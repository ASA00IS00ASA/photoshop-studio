#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gui_app.py —— Photoshop Studio 桌面版（Windows 原生窗口 + WebView2，不经浏览器）

用法:
    python gui_app.py            # 开发模式（控制台可见）
    python gui_app.py --mock     # 只看界面（假数据，连不上 Photoshop 也能预览设计）
    python gui_app.py --debug    # 打开 WebView2 开发者工具

结构:
    [WinForms 窗口 + WebView2]  ← HTML/CSS/JS（gui/web，无外链、离线可用）
            │  fetch 同源 HTTP
            ▼
    [内置 HTTP 服务：/ui/* 静态页 + /call 工具接口 + /llm/chat 模型代理]
            │  COM
            ▼
    [本机 Photoshop]
"""

from __future__ import annotations

import argparse
import ctypes
import json
import os
import sys
import threading
import traceback
import urllib.error
import urllib.request
from http.server import ThreadingHTTPServer
from urllib.parse import urlparse

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import photoshop_api as psapi  # noqa: E402

APP_NAME = "Photoshop Studio"
WEB_DIR = None  # 由 resolve_web_dir() 填充


def resource_root():
    """打包成 exe 后资源在 _MEIPASS；源码运行时在脚本目录。"""
    if getattr(sys, "frozen", False):
        return getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    return HERE


def resolve_web_dir():
    global WEB_DIR
    WEB_DIR = os.path.join(resource_root(), "gui", "web")
    return WEB_DIR


def user_dirs():
    home = os.path.expanduser("~")
    pictures = os.path.join(home, "Pictures")
    return home, (pictures if os.path.isdir(pictures) else home)


def setup_logging():
    """窗口模式没有控制台，把 stdout/stderr 落到日志文件，方便排查。"""
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    logdir = os.path.join(base, APP_NAME)
    try:
        os.makedirs(logdir, exist_ok=True)
        log_path = os.path.join(logdir, "studio.log")
        stream = open(log_path, "a", encoding="utf-8", buffering=1)
        if sys.stdout is None:
            sys.stdout = stream
        if sys.stderr is None:
            sys.stderr = stream
        print("\n=== %s %s ===" % (APP_NAME, psapi.VERSION))
        return log_path
    except Exception:
        return None


def show_error(text):
    try:
        ctypes.windll.user32.MessageBoxW(0, str(text), APP_NAME, 0x10)
    except Exception:
        pass


def fit_window_size(prefer=(1360, 860), margin=24, floor=(980, 640)):
    """把窗口收敛到当前显示器工作区内（逻辑像素，pywebview 会再按 DPI 缩放）。

    150% 缩放的笔记本上 1360×860 逻辑像素仍然放得下；更小的屏幕上自动收窄，
    避免窗口比桌面还大、右侧被切掉。"""
    try:
        from ctypes import wintypes

        user32 = ctypes.windll.user32
        try:
            dpi = int(user32.GetDpiForSystem())
        except Exception:
            dc = user32.GetDC(0)
            dpi = int(ctypes.windll.gdi32.GetDeviceCaps(dc, 88))
            user32.ReleaseDC(0, dc)
        scale = (dpi or 96) / 96.0
        rect = wintypes.RECT()
        if not user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(rect), 0):  # SPI_GETWORKAREA
            return prefer
        work_w = (rect.right - rect.left) / scale
        work_h = (rect.bottom - rect.top) / scale
        w = int(min(prefer[0], max(floor[0], work_w - margin * 2)))
        h = int(min(prefer[1], max(floor[1], work_h - margin * 2)))
        # 自己算居中位置：WinForms 的 CenterScreen 在这台 150% 缩放的机器上会把窗口底边压到任务栏下面
        return {
            "width": w,
            "height": h,
            "x": max(0, int((work_w - w) / 2)),
            "y": max(0, int((work_h - h) / 2)),
        }
    except Exception:
        return {"width": prefer[0], "height": prefer[1], "x": None, "y": None}


def single_instance_guard():
    """打包版只允许开一个实例；再次双击时把已有窗口叫到前台。

    返回 mutex 句柄（调用方需保持引用）；开发模式返回 None；已有实例返回 False。"""
    if not getattr(sys, "frozen", False):
        return None
    try:
        kernel32 = ctypes.windll.kernel32
        handle = kernel32.CreateMutexW(None, False, "Local\\PhotoshopStudio.SingleInstance")
        if handle and kernel32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
            user32 = ctypes.windll.user32
            hwnd = user32.FindWindowW(None, APP_NAME)
            if hwnd:
                user32.ShowWindow(hwnd, 9)  # SW_RESTORE
                user32.SetForegroundWindow(hwnd)
            return False
        return handle
    except Exception:
        return None


# ---------------------------------------------------------------------------
# HTTP 服务：静态 UI + 原有 API + 模型代理
# ---------------------------------------------------------------------------
MIME = {
    ".html": "text/html; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".js": "application/javascript; charset=utf-8",
    ".json": "application/json; charset=utf-8",
    ".svg": "image/svg+xml",
    ".png": "image/png",
    ".ico": "image/x-icon",
    ".woff2": "font/woff2",
}


class StudioHandler(psapi.APIHandler):
    server_version = "PhotoshopStudio/" + psapi.VERSION

    # -- 原始字节输出 ------------------------------------------------------
    def send_bytes(self, status, body, content_type):
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        try:
            self.wfile.write(body)
        except (BrokenPipeError, ConnectionResetError):
            pass

    def serve_static(self, rel):
        rel = (rel or "index.html").split("?")[0]
        target = os.path.normpath(os.path.join(WEB_DIR, rel))
        if not target.startswith(os.path.normpath(WEB_DIR)) or not os.path.isfile(target):
            return self.send_bytes(404, b"not found", "text/plain; charset=utf-8")
        with open(target, "rb") as fh:
            body = fh.read()
        ctype = MIME.get(os.path.splitext(target)[1].lower(), "application/octet-stream")
        self.send_bytes(200, body, ctype)

    # -- 路由 --------------------------------------------------------------
    def do_GET(self):
        path = urlparse(self.path).path
        if path == "/ui/config":
            home, pictures = user_dirs()
            return self._send(200, {"home": home, "pictures": pictures, "version": psapi.VERSION,
                                    "mock": False})
        if path.startswith("/ui/"):
            return self.serve_static(path[len("/ui/"):])
        return super().do_GET()

    def do_POST(self):
        if urlparse(self.path).path == "/llm/chat":
            return self.llm_chat()
        return super().do_POST()

    # -- 模型代理（绕过 WebView2 的跨域限制） ------------------------------
    def llm_chat(self):
        try:
            raw = self._read_body()
            data = json.loads(raw.decode("utf-8") or "{}")
        except Exception as exc:  # noqa: BLE001
            return self._send(400, {"ok": False, "error": "请求体不是合法 JSON: %s" % exc})

        base = str(data.get("base_url") or "https://api.openai.com/v1").rstrip("/")
        key = str(data.get("api_key") or "")
        if not key:
            return self._send(400, {"ok": False, "error": "未配置 API Key（右上角「设置」）"})
        payload = {
            "model": data.get("model") or "gpt-4o-mini",
            "messages": data.get("messages") or [],
            "temperature": 0,
        }
        tools = data.get("tools") or []
        if tools:
            # 兼容两种写法：OpenAI 的 {"type":"function","function":{...}} 和裸函数定义
            norm = []
            for item in tools:
                if isinstance(item, dict) and (item.get("function") or {}).get("name"):
                    norm.append(item)
                elif isinstance(item, dict) and item.get("name"):
                    norm.append({"type": "function", "function": item})
            if norm:
                payload["tools"] = norm
                payload["tool_choice"] = "auto"

        url = base if base.endswith("/chat/completions") else base + "/chat/completions"
        req = urllib.request.Request(
            url,
            data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
            headers={"Content-Type": "application/json", "Authorization": "Bearer " + key},
            method="POST",
        )
        try:
            with urllib.request.urlopen(req, timeout=180) as resp:
                body = json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", "replace")[:600]
            return self._send(502, {"ok": False, "error": "模型接口 HTTP %s：%s" % (exc.code, detail)})
        except Exception as exc:  # noqa: BLE001
            return self._send(502, {"ok": False, "error": "无法访问模型接口：%s" % exc})

        try:
            message = body["choices"][0]["message"]
        except Exception:  # noqa: BLE001
            return self._send(502, {"ok": False, "error": "模型返回结构异常：%s" % str(body)[:400]})
        return self._send(200, {"ok": True, "message": message, "usage": body.get("usage")})


# WebView2/Chromium 认为这些端口「不安全」，页面会直接报 ERR_UNSAFE_PORT。
# 随机端口偶尔会落进来（例如 6669），所以分配后要校验。
_UNSAFE_PORTS = {
    1, 7, 9, 11, 13, 15, 17, 19, 20, 21, 22, 23, 25, 37, 42, 43, 53, 69, 77, 79,
    87, 95, 101, 102, 103, 104, 109, 110, 111, 113, 115, 117, 119, 123, 135, 137,
    139, 143, 161, 179, 389, 427, 465, 512, 513, 514, 515, 526, 530, 531, 532,
    540, 548, 554, 556, 563, 587, 601, 636, 989, 990, 993, 995, 1719, 1720, 1723,
    2049, 3659, 4045, 5060, 5061, 6000, 6566, 6665, 6666, 6667, 6668, 6669, 6697,
    10080,
}


def start_api(host="127.0.0.1", port=0):
    resolve_web_dir()
    psapi.QUIET = True
    psapi.TOKEN = ""
    if psapi.BRIDGE is None:
        psapi.BRIDGE = psapi.PSBridge(launch=True)
    ThreadingHTTPServer.allow_reuse_address = True
    for _ in range(30):
        httpd = ThreadingHTTPServer((host, port), StudioHandler)
        actual = httpd.server_address[1]
        if actual in _UNSAFE_PORTS:
            httpd.server_close()
            if port:
                raise OSError("端口 %s 被 WebView2 列为不安全端口，请换一个（例如 18765）" % port)
            continue
        break
    else:
        raise OSError("连续多次分配到的端口都被 WebView2 拒绝，请用 --port 指定一个（例如 18765）")
    httpd.daemon_threads = True
    threading.Thread(target=httpd.serve_forever, name="studio-http", daemon=True).start()
    return httpd, actual


# ---------------------------------------------------------------------------
def main(argv=None):
    ap = argparse.ArgumentParser(description="Photoshop Studio 桌面版")
    ap.add_argument("--port", type=int, default=0, help="内置服务端口，0 = 自动分配")
    ap.add_argument("--debug", action="store_true", help="打开 WebView2 开发者工具")
    ap.add_argument("--mock", action="store_true", help="假数据界面预览（不连 Photoshop）")
    args = ap.parse_args(argv)

    setup_logging()
    resolve_web_dir()

    if single_instance_guard() is False:
        print("已有实例在运行，已切到前台。")
        return 0

    if not os.path.isdir(WEB_DIR):
        show_error("找不到界面资源目录：\n%s" % WEB_DIR)
        return 1

    if not psapi.HAS_COM:
        show_error("缺少 pywin32，请先运行：\n\npython -m pip install pywin32")
        return 1

    try:
        httpd, port = start_api(port=args.port)
    except OSError as exc:
        show_error("无法启动内置服务（端口 %s）：\n%s" % (args.port, exc))
        return 1

    url = "http://127.0.0.1:%d/ui/index.html" % port
    if args.mock:
        url += "?mock=1"
    print("UI:", url)

    try:
        import webview
    except ImportError:
        show_error("缺少 pywebview，请先运行：\n\npython -m pip install pywebview")
        return 1

    def after_gui():
        """冷启动时是本程序把 Photoshop 拉起来的，PS 的启动窗口会盖住 Studio；
        等连接建立后把 Studio 窗口叫回前台一次。"""
        import time

        try:
            win = webview.windows[0]
        except Exception:
            return
        t0 = time.time()
        connected = False
        while time.time() - t0 < 45:
            try:
                if psapi.BRIDGE is not None and psapi.BRIDGE.status().get("connected"):
                    connected = True
                    break
            except Exception:
                break
            time.sleep(0.4)
        if not connected:
            return
        if time.time() - t0 < 3.0:
            return  # 秒连上说明 Photoshop 本来就在跑，不要抢用户焦点
        try:
            hwnd = win.native.Handle.ToInt64()
        except Exception:
            try:
                hwnd = int(win.native.Handle)
            except Exception:
                return
        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32
        # PS 的欢迎页常是 TOPMOST，普通置顶压不下去；临时把自己的窗口也设为 TOPMOST 再恢复。
        for _ in range(3):
            try:
                win.on_top = True
                user32.ShowWindow(hwnd, 9)  # SW_RESTORE
                fg = user32.GetForegroundWindow()
                fg_tid = user32.GetWindowThreadProcessId(fg, None) if fg else 0
                cur_tid = kernel32.GetCurrentThreadId()
                if fg_tid and fg_tid != cur_tid:
                    user32.AttachThreadInput(cur_tid, fg_tid, True)
                user32.SetForegroundWindow(hwnd)
                if fg_tid and fg_tid != cur_tid:
                    user32.AttachThreadInput(cur_tid, fg_tid, False)
            except Exception:
                pass
            time.sleep(1.5)
            try:
                win.on_top = False
            except Exception:
                pass
            time.sleep(1.5)

    geo = fit_window_size()
    print("window: %dx%d at %s,%s (logical px)" % (geo["width"], geo["height"], geo["x"], geo["y"]))
    webview.create_window(
        APP_NAME,
        url,
        width=geo["width"],
        height=geo["height"],
        x=geo["x"],
        y=geo["y"],
        min_size=(min(980, geo["width"]), min(640, geo["height"])),
        background_color="#070A10",
        text_select=False,
    )
    try:
        # private_mode=False：模型接口地址 / API Key 存在 localStorage，重启后还在
        webview.start(after_gui, debug=args.debug, private_mode=False)
    except Exception:
        show_error("窗口启动失败：\n\n" + traceback.format_exc(limit=3))
        return 1
    finally:
        try:
            httpd.shutdown()
        except Exception:
            pass
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except Exception:
        show_error("启动失败：\n\n" + traceback.format_exc(limit=4))
        sys.exit(1)

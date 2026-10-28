#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
photoshop_api.py —— 让模型通过 HTTP/JSON 接口操作本机 Photoshop
================================================================

原理
----
    [模型 / 任意客户端]
            │  HTTP JSON  (127.0.0.1:8765)
            ▼
    [本脚本 photoshop_api.py]
            │  COM: Photoshop.Application.DoJavaScript()
            ▼
    [本机 Photoshop（ExtendScript 引擎）]
            ▼
        真实文档 / 图层 / 导出文件

模型只需要两件事：
    1. GET  /tools          取到 OpenAI function-calling 格式的工具清单
    2. POST /call           按 {"name": ..., "arguments": {...}} 调用工具

快速开始
--------
    python -m pip install pywin32
    python photoshop_api.py --selftest        # 端到端自检，生成一预览图
    python photoshop_api.py                   # 启动 API 服务

调用示例
--------
    curl http://127.0.0.1:8765/health
    curl http://127.0.0.1:8765/tools
    curl -X POST http://127.0.0.1:8765/call -H "Content-Type: application/json" ^
         -d "{\"name\":\"ps_doc_create\",\"arguments\":{\"width\":1200,\"height\":800}}"
    curl -X POST http://127.0.0.1:8765/jsx -H "Content-Type: text/plain" ^
         -d "app.documents.add(800, 600); 'ok';"

说明
----
* 首次调用会自动启动 Photoshop（可能会弹出登录/欢迎窗口，脚本会等待约 10 秒）。
* 所有 Photoshop 操作在主线程之外的一个 COM 工作线程串行执行，避免线程冲突。
* 服务默认只绑定 127.0.0.1；如需局域网访问请显式 --host 0.0.0.0 并设置 --token。
* 这是「真机真操作」，没有沙箱：调用会直接修改打开的文档，请自行备份重要文件。
"""

from __future__ import annotations

import argparse
import base64
import json
import os
import queue
import shutil
import sys
import tempfile
import threading
import time
import traceback
from concurrent.futures import Future, TimeoutError as FuturesTimeout
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

try:
    import pythoncom
    import pywintypes
    import win32com.client
    HAS_COM = True
except ImportError:  # 允许在没有 pywin32 的机器上查看 --help
    pythoncom = pywintypes = win32com = None
    HAS_COM = False

VERSION = "1.0.0"
START_TS = time.time()
TEMP_DIR = os.path.join(tempfile.gettempdir(), "photoshop_api")
os.makedirs(TEMP_DIR, exist_ok=True)

try:
    from PIL import Image  # 可选：用于把预览图缩到适合视觉模型的大小
except Exception:
    Image = None


# ---------------------------------------------------------------------------
# 工具清单（name -> Tool）
# ---------------------------------------------------------------------------
TOOLS: dict[str, "Tool"] = {}


class Tool:
    def __init__(self, name, description, parameters, required, handler, timeout):
        self.name = name
        self.description = description
        self.parameters = parameters
        self.required = required
        self.handler = handler
        self.timeout = timeout

    def schema(self):
        params = {"type": "object", "properties": self.parameters}
        if self.required:
            params["required"] = list(self.required)
        return {
            "type": "function",
            "function": {
                "name": self.name,
                "description": self.description,
                "parameters": params,
            },
        }


def tool(name, description, parameters=None, required=None, timeout=120.0):
    def deco(fn):
        TOOLS[name] = Tool(name, description, parameters or {}, required or [], fn, timeout)
        return fn
    return deco


def P(type_, description, enum=None, default=None, **extra):
    """构造一个 JSON Schema 属性。"""
    prop = {"type": type_, "description": description}
    if enum:
        prop["enum"] = list(enum)
    if default is not None:
        prop["default"] = default
    prop.update(extra)
    return prop


def j(value):
    """参数 -> JS 字面量（ensure_ascii 保证跨 COM 的中文安全）。"""
    return json.dumps(value, ensure_ascii=True)


def render(tpl, args):
    return tpl.replace("$ARGS", j(args))


# ---------------------------------------------------------------------------
# ExtendScript 前奏：JSON 序列化 + 通用辅助函数
# （ExtendScript 没有 JSON 对象，必须自带序列化器）
# ---------------------------------------------------------------------------
PRELUDE = r"""
function __j(v) {
  if (v === null || v === undefined) return "null";
  var t = typeof v;
  if (t === "number") return isFinite(v) ? String(v) : "null";
  if (t === "boolean") return v ? "true" : "false";
  if (t === "string") {
    var s = "", i, ch, c;
    for (i = 0; i < v.length; i++) {
      ch = v.charAt(i); c = v.charCodeAt(i);
      if (ch === '"') s += '\\"';
      else if (ch === '\\') s += '\\\\';
      else if (ch === '\n') s += '\\n';
      else if (ch === '\r') s += '\\r';
      else if (ch === '\t') s += '\\t';
      else if (c < 32) s += '\\u' + ('0000' + c.toString(16)).slice(-4);
      else s += ch;
    }
    return '"' + s + '"';
  }
  var i2, parts = [], out = [], key, val;
  if (v instanceof Array) {
    for (i2 = 0; i2 < v.length; i2++) parts.push(__j(v[i2]));
    return "[" + parts.join(",") + "]";
  }
  for (key in v) {
    if (!v.hasOwnProperty(key)) continue;
    val = v[key];
    if (typeof val === "function") continue;
    out.push(__j(String(key)) + ":" + __j(val));
  }
  return "{" + out.join(",") + "}";
}

function __enum(e) {
  var s = String(e);
  var i = s.lastIndexOf(".");
  return i >= 0 ? s.substring(i + 1) : s;
}

function __hex2rgb(h) {
  h = String(h === null || h === undefined ? "#000000" : h).replace(/^#/, "");
  if (h.length === 3) h = h.charAt(0) + h.charAt(0) + h.charAt(1) + h.charAt(1) + h.charAt(2) + h.charAt(2);
  return {
    r: parseInt(h.substring(0, 2), 16) || 0,
    g: parseInt(h.substring(2, 4), 16) || 0,
    b: parseInt(h.substring(4, 6), 16) || 0
  };
}

function __color(h) {
  var v = __hex2rgb(h);
  var c = new SolidColor();
  c.rgb.red = v.r; c.rgb.green = v.g; c.rgb.blue = v.b;
  return c;
}

function __doc(i) {
  var docs = app.documents;
  if (docs.length === 0) throw new Error("当前没有打开的文档");
  if (i === null || i === undefined) return app.activeDocument;
  var d = docs[Number(i)];
  if (!d) throw new Error("文档序号超出范围: " + i + "（共 " + docs.length + " 个）");
  return d;
}

function __docIndex(d) {
  var i;
  for (i = 0; i < app.documents.length; i++) if (app.documents[i] === d) return i;
  for (i = 0; i < app.documents.length; i++) if (app.documents[i].name === d.name) return i;
  return -1;
}

function __isGroup(l) {
  try { return l.typename === "LayerSet"; } catch (e) { return false; }
}

function __layer(d, name, index) {
  if (name !== null && name !== undefined && String(name) !== "") {
    var wanted = String(name), stack = [], i, l, found = null;
    for (i = 0; i < d.layers.length; i++) stack.push(d.layers[i]);
    while (stack.length > 0) {
      l = stack.shift();
      if (l.name === wanted) { found = l; break; }
      if (__isGroup(l)) { for (i = 0; i < l.layers.length; i++) stack.push(l.layers[i]); }
    }
    if (!found) throw new Error("找不到图层: " + wanted);
    return found;
  }
  if (index !== null && index !== undefined) {
    var top = d.layers[Number(index)];
    if (!top) throw new Error("图层序号超出范围: " + index + "（共 " + d.layers.length + " 个）");
    return top;
  }
  return d.activeLayer;
}

function __kind(l) {
  if (__isGroup(l)) return "group";
  try {
    // 不能直接和 LayerKind.XXX 做 === 比较：新版 Photoshop 里某些枚举成员
    // （如 SOLIDFILL/PATTERNFILL）访问或比较会抛「无效枚举值」。用字符串名判断。
    var s = String(l.kind).replace(/^LayerKind\.?/i, "").toLowerCase();
    if (!s || s === "undefined" || s === "null") return "unknown";
    return s;
  } catch (e) {
    return "unknown";
  }
}

function __layerJson(l) {
  var o = {
    name: l.name,
    kind: __kind(l),
    visible: !!l.visible,
    opacity: Math.round(l.opacity),
    blend: __enum(l.blendMode)
  };
  try {
    var b = l.bounds;
    if (b && b.length === 4) {
      o.bounds = [Math.round(b[0].as("px")), Math.round(b[1].as("px")), Math.round(b[2].as("px")), Math.round(b[3].as("px"))];
    } else {
      o.bounds = null;
    }
  } catch (e) {
    o.bounds = null;
  }
  if (__isGroup(l)) {
    o.layers = [];
    for (var i = 0; i < l.layers.length; i++) o.layers.push(__layerJson(l.layers[i]));
  } else {
    try {
      if (l.kind === LayerKind.TEXT) {
        o.text = l.textItem.contents;
        o.font = String(l.textItem.font);
      }
    } catch (e) {}
  }
  return o;
}

function __docJson(d, i) {
  var o = {
    index: i,
    name: d.name,
    width: Math.round(d.width.as("px")),
    height: Math.round(d.height.as("px")),
    resolution: Math.round(d.resolution),
    mode: __enum(d.mode),
    bits: __enum(d.bitsPerChannel),
    layer_count: d.layers.length,
    active: (d === app.activeDocument)
  };
  try { o.path = d.fullName ? d.fullName.fsName : null; } catch (e) { o.path = null; }
  try { o.profile = d.colorProfileName; } catch (e) { o.profile = null; }
  try {
    var sb = d.selection.bounds;
    o.selection = (sb && sb.length === 4)
      ? [Math.round(sb[0].as("px")), Math.round(sb[1].as("px")), Math.round(sb[2].as("px")), Math.round(sb[3].as("px"))]
      : null;
  } catch (e) { o.selection = null; }
  return o;
}

function __anchor(name) {
  var map = {
    "top-left": AnchorPosition.TOPLEFT, "top": AnchorPosition.TOPCENTER, "top-right": AnchorPosition.TOPRIGHT,
    "left": AnchorPosition.MIDDLELEFT, "center": AnchorPosition.MIDDLECENTER, "middle": AnchorPosition.MIDDLECENTER,
    "right": AnchorPosition.MIDDLERIGHT, "bottom-left": AnchorPosition.BOTTOMLEFT,
    "bottom": AnchorPosition.BOTTOMCENTER, "bottom-right": AnchorPosition.BOTTOMRIGHT
  };
  var k = String(name === null || name === undefined ? "center" : name).toLowerCase();
  return map[k] === undefined ? AnchorPosition.MIDDLECENTER : map[k];
}

function __absPath(p) {
  var f = new File(String(p));
  return f.fsName;
}
"""


# ---------------------------------------------------------------------------
# 结果 / 异常
# ---------------------------------------------------------------------------
class PSError(Exception):
    """Photoshop 端失败（未启动、脚本报错、超时等）。"""


def com_error_text(exc):
    """把 pywin32 的 COM 异常转成可读文本（含 PS 返回的中文报错）。"""
    try:
        hresult = exc.args[0] if exc.args else None
        excepinfo = exc.args[2] if len(exc.args) > 2 else None
        desc = None
        if excepinfo is not None:
            if isinstance(excepinfo, tuple):
                if len(excepinfo) > 2:
                    desc = excepinfo[2]
            else:
                desc = getattr(excepinfo, "strDescription", None)
        if not desc:
            desc = getattr(exc, "strerror", None) or str(exc)
        if isinstance(hresult, int):
            return "%s [0x%08X]" % (str(desc).strip(), hresult & 0xFFFFFFFF)
        return str(desc).strip()
    except Exception:
        return str(exc)


def _is_connection_error(exc):
    s = str(exc).lower()
    for key in (
        "rpc", "0x800706ba", "0x800706be", "0x80080005", "0x80040154",
        "server execution failed", "class not registered",
        "远程过程调用", "连接", "无法启动", "unavailable", "已断开",
    ):
        if key in s:
            return True
    return False


# ---------------------------------------------------------------------------
# COM 桥：所有 Photoshop 调用都在这一个工作线程里串行执行
# ---------------------------------------------------------------------------
class PSBridge:
    def __init__(self, launch=True):
        self._q: queue.Queue = queue.Queue()
        self._app = None
        self._launch = launch
        self._version = None
        self._connected = False
        self._thread = threading.Thread(target=self._worker, name="photoshop-com", daemon=True)
        self._thread.start()

    # -- 工作线程 ----------------------------------------------------------
    def _worker(self):
        pythoncom.CoInitialize()
        try:
            while True:
                fn, fut = self._q.get()
                if fn is None:
                    break
                if not fut.set_running_or_notify_cancel():
                    continue
                try:
                    fut.set_result(fn())
                except BaseException as exc:  # noqa: BLE001 - 原样传回调用方
                    fut.set_exception(exc)
        finally:
            pythoncom.CoUninitialize()

    def submit(self, fn, timeout):
        fut = Future()
        self._q.put((fn, fut))
        try:
            return fut.result(timeout=timeout)
        except FuturesTimeout:
            raise PSError("Photoshop 在 %.0f 秒内没有响应（可能弹出了模态对话框或正在处理大文件）" % timeout)

    # -- 连接 --------------------------------------------------------------
    def _get_app(self):
        if self._app is not None:
            try:
                _ = str(self._app.Name)
                return self._app
            except Exception:
                self._app = None
                self._connected = False

        if not self._launch:
            # 只连接已经运行的实例
            self._app = win32com.client.GetActiveObject("Photoshop.Application")
        else:
            try:
                self._app = win32com.client.GetActiveObject("Photoshop.Application")
            except Exception:
                self._app = win32com.client.Dispatch("Photoshop.Application")

        try:
            self._app.displayDialogs = 3          # DialogModes.NO：禁止弹窗卡死自动化
            self._app.preferences.rulerUnits = 1  # Units.PIXELS
            self._app.preferences.typeUnits = 2   # TypeUnits.POINTS
        except Exception:
            pass
        try:
            self._version = str(self._app.Version)
        except Exception:
            self._version = None
        self._connected = True
        return self._app

    # -- 执行 --------------------------------------------------------------
    def call_jsx(self, jsx, timeout=120.0):
        def go():
            return self._get_app().DoJavaScript(jsx, [], 1)  # 1 = 不弹调试器

        try:
            return self.submit(go, timeout)
        except PSError:
            raise
        except Exception as exc:
            if _is_connection_error(exc):
                self._app = None
                self._connected = False
                try:
                    return self.submit(go, timeout)
                except Exception as exc2:
                    raise PSError(com_error_text(exc2))
            raise PSError(com_error_text(exc))

    def status(self):
        return {"connected": bool(self._connected), "version": self._version}

    def probe(self):
        body = """
var out = {version: app.version, name: app.name, documents: app.documents.length};
__r = out;
"""
        data = run_jsx(body, timeout=180)
        return {"version": data.get("version"), "app": data.get("name"), "documents": data.get("documents")}


BRIDGE: "PSBridge | None" = None


def run_jsx(body, timeout=120.0):
    """执行一段 JSX 主体（自动套错误处理 + JSON 序列化），返回 Python 对象。"""
    if BRIDGE is None:
        raise PSError("服务尚未初始化")
    jsx = (
        PRELUDE
        + "\n;(function(){var __r = null;try{\n"
        + body
        + "\n}catch(__e){return __j({__error: String(__e), line: (__e && __e.line) || null});}\n"
        + "return __j(__r);})();"
    )
    raw = BRIDGE.call_jsx(jsx, timeout)
    if raw is None or (isinstance(raw, str) and raw.strip() == ""):
        return None
    try:
        data = json.loads(raw) if isinstance(raw, str) else None
    except ValueError:
        raise PSError("Photoshop 返回了无法解析的结果: %r" % (str(raw)[:400],))
    if isinstance(data, dict) and "__error" in data:
        msg = data.get("__error")
        line = data.get("line")
        raise PSError(str(msg) + ("（JSX 第 %s 行）" % line if line else ""))
    return data


def resolve_path(path):
    if not path:
        raise PSError("path 参数不能为空")
    return os.path.abspath(os.path.expanduser(str(path)))


# ---------------------------------------------------------------------------
# 工具实现
# ---------------------------------------------------------------------------
@tool("ps_status", "查看 Photoshop 连接状态、版本，以及当前打开的文档列表。不会启动 Photoshop。", {}, [], timeout=30)
def h_status(a):
    body = """
var out = { photoshop: app.name, version: app.version, documents: [] };
for (var i = 0; i < app.documents.length; i++) out.documents.push(__docJson(app.documents[i], i));
__r = out;
"""
    return run_jsx(body)


@tool("ps_doc_create", "新建一个文档（画布）。默认 72dpi、RGB、8 位、sRGB。",
      {
          "width": P("integer", "画布宽度（像素）", minimum=1),
          "height": P("integer", "画布高度（像素）", minimum=1),
          "resolution": P("number", "分辨率 dpi，默认 72", default=72),
          "name": P("string", "文档名称，默认 Untitled"),
          "background": P("string", "初始内容：white 白色 / transparent 透明 / color 背景色",
                          enum=["white", "transparent", "color"], default="white"),
          "color": P("string", "background=color 时使用的背景色，如 #ff7a18", default="#ffffff"),
      },
      ["width", "height"])
def h_doc_create(a):
    body = """
var o = $ARGS;
var bg = String(o.background || "white");
if (bg === "color") app.backgroundColor = __color(o.color || "#ffffff");
var fill = DocumentFill.WHITE;
if (bg === "transparent") fill = DocumentFill.TRANSPARENT;
else if (bg === "color") fill = DocumentFill.BACKGROUNDCOLOR;
var d = app.documents.add(
  UnitValue(o.width, "px"), UnitValue(o.height, "px"), o.resolution || 72,
  o.name || "Untitled", NewDocumentMode.RGB, fill, 1.0,
  BitsPerChannelType.EIGHT, "sRGB IEC61966-2.1");
__r = __docJson(d, __docIndex(d));
"""
    return run_jsx(render(body, a))


@tool("ps_doc_open", "打开磁盘上的图片 / PSD 文件作为新文档。",
      {"path": P("string", "文件路径，如 D:\\\\images\\\\a.png")},
      ["path"])
def h_doc_open(a):
    body = """
var o = $ARGS;
var f = new File(o.path);
if (!f.exists) throw new Error("文件不存在: " + o.path);
var d = app.open(f);
__r = __docJson(d, __docIndex(d));
"""
    return run_jsx(render(body, {"path": resolve_path(a.get("path"))}))


@tool("ps_doc_info", "获取指定文档（默认当前活动文档）的详细信息：尺寸、分辨率、颜色模式、图层数、选区等。",
      {"index": P("integer", "文档序号，默认当前活动文档")}, [])
def h_doc_info(a):
    body = """
var o = $ARGS;
var d = __doc(o.index);
__r = __docJson(d, __docIndex(d));
"""
    return run_jsx(render(body, {"index": a.get("index")}))


@tool("ps_doc_save", "把文档保存/导出为文件。支持 psd / png / jpg / tif；jpg 会自动合并图层。",
      {
          "path": P("string", "目标文件完整路径（相对路径按服务端工作目录解析）"),
          "format": P("string", "格式；不给则按扩展名推断", enum=["psd", "png", "jpg", "tif"]),
          "quality": P("integer", "JPEG 质量 0-100，默认 90（内部映射到 Photoshop 的 0-12 档）", default=90),
          "as_copy": P("boolean", "是否另存副本（不改动原文档路径），默认 true", default=True),
          "index": P("integer", "文档序号，默认当前活动文档"),
      },
      ["path"])
def h_doc_save(a):
    body = """
var o = $ARGS;
var d = __doc(o.index);
var path = String(o.path);
var fmt = o.format ? String(o.format).toLowerCase() : path.substring(path.lastIndexOf(".") + 1).toLowerCase();
if (fmt === "jpeg") fmt = "jpg";
if (fmt === "tiff") fmt = "tif";
var asCopy = (o.as_copy !== false);
var f = new File(path);
var folder = f.parent;
if (folder && !folder.exists) folder.create();
if (fmt === "psd") {
  var so = new PhotoshopSaveOptions();
  so.layers = true;
  so.embedColorProfile = true;
  d.saveAs(f, so, asCopy);
} else if (fmt === "png") {
  var po = new PNGSaveOptions();
  po.compression = 6;
  po.interlaced = false;
  d.saveAs(f, po, asCopy);
} else if (fmt === "jpg") {
  var jo = new JPEGSaveOptions();
  // Photoshop 的 JPEGSaveOptions.quality 是 0-12 档；对外统一用 0-100
  var q = (o.quality === undefined || o.quality === null) ? 90 : Number(o.quality);
  jo.quality = Math.max(0, Math.min(12, Math.round(q / 100 * 12)));
  jo.embedColorProfile = true;
  var dup = d.duplicate("__export__");
  try { dup.flatten(); } catch (e) {}
  dup.saveAs(f, jo, true);
  dup.close(SaveOptions.DONOTSAVECHANGES);
} else if (fmt === "tif") {
  var to = new TiffSaveOptions();
  to.imageCompression = TIFFEncoding.TIFFLZW;
  d.saveAs(f, to, asCopy);
} else {
  throw new Error("不支持的格式: " + fmt + "（可用 psd/png/jpg/tif）");
}
__r = { path: f.fsName, format: fmt, as_copy: asCopy };
"""
    res = run_jsx(render(body, {"path": resolve_path(a.get("path")), "format": a.get("format"),
                                "quality": a.get("quality"), "as_copy": a.get("as_copy"),
                                "index": a.get("index")}))
    try:
        res["bytes"] = os.path.getsize(res["path"]) if os.path.exists(res["path"]) else 0
    except OSError:
        res["bytes"] = 0
    return res


@tool("ps_doc_close", "关闭文档（默认不保存）。",
      {
          "index": P("integer", "文档序号，默认当前活动文档"),
          "save": P("boolean", "关闭前是否保存，默认 false", default=False),
      }, [])
def h_doc_close(a):
    body = """
var o = $ARGS;
var d = __doc(o.index);
var name = d.name;
d.close(o.save ? SaveOptions.SAVECHANGES : SaveOptions.DONOTSAVECHANGES);
__r = { closed: name };
"""
    return run_jsx(render(body, {"index": a.get("index"), "save": bool(a.get("save"))}))


@tool("ps_doc_resize_image", "调整文档像素尺寸（重采样）。width/height/resolution 至少给一个。",
      {
          "width": P("integer", "目标宽度（像素）"),
          "height": P("integer", "目标高度（像素）"),
          "resolution": P("number", "目标分辨率 dpi"),
          "index": P("integer", "文档序号，默认当前活动文档"),
      }, [])
def h_doc_resize_image(a):
    body = """
var o = $ARGS;
var d = __doc(o.index);
var curW = Math.round(d.width.as("px")), curH = Math.round(d.height.as("px"));
var w = o.width, h = o.height;
if (!w && !h && o.resolution) { var f = o.resolution / d.resolution; w = Math.round(curW * f); h = Math.round(curH * f); }
if (w && !h) h = Math.round(curH * w / curW);
if (h && !w) w = Math.round(curW * h / curH);
if (!w || !h) throw new Error("width / height / resolution 至少给一个");
d.resizeImage(UnitValue(w, "px"), UnitValue(h, "px"), o.resolution || undefined, ResampleMethod.BICUBICSMOOTHER);
__r = __docJson(d, __docIndex(d));
"""
    return run_jsx(render(body, {"width": a.get("width"), "height": a.get("height"),
                                 "resolution": a.get("resolution"), "index": a.get("index")}))


@tool("ps_doc_resize_canvas", "调整画布大小（不重采样，裁切或扩展）。",
      {
          "width": P("integer", "目标画布宽度（像素）"),
          "height": P("integer", "目标画布高度（像素）"),
          "anchor": P("string", "扩展/裁切基准点", default="center",
                      enum=["top-left", "top", "top-right", "left", "center", "right",
                            "bottom-left", "bottom", "bottom-right"]),
          "index": P("integer", "文档序号，默认当前活动文档"),
      },
      ["width", "height"])
def h_doc_resize_canvas(a):
    body = """
var o = $ARGS;
var d = __doc(o.index);
d.resizeCanvas(UnitValue(o.width, "px"), UnitValue(o.height, "px"), __anchor(o.anchor));
__r = __docJson(d, __docIndex(d));
"""
    return run_jsx(render(body, a))


@tool("ps_doc_flatten", "拼合所有图层（会丢失图层结构）。",
      {"index": P("integer", "文档序号，默认当前活动文档")}, [])
def h_doc_flatten(a):
    body = """
var o = $ARGS;
var d = __doc(o.index);
d.flatten();
__r = __docJson(d, __docIndex(d));
"""
    return run_jsx(render(body, {"index": a.get("index")}))


@tool("ps_layers_list", "列出文档的图层树（组会递归展开），含名称、类型、可见性、不透明度、混合模式、边界、文字内容。",
      {"index": P("integer", "文档序号，默认当前活动文档")}, [])
def h_layers_list(a):
    body = """
var o = $ARGS;
var d = __doc(o.index);
var out = { document: d.name, layers: [] };
for (var i = 0; i < d.layers.length; i++) out.layers.push(__layerJson(d.layers[i]));
__r = out;
"""
    return run_jsx(render(body, {"index": a.get("index")}))


@tool("ps_layer_add", "新建空白图层；可选给整个图层填充颜色。",
      {
          "name": P("string", "图层名"),
          "color": P("string", "填充颜色 #RRGGBB；不给则为透明图层"),
          "index": P("integer", "文档序号，默认当前活动文档"),
      }, [])
def h_layer_add(a):
    body = """
var o = $ARGS;
var d = __doc(o.index);
var l = d.artLayers.add();
if (o.name) l.name = String(o.name);
if (o.color) {
  d.selection.selectAll();
  d.selection.fill(__color(o.color), ColorBlendMode.NORMAL, 100, false);
  d.selection.deselect();
}
__r = __layerJson(l);
"""
    return run_jsx(render(body, a))


@tool("ps_layer_set", "修改图层属性：重命名、不透明度、混合模式、可见性、设为当前图层。",
      {
          "name": P("string", "目标图层名"),
          "layer_index": P("integer", "目标图层序号（与 name 二选一）"),
          "new_name": P("string", "新的图层名"),
          "opacity": P("integer", "不透明度 0-100", minimum=0, maximum=100),
          "blend": P("string", "混合模式，如 normal / multiply / screen / overlay / darken / lighten"),
          "visible": P("boolean", "是否可见"),
          "select": P("boolean", "是否设为当前活动图层"),
          "index": P("integer", "文档序号，默认当前活动文档"),
      }, [])
def h_layer_set(a):
    body = """
var o = $ARGS;
var d = __doc(o.index);
var l = __layer(d, o.name, o.layer_index);
if (o.new_name) l.name = String(o.new_name);
if (o.opacity !== null && o.opacity !== undefined) l.opacity = o.opacity;
if (o.blend) {
  var bm = BlendMode[String(o.blend).toUpperCase()];
  if (bm === undefined) throw new Error("不支持的混合模式: " + o.blend);
  l.blendMode = bm;
}
if (o.visible !== null && o.visible !== undefined) l.visible = !!o.visible;
if (o.select) d.activeLayer = l;
__r = __layerJson(l);
"""
    return run_jsx(render(body, a))


@tool("ps_layer_delete", "删除图层（按名称或序号）。",
      {
          "name": P("string", "目标图层名"),
          "layer_index": P("integer", "目标图层序号（与 name 二选一）"),
          "index": P("integer", "文档序号，默认当前活动文档"),
      }, [])
def h_layer_delete(a):
    body = """
var o = $ARGS;
var d = __doc(o.index);
var l = __layer(d, o.name, o.layer_index);
var nm = l.name;
l.remove();
__r = { deleted: nm };
"""
    return run_jsx(render(body, a))


@tool("ps_layer_duplicate", "复制图层。",
      {
          "name": P("string", "目标图层名"),
          "layer_index": P("integer", "目标图层序号（与 name 二选一）"),
          "new_name": P("string", "副本图层名"),
          "index": P("integer", "文档序号，默认当前活动文档"),
      }, [])
def h_layer_duplicate(a):
    body = """
var o = $ARGS;
var d = __doc(o.index);
var l = __layer(d, o.name, o.layer_index);
var nl = l.duplicate();
if (o.new_name) nl.name = String(o.new_name);
__r = __layerJson(nl);
"""
    return run_jsx(render(body, a))


@tool("ps_shape", "画一个填充形状（矩形或椭圆），在新建图层或当前图层上栅格化填充。坐标单位为文档像素。",
      {
          "shape": P("string", "形状类型", enum=["rect", "ellipse"], default="rect"),
          "x": P("number", "左上角 X（像素）", default=0),
          "y": P("number", "左上角 Y（像素）", default=0),
          "width": P("number", "宽度（像素）", minimum=1),
          "height": P("number", "高度（像素）", minimum=1),
          "color": P("string", "填充色 #RRGGBB", default="#000000"),
          "opacity": P("integer", "图层不透明度 0-100", default=100),
          "new_layer": P("boolean", "是否新建图层，默认 true", default=True),
          "name": P("string", "新图层名"),
          "index": P("integer", "文档序号，默认当前活动文档"),
      },
      ["width", "height"])
def h_shape(a):
    body = """
var o = $ARGS;
var d = __doc(o.index);
var x = Number(o.x || 0), y = Number(o.y || 0), w = Number(o.width), h = Number(o.height);
if (!(w > 0) || !(h > 0)) throw new Error("width / height 必须大于 0");
var layer;
if (o.new_layer === false) {
  layer = d.activeLayer;
} else {
  layer = d.artLayers.add();
  if (o.name) layer.name = String(o.name);
}
d.activeLayer = layer;
var pts = [], k, a2;
if (String(o.shape) === "ellipse") {
  var cx = x + w / 2, cy = y + h / 2, n = 128;
  for (k = 0; k < n; k++) {
    a2 = 2 * Math.PI * k / n;
    pts.push([cx + Math.cos(a2) * w / 2, cy + Math.sin(a2) * h / 2]);
  }
} else {
  pts = [[x, y], [x + w, y], [x + w, y + h], [x, y + h]];
}
d.selection.select(pts, SelectionType.REPLACE, 0, true);
d.selection.fill(__color(o.color || "#000000"), ColorBlendMode.NORMAL, 100, false);
d.selection.deselect();
if (o.opacity !== null && o.opacity !== undefined) layer.opacity = o.opacity;
__r = __layerJson(layer);
"""
    return run_jsx(render(body, a))


@tool("ps_text", "添加文字图层。x/y 是文字基线的左上点（像素）；中文建议 font 用 MicrosoftYaHei 或 SourceHanSansCN-Regular。",
      {
          "content": P("string", "文字内容，支持换行 \\n"),
          "font": P("string", "字体 PostScript 名，如 MicrosoftYaHei / MicrosoftYaHei-Bold / ArialMT / SimHei"),
          "size": P("number", "字号（pt），默认 24", default=24),
          "color": P("string", "文字颜色 #RRGGBB", default="#000000"),
          "x": P("number", "文字位置 X（像素）"),
          "y": P("number", "文字位置 Y（基线，像素）"),
          "align": P("string", "对齐", enum=["left", "center", "right"], default="left"),
          "opacity": P("integer", "图层不透明度 0-100"),
          "name": P("string", "图层名"),
          "index": P("integer", "文档序号，默认当前活动文档"),
      },
      ["content"])
def h_text(a):
    body = """
var o = $ARGS;
var d = __doc(o.index);
var l = d.artLayers.add();
l.kind = LayerKind.TEXT;
var nm = o.name ? String(o.name) : String(o.content === null || o.content === undefined ? "Text" : o.content);
if (nm.length > 40) nm = nm.substring(0, 40);
l.name = nm;
var t = l.textItem;
t.contents = String(o.content === null || o.content === undefined ? "" : o.content).replace(/\\r\\n/g, "\\r").replace(/\\n/g, "\\r");
var font_error = null;
if (o.font) { try { t.font = String(o.font); } catch (e) { font_error = String(e).substring(0, 160); } }
if (o.size) t.size = UnitValue(o.size, "pt");
t.color = __color(o.color || "#000000");
if (o.align) {
  var jm = { left: Justification.LEFT, center: Justification.CENTER, right: Justification.RIGHT };
  var jj = jm[String(o.align).toLowerCase()];
  if (jj !== undefined) t.justification = jj;
}
if ((o.x !== null && o.x !== undefined) || (o.y !== null && o.y !== undefined)) {
  t.position = [UnitValue(o.x || 0, "px"), UnitValue(o.y || 0, "px")];
}
if (o.opacity !== null && o.opacity !== undefined) l.opacity = o.opacity;
__r = __layerJson(l);
if (font_error) __r.font_error = font_error;
"""
    return run_jsx(render(body, a))


@tool("ps_fill", "用纯色填充整个图层（默认当前图层；无选区时自动全选）。",
      {
          "color": P("string", "填充色 #RRGGBB"),
          "name": P("string", "目标图层名"),
          "layer_index": P("integer", "目标图层序号"),
          "whole_layer": P("boolean", "true 填充整层；false 只填充当前选区", default=True),
          "index": P("integer", "文档序号，默认当前活动文档"),
      },
      ["color"])
def h_fill(a):
    body = """
var o = $ARGS;
var d = __doc(o.index);
if (o.name || o.layer_index !== null && o.layer_index !== undefined) {
  d.activeLayer = __layer(d, o.name, o.layer_index);
}
if (o.whole_layer !== false) d.selection.selectAll();
d.selection.fill(__color(o.color), ColorBlendMode.NORMAL, 100, false);
if (o.whole_layer !== false) d.selection.deselect();
__r = __layerJson(d.activeLayer);
"""
    return run_jsx(render(body, a))


@tool("ps_select", "选区操作：全选、取消选择、按矩形框选。",
      {
          "mode": P("string", "操作模式", enum=["all", "none", "rect"], default="all"),
          "x": P("number", "矩形选区 X"),
          "y": P("number", "矩形选区 Y"),
          "width": P("number", "矩形选区宽"),
          "height": P("number", "矩形选区高"),
          "index": P("integer", "文档序号，默认当前活动文档"),
      }, [])
def h_select(a):
    body = """
var o = $ARGS;
var d = __doc(o.index);
var mode = String(o.mode || "all");
if (mode === "all") {
  d.selection.selectAll();
} else if (mode === "none") {
  d.selection.deselect();
} else {
  var x = Number(o.x || 0), y = Number(o.y || 0), w = Number(o.width), h = Number(o.height);
  if (!(w > 0) || !(h > 0)) throw new Error("矩形选区需要 width / height");
  d.selection.select([[x, y], [x + w, y], [x + w, y + h], [x, y + h]], SelectionType.REPLACE, 0, true);
}
__r = __docJson(d, __docIndex(d));
"""
    return run_jsx(render(body, a))


@tool("ps_adjust", "图像调整：亮度对比度 / 色相饱和度 / 去色 / 反相 / 色阶 / 自动色阶 / 自动对比度 / 色调分离 / 阈值。作用于指定图层或当前图层。",
      {
          "kind": P("string", "调整类型",
                    enum=["brightness_contrast", "hue_saturation", "desaturate", "invert",
                          "levels", "auto_levels", "auto_contrast", "posterize", "threshold"],
                    default="brightness_contrast"),
          "brightness": P("number", "亮度 -150..150", default=0),
          "contrast": P("number", "对比度 -100..100", default=0),
          "hue": P("number", "色相 -180..180", default=0),
          "saturation": P("number", "饱和度 -100..100", default=0),
          "lightness": P("number", "明度 -100..100", default=0),
          "in_min": P("number", "色阶输入黑场 0-255", default=0),
          "in_max": P("number", "色阶输入白场 0-255", default=255),
          "gamma": P("number", "色阶中间调 0.1-9.99", default=1.0),
          "out_min": P("number", "色阶输出黑场 0-255", default=0),
          "out_max": P("number", "色阶输出白场 0-255", default=255),
          "levels": P("integer", "色调分离级数", default=4),
          "level": P("number", "阈值 0-255", default=128),
          "name": P("string", "目标图层名"),
          "layer_index": P("integer", "目标图层序号"),
          "index": P("integer", "文档序号，默认当前活动文档"),
      },
      ["kind"])
def h_adjust(a):
    body = """
var o = $ARGS;
var d = __doc(o.index);
if (o.name || (o.layer_index !== null && o.layer_index !== undefined)) {
  d.activeLayer = __layer(d, o.name, o.layer_index);
}
var target = d.activeLayer;
var kind = String(o.kind || "brightness_contrast");
if (kind === "brightness_contrast") {
  target.adjustBrightnessContrast(o.brightness || 0, o.contrast || 0);
} else if (kind === "hue_saturation") {
  var ad = new ActionDescriptor();
  ad.putBoolean(charIDToTypeID("Clrz"), false);
  var item = new ActionDescriptor();
  item.putInteger(charIDToTypeID("H   "), o.hue || 0);
  item.putInteger(charIDToTypeID("Strt"), o.saturation || 0);
  item.putInteger(charIDToTypeID("Lght"), o.lightness || 0);
  var list = new ActionList();
  list.putObject(charIDToTypeID("Hst2"), item);
  ad.putList(charIDToTypeID("Adjs"), list);
  executeAction(charIDToTypeID("HStr"), ad, DialogModes.NO);
} else if (kind === "desaturate") {
  target.desaturate();
} else if (kind === "invert") {
  target.invert();
} else if (kind === "levels") {
  target.adjustLevels(o.in_min === undefined ? 0 : o.in_min, o.in_max === undefined ? 255 : o.in_max,
                       o.gamma === undefined ? 1.0 : o.gamma,
                       o.out_min === undefined ? 0 : o.out_min, o.out_max === undefined ? 255 : o.out_max);
} else if (kind === "auto_levels") {
  target.autoLevels();
} else if (kind === "auto_contrast") {
  target.autoContrast();
} else if (kind === "posterize") {
  target.posterize(o.levels || 4);
} else if (kind === "threshold") {
  target.threshold(o.level === undefined ? 128 : o.level);
} else {
  throw new Error("不支持的调整类型: " + kind);
}
__r = __layerJson(target);
"""
    return run_jsx(render(body, a))


@tool("ps_filter", "滤镜：高斯模糊 / 动感模糊 / USM 锐化。作用于指定图层或当前图层。",
      {
          "kind": P("string", "滤镜类型", enum=["gaussian_blur", "motion_blur", "unsharp_mask"], default="gaussian_blur"),
          "radius": P("number", "半径（像素）", default=5),
          "angle": P("number", "动感模糊角度（度）", default=0),
          "amount": P("number", "USM 数量（%）", default=100),
          "threshold": P("number", "USM 阈值", default=0),
          "name": P("string", "目标图层名"),
          "layer_index": P("integer", "目标图层序号"),
          "index": P("integer", "文档序号，默认当前活动文档"),
      },
      ["kind"])
def h_filter(a):
    body = """
var o = $ARGS;
var d = __doc(o.index);
if (o.name || (o.layer_index !== null && o.layer_index !== undefined)) {
  d.activeLayer = __layer(d, o.name, o.layer_index);
}
var target = d.activeLayer;
var kind = String(o.kind || "gaussian_blur");
var radius = o.radius === undefined ? 5 : o.radius;
if (kind === "gaussian_blur") {
  var ok = false;
  try { target.applyGaussianBlur(radius); ok = true; } catch (e) {}
  if (!ok) {
    var desc = new ActionDescriptor();
    desc.putUnitDouble(charIDToTypeID("Rds "), charIDToTypeID("#Pxl"), radius);
    executeAction(charIDToTypeID("GsnB"), desc, DialogModes.NO);
  }
} else if (kind === "motion_blur") {
  var desc2 = new ActionDescriptor();
  desc2.putUnitDouble(charIDToTypeID("Angl"), charIDToTypeID("#Ang"), o.angle || 0);
  desc2.putUnitDouble(charIDToTypeID("Rds "), charIDToTypeID("#Pxl"), radius);
  executeAction(charIDToTypeID("MtnB"), desc2, DialogModes.NO);
} else if (kind === "unsharp_mask") {
  target.applyUnSharpMask(o.amount || 100, radius, o.threshold || 0);
} else {
  throw new Error("不支持的滤镜: " + kind);
}
__r = __layerJson(target);
"""
    return run_jsx(render(body, a))


@tool("ps_image_place", "把磁盘上的图片作为图层放入当前文档，可指定目标位置和尺寸。",
      {
          "path": P("string", "图片文件路径（png/jpg/tif/psd 等）"),
          "x": P("number", "左上角 X（像素），默认 0"),
          "y": P("number", "左上角 Y（像素），默认 0"),
          "width": P("number", "缩放后宽度（像素）"),
          "height": P("number", "缩放后高度（像素）"),
          "name": P("string", "图层名"),
          "index": P("integer", "文档序号，默认当前活动文档"),
      },
      ["path"])
def h_image_place(a):
    body = """
var o = $ARGS;
var d = __doc(o.index);
var f = new File(o.path);
if (!f.exists) throw new Error("文件不存在: " + o.path);
var src = app.open(f);
var sl = src.activeLayer;
var layer = sl.duplicate(d, ElementPlacement.PLACEATBEGINNING);
src.close(SaveOptions.DONOTSAVECHANGES);
app.activeDocument = d;
d.activeLayer = layer;
try { layer.isBackgroundLayer = false; } catch (e) {}
if (o.name) layer.name = String(o.name);
if (o.width || o.height) {
  var b = layer.bounds;
  var bw = b[2].as("px") - b[0].as("px");
  var bh = b[3].as("px") - b[1].as("px");
  var sx = o.width ? (o.width / bw) * 100 : 100;
  var sy = o.height ? (o.height / bh) * 100 : sx;
  layer.resize(sx, sy, AnchorPosition.TOPLEFT);
}
if ((o.x !== null && o.x !== undefined) || (o.y !== null && o.y !== undefined)) {
  var b2 = layer.bounds;
  layer.translate(UnitValue((o.x || 0) - b2[0].as("px"), "px"), UnitValue((o.y || 0) - b2[1].as("px"), "px"));
}
__r = __layerJson(layer);
"""
    return run_jsx(render(body, a))


@tool("ps_snapshot", "把文档导出成预览图 PNG，返回保存路径与（可选）base64 数据，供模型直接看图确认。",
      {
          "index": P("integer", "文档序号，默认当前活动文档"),
          "max_size": P("integer", "最长边像素上限，自动等比缩小，默认 1280", default=1280),
          "return_base64": P("boolean", "是否在结果里返回 base64 图片，默认 true", default=True),
          "path": P("string", "把预览图另存到指定路径（可选）"),
      }, [])
def h_snapshot(a):
    max_size = int(a.get("max_size") or 1280)
    tmp = os.path.join(TEMP_DIR, "snap_%d.png" % int(time.time() * 1000))
    body = """
var o = $ARGS;
var d = __doc(o.index);
var dup = d.duplicate("__snapshot__");
var w = Math.round(dup.width.as("px")), h = Math.round(dup.height.as("px"));
try { if (dup.mode !== DocumentMode.RGB) dup.changeMode(ChangeMode.RGB); } catch (e) {}
try { if (dup.bitsPerChannel !== BitsPerChannelType.EIGHT) dup.bitsPerChannel = BitsPerChannelType.EIGHT; } catch (e) {}
dup.flatten();
var po = new PNGSaveOptions();
po.compression = 6;
dup.saveAs(new File(o.path), po, true);
dup.close(SaveOptions.DONOTSAVECHANGES);
__r = { width: w, height: h, path: o.path };
"""
    data = run_jsx(render(body, {"index": a.get("index"), "path": tmp.replace("\\", "/")}))
    if not os.path.exists(tmp):
        raise PSError("预览图导出失败")

    final_path = tmp
    if a.get("path"):
        target = resolve_path(a.get("path"))
        parent = os.path.dirname(target)
        if parent and not os.path.isdir(parent):
            os.makedirs(parent, exist_ok=True)
        shutil.copyfile(tmp, target)
        final_path = target

    width, height = data.get("width"), data.get("height")
    if Image is not None and max_size:
        try:
            im = Image.open(final_path)
            if max(im.size) > max_size:
                ratio = max_size / float(max(im.size))
                size = (max(1, int(im.size[0] * ratio)), max(1, int(im.size[1] * ratio)))
                if im.mode not in ("RGB", "RGBA", "L"):
                    im = im.convert("RGB")
                im = im.resize(size, Image.LANCZOS)
                small = final_path[:-4] + "_preview.png"
                im.save(small, "PNG")
                final_path = small
        except Exception:
            pass

    result = {
        "path": final_path,
        "document_size": {"width": width, "height": height},
        "mime": "image/png",
    }
    with open(final_path, "rb") as fh:
        raw = fh.read()
    result["bytes"] = len(raw)
    if a.get("return_base64", True):
        result["image_base64"] = base64.b64encode(raw).decode("ascii")
        result["data_url"] = "data:image/png;base64," + result["image_base64"]
    return result


@tool("ps_jsx", "执行任意 ExtendScript(JSX) 代码，直接使用 app / app.activeDocument 等 Photoshop 对象。高级逃生舱。",
      {
          "code": P("string", "ExtendScript 代码；最后一条表达式的值会被返回"),
          "timeout": P("number", "超时秒数，默认 120", default=120),
      },
      ["code"], timeout=300)
def h_jsx(a):
    code = str(a.get("code") or "")
    if not code.strip():
        raise PSError("code 不能为空")
    timeout = float(a.get("timeout") or 120)
    raw = BRIDGE.call_jsx(PRELUDE + "\n;" + code, timeout)
    if raw is None:
        return {"value": None}
    text = str(raw)
    if len(text) > 20000:
        text = text[:20000] + "...(已截断)"
    return {"value": text}


# ---------------------------------------------------------------------------
# HTTP 服务
# ---------------------------------------------------------------------------
TOKEN = ""
QUIET = False


class APIHandler(BaseHTTPRequestHandler):
    server_version = "PhotoshopAPI/" + VERSION
    protocol_version = "HTTP/1.0"

    # -- 输出工具 ----------------------------------------------------------
    def _send(self, status, payload):
        body = json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization, X-API-Key")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.end_headers()
        try:
            self.wfile.write(body)
        except (BrokenPipeError, ConnectionResetError):
            pass

    def _read_body(self):
        length = int(self.headers.get("Content-Length") or 0)
        if length > 64 * 1024 * 1024:
            raise PSError("请求体过大（>64MB）")
        return self.rfile.read(length) if length > 0 else b""

    def _authorized(self):
        if not TOKEN:
            return True
        auth = self.headers.get("Authorization", "")
        return auth == "Bearer " + TOKEN or self.headers.get("X-API-Key") == TOKEN

    def _log(self, text):
        if not QUIET:
            print("[%s] %s" % (time.strftime("%H:%M:%S"), text), flush=True)

    # -- 路由 --------------------------------------------------------------
    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization, X-API-Key")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Content-Length", "0")
        self.end_headers()

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path.rstrip("/") or "/"
        query = parse_qs(parsed.query)
        if not self._authorized():
            return self._send(401, {"ok": False, "error": "未授权：需要 Bearer token"})
        if path == "/":
            return self._send(200, {
                "ok": True,
                "service": "photoshop-api",
                "version": VERSION,
                "photoshop": BRIDGE.status(),
                "tools": len(TOOLS),
                "endpoints": {
                    "GET /health": "服务与 Photoshop 连接状态（?connect=1 主动连接）",
                    "GET /tools": "OpenAI function-calling 工具清单",
                    "POST /call": '{"name": "...", "arguments": {...}}',
                    "POST /jsx": '{"code": "..."} 或纯文本 ExtendScript',
                },
                "cwd": os.getcwd(),
            })
        if path == "/health":
            force = str(query.get("connect", ["0"])[0]).lower() in ("1", "true", "yes", "on")
            info = {
                "ok": True,
                "service": "photoshop-api",
                "version": VERSION,
                "uptime_s": round(time.time() - START_TS, 1),
                "tools": len(TOOLS),
                "cwd": os.getcwd(),
                "photoshop": BRIDGE.status(),
            }
            if force:
                try:
                    connected = BRIDGE.probe()
                    info["photoshop"] = {"connected": True, **connected}
                except Exception as exc:
                    info["photoshop"] = {"connected": False, "error": str(exc)}
                    info["ok"] = False
            self._log("GET /health connected=%s" % info["photoshop"].get("connected"))
            return self._send(200, info)
        if path in ("/tools", "/tools/openai", "/v1/tools"):
            return self._send(200, {"tools": [t.schema() for t in TOOLS.values()]})
        return self._send(404, {"ok": False, "error": "未知路径: " + path})

    def do_POST(self):
        parsed = urlparse(self.path)
        path = parsed.path.rstrip("/") or "/"
        if not self._authorized():
            return self._send(401, {"ok": False, "error": "未授权：需要 Bearer token"})
        started = time.time()
        try:
            raw = self._read_body()
        except PSError as exc:
            return self._send(413, {"ok": False, "error": str(exc)})

        try:
            if path == "/call":
                data = json.loads(raw.decode("utf-8") or "{}")
                name = data.get("name") or data.get("tool")
                args = data.get("arguments") or data.get("args") or {}
                if not name:
                    return self._send(400, {"ok": False, "error": "缺少 name 字段"})
                if not isinstance(args, dict):
                    return self._send(400, {"ok": False, "error": "arguments 必须是对象"})
                entry = TOOLS.get(name)
                if entry is None:
                    return self._send(404, {"ok": False, "error": "未知工具: %s" % name,
                                            "available": sorted(TOOLS.keys())})
                missing = [key for key in entry.required if key not in args or args[key] is None]
                if missing:
                    return self._send(400, {"ok": False, "error": "缺少必填参数: " + ", ".join(missing),
                                            "required": list(entry.required)})
                result = entry.handler(args)
                elapsed = (time.time() - started) * 1000
                self._log("POST /call %-18s %6.0fms" % (name, elapsed))
                return self._send(200, {"ok": True, "tool": name, "result": result,
                                        "elapsed_ms": round(elapsed, 1)})

            if path == "/jsx":
                ctype = (self.headers.get("Content-Type") or "").lower()
                if "application/json" in ctype:
                    data = json.loads(raw.decode("utf-8") or "{}")
                    code = data.get("code") or ""
                    timeout = float(data.get("timeout") or 120)
                else:
                    code = raw.decode("utf-8")
                    timeout = 120.0
                if not code.strip():
                    return self._send(400, {"ok": False, "error": "code 不能为空"})
                value = BRIDGE.call_jsx(PRELUDE + "\n;" + code, timeout)
                elapsed = (time.time() - started) * 1000
                self._log("POST /jsx %6.0fms" % elapsed)
                return self._send(200, {"ok": True, "value": None if value is None else str(value),
                                        "elapsed_ms": round(elapsed, 1)})

            if path == "/connect":
                connected = BRIDGE.probe()
                return self._send(200, {"ok": True, "photoshop": {"connected": True, **connected}})

            return self._send(404, {"ok": False, "error": "未知路径: " + path})
        except PSError as exc:
            self._log("POST %s error: %s" % (path, str(exc)[:200]))
            return self._send(400, {"ok": False, "error": str(exc)})
        except json.JSONDecodeError as exc:
            return self._send(400, {"ok": False, "error": "JSON 解析失败: %s" % exc})
        except Exception as exc:  # noqa: BLE001 - 服务不能因单次调用崩溃
            traceback.print_exc()
            return self._send(500, {"ok": False, "error": "%s: %s" % (type(exc).__name__, exc)})

    def log_message(self, fmt, *args):  # 用简洁日志替代默认访问日志
        pass


def make_server(host, port):
    ThreadingHTTPServer.allow_reuse_address = True
    httpd = ThreadingHTTPServer((host, port), APIHandler)
    httpd.daemon_threads = True
    return httpd


# ---------------------------------------------------------------------------
# 端到端自检
# ---------------------------------------------------------------------------
def selftest():
    global BRIDGE
    if not HAS_COM:
        print("缺少 pywin32，请先: python -m pip install pywin32", file=sys.stderr)
        return 1
    BRIDGE = PSBridge(launch=True)
    here = os.path.dirname(os.path.abspath(__file__))
    preview = os.path.join(here, "selftest-preview.png")

    def step(name, args):
        t0 = time.time()
        res = TOOLS[name].handler(args)
        print("  [ok] %-18s %6.0fms" % (name, (time.time() - t0) * 1000))
        return res

    try:
        print("== Photoshop API 自检 ==")
        st = step("ps_status", {})
        print("  Photoshop:", st.get("version"), "| 已打开文档:", len(st.get("documents") or []))
        doc = step("ps_doc_create", {"width": 1200, "height": 800, "background": "white", "name": "photoshop-api-demo"})
        print("  画布:", doc["width"], "x", doc["height"])

        step("ps_shape", {"shape": "rect", "x": 0, "y": 0, "width": 1200, "height": 170, "color": "#0E1726", "name": "banner"})
        step("ps_text", {"content": "Photoshop × 模型 API", "font": "MicrosoftYaHei-Bold", "size": 58,
                         "color": "#FFFFFF", "x": 64, "y": 104, "name": "title"})
        step("ps_text", {"content": "本地 Photoshop 2026 · HTTP JSON 接口 · ExtendScript 执行层", "font": "MicrosoftYaHei",
                         "size": 22, "color": "#8FA3BF", "x": 66, "y": 146, "name": "subtitle"})

        cards = [("24", "工具接口", "#2F6FEB", 64), ("AI", "模型可直接调用", "#FF7A18", 434), ("PSD", "真实文档落盘", "#0E1726", 804)]
        for num, label, accent, x in cards:
            step("ps_shape", {"shape": "rect", "x": x, "y": 230, "width": 330, "height": 280, "color": "#F4F7FC", "name": "card%d" % (x // 10)})
            step("ps_shape", {"shape": "rect", "x": x, "y": 230, "width": 6, "height": 280, "color": accent, "name": "accent%d" % (x // 10)})
            step("ps_text", {"content": num, "font": "MicrosoftYaHei-Bold", "size": 52, "color": "#0E1726",
                             "x": x + 36, "y": 360, "name": "num%d" % (x // 10)})
            step("ps_text", {"content": label, "font": "MicrosoftYaHei", "size": 20, "color": "#5A6B85",
                             "x": x + 36, "y": 404, "name": "label%d" % (x // 10)})

        step("ps_shape", {"shape": "ellipse", "x": 880, "y": 520, "width": 240, "height": 240, "color": "#2F6FEB", "opacity": 90, "name": "circle-blue"})
        step("ps_shape", {"shape": "ellipse", "x": 980, "y": 620, "width": 160, "height": 160, "color": "#FF7A18", "opacity": 85, "name": "circle-orange"})
        step("ps_text", {"content": "POST /call  ·  GET /tools  ·  ps_snapshot 看图验收", "font": "MicrosoftYaHei",
                         "size": 20, "color": "#5A6B85", "x": 64, "y": 720, "name": "footer"})

        snap = step("ps_snapshot", {"max_size": 1280, "return_base64": False, "path": preview})
        print("  预览图:", snap["path"], snap["bytes"], "bytes")
        layers = step("ps_layers_list", {})
        print("  图层数:", len(layers["layers"]))
        step("ps_doc_close", {})
        print("\n自检通过。预览图: %s" % preview)
        return 0
    except PSError as exc:
        print("\n[自检失败] %s" % exc, file=sys.stderr)
        return 1


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------
def main(argv=None):
    global TOKEN, QUIET, BRIDGE
    ap = argparse.ArgumentParser(
        description="让模型通过 HTTP/JSON 接口操作本机 Photoshop",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="示例: python photoshop_api.py --port 8765\n      python photoshop_api.py --selftest",
    )
    ap.add_argument("--host", default="127.0.0.1", help="监听地址，默认 127.0.0.1")
    ap.add_argument("--port", type=int, default=8765, help="监听端口，默认 8765")
    ap.add_argument("--token", default=os.environ.get("PS_API_TOKEN", ""), help="访问令牌（可选）")
    ap.add_argument("--no-launch", action="store_true", help="只连接已运行的 Photoshop，不主动启动")
    ap.add_argument("--no-warmup", action="store_true", help="启动时不预热连接 Photoshop")
    ap.add_argument("--quiet", action="store_true", help="不打印每次调用日志")
    ap.add_argument("--selftest", action="store_true", help="跑一遍端到端自检并生成预览图，不启动服务")
    args = ap.parse_args(argv)

    QUIET = args.quiet
    TOKEN = args.token or ""

    if not HAS_COM:
        print("缺少 pywin32：请先运行  python -m pip install pywin32", file=sys.stderr)
        return 1

    if args.selftest:
        return selftest()

    BRIDGE = PSBridge(launch=not args.no_launch)
    try:
        httpd = make_server(args.host, args.port)
    except OSError as exc:
        print("端口 %s:%s 无法监听：%s" % (args.host, args.port, exc), file=sys.stderr)
        print("换个端口试试：python photoshop_api.py --port 8766", file=sys.stderr)
        return 1

    url = "http://%s:%d" % (args.host, args.port)
    print("=" * 62)
    print(" Photoshop API %s   %s" % (VERSION, url))
    print("   GET  %s/tools     模型工具清单" % url)
    print("   POST %s/call      {\"name\", \"arguments\"}" % url)
    print("   POST %s/jsx       原始 ExtendScript" % url)
    print("   GET  %s/health    健康检查" % url)
    print("   工具 %d 个 | 工作目录 %s" % (len(TOOLS), os.getcwd()))
    print("=" * 62, flush=True)

    if not args.no_warmup:
        def warmup():
            t0 = time.time()
            try:
                info = BRIDGE.probe()
                print("[warmup] 已连接 Photoshop %s（文档 %s 个，%.1fs）"
                      % (info.get("version"), info.get("documents"), time.time() - t0), flush=True)
            except Exception as exc:
                print("[warmup] 连接 Photoshop 失败：%s" % exc, flush=True)

        threading.Thread(target=warmup, name="warmup", daemon=True).start()

    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止")
    finally:
        httpd.server_close()
    return 0


if __name__ == "__main__":
    sys.exit(main())

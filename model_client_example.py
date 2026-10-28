#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
model_client_example.py —— 用任意 OpenAI 兼容模型（function calling）驱动本地 Photoshop

准备
----
1) 先启动 API 服务（另开一个终端）：
       python photoshop_api.py --port 8765

2) 配置模型（DeepSeek / OpenAI / 任何兼容 /v1/chat/completions 的服务都行）：
   Windows:
       set OPENAI_BASE_URL=https://api.deepseek.com/v1
       set OPENAI_API_KEY=sk-xxxxxxxx
       set OPENAI_MODEL=deepseek-chat
   PowerShell:
       $env:OPENAI_BASE_URL="https://api.deepseek.com/v1"
       $env:OPENAI_API_KEY="sk-xxxxxxxx"
       $env:OPENAI_MODEL="deepseek-chat"

运行
----
    python model_client_example.py --check        # 只检查 Photoshop API 是否连通
    python model_client_example.py "新建 1200x800 白色画布，写标题'你好'，导出 PNG 到 D:\\out\\a.png"

说明
----
* 只依赖标准库，不需要安装 openai SDK。
* ps_snapshot 返回的 base64 图片会从工具结果里剔除（避免撑爆上下文），
  只保留落盘路径；如果你用视觉模型，可以自己把 base64 拼成 image_url 消息。
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import urllib.error
import urllib.request

PS_API = os.environ.get("PS_API_URL", "http://127.0.0.1:8765").rstrip("/")
BASE_URL = os.environ.get("OPENAI_BASE_URL", "https://api.openai.com/v1").rstrip("/")
API_KEY = os.environ.get("OPENAI_API_KEY", "")
MODEL = os.environ.get("OPENAI_MODEL", "gpt-4o-mini")
MAX_STEPS = int(os.environ.get("MAX_STEPS", "20"))

SYSTEM_PROMPT = """你是 Photoshop 操作员，通过给定工具操作本机已运行的 Photoshop。
规则：
1. 先看清当前状态：不确定时调用 ps_status / ps_doc_info / ps_layers_list。
2. 文字、颜色、坐标都使用工具参数，不要臆造不存在的工具。
3. 中文字体推荐 MicrosoftYaHei 或 MicrosoftYaHei-Bold；西文可用 ArialMT。
4. 坐标单位是文档像素，原点在左上角；ps_text 的 x/y 是文字基线左上点。
5. 完成后用 ps_snapshot 导出预览图做自检；需要落盘时用 ps_doc_save。
6. 一次只做一步，拿到工具结果再继续；失败时读错误信息并修正参数。
7. 最终用简短中文说明你做了什么、文件保存在哪里。"""


def http_json(url, payload=None, headers=None, timeout=300):
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8") if payload is not None else None
    req = urllib.request.Request(url, data=data, headers=headers or {}, method="POST" if data else "GET")
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8"))


# ---------------------------------------------------------------- Photoshop API
def ps_health():
    return http_json(PS_API + "/health?connect=1", timeout=180)


def ps_tools():
    return http_json(PS_API + "/tools", timeout=60)["tools"]


def ps_call(name, arguments):
    try:
        return http_json(PS_API + "/call", {"name": name, "arguments": arguments or {}})
    except urllib.error.HTTPError as exc:
        return json.loads(exc.read().decode("utf-8"))
    except urllib.error.URLError as exc:
        raise SystemExit("连不上 Photoshop API（%s）：请先运行 python photoshop_api.py\n%s"
                         % (PS_API, exc))


# ---------------------------------------------------------------------- 模型
def chat(messages, tools):
    payload = {
        "model": MODEL,
        "messages": messages,
        "tools": tools,
        "tool_choice": "auto",
        "temperature": 0,
    }
    headers = {"Content-Type": "application/json"}
    if API_KEY:
        headers["Authorization"] = "Bearer " + API_KEY
    return http_json(BASE_URL + "/chat/completions", payload, headers, timeout=300)


def text_of(content):
    """OpenAI 字符串或 Anthropic 内容块数组 → 纯文本（只取正文块）。"""
    if isinstance(content, list):
        parts = []
        for b in content:
            if isinstance(b, dict) and b.get("type") in ("text", "output_text", "input_text"):
                parts.append(str(b.get("text") or ""))
        return "\n".join(p for p in parts if p)
    return str(content or "")


def reasoning_of(msg):
    """取出思考内容：reasoning_content / reasoning / thinking 块。"""
    parts = []
    for key in ("reasoning_content", "reasoning"):
        val = msg.get(key)
        if isinstance(val, str) and val.strip():
            parts.append(val.strip())
    content = msg.get("content")
    if isinstance(content, list):
        for b in content:
            if isinstance(b, dict) and b.get("type") in ("thinking", "redacted_thinking", "reasoning"):
                tx = b.get("thinking") or b.get("text")
                if tx:
                    parts.append(str(tx))
    return "\n\n".join(parts)


def shrink_tool_result(result):
    """去掉 base64 图片，避免上下文爆炸；保留路径和尺寸。"""
    if isinstance(result, dict) and result.get("data_url"):
        slim = {k: v for k, v in result.items() if k not in ("image_base64", "data_url")}
        slim["note"] = "预览图已保存到 path，如需看图请让用户打开该文件"
        return slim
    return result


def run_task(task):
    tools = ps_tools()
    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": task},
    ]
    print("模型: %s   Photoshop API: %s   工具: %d 个\n" % (MODEL, PS_API, len(tools)))

    for step in range(1, MAX_STEPS + 1):
        resp = chat(messages, tools)
        msg = resp["choices"][0].get("message") or {}
        think = reasoning_of(msg)
        if think:
            print("思考：\n" + think + "\n")
        # 只回传标准字段：reasoning_content 等不该进上下文，部分服务端会直接 400
        clean = {k: msg[k] for k in ("role", "content", "tool_calls") if k in msg}
        if clean.get("content") is None:
            clean["content"] = ""
        messages.append(clean)

        calls = msg.get("tool_calls") or []
        if not calls:
            print("模型最终答复：\n%s" % text_of(msg.get("content")).strip())
            return 0

        for call in calls:
            fn = call.get("function") or {}
            name = fn.get("name")
            try:
                args = json.loads(fn.get("arguments") or "{}")
            except json.JSONDecodeError:
                args = {}
            print("  [%d] %s %s" % (step, name, json.dumps(args, ensure_ascii=False)[:160]))
            body = ps_call(name, args)
            if body.get("ok"):
                result = shrink_tool_result(body.get("result"))
                content = json.dumps(result, ensure_ascii=False)
                print("      -> ok  %.0fms  %s" % (body.get("elapsed_ms", 0), content[:140].replace("\n", " ")))
            else:
                content = json.dumps(body, ensure_ascii=False)
                print("      -> 失败  %s" % content[:200])
            if len(content) > 12000:
                content = content[:12000] + " ...(已截断)"
            messages.append({"role": "tool", "tool_call_id": call.get("id"), "content": content})

    print("达到最大步数 %d，未完成任务" % MAX_STEPS)
    return 1


def check():
    print("Photoshop API:", PS_API)
    health = ps_health()
    ps = health.get("photoshop", {})
    print("  服务版本 :", health.get("version"))
    print("  已连接   :", ps.get("connected"), "| Photoshop", ps.get("version"))
    print("  工具数量 :", health.get("tools"))
    tools = ps_tools()
    print("  工具清单 :", ", ".join(t["function"]["name"] for t in tools))
    status = ps_call("ps_status", {})
    if status.get("ok"):
        docs = status["result"].get("documents") or []
        print("  ps_status: ok，当前文档 %d 个" % len(docs))
        for d in docs:
            print("    -", d["name"], "%sx%s" % (d["width"], d["height"]), d["mode"])
    else:
        print("  ps_status: 失败 ->", status.get("error"))
        return 1
    print("\n连通性正常。配置 OPENAI_* 环境变量后即可让模型直接操作 Photoshop。")
    return 0


def main():
    ap = argparse.ArgumentParser(description="用 OpenAI 兼容模型驱动本地 Photoshop")
    ap.add_argument("task", nargs="*", help="自然语言任务描述")
    ap.add_argument("--check", action="store_true", help="只检查 Photoshop API 连通性")
    args = ap.parse_args()

    if args.check:
        return check()
    task = " ".join(args.task).strip()
    if not task:
        task = "新建 1000x700 白色画布，画一个蓝色圆和一个橙色矩形，写一行标题，导出 PNG 到 D:\\PS\\demo.png"
        print("未提供任务，使用默认示例：%s\n" % task)
    if not API_KEY:
        print("提示：未设置 OPENAI_API_KEY，模型服务可能拒绝请求（本地 vLLM/Ollama 等可忽略）\n")
    return run_task(task)


if __name__ == "__main__":
    sys.exit(main())

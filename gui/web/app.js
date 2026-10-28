/* Photoshop Studio —— 前端逻辑（原生 JS，无框架） */
(() => {
  "use strict";

  // ---------------------------------------------------------------- 基础
  const $ = (sel, root = document) => root.querySelector(sel);
  const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));
  const MOCK = /[?&]mock=1/.test(location.search);
  const API = location.protocol === "file:" ? "http://127.0.0.1:8765" : location.origin;

  const settings = Object.assign(
    { base_url: "https://api.openai.com/v1", api_key: "", model: "gpt-4o-mini" },
    JSON.parse(localStorage.getItem("psstudio.settings") || "{}")
  );

  const state = {
    tools: [],
    docs: [],
    activeDoc: null,
    layers: [],
    previewTimer: null,
    busy: false,
    messages: [],
    config: { pictures: "" },
    startedAt: Date.now(),
  };

  const esc = (s) =>
    String(s == null ? "" : s).replace(/[&<>"']/g, (c) =>
      ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])
    );

  const stamp = () => new Date().toLocaleTimeString("zh-CN", { hour12: false });

  // ---------------------------------------------------------------- 提示
  function toast(text, kind = "info", ms = 2600) {
    const el = document.createElement("div");
    el.className = "toast " + kind;
    el.textContent = text;
    $("#toasts").appendChild(el);
    setTimeout(() => {
      el.classList.add("leaving");
      setTimeout(() => el.remove(), 220);
    }, ms);
  }

  function logLine(text, kind = "info") {
    const el = document.createElement("div");
    el.className = "log-line " + kind;
    el.innerHTML = `<span class="t">${stamp()}</span><span class="m">${esc(text)}</span>`;
    const box = $("#log");
    box.appendChild(el);
    box.scrollTop = box.scrollHeight;
  }

  // ---------------------------------------------------------------- 状态
  function setStatus(kind, text) {
    $("#status").dataset.state = kind;
    $("#statusText").textContent = text;
  }

  function setBusy(on) {
    state.busy = on;
    if (on) setStatus("busy", "处理中…");
  }

  // ---------------------------------------------------------------- API
  async function api(name, args = {}) {
    const t0 = performance.now();
    let body;
    try {
      const res = await fetch(API + "/call", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name, arguments: args }),
      });
      body = await res.json();
    } catch (err) {
      logLine(`${name} 连接失败：${err.message}`, "bad");
      setStatus("bad", "API 未连接");
      throw err;
    }
    const ms = Math.round(performance.now() - t0);
    if (!body.ok) {
      logLine(`${name} 失败 ${ms}ms · ${body.error}`, "bad");
      throw new Error(body.error || "调用失败");
    }
    logLine(`${name} ${ms}ms`, "ok");
    return body.result;
  }

  async function apiQuiet(name, args = {}) {
    try {
      return await api(name, args);
    } catch {
      return null;
    }
  }

  // ---------------------------------------------------------------- 模型代理
  async function llmChat(messages, tools) {
    const res = await fetch(API + "/llm/chat", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        base_url: settings.base_url,
        api_key: settings.api_key,
        model: settings.model,
        messages,
        tools,
      }),
    });
    const body = await res.json();
    if (!body.ok) throw new Error(body.error || "模型请求失败");
    return body;
  }

  // ---------------------------------------------------------------- 刷新
  async function refresh({ preview = true } = {}) {
    try {
      const health = await fetch(API + "/health?connect=1").then((r) => r.json());
      if (health.photoshop && health.photoshop.connected) {
        setStatus("ok", `Photoshop ${health.photoshop.version || ""}`.trim());
        $("#meta").textContent = `${health.photoshop.documents ?? 0} 文档 · ${health.tools} 工具`;
      } else {
        setStatus("bad", "Photoshop 未连接");
        $("#meta").textContent = "—";
      }
      state.docs = (health.photoshop && health.photoshop.documents_list) || [];
    } catch {
      setStatus("bad", "API 未连接");
    }

    const status = await apiQuiet("ps_status");
    if (status) {
      state.docs = status.documents || [];
      state.activeDoc = state.docs.find((d) => d.active) || state.docs[0] || null;
      $("#meta").textContent = `${state.docs.length} 文档 · ${state.tools.length} 工具`;
      setStatus("ok", `Photoshop ${status.version}`);
    }

    await refreshDoc(status);
    if (preview) schedulePreview(0);
  }

  async function refreshDoc(status) {
    if (!status) status = await apiQuiet("ps_status");
    if (status) {
      state.docs = status.documents || [];
      state.activeDoc = state.docs.find((d) => d.active) || state.docs[0] || null;
    }
    const doc = state.activeDoc;
    if (!doc) {
      $("#docName").textContent = "预览";
      $("#docMeta").textContent = "—";
      $("#layerList").innerHTML = '<div class="empty">还没有打开的文档</div>';
      $("#layerCount").textContent = "0";
      $("#preview").hidden = true;
      $("#stageEmpty").hidden = false;
      return;
    }
    $("#docName").textContent = doc.name || "未命名";
    $("#docMeta").textContent =
      `${doc.width} × ${doc.height} px · ${doc.resolution || 72} dpi · ${doc.mode || "RGB"} · ${doc.layer_count ?? 0} 图层`;
    const layers = await apiQuiet("ps_layers_list");
    if (layers && layers.layers) {
      state.layers = layers.layers;
      renderLayers(state.layers);
    }
  }

  function renderLayers(layers) {
    const box = $("#layerList");
    const flat = [];
    const walk = (list, depth) => {
      list.forEach((l) => {
        flat.push(Object.assign({}, l, { depth }));
        if (l.layers) walk(l.layers, depth + 1);
      });
    };
    walk(layers || [], 0);
    $("#layerCount").textContent = String(flat.length);
    if (!flat.length) {
      box.innerHTML = '<div class="empty">这个文档还没有图层</div>';
      return;
    }
    box.innerHTML = flat
      .map(
        (l) => `
      <div class="layer ${l.visible ? "" : "off"}" data-name="${esc(l.name)}" style="margin-left:${l.depth * 12}px">
        <button class="layer-eye" data-name="${esc(l.name)}" data-visible="${l.visible ? 1 : 0}" title="显示/隐藏">
          <svg class="ico"><use href="#${l.visible ? "i-eye" : "i-eye-off"}"/></svg>
        </button>
        <span class="layer-name" title="${esc(l.name)}">${esc(l.name)}</span>
        ${l.kind && l.kind !== "normal" ? `<span class="layer-kind">${esc(l.kind)}</span>` : ""}
        <span class="layer-op">${l.opacity != null ? l.opacity : 100}%</span>
      </div>`
      )
      .join("");

    $$(".layer", box).forEach((row) => {
      row.addEventListener("click", async (ev) => {
        if (ev.target.closest(".layer-eye")) return;
        await apiQuiet("ps_layer_set", { name: row.dataset.name, select: true });
        $$(".layer", box).forEach((r) => r.classList.toggle("active", r === row));
      });
    });

    $$(".layer-eye", box).forEach((btn) => {
      btn.addEventListener("click", async (ev) => {
        ev.stopPropagation();
        const visible = btn.dataset.visible === "1";
        const res = await apiQuiet("ps_layer_set", { name: btn.dataset.name, visible: !visible });
        if (res) refreshDoc();
      });
    });
  }

  // ---------------------------------------------------------------- 预览
  function schedulePreview(delay = 260) {
    if (MOCK) return;
    clearTimeout(state.previewTimer);
    state.previewTimer = setTimeout(takeSnapshot, delay);
  }

  async function takeSnapshot() {
    if (!state.activeDoc) return;
    const img = $("#preview");
    img.classList.add("loading");
    try {
      const snap = await api("ps_snapshot", { max_size: 1400, return_base64: true });
      img.src = snap.data_url || ("data:image/png;base64," + snap.image_base64);
      img.hidden = false;
      $("#stageEmpty").hidden = true;
      $("#snapTime").textContent = "预览 " + stamp();
    } catch (err) {
      toast("预览失败：" + err.message, "bad");
    } finally {
      img.classList.remove("loading");
    }
  }

  // ---------------------------------------------------------------- 画布
  const PRESETS = [
    { label: "1080² 社媒", w: 1080, h: 1080 },
    { label: "1920×1080", w: 1920, h: 1080 },
    { label: "A4 · 300dpi", w: 2480, h: 3508, dpi: 300 },
    { label: "800×800 头像", w: 800, h: 800 },
    { label: "2560×1440", w: 2560, h: 1440 },
  ];

  function renderPresets() {
    $("#presets").innerHTML = PRESETS.map(
      (p, i) => `<button class="chip" data-preset="${i}">${p.label}</button>`
    ).join("");
    $$("#presets .chip").forEach((btn) => {
      btn.addEventListener("click", () => {
        const p = PRESETS[Number(btn.dataset.preset)];
        $("#docW").value = p.w;
        $("#docH").value = p.h;
        $("#docDpi").value = p.dpi || 72;
      });
    });
  }

  async function createDoc() {
    const width = Number($("#docW").value) || 1200;
    const height = Number($("#docH").value) || 800;
    const resolution = Number($("#docDpi").value) || 72;
    const background = $("#docBg").value;
    const color = $("#docColor").value;
    try {
      setBusy(true);
      const doc = await api("ps_doc_create", { width, height, resolution, background, color });
      toast(`已创建 ${doc.width}×${doc.height} 画布`, "ok");
      await refresh();
    } catch (err) {
      toast(err.message, "bad");
    } finally {
      setBusy(false);
      refreshStatusOnly();
    }
  }

  // ---------------------------------------------------------------- 快捷操作
  const QUICK = [
    { label: "▭ 矩形", run: () => fillTool("ps_shape", { shape: "rect", x: 100, y: 100, width: 320, height: 200, color: "#2F6FEB" }) },
    { label: "◯ 椭圆", run: () => fillTool("ps_shape", { shape: "ellipse", x: 120, y: 120, width: 260, height: 260, color: "#FF7A18" }) },
    { label: "T 文字", run: () => fillTool("ps_text", { content: "标题文字", font: "MicrosoftYaHei-Bold", size: 48, color: "#0E1726", x: 100, y: 260 }) },
    { label: "高斯模糊", run: () => runQuick("ps_filter", { kind: "gaussian_blur", radius: 6 }) },
    { label: "USM 锐化", run: () => runQuick("ps_filter", { kind: "unsharp_mask", amount: 80, radius: 2 }) },
    { label: "去色", run: () => runQuick("ps_adjust", { kind: "desaturate" }) },
    { label: "反色", run: () => runQuick("ps_adjust", { kind: "invert" }) },
    { label: "拼合图层", run: () => runQuick("ps_doc_flatten", {}) },
  ];

  function renderQuick() {
    $("#quick").innerHTML = QUICK.map((q, i) => `<button class="chip" data-quick="${i}">${q.label}</button>`).join("");
    $$("#quick .chip").forEach((btn) => {
      btn.addEventListener("click", () => QUICK[Number(btn.dataset.quick)].run());
    });
  }

  async function runQuick(name, args) {
    try {
      setBusy(true);
      await api(name, args);
      toast("已执行 " + name, "ok");
      schedulePreview(120);
      await refreshDoc();
    } catch (err) {
      toast(err.message, "bad");
    } finally {
      setBusy(false);
      refreshStatusOnly();
    }
  }

  // ---------------------------------------------------------------- 工具直调
  function fillTool(name, args) {
    selectTab("tool");
    const select = $("#toolSelect");
    if (select.value !== name) select.value = name;
    onToolChange();
    $("#toolArgs").value = JSON.stringify(args, null, 2);
    toast("参数已填入，可修改后执行");
  }

  function schemaDefaults(tool) {
    const props = (tool && tool.parameters && tool.parameters.properties) || {};
    const out = {};
    Object.entries(props).forEach(([key, p]) => {
      if (p.default !== undefined) out[key] = p.default;
      else if (p.enum) out[key] = p.enum[0];
      else if (p.type === "integer" || p.type === "number") out[key] = 0;
      else if (p.type === "boolean") out[key] = false;
      else out[key] = "";
    });
    return out;
  }

  function onToolChange() {
    const tool = state.tools.find((t) => t.function.name === $("#toolSelect").value);
    const fn = tool ? tool.function : null;
    $("#toolDesc").textContent = fn ? fn.description : "—";
    $("#toolArgs").value = JSON.stringify(schemaDefaults(fn), null, 2);
  }

  function renderTools() {
    const select = $("#toolSelect");
    select.innerHTML = state.tools
      .map((t) => `<option value="${esc(t.function.name)}">${esc(t.function.name)}</option>`)
      .join("");
    select.addEventListener("change", onToolChange);
    onToolChange();
  }

  async function runTool() {
    const name = $("#toolSelect").value;
    let args = {};
    try {
      args = JSON.parse($("#toolArgs").value || "{}");
    } catch (err) {
      toast("参数不是合法 JSON", "bad");
      return;
    }
    $("#toolResult").textContent = "执行中…";
    try {
      setBusy(true);
      const res = await api(name, args);
      $("#toolResult").textContent = JSON.stringify(res, null, 2);
      toast(name + " 执行成功", "ok");
      schedulePreview(120);
      await refreshDoc();
    } catch (err) {
      $("#toolResult").textContent = "错误：" + err.message;
      toast(err.message, "bad");
    } finally {
      setBusy(false);
      refreshStatusOnly();
    }
  }

  // ---------------------------------------------------------------- 导出
  function guessDir() {
    return state.config.pictures || "D:\\PS";
  }

  function openExport(kind) {
    const d = new Date();
    const pad = (n) => String(n).padStart(2, "0");
    const ts = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}`;
    const ext = kind === "psd" ? "psd" : "png";
    const def = `${guessDir()}\\ps-studio-${ts}.${ext}`;
    modal({
      title: kind === "psd" ? "保存 PSD" : "导出 PNG",
      fields: [{ key: "path", label: "文件路径", value: def }],
      confirmText: "保存",
      icon: kind === "psd" ? "i-save" : "i-camera",
    }).then(async (values) => {
      if (!values) return;
      try {
        setBusy(true);
        const res = await api("ps_doc_save", { path: values.path, format: ext, as_copy: true });
        toast(`已保存 ${res.bytes} 字节`, "ok");
        logLine(`保存到 ${res.path}`, "ok");
      } catch (err) {
        toast(err.message, "bad");
      } finally {
        setBusy(false);
        refreshStatusOnly();
      }
    });
  }

  // ---------------------------------------------------------------- 设置
  function openSettings() {
    modal({
      title: "设置",
      icon: "i-settings",
      fields: [
        { key: "base_url", label: "模型接口 Base URL", value: settings.base_url, placeholder: "https://api.openai.com/v1" },
        { key: "api_key", label: "API Key", value: settings.api_key, type: "password", placeholder: "sk-…" },
        { key: "model", label: "模型名", value: settings.model, placeholder: "deepseek-chat / gpt-4o-mini" },
      ],
      note: "AI 对话通过本地服务代理请求该接口（避免浏览器跨域）。Key 仅保存在本机 localStorage。",
      confirmText: "保存",
    }).then((values) => {
      if (!values) return;
      settings.base_url = values.base_url.trim();
      settings.api_key = values.api_key.trim();
      settings.model = values.model.trim();
      localStorage.setItem("psstudio.settings", JSON.stringify(settings));
      toast("设置已保存", "ok");
    });
  }

  // ---------------------------------------------------------------- 弹窗
  function modal({ title, fields = [], note = "", confirmText = "确定", icon = "" }) {
    return new Promise((resolve) => {
      const root = $("#modalRoot");
      const body = fields
        .map(
          (f) => `<label>${esc(f.label)}
            <input data-key="${esc(f.key)}" type="${f.type || "text"}" value="${esc(f.value || "")}"
                   placeholder="${esc(f.placeholder || "")}" ${f.type === "password" ? 'autocomplete="off"' : ""}>
          </label>`
        )
        .join("");
      root.innerHTML = `
        <div class="modal-backdrop"></div>
        <div class="modal" role="dialog" aria-modal="true">
          <div class="modal-h"><h3>${esc(title)}</h3></div>
          <div class="modal-body">
            ${icon ? `<div class="row"><svg class="ico" style="width:18px;height:18px;color:var(--accent)"><use href="#${icon}"/></svg>
              <span class="note">${esc(note)}</span></div>` : ""}
            ${body}
            ${!icon && note ? `<p class="note">${esc(note)}</p>` : ""}
          </div>
          <div class="modal-f">
            <button class="btn ghost" data-act="cancel">取消</button>
            <button class="btn primary" data-act="ok">${esc(confirmText)}</button>
          </div>
        </div>`;
      root.classList.add("open");
      const close = (value) => {
        root.classList.remove("open");
        root.innerHTML = "";
        resolve(value);
      };
      const collect = () => {
        const out = {};
        $$("input[data-key]", root).forEach((inp) => (out[inp.dataset.key] = inp.value));
        return out;
      };
      root.querySelector('[data-act="ok"]').addEventListener("click", () => close(collect()));
      root.querySelector('[data-act="cancel"]').addEventListener("click", () => close(null));
      root.querySelector(".modal-backdrop").addEventListener("click", () => close(null));
      root.addEventListener("keydown", (e) => {
        if (e.key === "Escape") close(null);
        if (e.key === "Enter" && e.target.tagName === "INPUT") close(collect());
      });
      const first = root.querySelector("input");
      if (first) first.focus();
    });
  }

  // ---------------------------------------------------------------- 对话
  function chatAdd(role, html) {
    const empty = $("#chatEmpty");
    if (empty) empty.remove();
    const wrap = document.createElement("div");
    wrap.className = "msg " + role;
    const label = { user: "你", assistant: "模型", tool: "工具" }[role] || role;
    wrap.innerHTML = `<span class="who">${label}</span><div class="bubble">${html}</div>`;
    const box = $("#chat");
    box.appendChild(wrap);
    box.scrollTop = box.scrollHeight;
    return wrap;
  }

  function typingBubble() {
    const wrap = chatAdd("assistant", '<span class="typing"><i></i><i></i><i></i></span>');
    const box = $("#chat");
    box.scrollTop = box.scrollHeight;
    return wrap;
  }

  function summarize(res) {
    if (!res || typeof res !== "object") return "";
    const keys = ["name", "path", "width", "height", "layer_count", "bytes", "document"];
    const parts = [];
    keys.forEach((k) => {
      if (res[k] !== undefined) parts.push(`${k}=${res[k]}`);
    });
    if (res.document_size) parts.push(`size=${res.document_size.width}×${res.document_size.height}`);
    return parts.join(" · ");
  }

  async function sendChat() {
    const input = $("#chatInput");
    const text = input.value.trim();
    if (!text || state.busy) return;
    if (!settings.api_key) {
      toast("请先在「设置」里填写模型 API Key", "bad");
      openSettings();
      return;
    }
    input.value = "";
    chatAdd("user", esc(text));
    state.messages.push({ role: "user", content: text });

    setBusy(true);
    try {
      for (let step = 0; step < 10; step++) {
        const typing = typingBubble();
        let body;
        try {
          body = await llmChat(state.messages, state.tools);
        } catch (err) {
          typing.remove();
          chatAdd("assistant", esc("请求失败：" + err.message));
          break;
        }
        typing.remove();

        const msg = body.message || {};
        state.messages.push(msg);
        if (msg.content) chatAdd("assistant", esc(msg.content));

        const calls = msg.tool_calls || [];
        if (!calls.length) break;

        for (const call of calls) {
          const name = call.function && call.function.name;
          let args = {};
          try {
            args = JSON.parse((call.function && call.function.arguments) || "{}");
          } catch {
            args = {};
          }
          const item = chatAdd("tool", `调用 <b>${esc(name)}</b> …`);
          let content;
          try {
            const res = await api(name, args);
            const summary = summarize(res);
            item.querySelector(".bubble").innerHTML = `<span class="ok">✓</span> ${esc(name)} ${esc(summary)}`;
            content = JSON.stringify(res);
            schedulePreview(160);
          } catch (err) {
            item.querySelector(".bubble").innerHTML = `<span class="bad">✕</span> ${esc(name)} · ${esc(err.message)}`;
            content = JSON.stringify({ ok: false, error: err.message });
          }
          state.messages.push({ role: "tool", tool_call_id: call.id, content });
        }
        await refreshDoc();
      }
      await refreshDoc();
      schedulePreview(160);
    } finally {
      setBusy(false);
      refreshStatusOnly();
    }
  }

  // ---------------------------------------------------------------- 杂项
  let statusTimer = null;
  async function refreshStatusOnly() {
    try {
      const health = await fetch(API + "/health").then((r) => r.json());
      if (health.photoshop && health.photoshop.connected) {
        setStatus("ok", `Photoshop ${health.photoshop.version || ""}`.trim());
      } else {
        setStatus("idle", "Photoshop 未启动");
      }
    } catch {
      setStatus("bad", "API 未连接");
    }
  }

  function selectTab(name) {
    $$(".tab").forEach((t) => t.setAttribute("aria-selected", String(t.dataset.tab === name)));
    ["chat", "tool", "log"].forEach((key) => $("#tab-" + key).classList.toggle("hidden", key !== name));
  }

  // ---------------------------------------------------------------- 启动
  async function loadTools() {
    try {
      const body = await fetch(API + "/tools").then((r) => r.json());
      state.tools = body.tools || [];
      renderTools();
      $("#meta").textContent = state.tools.length + " 工具";
    } catch {
      state.tools = [];
    }
  }

  async function loadConfig() {
    try {
      state.config = await fetch(API + "/ui/config").then((r) => r.json());
    } catch {
      /* 忽略 */
    }
  }

  function bind() {
    $("#btnRefresh").addEventListener("click", () => refresh());
    $("#btnSettings").addEventListener("click", openSettings);
    $("#btnCreate").addEventListener("click", createDoc);
    $("#btnSnap").addEventListener("click", () => takeSnapshot());
    $("#btnExportPng").addEventListener("click", () => openExport("png"));
    $("#btnSavePsd").addEventListener("click", () => openExport("psd"));
    $("#btnRunTool").addEventListener("click", runTool);
    $("#btnChat").addEventListener("click", sendChat);

    $$(".tab").forEach((t) => t.addEventListener("click", () => selectTab(t.dataset.tab)));

    $("#chatInput").addEventListener("keydown", (e) => {
      if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
        e.preventDefault();
        sendChat();
      }
    });
    $("#toolArgs").addEventListener("keydown", (e) => {
      if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
        e.preventDefault();
        runTool();
      }
    });
    $("#docBg").addEventListener("change", (e) => {
      $("#docColor").disabled = e.target.value !== "color";
    });
    $("#docColor").disabled = true;
  }

  // ---------------------------------------------------------------- 设计预览用假数据
  function mockInit() {
    setStatus("ok", "Photoshop 27.0");
    $("#meta").textContent = "3 文档 · 23 工具";
    $("#docName").textContent = "night-mode-poster";
    $("#docMeta").textContent = "1200 × 800 px · 72 dpi · RGB · 6 图层";
    $("#layerCount").textContent = "6";
    const layers = [
      { name: "标题 · 夜间模式", kind: "text", visible: true, opacity: 100 },
      { name: "圆环 · 蓝", kind: "normal", visible: true, opacity: 90 },
      { name: "圆环 · 橙", kind: "normal", visible: true, opacity: 85 },
      { name: "装饰线", kind: "normal", visible: false, opacity: 60 },
      { name: "背景", kind: "normal", visible: true, opacity: 100 },
    ];
    state.layers = layers;
    $("#layerList").innerHTML = layers
      .map(
        (l, i) => `<div class="layer ${l.visible ? "" : "off"} ${i === 0 ? "active" : ""}" data-name="${l.name}">
          <button class="layer-eye"><svg class="ico"><use href="#${l.visible ? "i-eye" : "i-eye-off"}"/></svg></button>
          <span class="layer-name">${l.name}</span>
          <span class="layer-kind">${l.kind}</span>
          <span class="layer-op">${l.opacity}%</span>
        </div>`
      )
      .join("");

    const canvas = document.createElement("canvas");
    canvas.width = 1200;
    canvas.height = 800;
    const c = canvas.getContext("2d");
    const g = c.createLinearGradient(0, 0, 1200, 800);
    g.addColorStop(0, "#0E1726");
    g.addColorStop(1, "#16233A");
    c.fillStyle = g;
    c.fillRect(0, 0, 1200, 800);
    c.strokeStyle = "rgba(59,130,246,.85)";
    c.lineWidth = 26;
    c.beginPath();
    c.arc(880, 420, 220, 0, Math.PI * 2);
    c.stroke();
    c.strokeStyle = "rgba(255,122,24,.8)";
    c.beginPath();
    c.arc(980, 500, 140, 0, Math.PI * 2);
    c.stroke();
    c.fillStyle = "#E9EEF7";
    c.font = "700 92px 'Microsoft YaHei'";
    c.fillText("夜间模式", 92, 330);
    c.fillStyle = "#8FA3BF";
    c.font = "400 30px 'Microsoft YaHei'";
    c.fillText("Photoshop Studio · 本地控制台", 94, 392);
    c.fillStyle = "rgba(255,255,255,.06)";
    c.fillRect(92, 560, 700, 2);
    const img = $("#preview");
    img.src = canvas.toDataURL("image/png");
    img.hidden = false;
    $("#stageEmpty").hidden = true;
    $("#snapTime").textContent = "预览 " + stamp();

    chatAdd("user", "新建 1200×800 深色画布，画一蓝一橙两个圆环，标题写「夜间模式」，导出 PNG");
    chatAdd("assistant", "好的，我先建画布，再依次绘制圆环和标题，最后导出并生成预览图。");
    chatAdd("tool", '<span class="ok">✓</span> ps_doc_create name=night-mode-poster · 1200×800');
    chatAdd("tool", '<span class="ok">✓</span> ps_shape shape=ellipse · 220×220');
    chatAdd("assistant", "画好了，预览已更新。PNG 已导出到图片文件夹\\night-mode.png。");
    $("#chatEmpty")?.remove();

    ["ps_status 132ms", "ps_doc_create 244ms", "ps_shape 486ms", "ps_text 705ms", "ps_snapshot 1483ms"].forEach(
      (line, i) => logLine(line, "ok")
    );
    logLine("ps_layer_set 失败 · 找不到图层: 背景副本", "bad");
    $("#toolResult").textContent = JSON.stringify({ name: "night-mode-poster", width: 1200, height: 800, layer_count: 6 }, null, 2);
  }

  async function init() {
    renderPresets();
    renderQuick();
    bind();
    if (MOCK) {
      mockInit();
      return;
    }
    await loadConfig();
    await loadTools();
    await refresh();
    statusTimer = setInterval(refreshStatusOnly, 10000);
  }

  document.addEventListener("DOMContentLoaded", init);
})();

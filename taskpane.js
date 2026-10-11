/* Prescience by Psomas — the "Send to Prescience" button (task pane), add-in v1.0.

   Works in classic Outlook AND the new Outlook for Windows (Matt's rule of 2026-10-10): the
   pane is plain Office.js. It never touches the PC: with one press it hands the selected
   email (File) or ONE of its attachments (Compare) to the "Prescience Mail" Power Automate
   flow, which writes the file into the shared inbox folder; Prescience on the person's PC
   picks it up within a minute or two and puts it in the Upload or Compare box, exactly as
   if it had been dropped there.

   How the bytes travel: the pane reads them itself when Outlook allows (the attachment's
   content needs Mailbox 1.8; the whole email as EML needs Mailbox 1.14) and sends them
   base64-encoded; when it cannot, it sends the ids and the flow fetches the email or the
   attachment from the mailbox under the person's own connection.

   Nothing is stored on this page: the flow address lives in the user's Outlook roaming
   settings. Requests go as text/plain so the browser makes no CORS preflight; the flow
   answers with Access-Control-Allow-Origin: *.
*/
(function () {
  "use strict";
  const VERSION = "1.0";
  const $ = (id) => document.getElementById(id);
  const DOC_EXTS = [".pdf", ".docx", ".doc", ".rtf"];
  const state = { flowUrl: "", user: { name: "", email: "" }, email: null, busy: false, has18: false, has114: false };

  // ------------------------------------------------------------------ small helpers
  const pad = (n) => String(n).padStart(2, "0");
  const esc = (s) => String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const isoLocal = (d) => { const off = -d.getTimezoneOffset(), s = off >= 0 ? "+" : "-", a = Math.abs(off);
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}${s}${pad(Math.floor(a / 60))}:${pad(a % 60)}`; };
  const stamp = () => { const d = new Date(); return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`; };
  const rand = (n) => Array.from(crypto.getRandomValues(new Uint8Array(n))).map((b) => (b % 16).toString(16)).join("");
  const humanDateTime = (d) => `${pad(d.getMonth() + 1)}/${pad(d.getDate())}/${d.getFullYear()} ${d.getHours() % 12 || 12}:${pad(d.getMinutes())} ${d.getHours() < 12 ? "AM" : "PM"}`;
  const sizeText = (n) => !n ? "" : n < 1 << 20 ? `${Math.round(n / 1024).toLocaleString()} KB` : `${(n / (1 << 20)).toFixed(1)} MB`;
  const extOf = (name) => { const i = (name || "").lastIndexOf("."); return i >= 0 ? name.slice(i).toLowerCase() : ""; };
  const safeName = (name, dflt) => { const n = String(name || "").replace(/[<>:"/\\|?*\x00-\x1f]+/g, " ").replace(/\s+/g, " ").trim().replace(/^\.+|\.+$/g, ""); return (n || dflt).slice(0, 150); };
  function roam(k, v) { try { const rs = Office.context.roamingSettings; if (v === undefined) return rs.get(k); rs.set(k, v); rs.saveAsync(() => {}); } catch (e) { return undefined; } }
  function show(id, on) { $(id).classList.toggle("hidden", !on); }
  function setStatus(text, kind) { const el = $("status"); if (!text) { el.classList.add("hidden"); return; } el.className = "status " + (kind || ""); el.textContent = text; el.classList.remove("hidden"); }
  function page(name) {
    ["sec-setup", "sec-email", "sec-choose", "sec-pick", "sec-result", "sec-settings"].forEach((id) => show(id, false));
    if (name === "setup") show("sec-setup", true);
    else if (name === "settings") { show("sec-settings", true); renderSettings(); }
    else if (name === "pick") { show("sec-email", true); show("sec-pick", true); }
    else if (name === "result") show("sec-result", true);
    else { show("sec-email", true); show("sec-choose", true); }
  }
  function lock(on) { state.busy = on; ["btn-file", "btn-compare", "btn-pick-back", "btn-again"].forEach((id) => { const b = $(id); if (b && !(id === "btn-compare" && b.dataset.none === "1")) b.disabled = on; });
    document.querySelectorAll("#pick-list button").forEach((b) => { b.disabled = on; }); }

  // ------------------------------------------------------------------ the email
  function readItem() {
    const item = Office.context.mailbox && Office.context.mailbox.item; if (!item) return null;
    let restId = ""; try { restId = Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0); } catch (e) { restId = item.itemId || ""; }
    const atts = (item.attachments || []).map((a) => ({
      id: a.id || "", name: a.name || "attachment", size: a.size || 0, type: String(a.attachmentType || "file").toLowerCase(),
      inline: !!a.isInline, contentType: a.contentType || "" }));
    return { subject: item.subject || "", senderName: item.from ? item.from.displayName || "" : "", senderEmail: item.from ? (item.from.emailAddress || "").toLowerCase() : "",
      received: item.dateTimeCreated ? new Date(item.dateTimeCreated) : null, itemId: item.itemId || "", restId, attachments: atts };
  }
  const eligible = (a) => a.type === "file" && !a.inline && DOC_EXTS.includes(extOf(a.name));
  const compareChoice = (e) => { const good = (e ? e.attachments : []).filter(eligible); return { mode: !good.length ? "none" : good.length === 1 ? "one" : "many", good }; };

  function renderEmail() {
    const e = state.email;
    if (!e) { $("email-subject").textContent = "Open or select an email, then press the button again."; $("email-detail").textContent = ""; show("email-pill", false); $("btn-file").disabled = true; $("btn-compare").disabled = true; return; }
    $("email-subject").textContent = e.subject || "(no subject)";
    const n = e.attachments.filter((a) => !a.inline).length;
    const who = e.senderName && e.senderEmail && e.senderName.toLowerCase() !== e.senderEmail ? `${e.senderName} <${e.senderEmail}>` : (e.senderName || e.senderEmail || "(unknown sender)");
    $("email-detail").textContent = `From ${who}${e.received ? " · " + humanDateTime(e.received) : ""} · ${n ? n + " attachment" + (n !== 1 ? "s" : "") : "no attachments"}`;
    $("email-pill").textContent = "selected in Outlook"; show("email-pill", true);
    $("btn-file").disabled = state.busy;
    const c = compareChoice(e); const b = $("btn-compare"); b.dataset.none = c.mode === "none" ? "1" : "";
    if (c.mode === "none") { $("compare-hint").textContent = "This email has no PDF or Word attachment to compare."; b.textContent = "Compare an attachment"; b.disabled = true; }
    else if (c.mode === "one") { $("compare-hint").textContent = `${c.good[0].name} goes to Prescience's Compare box as the NEW contract. A comparison files nothing.`; b.textContent = `Compare ${short(c.good[0].name)}`; b.disabled = state.busy; }
    else { $("compare-hint").textContent = `This email carries ${c.good.length} documents — you'll pick the one to compare.`; b.textContent = `Pick one of ${c.good.length} attachments…`; b.disabled = state.busy; }
  }
  const short = (s, n = 30) => s.length <= n ? s : s.slice(0, n - 1) + "…";

  function renderPick() {
    const host = $("pick-list"); host.innerHTML = "";
    for (const a of compareChoice(state.email).good) {
      const row = document.createElement("div"); row.className = "att";
      const kind = document.createElement("span"); kind.className = "pill info"; kind.textContent = extOf(a.name) === ".pdf" ? "PDF" : "Word";
      const name = document.createElement("span"); name.className = "name"; name.textContent = a.name; name.title = a.name;
      const size = document.createElement("span"); size.className = "size"; size.textContent = sizeText(a.size);
      const btn = document.createElement("button"); btn.className = "primary mini"; btn.textContent = "Compare this one"; btn.addEventListener("click", () => send("compare", a));
      row.append(kind, name, size, btn); host.appendChild(row);
    }
  }

  // ------------------------------------------------------------------ reading the bytes (when Outlook lets us)
  function emailBase64() {
    return new Promise((resolve) => {
      const item = Office.context.mailbox.item;
      if (!state.has114 || !item || typeof item.getAsFileAsync !== "function") return resolve("");
      try { item.getAsFileAsync((r) => resolve(r.status === Office.AsyncResultStatus.Succeeded && typeof r.value === "string" ? r.value : "")); }
      catch (e) { resolve(""); }
    });
  }
  function attachmentBase64(att) {
    return new Promise((resolve) => {
      const item = Office.context.mailbox.item;
      if (!state.has18 || !item || typeof item.getAttachmentContentAsync !== "function") return resolve("");
      try {
        item.getAttachmentContentAsync(att.id, (r) => {
          const v = r.status === Office.AsyncResultStatus.Succeeded ? r.value : null;
          resolve(v && v.format === Office.MailboxEnums.AttachmentContentFormat.Base64 && typeof v.content === "string" ? v.content : "");
        });
      } catch (e) { resolve(""); }
    });
  }
  function restAttachmentId(id) { try { return Office.context.mailbox.convertToRestId(id, Office.MailboxEnums.RestVersion.v2_0); } catch (e) { return id; } }

  // ------------------------------------------------------------------ the flow
  const GENERIC = ["An action failed", "ActionFailed", "The execution of template action"];
  function shortError(raw) {
    if (raw === null || raw === undefined) return "The flow reported a failure.";
    const s = typeof raw === "string" ? raw.trim() : JSON.stringify(raw); let data = null;
    try { data = /^[\[{]/.test(s) ? JSON.parse(s) : null; } catch (e) { data = null; }
    const pick = (o) => { if (!o || typeof o !== "object") return ""; for (const k of ["message", "error"]) { let v = o[k]; if (v && typeof v === "object") v = v.message; if (typeof v === "string" && v.trim() && !GENERIC.some((g) => v.includes(g))) return v.trim(); } return ""; };
    if (Array.isArray(data)) { for (const it of data) { const m = it && it.error && it.error.message; if (m) return `${it.name || "a step"}: ${m}`.slice(0, 400); } }
    else if (data) { const m = pick(data); if (m) return m.slice(0, 400); }
    return s.slice(0, 400) || "The flow reported a failure.";
  }
  async function callFlow(payload, timeoutMs) {
    if (!state.flowUrl) throw new Error("The flow address isn't set — open Settings.");
    const ctrl = new AbortController(); const t = setTimeout(() => ctrl.abort(), timeoutMs || 115000);
    let res;
    try { res = await fetch(state.flowUrl, { method: "POST", headers: { "Content-Type": "text/plain;charset=UTF-8" }, body: JSON.stringify(payload), signal: ctrl.signal }); }
    catch (e) { throw new Error(e.name === "AbortError" ? "The flow didn't answer in time — try again in a moment." : "Couldn't reach the flow (" + e.message + "). Is the address right and the flow turned on?"); }
    finally { clearTimeout(t); }
    const text = await res.text(); let data;
    try { data = JSON.parse(text); if (typeof data === "string") data = JSON.parse(data); } catch (e) { data = { ok: res.ok, error: text.slice(0, 300) || ("HTTP " + res.status) }; }
    if (!res.ok && data.ok !== true) throw new Error(shortError(data.error || ("HTTP " + res.status)));
    if (data.ok === false) throw new Error(shortError(data.error));
    return data;
  }
  const basePayload = (kind) => ({ schema: "prescience-mail/1", kind, id: `pm-${kind}-${stamp()}-${rand(4)}`, submittedBy: state.user.email, submittedAt: isoLocal(new Date()) });

  // ------------------------------------------------------------------ the one action
  async function send(action, att) {
    if (state.busy || !state.email) return;
    lock(true);
    const e = state.email;
    try {
      setStatus(action === "upload" ? "Reading the email…" : `Reading ${att.name}…`, "busy");
      const content = action === "upload" ? await emailBase64() : await attachmentBase64(att);
      const name = action === "upload" ? safeName(e.subject, "email") + ".eml" : safeName(att.name, "attachment");
      setStatus("Sending to Prescience…", "busy");
      const res = await callFlow(Object.assign(basePayload("deliver"), {
        action, name, messageId: e.restId, attachmentId: att ? restAttachmentId(att.id) : "", contentBase64: content }), 115000);
      setStatus("", "");
      $("res-text").textContent = `${name} is on its way to Prescience's ${action === "upload" ? "Upload" : "Compare"} box.`;
      $("res-note").textContent = `It appears there within a minute or two — Prescience must be open on your PC, signed in as ${state.user.email}.`
        + (res.via === "mailbox" ? " (The flow fetched it from your mailbox.)" : "");
      page("result");
    } catch (err) {
      setStatus(err.message, "err");
    } finally { lock(false); renderEmail(); }
  }

  // ------------------------------------------------------------------ settings
  function renderSettings() { $("set-url").value = state.flowUrl || ""; $("about").textContent = `Prescience add-in ${VERSION} · signed in as ${state.user.email} · Outlook API: attachments ${state.has18 ? "yes" : "no"}, whole email ${state.has114 ? "yes" : "no"}`; }
  async function connect(url, fromSetup) {
    const msg = fromSetup ? $("setup-msg") : $("set-msg");
    url = (url || "").trim(); if (!/^https:\/\//i.test(url)) { msg.textContent = "That doesn't look like a flow address (it starts with https://)."; return false; }
    state.flowUrl = url; msg.textContent = "Connecting…";
    try { const r = await callFlow(basePayload("ping"), 40000); if (!r.mailbox || r.mailbox === "none") { msg.textContent = `The flow answered, but it has no mailbox connection for ${state.user.email} — Matt adds one in the flow. Files you fetch here still go through.`; }
      else msg.textContent = `Connected — the flow knows ${state.user.email} (mailbox “${r.mailbox}”).`; }
    catch (e) { msg.textContent = e.message; return false; }
    roam("prescienceFlowUrl", url); return true;
  }

  // ------------------------------------------------------------------ init
  function wire() {
    $("btn-file").addEventListener("click", () => send("upload", null));
    $("btn-compare").addEventListener("click", () => { const c = compareChoice(state.email); if (c.mode === "one") send("compare", c.good[0]); else if (c.mode === "many") { renderPick(); page("pick"); } });
    $("btn-pick-back").addEventListener("click", () => page("main"));
    $("btn-again").addEventListener("click", () => { state.email = readItem(); renderEmail(); page("main"); });
    $("setup-save").addEventListener("click", async () => { if (await connect($("setup-url").value, true)) page("main"); });
    $("set-save").addEventListener("click", async () => { if (await connect($("set-url").value, false)) setStatus("Flow address saved.", "ok"); });
    $("set-test").addEventListener("click", async () => { await connect($("set-url").value || state.flowUrl, false); });
    $("nav-main").addEventListener("click", (e) => { e.preventDefault(); page(state.flowUrl ? "main" : "setup"); });
    $("nav-settings").addEventListener("click", (e) => { e.preventDefault(); page("settings"); });
  }

  // a few pure helpers exposed for checking from a browser console (no data, no secrets)
  window.__pr = { eligible, compareChoice, safeName, shortError, extOf, version: VERSION };

  Office.onReady(() => {
    wire();
    try { state.has18 = Office.context.requirements.isSetSupported("Mailbox", "1.8"); } catch (e) { state.has18 = false; }
    try { state.has114 = Office.context.requirements.isSetSupported("Mailbox", "1.14"); } catch (e) { state.has114 = false; }
    try { const p = Office.context.mailbox.userProfile; state.user = { name: p.displayName || "", email: (p.emailAddress || "").toLowerCase() }; } catch (e) { state.user = { name: "", email: "" }; }
    state.flowUrl = roam("prescienceFlowUrl") || "";
    state.email = readItem(); renderEmail();
    try { Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => { state.email = readItem(); renderEmail(); if (!$("sec-settings").classList.contains("hidden") || !$("sec-setup").classList.contains("hidden")) return; page("main"); }); } catch (e) { /* pinning not supported here */ }
    page(state.flowUrl ? "main" : "setup");
  });
})();

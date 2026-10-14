/* Prescience by Psomas — the "Send to Prescience" button (task pane), add-in v2.0.

   v2.0 (Matt, 2026-10-14): File this email files the contract FROM INSIDE OUTLOOK. The pane
   asks Prescience's own questions (published by Prescience through the flow; the built-in
   five until it has), and Prescience on the person's PC does the filing with the very code its
   filing dialog uses, reporting back through a receipt the pane polls:

     File this email  ->  the email goes to the flow's queue (kind deliver, action prepare)
                          + the prescience: link wakes / starts Prescience (inside the click)
                      ->  receipts: received, prepared (documents, duplicates, the email's date),
                          suggestions (Egnyte AI, the employee directory) — filled in only where
                          the person has not answered
     File             ->  the answers go to the queue (action commit) + the link again
                      ->  receipt: filed (or duplicate -> File anyway, or error -> fix and File again)

   Works in classic Outlook AND the new Outlook (Matt's rule of 2026-10-10): plain Office.js.
   Nothing is stored on this page: the flow address lives in the user's Outlook roaming settings
   (or comes with the IT manifest as ?flow=), the form is cached in this pane's localStorage.
   ?dev=1 runs without Outlook against a fake flow (tools/fake_mail_flow.py).
*/
(function () {
  "use strict";
  const D = window.PR_DEFAULTS || { version: "2.0", link: { fileUrl: "prescience://file?job={jobId}", openUrl: "prescience://open" },
    form: { questions: [], entities: [] }, noise: [], docExts: [".pdf", ".docx", ".doc", ".rtf"], jobSchema: "prescience-file/1" };
  const VERSION = D.version || "2.0";
  const $ = (id) => document.getElementById(id);
  const DOC_EXTS = D.docExts || [".pdf", ".docx", ".doc", ".rtf"];
  const NOISE = new Set(D.noise || []);
  const state = { flowUrl: "", user: { name: "", email: "" }, email: null, busy: false, has18: false, has114: false,
    dev: false, form: null, job: null, controls: {} };

  // ------------------------------------------------------------------ small helpers
  const pad = (n) => String(n).padStart(2, "0");
  const isoLocal = (d) => { const off = -d.getTimezoneOffset(), s = off >= 0 ? "+" : "-", a = Math.abs(off);
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}${s}${pad(Math.floor(a / 60))}:${pad(a % 60)}`; };
  const ymd = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const stamp = () => { const d = new Date(); return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`; };
  const rand = (n) => Array.from(crypto.getRandomValues(new Uint8Array(n))).map((b) => (b % 16).toString(16)).join("");
  const humanDateTime = (d) => `${pad(d.getMonth() + 1)}/${pad(d.getDate())}/${d.getFullYear()} ${d.getHours() % 12 || 12}:${pad(d.getMinutes())} ${d.getHours() < 12 ? "AM" : "PM"}`;
  const mdy = (iso) => { const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso || ""); return m ? `${m[2]}/${m[3]}/${m[1]}` : (iso || ""); };
  const sizeText = (n) => !n ? "" : n < 1 << 20 ? `${Math.round(n / 1024).toLocaleString()} KB` : `${(n / (1 << 20)).toFixed(1)} MB`;
  const extOf = (name) => { const i = (name || "").lastIndexOf("."); return i >= 0 ? name.slice(i).toLowerCase() : ""; };
  const safeName = (name, dflt) => { const n = String(name || "").replace(/[<>:"/\\|?*\x00-\x1f]+/g, " ").replace(/_{2,}/g, "_").replace(/\s+/g, " ").trim().replace(/^\.+|\.+$/g, ""); return (n || dflt).slice(0, 140); };
  const textToB64 = (t) => { const bytes = new TextEncoder().encode(t); let s = ""; for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000)); return btoa(s); };
  const b64ToText = (b) => { const bin = atob(b); const bytes = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i); return new TextDecoder("utf-8").decode(bytes).replace(/^﻿/, ""); };
  const splitValues = (text) => { const out = [], seen = new Set(); for (const part of String(text || "").split(";")) { const v = part.trim(); if (v && !seen.has(v.toLowerCase())) { seen.add(v.toLowerCase()); out.push(v); } } return out; };
  function roam(k, v) { try { const rs = Office.context.roamingSettings; if (v === undefined) return rs.get(k); rs.set(k, v); rs.saveAsync(() => {}); } catch (e) { try { if (v === undefined) return localStorage.getItem("roam-" + k); localStorage.setItem("roam-" + k, v); } catch (e2) { /* nowhere to keep it */ } return undefined; } }
  function show(id, on) { const el = $(id); if (el) el.classList.toggle("hidden", !on); }
  function setStatus(text, kind) { const el = $("status"); if (!text) { el.classList.add("hidden"); return; } el.className = "status " + (kind || ""); el.textContent = text; el.classList.remove("hidden"); }
  function setProgress(text, kind) { const el = $("file-progress"); if (!text) { el.classList.add("hidden"); return; } el.className = "progress status " + (kind || ""); el.textContent = text; el.classList.remove("hidden"); }
  function page(name) {
    ["sec-setup", "sec-email", "sec-choose", "sec-pick", "sec-file", "sec-result", "sec-settings"].forEach((id) => show(id, false));
    if (name === "setup") show("sec-setup", true);
    else if (name === "settings") { show("sec-settings", true); renderSettings(); }
    else if (name === "pick") { show("sec-email", true); show("sec-pick", true); }
    else if (name === "file") { show("sec-email", true); show("sec-file", true); }
    else if (name === "result") show("sec-result", true);
    else { show("sec-email", true); show("sec-choose", true); }
  }

  // ------------------------------------------------------------------ entity names (mirror of entities.normalize_name)
  function normalizeName(display) {
    let t = String(display || "").normalize("NFKD").replace(/[̀-ͯ]/g, "");
    t = t.replace(/&/g, " and ");
    t = t.replace(/\b(?:[A-Za-z]\.){2,}/g, (m) => m.replace(/\./g, ""));
    t = t.toLowerCase().replace(/[^\p{L}\p{N}_\s]/gu, " ");
    const tokens = t.split(/\s+/).filter(Boolean);
    while (tokens.length && tokens[0] === "the") tokens.shift();
    while (tokens.length > 1 && NOISE.has(tokens[tokens.length - 1])) tokens.pop();
    return tokens.join(" ");
  }
  function exactMatch(query, names) { const q = normalizeName(query); if (!q) return null; for (const n of names || []) if (normalizeName(n) === q) return n; return null; }

  // ------------------------------------------------------------------ the email
  function readItem() {
    if (state.dev) return devItem();
    const item = Office.context.mailbox && Office.context.mailbox.item; if (!item) return null;
    let restId = ""; try { restId = Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0); } catch (e) { restId = item.itemId || ""; }
    const atts = (item.attachments || []).map((a) => ({ id: a.id || "", name: a.name || "attachment", size: a.size || 0,
      type: String(a.attachmentType || "file").toLowerCase(), inline: !!a.isInline, contentType: a.contentType || "" }));
    return { subject: item.subject || "", senderName: item.from ? item.from.displayName || "" : "", senderEmail: item.from ? (item.from.emailAddress || "").toLowerCase() : "",
      received: item.dateTimeCreated ? new Date(item.dateTimeCreated) : null, itemId: item.itemId || "", restId, attachments: atts };
  }
  function devItem() {
    return { subject: "RE: Lease for review", senderName: "Chad Wilson", senderEmail: "chad.wilson@psomas.com", received: new Date(2026, 9, 9, 9, 30),
      itemId: "dev-message", restId: "dev-message", attachments: [
        { id: "a1", name: "Master Agreement.pdf", size: 52000, type: "file", inline: false, contentType: "application/pdf" },
        { id: "a2", name: "Exhibit B.pdf", size: 21000, type: "file", inline: false, contentType: "application/pdf" }] };
  }
  const eligible = (a) => a.type === "file" && !a.inline && DOC_EXTS.includes(extOf(a.name));
  const compareChoice = (e) => { const good = (e ? e.attachments : []).filter(eligible); return { mode: !good.length ? "none" : good.length === 1 ? "one" : "many", good }; };
  const short = (s, n = 30) => s.length <= n ? s : s.slice(0, n - 1) + "…";

  function renderEmail() {
    const e = state.email;
    if (!e) { $("email-subject").textContent = "Open or select an email, then press the button again."; $("email-detail").textContent = ""; show("email-pill", false); $("btn-file").disabled = true; $("btn-compare").disabled = true; return; }
    $("email-subject").textContent = e.subject || "(no subject)";
    const n = e.attachments.filter((a) => !a.inline).length;
    const who = e.senderName && e.senderEmail && e.senderName.toLowerCase() !== e.senderEmail ? `${e.senderName} <${e.senderEmail}>` : (e.senderName || e.senderEmail || "(unknown sender)");
    $("email-detail").textContent = `From ${who}${e.received ? " · " + humanDateTime(e.received) : ""} · ${n ? n + " attachment" + (n !== 1 ? "s" : "") : "no attachments"}`;
    $("email-pill").textContent = state.dev ? "dev sample" : "selected in Outlook"; show("email-pill", true);
    $("btn-file").disabled = state.busy;
    const c = compareChoice(e); const b = $("btn-compare"); b.dataset.none = c.mode === "none" ? "1" : "";
    if (c.mode === "none") { $("compare-hint").textContent = "This email has no PDF or Word attachment to compare."; b.textContent = "Compare an attachment"; b.disabled = true; }
    else if (c.mode === "one") { $("compare-hint").textContent = `${c.good[0].name} goes to Prescience's Compare box as the NEW contract. A comparison files nothing.`; b.textContent = `Compare ${short(c.good[0].name)}`; b.disabled = state.busy; }
    else { $("compare-hint").textContent = `This email carries ${c.good.length} documents — you'll pick the one to compare.`; b.textContent = `Pick one of ${c.good.length} attachments…`; b.disabled = state.busy; }
  }

  function renderPick() {
    const host = $("pick-list"); host.innerHTML = "";
    for (const a of compareChoice(state.email).good) {
      const row = document.createElement("div"); row.className = "att";
      const kind = document.createElement("span"); kind.className = "pill info"; kind.textContent = extOf(a.name) === ".pdf" ? "PDF" : "Word";
      const name = document.createElement("span"); name.className = "name"; name.textContent = a.name; name.title = a.name;
      const size = document.createElement("span"); size.className = "size"; size.textContent = sizeText(a.size);
      const btn = document.createElement("button"); btn.className = "primary mini"; btn.textContent = "Compare this one"; btn.addEventListener("click", () => sendCompare(a));
      row.append(kind, name, size, btn); host.appendChild(row);
    }
  }

  // ------------------------------------------------------------------ reading the bytes (when Outlook lets us)
  function emailBase64() {
    return new Promise((resolve) => {
      if (state.dev) return resolve("");
      const item = Office.context.mailbox.item;
      if (!state.has114 || !item || typeof item.getAsFileAsync !== "function") return resolve("");
      try { item.getAsFileAsync((r) => resolve(r.status === Office.AsyncResultStatus.Succeeded && typeof r.value === "string" ? r.value : "")); }
      catch (e) { resolve(""); }
    });
  }
  function attachmentBase64(att) {
    return new Promise((resolve) => {
      if (state.dev) return resolve("");
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
  function restAttachmentId(id) { if (state.dev) return id; try { return Office.context.mailbox.convertToRestId(id, Office.MailboxEnums.RestVersion.v2_0); } catch (e) { return id; } }

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
  const basePayload = (kind) => ({ schema: "prescience-mail/1", kind, id: `pm-${kind}-${stamp()}-${rand(4)}`, submittedBy: state.user.email, submittedByName: state.user.name, submittedAt: isoLocal(new Date()) });

  // ------------------------------------------------------------------ the prescience: link (wakes / starts Prescience)
  function openLink(url) {
    if (state.dev) { window.__lastLink = url; return; }
    try { window.location.assign(url); } catch (e) { /* the result page offers the link too */ }
  }
  const fileLink = (jobId) => (D.link.fileUrl || "prescience://file?job={jobId}").replace("{jobId}", encodeURIComponent(jobId));

  // ------------------------------------------------------------------ the form Prescience published
  function cachedForm() { try { const t = localStorage.getItem("prescienceForm"); const f = t ? JSON.parse(t) : null; return f && Array.isArray(f.questions) ? f : null; } catch (e) { return null; } }
  async function refreshForm() {
    try {
      const r = await callFlow(basePayload("form"), 40000);
      if (r.ready && r.contentBase64) {
        const f = JSON.parse(b64ToText(r.contentBase64));
        if (f && Array.isArray(f.questions) && f.questions.length) { state.form = f; try { localStorage.setItem("prescienceForm", JSON.stringify(f)); } catch (e) { /* too big — fine */ } }
      }
    } catch (e) { /* keep what we have */ }
  }
  const currentForm = () => state.form || D.form;

  // ------------------------------------------------------------------ the questions
  function el(tag, cls, text) { const e = document.createElement(tag); if (cls) e.className = cls; if (text !== undefined) e.textContent = text; return e; }
  function markTouched(key) { if (state.job) state.job.touched[key] = true; const c = state.controls[key]; if (c && c.wrap) c.wrap.classList.remove("suggested"); validate(); }

  function buildEntity(q) {
    const wrap = el("div", "q");
    const lab = el("label", "qlabel"); lab.textContent = q.label; if (q.required) lab.appendChild(el("span", "req", " *"));
    const input = el("input"); input.type = "text"; input.id = "q-entity"; input.setAttribute("list", "entity-names"); input.autocomplete = "off";
    const list = el("datalist"); list.id = "entity-names";
    for (const n of currentForm().entities || []) { const o = document.createElement("option"); o.value = n; list.appendChild(o); }
    const status = el("div", "help");
    const upd = () => { const t = input.value.trim(); if (!t) { status.textContent = ""; status.className = "help"; return; }
      const hit = exactMatch(t, currentForm().entities || []);
      status.textContent = hit ? `✓ existing entity: ${hit}` : `＋ new entity will be created: ${t}`; status.className = hit ? "help ok" : "help warn"; };
    input.addEventListener("input", () => { upd(); markTouched(q.key); });
    wrap.append(lab, input, list, status);
    return { q, wrap, get: () => { const t = input.value.trim(); return exactMatch(t, currentForm().entities || []) || t; }, set: (v) => { input.value = v || ""; upd(); }, isEmpty: () => !input.value.trim() };
  }
  function buildChoice(q) {
    const wrap = el("div", "q");
    const lab = el("label", "qlabel"); lab.textContent = q.label; if (q.required) lab.appendChild(el("span", "req", " *"));
    wrap.appendChild(lab);
    const options = (q.options || []).slice();
    if (!q.multi) {
      const sel = el("select"); sel.id = "q-" + q.key;
      const add = (v, t) => { const o = document.createElement("option"); o.value = v; o.textContent = t; sel.appendChild(o); };
      add("", "— choose —"); options.forEach((o) => add(o, o)); add("__other__", "＋ Something else…");
      const other = el("input"); other.type = "text"; other.placeholder = `A new ${q.label.toLowerCase()} (added to the list for everyone)`; other.className = "hidden";
      sel.addEventListener("change", () => { other.classList.toggle("hidden", sel.value !== "__other__"); if (sel.value === "__other__") other.focus(); markTouched(q.key); });
      other.addEventListener("input", () => markTouched(q.key));
      wrap.append(sel, other);
      const setVal = (v) => { const want = (v || "").trim(); if (!want) { sel.value = ""; other.classList.add("hidden"); return; }
        const hit = options.find((o) => o.toLowerCase() === want.toLowerCase());
        if (hit) { sel.value = hit; other.classList.add("hidden"); } else { sel.value = "__other__"; other.value = want; other.classList.remove("hidden"); } };
      return { q, wrap, get: () => sel.value === "__other__" ? other.value.trim() : sel.value, set: setVal, isEmpty: () => !(sel.value === "__other__" ? other.value.trim() : sel.value) };
    }
    const box = el("div", "multibox"); box.id = "q-" + q.key;
    const addBox = (v, checked) => { const l = el("label", "check"); const c = el("input"); c.type = "checkbox"; c.value = v; c.checked = !!checked; c.addEventListener("change", () => markTouched(q.key)); l.append(c, document.createTextNode(" " + v)); box.appendChild(l); return c; };
    options.forEach((o) => addBox(o, false));
    const addRow = el("div", "addrow"); const addIn = el("input"); addIn.type = "text"; addIn.placeholder = `Another ${q.label.toLowerCase()}… (Enter adds it)`;
    addIn.addEventListener("keydown", (ev) => { if (ev.key === "Enter") { ev.preventDefault(); const v = addIn.value.trim(); if (!v) return; const hit = [...box.querySelectorAll("input")].find((c) => c.value.toLowerCase() === v.toLowerCase()); if (hit) hit.checked = true; else addBox(v, true); addIn.value = ""; markTouched(q.key); } });
    addRow.appendChild(addIn); wrap.append(box, addRow);
    const checked = () => [...box.querySelectorAll("input:checked")].map((c) => c.value);
    const setVal = (v) => { const want = splitValues(v); for (const c of box.querySelectorAll("input")) c.checked = false;
      for (const w of want) { const hit = [...box.querySelectorAll("input")].find((c) => c.value.toLowerCase() === w.toLowerCase()); if (hit) hit.checked = true; else addBox(w, true); } };
    return { q, wrap, get: () => checked(), set: setVal, isEmpty: () => !checked().length };
  }
  function buildSimple(q, type) {
    const wrap = el("div", "q");
    const lab = el("label", "qlabel"); lab.textContent = q.label; if (q.required) lab.appendChild(el("span", "req", " *"));
    const input = el("input"); input.type = type; input.id = "q-" + q.key;
    input.addEventListener("input", () => markTouched(q.key)); input.addEventListener("change", () => markTouched(q.key));
    wrap.append(lab, input);
    return { q, wrap, get: () => input.value.trim(), set: (v) => { input.value = v || ""; }, isEmpty: () => !input.value.trim() };
  }
  function renderQuestions() {
    const host = $("qform"); host.innerHTML = ""; state.controls = {};
    for (const q of currentForm().questions || []) {
      let c;
      if (q.kind === "entity") c = buildEntity(q);
      else if (q.kind === "choice") c = buildChoice(q);
      else if (q.kind === "date") c = buildSimple(q, "date");
      else c = buildSimple(q, "text");
      state.controls[q.key] = c; host.appendChild(c.wrap);
    }
    $("f-notes").value = "";
  }
  function missingAnswers() {
    const out = [];
    for (const q of currentForm().questions || []) { const c = state.controls[q.key]; if (q.required && (!c || c.isEmpty())) out.push(q.label); }
    return out;
  }
  function validate() {
    const job = state.job; if (!job) return;
    const missing = missingAnswers();
    $("file-problem").textContent = missing.length ? "Still needed: " + missing.join(", ") : "";
    $("btn-do-file").disabled = !!missing.length || job.committing || job.final;
  }
  function collectAnswers() {
    const out = { extra: {} }; const builtin = ["type", "group", "realm", "date", "entity"];
    for (const [key, c] of Object.entries(state.controls)) { const v = c.get(); if (builtin.includes(key)) out[key] = v; else out.extra[key] = v; }
    out.notes = $("f-notes").value.trim();
    return out;
  }
  function setIfUntouched(key, value) {
    const c = state.controls[key]; const job = state.job;
    if (!c || !value || (job && job.touched[key])) return false;
    c.set(value); return true;
  }
  function applySuggestions(sugg, by, note) {
    const applied = [];
    for (const [key, value] of Object.entries(sugg || {})) {
      const c = state.controls[key]; if (!c || !value) continue;
      if ((state.job && state.job.touched[key]) || !c.isEmpty()) continue;
      c.set(value); c.wrap.classList.add("suggested"); applied.push(c.q.label);
    }
    if (applied.length) {
      const who = [...new Set(Object.values(by || {}))].map((s) => s === "directory" ? "the employee directory" : "Egnyte AI").join(" and ");
      $("suggest-pill").textContent = `Suggested: ${applied.join(", ")}`; $("suggest-pill").title = `By ${who || "Prescience"}. ${note || ""}`.trim(); show("suggest-pill", true);
    }
    validate();
  }

  // ------------------------------------------------------------------ documents and duplicates (from Prescience's receipt)
  function renderDocs(docs) {
    const host = $("docs-list"); host.innerHTML = "";
    const atts = (docs || []).filter((d) => d.kind === "attachment");
    if (!atts.length) { show("docs-box", false); return; }
    const keep = state.job.docChoice || { main: null, leaveOut: [] };
    for (const d of docs) {
      const row = el("div", "att");
      if (d.kind === "email") { row.append(el("span", "pill neutral", "Email"), el("span", "name", "The email itself (always filed)")); host.appendChild(row); continue; }
      const kind = el("span", "pill info", extOf(d.name) === ".pdf" ? "PDF" : DOC_EXTS.includes(extOf(d.name)) ? "Word" : extOf(d.name).replace(".", "").toUpperCase() || "File");
      const name = el("span", "name", d.name); name.title = d.name;
      row.append(kind, name);
      if (d.isDocument) {
        const ml = el("label", "check mini"); const m = el("input"); m.type = "radio"; m.name = "doc-main"; m.value = String(d.index);
        m.checked = keep.main !== null ? keep.main === d.index : !!d.main;
        m.addEventListener("change", () => { state.job.docChoice = Object.assign(state.job.docChoice || { leaveOut: [] }, { main: d.index }); });
        ml.append(m, document.createTextNode(" main")); row.appendChild(ml);
      }
      const ll = el("label", "check mini"); const l = el("input"); l.type = "checkbox"; l.value = String(d.index); l.className = "doc-leave";
      l.checked = (keep.leaveOut || []).includes(d.index) || !!d.leftOut;
      l.addEventListener("change", () => { const set = new Set((state.job.docChoice && state.job.docChoice.leaveOut) || []); if (l.checked) set.add(d.index); else set.delete(d.index); state.job.docChoice = Object.assign(state.job.docChoice || { main: null }, { leaveOut: [...set] }); });
      ll.append(l, document.createTextNode(" leave out")); row.appendChild(ll);
      host.appendChild(row);
    }
    show("docs-box", true);
  }
  function chosenMain() { const r = document.querySelector('input[name="doc-main"]:checked'); return r ? parseInt(r.value, 10) : null; }
  function chosenLeaveOut() { return [...document.querySelectorAll("input.doc-leave:checked")].map((c) => parseInt(c.value, 10)); }
  function renderDuplicates(dups, held) {
    const host = $("dup-list"); host.innerHTML = "";
    for (const d of dups || []) {
      const line = el("div", "dupline");
      line.textContent = `${d.entity || "(entity?)"} · ${d.type || "?"}${d.date ? " · " + mdy(d.date) : ""} — ${d.label || d.match}${d.name ? " (" + d.name + ")" : ""}`;
      host.appendChild(line);
    }
    show("dup-box", !!(dups && dups.length));
    if (held) $("dup-box").classList.add("attention");
  }

  // ------------------------------------------------------------------ filing
  function startFiling() {
    if (state.busy || !state.email) return;
    const jobId = `pf-${stamp()}-${rand(8)}`;
    state.job = { id: jobId, seq: 0, receipt: {}, touched: {}, final: false, committing: false, startedAt: Date.now(), lastProgress: Date.now(), pollTimer: null, sentPrepare: false, suggestApplied: false, docChoice: null };
    openLink(fileLink(jobId));                   // inside the click: an app link needs the user's gesture
    renderQuestions();
    const e = state.email;
    if (e && e.received) setIfUntouched("date", ymd(e.received));   // the dialog's rule: the email's date (Prescience confirms it)
    $("dup-ok").checked = false; show("dup-box", false); $("dup-box").classList.remove("attention");
    show("docs-box", false); show("suggest-pill", false); show("lnk-open", false);
    $("btn-do-file").textContent = "File in Prescience";
    page("file"); validate();
    sendPrepare();
  }
  async function sendPrepare() {
    const job = state.job; if (!job) return;
    setProgress("Sending the email to Prescience…", "busy");
    try {
      const content = await emailBase64();
      const name = `${job.id}__${safeName(state.email.subject, "email")}.eml`;
      await callFlow(Object.assign(basePayload("deliver"), { action: "prepare", name, jobId: job.id, messageId: state.email.restId, attachmentId: "", contentBase64: content }), 115000);
      if (state.job !== job) return;
      job.sentPrepare = true; job.lastProgress = Date.now();
      setProgress("Sent — waiting for Prescience on your PC to read it… (you can answer the questions meanwhile)", "busy");
      schedulePoll(3000);
    } catch (err) {
      if (state.job !== job) return;
      setProgress("The email could not be sent to Prescience: " + err.message, "err");
      const retry = el("button", "secondary mini", "Try again"); retry.addEventListener("click", () => sendPrepare());
      $("file-progress").appendChild(retry);
    }
  }
  function schedulePoll(ms) { const job = state.job; if (!job) return; clearTimeout(job.pollTimer); job.pollTimer = setTimeout(pollStatus, ms); }
  async function pollStatus() {
    const job = state.job; if (!job || job.final) return;
    try {
      const r = await callFlow(Object.assign(basePayload("status"), { jobId: job.id }), 40000);
      if (state.job !== job) return;
      if (r.ready && r.contentBase64) {
        const rec = JSON.parse(b64ToText(r.contentBase64));
        if ((rec.seq || 0) > job.seq) { job.seq = rec.seq || 0; job.lastProgress = Date.now(); applyReceipt(rec); }
      }
    } catch (e) { /* a missed poll is fine — the next one catches up */ }
    if (state.job !== job || job.final) return;
    const idle = Date.now() - job.lastProgress;
    if (idle > 240000) {
      setProgress("Prescience hasn't picked this up yet. Is it open on your PC? Open it and it files this at once — or close this pane; it files it the next time Prescience opens.", "warn");
      show("lnk-open", true); schedulePoll(15000); return;
    }
    schedulePoll(job.committing ? 3000 : 5000);
  }
  function applyReceipt(r) {
    const job = state.job; job.receipt = r;
    if (r.documents) renderDocs(r.documents);
    if (r.defaults && r.defaults.date) setIfUntouched("date", r.defaults.date);
    if (r.suggestState === "done" && !job.suggestApplied) { job.suggestApplied = true; applySuggestions(r.suggestions, r.suggestedBy, r.suggestNote); }
    if (r.duplicates && r.duplicates.length && r.phase !== "filed") renderDuplicates(r.duplicates, r.phase === "duplicate");
    switch (r.phase) {
      case "received": setProgress("Prescience has the email — reading it…", "busy"); break;
      case "prepared": setProgress(r.suggestState === "pending" ? "Prescience has read the email ✓ — looking for suggested answers…" : "Prescience has read the email ✓ — answer the questions, then File.", "ok"); break;
      case "waiting": setProgress("Waiting for the email to reach Prescience…", "busy"); break;
      case "filing": setProgress("Filing in Prescience…", "busy"); break;
      case "filed": job.final = true; showFiled(r.filed || {}); return;
      case "duplicate":
        job.committing = false; $("btn-do-file").textContent = "File anyway"; $("dup-ok").checked = true;
        setProgress("Prescience held this back: it looks like something already on file. Press File anyway to file it, or Back to stop.", "warn"); break;
      case "error":
        if (job.committing || !job.sentPrepare || r.missing) { job.committing = false; }
        setProgress(r.error || "Something went wrong in Prescience.", "err");
        if (r.missing && r.missing.length) $("file-problem").textContent = "Prescience says still needed: " + r.missing.join(", ");
        break;
      default: break;
    }
    validate();
  }
  async function submitFiling() {
    const job = state.job; if (!job || job.committing || job.final) return;
    const missing = missingAnswers(); if (missing.length) { validate(); return; }
    job.committing = true; validate();
    openLink(fileLink(job.id));                  // the answers are on their way: wake Prescience again
    const req = { schema: D.jobSchema || "prescience-file/1", jobId: job.id, answers: collectAnswers(), main: chosenMain(), leaveOut: chosenLeaveOut(),
      duplicateOk: $("dup-ok").checked, submittedBy: state.user.email, submittedByName: state.user.name, submittedAt: isoLocal(new Date()) };
    setProgress("Sending your answers to Prescience…", "busy");
    try {
      await callFlow(Object.assign(basePayload("deliver"), { action: "commit", name: `${job.id}.json`, jobId: job.id, contentBase64: textToB64(JSON.stringify(req)) }), 115000);
      if (state.job !== job) return;
      job.lastProgress = Date.now(); setProgress("Filing in Prescience…", "busy"); schedulePoll(2500);
    } catch (err) {
      job.committing = false; validate();
      setProgress("Your answers could not be sent: " + err.message, "err");
    }
  }
  function showFiled(f) {
    $("res-head").textContent = "Filed in Prescience ✓";
    const recs = f.records || [];
    $("res-text").textContent = `${recs.length} document${recs.length !== 1 ? "s" : ""} filed under ${f.entity || "the entity"}${f.filedBy ? " — filed by " + f.filedBy : ""}.`;
    const list = $("res-list"); list.innerHTML = "";
    for (const r of recs) { const row = el("div", "att"); row.append(el("span", "pill " + (r.kind === "email" ? "neutral" : "info"), r.kind === "email" ? "Email" : "Doc"), el("span", "name", r.name)); list.appendChild(row); }
    const extras = [];
    if (f.added && f.added.length) extras.push(`Added to the pick lists: ${f.added.join(", ")}.`);
    if (f.leftOut && f.leftOut.length) extras.push(`Left out: ${f.leftOut.join(", ")}.`);
    $("res-note").textContent = extras.join(" ");
    setProgress("", ""); page("result");
  }
  function cancelFiling() { const job = state.job; if (job) { clearTimeout(job.pollTimer); job.final = true; } state.job = null; page("main"); }

  // ------------------------------------------------------------------ compare (unchanged since v1.1, plus the link)
  async function sendCompare(att) {
    if (state.busy || !state.email) return;
    state.busy = true; renderEmail();
    openLink(D.link.openUrl || "prescience://open");
    const e = state.email;
    try {
      setStatus(`Reading ${att.name}…`, "busy");
      const content = await attachmentBase64(att);
      const name = safeName(att.name, "attachment");
      setStatus("Sending to Prescience…", "busy");
      const res = await callFlow(Object.assign(basePayload("deliver"), { action: "compare", name, messageId: e.restId, attachmentId: restAttachmentId(att.id), contentBase64: content }), 115000);
      setStatus("", "");
      $("res-head").textContent = "Sent";
      $("res-text").textContent = `${name} is on its way to Prescience's Compare box.`;
      $("res-list").innerHTML = "";
      $("res-note").textContent = `Prescience opens with it in the Compare box — pick the entity and the agreements on file there.` + (res.via === "mailbox" ? " (The flow fetched it from your mailbox.)" : "");
      page("result");
    } catch (err) { setStatus(err.message, "err"); }
    finally { state.busy = false; renderEmail(); }
  }

  // ------------------------------------------------------------------ settings
  function renderSettings() { $("set-url").value = state.flowUrl || ""; $("about").textContent = `Prescience add-in ${VERSION} · signed in as ${state.user.email || "?"} · Outlook API: attachments ${state.has18 ? "yes" : "no"}, whole email ${state.has114 ? "yes" : "no"} · questions: ${state.form ? "from Prescience" + (state.form.publishedAt ? " (" + state.form.publishedAt.slice(0, 10) + ")" : "") : "built-in"}`; }
  async function connect(url, fromSetup) {
    const msg = fromSetup ? $("setup-msg") : $("set-msg");
    url = (url || "").trim(); if (!/^https?:\/\//i.test(url)) { msg.textContent = "That doesn't look like a flow address (it starts with https://)."; return false; }
    state.flowUrl = url; msg.textContent = "Connecting…";
    try { const r = await callFlow(basePayload("ping"), 40000); msg.textContent = (!r.mailbox || r.mailbox === "none") ? `The flow answered (it reads no mailbox for ${state.user.email}; filing from here still works).` : `Connected — the flow knows ${state.user.email}.`; }
    catch (e) { msg.textContent = e.message; return false; }
    roam("prescienceFlowUrl", url); refreshForm(); return true;
  }

  // ------------------------------------------------------------------ init
  function wire() {
    $("btn-file").addEventListener("click", startFiling);
    $("btn-compare").addEventListener("click", () => { const c = compareChoice(state.email); if (c.mode === "one") sendCompare(c.good[0]); else if (c.mode === "many") { renderPick(); page("pick"); } });
    $("btn-pick-back").addEventListener("click", () => page("main"));
    $("btn-do-file").addEventListener("click", submitFiling);
    $("btn-file-back").addEventListener("click", cancelFiling);
    $("lnk-open").addEventListener("click", (ev) => { ev.preventDefault(); if (state.job) openLink(fileLink(state.job.id)); });
    $("btn-again").addEventListener("click", () => { state.job = null; state.email = readItem(); renderEmail(); page("main"); });
    $("setup-save").addEventListener("click", async () => { if (await connect($("setup-url").value, true)) page("main"); });
    $("set-save").addEventListener("click", async () => { if (await connect($("set-url").value, false)) setStatus("Flow address saved.", "ok"); });
    $("set-test").addEventListener("click", async () => { await connect($("set-url").value || state.flowUrl, false); });
    $("nav-main").addEventListener("click", (e) => { e.preventDefault(); page(state.flowUrl ? (state.job && !state.job.final ? "file" : "main") : "setup"); });
    $("nav-settings").addEventListener("click", (e) => { e.preventDefault(); page("settings"); });
  }

  // pure helpers exposed for checking from a browser console (no data, no secrets)
  window.__pr = { eligible, compareChoice, safeName, shortError, extOf, normalizeName, exactMatch, splitValues, textToB64, b64ToText,
    missingAnswers, collectAnswers, fileLink, version: VERSION, state };

  function boot() {
    wire();
    const params = new URLSearchParams(location.search);
    state.dev = params.get("dev") === "1";
    if (!state.dev) {
      try { state.has18 = Office.context.requirements.isSetSupported("Mailbox", "1.8"); } catch (e) { state.has18 = false; }
      try { state.has114 = Office.context.requirements.isSetSupported("Mailbox", "1.14"); } catch (e) { state.has114 = false; }
      try { const p = Office.context.mailbox.userProfile; state.user = { name: p.displayName || "", email: (p.emailAddress || "").toLowerCase() }; } catch (e) { state.user = { name: "", email: "" }; }
    } else {
      state.user = { name: params.get("name") || "Test Person", email: (params.get("user") || "test@psomas.com").toLowerCase() };
    }
    // the manifest IT deploys carries the flow address (taskpane.html?flow=...); otherwise pasted once
    const fromManifest = (params.get("flow") || "").trim();
    if (/^https?:\/\//i.test(fromManifest)) { state.flowUrl = fromManifest; if (!state.dev && roam("prescienceFlowUrl") !== fromManifest) roam("prescienceFlowUrl", fromManifest); }
    else state.flowUrl = roam("prescienceFlowUrl") || "";
    state.form = cachedForm();
    state.email = readItem(); renderEmail();
    if (!state.dev) {
      try { Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => { if (state.job && !state.job.final) return; state.email = readItem(); renderEmail(); if (!$("sec-settings").classList.contains("hidden") || !$("sec-setup").classList.contains("hidden")) return; page("main"); }); } catch (e) { /* pinning not supported here */ }
    }
    page(state.flowUrl ? "main" : "setup");
    if (state.flowUrl) refreshForm();
  }
  if (new URLSearchParams(location.search).get("dev") === "1" || typeof Office === "undefined") { if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", boot); else boot(); }
  else Office.onReady(boot);
})();

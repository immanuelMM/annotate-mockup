// Annotator — ⌘-click (Ctrl-click on Windows) any button, card or section to attach
// a description. Saved notes render as callouts with an arrow pointing at the element.
// Notes can carry a custom number (e.g. "1", "2a", "Step 3") and can be connected
// card-to-card with arrows (Step 1 → Step 2 → …) to show a flow.
// Each box can be hidden individually (it collapses to a small numbered pin).
// A screenshot of the element is captured whenever a note is saved, and "Report" turns
// every note (this page or all annotated pages) into a printable PDF or a Word .docx.
// Notes are stored per page in localStorage (screenshots in IndexedDB) and can be exported / imported as JSON.
(() => {
  "use strict";
  if (window.__annotatorLoaded) return;
  window.__annotatorLoaded = true;

  const PAGE = decodeURIComponent(location.pathname.split("/").pop() || "page").replace(/\.dc\.html$|\.html?$/, "");
  const KEY = "annotator:" + PAGE;
  const ACCENT = "#2563eb";
  const H2C_URL = "https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js";
  const DOCX_URL = "https://cdn.jsdelivr.net/npm/docx@9.5.1/dist/index.iife.js";

  let notes = load();
  let visible = true;
  let editor = null;          // { el, note, box }
  let hoverEl = null;
  let modHeld = false;
  let linking = null;         // id of the note we're drawing a connection from

  const numOf = (note, i) => (note.num != null && String(note.num).trim() !== "" ? String(note.num).trim() : String(i + 1));

  function load() {
    try { return JSON.parse(localStorage.getItem(KEY) || "[]"); } catch { return []; }
  }
  function save() {
    try { localStorage.setItem(KEY, JSON.stringify(notes)); } catch {}
    updateToolbar();
  }

  // ---------- element targeting ----------
  const isOurs = (el) => el && el.closest && el.closest("[data-annotator-ui]");
  const textOf = (el) => (el.innerText || el.textContent || "").replace(/\s+/g, " ").trim().slice(0, 80);

  function selectorFor(el) {
    const parts = [];
    while (el && el.nodeType === 1 && el !== document.documentElement) {
      if (el.id) { parts.unshift("#" + CSS.escape(el.id)); break; }
      const parent = el.parentElement;
      if (!parent) break;
      const idx = Array.prototype.indexOf.call(parent.children, el) + 1;
      parts.unshift(el.tagName.toLowerCase() + ":nth-child(" + idx + ")");
      el = parent;
    }
    return parts.join(" > ");
  }

  // A label for elements without text (inputs, icons, images, progress bars…), so they can be found again even
  // when their position in the page structure changes (different wizard state, other rows rendered, …).
  function hintOf(el) {
    const a = (n) => (el.getAttribute(n) || "").trim();
    const h = a("placeholder") || a("aria-label") || a("title") || a("alt") || a("name") ||
      (el.tagName === "IMG" ? (a("src").split("/").pop() || "").split("?")[0] : "");
    return h.slice(0, 80);
  }
  const matches = (el, note) => el.tagName.toLowerCase() === note.tag &&
    (note.text ? textOf(el) === note.text : !note.hint || hintOf(el) === note.hint);
  // Is the element actually visible right now (on screen and not hidden behind e.g. a modal)?
  const seen = (el) => {
    const r = el.getBoundingClientRect();
    return !!(r.width || r.height) && r.bottom > 0 && r.top < innerHeight && r.right > 0 && r.left < innerWidth && !coveredAt(el, r);
  };

  const cache = new Map(), lastScan = new Map();
  function resolve(note) {
    const hit = cache.get(note.id), now = performance.now();
    const hitOk = !!(hit && hit.isConnected && matches(hit, note));
    const recent = now - (lastScan.get(note.id) || 0) < 400;
    if (hitOk && (recent || seen(hit))) return hit;
    if (!hitOk && hit === null && recent) return null; // nothing matched a moment ago: don't rescan every frame
    lastScan.set(note.id, now);
    // Candidates: the saved path, then every element with the same tag + text (or label). When the same text
    // appears more than once (e.g. a team name on the page and inside a modal), prefer the one you can see.
    const cands = [];
    let el = null;
    try { el = document.querySelector(note.selector); } catch {}
    if (el && !isOurs(el) && matches(el, note)) cands.push(el);
    if (note.text || note.hint) {
      const all = document.getElementsByTagName(note.tag);
      for (let i = 0; i < all.length; i++) if (!isOurs(all[i]) && !cands.includes(all[i]) && matches(all[i], note)) cands.push(all[i]);
    }
    const pick = cands.find(seen) || (hitOk ? hit : cands[0]) || null;
    cache.set(note.id, pick);
    return pick;
  }

  // ---------- styles ----------
  const style = document.createElement("style");
  style.textContent = `
  [data-annotator-ui]{font-family:Inter,system-ui,-apple-system,sans-serif;font-size:13px;line-height:1.4;color:#0e0e10;box-sizing:border-box}
  [data-annotator-ui] *{box-sizing:border-box}
  .ann-layer{position:fixed;inset:0;pointer-events:none;z-index:2147483000}
  .ann-layer svg{position:absolute;inset:0;width:100%;height:100%;overflow:visible}
  .ann-hl{position:fixed;pointer-events:none;z-index:2147483001;border:2px solid ${ACCENT};background:rgba(37,99,235,.08);border-radius:4px;transition:all .06s}
  .ann-ring{position:absolute;border:2px dashed ${ACCENT};border-radius:4px}
  .ann-card{position:absolute;pointer-events:auto;width:260px;min-width:140px;min-height:36px;max-width:70vw;max-height:70vh;background:#fff;border:1.5px solid ${ACCENT};border-radius:8px;
    padding:8px 10px;box-shadow:0 6px 20px rgba(20,20,40,.18);cursor:grab;white-space:pre-wrap;word-break:break-word;resize:both;overflow:auto}
  .ann-card:active{cursor:grabbing}
  .ann-card.ann-snapping{resize:none;box-shadow:none}
  .ann-card.ann-snapping .ann-acts{display:none !important}
  .ann-card.ann-snapping .ann-chip{display:inline-block;vertical-align:middle}
  .ann-card.ann-snapping .ann-ic{display:inline-block;height:16px;line-height:16px;text-align:center;vertical-align:-3px;margin-right:5px}
  .ann-num{float:left;margin:0 7px 0 -3px;min-width:17px;height:17px;padding:0 5px;border-radius:9px;background:${ACCENT};color:#fff;font-size:10px;font-weight:700;display:flex;align-items:center;justify-content:center;white-space:nowrap;cursor:text}
  .ann-card.ann-src{box-shadow:0 0 0 3px rgba(37,99,235,.35),0 6px 20px rgba(20,20,40,.18)}
  .ann-linking .ann-card:not(.ann-src){cursor:copy;outline:2px dashed ${ACCENT};outline-offset:3px}
  .ann-linking .ann-card:not(.ann-src):hover .ann-acts{display:none}
  .ann-link-hit{pointer-events:stroke;cursor:pointer}
  .ann-link-hit:hover + .ann-link{stroke:#e1251b}
  .ann-banner{position:fixed;left:50%;top:14px;transform:translateX(-50%);z-index:2147483003;background:${ACCENT};color:#fff;border-radius:8px;padding:7px 12px;font-weight:600;box-shadow:0 4px 14px rgba(20,20,40,.2);display:none}
  .ann-card .ann-acts{display:none;position:absolute;top:4px;right:4px;gap:4px;background:#fff;padding:2px;border-radius:6px}
  .ann-card:hover .ann-acts{display:flex}
  .ann-acts button,.ann-editor button,.ann-bar button{border:1px solid #d8d8de;background:#fff;border-radius:5px;font:inherit;font-size:12px;cursor:pointer;padding:3px 8px;color:#0e0e10}
  .ann-acts button{padding:1px 6px;font-size:11px;box-shadow:0 1px 3px rgba(0,0,0,.12)}
  .ann-editor{position:fixed;z-index:2147483002;width:300px;height:190px;min-width:220px;min-height:150px;max-width:80vw;max-height:80vh;background:#fff;border:1.5px solid ${ACCENT};border-radius:10px;padding:10px;box-shadow:0 10px 30px rgba(20,20,40,.25);resize:both;overflow:auto;display:flex;flex-direction:column}
  .ann-editor .ann-head{flex:none;display:flex;align-items:center;gap:8px;margin-bottom:6px}
  .ann-editor .ann-head label{flex:none;font-size:11px;font-weight:600;color:#6b6b72;display:flex;align-items:center;gap:4px}
  .ann-editor .ann-numin{width:64px;border:1px solid #d8d8de;border-radius:5px;padding:2px 6px;font:inherit;font-size:12px;outline:none}
  .ann-editor .ann-numin:focus{border-color:${ACCENT}}
  .ann-editor .ann-target{flex:1;min-width:0;font-size:11px;color:#6b6b72;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
  .ann-editor textarea{flex:1;width:100%;min-height:60px;resize:none;border:1px solid #d8d8de;border-radius:6px;padding:7px;font:inherit;outline:none}
  .ann-editor textarea:focus{border-color:${ACCENT};box-shadow:0 0 0 3px rgba(37,99,235,.15)}
  .ann-editor .ann-row{flex:none;display:flex;gap:6px;margin-top:8px;align-items:center}
  .ann-editor .ann-row .ann-sp{flex:1;font-size:11px;color:#9a9aa1}
  .ann-editor .ann-primary,.ann-bar .ann-primary{background:${ACCENT};border-color:${ACCENT};color:#fff}
  .ann-editor .ann-danger{color:#e1251b}
  .ann-bar{position:fixed;left:14px;bottom:14px;z-index:2147483002;display:flex;gap:6px;align-items:center;background:#fff;border:1px solid #e5e5e9;border-radius:8px;padding:6px 8px;box-shadow:0 4px 14px rgba(20,20,40,.14)}
  .ann-bar .ann-lbl{font-weight:600;font-size:12px;padding:0 4px}
  .ann-bar .ann-hint{font-size:11px;color:#6b6b72;padding-right:4px}
  .ann-pin{position:absolute;pointer-events:auto;min-width:18px;height:18px;padding:0 5px;border-radius:9px;background:#fff;border:1.5px dashed ${ACCENT};color:${ACCENT};font-size:10px;font-weight:700;display:flex;align-items:center;justify-content:center;white-space:nowrap;cursor:pointer;box-shadow:0 2px 6px rgba(20,20,40,.15)}
  .ann-pin:hover{background:${ACCENT};color:#fff}
  .ann-list{position:fixed;top:12px;right:12px;bottom:12px;z-index:2147483003;width:380px;max-width:calc(100vw - 24px);background:#fff;border:1px solid #e5e5e9;border-radius:12px;box-shadow:0 14px 40px rgba(20,20,40,.22);display:none;flex-direction:column;overflow:hidden}
  .ann-list.open{display:flex}
  .ann-list .lh{flex:none;display:flex;align-items:center;gap:8px;padding:12px 14px;border-bottom:1px solid #ececf0}
  .ann-list .lh b{font-size:14px;flex:1}
  .ann-list .lf{flex:none;display:flex;flex-wrap:wrap;gap:6px;align-items:center;padding:8px 14px;border-bottom:1px solid #ececf0;background:#fafafb;font-size:12px}
  .ann-list .lf label{display:flex;align-items:center;gap:5px;cursor:pointer;margin-right:auto}
  .ann-list .lb{flex:1;overflow:auto}
  .ann-list .row{display:flex;gap:9px;padding:10px 14px;border-bottom:1px solid #f1f1f3;align-items:flex-start}
  .ann-list .row:hover{background:#fafbff}
  .ann-list .row input[type=checkbox]{margin-top:3px}
  .ann-list .mn{flex:1;min-width:0}
  .ann-list .tl{display:flex;gap:6px;align-items:center}
  .ann-list .tl .t{font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
  .ann-list .nb{flex:none;min-width:17px;height:17px;padding:0 5px;border-radius:9px;background:${ACCENT};color:#fff;font-size:10px;font-weight:700;display:inline-flex;align-items:center;justify-content:center;white-space:nowrap}
  .ann-list .el{font-size:11px;color:#8a8a92;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;margin-top:2px}
  .ann-list .st{display:inline-block;font-size:10.5px;font-weight:600;border-radius:9px;padding:1px 7px;margin-top:4px}
  .ann-list .st.ok{background:#e7f5ed;color:#1f7a44}.ann-list .st.hid{background:#eef0f4;color:#555}
  .ann-list .st.off{background:#eaf1ff;color:#1d4ed8}.ann-list .st.cov{background:#fff4e0;color:#9a5b00}.ann-list .st.miss{background:#fde8e7;color:#b81d14}
  .ann-list .ac{display:flex;flex-wrap:wrap;gap:4px;margin-top:6px}
  .ann-list button{border:1px solid #d8d8de;background:#fff;border-radius:5px;font:inherit;font-size:11.5px;cursor:pointer;padding:2px 8px;color:#0e0e10}
  .ann-list button:disabled{opacity:.4;cursor:default}
  .ann-list button.del{color:#e1251b}
  .ann-list button.pr{background:${ACCENT};border-color:${ACCENT};color:#fff}
  .ann-list textarea{width:100%;min-height:70px;margin-top:6px;border:1px solid #d8d8de;border-radius:6px;padding:6px;font:inherit;font-size:12px;resize:vertical}
  .ann-list .ni{width:90px;border:1px solid #d8d8de;border-radius:5px;padding:2px 6px;font:inherit;font-size:12px;margin-top:6px}
  .ann-list .empty{padding:30px 14px;color:#8a8a92;text-align:center}
  .ann-card.ann-flash{animation:annCardFlash 1.2s ease 2}
  @keyframes annCardFlash{50%{box-shadow:0 0 0 6px rgba(37,99,235,.45),0 6px 20px rgba(20,20,40,.18)}}
  .ann-pop{position:fixed;left:14px;bottom:62px;z-index:2147483003;width:300px;background:#fff;border:1px solid #e5e5e9;border-radius:10px;padding:12px;box-shadow:0 10px 30px rgba(20,20,40,.2);display:none}
  .ann-pop h4{margin:0 0 8px;font-size:13px}
  .ann-pop label{display:flex;gap:8px;align-items:flex-start;padding:6px 4px;border-radius:6px;cursor:pointer}
  .ann-pop label:hover{background:#f5f7ff}
  .ann-pop small{display:block;color:#6b6b72;font-size:11px}
  .ann-pop .ann-row{display:flex;gap:6px;justify-content:flex-end;margin-top:10px}
  .ann-pop button{border:1px solid #d8d8de;background:#fff;border-radius:5px;font:inherit;font-size:12px;cursor:pointer;padding:4px 10px;color:#0e0e10}
  .ann-pop .ann-primary{background:${ACCENT};border-color:${ACCENT};color:#fff}
  body.ann-picking, body.ann-picking *{cursor:crosshair !important}
  `;

  // ---------- DOM scaffolding ----------
  const layer = document.createElement("div");
  layer.className = "ann-layer"; layer.setAttribute("data-annotator-ui", "");
  const svgNS = "http://www.w3.org/2000/svg";
  const svg = document.createElementNS(svgNS, "svg");
  svg.innerHTML = `<defs><marker id="ann-arrow" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="8" markerHeight="8" orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" fill="${ACCENT}"/></marker></defs>`;
  layer.appendChild(svg);

  const hl = document.createElement("div");
  hl.className = "ann-hl"; hl.setAttribute("data-annotator-ui", ""); hl.style.display = "none";

  const bar = document.createElement("div");
  bar.className = "ann-bar"; bar.setAttribute("data-annotator-ui", "");
  const banner = document.createElement("div");
  banner.className = "ann-banner"; banner.setAttribute("data-annotator-ui", "");
  banner.textContent = "Click another note to connect it · click the same pair again to disconnect · Esc to cancel";
  bar.innerHTML = `<span class="ann-lbl">Notes <span data-count>0</span></span>
    <span class="ann-hint">${/Mac|iPhone|iPad/.test(navigator.platform) ? "⌘" : "Ctrl"}+click to add</span>
    <button data-act="list" title="See every note on this page — also ones whose box you can't see — and edit, re-attach or delete them">All notes</button>
    <button data-act="toggle">Hide</button>
    <button data-act="unhide" title="Show the boxes you hid one by one">Show hidden</button>
    <button data-act="chain" title="Connect notes in number order: 1 → 2 → 3 …">Connect in order</button>
    <button data-act="report" class="ann-primary" title="Generate a PDF or Word report of every note, with screenshots">Report</button>
    <button data-act="export">Export</button>
    <button data-act="import">Import</button>
    <button data-act="clear">Clear</button>
    <input type="file" accept="application/json" style="display:none">`;
  const pop = document.createElement("div");
  pop.className = "ann-pop"; pop.setAttribute("data-annotator-ui", "");

  function mount() {
    style.textContent += ".ann-card " + HL_CSS + CHIP_CSS.replace(/\.ann-chip/g, ".ann-card .ann-chip").replace(/\.ann-ic/g, ".ann-card .ann-ic");
    document.head.appendChild(style);
    document.body.appendChild(layer);
    document.body.appendChild(hl);
    document.body.appendChild(bar);
    document.body.appendChild(banner);
    document.body.appendChild(pop);
    document.body.appendChild(list);
    updateToolbar();
    requestAnimationFrame(tick);
  }

  function updateToolbar() {
    bar.querySelector("[data-count]").textContent = notes.length;
    bar.querySelector('[data-act="toggle"]').textContent = visible ? "Hide" : "Show";
    const hidden = notes.filter((n) => n.hidden).length;
    const un = bar.querySelector('[data-act="unhide"]');
    un.style.display = hidden ? "" : "none"; un.textContent = "Show hidden (" + hidden + ")";
  }

  bar.addEventListener("click", (e) => {
    const act = e.target.closest("button")?.dataset.act;
    if (act === "toggle") { visible = !visible; updateToolbar(); }
    if (act === "chain") chainInOrder();
    if (act === "list") toggleList();
    if (act === "unhide") { notes.forEach((n) => { n.hidden = false; }); visible = true; save(); }
    if (act === "report") showReportMenu();
    if (act === "export") exportNotes();
    if (act === "import") bar.querySelector("input").click();
    if (act === "clear" && notes.length && confirm("Delete all " + notes.length + " notes on this page?")) { notes.forEach((n) => shotDel(n.id)); notes = []; save(); }
  });
  bar.querySelector("input").addEventListener("change", async (e) => {
    const f = e.target.files[0]; if (!f) return;
    try {
      const data = JSON.parse(await f.text());
      const incoming = Array.isArray(data) ? data : data.notes;
      if (!Array.isArray(incoming)) throw new Error();
      const ids = new Set(notes.map((n) => n.id));
      notes = notes.concat(incoming.filter((n) => n && n.id && !ids.has(n.id)));
      save();
      for (const [id, shot] of Object.entries(data.images || {})) if (shot && shot.src) await shotPut(id, shot);
    } catch { alert("That file isn't a valid notes export."); }
    e.target.value = "";
  });

  // ---------- picking (⌘/Ctrl + click) ----------
  const isMod = (e) => e.metaKey || e.ctrlKey;

  function setPicking(on) {
    modHeld = on;
    document.body.classList.toggle("ann-picking", on && !editor);
    if (!on) { hl.style.display = "none"; hoverEl = null; }
  }
  window.addEventListener("keydown", (e) => {
    if (e.key === "Meta" || e.key === "Control") setPicking(true);
    if (e.key === "Escape" && document.querySelector(".ann-linkform")) return; // the link dialog handles its own Esc
    if (e.key === "Escape" && editor) closeEditor();
    if (e.key === "Escape" && linking) setLinking(null);
    if (e.key === "Escape" && reattach) setReattach(null);
  }, true);
  window.addEventListener("keyup", (e) => { if (e.key === "Meta" || e.key === "Control") setPicking(false); }, true);
  window.addEventListener("blur", () => setPicking(false));

  document.addEventListener("mousemove", (e) => {
    if (!isMod(e) || editor) { if (hoverEl) setPicking(false); return; }
    if (!modHeld) setPicking(true);
    const el = e.target;
    if (isOurs(el) || el === document.body || el === document.documentElement) { hl.style.display = "none"; hoverEl = null; return; }
    hoverEl = el;
    const r = el.getBoundingClientRect();
    Object.assign(hl.style, { display: "block", left: r.left - 2 + "px", top: r.top - 2 + "px", width: r.width + 4 + "px", height: r.height + 4 + "px" });
  }, true);

  // Swallow the whole click sequence so the underlying button/link doesn't fire.
  const swallow = (e) => {
    if (!isMod(e) || isOurs(e.target)) return;
    e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
  };
  ["pointerdown", "mousedown", "pointerup", "mouseup", "auxclick"].forEach((t) => document.addEventListener(t, swallow, true));
  document.addEventListener("click", (e) => {
    if (!isMod(e) || isOurs(e.target)) return;
    swallow(e);
    const el = e.target;
    if (el === document.body || el === document.documentElement) return;
    if (reattach) { attachTo(reattach, el); return; }
    const existing = notes.find((n) => resolve(n) === el);
    openEditor(el, existing);
  }, true);

  // ---------- editor ----------
  function openEditor(el, note, focusNum) {
    closeEditor();
    setPicking(false);
    const box = document.createElement("div");
    box.className = "ann-editor"; box.setAttribute("data-annotator-ui", "");
    const label = el.tagName.toLowerCase() + (textOf(el) ? " · “" + textOf(el).slice(0, 40) + "”" : "");
    box.innerHTML = `<div class="ann-head"><label>No. <input class="ann-numin" maxlength="12"></label><div class="ann-target"></div></div>
      <textarea placeholder="First line = title, then the description…&#10;Add links with 🔗 Link · select text + 🖍 Highlight to flag a question."></textarea>
      <div class="ann-row">
        ${note ? '<button class="ann-danger" data-act="del">Delete</button>' : ""}
        <span class="ann-sp">${/Mac/.test(navigator.platform) ? "⌘" : "Ctrl"}+Enter to save</span>
        <button data-act="hl" title="Highlight the selected text — e.g. a question to check (${/Mac/.test(navigator.platform) ? "⌘⇧" : "Ctrl+Shift+"}H). Click again inside a highlight to remove it.">🖍 Highlight</button>
        <button data-act="link" title="Insert a link with its own title (${/Mac/.test(navigator.platform) ? "⌘" : "Ctrl"}+K)">🔗 Link</button>
        <button data-act="cancel">Cancel</button>
        <button class="ann-primary" data-act="save">Save</button>
      </div>`;
    box.querySelector(".ann-target").textContent = label;
    const ta = box.querySelector("textarea");
    ta.value = note ? note.body : "";
    const numIn = box.querySelector(".ann-numin");
    numIn.placeholder = note ? String(notes.indexOf(note) + 1) : String(notes.length + 1);
    numIn.value = note && note.num != null ? note.num : "";
    numIn.title = "Leave blank to number automatically";
    document.body.appendChild(box);

    const r = el.getBoundingClientRect();
    const W = 300, H = box.offsetHeight;
    let left = r.right + 14, top = r.top;
    if (left + W > innerWidth - 8) left = r.left - W - 14;
    if (left < 8) { left = Math.max(8, Math.min(r.left, innerWidth - W - 8)); top = r.bottom + 10; }
    if (top + H > innerHeight - 8) top = Math.max(8, innerHeight - H - 8);
    Object.assign(box.style, { left: left + "px", top: Math.max(8, top) + "px" });

    editor = { el, note, box };
    const ring = document.createElement("div");
    ring.className = "ann-hl"; ring.setAttribute("data-annotator-ui", "");
    Object.assign(ring.style, { left: r.left - 2 + "px", top: r.top - 2 + "px", width: r.width + 4 + "px", height: r.height + 4 + "px" });
    document.body.appendChild(ring);
    editor.ring = ring;

    const commit = () => {
      const body = ta.value.trim();
      if (!body) { if (note) remove(note.id); closeEditor(); return; }
      const num = numIn.value.trim() || null;
      let saved = note;
      if (note) { note.body = body; note.num = num; }
      else {
        const n = { id: Date.now().toString(36) + Math.random().toString(36).slice(2, 6), selector: selectorFor(el),
          tag: el.tagName.toLowerCase(), text: textOf(el), hint: hintOf(el), body, num, links: [], dx: null, dy: null, created: new Date().toISOString() };
        notes.push(n); cache.set(n.id, el); saved = n;
      }
      save(); visible = true; updateToolbar(); closeEditor();
      snapSoon(saved, 200); // after the callout has rendered, so the screenshot includes it
    };
    box.addEventListener("click", (e) => {
      const act = e.target.closest("button")?.dataset.act;
      if (act === "save") commit();
      if (act === "cancel") closeEditor();
      if (act === "link") { const r = e.target.closest("button").getBoundingClientRect(); editorLink(r.left, r.bottom + 6); }
      if (act === "hl") toggleHighlight();
      if (act === "del") { remove(note.id); closeEditor(); }
    });
    [ta, numIn].forEach((f) => f.addEventListener("keydown", (e) => { if (e.key === "Enter" && (e.metaKey || e.ctrlKey || f === numIn)) { e.preventDefault(); commit(); } }));
    // Insert / edit a link at the cursor: prefilled from the link under the cursor, or the selected text as its title.
    const editorLink = (x, y) => {
      const s = ta.selectionStart, en = ta.selectionEnd, hit = linkAt(ta.value, s, en);
      linkForm(document, x, y, hit ? { title: hit.label, url: hit.url } : { title: ta.value.slice(s, en).trim(), url: "" }, (title, url) => {
        const before = ta.value;
        ta.value = applyLink(before, s, en, hit, title, url);
        const pos = (hit ? hit.start : s) + fmtLink(title, url).length;
        ta.focus(); ta.setSelectionRange(pos, pos);
      }, () => ta.focus());
    };
    // Highlight toggle: inside a highlight → remove it; otherwise highlight the selection (each selected line
    // separately, keeping "- " bullets outside), or the current line when nothing is selected.
    const toggleHighlight = () => {
      const v = ta.value; let s = ta.selectionStart, en = ta.selectionEnd, m;
      HL_RE.lastIndex = 0;
      while ((m = HL_RE.exec(v))) {
        const a = m.index, b = a + m[0].length;
        if (s >= a && en <= b) { ta.value = v.slice(0, a) + m[1] + v.slice(b); ta.focus(); ta.setSelectionRange(a, a + m[1].length); return; }
      }
      if (s === en) { s = v.lastIndexOf("\n", s - 1) + 1; en = v.indexOf("\n", en); if (en < 0) en = v.length; }
      const wrapped = v.slice(s, en).split("\n").map((line) => {
        const mm = line.match(/^(\s*(?:[-•*]\s+)?)(.*?)(\s*)$/);
        return mm[2] ? mm[1] + "==" + mm[2].replace(/==/g, "") + "==" + mm[3] : line;
      }).join("\n");
      ta.value = v.slice(0, s) + wrapped + v.slice(en);
      ta.focus(); ta.setSelectionRange(s, s + wrapped.length);
    };
    ta.addEventListener("keydown", (e) => {
      if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === "k") { e.preventDefault(); const r = ta.getBoundingClientRect(); editorLink(r.left, r.bottom + 6); }
      if ((e.metaKey || e.ctrlKey) && e.shiftKey && e.key.toLowerCase() === "h") { e.preventDefault(); toggleHighlight(); }
    });
    setTimeout(() => (focusNum ? (numIn.focus(), numIn.select()) : ta.focus()), 0);
  }
  function closeEditor() {
    if (!editor) return;
    closeLinkUI(document);
    editor.box.remove(); editor.ring && editor.ring.remove();
    editor = null;
  }
  function remove(id) {
    shotDel(id);
    notes = notes.filter((n) => n.id !== id);
    notes.forEach((n) => { if (n.links) n.links = n.links.filter((t) => t !== id); });
    cache.delete(id); if (linking === id) setLinking(null); save();
  }

  // ---------- connections between notes ----------
  function setLinking(id) {
    linking = id;
    document.body.classList.toggle("ann-linking", !!id);
    banner.style.display = id ? "block" : "none";
    for (const [nid, v] of cards) v.card.classList.toggle("ann-src", nid === id);
  }
  function toggleLink(fromId, toId) {
    const a = notes.find((n) => n.id === fromId), b = notes.find((n) => n.id === toId);
    if (!a || !b || a === b) return;
    a.links = a.links || []; b.links = b.links || [];
    if (a.links.includes(toId)) a.links = a.links.filter((t) => t !== toId);
    else if (b.links.includes(fromId)) b.links = b.links.filter((t) => t !== fromId);
    else a.links.push(toId);
    save();
  }
  // Connect every note to the next one by number: 1 → 2 → 3 … (numbers like "2a" or "Step 3" sort by their digits).
  function ordered(list) {
    // Compare every number in the label in turn, so "FR-1.2" < "FR-1.10" < "FR-2.1"; labels without digits go last.
    const key = (n, i) => (numOf(n, i).match(/\d+/g) || []).map(Number);
    const cmp = (a, b) => {
      if (!a.length || !b.length) return (!a.length) - (!b.length);
      for (let j = 0; j < Math.max(a.length, b.length); j++) { const d = (a[j] ?? -1) - (b[j] ?? -1); if (d) return d; }
      return 0;
    };
    return list.map((n, i) => ({ n, i, k: key(n, i), label: numOf(n, i) })).sort((a, b) => cmp(a.k, b.k) || a.i - b.i);
  }
  function chainInOrder() {
    if (notes.length < 2) return;
    const sorted = ordered(notes).map((x) => x.n);
    for (let i = 0; i < sorted.length - 1; i++) {
      const a = sorted[i], b = sorted[i + 1];
      a.links = a.links || [];
      if (!a.links.includes(b.id) && !(b.links || []).includes(a.id)) a.links.push(b.id);
    }
    save();
  }
  document.addEventListener("pointerdown", (e) => { if (linking && !e.target.closest(".ann-card")) setLinking(null); }, true);

  // ---------- rendering callouts + arrows ----------
  const cards = new Map(); // id -> { card, path, ring }
  let drag = null;

  function ensure(note, i) {
    let v = cards.get(note.id);
    if (!v) {
      const card = document.createElement("div");
      card.className = "ann-card";
      card.innerHTML = `<span class="ann-num"></span><span class="ann-body"></span>
        <span class="ann-acts"><button data-act="link" title="Connect this note to another one">Connect →</button><button data-act="edit">Edit</button><button data-act="hide" title="Hide this box (click its pin to show it again)">Hide</button><button data-act="del">✕</button></span>`;
      const path = document.createElementNS(svgNS, "path");
      path.setAttribute("fill", "none"); path.setAttribute("stroke", ACCENT); path.setAttribute("stroke-width", "2");
      path.setAttribute("marker-end", "url(#ann-arrow)");
      const ring = document.createElement("div"); ring.className = "ann-ring";
      const pin = document.createElement("div"); pin.className = "ann-pin";
      pin.addEventListener("click", () => { const n = notes.find((x) => x.id === note.id); if (n) { n.hidden = false; save(); } });
      svg.appendChild(path); layer.appendChild(ring); layer.appendChild(pin); layer.appendChild(card);
      if (note.w) card.style.width = note.w + "px";
      if (note.h) card.style.height = note.h + "px";
      v = { card, path, ring, pin };
      cards.set(note.id, v);

      // Persist manual box resizes (drag from the bottom-right corner) per note.
      if (window.ResizeObserver) {
        let saveTimer = null;
        new ResizeObserver(() => {
          const n = notes.find((x) => x.id === note.id); if (!n || !card.offsetWidth) return; // 0×0 = box hidden, not resized
          if (n.w === card.offsetWidth && n.h === card.offsetHeight) return;
          const resized = n.w != null;
          n.w = card.offsetWidth; n.h = card.offsetHeight;
          clearTimeout(saveTimer); saveTimer = setTimeout(save, 250);
          if (resized) snapSoon(n);
        }).observe(card);
      }

      card.addEventListener("click", (e) => {
        const act = e.target.closest("button")?.dataset.act;
        const n = notes.find((x) => x.id === note.id);
        if (!n) return;
        if (linking && !act) { if (linking !== n.id) toggleLink(linking, n.id); setLinking(null); return; }
        if (act === "link") { setLinking(linking === n.id ? null : n.id); return; }
        if (e.target.closest(".ann-num")) { const el = resolve(n); if (el) openEditor(el, n, true); return; }
        if (act === "del") remove(n.id);
        if (act === "hide") { n.hidden = true; save(); }
        if (act === "edit") { const el = resolve(n); if (el) openEditor(el, n); }
      });
      card.addEventListener("dblclick", () => { const n = notes.find((x) => x.id === note.id); const el = n && resolve(n); if (el) openEditor(el, n); });
      card.addEventListener("pointerdown", (e) => {
        if (e.target.closest("button, .ann-num, a") || linking) return;
        // Bottom-right corner is the native resize handle — let the browser handle it, don't start a drag.
        const cr = card.getBoundingClientRect();
        if (e.clientX > cr.right - 16 && e.clientY > cr.bottom - 16) return;
        const n = notes.find((x) => x.id === note.id); const el = n && resolve(n); if (!el) return;
        const r = el.getBoundingClientRect();
        drag = { n, r, sx: e.clientX, sy: e.clientY, dx0: card.offsetLeft - r.right, dy0: card.offsetTop - r.top };
        card.setPointerCapture(e.pointerId); e.preventDefault();
      });
      card.addEventListener("pointermove", (e) => {
        if (!drag || drag.n.id !== note.id) return;
        drag.n.dx = drag.dx0 + e.clientX - drag.sx;
        drag.n.dy = drag.dy0 + e.clientY - drag.sy;
      });
      card.addEventListener("pointerup", () => {
        if (!drag || drag.n.id !== note.id) return;
        const moved = drag.n.dx !== drag.dx0 || drag.n.dy !== drag.dy0, n = drag.n;
        drag = null; save();
        if (moved) snapSoon(n); // keep the screenshot matching where the box now sits
      });
    }
    v.card.querySelector(".ann-num").textContent = numOf(note, i);
    v.pin.textContent = numOf(note, i);
    v.pin.title = "Show note " + numOf(note, i) + ": " + note.body.slice(0, 80);
    if (v.body !== note.body) { v.body = note.body; v.card.querySelector(".ann-body").innerHTML = linkifyHTML(note.body); } // only when changed, so chips stay clickable
    return v;
  }

  // Arrow from nearest card edge to nearest point on the element's border (shared with screenshots).
  function arrowGeom(r, x, y, cw, ch) {
    const cx = x + cw / 2, cy = y + ch / 2;
    const tx = Math.max(r.left, Math.min(cx, r.right)), ty = Math.max(r.top, Math.min(cy, r.bottom));
    const inside = cx > r.left && cx < r.right && cy > r.top && cy < r.bottom;
    const ex = inside ? r.left + r.width / 2 : tx, ey = inside ? r.top : ty;
    const sx = Math.max(x, Math.min(ex, x + cw)), sy = Math.max(y, Math.min(ey, y + ch));
    return { sx, sy, mx: (sx + ex) / 2, my: (sy + ey) / 2 - 18, ex, ey };
  }

  const links = new Map(); // "from>to" -> { hit, path }
  function edgePoint(b, tx, ty) {
    // Point where the line from b's centre toward (tx,ty) leaves box b.
    const cx = b.x + b.w / 2, cy = b.y + b.h / 2, dx = tx - cx, dy = ty - cy;
    if (!dx && !dy) return [cx, cy];
    const s = Math.min(dx ? (b.w / 2) / Math.abs(dx) : Infinity, dy ? (b.h / 2) / Math.abs(dy) : Infinity);
    return [cx + dx * s, cy + dy * s];
  }

  // Is the element hidden behind something else — e.g. a modal's overlay, or a sticky header it scrolled under?
  // Then its box/arrow/pin are hidden too, so boxes of the page behind an open modal don't pile on top of the
  // modal and bury the modal's own boxes. Checked by hit-testing a few points of the element (throttled).
  const coverCache = new Map(); // id -> { t, el, covered }
  function isCovered(id, el, r) {
    const now = performance.now(), c = coverCache.get(id);
    if (c && c.el === el && now - c.t < 150) return c.covered;
    const covered = coveredAt(el, r);
    coverCache.set(id, { t: now, el, covered });
    return covered;
  }
  function coveredAt(el, r) {
    const vw = document.documentElement.clientWidth, vh = document.documentElement.clientHeight;
    const pts = [[0.5, 0.5], [0.2, 0.25], [0.8, 0.25], [0.2, 0.75], [0.8, 0.75]]
      .map(([fx, fy]) => [r.left + r.width * fx, r.top + r.height * fy])
      .filter(([x, y]) => x >= 0 && y >= 0 && x < vw && y < vh);
    for (const [x, y] of pts) {
      const hit = document.elementsFromPoint(x, y).find((n) => !isOurs(n));
      if (!hit || hit === el || el.contains(hit) || hit.contains(el)) return false;
    }
    return pts.length > 0;
  }

  function tick() {
    const live = new Set();
    const boxes = new Map();
    notes.forEach((note, i) => {
      live.add(note.id);
      const v = ensure(note, i);
      const el = visible ? resolve(note) : null;
      const r = el && el.getBoundingClientRect();
      const onScreen = r && (r.width || r.height) && r.bottom > -40 && r.top < innerHeight + 40 && !isCovered(note.id, el, r);
      const shown = onScreen && !note.hidden;
      v.card.style.display = v.ring.style.display = shown ? "" : "none";
      v.path.style.display = shown ? "" : "none";
      v.pin.style.display = onScreen && note.hidden ? "" : "none";
      if (onScreen && note.hidden) Object.assign(v.pin.style, { left: Math.max(2, r.left - 9) + "px", top: Math.max(2, r.top - 9) + "px" });
      if (!shown) return;

      Object.assign(v.ring.style, { left: r.left - 3 + "px", top: r.top - 3 + "px", width: r.width + 6 + "px", height: r.height + 6 + "px" });

      const cw = v.card.offsetWidth, ch = v.card.offsetHeight;
      let x, y;
      if (note.dx != null) { x = r.right + note.dx; y = r.top + note.dy; }
      else {
        // Default placement: to the right, else left, else below.
        x = r.right + 40; y = r.top - ch - 16;
        if (x + cw > innerWidth - 8) x = r.left - cw - 40;
        if (x < 8) { x = Math.min(Math.max(8, r.left), innerWidth - cw - 8); y = r.bottom + 30; }
        if (y < 8) y = r.bottom + 30;
        // Don't land on top of a box already placed this frame (e.g. two notes on neighbouring lines):
        // move down past it, or up if that would leave the screen. Dragged boxes keep their spot.
        for (let tries = 0; tries < 12; tries++) {
          const hit = [...boxes.values()].find((b) => x < b.x + b.w + 8 && x + cw + 8 > b.x && y < b.y + b.h + 8 && y + ch + 8 > b.y);
          if (!hit) break;
          y = hit.y + hit.h + 10 + ch > innerHeight - 8 && hit.y - ch - 10 > 8 ? hit.y - ch - 10 : hit.y + hit.h + 10;
        }
      }
      v.card.style.left = x + "px"; v.card.style.top = y + "px";
      boxes.set(note.id, { x, y, w: cw, h: ch });

      const { sx, sy, mx, my, ex, ey } = arrowGeom(r, x, y, cw, ch);
      v.path.setAttribute("d", `M${sx},${sy} Q${mx},${my} ${ex},${ey}`);
      v.path.style.display = Math.hypot(ex - sx, ey - sy) < 6 ? "none" : "";
    });
    for (const [id, v] of cards) if (!live.has(id)) { v.card.remove(); v.path.remove(); v.ring.remove(); v.pin.remove(); cards.delete(id); }

    // Card-to-card connectors (dashed arrows).
    const liveLinks = new Set();
    notes.forEach((a) => (a.links || []).forEach((toId) => {
      const k = a.id + ">" + toId; liveLinks.add(k);
      let l = links.get(k);
      if (!l) {
        const hit = document.createElementNS(svgNS, "path");
        hit.setAttribute("class", "ann-link-hit"); hit.setAttribute("fill", "none");
        hit.setAttribute("stroke", "transparent"); hit.setAttribute("stroke-width", "12");
        const t = document.createElementNS(svgNS, "title"); t.textContent = "Click to remove this connection"; hit.appendChild(t);
        const path = document.createElementNS(svgNS, "path");
        path.setAttribute("class", "ann-link"); path.setAttribute("fill", "none"); path.setAttribute("stroke", ACCENT);
        path.setAttribute("stroke-width", "2"); path.setAttribute("stroke-dasharray", "6 4"); path.setAttribute("marker-end", "url(#ann-arrow)");
        hit.addEventListener("click", () => {
          const from = notes.find((n) => n.id === a.id);
          if (from && confirm("Remove this connection?")) { from.links = (from.links || []).filter((x) => x !== toId); save(); }
        });
        svg.appendChild(hit); svg.appendChild(path);
        l = { hit, path }; links.set(k, l);
      }
      const A = boxes.get(a.id), B = boxes.get(toId);
      const on = A && B && cards.get(a.id).card.style.display !== "none" && cards.get(toId).card.style.display !== "none";
      l.hit.style.display = l.path.style.display = on ? "" : "none";
      if (!on) return;
      const [sx, sy] = edgePoint(A, B.x + B.w / 2, B.y + B.h / 2);
      const [ex, ey] = edgePoint(B, A.x + A.w / 2, A.y + A.h / 2);
      // Gentle bend perpendicular to the line so overlapping connectors stay readable.
      const len = Math.hypot(ex - sx, ey - sy) || 1, bend = Math.min(40, len * 0.15);
      const mx = (sx + ex) / 2 - ((ey - sy) / len) * bend, my = (sy + ey) / 2 + ((ex - sx) / len) * bend;
      const d = `M${sx},${sy} Q${mx},${my} ${ex},${ey}`;
      l.hit.setAttribute("d", d); l.path.setAttribute("d", d);
    }));
    for (const [k, l] of links) if (!liveLinks.has(k)) { l.hit.remove(); l.path.remove(); links.delete(k); }

    if (editor && editor.ring && editor.el.isConnected) {
      const r = editor.el.getBoundingClientRect();
      Object.assign(editor.ring.style, { left: r.left - 2 + "px", top: r.top - 2 + "px", width: r.width + 4 + "px", height: r.height + 4 + "px" });
    }
    requestAnimationFrame(tick);
  }

  // ---------- screenshots (IndexedDB, keyed "<page>:<note id>") ----------
  let dbp = null;
  function db() {
    if (!dbp) dbp = new Promise((res, rej) => {
      const q = indexedDB.open("annotator", 1);
      q.onupgradeneeded = () => q.result.createObjectStore("shots");
      q.onsuccess = () => res(q.result); q.onerror = () => rej(q.error);
    });
    return dbp;
  }
  async function idb(mode, fn) {
    const d = await db();
    return new Promise((res, rej) => {
      const t = d.transaction("shots", mode), q = fn(t.objectStore("shots"));
      t.oncomplete = () => res(q && q.result); t.onerror = () => rej(t.error);
    });
  }
  const shotGet = (page, id) => idb("readonly", (st) => st.get(page + ":" + id)).catch(() => null);
  const shotPut = (id, shot) => idb("readwrite", (st) => st.put(shot, PAGE + ":" + id)).catch(() => {});
  const shotDel = (id) => idb("readwrite", (st) => st.delete(PAGE + ":" + id)).catch(() => {});

  const libs = {};
  function loadLib(url, global) {
    if (window[global]) return Promise.resolve(window[global]);
    return libs[url] || (libs[url] = new Promise((res, rej) => {
      const sc = document.createElement("script");
      sc.src = url; sc.setAttribute("data-annotator-ui", "");
      sc.onload = () => (window[global] ? res(window[global]) : rej(new Error(global + " missing")));
      sc.onerror = () => { delete libs[url]; rej(new Error("Couldn't load " + url)); };
      document.head.appendChild(sc);
    }));
  }

  // Screenshot of the element and its callout box exactly as shown on the page: the page is rendered without the
  // annotator UI, then the real card (rendered on its own), its dashed outline and its arrow are drawn at the same spots.
  const SHOT_V = 5; // bump when screenshots need retaking (v2: fixed shift on scrolled pages, v3: modals mid-animation, v4: modals on scrolled pages came out blank, v5: inline-flex tab bars hid their tabs)
  // html2canvas renders a copy of the page, where CSS animations start over — a modal with a pop-in/slide-in
  // animation would be captured mid-animation (shifted/scaled) while the outline is placed where it really is.
  // Freeze the copy in its final, settled state.
  function markFixed() {
    const out = [], all = document.body.getElementsByTagName("*");
    for (let i = 0; i < all.length; i++) {
      const el = all[i];
      if (isOurs(el) || getComputedStyle(el).position !== "fixed") continue;
      el.setAttribute("data-ann-fx", out.length);
      out.push({ el, r: el.getBoundingClientRect() });
    }
    return out;
  }
  function pinFixed(doc, fixed) {
    const win = doc.defaultView;
    fixed.forEach(({ r }, i) => { // outer ones first (document order), so inner ones can be placed relative to them
      const c = doc.querySelector('[data-ann-fx="' + i + '"]'); if (!c) return;
      const set = (k, v) => c.style.setProperty(k, v, "important");
      set("position", "absolute"); set("right", "auto"); set("bottom", "auto"); set("margin", "0"); set("transform", "none");
      set("width", r.width + "px"); set("height", r.height + "px"); set("box-sizing", "border-box");
      const op = c.offsetParent;
      let ox = 0, oy = 0;
      if (op && op !== doc.body && win.getComputedStyle(op).position !== "static") {
        const orc = op.getBoundingClientRect();
        ox = orc.left + win.scrollX + op.clientLeft - op.scrollLeft;
        oy = orc.top + win.scrollY + op.clientTop - op.scrollTop;
      }
      set("left", r.left + scrollX - ox + "px"); set("top", r.top + scrollY - oy + "px");
    });
  }
  function freezeClone(doc) {
    const st = doc.createElement("style");
    st.textContent = "*,*::before,*::after{animation:none!important;transition:none!important}";
    (doc.head || doc.documentElement).appendChild(st);
    fixInlineFlex(doc);
  }
  // html2canvas paints a display:inline-flex / inline-grid box's own background ON TOP of its children, so a
  // pill tab bar (All / My / Joined Stores, Pro / Sideline Stock) comes out as an empty grey bar with the
  // labels and selected pill hidden. An imperceptible opacity makes it paint the box first, then its children.
  function fixInlineFlex(doc) {
    const win = doc.defaultView, all = doc.body ? doc.body.getElementsByTagName("*") : [];
    for (let i = 0; i < all.length; i++) {
      const el = all[i], cs = win.getComputedStyle(el);
      if ((cs.display === "inline-flex" || cs.display === "inline-grid") && cs.opacity === "1") el.style.setProperty("opacity", "0.999", "important");
    }
  }
  async function capture(el, note) {
    const h2c = await loadLib(H2C_URL, "html2canvas");
    const r = el.getBoundingClientRect();
    if (!r.width || !r.height) return null;
    const vw = document.documentElement.clientWidth, vh = document.documentElement.clientHeight;
    if (r.bottom <= 0 || r.top >= vh || r.right <= 0 || r.left >= vw) return null; // off screen: nothing to capture
    const v = note && cards.get(note.id);
    const cardOn = !!(v && visible && !note.hidden && v.card.style.display !== "none" && v.card.offsetWidth);
    const cr = cardOn ? v.card.getBoundingClientRect() : null;
    let L = r.left - 6, T = r.top - 6, R = r.right + 6, B = r.bottom + 6;
    if (cr) { L = Math.min(L, cr.left); T = Math.min(T, cr.top); R = Math.max(R, cr.right); B = Math.max(B, cr.bottom); }
    const w = Math.min(vw, Math.max(R - L + 80, 760)), h = Math.min(vh, Math.max(B - T + 80, 440));
    const x = Math.max(0, Math.min((L + R) / 2 - w / 2, vw - w));
    const y = Math.max(0, Math.min((T + B) / 2 - h / 2, vh - h));
    const scale = Math.min(2, devicePixelRatio || 1);
    // x/y are document coordinates; scrollX/scrollY:0 stops html2canvas from applying the page scroll a second time
    // (otherwise any scrolled page comes out shifted and the outline/arrow miss the element).
    // Elements pinned to the screen (position:fixed — modals, overlays, toolbars) are drawn by html2canvas as if
    // the page weren't scrolled, while everything else is drawn scrolled — so on a scrolled page a modal ends up
    // far from where it's seen (the screenshot came out blank). Mark them, and in the copy re-place each one at
    // exactly where it's seen on screen, in page coordinates.
    const fixed = markFixed();
    let page;
    try {
      page = await h2c(document.body, {
      onclone: (doc) => {
        freezeClone(doc);
        if (!scrollX && !scrollY) return;
        // Modals often set body{overflow:hidden} (with body only one screen tall) to stop the page scrolling;
        // html2canvas then clips everything below the first screen, so a scrolled page renders blank.
        [doc.documentElement, doc.body].forEach((n) => n.style.setProperty("overflow", "visible", "important"));
        pinFixed(doc, fixed);
      },
      x: x + scrollX, y: y + scrollY, width: w, height: h, scale, scrollX: 0, scrollY: 0,
      useCORS: true, logging: false, backgroundColor: "#ffffff",
      ignoreElements: (n) => n.nodeType === 1 && n.hasAttribute("data-annotator-ui"),
      });
    } finally { fixed.forEach((f) => f.el.removeAttribute("data-ann-fx")); }
    let cardImg = null;
    if (cardOn) {
      v.card.classList.add("ann-snapping"); // hides the hover buttons + resize grip while rendering
      try { cardImg = await h2c(v.card, { scale, useCORS: true, logging: false, backgroundColor: null, onclone: freezeClone }); } catch {}
      v.card.classList.remove("ann-snapping");
    }
    const s = page.width / w, k = Math.min(1, 1400 / page.width), f = s * k;
    const out = document.createElement("canvas");
    out.width = Math.round(page.width * k); out.height = Math.round(page.height * k);
    const ctx = out.getContext("2d");
    ctx.drawImage(page, 0, 0, out.width, out.height);
    const box = { x: (r.left - x) * f, y: (r.top - y) * f, w: r.width * f, h: r.height * f };
    if (!cardImg) return { src: out.toDataURL("image/jpeg", 0.88), w: out.width, h: out.height, box, at: new Date().toISOString() };

    const X = (px) => (px - x) * f, Y = (py) => (py - y) * f;
    // Dashed outline (same as .ann-ring: 3px outside the element, 2px dashed).
    ctx.strokeStyle = ACCENT; ctx.lineWidth = 2 * f; ctx.setLineDash([6 * f, 4 * f]);
    ctx.beginPath();
    (ctx.roundRect ? ctx.roundRect.bind(ctx) : ctx.rect.bind(ctx))(X(r.left - 2), Y(r.top - 2), (r.width + 4) * f, (r.height + 4) * f, 4 * f);
    ctx.stroke(); ctx.setLineDash([]);
    // Arrow (same curve + arrowhead as the on-page SVG).
    const g = arrowGeom(r, cr.left, cr.top, cr.width, cr.height);
    if (Math.hypot(g.ex - g.sx, g.ey - g.sy) >= 6) {
      ctx.lineWidth = 2 * f; ctx.beginPath(); ctx.moveTo(X(g.sx), Y(g.sy)); ctx.quadraticCurveTo(X(g.mx), Y(g.my), X(g.ex), Y(g.ey)); ctx.stroke();
      const a = Math.atan2(g.ey - g.my, g.ex - g.mx), len = 14 * f, half = 7 * f;
      ctx.fillStyle = ACCENT; ctx.beginPath();
      ctx.moveTo(X(g.ex), Y(g.ey));
      ctx.lineTo(X(g.ex) - len * Math.cos(a) + half * Math.sin(a), Y(g.ey) - len * Math.sin(a) - half * Math.cos(a));
      ctx.lineTo(X(g.ex) - len * Math.cos(a) - half * Math.sin(a), Y(g.ey) - len * Math.sin(a) + half * Math.cos(a));
      ctx.closePath(); ctx.fill();
    }
    // The real callout box, with its shadow.
    ctx.save(); ctx.shadowColor = "rgba(20,20,40,.18)"; ctx.shadowBlur = 20 * f; ctx.shadowOffsetY = 6 * f;
    ctx.drawImage(cardImg, X(cr.left), Y(cr.top), cr.width * f, cr.height * f);
    ctx.restore();
    return { src: out.toDataURL("image/jpeg", 0.9), w: out.width, h: out.height, box, card: true, v: SHOT_V, at: new Date().toISOString() };
  }
  const snapTimers = new Map();
  function snapSoon(note, delay = 600) {
    clearTimeout(snapTimers.get(note.id));
    snapTimers.set(note.id, setTimeout(() => { snapTimers.delete(note.id); if (notes.includes(note)) snap(note); }, delay));
  }
  async function snap(note) {
    const el = note && resolve(note);
    if (!el) return null;
    try { const shot = await capture(el, note); if (shot) await shotPut(note.id, shot); return shot; }
    catch (err) { console.warn("[annotator] screenshot failed:", err); return null; }
  }

  async function exportNotes() {
    const images = {};
    for (const n of notes) { const shot = await shotGet(PAGE, n.id); if (shot) images[n.id] = shot; }
    const blob = new Blob([JSON.stringify({ page: PAGE, notes, images }, null, 2)], { type: "application/json" });
    download(blob, PAGE + " - notes.json");
  }
  function download(blob, name) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  }

  // ---------- "All notes" panel: every note on this page, with status + go to / edit / hide / re-attach / delete ----------
  const list = document.createElement("div");
  list.className = "ann-list"; list.setAttribute("data-annotator-ui", "");
  const picked = new Set();
  let editingId = null, reattach = null, listTimer = null;

  function noteStatus(n) {
    const el = resolve(n);
    if (!el) return ["miss", "Not found on this screen"];
    if (n.hidden) return ["hid", "Hidden (pin)"];
    const r = el.getBoundingClientRect();
    if (!(r.width || r.height)) return ["miss", "Element not shown"];
    if (r.bottom < 0 || r.top > innerHeight || r.right < 0 || r.left > innerWidth) return ["off", "Off screen — scroll to it"];
    if (coveredAt(el, r)) return ["cov", "Behind something (e.g. a modal)"];
    return ["ok", "Visible"];
  }
  function toggleList(force) {
    const open = force != null ? force : !list.classList.contains("open");
    list.classList.toggle("open", open);
    clearInterval(listTimer);
    if (open) { renderList(); listTimer = setInterval(() => { if (!editingId && !list.contains(document.activeElement)) renderList(); }, 800); }
  }
  function renderList() {
    const items = ordered(notes);
    const missing = items.filter((x) => noteStatus(x.n)[0] === "miss").length;
    for (const id of [...picked]) if (!notes.some((n) => n.id === id)) picked.delete(id);
    const html = `<div class="lh"><b>All notes on this page (${notes.length})</b><button data-a="close" title="Close">✕</button></div>
      <div class="lf"><label><input type="checkbox" data-a="all" ${picked.size && picked.size === notes.length ? "checked" : ""}> Select all</label>
        ${missing ? `<button data-a="pickmiss" title="Select the notes whose element can't be found on this screen">Select not found (${missing})</button>` : ""}
        <button data-a="delpicked" class="del" ${picked.size ? "" : "disabled"}>Delete selected (${picked.size})</button></div>
      <div class="lb">${items.length ? items.map((x) => {
        const n = x.n, [cls, label] = noteStatus(n), title = plainText(splitBody(n.body).title) || "(no title)";
        const target = n.tag + (n.text ? " · “" + n.text.slice(0, 40) + "”" : n.hint ? " · " + n.hint : "");
        if (editingId === n.id) return `<div class="row" data-id="${n.id}"><div class="mn">
          <div class="tl"><span class="nb">${esc(x.label)}</span><span class="t">Edit note</span></div>
          <input class="ni" placeholder="No. (e.g. FR-1.1)" value="${esc(n.num || "")}"><textarea>${esc(n.body)}</textarea>
          <div class="ac"><button data-a="save" class="pr">Save</button><button data-a="cancel">Cancel</button></div></div></div>`;
        return `<div class="row" data-id="${n.id}"><input type="checkbox" data-a="pick" ${picked.has(n.id) ? "checked" : ""}>
          <div class="mn"><div class="tl"><span class="nb">${esc(x.label)}</span><span class="t" title="${esc(n.body)}">${esc(title)}</span></div>
          <div class="el" title="${esc(target)}">${esc(target)}</div>
          <span class="st ${cls}">${esc(label)}</span>
          <div class="ac"><button data-a="go" ${cls === "miss" ? "disabled title='Its element is not on this screen — open the screen/modal it belongs to, or re-attach it'" : ""}>Go to</button>
            <button data-a="edit">Edit</button>
            <button data-a="vis">${n.hidden ? "Show" : "Hide"}</button>
            <button data-a="reattach" title="Then ⌘/Ctrl-click the element this note should point at">Re-attach</button>
            <button data-a="del" class="del">Delete</button></div></div></div>`;
      }).join("") : '<div class="empty">No notes on this page yet.</div>'}</div>`;
    // The panel refreshes itself every 800ms (statuses change as you scroll / open modals): only redraw when
    // something changed, and keep the list's scroll position so scrolling down it doesn't jump back to the top.
    if (html === list._html) return;
    const lb = list.querySelector(".lb"), top = lb ? lb.scrollTop : 0;
    list.innerHTML = html; list._html = html;
    const nb = list.querySelector(".lb"); if (nb) nb.scrollTop = top;
  }
  list.addEventListener("click", (e) => {
    const a = e.target.closest("[data-a]")?.dataset.a; if (!a) return;
    const row = e.target.closest(".row"), n = row && notes.find((x) => x.id === row.dataset.id);
    if (a === "close") return toggleList(false);
    if (a === "all") { if (e.target.checked) notes.forEach((x) => picked.add(x.id)); else picked.clear(); }
    if (a === "pickmiss") notes.forEach((x) => { if (noteStatus(x)[0] === "miss") picked.add(x.id); });
    if (a === "pick" && n) { if (e.target.checked) picked.add(n.id); else picked.delete(n.id); }
    if (a === "delpicked" && picked.size && confirm("Delete " + picked.size + " selected note" + (picked.size === 1 ? "" : "s") + "?")) { [...picked].forEach(remove); picked.clear(); }
    if (a === "del" && n && confirm("Delete note " + numOf(n, notes.indexOf(n)) + "?")) remove(n.id);
    if (a === "vis" && n) { n.hidden = !n.hidden; visible = true; save(); }
    if (a === "edit" && n) editingId = n.id;
    if (a === "cancel") editingId = null;
    if (a === "save" && n) {
      const body = row.querySelector("textarea").value.trim();
      if (body) { n.body = body; n.num = row.querySelector(".ni").value.trim() || null; save(); snapSoon(n); }
      editingId = null;
    }
    if (a === "go" && n) {
      const el = resolve(n); if (!el) return;
      if (n.hidden) { n.hidden = false; save(); }
      visible = true; updateToolbar();
      el.scrollIntoView({ behavior: "smooth", block: "center" });
      setTimeout(() => { const v = cards.get(n.id); if (v) { v.card.classList.remove("ann-flash"); void v.card.offsetWidth; v.card.classList.add("ann-flash"); } }, 450);
    }
    if (a === "reattach" && n) { setReattach(n.id); return; }
    renderList();
  });

  function setReattach(id) {
    reattach = id;
    banner.style.display = id ? "block" : "none";
    if (id) banner.textContent = "Re-attach note " + numOf(notes.find((n) => n.id === id), notes.findIndex((n) => n.id === id)) + ": " + (/Mac/.test(navigator.platform) ? "⌘" : "Ctrl") + "-click the element it should point at · Esc to cancel";
    else banner.textContent = "Click another note to connect it · click the same pair again to disconnect · Esc to cancel";
  }
  function attachTo(id, el) {
    const n = notes.find((x) => x.id === id); setReattach(null);
    if (!n) return;
    Object.assign(n, { selector: selectorFor(el), tag: el.tagName.toLowerCase(), text: textOf(el), hint: hintOf(el), dx: null, dy: null, hidden: false });
    cache.set(n.id, el); lastScan.delete(n.id);
    visible = true; save(); snapSoon(n, 300);
    if (list.classList.contains("open")) renderList();
  }

  // ---------- report (PDF via print / Word .docx) ----------
  function annotatedPages() {
    const out = [];
    for (let i = 0; i < localStorage.length; i++) {
      const k = localStorage.key(i);
      if (!k || k.indexOf("annotator:") !== 0) continue;
      const page = k.slice(10);
      let list = page === PAGE ? notes : null;
      if (!list) try { list = JSON.parse(localStorage.getItem(k) || "[]"); } catch { list = []; }
      if (Array.isArray(list) && list.length) out.push({ page, notes: list });
    }
    if (!out.some((p) => p.page === PAGE) && notes.length) out.push({ page: PAGE, notes });
    // First page = the one annotated first.
    const first = (p) => Math.min(...p.notes.map((n) => Date.parse(n.created) || Infinity));
    return out.sort((a, b) => first(a) - first(b));
  }

  function showReportMenu() {
    if (pop.style.display === "block") { pop.style.display = "none"; return; }
    const pages = annotatedPages();
    const total = pages.reduce((t, p) => t + p.notes.length, 0);
    if (!total) { alert("There are no notes yet. " + (/Mac/.test(navigator.platform) ? "⌘" : "Ctrl") + "+click an element to add one."); return; }
    if (pages.length < 2) { openReport(pages); return; }
    pop.innerHTML = `<h4>Generate report</h4>
      <label><input type="radio" name="ann-scope" value="page" checked><span>This page<small>${notes.length} note${notes.length === 1 ? "" : "s"} on ${esc(PAGE)}</small></span></label>
      <label><input type="radio" name="ann-scope" value="all"><span>All annotated pages<small>${total} notes on ${pages.length} pages, first page to last</small></span></label>
      <div class="ann-row"><button data-act="close">Cancel</button><button class="ann-primary" data-act="go">Open report</button></div>`;
    pop.style.display = "block";
    pop.onclick = (e) => {
      const act = e.target.closest("button")?.dataset.act;
      if (act === "close") pop.style.display = "none";
      if (act === "go") {
        const all = pop.querySelector('input[value="all"]').checked;
        pop.style.display = "none";
        openReport(all ? pages : pages.filter((p) => p.page === PAGE));
      }
    };
  }

  const esc = (t) => String(t == null ? "" : t).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c]);

  // ---------- links → chips (like smart chips in Google Docs) ----------
  // A pasted URL becomes a chip with an icon and a readable title; "[My title](https://…)" sets the title yourself.
  const LINK_RE = /\[([^\]\n]+)\]\((https?:\/\/[^\s)]+)\)|(https?:\/\/[^\s<>"']+)/g;
  const KINDS = {
    pdf: ["PDF", "#e1251b"], figma: ["F", "#a259ff"], confluence: ["C", "#1868db"], jira: ["J", "#2684ff"], monday: ["M", "#ff3d57"],
    gdoc: ["G", "#1a73e8"], video: ["▶", "#111111"], file: ["DOC", "#2b579a"], link: ["↗", "#6b6b72"],
  };
  function linkInfo(url, label) {
    let u; try { u = new URL(url); } catch { return { url, title: label || url, kind: "link" }; }
    const host = u.hostname.replace(/^www\./, "");
    let path = u.pathname; try { path = decodeURIComponent(path); } catch {}
    const file = path.split("/").filter(Boolean).pop() || "";
    let kind = "link", title = "", m;
    if (/\.pdf$/i.test(path)) { kind = "pdf"; title = file; }
    else if (/figma\.com$/.test(host)) { kind = "figma"; m = path.match(/\/(?:file|design|proto|board)\/[^/]+\/([^/]+)/); title = m ? m[1].replace(/-/g, " ") : "Figma design"; }
    else if (/atlassian\.net$/.test(host) && /\/wiki\//.test(path)) { kind = "confluence"; m = path.match(/\/pages\/\d+\/([^/]+)/); title = m ? m[1].replace(/\+/g, " ") : "Confluence page"; }
    else if (/atlassian\.net$|jira/.test(host)) { kind = "jira"; m = (path + u.search).match(/[A-Z][A-Z0-9]+-\d+/); title = m ? m[0] : "JIRA ticket"; }
    else if (/monday\.com$/.test(host)) { kind = "monday"; m = path.match(/\/pulses\/(\d+)/); title = m ? "Monday item " + m[1] : "Monday board"; }
    else if (host === "docs.google.com") { kind = "gdoc"; title = { document: "Google Doc", spreadsheets: "Google Sheet", presentation: "Google Slides" }[path.split("/")[1]] || "Google Drive file"; }
    else if (host === "drive.google.com") { kind = "gdoc"; title = "Google Drive file"; }
    else if (/\.(mp4|mov|webm)$/i.test(path)) { kind = "video"; title = file; }
    else if (/(^|\.)(loom\.com|youtube\.com|youtu\.be|vimeo\.com)$/.test(host)) { kind = "video"; title = "Video · " + host; }
    else if (/\.(docx?|xlsx?|pptx?|csv)$/i.test(path)) { kind = "file"; title = file; }
    else { title = host + (path.length > 1 ? path.replace(/\/$/, "") : ""); if (title.length > 60) title = host + " › " + file; }
    return { url, title: label || title || url, kind };
  }
  // Splits text into plain strings and link objects.
  function linkParts(text) {
    const out = [], s = String(text == null ? "" : text);
    let last = 0, m;
    LINK_RE.lastIndex = 0;
    while ((m = LINK_RE.exec(s))) {
      let url = m[2] || m[3], end = m.index + m[0].length;
      if (m[3]) {
        const hlEnd = url.indexOf("=="); if (hlEnd > 8) { end -= url.length - hlEnd; url = url.slice(0, hlEnd); }
        const trail = url.match(/[.,;:!?)\]]+$/); if (trail) { url = url.slice(0, -trail[0].length); end -= trail[0].length; }
      }
      if (m.index > last) out.push(s.slice(last, m.index));
      const label = m[1] ? m[1].trim() : "";
      out.push(Object.assign(linkInfo(url, label), { label, raw: s.slice(m.index, end), start: m.index, end }));
      last = end; LINK_RE.lastIndex = end;
    }
    if (last < s.length) out.push(s.slice(last));
    return out;
  }
  const chipHTML = (l, i) => { const [ic, c] = KINDS[l.kind] || KINDS.link; return `<a class="ann-chip" data-i="${i}" href="${esc(l.url)}" target="_blank" rel="noopener" title="${esc(l.url)}"><span class="ann-ic" style="background:${c}">${ic}</span>${esc(l.title)}</a>`; };
  // ---------- highlights: ==text== (like a highlighter pen; used to flag questions to check) ----------
  const HL_RE = /==([^\n]+?)==/g;
  function hlParts(text) {
    const out = [], s = String(text == null ? "" : text);
    let last = 0, m;
    HL_RE.lastIndex = 0;
    while ((m = HL_RE.exec(s))) {
      if (m.index > last) out.push({ text: s.slice(last, m.index), hl: false });
      out.push({ text: m[1], hl: true });
      last = m.index + m[0].length;
    }
    if (last < s.length) out.push({ text: s.slice(last), hl: false });
    return out;
  }
  const hasHL = (text) => { HL_RE.lastIndex = 0; return HL_RE.test(String(text || "")); };
  // Text → HTML with yellow highlights and link chips (chip data-i counts links across the whole text).
  const linkifyHTML = (text) => {
    let i = 0;
    return hlParts(text).map((seg) => {
      const inner = linkParts(seg.text).map((p) => (typeof p === "string" ? esc(p) : chipHTML(p, i++))).join("");
      return seg.hl ? `<mark class="ann-mark">${inner}</mark>` : inner;
    }).join("");
  };
  const HL_CSS = `mark.ann-mark{background:#fde047;color:inherit;padding:0 2px;border-radius:2px;-webkit-print-color-adjust:exact;print-color-adjust:exact}`;

  // ---------- "Insert link" dialog: Title + URL (like Google Docs) ----------
  // Links are stored in the text as [Title](https://…), or just the URL when no title is given.
  const fmtLink = (title, url) => { const tt = String(title || "").replace(/[\[\]\n]/g, " ").trim(); return tt ? "[" + tt + "](" + url + ")" : url; };
  const linkList = (text) => linkParts(text).filter((p) => typeof p !== "string");
  // Replaces the link under the selection (or the selected text) with a new link; returns the new text.
  function applyLink(text, s, e, hit, title, url) {
    const [a, b] = hit ? [hit.start, hit.end] : [s, e];
    return text.slice(0, a) + fmtLink(title, url) + text.slice(b);
  }
  const linkAt = (text, s, e) => linkList(text).find((p) => p.start <= s && e <= p.end) || null;

  const LINKFORM_CSS = `.ann-linkform,.ann-chippop{position:fixed;z-index:2147483646;background:#fff;border:1px solid #dadce0;border-radius:10px;box-shadow:0 8px 28px rgba(20,20,40,.22);font:13px/1.4 Inter,system-ui,-apple-system,sans-serif;color:#0e0e10}
  .ann-linkform{width:340px;padding:12px}
  .ann-linkform h5{margin:0 0 8px;font-size:13px}
  .ann-linkform label{display:block;font-size:11px;font-weight:600;color:#6b6b72;margin-bottom:8px}
  .ann-linkform input{display:block;width:100%;box-sizing:border-box;margin-top:3px;border:1px solid #d8d8de;border-radius:6px;padding:6px 8px;font:inherit;font-size:13px;color:#0e0e10;outline:none}
  .ann-linkform input:focus{border-color:${ACCENT};box-shadow:0 0 0 3px rgba(37,99,235,.15)}
  .ann-linkform .r,.ann-chippop .r{display:flex;gap:6px;align-items:center}
  .ann-linkform .err{flex:1;color:#e1251b;font-size:11px}
  .ann-linkform button,.ann-chippop button{border:1px solid #d8d8de;background:#fff;border-radius:6px;font:inherit;font-size:12px;cursor:pointer;padding:4px 10px;color:#0e0e10}
  .ann-linkform button.p{background:${ACCENT};border-color:${ACCENT};color:#fff}
  .ann-chippop{padding:8px;max-width:420px}
  .ann-chippop .u{flex:1;min-width:0;color:#1155cc;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;padding:0 6px;text-decoration:none}
  .ann-chippop .u:hover{text-decoration:underline}
  .ann-chippop .del{color:#e1251b}`;
  function ensureLinkCSS(doc) {
    if (doc.getElementById("ann-linkform-css")) return;
    const st = doc.createElement("style"); st.id = "ann-linkform-css"; st.textContent = LINKFORM_CSS; doc.head.appendChild(st);
  }
  function closeLinkUI(doc) { doc.querySelectorAll(".ann-linkform,.ann-chippop").forEach((n) => n.remove()); }
  const placeAt = (doc, box, x, y) => {
    const win = doc.defaultView, W = box.offsetWidth, H = box.offsetHeight;
    box.style.left = Math.max(8, Math.min(x, win.innerWidth - W - 8)) + "px";
    box.style.top = Math.max(8, y + H > win.innerHeight - 8 ? y - H - 36 : y) + "px";
  };
  // Opens the Title + URL dialog at (x, y) in `doc`; onApply(title, url) runs with a valid http(s) URL.
  function linkForm(doc, x, y, init, onApply, onCancel) {
    ensureLinkCSS(doc); closeLinkUI(doc);
    const f = doc.createElement("div");
    f.className = "ann-linkform"; f.setAttribute("data-annotator-ui", "");
    f.innerHTML = `<h5>${init.url ? "Edit link" : "Insert link"}</h5>
      <label>Title<input name="t" placeholder="e.g. Team Store Creation Flow"></label>
      <label>Link<input name="u" placeholder="Paste a URL — https://…"></label>
      <div class="r"><span class="err"></span><button data-a="c">Cancel</button><button data-a="s" class="p">Apply</button></div>`;
    const ti = f.querySelector('[name="t"]'), ui = f.querySelector('[name="u"]'), err = f.querySelector(".err");
    ti.value = init.title || ""; ui.value = init.url || "";
    doc.body.appendChild(f); placeAt(doc, f, x, y);
    const close = (cancelled) => { f.remove(); if (cancelled && onCancel) onCancel(); };
    const apply = () => {
      let url = ui.value.trim();
      if (url && !/^https?:\/\//i.test(url) && /^[\w-]+(\.[\w-]+)+/.test(url)) url = "https://" + url; // "figma.com/…" → https://figma.com/…
      if (!/^https?:\/\/\S+$/i.test(url)) { err.textContent = "Enter a valid link (https://…)"; ui.focus(); return; }
      close(false); onApply(ti.value.trim(), url);
    };
    f.addEventListener("click", (e) => { const a = e.target.closest("button")?.dataset.a; if (a === "s") apply(); if (a === "c") close(true); });
    f.addEventListener("keydown", (e) => {
      if (e.key === "Enter") { e.preventDefault(); apply(); }
      if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); close(true); }
    });
    setTimeout(() => (init.url && !init.title ? ti : init.url ? ti : ui).focus(), 0);
    return f;
  }
  const plainText = (text) => linkParts(String(text == null ? "" : text).replace(HL_RE, "$1")).map((p) => (typeof p === "string" ? p : p.title)).join("");
  const CHIP_CSS = `.ann-chip{display:inline-flex;align-items:center;gap:5px;max-width:100%;vertical-align:baseline;padding:1px 8px 1px 3px;margin:1px 0;border:1px solid #dadce0;border-radius:12px;background:#f1f3f4;color:#1f1f1f;font-size:.92em;font-weight:500;line-height:1.5;text-decoration:none;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;cursor:pointer}
  .ann-chip:hover{background:#e8eaed;border-color:#c4c7cc}
  .ann-ic{flex:none;display:inline-flex;align-items:center;justify-content:center;min-width:16px;height:16px;padding:0 3px;border-radius:4px;color:#fff;font-size:8px;font-weight:800;letter-spacing:.02em}
  @media print{.ann-chip{border-color:#c8c8c8;-webkit-print-color-adjust:exact;print-color-adjust:exact}}`;

  // ---------- FRD (Functional Requirements Document) ----------
  // Header fields, versions, section titles, phased-in rows and references are edited in the
  // report window itself and remembered in localStorage (shared by every page on the site).
  const META_KEY = "annotator-frd";
  const META_FIELDS = { project: "", type: "Change Request", stakeholder: "", uiux: "", ba: "", pm: "", dev: "", sqa: "", other: "", priority: "", brand: "" };
  const REFS = ["Final Design Figma Link", "Monday ticket", "Hand-off Video", "Technical Specs", "JIRA Ticket"];
  const loadMeta = () => { try { return JSON.parse(localStorage.getItem(META_KEY) || "{}"); } catch { return {}; } };
  const saveMeta = (m) => { try { localStorage.setItem(META_KEY, JSON.stringify(m)); } catch {} };

  function frdDefaults(groups, date) {
    // Project name defaults to the words every page name starts with ("Team Stores Workspace" + "Team Stores Landing" → "Team Stores").
    const words = groups.map((g) => g.page.split(/\s+/));
    const common = [];
    for (let i = 0; words.every((w) => w[i] && w[i] === words[0][i]); i++) common.push(words[0][i]);
    const d = Object.assign({}, META_FIELDS, {
      project: (groups.length > 1 && common.length ? common.join(" ") : groups[0].page).replace(/\s*[-–—]\s*$/, ""),
      "ver:title": "v1.0.0 - Minor Release", "ver:0:0": "v1.0.0", "ver:0:1": "n/a", "ver:0:2": "INITIAL RELEASE", "ver:0:3": date,
    });
    groups.forEach((g, gi) => { d["sec:" + g.page] = gi + 1 + ". " + g.page; });
    return d;
  }

  // First line of a note = requirement title; remaining lines = description ("-", "•", "*" lines become bullets).
  function splitBody(body) {
    const lines = String(body || "").split("\n").map((l) => l.trim()).filter(Boolean);
    return {
      title: lines[0] || "",
      lines: lines.slice(1).map((l) => { const m = l.match(/^[-•*]\s*(.*)$/); return m ? { bullet: true, text: m[1] } : { bullet: false, text: l }; }),
    };
  }

  // Draws the numbered callout (label + title, arrow, dashed outline) onto the raw screenshot, like the on-page callout.
  function compose(shot, label, title) {
    if (!shot || !shot.box || shot.card) return Promise.resolve(shot); // .card = the real callout is already in the image
    return new Promise((res) => {
      const img = new Image();
      img.onerror = () => res(shot);
      img.onload = () => {
        const c = document.createElement("canvas"); c.width = shot.w; c.height = shot.h;
        const ctx = c.getContext("2d"), b = shot.box, u = shot.w / 900;
        ctx.drawImage(img, 0, 0);
        ctx.strokeStyle = ACCENT; ctx.lineWidth = 2 * u; ctx.setLineDash([6 * u, 4 * u]);
        ctx.strokeRect(b.x - 4 * u, b.y - 4 * u, b.w + 8 * u, b.h + 8 * u); ctx.setLineDash([]);

        const t = title.length > 70 ? title.slice(0, 67) + "…" : title;
        ctx.font = `700 ${11 * u}px Inter, system-ui, sans-serif`;
        const pw = ctx.measureText(label).width + 14 * u;
        ctx.font = `500 ${12.5 * u}px Inter, system-ui, sans-serif`;
        const tw = Math.min(ctx.measureText(t).width, c.width - pw - 60 * u);
        const bw = pw + tw + 28 * u, bh = 30 * u, gap = 30 * u;
        let bx = Math.min(Math.max(8 * u, b.x), c.width - bw - 8 * u), by = b.y - bh - gap, above = true;
        if (by < 8 * u) { by = b.y + b.h + gap; above = false; }
        if (by + bh > c.height - 8 * u) { by = Math.max(8 * u, b.y + 8 * u); above = null; } // no room: overlay inside

        if (above !== null) {
          const sx = bx + Math.min(bw / 2, 60 * u), sy = above ? by + bh : by;
          const ex = Math.max(b.x + 6 * u, Math.min(sx + 30 * u, b.x + b.w - 6 * u)), ey = above ? b.y - 5 * u : b.y + b.h + 5 * u;
          ctx.strokeStyle = ACCENT; ctx.lineWidth = 2 * u;
          ctx.beginPath(); ctx.moveTo(sx, sy); ctx.quadraticCurveTo(sx, ey, ex, ey); ctx.stroke();
          const dir = ex >= sx ? 1 : -1;
          ctx.fillStyle = ACCENT; ctx.beginPath();
          ctx.moveTo(ex, ey); ctx.lineTo(ex - dir * 8 * u, ey - 4.5 * u); ctx.lineTo(ex - dir * 8 * u, ey + 4.5 * u); ctx.closePath(); ctx.fill();
        }
        const rr = (x, y, w, h, r) => { ctx.beginPath(); ctx.roundRect ? ctx.roundRect(x, y, w, h, r) : ctx.rect(x, y, w, h); };
        ctx.save(); ctx.shadowColor = "rgba(20,20,40,.22)"; ctx.shadowBlur = 12 * u; ctx.shadowOffsetY = 3 * u;
        rr(bx, by, bw, bh, 7 * u); ctx.fillStyle = "#fff"; ctx.fill(); ctx.restore();
        ctx.strokeStyle = ACCENT; ctx.lineWidth = 1.5 * u; rr(bx, by, bw, bh, 7 * u); ctx.stroke();
        rr(bx + 8 * u, by + 7 * u, pw, 16 * u, 8 * u); ctx.fillStyle = ACCENT; ctx.fill();
        ctx.fillStyle = "#fff"; ctx.textBaseline = "middle"; ctx.font = `700 ${11 * u}px Inter, system-ui, sans-serif`;
        ctx.fillText(label, bx + 15 * u, by + 15.5 * u);
        ctx.fillStyle = "#0e0e10"; ctx.font = `500 ${12.5 * u}px Inter, system-ui, sans-serif`;
        ctx.fillText(t, bx + pw + 16 * u, by + 15.5 * u, tw);
        res({ src: c.toDataURL("image/jpeg", 0.88), w: c.width, h: c.height });
      };
      img.src = shot.src;
    });
  }

  async function openReport(pages) {
    // Open the window synchronously (inside the click) so pop-up blockers allow it.
    const w = window.open("", "_blank");
    if (!w) { alert("Allow pop-ups for this page to open the report."); return; }
    w.document.write('<title>Preparing requirements document…</title><p id="s" style="font:14px system-ui;padding:40px;color:#444">Preparing requirements document…</p>');
    const status = (t) => { const el = w.document.getElementById("s"); if (el) el.textContent = t; };

    const groups = [];
    let missing = 0;
    for (const [gi, pg] of pages.entries()) {
      const items = ordered(pg.notes);
      // Custom numbers (e.g. "FR-1.1") are used as-is; auto-numbered notes become FR-<section>.<n>.
      const frById = new Map(items.map((x, i) => [x.n.id, x.n.num ? x.label : "FR-" + (gi + 1) + "." + (i + 1)]));
      const rows = [];
      for (const [idx, x] of items.entries()) {
        status("Collecting screenshots… " + (idx + 1) + " / " + items.length + (pages.length > 1 ? " (" + pg.page + ")" : ""));
        let shot = await shotGet(pg.page, x.n.id);
        // Only this page's elements can be captured now: take missing screenshots, and retake older ones that
        // don't show the real callout box yet or came from a buggy version (if that element is on screen).
        if (pg.page === PAGE && (!shot || !shot.card || shot.v !== SHOT_V)) shot = (await snap(x.n)) || shot;
        if (!shot) missing++;
        const fr = frById.get(x.n.id), parts = splitBody(x.n.body);
        parts.hls = [parts.title].concat(parts.lines.map((l) => l.text)).flatMap((txt) => hlParts(txt).filter((sg) => sg.hl).map((sg) => plainText(sg.text)));
        parts.plainTitle = plainText(parts.title);
        rows.push({
          fr, title: parts.title, lines: parts.lines, hls: parts.hls,
          target: x.n.tag + (x.n.text ? " · “" + x.n.text + "”" : ""),
          next: (x.n.links || []).map((id) => frById.get(id)).filter(Boolean),
          shot: shot && (await compose(shot, fr, parts.plainTitle)),
        });
      }
      groups.push({ page: pg.page, rows });
    }
    const date = new Date().toLocaleDateString(undefined, { year: "numeric", month: "long", day: "numeric" });
    const defs = frdDefaults(groups, date);
    const meta = loadMeta();
    const val = (k) => (meta[k] != null ? meta[k] : defs[k] != null ? defs[k] : "");
    const count = groups.reduce((t, g) => t + g.rows.length, 0);

    w.document.open();
    w.document.write(frdHTML({ val, groups, count, missing }));
    w.document.close();
    const d = w.document;
    d.title = plainText(val("project")) + " — Requirements Document";
    // Editable cells show links as chips; while a cell is being edited it shows the raw text (URLs / [Title](url)).
    ensureLinkCSS(d);
    const setCell = (el, text) => {
      const k = el.dataset.k, m = loadMeta(); m[k] = text; saveMeta(m); meta[k] = text;
      if (d.activeElement !== el) el.innerHTML = linkifyHTML(text); else el.textContent = text;
    };
    // Caret/selection as character offsets inside a (plain-text, being edited) cell.
    const caretIn = (el) => {
      const sel = w.getSelection(); if (!sel.rangeCount || !el.contains(sel.anchorNode)) { const n = el.innerText.length; return [n, n]; }
      const r = sel.getRangeAt(0), pre = d.createRange(); pre.selectNodeContents(el); pre.setEnd(r.startContainer, r.startOffset);
      const s = pre.toString().length; return [s, s + r.toString().length];
    };
    const cellLink = (el, x, y) => {
      const text = el.innerText.replace(/\n$/, ""), [s, en] = caretIn(el), hit = linkAt(text, s, en);
      linkForm(d, x, y, hit ? { title: hit.label, url: hit.url } : { title: text.slice(s, en).trim(), url: "" },
        (title, url) => { el.blur(); setCell(el, applyLink(text, s, en, hit, title, url)); });
    };
    d.getElementById("lnk").addEventListener("mousedown", (e) => e.preventDefault()); // keep the cell focused
    d.getElementById("lnk").onclick = (e) => {
      const el = d.activeElement && d.activeElement.closest && d.activeElement.closest("[data-k]");
      const r = e.currentTarget.getBoundingClientRect();
      if (!el) { w.alert("Click into a white cell first (e.g. References → Location), then click “Insert link”."); return; }
      const cr = el.getBoundingClientRect(); cellLink(el, cr.left, cr.bottom + 6); void r;
    };
    d.addEventListener("mousedown", (e) => { if (!e.target.closest(".ann-linkform,.ann-chippop,a.ann-chip")) d.querySelectorAll(".ann-chippop").forEach((n) => n.remove()); });
    d.querySelectorAll("[data-k]").forEach((el) => {
      const k = el.dataset.k;
      // Clicking a chip (cell not being edited) shows a small popup: Open · Edit · Remove.
      el.addEventListener("mousedown", (e) => {
        const a = e.target.closest("a.ann-chip");
        if (!a || d.activeElement === el) return;
        e.preventDefault();
        closeLinkUI(d);
        const i = +a.dataset.i, link = linkList(val(k))[i]; if (!link) return;
        const pop = d.createElement("div"); pop.className = "ann-chippop";
        pop.innerHTML = `<div class="r"><a class="u" href="${esc(link.url)}" target="_blank" rel="noopener" title="${esc(link.url)}">${esc(link.url)}</a>
          <button data-a="open">Open</button><button data-a="edit">Edit</button><button data-a="del" class="del">Remove</button></div>`;
        d.body.appendChild(pop);
        const ar = a.getBoundingClientRect(); placeAt(d, pop, ar.left, ar.bottom + 6);
        pop.addEventListener("click", (ev) => {
          const act = ev.target.closest("button")?.dataset.a; if (!act) return;
          pop.remove();
          const text = val(k);
          if (act === "open") w.open(link.url, "_blank", "noopener");
          if (act === "del") setCell(el, (text.slice(0, link.start) + text.slice(link.end)).replace(/[ \t]{2,}/g, " ").trim());
          if (act === "edit") linkForm(d, ar.left, ar.bottom + 6, { title: link.label, url: link.url }, (title, url) => setCell(el, applyLink(text, 0, 0, link, title, url)));
        });
        a.addEventListener("click", (ev) => ev.preventDefault(), { once: true });
      });
      el.addEventListener("keydown", (e) => {
        if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === "k") { e.preventDefault(); const cr = el.getBoundingClientRect(); cellLink(el, cr.left, cr.bottom + 6); }
      });
      el.addEventListener("focus", () => { if (el.querySelector("a")) el.textContent = val(k); });
      el.addEventListener("blur", () => { el.innerHTML = linkifyHTML(val(k)); });
      el.addEventListener("paste", (e) => { e.preventDefault(); d.execCommand("insertText", false, e.clipboardData.getData("text/plain")); });
      el.addEventListener("input", () => {
        const m = loadMeta(); m[k] = el.innerText.trim(); saveMeta(m); meta[k] = m[k];
        if (k === "project") d.title = plainText(val("project")) + " — Requirements Document";
      });
    });
    // Links inside the (non-editable) requirement rows open normally; target=_blank already set.
    d.getElementById("pdf").onclick = () => w.print();
    // Highlight navigation: the "To check" list and the toolbar button jump to each highlight in the FR rows.
    const marks = () => Array.from(d.querySelectorAll(".req mark.ann-mark"));
    let hlAt = -1;
    const goHL = (i) => {
      const ms = marks(); if (!ms.length) return;
      hlAt = ((i % ms.length) + ms.length) % ms.length;
      const m = ms[hlAt]; m.scrollIntoView({ behavior: "smooth", block: "center" });
      m.classList.remove("flash"); void m.offsetWidth; m.classList.add("flash");
    };
    if (d.getElementById("hlnav")) {
      d.getElementById("hlnav").onclick = () => goHL(hlAt + 1);
      d.querySelectorAll(".tocheck li").forEach((li, i) => { li.onclick = () => goHL(i); });
      d.getElementById("hlexp").onchange = (e) => d.body.classList.toggle("nohl", !e.target.checked);
    }
    d.getElementById("docx").onclick = async (e) => {
      const btn = e.currentTarget; btn.disabled = true; btn.textContent = "Building…";
      const hlOn = !d.getElementById("hlexp") || d.getElementById("hlexp").checked;
      try { download(await buildDocx({ val, groups, hlOn }), (plainText(val("project")) || "Team Store").replace(/[\\/:*?"<>|]/g, "-") + " - Requirements Document.docx"); }
      catch (err) { w.alert("Couldn't build the Word file: " + err.message); }
      btn.disabled = false; btn.textContent = "Download Word (.docx)";
    };
  }

  function descHTML(r) {
    let h = "", ul = [];
    const flush = () => { if (ul.length) { h += "<ul>" + ul.join("") + "</ul>"; ul = []; } };
    r.lines.forEach((l) => { if (l.bullet) ul.push("<li>" + linkifyHTML(l.text) + "</li>"); else { flush(); h += "<p>" + linkifyHTML(l.text) + "</p>"; } });
    flush();
    if (r.next.length) h += `<p class="nx">Next: ${esc(r.next.join(", "))}</p>`;
    return h || '<p class="mut">—</p>';
  }

  function frdHTML({ val, groups, count, missing }) {
    const hlCount = groups.reduce((t, g) => t + g.rows.reduce((u, r) => u + r.hls.length, 0), 0);
    const ed = (k, attrs) => `<td class="ed" contenteditable="true" data-k="${esc(k)}"${attrs || ""}>${linkifyHTML(val(k))}</td>`;
    const lb = (t, attrs) => `<td class="lb"${attrs || ""}>${esc(t)}</td>`;
    const verRows = [0, 1, 2].map((r) => `<tr>${[0, 1, 2, 3].map((c) => ed("ver:" + r + ":" + c, ' class="c"')).join("")}</tr>`).join("");
    const phRows = [0, 1].map((r) => `<tr>${[0, 1, 2, 3].map((c) => ed("ph:" + r + ":" + c)).join("")}</tr>`).join("");
    const fr = groups.map((g) => `<tbody class="sec"><tr><td class="sh ed" colspan="3" contenteditable="true" data-k="${esc("sec:" + g.page)}">${linkifyHTML(val("sec:" + g.page))}</td></tr>` +
      g.rows.map((r) => `
        <tr class="shot"><td colspan="3">${r.shot ? `<img src="${r.shot.src}" alt="${esc(r.fr)}">` : `<div class="noimg">No screenshot yet — open this screen and save the note again to capture one.</div>`}</td></tr>
        <tr class="req"><td class="frno">${esc(r.fr)}</td><td class="ttl"><b>${linkifyHTML(r.title)}</b><div class="el">${esc(r.target)}</div></td><td class="desc">${descHTML(r)}</td></tr>`).join("") + "</tbody>").join("");
    return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Requirements Document</title>
<style>
  @page{size:A4;margin:14mm}
  *{box-sizing:border-box}
  body{margin:0;background:#eef0f3;color:#111;font:13px/1.45 Calibri,Carlito,Arial,sans-serif}
  .bar{position:sticky;top:0;z-index:2;display:flex;gap:8px;align-items:center;padding:10px 20px;background:#fff;border-bottom:1px solid #e5e5e9;font-family:Inter,system-ui,sans-serif}
  .bar .sp{flex:1;color:#6b6b72;font-size:12px}
  .bar button.hl{background:#fde047;border-color:#e8c40c}
  .bar button{font:inherit;font-size:13px;border:1px solid #d8d8de;background:#fff;border-radius:6px;padding:6px 12px;cursor:pointer}
  .bar button.p{background:${ACCENT};border-color:${ACCENT};color:#fff}
  main{max-width:880px;margin:24px auto;background:#fff;padding:36px 40px;box-shadow:0 2px 12px rgba(20,20,40,.08)}
  h1{font-size:14px;text-align:center;margin:0 0 14px;letter-spacing:.02em}
  table{width:100%;border-collapse:collapse;margin:0 0 26px;table-layout:fixed}
  td{border:1px solid #bfbfbf;padding:5px 8px;vertical-align:top;word-wrap:break-word}
  .lb{background:#d9d9d9;font-weight:700}
  .th td{background:#ececec;font-weight:700;text-align:center}
  .c{text-align:center}
  .ed{min-height:24px;outline:none}
  .ed:hover{background:#fffbe6}
  .ed:focus{background:#fff8d6;box-shadow:inset 0 0 0 2px #f2c94c}
  .ed:empty::before{content:"\\00a0"}
  .sub{font-weight:400;margin-top:4px}
  .sh{background:#d9d9d9;font-weight:700}
  .shot td{padding:10px;text-align:center}
  .shot img{max-width:100%;border:1px solid #d0d0d0}
  .noimg{padding:18px;border:1px dashed #c8c8c8;color:#8a8a8a;font-size:12px}
  .frno{font-weight:700;white-space:nowrap}
  .ttl .el{color:#7a7a7a;font-size:11px;margin-top:3px}
  .desc p{margin:0 0 4px}.desc ul{margin:0 0 4px;padding-left:18px}
  .nx{color:#555;font-size:12px}.mut{color:#aaa}
  ${HL_CSS}
  mark.ann-mark.flash{animation:annFlash 1.4s ease 2}
  @keyframes annFlash{0%,100%{box-shadow:0 0 0 0 rgba(245,158,11,0)}50%{box-shadow:0 0 0 5px rgba(245,158,11,.55)}}
  .tocheck{border:1px solid #f5d565;background:#fffbea;border-radius:8px;padding:12px 16px;margin:0 0 22px;font-family:Inter,system-ui,sans-serif}
  .tocheck h4{margin:0 0 6px;font-size:13px}
  .tocheck ol{margin:0;padding-left:20px}
  .tocheck li{margin:3px 0;cursor:pointer}
  .tocheck li:hover span{text-decoration:underline}
  .tocheck b{font-size:12px;margin-right:6px}
  .bar label{font-size:12px;color:#444;display:flex;gap:5px;align-items:center;cursor:pointer}
  body.nohl mark.ann-mark{background:none;padding:0}
  ${CHIP_CSS}
  @media print{
    body{background:#fff}
    .bar{display:none}
    main{margin:0;padding:0;box-shadow:none;max-width:none}
    .ed:hover,.ed:focus{background:none;box-shadow:none}
    .tocheck{display:none}
    .lb,.sh,.th td{-webkit-print-color-adjust:exact;print-color-adjust:exact}
    tr{break-inside:avoid}
    .shot{break-after:avoid}
  }
</style></head><body>
<div class="bar"><span class="sp">${count} functional requirement${count === 1 ? "" : "s"}${missing ? " · " + missing + " without screenshot" : ""} · Click any white cell to edit it — edits are remembered · 🔗 Insert link / ⌘K adds a titled link · For PDF, choose “Save as PDF” in the print dialog</span>
  ${hlCount ? `<button id="hlnav" class="hl" title="Jump to the next highlighted question">🖍 ${hlCount} to check · Next ↓</button><label title="Untick to print / export without the yellow highlights"><input type="checkbox" id="hlexp" checked> Highlights in export</label>` : ""}
  <button id="lnk" title="Insert a link with a title into the cell you're editing (⌘K)">🔗 Insert link</button><button id="docx">Download Word (.docx)</button><button id="pdf" class="p">Save as PDF</button></div>
<main>
  ${hlCount ? `<div class="tocheck"><h4>🖍 To check — ${hlCount} highlighted item${hlCount === 1 ? "" : "s"}</h4><ol>${groups.flatMap((g) => g.rows.flatMap((r) => r.hls.map((h) => `<li><b>${esc(r.fr)}</b><span>${esc(h)}</span></li>`))).join("")}</ol></div>` : ""}
  <h1>REQUIREMENTS DOCUMENT</h1>
  <table><colgroup><col style="width:34%"><col style="width:33%"><col style="width:33%"></colgroup>
    <tr>${lb("PROJECT NAME:", ' colspan="2"')}${lb("Type")}</tr>
    <tr>${ed("project", ' colspan="2" style="font-weight:700"')}${ed("type")}</tr>
    <tr>${lb("STAKEHOLDER")}${lb("US UI/UX:", ' colspan="2"')}</tr>
    <tr>${ed("stakeholder")}${ed("uiux", ' colspan="2"')}</tr>
    <tr>${lb("BA AUTHOR:", ' colspan="3"')}</tr><tr>${ed("ba", ' colspan="3"')}</tr>
    <tr>${lb("PROJECT MANAGER:", ' colspan="3"')}</tr><tr>${ed("pm", ' colspan="3"')}</tr>
    <tr>${lb("DEVELOPMENT TEAM: (INCLUDING UI/UX)", ' colspan="3"')}</tr><tr>${ed("dev", ' colspan="3"')}</tr>
    <tr>${lb("SQA TEAM:", ' colspan="3"')}</tr><tr>${ed("sqa", ' colspan="3"')}</tr>
    <tr>${lb("OTHER TEAMS:", ' colspan="3"')}</tr><tr>${ed("other", ' colspan="3"')}</tr>
    <tr>${lb("PRIORITY:")}${lb("BRAND:", ' colspan="2"')}</tr>
    <tr>${ed("priority")}${ed("brand", ' colspan="2"')}</tr>
  </table>
  <table>
    <tr><td class="lb" colspan="4">VERSIONS<div class="sub ed" contenteditable="true" data-k="ver:title">${linkifyHTML(val("ver:title"))}</div></td></tr>
    <tr class="th"><td>VERSION NO.</td><td>BR/FR NO.</td><td>REVISION REMARKS</td><td>LAST UPDATED DATE</td></tr>
    ${verRows}
  </table>
  <table><colgroup><col style="width:13%"><col style="width:27%"><col style="width:60%"></colgroup>
    <tr><td class="lb" colspan="3">FUNCTIONAL REQUIREMENTS (FR)</td></tr>
    ${fr}
  </table>
  <table>
    <tr><td class="lb" colspan="4">PHASED IN REQUIREMENTS</td></tr>
    <tr class="th"><td>FR NO.</td><td>REQUIREMENT</td><td>TARGET PHASE</td><td>REMARKS</td></tr>
    ${phRows}
  </table>
  <table><colgroup><col style="width:30%"><col style="width:70%"></colgroup>
    <tr><td class="lb" colspan="2">REFERENCES</td></tr>
    <tr class="th"><td>NAME</td><td>LOCATION</td></tr>
    ${REFS.map((n, i) => `<tr><td>${esc(n)}</td>${ed("ref:" + i)}</tr>`).join("")}
  </table>
</main></body></html>`;
  }

  async function buildDocx({ val, groups, hlOn = true }) {
    const D = await loadLib(DOCX_URL, "docx");
    // Page geometry and table styling mirror the PDF (@page A4, margin 14mm; td border 1px #bfbfbf, padding 5px 8px).
    const MARGIN = 794; // 14mm in twips
    const FULL = 11906 - 2 * MARGIN; // A4 text width in twips
    const pct = (...p) => { const w = p.map((x) => Math.round((FULL * x) / 100)); w[w.length - 1] += FULL - w.reduce((s, x) => s + x, 0); return w; };
    const line = { style: D.BorderStyle.SINGLE, size: 6, color: "BFBFBF" };
    const borders = { top: line, bottom: line, left: line, right: line, insideHorizontal: line, insideVertical: line };
    const run = (text, o = {}) => new D.TextRun({ text: text || "", bold: o.bold, size: o.size || 20, color: o.color, font: "Calibri", highlight: o.hl ? "yellow" : undefined });
    // Text with links → runs; each link becomes a clickable, shaded "chip" (icon label + title).
    const runs = (text, o = {}) => hlParts(text).flatMap((seg) => linkRuns(seg.text, Object.assign({}, o, { hl: seg.hl && hlOn })));
    const linkRuns = (text, o = {}) => linkParts(text).map((p) => {
      if (typeof p === "string") return run(p, o);
      const [ic, c] = KINDS[p.kind] || KINDS.link, shade = { type: D.ShadingType.CLEAR, fill: "F1F3F4", color: "auto" };
      return new D.ExternalHyperlink({ link: p.url, children: [
        new D.TextRun({ text: " " + ic + " ", bold: true, size: 14, color: "FFFFFF", font: "Calibri", shading: { type: D.ShadingType.CLEAR, fill: c.slice(1), color: "auto" } }),
        new D.TextRun({ text: " " + p.title + " ", size: o.size || 20, bold: o.bold, color: "1155CC", underline: {}, font: "Calibri", shading: o.hl ? undefined : shade, highlight: o.hl ? "yellow" : undefined }),
      ] });
    });
    const para = (text, o = {}) => new D.Paragraph({ alignment: o.center ? D.AlignmentType.CENTER : undefined, children: runs(text, o) });
    // Cells are described first and materialised in table(), which knows the column widths (fixed widths keep every table full-width).
    const cell = (children, o = {}) => ({ children: Array.isArray(children) ? children : [children], span: o.span || 1, fill: o.fill, pad: o.pad });
    const lbl = (t, span, fill = "D9D9D9") => cell(para(t, { bold: true }), { span, fill });
    const edc = (k, span, o = {}) => cell(para(val(k), o), { span });
    const table = (cols, rows) => new D.Table({
      width: { size: FULL, type: D.WidthType.DXA }, columnWidths: cols, layout: D.TableLayoutType.FIXED, borders,
      rows: rows.map((r) => {
        let col = 0;
        return new D.TableRow({ cantSplit: true, children: r.map((c) => {
          const w = cols.slice(col, col + c.span).reduce((s, x) => s + x, 0); col += c.span;
          return new D.TableCell({
            children: c.children, columnSpan: c.span > 1 ? c.span : undefined, width: { size: w, type: D.WidthType.DXA },
            shading: c.fill ? { fill: c.fill, type: D.ShadingType.CLEAR, color: "auto" } : undefined,
            margins: c.pad || { top: 75, bottom: 75, left: 120, right: 120 },
          });
        }) });
      }),
    });
    const gap = () => new D.Paragraph({ spacing: { before: 0, after: 0, line: 390, lineRule: D.LineRuleType.EXACT }, children: [] }); // 26px, like table margin-bottom
    const thirds = pct(34, 33, 33), quarters = pct(25, 25, 25, 25);

    const kids = [
      new D.Paragraph({ alignment: D.AlignmentType.CENTER, spacing: { after: 210 }, children: runs("REQUIREMENTS DOCUMENT", { bold: true, size: 22 }) }),
      table(thirds, [
        [lbl("PROJECT NAME:", 2), lbl("Type")], [edc("project", 2, { bold: true }), edc("type")],
        [lbl("STAKEHOLDER"), lbl("US UI/UX:", 2)], [edc("stakeholder"), edc("uiux", 2)],
        [lbl("BA AUTHOR:", 3)], [edc("ba", 3)], [lbl("PROJECT MANAGER:", 3)], [edc("pm", 3)],
        [lbl("DEVELOPMENT TEAM: (INCLUDING UI/UX)", 3)], [edc("dev", 3)], [lbl("SQA TEAM:", 3)], [edc("sqa", 3)],
        [lbl("OTHER TEAMS:", 3)], [edc("other", 3)], [lbl("PRIORITY:"), lbl("BRAND:", 2)], [edc("priority"), edc("brand", 2)],
      ]), gap(),
      table(quarters, [
        [cell([para("VERSIONS", { bold: true }), para(val("ver:title"))], { span: 4, fill: "D9D9D9" })],
        ["VERSION NO.", "BR/FR NO.", "REVISION REMARKS", "LAST UPDATED DATE"].map((t) => cell(para(t, { bold: true, center: true }), { fill: "ECECEC" })),
        ...[0, 1, 2].map((r) => [0, 1, 2, 3].map((c) => edc("ver:" + r + ":" + c, undefined, { center: true }))),
      ]), gap(),
    ];

    const frCols = pct(13, 27, 60), frRows = [[lbl("FUNCTIONAL REQUIREMENTS (FR)", 3)]];
    groups.forEach((g) => {
      frRows.push([cell(para(val("sec:" + g.page), { bold: true }), { span: 3, fill: "D9D9D9" })]);
      g.rows.forEach((r) => {
        if (r.shot) {
          const bin = atob(r.shot.src.split(",")[1]), bytes = new Uint8Array(bin.length);
          for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
          const W = Math.min(Math.floor((FULL - 300) / 15), r.shot.w), H = Math.round((r.shot.h * W) / r.shot.w); // fit cell width minus 10px padding
          frRows.push([cell(new D.Paragraph({ alignment: D.AlignmentType.CENTER, children: [new D.ImageRun({ type: "jpg", data: bytes, transformation: { width: W, height: H } })] }), { span: 3, pad: { top: 150, bottom: 150, left: 150, right: 150 } })]);
        }
        const desc = r.lines.map((l) => para((l.bullet ? "•  " : "") + l.text));
        if (r.next.length) desc.push(para("Next: " + r.next.join(", "), { color: "555555", size: 18 }));
        frRows.push([
          cell(para(r.fr, { bold: true })),
          cell([para(r.title, { bold: true }), para(r.target, { color: "7A7A7A", size: 16 })]),
          cell(desc.length ? desc : [para("—", { color: "AAAAAA" })]),
        ]);
      });
    });
    kids.push(table(frCols, frRows), gap());
    kids.push(table(quarters, [
      [lbl("PHASED IN REQUIREMENTS", 4)],
      ["FR NO.", "REQUIREMENT", "TARGET PHASE", "REMARKS"].map((t) => cell(para(t, { bold: true, center: true }), { fill: "ECECEC" })),
      ...[0, 1].map((r) => [0, 1, 2, 3].map((c) => edc("ph:" + r + ":" + c))),
    ]), gap());
    kids.push(table(pct(30, 70), [
      [lbl("REFERENCES", 2)],
      [cell(para("NAME", { bold: true }), { fill: "ECECEC" }), cell(para("LOCATION", { bold: true }), { fill: "ECECEC" })],
      ...REFS.map((n, i) => [cell(para(n)), edc("ref:" + i)]),
    ]));

    const doc = new D.Document({
      title: val("project") + " — Requirements Document", creator: "Annotator",
      sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN } } }, children: kids }],
    });
    return D.Packer.toBlob(doc);
  }

  // Keep notes in sync if another tab edits them.
  window.addEventListener("storage", (e) => { if (e.key === KEY) { notes = load(); updateToolbar(); } });

  if (document.body) mount(); else document.addEventListener("DOMContentLoaded", mount);
})();

---
name: annotate-mockup
description: Install the ⌘-click annotation layer on an HTML mockup, so anyone viewing it can ⌘/Ctrl-click any button, card or section, type a description, and see it as a numbered callout with an arrow pointing at that element. Use when the user provides a new HTML mockup/prototype (.html or .dc.html) and asks to add annotations, notes, descriptions, "the annotator", or "the comment/description feature", or types /annotate-mockup.
---

# Annotate mockup

Adds `annotator.js` (bundled next to this file) to one or more HTML mockups. The script is framework-free vanilla JS: it works on plain HTML, React/`.dc.html` pages, or anything else rendered in the light DOM. It never modifies the page's own code.

## What the user gets

- Hold **⌘** (Ctrl on Windows) → hovered element is outlined in blue.
- **⌘-click** → description box opens; the element's own click does NOT fire.
- **Save** (or ⌘+Enter) → numbered callout + curved arrow pointing at the element. Esc cancels.
- Callouts are draggable, follow the element on scroll/resize, hide when the element isn't on screen **or is covered** (e.g. by an open modal's overlay or a sticky header) — so with a modal open only the modal's own boxes show, and the page's come back when it closes. New boxes are placed so they don't stack on top of each other. Each note re-finds its element by its saved path, then by tag + text, or for text-less elements (inputs, icons, images) by a saved label (placeholder / aria-label / title / alt / name); when several elements match (e.g. the same team name on the page and inside a modal), the visible one wins.
- Hover a callout → Edit / ✕. Double-click also edits. ⌘-clicking an annotated element reopens its note.
- **Custom numbers:** the editor has a **No.** field (e.g. `1`, `2a`, `Step 3`); leave it blank to number automatically. Clicking a callout's number badge jumps straight to editing it.
- **Connect callouts:** hover a callout → **Connect →**, then click another callout to draw a dashed arrow between them (Step 1 → Step 2 …). Do the same pair again to disconnect, or click the arrow to remove it. Esc cancels. The toolbar's **Connect in order** links every note by its number: 1 → 2 → 3 …
- **Hide one box:** hover a callout → **Hide**. It collapses to a small dashed numbered pin on the element; click the pin to show it again. The toolbar's **Show hidden (N)** brings them all back.
- **Screenshots:** a moment after a note is saved, and again after its box is dragged or resized, a screenshot is captured (html2canvas, loaded from jsDelivr on first use) and stored in IndexedDB. It shows the element *and its callout exactly as on the page*: the real box (same size, position, text and link chips), the dashed outline and the arrow. The report uses these as-is; older screenshots without the box are retaken when the report runs if that element is on screen.
- **Report = FRD (Requirements Document):** the toolbar's **Report** button opens a Functional Requirements Document in the team's BA format: title "REQUIREMENTS DOCUMENT"; header table (Project name, Type, Stakeholder, US UI/UX, BA author, Project manager, Development team, SQA team, Other teams, Priority, Brand); Versions table; **FUNCTIONAL REQUIREMENTS (FR)** with one numbered section per page (e.g. "1. Team Stores Landing"), and for each note a screenshot with its callout (FR label + title, arrow, dashed outline) drawn on, followed by an `FR No. | Title | Description` row; then Phased-in requirements and References (Figma, Monday ticket, Hand-off video, Technical specs, JIRA). Note format: **first line = FR title**, following lines = description (lines starting with `-`/`•` become bullets); connections show as "Next: FR-x.y". Custom numbers (e.g. `FR-1.1`) are used as-is and sorted naturally (FR-1.2 < FR-1.10); auto-numbered notes become `FR-<section>.<n>`. All white cells (header fields, versions, section titles, phased-in, reference links) are editable in the report and remembered in localStorage key `annotator-frd`. If other mockup pages on the same site have notes, pick "This page" or "All annotated pages" (sections ordered by when each page was first annotated). Output: **Save as PDF** (print dialog) or **Download Word (.docx)** (built with the `docx` library from jsDelivr; opens in Word / Google Docs). Notes with no screenshot are captured when the report runs, if their element is on screen.
- **Links as chips + titled links:** any URL in a note or in an editable report cell shows as a clickable chip (icon + readable title), like smart chips in Google Docs. To give a link its own title, use **🔗 Link** in the note editor (or ⌘/Ctrl+K), or **🔗 Insert link** / ⌘K in the report while a white cell is being edited: a dialog asks for **Title** and **Link**, prefilled from the link under the cursor or the selected text; it is stored as `[Title](https://…)`. Without a title, the chip title comes from the URL (PDF/Office file name, Confluence page title, JIRA key, Figma file name, Monday item, Google Doc/Sheet/Slides, video). In the report, clicking a chip in an editable cell shows **Open · Edit · Remove**; chips elsewhere open directly. In the Word file they are real hyperlinks.
- **Highlights (questions to check):** in the note editor select text and click **🖍 Highlight** (or ⌘/Ctrl+Shift+H); with nothing selected it highlights the current line; inside a highlight it removes it. Stored as `==text==`; shown yellow on the callout, in its screenshot, in the report, PDF and Word (real Word highlight). The report shows a **🖍 To check** list (screen only, not printed) with each highlighted item and its FR number — click one to jump to it — plus a **N to check · Next ↓** button in its toolbar and a **Highlights in export** checkbox to print/export without the yellow.
- **All notes panel:** toolbar **All notes** opens a side panel listing every note on the page (sorted by number), even ones whose box can't be seen, each with a status — Visible · Hidden (pin) · Off screen · Behind something (e.g. a modal) · Not found on this screen / Element not shown — and actions **Go to** (scrolls to it and flashes the box), **Edit** (inline: number + text, works even when the element is missing), **Hide/Show**, **Re-attach** (then ⌘/Ctrl-click the element the note should point at — fixes notes that lost their element) and **Delete**. Checkboxes + **Select all** / **Select not found** + **Delete selected** remove several at once.
- Bottom-left toolbar: note count · All notes · Hide/Show · Show hidden · Connect in order · Report · Export (JSON, includes screenshots) · Import · Clear.
- Notes persist in `localStorage`, keyed per page file name. Sharing between people/machines = Export → Import.

## Steps

1. **Find the target file(s).** Use the path(s) the user gave. If none were given, look for recently added `.html` / `.dc.html` files in the working directory and confirm which ones. Read the project's `CLAUDE.md` first — if it says a change must apply to several related files (e.g. "both workspace files"), install on all of them.

2. **Skip if already installed.** `grep -l "annotator.js\|__annotatorLoaded" <file>` — if it matches, tell the user it's already there and stop for that file.

3. **Copy the script** next to the HTML file (same folder), unless an `annotator.js` already exists there:
   ```bash
   cp ~/.claude/skills/annotate-mockup/annotator.js "<folder of the html>/annotator.js"
   ```
   If one already exists, `diff` it against the bundled copy; if they differ, ask before overwriting.

4. **Add one script tag** inside `<head>`:
   - If the page has `<script src="./support.js"></script>` (dc-runtime pages), put it on the line right after that.
   - Otherwise put it immediately before `</head>`.
   ```html
   <script src="./annotator.js?v=<first 8 chars of shasum annotator.js>"></script>
   ```
   The `?v=` stops browsers from running a cached old copy after the script is updated. Adjust the relative path if the HTML and the script are in different folders. Change nothing else in the file.

   **Single-file exception:** if the mockup must stay one self-contained file (the user says so, or it will be published/sent as a lone `.html`), inline the script instead: `<script>` + contents of `annotator.js` + `</script>` before `</head>`.

5. **Verify.** `node --check annotator.js`, then, if a browser is available (Playwright with `channel: 'chrome'`, installed in the scratchpad — not the project), serve the folder (`python3 -m http.server`), load the page and check:
   - the `.ann-bar` toolbar is visible,
   - ⌘-click on a real button (dismiss any welcome/intro modal first; exclude elements inside `[data-annotator-ui]`) opens `.ann-editor`, and the button's own action did not run,
   - after saving, `.ann-card` is visible, and it still shows after a reload.
   Screenshot and look at it. Stop the server afterwards and don't leave test files in the project.

6. **Report** which files got the tag, where `annotator.js` was copied, whether the browser check passed, and remind the user that notes live in their browser only (Export/Import to share).

## Removing it

Delete the `<script src="./annotator.js"></script>` line (or the inlined block). Delete `annotator.js` only if no other page in that folder still references it.

## Updating the script

`annotator.js` in this skill folder is the master copy. After improving it, re-copy it to the projects that use it (`grep -rl "annotator.js" --include="*.html" <dir>` finds them), then bump the cache-busting version in each page's tag so browsers don't keep the old copy:
```bash
V=$(shasum annotator.js | cut -c1-8)
sed -i '' -E "s#<script src=\"\./annotator\.js(\?v=[a-z0-9]+)?\"></script>#<script src=\"./annotator.js?v=$V\"></script>#" "<page>.html"
```
If a change fixes how screenshots are taken, also bump `SHOT_V` in the script so the report retakes screenshots from older versions.

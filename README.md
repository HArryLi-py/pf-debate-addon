# PF Debate — Google Docs Add-on

A Google Docs Editor add-on (Google Apps Script) for Public Forum debate. Centres
the debater's writing workflow: format speech headings/blocks, and generate
evidence cards into a separate "PF Evidence Cards" doc with a cross-doc link
back from the debate doc.

## What it does (Phase 1)

Three sidebar buttons, plus a card-generation form:

1. **标题 (Title)** — select text in the Doc → the whole paragraph becomes a
   centered 27pt bold-underlined heading (so it shows in the native Outline
   panel).
2. **blocks** — select text → gray 24pt underline.
3. **生成 evidence card** — fill the form (title / citation / quote / highlight)
   → appends a card to a separate **"PF Evidence Cards"** Google Doc (created
   on first use, ID remembered in User Properties) and inserts a hyperlink at
   the cursor in the debate doc that opens that card's doc.

Card layout in the Evidence Cards doc:
- Title — 24pt, bold, underline (heading)
- Citation — 12pt
- Quote — 9pt (small); the highlighted excerpt (bold + underline + yellow
  background) is matched within the quote text
- blank line between cards

## Install (container-bound script)

1. Open a Google Doc → **Extensions → Apps Script** (binds the script to the Doc).
2. Enable the Advanced Google Docs API? **Not needed** — this version uses only
   `DocumentApp`, no Advanced Service, no `serviceId` headache.
3. Paste / `clasp push` these files: `appsscript.json`, `Code.gs`, `Sidebar.html`.
4. Run `onOpen` once → OAuth consent (scopes: `documents`,
   `script.container.ui`, `drive.file` — the last is for creating the Evidence
   Cards doc; first "生成 card" click triggers the consent).
5. Reload the Doc → **Extensions → Add-ons → PF Debate → Open sidebar**.

> The sidebar UI is a light theme with a premium-blue primary action, hover/press
> motion, and a one-off success flash on card generation. Respects
> `prefers-reduced-motion`.

## Files

| File | Purpose |
|------|---------|
| `appsscript.json` | Manifest: scopes (documents, script.container.ui, drive.file), V8. No Advanced Service. |
| `Code.gs` | Menu/sidebar + `applyTitle` / `applyBlocks` / `generateCard` + helpers. |
| `Sidebar.html` | Sidebar UI: 2 format buttons + card form + status. |
| `schemas/` | JSON schemas: evidence card, speech outline, flow grid, NSDA PF round timing, judge ballot. |
| `spike/` | Firebase-in-sandbox spike (gates Phase 2 — synced timer/chat across debaters). |

## Scope & roadmap

- **Phase 1 (this):** formatting buttons + evidence-card generation into a
  separate doc + cross-doc link.
- **Phase 2 (gated by `spike/`):** Firebase/Supabase-backed shared timer, NSDA
  round phase sequencer, text chat, presence across 4–5 debaters.
- **Phase 3:** external WebRTC audio room (LiveKit/Jitsi) launched from the
  sidebar — audio can't run inside the Docs add-on sandbox (getUserMedia is
  hard-blocked by the HtmlService iframe).

## Honest constraints

- **No drag-from-sidebar-to-Doc-canvas** — Apps Script sidebars are sandboxed
  iframes; insert is click-at-cursor, not drag.
- **No live audio in the sidebar** — platform hard wall; audio is Phase 3
  external.
- **Cross-doc link opens the Evidence Cards doc** (not yet an exact-card
  bookmark anchor — the card title is a heading, so the native Outline lists
  cards for quick jump).

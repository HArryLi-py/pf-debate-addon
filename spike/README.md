# Firebase-in-sandbox spike

The ONE open technical risk before committing to the Phase 2 architecture
(Firebase-backed shared timer / phase / chat across 4–5 debaters' sidebars):
**can a Google Docs add-on sidebar (HtmlService IFRAME sandbox) sustain a
client-side realtime connection to Firebase RTDB with sub-second propagation
between two sidebars on the same Doc?**

Evidence leans yes (sandbox attrs include `allow-same-origin`; WebSocket isn't
CORS-enforced; Google's own 2015 Firebase-in-sidebar blog; no restrictive CSP
blocking gstatic/firebaseio). But there's no 2024–2026 end-to-end confirmation,
so we test instead of assuming.

## Setup (≈30 min)

1. **Make a test Firebase project.** console.firebase.google.com → new project
   → Realtime Database → create (test mode). Note the `databaseURL`
   (`https://<project>.firebaseio.com`).
2. **Open RTDB rules** → paste `firebase-rules.json` → publish. (Open r/w on
   `/spike` for the test only — never ship.)
3. **Bind the spike script to a Google Doc.** Open a Doc → Extensions →
   Apps Script. Paste `appsscript.json`, `Code.gs`, `Sidebar.html`. Edit the
   `FIREBASE_CONFIG.dbHost` in `Code.gs` to your `databaseURL`.
4. **Authorize.** Run `onOpen` once → consent (documents,
   script.container.ui, script.external_request). Reload the Doc.
5. **Open two sidebars in two Google accounts.** Account A owns the Doc +
   shares edit with Account B. Both open the sidebar (Extensions → Add-ons →
   Firebase Spike → Open realtime sidebar). The `roundId` is derived from the
   Doc id, so both sidebars subscribe to the same node.

## Run + interpret

- **Tier 1** auto-loads Firebase v9 modular SDK via ESM `import()` from gstatic
  and patches `window.WebSocket` to detect the transport actually used.
- Write a message in one sidebar; watch the other receive it.
- **SUCCESS = transport badge reads `websocket` AND propagation < 1000 ms.**
  Both conditions — "connected" alone is NOT success.
- If transport reads `long-polling` / `compat-umd` / `rest-stream`, it
  technically connected but you have NOT met the sub-second bar (long-polling
  adds 1–3s; REST stream 200–500ms+; server-poll 250–400ms+).

## Fallbacks (auto-activate if a prior tier fails)

- **Tier 1b:** v9-compat UMD via plain `<script src>` (only if ESM blocked).
- **Tier 2:** Firebase REST streaming via `EventSource` (SSE). ~200–500ms+.
- **Tier 3:** `google.script.run` → server `UrlFetchApp` → Firebase REST, polled
  at 400ms. ~250ms RTT, **not realtime**. Last resort — also what you'd ship if
  WebSocket is fully blocked (acceptable for chat + turn-indication, marginal
  for sub-second timer sync).

## Notes

- Firebase v9 modular ESM uses `import()` inside a classic `<script>`, NOT
  `<script type="module">`. A rejected `import()` is catchable → clean fallback.
  A top-level module script that fails to load is NOT catchable.
- Sub-second propagation REQUIRES WebSocket (Tier 1). If the corporate/Workspace
  network blocks `wss://*.firebaseio.com`, the SDK sits in a 20–30s retry before
  falling back to long-polling internally — the 6s transport-resolve timer may
  report `long-polling` prematurely; watch the log for a late `ws open` line.
- The spike uses an unauthenticated open-rules test node. Don't add Firebase
  Auth to this spike — it adds a second unknown (can the sandbox run the auth
  popup?) and muddies the single risk under test. Auth comes in Phase 2 proper
  (custom tokens minted by the Apps Script server from the user's email).
- If BOTH ESM and UMD fail to load (`script onerror`), check the Workspace
  admin console → Apps Script → URL allowlist before concluding the
  architecture is broken — that's a policy block, not a sandbox limit.
- EventSource (Tier 2) can't set Authorization headers, so it only works while
  RTDB rules are open. Once rules require a token, Tier 2 must switch to
  `fetch()`+`ReadableStream` on the `.json` stream, or upgrade to Tier 3.

## Decision after the spike

- **Tier 1 websocket, <1s** → Phase 2 architecture confirmed: Firebase RTDB
  loaded via CDN ESM into the sidebar, sub-second shared timer/phase/chat.
- **Tier 1b/2 only** → use Firebase but accept higher latency, or move to
  Supabase Realtime (different transport).
- **Only Tier 3** → realtime-in-sidebar is dead. Shared state goes through a
  server-side `UrlFetchApp` proxy + polling (200ms–1s+); timer sync will be
  rough. Consider whether that's acceptable, or scope Phase 2 down.

/**
 * PF Debate Add-on — Phase 1: formatting buttons + evidence card appended to
 * the current (debate) doc. No separate evidence doc, no Advanced Docs Service,
 * no drive.file scope.
 *
 * Three sidebar buttons:
 *   1. 标题 (Title)  — whole paragraph → centered 27pt bold underline.
 *   2. blocks       — whole paragraph → centered 24pt gray underline.
 *   3. 生成 evidence card — append a card (title/citation/quote+highlight) to
 *      the bottom of THIS doc.
 *
 * Why no heading on the Title button: applying HEADING1 was clobbering the
 * center + bold (heading style overrode direct formatting in render). Pure
 * direct formatting now — alignment + font size + bold + underline — so it
 * reliably centers and bolds. (Card title still uses HEADING2 for Outline
 * structure; that one isn't centered so no conflict.)
 */

/* ============================== MENU + SIDEBAR ============================== */
function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
    .addItem('Open sidebar', 'showSidebar')
    .addSeparator()
    .addItem('Title selection (centered bold underline)', 'applyTitle')
    .addItem('Block selection (centered gray underline)', 'applyBlocks')
    .addItem('Normal text (13pt no bold no underline)', 'applyNormal')
    .addToUi();
}
function onInstall(e) { onOpen(e); }

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('PF Debate')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  DocumentApp.getUi().showSidebar(html);
}

/* Open a separate modeless dialog for chat + round-phase (WeChat-style chat).
   Stays open alongside the main sidebar — non-blocking. */
function showChatDialog() {
  var html = HtmlService.createHtmlOutputFromFile('Chat')
    .setTitle('PF Debate · 聊天')
    .setWidth(400).setHeight(560)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  DocumentApp.getUi().showModelessDialog(html, 'PF Debate · 聊天');
}

/* ====================== SELF-TEST (sidebar calls on load) ================= */
function ping() {
  var d = DocumentApp.getActiveDocument();
  var sel = d.getSelection();
  return 'pong ✓ doc="' + d.getName() + '" selection=' + (sel ? 'yes' : 'none') + ' cursor=' + (d.getCursor() ? 'yes' : 'none');
}

/* ====================== PHASE 2: FIREBASE CONFIG / ROUND ID ===============
   The Firebase databaseURL is stored per-user (UserProperties) so debaters
   don't re-enter it each session. The round ID is derived from the Doc ID, so
   everyone on the same Doc shares the same sync node. */
var FB_URL_KEY = 'pf_fb_url';
function getFbUrl() { return PropertiesService.getUserProperties().getProperty(FB_URL_KEY) || ''; }
function setFbUrl(url) { PropertiesService.getUserProperties().setProperty(FB_URL_KEY, url || ''); return 'saved'; }
function getRoundId() { try { return 'doc-' + DocumentApp.getActiveDocument().getId().slice(-8); } catch (e) { return 'round-default'; } }

/* ====================== SETTINGS ==========================================
   cardDestination: 'tab' (evidence cards tab) or 'append' (current doc bottom)
   language: 'zh' or 'en' */
var SETTINGS_KEY = 'pf_settings';
function getSettings() {
  var s = PropertiesService.getUserProperties().getProperty(SETTINGS_KEY);
  return s ? JSON.parse(s) : { cardDestination: 'tab', language: 'zh', titleSize: 27, blockSize: 24 };
}
function setSettings(settings) {
  PropertiesService.getUserProperties().setProperty(SETTINGS_KEY, JSON.stringify(settings));
  return 'saved';
}

/* ====================== PHASE 2: CHAT IDENTITY ============================
   Each user gets a persistent display name (UserProperties) for the chat.
   Generated on first use; editable from the sidebar. */
var MY_NAME_KEY = 'pf_my_name';
function getMyName() {
  var p = PropertiesService.getUserProperties();
  var n = p.getProperty(MY_NAME_KEY);
  if (n) return n;
  n = 'Debater-' + Math.random().toString(36).slice(2, 6);
  p.setProperty(MY_NAME_KEY, n);
  return n;
}
function setMyName(name) {
  PropertiesService.getUserProperties().setProperty(MY_NAME_KEY, name || 'Debater');
  return 'saved';
}

var MY_AVATAR_KEY = 'pf_my_avatar';
function getMyAvatar() { return PropertiesService.getUserProperties().getProperty(MY_AVATAR_KEY) || ''; }
function setMyAvatar(url) { PropertiesService.getUserProperties().setProperty(MY_AVATAR_KEY, url || ''); return 'saved'; }

/* ====================== BUTTON 1: TITLE (选中→标题) ======================
   Whole paragraph → centered + 27pt + bold + underline. Handles Paragraph and
   ListItem; falls back to cursor's paragraph. Styles both the selection runs
   and the whole paragraph so it works even on partial selections. */
function applyTitle() {
  var sz = getSettings().titleSize || 27;
  var doc = DocumentApp.getActiveDocument();
  var paras = _targetParagraphs(doc);
  var runs = _styleSelectionRuns(doc, function (t, s, e) {
    t.setFontSize(s, e, sz).setBold(s, e, true).setUnderline(s, e, true).setForegroundColor(s, e, '#000000');
  });
  paras.forEach(function (p) {
    try { p.setHeading(DocumentApp.ParagraphHeading.HEADING1); } catch (e) {}
    try { p.setAlignment(DocumentApp.HorizontalAlignment.CENTER); } catch (e) {}
    try { var t = p.editAsText(); if (t) t.setFontSize(sz).setBold(true).setUnderline(true).setForegroundColor('#000000'); } catch (e) {}
  });
  if (!paras.length && !runs) throw new Error('没找到可改的段落。先在 Doc 里用鼠标拖选一段字，再点按钮。');
  return '标题 ✓ 段落 ' + paras.length + ' / 文本段 ' + runs + '（居中 + ' + sz + 'pt + 加粗 + 下划线 + 黑字 + 进大纲）';
}

/* ====================== BUTTON 2: BLOCKS (选中→灰块) =====================
   Whole paragraph → centered + gray + underline. Size from settings. */
function applyBlocks() {
  var sz = getSettings().blockSize || 24;
  var doc = DocumentApp.getActiveDocument();
  var paras = _targetParagraphs(doc);
  var runs = _styleSelectionRuns(doc, function (t, s, e) {
    t.setFontSize(s, e, sz).setUnderline(s, e, true).setForegroundColor(s, e, '#666666');
  });
  paras.forEach(function (p) {
    try { p.setHeading(DocumentApp.ParagraphHeading.HEADING2); } catch (e) {}
    try { p.setAlignment(DocumentApp.HorizontalAlignment.CENTER); } catch (e) {}
    try { var t = p.editAsText(); if (t) t.setFontSize(sz).setUnderline(true).setForegroundColor('#666666'); } catch (e) {}
  });
  if (!paras.length && !runs) throw new Error('没找到可改的段落。先在 Doc 里用鼠标拖选一段字，再点按钮。');
  return 'blocks ✓ 段落 ' + paras.length + ' / 文本段 ' + runs + '（居中 + ' + sz + 'pt + 灰字 + 下划线 + 进大纲 H2）';
}

/* ====================== BUTTON 3: 写正文 (选中→正常文字) =====================
   13pt, no bold, no underline, black. For body text after headings. */
function applyNormal() {
  var doc = DocumentApp.getActiveDocument();
  var paras = _targetParagraphs(doc);
  var runs = _styleSelectionRuns(doc, function (t, s, e) {
    t.setFontSize(s, e, 13).setBold(s, e, false).setUnderline(s, e, false).setForegroundColor(s, e, '#000000');
  });
  paras.forEach(function (p) {
    try { p.setHeading(DocumentApp.ParagraphHeading.NORMAL); } catch (e) {}
    try { var t = p.editAsText(); if (t) t.setFontSize(13).setBold(false).setUnderline(false).setForegroundColor('#000000'); } catch (e) {}
  });
  if (!paras.length && !runs) throw new Error('没找到可改的段落。先在 Doc 里用鼠标拖选一段字，再点按钮。');
  return '正文 ✓ 段落 ' + paras.length + ' / 文本段 ' + runs + '（13pt + 不加粗 + 不下划线 + 黑字）';
}

/** Style each text run in the current selection via fn(textRun, start, end). */
function _styleSelectionRuns(doc, fn) {
  var sel = doc.getSelection();
  if (!sel) return 0;
  var n = 0;
  sel.getRangeElements().forEach(function (re) {
    var el = re.getElement();
    if (el.getType() !== DocumentApp.ElementType.TEXT) return;
    var t = el.asText();
    var s = re.isPartial() ? re.getStartOffset() : 0;
    var e = re.isPartial() ? re.getEndOffsetInclusive() : t.getText().length - 1;
    if (e >= s) { fn(t, s, e); n++; }
  });
  return n;
}

/** Paragraph/List-Item containers of the selection; falls back to the cursor's. */
function _targetParagraphs(doc) {
  var paras = [];
  var seen = [];
  var add = function (p) { if (p && seen.indexOf(p) === -1) { seen.push(p); paras.push(p); } };
  var walk = function (el) {
    var cur = el;
    while (cur) {
      var tp = cur.getType();
      if (tp === DocumentApp.ElementType.PARAGRAPH || tp === DocumentApp.ElementType.LIST_ITEM) { add(cur); return; }
      cur = cur.getParent ? cur.getParent() : null;
    }
  };
  var sel = doc.getSelection();
  if (sel) sel.getRangeElements().forEach(function (re) { walk(re.getElement()); });
  if (!paras.length) { var c = doc.getCursor(); if (c) walk(c.getElement()); }
  return paras;
}

/* ================ BUTTON 3: GENERATE EVIDENCE CARD (生成卡片) ==============
   Card goes into a tab named "evidence cards" (user creates it once — Apps
   Script can't create tabs programmatically). A "card N" hyperlink is inserted
   at the cursor in the current tab → links to that tab. Highlight in the quote
   is rendered bigger (13pt) than the rest of the quote (9pt). */

var EV_TAB_NAME = 'evidence cards';
var CARD_COUNT_KEY = 'pf_card_count';

function generateCard(form) {
  if (!form || !form.title || !form.citation || !form.quote) {
    throw new Error('标题、citation、quote 都要填。');
  }
  var doc = DocumentApp.getActiveDocument();
  var settings = getSettings();
  var editUrl = 'https://docs.google.com/document/d/' + doc.getId() + '/edit';

  // Determine where to put the card
  var targetBody, evTab = null, targetLabel;
  if (settings.cardDestination === 'append') {
    targetBody = doc.getBody();
    targetLabel = '本文档底部';
  } else {
    evTab = _findEvidenceTab(doc);
    if (!evTab) {
      throw new Error('没找到名为 "evidence cards" 的标签页。请在 Doc 左下角标签栏点 ＋ 新建一个标签页，名字填 "evidence cards"，或者在设置里改成"本文档底部"。');
    }
    try { targetBody = evTab.asDocumentTab().getBody(); } catch (e) { throw new Error('无法读取 evidence cards 标签页：' + e.message); }
    targetLabel = 'evidence cards 标签页';
  }

  // increment card counter
  var props = PropertiesService.getUserProperties();
  var n = parseInt(props.getProperty(CARD_COUNT_KEY) || '0', 10) + 1;
  props.setProperty(CARD_COUNT_KEY, String(n));

  // append the card
  targetBody.appendParagraph('');
  var head = targetBody.appendParagraph('card ' + n + ': ' + form.title);
  try { head.setHeading(DocumentApp.ParagraphHeading.HEADING2); } catch (e) {}
  head.editAsText().setFontSize(24).setBold(true).setUnderline(true).setForegroundColor('#000000');
  targetBody.appendParagraph(form.citation).editAsText().setFontSize(12);
  var quotePara = targetBody.appendParagraph(form.quote);
  var qt = quotePara.editAsText();
  qt.setFontSize(9);
  var highlightFound = false;
  if (form.highlight) {
    var idx = form.quote.indexOf(form.highlight);
    if (idx >= 0) {
      var end = idx + form.highlight.length;
      var hlColor = form.highlightColor || '#FFF3A0';
      qt.setFontSize(idx, end - 1, 20)
        .setBold(idx, end - 1, true)
        .setUnderline(idx, end - 1, true)
        .setBackgroundColor(idx, end - 1, hlColor);
      highlightFound = true;
    }
  }

  // build a link to the exact card
  var linkUrl = editUrl;
  var method = 'doc-only';
  // Method A: bookmark via editAsText().createPosition(0) — fixed from head.createPosition(0)
  try {
    var pos = head.editAsText().createPosition(0);
    var bm = (settings.cardDestination === 'append') ? doc.getActiveTab().asDocumentTab().addBookmark(pos) : evTab.asDocumentTab().addBookmark(pos);
    var bmId = bm.getId();
    if (bmId) {
      linkUrl = editUrl + '#bookmark=' + bmId;
      method = 'bookmark';
    }
  } catch (eB) {
    // Method B: tab-level fallback (only for tab destination)
    try {
      if (evTab) {
        var tabId = evTab.getId();
        if (tabId) { linkUrl = editUrl + '#tab=h.' + tabId; method = 'tab'; }
      }
    } catch (eC) {}
  }
  _insertLinkAtCursor(doc, 'card ' + n, linkUrl);

  return 'card ' + n + ' ✓ 已添加到' + targetLabel + '。链接方式=' + method + '。' +
    (form.highlight && !highlightFound ? '（highlight 在 quote 里没找到，没加高亮。）' : '');
}

/** Find a tab whose title matches EV_TAB_NAME (case-insensitive). */
function _findEvidenceTab(doc) {
  var tabs = doc.getTabs();
  for (var i = 0; i < tabs.length; i++) {
    var t = tabs[i];
    var title = '';
    try { title = t.getTitle() || ''; } catch (e) { try { title = t.asDocumentTab().getTitle() || ''; } catch (e2) {} }
    if (title && title.toLowerCase() === EV_TAB_NAME) return t;
  }
  return null;
}

/** Insert `text` as a hyperlink (→ url) at the cursor; fall back to appending on the active tab. */
function _insertLinkAtCursor(doc, text, url) {
  var cursor = doc.getCursor();
  if (cursor) {
    try {
      var el = cursor.getElement();
      var start = cursor.getOffset();
      cursor.insertText(text);
      el.asText().setLinkUrl(start, start + text.length - 1, url);
      return;
    } catch (e) { /* fall through */ }
  }
  var body;
  try { body = doc.getActiveTab().asDocumentTab().getBody(); } catch (e) { body = doc.getBody(); }
  var p = body.appendParagraph(text);
  p.editAsText().setLinkUrl(url);
}

/* ====================== P0: AGORA TOKEN SERVER-SIDE ======================
   AccessToken2 minting on server. jsSHA.gs provides HMAC-SHA256.
   Web App: doGet?channel=xxx&uid=123 → {token, appId}
   App Certificate stays in Code.gs — never exposed to front-end. */
var AGORA_APP_ID = '6fb557c64dc04c578ea45bbc0f29cac0';
var AGORA_APP_CERT = '8d95c2dad6214253b87e55592de978da';

function _agoraHmac(keyBytes, msgBytes) {
  var sha = new jsSHA("SHA-256", "ARRAY");
  sha.setHMACKey(keyBytes, "ARRAY");
  var hmacHex = sha.getHMAC(msgBytes, "ARRAY", "HEX");
  return hmacHex.match(/.{2}/g).map(function(h) { return parseInt(h, 16); });
}
function _agoraBB() { var b = []; this.u16 = function(v) { b.push(v & 255, v >> 8 & 255); return this; }; this.u32 = function(v) { b.push(v & 255, v >> 8 & 255, v >> 16 & 255, v >>> 24 & 255); return this; }; this.bytes = function(a) { this.u16(a.length); for (var i = 0; i < a.length; i++) b.push(a[i] & 255); return this; }; this.str = function(s) { var u = unescape(encodeURIComponent(s)); var a = []; for (var i = 0; i < u.length; i++) a.push(u.charCodeAt(i)); return this.bytes(a); }; this.map32 = function(m) { var k = Object.keys(m); this.u16(k.length); for (var i = 0; i < k.length; i++) { this.u16(parseInt(k[i], 10)); this.u32(m[k[i]]); } return this; }; this.pack = function() { return b; }; return this; }
function _agoraCat(arrs) { var t = 0; arrs.forEach(function(a) { t += a.length; }); var o = new Array(t); var f = 0; arrs.forEach(function(a) { for (var i = 0; i < a.length; i++) o[f++] = a[i]; }); return o; }
function _agoraAdler(b) { var a = 1, c = 0; for (var i = 0; i < b.length; i++) { a = (a + b[i]) % 65521; c = (c + a) % 65521; } return ((c << 16) | a) >>> 0; }
function _agoraZlibDef(d) { var l = d.length, n = (~l) & 65535, o = [120, 1, 1, l & 255, l >> 8 & 255, n & 255, n >> 8 & 255]; for (var i = 0; i < l; i++) o.push(d[i]); var ad = _agoraAdler(d); o.push(ad >>> 24 & 255, ad >> 16 & 255, ad >> 8 & 255, ad & 255); return o; }
function _agoraPackSvc(cn, uid, exp) { var t = new _agoraBB().u16(1).pack(); var p = { 1: exp, 2: exp, 3: exp, 4: exp }; var pb = new _agoraBB().map32(p).pack(); var cnb = new _agoraBB().str(cn).pack(); var ub = new _agoraBB().str(uid === 0 ? '' : String(uid)).pack(); return _agoraCat([t, pb, cnb, ub]); }
function _utf8Bytes(s) { var u = unescape(encodeURIComponent(s)); var a = []; for (var i = 0; i < u.length; i++) a.push(u.charCodeAt(i)); return a; }
function _agoraMkToken(cn, uid) {
  var ts = Math.floor(Date.now() / 1000), exp = 3600, salt = Math.floor(Math.random() * 99999999) + 1;
  var s1 = _agoraHmac(new _agoraBB().u32(ts).pack(), _utf8Bytes(AGORA_APP_CERT));
  var sg = _agoraHmac(new _agoraBB().u32(salt).pack(), s1);
  var svc = _agoraPackSvc(cn, uid, exp);
  var si = _agoraCat([new _agoraBB().str(AGORA_APP_ID).u32(ts).u32(exp).u32(salt).u16(1).pack(), svc]);
  var sig = _agoraHmac(sg, si);
  var ct = _agoraCat([new _agoraBB().bytes(sig).pack(), si]);
  return '007' + Utilities.base64Encode(_agoraZlibDef(ct));
}
function doGet(e) {
  var channel = (e && e.parameter && e.parameter.channel) || 'pfdebate-default';
  var uid = parseInt((e && e.parameter && e.parameter.uid) || '0', 10);
  var token = _agoraMkToken(channel, uid);
  return ContentService.createTextOutput(JSON.stringify({ token: token, appId: AGORA_APP_ID })).setMimeType(ContentService.MimeType.JSON);
}

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
    .addItem('Title selection (27pt centered bold underline)', 'applyTitle')
    .addItem('Block selection (24pt centered gray underline)', 'applyBlocks')
    .addToUi();
}
function onInstall(e) { onOpen(e); }

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('PF Debate')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  DocumentApp.getUi().showSidebar(html);
}

/* ====================== SELF-TEST (sidebar calls on load) ================= */
function ping() {
  var d = DocumentApp.getActiveDocument();
  var sel = d.getSelection();
  return 'pong ✓ doc="' + d.getName() + '" selection=' + (sel ? 'yes' : 'none') + ' cursor=' + (d.getCursor() ? 'yes' : 'none');
}

/* ====================== BUTTON 1: TITLE (选中→标题) ======================
   Whole paragraph → centered + 27pt + bold + underline. Handles Paragraph and
   ListItem; falls back to cursor's paragraph. Styles both the selection runs
   and the whole paragraph so it works even on partial selections. */
function applyTitle() {
  var doc = DocumentApp.getActiveDocument();
  var paras = _targetParagraphs(doc);
  var runs = _styleSelectionRuns(doc, function (t, s, e) {
    t.setFontSize(s, e, 27).setBold(s, e, true).setUnderline(s, e, true).setForegroundColor(s, e, '#000000');
  });
  paras.forEach(function (p) {
    try { p.setHeading(DocumentApp.ParagraphHeading.HEADING1); } catch (e) {}
    try { p.setAlignment(DocumentApp.HorizontalAlignment.CENTER); } catch (e) {}
    try { var t = p.editAsText(); if (t) t.setFontSize(27).setBold(true).setUnderline(true).setForegroundColor('#000000'); } catch (e) {}
  });
  if (!paras.length && !runs) throw new Error('没找到可改的段落。先在 Doc 里用鼠标拖选一段字，再点按钮。');
  return '标题 ✓ 段落 ' + paras.length + ' / 文本段 ' + runs + '（居中 + 27pt + 加粗 + 下划线 + 黑字 + 进大纲）';
}

/* ====================== BUTTON 2: BLOCKS (选中→灰块) =====================
   Whole paragraph → centered + 24pt + gray + underline. */
function applyBlocks() {
  var doc = DocumentApp.getActiveDocument();
  var paras = _targetParagraphs(doc);
  var runs = _styleSelectionRuns(doc, function (t, s, e) {
    t.setFontSize(s, e, 24).setUnderline(s, e, true).setForegroundColor(s, e, '#666666');
  });
  paras.forEach(function (p) {
    try { p.setAlignment(DocumentApp.HorizontalAlignment.CENTER); } catch (e) {}
    try { var t = p.editAsText(); if (t) t.setFontSize(24).setUnderline(true).setForegroundColor('#666666'); } catch (e) {}
  });
  if (!paras.length && !runs) throw new Error('没找到可改的段落。先在 Doc 里用鼠标拖选一段字，再点按钮。');
  return 'blocks ✓ 段落 ' + paras.length + ' / 文本段 ' + runs + '（居中 + 24pt + 灰字 + 下划线）';
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

  // 1. find the "evidence cards" tab (user must create it once)
  var evTab = _findEvidenceTab(doc);
  if (!evTab) {
    throw new Error('没找到名为 "' + EV_TAB_NAME + '" 的标签页。请在 Doc 左下角标签栏点 ＋ 新建一个标签页，名字填 "evidence cards"，然后再点生成。');
  }
  var evBody;
  try { evBody = evTab.asDocumentTab().getBody(); } catch (e) { throw new Error('无法读取 evidence cards 标签页：' + e.message); }

  // 2. increment card counter
  var props = PropertiesService.getUserProperties();
  var n = parseInt(props.getProperty(CARD_COUNT_KEY) || '0', 10) + 1;
  props.setProperty(CARD_COUNT_KEY, String(n));

  // 3. append the card to the evidence tab
  evBody.appendParagraph(''); // blank line separator
  var head = evBody.appendParagraph('card ' + n + ': ' + form.title);
  try { head.setHeading(DocumentApp.ParagraphHeading.HEADING2); } catch (e) {}
  head.editAsText().setFontSize(24).setBold(true).setUnderline(true).setForegroundColor('#000000');
  evBody.appendParagraph(form.citation).editAsText().setFontSize(12);
  var quotePara = evBody.appendParagraph(form.quote);
  var qt = quotePara.editAsText();
  qt.setFontSize(9);
  var highlightFound = false;
  if (form.highlight) {
    var idx = form.quote.indexOf(form.highlight);
    if (idx >= 0) {
      var end = idx + form.highlight.length; // exclusive
      var hlColor = form.highlightColor || '#FFF3A0';
      qt.setFontSize(idx, end - 1, 20)                       // highlight bigger than the 9pt quote
        .setBold(idx, end - 1, true)
        .setUnderline(idx, end - 1, true)
        .setBackgroundColor(idx, end - 1, hlColor);
      highlightFound = true;
    }
  }

  // 4. bookmark the card title heading, then insert a "card N" hyperlink → that bookmark
  //    (clicking jumps straight to the card's position in the evidence cards tab)
  var linkUrl = doc.getUrl();
  try {
    var pos = head.createPosition(0);
    var bm = evTab.asDocumentTab().addBookmark(pos);
    var bmId = bm.getId();
    if (bmId) linkUrl = doc.getUrl() + '#bookmark=id.' + bmId;
  } catch (e) {
    // fallback: link to the tab if bookmark creation isn't supported
    try { var tabId = evTab.getId(); if (tabId) linkUrl = doc.getUrl() + '#tab=h.' + tabId; } catch (e2) {}
  }
  _insertLinkAtCursor(doc, 'card ' + n, linkUrl);

  return 'card ' + n + ' ✓ 已生成：evidence cards 标签页加了卡片，当前 Doc 光标处插了超链接（点开跳到对应卡片位置）。' +
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

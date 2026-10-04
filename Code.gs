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
    t.setFontSize(s, e, 27).setBold(s, e, true).setUnderline(s, e, true);
  });
  paras.forEach(function (p) {
    try { p.setAlignment(DocumentApp.Alignment.CENTER); } catch (e) {}
    try { var t = p.editAsText(); if (t) t.setFontSize(27).setBold(true).setUnderline(true); } catch (e) {}
  });
  if (!paras.length && !runs) throw new Error('没找到可改的段落。先在 Doc 里用鼠标拖选一段字，再点按钮。');
  return '标题 ✓ 段落 ' + paras.length + ' / 文本段 ' + runs + '（居中 + 27pt + 加粗 + 下划线）';
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
    try { p.setAlignment(DocumentApp.Alignment.CENTER); } catch (e) {}
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
   Append the card to the BOTTOM of the current doc. No separate doc, no
   hyperlink, no drive.file. Cards accumulate with a blank line between. */

function generateCard(form) {
  if (!form || !form.title || !form.citation || !form.quote) {
    throw new Error('标题、citation、quote 都要填。');
  }
  var body = DocumentApp.getActiveDocument().getBody();

  // blank line separator before each card
  body.appendParagraph('');

  // Title: 24pt bold underline + heading (heading for Outline structure; not centered)
  var titlePara = body.appendParagraph(form.title);
  titlePara.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  titlePara.editAsText().setFontSize(24).setBold(true).setUnderline(true);

  // Citation: 12pt
  body.appendParagraph(form.citation).editAsText().setFontSize(12);

  // Quote: 9pt; bold+underline+yellow the highlighted excerpt (matched within the quote)
  var quotePara = body.appendParagraph(form.quote);
  var qt = quotePara.editAsText();
  qt.setFontSize(9);
  var highlightFound = false;
  if (form.highlight) {
    var idx = form.quote.indexOf(form.highlight);
    if (idx >= 0) {
      var end = idx + form.highlight.length; // exclusive
      qt.setBold(idx, end - 1, true)
        .setUnderline(idx, end - 1, true)
        .setBackgroundColor(idx, end - 1, '#FFF3A0');
      highlightFound = true;
    }
  }
  return 'Card 已追加到本文档底部。往下滚查看。' +
    (form.highlight && !highlightFound ? '（highlight 文字在 quote 里没找到，没加高亮。）' : '');
}

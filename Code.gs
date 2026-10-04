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

/* ====================== BUTTON 1: TITLE (选中→标题) ======================
   Whole paragraph → centered + 27pt + bold + underline. Pure direct formatting. */
function applyTitle() {
  return _formatSelection(
    function (p) {
      p.setAlignment(DocumentApp.Alignment.CENTER);
      var t = p.editAsText();
      if (t) t.setFontSize(27).setBold(true).setUnderline(true);
    },
    function (t, s, e) {
      t.setFontSize(s, e, 27).setBold(s, e, true).setUnderline(s, e, true);
    }
  );
}

/* ====================== BUTTON 2: BLOCKS (选中→灰块) =====================
   Whole paragraph → centered + 24pt + gray + underline. */
function applyBlocks() {
  return _formatSelection(
    function (p) {
      p.setAlignment(DocumentApp.Alignment.CENTER);
      var t = p.editAsText();
      if (t) t.setFontSize(24).setUnderline(true).setForegroundColor('#666666');
    },
    function (t, s, e) {
      t.setFontSize(s, e, 24).setUnderline(s, e, true).setForegroundColor(s, e, '#666666');
    }
  );
}

/** Apply paraFn(paragraph) to each paragraph containing the selection;
 *  textFn(textRun, start, end) as a fallback for text in lists/tables. */
function _formatSelection(paraFn, textFn) {
  var doc = DocumentApp.getActiveDocument();
  var sel = doc.getSelection();
  if (!sel) throw new Error('先在 Doc 里选中文字。');
  var done = [];
  sel.getRangeElements().forEach(function (re) {
    var el = re.getElement();
    var parent = el.getParent();
    if (parent && parent.getType() === DocumentApp.ElementType.PARAGRAPH) {
      var p = parent.asParagraph();
      if (done.indexOf(p) !== -1) return;
      done.push(p);
      paraFn(p);
    } else if (el.getType() === DocumentApp.ElementType.TEXT) {
      var t = el.asText();
      var s = re.isPartial() ? re.getStartOffset() : 0;
      var e = re.isPartial() ? re.getEndOffsetInclusive() : t.getText().length - 1;
      if (e >= s) textFn(t, s, e);
    }
  });
  return done.length + ' paragraph(s).';
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

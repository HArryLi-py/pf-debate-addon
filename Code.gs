/**
 * PF Debate Add-on — Phase 1 (revised v2): three formatting buttons +
 * generate evidence cards into a separate "PF Evidence Cards" doc, with a
 * cross-doc hyperlink back to each card.
 *
 * Three sidebar buttons:
 *   1. 标题 (Title)  — selection → centered 27pt bold underline + HEADING1
 *      (so it shows in the Doc's native Outline panel).
 *   2. blocks       — selection → gray 24pt underline.
 *   3. 生成 evidence card — from the form (title/citation/quote/highlight):
 *      append a card to a separate Evidence-Cards doc (created on first use,
 *      ID remembered), and insert a hyperlink at the cursor in the debate doc
 *      that opens that card's doc.
 *
 * Card layout (in the Evidence-Cards doc):
 *   - Title:     24pt, bold, underline (heading → Outline)
 *   - Citation:  12pt, not bold
 *   - Quote:     9pt (small); the highlighted excerpt (bold + underline + yellow
 *                background) is found within the quote text.
 *   - blank line before each new card.
 *
 * All via DocumentApp — no Advanced Docs Service needed (no more table/border
 * batchUpdate, no serviceId "找不到服务" issue). Scopes: documents,
 * script.container.ui, drive.file (drive.file for creating the Evidence doc).
 */

/* ============================== MENU + SIDEBAR ============================== */

function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
    .addItem('Open sidebar', 'showSidebar')
    .addSeparator()
    .addItem('Title selection (27pt centered bold underline)', 'applyTitle')
    .addItem('Block selection (24pt gray underline)', 'applyBlocks')
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
   Whole paragraph(s) containing the selection → HEADING1 + centered +
   27pt + bold + underline. Heading first, then alignment, then text-run
   overrides, so the heading style doesn't clobber center/bold/size. */
function applyTitle() {
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
      p.setHeading(DocumentApp.ParagraphHeading.HEADING1);
      p.setAlignment(DocumentApp.Alignment.CENTER);
      var t = p.editAsText();
      if (t) t.setFontSize(27).setBold(true).setUnderline(true);
    } else if (el.getType() === DocumentApp.ElementType.TEXT) {
      // fallback when selection isn't in a plain paragraph (list/table): style the range
      var t2 = el.asText();
      var s = re.isPartial() ? re.getStartOffset() : 0;
      var e = re.isPartial() ? re.getEndOffsetInclusive() : t2.getText().length - 1;
      if (e >= s) t2.setFontSize(s, e, 27).setBold(s, e, true).setUnderline(s, e, true);
    }
  });
  return done.length + ' paragraph(s) titled.';
}

/* ====================== BUTTON 2: BLOCKS (选中→灰块) =====================
   Gray + 24pt + underline. */
function applyBlocks() {
  return _withSelection(function (t, s, e, para) {
    t.setFontSize(s, e, 24).setUnderline(s, e, true).setForegroundColor(s, e, '#666666');
  });
}

/** Apply fn(textRun, startOffset, endOffsetInclusive, paragraphOrNull) to each
 *  text range in the current selection. */
function _withSelection(fn) {
  var doc = DocumentApp.getActiveDocument();
  var sel = doc.getSelection();
  if (!sel) throw new Error('先在 Doc 里选中文字。');
  var n = 0;
  sel.getRangeElements().forEach(function (re) {
    var el = re.getElement();
    if (el.getType() !== DocumentApp.ElementType.TEXT) return;
    var t = el.asText();
    var s = re.isPartial() ? re.getStartOffset() : 0;
    var e = re.isPartial() ? re.getEndOffsetInclusive() : t.getText().length - 1;
    if (e < s) return;
    var parent = el.getParent();
    var para = (parent && parent.getType() === DocumentApp.ElementType.PARAGRAPH)
      ? parent.asParagraph() : null;
    fn(t, s, e, para);
    n++;
  });
  return n + ' run(s) styled.';
}

/* ================ BUTTON 3: GENERATE EVIDENCE CARD (生成卡片) ==============
   Append a card to the shared "PF Evidence Cards" doc + insert a hyperlink
   at the cursor in the debate doc that opens the Evidence doc. */

var EV_DOC_ID_KEY = 'pf_evcards_doc_id';

function generateCard(form) {
  if (!form || !form.title || !form.citation || !form.quote) {
    throw new Error('标题、citation、quote 都要填。');
  }

  // 1. get/create the Evidence-Cards doc (separate Google Doc).
  var evDoc = _getOrCreateEvCardDoc();
  var evUrl = evDoc.getUrl();
  var body = evDoc.getBody();

  // 2. blank-line separator before each card.
  body.appendParagraph('');

  // 3. Title: 24pt bold underline + heading (→ Outline; also a navigable anchor).
  var titlePara = body.appendParagraph(form.title);
  titlePara.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  titlePara.editAsText().setFontSize(24).setBold(true).setUnderline(true);

  // 4. Citation: 12pt, not bold.
  body.appendParagraph(form.citation).editAsText().setFontSize(12);

  // 5. Quote: 9pt small; bold+underline+yellow the highlighted excerpt (found within).
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
  evDoc.saveAndClose();

  // 6. Insert a hyperlink at the cursor in the (debate) doc → opens the Evidence doc.
  var bound = DocumentApp.getActiveDocument();
  _insertCardLink(bound, form.title, evUrl);

  return 'Card "' + form.title + '" added to the Evidence doc.' +
    (form.highlight && !highlightFound ? ' (注意：highlight 文字在 quote 里没找到，没加高亮。)' : '') +
    ' 链接已插到光标处。';
}

/** Get the shared Evidence-Cards doc by stored ID; create it on first use. */
function _getOrCreateEvCardDoc() {
  var props = PropertiesService.getUserProperties();
  var id = props.getProperty(EV_DOC_ID_KEY);
  if (id) {
    try { return DocumentApp.openById(id); } catch (e) { /* recreate below */ }
  }
  var doc = DocumentApp.create('PF Evidence Cards');
  props.setProperty(EV_DOC_ID_KEY, doc.getId());
  return doc;
}

/** Insert `text` as a hyperlink (→ url) at the cursor in boundDoc; fall back to append. */
function _insertCardLink(boundDoc, text, url) {
  var cursor = boundDoc.getCursor();
  if (cursor) {
    try {
      var el = cursor.getElement();
      var start = cursor.getOffset();
      cursor.insertText(text);
      el.asText().setLinkUrl(start, start + text.length - 1, url);
      return;
    } catch (e) { /* fall through to append */ }
  }
  var p = boundDoc.getBody().appendParagraph(text);
  p.editAsText().setLinkUrl(url);
}

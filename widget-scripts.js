/* ==========================================================================
   WIDGET SCRIPTS — optional, per-report JS, sibling to widgets.css
   ==========================================================================
   widgets.css catalogs CSS-only components. These widgets are JS-heavy
   enough that they didn't fit that file's format — they live here
   instead, using the same "copy the labeled block you need" pattern.

   Nothing in this file is linked automatically by boilerplate.html. Copy
   the section you want into your dashboard's own <script>, following the
   per-section notes on what's safe to reuse verbatim vs. what needs
   rewriting for your report's own metric/brand vocabulary.

   Sections in this file:
     A. PPTX EXPORT ENGINE            — fully reusable, generic contract
        (single-source: UK_Brand_Pulse)
     B. EXCEL/CSV INGESTION HELPERS   — fully reusable date/value parsing
        (single-source: UK_Brand_Pulse)
     C. AI DATA ASSISTANT             — reusable shell, bespoke domain logic
        (single-source: UK_Brand_Pulse)
     D. COMPARE TOGGLE (MoM/Mo3M/YoY) — fully reusable lag math + contract
        (single-source: 202607_Product_Pulse)
     E. BRAND-SPLIT OVERLAY CHART     — reusable dataset builder, bespoke
        data-lookup glue (single-source: 202607_Product_Pulse)
     F. DATA DICTIONARY PAGE (JS)     — reusable render pattern, bespoke
        content (single-source: 202607_Product_Pulse; CSS lives in
        widgets.css under "DATA DICTIONARY PAGE")
     G. CHART EMBED AS RAW HTML       — NOT sourced from any dashboard; no
        source dashboard had this, so it's built fresh rather than
        extracted — read its note before assuming it's battle-tested
   ========================================================================== */


/* ==========================================================================
   A. PPTX EXPORT ENGINE  (verbatim, fully reusable — single-source)
   ==========================================================================
   Builds a real .pptx file (an OOXML zip) entirely in the browser: no
   server, no third-party export API. The chart that lands on the slide is
   a NATIVE PowerPoint chart object (not a pasted image) — recolour it,
   retype the title, switch chart type, and "Edit Data" opens a genuine
   embedded workbook (via SheetJS) with the real numbers in it.

   DEPENDENCIES — add both, only if you're using this widget:
     <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
   and the "fflate" zip library. UK_Brand_Pulse inlines fflate directly in
   a <script id="fflate-inline"> tag rather than pulling from a CDN — copy
   that inline block verbatim (it's the umd build of https://github.com/101arrowz/fflate,
   unmodified) so the exporter still works from a standalone copy of the
   dashboard with no network access at export time.

   WIRING — two things call into this engine:
     1. A button (id="chart-modal-ppt" in the source; reuses the existing
        .chart-modal-download class from widgets.css, no new CSS needed)
        inside your chart modal, shown only when the current chart's spec
        has a pptx() function:
          document.getElementById('chart-modal-ppt').style.display =
            (spec.pptx && !blocked) ? '' : 'none';
        ...and wired to call downloadModalPptx() on click.
     2. Each chart TYPE that wants PPT export adds a `pptx: () => {...}`
        function to its spec object (alongside whatever `config`/`csv` it
        already returns), returning:
          {
            title:      string,               // slide title
            subtitle:   string,                // slide subtitle line
            footnote:   string,                // small grey footer line
            seriesName: string,                // legend / column-2 header
            color:      '#RRGGBB',             // area fill + line colour
            rows: [ { month: 'Jul-26', value: 42.3, base: 1024 }, ... ]
                                                // value is a PERCENTAGE
                                                // NUMBER (42.3, not 0.423);
                                                // value may be null for a
                                                // gap; base is respondent n
                                                // for that month, used only
                                                // in the footnote base-range
          }
     MODAL_SPEC (whatever object your chart-modal already tracks as "the
     currently open chart's spec") is what downloadModalPptx() reads
     spec.pptx() from — adjust that one reference if your modal code names
     it something else.

   Only produces a single-series AREA chart today (that's what the source
   report needed). Extending pptChartXml() to line/bar/multi-series is a
   templating exercise in the same OOXML chart1.xml — not started here.

   Everything below is verbatim from source except three literal strings —
   'UK Brand Pulse' appears as the docProps creator/application name and as
   the theme name; rename those three for your own dashboard.
   ========================================================================== */

// POWERPOINT EXPORT — area chart as a NATIVE, EDITABLE chart
// Builds a .pptx (an OOXML zip) in the browser using the fflate
// build already inlined above, plus SheetJS for the embedded
// workbook that makes PowerPoint's "Edit Data" work. No new
// dependencies, so this still works from a standalone copy.
// The chart lands as a real chart object — recolour it, retype
// the title, switch it to a line/column, edit the numbers.
// ============================================================
const PPT_IN = 914400;                              // EMU per inch
const PPT_W = 12192000, PPT_H = 6858000;            // 16:9 slide
const PPT_MARGIN = Math.round(0.55 * PPT_IN);
const PPT_BODY_W = PPT_W - 2 * PPT_MARGIN;
const PPT_FONT = 'Calibri';                         // ships with Office on Win + Mac

function pptEsc(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&apos;')
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
}
function pptHex(c, fallback) {
  const m = String(c || '').replace('#', '').toUpperCase();
  return /^[0-9A-F]{6}$/.test(m) ? m : (fallback || '15B5DB');
}
const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

// ---------- text box helper ----------
function pptTextBox(id, name, x, y, cx, cy, runs, opts) {
  opts = opts || {};
  const body = runs.map(r =>
    '<a:r><a:rPr lang="en-GB" sz="' + r.sz + '"' + (r.b ? ' b="1"' : '') + ' dirty="0">'
    + '<a:solidFill><a:srgbClr val="' + r.color + '"/></a:solidFill>'
    + '<a:latin typeface="' + PPT_FONT + '"/><a:cs typeface="' + PPT_FONT + '"/></a:rPr>'
    + '<a:t>' + pptEsc(r.text) + '</a:t></a:r>').join('');
  return '<p:sp><p:nvSpPr><p:cNvPr id="' + id + '" name="' + pptEsc(name) + '"/>'
    + '<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
    + '<p:spPr><a:xfrm><a:off x="' + x + '" y="' + y + '"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm>'
    + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
    + '<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="' + (opts.anchor || 't') + '"><a:noAutofit/></a:bodyPr>'
    + '<a:lstStyle/><a:p><a:pPr algn="' + (opts.algn || 'l') + '"/>' + body + '</a:p></p:txBody></p:sp>';
}

// ---------- chart1.xml : the native chart ----------
function pptChartXml(o) {
  const n = o.rows.length;
  const col = pptHex(o.color);
  const cats = o.rows.map((r, i) => '<c:pt idx="' + i + '"><c:v>' + pptEsc(r.month) + '</c:v></c:pt>').join('');
  const vals = o.rows.map((r, i) => r.value == null ? ''
    : '<c:pt idx="' + i + '"><c:v>' + r.value + '</c:v></c:pt>').join('');
  const lastRow = n + 1;
  // Keep the category axis readable: aim for ~12 labels however long the series is.
  const skip = Math.max(1, Math.ceil(n / 12));
  const axTxPr = '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000">'
    + '<a:solidFill><a:srgbClr val="4A5170"/></a:solidFill>'
    + '<a:latin typeface="' + PPT_FONT + '"/></a:defRPr></a:pPr><a:endParaRPr lang="en-GB"/></a:p></c:txPr>';

  return XML_HEAD
    + '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"'
    + ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="' + NS_R + '">'
    + '<c:date1904 val="0"/><c:lang val="en-GB"/><c:roundedCorners val="0"/>'
    + '<c:chart><c:autoTitleDeleted val="1"/>'
    + '<c:plotArea><c:layout/>'
    + '<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>'
    + '<c:ser><c:idx val="0"/><c:order val="0"/>'
    + '<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/>'
    + '<c:pt idx="0"><c:v>' + pptEsc(o.seriesName) + '</c:v></c:pt></c:strCache></c:strRef></c:tx>'
    + '<c:spPr><a:solidFill><a:srgbClr val="' + col + '"><a:alpha val="28000"/></a:srgbClr></a:solidFill>'
    + '<a:ln w="22225"><a:solidFill><a:srgbClr val="' + col + '"/></a:solidFill></a:ln></c:spPr>'
    + '<c:dLbls><c:delete val="1"/></c:dLbls>'
    + '<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$' + lastRow + '</c:f>'
    + '<c:strCache><c:ptCount val="' + n + '"/>' + cats + '</c:strCache></c:strRef></c:cat>'
    + '<c:val><c:numRef><c:f>Sheet1!$B$2:$B$' + lastRow + '</c:f>'
    + '<c:numCache><c:formatCode>0.0"%"</c:formatCode><c:ptCount val="' + n + '"/>' + vals + '</c:numCache></c:numRef></c:val>'
    + '</c:ser>'
    + '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/>'
    + '<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>'
    + '<c:axId val="111111111"/><c:axId val="222222222"/></c:areaChart>'
    // category axis
    + '<c:catAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
    + '<c:delete val="0"/><c:axPos val="b"/>'
    + '<c:numFmt formatCode="General" sourceLinked="0"/>'
    + '<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
    + '<c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="CBD0DC"/></a:solidFill></a:ln></c:spPr>'
    + axTxPr
    + '<c:crossAx val="222222222"/><c:crosses val="autoZero"/><c:auto val="1"/>'
    + '<c:lblAlgn val="ctr"/><c:lblOffset val="100"/>'
    + '<c:tickLblSkip val="' + skip + '"/><c:tickMarkSkip val="' + skip + '"/><c:noMultiLvlLbl val="1"/></c:catAx>'
    // value axis
    // Pin the floor at zero to match the dashboard (beginAtZero) — an auto-scaled
    // axis on a tracker exaggerates month-to-month movement. Amendable in PowerPoint.
    + '<c:valAx><c:axId val="222222222"/>'
    + '<c:scaling><c:orientation val="minMax"/><c:min val="0"/></c:scaling>'
    + '<c:delete val="0"/><c:axPos val="l"/>'
    + '<c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="EEF0F5"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>'
    + '<c:numFmt formatCode="0&quot;%&quot;" sourceLinked="0"/>'
    + '<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
    + '<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>'
    + axTxPr
    + '<c:crossAx val="111111111"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>'
    + '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'
    + '</c:plotArea><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>'
    + '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'
    + '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr><a:latin typeface="' + PPT_FONT + '"/></a:defRPr></a:pPr><a:endParaRPr lang="en-GB"/></a:p></c:txPr>'
    + (o.hasWorkbook ? '<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>' : '')
    + '</c:chartSpace>';
}

// ---------- slide1.xml ----------
function pptSlideXml(o) {
  const titleY = Math.round(0.42 * PPT_IN);
  const subY = Math.round(1.03 * PPT_IN);
  const ruleY = Math.round(1.46 * PPT_IN);
  const chartY = Math.round(1.64 * PPT_IN);
  const chartH = Math.round(4.62 * PPT_IN);
  const footY = Math.round(6.44 * PPT_IN);

  const rule = '<p:sp><p:nvSpPr><p:cNvPr id="4" name="Rule"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
    + '<p:spPr><a:xfrm><a:off x="' + PPT_MARGIN + '" y="' + ruleY + '"/>'
    + '<a:ext cx="' + PPT_BODY_W + '" cy="12700"/></a:xfrm>'
    + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
    + '<a:solidFill><a:srgbClr val="E5E7ED"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr>'
    + '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-GB"/></a:p></p:txBody></p:sp>';

  const frame = '<p:graphicFrame><p:nvGraphicFramePr>'
    + '<p:cNvPr id="5" name="' + pptEsc(o.title) + ' chart"/>'
    + '<p:cNvGraphicFramePr><a:graphicFrameLocks/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>'
    + '<p:xfrm><a:off x="' + PPT_MARGIN + '" y="' + chartY + '"/>'
    + '<a:ext cx="' + PPT_BODY_W + '" cy="' + chartH + '"/></p:xfrm>'
    + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
    + '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="' + NS_R + '" r:id="rId2"/>'
    + '</a:graphicData></a:graphic></p:graphicFrame>';

  return XML_HEAD
    + '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="' + NS_R + '"'
    + ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
    + '<p:cSld><p:spTree>'
    + '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    + '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
    + '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
    + pptTextBox(2, 'Title', PPT_MARGIN, titleY, PPT_BODY_W, Math.round(0.55 * PPT_IN),
        [{ text: o.title, sz: 2400, b: true, color: '0F1F4A' }])
    + pptTextBox(3, 'Subtitle', PPT_MARGIN, subY, PPT_BODY_W, Math.round(0.34 * PPT_IN),
        [{ text: o.subtitle, sz: 1200, color: '4A5170' }])
    + rule
    + frame
    + pptTextBox(6, 'Footnote', PPT_MARGIN, footY, PPT_BODY_W, Math.round(0.4 * PPT_IN),
        [{ text: o.footnote, sz: 900, color: '7B829B' }])
    + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';
}

// ---------- embedded workbook (drives PowerPoint's "Edit Data") ----------
function pptWorkbook(o) {
  if (typeof XLSX === 'undefined') return null;
  try {
    const aoa = [['Month', o.seriesName, 'Base (n)']];
    o.rows.forEach(r => aoa.push([r.month, r.value, r.base]));
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [{ wch: 12 }, { wch: 22 }, { wch: 10 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    return new Uint8Array(XLSX.write(wb, { bookType: 'xlsx', type: 'array' }));
  } catch (e) { return null; }
}

// ---------- static package parts ----------
function pptStaticParts(hasWorkbook) {
  const rel = (id, type, target) =>
    '<Relationship Id="' + id + '" Type="' + NS_R + '/' + type + '" Target="' + target + '"/>';
  const P = {};

  P['[Content_Types].xml'] = XML_HEAD
    + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + (hasWorkbook ? '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>' : '')
    + '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
    + '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
    + '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
    + '<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
    + '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
    + '<Override PartName="/ppt/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
    + '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
    + '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
    + '</Types>';

  P['_rels/.rels'] = XML_HEAD
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + rel('rId1', 'officeDocument', 'ppt/presentation.xml')
    + '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
    + rel('rId3', 'extended-properties', 'docProps/app.xml')
    + '</Relationships>';

  P['ppt/presentation.xml'] = XML_HEAD
    + '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="' + NS_R + '"'
    + ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1">'
    + '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
    + '<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>'
    + '<p:sldSz cx="' + PPT_W + '" cy="' + PPT_H + '"/>'
    + '<p:notesSz cx="' + PPT_H + '" cy="' + PPT_W + '"/>'
    + '</p:presentation>';

  P['ppt/_rels/presentation.xml.rels'] = XML_HEAD
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + rel('rId1', 'slideMaster', 'slideMasters/slideMaster1.xml')
    + rel('rId2', 'slide', 'slides/slide1.xml')
    + rel('rId3', 'theme', 'theme/theme1.xml')
    + '</Relationships>';

  const emptyTree = '<p:spTree>'
    + '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    + '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
    + '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree>';

  P['ppt/slideMasters/slideMaster1.xml'] = XML_HEAD
    + '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="' + NS_R + '"'
    + ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
    + '<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
    + '<a:effectLst/></p:bgPr></p:bg>' + emptyTree + '</p:cSld>'
    + '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2"'
    + ' accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6"'
    + ' hlink="hlink" folHlink="folHlink"/>'
    + '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>'
    + '<p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="2400"/></a:lvl1pPr></p:titleStyle>'
    + '<p:bodyStyle><a:lvl1pPr><a:defRPr sz="1400"/></a:lvl1pPr></p:bodyStyle>'
    + '<p:otherStyle><a:lvl1pPr><a:defRPr sz="1400"/></a:lvl1pPr></p:otherStyle></p:txStyles>'
    + '</p:sldMaster>';

  P['ppt/slideMasters/_rels/slideMaster1.xml.rels'] = XML_HEAD
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + rel('rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml')
    + rel('rId2', 'theme', '../theme/theme1.xml')
    + '</Relationships>';

  P['ppt/slideLayouts/slideLayout1.xml'] = XML_HEAD
    + '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="' + NS_R + '"'
    + ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">'
    + '<p:cSld name="Blank">' + emptyTree + '</p:cSld>'
    + '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>';

  P['ppt/slideLayouts/_rels/slideLayout1.xml.rels'] = XML_HEAD
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + rel('rId1', 'slideMaster', '../slideMasters/slideMaster1.xml')
    + '</Relationships>';

  P['ppt/slides/_rels/slide1.xml.rels'] = XML_HEAD
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + rel('rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml')
    + rel('rId2', 'chart', '../charts/chart1.xml')
    + '</Relationships>';

  P['ppt/charts/_rels/chart1.xml.rels'] = XML_HEAD
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + (hasWorkbook ? '<Relationship Id="rId1" Type="' + NS_R + '/package" Target="../embeddings/data1.xlsx"/>' : '')
    + '</Relationships>';

  const themeFont = '<a:latin typeface="' + PPT_FONT + '"/><a:ea typeface=""/><a:cs typeface=""/>';
  const fill = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
  P['ppt/theme/theme1.xml'] = XML_HEAD
    + '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="UK Brand Pulse">'
    + '<a:themeElements>'
    + '<a:clrScheme name="UK Brand Pulse">'
    + '<a:dk1><a:srgbClr val="1A2238"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>'
    + '<a:dk2><a:srgbClr val="0F1F4A"/></a:dk2><a:lt2><a:srgbClr val="F7F8FB"/></a:lt2>'
    + '<a:accent1><a:srgbClr val="15B5DB"/></a:accent1><a:accent2><a:srgbClr val="EE3A7C"/></a:accent2>'
    + '<a:accent3><a:srgbClr val="0F1F4A"/></a:accent3><a:accent4><a:srgbClr val="4A5170"/></a:accent4>'
    + '<a:accent5><a:srgbClr val="7B829B"/></a:accent5><a:accent6><a:srgbClr val="CBD0DC"/></a:accent6>'
    + '<a:hlink><a:srgbClr val="15B5DB"/></a:hlink><a:folHlink><a:srgbClr val="EE3A7C"/></a:folHlink>'
    + '</a:clrScheme>'
    + '<a:fontScheme name="UK Brand Pulse">'
    + '<a:majorFont>' + themeFont + '</a:majorFont>'
    + '<a:minorFont>' + themeFont + '</a:minorFont>'
    + '</a:fontScheme>'
    + '<a:fmtScheme name="UK Brand Pulse">'
    + '<a:fillStyleLst>' + fill + fill + fill + '</a:fillStyleLst>'
    + '<a:lnStyleLst>'
    + '<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">' + fill + '<a:prstDash val="solid"/></a:ln>'
    + '<a:ln w="19050" cap="flat" cmpd="sng" algn="ctr">' + fill + '<a:prstDash val="solid"/></a:ln>'
    + '<a:ln w="28575" cap="flat" cmpd="sng" algn="ctr">' + fill + '<a:prstDash val="solid"/></a:ln>'
    + '</a:lnStyleLst>'
    + '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle>'
    + '<a:effectStyle><a:effectLst/></a:effectStyle>'
    + '<a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
    + '<a:bgFillStyleLst>' + fill + fill + fill + '</a:bgFillStyleLst>'
    + '</a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>';

  return P;
}

// ---------- package builder ----------
function buildPptx(o) {
  const wbBytes = pptWorkbook(o);
  o.hasWorkbook = !!wbBytes;
  const parts = pptStaticParts(o.hasWorkbook);
  parts['ppt/slides/slide1.xml'] = pptSlideXml(o);
  parts['ppt/charts/chart1.xml'] = pptChartXml(o);

  const stamp = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  parts['docProps/core.xml'] = XML_HEAD
    + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"'
    + ' xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"'
    + ' xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
    + '<dc:title>' + pptEsc(o.title) + '</dc:title>'
    + '<dc:subject>' + pptEsc(o.subtitle) + '</dc:subject>'
    + '<dc:creator>UK Brand Pulse</dc:creator><cp:lastModifiedBy>UK Brand Pulse</cp:lastModifiedBy>'
    + '<dcterms:created xsi:type="dcterms:W3CDTF">' + stamp + '</dcterms:created>'
    + '<dcterms:modified xsi:type="dcterms:W3CDTF">' + stamp + '</dcterms:modified>'
    + '</cp:coreProperties>';
  parts['docProps/app.xml'] = XML_HEAD
    + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"'
    + ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
    + '<Application>UK Brand Pulse</Application><Slides>1</Slides></Properties>';

  const zipInput = {};
  Object.keys(parts).forEach(k => { zipInput[k] = fflate.strToU8(parts[k]); });
  if (wbBytes) zipInput['ppt/embeddings/data1.xlsx'] = [wbBytes, { level: 0 }];
  return fflate.zipSync(zipInput, { level: 6, mtime: new Date() });
}

// ---------- wire-up ----------
function downloadModalPptx() {
  if (!MODAL_SPEC || !MODAL_SPEC.pptx) return;
  try {
    const o = MODAL_SPEC.pptx();
    if (!o.rows.length) { showToast('Nothing to export for this chart.', 3000); return; }
    const bytes = buildPptx(o);
    const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = (MODAL_SPEC.csvName || 'chart').replace(/[^a-z0-9]+/gi, '_') + '.pptx';
    a.click();
    URL.revokeObjectURL(a.href);
    showToast(o.hasWorkbook
      ? 'Downloaded PowerPoint \u00b7 editable chart (right-click \u2192 Edit Data)'
      : 'Downloaded PowerPoint \u00b7 chart is editable, but the data sheet could not be embedded', 4200);
  } catch (e) {
    console.error(e);
    showToast('Could not build the PowerPoint file.', 3200);
  }
}
// downloadModalPptx() calls showToast(...) — a small transient-message helper
// the source dashboard already has elsewhere for its other export buttons;
// wire this to whatever your dashboard uses for toast/snackbar messages, or
// stub it to console.log if you don't have one yet.


/* ==========================================================================
   B. EXCEL/CSV INGESTION HELPERS  (verbatim, fully reusable — single-source)
   ==========================================================================
   Three small, dependency-light pieces for any dashboard that ingests a
   wide monthly-tracker file (Excel or CSV) and needs to cope with however
   a real analyst's export actually formats dates and percentages, rather
   than assuming one clean format:

     normMonth(cell)  — normalises ANY of: a JS Date, an Excel date serial
                         number, "2026-05", "05/2026", "May 2026"/"May-26",
                         or "202605" → canonical "YYYY-MM" string, or null
                         if nothing matched.
     parseValue(cell)  — normalises a percentage cell that might arrive as
                         "42%", "42", "0.42", or "4,200" → a 0–1 fraction,
                         or null. (Uses `clamp(v,0,1)` — bring your own
                         clamp(v,lo,hi) helper, or inline it.)
     parseAnyFile(file) — routes an uploaded File to an Excel or CSV parser
                         by extension and returns a Promise<boolean>.

   parseAnyFile expects two functions you write per-dashboard:
     parseWorkbookFile(file) — for .xlsx/.xlsm/.xlsb/.xls, shown below as a
                                skeleton: read the file, hand the workbook
                                to your own onWorkbook(wb, filename), which
                                knows which sheet/columns your report's
                                template uses. (Requires the SheetJS script
                                tag noted in section A.)
     parseCsvFile(file)      — your existing CSV path (PapaParse or similar)
                                — not included here, use what you already have.
   ========================================================================== */

// Normalise a month cell to canonical "YYYY-MM". Accepts the template's text
// dates plus the variants a real export tends to carry: Excel date serials
// (numbers, when SheetJS reads with raw:true), JS Date objects, "YYYY/MM",
// "MM/YYYY", and month names ("May 2026", "May-26", "September 2025").
const MONTH_NAMES = { jan:1, january:1, feb:2, february:2, mar:3, march:3, apr:4, april:4, may:5, jun:6, june:6, jul:7, july:7, aug:8, august:8, sep:9, sept:9, september:9, oct:10, october:10, nov:11, november:11, dec:12, december:12 };
function ymOut(y, mo) { if (!(y >= 1990 && y <= 2100) || !(mo >= 1 && mo <= 12)) return null; return y + '-' + String(mo).padStart(2, '0'); }
function normMonth(s) {
  if (s === null || s === undefined || s === '') return null;
  // JS Date object (e.g. when a workbook is read with cellDates).
  if (s instanceof Date && !isNaN(s.getTime())) return ymOut(s.getUTCFullYear(), s.getUTCMonth() + 1);
  // Excel serial date number (days since 1899-12-30). Guard to a plausible range.
  if (typeof s === 'number' && isFinite(s)) {
    if (s > 20000 && s < 80000) { const d = new Date(Date.UTC(1899, 11, 30) + Math.round(s) * 86400000); if (!isNaN(d.getTime())) return ymOut(d.getUTCFullYear(), d.getUTCMonth() + 1); }
    return null;
  }
  s = String(s).trim();
  let m = s.match(/^(\d{4})[-/.](\d{1,2})(?:[-/.]\d{1,2})?$/); if (m) return ymOut(+m[1], +m[2]);   // 2026-05, 2026/5, 2026-05-12
  m = s.match(/^(\d{1,2})[-/.](\d{4})$/); if (m) return ymOut(+m[2], +m[1]);                          // 05/2026
  m = s.match(/^([A-Za-z]{3,9})[ \-/.]*'?(\d{2,4})$/);                                                 // May 2026, May-26, September 2025
  if (m) { const mo = MONTH_NAMES[m[1].toLowerCase()]; if (mo) { let y = +m[2]; if (y < 100) y += 2000; return ymOut(y, mo); } }
  m = s.match(/^(\d{4})(\d{2})$/); if (m) return ymOut(+m[1], +m[2]);                                  // 202605
  const d = new Date(s); if (!isNaN(d.getTime())) return ymOut(d.getFullYear(), d.getMonth() + 1);
  return null;
}
function parseValue(raw) { if (raw === null || raw === undefined || raw === '') return null; let v = parseFloat(String(raw).replace('%', '').replace(/,/g, '').trim()); if (isNaN(v)) return null; if (v > 1.5) v /= 100; return clamp(v, 0, 1); }

// Route an uploaded file to the Excel or CSV parser by extension. Returns a Promise
// that resolves once the file has been parsed and merged (or failed).
function parseAnyFile(file) {
  if (!file) return Promise.resolve(false);
  const name = (file.name || '').toLowerCase();
  if (/\.(xlsx|xlsm|xlsb|xls)$/.test(name)) return parseWorkbookFile(file);
  return parseCsvFile(file);
}

// Skeleton for the Excel side of parseAnyFile — read the file, hand the parsed
// workbook to your own onWorkbook(wb, filename) which knows your report's sheet
// layout (uploadError/setDzStatus are the dropzone-widget status helpers from
// widgets.css's Multi-file upload section — reuse those, or your own equivalent).
function parseWorkbookFile(file) {
  return new Promise(resolve => {
    if (typeof XLSX === 'undefined') { uploadError('Excel reader failed to load — check your connection and reload, or upload the CSV version instead.'); resolve(false); return; }
    setDzStatus('Reading ' + file.name + ' …', 'loading');
    const reader = new FileReader();
    reader.onload = e => { try { const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' }); onWorkbook(wb, file.name); resolve(true); } catch (err) { uploadError('Could not read workbook: ' + (err && err.message ? err.message : err)); resolve(false); } };
    reader.onerror = () => { uploadError('Could not read file.'); resolve(false); };
    reader.readAsArrayBuffer(file);
  });
}


/* ==========================================================================
   C. AI DATA ASSISTANT — full widget  (single-source; CSS lives in
   widgets.css under "AI DATA ASSISTANT — panel shell"; markup + JS below)
   ==========================================================================
   A floating "Ask the data" launcher that opens a chat panel answering
   questions about the dashboard's own figures and methodology — entirely
   in-browser, no network call, no LLM. It works by pattern-matching the
   question against the dashboard's own brand/metric/date vocabulary and
   reading the same lookup functions the dashboard itself uses, then
   rendering a templated answer (figure + delta + significance + a small
   sparkline where relevant). It can also deep-link: an "Show me" chip on
   an answer can jump the user to the matching dashboard view.

   REUSE TIERS — be honest with yourself about which bucket each piece of
   code you copy falls into, because the middle tier is the one people most
   often accidentally treat as tier 1:

   1. VERBATIM REUSABLE (no report-specific vocabulary in these at all):
      collapse(), words(), esc2() — string helpers
      loadDiag()/saveDiag()/logDiag() and buildDiagPanel()/renderDiagList()/
        openDiagPanel()/closeDiagPanel()/toggleDiagPanel() — the "log
        questions I couldn't answer, let me review them later" diagnostics
        drawer, opened with Ctrl+Alt+D or ?diag=1. Fully generic.
      scrollDown(), addUser(), copyPlain(), addBot(), seed(), run(), send(),
        openPanel(), closePanel() — the chat-panel mechanics (render a
        message, scroll, wire the input box). Fully generic.
      The HTML markup below (launcher button + panel shell) — generic IDs
        only, copy as-is.

   2. ADAPT THE SHAPE, REWRITE THE VOCABULARY (the pattern is reusable, the
      brand/metric names and matching lists inside are this report's own —
      you WILL need to redo the content of these for your dashboard):
      detectMarket(), brandIndex(), brandsInText(), inferMarket(),
        resolveMetric(), resolveMonth(), resolveTwoMonths(), ym(), inData(),
        parseSubgroups(), audKeyFromParsed(), crossMarketPairs() — all of
        these hold this report's specific brand-name aliases (e.g. "lads" →
        "Ladbrokes"), metric labels, and vertical names. Same job, new list,
        for your dashboard.
      pointAnswer(), seriesAnswer(), rankAnswer(), compareAnswer(),
        crossCompareAnswer(), periodCompareAnswer(), followUpChips(),
        exampleChips() — the answer TEMPLATES (what HTML a "what's X's
        score" or "X vs Y" question renders). Reusable as a pattern, but
        the copy inside references this report's own metric names.
      askEngine() — the central router: mostly generic intent-detection
        plumbing, but it calls out to the tier-2 functions above and its
        keyword lists (methoHit regex etc.) are tuned to this report's
        own terminology.
      applyToDashboard() — the "jump to dashboard" deep-link; touches your
        dashboard's own filter-state setters, so this needs rewiring to
        whatever your dashboard calls to change vertical/brand/date.

   3. FULLY BESPOKE (don't copy the TEXT, only the idea that this exists):
      methodologyAnswer() and definitionAnswer() return hardcoded prose
        describing THIS report's significance test and metric definitions.
        Write your own description of your own methodology.

   Everything below is otherwise verbatim from source, included in full
   (all three tiers) because even the tier-2/3 functions are a genuinely
   useful reference for the SHAPE of a working implementation — faster to
   adapt than to write from nothing.
   ========================================================================== */

<button class="assistant-launcher" id="a-launch" aria-label="Open the data assistant" aria-haspopup="dialog">
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"></path></svg>
  Ask the data
</button>

<div class="assistant-panel" id="a-panel" role="dialog" aria-modal="false" aria-label="Data and methodology assistant">
  <div class="a-head">
    <div style="flex:1"><div class="a-title">Data &amp; Methodology Assistant</div><div class="a-sub">Runs in your browser · nothing leaves this file</div></div>
    <button id="a-close" aria-label="Close assistant"><svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><line x1="6" y1="6" x2="18" y2="18"></line><line x1="18" y1="6" x2="6" y2="18"></line></svg></button>
  </div>
  <div class="a-body" id="a-body"></div>
  <div class="a-foot">
    <input id="a-input" type="text" placeholder="Ask about a figure or the methodology…" autocomplete="off" aria-label="Ask a question">
    <button id="a-send" aria-label="Send"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="22" y1="2" x2="11" y2="13"></line><polygon points="22 2 15 22 11 13 2 9 22 2"></polygon></svg></button>
  </div>
</div>

<script>
(function(){
  "use strict";
  // ---------- small text helpers ----------
  function collapse(s){return String(s==null?'':s).toLowerCase().replace(/[^a-z0-9]+/g,'');}
  function words(s){return String(s==null?'':s).toLowerCase().replace(/[^a-z0-9\s]+/g,' ').split(/\s+/).filter(Boolean);}
  function esc2(s){return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}

  // ---------- conversation memory (follow-up questions) ----------
  // Remembers the last successfully-resolved market/brand(s)/metric/date so a
  // follow-up like "and last month?" or "what about Paddy Power?" doesn't need
  // to repeat the whole question. Cleared on panel close is NOT done on purpose
  // — stakeholders expect the thread to keep context while the panel is open.
  var LAST=null;
  function rememberContext(mk,brands,metric,date){ LAST={mk:mk,brands:(brands||[]).slice(),metric:metric,date:date}; }

  // ---------- private diagnostics log (never shown to stakeholders) ----------
  // Logs questions the assistant couldn't resolve, or had to ask for
  // clarification on, so the dashboard owner can review them later and expand
  // the synonym/metric tables. Stored only in this browser's localStorage —
  // nothing is sent anywhere. Viewing the log requires a deliberate action
  // (see openDiagPanel below) that a stakeholder would not stumble into.
  var DIAG_KEY='bp_assistant_diag_log_v1';
  function loadDiag(){ try{ return JSON.parse(window.localStorage.getItem(DIAG_KEY)||'[]'); }catch(e){ return []; } }
  function saveDiag(list){ try{ window.localStorage.setItem(DIAG_KEY, JSON.stringify(list.slice(-300))); }catch(e){} }
  function logDiag(entry){
    var list=loadDiag();
    list.push({t:new Date().toISOString(), q:entry.q, reason:entry.reason});
    saveDiag(list);
  }

  // ---------- market resolution ----------
  function detectMarket(low){
    if(/\bbingo\b/.test(low))return 'Bingo';
    if(/\b(casino|gaming|vegas|slots?)\b/.test(low))return 'Casino';
    if(/\b(sports?|betting|bet)\b/.test(low))return 'Sports';
    return null;
  }

  // ---------- brand index (per market: collapsed key -> canonical) ----------
  var BRAND_SYN={
    Sports:{skybet:'SkyBet',paddy:'Paddy Power',pp:'Paddy Power',betfair:'Betfair',bet365:'Bet365',b365:'Bet365',coral:'Coral',ladbrokes:'Ladbrokes',lads:'Ladbrokes',williamhill:'William Hill',willhill:'William Hill',wh:'William Hill'},
    Casino:{skyvegas:'Sky Vegas',vegas:'Sky Vegas',skycasino:'Sky Casino',paddypowergames:'Paddy Power Games',ppgames:'Paddy Power Games',paddy:'Paddy Power Games',betfaircasino:'Betfair Casino',betfair:'Betfair Casino',tombolaarcade:'Tombola Arcade',tombola:'Tombola Arcade',bet365:'Bet365 Casino',b365:'Bet365 Casino',ladbrokes:'Ladbrokes Casino',williamhill:'William Hill Vegas',willhill:'William Hill Vegas','888':'888 Casino',virgin:'Virgin Games',virgingames:'Virgin Games',jackpotjoy:'Jackpot Joy'},
    Bingo:{skybingo:'Sky Bingo',paddy:'Paddy Power',pp:'Paddy Power',betfair:'Betfair',tombola:'Tombola',foxy:'Foxy Bingo',foxybingo:'Foxy Bingo',mecca:'Mecca Bingo',meccabingo:'Mecca Bingo',gala:'Gala Bingo',galabingo:'Gala Bingo',buzz:'Buzz Bingo',buzzbingo:'Buzz Bingo',jackpotjoy:'Jackpot Joy'}
  };
  function brandIndex(mk){
    var m={};
    brandsFor(mk).forEach(function(b){ m[collapse(b)]=b; var l=brandLabel(b); if(l)m[collapse(l)]=b; });
    var syn=BRAND_SYN[mk]||{}; Object.keys(syn).forEach(function(k){ if(m[k]===undefined) m[k]=syn[k]; });
    return m;
  }
  // find brands mentioned, resolved within a chosen market
  function brandsInText(low, mk){
    var idx=brandIndex(mk), toks=words(low), found=[], seen={};
    for(var i=0;i<toks.length;i++){
      for(var n=Math.min(3,toks.length-i);n>=1;n--){
        var key=collapse(toks.slice(i,i+n).join(' '));
        if(idx[key]!==undefined){ var b=idx[key]; if(!seen[b]){seen[b]=1;found.push(b);} i+=n-1; break; }
      }
    }
    return found;
  }
  // infer market when none stated: prefer the market with the most specific
  // (longest) brand-phrase match. If two+ markets tie on the same best length,
  // that's a genuine ambiguity (e.g. "Paddy" exists in Sports, Casino AND
  // Bingo) — the caller should ask rather than silently guess.
  var LAST_MARKET_TIES=[];
  function inferMarket(low, preferred){
    var order=[preferred||MARKET].concat(MARKETS.filter(function(m){return m!==(preferred||MARKET);}));
    var bestM=null,bestLen=0,ties=[],toks=words(low);
    order.forEach(function(mk){
      var idx=brandIndex(mk),hitLen=0;
      for(var i=0;i<toks.length;i++){
        for(var n=Math.min(3,toks.length-i);n>=1;n--){
          var key=collapse(toks.slice(i,i+n).join(' '));
          if(idx[key]!==undefined){ if(key.length>hitLen)hitLen=key.length; break; }
        }
      }
      if(hitLen>0){
        if(hitLen>bestLen){ bestLen=hitLen; bestM=mk; ties=[mk]; }
        else if(hitLen===bestLen){ ties.push(mk); }
      }
    });
    // A genuine ambiguity is when 2+ markets tie AND the caller's preferred
    // market isn't one of them (if it is, that preference breaks the tie).
    if(ties.length>1 && preferred && ties.indexOf(preferred)>=0){ bestM=preferred; ties=[]; }
    LAST_MARKET_TIES = ties.length>1 ? ties : [];
    return bestM;
  }

  // ---------- metric resolution ----------
  var HEADLINE_SYN=[
    [/most[\s-]?often|most[\s-]?used|mostly use|primary brand|main brand/,'Most Often Used'],
    [/used (in|last|it)|\busage\b|played with|\bused\b/,'Used in Last Month'],
    [/consider/,'Consideration (T2B)'],
    [/spont\w* ?ad|ad awareness|advertis\w* awareness|seen advertis/,'Spontaneous Ad Awareness - All Mentions'],
    [/total spont|spont\w* awareness|unprompted/,'Total Spontaneous Awareness'],
    [/prompt|aided|recogni/,'Prompted Awareness'],
    [/awareness/,'Prompted Awareness']
  ];
  var STOP={associate:1,'with':1,best:1,at:1,the:1,a:1,an:1,of:1,to:1,in:1,me:1,my:1,you:1,your:1,they:1,them:1,their:1,and:1,'for':1,is:1,are:1,on:1,it:1,that:1,this:1,brand:1,brands:1,i:1,we:1};
  var LAST_METRIC_ALTS=[];
  function resolveMetric(mk, low){
    LAST_METRIC_ALTS=[];
    for(var i=0;i<HEADLINE_SYN.length;i++){ if(HEADLINE_SYN[i][0].test(low) && metricsFor(mk).indexOf(HEADLINE_SYN[i][1])>=0) return HEADLINE_SYN[i][1]; }
    var mtoks={}; words(low).forEach(function(w){mtoks[w]=1;});
    var best=null,bestScore=0,second=null,secondScore=0;
    metricsFor(mk).forEach(function(metric){
      if(HEADLINE_METRICS.indexOf(metric)>=0) return;
      var mt=words(metric).filter(function(w){return !STOP[w] && w.length>=3;});
      if(!mt.length)return;
      var hit=0,distinct=0; mt.forEach(function(w){ if(mtoks[w]){hit++; if(w.length>=5)distinct++;} });
      if(!hit)return;
      var score=hit/mt.length + hit*0.05 + distinct*0.15;
      if(score>bestScore){ second=best;secondScore=bestScore; best=metric;bestScore=score; }
      else if(score>secondScore){ second=metric; secondScore=score; }
    });
    // Flag ambiguity only when two genuinely different, individually-plausible
    // metrics are within a hair of each other — not just because a second
    // metric scored a little lower.
    if(best && second && bestScore>=0.45 && secondScore>=0.45 && (bestScore-secondScore)<=0.08){
      LAST_METRIC_ALTS=[best,second];
    }
    return (best && bestScore>=0.45)?best:null;
  }

  // ---------- month resolution ----------
  var MO={jan:1,january:1,feb:2,february:2,mar:3,march:3,apr:4,april:4,may:5,jun:6,june:6,jul:7,july:7,aug:8,august:8,sep:9,sept:9,september:9,oct:10,october:10,nov:11,november:11,dec:12,december:12};
  function ym(y,mo){ if(mo<1||mo>12)return null; return y+'-'+String(mo).padStart(2,'0'); }
  function inData(d){ return d && DATES.indexOf(d)>=0 ? d : null; }
  function resolveMonth(low){
    if(/\b(latest|current|most recent|newest|this month|right now|now)\b/.test(low)) return {date:PERIOD};
    if(/\blast month\b|\bprevious month\b|\bmonth before\b/.test(low)){var d=monthsAgo(PERIOD,1);return {date:inData(d),req:d};}
    if(/\blast year\b|\ba year ago\b|\byear ago\b|\bsame month last year\b/.test(low)){var y=monthsAgo(PERIOD,12);return {date:inData(y),req:y};}
    var m=low.match(/\b(\d{4})[-\/.](\d{1,2})\b/); if(m){var a=ym(+m[1],+m[2]);return {date:inData(a),req:a};}
    m=low.match(/\b(\d{1,2})[-\/.](\d{4})\b/); if(m){var b=ym(+m[2],+m[1]);return {date:inData(b),req:b};}
    m=low.match(/\b([a-z]{3,9})[ '\-\/]*'?(\d{2,4})\b/); if(m&&MO[m[1]]){var yy=+m[2];if(yy<100)yy+=2000;var c=ym(yy,MO[m[1]]);return {date:inData(c),req:c};}
    var g=low.match(/\b([a-z]{3,9})\b/g);
    if(g){for(var i=0;i<g.length;i++){var w=g[i]; if(w==='may'||w==='march')continue; if(MO[w]){var mm=MO[w];var cand=DATES.filter(function(x){return +x.split('-')[1]===mm;}); if(cand.length)return {date:cand[cand.length-1]};}}}
    return {date:null};
  }
  // Find TWO distinct explicit months in one question, in the order they were
  // typed — powers "compare January vs June" style custom-period comparisons
  // (as opposed to the dashboard's fixed MoM/YoY setting).
  function resolveTwoMonths(low){
    var found=[], re=/\b(\d{4})[-\/.](\d{1,2})\b|\b(\d{1,2})[-\/.](\d{4})\b|\b([a-z]{3,9})[ '\-\/]*'?(\d{2,4})\b|\b([a-z]{3,9})\b/g, m;
    while((m=re.exec(low))){
      var d=null;
      if(m[1]&&m[2]) d=inData(ym(+m[1],+m[2]));
      else if(m[3]&&m[4]) d=inData(ym(+m[4],+m[3]));
      else if(m[5]&&MO[m[5]]){ var yy=+m[6]; if(yy<100)yy+=2000; d=inData(ym(yy,MO[m[5]])); }
      else if(m[7]&&m[7]!=='may'&&m[7]!=='march'&&MO[m[7]]){
        var mm=MO[m[7]], cand=DATES.filter(function(x){return +x.split('-')[1]===mm;}), pick=null;
        for(var i=cand.length-1;i>=0;i--){ if(found.every(function(f){return f.date!==cand[i];})){ pick=cand[i]; break; } }
        d=pick||(cand.length?cand[cand.length-1]:null);
      }
      if(d && found.every(function(f){return f.date!==d;})) found.push({idx:m.index,date:d});
    }
    found.sort(function(a,b){return a.idx-b.idx;});
    return found.length>=2 ? {d1:found[0].date,d2:found[1].date} : null;
  }

  // ---------- formatting ----------
  function arrow(dir){return dir>0?'<span class="a-up">▲</span>':(dir<0?'<span class="a-dn">▼</span>':'<span class="a-flat">▬</span>');}
  function ctxNote(){
    var src = (typeof DATA_SOURCE!=='undefined' && DATA_SOURCE==='uploaded') ? 'your uploaded data' : 'sample data';
    return 'Context: '+esc2(filterContextLabel())+' · '+compareLabel()+' · '+src;
  }
  function mLabel(mk,metric){ try{var l=comparisonMetricLabel(mk,metric); return l||metric;}catch(e){return metric;} }

  // ---------- methodology knowledge base ----------
  var KB=[
    {k:['significance','significant','confidence','z test','ztest','z-test','z score','1.96','two proportion','two-proportion','stat sig','statistically','how do you know','flagged'],
     a:function(){return 'Differences are flagged at the <strong>95% confidence level</strong> with a two-sided <strong>two-proportion z-test</strong> — a change is marked significant when |z| ≥ 1.96. Each brand is tested against its own comparison period (previous month for MoM, the same month a year earlier for YoY); in Brand Comparison it is also tested against the month\u2019s leading brand. '+arrow(1)+' means a significant rise, '+arrow(-1)+' a significant fall.';}},
    {k:['base','sample size','sample','denominator','respondents','how many people','n ='],
     a:function(){return 'Significance uses each month\u2019s survey base as the denominator. '+((typeof DATA_SOURCE!=='undefined'&&DATA_SOURCE==='uploaded')?'With your uploaded data the <strong>Base</strong> column from your file is used, per month and subgroup.':'With the built-in sample data the bases are roughly <strong>1,300</strong> (Sports), <strong>1,100</strong> (Casino) and <strong>760</strong> (Bingo) per month.')+' As a rule of thumb, a month-on-month move of about <strong>3–4 percentage points</strong> around a 30% score sits near the threshold of significance.';}},
    {k:['mom','yoy','month on month','month-on-month','year on year','year-on-year','comparison period','previous month','year earlier','compare period'],
     a:function(){return 'There are two comparison modes. <strong>MoM</strong> compares each month with the one before; <strong>YoY</strong> compares it with the same month a year earlier. The dashboard is currently set to <strong>'+compareLabel()+'</strong>, and my change figures and arrows follow that setting.';}},
    {k:['net','weighted','subgroup','combine','combined','aggregate','aggregated','how are subgroup'],
     a:function(){return 'Audiences are <strong>not combined</strong>. The <strong>Audience</strong> selector picks exactly one audience at a time and every figure is read only from that audience\u2019s own tab — there is no NET or averaging across audiences.';}},
    {k:['portfolio','audience view','market view','flutter customer','segment view'],
     a:function(){return 'The single <strong>Audience</strong> selector chooses one audience: the <strong>Total Market</strong>, a market split (gender, age band, frequency, player segment or an \u201COther\u201D split), or a <strong>Portfolio</strong> segment (Mass Next Gen / Mass 30+ / HVC, shown on their rolling windows). Each is read only from its own tab. Right now: <strong>'+esc2(filterContextLabel())+'</strong>.';}},
    {k:['dummy','synthetic','generated','real data','data source','where does the data','made up','fake'],
     a:function(){return (typeof DATA_SOURCE!=='undefined'&&DATA_SOURCE==='uploaded')?'The dashboard is showing <strong>your uploaded tracker data</strong>. I read every figure directly from those values — I never estimate or invent numbers.':'The dashboard is currently showing <strong>built-in sample data</strong> (deterministic dummy values for demonstration). Load your tracker export to replace it. Either way, I only ever read values already in the dashboard — I never invent them.';}},
    {k:['filter','filtered','current view','context','which cut'],
     a:function(){return 'By default a figure reflects the dashboard\u2019s current context: <strong>'+esc2(filterContextLabel())+'</strong>. You can also name an audience directly in your question — for example \u201Camong 18-29s\u201D, \u201Cfor women\u201D, \u201Ccore players\u201D or a player segment — and I\u2019ll apply just that cut for the answer, then leave your dashboard filters untouched. Subgroup figures only exist where your uploaded data carries that breakdown.';}}
  ];
  function methodologyAnswer(low){
    var best=null,bestScore=0;
    KB.forEach(function(e){ var s=0; e.k.forEach(function(k){ if(low.indexOf(k)>=0) s++; }); if(s>bestScore){bestScore=s;best=e;} });
    return bestScore>=1 ? (typeof best.a==='function'?best.a():best.a) : null;
  }
  var DEFS={
    'Spontaneous Ad Awareness - All Mentions':'Unprompted recall of having seen or heard advertising for the brand in the category.',
    'Total Spontaneous Awareness':'Unprompted recall of the brand (including its advertising), with duplicate mentions removed.',
    'Prompted Awareness':'Aided awareness — the brand is recognised from a shown list. The broadest awareness measure.',
    'Consideration (T2B)':'Top-2-Box consideration — the share who would consider the brand first, or among the brands they\u2019d consider ahead of others.',
    'Used in Last Month':'The share who report having used or played with the brand in the last month.',
    'Most Often Used':'The share who name the brand as the one they use most often in the category.'
  };
  function definitionAnswer(mk,metric){
    if(DEFS[metric]) return '<strong>'+esc2(metric)+'</strong> — '+DEFS[metric];
    if(/^Associate with:/.test(metric)) return '<strong>'+esc2(metric)+'</strong> — the share of people who associate the brand with \u201C'+esc2(metric.replace(/^Associate with:\s*/,''))+'\u201D.';
    if(/^Best at:/.test(metric)) return '<strong>'+esc2(metric)+'</strong> — the share who pick the brand as best at \u201C'+esc2(metric.replace(/^Best at:\s*/,''))+'\u201D.';
    return '<strong>'+esc2(metric)+'</strong> is one of the tracked attribute statements; the figure is the share of people who give the brand that response.';
  }

  // ---------- answer builders ----------
  function pointAnswer(mk,brand,metric,date){
    var v=valueAt(mk,brand,metric,date);
    if(v===null||v===undefined||isNaN(v)) return {html:'I don\u2019t have a value for <strong>'+esc2(brandLabel(brand))+' · '+esc2(mLabel(mk,metric))+'</strong> ('+mk+') in '+prettyDate(date)+'. It may not be covered for that month or subgroup.'};
    var cmp=compareDate(date), pv=valueAt(mk,brand,metric,cmp), sig=sigVsCompare(mk,brand,metric,date);
    var delta=(pv===null||pv===undefined||isNaN(pv))?null:v-pv;
    var line='<strong>'+esc2(brandLabel(brand))+'</strong> · '+esc2(mLabel(mk,metric))+' <span style="color:var(--ink-3)">('+mk+')</span><br>'+prettyDate(date)
      +'<span class="a-fig">'+fmtPct(v)+'</span>';
    if(delta!==null){ line+='vs '+compareLabel()+' '+fmtPct(pv)+' &nbsp;'+arrow(sig.dir)+' '+fmtPP(delta)+(sig.dir!==0?' (significant)':' (not significant)'); }
    line+='<div class="a-note">'+ctxNote()+'</div>';
    return {html:line};
  }
  function seriesAnswer(mk,brand,metric,date,n){
    var win=windowDates(date,n||12), vals=win.map(function(d){return valueAt(mk,brand,metric,d);});
    var valid=vals.filter(function(x){return x!==null&&!isNaN(x);});
    if(!valid.length) return {html:'I don\u2019t have a trend for <strong>'+esc2(brandLabel(brand))+' · '+esc2(mLabel(mk,metric))+'</strong> ('+mk+').'};
    var mx=Math.max.apply(null,valid), first=null,last=null;
    for(var i=0;i<vals.length;i++){ if(vals[i]!==null&&!isNaN(vals[i])){ if(first===null)first={d:win[i],v:vals[i]}; last={d:win[i],v:vals[i]}; } }
    var bars=vals.map(function(v,i){ var h=(v===null||isNaN(v))?0:Math.round((v/mx)*44); return '<div class="bar'+(i===vals.length-1?' last':'')+'" style="height:'+Math.max(2,h)+'px" title="'+shortDate(win[i])+': '+fmtPct(v)+'"></div>'; }).join('');
    var d=last.v-first.v;
    var html='<strong>'+esc2(brandLabel(brand))+'</strong> · '+esc2(mLabel(mk,metric))+' <span style="color:var(--ink-3)">('+mk+')</span><br><span style="font-size:11.5px;color:var(--ink-3)">'+shortDate(first.d)+' → '+shortDate(last.d)+'</span>'
      +'<div class="a-spark">'+bars+'</div>'
      +'<div class="a-sparklab"><span>'+fmtPct(first.v)+'</span><span>'+(d>=0?'+':'')+ (d*100).toFixed(1)+'pp</span><span>'+fmtPct(last.v)+'</span></div>'
      +'<div class="a-note">Trend of the '+valid.length+' most recent available months (not significance-tested). '+ctxNote()+'</div>';
    return {html:html};
  }
  function rankAnswer(mk,metric,date,asc){
    var rows=brandsFor(mk).map(function(b){return {b:b,v:valueAt(mk,b,metric,date),s:sigVsCompare(mk,b,metric,date)};}).filter(function(r){return r.v!==null&&!isNaN(r.v);});
    if(!rows.length) return {html:'I don\u2019t have values for '+esc2(mLabel(mk,metric))+' in '+mk+' for '+prettyDate(date)+'.'};
    rows.sort(function(a,z){return asc?a.v-z.v:z.v-a.v;});
    var t='<strong>'+esc2(mLabel(mk,metric))+'</strong> · '+mk+' · '+prettyDate(date)+'<table class="a-table"><tbody>';
    rows.forEach(function(r){ t+='<tr><td>'+esc2(brandLabel(r.b))+'</td><td class="v">'+fmtPct(r.v)+'</td><td class="d">'+arrow(r.s.dir)+'</td></tr>'; });
    t+='</tbody></table><div class="a-note">Ranked '+(asc?'lowest→highest':'highest→lowest')+'. Arrows = significance vs '+compareLabel()+'. '+ctxNote()+'</div>';
    return {html:t};
  }
  function compareAnswer(mk,brands,metric,date){
    var rows=brands.map(function(b){return {b:b,v:valueAt(mk,b,metric,date)};}).filter(function(r){return r.v!==null&&!isNaN(r.v);});
    if(rows.length<2) return {html:'I need two brands with data to compare. Try naming both, e.g. \u201CSky Bet vs Paddy Power consideration\u201D.'};
    rows.sort(function(a,z){return z.v-a.v;});
    var base=baseAt(mk,date), lead=rows[0];
    var t='<strong>'+esc2(mLabel(mk,metric))+'</strong> · '+mk+' · '+prettyDate(date)+'<table class="a-table"><tbody>';
    rows.forEach(function(r){
      var gapSig=''; if(r.b!==lead.b){ var st=sigTest(lead.v,base,r.v,base); gapSig=st.dir>0?' '+arrow(1):''; }
      t+='<tr><td>'+esc2(brandLabel(r.b))+(r.b===lead.b?' <span style="font-size:10px;color:var(--ink-3)">(lead)</span>':'')+'</td><td class="v">'+fmtPct(r.v)+'</td><td class="d">'+(r.b===lead.b?'':fmtPP(r.v-lead.v)+gapSig)+'</td></tr>';
    });
    t+='</tbody></table><div class="a-note">Gap vs leader shown in pp. '+arrow(1)+' = leader significantly ahead of that brand. '+ctxNote()+'</div>';
    return {html:t, plain: rows.map(function(r){return brandLabel(r.b)+': '+fmtPct(r.v);}).join(', ')};
  }
  // Compares brands that live in DIFFERENT markets (e.g. "Bet365 vs Bet365
  // Casino consideration") — each brand's figure is read from its own
  // market's tracker/base, so significance uses each side's own base size.
  function crossCompareAnswer(pairs,metric,date){
    var rows=pairs.map(function(p){return {mk:p.mk,b:p.brand,v:valueAt(p.mk,p.brand,metric,date)};}).filter(function(r){return r.v!==null&&!isNaN(r.v);});
    if(rows.length<2) return {html:'I need at least two brands with data for that metric to compare across markets.'};
    rows.sort(function(a,z){return z.v-a.v;});
    var lead=rows[0];
    var t='<strong>'+esc2(mLabel(lead.mk,metric))+'</strong> · across markets · '+prettyDate(date)+'<table class="a-table"><tbody>';
    rows.forEach(function(r){
      var gapSig=''; if(r!==lead){ var st=sigTest(lead.v,baseAt(lead.mk,date),r.v,baseAt(r.mk,date)); gapSig=st.dir>0?' '+arrow(1):''; }
      t+='<tr><td>'+esc2(brandLabel(r.b))+' <span style="font-size:10px;color:var(--ink-3)">('+r.mk+')</span>'+(r===lead?' <span style="font-size:10px;color:var(--ink-3)">(lead)</span>':'')+'</td><td class="v">'+fmtPct(r.v)+'</td><td class="d">'+(r===lead?'':fmtPP(r.v-lead.v)+gapSig)+'</td></tr>';
    });
    t+='</tbody></table><div class="a-note">Cross-market comparison — each brand\u2019s figure and significance use its own market\u2019s tracker and base size. '+arrow(1)+' = leader significantly ahead. '+ctxNote()+'</div>';
    return {html:t, plain: rows.map(function(r){return brandLabel(r.b)+' ('+r.mk+'): '+fmtPct(r.v);}).join(', '), lead:{mk:lead.mk,brand:lead.b}};
  }
  // Compares ONE brand/metric across two arbitrary, explicitly-named months
  // (e.g. "Bet365 consideration, January vs June") — independent of the
  // dashboard's fixed MoM/YoY comparison setting.
  function periodCompareAnswer(mk,brand,metric,d1,d2){
    var v1=valueAt(mk,brand,metric,d1), v2=valueAt(mk,brand,metric,d2);
    if(v1===null||isNaN(v1)||v2===null||isNaN(v2)) return {html:'I don\u2019t have values for <strong>'+esc2(brandLabel(brand))+' · '+esc2(mLabel(mk,metric))+'</strong> in both of those months.'};
    var st=sigTest(v2,baseAt(mk,d2),v1,baseAt(mk,d1)), delta=v2-v1;
    var html='<strong>'+esc2(brandLabel(brand))+'</strong> · '+esc2(mLabel(mk,metric))+' <span style="color:var(--ink-3)">('+mk+')</span>'
      +'<table class="a-table"><tbody>'
      +'<tr><td>'+esc2(prettyDate(d1))+'</td><td class="v">'+fmtPct(v1)+'</td><td class="d"></td></tr>'
      +'<tr><td>'+esc2(prettyDate(d2))+'</td><td class="v">'+fmtPct(v2)+'</td><td class="d">'+fmtPP(delta)+' '+arrow(st.dir)+'</td></tr>'
      +'</tbody></table>'
      +'<div class="a-note">Direct comparison of the two months you named (separate from the dashboard\u2019s MoM/YoY setting). '+(st.dir!==0?'Significant':'Not significant')+' difference. '+ctxNote()+'</div>';
    return {html:html, plain: brandLabel(brand)+' '+mLabel(mk,metric)+': '+prettyDate(d1)+' '+fmtPct(v1)+' vs '+prettyDate(d2)+' '+fmtPct(v2)};
  }

  // ---------- proactive follow-up suggestions ----------
  // Shown under every successful data answer (not just failed ones) so
  // stakeholders can go deeper with one tap instead of typing a new question.
  function followUpChips(mk,brands,metric,kind){
    var chips=[], b=brands&&brands[0];
    if(!metric) return chips;
    if(b && kind!=='series') chips.push(brandLabel(b)+' '+mLabel(mk,metric)+' trend');
    if(kind!=='rank') chips.push('Top brands for '+mLabel(mk,metric)+' in '+mk);
    if(b){
      var others=brandsFor(mk).filter(function(x){return x!==b;});
      if(others.length && kind!=='compare') chips.push(brandLabel(b)+' vs '+brandLabel(others[0])+' '+mLabel(mk,metric));
    }
    return chips.slice(0,3);
  }

  // ---------- subgroup parsing (drives the dashboard's OWN filter state) ----------
  // We read audience/subgroup terms straight from the question, then set the
  // dashboard's AUDIENCE_SEL for the duration of the lookup and restore
  // them — so the figure is produced by exactly the same path as the on-screen
  // filters, never a re-implemented calculation.
  function parseSubgroups(low, market){
    var filters={}, audience=null, pseg=null, note=null;
    // gender (single)
    if(/\b(female|females|women|woman|ladies)\b/.test(low)) filters.gender='Female';
    else if(/\b(male|males|men|man|gents)\b/.test(low)) filters.gender='Male';
    // age (multi)
    var ages=[], at=low;
    // explicit ranges first; strip them so their digits don't re-trigger single bands
    if(/under\s*40/.test(at)){ ages.push('18-29','30-39'); at=at.replace(/under\s*40s?/g,' '); }
    if(/over\s*40|40\s*plus/.test(at)){ ages.push('40-49','50+'); at=at.replace(/over\s*40s?|40\s*plus/g,' '); }
    if(/under\s*30/.test(at)){ ages.push('18-29'); at=at.replace(/under\s*30s?/g,' '); }
    // single bands on the residual text
    if(/18\D{0,4}29|gen\s?z|youngest|younger/.test(at)) ages.push('18-29');
    if(/30\D{0,4}39|\b30s\b|thirties/.test(at)) ages.push('30-39');
    if(/40\D{0,4}49|\b40s\b|forties/.test(at)) ages.push('40-49');
    if(/50\s*\+|50\s*plus|over\s*50|\b50s\b|fifties|older|oldest|(^|\s)50\b/.test(at)) ages.push('50+');
    var AORD=['18-29','30-39','40-49','50+']; ages=AORD.filter(function(a){return ages.indexOf(a)>=0;});
    if(ages.length) filters.age=ages;
    // frequency (multi)
    var freq=[];
    if(/\bcore\b|high[\s-]?freq|frequent|heavy (player|bettor|user)|2\s*\+?\s*times|twice a week|regular(ly)?\b/.test(low)) freq.push('Core (2+ times per week)');
    if(/once a week|less frequent|infrequent|light (player|bettor|user)|occasional|casual\b/.test(low)) freq.push('Mass (Once a week or less)');
    if(freq.length) filters.frequency=freq;
    // player segment (multi)
    var seg=[];
    if(/\bsolus\b/.test(low)) seg.push('Solus');
    if(/gaming[\s-]?led/.test(low)) seg.push('Gaming-Led MPP');
    if(/bet(ting)?[\s-]?led/.test(low)) seg.push('Bet-Led MPP');
    if(seg.length) filters.segment=seg;
    // product "Other" (multi, market-specific) — require the "-led" cue to avoid
    // colliding with metric words like "slots" or "live".
    var opts=[]; try{opts=dimOptions('product',market)||[];}catch(e){}
    var prod=[];
    function addProd(re,val){ if(opts.indexOf(val)>=0 && re.test(low)) prod.push(val); }
    addProd(/football[\s-]?led/,'Football-Led');
    addProd(/racing[\s-]?led/,'Racing-Led');
    addProd(/sas[\s-]?led/,'SAS-Led');
    addProd(/slots?[\s-]?led/,'Slots-Led');
    addProd(/table[\s-]?games?[\s-]?led/,'Table Games-Led');
    addProd(/live[\s-]?led/,'Live-Led');
    addProd(/bingo[\s-]?led/,'Bingo-Led Slots');
    if(prod.length) filters.product=prod;
    // portfolio audience
    if(/next[\s-]?gen/.test(low)){ audience='portfolio'; pseg='nextgen'; }
    else if(/mass\s*30|mass\s*thirty/.test(low)){ audience='portfolio'; pseg='mass30'; }
    else if(/\bhvc\b|high[\s-]?value/.test(low)){ audience='portfolio'; pseg='hvc'; }
    else if(/\bportfolio\b/.test(low)){ audience='portfolio'; }
    // portfolio view isn't cross-filtered by subgroup here — don't silently combine
    if(audience==='portfolio' && Object.keys(filters).length){
      note='Note: the Portfolio view can\u2019t be combined with subgroup filters, so I answered for the '+(pseg?portfolioSegLabel(pseg):portfolioSegLabel())+' portfolio segment only.';
      filters={};
    }
    var any=(Object.keys(filters).length>0)||!!audience;
    var cleaned=low
      .replace(/\b(fe)?males?\b|\bwomen\b|\bwoman\b|\bmen\b|\bman\b|ladies|gents/g,' ')
      .replace(/18\D{0,4}29|30\D{0,4}39|40\D{0,4}49|50\s*\+|50\s*plus|over\s*50|under\s*30|under\s*40|over\s*40|40\s*plus|\b\d0s\b|thirties|forties|fifties|younger|youngest|older|oldest|gen\s?z/g,' ')
      .replace(/\bcore\b|once a week|less frequent|infrequent|occasional|casual|frequent|regularly|heavy|light/g,' ')
      .replace(/\bsolus\b|gaming[\s-]?led|bet(ting)?[\s-]?led/g,' ')
      .replace(/(football|racing|sas|slots?|table[\s-]?games?|live|bingo)[\s-]?led/g,' ')
      .replace(/next[\s-]?gen|mass\s*30|mass\s*thirty|\bhvc\b|high[\s-]?value|\bportfolio\b/g,' ');
    return {any:any, filters:filters, audience:audience, seg:pseg, note:note, cleaned:cleaned};
  }
  // Collapse a parsed question to ONE audience key (audiences are never combined).
  function audKeyFromParsed(parsed){
    if(!parsed) return null;
    if(parsed.audience==='portfolio' && parsed.seg) return 'port:'+parsed.seg;
    var f=parsed.filters||{};
    function first(v){ return Array.isArray(v)?v[0]:v; }
    if(f.product)   return 'product:'+first(f.product);      // most specific first
    if(f.segment)   return 'segment:'+first(f.segment);
    if(f.frequency) return 'frequency:'+first(f.frequency);
    if(f.age)       return 'age:'+first(f.age);
    if(f.gender)    return 'gender:'+first(f.gender);
    return null;
  }
  // Run fn() with the single parsed audience applied to the live dashboard state,
  // then restore. If the question named more than one cut, only the closest single
  // audience is used (with a note) — nothing is NETted.
  function withSub(parsed, fn){
    if(!parsed || !parsed.any) return fn();
    var snap=AUDIENCE_SEL;
    try{
      var key=audKeyFromParsed(parsed);
      if(key && audienceList(MARKET).some(function(a){return a.key===key;})) AUDIENCE_SEL=key;
      var f=parsed.filters||{}, cuts=0; ['gender','age','frequency','segment','product'].forEach(function(d){ if(f[d]) cuts++; });
      if(parsed.audience==='portfolio') cuts++;
      if(cuts>1 && !parsed.note){ var m=audienceList(MARKET).find(function(a){return a.key===key;}); parsed.note='Note: audiences can\u2019t be combined here — each is read only from its own tab — so I answered for the single closest audience ('+((m&&m.label)||'Total Market')+') only.'; }
      return fn();
    } finally { AUDIENCE_SEL=snap; }
  }
  function pack(parsed, res){ if(parsed && parsed.note && res && res.html){ res.html+='<div class="a-note">'+parsed.note+'</div>'; } return res; }

  // Cross-market brand pairs, used only as a FALLBACK when the primary
  // (single-market) resolution can't find two brands to compare — this keeps
  // ordinary same-market comparisons (which resolve fine on their own) from
  // ever being mis-routed here just because a short name (e.g. "Paddy") is
  // also a synonym in another market.
  function crossMarketPairs(low){
    var pairs=[],seen={};
    MARKETS.forEach(function(m){
      brandsInText(low,m).forEach(function(b){ var k=m+'|'+b; if(!seen[k]){seen[k]=1; pairs.push({mk:m,brand:b});} });
    });
    var distinct={}; pairs.forEach(function(p){distinct[p.mk]=1;});
    return Object.keys(distinct).length>=2 ? pairs : null;
  }
  function marketClarification(ties, rawText){
    var chips=ties.slice(0,3).map(function(mk){ return mk+': '+rawText; });
    var labelled=ties.map(function(mk){return '<strong>'+esc2(mk)+'</strong>';}).join(' or ');
    return {html:'That brand name could be in more than one market here — '+labelled+'. Tap one to pick it, or add the market name to your question next time.', chips:chips};
  }
  function metricClarification(mk, alts, brands, dateWord){
    var pre = brands.length ? brandLabel(brands[0])+' ' : '';
    var chips = alts.map(function(m){ return pre+mLabel(mk,m)+' '+(dateWord||'latest'); });
    return {html:'I found two metrics that could match that — did you mean <strong>'+esc2(mLabel(mk,alts[0]))+'</strong> or <strong>'+esc2(mLabel(mk,alts[1]))+'</strong>?', chips:chips};
  }

  // ---------- main engine ----------
  function askEngine(text){
    var low=' '+text.toLowerCase().replace(/[’]/g,"'")+' ';
    // methodology / definition first when clearly asked
    var wantsDef=/\b(define|definition|what (is|are|does)|meaning|mean\b|explain)\b/.test(low);
    var methoHit=/(significan|confidence|z[\s-]?test|\bbase\b|sample size|methodolog|two[\s-]?proportion|\bmom\b|\byoy\b|month[\s-]on[\s-]month|year[\s-]on[\s-]year|\bnet\b|weighted|portfolio|audience view|dummy|synthetic|data source|which cut|\bfilter)/.test(low);

    var preferredMk = (LAST && LAST.mk) || MARKET;
    var explicitMk = detectMarket(low);
    var mk;
    if(explicitMk){ mk=explicitMk; }
    else {
      var inferred=inferMarket(low, preferredMk);
      if(LAST_MARKET_TIES.length>1){ logDiag({q:text,reason:'ambiguous market: '+LAST_MARKET_TIES.join('/')}); return marketClarification(LAST_MARKET_TIES, text); }
      mk = inferred || preferredMk;
    }
    var parsed=parseSubgroups(low, mk);
    var clow=parsed.any ? parsed.cleaned : low;
    var brands=brandsInText(clow,mk);
    var metric=resolveMetric(mk,clow);
    var metricAlts=LAST_METRIC_ALTS.slice();

    // ---- follow-up context: fill in anything this question didn't mention ----
    var usedLastBrand=false, usedLastMetric=false, usedLastDate=false;
    if(!brands.length && LAST && LAST.mk===mk && LAST.brands && LAST.brands.length){ brands=LAST.brands; usedLastBrand=true; }
    if(!metric && LAST && LAST.mk===mk && LAST.metric){ metric=LAST.metric; usedLastMetric=true; }

    // definition of a metric ("what does consideration mean")
    if(wantsDef && metric && !brands.length){ return {html:definitionAnswer(mk,metric)}; }
    // pure methodology
    if((methoHit && !(brands.length&&metric)) || (wantsDef && !metric && !brands.length)){
      var ma=methodologyAnswer(low); if(ma) return {html:ma};
    }

    // ask, rather than silently guess, when two metrics are a close match
    if(metric && metricAlts.length===2 && !wantsDef && !methoHit){
      return metricClarification(mk, metricAlts, brands, null);
    }

    // data intents
    var mres=resolveMonth(clow);
    if(mres.date===null && mres.req){ return {html:'I only hold data from <strong>'+prettyDate(DATES[0])+'</strong> to <strong>'+prettyDate(DATES[DATES.length-1])+'</strong>, so I can\u2019t give you '+prettyDate(mres.req)+'.'}; }
    var date=mres.date;
    if(date===null){
      if(LAST && LAST.mk===mk && LAST.date){ date=LAST.date; usedLastDate=true; } else date=PERIOD;
    }

    var isTrend=/\btrend|over time|evolution|history|trajectory|movement|last \d+ months|past \d+ months|over the (last|past)\b/.test(low);
    var winM=low.match(/(?:last|past)\s+(\d{1,2})\s+months?/); var winN=winM?Math.max(2,Math.min(25,+winM[1])):12;
    var isRank=/\b(top|highest|lowest|best|worst|leading|leader|leads|rank|ranking|which brand)\b/.test(low);
    var isCompareWord=/\b(vs|versus|compare|compared to|against|difference between)\b/.test(low);
    var isCompare=isCompareWord && brands.length>=2;

    function ctxSuffix(){
      var bits=[]; if(usedLastBrand)bits.push('brand'); if(usedLastMetric)bits.push('metric'); if(usedLastDate)bits.push('month');
      return bits.length ? '<div class="a-note">Carried over from your last question: '+bits.join(', ')+'.</div>' : '';
    }
    // ctxOverride lets a branch (e.g. cross-market) report a different
    // market/brand/date than the ambient mk/date — needed whenever the
    // answer isn't anchored to a single market the normal way.
    function finish(res, kind, brandsUsed, ctxOverride){
      if(!res) return res;
      if(res.html) res.html += ctxSuffix();
      res = pack(parsed, res);
      if(metric){
        var useMk=(ctxOverride&&ctxOverride.mk)||mk;
        var useBrands=(ctxOverride&&ctxOverride.brand)?[ctxOverride.brand]:(brandsUsed||brands);
        var useDate=(ctxOverride&&ctxOverride.date)||date;
        res.chips = followUpChips(useMk, useBrands, metric, kind);
        res.ctx = {mk:useMk, brand:useBrands[0]||null, metric:metric, date:useDate, audKey:audKeyFromParsed(parsed)};
        rememberContext(useMk, useBrands, metric, useDate);
      }
      return res;
    }

    // custom two-period comparison for a single brand ("Jan vs June")
    if(metric && brands.length===1 && isCompareWord){
      var twoM=resolveTwoMonths(clow);
      if(twoM) return finish(withSub(parsed, function(){return periodCompareAnswer(mk,brands[0],metric,twoM.d1,twoM.d2);}), 'period', brands, {mk:mk,brand:brands[0],date:twoM.d2});
    }

    if(metric && brands.length>=2 && (isCompare||(!isTrend&&!isRank))) return finish(withSub(parsed, function(){return compareAnswer(mk,brands,metric,date);}), 'compare', brands);

    // cross-market comparison — fallback only, when the single-market path
    // couldn't find two brands but the question clearly wants a comparison
    if(metric && isCompareWord && brands.length<2){
      var cmPairs=crossMarketPairs(clow);
      var validPairs=cmPairs?cmPairs.filter(function(p){return metricsFor(p.mk).indexOf(metric)>=0;}):[];
      if(validPairs.length>=2){
        var cc=crossCompareAnswer(validPairs,metric,date);
        return finish(cc, 'compare', validPairs.map(function(p){return p.brand;}), cc.lead);
      }
    }

    if(metric && isRank && !brands.length) return finish(withSub(parsed, function(){return rankAnswer(mk,metric,date,/\b(lowest|worst)\b/.test(low));}), 'rank', brands);
    if(metric && brands.length===1 && isTrend) return finish(withSub(parsed, function(){return seriesAnswer(mk,brands[0],metric,date,winN);}), 'series', brands);
    if(metric && brands.length===1) return finish(withSub(parsed, function(){return pointAnswer(mk,brands[0],metric,date);}), 'point', brands);
    if(metric && !brands.length && isRank) return finish(withSub(parsed, function(){return rankAnswer(mk,metric,date,/\b(lowest|worst)\b/.test(low));}), 'rank', brands);

    // last-chance methodology
    var ma2=methodologyAnswer(low); if(ma2) return {html:ma2};

    // couldn't fill the slots — guide the user, and privately log the miss
    var miss=[];
    if(!brands.length && !isRank) miss.push('a brand');
    if(!metric) miss.push('a metric');
    logDiag({q:text, reason:'unresolved'+(miss.length?' — missing '+miss.join(' and '):'')});
    return {html:'I couldn\u2019t quite pin that down'+(miss.length?' — I was missing '+miss.join(' and ')+'.':'.')+' I can fetch a specific figure, a trend, a ranking, or a comparison, and I can explain the methodology.', help:true};
  }

  // ---------- UI ----------
  var panel=document.getElementById('a-panel'), body=document.getElementById('a-body'),
      input=document.getElementById('a-input'), seeded=false;
  function scrollDown(){ body.scrollTop=body.scrollHeight; }
  function addUser(t){ var d=document.createElement('div'); d.className='a-msg user'; d.textContent=t; body.appendChild(d); scrollDown(); }

  // Copies a plain-text figure to the clipboard (falls back to a hidden
  // textarea + execCommand for older/locked-down browsers).
  function copyPlain(text, btn){
    var v=text||'';
    function ok(){ if(!btn) return; var was=btn.textContent; btn.textContent='Copied'; setTimeout(function(){btn.textContent=was;},1200); }
    if(navigator.clipboard && navigator.clipboard.writeText){ navigator.clipboard.writeText(v).then(ok, function(){}); }
    else { try{ var ta=document.createElement('textarea'); ta.value=v; document.body.appendChild(ta); ta.select(); document.execCommand('copy'); document.body.removeChild(ta); ok(); }catch(e){} }
  }
  // Pushes an answer's market/brand/metric/month/audience back onto the live
  // dashboard filters, then re-renders — so a stakeholder can go from "what
  // the assistant said" to "see it in the dashboard" in one tap.
  function applyToDashboard(ctx){
    if(!ctx || !ctx.mk || !ctx.metric) return;
    if(ctx.mk!==MARKET){ MARKET=ctx.mk; syncMarketDependentControls(); }
    if(comparisonMetrics(MARKET).some(function(r){return r.metric===ctx.metric;})) METRIC=ctx.metric;
    var pickBrands=COMPARISON_BRANDS[MARKET]||brandsFor(MARKET);
    if(ctx.brand && pickBrands.indexOf(ctx.brand)>=0) BRAND=ctx.brand;
    if(ctx.date && DATES.indexOf(ctx.date)>=0) PERIOD=ctx.date;
    AUDIENCE_SEL=(ctx.audKey && audienceList(MARKET).some(function(a){return a.key===ctx.audKey;}))?ctx.audKey:'total';
    try{
      document.getElementById('market-select').value=MARKET;
      document.getElementById('period-select').value=PERIOD;
      syncMarketDependentControls();
      document.getElementById('metric-select').value=METRIC;
      document.getElementById('audience-select').value=AUDIENCE_SEL;
    }catch(e){}
    renderAll();
    if(typeof showToast==='function') showToast('Dashboard updated to match this answer');
  }
  function addBot(html,chips,ctx,plain){
    var d=document.createElement('div'); d.className='a-msg bot'; d.innerHTML=html;
    if(chips&&chips.length){ var c=document.createElement('div'); c.className='a-chips';
      chips.forEach(function(q){ var b=document.createElement('button'); b.className='a-chip'; b.textContent=q; b.addEventListener('click',function(){run(q);}); c.appendChild(b); });
      d.appendChild(c); }
    if(plain || (ctx&&ctx.mk&&ctx.metric)){
      var act=document.createElement('div'); act.className='a-chips'; act.style.marginTop='2px';
      if(plain){
        var cb=document.createElement('button'); cb.className='a-chip'; cb.textContent='Copy figure';
        cb.addEventListener('click',function(){ copyPlain(plain, cb); });
        act.appendChild(cb);
      }
      if(ctx && ctx.mk && ctx.metric){
        var ab=document.createElement('button'); ab.className='a-chip'; ab.textContent='Apply to dashboard';
        ab.addEventListener('click',function(){ applyToDashboard(ctx); });
        act.appendChild(ab);
      }
      d.appendChild(act);
    }
    body.appendChild(d); scrollDown();
  }
  function exampleChips(){
    var mk=MARKET, focus=null;
    try{ focus=brandsFor(mk).filter(function(b){return brandMeta(mk,b).focus;})[0]; }catch(e){}
    focus=focus||brandsFor(mk)[0];
    var comp=brandsFor(mk).filter(function(b){return b!==focus;})[0]||brandsFor(mk)[0];
    return [
      brandLabel(focus)+' consideration in '+mk+' latest',
      brandLabel(focus)+' consideration among 18-29s',
      'Top brands for prompted awareness in '+mk,
      brandLabel(focus)+' vs '+brandLabel(comp)+' most often used',
      brandLabel(focus)+' used in last month trend',
      'How do you decide what\u2019s significant?'
    ];
  }
  function seed(){
    if(seeded)return; seeded=true;
    addBot('Hi — I can pull any figure that\u2019s in this dashboard and explain how it\u2019s calculated. I read straight from the loaded data, so I\u2019ll never invent a number. You can also name an audience in your question — e.g. \u201Camong 18-29s\u201D, \u201Cfor women\u201D or \u201Ccore players\u201D — and I\u2019ll apply it. Once we\u2019re talking about a brand, feel free to just ask follow-ups like \u201Cand last month?\u201D without repeating everything. Try one of these, or just ask:', exampleChips());
  }
  function run(q){
    addUser(q);
    var r;
    try{ r=askEngine(q); }
    catch(e){ r={html:'Sorry — I hit a snag reading that. Try rephrasing?'}; try{ logDiag({q:q, reason:'error: '+((e&&e.message)||e)}); }catch(e2){} }
    addBot(r.html, r.chips || (r.help?exampleChips():null), r.ctx, r.plain);
  }
  function send(){ var v=input.value.trim(); if(!v)return; input.value=''; run(v); }

  function openPanel(){ panel.classList.add('open'); seed(); setTimeout(function(){input.focus();},60); }
  function closePanel(){ panel.classList.remove('open'); }
  document.getElementById('a-launch').addEventListener('click',openPanel);
  document.getElementById('a-close').addEventListener('click',closePanel);
  document.getElementById('a-send').addEventListener('click',send);
  input.addEventListener('keydown',function(e){ if(e.key==='Enter'){e.preventDefault();send();} });
  document.addEventListener('keydown',function(e){ if(e.key==='Escape'&&panel.classList.contains('open'))closePanel(); });

  // ---------- private diagnostics panel ----------
  // Shows this browser's own log of questions the assistant couldn't resolve
  // or had to ask for clarification on — for the dashboard owner to review
  // and use to improve the synonym/metric tables. Nothing is sent anywhere;
  // it's just a read-out of localStorage. It is NOT linked from the normal
  // UI — it only opens via Ctrl+Alt+D, or by loading the file with ?diag=1
  // in the URL — so a stakeholder using the dashboard normally never sees it.
  var diagPanel=null;
  function buildDiagPanel(){
    if(diagPanel) return diagPanel;
    var p=document.createElement('div');
    p.style.cssText='position:fixed;right:22px;bottom:96px;z-index:200;width:360px;max-width:calc(100vw - 44px);max-height:60vh;background:#fff;border:1px solid var(--line);border-radius:12px;box-shadow:0 24px 64px rgba(15,31,74,.28);display:none;flex-direction:column;overflow:hidden;font-family:var(--font-body);';
    p.innerHTML='<div style="padding:10px 12px;background:var(--navy);color:#fff;display:flex;justify-content:space-between;align-items:center;font-family:var(--font-display);font-weight:600;font-size:13px;">Assistant diagnostics (private)<button id="diag-close" aria-label="Close diagnostics" style="background:transparent;border:none;color:#fff;cursor:pointer;font-size:16px;line-height:1;">×</button></div>'
      +'<div id="diag-list" style="padding:10px 12px;overflow-y:auto;font-size:12px;color:var(--ink);flex:1;"></div>'
      +'<div style="padding:8px 12px;border-top:1px solid var(--line);display:flex;gap:8px;">'
      +'<button id="diag-copy" style="flex:1;padding:6px;border:1px solid var(--line-2);border-radius:6px;background:#fff;cursor:pointer;font-size:12px;">Copy log</button>'
      +'<button id="diag-clear" style="padding:6px 10px;border:1px solid var(--line-2);border-radius:6px;background:#fff;cursor:pointer;font-size:12px;">Clear</button></div>';
    document.body.appendChild(p);
    p.querySelector('#diag-close').addEventListener('click', closeDiagPanel);
    p.querySelector('#diag-clear').addEventListener('click', function(){ saveDiag([]); renderDiagList(); });
    p.querySelector('#diag-copy').addEventListener('click', function(){
      var list=loadDiag();
      var txt=list.map(function(e){return e.t+' — "'+e.q+'" — '+e.reason;}).join('\n');
      copyPlain(txt, p.querySelector('#diag-copy'));
    });
    diagPanel=p;
    return p;
  }
  function renderDiagList(){
    var list=loadDiag().slice().reverse();
    var el=diagPanel.querySelector('#diag-list');
    if(!list.length){ el.innerHTML='<div style="color:var(--ink-3)">No unresolved questions logged yet.</div>'; return; }
    el.innerHTML=list.map(function(e){
      return '<div style="padding:6px 0;border-bottom:1px solid var(--line)"><div style="font-weight:600;">'+esc2(e.q)+'</div><div style="color:var(--ink-3);font-size:11px;">'+esc2(e.reason)+' · '+new Date(e.t).toLocaleString('en-GB')+'</div></div>';
    }).join('');
  }
  function openDiagPanel(){ buildDiagPanel(); renderDiagList(); diagPanel.style.display='flex'; }
  function closeDiagPanel(){ if(diagPanel) diagPanel.style.display='none'; }
  function toggleDiagPanel(){ if(diagPanel && diagPanel.style.display==='flex') closeDiagPanel(); else openDiagPanel(); }
  document.addEventListener('keydown', function(e){ if(e.ctrlKey && e.altKey && (e.key==='d'||e.key==='D')){ e.preventDefault(); toggleDiagPanel(); } });
  try{ if(/(?:^|[?&])diag=1(?:&|$)/.test(window.location.search)) openDiagPanel(); }catch(e){}
})();
</script>


/* ==========================================================================
   D. COMPARE TOGGLE — MoM / Mo3M / YoY  (verbatim, fully reusable —
   single-source: 202607_Product_Pulse)
   ==========================================================================
   The lag-based math behind a "compare to N months ago" toggle. Genuinely
   generic — three lines of lookup table plus one delta function — but it's
   the piece both dashboards' Compare segmented controls (the "Absolute /
   MoM % / YoY %" buttons already in boilerplate.html's filters bar) need
   in order to actually DO something, rather than just sit there as UI.

   CONTRACT: you need one plain array of numbers per series — one value per
   month, oldest first, gaps as null — plus the index of the month
   currently being viewed. Everything else follows from that.

     var COMPARE_LAG = { mom: 1, mo3m: 3, yoy: 12 };

     // % change from `lag` months before index `pIdx` in a monthly series.
     // Returns null if there's no data far enough back (start of series)
     // or the earlier value is null/zero/NaN (can't divide by it).
     function deltaPct(cur, prior) {
       if (cur == null || prior == null || prior === 0 || isNaN(cur) || isNaN(prior)) return null;
       return (cur - prior) / prior;
     }
     function deltaAtLag(series, pIdx, lag) {
       var p = pIdx - lag;
       return p >= 0 ? deltaPct(series[pIdx], series[p]) : null;
     }

   WIRING — add Mo3M to the boilerplate's existing Compare segmented
   control (it currently only has Absolute/MoM %/YoY %):

     <div class="segmented" role="radiogroup" aria-label="Compare">
       <button type="button" data-mode="absolute" class="active" role="radio" aria-checked="true">Absolute</button>
       <button type="button" data-mode="mom" role="radio" aria-checked="false">MoM %</button>
       <button type="button" data-mode="mo3m" role="radio" aria-checked="false">Mo3M %</button>
       <button type="button" data-mode="yoy" role="radio" aria-checked="false">YoY %</button>
     </div>

   Then, wherever your dashboard currently reads `state.compare` (or
   whatever you call the active mode) to decide what to render: if the
   mode is 'absolute', render the raw series; otherwise look up
   COMPARE_LAG[mode] and call deltaAtLag(series, pIdx, lag) for each point
   you need a delta for. That's the entire feature — there's no
   chart-library-specific part; deltaAtLag returns a plain number (or
   null) whether you're feeding a Chart.js dataset, a table cell, or a
   scorecard delta.

   Rendering a delta once you have it — pair with the .threshold-chip
   classes in widgets.css, or the small inline chip pattern the source
   uses (adapt colours/wording, these two are illustrative, not a class
   you need to add anywhere):

     function deltaPill(d) {
       if (d === null) return '<span class="nsm-delta flat">n/a</span>';
       var cls = d > 0.0005 ? 'pos' : (d < -0.0005 ? 'neg' : 'flat');
       var arrow = d > 0.0005 ? '\u25b2' : (d < -0.0005 ? '\u25bc' : '\u2013');
       return '<span class="nsm-delta ' + cls + '">' + arrow + ' ' + (Math.abs(d) * 100).toFixed(1) + '%</span>';
     }

   .nsm-delta isn't a widgets.css class — style it yourself the same way
   you'd style any .pos/.neg/.flat text (var(--pos-fg)/var(--neg-fg)/
   var(--ink-3) are the existing tokens for that), or swap in
   .threshold-chip directly if a chip look fits your layout better than
   inline coloured text.
   ========================================================================== */


/* ==========================================================================
   E. BRAND-SPLIT OVERLAY CHART  (mixed — dataset builder is verbatim/
   reusable, everything upstream of it is bespoke data-lookup glue —
   single-source: 202607_Product_Pulse)
   ==========================================================================
   "Brand-split comparisons" — several brands drawn as overlaid line series
   on ONE chart, not side-by-side small multiples (that's the separate
   "Multi-chart comparison panel" widget in widgets.css — a different
   thing that looks superficially similar in a screenshot).

   The reusable core is small: one function that turns a brand + a plain
   array of numbers into one Chart.js dataset object, coloured with that
   brand's own token from design-tokens.css. Call it once per selected
   brand, collect the results into one array, hand that whole array to a
   single `new Chart(canvas, { data: { labels, datasets: [...] } })` call
   — that's the entire "overlay" mechanic:

     // BRAND_COLOR should already exist wherever your dashboard defines
     // its brand list — map each brand key to the corresponding
     // design-tokens.css var (getComputedStyle(document.documentElement)
     // .getPropertyValue('--pp') etc.), don't hardcode hex here.
     function overlayDataset(brandKey, brandLabel, color, data, opts) {
       opts = opts || {};
       var ds = {
         label: brandLabel, data: data, borderColor: color,
         backgroundColor: color, borderWidth: 2, tension: 0.4,
         cubicInterpolationMode: 'monotone', pointRadius: 2,
         pointHoverRadius: 5, spanGaps: true
       };
       if (opts.stacked) {
         ds.backgroundColor = color; // add alpha yourself if you want a fill tint
         ds.fill = opts.index === 0 ? 'origin' : '-1';
         ds.borderWidth = 1.2; ds.pointRadius = 0; ds.pointHoverRadius = 3;
       }
       return ds;
     }

     // Given the brands currently selected for overlay, build the chart:
     var datasets = selectedBrands.map(function (b, i) {
       return overlayDataset(b.key, b.label, b.color, b.monthlyValues, { index: i });
     });
     new Chart(canvas, { type: 'line', data: { labels: months, datasets: datasets } });

   Everything upstream of `b.monthlyValues` — pulling the right metric,
   for the right brand, for the right vertical, out of your dashboard's
   own data model — is inherently bespoke to your report's data shape.
   The source's version of that step (scoped(), scopedDriver()) is
   hundreds of lines tied to its specific pyramid/tier structure and isn't
   reproduced here; write your own lookup that ends with "an array of
   numbers, one per month, for this brand" and everything above this line
   works unchanged.

   BRAND MULTI-SELECT UI — the "which brands are currently overlaid"
   picker uses the `.brand-split` pills in widgets.css (not the plain
   `.brand-pill-group`, which is single-select). Wiring:

     function renderBrandSplitPills(containerEl, allBrands, selectedKeys, onChange) {
       containerEl.innerHTML = allBrands.map(function (b) {
         var on = selectedKeys.indexOf(b.key) !== -1;
         var canRemove = on && selectedKeys.length > 1;
         var addBtn = !on
           ? '<button type="button" class="bs-add" data-brand="' + b.key + '" aria-label="Add ' + b.label + '">+</button>'
           : canRemove
             ? '<button type="button" class="bs-add" data-brand="' + b.key + '" aria-label="Remove ' + b.label + '">\u2212</button>'
             : '';
         return '<span class="brand-split brand-badge ' + b.badgeClass + (on ? ' active' : '') + (on && !addBtn ? ' solo' : '') + '" data-brand="' + b.key + '">' +
           '<button type="button" class="bs-main" data-brand="' + b.key + '" aria-label="Show only ' + b.label + '">' + b.label + '</button>' +
           addBtn + '</span>';
       }).join('');
       containerEl.querySelectorAll('.bs-main').forEach(function (btn) {
         btn.addEventListener('click', function () { onChange([btn.getAttribute('data-brand')]); });
       });
       containerEl.querySelectorAll('.bs-add').forEach(function (btn) {
         btn.addEventListener('click', function () {
           var key = btn.getAttribute('data-brand');
           var i = selectedKeys.indexOf(key);
           var next = selectedKeys.slice();
           if (i === -1) next.push(key); else if (next.length > 1) next.splice(i, 1);
           onChange(next);
         });
       });
     }

   `b.badgeClass` should be one of the sbg/pp/bfu/bfe/tom/tomg/core classes
   from design-tokens.css (or your own equivalent) so the pill picks up the
   corrected brand colour automatically — don't hardcode a colour here
   either.
   ========================================================================== */


/* ==========================================================================
   F. DATA DICTIONARY PAGE — render pattern  (mixed — same tiering as the
   AI assistant — single-source: 202607_Product_Pulse; CSS lives in
   widgets.css under "DATA DICTIONARY PAGE")
   ==========================================================================
   The source's renderDefinitions() is ~200 lines because it walks a very
   specific three-tier metric taxonomy (Volume/Conversion/Value) plus a
   raw-feed-field appendix, and isn't reproduced verbatim here — it's tied
   too tightly to that report's own metrics to be a useful copy/paste. What
   IS worth reusing is the shape:

     1. One lookup object per "kind" of thing you're documenting — e.g.
        METRIC_DESC = { groupName: { metricKey: 'plain-English definition', ... } },
        RAW_FIELD_DESC = { 'raw feed column name': 'what it means', ... },
        and optionally a SOURCE map (metricKey -> 'GA4' / 'Redshift' /
        'Calculated field' / whatever provenance labels matter to your team).
     2. A small esc() helper (escape &, <, > — never trust definition text
        into innerHTML unescaped) and a tbl(inner) helper that wraps a
        <table> in .dd-table-wrap:
          function esc(s) { return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
          function tbl(inner) { return '<div class="dd-table-wrap"><table class="dd-table">' + inner + '</table></div>'; }
     3. One render function that builds a `.dd-lead` intro paragraph, then
        loops your lookup object(s) into `.dd-block` sections, each with an
        `<h3>` and one `tbl(...)` call — <thead><tr><th>Metric</th>
        <th>Definition</th>[<th>Source</th>]</tr></thead>, one <tr> per
        entry in the lookup object.
     4. Wire it into your section-switching the same way as any other
        section: a sidenav button, a matching <section>, and a call to
        your render function inside whatever dispatch function
        (renderActive() in the source) runs when that section becomes
        visible.

   Use `.dd-tag` for short inline labels inside definition text (the
   source uses it for "MoM"/"Mo3M"/"YoY" glossary rows), `.dd-tag.nsm` for
   a "this is a North Star Metric" flag, and `.mono` when quoting a raw
   feed field name verbatim so it reads unambiguously as a literal string
   rather than prose.
   ========================================================================== */


/* ==========================================================================
   G. CHART EMBED AS RAW HTML  (NOT sourced from any dashboard)
   ==========================================================================
   None of the three source dashboards actually built this — it was on the
   wishlist ("Raw HTML export, alongside the standard PNG/CSV — useful if
   someone downstream needs to embed the dashboard elsewhere") but never
   implemented anywhere Claude could extract it from. Everything below is
   written fresh to fit the existing conventions, not pulled from a
   working file — treat it as a reasonable starting point to test against
   your own chart data, not as something already proven in production the
   way the PPTX engine or AI assistant are.

   The idea: alongside "Download PNG" / "Download CSV" in a chart's export
   menu, a "Download HTML" option that produces a small, genuinely
   standalone .html file — a bare page with one <canvas>, a Chart.js CDN
   tag, and this chart's exact config inlined as JSON — so a recipient can
   open it directly, or drop the whole file into a CMS/email that allows
   raw HTML, and get a live (not static-image) version of just this one
   chart, with no dependency on the rest of the dashboard.

   CONTRACT: reuse whatever Chart.js `config` object your chart-modal
   already builds for PNG/CSV export (type, data, options) — this doesn't
   need a second copy of your chart-building logic, only that object.

     function buildChartEmbedHtml(chartConfig, title) {
       var configJson = JSON.stringify(chartConfig).replace(/</g, '\\u003c'); // no early </script>
       return '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
         '<title>' + (title ? String(title).replace(/</g, '&lt;') : 'Chart') + '</title>' +
         '<style>html,body{margin:0;padding:16px;font-family:sans-serif;background:#fff;}' +
         '#wrap{max-width:900px;margin:0 auto;height:480px;}</style>' +
         '</head><body><div id="wrap"><canvas id="c"></canvas></div>' +
         '<script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>' +
         '<script>new Chart(document.getElementById("c"), ' + configJson + ');</script>' +
         '</body></html>';
     }

     function downloadChartEmbedHtml(chartConfig, title) {
       var html = buildChartEmbedHtml(chartConfig, title);
       var blob = new Blob([html], { type: 'text/html' });
       var a = document.createElement('a');
       a.href = URL.createObjectURL(blob);
       a.download = (title || 'chart').replace(/[^a-z0-9]+/gi, '_') + '.html';
       a.click();
       URL.revokeObjectURL(a.href);
     }

   Add a fourth button to the existing chart-modal export menu — no new
   CSS, it's the same `.export-opt` class the CSV/clipboard/PNG buttons
   already use:

     <button type="button" class="export-opt" data-export="html" role="menuitem">
       <span class="opt-icon" aria-hidden="true">&lt;/&gt;</span>Download HTML
     </button>

   CAVEATS worth testing before you rely on this: the exported file
   depends on the Chart.js CDN being reachable wherever it's opened
   (unlike the PPTX engine, this is NOT offline-safe as written — inline
   the Chart.js bundle the same way the AI assistant's fflate dependency
   is inlined, if that matters for your use case). It also only carries
   the chart itself, not any surrounding page chrome, title styling, or
   your brand fonts — Chart.js's own default fonts will render instead
   unless you set them explicitly in the config's `options.font`.
   ========================================================================== */

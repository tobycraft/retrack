// Diff/OOXML engine shared by taskpane.html (task pane button) and
// commands.html (ribbon button). Pure functions, no Office.js dependency —
// callers pass in the selection text/OOXML and clipboard text.

// ── Tokenizer ─────────────────────────────────────────────────────────────
// Word sequences | punctuation sequences | whitespace sequences
function tokenize(text) {
  return text.match(/\w+|[^\w\s]+|\s+/g) || [];
}

// ── Myers diff ────────────────────────────────────────────────────────────
// Returns [{type: 'equal'|'delete'|'insert', val: token}, ...]
function myersDiff(a, b) {
  const n = a.length, m = b.length;
  if (n === 0 && m === 0) return [];
  if (n === 0) return b.map(val => ({ type: 'insert', val }));
  if (m === 0) return a.map(val => ({ type: 'delete', val }));

  const max = n + m, off = max;
  const v = new Array(2 * max + 1).fill(0);
  const trace = [];

  outer:
  for (let d = 0; d <= max; d++) {
    trace.push(v.slice());
    for (let k = -d; k <= d; k += 2) {
      let x = (k === -d || (k !== d && v[k - 1 + off] < v[k + 1 + off]))
        ? v[k + 1 + off]
        : v[k - 1 + off] + 1;
      let y = x - k;
      while (x < n && y < m && a[x] === b[y]) { x++; y++; }
      v[k + off] = x;
      if (x >= n && y >= m) break outer;
    }
  }

  const edits = [];
  let x = n, y = m;

  for (let d = trace.length - 1; d > 0; d--) {
    const prev = trace[d];
    const k = x - y;
    const prevK = (k === -d || (k !== d && prev[k - 1 + off] < prev[k + 1 + off]))
      ? k + 1 : k - 1;
    const prevX  = prev[prevK + off];
    const afterX = prevK === k - 1 ? prevX + 1 : prevX;
    while (x > afterX) { edits.push({ type: 'equal',  val: a[x - 1] }); x--; y--; }
    if (prevK === k - 1) { edits.push({ type: 'delete', val: a[x - 1] }); x--; }
    else                 { edits.push({ type: 'insert', val: b[y - 1] }); y--; }
  }
  while (x > 0) { edits.push({ type: 'equal', val: a[x - 1] }); x--; y--; }

  return edits.reverse();
}

// ── OOXML helpers ─────────────────────────────────────────────────────────
function extractRPr(ooxml) {
  const m = ooxml.match(/<w:rPr(?:\s[^>]*)?\/>|<w:rPr(?:[^>]*)?>[\s\S]*?<\/w:rPr>/);
  return m ? m[0] : '';
}

// Reuse the source document's namespace declarations to cover any w14:/w15:/etc. in rPr.
function extractDocumentOpenTag(ooxml) {
  const m = ooxml.match(/<w:document\b[^>]*>/);
  return m ? m[0] : '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">';
}

function esc(s) {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function buildOoxml(edits, rPr, docTag, author, date) {
  let id = 1;
  let runs = '';
  let i = 0;

  while (i < edits.length) {
    const type = edits[i].type;
    let text = '';
    while (i < edits.length && edits[i].type === type) { text += edits[i].val; i++; }
    if (type === 'equal')
      runs += `<w:r>${rPr}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
    else if (type === 'delete')
      runs += `<w:del w:id="${id++}" w:author="${esc(author)}" w:date="${date}">`
            + `<w:r>${rPr}<w:delText xml:space="preserve">${esc(text)}</w:delText></w:r>`
            + `</w:del>`;
    else
      runs += `<w:ins w:id="${id++}" w:author="${esc(author)}" w:date="${date}">`
            + `<w:r>${rPr}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`
            + `</w:ins>`;
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
       + `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">`
       + `<pkg:part pkg:name="/_rels/.rels"`
       +   ` pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"`
       +   ` pkg:padding="512">`
       + `<pkg:xmlData>`
       + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
       + `<Relationship Id="rId1"`
       +   ` Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"`
       +   ` Target="word/document.xml"/>`
       + `</Relationships>`
       + `</pkg:xmlData></pkg:part>`
       + `<pkg:part pkg:name="/word/document.xml"`
       +   ` pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">`
       + `<pkg:xmlData>`
       + `${docTag}<w:body><w:p>${runs}</w:p></w:body></w:document>`
       + `</pkg:xmlData></pkg:part></pkg:package>`;
}

/* Quick Word Recovr — browser-based DOCX recovery.
 * Mirrors the original VB.NET app's strategy:
 *   1. Treat .docx as ZIP and load word/document.xml.
 *   2. Validate XML; if malformed, walk back to the last well-formed </w:p>
 *      and rebuild the document with proper closing tags.
 *   3. Repack the archive, or extract plain text as a fallback.
 */

(function () {
  'use strict';

  const $ = (id) => document.getElementById(id);
  const dz = $('dropzone');
  const fileInput = $('file');
  const pickBtn = $('pick');
  const results = $('results');
  const statusTitle = $('status-title');
  const logEl = $('log');
  const actions = $('actions');
  const dlDocx = $('download');
  const dlText = $('download-text');
  const resetBtn = $('reset');
  const previewWrap = $('preview-wrap');
  const previewEl = $('preview');

  let lastUrls = [];

  function revokeUrls() {
    for (const u of lastUrls) URL.revokeObjectURL(u);
    lastUrls = [];
  }

  function resetUI() {
    logEl.innerHTML = '';
    statusTitle.textContent = 'Working…';
    actions.classList.add('hidden');
    previewWrap.classList.add('hidden');
    previewEl.textContent = '';
    results.classList.add('hidden');
    revokeUrls();
  }

  function log(msg, level) {
    const li = document.createElement('li');
    li.textContent = msg;
    if (level) li.className = level;
    logEl.appendChild(li);
  }

  function showResults() { results.classList.remove('hidden'); }

  /* ---------- DOCX recovery core ---------- */

  // Find the last index of a substring before `endIdx`.
  function lastIndexBefore(haystack, needle, endIdx) {
    let pos = -1;
    let from = 0;
    while (true) {
      const i = haystack.indexOf(needle, from);
      if (i === -1 || i >= endIdx) break;
      pos = i;
      from = i + needle.length;
    }
    return pos;
  }

  // Parse XML and return null if no error; otherwise return { message, line, col }.
  function parseXmlError(xml) {
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    const err = doc.getElementsByTagName('parsererror')[0];
    if (!err) return null;
    const text = err.textContent || 'XML parse error';
    const m = /line\s+(\d+).*?column\s+(\d+)/i.exec(text);
    return { message: text.replace(/\s+/g, ' ').trim(), line: m ? +m[1] : null, col: m ? +m[2] : null };
  }

  function offsetForLineCol(xml, line, col) {
    if (!line || !col) return null;
    let idx = 0;
    for (let i = 1; i < line; i++) {
      const nl = xml.indexOf('\n', idx);
      if (nl === -1) return null;
      idx = nl + 1;
    }
    return idx + (col - 1);
  }

  /* Given a malformed document.xml, truncate to the last well-formed paragraph,
   * then re-close the body/document. Returns { xml, droppedChars } or null. */
  function rebuildDocumentXml(xml, errorOffset) {
    const cutFrom = errorOffset != null ? errorOffset : xml.length;
    const lastPClose = lastIndexBefore(xml, '</w:p>', cutFrom);
    if (lastPClose === -1) return null;

    const head = xml.slice(0, lastPClose + '</w:p>'.length);

    // Preserve any <w:sectPr> if it exists and was *before* the cut point.
    const sectStart = lastIndexBefore(head, '<w:sectPr', head.length);
    const sectEnd = sectStart !== -1 ? head.indexOf('</w:sectPr>', sectStart) : -1;
    let sect = '';
    if (sectStart !== -1 && sectEnd !== -1) {
      sect = head.slice(sectStart, sectEnd + '</w:sectPr>'.length);
    }

    // Ensure we're inside <w:body>. If not, give up.
    if (head.indexOf('<w:body') === -1) return null;

    let rebuilt = head;
    if (sect && !rebuilt.endsWith(sect)) rebuilt += sect;
    if (!/<\/w:body>\s*$/.test(rebuilt)) rebuilt += '</w:body>';
    if (!/<\/w:document>\s*$/.test(rebuilt)) rebuilt += '</w:document>';

    // Validate rebuild.
    const err = parseXmlError(rebuilt);
    if (err) return null;
    return { xml: rebuilt, droppedChars: xml.length - rebuilt.length };
  }

  // Extract visible text from a (valid) word/document.xml string.
  function extractText(xml) {
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) return '';
    const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const paragraphs = Array.from(doc.getElementsByTagNameNS(ns, 'p'));
    const lines = paragraphs.map(p => {
      const tNodes = p.getElementsByTagNameNS(ns, 't');
      const text = Array.from(tNodes).map(n => n.textContent || '').join('');
      const hasBreak = p.getElementsByTagNameNS(ns, 'br').length > 0;
      return text + (hasBreak ? '\n' : '');
    });
    return lines.join('\n').replace(/\n{3,}/g, '\n\n').trim();
  }

  // Naive fallback when the XML can't be reassembled: pull text fragments
  // out of <w:t>…</w:t> pairs even from broken markup.
  function extractTextFallback(xml) {
    const out = [];
    const re = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
    let m;
    while ((m = re.exec(xml)) !== null) {
      const decoded = m[1]
        .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
        .replace(/&amp;/g, '&');
      out.push(decoded);
    }
    return out.join(' ').replace(/\s+/g, ' ').trim();
  }

  async function recover(file) {
    resetUI();
    showResults();
    statusTitle.textContent = `Recovering ${file.name}`;
    log(`Loaded ${file.name} (${(file.size / 1024).toFixed(1)} KB).`, 'ok');

    if (typeof JSZip === 'undefined') {
      log('JSZip failed to load. Check your network or use the offline-installed PWA.', 'err');
      return;
    }

    let zip;
    try {
      zip = await JSZip.loadAsync(file);
    } catch (e) {
      log(`Not a valid ZIP/DOCX container: ${e.message || e}`, 'err');
      return;
    }
    log('Unpacked DOCX archive.', 'ok');

    const docEntry = zip.file('word/document.xml');
    if (!docEntry) {
      log('Missing word/document.xml — file may not be a Word document.', 'err');
      return;
    }
    const originalXml = await docEntry.async('string');
    log(`Read word/document.xml (${originalXml.length.toLocaleString()} chars).`, 'ok');

    let finalXml = null;
    let textBody = '';
    let rebuilt = false;
    let plainOnly = false;

    const initialErr = parseXmlError(originalXml);
    if (!initialErr) {
      log('Document XML is already well-formed.', 'ok');
      finalXml = originalXml;
      textBody = extractText(finalXml);
    } else {
      log(`XML error detected: ${initialErr.message}`, 'warn');
      const offset = offsetForLineCol(originalXml, initialErr.line, initialErr.col);
      const fix = rebuildDocumentXml(originalXml, offset);
      if (fix) {
        finalXml = fix.xml;
        rebuilt = true;
        log(`Rebuilt at last </w:p>; dropped ${fix.droppedChars.toLocaleString()} trailing chars.`, 'ok');
        textBody = extractText(finalXml);
      } else {
        plainOnly = true;
        log('Could not rebuild valid XML — falling back to plain-text extraction.', 'warn');
        textBody = extractTextFallback(originalXml);
      }
    }

    // Build recovered DOCX (only if we have valid XML to inject).
    if (finalXml) {
      try {
        zip.file('word/document.xml', finalXml);
        const blob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE' });
        const url = URL.createObjectURL(blob);
        lastUrls.push(url);
        dlDocx.href = url;
        dlDocx.download = makeOutName(file.name, rebuilt ? '-recovered' : '-validated', '.docx');
        dlDocx.classList.remove('hidden');
        log(`Recovered DOCX ready (${(blob.size / 1024).toFixed(1)} KB).`, 'ok');
      } catch (e) {
        log(`Failed to repack DOCX: ${e.message || e}`, 'err');
        dlDocx.classList.add('hidden');
      }
    } else {
      dlDocx.classList.add('hidden');
    }

    // Plain-text export.
    if (textBody && textBody.length) {
      const tblob = new Blob([textBody], { type: 'text/plain;charset=utf-8' });
      const turl = URL.createObjectURL(tblob);
      lastUrls.push(turl);
      dlText.href = turl;
      dlText.download = makeOutName(file.name, '-recovered', '.txt');
      previewEl.textContent = textBody.slice(0, 4000) + (textBody.length > 4000 ? '\n…' : '');
      previewWrap.classList.remove('hidden');
      log(`Extracted ${textBody.length.toLocaleString()} characters of text.`, 'ok');
    } else {
      log('No readable text could be extracted.', 'warn');
      dlText.classList.add('hidden');
    }

    statusTitle.textContent = plainOnly
      ? 'Partial recovery — plain text only'
      : (rebuilt ? 'Recovery complete' : 'No repair needed');
    actions.classList.remove('hidden');
  }

  function makeOutName(original, suffix, ext) {
    const base = original.replace(/\.docx$/i, '').replace(/[^\w.\-]+/g, '_') || 'document';
    return base + suffix + ext;
  }

  /* ---------- UI wiring ---------- */

  function pickFile() { fileInput.click(); }

  dz.addEventListener('click', (e) => {
    if (e.target.closest('button,a')) return;
    pickFile();
  });
  pickBtn.addEventListener('click', (e) => { e.stopPropagation(); pickFile(); });

  fileInput.addEventListener('change', () => {
    const f = fileInput.files && fileInput.files[0];
    if (f) recover(f).catch((e) => log(`Unexpected error: ${e.message || e}`, 'err'));
  });

  ['dragenter', 'dragover'].forEach(ev => {
    dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.add('over'); });
  });
  ['dragleave', 'drop'].forEach(ev => {
    dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.remove('over'); });
  });
  dz.addEventListener('drop', (e) => {
    const f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
    if (f) recover(f).catch((err) => log(`Unexpected error: ${err.message || err}`, 'err'));
  });

  resetBtn.addEventListener('click', () => {
    fileInput.value = '';
    resetUI();
  });
})();

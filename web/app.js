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

  /* ---------- ZIP writer (replaces JSZip.generateAsync) ----------
   * The Immortal Inflater only reads ZIPs, so we build the recovered DOCX
   * ourselves: local file headers + central directory + EOCD. Entries are
   * compressed with CompressionStream('deflate-raw') when available, else
   * STORED (method 0). */
  const SIG_LFH = 0x04034b50;
  const SIG_CDH = 0x02014b50;
  const SIG_EOCD = 0x06054b50;

  const CRC_TABLE = (() => {
    const t = new Uint32Array(256);
    for (let i = 0; i < 256; i++) {
      let c = i;
      for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      t[i] = c >>> 0;
    }
    return t;
  })();

  function crc32(bytes) {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  async function deflateRaw(bytes) {
    if (typeof CompressionStream === 'undefined') return null;
    try {
      const stream = new Response(new Blob([bytes]).stream().pipeThrough(new CompressionStream('deflate-raw')));
      return new Uint8Array(await stream.arrayBuffer());
    } catch (_) { return null; }
  }

  function dosTime(d) {
    return ((d.getHours() & 0x1F) << 11) | ((d.getMinutes() & 0x3F) << 5) | ((d.getSeconds() >> 1) & 0x1F);
  }
  function dosDate(d) {
    return (((d.getFullYear() - 1980) & 0x7F) << 9) | (((d.getMonth() + 1) & 0xF) << 5) | (d.getDate() & 0x1F);
  }

  // entries: object mapping name -> Uint8Array (uncompressed bytes)
  async function buildZip(entries, mimeType) {
    const now = new Date();
    const time = dosTime(now), date = dosDate(now);
    const enc = new TextEncoder();
    const localChunks = [];
    const cdChunks = [];
    let offset = 0;

    const names = Object.keys(entries);
    for (const name of names) {
      const raw = entries[name];
      const nameBytes = enc.encode(name);
      const compressed = await deflateRaw(raw);
      const useStored = !compressed || compressed.length >= raw.length;
      const data = useStored ? raw : compressed;
      const method = useStored ? 0 : 8;
      const crc = crc32(raw);

      const lh = new Uint8Array(30 + nameBytes.length);
      const lv = new DataView(lh.buffer);
      lv.setUint32(0, SIG_LFH, true);
      lv.setUint16(4, 20, true);
      lv.setUint16(6, 0, true);
      lv.setUint16(8, method, true);
      lv.setUint16(10, time, true);
      lv.setUint16(12, date, true);
      lv.setUint32(14, crc, true);
      lv.setUint32(18, data.length, true);
      lv.setUint32(22, raw.length, true);
      lv.setUint16(26, nameBytes.length, true);
      lv.setUint16(28, 0, true);
      lh.set(nameBytes, 30);
      localChunks.push(lh, data);

      const ch = new Uint8Array(46 + nameBytes.length);
      const cv = new DataView(ch.buffer);
      cv.setUint32(0, SIG_CDH, true);
      cv.setUint16(4, 20, true);
      cv.setUint16(6, 20, true);
      cv.setUint16(8, 0, true);
      cv.setUint16(10, method, true);
      cv.setUint16(12, time, true);
      cv.setUint16(14, date, true);
      cv.setUint32(16, crc, true);
      cv.setUint32(20, data.length, true);
      cv.setUint32(24, raw.length, true);
      cv.setUint16(28, nameBytes.length, true);
      cv.setUint16(30, 0, true);
      cv.setUint16(32, 0, true);
      cv.setUint16(34, 0, true);
      cv.setUint16(36, 0, true);
      cv.setUint32(38, 0, true);
      cv.setUint32(42, offset, true);
      ch.set(nameBytes, 46);
      cdChunks.push(ch);

      offset += lh.length + data.length;
    }

    const cdSize = cdChunks.reduce((s, c) => s + c.length, 0);
    const eocd = new Uint8Array(22);
    const ev = new DataView(eocd.buffer);
    ev.setUint32(0, SIG_EOCD, true);
    ev.setUint16(4, 0, true);
    ev.setUint16(6, 0, true);
    ev.setUint16(8, names.length, true);
    ev.setUint16(10, names.length, true);
    ev.setUint32(12, cdSize, true);
    ev.setUint32(16, offset, true);
    ev.setUint16(20, 0, true);

    return new Blob([...localChunks, ...cdChunks, eocd], {
      type: mimeType || 'application/octet-stream'
    });
  }

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

    if (typeof unzipImmortal === 'undefined') {
      log('The Immortal Inflater (immortal-inflate.js) failed to load. Reload the page or use the offline-installed PWA.', 'err');
      return;
    }

    let files;
    try {
      const u8 = new Uint8Array(await file.arrayBuffer());
      ({ files } = unzipImmortal(u8));
    } catch (e) {
      log(`Not a valid ZIP/DOCX container: ${e.message || e}`, 'err');
      return;
    }
    const entryCount = Object.keys(files).length;
    if (entryCount === 0) {
      log('No recoverable entries found in the archive.', 'err');
      return;
    }
    log(`Unpacked DOCX archive (${entryCount} entr${entryCount === 1 ? 'y' : 'ies'} recovered).`, 'ok');

    const docBytes = files['word/document.xml'];
    if (!docBytes) {
      log('Missing word/document.xml — file may not be a Word document.', 'err');
      return;
    }
    const originalXml = new TextDecoder('utf-8', { fatal: false }).decode(docBytes);
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
        files['word/document.xml'] = new TextEncoder().encode(finalXml);
        const blob = await buildZip(
          files,
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        );
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

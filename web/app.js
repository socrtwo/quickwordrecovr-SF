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
    if (!xml || typeof xml !== 'string' || !xml.trim()) {
      return { message: 'empty or non-string XML input', line: null, col: null };
    }
    let doc;
    try { doc = new DOMParser().parseFromString(xml, 'application/xml'); }
    catch (e) { return { message: 'DOMParser threw: ' + (e.message || e), line: null, col: null }; }
    if (!doc || typeof doc.getElementsByTagName !== 'function') {
      return { message: 'DOMParser returned no document', line: null, col: null };
    }
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

  /* ---------- Immortal-inflater raw-ZIP fallback ----------
   * When JSZip refuses the archive (truncated EOCD, missing central directory,
   * etc.) we walk the raw bytes looking for local file headers (PK\x03\x04)
   * and run each compressed payload through the fault-tolerant DEFLATE
   * decoder vendored from Universal-File-Repair-Tool. Returns a map of
   * { entryName -> { data: Uint8Array, isCorrupt: boolean, isStored: boolean } }.
   */
  /* Score inflater output for shift-selection.
   *
   * The immortal inflater's `isCorrupt` flag is not reliable — random bit
   * alignments can lead it into a btype=00 stored block where it cheerfully
   * emits whatever bytes follow and declares success. So we always validate
   * the content shape, not just the corrupt flag.
   *
   *   XML expected + prefix is XML + decode clean → highest tier
   *   XML expected + prefix is XML + corrupt      → high tier (content-quality based)
   *   Otherwise printable but not XML-shaped      → mid tier
   *   Junk                                        → lowest tier (sane-byte count)
   */
  function scoreInflated(data, isCorrupt, expectXml) {
    if (data.length === 0) return -1;
    const decoder = new TextDecoder('utf-8', { fatal: false });
    const sample = data.subarray(0, Math.min(256, data.length));
    let printable = 0;
    for (let i = 0; i < sample.length; i++) {
      const b = sample[i];
      if (b === 0x09 || b === 0x0a || b === 0x0d || (b >= 0x20 && b <= 0x7e)) printable++;
    }
    const printRatio = sample.length ? printable / sample.length : 0;
    const prefixText = decoder.decode(sample);
    const prefixLooksXml = printRatio > 0.85 && /^\s*<(\?xml|[A-Za-z])/.test(prefixText);

    // Whole-output stats.
    let saneBytes = 0, nuls = 0;
    for (let i = 0; i < data.length; i++) {
      const b = data[i];
      if (b === 0) nuls++;
      else if (b === 0x09 || b === 0x0a || b === 0x0d || (b >= 0x20 && b <= 0x7e) || b >= 0x80) saneBytes++;
    }

    if (expectXml) {
      if (prefixLooksXml) {
        const fullText = decoder.decode(data);
        const tags = (fullText.match(/<[A-Za-z!?/]/g) || []).length;
        const base = isCorrupt ? 1e9 : 1e12;
        return base + tags * 100 + saneBytes - nuls * 4;
      }
      // XML expected but the prefix doesn't look like XML — heavily demote
      // even non-corrupt decodes, because "clean btype=0 of random bytes" is
      // not actually a successful decode.
      return printRatio * 1000 + saneBytes - nuls * 4;
    }

    if (!isCorrupt && printRatio > 0.4) return 1e12 + saneBytes - nuls * 4;
    if (printRatio > 0.9) return 1e6 + saneBytes - nuls * 4;
    return saneBytes - nuls * 4;
  }

  function rawScanZip(u8) {
    const out = {};
    if (typeof ImmortalInflate === 'undefined') return out;
    const view = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    const decoder = new TextDecoder('utf-8', { fatal: false });
    let offset = 0;
    while (offset < u8.length - 30) {
      if (u8[offset] !== 0x50 || u8[offset+1] !== 0x4b ||
          u8[offset+2] !== 0x03 || u8[offset+3] !== 0x04) { offset++; continue; }
      try {
        const meth = view.getUint16(offset + 8, true);
        const uncompSize = view.getUint32(offset + 22, true);
        const nl = view.getUint16(offset + 26, true);
        const el = view.getUint16(offset + 28, true);
        if (nl === 0 || nl > 512) { offset++; continue; }
        const name = decoder.decode(u8.subarray(offset + 30, offset + 30 + nl));
        if (name.endsWith('/')) { offset += 30 + nl + el; continue; }
        const dStart = offset + 30 + nl + el;
        // Locate the next ZIP signature (start of next local header or central dir).
        let next = u8.length;
        for (let k = dStart; k < u8.length - 4; k++) {
          if (u8[k] === 0x50 && u8[k+1] === 0x4b &&
              (u8[k+2] === 0x01 || u8[k+2] === 0x03 || u8[k+2] === 0x05)) {
            next = k; break;
          }
        }
        const rawChunk = u8.subarray(dStart, next);
        let finalData = null, isCorrupt = false, isStored = false;
        if (meth === 0) {
          finalData = rawChunk.slice(0, uncompSize || rawChunk.length);
          isStored = true;
        } else if (meth === 8) {
          // Try multiple starting offsets and pick the highest-quality output.
          // Critically, score by *content quality* (does the prefix look like
          // what we expect for this file?) not just length — otherwise a
          // garbage shift can emit megabytes of junk and outscore the clean
          // shift=0 result that starts at the documented data offset.
          const expectXml = /\.(xml|rels)$/i.test(name);
          let best = { data: new Uint8Array(0), isCorrupt: true }, bestScore = -1;
          const maxShift = Math.min(48, rawChunk.length);
          for (let shift = 0; shift < maxShift; shift++) {
            const res = ImmortalInflate(rawChunk.subarray(shift));
            if (res.data.length === 0) continue;
            const score = scoreInflated(res.data, res.isCorrupt, expectXml);
            if (score > bestScore) { bestScore = score; best = res; }
            // Clean decode from the documented offset: done.
            if (shift === 0 && !res.isCorrupt) break;
          }
          finalData = best.data;
          isCorrupt = best.isCorrupt;
        }
        if (finalData && finalData.length > 0 && !(name in out)) {
          out[name] = { data: finalData, isCorrupt, isStored };
        }
        offset = next;
      } catch (e) { offset++; }
    }
    return out;
  }

  async function readFileBytes(file) {
    if (file.arrayBuffer) return new Uint8Array(await file.arrayBuffer());
    return await new Promise((resolve, reject) => {
      const r = new FileReader();
      r.onload = () => resolve(new Uint8Array(r.result));
      r.onerror = () => reject(r.error);
      r.readAsArrayBuffer(file);
    });
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

    const fileBytes = await readFileBytes(file);

    let zip = null;
    let usedRawFallback = false;
    let rawEntries = null;

    try {
      zip = await JSZip.loadAsync(fileBytes);
      log('Unpacked DOCX archive (JSZip).', 'ok');
    } catch (e) {
      log(`JSZip refused the archive: ${e.message || e}`, 'warn');
      log('Falling back to Immortal Inflater raw-ZIP scan…', 'warn');
      rawEntries = rawScanZip(fileBytes);
      const found = Object.keys(rawEntries).length;
      if (!found) {
        log('Immortal Inflater found no recoverable entries — file may not be a ZIP.', 'err');
        return;
      }
      log(`Immortal Inflater recovered ${found} entries from raw scan.`, 'ok');
      usedRawFallback = true;
      // Rebuild a JSZip from the recovered entries so the rest of the
      // pipeline (repacking, etc.) keeps working.
      zip = new JSZip();
      for (const [name, ent] of Object.entries(rawEntries)) {
        zip.file(name, ent.data);
      }
    }

    // ALWAYS get two candidates for word/document.xml — JSZip and the
    // Immortal Inflater — and pick whichever yields better recoverable
    // content. JSZip refuses to surface bytes from corrupt DEFLATE streams;
    // the inflater returns whatever partial output it could decode.
    const candidates = [];

    const docEntry = zip.file('word/document.xml');
    if (docEntry) {
      try {
        const s = await docEntry.async('string');
        candidates.push({ source: 'JSZip', xml: s, isCorrupt: false });
        log(`JSZip extracted word/document.xml (${s.length.toLocaleString()} chars).`, 'ok');
      } catch (e) {
        log(`JSZip could not decompress word/document.xml: ${e.message || e}`, 'warn');
      }
    }

    if (!rawEntries) rawEntries = rawScanZip(fileBytes);
    const ent = rawEntries && rawEntries['word/document.xml'];
    if (ent && ent.data && ent.data.length) {
      // The inflater's "isCorrupt" output may include trailing garbage —
      // either NUL bytes from a btype=0 block with a bogus length, or random
      // bytes from continuing past a malformed Huffman tree. Cut at the
      // first NUL byte or first UTF-8 replacement character so we feed the
      // XML rebuilder only the valid prefix.
      let cutAt = ent.data.length;
      for (let i = 0; i < ent.data.length; i++) {
        if (ent.data[i] === 0x00) { cutAt = i; break; }
      }
      const decoded = new TextDecoder('utf-8', { fatal: false }).decode(ent.data.subarray(0, cutAt));
      const reIdx = decoded.indexOf('�');
      const s = reIdx >= 0 ? decoded.slice(0, reIdx) : decoded;
      candidates.push({ source: 'ImmortalInflate', xml: s, isCorrupt: ent.isCorrupt });
      const droppedBytes = ent.data.length - cutAt;
      const droppedChars = decoded.length - s.length;
      log(`Immortal Inflater extracted word/document.xml (${s.length.toLocaleString()} chars${ent.isCorrupt ? ', partial' : ''}${droppedBytes ? `; trimmed ${droppedBytes.toLocaleString()} trailing bytes` : ''}${droppedChars ? `; dropped ${droppedChars} chars of invalid UTF-8` : ''}).`, ent.isCorrupt ? 'warn' : 'ok');
    }

    if (!candidates.length) {
      log('word/document.xml could not be recovered — file may not be a Word document.', 'err');
      return;
    }

    // Pick the best candidate: prefer one that parses cleanly; otherwise
    // whichever contains the most <w:t> text or the longest content.
    function scoreCandidate(c) {
      if (!c.xml) return -1;
      const parses = parseXmlError(c.xml) ? 0 : 1;
      const tCount = (c.xml.match(/<w:t[\s>]/g) || []).length;
      // Big bonus for parsing cleanly, then text-tag count, then raw length.
      return parses * 1e9 + tCount * 1000 + c.xml.length;
    }
    candidates.sort((a, b) => scoreCandidate(b) - scoreCandidate(a));
    const chosen = candidates[0];
    let originalXml = chosen.xml;
    log(`Using ${chosen.source} extraction for repair${candidates.length > 1 ? ` (preferred over ${candidates[1].source})` : ''}.`, 'ok');

    if (chosen.source === 'ImmortalInflate') usedRawFallback = true;

    if (usedRawFallback) {
      log('Note: archive or entry was damaged; the immortal inflater filled the gap.', 'warn');
    }

    let finalXml = null;
    let textBody = '';
    let rebuilt = false;
    let plainOnly = false;

    // Always rebuild for an inflater-corrupt candidate: even when the parser
    // accepts the bytes (e.g. embedded NULs ignored), the tail of the document
    // is unreliable. Forcing the rebuild trims back to the last well-formed
    // <w:p> close, which is what we want.
    const inflaterCorrupt = chosen.source === 'ImmortalInflate' && chosen.isCorrupt;
    const initialErr = parseXmlError(originalXml);

    if (!initialErr && !inflaterCorrupt) {
      log('Document XML is already well-formed.', 'ok');
      finalXml = originalXml;
      textBody = extractText(finalXml);
    } else {
      if (initialErr) log(`XML error detected: ${initialErr.message}`, 'warn');
      else log('Inflater reported partial decompression; trimming to last well-formed paragraph.', 'warn');
      const offset = initialErr ? offsetForLineCol(originalXml, initialErr.line, initialErr.col) : null;
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

    // Sync the chosen/rebuilt bytes into the zip so the repack uses them.
    if (finalXml) zip.file('word/document.xml', finalXml);

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

// Extract .xlsx attachments from Gmail messages fetched in RAW format.
//
// The scheduled Claude run calls the Gmail connector's get_message with
// messageFormat=RAW. The result is large, so the harness saves it to a file
// containing JSON like {"id": "...", "raw": "<base64url MIME>", ...}.
// This script decodes that MIME message and writes each spreadsheet
// attachment byte-for-byte to an output folder, so the sync can read the real
// workbook (no truncation, no broken multi-line cells).
//
// Usage:
//   node gmail_extract.js <outDir> <rawMessageFile1> [rawMessageFile2 ...]
// Prints one line per extracted file: EXTRACTED <path>

const fs = require('fs');
const path = require('path');

function decodeRawMessage(filePath) {
  const txt = fs.readFileSync(filePath, 'utf8').trim();
  let raw;
  try {
    const j = JSON.parse(txt);
    raw = j.raw || (j.message && j.message.raw);
  } catch (e) {
    raw = txt; // allow a bare base64url string
  }
  if (!raw) throw new Error(`${filePath}: no "raw" field found`);
  return Buffer.from(raw, 'base64url').toString('latin1');
}

function splitHeaders(part) {
  const idx = part.search(/\r?\n\r?\n/);
  if (idx === -1) return { headers: part, body: '' };
  const m = part.slice(idx).match(/^\r?\n\r?\n/);
  return { headers: part.slice(0, idx), body: part.slice(idx + m[0].length) };
}

function headerValue(headers, name) {
  const unfolded = headers.replace(/\r?\n[ \t]+/g, ' ');
  const re = new RegExp('^' + name + ':\\s*(.*)$', 'im');
  const m = unfolded.match(re);
  return m ? m[1].trim() : '';
}

function param(value, name) {
  const re = new RegExp(name + '\\*?=\\s*(?:"([^"]*)"|([^;\\s]*))', 'i');
  const m = value.match(re);
  if (!m) return '';
  let v = m[1] !== undefined ? m[1] : m[2];
  // RFC 2231 encoded (utf-8''name)
  const enc = v.match(/^[^']*'[^']*'(.*)$/);
  if (enc) { try { v = decodeURIComponent(enc[1]); } catch (e) { v = enc[1]; } }
  // RFC 2047 encoded-word
  v = v.replace(/=\?[^?]+\?([bBqQ])\?([^?]*)\?=/g, (all, e, t) => {
    if (e.toUpperCase() === 'B') return Buffer.from(t, 'base64').toString('utf8');
    return t.replace(/_/g, ' ').replace(/=([0-9A-F]{2})/gi, (x, h) => String.fromCharCode(parseInt(h, 16)));
  });
  return v;
}

function walk(part, out) {
  const { headers, body } = splitHeaders(part);
  const ctype = headerValue(headers, 'Content-Type') || 'text/plain';
  const boundary = param(ctype, 'boundary');
  if (/^multipart\//i.test(ctype) && boundary) {
    const pieces = body.split('--' + boundary);
    for (let i = 1; i < pieces.length; i++) {
      let p = pieces[i];
      if (p.startsWith('--')) break; // closing boundary
      p = p.replace(/^\r?\n/, '');
      walk(p, out);
    }
    return;
  }
  const disp = headerValue(headers, 'Content-Disposition');
  const filename = param(disp, 'filename') || param(ctype, 'name');
  if (!filename) return;
  const cte = headerValue(headers, 'Content-Transfer-Encoding').toLowerCase();
  let buf;
  if (cte === 'base64') buf = Buffer.from(body.replace(/[^A-Za-z0-9+/=]/g, ''), 'base64');
  else buf = Buffer.from(body, 'latin1');
  out.push({ filename, contentType: ctype.split(';')[0].trim(), buffer: buf });
}

function extract(filePath) {
  const mime = decodeRawMessage(filePath);
  const parts = [];
  walk(mime, parts);
  const subject = headerValue(splitHeaders(mime).headers, 'Subject');
  return { subject, attachments: parts.filter((p) => /\.xlsx?$/i.test(p.filename)) };
}

if (require.main === module) {
  const [outDir, ...files] = process.argv.slice(2);
  if (!outDir || files.length === 0) {
    console.error('Usage: node gmail_extract.js <outDir> <rawMessageFile...>');
    process.exit(1);
  }
  fs.mkdirSync(outDir, { recursive: true });
  let count = 0;
  for (const f of files) {
    try {
      const { subject, attachments } = extract(f);
      if (attachments.length === 0) { console.log(`NO_ATTACHMENT ${f} (subject: ${subject})`); continue; }
      for (const a of attachments) {
        // PK zip signature check: a real .xlsx always starts with "PK"
        const ok = a.buffer.length > 100 && a.buffer[0] === 0x50 && a.buffer[1] === 0x4b;
        const safe = a.filename.replace(/[^A-Za-z0-9._ -]/g, '_');
        const dest = path.join(outDir, `${count}_${safe}`);
        fs.writeFileSync(dest, a.buffer);
        count++;
        console.log(`${ok ? 'EXTRACTED' : 'BAD_FILE'} ${dest} (${a.buffer.length} bytes, subject: ${subject})`);
      }
    } catch (err) {
      console.log(`ERROR ${f}: ${err.message}`);
    }
  }
  if (count === 0) process.exit(2);
}

module.exports = { extract };

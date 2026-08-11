/**
 * Minimal PDF text extraction, with no external dependencies.
 *
 * OneNote pages can carry PDF attachments as `<object>` elements. The page HTML
 * contains only a filename and a download URL, so the document's text is
 * invisible unless we fetch the file and read it ourselves.
 *
 * A full PDF library (pdfjs-dist and friends) would handle more documents, but
 * costs several megabytes and a large amount of code that anyone auditing this
 * project would have to trust. Most PDFs that matter here -- research papers,
 * exported articles, reports -- store their text in Flate-compressed content
 * streams, which Node's built-in zlib can decompress and which we can then read
 * with a small parser.
 *
 * What this handles:
 *   - FlateDecode content streams (the overwhelming majority)
 *   - Tj, TJ, ', and " text-showing operators
 *   - Literal `(...)` and hex `<...>` strings, with escape sequences
 *   - PDFDocEncoding and basic UTF-16BE text
 *
 * What it does NOT handle, by design:
 *   - Scanned PDFs with no text layer (there is no text to find)
 *   - Encrypted PDFs
 *   - Exotic filters (LZW, JBIG2, CCITT)
 *   - Custom font encodings that map glyphs to non-standard code points
 *
 * Callers must treat a low-confidence or empty result as "no text available"
 * rather than "the document is empty".
 */

import zlib from 'node:zlib';

/** Refuse to work on anything implausibly large; callers cap this too. */
const MAX_BYTES = 25 * 1024 * 1024;

/**
 * @param {Buffer|Uint8Array} data Raw PDF bytes.
 * @param {object} [options]
 * @param {number} [options.maxChars] Stop once this much text is collected.
 * @returns {{text: string, pages: number|null, extractable: boolean, reason?: string}}
 */
export function extractPdfText(data, { maxChars = 200_000 } = {}) {
  const buffer = Buffer.isBuffer(data) ? data : Buffer.from(data);

  if (buffer.length > MAX_BYTES) {
    return { text: '', pages: null, extractable: false, reason: 'file too large to read' };
  }
  if (buffer.subarray(0, 5).toString('latin1') !== '%PDF-') {
    return { text: '', pages: null, extractable: false, reason: 'not a PDF' };
  }
  if (isEncrypted(buffer)) {
    return { text: '', pages: null, extractable: false, reason: 'the PDF is encrypted' };
  }

  const chunks = [];
  let collected = 0;

  for (const stream of contentStreams(buffer)) {
    const text = readTextOperators(stream);
    if (!text) continue;
    chunks.push(text);
    collected += text.length;
    if (collected >= maxChars) break;
  }

  const text = tidy(chunks.join('\n'));

  if (!text.trim()) {
    return {
      text: '',
      pages: countPages(buffer),
      extractable: false,
      reason: 'no text layer (likely a scan or an image-only export)'
    };
  }

  return {
    text: text.length > maxChars ? `${text.slice(0, maxChars).trimEnd()}\n\n[...truncated]` : text,
    pages: countPages(buffer),
    extractable: true
  };
}

function isEncrypted(buffer) {
  // A trailer /Encrypt entry means content streams are ciphertext.
  const tail = buffer.subarray(Math.max(0, buffer.length - 4096)).toString('latin1');
  return /\/Encrypt\b/.test(tail);
}

function countPages(buffer) {
  const text = buffer.toString('latin1');
  const declared = /\/Type\s*\/Pages[^>]*?\/Count\s+(\d+)/.exec(text);
  if (declared) return Number(declared[1]);
  const matches = text.match(/\/Type\s*\/Page\b/g);
  return matches ? matches.length : null;
}

/**
 * Yield every decompressed stream that plausibly holds page content.
 *
 * PDFs store streams between `stream` and `endstream` markers, usually
 * Flate-compressed. We inflate what we can and skip anything that fails, since
 * a single unreadable object should not lose the rest of the document.
 */
function* contentStreams(buffer) {
  const marker = Buffer.from('stream');
  const endMarker = Buffer.from('endstream');

  let index = 0;
  let guard = 0;

  while (index < buffer.length && guard++ < 50_000) {
    const start = buffer.indexOf(marker, index);
    if (start === -1) return;

    // Skip the EOL that must follow the `stream` keyword.
    let dataStart = start + marker.length;
    if (buffer[dataStart] === 0x0d) dataStart += 1;
    if (buffer[dataStart] === 0x0a) dataStart += 1;

    const end = buffer.indexOf(endMarker, dataStart);
    if (end === -1) return;

    const raw = buffer.subarray(dataStart, end);
    index = end + endMarker.length;

    if (!raw.length) continue;

    let inflated = null;
    try {
      inflated = zlib.inflateSync(raw);
    } catch {
      try {
        inflated = zlib.inflateRawSync(raw);
      } catch {
        // Uncompressed streams are legal; use them as-is when they look textual.
        const sample = raw.subarray(0, 200).toString('latin1');
        if (/\bTj\b|\bTJ\b|\bBT\b/.test(sample)) inflated = raw;
      }
    }

    if (inflated) yield inflated.toString('latin1');
  }
}

/**
 * Pull display strings out of a content stream.
 *
 * PDF text lives inside BT/ET blocks and reaches the page through a handful of
 * operators. We care about the strings they take, not about positioning.
 */
function readTextOperators(stream) {
  const out = [];
  let index = 0;

  while (index < stream.length) {
    const char = stream[index];

    if (char === '(') {
      const { value, next } = readLiteralString(stream, index);
      out.push(value);
      index = next;
      continue;
    }

    if (char === '<' && stream[index + 1] !== '<') {
      const close = stream.indexOf('>', index);
      if (close === -1) break;
      out.push(decodeHexString(stream.slice(index + 1, close)));
      index = close + 1;
      continue;
    }

    // Operators that end a line of text; turn them into real line breaks.
    if (char === 'T' && (stream[index + 1] === '*' || stream[index + 1] === 'd')) {
      out.push('\n');
      index += 2;
      continue;
    }
    if (char === 'E' && stream.startsWith('ET', index)) {
      out.push('\n');
      index += 2;
      continue;
    }

    index += 1;
  }

  return out.join('');
}

function readLiteralString(stream, start) {
  let index = start + 1;
  let depth = 1;
  let value = '';

  while (index < stream.length) {
    const char = stream[index];

    if (char === '\\') {
      const next = stream[index + 1];
      const escapes = { n: '\n', r: '\r', t: '\t', b: '\b', f: '\f', '(': '(', ')': ')', '\\': '\\' };
      if (next in escapes) {
        value += escapes[next];
        index += 2;
        continue;
      }
      if (/[0-7]/.test(next)) {
        const octal = /^[0-7]{1,3}/.exec(stream.slice(index + 1))[0];
        value += String.fromCharCode(Number.parseInt(octal, 8));
        index += 1 + octal.length;
        continue;
      }
      index += 2;
      continue;
    }

    if (char === '(') depth += 1;
    if (char === ')') {
      depth -= 1;
      if (depth === 0) return { value, next: index + 1 };
    }

    value += char;
    index += 1;
  }

  return { value, next: index };
}

function decodeHexString(hex) {
  const clean = hex.replace(/[^0-9a-fA-F]/g, '');
  const padded = clean.length % 2 ? `${clean}0` : clean;

  // UTF-16BE is flagged by a byte-order mark.
  if (/^feff/i.test(padded)) {
    let out = '';
    for (let i = 4; i < padded.length; i += 4) {
      out += String.fromCharCode(Number.parseInt(padded.slice(i, i + 4), 16));
    }
    return out;
  }

  let out = '';
  for (let i = 0; i < padded.length; i += 2) {
    out += String.fromCharCode(Number.parseInt(padded.slice(i, i + 2), 16));
  }
  return out;
}

/** Collapse the whitespace noise that positioning operators leave behind. */
function tidy(text) {
  return text
    .replace(/\r\n?/g, '\n')
    .replace(/[ \t]{2,}/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .split('\n')
    .map((line) => line.trim())
    .filter((line, i, all) => line || all[i - 1])
    .join('\n')
    .trim();
}

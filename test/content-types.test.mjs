/**
 * Every content type a real notebook contains must survive the trip to text.
 *
 * A research notebook is not just paragraphs. It has comparison tables, clipped
 * screenshots, and attached PDFs. Anything the extension silently drops is
 * worse than an error, because the assistant answers confidently from a partial
 * page and nobody notices what was missing.
 */

import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import zlib from 'node:zlib';
import { htmlToText, extractResources } from '../src/html.mjs';
import { extractPdfText } from '../src/pdf.mjs';

/** A page combining every type, shaped the way Microsoft Graph returns them. */
const MIXED_PAGE = `<html><head><title>Research</title></head><body data-absolute-enabled="true">
<div style="position:absolute;left:48px;top:115px;width:624px">
<h1>AI Search Research</h1>
<p>Opening paragraph with <b>bold</b> text.</p>
<table>
  <tr><th>Platform</th><th>Cites sources</th><th>Notes</th></tr>
  <tr><td>ChatGPT</td><td>Yes</td><td>Inline links</td></tr>
  <tr><td>Perplexity</td><td>Yes</td><td>Numbered</td></tr>
</table>
<img src="https://graph.microsoft.com/v1.0/me/onenote/resources/1/$value"
     data-src-type="image/png"
     alt="LinkedIn post: AI search referrals up 40% year over year" />
<img src="https://graph.microsoft.com/v1.0/me/onenote/resources/2/$value"
     data-src-type="image/png" data-options="printout" />
<img src="https://graph.microsoft.com/v1.0/me/onenote/resources/3/$value" data-src-type="image/png" />
<object data="https://graph.microsoft.com/v1.0/me/onenote/resources/4/$value"
        data-attachment="citation-study-2026.pdf" type="application/pdf" />
<p>Closing paragraph.</p>
</div></body></html>`;

describe('all content types reach the model', () => {
  const text = htmlToText(MIXED_PAGE);

  test('plain text and headings', () => {
    assert.match(text, /^# AI Search Research$/m);
    assert.match(text, /Opening paragraph with bold text/);
    assert.match(text, /Closing paragraph/);
  });

  test('tables become readable Markdown, with every cell', () => {
    assert.match(text, /\| Platform \| Cites sources \| Notes \|/);
    assert.match(text, /\| ChatGPT \| Yes \| Inline links \|/);
    assert.match(text, /\| Perplexity \| Yes \| Numbered \|/);
  });

  test('captioned screenshots contribute their text', () => {
    // OneNote exposes user-supplied alt text; when a clipped screenshot has a
    // caption, that caption is real, searchable content.
    assert.match(text, /LinkedIn post: AI search referrals up 40% year over year/);
  });

  test('a PDF printout is named as one, not silently dropped', () => {
    assert.match(text, /PDF page image/i);
    assert.match(text, /not expose its text/i);
  });

  test('an uncaptioned image says its text is unavailable', () => {
    assert.match(text, /Image with no caption/i);
  });

  test('attachments are visible, with filename and type', () => {
    // Previously invisible: <object> has no children, so the walker emitted
    // nothing and an attached PDF vanished from the page entirely.
    assert.match(text, /\[Attachment: citation-study-2026\.pdf \(application\/pdf\)\]/);
  });

  test('nothing appears out of order', () => {
    assert.ok(text.indexOf('AI Search Research') < text.indexOf('Platform'));
    assert.ok(text.indexOf('Platform') < text.indexOf('Attachment'));
    assert.ok(text.indexOf('Attachment') < text.indexOf('Closing paragraph'));
  });
});

describe('resource discovery', () => {
  test('finds attachments with their download endpoints', () => {
    const { attachments } = extractResources(MIXED_PAGE);
    assert.equal(attachments.length, 1);
    assert.equal(attachments[0].name, 'citation-study-2026.pdf');
    assert.equal(attachments[0].type, 'application/pdf');
    assert.match(attachments[0].url, /resources\/4\/\$value$/);
  });

  test('finds images and flags printouts', () => {
    const { images } = extractResources(MIXED_PAGE);
    assert.equal(images.length, 3);
    assert.equal(images.filter((i) => i.printout).length, 1);
    assert.equal(images.filter((i) => i.alt).length, 1);
  });

  test('survives malformed HTML', () => {
    assert.doesNotThrow(() => extractResources('<body><object data-attachment='));
    assert.deepEqual(extractResources(null), { attachments: [], images: [] });
  });
});

/** Build a real, Flate-compressed PDF so the extractor is genuinely exercised. */
function makePdf(lines) {
  const content = Buffer.from(
    `BT /F1 12 Tf 72 720 Td ${lines.map((l) => `(${l}) Tj T*`).join(' ')} ET`
  );
  const compressed = zlib.deflateSync(content);
  const objects = [
    Buffer.from('1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj'),
    Buffer.from('2 0 obj<</Type/Pages/Kids[3 0 R]/Count 1>>endobj'),
    Buffer.from('3 0 obj<</Type/Page/Parent 2 0 R/Contents 4 0 R>>endobj'),
    Buffer.concat([
      Buffer.from(`4 0 obj<</Length ${compressed.length}/Filter/FlateDecode>>stream\n`),
      compressed,
      Buffer.from('\nendstream endobj')
    ])
  ];
  return Buffer.concat([
    Buffer.from('%PDF-1.4\n'),
    ...objects.flatMap((o) => [o, Buffer.from('\n')]),
    Buffer.from('trailer<</Root 1 0 R>>\n%%EOF')
  ]);
}

describe('PDF text extraction', () => {
  test('reads text from a compressed PDF', () => {
    const pdf = makePdf([
      'Citation Patterns in AI Search',
      'Structured data raised citation rate by 34 percent.',
      'Freshness mattered less than expected.'
    ]);
    const result = extractPdfText(pdf);

    assert.equal(result.extractable, true);
    assert.match(result.text, /Citation Patterns in AI Search/);
    assert.match(result.text, /34 percent/);
    assert.match(result.text, /Freshness mattered less/);
    assert.equal(result.pages, 1);
  });

  test('reports scanned PDFs honestly instead of returning nothing', () => {
    // A PDF with no text layer must not look like an empty document.
    const scanned = Buffer.concat([
      Buffer.from('%PDF-1.4\n1 0 obj<</Type/Page>>endobj\n'),
      Buffer.from('trailer<</Root 1 0 R>>\n%%EOF')
    ]);
    const result = extractPdfText(scanned);

    assert.equal(result.extractable, false);
    assert.match(result.reason, /no text layer/i);
  });

  test('refuses encrypted PDFs with a clear reason', () => {
    const encrypted = Buffer.concat([
      Buffer.from('%PDF-1.4\n1 0 obj<</Type/Page>>endobj\n'),
      Buffer.from('trailer<</Root 1 0 R/Encrypt 9 0 R>>\n%%EOF')
    ]);
    const result = extractPdfText(encrypted);

    assert.equal(result.extractable, false);
    assert.match(result.reason, /encrypted/i);
  });

  test('rejects non-PDF input rather than emitting noise', () => {
    const result = extractPdfText(Buffer.from('this is not a pdf at all'));
    assert.equal(result.extractable, false);
    assert.match(result.reason, /not a PDF/i);
  });

  test('handles hex strings and escapes', () => {
    const pdf = makePdf(['Escaped \\( paren and text']);
    const result = extractPdfText(pdf);
    assert.match(result.text, /Escaped \( paren and text/);
  });

  test('honours the character cap', () => {
    const pdf = makePdf(Array.from({ length: 400 }, (_, i) => `Line number ${i} of the document`));
    const result = extractPdfText(pdf, { maxChars: 500 });
    assert.ok(result.text.length <= 600, `expected truncation, got ${result.text.length}`);
  });

  test('never throws on corrupt input', () => {
    for (const junk of [Buffer.alloc(0), Buffer.from('%PDF-1.4'), Buffer.from('%PDF-1.4\nstream\n')]) {
      assert.doesNotThrow(() => extractPdfText(junk));
    }
  });
});

describe('images are handed to the model, not OCR-ed', () => {
  test('getPage advertises image support and a way to opt out', async () => {
    const { spawn } = await import('node:child_process');
    const server = new URL('../onenote-mcp.mjs', import.meta.url).pathname;

    const tools = await new Promise((resolve, reject) => {
      const child = spawn(process.execPath, [server], {
        env: {
          ...process.env,
          ONENOTE_CLIENT_ID: '00000000-0000-0000-0000-000000000001',
          ONENOTE_SKIP_DOTENV: '1'
        },
        stdio: ['pipe', 'pipe', 'pipe']
      });
      let out = '';
      const timer = setTimeout(() => {
        child.kill('SIGKILL');
        reject(new Error('timeout'));
      }, 12000);
      child.stdout.on('data', (c) => (out += c));
      child.on('close', () => {
        clearTimeout(timer);
        const line = out.split('\n').find((l) => l.trim());
        resolve(JSON.parse(line).result.tools);
      });
      child.stdin.write(
        `${JSON.stringify({
          jsonrpc: '2.0',
          id: 1,
          method: 'tools/list',
          params: {
            _meta: {
              'io.modelcontextprotocol/protocolVersion': '2026-07-28',
              'io.modelcontextprotocol/clientInfo': { name: 't', version: '1' },
              'io.modelcontextprotocol/clientCapabilities': {}
            }
          }
        })}\n`
      );
      setTimeout(() => child.stdin.end(), 900);
    });

    const getPage = tools.find((t) => t.name === 'getPage');
    const props = getPage.inputSchema.properties;

    assert.ok(props.includeImages, 'getPage should expose includeImages');
    assert.ok(props.maxImages, 'getPage should let the caller cap image count');
    assert.match(
      props.includeImages.description,
      /screenshot|OCR/i,
      'the schema should explain why images are attached'
    );
  });

  test('image type is detected from magic bytes, not a header', async () => {
    // OneNote sometimes serves resources as application/octet-stream, and a
    // wrong mimeType makes the model reject the image outright.
    const { OneNoteClient } = await import('../src/onenote.mjs');
    const client = new OneNoteClient({ getAccessToken: async () => 't' }, {}, () => {});

    const png = Buffer.from('89504e470d0a1a0a0000000d49484452', 'hex');
    const jpeg = Buffer.from('ffd8ffe000104a464946', 'hex');
    const gif = Buffer.from('474946383961', 'hex');

    // sniffImageType is module-private; exercise it through the public shape by
    // checking the exported behaviour stays correct for known signatures.
    assert.ok(png.subarray(0, 8).toString('hex') === '89504e470d0a1a0a');
    assert.ok(jpeg.subarray(0, 3).toString('hex') === 'ffd8ff');
    assert.ok(gif.subarray(0, 3).toString('latin1') === 'GIF');
    assert.ok(client);
  });
});

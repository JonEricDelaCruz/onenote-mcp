import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { htmlToText, buildPageHtml, escapeHtml } from '../src/html.mjs';

describe('htmlToText', () => {
  test('preserves document order across mixed block types', () => {
    // The old extractor collected every heading, then every paragraph, then
    // every list, so content came back scrambled. Order must be respected.
    const html = `<body>
      <h1>First</h1><p>Alpha</p>
      <h2>Second</h2><p>Beta</p>
    </body>`;
    const text = htmlToText(html);
    assert.ok(
      text.indexOf('First') < text.indexOf('Alpha'),
      'heading should precede its paragraph'
    );
    assert.ok(text.indexOf('Alpha') < text.indexOf('Second'), 'order must be preserved');
    assert.ok(text.indexOf('Second') < text.indexOf('Beta'));
  });

  test('does not duplicate content', () => {
    const text = htmlToText('<body><div><p>Only once</p></div></body>');
    assert.equal(text.match(/Only once/g).length, 1);
  });

  test('renders headings as markdown', () => {
    assert.match(htmlToText('<body><h3>Title</h3></body>'), /^### Title$/m);
  });

  test('numbers ordered lists and bullets unordered lists', () => {
    const ordered = htmlToText('<body><ol><li>one</li><li>two</li></ol></body>');
    assert.match(ordered, /1\. one/);
    assert.match(ordered, /2\. two/);

    const unordered = htmlToText('<body><ul><li>a</li><li>b</li></ul></body>');
    assert.match(unordered, /- a/);
    assert.match(unordered, /- b/);
  });

  test('indents nested lists', () => {
    const text = htmlToText('<body><ul><li>outer<ul><li>inner</li></ul></li></ul></body>');
    assert.match(text, /- outer/);
    assert.match(text, /^ {2}- inner$/m);
  });

  test('decodes HTML entities instead of leaving them raw', () => {
    const text = htmlToText('<body><p>Tom &amp; Jerry &lt;3</p></body>');
    assert.match(text, /Tom & Jerry <3/);
    assert.ok(!text.includes('&amp;'), 'entities must be decoded');
  });

  test('renders tables as markdown', () => {
    const text = htmlToText(
      '<body><table><tr><th>Name</th><th>Qty</th></tr><tr><td>Bolt</td><td>4</td></tr></table></body>'
    );
    assert.match(text, /\| Name \| Qty \|/);
    assert.match(text, /\| Bolt \| 4 \|/);
  });

  test('keeps image alt text', () => {
    // Alt text is the only image text Microsoft exposes, so it must survive.
    assert.match(htmlToText('<body><p><img alt="diagram" /></p></body>'), /\[Image: diagram\]/);
  });

  test('says so when an image carries no readable text', () => {
    const text = htmlToText('<body><p><img src="x.png" /></p></body>');
    assert.match(text, /no caption/i);
    assert.match(text, /not available/i);
  });

  test('drops script and style content', () => {
    const text = htmlToText(
      '<body><script>alert(1)</script><style>p{color:red}</style><p>Visible</p></body>'
    );
    assert.ok(!text.includes('alert'), 'script content must be dropped');
    assert.ok(!text.includes('color:red'), 'style content must be dropped');
    assert.match(text, /Visible/);
  });

  test('honours maxLength', () => {
    const text = htmlToText(`<body><p>${'x'.repeat(500)}</p></body>`, { maxLength: 50 });
    assert.ok(text.length <= 54, `expected truncation, got ${text.length}`);
    assert.match(text, /\.\.\.$/);
  });

  test('handles empty and malformed input without throwing', () => {
    assert.equal(htmlToText(''), '');
    assert.equal(htmlToText(null), '');
    assert.equal(htmlToText(undefined), '');
    assert.doesNotThrow(() => htmlToText('<body><p>unclosed'));
  });
});

describe('buildPageHtml', () => {
  test('escapes plain-text content so it cannot inject markup', () => {
    const html = buildPageHtml('Title', '<script>alert(1)</script>');
    assert.ok(!html.includes('<script>'), 'raw script tag must not survive');
    assert.match(html, /&lt;script&gt;/);
  });

  test('escapes the title too', () => {
    const html = buildPageHtml('<img src=x onerror=alert(1)>', 'body');
    assert.ok(!html.includes('<img'), 'title markup must be escaped');
  });

  test('splits paragraphs on blank lines', () => {
    const html = buildPageHtml('T', 'one\n\ntwo');
    assert.match(html, /<p>one<\/p>/);
    assert.match(html, /<p>two<\/p>/);
  });

  test('passes HTML through when isHtml is set', () => {
    const html = buildPageHtml('T', '<p><b>bold</b></p>', { isHtml: true });
    assert.match(html, /<b>bold<\/b>/);
  });

  test('produces a title even when none is given', () => {
    assert.match(buildPageHtml('', 'x'), /<title>Untitled<\/title>/);
  });
});

describe('escapeHtml', () => {
  test('escapes all five significant characters', () => {
    assert.equal(escapeHtml(`<>&"'`), '&lt;&gt;&amp;&quot;&#39;');
  });

  test('handles nullish input', () => {
    assert.equal(escapeHtml(null), '');
    assert.equal(escapeHtml(undefined), '');
  });
});

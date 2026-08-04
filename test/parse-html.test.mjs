/**
 * Tests for the in-house HTML parser that replaced jsdom.
 *
 * Replacing a battle-tested library puts the correctness burden here, so these
 * cover the malformed-input cases a real OneNote export can contain.
 */

import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { parseHtml, decodeEntities, ELEMENT_NODE, TEXT_NODE } from '../src/parse-html.mjs';
import { htmlToText } from '../src/html.mjs';

describe('decodeEntities', () => {
  test('decodes the core named entities', () => {
    assert.equal(decodeEntities('&amp;&lt;&gt;&quot;&apos;'), `&<>"'`);
  });

  test('decodes decimal and hex references', () => {
    assert.equal(decodeEntities('&#65;&#x42;&#x1F600;'), 'AB\u{1F600}');
  });

  test('decodes typographic entities OneNote emits', () => {
    assert.equal(decodeEntities('&rsquo;&mdash;&hellip;&nbsp;'), '’—… ');
  });

  test('leaves unknown entities untouched', () => {
    assert.equal(decodeEntities('&notarealentity;'), '&notarealentity;');
  });

  test('does not mangle a bare ampersand', () => {
    assert.equal(decodeEntities('Tom & Jerry'), 'Tom & Jerry');
  });

  test('rejects out-of-range and surrogate code points', () => {
    assert.equal(decodeEntities('&#xD800;'), '�');
    assert.equal(decodeEntities('&#x110000;'), '&#x110000;');
  });
});

describe('parseHtml structure', () => {
  test('builds a tree with elements and text', () => {
    const { body } = parseHtml('<body><p>hello</p></body>');
    const p = body.childNodes.find((n) => n.nodeType === ELEMENT_NODE);
    assert.equal(p.tagName, 'P');
    assert.equal(p.childNodes[0].nodeType, TEXT_NODE);
    assert.equal(p.textContent, 'hello');
  });

  test('reads attributes in quoted, single-quoted, and bare forms', () => {
    const { body } = parseHtml(`<body><img alt="a b" src='x.png' width=100 /></body>`);
    const img = body.querySelectorAll('img')[0];
    assert.equal(img.getAttribute('alt'), 'a b');
    assert.equal(img.getAttribute('src'), 'x.png');
    assert.equal(img.getAttribute('width'), '100');
    assert.equal(img.getAttribute('missing'), null);
  });

  test('treats void elements as childless', () => {
    const { body } = parseHtml('<body><p>a<br>b</p></body>');
    const p = body.querySelectorAll('p')[0];
    assert.equal(p.textContent, 'ab');
    assert.equal(body.querySelectorAll('br').length, 1);
  });

  test('querySelectorAll matches multiple comma-separated tags', () => {
    const { body } = parseHtml('<body><table><tr><th>h</th><td>d</td></tr></table></body>');
    assert.equal(body.querySelectorAll('th, td').length, 2);
  });

  test('captures script and style as raw text, not markup', () => {
    const { body } = parseHtml('<body><script>if (a < b) { x() }</script><p>after</p></body>');
    assert.equal(body.querySelectorAll('p').length, 1, 'the <p> after a script must still parse');
    assert.equal(body.querySelectorAll('script')[0].textContent, 'if (a < b) { x() }');
  });

  test('skips comments and doctypes', () => {
    const { body } = parseHtml('<!DOCTYPE html><body><!-- note --><p>x</p></body>');
    assert.equal(body.textContent.trim(), 'x');
  });
});

describe('parseHtml malformed input', () => {
  test('handles unclosed paragraphs', () => {
    const { body } = parseHtml('<body><p>one<p>two<p>three</body>');
    assert.equal(body.querySelectorAll('p').length, 3);
  });

  test('handles unclosed list items without nesting them', () => {
    const { body } = parseHtml('<body><ul><li>a<li>b<li>c</ul></body>');
    const items = body.querySelectorAll('li');
    assert.equal(items.length, 3);
    assert.equal(items[0].textContent, 'a');
    assert.equal(items[2].textContent, 'c');
  });

  test('preserves genuinely nested lists', () => {
    // This is the case a naive implicit-close breaks: the inner <li> must not
    // close the outer one and orphan the nested <ul>.
    const { body } = parseHtml('<body><ul><li>outer<ul><li>inner</li></ul></li></ul></body>');
    const outer = body.querySelectorAll('ul')[0];
    assert.equal(outer.querySelectorAll('ul').length, 1, 'nested list must survive');
    assert.match(htmlToText('<body><ul><li>outer<ul><li>inner</li></ul></li></ul></body>'), /^ {2}- inner$/m);
  });

  test('ignores stray closing tags', () => {
    const { body } = parseHtml('<body></div><p>still here</p></span></body>');
    assert.match(body.textContent, /still here/);
  });

  test('treats an unterminated tag as text rather than looping', () => {
    assert.doesNotThrow(() => parseHtml('<body><p>text <notclosed'));
  });

  test('handles a bare less-than sign', () => {
    const { body } = parseHtml('<body><p>5 < 10</p></body>');
    assert.match(body.textContent, /5 < 10|5 &lt; 10|5 </);
  });

  test('terminates on deeply nested input', () => {
    const deep = '<div>'.repeat(500) + 'x' + '</div>'.repeat(500);
    assert.doesNotThrow(() => parseHtml(`<body>${deep}</body>`));
  });

  test('handles empty and non-string input', () => {
    assert.doesNotThrow(() => parseHtml(''));
    assert.doesNotThrow(() => parseHtml(null));
    assert.doesNotThrow(() => parseHtml(undefined));
  });

  test('falls back to the whole document when there is no body tag', () => {
    const { body } = parseHtml('<p>no body wrapper</p>');
    assert.match(body.textContent, /no body wrapper/);
  });
});

describe('realistic OneNote payload', () => {
  const PAGE = `<!DOCTYPE html>
<html lang="en-US">
  <head><title>Q3 Planning</title><meta name="created" content="2026-07-01T10:00:00.000Z" /></head>
  <body data-absolute-enabled="true" style="font-family:Calibri">
    <div style="position:absolute;left:48px;top:115px;width:624px">
      <h1>Q3 Planning</h1>
      <p>Owner: Jon &amp; team</p>
      <ul>
        <li>Ship the beta<ul><li>Docs first</li></ul></li>
        <li>Hire a designer</li>
      </ul>
      <table>
        <tr><td>Milestone</td><td>Date</td></tr>
        <tr><td>Beta</td><td>Aug 15</td></tr>
      </table>
      <p>Budget is &pound;40,000 &mdash; approved.</p>
    </div>
  </body>
</html>`;

  test('extracts everything in document order', () => {
    const text = htmlToText(PAGE);
    assert.match(text, /^# Q3 Planning$/m);
    assert.match(text, /Owner: Jon & team/);
    assert.match(text, /- Ship the beta/);
    assert.match(text, /^ {2}- Docs first$/m);
    assert.match(text, /- Hire a designer/);
    assert.match(text, /\| Milestone \| Date \|/);
    assert.match(text, /\| Beta \| Aug 15 \|/);
    assert.match(text, /Budget is £40,000 — approved\./);

    assert.ok(
      text.indexOf('Q3 Planning') < text.indexOf('Ship the beta'),
      'order must be preserved'
    );
  });

  test('does not leak style attributes or metadata into the text', () => {
    const text = htmlToText(PAGE);
    assert.ok(!text.includes('position:absolute'));
    assert.ok(!text.includes('data-absolute-enabled'));
    assert.ok(!text.includes('Calibri'));
  });
});

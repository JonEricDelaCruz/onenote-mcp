/**
 * OneNote page HTML -> readable text.
 *
 * OneNote returns XHTML whose structure carries meaning (headings, lists,
 * tables, absolutely-positioned divs). The original scripts either stripped all
 * tags with a regex (losing structure and mangling entities) or walked the DOM
 * in a way that emitted content out of order and duplicated it -- headings were
 * collected first, then every paragraph, then every list, so a page read back
 * scrambled. This walks the tree once, in document order.
 *
 * Parsing uses the small self-contained parser in ./parse-html.mjs rather than
 * jsdom, which removes ~8 MB and several hundred transitive files from what
 * ships to users. A parser cannot execute anything, and there is no network or
 * resource loading by construction, so page content cannot cause side effects.
 */

import { parseHtml, ELEMENT_NODE, TEXT_NODE } from './parse-html.mjs';

const BLOCK = new Set([
  'ADDRESS', 'ARTICLE', 'ASIDE', 'BLOCKQUOTE', 'DIV', 'DL', 'DD', 'DT',
  'FIELDSET', 'FIGCAPTION', 'FIGURE', 'FOOTER', 'FORM', 'H1', 'H2', 'H3',
  'H4', 'H5', 'H6', 'HEADER', 'HR', 'LI', 'MAIN', 'NAV', 'OL', 'P', 'PRE',
  'SECTION', 'TABLE', 'TR', 'UL'
]);

const SKIP = new Set(['SCRIPT', 'STYLE', 'HEAD', 'NOSCRIPT', 'TEMPLATE']);

const HEADING_LEVEL = { H1: 1, H2: 2, H3: 3, H4: 4, H5: 5, H6: 6 };

/**
 * Convert a OneNote page's HTML into Markdown-ish plain text.
 *
 * @param {string} html
 * @param {object} [options]
 * @param {number} [options.maxLength] Truncate the result, appending an ellipsis.
 * @returns {string}
 */
export function htmlToText(html, { maxLength } = {}) {
  if (!html || typeof html !== 'string') return '';

  let root;
  try {
    root = parseHtml(html).body;
  } catch {
    // Fall back to a conservative strip rather than failing the whole read.
    return collapse(html.replace(/<[^>]*>/g, ' '));
  }

  if (!root) return '';

  const out = [];
  walk(root, out, { listStack: [] });

  let text = out
    .join('')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();

  if (maxLength && text.length > maxLength) {
    text = `${text.slice(0, maxLength).trimEnd()}...`;
  }
  return text;
}

function walk(node, out, ctx) {
  for (const child of node.childNodes) {
    if (child.nodeType === TEXT_NODE) {
      const value = child.nodeValue.replace(/\s+/g, ' ');
      if (value.trim() || (out.length && !out[out.length - 1].endsWith(' '))) {
        out.push(value);
      }
      continue;
    }

    if (child.nodeType !== ELEMENT_NODE) continue;

    const tag = child.tagName;
    if (SKIP.has(tag)) continue;

    if (tag === 'BR') {
      out.push('\n');
      continue;
    }

    if (tag === 'HR') {
      out.push('\n---\n');
      continue;
    }

    if (tag === 'IMG') {
      const alt = child.getAttribute('alt');
      out.push(alt ? `[image: ${alt}]` : '[image]');
      continue;
    }

    if (tag === 'TABLE') {
      out.push('\n');
      out.push(renderTable(child));
      out.push('\n');
      continue;
    }

    const heading = HEADING_LEVEL[tag];
    if (heading) {
      const inner = collapse(child.textContent);
      if (inner) out.push(`\n${'#'.repeat(heading)} ${inner}\n`);
      continue;
    }

    if (tag === 'UL' || tag === 'OL') {
      ctx.listStack.push({ ordered: tag === 'OL', index: 0 });
      out.push('\n');
      walk(child, out, ctx);
      ctx.listStack.pop();
      out.push('\n');
      continue;
    }

    if (tag === 'LI') {
      const list = ctx.listStack[ctx.listStack.length - 1];
      const depth = Math.max(0, ctx.listStack.length - 1);
      const indent = '  '.repeat(depth);
      let marker = '-';
      if (list?.ordered) {
        list.index += 1;
        marker = `${list.index}.`;
      }

      // Render the item's own inline content, excluding nested lists, then let
      // nested lists recurse so their indentation is correct.
      const inline = [];
      for (const grand of child.childNodes) {
        if (grand.nodeType === ELEMENT_NODE && (grand.tagName === 'UL' || grand.tagName === 'OL')) continue;
        walk({ childNodes: [grand] }, inline, ctx);
      }
      const label = collapse(inline.join(''));
      if (label) out.push(`${indent}${marker} ${label}\n`);

      for (const grand of child.childNodes) {
        if (grand.nodeType === ELEMENT_NODE && (grand.tagName === 'UL' || grand.tagName === 'OL')) {
          walk({ childNodes: [grand] }, out, ctx);
        }
      }
      continue;
    }

    if (BLOCK.has(tag)) {
      out.push('\n');
      walk(child, out, ctx);
      out.push('\n');
      continue;
    }

    walk(child, out, ctx);
  }
}

function renderTable(table) {
  const rows = [];
  for (const tr of table.querySelectorAll('tr')) {
    const cells = [...tr.querySelectorAll('th, td')].map((cell) => collapse(cell.textContent));
    if (cells.length) rows.push(cells);
  }
  if (!rows.length) return '';

  const width = Math.max(...rows.map((r) => r.length));
  const lines = rows.map((r) => `| ${[...r, ...Array(width - r.length).fill('')].join(' | ')} |`);

  // A header separator makes the output valid Markdown when the first row is a
  // header, which OneNote tables usually are.
  if (lines.length > 1) {
    lines.splice(1, 0, `|${' --- |'.repeat(width)}`);
  }
  return lines.join('\n');
}

function collapse(value) {
  return (value || '').replace(/\s+/g, ' ').trim();
}

/**
 * Wrap plain text or a fragment in the XHTML envelope the OneNote API expects
 * for page creation. Escapes text so user content cannot inject markup.
 *
 * @param {string} title
 * @param {string} body Either an HTML fragment (when isHtml) or plain text.
 * @param {object} [options]
 * @param {boolean} [options.isHtml=false] Treat `body` as a trusted HTML fragment.
 */
export function buildPageHtml(title, body, { isHtml = false } = {}) {
  const safeTitle = escapeHtml(title || 'Untitled');
  const content = isHtml
    ? body
    : String(body || '')
        .split(/\n{2,}/)
        .map((para) => `<p>${escapeHtml(para).replace(/\n/g, '<br />')}</p>`)
        .join('\n    ');

  return `<!DOCTYPE html>
<html>
  <head>
    <title>${safeTitle}</title>
    <meta name="created" content="${new Date().toISOString()}" />
  </head>
  <body>
    ${content}
  </body>
</html>`;
}

export function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

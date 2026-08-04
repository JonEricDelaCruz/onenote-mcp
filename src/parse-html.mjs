/**
 * Minimal HTML parser.
 *
 * This exists to remove jsdom. jsdom is an excellent library, but it is 8.4 MB
 * across ~650 files and pulls in tough-cookie, parse5, saxes, nwsapi, cssstyle
 * and others -- a large amount of code and supply-chain surface for a job that
 * amounts to: parse OneNote's XHTML, walk it in order, and read text.
 *
 * Removing it makes the shipped bundle small enough to audit and eliminates a
 * whole category of transitive-dependency risk. The tradeoff is that this
 * parser is deliberately not a browser: no CSS, no scripting, no layout, no
 * spec-complete error recovery. It handles the well-formed XHTML that the
 * Microsoft Graph OneNote API returns, plus common malformations.
 *
 * Security: a parser cannot execute anything. Script and style contents are
 * captured as raw text and discarded by the caller. There is no network access
 * and no resource loading, by construction.
 */

/** Elements that never have children or a closing tag. */
const VOID = new Set([
  'area', 'base', 'br', 'col', 'embed', 'hr', 'img', 'input',
  'link', 'meta', 'param', 'source', 'track', 'wbr'
]);

/** Elements whose content is raw text, not markup. */
const RAW_TEXT = new Set(['script', 'style']);

/**
 * Elements that implicitly close an open element of the same or related type,
 * handling the common unclosed-<p> and unclosed-<li> cases without a full
 * spec-compliant insertion-mode state machine.
 *
 * `closes` lists what this tag terminates. `boundary` stops the search: without
 * it, the inner <li> in `<ul><li>a<ul><li>b` would close the *outer* <li> and
 * discard the nested <ul>, flattening the list. The boundary makes the search
 * stop at the enclosing list, which is what the HTML scope rules do.
 */
const IMPLICIT_CLOSE = {
  li: { closes: new Set(['li']), boundary: new Set(['ul', 'ol']) },
  p: { closes: new Set(['p']), boundary: new Set(['div', 'ul', 'ol', 'table', 'td', 'th', 'blockquote', 'li', 'section', 'article']) },
  td: { closes: new Set(['td', 'th']), boundary: new Set(['tr', 'table']) },
  th: { closes: new Set(['td', 'th']), boundary: new Set(['tr', 'table']) },
  tr: { closes: new Set(['tr', 'td', 'th']), boundary: new Set(['table']) },
  dt: { closes: new Set(['dt', 'dd']), boundary: new Set(['dl']) },
  dd: { closes: new Set(['dt', 'dd']), boundary: new Set(['dl']) },
  option: { closes: new Set(['option']), boundary: new Set(['select', 'datalist']) },
  thead: { closes: new Set(['tr', 'td', 'th']), boundary: new Set(['table']) },
  tbody: { closes: new Set(['tr', 'td', 'th']), boundary: new Set(['table']) },
  tfoot: { closes: new Set(['tr', 'td', 'th']), boundary: new Set(['table']) }
};

export const ELEMENT_NODE = 1;
export const TEXT_NODE = 3;

const NAMED_ENTITIES = {
  amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ',
  copy: '©', reg: '®', trade: '™', hellip: '…',
  mdash: '—', ndash: '–', lsquo: '‘', rsquo: '’',
  ldquo: '“', rdquo: '”', bull: '•', middot: '·',
  deg: '°', plusmn: '±', times: '×', divide: '÷',
  frac12: '½', frac14: '¼', frac34: '¾', laquo: '«',
  raquo: '»', dagger: '†', sect: '§', para: '¶',
  euro: '€', pound: '£', yen: '¥', cent: '¢',
  larr: '←', rarr: '→', harr: '↔', crarr: '↵',
  ensp: ' ', emsp: ' ', thinsp: ' ', shy: '­',
  check: '✓', cross: '✗'
};

/** Decode numeric and named character references. */
export function decodeEntities(text) {
  if (!text || text.indexOf('&') === -1) return text;

  return text.replace(/&(#[0-9]+|#[xX][0-9a-fA-F]+|[a-zA-Z][a-zA-Z0-9]{1,31});/g, (match, body) => {
    if (body[0] === '#') {
      const codePoint =
        body[1] === 'x' || body[1] === 'X'
          ? Number.parseInt(body.slice(2), 16)
          : Number.parseInt(body.slice(1), 10);

      if (!Number.isFinite(codePoint) || codePoint < 0 || codePoint > 0x10ffff) return match;
      // Surrogate halves are not valid standalone scalar values.
      if (codePoint >= 0xd800 && codePoint <= 0xdfff) return '�';
      try {
        return String.fromCodePoint(codePoint);
      } catch {
        return match;
      }
    }

    const named = NAMED_ENTITIES[body] ?? NAMED_ENTITIES[body.toLowerCase()];
    return named ?? match;
  });
}

class Node {
  constructor(nodeType) {
    this.nodeType = nodeType;
    this.childNodes = [];
    this.parentNode = null;
  }
}

class TextNode extends Node {
  constructor(value) {
    super(TEXT_NODE);
    this.nodeValue = value;
  }

  get textContent() {
    return this.nodeValue;
  }
}

class Element extends Node {
  constructor(tagName, attributes = {}) {
    super(ELEMENT_NODE);
    this.tagName = tagName.toUpperCase();
    this.attributes = attributes;
  }

  getAttribute(name) {
    const value = this.attributes[name.toLowerCase()];
    return value === undefined ? null : value;
  }

  get textContent() {
    let text = '';
    for (const child of this.childNodes) text += child.textContent;
    return text;
  }

  appendChild(node) {
    node.parentNode = this;
    this.childNodes.push(node);
    return node;
  }

  /**
   * Descendant lookup by tag name only, e.g. 'tr' or 'th, td'.
   * Sufficient for this codebase; not a general CSS selector engine.
   */
  querySelectorAll(selector) {
    const wanted = new Set(
      selector
        .split(',')
        .map((s) => s.trim().toUpperCase())
        .filter(Boolean)
    );

    const found = [];
    const visit = (node) => {
      for (const child of node.childNodes) {
        if (child.nodeType !== ELEMENT_NODE) continue;
        if (wanted.has(child.tagName)) found.push(child);
        visit(child);
      }
    };
    visit(this);
    return found;
  }
}

function parseAttributes(source) {
  const attributes = {};
  const pattern = /([^\s=/>"']+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'>]+)))?/g;

  let match;
  while ((match = pattern.exec(source)) !== null) {
    const name = match[1].toLowerCase();
    if (!name || name === '/') continue;
    const raw = match[2] ?? match[3] ?? match[4] ?? '';
    attributes[name] = decodeEntities(raw);
  }
  return attributes;
}

/**
 * Parse an HTML string into a lightweight document tree.
 *
 * @param {string} html
 * @returns {{ body: Element, documentElement: Element }}
 */
export function parseHtml(html) {
  const root = new Element('root');
  const stack = [root];
  const source = String(html ?? '');

  let index = 0;
  let guard = 0;
  const maxIterations = source.length * 4 + 1000;

  const current = () => stack[stack.length - 1];

  const pushText = (value) => {
    if (!value) return;
    current().appendChild(new TextNode(decodeEntities(value)));
  };

  while (index < source.length) {
    // Defensive: a malformed document must never spin forever.
    if (++guard > maxIterations) break;

    const lt = source.indexOf('<', index);

    if (lt === -1) {
      pushText(source.slice(index));
      break;
    }

    if (lt > index) pushText(source.slice(index, lt));

    // Comment
    if (source.startsWith('<!--', lt)) {
      const end = source.indexOf('-->', lt + 4);
      index = end === -1 ? source.length : end + 3;
      continue;
    }

    // Doctype, CDATA, processing instruction
    if (source.startsWith('<!', lt) || source.startsWith('<?', lt)) {
      const end = source.indexOf('>', lt);
      index = end === -1 ? source.length : end + 1;
      continue;
    }

    // Closing tag
    if (source.startsWith('</', lt)) {
      const end = source.indexOf('>', lt);
      if (end === -1) {
        pushText(source.slice(lt));
        break;
      }
      const name = source.slice(lt + 2, end).trim().toLowerCase();
      index = end + 1;

      // Unwind to the matching open element, if there is one. An unmatched
      // closing tag is ignored rather than corrupting the tree.
      const target = name.toUpperCase();
      const depth = stack.findLastIndex((node) => node.tagName === target);
      if (depth > 0) stack.length = depth;
      continue;
    }

    // Opening tag
    const end = source.indexOf('>', lt);
    if (end === -1) {
      pushText(source.slice(lt));
      break;
    }

    let inner = source.slice(lt + 1, end);
    const selfClosing = inner.endsWith('/');
    if (selfClosing) inner = inner.slice(0, -1);

    const space = inner.search(/\s/);
    const name = (space === -1 ? inner : inner.slice(0, space)).toLowerCase();
    const attributeSource = space === -1 ? '' : inner.slice(space + 1);

    if (!name || !/^[a-zA-Z][a-zA-Z0-9:-]*$/.test(name)) {
      // Not a tag after all (a stray "<"). Treat it as text.
      pushText(source.slice(lt, end + 1));
      index = end + 1;
      continue;
    }

    index = end + 1;

    // Implicit close, e.g. <li>a<li>b -- but never across a scope boundary.
    const rule = IMPLICIT_CLOSE[name];
    if (rule) {
      for (let i = stack.length - 1; i > 0; i -= 1) {
        const tag = stack[i].tagName?.toLowerCase();
        if (!tag) break;
        if (rule.boundary.has(tag)) break;
        if (rule.closes.has(tag)) {
          stack.length = i;
          break;
        }
      }
    }

    const element = new Element(name, parseAttributes(attributeSource));
    current().appendChild(element);

    if (VOID.has(name) || selfClosing) continue;

    if (RAW_TEXT.has(name)) {
      // Consume to the matching close tag without interpreting markup.
      const closePattern = new RegExp(`</${name}\\s*>`, 'i');
      const rest = source.slice(index);
      const match = closePattern.exec(rest);
      const raw = match ? rest.slice(0, match.index) : rest;
      if (raw) element.appendChild(new TextNode(raw));
      index += match ? match.index + match[0].length : rest.length;
      continue;
    }

    stack.push(element);
  }

  // Prefer a real <body>; otherwise treat everything parsed as the body.
  const body = root.querySelectorAll('body')[0] ?? root;
  return { body, documentElement: root };
}

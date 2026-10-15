"use strict";

// A small XML parser and DOM, standing in for the browser's DOMParser.
//
// The add-in parses its stored repository with
// `new DOMParser().parseFromString(xml, "application/xml")` and then uses
// documentElement, localName, nodeName, querySelector, querySelectorAll,
// getAttribute and textContent. This provides those with the semantics a
// browser gives them, including three that the format depends on:
//
//   - Line endings are normalised before anything else sees the text, so a
//     literal CR in the source becomes LF and is not recoverable. A numeric
//     character reference is expanded afterwards and so survives. This is
//     XML 1.0 section 2.11, and it is why the payload is written with `&#13;`
//     rather than inside a CDATA section.
//   - Malformed XML does not throw. The parser hands back a document whose
//     root element is <parsererror>, which is why the add-in has to check the
//     root element rather than assume a successful parse.
//   - textContent concatenates CDATA sections with ordinary text, so
//     repositories written by the original code still read correctly.
//
// It is deliberately strict about well-formedness: a test that feeds in
// damaged XML has to get the same answer a browser would.

const ELEMENT_NODE = 1;
const TEXT_NODE = 3;
const DOCUMENT_NODE = 9;

const PARSERERROR_NS = "http://www.w3.org/1999/xhtml";

const NAMED_ENTITIES = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" };

class XmlError extends Error {}

function decodeEntities(s) {
  return s.replace(/&([^;<&\s]*);?/g, (whole, body) => {
    if (!whole.endsWith(";")) throw new XmlError(`unterminated entity reference near "${whole}"`);
    if (body[0] === "#") {
      const code =
        body[1] === "x" || body[1] === "X"
          ? parseInt(body.slice(2), 16)
          : parseInt(body.slice(1), 10);
      if (!Number.isFinite(code)) throw new XmlError(`bad character reference "${whole}"`);
      return String.fromCodePoint(code);
    }
    if (!Object.prototype.hasOwnProperty.call(NAMED_ENTITIES, body)) {
      throw new XmlError(`undefined entity "${whole}"`);
    }
    return NAMED_ENTITIES[body];
  });
}

class XmlNode {
  constructor(nodeType) {
    this.nodeType = nodeType;
    this.childNodes = [];
    this.parentNode = null;
  }
}

class XmlText extends XmlNode {
  constructor(data) {
    super(TEXT_NODE);
    this.data = data;
  }
  get textContent() {
    return this.data;
  }
}

class XmlElement extends XmlNode {
  constructor(name, attrs) {
    super(ELEMENT_NODE);
    this.nodeName = name;
    this.tagName = name;
    const colon = name.indexOf(":");
    this.prefix = colon >= 0 ? name.slice(0, colon) : null;
    this.localName = colon >= 0 ? name.slice(colon + 1) : name;
    this._attrs = attrs;
  }
  getAttribute(name) {
    return Object.prototype.hasOwnProperty.call(this._attrs, name) ? this._attrs[name] : null;
  }
  hasAttribute(name) {
    return Object.prototype.hasOwnProperty.call(this._attrs, name);
  }
  get children() {
    return this.childNodes.filter((n) => n.nodeType === ELEMENT_NODE);
  }
  get textContent() {
    let out = "";
    for (const child of this.childNodes) out += child.textContent;
    return out;
  }
  querySelector(selector) {
    return querySelectorAll(this, selector)[0] ?? null;
  }
  querySelectorAll(selector) {
    return querySelectorAll(this, selector);
  }
}

class XmlDocument extends XmlNode {
  constructor() {
    super(DOCUMENT_NODE);
  }
  get documentElement() {
    return this.childNodes.find((n) => n.nodeType === ELEMENT_NODE) || null;
  }
  get textContent() {
    const root = this.documentElement;
    return root ? root.textContent : "";
  }
  querySelector(selector) {
    return querySelectorAll(this, selector)[0] ?? null;
  }
  querySelectorAll(selector) {
    return querySelectorAll(this, selector);
  }
}

// A CSS subset: type selectors joined by descendant (space) or child (>)
// combinators, which is all the add-in uses. Matching is by local name,
// case-sensitively, which is what a browser does for an XML document.
function querySelectorAll(root, selector) {
  const parts = selector
    .split(">")
    .map((s) => s.trim())
    .filter(Boolean);
  const steps = [];
  parts.forEach((part, i) => {
    const names = part.split(/\s+/).filter(Boolean);
    names.forEach((name, j) => {
      steps.push({ name, child: i > 0 && j === 0 });
    });
  });
  if (steps.length === 0) return [];

  let current = [root];
  for (const step of steps) {
    const next = [];
    for (const node of current) {
      const candidates = step.child ? childElements(node) : descendantElements(node);
      for (const el of candidates) {
        if (el.localName === step.name && !next.includes(el)) next.push(el);
      }
    }
    current = next;
  }
  return current;
}

function childElements(node) {
  return node.childNodes.filter((n) => n.nodeType === ELEMENT_NODE);
}

function descendantElements(node) {
  const out = [];
  const walk = (n) => {
    for (const child of n.childNodes) {
      if (child.nodeType !== ELEMENT_NODE) continue;
      out.push(child);
      walk(child);
    }
  };
  walk(node);
  return out;
}

const ATTR_RE = /([^\s=/>]+)\s*=\s*("([^"]*)"|'([^']*)')/g;

function parseAttributes(src) {
  const attrs = Object.create(null);
  let m;
  ATTR_RE.lastIndex = 0;
  while ((m = ATTR_RE.exec(src)) !== null) {
    const raw = m[3] !== undefined ? m[3] : m[4];
    // Attribute-value normalisation: tab, newline and carriage return in the
    // source all become a space before the value is reported.
    attrs[m[1]] = decodeEntities(raw.replace(/[\t\n\r]/g, " "));
  }
  return attrs;
}

/** XML 1.0 section 2.11: CRLF and lone CR both become LF, before anything else. */
function normaliseLineEndings(src) {
  return src.replace(/\r\n?/g, "\n");
}

function parseXml(raw) {
  const src = normaliseLineEndings(raw);
  const doc = new XmlDocument();
  const stack = [doc];
  let i = 0;
  let rootSeen = false;

  const push = (node) => {
    const parent = stack[stack.length - 1];
    if (parent === doc && node.nodeType === TEXT_NODE && node.data.trim() !== "") {
      throw new XmlError("text at the top level of the document");
    }
    node.parentNode = parent;
    parent.childNodes.push(node);
  };

  while (i < src.length) {
    const lt = src.indexOf("<", i);
    if (lt < 0) {
      const tail = src.slice(i);
      if (tail) push(new XmlText(decodeEntities(tail)));
      break;
    }
    if (lt > i) push(new XmlText(decodeEntities(src.slice(i, lt))));

    if (src.startsWith("<!--", lt)) {
      const end = src.indexOf("-->", lt);
      if (end < 0) throw new XmlError("unterminated comment");
      i = end + 3;
      continue;
    }
    if (src.startsWith("<![CDATA[", lt)) {
      const end = src.indexOf("]]>", lt);
      if (end < 0) throw new XmlError("unterminated CDATA section");
      push(new XmlText(src.slice(lt + 9, end))); // CDATA is not entity-decoded
      i = end + 3;
      continue;
    }
    if (src.startsWith("<?", lt)) {
      const end = src.indexOf("?>", lt);
      if (end < 0) throw new XmlError("unterminated processing instruction");
      i = end + 2;
      continue;
    }
    if (src.startsWith("<!", lt)) {
      const end = src.indexOf(">", lt);
      if (end < 0) throw new XmlError("unterminated declaration");
      i = end + 1;
      continue;
    }
    if (src.startsWith("</", lt)) {
      const end = src.indexOf(">", lt);
      if (end < 0) throw new XmlError("unterminated end tag");
      const name = src.slice(lt + 2, end).trim();
      const open = stack[stack.length - 1];
      if (open === doc) throw new XmlError(`end tag </${name}> with nothing open`);
      if (open.nodeName !== name) {
        throw new XmlError(`end tag </${name}> does not match <${open.nodeName}>`);
      }
      stack.pop();
      i = end + 1;
      continue;
    }

    // Start tag. Find the matching ">", skipping any inside quoted attributes.
    let j = lt + 1;
    let quote = null;
    for (; j < src.length; j++) {
      const c = src[j];
      if (quote) {
        if (c === quote) quote = null;
      } else if (c === '"' || c === "'") {
        quote = c;
      } else if (c === ">") {
        break;
      }
    }
    if (j >= src.length) throw new XmlError("unterminated start tag");

    let inner = src.slice(lt + 1, j);
    const selfClosing = inner.endsWith("/");
    if (selfClosing) inner = inner.slice(0, -1);

    const nameMatch = /^([^\s/>]+)/.exec(inner);
    if (!nameMatch) throw new XmlError("malformed start tag");
    if (stack.length === 1) {
      if (rootSeen) throw new XmlError("more than one root element");
      rootSeen = true;
    }
    const el = new XmlElement(nameMatch[1], parseAttributes(inner.slice(nameMatch[1].length)));
    push(el);
    if (!selfClosing) stack.push(el);
    i = j + 1;
  }

  if (stack.length > 1) throw new XmlError(`unclosed element <${stack[stack.length - 1].nodeName}>`);
  if (!rootSeen) throw new XmlError("no root element");
  return doc;
}

/** The document a browser returns instead of throwing. */
function parsererrorDocument(message) {
  const doc = new XmlDocument();
  const el = new XmlElement("parsererror", { xmlns: PARSERERROR_NS });
  el.parentNode = doc;
  doc.childNodes.push(el);
  const text = new XmlText(message);
  text.parentNode = el;
  el.childNodes.push(text);
  return doc;
}

class DOMParser {
  parseFromString(src, mime) {
    if (mime !== "application/xml" && mime !== "text/xml") {
      throw new Error(`xml shim: unsupported mime type ${mime}`);
    }
    try {
      return parseXml(src);
    } catch (e) {
      if (e instanceof XmlError) return parsererrorDocument(e.message);
      throw e;
    }
  }
}

/** The parse function `parseRepoXml` takes, wired to this shim. */
function parse(xml) {
  return new DOMParser().parseFromString(xml, "application/xml");
}

module.exports = { DOMParser, parse, parseXml, XmlElement, XmlText, XmlDocument, XmlError };

/*
 * Macrotec Font: HTML transform.
 *
 * Applies the Macrotec house email format (Century Gothic 10pt) to the new
 * part of a draft and leaves any quoted thread below it untouched.
 *
 * Pure function: takes an HTML string and a DOMParser-capable document
 * factory, returns the new HTML string. No network access, no Office.js,
 * so it can be unit tested in Node with jsdom.
 */
(function (root) {
  "use strict";

  var FONT_FAMILY = "'Century Gothic',Arial,Helvetica,sans-serif";
  var FONT_SIZE = "10pt";

  // Elements that mark the start of a quoted reply or forward. Everything
  // from the first match onwards is left exactly as it is.
  var QUOTE_BOUNDARY_SELECTORS = [
    "#mail-editor-reference-message-container", // new Outlook / OWA reply
    "#divRplyFwdMsg", // classic Outlook / OWA reply header
    "#appendonsend", // OWA marker above quoted content
    ".gmail_quote", // quoted Gmail content
    "blockquote",
  ];

  // Elements that get the house font written into them.
  var TEXT_ELEMENTS = [
    "div", "p", "span", "font", "li", "ol", "ul", "td", "th", "a",
    "b", "strong", "i", "em", "u", "strike", "h1", "h2", "h3", "h4", "h5", "h6",
  ];

  var LIST_TAGS = { OL: true, UL: true };

  function setStyles(el, styles) {
    Object.keys(styles).forEach(function (prop) {
      el.style.setProperty(prop, styles[prop]);
    });
  }

  function houseFont(el) {
    setStyles(el, { "font-family": FONT_FAMILY, "font-size": FONT_SIZE });
    if (el.tagName === "FONT") {
      el.removeAttribute("face");
      el.removeAttribute("size");
    }
  }

  function isBlankLine(el) {
    if (!el || el.tagName !== "DIV") return false;
    var text = (el.textContent || "").replace(/ /g, "").trim();
    if (text !== "") return false;
    // A blank line is a div holding nothing but an optional <br>.
    var children = el.children;
    return children.length === 0 || (children.length === 1 && children[0].tagName === "BR");
  }

  function makeBlankLine(doc) {
    var div = doc.createElement("div");
    div.appendChild(doc.createElement("br"));
    houseFont(div);
    return div;
  }

  // Header block Outlook for Mac and others insert above a quoted reply:
  // a div with a top border whose text starts "From:".
  function isReplyHeader(el) {
    if (!el || el.tagName !== "DIV") return false;
    var style = (el.getAttribute("style") || "").toLowerCase();
    var text = (el.textContent || "").trim();
    return style.indexOf("border-top") !== -1 && /^(From|Van|De):/.test(text);
  }

  function findBoundary(body) {
    var candidates = [];
    QUOTE_BOUNDARY_SELECTORS.forEach(function (sel) {
      var found = body.querySelector(sel);
      if (found) candidates.push(found);
    });
    Array.prototype.forEach.call(body.querySelectorAll("div"), function (div) {
      if (isReplyHeader(div)) candidates.push(div);
    });
    if (candidates.length === 0) return null;
    // Earliest in document order wins.
    candidates.sort(function (a, b) {
      return a.compareDocumentPosition(b) & 4 /* FOLLOWING */ ? -1 : 1;
    });
    return candidates[0];
  }

  // True when el sits before the boundary and is not one of its ancestors.
  function isNewContent(el, boundary) {
    if (!boundary) return true;
    if (el === boundary) return false;
    var pos = boundary.compareDocumentPosition(el);
    var precedes = (pos & 2) /* PRECEDING */ !== 0;
    var contains = (pos & 8) /* CONTAINS: el is an ancestor */ !== 0;
    return precedes && !contains;
  }

  function wrapLooseText(doc, body, boundary) {
    // Text sitting directly in body (or in an ancestor of the boundary)
    // has no element to carry the font, so wrap it in a span.
    var walker = doc.createTreeWalker(body, 4 /* SHOW_TEXT */);
    var loose = [];
    var node;
    while ((node = walker.nextNode())) {
      if (!node.nodeValue.trim()) continue;
      var parent = node.parentNode;
      var parentIsContainer =
        parent === body || (boundary && parent.contains(boundary));
      if (!parentIsContainer) continue;
      if (boundary && !(boundary.compareDocumentPosition(node) & 2)) continue;
      loose.push(node);
    }
    loose.forEach(function (textNode) {
      var span = doc.createElement("span");
      textNode.parentNode.insertBefore(span, textNode);
      span.appendChild(textNode);
    });
  }

  function normaliseLists(doc, lists) {
    lists.forEach(function (list) {
      // Nested lists keep Outlook's own spacing.
      if (list.parentElement && list.parentElement.tagName === "LI") return;
      setStyles(list, { "margin-top": "0", "margin-bottom": "0" });

      // Exactly one blank line before and after the list, like a hand-typed email.
      var prev = list.previousElementSibling;
      if (prev && !isBlankLine(prev) && !LIST_TAGS[prev.tagName]) {
        list.parentNode.insertBefore(makeBlankLine(doc), list);
      }
      var next = list.nextElementSibling;
      if (next && !isBlankLine(next) && !LIST_TAGS[next.tagName]) {
        list.parentNode.insertBefore(makeBlankLine(doc), next);
      }

      Array.prototype.forEach.call(list.children, function (li) {
        if (li.tagName !== "LI") return;
        setStyles(li, { "margin-top": "0", "margin-bottom": "0" });
        // Paragraphs inside list items cause the wide gaps seen in
        // add-in drafts; strip their margins.
        Array.prototype.forEach.call(li.querySelectorAll("p"), function (p) {
          setStyles(p, { "margin-top": "0", "margin-bottom": "0" });
        });
      });
    });
  }

  /**
   * Apply the house font to the new part of an email body.
   * @param {string} html  Body HTML from Office.js body.getAsync.
   * @param {function(string): Document} parse  Returns a Document for html.
   * @returns {{html: string, changed: number}}
   */
  function applyHouseFont(html, parse) {
    if (typeof html !== "string") throw new TypeError("html must be a string");
    var doc = parse(html);
    var body = doc.body;
    if (!body) return { html: html, changed: 0 };

    var boundary = findBoundary(body);
    wrapLooseText(doc, body, boundary);

    var selector = TEXT_ELEMENTS.join(",");
    var targets = Array.prototype.filter.call(
      body.querySelectorAll(selector),
      function (el) { return isNewContent(el, boundary); }
    );
    targets.forEach(houseFont);

    var lists = targets.filter(function (el) { return LIST_TAGS[el.tagName]; });
    normaliseLists(doc, lists);

    // Keep a full document (with any <head> styles the quoted thread relies
    // on) when Outlook gave us one; otherwise return just the body content.
    var out = /<html[\s>]/i.test(html) ? doc.documentElement.outerHTML : body.innerHTML;
    return { html: out, changed: targets.length };
  }

  var api = { applyHouseFont: applyHouseFont, FONT_FAMILY: FONT_FAMILY, FONT_SIZE: FONT_SIZE };
  if (typeof module !== "undefined" && module.exports) {
    module.exports = api;
  } else {
    root.MacrotecFont = api;
  }
})(typeof self !== "undefined" ? self : this);

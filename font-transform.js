/*
 * Email Composer: HTML transform.
 *
 * Applies the Macrotec house email format (Century Gothic 10pt, black text,
 * one blank line around lists and after the sign-off) to the new part of a
 * draft and leaves any quoted thread below it untouched.
 *
 * Pure function: takes an HTML string and a DOMParser-capable document
 * factory, returns the new HTML string. No network access, no Office.js,
 * so it can be unit tested in Node with jsdom.
 */
(function (root) {
  "use strict";

  var FONT_FAMILY = "'Century Gothic',Arial,Helvetica,sans-serif";
  var FONT_SIZE = "10pt";
  // Outlook writes this colour into text typed in the compose window. Text
  // without a colour shows in the reader's default instead, which in dark
  // mode is a dimmer grey than typed text, so set it explicitly.
  var TEXT_COLOUR = "rgb(0,0,0)";

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

  // Block elements that get the default text colour. Inline elements and
  // links inherit it, so coloured words and blue links keep their colour.
  var COLOUR_TAGS = {
    DIV: true, P: true, LI: true, OL: true, UL: true, TD: true, TH: true,
    H1: true, H2: true, H3: true, H4: true, H5: true, H6: true,
  };

  function setStyles(el, styles) {
    Object.keys(styles).forEach(function (prop) {
      el.style.setProperty(prop, styles[prop]);
    });
  }

  // True when el or an ancestor below body already sets a text colour.
  function hasColour(el, body) {
    for (var node = el; node && node !== body; node = node.parentElement) {
      if (node.style.color || node.getAttribute("color")) return true;
    }
    return false;
  }

  function houseFont(el, body) {
    setStyles(el, { "font-family": FONT_FAMILY, "font-size": FONT_SIZE });
    if (COLOUR_TAGS[el.tagName] && !hasColour(el, body)) {
      setStyles(el, { color: TEXT_COLOUR });
    }
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

  // Looser test used for the sign-off gap: an empty div or paragraph in any
  // form Outlook writes it (<br>, &nbsp;, an empty span), with no images,
  // tables, rules or lists inside.
  function isEmptyLine(el) {
    if (!el || (el.tagName !== "DIV" && el.tagName !== "P")) return false;
    var text = (el.textContent || "").replace(/[\s ​]/g, "");
    if (text !== "") return false;
    return !el.querySelector("img,table,hr,ol,ul,iframe,object,svg,video");
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

  // The child of parent that is target or contains it.
  function childContaining(parent, target) {
    var node = target;
    while (node && node.parentNode !== parent) node = node.parentNode;
    return node;
  }

  function hasContent(node) {
    if (node.nodeType === 3) return node.nodeValue.replace(/[\s ​]/g, "") !== "";
    return node.nodeType === 1 && !isEmptyLine(node) && node.tagName !== "BR";
  }

  /*
   * Collapse the blank lines at the end of the new text (after the sign-off,
   * above the signature Exclaimer adds on send or the quoted reply) to one.
   * It only removes lines, never adds one. Returns the number removed.
   */
  function trimSignOffGap(body, boundary) {
    var container = body;
    var end = null;
    if (boundary) {
      // Find the level where the new text and the boundary are siblings:
      // step into wrappers that hold both.
      end = childContaining(body, boundary);
      while (end && end !== boundary) {
        var inner = childContaining(end, boundary);
        var before = inner.previousSibling;
        while (before && !hasContent(before)) before = before.previousSibling;
        if (!before) break;
        container = end;
        end = inner;
      }
    }

    var blanks = [];
    var node = end ? end.previousSibling : container.lastChild;
    while (node) {
      var isBlank =
        (node.nodeType === 3 && !hasContent(node)) ||
        node.nodeType === 8 ||
        (node.nodeType === 1 && (isEmptyLine(node) || node.tagName === "BR"));
      if (!isBlank) break;
      if (node.nodeType === 1) blanks.push(node);
      node = node.previousSibling;
    }
    // Nothing but blank lines: an empty draft, leave it alone.
    if (!node || blanks.length < 2) return 0;

    // blanks runs bottom-up; keep the one directly under the sign-off.
    var extra = blanks.slice(0, blanks.length - 1);
    extra.forEach(function (el) { el.parentNode.removeChild(el); });
    return extra.length;
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
    targets.forEach(function (el) { houseFont(el, body); });

    var lists = targets.filter(function (el) { return LIST_TAGS[el.tagName]; });
    normaliseLists(doc, lists);
    trimSignOffGap(body, boundary);

    // Keep a full document (with any <head> styles the quoted thread relies
    // on) when Outlook gave us one; otherwise return just the body content.
    var out = /<html[\s>]/i.test(html) ? doc.documentElement.outerHTML : body.innerHTML;
    return { html: out, changed: targets.length };
  }

  var api = {
    applyHouseFont: applyHouseFont,
    FONT_FAMILY: FONT_FAMILY,
    FONT_SIZE: FONT_SIZE,
    TEXT_COLOUR: TEXT_COLOUR,
  };
  if (typeof module !== "undefined" && module.exports) {
    module.exports = api;
  } else {
    root.MacrotecFont = api;
  }
})(typeof self !== "undefined" ? self : this);

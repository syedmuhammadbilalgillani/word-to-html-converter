import React, { useState } from "react";

// --- Custom Sanitizer ---
const DEFAULT_ALLOWED_TAGS = [
  "html", "head", "body", "title", "meta", "style", "link", "script", "noscript",
  "div", "span", "p", "b", "i", "u", "strong", "em", "br", "hr",
  "h1", "h2", "h3", "h4", "h5", "h6",
  "ul", "ol", "li", "dl", "dt", "dd",
  "table", "thead", "tbody", "tfoot", "tr", "td", "th",
  "a", "img", "blockquote", "pre", "code", "nav", "section", "article", "aside", "header", "footer"
];

const DEFAULT_ALLOWED_ATTRS = ["href", "src", "alt", "title", "style", "class", "id", "name"];

function sanitizeHtml(html: string, allowedTags = DEFAULT_ALLOWED_TAGS, allowedAttrs = DEFAULT_ALLOWED_ATTRS): string {
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, "text/html");

  function clean(node: HTMLElement | ChildNode) {
    if (node.nodeType === 1) {
      const el = node as HTMLElement;
      if (!allowedTags.includes(el.tagName.toLowerCase())) {
        el.remove();
        return;
      }
      [...el.attributes].forEach(attr => {
        if (!allowedAttrs.includes(attr.name.toLowerCase())) {
          el.removeAttribute(attr.name);
        }
      });
    }
    [...node.childNodes].forEach(clean);
  }

  doc.body.childNodes.forEach(clean);
  return doc.body.innerHTML;
}

// --- Custom Beautifier ---
function beautifyHtml(html: string, indentSize = 2): string {
  let formatted = '';
  const pad = ' '.repeat(indentSize);
  let depth = 0;

  html.replace(/>\s*</g, '><')
    .split(/(?=<)/g)
    .forEach(line => {
      if (line.match(/^<\//)) depth--;
      formatted += pad.repeat(depth) + line + '\n';
      if (line.match(/^<\w[^>]*[^/]>/)) depth++;
    });

  return formatted.trim();
}

// --- Inline Style Extractor to CSS ---
function extractStyles(html: string): { css: string; html: string } {
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, "text/html");
  const styles: string[] = [];
  let styleIndex = 0;

  const walk = (node: HTMLElement | ChildNode) => {
    if (node.nodeType === 1) {
      const el = node as HTMLElement;
      const style = el.getAttribute("style");
      if (style) {
        const className = `custom-style-${styleIndex++}`;
        styles.push(`.${className} { ${style} }`);
        el.removeAttribute("style");
        el.classList.add(className);
      }
    }
    [...node.childNodes].forEach(walk);
  };

  doc.body.childNodes.forEach(walk);

  return {
    css: styles.join("\n"),
    html: doc.body.innerHTML
  };
}

export function HTMLInlineCSS() {
  const [htmlOutput, setHtmlOutput] = useState<string>("");
  const [cssOutput, setCssOutput] = useState<string>("");

  const handlePaste = (e: React.ClipboardEvent<HTMLDivElement>) => {
    e.preventDefault();
    const pastedHTML =
      e.clipboardData.getData("text/html") ||
      e.clipboardData.getData("text/plain");

    const cleanHTML = sanitizeHtml(pastedHTML);
    const { css, html } = extractStyles(cleanHTML);

    setCssOutput(css);
    setHtmlOutput(beautifyHtml(html));
  };

  const handleCopy = () => {
    navigator.clipboard.writeText(htmlOutput);
    alert("HTML copied to clipboard!");
  };

  return (
    <div className="p-6 max-w-4xl mx-auto">
      <h1 className="text-2xl font-bold mb-4">Word to HTML Converter</h1>

      <div
        contentEditable
        onPaste={handlePaste}
        className="border border-gray-400 p-4 mb-4 min-h-[150px] rounded-lg bg-white shadow-sm"
      >
        <p className="text-gray-500">Paste your content here...</p>
      </div>

      <div className="mb-2 flex items-center justify-between">
        <h2 className="text-xl font-semibold">Generated HTML</h2>
        <button
          onClick={handleCopy}
          className="bg-blue-500 text-white px-4 py-1 rounded hover:bg-blue-600"
        >
          Copy HTML
        </button>
      </div>

      <textarea
        className="w-full h-64 p-4 border rounded font-mono text-sm mb-4"
        value={htmlOutput}
        readOnly
      />

      <h2 className="text-xl font-semibold mb-2">Generated CSS</h2>
      <textarea
        className="w-full h-40 p-4 border rounded font-mono text-sm"
        value={cssOutput}
        readOnly
      />
    </div>
  );
}

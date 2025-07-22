import { Code, Copy, Download, Eye, RefreshCw, Sun, Moon } from "lucide-react";
import { useState } from "react";

// Enhanced sanitizer for Word documents
const ALLOWED_TAGS = [
  "div",
  "span",
  "p",
  "b",
  "i",
  "u",
  "strong",
  "em",
  "br",
  "hr",
  "h1",
  "h2",
  "h3",
  "h4",
  "h5",
  "h6",
  "ul",
  "ol",
  "li",
  "dl",
  "dt",
  "dd",
  "table",
  "thead",
  "tbody",
  "tfoot",
  "tr",
  "td",
  "th",
  "a",
  "img",
  "blockquote",
  "pre",
  "code",
  "sup",
  "sub",
];

const ALLOWED_ATTRS = ["href", "src", "alt", "title", "class", "id"];

// Word-specific cleanup patterns
const WORD_CLEANUP_PATTERNS = [
  /<o:p\s*\/?>|<\/o:p>/gi,
  /<w:[^>]*>|<\/w:[^>]*>/gi,
  /<m:[^>]*>|<\/m:[^>]*>/gi,
  /mso-[^=]*="[^"]*"/gi,
  /lang="[^"]*"/gi,
  /style="[^"]*"/gi,
  /class="Mso[^"]*"/gi,
  /<\?xml[^>]*>/gi,
  /<meta[^>]*>/gi,
  /<p[^>]*>\s*<\/p>/gi,
  /<span[^>]*>\s*<\/span>/gi,
];

function cleanWordHtml(html: string) {
  let cleaned = html;
  WORD_CLEANUP_PATTERNS.forEach((pattern) => {
    cleaned = cleaned.replace(pattern, "");
  });
  cleaned = cleaned.replace(/\s+/g, " ");
  cleaned = cleaned.replace(/>\s+</g, "><");
  return cleaned;
}

function sanitizeHtml(
  html: string,
  allowedTags = ALLOWED_TAGS,
  allowedAttrs = ALLOWED_ATTRS,
  removeSpanBr = false
) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, "text/html");

  function clean(node) {
    if (node.nodeType === 1) {
      const el = node;
      const tagName = el.tagName.toLowerCase();
      if (removeSpanBr && (tagName === "span" || tagName === "br")) {
        const parent = el.parentNode;
        while (el.firstChild) {
          parent.insertBefore(el.firstChild, el);
        }
        el.remove();
        return;
      }
      if (!allowedTags.includes(tagName)) {
        const parent = el.parentNode;
        while (el.firstChild) {
          parent.insertBefore(el.firstChild, el);
        }
        el.remove();
        return;
      }
      [...el.attributes].forEach((attr) => {
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

function htmlToJsx(html: string) {
  let jsx = html;
  const attributeMap = {
    class: "className",
    for: "htmlFor",
    tabindex: "tabIndex",
    readonly: "readOnly",
    maxlength: "maxLength",
    cellpadding: "cellPadding",
    cellspacing: "cellSpacing",
    rowspan: "rowSpan",
    colspan: "colSpan",
    usemap: "useMap",
    frameborder: "frameBorder",
    contenteditable: "contentEditable",
  };

  Object.entries(attributeMap).forEach(([htmlAttr, jsxAttr]) => {
    const regex = new RegExp(`\\b${htmlAttr}=`, "gi");
    jsx = jsx.replace(regex, `${jsxAttr}=`);
  });

  jsx = jsx.replace(/style="([^"]+)"/g, (match, styleContent) => {
    const styleObj = styleContent
      .split(";")
      .filter((prop) => prop.trim())
      .map((prop) => {
        const [key, value] = prop.split(":").map((s) => s.trim());
        if (!key || !value) return "";
        const camelKey = key.replace(/-([a-z])/g, (match, letter) =>
          letter.toUpperCase()
        );
        const processedValue = /^\d+$/.test(value) ? value : `'${value}'`;
        return `${camelKey}: ${processedValue}`;
      })
      .filter((prop) => prop)
      .join(", ");
    return styleObj ? `style={{${styleObj}}}` : "";
  });

  const selfClosingTags = ["br", "hr", "img", "input", "meta", "link"];
  selfClosingTags.forEach((tag) => {
    const regex = new RegExp(`<${tag}([^>]*?)>`, "gi");
    jsx = jsx.replace(regex, (match, attributes) => `<${tag}${attributes} />`);
  });

  jsx = jsx.replace(/<!--([\s\S]*?)-->/g, "{/* $1 */}");
  return jsx;
}
function generateCompleteJsx(
  bodyContent: string,
  componentName = "ConvertedComponent"
) {
  const jsxContent = htmlToJsx(bodyContent);
  return `import React from 'react';

const ${componentName} = () => {
  return (
    <>
${jsxContent
  .split("\n")
  .map((line) => "      " + line)
  .join("\n")}
    </>
  );
};

export default ${componentName};`;
}

function generateCompleteHtml(bodyContent, title = "Converted Document") {
  return `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${title}</title>
    <style>
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
            line-height: 1.6;
            max-width: 800px;
            margin: 0 auto;
            padding: 20px;
            color: #333;
        }
        .dark-mode {
            background-color: #1a1a1a;
            color: #e0e0e0;
        }
        h1, h2, h3, h4, h5, h6 {
            margin-top: 1.5em;
            margin-bottom: 0.5em;
            font-weight: 600;
        }
        p {
            margin-bottom: 1em;
        }
        ul, ol {
            margin-bottom: 1em;
            padding-left: 2em;
        }
        table {
            border-collapse: collapse;
            width: 100%;
            margin-bottom: 1em;
        }
        th, td {
            border: 1px solid #ddd;
            padding: 8px;
            text-align: left;
        }
        th {
            background-color: #f2f2f2;
            font-weight: 600;
        }
        .dark-mode th {
            background-color: #333;
            border-color: #555;
        }
        .dark-mode td {
            border-color: #555;
        }
        blockquote {
            border-left: 4px solid #ddd;
            margin-left: 0;
            padding-left: 1em;
            color: #666;
        }
        .dark-mode blockquote {
            border-color: #555;
            color: #aaa;
        }
        code {
            background-color: #f4f4f4;
            padding: 2px 4px;
            border-radius: 3px;
            font-family: monospace;
        }
        .dark-mode code {
            background-color: #333;
        }
        pre {
            background-color: #f4f4f4;
            padding: 1em;
            border-radius: 5px;
            overflow-x: auto;
        }
        .dark-mode pre {
            background-color: #333;
        }
    </style>
</head>
<body>
${bodyContent}
</body>
</html>`;
}

function beautifyHtml(html, indentSize = 2) {
  let formatted = "";
  const pad = " ".repeat(indentSize);
  let depth = 0;
  const tokens = html.split(/(<[^>]*>)/g).filter((token) => token.length > 0);

  tokens.forEach((token) => {
    if (token.match(/<\/\w+>/)) {
      depth--;
      if (formatted.endsWith("\n")) {
        formatted += pad.repeat(Math.max(0, depth)) + token + "\n";
      } else {
        formatted += token + "\n";
      }
    } else if (token.match(/<\w+[^>]*>/)) {
      if (formatted.endsWith("\n")) {
        formatted += pad.repeat(depth) + token;
      } else {
        formatted += "\n" + pad.repeat(depth) + token;
      }
      if (!token.match(/<\w+[^>]*\/>/)) {
        depth++;
      }
      if (token.match(/<(br|hr|img|input)[^>]*>/i)) {
        formatted += "\n";
      }
    } else if (token.trim()) {
      if (formatted.endsWith(">")) {
        formatted += token;
      } else {
        formatted += token;
      }
    }
  });
  return formatted.trim();
}

export default function WordToHtmlConverter() {
  const [htmlOutput, setHtmlOutput] = useState("");
  const [bodyContent, setBodyContent] = useState("");
  const [activeTab, setActiveTab] = useState("preview");
  const [isProcessing, setIsProcessing] = useState(false);
  const [outputFormat, setOutputFormat] = useState("complete");
  const [outputType, setOutputType] = useState("html");
  const [removeSpanBr, setRemoveSpanBr] = useState(false);
  const [customTags, setCustomTags] = useState("");
  const [customAttrs, setCustomAttrs] = useState("");
  const [fontSize, setFontSize] = useState(16);
  const [isDarkMode, setIsDarkMode] = useState(false);
  const [errorMessage, setErrorMessage] = useState("");
  const [wordCleanedHtml, setWordCleanedHtml] = useState("");

  const handlePaste = async (e) => {
    e.preventDefault();
    setIsProcessing(true);
    setErrorMessage("");

    try {
      const clipboardData = e.clipboardData;
      let pastedHTML = clipboardData.getData("text/html");

      if (!pastedHTML) {
        const pastedText = clipboardData.getData("text/plain");
        pastedHTML = pastedText
          .split(" Veronika")
          .map((line) => (line.trim() ? `<p>${line}</p>` : ""))
          .join("");
      }

      const wordCleaned = cleanWordHtml(pastedHTML);
      setWordCleanedHtml(wordCleaned);
      const tagsToUse = customTags
        ? customTags
            .split(",")
            .map((tag) => tag.trim().toLowerCase())
            .filter((tag) => tag)
        : ALLOWED_TAGS;
      const attrsToUse = customAttrs
        ? customAttrs
            .split(",")
            .map((attr) => attr.trim().toLowerCase())
            .filter((attr) => attr)
        : ALLOWED_ATTRS;
      const sanitized = sanitizeHtml(
        wordCleaned,
        tagsToUse,
        attrsToUse,
        removeSpanBr
      );
      setBodyContent(sanitized);

      let finalOutput;
      if (outputType === "jsx") {
        finalOutput =
          outputFormat === "complete"
            ? generateCompleteJsx(sanitized, "ConvertedComponent")
            : htmlToJsx(sanitized);
      } else {
        finalOutput =
          outputFormat === "complete"
            ? generateCompleteHtml(sanitized, "Converted Document")
            : sanitized;
      }

      const beautified = beautifyHtml(finalOutput);
      setHtmlOutput(beautified);
      setActiveTab("preview");
    } catch (error) {
      console.error("Error processing pasted content:", error);
      setErrorMessage("Error processing pasted content: " + error.message);
    } finally {
      setIsProcessing(false);
    }
  };

  const handleCopy = async () => {
    try {
      await navigator.clipboard.writeText(htmlOutput);
      alert("Content copied to clipboard!");
    } catch (error) {
      console.error("Failed to copy:", error);
      setErrorMessage("Failed to copy to clipboard: " + error.message);
    }
  };

  const handleDownload = () => {
    const fileExtension = outputType === "jsx" ? "jsx" : "html";
    const mimeType = outputType === "jsx" ? "text/javascript" : "text/html";
    const fileName =
      outputType === "jsx" ? "ConvertedComponent.jsx" : "converted.html";
    const blob = new Blob([htmlOutput], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const handleClear = () => {
    setHtmlOutput("");
    setBodyContent("");
    setErrorMessage("");
    setWordCleanedHtml("");
  };

  return (
    <div
      className={`min-h-screen ${
        isDarkMode ? "bg-gray-800" : "bg-gray-50"
      } p-6`}
    >
      <div className="max-w-6xl mx-auto">
        <div
          className={`bg-white ${
            isDarkMode ? "dark:bg-gray-900 dark:text-white" : ""
          } rounded-lg shadow-lg overflow-hidden`}
        >
          <div className="bg-gradient-to-r from-blue-600 to-purple-600 text-white p-6">
            <h1 className="text-3xl font-bold mb-2">
              Word to HTML/JSX Converter
            </h1>
            <p className="text-blue-100">
              Convert Word documents to clean HTML or React JSX components with
              live preview
            </p>
          </div>

          <div className="p-6 border-b border-gray-200">
            <h2 className="text-xl font-semibold mb-4 text-gray-800 dark:text-gray-200">
              Paste Your Content
            </h2>
            <div
              contentEditable
              onPaste={handlePaste}
              className={`border-2 border-dashed ${
                isDarkMode ? "border-gray-600" : "border-gray-300"
              } rounded-lg p-6 min-h-[200px] focus:border-blue-500 focus:outline-none transition-colors`}
              style={{ minHeight: "200px" }}
            >
              {isProcessing ? (
                <div className="flex items-center justify-center text-gray-500 dark:text-gray-400">
                  <RefreshCw className="animate-spin mr-2" size={20} />
                  Processing...
                </div>
              ) : (
                <div className="text-gray-500 dark:text-gray-400">
                  <p className="text-lg mb-2">
                    📋 Paste your Word content here
                  </p>
                  <p className="text-sm">
                    Copy content from Word and paste it here. The converter will
                    automatically clean and format the HTML.
                  </p>
                </div>
              )}
            </div>
          </div>

          <div className="p-6">
            <div className="flex flex-wrap items-center justify-between mb-4 gap-4">
              <div className="flex space-x-2">
                <button
                  onClick={() => setActiveTab("preview")}
                  className={`flex items-center px-4 py-2 rounded-lg transition-colors ${
                    activeTab === "preview"
                      ? "bg-blue-500 text-white"
                      : isDarkMode
                      ? "bg-gray-700 text-gray-200 hover:bg-gray-600"
                      : "bg-gray-200 text-gray-700 hover:bg-gray-300"
                  }`}
                >
                  <Eye size={16} className="mr-2" />
                  Preview
                </button>
                <button
                  onClick={() => setActiveTab("code")}
                  className={`flex items-center px-4 py-2 rounded-lg transition-colors ${
                    activeTab === "code"
                      ? "bg-blue-500 text-white"
                      : isDarkMode
                      ? "bg-gray-700 text-gray-200 hover:bg-gray-600"
                      : "bg-gray-200 text-gray-700 hover:bg-gray-300"
                  }`}
                >
                  <Code size={16} className="mr-2" />
                  {outputType === "jsx" ? "JSX Code" : "HTML Code"}
                </button>
              </div>

              <div className="flex flex-wrap items-center space-x-4 gap-4">
                <div className="flex items-center space-x-2">
                  <label className="text-sm font-medium text-gray-700 dark:text-gray-200">
                    Type:
                  </label>
                  <select
                    value={outputType}
                    onChange={(e) => {
                      setOutputType(e.target.value);
                      if (bodyContent) {
                        let finalOutput;
                        if (e.target.value === "jsx") {
                          finalOutput =
                            outputFormat === "complete"
                              ? generateCompleteJsx(
                                  bodyContent,
                                  "ConvertedComponent"
                                )
                              : htmlToJsx(bodyContent);
                        } else {
                          finalOutput =
                            outputFormat === "complete"
                              ? generateCompleteHtml(
                                  bodyContent,
                                  "Converted Document"
                                )
                              : bodyContent;
                        }
                        setHtmlOutput(beautifyHtml(finalOutput));
                      }
                    }}
                    className="px-3 py-1 border border-gray-300 rounded text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 dark:border-gray-600 dark:bg-gray-700 dark:text-white"
                  >
                    <option value="html">HTML</option>
                    <option value="jsx">JSX/React</option>
                  </select>
                </div>

                <div className="flex items-center space-x-2">
                  <label className="text-sm font-medium text-gray-700 dark:text-gray-200">
                    Format:
                  </label>
                  <select
                    value={outputFormat}
                    onChange={(e) => {
                      setOutputFormat(e.target.value);
                      if (bodyContent) {
                        let finalOutput;
                        if (outputType === "jsx") {
                          finalOutput =
                            e.target.value === "complete"
                              ? generateCompleteJsx(
                                  bodyContent,
                                  "ConvertedComponent"
                                )
                              : htmlToJsx(bodyContent);
                        } else {
                          finalOutput =
                            e.target.value === "complete"
                              ? generateCompleteHtml(
                                  bodyContent,
                                  "Converted Document"
                                )
                              : bodyContent;
                        }
                        setHtmlOutput(beautifyHtml(finalOutput));
                      }
                    }}
                    className="px-3 py-1 border border-gray-300 rounded text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 dark:border-gray-600 dark:bg-gray-700 dark:text-white"
                  >
                    <option value="complete">
                      {outputType === "jsx"
                        ? "Complete Component"
                        : "Complete HTML"}
                    </option>
                    <option value="body-only">
                      {outputType === "jsx"
                        ? "JSX Content Only"
                        : "Body Content Only"}
                    </option>
                  </select>
                </div>

                <div className="flex items-center space-x-2">
                  <input
                    type="checkbox"
                    id="removeSpanBr"
                    checked={removeSpanBr}
                    onChange={(e) => {
                      setRemoveSpanBr(e.target.checked);
                      if (bodyContent) {
                        const tagsToUse = customTags
                          ? customTags
                              .split(",")
                              .map((tag) => tag.trim().toLowerCase())
                              .filter((tag) => tag)
                          : ALLOWED_TAGS;
                        const attrsToUse = customAttrs
                          ? customAttrs
                              .split(",")
                              .map((attr) => attr.trim().toLowerCase())
                              .filter((attr) => attr)
                          : ALLOWED_ATTRS;
                        const sanitized = sanitizeHtml(
                          wordCleanedHtml,
                          tagsToUse,
                          attrsToUse,
                          e.target.checked
                        );
                        setBodyContent(sanitized);
                        let finalOutput;
                        if (outputType === "jsx") {
                          finalOutput =
                            outputFormat === "complete"
                              ? generateCompleteJsx(
                                  sanitized,
                                  "ConvertedComponent"
                                )
                              : htmlToJsx(sanitized);
                        } else {
                          finalOutput =
                            outputFormat === "complete"
                              ? generateCompleteHtml(
                                  sanitized,
                                  "Converted Document"
                                )
                              : sanitized;
                        }
                        setHtmlOutput(beautifyHtml(finalOutput));
                      }
                    }}
                    className="h-4 w-4 text-blue-500 focus:ring-blue-500 border-gray-300 rounded dark:border-gray-600"
                  />
                  <label
                    htmlFor="removeSpanBr"
                    className="text-sm font-medium text-gray-700 dark:text-gray-200"
                  >
                    Remove &lt;span&gt; and &lt;br&gt;
                  </label>
                </div>

                <div className="flex items-center space-x-2">
                  <label className="text-sm font-medium text-gray-700 dark:text-gray-200">
                    Font Size:
                  </label>
                  <input
                    type="range"
                    min="12"
                    max="24"
                    value={fontSize}
                    onChange={(e) => setFontSize(Number(e.target.value))}
                    className="w-24"
                  />
                  <span className="text-sm text-gray-600 dark:text-gray-400">
                    {fontSize}px
                  </span>
                </div>

                <button
                  onClick={() => setIsDarkMode(!isDarkMode)}
                  className={`flex items-center px-4 py-2 rounded-lg transition-colors ${
                    isDarkMode
                      ? "bg-gray-700 text-white hover:bg-gray-600"
                      : "bg-gray-200 text-gray-700 hover:bg-gray-300"
                  }`}
                >
                  {isDarkMode ? (
                    <Sun size={16} className="mr-2" />
                  ) : (
                    <Moon size={16} className="mr-2" />
                  )}
                  {isDarkMode ? "Light Mode" : "Dark Mode"}
                </button>
              </div>

              <div className="flex space-x-2">
                <button
                  onClick={handleCopy}
                  className="flex items-center px-4 py-2 bg-green-500 text-white rounded-lg hover:bg-green-600 transition-colors"
                >
                  <Copy size={16} className="mr-2" />
                  Copy {outputType === "jsx" ? "JSX" : "HTML"}
                </button>
                <button
                  onClick={handleDownload}
                  className="flex items-center px-4 py-2 bg-purple-500 text-white rounded-lg hover:bg-purple-600 transition-colors"
                >
                  <Download size={16} className="mr-2" />
                  Download
                </button>
                <button
                  onClick={handleClear}
                  className="flex items-center px-4 py-2 bg-red-500 text-white rounded-lg hover:bg-red-600 transition-colors"
                >
                  <RefreshCw size={16} className="mr-2" />
                  Clear
                </button>
              </div>
            </div>

            {errorMessage && (
              <div className="mt-4 p-4 bg-red-100 text-red-700 rounded-lg">
                <p>{errorMessage}</p>
              </div>
            )}

            <div className="mt-4 grid md:grid-cols-2 gap-4">
              <div>
                <label className="text-sm font-medium text-gray-700 dark:text-gray-200">
                  Custom Allowed Tags (comma-separated):
                </label>
                0s0n
                <input
                  type="text"
                  value={customTags}
                  onChange={(e) => {
                    setCustomTags(e.target.value);
                    if (bodyContent) {
                      const tagsToUse = e.target.value
                        ? e.target.value
                            .split(",")
                            .map((tag) => tag.trim().toLowerCase())
                            .filter((tag) => tag)
                        : ALLOWED_TAGS;
                      const attrsToUse = customAttrs
                        ? customAttrs
                            .split(",")
                            .map((attr) => attr.trim().toLowerCase())
                            .filter((attr) => attr)
                        : ALLOWED_ATTRS;
                      const sanitized = sanitizeHtml(
                        wordCleaned,
                        tagsToUse,
                        attrsToUse,
                        removeSpanBr
                      );
                      setBodyContent(sanitized);
                      let finalOutput;
                      if (outputType === "jsx") {
                        finalOutput =
                          outputFormat === "complete"
                            ? generateCompleteJsx(
                                sanitized,
                                "ConvertedComponent"
                              )
                            : htmlToJsx(sanitized);
                      } else {
                        finalOutput =
                          outputFormat === "complete"
                            ? generateCompleteHtml(
                                sanitized,
                                "Converted Document"
                              )
                            : sanitized;
                      }
                      setHtmlOutput(beautifyHtml(finalOutput));
                    }
                  }}
                  placeholder="e.g., div,p,h1,h2"
                  className="mt-1 w-full px-3 py-1 border border-gray-300 rounded text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 dark:border-gray-600 dark:bg-gray-700 dark:text-white"
                />
              </div>
              <div>
                <label className="text-sm font-medium text-gray-700 dark:text-gray-200">
                  Custom Allowed Attributes (comma-separated):
                </label>
                <input
                  type="text"
                  value={customAttrs}
                  onChange={(e) => {
                    setCustomAttrs(e.target.value);
                    if (bodyContent) {
                      const tagsToUse = customTags
                        ? customTags
                            .split(",")
                            .map((tag) => tag.trim().toLowerCase())
                            .filter((tag) => tag)
                        : ALLOWED_TAGS;
                      const attrsToUse = e.target.value
                        ? e.target.value
                            .split(",")
                            .map((attr) => attr.trim().toLowerCase())
                            .filter((attr) => attr)
                        : ALLOWED_ATTRS;
                      const sanitized = sanitizeHtml(
                        wordCleaned,
                        tagsToUse,
                        attrsToUse,
                        removeSpanBr
                      );
                      setBodyContent(sanitized);
                      let finalOutput;
                      if (outputType === "jsx") {
                        finalOutput =
                          outputFormat === "complete"
                            ? generateCompleteJsx(
                                sanitized,
                                "ConvertedComponent"
                              )
                            : htmlToJsx(sanitized);
                      } else {
                        finalOutput =
                          outputFormat === "complete"
                            ? generateCompleteHtml(
                                sanitized,
                                "Converted Document"
                              )
                            : sanitized;
                      }
                      setHtmlOutput(beautifyHtml(finalOutput));
                    }
                  }}
                  placeholder="e.g., href,src,alt"
                  className="mt-1 w-full px-3 py-1 border border-gray-300 rounded text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 dark:border-gray-600 dark:bg-gray-700 dark:text-white"
                />
              </div>
            </div>
          </div>

          {htmlOutput && (
            <div className="border rounded-lg overflow-hidden">
              {activeTab === "preview" ? (
                <div
                  className={`p-6 min-h-[400px] ${
                    isDarkMode ? "dark-mode" : "bg-white"
                  }`}
                >
                  <div
                    className="prose max-w-none"
                    dangerouslySetInnerHTML={{ __html: bodyContent }}
                    style={{
                      fontFamily: "system-ui, -apple-system, sans-serif",
                      lineHeight: "1.6",
                      color: isDarkMode ? "#e0e0e0" : "#333",
                      fontSize: `${fontSize}px`,
                    }}
                  />
                </div>
              ) : (
                <div className="bg-gray-900 text-gray-100 p-4">
                  <pre className="whitespace-pre-wrap font-mono text-sm overflow-x-auto">
                    <code>{htmlOutput}</code>
                  </pre>
                </div>
              )}
            </div>
          )}

          {htmlOutput && (
            <div className="mt-4 text-sm text-gray-600 dark:text-gray-400 flex flex-wrap gap-4">
              <span>Characters: {htmlOutput.length}</span>
              <span>Lines: {htmlOutput.split("\n").length}</span>
              <span>Tags: {(htmlOutput.match(/<[^>]+>/g) || []).length}</span>
            </div>
          )}
        </div>

        <div className="mt-8 bg-white dark:bg-gray-900 rounded-lg shadow-lg p-6">
          <h3 className="text-lg font-semibold mb-4 text-gray-800 dark:text-gray-200">
            How to Use
          </h3>
          <div className="grid md:grid-cols-2 gap-6">
            <div>
              <h4 className="font-medium text-gray-700 dark:text-gray-200 mb-2">
                📝 From Microsoft Word
              </h4>
              <ol className="list-decimal list-inside text-sm text-gray-600 dark:text-gray-400 space-y-1">
                <li>Select content in your Word document</li>
                <li>Copy it (Ctrl+C or Cmd+C)</li>
                <li>Paste into the input area above</li>
                <li>View the cleaned HTML output</li>
              </ol>
            </div>
            <div>
              <h4 className="font-medium text-gray-700 dark:text-gray-200 mb-2">
                ✨ Features
              </h4>
              <ul className="list-disc list-inside text-sm text-gray-600 dark:text-gray-400 space-y-1">
                <li>Convert to HTML or React JSX components</li>
                <li>Removes Word-specific formatting</li>
                <li>Option to remove &lt;span&gt; and &lt;br&gt; tags</li>
                <li>Customizable allowed tags and attributes</li>
                <li>Adjustable font size in preview</li>
                <li>Light/Dark mode for preview</li>
                <li>Cleans unnecessary attributes</li>
                <li>Beautifies code structure</li>
                <li>Live preview of converted content</li>
                <li>JSX converts HTML attributes (class → className)</li>
                <li>Self-closing tags for JSX compatibility</li>
                <li>Copy to clipboard or download as file</li>
              </ul>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
// import { Code, Copy, Download, Eye, RefreshCw } from "lucide-react";
// import { useState } from "react";

// // Enhanced sanitizer for Word documents
// const ALLOWED_TAGS = [
//   "div",
//   "span",
//   "p",
//   "b",
//   "i",
//   "u",
//   "strong",
//   "em",
//   "br",
//   "hr",
//   "h1",
//   "h2",
//   "h3",
//   "h4",
//   "h5",
//   "h6",
//   "ul",
//   "ol",
//   "li",
//   "dl",
//   "dt",
//   "dd",
//   "table",
//   "thead",
//   "tbody",
//   "tfoot",
//   "tr",
//   "td",
//   "th",
//   "a",
//   "img",
//   "blockquote",
//   "pre",
//   "code",
//   "sup",
//   "sub",
// ];

// const ALLOWED_ATTRS = ["href", "src", "alt", "title", "class", "id"];

// // Word-specific cleanup patterns
// const WORD_CLEANUP_PATTERNS = [
//   // Remove Word-specific namespaces and tags
//   /<o:p\s*\/?>|<\/o:p>/gi,
//   /<w:[^>]*>|<\/w:[^>]*>/gi,
//   /<m:[^>]*>|<\/m:[^>]*>/gi,
//   // Remove Word-specific attributes
//   /mso-[^=]*="[^"]*"/gi,
//   /lang="[^"]*"/gi,
//   // Remove empty style attributes
//   /style="[^"]*"/gi,
//   // Remove Word-specific classes
//   /class="Mso[^"]*"/gi,
//   // Remove XML declarations
//   /<\?xml[^>]*>/gi,
//   // Remove Word document properties
//   /<meta[^>]*>/gi,
//   // Remove empty paragraphs with only whitespace
//   /<p[^>]*>\s*<\/p>/gi,
//   // Remove span tags with no content or styling
//   /<span[^>]*>\s*<\/span>/gi,
// ];

// function cleanWordHtml(html: any) {
//   let cleaned = html;

//   // Apply Word-specific cleanup patterns
//   WORD_CLEANUP_PATTERNS.forEach((pattern) => {
//     cleaned = cleaned.replace(pattern, "");
//   });

//   // Remove excessive whitespace
//   cleaned = cleaned.replace(/\s+/g, " ");
//   cleaned = cleaned.replace(/>\s+</g, "><");

//   return cleaned;
// }

// function sanitizeHtml(
//   html: any,
//   allowedTags = ALLOWED_TAGS,
//   allowedAttrs = ALLOWED_ATTRS
// ) {
//   const parser = new DOMParser();
//   const doc = parser.parseFromString(html, "text/html");

//   function clean(node: any) {
//     if (node.nodeType === 1) {
//       const el = node;
//       if (!allowedTags.includes(el.tagName.toLowerCase())) {
//         // Replace with its content instead of removing
//         const parent = el.parentNode;
//         while (el.firstChild) {
//           parent.insertBefore(el.firstChild, el);
//         }
//         el.remove();
//         return;
//       }

//       // Clean attributes
//       [...el.attributes].forEach((attr) => {
//         if (!allowedAttrs.includes(attr.name.toLowerCase())) {
//           el.removeAttribute(attr.name);
//         }
//       });
//     }
//     [...node.childNodes].forEach(clean);
//   }

//   doc.body.childNodes.forEach(clean);
//   return doc.body.innerHTML;
// }

// // JSX Converter Functions
// function htmlToJsx(html: any) {
//   let jsx = html;

//   // Convert HTML attributes to JSX attributes
//   const attributeMap = {
//     class: "className",
//     for: "htmlFor",
//     tabindex: "tabIndex",
//     readonly: "readOnly",
//     maxlength: "maxLength",
//     cellpadding: "cellPadding",
//     cellspacing: "cellSpacing",
//     rowspan: "rowSpan",
//     colspan: "colSpan",
//     usemap: "useMap",
//     frameborder: "frameBorder",
//     contenteditable: "contentEditable",
//   };

//   // Replace HTML attributes with JSX equivalents
//   Object.entries(attributeMap).forEach(([htmlAttr, jsxAttr]) => {
//     const regex = new RegExp(`\\b${htmlAttr}=`, "gi");
//     jsx = jsx.replace(regex, `${jsxAttr}=`);
//   });

//   // Convert style attributes to JSX style objects
//   jsx = jsx.replace(/style="([^"]+)"/g, (match: any, styleContent: any) => {
//     const styleObj = styleContent
//       .split(";")
//       .filter((prop: any) => prop.trim())
//       .map((prop: any) => {
//         const [key, value] = prop.split(":").map((s: any) => s.trim());
//         if (!key || !value) return "";

//         // Convert kebab-case to camelCase
//         const camelKey = key.replace(/-([a-z])/g, (match: any, letter: any) =>
//           letter.toUpperCase()
//         );

//         // Handle numeric values
//         const processedValue = /^\d+$/.test(value) ? value : `'${value}'`;

//         return `${camelKey}: ${processedValue}`;
//       })
//       .filter((prop: any) => prop)
//       .join(", ");

//     return styleObj ? `style={{${styleObj}}}` : "";
//   });

//   // Self-closing tags
//   const selfClosingTags = ["br", "hr", "img", "input", "meta", "link"];
//   selfClosingTags.forEach((tag) => {
//     const regex = new RegExp(`<${tag}([^>]*?)>`, "gi");
//     jsx = jsx.replace(regex, (match: any, attributes: any) => {
//       return `<${tag}${attributes} />`;
//     });
//   });

//   // Convert comments to JSX comments
//   jsx = jsx.replace(/<!--([\s\S]*?)-->/g, "{/* $1 */}");

//   return jsx;
// }

// function generateCompleteJsx(
//   bodyContent: any,
//   componentName = "ConvertedComponent"
// ) {
//   const jsxContent = htmlToJsx(bodyContent);

//   return `import React from 'react';

// const ${componentName} = () => {
//   return (
//     <>
// ${jsxContent
//   .split("\n")
//   .map((line: any) => "      " + line)
//   .join("\n")}
//     </>
//   );
// };

// export default ${componentName};`;
// }

// function generateCompleteHtml(bodyContent: any, title = "Converted Document") {
//   return `<!DOCTYPE html>
// <html lang="en">
// <head>
//     <meta charset="UTF-8">
//     <meta name="viewport" content="width=device-width, initial-scale=1.0">
//     <title>${title}</title>
//     <style>
//         body {
//             font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
//             line-height: 1.6;
//             max-width: 800px;
//             margin: 0 auto;
//             padding: 20px;
//             color: #333;
//         }
//         h1, h2, h3, h4, h5, h6 {
//             margin-top: 1.5em;
//             margin-bottom: 0.5em;
//             font-weight: 600;
//         }
//         p {
//             margin-bottom: 1em;
//         }
//         ul, ol {
//             margin-bottom: 1em;
//             padding-left: 2em;
//         }
//         table {
//             border-collapse: collapse;
//             width: 100%;
//             margin-bottom: 1em;
//         }
//         th, td {
//             border: 1px solid #ddd;
//             padding: 8px;
//             text-align: left;
//         }
//         th {
//             background-color: #f2f2f2;
//             font-weight: 600;
//         }
//         blockquote {
//             border-left: 4px solid #ddd;
//             margin-left: 0;
//             padding-left: 1em;
//             color: #666;
//         }
//         code {
//             background-color: #f4f4f4;
//             padding: 2px 4px;
//             border-radius: 3px;
//             font-family: monospace;
//         }
//         pre {
//             background-color: #f4f4f4;
//             padding: 1em;
//             border-radius: 5px;
//             overflow-x: auto;
//         }
//     </style>
// </head>
// <body>
// ${bodyContent}
// </body>
// </html>`;
// }

// function beautifyHtml(html: any, indentSize = 2) {
//   let formatted = "";
//   const pad = " ".repeat(indentSize);
//   let depth = 0;

//   // Split by tags but keep them
//   const tokens = html
//     .split(/(<[^>]*>)/g)
//     .filter((token: any) => token.length > 0);

//   tokens.forEach((token: any) => {
//     if (token.match(/<\/\w+>/)) {
//       // Closing tag
//       depth--;
//       if (formatted.endsWith("\n")) {
//         formatted += pad.repeat(Math.max(0, depth)) + token + "\n";
//       } else {
//         formatted += token + "\n";
//       }
//     } else if (token.match(/<\w+[^>]*>/)) {
//       // Opening tag
//       if (formatted.endsWith("\n")) {
//         formatted += pad.repeat(depth) + token;
//       } else {
//         formatted += "\n" + pad.repeat(depth) + token;
//       }

//       if (!token.match(/<\w+[^>]*\/>/)) {
//         depth++;
//       }

//       if (token.match(/<(br|hr|img|input)[^>]*>/i)) {
//         formatted += "\n";
//       }
//     } else if (token.trim()) {
//       // Text content
//       if (formatted.endsWith(">")) {
//         formatted += token;
//       } else {
//         formatted += token;
//       }
//     }
//   });

//   return formatted.trim();
// }

// export default function WordToHtmlConverter() {
//   const [htmlOutput, setHtmlOutput] = useState("");
//   const [bodyContent, setBodyContent] = useState("");
//   const [activeTab, setActiveTab] = useState("preview");
//   const [isProcessing, setIsProcessing] = useState(false);
//   const [outputFormat, setOutputFormat] = useState("complete"); // "complete" or "body-only"
//   const [outputType, setOutputType] = useState("html"); // "html" or "jsx"

//   const handlePaste = async (e: any) => {
//     e.preventDefault();
//     setIsProcessing(true);

//     try {
//       const clipboardData = e.clipboardData;
//       let pastedHTML = clipboardData.getData("text/html");

//       if (!pastedHTML) {
//         // Fallback to plain text
//         const pastedText = clipboardData.getData("text/plain");
//         pastedHTML = pastedText
//           .split("\n")
//           .map((line: any) => (line.trim() ? `<p>${line}</p>` : ""))
//           .join("");
//       }

//       // Clean Word-specific formatting
//       const wordCleaned = cleanWordHtml(pastedHTML);

//       // Sanitize HTML
//       const sanitized = sanitizeHtml(wordCleaned);

//       // Store body content for preview
//       setBodyContent(sanitized);

//       // Generate output based on type and format selection
//       let finalOutput;

//       if (outputType === "jsx") {
//         finalOutput =
//           outputFormat === "complete"
//             ? generateCompleteJsx(sanitized, "ConvertedComponent")
//             : htmlToJsx(sanitized);
//       } else {
//         finalOutput =
//           outputFormat === "complete"
//             ? generateCompleteHtml(sanitized, "Converted Document")
//             : sanitized;
//       }

//       // Beautify the output
//       const beautified = beautifyHtml(finalOutput);

//       setHtmlOutput(beautified);

//       // Auto-switch to preview tab
//       setActiveTab("preview");
//     } catch (error) {
//       console.error("Error processing pasted content:", error);
//       alert("Error processing pasted content. Please try again.");
//     } finally {
//       setIsProcessing(false);
//     }
//   };

//   const handleCopy = async () => {
//     try {
//       await navigator.clipboard.writeText(htmlOutput);
//       alert("HTML copied to clipboard!");
//     } catch (error) {
//       console.error("Failed to copy:", error);
//       alert("Failed to copy to clipboard");
//     }
//   };

//   const handleDownload = () => {
//     const fileExtension = outputType === "jsx" ? "jsx" : "html";
//     const mimeType = outputType === "jsx" ? "text/javascript" : "text/html";
//     const fileName =
//       outputType === "jsx" ? "ConvertedComponent.jsx" : "converted.html";

//     const blob = new Blob([htmlOutput], { type: mimeType });
//     const url = URL.createObjectURL(blob);
//     const a = document.createElement("a");
//     a.href = url;
//     a.download = fileName;
//     document.body.appendChild(a);
//     a.click();
//     document.body.removeChild(a);
//     URL.revokeObjectURL(url);
//   };

//   const handleClear = () => {
//     setHtmlOutput("");
//     setBodyContent("");
//   };

//   return (
//     <div className="min-h-screen bg-gray-50 p-6">
//       <div className="max-w-6xl mx-auto">
//         <div className="bg-white rounded-lg shadow-lg overflow-hidden">
//           {/* Header */}
//           <div className="bg-gradient-to-r from-blue-600 to-purple-600 text-white p-6">
//             <h1 className="text-3xl font-bold mb-2">
//               Word to HTML/JSX Converter
//             </h1>
//             <p className="text-blue-100">
//               Convert Word documents to clean HTML or React JSX components with
//               live preview
//             </p>
//           </div>

//           {/* Input Section */}
//           <div className="p-6 border-b border-gray-200">
//             <h2 className="text-xl font-semibold mb-4 text-gray-800">
//               Paste Your Content
//             </h2>
//             <div
//               contentEditable
//               onPaste={handlePaste}
//               className="border-2 border-dashed border-gray-300 rounded-lg p-6 min-h-[200px] focus:border-blue-500 focus:outline-none transition-colors"
//               style={{ minHeight: "200px" }}
//             >
//               {isProcessing ? (
//                 <div className="flex items-center justify-center text-gray-500">
//                   <RefreshCw className="animate-spin mr-2" size={20} />
//                   Processing...
//                 </div>
//               ) : (
//                 <div className="text-gray-500">
//                   <p className="text-lg mb-2">
//                     📋 Paste your Word content here
//                   </p>
//                   <p className="text-sm">
//                     Copy content from Word and paste it here. The converter will
//                     automatically clean and format the HTML.
//                   </p>
//                 </div>
//               )}
//             </div>
//           </div>

//           {/* Output Section */}
//           {htmlOutput && (
//             <div className="p-6">
//               {/* Controls */}
//               <div className="flex flex-wrap items-center justify-between mb-4 gap-4">
//                 <div className="flex space-x-2">
//                   <button
//                     onClick={() => setActiveTab("preview")}
//                     className={`flex items-center px-4 py-2 rounded-lg transition-colors ${
//                       activeTab === "preview"
//                         ? "bg-blue-500 text-white"
//                         : "bg-gray-200 text-gray-700 hover:bg-gray-300"
//                     }`}
//                   >
//                     <Eye size={16} className="mr-2" />
//                     Preview
//                   </button>
//                   <button
//                     onClick={() => setActiveTab("code")}
//                     className={`flex items-center px-4 py-2 rounded-lg transition-colors ${
//                       activeTab === "code"
//                         ? "bg-blue-500 text-white"
//                         : "bg-gray-200 text-gray-700 hover:bg-gray-300"
//                     }`}
//                   >
//                     <Code size={16} className="mr-2" />
//                     {outputType === "jsx" ? "JSX Code" : "HTML Code"}
//                   </button>
//                 </div>

//                 <div className="flex items-center space-x-4">
//                   {/* Output Type Toggle */}
//                   <div className="flex items-center space-x-2">
//                     <label className="text-sm font-medium text-gray-700">
//                       Type:
//                     </label>
//                     <select
//                       value={outputType}
//                       onChange={(e) => {
//                         setOutputType(e.target.value);
//                         // Regenerate output with new type
//                         if (bodyContent) {
//                           let finalOutput;
//                           if (e.target.value === "jsx") {
//                             finalOutput =
//                               outputFormat === "complete"
//                                 ? generateCompleteJsx(
//                                     bodyContent,
//                                     "ConvertedComponent"
//                                   )
//                                 : htmlToJsx(bodyContent);
//                           } else {
//                             finalOutput =
//                               outputFormat === "complete"
//                                 ? generateCompleteHtml(
//                                     bodyContent,
//                                     "Converted Document"
//                                   )
//                                 : bodyContent;
//                           }
//                           setHtmlOutput(beautifyHtml(finalOutput));
//                         }
//                       }}
//                       className="px-3 py-1 border border-gray-300 rounded text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
//                     >
//                       <option value="html">HTML</option>
//                       <option value="jsx">JSX/React</option>
//                     </select>
//                   </div>

//                   {/* Output Format Toggle */}
//                   <div className="flex items-center space-x-2">
//                     <label className="text-sm font-medium text-gray-700">
//                       Format:
//                     </label>
//                     <select
//                       value={outputFormat}
//                       onChange={(e) => {
//                         setOutputFormat(e.target.value);
//                         // Regenerate output with new format
//                         if (bodyContent) {
//                           let finalOutput;
//                           if (outputType === "jsx") {
//                             finalOutput =
//                               e.target.value === "complete"
//                                 ? generateCompleteJsx(
//                                     bodyContent,
//                                     "ConvertedComponent"
//                                   )
//                                 : htmlToJsx(bodyContent);
//                           } else {
//                             finalOutput =
//                               e.target.value === "complete"
//                                 ? generateCompleteHtml(
//                                     bodyContent,
//                                     "Converted Document"
//                                   )
//                                 : bodyContent;
//                           }
//                           setHtmlOutput(beautifyHtml(finalOutput));
//                         }
//                       }}
//                       className="px-3 py-1 border border-gray-300 rounded text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
//                     >
//                       <option value="complete">
//                         {outputType === "jsx"
//                           ? "Complete Component"
//                           : "Complete HTML"}
//                       </option>
//                       <option value="body-only">
//                         {outputType === "jsx"
//                           ? "JSX Content Only"
//                           : "Body Content Only"}
//                       </option>
//                     </select>
//                   </div>

//                   <div className="flex space-x-2">
//                     <button
//                       onClick={handleCopy}
//                       className="flex items-center px-4 py-2 bg-green-500 text-white rounded-lg hover:bg-green-600 transition-colors"
//                     >
//                       <Copy size={16} className="mr-2" />
//                       Copy {outputType === "jsx" ? "JSX" : "HTML"}
//                     </button>
//                     <button
//                       onClick={handleDownload}
//                       className="flex items-center px-4 py-2 bg-purple-500 text-white rounded-lg hover:bg-purple-600 transition-colors"
//                     >
//                       <Download size={16} className="mr-2" />
//                       Download
//                     </button>
//                     <button
//                       onClick={handleClear}
//                       className="flex items-center px-4 py-2 bg-red-500 text-white rounded-lg hover:bg-red-600 transition-colors"
//                     >
//                       <RefreshCw size={16} className="mr-2" />
//                       Clear
//                     </button>
//                   </div>
//                 </div>
//               </div>

//               {/* Output Display */}
//               <div className="border rounded-lg overflow-hidden">
//                 {activeTab === "preview" ? (
//                   <div className="bg-white p-6 min-h-[400px]">
//                     <div
//                       className="prose max-w-none"
//                       dangerouslySetInnerHTML={{ __html: bodyContent }}
//                       style={{
//                         fontFamily: "system-ui, -apple-system, sans-serif",
//                         lineHeight: "1.6",
//                         color: "#333",
//                       }}
//                     />
//                   </div>
//                 ) : (
//                   <div className="bg-gray-900 text-gray-100 p-4">
//                     <pre className="whitespace-pre-wrap font-mono text-sm overflow-x-auto">
//                       <code>{htmlOutput}</code>
//                     </pre>
//                   </div>
//                 )}
//               </div>

//               {/* Stats */}
//               <div className="mt-4 text-sm text-gray-600 flex flex-wrap gap-4">
//                 <span>Characters: {htmlOutput.length}</span>
//                 <span>Lines: {htmlOutput.split("\n").length}</span>
//                 <span>Tags: {(htmlOutput.match(/<[^>]+>/g) || []).length}</span>
//               </div>
//             </div>
//           )}
//         </div>

//         {/* Instructions */}
//         <div className="mt-8 bg-white rounded-lg shadow-lg p-6">
//           <h3 className="text-lg font-semibold mb-4 text-gray-800">
//             How to Use
//           </h3>
//           <div className="grid md:grid-cols-2 gap-6">
//             <div>
//               <h4 className="font-medium text-gray-700 mb-2">
//                 📝 From Microsoft Word
//               </h4>
//               <ol className="list-decimal list-inside text-sm text-gray-600 space-y-1">
//                 <li>Select content in your Word document</li>
//                 <li>Copy it (Ctrl+C or Cmd+C)</li>
//                 <li>Paste into the input area above</li>
//                 <li>View the cleaned HTML output</li>
//               </ol>
//             </div>
//             <div>
//               <h4 className="font-medium text-gray-700 mb-2">✨ Features</h4>
//               <ul className="list-disc list-inside text-sm text-gray-600 space-y-1">
//                 <li>Convert to HTML or React JSX components</li>
//                 <li>Removes Word-specific formatting</li>
//                 <li>Cleans unnecessary attributes</li>
//                 <li>Beautifies code structure</li>
//                 <li>Live preview of converted content</li>
//                 <li>JSX converts HTML attributes (class → className)</li>
//                 <li>Self-closing tags for JSX compatibility</li>
//                 <li>Copy to clipboard or download as file</li>
//               </ul>
//             </div>
//           </div>
//         </div>
//       </div>
//     </div>
//   );
// }

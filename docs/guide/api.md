---
lang: en-US
title: Library Entry Points
description: Guide for selecting the relevant parsing method.
---

## Input Scope and Tag Compliance

The library has been designed for maximum compatibility, gracefully handling a wide range of HTML constructs.

* **Broad Support**: The library provides faithful conversion for a wide range of HTML constructs, including both legacy (e.g., `font`) and modern semantic tags (`article`).

* **Ignored Elements**: For cleaner output, certain elements are intentionally ignored or passed through without action. This includes interactive form controls (`button`, `input`, `select`) and document metadata tags (`script`, `head`, `xml`).

## API Entry Points

The library provides multiple methods to drive the conversion. Your choice depends on whether you are **building a document**, **populating a template**, or simply **translating content**.

<llm-only>
@c HtmlConverter "Primary entry point for converting HTML into OpenXml Word document. The converter can be used with newly created documents as well as existing document templates. When a template is used, styles, themes, numbering definitions, bookmarks, and other Word settings are automatically reused. Supports both left-to-right (LTR) and right-to-left (LTR) visual flow content.
 .ctor (MainDocumentPart mainPart, IWebRequest? webRequester = null) "Create a converter bound to a Word document. Reuse the same HtmlConverter instance for the lifetime of a document to avoid reloading cached configuration such as styles and bookmarks. Do not use the same converter instance with multiple documents."
 .m Parse:IList&lt;OpenXmlCompositeElement&gt;(string html) "Convert HTML into OpenXml elements and return them to the caller. Use this method when your HTML is simple and don't need to download any external resources. Parse() is equivalent to ParseAsync() and is retained for backward compatibility. New code should prefer ParseBody(), ParseHeader() or ParseFooter() to make the target document section explicit."
 .m ParseAsync:Task&lt;IEnumerable&gt;OpenXmlCompositeElement&gt;&gt;(string html, CancellationToken cancellationToken = null) "Convert HTML into OpenXml elements and return them to the caller. However, the conversion process itself may still update the underlying document by creating styles, numbering definitions, image parts, bookmarks, or other resources required by the generated content. Use this method when you need full control over the insertion point or need to inspect or modify the generated elements before adding them to the document. For most scenarios, prefer ParseBody()."
 .m ParseAsync:Task&lt;IEnumerable&gt;OpenXmlCompositeElement&gt;&gt;(string html, ParallelOptions parallelOptions) "Convert HTML into OpenXml elements and return them to the caller. Use this method when you need full control over the insertion point or need to inspect or modify the generated elements before adding them to the document. For most scenarios, prefer ParseBody()."
</llm-only>

### `ParseBody()`

* **Purpose:** Appends the converted HTML content into the document's main body part (`MainDocumentPart`). This is the most common use case.

* **Use Case:** Generating the primary flow of a document (e.g., reports, contracts).

### `ParseHeader()` / `ParseFooter()`

* **Purpose:** Inserts the converted HTML into the dedicated header or footer section. This content is used to create repeating elements like page numbers and chapter titles across the document body.

* **Use Case:** Applying consistent branding or required metadata to every page.

### `Parse()` / `ParseAsync()`

Purpose: Converts HTML into a standalone, actionable collection of OpenXML elements. This output is not automatically attached to any document.

* **Purpose:** Converts the HTML into a list of independent OpenXML elements. This output is ready to be inserted into any part of the DOCX package, offering you full control over its insertion point and allowing mid-conversion inspection or modification.

* **Use Case:** You are building a custom element, merging snippets from multiple sources, or needing to inspect and modify the generated code before injecting it into a document.

* **Note on Synchronicity**: Use `Parse()` for basic, non-resource-dependent conversions. Use `ParseAsync()` when the content requires external resource fetching (images, links) or longer processing time.

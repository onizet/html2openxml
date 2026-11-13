---
url: /html2openxml/guide/api.md
description: Guide for selecting the relevant parsing method.
---
## Input Scope and Tag Compliance

The library has been designed for maximum compatibility, gracefully handling a wide range of HTML constructs.

* **Broad Support**: The library provides faithful conversion for a wide range of HTML constructs, including both legacy (e.g., `font`) and modern semantic tags (`article`).

* **Ignored Elements**: For cleaner output, certain elements are intentionally ignored or passed through without action. This includes interactive form controls (`button`, `input`, `select`) and document metadata tags (`script`, `head`, `xml`).

## API Entry Points

The library provides multiple methods to drive the conversion. Your choice depends on whether you are **building a document**, **populating a template**, or simply **translating content**.

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

## Advanced Extension Hooks (Optional)

`HtmlConverter` can be extended by overriding two virtual methods that participate in the HTML parsing pipeline.
These hooks are intended for advanced scenarios such as CSS inlining, DOM transformations, custom AngleSharp integrations, URL rewriting, or document validation before conversion.

### `GetParserConfiguration()`

* **Purpose:** Provides the `AngleSharp` configuration used to parse the input HTML before conversion begins.

* **Use Case:** Override this method to register AngleSharp extensions, services, or features required during
  document parsing. Typical scenario include enabling support for CSS processing, custom resource loaders, or other AngleSharp integrations.

* **Default Behavior:** Returns `Configuration.Default`, providing standard HTML parsing behaviour.

```csharp
protected override IConfiguration GetParserConfiguration()
{
    return Configuration.Default
        .WithCss();
}
```

### `OnDocumentLoadedAsync()`

* **Purpose:** Executes after the HTML has been parsed into an AngleSharp document but before any OpenXML elements or image resources are generated.

* **Use Case:** Inspect or alter the DOM before conversion. Typical sccenarios include:
  * Inlining CSS stylesheets
  * Rewriting image URLs
  * Removing unsupported elements
  * Adding custom attributes
  * Performing document validation

* **Default Behavior:** Performs no action.

* **Note:** Any changes made to the DOM are used by all subsequent conversion stages, including the image discovery,
  resource loading, and OpenXML conversion.

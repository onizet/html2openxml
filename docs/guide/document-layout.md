---
lang: en-US
title: Document Layout
description: Generate cover pages, tables of contents, page breaks, and custom page layouts.
---

`HtmlToOpenXml` provides helpers for assembling professional Word documents with cover pages, tables of contents, page breaks, and mixed page orientations.

These features integrate with Word's native document structure and can be combined with HTML conversion, templates, bookmarks, and heading styles.

## Cover Page

A cover page is stored as a Word content control and always appears at the beginning of the document.

### Create or Update a Cover Page

```csharp
HtmlConverter converter = new(mainPart);

await converter.SetCoverPage("""
    <p>Annual Report 2026</p>
    <p>Confidential</p>
""");
```

If a cover page already exists, its content is replaced.

### Retrieve the Existing Cover Page

```csharp
var coverPage = converter.GetCoverPage();
```

Returns `null` when no cover page is present.

---

## Table of Contents

`HtmlToOpenXml` can generate a native Microsoft Word table of contents from document headings.

### Create or Update a Table of Contents

```csharp
converter.SetTableOfContents();
```

By default, the table of contents is inserted immediately after the cover page.

### Limit the Heading Depth

```csharp
converter.SetTableOfContents(maxLevel: 3);
```

This includes:

- Heading 1
- Heading 2
- Heading 3

Headings below the specified level are ignored.

---

## Positioning

The insertion location can be controlled using the `DocumentPosition` enumeration.

```csharp
converter.SetTableOfContents(
    maxLevel: 3,
    position: DocumentPosition.AfterCover);
```

### Available Positions

```csharp
DocumentPosition.DocumentStart
DocumentPosition.AfterCover
DocumentPosition.DocumentEnd
```

| Position | Description |
|-----------|-------------|
| `DocumentStart` | Inserts the table of contents at the beginning of the document. |
| `AfterCover` | Inserts the table of contents immediately after the cover page. If no cover page exists, the TOC is inserted at the beginning of the document. |
| `DocumentEnd` | Inserts the table of contents at the end of the document. |

---

## Page Breaks

Page breaks can be inserted before or after any OpenXml element.

### Insert a Page Break Before an Element

```csharp
var toc = converter.SetTableOfContents();

converter.AddPageBreakBefore(toc);
```

### Insert a Page Break After an Element

```csharp
var toc = converter.SetTableOfContents();

converter.AddPageBreakAfter(toc);
```

A common scenario is to start the report content on a new page immediately after the table of contents:

```csharp
var toc = converter.SetTableOfContents();

converter.AddPageBreakAfter(toc);
```

---

## Page Orientation and Paper Size

Documents often require a mixture of portrait and landscape pages. This is particularly useful for:

- Wide tables
- Financial statements
- Large diagrams
- Charts
- Technical appendices

### Add a Landscape Page

```csharp
converter.AddPage(PageOrientation.Landscape);
```

### Add a Portrait Page

```csharp
converter.AddPage(PageOrientation.Portrait);
```

A new page is always created, even when the requested orientation already matches the current page orientation.

---

## Paper Size

When creating a new section, a paper size can also be specified.

```csharp
converter.AddPage(
    PageOrientation.Landscape,
    PaperSize.A3);
```

### Supported Paper Sizes

```csharp
PaperSize.A3
PaperSize.A4
PaperSize.A5
PaperSize.A6

PaperSize.B3
PaperSize.B4
PaperSize.B5
PaperSize.B6

PaperSize.Letter
PaperSize.Legal
```

---

## Creating a New Document

`WordDocumentFactory.Create()` can initialize a document using a specific paper size and orientation.

### A4 Portrait

```csharp
using var package = WordDocumentFactory.Create(
    stream,
    PageOrientation.Portrait,
    PaperSize.A4);
```

### Letter Landscape

```csharp
using var package = WordDocumentFactory.Create(
    stream,
    PageOrientation.Landscape,
    PaperSize.Letter);
```

---

## Typical Annual Report Layout

```csharp
await converter.SetCoverPage("""
    <p>Annual Report 2026</p>
""");

await converter.ParseBody("""
    <h1>Executive Summary</h1>

    <h1>Company Overview</h1>
    <h2>Mission</h2>
    <h2>Key Figures</h2>

    <h1>Financial Performance</h1>
    <h2>Revenue</h2>
    <h2>Expenses</h2>
    <h2>Profitability</h2>

    <h1>Future Outlook</h1>

    <h1>Appendices</h1>
""");

// Generate the TOC after all headings have been created
var toc = converter.SetTableOfContents();

// Start the report body on a new page
converter.AddPageBreakAfter(toc);
```

Result:

```text
Cover Page
Table of Contents
Document Content
```

---

## How Headings Are Detected

The table of contents is generated from Word heading styles.

The converter automatically maps:

```html
<h1>Introduction</h1>
<h2>Overview</h2>
<h3>Details</h3>
```

to:

```text
Heading 1
Heading 2
Heading 3
```

Only headings whose level is less than or equal to `maxLevel` are included in the generated table of contents.

---

## Hyperlinks

Each TOC entry is generated as an internal hyperlink.

Clicking an entry in Word navigates directly to the corresponding heading.

Any required bookmarks are generated automatically during TOC creation.

---

## Page Numbers

`HtmlToOpenXml` intentionally does **not** generate page numbers inside table of contents entries.

Page numbering depends on Word's layout engine and is influenced by factors such as: installed fonts, page orientation, margins, printer configuration, ...

Because these values are not available during OpenXML generation, reliable page numbers cannot be computed in advance.

The generated TOC therefore focuses on document navigation and heading structure.

### Generate Page Numbers

If page numbers are required, refresh the table of contents from Microsoft Word:

```text
References → Update Table
```

or

```text
Right-click → Update Field
```

Word will calculate the final page numbers and regenerate the table contents.

---
lang: en-US
title: Cover Pages and Table of Contents
description: Document extra document structure feature.
---

`HtmlToOpenXml` can generate Microsoft Word cover pages and tables of contents directly from your document.

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

HtmlToOpenXml can generate a native Word table of contents from Word heading styles.

### Create or Update a Table of Contents

```csharp
converter.SetTableOfContents();
```

By default, the TOC is inserted immediately after the cover page.

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

The insertion location can be controlled using the `DocumentLocation` enumeration.

```csharp
converter.SetTableOfContents(
    maxLevel: 3,
    position: DocumentLocation.AfterCover);
```

### Available Locations

```csharp
DocumentLocation.DocumentStart
DocumentLocation.AfterCover
DocumentLocation.DocumentEnd
```

| Location | Description |
|-----------|-------------|
| `DocumentStart` | Inserts the TOC at the beginning of the document. |
| `AfterCover` | Inserts the TOC immediately after the cover page. If no cover page exists, the TOC is inserted at the beginning of the document. |
| `DocumentEnd` | Inserts the TOC at the end of the document. |

---

## Typical Document Layout

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

converter.SetTableOfContents();
```

Result:

```text
Cover Page
Table of Contents
Document Content
```

---

## Heading Detection

The TOC is generated from paragraphs using Word heading styles.

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

Only headings whose level is less than or equal to `maxLevel` are included.

---

## Hyperlinks

Each TOC entry is generated as an internal hyperlink.

Clicking an entry in Word navigates directly to the corresponding heading.

Any required bookmarks are created automatically during TOC generation.

---

## Page Numbers

HtmlToOpenXml intentionally does **not** generate page numbers in TOC entries.

Page numbers are determined by Word's layout engine and depend on factors such as:

- installed fonts
- page orientation
- margins
- printer settings
- Word version

Because these values are not available during OpenXML generation, reliable page numbers cannot be computed ahead of time.

The generated TOC therefore focuses on document navigation and heading structure.

### Generating Page Numbers

If page numbers are required, update the table of contents from Microsoft Word:

```text
References → Update Table
```

or

```text
Right-click → Update Field
```

Word will then calculate the correct page numbers and rebuild the table.

---

## Compatibility

The generated TOC is immediately usable after document generation:

- Heading links are clickable.
- Bookmarks are generated automatically.
- Cover pages are preserved.
- Word recognizes the TOC as a native table of contents.

Refreshing the TOC is only required when page numbers are needed.
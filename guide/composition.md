---
url: /html2openxml/guide/composition.md
description: Guide for advanced page layout.
---
To manage the visual flow and chapter structure of your generated Word document, use these CSS-like page layout attributes to control breaks and orientation.

## Page Break

Page breaks are controlled via the `page-break-before` or `page-break-after` attributes, which accept the value **always**.
These directional controls apply to various block elements (`p`, `div`, `pre`, `span`) or the document body.

```html
<div style="page-break-after:always">
    Lorem ipsum dolor sit amet,
    consectetur adipiscing elit.
</div>

Duis dictum leo quis ipsum tempor nec ultrices sapien elementum.
```

***

## Page Breaks

Page breaks can be inserted in the current flow of the Body.

```csharp
converter.AddPageBreak();
```

### HTML-Driven Page Breaks

Page breaks can also be controlled directly from HTML using standard CSS properties:

```html
<div style="page-break-before: always">
    Appendix A
</div>
```

```html
<div style="page-break-after: always">
    End of chapter
</div>
```

These properties are supported on common block elements such as:

* `div`
* `p`
* `pre`
* `span`
* `body`

***

## Page Orientation and Paper Size

Documents often require a mixture of portrait and landscape pages.

Typical scenarios include:

* Wide tables
* Financial statements
* Large diagrams
* Charts
* Technical appendices

### Add a Landscape Page

```csharp
converter.AddPage(PageOrientation.Landscape);
```

### Add a Portrait Page

```csharp
converter.AddPage(PageOrientation.Portrait);
```

A new page is always created, even when the requested orientation already matches the current page orientation.

### Specify a Paper Size

When creating a new page, a paper size can also be specified:

```csharp
converter.AddPage(
    PageOrientation.Landscape,
    PaperSize.A3);
```

#### Supported Paper Sizes

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

***

## Creating a New Document

`WordDocumentFactory.Create()` can initialize a document using a specific paper size and orientation.

### A4 Portrait

```csharp
using var document = WordDocumentFactory.Create(stream,
    PageOrientation.Portrait,
    PaperSize.A4);
```

### Letter Landscape

```csharp
using var document = WordDocumentFactory.Create(stream,
    PageOrientation.Landscape,
    PaperSize.Letter);
```

***

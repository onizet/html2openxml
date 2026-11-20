# HtmlToOpenXml&nbsp;[![Latest version](https://img.shields.io/nuget/v/HtmlToOpenXml.svg)](https://www.nuget.org/packages/HtmlToOpenXml/)&nbsp;[![Download Counts](https://img.shields.io/nuget/dt/HtmlToOpenXml.svg)](https://www.nuget.org/packages/HtmlToOpenXml/)&nbsp;[![MIT License](https://img.shields.io/badge/license-MIT-blue.svg)](https://github.com/onizet/html2openxml/blob/dev/LICENSE)

HtmlToOpenXml is a mature, high-fidelity converter that translates complex HTML structures into native OpenXML elements, ensuring your content renders exactly as intended across all Microsoft Word users.

Originally created in 2009 to transform user-generated content into templated Word documents, it has evolved
into a mature HTML-to-OpenXML converter supporting styles, numbering, images, bookmarks, page layout, tables,
document template, and custom resource resolution.

![Logo](https://raw.githubusercontent.com/onizet/html2openxml/master/docs/images/html2openxml.png)

## Features

- HTML → DOCX conversion
- Word template integration
- Cover pages
- Native Word table of contents generation
- Portrait and landscape page sections
- Headers and footers
- Tables with RowSpan and ColSpan support
- Numbered and bulleted lists
- Images (PNG, JPEG, SVG, Data URI)
- Internal hyperlinks and bookmarks
- Right-to-left (RTL) documents
- Markdown support (optional package)
- CSS stylesheet support (optional package)

Supports **.NET Framework 4.6.2**, **.NET Standard 2.0**, **.NET 8** and **.NET 10**.

Built on top of [DocumentFormat.OpenXml](https://www.nuget.org/packages/DocumentFormat.OpenXml/) and [AngleSharp](https://www.nuget.org/packages/AngleSharp).

| Package | NuGet Stable | Downloads |
| ------- | ------------ | --------- |
| [HtmlToOpenXml](https://www.nuget.org/packages/HtmlToOpenXml/) | [![HtmlToOpenXml](https://img.shields.io/nuget/v/HtmlToOpenXml.svg)](https://www.nuget.org/packages/HtmlToOpenXml/) | [![HtmlToOpenXml](https://img.shields.io/nuget/dt/HtmlToOpenXml.svg)](https://www.nuget.org/packages/HtmlToOpenXml/) |
| [HtmlToOpenXml.CssInliner](https://www.nuget.org/packages/HtmlToOpenXml.CssInliner/) | [![HtmlToOpenXml.CssInliner](https://img.shields.io/nuget/v/HtmlToOpenXml.CssInliner.svg)](https://www.nuget.org/packages/HtmlToOpenXml.CssInliner/) | [![HtmlToOpenXml.CssInliner](https://img.shields.io/nuget/dt/HtmlToOpenXml.CssInliner.svg)](https://www.nuget.org/packages/HtmlToOpenXml.CssInliner/) |
| [HtmlToOpenXml.Markdown](https://www.nuget.org/packages/HtmlToOpenXml.Markdown/) | [![HtmlToOpenXml.Markdown](https://img.shields.io/nuget/v/HtmlToOpenXml.Markdown.svg)](https://www.nuget.org/packages/HtmlToOpenXml.Markdown/) | [![HtmlToOpenXml.Markdown](https://img.shields.io/nuget/dt/HtmlToOpenXml.Markdown.svg)](https://www.nuget.org/packages/HtmlToOpenXml.Markdown/) |

> **Note**
> The package was renamed from `HtmlToOpenXml.dll` to `HtmlToOpenXml`.
> NuGet download statistics were unfortunately not transferred to the new package.

## Quick Start

When creating a blank document:

```csharp
await using var generatedDocument = new MemoryStream();
using var package = WordDocumentFactory.Create(generatedDocument);

HtmlConverter converter = new(package);
await converter.ParseBody(html);
```

Using an existing template:

```csharp
await using var generatedDocument = new MemoryStream();
await templateFileStream.CopyToAsync(generatedDocument);
generatedDocument.Position = 0L;
using var package = WordprocessingDocument.Open(generatedDocument, true);
HtmlConverter converter = new(package);
await converter.ParseBody(html);
```

Advanced document features:

```csharp
await converter.SetCoverPage("<h1>Annual Report 2026</h1>");

converter.AddPage(PageOrientation.Landscape);
await converter.ParseBody(wideTableHtml);

// generate the TOC after all document headings have been created
var toc = converter.SetTableOfContents();
converter.AddPageBreakAfter(toc);
```

## Documentation

[**Documentation**](https://onizet.github.io/html2openxml/) covers from the minimal code to advanced scenarios.

### Guides & Tutorials

- [Getting Started Guide](https://onizet.github.io/html2openxml/guide/quickstart.md)
- [LLM Guidance](https://onizet.github.io/html2openxml/llmdoc.md)

### API Reference

- [API Contract](https://onizet.github.io/html2openxml/guide/api.md)

### Document Structure

- [Cover Pages and Table of Contents](https://onizet.github.io/html2openxml/guide/document-layout.md)

### Technical Deep Dive

- [List Numbering](https://onizet.github.io/html2openxml/guide/numbering.md)
- [Advanced Tables (RowSpan, ColSpan, ColGroup)](https://onizet.github.io/html2openxml/guide/tables.md)
- [Image Processing](https://onizet.github.io/html2openxml/guide/images.md)
- [CSS Stylesheet support](https://onizet.github.io/html2openxml/guide/styling.html#note-on-external-css)
- [Performance notes and memory allocations](https://onizet.github.io/html2openxml/appendix/performance.md)

## How to implement or debug features

Useful references for both OpenXML and HTML:

- [MDN](https://developer.mozilla.org/en-US/docs/Web/HTML)
- [W3Schools](https://www.w3schools.com/html/default.asp)
- [OpenXml MSDN](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing?view=openxml-3.0.1)

Open Microsoft Word and design your expected output. Save as a DOCX file, then rename as a ZIP. Extract the content and inspect those files:
`document.xml`, `numbering.xml` (for list) and `styles.xml`.

## Acknowledgements

Thank you to all contributors that share their bug fixes (in no particular order): scwebgroup, ddforge, daviderapicavoli, worstenbrood, jodybullen, BenBurns, OleK, scarhand, imagremlin, antgraf, mdeclercq, pauldbentley, xjpmauricio, jairoXXX, giorand, bostjanKlemenc, AaronLS, taishmanov, AyUsH18102001.
And thanks to David Podhola for the Nuget package.

Logo provided with the permission of [Enhanced Labs Design Studio](http://www.enhancedlabs.com).

## Support

This project is open source and I maintain it in my spare time.

Bug reports, feature requests, and pull requests are always welcome.
Please target the `dev` branch when submitting pull requests.

If you have questions, don't hesitate to get in touch.

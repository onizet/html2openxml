# HtmlToOpenXml&nbsp;[![Latest version](https://img.shields.io/nuget/v/HtmlToOpenXml.svg)](https://www.nuget.org/packages/HtmlToOpenXml/)&nbsp;[![Download Counts](https://img.shields.io/nuget/dt/HtmlToOpenXml.svg)](https://www.nuget.org/packages/HtmlToOpenXml/)&nbsp;[![MIT License](https://img.shields.io/badge/license-MIT-blue.svg)](https://github.com/onizet/html2openxml/blob/dev/LICENSE)

HtmlToOpenXml is a mature library for generating native OpenXML Word documents from HTML. It transforms simple or complete HTML into fully editable Word content while preserving document styles, numbering definitions, templates, and layout.

Originally created in 2009 to transform user-generated content into templated Word documents, it has evolved
into a mature HTML-to-OpenXML converter supporting styles, numbering, images, bookmarks, page layout, tables,
document template, and custom resource resolution.

![Logo](https://raw.githubusercontent.com/onizet/html2openxml/master/docs/images/html2openxml.png)

## Features

- HTML → DOCX conversion
- Word template integration
- Portrait and landscape page sections
- Headers and footers
- Tables with RowSpan and ColSpan support
- Numbered and bulleted lists
- Images (PNG, JPEG, SVG, Data URI)
- Internal hyperlinks and bookmarks
- Right-to-left (RTL) documents
- CSS stylesheet support (optional package)

Supports **.NET Framework 4.6.2**, **.NET Standard 2.0**, **.NET 8** **.NET 10**.

Built on top of [DocumentFormat.OpenXml](https://www.nuget.org/packages/DocumentFormat.OpenXml/) and [AngleSharp](https://www.nuget.org/packages/AngleSharp).

| Package | NuGet Stable | Downloads |
| ------- | ------------ | --------- |
| [HtmlToOpenXml](https://www.nuget.org/packages/HtmlToOpenXml/) | [![HtmlToOpenXml](https://img.shields.io/nuget/v/HtmlToOpenXml.svg)](https://www.nuget.org/packages/HtmlToOpenXml/) | [![HtmlToOpenXml](https://img.shields.io/nuget/dt/HtmlToOpenXml.svg)](https://www.nuget.org/packages/HtmlToOpenXml/) |
| [HtmlToOpenXml.CssInliner](https://www.nuget.org/packages/HtmlToOpenXml.CssInliner/) | [![HtmlToOpenXml.CssInliner](https://img.shields.io/nuget/v/HtmlToOpenXml.CssInliner.svg)](https://www.nuget.org/packages/HtmlToOpenXml.CssInliner/) | [![HtmlToOpenXml.CssInliner](https://img.shields.io/nuget/dt/HtmlToOpenXml.CssInliner.svg)](https://www.nuget.org/packages/HtmlToOpenXml.CssInliner/) |

> **Note**
> The package was renamed from `HtmlToOpenXml.dll` to `HtmlToOpenXml`.
> NuGet download statistics were unfortunately not transferred to the new package.

## Quick Start

When creating a blank document:

```csharp
await using var generatedDocument = new MemoryStream();
using var package = WordprocessingDocument.Create(generatedDocument, WordprocessingDocumentType.Document);

HtmlConverter converter = new(package);
await converter.ParseBody(html);
```

When inserting inside an existing document:

```csharp
await using var generatedDocument = new MemoryStream();
await templateFileStream.CopyToAsync(generatedDocument);
generatedDocument.Position = 0L;
using var package = WordprocessingDocument.Open(generatedDocument, true);
HtmlConverter converter = new(package);
await converter.ParseBody(html);
```

## Documentation

[**Documentation**](https://onizet.github.io/html2openxml/) covers from the minimal code to advanced scenarios.

### Guides & Tutorials

* [Getting Started Guide](https://onizet.github.io/html2openxml/guide/quickstart.md)
* [LLM Guidance](https://onizet.github.io/html2openxml/llmdoc.md)

### API Reference

* [API Contract](https://onizet.github.io/html2openxml/guide/api.md)

### Technical Deep Dive

* [Numbering List](https://onizet.github.io/html2openxml/guide/numbering.md)
* [Advanced Table Structures (Row/ColSpan, ColGroup)](https://onizet.github.io/html2openxml/guide/tables.md)
* [Image Processing](https://onizet.github.io/html2openxml/guide/images.md)
* [CSS Stylesheet support](https://onizet.github.io/html2openxml/guide/styling.html#note-on-external-css)
* [Performance notes and memory allocations](https://onizet.github.io/html2openxml/appendix/performance.md)

## How to implement or debug features

My reference bibles cover both OpenXml and HTML:

* [MDN](https://developer.mozilla.org/en-US/docs/Web/HTML)
* [W3Schools](https://www.w3schools.com/html/default.asp)
* [OpenXml MSDN](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing?view=openxml-3.0.1)

Open MS Word or Apple Pages and design your expected output. Save as a DOCX file, then rename as a ZIP. Extract the content and inspect those files:
`document.xml`, `numbering.xml` (for list) and `styles.xml`.

## Acknowledgements

Thank you to all contributors that share their bug fixes (in no particular order): scwebgroup, ddforge, daviderapicavoli, worstenbrood, jodybullen, BenBurns, OleK, scarhand, imagremlin, antgraf, mdeclercq, pauldbentley, xjpmauricio, jairoXXX, giorand, bostjanKlemenc, AaronLS, taishmanov, AyUsH18102001.
And thanks to David Podhola for the Nuget package.

Logo provided with the permission of [Enhanced Labs Design Studio](http://www.enhancedlabs.com).

## Support

This project is open source and I do my best to support it in my spare time. I'm always happy to receive Pull Request and grateful for the time you have taken. Please target branch `dev` only.
If you have questions, don't hesitate to get in touch with me!

![Latest version](https://img.shields.io/nuget/v/HtmlToOpenXml.svg)
![Download Counts](https://img.shields.io/nuget/dt/HtmlToOpenXml.svg)
[![MIT License](https://img.shields.io/badge/license-MIT-blue.svg)](https://github.com/onizet/html2openxml/blob/dev/LICENSE)

# What is HtmlToOpenXml?

HtmlToOpenXml is a mature, high-fidelity converter that translates complex HTML structures into native OpenXML elements, ensuring your content renders exactly as intended across all Microsoft Word users.

Originally created in 2009 to transform user-generated content into templated Word documents, it has evolved
into a mature HTML-to-OpenXML converter supporting styles, numbering, images, bookmarks, page layout, tables,
document template, and custom resource resolution.

Supports **.Net Framework 4.6.2**, **.NET Standard 2.0**, **.NET 8** **.NET 10** which are all LTS.

Built on top of [DocumentFormat.OpenXml](https://www.nuget.org/packages/DocumentFormat.OpenXml/) and [AngleSharp](https://www.nuget.org/packages/AngleSharp).

-> [Official Nuget Package](https://www.nuget.org/packages/HtmlToOpenXml) (❗ Update your bookmark, suffix `.dll` removed)

## Quick Start

When creating a blank document:

```csharp
await using var generatedDocument = new MemoryStream();
using var package = WordprocessingDocument.Create(generatedDocument, WordprocessingDocumentType.Document);

var mainPart = package.AddMainDocumentPart();
new Document(new Body()).Save(mainPart);
HtmlConverter converter = new(mainPart);
await converter.ParseBody(html);
```

When inserting inside an existing document:

```csharp
await using var generatedDocument = new MemoryStream();
await templateFileStream.CopyToAsync(generatedDocument);
using var package = WordprocessingDocument.Open(generatedDocument, true);
MainDocumentPart? mainPart = package.MainDocumentPart;
if (mainPart == null)
{
    mainPart = package.AddMainDocumentPart();
    new Document(new Body()).Save(mainPart);
}
HtmlConverter converter = new(mainPart);
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

* [Numbering List](https://onizet.github.io/html2openxml/numbering.md)
* [Advanced Table Structures (Row/ColSpan, ColGroup)](https://onizet.github.io/html2openxml/guide/tables.md)
* [Image Processing](https://onizet.github.io/html2openxml/guide/images.md)
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

![Latest version](https://img.shields.io/nuget/v/HtmlToOpenXml.dll.svg)
![Download Counts](https://img.shields.io/nuget/dt/HtmlToOpenXml.dll.svg)
[![MIT License](https://img.shields.io/badge/license-MIT-blue.svg)](https://github.com/onizet/html2openxml/blob/dev/LICENSE)

# What is HtmlToOpenXml?

HtmlToOpenXml is a small .Net library that convert simple or advanced HTML to plain OpenXml components. This program has started in 2009, initially to convert user's comments into Word.

This library supports both **.Net Framework 4.6.2**, **.NET Standard 2.0** and **.NET 8** which are all LTS.

Depends on [DocumentFormat.OpenXml](https://www.nuget.org/packages/DocumentFormat.OpenXml/) and [AngleSharp](https://www.nuget.org/packages/AngleSharp).

-> [Official Nuget Package](https://www.nuget.org/packages/HtmlToOpenXml.dll)

## See Also

* [Documentation](https://github.com/onizet/html2openxml/wiki)
* [How to deliver a generated DOCX from server Asp.Net/SharePoint?](https://github.com/onizet/html2openxml/wiki/Serves-a-generated-docx-from-the-server)
* [Prevent Document Edition](https://github.com/onizet/html2openxml/wiki/Prevent-Document-Edition)
* [Convert dotx to docx](https://github.com/onizet/html2openxml/wiki/Convert-.dotx-to-.docx)

## Supported Html tags

Refer to [w3schools’ tag](http://www.w3schools.com/tags/default.asp) list to see their meaning

* `a`
* `h1-h6`
* `abbr` and `acronym`
* `b`, `i`, `u`, `s`, `del`, `ins`, `em`, `strike`, `strong`
* `br` and `hr`
* `img`, `figcaption` and `svg`
* `table`, `td`, `tr`, `th`, `tbody`, `thead`, `tfoot`, `caption` and `col`
* `cite`
* `div`, `span`, `time`, `font` and `p`
* `pre`
* `sub` and `sup`
* `ul`, `ol` and `li`
* `dd` and `dt`
* `q`, `blockquote`, `dfn`
* `article`, `aside`, `section` are considered like `div`

Javascript (`script`), CSS `style`, `meta`, comments, buttons and input controls are ignored.
Other tags are treated like `div`.

In v1 and v2, Javascript (`script`), CSS `style`, `meta`, comments and other not supported tags does not generate an error but are **ignored**.

## Html Parser

In v3, the parsing of the Html relies on AngleSharp package, which follows the W3C specifications and actively supports Html5.

In v1 and v2, the parsing of the Html was done using a custom Regex-based enumerator and was more flexible, but leaving a complex code, hard to maintain.

## How to implement or debug features

My reference bibles cover both OpenXml and HTML:

* [MDN](https://developer.mozilla.org/en-US/docs/Web/HTML)
* [W3Schools](https://www.w3schools.com/html/default.asp)
* [OpenXml MSDN](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing?view=openxml-3.0.1)

Open MS Word or Apple Pages and design your expected output. Save as a DOCX file, then rename as a ZIP. Extract the content and inspect those files:
`document.xml`, `numbering.xml` (for list) and `styles.xml`.

## Acknowledgements

Thank you to all contributors that share their bug fixes (in no particular order): scwebgroup, ddforge, daviderapicavoli, worstenbrood, jodybullen, BenBurns, OleK, scarhand, imagremlin, antgraf, mdeclercq, pauldbentley, xjpmauricio, jairoXXX, giorand, bostjanKlemenc, AaronLS, taishmanov.
And thanks to David Podhola for the Nuget package.

Logo provided with the permission of [Enhanced Labs Design Studio](http://www.enhancedlabs.com).

## Support

This project is open source and I do my best to support it in my spare time. I'm always happy to receive Pull Request and grateful for the time you have taken. Please target branch `dev` only.
If you have questions, don't hesitate to get in touch with me!

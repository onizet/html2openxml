[![Latest version](https://img.shields.io/nuget/v/HtmlToOpenXml.dll.svg)](https://www.nuget.org/packages/HtmlToOpenXml.dll)
[![License MIT](https://img.shields.io/badge/license-MIT-green.svg)](https://github.com/onizet/html2openxml/blob/master/LICENSE)
[![NuGet](https://img.shields.io/nuget/dt/HtmlToOpenXml.dll.svg)]()
[![codecov.io](https://codecov.io/github/onizet/html2openxml/coverage.svg?branch=master)](https://codecov.io/github/onizet/html2openxml?branch=master)

[![netstandard1.3](https://img.shields.io/badge/netstandard-1.3-brightgreen.svg)](https://img.shields.io/badge/netstandard-1.3-brightgreen.svg)
[![netstandard2.0](https://img.shields.io/badge/netstandard-2.0-brightgreen.svg)](https://img.shields.io/badge/netstandard-2.0-brightgreen.svg)
[![net46](https://img.shields.io/badge/net-46-brightgreen.svg)](https://img.shields.io/badge/net-46-brightgreen.svg)

# What is Html2OpenXml?

Html2OpenXml is a small .Net library that convert simple or advanced HTML to plain OpenXml components. This program has started in 2009, initially to convert user's comments from SharePoint to Word.

This library supports both **.Net Core (netstandard1.3)** and **.Net Framework 4.6**.

Depends on [DocumentFormat.OpenXml](https://www.nuget.org/packages/DocumentFormat.OpenXml/).

## See Also

* [Documentation](https://github.com/onizet/html2openxml/wiki)
* [How to deliver a generated DOCX from server Asp.Net/SharePoint?](https://github.com/onizet/html2openxml/wiki/Serves-a-generated-docx-from-the-server)
* [Prevent Document Edition](https://github.com/onizet/html2openxml/wiki/Prevent-Document-Edition)
* [Convert dotx to docx](https://github.com/onizet/html2openxml/wiki/Convert-.dotx-to-.docx)

## Supported Html tags

Refer to [w3schools’ tag](http://www.w3schools.com/tags/default.asp) list to see their meaning

* &lt;a&gt;
* &lt;h1-h6&gt;
* &lt;abbr&gt; and &lt;acronym&gt;
* &lt;b&gt;, &lt;i&gt;, &lt;u&gt;, &lt;s&gt;, &lt;del&gt;, &lt;ins&gt;, &lt;em&gt;, &lt;strike&gt;, &lt;strong&gt;
* &lt;br&gt; and &lt;hr&gt;
* &lt;img&gt;, &lt;figcaption&gt;
* &lt;table&gt;, &lt;td&gt;, &lt;tr&gt;, &lt;th&gt;, &lt;tbody&gt;, &lt;thead&gt;, &lt;tfoot&gt; and &lt;caption&gt;
* &lt;cite&gt;
* &lt;div&gt;, &lt;span&gt;, &lt;font&gt; and &lt;p&gt;
* &lt;pre&gt;
* &lt;sub&gt; and &lt;sup&gt;
* &lt;ul&gt;, &lt;ol&gt; and &lt;li&gt;
* &lt;dd&gt; and &lt;dt&gt;
* &lt;q&gt; and &lt;blockquote&gt; (since 1.5)
* &lt;article&gt;, &lt;aside&gt;, &lt;section&gt; are considered like &lt;div&gt;

Javascript (&lt;script&gt;), CSS &lt;style&gt;, &lt;meta&gt; and other not supported tags does not generate an error but are **ignored**.

## Tolerance for bad formed HTML

The parsing of the Html is done using a custom Regex-based enumerator. These are supported:

|   | samples |
| ------------- | ------------- |
| Ignore case | &lt;span&gt;Some text&lt;SPAN&gt; |
| Missing closing tag or invalid tag position | &lt;i&gt;Here&lt;b&gt; is &lt;/i&gt; some&lt;/b&gt; bad formed html. |
| no need to be XHTML compliant | Both &lt;br&gt; and &lt;br/&gt; are valid |
| Color | red, #ff0000, #f00, rgb(255,0,0,.5), hsl(0, 100%, 50%) are all the red color |
| Attributes | &lt;table id=table1&gt; or &lt;table id="table1"&gt; |

## Acknowledgements

Thank you to all contributors that share their bug fixes: scwebgroup, ddforge, daviderapicavoli, worstenbrood, jodybullen, BenBurns, OleK, scarhand, imagremlin, antgraf, mdeclercq, pauldbentley, xjpmauricio, jairoXXX, giorand, bostjanKlemenc, AaronLS, taishmanov.
And thanks to David Podhola for the Nuget package.

Logo provided with the permission of [Enhanced Labs Design Studio](http://www.enhancedlabs.com).

## Support

This project is open source and I do my best to support it in my spare time. I'm always happy to receive Pull Request and grateful for the time you have taken
If you have questions, don't hesitate to get in touch with me!

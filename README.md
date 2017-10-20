[![Latest version](https://img.shields.io/nuget/v/HtmlToOpenXml.dll.svg)](https://www.nuget.org/packages/HtmlToOpenXml.dll)
[![License MSPL](https://img.shields.io/badge/license-MSPL-green.svg)](https://github.com/onizet/html2openxml/blob/master/License.txtl)
[![NuGet](https://img.shields.io/nuget/dt/HtmlToOpenXml.dll.svg)]()

## What is Html2OpenXml?

Httml2OpenXml is a small .Net library that convert simple or advanced HTML to plain OpenXml components. This program has started in 2009, initially to convert user's comments from SharePoint to Word.

This library supports both **.NetCore (netstandard1.4)** and **.Net Framework 3.5, 4.0 and 4.6**.

Depends on either the [DocumentFormat.OpenXml](https://www.nuget.org/packages/DocumentFormat.OpenXml/).

### See Also

* [Documentation](https://github.com/onizet/html2openxml/wiki)
* [How to deliver a generated DOCX from server Asp.Net/SharePoint?](https://github.com/onizet/html2openxml/wiki/Serves-a-generated-docx-from-the-server)
* [Prevent Document Edition](https://github.com/onizet/html2openxml/wiki/Prevent-Document-Edition)
* [Convert dotx to docx](https://github.com/onizet/html2openxml/wiki/Convert-.dotx-to-.docx)

### Supported Html tags
Refer to [w3schools’ tag](http://www.w3schools.com/tags/default.asp) list to see their meaning
*	&lt;a&gt;
*	&lt;h1-h6&gt;
*	&lt;abbr&gt; and &lt;acronym&gt;
*	&lt;b&gt;, &lt;i&gt;, &lt;u&gt;, &lt;s&gt;, &lt;del&gt;, &lt;ins&gt;, &lt;em&gt;, &lt;strike&gt;, &lt;strong&gt;
*	&lt;br&gt; and &lt;hr&gt;
*	&lt;img&gt;, &lt;figcaption&gt;
*	&lt;table&gt;, &lt;td&gt;, &lt;tr&gt;, &lt;th&gt;, &lt;tbody&gt;, &lt;thead&gt;, &lt;tfoot&gt; and &lt;caption&gt;
*	&lt;cite&gt;
*	&lt;div&gt;, &lt;span&gt;, &lt;font&gt; and &lt;p&gt;
*	&lt;pre&gt;
*	&lt;sub&gt; and &lt;sup&gt;
*	&lt;ul&gt;, &lt;ol&gt; and &lt;li&gt;
*	&lt;dd&gt; and &lt;dt&gt;
* &lt;q&gt; and &lt;blockquote&gt; (since 1.5)

Javascript (&lt;script&gt;), CSS &lt;style&gt;, &lt;meta&gt; and other not supported tags does not generate an error but are **ignored**.

### Tolerance for bad formed HTML
The parsing of the Html is done using a custom Regex-based enumerator. These are supported:

<table>
<tr><th></th><th>samples</th></tr>
<tr>
  <td>Ignore case</td>
  <td>&lt;span&gt;Some text&lt;SPAN&gt;</td>
</tr>
<tr>
  <td>Missing closing tag or invalid tag position</td>
  <td>&lt;i&gt;Here&lt;b&gt; is &lt;/i&gt; some&lt;/b&gt; bad formed html.</td>
</tr>
<tr>
  <td>no need to be XHTML compliant</td>
  <td>Both &lt;br&gt; and &lt;br/&gt; are valid</td>
</tr>
<tr>
  <td>Color</td>
  <td>red, #ff0000, ff0000, rgb(255,0,0) are all the red color</td>
</tr>
<tr>
  <td>Attributes</td>
  <td>&lt;table id=table1&gt; or &lt;table id="table1"&gt;</td>
</tr>
</table>

### Acknowledgements

Thank you to all contributors that share their bug fixes: scwebgroup, ddforge, daviderapicavoli, worstenbrood, jodybullen, BenBurns, OleK, scarhand, imagremlin, antgraf, mdeclercq, pauldbentley, xjpmauricio, jairoXXX, giorand, bostjanKlemenc, AaronLS.
And thanks to David Podhola for the Nuget package.

Logo provided with the permission of [Enhanced Labs Design Studio](http://www.enhancedlabs.com).

### Support

This project is open source and I do my best to support it in my spare time. I'm always happy to receive Pull Request and grateful for the time you have taken
If you have questions, don't hesitate to get in touch with me!

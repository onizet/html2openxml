/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml - All Rights Reserved
 * 
 * This source is subject to the Microsoft Permissive License.
 * Please see the License.txt file for more information.
 * All other rights reserved.
 * 
 * THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY 
 * KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
 * PARTICULAR PURPOSE.
 */
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace HtmlToOpenXml
{
	/// <summary>
	/// Splits an html chunk of text and provide a way to enumerate through its tags.
	/// </summary>
	[System.Diagnostics.DebuggerDisplay("HtmlEnumerator. Current: {Current}")]
	sealed class HtmlEnumerator : IEnumerator<String>
	{
		private static Regex
            stripTagRegex = new Regex(@"(</?\w+)");          // extract the name of a tag without its attributes but with the < >

		private IEnumerator<String> en;
		private String current, currentTag;
		private HtmlAttributeCollection attributes, styleAttributes;


		/// <summary>
		/// Constructor.
		/// </summary>
		public HtmlEnumerator(String html)
		{
			// Clean a bit the html before processing

			// Remove Script tags, doctype, comments, css style, controls and html head part
			html = Regex.Replace(html, @"<xml.+?</xml>|<!--.+?-->|<script.+?</script>|<style.+?</style>|<head.+</head>|<!.+?>|<input.+?/>|<select.+?</select>|<textarea.+?</textarea>|<button.+?</button>", String.Empty,
								 RegexOptions.IgnoreCase | RegexOptions.Singleline);

			// Removes tabs and whitespace inside and before|next the line-breaking tags (p, div, br and body)
            // to preserve first whitespaces on the beginning of a 'pre' tag, we use '\bp\b' tag to exclude matching <pre> (by giorand, bug #13800)
            html = Regex.Replace(html, @"(\s*)(</?(\bp\b|div|br|body)[^>]*/?>)(\s*)", "$2", RegexOptions.Multiline| RegexOptions.IgnoreCase);

			// Preserves whitespaces inside Pre tags.
			html = Regex.Replace(html, "(<pre.*?>)(.+?)</pre>", PreserveWhitespacesInPre, RegexOptions.Singleline| RegexOptions.IgnoreCase);

			// Remove tabs and whitespace at the beginning of the lines
			html = Regex.Replace(html, @"^\s+", String.Empty, RegexOptions.Multiline);
			// and now at the end of the lines
			html = Regex.Replace(html, @"\s+$", String.Empty, RegexOptions.Multiline);

			// Replace xml header by xml tag for further processing
			html = Regex.Replace(html, @"<\?xml:namespace.+?>", "<xml>", RegexOptions.Singleline| RegexOptions.IgnoreCase);

			// Ensure order of table elements are respected: thead, tbody and tfooter
			// we select only the table that contains at least a tfoot or thead tag
			//html = Regex.Replace(html, @"<table.*?>(\s+</?(?=(thead|tbody|tfoot))).+?</\2>\s+</table>", PreserveTablePartOrder, RegexOptions.Singleline);
			html = Regex.Replace(html, "(<table.*?>)(.*?)(</table>)", PreserveTablePartOrder, RegexOptions.Singleline | RegexOptions.IgnoreCase);

			// Split our html using the tags
			String[] lines = Regex.Split(html, @"(</?\w+[^>]*/?>)", RegexOptions.Singleline);

			this.en = (lines as IEnumerable<String>).GetEnumerator();
		}

		public void Dispose()
		{
			en.Dispose();
		}

		//__________________________________________________________________________
		//
		// Private Implementation

		#region PreserveWhitespacesInPre

		private static String PreserveWhitespacesInPre(Match match)
		{
			// Convert new lines in <pre> to <br> tags for easier processing
			string innerHtml = Regex.Replace(match.Groups[2].Value, "\r?\n", "<br>");
			// Remove any whitespace at the end of the pre
			innerHtml = Regex.Replace(innerHtml, @"(<br>|\s+)$", String.Empty);
			return match.Groups[1].Value + innerHtml + "</pre>";
		}

		#endregion

		#region PreserveTablePartOrder

		private static String PreserveTablePartOrder(Match match)
		{
			// ensure the order of the table elements are set in the correct order.
			// bug #11016 reported by pauldbentley

			var sb = new System.Text.StringBuilder();
			sb.Append(match.Groups[1].Value);

			Regex tableSplitReg = new Regex(@"(<(?=(caption|colgroup|thead|tbody|tfoot|tr)).*?>.+?</\2>)", RegexOptions.Singleline | RegexOptions.IgnoreCase);
			MatchCollection matches = tableSplitReg.Matches(match.Groups[2].Value);

			foreach (String tagOrder in new [] { "caption", "colgroup", "thead", "tbody", "tfoot", "tr" })
			foreach (Match m in matches)
			{
				if (m.Groups[2].Value.Equals(tagOrder, StringComparison.OrdinalIgnoreCase))
					sb.Append(m.Groups[1].Value);
			}

			sb.Append(match.Groups[3].Value);
			return sb.ToString();
		}

		#endregion

		//__________________________________________________________________________
		//
		// Public Functionality

		public void Reset()
		{
			en.Reset();
		}

		/// <summary>
		/// Use as MoveNext() but this function will stop once the current value is equals to tag.
		/// </summary>
		/// <param name="tag">The tag to stop on (Optional).</param>
		/// <returns>
		/// If tag is null, it returns true if the enumerator was successfully advanced to the next element; false
		/// if the enumerator has passed the end of the collection.<br/>
		/// If tag is not null, it returns false as long as the tag was not found.
		/// </returns>
		public bool MoveUntilMatch(String tag)
		{
			current = currentTag = null;
			attributes = styleAttributes = null;
			bool success;

			// Ignore empty lines
			while ((success = en.MoveNext()) && (current = en.Current.Trim('\n', '\r')).Length == 0) ;

			if (success && tag != null)
				return !current.Equals(tag, StringComparison.CurrentCultureIgnoreCase);

			return success;
		}

		public bool MoveNext()
		{
			return MoveUntilMatch(null);
		}

		/// <summary>
		/// Gets an attribute in the Style attribute of a Html tag.
		/// </summary>
		public HtmlAttributeCollection StyleAttributes
		{
			get { return styleAttributes ?? (styleAttributes = HtmlAttributeCollection.ParseStyle(this.Attributes["style"])); }
		}

		/// <summary>
		/// Gets an attribute from a Html tag.
		/// </summary>
		public HtmlAttributeCollection Attributes
		{
			get { return attributes ?? (attributes = HtmlAttributeCollection.Parse(current)); }
		}

		/// <summary>
		/// Gets whether the current element is an Html tag or not.
		/// </summary>
		public bool IsCurrentHtmlTag
		{
			get { return current[0] == '<'; }
		}

		/// <summary>
		/// Gets whether the current element is an Html tag that is closed (example: &lt;td/&gt;).
		/// </summary>
		public bool IsSelfClosedTag
		{
			get { return this.IsCurrentHtmlTag && current.EndsWith("/>", StringComparison.Ordinal); }
		}

		/// <summary>
		/// If <see cref="HtmlEnumerator.Current"/> property is a Html tag, it returns the name of that tag.
		/// </summary>
		public String CurrentTag
		{
			get
			{
				if(currentTag == null)
				{
					Match m = stripTagRegex.Match(current);
					currentTag = m.Success ? m.Groups[1].Value + ">" : null;
				}
				return currentTag;
			}
		}

		/// <summary>
		/// Gets the expected closing tag for the current tag.
		/// </summary>
		public String ClosingCurrentTag
		{
			get
			{
				if (IsSelfClosedTag) return this.CurrentTag;
				return this.CurrentTag.Insert(1, "/");
			}
		}

		/// <summary>
		/// Gets the line or tag at the current position of the enumerator.
		/// </summary>
		public String Current
		{
			get { return current; }
		}

		Object System.Collections.IEnumerator.Current
		{
			get { return current; }
		}
	}
}
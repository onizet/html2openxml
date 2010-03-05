using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Globalization;

namespace NotesFor.HtmlToOpenXml
{
	/// <summary>
	/// Split an html piece of text and provide a way to enumerate its tags.
	/// </summary>
	[System.Diagnostics.DebuggerDisplay("HtmlEnumerator. Current: {Current}")]
	sealed class HtmlEnumerator : IEnumerator<String>
	{
		private static Regex
			stripTagRegex,          // extract the name of a tag without its attributes but with the < >
			beginOfLineTrimRegex;   // remove whitespaces at the beginning of any new lines.

		private IEnumerator<String> en;
		private String current, currentTag;
		private HtmlAttributeCollection attributes, styleAttributes;

		static HtmlEnumerator()
		{
			stripTagRegex = new Regex(@"(</?\w+)", RegexOptions.Compiled);
			beginOfLineTrimRegex = new Regex(@"\r?\n\s*([^<])", RegexOptions.Compiled);
		}

		/// <summary>
		/// Constructor.
		/// </summary>
		public HtmlEnumerator(String html)
		{
			// Clean a bit the html before processing

			// Remove Script tags, doctype, comments, css style, controls and html head part
			html = Regex.Replace(html, @"<!--.+?-->|<script.+?</script>|<style.+?</style>|<head.+</head>|<!.+?>|<input.+?/>|<select.+?</select>|<textarea.+?</textarea>|<button.+?</button>", String.Empty,
								 RegexOptions.IgnoreCase | RegexOptions.Singleline);

			// Preserves whitespaces inside Pre tags.
			html = Regex.Replace(html, "(<pre.*?>)(.+?)</pre>", PreserveWhitespacesInPre, RegexOptions.Singleline);

			// Remove tabs and whitespace at the beginning of the lines
			html = Regex.Replace(html, @"^\s+", String.Empty, RegexOptions.Multiline);

			// Split our html using the tags
			String[] lines = Regex.Split(html, @"(</?\w+[^>]*/?>)", RegexOptions.Singleline);

			this.en = (lines as IEnumerable<String>).GetEnumerator();
		}

		public void Dispose()
		{
			en.Dispose();
		}

		public void Reset()
		{
			en.Reset();
		}

		private static String PreserveWhitespacesInPre(Match match)
		{
			// Convert new lines in <pre> to <br> tags for easier processing
			string innerHtml = Regex.Replace(match.Groups[2].Value, "\r?\n", "<br>");
			// Remove any whitespace at the beginning or end of the pre
			innerHtml = Regex.Replace(innerHtml, "^<br>|<br>$", String.Empty);
			return match.Groups[1].Value + innerHtml + "</pre>";
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

			// Ignore empty lines and remove tabs and whitespace at the beginning of the lines
			// (unless the PreserveWhitespaces property is enabled)
			if (this.PreserveWhitespaces)
			{
				while ((success = en.MoveNext()) && en.Current.Length == 0) ;
				if (success) current = en.Current;
			}
			else
			{
				while ((success = en.MoveNext()) && en.Current.Trim().Length == 0) ;

				// Remove tabs and whitespace at the beginning of the lines
				if(success) current = beginOfLineTrimRegex.Replace(en.Current, " $1");
			}

			if (success && tag != null)
				return !current.Equals(tag, StringComparison.CurrentCultureIgnoreCase);

			return success;
		}

		bool System.Collections.IEnumerator.MoveNext()
		{
			return MoveUntilMatch(null);
		}

		/// <summary>
		/// Gets an attribute in the Style attribute of a Html tag.
		/// </summary>
		public HtmlAttributeCollection StyleAttributes
		{
			get { return styleAttributes ?? (styleAttributes = new HtmlAttributeCollection(this.Attributes["style"], true)); }
		}

		/// <summary>
		/// Gets an attribute from a Html tag.
		/// </summary>
		public HtmlAttributeCollection Attributes
		{
			get { return attributes ?? (attributes = new HtmlAttributeCollection(current, false)); }
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
			get { return this.IsCurrentHtmlTag && current.EndsWith("/>"); }
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
		/// Gets or sets whether the enumerator should preserve white spaces at the beginning of the lines.
		/// </summary>
		public bool PreserveWhitespaces { get; set; }

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
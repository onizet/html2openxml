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
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml
{
	/// <summary>
	/// Defines the styles to apply on OpenXml elements.
	/// </summary>
	public sealed class HtmlDocumentStyle
	{
		/// <summary>
		/// Occurs when a Style is missing in the MainDocumentPart but will be used during the conversion process.
		/// </summary>
		public event EventHandler<StyleEventArgs> StyleMissing;

		/// <summary>
		/// Contains the default styles for new OpenXML elements
		/// </summary>
		public DefaultStyles DefaultStyles { get { return this._defaultStyles; } }

		private DefaultStyles _defaultStyles = new DefaultStyles();
		private RunStyleCollection _runStyle;
		private TableStyleCollection _tableStyle;
		private ParagraphStyleCollection _paraStyle;
        private NumberingListStyleCollection _listStyle;
		private OpenXmlDocumentStyleCollection _knownStyles;
		private MainDocumentPart _mainPart;


		internal HtmlDocumentStyle(MainDocumentPart mainPart)
		{
			PrepareStyles(mainPart);
			_tableStyle = new TableStyleCollection(this);
			_runStyle = new RunStyleCollection(this);
			_paraStyle = new ParagraphStyleCollection(this);
            this.QuoteCharacters = QuoteChars.IE;
			this._mainPart = mainPart;
		}

		//____________________________________________________________________
		//

		#region PrepareStyles

		/// <summary>
		/// Preload the styles in the document to match localized style name.
		/// </summary>
		internal void PrepareStyles(MainDocumentPart mainPart)
		{
			_knownStyles = new OpenXmlDocumentStyleCollection();
			if (mainPart.StyleDefinitionsPart == null) return;

			Styles styles = mainPart.StyleDefinitionsPart.Styles;

			foreach (var s in styles.Elements<Style>())
			{
				StyleName n = s.StyleName;
				if (n != null)
				{
					String name = n.Val.Value;
					if (name != s.StyleId) _knownStyles[name] = s;
				}

				_knownStyles.Add(s.StyleId, s);
			}
		}

		#endregion

		#region GetStyle

		/// <summary>
		/// Helper method to obtain the StyleId of a named style (invariant or localized name).
		/// </summary>
		/// <param name="name">The name of the style to look for.</param>
		/// <param name="styleType">True to obtain the character version of the given style.</param>
		/// <param name="ignoreCase">Indicate whether the search should be performed with the case-sensitive flag or not.</param>
		/// <returns>If not found, returns the given name argument.</returns>
		public String GetStyle(string name, StyleValues styleType = StyleValues.Paragraph, bool ignoreCase = false)
		{
			Style style;

			// OpenXml is case-sensitive but CSS is not.
			// We will try to find the styles another time with case-insensitive:
			if (ignoreCase)
			{
				if (!_knownStyles.TryGetValueIgnoreCase(name, styleType, out style))
				{
					if (StyleMissing != null)
					{
						StyleMissing(this, new StyleEventArgs(name, _mainPart, styleType));
						if (_knownStyles.TryGetValueIgnoreCase(name, styleType, out style))
							return style.StyleId;
					}
					return null; // null means we ignore this style (css class)
				}

				return style.StyleId;
			}
			else
			{
				if (!_knownStyles.TryGetValue(name, out style))
				{
					if (!EnsureKnownStyle(name, out style))
					{
						StyleMissing?.Invoke(this, new StyleEventArgs(name, _mainPart, styleType));
						return name;
					}
				}

				if (styleType == StyleValues.Character && !style.Type.Equals<StyleValues>(StyleValues.Character))
				{
					LinkedStyle linkStyle = style.GetFirstChild<LinkedStyle>();
					if (linkStyle != null) return linkStyle.Val;
				}
				return style.StyleId;
			}
		}

		#endregion

		#region DoesStyleExists

		/// <summary>
		/// Gets whether the given style exists in the document.
		/// </summary>
		public bool DoesStyleExists(string name)
		{
			return _knownStyles.ContainsKey(name);
		}

		#endregion

		#region AddStyle

		/// <summary>
		/// Add a new style inside the document and refresh the style cache.
		/// </summary>
		internal void AddStyle(String name, Style style)
		{
			_knownStyles[name] = style;
			if (_mainPart.StyleDefinitionsPart == null)
				_mainPart.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
			_mainPart.StyleDefinitionsPart.Styles.Append(style);
		}

		#endregion

        #region EnsureKnownStyle

        /// <summary>
        /// Try to insert the style in the document if it is a known style.
        /// </summary>
        private bool EnsureKnownStyle(string styleName, out Style style)
        {
			style = null;
			string xml = PredefinedStyles.GetOuterXml(styleName);
			if (xml == null) return false;
			this.AddStyle(styleName, style = new Style(xml));
			return true;
        }

        #endregion

		//____________________________________________________________________
		//

		internal RunStyleCollection Runs
		{
			[System.Diagnostics.DebuggerHidden()]
			get { return _runStyle; }
		}
		internal TableStyleCollection Tables
		{
			[System.Diagnostics.DebuggerHidden()]
			get { return _tableStyle; }
		}
		internal ParagraphStyleCollection Paragraph
		{
			[System.Diagnostics.DebuggerHidden()]
			get { return _paraStyle; }
		}
        internal NumberingListStyleCollection NumberingList
        {
			// use lazy loading to avoid injecting NumberListDefinition if not required
            [System.Diagnostics.DebuggerHidden()]
            get { return _listStyle ?? (_listStyle = new NumberingListStyleCollection(_mainPart)); }
        }

		//____________________________________________________________________
		//

        /// <summary>
        /// Gets or sets the beginning and ending characters used in the &lt;q&gt; tag.
        /// </summary>
        public QuoteChars QuoteCharacters { get; set; }
	}
}
using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.Collections.Specialized;

namespace NotesFor.HtmlToOpenXml
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

		private RunStyleCollection runStyle;
		private TableStyleCollection tableStyle;
		private ParagraphStyleCollection paraStyle;
        private NumberingListStyleCollection listStyle;
		private Dictionary<String, Style> knownStyles;
		private MainDocumentPart mainPart;



		internal HtmlDocumentStyle(MainDocumentPart mainPart)
		{
			PrepareStyles(mainPart);
			runStyle = new RunStyleCollection();
			tableStyle = new TableStyleCollection();
			paraStyle = new ParagraphStyleCollection();
            listStyle = new NumberingListStyleCollection(mainPart);
            this.QuoteCharacters = QuoteChars.IE;
			this.mainPart = mainPart;
		}

		//____________________________________________________________________
		//

		#region PrepareStyles

		/// <summary>
		/// Preload the styles in the document to match localized style name.
		/// </summary>
		internal void PrepareStyles(MainDocumentPart mainPart)
		{
			knownStyles = new Dictionary<String, Style>();
			if (mainPart.StyleDefinitionsPart == null) return;

			Styles styles = mainPart.StyleDefinitionsPart.Styles;

			foreach (var s in styles.Elements<Style>())
			{
				StyleName n = s.GetFirstChild<StyleName>();
				if (n != null)
				{
					String name = n.Val.Value;
					if (name != s.StyleId) knownStyles[name] = s;
				}

				knownStyles.Add(s.StyleId, s);
			}
		}

		#endregion

		#region GetStyle

		/// <summary>
		/// Helper method to obtain the StyleId of a named style (invariant or localized name).
		/// </summary>
		/// <param name="name">The name of the style to look for.</param>
		/// <param name="characterType">True to obtain the character version of the given style.</param>
		/// <returns>If not found, returns the given name argument.</returns>
		public String GetStyle(string name, bool characterType = false)
		{
			Style style;
			if (!knownStyles.TryGetValue(name, out style))
			{
				if (StyleMissing != null) StyleMissing(this, new StyleEventArgs(name, mainPart));
				return name;
			}

			if (characterType && !style.Type.Equals<StyleValues>(StyleValues.Character))
			{
				LinkedStyle linkStyle = style.GetFirstChild<LinkedStyle>();
				if (linkStyle != null) return linkStyle.Val;
			}
			return style.StyleId;
		}

		#endregion

		#region DoesStyleExists

		/// <summary>
		/// Gets whether the given style exists in the document.
		/// </summary>
		public bool DoesStyleExists(string name)
		{
			return knownStyles.ContainsKey(name);
		}

		#endregion

		#region AddStyle

		/// <summary>
		/// Add a new style inside the document and refresh the style cache.
		/// </summary>
		internal void AddStyle(String name, Style style)
		{
			knownStyles[name] = style;
			if (mainPart.StyleDefinitionsPart == null)
				mainPart.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
			mainPart.StyleDefinitionsPart.Styles.Append(style);
		}

		#endregion

		//____________________________________________________________________
		//

		internal RunStyleCollection Runs
		{
			[System.Diagnostics.DebuggerHidden()]
			get { return runStyle; }
		}
		internal TableStyleCollection Tables
		{
			[System.Diagnostics.DebuggerHidden()]
			get { return tableStyle; }
		}
		internal ParagraphStyleCollection Paragraph
		{
			[System.Diagnostics.DebuggerHidden()]
			get { return paraStyle; }
		}
        internal NumberingListStyleCollection NumberingList
        {
            [System.Diagnostics.DebuggerHidden()]
            get { return listStyle; }
        }

		//____________________________________________________________________
		//

		/// <summary>
		/// Gets the default StyleId to apply on the any new paragraph.
		/// </summary>
		internal String DefaultParagraphStyle
		{
			get { return paraStyle.DefaultParagraphStyle; }
			set { paraStyle.DefaultParagraphStyle = value; }
		}

		/// <summary>
		/// Gets or sets the default paragraph style to apply on any new runs.
		/// </summary>
		public String DefaultStyle
		{
			get { return DefaultParagraphStyle ?? runStyle.DefaultRunStyle; }
			set
			{
				if (String.IsNullOrEmpty(value))
				{
					runStyle.DefaultRunStyle = null;
					this.DefaultParagraphStyle = null;
					return;
				}

				Style s;
				if (!knownStyles.TryGetValue(value, out s))
				{
					this.DefaultParagraphStyle = value;
				}
				else
				{
					if (s.Type.Equals<StyleValues>(StyleValues.Paragraph))
						this.DefaultParagraphStyle = s.StyleId;
					else
						runStyle.DefaultRunStyle = s.StyleId;
				}
			}
		}

        /// <summary>
        /// Gets or sets the beginning and ending characters used in the &lt;q&gt; tag.
        /// </summary>
        public QuoteChars QuoteCharacters { get; set; }
	}
}
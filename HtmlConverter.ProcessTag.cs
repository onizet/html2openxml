using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace NotesFor.HtmlToOpenXml
{
	using a = DocumentFormat.OpenXml.Drawing;
	using pic = DocumentFormat.OpenXml.Drawing.Pictures;

	partial class HtmlConverter
	{
		//____________________________________________________________________
		//
		// Processing known tags

		#region ProcessAcronym

		private void ProcessAcronym(HtmlEnumerator en)
		{
			// Transform the inline acronym/abbreviation to a reference to a foot note.

			string title = en.Attributes["title"];
			if (title == null) return;

			AlternateProcessHtmlChunks(en, en.CurrentTag.Replace("<", "</"));

			if (elements.Count > 0 && elements[0] is Run)
			{
				string defaultRefStyle, runStyle;
				FootnoteEndnoteReferenceType reference;

				if (this.AcronymPosition == AcronymPosition.PageEnd)
				{
					reference = new FootnoteReference() { Id = AddFootnoteReference(title) };
					defaultRefStyle = "footnote text";
					runStyle = "footnote reference";
				}
				else
				{
					reference = new EndnoteReference() { Id = AddEndnoteReference(title) };
					defaultRefStyle = "endnote text";
					runStyle = "endnote reference";
				}


				Run run;
				elements.Add(
					run = new Run(
						new RunProperties(
							new RunStyle() { Val = htmlStyles.GetStyle(runStyle, true) }),
						reference));

				if (!htmlStyles.DoesStyleExists(defaultRefStyle))
				{
					// Force the superscript style because if the footnote text style does not exists,
					// the rendering will be awful.
					run.InsertInProperties(new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript });
				}
			}
		}

		#endregion

		#region ProcessBold

		private void ProcessBold(HtmlEnumerator en)
		{
			htmlStyles.Runs.BeginTag("<b>", new Bold());
		}

		#endregion

		#region ProcessBody

		private void ProcessBody(HtmlEnumerator en)
		{
			List<OpenXmlElement> styleAttributes = new List<OpenXmlElement>();
			ProcessContainerAttributes(en, styleAttributes);

			if (styleAttributes.Count > 0)
				htmlStyles.Runs.BeginTag("<body>", styleAttributes.ToArray());
		}

		#endregion

		#region ProcessBr

		private void ProcessBr(HtmlEnumerator en)
		{
			elements.Add(new Run(new Break()));
		}

		#endregion

		#region ProcessCite

		private void ProcessCite(HtmlEnumerator en)
		{
			htmlStyles.Runs.BeginTag("<cite>", new RunStyle() { Val = htmlStyles.GetStyle("Quote", true) });
		}

		#endregion

		#region ProcessDefinitionList

		private void ProcessDefinitionList(HtmlEnumerator en)
		{
			ProcessParagraph(en);
			currentParagraph.InsertInProperties(
				 new SpacingBetweenLines() { After = "0" });
		}

		#endregion

		#region ProcessDefinitionListItem

		private void ProcessDefinitionListItem(HtmlEnumerator en)
		{
			AlternateProcessHtmlChunks(en, "</dd>");

			currentParagraph = htmlStyles.Paragraph.NewParagraph();
			currentParagraph.Append(elements);
			currentParagraph.InsertInProperties(
				   new Indentation() { FirstLine = "708" },
				   new SpacingBetweenLines() { After = "0" }
			);

			// Restore the original elements list
			AddParagraph(currentParagraph);
			this.elements.Clear();
		}

		#endregion

		#region ProcessDiv

		private void ProcessDiv(HtmlEnumerator en)
		{
			// The way the browser consider <div> is like a simple Break. But in case of any attributes that targets
			// the paragraph, we don't want to apply the style on the old paragraph but on a new one.
			if (en.Attributes.Count == 0 || (en.StyleAttributes["text-align"] == null && en.Attributes["align"] == null))
			{
				CompleteCurrentParagraph();
				Paragraph previousParagraph = currentParagraph;
				currentParagraph = htmlStyles.Paragraph.NewParagraph();

				List<OpenXmlElement> runStyleAttributes = new List<OpenXmlElement>();
				bool newParagraph = ProcessContainerAttributes(en, runStyleAttributes);

				if (runStyleAttributes.Count > 0)
					htmlStyles.Runs.BeginTag(en.CurrentTag, runStyleAttributes);

				// Any changes that requires a new paragraph?
				if (!newParagraph && previousParagraph.HasChild<Run>())
				{
					ProcessBr(en);
					currentParagraph = previousParagraph;
				}
				else
				{
					if (newParagraph)
					{
						// Insert before the break, complete this paragraph and start a new one
						this.paragraphs.Insert(this.paragraphs.Count-1, currentParagraph);
						AlternateProcessHtmlChunks(en, "</div>");
						CompleteCurrentParagraph();
						AddParagraph(currentParagraph = htmlStyles.Paragraph.NewParagraph());
					}
					else
					{
						AddParagraph(currentParagraph);
					}
				}
			}
			else
			{
				ProcessParagraph(en);
			}
		}

		#endregion

		#region ProcessFont

		private void ProcessFont(HtmlEnumerator en)
		{
			List<OpenXmlElement> styleAttributes = new List<OpenXmlElement>();
			ProcessContainerAttributes(en, styleAttributes);

			string attrValue = en.Attributes["size"];
			if (attrValue != null)
			{
				uint fontSize = ConverterUtility.ConvertToFontSize(attrValue);
				if (fontSize != 0L)
					styleAttributes.Add(new FontSize { Val = fontSize.ToString(CultureInfo.InvariantCulture) });
			}

			if (styleAttributes.Count > 0)
				htmlStyles.Runs.MergeTag("<font>", styleAttributes);
		}

		#endregion

		#region ProcessHeading

		private void ProcessHeading(HtmlEnumerator en)
		{
			char level = en.Current[2];

			AlternateProcessHtmlChunks(en, "</h" + level + ">");
			Paragraph p = new Paragraph(elements);
			p.InsertInProperties(
				new ParagraphStyleId() { Val = htmlStyles.GetStyle("heading " + level, false) });

			this.elements.Clear();
			AddParagraph(p);
			AddParagraph(currentParagraph = htmlStyles.Paragraph.NewParagraph());
		}

		#endregion

		#region ProcessHorizontalLine

		private void ProcessHorizontalLine(HtmlEnumerator en)
		{
			// Insert an horizontal line as it stands in many emails.

			CompleteCurrentParagraph();

			UInt32 hrSize = 4U;

			// If the previous paragraph contains a bottom border, we should toggle the size of this <hr> to 0U or 4U
			// or Word will display only the last border.
			// (see Remarks: http://msdn.microsoft.com/en-us/library/documentformat.openxml.wordprocessing.bottomborder%28office.14%29.aspx)
			if (paragraphs.Count > 1)
			{
				ParagraphProperties prop = paragraphs[paragraphs.Count - 2].GetFirstChild<ParagraphProperties>();
				if (prop != null)
				{
					ParagraphBorders borders = prop.GetFirstChild<ParagraphBorders>();
					if (borders != null && borders.HasChild<BottomBorder>())
					{
						if (borders.GetFirstChild<BottomBorder>().Size == 4U) hrSize = 0U;
						else hrSize = 4U;
					}
				}
			}


			currentParagraph.InsertInProperties(
				new ParagraphBorders(
					new BottomBorder() { Val = BorderValues.Single, Size = hrSize }));
			AddParagraph(currentParagraph = htmlStyles.Paragraph.NewParagraph());
		}

		#endregion

		#region ProcessImage

		private void ProcessImage(HtmlEnumerator en)
		{
			Drawing drawing = null;

			if (this.ImageProcessing != ImageProcessing.Ignore)
			{
				string src = en.Attributes["src"];
				Uri uri;

				if (src != null && Uri.TryCreate(src, UriKind.RelativeOrAbsolute, out uri))
				{
					string alt = en.Attributes["alt"];
					bool process = true;

					if (!uri.IsAbsoluteUri && this.BaseImageUrl != null)
						uri = new Uri(this.BaseImageUrl, uri);

					Size preferredSize = Size.Empty;
					if (en.Attributes["width"] != null || en.Attributes["height"] != null)
					{
						Unit wu = en.Attributes.GetAsUnit("width");
						Unit hu = en.Attributes.GetAsUnit("height");

						// only keep if the unit are px (otherwise it's dependent of the user's font size)
						if (wu.IsValid && wu.Value > 0 && wu.Type == "px")
							preferredSize.Width = wu.Value;
						if (hu.IsValid && hu.Value > 0 && hu.Type == "px")
							preferredSize.Height = hu.Value;
					}

					if (process)
						drawing = AddImagePart(uri, src, alt, preferredSize);
				}
			}

			if (drawing != null)
				elements.Add(new Run(drawing));
		}

		#endregion

		#region ProcessItalic

		private void ProcessItalic(HtmlEnumerator en)
		{
			htmlStyles.Runs.BeginTag("<i>", new Italic());
		}

		#endregion

		#region ProcessLi

		private void ProcessLi(HtmlEnumerator en)
		{
			CompleteCurrentParagraph();

			// Save the new paragraph reference to support nested numbering list.
			Paragraph p = htmlStyles.Paragraph.NewParagraph();
			currentParagraph = p;
			currentParagraph.InsertInProperties(
				new SpacingBetweenLines() { After = "0" },
				new NumberingProperties(
					new NumberingLevelReference() { Val = numberLevelRef - 1 },
					new NumberingId() { Val = numberingId }
				)
			);

			// Restore the original elements list
			AddParagraph(currentParagraph);

			// Continue to process the html until we found </li>
			AlternateProcessHtmlChunks(en, "</li>");
			p.Append(elements);
			this.elements.Clear();
		}

		#endregion

		#region ProcessLink

		private void ProcessLink(HtmlEnumerator en)
		{
			String att = en.Attributes["href"];
			Hyperlink h = null;
			Uri uri = null;


			if (!String.IsNullOrEmpty(att))
			{
				// is it an anchor?
				if (att[0] == '#' && att.Length > 1)
				{
					// Always accept _top anchor
					if (!this.ExcludeLinkAnchor || att == "#_top")
					{
						h = new Hyperlink(
							) { History = true, Anchor = att.Substring(1) };
					}
				}
				// ensure the links does not start with javascript:
				else if (Uri.TryCreate(att, UriKind.Absolute, out uri) && uri.Scheme != "javascript")
				{
					HyperlinkRelationship extLink = mainPart.AddHyperlinkRelationship(uri, true);

					h = new Hyperlink(
						) { History = true, Id = extLink.Id };
				}
			}

			if (h == null)
			{
				// link to a broken url, simply process the content of the tag
				ProcessHtmlChunks(en, "</a>");
				return;
			}

			AlternateProcessHtmlChunks(en, "</a>");

			if (elements.Count > 0)
			{
				// Let's see whether the link tag include an image inside its body.
				// If so, the Hyperlink OpenXmlElement is lost and we'll keep only the images
				// and applied a HyperlinkOnClick attribute.
				List<OpenXmlElement> imageInLink = elements.FindAll(e => { return e.HasChild<Drawing>(); });
				if (imageInLink.Count != 0)
				{
					for (int i = 0; i < imageInLink.Count; i++)
					{
						// Retrieves the "alt" attribute of the image and apply it as the link's tooltip
						Drawing d = imageInLink[i].GetFirstChild<Drawing>();
						var enDp = d.Descendants<pic.NonVisualDrawingProperties>().GetEnumerator();
						String alt;
						if (enDp.MoveNext()) alt = enDp.Current.Description;
						else alt = null;

						d.InsertInDocProperties(
							 new a.HyperlinkOnClick() { Id = h.Id ?? h.Anchor, Tooltip = alt });
					}

					CompleteCurrentParagraph();
					AddParagraph(currentParagraph = htmlStyles.Paragraph.NewParagraph());
				}
				else
				{
					// Append the processed elements and put them to the Run of the Hyperlink
					h.Append(elements);

					if (!htmlStyles.DoesStyleExists("Hyperlink"))
					{
						htmlStyles.AddStyle("Hyperlink", new Style(
							new StyleName() { Val = "Hyperlink" },
							new UnhideWhenUsed(),
							new StyleRunProperties(
								new DocumentFormat.OpenXml.Wordprocessing.Color() { Val = "0000FF", ThemeColor = ThemeColorValues.Hyperlink },
								new Underline() { Val = UnderlineValues.Single }
							)
						) { Type = StyleValues.Character, StyleId = "Hyperlink" });
					}

					h.GetFirstChild<Run>().InsertInProperties(
						new RunStyle() { Val = htmlStyles.GetStyle("Hyperlink", true) });

					this.elements.Clear();

					// Append the hyperlink
					elements.Add(h);
				}
			}
		}

		#endregion

		#region ProcessNumberingList

		private void ProcessNumberingList(HtmlEnumerator en)
		{
			EnsureNumberingIds();

			if (en.Current.Equals("<ul>", StringComparison.InvariantCultureIgnoreCase))
				numberingId = numberingUlId;
			else
				numberingId = numberingOlId;

			numberLevelRef++;
		}

		#endregion

		#region ProcessParagraph

		private void ProcessParagraph(HtmlEnumerator en)
		{
			CompleteCurrentParagraph();
			AddParagraph(currentParagraph = htmlStyles.Paragraph.NewParagraph());

			// Respect this order: this is the way the browsers apply them
			String attrValue = en.StyleAttributes["text-align"];
			if (attrValue == null) attrValue = en.Attributes["align"];

			if (attrValue != null)
			{
				JustificationValues? align = ConverterUtility.FormatParagraphAlign(attrValue);
				if (align.HasValue)
				{
					currentParagraph.InsertInProperties(new Justification { Val = align });
				}
			}

			List<OpenXmlElement> styleAttributes = new List<OpenXmlElement>();
			bool newParagraph = ProcessContainerAttributes(en, styleAttributes);

			if (styleAttributes.Count > 0)
				htmlStyles.Runs.BeginTag(en.CurrentTag, styleAttributes.ToArray());

			if (newParagraph)
			{
				AlternateProcessHtmlChunks(en, "</p>");

				CompleteCurrentParagraph();
				AddParagraph(currentParagraph = htmlStyles.Paragraph.NewParagraph());
			}
		}

		#endregion

		#region ProcessPre

		private void ProcessPre(HtmlEnumerator en)
		{
			CompleteCurrentParagraph();
			currentParagraph = htmlStyles.Paragraph.NewParagraph();

			// Oftenly, <pre> tag are used to renders some code examples. They look better inside a table
			if (this.RenderPreAsTable)
			{
				Table currentTable = new Table(
					new TableProperties(
						new TableStyle() { Val = htmlStyles.GetStyle("Table Grid", false) },
						new TableWidth() { Type = TableWidthUnitValues.Pct, Width = "5000" }), // 100% * 50
					new TableGrid(
						new GridColumn() { Width = "5610" }),
					new TableRow(
						new TableCell(
							// Ensure the border lines are visible (regardless of the style used)
							new TableCellProperties(
								new TableCellBorders(
									new TopBorder() { Val = BorderValues.Single },
									new LeftBorder() { Val = BorderValues.Single },
									new BottomBorder() { Val = BorderValues.Single },
									new RightBorder() { Val = BorderValues.Single })),
							currentParagraph))
				);

				AddParagraph(currentTable);
				tables.NewContext(currentTable);
			}
			else
			{
				AddParagraph(currentParagraph);
			}

			// Process the entire <pre> tag and append it to the document
			List<OpenXmlElement> styleAttributes = new List<OpenXmlElement>();
			ProcessContainerAttributes(en, styleAttributes);

			if (styleAttributes.Count > 0)
				htmlStyles.Runs.BeginTag("<pre>", styleAttributes.ToArray());

			AlternateProcessHtmlChunks(en, "</pre>");

			if (styleAttributes.Count > 0)
				htmlStyles.Runs.EndTag("<pre>");

			if (RenderPreAsTable)
				tables.CloseContext();

			currentParagraph.Append(elements);
			elements.Clear();
		}

		#endregion

		#region ProcessSpan

		private void ProcessSpan(HtmlEnumerator en)
		{
			// A span style attribute can contains many information: font color, background color, font size,
			// font family, ...
			// We'll check for each of these and add apply them to the next build runs.

			List<OpenXmlElement> styleAttributes = new List<OpenXmlElement>();
			bool newParagraph = ProcessContainerAttributes(en, styleAttributes);

			if (styleAttributes.Count > 0)
				htmlStyles.Runs.MergeTag("<span>", styleAttributes);

			if (newParagraph)
			{
				AlternateProcessHtmlChunks(en, "</span>");

				CompleteCurrentParagraph();
				AddParagraph(currentParagraph = htmlStyles.Paragraph.NewParagraph());
			}
		}

		#endregion

		#region ProcessStrike

		private void ProcessStrike(HtmlEnumerator en)
		{
			htmlStyles.Runs.BeginTag(en.CurrentTag, new Strike());
		}

		#endregion

		#region ProcessSubscript

		private void ProcessSubscript(HtmlEnumerator en)
		{
			htmlStyles.Runs.BeginTag("<sub>", new VerticalTextAlignment() { Val = VerticalPositionValues.Subscript });
		}

		#endregion

		#region ProcessSuperscript

		private void ProcessSuperscript(HtmlEnumerator en)
		{
			htmlStyles.Runs.BeginTag("<sup>", new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript });
		}

		#endregion

		#region ProcessUnderline

		private void ProcessUnderline(HtmlEnumerator en)
		{
			htmlStyles.Runs.BeginTag(en.CurrentTag, new Underline { Val = UnderlineValues.Single });
		}

		#endregion

		#region ProcessTable

		private void ProcessTable(HtmlEnumerator en)
		{
			IList<OpenXmlElement> properties = new List<OpenXmlElement>();

			int? border = en.Attributes.GetAsInt("border");
			if (border.HasValue && border.Value > 0)
			{
				// If the border has been specified, we display the Table Grid style which display
				// its grid lines. Otherwise the default table style hides the grid lines.
				if (htmlStyles.DoesStyleExists("Table Grid"))
					properties.Add(new TableStyle() { Val = htmlStyles.GetStyle("Table Grid", false) });
				else
				{
					properties.Add(new TableBorders(
						new TopBorder { Val = BorderValues.Single },
						new LeftBorder { Val = BorderValues.Single },
						new RightBorder { Val = BorderValues.Single },
						new BottomBorder { Val = BorderValues.Single },
						new InsideHorizontalBorder { Val = BorderValues.Single },
						new InsideVerticalBorder { Val = BorderValues.Single }
					));
				}
			}

			Unit unit = en.StyleAttributes.GetAsUnit("width");
			if (!unit.IsValid) unit = en.Attributes.GetAsUnit("width");

			if (unit.IsValid)
			{
				switch (unit.Type)
				{
					case "%":
						properties.Add(new TableWidth() { Type = TableWidthUnitValues.Pct, Width = (unit.Value * 50).ToString(CultureInfo.InvariantCulture) }); break;
					case "pt":
						properties.Add(new TableWidth() { Type = TableWidthUnitValues.Dxa, Width = (unit.Value * 20).ToString(CultureInfo.InvariantCulture) }); break;
					case "px":
						properties.Add(new TableWidth() { Type = TableWidthUnitValues.Dxa, Width = (unit.Value).ToString(CultureInfo.InvariantCulture) }); break;
				}
			}
			else
			{
				properties.Add(new TableWidth() { Type = TableWidthUnitValues.Pct, Width = "5000" }); // 100% * 50
			}

			string align = en.Attributes["align"];
			if (align != null)
			{
				JustificationValues? halign = ConverterUtility.FormatParagraphAlign(align);
				if (halign.HasValue)
					properties.Add(new TableJustification() { Val = halign.Value.ToTableRowAlignment() });
			}

			List<OpenXmlElement> runStyleAttributes = new List<OpenXmlElement>();
			htmlStyles.Runs.ProcessCommonRunAttributes(en, runStyleAttributes);
			if (runStyleAttributes.Count > 0)
				htmlStyles.Runs.BeginTag(en.CurrentTag, runStyleAttributes.ToArray());

			Table currentTable = new Table(
				new TableProperties(properties));

			if (tables.HasContext)
			{
				TableCell currentCell = tables.CurrentTable.GetLastChild<TableRow>().GetLastChild<TableCell>();
				currentCell.Append(new Paragraph(elements));
				currentCell.Append(currentTable);
				elements.Clear();
			}
			else
			{
				CompleteCurrentParagraph();
				this.paragraphs.Add(currentTable);
			}

			tables.NewContext(currentTable);
		}

		#endregion

		#region ProcessTableCaption

		private void ProcessTableCaption(HtmlEnumerator en)
		{
			if (!tables.HasContext) return;

			string att = en.StyleAttributes["text-align"];
			if (att == null) att = en.Attributes["align"];

			ProcessHtmlChunks(en, "</caption>");

			var runStyleId = htmlStyles.GetStyle("Subtle Reference", true);
			var legend = new Paragraph(
					new ParagraphProperties(
						new ParagraphStyleId() { Val = htmlStyles.GetStyle("caption", false) },
						new ParagraphMarkRunProperties(
							new RunStyle() { Val = runStyleId })),
					new Run(
						new RunProperties(
							new RunStyle() { Val = runStyleId }),
						new FieldChar() { FieldCharType = FieldCharValues.Begin }),
					new Run(
						new RunProperties(
							new RunStyle() { Val = runStyleId }),
						new FieldCode(" SEQ Tableau \\* ARABIC ") { Space = SpaceProcessingModeValues.Preserve }),
					new Run(
						new RunProperties(
							new RunStyle() { Val = runStyleId }),
						new FieldChar() { FieldCharType = FieldCharValues.End })
				);
			legend.Append(elements);
			elements.Clear();

			if (att != null)
			{
				JustificationValues? align = ConverterUtility.FormatParagraphAlign(att);
				if (align.HasValue)
				{
					legend.InsertInProperties(new Justification { Val = align });
				}
			}
			else
			{
				// If no particular alignement has been specified for the legend, we will align the legend
				// relative to the owning table
				TableProperties props = tables.CurrentTable.GetFirstChild<TableProperties>();
				if (props != null)
				{
					TableJustification justif = props.GetFirstChild<TableJustification>();
					if (justif != null) legend.InsertInProperties(new Justification { Val = justif.Val.Value.ToJustification() });
				}
			}

			if (this.TableCaptionPosition == CaptionPositionValues.Above)
			{
				this.paragraphs.Insert(this.paragraphs.Count - 1, legend);
			}
			else
			{
				this.paragraphs.Add(legend);
			}
		}

		#endregion

		#region ProcessTableRow

		private void ProcessTableRow(HtmlEnumerator en)
		{
			List<OpenXmlElement> styleAttributes = new List<OpenXmlElement>();
			List<OpenXmlElement> runStyleAttributes = new List<OpenXmlElement>();

			htmlStyles.Tables.ProcessCommonAttributes(en, styleAttributes);

			TableRow row = new TableRow();
			if (styleAttributes.Count > 0)
				row.Append(new TableRowProperties(styleAttributes));

			htmlStyles.Runs.ProcessCommonRunAttributes(en, runStyleAttributes);
			if (runStyleAttributes.Count > 0)
				htmlStyles.Runs.BeginTag("<tr>", runStyleAttributes.ToArray());

			tables.CurrentTable.Append(row);
			tables.CellPosition = new Point(0, tables.CellPosition.Y + 1);
		}

		#endregion

		#region ProcessTableColumn

		private void ProcessTableColumn(HtmlEnumerator en)
		{
			List<OpenXmlElement> styleAttributes = new List<OpenXmlElement>();
			List<OpenXmlElement> runStyleAttributes = new List<OpenXmlElement>();

			Unit unit = en.StyleAttributes.GetAsUnit("width");
			if (!unit.IsValid) unit = en.Attributes.GetAsUnit("width");

			if (unit.IsValid)
			{
				switch (unit.Type)
				{
					case "%":
						styleAttributes.Add(new TableCellWidth() { Type = TableWidthUnitValues.Pct, Width = (unit.Value * 50).ToString(CultureInfo.InvariantCulture) });
						break;
					case "pt":
						styleAttributes.Add(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = (unit.Value * 20).ToString(CultureInfo.InvariantCulture) });
						break;
					case "px":
						styleAttributes.Add(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = (unit.Value).ToString(CultureInfo.InvariantCulture) });
						break;
				}
			}

			int? colspan = en.Attributes.GetAsInt("colspan");
			if (colspan.HasValue)
			{
				styleAttributes.Add(new GridSpan() { Val = colspan });
			}

			int? rowspan = en.Attributes.GetAsInt("rowspan");
			if (rowspan.HasValue)
			{
				styleAttributes.Add(new VerticalMerge() { Val = MergedCellValues.Restart });
				tables.RowSpan[tables.CellPosition] = rowspan.Value - 1;
			}

			htmlStyles.Runs.ProcessCommonRunAttributes(en, runStyleAttributes);

			// Manage vertical text (only for table cell)
			string direction = en.StyleAttributes["writing-mode"];
			if (direction != null)
			{
				switch (direction)
				{
					case "tb-lr":
						styleAttributes.Add(new TextDirection() { Val = TextDirectionValues.BottomToTopLeftToRight });
						styleAttributes.Add(new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center });
						htmlStyles.Tables.BeginTagForParagraph("<td>", new Justification() { Val = JustificationValues.Center });
						break;
					case "tb-rl":
						styleAttributes.Add(new TextDirection() { Val = TextDirectionValues.TopToBottomRightToLeft });
						styleAttributes.Add(new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center });
						htmlStyles.Tables.BeginTagForParagraph("<td>", new Justification() { Val = JustificationValues.Center });
						break;
				}
			}

			htmlStyles.Tables.ProcessCommonAttributes(en, styleAttributes);
			if (runStyleAttributes.Count > 0)
				htmlStyles.Runs.BeginTag("<td>", runStyleAttributes.ToArray());

			TableCell cell = new TableCell(
				new TableCellProperties(styleAttributes));
			tables.CurrentTable.GetLastChild<TableRow>().Append(cell);

			if (en.IsSelfClosedTag) // Force a call to ProcessClosingTableColumn
				ProcessClosingTableColumn(en);
			else
			{
				// we create a new currentParagraph to add new runs inside the TableCell
				cell.Append(currentParagraph = new Paragraph());
			}
		}

		#endregion

		#region ProcessTablePart

		private void ProcessTablePart(HtmlEnumerator en)
		{
			List<OpenXmlElement> styleAttributes = new List<OpenXmlElement>();

			htmlStyles.Tables.ProcessCommonAttributes(en, styleAttributes);

			if (styleAttributes.Count > 0)
				htmlStyles.Tables.BeginTag(en.CurrentTag, styleAttributes.ToArray());
		}

		#endregion

		// Closing tags

		#region ProcessClosingBold

		private void ProcessClosingBold(HtmlEnumerator en)
		{
			htmlStyles.Runs.EndTag("<b>");
		}

		#endregion

		#region ProcessClosingDiv

		private void ProcessClosingDiv(HtmlEnumerator en)
		{
			// Mimic the rendering of the browser:
			ProcessBr(en);
			ProcessClosingTag(en);
		}

		#endregion

		#region ProcessClosingItalic

		private void ProcessClosingItalic(HtmlEnumerator en)
		{
			htmlStyles.Runs.EndTag("<i>");
		}

		#endregion

		#region ProcessClosingTag

		private void ProcessClosingTag(HtmlEnumerator en)
		{
			htmlStyles.Runs.EndTag(en.CurrentTag.Replace("/", ""));
		}

		#endregion

		#region ProcessClosingNumberingList

		private void ProcessClosingNumberingList(HtmlEnumerator en)
		{
			numberLevelRef--;

			// If we are no more inside a list, we move to another paragraph (as we created
			// one for containing all the <li>. This will ensure the next run will not be added to the <li>.
			if (numberLevelRef == 0)
				AddParagraph(currentParagraph = htmlStyles.Paragraph.NewParagraph());
		}

		#endregion

		#region ProcessClosingParagraph

		private void ProcessClosingParagraph(HtmlEnumerator en)
		{
			CompleteCurrentParagraph();
			AddParagraph(currentParagraph = htmlStyles.Paragraph.NewParagraph());

			string tag = en.CurrentTag.Replace("/", "");
			htmlStyles.Runs.EndTag(tag);
			htmlStyles.Paragraph.EndTag(tag);
		}

		#endregion

		#region ProcessClosingTable

		private void ProcessClosingTable(HtmlEnumerator en)
		{
			htmlStyles.Tables.EndTag("<table>");
			htmlStyles.Runs.EndTag("<table>");

			TableRow row = tables.CurrentTable.GetFirstChild<TableRow>();
			// Is this a misformed or empty table?
			if (row == null) return;

			// Count the number of tableCell and add as much GridColumn as we need.
			TableGrid grid = new TableGrid();
			for (int i = 0; i < row.ChildElements.Count; i++)
			{
				if (row.ChildElements[i] is TableCell)
				{
					grid.Append(new GridColumn());

					// If that column contains some span, we need to count them also
					GridSpan span = row.ChildElements[i].GetFirstChild<GridSpan>();
					if (span != null)
					{
						for (int j = span.Val; j > 0; j++)
							grid.Append(new GridColumn());
					}
				}
			}

			tables.CurrentTable.PrependChild<TableGrid>(grid);
			tables.CloseContext();

			if (!tables.HasContext)
				AddParagraph(currentParagraph = htmlStyles.Paragraph.NewParagraph());
		}

		#endregion

		#region ProcessClosingTablePart

		private void ProcessClosingTablePart(HtmlEnumerator en)
		{
			string closingTag = en.CurrentTag.Replace("/", "");

			htmlStyles.Tables.EndTag(closingTag);
			htmlStyles.Tables.EndTagForParagraph(closingTag);
		}

		#endregion

		#region ProcessClosingTableRow

		private void ProcessClosingTableRow(HtmlEnumerator en)
		{
			TableRow row = tables.CurrentTable.GetLastChild<TableRow>();

			// Add empty columns to fill rowspan
			if (tables.RowSpan.Count > 0)
			{
				int rowIndex = tables.CellPosition.Y;

				for (int i = 0; i < tables.RowSpan.Count; i++)
				{
					Point position = tables.RowSpan.Keys[i];
					if (position.Y == rowIndex) continue;

					row.InsertAt<TableCell>(new TableCell(new TableCellProperties(
												new VerticalMerge()),
											new Paragraph()),
						position.X);

					int span = tables.RowSpan[position];
					if (span == 1) { tables.RowSpan.RemoveAt(i); i--; }
					else tables.RowSpan[position] = span - 1;
				}
			}

			htmlStyles.Tables.EndTagForParagraph("<tr>");
			htmlStyles.Runs.EndTag("<tr>");
		}

		#endregion

		#region ProcessClosingTableColumn

		private void ProcessClosingTableColumn(HtmlEnumerator en)
		{
			TableCell cell = tables.CurrentTable.GetLastChild<TableRow>().GetLastChild<TableCell>();

			// As we add automatically a paragraph to the cell once we create it, we'll remove it if finally, it was not used.
			// For all the other children, we will ensure there is no more empty paragraphs (similarly to what we do at the end
			// of the convert processing).
			foreach(Paragraph p in cell.Elements<Paragraph>())
			{
				if (!p.HasChild<Run>()) p.Remove();
			}

			// We add this paragraph regardless it has elements or not. A TableCell requires at least a Paragraph.
			// The append should occur after the previous foreach()
			cell.Append(new Paragraph(elements));

			htmlStyles.Tables.ApplyTags(cell);

			// Reset all our variables and move to next cell
			this.elements.Clear();
			htmlStyles.Tables.EndTagForParagraph("<td>");
			htmlStyles.Runs.EndTag("<td>");

			Point pos = tables.CellPosition;
			pos.X++;
			tables.CellPosition = pos;
		}

		#endregion
	}
}

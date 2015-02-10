/* Copyright (C) Olivier Nizet http://html2openxml.codeplex.com - All Rights Reserved
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
using System.Drawing;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace NotesFor.HtmlToOpenXml
{
	using a = DocumentFormat.OpenXml.Drawing;
	using pic = DocumentFormat.OpenXml.Drawing.Pictures;
	using wBorder = DocumentFormat.OpenXml.Wordprocessing.Border;

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

			AlternateProcessHtmlChunks(en, en.ClosingCurrentTag);

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
						new RunProperties {
							RunStyle = new RunStyle() { Val = htmlStyles.GetStyle(runStyle, StyleValues.Character) }
						},
						reference));

				if (!htmlStyles.DoesStyleExists(defaultRefStyle))
				{
					// Force the superscript style because if the footnote text style does not exists,
					// the rendering will be awful.
					run.InsertInProperties(prop =>
						prop.VerticalTextAlignment = new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript });
				}
			}
		}

		#endregion

		#region ProcessBold

		private void ProcessBold(HtmlEnumerator en)
		{
			htmlStyles.Runs.BeginTag(en.CurrentTag, new Bold());
		}

		#endregion

		#region ProcessBlockQuote

		private void ProcessBlockQuote(HtmlEnumerator en)
		{
			CompleteCurrentParagraph(true);

			// for nested paragraphs:
			htmlStyles.Paragraph.BeginTag(en.CurrentTag, new ParagraphStyleId() { Val = htmlStyles.GetStyle("IntenseQuote") });

			// if the style was not yet defined, we force the indentation
			if (!htmlStyles.DoesStyleExists("IntenseQuote"))
				htmlStyles.Paragraph.BeginTag(en.CurrentTag, new Indentation() { Left = "708" });

			// TODO: handle attribute cite and create footnote
		}

		#endregion

		#region ProcessBody

		private void ProcessBody(HtmlEnumerator en)
		{
			List<OpenXmlElement> styleAttributes = new List<OpenXmlElement>();
			htmlStyles.Paragraph.ProcessCommonAttributes(en, styleAttributes);

			if (styleAttributes.Count > 0)
				htmlStyles.Runs.BeginTag(en.CurrentTag, styleAttributes.ToArray());

			// Unsupported W3C attribute but claimed by users. Specified at <body> level, the page
			// orientation is applied on the whole document
			string attr = en.StyleAttributes["page-orientation"];
			if (attr != null)
			{
				PageOrientationValues orientation = ConverterUtility.ConvertToPageOrientation(attr);

                SectionProperties sectionProperties = mainPart.Document.Body.GetFirstChild<SectionProperties>();
                if (sectionProperties == null || sectionProperties.GetFirstChild<PageSize>() == null)
                {
                    mainPart.Document.Body.Append(HtmlConverter.ChangePageOrientation(orientation));
                }
                else
                {
                    PageSize pageSize = sectionProperties.GetFirstChild<PageSize>();
                    if (!pageSize.Compare(orientation))
                    {
                        SectionProperties validSectionProp = ChangePageOrientation(orientation);
                        if (pageSize != null) pageSize.Remove();
                        sectionProperties.PrependChild(validSectionProp.GetFirstChild<PageSize>().CloneNode(true));
                    }
                }
            }
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
			htmlStyles.Runs.BeginTag(en.CurrentTag, new RunStyle() { Val = htmlStyles.GetStyle("Quote", StyleValues.Character) });
		}

		#endregion

		#region ProcessDefinitionList

		private void ProcessDefinitionList(HtmlEnumerator en)
		{
			ProcessParagraph(en);
			currentParagraph.InsertInProperties(prop => prop.SpacingBetweenLines = new SpacingBetweenLines() { After = "0" });
		}

		#endregion

		#region ProcessDefinitionListItem

		private void ProcessDefinitionListItem(HtmlEnumerator en)
		{
			AlternateProcessHtmlChunks(en, "</dd>");

			currentParagraph = htmlStyles.Paragraph.NewParagraph();
			currentParagraph.Append(elements);
			currentParagraph.InsertInProperties(prop => {
				prop.Indentation = new Indentation() { FirstLine = "708" };
				prop.SpacingBetweenLines = new SpacingBetweenLines() { After = "0" };
			});

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
			if (en.Attributes.Count == 0 || (en.StyleAttributes["text-align"] == null && en.Attributes["align"] == null && en.StyleAttributes.GetAsBorder("border").IsEmpty))
			{
				List<OpenXmlElement> runStyleAttributes = new List<OpenXmlElement>();
				bool newParagraph = ProcessContainerAttributes(en, runStyleAttributes);
				CompleteCurrentParagraph(newParagraph);

				if (runStyleAttributes.Count > 0)
					htmlStyles.Runs.BeginTag(en.CurrentTag, runStyleAttributes);

				// Any changes that requires a new paragraph?
				if (newParagraph)
				{
					// Insert before the break, complete this paragraph and start a new one
					this.paragraphs.Insert(this.paragraphs.Count - 1, currentParagraph);
					AlternateProcessHtmlChunks(en, en.ClosingCurrentTag);
					CompleteCurrentParagraph();
				}
			}
			else
			{
				// treat div as a paragraph
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
				Unit fontSize = ConverterUtility.ConvertToFontSize(attrValue);
				if (fontSize.IsValid)
					styleAttributes.Add(new FontSize { Val = (fontSize.ValueInPoint * 2).ToString(CultureInfo.InvariantCulture) });
			}

			attrValue = en.Attributes["face"];
			if (attrValue != null)
			{
				// Set HightAnsi. Bug fixed by xjpmauricio on http://html2openxml.codeplex.com/discussions/285439
				// where characters with accents were always using fallback font
				styleAttributes.Add(new RunFonts { Ascii = attrValue, HighAnsi = attrValue });
			}

			if (styleAttributes.Count > 0)
				htmlStyles.Runs.MergeTag(en.CurrentTag, styleAttributes);
		}

		#endregion

		#region ProcessHeading

		private void ProcessHeading(HtmlEnumerator en)
		{
			char level = en.Current[2];
			string clsName = "heading " + level;

			// ensure style exists or the heading will be displayed as a normal text
			if (!htmlStyles.DoesStyleExists(clsName))
			{
				String normalStyleName = htmlStyles.GetStyle("Normal", StyleValues.Paragraph);
				htmlStyles.AddStyle(clsName, new Style(
					new StyleName() { Val = clsName },
					new BasedOn() { Val = normalStyleName },
					new NextParagraphStyle() { Val = normalStyleName },
					new UnhideWhenUsed(),
					new PrimaryStyle(),
					new StyleParagraphProperties() {
						KeepNext = new KeepNext(),
						KeepLines = new KeepLines(),
						SpacingBetweenLines = new SpacingBetweenLines(){ Before = "200", After = "0" },
						OutlineLevel = new OutlineLevel() { Val = (int) Char.GetNumericValue(level) - 1 /* outline starts at 0 */ }
					},
					new StyleRunProperties(PredefinedStyles.ResourceManager.GetString(clsName))
				) { Type = StyleValues.Paragraph, StyleId = clsName });
			}

			// support also style attributes for heading (in case of theme override)
			List<OpenXmlElement> styleAttributes = new List<OpenXmlElement>();
			htmlStyles.Paragraph.ProcessCommonAttributes(en, styleAttributes);

			AlternateProcessHtmlChunks(en, "</h" + level + ">");
			Paragraph p = new Paragraph(elements);
			p.InsertInProperties(prop =>
				prop.ParagraphStyleId = new ParagraphStyleId() { Val = htmlStyles.GetStyle(clsName, StyleValues.Paragraph) });

			htmlStyles.Paragraph.ApplyTags(p);
			htmlStyles.Paragraph.EndTag("<h" + level + ">");

			this.elements.Clear();
			AddParagraph(p);
			AddParagraph(currentParagraph = htmlStyles.Paragraph.NewParagraph());
		}

		#endregion

		#region ProcessHorizontalLine

		private void ProcessHorizontalLine(HtmlEnumerator en)
		{
			// Insert an horizontal line as it stands in many emails.
            CompleteCurrentParagraph(true);

			// If the previous paragraph contains a bottom border or is a Table, we add some spacing between the <hr>
			// and the previous element or Word will display only the last border.
			// (see Remarks: http://msdn.microsoft.com/en-us/library/documentformat.openxml.wordprocessing.bottomborder%28office.14%29.aspx)
            if (paragraphs.Count > 2)
            {
                OpenXmlCompositeElement previousElement = paragraphs[paragraphs.Count - 2];
                bool addSpacing = false;
                ParagraphProperties prop = previousElement.GetFirstChild<ParagraphProperties>();
                if (prop != null)
                {
                    if (prop.ParagraphBorders != null && prop.ParagraphBorders.BottomBorder != null
                        && prop.ParagraphBorders.BottomBorder.Size > 0U)
                            addSpacing = true;
                }
                else
                {
                    if (previousElement is Table)
                        addSpacing = true;
                }


                if (addSpacing)
                {
                    currentParagraph.InsertInProperties(p => p.SpacingBetweenLines = new SpacingBetweenLines() { Before = "240" });
                }
			}

			// if this paragraph has no children, it will be deleted in RemoveEmptyParagraphs()
			// in order to kept the <hr>, we force an empty run
            currentParagraph.Append(new Run());

            currentParagraph.InsertInProperties(prop => 
				prop.ParagraphBorders = new ParagraphBorders {
					TopBorder = new TopBorder() { Val = BorderValues.Single, Size = 4U }
				});
		}

		#endregion

		#region ProcessHtml

		private void ProcessHtml(HtmlEnumerator en)
		{
			List<OpenXmlElement> styleAttributes = new List<OpenXmlElement>();
			htmlStyles.Paragraph.ProcessCommonAttributes(en, styleAttributes);

			if (styleAttributes.Count > 0)
				htmlStyles.Runs.BeginTag(en.CurrentTag, styleAttributes.ToArray());
		}

		#endregion

		#region ProcessFigureCaption

		private void ProcessFigureCaption(HtmlEnumerator en)
		{
			this.CompleteCurrentParagraph(true);
			EnsureCaptionStyle();

			currentParagraph.Append(
					new ParagraphProperties {
						ParagraphStyleId = new ParagraphStyleId() { Val = htmlStyles.GetStyle("caption", StyleValues.Paragraph) },
						KeepNext = new KeepNext()
					},
					new Run(
						new Text("Figure ") { Space = SpaceProcessingModeValues.Preserve }
					),
					new SimpleField(
						new Run(
							new Text(AddFigureCaption().ToString(CultureInfo.InvariantCulture)))
					) { Instruction = " SEQ Figure \\* ARABIC " }
				);

			ProcessHtmlChunks(en, "</figcaption>");

			if (elements.Count > 0) // any caption?
			{
				Text t = (elements[0] as Run).GetFirstChild<Text>();
				t.Text = " " + t.InnerText; // append a space after the numero of the picture
			}

			this.CompleteCurrentParagraph(true);
		}

		#endregion

		#region ProcessImage

		private void ProcessImage(HtmlEnumerator en)
		{
			if (this.ImageProcessing == ImageProcessing.Ignore) return;

			Drawing drawing = null;
			wBorder border = new wBorder() { Val = BorderValues.None };
			string src = en.Attributes["src"];
			Uri uri = null;

			// Bug reported by Erik2014. Inline 64 bit images can be too big and Uri.TryCreate will fail silently with a SizeLimit error.
			// To circumvent this buffer size, we will work either on the Uri, either on the original src.
			if (src != null && (DataUri.IsWellFormed(src) || Uri.TryCreate(src, UriKind.RelativeOrAbsolute, out uri)))
			{
				string alt = en.Attributes["title"] ?? en.Attributes["alt"];
				bool process = true;

				if (uri != null && !uri.IsAbsoluteUri && this.BaseImageUrl != null)
					uri = new Uri(this.BaseImageUrl, uri);

				Size preferredSize = Size.Empty;
				Unit wu = en.Attributes.GetAsUnit("width");
				if (!wu.IsValid) wu = en.StyleAttributes.GetAsUnit("width");
				Unit hu = en.Attributes.GetAsUnit("height");
				if (!hu.IsValid) hu = en.StyleAttributes.GetAsUnit("height");

				// % is not supported
				if (wu.IsValid && wu.Value > 0 && wu.Type != UnitMetric.Percent)
				{
					preferredSize.Width = wu.ValueInPx;
				}
				if (hu.IsValid && hu.Value > 0 && wu.Type != UnitMetric.Percent)
				{
					// Image perspective skewed. Bug fixed by ddeforge on http://html2openxml.codeplex.com/discussions/350500
					preferredSize.Height = hu.ValueInPx;
				}

				SideBorder attrBorder = en.StyleAttributes.GetAsSideBorder("border");
				if (attrBorder.IsValid)
				{
					border.Val = attrBorder.Style;
					border.Color = attrBorder.Color.ToHexString();
					border.Size = (uint) attrBorder.Width.ValueInPx * 4;
				}

				if (process)
					drawing = AddImagePart(uri, src, alt, preferredSize);
			}

			if (drawing != null)
			{
				Run run = new Run(drawing);
				if (border.Val != BorderValues.None) run.InsertInProperties(prop => prop.Border = border);
				elements.Add(run);
			}
		}

		#endregion

		#region ProcessItalic

		private void ProcessItalic(HtmlEnumerator en)
		{
			htmlStyles.Runs.BeginTag(en.CurrentTag, new Italic());
		}

		#endregion

		#region ProcessLi

		private void ProcessLi(HtmlEnumerator en)
		{
			CompleteCurrentParagraph(false);
			currentParagraph = htmlStyles.Paragraph.NewParagraph();

			int numberingId = htmlStyles.NumberingList.ProcessItem(en);
			int level = htmlStyles.NumberingList.LevelIndex;

			// Save the new paragraph reference to support nested numbering list.
			Paragraph p = currentParagraph;
			currentParagraph.InsertInProperties(prop => {
				prop.ParagraphStyleId = new ParagraphStyleId() { Val = htmlStyles.GetStyle("ListParagraph", StyleValues.Paragraph) };
				prop.SpacingBetweenLines = new SpacingBetweenLines() { After = "0" };
				prop.Indentation = new Indentation() { Hanging = "357", Left = (level * 357).ToString(CultureInfo.InvariantCulture) };
				prop.NumberingProperties = new NumberingProperties {
					NumberingLevelReference = new NumberingLevelReference() { Val = level - 1 },
					NumberingId = new NumberingId() { Val = numberingId }
				};
			});

			// Restore the original elements list
			AddParagraph(currentParagraph);

			// Continue to process the html until we found </li>
			HtmlStyles.Paragraph.ApplyTags(currentParagraph);
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
				// handle link where the http:// is missing and that starts directly with www
				if(att.StartsWith("www.", StringComparison.OrdinalIgnoreCase))
					att = Uri.UriSchemeHttp + Uri.SchemeDelimiter + att;

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

			if (elements.Count == 0) return;

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
			}

			// Append the processed elements and put them to the Run of the Hyperlink
			h.Append(elements);

			// can't use GetFirstChild<Run> or we may find the one containing the image
			foreach (var el in h.ChildElements)
			{
				Run run = el as Run;
				if (run != null && !run.HasChild<Drawing>())
				{
					if (!htmlStyles.DoesStyleExists("Hyperlink"))
					{
						htmlStyles.AddStyle("Hyperlink", new Style(
							new StyleName() { Val = "Hyperlink" },
							new UnhideWhenUsed(),
							new StyleRunProperties(PredefinedStyles.HyperLink)
						) { Type = StyleValues.Character, StyleId = "Hyperlink" });
					}

					run.InsertInProperties(prop =>
						prop.RunStyle = new RunStyle() { Val = htmlStyles.GetStyle("Hyperlink", StyleValues.Character) });
					break;
				}
			}

			this.elements.Clear();

			// Append the hyperlink
			elements.Add(h);

			if (imageInLink.Count > 0) CompleteCurrentParagraph(true);
		}

		#endregion

		#region ProcessNumberingList

		private void ProcessNumberingList(HtmlEnumerator en)
		{
			htmlStyles.NumberingList.BeginList(en);
		}

		#endregion

		#region ProcessParagraph

		private void ProcessParagraph(HtmlEnumerator en)
		{
			CompleteCurrentParagraph(true);

			// Respect this order: this is the way the browsers apply them
			String attrValue = en.StyleAttributes["text-align"];
			if (attrValue == null) attrValue = en.Attributes["align"];

			if (attrValue != null)
			{
				JustificationValues? align = ConverterUtility.FormatParagraphAlign(attrValue);
				if (align.HasValue)
				{
					currentParagraph.InsertInProperties(prop => prop.Justification = new Justification { Val = align });
				}
			}

			List<OpenXmlElement> styleAttributes = new List<OpenXmlElement>();
			bool newParagraph = ProcessContainerAttributes(en, styleAttributes);

			if (styleAttributes.Count > 0)
				htmlStyles.Runs.BeginTag(en.CurrentTag, styleAttributes.ToArray());

			if (newParagraph)
			{
				AlternateProcessHtmlChunks(en, en.ClosingCurrentTag);
				ProcessClosingParagraph(en);
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
                    new TableProperties
                    {
                        TableStyle = new TableStyle() { Val = htmlStyles.GetStyle("Table Grid", StyleValues.Paragraph) },
                        TableWidth = new TableWidth() { Type = TableWidthUnitValues.Pct, Width = "5000" }, // 100% * 50
                    },
                    new TableGrid(
                        new GridColumn() { Width = "5610" }),
                    new TableRow(
                        new TableCell(
                    // Ensure the border lines are visible (regardless of the style used)
                            new TableCellProperties
                            {
                                TableCellBorders = new TableCellBorders(
                                   new TopBorder() { Val = BorderValues.Single },
                                   new LeftBorder() { Val = BorderValues.Single },
                                   new BottomBorder() { Val = BorderValues.Single },
                                   new RightBorder() { Val = BorderValues.Single })
                            },
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
				htmlStyles.Runs.BeginTag(en.CurrentTag, styleAttributes.ToArray());

			AlternateProcessHtmlChunks(en, "</pre>");

			if (styleAttributes.Count > 0)
				htmlStyles.Runs.EndTag(en.CurrentTag);

			if (RenderPreAsTable)
				tables.CloseContext();

			CompleteCurrentParagraph();
		}

		#endregion

		#region ProcessQuote

		private void ProcessQuote(HtmlEnumerator en)
		{
			// The browsers render the quote tag between a kind of separators.
			// We add the Quote style to the nested runs to match more Word.

			htmlStyles.Runs.BeginTag(en.CurrentTag, new RunStyle() { Val = htmlStyles.GetStyle("Quote", StyleValues.Character) });

			Run run = new Run(
				new Text(" " + HtmlStyles.QuoteCharacters.chars[0]) { Space = SpaceProcessingModeValues.Preserve }
			);

			htmlStyles.Runs.ApplyTags(run);
			elements.Add(run);
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
				htmlStyles.Runs.MergeTag(en.CurrentTag, styleAttributes);

			if (newParagraph)
			{
				AlternateProcessHtmlChunks(en, en.ClosingCurrentTag);
				CompleteCurrentParagraph(true);
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
			htmlStyles.Runs.BeginTag(en.CurrentTag, new VerticalTextAlignment() { Val = VerticalPositionValues.Subscript });
		}

		#endregion

		#region ProcessSuperscript

		private void ProcessSuperscript(HtmlEnumerator en)
		{
			htmlStyles.Runs.BeginTag(en.CurrentTag, new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript });
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
			TableProperties properties = new TableProperties();
			Table currentTable = new Table(properties);

			int? border = en.Attributes.GetAsInt("border");
			if (border.HasValue && border.Value > 0)
			{
				// If the border has been specified, we display the Table Grid style which display
				// its grid lines. Otherwise the default table style hides the grid lines.
				if (htmlStyles.DoesStyleExists("Table Grid"))
					properties.TableStyle = new TableStyle() { Val = htmlStyles.GetStyle("Table Grid", StyleValues.Paragraph) };
				else
				{
					properties.TableBorders = new TableBorders() {
						TopBorder = new TopBorder { Val = BorderValues.Single },
						LeftBorder = new LeftBorder { Val = BorderValues.Single },
						RightBorder = new RightBorder { Val = BorderValues.Single },
						BottomBorder = new BottomBorder { Val = BorderValues.Single },
						InsideHorizontalBorder = new InsideHorizontalBorder { Val = BorderValues.Single },
						InsideVerticalBorder = new InsideVerticalBorder { Val = BorderValues.Single }
					};
				}
			}

			Unit unit = en.StyleAttributes.GetAsUnit("width");
			if (!unit.IsValid) unit = en.Attributes.GetAsUnit("width");

			if (unit.IsValid)
			{
				switch (unit.Type)
				{
					case UnitMetric.Percent:
						properties.TableWidth = new TableWidth() { Type = TableWidthUnitValues.Pct, Width = (unit.Value * 50).ToString(CultureInfo.InvariantCulture) }; break;
					case UnitMetric.Point:
						properties.TableWidth = new TableWidth() { Type = TableWidthUnitValues.Dxa, Width = unit.ValueInDxa.ToString(CultureInfo.InvariantCulture) }; break;
					case UnitMetric.Pixel:
						properties.TableWidth = new TableWidth() { Type = TableWidthUnitValues.Dxa, Width = unit.ValueInDxa.ToString(CultureInfo.InvariantCulture) }; break;
				}
			}
			else
			{
				// Use Auto=0 instead of Pct=auto
				// bug reported by scarhand (https://html2openxml.codeplex.com/workitem/12494)
				properties.TableWidth = new TableWidth() { Type = TableWidthUnitValues.Auto, Width = "0" };
			}

			string classValue = en.Attributes["class"];
			if (classValue != null)
			{
				classValue = htmlStyles.GetStyle(classValue, StyleValues.Table, ignoreCase: true);
				if (classValue != null)
					properties.TableStyle = new TableStyle() { Val = classValue };
			}

			string align = en.Attributes["align"];
			if (align != null)
			{
				JustificationValues? halign = ConverterUtility.FormatParagraphAlign(align);
				if (halign.HasValue)
					properties.TableJustification = new TableJustification() { Val = halign.Value.ToTableRowAlignment() };
			}

			// only if the table is left aligned, we can handle some left margin indentation
			// Right margin + Right align has no equivalent in OpenXml
			if (align == null || align == "left")
			{
				Margin margin = en.StyleAttributes.GetAsMargin("margin");

				if (margin.IsValid)
				{
					// OpenXml doesn't support table margin in Percent, but Html does
					// the margin part has been implemented by Olek (patch #8457)

					TableCellMarginDefault cellMargin = new TableCellMarginDefault();
					if (margin.Left.IsValid && margin.Left.Type != UnitMetric.Percent)
						cellMargin.TableCellLeftMargin = new TableCellLeftMargin() { Type = TableWidthValues.Dxa, Width = (short) margin.Left.ValueInDxa };
					if (margin.Right.IsValid && margin.Right.Type != UnitMetric.Percent)
						cellMargin.TableCellRightMargin = new TableCellRightMargin() { Type = TableWidthValues.Dxa, Width = (short) margin.Right.ValueInDxa };
					if (margin.Top.IsValid && margin.Top.Type != UnitMetric.Percent)
						cellMargin.TopMargin = new TopMargin() { Type = TableWidthUnitValues.Dxa, Width = margin.Top.ValueInDxa.ToString(CultureInfo.InvariantCulture) };
					if (margin.Bottom.IsValid && margin.Bottom.Type != UnitMetric.Percent)
						cellMargin.BottomMargin = new BottomMargin() { Type = TableWidthUnitValues.Dxa, Width = margin.Bottom.ValueInDxa.ToString(CultureInfo.InvariantCulture) };

					if (cellMargin.HasChildren)
						properties.TableCellMarginDefault = cellMargin;
				}
			}

			int? spacing = en.Attributes.GetAsInt("cellspacing");
			if (spacing.HasValue)
				properties.TableCellSpacing = new TableCellSpacing { Type = TableWidthUnitValues.Dxa, Width = spacing.Value.ToString(CultureInfo.InvariantCulture) };

			int? padding = en.Attributes.GetAsInt("cellpadding");
			if (padding.HasValue)
				properties.TableIndentation = new TableIndentation { Type = TableWidthUnitValues.Dxa, Width = padding.Value };

			List<OpenXmlElement> runStyleAttributes = new List<OpenXmlElement>();
			htmlStyles.Tables.ProcessCommonAttributes(en, runStyleAttributes);
			if (runStyleAttributes.Count > 0)
				htmlStyles.Runs.BeginTag(en.CurrentTag, runStyleAttributes.ToArray());


			// are we currently inside another table?
			if (tables.HasContext)
			{
				// Okay we will insert nested table but beware the paragraph inside TableCell should contains at least 1 run.

				TableCell currentCell = tables.CurrentTable.GetLastChild<TableRow>().GetLastChild<TableCell>();
				if (elements.Count == 0) elements.Add(new Run(new Text("")));
				currentCell.Append(new Paragraph(elements), currentTable);
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

			var runStyleId = htmlStyles.GetStyle("Subtle Reference", StyleValues.Character);
			var legend = new Paragraph(
					new ParagraphProperties {
						ParagraphStyleId = new ParagraphStyleId() { Val = htmlStyles.GetStyle("caption", StyleValues.Paragraph) },
						ParagraphMarkRunProperties = new ParagraphMarkRunProperties(
							new RunStyle() { Val = runStyleId })
					},
					new Run(
						new RunProperties {
							RunStyle = new RunStyle() { Val = runStyleId }
						},
						new FieldChar() { FieldCharType = FieldCharValues.Begin }),
					new Run(
						new RunProperties {
							RunStyle = new RunStyle() { Val = runStyleId }
						},
						new FieldCode(" SEQ TABLE \\* ARABIC ") { Space = SpaceProcessingModeValues.Preserve }),
					new Run(
						new RunProperties {
							RunStyle = new RunStyle() { Val = runStyleId }
						},
						new FieldChar() { FieldCharType = FieldCharValues.End })
				);
			legend.Append(elements);
			elements.Clear();

			if (att != null)
			{
				JustificationValues? align = ConverterUtility.FormatParagraphAlign(att);
				if (align.HasValue)
					legend.InsertInProperties(prop => prop.Justification = new Justification { Val = align } );
			}
			else
			{
				// If no particular alignement has been specified for the legend, we will align the legend
				// relative to the owning table
				TableProperties props = tables.CurrentTable.GetFirstChild<TableProperties>();
				if (props != null)
				{
					TableJustification justif = props.GetFirstChild<TableJustification>();
					if (justif != null) legend.InsertInProperties(prop =>
						prop.Justification = new Justification { Val = justif.Val.Value.ToJustification() });
				}
			}

			if (this.TableCaptionPosition == CaptionPositionValues.Above)
				this.paragraphs.Insert(this.paragraphs.Count - 1, legend);
			else
				this.paragraphs.Add(legend);

			EnsureCaptionStyle();
		}

		#endregion

		#region ProcessTableRow

		private void ProcessTableRow(HtmlEnumerator en)
		{
			// in case the html is bad-formed and use <tr> outside a <table> tag, we will ensure
			// a table context exists.
			if (!tables.HasContext) return;

			TableRowProperties properties = new TableRowProperties();
			List<OpenXmlElement> runStyleAttributes = new List<OpenXmlElement>();

			htmlStyles.Tables.ProcessCommonAttributes(en, runStyleAttributes);


			Unit unit = en.StyleAttributes.GetAsUnit("height");
			if (!unit.IsValid) unit = en.Attributes.GetAsUnit("height");

			if (unit.IsValid)
			{
				switch (unit.Type)
				{
					case UnitMetric.Point:
						properties.Append(new TableRowHeight() { HeightType = HeightRuleValues.AtLeast, Val = (uint) (unit.Value * 20) });
						break;
					case UnitMetric.Pixel:
						properties.Append(new TableRowHeight() { HeightType = HeightRuleValues.AtLeast, Val = (uint) unit.ValueInDxa });
						break;
				}
			}

			TableRow row = new TableRow();
			if (properties.HasChildren) row.AppendChild(properties);

			htmlStyles.Runs.ProcessCommonAttributes(en, runStyleAttributes);
			if (runStyleAttributes.Count > 0)
				htmlStyles.Runs.BeginTag(en.CurrentTag, runStyleAttributes.ToArray());

			tables.CurrentTable.Append(row);
			tables.CellPosition = new Point(0, tables.CellPosition.Y + 1);
			tables.IndexColumn = 0;
		}

		#endregion

		#region ProcessTableColumn

		private void ProcessTableColumn(HtmlEnumerator en)
		{
			if (!tables.HasContext) return;

			TableCellProperties properties = new TableCellProperties();

			List<OpenXmlElement> styleAttributes = new List<OpenXmlElement>();
			List<OpenXmlElement> runStyleAttributes = new List<OpenXmlElement>();

			Unit unit = en.StyleAttributes.GetAsUnit("width");
			if (!unit.IsValid) unit = en.Attributes.GetAsUnit("width");

			if (unit.IsValid)
			{
				switch (unit.Type)
				{
					case UnitMetric.Percent:
						properties.TableCellWidth = new TableCellWidth() { Type = TableWidthUnitValues.Pct, Width = (unit.Value * 50).ToString(CultureInfo.InvariantCulture) };
						break;
					case UnitMetric.Point:
						properties.TableCellWidth = new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = (unit.Value * 20).ToString(CultureInfo.InvariantCulture) };
						break;
					case UnitMetric.Pixel:
						properties.TableCellWidth = new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = (unit.Value).ToString(CultureInfo.InvariantCulture) };
						break;
				}
			}

			// any code related to tables.IndexColumn is deal to handle row span (implemented by daviderapicavoli)
			tables.IndexColumn++;

			// fix an issue when specifying the RowSpan or ColSpan=1 (reported by imagremlin)
			int? colspan = en.Attributes.GetAsInt("colspan");
			if (colspan.HasValue && colspan.Value > 1)
			{
				properties.GridSpan = new GridSpan() { Val = colspan };
				tables.IndexColumn += colspan.Value - 1;
			}

			int? rowspan = en.Attributes.GetAsInt("rowspan");
			if (rowspan.HasValue && rowspan.Value > 1)
			{
				properties.VerticalMerge = new VerticalMerge() { Val = MergedCellValues.Restart };

				Point p = new Point(tables.CellPosition.X, tables.CellPosition.Y);
				if (p.X > 0) p.X = tables.IndexColumn - 1;
				tables.RowSpan[p] = rowspan.Value - 1;
			}

			// Manage vertical text (only for table cell)
			string direction = en.StyleAttributes["writing-mode"];
			if (direction != null)
			{
				switch (direction)
				{
					case "tb-lr":
						properties.TextDirection = new TextDirection() { Val = TextDirectionValues.BottomToTopLeftToRight };
						properties.TableCellVerticalAlignment = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };
						htmlStyles.Tables.BeginTagForParagraph(en.CurrentTag, new Justification() { Val = JustificationValues.Center });
						break;
					case "tb-rl":
						properties.TextDirection = new TextDirection() { Val = TextDirectionValues.TopToBottomRightToLeft };
						properties.TableCellVerticalAlignment = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };
						htmlStyles.Tables.BeginTagForParagraph(en.CurrentTag, new Justification() { Val = JustificationValues.Center });
						break;
				}
			}

			var padding = en.StyleAttributes.GetAsMargin("padding");
			if (!padding.IsEmpty)
			{
				TableCellMargin cellMargin = new TableCellMargin();
				var cellMarginSide = new List<KeyValuePair<Unit, TableWidthType>>();
				cellMarginSide.Add(new KeyValuePair<Unit, TableWidthType>(padding.Top, new TopMargin()));
				cellMarginSide.Add(new KeyValuePair<Unit, TableWidthType>(padding.Right, new RightMargin()));
				cellMarginSide.Add(new KeyValuePair<Unit, TableWidthType>(padding.Bottom, new BottomMargin()));
				cellMarginSide.Add(new KeyValuePair<Unit, TableWidthType>(padding.Left, new LeftMargin()));

				foreach (var pair in cellMarginSide)
				{
					if (!pair.Key.IsValid || pair.Key.Value == 0) continue;
					if (pair.Key.Type == UnitMetric.Percent)
					{
						pair.Value.Width = (pair.Key.Value * 50).ToString(CultureInfo.InvariantCulture);
						pair.Value.Type = TableWidthUnitValues.Pct;
					}
					else
					{
						pair.Value.Width = pair.Key.ValueInDxa.ToString(CultureInfo.InvariantCulture);
						pair.Value.Type = TableWidthUnitValues.Dxa;
					}

					cellMargin.Append(pair.Value);
				}

				properties.TableCellMargin = cellMargin;
			}

			htmlStyles.Tables.ProcessCommonAttributes(en, styleAttributes);
			if (styleAttributes.Count > 0)
				htmlStyles.Tables.BeginTag(en.CurrentTag, styleAttributes);
			if (runStyleAttributes.Count > 0)
				htmlStyles.Runs.BeginTag(en.CurrentTag, runStyleAttributes.ToArray());

			TableCell cell = new TableCell();
			if (properties.HasChildren) cell.AppendChild(properties);

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

		#region ProcessXmlDataIsland

		private void ProcessXmlDataIsland(HtmlEnumerator en)
		{
			// Process inner Xml data island and do nothing.
			// The Xml has this format:
			/* <?xml:namespace prefix=o ns="urn:schemas-microsoft-com:office:office">
			   <globalGuideLine>
				   <employee>
					  <FirstName>Austin</FirstName>
					  <LastName>Hennery</LastName>
				   </employee>
			   </globalGuideLine>
			 */

			// Move to the first root element of the Xml then process until the end of the xml chunks.
			while (en.MoveNext() && !en.IsCurrentHtmlTag) ;

			if (en.Current != null)
			{
				string xmlRootElement = en.ClosingCurrentTag;
				while (en.MoveUntilMatch(xmlRootElement)) ;
			}
		}

		#endregion

		// Closing tags

		#region ProcessClosingBlockQuote

		private void ProcessClosingBlockQuote(HtmlEnumerator en)
		{
			CompleteCurrentParagraph(true);
			htmlStyles.Paragraph.BeginTag("<blockquote>");
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

		#region ProcessClosingTag

		private void ProcessClosingTag(HtmlEnumerator en)
		{
			string openingTag = en.CurrentTag.Replace("/", "");
			htmlStyles.Runs.EndTag(openingTag);
			htmlStyles.Paragraph.EndTag(openingTag);
		}

		#endregion

		#region ProcessClosingNumberingList

		private void ProcessClosingNumberingList(HtmlEnumerator en)
		{
			htmlStyles.NumberingList.EndList();

			// If we are no more inside a list, we move to another paragraph (as we created
			// one for containing all the <li>. This will ensure the next run will not be added to the <li>.
			if (htmlStyles.NumberingList.LevelIndex == 0)
				AddParagraph(currentParagraph = htmlStyles.Paragraph.NewParagraph());
		}

		#endregion

		#region ProcessClosingParagraph

		private void ProcessClosingParagraph(HtmlEnumerator en)
		{
			CompleteCurrentParagraph(true);

			string tag = en.CurrentTag.Replace("/", "");
			htmlStyles.Runs.EndTag(tag);
			htmlStyles.Paragraph.EndTag(tag);
		}

		#endregion

		#region ProcessClosingQuote

		private void ProcessClosingQuote(HtmlEnumerator en)
		{
			Run run = new Run(
				new Text(HtmlStyles.QuoteCharacters.chars[1]) { Space = SpaceProcessingModeValues.Preserve }
			);
			htmlStyles.Runs.ApplyTags(run);
			elements.Add(run);

			htmlStyles.Runs.EndTag("<q>");
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

			tables.CurrentTable.InsertAt<TableGrid>(grid, 1);
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
		}

		#endregion

		#region ProcessClosingTableRow

		private void ProcessClosingTableRow(HtmlEnumerator en)
		{
			if (!tables.HasContext) return;
			TableRow row = tables.CurrentTable.GetLastChild<TableRow>();
			if (row == null) return;

			// Word will not open documents with empty rows (reported by scwebgroup)
			if (row.GetFirstChild<TableCell>() == null)
			{
				row.Remove();
				return;
			}

			// Add empty columns to fill rowspan
			if (tables.RowSpan.Count > 0)
			{
				int rowIndex = tables.CellPosition.Y;

				for (int i = 0; i < tables.RowSpan.Count; i++)
				{
					Point position = tables.RowSpan.Keys[i];
					if (position.Y == rowIndex) continue;

					row.InsertAt<TableCell>(new TableCell(new TableCellProperties {
												VerticalMerge = new VerticalMerge() },
											new Paragraph()),
						position.X);

					int span = tables.RowSpan[position];
					if (span == 1) { tables.RowSpan.RemoveAt(i); i--; }
					else tables.RowSpan[position] = span - 1;
				}
			}

			htmlStyles.Tables.EndTag("<tr>");
			htmlStyles.Runs.EndTag("<tr>");
		}

		#endregion

		#region ProcessClosingTableColumn

		private void ProcessClosingTableColumn(HtmlEnumerator en)
		{
			if (!tables.HasContext)
			{
				// When the Html is bad-formed and doesn't contain <table>, the browser renders the column separated by a space.
				// So we do the same here
				Run run = new Run(new Text(" ") { Space = SpaceProcessingModeValues.Preserve });
				htmlStyles.Runs.ApplyTags(run);
				elements.Add(run);
				return;
			}
			TableCell cell = tables.CurrentTable.GetLastChild<TableRow>().GetLastChild<TableCell>();

			// As we add automatically a paragraph to the cell once we create it, we'll remove it if finally, it was not used.
			// For all the other children, we will ensure there is no more empty paragraphs (similarly to what we do at the end
			// of the convert processing).
			// use a basic loop instead of foreach to allow removal (bug reported by antgraf)
			for (int i=0; i<cell.ChildElements.Count; )
			{
				Paragraph p = cell.ChildElements[i] as Paragraph;
				// care of hyperlinks as they are not inside Run (bug reported by mdeclercq http://html2openxml.codeplex.com/workitem/11162)
				if (p != null && !p.HasChild<Run>() && !p.HasChild<Hyperlink>()) p.Remove();
				else i++;
			}

			// We add this paragraph regardless it has elements or not. A TableCell requires at least a Paragraph, as the last child of
			// of a table cell.
			// additional check for a proper cleaning (reported by antgraf http://html2openxml.codeplex.com/discussions/272744)
			if (!(cell.LastChild is Paragraph) || elements.Count > 0) cell.Append(new Paragraph(elements));

			htmlStyles.Tables.ApplyTags(cell);

			// Reset all our variables and move to next cell
			this.elements.Clear();
			String openingTag = en.CurrentTag.Replace("/", "");
			htmlStyles.Tables.EndTag(openingTag);
			htmlStyles.Runs.EndTag(openingTag);

			Point pos = tables.CellPosition;
			pos.X++;
			tables.CellPosition = pos;
		}

		#endregion
	}
}
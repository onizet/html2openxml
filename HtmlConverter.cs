using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace NotesFor.HtmlToOpenXml
{
	using a = DocumentFormat.OpenXml.Drawing;
	using pic = DocumentFormat.OpenXml.Drawing.Pictures;
	using Point = System.Drawing.Point;
	using wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
    using System.Globalization;


	/// <summary>
	/// Helper class to convert some Html text to OpenXml elements.
	/// </summary>
	public class HtmlConverter
	{
		/// <summary>
		/// Occurs when an image tag was detected and you want to manage yourself the download of the data.
		/// </summary>
		public event EventHandler<ProvisionImageEventArgs> ProvisionImage;

		sealed class CachedImagePart
		{
			public ImagePart Part;
			public Int32 Width;
			public Int32 Height;
		}

		private MainDocumentPart mainPart;

		/// <summary>The list of paragraphs that will be returned.</summary>
		private IList<OpenXmlCompositeElement> paragraphs;
		/// <summary>Holds the elements to append to the current paragraph.</summary>
		private List<OpenXmlElement> elements;
		private Paragraph currentParagraph;
		private Int32 numberingId, numberLevelRef, footnotesRef=1, endnotesRef=1;
		private Dictionary<String, Action<HtmlEnumerator>> knownTags;
		private Dictionary<Uri, CachedImagePart> knownImageParts;
		private List<String> bookmarks;
		private TableContext tables;
		private HtmlDocumentStyle htmlStyles;
		private uint drawingObjId, imageObjId = UInt32.MinValue;
		private int numberingUlId, numberingOlId = Int32.MinValue;
		private Uri baseImageUri;



		/// <summary>
		/// Constructor.
		/// </summary>
		/// <param name="maker">The mainDocumentPart of a document where to write the conversion to.</param>
		/// <remarks>We preload some configuration from inside the document such as style, bookmarks,...</remarks>
		public HtmlConverter(MainDocumentPart mainPart)
		{
			this.mainPart = mainPart;
			this.RenderPreAsTable = true;
			this.ImageProcessing = ImageProcessing.AutomaticDownload;
			knownTags = InitKnownTags();
			htmlStyles = new HtmlDocumentStyle(mainPart);
			knownImageParts = new Dictionary<Uri, CachedImagePart>();
		}

		/// <summary>
		/// Start the parse processing.
		/// </summary>
		/// <returns>Returns a list of parsed paragraph.</returns>
		public IList<OpenXmlCompositeElement> Parse(String html)
		{
			if (String.IsNullOrEmpty(html))
				return new Paragraph[0];

			// Reset:
			elements = new List<OpenXmlElement>();
			paragraphs = new List<OpenXmlCompositeElement>();
			tables = new TableContext();
			htmlStyles.Runs.Reset();
			currentParagraph = null;
			numberLevelRef = 0;

			// Start a new processing
			paragraphs.Add(currentParagraph = htmlStyles.Paragraph.NewParagraph());
			if (htmlStyles.DefaultParagraphStyle != null)
			{
				currentParagraph.Append(new ParagraphProperties(
					new ParagraphStyleId { Val = htmlStyles.DefaultParagraphStyle }
				));
			}

			HtmlEnumerator en = new HtmlEnumerator(html);
			ProcessHtmlChunks(en, null);

			if (elements.Count > 0)
				this.currentParagraph.Append(elements);

			// Remove empty paragraph unless 2 tables are side by side
			// These paragraph could be empty due to misformed html or spaces in the html source
			for (int i = 0; i < paragraphs.Count; i++)
			{
				if (paragraphs[i].HasChildren) continue;

				// If the paragraph is between 2 tables, we don't remove it (it provides some
				// separation or Word will merge the two tables)
				if (i > 0 && i + 1 < paragraphs.Count - 1
					&& paragraphs[i - 1].LocalName == "tbl"
					&& paragraphs[i + 1].LocalName == "tbl") continue;

				paragraphs.RemoveAt(i);
				i--;
			}

			return paragraphs;
		}


		#region ProcessHtmlChunks

		private void ProcessHtmlChunks(HtmlEnumerator en, String endTag)
		{
			while (en.MoveUntilMatch(endTag))
			{
				if (en.IsCurrentHtmlTag)
				{
					Action<HtmlEnumerator> action;
					if (knownTags.TryGetValue(en.CurrentTag, out action))
					{
						action(en);
					}

					// else unknow or not yet implemented - we ignore
				}
				else
				{
					// apply the previously discovered style
					Run run = new Run(
						new Text(HttpUtility.HtmlDecode(en.Current)) { Space = SpaceProcessingModeValues.Preserve }
					);
					htmlStyles.Runs.ApplyTags(run);
					elements.Add(run);
				}
			}
		}

		#endregion

		#region AlternateProcessHtmlChunks

		/// <summary>
		/// Save the actual list and restart with a new one.
		/// Continue to process until we found endTag.
		/// </summary>
		private void AlternateProcessHtmlChunks(HtmlEnumerator en, string endTag)
		{
			if(elements.Count > 0) CompleteCurrentParagraph();
			ProcessHtmlChunks(en, endTag);
		}

		#endregion

		#region AddFootnoteReference

		/// <summary>
		/// Add a note to the FootNotes part and ensure it exists.
		/// </summary>
		/// <param name="description">The description of an acronym, abbreviation, some book references, ...</param>
		/// <returns>Returns the id of the footnote reference.</returns>
		private int AddFootnoteReference(string description)
		{
			FootnotesPart fpart = mainPart.FootnotesPart;
			if (fpart == null)
				fpart = mainPart.AddNewPart<FootnotesPart>();

			if (fpart.Footnotes == null)
			{
				// Insert a new Footnotes reference
				new Footnotes(
					new Footnote(
						new Paragraph(
							new ParagraphProperties(
								new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }),
							new Run(
								new SeparatorMark())
						)
					) { Type = FootnoteEndnoteValues.Separator, Id = -1 },
					new Footnote(
						new Paragraph(
							new ParagraphProperties(
								new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }),
							new Run(
								new ContinuationSeparatorMark())
						)
					) { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 }).Save(fpart);
				footnotesRef = 1;
			}
			else
			{
				// The footnotesRef Id is a required field and should be unique. You can assign yourself some hard-coded
				// value but that's absolutely not safe. We will loop through the existing Footnote
				// to retrieve the highest Id.
				foreach (var p in fpart.Footnotes.Elements<Footnote>())
				{
					if (p.Id.HasValue && p.Id > footnotesRef) footnotesRef = (int)p.Id.Value;
				}
				footnotesRef++;
			}

			fpart.Footnotes.Append(
				new Footnote(
					new Paragraph(
						new ParagraphProperties(
							new ParagraphStyleId() { Val = htmlStyles.GetStyle("footnote text", false) }),
						new Run(
							new RunProperties(
								new RunStyle() { Val = htmlStyles.GetStyle("footnote reference", true) }),
							new FootnoteReferenceMark()),
						new Run(
							// Word insert automatically a space before the definition to separate the reference number
							// with its description
                            new Text(" " + description) { Space = SpaceProcessingModeValues.Preserve })
					)
				) { Id = footnotesRef });
			fpart.Footnotes.Save();

			return footnotesRef;
		}

		#endregion

		#region AddEndnoteReference

		/// <summary>
		/// Add a note to the Endnotes part and ensure it exists.
		/// </summary>
		/// <param name="description">The description of an acronym, abbreviation, some book references, ...</param>
		/// <returns>Returns the id of the endnote reference.</returns>
		private int AddEndnoteReference(string description)
		{
			EndnotesPart fpart = mainPart.EndnotesPart;
			if (fpart == null)
				fpart = mainPart.AddNewPart<EndnotesPart>();

			if (fpart.Endnotes == null)
			{
				// Insert a new Footnotes reference
				new Endnotes(
					new Endnote(
						new Paragraph(
							new ParagraphProperties(
								new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }),
							new Run(
								new SeparatorMark())
						)
					) { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = -1 },
					new Endnote(
						new Paragraph(
							new ParagraphProperties(
								new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }),
							new Run(
								new ContinuationSeparatorMark())
						)
					) { Id = 0 }).Save(fpart);
				endnotesRef = 1;
			}
			else
			{
				// The footnotesRef Id is a required field and should be unique. You can assign yourself some hard-coded
				// value but that's absolutely not safe. We will loop through the existing Footnote
				// to retrieve the highest Id.
				foreach (var p in fpart.Endnotes.Elements<Endnote>())
				{
					if (p.Id.HasValue && p.Id > footnotesRef) endnotesRef = (int)p.Id.Value;
				}
				endnotesRef++;
			}

			fpart.Endnotes.Append(
				new Endnote(
					new Paragraph(
						new ParagraphProperties(
							new ParagraphStyleId() { Val = htmlStyles.GetStyle("endnote text", false) }),
						new Run(
							new RunProperties(
								new RunStyle() { Val = htmlStyles.GetStyle("endnote reference", true) }),
							new FootnoteReferenceMark()),
						new Run(
							// Word insert automatically a space before the definition to separate the reference number
							// with its description
                            new Text(" " + description) { Space = SpaceProcessingModeValues.Preserve })
					)
				) { Id = endnotesRef });
			fpart.Endnotes.Save();

			return endnotesRef;
		}

		#endregion

		#region CreateImage

		private Drawing AddImagePart(Uri imageUrl, String imageSource, String alt)
		{
			if (imageObjId == UInt32.MinValue)
			{
				// In order to add images in the document, we need to asisgn an unique id
				// to each Drawing object. So we'll loop through all of the existing <wp:docPr> elements
				// to find the largest Id, then increment it for each new image.

				drawingObjId = 1; // 1 is the minimum ID set by MS Office.
				imageObjId = 1;
				foreach (var d in mainPart.Document.Body.Descendants<Drawing>())
				{
					if (d.Inline.DocProperties.Id > drawingObjId) drawingObjId = d.Inline.DocProperties.Id;

					var nvPr = d.Inline.Graphic.GraphicData.GetFirstChild<pic.NonVisualPictureProperties>();
					if(nvPr != null && nvPr.NonVisualDrawingProperties.Id > imageObjId)
						imageObjId = nvPr.NonVisualDrawingProperties.Id;
				}
				if (drawingObjId > 1) drawingObjId++;
				if (imageObjId > 1) imageObjId++;
			}


			// Cache all the ImagePart processed to avoid downloading the same image.
			CachedImagePart imagePart;
			if(!knownImageParts.TryGetValue(imageUrl, out imagePart))
			{
				ProvisionImageEventArgs e = new ProvisionImageEventArgs(imageUrl);
				if (this.ImageProcessing == ImageProcessing.AutomaticDownload && imageUrl.IsAbsoluteUri)
				{
					e.Data = ConverterUtility.DownloadData(imageUrl);
				}
				else
				{
					RaiseProvisionImage(e);
				}

				if(e.Data == null) return null;

				if (!e.ImageExtension.HasValue)
				{
					e.ImageExtension = ConverterUtility.GetImagePartTypeForImageUrl(imageUrl);
					if (!e.ImageExtension.HasValue) return null;
				}

				ImagePart ipart = mainPart.AddImagePart(e.ImageExtension.Value);
				imagePart = new CachedImagePart() { Part = ipart };

				using (Stream outputStream = ipart.GetStream(FileMode.Create))
				{
					outputStream.Write(e.Data, 0, e.Data.Length);
					outputStream.Seek(0L, SeekOrigin.Begin);

					if (e.ImageSize.IsEmpty)
					{
						e.ImageSize = ConverterUtility.GetImageSize(outputStream);						
					}

					imagePart.Width = e.ImageSize.Width;
					imagePart.Height = e.ImageSize.Height;
				}

				knownImageParts.Add(imageUrl, imagePart);
			}

			String imagePartId = mainPart.GetIdOfPart(imagePart.Part);

			/* Compute width and height in English Metrics Units.
			 * There are 360000 EMUs per centimeter, 914400 EMUs per inch, 12700 EMUs per point
			 * widthInEmus = widthInPixels / HorizontalResolutionInDPI * 914400
			 * heightInEmus = heightInPixels / VerticalResolutionInDPI * 914400
			 * 
			 * According to 1 px ~= 9525 EMU -> 914400 EMU per inch / 9525 EMU = 96 dpi
			 * So Word use 96 DPI printing which seems fair.
			 * http://hastobe.net/blogs/stevemorgan/archive/2008/09/15/howto-insert-an-image-into-a-word-document-and-display-it-using-openxml.aspx
			 */
			long widthInEmus = (long)((double)imagePart.Width / 96 * 914400L);
			long heightInEmus = (long)((double)imagePart.Height / 96 * 914400L);

			++drawingObjId;
			++imageObjId;

			var img = new Drawing(
				new wp.Inline(
					new wp.Extent() { Cx = widthInEmus, Cy = heightInEmus },
					new wp.EffectExtent() { LeftEdge = 19050L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
					new wp.DocProperties() { Id = drawingObjId, Name = imageSource, Description = String.Empty },
					new wp.NonVisualGraphicFrameDrawingProperties(
						new a.GraphicFrameLocks() { NoChangeAspect = true }),
					new a.Graphic(
						new a.GraphicData(
							new pic.Picture(
								new pic.NonVisualPictureProperties(
									new pic.NonVisualDrawingProperties() { Id = imageObjId, Name = imageSource, Description = alt },
									new pic.NonVisualPictureDrawingProperties(
										new a.PictureLocks() { NoChangeAspect = true, NoChangeArrowheads = true })),
								new pic.BlipFill(
									new a.Blip() { Embed = imagePartId },
									new a.SourceRectangle(),
									new a.Stretch(
										new a.FillRectangle())),
								new pic.ShapeProperties(
									new a.Transform2D(
										new a.Offset() { X = 0L, Y = 0L },
										new a.Extents() { Cx = widthInEmus, Cy = heightInEmus }),
									new a.PresetGeometry(
										new a.AdjustValueList()
									) { Preset = a.ShapeTypeValues.Rectangle }
								) { BlackWhiteMode = a.BlackWhiteModeValues.Auto })
						) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
				) { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U }
			);

			return img;
		}

		#endregion

		#region InitKnownTags

		private Dictionary<String, Action<HtmlEnumerator>> InitKnownTags()
		{
			// A complete list of HTML tags can be found here: http://www.w3schools.com/tags/default.asp

			var knownTags = new Dictionary<String, Action<HtmlEnumerator>>(StringComparer.InvariantCultureIgnoreCase) {
				{ "<a>", ProcessLink },
				{ "<abbr>" , ProcessAcronym },
				{ "<acronym>" , ProcessAcronym },
				{ "<b>", ProcessBold },
				{ "<body>", ProcessBody },
				{ "<br>", ProcessBr },
				{ "<caption>", ProcessTableCaption },
				{ "<cite>", ProcessCite },
				{ "<del>", ProcessStrike },
				{ "<div>", ProcessDiv },
				{ "<dd>", ProcessDefinitionListItem },
				{ "<dt>", ProcessDefinitionList },
				{ "<em>", ProcessItalic },
				{ "<font>", ProcessFont },
				{ "<h1>", ProcessHeading },
				{ "<h2>", ProcessHeading },
				{ "<h3>", ProcessHeading },
				{ "<h4>", ProcessHeading },
				{ "<h5>", ProcessHeading },
				{ "<h6>", ProcessHeading },
				{ "<hr>", ProcessHorizontalLine },
				{ "<i>", ProcessItalic },
				{ "<img>", ProcessImage },
				{ "<ins>", ProcessUnderline },
				{ "<li>", ProcessLi },
				{ "<ol>", ProcessNumberingList },
				{ "<p>", ProcessParagraph },
				{ "<pre>", ProcessPre },
				{ "<span>", ProcessSpan },
				{ "<s>", ProcessStrike },
				{ "<strike>", ProcessStrike },
				{ "<strong>", ProcessBold },
				{ "<sub>", ProcessSubscript },
				{ "<sup>", ProcessSuperscript },
				{ "<table>", ProcessTable },
				{ "<tbody>", ProcessTablePart },
				{ "<td>", ProcessTableColumn },
				{ "<tfoot>", ProcessTablePart },
				{ "<th>", ProcessTableColumn },
				{ "<thead>", ProcessTablePart },
				{ "<tr>", ProcessTableRow },
				{ "<u>", ProcessUnderline },
				{ "<ul>", ProcessNumberingList },

				// closing tag
				{ "</b>", ProcessClosingBold },
				{ "</body>", ProcessClosingTag },
				{ "</cite>", ProcessClosingTag },
				{ "</del>", ProcessClosingTag },
				{ "</div>", ProcessClosingDiv },
				{ "</em>", ProcessClosingItalic },
				{ "</font>", ProcessClosingTag },
				{ "</i>", ProcessClosingItalic },
				{ "</ins>", ProcessClosingTag },
				{ "</p>", ProcessClosingParagraph },
				{ "</span>", ProcessClosingTag },
				{ "</s>", ProcessClosingTag },
				{ "</strike>", ProcessClosingTag },
				{ "</strong>", ProcessClosingBold },
				{ "</ol>", ProcessClosingNumberingList },
				{ "</sub>", ProcessClosingTag },
				{ "</sup>", ProcessClosingTag },
				{ "</table>", ProcessClosingTable },
				{ "</tbody>", ProcessClosingTablePart },
				{ "</tfoot>", ProcessClosingTablePart },
				{ "</thead>", ProcessClosingTablePart },
				{ "</td>", ProcessClosingTableColumn },
				{ "</tr>", ProcessClosingTableRow },
				{ "</u>", ProcessClosingTag },
				{ "</ul>", ProcessClosingNumberingList },
			};

			return knownTags;
		}

		#endregion

		#region Bookmarks

		private List<String> Bookmarks
		{
			get
			{
				if (bookmarks == null)
				{
					bookmarks = new List<String>();
					var en = mainPart.Document.Body.Descendants<BookmarkStart>().GetEnumerator();
					while (en.MoveNext())
						bookmarks.Add(en.Current.Name.Value);
					bookmarks.Sort(StringComparer.Ordinal);
				}
				return bookmarks;
			}
		}

		#endregion

		#region CompleteCurrentParagraph

		/// <summary>
		/// Push the elements members to the current paragraph and reset the elements collection.
		/// </summary>
		private void CompleteCurrentParagraph()
		{
			this.currentParagraph.Append(elements);
			htmlStyles.Paragraph.ApplyTags(currentParagraph);
			elements.Clear();
		}

		#endregion

		#region RefreshStyle

		/// <summary>
		/// Refresh the cache of styles presents in the document.
		/// </summary>
		public void RefreshStyles()
		{
			htmlStyles.PrepareStyles(mainPart);
		}

		#endregion

		#region ProcessContainerAttributes

		/// <summary>
		/// There is a few attributes shared by a large number of tags. This method will check them for a limited
		/// number of tags (&lt;p&gt;, &lt;pre&gt;, &lt;div&gt;, &lt;span&gt; and &lt;body&gt;,).
		/// </summary>
		private bool ProcessContainerAttributes(HtmlEnumerator en, IList<OpenXmlElement> styleAttributes)
		{
			if (en.Attributes.Count == 0) return false;

			bool paragraphPropertiesChanged = false;
			List<OpenXmlElement> containerStyleAttributes = new List<OpenXmlElement>();

			string attrValue = en.Attributes["lang"];
			if (attrValue != null && attrValue.Length > 0)
			{
				containerStyleAttributes.Add(
					new ParagraphMarkRunProperties(
						new Languages() { Val = attrValue }));
			}


			// Not applicable to a table : page break
			if (!tables.HasContext)
			{
				attrValue = en.StyleAttributes["page-break-after"];
				if (attrValue == "always")
				{
					paragraphs.Add(new Paragraph(
						new Run(
							new Break() { Type = BreakValues.Page })));
					paragraphPropertiesChanged = true;
				}

				attrValue = en.StyleAttributes["page-break-before"];
				if (attrValue == "always")
				{
					paragraphs.Insert(paragraphs.Count - 1, new Paragraph(
						new Run(
							new Break() { Type = BreakValues.Page })));
					paragraphPropertiesChanged = true;
				}
			}

			attrValue = en.StyleAttributes["text-align"];
			if (attrValue != null && en.CurrentTag != "<font>")
			{
				JustificationValues? align = ConverterUtility.FormatParagraphAlign(attrValue);
				if (align.HasValue)
				{
					containerStyleAttributes.Add(new Justification { Val = align });
				}
			}

			htmlStyles.Paragraph.BeginTag(en.CurrentTag, containerStyleAttributes);

			// Process general run styles
			htmlStyles.Runs.ProcessCommonRunAttributes(en, styleAttributes);

			return paragraphPropertiesChanged;
		}

		#endregion

		#region EnsureNumberingIds

		private void EnsureNumberingIds()
		{
			if (numberingOlId != Int32.MinValue) return;

			// Ensure the numbering.xml file exists or any numbering or bullets list will results
			// in simple numbering list (1.   2.   3...)
			if (mainPart.NumberingDefinitionsPart == null || mainPart.NumberingDefinitionsPart.Numbering == null)
			{
				// This minimal numbering definition has been inspired by the documentation OfficeXMLMarkupExplained_en.docx
				// http://www.microsoft.com/downloads/details.aspx?FamilyID=6f264d0b-23e8-43fe-9f82-9ab627e5eaa3&displaylang=en

				NumberingDefinitionsPart numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
				new Numbering(
					new AbstractNum(
						new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
						new Level(
							new NumberingFormat() { Val = NumberFormatValues.Bullet },
							new LevelText() { Val = "•" },
							new PreviousParagraphProperties(
								new Indentation() { Left = "420", Hanging = "360" })
						) { LevelIndex = 0 }
					) { AbstractNumberId = 0 },
					new AbstractNum(
						new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
						new Level(
							new StartNumberingValue() { Val = 1 },
							new NumberingFormat() { Val = NumberFormatValues.Decimal },
							new LevelText() { Val = "%1." },
							new PreviousParagraphProperties(
								new Indentation() { Left = "420", Hanging = "360" })
						) { LevelIndex = 0 }
					) { AbstractNumberId = 1 },
                    new NumberingInstance(
						new AbstractNumId() { Val = 0 }
					) { NumberID = 1 },
                    new NumberingInstance(
						new AbstractNumId() { Val = 1 }
					) { NumberID = 2 }).Save(numberingPart);

				numberingUlId = 1;
				numberingOlId = 2;
			}
			else
			{
				// else find back the numbering for (un)ordered list

				Numbering numbering = mainPart.NumberingDefinitionsPart.Numbering;
				numberingOlId = numbering.FindNumberIDByFormat(NumberFormatValues.Decimal);
				numberingUlId = numbering.FindNumberIDByFormat(NumberFormatValues.Bullet);
			}
		}

		#endregion

		// Events

		#region RaiseProvisionImage

		/// <summary>
		/// Raises the ProvisionImage event.
		/// </summary>
		protected virtual void RaiseProvisionImage(ProvisionImageEventArgs e)
		{
			if (ProvisionImage != null) ProvisionImage(this, e);
		}

		#endregion

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
				 new SpacingBetweenLines(){ After = "0" });
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
				   new SpacingBetweenLines(){ After = "0" }
			);

			// Restore the original elements list
			this.paragraphs.Add(currentParagraph);
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
				bool paraPropsWasChanged = ProcessContainerAttributes(en, runStyleAttributes);

				if(runStyleAttributes.Count > 0)
					htmlStyles.Runs.BeginTag(en.CurrentTag, runStyleAttributes);

				// Any changes that requires a new paragraph?
				if (!paraPropsWasChanged && previousParagraph.HasChild<Run>())
				{
					ProcessBr(en);
					currentParagraph = previousParagraph;
				}
				else
				{
					this.paragraphs.Add(currentParagraph);
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

			if(styleAttributes.Count > 0)
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
			this.paragraphs.Add(p);
			this.paragraphs.Add(currentParagraph = htmlStyles.Paragraph.NewParagraph());
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
			this.paragraphs.Add(currentParagraph = htmlStyles.Paragraph.NewParagraph());
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

					if (!uri.IsAbsoluteUri && this.BaseImageUrl != null)
						uri = new Uri(this.BaseImageUrl, uri);

					drawing = AddImagePart(uri, src, alt);
				}
			}

			if (!en.IsSelfClosedTag) AlternateProcessHtmlChunks(en, "</img>");

			if(drawing != null)
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
			// Continue to process the html until we found </li>
			AlternateProcessHtmlChunks(en, "</li>");

			currentParagraph = htmlStyles.Paragraph.NewParagraph();
			currentParagraph.Append(elements);
			currentParagraph.InsertInProperties(
				new SpacingBetweenLines() { After = "0" },
				new NumberingProperties(
					new NumberingLevelReference() { Val = numberLevelRef - 1 },
					new NumberingId() { Val = numberingId }
				)
			);

			// Restore the original elements list
			this.paragraphs.Add(currentParagraph);
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
					this.paragraphs.Add(currentParagraph = htmlStyles.Paragraph.NewParagraph());
				}
				else
				{
					// Append the processed elements and put them to the Run of the Hyperlink
					h.Append(elements);

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
			this.paragraphs.Add(currentParagraph = htmlStyles.Paragraph.NewParagraph());

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
			ProcessContainerAttributes(en, styleAttributes);

			if(styleAttributes.Count > 0)
				htmlStyles.Runs.BeginTag(en.CurrentTag, styleAttributes.ToArray());
		}

		#endregion

		#region ProcessPre

		private void ProcessPre(HtmlEnumerator en)
		{
			en.PreserveWhitespaces = true;

			CompleteCurrentParagraph();
			currentParagraph = htmlStyles.Paragraph.NewParagraph();

			// Oftenly, <pre> tag are used to renders some code examples. They look better inside a table
			if (this.RenderPreAsTable)
			{
				Table currentTable = new Table(
					new TableProperties(
						new TableStyle() { Val = htmlStyles.GetStyle("Table Grid", false) },
						new TableWidth() { Type = TableWidthUnitValues.Pct, Width = "5000" }), // 100% * 500
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

				this.paragraphs.Add(currentTable);
				tables.NewContext(currentTable);
			}
			else
			{
				this.paragraphs.Add(currentParagraph);
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
			this.paragraphs.Add(currentParagraph = htmlStyles.Paragraph.NewParagraph());
			en.PreserveWhitespaces = false;
		}

		#endregion

		#region ProcessSpan

		private void ProcessSpan(HtmlEnumerator en)
		{
			// A span style attribute can contains many information: font color, background color, font size,
			// font family, ...
			// We'll check for each of these and add apply them to the next build runs.

			List<OpenXmlElement> styleAttributes = new List<OpenXmlElement>();
			ProcessContainerAttributes(en, styleAttributes);

			if (styleAttributes.Count > 0)
				htmlStyles.Runs.MergeTag("<span>", styleAttributes);
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
			if (border.HasValue)
			{
				// If the border has been specified, we display the Table Grid style which display
				// its grid lines. Otherwise the default table style hides the grid lines.
				properties.Add(new TableStyle() { Val = htmlStyles.GetStyle("Table Grid", false) });
			}

			Unit unit = en.StyleAttributes.GetAsUnit("width");
			if (!unit.IsValid) unit = en.Attributes.GetAsUnit("width");

			if (unit.IsValid)
			{
				switch (unit.Type)
				{
					case "%":
                        properties.Add(new TableWidth() { Type = TableWidthUnitValues.Pct, Width = (unit.Value * 500).ToString(CultureInfo.InvariantCulture) }); break;
					case "pt":
                        properties.Add(new TableWidth() { Type = TableWidthUnitValues.Dxa, Width = (unit.Value * 20).ToString(CultureInfo.InvariantCulture) }); break;
					case "px":
                        properties.Add(new TableWidth() { Type = TableWidthUnitValues.Dxa, Width = (unit.Value).ToString(CultureInfo.InvariantCulture) }); break;
				}
			}
			else
			{
				properties.Add(new TableWidth() { Type = TableWidthUnitValues.Pct, Width = "5000" }); // 100% * 500
			}

			string align = en.Attributes["align"];
			if (align != null)
			{
				JustificationValues? halign = ConverterUtility.FormatParagraphAlign(align);
				if(halign.HasValue)
					properties.Add(new TableJustification(){ Val = halign.Value.ToTableRowAlignment() });
			}

			List<OpenXmlElement> runStyleAttributes = new List<OpenXmlElement>();
			htmlStyles.Runs.ProcessCommonRunAttributes(en, runStyleAttributes);
			if(runStyleAttributes.Count > 0)
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
				this.paragraphs.Insert(this.paragraphs.Count-1, legend);
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
			htmlStyles.Tables.ProcessCommonAttributes(en, styleAttributes);

			TableRow row = new TableRow();
			if(styleAttributes.Count > 0)
				row.Append(new TableRowProperties(styleAttributes));

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
			if(!unit.IsValid) unit = en.Attributes.GetAsUnit("width");

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
				styleAttributes.Add(new GridSpan() { Val =  colspan });
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
			if(runStyleAttributes.Count > 0)
				htmlStyles.Runs.BeginTag("<td>", runStyleAttributes.ToArray());

			TableCell cell = new TableCell(
				new TableCellProperties(styleAttributes));
			tables.CurrentTable.GetLastChild<TableRow>().Append(cell);

			if (en.IsSelfClosedTag) // Force a call to ProcessClosingTableColumn
				ProcessClosingTableColumn(en);
		}

		#endregion

		#region ProcessTablePart

		private void ProcessTablePart(HtmlEnumerator en)
		{
			List<OpenXmlElement> styleAttributes = new List<OpenXmlElement>();

			htmlStyles.Tables.ProcessCommonAttributes(en, styleAttributes);

			if(styleAttributes.Count > 0)
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
			// one for containing all the <li>
			if (numberLevelRef == 0)
				this.paragraphs.Add(currentParagraph = htmlStyles.Paragraph.NewParagraph());
		}

		#endregion

		#region ProcessClosingParagraph

		private void ProcessClosingParagraph(HtmlEnumerator en)
		{
			CompleteCurrentParagraph();
			this.paragraphs.Add(currentParagraph = htmlStyles.Paragraph.NewParagraph());

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
				if(row.ChildElements[i] is TableCell)
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

			if(!tables.HasContext)
				this.paragraphs.Add(currentParagraph = htmlStyles.Paragraph.NewParagraph());
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
		}

		#endregion

		#region ProcessClosingTableColumn

		private void ProcessClosingTableColumn(HtmlEnumerator en)
		{
			TableCell cell = tables.CurrentTable.GetLastChild<TableRow>().GetLastChild<TableCell>();
			Paragraph p = htmlStyles.Paragraph.NewParagraph();
			p.Append(elements);
			cell.Append(p);

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

		//____________________________________________________________________
		//
		// Configuration

		/// <summary>
		/// Gets or sets where to render the acronym or abbreviation tag.
		/// </summary>
		public AcronymPosition AcronymPosition { get; set; }

		/// <summary>
		/// Gets or sets whether the &lt;div&gt; tag should be processed as &lt;p&gt;. It depends whether you consider &lt;div&gt;
		/// as part of the layout or as part of a text field.
		/// </summary>
		public bool ConsiderDivAsParagraph { get; set; }

		/// <summary>
		/// Gets or sets whether anchor links are included or not in the conversion.
		/// </summary>
		/// <remarks>An anchor is a term used to define a hyperlink destination inside a document.
		/// <see cref="http://www.w3schools.com/HTML/html_links.asp"/>.
		/// <br/>
		/// It exists some predefined anchors used by Word such as _top to refer to the top of the document.
		/// The anchor <i>#_top</i> is always accepted regardless this property value.
		/// For others anchors like refering to your own bookmark or a title, add a 
		/// <see cref="DocumentFormat.OpenXml.Wordprocessing.BookmarkStart"/> and 
		/// <see cref="DocumentFormat.OpenXml.Wordprocessing.BookmarkEnd"/> elements
		/// and set the value of href to <i>#&lt;name of your bookmark&gt;</i>.
		/// </remarks>
		public bool ExcludeLinkAnchor { get; set; }

		/// <summary>
		/// Gets the Html styles manager mapping to OpenXml style properties.
		/// </summary>
		public HtmlDocumentStyle HtmlStyles
		{
			get { return htmlStyles; }
		}

		/// <summary>
		/// Gets or sets how the &lt;img&gt; tag should be handled.
		/// </summary>
		public ImageProcessing ImageProcessing { get; set; }

		/// <summary>
		/// Gets or sets the base Uri used to automaticaly resolve relative images 
		/// if used with ImageProcessing = AutomaticDownload.
		/// </summary>
		public Uri BaseImageUrl
		{
			get { return this.baseImageUri; }
			set
			{
				if (value != null && !value.IsAbsoluteUri)
					throw new ArgumentException("BaseImageUrl should be an absolute Uri");
				this.baseImageUri = value;
			}
		}

		/// <summary>
		/// Gets or sets where the Legend tag (&lt;caption&gt;) should be rendered (above or below the table).
		/// </summary>
		public CaptionPositionValues TableCaptionPosition { get; set; }

		/// <summary>
		/// Gets or sets whether the &lt;pre&gt; tag should be rendered as a table.
		/// </summary>
		/// <remarks>The table will contains only one cell.</remarks>
		public bool RenderPreAsTable { get; set; }
	}
}

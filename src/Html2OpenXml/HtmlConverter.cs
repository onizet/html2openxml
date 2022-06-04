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
using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml.IO;

namespace HtmlToOpenXml
{
    using a = DocumentFormat.OpenXml.Drawing;
    using pic = DocumentFormat.OpenXml.Drawing.Pictures;
    using wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;



	/// <summary>
	/// Helper class to convert some Html text to OpenXml elements.
	/// </summary>
	public partial class HtmlConverter
	{
		private MainDocumentPart _mainPart;
		/// <summary>The list of paragraphs that will be returned.</summary>
		private IList<OpenXmlCompositeElement> _paragraphs;
		/// <summary>Holds the elements to append to the current paragraph.</summary>
		private List<OpenXmlElement> _elements;
		private Paragraph _currentParagraph;
		private Int32 _footnotesRef = 1, _endnotesRef = 1, _figCaptionRef = -1;
		private Dictionary<String, Action<HtmlEnumerator>> _knownTags;
        private ImagePrefetcher _imagePrefetcher;
        private TableContext _tables;
        private readonly HtmlDocumentStyle _htmlStyles;
        private readonly IWebRequest _webRequester;
        private uint drawingObjId, _imageObjId;



		/// <summary>
		/// Constructor.
		/// </summary>
		/// <param name="mainPart">The mainDocumentPart of a document where to write the conversion to.</param>
		/// <remarks>We preload some configuration from inside the document such as style, bookmarks,...</remarks>
        public HtmlConverter(MainDocumentPart mainPart) : this(mainPart, null)
        {
        }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="mainPart">The mainDocumentPart of a document where to write the conversion to.</param>
        /// <param name="webRequester">Factory to download the images.</param>
        /// <remarks>We preload some configuration from inside the document such as style, bookmarks,...</remarks>
        public HtmlConverter(MainDocumentPart mainPart, IWebRequest webRequester = null)
        {
            this._knownTags = InitKnownTags();
            this._mainPart = mainPart ?? throw new ArgumentNullException("mainPart");
            this._htmlStyles = new HtmlDocumentStyle(mainPart);
            this._webRequester = webRequester ?? new DefaultWebRequest();
        }

		/// <summary>
		/// Start the parse processing.
		/// </summary>
		/// <returns>Returns a list of parsed paragraph.</returns>
        public IList<OpenXmlCompositeElement> Parse(String html)
		{
			if (String.IsNullOrEmpty(html))
				return new Paragraph[0];

			// ensure a body exists to avoid any errors when trying to access it
			if (_mainPart.Document == null)
				new Document(new Body()).Save(_mainPart);
			else if (_mainPart.Document.Body == null)
				_mainPart.Document.Body = new Body();

			// Reset:
			_elements = new List<OpenXmlElement>();
			_paragraphs = new List<OpenXmlCompositeElement>();
			_tables = new TableContext();
			_htmlStyles.Runs.Reset();
			_currentParagraph = null;

			// Start a new processing
			_paragraphs.Add(_currentParagraph = _htmlStyles.Paragraph.NewParagraph());
			if (_htmlStyles.DefaultStyles.ParagraphStyle != null)
			{
				_currentParagraph.ParagraphProperties = new ParagraphProperties {
					ParagraphStyleId = new ParagraphStyleId { Val = _htmlStyles.DefaultStyles.ParagraphStyle }
				};
			}

			HtmlEnumerator en = new HtmlEnumerator(html);
			ProcessHtmlChunks(en, null);

            if (_elements.Count > 0)
                this._currentParagraph.Append(_elements);

			// As the Parse method is public, to avoid changing the type of the return value, I use this proxy
			// that will allow me to call the recursive method RemoveEmptyParagraphs with no major changes, impacting the client.
			RemoveEmptyParagraphs();

			return _paragraphs;
		}

        /// <summary>
		/// Start the parse processing and append the converted paragraphs into the Body of the document.
		/// </summary>
        public void ParseHtml(String html)
        {
            // This method exists because we may ensure the SectionProperties remains the last element of the body.
            // It's mandatory when dealing with page orientation

            var paragraphs = Parse(html);

			Body body = _mainPart.Document.Body;
			SectionProperties sectionProperties = body.GetLastChild<SectionProperties>();
			for (int i = 0; i < paragraphs.Count; i++)
				body.Append(paragraphs[i]);

			// move the paragraph with BookmarkStart `_GoBack` as the last child
			var p = body.GetFirstChild<Paragraph>();
			if (p != null && p.HasChild<BookmarkStart>())
			{
				p.Remove();
				body.Append(p);
			}

			// Push the sectionProperties as the last element of the Body
			// (required by OpenXml schema to avoid the bad formatting of the document)
			if (sectionProperties != null)
			{
				sectionProperties.Remove();
				body.AddChild(sectionProperties);
			}
		}

		#region RemoveEmptyParagraphs

		/// <summary>
		/// Remove empty paragraph unless 2 tables are side by side.
		/// These paragraph could be empty due to misformed html or spaces in the html source.
		/// </summary>
		private void RemoveEmptyParagraphs()
		{
			bool hasRuns;

			for (int i = 0; i < _paragraphs.Count; i++)
			{
				OpenXmlCompositeElement p = _paragraphs[i];

				// If the paragraph is between 2 tables, we don't remove it (it provides some
				// separation or Word will merge the two tables)
				if (i > 0 && i + 1 < _paragraphs.Count - 1
					&& _paragraphs[i - 1].LocalName == "tbl"
					&& _paragraphs[i + 1].LocalName == "tbl") continue;

				if (p.HasChildren)
				{
					if (!(p is Paragraph)) continue;

					// Has this paragraph some other elements than ParagraphProperties?
					// This code ensure no default style or attribute on empty div will stay
					hasRuns = false;
					for (int j = p.ChildElements.Count - 1; j >= 0; j--)
					{
						ParagraphProperties prop = p.ChildElements[j] as ParagraphProperties;
						if (prop == null || prop.SectionProperties != null)
						{
							hasRuns = true;
							break;
						}
					}

					if (hasRuns) continue;
				}

				_paragraphs.RemoveAt(i);
				i--;
			}
		}

		#endregion

		#region ProcessHtmlChunks

		private void ProcessHtmlChunks(HtmlEnumerator en, String endTag)
		{
			while (en.MoveUntilMatch(endTag))
			{
				if (en.IsCurrentHtmlTag)
				{
					Action<HtmlEnumerator> action;
					if (_knownTags.TryGetValue(en.CurrentTag, out action))
					{
						if (Logging.On) Logging.PrintVerbose(en.Current);
						action(en);
					}

					// else unknown or not yet implemented - we ignore
				}
				else
				{
					Run run = new Run(
						new Text(HttpUtility.HtmlDecode(en.Current)) { Space = SpaceProcessingModeValues.Preserve }
					);
					// apply the previously discovered style
					_htmlStyles.Runs.ApplyTags(run);
					_elements.Add(run);
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
			if (_elements.Count > 0) CompleteCurrentParagraph();
			ProcessHtmlChunks(en, endTag);
		}

		#endregion

		#region AddParagraph

		/// <summary>
		/// Add a new paragraph, table, ... to the list of processed paragrahs. This method takes care of 
		/// adding the new element to the current table if it exists.
		/// </summary>
		private void AddParagraph(OpenXmlCompositeElement element)
		{
			if (_tables.HasContext)
			{
				TableRow row = _tables.CurrentTable.GetLastChild<TableRow>();
				if (row == null)
				{
					_tables.CurrentTable.Append(row = new TableRow());
					_tables.CellPosition = new CellPosition(_tables.CellPosition.Row + 1, 0);
				}
                TableCell cell = row.GetLastChild<TableCell>();
                if (cell == null) // ensure cell exists (issue #13982 reported by Willu)
                {
                    row.Append(cell = new TableCell());
                }
                cell.Append(element);
			}
			else
				this._paragraphs.Add(element);
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
			FootnotesPart fpart = _mainPart.FootnotesPart;
			if (fpart == null)
				fpart = _mainPart.AddNewPart<FootnotesPart>();

			if (fpart.Footnotes == null)
			{
				// Insert a new Footnotes reference
				new Footnotes(
					new Footnote(
						new Paragraph(
							new ParagraphProperties {
								SpacingBetweenLines = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }
							},
							new Run(
								new SeparatorMark())
						)
					) { Type = FootnoteEndnoteValues.Separator, Id = -1 },
					new Footnote(
						new Paragraph(
							new ParagraphProperties {
								SpacingBetweenLines = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }
							},
							new Run(
								new ContinuationSeparatorMark())
						)
					) { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 }).Save(fpart);
				_footnotesRef = 1;
			}
			else
			{
				// The footnotesRef Id is a required field and should be unique. You can assign yourself some hard-coded
				// value but that's absolutely not safe. We will loop through the existing Footnote
				// to retrieve the highest Id.
				foreach (var fn in fpart.Footnotes.Elements<Footnote>())
				{
					if (fn.Id.HasValue && fn.Id > _footnotesRef) _footnotesRef = (int) fn.Id.Value;
				}
				_footnotesRef++;
			}


            Paragraph p;
			fpart.Footnotes.Append(
				new Footnote(
					p = new Paragraph(
						new ParagraphProperties {
							ParagraphStyleId = new ParagraphStyleId() { Val = _htmlStyles.GetStyle(_htmlStyles.DefaultStyles.FootnoteTextStyle, StyleValues.Paragraph) }
						},
						new Run(
							new RunProperties {
								RunStyle = new RunStyle() { Val = _htmlStyles.GetStyle(_htmlStyles.DefaultStyles.FootnoteReferenceStyle, StyleValues.Character) }
							},
							new FootnoteReferenceMark()),
						new Run(
				        // Word insert automatically a space before the definition to separate the
                        // reference number with its description
							new Text(" ") { Space = SpaceProcessingModeValues.Preserve })
					)
				) { Id = _footnotesRef });


            // Description in footnote reference can be plain text or a web protocols/file share (like \\server01)
            Uri uriReference;
            Regex linkRegex = new Regex(@"^((https?|ftps?|mailto|file)://|[\\]{2})(?:[\w][\w.-]?)");
            if (linkRegex.IsMatch(description) && Uri.TryCreate(description, UriKind.Absolute, out uriReference))
            {
                // when URI references a network server (ex: \\server01), System.IO.Packaging is not resolving the correct URI and this leads
                // to a bad-formed XML not recognized by Word. To enforce the "original URI", a fresh new instance must be created
                uriReference = new Uri(uriReference.AbsoluteUri, UriKind.Absolute);
                HyperlinkRelationship extLink = fpart.AddHyperlinkRelationship(uriReference, true);
                var h = new Hyperlink(
                    ) { History = true, Id = extLink.Id };

                h.Append(new Run(
                    new RunProperties {
                        RunStyle = new RunStyle() { Val = _htmlStyles.GetStyle(_htmlStyles.DefaultStyles.HyperlinkStyle, StyleValues.Character) }
                    },
                    new Text(description)));
                p.Append(h);
            }
            else
            {
                p.Append(new Run(
                    new Text(description) { Space = SpaceProcessingModeValues.Preserve }));
            }

			fpart.Footnotes.Save();

			return _footnotesRef;
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
			EndnotesPart fpart = _mainPart.EndnotesPart;
			if (fpart == null)
				fpart = _mainPart.AddNewPart<EndnotesPart>();

			if (fpart.Endnotes == null)
			{
				// Insert a new Footnotes reference
				new Endnotes(
					new Endnote(
						new Paragraph(
							new ParagraphProperties {
								SpacingBetweenLines = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }
							},
							new Run(
								new SeparatorMark())
						)
					) { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = -1 },
					new Endnote(
						new Paragraph(
							new ParagraphProperties {
								SpacingBetweenLines = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }
							},
							new Run(
								new ContinuationSeparatorMark())
						)
					) { Id = 0 }).Save(fpart);
				_endnotesRef = 1;
			}
			else
			{
				// The footnotesRef Id is a required field and should be unique. You can assign yourself some hard-coded
				// value but that's absolutely not safe. We will loop through the existing Footnote
				// to retrieve the highest Id.
				foreach (var p in fpart.Endnotes.Elements<Endnote>())
				{
					if (p.Id.HasValue && p.Id > _footnotesRef) _endnotesRef = (int) p.Id.Value;
				}
				_endnotesRef++;
			}

			fpart.Endnotes.Append(
				new Endnote(
					new Paragraph(
						new ParagraphProperties {
							ParagraphStyleId = new ParagraphStyleId() { Val = _htmlStyles.GetStyle(_htmlStyles.DefaultStyles.EndnoteTextStyle, StyleValues.Paragraph) }
						},
						new Run(
							new RunProperties {
								RunStyle = new RunStyle() { Val = _htmlStyles.GetStyle(_htmlStyles.DefaultStyles.EndnoteReferenceStyle, StyleValues.Character) }
							},
							new FootnoteReferenceMark()),
						new Run(
				// Word insert automatically a space before the definition to separate the reference number
				// with its description
							new Text(" " + description) { Space = SpaceProcessingModeValues.Preserve })
					)
				) { Id = _endnotesRef });

			fpart.Endnotes.Save();

			return _endnotesRef;
		}

		#endregion

		#region AddFigureCaption

		/// <summary>
		/// Add a new figure caption to the document.
		/// </summary>
		/// <returns>Returns the id of the new figure caption.</returns>
		private int AddFigureCaption()
		{
			if (_figCaptionRef == -1)
			{
				_figCaptionRef = 0;
				foreach (var p in _mainPart.Document.Descendants<SimpleField>())
				{
					if (p.Instruction == " SEQ Figure \\* ARABIC ")
						_figCaptionRef++;
				}
			}
			_figCaptionRef++;
			return _figCaptionRef;
		}

		#endregion

		#region AddImagePart

		private Drawing AddImagePart(String imageSource, String alt, Size preferredSize)
		{
			if (_imageObjId == UInt32.MinValue)
			{
				// In order to add images in the document, we need to asisgn an unique id
				// to each Drawing object. So we'll loop through all of the existing <wp:docPr> elements
				// to find the largest Id, then increment it for each new image.

				drawingObjId = 1; // 1 is the minimum ID set by MS Office.
				_imageObjId = 1;
				foreach (var d in _mainPart.Document.Body.Descendants<Drawing>())
				{
					if (d.Inline == null) continue; // fix some rare issue where Inline is null (reported by scwebgroup)
					if (d.Inline.DocProperties.Id > drawingObjId) drawingObjId = d.Inline.DocProperties.Id;

					var nvPr = d.Inline.Graphic.GraphicData.GetFirstChild<pic.NonVisualPictureProperties>();
					if (nvPr != null && nvPr.NonVisualDrawingProperties.Id > _imageObjId)
						_imageObjId = nvPr.NonVisualDrawingProperties.Id;
				}
				if (drawingObjId > 1) drawingObjId++;
				if (_imageObjId > 1) _imageObjId++;
			}

            // Cache all the ImagePart processed to avoid downloading the same image.
            if (_imagePrefetcher == null)
                _imagePrefetcher = new ImagePrefetcher(_mainPart, _webRequester);

            HtmlImageInfo iinfo = _imagePrefetcher.Download(imageSource);

            if (iinfo == null)
                return null;

			if (preferredSize.IsEmpty)
			{
				preferredSize = iinfo.Size;
			}
			else if (preferredSize.Width <= 0 || preferredSize.Height <= 0)
			{
				Size actualSize = iinfo.Size;
				preferredSize = ImageHeader.KeepAspectRatio(actualSize, preferredSize);
			}

			long widthInEmus = new Unit(UnitMetric.Pixel, preferredSize.Width).ValueInEmus;
			long heightInEmus = new Unit(UnitMetric.Pixel, preferredSize.Height).ValueInEmus;

			++drawingObjId;
			++_imageObjId;

			var img = new Drawing(
				new wp.Inline(
					new wp.Extent() { Cx = widthInEmus, Cy = heightInEmus },
					new wp.EffectExtent() { LeftEdge = 19050L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
					new wp.DocProperties() { Id = drawingObjId, Name = "Picture " + _imageObjId, Description = String.Empty },
					new wp.NonVisualGraphicFrameDrawingProperties {
						GraphicFrameLocks = new a.GraphicFrameLocks() { NoChangeAspect = true }
					},
					new a.Graphic(
						new a.GraphicData(
							new pic.Picture(
								new pic.NonVisualPictureProperties {
									NonVisualDrawingProperties = new pic.NonVisualDrawingProperties() { Id = _imageObjId, Name = DataUri.IsWellFormed(imageSource) ? string.Empty : imageSource, Description = alt },
									NonVisualPictureDrawingProperties = new pic.NonVisualPictureDrawingProperties(
										new a.PictureLocks() { NoChangeAspect = true, NoChangeArrowheads = true })
								},
								new pic.BlipFill(
									new a.Blip() { Embed = iinfo.ImagePartId },
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
				) { DistanceFromTop = (UInt32Value) 0U, DistanceFromBottom = (UInt32Value) 0U, DistanceFromLeft = (UInt32Value) 0U, DistanceFromRight = (UInt32Value) 0U }
			);

			return img;
		}

		#endregion

		#region InitKnownTags

		private Dictionary<String, Action<HtmlEnumerator>> InitKnownTags()
		{
			// A complete list of HTML tags can be found here: http://www.w3schools.com/tags/default.asp

			var knownTags = new Dictionary<String, Action<HtmlEnumerator>>(StringComparer.OrdinalIgnoreCase) {
				{ "<a>", ProcessLink },
				{ "<abbr>", ProcessAcronym },
				{ "<acronym>", ProcessAcronym },
                { "<article>", ProcessDiv },
                { "<aside>", ProcessDiv },
				{ "<b>", ProcessHtmlElement<Bold> },
                { "<blockquote>", ProcessBlockQuote },
				{ "<body>", ProcessBody },
				{ "<br>", ProcessBr },
				{ "<caption>", ProcessTableCaption },
				{ "<cite>", ProcessCite },
				{ "<del>", ProcessHtmlElement<Strike> },
				{ "<div>", ProcessDiv },
				{ "<dd>", ProcessDefinitionListItem },
				{ "<dt>", ProcessDefinitionList },
				{ "<em>", ProcessHtmlElement<Italic> },
				{ "<font>", ProcessFont },
				{ "<h1>", ProcessHeading },
				{ "<h2>", ProcessHeading },
				{ "<h3>", ProcessHeading },
				{ "<h4>", ProcessHeading },
				{ "<h5>", ProcessHeading },
				{ "<h6>", ProcessHeading },
				{ "<hr>", ProcessHorizontalLine },
                { "<html>", ProcessHtml },
                { "<figcaption>", ProcessFigureCaption },
				{ "<i>", ProcessHtmlElement<Italic> },
				{ "<img>", ProcessImage },
				{ "<ins>", ProcessUnderline },
				{ "<li>", ProcessLi },
				{ "<ol>", ProcessNumberingList },
				{ "<p>", ProcessParagraph },
				{ "<pre>", ProcessPre },
                { "<q>", ProcessQuote },
				{ "<span>", ProcessSpan },
                { "<section>", ProcessDiv },
                { "<s>", ProcessHtmlElement<Strike> },
				{ "<strike>", ProcessHtmlElement<Strike> },
				{ "<strong>", ProcessHtmlElement<Bold> },
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
				{ "<xml>", ProcessXmlDataIsland },

				// closing tag
                { "</article>", ProcessClosingDiv },
                { "</aside>", ProcessClosingDiv },
                { "</b>", ProcessClosingTag },
				{ "</body>", ProcessClosingTag },
				{ "</cite>", ProcessClosingTag },
				{ "</del>", ProcessClosingTag },
				{ "</div>", ProcessClosingDiv },
				{ "</em>", ProcessClosingTag },
				{ "</font>", ProcessClosingTag },
                { "</html>", ProcessClosingTag },
				{ "</i>", ProcessClosingTag },
				{ "</ins>", ProcessClosingTag },
				{ "</ol>", ProcessClosingNumberingList },
                { "</p>", ProcessClosingParagraph },
                { "</q>", ProcessClosingQuote },
				{ "</span>", ProcessClosingTag },
				{ "</s>", ProcessClosingTag },
                { "</section>", ProcessClosingDiv },
                { "</strike>", ProcessClosingTag },
				{ "</strong>", ProcessClosingTag },
				{ "</sub>", ProcessClosingTag },
				{ "</sup>", ProcessClosingTag },
				{ "</table>", ProcessClosingTable },
				{ "</tbody>", ProcessClosingTablePart },
				{ "</tfoot>", ProcessClosingTablePart },
				{ "</thead>", ProcessClosingTablePart },
				{ "</td>", ProcessClosingTableColumn },
				{ "</th>", ProcessClosingTableColumn },
				{ "</tr>", ProcessClosingTableRow },
				{ "</u>", ProcessClosingTag },
				{ "</ul>", ProcessClosingNumberingList },
			};

			return knownTags;
		}

		#endregion

		#region CompleteCurrentParagraph

		/// <summary>
		/// Push the elements members to the current paragraph and reset the elements collection.
		/// </summary>
		/// <param name="createNew">True to automatically create a new paragraph, stored in the instance member <see cref="_currentParagraph"/>.</param>
		private void CompleteCurrentParagraph(bool createNew = false)
		{
			_htmlStyles.Paragraph.ApplyTags(_currentParagraph);
			this._currentParagraph.Append(_elements);
			_elements.Clear();

			if (createNew && _currentParagraph.ChildElements.Count > 0)
				AddParagraph(_currentParagraph = _htmlStyles.Paragraph.NewParagraph());
		}

		#endregion

		#region RefreshStyle

		/// <summary>
		/// Refresh the cache of styles presents in the document.
		/// </summary>
		public void RefreshStyles()
		{
			_htmlStyles.PrepareStyles(_mainPart);
		}

		#endregion

		#region ProcessContainerAttributes

		/// <summary>
		/// There is a few attributes shared by a large number of tags. This method will check them for a limited
		/// number of tags (&lt;p&gt;, &lt;pre&gt;, &lt;div&gt;, &lt;span&gt; and &lt;body&gt;).
		/// </summary>
		/// <returns>Returns true if the processing of this tag should generate a new paragraph.</returns>
		private bool ProcessContainerAttributes(HtmlEnumerator en, IList<OpenXmlElement> styleAttributes)
		{
			bool newParagraph = false;

			// Not applicable to a table : page break
			if (!_tables.HasContext || en.CurrentTag == "<pre>")
			{
				String attrValue = en.StyleAttributes["page-break-after"];
				if (attrValue == "always")
				{
					_paragraphs.Add(new Paragraph(
						new Run(
							new Break() { Type = BreakValues.Page })));
				}

				attrValue = en.StyleAttributes["page-break-before"];
				if (attrValue == "always")
				{
					_elements.Add(
						new Run(
							new Break() { Type = BreakValues.Page })
					);
					_elements.Add(new Run(
							new LastRenderedPageBreak())
					);
				}
			}

            // support left and right padding
            var padding = en.StyleAttributes.GetAsMargin("padding");
            if (!padding.IsEmpty && (padding.Left.IsFixed || padding.Right.IsFixed))
			{
                Indentation indentation = new Indentation();
                if (padding.Left.Value > 0) indentation.Left = padding.Left.ValueInDxa.ToString(CultureInfo.InvariantCulture);
                if (padding.Right.Value > 0) indentation.Right = padding.Right.ValueInDxa.ToString(CultureInfo.InvariantCulture);

			    _currentParagraph.InsertInProperties(prop => prop.Indentation = indentation);
			}

			newParagraph |= _htmlStyles.Paragraph.ProcessCommonAttributes(en, styleAttributes);
			return newParagraph;
		}

		#endregion

		#region ChangePageOrientation

		/// <summary>
		/// Generate the required OpenXml element for handling page orientation.
		/// </summary>
		private static SectionProperties ChangePageOrientation(PageOrientationValues orientation)
		{
			PageSize pageSize = new PageSize() { Width = (UInt32Value) 16838U, Height = (UInt32Value) 11906U };
			if (orientation == PageOrientationValues.Portrait)
			{
				UInt32Value swap = pageSize.Width;
				pageSize.Width = pageSize.Height;
				pageSize.Height = swap;
			}
			else
			{
				pageSize.Orient = orientation;
			}

			return new SectionProperties (
				pageSize,
				new PageMargin() {
					Top = 1417, Right = (UInt32Value) 1417U, Bottom = 1417, Left = (UInt32Value) 1417U,
					Header = (UInt32Value) 708U, Footer = (UInt32Value) 708U, Gutter = (UInt32Value) 0U
				},
				new Columns() { Space = "708" },
				new DocGrid() { LinePitch = 360 }
			);
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
		/// Gets or sets whether the &lt;div&gt; tag should be processed as &lt;p&gt; (default false). It depends whether you consider &lt;div&gt;
		/// as part of the layout or as part of a text field.
		/// </summary>
		public bool ConsiderDivAsParagraph { get; set; }

		/// <summary>
		/// Gets or sets whether anchor links are included or not in the conversion.
		/// </summary>
		/// <remarks>An anchor is a term used to define a hyperlink destination inside a document.
		/// <see href="http://www.w3schools.com/HTML/html_links.asp"/>.
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
			get { return _htmlStyles; }
		}

        /// <summary>
        /// Gets or sets how the &lt;img&gt; tag should be handled.
        /// </summary>
        [Obsolete("Provide a IWebRequest implementation or use DefaultWebRequest")]
        public ImageProcessing ImageProcessing { get; set; } = ImageProcessing.AutomaticDownload;

        /// <summary>
        /// Gets or sets the base Uri used to automaticaly resolve relative images 
        /// if used with ImageProcessing = AutomaticDownload.
        /// </summary>
        [Obsolete("Provide a IWebRequest implementation or use DefaultWebRequest.BaseImageUrl")]
        public Uri BaseImageUrl
        {
            get { return (_webRequester as DefaultWebRequest)?.BaseImageUrl; }
            set
            {
                if (value != null)
                {
                    if (!value.IsAbsoluteUri)
                        throw new ArgumentException("BaseImageUrl should be an absolute Uri");
                    // in case of local uri (file:///) we need to be sure the uri ends with '/' or the
                    // combination of uri = new Uri(@"C:\users\demo\images", "pic.jpg");
                    // will eat the images part
                    if (value.IsFile && value.LocalPath[value.LocalPath.Length - 1] != '/')
                        value = new Uri(value.OriginalString + '/');
                }
                if (_webRequester is DefaultWebRequest wr)
                    wr.BaseImageUrl = value;
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
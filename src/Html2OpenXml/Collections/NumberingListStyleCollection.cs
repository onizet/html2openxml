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
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml
{
	sealed class NumberingListStyleCollection
	{
		private MainDocumentPart mainPart;
		private int nextInstanceID, levelDepth;
        private int maxlevelDepth = 0;
        private bool firstItem;
		private Dictionary<String, Int32> knownAbsNumIds;
		private Stack<KeyValuePair<Int32, int>> numInstances;


		public NumberingListStyleCollection(MainDocumentPart mainPart)
		{
			this.mainPart = mainPart;
			this.numInstances = new Stack<KeyValuePair<Int32, int>>();
			InitNumberingIds();
		}


		#region InitNumberingIds

		private void InitNumberingIds()
		{
			NumberingDefinitionsPart numberingPart = mainPart.NumberingDefinitionsPart;
			int absNumIdRef = 0;

			// Ensure the numbering.xml file exists or any numbering or bullets list will results
			// in simple numbering list (1.   2.   3...)
			if (numberingPart == null)
				numberingPart = numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();

			if (mainPart.NumberingDefinitionsPart.Numbering == null)
			{
				new Numbering().Save(numberingPart);
			}
			else
			{
				// The absNumIdRef Id is a required field and should be unique. We will loop through the existing Numbering definition
				// to retrieve the highest Id and reconstruct our own list definition template.
				foreach (var abs in numberingPart.Numbering.Elements<AbstractNum>())
				{
					if (abs.AbstractNumberId.HasValue && abs.AbstractNumberId > absNumIdRef)
						absNumIdRef = abs.AbstractNumberId;
				}
				absNumIdRef++;
			}

			// This minimal numbering definition has been inspired by the documentation OfficeXMLMarkupExplained_en.docx
			// http://www.microsoft.com/downloads/details.aspx?FamilyID=6f264d0b-23e8-43fe-9f82-9ab627e5eaa3&displaylang=en

			AbstractNum[] absNumChildren = new [] {
				//8 kinds of abstractnum + 1 multi-level.
				new AbstractNum(
					new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
					new Level {
						StartNumberingValue = new StartNumberingValue() { Val = 1 },
						NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.Decimal },
						LevelIndex = 0,
						LevelText = new LevelText() { Val = "%1." },
						PreviousParagraphProperties = new PreviousParagraphProperties {
							Indentation = new Indentation() { Left = "420", Hanging = "360" }
						}
					}
				) { AbstractNumberId = absNumIdRef, AbstractNumDefinitionName = new AbstractNumDefinitionName() { Val = "decimal" } },
				new AbstractNum(
					new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
					new Level {
						NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.Bullet },
						LevelIndex = 0,
						LevelText = new LevelText() { Val = "•" },
						PreviousParagraphProperties = new PreviousParagraphProperties {
							Indentation = new Indentation() { Left = "420", Hanging = "360" }
						}
					}
				) { AbstractNumberId = absNumIdRef + 1, AbstractNumDefinitionName = new AbstractNumDefinitionName() { Val = "disc" } },
				new AbstractNum(
					new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
					new Level {
						NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.Bullet },
						LevelIndex = 0,
						LevelText = new LevelText() { Val = "▪" },
						PreviousParagraphProperties = new PreviousParagraphProperties {
							Indentation = new Indentation() { Left = "420", Hanging = "360" }
						}
					}
				) { AbstractNumberId = absNumIdRef + 2, AbstractNumDefinitionName = new AbstractNumDefinitionName() { Val = "square" } },
				new AbstractNum(
					new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
					new Level {
						NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.Bullet },
						LevelIndex = 0,
						LevelText = new LevelText() { Val = "o" },
						PreviousParagraphProperties = new PreviousParagraphProperties {
							Indentation = new Indentation() { Left = "420", Hanging = "360" }
						}
					}
				) { AbstractNumberId = absNumIdRef + 3, AbstractNumDefinitionName = new AbstractNumDefinitionName() { Val = "circle" } },
				new AbstractNum(
					new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
					new Level {
						StartNumberingValue = new StartNumberingValue() { Val = 1 },
						NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.UpperLetter },
						LevelIndex = 0,
						LevelText = new LevelText() { Val = "%1." },
						PreviousParagraphProperties = new PreviousParagraphProperties {
							Indentation = new Indentation() { Left = "420", Hanging = "360" }
						}
					}
				) { AbstractNumberId = absNumIdRef + 4, AbstractNumDefinitionName = new AbstractNumDefinitionName() { Val = "upper-alpha" } },
				new AbstractNum(
					new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
					new Level {
						StartNumberingValue = new StartNumberingValue() { Val = 1 },
						NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.LowerLetter },
						LevelIndex = 0,
						LevelText = new LevelText() { Val = "%1." },
						PreviousParagraphProperties = new PreviousParagraphProperties {
							Indentation = new Indentation() { Left = "420", Hanging = "360" }
						}
					}
				) { AbstractNumberId = absNumIdRef + 5, AbstractNumDefinitionName = new AbstractNumDefinitionName() { Val = "lower-alpha" } },
				new AbstractNum(
					new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
					new Level {
						StartNumberingValue = new StartNumberingValue() { Val = 1 },
						NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.UpperRoman },
						LevelIndex = 0,
						LevelText = new LevelText() { Val = "%1." },
						PreviousParagraphProperties = new PreviousParagraphProperties {
							Indentation = new Indentation() { Left = "420", Hanging = "360" }
						}
					}
				) { AbstractNumberId = absNumIdRef + 6, AbstractNumDefinitionName = new AbstractNumDefinitionName() { Val = "upper-roman" } },
				new AbstractNum(
					new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
					new Level {
						StartNumberingValue = new StartNumberingValue() { Val = 1 },
						NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.LowerRoman },
						LevelIndex = 0,
						LevelText = new LevelText() { Val = "%1." },
						PreviousParagraphProperties = new PreviousParagraphProperties {
							Indentation = new Indentation() { Left = "420", Hanging = "360" }
						}
					}
				) { AbstractNumberId = absNumIdRef + 7, AbstractNumDefinitionName = new AbstractNumDefinitionName() { Val = "lower-roman" } },
				// decimal-title-multi
				// WARNING: only use this for headings
				new AbstractNum(
					new MultiLevelType() { Val = MultiLevelValues.Multilevel },
					new Level {
						StartNumberingValue = new StartNumberingValue() { Val = 1 },
						NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.Decimal },
						LevelIndex = 0,
						LevelText = new LevelText() { Val = "%1." }
					},
					new Level {
						StartNumberingValue = new StartNumberingValue() { Val = 1 },
						NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.Decimal },
						LevelIndex = 1,
						LevelText = new LevelText() { Val = "%1.%2." }
					},
					new Level {
						StartNumberingValue = new StartNumberingValue() { Val = 1 },
						NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.Decimal },
						LevelIndex = 2,
						LevelText = new LevelText() { Val = "%1.%2.%3." }
					}
				) { AbstractNumberId = absNumIdRef + 8, AbstractNumDefinitionName = new AbstractNumDefinitionName() { Val = "decimal-title-multi" } }
			};

			// Check if we have already initialized our abstract nums
			// if that is the case, we should not add them again.
			// This supports a use-case where the HtmlConverter is called multiple times
			// on document generation, and needs to continue existing lists
			bool addNewAbstractNums = false;
			IEnumerable<AbstractNum> existingAbstractNums = numberingPart.Numbering.ChildElements.ToList().Where(e => e != null && e is AbstractNum).Cast<AbstractNum>();

			if (existingAbstractNums.Count() >= absNumChildren.Length) // means we might have added our own already
			{
				foreach (var abstractNum in absNumChildren)
				{
					// Check if we can find this in the existing document
					addNewAbstractNums = addNewAbstractNums 
						|| !existingAbstractNums.Any(a => a.AbstractNumDefinitionName != null && a.AbstractNumDefinitionName.Val.Value == abstractNum.AbstractNumDefinitionName.Val.Value);
				}
			} else {
				addNewAbstractNums = true;
			}

			if (addNewAbstractNums)
			{
				// this is not documented but MS Word needs that all the AbstractNum are stored consecutively.
				// Otherwise, it will apply the "NoList" style to the existing ListInstances.
				// This is the reason why I insert all the items after the last AbstractNum.
				int lastAbsNumIndex = 0;
				if (absNumIdRef > 0)
				{
					lastAbsNumIndex = numberingPart.Numbering.ChildElements.Count-1;
					for (; lastAbsNumIndex >= 0; lastAbsNumIndex--)
					{
						if(numberingPart.Numbering.ChildElements[lastAbsNumIndex] is AbstractNum)
							break;
					}
				}

				for (int i = 0; i < absNumChildren.Length; i++)
					numberingPart.Numbering.InsertAt(absNumChildren[i], i + lastAbsNumIndex);

				knownAbsNumIds = absNumChildren
					.Cast<AbstractNum>()
					.ToDictionary(a => a.AbstractNumDefinitionName.Val.Value, a => a.AbstractNumberId.Value);
			} 
			else
			{
				knownAbsNumIds = existingAbstractNums.ToList()
					.Where(a => a.AbstractNumDefinitionName != null && a.AbstractNumDefinitionName.Val != null)
					.ToDictionary(a => a.AbstractNumDefinitionName.Val.Value, a => a.AbstractNumberId.Value);
			}

			// compute the next list instance ID seed. We start at 1 because 0 has a special meaning: 
			// The w:numId can contain a value of 0, which is a special value that indicates that numbering was removed
			// at this level of the style hierarchy. While processing this markup, if the w:val='0',
			// the paragraph does not have a list item (http://msdn.microsoft.com/en-us/library/ee922775(office.14).aspx)
			nextInstanceID = 1;
			foreach (NumberingInstance inst in numberingPart.Numbering.Elements<NumberingInstance>())
			{
				if (inst.NumberID.Value > nextInstanceID) nextInstanceID = inst.NumberID;
			}
			numInstances.Push(new KeyValuePair<int, int>(nextInstanceID, -1));

			numberingPart.Numbering.Save();
		}

		#endregion

		#region BeginList

		public void BeginList(HtmlEnumerator en)
		{
			// lookup for a predefined list style in the template collection
			String type = en.StyleAttributes["list-style-type"];
			bool orderedList = en.CurrentTag.Equals("<ol>", StringComparison.OrdinalIgnoreCase);

			CreateList(type, orderedList);
		}

		#endregion

		#region EndList

		public void EndList()
		{
			if (levelDepth > 0)
				numInstances.Pop();  // decrement for nested list
			levelDepth--;
			firstItem = true;
		}

		#endregion

		#region CreateList

		public int CreateList(String type, bool orderedList, int numberOfLevels = 1)
		{
			int absNumId = GetAbsNumIdFromType(type, orderedList);
			int prevAbsNumId = numInstances.Peek().Value;

			firstItem = true;
			levelDepth++;
            if (levelDepth > maxlevelDepth)
            {
                maxlevelDepth = levelDepth;
            }

            // save a NumberingInstance if the nested list style is the same as its ancestor.
            // this allows us to nest <ol> and restart the indentation to 1.
            int currentInstanceId = this.InstanceID;
            if (levelDepth > 1 && absNumId == prevAbsNumId && orderedList)
            {
                EnsureMultilevel(absNumId);
            }
            else
            {
                // For unordered lists (<ul>), create only one NumberingInstance per level
                // (MS Word does not tolerate hundreds of identical NumberingInstances)
                if (orderedList || (levelDepth >= maxlevelDepth))
                {
                    currentInstanceId = ++nextInstanceID;
                    Numbering numbering = mainPart.NumberingDefinitionsPart.Numbering;

					List<OpenXmlElement> elements = new List<OpenXmlElement>();

					elements.Add(new AbstractNumId() { Val = absNumId });

					for (int i = 0; i < numberOfLevels; i++)
					{
						elements.Add(new LevelOverride(
							new StartOverrideNumberingValue() {
								Val = 1
							})
							{
								LevelIndex = i
							});
					}

                    numbering.Append(
                        new NumberingInstance(
                            elements
                        )
                        { NumberID = currentInstanceId });
                }
            }

			numInstances.Push(new KeyValuePair<int, int>(currentInstanceId, absNumId));

			return currentInstanceId;
		}

		#endregion

		#region GetAbsNumIdFromType

		public int GetAbsNumIdFromType(String type, bool orderedList)
		{
			int absNumId;

			if (type == null || !knownAbsNumIds.TryGetValue(type.ToLowerInvariant(), out absNumId))
			{
				if (orderedList)
					absNumId = knownAbsNumIds["decimal"];
				else
					absNumId = knownAbsNumIds["disc"];
			}

			return absNumId;
		}

		#endregion

		#region ProcessItem

		public int ProcessItem(HtmlEnumerator en)
		{
			if (!firstItem) return this.InstanceID;

			firstItem = false;

			// in case a margin has been specifically specified, we need to create a new list template
			// on the fly with a different AbsNumId, in order to let Word doesn't merge the style with its predecessor.
			Margin margin = en.StyleAttributes.GetAsMargin("margin");
			if (margin.Left.Value > 0 && margin.Left.Type == UnitMetric.Pixel)
			{
				Numbering numbering = mainPart.NumberingDefinitionsPart.Numbering;
				foreach (AbstractNum absNum in numbering.Elements<AbstractNum>())
				{
					if (absNum.AbstractNumberId == numInstances.Peek().Value)
					{
						Level lvl = absNum.GetFirstChild<Level>();
						Int32 currentNumId = ++nextInstanceID;

						numbering.Append(
							new AbstractNum(
									new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
									new Level {
										StartNumberingValue = new StartNumberingValue() { Val = 1 },
										NumberingFormat = new NumberingFormat() { Val = lvl.NumberingFormat.Val },
										LevelIndex = 0,
										LevelText = new LevelText() { Val = lvl.LevelText.Val }
									}
								) { AbstractNumberId = currentNumId });
						numbering.Save(mainPart.NumberingDefinitionsPart);
						numbering.Append(
							new NumberingInstance(
									new AbstractNumId() { Val = currentNumId }
								) { NumberID = currentNumId });
						numbering.Save(mainPart.NumberingDefinitionsPart);
						mainPart.NumberingDefinitionsPart.Numbering.Reload();
						break;
					}
				}
			}

			return this.InstanceID;
		}

		#endregion

		#region EnsureMultilevel

		/// <summary>
		/// Find a specified AbstractNum by its ID and update its definition to make it multi-level.
		/// </summary>
		private void EnsureMultilevel(int absNumId)
		{
			AbstractNum absNumMultilevel = null;
			foreach (AbstractNum absNum in mainPart.NumberingDefinitionsPart.Numbering.Elements<AbstractNum>())
			{
				if (absNum.AbstractNumberId == absNumId)
				{
					absNumMultilevel = absNum;
					break;
				}
			}


			if (absNumMultilevel != null && absNumMultilevel.MultiLevelType.Val == MultiLevelValues.SingleLevel)
			{
				Level level1 = absNumMultilevel.GetFirstChild<Level>();
				absNumMultilevel.MultiLevelType.Val = MultiLevelValues.Multilevel;

				// skip the first level, starts to 2
				for (int i = 2; i < 10; i++)
				{
					absNumMultilevel.Append(new Level {
						StartNumberingValue = new StartNumberingValue() { Val = 1 },
						NumberingFormat = new NumberingFormat() { Val = level1.NumberingFormat.Val },
						LevelIndex = i - 1,
						LevelText = new LevelText() { Val = "%" + i + "." },
						PreviousParagraphProperties = new PreviousParagraphProperties {
							Indentation = new Indentation() { Left = (720 * i).ToString(CultureInfo.InvariantCulture), Hanging = "360" }
						}
					});
				}
			}
		}

		#endregion

		//____________________________________________________________________
		//
		// Properties Implementation

		/// <summary>
		/// Gets the depth level of the current list instance.
		/// </summary>
		public Int32 LevelIndex
		{
			get { return this.levelDepth; }
		}

		/// <summary>
		/// Gets the ID of the current list instance.
		/// </summary>
		private Int32 InstanceID
		{
			get { return this.numInstances.Peek().Key; }
		}
	}
}
using System;
using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;

namespace NotesFor.HtmlToOpenXml
{
	sealed class NumberingListStyleCollection
	{
		private MainDocumentPart mainPart;
		private int nextInstanceID, absNumId, levelDepth;
		private bool firstItem;
		private Dictionary<String, Int32> knonwAbsNumIds;
		private Stack<Int32> numInstances;


		public NumberingListStyleCollection(MainDocumentPart mainPart)
		{
			this.mainPart = mainPart;
			this.absNumId = -1;
			this.numInstances = new Stack<Int32>();
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

			numberingPart.Numbering.Append(
				//8 kinds of abstractnum + 1 multi-level.
				new AbstractNum(
					new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
					new Level(
						new StartNumberingValue() { Val = 1 },
						new NumberingFormat() { Val = NumberFormatValues.Decimal },
						new LevelText() { Val = "%1." },
						new PreviousParagraphProperties(
							new Indentation() { Left = "420", Hanging = "360" })
					) { LevelIndex = 0 }
				) { AbstractNumberId = absNumIdRef },
				new AbstractNum(
					new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
					new Level(
						new NumberingFormat() { Val = NumberFormatValues.Bullet },
						new LevelText() { Val = "•" },
						new PreviousParagraphProperties(
							new Indentation() { Left = "420", Hanging = "360" })
					) { LevelIndex = 0 }
				) { AbstractNumberId = absNumIdRef + 1 },
				new AbstractNum(
					new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
					new Level(
						new NumberingFormat() { Val = NumberFormatValues.Bullet },
						new LevelText() { Val = "▪" },
						new PreviousParagraphProperties(
							new Indentation() { Left = "420", Hanging = "360" })
					) { LevelIndex = 0 }
				) { AbstractNumberId = absNumIdRef + 2 },
				new AbstractNum(
					new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
					new Level(
						new NumberingFormat() { Val = NumberFormatValues.Bullet },
						new LevelText() { Val = "o" },
						new PreviousParagraphProperties(
							new Indentation() { Left = "420", Hanging = "360" })
					) { LevelIndex = 0 }
				) { AbstractNumberId = absNumIdRef + 3 },
				new AbstractNum(
					new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
					new Level(
						new StartNumberingValue() { Val = 1 },
						new NumberingFormat() { Val = NumberFormatValues.UpperLetter },
						new LevelText() { Val = "%1." },
						new PreviousParagraphProperties(
							new Indentation() { Left = "420", Hanging = "360" })
					) { LevelIndex = 0 }
				) { AbstractNumberId = absNumIdRef + 4 },
				new AbstractNum(
					new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
					new Level(
						new StartNumberingValue() { Val = 1 },
						new NumberingFormat() { Val = NumberFormatValues.LowerLetter },
						new LevelText() { Val = "%1." },
						new PreviousParagraphProperties(
							new Indentation() { Left = "420", Hanging = "360" })
					) { LevelIndex = 0 }
				) { AbstractNumberId = absNumIdRef + 5 },
				new AbstractNum(
					new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
					new Level(
						new StartNumberingValue() { Val = 1 },
						new NumberingFormat() { Val = NumberFormatValues.UpperRoman },
						new LevelText() { Val = "%1." },
						new PreviousParagraphProperties(
							new Indentation() { Left = "420", Hanging = "360" })
					) { LevelIndex = 0 }
				) { AbstractNumberId = absNumIdRef + 6 },
				new AbstractNum(
					new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
					new Level(
						new StartNumberingValue() { Val = 1 },
						new NumberingFormat() { Val = NumberFormatValues.LowerRoman },
						new LevelText() { Val = "%1." },
						new PreviousParagraphProperties(
							new Indentation() { Left = "420", Hanging = "360" })
					) { LevelIndex = 0 }
				) { AbstractNumberId = absNumIdRef + 7 }
			);

			// initializes the lookup
			knonwAbsNumIds = new Dictionary<String, Int32>() {
				{ "disc", absNumIdRef+1 }, { "square", absNumIdRef+2 }, { "circle" , absNumIdRef+3 },
				{ "upper-alpha", absNumIdRef+4 }, { "lower-alpha", absNumIdRef+5 },
				{ "upper-roman", absNumIdRef+6 }, { "lower-roman", absNumIdRef+7 },
				{ "decimal", absNumIdRef }
			};

			// compute the next list instance ID seed
			nextInstanceID = 4; // 4 stands for the default value from Word
			foreach (NumberingInstance inst in numberingPart.Numbering.Elements<NumberingInstance>())
			{
				if (inst.NumberID.Value > nextInstanceID) nextInstanceID = inst.NumberID;
			}
			numInstances.Push(nextInstanceID);

			numberingPart.Numbering.Save(numberingPart);
		}

		#endregion

		#region BeginList

		public void BeginList(HtmlEnumerator en)
		{
			int prevAbsNumId = absNumId;

			// lookup for a predefined list style in the template collection
			String type = en.StyleAttributes["list-style-type"];
			bool orderedList = en.CurrentTag.Equals("<ol>", StringComparison.InvariantCultureIgnoreCase);
			if (type == null || !knonwAbsNumIds.TryGetValue(type.ToLowerInvariant(), out absNumId))
			{
				if (orderedList)
					absNumId = knonwAbsNumIds["decimal"];
				else
					absNumId = knonwAbsNumIds["disc"];
			}

			firstItem = true;
			levelDepth++;

			// save a NumberingInstance if the nested list style is the same as its ancestor.
			// this allows us to nest <ol> and restart the identation to 1.
			int currentInstanceId = this.InstanceID;
			if (levelDepth > 1 && absNumId == prevAbsNumId && orderedList)
			{
				EnsureMultilevel(absNumId);
			}
			else
			{
				currentInstanceId = ++nextInstanceID;
				Numbering numbering = mainPart.NumberingDefinitionsPart.Numbering;
				numbering.Append(
					new NumberingInstance(
						new AbstractNumId() { Val = absNumId },
						new LevelOverride(
							new StartOverrideNumberingValue() { Val = 1 }
						)
					) { NumberID = currentInstanceId });
			}

			numInstances.Push(currentInstanceId);
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

		#region ProcessItem

		public int ProcessItem(HtmlEnumerator en)
		{
			if (!firstItem) return this.InstanceID;

			firstItem = false;

			// in case a margin has been specifically specified, we need to create a new list template
			// on the fly with a different AbsNumId, in order to let Word doesn't merge the style with its predecessor.
			Margin margin = en.StyleAttributes.GetAsMargin("margin");
			if (margin.Left.IsValid && margin.Left.Value > 0 && margin.Left.Type == UnitMetric.Pixel)
			{
				Numbering numbering = mainPart.NumberingDefinitionsPart.Numbering;
				foreach (AbstractNum absNum in numbering.Elements<AbstractNum>())
				{
					if (absNum.AbstractNumberId == absNumId)
					{
						Level lvl = absNum.GetFirstChild<Level>();
						Int32 currentNumId = ++nextInstanceID;

						numbering.Append(
							new AbstractNum(
									new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
									new Level(
										new StartNumberingValue() { Val = 1 },
										new NumberingFormat() { Val = lvl.NumberingFormat.Val },
										new LevelText() { Val = lvl.LevelText.Val }
									) { LevelIndex = 0 }
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
					absNumMultilevel.Append(new Level(
						new StartNumberingValue() { Val = 1 },
						new NumberingFormat() { Val = level1.NumberingFormat.Val },
						new LevelText() { Val = "%" + i + "." },
						new PreviousParagraphProperties(
							new Indentation() { Left = (720 * i).ToString(CultureInfo.InvariantCulture), Hanging = "360" })
					) { LevelIndex = i - 1 });
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
			get { return this.numInstances.Peek(); }
		}
	}
}
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
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace NotesFor.HtmlToOpenXml
{
	using wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;

	/// <summary>
	/// Helper class that provide some extension methods to OpenXml SDK.
	/// </summary>
    [System.Diagnostics.DebuggerStepThrough]
	static class OpenXmlExtension
    {
        public static bool HasChild<T>(this OpenXmlElement element) where T : OpenXmlElement
        {
            return element.GetFirstChild<T>() != null;
        }

		public static T GetLastChild<T>(this OpenXmlElement element) where T : OpenXmlElement
		{
			if (element == null) return null;

			for (int i = element.ChildElements.Count - 1; i >= 0; i--)
			{
				if (element.ChildElements[i] is T)
					return element.ChildElements[i] as T;
			}

			return null;
		}

        public static bool Equals<T>(this EnumValue<T> value, T comparand) where T : struct
        {
            return value != null && value.Value.Equals(comparand);
        }

		public static void InsertInProperties(this Paragraph p, Action<ParagraphProperties> @delegate)
		{
			ParagraphProperties prop = p.GetFirstChild<ParagraphProperties>();
			if (prop == null) p.PrependChild<ParagraphProperties>(prop = new ParagraphProperties());

			@delegate(prop);
		}

		public static void InsertInProperties(this Run r, Action<RunProperties> @delegate)
		{
			RunProperties prop = r.GetFirstChild<RunProperties>();
			if (prop == null) r.PrependChild<RunProperties>(prop = new RunProperties());

			@delegate(prop);
		}

		public static void InsertInDocProperties(this Drawing d, params OpenXmlElement[] newChildren)
		{
			wp.Inline inline = d.GetFirstChild<wp.Inline>();
			wp.DocProperties prop = inline.GetFirstChild<wp.DocProperties>();

			if (prop == null) inline.Append(prop = new wp.DocProperties());
			prop.Append(newChildren);
		}

        public static bool Compare(this PageSize pageSize, PageOrientationValues orientation)
        {
            PageOrientationValues pageOrientation;

            if (pageSize.Orient != null) pageOrientation = pageSize.Orient.Value;
            else if (pageSize.Width > pageSize.Height) pageOrientation = PageOrientationValues.Landscape;
            else pageOrientation = PageOrientationValues.Portrait;

            return pageOrientation == orientation;
        }

		// needed since December 2009 CTP refactoring, where casting is not anymore an option

		public static TableRowAlignmentValues ToTableRowAlignment(this JustificationValues val)
		{
			if (val == JustificationValues.Center) return TableRowAlignmentValues.Center;
			else if (val == JustificationValues.Right) return TableRowAlignmentValues.Right;
			else return TableRowAlignmentValues.Left;
		}
		public static JustificationValues ToJustification(this TableRowAlignmentValues val)
		{
			if (val == TableRowAlignmentValues.Left) return JustificationValues.Left;
			else if (val == TableRowAlignmentValues.Center) return JustificationValues.Center;
			else return JustificationValues.Right;
		}
    }
}
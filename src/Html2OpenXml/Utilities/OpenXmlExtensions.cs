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
using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml;

/// <summary>
/// Helper class that provide some extension methods to OpenXml SDK.
/// </summary>
[System.Diagnostics.DebuggerStepThrough]
static class OpenXmlExtensions
{
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool HasChild<T>(this OpenXmlElement element) where T : OpenXmlElement
    {
        return element.GetFirstChild<T>() != null;
    }

    public static T? GetLastChild<T>(this OpenXmlElement element) where T : OpenXmlElement
    {
        if (element == null) return null;

        for (int i = element.ChildElements.Count - 1; i >= 0; i--)
        {
            if (element.ChildElements[i] is T t)
                return t;
        }

        return null;
    }

    public static SdtBlock? FindNamedSdtBlock(this Body body, string name)
    {
        foreach (var stdBlock in body.Descendants<SdtBlock>())
        {
            if (stdBlock.SdtProperties?
                .GetFirstChild<SdtContentDocPartObject>()?
                .DocPartGallery?.Val?.Value == name)
            {
                return stdBlock;
            }
        }

        return null;
    }

    public static void InsertAt<T>(this Body body,
        T element, DocumentLocation position) where T: OpenXmlElement
    {
        if (position == DocumentLocation.AfterCover)
        {
            var coverPage = body.GetCoverPage();

            // if the cover page does not exist, insert at the start of the document
            if (coverPage != null)
                body.InsertAfter(element, coverPage);
            else
                position = DocumentLocation.DocumentStart;
        }

        if (position == DocumentLocation.DocumentStart)
            body.InsertAt(element, 0);
        else if (position == DocumentLocation.DocumentEnd)
            body.AddChild(element);
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
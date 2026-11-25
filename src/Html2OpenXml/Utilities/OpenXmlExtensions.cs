/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml
 * Licensed under the MIT License.
 * Please see the LICENSE file in the project root fo full license information
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

    /// <summary>
    /// Determines whether the specified page size represents a landscape orientation.
    /// </summary>
    public static bool IsLandscape(this PageSize pageSize)
    {
        return pageSize.Width?.Value >= pageSize.Height?.Value;
    }

    /// <summary>
    /// Swap the orientation of a page.
    /// </summary>
    public static void FlipPage(this PageSize pageSize)
    {
        (pageSize.Height, pageSize.Width) = (pageSize.Width, pageSize.Height);
        pageSize.Orient = pageSize.IsLandscape() ? PageOrientationValues.Landscape : PageOrientationValues.Portrait;
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
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
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml;

/// <summary>
/// Provides information about a style requested during conversion that could not
/// be found in the Word document.
///
/// <para>
/// The <see cref="WordDocumentStyle.StyleMissing" /> event uses these arguments to
/// indicate which style should be dynimacally provisioned and which OpenXml style type is expected.
/// </para>
/// </summary>
public class StyleEventArgs : EventArgs
{
    internal StyleEventArgs(string styleId, StyleValues type)
    {
        Name = styleId;
        Type = type;
    }

    /// <summary>
    /// Gets the identifier of the missing style.
    /// Use this value when creating and registering the style.
    /// </summary>
    public string Name { get; init; }

    /// <summary>
    /// Gets the OpenXml style type expected by the converter,
    /// such as paragraph, table or character.
    /// </summary>
    public StyleValues Type { get; init; }
}
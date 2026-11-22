/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml
 * Licensed under the MIT License.
 * Please see the LICENSE file in the project root fo full license information
 */
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml;

/// <summary>
/// Provides information about a style requested during conversion that could not
/// be found in the Word document.
///
/// <para>
/// The <see cref="WordDocumentStyle.StyleMissing" /> event uses these arguments to
/// indicate which style should be dynamically provisioned and which OpenXml style type is expected.
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
    public string Name { get; private set; }

    /// <summary>
    /// Gets the OpenXml style type expected by the converter,
    /// such as paragraph, table or character.
    /// </summary>
    public StyleValues Type { get; private set; }
}
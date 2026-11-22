/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml
 * Licensed under the MIT License.
 * Please see the LICENSE file in the project root fo full license information
 */
namespace HtmlToOpenXml;

/// <summary>
/// Specifies some decoration to apply to a text.
/// </summary>
[Flags]
enum TextDecoration
{
    None = 0,
    Underline = 1,
    LineThrough = 2,
    Double = 4,
    Dotted = 8,
    Dashed = 16,
    Wave = 32
}

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

namespace HtmlToOpenXml
{
    /// <summary>
    /// Specifies the position of an acronym or abbreviation in the resulting conversion.
    /// </summary>
    public enum AcronymPosition
    {
        /// <summary>
        /// Position at the end of the page.
        /// </summary>
        PageEnd = 0,
        /// <summary>
        /// Position at the end of the document.
        /// </summary>
        DocumentEnd = 1,
    }

    /// <summary>
    /// Predefined quote style as defined by the browser (used for the &lt;q&gt; tag).
    /// </summary>
    /// <remarks>
    /// Initializes a new instance of <see cref="QuoteChars"/> class.
    /// </remarks>
    /// <param name="begin">The characters at the beginning of a quote.</param>
    /// <param name="end">The characters at the end of a quote.</param>
    public readonly struct QuoteChars(string begin, string end)
    {
        /// <summary>Internet Explorer style: « abc » </summary>
        public static readonly QuoteChars IE = new QuoteChars("« ", " »");
        /// <summary>Firefox style: “abc”</summary>
        public static readonly QuoteChars Gecko = new QuoteChars("“", "”");
        /// <summary>Chrome/Safari/Opera style: "abc"</summary>
        public static readonly QuoteChars WebKit = new QuoteChars("\"", "\"");

        internal string Prefix { get; } = begin;
        internal string Suffix { get; } = end;
    }
}
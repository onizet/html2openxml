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

namespace HtmlToOpenXml.Tests
{
    /// <summary>
    /// Helper class that provide some extension methods to String.
    /// </summary>
    [System.Diagnostics.DebuggerStepThrough]
    static class StringExtensions
    {
        public static string Repeat(this string text, uint n)
        {
            var textAsSpan = text.AsSpan();
            var span = new Span<char>(new char[textAsSpan.Length * (int)n]);
            for (var i = 0; i < n; i++)
            {
                textAsSpan.CopyTo(span.Slice(i * textAsSpan.Length, textAsSpan.Length));
            }

            return span.ToString();
        }
    }
}

/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml
 * Licensed under the MIT License.
 * Please see the LICENSE file in the project root fo full license information
 */
using DocumentFormat.OpenXml;

namespace HtmlToOpenXml.Tests
{
    /// <summary>
    /// Helper class that provide some extension methods to OpenXml SDK.
    /// </summary>
    [System.Diagnostics.DebuggerStepThrough]
    static class OpenXmlExtensions
    {
        public static bool HasChild<T>(this OpenXmlElement element) where T : OpenXmlElement
        {
            return element.GetFirstChild<T>() != null;
        }
    }
}
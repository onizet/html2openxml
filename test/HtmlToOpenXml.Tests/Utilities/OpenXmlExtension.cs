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
using DocumentFormat.OpenXml;

namespace HtmlToOpenXml.Tests
{
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
    }
}
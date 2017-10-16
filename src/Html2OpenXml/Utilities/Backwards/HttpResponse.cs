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
    /// Represents a downloaded resource.
    /// </summary>
    sealed class HttpResponse
    {
        /// <summary>
        /// Gets or sets the binary response body.
        /// </summary>
        public byte[] Body { get; set; }

        /// <summary>
        /// Gets or sets the content type of the response.
        /// </summary>
        public string ContentType { get; set; }
    }
}
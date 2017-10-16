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
using System;
using System.Collections.Generic;
using System.Drawing;

namespace HtmlToOpenXml
{
    sealed class HtmlTableSpan : IComparable<HtmlTableSpan>
    {
        public Point CellOrigin;
        public int RowSpan;
        public int ColSpan;

        public HtmlTableSpan(Point origin)
        {
            this.CellOrigin = origin;
        }

        public int CompareTo(HtmlTableSpan other)
        {
            if (other == null) return -1;
            int rc = this.CellOrigin.Y.CompareTo(other.CellOrigin.Y);
            if (rc != 0) return rc;
            return this.CellOrigin.X.CompareTo(other.CellOrigin.X);
        }
    }
}
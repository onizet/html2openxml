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

namespace HtmlToOpenXml
{
    sealed class HtmlTableSpan : IComparable<HtmlTableSpan>
    {
        public CellPosition CellOrigin;
        public int RowSpan;
        public int ColSpan;

        public HtmlTableSpan(CellPosition origin)
        {
            this.CellOrigin = origin;
        }

        public int CompareTo(HtmlTableSpan? other)
        {
            if (other == null) return -1;
            int rc = this.CellOrigin.Row.CompareTo(other.CellOrigin.Row);
            if (rc != 0) return rc;
            return this.CellOrigin.Column.CompareTo(other.CellOrigin.Column);
        }
    }
}
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
    /// Represents the location of cell in a table (2d matrix).
    /// </summary>
    struct CellPosition
    {
        public static readonly CellPosition Empty = new CellPosition();


        /// <summary>
        /// Initializes a new instance of the <see cref='HtmlToOpenXml.CellPosition'/> class from
        /// the specified location.
        /// </summary>
        public CellPosition(int row, int column)
        {
            this.Row = row;
            this.Column = column;
        }

        /// <summary>
        /// Translates this position by the specified amount.
        /// </summary>
        public void Offset(int dr, int dc)
        {
            unchecked
            {
                Row += dr;
                Column += dc;
            }
        }

        /// <summary>
        /// Gets the horizontal coordinate of this position.
        /// </summary>
        public int Row { get; set; }

        /// <summary>
        /// Gets the vertical coordinate of this position.
        /// </summary>
        public int Column { get; set; }
    }
}
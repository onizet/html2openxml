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
using System.ComponentModel;
using System.Globalization;

namespace HtmlToOpenXml
{
	using w = DocumentFormat.OpenXml.Wordprocessing;
	using System.Collections.Generic;


	/// <summary>
	/// Represents a Html Border with the 4 sides.
	/// </summary>
	struct HtmlBorder
	{
		private SideBorder[] _sides;


		public HtmlBorder(SideBorder all)
		{
			if (!all.IsValid) _sides = null;
			else this._sides = new[] { all, all, all, all };
		}

		private void EnsureSides()
		{
			if(this._sides == null) _sides = new SideBorder[4];
		}

		//____________________________________________________________________
		//

		/// <summary>
		/// Gets or sets the border of the bottom side.
		/// </summary>
		public SideBorder Bottom
		{
			get { return _sides == null ? SideBorder.Empty : _sides[2]; }
			set { EnsureSides(); _sides[2] = value; }
		}

		/// <summary>
		/// Gets or sets the border of the left side.
		/// </summary>
		public SideBorder Left
		{
			get { return _sides == null ? SideBorder.Empty : _sides[3]; }
			set { EnsureSides(); _sides[3] = value; }
		}

		/// <summary>
		/// Gets or sets the border of the top side.
		/// </summary>
		public SideBorder Top
		{
			get { return _sides == null ? SideBorder.Empty : _sides[0]; }
			set { EnsureSides(); _sides[0] = value; }
		}

		/// <summary>
		/// Gets or sets the border of the right side.
		/// </summary>
		public SideBorder Right
		{
			get { return _sides == null ? SideBorder.Empty : _sides[1]; }
			set { EnsureSides(); _sides[1] = value; }
		}

		/// <summary>
		/// Gets whether at least one side has been specified.
		/// </summary>
		public bool IsEmpty
		{
			get { return _sides == null || !(Left.IsValid || Right.IsValid || Bottom.IsValid || Top.IsValid); }
		}
	}
}
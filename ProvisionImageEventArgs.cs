/* Copyright (C) Olivier Nizet http://html2openxml.codeplex.com - All Rights Reserved
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
using DocumentFormat.OpenXml.Packaging;

namespace NotesFor.HtmlToOpenXml
{
	/// <summary>
	/// The event arguments used for a ProvisionImage event.
	/// </summary>
	public class ProvisionImageEventArgs : System.ComponentModel.CancelEventArgs
	{
		private HtmlImageInfo info;


		internal ProvisionImageEventArgs(Uri uri, HtmlImageInfo info)
		{
			this.ImageUrl = uri;
			this.info = info;
		}

        public void Provision(byte[] data)
        {
            this.info.RawData = data;
        }

        //____________________________________________________________________
        //

		/// <summary>
		/// Gets the value of the href tag.
		/// </summary>
		public Uri ImageUrl { get; private set; }

		/// <summary>
		/// Gets the styles definition part located inside MainDocumentPart.
		/// </summary>
        [Obsolete("Use Provision(data). Refactoring to match Coding Rule: 'Properties should not return arrays'.")]
        public byte[] Data
		{
			get { return info.RawData; }
			set { info.RawData = value; }
		}

		/// <summary>
		/// Gets or sets the format of the image.
		/// </summary>
		public ImagePartType? ImageExtension
		{
			get { return info.Type; }
			set { info.Type = value; }
		}

		/// <summary>
		/// Gets or sets the width and height (in pixels) of the image as it should be displayed in the document.
		/// </summary>
		public System.Drawing.Size ImageSize
		{
			get { return info.Size; }
			set { info.Size = value; }
		}
	}
}
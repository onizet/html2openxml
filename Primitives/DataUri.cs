﻿/* Copyright (C) Olivier Nizet http://html2openxml.codeplex.com - All Rights Reserved
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
using System.Text.RegularExpressions;
using System.Text;

namespace NotesFor.HtmlToOpenXml
{
	/// <summary>
	/// Represents an URI that includes inline data as if they were external resources.
	/// </summary>
	[System.Diagnostics.DebuggerDisplay("{Mime,nq}")]
	sealed class DataUri
	{
		private String mime;
		private byte[] data;


		private DataUri(String mime, byte[] data)
		{
			this.mime = mime;
			this.data = data;
		}


		public static DataUri Parse(String uri)
		{
			// expected format: data:[<MIME-type>][;charset=<encoding>][;base64],<data>
			// The encoding is indicated by ;base64. If it's present the data is encoded as base64. Without it the data (as a sequence of octets)
			// is represented using ASCII encoding for octets inside the range of safe URL characters and using the standard %xx hex encoding
			// of URLs for octets outside that range. If <MIME-type> is omitted, it defaults to text/plain;charset=US-ASCII.
			// (As a shorthand, the type can be omitted but the charset parameter supplied.)
			// Some browsers (Chrome, Opera, Safari, Firefox) accept a non-standard ordering if both ;base64 and ;charset are supplied,
			// while Internet Explorer requires that the charset's specification must precede the base64 token.
			// http://en.wikipedia.org/wiki/Data_URI_scheme

			// We will stick for IE compliance for the moment...

			Match match = Regex.Match(uri,
				@"data\:(?<mime>\w+/\w+)?(?:;charset=(?<charset>[a-zA-Z_0-9-]+))?(?<base64>;base64)?,(?<data>.*)",
				RegexOptions.IgnoreCase| RegexOptions.Singleline);

			if (!match.Success) return null;

			byte[] rawData = null;
			String mime;
			Encoding charSet = Encoding.ASCII;

			// if mime is omitted, set default value as it stands in the norm
			if (match.Groups["mime"].Length == 0)
				mime = "text/plain";
			else
				mime = match.Groups["mime"].Value;

			if (match.Groups["charset"].Length > 0)
			{
				try
				{
					charSet = Encoding.GetEncoding(match.Groups["charset"].Value);
				}
				catch (ArgumentException)
				{
					// charSet was not recognized
					return null;
				}
			}

			// is it encoded in base64?
			if (match.Groups["base64"].Length > 0)
			{
				// be careful that the raw data is encoded for url (standard %xx hex encoding)
				String base64 = HttpUtility.HtmlDecode(match.Groups["data"].Value);

				try
				{
					rawData = Convert.FromBase64String(base64);
				}
				catch (FormatException)
				{
					// invalid base64
					return null;
				}
			}
			else
			{
				// the <data> represents some text (like html snippet) and must be decoded.
				String raw = HttpUtility.UrlDecode(match.Groups["data"].Value);
				try
				{
					// we convert back to UTF-8 for easier processing later and to have a "referential" encoding
					rawData = Encoding.Convert(charSet, Encoding.UTF8, charSet.GetBytes(raw));
				}
				catch (ArgumentException)
				{
					return null;
				}
			}

			return new DataUri(mime, rawData);
		}

		//____________________________________________________________________
		//

		/// <summary>
		/// Gets the MIME type of the encoded data.
		/// </summary>
		public String Mime
		{
			get { return mime; }
		}

		/// <summary>
		/// Gets the decoded data.
		/// </summary>
		public byte[] Data
		{
			get { return data; }
		}
	}
}
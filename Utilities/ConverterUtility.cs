using System;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using System.Drawing;

namespace NotesFor.HtmlToOpenXml
{
	/// <summary>
	/// Provides some utilies methods for translating Http attributes to OpenXml elements.
	/// </summary>
	static class ConverterUtility
	{
		#region FormatParagraphAlign

		/// <summary>
		/// Convert the Html text align attribute (horizontal alignement) to its corresponding OpenXml value.
		/// </summary>
		public static JustificationValues? FormatParagraphAlign(string htmlAlign)
		{
			switch (htmlAlign)
			{
				case "left": return JustificationValues.Left;
				case "right": return JustificationValues.Right;
				case "center": return JustificationValues.Center;
				case "justify": return JustificationValues.Both;
			}

			return null;
		}

		#endregion

		#region FormatVAlign

		/// <summary>
		/// Convert the Html vertical-align attribute to its corresponding OpenXml value.
		/// </summary>
		public static TableVerticalAlignmentValues? FormatVAlign(string htmlAlign)
		{
			switch (htmlAlign)
			{
				case "top": return TableVerticalAlignmentValues.Top;
				case "middle": return TableVerticalAlignmentValues.Center;
				case "bottom": return TableVerticalAlignmentValues.Bottom;
			}

			return null;
		}

		#endregion

		#region ConvertToFontSize

		/// <summary>
		/// Convert Html regular font-size to OpenXml font value (expressed in point).
		/// </summary>
		public static UInt32 ConvertToFontSize(string htmlSize)
		{
			switch (htmlSize)
			{
				case "1":
				case "xx-small": return 15u;
				case "2":
				case "x-small": return 20u;
				case "4":
				case "medium": return 27u;
				case "5":
				case "large": return 36u;
				case "6":
				case "x-large": return 48u;
				case "7":
				case "xx-large": return 72u;
				case "3":
				case "small":
				default: return 0u;
			}
		}

		#endregion

		#region ConvertToHighlightColor

		/// <summary>
		/// Convert an Html color to its hightlight color representation in OpenXml.
		/// </summary>
		/// <remarks>
		/// As OpenXml supports a limited number of highlight colors, we will check whether the
		/// Html color is near a known color. If no near color are satisfied, it will returns
		/// HighlightColorValues.None.
		/// </remarks>
		public static HighlightColorValues ConvertToHighlightColor(System.Drawing.Color c)
		{
			const int tolerance = 55;

			if (AreColorClose(c, System.Drawing.Color.White, tolerance))
				return HighlightColorValues.White;
			if (AreColorClose(c, System.Drawing.Color.Yellow, tolerance))
				return HighlightColorValues.Yellow;
			if (AreColorClose(c, System.Drawing.Color.Red, tolerance))
				return HighlightColorValues.Red;
			if (AreColorClose(c, System.Drawing.Color.Blue, tolerance))
				return HighlightColorValues.Blue;
			if (AreColorClose(c, System.Drawing.Color.Lime, tolerance))
				return HighlightColorValues.Green;
			if (AreColorClose(c, System.Drawing.Color.Cyan, tolerance))
				return HighlightColorValues.Cyan;
			if (AreColorClose(c, System.Drawing.Color.Fuchsia, tolerance))
				return HighlightColorValues.Magenta;
			if (AreColorClose(c, System.Drawing.Color.Silver, tolerance))
				return HighlightColorValues.LightGray;
			if (AreColorClose(c, System.Drawing.Color.Navy, tolerance))
				return HighlightColorValues.DarkBlue;
			if (AreColorClose(c, System.Drawing.Color.Olive, tolerance))
				return HighlightColorValues.DarkYellow;
			if (AreColorClose(c, System.Drawing.Color.Teal, tolerance))
				return HighlightColorValues.DarkCyan;
			if (AreColorClose(c, System.Drawing.Color.Maroon, tolerance))
				return HighlightColorValues.DarkRed;
			if (AreColorClose(c, System.Drawing.Color.Green, tolerance))
				return HighlightColorValues.DarkGreen;
			if (AreColorClose(c, System.Drawing.Color.Purple, tolerance))
				return HighlightColorValues.DarkMagenta;
			if (AreColorClose(c, System.Drawing.Color.Gray, tolerance))
				return HighlightColorValues.DarkGray;
			if (AreColorClose(c, System.Drawing.Color.Black, tolerance))
				return HighlightColorValues.Black;

			return HighlightColorValues.None;
		}

		#endregion

		#region ConvertToForeColor

		public static System.Drawing.Color ConvertToForeColor(string htmlColor)
		{
			System.Drawing.Color color;

			// The Html allows to write color in hexa without the preceding '#'
			// I just ensure it's a correct hexadecimal value (length=6 and first character should be
			// a digit or an hexa letter)
			if (htmlColor.Length == 6 && (Char.IsDigit(htmlColor[0]) || (htmlColor[0] >= 'a' && htmlColor[0] <= 'f')
				|| (htmlColor[0] >= 'A' && htmlColor[0] <= 'F')))
			{
				try
				{
					color = System.Drawing.Color.FromArgb(
						Convert.ToInt32(htmlColor.Substring(0, 2), 16),
						Convert.ToInt32(htmlColor.Substring(2, 2), 16),
						Convert.ToInt32(htmlColor.Substring(4, 2), 16));
				}
				catch (System.FormatException)
				{
					// If the conversion failed, that should be a named color
					// Let the framework dealing with it
					color = System.Drawing.ColorTranslator.FromHtml(htmlColor);
				}
			}
			else
			{
				color = System.Drawing.ColorTranslator.FromHtml(htmlColor);
			}

			return color;
		}

		#endregion

		#region AreColorClose

		/// <summary>
		/// Check whether two colors are close (as a Magic Wand tool performs).
		/// </summary>
		/// <remarks>
		/// This algorithm comes from the source code of Paint.Net, in FloodToolBase.cs, CheckColor method.
		/// </remarks>
		/// <param name="tolerance">From 0 to 100.</param>
		public static bool AreColorClose(System.Drawing.Color a, System.Drawing.Color b, int tolerance)
		{
			int sum = 0;
			int diff;

			diff = a.R - b.R;
			sum += (1 + diff * diff) * a.A / 256;

			diff = a.G - b.G;
			sum += (1 + diff * diff) * a.A / 256;

			diff = a.B - b.B;
			sum += (1 + diff * diff) * a.A / 256;

			diff = a.A - b.A;
			sum += diff * diff;

			return (sum <= tolerance * tolerance * 4);
		}

		#endregion


		#region DownloadData

		/// <summary>
		/// Download some data located at the specified url.
		/// </summary>
		public static byte[] DownloadData(Uri uri)
		{
			System.Net.WebClient webClient = new System.Net.WebClient();

			try
			{
				return webClient.DownloadData(uri);
			}
			catch (System.Net.WebException)
			{
				return null;
			}
		}

		#endregion

		#region GetImagePartTypeForImageUrl

		private static Dictionary<String, ImagePartType> knownExtensions;

		/// <summary>
		/// Gets the OpenXml ImagePartType associated to an image.
		/// </summary>
		public static ImagePartType? GetImagePartTypeForImageUrl(Uri uri)
		{
			if (knownExtensions == null)
			{
				// Map extension to ImagePartType
				knownExtensions = new Dictionary<String, ImagePartType>(10);
				knownExtensions.Add(".gif", ImagePartType.Gif);
				knownExtensions.Add(".bmp", ImagePartType.Bmp);
				knownExtensions.Add(".emf", ImagePartType.Emf);
				knownExtensions.Add(".ico", ImagePartType.Icon);
				knownExtensions.Add(".jpeg", ImagePartType.Jpeg);
				knownExtensions.Add(".jpg", ImagePartType.Jpeg);
				knownExtensions.Add(".pcx", ImagePartType.Pcx);
				knownExtensions.Add(".png", ImagePartType.Png);
				knownExtensions.Add(".tiff", ImagePartType.Tiff);
				knownExtensions.Add(".wmf", ImagePartType.Wmf);
			}

			String extension = System.IO.Path.GetExtension(uri.AbsoluteUri);
			ImagePartType type;
			if (knownExtensions.TryGetValue(extension, out type)) return type;
			return null;
		}

		#endregion

		#region GetImageSize

		/// <summary>
		/// Loads an image from a stream and grab its size.
		/// </summary>
		internal static Size GetImageSize(Stream imageStream)
		{
			// Read only the size of the image using a little API found on codeproject.
			using (BinaryReader breader = new BinaryReader(imageStream))
				return ImageHeader.GetDimensions(breader);
		}

		#endregion

		//____________________________________________________________________
		//
		// Private Implementation

		static char[] hexDigits = {
         '0', '1', '2', '3', '4', '5', '6', '7',
         '8', '9', 'A', 'B', 'C', 'D', 'E', 'F'};


		#region Color ToHexString

		/// <summary>
		/// Convert a .Net Color to a hex string.
		/// </summary>
		public static string ToHexString(this System.Drawing.Color color)
		{
			// http://www.cambiaresearch.com/c4/24c09e15-2941-4ad2-8695-00b1b4029f4d/Convert-dotnet-Color-to-Hex-String.aspx

			byte[] bytes = new byte[3];
			bytes[0] = color.R;
			bytes[1] = color.G;
			bytes[2] = color.B;
			char[] chars = new char[bytes.Length * 2];
			for (int i = 0; i < bytes.Length; i++)
			{
				int b = bytes[i];
				chars[i * 2] = hexDigits[b >> 4];
				chars[i * 2 + 1] = hexDigits[b & 0xF];
			}
			return new string(chars);
		}

		#endregion
	}
}

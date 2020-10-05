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
using System.Text;
using System.IO;
using System.Globalization;
using System.Collections.Generic;

namespace HtmlToOpenXml
{
	/// <summary>
	/// Helper class that can be found in System.Web.HttpUtility. This class allows us to avoid a reference to System.Web.
	/// </summary>
	static class HttpUtility
	{
		/// <summary>The common characters considered as white space.</summary>
		internal static readonly char[] WhiteSpaces = { ' ', '\t', '\r', '\u00A0', '\u0085' };

		static class HtmlEntities
		{
			private static string[] entitiesList = new string[] { 
				"\"-quot", "&-amp", "<-lt", ">-gt", "\x00a0-nbsp", "\x00a1-iexcl", "\x00a2-cent", "\x00a3-pound", "\x00a4-curren", "\x00a5-yen", "\x00a6-brvbar", "\x00a7-sect", "\x00a8-uml", "\x00a9-copy", "\x00aa-ordf", "\x00ab-laquo", 
				"\x00ac-not", "\x00ad-shy", "\x00ae-reg", "\x00af-macr", "\x00b0-deg", "\x00b1-plusmn", "\x00b2-sup2", "\x00b3-sup3", "\x00b4-acute", "\x00b5-micro", "\x00b6-para", "\x00b7-middot", "\x00b8-cedil", "\x00b9-sup1", "\x00ba-ordm", "\x00bb-raquo", 
				"\x00bc-frac14", "\x00bd-frac12", "\x00be-frac34", "\x00bf-iquest", "\x00c0-Agrave", "\x00c1-Aacute", "\x00c2-Acirc", "\x00c3-Atilde", "\x00c4-Auml", "\x00c5-Aring", "\x00c6-AElig", "\x00c7-Ccedil", "\x00c8-Egrave", "\x00c9-Eacute", "\x00ca-Ecirc", "\x00cb-Euml", 
				"\x00cc-Igrave", "\x00cd-Iacute", "\x00ce-Icirc", "\x00cf-Iuml", "\x00d0-ETH", "\x00d1-Ntilde", "\x00d2-Ograve", "\x00d3-Oacute", "\x00d4-Ocirc", "\x00d5-Otilde", "\x00d6-Ouml", "\x00d7-times", "\x00d8-Oslash", "\x00d9-Ugrave", "\x00da-Uacute", "\x00db-Ucirc", 
				"\x00dc-Uuml", "\x00dd-Yacute", "\x00de-THORN", "\x00df-szlig", "\x00e0-agrave", "\x00e1-aacute", "\x00e2-acirc", "\x00e3-atilde", "\x00e4-auml", "\x00e5-aring", "\x00e6-aelig", "\x00e7-ccedil", "\x00e8-egrave", "\x00e9-eacute", "\x00ea-ecirc", "\x00eb-euml", 
				"\x00ec-igrave", "\x00ed-iacute", "\x00ee-icirc", "\x00ef-iuml", "\x00f0-eth", "\x00f1-ntilde", "\x00f2-ograve", "\x00f3-oacute", "\x00f4-ocirc", "\x00f5-otilde", "\x00f6-ouml", "\x00f7-divide", "\x00f8-oslash", "\x00f9-ugrave", "\x00fa-uacute", "\x00fb-ucirc", 
				"\x00fc-uuml", "\x00fd-yacute", "\x00fe-thorn", "\x00ff-yuml", "Œ-OElig", "œ-oelig", "Š-Scaron", "š-scaron", "Ÿ-Yuml", "ƒ-fnof", "ˆ-circ", "˜-tilde", "Α-Alpha", "Β-Beta", "Γ-Gamma", "Δ-Delta", 
				"Ε-Epsilon", "Ζ-Zeta", "Η-Eta", "Θ-Theta", "Ι-Iota", "Κ-Kappa", "Λ-Lambda", "Μ-Mu", "Ν-Nu", "Ξ-Xi", "Ο-Omicron", "Π-Pi", "Ρ-Rho", "Σ-Sigma", "Τ-Tau", "Υ-Upsilon", 
				"Φ-Phi", "Χ-Chi", "Ψ-Psi", "Ω-Omega", "α-alpha", "β-beta", "γ-gamma", "δ-delta", "ε-epsilon", "ζ-zeta", "η-eta", "θ-theta", "ι-iota", "κ-kappa", "λ-lambda", "μ-mu", 
				"ν-nu", "ξ-xi", "ο-omicron", "π-pi", "ρ-rho", "ς-sigmaf", "σ-sigma", "τ-tau", "υ-upsilon", "φ-phi", "χ-chi", "ψ-psi", "ω-omega", "ϑ-thetasym", "ϒ-upsih", "ϖ-piv", 
				" -ensp", " -emsp", " -thinsp", "‌-zwnj", "‍-zwj", "‎-lrm", "‏-rlm", "–-ndash", "—-mdash", "‘-lsquo", "’-rsquo", "‚-sbquo", "“-ldquo", "”-rdquo", "„-bdquo", "†-dagger", 
				"‡-Dagger", "•-bull", "…-hellip", "‰-permil", "′-prime", "″-Prime", "‹-lsaquo", "›-rsaquo", "‾-oline", "⁄-frasl", "€-euro", "ℑ-image", "℘-weierp", "ℜ-real", "™-trade", "ℵ-alefsym", 
				"←-larr", "↑-uarr", "→-rarr", "↓-darr", "↔-harr", "↵-crarr", "⇐-lArr", "⇑-uArr", "⇒-rArr", "⇓-dArr", "⇔-hArr", "∀-forall", "∂-part", "∃-exist", "∅-empty", "∇-nabla", 
				"∈-isin", "∉-notin", "∋-ni", "∏-prod", "∑-sum", "−-minus", "∗-lowast", "√-radic", "∝-prop", "∞-infin", "∠-ang", "∧-and", "∨-or", "∩-cap", "∪-cup", "∫-int", 
				"∴-there4", "∼-sim", "≅-cong", "≈-asymp", "≠-ne", "≡-equiv", "≤-le", "≥-ge", "⊂-sub", "⊃-sup", "⊄-nsub", "⊆-sube", "⊇-supe", "⊕-oplus", "⊗-otimes", "⊥-perp", 
				"⋅-sdot", "⌈-lceil", "⌉-rceil", "⌊-lfloor", "⌋-rfloor", "〈-lang", "〉-rang", "◊-loz", "♠-spades", "♣-clubs", "♥-hearts", "♦-diams"
			 };
			private static Dictionary<String,Char> entitiesLookupTable;
			private static readonly object SyncObject = new object();

			internal static char Lookup(string entity)
			{
				if (entitiesLookupTable == null)
				{
					lock (SyncObject)
					{
						if (entitiesLookupTable == null)
						{
							Dictionary<String, Char> hashtable = new Dictionary<String, Char>(entitiesList.Length);
							foreach (string str in entitiesList)
								hashtable[str.Substring(2)] = str[0];

							entitiesLookupTable = hashtable;
						}
					}
				}

				char ch;
				if (!entitiesLookupTable.TryGetValue(entity, out ch)) return '\0';
				return ch;
			}
		}

		#region HtmlDecode

		/// <summary>
		/// Converts a string that has been HTML-encoded for HTTP transmission into a decoded string.
		/// </summary>
		public static string HtmlDecode(string s)
		{
			if (s == null) return null;

			if (s.IndexOf('&') < 0) return s;

			StringBuilder sb = new StringBuilder();
			StringWriter output = new StringWriter(sb, CultureInfo.InvariantCulture);
			HtmlDecode(s, output);
			return sb.ToString();
		}

		/// <summary>
		/// Converts a string that has been HTML-encoded into a decoded string, and sends
		/// the decoded string to a System.IO.TextWriter output stream.
		/// </summary>
		public static void HtmlDecode(string s, TextWriter output)
		{
			if (s == null) return;
			if (s.IndexOf('&') < 0) output.Write(s);

			char[] entityEndingChars = new char[] { ';', '&' };

			int length = s.Length;
			for (int i = 0; i < length; i++)
			{
				char ch = s[i];
				if (ch != '&')
				{
					output.Write(ch);
					continue;
				}


				int endIndex = s.IndexOfAny(entityEndingChars, i + 1);
				if (endIndex > 0 && s[endIndex] == ';')
				{
					string entity = s.Substring(i + 1, endIndex - i - 1);
					if (entity.Length > 0 && entity[0] == '#')
					{
						bool success;
						int result;
						if (entity[1] == 'x' || entity[1] == 'X')
						{
							success = Int32.TryParse(entity.Substring(2), NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out result);
						}
						else
						{
							success = Int32.TryParse(entity.Substring(1), out result);
						}

						if (success)
						{
							if (IsLegalXmlChar(result)) output.Write((char) result);
							i = endIndex;
						}
						else
						{
							i++;
						}
					}
					else
					{
						i = endIndex;
						ch = HtmlEntities.Lookup(entity);
						if (ch != '\0') output.Write(ch);
						else
						{
							output.Write('&');
							output.Write(entity);
							output.Write(';');
						}
					}
				}
			}
		}

		#endregion

		#region UrlDecode

		/// <summary>
		/// Converts a string that represents an Html-encoded URL to a decoded string.
		/// </summary>
		public static string UrlDecode(string text)
		{
			// pre-process for + sign space formatting since System.Uri doesn't handle it
			// plus literals are encoded as %2b normally so this should be safe
			// http://www.west-wind.com/weblog/posts/2009/Feb/05/Html-and-Uri-String-Encoding-without-SystemWeb
			text = text.Replace("+", " ");
			return System.Uri.UnescapeDataString(text);
		}

		#endregion

		// Utilities methods

		#region IsLegalXmlChar

		/// <summary>
		/// Gets whether a given character is allowed by XML 1.0.
		/// </summary>
		private static bool IsLegalXmlChar(int character)
		{
			// http://seattlesoftware.wordpress.com/2008/09/11/hexadecimal-value-0-is-an-invalid-character/

			return (
				 character == 0x9 /* == '\t' == 9   */       ||
				 character == 0xA /* == '\n' == 10  */       ||
				 character == 0xD /* == '\r' == 13  */       ||
				(character >= 0x20 && character <= 0xD7FF) ||
				(character >= 0xE000 && character <= 0xFFFD) ||
				(character >= 0x10000 && character <= 0x10FFFF)
			);
		}

		#endregion
	}
}
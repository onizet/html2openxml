using System.Globalization;
using System.Reflection;
using System.Resources;

namespace HtmlToOpenXml
{
    /// <summary>
    /// Helper class to get chunks of OpenXml predefined style.
    /// </summary>
    internal class PredefinedStyles
    {
        private static ResourceManager? resourceMan;

        public const string Caption = "Caption";
        public const string EndnoteText = "EndnoteText";
        public const string EndnoteReference = "EndnoteReference";
        public const string FootnoteText = "FootnoteText";
        public const string FootnoteReference = "FootnoteReference";
        public const string Heading = "Heading";
        public const string Hyperlink = "Hyperlink";
        public const string IntenseQuote = "IntenseQuote";
        public const string ListParagraph = "ListParagraph";
        public const string Quote = "Quote";
        public const string QuoteChar = "QuoteChar";
        public const string TableGrid = "TableGrid";


        /*/// <summary>
        /// Looks up a localized string similar to Caption.
        /// </summary>
        internal static string Caption {
            get {
                return ResourceManager.GetString("Caption", CultureInfo.InvariantCulture);
            }
        }

        /// <summary>
        ///   Looks up a localized string similar HyperLink.
        /// </summary>
        internal static string HyperLink {
            get {
                return ResourceManager.GetString("HyperLink", CultureInfo.InvariantCulture);
            }
        }

        /// <summary>
        ///   Looks up a localized string similar EndnoteText.
        /// </summary>
        internal static string EndnoteText {
            get {
                return ResourceManager.GetString("EndnoteText", CultureInfo.InvariantCulture);
            }
        }

        /// <summary>
        ///   Looks up a localized string similar EndnoteReference.
        /// </summary>
        internal static string EndnoteReference {
            get {
                return ResourceManager.GetString("EndnoteReference", CultureInfo.InvariantCulture);
            }
        }

        /// <summary>
        ///   Looks up a localized string similar FootnoteText.
        /// </summary>
        internal static string FootnoteText {
            get {
                return ResourceManager.GetString("FootnoteText", CultureInfo.InvariantCulture);
            }
        }

        /// <summary>
        ///   Looks up a localized string similar FootnoteReference.
        /// </summary>
        internal static string FootnoteReference {
            get {
                return ResourceManager.GetString("FootnoteReference", CultureInfo.InvariantCulture);
            }
        }

        /// <summary>
        ///   Looks up a localized string similar IntenseQuote.
        /// </summary>
        internal static string IntenseQuote {
            get {
                return ResourceManager.GetString("IntenseQuote", CultureInfo.InvariantCulture);
            }
        }

        /// <summary>
        ///   Looks up a localized string similar ListParagraph.
        /// </summary>
        internal static string ListParagraph {
            get {
                return ResourceManager.GetString("ListParagraph", CultureInfo.InvariantCulture);
            }
        }

        /// <summary>
        ///   Looks up a localized string similar Quote.
        /// </summary>
        internal static string Quote {
            get {
                return ResourceManager.GetString("Quote", CultureInfo.InvariantCulture);
            }
        }

        /// <summary>
        ///   Looks up a localized string similar QuoteChar.
        /// </summary>
        internal static string QuoteChar {
            get {
                return ResourceManager.GetString("QuoteChar", CultureInfo.InvariantCulture);
            }
        }

        /// <summary>
        ///   Looks up a localized string similar TableGrid.
        /// </summary>
        internal static string TableGrid {
            get {
                return ResourceManager.GetString("TableGrid", CultureInfo.InvariantCulture);
            }
        }*/

        /// <summary>
        /// Retrieves the embedded resource.
        /// </summary>
        /// <param name="styleName">The key name of the resource to find.</param>
        public static string? GetOuterXml(string styleName)
        {
            return ResourceManager.GetString(styleName);
        }


        /// <summary>
        /// Returns the cached ResourceManager instance used by this class.
        /// </summary>
        private static ResourceManager ResourceManager
        {
            get
            {
                if (resourceMan is null)
                {
                    ResourceManager temp = new("HtmlToOpenXml.PredefinedStyles",
                        typeof(PredefinedStyles).GetTypeInfo().Assembly);
                    resourceMan = temp;
                }
                return resourceMan;
            }
        }
    }
}

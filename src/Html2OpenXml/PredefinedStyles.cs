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
        private static global::System.Resources.ResourceManager? resourceMan;

        /// <summary>
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

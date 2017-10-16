using System.ComponentModel;
using System.Resources;
using System.Reflection;

namespace HtmlToOpenXml
{
    internal partial class PredefinedStyles
    {
        /// <summary>
        ///   Returns the cached ResourceManager instance used by this class.
        /// </summary>
        [EditorBrowsableAttribute(EditorBrowsableState.Advanced)]
        internal static ResourceManager ResourceManager
        {
            get
            {
                if (object.ReferenceEquals(resourceMan, null))
                {
                    ResourceManager temp = new ResourceManager("HtmlToOpenXml.Properties.PredefinedStyles",
#if FEATURE_REFLECTION
                        typeof(PredefinedStyles).Assembly);
#else
                        typeof(PredefinedStyles).GetTypeInfo().Assembly);
#endif
                    resourceMan = temp;
                }
                return resourceMan;
            }
        }
    }
}

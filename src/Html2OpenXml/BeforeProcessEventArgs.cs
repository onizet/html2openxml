using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HtmlToOpenXml
{
    /// <summary>
    /// The event arguments used for a BeforeProcess event.
    /// </summary>
    public class BeforeProcessEventArgs : EventArgs
    {
        /// <summary>
        /// ctor fot this event args
        /// </summary>
        /// <param name="current"></param>
        /// <param name="tag"></param>
        internal BeforeProcessEventArgs(IHtmlAttributeCollection current, string tag)
        {
            Current = current;
            Tag = tag;
        }
        /// <summary>
        /// The attribute list of the current Tag
        /// </summary>
        public IHtmlAttributeCollection Current { get; private set; }
        /// <summary>
        /// Tag name of the current element
        /// </summary>
        public string Tag { get; private set; }
    }
}

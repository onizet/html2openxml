using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HtmlToOpenXml
{
    /// <summary>
    /// The event arguments used for a AfterProcess event.
    /// </summary>
    public class AfterProcessEventArgs : EventArgs
    {
        /// <summary>
        /// ctor fot this event args
        /// </summary>
        /// <param name="current"></param>
        /// <param name="htmlAttributes"></param>
        /// <param name="currentParagraph"></param>
        /// <param name="tag"></param>
        internal AfterProcessEventArgs(OpenXmlElement current, IHtmlAttributeCollection htmlAttributes, Paragraph currentParagraph, string tag)
        {
            Current = current;
            HtmlAttributes = htmlAttributes;
            CurrentParagraph = currentParagraph;
            Tag = tag;
        }
        /// <summary>
        /// The current OpenXmlElement
        /// </summary>
        public OpenXmlElement Current { get; private set; }
        /// <summary>
        /// The attribute list of the current Tag
        /// </summary>
        public IHtmlAttributeCollection HtmlAttributes { get; private set; }
        /// <summary>
        /// Tag name of the current element
        /// </summary>
        public Paragraph CurrentParagraph { get; private set; }
        /// <summary>
        /// Tag name of the current element
        /// </summary>
        public string Tag { get; private set; }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HtmlToOpenXml
{
    /// <summary>
    /// Interface to use the HtmlAttributeCollection outsite of this project
    /// </summary>
    public interface IHtmlAttributeCollection
    {
        /// <summary>
        /// Indexer to the collection
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        string this[string name] { get; set; }
    }
}

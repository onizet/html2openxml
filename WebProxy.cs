using System;
using System.Net;

namespace NotesFor.HtmlToOpenXml
{
    /// <summary>
    /// Represents the configuration used to download some data such as the images.
    /// </summary>
    public sealed class WebProxy
    {
        /// <summary>
        /// Gets or sets the credentials to submit to the proxy server for authentication.
        /// </summary>
        public ICredentials Credentials { get; set; }

        /// <summary>
        /// Gets or sets the proxy access.
        /// </summary>
        public IWebProxy Proxy { get; set; }
    }
}

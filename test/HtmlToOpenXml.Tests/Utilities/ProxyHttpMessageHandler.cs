/*
 * Copyright (c) 2017 Deal Stream sàrl. All rights reserved
 */
using System.Net.Http;

namespace HtmlToOpenXml.Tests
{
    public class ProxyHttpMessageHandler : HttpMessageHandler
    {
        private readonly Func<Uri, Task<HttpResponseMessage>> _getResponseFunc;

        public ProxyHttpMessageHandler(Func<Uri, Task<HttpResponseMessage>> getResponseFunc)
        {
            _getResponseFunc = getResponseFunc;
        }

        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            return await _getResponseFunc(request.RequestUri!);
        }
    }
}
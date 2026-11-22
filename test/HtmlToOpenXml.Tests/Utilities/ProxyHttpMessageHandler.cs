/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml
 * Licensed under the MIT License.
 * Please see the LICENSE file in the project root fo full license information
 */
using System.Net.Http;

namespace HtmlToOpenXml.Tests;

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

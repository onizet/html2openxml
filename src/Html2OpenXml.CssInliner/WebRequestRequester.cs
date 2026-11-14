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
using AngleSharp.Io;

namespace HtmlToOpenXml.CssInliner;

/// <summary>
/// Proxy implementation of the AngleSharp requestor to reuse HtmlToConverter resource loader.
/// </summary>
sealed class WebRequestRequester(IO.IWebRequest webRequest) : BaseRequester
{
    public override bool SupportsProtocol(string protocol)
        => webRequest.SupportsProtocol(protocol);

    protected override async Task<IResponse?> PerformRequestAsync(Request request, CancellationToken cancel)
    {
        var resource = await webRequest.FetchAsync(request.Address, cancel);

        if (resource is null)
        {
            return VirtualResponse.Create(response => {
                response.Status(System.Net.HttpStatusCode.NotFound);
                response.Address(request.Address);
            });
        }

        return VirtualResponse.Create(response => {
            response.Address(request.Address);
            response.Status((int) resource.StatusCode);

            foreach (var header in resource.Headers)
            {
                response.Header(header.Key, header.Value);
            }

            if (resource.Content != null)
                response.Content(resource.Content);
        });
    }
}
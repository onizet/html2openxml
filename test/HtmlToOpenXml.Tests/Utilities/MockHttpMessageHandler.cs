/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml
 * Licensed under the MIT License.
 * Please see the LICENSE file in the project root fo full license information
 */
using Moq;
using Moq.Protected;
using System.Net.Http;

namespace HtmlToOpenXml.Tests;

public class MockHttpMessageHandler
{
    private readonly Mock<HttpMessageHandler> mockMessageHandler;


    public MockHttpMessageHandler()
    {
        mockMessageHandler = new Mock<HttpMessageHandler>();
        mockMessageHandler.Protected()
            .Setup<Task<HttpResponseMessage>>("SendAsync", ItExpr.IsAny<HttpRequestMessage>(), ItExpr.IsAny<CancellationToken>())
            .ReturnsAsync(new HttpResponseMessage());
    }

    public IO.IWebRequest GetWebRequest()
    {
        return new IO.DefaultWebRequest(new HttpClient(mockMessageHandler.Object));
    }

    public void AssertNeverCalled()
    {
        mockMessageHandler.Protected()
            .Verify("SendAsync", Times.Never(),
            ItExpr.IsAny<HttpRequestMessage>(), ItExpr.IsAny<CancellationToken>());
    }
}

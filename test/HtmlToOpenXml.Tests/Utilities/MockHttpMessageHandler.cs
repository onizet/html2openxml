/*
 * Copyright (c) 2017 Deal Stream sàrl. All rights reserved
 */
using Moq;
using Moq.Protected;
using System.Net.Http;

namespace HtmlToOpenXml.Tests
{
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
}
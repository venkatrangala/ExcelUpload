using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Mmu.Common.Api.Service.Interfaces;
using MMU.Functions.Helpers;
//using Mmu.Unit4ApiPublisher.Helpers;
//using Mmu.Unit4ApiPublisher.Interfaces;

namespace Mmu.Functions.Services
{
    public class CookieMessageFactory : IHttpRequestMessageFactory
    {
        private readonly ITokenService<CookieInfo> _tokenService;

        public CookieMessageFactory(ITokenService<CookieInfo> tokenService)
        {
            _tokenService = tokenService;
        }

        public async Task<HttpRequestMessage> CreateMessage(HttpMethod method, Uri uri, string payload)
        {

            var message = new HttpRequestMessage(method, uri)
            {
                Content = new StringContent(payload, Encoding.UTF8, "application/json")
            };

            var tokenInfo = await _tokenService.GetToken();

            message.Headers.Add("Cookie", $".MOSAICANON={tokenInfo.Cookies.MosaicAnon}; .MosaicAuthorization={tokenInfo.Cookies.MosaicAuthorization}");

            return message;

        }
    }
}

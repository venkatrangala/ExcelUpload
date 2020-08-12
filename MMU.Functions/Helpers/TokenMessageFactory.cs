using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.Options;
using Mmu.Common.Api.Service.Interfaces;
using Mmu.Common.Api.Service.Models;
using MMU.Functions.Interfaces;

namespace MMU.Functions.Helpers
{
    public interface IHttpRequestMessageFactory
    {
        Task<HttpRequestMessage> CreateMessage(HttpMethod method, Uri uri, string payload);
    }

    public class TokenMessageFactory : IHttpRequestMessageFactory
    {
        private readonly ITokenService<TokenInfo> _tokenService;
        private readonly EndPointConfigU4 _config;

        public TokenMessageFactory(ITokenService<TokenInfo> tokenService, IOptions<EndPointConfigU4> options)
        {
            _tokenService = tokenService;
            _config = options.Value;
        }

        public async Task<HttpRequestMessage> CreateMessage(HttpMethod method, Uri uri, string payload)
        {

            var message = new HttpRequestMessage(method, uri)
            {
                Content = new StringContent(payload, Encoding.UTF8, "application/json")
            };

            var tokenInfo = await _tokenService.GetToken();

            message.Headers.Add("Authorization", $"Bearer {tokenInfo.Access_Token}");
            message.Headers.Add("unit4_id", _config.Unit4IdClaim);

            return message;

        }
    }
}

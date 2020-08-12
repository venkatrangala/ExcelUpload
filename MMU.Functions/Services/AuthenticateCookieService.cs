
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Extensions.Options;
using Mmu.Common.Api.Service.Interfaces;
using Mmu.Common.Api.Service.Models;
using MMU.Functions.Helpers;
//using Mmu.Unit4ApiPublisher.Helpers;
//using Mmu.Unit4ApiPublisher.Interfaces;

namespace Mmu.Functions.Services
{
    public class AuthenticateCookieService : IAuthenticateService<CookieInfo>
    {
        private static HttpClient _client;
        private readonly EndpointConfig _config;

        public AuthenticateCookieService(IHttpClientProvider clientProvider, IOptions<EndpointConfig> options)
        {
            _client = clientProvider.HttpClient;
            _config = options.Value;
        }

        public async Task<CookieInfo> Authenticate()
        {
            var message = new HttpRequestMessage(HttpMethod.Post, _config.LoginUri)
            {
                Content = new FormUrlEncodedContent(new[]
                {
                    new KeyValuePair<string, string>("grant_type", "password"),
                    new KeyValuePair<string, string>("username", _config.Username),
                    new KeyValuePair<string, string>("password", _config.Password),
                })
            };

            message.Headers.Add("Cookie", "GetCookies=.MOSAICANON; GetCookies=.MosaicAuthorization");
            message.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            var responseMessage = await _client.SendAsync(message);

            responseMessage.EnsureSuccessStatusCode();

            var u4SmCookies = SortOutCookies(responseMessage.Headers.GetValues("Set-Cookie"));

            var user = responseMessage.Content.ReadAsStringAsync().Result.Replace("\"", "");

            if (string.IsNullOrEmpty(u4SmCookies.MosaicAuthorization) && string.IsNullOrEmpty(user))
            {
                throw new ApplicationException("Unable to authenticate with U4SM API");
            }

            return new CookieInfo()
            {
                Cookies = u4SmCookies,
                ExpiryDate = DateTime.UtcNow.AddDays(30)
            };
        }

        private static Cookies SortOutCookies(IEnumerable<string> cookies)
        {
            var u4SmCookies = new Cookies();

            foreach (var cookie in cookies)
            {
                if (cookie.Contains(".MOSAICANON"))
                {
                    u4SmCookies.MosaicAnon = cookie.Replace(".MOSAICANON=", "");
                }

                if (cookie.Contains(".MosaicAuthorization"))
                {
                    u4SmCookies.MosaicAuthorization =
                        cookie.Remove(cookie.IndexOf(";", StringComparison.Ordinal))
                            .Replace(".MosaicAuthorization=", "");
                }
            }

            return u4SmCookies;
        }
    }
}
using System;
using System.Net.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;
using Mmu.Common.Api.Service.Authentication;
using Mmu.Common.Api.Service.Extensions;
using Mmu.Common.Api.Service.Helper;
using Mmu.Common.Api.Service.Interfaces;
using Mmu.Common.Api.Service.Models;
using Mmu.Functions.Services;
using MMU.Functions.Helpers;
//using Mmu.Unit4ApiPublisher.Helpers;
//using Mmu.Unit4ApiPublisher.Interfaces;
//using Mmu.Unit4ApiPublisher.Services;

namespace Mmu.Functions.Extensions
{

    /// <summary>
    /// Extension methods for setting up Unit4 Wrapper services in an <see cref="T:Microsoft.Extensions.DependencyInjection.IServiceCollection" />.
    /// </summary>
    public static class ServiceCollectionExtensions
    {
        /// <summary>
        /// Adds Unit4 Wrapper services to the specified<see cref="T:Microsoft.Extensions.DependencyInjection.IServiceCollection" />
        /// using the given<see cref= "T:Microsoft.Extensions.Configuration.IConfigurationRoot" />.
        /// </ summary >
        /// < param name= "services" > The < see cref= "T:Microsoft.Extensions.DependencyInjection.IServiceCollection" /> to add services to.</param>
        /// <param name = "configuration" > The < see cref= "T:Microsoft.Extensions.Configuration.IConfigurationRoot" /> to read configuration values from.</param>
        /// <returns>The<see cref="T:Microsoft.Extensions.DependencyInjection.IServiceCollection" /> so that additional calls can be chained.</returns>
        public static IServiceCollection AddU4Service(this IServiceCollection services, IConfiguration configuration)
        {
            IOptions<EndPointConfigU4> options = Options.Create(new EndPointConfigU4());

            if (configuration["Unit4Api:GrantType"] != null &&
            configuration["Unit4Api:GrantType"].Equals("password", StringComparison.InvariantCultureIgnoreCase))
            {
                options.Value.Password = configuration["Unit4Api:Password"];
                options.Value.Username = configuration["Unit4Api:Username"];
                options.Value.BaseUrl = configuration["Unit4Api:BaseUri"];
                options.Value.LoginUri = configuration["Unit4Api:LoginUri"];
                services.AddSingleton<IHttpRequestMessageFactory, CookieMessageFactory>();
                services.AddCookieAuthenticationService(options);
            }
            else
            {
                options.Value.BaseUrl = configuration["Unit4Api:BaseUri"];
                options.Value.LoginUri = configuration["Unit4Api:TokenLoginUri"];
                options.Value.ClientId = configuration["Unit4Api:ClientId"];
                options.Value.ClientSecret = configuration["Unit4Api:ClientSecret"];
                options.Value.Scope = configuration["Unit4Api:Scope"];
                options.Value.Unit4IdClaim = configuration["Unit4Api:IdClaim"];
                services.AddSingleton<IHttpRequestMessageFactory, TokenMessageFactory>();

                services.AddSingleton(options);
                services.AddTokenAuthenticationService(options);
            }

           // services.AddScoped<IUnit4Service, Unit4Service>();

            return services;
        }

        public static IServiceCollection AddCookieAuthenticationService(this IServiceCollection services, IOptions<EndpointConfig> config)
        {
            services.AddSingleton(config);

            services.AddSingleton<IHttpClientProvider>(new HttpClientProvider(new HttpClientHandler { UseCookies = false }));
            services.AddSingleton<IAuthenticateService<CookieInfo>, AuthenticateCookieService>();
            services.AddSingleton<ITokenService<CookieInfo>, TokenService<CookieInfo>>();
            return services;
        }

    }
}
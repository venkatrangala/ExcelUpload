using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using MMU.Functions.Helpers;
using Microsoft.Extensions.Configuration;
using Mmu.Integration.Common.Utilities.Management.Interfaces;
using Mmu.Integration.Common.Utilities.Data.Interfaces;
using Mmu.Common.Api.Service.Interfaces;
using Microsoft.Extensions.Options;
using Mmu.Common.Api.Service.Models;
using System.Net.Http;

namespace MMU.Functions
{
    public class ProcessExcelFunction
    {
        private readonly IConfiguration _configuration;
        private readonly ILoggerInjector _loggerProvider;
        private readonly IDataService _dataService;
        //private readonly ITokenService<TokenInfo> _tokenService;
        private readonly IHttpRequestMessageFactory _messageFactory;
        private IHttpClientProvider _httpClient;
        //private readonly IOptions<EndPointConfigU4> _config;

        public ProcessExcelFunction(ILoggerInjector loggerProvider, IDataService dataService, IConfiguration configuration,
            IHttpRequestMessageFactory messageFactory,
            ITokenService<TokenInfo> tokenService,
            IOptions<EndPointConfigU4> options,
            IHttpClientProvider httpClientProvider) //IOptions<AppSettings> appSettings, ILogger<ExcelProcessingHelper> logger,
        {
            _dataService = dataService;
            _loggerProvider = loggerProvider;
            _configuration = configuration;
            _messageFactory = messageFactory;
            _httpClient = httpClientProvider;
            //_tokenService = tokenService;
            //_config = options;
            //_appSettings = appSettings.Value;
            //_logger = logger;
        }

        [FunctionName("ProcessExcel")]
        public async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {

            ExcelProcessingHelper excelProcessingHelper = new ExcelProcessingHelper(_loggerProvider, _dataService, _configuration, _messageFactory, _httpClient); //, _tokenService, _config);

            string blobName = "UpdateCourseOfferingCourseOfferingTemplate.xlsx";

            await excelProcessingHelper.ReadFilesFromBlob(blobName);

            log.LogInformation("C# HTTP trigger function processed a request.");

            string name = req.Query["name"];

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);
            name = name ?? data?.name;

            string responseMessage = string.IsNullOrEmpty(name)
                ? "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response."
                : $"Hello, {name}. This HTTP triggered function executed successfully.";

            return new OkObjectResult(responseMessage);
        }
    }
}

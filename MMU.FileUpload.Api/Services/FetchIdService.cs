using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using MMU.FileUpload.Api.Helpers;
using Newtonsoft.Json;
using WebApi.Helpers;

namespace MMU.FileUpload.Api.Services
{
    public interface IFetchIdService
    {
        Task<int> FetchIdFromCourseIdAndAcademicPeriod(string CourseId, string AcademicPeriod);
    }

    public class FetchIdService : IFetchIdService
    {
        private readonly AppSettings _appSettings;
        private readonly HttpClient _httpClient;
        private readonly ILogger<FetchIdService> _logger;

        public FetchIdService(IOptions<AppSettings> appSettings, IHttpClientFactory httpClientFactory, ILogger<FetchIdService> logger)
        {
            _appSettings = appSettings.Value;
            _httpClient = httpClientFactory.CreateClient("FetchIdClient");
            _logger = logger;
        }
        public async Task<int> FetchIdFromCourseIdAndAcademicPeriod(string courseId, string academicPeriod)
        {

            return 123;

            //TODO
            var httpResponse = await _httpClient.GetAsync(_appSettings.BaseUrl + "courseId="+ courseId + "&academicPeriod="+ academicPeriod);

            if (!httpResponse.IsSuccessStatusCode)
            {
                _logger.LogError("invalid request");
                throw new Exception($"Cannot retrieve Id for courseId {courseId} and academicPeriod {academicPeriod}");
            }

            var content = await httpResponse.Content.ReadAsStringAsync();
            var recordId = JsonConvert.DeserializeObject<int>(content);

            return recordId;
        }
    }
}

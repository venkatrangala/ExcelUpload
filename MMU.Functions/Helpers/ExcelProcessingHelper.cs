using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Mmu.Common.Api.Service.Interfaces;
using Mmu.Common.Api.Service.Models;
using Mmu.Integration.Common.Utilities.Data;
using Mmu.Integration.Common.Utilities.Data.Interfaces;
using Mmu.Integration.Common.Utilities.Management.Interfaces;
using MMU.Functions.Models;
using Newtonsoft.Json.Linq;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace MMU.Functions.Helpers
{
    public class ExcelProcessingHelper
    {
        //private readonly AppSettings _appSettings;
        //private readonly ILogger<ExcelProcessingHelper> _logger;
        private readonly ILoggerInjector _loggerProvider;
        private readonly IDataService _dataService;
        private readonly IConfiguration _configuration;
        private readonly HttpClient _httpClient;
        private readonly ITokenService<TokenInfo> _tokenService;
        private readonly IHttpRequestMessageFactory _messageFactory;
        private readonly EndPointConfigU4 _config;
        public ExcelProcessingHelper(ILoggerInjector loggerProvider,
            IDataService dataService, IConfiguration configuration,
            IHttpRequestMessageFactory messageFactory,
            ITokenService<TokenInfo> tokenService,
            IOptions<EndPointConfigU4> options)//IOptions<AppSettings> appSettings, ILogger<ExcelProcessingHelper> logger,
        {
            _dataService = dataService;
            _loggerProvider = loggerProvider;
            _configuration = configuration;
            _messageFactory = messageFactory;
            _tokenService = tokenService;
            _config = options.Value;
            //_appSettings = appSettings.Value;
            //_logger = logger;
        }

        //[HttpGet("ReadFilesFromBlob")]
        //[AllowAnonymous]
        public async Task<IActionResult> ReadFilesFromBlob(string blobName)
        {
            var containerName = "excel";
            //string blobName = "rv.xlsx";
            var azureStorageBlobOptions = new AzureStorageBlobOptions(_configuration);

            //const string folderName = "ExcelUploads";
            var folderPath = Path.Combine(Directory.GetCurrentDirectory());
            var fileEntries = Directory.GetFiles(folderPath);
            var fileName = folderPath + "\\ExcelFile.xlsx";

            var ms = await azureStorageBlobOptions.GetAsync(containerName, blobName);
            ms.Seek(0, SeekOrigin.Begin);

            //Copy the memoryStream from Blob on to local file
            await using (var fs = new FileStream(fileName ?? throw new InvalidOperationException(), FileMode.OpenOrCreate))
            {
                await ms.CopyToAsync(fs);
                fs.Flush();
            }


            UpdateExcelForBlobStorageByName(fileName);

            await using var send = new FileStream(fileName, FileMode.Open, FileAccess.Read);

            await using (var memoryStreamToUpdateBlob = new MemoryStream())
            {
                await send.CopyToAsync(memoryStreamToUpdateBlob);
                memoryStreamToUpdateBlob.Position = 0;
                memoryStreamToUpdateBlob.Seek(0, SeekOrigin.Begin);
                await azureStorageBlobOptions.UpdateFileAsync(memoryStreamToUpdateBlob, containerName, blobName);
            }

            ms.Close();

            return null;
        }

        /// <summary>
        /// Read the rows and update the column data
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public string UpdateExcelForBlobStorageByName(string fileName)
        {
            try
            {
                switch (fileName.ToLower())
                {
                    case "file1":
                        break;
                    default:
                        break;
                }

                //TODO: Where to store the sheetnames
                string sheetName = "COT_Data_input_sheet";
                switch (sheetName.ToLower())
                {
                    case "cot_data_input_sheet":
                        ProcessCoDataInputExcelSheetAsync(fileName, sheetName);
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
            }
            return null;
        }

        private async Task ProcessCoDataInputExcelSheetAsync(string fileName, string sheetName)
        {
            //CourseId
            //CourseTitle
            //StartDate
            //EndDate
            //MinEnrolled
            //MaxEnrolled
            //PriceGroupId
            //CourseLevelId
            //EnrollmentModeId
            //Id
            //Result
            //Reason

            using FileStream rstr = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            IWorkbook workbook = new XSSFWorkbook(rstr);
            var sheet = workbook.GetSheet(sheetName);
            IRow headerRow = sheet.GetRow(1); //Send Row is Header 
            int cellCount = headerRow.LastCellNum;

            for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++) //Values from 2nd row
            {
                IRow row = sheet.GetRow(i);
                if (row == null) continue;
                if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;

                string courseId = string.Empty;
                string courseTitle = string.Empty;
                string startDate = string.Empty;
                string endDate = string.Empty;
                string successValue = string.Empty;
                string invalidDates = string.Empty;
                dynamic errorValue = null;
                int recordId = 0;
                int id;

                Microsoft.AspNetCore.Http.HttpResponse httpResponse;
                HttpResponseMessage httpResponseMessage;
                var updateResult = 0;
                int j = 1; //Leave the first column blank
                //Columns from each row
                for (j = row.FirstCellNum; j < cellCount; j++)
                {
                    if (row.GetCell(j) != null)
                    {
                        var temp1 = new CellReference(row.GetCell(j));
                        var reference = temp1.FormatAsString();
                        ICell cell;

                        //Get the CourseId & AcademicPeriod & fetch RecordID
                        if (reference.StartsWith("B")) //CourseId
                        {
                            courseId = row.GetCell(j).StringCellValue;
                        }
                        //if (reference.StartsWith("C")) //Course Title
                        //{
                        //    courseTitle = row.GetCell(j).StringCellValue;
                        //}

                        //If we got CourseId & courseTitle then fetch RecordID
                        if (!string.IsNullOrEmpty(courseId))
                        {  //TODO: Fetch RecordID
                           //var query = @"Select  Co.Id
                           //     from ACCourseOffering CO
                           //     left join ACAcademicPeriod AP on AP.ID = CO.AcademicPeriodID
                           //     left join ACCourseLevel ACL on CO.CourseLevelID = ACL.Id
                           //     left join BCPriceGroup Price on Co.PriceGroupID = Price. ID
                           //     left join ACCourseOffModes ACM on ACM.CourseOfferingID = CO.Id
                           //     left join ACEnrollmentMode EM on EM.id = ACM.CourseEnrollmentModeID
                           //     Where Co.CourseID  = '" + courseId + "' and AP.BusinessMeaningName = '" + academicPeriod + "'";

                            if (recordId <= 0)
                            {
                                var query = @"Select cot.id from ACCourseofferingTemplate cot
                                            Where cot.courseid ='" + courseId + "'";

                                try
                                {
                                    recordId = _dataService.Query<int>("u4clone", query).FirstOrDefault();
                                }
                                catch (Exception ex)
                                {
                                    //We dont need to handle as we just update the record with no ID
                                }
                            }

                            if (recordId != null && Convert.ToInt32(recordId) > 0)
                            {
                                if (reference.StartsWith("D"))
                                {
                                    startDate = row.GetCell(j).DateCellValue.ToString();
                                }

                                if (reference.StartsWith("E"))
                                {
                                    endDate = row.GetCell(j).DateCellValue.ToString();
                                }

                                //TODO: Validate Dates
                                if (!string.IsNullOrEmpty(startDate) && !string.IsNullOrEmpty(endDate))
                                {
                                    //if (reference.StartsWith("K") )//CourseLevelid
                                    {
                                        bool validDates = ValidateDates(startDate, endDate);

                                        if (validDates)
                                        {
                                            //TODO: Call Api to Update Start & End Dates

                                            try
                                            {
                                                //TODO: What type of response do we get here 
                                                //httpResponse = httpClient call here
                                                var apiUri = new Uri("https://u4sm-preview-mmu.unit4cloud.com/U4SMapi/api/CourseOfferingTemplate/put?id=1");

                                                var payload = new CourseOfferingTemplate
                                                {
                                                    Id = recordId,
                                                    StartDate = Convert.ToDateTime(startDate),
                                                    EndDate = Convert.ToDateTime(endDate)
                                                };

                                                var message = new HttpRequestMessage(HttpMethod.Put, apiUri)
                                                {
                                                    Content = new StringContent(payload.ToString(), Encoding.UTF8, "application/json")
                                                };

                                                var tokenInfo = await _tokenService.GetToken();

                                                message.Headers.Add("Authorization", $"Bearer {tokenInfo.Access_Token}");
                                                message.Headers.Add("unit4_id", _config.Unit4IdClaim);


                                                // var message = await _messageFactory.CreateMessage(HttpMethod.Put, apiUri, payload.ToString());
                                                //Send the request and wait for the response
                                                //var result = await _unit4Service.SendAsync(request.Path, queryPayload.Item1, queryPayload.Item2, request.Path.Contains("update"));

                                                httpResponseMessage = await _httpClient.SendAsync(message);

                                                if (httpResponseMessage.StatusCode.Equals(System.Net.HttpStatusCode.OK))
                                                {
                                                    //TODO: Update success
                                                    successValue = "Success";
                                                }
                                                else //failure
                                                {
                                                    errorValue = httpResponseMessage.Content.ReadAsStringAsync();
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                errorValue = ex;
                                            }
                                        }
                                        else
                                        {
                                            invalidDates = "Invalid Start / End Dates";
                                            successValue = "Fail";
                                        }
                                    }

                                    if (reference.StartsWith("L")) // Result : Success Or Failure
                                    {
                                        using FileStream wstr = new FileStream(fileName, FileMode.Create, FileAccess.Write);
                                        cell = row.GetCell(j);
                                        cell.SetCellValue(successValue != null ? successValue : "Fail");
                                        workbook.Write(wstr);
                                        wstr.Close();
                                    }

                                    if (reference.StartsWith("M") && errorValue != null) // Error Reason
                                    {
                                        using FileStream wstr = new FileStream(fileName, FileMode.Create, FileAccess.Write);
                                        cell = row.GetCell(j);
                                        cell.SetCellValue(errorValue);
                                        workbook.Write(wstr);
                                        wstr.Close();
                                    }
                                }
                            }
                            else
                            {
                                //Log on to Excel and continue
                                //TODO: What needs doing?
                                if (reference.StartsWith("M") && (recordId == 0)) // Error Reason
                                {
                                    using FileStream noIDStream = new FileStream(fileName, FileMode.Create, FileAccess.Write);
                                    cell = row.GetCell(j);
                                    cell.SetCellValue("ID does not exist");
                                    workbook.Write(noIDStream);
                                    noIDStream.Close();
                                }
                            }
                        }
                    }

                    rstr.Close();
                }
            }
        }

        private bool ValidateDates(string startDate, string endDate)
        {
            return DateTime.TryParse(startDate, out DateTime _) == true &&
                         (DateTime.TryParse(endDate, out DateTime _) == true);

        }
    }
}

using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Mmu.Common.Api.Service.Interfaces;
using Mmu.Integration.Common.Utilities.Data.Interfaces;
using Mmu.Integration.Common.Utilities.Management.Interfaces;
using MMU.Functions.Models;
using Newtonsoft.Json;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
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
        private HttpClient _httpClient;
        //private readonly ITokenService<TokenInfo> _tokenService;
        private readonly IHttpRequestMessageFactory _messageFactory;
        //private readonly EndPointConfigU4 _config;
        public ExcelProcessingHelper(ILoggerInjector loggerProvider,
            IDataService dataService, IConfiguration configuration,
            IHttpRequestMessageFactory messageFactory,
            IHttpClientProvider httpClientProvider
            //ITokenService<TokenInfo> tokenService,
            //IOptions<EndPointConfigU4> options
            )//IOptions<AppSettings> appSettings, ILogger<ExcelProcessingHelper> logger,
        {
            _dataService = dataService;
            _loggerProvider = loggerProvider;
            _configuration = configuration;
            _messageFactory = messageFactory;
            _httpClient = httpClientProvider.HttpClient;
            //_tokenService = tokenService;
            //_config = options.Value;
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
            //var fileEntries = Directory.GetFiles(folderPath);
            var fileName = folderPath + "\\ExcelFile.xlsx";

            var ms = await azureStorageBlobOptions.GetAsync(containerName, blobName);
            ms.Seek(0, SeekOrigin.Begin);

            //Copy the memoryStream from Blob on to local file
            await using (var fs = new FileStream(fileName ?? throw new InvalidOperationException(), FileMode.OpenOrCreate))
            {
                await ms.CopyToAsync(fs);
                fs.Flush();
            }

            await UpdateExcelForBlobStorageByNameAsync(fileName);

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
        public async Task<string> UpdateExcelForBlobStorageByNameAsync(string fileName)
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
                        await ProcessCoDataInputExcelSheetAsync(fileName, sheetName);
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
            //TODO: Fetch using DataService
            //Select BusinessMeaningName, id PriceGroupId from BCPriceGroup
            //Select BusinessMeaningName,id Courselevelid from ACCourselevel
            //Select BusinessMeaningName,id EnrollmentmodeId from ACEnrollmentMode

            var priceGroupQuery = @"Select Id, BusinessMeaningName from BCPriceGroup";
            var courseLevelQuery = @"Select Id, BusinessMeaningName from ACCourselevel";
            var enrollmentModelQuery = @"Select Id, BusinessMeaningName from ACEnrollmentMode";

            var priceGroupList = new List<BusinessNames>();
            var courseLevelList = new List<BusinessNames>();
            var enrollmentModelList = new List<BusinessNames>();

            try
            {
                //recordId = _dataService.Query<int>("u4clone", query).FirstOrDefault();
                priceGroupList = _dataService.Query<BusinessNames>("u4clone", priceGroupQuery).ToList();
                courseLevelList = _dataService.Query<BusinessNames>("u4clone", courseLevelQuery).ToList();
                enrollmentModelList = _dataService.Query<BusinessNames>("u4clone", enrollmentModelQuery).ToList();
            }
            catch (Exception ex)
            {
                throw;
                //We dont need to handle as we just update the record with no ID
            }


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
            IRow headerRow = sheet.GetRow(1); //Header 
            int cellCount = headerRow.LastCellNum;

            string courseIdColumn = "B";
            string MinEnrolledColumn = "D";
            string MaxEnrolledColumn = "E";
            string PriceGroupIdColumn = "F";
            string CourseLevelIdColumn = "G";
            string EnrollmentModeIdColumn = "H";
            string RecordIdColumn = "I";
            string ResultColumn = "J";
            string ErrorColumn = "K";

            for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++) //Values from 2nd row
            {
                IRow row = sheet.GetRow(i);
                if (row == null) continue;
                if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;

                string courseId = string.Empty;
                //string courseTitle = string.Empty;
                //string startDate = string.Empty;
                //string endDate = string.Empty;
                string successValue = string.Empty;
                string invalidDates = string.Empty;
                dynamic errorValue = null;
                int recordId = 0;
                //int id;

                var payload = new CourseOfferingTemplate();
                HttpResponseMessage httpResponseMessage;
                //var updateResult = 0;

                int j = 1; //Leave the first column blank
                //Columns from each row
                for (j = row.FirstCellNum; j < cellCount; j++)
                {
                    if (row.GetCell(j) != null)
                    {
                        //var temp1 = new CellReference(row.GetCell(j));
                        //var reference = temp1.FormatAsString();
                        var reference = new CellReference(row.GetCell(j)).FormatAsString();
                        ICell cell;

                        //Get the CourseId & fetch RecordID
                        if (reference.StartsWith(courseIdColumn)) //CourseId
                        {
                            courseId = row.GetCell(j).StringCellValue.Trim();
                        }

                        //if (reference.StartsWith("C")) //Course Title
                        //{
                        //    courseTitle = row.GetCell(j).StringCellValue;
                        //}

                        //If we got CourseId then fetch RecordID
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
                                var query = $"Select cot.id from ACCourseofferingTemplate cot Where cot.courseid ='{courseId}'";

                                try
                                {
                                    recordId = _dataService.Query<int>("u4clone", query).FirstOrDefault();
                                    //recordId = 2;
                                    payload.Id = recordId;
                                }
                                catch (Exception ex)
                                {
                                    //We dont need to handle as we just update the record with no ID
                                }
                            }

                            if (recordId > 0)
                            {
                                //if (reference.StartsWith("D"))
                                //{
                                //    startDate = row.GetCell(j).DateCellValue.ToString();
                                //}

                                //if (reference.StartsWith("E"))
                                //{
                                //    endDate = row.GetCell(j).DateCellValue.ToString();
                                //}

                                if (string.IsNullOrEmpty(successValue) && string.IsNullOrEmpty(errorValue))
                                {
                                    if (reference.StartsWith(MinEnrolledColumn)) //MinEnrolled
                                    {
                                        payload.MinEnrolled = Convert.ToInt32(row.GetCell(j).StringCellValue);
                                    }
                                    if (reference.StartsWith(MaxEnrolledColumn)) //MaxEnrolled
                                    {
                                        payload.MaxEnrolled = Convert.ToInt32(row.GetCell(j).StringCellValue);
                                    }
                                    if (reference.StartsWith(PriceGroupIdColumn)) //PriceGroupId
                                    {
                                        //TODO:
                                        //payload.PriceGroupId = Convert.ToInt32(row.GetCell(j).StringCellValue);
                                        payload.PriceGroupId = Convert.ToInt32(priceGroupList.Where(x => x.BusinessMeaningName == row.GetCell(j).StringCellValue).Select(y => y.Id));
                                    }
                                    if (reference.StartsWith(CourseLevelIdColumn)) //CourseLevelId
                                    {
                                        //TODO:
                                        //payload.CourseLevelId = Convert.ToInt32(row.GetCell(j).StringCellValue);
                                        payload.CourseLevelId = Convert.ToInt32(courseLevelList.Where(x => x.BusinessMeaningName == row.GetCell(j).StringCellValue).Select(y => y.Id));
                                    }
                                    if (reference.StartsWith(EnrollmentModeIdColumn)) //EnrollmentModeId
                                    {
                                        //TODO:
                                        //payload.EnrollmentModeId = Convert.ToInt32(row.GetCell(j).StringCellValue);
                                        payload.EnrollmentModeId = Convert.ToInt32(enrollmentModelList.Where(x => x.BusinessMeaningName == row.GetCell(j).StringCellValue).Select(y => y.Id));
                                    }

                                    //bool validDates = ValidateDates(startDate, endDate);

                                    //if (validDates)

                                    //Call Api to Update Start & End Dates
                                    try
                                    {
                                        //TODO
                                        //var apiUri = new Uri("https://u4sm-preview-mmu.unit4cloud.com/U4SMapi/api/CourseOfferingTemplate/put?id=1");

                                        var apiUri = new Uri("https://u4sm-preview-mmu.unit4cloud.com/U4SMapi/api/CourseOfferingTemplate/put?id=" + recordId);
                                        //https://u4sm-accept05-mmu-sit.unit4cloud.com/U4SMapi/api/CourseOffering/get?id=1

                                        //var payload = new CourseOfferingTemplate
                                        //{
                                        //    Id = recordId,
                                        //    MinEnrolled = 1,
                                        //    MaxEnrolled = 1,
                                        //    PriceGroupId = 1,
                                        //    CourseLevelId = 1,
                                        //    EnrollmentModeId = 1
                                        //    //StartDate = Convert.ToDateTime(startDate),
                                        //    //EndDate = Convert.ToDateTime(endDate)
                                        //};

                                        var payloadString = JsonConvert.SerializeObject(payload);
                                        //TODO: Uncomment for live
                                        var message = await _messageFactory.CreateMessage(HttpMethod.Put, apiUri, payloadString);
                                        httpResponseMessage = await _httpClient.SendAsync(message);
                                        if (httpResponseMessage.StatusCode.Equals(System.Net.HttpStatusCode.OK))
                                        //if (1 == 1)
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

                                    //else
                                    //{
                                    //    invalidDates = "Invalid Start / End Dates";
                                    //    successValue = "Fail";
                                    //}
                                }

                                if (reference.StartsWith(RecordIdColumn)) // RecordId
                                {
                                    using FileStream wstr = new FileStream(fileName, FileMode.Create, FileAccess.Write);
                                    cell = row.GetCell(j);
                                    cell.SetCellValue(recordId);
                                    workbook.Write(wstr);
                                    wstr.Close();
                                }

                                if (reference.StartsWith(ResultColumn)) // Result : Success Or Failure
                                {
                                    using FileStream wstr = new FileStream(fileName, FileMode.Create, FileAccess.Write);
                                    cell = row.GetCell(j);
                                    cell.SetCellValue(successValue != null ? successValue : "Fail");
                                    workbook.Write(wstr);
                                    wstr.Close();
                                }

                                if (reference.StartsWith(ErrorColumn) && errorValue != null) // Error Reason
                                {
                                    using FileStream wstr = new FileStream(fileName, FileMode.Create, FileAccess.Write);
                                    cell = row.GetCell(j);
                                    cell.SetCellValue(errorValue);
                                    workbook.Write(wstr);
                                    wstr.Close();
                                }
                            }
                            else
                            {
                                //Log error to Excel and continue
                                if (reference.StartsWith(ErrorColumn) && (recordId == 0)) // Error Reason
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
                }
            }
            rstr.Close();
        }

        private bool ValidateDates(string startDate, string endDate)
        {
            return DateTime.TryParse(startDate, out DateTime _) == true &&
                         (DateTime.TryParse(endDate, out DateTime _) == true);

        }
    }
}

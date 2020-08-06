using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Mmu.Integration.Common.Utilities.Data;
using Mmu.Integration.Common.Utilities.Data.Interfaces;
using Mmu.Integration.Common.Utilities.Management.Interfaces;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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

        public ExcelProcessingHelper(ILoggerInjector loggerProvider,
            IDataService dataService, IConfiguration configuration) //IOptions<AppSettings> appSettings, ILogger<ExcelProcessingHelper> logger,
        {
            _dataService = dataService;
            _loggerProvider = loggerProvider;
            _configuration = configuration;
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
            string sheetName = "CO_Data_input_sheet";
            using FileStream rstr = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            IWorkbook workbook = new XSSFWorkbook(rstr);
            var sheet = workbook.GetSheet(sheetName);
            IRow headerRow = sheet.GetRow(0);
            int cellCount = headerRow.LastCellNum;
            

            for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row == null) continue;
                if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;

                string courseId = string.Empty;
                string academicPeriod = string.Empty;
                int id;

                
                int j = 0;

                for (j = row.FirstCellNum; j < cellCount; j++)
                {
                    if (row.GetCell(j) != null)
                    {
                        //if (!string.IsNullOrEmpty(row.GetCell(j).ToString()) & !string.IsNullOrWhiteSpace(row.GetCell(j).ToString()))
                        {
                            var temp1 = new CellReference(row.GetCell(j));
                            var reference = temp1.FormatAsString();
                            ICell cell;

                            //Get the CourseId & AcademicPeriod & fetch RecordID
                            if (reference.StartsWith("A")) //CourseId
                            {
                                courseId = row.GetCell(j).StringCellValue;
                            }
                            if (reference.StartsWith("B")) //AcademicPeriod
                            {
                                academicPeriod = row.GetCell(j).StringCellValue;
                            }

                            //If we got CourseId & AcademicPeriod then fetch RecordID
                            if (!string.IsNullOrEmpty(courseId) && !string.IsNullOrEmpty(academicPeriod))
                            {  //TODO: Fetch RecordID
                                //var query = @"Select  Co.Id
                                //     from ACCourseOffering CO
                                //     left join ACAcademicPeriod AP on AP.ID = CO.AcademicPeriodID
                                //     left join ACCourseLevel ACL on CO.CourseLevelID = ACL.Id
                                //     left join BCPriceGroup Price on Co.PriceGroupID = Price. ID
                                //     left join ACCourseOffModes ACM on ACM.CourseOfferingID = CO.Id
                                //     left join ACEnrollmentMode EM on EM.id = ACM.CourseEnrollmentModeID
                                //     Where Co.CourseID  = '" + courseId + "' and AP.BusinessMeaningName = '" + academicPeriod + "'";

                                //var id = _dataService.Query<int>("vertical_replica_preview", query);
                                id = 101;

                                if (id != null && id > 0)
                                {


                                    if (reference.StartsWith("D"))
                                    {
                                        using FileStream wstr = new FileStream(fileName, FileMode.Create, FileAccess.Write);
                                        cell = row.GetCell(j);
                                        cell.SetCellValue(DateTime.Now.ToShortDateString());
                                        workbook.Write(wstr);
                                        wstr.Close();
                                    }
                                    if (reference.StartsWith("E"))
                                    {
                                        using FileStream wstr = new FileStream(fileName, FileMode.Create, FileAccess.Write);
                                        cell = row.GetCell(j);
                                        cell.SetCellValue(DateTime.Now.AddYears(1).ToShortDateString());
                                        workbook.Write(wstr);
                                        wstr.Close();
                                    }
                                    if (reference.StartsWith("I"))//CourseLevelid
                                    {
                                        using FileStream wstr = new FileStream(fileName, FileMode.Create, FileAccess.Write);
                                        cell = row.GetCell(j);
                                        cell.SetCellValue(id);
                                        workbook.Write(wstr);
                                        wstr.Close();
                                    }
                                }

                                //if (reference.StartsWith("D"))
                                //{
                                //    using FileStream wstr = new FileStream(fileName, FileMode.Create, FileAccess.Write);
                                //    cell = row.GetCell(j);
                                //    cell.SetCellValue(DateTime.Now.ToShortDateString());
                                //    workbook.Write(wstr);
                                //    wstr.Close();
                                //}
                                //if (reference.StartsWith("E"))
                                //{
                                //    using FileStream wstr = new FileStream(fileName, FileMode.Create, FileAccess.Write);
                                //    cell = row.GetCell(j);
                                //    cell.SetCellValue(DateTime.Now.AddYears(1).ToShortDateString());
                                //    workbook.Write(wstr);
                                //    wstr.Close();
                                //}


                            }
                            else
                            {
                             //Log on to Excel and continue
                                //TODO: What needs doing?
                            }

                        }
                    }
                }

                
            }

            
            rstr.Close();

            return null;
        }

    }
}

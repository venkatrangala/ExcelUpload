using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
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
        private readonly AppSettings _appSettings;
        private readonly ILogger<ExcelProcessingHelper> _logger;
        private readonly IConfiguration _configuration;

        public ExcelProcessingHelper(IOptions<AppSettings> appSettings, ILogger<ExcelProcessingHelper> logger, IConfiguration configuration)
        {
            _configuration = configuration;
            _appSettings = appSettings.Value;
            _logger = logger;
        }

        //[HttpGet("ReadFilesFromBlob")]
        //[AllowAnonymous]
        public async Task<IActionResult> ReadFilesFromBlob(string blobName)
        {
            var containerName = "excel";
            //string blobName = "rv.xlsx";
            var azureStorageBlobOptions = new AzureStorageBlobOptions(_configuration);

            const string folderName = "ExcelUploads";
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), folderName);
            var fileEntries = Directory.GetFiles(folderPath);
            var fileName = fileEntries.FirstOrDefault();

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
        public static string UpdateExcelForBlobStorageByName(string fileName)
        {
            //IRow row;
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
                for (int j = row.FirstCellNum; j < cellCount; j++)
                {
                    if (row.GetCell(j) != null)
                    {
                        if (!string.IsNullOrEmpty(row.GetCell(j).ToString()) & !string.IsNullOrWhiteSpace(row.GetCell(j).ToString()))
                        {
                            var temp1 = new CellReference(row.GetCell(j));
                            var reference = temp1.FormatAsString();
                            ICell cell;
                            //Get the CourseId & AcademicPeriod & fetch RecordID
                            string courseId = string.Empty;
                            string academicPeriod = string.Empty;
                            if (reference.StartsWith("A")) //CourseId
                            {
                                courseId = row.GetCell(j).StringCellValue;
                            }
                            if (reference.StartsWith("B")) //AcademicPeriod
                            {
                                academicPeriod = row.GetCell(j).StringCellValue;
                            }

                            //If we got CourseId & AcademicPeriod then fetch RecordID
                            //TODO: Fetch RecordID

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
                                cell.SetCellValue("Dev");
                                workbook.Write(wstr);
                                wstr.Close();
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

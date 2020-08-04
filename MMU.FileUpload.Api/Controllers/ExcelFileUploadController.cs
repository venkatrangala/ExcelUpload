using System;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using AutoMapper;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using MMU.FileUpload.Api.Helpers;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using WebApi.Helpers;
using WebApi.Services;

namespace MMU.FileUpload.Api.Controllers
{
    //[Authorize]
    [ApiController]
    [Route("api/[controller]")]
    public class ExcelFileUploadController : ControllerBase
    {
        private IFileUploadService _fileUploadService;
        private IMapper _mapper;
        private readonly AppSettings _appSettings;
        private readonly ILogger<ExcelFileUploadController> _logger;
        private readonly IConfiguration _configuration;
        public ExcelFileUploadController(
            IFileUploadService fileUploadService,
            IMapper mapper,
            IOptions<AppSettings> appSettings, ILogger<ExcelFileUploadController> logger, IConfiguration configuration)
        {
            _configuration = configuration;
            _fileUploadService = fileUploadService;
            _mapper = mapper;
            _appSettings = appSettings.Value;
            _logger = logger;
        }

        [HttpPost("UploadFilesToBlob"), DisableRequestSizeLimit]
        [AllowAnonymous]
        public async Task<IActionResult> UploadFilesToBlob() //Blob Storage
        {
            try
            {
                string containerName = "excel";
                var azureStorageBlobOptions = new AzureStorageBlobOptions(_configuration);

                var files = Request.Form.Files;
                await azureStorageBlobOptions.UploadFileAsync(files.FirstOrDefault(), containerName);

                return Ok("File uploaded successfully.");
            }
            catch (Exception ex)
            {
                return StatusCode(500, "Internal server error");
            }
        }

        [HttpGet("ReadFilesFromBlob")]
        [AllowAnonymous]
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
            
            return Ok();
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


        //LOCAL STORAGE CODE FROM HERE BELOW

        [HttpGet("ReadFiles")]
        [AllowAnonymous]
        public IActionResult ReadFiles()
        {
            const string folderName = "ExcelUploads";
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), folderName);

            var fileEntries = Directory.GetFiles(folderPath);

            foreach (var fileName in fileEntries)
            {
                if (fileName.Length > 0)//ProcessFiles
                {
                    ReadExcelByName(fileName);
                }
            }
            return Ok();
        }

        [HttpPost("UploadFiles"), DisableRequestSizeLimit]
        [AllowAnonymous]
        public async Task<IActionResult> UploadFiles() //Local Storage
        {
            try
            {
                var files = Request.Form.Files;

                //Local Directory File Upload below
                const string folderName = "ExcelUploads";

                var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);

                if (files.Any(f => f.Length == 0))
                {
                    return BadRequest();
                }

                foreach (var file in files)
                {
                    var fileName = ContentDispositionHeaderValue.Parse(file.ContentDisposition).FileName.Trim('"');
                    var fullPath = Path.Combine(pathToSave, fileName);
                    var dbPath = Path.Combine(folderName, fileName);

                    await using var stream = new FileStream(fullPath, FileMode.Create);
                    await file.CopyToAsync(stream);
                }

                return Ok("File uploaded successfully.");
            }
            catch (Exception ex)
            {
                return StatusCode(500, "Internal server error");
            }
        }

        /// <summary>
        /// Read the rows and update the column data
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static string ReadExcelByName(string fileName)
        {
            IWorkbook workbook;
            ISheet sheet;
            //IRow row;
            ICell cell;
            string sheetName = "CO_Data_input_sheet";
            using FileStream rstr = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            workbook = new XSSFWorkbook(rstr);
            sheet = workbook.GetSheet(sheetName);
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
                        if (!string.IsNullOrEmpty(row.GetCell(j).ToString()) & ((!string.IsNullOrWhiteSpace(row.GetCell(j).ToString()))))
                        {
                            var temp1 = new CellReference(row.GetCell(j));
                            var reference = temp1.FormatAsString();
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

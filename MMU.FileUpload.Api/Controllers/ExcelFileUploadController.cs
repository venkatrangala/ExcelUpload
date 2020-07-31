using System;
using System.Collections.Generic;
using Microsoft.AspNetCore.Mvc;
using AutoMapper;
using System.IdentityModel.Tokens.Jwt;
using WebApi.Helpers;
using Microsoft.Extensions.Options;
using System.Text;
using Microsoft.IdentityModel.Tokens;
using System.Security.Claims;
using Microsoft.AspNetCore.Authorization;
using WebApi.Services;
using WebApi.Models;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Logging;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.Linq;
using System.Net.Http.Headers;
using Newtonsoft.Json;
using System.Data;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using MMU.FileUpload.Api.Helpers;
using NPOI.SS.Util;

namespace WebApi.Controllers
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

        [HttpPost("UploadFiles"), DisableRequestSizeLimit]
        [AllowAnonymous]
        public async Task<IActionResult> UploadFiles() //Blob Storage
        {
            try
            {
                AzureStorageBlobOptions azureStorageBlobOptions = new AzureStorageBlobOptions(_configuration);

                var files = Request.Form.Files;
                await azureStorageBlobOptions.UploadFileAsync(files.FirstOrDefault());
                
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

        [HttpGet("ReadFilesFromBlob")]
        [AllowAnonymous]
        public async Task<IActionResult> ReadFilesFromBlob()
        {

            AzureStorageBlobOptions azureStorageBlobOptions = new AzureStorageBlobOptions(_configuration);

            var fileName = "rv.xlsx";
            var fileMemoryStream = await azureStorageBlobOptions.GetAsync(fileName);

            //return fileMemoryStream;


            //FileStream rstr = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            //fileMemoryStream.CopyTo(rstr);
            var ms = new NpoiMemoryStream();
            ms.AllowClose = false;
            fileMemoryStream.CopyTo(ms);

            IWorkbook workbook;
            ISheet sheet;
            //IRow row;
            ICell cell;
            string sheetName = "CO_Data_input_sheet";
            //using FileStream rstr = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            //https://peter.grman.at/handling-excel-files-in-web-applications/
            fileMemoryStream.Position = 0;
            
            MemoryStream temp = new MemoryStream();
            fileMemoryStream.CopyTo(temp);
            using MemoryStream rstr = fileMemoryStream;
            //using (MemoryStream rstr = fileMemoryStream)
            {
                workbook = new XSSFWorkbook();
                workbook.Write(rstr);
                //rstr = new MemoryStream();
                //temp.CopyTo(rstr);
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
                            if (!string.IsNullOrEmpty(row.GetCell(j).ToString()) &
                                ((!string.IsNullOrWhiteSpace(row.GetCell(j).ToString()))))
                            {
                                var temp1 = new CellReference(row.GetCell(j));
                                var reference = temp1.FormatAsString();
                                if (reference.StartsWith("D"))
                                {
                                    //using FileStream wstr = new FileStream(fileName, FileMode.Create, FileAccess.Write);
                                    //MemoryStream wstr = fileMemoryStream;
                                    cell = row.GetCell(j);
                                    cell.SetCellValue(DateTime.Now.ToShortDateString());
                                    workbook.Write(rstr);
                                    //wstr.Close();
                                }

                                if (reference.StartsWith("E"))
                                {
                                    //using FileStream wstr = new FileStream(fileName, FileMode.Create, FileAccess.Write);
                                    //MemoryStream wstr = fileMemoryStream;
                                    cell = row.GetCell(j);
                                    cell.SetCellValue("Dev");
                                    workbook.Write(rstr);
                                    //wstr.Close();
                                }
                            }
                        }
                    }
                }
                
                fileMemoryStream.Position = 0;
                await azureStorageBlobOptions.UpdateFileAsync(fileMemoryStream);
                rstr.Close();
            }

            return null;

            //const string folderName = "ExcelUploads";
            //var folderPath = Path.Combine(Directory.GetCurrentDirectory(), folderName);

            //var fileEntries = Directory.GetFiles(folderPath);

            //foreach (var fileName in fileEntries)
            //{
            //    if (fileName.Length > 0)//ProcessFiles
            //    {
            //        ReadExcelByName(fileName);
            //    }
            //}
            return Ok();
        }

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

        static string ReadExcelByNameBkp(string fileName)
        {
            DataTable dtTable = new DataTable();
            List<string> rowList = new List<string>();
            ISheet sheet;
            using (var stream = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite))
            {
                stream.Position = 0;
                XSSFWorkbook xssWorkbook = new XSSFWorkbook(stream);
                //sheet = xssWorkbook.GetSheet("CO_Data_input_Sheet");
                sheet = xssWorkbook.GetSheet("COT_Data_input_sheet");
                IRow headerRow = sheet.GetRow(0);
                int cellCount = headerRow.LastCellNum;
                for (int j = 0; j < cellCount; j++)
                {
                    ICell cell = headerRow.GetCell(j);
                    if (cell == null || string.IsNullOrWhiteSpace(cell.ToString())) continue;
                    {
                        dtTable.Columns.Add(cell.ToString());
                    }
                }
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
                                rowList.Add(row.GetCell(j).ToString());
                            }
                        }
                    }

                    if (rowList.Count > 0)
                        dtTable.Rows.Add(rowList.ToArray());
                    rowList.Clear();
                }
            }

            return JsonConvert.SerializeObject(dtTable);
        }

        static string ReadExcel(string fileName)
        {
            DataTable dtTable = new DataTable();
            List<string> rowList = new List<string>();
            ISheet sheet;
            using (var stream = new FileStream(fileName, FileMode.Open))
            {
                stream.Position = 0;
                XSSFWorkbook xssWorkbook = new XSSFWorkbook(stream);
                sheet = xssWorkbook.GetSheetAt(0);
                IRow headerRow = sheet.GetRow(0);
                int cellCount = headerRow.LastCellNum;
                for (int j = 0; j < cellCount; j++)
                {
                    ICell cell = headerRow.GetCell(j);
                    if (cell == null || string.IsNullOrWhiteSpace(cell.ToString())) continue;
                    {
                        dtTable.Columns.Add(cell.ToString());
                    }
                }
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
                                rowList.Add(row.GetCell(j).ToString());
                            }
                        }
                    }

                    if (rowList.Count > 0)
                        dtTable.Rows.Add(rowList.ToArray());
                    rowList.Clear();
                }
            }

            return JsonConvert.SerializeObject(dtTable);
        }
    }


    public class NpoiMemoryStream : MemoryStream
    {
        public NpoiMemoryStream()
        {
            // We always want to close streams by default to
            // force the developer to make the conscious decision
            // to disable it.  Then, they're more apt to remember
            // to re-enable it.  The last thing you want is to
            // enable memory leaks by default.  ;-)
            AllowClose = true;
        }

        public bool AllowClose { get; set; }

        public override void Close()
        {
            if (AllowClose)
                base.Close();
        }
    }

}

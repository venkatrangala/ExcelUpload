using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using Microsoft.WindowsAzure.Storage.RetryPolicies;

namespace MMU.FileUpload.Api.Helpers
{
    public class AzureStorageBlobOptions
    {
        private readonly IConfiguration _configuration;

        public AzureStorageBlobOptions(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        public async Task<IActionResult> UploadFileAsync(IFormFile file)
        {
            var cloudStorageAccount =
                CloudStorageAccount.Parse("UseDevelopmentStorage=true;");
            //_configuration["AzureStorage:ConnectionString"]);

            var cloudBlobClient =
                cloudStorageAccount.CreateCloudBlobClient();
            string containerName = "excel"; // + Guid.NewGuid();

            // Create a container for organizing blobs within the storage account.
            Console.WriteLine("1. Creating Container");
            CloudBlobContainer cloudBlobContainer = cloudBlobClient.GetContainerReference(containerName);
            //var cloudBlobContainer = cloudBlobClient.GetContainerReference(containerName);
            try
            {
                // The call below will fail if the sample is configured to use the storage emulator in the connection string, but 
                // the emulator is not running.
                // Change the retry policy for this call so that if it fails, it fails quickly.
                BlobRequestOptions requestOptions = new BlobRequestOptions() { RetryPolicy = new NoRetry() };
                await cloudBlobContainer.CreateIfNotExistsAsync(requestOptions, null);
            }
            catch (StorageException)
            {
                Console.WriteLine("If you are running with the default connection string, please make sure you have started the storage emulator. Press the Windows key and type Azure Storage to select and run it from the list of applications - then restart the sample.");
                Console.ReadLine();
                throw;
            }

            //_configuration["AzureStorage:FilePath"]);
            var blobName = file.FileName.ToLower();
            blobName = blobName.Replace("\"", "");

            var cloudBlockBlob =
                cloudBlobContainer.GetBlockBlobReference(blobName);

            var temp = file.ContentType;
            cloudBlockBlob.Properties.ContentType = file.ContentType;

            try
            {
                await using var fileStream = file.OpenReadStream();
                await cloudBlockBlob.UploadFromStreamAsync(fileStream);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }


            return new OkObjectResult(new { name = blobName });
        }


        public async Task<IActionResult> UpdateFileAsync(MemoryStream memoryStream)
        {
            var cloudStorageAccount =
                CloudStorageAccount.Parse("UseDevelopmentStorage=true;");
            //_configuration["AzureStorage:ConnectionString"]);

            var cloudBlobClient =
                cloudStorageAccount.CreateCloudBlobClient();
            string containerName = "excel"; // + Guid.NewGuid();

            // Create a container for organizing blobs within the storage account.
            Console.WriteLine("1. Creating Container");
            CloudBlobContainer cloudBlobContainer = cloudBlobClient.GetContainerReference(containerName);
            //var cloudBlobContainer = cloudBlobClient.GetContainerReference(containerName);
            try
            {
                // The call below will fail if the sample is configured to use the storage emulator in the connection string, but 
                // the emulator is not running.
                // Change the retry policy for this call so that if it fails, it fails quickly.
                BlobRequestOptions requestOptions = new BlobRequestOptions() { RetryPolicy = new NoRetry() };
                await cloudBlobContainer.CreateIfNotExistsAsync(requestOptions, null);
            }
            catch (StorageException)
            {
                Console.WriteLine("If you are running with the default connection string, please make sure you have started the storage emulator. Press the Windows key and type Azure Storage to select and run it from the list of applications - then restart the sample.");
                Console.ReadLine();
                throw;
            }

            //_configuration["AzureStorage:FilePath"]);
            var blobName = "rv.xlsx";
            
            var cloudBlockBlob =
                cloudBlobContainer.GetBlockBlobReference(blobName);

            cloudBlockBlob.Properties.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

            try
            {
                // Create or overwrite the "myblob" blob with contents from a local file.
                memoryStream.Position = 0;
                using (var fileStream = memoryStream)
                {
                    cloudBlockBlob.UploadFromStreamAsync(fileStream);
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }


            return new OkObjectResult(new { name = blobName });
        }

        public async Task<MemoryStream> GetAsync(
            string name)
        {
            var cloudStorageAccount =
                CloudStorageAccount.Parse("UseDevelopmentStorage=true;");
            //CloudStorageAccount.Parse(
                    //_configuration["AzureStorage:ConnectionString"]);

            var cloudBlobClient =
                cloudStorageAccount.CreateCloudBlobClient();

            string containerName = "excel"; // + Guid.NewGuid();

            CloudBlobContainer cloudBlobContainer = cloudBlobClient.GetContainerReference(containerName);
            
            var blobName = name;

            var cloudBlockBlob =
                cloudBlobContainer.GetBlockBlobReference(blobName);

            var ms = new MemoryStream();
            await cloudBlockBlob.DownloadToStreamAsync(ms);
            ms.Seek(0, SeekOrigin.Begin);
            //var ms = new MemoryStream();
            //await cloudBlockBlob.DownloadToStreamAsync(ms);
            //https://stackoverflow.com/questions/8624071/save-and-load-memorystream-to-from-a-file
            //using (FileStream file = new FileStream("file.xlsx", FileMode.Create, System.IO.FileAccess.Write))
            //{
            //    byte[] bytes = new byte[ms.Length];
            //    ms.Read(bytes, 0, (int)ms.Length);
            //    file.Write(bytes, 0, bytes.Length);
            //    ms.Close();
            //}

            return ms;
            //return new FileContentResult(ms.ToArray(), cloudBlockBlob.Properties.ContentType);
        }
    }
}
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;

namespace MMU.Functions.Helpers
{
    public static class HttpContentExtensions
    {
        /// <summary>
        /// Clones and closes the HttpContent Stream
        /// </summary>
        /// <param name="content">The HttpContent to clone</param>
        /// <returns>The contents of the response positioned at the start</returns>
        public static async Task<Stream> CopyContentAsync(this HttpContent content)
        {
            var clonedStream = new MemoryStream();
            var contentStream = await content.ReadAsStreamAsync();
            await contentStream.CopyToAsync(clonedStream);
            clonedStream.Position = 0;
            contentStream.Close();
            return clonedStream;
        }
    }

}
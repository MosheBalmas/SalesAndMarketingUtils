using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;

namespace SalesAndMarketingUtilsServices
{
    public class FileUploadingController : ApiController
    {
        [HttpPost]
        [Route("api/FileUploading/UploadFile")]
        public async Task<Tuple <string,string>> UploadFile()
        {
            var ctx = HttpContext.Current;
            var root = ctx.Server.MapPath("~/App_Data");
            var provider = new MultipartFormDataStreamProvider(root);
            string filePath = "";

            byte[] resFile;
            try
            {
                await Request.Content.ReadAsMultipartAsync(provider);
               
                foreach (var file in provider.FileData)
                {
                    var name = file.Headers
                        .ContentDisposition
                        .FileName;

                    name = name.Trim('"');

                    var localFileName = file.LocalFileName;
                    filePath = Path.Combine(root, name);

                    File.Move(localFileName, filePath);

                    

                }

                return new Tuple<string, string>(filePath, "File uploaded!");


            }
            catch (Exception e)
            {
                return  new Tuple<string, string>("",  $"Error: {e.Message}");
            }

            
        }
    }
}
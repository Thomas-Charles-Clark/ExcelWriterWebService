using ExcelWriterWebService.Models;
using ExcelWriterWebService.Utilities;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web.Http;

namespace ExcelWriterWebService.Controllers
{
    public class ExcelController : ApiController
    {
        [Route("api/excel/CreateFile")]
        [HttpPost]
        public IHttpActionResult CreateFile([FromBody]PassExcelCreation objCreation)
        {
            try
            {
                if (Framework.WriteFile(
                            FileName: objCreation.FileName,
                            FileType: objCreation.FileType,
                            Destination: objCreation.Destination,
                            Contents: objCreation.Contents ?? Framework.FetchWvAJSON(),
                            UppercaseHeadings: objCreation.UppercaseHeadings,
                            AutoSizeColumns: objCreation.AutoSizeColumns,
                            HeadingStyleID: objCreation.HeadingStyleID,
                            StandardStyleID: objCreation.StandardStyleID,
                            IndentedHeadingStyleID: objCreation.IndentedHeadingStyleID,
                            IndentedStyleID: objCreation.IndentedStyleID))
                {
                    return Ok(new { objCreation.FileName });
                }
                else
                {
                    throw new Exception($"FAILED writing {objCreation.FileName}");
                }
            }
            catch (Exception e)
            {
                return Content(HttpStatusCode.InternalServerError, e);
            }
        }

        [Route("api/excel/GetFileStream")]
        [HttpPost]
        public HttpResponseMessage GetFileStream([FromBody]PassExcelCreation objCreation)
        {
            var bytes = Framework.WriteFileStream(
                    FileName: objCreation.FileName,
                    FileType: objCreation.FileType,
                    Contents: objCreation.Contents ?? Framework.FetchJSON2(),
                    UppercaseHeadings: objCreation.UppercaseHeadings,
                    AutoSizeColumns: objCreation.AutoSizeColumns,
                    HeadingStyleID: objCreation.HeadingStyleID,
                    StandardStyleID: objCreation.StandardStyleID,
                    IndentedHeadingStyleID: objCreation.IndentedHeadingStyleID,
                    IndentedStyleID: objCreation.IndentedStyleID);

            HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
            result.Content = new ByteArrayContent(bytes);
            result.Content.Headers.ContentType =
                new MediaTypeHeaderValue("application/octet-stream");

            return result;
        }
    }
}

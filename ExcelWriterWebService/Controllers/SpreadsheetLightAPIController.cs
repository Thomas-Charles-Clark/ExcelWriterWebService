using ExcelWriterWebService.Models;
using ExcelWriterWebService.Utilities;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Web.Http;

namespace ExcelWriterWebService.Controllers
{
    public class SpreadsheetLightAPIController : ApiController
    {
        [Route("api/spreadsheet-light/CreateFile")]
        [HttpPost]
        public async Task<IHttpActionResult> CreateFile([FromBody]PassExcelCreation objCreation)
        {
            try
            {
                if (await Framework.WriteSpreadsheetLightFile(
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

        [Route("api/spreadsheet-light/GetFileStream")]
        [HttpPost]
        public async Task<HttpResponseMessage> GetFileStream([FromBody]PassExcelCreation objCreation)
        {
            var stream = await Framework.WriteSpreadsheetLightStream(
                    FileName: objCreation.FileName,
                    FileType: objCreation.FileType,
                    Contents: objCreation.Contents ?? Framework.FetchJSON2(),
                    UppercaseHeadings: objCreation.UppercaseHeadings,
                    AutoSizeColumns: objCreation.AutoSizeColumns,
                    HeadingStyleID: objCreation.HeadingStyleID,
                    StandardStyleID: objCreation.StandardStyleID,
                    IndentedHeadingStyleID: objCreation.IndentedHeadingStyleID,
                    IndentedStyleID: objCreation.IndentedStyleID);

            var bytes = stream.GetBuffer();

            HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
            result.Content = new ByteArrayContent(bytes);
            result.Content.Headers.ContentType =
                new MediaTypeHeaderValue("application/octet-stream");

            return result;
        }

		[Route("api/spreadsheet-light/CreateFromJsonFile")]
		[HttpPost]
		public async Task<IHttpActionResult> CreateFromJsonFile([FromBody] PassExcelCreationFromJsonFile objCreation)
		{
			try
			{
				if (await Framework.WriteSpreadsheetLightFile(
							FileName: objCreation.FileName,
							FileType: objCreation.FileType,
							Destination: objCreation.Destination,
							Contents: System.IO.File.ReadAllText(objCreation.OriginFilePath),
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
	}
}

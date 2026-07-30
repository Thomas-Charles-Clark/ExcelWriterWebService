using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Security;
using System.Web.Http;



using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using Newtonsoft.Json;
using System.Xml;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Text;
using System.Runtime.CompilerServices;


using LargeXlsx;
using System.ComponentModel;
using DocumentFormat.OpenXml.Wordprocessing;
using ExcelWriterWebService.Models;
using System.Web;


// Thanks, Google.
public partial class ExcelStreamingAPIController : ApiController
{
	private static readonly Regex InvalidXMLCharsRegex = new Regex("[^\x09\x0A\x0D\x20-\x7E\x85\xA0-\uD7FF\uE000-\uFDCF\uFDE0-\uFFFD]", RegexOptions.Compiled);

	[HttpPost]
	[Route("api/excel-streaming/StreamJsonToXlsx")]
	public async Task<IHttpActionResult> StreamJsonToXlsx()
	{
#if DEBUG
		System.Diagnostics.Debug.WriteLine("api/excel-streaming/StreamJsonToXlsx hit");
#endif

		// 1. Extract custom destination file path header safely
		if (!Request.Headers.TryGetValues("DestinationFullFilePath", out var headerValues))
		{
			return BadRequest("Missing 'DestinationFullFilePath' header.");
		}
		var destinationFullFilePath = headerValues.FirstOrDefault();

		if (string.IsNullOrWhiteSpace(destinationFullFilePath))
		{
			return BadRequest("Invalid 'DestinationFullFilePath' value.");
		}

		// 2. Ensure destination directory exists
		var directory = Path.GetDirectoryName(destinationFullFilePath);
		if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
		{
			Directory.CreateDirectory(directory);
		}

		try
		{
			// 3. Open the incoming raw HTTP body stream directly (No more multipart temp files!)
			using (Stream rawHttpStream = await Request.Content.ReadAsStreamAsync())
			// 4. Open the target Excel file stream
			using (FileStream diskTargetStream = new FileStream(destinationFullFilePath, FileMode.Create, FileAccess.ReadWrite, FileShare.None, 32768, useAsync: true))  // 128kb buffer; IO on the Files server is costly
			using (var writer = new XlsxWriter(diskTargetStream))
			{	
				writer.BeginWorksheet("Data Export");

				// Force the StreamReader to explicitly pair bytes as 2-byte UTF-16LE units
				var utf16LeNoBom = new UnicodeEncoding(bigEndian: false, byteOrderMark: false, throwOnInvalidBytes: false);

				// 2. FORCE a managed byte buffer layer (e.g., 64KB) to absorb network boundary clipping safely
				using (Stream incomingJsonStreamBuffer = new BufferedStream(rawHttpStream, 32768))
				// 3. Initialize the readers on top of the safely padded stream
				using (var streamReader = new StreamReader(incomingJsonStreamBuffer, utf16LeNoBom))
				using (var jsonReader = new JsonTextReader(streamReader))
				{
					bool isFirstObject = true;
					var currentObjectValues = new List<object>();

					while (await jsonReader.ReadAsync())
					{
						if (jsonReader.TokenType == JsonToken.StartObject)
						{
							currentObjectValues.Clear();
						}

						// Pass 1: Write out structural Excel headers (Only during the very first record)
						if (isFirstObject && jsonReader.TokenType == JsonToken.PropertyName)
						{
							string propName = jsonReader.Value?.ToString() ?? "";

							if (currentObjectValues.Count == 0)
							{
								writer.BeginRow(); // Initialize Header Row
							}

							writer.Write(propName);
						}

						// Buffer primitive values for the current processing record row
						if (jsonReader.TokenType == JsonToken.String ||
							jsonReader.TokenType == JsonToken.Integer ||
							jsonReader.TokenType == JsonToken.Float ||
							jsonReader.TokenType == JsonToken.Boolean ||
							jsonReader.TokenType == JsonToken.Date||
							jsonReader.TokenType == JsonToken.Bytes ||
							jsonReader.TokenType == JsonToken.Null)
						{
							currentObjectValues.Add(jsonReader.Value);
						}

						if (jsonReader.TokenType == JsonToken.EndObject)
						{
							writer.BeginRow(); // Open a clean row block for data values

							foreach (var val in currentObjectValues)
							{
								// 1. Handle Nulls cleanly first
								if (val == null)
								{
									writer.Write("");
								}
								// 2. Exact matches for exact overloads (No casting overhead or precision loss)
								else if (val is int intVal)
								{
									writer.Write(intVal);
								}
								else if (val is decimal decimalVal)
								{
									writer.Write(decimalVal); // Preserves exact currency/financial values
								}
								else if (val is double doubleVal)
								{
									writer.Write(doubleVal);
								}
								else if (val is bool booleanVal)
								{
									writer.Write(booleanVal);
								}
								else if (val is DateTime dateTimeVal)
								{
									writer.Write(dateTimeVal); // Leverages native Util.DateToDouble conversion!
								}
								// 3. Smart Widening Casts for unrepresented C# integer types
								else if (val is long || val is short || val is byte)
								{
									// Json.NET will sometimes map small ints to int, large to long.
									// If it hits a true long, Excel can safely store it as a decimal or double.
									writer.Write(Convert.ToDecimal(val));
								}
								else if (val is float floatVal)
								{
									writer.Write(Convert.ToDouble(floatVal));
								}
								// 4. Fallback for strings or any other unhandled type
								else
								{
									writer.Write(val.ToString());
								}
							}

							isFirstObject = false;
						}
					}
				}
			}
		}
		catch (Exception exception)
		{
			// Handle network drops or parsing issues gracefully

			// Log your exception here
			var db = new DatabaseContext();

			db.SaveToErrorLog(
				memberId: 0,
				exceptionType: exception.GetType().ToString(),
				exceptionMessage: exception.Message,
				exceptionSource: exception.Source,
				stackTrace: exception.StackTrace,
				innerExceptionType: exception.InnerException?.GetType().ToString(),
				innerExceptionMessage: exception.InnerException?.Message,
				innerExceptionSource: exception.InnerException?.Source,
				innerExceptionStackTrace: exception.InnerException?.StackTrace,
				HttpContext.Current.Request.Url.AbsoluteUri,
				HttpContext.Current.Request.UserAgent,
				HttpContext.Current.Request.UserHostAddress,
				DateTime.Now
			);

			return InternalServerError(exception);
		}

		return Ok(new
		{
			Message = "File successfully generated and saved to files server.",
			FileName = Path.GetFileName(destinationFullFilePath),
			SavedPath = destinationFullFilePath
		});
	}
}
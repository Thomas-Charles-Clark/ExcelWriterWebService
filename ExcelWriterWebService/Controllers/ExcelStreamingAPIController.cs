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


// Thanks, Google.
public partial class ExcelStreamingAPIController : ApiController
{
	private static readonly Regex InvalidXMLCharsRegex = new Regex("[^\x09\x0A\x0D\x20-\x7E\x85\xA0-\uD7FF\uE000-\uFDCF\uFDE0-\uFFFD]", RegexOptions.Compiled);

	[HttpPost]
	[Route("api/excel-streaming/StreamJsonToXlsx")]
	public async Task<IHttpActionResult> StreamJsonToXlsx()
	{
		// 1. Enforce multipart request type
		if (!Request.Content.IsMimeMultipartContent())
		{
			return StatusCode(HttpStatusCode.UnsupportedMediaType);
		}

		// Extract custom header safely
		IEnumerable<string> headerValues;
		if (!Request.Headers.TryGetValues("DestinationFullFilePath", out headerValues))
		{
			return BadRequest("Missing 'DestinationFullFilePath' header.");
		}
		var destinationFullFilePath = headerValues.FirstOrDefault();

		if (!System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(destinationFullFilePath)))
		{
			System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(destinationFullFilePath));
		}

		// 2. Read the multipart contents as a true stream without full buffering
		var streamProvider = await Request.Content.ReadAsMultipartAsync();

		// Find the file content block without downloading it to memory first
		var jsonContent = streamProvider.Contents.FirstOrDefault(c => !string.IsNullOrEmpty(c.Headers.ContentDisposition.FileName));

		if (jsonContent == null)
		{
			return BadRequest("Missing streamed JSON object.");
		}

		// 3. Open the incoming HTTP network stream and the target disk stream
		// Using true async I/O handles high-throughput operations smoothly
		using (Stream incomingJsonStream = await jsonContent.ReadAsStreamAsync())
		using (FileStream diskTargetStream = new FileStream(destinationFullFilePath, FileMode.Create, FileAccess.ReadWrite, FileShare.None, 4096, useAsync: true))
		{
			// Execute the processing. The JSON tokens flow straight from the network card 
			// through OpenXmlWriter and directly onto your hard drive. 
			GenerateExcelToDiskPath(incomingJsonStream, diskTargetStream);
		}

		// 4. Send success metadata packet downstream
		return Ok(new
		{
			Message = "File successfully generated and saved to files server.",
			FileName = Path.GetFileName(destinationFullFilePath),
			SavedPath = destinationFullFilePath
		});
	}

	private void GenerateExcelToDiskPath(Stream jsonInputStream, Stream fileDiskStream)
	{
		// SpreadsheetDocument processes elements directly into the FileStream channel
		using (var spreadsheetDocument = SpreadsheetDocument.Create(fileDiskStream, SpreadsheetDocumentType.Workbook))
		{
			var workbookPart = spreadsheetDocument.AddWorkbookPart();
			workbookPart.Workbook = new Workbook();
			var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

			using (var xmlWriter = OpenXmlWriter.Create(worksheetPart, Encoding.UTF8))
			{
				xmlWriter.WriteStartElement(new Worksheet());
				xmlWriter.WriteStartElement(new SheetData());

				using (var streamReader = new StreamReader(jsonInputStream))
				using (var jsonReader = new JsonTextReader(streamReader))
				{
					uint rowIndex = 1;
					int colIndex = 0;
					bool insideObject = false;

					while (jsonReader.Read())
					{
						if (jsonReader.TokenType == JsonToken.StartObject)
						{
							insideObject = true;
							colIndex = 0;

							var rowAttributes = new[] { new OpenXmlAttribute("r", null, rowIndex.ToString()) };
							xmlWriter.WriteStartElement(new Row(), rowAttributes);
						}
						else if (insideObject && IsValueToken(jsonReader.TokenType))
						{
							colIndex++;
							WriteInferredCell(xmlWriter, jsonReader.Value, jsonReader.TokenType, (uint)colIndex, rowIndex);
						}
						else if (jsonReader.TokenType == JsonToken.EndObject)
						{
							xmlWriter.WriteEndElement(); // Close Row
							rowIndex++;
							insideObject = false;
						}
					}
				}

				xmlWriter.WriteEndElement(); // Close SheetData
				xmlWriter.WriteEndElement(); // Close Worksheet
			}

			var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());
			var sheet = new Sheet()
			{
				Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
				SheetId = 1,
				Name = "Data Export"
			};
			sheets.Append(sheet);
			workbookPart.Workbook.Save();
		}
	}

	private bool IsValueToken(JsonToken tokenType)
	{
		return tokenType == JsonToken.String ||
			   tokenType == JsonToken.Integer ||
			   tokenType == JsonToken.Float ||
			   tokenType == JsonToken.Boolean ||
			   tokenType == JsonToken.Null;
	}

	private void WriteInferredCell(OpenXmlWriter writer, object value, JsonToken tokenType, uint colIndex, uint rowIndex)
	{
		string cellReference = GetCellReference((int)colIndex, rowIndex);

		if (value == null || tokenType == JsonToken.Null)
		{
			var cellAttributes = new[] { new OpenXmlAttribute("r", null, cellReference) };
			writer.WriteStartElement(new Cell(), cellAttributes);
			writer.WriteEndElement();
			return;
		}

		string typeAttr = "str";
		string stringValue = value.ToString();

		switch (tokenType)
		{
			case JsonToken.Integer:
			case JsonToken.Float:
				typeAttr = "n";
				stringValue = Convert.ToString(value);
				break;

			case JsonToken.Boolean:
				typeAttr = "b";
				stringValue = (bool)value ? "1" : "0";
				break;

			case JsonToken.String:
			default:
				typeAttr = "str";
				break;
		}

		var attributes = new[] {
			new OpenXmlAttribute("r", null, cellReference),
			new OpenXmlAttribute("t", null, typeAttr)
		};

		writer.WriteStartElement(new Cell(), attributes);

		// Keeps only valid XML 1.0 characters
		string cleanValue = Regex.Replace(stringValue, @"[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F-\u0084\u0086-\u009F\uFFFE\uFFFF]", "");

		writer.WriteElement(new CellValue(cleanValue));
		writer.WriteEndElement();
	}

	static string RemoveInvalidXmlChars(string text)
	{
		var validXmlChars = text.Where(ch => XmlConvert.IsXmlChar(ch)).ToArray();
		return new string(validXmlChars);
	}

	private string GetCellReference(int columnIndex, uint rowIndex)
	{
		int dividend = columnIndex;
		string columnName = String.Empty;
		int modifier;

		while (dividend > 0)
		{
			modifier = (dividend - 1) % 26;
			columnName = Convert.ToChar(65 + modifier).ToString() + columnName;
			dividend = (dividend - modifier) / 26;
		}

		return $"{columnName}{rowIndex}";
	}
}
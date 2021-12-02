using Dapper;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using SimpleImpersonation;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;

namespace ExcelWriterWebService.Utilities
{
    public static class StringExtension
    {
        public static string TrimReduceTruncate(this string str, int length)
        {
            if (str == null)
            {
                return "";
            }

            var result = ConvertWhitespacesToSingleSpaces(str).Trim();

            if (result.Length > length)
            {
                return result.Substring(0, length);
            }
            else
            {
                return result;
            }
        }

        public static string ConvertWhitespacesToSingleSpaces(this string value)
        {
            return Regex.Replace(value, @"\s+", " ");
        }

        public static string SplitCamelCase(this string input)
        {
            return Regex.Replace(input, "([A-Z])", " $1", RegexOptions.Compiled).Trim();
        }
    }

    public static partial class Framework
    {
        public static byte[] StringToByteArray(String hex)
        {
            int NumberChars = hex.Length;
            byte[] bytes = new byte[NumberChars / 2];
            for (int i = 0; i < NumberChars; i += 2)
                bytes[i / 2] = Convert.ToByte(hex.Substring(i, 2), 16);
            return bytes;
        }

        public static string FetchJSON2()
        {
            using (var conn = new SqlConnection("Data Source=vm-sql1;Initial Catalog=Flock;User=dg;Password=mattm00"))
            {
                try
                {
                    // Cushions, Fabric, Wallpaper, Pattern Books.
                    return conn.Query<string>($@"SELECT (SELECT 
TOP 100
TA.*
FROM dbo.TradeAccount AS TA
ORDER BY TA.TradeAccountID DESC
	FOR JSON PATH, INCLUDE_NULL_VALUES)").Single();
                }
                finally
                {
                    conn.Close();
                }
            }
        }


        public static bool WriteFile(
            string FileName,
            string FileType,
            string Contents,
            string Destination,
            bool UppercaseHeadings,
            bool AutoSizeColumns,
            byte? HeadingStyleID,
            byte? StandardStyleID,
            byte? IndentedHeadingStyleID,
            byte? IndentedStyleID)
        {
            var credentials = new UserCredentials("DESIGNERSGUILD\\SQLAgentFileAccess", "34949C02-4677-4711-999C-8E038E26B10F");

            var data = JsonConvert.DeserializeObject<IEnumerable<ExpandoObject>>(Contents, new ExpandoObjectConverter());

            var fullPath = $"{Destination}\\{FileName}.{FileType}";

            switch (FileType)
            {
                case "csv":
                    {
                        Impersonation.RunAsUser(credentials, LogonType.Interactive, () =>
                        {
                            using (var fs = System.IO.File.CreateText(fullPath))
                            {
                                var counter = 0;

                                foreach (var item in data)
                                {
                                    IDictionary<string, object> propertyValues = item;

                                    if (counter == 0)
                                    {
                                        if (UppercaseHeadings)
                                            fs.WriteLine(String.Join(",", propertyValues.Keys.Select(p => p?.Replace(",", " ").Replace("\r\n", " ").Replace("\n", " ").ToUpperInvariant())));
                                        else
                                            fs.WriteLine(String.Join(",", propertyValues.Keys.Select(p => p?.Replace(",", " ").Replace("\r\n", " ").Replace("\n", " "))));
                                    }

                                    counter++;

                                    fs.WriteLine(String.Join(",", propertyValues.Values.Select(p => p?.ToString().Replace(",", " ").Replace("\r\n", " ").Replace("\n", " "))));
                                }
                            }
                        });
                    }
                    break;

                case "zip":
                case "xlsx":
                    {
                        Impersonation.RunAsUser(credentials, LogonType.Interactive, () =>
                        {
                            if (System.IO.File.Exists(fullPath))
                            {
                                System.IO.File.Delete(fullPath);
                            }

                            using (var fs = new System.IO.FileStream(fullPath, System.IO.FileMode.OpenOrCreate))
                            {
                                using (var ts = new System.IO.MemoryStream(Framework.StringToByteArray(Framework.TemplateHex)))
                                {
                                    using (var fastExcel = new FastExcel.FastExcel(ts, fs))
                                    {
                                        fastExcel.Write(
                                            rows: data,
                                            sheetName: "Sheet1",
                                            uppercaseHeadings: UppercaseHeadings,
                                            autoWidthColumns: AutoSizeColumns,
                                            headingStyleId: HeadingStyleID,
                                            standardStyleId: StandardStyleID,
                                            indentedHeadingStyleId: IndentedHeadingStyleID,
                                            indentedStyleId: IndentedStyleID);
                                    }
                                }
                            }
                        });
                    }
                    break;
            }

            return true;
        }

        public static byte[] WriteFileStream(
            string FileName,
            string FileType,
            string Contents,
            bool UppercaseHeadings,
            bool AutoSizeColumns,
            byte? HeadingStyleID,
            byte? StandardStyleID,
            byte? IndentedHeadingStyleID,
            byte? IndentedStyleID)
        {
            var credentials = new UserCredentials("DESIGNERSGUILD\\SQLAgentFileAccess", "34949C02-4677-4711-999C-8E038E26B10F");

            var data = JsonConvert.DeserializeObject<IEnumerable<ExpandoObject>>(Contents, new ExpandoObjectConverter());

            switch (FileType)
            {
                case "csv":
                    {
                        return Impersonation.RunAsUser(credentials, LogonType.Interactive, () =>
                        {
                            using (var ms = new MemoryStream())
                            {
                                using (var sw = new StreamWriter(ms))
                                {
                                    var counter = 0;

                                    foreach (var item in data)
                                    {
                                        IDictionary<string, object> propertyValues = item;

                                        if (counter == 0)
                                        {
                                            if (UppercaseHeadings)
                                            {
                                                sw.WriteLine(String.Join(",", propertyValues.Keys.Select(p => p?.Replace(",", " ").Replace("\r\n", " ").Replace("\n", " ").ToUpperInvariant())));
                                            }
                                            else
                                            {
                                                sw.WriteLine(String.Join(",", propertyValues.Keys.Select(p => p?.Replace(",", " ").Replace("\r\n", " ").Replace("\n", " "))));
                                            }
                                        }

                                        counter++;

                                        sw.WriteLine(String.Join(",", propertyValues.Values.Select(p => p?.ToString().Replace(",", " ").Replace("\r\n", " ").Replace("\n", " "))));
                                    }

                                    sw.Flush();

                                    return ms.ToArray();
                                }
                            }
                        });
                    }

                case "xlsx":
                    {
                        return Impersonation.RunAsUser(credentials, LogonType.Interactive, () =>
                        {
                            using (var ms = new MemoryStream())
                            {
                                using (var ts = new MemoryStream(Framework.StringToByteArray(Framework.TemplateHex)))
                                {
                                    using (var fastExcel = new FastExcel.FastExcel(ts, ms))
                                    {
                                        fastExcel.Write(
                                            rows: data,
                                            sheetName: "Sheet1",
                                            uppercaseHeadings: UppercaseHeadings,
                                            autoWidthColumns: AutoSizeColumns,
                                            headingStyleId: HeadingStyleID,
                                            standardStyleId: StandardStyleID,
                                            indentedHeadingStyleId: IndentedHeadingStyleID,
                                            indentedStyleId: IndentedStyleID);
                                    }
                                }

                                return ms.ToArray();
                            }
                        });
                    }

                default:
                    return new byte[] { };
            }
        }
    }
}
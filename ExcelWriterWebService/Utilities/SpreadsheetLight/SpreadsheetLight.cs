using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Web;

using Microsoft.Ajax.Utilities;

using Newtonsoft.Json;
using Newtonsoft.Json.Converters;

using SpreadsheetLight;

using ExcelWriterWebService.Models;
using FastExcel;
using System.Security.AccessControl;
using SpreadsheetLight.Drawing;
using System.Threading.Tasks;
using System.IO;

namespace ExcelWriterWebService.Utilities
{
    public static partial class Framework
    {
        public static async Task<bool> WriteSpreadsheetLightFile(string FileName, string FileType, string Destination, string Contents, bool UppercaseHeadings, bool AutoSizeColumns, byte? HeadingStyleID, byte? StandardStyleID, byte? IndentedHeadingStyleID, byte? IndentedStyleID)
        {
            var data = JsonConvert.DeserializeObject<IEnumerable<ExpandoObject>>(Contents, new ExpandoObjectConverter());

            var columnHeadings = new List<string>();

            var rows = PopulateRowsFromIEnumerable(data, ref columnHeadings, UppercaseHeadings, AutoSizeColumns, HeadingStyleID, StandardStyleID, IndentedHeadingStyleID, IndentedStyleID);

            using (var doc = new SLDocument())
            {
                var style = doc.CreateStyle();

                style.SetVerticalAlignment(DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center);

                if (HeadingStyleID != null)
                {
                    doc.ApplyNamedCellStyleToRow(1, (SLNamedCellStyleValues)(int)(HeadingStyleID.Value));
                }

                for (var i = 0; i < columnHeadings.Count; i++)
                {
                    var heading = columnHeadings.ElementAt(i);

                    if (UppercaseHeadings)
                    {
                        doc.SetCellValue(1, i + 1, heading.ToUpperInvariant());
                    }
                    else
                    {
                        doc.SetCellValue(1, i + 1, heading);
                    }

                    if (heading == "ImageID")
                    {
                        doc.SetColumnWidth(i + 1, 13);
                    }
                }

                // there's a row in here for the column headings...
                // but we've just added them so remove that first row
                rows = rows.Skip(1).Take(rows.Count());

                for(var i = 0; i < rows.Count(); i++)
                {
                    var row = rows.ElementAt(i);

                    doc.SetRowHeight(i + 2, 68);

                    for (var j = 0; j < row.Cells.Count(); j++)
                    {
                        var cell = row.Cells.ElementAt(j);

                        var cellType = "s";

                        var styleIndex = 0;

                        var Value = cell.Value;

                        var ColumnHeading = columnHeadings.ElementAt(j);

                        var columnHeadingLowered = ColumnHeading?.ToLowerInvariant();

                        if (Value != null)
                        {
                            object value = Value;

                            string stringedValue = $"{value}";

                            //if (ColumnHeading?.ToLowerInvariant().Contains("reference") == false)
                            //{
                            if (columnHeadingLowered.Contains("code") == true ||
                                columnHeadingLowered.Contains("reference") == true ||
                                columnHeadingLowered.Contains("orderid") == true ||
                                columnHeadingLowered.Contains("orderheadid") == true)
                            {
                                cellType = "s";
                            }
                            else if (columnHeadingLowered == "imageid")
                            {
                                cellType = "i";
                            }
                            else
                            {
                                if (Byte.TryParse(stringedValue, out _) ||
                                    Int16.TryParse(stringedValue, out _) ||
                                    Int32.TryParse(stringedValue, out _) ||
                                    Int64.TryParse(stringedValue, out _))
                                {
                                    cellType = "n";
                                    styleIndex = 41;
                                }

                                if (stringedValue.Contains(".") || ColumnHeading?.Contains("Currency") == true || ColumnHeading?.Contains("Sterling") == true)
                                {
                                    if (Decimal.TryParse(stringedValue, out _))
                                    {
                                        cellType = "n";
                                        styleIndex = 39;
                                    }
                                }

                                if (DateTime.TryParseExact(stringedValue, "yyyy-MM-dd", System.Globalization.DateTimeFormatInfo.InvariantInfo, System.Globalization.DateTimeStyles.None, out DateTime dt))
                                {
                                    //value = dt.ToString("dd/MM/yyyy");
                                    //cellType = "d";
                                    styleIndex = 40;
                                }
                            }
                        }

                        var x = j + 1;
                        var y = i + 2;

                        doc.SetCellStyle(y, x, style);

                        switch (cellType)
                        {
                            case "s":
                                doc.SetCellValue(y, x, Value.ToString());
                                break;
                            case "n":
                                doc.SetCellValueNumeric(y, x, Value.ToString());
                                break;
                            case "d":
                                doc.SetCellValue(y, x, (DateTime.Parse(Value.ToString())));
                                break;
                            case "i":
                                if (Value != null)
                                {
                                    doc.SetCellValue(y, x, $"=_xlfn.IMAGE(\"https://www.designersguild.com/image-nonwebp/90/{Value}\")");
                                }
                                break;
                            default:
                                break;
                        }
                    }
                }
                
                for (var i = 0; i < columnHeadings.Count; i++)
                {
                    if (AutoSizeColumns)
                    {
                        doc.AutoFitColumn(i + 1);
                    }

                    if (StandardStyleID != null)
                    {
                        doc.ApplyNamedCellStyleToColumn(i + 1, (SLNamedCellStyleValues)(int)(StandardStyleID.Value));
                    }
                }

                doc.FlattenAllSharedCellFormula();

                doc.SaveAs($"{Destination}\\{System.IO.Path.GetFileNameWithoutExtension(FileName)}.xlsx");
            }

            return true;
        }

        public static async Task<MemoryStream> WriteSpreadsheetLightStream(string FileName, string FileType, string Contents, bool UppercaseHeadings, bool AutoSizeColumns, byte? HeadingStyleID, byte? StandardStyleID, byte? IndentedHeadingStyleID, byte? IndentedStyleID)
        {
            var data = JsonConvert.DeserializeObject<IEnumerable<ExpandoObject>>(Contents, new ExpandoObjectConverter());

            var columnHeadings = new List<string>();

            var rows = PopulateRowsFromIEnumerable(data, ref columnHeadings, UppercaseHeadings, AutoSizeColumns, HeadingStyleID, StandardStyleID, IndentedHeadingStyleID, IndentedStyleID);

            var memoryStream = new MemoryStream();

            using (var doc = new SLDocument())
            {
                for (var i = 0; i < columnHeadings.Count; i++)
                {
                    var heading = columnHeadings.ElementAt(i);

                    if (UppercaseHeadings)
                    {
                        doc.SetCellValue(1, i + 1, heading.ToUpperInvariant());
                    }
                    else
                    {
                        doc.SetCellValue(1, i + 1, heading);
                    }

                    if (HeadingStyleID != null)
                    {
                        doc.ApplyNamedCellStyleToColumn(i + 1, (SLNamedCellStyleValues)(int)(HeadingStyleID.Value));
                    }
                }

                // there's a row in here for the column headings...
                // but we've just added them so remove that first row
                rows = rows.Skip(1).Take(rows.Count());

                for (var i = 0; i < rows.Count(); i++)
                {
                    var row = rows.ElementAt(i);

                    for (var j = 0; j < row.Cells.Count(); j++)
                    {
                        var cell = row.Cells.ElementAt(j);

                        var cellType = "s";

                        var styleIndex = 0;

                        var Value = cell.Value;

                        var ColumnHeading = columnHeadings.ElementAt(j);

                        if (Value != null)
                        {
                            object value = Value;

                            string stringedValue = $"{value}";

                            //if (ColumnHeading?.ToLowerInvariant().Contains("reference") == false)
                            //{
                            if (ColumnHeading?.ToLowerInvariant().Contains("code") == true ||
                                ColumnHeading?.ToLowerInvariant().Contains("reference") == true ||
                                ColumnHeading?.EndsWith("ID") == true)
                            {
                                cellType = "s";

                                if (ColumnHeading == "ImageID")
                                {
                                    cellType = "i";
                                }
                            }
                            else
                            {
                                if (Byte.TryParse(stringedValue, out _) ||
                                    Int16.TryParse(stringedValue, out _) ||
                                    Int32.TryParse(stringedValue, out _) ||
                                    Int64.TryParse(stringedValue, out _))
                                {
                                    cellType = "n";
                                    styleIndex = 41;
                                }

                                if (stringedValue.Contains(".") || ColumnHeading?.Contains("Currency") == true || ColumnHeading?.Contains("Sterling") == true)
                                {
                                    if (Decimal.TryParse(stringedValue, out _))
                                    {
                                        cellType = "n";
                                        styleIndex = 39;
                                    }
                                }

                                if (DateTime.TryParseExact(stringedValue, "yyyy-MM-dd", System.Globalization.DateTimeFormatInfo.InvariantInfo, System.Globalization.DateTimeStyles.None, out DateTime dt))
                                {
                                    //value = dt.ToString("dd/MM/yyyy");
                                    //cellType = "d";
                                    styleIndex = 40;
                                }
                            }
                        }

                        var x = j + 1;
                        var y = i + 2;

                        doc.SetRowHeight(y, 68);

                        var style = doc.CreateStyle();

                        style.SetVerticalAlignment(DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center);

                        doc.SetCellStyle(y, x, style);

                        doc.AutoFitColumn(x);

                        if (cell.ColumnHeading == "ImageID")
                        {
                            doc.SetColumnWidth(x, 13);
                        }

                        switch (cellType)
                        {
                            case "s":
                                doc.SetCellValue(y, x, Value.ToString());
                                break;
                            case "n":
                                doc.SetCellValueNumeric(y, x, Value.ToString());
                                break;
                            case "d":
                                doc.SetCellValue(y, x, (DateTime.Parse(Value.ToString())));
                                break;
                            case "i":
                                doc.SetCellValue(y, x, $"=_xlfn.IMAGE(\"https://www.designersguild.com/image-nonwebp/90/{Value}\")");
                                break;
                            default:
                                break;
                        }
                    }
                }

                doc.FlattenAllSharedCellFormula();

                doc.SaveAs(memoryStream);
            }

            return memoryStream;
        }

        // Yoink
        private static IEnumerable<Row> PopulateRowsFromIEnumerable(IEnumerable<ExpandoObject> rows, ref List<string> columnHeadings, bool UppercaseHeadings, bool AutoSizeColumns, byte? HeadingStyleID, byte? StandardStyleID, byte? IndentedHeadingStyleID, byte? IndentedStyleID)
        {
            int overallRowCounter = 1;
            int indentationCounter = 0;

            var result = WriteExpandoObjectList(rows, ref columnHeadings, ref overallRowCounter, ref indentationCounter, UppercaseHeadings, AutoSizeColumns, HeadingStyleID, StandardStyleID, IndentedHeadingStyleID, IndentedStyleID, null);

            //if (autoWidthColumns)
            //{
            //    for (var i = 0; i < result.Select(p => p.Cells.Count()).Max(); i++)
            //    {
            //        Columns.Add(new Column()
            //        {
            //            Index = i + 1,
            //            LongestContentWidth = null
            //        });
            //    }

            //    Columns.ForEach(c =>
            //    {
            //        c.LongestContentWidth = result.SelectMany(p => p.Cells.Where(q => q.ColumnNumber - 1 == c.Index).Select(q => q.Length)).Max() + 3;
            //    });
            //}

            return result;
        }

        private static List<Row> WriteExpandoObjectList(IEnumerable<ExpandoObject> data, ref List<string> columnHeadings, ref int overallRowCounter, ref int indentationCounter, bool UppercaseHeadings, bool AutoSizeColumns, byte? HeadingStyleID, byte? StandardStyleID, byte? IndentedHeadingStyleID, byte? IndentedStyleID, string subTableLabel = null)
        {
            List<Row> rows = new List<Row>();

            int thisTableRowCounter = 1;

            foreach (var item in data)
            {
                IDictionary<string, object> propertyValues = item;

                if (thisTableRowCounter == 1)
                {
                    var cells = new List<Cell>();

                    thisTableRowCounter++;

                    var headerCellCounter = 1;

                    for (var i = 0; i < indentationCounter; i++)
                    {
                        cells.Add(new Cell(headerCellCounter++, null));
                    }

                    var headerStyleIndex = indentationCounter == 0 ? HeadingStyleID : IndentedHeadingStyleID; // 15 : 5

                    var keys = propertyValues.Keys.ToList();
                    
                    for (var i = 0; i < keys.Count; i++)
                    {
                        var k = keys[i];

                        if (propertyValues[k] as List<object> != null)
                        {
                            // We have a thing we want to indent and display as related data but of a different set
                            // The first step here is, yes, to NOT write a column in teh main set for these types
                        }
                        else
                        {
                            var heading = k?.Replace(",", " ").Replace("\r\n", " ").Replace("\n", " ");//?.SplitCamelCase();

                            if (UppercaseHeadings)
                                cells.Add(new Cell(headerCellCounter++, heading.ToUpperInvariant(), headerStyleIndex));
                            else
                                cells.Add(new Cell(headerCellCounter++, heading, headerStyleIndex));

                            columnHeadings.Add(heading);
                        }
                    }

                    var headerRow = new Row(overallRowCounter++, cells);

                    rows.Add(headerRow);
                }

                var dataCells = new List<Cell>();

                var dataCellCounter = 1;

                // what about having MULTIPLE properties which could be arrays, think about that
                List<Tuple<string, List<object>>> AdditionalDataSets = new List<Tuple<string, List<object>>>();

                for (var i = 0; i < indentationCounter; i++)
                {
                    var h = columnHeadings.Count > i ? columnHeadings[i] : null;

                    if (i == indentationCounter - 1 && !String.IsNullOrEmpty(subTableLabel))
                    {
                        dataCells.Add(new Cell(dataCellCounter++, UppercaseHeadings ? subTableLabel.ToUpperInvariant() : subTableLabel, IndentedHeadingStyleID, h)); // bold it
                    }
                    else
                    {
                        dataCells.Add(new Cell(dataCellCounter++, null, null, h));
                    }
                }

                var styleIndex = indentationCounter == 0 ? StandardStyleID : IndentedStyleID; // 13 : 26

                var moreKeys = propertyValues.Keys.ToList();

                for (var i = 0; i < moreKeys.Count; i++)
                {
                    var k = moreKeys[i];

                    var v = propertyValues[k];

                    if (v is List<object> list)
                    {
                        AdditionalDataSets.Add(new Tuple<string, List<object>>(k, list));
                    }
                    else
                    {
                        var h = columnHeadings.Count > (dataCellCounter - 1) ? columnHeadings[dataCellCounter - 1] : null;

                        dataCells.Add(new Cell(dataCellCounter++, v?.ToString().Replace(",", " ").Replace("\r\n", " ").Replace("\n", " ") ?? "", styleIndex, h));
                    }
                }

                var dataRow = new Row(overallRowCounter++, dataCells);

                rows.Add(dataRow);

                if (AdditionalDataSets.Where(p => p.Item2.Count > 0).Any())
                {
                    indentationCounter += 2;

                    for (var i = 0; i < AdditionalDataSets.Count; i++)
                    {
                        var c = AdditionalDataSets[i];

                        if (c.Item2.Count > 0)
                        {
                            rows.Add(new Row(overallRowCounter++, new List<Cell>()));

                            var eo = JsonConvert.DeserializeObject<IEnumerable<ExpandoObject>>(JsonConvert.SerializeObject(c.Item2), new ExpandoObjectConverter());

                            rows.AddRange(WriteExpandoObjectList(eo, ref columnHeadings, ref overallRowCounter, ref indentationCounter, UppercaseHeadings, AutoSizeColumns, HeadingStyleID, StandardStyleID, IndentedHeadingStyleID, IndentedStyleID, c.Item1));
                        }
                    }

                    indentationCounter = 0;

                    overallRowCounter += 2;
                }
            }

            return rows;
        }

    }
}
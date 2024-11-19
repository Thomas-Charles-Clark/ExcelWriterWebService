using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
 
namespace ExcelWriterWebService.Models
{
    public class PassExcelCreation
    {
        public string FileName { get; set; }
        public string FileType { get; set; }
        public string Destination { get; set; }
        public string Contents { get; set; }
        public bool UppercaseHeadings { get; set; }
        public bool AutoSizeColumns { get; set; }
        public byte? HeadingStyleID { get; set; }
        public byte? StandardStyleID { get; set; }
        public byte? IndentedHeadingStyleID { get; set; }
        public byte? IndentedStyleID { get; set; }

        public PassExcelCreation()
        {

        }
    }
}
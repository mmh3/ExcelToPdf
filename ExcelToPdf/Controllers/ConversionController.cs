using System.Collections.Generic;
using System.Web.Http;

namespace ExcelToPdf.Controllers
{
    public class ConversionController : ApiController
    {
        public class ConversionArgs
        {
            public string sourceFilename { get; set; }
        }

        [HttpPost]
        [ActionName("Convert")]
        public IHttpActionResult Convert(string id)
        {
            // There are issues with sending in special characters in the path (:, /, .).
            // Convert -- => :/ and - => /. Also, assume any file passed in is a ".xlsx" file type.
            var filename = id.Replace("--", ":/");
            filename = filename.Replace("-", "/");
            filename += ".xlsx";

            var excelInteropExcelToPdfConverter = new ExcelInteropExcelToPdfConverter();
            excelInteropExcelToPdfConverter.ConvertToPdf(new List<string> { filename });

            return Ok();
        }
    }
}

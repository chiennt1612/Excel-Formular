using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using Utils;

namespace Excel_Formular.Pages
{
    public class IndexModel : PageModel
    {
        private readonly ILogger<IndexModel> _logger;

        public IndexModel(ILogger<IndexModel> logger)
        {
            _logger = logger;
        }

        public void OnGet()
        {
            ExcelPackage package = new ExcelPackage(new System.IO.FileInfo(System.IO.Directory.GetCurrentDirectory () + @"\wwwroot\Template\temp.xlsx"));
            dynamic DLanguage = JObject.Parse(System.IO.File.ReadAllText(System.IO.Directory.GetCurrentDirectory() + @"\DLanguge.json"));
            string DataParam = System.IO.File.ReadAllText(System.IO.Directory.GetCurrentDirectory() + @"\Param.json");
            string DataAll = System.IO.File.ReadAllText(System.IO.Directory.GetCurrentDirectory() + @"\Data.json");
            string DataPivot = System.IO.File.ReadAllText(System.IO.Directory.GetCurrentDirectory() + @"\Pivot.json");
            ReadExcelForm ac = new ReadExcelForm(package, DLanguage, DataParam.ToUpper(), DataAll, DataPivot);
            package.Compression = CompressionLevel.BestSpeed;
            package.SaveAs(new System.IO.FileInfo(System.IO.Directory.GetCurrentDirectory() + @"\wwwroot\Template\result.xlsx"));
        }
    }
}

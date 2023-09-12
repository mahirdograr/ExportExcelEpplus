using ExportExcel.Models;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Diagnostics;

namespace ExportExcel.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private AppDbContext _appDbContext;

        public HomeController(ILogger<HomeController> logger, AppDbContext appDbContext)
        {
            _logger = logger;
            _appDbContext = appDbContext;
        }

        public IActionResult Index()
        {
            var employees = _appDbContext.Comments.ToList();

            return View(employees);
        }


        [HttpPost]
        public IActionResult GenerateExcel()
        {
            var comments = _appDbContext.Comments.ToList();

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Comments");

                // Başlık satırının eklenmesi ve şekillenmesi
                worksheet.Cells["A1:B1"].Style.Font.Bold = true; // Kalın yazı
                worksheet.Cells["A1:B1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid; // Dolgu rengi kullanılması
                worksheet.Cells["A1:B1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Blue); // Mavi arka plan rengi kullanılması

                worksheet.Cells["A1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                worksheet.Cells["B1"].Style.Font.Color.SetColor(System.Drawing.Color.White);


                worksheet.Cells["A1"].Value = "ID";
                worksheet.Cells["B1"].Value = "CommentText";

                int row = 2; // Verilerin başlayacağı satır

                int toplamID = comments.Sum(comment => comment.CommentID); // ID'lerin toplamını hesapla

                foreach (var comment in comments)
                {
                    worksheet.Cells["A" + row].Value = comment.CommentID;
                    worksheet.Cells["B" + row].Value = comment.CommentContent;

                    row++;
                }

                // Toplam ID'yi yeni bir hücrede yazar
                worksheet.Cells["C1"].Value = "Toplam ID: " + toplamID;

                // Excel dosyası oluşuyır
                var stream = new MemoryStream(package.GetAsByteArray());
                stream.Seek(0, SeekOrigin.Begin);

                string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                string fileName = "Comments.xlsx";

                return File(stream, contentType, fileName);
            }
        }




        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
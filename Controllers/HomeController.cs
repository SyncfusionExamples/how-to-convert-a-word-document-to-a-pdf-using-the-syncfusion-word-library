using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using WordToPDF.Models;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;

namespace WordToPDF.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }


        public IActionResult ConvertWordToPDF()
        {
            using(FileStream docStream = new FileStream(Path.GetFullPath("Data/Template.docx"), FileMode.Open, FileAccess.Read))
            {
                using (WordDocument wordDocument = new WordDocument(docStream, FormatType.Docx))
                {
                    using(DocIORenderer renderer = new DocIORenderer())
                    {
                        renderer.Settings.AutoTag = true;
                        renderer.Settings.PdfConformanceLevel = PdfConformanceLevel.Pdf_A3A;
                        renderer.Settings.EmbedCompleteFonts = true;

                        PdfDocument pdfDocument = renderer.ConvertToPDF(wordDocument);
                        MemoryStream memoryStream = new MemoryStream();
                        pdfDocument.Save(memoryStream);
                        memoryStream.Position= 0;
                        return File(memoryStream, "application/pdf", "Sample.pdf");
                    }
                }
            }
        }

        public IActionResult FillablePDF()
        {
            using (FileStream docStream = new FileStream(Path.GetFullPath("Data/FormFields.docx"), FileMode.Open, FileAccess.Read))
            {
                using (WordDocument wordDocument = new WordDocument(docStream, FormatType.Docx))
                {
                    using (DocIORenderer renderer = new DocIORenderer())
                    {
                        renderer.Settings.PreserveFormFields = true;
                        PdfDocument pdfDocument = renderer.ConvertToPDF(wordDocument);
                        MemoryStream memoryStream = new MemoryStream();
                        pdfDocument.Save(memoryStream);
                        memoryStream.Position = 0;
                        return File(memoryStream, "application/pdf", "Sample.pdf");
                    }
                }
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
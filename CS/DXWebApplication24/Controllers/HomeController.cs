using System.Web;
using System.Web.Mvc;
using DevExpress.Web.Mvc;
using System.IO;
using DevExpress.Spreadsheet;
using DevExpress.XtraPrinting;

namespace DXWebApplication24.Controllers {
    public class HomeController : Controller {

        public ActionResult Index() {
            return View();
        }

        public ActionResult SpreadsheetPartial() {
            return PartialView("_SpreadsheetPartial");
        }

        public FileStreamResult SpreadsheetPartialDownload() {
            return SpreadsheetExtension.DownloadFile("Spreadsheet");
        }

        public FileStreamResult Export() {
            Stream pdfStream = GetPdfStream();
            HttpContext.Response.AddHeader("content-disposition", "attachment; filename=Document.pdf");
            return new FileStreamResult(pdfStream, "application/pdf");
        }

        Stream GetPdfStream() {
            MemoryStream ms = new MemoryStream();
            IWorkbook workBook = SpreadsheetExtension.GetCurrentDocument("Spreadsheet");

            workBook.SaveDocument(ms, DocumentFormat.Xlsm);
            ms.Position = 0;
            Workbook docServer = new Workbook();
            docServer.LoadDocument(ms, DocumentFormat.Xlsm);
            ms.Position = 0;
            PrintableComponentLink link = new PrintableComponentLink(new PrintingSystem());
            link.Component = docServer;
            link.ExportToPdf(ms);
            ms.Position = 0;

            return ms;
        }
    }
}
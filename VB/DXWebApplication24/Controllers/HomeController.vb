Imports Microsoft.VisualBasic
Imports System.Web
Imports System.Web.Mvc
Imports DevExpress.Web.Mvc
Imports System.IO
Imports DevExpress.Spreadsheet
Imports DevExpress.XtraPrinting

Namespace DXWebApplication24.Controllers
	Public Class HomeController
		Inherits Controller

		Public Function Index() As ActionResult
			Return View()
		End Function

		Public Function SpreadsheetPartial() As ActionResult
			Return PartialView("_SpreadsheetPartial")
		End Function

		Public Function SpreadsheetPartialDownload() As FileStreamResult
			Return SpreadsheetExtension.DownloadFile("Spreadsheet")
		End Function

		Public Function Export() As FileStreamResult
			Dim pdfStream As Stream = GetPdfStream()
			HttpContext.Response.AddHeader("content-disposition", "attachment; filename=Document.pdf")
			Return New FileStreamResult(pdfStream, "application/pdf")
		End Function

		Private Function GetPdfStream() As Stream
			Dim ms As New MemoryStream()
			Dim workBook As IWorkbook = SpreadsheetExtension.GetCurrentDocument("Spreadsheet")

			workBook.SaveDocument(ms, DocumentFormat.Xlsm)
			ms.Position = 0
			Dim docServer As New Workbook()
			docServer.LoadDocument(ms, DocumentFormat.Xlsm)
			ms.Position = 0
			Dim link As New PrintableComponentLink(New PrintingSystem())
			link.Component = docServer
			link.ExportToPdf(ms)
			ms.Position = 0

			Return ms
		End Function
	End Class
End Namespace
@Html.DevExpress().Spreadsheet(
    Sub(settings)
            settings.Name = "Spreadsheet"
            settings.CallbackRouteValues = New With {Key .Controller = "Home", Key .Action = "SpreadsheetPartial"}
            settings.DownloadRouteValues = New With {Key .Controller = "Home", Key .Action = "SpreadsheetPartialDownload"}
            settings.RibbonMode = SpreadsheetRibbonMode.Ribbon
    End Sub).Open(Server.MapPath("Documents\Sample.xlsx")).GetHtml()
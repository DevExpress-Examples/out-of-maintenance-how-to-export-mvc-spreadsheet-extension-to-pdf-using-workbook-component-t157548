@Code
    ViewBag.Title = "Index"
End Code

<h2>Index</h2>

@Using (Html.BeginForm("Export", "Home"))
    @Html.Action("SpreadsheetPartial")
    @<input type="submit" value="Export To PDF" />
End Using
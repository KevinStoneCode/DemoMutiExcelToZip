
Imports System.IO
Imports Ionic.Zip
Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel

Partial Class _Default
    Inherits System.Web.UI.Page

    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim workbook As New XSSFWorkbook()
        Dim sheet As ISheet = workbook.CreateSheet()
        Dim headerRow As IRow = sheet.CreateRow(0)
        headerRow.CreateCell(0).SetCellValue("test")
        headerRow.CreateCell(1).SetCellValue("測試中文")

        Dim workbook2 As New XSSFWorkbook()
        Dim sheet2 As ISheet = workbook2.CreateSheet()
        Dim headerRow2 As IRow = sheet2.CreateRow(0)
        headerRow2.CreateCell(0).SetCellValue("test2")
        headerRow2.CreateCell(1).SetCellValue("測試中文2")


        Dim filename As String = "test中文.zip"
        Response.Clear()
        Response.ContentType = "application/zip"
        Response.HeaderEncoding = System.Text.Encoding.GetEncoding("utf-8")
        Response.AddHeader("content-disposition", String.Format("attachment;filename=""{0}""", filename))

        Dim ms As MemoryStream = New MemoryStream()
        '' leaves MemoryStream open
        workbook.Write(ms, True)
        ms.Position = 0

        Dim ms2 As MemoryStream = New MemoryStream()
        workbook2.Write(ms2, True)
        ms2.Position = 0

        Using zip As ZipFile = New ZipFile()
            '' 設定編碼格式，檔名如果是中文檔案才不會消失
            zip.AlternateEncoding = Encoding.GetEncoding("utf-8")
            zip.AlternateEncodingUsage = ZipOption.Always

            zip.AddEntry("中文檔名.xlsx", ms)
            zip.AddEntry("testfilename.xlsx", ms2)

            zip.Save(Response.OutputStream)
        End Using

        ms.Close()
        ms.Dispose()
        ms2.Close()
        ms2.Dispose()
        workbook = Nothing
        workbook2 = Nothing
        Response.End()
    End Sub
End Class

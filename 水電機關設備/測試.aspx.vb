Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel.XlDirection
Imports Microsoft.Office.Interop.Excel.XlPageBreak
Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Word
Imports Microsoft.VisualBasic.Logging
Imports DevExpress.Pdf
Imports Spire.Pdf
Imports System.IO
Imports System.Text
Imports System.Math
Imports System.Diagnostics
Imports System.Data.Pdf
Imports System.Data.Sql
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Text.RegularExpressions
Imports System.Web.UI.WebControls
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Data.OleDb
Imports System.Drawing.Drawing2D
Imports System.Drawing.Printing
Imports org.pdfbox.pdmodel
Imports org.pdfbox.util
Partial Class 測試
    Inherits System.Web.UI.Page
Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
        If Session("水_Uid")="3855"
                測試.Visible=True
            End If
        Else
            Me.Label1.Text = ""
            Me.Label2.Text = ""
        End If
    End Sub
' Dim myPDF as Object
Protected Sub Test(ByVal sender As Object, ByVal e As System.EventArgs)
    ' myPDF = CreateObject("acroexch.pddoc")  
    ' 'once again open the file  
    ' Dim openResult as Object = myPDF.Open(".\data\1100928秘書室便簽(新).pdf")  
    ' For pagenumber = 0 To openResult.pageCount - 1  
    '     getPDFTextFromPage(pagenumber)  
    ' Next  
    ' myPDF = Nothing
    '----
    Dim s As string = "C:\水電機關設備測試\data\PDF\通過認證時數證書-何權哲11109141423.pdf"
    Label1.Text=s
    Dim pdffile as FileInfo = new FileInfo(s)
    if (pdffile.Exists)
        Dim file as FileInfo = new FileInfo("c:\mis2000lab_example.txt")
        pdf2txt(pdffile, file)
    else
        label2.Text="The File is NOT Exist."
    End If
        '----Spire.Pdf
        ' Dim doc As PdfDocument
        ' doc.LoadFromFile("sample.pdf")
        ' Dim content As StringBuilder
        ' For Each page As PdfPageBase In doc.Pages
        ' content.Append(page.ExtractText())
        ' Next
        ' Dim fileName As String = "獲取文本.txt"
        ' File.WriteAllText(fileName, content.ToString())
        ' System.Diagnostics.Process.Start("獲取文本.txt")
        '----
        ' For Each PostedFile As HttpPostedFile In FileUpload1.PostedFiles
        '     Dim MyGUID As String = Guid.NewGuid().ToString("N")
        '     Dim Myfiles As String = MapPath(".\data\Temp\") & MyGUID
        '     PostedFile.SaveAs(Myfiles)
        '     Try
        '         File.Copy(Myfiles, MapPath(".\data\報表作業檔案\") & PostedFile.FileName, False)
        '     Catch
        '     End Try
        '     'PostedFile.FileName
        '     'Set the appropriate ContentType.
        '     Response.ContentType = "Application/pdf"
        '     'Get the physical path to the file.
        '     Dim FilePath As String = MapPath(".\data\報表作業檔案\") & PostedFile.FileName
        '     'Write the file directly to the HTTP output stream.
        '     Response.WriteFile(FilePath)
        '     System.IO.File.Delete(Myfiles)
        '     Response.End()
        '     System.IO.File.Delete(FilePath)
        ' Next
    ' End Sub
    ' Sub getPDFTextFromPage(pagenumber )  
    '     Dim myPDFPage as Object = myPDF.AcquirePage(pagenumber)  
    '     Dim myPageHilite as Object = CreateObject("acroexch.hilitelist")  
    '     Dim hiliteResult as Object = myPageHilite.Add(0, 9000)  
    '     Dim pageSelect as Object = myPDFPage.CreatePageHilite(myPageHilite)  
    '     Dim i As Integer  
    '     For i = 0 To pageSelect.GetNumText - 1  
    '         Dim pdfData as Object = pdfData & pageSelect.GetText(i)  
    '     Next  
    '     'clean up  
    '     myPDFPage = Nothing  
    '     myPageHilite = Nothing  
    '     pageSelect = Nothing  
        'getPDFTextFromPage = getPDFTextFromPage=pdfData  
    End Sub
    Public Sub pdf2txt(ByVal file As FileInfo,ByVal txtfile As FileInfo)'以月取日，收尋，日可不留白
        Dim doc As PDDocument = PDDocument.load(file.FullName)
        Dim pdfStripper As PDFTextStripper = new PDFTextStripper()
        Dim text As string = pdfStripper.getText(doc)
        Dim swPdfChange As StreamWriter = new StreamWriter(txtfile.FullName, false, Encoding.GetEncoding(65001))
        swPdfChange.Write(text)
        swPdfChange.Close()
    End Sub
End Class
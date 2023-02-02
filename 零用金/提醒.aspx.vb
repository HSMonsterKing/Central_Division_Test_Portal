Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel.XlDirection
Imports Microsoft.Office.Interop.Excel.XlPageBreak
Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.VisualBasic.Logging
Imports System.IO
Imports System.Math
Imports System.Diagnostics
Imports System.Data.Sql
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Text.RegularExpressions
Imports System.Web.UI.WebControls
Imports System.Drawing
Partial Class 提醒
    Inherits System.Web.UI.Page
    Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        data.ConnectionString = con_14
        If Not Page.IsPostBack Then
        End If
        If Session("atype") = "IsUserLogin" Or Session("atype") = "all"
            Panel3.Visible=true
            data.ConnectionString = con_14
            data.SelectCommand = "SELECT * FROM 分配明細表 WHERE (零用金使用單位 = '秘書室') AND (金額/2)> (SELECT top 1 餘額 FROM 收支備查簿 Where _種類='A' order by id Desc)"
            data_dv = data.Select(New DataSourceSelectArguments)
            If data_dv.Count>0
                Label3.text="本部零用金已消耗二分之一，請盡速核銷"
            End If 
        End If 
        If Session("atype") = "IsDirectorLogin" Or Session("atype") = "all"
        Panel4.Visible=true
        End If 
        If Session("atype") = "IsAccountantLogin" Or Session("atype") = "all"
        Panel5.Visible=true
        End If 
    End Sub
End Class
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
Imports System.Drawing.Imaging
Imports System.Data.OleDb
Partial Class 照片
    Inherits System.Web.UI.Page
    Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        data.ConnectionString = con_14
        If Not Page.IsPostBack Then
            Image1.ImageUrl=Session("水_照片")
        Else
            Me.Label1.Text = ""
            Me.Label2.Text = ""
        End If
    End Sub
    Protected Sub 原始照片_OnClick(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim 壓縮照片 As String=Session("水_照片")
        If O_Image.Visible = False
            data.ConnectionString = con_14
            data.SelectCommAnd = "SELECT 原始照片 FROM 水電機關設備資料表 Where 壓縮照片 = '" & 壓縮照片 & " ' AND  原始照片 Is Not NULL"
            data_dv = data.Select(New DataSourceSelectArguments)
            If data_dv.count > 0
                O_Image.Visible=True
                O_Image.ImageUrl=data_dv(0)("原始照片").ToString
                Image1.ImageUrl=""
                Image1.Visible=False
                返回.Visible=True
            Else
                label2.Text="此圖片即為原始照片"
            End If
            原始照片.Visible=False
        End If
    End Sub
Protected Sub 返回_OnClick(ByVal sender As Object, ByVal e As System.EventArgs)'
        Dim 壓縮照片 As String=Session("水_照片")
        If O_Image.Visible = True
            O_Image.Visible=False
            O_Image.ImageUrl=""
            Image1.ImageUrl=壓縮照片
            Image1.Visible=True
            原始照片.Visible=True
            返回.Visible=False
        End If
    End Sub
End Class
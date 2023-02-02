Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel.XlDirection
Imports Microsoft.Office.Interop.Excel.XlPageBreak
Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.VisualBasic.Logging
' Imports Microsoft.Office.Interop.Outlook.O1Mail
' Imports OutLook = Microsoft.Office.Interop.Outlook
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
Partial Class 中分入口網
    Inherits System.Web.UI.Page
    ' Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
    ' Dim data As New SqlDataSource
    ' Dim data_dv As Data.DataView
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ' data.ConnectionString = CFW_wf2
        If Not Page.IsPostBack Then
            Dim r1, r2, r3 As Integer
            Dim _登入帳號 As String
            _登入帳號 = Request.ServerVariables("REMOTE_HOST")
            If Len(Trim(_登入帳號)) <= 0 Then
                _登入帳號 = System.Security.Principal.WindowsIdentity.GetCurrent().Name
            End If
            r1 = Len(_登入帳號)
            For i As Integer = 1 To r1
                If Mid(_登入帳號, i, 1) = "\" Then
                    r2 = i
                End If
            Next
            r3 = r1 - r2
            _登入帳號 = Right(_登入帳號, r3)
            If _登入帳號="10.52.3.155" Or _登入帳號="10.52.0.178"
                出納對帳系統.Visible=true
                稽催寄送.Visible=true
                大宗郵件.Visible=true
            End If
            If _登入帳號="10.52.10.210"
                零用金.Visible=false
                水電設備管理.Visible=false
                出納對帳系統.Visible=true
                大宗郵件.Visible=true
            End If
            If  _登入帳號="10.52.10.63" Or _登入帳號="10.52.10.64"
                Response.Redirect("~/零用金/Default.aspx")                
            End If
            If  _登入帳號="10.52.10.170"
                Response.Redirect("~/大宗郵件/Default.aspx")                
            End If
            If _登入帳號="10.52.10.79" Or _登入帳號="10.52.10.180" Or _登入帳號="10.52.3.91" Or _登入帳號="10.52.3.92" Or _登入帳號="10.52.3.93"
                Response.Redirect("~/出納對帳系統/Default.aspx")
            End If
            If _登入帳號="10.52.10.167"
                Response.Redirect("~/稽催寄送/Default.aspx")
            End If
        Else
            Me.Label1.Text = ""
            Me.Label2.Text = ""
        End If 
    End Sub
    Protected Sub 零用金_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Response.Redirect("~/零用金/Default.aspx")
    End Sub
    Protected Sub 水電設備管理_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Response.Redirect("~/水電機關設備/Default.aspx")
    End Sub
    Protected Sub 出納對帳系統_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Response.Redirect("~/出納對帳系統/Default.aspx")
    End Sub
    Protected Sub 大宗郵件_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Response.Redirect("~/大宗郵件/Default.aspx")
    End Sub
    Protected Sub 稽催寄送_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Response.Redirect("~/稽催寄送/Default.aspx")
    End Sub
End Class
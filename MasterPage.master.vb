Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel.XlDirection
Imports Microsoft.VisualBasic.Logging
Imports System.IO
Imports System.Drawing
Imports System.Diagnostics
Imports System.Data.Sql
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Text.RegularExpressions
Partial Class MasterPage
    Inherits System.Web.UI.MasterPage
    ' Dim CFW_wf2 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ' data.ConnectionString = CFW_wf2
        If Not Page.IsPostBack Then
        End If 
    End Sub
End Class

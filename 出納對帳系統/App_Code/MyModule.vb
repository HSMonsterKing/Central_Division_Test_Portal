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
Public Module MyModule
    Public Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub
    Public Function taiwancalendarto(ByVal _年月日 As String) As String
        If _年月日 Is Nothing
            Return _年月日
        End If
        Dim regex as new Regex("^[0-9]{1,3}/[0-9]{1,2}/[0-9]{1,2}")
        If regex.IsMatch(_年月日)
            _年月日 = _年月日.Split(" ")(0)
            Return (CLng(_年月日.Split("/")(0) + 1911).ToString() & "/" & CLng(_年月日.Split("/")(1)).ToString("00") & "/" & CLng(_年月日.Split("/")(2)).ToString("00"))
        Else
            Return _年月日
        End If
    End Function
    Public Function totaiwancalendar(ByVal _年月日 As String) As String
        If _年月日 Is Nothing
            Return _年月日
        End If
        Dim regex as new Regex("^[0-9]{4}/[0-9]{1,2}/[0-9]{1,2}")
        If regex.IsMatch(_年月日)
            _年月日 = _年月日.Split(" ")(0)
            Return (CLng(_年月日.Split("/")(0) - 1911).ToString() & "/" & CLng(_年月日.Split("/")(1)).ToString("00") & "/" & CLng(_年月日.Split("/")(2)).ToString("00"))
        Else
            Return _年月日
        End If
    End Function
End Module
Attribute VB_Name = "Mod_Reporte"
Option Explicit
Public Declare Function CreateFieldDefFile Lib "p2smon.dll" (oRst As Object, ByVal FileName As String, ByVal bOverWriteExistingFile As Integer) As Long

Public Function ArrFormulas(ByVal STRnombreParametro As String, ByVal STRvalorParametro As Variant)
   ArrFormulas = Array(STRnombreParametro, STRvalorParametro)
End Function

Public Function SourceReport(ByVal STRnombre As String, ByVal STRruta As String, ByVal RSdata As ADODB.Recordset)
   SourceReport = Array(STRnombre, STRruta, RSdata)
End Function

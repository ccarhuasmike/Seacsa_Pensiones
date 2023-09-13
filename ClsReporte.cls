VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function CargaReporte(ByVal vruta As String, ByVal vnombre As String, ByVal vtitulo As String, _
                        ByVal vrs As ADODB.Recordset, ByVal mostrarSetup As Boolean, _
                        ParamArray Parametros() As Variant) As Boolean

On Error GoTo mierror

'MARCO VARGAS 11-03-2010
Dim crApplication As New CRAXDRT.Application
Dim crpReport As CRAXDRT.Report ' Variable para abrir el reporte
Dim crpDatabase As CRAXDRT.Database ' Variable para la Base de Datos
Dim crpTables As CRAXDRT.DatabaseTables ' Variable para tablas de la base de datos
Dim crpTable As CRAXDRT.DatabaseTable
Dim ARRparametro As Variant
Dim STRnomparametro As String
Dim objReporte As New frmReporte

'Set crpApplication = CreateObject("crystal.CRPE.application")
Set crpReport = crApplication.OpenReport(vruta & vnombre)

Dim y As Integer
y = 1
   ARRparametro = Parametros()
   For INTc = LBound(ARRparametro) To UBound(ARRparametro)
          STRnomparametro = ARRparametro(INTc)(0)
         crpReport.ParameterFields(y).AddCurrentValue ARRparametro(INTc)(1)
         'crpReport.FormulaFields.Item(y).Text = "'" & ARRparametro(INTc)(1) & "'"
      y = y + 1
   Next INTc
   
sigue:
Set crpDatabase = crpReport.Database
Set crpTables = crpDatabase.Tables
Set crpTable = crpTables.Item(1)


Call crpTable.SetPrivateData(3, vrs)

objReporte.Setup = mostrarSetup
objReporte.Save = mostrarGuardar
objReporte.SetReporte crpReport
objReporte.Caption = vtitulo
objReporte.Show


Set objReporte = Nothing
Set crApplication = Nothing
Set crpReport = Nothing
Set crpDatabase = Nothing
Set crParamDefs = Nothing
Set crParamDef = Nothing

CargaReporte = True
Exit Function
mierror:
    If Err.Number = 9 Then
      GoTo sigue
    End If
    MsgBox Err.Description & " " & Err.Number, vbExclamation, "ReporteRS"
    
    
End Function



Public Function CargaReporte_SubReporte1(ByVal vruta As String, ByVal vnombre As String, ByVal vtitulo As String, _
                        ByVal vrs As ADODB.Recordset, ByVal mostrarSetup As Boolean, _
                        ByVal Name_subreport As String, ByVal rsSubReporte1 As ADODB.Recordset, _
                        ParamArray Parametros() As Variant) As Boolean

On Error GoTo mierror

'variables para carga del reporte
Dim crApplication As New CRAXDRT.Application
Dim crpReport As CRAXDRT.Report ' Variable para abrir el reporte
Dim crpDatabase As CRAXDRT.Database ' Variable para la Base de Datos
Dim crpTables As CRAXDRT.DatabaseTables ' Variable para tablas de la base de datos
Dim crpTable As CRAXDRT.DatabaseTable

'para sub reportes
Dim crpSections As CRAXDRT.Sections
Dim crpSection As CRAXDRT.section
Dim crpReportObjs As CRAXDRT.ReportObjects
Dim crpSubReportObjs As CRAXDRT.SubreportObject
Dim crpSubReport As CRAXDRT.Report

'otras variables
Dim ARRparametro As Variant
Dim STRnomparametro As String
Dim objReporte As New frmReporte


'Set crpApplication = CreateObject("crystal.crpe.application")
Set crpReport = crApplication.OpenReport(vruta & vnombre)

'leyendo formulas
Dim y As Integer
y = 1
   ARRparametro = Parametros()
   For INTc = LBound(ARRparametro) To UBound(ARRparametro)
          STRnomparametro = ARRparametro(INTc)(0)
          crpReport.FormulaFields.Item(y).Text = "'" & ARRparametro(INTc)(1) & "'"
      y = y + 1
   Next INTc
 
Set crpDatabase = crpReport.Database
Set crpTables = crpDatabase.Tables
Set crpTable = crpTables.Item(1)
'call crpTable.SetPrivateData(3, vrs)
crpTable.SetDataSource vrs, 3

Dim z As Integer
Dim x As Integer


'codigo para sub reportes
Set crpSections = crpReport.Sections

For x = 1 To crpReport.Sections.Count
    Set crpSection = crpSections.Item(x)
    Set crpReportObjs = crpSection.ReportObjects
    For z = 1 To crpReportObjs.Count
        If crpReportObjs.Item(z).Kind = 5 Then  '5=subreporte
            Set crpSubReportObjs = crpReportObjs.Item(z)
            Set crpSubReport = crpReport.OpenSubreport(Name_subreport)
            Set crpDatabase = crpSubReport.Database
            Set crpTables = crpDatabase.Tables
            Set crpTable = crpTables.Item(1)
            crpTable.SetDataSource rsSubReporte1, 3
        End If
    Next
Next


objReporte.Setup = mostrarSetup
objReporte.Save = mostrarGuardar
objReporte.SetReporte crpReport
objReporte.Caption = vtitulo
objReporte.Show


Set objReporte = Nothing
Set crApplication = Nothing
Set crpReport = Nothing
Set crpDatabase = Nothing
Set crParamDefs = Nothing
Set crParamDef = Nothing

CargaReporte_SubReporte1 = True
Exit Function
mierror:
    MsgBox Err.Description, vbExclamation, "ReporteRS"
    
    
End Function



Public Function CargaReporte_SubReporte2(ByVal vruta As String, ByVal vnombre As String, ByVal vtitulo As String, _
                        ByVal vrs As ADODB.Recordset, ByVal mostrarSetup As Boolean, _
                        ByVal Name_subreport As String, ByVal rsSubReporte1 As ADODB.Recordset, _
                        ByVal Name_subreport2 As String, ByVal rsSubReporte2 As ADODB.Recordset, _
                        ParamArray Parametros() As Variant) As Boolean

On Error GoTo mierror

'variables para carga del reporte
Dim crApplication As New CRAXDRT.Application
Dim crpReport As CRAXDRT.Report ' Variable para abrir el reporte
Dim crpDatabase As CRAXDRT.Database ' Variable para la Base de Datos
Dim crpTables As CRAXDRT.DatabaseTables ' Variable para tablas de la base de datos
Dim crpTable As CRAXDRT.DatabaseTable

'para sub reportes
Dim crpSections As CRAXDRT.Sections
Dim crpSection As CRAXDRT.section
Dim crpReportObjs As CRAXDRT.ReportObjects
Dim crpSubReportObjs As CRAXDRT.SubreportObject
Dim crpSubReport As CRAXDRT.Report

'otras variables
Dim ARRparametro As Variant
Dim STRnomparametro As String
Dim objReporte As New frmReporte


'Set crpApplication = CreateObject("crystal.crpe.application")
Set crpReport = crApplication.OpenReport(vruta & vnombre)

'leyendo formulas
Dim y As Integer
y = 1
   ARRparametro = Parametros()
   For INTc = LBound(ARRparametro) To UBound(ARRparametro)
          STRnomparametro = ARRparametro(INTc)(0)
          crpReport.FormulaFields.Item(y).Text = "'" & ARRparametro(INTc)(1) & "'"
      y = y + 1
   Next INTc
 
Set crpDatabase = crpReport.Database
Set crpTables = crpDatabase.Tables
Set crpTable = crpTables.Item(1)
crpTable.SetDataSource vrs, 3

Dim z As Integer
Dim x As Integer


'codigo para sub reportes
Set crpSections = crpReport.Sections

For x = 1 To crpReport.Sections.Count
    Set crpSection = crpSections.Item(x)
    Set crpReportObjs = crpSection.ReportObjects
    For z = 1 To crpReportObjs.Count
        If crpReportObjs.Item(z).Kind = 5 Then  '5=subreporte
            Set crpSubReportObjs = crpReportObjs.Item(z)
            If crpSubReportObjs.SubreportName = Name_subreport Then
                Set crpSubReport = crpReport.OpenSubreport(Name_subreport)
                Set crpDatabase = crpSubReport.Database
                Set crpTables = crpDatabase.Tables
                Set crpTable = crpTables.Item(1)
                crpTable.SetDataSource rsSubReporte1, 3
            ElseIf UCase(crpSubReportObjs.SubreportName) = UCase(Name_subreport2) Then
                Set crpSubReport = crpReport.OpenSubreport(Name_subreport2)
                Set crpDatabase = crpSubReport.Database
                Set crpTables = crpDatabase.Tables
                Set crpTable = crpTables.Item(1)
                crpTable.SetDataSource rsSubReporte2, 3
            End If
        End If
    Next
Next


objReporte.Setup = mostrarSetup
objReporte.Save = mostrarGuardar
objReporte.SetReporte crpReport
objReporte.Caption = vtitulo
objReporte.Show


Set objReporte = Nothing
Set crApplication = Nothing
Set crpReport = Nothing
Set crpDatabase = Nothing
Set crParamDefs = Nothing
Set crParamDef = Nothing

CargaReporte_SubReporte2 = True
Exit Function
mierror:
    MsgBox Err.Description, vbExclamation, "ReporteRS"
    
    
End Function

Public Function CargaReporte_toPdf(ByVal vruta As String, ByVal vnombre As String, ByVal vtitulo As String, vPeriodo As String, ByVal vPoliza As String, _
                        ByVal vrs As ADODB.Recordset, ByVal mostrarSetup As Boolean, ByVal ruta As String, _
                        ParamArray Parametros() As Variant) As Boolean

On Error GoTo mierror


CargaReporte_toPdf = True

'MARCO VARGAS 11-03-2010
Dim crApplication As New CRAXDRT.Application
Dim crpReport As CRAXDRT.Report ' Variable para abrir el reporte
Dim crpDatabase As CRAXDRT.Database ' Variable para la Base de Datos
Dim crpTables As CRAXDRT.DatabaseTables ' Variable para tablas de la base de datos
Dim crpTable As CRAXDRT.DatabaseTable
Dim ARRparametro As Variant
Dim STRnomparametro As String
Dim objReporte As New frmReporte

'Set crpApplication = CreateObject("crystal.CRPE.application")
Set crpReport = crApplication.OpenReport(vruta & vnombre)
    
    
Dim y As Integer
y = 1
   ARRparametro = Parametros()
   For INTc = LBound(ARRparametro) To UBound(ARRparametro)
          STRnomparametro = ARRparametro(INTc)(0)
         crpReport.ParameterFields(y).AddCurrentValue ARRparametro(INTc)(1)
         'crpReport.FormulaFields.Item(y).Text = "'" & ARRparametro(INTc)(1) & "'"
      y = y + 1
   Next INTc
   
sigue:
Set crpDatabase = crpReport.Database
Set crpTables = crpDatabase.Tables
Set crpTable = crpTables.Item(1)
            

            
'Set crpReport = Nothing

Call crpTable.SetPrivateData(3, vrs)

'objReporte.Setup = mostrarSetup
'objReporte.Save = mostrarGuardar
'objReporte.SetReporte crpReport
'objReporte.Caption = vtitulo
'objReporte.Show

'SetReportDatabase crpReport
With crpReport
    .ExportOptions.PDFExportAllPages = True
    .ExportOptions.DestinationType = crEDTDiskFile
    .ExportOptions.DiskFileName = ruta 'App.Path & "\PDF\" & vPeriodo & "\Boleta_" & vPoliza & ".pdf"
    .ExportOptions.FormatType = crEFTPortableDocFormat
    .DisplayProgressDialog = False
    .Export False
    
End With
            
Set objReporte = Nothing
Set crApplication = Nothing
Set crpReport = Nothing
Set crpDatabase = Nothing
Set crParamDefs = Nothing
Set crParamDef = Nothing

Exit Function
mierror:
    If Err.Number = 9 Then
      GoTo sigue
    End If
    MsgBox Err.Description & " " & Err.Number, vbExclamation, "ReporteRS"
    
    CargaReporte_toPdf = False
    
End Function

Public Function EjecutaReporte(ByVal vruta As String, ByVal vnombre As String, ByVal vtitulo As String, _
                        ByVal vrs As ADODB.Recordset, ByVal flgSRpt As Integer, ByVal RsSubRpt As ADODB.Recordset, ParamArray Parametros() As Variant) As Boolean
'    Dim objReporte As New frmReporte

Dim crApplication As New CRAXDRT.Application
Dim objReporte As New frmReporte
Dim CrReport As CRAXDRT.Report
Dim CrSubreport As CRAXDRT.Report

Dim crDatabase As CRAXDRT.Database
Dim CrpSRDB As CRAXDRT.Database
    
Dim crRepObjs As CRAXDRT.ReportObjects
Dim crSubRepObj As CRAXDRT.SubreportObject

Dim crSecs As CRAXDRT.Sections
Dim crSec As CRAXDRT.section

Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions
Dim crParamDef As CRAXDRT.ParameterFieldDefinition

    Set CrReport = crApplication.OpenReport(vruta & vnombre)
    Set crParamDefs = CrReport.ParameterFields
    For Each crParamDef In crParamDefs
        With crParamDef
            Select Case UCase(.ParameterFieldName)
                Case "NOMBRECOMPANIA"
                    .SetCurrentValue Parametros(0)(1)
                Case "NOMBRESISTEMA"
                    .SetCurrentValue Parametros(1)(1)
                Case "NOMBRESUBSISTEMA"
                    .SetCurrentValue Parametros(2)(1)
                Case "TIPPENSION"
                    .SetCurrentValue Parametros(3)(1)
                Case "AFP"
                    .SetCurrentValue Parametros(4)(1)
                Case "TIPRTA"
                    .SetCurrentValue Parametros(5)(1)
                Case "NOMBRECAUSANTE"
                    .SetCurrentValue Parametros(6)(1)
                Case "RUTCAUSANTE"
                    .SetCurrentValue Parametros(7)(1)
                Case "DIRECCION"
                    .SetCurrentValue Parametros(8)(1)
                Case "FONO"
                    .SetCurrentValue Parametros(9)(1)
                Case "COMUNA"
                    .SetCurrentValue Parametros(10)(1)
                Case "PROVINCIA"
                    .SetCurrentValue Parametros(11)(1)
                Case "REGION"
                    .SetCurrentValue Parametros(12)(1)
                Case "ORIGEN"
                    .SetCurrentValue Parametros(13)(1)
                Case "CODMONEDA"
                    .SetCurrentValue Parametros(14)(1)
                Case "MOTIVOENDOSO"
                    .SetCurrentValue Parametros(15)(1)
                Case "GLSFACTORENDOSO"
                    .SetCurrentValue Parametros(16)(1)
                Case "MTORTAMOD"
                    .SetCurrentValue Parametros(17)(1)
                Case "MTOPENSION"
                    .SetCurrentValue Parametros(18)(1)
                Case "FECHAVIGENDOSO"
                    .SetCurrentValue Parametros(19)(1)
                Case "MTORTAORI"
                    .SetCurrentValue Parametros(20)(1)
                Case "CODMONEDACOR"
                    .SetCurrentValue Parametros(21)(1)
                Case "FLGSRPT"
                    .SetCurrentValue flgSRpt
            End Select
        End With
    Next
    
    Dim RsSubRp As ADODB.Recordset
    Set crDatabase = CrReport.Database
    Call crDatabase.Tables(1).SetPrivateData(3, vrs)
    
    Set crSecs = CrReport.Sections
    rx = 0
    If flgSRpt = 1 Then
        For Ix = 1 To crSecs.Count
            Set crSec = crSecs.Item(Ix)
            Set crRepObjs = crSec.ReportObjects
            For Jx = 1 To crRepObjs.Count
                If crRepObjs.Item(Jx).Kind = crSubreportObject Then
    '                Set RsSubRp = Devuelve_Select(vCadSql(rx))
                    Set CrSubreport = CrReport.OpenSubreport(crRepObjs(Jx).SubreportName)
                    Set CrpSRDB = CrSubreport.Database
                    Call CrpSRDB.Tables(1).SetPrivateData(3, RsSubRpt)
    '                rx = rx + 1
                End If
            Next
        Next
    End If
    
    CrReport.ReportTitle = vtitulo
    objReporte.Setup = True
    objReporte.SetReporte CrReport
    objReporte.Show 1
    
    Set objReporte = Nothing
    Set crApplication = Nothing
    Set CrReport = Nothing
    Set crDatabase = Nothing
    Set crParamDefs = Nothing
    Set crParamDef = Nothing
    EjecutaReporte = True
    Exit Function
End Function

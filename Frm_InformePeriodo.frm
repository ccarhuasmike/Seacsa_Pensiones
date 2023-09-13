VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "Crystl32.OCX"
Begin VB.Form Frm_InformePeriodo 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   5880
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   5655
      Begin VB.TextBox txtFecmanu 
         Height          =   285
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Vencimiento"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Carta"
      Height          =   675
      Left            =   2040
      Picture         =   "Frm_InformePeriodo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   720
   End
   Begin VB.Frame Fra_Operaciones 
      Height          =   1095
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   5655
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   840
         Picture         =   "Frm_InformePeriodo.frx":06BA
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   255
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   3000
         Picture         =   "Frm_InformePeriodo.frx":0D74
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   4080
         Picture         =   "Frm_InformePeriodo.frx":142E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
      Begin Crystal.CrystalReport Rpt_Reporte 
         Left            =   240
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowState     =   2
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.Frame Fra_Datos 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.OptionButton op2 
         Caption         =   "Sin Renovaciones"
         Height          =   255
         Left            =   3000
         TabIndex        =   11
         Top             =   855
         Width           =   1830
      End
      Begin VB.OptionButton op1 
         Caption         =   "Con Renovaciones"
         Height          =   255
         Left            =   990
         TabIndex        =   10
         Top             =   840
         Width           =   1770
      End
      Begin VB.TextBox Txt_Desde 
         Height          =   285
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Txt_Hasta 
         Height          =   285
         Left            =   4080
         MaxLength       =   10
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   8
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Lbl_Contrato 
         Caption         =   "Período (Desde - Hasta)  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   2295
      End
   End
End
Attribute VB_Name = "Frm_InformePeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vlRegistro As New ADODB.Recordset
Dim vlRsStruc As New ADODB.Recordset
Dim vlArchivo       As String
Dim vlFechaInicio As String
Dim vlFechaTermino As String
Dim vlFechaVenManu As String

Dim vlGlosaEdadBen As String

Dim vlFecIni24 As String
Dim vlFecTer24 As String

Dim vlFecIni18 As String
Dim vlFecTer18 As String

Const clEdad18 As String = 18
Const clEdad24 As String = 24
Private Sub crear_rs_struc()
   
   Set vlRsStruc = New ADODB.Recordset
   
   vlRsStruc.Fields.Append "NUM_POLIZA", adVarChar, 10
   vlRsStruc.Fields.Append "NUM_ORDEN", adVarChar, 10
   vlRsStruc.Fields.Append "GLS_TIPOIDENCOR", adVarChar, 10
   vlRsStruc.Fields.Append "NUM_IDENBEN", adVarChar, 16
   vlRsStruc.Fields.Append "GLS_PATBEN", adVarChar, 20
   vlRsStruc.Fields.Append "GLS_MATBEN", adVarChar, 20
   vlRsStruc.Fields.Append "GLS_NOMBEN", adVarChar, 25
   vlRsStruc.Fields.Append "GLS_NOMSEGBEN", adVarChar, 20
   vlRsStruc.Fields.Append "FEC_INICER", adVarChar, 8
   vlRsStruc.Fields.Append "FEC_TERCER", adVarChar, 8
   vlRsStruc.Fields.Append "DIFERENCIA", adVarChar, 8
   vlRsStruc.Fields.Append "DIRECCION", adVarChar, 100
   vlRsStruc.Fields.Append "DISTRITO", adVarChar, 50
   vlRsStruc.Fields.Append "NUMPAG", adVarChar, 5
      
   vlRsStruc.Open
  

End Sub
Function flInformeCerEst()
On Error GoTo Err_flInformeCerEst
'Certificados de Supervivencia
   Screen.MousePointer = 11
   
   'marco 11/03/2010
   Dim cadena As String
   Dim objRep As New ClsReporte
   vlFechaInicio = Format(CDate(Trim(Txt_Desde.Text)), "yyyymmdd")
   vlFechaTermino = Format(CDate(Trim(Txt_Hasta.Text)), "yyyymmdd")
   vgPalabra = ""
   vgPalabra = Txt_Desde.Text & "   *   " & Txt_Hasta.Text
   
'   cadena = "select c.NUM_POLIZA,c.NUM_ORDEN,t.GLS_TIPOIDENCOR,b.NUM_IDENBEN,b.GLS_PATBEN,b.GLS_MATBEN,b.GLS_NOMBEN,b.GLS_NOMSEGBEN," & _
'            " c.FEC_INICER,TRUNC(TO_DATE(SUBSTR(FEC_TERCER,7,2)||SUBSTR(FEC_TERCER,5,2)||SUBSTR(FEC_TERCER,1,4),'dd/mm/yyyy')-sysdate) AS DIFERENCIA " & _
'            " from pp_tmae_certificado c left join pp_tmae_ben b on c.NUM_POLIZA=b.NUM_POLIZA " & _
'            " and c.NUM_ENDOSO=b.NUM_ENDOSO and c.NUM_ORDEN=b.NUM_ORDEN left join ma_tpar_tipoiden t on " & _
'            " b.cod_tipoidenben=t.cod_tipoiden where c.fec_tercer>='" & vlFechaInicio & "' and c.fec_tercer<='" & vlFechaTermino & "'"
            
'    Dim cm As ADODB.Command
'    Set cm = New ADODB.Command
'    cm.ActiveConnection = vgConexionBD
'    cm.CommandType = adCmdText
'    cm.CommandText = "{call RENTASVD.PK_LISTA_CERT_POR_CADUCAR.LISTAR}"
'    cm.Execute
'
'    vlRegistro.Open "PK_LISTA_CERT_POR_CADUCAR.LISTAR", vgConexionBD, adOpenStatic, adLockReadOnly
    
    'cadena = "{Call PK_LISTA_CERT_POR_CADUCAR.LISTAR({ResultSet 0, lista1})}"
    
    Call crear_rs_struc
    Dim var As String
    var = IIf(op1.Value = True, "S", "N")
    cadena = "{CALL PK_LISTA_CERT_POR_CADUCAR.LISTAR(" & vlFechaInicio & "," & vlFechaTermino & ",'" & var & "')}"
    
    'Set vlRegistro = Server.CreateObject("ADODB.Recordset")
    Set vlRegistro = New ADODB.Recordset
    vlRegistro.Open cadena, vgConexionBD, 3, 1
    
    Dim counta As Integer
    
    counta = 1
    
        If Not vlRegistro.EOF Then
            Do Until vlRegistro.EOF
                vlRsStruc.AddNew
                
                vlRsStruc.Fields("NUM_POLIZA").Value = Trim(vlRegistro!num_poliza)
                vlRsStruc.Fields("NUM_ORDEN").Value = vlRegistro!Num_Orden
                vlRsStruc.Fields("GLS_TIPOIDENCOR").Value = vlRegistro!gls_tipoidencor
                vlRsStruc.Fields("NUM_IDENBEN").Value = vlRegistro!Num_IdenBen
                vlRsStruc.Fields("GLS_PATBEN").Value = vlRegistro!Gls_PatBen
                vlRsStruc.Fields("GLS_MATBEN").Value = vlRegistro!Gls_MatBen
                vlRsStruc.Fields("GLS_NOMBEN").Value = vlRegistro!Gls_NomBen
                vlRsStruc.Fields("GLS_NOMSEGBEN").Value = IIf(IsNull(vlRegistro!Gls_NomSegBen), "", vlRegistro!Gls_NomSegBen)
                vlRsStruc.Fields("FEC_INICER").Value = vlRegistro!fec_inicer
                vlRsStruc.Fields("FEC_TERCER").Value = vlRegistro!FEC_TERCER
                vlRsStruc.Fields("DIFERENCIA").Value = vlRegistro!DIFERENCIA
                vlRsStruc.Fields("DIRECCION").Value = vlRegistro!Gls_DirBen
                vlRsStruc.Fields("DISTRITO").Value = vlRegistro!gls_comuna
                vlRsStruc.Fields("NUMPAG").Value = counta
                counta = counta + 1
                vlRsStruc.Update
                vlRegistro.MoveNext
            Loop
        End If
    
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(vlRsStruc, Replace(UCase(strRpt & "Estructura\PP_Rpt_InfoCerSupCaducar.rpt"), ".RPT", ".TTX"), 1)
    
        'If objRep.CargaReporte(strRpt & "PP_Rpt_InfoCerSupCaducar.rpt", "Informe Certificados de Supervivencia por Caducar", "Informe Certificados de Supervivencia por Caducar", vlRsStruc, True) = False Then
    If objRep.CargaReporte(strRpt & "", "PP_Rpt_InfoCerSupCaducar.rpt", "Informe Certificados de Supervivencia por Caducar", vlRsStruc, True, _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("NombreSistema", vgNombreSistema), _
                            ArrFormulas("NombreSubSistema", vgNombreSubSistema), _
                            ArrFormulas("Periodo", vgPalabra), _
                            ArrFormulas("GlosaQuiebra", vgGlosaQuiebra)) = False Then
                            
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Function
    End If
    
    'fin marco
    
'
'   vlArchivo = strRpt & "PP_Rpt_InfoCerSupCaducar.rpt"   '\Reportes
'   If Not fgExiste(vlArchivo) Then     ', vbNormal
'      MsgBox "Archivo de Reporte de Certificados de Estudio por Caducar no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
'      Screen.MousePointer = 0
'      Exit Function
'   End If
'
'   Call fgVigenciaQuiebra(Txt_Desde)
'
'   vlFechaInicio = Format(CDate(Trim(Txt_Desde.Text)), "yyyymmdd")
'   vlFechaTermino = Format(CDate(Trim(Txt_Hasta.Text)), "yyyymmdd")
'
'   vgQuery = ""
'   vgQuery = vgQuery & "{PP_TMAE_CERTIFICADO.fec_tercer} >= '" & Trim(vlFechaInicio) & "' AND "
'   vgQuery = vgQuery & "{PP_TMAE_CERTIFICADO.fec_tercer} <= '" & Trim(vlFechaTermino) & "' "
'
'
'   Rpt_Reporte.Reset
'   Rpt_Reporte.WindowState = crptMaximized
'   Rpt_Reporte.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
'   Rpt_Reporte.Connect = vgRutaDataBase       ' o App.Path & "\Nestle.mdb"
'   Rpt_Reporte.SelectionFormula = vgQuery
'
'   vgPalabra = ""
'   vgPalabra = Txt_Desde.Text & "   *   " & Txt_Hasta.Text
'
'   Rpt_Reporte.Formulas(0) = ""
'   Rpt_Reporte.Formulas(1) = ""
'   Rpt_Reporte.Formulas(2) = ""
'   Rpt_Reporte.Formulas(3) = ""
'   Rpt_Reporte.Formulas(4) = ""
'
'   Rpt_Reporte.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
'   Rpt_Reporte.Formulas(1) = "NombreSistema= '" & vgNombreSistema & "'"
'   Rpt_Reporte.Formulas(2) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
'   Rpt_Reporte.Formulas(3) = "Periodo = '" & vgPalabra & "'"
'   Rpt_Reporte.Formulas(4) = "GlosaQuiebra = '" & vgGlosaQuiebra & "'"
'
'   Rpt_Reporte.SubreportToChange = ""
'   Rpt_Reporte.Destination = crptToWindow
'   Rpt_Reporte.WindowState = crptMaximized
'   Rpt_Reporte.WindowTitle = "Informe Certificados de Supervivencia por Caducar"
''   Rpt_Reporte.SelectionFormula = ""
'   Rpt_Reporte.Action = 1
'   Screen.MousePointer = 0

Exit Function
Err_flInformeCerEst:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Private Sub Cmd_Imprimir_Click()
On Error GoTo errImprimir
    
    If Trim(Txt_Desde) = "" Then
        MsgBox "Falta Ingresar Fecha Desde ", vbCritical, "Falta Información"
        Txt_Desde.SetFocus
        Exit Sub
    Else
        If Not IsDate(Txt_Desde) Then
            MsgBox "Fecha Desde no es una Fecha válida", vbCritical, "Falta Información"
            Txt_Desde.SetFocus
            Exit Sub
        End If
    End If
    If Trim(Txt_Hasta) = "" Then
        MsgBox "Falta Ingresar Fecha Hasta", vbCritical, "Falta Información"
        Txt_Hasta.SetFocus
        Exit Sub
    Else
        If Not IsDate(Txt_Hasta) Then
            MsgBox "Fecha Hasta no es una fecha Válida", vbCritical, "Falta Información"
            Txt_Hasta = ""
            Exit Sub
        End If
    End If
        
    Screen.MousePointer = 11
    
    vlFechaInicio = Format(CDate(Trim(Txt_Desde)), "yyyymmdd")
    vlFechaTermino = Format(CDate(Trim(Txt_Hasta)), "yyyymmdd")


    



    'Permite imprimir la Opción Indicada a través del Menú
    Select Case vgNombreInformePeriodoSeleccionado
        Case "InfCerEstCad"    'Informe de Certificados de Estudio por Caducar
            Call flInformeCerEst
        Case "InfBen18"    'Informe de Beneficiarios por cumplir 18 años de edad
            Call flInformeBen18
'        Case "InfBen24"    'Informe de Beneficiarios por cumplir 24 años de edad
'            Call flInformeBen24
        
    End Select
    
    Screen.MousePointer = 0

Exit Sub
errImprimir:
    Screen.MousePointer = 0
    If Err.Number <> 0 Then
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End If
End Sub


Function flInformeBen18()
On Error GoTo Err_flInformeBen18
                    
   Screen.MousePointer = 11
   
   vlArchivo = strRpt & "PP_Rpt_InfoBenEdad18_24.rpt"   '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
      MsgBox "Archivo de Reporte de Beneficiarios por cumplir 24 Años de Edad no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
      Screen.MousePointer = 0
      Exit Function
   End If
   
   Call fgVigenciaQuiebra(Txt_Desde)
   
   vlFechaInicio = Format(CDate(Trim(Txt_Desde.Text)), "yyyymmdd")
   vlFechaTermino = Format(CDate(Trim(Txt_Hasta.Text)), "yyyymmdd")
   
   vlFecIni18 = ""
   vlFecTer18 = ""
   
   vlFecIni18 = DateSerial(Mid(vlFechaInicio, 1, 4) - clEdad18, Mid(vlFechaInicio, 5, 2), Mid(vlFechaInicio, 7, 2))
   vlFecTer18 = DateSerial(Mid(vlFechaTermino, 1, 4) - clEdad18, Mid(vlFechaTermino, 5, 2), Mid(vlFechaTermino, 7, 2))
   
   vlFecIni18 = Format(CDate(Trim(vlFecIni18)), "yyyymmdd")
   vlFecTer18 = Format(CDate(Trim(vlFecTer18)), "yyyymmdd")

   vgQuery = ""
   vgQuery = vgQuery & "{PP_TMAE_BEN.fec_nacben} >= '" & Trim(vlFecIni18) & "' AND "
   vgQuery = vgQuery & "{PP_TMAE_BEN.fec_nacben} <= '" & Trim(vlFecTer18) & "' "
          
   Rpt_Reporte.Reset
   Rpt_Reporte.WindowState = crptMaximized
   Rpt_Reporte.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
   Rpt_Reporte.Connect = vgRutaDataBase       ' o App.Path & "\Nestle.mdb"
   Rpt_Reporte.SelectionFormula = vgQuery
   
   vgPalabra = ""
   vgPalabra = Txt_Desde.Text & "   *   " & Txt_Hasta.Text
   
   vlGlosaEdadBen = "Beneficiarios por Cumplir 18 Años."
   
   Rpt_Reporte.Formulas(0) = ""
   Rpt_Reporte.Formulas(1) = ""
   Rpt_Reporte.Formulas(2) = ""
   Rpt_Reporte.Formulas(3) = ""
   Rpt_Reporte.Formulas(4) = ""
   Rpt_Reporte.Formulas(5) = ""
   
   Rpt_Reporte.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
   Rpt_Reporte.Formulas(1) = "NombreSistema= '" & vgNombreSistema & "'"
   Rpt_Reporte.Formulas(2) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
   Rpt_Reporte.Formulas(3) = "Periodo = '" & vgPalabra & "'"
   Rpt_Reporte.Formulas(4) = "GlosaEdadBen = '" & vlGlosaEdadBen & "'"
   Rpt_Reporte.Formulas(5) = "GlosaQuiebra = '" & vgGlosaQuiebra & "'"
   
   Rpt_Reporte.SubreportToChange = ""
   Rpt_Reporte.Destination = crptToWindow
   Rpt_Reporte.WindowState = crptMaximized
   Rpt_Reporte.WindowTitle = "Informe Beneficiarios por Cumplir Mayoría de Edad"
'   Rpt_Reporte.SelectionFormula = ""
   Rpt_Reporte.Action = 1
   Screen.MousePointer = 0

Exit Function
Err_flInformeBen18:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function


Private Sub Cmd_Limpiar_Click()
On Error GoTo Err_Limpiar

    Txt_Desde = ""
    Txt_Hasta = ""
    Txt_Desde.SetFocus

Exit Sub
Err_Limpiar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub cmd_salir_Click()
On Error GoTo Err_Salir

    Screen.MousePointer = 11
    Unload Me
    Screen.MousePointer = 0

Exit Sub
Err_Salir:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Command1_Click()
If Trim(Txt_Desde) = "" Then
        MsgBox "Falta Ingresar Fecha Desde ", vbCritical, "Falta Información"
        Txt_Desde.SetFocus
        Exit Sub
    Else
        If Not IsDate(Txt_Desde) Then
            MsgBox "Fecha Desde no es una Fecha válida", vbCritical, "Falta Información"
            Txt_Desde.SetFocus
            Exit Sub
        End If
    End If
    If Trim(Txt_Hasta) = "" Then
        MsgBox "Falta Ingresar Fecha Hasta", vbCritical, "Falta Información"
        Txt_Hasta.SetFocus
        Exit Sub
    Else
        If Not IsDate(Txt_Hasta) Then
            MsgBox "Fecha Hasta no es una fecha Válida", vbCritical, "Falta Información"
            Txt_Hasta = ""
            Exit Sub
        End If
    End If
    
    On Error GoTo Err_flInformeCerEst
'Certificados de Supervivencia
   Screen.MousePointer = 11
   
   'marco 11/03/2010
   Dim cadena As String
   Dim objRep As New ClsReporte
   vlFechaInicio = Format(CDate(Trim(Txt_Desde.Text)), "yyyymmdd")
   vlFechaTermino = Format(CDate(Trim(Txt_Hasta.Text)), "yyyymmdd")
   vlFechaVenManu = Format(CDate(Trim(txtFecmanu.Text)), "yyyymmdd")
   
   
   vgPalabra = ""
   vgPalabra = Txt_Desde.Text & "   *   " & Txt_Hasta.Text
   
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "PK_LISTA_CERT_POR_CADUCAR_CTA.LISTAR(" & vlFechaInicio & "," & vlFechaTermino & ")", vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PP_Rpt_FormatoRenova.rpt"), ".RPT", ".TTX"), 1)
    
        
    If objRep.CargaReporte(strRpt & "", "PP_Rpt_FormatoRenova.rpt", "Formato de carta a pensionistas para renovación", rs, True, _
                            ArrFormulas("FecVenManu", vlFechaVenManu)) = False Then
                            
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
    End If
    
    'fin marco

Exit Sub
Err_flInformeCerEst:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Form_Load()
On Error GoTo Err_Carga
    
    Frm_InformePeriodo.Left = 0
    Frm_InformePeriodo.Top = 0
    op2.Value = True
Exit Sub
Err_Carga:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Txt_Desde_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    If (Txt_Desde <> "") Then
        If Not IsDate(Trim(Txt_Desde)) Then
            MsgBox "La Fecha ingresada no es una fecha válida.", vbCritical, "Error de Datos"
            Txt_Desde.SetFocus
            Exit Sub
        End If
        If Txt_Hasta <> "" Then
            If (CDate(Trim(Txt_Desde)) > CDate(Trim(Txt_Hasta))) Then
                MsgBox "La Fecha de Término de Perido es mayor a la fecha de Inicio", vbCritical, "Error de Datos"
                Exit Sub
            End If
        End If
        If (Year(CDate(Trim(Txt_Desde))) < 1900) Then
            MsgBox "Error en la Fecha ingresada es menor a la mínima fecha que se puede ingresar (1900).", vbCritical, "Error de Datos"
            Txt_Desde.SetFocus
            Exit Sub
        End If
        Txt_Desde.Text = Format(CDate(Trim(Txt_Desde)), "yyyymmdd")
        vlFechaInicio = Trim(Txt_Desde)
        Txt_Desde.Text = DateSerial(Mid((Txt_Desde.Text), 1, 4), Mid((Txt_Desde.Text), 5, 2), Mid((Txt_Desde.Text), 7, 2))
    End If
Txt_Hasta.SetFocus
End If
End Sub

Private Sub Txt_Desde_LostFocus()
If (Txt_Desde <> "") Then
    If Not IsDate(Trim(Txt_Desde)) Then
        MsgBox "La Fecha ingresada no es una fecha válida.", vbCritical, "Error de Datos"
        Txt_Desde.SetFocus
        Exit Sub
    End If
    If Txt_Hasta <> "" Then
        If (CDate(Trim(Txt_Desde)) > CDate(Trim(Txt_Hasta))) Then
            MsgBox "La Fecha de Término de Perido es mayor a la fecha de Inicio", vbCritical, "Error de Datos"
            Exit Sub
        End If
    End If
    If (Year(CDate(Trim(Txt_Desde))) < 1900) Then
        MsgBox "Error en la Fecha ingresada es menor a la mínima fecha que se puede ingresar (1900).", vbCritical, "Error de Datos"
        Txt_Desde.SetFocus
        Exit Sub
    End If
    Txt_Desde.Text = Format(CDate(Trim(Txt_Desde)), "yyyymmdd")
    vlFechaInicio = Trim(Txt_Desde)
    Txt_Desde.Text = DateSerial(Mid((Txt_Desde.Text), 1, 4), Mid((Txt_Desde.Text), 5, 2), Mid((Txt_Desde.Text), 7, 2))
End If
End Sub

Private Sub Txt_Hasta_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    If (Txt_Hasta <> "") Then
        If Not IsDate(Trim(Txt_Hasta)) Then
            MsgBox "La Fecha ingresada no es una fecha válida.", vbCritical, "Error de Datos"
            Txt_Hasta.SetFocus
            Exit Sub
        End If
        If Txt_Desde <> "" Then
            If (CDate(Trim(Txt_Desde)) > CDate(Trim(Txt_Hasta))) Then
                MsgBox "La Fecha de Término de Perido es mayor a la fecha de Inicio", vbCritical, "Error de Datos"
                Exit Sub
            End If
        End If
        If (Year(CDate(Trim(Txt_Hasta))) < 1900) Then
            MsgBox "Error en la Fecha ingresada es menor a la mínima fecha que se puede ingresar (1900).", vbCritical, "Error de Datos"
            Txt_Hasta.SetFocus
            Exit Sub
        End If
        Txt_Hasta.Text = Format(CDate(Trim(Txt_Hasta)), "yyyymmdd")
        vlFechaTermino = Trim(Txt_Hasta)
        Txt_Hasta.Text = DateSerial(Mid((Txt_Hasta.Text), 1, 4), Mid((Txt_Hasta.Text), 5, 2), Mid((Txt_Hasta.Text), 7, 2))
    End If
Cmd_Imprimir.SetFocus
End If
End Sub

Private Sub Txt_Hasta_LostFocus()
If (Txt_Hasta <> "") Then
    If Not IsDate(Trim(Txt_Hasta)) Then
        MsgBox "La Fecha ingresada no es una fecha válida.", vbCritical, "Error de Datos"
        Txt_Hasta.SetFocus
        Exit Sub
    End If
    If Txt_Desde <> "" Then
        If (CDate(Trim(Txt_Desde)) > CDate(Trim(Txt_Hasta))) Then
            MsgBox "La Fecha de Término de Perido es mayor a la fecha de Inicio", vbCritical, "Error de Datos"
            Exit Sub
        End If
    End If
    If (Year(CDate(Trim(Txt_Hasta))) < 1900) Then
        MsgBox "Error en la Fecha ingresada es menor a la mínima fecha que se puede ingresar (1900).", vbCritical, "Error de Datos"
        Txt_Hasta.SetFocus
        Exit Sub
    End If
    Txt_Hasta.Text = Format(CDate(Trim(Txt_Hasta)), "yyyymmdd")
    vlFechaTermino = Trim(Txt_Hasta)
    Txt_Hasta.Text = DateSerial(Mid((Txt_Hasta.Text), 1, 4), Mid((Txt_Hasta.Text), 5, 2), Mid((Txt_Hasta.Text), 7, 2))
End If
End Sub



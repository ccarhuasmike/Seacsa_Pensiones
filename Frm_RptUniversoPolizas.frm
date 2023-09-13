VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Frm_RptUniversoPolizas 
   Caption         =   "Univero de Pólizas"
   ClientHeight    =   2460
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6375
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   855
         Left            =   5160
         Picture         =   "Frm_RptUniversoPolizas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   855
         Left            =   4080
         Picture         =   "Frm_RptUniversoPolizas.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   960
         Width           =   855
      End
      Begin MSComctlLib.ProgressBar BarraProgreso 
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblMensaje 
         Caption         =   "lblmensale"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label lblprogreso 
         Caption         =   "Label2"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   960
         Width           =   3615
      End
   End
   Begin VB.Label lbltitulorpt 
      Alignment       =   2  'Center
      Caption         =   "Univero de Pólizas Emitidas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "Frm_RptUniversoPolizas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExportar_Click()

    Me.cmdExportar.Enabled = False
    Me.cmdSalir.Enabled = False
    
    Call CreaReporte
    
    Me.cmdExportar.Enabled = True
    Me.cmdSalir.Enabled = True

End Sub
Private Sub CreaReporte()

On Error GoTo Gestionaerror
    Dim vTotaReg As Integer
    Dim vTotaFilas As Integer
          
    Dim RS      As ADODB.Recordset, cmd As ADODB.Command
    Dim conn    As ADODB.Connection
    Set conn = New ADODB.Connection
    Set RS = New ADODB.Recordset ' CreateObject("ADODB.Recordset")
    Set objCmd = New ADODB.Command ' CreateObject("ADODB.Command")
    lblMensaje.Caption = "Obteniendo Datos, espere un momento....."
    Me.Refresh
    MousePointer = vbHourglass
        
    conn.Provider = "OraOLEDB.Oracle"
    conn.ConnectionString = "PLSQLRSet=1;Data Source=" & vgNombreBaseDatos & ";" & "User ID=" & vgNombreUsuario & ";Password=" & vgPassWord & ";"
    conn.CursorLocation = adUseClient
    conn.Open
        
    Set objCmd = CreateObject("ADODB.Command")
    Set objCmd.ActiveConnection = conn
      
    objCmd.CommandText = "PKG_Informes_Administrativos.SP_PROC_UNIVERSO_POLIZAS"
    objCmd.CommandType = adCmdStoredProc
        
    Set RS = objCmd.Execute
        
    vTotaReg = RS.RecordCount
    vTotaCampos = RS.Fields.Count
    
    MousePointer = vbDefault
   
    If vTotaReg = 0 Then
        lblMensaje.Caption = ""
        MsgBox "No se encontró información", vbDefaultButton1, "No hay datos"
        Exit Sub
    End If

    BarraProgreso.Min = 0
    BarraProgreso.Max = vTotaReg
    
    Dim Obj_Excel As Object
    Dim Obj_Libro As Object
    Dim Obj_Hoja As Object
    Dim vFila As Integer
    Dim vColumna As Integer
            
    Set Obj_Excel = CreateObject("Excel.Application")
    Set Obj_Libro = Obj_Excel.Workbooks.Add
    Set Obj_Hoja = Obj_Libro.Worksheets.Add
  
    vFila = 0
    vColumna = 0
    lblMensaje.Caption = "Escribiendo filas, espere un momento..."
    Dim Titulos(85) As String

    Titulos(0) = "IDKEY"
    Titulos(1) = "PÓLIZA"
    Titulos(2) = "ORDEN"
    Titulos(3) = "ROL"
    Titulos(4) = "CUSPP"
    Titulos(5) = "FECHA CÁLCULO"
    Titulos(6) = "FECHA CIERRE"
    Titulos(7) = "NUM COT"
    Titulos(8) = "COD AFP"
    Titulos(9) = "AFP"
    Titulos(10) = "FECHA INICIO PAGO"
    Titulos(11) = "FECHA DEV"
    Titulos(12) = "FECHA DEV SOL"
    Titulos(13) = "COD SITINV"
    Titulos(14) = "TIPO DOC"
    Titulos(15) = "DNI BEN"
    Titulos(16) = "NOMBRE BEN"
    Titulos(17) = "COD_SEXO"
    Titulos(18) = "FECHA NAC BEN"
    Titulos(19) = "FECHA FALL BEN"
    Titulos(20) = "CORREO BEN"
    Titulos(21) = "TELEFONO BEN"
    Titulos(22) = "DIRECCIÓN BEN"
    Titulos(23) = "COMUNA"
    Titulos(24) = "PROVINCIA"
    Titulos(25) = "REGIÓN"
    Titulos(26) = "TIPO PEN"
    Titulos(27) = "TIPO JUB"
    Titulos(28) = "RENTA"
    Titulos(29) = "MODALIDAD"
    Titulos(30) = "NUM MES DIF"
    Titulos(31) = "NUM MES GAR"
    Titulos(32) = "COD DERCRE"
    Titulos(33) = "GRATIFICACIÓN"
    Titulos(34) = "TIPO COBRANTE"
    Titulos(35) = "IDEN COBRANTE"
    Titulos(36) = "NOMBRE COBRANTE"
    Titulos(37) = "DIRECCIÓN COBRANTE"
    Titulos(38) = "DPTO COBRANTE"
    Titulos(39) = "PROVINCIA COBRANTE"
    Titulos(40) = "DISTRITO COBRANTE"
    Titulos(41) = "TELÉFONO COBRANTE"
    Titulos(42) = "CORRE COBRANTE"
    Titulos(43) = "FORMA PAGO"
    Titulos(44) = "BANCO"
    Titulos(45) = "TIPO CUENTA"
    Titulos(46) = "NUM CUENTA"
    Titulos(47) = "NUM CUENTA CCI"
    Titulos(48) = "PRIINF"
    Titulos(49) = "FECHA SOLICITUD"
    Titulos(50) = "CASA VTA"
    Titulos(51) = "FECHA ACEPTA"
    Titulos(52) = "TCTrasAFP"
    Titulos(53) = "PRIRECAL"
    Titulos(54) = "MTO PRICIA"
    Titulos(55) = "PENSIÓN EMITIDA"
    Titulos(56) = "FECHA EMISIÓN"
    Titulos(57) = "FECHA TRASPASO"
    Titulos(58) = "NUM IDEN COR"
    Titulos(59) = "NOMBRE ASESOR"
    Titulos(60) = "CORCOMREAL"
    Titulos(61) = "IND COB"
    Titulos(62) = "IPC"
    Titulos(63) = "TASAMER"
    Titulos(64) = "PRC PENSIÓN"
    Titulos(65) = "PENSIÓN"
    Titulos(66) = "PRC PENSIÓN GAR"
    Titulos(67) = "MTO PENSIÓN GAR"
    Titulos(68) = "PRC TASACE"
    Titulos(69) = "COD COBERCON"
    Titulos(70) = "RMB"
    Titulos(71) = "RMF"
    Titulos(72) = "PRC TASATIR"
    Titulos(73) = "PRC PERCON"
    Titulos(74) = "RETENCIÓN"
    Titulos(75) = "INI_CERT_SOB"
    Titulos(76) = "FIN_CERT_SOB"
    Titulos(77) = "INI_CERT_EST"
    Titulos(78) = "FIN_CERT_EST"
    Titulos(79) = "CHECK_CER_EST"
    Titulos(80) = "MONEDA ORI"
    Titulos(81) = "MONEDA"
    Titulos(82) = "ESTADO DERECHO EMI"
    Titulos(83) = "SITUACION"
    Titulos(84) = "EXONERADO"

    vFila = 1
    Obj_Hoja.Cells(vFila, 1) = "Univero de Pólizas Emitidas"
    Obj_Hoja.Rows(vFila).Font.Bold = True
    
    Obj_Hoja.Columns("A").NumberFormat = "@"
    Obj_Hoja.Columns("B").NumberFormat = "@"
    Obj_Hoja.Columns("E").NumberFormat = "@"
    Obj_Hoja.Columns("H").NumberFormat = "@"
    Obj_Hoja.Columns("I").NumberFormat = "@"
    Obj_Hoja.Columns("P").NumberFormat = "@"
    Obj_Hoja.Columns("U").NumberFormat = "@"
    Obj_Hoja.Columns("V").NumberFormat = "@"
    Obj_Hoja.Columns("AJ").NumberFormat = "@"
    Obj_Hoja.Columns("AP").NumberFormat = "@"
    Obj_Hoja.Columns("AQ").NumberFormat = "@"
    Obj_Hoja.Columns("AU").NumberFormat = "@"
    Obj_Hoja.Columns("AV").NumberFormat = "@"
    Obj_Hoja.Columns("AW").NumberFormat = "######,##0.00"
    Obj_Hoja.Columns("AY").NumberFormat = "######,##0.00"
    Obj_Hoja.Columns("BA:BD").NumberFormat = "######,##0.00"
    Obj_Hoja.Columns("BG").NumberFormat = "@"
    Obj_Hoja.Columns("BI").NumberFormat = "######,##0.00"
    Obj_Hoja.Columns("BJ").NumberFormat = "@"
    Obj_Hoja.Columns("BL:BQ").NumberFormat = "######,##0.00"
    Obj_Hoja.Columns("BU:BW").NumberFormat = "######,##0.00"
      
    vFila = 3
    
    For i = 0 To UBound(Titulos)
    
          Obj_Hoja.Cells(vFila, i + 1) = Titulos(i)
    
    Next
    
          
    Do While Not RS.EOF
        vFila = vFila + 1
        For vColumna = 0 To vTotaCampos - 1
            Obj_Hoja.Cells(vFila, vColumna + 1) = RS.Fields(vColumna).Value
        Next
    
        BarraProgreso.Value = vFila - 3
        lblprogreso.Caption = "Procesando " & BarraProgreso.Value & " de " & vTotaReg & " Registros."
        Me.Refresh
        RS.MoveNext
    
    Loop
      
    RS.Close
    Set RS = Nothing
    conn.Close
    
    Obj_Hoja.Columns("A:CG").AutoFit
    Obj_Hoja.Rows(2).Font.Bold = True
    Obj_Hoja.Columns("A").ColumnWidth = 18
    Obj_Hoja.Range("A3:CG3").Borders.LineStyle = xlContinuous
    Obj_Hoja.Range("A3:CG3").Interior.Color = RGB(0, 32, 96)
    Obj_Hoja.Range("A3:CG3").Font.Color = RGB(237, 125, 49)
    Obj_Hoja.Range("A3:CG3").RowHeight = 34.5
    Obj_Hoja.Range("A3:CG3").VerticalAlignment = xlVAlignCenter
    Obj_Hoja.Range("A1:A1").Select
    
    Obj_Excel.Visible = True
    
    Set Obj_Hoja = Nothing
    Set Obj_Libro = Nothing
    Set Obj_Excel = Nothing
    
    MousePointer = vbDefault
    
    BarraProgreso.Value = 0
    lblMensaje.Caption = ""
    lblprogreso.Caption = ""
    
Gestionaerror:
    If Err.Number <> 0 Then
        GestiónError
        Resume Next
    End If
End Sub

Private Sub GestiónError()
    MsgBox ("Se ha producido un error. Tipo de error = " & Err.Number & ". Descripción: " & Err.Description)
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.lblMensaje.Caption = ""
    Me.lblprogreso.Caption = ""
End Sub

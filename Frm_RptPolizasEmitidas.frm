VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_RptPoliasEmitidas 
   Caption         =   "Pólizas Emitidas"
   ClientHeight    =   3195
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6675
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   6675
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   855
      Left            =   5160
      Picture         =   "Frm_RptPolizasEmitidas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1920
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6375
      Begin MSComCtl2.DTPicker dtFechaFin 
         Height          =   375
         Left            =   4560
         TabIndex        =   11
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   83755009
         CurrentDate     =   44473
      End
      Begin MSComCtl2.DTPicker dtFechaInicio 
         Height          =   375
         Left            =   1320
         TabIndex        =   10
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   83755009
         CurrentDate     =   44473
      End
      Begin MSComDlg.CommonDialog VentSelecFichero 
         Left            =   2400
         Top             =   1560
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Ubicacion del archivo"
      End
      Begin VB.CommandButton cmdtxt 
         Caption         =   "TXT"
         Height          =   855
         Left            =   1080
         Picture         =   "Frm_RptPolizasEmitidas.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSComctlLib.ProgressBar BarraProgreso 
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1200
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   855
         Left            =   4080
         Picture         =   "Frm_RptPolizasEmitidas.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Inicio:"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblprogreso 
         Caption         =   "Label2"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   1560
         Width           =   3615
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
         TabIndex        =   3
         Top             =   840
         Width           =   4335
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Fin:"
         Height          =   375
         Left            =   3480
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Label lbltitulorpt 
      Alignment       =   2  'Center
      Caption         =   "Pólizas Emitidas"
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
      TabIndex        =   7
      Top             =   0
      Width           =   6135
   End
   Begin VB.Label lblTipo_rpt 
      Height          =   375
      Left            =   5760
      TabIndex        =   6
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "Frm_RptPoliasEmitidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const adReadAll = -1
Private Const adSaveCreateOverWrite = 2
Private Const adTypeBinary = 1
Private Const adTypeText = 2
Private Const adWriteChar = 0
Private Sub cmdExportar_Click()
    
    Me.cmdExportar.Enabled = False
    Me.cmdSalir.Enabled = False

    dtFechaInicio.Enabled = False
    dtFechaFin.Enabled = False
 
    
    Call CreaReporte
  
    
    Me.cmdExportar.Enabled = True
    Me.cmdSalir.Enabled = True
    dtFechaInicio.Enabled = True
    dtFechaFin.Enabled = True
    
     

End Sub
Private Sub CreaReporte()

On Error GoTo Gestionaerror


   Dim vTotaReg As Integer
        Dim vTotaFilas As Integer
        
'          Me.lblinicio.Caption = Time
'          Me.Refresh
          
          
        Dim rs      As ADODB.Recordset, cmd As ADODB.Command
        Dim conn    As ADODB.Connection
        Set conn = New ADODB.Connection
        Set rs = New ADODB.Recordset ' CreateObject("ADODB.Recordset")
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
      
        objCmd.CommandText = "PKG_Informes_Administrativos.SP_PROC_POLIZAS_EMITIDAS"
        objCmd.CommandType = adCmdStoredProc
        
        Set param1 = objCmd.CreateParameter("p_fec_ini", adVarChar, adParamInput, 8, Format(dtFechaInicio.Value, "yyyyMMdd"))
        Set param2 = objCmd.CreateParameter("p_fec_fin", adVarChar, adParamInput, 8, Format(dtFechaFin.Value, "yyyyMMdd"))
        
        objCmd.Parameters.Append param1
        objCmd.Parameters.Append param2
        
        Set rs = objCmd.Execute
'        rs.MoveLast
'        rs.MoveFirst
        
        vTotaReg = rs.RecordCount
        vTotaCampos = rs.Fields.Count
    
        MousePointer = vbDefault
       
   
        If vTotaReg = 0 Then
              lblMensaje.Caption = ""
              MsgBox "No se encontró información para el periodo indicado", vbDefaultButton1, "No hay datos"
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
 
    Dim Titulos(41) As String

        Titulos(0) = "NUM_POLIZA"
        Titulos(1) = "NUM_ORDEN"
        Titulos(2) = "COD_PARENTESCO"
        Titulos(3) = "DESC_PARENTESCO"
        Titulos(4) = "DNI"
        Titulos(5) = "NOMBEN"
        Titulos(6) = "NOMSEGBEN"
        Titulos(7) = "PATBEN"
        Titulos(8) = "MATBEN"
        Titulos(9) = "COD_SEXO"
        Titulos(10) = "FEC_FALLBEN"
        Titulos(11) = "NOMBRES_REP"
        Titulos(12) = "COD_TIPOIDENREP,"
        Titulos(13) = "NUM_IDENREP"
        Titulos(14) = "COD_SEXOREP"
        Titulos(15) = "GLS_FONOREP"
        Titulos(16) = "GLS_FONO2REP"
        Titulos(17) = "NOM_COBRANTE"
        Titulos(18) = "DIRECCION_COBRANTE"
        Titulos(19) = "GLS_COMUNA"
        Titulos(20) = "GLS_PROVINCIA"
        Titulos(21) = "GLS_REGION"
        Titulos(22) = "CORREOBEN"
        Titulos(23) = "TELEFONO"
        Titulos(24) = "COD_AFP"
        Titulos(25) = "DESC_AFP"
        Titulos(26) = "FEC_INI_RRVV"
        Titulos(27) = "TIP_PENSION_ORIGEN"
        Titulos(28) = "TIP_PENSION_ACTUAL"
        Titulos(29) = "ENTIDAD_PAGA"
        Titulos(30) = "MES_INICIO"
        Titulos(31) = "COB_ACTUAL"
        Titulos(32) = "COD_CUSPP"
        Titulos(33) = "DNI_ASEGURADO"
        Titulos(34) = "NOMBRE_ASEGURADO"
        Titulos(35) = "PER_GARANTIZADO"
        Titulos(36) = "NORMA_JUBILACION"
        Titulos(37) = "FEC_EMISION"
        Titulos(38) = "MTO_PRIMA"
        Titulos(39) = "PENSION_BASE"
        Titulos(40) = "MONEDA"

    
    vFila = 1
    Obj_Hoja.Cells(vFila, 1) = "Pólizas Emitidas"
    Obj_Hoja.rows(vFila).Font.Bold = True
    vFila = 2
    Obj_Hoja.Cells(vFila, 1) = "Rango de fechas: del " & dtFechaInicio.Value & " al " & Me.dtFechaFin.Value
    Obj_Hoja.rows(vFila).Font.Bold = True
    Obj_Hoja.Columns("A").NumberFormat = "@"
    Obj_Hoja.Columns("E").NumberFormat = "@"
    Obj_Hoja.Columns("R").NumberFormat = "@"
    Obj_Hoja.Columns("K").NumberFormat = "@"
    Obj_Hoja.Columns("N").NumberFormat = "@"
    Obj_Hoja.Columns("X").NumberFormat = "@"
    Obj_Hoja.Columns("AA").NumberFormat = "@"
    Obj_Hoja.Columns("AH").NumberFormat = "@"
    Obj_Hoja.Columns("AM:AN").NumberFormat = "######,##0.00"
  
   
    vFila = 3
    
    For i = 0 To UBound(Titulos)
    
          Obj_Hoja.Cells(vFila, i + 1) = Titulos(i)
    
    Next
    
          
      Do While Not rs.EOF
               vFila = vFila + 1
              For vColumna = 0 To vTotaCampos - 1
                Obj_Hoja.Cells(vFila, vColumna + 1) = rs.Fields(vColumna).Value
                
              Next
    
            BarraProgreso.Value = vFila - 3
            lblprogreso.Caption = "Procesando " & BarraProgreso.Value & " de " & vTotaReg & " Registros."
            Me.Refresh
            
        rs.MoveNext
        
      Loop
      
     rs.Close
     Set rs = Nothing
     conn.Close
     
     Obj_Hoja.Columns("A:AO").AutoFit
     Obj_Hoja.rows(3).Font.Bold = True
     Obj_Hoja.Columns("A").ColumnWidth = 18
     Obj_Hoja.Range("A3:AO3").Borders.LineStyle = xlContinuous
     Obj_Hoja.Range("A3:AO3").Interior.Color = RGB(0, 32, 96)
     Obj_Hoja.Range("A3:AO3").Font.Color = RGB(237, 125, 49)
     Obj_Hoja.Range("A3:AO3").RowHeight = 34.5
     Obj_Hoja.Range("A3:AO3").VerticalAlignment = xlVAlignCenter
    
     Obj_Hoja.Range("A1:A1").Select
 
     Obj_Excel.Visible = True
      
'       Me.lblfinexcel.Caption = Time
'       Me.lblfinexcel.Caption = Format((CDate(lblinicio.Caption) - CDate(lblfinexcel.Caption)), "hh:mm:ss")
         
    
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




Private Sub cmdsalir_Click()
Unload Me
 
End Sub







Private Sub Form_Load()


Dim AnnioActual As Integer
Dim MesActual As Integer
Dim ItemAnioMes As String
Dim ItemInicio As String




lblMensaje.Caption = ""
lblprogreso.Caption = ""

dtFechaInicio.Value = Date
dtFechaFin.Value = Date

End Sub
Private Sub GestiónError()
    MsgBox ("Se ha producido un error. Tipo de error = " & Err.Number & ". Descripción: " & Err.Description)
End Sub














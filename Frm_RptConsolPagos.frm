VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_RptConsolPagos 
   Caption         =   "Consolidados Pagos"
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
      Picture         =   "Frm_RptConsolPagos.frx":0000
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
         TabIndex        =   12
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   85393409
         CurrentDate     =   44473
      End
      Begin MSComCtl2.DTPicker dtFechaInicio 
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   85393409
         CurrentDate     =   44473
      End
      Begin MSComDlg.CommonDialog VentSelecFichero 
         Left            =   2040
         Top             =   1800
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Ubicacion del archivo"
      End
      Begin VB.CommandButton cmdtxt 
         Caption         =   "TXT"
         Height          =   855
         Left            =   3120
         Picture         =   "Frm_RptConsolPagos.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1560
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
         Picture         =   "Frm_RptConsolPagos.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Fin:"
         Height          =   375
         Left            =   3480
         TabIndex        =   10
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
         Caption         =   "Fecha Inicio:"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Label lbltitulorpt 
      Alignment       =   2  'Center
      Caption         =   "Consolidado de Pagos"
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
Attribute VB_Name = "Frm_RptConsolPagos"
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
      
        objCmd.CommandText = "PKG_Informes_Administrativos.SP_PROC_CONSOLIDADO_PAGOS"
        objCmd.CommandType = adCmdStoredProc
        
        Set param1 = objCmd.CreateParameter("p_fec_ini", adVarChar, adParamInput, 8, Format(dtFechaInicio.Value, "yyyyMMdd"))
        Set param2 = objCmd.CreateParameter("p_fec_fin", adVarChar, adParamInput, 8, Format(dtFechaFin.Value, "yyyyMMdd"))
        
        objCmd.Parameters.Append param1
        objCmd.Parameters.Append param2
        
        Set RS = objCmd.Execute
'        rs.MoveLast
'        rs.MoveFirst
        
        vTotaReg = RS.RecordCount
        vTotaCampos = RS.Fields.Count
    
        MousePointer = vbDefault
       
   
        If vTotaReg = 0 Then
              lblMensaje.Caption = ""
              MsgBox "No se encontró información para el rango de fechas indicado", vbDefaultButton1, "No hay datos"
          Exit Sub
          
        End If
'         Me.lblfin.Caption = Time
'         Me.lblfin.Caption = Format((CDate(lblinicio.Caption) - CDate(lblfin.Caption)), "hh:mm:ss")
'

         
'        Me.lblinicio.Caption = Time
         
        
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
 
    Dim Titulos(29) As String
     
    Titulos(0) = "NUM_POLIZA"
    Titulos(1) = "COD_CUSPP"
    Titulos(2) = "AFP"
    Titulos(3) = "PRESTACION"
    Titulos(4) = "PRESTACION_ACTUAL"
    Titulos(5) = "TIPO_SOLICITUD"
    Titulos(6) = "NOMBRE_TIT"
    Titulos(7) = "TIPO_BEN"
    Titulos(8) = "NOMBRE_BEN"
    Titulos(9) = "COD_TIPOIDENBEN"
    Titulos(10) = "NUM_IDENBEN"
    Titulos(11) = "NOMBRE_COB"
    Titulos(12) = "TIP_DOC_COB"
    Titulos(13) = "NUM_DOC_COB"
    Titulos(14) = "MONEDA"
    Titulos(15) = "TIPO_PAGO"
    Titulos(16) = "TIPO_PAGO_DET"
    Titulos(17) = "MTO_BRUTO"
    Titulos(18) = "MTO_ESS"
    Titulos(19) = "MTO_RETJUD"
    Titulos(20) = "MTO_LIQPAGAR"
    Titulos(21) = "NUM_PERPAGO"
    Titulos(22) = "VIA_PAGO"
    Titulos(23) = "BANCO"
    Titulos(24) = "TIP_CUENTA"
    Titulos(25) = "NUM_CUENTA"
    Titulos(26) = "FEC_PAGO"
    Titulos(27) = "NUM_MEMO"
    Titulos(28) = "TIPO"

    vFila = 1
    Obj_Hoja.Cells(vFila, 1) = "Consolidado de Pagos"
    Obj_Hoja.rows(vFila).Font.Bold = True
    vFila = 2
    Obj_Hoja.Cells(vFila, 1) = "Rango de fechas: del " & dtFechaInicio.Value & " al " & Me.dtFechaFin.Value
    Obj_Hoja.rows(vFila).Font.Bold = True
    Obj_Hoja.Columns("A").NumberFormat = "@"
    Obj_Hoja.Columns("N").NumberFormat = "@"
    Obj_Hoja.Columns("Z").NumberFormat = "@"
    Obj_Hoja.Columns("AA").NumberFormat = "@"
    Obj_Hoja.Columns("AB").NumberFormat = "@"
   
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
     
     Obj_Hoja.Columns("A:AC").AutoFit
     Obj_Hoja.rows(3).Font.Bold = True
     Obj_Hoja.Range("A3:AC3").Borders.LineStyle = xlContinuous
     Obj_Hoja.Range("A3:AC3").Interior.Color = RGB(0, 32, 96)
     Obj_Hoja.Range("A3:AC3").Font.Color = RGB(237, 125, 49)
     Obj_Hoja.Range("A3:AC3").RowHeight = 34.5
     Obj_Hoja.Range("A3:AC3").VerticalAlignment = xlVAlignCenter
     Obj_Hoja.Columns("A").ColumnWidth = 18
     Obj_Hoja.Columns("R:U").NumberFormat = "######,##0.00"

     
'      Obj_Hoja.Columns("B").NumberFormat = "@"
'    Obj_Hoja.Columns("I").NumberFormat = "@"
'    Obj_Hoja.Columns("P").NumberFormat = "@"
'    Obj_Hoja.Columns("Y:AC").NumberFormat = "######,##0.00"
'    Obj_Hoja.Columns("X").NumberFormat = "dd/mm/yyyy"
'    Obj_Hoja.Columns("AD").NumberFormat = "@"
'    Obj_Hoja.Columns("AD").HorizontalAlignment = xlHAlignRight
'    Obj_Hoja.Columns("AH").NumberFormat = "@"
'    Obj_Hoja.Columns("AH").HorizontalAlignment = xlHAlignRight
'     Obj_Hoja.Columns("O").NumberFormat = "@"
'        Obj_Hoja.Columns("S").NumberFormat = "@"
'        Obj_Hoja.Columns("F").NumberFormat = "@"
'        Obj_Hoja.Columns("S").HorizontalAlignment = xlHAlignRight
'        Obj_Hoja.Columns("E").HorizontalAlignment = xlHAlignRight
'        Obj_Hoja.Columns("F").HorizontalAlignment = xlHAlignRight
'        Obj_Hoja.Columns("AM").HorizontalAlignment = xlHAlignRight
'        Obj_Hoja.Columns("W:X").NumberFormat = "######,##0.00"
'        Obj_Hoja.Columns("Z:AH").NumberFormat = "######,##0.00"
'        Obj_Hoja.Columns("AJ").NumberFormat = "######,##0.00"
 
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

Private Sub cmdSalir_Click()
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













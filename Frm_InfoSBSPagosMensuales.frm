VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Frm_InfoSBSPagosMensuales 
   Caption         =   "Informe SBS - Pagos Mensuales"
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
      Picture         =   "Frm_InfoSBSPagosMensuales.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1920
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6375
      Begin MSComDlg.CommonDialog VentSelecFichero 
         Left            =   5400
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Ubicacion del archivo"
      End
      Begin VB.CommandButton cmdtxt 
         Caption         =   "TXT"
         Height          =   855
         Left            =   3960
         Picture         =   "Frm_InfoSBSPagosMensuales.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1560
         Width           =   855
      End
      Begin MSComctlLib.ProgressBar BarraProgreso 
         Height          =   255
         Left            =   360
         TabIndex        =   5
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
         Left            =   3000
         Picture         =   "Frm_InfoSBSPagosMensuales.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1560
         Width           =   855
      End
      Begin VB.ComboBox CmbMesExtrae 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblprogreso 
         Caption         =   "Label2"
         Height          =   255
         Left            =   360
         TabIndex        =   6
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
         TabIndex        =   4
         Top             =   840
         Width           =   4335
      End
      Begin VB.Label Label1 
         Caption         =   "Periodo:"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Label lbltitulorpt 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      TabIndex        =   8
      Top             =   0
      Width           =   6135
   End
   Begin VB.Label lblTipo_rpt 
      Height          =   375
      Left            =   5760
      TabIndex        =   7
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "Frm_InfoSBSPagosMensuales"
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
    Me.CmbMesExtrae.Enabled = False
    Me.cmdtxt.Enabled = False
    
    
        Select Case Me.lblTipo_rpt.Caption
        Case 1
            Call Trama_AFP
        Case 2
            Call Oficio_multiple
        Case 3
            Call Formato0730
        End Select
  
    
    Me.cmdExportar.Enabled = True
    Me.cmdSalir.Enabled = True
    Me.CmbMesExtrae.Enabled = True
    Me.cmdtxt.Enabled = True
    
     

End Sub
Private Sub Oficio_multiple()

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
      
        objCmd.CommandText = "PKG_SBSINFORMES.oficio_multiple"
        'objCmd.CommandText = "PKG_INFORMESBS_TRAMAS.oficio_multiple"
        objCmd.CommandType = adCmdStoredProc
        
        Set param1 = objCmd.CreateParameter("p_periodo", adVarChar, adParamInput, 6, CmbMesExtrae.Text)
        objCmd.Parameters.Append param1
        
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
 
    Dim Titulos(30) As String

 Titulos(0) = "AFP"
 Titulos(1) = "AÑO"
 Titulos(2) = "MES"
 Titulos(3) = "CUSPP"
 Titulos(4) = "TIPO DE BENEFICIO"
 Titulos(5) = "TIPO DE SOLICITUD"
 Titulos(6) = "ETAPA DE PAGO"
 Titulos(7) = "MODALIDAD"
 Titulos(8) = "PERIODO DE PAGO"
 Titulos(9) = "ORIGEN DEL FONDO"
 Titulos(10) = "APELLIDOS Y NOMBRES DEL AFILIADO"
 Titulos(11) = "APELLIDOS Y NOMBRES DEL BENEFICIARIO"
 Titulos(12) = "PARENTESCO"
 Titulos(13) = "MONEDA"
 Titulos(14) = "MONTO BRUTO TOTAL"
 Titulos(15) = "PENSION BRUTA MENSUAL"
 Titulos(16) = "N° DE MESES DEVENGADOS"
 Titulos(17) = "REINTEGRO 20%"
 Titulos(18) = "MONTO DE OTRO ABONO"
 Titulos(19) = "CONCEPTO DEL OTRO ABONO"
 Titulos(20) = "ESSALUD"
 Titulos(21) = "APORTE A CIC"
 Titulos(22) = "APORTE DE SEGURO PREVISIONAL"
 Titulos(23) = "COMISION AFP"
 Titulos(24) = "RETENCION JUDICIAL"
 Titulos(25) = "MONTO DE OTRO DESCUENTO"
 Titulos(26) = "CONCEPTO DEL OTRO DESCUENTO"
 Titulos(27) = "PENSION NETA MENSUAL"
 Titulos(28) = "FORMA DE PAGO"
 Titulos(29) = "BANCO"
 Titulos(30) = "AGENCIA"

    
    
    vFila = 1
    Obj_Hoja.Cells(vFila, 1) = "Informe SBS - Oficio Multiple 24729"
    Obj_Hoja.rows(vFila).Font.Bold = True
    vFila = 2
    Obj_Hoja.Cells(vFila, 1) = "Periodo: " & CmbMesExtrae.Text
    Obj_Hoja.rows(vFila).Font.Bold = True
   
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
     
     Obj_Hoja.Columns("A:AE").AutoFit
     Obj_Hoja.rows(3).Font.Bold = True
     Obj_Hoja.Range("A3:AE3").Borders.LineStyle = xlContinuous
     Obj_Hoja.Range("A3:AE3").Interior.Color = RGB(0, 32, 96)
     Obj_Hoja.Range("A3:AE3").Font.Color = RGB(237, 125, 49)
     Obj_Hoja.Range("A3:AE3").RowHeight = 34.5
     Obj_Hoja.Range("A3:AE3").VerticalAlignment = xlVAlignCenter
     
     Obj_Hoja.Columns("C").NumberFormat = "00"
     Obj_Hoja.Columns("O:P").NumberFormat = "######,##0.00"
     Obj_Hoja.Columns("U").NumberFormat = "######,##0.00"
     Obj_Hoja.Columns("AB").NumberFormat = "######,##0.00"
     
     
  
     
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
Private Sub Trama_AFP()

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
         objCmd.CommandText = "PKG_SBSINFORMES.pago_mensual"
       ' objCmd.CommandText = "PKG_INFORMESBS_TRAMAS.pago_mensual"
        objCmd.CommandType = adCmdStoredProc
        
        Set param1 = objCmd.CreateParameter("p_periodo", adVarChar, adParamInput, 6, CmbMesExtrae.Text)
        objCmd.Parameters.Append param1
        
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
      
      
      
      
      
    Dim Titulos(39) As String

        Titulos(0) = "CIA"
        Titulos(1) = "NUM_POLIZA"
        Titulos(2) = "AFP"
        Titulos(3) = "COD_CUSPP"
        Titulos(4) = "NOMBRE_TIT"
        Titulos(5) = "TIPO_BEN"
        Titulos(6) = "NOMBRE_BEN"
        Titulos(7) = "TIPO_DOC"
        Titulos(8) = "NUM_IDENBEN"
        Titulos(9) = "FEC_NACBEN"
        Titulos(10) = "COD_SEXO"
        Titulos(11) = "CONTINUIDAD_ESTUDIOS"
        Titulos(12) = "INV_TOTAL"
        Titulos(13) = "NOMBRE_COB"
        Titulos(14) = "TIPO_DOC_COB"
        Titulos(15) = "NUM_DOC_COB"
        Titulos(16) = "PRESTACION"
        Titulos(17) = "MODALIDAD"
        Titulos(18) = "FEC_FINPERGAR"
        Titulos(19) = "MON_PAGO"
        Titulos(20) = "MON_CARACTERISTICA"
        Titulos(21) = "GRATIFICACION"
        Titulos(22) = "TIPO_PAGO"
        Titulos(23) = "DEVENGUE"
        Titulos(24) = "MTO_BRUTO"
        Titulos(25) = "MTO_ESS"
        Titulos(26) = "MTO_RETJUD"
        Titulos(27) = "MTO_OTROS"
        Titulos(28) = "MTO_LIQPAGAR"
        Titulos(29) = "NUM_PERPAGO"
        Titulos(30) = "MODALIDAD DE PAGO"
        Titulos(31) = "BANCO"
        Titulos(32) = "TIPO_CUENTA"
        Titulos(33) = "NUMERO_CUENTA"
        Titulos(34) = "AGENCIA"
        Titulos(35) = "OBSERVACION"
        Titulos(36) = "TIPO"
        Titulos(37) = "PRESTACION"
        Titulos(38) = "TIPO_JUBILACION"
    
    
    vFila = 1
    Obj_Hoja.Cells(vFila, 1) = "Informe SBS - Pago Mensual"
    Obj_Hoja.rows(vFila).Font.Bold = True
    vFila = 2
    Obj_Hoja.Cells(vFila, 1) = "Periodo: " & CmbMesExtrae.Text
    Obj_Hoja.rows(vFila).Font.Bold = True
    
    Obj_Hoja.Columns("B").NumberFormat = "@"
    Obj_Hoja.Columns("I").NumberFormat = "@"
    Obj_Hoja.Columns("P").NumberFormat = "@"
    Obj_Hoja.Columns("Y:AC").NumberFormat = "######,##0.00"
    Obj_Hoja.Columns("X").NumberFormat = "dd/mm/yyyy"
    Obj_Hoja.Columns("AD").NumberFormat = "@"
    Obj_Hoja.Columns("AD").HorizontalAlignment = xlHAlignRight
    Obj_Hoja.Columns("AH").NumberFormat = "@"
    Obj_Hoja.Columns("AH").HorizontalAlignment = xlHAlignRight
   
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
     
     Obj_Hoja.Columns("A:AM").AutoFit
     Obj_Hoja.rows(3).Font.Bold = True
     Obj_Hoja.Range("A3:AM3").Borders.LineStyle = xlContinuous
     Obj_Hoja.Range("A3:AM3").Interior.Color = RGB(0, 32, 96)
     Obj_Hoja.Range("A3:AM3").Font.Color = RGB(237, 125, 49)
     Obj_Hoja.Range("A1:A1").Select
 
      Obj_Excel.Visible = True
      
'       Me.lblfinexcel.Caption = Time
'       Me.lblfinexcel.Caption = Format((CDate(lblinicio.Caption) - CDate(lblfinexcel.Caption)), "hh:mm:ss")
'
    
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

Private Sub cmdtxt_Click()

    Me.cmdExportar.Enabled = False
    Me.cmdSalir.Enabled = False
    Me.CmbMesExtrae.Enabled = False
    Me.cmdtxt.Enabled = False
    
    Call GeneraArchivo

    Me.cmdExportar.Enabled = True
    Me.cmdSalir.Enabled = True
    Me.CmbMesExtrae.Enabled = True
    Me.cmdtxt.Enabled = True
    




End Sub

Private Sub Form_Activate()



 Select Case Me.lblTipo_rpt.Caption
        Case 1
            Me.lbltitulorpt.Caption = "Trama AFP"
            Me.Caption = "Trama AFP"
            cmdtxt.Enabled = False
        Case 2
      Me.lbltitulorpt.Caption = "Oficio Multiple 24729"
            Me.Caption = "Oficio Multiple 24729"
            
        Case 3
      
            
              
              Me.lbltitulorpt.Caption = "Informe 0730"
            Me.Caption = "Informe 0730"
            cmdtxt.Enabled = True
        End Select
    

End Sub

Private Sub Form_Load()


Dim AnnioActual As Integer
Dim MesActual As Integer
Dim ItemAnioMes As String
Dim ItemInicio As String




lblMensaje.Caption = ""
lblprogreso.Caption = ""




ItemInicio = "201801"

AnnioActual = Format(Now, "YYYY")
MesActual = Format(Now, "MM")

ItemAnioMes = AnnioActual & Right(String(2, "0") & MesActual, 2)



Dim i As Long

For i = ItemInicio To ItemAnioMes

CmbMesExtrae.AddItem (i)

If Right(i, 2) = 12 Then
    i = i + 88
End If

Next

CmbMesExtrae.Text = i - 1

End Sub
Private Sub GestiónError()
    MsgBox ("Se ha producido un error. Tipo de error = " & Err.Number & ". Descripción: " & Err.Description)
End Sub
Private Sub Formato0730()

On Error GoTo Gestionaerror


   Dim vTotaReg As Integer
        Dim vTotaFilas As Integer
          
          
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
      
        objCmd.CommandText = "PKG_SBSINFORMES.formato0730_SBS"
        objCmd.CommandType = adCmdStoredProc
        
        Set param1 = objCmd.CreateParameter("p_periodo", adVarChar, adParamInput, 6, CmbMesExtrae.Text)
        objCmd.Parameters.Append param1
        
        Set rs = objCmd.Execute

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
 
    Dim Titulos(40) As String

 Titulos(0) = "FILA"
 Titulos(1) = "CUSPP"
 Titulos(2) = "TIPO BENEFICIO"
 Titulos(3) = "TIPO SOLICITUD"
 Titulos(4) = "FECHA SOLICITUD"
 Titulos(5) = "NUMERO SOLICITUD"
 Titulos(6) = "ETAPA PAGO"
 Titulos(7) = "MODALIDAD"
 Titulos(8) = "PERIODO PAGO"
 Titulos(9) = "ORIGEN FONDO"
 Titulos(10) = "APELLIDO PATERNO AFILIADO"
 Titulos(11) = "APELLIDO MATERNO AFILIADO"
 Titulos(12) = "NOMBRES AFILIADO"
 Titulos(13) = "TIPO DOCUMENTO BENEFICIARIO"
 Titulos(14) = "NUMERO DOCUMENTO BENEFICIARIO"
 Titulos(15) = "APELLIDO PATERNO BENEFICIARIO"
 Titulos(16) = "APELLIDO MATERNO BENEFICIARIO"
 Titulos(17) = "NOMBRES BENEFICIARIO"
 Titulos(18) = "FECHA NACIMIENTO BENEFICIARIO"
 Titulos(19) = "PARENTESCO"
 Titulos(20) = "SEXO"
 Titulos(21) = "MONEDA"
 Titulos(22) = "MONTO BRUTO TOTAL"
 Titulos(23) = "PENSION BRUTA MENSUAL"
 Titulos(24) = "MESES DEVENGADOS"
 Titulos(25) = "REINTEGRO 20%"
 Titulos(26) = "MONTO OTRO ABONO"
 Titulos(27) = "CONCEPTO OTRO ABONO"
 Titulos(28) = "ESSALUD"
 Titulos(29) = "APORTE CIC"
 Titulos(30) = "APORTE SEGURO PREVISIONAL"
 Titulos(31) = "COMISION AFP"
 Titulos(32) = "RETENCION JUDICIAL"
 Titulos(33) = "OTRO DESCUENTO"
 Titulos(34) = "CONCEPTO OTRO DESCUENTO"
 Titulos(35) = "PENSION NETA MENSUAL"
 Titulos(36) = "FORMA PAGO"
 Titulos(37) = "BANCO"
 Titulos(38) = "AGENCIA"
 Titulos(39) = "NOMBRE OTRO BANCO"
 Titulos(40) = "SITUACION COBERTURA"
 
    
    vFila = 1
    Obj_Hoja.Cells(vFila, 1) = "Informe SBS - FORMATO 0730"
    Obj_Hoja.rows(vFila).Font.Bold = True
    vFila = 2
    Obj_Hoja.Cells(vFila, 1) = "Periodo: " & CmbMesExtrae.Text
    Obj_Hoja.rows(vFila).Font.Bold = True
   
    vFila = 3
    
    For i = 0 To UBound(Titulos)
    
          Obj_Hoja.Cells(vFila, i + 1) = Titulos(i)
    
    Next
  
     
        Obj_Hoja.Columns("O").NumberFormat = "@"
        Obj_Hoja.Columns("S").NumberFormat = "@"
        Obj_Hoja.Columns("F").NumberFormat = "@"
        Obj_Hoja.Columns("S").HorizontalAlignment = xlHAlignRight
        Obj_Hoja.Columns("E").HorizontalAlignment = xlHAlignRight
        Obj_Hoja.Columns("F").HorizontalAlignment = xlHAlignRight
        Obj_Hoja.Columns("AM").HorizontalAlignment = xlHAlignRight
        Obj_Hoja.Columns("W:X").NumberFormat = "######,##0.00"
        Obj_Hoja.Columns("Z:AH").NumberFormat = "######,##0.00"
        Obj_Hoja.Columns("AJ").NumberFormat = "######,##0.00"
      
        Dim FilaPrint As String
       
      Do While Not rs.EOF
               vFila = vFila + 1
              For vColumna = 0 To vTotaCampos - 1
              
                Select Case vColumna
                    Case 4, 18 'Fecha
                        Dim vfecha As String
                        Dim vFfinal As String
          
                        If Len((Trim(rs.Fields(vColumna).Value))) > 0 Then
                           vfecha = rs.Fields(vColumna).Value
                           vfinal = Mid(vfecha, 7, 2) & "/" & Mid(vfecha, 5, 2) & "/" & Mid(vfecha, 1, 4)
                           Obj_Hoja.Cells(vFila, vColumna + 1) = vfinal
                         Else
                            Obj_Hoja.Cells(vFila, vColumna + 1) = " "
                        End If
             
                          
                                  
                    Case Else
                    
                    Obj_Hoja.Cells(vFila, vColumna + 1) = rs.Fields(vColumna).Value
                End Select
                
             Next
                
             For vColumna = 0 To vTotaCampos - 1
      
                '***********INI Columna de archivo texto
                 Select Case vColumna
                 Case 2, 7, 37 'Tipo Beneficio, Modalidad, banco
                    FilaPrint = FilaPrint & Mid(rs.Fields(vColumna).Value, 1, 3)
                Case 3, 6, 19, 20, 21, 36, 40 'Tipo solicitud, Etapa de Pago, parentesco,sexo, moneda, forma pago, Situacion Cobertura
                    FilaPrint = FilaPrint & Mid(rs.Fields(vColumna).Value, 1, 1)
                Case 8, 13 'Periodo pago, tipo documento
                    FilaPrint = FilaPrint & Mid(rs.Fields(vColumna).Value, 1, 2)
                Case 9 'origen del fondo
                    FilaPrint = FilaPrint & Mid(rs.Fields(vColumna).Value, 1, 5)
                Case 22, 23, 28, 32, 35 'mnto bruto total, pension bruta mensual,essalud, monto judicial, pension neta mensual
                    FilaPrint = FilaPrint & Replace(Format(rs.Fields(vColumna).Value, "0000000000000000.00"), ".", "")
                 Case 24 'Meses devengados
                    FilaPrint = FilaPrint & Format(rs.Fields(vColumna).Value, "000")
                Case Else
                 FilaPrint = FilaPrint & rs.Fields(vColumna).Value
               End Select
                      
           Next
           Obj_Hoja.Cells(vFila, 42) = FilaPrint
            FilaPrint = ""
          '***********FIN Columna de archivo texto
    
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
     Obj_Hoja.Range("A3:AO3").Borders.LineStyle = xlContinuous
     Obj_Hoja.Range("A3:AO3").Interior.Color = RGB(0, 32, 96)
     Obj_Hoja.Range("A3:AO3").Font.Color = RGB(237, 125, 49)
     Obj_Hoja.Range("A3:AO3").RowHeight = 34.5
     Obj_Hoja.Range("A3:AO3").VerticalAlignment = xlVAlignCenter


      Obj_Hoja.Columns("A").ColumnWidth = 5.57
      Obj_Hoja.Columns("D").ColumnWidth = 48
      Obj_Hoja.Columns("H").ColumnWidth = 48
      Obj_Hoja.Columns("I").ColumnWidth = 48
  
     
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
Private Sub GeneraArchivo()

       Dim variable As String
  
        Dim vTotaReg As Integer
        Dim vTotaFilas As Integer
        Dim NombreArchivo As String
        Dim PrimeraLinea As String
        
        
        
        Dim ultimodiaMes As Integer
        Dim anio As Integer
        Dim mes As Integer
                
        anio = Mid(CmbMesExtrae.Text, 1, 4)
        mes = Mid(CmbMesExtrae.Text, 5, 2)

       ultimodiaMes = Day(DateSerial(anio, mes + 1, 0))
       
        NombreArchivo = "01" & CStr(anio - 2000) & Format(CStr(mes), "00") & Format(CStr(ultimodiaMes), "00") & "." & "730"
    
        PrimeraLinea = "07300100209" & anio & Format(CStr(mes), "00") & Format(CStr(ultimodiaMes), "00") & "0120" & "|"
        
        VentSelecFichero.DialogTitle = "Ubicación del archivo"
        VentSelecFichero.InitDir = "C:\"
        VentSelecFichero.Flags = cdlOFNHideReadOnly
        'VentSelecFichero.Filter = "Archivos de texto|*.txt"
        VentSelecFichero.FileName = "C:\" & NombreArchivo
        VentSelecFichero.ShowOpen
        variable = VentSelecFichero.FileName
   
          
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
      
        objCmd.CommandText = "PKG_SBSINFORMES.formato0730_SBS"
        objCmd.CommandType = adCmdStoredProc
        
        Set param1 = objCmd.CreateParameter("p_periodo", adVarChar, adParamInput, 6, CmbMesExtrae.Text)
        objCmd.Parameters.Append param1
        
        Set rs = objCmd.Execute

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
      
      
      
    Static obj As Object
    If obj Is Nothing Then Set obj = VBA.CreateObject("ADODB.Stream")
    obj.Open
    obj.Charset = "windows-1252"
    obj.Type = 2
   
        Dim FilaPrint As String
        
        obj.WriteText PrimeraLinea & vbCrLf
               
      
        Do While Not rs.EOF
               vFila = vFila + 1
          
              For vColumna = 0 To vTotaCampos - 1
              
               Select Case vColumna
                 Case 2, 7, 37 'Tipo Beneficio, Modalidad, banco
                    FilaPrint = FilaPrint & Mid(rs.Fields(vColumna).Value, 1, 3)
                Case 3, 6, 19, 20, 21, 36, 40 'Tipo solicitud, Etapa de Pago, parentesco,sexo, moneda, forma pago, Situacion Cobertura
                    FilaPrint = FilaPrint & Mid(rs.Fields(vColumna).Value, 1, 1)
                Case 8, 13 'Periodo pago, tipo documento
                    FilaPrint = FilaPrint & Mid(rs.Fields(vColumna).Value, 1, 2)
                Case 9 'origen del fondo
                    FilaPrint = FilaPrint & Mid(rs.Fields(vColumna).Value, 1, 5)
                Case 22, 23, 28, 32, 35 'mnto bruto total, pension bruta mensual,essalud, monto judicial, pension neta mensual
                    FilaPrint = FilaPrint & Replace(Format(rs.Fields(vColumna).Value, "0000000000000000.00"), ".", "")
                 Case 24 'Meses devengados
                    FilaPrint = FilaPrint & Format(rs.Fields(vColumna).Value, "000")
                Case Else
                 FilaPrint = FilaPrint & rs.Fields(vColumna).Value
               End Select
              
              Next
              
              obj.WriteText FilaPrint & vbCrLf
              FilaPrint = ""
            
              
            BarraProgreso.Value = vFila
            lblprogreso.Caption = "Procesando " & BarraProgreso.Value & " de " & vTotaReg & " Registros."
            Me.Refresh
              
            rs.MoveNext
    Loop
    
    
   If Dir(variable, vbArchive) <> "" Then
    Kill (variable)
   End If
 
    
   obj.SaveToFile variable
   obj.Close


End Sub












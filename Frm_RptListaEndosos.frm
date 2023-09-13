VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Frm_RptListaEndosos 
   Caption         =   "Lista de Endosos"
   ClientHeight    =   2460
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   6435
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   6375
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   855
         Left            =   4080
         Picture         =   "Frm_RptListaEndosos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   855
         Left            =   5160
         Picture         =   "Frm_RptListaEndosos.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   960
         Width           =   855
      End
      Begin MSComctlLib.ProgressBar BarraProgreso 
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblprogreso 
         Caption         =   "Label2"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   960
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
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Label lbltitulorpt 
      Alignment       =   2  'Center
      Caption         =   "Lista de Endosos"
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
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "Frm_RptListaEndosos"
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
      
    objCmd.CommandText = "PKG_Informes_Administrativos.SP_LISTA_ENDOSOS"
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
    Dim Titulos(9) As String

    Titulos(0) = "PÓLIZA"
    Titulos(1) = "ENDOSO"
    Titulos(2) = "TIPO"
    Titulos(3) = "CAUSA"
    Titulos(4) = "PER DIFEREIDO"
    Titulos(5) = "MTO PENSIÓN"
    Titulos(6) = "PER GARANTIZADO"
    Titulos(7) = "MTO PENSIÓN GARANTIZADO"
    Titulos(8) = "FECHA CREA"

    vFila = 1
    Obj_Hoja.Cells(vFila, 1) = "Lista de Endosos"
    Obj_Hoja.Rows(vFila).Font.Bold = True
    
    Obj_Hoja.Columns("A").NumberFormat = "@"
    Obj_Hoja.Columns("B").NumberFormat = "@"
    Obj_Hoja.Columns("F").NumberFormat = "######,##0.00"
    Obj_Hoja.Columns("H").NumberFormat = "######,##0.00"
      
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
    
    Obj_Hoja.Columns("A:I").AutoFit
    Obj_Hoja.Rows(2).Font.Bold = True
    Obj_Hoja.Columns("A").ColumnWidth = 18
    Obj_Hoja.Range("A3:I3").Borders.LineStyle = xlContinuous
    Obj_Hoja.Range("A3:I3").Interior.Color = RGB(0, 32, 96)
    Obj_Hoja.Range("A3:I3").Font.Color = RGB(237, 125, 49)
    Obj_Hoja.Range("A3:I3").RowHeight = 34.5
    Obj_Hoja.Range("A3:I3").VerticalAlignment = xlVAlignCenter
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

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.lblMensaje.Caption = ""
    Me.lblprogreso.Caption = ""
End Sub

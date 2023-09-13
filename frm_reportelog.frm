VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_reportelog 
   Caption         =   "Reporte de Log"
   ClientHeight    =   5970
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12225
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   12225
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Filtros de Busqueda"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      Begin VB.Frame Frame3 
         Caption         =   "Modulo"
         Height          =   735
         Left            =   480
         TabIndex        =   9
         Top             =   240
         Width           =   3375
         Begin VB.ComboBox cmb_modulog 
            Height          =   315
            Left            =   240
            TabIndex        =   10
            Text            =   "TODOS"
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Usuario"
         Height          =   735
         Left            =   4080
         TabIndex        =   7
         Top             =   240
         Width           =   2655
         Begin VB.TextBox txt_usulog 
            Height          =   375
            Left            =   240
            TabIndex        =   8
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Fec. Inicio"
         Height          =   735
         Left            =   7200
         TabIndex        =   5
         Top             =   240
         Width           =   2055
         Begin MSMask.MaskEdBox txt_fecinilog 
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   10
            Format          =   "ddddd"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Fec. Fin"
         Height          =   735
         Left            =   9720
         TabIndex        =   3
         Top             =   240
         Width           =   2055
         Begin MSMask.MaskEdBox txt_fecfinlog 
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   10
            Format          =   "ddddd"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
      End
      Begin VB.CommandButton btn_buslog 
         Caption         =   "Buscar"
         Height          =   495
         Left            =   9360
         TabIndex        =   2
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton btn_explog 
         Caption         =   "Exportar"
         Height          =   495
         Left            =   10680
         TabIndex        =   1
         Top             =   1200
         Width           =   1095
      End
   End
   Begin TabDlg.SSTab fram_audlog 
      Height          =   4215
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   7435
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Reporte de Auditoria"
      TabPicture(0)   =   "frm_reportelog.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grilla_audlog"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Reporte de Sesion"
      TabPicture(1)   =   "frm_reportelog.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grilla_seslog"
      Tab(1).ControlCount=   1
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla_seslog 
         Height          =   2775
         Left            =   -74640
         TabIndex        =   12
         Top             =   600
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   4895
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla_audlog 
         Height          =   3255
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   5741
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "frm_reportelog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'CORPTEC
Dim fla_fecha As Integer

Public Function flCarga_sesLog()
   Dim objCommand As ADODB.Command
  
   Dim vlSql, valor, usuario, sistema, modulo, tipo, fec_ini, fec_fin As String
   sistema = "SEACSA"
   modulo = cmb_modulog.Text
   
   If modulo = "TODOS" Then
     modulo = ""
   End If
   valor = "Nulo"
   usuario = UCase(txt_usulog.Text)
   fec_ini = Format(txt_fecinilog.Text, "yyyymmdd")
   fec_fin = Format(txt_fecfinlog.Text, "yyyymmdd")
   
   
    Dim vgRs  As ADODB.Recordset

    vlSql = "SELECT    COALESCE(FEC_INI,'" & valor & "') AS FEC_INI , COALESCE(HOR_INI,'" & valor & "') AS HOR_INI, "
    vlSql = vlSql & "  COALESCE(FEC_FIN,'" & valor & "') AS FEC_FIN,  COALESCE(HOR_FIN, '" & valor & "') AS HOR_FIN, "
    vlSql = vlSql & "  COALESCE(COD_USUARIO,'" & valor & "') AS COD_USUARIO , "
    vlSql = vlSql & "  COALESCE(GLS_MODULO,'" & valor & "') AS GLS_MODULO"
    vlSql = vlSql & " FROM LOG_SESSION"
    vlSql = vlSql & " WHERE FEC_INI >= " & fec_ini & " "
    vlSql = vlSql & " AND   FEC_INI <= " & fec_fin & " "
    If (sistema <> "") Then
    vlSql = vlSql & " AND GLS_SISTEMA like '%" & sistema & "%' "
    End If
    
    If (modulo <> "") Then
    vlSql = vlSql & " AND GLS_MODULO like '%" & modulo & "%' "
    End If
    
    If (usuario <> "") Then
    vlSql = vlSql & " AND COD_USUARIO like '%" & usuario & "%' "
    End If
    
    vlSql = vlSql & " ORDER BY  FEC_INI DESC "
    
  ' MsgBox vlSQL
    Set vgRs = vgConexionBD.Execute(vlSql)
  
    If Not (vgRs.EOF) Then
            Call p_llena_grilla_seslog(vgRs)
    Else
            MsgBox "No se encontró información", vbCritical, "¡ERROR!..."
    End If
   

End Function


Public Function flCarga_Log()
   Dim objCommand As ADODB.Command
   Dim objRs As ADODB.Recordset
  
   Dim vlSql, valor, usuario, sistema, modulo, fec_ini, fec_fin As String
   sistema = "SEACSA"
   modulo = cmb_modulog.Text
   valor = "Nulo"
   If modulo = "TODOS" Then
     modulo = ""
   End If
   
   usuario = UCase(txt_usulog.Text)
   fec_ini = Format(txt_fecinilog.Text, "yyyymmdd")
   fec_fin = Format(txt_fecfinlog.Text, "yyyymmdd")
   Dim vgRs  As ADODB.Recordset

    vlSql = "SELECT    COALESCE(FEC_INS,'" & valor & "') AS FEC_INS , COALESCE(HOR_INS,'" & valor & "') AS HOR_INS, "
    vlSql = vlSql & "  COALESCE(FEC_FIN,'" & valor & "') AS FEC_FIN,  COALESCE(HOR_FIN, '" & valor & "') AS HOR_FIN, "
    vlSql = vlSql & "  COALESCE(COD_USUARIO,'" & valor & "') AS COD_USUARIO , COD_TIPO, "
    vlSql = vlSql & "  COALESCE(COD_TRANS,'" & valor & "') AS COD_TRANS , "
    vlSql = vlSql & "  COALESCE(GLS_MODULO,'" & valor & "') AS GLS_MODULO ,  COALESCE(GLS_TABLA,'" & valor & "') AS GLS_TABLA,"
    vlSql = vlSql & "  COALESCE(COD_IDTABLA,'" & valor & "') AS COD_IDTABLA, COALESCE(GLS_VALORTABLA,'" & valor & "') AS  GLS_VALORTABLA"
    vlSql = vlSql & " FROM LOG_TABLAS"
    vlSql = vlSql & " WHERE FEC_INS >= " & fec_ini & " "
    vlSql = vlSql & " AND   FEC_INS <= " & fec_fin & " "
    If (sistema <> "") Then
    vlSql = vlSql & " AND GLS_SISTEMA like '%" & sistema & "%' "
    End If
    
    If (modulo <> "") Then
    vlSql = vlSql & " AND GLS_MODULO like '%" & modulo & "%' "
    End If
    
    If (usuario <> "") Then
    vlSql = vlSql & " AND COD_USUARIO like '%" & usuario & "%' "
    End If
    vlSql = vlSql & " ORDER BY  FEC_INS DESC "
  '  MsgBox vlSQL
    Set vgRs = vgConexionBD.Execute(vlSql)
  
    If Not (vgRs.EOF) Then
            Call p_llena_grilla_log(vgRs)
    Else
            MsgBox "No se encontró información", vbCritical, "¡ERROR!..."
    End If
   
End Function



Public Sub p_titulo_grilla_log()
    grilla_audlog.Cols = 11
    
    grilla_audlog.ColWidth(0) = 950:     grilla_audlog.ColAlignmentFixed(0) = 4
    grilla_audlog.ColWidth(1) = 950:     grilla_audlog.ColAlignmentFixed(1) = 4
    grilla_audlog.ColWidth(2) = 950:     grilla_audlog.ColAlignmentFixed(2) = 4
    grilla_audlog.ColWidth(3) = 950:     grilla_audlog.ColAlignmentFixed(3) = 4
    grilla_audlog.ColWidth(4) = 950:     grilla_audlog.ColAlignmentFixed(4) = 4
    grilla_audlog.ColWidth(5) = 950:     grilla_audlog.ColAlignmentFixed(5) = 4
    grilla_audlog.ColWidth(6) = 1350:    grilla_audlog.ColAlignmentFixed(6) = 4
    grilla_audlog.ColWidth(7) = 1350:    grilla_audlog.ColAlignmentFixed(7) = 4
    grilla_audlog.ColWidth(8) = 1350:    grilla_audlog.ColAlignmentFixed(8) = 4
    grilla_audlog.ColWidth(9) = 1350:    grilla_audlog.ColAlignmentFixed(9) = 4
    grilla_audlog.ColWidth(10) = 1500:   grilla_audlog.ColAlignmentFixed(10) = 4
    
    grilla_audlog.Row = 0
    grilla_audlog.Col = 0: grilla_audlog.Text = "Fecha Inicio"
    grilla_audlog.Col = 1: grilla_audlog.Text = "Hora Inicio"
    grilla_audlog.Col = 2: grilla_audlog.Text = "Fecha Fin"
    grilla_audlog.Col = 3: grilla_audlog.Text = "Hora Fin"
    grilla_audlog.Col = 4: grilla_audlog.Text = "Usuario"
    grilla_audlog.Col = 5: grilla_audlog.Text = "Tipo"
    grilla_audlog.Col = 6: grilla_audlog.Text = "Modulo"
    grilla_audlog.Col = 7: grilla_audlog.Text = "Actividad"
    grilla_audlog.Col = 8: grilla_audlog.Text = "Tabla"
    grilla_audlog.Col = 9: grilla_audlog.Text = "ID Tabla"
    grilla_audlog.Col = 10: grilla_audlog.Text = "Nombre ID Tabla"
End Sub

Public Sub p_titulo_grilla_seslog()
    grilla_seslog.Cols = 6
    
    grilla_seslog.ColWidth(0) = 1050:     grilla_seslog.ColAlignmentFixed(0) = 4
    grilla_seslog.ColWidth(1) = 1050:     grilla_seslog.ColAlignmentFixed(1) = 4
    grilla_seslog.ColWidth(2) = 1050:     grilla_seslog.ColAlignmentFixed(2) = 4
    grilla_seslog.ColWidth(3) = 1050:     grilla_seslog.ColAlignmentFixed(3) = 4
    grilla_seslog.ColWidth(4) = 1050:     grilla_seslog.ColAlignmentFixed(4) = 4
    grilla_seslog.ColWidth(5) = 1050:     grilla_seslog.ColAlignmentFixed(5) = 4
    
    grilla_seslog.Row = 0
    grilla_seslog.Col = 0: grilla_seslog.Text = "Fecha Inicio"
    grilla_seslog.Col = 1: grilla_seslog.Text = "Hora Inicio"
    grilla_seslog.Col = 2: grilla_seslog.Text = "Fecha Fin"
    grilla_seslog.Col = 3: grilla_seslog.Text = "Hora Fin"
    grilla_seslog.Col = 4: grilla_seslog.Text = "Usuario"
    grilla_seslog.Col = 5: grilla_seslog.Text = "Modulo"
  
   
End Sub
Private Sub p_llena_grilla_log(rsLog As ADODB.Recordset)
    grilla_audlog.Clear
    Call p_titulo_grilla_log
    rsLog.MoveFirst
    
    grilla_audlog.Row = 1
   
    Do While Not rsLog.EOF
       
            grilla_audlog.Col = 0
            If rsLog.Fields("FEC_INS") = "Nulo" Then
              grilla_audlog.Text = ""
              Else
              grilla_audlog.Text = rsLog.Fields("FEC_INS")
            End If
            
            grilla_audlog.Col = 1
            If rsLog.Fields("HOR_INS") = "Nulo" Then
              grilla_audlog.Text = ""
              Else
              grilla_audlog.Text = rsLog.Fields("HOR_INS")
            End If
            
            grilla_audlog.Col = 2
            If rsLog.Fields("FEC_FIN") = "Nulo" Then
              grilla_audlog.Text = ""
              Else
              grilla_audlog.Text = rsLog.Fields("FEC_FIN")
            End If
            
            grilla_audlog.Col = 3
             If rsLog.Fields("HOR_FIN") = "Nulo" Then
              grilla_audlog.Text = ""
              Else
              grilla_audlog.Text = rsLog.Fields("HOR_FIN")
            End If
            
            grilla_audlog.Col = 4
            If rsLog.Fields("COD_USUARIO") = "Nulo" Then
              grilla_audlog.Text = ""
              Else
              grilla_audlog.Text = rsLog.Fields("COD_USUARIO")
            End If
            
            grilla_audlog.Col = 5
            If rsLog.Fields("COD_TIPO") = "A" Then
            grilla_audlog.Text = "ARCHIVO"
            End If
             If rsLog.Fields("COD_TIPO") = "A" Then
            grilla_audlog.Text = "TABLA"
            End If
             If rsLog.Fields("COD_TIPO") = "P" Then
            grilla_audlog.Text = "PROCESO"
            End If
            
            grilla_audlog.Col = 6
            If rsLog.Fields("GLS_MODULO") = "Nulo" Then
              grilla_audlog.Text = ""
              Else
              grilla_audlog.Text = rsLog.Fields("GLS_MODULO")
            End If
            
            grilla_audlog.Col = 7
            If rsLog.Fields("COD_TRANS") = "Nulo" Then
              grilla_audlog.Text = ""
              Else
              grilla_audlog.Text = rsLog.Fields("COD_TRANS")
            End If
            
            grilla_audlog.Col = 8
            If rsLog.Fields("GLS_TABLA") = "Nulo" Then
              grilla_audlog.Text = ""
              Else
              grilla_audlog.Text = rsLog.Fields("GLS_TABLA")
            End If
           
            grilla_audlog.Col = 9
            If rsLog.Fields("COD_IDTABLA") = "Nulo" Then
              grilla_audlog.Text = ""
              Else
              grilla_audlog.Text = rsLog.Fields("COD_IDTABLA")
            End If
            
            grilla_audlog.Col = 10
            If rsLog.Fields("GLS_VALORTABLA") = "Nulo" Then
              grilla_audlog.Text = ""
              Else
              grilla_audlog.Text = rsLog.Fields("GLS_VALORTABLA")
            End If
            
            
            grilla_audlog.Rows = grilla_audlog.Rows + 1
            grilla_audlog.Row = grilla_audlog.Rows - 1
          
        
        rsLog.MoveNext
    Loop
End Sub

Private Sub p_llena_grilla_seslog(rsLog As ADODB.Recordset)
    grilla_seslog.Clear
    Call p_titulo_grilla_seslog
    rsLog.MoveFirst
    
    grilla_seslog.Row = 1
   
    Do While Not rsLog.EOF
       
            grilla_seslog.Col = 0
             If rsLog.Fields("FEC_INI") = "Nulo" Then
              grilla_seslog.Text = ""
              Else
              grilla_seslog.Text = rsLog.Fields("FEC_INI")
            End If
            
            grilla_seslog.Col = 1
             If rsLog.Fields("HOR_INI") = "Nulo" Then
              grilla_seslog.Text = ""
              Else
              grilla_seslog.Text = rsLog.Fields("HOR_INI")
            End If
            
            grilla_seslog.Col = 2
            If rsLog.Fields("FEC_FIN") = "Nulo" Then
              grilla_seslog.Text = ""
              Else
              grilla_seslog.Text = rsLog.Fields("FEC_FIN")
            End If
            
            grilla_seslog.Col = 3
            If rsLog.Fields("HOR_FIN") = "Nulo" Then
              grilla_seslog.Text = ""
              Else
              grilla_seslog.Text = rsLog.Fields("HOR_FIN")
            End If
            
            grilla_seslog.Col = 4
            If rsLog.Fields("COD_USUARIO") = "Nulo" Then
              grilla_seslog.Text = ""
              Else
              grilla_seslog.Text = rsLog.Fields("COD_USUARIO")
            End If
            
            grilla_seslog.Col = 5
            If rsLog.Fields("GLS_MODULO") = "Nulo" Then
              grilla_seslog.Text = ""
              Else
              grilla_seslog.Text = rsLog.Fields("GLS_MODULO")
            End If
           
            grilla_seslog.Rows = grilla_seslog.Rows + 1
            grilla_seslog.Row = grilla_seslog.Rows - 1
          
        rsLog.MoveNext
    Loop
End Sub

Public Sub reinicia_grilla_seslog()
    grilla_seslog.Clear
    grilla_seslog.Rows = 2
    grilla_seslog.Cols = 2
End Sub
Public Sub reinicia_grilla_log()
    grilla_audlog.Clear
    grilla_audlog.Rows = 2
    grilla_audlog.Cols = 2
End Sub

Private Sub expor_log()
        Dim i As Long, j As Long
        Dim objExcel As Object
        Dim objWorkbook As Object
        On Error Resume Next ' por si se cierra Excel antes de cargar los datos
        Set objExcel = CreateObject("Excel.Application")
        objExcel.Visible = True
        Set objWorkbook = objExcel.Workbooks.Add
        For i = 0 To grilla_audlog.Rows - 1
        grilla_audlog.Row = i
        For j = 0 To grilla_audlog.Cols - 1
        grilla_audlog.Col = j
        objWorkbook.ActiveSheet.Cells(i + 1, j + 1).Value = grilla_audlog.Text
        Next
        Next
        objExcel.Cells.Select
        objExcel.Selection.EntireColumn.AutoFit ' Ancho de columna
        objExcel.Range("A1").Select
       '' objExcel.ActiveWindow.SelectedSheets.PrintPreview ' Previsualizar informe
        Set objWorkbook = Nothing
        Set objExcel = Nothing
End Sub

Private Sub expor_seslog()
        Dim i As Long, j As Long
        Dim objExcel As Object
        Dim objWorkbook As Object
        On Error Resume Next ' por si se cierra Excel antes de cargar los datos
        Set objExcel = CreateObject("Excel.Application")
        objExcel.Visible = True
        Set objWorkbook = objExcel.Workbooks.Add
        For i = 0 To grilla_seslog.Rows - 1
        grilla_seslog.Row = i
        For j = 0 To grilla_seslog.Cols - 1
        grilla_seslog.Col = j
        objWorkbook.ActiveSheet.Cells(i + 1, j + 1).Value = grilla_seslog.Text
        Next
        Next
        objExcel.Cells.Select
        objExcel.Selection.EntireColumn.AutoFit ' Ancho de columna
        objExcel.Range("A1").Select
       '' objExcel.ActiveWindow.SelectedSheets.PrintPreview ' Previsualizar informe
        Set objWorkbook = Nothing
        Set objExcel = Nothing
End Sub

Private Sub btn_buslog_Click()
   Call verifica_fecha_log
   If fla_fecha = 0 Then
            If fram_audlog.Tab = 1 Then
             Call reinicia_grilla_seslog
             Call flCarga_sesLog
            Else
             Call reinicia_grilla_log
             Call flCarga_Log
            End If
    End If
    
End Sub

Private Sub verifica_fecha_log()
    fla_fecha = 0
    Dim fecmayor As Integer
    fecmayor = 0
    
    If Not (IsDate(txt_fecinilog.Text)) Or Trim(txt_fecinilog.Text) = "__/__/____" Then
        MsgBox "Debe ingresar la Fecha de Inicio.", vbCritical, "¡ERROR!..."
        txt_fecinilog.SetFocus
        fla_fecha = 1
        fecmayor = 1
    End If
    If Not (IsDate(txt_fecfinlog.Text)) Or Trim(txt_fecfinlog.Text) = "__/__/____" Then
        MsgBox "Debe ingresar la Fecha de  Fin.", vbCritical, "¡ERROR!..."
        fla_fecha = 1
        fecmayor = 1
    End If
    
    If fecmayor = 0 Then
    If DateValue(txt_fecinilog.Text) > DateValue(txt_fecfinlog.Text) Then
        MsgBox "Fecha de Inicio debe ser menor a Fecha Fin .", vbCritical, "¡ERROR!..."
        txt_fecinilog.SetFocus
        fla_fecha = 1
    End If
        
    End If
  
   
End Sub

Private Sub btn_explog_Click()
    If fram_audlog.Tab = 1 Then
      Call expor_seslog
     Else
      Call expor_log
     End If
End Sub

Private Sub Carga_combolog()
  cmb_modulog.AddItem "TODOS"
  cmb_modulog.AddItem "COTIZACION"
  cmb_modulog.AddItem "PENSIONES"
  cmb_modulog.AddItem "PRODUCCION"
  cmb_modulog.AddItem "RESERVAS"
     
End Sub

Private Sub Form_Load()
  Call Carga_combolog
  cmb_modulog.ListIndex = 2
  txt_fecinilog.Text = Format(Now, "dd/mm/yyyy")
  txt_fecfinlog.Text = Format(Now, "dd/mm/yyyy")
End Sub




VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReliqMasiva 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reliquidación por rango de fechas"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13530
   Icon            =   "frmReliqMasiva.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   13530
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   120
      TabIndex        =   10
      Top             =   1215
      Width           =   13320
      _ExtentX        =   23495
      _ExtentY        =   11033
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Detalle de certificados pendientes o reliquidados"
      TabPicture(0)   =   "frmReliqMasiva.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkTodos"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraDetalle"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtMontoTot"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtCuotas"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Pólizas con certificados vencidos"
      TabPicture(1)   =   "frmReliqMasiva.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdImprimirPend"
      Tab(1).Control(1)=   "DtgCertVenc"
      Tab(1).Control(2)=   "Label5"
      Tab(1).Control(3)=   "Label4"
      Tab(1).ControlCount=   4
      Begin VB.TextBox txtCuotas 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   9180
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   5805
         Width           =   810
      End
      Begin VB.TextBox txtMontoTot 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   11625
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   5790
         Width           =   1395
      End
      Begin VB.CommandButton cmdImprimirPend 
         Caption         =   "Imprimir"
         Height          =   435
         Left            =   -69195
         TabIndex        =   23
         Top             =   5640
         Width           =   2730
      End
      Begin MSDataGridLib.DataGrid DtgCertVenc 
         Height          =   4485
         Left            =   -74820
         TabIndex        =   20
         Top             =   990
         Width           =   12945
         _ExtentX        =   22834
         _ExtentY        =   7911
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame3 
         Caption         =   "DETALLE DE AVANCE DEL PROCESO DE RELIQUIDACIÓN"
         Height          =   1080
         Left            =   4590
         TabIndex        =   18
         Top             =   2880
         Visible         =   0   'False
         Width           =   4950
         Begin MSComctlLib.ProgressBar pbAbanse 
            Height          =   345
            Left            =   270
            TabIndex        =   19
            Top             =   435
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   609
            _Version        =   393216
            Appearance      =   1
         End
      End
      Begin VB.Frame fraDetalle 
         Height          =   2490
         Left            =   90
         TabIndex        =   15
         Top             =   3225
         Width           =   13065
         Begin MSFlexGridLib.MSFlexGrid msfDetalle 
            Height          =   1860
            Left            =   105
            TabIndex        =   16
            Top             =   525
            Width           =   12855
            _ExtentX        =   22675
            _ExtentY        =   3281
            _Version        =   393216
         End
         Begin VB.Label Label3 
            Caption         =   "Detalle de Cuotas Pendientes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   165
            TabIndex        =   17
            Top             =   195
            Width           =   3495
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2520
         Left            =   75
         TabIndex        =   12
         Top             =   720
         Width           =   13080
         Begin MSFlexGridLib.MSFlexGrid MSHLista 
            Height          =   1935
            Left            =   105
            TabIndex        =   13
            Top             =   495
            Width           =   12870
            _ExtentX        =   22701
            _ExtentY        =   3413
            _Version        =   393216
            FixedCols       =   0
            MergeCells      =   1
         End
         Begin VB.Label lblTitulo1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   150
            TabIndex        =   14
            Top             =   180
            Width           =   4980
         End
      End
      Begin VB.CheckBox chkTodos 
         Caption         =   "Marcar Todos"
         Height          =   315
         Left            =   360
         TabIndex        =   11
         Top             =   450
         Width           =   1410
      End
      Begin VB.Label Label6 
         Caption         =   "Monto Total :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10125
         TabIndex        =   27
         Top             =   5820
         Width           =   1500
      End
      Begin VB.Label Label7 
         Caption         =   "Numero de cuotas pendientes :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5580
         TabIndex        =   26
         Top             =   5820
         Width           =   3525
      End
      Begin VB.Label Label5 
         Caption         =   "Para ordenar la lista por cualquiera de los campos hacer click en cualquier cabecera"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   -74820
         TabIndex        =   22
         Top             =   720
         Width           =   6375
      End
      Begin VB.Label Label4 
         Caption         =   "Listado de pólizas con certificados cuya fecha de inicio y fecha de termino sea menor a la fecha de pago"
         Height          =   315
         Left            =   -74820
         TabIndex        =   21
         Top             =   465
         Width           =   9720
      End
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Height          =   990
      Left            =   7755
      Picture         =   "frmReliqMasiva.frx":047A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   165
      Width           =   960
   End
   Begin VB.OptionButton optReliquidados 
      Caption         =   "Reliquidados"
      Height          =   195
      Left            =   5070
      TabIndex        =   8
      Top             =   675
      Width           =   1365
   End
   Begin VB.OptionButton optPendiente 
      Caption         =   "Pendientes"
      Height          =   225
      Left            =   5070
      TabIndex        =   7
      Top             =   300
      Width           =   1335
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   990
      Left            =   9810
      Picture         =   "frmReliqMasiva.frx":08BC
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   165
      Width           =   990
   End
   Begin VB.CommandButton cmdReliquidar 
      Caption         =   "Reliquidar"
      Height          =   990
      Left            =   8790
      Picture         =   "frmReliqMasiva.frx":5C9E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   165
      Width           =   945
   End
   Begin VB.Frame Frame1 
      Height          =   990
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   4695
      Begin MSComCtl2.DTPicker dtpInicio 
         Height          =   300
         Left            =   2865
         TabIndex        =   3
         Top             =   210
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   393216
         Format          =   50069505
         CurrentDate     =   40302
      End
      Begin MSComCtl2.DTPicker dtpFinal 
         Height          =   300
         Left            =   2865
         TabIndex        =   4
         Top             =   600
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   50069505
         CurrentDate     =   40302
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha de pago del mes actual :"
         Height          =   210
         Left            =   375
         TabIndex        =   6
         Top             =   645
         Width           =   2325
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de certificado Desde :"
         Height          =   240
         Left            =   375
         TabIndex        =   5
         Top             =   270
         Width           =   2280
      End
   End
End
Attribute VB_Name = "frmReliqMasiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim rsStruc As ADODB.Recordset
Dim RSGEN As ADODB.Recordset

Private rsdet As ADODB.Recordset
Private TipLetra As String
Private C_sSeleccion As String
Private C_sTipLetraSel As String
Private vlConceptoPEMenor As String
Private vlConceptoPEMayor As String
Private vlConceptoPRentaRetro As String
Private filtro As String


Private Sub chkTodos_Click()
Dim i As Integer
    Select Case chkTodos.Value
    Case 1:
           For i = 1 To MSHLista.Rows - 1
                MSHLista.Col = 0
                MSHLista.Row = i
                MSHLista.CellAlignment = 4
                MSHLista.CellFontName = C_sTipLetraSel
                MSHLista.CellForeColor = &HFF&
                MSHLista.TextMatrix(i, 0) = C_sSeleccion
           Next i
    Case 0:
           For i = 1 To MSHLista.Rows - 1
                MSHLista.TextMatrix(i, 0) = ""
           Next i
    End Select
End Sub

Private Sub cmdBuscar_Click()
If optReliquidados.Value = True Then
    cmdReliquidar.Enabled = False
    lblTitulo1.Caption = "Listado de certificados reliquidados"
    txtCuotas.Text = Format(0, "###,##0.00")
    txtMontoTot.Text = Format(0, "###,##0.00")
Else
    cmdReliquidar.Enabled = True
    lblTitulo1.Caption = "Listado de certificados pendientes de pago"
End If
Call CargaDatos
End Sub

Private Sub Command3_Click()
fraDetalle.Visible = False
End Sub

Function flObtieneNumReliq() As Double
Dim vlSql As String
Dim vlTB As ADODB.Recordset
vlSql = "SELECT MAX(num_reliq) as reliq FROM PP_TMAE_RELIQ"
Set vlTB = vgConexionBD.Execute(vlSql)
If Not vlTB.EOF Then
    flObtieneNumReliq = IIf(IsNull(vlTB!reliq), 0, vlTB!reliq) + 1
Else
    flObtieneNumReliq = 1
End If
End Function

Private Sub cmdImprimir_Click()

   If optPendiente.Value = True Then
        Call ReportePendientes
   Else
        Call ReporteReliquidados
   End If
    
End Sub

Private Sub Crea_Estructura()

Set rsStruc = New ADODB.Recordset
    With rsStruc.Fields
        .Append "AFP", adVarChar, 50
        .Append "Cod_Pensión", adVarChar, 50
        .Append "num_poliza", adVarChar, 50
        
        .Append "gls_nomben", adVarChar, 50
        .Append "gls_patben", adVarChar, 50
        .Append "gls_matben", adVarChar, 50
        
        .Append "Cuotas_retenidas", adVarChar, 50
        .Append "Fecha_Reliq", adDate
        .Append "mto_pension", adVarNumeric
        
        .Append "Essalud", adVarNumeric
        .Append "retensión", adVarNumeric
        .Append "Periodo_de_Pago", adVarChar, 50
        
    End With

End Sub

Private Sub ReporteReliquidados()

Dim objRep As New ClsReporte

    On Error GoTo mierror
   
    'Call Crea_Estructura
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

    rs.Open "PP_LISTA_RELIQUIDADOS.LISTAR('" & Format(dtpInicio, "yyyyMMdd") & "','" & Format(dtpFinal, "yyyyMMdd") & "')", vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF Then
        MsgBox "No hay datos para mostrar", vbExclamation, "Pago de pensiones"
        Exit Sub
    End If
    
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PP_Rpt_Cert_Reliquidados.rpt"), ".RPT", ".TTX"), 1)
    
        
    If objRep.CargaReporte(strRpt & "", "PP_Rpt_Cert_Reliquidados.rpt", "Informe Certificados Reliquidados", rs, True, _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("NombreSistema", vgNombreSistema), _
                            ArrFormulas("NombreSubSistema", vgNombreSubSistema)) = False Then
                            
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
    End If
    
    Exit Sub
mierror:
    MsgBox "No se pudo cargar el reporte", vbInformation
    
End Sub

Private Function Crea_temporal_Reliq() As Boolean
Dim cad As String
Dim rs As ADODB.Recordset
On Error GoTo mierror

cad = "select * from all_tables where TABLE_NAME='PT_TMP_RELQPEND" & vgUsuario & "'"
Set rs = vgConexionBD.Execute(cad)
If Not rs.EOF Then
    cad = "DELETE FROM PT_TMP_RELQPEND" & vgUsuario
    vgConexionBD.Execute (cad)
Else
    cad = "create table PT_TMP_RELQPEND" & vgUsuario & "(NUM_POLIZA VARCHAR2(10),NUM_ENDOSO NUMBER,NUM_ORDEN NUMBER,GLS_NOMBEN VARCHAR2(50),GLS_PATBEN VARCHAR2(50),GLS_MATBEN VARCHAR2(50)," & _
            " COD_ESTPENSION VARCHAR(10),COD_TIPOIDENBEN VARCHAR(10),NUM_IDENBEN VARCHAR(20),PRC_PENSION NUMBER,FEC_NACBEN VARCHAR(10),COD_MONEDA VARCHAR(10)," & _
            " COD_TIPPENSION VARCHAR(10),FECHA_RELIQ VARCHAR2(10), PENSIÓN NUMBER,RETENSIÓN NUMBER,ESSALUD NUMBER,FEC_INICER VARCHAR(10),FEC_TERCER VARCHAR(10)," & _
            " MTO_PENSION NUMBER,AFP VARCHAR(50),DESC_MONEDA VARCHAR(50),CUOTAS_RETENIDAS NUMBER,CODPAR VARCHAR2(10),SITINV VARCHAR2(10),RANGO_PENDIENTE VARCHAR(1000),LARGO NUMBER,EDAD NUMBER)"
            
    vgConexionBD.Execute (cad)
End If
Crea_temporal_Reliq = True
Exit Function
mierror:
    MsgBox "Problemas al crear el temporal", vbCritical
End Function

Private Sub ReportePendientes()
Dim estadoRep As Integer
Dim objRep As New ClsReporte
Dim RANGOTOTAL As String
Dim RANGOINI As String
Dim RANGOFIN As String

    On Error GoTo mierror
    If Crea_temporal_Reliq = False Then
        MsgBox "No se pudo crear temporal", vbExclamation, "Pensiones"
        Exit Sub
    End If
    DoEvents
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    vgConexionBD.Execute "PD_TMP_RELIQ_PEND('" & Format(dtpInicio, "yyyyMMdd") & "','" & Format(dtpFinal, "yyyyMMdd") & "','" & vgUsuario & "')"
    'rs.Open "PP_LISTA_RELIQUIDACION.LISTAR('" & Format(dtpInicio, "yyyyMMdd") & "','" & Format(dtpFinal, "yyyyMMdd") & "')", vgConexionBD, adOpenForwardOnly, adLockReadOnly
    'Set rs = vgConexionBD.Execute("PD_TMP_RELIQ_PEND('" & Format(dtpInicio, "yyyyMMdd") & "','" & Format(dtpFinal, "yyyyMMdd") & "','" & vgUsuario & "')")
    rs.Open "select T.cod_moneda as MONEDA,T.AFP,T.cod_estpension as COD_PENSIÓN,T.NUM_POLIZA,T.GLS_NOMBEN,T.GLS_PATBEN,T.GLS_MATBEN," & _
            " max(T.CUOTAS_RETENIDAS)as CUOTAS_RETENIDAS,T.FECHA_RELIQ,T.MTO_PENSION,T.PENSIÓN,T.RETENSIÓN,T.ESSALUD," & _
            " T.RANGO_PENDIENTE as PERIODO_DE_PAGO,T.DESC_MONEDA,T.LARGO,T.EDAD,T.CODPAR from PT_TMP_RELQPEND" & vgUsuario & " T" & _
            " WHERE LARGO=(SELECT MAX(LARGO) FROM PT_TMP_RELQPEND" & vgUsuario & _
            " WHERE NUM_POLIZA=T.NUM_POLIZA AND NUM_ENDOSO=T.NUM_ENDOSO AND NUM_ORDEN=T.NUM_ORDEN)" & _
            " group by T.cod_moneda,T.AFP,T.cod_estpension,T.NUM_POLIZA,T.GLS_NOMBEN,T.GLS_PATBEN,T.GLS_MATBEN,T.FECHA_RELIQ," & _
            " T.Mto_Pension , T.PENSIÓN, T.RETENSIÓN, T.ESSALUD, T.fec_inicer, T.DESC_MONEDA, T.RANGO_PENDIENTE, T.LARGO,T.EDAD,T.CODPAR", vgConexionBD, adOpenForwardOnly, adLockReadOnly
    '(to_char(to_date(fec_inicer,'yyyyMMdd'),'yyyymm') ||'-'||'" & Format(dtpFinal, "yyyymm") & "')
    estadoRep = 0
    If rs.EOF Then
        MsgBox "No hay datos para mostrar", vbExclamation, "Pago de pensiones"
        Exit Sub
    Else
        Do While Not rs.EOF
            If CDbl(rs!CUOTAS_RETENIDAS) > 0 Then
                estadoRep = estadoRep + 1
                Exit Do
            End If
            rs.MoveNext
        Loop
        If estadoRep = 0 Then
            MsgBox "No hay datos para mostrar", vbExclamation, "Pago de pensiones"
            Exit Sub
        End If
    End If
    
    Call CREAR_ESTRUCTURA
   
    Do While Not rs.EOF
        If rs.Fields("CUOTAS_RETENIDAS") > 0 Then
            If (rs.Fields("Edad") < 18 And rs.Fields("CODPAR") >= "30" And rs.Fields("CODPAR") <= "35") Then
                rsStruc.AddNew
                rsStruc.Fields("MONEDA").Value = rs!Moneda
                rsStruc.Fields("AFP").Value = rs!AFP
                rsStruc.Fields("COD_PENSIÓN").Value = rs!COD_PENSIÓN
                rsStruc.Fields("NUM_POLIZA").Value = rs!num_poliza
                rsStruc.Fields("GLS_NOMBEN").Value = rs!Gls_NomBen
                rsStruc.Fields("GLS_PATBEN").Value = rs!Gls_PatBen
                rsStruc.Fields("GLS_MATBEN").Value = rs!Gls_MatBen
                rsStruc.Fields("CUOTAS_RETENIDAS").Value = rs!CUOTAS_RETENIDAS
                rsStruc.Fields("FECHA_RELIQ").Value = rs!FECHA_RELIQ
                rsStruc.Fields("MTO_PENSION").Value = rs!Mto_Pension
                rsStruc.Fields("PENSIÓN").Value = rs!PENSIÓN
                rsStruc.Fields("RETENSIÓN").Value = rs!RETENSIÓN
                rsStruc.Fields("ESSALUD").Value = rs!ESSALUD
                
                RANGOTOTAL = Mid(rs!PERIODO_DE_PAGO, 1, Len(rs!PERIODO_DE_PAGO) - 1)
                RANGOINI = Mid(RANGOTOTAL, 1, 7)
                RANGOFIN = Mid(RANGOTOTAL, Len(RANGOTOTAL) - 6, Len(RANGOTOTAL))
                rsStruc.Fields("PERIODO_DE_PAGO").Value = RANGOINI & "-" & RANGOFIN
                rsStruc.Fields("DESC_MONEDA").Value = rs!DESC_MONEDA
                rsStruc.Update
            End If
            If (rs.Fields("Edad") > 18 And rs.Fields("CODPAR") >= "10") Then
                rsStruc.AddNew
                rsStruc.Fields("MONEDA").Value = rs!Moneda
                rsStruc.Fields("AFP").Value = rs!AFP
                rsStruc.Fields("COD_PENSIÓN").Value = rs!COD_PENSIÓN
                rsStruc.Fields("NUM_POLIZA").Value = rs!num_poliza
                rsStruc.Fields("GLS_NOMBEN").Value = rs!Gls_NomBen
                rsStruc.Fields("GLS_PATBEN").Value = rs!Gls_PatBen
                rsStruc.Fields("GLS_MATBEN").Value = rs!Gls_MatBen
                rsStruc.Fields("CUOTAS_RETENIDAS").Value = rs!CUOTAS_RETENIDAS
                rsStruc.Fields("FECHA_RELIQ").Value = rs!FECHA_RELIQ
                rsStruc.Fields("MTO_PENSION").Value = rs!Mto_Pension
                rsStruc.Fields("PENSIÓN").Value = rs!PENSIÓN
                rsStruc.Fields("RETENSIÓN").Value = rs!RETENSIÓN
                rsStruc.Fields("ESSALUD").Value = rs!ESSALUD

                RANGOTOTAL = Mid(rs!PERIODO_DE_PAGO, 1, Len(rs!PERIODO_DE_PAGO) - 1)
                RANGOINI = Mid(RANGOTOTAL, 1, 7)
                RANGOFIN = Mid(RANGOTOTAL, Len(RANGOTOTAL) - 6, Len(RANGOTOTAL))
                rsStruc.Fields("PERIODO_DE_PAGO").Value = RANGOINI & "-" & RANGOFIN
                rsStruc.Fields("DESC_MONEDA").Value = rs!DESC_MONEDA
                rsStruc.Update
            End If
        End If
        rs.MoveNext
    Loop
    
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(rsStruc, Replace(UCase(strRpt & "Estructura\PP_Rpt_Cert_Pend_Reliq.rpt"), ".RPT", ".TTX"), 1)
    
    If objRep.CargaReporte(strRpt & "", "PP_Rpt_Cert_Pend_Reliq.rpt", "Informe Certificados Pendientes de reliquidación", rsStruc, True, _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("NombreSistema", vgNombreSistema), _
                            ArrFormulas("NombreSubSistema", vgNombreSubSistema)) = False Then
                            
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
    End If
    
    Exit Sub
mierror:
    MsgBox "No se pudo cargar el reporte", vbInformation
    
End Sub
Private Sub CREAR_ESTRUCTURA()

Set rsStruc = New ADODB.Recordset
With rsStruc.Fields
    .Append "MONEDA", adVarChar, 20
    .Append "AFP", adVarChar, 20
    .Append "COD_PENSIÓN", adVarChar, 20
    .Append "NUM_POLIZA", adVarChar, 20
    .Append "GLS_NOMBEN", adVarChar, 50
    .Append "GLS_PATBEN", adVarChar, 50
    .Append "GLS_MATBEN", adVarChar, 50
    .Append "CUOTAS_RETENIDAS", adVarChar, 20
    .Append "FECHA_RELIQ", adVarChar, 50
    .Append "MTO_PENSION", adDouble
    .Append "PENSIÓN", adDouble
    .Append "RETENSIÓN", adDouble
    .Append "ESSALUD", adDouble
    .Append "PERIODO_DE_PAGO", adVarChar, 20
    .Append "DESC_MONEDA", adVarChar, 30
    .Append "EDAD", adDouble
    .Append "COD_PAR", adVarChar, 2
    
End With
rsStruc.Open
End Sub


Private Sub cmdImprimirPend_Click()
Dim objRep As New ClsReporte

    On Error GoTo mierror
   
    'Call Crea_Estructura
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

    rs.Open "PP_LISTA_CERT_PEND.LISTAR('" & Format(dtpInicio, "yyyyMMdd") & "','" & Format(dtpFinal, "yyyyMMdd") & "')", vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF Then
        MsgBox "No hay datos para mostrar", vbExclamation, "Pago de pensiones"
        Exit Sub
    End If
    
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PP_Rpt_Cert_PendSinCert.rpt"), ".RPT", ".TTX"), 1)
    
        
    If objRep.CargaReporte(strRpt & "", "PP_Rpt_Cert_PendSinCert.rpt", "Informe de Pólizas sin certificados o vencidos", rs, True, _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("NombreSistema", vgNombreSistema), _
                            ArrFormulas("NombreSubSistema", vgNombreSubSistema)) = False Then
                            
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
    End If
    
    Exit Sub
mierror:
    MsgBox "No se pudo cargar el reporte", vbInformation
    
End Sub

Private Sub cmdReliquidar_Click()

Dim X, i As Integer
Dim vlOrden As Integer
Dim vlSql, vlSql2 As String
Dim vlSqlCompl As String
Dim iNumReliq As Integer
Dim fechaDesde, fechahasta As String
Dim marca As Boolean

On Error GoTo mierror

Set rsdet = New ADODB.Recordset

If Not fgConexionBaseDatos(vgConexionTransac) Then
    MsgBox "Error en Conexion a la Base de Datos", vbCritical, Me.Caption
    Exit Sub
End If

If MSHLista.Rows = 1 And optPendiente.Value = True Then
    MsgBox "No hay datos a reliquidar", vbCritical, Me.Caption
    Exit Sub
End If

If SSTab1.Tab = 1 Then
    Exit Sub
End If

For X = 1 To MSHLista.Rows - 1
    If MSHLista.TextMatrix(X, 0) <> "" Then
        marca = True
    End If
Next X

If marca = False Then
    MsgBox "No ha seleccionado registros a reliquidar", vbInformation, "Pensiones"
    Exit Sub
End If
    

Frame3.Visible = True
DoEvents

Screen.MousePointer = 11
pbAbanse.Max = MSHLista.Rows
pbAbanse.Min = 0
pbAbanse.Value = 0

    
 
    For X = 1 To MSHLista.Rows - 1
        pbAbanse.Value = pbAbanse.Value + 1
        DoEvents
        If MSHLista.TextMatrix(X, 0) <> "" Then
            'obtiene datos de beneficiarios por póliza
            Set rsdet = New ADODB.Recordset
            rsdet.CursorLocation = adUseClient
            rsdet.Open "PP_LISTA_DATOS_BENE.LISTAR('" & MSHLista.TextMatrix(MSHLista.Row, 1) & "','" & MSHLista.TextMatrix(MSHLista.Row, 3) & "')", vgConexionBD, adOpenStatic, adLockReadOnly
        
            'OBTIENE EL NUMERO MAXIMO DE RELIQUIDACION
            iNumReliq = flObtieneNumReliq
            DoEvents
            
            'carga detalle
            msfDetalle.Rows = 1
            Call CargaDetalle(MSHLista.TextMatrix(X, 1), MSHLista.TextMatrix(X, 3))
            
            DoEvents
            
            vgConexionTransac.BeginTrans
            
            For i = 1 To msfDetalle.Rows - 1
                If i = 1 Then
                    fechaDesde = Mid(msfDetalle.TextMatrix(i, 2), 4, 4) & Mid(msfDetalle.TextMatrix(i, 2), 1, 2)
                Else
                    fechahasta = Mid(msfDetalle.TextMatrix(i, 2), 4, 4) & Mid(msfDetalle.TextMatrix(i, 2), 1, 2)
                End If
            Next i
            
            'Graba la Reliquidación
            vlSql = "INSERT INTO pp_tmae_reliq (num_reliq, num_poliza, num_endoso, fec_reliq,num_perdesde, num_perhasta, gls_observacion)"
            vlSql = vlSql & " VALUES (" & iNumReliq & ",'" & MSHLista.TextMatrix(X, 1) & "'," & MSHLista.TextMatrix(X, 2) & ",TO_CHAR(SYSDATE,'YYYYMMDD')" & ",'"
            vlSql = vlSql & fechaDesde & "','" & fechahasta & "','')"
            vgConexionTransac.Execute (vlSql)
            
            'Graba Beneficiarios de la Reliquidacion
            rsdet.MoveFirst
            vlSql = "INSERT INTO pp_tmae_benreliq(num_reliq, num_poliza, num_endoso, num_orden,cod_indreliq, num_perdesde, num_perhasta,cod_indpension)"
            vlSql = vlSql & " VALUES (" & iNumReliq & ",'" & MSHLista.TextMatrix(X, 1) & "'," & MSHLista.TextMatrix(X, 2) & ","
            
            
            If Not rsdet.EOF Then
                Do While Not rsdet.EOF
                    vlOrden = MSHLista.TextMatrix(X, 3)
                    vlSqlCompl = vlOrden & ",1,'" & fechaDesde & "','" & fechahasta & "',1)"
                    vlSql2 = vlSql & vlSqlCompl
                    vgConexionTransac.Execute (vlSql2)
                    rsdet.MoveNext
                Loop
            End If
            
            'Graba Beneficiarios reliquidados
            vlSql = "INSERT INTO pp_tmae_detcalcreliq"
            vlSql = vlSql & " (num_reliq, num_orden, num_perpago,cod_conhabdes, fec_inipago, fec_terpago,mto_conhabdes, cod_moneda, mto_conhabdesant,mto_diferencia)"
            vlSql = vlSql & " VALUES (" & iNumReliq & ","
            For i = 1 To msfDetalle.Rows - 1
               vlSqlCompl = MSHLista.TextMatrix(X, 3) & ",'" & Mid(msfDetalle.TextMatrix(i, 2), 4, 4) & Mid(msfDetalle.TextMatrix(i, 2), 1, 2) & "','" & Trim(Mid(msfDetalle.TextMatrix(i, 3), 1, InStr(1, msfDetalle.TextMatrix(i, 3), "-") - 1)) & "',"
               vlSqlCompl = vlSqlCompl & "'" & Format(msfDetalle.TextMatrix(i, 4), "YYYYMMDD") & "','" & Format(msfDetalle.TextMatrix(i, 5), "YYYYMMDD") & "'," & Str(msfDetalle.TextMatrix(i, 7)) & ", "
               vlSqlCompl = vlSqlCompl & "'" & Trim(msfDetalle.TextMatrix(i, 9)) & "'," & Str(msfDetalle.TextMatrix(i, 6)) & "," & Str(msfDetalle.TextMatrix(i, 8)) & ")"
               vlSql2 = vlSql & vlSqlCompl
               vgConexionTransac.Execute (vlSql2)
            Next i
            
            'Graba Modalidad de Pago
            vlOrden = 0
            rsdet.MoveFirst
            vlSql = "INSERT INTO pp_tmae_detpagoreliq(num_reliq, num_orden, cod_conhabdes,fec_inihabdes, fec_terhabdes, num_cuotas,mto_cuota, mto_total, cod_moneda, mto_ultcuota)"
            vlSql = vlSql & " VALUES (" & iNumReliq & ","
            Do While Not rsdet.EOF
                vlOrden = MSHLista.TextMatrix(X, 3)
                vlSqlCompl = vlOrden & "," & Trim(Mid(vlConceptoPEMayor, 1, InStr(1, vlConceptoPEMayor, "-") - 1)) & ",'" & Format("01/" & Format(Month(dtpFinal.Value), "00") & "/" & Year(dtpFinal.Value), "YYYYMMDD") & "','" & Format(dtpFinal.Value, "YYYYMMDD") & "',"
                vlSqlCompl = vlSqlCompl & "1," & txtCuotas.Text & "," & CDbl(txtMontoTot.Text) & ",'" & msfDetalle.TextMatrix(1, 9) & "',0)"
                vlSql2 = vlSql & vlSqlCompl
                vgConexionTransac.Execute (vlSql2)
                rsdet.MoveNext
            Loop
            
            'Graba Haber y Descuento en Tabla General de Haberes y Descuentos
            vlOrden = 0
            rsdet.MoveFirst
            vlSql = "INSERT INTO pp_tmae_habdes(num_poliza, num_endoso, num_orden,cod_conhabdes, fec_inihabdes,fec_terhabdes, num_cuotas, mto_cuota, "
            vlSql = vlSql & " mto_total, cod_moneda, cod_motsushabdes, fec_sushabdes,gls_obshabdes, cod_usuariocrea, fec_crea,hor_crea, num_reliq) VALUES ('"
            vlSql = vlSql & MSHLista.TextMatrix(X, 1) & "'," & MSHLista.TextMatrix(X, 2) & ","
            
            Do While Not rsdet.EOF
                vlOrden = MSHLista.TextMatrix(X, 3)
                vlSqlCompl = vlOrden & ",'" & Trim(Mid(vlConceptoPEMayor, 1, InStr(1, vlConceptoPEMayor, "-") - 1)) & "',"
                vlSqlCompl = vlSqlCompl & "'" & Format("01/" & Format(Month(dtpFinal.Value), "00") & "/" & Year(dtpFinal.Value), "YYYYMMDD") & "','" & Format(dtpFinal.Value, "YYYYMMDD") & "',1," & CDbl(txtMontoTot.Text) & "," & CDbl(txtMontoTot.Text) & ","
                vlSqlCompl = vlSqlCompl & "'" & msfDetalle.TextMatrix(1, 9) & "','00', NULL, NULL,'" & vgUsuario & "','" & Format(Date, "yyyymmdd") & "','" & Format(Time, "hhmmss") & "'," & iNumReliq & ")"
                vlSql2 = vlSql & vlSqlCompl
                vgConexionTransac.Execute (vlSql2)
                rsdet.MoveNext
            Loop
            
            'Graba Haber y Descuento en Tabla General de Haberes y Descuentos el CONCEPTO 06 RETROACTIVA
            vlOrden = 0
            rsdet.MoveFirst
            vlSql = "INSERT INTO pp_tmae_habdes(num_poliza, num_endoso, num_orden,cod_conhabdes, fec_inihabdes,fec_terhabdes, num_cuotas, mto_cuota, "
            vlSql = vlSql & " mto_total, cod_moneda, cod_motsushabdes, fec_sushabdes,gls_obshabdes, cod_usuariocrea, fec_crea,hor_crea, num_reliq) VALUES ('"
            vlSql = vlSql & MSHLista.TextMatrix(X, 1) & "'," & MSHLista.TextMatrix(X, 2) & ","
            
            Do While Not rsdet.EOF
                vlOrden = MSHLista.TextMatrix(X, 3)
                vlSqlCompl = vlOrden & ",'" & Trim(Mid(vlConceptoPRentaRetro, 1, InStr(1, vlConceptoPRentaRetro, "-") - 1)) & "',"
                vlSqlCompl = vlSqlCompl & "'" & Format("01/" & Format(Month(dtpFinal.Value), "00") & "/" & Year(dtpFinal.Value), "YYYYMMDD") & "','" & Format(dtpFinal.Value, "YYYYMMDD") & "',1," & CDbl(txtMontoTot.Text) & "," & CDbl(txtMontoTot.Text) & ","
                vlSqlCompl = vlSqlCompl & "'" & msfDetalle.TextMatrix(1, 9) & "','00', NULL, NULL,'" & vgUsuario & "','" & Format(Date, "yyyymmdd") & "','" & Format(Time, "hhmmss") & "'," & iNumReliq & ")"
                vlSql2 = vlSql & vlSqlCompl
                vgConexionTransac.Execute (vlSql2)
                rsdet.MoveNext
            Loop
            
            'Actualiza Tabla por la cual se hizo la Reliquidación
            'con la query traspasada desde el Formulario que Llama a la Reliquidación
            
            vlSql = "update pp_tmae_certificado SET num_reliq = " & iNumReliq
            vlSql = vlSql & " where num_poliza='" & MSHLista.TextMatrix(X, 1) & "' and  num_endoso =(select max(num_endoso) from pp_tmae_certificado where num_poliza='" & MSHLista.TextMatrix(X, 1) & "' and num_orden='" & MSHLista.TextMatrix(X, 3) & "') AND num_orden='" & MSHLista.TextMatrix(X, 3) & "'"
            vgConexionTransac.Execute (vlSql)
            
            vgConexionTransac.CommitTrans
            
        End If
    Next X
    


Frame3.Visible = False
vgConexionTransac.Close
Set vgConexionTransac = Nothing

MsgBox "Proceso de Reliquidación finalizo con éxito", vbInformation

Call CargaDatos
Screen.MousePointer = 1

Exit Sub
mierror:
    vgConexionTransac.RollbackTrans
    MsgBox "No se pudo reliquidar", vbExclamation
    Screen.MousePointer = 1
    Frame3.Visible = False
End Sub


Private Sub DtgCertVenc_HeadClick(ByVal ColIndex As Integer)
filtro = DtgCertVenc.Columns(ColIndex).Caption
RSGEN.Sort = filtro
End Sub

Private Sub Form_Load()

Dim fechMenor As Date
Dim fech As Date
C_sSeleccion = ">"
C_sTipLetraSel = "Monotype Sorts"
TipLetra = "VERDANA"

Me.Top = 0
Me.Left = 0


dtpFinal.Value = FechaPago 'DateAdd("d", -1, fech)
fechMenor = DateAdd("m", -1, dtpFinal)

fech = DateAdd("m", 1, fechMenor)
dtpInicio.Value = "01/" & Format(Month(fech), "00") & "/" & Year(fech)

SSTab1.Tab = 0
optPendiente.Value = True
Call CargaDatos
Call Init_Flex(MSHLista)


End Sub
Private Function FechaPago() As Date
Dim FechaFin As Date
Dim fechafin2 As Date
Dim rsf As ADODB.Recordset
Set rsf = New ADODB.Recordset
FechaFin = FechaServidor
rsf.Open "select to_date(max(fec_calpagoreg),'yyyy/mm/dd') as fecha from PP_TMAE_PROPAGOPEN where cod_estadoreg='C'", vgConexionBD, adOpenStatic, adLockReadOnly
If Not rsf.EOF Then
    fechafin2 = Format(rsf!fecha, "dd/mm/yyyy")
    FechaPago = DateAdd("d", -1, fechafin2)
Else
    MsgBox "No hay fecha programada en el calendario para este mes", vbExclamation, "Pago de pensiones"
    cmdBuscar.Enabled = False
    cmdReliquidar.Enabled = False
    cmdImprimir.Enabled = False
End If
End Function

Private Function FechaPagoAbierto() As Date
Dim FechaFin As Date
Dim fechafin2 As Date
Dim rsf As ADODB.Recordset
Set rsf = New ADODB.Recordset
FechaFin = FechaServidor
rsf.Open "select to_date(max(fec_calpagoreg),'yyyy/mm/dd')as fecha from PP_TMAE_PROPAGOPEN", vgConexionBD, adOpenStatic, adLockReadOnly
If Not rsf.EOF Then
    fechafin2 = Format(rsf!fecha, "dd/mm/yyyy")
    FechaPagoAbierto = DateAdd("d", -1, fechafin2)
Else
    MsgBox "No hay fecha programada en el calendario para este mes", vbExclamation, "Pago de pensiones"
    cmdBuscar.Enabled = False
    cmdReliquidar.Enabled = False
    cmdImprimir.Enabled = False
End If
End Function


Private Sub CargaDatos()


On Error GoTo mierror

Set RSGEN = New ADODB.Recordset
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient

chkTodos.Value = 0
Call LmpGrilla
Call LmpGrillaDetalle

If optPendiente.Value = True Then
    rs.Open "PP_LISTA_RELIQUIDACION.LISTAR('" & Format(dtpInicio, "yyyyMMdd") & "','" & Format(dtpFinal, "yyyyMMdd") & "')", vgConexionBD, adOpenStatic, adLockReadOnly

    While Not rs.EOF
        If Not IsNull(rs.Fields("Cuotas_retenidas")) And rs.Fields("Cuotas_retenidas") > 0 Then
        
        'If (rs.Fields("Edad") < 18 And rs.Fields("cod_par") >= "30" And rs.Fields("cod_par") <= "35") Then
        MSHLista.AddItem (vbTab & (rs.Fields("num_poliza")) & vbTab & _
                             (rs.Fields("num_endoso")) & vbTab & _
                             (rs.Fields("num_orden")) & vbTab & _
                             (rs.Fields("cod_tipoidenben")) & vbTab & _
                             (rs.Fields("num_idenben")) & vbTab & _
                             (rs.Fields("gls_nomben")) & vbTab & _
                             (rs.Fields("gls_patben")) & vbTab & _
                             (rs.Fields("gls_matben")) & vbTab & _
                             (Format(rs.Fields("mto_pension"), "###,##0.00")))
                             
        'End If
        'If (rs.Fields("Edad") > 18 And rs.Fields("cod_par") >= "10") Then
        'MSHLista.AddItem (vbTab & (rs.Fields("num_poliza")) & vbTab & _
        '                     (rs.Fields("num_endoso")) & vbTab & _
        '                     (rs.Fields("num_orden")) & vbTab & _
        '                     (rs.Fields("cod_tipoidenben")) & vbTab & _
        '                     (rs.Fields("num_idenben")) & vbTab & _
        '                     (rs.Fields("gls_nomben")) & vbTab & _
        '                     (rs.Fields("gls_patben")) & vbTab & _
        '                     (rs.Fields("gls_matben")) & vbTab & _
        '                     (Format(rs.Fields("mto_pension"), "###,##0.00")))
        'End If
            
                             ' & vbTab & _
                             '(rs.Fields("Cuotas_retenidas")) & vbTab & _
                             '(Format((rs.Fields("mto_pension") * rs.Fields("Cuotas_retenidas")), "###,##0.00")))
        End If
        rs.MoveNext
    Wend
Else
    rs.Open "PP_LISTA_RELIQUIDADOS.LISTAR('" & Format(dtpInicio, "yyyyMMdd") & "','" & Format(dtpFinal, "yyyyMMdd") & "')", vgConexionBD, adOpenStatic, adLockReadOnly

    While Not rs.EOF
    'If rs.Fields("Cuotas_retenidas") = 0 Then
        MSHLista.AddItem (vbTab & (rs.Fields("num_poliza")) & vbTab & _
                         (rs.Fields("num_endoso")) & vbTab & _
                         (rs.Fields("num_orden")) & vbTab & _
                         (rs.Fields("cod_tipoidenben")) & vbTab & _
                         (rs.Fields("num_idenben")) & vbTab & _
                         (rs.Fields("gls_nomben")) & vbTab & _
                         (rs.Fields("gls_patben")) & vbTab & _
                         (rs.Fields("gls_matben")) & vbTab & _
                         (Format(rs.Fields("mto_pension"), "###,##0.00")))
                         '& vbTab & _
'                         (0) & vbTab & _
'                         (Format((rs.Fields("mto_pension") * 0), "###,##0.00")))
    'End If
    rs.MoveNext
Wend
End If

If Not rs.EOF Then
    If optPendiente.Value = True Then
        lblTitulo1.Caption = "Listado de Pólizas Pendientes"
    Else
        lblTitulo1.Caption = "Listado de Pólizas reliquidadas"
    End If
    MSHLista.Row = 1
End If

Set RSGEN = New ADODB.Recordset
RSGEN.CursorLocation = adUseClient
RSGEN.Open "PP_LISTA_CERT_PEND.LISTAR('" & Format(dtpInicio, "yyyyMMdd") & "','" & Format(dtpFinal, "yyyyMMdd") & "')", vgConexionBD, adOpenStatic, adLockReadOnly

If Not RSGEN.EOF Then
    Set DtgCertVenc.DataSource = RSGEN
    DtgCertVenc.Refresh
Else
    Exit Sub
End If

Call FormateaDatagrid

Exit Sub
mierror:
    MsgBox "No se pudo cargar el listado", vbInformation
    
End Sub
Private Sub FormateaDatagrid()
    With DtgCertVenc
        .Columns(0).Alignment = dbgCenter
        .Columns(0).Width = "1140.095"
        
        .Columns(1).Alignment = dbgCenter
        .Columns(1).Width = "1305.071"
        
        .Columns(2).Alignment = dbgCenter
        .Columns(2).Width = "1244.976"
        
        .Columns(3).Alignment = dbgLeft
        .Columns(3).Width = "4305.26"
        
        .Columns(4).Alignment = dbgCenter
        .Columns(4).Width = "959.8111"
        
        .Columns(5).Alignment = dbgCenter
        .Columns(5).Width = "1305.071"
        
        .Columns(6).Alignment = dbgCenter
        .Columns(6).Width = "1530.142"
        
        .Columns(7).Alignment = dbgCenter
        .Columns(7).Width = "854.9292"
        
        .Columns(8).Alignment = dbgCenter
        .Columns(8).Width = "1305.071"
        
        .Columns(9).Alignment = dbgRight
        .Columns(9).Width = "1305.071"
        .Columns(9).NumberFormat = "###,##0.00"
        
        .Columns(10).Alignment = dbgRight
        .Columns(10).Width = "929.7639"
        .Columns(10).NumberFormat = "###,##0.00"
        
        .Columns(11).Alignment = dbgCenter
        .Columns(11).Width = "1725.165"
        
        .Columns(12).Alignment = dbgRight
        .Columns(12).Width = "1305.071"
        .Columns(12).NumberFormat = "###,##0.00"
        
        .Columns(13).Alignment = dbgCenter
        .Columns(13).Width = "1184.882"
        
        .Columns(14).Alignment = dbgCenter
        .Columns(14).Width = "1739.906"
        
        .Columns(15).Alignment = dbgCenter
        .Columns(15).Width = "1260.284"
        
    End With
End Sub

Function LmpGrilla()

    MSHLista.Clear
    MSHLista.Rows = 1
    MSHLista.RowHeight(0) = 250
    MSHLista.Row = 0
    MSHLista.Cols = 10

    MSHLista.Col = 0
    MSHLista.Text = "Sel"
    MSHLista.ColWidth(0) = 350
    MSHLista.ColAlignment(0) = 1
    MSHLista.BackColor = &H8000000F
    
    MSHLista.Col = 1
    MSHLista.Text = "Póliza"
    MSHLista.ColWidth(1) = 1000
    MSHLista.ColAlignment(1) = 1

    MSHLista.Col = 2
    MSHLista.Text = "Endoso"
    MSHLista.ColWidth(2) = 700
    MSHLista.ColAlignment(2) = 3

    MSHLista.Col = 3
    MSHLista.Text = "Orden"
    MSHLista.ColWidth(3) = 600
    MSHLista.ColAlignment(3) = 3

    MSHLista.Col = 4
    MSHLista.Text = "Tip.Doc"
    MSHLista.ColWidth(4) = 700
    MSHLista.ColAlignment(4) = 3
    
    MSHLista.Col = 5
    MSHLista.Text = "DNI"
    MSHLista.ColWidth(5) = 900
    MSHLista.ColAlignment(5) = 1
    
    MSHLista.Col = 6
    MSHLista.Text = "Nombre"
    MSHLista.ColWidth(6) = 2200

    MSHLista.Col = 7
    MSHLista.Text = "Apellido Paterno"
    MSHLista.ColWidth(7) = 1600

    MSHLista.Col = 8
    MSHLista.Text = "Apellido Materno"
    MSHLista.ColWidth(8) = 1600
    
    MSHLista.Col = 9
    MSHLista.Text = "Mto. Cuota sin Essalud"
    MSHLista.ColWidth(9) = 1700
    'MSHLista.FormatString = "###,##0.00"
    MSHLista.ColAlignment(9) = 6
    
'    MSHLista.Col = 10
'    MSHLista.Text = "nto 2%"
'    MSHLista.ColWidth(10) = 1700
'    MSHLista.ColAlignment(10) = 6
    
    
'    MSHLista.Col = 10
'    MSHLista.Text = "Ctas.Ret."
'    MSHLista.ColWidth(10) = 800
'    MSHLista.ColAlignment(10) = 3
'
'    MSHLista.Col = 11
'    MSHLista.Text = "Monto Total"
'    MSHLista.ColWidth(11) = 1200
'    'MSHLista.FormatString = "###,##0.00"
'    MSHLista.ColAlignment(11) = 6

End Function
    
    
Function LmpGrillaDetalle()

    msfDetalle.Clear
    msfDetalle.Rows = 1
    msfDetalle.RowHeight(0) = 250
    msfDetalle.Row = 0
    msfDetalle.Cols = 10

    msfDetalle.Col = 0
    msfDetalle.Text = ""
    msfDetalle.ColWidth(0) = 1
    msfDetalle.ColAlignment(0) = 1
    
    msfDetalle.Col = 1
    msfDetalle.Text = "N° Orden"
    msfDetalle.ColWidth(1) = 1000
    msfDetalle.ColAlignment(1) = 3

    msfDetalle.Col = 2
    msfDetalle.Text = "Periodo"
    msfDetalle.ColWidth(2) = 900
    msfDetalle.ColAlignment(2) = 3

    msfDetalle.Col = 3
    msfDetalle.Text = "Concepto"
    msfDetalle.ColWidth(3) = 2500
    msfDetalle.ColAlignment(3) = 3

    msfDetalle.Col = 4
    msfDetalle.Text = "Fecha Inicio"
    msfDetalle.ColWidth(4) = 1100
    msfDetalle.ColAlignment(4) = 3
    
    msfDetalle.Col = 5
    msfDetalle.Text = "Fecha Final"
    msfDetalle.ColWidth(5) = 1100
    msfDetalle.ColAlignment(5) = 3
    
    msfDetalle.Col = 6
    msfDetalle.Text = "Monto Ant."
    msfDetalle.ColWidth(6) = 1200
    msfDetalle.ColAlignment(6) = 6

    msfDetalle.Col = 7
    msfDetalle.Text = "Monto Act."
    msfDetalle.ColWidth(7) = 1200
    msfDetalle.ColAlignment(7) = 6

    msfDetalle.Col = 8
    msfDetalle.Text = "Diferencia"
    msfDetalle.ColWidth(8) = 1000
    msfDetalle.ColAlignment(8) = 6
    
    msfDetalle.Col = 9
    msfDetalle.Text = "Moneda"
    msfDetalle.ColWidth(9) = 800
    msfDetalle.ColAlignment(9) = 3

End Function

Public Sub Init_Flex(EsteFlex As MSFlexGrid)
 With EsteFlex
'  .FocusRect = flexFocusNone
'  .HighLight = flexHighlightWithFocus
'  .Gridlines = flexGridNone
'  .SelectionMode = flexSelectionByRow
'  .ScrollTrack = True
'  .Font = C_sTipLetraSel 'TipLetra
 End With
End Sub

Private Sub MSHLista_Click()
Dim i As Integer

If MSHLista.Col = 0 Then
    If MSHLista.TextMatrix(MSHLista.Row, 0) = C_sSeleccion Then
     MSHLista.TextMatrix(MSHLista.Row, 0) = ""
    Else
      MSHLista.Col = 0
      MSHLista.CellAlignment = 4
      MSHLista.CellFontName = C_sTipLetraSel
      MSHLista.CellForeColor = &HFF&
      MSHLista.TextMatrix(MSHLista.Row, 0) = C_sSeleccion
    End If
    MSHLista.Col = 1
    MSHLista.ColSel = MSHLista.Cols - 1
End If

End Sub

Private Sub MSHLista_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeySpace Then
  MSHLista_Click
 End If
End Sub

Private Sub CargaDetalle(ByVal Poliza As String, ByVal Orden As Long)
Dim vlPerDesde As Date, vlPerHasta As Date
Dim vlOrden As Long, vlTotAsigFam As Double 'Monto Total de Asignación Familiar
Dim vlFechaHasta As Date, vlConcepto As String, vlMoneda As String
Dim vlFila As Long
Dim vlMontoAnterior As Double, vlDiferencia As Double
Dim i As Integer
Dim vlVarCarga As Double, vlNumCargas As Integer
Dim vlAsigFamiliar As Double
Dim vlCalcPension As Boolean
Dim vlTipoPension As String
Dim vlIndPension As Boolean, vlIndGarantia As Boolean  'Indica si se debe Pagar Pensión o Garantía Estatal
Dim vlFecTerPagoPenGar As String, vlPensionGar As Double
Dim vlPension As Double


Dim vlPensionNormal As Double
Dim vlSql As String, vlTB As ADODB.Recordset
Dim vlCodPar As String, vlFecHasta As Date
Dim vlEdad As Integer, vlEdadAños As Integer, bResp As Integer
Dim vlFecIniPago As String, vlTB2 As ADODB.Recordset
                    
Dim rsdet As New ADODB.Recordset
Dim rsdet2 As New ADODB.Recordset
Dim CantCtas As Integer
Dim MtoCtas As Double

On Error GoTo mierror
    
If optReliquidados.Value = True Then
    Exit Sub
End If
msfDetalle.Rows = 1
If MSHLista.Rows - 1 > 0 Then
'    Set rsdet = New ADODB.Recordset
    
    vlConceptoPEMenor = fgObtieneDescripcionConcepto(stDatGenerales.Cod_ConceptoPensionCobro)
    vlConceptoPEMayor = fgObtieneDescripcionConcepto(stDatGenerales.Cod_ConceptoPensionPago)
    vlConceptoPRentaRetro = fgObtieneDescripcionConcepto(stDatGenerales.Cod_ConceptoPensionPagoRetro)
    
    rsdet.CursorLocation = adUseClient
    rsdet.Open "PP_LISTA_DATOS_BENE.LISTAR('" & Poliza & "','" & Orden & "')", vgConexionBD, adOpenStatic, adLockReadOnly
    
    If Not rsdet.EOF Then
    Do While Not rsdet.EOF
        CantCtas = 0
        MtoCtas = 0
    
         vlTipoPension = rsdet!Cod_TipPension
         vlPensionNormal = rsdet!Mto_Pension
         vlOrden = vlOrden + 1
         'vlFecIniPago = DateSerial(Mid(rsdet!Fec_IniPagoPen, 1, 4), Mid(rsdet!Fec_IniPagoPen, 5, 2), Mid(rsdet!Fec_IniPagoPen, 7, 2))
         
         
        'verifico desde ultimo mes de pago+1 del bene... si no tiene desde el ultimo mes de pago+1 del titular
        If Orden = 1 Then
            vlFecIniPago = DateSerial(Mid(rsdet!Fec_IniPagoPen, 1, 4), Mid(rsdet!Fec_IniPagoPen, 5, 2), Mid(rsdet!Fec_IniPagoPen, 7, 2))
            vlPerDesde = Mid(rsdet!Fec_IniPagoPen, 7, 2) & "/" & Mid(rsdet!Fec_IniPagoPen, 5, 2) & "/" & Mid(rsdet!Fec_IniPagoPen, 1, 4)
        Else
'            Dim rs As ADODB.Recordset
'            Set rs = New ADODB.Recordset
'            rs.Open "select max(num_perpago) as PagoUltimo from pp_tmae_pagopendef where num_poliza='" & Poliza & "' and cod_conhabdes='01' " & _
'                    " and num_orden=" & Orden & " and cod_tipreceptor <> 'R' order by num_perpago", vgConexionBD, adOpenStatic, adLockReadOnly
'
'            If Not rs.EOF Then
'                If IsNull(rs!PagoUltimo) = True Then
                        Dim rs2 As ADODB.Recordset
                        Set rs2 = New ADODB.Recordset
                        rs2.Open "select max(num_perpago) + 1 as PagoUltimo from pp_tmae_pagopendef where num_poliza='" & Poliza & "' and cod_conhabdes='01' " & _
                                " and num_orden=" & Orden & " and cod_tipreceptor <> 'R' order by num_perpago", vgConexionBD, adOpenStatic, adLockReadOnly
                            
                        If Not rs2.EOF Then
                            If IsNull(rs2!PagoUltimo) = True Then
                                vlFecIniPago = DateSerial(Mid(rsdet!Fec_IniPagoPen, 1, 4), Mid(rsdet!Fec_IniPagoPen, 5, 2), Mid(rsdet!Fec_IniPagoPen, 7, 2))
                                vlPerDesde = Mid(rsdet!Fec_IniPagoPen, 7, 2) & "/" & Mid(rsdet!Fec_IniPagoPen, 5, 2) & "/" & Mid(rsdet!Fec_IniPagoPen, 1, 4)
                            Else
                                vlFecIniPago = DateSerial(Mid(rs2!PagoUltimo, 1, 4), Mid(rs2!PagoUltimo, 5, 2), "01")
                                vlPerDesde = "01/" & Mid(rs2!PagoUltimo, 5, 2) & "/" & Mid(rs2!PagoUltimo, 1, 4)
                            End If
                        End If
'                Else
'                    vlFecIniPago = DateSerial(Mid(rs!PagoUltimo, 1, 4), Mid(rs!PagoUltimo, 5, 2), "01")
'                    vlPerDesde = "01/" & Mid(rs!PagoUltimo, 5, 2) & "/" & Mid(rs!PagoUltimo, 1, 4)
'                End If
'
'            End If
        End If
'            vlFecIniPago = DateSerial(Mid(rsdet!Fec_IniPagoPen, 1, 4), Mid(rsdet!Fec_IniPagoPen, 5, 2), Mid(rsdet!Fec_IniPagoPen, 7, 2))
'            vlPerDesde = Mid(rsdet!Fec_IniPagoPen, 7, 2) & "/" & Mid(rsdet!Fec_IniPagoPen, 5, 2) & "/" & Mid(rsdet!Fec_IniPagoPen, 1, 4)

         vlCodPar = rsdet!Cod_Par
         If Not IsNull(rsdet!Fec_TerPagoPenGar) Then
             vlFecTerPagoPenGar = rsdet!Fec_TerPagoPenGar
             vlPensionGar = rsdet!Mto_Pension
         End If
         
         vlPerHasta = "01/" & Format(Month(dtpFinal.Value), "00") & "/" & Year(dtpFinal.Value)
    
            
            Do While vlPerDesde <= vlPerHasta
            
            'If vlPerDesde = "01/12/2011" Then
            'MsgBox "dfrd"
            'End If
                If ValidaReliquidacion(Poliza, Year(vlPerDesde) & Format(Month(vlPerDesde), "00"), Orden) Then
                    GoTo Siguiente
                End If
                If CDate(vlFecIniPago) > vlPerDesde Then
                    GoTo Siguiente 'Aun no le corresponde pension
                End If
                
                
                vlSql = " select to_date(x.num_perpago||'01', 'yyyy/mm/dd' ) as pagofec, sum(mto_conhabdes) as monto from pp_tmae_pagopendef x"
                vlSql = vlSql & " left join pp_tmae_ben c on x.num_poliza=c.num_poliza and x.num_orden=c.num_orden"
                vlSql = vlSql & " where x.num_poliza = '" & Poliza & "' and"
                vlSql = vlSql & " x.cod_tipreceptor <> 'R' AND"
                vlSql = vlSql & " x.cod_conhabdes = '01' AND"
                vlSql = vlSql & " c.num_orden = '" & Orden & "' and"
                vlSql = vlSql & " c.num_endoso= " & MSHLista.TextMatrix(MSHLista.Row, 2) & " and"
                vlSql = vlSql & " SUBSTR(x.fec_inipago,1,6)=SUBSTR('" & Format(vlPerDesde, "yyyymmdd") & "',1,6)"
                vlSql = vlSql & " group by x.num_perpago"
                vlSql = vlSql & " order by 1"
                Set rsdet2 = vgConexionBD.Execute(vlSql)
                
                If Not rsdet2.EOF Then
                        If rsdet2!pagofec = vlPerDesde Then
                            'MsgBox "fECHA EXISTE" & rsdet2!pagofec
                        End If
                Else
                        'If rsdet!Cod_Moneda = "NS" Then
                        If rsdet!Cod_TipReajuste <> cgSINAJUSTE Then 'hqr 21/02/2011
                            'Obtiene Monto de la Pensión Actualizada
                            'vlSql = "SELECT a.mto_pension "
                            'vlSql = vlSql & "FROM pp_tmae_pensionact a "
                            'vlSql = vlSql & "WHERE a.num_poliza = '" & Poliza & "' "
                            'vlSql = vlSql & "AND a.num_endoso = " & MSHLista.TextMatrix(MSHLista.Row, 2) & " "
                            'vlSql = vlSql & "AND a.fec_desde = ("
                            'vlSql = vlSql & "SELECT max(fec_desde) FROM pp_tmae_pensionact b "
                            'vlSql = vlSql & "WHERE b.num_poliza = a.num_poliza "
                            'vlSql = vlSql & "AND b.num_endoso = a.num_endoso "
                            'vlSql = vlSql & "AND b.fec_desde <= '" & Format(vlPerDesde, "yyyymmdd") & "'"
                            'vlSql = vlSql & ")"
                            
                             vlSql = "select  PP_FUNCION_AJUSTE_PENSION ('" & Format(vlPerDesde, "yyyymmdd") & "', '" & Poliza & "') as Mto_Pension from dual"
                            
                            Set rsdet2 = vgConexionBD.Execute(vlSql)
                            If Not rsdet2.EOF Then
                                vlPensionNormal = rsdet2!Mto_Pension '/ ((rsdet!Mto_PlanSalud / 100) + 1)
                            Else
                                vlPensionNormal = 0
                            End If
                         End If
                
                 vlPension = Format(vlPensionNormal * rsdet!Prc_Pension / 100, "#0.00")
                'vlPension = Format(vlPensionNormal * rsdet!Mto_Pension / 100, "#0.00")
                 If vlFecTerPagoPenGar <> "" Then 'Hay fecha de Pago Garantizado
                     If vlFecTerPagoPenGar >= Format(vlPerDesde, "yyyymmdd") Then
                         vlPension = Format(vlPensionNormal * rsdet!Prc_PensionGar / 100, "#0.00")
                     End If
                 End If
                
                 vlFecHasta = DateAdd("d", -1, DateAdd("m", 1, vlPerDesde))
                 bResp = fgCalculaEdad(rsdet!Fec_NacBen, vlFecHasta)
                 If bResp = "-1" Then 'Error
                     Exit Sub
                 End If
                 vlEdad = bResp
                 vlEdadAños = fgConvierteEdadAños(vlEdad)
                
                 'Si son Hijos se Calcula la Edad y se Verifica Certificado de Estudios
                 If vlCodPar >= 30 And vlCodPar <= 35 Then 'Hijos
                     If vlEdad >= stDatGenerales.MesesEdad18 And rsdet!Cod_SitInv = "N" Then 'Hijos Sanos
                        'OBS: Se asume que el mes de los 18 años se paga completo
                            GoTo Siguiente
                        End If
                 End If
                 
                 'If vlIndPension Then  SE COMENTA POR Q NO TIENE EL CHECK DE LA ANTIGUA GRILLA
                    'Obtener Pensión Anterior
                    
                 
                        vlMontoAnterior = fgObtieneMontoConcepto(Poliza, Orden, vlOrden, stDatGenerales.Cod_ConceptoPension, Format(vlPerDesde, "yyyymm"), "PE") 'Estaría en Pesos
                        If vlMontoAnterior > 0 Then
                             vlMontoAnterior = ObtieneMontoEndoso(Poliza, Orden)
                        End If
                   
                    
                    
                     'vlDiferencia = Format(vlMontoAnterior - vlPension, "#0.00")
                     vlDiferencia = Format(vlMontoAnterior - vlPension, "#0.00")
                     
                     If Abs(vlDiferencia) = vlPension Then vlDiferencia = 0
                     
                     'If vlDiferencia <> 0 Then
                         'variables en duro -- se deben cambiar
                         
                         
                         If vlDiferencia > 0 Then
                             vlConcepto = vlConceptoPEMenor
                         Else
                             vlConcepto = vlConceptoPEMayor
                         End If
                         vlDiferencia = Abs(vlDiferencia)
                         vlMoneda = rsdet!Cod_Moneda
                         
                          msfDetalle.AddItem (vbTab & vlOrden & vbTab & _
                          (Format(vlPerDesde, "mm/yyyy")) & vbTab & _
                          (vlConcepto) & vbTab & _
                          (vlPerDesde) & vbTab & _
                          (vlFecHasta) & vbTab & _
                          (vlMontoAnterior) & vbTab & _
                          (Format(vlPension - (vlPension * (rsdet!Mto_PlanSalud / 100)), "###,##0.00")) & vbTab & _
                          (Format(vlDiferencia - (vlPension * (rsdet!Mto_PlanSalud / 100)), "###,##0.00")) & vbTab & _
                          (vlMoneda) & vbTab & _
                          (MSHLista.TextMatrix(MSHLista.Row, 3)))
                          
                          CantCtas = CantCtas + 1
                          MtoCtas = MtoCtas + Format(vlPension - (vlPension * (rsdet!Mto_PlanSalud / 100)), "###,##0.00")
                     
                      'End If
                    'End If
                
                    'MsgBox "fECHA no EXISTE" & vlPerDesde
                
                End If
            
                        
Siguiente:
                vlPerDesde = DateAdd("m", 1, vlPerDesde) 'Incrementa el Periodo
            Loop
            rsdet.MoveNext
        Loop
        
        txtCuotas.Text = CantCtas
        txtMontoTot.Text = Format(MtoCtas, "###,##0.00")
    End If
End If

Exit Sub
mierror:
    MsgBox "No pudo cargar detalle", vbInformation
    
End Sub

Private Function ObtieneMontoEndoso(ByVal Poliza As String, ByVal Orden As Integer) As Double

Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

On Error GoTo mierror
    
    rs.Open "select max(num_endoso),num_orden,(case when prc_pensiongar>0 then mto_pensiongar else mto_pension end)as mto_pension from pp_tmae_ben where num_poliza='" & Poliza & "' and num_orden=" & Orden & "" & _
            " AND NUM_ENDOSO=(SELECT MAX(NUM_ENDOSO) FROM pp_tmae_certificado where num_poliza='" & Poliza & "') group by num_orden,mto_pension,prc_pensiongar,mto_pensiongar", vgConexionBD, adOpenStatic, adLockReadOnly
            
    If Not rs.EOF Then
        If IsNull(rs!Mto_Pension) Then
             ObtieneMontoEndoso = 0
        Else
             ObtieneMontoEndoso = CDbl(rs!Mto_Pension)
        End If
    Else
         ObtieneMontoEndoso = 0
    End If

Exit Function
mierror:
MsgBox "No se pudo Obtener el monto de pago del endoso", vbInformation, "Pensiones"


End Function


Private Sub CargaPagos()
Dim Orden As Integer

On Error GoTo mierror

Do While Not rsdet.EOF
    Orden = Orden + 1
    
    rsdet.MoveNext
Loop

Exit Sub
mierror:
    MsgBox "No pudo cargar detalle", vbInformation
    
End Sub

Private Function ValidaReliquidacion(ByVal Poliza As String, ByVal DESDE As String, ByVal Orden As Integer) As Boolean
Dim rsreliq As ADODB.Recordset
Dim cadena As String

Set rsreliq = New ADODB.Recordset
cadena = "select NUM_PERPAGO from pp_tmae_reliq R INNER JOIN pp_tmae_detcalcreliq D ON R.NUM_RELIQ=D.NUM_RELIQ WHERE R.NUM_POLIZA='" & Poliza & "' AND D.NUM_PERPAGO='" & DESDE & "' and D.NUM_ORDEN=" & Orden & ""
rsreliq.Open cadena, vgConexionBD, adOpenStatic, adLockReadOnly
If Not rsreliq.EOF Then
    ValidaReliquidacion = True
End If
    
End Function


Private Sub MSHLista_SelChange()
If MSHLista.Col > 0 Then
Call CargaDetalle(MSHLista.TextMatrix(MSHLista.Row, 1), MSHLista.TextMatrix(MSHLista.Row, 3))
End If
End Sub


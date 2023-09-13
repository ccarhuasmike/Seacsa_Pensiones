VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmMemos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso Manual de pensiones"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   12045
   Begin VB.Frame fraListado 
      Height          =   6345
      Left            =   480
      TabIndex        =   7
      Top             =   6000
      Width           =   11805
      Begin VB.Frame Frame2 
         Height          =   1215
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   11460
         Begin VB.ComboBox Cmb_BuscaTipReg 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   71
            Top             =   555
            Width           =   2835
         End
         Begin VB.CommandButton Command3 
            Caption         =   "E&xportar"
            Height          =   750
            Left            =   10320
            Picture         =   "frmMemos.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   240
            Width           =   1005
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Memorandum"
            Height          =   750
            Left            =   9120
            Picture         =   "frmMemos.frx":0270
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton Cmd_Eliminar 
            Caption         =   "&Eliminar"
            Height          =   735
            Left            =   8280
            Picture         =   "frmMemos.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton Cmd_Imprimir 
            Caption         =   "&Listado"
            Height          =   750
            Left            =   10320
            Picture         =   "frmMemos.frx":0C6C
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   240
            Width           =   765
         End
         Begin VB.TextBox txtFiltroMes 
            Height          =   315
            Left            =   1440
            TabIndex        =   2
            Top             =   555
            Width           =   585
         End
         Begin VB.TextBox txtFiltroAnio 
            Height          =   315
            Left            =   240
            TabIndex        =   1
            Top             =   555
            Width           =   960
         End
         Begin VB.CommandButton Cmd_OK 
            Caption         =   "&Ingreso"
            Height          =   735
            Left            =   7440
            Picture         =   "frmMemos.frx":1326
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "Buscar"
            Height          =   735
            Left            =   6600
            Picture         =   "frmMemos.frx":1768
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   240
            Width           =   765
         End
         Begin VB.Label Label25 
            Caption         =   "Tipo de Regularización :"
            Height          =   255
            Left            =   2400
            TabIndex        =   72
            Top             =   360
            Width           =   2040
         End
         Begin VB.Label Label18 
            Caption         =   "Mes:"
            Height          =   300
            Left            =   1440
            TabIndex        =   55
            Top             =   360
            Width           =   510
         End
         Begin VB.Label Label3 
            Caption         =   "Año:"
            Height          =   240
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   825
         End
      End
      Begin MSDataGridLib.DataGrid dtgMemos 
         Height          =   4695
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   11460
         _ExtentX        =   20214
         _ExtentY        =   8281
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   2
         RowDividerStyle =   3
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
   End
   Begin VB.Frame fraDatos 
      Height          =   6345
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   11850
      Begin VB.TextBox txtMntoPensionMensual 
         Height          =   375
         Left            =   10520
         MaxLength       =   10
         TabIndex        =   79
         Top             =   2880
         Width           =   855
      End
      Begin VB.ComboBox cmbDevengues 
         Height          =   315
         ItemData        =   "frmMemos.frx":1BAA
         Left            =   8160
         List            =   "frmMemos.frx":1BAC
         TabIndex        =   77
         Text            =   "0"
         Top             =   2880
         Width           =   615
      End
      Begin VB.CheckBox chkEssalud 
         Caption         =   "Descontar ESSALUD"
         Height          =   375
         Left            =   4920
         TabIndex        =   73
         Top             =   2880
         Width           =   1935
      End
      Begin VB.ComboBox Cmb_TipoReg 
         BackColor       =   &H00E0FFFF&
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   2880
         Width           =   2835
      End
      Begin VB.Frame Frame4 
         Caption         =   "Datos a Ingresar"
         Enabled         =   0   'False
         Height          =   2145
         Left            =   180
         TabIndex        =   38
         Top             =   3240
         Width           =   11445
         Begin VB.TextBox txtMesesAfectados 
            Height          =   345
            Left            =   1725
            MaxLength       =   100
            TabIndex        =   19
            Top             =   1680
            Width           =   4320
         End
         Begin VB.TextBox txtNumeroMemo 
            Height          =   315
            Left            =   1725
            MaxLength       =   30
            TabIndex        =   18
            Top             =   1290
            Width           =   2715
         End
         Begin MSMask.MaskEdBox txtFechaPago 
            Height          =   330
            Left            =   1725
            TabIndex        =   17
            Top             =   870
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   582
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtLIquidoFinal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            Height          =   330
            Left            =   9705
            TabIndex        =   63
            Top             =   1215
            Width           =   1410
         End
         Begin VB.TextBox txtDescFinal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            Height          =   360
            Left            =   9705
            TabIndex        =   62
            Top             =   825
            Width           =   1410
         End
         Begin VB.TextBox txtHaberFinal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            Height          =   360
            Left            =   9705
            TabIndex        =   61
            Top             =   465
            Width           =   1410
         End
         Begin VB.TextBox txtLiqReg 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            Height          =   330
            Left            =   8010
            TabIndex        =   59
            Top             =   1215
            Width           =   1305
         End
         Begin VB.TextBox txtDescReg 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            Height          =   360
            Left            =   8010
            TabIndex        =   58
            Top             =   840
            Width           =   1305
         End
         Begin VB.TextBox txtHaberReg 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            ForeColor       =   &H80000001&
            Height          =   360
            Left            =   8010
            TabIndex        =   57
            Top             =   465
            Width           =   1305
         End
         Begin VB.TextBox txtLiquido 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            Height          =   345
            Left            =   6210
            TabIndex        =   52
            Top             =   1215
            Width           =   1365
         End
         Begin VB.TextBox txtDescuento 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            Height          =   360
            Left            =   6210
            TabIndex        =   50
            Top             =   840
            Width           =   1365
         End
         Begin VB.TextBox txtMes 
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2880
            TabIndex        =   24
            Top             =   450
            Width           =   495
         End
         Begin VB.TextBox txtAnio 
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            Height          =   330
            Left            =   1740
            TabIndex        =   23
            Top             =   465
            Width           =   720
         End
         Begin VB.TextBox txtHaber 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   6210
            TabIndex        =   20
            Top             =   465
            Width           =   1365
         End
         Begin VB.TextBox txtTC 
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3825
            TabIndex        =   41
            Top             =   450
            Width           =   660
         End
         Begin VB.Label Label23 
            Caption         =   "Meses Afectados:"
            Height          =   240
            Left            =   210
            TabIndex        =   67
            Top             =   1725
            Width           =   1425
         End
         Begin VB.Label Label22 
            Caption         =   "Numero de Memo "
            Height          =   240
            Left            =   210
            TabIndex        =   65
            Top             =   1305
            Width           =   1455
         End
         Begin VB.Label Label21 
            Caption         =   "Valores Consolidados"
            Height          =   255
            Left            =   9675
            TabIndex        =   64
            Top             =   165
            Width           =   1590
         End
         Begin VB.Label Label20 
            Caption         =   "Valores a Regularizar"
            Height          =   240
            Left            =   6120
            TabIndex        =   60
            Top             =   165
            Width           =   1605
         End
         Begin VB.Label Label19 
            Caption         =   "Valores Actuales"
            Height          =   270
            Left            =   8040
            TabIndex        =   56
            Top             =   165
            Width           =   1350
         End
         Begin VB.Label Label16 
            Caption         =   "Monto Liquido Pagar"
            Height          =   255
            Left            =   4605
            TabIndex        =   51
            Top             =   1245
            Width           =   1605
         End
         Begin VB.Label Label15 
            Caption         =   "Monto Descuento"
            Height          =   270
            Left            =   4620
            TabIndex        =   49
            Top             =   870
            Width           =   1590
         End
         Begin VB.Label Label11 
            Caption         =   "Mes:"
            Height          =   255
            Left            =   2475
            TabIndex        =   48
            Top             =   510
            Width           =   450
         End
         Begin VB.Label Label10 
            Caption         =   "Año:"
            Height          =   240
            Left            =   1275
            TabIndex        =   47
            Top             =   510
            Width           =   465
         End
         Begin VB.Label Label4 
            Caption         =   "Periodo:"
            Height          =   240
            Left            =   225
            TabIndex        =   46
            Top             =   495
            Width           =   870
         End
         Begin VB.Label Label12 
            Caption         =   "Monto Haber"
            Height          =   270
            Left            =   4620
            TabIndex        =   42
            Top             =   495
            Width           =   1545
         End
         Begin VB.Label Label9 
            Caption         =   "T/C"
            Height          =   255
            Left            =   3450
            TabIndex        =   40
            Top             =   480
            Width           =   465
         End
         Begin VB.Label Label8 
            Caption         =   "Fecha de Pago"
            Height          =   285
            Left            =   210
            TabIndex        =   39
            Top             =   885
            Width           =   1170
         End
      End
      Begin VB.CommandButton Cmd_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   750
         Left            =   6675
         Picture         =   "frmMemos.frx":1BAE
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   5520
         Width           =   750
      End
      Begin VB.CommandButton cmd_Aceptar 
         Caption         =   "&Aceptar"
         Height          =   735
         Left            =   4365
         Picture         =   "frmMemos.frx":2188
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   5520
         Width           =   735
      End
      Begin VB.Frame framPersonal 
         Caption         =   "Datos Personales"
         Height          =   1830
         Left            =   210
         TabIndex        =   25
         Top             =   960
         Width           =   11445
         Begin VB.TextBox txtDoc 
            BackColor       =   &H80000000&
            Height          =   285
            Left            =   7635
            TabIndex        =   27
            Top             =   420
            Width           =   1230
         End
         Begin VB.TextBox txtDirección 
            BackColor       =   &H80000000&
            Height          =   330
            Left            =   1410
            TabIndex        =   26
            Top             =   1365
            Width           =   7440
         End
         Begin MSDataListLib.DataCombo dtcNombres 
            Height          =   315
            Left            =   1425
            TabIndex        =   16
            Top             =   420
            Width           =   5220
            _ExtentX        =   9208
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   "DataCombo1"
         End
         Begin VB.Label Label27 
            Caption         =   "ESSALUD"
            Height          =   255
            Left            =   9480
            TabIndex        =   75
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label TxtEssalud 
            BackColor       =   &H80000000&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   10320
            TabIndex        =   74
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label lblMoneda 
            BackColor       =   &H80000000&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   10320
            TabIndex        =   45
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label14 
            Caption         =   "Label14"
            Height          =   15
            Left            =   9480
            TabIndex        =   44
            Top             =   1065
            Width           =   45
         End
         Begin VB.Label Label13 
            Caption         =   "Moneda"
            Height          =   255
            Left            =   9570
            TabIndex        =   43
            Top             =   885
            Width           =   675
         End
         Begin VB.Label Label1 
            Caption         =   "Nombre"
            Height          =   240
            Left            =   225
            TabIndex        =   37
            Top             =   480
            Width           =   990
         End
         Begin VB.Label lbldoc 
            Caption         =   "Nro.Doc."
            Height          =   240
            Left            =   6675
            TabIndex        =   36
            Top             =   465
            Width           =   810
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo de pensión"
            Height          =   240
            Left            =   195
            TabIndex        =   35
            Top             =   945
            Width           =   1200
         End
         Begin VB.Label lblTipoPensión 
            BackColor       =   &H80000000&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   2445
            TabIndex        =   34
            Top             =   900
            Width           =   4185
         End
         Begin VB.Label Label5 
            Caption         =   "Endoso"
            Height          =   300
            Left            =   6675
            TabIndex        =   33
            Top             =   915
            Width           =   720
         End
         Begin VB.Label lblEndoso 
            BackColor       =   &H80000000&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   7620
            TabIndex        =   32
            Top             =   855
            Width           =   1245
         End
         Begin VB.Label Label6 
            Caption         =   "Orden:"
            Height          =   255
            Left            =   9570
            TabIndex        =   31
            Top             =   480
            Width           =   570
         End
         Begin VB.Label lblOrden 
            BackColor       =   &H80000000&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   10320
            TabIndex        =   30
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Dirección:"
            Height          =   285
            Left            =   210
            TabIndex        =   29
            Top             =   1440
            Width           =   990
         End
         Begin VB.Label lblCodTipPension 
            BackColor       =   &H80000000&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1440
            TabIndex        =   28
            Top             =   900
            Width           =   915
         End
      End
      Begin VB.Frame Frame1 
         Height          =   750
         Left            =   210
         TabIndex        =   11
         Top             =   195
         Width           =   4530
         Begin VB.OptionButton OptDNI 
            Caption         =   "DNI"
            Height          =   255
            Left            =   180
            TabIndex        =   12
            Top             =   285
            Width           =   780
         End
         Begin VB.OptionButton OptPoliza 
            Caption         =   "Póliza"
            Height          =   330
            Left            =   1155
            TabIndex        =   13
            Top             =   240
            Width           =   960
         End
         Begin VB.TextBox txtDniPol 
            Height          =   300
            Left            =   2700
            TabIndex        =   14
            Top             =   270
            Width           =   1650
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar"
         Height          =   675
         Left            =   4875
         Picture         =   "frmMemos.frx":25CA
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   270
         Width           =   765
      End
      Begin VB.Label lblMtoPension 
         Caption         =   "Mto Pension Normal"
         Height          =   255
         Left            =   8880
         TabIndex        =   78
         Top             =   2955
         Width           =   1575
      End
      Begin VB.Label lblcmdDev 
         Caption         =   " Devengues"
         Height          =   255
         Left            =   7200
         TabIndex        =   76
         Top             =   2955
         Width           =   1095
      End
      Begin VB.Label Label24 
         Caption         =   "Tipo de Regularización :"
         Height          =   240
         Left            =   120
         TabIndex        =   68
         Top             =   2940
         Width           =   1770
      End
      Begin VB.Label Label17 
         Caption         =   "Numero de Póliza:"
         Height          =   330
         Left            =   8385
         TabIndex        =   54
         Top             =   555
         Width           =   1365
      End
      Begin VB.Label lblPoliza 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   9795
         TabIndex        =   53
         Top             =   420
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmMemos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cmd As ADODB.Command
Private rs_grilla As ADODB.Recordset
Public rs_Temp As ADODB.Recordset
Public sqlQuery As String

Private Sub chkEssalud_Click()
     Call calcula
End Sub

Private Sub Cmb_TipoReg_Click()
    Dim strCod As String
    Dim strCodTip As String
    
    strCod = fgObtenerCodigo_TextoCompuesto(Me.Cmb_TipoReg.Text)
    'Frame4.Enabled = False
    If strCod = "A" Then
        chkEssalud.Value = 0
        TxtEssalud.Caption = ""
    End If

    If strCod <> "" Then
        Frame4.Enabled = True
    End If
        
    vgSql = "Select NVL(max(cod_memo),0) as maximo From pp_tmae_regmemos where Cod_Tipopago='" & strCod & "'"
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
        If vgRs!maximo <> 0 Then
            txtNumeroMemo.Text = CDbl(vgRs!maximo) + 1
        Else
            txtNumeroMemo.Text = 1
        End If
    Else
        txtNumeroMemo.Text = 1
    End If
    
    Dim strCodTipReg As String
    strCodTipReg = fgObtenerCodigo_TextoCompuesto(Me.Cmb_TipoReg.Text)
    
    If strCodTipReg = "M" Then
    
        Me.lblcmdDev.Visible = True
        Me.cmbDevengues.Visible = True
        Me.lblMtoPension.Visible = True
        Me.txtMntoPensionMensual.Visible = True
    
    End If
    
     
End Sub

Private Sub cmd_Aceptar_Click()
    
Dim rs As ADODB.Recordset
Dim rsRep As ADODB.Recordset
Dim varPeriodo As String

On Error GoTo mierror


If rs_Temp!Fec_FallBen <> "" Or rs_Temp!Fec_FallBen = Null Then
    If MsgBox("La persona seleccionada Tiene fecha de Fallecimiento, Desea Agregar la Regularización", vbYesNo + vbCritical, "Ingreso de Regularizaciones") = vbNo Then
        Exit Sub
    End If
End If
    
If lblPoliza.Caption = "" Or txtHaber.Text = "0" Or IsDate(txtFechaPago.FormattedText) = False Then
    MsgBox "Falta Ingresar Datos", vbCritical, "Ingreso de Memos"
    Exit Sub
End If

If dtcNombres.Text = "" Then
    MsgBox "Elegir un Beneficiario", vbCritical, "Ingreso de Memos"
    dtcNombres.SetFocus
    Exit Sub
End If

Dim strCodTipReg As String
strCodTipReg = fgObtenerCodigo_TextoCompuesto(Me.Cmb_TipoReg.Text)

If Trim(strCodTipReg) = "" Then
    MsgBox "Se debe seleccionar un Tipo de Regularización", vbCritical, "Ingreso de Memos"
    Me.Cmb_TipoReg.SetFocus
    Exit Sub
End If
'MAFM

Set cmd = New ADODB.Command
Set rs = New ADODB.Recordset
cmd.ActiveConnection = vgConexionBD
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "PP_LISTA_PDT.Busca_MEMOS"
cmd.Parameters.Append cmd.CreateParameter("periodo", adVarChar, adParamInput, 6, txtAnio & txtMes)
cmd.Parameters.Append cmd.CreateParameter("poliza", adInteger, adParamInput, , CDbl(lblPoliza.Caption))
cmd.Parameters.Append cmd.CreateParameter("orden", adInteger, adParamInput, , CInt(lblOrden.Caption))
Set rs = cmd.Execute
'rs.CursorLocation = adUseClient
'rs.Open "PP_LISTA_PDT.Busca_MEMOS('" & txtAnio & txtMes & "'," & CDbl(lblPoliza.Caption) & "," & CInt(lblOrden.Caption) & ")", vgConexionBD, adOpenStatic, adLockReadOnly
If Not rs.EOF Then
    If rs!cantidad > 0 Then
        MsgBox "Ya se ingreso una Regularización para esta persona verifique si es necesario eliminar el memo para Ingresarlo nuevamente", vbCritical, "Ingreso de Memos"
        Exit Sub
    End If
End If


Set cmd = New ADODB.Command
Set rs = New ADODB.Recordset
cmd.ActiveConnection = vgConexionBD
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "PP_LISTA_PDT.Busca_num_MEMOS"
cmd.Parameters.Append cmd.CreateParameter("memo", adVarChar, adParamInput, 30, Trim(txtNumeroMemo.Text))
Set rs = cmd.Execute
'rs.CursorLocation = adUseClient
'rs.Open "PP_LISTA_PDT.Busca_num_MEMOS(" & CInt(txtNumeroMemo.Text) & ")", vgConexionBD, adOpenStatic, adLockReadOnly

If Not rs.EOF Then
    'If CDbl(txtNumeroMemo.Text) = rs!Memo Or CDbl(txtNumeroMemo.Text) - 1 = rs!Memo Then
        'MsgBox "Ya se ingreso una Regularización con ese numero de memorandum. El ultimo numero guardado es : " & rs!Memo, vbCritical, "Ingreso de Memos"
        'Exit Sub
        'vgSql = "Select max(Numeromemo)as maximo From pp_tmae_regmemos where Cod_Tipopago='M'"
        'Set vgRs = vgConexionBD.Execute(vgSql)
        'If Not vgRs.EOF Then
        '    MsgBox "Ya se ingreso una Regularización con ese numero de memorandum. El ultimo numero guardado es : " & RS!maximo, vbCritical, "Ingreso de Memos"
        '    Exit Sub
        'End If
    'End If
End If

Set cmd.ActiveConnection = Nothing

Dim fecPago As String

fecPago = Format(txtFechaPago.Text, "yyyymmdd")
varPeriodo = txtAnio & txtMes
If Not rs_Temp.EOF Then

    Set cmd = New ADODB.Command
    Set rs = New ADODB.Recordset
    
    cmd.ActiveConnection = vgConexionBD
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "Pp_Insert_Memos"
    
    
    Dim vDevengues As Integer
    Dim vmto_pension As Double
 
        vDevengues = Val(cmbDevengues.Text)
  
    
    If Len(txtMntoPensionMensual.Text) > 0 Then
        vmto_pension = CDbl(txtMntoPensionMensual.Text)
    Else
        vmto_pension = 0
    End If
    
    cmd.Parameters.Append cmd.CreateParameter("Par_Poliza", adInteger, adParamInput, , CDbl(lblPoliza.Caption))
    cmd.Parameters.Append cmd.CreateParameter("PAR_ORDEN", adInteger, adParamInput, , lblOrden.Caption)
    cmd.Parameters.Append cmd.CreateParameter("PAR_PERIODO", adVarChar, adParamInput, 6, varPeriodo)
    cmd.Parameters.Append cmd.CreateParameter("Par_Fechapago", adVarChar, adParamInput, 8, fecPago)
    cmd.Parameters.Append cmd.CreateParameter("Haber", adDouble, adParamInput, , CDbl(txtHaber.Text))
    cmd.Parameters.Append cmd.CreateParameter("Descuento", adDouble, adParamInput, , CDbl(txtDescuento.Text))
    cmd.Parameters.Append cmd.CreateParameter("Liquido", adDouble, adParamInput, , CDbl(txtLiquido.Text))
    cmd.Parameters.Append cmd.CreateParameter("Numeromemo", adVarChar, adParamInput, 30, txtNumeroMemo.Text)
    cmd.Parameters.Append cmd.CreateParameter("MesesAfectados", adVarChar, adParamInput, 100, txtMesesAfectados.Text)
    cmd.Parameters.Append cmd.CreateParameter("P_TipoReg varchar2", adVarChar, adParamInput, 2, strCodTipReg)
    cmd.Parameters.Append cmd.CreateParameter("pnum_devengues", adInteger, adParamInput, , vDevengues)
    cmd.Parameters.Append cmd.CreateParameter("pmto_pension", adDouble, adParamInput, , vmto_pension)
     
    cmd.Execute

    'vgConexionBD.Execute "Pp_Insert_Memos(" & CDbl(lblPoliza.Caption) & "," & lblOrden.Caption & ",'" & varPeriodo & "','" & Format(txtFechaPago.Text, "yyyymmdd") & "'," & CDbl(txtHaber.Text) & "," & CDbl(txtDescuento.Text) & "," & CDbl(txtLiquido.Text) & ",'" & txtNumeroMemo.Text & "','" & txtMesesAfectados.Text & "')"
    MsgBox "Registro Insertado en la Planilla", vbInformation, "Memos"
End If

If lblCodTipPension.Caption = "08" Or lblCodTipPension.Caption = "09" Or lblCodTipPension.Caption = "10" Or lblCodTipPension.Caption = "11" Or lblCodTipPension.Caption = "12" Then
    If dtcNombres.VisibleCount > 0 Then
        If MsgBox("Desea Ingresar a otra persona para este memorandum", vbQuestion + vbYesNo, "Ingreso de Regularizaciones") = vbYes Then
            
            Exit Sub
        End If
    End If
End If


strCodTipReg = fgObtenerCodigo_TextoCompuesto(Me.Cmb_TipoReg.Text)

If strCodTipReg = "M" Then

    Call reportMemoRegularizacion(lblPoliza.Caption, txtNumeroMemo.Text, lblOrden.Caption)

Else

         
        Dim objRep As New ClsReporte
        
        Set rsRep = New ADODB.Recordset
        rsRep.CursorLocation = adUseClient
        rsRep.Open "PP_LISTA_PDT.Lista_Datos_Reportes(" & CDbl(lblPoliza.Caption) & ",'" & txtNumeroMemo & "'," & CInt(lblEndoso.Caption) & ")", vgConexionBD, adOpenStatic, adLockReadOnly
        
        Dim FechaPago As String
        FechaPago = Mid(txtFechaPago.Text, 1, 2) & " de " & Format(CDate(txtFechaPago.Text), "mmmm") & " de " & Format(CDate(txtFechaPago.Text), "yyyy")
        
        Dim LNGa As Long
        LNGa = CreateFieldDefFile(rsRep, Replace(UCase(strRpt & "Estructura\Rpt_formato1.rpt"), ".RPT", ".TTX"), 1)
             
        If objRep.CargaReporte(strRpt, "Rpt_formato2.rpt", "Informe de Memos ingresados al sistema", rsRep, True, _
                                ArrFormulas("pm_NumMemo", Format(txtNumeroMemo.Text, "00000")), _
                                ArrFormulas("pm_Haber", CDbl(txtHaber.Text)), _
                                ArrFormulas("pm_Descuento", CDbl(txtDescuento.Text)), _
                                ArrFormulas("pm_Liquido", CDbl(txtLiquido.Text)), _
                                ArrFormulas("pm_FechaPago", FechaPago), _
                                ArrFormulas("pm_Poliza", CDbl(lblPoliza.Caption)), _
                                ArrFormulas("pm_MesesAfec", txtMesesAfectados.Text), _
                                ArrFormulas("pm_anio", Mid(Format(CDate(txtFechaPago.Text), "yyyy"), 3, 2)), _
                                ArrFormulas("pm_TipoReg", Right(Trim(Me.Cmb_TipoReg.Text), Len(Trim(Me.Cmb_TipoReg.Text)) - 4))) = False Then
                                
            MsgBox "No se pudo abrir el reporte", vbInformation
            Exit Sub
        End If
        DoEvents
        
        
        Dim memo2 As String
        memo2 = Format(CDbl(txtNumeroMemo.Text) + 1, "00000")
        
        LNGa = CreateFieldDefFile(rsRep, Replace(UCase(strRpt & "Estructura\Rpt_formato1.rpt"), ".RPT", ".TTX"), 1)
             
        If objRep.CargaReporte(strRpt, "Rpt_formato1.rpt", "Informe de Memos ingresados al sistema", rsRep, True, _
                                ArrFormulas("pm_NumMemo", memo2), _
                                ArrFormulas("pm_Haber", CDbl(txtHaber.Text)), _
                                ArrFormulas("pm_Descuento", CDbl(txtDescuento.Text)), _
                                ArrFormulas("pm_Liquido", CDbl(txtLiquido.Text)), _
                                ArrFormulas("pm_FechaPago", FechaPago), _
                                ArrFormulas("pm_Poliza", CDbl(lblPoliza.Caption)), _
                                ArrFormulas("pm_anio", Mid(Format(CDate(txtFechaPago.Text), "yyyy"), 3, 2)), _
                                ArrFormulas("pm_TipoReg", Right(Trim(Me.Cmb_TipoReg.Text), Len(Trim(Me.Cmb_TipoReg.Text)) - 4))) = False Then
                                
            MsgBox "No se pudo abrir el reporte", vbInformation
            Exit Sub
        End If
        
End If
DoEvents


    
txtFiltroAnio.Text = txtAnio.Text
txtFiltroMes.Text = txtMes.Text
Call Buscar
fraDatos.Visible = False
fraListado.Visible = True

Call Limpiar

Exit Sub
mierror:
    Screen.MousePointer = 0
    If Err.Number <> 0 Then
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End If
End Sub

Private Sub Cmd_Cancelar_Click()
fraDatos.Visible = False
fraListado.Visible = True
Call Limpiar
End Sub

Private Function ValidaPeriodo() As Boolean


ValidaPeriodo = False
If txtFiltroAnio.Text = "" Then
    MsgBox "Falta Ingresar año", vbInformation, "Ingreso de Memos"
    txtFiltroAnio.SetFocus
    Exit Function
End If
If txtFiltroMes.Text = "" Then
    MsgBox "Falta Ingresar el mes", vbInformation, "Ingreso de Memos"
    txtFiltroMes.SetFocus
    Exit Function
End If
If Len(txtFiltroAnio.Text) < 4 Then
    MsgBox "año mal Ingresado", vbInformation, "Ingreso de Memos"
    txtFiltroAnio.SetFocus
    Exit Function
End If
If Len(txtFiltroMes.Text) < 2 Then
    MsgBox "Mes mal Ingresado", vbInformation, "Ingreso de Memos"
    txtFiltroMes.SetFocus
    Exit Function
End If
If Len(txtFiltroAnio.Text) > 4 Then
    MsgBox "año mal Ingresado", vbInformation, "Ingreso de Memos"
    txtFiltroAnio.SetFocus
    Exit Function
End If
If Len(txtFiltroMes.Text) > 2 Then
    MsgBox "Mes mal Ingresado", vbInformation, "Ingreso de Memos"
    txtFiltroMes.SetFocus
    Exit Function
End If




ValidaPeriodo = True

End Function

Private Sub Cmd_Eliminar_Click()

On Error GoTo mierror

Dim rsFech As New ADODB.Recordset
Dim fech As Date
rsFech.CursorLocation = adUseClient
rsFech.Open "select sysdate as fecha from dual", vgConexionBD, adOpenStatic, adLockReadOnly
fech = Format(rsFech!fecha, "dd/mm/yyyy")

If Format(fech, "yyyy") <> txtFiltroAnio.Text Then
    MsgBox "Solo se puede eliminar registros en el año y mes actual", vbInformation, "Ingreso de Memos"
    txtFiltroAnio.SetFocus
    Exit Sub
End If

If Format(fech, "mm") <> txtFiltroMes.Text Then
    MsgBox "Solo se puede eliminar registros en el año y mes actual", vbInformation, "Ingreso de Memos"
    txtFiltroAnio.SetFocus
    Exit Sub
End If


If dtgMemos.ApproxCount > 0 Then

    If MsgBox("Desea eliminar la Regularización Ingresada para el periodo " & txtFiltroMes & "/" & txtFiltroAnio, vbQuestion + vbYesNo, "Ingreso de Regularizaciones") = vbYes Then
        vgConexionBD.Execute "PP_LISTA_PDT.Eliminar_Memos('" & dtgMemos.Columns(11) & "'," & CDbl(dtgMemos.Columns(0)) & "," & CInt(dtgMemos.Columns(1)) & ")"
        Call Buscar
    Else
        Exit Sub
    End If
    
    MsgBox "Se elimino la regularización seleccionada", vbInformation, "Ingreso de Regularizaciones"
End If
Exit Sub
mierror:
MsgBox "No se puede eliminar, consulte con sistemas", vbCritical, "Regularizaciones"
End Sub

Private Sub Cmd_Imprimir_Click()

Dim RANGO As String

On Error GoTo mierror
    Dim objRep As New ClsReporte
    
    RANGO = "Periodo del : " & txtFiltroAnio & txtFiltroMes
    
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(rs_grilla, Replace(UCase(strRpt & "Estructura\Rpt_Memos.rpt"), ".RPT", ".TTX"), 1)
         
    If objRep.CargaReporte(strRpt, "Rpt_Memos.rpt", "Informe de Memos ingresados al sistema", rs_grilla, True, _
                            ArrFormulas("pm_periodo", RANGO)) = False Then
                            
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
    End If
    DoEvents

    
Exit Sub
mierror:
    MsgBox "No se pudo cargar el reporte", vbExclamation
End Sub

Private Sub Cmd_OK_Click()

If ValidaPeriodo = False Then
    Exit Sub
End If

Dim rsFech As New ADODB.Recordset
Dim fech As Date
rsFech.CursorLocation = adUseClient
rsFech.Open "select sysdate as fecha from dual", vgConexionBD, adOpenStatic, adLockReadOnly
fech = Format(rsFech!fecha, "dd/mm/yyyy")

'If Format(fech, "yyyy") <> txtFiltroAnio.Text Then
'    MsgBox "Solo se ingresará en el año y mes actual", vbInformation, "Ingreso de Memos"
'    txtFiltroAnio.SetFocus
'    Exit Sub
'End If

'If Format(fech, "mm") <> txtFiltroMes.Text Then
'    MsgBox "Solo se ingresará en el año y mes actual", vbInformation, "Ingreso de Memos"
'    txtFiltroAnio.SetFocus
'    Exit Sub
'End If

Call Limpiar

fraDatos.Visible = True
fraListado.Visible = False

txtAnio.Text = txtFiltroAnio.Text
txtMes.Text = txtFiltroMes.Text
OptDNI.SetFocus

'MAFM - 05/08/2019 - Se hereda la selección de tipo de reg. de la búsqueda al control del registro.
If Me.Cmb_BuscaTipReg.ListIndex > 0 Then
    Dim strCodTR_Busq As String
    Dim intIndexTR As Integer
    strCodTR_Busq = fgObtenerCodigo_TextoCompuesto(Me.Cmb_BuscaTipReg.Text)
    intIndexTR = fgBuscarPosicionCodigoCombo(strCodTR_Busq, Me.Cmb_TipoReg)
    Me.Cmb_TipoReg.ListIndex = intIndexTR
End If
'MAFM

End Sub

Private Function VALIDA() As Boolean


VALIDA = True
End Function

Private Sub cmdBuscar_Click()

If ValidaPeriodo = False Then
    Exit Sub
End If

Call Buscar

End Sub

Private Sub Buscar()
    Dim Periodo, strCodTR As String
    Dim param1  As New ADODB.Parameter
    On Error GoTo mierror
    
    Set cmd = New ADODB.Command
    Set rs_grilla = New ADODB.Recordset

    Periodo = txtFiltroAnio.Text & Format(txtFiltroMes.Text, "00")

'    cmd.ActiveConnection = vgConexionBD
'    cmd.CommandType = adCmdStoredProc
'    cmd.CommandText = "PP_LISTA_PDT.LISTA_MEMOS"
'    Set param1 = cmd.CreateParameter("periodo", adChar, adParamInput, 6, Periodo)
'    cmd.Parameters.Append param1
'
'    rs_grilla.CursorType = adOpenStatic
'    rs_grilla.CursorLocation = adUseClient
'    rs_grilla.LockType = 3
'    Set rs_grilla = cmd.Execute
    rs_grilla.CursorLocation = adUseClient
    'MAFM - 05/08/2019 - Se modifica la búsqueda de memos, si se tiene un Tipo de Reg. invoca otro SP.
    If Me.Cmb_BuscaTipReg.ListIndex > 0 Then
        strCodTR = fgObtenerCodigo_TextoCompuesto(Me.Cmb_BuscaTipReg.Text)
        rs_grilla.Open "PP_LISTA_PDT.LISTA_MEMOS_xTIPOPAGO('" & Periodo & "','" & strCodTR & "')", vgConexionBD, adOpenStatic, adLockReadOnly
    Else
        rs_grilla.Open "PP_LISTA_PDT.LISTA_MEMOS('" & Periodo & "')", vgConexionBD, adOpenStatic, adLockReadOnly
    End If
    'MAFM
    'rs_grilla.CursorLocation = adUseClient
    'rs_grilla.Open "PP_LISTA_PDT.LISTA_MEMOS('" & Periodo & "')", vgConexionBD, adOpenStatic, adLockReadOnly
    If Not rs_grilla.EOF Then
        Set dtgMemos.DataSource = rs_grilla
    End If
    
    Exit Sub
mierror:
    Screen.MousePointer = 0
    If Err.Number <> 0 Then
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End If
        'MsgBox "No se pudo Listar", vbExclamation, "Lista de Memos"
        
End Sub

Private Sub Command1_Click()

On Error GoTo mierror

Set rs_Temp = New ADODB.Recordset
rs_Temp.CursorLocation = adUseClient

Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    rst.CursorType = adOpenKeyset
    rst.LockType = adLockBatchOptimistic
    rst.CursorLocation = adUseClient
    

If OptDNI.Value = True Then
    rs_Temp.Open "PP_LISTA_PDT.LISTA_DATOS_MEMOS_DNI('" & txtDniPol.Text & "')", vgConexionBD, adOpenStatic, adLockReadOnly
ElseIf OptPoliza.Value = True Then
    rs_Temp.Open "PP_LISTA_PDT.LISTA_DATOS_MEMOS_POL(" & txtDniPol.Text & ")", vgConexionBD, adOpenStatic, adLockReadOnly
End If


    
If Not rs_Temp.EOF Then
    
    Set dtcNombres.RowSource = rs_Temp
        dtcNombres.ListField = "NombCombo"  '"NOMBRE"
        dtcNombres.BoundColumn = "NUM_IDENBEN"
        lblPoliza.Caption = rs_Temp!num_poliza
    If rs_Temp.RecordCount = 1 Then
        'dtcNombres.SelectedItem = 1
        
    End If
    
Else
    MsgBox "No hay información", vbInformation, "Ingreso de Memos"
    Exit Sub
End If

txtHaber.Text = 0
txtDescuento.Text = 0
txtLiquido.Text = 0

Exit Sub
mierror:
    Screen.MousePointer = 0
    If Err.Number <> 0 Then
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End If
    'MsgBox "No se pudo consultar datos", vbCritical, "Ingreso de Memos"
    
End Sub


Private Sub Command2_Click()

Dim rsRep As ADODB.Recordset
On Error GoTo mierror


Call reportMemoRegularizacion(CDbl(dtgMemos.Columns(0)), dtgMemos.Columns(14), CInt(dtgMemos.Columns(13)))

Dim objRep As New ClsReporte

Set rsRep = New ADODB.Recordset
rsRep.CursorLocation = adUseClient
rsRep.Open "PP_LISTA_PDT.Lista_Datos_Reportes(" & CDbl(dtgMemos.Columns(0)) & ",'" & dtgMemos.Columns(14) & "'," & CInt(dtgMemos.Columns(13)) & ")", vgConexionBD, adOpenStatic, adLockReadOnly

If Not rsRep.EOF Then
    'MAFM - 05/08/2019 - Se obtiene el tipo de reg. de la grilla.
    Dim StrTipoReg As String
    StrTipoReg = Trim(dtgMemos.Columns(15))
    'MAFM
    
    Dim FechaPago As String
    
    FechaPago = Mid(rsRep!fecha_pago, 1, 2) & " de " & Format(rsRep!fecha_pago, "mmmm") & " de " & Format(rsRep!fecha_pago, "yyyy")
    
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(rsRep, Replace(UCase(strRpt & "Estructura\Rpt_formato1.rpt"), ".RPT", ".TTX"), 1)
         
    If objRep.CargaReporte(strRpt, "Rpt_formato2.rpt", "Informe de Memos ingresados al sistema", rsRep, True, _
                            ArrFormulas("pm_NumMemo", dtgMemos.Columns(14).Text), _
                            ArrFormulas("pm_Haber", CDbl(dtgMemos.Columns(7))), _
                            ArrFormulas("pm_Descuento", CDbl(dtgMemos.Columns(8))), _
                            ArrFormulas("pm_Liquido", CDbl(dtgMemos.Columns(9))), _
                            ArrFormulas("pm_FechaPago", FechaPago), _
                            ArrFormulas("pm_Poliza", CDbl(dtgMemos.Columns(0))), _
                            ArrFormulas("pm_MesesAfec", dtgMemos.Columns(12).Text), _
                            ArrFormulas("pm_anio", Mid(Format(CDate(dtgMemos.Columns(2)), "yyyy"), 3, 2)), _
                            ArrFormulas("pm_TipoReg", Right(StrTipoReg, Len(StrTipoReg) - 4))) = False Then
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
    End If
    DoEvents
    
    Dim memo2 As String
    memo2 = Format(CDbl(dtgMemos.Columns(14)) + 1, "00000")
    
    LNGa = CreateFieldDefFile(rsRep, Replace(UCase(strRpt & "Estructura\Rpt_formato1.rpt"), ".RPT", ".TTX"), 1)
         
    If objRep.CargaReporte(strRpt, "Rpt_formato1.rpt", "Informe de Memos ingresados al sistema", rsRep, True, _
                            ArrFormulas("pm_NumMemo", memo2), _
                            ArrFormulas("pm_Haber", CDbl(dtgMemos.Columns(7))), _
                            ArrFormulas("pm_Descuento", CDbl(dtgMemos.Columns(8))), _
                            ArrFormulas("pm_Liquido", CDbl(dtgMemos.Columns(9))), _
                            ArrFormulas("pm_FechaPago", FechaPago), _
                            ArrFormulas("pm_Poliza", CDbl(dtgMemos.Columns(0))), _
                            ArrFormulas("pm_anio", Mid(Format(CDate(dtgMemos.Columns(2)), "yyyy"), 3, 2)), _
                            ArrFormulas("pm_TipoReg", Right(StrTipoReg, Len(StrTipoReg) - 4))) = False Then
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
    End If
    DoEvents
    
Else
    MsgBox "No se pudo Imprimir, No hay datos", vbExclamation, ""
    Exit Sub
End If
                
Exit Sub
mierror:

       MsgBox "No se pudo Imprimir", vbExclamation, ""
       
       
End Sub

Private Sub Command3_Click()
    Dim frmRep As New Frm_RepLiqPagoPen
    frmRep.Show 1
End Sub

Private Sub Command4_Click()
    Call reportMemoRegularizacion(lblPoliza.Caption, txtNumeroMemo.Text, lblOrden.Caption)
End Sub

Private Sub dtcNombres_Change()

Dim rspago As ADODB.Recordset

On Error GoTo mierror

If dtcNombres.BoundText <> "" Then
rs_Temp.MoveFirst

    rs_Temp.Find "NUM_IDENBEN=" & dtcNombres.BoundText
    
'    If rs_Temp!Cod_DerPen = "10" Or rs_Temp!Cod_EstPension = "10" Then
'        MsgBox "La persona seleccionada No tiene derecho a pensión", vbCritical, "Ingreso de memos"
'        Exit Sub
'    End If
    
        txtDoc.Text = rs_Temp!Num_IdenBen
        lblOrden.Caption = rs_Temp!Num_Orden
        txtDirección.Text = rs_Temp!Gls_DirBen
    
        txtDoc.Text = rs_Temp!Num_IdenBen
        lblOrden.Caption = rs_Temp!Num_Orden
        txtDirección.Text = rs_Temp!Gls_DirBen
        
        lblCodTipPension.Caption = rs_Temp!Cod_TipPension
        lblTipoPensión.Caption = rs_Temp!GLS_ELEMENTO
        lblEndoso.Caption = rs_Temp!num_endoso
        lblMoneda.Caption = rs_Temp!Cod_Moneda
        lblCodTipPension.Caption = rs_Temp!Cod_TipPension
        lblTipoPensión.Caption = rs_Temp!GLS_ELEMENTO
        '**INI GCP CORRECCION PARA TOMAR EL INDICADOR DE DESCUENTO ESSALUD AHORA cod_isapre
        'If rs_Temp!IND_BENDES = "S" Then
        If rs_Temp!cod_isapre = "00" Then
        '**FIN GCP CORRECCION PARA TOMAR EL INDICADOR DE DESCUENTO ESSALUD AHORA cod_isapre
            chkEssalud.Value = 0
            TxtEssalud.Caption = "NO"
        Else
            chkEssalud.Value = 1
            TxtEssalud.Caption = "SI"
        End If
      
        
        Set rspago = New ADODB.Recordset
        rspago.CursorLocation = adUseClient
        rspago.Open "PP_LISTA_PDT.LISTA_DATOS_PAGOS('" & txtAnio.Text & txtMes.Text & "'," & CDbl(lblPoliza.Caption) & "," & CInt(lblOrden.Caption) & ")", vgConexionBD, adOpenStatic, adLockReadOnly
        
        If Not rspago.EOF Then
            txtHaberReg.Text = IIf(IsNull(rspago!Mto_Haber), 0, rspago!Mto_Haber)
            txtDescReg.Text = IIf(IsNull(rspago!Mto_Descuento), 0, rspago!Mto_Haber)
            txtLiqReg.Text = IIf(IsNull(rspago!Mto_LiqPagar), 0, rspago!Mto_Haber)
        End If
        
        If lblMoneda.Caption = "US" Then
             txtTC.Text = fTipoCambioSBS("US", Format$(txtAnio.Text, "0000") & Format$(txtMes.Text, "00"))
        Else
             txtTC.Text = 0
        End If

End If

Exit Sub
mierror:
        MsgBox "Hay problemas al consultar la persona", vbInformation, "Ingreso de Memos"
        
End Sub





Private Sub Form_Load()
frmMemos.Top = 0
frmMemos.Left = 0

fraListado.Height = 6345
fraListado.Left = 120
fraListado.Width = 11805
fraListado.Top = 0

    
OptDNI.Value = True
fraListado.Visible = True
fraDatos.Visible = False

txtFechaPago.Format = "dd/mm/yyyy"
txtFechaPago.Mask = "##/##/####"
'txtFiltroAnio.SetFocus

'MAFM - 01/08/2019 - Carga los combos de tipo de reg. para la búsqueda y el registro de Memos.
fgComboTipoRegularizacion Me.Cmb_TipoReg, True
fgComboTipoRegularizacion Me.Cmb_BuscaTipReg, True
Frame4.Enabled = False
'MAFM


Me.lblcmdDev.Visible = False
Me.cmbDevengues.Visible = False
Me.lblMtoPension.Visible = False
Me.txtMntoPensionMensual.Visible = False

Dim x As Integer
For x = 0 To 300

    cmbDevengues.AddItem (x)


Next




End Sub


Private Sub MaskMtoPensionNormal_Change()

End Sub



Private Sub txtAnio_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Case vbKeyBack
Case 13
    SendKeys "{TAB}"
Case Else
    KeyAscii = 0
End Select
End Sub

Private Sub txtDniPol_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Case vbKeyBack
Case 13
    SendKeys "{TAB}"
Case Else
    KeyAscii = 0
End Select
End Sub

Private Sub txtDniPol_LostFocus()

If OptDNI.Value = True Then
    txtDniPol.Text = Format(txtDniPol.Text, "00000000")
End If

End Sub

Private Sub txtFechaPago_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Case vbKeyBack
Case 13
    SendKeys "{TAB}"
    SendKeys "{HOME}+{END}"
Case Else
    KeyAscii = 0
End Select
End Sub





Private Sub txtFiltroAnio_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Case vbKeyBack
Case 13
    SendKeys "{TAB}"
Case Else
    KeyAscii = 0
End Select
End Sub

Private Sub txtFiltroMes_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Case vbKeyBack
Case 13
    SendKeys "{TAB}"
Case Else
    KeyAscii = 0
End Select
End Sub

Private Sub txtHaber_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Case 46
Case vbKeyBack
Case 13
    SendKeys "{TAB}"
Case Else
    KeyAscii = 0
End Select
End Sub
Private Sub txtHaber_LostFocus()

   Call calcula

End Sub

Private Sub calcula()
     If chkEssalud.Value = 1 Then
        txtDescuento.Text = Round(CDbl(txtHaber.Text) * 0.04, 2)
        txtLiquido.Text = Round(CDbl(txtHaber.Text) - CDbl(txtDescuento.Text), 2)
        
    Else
        txtDescuento.Text = 0
        txtLiquido.Text = Round(CDbl(txtHaber.Text) - CDbl(txtDescuento.Text), 2)
    End If
    
    If IsNumeric(txtHaberReg.Text) And txtHaberReg.Text >= 0 Then
        txtHaberFinal.Text = CDbl(txtHaberReg.Text) + CDbl(txtHaber.Text)
        txtDescFinal.Text = CDbl(txtDescReg.Text) + CDbl(txtDescuento.Text)
        txtLIquidoFinal.Text = CDbl(txtLiqReg.Text) + CDbl(txtLiquido.Text)
    End If

End Sub

Private Sub txtMes_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Case vbKeyBack
Case 13
    SendKeys "{TAB}"
Case Else
    KeyAscii = 0
End Select
End Sub



Private Function fTipoCambioSBS(ByVal sCodMon As String, ByVal dFecha As String) As Double

Dim rs_TC As ADODB.Recordset
Dim Sql As String
Dim fecha As Date
Dim fechas As String
    'saco fecha de trabajo que es la q se registra a fin de mes
    fecha = "01/" & Mid(dFecha, 5, 2) & "/" & Mid(dFecha, 1, 4)
    fechas = DateAdd("m", -1, fecha)
    dFecha = Mid(fechas, 7, 4) & Mid(fechas, 4, 2)
    'Para tipo de cambio del mes
    Sql = "SELECT mto_moneda FROM MA_TVAL_MONEDA_SBS WHERE "
    Sql = Sql & "cod_moneda = '" & sCodMon & "' AND "
    Sql = Sql & "fec_moneda = '" & dFecha & "'"
    Set rs_TC = New ADODB.Recordset
    Set rs_TC = vgConexionBD.Execute(Sql)
    If rs_TC.EOF Then
        fTipoCambioSBS = 0
    Else
        fTipoCambioSBS = rs_TC!Mto_Moneda
    End If

End Function

Private Sub Limpiar()

txtDniPol.Text = ""
Set dtcNombres.RowSource = Nothing
dtcNombres.Text = ""
lblPoliza.Caption = ""
txtDoc.Text = ""
lblOrden.Caption = ""
lblCodTipPension.Caption = ""
lblTipoPensión.Caption = ""
lblEndoso.Caption = ""
lblMoneda.Caption = ""
txtDirección.Text = ""
chkEssalud.Value = 0
txtAnio.Text = ""
txtMes.Text = ""
txtTC.Text = ""
txtFechaPago.Mask = ""
txtFechaPago.Text = ""
txtFechaPago.Format = "dd/mm/yyyy"
txtFechaPago.Mask = "##/##/####"

txtHaber.Text = 0
txtHaberReg.Text = 0
txtHaberFinal.Text = 0

txtDescuento.Text = 0
txtDescReg.Text = 0
txtDescFinal.Text = 0

txtLiquido.Text = 0
txtLiqReg.Text = 0
txtLIquidoFinal.Text = 0

txtNumeroMemo.Text = ""
txtMesesAfectados.Text = ""
'Set rs_Temp = New ADODB.Recordset
'OptDNI.SetFocus

End Sub


Private Sub txtMesesAfectados_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 13
    SendKeys "{TAB}"
    SendKeys "{HOME}+{END}"
End Select
End Sub

Private Sub txtMntoPensionMensual_KeyPress(KeyAscii As Integer)
 
 ' Salimos si se ha pulsado la tecla de Retroceso
  If KeyAscii = 8 Then Exit Sub
  ' Salimos si es de 0 a 9
  If InStr("0123456789", Chr$(KeyAscii)) Then Exit Sub
  ' Si es punto y no está en el contenido salimos
  If KeyAscii = 46 And InStr(txtMntoPensionMensual.Text, ".") = 0 Then Exit Sub
  ' Borramos el Caracter introducido
  KeyAscii = 0
End Sub

Private Sub txtNumeroMemo_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 13
    SendKeys "{TAB}"
    SendKeys "{HOME}+{END}"
End Select
End Sub
Private Sub reportMemoRegularizacion(ByVal Par_Poliza As String, ByVal Par_memo As String, ByVal POrden As Integer)
Dim rs As ADODB.Recordset
Dim conn As ADODB.Connection

Set rs = New ADODB.Recordset
Set conn = New ADODB.Connection



Dim objRep As New ClsReporte
 conn.Provider = "OraOLEDB.Oracle"
                conn.ConnectionString = "PLSQLRSet=1;Data Source=" & vgNombreBaseDatos & ";" & "User ID=" & vgNombreUsuario & ";Password=" & vgPassWord & ";"
                conn.CursorLocation = adUseClient
                conn.Open
                
                
 Call DatosReporteMemo(Par_Poliza, Par_memo, POrden, rs, conn)
 
 Dim FechaPago As String
 FechaPago = Mid(txtFechaPago.Text, 1, 2) & " de " & Format(CDate(txtFechaPago.Text), "mmmm") & " de " & Format(CDate(txtFechaPago.Text), "yyyy")


If objRep.CargaReporte(strRpt, "RPT_MemoRegularizacion.rpt", "Informe de Memos ingresados al sistema", rs, True, _
            ArrFormulas("moneda", "SOLES"), _
            ArrFormulas("montobruto", CDbl(txtHaber.Text)), _
            ArrFormulas("essalud", CDbl(txtDescuento.Text)), _
            ArrFormulas("montoneto", CDbl(txtLiquido.Text)), _
            ArrFormulas("usuario", vgUsuario), _
            ArrFormulas("numeromeno", "75"), _
            ArrFormulas("fechalarga", FechaPago)) = False Then

           MsgBox "No se pudo abrir el reporte", vbInformation
           Exit Sub
                        
                        
End If
 DoEvents
 
 
 If objRep.CargaReporte(strRpt, "RPT_MemoRegularizacion.rpt", "Informe de Memos ingresados al sistema", rs, True, _
            ArrFormulas("moneda", "SOLES"), _
            ArrFormulas("montobruto", CDbl(txtHaber.Text)), _
            ArrFormulas("essalud", CDbl(txtDescuento.Text)), _
            ArrFormulas("montoneto", CDbl(txtLiquido.Text)), _
            ArrFormulas("usuario", vgUsuario), _
            ArrFormulas("numeromeno", "76"), _
            ArrFormulas("fechalarga", FechaPago)) = False Then

           MsgBox "No se pudo abrir el reporte", vbInformation
           Exit Sub
                        
                        
End If
 DoEvents

   conn.Close

   Set rs = Nothing
   Set conn = Nothing

       

End Sub
Private Sub DatosReporteMemo(ByVal Par_Poliza As String, ByVal Par_memo As String, ByVal Par_Orden As Integer, ByRef rs As ADODB.Recordset, ByRef conn As ADODB.Connection)

       
                Dim texto As String
                
                Dim objCmd As ADODB.Command
                Set rs = New ADODB.Recordset
         
                
                Dim param1 As ADODB.Parameter
                Dim param2 As ADODB.Parameter
                
                
                Set objCmd = CreateObject("ADODB.Command")
                Set objCmd.ActiveConnection = conn
                
                objCmd.CommandText = "PP_LISTA_PDT.Lista_MemoRgularizacion"
                objCmd.CommandType = adCmdStoredProc
                
                    
                Set param1 = objCmd.CreateParameter("Par_Poliza", adVarChar, adParamInput, 15, Par_Poliza)
                objCmd.Parameters.Append param1
                
          
                Set param2 = objCmd.CreateParameter("Par_Orden", adInteger, adParamInput, 2, Par_Orden)
                objCmd.Parameters.Append param2
                
        
               Set rs = objCmd.Execute
                
                   
      
    
End Sub

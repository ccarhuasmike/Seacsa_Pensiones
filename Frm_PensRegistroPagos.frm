VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Frm_PensRegistroPagos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantención de Pagos a Terceros - Gastos de Sepelio."
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   9390
   Begin VB.TextBox txtFecha_Solicitud 
      Height          =   285
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   74
      Top             =   4080
      Width           =   1200
   End
   Begin VB.Frame Fra_AntRecep 
      Caption         =   "  Antecedentes del Receptor  "
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
      Height          =   2775
      Left            =   120
      TabIndex        =   61
      Top             =   1080
      Width           =   9135
      Begin VB.TextBox Txt_NumIdentRep 
         Height          =   285
         Left            =   3900
         MaxLength       =   25
         TabIndex        =   9
         Top             =   650
         Width           =   1845
      End
      Begin VB.TextBox Txt_TelefRecep 
         Height          =   285
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   19
         Top             =   2220
         Width           =   1800
      End
      Begin VB.TextBox Txt_NombRecep 
         Height          =   285
         Left            =   1575
         MaxLength       =   25
         TabIndex        =   10
         Top             =   975
         Width           =   2805
      End
      Begin VB.TextBox Txt_DomicRecep 
         Height          =   285
         Left            =   1575
         MaxLength       =   50
         TabIndex        =   14
         Top             =   1600
         Width           =   7185
      End
      Begin VB.TextBox Txt_EmailRecep 
         Height          =   285
         Left            =   4080
         MaxLength       =   40
         TabIndex        =   20
         Top             =   2220
         Width           =   4680
      End
      Begin VB.TextBox Txt_ApPaterno 
         Height          =   285
         Left            =   1575
         MaxLength       =   20
         TabIndex        =   12
         Top             =   1290
         Width           =   2805
      End
      Begin VB.TextBox Txt_ApMaterno 
         Height          =   285
         Left            =   5745
         MaxLength       =   20
         TabIndex        =   13
         Top             =   1290
         Width           =   3015
      End
      Begin VB.ComboBox Cmb_NumIdentRep 
         BackColor       =   &H00E0FFFF&
         Height          =   315
         Left            =   1575
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   630
         Width           =   2235
      End
      Begin VB.TextBox Txt_NombSegRecep 
         Height          =   285
         Left            =   5745
         MaxLength       =   25
         TabIndex        =   11
         Top             =   975
         Width           =   3015
      End
      Begin VB.CommandButton Cmd_BuscarDir 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Efectuar Busqueda de Dirección"
         Top             =   1900
         Width           =   300
      End
      Begin VB.ComboBox Cmb_TipoPersona 
         BackColor       =   &H00E0FFFF&
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   2235
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Email"
         Height          =   255
         Index           =   13
         Left            =   3540
         TabIndex        =   72
         Top             =   2220
         Width           =   465
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Tipo de persona"
         Height          =   270
         Index           =   9
         Left            =   240
         TabIndex        =   71
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Teléfono"
         Height          =   270
         Index           =   19
         Left            =   240
         TabIndex        =   70
         Top             =   2220
         Width           =   840
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Ubicación"
         Height          =   270
         Index           =   18
         Left            =   240
         TabIndex        =   69
         Top             =   1900
         Width           =   825
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Domicilio"
         Height          =   285
         Index           =   14
         Left            =   240
         TabIndex        =   68
         Top             =   1600
         Width           =   810
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "1er. Nombre"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   67
         Top             =   975
         Width           =   885
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Antecedentes del Receptor"
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
         Height          =   285
         Index           =   26
         Left            =   240
         TabIndex        =   66
         Top             =   0
         Width           =   2445
      End
      Begin VB.Label Lbl_Nombre 
         BackStyle       =   0  'Transparent
         Caption         =   "Ap. Paterno"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   65
         Top             =   1290
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         BackStyle       =   0  'Transparent
         Caption         =   "Ap. Materno"
         Height          =   255
         Index           =   5
         Left            =   4710
         TabIndex        =   64
         Top             =   1290
         Width           =   945
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Ident."
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   63
         Top             =   630
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         BackStyle       =   0  'Transparent
         Caption         =   "2do. Nombre"
         Height          =   255
         Index           =   6
         Left            =   4680
         TabIndex        =   62
         Top             =   960
         Width           =   945
      End
      Begin VB.Label Lbl_Distrito 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   6360
         TabIndex        =   17
         Top             =   1900
         Width           =   2175
      End
      Begin VB.Label Lbl_Provincia 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   3960
         TabIndex        =   16
         Top             =   1900
         Width           =   2295
      End
      Begin VB.Label Lbl_Departamento 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   15
         Top             =   1900
         Width           =   2295
      End
   End
   Begin VB.Frame Fra_DetPgo 
      Caption         =   "  Detalle de Pago  "
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
      Height          =   3015
      Left            =   120
      TabIndex        =   52
      Top             =   3840
      Width           =   4515
      Begin VB.TextBox Txt_Ruc 
         Height          =   285
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   26
         Top             =   2520
         Width           =   1485
      End
      Begin VB.TextBox Txt_FechaRecep 
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   21
         Top             =   600
         Width           =   1200
      End
      Begin VB.TextBox Txt_NroFactura 
         Height          =   285
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   25
         Top             =   2205
         Width           =   1485
      End
      Begin VB.TextBox Txt_FechaPgo 
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   22
         Top             =   915
         Width           =   1200
      End
      Begin VB.TextBox Txt_ValorPagado 
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   23
         Top             =   1530
         Width           =   1485
      End
      Begin VB.ComboBox Cmb_TipoDctoPago 
         BackColor       =   &H00E0FFFF&
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1845
         Width           =   1875
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de Solicitud"
         Height          =   255
         Left            =   240
         TabIndex        =   75
         Top             =   280
         Width           =   1575
      End
      Begin VB.Label Lbl_ValorCobrado 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   73
         Top             =   1230
         Width           =   1215
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "RUC Empresa Funeraria"
         Height          =   255
         Index           =   25
         Left            =   240
         TabIndex        =   60
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Detalle de Pago "
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
         Index           =   27
         Left            =   240
         TabIndex        =   59
         Top             =   0
         Width           =   1560
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Valor Cobrado en S/."
         Height          =   255
         Index           =   22
         Left            =   255
         TabIndex        =   58
         Top             =   1200
         Width           =   1500
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Fecha de Recepción"
         Height          =   255
         Index           =   20
         Left            =   255
         TabIndex        =   57
         Top             =   615
         Width           =   1560
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Dcto. Pago"
         Height          =   255
         Index           =   24
         Left            =   255
         TabIndex        =   56
         Top             =   2220
         Width           =   1110
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Fecha de Pago"
         Height          =   255
         Index           =   21
         Left            =   255
         TabIndex        =   55
         Top             =   915
         Width           =   1695
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Valor Pagado en S/."
         Height          =   255
         Index           =   23
         Left            =   240
         TabIndex        =   54
         Top             =   1530
         Width           =   1470
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Tipo Dcto. Pago"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   53
         Top             =   1890
         Width           =   1815
      End
   End
   Begin VB.Frame Fra_FormPgo 
      Caption         =   "  Forma de Pago  "
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
      Height          =   2415
      Left            =   4800
      TabIndex        =   45
      Top             =   3840
      Width           =   4455
      Begin VB.ComboBox Cmb_Sucursal 
         BackColor       =   &H00E0FFFF&
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   720
         Width           =   2955
      End
      Begin VB.TextBox Txt_NumCta 
         Height          =   285
         Left            =   960
         MaxLength       =   15
         TabIndex        =   31
         Top             =   1750
         Width           =   2940
      End
      Begin VB.ComboBox Cmb_TipCuenta 
         BackColor       =   &H00E0FFFF&
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1050
         Width           =   2940
      End
      Begin VB.ComboBox Cmb_ViaPago 
         BackColor       =   &H00E0FFFF&
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   375
         Width           =   2955
      End
      Begin VB.ComboBox Cmb_Banco 
         BackColor       =   &H00E0FFFF&
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1380
         Width           =   2940
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Forma de Pago"
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
         Index           =   28
         Left            =   240
         TabIndex        =   51
         Top             =   0
         Width           =   1305
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Sucursal"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   50
         Top             =   720
         Width           =   810
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N°Cuenta"
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   49
         Top             =   1770
         Width           =   795
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Banco"
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   48
         Top             =   1440
         Width           =   825
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Vía Pago"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   47
         Top             =   375
         Width           =   810
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Tipo Cta."
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   46
         Top             =   1080
         Width           =   825
      End
   End
   Begin VB.Frame Fra_Botones 
      Height          =   975
      Left            =   120
      TabIndex        =   39
      Top             =   6840
      Width           =   9135
      Begin VB.CommandButton Cmd_Grabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   1320
         Picture         =   "Frm_PensRegistroPagos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   4575
         Picture         =   "Frm_PensRegistroPagos.frx":06BA
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   6720
         Picture         =   "Frm_PensRegistroPagos.frx":0D74
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   3480
         Picture         =   "Frm_PensRegistroPagos.frx":0E6E
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   195
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Eliminar 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   2400
         Picture         =   "Frm_PensRegistroPagos.frx":1528
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Eliminar Receptor"
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   5625
         Picture         =   "Frm_PensRegistroPagos.frx":186A
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   200
         Width           =   730
      End
      Begin Crystal.CrystalReport Rpt_Imprimir 
         Left            =   8520
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.Frame Fra_Poliza 
      Caption         =   "Póliza / Pensionado"
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
      Height          =   1095
      Left            =   120
      TabIndex        =   38
      Top             =   0
      Width           =   9135
      Begin VB.TextBox Txt_PenPoliza 
         Height          =   285
         Left            =   840
         MaxLength       =   10
         TabIndex        =   0
         Top             =   360
         Width           =   1185
      End
      Begin VB.TextBox Txt_PenNumIdent 
         Height          =   285
         Left            =   5160
         MaxLength       =   16
         TabIndex        =   2
         Top             =   360
         Width           =   1875
      End
      Begin VB.CommandButton Cmd_BuscarPol 
         Height          =   375
         Left            =   8280
         Picture         =   "Frm_PensRegistroPagos.frx":1E44
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Buscar Póliza"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Cmd_Buscar 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8280
         Picture         =   "Frm_PensRegistroPagos.frx":1F46
         TabIndex        =   6
         ToolTipText     =   "Buscar Póliza"
         Top             =   600
         Width           =   615
      End
      Begin VB.ComboBox Cmb_PenNumIdent 
         BackColor       =   &H00E0FFFF&
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   2235
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Póliza"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   44
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Lbl_PenNombre 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   720
         Width           =   7335
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Ident."
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   42
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "N° End"
         Height          =   195
         Index           =   42
         Left            =   7080
         TabIndex        =   41
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Lbl_End 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   7680
         TabIndex        =   3
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Póliza / Pensionado"
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
         Index           =   4
         Left            =   120
         TabIndex        =   40
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   43
         Top             =   720
         Width           =   855
      End
   End
End
Attribute VB_Name = "Frm_PensRegistroPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vlCodViaPag As String
Dim vlAfp As String
Dim vlSwSeleccionado As Boolean

Dim vlCodTipoIdenBenCau As String
Dim vlNumIdenBenCau As String

Const clCodSinDerPen As String * 2 = "10"

Dim vlRegistro    As ADODB.Recordset
Dim vlI           As Integer, vlCodConcPgo  As String
Dim vlCodDir      As Integer, vlCodComuna   As String
Dim vlSw          As Boolean, vlCodViaPgo   As String
Dim vlCodSucursal As String, vlCodTipCuenta As String
Dim vlCodBco      As String, vlPasa         As Boolean
Dim vlRut         As String, vlFechaRecep   As String
Dim vlFechaP      As String, vlCodDerpen    As String
Dim vlFechaPgo    As String, vlTotal        As Double
Dim vlMtoCuoMor   As Double
Dim vlTipPersona  As String, vlTipDocPago As String
Dim vlNombreRegion As String
Dim vlNombreProvincia As String
Dim vlNombreComuna As String
Dim vlFechaSolicitud   As String

Dim vlCodTipoIdenBenRec As String
Dim vlNumIdenBenRec As String

Private Sub Cmb_PenNumIdent_Click()
If (Cmb_PenNumIdent <> "") Then
    vlPosicionTipoIden = Cmb_PenNumIdent.ListIndex
    vlLargoTipoIden = Cmb_PenNumIdent.ItemData(vlPosicionTipoIden)
    If (vlLargoTipoIden = 0) Then
        Txt_PenNumIdent.Text = "0"
        Txt_PenNumIdent.Enabled = False
    Else
        Txt_PenNumIdent = ""
        Txt_PenNumIdent.Enabled = True
        Txt_PenNumIdent.MaxLength = vlLargoTipoIden
        If (Txt_PenNumIdent <> "") Then Txt_PenNumIdent.Text = Mid(Txt_PenNumIdent, 1, vlLargoTipoIden)
    End If
End If
End Sub

Private Sub Cmd_Buscar_Click()
On Error GoTo Err_CmdBuscarClick

    Frm_Busqueda.flInicio ("Frm_PensRegistroPagos")
    
Exit Sub
Err_CmdBuscarClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Function flRecibe(vlNumPoliza, vlCodTipoIden, vlNumIden, vlNumEndoso)
    Txt_PenPoliza = vlNumPoliza
    Call fgBuscarPosicionCodigoCombo(vlCodTipoIden, Cmb_PenNumIdent)
    Txt_PenNumIdent = vlNumIden
    Lbl_End = vlNumEndoso
    Cmd_BuscarPol_Click
End Function

Private Sub Cmd_BuscarDir_Click()
On Error GoTo Err_Buscar

    Frm_BusDireccion.flInicio ("Frm_PensRegistroPagos")
    
Exit Sub
Err_Buscar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Function flRecibeDireccion(iNomDepartamento As String, iNomProvincia As String, iNomDistrito As String, iCodDir As String)
'FUNCION QUE RECIBE LOS DATOS DEL FORMULARIO DE BUSQUEDA de Dirección
    
    Lbl_Departamento = Trim(iNomDepartamento)
    Lbl_Provincia = Trim(iNomProvincia)
    Lbl_Distrito = Trim(iNomDistrito)
    vlCodDir = iCodDir
    Txt_TelefRecep.SetFocus

End Function

Private Sub Cmd_Eliminar_Click()
Dim vlOperacion As String

    Screen.MousePointer = 11
    
    vlCodTipoIdenBenRec = fgObtenerCodigo_TextoCompuesto(Cmb_PenNumIdent)
    vlNumIdenBenRec = Trim(Txt_NumIdentRep)
    vlCodConcPgo = cgPagoTerceroCuoMor
    
    vlOperacion = ""
    vgSql = ""
   'Verifica la existencia de la póliza
    vgSql = "SELECT NUM_POLIZA, COD_CONPAGO FROM PP_TMAE_PAGTERCUOMOR WHERE "
    vgSql = vgSql & "NUM_POLIZA = '" & Txt_PenPoliza & "' AND "
    vgSql = vgSql & "COD_CONPAGO = '" & vlCodConcPgo & "' AND "
    vgSql = vgSql & "COD_TIPOIDENSOLICITA = " & vlCodTipoIdenBenRec & " AND "
    vgSql = vgSql & "NUM_IDENSOLICITA = '" & vlNumIdenBenRec & "'"
    Set vlRegistro = vgConexionBD.Execute(vgSql)
    If Not vlRegistro.EOF Then
       vlOperacion = "E"
    End If
    vlRegistro.Close
    
    If (vlOperacion = "E") Then
        vgRes = MsgBox(" ¿ Esta seguro que desea Eliminar los Datos ? ", vbQuestion + vbYesNo + 256, "Operación de Eliminación")
        If vgRes <> 6 Then
            Screen.MousePointer = 0
            Exit Sub
        End If
        
        vgQuery = "DELETE FROM PP_TMAE_PAGTERCUOMOR WHERE "
        vgQuery = vgQuery & "NUM_POLIZA = '" & Txt_PenPoliza & "' And "
        vgQuery = vgQuery & "COD_CONPAGO = '" & vlCodConcPgo & "' and "
        vgQuery = vgQuery & "COD_TIPOIDENSOLICITA = " & vlCodTipoIdenBenRec & " AND "
        vgQuery = vgQuery & "NUM_IDENSOLICITA = '" & vlNumIdenBenRec & "'"
        vgConexionBD.Execute (vgQuery)
                 
'''        MSF_GrillaHistorica.Rows = 1
'''        MSF_GrillaHistorica.Rows = 2
        flCargarHistorico
        Cmd_Limpiar_Click
        Cmb_TipoPersona.SetFocus
    End If
    Screen.MousePointer = 0
    
Exit Sub
Err_Eliminar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Imprimir_Click()
 If Fra_Poliza.Enabled = False Then
    flImpresion
 End If
End Sub

'--------------------- Nº Póliza ------------------------------
Private Sub Txt_PenPoliza_KeyPress(KeyAscii As Integer)
On Error GoTo Err_TxtPenPolizaKeyPress

    If KeyAscii = 13 Then
       
        If Trim(Txt_PenPoliza.Text) = "" Then
          'MsgBox "Debe Ingresar Número de Póliza.", vbCritical, "Error de Datos"
          'Txt_PenPoliza.SetFocus
          'Exit Sub
        End If
        Txt_PenPoliza = UCase(Trim(Txt_PenPoliza))
        Txt_PenPoliza = Format(Txt_PenPoliza, "0000000000")
        Cmb_PenNumIdent.SetFocus
    End If
    
Exit Sub
Err_TxtPenPolizaKeyPress:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub
'--------------------- Tipo Identificación ------------------------------
Private Sub Cmb_PenNumIdent_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        If (Txt_PenNumIdent.Enabled = True) Then
            Txt_PenNumIdent.SetFocus
        Else
            Cmd_BuscarPol.SetFocus
        End If
    End If
End Sub
'--------------------- Nº Identificación ------------------------------
Private Sub Txt_PenNumIdent_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        If (Txt_PenNumIdent.Enabled = True) Then
            Cmd_BuscarPol.SetFocus
        Else
            Cmd_BuscarPol.SetFocus
        End If
    End If
End Sub
'--------------------- Tipo Persona ------------------------------
Private Sub Cmb_TipoPersona_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Cmb_TipoPersona.Text <> "" Then
          Cmb_NumIdentRep.SetFocus
       End If
    End If
End Sub
'--------------------- Tipo Identificación Receptor ---------------------
Private Sub Cmb_NumIdentRep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Cmb_NumIdentRep.Text <> "" Then
          Txt_NumIdentRep.SetFocus
       End If
    End If
End Sub
Private Sub Cmb_NumIdentRep_Click()
If (Cmb_NumIdentRep <> "") Then
    vlPosicionTipoIden = Cmb_NumIdentRep.ListIndex
    vlLargoTipoIden = Cmb_NumIdentRep.ItemData(vlPosicionTipoIden)
    If (vlLargoTipoIden = 0) Then
        Txt_NumIdentRep.Text = "0"
        Txt_NumIdentRep.Enabled = False
    Else
        Txt_NumIdentRep = ""
        Txt_NumIdentRep.Enabled = True
        Txt_NumIdentRep.MaxLength = vlLargoTipoIden
        If (Txt_NumIdentRep <> "") Then Txt_NumIdentRep.Text = Mid(Txt_NumIdentRep, 1, vlLargoTipoIden)
    End If
End If
End Sub
'--------------------- Nº Identificación Receptor ------------------------
Private Sub Txt_NumIdentRep_GotFocus()
    Txt_NumIdentRep.SelStart = 0
    Txt_NumIdentRep.SelLength = Len(Txt_NumIdentRep)
End Sub
Private Sub Txt_NumIdentRep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Txt_NumIdentRep.Text <> "" Then
            Txt_NumIdentRep = Trim(UCase(Txt_NumIdentRep))
       End If
       Txt_NombRecep.SetFocus
    End If
End Sub
Private Sub Txt_NumIdentRep_LostFocus()
   If (Trim(Txt_NumIdentRep)) <> "" Then
       Txt_NumIdentRep = UCase(Trim(Txt_NumIdentRep))
   End If
End Sub
'--------------------- Primer Nombre Receptor ---------------------
Private Sub Txt_NombRecep_GotFocus()
    Txt_NombRecep.SelStart = 0
    Txt_NombRecep.SelLength = Len(Txt_NombRecep)
End Sub
Private Sub Txt_NombRecep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If (Trim(Txt_NombRecep)) <> "" Then
          Txt_NombRecep = UCase(Trim(Txt_NombRecep))
          Txt_NombSegRecep.SetFocus
      End If
   End If
End Sub
Private Sub Txt_NombRecep_LostFocus()
   If (Trim(Txt_NombRecep)) <> "" Then
       Txt_NombRecep = UCase(Trim(Txt_NombRecep))
   End If
End Sub
'--------------------- Segundo Nombre Receptor ---------------------
Private Sub Txt_NombSegRecep_GotFocus()
    Txt_NombSegRecep.SelStart = 0
    Txt_NombSegRecep.SelLength = Len(Txt_NombSegRecep)
End Sub
Private Sub Txt_NombSegRecep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If (Trim(Txt_NombSegRecep)) <> "" Then
          Txt_NombSegRecep = UCase(Trim(Txt_NombSegRecep))
      End If
      Txt_ApPaterno.SetFocus
   End If
End Sub
Private Sub Txt_NombSegRecep_LostFocus()
   If (Trim(Txt_NombSegRecep)) <> "" Then
       Txt_NombSegRecep = UCase(Trim(Txt_NombSegRecep))
   End If
End Sub
'--------------------- Apellido Paterno Receptor ---------------------
Private Sub Txt_ApPaterno_GotFocus()
    Txt_ApPaterno.SelStart = 0
    Txt_ApPaterno.SelLength = Len(Txt_ApPaterno)
End Sub
Private Sub Txt_ApPaterno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If (Trim(Txt_ApPaterno)) <> "" Then
          Txt_ApPaterno = UCase(Trim(Txt_ApPaterno))
          Txt_ApMaterno.SetFocus
      End If
   End If
End Sub
Private Sub Txt_ApPaterno_LostFocus()
   If (Trim(Txt_ApPaterno)) <> "" Then
       Txt_ApPaterno = UCase(Trim(Txt_ApPaterno))
   End If
End Sub
'--------------------- Apellido Materno Receptor ---------------------
Private Sub Txt_ApMaterno_GotFocus()
    Txt_ApMaterno.SelStart = 0
    Txt_ApMaterno.SelLength = Len(Txt_ApMaterno)
End Sub
Private Sub Txt_Apmaterno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If (Trim(Txt_ApMaterno)) <> "" Then
          Txt_ApMaterno = UCase(Trim(Txt_ApMaterno))
      End If
      Txt_DomicRecep.SetFocus
   End If
End Sub
Private Sub Txt_ApMaterno_LostFocus()
   If (Trim(Txt_ApMaterno)) <> "" Then
       Txt_ApMaterno = UCase(Trim(Txt_ApMaterno))
   End If
End Sub
'--------------------- Domicilio Receptor ---------------------
Private Sub Txt_DomicRecep_GotFocus()
    Txt_DomicRecep.SelStart = 0
    Txt_DomicRecep.SelLength = Len(Txt_DomicRecep)
End Sub
Private Sub Txt_DomicRecep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If (Trim(Txt_DomicRecep)) <> "" Then
          Txt_DomicRecep = UCase(Trim(Txt_DomicRecep))
          Cmd_BuscarDir.SetFocus
      End If
   End If
End Sub
Private Sub Txt_DomicRecep_LostFocus()
   If (Trim(Txt_DomicRecep)) <> "" Then
       Txt_DomicRecep = UCase(Trim(Txt_DomicRecep))
   End If
End Sub
'--------------------- Telefono Receptor ---------------------
Private Sub Txt_TelefRecep_GotFocus()
    Txt_TelefRecep.SelStart = 0
    Txt_TelefRecep.SelLength = Len(Txt_TelefRecep)
End Sub
Private Sub Txt_TelefRecep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Txt_EmailRecep.SetFocus
    End If
End Sub
Private Sub Txt_TelefRecep_LostFocus()
   If (Trim(Txt_TelefRecep)) <> "" Then
       Txt_TelefRecep = UCase(Trim(Txt_TelefRecep))
   End If
End Sub
'--------------------- E-Mail Receptor ---------------------
Private Sub Txt_EmailRecep_GotFocus()
    Txt_EmailRecep.SelStart = 0
    Txt_EmailRecep.SelLength = Len(Txt_EmailRecep)
End Sub
Private Sub Txt_EmailRecep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Txt_EmailRecep = UCase(Trim(Txt_EmailRecep))
      Txt_FechaRecep.SetFocus
   End If
End Sub
Private Sub Txt_EmailRecep_LostFocus()
   If (Trim(Txt_EmailRecep)) <> "" Then
       Txt_EmailRecep = UCase(Trim(Txt_EmailRecep))
   End If
End Sub
'--------------------- Fecha de Recepción ---------------------
Private Sub Txt_FechaRecep_GotFocus()
    Txt_FechaRecep.SelStart = 0
    Txt_FechaRecep.SelLength = Len(Txt_FechaRecep)
End Sub
Private Sub Txt_FechaRecep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(Txt_FechaRecep <> "") Then
       If (flValidaFecha(Txt_FechaRecep) = False) Then
          Txt_FechaRecep = ""
          Exit Sub
       End If
       Txt_FechaRecep.Text = Format(CDate(Trim(Txt_FechaRecep)), "yyyymmdd")
       Txt_FechaRecep.Text = DateSerial(Mid((Txt_FechaRecep.Text), 1, 4), Mid((Txt_FechaRecep.Text), 5, 2), Mid((Txt_FechaRecep.Text), 7, 2))
       Txt_FechaPgo.SetFocus
    End If
End Sub
Private Sub Txt_FechaRecep_LostFocus()
   If (Trim(Txt_FechaRecep)) <> "" Then
       If (flValidaFecha(Txt_FechaRecep) = False) Then
          Txt_FechaRecep = ""
          Exit Sub
      End If
      Txt_FechaRecep.Text = Format(CDate(Trim(Txt_FechaRecep)), "yyyymmdd")
      Txt_FechaRecep.Text = DateSerial(Mid((Txt_FechaRecep.Text), 1, 4), Mid((Txt_FechaRecep.Text), 5, 2), Mid((Txt_FechaRecep.Text), 7, 2))
   End If
End Sub
'--------------------- Fecha de Pago ---------------------
Private Sub Txt_FechaPgo_GotFocus()
    Txt_FechaPgo.SelStart = 0
    Txt_FechaPgo.SelLength = Len(Txt_FechaPgo)
End Sub
Private Sub Txt_FechaPgo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(Txt_FechaPgo <> "") Then
        If (flValidaFecha(Txt_FechaPgo) = False) Then
           Txt_FechaPgo = ""
           Lbl_ValorCobrado = ""
           Exit Sub
        Else
            Txt_FechaPgo.Text = Format(CDate(Trim(Txt_FechaPgo)), "yyyymmdd")
            Lbl_ValorCobrado = Format(ObtieneMtoCuoMor(Txt_FechaPgo), "#,#0.00")
            Txt_FechaPgo.Text = DateSerial(Mid((Txt_FechaPgo.Text), 1, 4), Mid((Txt_FechaPgo.Text), 5, 2), Mid((Txt_FechaPgo.Text), 7, 2))
'            If (CDate(Txt_FechaRecep) < CDate(Txt_FechaPgo)) Then
'               MsgBox "La Fecha de Pago debe ser menor a la Fecha de Recepción", vbCritical, "Operación Cancelada"
'               Txt_FechaPgo.SetFocus
'               Exit Sub
'            End If
        End If
        Txt_ValorPagado.SetFocus
    End If
End Sub
Private Sub Txt_FechaPgo_LostFocus()
     If (flValidaFecha(Txt_FechaPgo) = False) Then
        Txt_FechaPgo = ""
        Lbl_ValorCobrado = ""
        Exit Sub
    Else
        Txt_FechaPgo.Text = Format(CDate(Trim(Txt_FechaPgo)), "yyyymmdd")
        Lbl_ValorCobrado = Format(ObtieneMtoCuoMor(Txt_FechaPgo), "#,#0.00")
        Txt_FechaPgo.Text = DateSerial(Mid((Txt_FechaPgo.Text), 1, 4), Mid((Txt_FechaPgo.Text), 5, 2), Mid((Txt_FechaPgo.Text), 7, 2))
'        If (CDate(Txt_FechaRecep) < CDate(Txt_FechaPgo)) Then
'           MsgBox "La Fecha de Pago debe ser menor a la Fecha de Recepción", vbCritical, "Operación Cancelada"
'           Txt_FechaPgo.SetFocus
'           Exit Sub
'        End If
    End If
End Sub
'--------------------- Valor Pagado ---------------------
Private Sub Txt_ValorPagado_GotFocus()
    Txt_ValorPagado.SelStart = 0
    Txt_ValorPagado.SelLength = Len(Txt_ValorPagado)
End Sub
Private Sub Txt_ValorPagado_Change()
If Not IsNumeric(Txt_ValorPagado) Then
    Txt_ValorPagado = ""
End If
End Sub
Private Sub Txt_ValorPagado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If IsNumeric(Txt_ValorPagado) Then
        Txt_ValorPagado = Format(Txt_ValorPagado, "#,#0.00")
        If (CDbl(Lbl_ValorCobrado) < CDbl(Txt_ValorPagado)) Then 'MC - 12/10/2007
           MsgBox "El Valor a Pagar debe ser menor al Valor a Cobrar", vbCritical, "Operación Cancelada"
           Txt_ValorPagado.SetFocus
           Exit Sub
        End If
        Cmb_TipoDctoPago.SetFocus
    End If
End If
End Sub
Private Sub Txt_ValorPagado_LostFocus()
    If IsNumeric(Txt_ValorPagado) Then
        Txt_ValorPagado = Format(Txt_ValorPagado, "#,#0.00")
    End If
End Sub
'--------------------- Tipo Documento de Pago ---------------------
Private Sub Cmb_TipoDctoPago_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Cmb_TipoDctoPago.Text <> "" Then
          Txt_NroFactura.SetFocus
       End If
    End If
End Sub
'--------------------- Nro Documento de Pago ---------------------
Private Sub Txt_NroFactura_GotFocus()
    Txt_NroFactura.SelStart = 0
    Txt_NroFactura.SelLength = Len(Txt_NroFactura)
End Sub
Private Sub Txt_NroFactura_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Txt_NroFactura = UCase(Trim(Txt_NroFactura))
       Txt_Ruc.SetFocus
    End If
End Sub
Private Sub Txt_NroFactura_LostFocus()
   If (Trim(Txt_NroFactura)) <> "" Then
       Txt_NroFactura = UCase(Trim(Txt_NroFactura))
   End If
End Sub
'--------------------- Ruc Empresa ---------------------
Private Sub TTxt_Ruc_GotFocus()
    Txt_Ruc.SelStart = 0
    Txt_Ruc.SelLength = Len(Txt_Ruc)
End Sub
Private Sub Txt_Ruc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Txt_Ruc = UCase(Trim(Txt_Ruc))
   Cmb_ViaPago.SetFocus
End If
End Sub
Private Sub Txt_Ruc_LostFocus()
   If (Trim(Txt_Ruc)) <> "" Then
       Txt_Ruc = UCase(Trim(Txt_Ruc))
   End If
End Sub
'--------------------- Via Pago ---------------------
Private Sub Cmb_ViaPago_Click()
    vlAfp = "241"
    
    vlCodViaPag = Trim(Mid(Cmb_ViaPago.Text, 1, (InStr(1, Cmb_ViaPago.Text, "-") - 1)))
    If vlCodViaPag = "01" Or vlCodViaPag = "04" Then 'caja
        If vlSw = False Then
            If (vlCodViaPag = "04") Then
                vgTipoSucursal = cgTipoSucursalAfp
            Else
                vgTipoSucursal = cgTipoSucursalSuc
            End If
            fgComboSucursal Cmb_Sucursal, vgTipoSucursal
        
            Cmb_TipCuenta.ListIndex = 0
            Cmb_Banco.ListIndex = 0
            If (vlCodViaPag = "04") Then
                vgPalabra = fgObtenerCodigo_TextoCompuesto(vlAfp)
                Call fgBuscarPosicionCodigoCombo(vgPalabra, Cmb_Sucursal)
            End If
            Cmb_TipCuenta.Enabled = False
            Cmb_Banco.Enabled = False
            If vlSwNumCta = False Then
                Txt_NumCta = ""
            End If
            Txt_NumCta.Enabled = False
            Cmb_Sucursal.Enabled = True
        End If
    Else
        If (vlCodViaPag = "00" Or vlCodViaPag = "05") And vlSw = False Then 'sin información
            vgTipoSucursal = cgTipoSucursalSuc
            fgComboSucursal Cmb_Sucursal, vgTipoSucursal
            
            Cmb_TipCuenta.ListIndex = 0
            Cmb_Banco.ListIndex = 0
            Cmb_Sucursal.ListIndex = 0
            Cmb_TipCuenta.Enabled = False
            Cmb_Banco.Enabled = False
            Cmb_Sucursal.Enabled = False
            If vlSwNumCta = False Then
                Txt_NumCta = ""
            End If
            Txt_NumCta.Enabled = False
        Else
            If vlSw = False Then
                
                vgTipoSucursal = cgTipoSucursalSuc
                fgComboSucursal Cmb_Sucursal, vgTipoSucursal
                
                If vlCodViaPag = "02" Or vlCodViaPag = "03" Then
                    Cmb_Sucursal.ListIndex = 0
                    Cmb_Sucursal.Enabled = False
                    Cmb_TipCuenta.Enabled = True
                    Cmb_Banco.Enabled = True
                    Txt_NumCta.Enabled = True
                Else
                    Cmb_TipCuenta.ListIndex = 0
                    Cmb_Banco.ListIndex = 0
                    Cmb_Sucursal.ListIndex = 0
                    Cmb_TipCuenta.Enabled = True
                    Cmb_Banco.Enabled = True
                    Cmb_Sucursal.Enabled = True
                    Txt_NumCta = ""
                    Txt_NumCta.Enabled = True
                End If
            End If
        End If
    End If
End Sub
Private Sub Cmb_ViaPago_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If (Cmb_Sucursal.Enabled = True) Then
            Cmb_Sucursal.SetFocus
        Else
            Cmd_Grabar.SetFocus
        End If
    End If
End Sub
'--------------------- Combo Sucursal ---------------------
Private Sub Cmb_Sucursal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If vlCodViaPgo = "02" Or vlCodViaPgo = "03" Then
      Cmb_TipCuenta.SetFocus
   Else
      Cmd_Grabar.SetFocus
   End If
End If
End Sub
'--------------------- Tipo Cuenta ---------------------
Private Sub Cmb_TipCuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If vlCodViaPgo = "02" Or vlCodViaPgo = "03" Then
      Cmb_Banco.SetFocus
   End If
End If
End Sub
'--------------------- Banco ---------------------------
Private Sub Cmb_Banco_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If vlCodViaPgo = "02" Or vlCodViaPgo = "03" Then
      Txt_NumCta.SetFocus
   End If
End If
End Sub
'--------------------- Nro Cuenta ---------------------
Private Sub Txt_NumCta_GotFocus()
    Txt_NumCta.SelStart = 0
    Txt_NumCta.SelLength = Len(Txt_NumCta)
End Sub
Private Sub Txt_NumCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Txt_NumCta = UCase(Trim(Txt_NumCta))
       Cmd_Grabar.SetFocus
    End If
End Sub

Private Sub Cmd_BuscarPol_Click()
On Error GoTo Err_Buscar

  If Trim(Txt_PenPoliza) <> "" Or Trim(Txt_PenNumIdent) <> "" Then
       If Trim(Txt_PenNumIdent) <> "" Then
          If Trim(Cmb_PenNumIdent) = "" Then
             MsgBox "Debe Tipo de Identificación.", vbCritical, "Error de Datos"
             Cmb_PenNumIdent.SetFocus
             Exit Sub
          End If
          Txt_PenNumIdent = Trim(UCase(Txt_PenNumIdent))
       End If
       'Permite Buscar los Datos del Beneficiario
       Call flValidarBen
       
   Else
     MsgBox "Debe ingresar el NºPóliza o Rut del Pensionado", vbCritical, "Error de Datos"
     Txt_PenPoliza.SetFocus
   End If

Exit Sub
Err_Buscar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Cancelar_Click()
    flLimpia

   Txt_PenPoliza.Enabled = True
   Cmb_PenNumIdent.Enabled = True
   Txt_PenNumIdent.Enabled = True
   Txt_PenPoliza = ""
    If (Cmb_PenNumIdent.ListCount <> 0) Then
        Cmb_PenNumIdent.ListIndex = 0
    End If
   Txt_PenNumIdent = ""
   Lbl_End = ""
   Lbl_PenNombre = ""
   
   Fra_Poliza.Enabled = True
   flDeshabilitarIngreso
   Txt_PenPoliza.SetFocus
End Sub

Private Sub cmd_grabar_Click()
Dim vlOperacion As String
On Error GoTo Err_Registrar

    If Fra_Poliza.Enabled = True Then
       Exit Sub
    End If
    
''    If vlCodDerpen = "10" Then
''       MsgBox " El Beneficiario Seleccionado No Tiene Derecho a Pensión " & Chr(13) & _
''              "          Sólo podrá Consultar los Datos del Registro", vbCritical, "Operación Cancelada"
''       Exit Sub
''    End If
        
    If (Cmb_PenNumIdent = "") Then
        MsgBox "Debe ingresar el Tipo de Identificación del Receptor", vbCritical, "Error de Datos"
        Cmb_PenNumIdent.SetFocus
        Exit Sub
    End If
    
    If Trim(Txt_PenNumIdent = "") Then
       MsgBox "Debe ingresar el número de identificación del Receptor", vbCritical, "Error de Datos"
       Txt_PenNumIdent.SetFocus
       Exit Sub
    End If
    
    If Trim(Txt_NombRecep) = "" Then
       MsgBox "Debe ingresar Nombre de Receptor", vbCritical, "Error de Datos"
       Txt_NombRecep.SetFocus
       Exit Sub
    End If
       
    If Trim(Txt_ApPaterno) = "" Then
       MsgBox "Debe ingresar Apellido Paterno del Receptor", vbCritical, "Error de Datos"
       Txt_ApPaterno.SetFocus
       Exit Sub
    End If
    
    If Trim(Txt_DomicRecep) = "" Then
       MsgBox "Debe ingresar Domicilio del Receptor", vbCritical, "Error de Datos"
       Txt_ApPaterno.SetFocus
       Exit Sub
    End If
    
   If Trim(txtFecha_Solicitud) = "" Then
       MsgBox "Debe Ingresar Fecha de Solicitud ", vbCritical, "Operación Cancelada"
       Txt_FechaRecep.SetFocus
       Exit Sub
    Else
      If (flValidaFecha(txtFecha_Solicitud) = False) Then
          txtFecha_Solicitud.SetFocus
          Exit Sub
      End If
    End If
    
    If Trim(Txt_FechaRecep) = "" Then
       MsgBox "Debe Ingresar Fecha de Recepción ", vbCritical, "Operación Cancelada"
       Txt_FechaRecep.SetFocus
       Exit Sub
    Else
      If (flValidaFecha(Txt_FechaRecep) = False) Then
          Txt_FechaRecep.SetFocus
          Exit Sub
      End If
    End If
    
    If Trim(Txt_FechaPgo) = "" Then
       MsgBox "Debe Ingresar Fecha de Pago ", vbCritical, "Operación Cancelada"
       Txt_FechaPgo.SetFocus
       Exit Sub
    Else
      If (flValidaFecha(Txt_FechaPgo) = False) Then
          Txt_FechaPgo.SetFocus
          Exit Sub
      End If
    End If
    
'    If (CDate(Txt_FechaRecep) < CDate(Txt_FechaPgo)) Then
'       MsgBox "La Fecha de Pago debe ser menor a la Fecha de Recepción", vbCritical, "Operación Cancelada"
'       Txt_FechaPgo.SetFocus
'       Exit Sub
'    End If
    
    If Trim(Lbl_ValorCobrado) = "" Then
       MsgBox "Debe Ingresar el Valor Cobrado ", vbCritical, "Operación Cancelada"
       Exit Sub
    End If
    
    If Trim(Txt_ValorPagado) = "" Then
       MsgBox "Debe Ingresar el Valor Pagado ", vbCritical, "Operación Cancelada"
       Txt_ValorPagado.SetFocus
       Exit Sub
    End If
    
    If (CDbl(Lbl_ValorCobrado) < CDbl(Txt_ValorPagado)) Then 'MC - 12/10/2007
       MsgBox "El Valor a Pagar debe ser menor al Valor a Cobrar", vbCritical, "Operación Cancelada"
       Txt_ValorPagado.SetFocus
       Exit Sub
    End If
    
    If Trim(Txt_NroFactura) = "" Then
       MsgBox "Debe Ingresar el Nro de Factura", vbCritical, "Operación Cancelada"
       Txt_NroFactura.SetFocus
       Exit Sub
    End If
    
    If Trim(Txt_Ruc) = "" Then
       MsgBox "Debe Ingresar el RUC de la Empresa Funeraria", vbCritical, "Operación Cancelada"
       Txt_Ruc.SetFocus
       Exit Sub
    End If
    
    Screen.MousePointer = 11
    
    vlOperacion = ""
    ''**vlCodConcPgo = Trim(Mid(Cmb_ConceptoPgo.Text, 1, (InStr(1, Cmb_ConceptoPgo.Text, "-") - 1)))
    vlCodViaPgo = Trim(Mid(Cmb_ViaPago.Text, 1, (InStr(1, Cmb_ViaPago.Text, "-") - 1)))
    vlCodSucursal = Trim(Mid(Cmb_Sucursal.Text, 1, (InStr(1, Cmb_Sucursal.Text, "-") - 1)))
    vlCodTipCuenta = Trim(Mid(Cmb_TipCuenta.Text, 1, (InStr(1, Cmb_TipCuenta.Text, "-") - 1)))
    vlCodBco = Trim(Mid(Cmb_Banco.Text, 1, (InStr(1, Cmb_Banco.Text, "-") - 1)))
    vlTipPersona = Trim(Mid(Cmb_TipoPersona.Text, 1, (InStr(1, Cmb_TipoPersona.Text, "-") - 1)))
    vlTipDocPago = Trim(Mid(Cmb_TipoDctoPago.Text, 1, (InStr(1, Cmb_TipoDctoPago.Text, "-") - 1)))
    
    If vlCodViaPgo = "00" Then
       MsgBox "Debe Seleccionar Forma de Pago", vbCritical, "Operación Cancelada"
       Cmb_ViaPago.SetFocus
       Screen.MousePointer = 0
       Exit Sub
    End If
    If vlCodViaPgo = "01" Then
       If vlCodSucursal = "0000" Then
            MsgBox "Debe seleccionar la Sucursal de la Vía de Pago", vbCritical, "Falta Información"
            Cmb_Sucursal.SetFocus
            Screen.MousePointer = 0
            Exit Sub
       End If
    End If
    If vlCodViaPgo = "02" Or vlCodViaPgo = "03" Then
       If vlCodTipCuenta = "00" Then
          MsgBox "Debe seleccionar el tipo de Cuenta", vbCritical, "Falta Información"
          Cmb_TipCuenta.SetFocus
          Screen.MousePointer = 0
          Exit Sub
       End If
       If vlCodBco = "00" Then
          MsgBox "Debe seleccionar el Banco", vbCritical, "Falta Información"
          Screen.MousePointer = 0
          Exit Sub
       End If
       If Trim(Txt_NumCta) = "" Then
          MsgBox "Debe ingresar el número de cuenta", vbCritical, "Falta Información"
          Txt_NumCta.SetFocus
          Screen.MousePointer = 0
          Exit Sub
       End If
    End If
    
    vlCodTipoIdenBenRec = fgObtenerCodigo_TextoCompuesto(Cmb_PenNumIdent)
    vlNumIdenBenRec = Trim(Txt_NumIdentRep)
    
    vlFechaRecep = Txt_FechaRecep
    vlFechaRecep = Format(vlFechaRecep, "yyyymmdd")
    
    vlFechaSolicitud = Me.txtFecha_Solicitud
    vlFechaSolicitud = Format(txtFecha_Solicitud, "yyyymmdd")
    
    vlFechaPgo = Txt_FechaPgo
    vlFechaPgo = Format(vlFechaPgo, "yyyymmdd")
    
    vgSql = ""
    vgSql = "SELECT NUM_POLIZA FROM PP_TMAE_PAGTERCUOMOR WHERE "
    vgSql = vgSql & "NUM_POLIZA = '" & Trim(Txt_PenPoliza) & "' AND "
    vgSql = vgSql & "COD_CONPAGO = '" & cgPagoTerceroCuoMor & "' "
''    vgSql = vgSql & "AND COD_TIPOIDENSOLICITA = " & vlCodTipoIdenBenRec & " "
''    vgSql = vgSql & "AND NUM_IDENSOLICITA = '" & vlNumIdenBenRec & "'"
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not (vgRs.EOF) Then
        vlOperacion = "A"
    Else
        vlOperacion = "I"
    End If
    vgRs.Close
    
    If vlOperacion = "A" Then
        vlResp = MsgBox(" ¿ Está seguro que desea Modificar los Datos ?", 4 + 32 + 256, "Actualización")
        If vlResp <> 6 Then
            Screen.MousePointer = 0
            Exit Sub
        End If
            
        'Actualiza los Datos en la tabla TMAE_PAGTERCERO
        Sql = ""
        Sql = "UPDATE PP_TMAE_PAGTERCUOMOR SET"
        Sql = Sql & " NUM_ENDOSO = " & Trim(Lbl_End) & ","
        Sql = Sql & " COD_TIPOPERSONA = '" & vlTipPersona & "',"
        Sql = Sql & " COD_TIPOIDENSOLICITA =" & vlCodTipoIdenBenRec & ","
        Sql = Sql & " NUM_IDENSOLICITA ='" & vlNumIdenBenRec & "',"
        Sql = Sql & " GLS_NOMSOLICITA = '" & Trim(Txt_NombRecep) & "',"
        If Trim(Txt_NombSegRecep) <> "" Then
            Sql = Sql & " GLS_NOMSEGSOLICITA = '" & Trim(Txt_NombSegRecep) & "',"
        Else
            Sql = Sql & " GLS_NOMSEGSOLICITA = NULL,"
        End If
        Sql = Sql & " GLS_PATSOLICITA = '" & Trim(Txt_ApPaterno) & "',"
        If Trim(Txt_ApMaterno) <> "" Then
            Sql = Sql & " GLS_MATSOLICITA = '" & Trim(Txt_ApMaterno) & "',"
        Else
            Sql = Sql & " GLS_MATSOLICITA = '" & Trim(Txt_ApMaterno) & "',"
        End If
        Sql = Sql & " GLS_DIRSOLICITA = '" & Trim(Txt_DomicRecep) & "',"
        Sql = Sql & " COD_DIRECCION = " & vlCodDir & ","
        If Trim(Txt_TelefRecep) <> "" Then
           Sql = Sql & " GLS_FONOSOLICITA = '" & Trim(Txt_TelefRecep) & "',"
        Else
           Sql = Sql & " GLS_FONOSOLICITA = NULL,"
        End If
                
        If Trim(Txt_EmailRecep) <> "" Then
            Sql = Sql & " GLS_CORREOSOLICITA = '" & Trim(Txt_EmailRecep) & "',"
        Else
            Sql = Sql & " GLS_CORREOSOLICITA = NULL,"
        End If
        
        Sql = Sql & " FEC_SOLPAGO = '" & vlFechaRecep & "',"
        Sql = Sql & " FEC_PAGO = '" & vlFechaPgo & "',"
        Sql = Sql & " MTO_COBRA = " & str(Lbl_ValorCobrado) & ","
        Sql = Sql & " MTO_PAGO = " & str(Txt_ValorPagado) & ","
        Sql = Sql & " COD_TIPODCTOPAGO = '" & vlTipDocPago & "',"
        If Trim(Txt_NroFactura) <> "" Then
           Sql = Sql & " NUM_DCTOPAGO = '" & Trim(Txt_NroFactura) & "',"
        Else
           Sql = Sql & " NUM_DCTOPAGO = NULL,"
        End If
        Sql = Sql & " COD_TIPOIDENFUN =" & CInt(cgTipoIdenRuc) & ","
        Sql = Sql & " NUM_IDENFUN = '" & Trim(Txt_Ruc) & "',"
        Sql = Sql & " COD_VIAPAGO = '" & Trim(vlCodViaPgo) & "',"
        Sql = Sql & " COD_SUCURSAL = '" & Trim(vlCodSucursal) & "',"
        Sql = Sql & " COD_TIPCUENTA = '" & Trim(vlCodTipCuenta) & "',"
        Sql = Sql & " COD_BANCO = '" & Trim(vlCodBco) & "',"
        If Trim(Txt_NumCta) <> "" Then
           Sql = Sql & " NUM_CUENTA = '" & Trim(Txt_NumCta) & "',"
        Else
           Sql = Sql & " NUM_CUENTA = NULL,"
        End If
        Sql = Sql & " COD_USUARIOMODI = '" & (vgUsuario) & "',"
        Sql = Sql & " FEC_MODI = '" & Format(Date, "yyyymmdd") & "',"
        Sql = Sql & " HOR_MODI = '" & Format(Time, "hhmmss") & "',"
        Sql = Sql & " FEC_SOL_AFP =  '" & vlFechaSolicitud & "'"
 
        Sql = Sql & " WHERE"
        Sql = Sql & " NUM_POLIZA = '" & Trim(Txt_PenPoliza) & "' and "
        Sql = Sql & " COD_CONPAGO = '" & cgPagoTerceroCuoMor & "' "
        ''**Sql = Sql & "and RUT_SOLICITA = " & vlRut & ""
        vgConexionBD.Execute (Sql)
    Else
        'Inserta los Datos en la Tabla TMAE_PAGTERCERO
        vlResp = MsgBox(" ¿ Está seguro que desea Ingresar los Datos ?", 4 + 32 + 256, "Actualización")
        If vlResp <> 6 Then
            Screen.MousePointer = 0
            Exit Sub
        End If
        Sql = ""
        Sql = "INSERT INTO PP_TMAE_PAGTERCUOMOR ("
        Sql = Sql & "NUM_POLIZA,NUM_ENDOSO,COD_CONPAGO,COD_TIPOPERSONA,"
        Sql = Sql & "COD_TIPOIDENSOLICITA,NUM_IDENSOLICITA,GLS_NOMSOLICITA,"
        Sql = Sql & "GLS_NOMSEGSOLICITA,GLS_PATSOLICITA,GLS_MATSOLICITA,"
        Sql = Sql & "GLS_DIRSOLICITA,COD_DIRECCION,GLS_FONOSOLICITA,"
        Sql = Sql & "GLS_CORREOSOLICITA,FEC_SOLPAGO,FEC_PAGO,"
        Sql = Sql & "COD_TIPODCTOPAGO,NUM_DCTOPAGO,MTO_COBRA,MTO_PAGO,"
        Sql = Sql & "COD_TIPOIDENFUN,NUM_IDENFUN,COD_VIAPAGO,COD_TIPCUENTA,COD_BANCO,"
        Sql = Sql & "NUM_CUENTA,COD_SUCURSAL,COD_USUARIOCREA,FEC_CREA,HOR_CREA, FEC_SOL_AFP "
        Sql = Sql & ") VALUES ("
        Sql = Sql & "'" & Trim(Txt_PenPoliza) & "',"
        Sql = Sql & "" & Trim(Lbl_End) & ","
        Sql = Sql & "'" & cgPagoTerceroCuoMor & "',"
        Sql = Sql & "'" & vlTipPersona & "',"
        Sql = Sql & "" & vlCodTipoIdenBenRec & ","
        Sql = Sql & "'" & vlNumIdenBenRec & "',"
        Sql = Sql & "'" & Trim(Txt_NombRecep) & "',"
        Sql = Sql & "'" & Trim(Txt_NombSegRecep) & "',"
        Sql = Sql & "'" & Trim(Txt_ApPaterno) & "',"
        Sql = Sql & "'" & Trim(Txt_ApMaterno) & "',"
        Sql = Sql & "'" & Trim(Txt_DomicRecep) & "',"
        Sql = Sql & "" & vlCodDir & ","
        
        If Trim(Txt_TelefRecep) <> "" Then
           Sql = Sql & "'" & Trim(Txt_TelefRecep) & "',"
        Else
           Sql = Sql & "NULL,"
        End If
                
        If Trim(Txt_EmailRecep) <> "" Then
            Sql = Sql & "'" & Trim(Txt_EmailRecep) & "',"
        Else
            Sql = Sql & "NULL,"
        End If
        
        Sql = Sql & "'" & vlFechaRecep & "',"
        Sql = Sql & "'" & vlFechaPgo & "',"
        Sql = Sql & "'" & vlTipDocPago & "',"
        Sql = Sql & "'" & Trim(Txt_NroFactura) & "',"
        Sql = Sql & "" & str(Lbl_ValorCobrado) & ","
        Sql = Sql & "" & str(Txt_ValorPagado) & ","
        Sql = Sql & "" & CInt(cgTipoIdenRuc) & ","
        Sql = Sql & "'" & Trim(Txt_Ruc) & "',"
        Sql = Sql & "'" & vlCodViaPgo & "',"
        Sql = Sql & "'" & Trim(vlCodTipCuenta) & "',"
        Sql = Sql & "'" & Trim(vlCodBco) & "',"
        If Trim(Txt_NumCta) <> "" Then
           Sql = Sql & "'" & Trim(Txt_NumCta) & "',"
        Else
           Sql = Sql & "NULL,"
        End If
        Sql = Sql & "'" & Trim(vlCodSucursal) & "',"
        Sql = Sql & "'" & (vgUsuario) & "',"
        Sql = Sql & "'" & Format(Date, "yyyymmdd") & "',"
        Sql = Sql & "'" & Format(Time, "hhmmss") & "',"
        Sql = Sql & "'" & vlFechaSolicitud & "'"
        Sql = Sql & ")"
        vgConexionBD.Execute (Sql)
    End If
     
    If (vlOperacion <> "") Then
'''        MSF_GrillaHistorica.Rows = 1
'''        MSF_GrillaHistorica.Rows = 2
        flCargarHistorico
    End If
    Screen.MousePointer = 0

Exit Sub
Err_Registrar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Limpiar_Click()
    Call flLimpia
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

Private Sub Form_Load()
On Error GoTo Err_Cargar

    Me.Top = 0
    Me.Left = 0

    Call fgComboGeneral(vgCodTabla_TipPer, Cmb_TipoPersona)
    Call fgComboGeneral(vgCodTabla_TipCta, Cmb_TipCuenta)
    Call fgComboGeneral(vgCodTabla_Bco, Cmb_Banco)
    Call fgComboGeneral(vgCodTabla_ViaPago, Cmb_ViaPago)
    Call fgComboSucursal(Cmb_Sucursal, "S")
    
    fgComboTipoPago Cmb_TipoDctoPago
    fgComboTipoIdentificacion Cmb_PenNumIdent
    fgComboTipoIdentificacion Cmb_NumIdentRep
    
    Call fgCargarTablaMoneda(vgCodTabla_TipMon, egTablaMoneda(), vgNumeroTotalTablasMoneda)


Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub



Private Sub txt_pennumident_lostfocus()
    Txt_PenNumIdent = Trim(UCase(Txt_PenNumIdent))
End Sub



Private Sub Txt_PenPoliza_LostFocus()
    Txt_PenPoliza = UCase(Trim(Txt_PenPoliza))
    Txt_PenPoliza = Format(Txt_PenPoliza, "0000000000")
End Sub


Function flHabilitarIngreso()

    Fra_Poliza.Enabled = False
    Fra_AntRecep.Enabled = True
    Fra_DetPgo.Enabled = True
    Fra_FormPgo.Enabled = True

End Function

Function flDeshabilitarIngreso()

    'Desactivar Todos los Controles del Formulario
    Fra_Poliza.Enabled = True
    Fra_AntRecep.Enabled = False
    Fra_DetPgo.Enabled = False
    Fra_FormPgo.Enabled = False
    
End Function

Function flValidarBen()
Dim fechafallesimiento As String
Dim cantBene As Integer
Dim vlRegAux As ADODB.Recordset
On Error GoTo Err_Validar

   Screen.MousePointer = 11
      
    vlCodTipoIdenBenCau = fgObtenerCodigo_TextoCompuesto(Cmb_PenNumIdent)
    vlNumIdenBenCau = Trim(Txt_PenNumIdent)
    
   'Verificar Número de Póliza, y saca el último Endoso
   vgPalabra = ""
   vgSql = ""
   If Txt_PenPoliza <> "" And Cmb_PenNumIdent <> "" And Txt_PenNumIdent <> "" Then
        vgPalabra = "NUM_POLIZA = '" & Txt_PenPoliza & "' AND "
        vgPalabra = vgPalabra & "cod_tipoidenBEN = " & (vlCodTipoIdenBenCau) & " and "
        vgPalabra = vgPalabra & "num_idenben = '" & (vlNumIdenBenCau) & "' "
   Else
     If Txt_PenPoliza <> "" Then
        vgSql = "SELECT COUNT(NUM_POLIZA) AS REG_POLIZA"
        vgSql = vgSql & " FROM PP_TMAE_BEN WHERE"
        vgSql = vgSql & " NUM_POLIZA = '" & Txt_PenPoliza & "' "
        ''vgSql = vgSql & " and COD_ESTPENSION <> '10'"
        Set vlRegistro = vgConexionBD.Execute(vgSql)
        If Not vlRegistro.EOF Then
           If (vlRegistro!reg_poliza) > 0 Then
               vgPalabra = "NUM_POLIZA = '" & Txt_PenPoliza & "' "
               ''vgPalabra = vgPalabra & " AND COD_ESTPENSION <> '10'"
           Else
               vgPalabra = "NUM_POLIZA = '" & Txt_PenPoliza & "'"
           End If
        Else
           vgPalabra = "NUM_POLIZA = '" & Txt_PenPoliza & "'"
        End If
     Else
        If (Cmb_PenNumIdent.Text <> "") And (Txt_PenNumIdent.Text <> "") Then
            vgPalabra = "cod_tipoidenBEN = " & (vlCodTipoIdenBenCau) & " "
            vgPalabra = vgPalabra & "AND num_idenben = '" & (vlNumIdenBenCau) & "' "
        End If
     End If
   End If
                
   'verifica si es unico venefiociario y tiene fecha de fallesimiento para que no haga la siguiente validación.
    cantBene = 0
    vgSql = "select fec_fallben from pp_tmae_ben where num_poliza='" & Trim(Txt_PenPoliza) & "' and num_endoso in(select max(num_endoso) from pp_tmae_ben where num_poliza='" & Trim(Txt_PenPoliza) & "') and num_orden=1"
    Set vlRegAux = vgConexionBD.Execute(vgSql)
    If Not vlRegAux.EOF Then
        Do While Not vlRegAux.EOF
            fechafallesimiento = "" & vlRegAux!Fec_FallBen
            vlRegAux.MoveNext
            cantBene = cantBene + 1
        Loop
        vlRegAux.Close
    End If
    
    If (cantBene > 1) Or (fechafallesimiento = "" And cantBene = 1) Then
    'Verifica que la Póliza corresponda a una Sobrevivencia de ....
        vgSql = "SELECT cod_tippension from pp_tmae_poliza "
        vgSql = vgSql & "WHERE "
        vgSql = vgSql & "num_poliza = '" & Trim(Txt_PenPoliza) & "' "
        vgSql = vgSql & "ORDER BY num_endoso desc "
        Set vlRegAux = vgConexionBD.Execute(vgSql)
        If Not vlRegAux.EOF Then
            If (vlRegAux!Cod_TipPension <> "09") And (vlRegAux!Cod_TipPension <> "10") And (vlRegAux!Cod_TipPension <> "11") And (vlRegAux!Cod_TipPension <> "12") Then
                vlRegAux.Close
                MsgBox "El Causante de la Póliza No se encuentra Fallecido o el Tipo de Pensión de la Póliza no tiene derecho a este Beneficio.", vbCritical, "Operación Cancelada"
                Cmd_Cancelar_Click
                Screen.MousePointer = 0
                Exit Function
            End If
        End If
        vlRegAux.Close
    End If
'F--- ABV 27/10/2007 ---
      
   vgSql = ""
   vgSql = "SELECT NUM_POLIZA,NUM_ENDOSO,NUM_ORDEN,COD_TIPOIDENBEN,NUM_IDENBEN,"
   vgSql = vgSql & " MTO_PENSION,COD_ESTPENSION,GLS_NOMBEN,GLS_PATBEN,"
   vgSql = vgSql & " GLS_MATBEN,FEC_MATRIMONIO FROM PP_TMAE_BEN"
   vgSql = vgSql & " Where "
   vgSql = vgSql & vgPalabra
   vgSql = vgSql & " ORDER BY num_orden asc, NUM_ENDOSO DESC "
   Set vlRegistro = vgConexionBD.Execute(vgSql)
   If Not vlRegistro.EOF Then
''      If (vlRegistro!Cod_EstPension) = "10" Then
''          MsgBox " El Beneficiario Seleccionado No Tiene Derecho a Pensión " & Chr(13) & _
''                 "          Sólo podrá Consultar los Datos del Registro", vbInformation, "Información"
''      End If
      Txt_PenPoliza = (vlRegistro!num_poliza)
      Call fgBuscarPosicionCodigoCombo(vlRegistro!Cod_TipoIdenBen, Cmb_PenNumIdent)
      Txt_PenNumIdent = (vlRegistro!Num_IdenBen)
      vlNumEndoso = (vlRegistro!num_endoso)
      Lbl_End = vlNumEndoso
      vlCodDerpen = (vlRegistro!Cod_EstPension)
      vlNumOrden = (vlRegistro!Num_Orden)
      vlNombre = (vlRegistro!Gls_NomBen) + " " + IIf(IsNull(vlRegistro!Gls_NomBen), "", (vlRegistro!Gls_NomBen)) + " " + (vlRegistro!Gls_PatBen) + " " + IIf(IsNull(vlRegistro!Gls_MatBen), "", (vlRegistro!Gls_MatBen))
      Lbl_PenNombre = vlNombre
      ''Fra_Poliza.Enabled = False
      flHabilitarIngreso
   
   Else
      Lbl_End = ""
      MsgBox "El Beneficiario/Pensionado No tiene Derecho a Pensión o No se encuentra registrado.", vbCritical, "Error de Datos"
      Txt_PenPoliza.SetFocus
   End If
   vlRegistro.Close
     
'I--- ABV 27/10/2007 ---
'    'Verifica si el Causante se encuentra fallecido
'    vgSql = "Select cod_estpension from pp_tmae_ben "
'    vgSql = vgSql & "Where cod_par='99' and "
'    vgSql = vgSql & "num_poliza = '" & Trim(Txt_PenPoliza) & "' "
'    vgSql = vgSql & "order by num_endoso desc"
'    Set vlRegistro = vgConexionBD.Execute(vgSql)
'    If Not vlRegistro.EOF Then
'      If (vlRegistro!Cod_EstPension <> "10") Then
'        MsgBox "El Causante de la Póliza No se encuentra Fallecido. Imposible realizar Operación.", vbCritical, "Causante No Fallecido"
'        Cmd_Cancelar_Click
'        Screen.MousePointer = 0
'        Exit Function
'      End If
'    End If
'    vlRegistro.Close
'F--- ABV 27/10/2007 ---
     
   If (Lbl_End <> "") And (Lbl_PenNombre <> "") And (Txt_PenPoliza <> "") And (Cmb_PenNumIdent <> "") And (Txt_PenNumIdent <> "") Then
       flCargarHistorico
       Cmb_TipoPersona.SetFocus
   End If
     
   Screen.MousePointer = 0
     
Exit Function
Err_Validar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flCargarHistorico()
On Error GoTo Err_Carga

  vgSql = ""
  vgSql = "SELECT num_poliza,num_endoso,cod_conpago,cod_tipoidensolicita,"
  vgSql = vgSql & "num_idensolicita,cod_tipopersona,gls_nomsolicita,gls_nomsegsolicita,"
  vgSql = vgSql & "gls_patsolicita,gls_matsolicita,gls_dirsolicita,cod_direccion,"
  vgSql = vgSql & "gls_fonosolicita,gls_correosolicita,fec_solpago,fec_pago,"
  vgSql = vgSql & "cod_tipodctopago,num_dctopago,mto_cobra,mto_pago,cod_tipoidenfun,"
  vgSql = vgSql & "num_idenfun,cod_viapago,cod_tipcuenta,cod_banco,num_cuenta,cod_sucursal, FEC_SOL_AFP "
  vgSql = vgSql & "FROM PP_TMAE_PAGTERCUOMOR "
  vgSql = vgSql & " Where "
  vgSql = vgSql & " NUM_POLIZA = '" & Trim(Txt_PenPoliza) & "' "
  Set vlRegistro = vgConexionBD.Execute(vgSql)
  If Not vlRegistro.EOF Then
     
        Cmb_TipoPersona.ListIndex = fgBuscarPosicionCodigoCombo(Trim(vlRegistro!COD_TIPOPERSONA), Cmb_TipoPersona)
        Cmb_NumIdentRep.ListIndex = fgBuscarPosicionCodigoCombo(Trim(vlRegistro!cod_tipoidensolicita), Cmb_NumIdentRep)
        Txt_NumIdentRep = Trim(vlRegistro!num_idensolicita)
        Txt_NombRecep = Trim(vlRegistro!gls_nomsolicita)
        Txt_NombSegRecep = IIf(IsNull(vlRegistro!gls_nomsegsolicita), "", Trim(vlRegistro!gls_nomsegsolicita))
        Txt_ApPaterno = Trim(vlRegistro!gls_patsolicita)
        Txt_ApMaterno = IIf(IsNull(vlRegistro!gls_matsolicita), "", Trim(vlRegistro!gls_matsolicita))
        Txt_DomicRecep = Trim(vlRegistro!gls_dirsolicita)
        
        vlCodDir = (vlRegistro!Cod_Direccion)
        Call fgBuscarNombreProvinciaRegion(vlCodDir)
        vlNombreRegion = vgNombreRegion
        vlNombreProvincia = vgNombreProvincia
        vlNombreComuna = vgNombreComuna
       
        Lbl_Departamento = vlNombreRegion
        Lbl_Provincia = vlNombreProvincia
        Lbl_Distrito = vlNombreComuna
        Txt_TelefRecep = IIf(IsNull(vlRegistro!gls_fonosolicita), "", Trim(vlRegistro!gls_fonosolicita))
        Txt_EmailRecep = IIf(IsNull(vlRegistro!gls_correosolicita), "", Trim(vlRegistro!gls_correosolicita))
        Txt_FechaRecep = DateSerial(Mid((vlRegistro!fec_solpago), 1, 4), Mid((vlRegistro!fec_solpago), 5, 2), Mid((vlRegistro!fec_solpago), 7, 2))
        
        If IsNull(vlRegistro!FEC_SOL_AFP) Then
            
            txtFecha_Solicitud = ""
        Else
              txtFecha_Solicitud = DateSerial(Mid((vlRegistro!FEC_SOL_AFP), 1, 4), Mid((vlRegistro!FEC_SOL_AFP), 5, 2), Mid((vlRegistro!FEC_SOL_AFP), 7, 2))
        
        End If
    
        Txt_FechaPgo = DateSerial(Mid((vlRegistro!Fec_Pago), 1, 4), Mid((vlRegistro!Fec_Pago), 5, 2), Mid((vlRegistro!Fec_Pago), 7, 2))
        Lbl_ValorCobrado = Format(Trim(vlRegistro!mto_cobra), "#,#0.00")
        Txt_ValorPagado = Format(Trim(vlRegistro!mto_pago), "#,#0.00")
        Cmb_TipoDctoPago.ListIndex = fgBuscarPosicionCodigoCombo(Trim(vlRegistro!cod_tipodctopago), Cmb_TipoDctoPago)
        Txt_NroFactura = Trim(vlRegistro!num_dctopago)
        Txt_Ruc = Trim(vlRegistro!num_idenfun)
        Cmb_ViaPago.ListIndex = fgBuscarPosicionCodigoCombo(Trim(vlRegistro!Cod_ViaPago), Cmb_ViaPago)
        Cmb_Sucursal.ListIndex = fgBuscarPosicionCodigoCombo(Trim(vlRegistro!Cod_Sucursal), Cmb_Sucursal)
        Cmb_TipCuenta.ListIndex = fgBuscarPosicionCodigoCombo(Trim(vlRegistro!Cod_TipCuenta), Cmb_TipCuenta)
        Cmb_Banco.ListIndex = fgBuscarPosicionCodigoCombo(Trim(vlRegistro!Cod_Banco), Cmb_Banco)
        Txt_NumCta = IIf(IsNull(vlRegistro!Num_Cuenta), "", Trim(vlRegistro!Num_Cuenta))

  End If
  vlRegistro.Close
     
  Screen.MousePointer = 0

Exit Function
Err_Carga:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flLimpia()
   vlPasa = True
   If Cmb_TipoPersona.Text <> "" Then
      Cmb_TipoPersona.ListIndex = 0
   End If
   If Cmb_NumIdentRep.Text <> "" Then
      Cmb_NumIdentRep.ListIndex = 0
   End If
   If Cmb_TipoDctoPago.Text <> "" Then
      Cmb_TipoDctoPago.ListIndex = 0
   End If
   If Cmb_ViaPago.Text <> "" Then
      Cmb_ViaPago.ListIndex = 0
   End If
   If Cmb_Sucursal.Text <> "" Then
      Cmb_Sucursal.ListIndex = 0
   End If
   If Cmb_TipCuenta.Text <> "" Then
    Cmb_TipCuenta.ListIndex = 0
   End If
   If Cmb_Banco.Text <> "" Then
    Cmb_Banco.ListIndex = 0
   End If
   
   Txt_NumIdentRep = ""
   Txt_NombRecep = ""
   Txt_NombSegRecep = ""
   Txt_ApPaterno = ""
   Txt_ApMaterno = ""
   Txt_DomicRecep = ""
   Lbl_Departamento = ""
   Lbl_Provincia = ""
   Lbl_Distrito = ""
   Txt_TelefRecep = ""
   Txt_EmailRecep = ""
   Txt_FechaRecep = ""
   Txt_FechaPgo = ""
   Lbl_ValorCobrado = ""
   Txt_ValorPagado = ""
   Txt_NroFactura = ""
   Txt_Ruc = ""
   Txt_NumCta = ""

End Function

Function flValidaFecha(iFecha)
On Error GoTo Err_valfecha
      flValidaFecha = False
     
     'valida que la fecha este correcta
      If Trim(iFecha <> "") Then
         If Not IsDate(iFecha) Then
                MsgBox "La Fecha ingresada no es una fecha válida.", vbCritical, "Dato Incorrecto"
                Exit Function
         End If

         If (Year(iFecha) < 1900) Then
             MsgBox "La Fecha ingresada es menor a la mínima que se puede ingresar (1900).", vbCritical, "Dato Incorrecto"
             Exit Function
         End If
         flValidaFecha = True
     End If

Exit Function
Err_valfecha:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(2) & Err.Description & " ]"
    End Select
End Function

Sub flImpresion()
Dim vlArchivo As String
Dim vlTIdent, vlNIdent, vlNom As String
Err.Clear
On Error GoTo Errores1
   
   Screen.MousePointer = 11
   
   vlArchivo = strRpt & "PP_Rpt_PensRegistroPagos.rpt"   '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
       MsgBox "Archivo de Reporte de Listados no se encuentra en el directorio de la Aplicación.", 16, "Archivo no encontrado"
       Screen.MousePointer = 0
       Exit Sub
   End If
  
    'Busca Información del causante
    vgSql = "Select cod_tipoidenben,num_idenben,gls_nomben,gls_nomsegben,gls_patben,"
    vgSql = vgSql & "gls_matben from pp_tmae_ben Where cod_par='99' and "
    vgSql = vgSql & "num_poliza = '" & Trim(Txt_PenPoliza) & "' "
    vgSql = vgSql & "order by num_endoso desc"
    Set vlRegistro = vgConexionBD.Execute(vgSql)
    If Not vlRegistro.EOF Then
        vlTIdent = fgBuscarNombreTipoIden(vlRegistro!Cod_TipoIdenBen, False)
        vlNIdent = vlRegistro!Num_IdenBen
        'I - MC 24/01/2008
        ''vlNom = vlRegistro!Gls_NomBen + " " + IIf(IsNull(vlRegistro!Gls_NomBen), "", (vlRegistro!Gls_NomBen)) + " " + vlRegistro!Gls_PatBen + " " + IIf(IsNull(vlRegistro!Gls_MatBen), "", (vlRegistro!Gls_MatBen))
        vlNom = vlRegistro!Gls_NomBen + " " + IIf(IsNull(vlRegistro!Gls_NomSegBen), "", (vlRegistro!Gls_NomSegBen)) + " " + vlRegistro!Gls_PatBen + " " + IIf(IsNull(vlRegistro!Gls_MatBen), "", (vlRegistro!Gls_MatBen))
        'F - MC 24/01/2008
    End If
    vlRegistro.Close
  
   vgQuery = ""
   vgQuery = "{PP_TMAE_PAGTERCUOMOR.NUM_POLIZA} = '" & Trim(Txt_PenPoliza) & "' AND "
   vgQuery = vgQuery & "{MA_TPAR_TABCODViaPago.COD_TABLA} = '" & vgCodTabla_ViaPago & "' AND "
   vgQuery = vgQuery & "{MA_TPAR_TABCODTipCta.COD_TABLA} = '" & vgCodTabla_TipCta & "' AND "
   vgQuery = vgQuery & "{MA_TPAR_TABCODCodBco.COD_TABLA} = '" & vgCodTabla_Bco & "' AND "
   vgQuery = vgQuery & "{MA_TPAR_TABCODTipPer.COD_TABLA} = '" & vgCodTabla_TipPer & "' "

   Rpt_Imprimir.Reset
   Rpt_Imprimir.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
   Rpt_Imprimir.Connect = vgRutaDataBase
   Rpt_Imprimir.SelectionFormula = vgQuery
   Rpt_Imprimir.Formulas(0) = ""
   Rpt_Imprimir.Formulas(1) = ""
   Rpt_Imprimir.Formulas(2) = ""
   Rpt_Imprimir.Formulas(3) = ""
   Rpt_Imprimir.Formulas(4) = ""
   Rpt_Imprimir.Formulas(5) = ""
   Rpt_Imprimir.Formulas(6) = ""
   
   Rpt_Imprimir.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
   Rpt_Imprimir.Formulas(1) = "NombreSistema= '" & vgNombreSistema & "'"
   Rpt_Imprimir.Formulas(2) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
   Ident = Trim(vlTIdent) + "-" + Trim(vlNIdent)
   Rpt_Imprimir.Formulas(3) = "Poliza = '" & Trim(Txt_PenPoliza) & "'"
   Rpt_Imprimir.Formulas(4) = "Endoso = '" & Trim(Lbl_End) & "'"
   Rpt_Imprimir.Formulas(5) = "Identificacion = '" & Trim(Ident) & "'"
   Rpt_Imprimir.Formulas(6) = "Nombre_Bene = '" & Trim(vlNom) & "'"
   
   Rpt_Imprimir.Destination = crptToWindow
   Rpt_Imprimir.WindowState = crptMaximized
   Rpt_Imprimir.WindowTitle = "Informe Pagos a Terceros Gastos de Sepelio"
   Rpt_Imprimir.Action = 1
   
   Screen.MousePointer = 0
   
Exit Sub
Errores1:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " " & "Numero de Error : " & Err.Number & Err.HelpFile & Err.HelpContext, vbExclamation, "Mensaje de error ..."
        Screen.MousePointer = 0
    End If
End Sub

Private Function ObtieneMtoCuoMor(iFecPago) As Double
'Busca el monto de cuota mortuoria a cobrar, según fecha de pago

    vgSql = ""
    vgSql = "select mto_cuomor from ma_tval_cuomor "
    vgSql = vgSql & "Where cod_moneda='" & vgMonedaCodOfi & "' and "
    vgSql = vgSql & "'" & iFecPago & "' between fec_inicuomor and fec_tercuomor "
    Set vlRegistro = vgConexionBD.Execute(vgSql)
    If Not vlRegistro.EOF Then
        ObtieneMtoCuoMor = Format(vlRegistro!mto_CUOMOR, "#0.00")
    End If
    vlRegistro.Close
End Function

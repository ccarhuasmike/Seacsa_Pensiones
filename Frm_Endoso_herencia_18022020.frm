VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Frm_EndosoHerencia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenedor de Endoso Herencia"
   ClientHeight    =   4035
   ClientLeft      =   4425
   ClientTop       =   1860
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   10215
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Datos de la Herencia"
      Height          =   1695
      Left            =   120
      TabIndex        =   147
      Top             =   1080
      Width           =   9975
      Begin VB.TextBox txtnum_mesgarres 
         Height          =   285
         Left            =   8550
         MaxLength       =   13
         TabIndex        =   150
         Text            =   "1"
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtmto_pensioncal 
         Height          =   285
         Left            =   4830
         MaxLength       =   13
         TabIndex        =   149
         Top             =   600
         Width           =   1785
      End
      Begin VB.TextBox txtgls_docjud 
         Height          =   285
         Left            =   4830
         MaxLength       =   300
         TabIndex        =   148
         Top             =   960
         Width           =   1785
      End
      Begin VB.Label lblMoneda 
         Height          =   255
         Left            =   4920
         TabIndex        =   154
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "Garantizado Restante (Meses)"
         Height          =   285
         Left            =   8520
         TabIndex        =   153
         Top             =   480
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label17 
         Caption         =   "Monto Liquidar"
         Height          =   225
         Left            =   2280
         TabIndex        =   152
         Top             =   570
         Width           =   1965
      End
      Begin VB.Label Label19 
         Caption         =   "Fecha Solicita"
         Height          =   315
         Left            =   2280
         TabIndex        =   151
         Top             =   930
         Width           =   1605
      End
   End
   Begin VB.Frame Fra_Operacion 
      Height          =   1035
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      Width           =   9975
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "Endoso"
         Enabled         =   0   'False
         Height          =   675
         Left            =   3360
         Picture         =   "Frm_Endoso_herencia.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Enabled         =   0   'False
         Height          =   675
         Left            =   4560
         Picture         =   "Frm_Endoso_herencia.frx":06BA
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Limpiar Formulario"
         Top             =   4440
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Modificar 
         Enabled         =   0   'False
         Height          =   675
         Left            =   1920
         Picture         =   "Frm_Endoso_herencia.frx":0D74
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Modificar lo Calculado"
         Top             =   4440
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Cancelar 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   675
         Left            =   4200
         Picture         =   "Frm_Endoso_herencia.frx":11B6
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5040
         Picture         =   "Frm_Endoso_herencia.frx":1790
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Grabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   675
         Left            =   2760
         Picture         =   "Frm_Endoso_herencia.frx":188A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Grabar Datos"
         Top             =   4440
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Eliminar 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Height          =   675
         Left            =   3720
         Picture         =   "Frm_Endoso_herencia.frx":1F44
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   4440
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Aprobar 
         Caption         =   "&Aprobar"
         Height          =   675
         Left            =   2520
         Picture         =   "Frm_Endoso_herencia.frx":2286
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Grabar Datos"
         Top             =   240
         Width           =   720
      End
      Begin Crystal.CrystalReport Rpt_Reporte 
         Left            =   120
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowState     =   2
         PrintFileLinesPerPage=   60
      End
      Begin VB.Line Lin_Borde 
         BorderColor     =   &H00808080&
         Index           =   2
         X1              =   240
         X2              =   960
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Lin_Borde 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   240
         X2              =   960
         Y1              =   1050
         Y2              =   1050
      End
   End
   Begin VB.Frame Fra_Poliza 
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
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   9975
      Begin VB.CommandButton Cmd_BuscarPendiente 
         Caption         =   "PRE"
         Height          =   680
         Left            =   9120
         Picture         =   "Frm_Endoso_herencia.frx":2940
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Buscar Póliza"
         Top             =   150
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Cmd_BuscarPol 
         Height          =   705
         Left            =   7920
         Picture         =   "Frm_Endoso_herencia.frx":2A42
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Buscar Póliza"
         Top             =   180
         Width           =   615
      End
      Begin VB.TextBox Txt_PenPoliza 
         Height          =   285
         Left            =   840
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1185
      End
      Begin VB.TextBox Txt_PenNumIdent 
         Height          =   285
         Left            =   4560
         MaxLength       =   16
         TabIndex        =   2
         Top             =   240
         Width           =   1635
      End
      Begin VB.ComboBox Cmb_PenNumIdent 
         BackColor       =   &H00E0FFFF&
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1635
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Póliza"
         Height          =   255
         Index           =   80
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Lbl_PenNombre 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   600
         Width           =   6975
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Ident."
         Height          =   255
         Index           =   0
         Left            =   2160
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "N° End. Actual"
         Height          =   195
         Index           =   42
         Left            =   6240
         TabIndex        =   6
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Lbl_EndosoActual 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   7320
         TabIndex        =   3
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label8 
         Caption         =   " Póliza que se desea Endosar"
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
         Left            =   120
         TabIndex        =   5
         Top             =   0
         Width           =   2655
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   43
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   855
      End
   End
   Begin TabDlg.SSTab SSTab_PolizaModificada 
      Height          =   1275
      Left            =   240
      TabIndex        =   22
      Top             =   6720
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   2249
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   128
      TabCaption(0)   =   "Póliza a Endosar"
      TabPicture(0)   =   "Frm_Endoso_herencia.frx":2B44
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Fra_PM"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Beneficiarios a Endosar"
      TabPicture(1)   =   "Frm_Endoso_herencia.frx":2B60
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Cmd_BMCalcular"
      Tab(1).Control(1)=   "Cmd_BMRestar"
      Tab(1).Control(2)=   "Cmd_BMSumar"
      Tab(1).Control(3)=   "Fra_BM"
      Tab(1).Control(4)=   "frmDatosGen"
      Tab(1).Control(5)=   "Msf_BMGrilla"
      Tab(1).ControlCount=   6
      Begin VB.CommandButton Cmd_BMCalcular 
         Enabled         =   0   'False
         Height          =   450
         Left            =   -66300
         Picture         =   "Frm_Endoso_herencia.frx":2B7C
         Style           =   1  'Graphical
         TabIndex        =   141
         ToolTipText     =   "Calcular Porcentajes"
         Top             =   3120
         Width           =   495
      End
      Begin VB.CommandButton Cmd_BMRestar 
         Enabled         =   0   'False
         Height          =   450
         Left            =   -66300
         Picture         =   "Frm_Endoso_herencia.frx":313E
         Style           =   1  'Graphical
         TabIndex        =   140
         ToolTipText     =   "Quitar Beneficiario"
         Top             =   2640
         Width           =   495
      End
      Begin VB.CommandButton Cmd_BMSumar 
         Enabled         =   0   'False
         Height          =   450
         Left            =   -66300
         Picture         =   "Frm_Endoso_herencia.frx":32C8
         Style           =   1  'Graphical
         TabIndex        =   139
         ToolTipText     =   "Agregar Beneficiario"
         Top             =   2160
         Width           =   495
      End
      Begin VB.Frame Fra_BM 
         Caption         =   "  Antecedentes Personales  "
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
         Height          =   2295
         Left            =   -74880
         TabIndex        =   109
         Top             =   3720
         Width           =   9135
         Begin VB.CheckBox chkContinuarpago 
            Caption         =   "Continuar el pago"
            Enabled         =   0   'False
            Height          =   255
            Left            =   4800
            TabIndex        =   143
            Top             =   270
            Width           =   2325
         End
         Begin VB.ComboBox cmbExcluesion 
            BackColor       =   &H00E0FFFF&
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "Frm_Endoso_herencia.frx":3452
            Left            =   1080
            List            =   "Frm_Endoso_herencia.frx":345F
            Style           =   2  'Dropdown List
            TabIndex        =   122
            Top             =   1260
            Width           =   3135
         End
         Begin VB.TextBox txt_BMPrcGar 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3480
            TabIndex        =   121
            Top             =   1905
            Width           =   735
         End
         Begin VB.CommandButton cmdIngNombres 
            Caption         =   "Nombres >>"
            Height          =   375
            Left            =   8640
            TabIndex        =   120
            Top             =   2880
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.ComboBox Cmb_BMDerPen 
            BackColor       =   &H00E0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   119
            Top             =   915
            Width           =   3135
         End
         Begin VB.ComboBox Cmb_BMSexo 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   118
            Top             =   570
            Width           =   3135
         End
         Begin VB.TextBox Txt_BMNomBen 
            Height          =   285
            Left            =   1080
            MaxLength       =   25
            TabIndex        =   117
            Top             =   2760
            Width           =   3135
         End
         Begin VB.TextBox Txt_BMNomSegBen 
            Height          =   285
            Left            =   1080
            MaxLength       =   25
            TabIndex        =   116
            Top             =   3120
            Width           =   3135
         End
         Begin VB.TextBox Txt_BMFecMat 
            Height          =   285
            Left            =   5760
            MaxLength       =   10
            TabIndex        =   115
            Top             =   2520
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox Txt_BMMtoPensionGar 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            TabIndex        =   114
            Top             =   1905
            Width           =   1095
         End
         Begin VB.TextBox Txt_BMFecNac 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   113
            Top             =   1590
            Width           =   1095
         End
         Begin VB.TextBox Txt_BMMatBen 
            Height          =   285
            Left            =   5280
            MaxLength       =   20
            TabIndex        =   112
            Top             =   3120
            Width           =   3135
         End
         Begin VB.TextBox Txt_BMPatBen 
            Height          =   285
            Left            =   5280
            MaxLength       =   20
            TabIndex        =   111
            Top             =   2760
            Width           =   3135
         End
         Begin VB.ComboBox Cmb_BMPar 
            BackColor       =   &H00E0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   110
            Top             =   240
            Width           =   3135
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Exclusion"
            Height          =   255
            Index           =   103
            Left            =   120
            TabIndex        =   138
            Top             =   1290
            Width           =   975
         End
         Begin VB.Line Line3 
            X1              =   4560
            X2              =   4560
            Y1              =   240
            Y2              =   2040
         End
         Begin VB.Label lbl_BMPrcGar 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   3480
            TabIndex        =   137
            Top             =   1905
            Width           =   735
         End
         Begin VB.Label Label18 
            Caption         =   "% Pen. Gar."
            Height          =   255
            Left            =   2550
            TabIndex        =   136
            Top             =   1935
            Width           =   855
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2do. Nombre"
            Height          =   195
            Index           =   69
            Left            =   120
            TabIndex        =   135
            Top             =   3120
            Width           =   915
         End
         Begin VB.Label Lbl_MonPension 
            Caption         =   "(TM)"
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   2220
            TabIndex        =   134
            Top             =   1920
            Width           =   495
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pensión Gar."
            Height          =   195
            Index           =   91
            Left            =   120
            TabIndex        =   133
            Top             =   1905
            Width           =   1035
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Mat."
            Height          =   195
            Index           =   90
            Left            =   4920
            TabIndex        =   132
            Top             =   2640
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Ap. Materno"
            Height          =   255
            Index           =   86
            Left            =   4320
            TabIndex        =   131
            Top             =   3120
            Width           =   915
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Ap. Paterno"
            Height          =   255
            Index           =   85
            Left            =   4320
            TabIndex        =   130
            Top             =   2760
            Width           =   915
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1er. Nombre"
            Height          =   195
            Index           =   84
            Left            =   120
            TabIndex        =   129
            Top             =   2760
            Width           =   870
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Nac."
            Height          =   195
            Index           =   74
            Left            =   120
            TabIndex        =   128
            Top             =   1620
            Width           =   840
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Parentesco"
            Height          =   255
            Index           =   72
            Left            =   120
            TabIndex        =   127
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sexo"
            Height          =   255
            Index           =   67
            Left            =   120
            TabIndex        =   126
            Top             =   540
            Width           =   975
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Dº Pensión"
            Height          =   255
            Index           =   66
            Left            =   120
            TabIndex        =   125
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Lbl_Nombre 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Dº Crecer"
            Height          =   255
            Index           =   65
            Left            =   3480
            TabIndex        =   124
            Top             =   2520
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Lbl_BMDerAcrecer 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   4320
            TabIndex        =   123
            Top             =   2520
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.Frame Fra_PM 
         Caption         =   " Antecedentes de la Póliza "
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
         Height          =   5115
         Left            =   330
         TabIndex        =   62
         Top             =   660
         Width           =   9375
         Begin VB.ComboBox Cmb_PMTipPen 
            BackColor       =   &H00E0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   73
            Top             =   1080
            Width           =   3375
         End
         Begin VB.TextBox Txt_PMIniVig 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   72
            Top             =   2160
            Width           =   1065
         End
         Begin VB.ComboBox Cmb_PMEstVig 
            BackColor       =   &H00E0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   71
            Top             =   1800
            Width           =   2775
         End
         Begin VB.ComboBox Cmb_PMTipRta 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   70
            Top             =   2520
            Width           =   2775
         End
         Begin VB.ComboBox Cmb_PMMod 
            BackColor       =   &H00E0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1650
            Style           =   2  'Dropdown List
            TabIndex        =   69
            Top             =   3240
            Width           =   2655
         End
         Begin VB.TextBox Txt_PMMesDif 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            MaxLength       =   3
            TabIndex        =   68
            Top             =   2880
            Width           =   585
         End
         Begin VB.TextBox Txt_PMMesGar 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            MaxLength       =   3
            TabIndex        =   67
            Top             =   3600
            Width           =   585
         End
         Begin VB.TextBox Txt_PMTasaCto 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Left            =   6720
            MaxLength       =   6
            TabIndex        =   66
            Top             =   1800
            Width           =   585
         End
         Begin VB.TextBox Txt_PMTasaVta 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Left            =   6720
            MaxLength       =   6
            TabIndex        =   65
            Top             =   2160
            Width           =   585
         End
         Begin VB.TextBox Txt_PMMtoPrima 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Left            =   6720
            MaxLength       =   13
            TabIndex        =   64
            Top             =   1440
            Width           =   1185
         End
         Begin VB.TextBox Txt_PMTerVig 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Left            =   3000
            MaxLength       =   10
            TabIndex        =   63
            Top             =   2160
            Width           =   1065
         End
         Begin VB.Label lblTC 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            TabIndex        =   108
            Top             =   4320
            Width           =   735
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "TC"
            Height          =   255
            Index           =   110
            Left            =   240
            TabIndex        =   107
            Top             =   4320
            Width           =   1095
         End
         Begin VB.Label Lbl_PMPrcFacPenElla 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   7560
            TabIndex        =   106
            Top             =   3240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Lbl_PMMtoFacPenElla 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   7560
            TabIndex        =   105
            Top             =   2880
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Lbl_PMDerGratificacion 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   6720
            TabIndex        =   104
            Top             =   3600
            Width           =   735
         End
         Begin VB.Label Lbl_PMDerCrecer 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   6720
            TabIndex        =   103
            Top             =   3240
            Width           =   735
         End
         Begin VB.Label Lbl_PMCoberCon 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   6720
            TabIndex        =   102
            Top             =   2880
            Width           =   735
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Gratificación"
            Height          =   255
            Index           =   102
            Left            =   5160
            TabIndex        =   101
            Top             =   3600
            Width           =   1575
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Derecho Crecer"
            Height          =   255
            Index           =   101
            Left            =   5160
            TabIndex        =   100
            Top             =   3240
            Width           =   1575
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Cob. Cónyuge"
            Height          =   255
            Index           =   100
            Left            =   5160
            TabIndex        =   99
            Top             =   2880
            Width           =   1575
         End
         Begin VB.Label Lbl_PMIndCobertura 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            TabIndex        =   98
            Top             =   3960
            Width           =   735
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Ind. Cobertura"
            Height          =   255
            Index           =   99
            Left            =   240
            TabIndex        =   97
            Top             =   3960
            Width           =   1335
         End
         Begin VB.Label Lbl_PMMonedaPrima 
            Caption         =   "S/."
            Height          =   285
            Left            =   7920
            TabIndex        =   96
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label Lbl_PMFecDevengue 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   6720
            TabIndex        =   95
            Top             =   2520
            Width           =   1095
         End
         Begin VB.Label Lbl_PMFecEmision 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   6720
            TabIndex        =   94
            Top             =   3960
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Fecha de Devengue"
            Height          =   255
            Index           =   94
            Left            =   5160
            TabIndex        =   93
            Top             =   2520
            Width           =   1575
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Fecha de Emisión"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   93
            Left            =   5160
            TabIndex        =   92
            Top             =   3960
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Tipo de Pensión "
            Height          =   255
            Index           =   21
            Left            =   240
            TabIndex        =   91
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Nº Beneficiarios"
            Height          =   255
            Index           =   22
            Left            =   240
            TabIndex        =   90
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Lbl_Nombre 
            Alignment       =   2  'Center
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   23
            Left            =   2760
            TabIndex        =   89
            Top             =   1560
            Width           =   255
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Periodo Vigencia"
            Height          =   255
            Index           =   37
            Left            =   240
            TabIndex        =   88
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Estado Vigencia"
            Height          =   255
            Index           =   44
            Left            =   240
            TabIndex        =   87
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Tipo de Renta"
            Height          =   255
            Index           =   49
            Left            =   240
            TabIndex        =   86
            Top             =   2520
            Width           =   1455
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Meses Gar."
            Height          =   255
            Index           =   50
            Left            =   240
            TabIndex        =   85
            Top             =   3600
            Width           =   975
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Meses Dif."
            Height          =   255
            Index           =   51
            Left            =   240
            TabIndex        =   84
            Top             =   2880
            Width           =   1095
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Modalidad"
            Height          =   255
            Index           =   52
            Left            =   240
            TabIndex        =   83
            Top             =   3240
            Width           =   1380
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Tasa de Venta"
            Height          =   255
            Index           =   54
            Left            =   5160
            TabIndex        =   82
            Top             =   2160
            Width           =   1095
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Tasa Cto. Equiv."
            Height          =   255
            Index           =   55
            Left            =   5160
            TabIndex        =   81
            Top             =   1815
            Width           =   1335
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Monto Prima"
            Height          =   255
            Index           =   57
            Left            =   5160
            TabIndex        =   80
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "%"
            Height          =   255
            Index           =   58
            Left            =   7440
            TabIndex        =   79
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "%"
            Height          =   255
            Index           =   59
            Left            =   7440
            TabIndex        =   78
            Top             =   2160
            Width           =   375
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "%"
            Height          =   255
            Index           =   60
            Left            =   4560
            TabIndex        =   77
            Top             =   2640
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Lbl_PMNumCar 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            TabIndex        =   76
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Porc. Cubierto"
            Height          =   255
            Index           =   47
            Left            =   2400
            TabIndex        =   75
            Top             =   3720
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "%"
            Height          =   255
            Index           =   48
            Left            =   2760
            TabIndex        =   74
            Top             =   3600
            Visible         =   0   'False
            Width           =   375
         End
      End
      Begin VB.Frame frmDatosGen 
         Caption         =   "Datos Generales"
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
         Height          =   1695
         Left            =   -74880
         TabIndex        =   23
         Top             =   2040
         Width           =   8535
         Begin VB.TextBox Txt_BMNumIden 
            Enabled         =   0   'False
            Height          =   285
            Left            =   6360
            MaxLength       =   10
            TabIndex        =   42
            Top             =   195
            Width           =   1815
         End
         Begin VB.ComboBox Cmb_BMNumIdent 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   4920
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   180
            Width           =   1395
         End
         Begin VB.TextBox txtBenEmail 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4920
            TabIndex        =   40
            Top             =   765
            Width           =   3255
         End
         Begin VB.TextBox txtDistritoBen 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4920
            TabIndex        =   39
            Top             =   1350
            Width           =   3255
         End
         Begin VB.TextBox txtcodDirBen 
            Height          =   285
            Left            =   3960
            TabIndex        =   38
            Top             =   2760
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton Cmd_BuscarDir 
            Caption         =   "?"
            Enabled         =   0   'False
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
            Left            =   8220
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Efectuar Busqueda de Dirección"
            Top             =   1350
            Width           =   300
         End
         Begin VB.TextBox txtTelfon2 
            Enabled         =   0   'False
            Height          =   285
            Left            =   6840
            TabIndex        =   36
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtFecdevBen 
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   8880
            TabIndex        =   35
            Top             =   1920
            Width           =   1215
         End
         Begin VB.ComboBox cmbSexoBen 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   8760
            TabIndex        =   34
            Top             =   2880
            Width           =   1575
         End
         Begin VB.TextBox txtFecnacBen 
            Height          =   285
            Left            =   4440
            TabIndex        =   33
            Top             =   3240
            Width           =   1215
         End
         Begin VB.ComboBox cmbTipoPreBen 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   8880
            TabIndex        =   32
            Top             =   3480
            Width           =   3015
         End
         Begin VB.ComboBox cmbAfpBen 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   5760
            TabIndex        =   31
            Top             =   2760
            Width           =   2655
         End
         Begin VB.TextBox txtCusspBen 
            Height          =   285
            Left            =   6360
            TabIndex        =   30
            Top             =   3240
            Width           =   1335
         End
         Begin VB.TextBox txtTelBen 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4920
            TabIndex        =   29
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txt_dirben 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4920
            MaxLength       =   120
            TabIndex        =   28
            Top             =   1050
            Width           =   3255
         End
         Begin VB.TextBox txtApematBen 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   27
            Top             =   1350
            Width           =   2985
         End
         Begin VB.TextBox txtApepatBen 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   26
            Top             =   1050
            Width           =   2985
         End
         Begin VB.TextBox txtNomsegBen 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   25
            Top             =   765
            Width           =   2985
         End
         Begin VB.TextBox txtNomben 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   24
            Top             =   480
            Width           =   2985
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            Caption         =   "Nº Endoso a Crear"
            Height          =   195
            Index           =   79
            Left            =   2040
            TabIndex        =   145
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label Lbl_EndCrear 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   3480
            TabIndex        =   144
            Top             =   180
            Width           =   600
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Orden"
            Height          =   255
            Index           =   71
            Left            =   120
            TabIndex        =   61
            Top             =   195
            Width           =   975
         End
         Begin VB.Label Lbl_BMNumOrd 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            TabIndex        =   60
            Top             =   195
            Width           =   495
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Ident."
            Height          =   255
            Index           =   68
            Left            =   4200
            TabIndex        =   59
            Top             =   195
            Width           =   735
         End
         Begin VB.Label Label15 
            Caption         =   "E-mail"
            Height          =   255
            Left            =   4200
            TabIndex        =   58
            Top             =   765
            Width           =   615
         End
         Begin VB.Label Lbl_Distrito 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   120
            TabIndex        =   57
            Top             =   3240
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Distrito"
            Height          =   255
            Index           =   12
            Left            =   4200
            TabIndex        =   56
            Top             =   1350
            Width           =   615
         End
         Begin VB.Label Label14 
            Caption         =   "/"
            Height          =   255
            Left            =   6480
            TabIndex        =   55
            Top             =   480
            Width           =   255
         End
         Begin VB.Line Line2 
            X1              =   120
            X2              =   8400
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   8400
            Y1              =   2400
            Y2              =   2400
         End
         Begin VB.Label Label13 
            Caption         =   "Fech.Dev."
            Height          =   255
            Left            =   8760
            TabIndex        =   54
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label Label12 
            Caption         =   "Sexo"
            Height          =   255
            Left            =   8880
            TabIndex        =   53
            Top             =   2400
            Width           =   615
         End
         Begin VB.Label Label11 
            Caption         =   "Fech.Nac."
            Height          =   255
            Left            =   3600
            TabIndex        =   52
            Top             =   3240
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "Tipo de Pens."
            Height          =   255
            Left            =   9000
            TabIndex        =   51
            Top             =   3480
            Width           =   1095
         End
         Begin VB.Label Label9 
            Caption         =   "AFP"
            Height          =   255
            Left            =   5400
            TabIndex        =   50
            Top             =   2760
            Width           =   495
         End
         Begin VB.Label Label7 
            Caption         =   "Cuspp"
            Height          =   255
            Left            =   5640
            TabIndex        =   49
            Top             =   3240
            Width           =   855
         End
         Begin VB.Label Label6 
            Caption         =   "Telf 1."
            Height          =   255
            Left            =   4200
            TabIndex        =   48
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label5 
            Caption         =   "Direccion"
            Height          =   255
            Left            =   4200
            TabIndex        =   47
            Top             =   1050
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Apellido Mat."
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   1350
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Apellido Pat."
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   1050
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Seg. Nom"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   765
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Nombre "
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   480
            Width           =   1095
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Msf_BMGrilla 
         Height          =   1635
         Left            =   -74880
         TabIndex        =   142
         Top             =   360
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   2884
         _Version        =   393216
         Rows            =   1
         Cols            =   6
         FixedCols       =   0
         BackColor       =   14745599
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Label lblRowsAfecctedBen 
      Caption         =   "lblRowsAfecctedBen"
      Height          =   345
      Left            =   3240
      TabIndex        =   146
      Top             =   3000
      Width           =   705
   End
End
Attribute VB_Name = "Frm_EndosoHerencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim vl_click_event As String
Dim vgPalabra As String
'Parámetros de Impresión
Dim vlNombre As String, vlNombreSeg As String
Dim vlPaterno As String, vlMaterno As String
Dim vlNombreTipoIden As String
Dim vlImpCodMoneda As String, vlImpNomMoneda As String
Dim vlMantPerGar As String
Dim vlGlobalNumPoliza As String
Dim vlGlobalNumEndoso As String
Dim vlRptNumEndosoPol As String
Dim vlRptNumEndosoEnd As String
' Variables Beneficiario
Dim vlnum_endoso_actual As Integer
Dim vlNum_Poliza  As String
Dim vlnum_endoso  As Integer
Dim vlnum_orden  As Integer
Dim vlnum_idenben  As String
Dim vlcod_tipoidenben  As String
Dim vlgls_nomben  As String
Dim vlgls_nomsegben  As String
Dim vlgls_patben  As String
Dim vlgls_matben  As String
Dim vlcod_grufam  As String
Dim vlcod_par  As String
Dim vlcod_sexo  As String
Dim vlcod_sitinv  As String
Dim vlcod_dercre  As String
Dim vlcod_derpen  As String
Dim vlfec_nacben  As String
Dim vlmto_pension As String
Dim vlmto_pensiongar  As String
Dim vlcod_estpension  As String
Dim vlfec_inipagopen  As String
Dim vlfec_terpagopengar  As String
Dim vlprc_pensiongar  As String
Dim vlgls_telben2  As String
Dim vlcod_direccion  As String
Dim vlgls_dirben  As String
Dim vlgls_fonoben  As String
Dim vlgls_correoben  As String

Dim vlcod_cauinv As Integer
Dim vlcod_motreqpen As String

Dim vlprc_pension As String

'Variables endoso
Dim vlfec_solendoso As String
Dim vlfec_endoso As String
Dim vlcod_cauendoso As String
Dim vlcod_tipendoso As String
Dim vlmto_diferencia As Double
Dim vlmto_pensionori As Double
Dim vlmto_pensioncal As Double
Dim vlfec_efecto As String
Dim vlprc_factor As String
Dim vlgls_observacion As String
Dim vlcod_usuariocrea As String
Dim vlfec_crea As String
Dim vlhor_crea As String
Dim vlcod_usuariomodi As String
Dim vlfec_modi As String
Dim vlhor_modi As String
Dim vlfec_finefecto As String
Dim vlcod_tipreajuste As String
Dim vlcod_estado_endoso As String

'begin---  variables auxialeres ----
Dim vlxcod_tipoidenben As String
Dim vlxnum_idenben As String
Dim vlxgls_nomben As String
Dim vlxgls_nomsegben As String
Dim vlxgls_patben As String
Dim vlxgls_matben As String
Dim vlxcod_direccion As String
Dim vlxgls_dirben As String
'end---------------------------------

Dim vlnum_mesgarres As String
Dim vlgls_docjud As String

Dim vg_tipo_busqueda As String
' variable de poliza
Dim vlcod_moneda As String

Dim vgSql As String
    
Function flDefineColumnsGridBeneficiario(iGrilla As MSFlexGrid)

On Error GoTo Err_flInicializaGrillaBenef
    
    iGrilla.Clear
    'mvg 20170904
    iGrilla.Cols = 25
    iGrilla.rows = 1
    iGrilla.RowHeight(0) = 250
    iGrilla.row = 0

    '----------------------------------------
    iGrilla.Col = 0
    iGrilla.Text = "Nº Orden"
    iGrilla.ColWidth(0) = 700

    iGrilla.Col = 1
    iGrilla.Text = "Tipo Ident."
    iGrilla.ColWidth(1) = 1000

    iGrilla.Col = 2
    iGrilla.Text = "N° Ident."
    iGrilla.ColWidth(2) = 1200

    iGrilla.Col = 3
    iGrilla.Text = "Nombre"
    iGrilla.ColWidth(3) = 1500

    iGrilla.Col = 4
    iGrilla.Text = " 2do. Nombre"
    iGrilla.ColWidth(4) = 1500

    iGrilla.Col = 5
    iGrilla.Text = "Ap. Paterno"
    iGrilla.ColWidth(5) = 1000

    iGrilla.Col = 6
    iGrilla.Text = "Ap. Materno"
    iGrilla.ColWidth(6) = 1000
    
    iGrilla.Col = 7
    iGrilla.Text = "Gru.Fam."
    iGrilla.ColWidth(7) = 700

    iGrilla.Col = 8
    iGrilla.Text = "Par."
    iGrilla.ColWidth(8) = 500

    iGrilla.Col = 9
    iGrilla.Text = "Sexo"
    iGrilla.ColWidth(9) = 500

    iGrilla.Col = 10
    iGrilla.Text = "Sit. Inv."
    iGrilla.ColWidth(10) = 600

    iGrilla.Col = 11
    iGrilla.Text = "Dº Crecer"
    iGrilla.ColWidth(11) = 800
    
    iGrilla.Col = 12
    iGrilla.Text = "Dº Pen." 'cod_estpension
    iGrilla.ColWidth(12) = 600
    
    iGrilla.Col = 13
    iGrilla.Text = "Fec.Nac."
    iGrilla.ColWidth(13) = 1000
    
    iGrilla.Col = 14
    iGrilla.Text = "Mto.Pensión gar"
    iGrilla.ColWidth(14) = 1000
    
    iGrilla.Col = 15
    iGrilla.Text = "cod EstPension"
    iGrilla.ColWidth(15) = 1000
    
    iGrilla.Col = 16
    iGrilla.Text = "Fec. IniPagoPen"
    iGrilla.ColWidth(16) = 1000
    
    iGrilla.Col = 17
    iGrilla.Text = "Fec. TerPagoPenGar"
    iGrilla.ColWidth(17) = 1000
    
    iGrilla.Col = 18
    iGrilla.Text = "Prc. PensionGar"
    iGrilla.ColWidth(18) = 1000
    
    iGrilla.Col = 19
    iGrilla.Text = "Gls. Telben2"
    iGrilla.ColWidth(19) = 1000
    
    iGrilla.Col = 20
    iGrilla.Text = "hubigeo" ' cod_direccion
    iGrilla.ColWidth(20) = 1000
    
    iGrilla.Col = 21
    iGrilla.Text = "direccion" ' gls_dirben
    iGrilla.ColWidth(21) = 1500
    
    iGrilla.Col = 22
    iGrilla.Text = "Telefono" ' gls_fonoben
    iGrilla.ColWidth(22) = 1000
    
    iGrilla.Col = 23
    iGrilla.Text = "Correo" ' gls_correoben
    iGrilla.ColWidth(23) = 1500
    
     iGrilla.Col = 24
    iGrilla.Text = "Mto. Pension" ' gls_correoben
    iGrilla.ColWidth(24) = 1500
  

Exit Function
Err_flInicializaGrillaBenef:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Private Sub chkContinuarpago_Click()
    If chkContinuarpago.Value = 1 Then
     cmbExcluesion.ListIndex = fgBuscarPosicionCodigoCombo(Trim("10"), cmbExcluesion)
       Else
     cmbExcluesion.ListIndex = fgBuscarPosicionCodigoCombo(Trim("99"), cmbExcluesion)
    End If
End Sub



Private Sub Cmd_Aprobar_Click()

If Not IsNumeric(txtmto_pensioncal.Text) Then
    MsgBox "Debe ingresar el Monto liquidar. Edite y realice los cambios", vbCritical, "Error de Datos"
    Exit Sub
End If

If Val(txtmto_pensioncal.Text) < 1 Then
    MsgBox "El Monto liquidar debe ser mayor a cero. Edite y realice los cambios", vbCritical, "Error de Datos"
    Exit Sub
End If

If Not IsNumeric(txtnum_mesgarres.Text) Then
    MsgBox "Debe ingresar mes garantizado restante. Edite y realice los cambios", vbCritical, "Error de Datos"
    Exit Sub
End If

If txtgls_docjud.Text = "" Then
 MsgBox "Debe ingresar número de documento judicial. Edite y realice los cambios", vbCritical, "Error de Datos"
    txtgls_docjud.SetFocus
    Exit Sub
End If



 If Msf_BMGrilla.rows > 1 Then
     vlmto_pension = Val(txtmto_pensioncal.Text) / (Msf_BMGrilla.rows - 1)
 Else
     vlmto_pension = Val(txtmto_pensioncal.Text)
 End If

'Dim resp As String
'If Msf_BMGrilla.rows > 1 Then
' vlPos = 1
' Msf_BMGrilla.Col = 24
' While vlPos <= (Msf_BMGrilla.rows - 1)
'     Msf_BMGrilla.row = vlPos
'     Msf_BMGrilla.Col = 0
'     resp = Msf_BMGrilla.TextMatrix(vlPos, 24)
'     If resp = "0" Or resp = "" Or vlmto_pension <> resp Then
'          MsgBox "Para realizar la aprobación, la pension del beneficiario debe tener la proporcion igualitaria sobre el monto a liquidar, por favor presione el boton modificar y realice el calculo.", vbCritical, "Proceso Cancelado"
'     Exit Sub
'     End If
'
'     vlPos = vlPos + 1
' Wend
' End If
 
vgRes = MsgBox(" ¿ Está seguro que desea Aprobar Definitivamente. ?", 4 + 32 + 256, "Operación de Aprobación")
If vgRes <> 6 Then
    Cmd_Salir.SetFocus
    Screen.MousePointer = 0
    Exit Sub
End If
   
     
 If (fl_aprobar = True) Then
        MsgBox "El Procedimiento de Aprobación se realizó Correctamente.", vbInformation, "Proceso Finalizado"
        Cmd_Modificar.Enabled = False
        Cmd_Grabar.Enabled = False
        Cmd_Aprobar.Enabled = False
        Cmd_Limpiar.Enabled = False
        Cmd_BMCalcular.Enabled = False
 Else
        MsgBox "El Procedimiento de Aprobación no ha podido realizarse.", vbCritical, "Proceso Cancelado"
 End If
End Sub

Private Sub Cmd_BMCalcular_Click()
    
    If Msf_BMGrilla.rows <= 1 Then
         MsgBox "Debe ingresar un beneficiario.", vbCritical, "Error de Datos"
         Exit Sub
     End If
 
    Txt_BMMtoPensionGar = Trim(Txt_BMMtoPensionGar)
    If Txt_BMMtoPensionGar.Text = "" Then
        Txt_BMMtoPensionGar.Text = "0"
    End If
    
    txt_BMPrcGar = Trim(txt_BMPrcGar)
    If txt_BMPrcGar.Text = "" Then
        txt_BMPrcGar.Text = "0"
    End If
    
     If Not IsNumeric(txtmto_pensioncal.Text) Then
         MsgBox "Debe ingresar el Monto liquidar.", vbCritical, "Error de Datos"
         txtmto_pensioncal.SetFocus
         Exit Sub
     End If
     
     If Val(txtmto_pensioncal.Text) < 1 Then
        MsgBox "El monto a liquidar debe ser mayor a cero.", vbCritical, "Error de Datos"
         txtmto_pensioncal.SetFocus
         Exit Sub
     End If
     
     If Msf_BMGrilla.rows > 1 Then
     vlmto_pension = Val(txtmto_pensioncal.Text) / (Msf_BMGrilla.rows - 1)
     Else
     vlmto_pension = Val(txtmto_pensioncal.Text)
     End If

    If Msf_BMGrilla.rows > 1 Then
     vlPos = 1
     Msf_BMGrilla.Col = 24
     While vlPos <= (Msf_BMGrilla.rows - 1)
         Msf_BMGrilla.row = vlPos
         Msf_BMGrilla.Col = 0
         Msf_BMGrilla.TextMatrix(vlPos, 24) = vlmto_pension
         Msf_BMGrilla.TextMatrix(vlPos, 14) = Txt_BMMtoPensionGar.Text ' Mto.Pensión gar
         Msf_BMGrilla.TextMatrix(vlPos, 18) = txt_BMPrcGar.Text 'Prc. PensionGar
         vlPos = vlPos + 1
     Wend
     End If

End Sub

Private Sub Cmd_BMRestar_Click()
 
      
      vgRes = MsgBox(" ¿ Está seguro que desea Eliminar este Beneficiario ?", 4 + 32 + 256, "Operación de Ingreso")
        If vgRes <> 6 Then
            Screen.MousePointer = 0
            Exit Sub
        End If
           
    
    If Msf_BMGrilla.rows <= 2 Then
     Call flDefineColumnsGridBeneficiario(Msf_BMGrilla)
     Lbl_BMNumOrd = ""
     Exit Sub
    End If
   
        
    Msf_BMGrilla.RemoveItem Msf_BMGrilla.row
    
    'ReAsignar Números de Orden a Registros Nuevos
    Msf_BMGrilla.row = 1
    Msf_BMGrilla.Col = 0
    vlPos = Msf_BMGrilla.row
    While vlPos <= Msf_BMGrilla.rows - 1
        If Trim(Msf_BMGrilla.Text) <> vlPos Then
            'Si el número de línea es distinto al número de orden
            Msf_BMGrilla.Text = vlPos
        End If
        vlPos = vlPos + 1
        If Msf_BMGrilla.row < Msf_BMGrilla.rows - 1 Then
            Msf_BMGrilla.row = vlPos
        End If
    Wend
    
    Msf_BMGrilla.row = 1
    
End Sub

Private Sub Cmd_BMSumar_Click()

 'Validar Código de Identificación del Beneficiario
    If (Cmb_BMNumIdent.Text) = "" Then
        MsgBox "Debe seleccionar el Tipo de Identificación del Beneficiario.", vbCritical, "Error de Datos"
        Cmb_BMNumIdent.SetFocus
        Exit Sub
    End If
    'Validar Número de Identificación del Beneficiario
    Txt_BMNumIden = Trim(UCase(Txt_BMNumIden))
    If Txt_BMNumIden.Text = "" Then
       MsgBox "Debe Ingresar el Número de Identificación del Beneficiario.", vbCritical, "Error de Datos"
       Txt_BMNumIden.SetFocus
       Exit Sub
    End If
    
    'Valida Nombre del Beneficiario
    txtNomben.Text = Trim(UCase(txtNomben.Text))
    If txtNomben.Text = "" Then
       MsgBox "Debe Ingresar Nombre.", vbCritical, "Error de Datos"
       'frmDatosGen.Visible = True
       txtNomben.SetFocus
       Exit Sub
    End If
    
    'Valida Apellido Paterno del Beneficiario
    txtApepatBen.Text = Trim(UCase(txtApepatBen.Text))
    If txtApepatBen.Text = "" Then
       MsgBox "Debe Ingresar Apellido Paterno.", vbCritical, "Error de Datos"
       txtApepatBen.SetFocus
       vlSwCalIntOK = False
       Exit Sub
    End If
    
     'Valida direccion  del Beneficiario
    txt_dirben.Text = Trim(UCase(txt_dirben.Text))
    If txt_dirben.Text = "" Then
       MsgBox "Debe Ingresar direccion.", vbCritical, "Error de Datos"
       txt_dirben.SetFocus
       vlSwCalIntOK = False
       Exit Sub
    End If
    
    'Valida distrito del Beneficiario
    txtDistritoBen.Text = Trim(UCase(txtDistritoBen.Text))
    If txtDistritoBen.Text = "" Then
       MsgBox "Debe Ingresar distrito.", vbCritical, "Error de Datos"
       txtDistritoBen.SetFocus
       vlSwCalIntOK = False
       Exit Sub
    End If
    
    'Valida Fecha de Nacimiento del Beneficiario
    If (Trim(Txt_BMFecNac) = "") Then
       MsgBox "Debe ingresar una Fecha de Nacimiento de Beneficiario", vbCritical, "Error de Datos"
       Txt_BMFecNac.SetFocus
       
       Exit Sub
    End If
    
    If Not IsDate(Txt_BMFecNac.Text) Then
       MsgBox "La Fecha Ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
       Txt_BMFecNac.SetFocus
      
       Exit Sub
    End If
    If (CDate(Txt_BMFecNac) > CDate(Date)) Then
       MsgBox "La Fecha Ingresada es Mayor a la Fecha Actual", vbCritical, "Error de Datos"
       Txt_BMFecNac.SetFocus
       
       Exit Sub
    End If
    If (Year(Txt_BMFecNac) < 1900) Then
       MsgBox "La Fecha Ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
       Txt_BMFecNac.SetFocus
      
       Exit Sub
    End If
        
         vgRes = MsgBox(" ¿ Está seguro que desea Ingresar este Beneficiario ?", 4 + 32 + 256, "Operación de Ingreso")
            If vgRes <> 6 Then
                Screen.MousePointer = 0
                Exit Sub
            End If
        
        Lbl_EndCrear = (Val(vlnum_endoso_actual) + 1)
        vlnum_endoso = Lbl_EndosoActual
        Call fl_insert_update_ItemGridBeneficiario
   
        Msf_BMGrilla.Enabled = True
        Lbl_BMNumOrd = ""

End Sub

Private Sub Cmd_BuscarDir_Click()
On Error GoTo Err_Buscar
     
    Frm_BusDireccion.flInicio ("Frm_EndosoHerencia")
    
Exit Sub
Err_Buscar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub


Private Sub Cmd_BuscarPendiente_Click()

If (Txt_PenPoliza.Text = "") And fgObtenerCodigo_TextoCompuesto(Cmb_PenNumIdent) = 0 And (Txt_PenNumIdent.Text = "") Then
    MsgBox "Debe ingresar una poliza o documento", vbCritical, "Error de Datos"
     Exit Sub
 End If

If (Txt_PenPoliza.Text = "") Then
    MsgBox "Debe ingresar el numero de poliza ", vbCritical, "Error de Datos"
     Txt_PenPoliza.SetFocus
  Exit Sub
End If

If fgObtenerCodigo_TextoCompuesto(Cmb_PenNumIdent) > 0 And (Txt_PenNumIdent.Text = "") Then
  MsgBox "Debe ingresar el numero del documeto ", vbCritical, "Error de Datos"
  Txt_PenNumIdent.SetFocus
  Exit Sub
End If

Txt_PenPoliza = UCase(Trim(Txt_PenPoliza))
Txt_PenPoliza = Format(Txt_PenPoliza, "0000000000")

If fl_existe_pp_tmae_endben_pendiente_aprobar(Txt_PenPoliza.Text) = False And fl_existe_pp_tmae_endpoliza_pendiente_aprobar(Txt_PenPoliza.Text) = False Then
    MsgBox "El Beneficiario o la Póliza Ingresados, No Tienen Registro de Endoso Preliminar", vbCritical, "Inexistencia de Endoso Preliminar"
    Exit Sub
End If
Call flDefineColumnsGridBeneficiario(Msf_BMGrilla)
Call flBuscarPoliza("P", "pp_tmae_endben")
 
End Sub

Sub flBuscarPoliza(p_tipo_busqueda As String, p_name_table As String)
vg_tipo_busqueda = p_tipo_busqueda

vgSql = "SELECT num_poliza,num_endoso,num_orden,num_idenben,cod_tipoidenben,gls_nomben,gls_nomsegben,gls_patben,gls_matben, cod_grufam,cod_par,(select distinct cod_moneda from pp_tmae_poliza where num_poliza=a.num_poliza) cod_moneda,"
vgSql = vgSql & " cod_sexo,cod_sitinv,cod_dercre,cod_derpen,fec_nacben,mto_pensiongar,cod_estpension,fec_inipagopen, fec_terpagopengar,prc_pensiongar,"
vgSql = vgSql & " gls_telben2,cod_direccion,gls_dirben,gls_fonoben,gls_correoben, mto_pension, cod_MotReqPen, prc_pension FROM " & p_name_table & " a where num_poliza = '" & Txt_PenPoliza.Text & "' "
vgSql = vgSql & " and num_endoso = (SELECT MAX (num_endoso) as número FROM " & p_name_table & " where num_poliza = " & Txt_PenPoliza.Text & " ) order by num_orden "

If p_tipo_busqueda = "P" Then

    vgSql = "SELECT e.num_poliza,e.num_endoso,e.num_orden,e.num_idenben,e.cod_tipoidenben,e.gls_nomben,e.gls_nomsegben,e.gls_patben,e.gls_matben, e.cod_grufam,e.cod_par,"
    vgSql = vgSql & "e.cod_sexo,e.cod_sitinv,e.cod_dercre,e.cod_derpen,e.fec_nacben,e.mto_pensiongar,e.cod_estpension,e.fec_inipagopen, e.fec_terpagopengar,e.prc_pensiongar,"
    vgSql = vgSql & "e.gls_telben2,e.cod_direccion,e.gls_dirben,e.gls_fonoben,e.gls_correoben, e.mto_pension, e.cod_MotReqPen, e.prc_pension FROM pp_tmae_endben e "
    vgSql = vgSql & "where e.num_poliza = '" & Txt_PenPoliza.Text & "' and not exists ( select * from pp_tmae_ben a where a.num_poliza = '" & Txt_PenPoliza.Text & "' and a.num_endoso = (e.num_endoso - 1) and a.num_orden = e.num_orden )"
    vgSql = vgSql & "order by e.num_orden"

End If


Set vgRs2 = vgConexionBD.Execute(vgSql)
    Do While Not vgRs2.EOF
 
        vlNum_Poliza = vgRs2!num_poliza
        'vlnum_endoso = IIf(IsNull(vgRs2!num_endoso), 0, vgRs2!num_endoso)
        vlnum_orden = IIf(IsNull(vgRs2!Num_Orden), 0, vgRs2!Num_Orden)
        vlcod_tipoidenben = IIf(IsNull(vgRs2!Cod_TipoIdenBen), 0, vgRs2!Cod_TipoIdenBen)
        
        vlnum_idenben = IIf(IsNull(vgRs2!Num_IdenBen), "", vgRs2!Num_IdenBen)
        vlgls_nomben = IIf(IsNull(vgRs2!Gls_NomBen), "", vgRs2!Gls_NomBen)
        vlgls_nomsegben = IIf(IsNull(vgRs2!Gls_NomSegBen), "", vgRs2!Gls_NomSegBen)
        vlgls_patben = IIf(IsNull(vgRs2!Gls_PatBen), "", vgRs2!Gls_PatBen)
        vlgls_matben = IIf(IsNull(vgRs2!Gls_MatBen), "", vgRs2!Gls_MatBen)
        
        If vlnum_orden = 1 Then
            vlxcod_tipoidenben = IIf(IsNull(vgRs2!Cod_TipoIdenBen), 0, vgRs2!Cod_TipoIdenBen)
            vlxnum_idenben = IIf(IsNull(vgRs2!Num_IdenBen), "", vgRs2!Num_IdenBen)
            vlxgls_nomben = IIf(IsNull(vgRs2!Gls_NomBen), "", vgRs2!Gls_NomBen)
            vlxgls_nomsegben = IIf(IsNull(vgRs2!Gls_NomSegBen), "", vgRs2!Gls_NomSegBen)
            vlxgls_patben = IIf(IsNull(vgRs2!Gls_PatBen), "", vgRs2!Gls_PatBen)
            vlxgls_matben = IIf(IsNull(vgRs2!Gls_MatBen), "", vgRs2!Gls_MatBen)
            vlxcod_direccion = IIf(IsNull(vgRs2!Cod_Direccion), "", vgRs2!Cod_Direccion)
            vlxgls_dirben = IIf(IsNull(vgRs2!Gls_DirBen), "", vgRs2!Gls_DirBen)
        End If
        
        vlcod_grufam = IIf(IsNull(vgRs2!Cod_GruFam), "", vgRs2!Cod_GruFam)
        vlcod_par = IIf(IsNull(vgRs2!Cod_Par), "", vgRs2!Cod_Par)
        vlcod_sexo = IIf(IsNull(vgRs2!Cod_Sexo), "", vgRs2!Cod_Sexo)
        vlcod_sitinv = IIf(IsNull(vgRs2!Cod_SitInv), "", vgRs2!Cod_SitInv)
        vlcod_dercre = IIf(IsNull(vgRs2!Cod_DerCre), "", vgRs2!Cod_DerCre)
        vlcod_derpen = IIf(IsNull(vgRs2!Cod_DerPen), "", vgRs2!Cod_DerPen)
        vlfec_nacben = DateSerial(Mid((vgRs2!Fec_NacBen), 1, 4), Mid((vgRs2!Fec_NacBen), 5, 2), Mid((vgRs2!Fec_NacBen), 7, 2))
        vlmto_pensiongar = IIf(IsNull(vgRs2!Mto_PensionGar), "0", vgRs2!Mto_PensionGar)
        vlcod_estpension = IIf(IsNull(vgRs2!Cod_EstPension), "", vgRs2!Cod_EstPension)
        vlfec_inipagopen = IIf(IsNull(vgRs2!Fec_IniPagoPen), "", vgRs2!Fec_IniPagoPen)
        vlfec_terpagopengar = IIf(IsNull(vgRs2!Fec_TerPagoPenGar), "", vgRs2!Fec_TerPagoPenGar)
        vlprc_pensiongar = IIf(IsNull(vgRs2!Prc_PensionGar), "0", vgRs2!Prc_PensionGar)
        vlgls_telben2 = IIf(IsNull(vgRs2!Gls_Telben2), "", vgRs2!Gls_Telben2)
        vlcod_direccion = IIf(IsNull(vgRs2!Cod_Direccion), "", vgRs2!Cod_Direccion)
        vlgls_dirben = IIf(IsNull(vgRs2!Gls_DirBen), "", vgRs2!Gls_DirBen)
        vlgls_fonoben = IIf(IsNull(vgRs2!Gls_FonoBen), "", vgRs2!Gls_FonoBen)
        vlgls_correoben = IIf(IsNull(vgRs2!Gls_CorreoBen), "", vgRs2!Gls_CorreoBen)
        
        vlcod_motreqpen = IIf(IsNull(vgRs2!Cod_MotReqPen), "0", vgRs2!Cod_MotReqPen)
        vlprc_pension = IIf(IsNull(vgRs2!Prc_Pension), "0", vgRs2!Prc_Pension)
        vlmto_pension = 0
        
        Call fgBuscarPosicionCodigoCombo(vlcod_tipoidenben, Cmb_PenNumIdent)
        cmbExcluesion.ListIndex = fgBuscarPosicionCodigoCombo(vlcod_estpension, cmbExcluesion)
        If vlcod_estpension = "10" Then
            chkContinuarpago.Value = 1
        End If
        Txt_PenNumIdent.Text = vlnum_idenben
        Lbl_End = vlnum_endoso
        lblMoneda = vgRs2!Cod_Moneda
        vlNombreCompleto = fgFormarNombreCompleto(vlgls_nomben, vlgls_nomsegben, vlgls_patben, vlgls_matben)
        Lbl_PenNombre.Caption = vlNombreCompleto
        
        If p_tipo_busqueda = "P" Then
        
        vlmto_pension = IIf(IsNull(vgRs2!Mto_Pension), "0", vgRs2!Mto_Pension)
        
        Msf_BMGrilla.AddItem vlnum_orden & vbTab & vlcod_tipoidenben & vbTab & vlnum_idenben & vbTab & vlgls_nomben & vbTab & _
        vlgls_nomsegben & vbTab & vlgls_patben & vbTab & vlgls_matben & vbTab & vlcod_grufam & vbTab & vlcod_par & vbTab & vlcod_sexo & vbTab & _
        vlcod_sitinv & vbTab & vlcod_dercre & vbTab & vlcod_derpen & vbTab & vlfec_nacben & vbTab & vlmto_pensiongar & vbTab & vlcod_estpension & _
        vbTab & vlfec_inipagopen & vbTab & vlfec_terpagopengar & vbTab & vlprc_pensiongar & vbTab & vlgls_telben2 & vbTab & vlcod_direccion & vbTab & _
        vlgls_dirben & vbTab & vlgls_fonoben & vbTab & vlgls_correoben & vbTab & vlmto_pension

        End If

        vgRs2.MoveNext
        Loop
        
        If p_tipo_busqueda = "P" Then
            Call fl_cargar_valor_controles_beneficiario
        End If

        
    '' Poliza

vgSql = ""
vgSql = vgSql & " SELECT num_poliza,num_endoso,cod_tippension,cod_estado, cod_tipren,cod_modalidad,num_cargas,fec_vigencia,"
vgSql = vgSql & " fec_tervigencia,mto_prima,mto_pension,num_mesdif, num_mesgar,prc_tasace,prc_tasavta,prc_tasaintpergar,"
vgSql = vgSql & " fec_dev,fec_inipencia,cod_moneda,mto_valmoneda,ind_cob,cod_cobercon,mto_facpenella,prc_facpenella, cod_dercre,"
vgSql = vgSql & " cod_dergra , cod_tipreajuste, mto_valreajustetri, mto_valreajustemen, '0' as mto_pensiongar, fec_finpergar"
vgSql = vgSql & " FROM PP_TMAE_POLIZA a WHERE a.num_poliza = " & Txt_PenPoliza.Text

'If p_tipo_busqueda = "P" Then
vgSql = vgSql & " AND num_endoso = (select max(num_endoso) from PP_TMAE_POLIZA where num_poliza=a.num_poliza)"
'Else
'    vgSql = vgSql & " AND num_endoso = 1"
'End If

 Set vgRs2 = vgConexionBD.Execute(vgSql)
 Do While Not vgRs2.EOF
 
   Lbl_EndosoActual = IIf(IsNull(vgRs2!num_endoso), 0, vgRs2!num_endoso)
   vlnum_endoso_actual = IIf(IsNull(vgRs2!num_endoso), 0, vgRs2!num_endoso)
   
   Cmb_PMTipPen.ListIndex = fgBuscarPosicionCodigoCombo(Trim(vgRs2!Cod_TipPension), Cmb_PMTipPen)
   Cmb_PMEstVig.ListIndex = fgBuscarPosicionCodigoCombo(Trim(vgRs2!Cod_Estado), Cmb_PMEstVig)
   Cmb_PMTipRta.ListIndex = fgBuscarPosicionCodigoCombo(Trim(vgRs2!Cod_TipRen), Cmb_PMTipRta)
   Cmb_PMMod.ListIndex = fgBuscarPosicionCodigoCombo(Trim(vgRs2!Cod_Modalidad), Cmb_PMMod)
   Txt_PMIniVig = DateSerial(Mid((vgRs2!Fec_Vigencia), 1, 4), Mid((vgRs2!Fec_Vigencia), 5, 2), Mid((vgRs2!Fec_Vigencia), 7, 2))
   Txt_PMTerVig = DateSerial(Mid((vgRs2!Fec_TerVigencia), 1, 4), Mid((vgRs2!Fec_TerVigencia), 5, 2), Mid((vgRs2!Fec_TerVigencia), 7, 2))
   Txt_PMMesDif = vgRs2!Num_MesDif
   
    Txt_PMMesGar = vgRs2!Num_MesGar
    Txt_PMMtoPrima = Format(vgRs2!Mto_Prima, "###,###,##0.00")
    vlmto_pensioncal = IIf(IsNull(vgRs2!Mto_Prima), "0", vgRs2!Mto_Prima)
    Txt_PMTasaCto = Format(vgRs2!Prc_TasaCe, "##0.000")
    Txt_PMTasaVta = Format(vgRs2!Prc_TasaVta, "##0.000")
    Txt_PMTasaPerGar = Format(vgRs2!Prc_TasaIntPerGar, "##0.00")
    
    Lbl_PMFecDevengue = DateSerial(Mid((vgRs2!fec_dev), 1, 4), Mid((vgRs2!fec_dev), 5, 2), Mid((vgRs2!fec_dev), 7, 2))
    Lbl_PMIndCobertura = vgRs2!Ind_Cob
    Lbl_PMCoberCon = vgRs2!Cod_CoberCon
    Lbl_PMDerCrecer = vgRs2!Cod_DerCre
    Lbl_PMDerGratificacion = vgRs2!Cod_DerGra
    Lbl_PMMtoFacPenElla = vgRs2!Mto_FacPenElla
    Lbl_PMPrcFacPenElla = vgRs2!Prc_FacPenElla

    lblTC = Format(vgRs2!Mto_ValMoneda, "##0.00")
    vlcod_moneda = vgRs2!Cod_Moneda
        
  vgRs2.MoveNext
 Loop

If p_tipo_busqueda = "P" Then
    vgSql = "select num_endoso, num_mesgarres, gls_docjud, mto_pensioncal  from PP_TMAE_ENDENDOSO e where e.num_poliza = '" & Txt_PenPoliza.Text & "' and  num_endoso = " & Lbl_EndosoActual
    Set vgRs2 = vgConexionBD.Execute(vgSql)

    Do While Not vgRs2.EOF
        txtnum_mesgarres.Text = IIf(IsNull(vgRs2!num_mesgarres), "", vgRs2!num_mesgarres)
        txtgls_docjud.Text = IIf(IsNull(vgRs2!gls_docjud), "", vgRs2!gls_docjud)
        txtmto_pensioncal.Text = IIf(IsNull(vgRs2!mto_pensioncal), 0, vgRs2!mto_pensioncal)
    vgRs2.MoveNext
    Loop

End If

 Lbl_EndCrear = ((Lbl_EndosoActual) + 1)
 
 Cmd_BMRestar.Enabled = False
 Cmd_BuscarPol.Enabled = False
 Cmd_BuscarPendiente.Enabled = False
 Cmd_Modificar.Enabled = True
 Cmd_Eliminar.Enabled = True
 Cmd_Imprimir.Enabled = True
 Cmd_Limpiar.Enabled = True
 Cmd_Cancelar.Enabled = True
 
 
' If p_tipo_busqueda = "P" Then
'    Cmd_Aprobar.Enabled = True
' Else
'    Cmd_Aprobar.Enabled = False
'
' End If
 
 Txt_PenPoliza.Enabled = False
 Cmb_PenNumIdent.Enabled = False
 Txt_PenNumIdent.Enabled = False
 
 


Exit Sub
Err_CmdBuscarPolClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Sub fl_begin_trans_rollbak()
        vgConectarBD.RollbackTrans
        vgConectarBD.Close
        Cmd_Salir.SetFocus
        Screen.MousePointer = 0
        Exit Sub
End Sub

Function fl_grabar(iNumPoliza As String, inumendoso As String) As Boolean
On Error GoTo Err_Aprobar
    'Abrir la Conexión
    
    fl_grabar = False
    
    If Not fgConexionBaseDatos(vgConectarBD) Then
        MsgBox "Falló la conexión con la Base de Datos.", vbCritical, "Error de Conexión"
        Exit Function
    End If
    'Comenzar la Transacción
    vgConectarBD.BeginTrans
    
    ' Insert endozo temporal
       fl_grabar = fl_grabar_insertar_IN_pp_tmae_endendoso
       If fl_grabar = False Then Call fl_begin_trans_rollbak
       
    'BENEFICIARIOS
    '---------------------
    'Eliminar Beneficiario
    vgSql = "DELETE FROM PP_TMAE_ENDBEN "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "num_poliza = '" & iNumPoliza & "' AND "
    vgSql = vgSql & "num_endoso = " & Lbl_EndCrear
    vgConectarBD.Execute vgSql
    
    Dim lRecordsAffected As Long
    
    
    If fl_insertar_temp_IN_select_pp_tmae_endben = False Then Call fl_begin_trans_rollbak
    
    'Ingresar los Beneficiarios
    If Msf_BMGrilla.rows > 0 Then
        vlPos = 1
        Msf_BMGrilla.Col = 0
        While vlPos <= (Msf_BMGrilla.rows - 1)
            Msf_BMGrilla.row = vlPos
            Call fl_cargar_valor_variables_grid_beneficiarios(Msf_BMGrilla.row)
            Msf_BMGrilla.Col = 0
            fl_grabar = fl_Insertar_temp_beneficiario(vlPos)
            If fl_grabar = False Then Call fl_begin_trans_rollbak

            vlPos = vlPos + 1
        Wend
        
    Else
        fl_grabar = False
        vgConectarBD.RollbackTrans
        vgConectarBD.Close
        MsgBox "No existen Beneficiarios a cargar desde la Grilla de Beneficiarios Modificados.", vbCritical, "Beneficiarios Inexistentes"
        Screen.MousePointer = 0
        Exit Function
    
    End If
        
    vgConectarBD.CommitTrans
    vgConectarBD.Close
    fl_grabar = True
    
Exit Function
Err_Aprobar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        fl_grabar = False
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Private Sub Cmd_BuscarPol_Click()

If (Txt_PenPoliza.Text = "") And fgObtenerCodigo_TextoCompuesto(Cmb_PenNumIdent) = 0 And (Txt_PenNumIdent.Text = "") Then
    MsgBox "Debe ingresar una poliza o documento", vbCritical, "Error de Datos"
     Exit Sub
 End If

If (Txt_PenPoliza.Text = "") Then
    MsgBox "Debe ingresar el numero de poliza ", vbCritical, "Error de Datos"
     Txt_PenPoliza.SetFocus
  Exit Sub
End If

If fgObtenerCodigo_TextoCompuesto(Cmb_PenNumIdent) > 0 And (Txt_PenNumIdent.Text = "") Then
  MsgBox "Debe ingresar el numero del documeto ", vbCritical, "Error de Datos"
  Txt_PenNumIdent.SetFocus
  Exit Sub
End If

Txt_PenPoliza = UCase(Trim(Txt_PenPoliza))
Txt_PenPoliza = Format(Txt_PenPoliza, "0000000000")

If fl_existe_poliza_in_pp_tmae_ben(Txt_PenPoliza.Text) = False Then
 MsgBox "La Póliza Ingresada no existe, Ingrese otro e intente nuevamente.", vbCritical, "Existencia de Póliza Preliminar"
        Exit Sub
End If

If fl_existe_pp_tmae_endpoliza_pendiente_aprobar(Txt_PenPoliza.Text) = True Then
         MsgBox "El Beneficiario o la Póliza Ingresada, tienen un registro de Endoso Pendiente. No es posible crear un nuevo Endoso hasta finalizar el que se encuentra en Etapa Preliminar.", vbCritical, "Existencia de Póliza Preliminar"
    Exit Sub
End If

If fl_existe_pp_tmae_endben_pendiente_aprobar(Txt_PenPoliza.Text) = True Then
         MsgBox "El Beneficiario o la Póliza Ingresada, tienen un registro de Endoso Pendiente. No es posible crear un nuevo Endoso hasta finalizar el que se encuentra en Etapa Preliminar.", vbCritical, "Existencia de Póliza Preliminar"
    Exit Sub
End If


'Call Txt_PenPoliza_KeyPress(13)

Call flDefineColumnsGridBeneficiario(Msf_BMGrilla)
Call flBuscarPoliza("E", "pp_tmae_ben")
End Sub

Private Sub Cmd_Cancelar_Click()
Habilitar (False)
Msf_BMGrilla.rows = 1
Txt_PenPoliza.Text = ""

Txt_PenNumIdent.Text = ""

Cmd_BuscarPol.Enabled = True
Cmd_BuscarPendiente.Enabled = True
Cmb_PenNumIdent.ListIndex = fgBuscarPosicionCodigoCombo(0, Cmb_BMNumIdent)
Lbl_End = ""
Lbl_PenNombre = ""
Lbl_BMNumOrd = ""
Call Limpiar
Txt_PenPoliza.Enabled = True
Cmb_PenNumIdent.Enabled = True
Txt_PenNumIdent.Enabled = True
Cmd_Modificar.Enabled = False
Cmd_Grabar.Enabled = False
Cmd_Eliminar.Enabled = False
'Cmd_Aprobar.Enabled = False
Cmd_Imprimir.Enabled = False
Cmd_Limpiar.Enabled = False
Cmd_Cancelar.Enabled = False
Cmd_BMRestar.Enabled = False
 chkContinuarpago.Enabled = False
Cmd_BMCalcular.Enabled = False
Cmd_BMSumar.Enabled = False

Msf_BMGrilla.Enabled = True
Lbl_EndosoActual = ""

chkContinuarpago.Value = 0


End Sub

Private Sub cmd_grabar_Click()

' If Msf_BMGrilla.rows <= 1 Then
'     MsgBox "Debe Ingresar un Beneficiario.", vbCritical, "Error de Datos"
'     Exit Sub
' End If


'Validar Código de Identificación del Beneficiario
'    If (Cmb_BMNumIdent.Text) = "" Then
'        MsgBox "Debe seleccionar el Tipo de Identificación del Beneficiario.", vbCritical, "Error de Datos"
'        Cmb_BMNumIdent.SetFocus
'        Exit Sub
'    End If
'    'Validar Número de Identificación del Beneficiario
'    Txt_BMNumIden = Trim(UCase(Txt_BMNumIden))
'    If Txt_BMNumIden.Text = "" Then
'       MsgBox "Debe Ingresar el Número de Identificación del Beneficiario.", vbCritical, "Error de Datos"
'       Txt_BMNumIden.SetFocus
'       Exit Sub
'    End If
    
    
    'Valida Nombre del Beneficiario
'    txtNomben.Text = Trim(UCase(txtNomben.Text))
'    If txtNomben.Text = "" Then
'       MsgBox "Debe Ingresar Nombre.", vbCritical, "Error de Datos"
'       'frmDatosGen.Visible = True
'       txtNomben.SetFocus
'       Exit Sub
'    End If
'
'    'Valida Apellido Paterno del Beneficiario
'    txtApepatBen.Text = Trim(UCase(txtApepatBen.Text))
'    If txtApepatBen.Text = "" Then
'       MsgBox "Debe Ingresar Apellido Paterno.", vbCritical, "Error de Datos"
'       txtApepatBen.SetFocus
'       vlSwCalIntOK = False
'       Exit Sub
'    End If
'
'
'    'Valida Fecha de Nacimiento del Beneficiario
'    If (Trim(Txt_BMFecNac) = "") Then
'       MsgBox "Debe ingresar una Fecha de Nacimiento de Beneficiario", vbCritical, "Error de Datos"
'       Txt_BMFecNac.SetFocus
'
'       Exit Sub
'    End If
'    If Not IsDate(Txt_BMFecNac.Text) Then
'       MsgBox "La Fecha Ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
'       Txt_BMFecNac.SetFocus
'       Exit Sub
'    End If
'    If (CDate(Txt_BMFecNac) > CDate(Date)) Then
'       MsgBox "La Fecha Ingresada es Mayor a la Fecha Actual", vbCritical, "Error de Datos"
'       Txt_BMFecNac.SetFocus
'       Exit Sub
'    End If
'    If (Year(Txt_BMFecNac) < 1900) Then
'       MsgBox "La Fecha Ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
'       Txt_BMFecNac.SetFocus
'       Exit Sub
'    End If
'
'    Txt_BMMtoPensionGar = Trim(Txt_BMMtoPensionGar)
'    If Txt_BMMtoPensionGar.Text = "" Then
'        Txt_BMMtoPensionGar.Text = "0"
'    End If
'
'    txt_BMPrcGar = Trim(txt_BMPrcGar)
'    If txt_BMPrcGar.Text = "" Then
'        txt_BMPrcGar.Text = "0"
'    End If
'
'   Dim lv_message As String
'
If vg_tipo_busqueda = "P" Then

    If Not IsNumeric(txtmto_pensioncal.Text) Then
        MsgBox "Debe ingresar el Monto liquidar. Edite y realice los cambios", vbCritical, "Error de Datos"
    Exit Sub
    End If
    
    If Val(txtmto_pensioncal.Text) < 1 Then
        MsgBox "El Monto liquidar debe ser mayor a cero.", vbCritical, "Error de Datos"
        txtmto_pensioncal.SetFocus
        Exit Sub
    End If
    
    lv_message = " ¿ Está seguro que desea grabar el Endoso ?"
    
    Else
    
    lv_message = " ¿ Está seguro que desea grabar el Endoso como Preliminar ?"
   
End If
   
   Dim b_isCorrect As Boolean
   b_isCorrect = True
   
    txtmto_pensioncal = Trim(txtmto_pensioncal)
    If txtmto_pensioncal.Text = "" Then
        txtmto_pensioncal.Text = "0"
    End If
    
    txtnum_mesgarres = Trim(txtnum_mesgarres)
    If txtnum_mesgarres.Text = "" Then
        txtnum_mesgarres.Text = "0"
    End If
    
    txtgls_docjud = Trim(txtgls_docjud)
    If txtgls_docjud.Text = "" Then
        txtgls_docjud.Text = "0"
    End If


' If Msf_BMGrilla.rows > 1 Then
'     vlmto_pension = Val(txtmto_pensioncal.Text) / (Msf_BMGrilla.rows - 1)
' Else
'     vlmto_pension = Val(txtmto_pensioncal.Text)
' End If
'
'    Dim resp As String
'    If Msf_BMGrilla.rows > 1 Then
'     vlPos = 1
'     Msf_BMGrilla.Col = 24
'     Do While vlPos <= (Msf_BMGrilla.rows - 1)
'         Msf_BMGrilla.row = vlPos
'         Msf_BMGrilla.Col = 0
'         resp = Msf_BMGrilla.TextMatrix(vlPos, 24)
'         If resp = "0" Or resp = "" Or vlmto_pension <> resp Then
'            b_isCorrect = False
'            Exit Do
'         End If
'         vlPos = vlPos + 1
'     Loop
'     End If
          

'If vg_tipo_busqueda = "P" Then
'
'    If b_isCorrect = False Then
'      lv_message = " Debe realizar el recálculo, el monto de pension del beneficiario no esta en la proporcion del monto liquidar"
'      MsgBox lv_message, vbCritical, "Error de Datos"
'      Exit Sub
'      Else
'
'      vgRes = MsgBox(lv_message, 4 + 32 + 256, "Operación de Grabación Endoso")
'      If vgRes <> 6 Then
'         Screen.MousePointer = 0
'         Exit Sub
'      End If
'
'    End If
'
'Else
'    If b_isCorrect = False Then
'      lv_message = " ¿ Debe realizar el recálculo para la pension de los beneficiarios, Está seguro que desea grabar el Endoso como Preliminar ?"
'      vgRes = MsgBox(lv_message, 4 + 32 + 256, "Operación de Grabación Endoso")
'      If vgRes <> 6 Then
'        Screen.MousePointer = 0
'        Exit Sub
'       End If
'    End If
'End If

                   
Screen.MousePointer = 11

If (fl_grabar(Me.Txt_PenPoliza.Text, Lbl_EndosoActual) = True) Then
    MsgBox "El proceso de Grabación ha finalizado Correctamente.", vbInformation, "Estado del Proceso"
    
    Cmd_BMSumar.Enabled = False
    Cmd_BMRestar.Enabled = False
    Cmd_BMCalcular.Enabled = False
    
    Screen.MousePointer = 0
    If vg_tipo_busqueda = "P" Then
        'Cmd_Aprobar.Enabled = True
        Cmd_Grabar.Enabled = False
     Else
        'Cmd_Aprobar.Enabled = False
        Cmd_Modificar.Enabled = False
        Cmd_Grabar.Enabled = False
        'Cmd_Aprobar.Enabled = False
        Cmd_Limpiar.Enabled = False
        'Call Cmd_Cancelar_Click
     End If
     
     Cmd_BuscarDir.Enabled = False
     'txtnum_mesgarres.Enabled = False
     'txtmto_pensioncal.Enabled = False
     txtgls_docjud.Enabled = False
    
    
Else
    MsgBox "El proceso de Grabación ha sido cancelado por Problemas en el Proceso.", vbCritical, "Estado del Proceso"
End If


End Sub
Sub Limpiar()
        txtNomben.Text = ""
        txtNomsegBen.Text = ""
        txtApepatBen.Text = ""
        txtApematBen.Text = ""
        
        Txt_BMNumIden.Text = ""
        txtTelBen.Text = ""
        txtTelfon2.Text = ""
        txtBenEmail.Text = ""
        txt_dirben.Text = ""
        txtDistritoBen.Text = ""
        Txt_BMFecNac.Text = ""
        Txt_BMMtoPensionGar.Text = ""
        txtnum_mesgarres.Text = ""
        txtmto_pensioncal.Text = ""
        txtgls_docjud.Text = ""
    
End Sub
Sub Habilitar(Value As Boolean)
        
        txtNomben.Enabled = Value
        txtNomsegBen.Enabled = Value
        txtApepatBen.Enabled = Value
        txtApematBen.Enabled = Value
        Txt_BMNumIden.Enabled = Value
        txtTelBen.Enabled = Value
        txtTelfon2.Enabled = Value
        txtBenEmail.Enabled = Value
        txt_dirben.Enabled = Value
        txtDistritoBen.Enabled = Value
        Txt_BMFecNac.Enabled = Value
        Txt_BMMtoPensionGar.Enabled = Value
        
'        txtnum_mesgarres.Enabled = Value
'        txtmto_pensioncal.Enabled = Value
'        txtgls_docjud.Enabled = Value
        
        txt_BMPrcGar.Enabled = Value
End Sub

Private Sub Cmd_Imprimir_Click()

 'Verificar el Número de Póliza
    If (Trim(Txt_PenPoliza) = "") Then
        MsgBox "Debe indicar el Número de Póliza sobre la cual se quiere realizar la Operación.", vbCritical, "Error de Datos"
        Exit Sub
    End If
    'Verificar el Número de Endoso
    If Not IsNumeric(Trim(Lbl_EndosoActual)) Then
        MsgBox "Debe indicar el Número de Endoso sobre el cual se quiere realizar la Operación.", vbCritical, "Error de Datos"
        Exit Sub
    End If
      
    strRpt = "C:\\Sistemas Rtas Vit\\Reportes\\"
    vlGlobalNumPoliza = Trim(Txt_PenPoliza)
    vlGlobalNumEndoso = Trim(Lbl_EndosoActual)
    vlRptNumEndosoPol = Lbl_EndosoActual
    vlRptNumEndosoEnd = Lbl_EndosoActual
    
 If vg_tipo_busqueda = "P" Then
         
        Call fl_imprimir_poliza
        Call flImprimirEndosoPrev
        
    Else
         
        Call fl_imprimir_endoso
 End If

End Sub

Private Sub Cmd_Limpiar_Click()

        Cmb_BMPar.ListIndex = fgBuscarPosicionCodigoCombo("50", Cmb_BMPar)
        Cmb_BMDerPen.ListIndex = fgBuscarPosicionCodigoCombo("99", Cmb_BMDerPen)
        cmbExcluesion.ListIndex = fgBuscarPosicionCodigoCombo("99", cmbExcluesion)
        Habilitar (True)
        Call Limpiar
        
       Msf_BMGrilla.Enabled = False
       Lbl_BMNumOrd = ""
       Cmd_BMSumar.Enabled = True
       Cmd_BMRestar.Enabled = True
       Cmd_Grabar.Enabled = True
       
       txtnum_mesgarres.Enabled = True
        txtmto_pensioncal.Enabled = True
        txtgls_docjud.Enabled = True
       
       If vg_tipo_busqueda = "P" Then
          Cmd_Aprobar.Enabled = False
       End If
       chkContinuarpago.Enabled = True
       Cmd_BMCalcular.Enabled = True
       chkContinuarpago.Value = 0
       
       Cmd_BuscarDir.Enabled = True
 
        
End Sub

Private Sub Cmd_Modificar_Click()
Habilitar (True)

Cmd_BMCalcular.Enabled = True

Cmd_Grabar.Enabled = True
Cmd_Eliminar.Enabled = False
Cmd_BMSumar.Enabled = True
Cmd_Aprobar.Enabled = False

txtnum_mesgarres.Enabled = True
txtmto_pensioncal.Enabled = True
txtgls_docjud.Enabled = True
 
If vg_tipo_busqueda = "P" Then
   Cmd_BMRestar.Enabled = True
   Cmd_Aprobar.Enabled = False
End If
 chkContinuarpago.Enabled = True
'   Cmd_BMCalcular.Enabled = True

Cmd_BuscarDir.Enabled = True
 
End Sub

Private Sub cmd_salir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    'poliza
    flComboTipoIdentificacion Cmb_PenNumIdent
    fgComboGeneral vgCodTabla_TipPen, Cmb_PMTipPen
    fgComboGeneral vgCodTabla_TipVigPol, Cmb_PMEstVig
    fgComboGeneral vgCodTabla_TipRen, Cmb_PMTipRta
    fgComboGeneral vgCodTabla_AltPen, Cmb_PMMod
    
    ' beneficiario
    flComboTipoIdentificacion Cmb_BMNumIdent
    fgComboGeneral vgCodTabla_Par, Cmb_BMPar
'    fgComboGeneral vgCodTabla_GruFam, Cmb_BMGrupFam
    fgComboGeneral vgCodTabla_Sexo, Cmb_BMSexo
'    fgComboGeneral vgCodTabla_SitInv, Cmb_BMSitInv
    
'    fgComboGeneral vgCodTabla_TipEnd, Cmb_EndTipoEnd
    fgComboGeneral vgCodTabla_DerPen, Cmb_BMDerPen
    fgComboGeneral vgCodTabla_TipPen, cmbTipoPreBen
'    fgComboGeneral vgCodTabla_AFP, cmbAfpBen

   
End Sub

'Este debe ir despues de flCargaVariables_GridBeneficiarios
Sub fl_cargar_valor_controles_beneficiario()
On Error GoTo Err_SetControles_Beneficiario

        txtcodDirBen.Text = vlcod_direccion
        Lbl_BMNumOrd.Caption = vlnum_orden
        Txt_BMNumIden.Text = vlnum_idenben
        Cmb_BMNumIdent.ListIndex = fgBuscarPosicionCodigoCombo(Trim(vlcod_tipoidenben), Cmb_BMNumIdent)
        txtNomben.Text = vlgls_nomben
        txtNomsegBen.Text = vlgls_nomsegben
        txtApepatBen.Text = vlgls_patben
        txtApematBen.Text = vlgls_matben
        'vlcod_grufam
        Cmb_BMPar.ListIndex = fgBuscarPosicionCodigoCombo(Trim(vlcod_par), Cmb_BMPar)
        Cmb_BMSexo.ListIndex = fgBuscarPosicionCodigoCombo(Trim(vlcod_sexo), Cmb_BMSexo)
        Cmb_BMDerPen.ListIndex = fgBuscarPosicionCodigoCombo(Trim(vlcod_derpen), Cmb_BMDerPen)
        cmbExcluesion.ListIndex = fgBuscarPosicionCodigoCombo(Trim(vlcod_estpension), cmbExcluesion)
        
        txtTelBen.Text = vlgls_fonoben
        txtTelfon2.Text = vlgls_telben2
        txtBenEmail.Text = vlgls_correoben
        txt_dirben.Text = vlgls_dirben
        
        Txt_BMMtoPensionGar.Text = Val(vlmto_pensiongar)
        txt_BMPrcGar.Text = Val(vlprc_pensiongar)
       
        Txt_BMFecNac.Text = vlfec_nacben
        
        
        vgSql = "select gls_comuna || '/' || gls_provincia || '/' || gls_region as distrito from ma_tpar_comuna a"
        vgSql = vgSql & " join ma_tpar_provincia b on a.cod_provincia=b.cod_provincia"
        vgSql = vgSql & " join ma_tpar_region c on a.cod_region=c.cod_region"
        vgSql = vgSql & " Where Cod_Direccion = '" & vlcod_direccion & "'"
        Set vlRegistro = vgConexionBD.Execute(vgSql)
        If Not vlRegistro.EOF Then
            txtDistritoBen.Text = vlRegistro!distrito
        End If
Exit Sub
Err_SetControles_Beneficiario:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub
Function fl_cargar_valor_variables_grid_beneficiarios(iPosicion As Integer)
On Error GoTo Err_flCargaVariables_GridBeneficiarios

        vlNum_Poliza = Txt_PenPoliza.Text
        vlnum_orden = Msf_BMGrilla.TextMatrix(iPosicion, 0)
        vlcod_tipoidenben = Msf_BMGrilla.TextMatrix(iPosicion, 1)
        vlnum_idenben = Msf_BMGrilla.TextMatrix(iPosicion, 2)
        vlgls_nomben = Msf_BMGrilla.TextMatrix(iPosicion, 3)
        vlgls_nomsegben = Msf_BMGrilla.TextMatrix(iPosicion, 4)
        vlgls_patben = Msf_BMGrilla.TextMatrix(iPosicion, 5)
        vlgls_matben = Msf_BMGrilla.TextMatrix(iPosicion, 6)
        vlcod_grufam = Msf_BMGrilla.TextMatrix(iPosicion, 7)
        vlcod_par = Msf_BMGrilla.TextMatrix(iPosicion, 8)
        vlcod_sexo = Msf_BMGrilla.TextMatrix(iPosicion, 9)
        vlcod_sitinv = Msf_BMGrilla.TextMatrix(iPosicion, 10)
        vlcod_dercre = Msf_BMGrilla.TextMatrix(iPosicion, 11)
        vlcod_derpen = Msf_BMGrilla.TextMatrix(iPosicion, 12)
        vlfec_nacben = Msf_BMGrilla.TextMatrix(iPosicion, 13)
        vlmto_pensiongar = Msf_BMGrilla.TextMatrix(iPosicion, 14)
        vlcod_estpension = Msf_BMGrilla.TextMatrix(iPosicion, 15)
        vlfec_inipagopen = Msf_BMGrilla.TextMatrix(iPosicion, 16)
        vlfec_terpagopengar = Msf_BMGrilla.TextMatrix(iPosicion, 17)
        vlprc_pensiongar = Msf_BMGrilla.TextMatrix(iPosicion, 18)
        vlgls_telben2 = Msf_BMGrilla.TextMatrix(iPosicion, 19)
        vlcod_direccion = Msf_BMGrilla.TextMatrix(iPosicion, 20)
        vlgls_dirben = Msf_BMGrilla.TextMatrix(iPosicion, 21)
        vlgls_fonoben = Msf_BMGrilla.TextMatrix(iPosicion, 22)
        vlgls_correoben = Msf_BMGrilla.TextMatrix(iPosicion, 23)
   
Exit Function
Err_flCargaVariables_GridBeneficiarios:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function



Private Sub Msf_BMGrilla_Click()
On Error GoTo Err_Msf_BOGrilla_Click
    
    Msf_BMGrilla.Col = 0
    If (Msf_BMGrilla.Text = "") Or (Msf_BMGrilla.row = 0) Then
        MsgBox "No existen Detalles", vbExclamation, "Información"
        Exit Sub
    Else
        
        Call fl_cargar_valor_variables_grid_beneficiarios(Msf_BMGrilla.row)
        Call fl_cargar_valor_controles_beneficiario
        
    End If

Exit Sub
Err_Msf_BOGrilla_Click:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Function flObtieneFechaEfecto() As String

On Error GoTo Err_ValidaFechaEfecto
        
        'Determinar si el periodo a registrar es posterior al que se desea ingresar
        vgSql = "select num_perpago || '01' fechaEfec from pp_tmae_propagopen where cod_estadoreg in ('A', 'P') "
        Set vgRs2 = vgConexionBD.Execute(vgSql)
        If Not vgRs2.EOF Then
            flObtieneFechaEfecto = vgRs2!fechaEfec
        End If
        vgRs2.Close

Exit Function
Err_ValidaFechaEfecto:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function
Function flObtieneFec_terpagopengar(pnum_poliza As String) As String

On Error GoTo Err_ValidaFechaEfecto
        'Determinar si el periodo a registrar es posterior al que se desea ingresar
        vgSql = "select to_char(TO_DATE(FEC_DEV, 'yyyy/mm/dd') + NUM_MESGAR,'yyyymmdd') as fec_terpagopengar "
        vgSql = vgSql & " from PP_TMAE_POLIZA p where p.num_poliza = " & pnum_poliza & ""
        vgSql = vgSql & " and num_endoso = (select max(num_endoso) from PP_TMAE_BEN where num_poliza= p.num_poliza)"
        
        Set vgRs2 = vgConexionBD.Execute(vgSql)
        If Not vgRs2.EOF Then
            flObtieneFec_terpagopengar = vgRs2!Fec_TerPagoPenGar
        End If
        vgRs2.Close

Exit Function
Err_ValidaFechaEfecto:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function
'FUNCTION QUE RETORNA LA DIRECCION DEL CONTROL BUSCAR DIRECCION
Function flRecibeDireccion(iNomDepartamento As String, iNomProvincia As String, iNomDistrito As String, iCodDir As String)
    Dim vStringDep As String
    Dim vStringProv As String
    Dim vStringDis As String
    Dim vlCodDir As String

    vStringDep = Trim(iNomDepartamento)
    vStringProv = Trim(iNomProvincia)
    vStringDis = Trim(iNomDistrito)
    vlCodDir = iCodDir

    txtcodDirBen = vlCodDir
    txtDistritoBen.Text = vStringDis & "/" & vStringProv & "/" & vStringDep
    
    txt_dirben.SetFocus
    
    ' Msf_BMGrilla.Enabled = True

End Function

'Function flExisteEndosoPendienteAprobacion(pnum_poliza As String) As Boolean
'On Error GoTo Err_flExisteAprobacionPendiente
'        flExisteAprobacionPendiente = False
'        vgSql = "SELECT num_poliza FROM PP_TMAE_ENDBEN where num_poliza = '" & pnum_poliza & "'"
'        Set vgRs2 = vgConexionBD.Execute(vgSql)
'        If Not vgRs2.EOF Then
'            flExisteAprobacionPendiente = True
'        End If
'        vgRs2.Close
'Exit Function
'Err_flExisteAprobacionPendiente:
'    Screen.MousePointer = 0
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
'End Function

Function fl_existe_pp_tmae_endben_pendiente_aprobar(pnum_poliza As String) As Boolean
On Error GoTo Err_flExisteEndosoPreliminar
         fl_existe_pp_tmae_endben_pendiente_aprobar = False
  
        vgSql = "SELECT NUM_POLIZA FROM PP_TMAE_ENDBEN b WHERE b.num_poliza = '" & pnum_poliza & "'"
        
        Set vgRs2 = vgConexionBD.Execute(vgSql)
        If Not vgRs2.EOF Then
            fl_existe_pp_tmae_endben_pendiente_aprobar = True
        End If
        vgRs2.Close
Exit Function
Err_flExisteEndosoPreliminar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function
Function fl_existe_pp_tmae_endpoliza_pendiente_aprobar(pnum_poliza As String) As Boolean
On Error GoTo Err_flExisteEndosoPreliminar
         fl_existe_pp_tmae_endpoliza_pendiente_aprobar = False
  
        vgSql = "SELECT NUM_POLIZA FROM PP_TMAE_ENDPOLIZA b WHERE b.num_poliza = '" & pnum_poliza & "'"
        
        Set vgRs2 = vgConexionBD.Execute(vgSql)
        If Not vgRs2.EOF Then
            fl_existe_pp_tmae_endpoliza_pendiente_aprobar = True
        End If
        vgRs2.Close
Exit Function
Err_flExisteEndosoPreliminar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function


Function fl_existe_poliza_in_pp_tmae_ben(pnum_poliza As String) As Boolean
On Error GoTo Err_fl_existe_poliza_in_pp_tmae_ben
        fl_existe_poliza_in_pp_tmae_ben = False
  
        vgSql = "SELECT NUM_POLIZA FROM PP_TMAE_BEN b WHERE b.num_poliza = '" & pnum_poliza & "'"
        
        Set vgRs2 = vgConexionBD.Execute(vgSql)
        If Not vgRs2.EOF Then
            fl_existe_poliza_in_pp_tmae_ben = True
        End If
        
        vgRs2.Close
        
       
Exit Function
Err_fl_existe_poliza_in_pp_tmae_ben:
    Screen.MousePointer = 0
    fl_existe_poliza_in_pp_tmae_ben = False
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'Function fl_existe_poliza_in_pp_tmae_endoso(pnum_poliza As String) As Boolean
'On Error GoTo Err_fl_existe_poliza_in_pp_tmae_endoso
'        fl_existe_poliza_in_pp_tmae_endoso = False
'
'        vgSql = "select num_poliza from pp_tmae_endoso e where e.num_poliza = '" & pnum_poliza & "' and  num_endoso = (select max(num_endoso) from pp_tmae_endoso where num_poliza = e.num_poliza)"
'
'        Set vgRs2 = vgConexionBD.Execute(vgSql)
'        If Not vgRs2.EOF Then
'            fl_existe_poliza_in_pp_tmae_endoso = True
'        End If
'
'        vgRs2.Close
'
'Exit Function
'Err_fl_existe_poliza_in_pp_tmae_endoso:
'    Screen.MousePointer = 0
'    fl_existe_poliza_in_pp_tmae_endoso = False
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
'End Function

Function fl_Insertar_temp_beneficiario(ByVal pnro_orden As Integer) As Boolean
On Error GoTo Err_fl_Insertar_temp_beneficiario
fl_Insertar_temp_beneficiario = False

    vgSql = ""
    vgSql = " insert into PP_TMAE_ENDBEN (num_Poliza, num_endoso, num_orden,num_idenben,cod_tipoidenben,gls_nomben,gls_nomsegben,gls_patben,gls_matben,cod_grufam,cod_par, "
    vgSql = vgSql & " cod_sexo,cod_sitinv,cod_dercre,cod_derpen,fec_nacben,mto_pensiongar,cod_estpension,fec_inipagopen,fec_terpagopengar,prc_pensiongar, "
    vgSql = vgSql & " gls_telben2,cod_direccion,gls_dirben,gls_fonoben,gls_correoben, cod_cauinv, cod_motreqpen, mto_pension, prc_pension)"
    
    vgSql = vgSql & " VALUES( "
   
    vgSql = vgSql & "'" & vlNum_Poliza & "' "
    vgSql = vgSql & ",'" & Lbl_EndCrear & "' "
    vgSql = vgSql & ",'" & ((lblRowsAfecctedBen) + pnro_orden) & "' "
    vgSql = vgSql & ",'" & vlnum_idenben & "' "
    vgSql = vgSql & ",'" & vlcod_tipoidenben & "' "
    vgSql = vgSql & ",'" & vlgls_nomben & "' "
    vgSql = vgSql & ",'" & vlgls_nomsegben & "' "
    vgSql = vgSql & ",'" & vlgls_patben & "' "
    vgSql = vgSql & ",'" & vlgls_matben & "' "
    vgSql = vgSql & ",'" & vlcod_grufam & "' "
    vgSql = vgSql & ",'" & vlcod_par & "' "
    vgSql = vgSql & ",'" & vlcod_sexo & "' "
    vgSql = vgSql & ",'" & vlcod_sitinv & "' "
    vgSql = vgSql & ",'" & vlcod_dercre & "' "
    vgSql = vgSql & ",'" & vlcod_derpen & "' "
    vgSql = vgSql & ",'" & Format(CDate(Trim(vlfec_nacben)), "yyyymmdd") & "' "
    vgSql = vgSql & "," & IIf(vlmto_pensiongar = "", 0, vlmto_pensiongar) & " "
    vgSql = vgSql & ",'" & vlcod_estpension & "' "
    vgSql = vgSql & ",'" & vlfec_inipagopen & "' "
    vgSql = vgSql & ",'" & vlfec_terpagopengar & "' "
    vgSql = vgSql & ",'" & IIf(vlprc_pensiongar = "", 0, vlprc_pensiongar) & "' "
    vgSql = vgSql & ",'" & vlgls_telben2 & "' "
    vgSql = vgSql & ",'" & vlcod_direccion & "' "
    vgSql = vgSql & ",'" & vlgls_dirben & "' "
    vgSql = vgSql & ",'" & vlgls_fonoben & "' "
    vgSql = vgSql & ",'" & vlgls_correoben & "' "
    
    vgSql = vgSql & ",'" & vlcod_cauinv & "' "
    vgSql = vgSql & ",'" & vlcod_motreqpen & "' "
    vgSql = vgSql & ",'" & vlmto_pension & "' "
    vgSql = vgSql & ",'" & vlprc_pension & "' "
    
    vgSql = vgSql & " ) "
  
    vgConectarBD.Execute vgSql
    
    fl_Insertar_temp_beneficiario = True
  
Exit Function
Err_fl_Insertar_temp_beneficiario:
    Screen.MousePointer = 0
    fl_Insertar_temp_beneficiario = False
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fl_insertar_temp_IN_select_pp_tmae_endben() As Boolean
On Error GoTo Err_fl_Insertar_temp_beneficiario
fl_insertar_temp_IN_select_pp_tmae_endben = False
 
    vgSql = "insert into pp_tmae_endben"
    vgSql = vgSql & "("
    vgSql = vgSql & " num_Poliza, num_endoso, num_orden,num_idenben,cod_tipoidenben,gls_nomben,gls_nomsegben,gls_patben,gls_matben,cod_grufam,cod_par,"
    vgSql = vgSql & " cod_sexo,cod_sitinv,cod_dercre,cod_derpen,fec_nacben,mto_pensiongar,cod_estpension,fec_inipagopen,fec_terpagopengar,prc_pensiongar,"
    vgSql = vgSql & " Gls_Telben2 , Cod_Direccion, Gls_DirBen, Gls_FonoBen, Gls_CorreoBen, Cod_CauInv, Cod_MotReqPen, Mto_Pension, Prc_Pension"
    vgSql = vgSql & " )"
    vgSql = vgSql & " select"
    vgSql = vgSql & " num_Poliza, " & Lbl_EndCrear & " as num_endoso , num_orden,num_idenben,cod_tipoidenben,gls_nomben,gls_nomsegben,gls_patben,gls_matben,cod_grufam,cod_par,"
    vgSql = vgSql & " cod_sexo,cod_sitinv,cod_dercre,cod_derpen,fec_nacben,mto_pensiongar,cod_estpension,fec_inipagopen,fec_terpagopengar,prc_pensiongar,"
    vgSql = vgSql & " Gls_Telben2 , Cod_Direccion, Gls_DirBen, Gls_FonoBen, Gls_CorreoBen, Cod_CauInv, Cod_MotReqPen, Mto_Pension, Prc_Pension"
    vgSql = vgSql & " From pp_tmae_ben"
    vgSql = vgSql & " where num_poliza ='" & Txt_PenPoliza.Text & "' and num_endoso = " & Lbl_EndosoActual

    vgConectarBD.Execute vgSql, lRecordsAffected
    
    lblRowsAfecctedBen = lRecordsAffected
    
    fl_insertar_temp_IN_select_pp_tmae_endben = True
  
Exit Function
Err_fl_Insertar_temp_beneficiario:
    Screen.MousePointer = 0
    fl_insertar_temp_IN_select_pp_tmae_endben = False
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fl_Actualizar_temp_beneficiario() As Boolean
'*On Error GoTo Err_flInsertarBeneficiario

    fl_Actualizar_temp_beneficiario = False
   
    vgSql = ""
    vgSql = " UPDATE PP_TMAE_ENDBEN SET "
    vgSql = vgSql & " cod_TipoIdenben = " & vlcod_tipoidenben & ", "
    vgSql = vgSql & " num_Idenben = '" & vlnum_idenben & "', "
    vgSql = vgSql & " gls_nomben = '" & vlgls_nomben & "', "
    vgSql = vgSql & " gls_patben = '" & vlgls_patben & "', "
    vgSql = vgSql & " gls_matben = '" & vlgls_matben & "', "
    vgSql = vgSql & " cod_sexo = '" & vlcod_sexo & "', "
    vgSql = vgSql & " mto_pension = " & vlmto_pension & ", "
    vgSql = vgSql & " fec_nacben = '" & Format(CDate(Trim(vlfec_nacben)), "yyyymmdd") & "', "
    vgSql = vgSql & " cod_estpension = '" & vlcod_estpension & "', "
    vgSql = vgSql & " fec_inipagopen = '" & vlfec_inipagopen & "' "
    vgSql = vgSql & " WHERE "
    vgSql = vgSql & " num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
    vgSql = vgSql & " num_endoso = " & Trim(Lbl_EndosoActual) & " AND "
    vgSql = vgSql & " num_orden = " & vlnum_orden & " "
    vgConectarBD.Execute vgSql
   
   fl_Actualizar_temp_beneficiario = True

Exit Function
Err_flInsertarBeneficiario:
    Screen.MousePointer = 0
    fl_Actualizar_temp_beneficiario = False
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function



Function fl_insert_update_ItemGridBeneficiario()
On Error GoTo Err_fl_insert_update_ItemGridBeneficiario

        vlnum_orden = Trim(Msf_BMGrilla.rows)
        vlcod_tipoidenben = Trim(Mid(Cmb_BMNumIdent, 1, (InStr(1, Cmb_BMNumIdent, "-") - 1)))
        vlnum_idenben = Txt_BMNumIden.Text
        vlgls_nomben = txtNomben.Text
        vlgls_nomsegben = txtNomsegBen.Text
        vlgls_patben = txtApepatBen.Text
        vlgls_matben = txtApematBen.Text
        vlcod_grufam = "01"
        vlcod_par = Trim(Mid(Cmb_BMPar, 1, (InStr(1, Cmb_BMPar, "-") - 1)))
        vlcod_sexo = Trim(Mid(Cmb_BMSexo, 1, (InStr(1, Cmb_BMSexo, "-") - 1)))
        vlcod_sitinv = "N"
        vlcod_dercre = "N"
        vlcod_derpen = Trim(Mid(Cmb_BMDerPen, 1, (InStr(1, Cmb_BMDerPen, "-") - 1)))
        vlfec_nacben = Txt_BMFecNac.Text
        vlmto_pensiongar = 0
        If Txt_BMMtoPensionGar.Text <> "" Then
            vlmto_pensiongar = Txt_BMMtoPensionGar.Text
        End If
        vlcod_estpension = Trim(Mid(cmbExcluesion, 1, (InStr(1, cmbExcluesion, "-") - 1)))
        vlfec_inipagopen = flObtieneFechaEfecto
        vlfec_terpagopengar = flObtieneFec_terpagopengar(Txt_PenPoliza.Text)
        
        vlprc_pensiongar = 0
        If chkContinuarpago.Value <> 1 Then
           vlprc_pensiongar = txt_BMPrcGar.Text
        End If
        
        vlgls_telben2 = txtTelfon2.Text
        vlcod_direccion = Trim(IIf(txtcodDirBen = "", 0, txtcodDirBen))
        vlgls_dirben = txt_dirben.Text
        vlgls_fonoben = txtTelBen.Text
        vlgls_correoben = txtBenEmail.Text
        
        If Lbl_BMNumOrd = "" Then
        
        Msf_BMGrilla.AddItem vlnum_orden & vbTab & vlcod_tipoidenben & vbTab & vlnum_idenben & vbTab & vlgls_nomben & vbTab & _
        vlgls_nomsegben & vbTab & vlgls_patben & vbTab & vlgls_matben & vbTab & vlcod_grufam & vbTab & vlcod_par & vbTab & vlcod_sexo & vbTab & _
        vlcod_sitinv & vbTab & vlcod_dercre & vbTab & vlcod_derpen & vbTab & vlfec_nacben & vbTab & vlmto_pensiongar & vbTab & vlcod_estpension & _
        vbTab & vlfec_inipagopen & vbTab & vlfec_terpagopengar & vbTab & vlprc_pensiongar & vbTab & vlgls_telben2 & vbTab & vlcod_direccion & vbTab & _
        vlgls_dirben & vbTab & vlgls_fonoben & vbTab & vlgls_correoben
        
        Lbl_BMNumOrd = vlnum_orden
        
          Else
          
          Msf_BMGrilla.row = Lbl_BMNumOrd
    '------
    Msf_BMGrilla.Col = 0
    Msf_BMGrilla.Text = Lbl_BMNumOrd
    Msf_BMGrilla.Col = 1
    Msf_BMGrilla.Text = vlcod_tipoidenben
    Msf_BMGrilla.Col = 2
    Msf_BMGrilla.Text = vlnum_idenben
    Msf_BMGrilla.Col = 3
    Msf_BMGrilla.Text = vlgls_nomben
    Msf_BMGrilla.Col = 4
    Msf_BMGrilla.Text = vlgls_nomsegben
    Msf_BMGrilla.Col = 5
    Msf_BMGrilla.Text = vlgls_patben
    Msf_BMGrilla.Col = 6
    Msf_BMGrilla.Text = vlgls_matben
    Msf_BMGrilla.Col = 7
    Msf_BMGrilla.Text = vlcod_grufam
    Msf_BMGrilla.Col = 8
    Msf_BMGrilla.Text = vlcod_par
    Msf_BMGrilla.Col = 9
    Msf_BMGrilla.Text = vlcod_sexo
    Msf_BMGrilla.Col = 10
    Msf_BMGrilla.Text = vlcod_sitinv
    Msf_BMGrilla.Col = 11
    Msf_BMGrilla.Text = vlcod_dercre
    Msf_BMGrilla.Col = 12
    Msf_BMGrilla.Text = vlcod_derpen
    Msf_BMGrilla.Col = 13
    Msf_BMGrilla.Text = vlfec_nacben
    Msf_BMGrilla.Col = 14
    Msf_BMGrilla.Text = vlmto_pensiongar
    Msf_BMGrilla.Col = 15
    Msf_BMGrilla.Text = vlcod_estpension
    Msf_BMGrilla.Col = 16
    Msf_BMGrilla.Text = vlfec_inipagopen
    Msf_BMGrilla.Col = 17
    Msf_BMGrilla.Text = vlfec_terpagopengar
    Msf_BMGrilla.Col = 18
    Msf_BMGrilla.Text = vlprc_pensiongar
    Msf_BMGrilla.Col = 19
    Msf_BMGrilla.Text = vlgls_telben2
    Msf_BMGrilla.Col = 20
    Msf_BMGrilla.Text = vlcod_direccion
    Msf_BMGrilla.Col = 21
    Msf_BMGrilla.Text = vlgls_dirben
    Msf_BMGrilla.Col = 22
    Msf_BMGrilla.Text = vlgls_fonoben
    Msf_BMGrilla.Col = 23
    Msf_BMGrilla.Text = vlgls_correoben
          
        End If
      
  
   
Exit Function
Err_fl_insert_update_ItemGridBeneficiario:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fl_aprobar_insertar_IN_pp_tmae_endoso()
'*On Error GoTo Err_flInsertarEndoso
Dim vlExiste As Integer

vlFecCrea = Format(Date, "yyyymmdd")
vlHorCrea = Format(Time, "hhmmss")
fl_aprobar_insertar_IN_pp_tmae_endoso = False

vgSql = "select count(*) existe from PP_TMAE_ENDOSO where num_poliza = '" & Txt_PenPoliza.Text & "'"
vgConectarBD.Execute vgSql

Set vgRs2 = vgConexionBD.Execute(vgSql)
If Not vgRs2.EOF Then
    vlExiste = vgRs2!existe
Else
    
End If

If vlExiste <> 0 Then
    vgSql = " insert into PP_TMAE_ENDOSO "
    vgSql = vgSql & " select num_poliza,num_endoso + 1 ,'" & vlFecCrea & "','" & vlFecCrea & "','29','S',mto_diferencia,cod_moneda," & txtmto_pensioncal.Text & "," & txtmto_pensioncal.Text & ",'" & vlFecCrea & "',prc_factor,'" & txtgls_docjud.Text & "',"
    vgSql = vgSql & " cod_usuariocrea,fec_crea,hor_crea,'" & vgUsuario & "','" & vlFecCrea & "','" & vlHorCrea & "',fec_finefecto,'L',cod_tipreajuste,mto_valreajustetri,mto_valreajustemen from PP_TMAE_ENDOSO"
    vgSql = vgSql & " where num_poliza = '" & Txt_PenPoliza.Text & "' and num_endoso = " & Lbl_EndosoActual - 1
Else
    vgSql = " insert into PP_TMAE_ENDOSO "
    vgSql = vgSql & " select num_poliza,1,'" & vlFecCrea & "','" & vlFecCrea & "','29', 'S', 0, cod_moneda," & txtmto_pensioncal.Text & "," & txtmto_pensioncal.Text & ",'" & vlFecCrea & "', 0,'" & txtgls_docjud.Text & "',"
    vgSql = vgSql & " '" & vgUsuario & "','" & vlFecCrea & "','" & vlHorCrea & "',cod_usuariomodi,fec_modi,hor_modi,null, 'L', cod_tipreajuste,mto_valreajustetri,mto_valreajustemen"
    vgSql = vgSql & " from pp_tmae_poliza where num_poliza = '" & Txt_PenPoliza.Text & "' and num_endoso = " & Lbl_EndosoActual
End If

vgRs2.Close
vgConectarBD.Execute vgSql
  
fl_aprobar_insertar_IN_pp_tmae_endoso = True
    
Exit Function
Err_flInsertarEndoso:
    Screen.MousePointer = 0
    fl_aprobar_insertar_IN_pp_tmae_endoso = False
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function fl_existe_tmae_liqpagopendef() As Boolean

On Error GoTo Err_ValidaFechaEfecto
        
        fl_existe_tmae_liqpagopendef = False
        vgSql = "select num_perpago from pp_tmae_liqpagopendef where num_poliza = '" & Txt_PenPoliza.Text & "' and num_perpago = '" & Format(Date, "yyyymm") & "' and cod_tipopago='H'"
        Set vgRs2 = vgConexionBD.Execute(vgSql)
        If Not vgRs2.EOF Then
            fl_existe_tmae_liqpagopendef = True
        End If
        vgRs2.Close

Exit Function
Err_ValidaFechaEfecto:
    Screen.MousePointer = 0
    fl_existe_tmae_liqpagopendef = False
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fl_aprobar_insertar_IN_pp_tmae_liqpagopendef() As Boolean
'*On Error GoTo Err_flInsertarEndoso

fl_aprobar_insertar_IN_pp_tmae_liqpagopendef = False

Dim vlcod_tippension As String
vlcod_tippension = Trim(Mid(Cmb_PMTipPen, 1, (InStr(1, Cmb_PMTipPen, "-") - 1)))

vgSql = "insert into pp_tmae_liqpagopendef("
vgSql = vgSql & "num_perpago,num_poliza,num_endoso,num_orden,cod_tipopago,gls_direccion,cod_direccion,cod_tippension,fec_pago,cod_viapago,num_idenreceptor,"
vgSql = vgSql & "cod_tipoidenreceptor,gls_nomreceptor,gls_nomsegreceptor,gls_patreceptor,gls_matreceptor,cod_tipreceptor,mto_haber,mto_liqpagar,mto_pension,"
vgSql = vgSql & "cod_moneda) values ("

vgSql = vgSql & "  '" & Format(Date, "yyyymm") & "'"
vgSql = vgSql & ", '" & Txt_PenPoliza.Text & "'"
vgSql = vgSql & ", " & CInt(Lbl_EndosoActual) + 1 & ""
vgSql = vgSql & ", " & "1"
vgSql = vgSql & ", '" & "H" & "'"
vgSql = vgSql & ", '" & vlxgls_dirben & "'"
vgSql = vgSql & ", '" & vlxcod_direccion & "'"
vgSql = vgSql & ", '" & vlcod_tippension & "'"
vgSql = vgSql & ", '" & Format(Date, "yyyymmdd") & "'"
vgSql = vgSql & ", '" & "04" & "'"
vgSql = vgSql & ", '" & vlxnum_idenben & "'"
vgSql = vgSql & ", '" & "1" & "'"
vgSql = vgSql & ", '" & vlxgls_nomben & "'"
vgSql = vgSql & ", '" & vlxgls_nomsegben & "'"
vgSql = vgSql & ", '" & vlxgls_patben & "'"
vgSql = vgSql & ", '" & vlxgls_matben & "'"
vgSql = vgSql & ", '" & "H" & "'"
vgSql = vgSql & ", " & txtmto_pensioncal.Text & ""
vgSql = vgSql & ", " & txtmto_pensioncal.Text & ""
vgSql = vgSql & ", " & txtmto_pensioncal.Text & ""
vgSql = vgSql & ", '" & "NS" & "')"


vgConectarBD.Execute vgSql
  
fl_aprobar_insertar_IN_pp_tmae_liqpagopendef = True
    
Exit Function
Err_flInsertarEndoso:
    Screen.MousePointer = 0
    fl_aprobar_insertar_IN_pp_tmae_liqpagopendef = False
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function fl_grabar_insertar_IN_pp_tmae_endendoso()
'*On Error GoTo Err_flInsertarEndoso

fl_grabar_insertar_IN_pp_tmae_endendoso = False
    
    vgSql = "DELETE FROM PP_TMAE_ENDENDOSO WHERE num_poliza ='" & Txt_PenPoliza.Text & "' and num_endoso=" & Lbl_EndosoActual
    vgConectarBD.Execute vgSql
    
    vgSql = "INSERT INTO PP_TMAE_ENDENDOSO(num_poliza,num_endoso,fec_solendoso,cod_cauendoso,cod_tipendoso,"
    vgSql = vgSql & "cod_moneda,mto_pensionori,mto_pensioncal ,fec_efecto,cod_usuariocrea,fec_crea,hor_crea, "
    vgSql = vgSql & "cod_estado, num_mesgarres,gls_docjud, FEC_ENDOSO) "
    vgSql = vgSql & "values("
    vgSql = vgSql & "'" & Txt_PenPoliza.Text & "',"
    vgSql = vgSql & "'" & Lbl_EndosoActual & "',"
    vgSql = vgSql & "'" & Format(CDate(Trim(Date)), "yyyymmdd") & "',"
    vgSql = vgSql & "'" & "28" & "',"
    vgSql = vgSql & "'" & "S" & "',"
    vgSql = vgSql & "'" & vlcod_moneda & "',"
    vgSql = vgSql & "" & vlmto_pensioncal & ","
    vgSql = vgSql & "" & txtmto_pensioncal.Text & ","
    vgSql = vgSql & "'" & flObtieneFechaEfecto & "',"
    vgSql = vgSql & "'" & vgUsuario & "',"
    vgSql = vgSql & "'" & Format(Date, "yyyymmdd") & "',"
    vgSql = vgSql & "'" & Format(Time, "hhmmss") & "',"
    vgSql = vgSql & "'" & "P" & "',"
    vgSql = vgSql & "'" & txtnum_mesgarres.Text & "',"
    vgSql = vgSql & "'" & txtgls_docjud.Text & "',"
    vgSql = vgSql & "'" & Format(CDate(Trim(Date)), "yyyymmdd") & "'"
    vgSql = vgSql & ")"
 
vgConectarBD.Execute vgSql
  
fl_grabar_insertar_IN_pp_tmae_endendoso = True

    
Exit Function
Err_flInsertarEndoso:
    Screen.MousePointer = 0
    fl_grabar_insertar_IN_pp_tmae_endendoso = False
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function fl_aprobar_insertar_IN_pp_tmae_ben() As Boolean
'*On Error GoTo Err_flActualizarEndoso
fl_aprobar_insertar_IN_pp_tmae_ben = False

vlFecCrea = Format(Date, "yyyymmdd")
vlHorCrea = Format(Time, "hhmmss")

vgSql = ""
vgSql = "insert into pp_tmae_ben ("
vgSql = vgSql & " num_poliza,num_endoso,num_orden,fec_ingreso,cod_tipoidenben,num_idenben,gls_nomben,gls_nomsegben,gls_patben,gls_matben,gls_dirben,cod_direccion,gls_fonoben,gls_correoben,cod_grufam,cod_par,cod_sexo,cod_sitinv,cod_dercre,"
vgSql = vgSql & " cod_derpen,cod_cauinv,fec_nacben,fec_nachm,fec_invben,cod_motreqpen,mto_pension,mto_pensiongar,prc_pension,prc_pensionleg,cod_inssalud,cod_modsalud,mto_plansalud,cod_estpension,cod_viapago,cod_banco,cod_tipcuenta,num_cuenta,"
vgSql = vgSql & " cod_sucursal,fec_fallben,fec_matrimonio,cod_caususben,fec_susben,fec_inipagopen,fec_terpagopengar,cod_usuariocrea,fec_crea,hor_crea,cod_usuariomodi,fec_modi,hor_modi,prc_pensiongar,gls_telben2,cod_tipcta,cod_monbco,num_ctabco,"
vgSql = vgSql & " ind_bolelec,num_cuenta_cci,cons_trainfo,cons_datcomer )"
vgSql = vgSql & " SELECT num_poliza,num_endoso + 1,num_orden,fec_ingreso,cod_tipoidenben,num_idenben,gls_nomben,gls_nomsegben,gls_patben,gls_matben,gls_dirben,cod_direccion,gls_fonoben,gls_correoben,cod_grufam,cod_par,cod_sexo,cod_sitinv,cod_dercre,"
vgSql = vgSql & " 10,cod_cauinv,fec_nacben,fec_nachm,fec_invben,cod_motreqpen,0,0,prc_pension,prc_pensionleg,cod_inssalud,cod_modsalud,mto_plansalud,10,cod_viapago,cod_banco,cod_tipcuenta,num_cuenta,"
vgSql = vgSql & " cod_sucursal,fec_fallben,fec_matrimonio,cod_caususben,fec_susben,fec_inipagopen,fec_terpagopengar,cod_usuariocrea,fec_crea,hor_crea,'" & vgUsuario & "','" & vlFecCrea & "','" & vlHorCrea & "',prc_pensiongar,gls_telben2,cod_tipcta,cod_monbco,num_ctabco,"
vgSql = vgSql & " ind_bolelec,num_cuenta_cci,cons_trainfo,cons_datcomer FROM pp_tmae_ben where num_poliza = '" & Txt_PenPoliza & "' and num_endoso = " & Lbl_EndosoActual
vgConectarBD.Execute (vgSql)
fl_aprobar_insertar_IN_pp_tmae_ben = True

Exit Function
Err_flActualizarEndoso:
    Screen.MousePointer = 0
    fl_aprobar_insertar_IN_pp_tmae_ben = False
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fl_aprobar_eliminar_tablas_temp() As Boolean
'*On Error GoTo Err_flEliminarRegistrosTempEnd
    fl_aprobar_eliminar_tablas_temp = False
   
    vgSql = "DELETE PP_TMAE_ENDBEN "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "num_poliza = '" & Txt_PenPoliza.Text & "' AND "
    vgSql = vgSql & "num_endoso = " & Lbl_EndCrear & " "
    vgConectarBD.Execute (vgSql)

    
    vgSql = "DELETE PP_TMAE_ENDPOLIZA "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "num_poliza = '" & Txt_PenPoliza.Text & "' AND "
    vgSql = vgSql & "num_endoso = " & Lbl_EndCrear & " "
    vgConectarBD.Execute (vgSql)
    
    
    vgSql = "DELETE PP_TMAE_ENDENDOSO "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "num_poliza = '" & Txt_PenPoliza.Text & "' AND "
    vgSql = vgSql & "num_endoso = " & Lbl_EndosoActual & " "
    vgConectarBD.Execute (vgSql)
    
     fl_aprobar_eliminar_tablas_temp = True
     
Exit Function
Err_flEliminarRegistrosTempEnd:
    Screen.MousePointer = 0
     fl_aprobar_eliminar_tablas_temp = False
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function
'Call fl_eliminar_registros_temp(Txt_PenPoliza, Lbl_EndosoActual)

Function fl_aprobar_insertar_IN_pp_tmae_poliza() As Boolean
'*On Error GoTo Err_flInsertarPoliza

fl_aprobar_insertar_IN_pp_tmae_poliza = False

vgSql = " insert into pp_tmae_poliza("
vgSql = vgSql & " Num_poliza,Num_endoso,Cod_afp,Cod_tippension,Cod_estado,Cod_tipren,Cod_modalidad,Num_cargas,Fec_vigencia,Fec_tervigencia,Mto_prima,Mto_pension,Mto_pensiongar,Num_mesdif,Num_mesgar,Prc_tasace,Prc_tasavta,Prc_tasactorea,Prc_tasaintpergar,Fec_inipagopen,Cod_usuariocrea,"
vgSql = vgSql & " Fec_crea,Hor_crea,Cod_usuariomodi,Fec_modi,Hor_modi,Cod_tiporigen,Num_indquiebra,Cod_cuspp,Ind_cob,Cod_moneda,Mto_valmoneda,Cod_cobercon,Mto_facpenella,Prc_facpenella,Cod_dercre,Cod_dergra,Prc_tasatir,Fec_emision,Fec_dev,Fec_inipencia,Fec_pripago,Fec_finperdif,Fec_finpergar,"
vgSql = vgSql & " Fec_efecto,Cod_tipreajuste,Mto_valreajustetri,Mto_valreajustemen,Fec_devsol,Ind_bendes,Ind_bolelec,Ind_heren)"
vgSql = vgSql & " select"
vgSql = vgSql & " Num_poliza,(num_endoso) + 1 as num_endoso,Cod_afp,Cod_tippension,9,Cod_tipren,Cod_modalidad,Num_cargas,Fec_vigencia,Fec_tervigencia,Mto_prima,Mto_pension,Mto_pensiongar,Num_mesdif,Num_mesgar,Prc_tasace,Prc_tasavta,Prc_tasactorea,Prc_tasaintpergar,Fec_inipagopen,"
vgSql = vgSql & " Cod_usuariocrea,Fec_crea,Hor_crea,Cod_usuariomodi,Fec_modi,Hor_modi,Cod_tiporigen,Num_indquiebra,Cod_cuspp,Ind_cob,Cod_moneda,Mto_valmoneda,Cod_cobercon,Mto_facpenella,Prc_facpenella,Cod_dercre,Cod_dergra,Prc_tasatir,Fec_emision,Fec_dev,Fec_inipencia,Fec_pripago,Fec_finperdif,"
vgSql = vgSql & " Fec_finpergar , Fec_efecto, Cod_TipReajuste, Mto_ValReajusteTri, Mto_ValReajusteMen, Fec_devsol, Ind_bendes, ind_bolelec, 2"
vgSql = vgSql & " From pp_tmae_poliza where num_poliza = '" & Txt_PenPoliza.Text & "' and num_endoso = " & Lbl_EndosoActual
vgConectarBD.Execute vgSql

fl_aprobar_insertar_IN_pp_tmae_poliza = True
Exit Function
Err_flInsertarPoliza:
    Screen.MousePointer = 0
    fl_aprobar_insertar_IN_pp_tmae_poliza = False
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fl_aprobar() As Boolean
On Error GoTo Err_Aprobar
   
    fl_aprobar = False
    
    If Not fgConexionBaseDatos(vgConectarBD) Then
        MsgBox "Falló la conexión con la Base de Datos.", vbCritical, "Error de Conexión"
        Exit Function
    End If
    'Comenzar la Transacción
    vgConectarBD.BeginTrans
   
    If fl_aprobar_insertar_IN_pp_tmae_poliza = False Then Call fl_begin_trans_rollbak
    If fl_aprobar_insertar_IN_pp_tmae_endoso = False Then Call fl_begin_trans_rollbak
    If fl_aprobar_insertar_IN_pp_tmae_ben = False Then Call fl_begin_trans_rollbak
    'If fl_aprobar_eliminar_tablas_temp = False Then Call fl_begin_trans_rollbak
    If fl_existe_tmae_liqpagopendef = False Then If fl_aprobar_insertar_IN_pp_tmae_liqpagopendef = False Then Call fl_begin_trans_rollbak
        
    vgConectarBD.CommitTrans
    vgConectarBD.Close
    
    fl_aprobar = True
    
Exit Function
Err_Aprobar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        'vgConectarBD.RollbackTrans
        'vgConectarBD.Close
        fl_aprobar = False
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'dextre
Function flComboTipoIdentificacion(vlCombo As ComboBox)
Dim vlRegCombo As ADODB.Recordset
Dim vlcont As Long
On Error GoTo Err_Combo

    vlCombo.Clear
    vgSql = ""
    vlcont = 0
    vgQuery = "SELECT 0 as codigo, 'Seleccione' as Nombre, 8 as largo FROM dual union "
    vgQuery = vgQuery & "SELECT cod_tipoiden as codigo, gls_tipoidencor as Nombre, "
    vgQuery = vgQuery & "num_lartipoiden as largo "
    vgQuery = vgQuery & "FROM MA_TPAR_TIPOIDEN "
    vgQuery = vgQuery & "WHERE cod_tipoiden <> 0 "
    vgQuery = vgQuery & "ORDER BY codigo "
    Set vlRegCombo = vgConexionBD.Execute(vgQuery)
    If Not (vlRegCombo.EOF) Then
        While Not (vlRegCombo.EOF)
            vlCombo.AddItem Space(2 - Len(vlRegCombo!Codigo)) & vlRegCombo!Codigo & " - " & (Trim(vlRegCombo!Nombre))
            vlcont = vlCombo.ListCount - 1
            vlCombo.ItemData(vlcont) = (vlRegCombo!largo)
            vlRegCombo.MoveNext
        Wend
    End If
    vlRegCombo.Close
        
    'Colocar el Tipo de Identificación "Sin Información" al final
    vgQuery = "SELECT cod_tipoiden as codigo, gls_tipoidencor as Nombre, "
    vgQuery = vgQuery & "num_lartipoiden as largo "
    vgQuery = vgQuery & "FROM MA_TPAR_TIPOIDEN "
    vgQuery = vgQuery & "WHERE cod_tipoiden = 0 "
    vgQuery = vgQuery & "ORDER BY codigo "
    Set vlRegCombo = vgConexionBD.Execute(vgQuery)
    If Not (vlRegCombo.EOF) Then
        While Not (vlRegCombo.EOF)
            vlCombo.AddItem Space(2 - Len(vlRegCombo!Codigo)) & vlRegCombo!Codigo & " - " & (Trim(vlRegCombo!Nombre))
            vlcont = vlCombo.ListCount - 1
            vlCombo.ItemData(vlcont) = (vlRegCombo!largo)
            vlRegCombo.MoveNext
        Wend
    End If
    vlRegCombo.Close
        
    If vlCombo.ListCount <> 0 Then
        vlCombo.ListIndex = 0
    End If

Exit Function
Err_Combo:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fl_existe_beneficiario(pnum_orden As Integer) As String

On Error GoTo Err_ValidaFechaEfecto
        
        fl_existe_beneficiario = False
        vgSql = "select num_poliza from pp_tmae_endben where num_poliza = '" & Txt_PenPoliza & "' and num_endoso = " & Lbl_EndosoActual & " and num_orden = " & pnum_orden
        Set vgRs2 = vgConexionBD.Execute(vgSql)
        If Not vgRs2.EOF Then
            fl_existe_beneficiario = True
        End If
        vgRs2.Close

Exit Function
Err_ValidaFechaEfecto:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'--------------------------------------------------------- print ----------------------------------------------



Function fl_imprimir_endoso()
On Error GoTo Err_flImprimirEndosoDef

    Dim rsLiq As ADODB.Recordset
    Dim objRep As New ClsReporte
    Dim LNGa As Long
    
    'Imprimir Reporte de Poliza Original
    Screen.MousePointer = 11
    
     vlRptNumEndosoPol = 0
    vlRptNumEndosoEnd = 0
    
   
    
        'Obtener el último Endoso creado
        vgSql = "SELECT num_endoso "
        vgSql = vgSql & "FROM PP_TMAE_POLIZA "
        vgSql = vgSql & "WHERE "
        vgSql = vgSql & "num_poliza = '" & vlGlobalNumPoliza & "' "
        vgSql = vgSql & "ORDER BY num_endoso DESC "
        Set vgRegistro = vgConexionBD.Execute(vgSql)
        If Not vgRegistro.EOF Then
           'Ultimo Endoso
           vlRptNumEndosoEnd = (vgRegistro!num_endoso)
        Else
            vgRegistro.Close
            MsgBox "No existen endosos para la Póliza indicada.", vbCritical, "Operación Cancelada"
            Exit Function
        End If
        vgRegistro.Close
        
        vlGlobalNumPoliza = Txt_PenPoliza.Text
        vlRptNumEndosoEnd = Lbl_EndosoActual
        
        'Obtener el Penúltimo Endoso
        vgSql = "SELECT num_endoso "
        vgSql = vgSql & "FROM PP_TMAE_POLIZA "
        vgSql = vgSql & "WHERE "
        vgSql = vgSql & "num_poliza = '" & vlGlobalNumPoliza & "' "
        vgSql = vgSql & "and num_endoso = " & vlRptNumEndosoEnd & " "
        vgSql = vgSql & "ORDER BY num_endoso DESC "
        Set vgRegistro = vgConexionBD.Execute(vgSql)
        If Not vgRegistro.EOF Then
            'Penultimo Endoso
            vlRptNumEndosoPol = (vgRegistro!num_endoso) - 1
        Else
            vgRegistro.Close
            MsgBox "No existe endoso anterior al " & CStr(vlRptNumEndosoEnd) & " para la Póliza indicada.", vbCritical, "Operación Cancelada"
            Exit Function
        End If
        vgRegistro.Close
      
   

   vlArchivo = strRpt & "PP_Rpt_EndDefEndoso_2.rpt"   '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
      MsgBox "Archivo de Reporte de Endoso de Renta Vitalicia no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
      Screen.MousePointer = 0
      Exit Function
   End If
      
     vlRptCodTipPension = ""
   'vlRptNumEndoso = 0
   vlRptCodAfp = ""
   vlRptCodTipRen = ""
   
   vgSql = ""
   vgSql = "SELECT num_endoso,cod_afp,cod_tipren,cod_tippension "
   vgSql = vgSql & ",cod_moneda,mto_valmoneda "
   vgSql = vgSql & "FROM PP_TMAE_POLIZA "
   vgSql = vgSql & "WHERE "
   vgSql = vgSql & "num_poliza = '" & vlGlobalNumPoliza & "' AND "
   vgSql = vgSql & "num_endoso = " & vlRptNumEndosoPol & " "
   'vgSql = vgSql & "ORDER BY num_endoso DESC "
   Set vgRegistro = vgConexionBD.Execute(vgSql)
   If Not vgRegistro.EOF Then
      'vlRptNumEndoso = (vgRegistro!num_endoso)
      vlRptCodAfp = Trim(vgRegistro!cod_afp) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_AFP, Trim(vgRegistro!cod_afp)))
      vlRptCodTipRen = Trim(vgRegistro!Cod_TipRen) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_TipRen, Trim(vgRegistro!Cod_TipRen)))
      vlRptCodTipPension = Trim(vgRegistro!Cod_TipPension) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_TipPen, Trim(vgRegistro!Cod_TipPension)))
      vlImpCodMoneda = vgRegistro!Cod_Moneda
      vlImpNomMoneda = fgBuscarGlosaElemento(vgCodTabla_TipMon, vlImpCodMoneda)
   End If
   vgRegistro.Close
   
   vlRptCodInsSalud = ""
   vlRptNomBen = ""
   vlRptRutBen = ""
   vlRptGlsDirBen = ""
   vlRptFonoBen = ""
   vlRptCodDireccion = "0"
   
   vgSql = ""
   vgSql = "SELECT cod_inssalud,gls_nomben,gls_nomsegben,gls_patben,gls_matben,Cod_TipoIdenben, "
   vgSql = vgSql & "Num_Idenben,cod_direccion,gls_dirben,gls_fonoben "
   vgSql = vgSql & "FROM PP_TMAE_BEN "
   vgSql = vgSql & "WHERE "
   vgSql = vgSql & "num_poliza = '" & vlGlobalNumPoliza & "' AND "
   vgSql = vgSql & "num_endoso = " & (vlRptNumEndosoPol) & " AND "
   vgSql = vgSql & "cod_par = '" & Trim(clCodParCau) & "' "
   
  Set vgRegistro = vgConexionBD.Execute(vgSql)
   If Not vgRegistro.EOF Then
        If Not IsNull(vgRegistro!Cod_InsSalud) Then
            vlRptCodInsSalud = " " & Trim(vgRegistro!Cod_InsSalud) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_InsSal, Trim(vgRegistro!Cod_InsSalud)))
        End If
        If Not IsNull(vgRegistro!Gls_NomBen) And Not IsNull(vgRegistro!Gls_PatBen) Then
            vlNombre = vgRegistro!Gls_NomBen
            vlNombreSeg = IIf(IsNull(vgRegistro!Gls_NomSegBen), "", vgRegistro!Gls_NomSegBen)
            vlPaterno = vgRegistro!Gls_PatBen
            'I - MC 24/01/2008
            ''vlMaterno = IIf(IsNull(vgRegistro!Gls_PatBen), "", vgRegistro!Gls_PatBen)
            vlMaterno = IIf(IsNull(vgRegistro!Gls_MatBen), "", vgRegistro!Gls_MatBen)
            'F - MC 24/01/2008
            vlNombreCompleto = fgFormarNombreCompleto(vlNombre, vlNombreSeg, vlPaterno, vlMaterno)
            
            vlRptNomBen = vlNombreCompleto
        End If
        If Not IsNull(vgRegistro!Cod_TipoIdenBen) And Not IsNull(vgRegistro!Num_IdenBen) Then
            vlNombreTipoIden = fgBuscarNombreTipoIden(vgRegistro!Cod_TipoIdenBen)
            vlRptRutBen = vlNombreTipoIden & " - " & (Trim(vgRegistro!Num_IdenBen))
        End If
        If Not IsNull(vgRegistro!Gls_DirBen) Then
            vlRptGlsDirBen = Trim(vgRegistro!Gls_DirBen)
        End If
        If Not IsNull(vgRegistro!Gls_FonoBen) Then
            vlRptFonoBen = Trim(vgRegistro!Gls_FonoBen)
        Else
            vlRptFonoBen = " "
        End If
        If Not IsNull(vgRegistro!Cod_Direccion) Then
            vlRptCodDireccion = Trim(vgRegistro!Cod_Direccion)
        End If
   End If
   vgRegistro.Close


'RRR DATOS DEL ULTIMO ENDOSO

   vlRptNomBenNue = ""
   vlRptGlsDirBenNue = ""
   vlRptFonoBenNue = ""
   vlRptCodDireccionNue = "0"
   vlRptNomBeNue = ""
   
   vgSql = ""
   vgSql = "SELECT cod_inssalud,gls_nomben,gls_nomsegben,gls_patben,gls_matben,Cod_TipoIdenben, "
   vgSql = vgSql & "Num_Idenben,cod_direccion,gls_dirben,gls_fonoben, gls_correoben "
   vgSql = vgSql & "FROM PP_TMAE_BEN "
   vgSql = vgSql & "WHERE "
   vgSql = vgSql & "num_poliza = '" & vlGlobalNumPoliza & "' AND "
   vgSql = vgSql & "num_endoso = " & (vlRptNumEndosoEnd) & " AND "
   'vgSql = vgSql & "cod_derpen = '" & Trim(clCodParCau) & "' AND " 'MateriaGris-JRios 11/01/2018
   vgSql = vgSql & "num_orden = " & Trim(Lbl_BMNumOrd) & ""
   Set vgRegistro = vgConexionBD.Execute(vgSql)
   If Not vgRegistro.EOF Then
       
        'While Not vgRegistro.EOF
        
            If Not IsNull(vgRegistro!Gls_NomBen) And Not IsNull(vgRegistro!Gls_PatBen) Then
                vlNombre = vgRegistro!Gls_NomBen
                vlNombreSeg = IIf(IsNull(vgRegistro!Gls_NomSegBen), "", vgRegistro!Gls_NomSegBen)
                vlPaterno = vgRegistro!Gls_PatBen
                'I - MC 24/01/2008
                ''vlMaterno = IIf(IsNull(vgRegistro!Gls_PatBen), "", vgRegistro!Gls_PatBen)
                vlMaterno = IIf(IsNull(vgRegistro!Gls_MatBen), "", vgRegistro!Gls_MatBen)
                'F - MC 24/01/2008
                vlNombreCompleto = fgFormarNombreCompleto(vlNombre, vlNombreSeg, vlPaterno, vlMaterno)
                
                vlRptNomBeNue = vlNombreCompleto 'vlRptNomBeNue & vbCrLf & vlNombreCompleto
            End If
            'vgRegistro.MoveNext
        'Wend
        'vgRegistro.MoveFirst
          
        If Not IsNull(vgRegistro!Gls_DirBen) Then
            vlRptGlsDirBenNue = Trim(vgRegistro!Gls_DirBen)
        End If
        If Not IsNull(vgRegistro!Gls_FonoBen) Then
            vlRptFonoBenNue = Trim(vgRegistro!Gls_FonoBen)
        Else
            vlRptFonoBenNue = " "
        End If
        If Not IsNull(vgRegistro!Cod_Direccion) Then
            vlRptCodDireccionNue = Trim(vgRegistro!Cod_Direccion)
        End If
        If Not IsNull(vgRegistro!Gls_CorreoBen) Then
            vlRptCorreo = Trim(vgRegistro!Gls_CorreoBen)
        End If
        If Not IsNull(vgRegistro!Num_IdenBen) Then
            vlRptDocNum = Trim(vgRegistro!Num_IdenBen)
        End If
 
   End If
   vgRegistro.Close

    vgSql = ""
    vgSql = "SELECT c.gls_comuna,p.gls_provincia,r.gls_region "
    vgSql = vgSql & "FROM MA_TPAR_COMUNA c,MA_TPAR_PROVINCIA p,MA_TPAR_REGION r "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "c.cod_direccion = '" & vlRptCodDireccionNue & "' AND "
    vgSql = vgSql & "c.cod_provincia = p.cod_provincia AND "
    vgSql = vgSql & "p.cod_region = r.cod_region AND "
    vgSql = vgSql & "c.cod_region = r.cod_region "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
         If Not IsNull(vgRegistro!gls_comuna) Then vlRptComuna = Trim(vgRegistro!gls_comuna)
         If Not IsNull(vgRegistro!gls_provincia) Then vlRptProvincia = Trim(vgRegistro!gls_provincia)
         If Not IsNull(vgRegistro!gls_region) Then vlRptRegion = Trim(vgRegistro!gls_region)
         vlRptGlsDirBenNue = vlRptGlsDirBenNue & " - " & vlRptComuna & "/" & vlRptProvincia & "/" & vlRptRegion
    End If
    vgRegistro.Close

'RRR

   vlRptComuna = ""
   vlRptProvincia = ""
   vlRptRegion = ""
   
    vgSql = ""
    vgSql = "SELECT c.gls_comuna,p.gls_provincia,r.gls_region "
    vgSql = vgSql & "FROM MA_TPAR_COMUNA c,MA_TPAR_PROVINCIA p,MA_TPAR_REGION r "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "c.cod_direccion = '" & vlRptCodDireccion & "' AND "
    vgSql = vgSql & "c.cod_provincia = p.cod_provincia AND "
    vgSql = vgSql & "p.cod_region = r.cod_region AND "
    vgSql = vgSql & "c.cod_region = r.cod_region "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
        If Not IsNull(vgRegistro!gls_comuna) Then vlRptComuna = Trim(vgRegistro!gls_comuna)
        If Not IsNull(vgRegistro!gls_provincia) Then vlRptProvincia = Trim(vgRegistro!gls_provincia)
        If Not IsNull(vgRegistro!gls_region) Then vlRptRegion = Trim(vgRegistro!gls_region)
    End If
    vgRegistro.Close
    
    vlRptNomInter = "-"
    vlRptRutInter = "-"
    vlRptComInter = "-"
   
    Dim vlRptFechaVigEndoso As String
      
   vlRptGlsCauEndoso = ""
   vlRptGlsFactorEndoso = ""
   vlRptMtoRtaMod = 0
   vlRptMtoPension = 0
      
   vgSql = ""
   vgSql = "SELECT mto_pensionori,mto_pensioncal,fec_efecto, "
   vgSql = vgSql & "cod_cauendoso,cod_tipendoso "
   vgSql = vgSql & "FROM PP_TMAE_ENDOSO "
   vgSql = vgSql & "WHERE "
   vgSql = vgSql & "num_poliza = '" & Trim(Txt_PenPoliza) & "' AND "
  
   vgSql = vgSql & "num_endoso = " & vlRptNumEndosoEnd - 1 & " "
 
   Set vgRegistro = vgConexionBD.Execute(vgSql)
   If Not vgRegistro.EOF Then
      'vlRptNumEndoso = (vgRegistro!num_endoso)
      '****mvg agrego la ultima else
      If vgRegistro!cod_cauendoso = "14" Then
            vlRptGlsCauEndoso = Trim(fgBuscarGlosaElemento(vgCodTabla_CauEnd, Trim(vgRegistro!cod_cauendoso))) & vbCrLf & "Nombres Actualizados: " & vlRptNomBeNue
      ElseIf vgRegistro!cod_cauendoso = "15" Then
            vlRptGlsCauEndoso = Trim(fgBuscarGlosaElemento(vgCodTabla_CauEnd, Trim(vgRegistro!cod_cauendoso))) & vbCrLf & "La nueva direccion es: " & vlRptGlsDirBenNue
      ElseIf vgRegistro!cod_cauendoso = "16" Then
            vlRptGlsCauEndoso = Trim(fgBuscarGlosaElemento(vgCodTabla_CauEnd, Trim(vgRegistro!cod_cauendoso))) & vbCrLf & "El nuevo Nùmero Telefonico es: " & vlRptFonoBenNue
      ElseIf vgRegistro!cod_cauendoso = "27" Then
            vlRptGlsCauEndoso = Trim(fgBuscarGlosaElemento(vgCodTabla_CauEnd, Trim(vgRegistro!cod_cauendoso))) & vbCrLf & "Nombre    : " & vlRptNomBeNue
            vlRptGlsCauEndoso = vlRptGlsCauEndoso & vbCrLf & "Nro. documento : " & vlRptDocNum & vbCrLf & "E-mail       : " & vlRptCorreo
            vlRptGlsCauEndoso = vlRptGlsCauEndoso & vbCrLf & "Direccion : " & vlRptGlsDirBenNue & vbCrLf & "Telèfono   : " & vlRptFonoBenNue
      Else
            vlRptGlsCauEndoso = Trim(fgBuscarGlosaElemento(vgCodTabla_CauEnd, Trim(vgRegistro!cod_cauendoso)))
      End If
            
      vlRptFechaVigEndoso = DateSerial(Mid((vgRegistro!FEC_EFECTO), 1, 4), Mid((vgRegistro!FEC_EFECTO), 5, 2), Mid((vgRegistro!FEC_EFECTO), 7, 2))
      vlRptMtoPension = Format((vgRegistro!mto_pensioncal), "###,###,##0.00")
      vlRptMtoPensionOri = Format((vgRegistro!mto_pensionori), "###,###,##0.00")
      If vlRptMtoPensionOri <> vlRptMtoPension Then
         If vlRptMtoPensionOri > vlRptMtoPension Then
            vlRptMtoRtaMod = Format((vlRptMtoPensionOri - vlRptMtoPension), "#0.00")
            vlRptGlsFactorEndoso = Trim(clRptDisminuye)
         End If
         If vlRptMtoPensionOri < vlRptMtoPension Then
            vlRptMtoRtaMod = Format((vlRptMtoPension - vlRptMtoPensionOri), "#0.00")
            vlRptGlsFactorEndoso = Trim(clRptAumenta)
         End If
      Else
          
          vlRptMtoRtaMod = vlRptMtoPensionOri
          'CMV 20050928 F
          vlRptGlsFactorEndoso = Trim(clRptMantiene)
      End If
      
      
   End If
      
    vgSql = ""
    vgSql = " select a.num_poliza, a.num_endoso, b.num_idenben, a.fec_vigencia, a.mto_prima, a.mto_pension, a.num_mesdif, a.num_mesgar,"
    vgSql = vgSql & " Gls_NomBen , Gls_NomSegBen, b.Gls_PatBen, b.Gls_MatBen, Cod_Par, Cod_Sexo, Cod_SitInv, b.Mto_Pension as Mto_ben, c.gls_elemento, d.gls_tipoiden, b.fec_nacben,e.cod_tipendoso, b.gls_dirben"
    ',f.gls_elemento AS tipoendoso
    vgSql = vgSql & " from pp_tmae_poliza a"
    vgSql = vgSql & " join pp_tmae_ben b on a.num_poliza=b.num_poliza and a.num_endoso=b.num_endoso"
    vgSql = vgSql & " join ma_tpar_tabcod c on b.cod_par=c.cod_elemento and c.cod_tabla='PA'"
    vgSql = vgSql & " join ma_tpar_tipoiden d on b.cod_tipoidenben=d.cod_tipoiden"
    vgSql = vgSql & " join pp_tmae_endoso e on a.num_poliza=e.num_poliza and a.num_endoso=(e.num_endoso + 1) "
    'vgSql = vgSql & " JOIN ma_tpar_tabcod F on e.cod_cauendoso=f.cod_elemento and f.cod_tabla='CE'"
    vgSql = vgSql & " Where a.num_poliza = '" & vlGlobalNumPoliza & "' And a.num_endoso = " & (vlRptNumEndosoEnd) & ""
    vgSql = vgSql & " order by b.num_orden"
      
    Set rsLiq = New ADODB.Recordset
    rsLiq.CursorLocation = adUseClient
    rsLiq.Open vgSql, vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
    vlRutCliente = flRutCliente 'vgRutCliente + " - " + vgDgvCliente
    
    LNGa = CreateFieldDefFile(rsLiq, Replace(UCase(strRpt & "Estructura\PP_Rpt_EndDefEndoso.rpt"), ".RPT", ".TTX"), 1)
    
    'vlRptFechaVigEndoso = "01/01/2012" ' IIf(IsNull(vlRptFechaVigEndoso), " ", vlRptFechaVigEndoso)

    If objRep.CargaReporte(strRpt & "", "PP_Rpt_EndDefEndoso_2.rpt", "Informe de Liquidación de Rentas Vitalicias", rsLiq, True, _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("NombreSistema", vgNombreSistema), _
                            ArrFormulas("NombreSubSistema", vgNombreSubSistema), _
                            ArrFormulas("TipPension", vlRptCodTipPension), _
                            ArrFormulas("Afp", vlRptCodAfp), _
                            ArrFormulas("TipRta", vlRptCodTipRen), _
                            ArrFormulas("NombreCausante", vlRptNomBen), _
                            ArrFormulas("RutCausante", vlRptRutBen), _
                            ArrFormulas("Direccion", vlRptGlsDirBen), _
                            ArrFormulas("Fono", vlRptFonoBen), _
                            ArrFormulas("Comuna", vlRptComuna), _
                            ArrFormulas("Provincia", vlRptProvincia), _
                            ArrFormulas("Region", vlRptRegion), _
                            ArrFormulas("Origen", clRptOrigen), _
                            ArrFormulas("CodMoneda", cgCodTipMonedaUF), _
                            ArrFormulas("MotivoEndoso", vlRptGlsCauEndoso), _
                            ArrFormulas("GlsFactorEndoso", vlRptGlsFactorEndoso), _
                            ArrFormulas("MtoRtaMod", str(vlRptMtoRtaMod)), _
                            ArrFormulas("MtoPension", str(vlRptMtoPension)), _
                            ArrFormulas("FechaVigEndoso", vlRptFechaVigEndoso), _
                            ArrFormulas("MtoRtaOri", str(vlRptMtoPensionOri)), _
                            ArrFormulas("CodMonedaCor", vlImpCodMoneda)) = False Then

        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Function
    End If
      
   Screen.MousePointer = 0

Exit Function
Err_flImprimirEndosoDef:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
    
End Function
    
Function fl_imprimir_poliza()
On Error GoTo Err_flImprimirPoliza

'Imprimir Reporte de Poliza Original
   Screen.MousePointer = 11

   vlArchivo = strRpt & "PP_Rpt_EndPoliza.rpt"   '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
      MsgBox "Archivo de Reporte de Endoso Póliza de Renta Vitalicia no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
      Screen.MousePointer = 0
      Exit Function
   End If
   
   vlRptCodTipPension = ""
   'vlRptNumEndoso = 0
   vlRptCodAfp = ""
   vlRptCodTipRen = ""
   
   vgSql = ""
   vgSql = "SELECT num_endoso,cod_afp,cod_tipren,cod_tippension "
   vgSql = vgSql & ",cod_moneda,mto_valmoneda "
   vgSql = vgSql & "FROM PP_TMAE_POLIZA "
   vgSql = vgSql & "WHERE "
   vgSql = vgSql & "num_poliza = '" & vlGlobalNumPoliza & "' AND "
   vgSql = vgSql & "num_endoso = " & vlRptNumEndosoPol & " "
   'vgSql = vgSql & "ORDER BY num_endoso DESC "
   Set vgRegistro = vgConexionBD.Execute(vgSql)
   If Not vgRegistro.EOF Then
      'vlRptNumEndoso = (vgRegistro!num_endoso)
      vlRptCodAfp = Trim(vgRegistro!cod_afp) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_AFP, Trim(vgRegistro!cod_afp)))
      vlRptCodTipRen = Trim(vgRegistro!Cod_TipRen) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_TipRen, Trim(vgRegistro!Cod_TipRen)))
      vlRptCodTipPension = Trim(vgRegistro!Cod_TipPension) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_TipPen, Trim(vgRegistro!Cod_TipPension)))
      vlImpCodMoneda = vgRegistro!Cod_Moneda
      vlImpNomMoneda = fgBuscarGlosaElemento(vgCodTabla_TipMon, vlImpCodMoneda)
   End If
   vgRegistro.Close
   
   vlRptCodInsSalud = ""
   vlRptNomBen = ""
   vlRptRutBen = ""
   vlRptGlsDirBen = ""
   vlRptFonoBen = ""
   vlRptCodDireccion = "0"
   
   vgSql = ""
   vgSql = "SELECT cod_inssalud,gls_nomben,gls_nomsegben,gls_patben,gls_matben,Cod_TipoIdenben, "
   vgSql = vgSql & "Num_Idenben,cod_direccion,gls_dirben,gls_fonoben "
   vgSql = vgSql & "FROM PP_TMAE_BEN "
   vgSql = vgSql & "WHERE "
   vgSql = vgSql & "num_poliza = '" & vlGlobalNumPoliza & "' AND "
   vgSql = vgSql & "num_endoso = " & (vlRptNumEndosoPol) & " AND "
   vgSql = vgSql & "cod_par = '" & Trim(clCodParCau) & "' "
   Set vgRegistro = vgConexionBD.Execute(vgSql)
   If Not vgRegistro.EOF Then
        If Not IsNull(vgRegistro!Cod_InsSalud) Then
            vlRptCodInsSalud = " " & Trim(vgRegistro!Cod_InsSalud) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_InsSal, Trim(vgRegistro!Cod_InsSalud)))
        End If
        If Not IsNull(vgRegistro!Gls_NomBen) And Not IsNull(vgRegistro!Gls_PatBen) Then
            vlNombre = vgRegistro!Gls_NomBen
            vlNombreSeg = IIf(IsNull(vgRegistro!Gls_NomSegBen), "", vgRegistro!Gls_NomSegBen)
            vlPaterno = vgRegistro!Gls_PatBen
            'I - MC 24/01/2008
            ''vlMaterno = IIf(IsNull(vgRegistro!Gls_PatBen), "", vgRegistro!Gls_PatBen)
            vlMaterno = IIf(IsNull(vgRegistro!Gls_MatBen), "", vgRegistro!Gls_MatBen)
            'F - MC 24/01/2008
            vlNombreCompleto = fgFormarNombreCompleto(vlNombre, vlNombreSeg, vlPaterno, vlMaterno)
            
            vlRptNomBen = vlNombreCompleto
        End If
        If Not IsNull(vgRegistro!Cod_TipoIdenBen) And Not IsNull(vgRegistro!Num_IdenBen) Then
            vlNombreTipoIden = fgBuscarNombreTipoIden(vgRegistro!Cod_TipoIdenBen)
            vlRptRutBen = vlNombreTipoIden & " - " & (Trim(vgRegistro!Num_IdenBen))
        End If
        If Not IsNull(vgRegistro!Gls_DirBen) Then
            vlRptGlsDirBen = Trim(vgRegistro!Gls_DirBen)
        End If
        If Not IsNull(vgRegistro!Gls_FonoBen) Then
            vlRptFonoBen = Trim(vgRegistro!Gls_FonoBen)
        Else
            vlRptFonoBen = " "
        End If
        If Not IsNull(vgRegistro!Cod_Direccion) Then
            vlRptCodDireccion = Trim(vgRegistro!Cod_Direccion)
        End If
   End If
   vgRegistro.Close

   vlRptComuna = ""
   vlRptProvincia = ""
   vlRptRegion = ""
   
    vgSql = ""
    vgSql = "SELECT c.gls_comuna,p.gls_provincia,r.gls_region "
    vgSql = vgSql & "FROM MA_TPAR_COMUNA c,MA_TPAR_PROVINCIA p,MA_TPAR_REGION r "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "c.cod_direccion = '" & vlRptCodDireccion & "' AND "
    vgSql = vgSql & "c.cod_provincia = p.cod_provincia AND "
    vgSql = vgSql & "p.cod_region = r.cod_region AND "
    vgSql = vgSql & "c.cod_region = r.cod_region "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
        If Not IsNull(vgRegistro!gls_comuna) Then vlRptComuna = Trim(vgRegistro!gls_comuna)
        If Not IsNull(vgRegistro!gls_provincia) Then vlRptProvincia = Trim(vgRegistro!gls_provincia)
        If Not IsNull(vgRegistro!gls_region) Then vlRptRegion = Trim(vgRegistro!gls_region)
    End If
    vgRegistro.Close
    
    vlRptNomInter = "-"
    vlRptRutInter = "-"
    vlRptComInter = "-"
      
    vgQuery = ""
    vgQuery = vgQuery & "{PP_TMAE_POLIZA.num_poliza} = '" & vlGlobalNumPoliza & "' AND "
    vgQuery = vgQuery & "{PP_TMAE_POLIZA.num_endoso} = " & (vlRptNumEndosoPol) & " "
           
   Rpt_Reporte.Reset
   Rpt_Reporte.WindowState = crptMaximized
   Rpt_Reporte.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
   Rpt_Reporte.Connect = vgRutaDataBase       ' o App.Path & "\Nestle.mdb"
   Rpt_Reporte.SelectionFormula = vgQuery

   Rpt_Reporte.Formulas(0) = ""
   Rpt_Reporte.Formulas(1) = ""
   Rpt_Reporte.Formulas(2) = ""
   Rpt_Reporte.Formulas(3) = ""
   Rpt_Reporte.Formulas(4) = ""
   Rpt_Reporte.Formulas(5) = ""
   Rpt_Reporte.Formulas(6) = ""
   Rpt_Reporte.Formulas(7) = ""
   Rpt_Reporte.Formulas(8) = ""
   Rpt_Reporte.Formulas(9) = ""
   Rpt_Reporte.Formulas(10) = ""
   Rpt_Reporte.Formulas(11) = ""
   Rpt_Reporte.Formulas(12) = ""
   Rpt_Reporte.Formulas(13) = ""
   Rpt_Reporte.Formulas(14) = ""
   Rpt_Reporte.Formulas(15) = ""
   Rpt_Reporte.Formulas(16) = ""
   Rpt_Reporte.Formulas(17) = ""
   Rpt_Reporte.Formulas(18) = ""
   Rpt_Reporte.Formulas(19) = ""
   Rpt_Reporte.Formulas(20) = ""
   Rpt_Reporte.Formulas(21) = ""
   Rpt_Reporte.Formulas(22) = ""
   Rpt_Reporte.Formulas(23) = ""
         
   Rpt_Reporte.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
   Rpt_Reporte.Formulas(1) = "NombreSistema= '" & vgNombreSistema & "'"
   Rpt_Reporte.Formulas(2) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
   Rpt_Reporte.Formulas(3) = "TipPension = '" & vlRptCodTipPension & "'"
   Rpt_Reporte.Formulas(4) = "Afp = '" & vlRptCodAfp & "'"
   Rpt_Reporte.Formulas(5) = "TipRta = '" & vlRptCodTipRen & "'"
   Rpt_Reporte.Formulas(6) = "InstSalud = '" & vlRptCodInsSalud & "'"
   Rpt_Reporte.Formulas(7) = "NombreCausante = '" & vlRptNomBen & "'"
   Rpt_Reporte.Formulas(8) = "RutCausante = '" & vlRptRutBen & "'"
   Rpt_Reporte.Formulas(9) = "Direccion = '" & vlRptGlsDirBen & "'"
   Rpt_Reporte.Formulas(10) = "Fono = '" & vlRptFonoBen & "'"
   Rpt_Reporte.Formulas(11) = "Comuna = '" & vlRptComuna & "'"
   Rpt_Reporte.Formulas(12) = "Provincia = '" & vlRptProvincia & "'"
   Rpt_Reporte.Formulas(13) = "Region = '" & vlRptRegion & "'"
   Rpt_Reporte.Formulas(14) = "NombreInter = '" & vlRptNomInter & "'"
   Rpt_Reporte.Formulas(15) = "RutInter = '" & vlRptRutInter & "'"
   Rpt_Reporte.Formulas(16) = "ComisionInter = '" & vlRptComInter & "'"
   Rpt_Reporte.Formulas(17) = "Origen = '" & clRptOrigen & "'"
   Rpt_Reporte.Formulas(18) = "CodMoneda = '" & vlImpNomMoneda & "'"
   Rpt_Reporte.Formulas(19) = "CodMonedaCor = '" & vlImpCodMoneda & "'"
   Rpt_Reporte.Formulas(20) = "MtoPension = " & str(Lbl_EndMtoPensionOrig) & " "

   Rpt_Reporte.SubreportToChange = ""
   Rpt_Reporte.Destination = crptToWindow
   Rpt_Reporte.WindowState = crptMaximized
   Rpt_Reporte.WindowTitle = "Póliza de Endoso Original"
   'Rpt_Reporte.SelectionFormula = ""
   Rpt_Reporte.Action = 1
   Screen.MousePointer = 0

Exit Function
Err_flImprimirPoliza:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
    
End Function

Function flImprimirEndosoPrev()
On Error GoTo Err_flImprimirEndosoPrev

'Imprimir Reporte de Poliza Original
   Screen.MousePointer = 11

   vlArchivo = strRpt & "PP_Rpt_EndPreEndoso.rpt"   '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
      MsgBox "Archivo de Reporte de Endoso de Renta Vitalicia Preliminar no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
      Screen.MousePointer = 0
      Exit Function
   End If
    
    vlcobertura = ""
    
    'Buscar Antecedentes para formar la Cobertura
    vgSql = "SELECT  p.num_poliza,p.cod_tippension,"
    vgSql = vgSql & "p.cod_tipren,p.num_mesdif,p.cod_modalidad,p.num_mesgar,"
    vgSql = vgSql & "p.mto_pension,t.gls_elemento as gls_pension,"
    vgSql = vgSql & "r.gls_elemento as gls_renta,m.gls_elemento as gls_modalidad,"
    vgSql = vgSql & "p.cod_moneda, "
    vgSql = vgSql & "p.cod_cobercon, b.gls_cobercon,p.cod_dercre, p.cod_dergra "
    vgSql = vgSql & "FROM "
    vgSql = vgSql & "pp_tmae_endpoliza p, ma_tpar_tabcod t, ma_tpar_tabcod r, "
    vgSql = vgSql & "ma_tpar_tabcod m, ma_tpar_cobercon b "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "p.num_poliza = '" & vlGlobalNumPoliza & "' AND "
    vgSql = vgSql & "t.cod_tabla = '" & Trim(vgCodTabla_TipPen) & "' AND "
    vgSql = vgSql & "t.cod_elemento = p.cod_tippension AND "
    vgSql = vgSql & "r.cod_tabla = '" & Trim(vgCodTabla_TipRen) & "' AND "
    vgSql = vgSql & "r.cod_elemento = p.cod_tipren AND "
    vgSql = vgSql & "m.cod_tabla = '" & Trim(vgCodTabla_AltPen) & "' AND "
    vgSql = vgSql & "m.cod_elemento = p.cod_modalidad AND "
    vgSql = vgSql & "p.cod_cobercon = b.cod_cobercon"
    Set vlRegistro = vgConexionBD.Execute(vgSql)
    If Not (vlRegistro.EOF) Then
        vlcobertura = vlRegistro!Gls_Renta
        If vlRegistro!Cod_Modalidad = 1 Then
            If Not IsNull(vlRegistro!gls_modalidad) Then
                vlcobertura = vlcobertura & " " & vlRegistro!gls_modalidad
            End If
        Else
            If Not IsNull(vlRegistro!gls_modalidad) Then
                vlcobertura = vlcobertura & " CON P. " & vlRegistro!gls_modalidad
            End If
        End If
        If vlRegistro!Cod_CoberCon <> 0 Then
            If Not IsNull(vlRegistro!GLS_COBERCON) Then
                vlcobertura = vlcobertura & " CON " & vlRegistro!GLS_COBERCON
            End If
        End If
        If vlRegistro!Cod_DerCre = "S" Then
            vlcobertura = vlcobertura & " CON D.CRECER"
        End If
        
        If vlRegistro!Cod_DerGra = "S" Then
            vlcobertura = vlcobertura & " Y CON GRATIFICACIÓN"
        End If
    End If
    vlRegistro.Close
    

   vlRptGlsCauEndoso = ""
   vlRptGlsFactorEndoso = ""
   vlRptMtoRtaMod = 0
   vlRptMtoPension = 0
      
   vgSql = ""
   vgSql = "SELECT mto_pensionori,mto_pensioncal,fec_efecto, "
   vgSql = vgSql & "cod_cauendoso,cod_tipendoso "
   vgSql = vgSql & "FROM PP_TMAE_ENDOSO "
   vgSql = vgSql & "WHERE "
   vgSql = vgSql & "num_poliza = '" & vlGlobalNumPoliza & "' AND "
   vgSql = vgSql & "num_endoso = " & vlRptNumEndosoEnd & " "
   Set vgRegistro = vgConexionBD.Execute(vgSql)
   If Not vgRegistro.EOF Then
      'vlRptNumEndoso = (vgRegistro!num_endoso)
      vlRptGlsCauEndoso = Trim(fgBuscarGlosaElemento(vgCodTabla_CauEnd, Trim(vgRegistro!cod_cauendoso)))
      vlRptFechaVigEndoso = DateSerial(Mid((vgRegistro!FEC_EFECTO), 1, 4), Mid((vgRegistro!FEC_EFECTO), 5, 2), Mid((vgRegistro!FEC_EFECTO), 7, 2))
      vlRptMtoPension = Format((vgRegistro!mto_pensioncal), "#,#0.00")
      vlRptMtoPensionOri = Format((vgRegistro!mto_pensionori), "#,#0.00")
      If vlRptMtoPensionOri <> vlRptMtoPension Then
         If vlRptMtoPensionOri > vlRptMtoPension Then
            vlRptMtoRtaMod = Format((vlRptMtoPensionOri - vlRptMtoPension), "#0.00")
            vlRptGlsFactorEndoso = Trim(clRptDisminuye)
         End If
         If vlRptMtoPensionOri < vlRptMtoPension Then
            vlRptMtoRtaMod = Format((vlRptMtoPension - vlRptMtoPensionOri), "#0.00")
            vlRptGlsFactorEndoso = Trim(clRptAumenta)
         End If
      Else
         
          vlRptMtoRtaMod = vlRptMtoPensionOri
          'CMV 20050928 F
          vlRptGlsFactorEndoso = Trim(clRptMantiene)
      End If
   End If
   vgRegistro.Close
      
   vgQuery = ""
   vgQuery = vgQuery & "{PP_TMAE_ENDPOLIZA.num_poliza} = '" & vlGlobalNumPoliza & "' AND "
   vgQuery = vgQuery & "{PP_TMAE_ENDPOLIZA.num_endoso} = " & (vlRptNumEndosoEnd) & " "
           
   Rpt_Reporte.Reset
   Rpt_Reporte.WindowState = crptMaximized
   Rpt_Reporte.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
   Rpt_Reporte.Connect = vgRutaDataBase       ' o App.Path & "\Nestle.mdb"
   Rpt_Reporte.SelectionFormula = vgQuery

   Rpt_Reporte.Formulas(0) = ""
   Rpt_Reporte.Formulas(1) = ""
   Rpt_Reporte.Formulas(2) = ""
   Rpt_Reporte.Formulas(3) = ""
   Rpt_Reporte.Formulas(4) = ""
   Rpt_Reporte.Formulas(5) = ""
   Rpt_Reporte.Formulas(6) = ""
   Rpt_Reporte.Formulas(7) = ""
   Rpt_Reporte.Formulas(8) = ""
   Rpt_Reporte.Formulas(9) = ""
   Rpt_Reporte.Formulas(10) = ""
   Rpt_Reporte.Formulas(11) = ""
   Rpt_Reporte.Formulas(12) = ""
   Rpt_Reporte.Formulas(13) = ""
   Rpt_Reporte.Formulas(14) = ""
   Rpt_Reporte.Formulas(15) = ""
   Rpt_Reporte.Formulas(16) = ""
   Rpt_Reporte.Formulas(17) = ""
   Rpt_Reporte.Formulas(18) = ""
   Rpt_Reporte.Formulas(19) = ""
   Rpt_Reporte.Formulas(20) = ""
   Rpt_Reporte.Formulas(21) = ""
   Rpt_Reporte.Formulas(22) = ""
   Rpt_Reporte.Formulas(23) = ""
   Rpt_Reporte.Formulas(24) = ""
         
   Rpt_Reporte.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
   Rpt_Reporte.Formulas(1) = "NombreSistema= '" & vgNombreSistema & "'"
   Rpt_Reporte.Formulas(2) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
   Rpt_Reporte.Formulas(3) = "TipPension = '" & vlRptCodTipPension & "'"
   Rpt_Reporte.Formulas(4) = "Afp = '" & vlRptCodAfp & "'"
'   Rpt_Reporte.Formulas(5) = "TipRta = '" & vlRptCodTipRen & "'"
   Rpt_Reporte.Formulas(5) = "Concatenar = '" & vlcobertura & "'"
   Rpt_Reporte.Formulas(6) = "InstSalud = '" & vlRptCodInsSalud & "'"
   Rpt_Reporte.Formulas(7) = "NombreCausante = '" & vlRptNomBen & "'"
   Rpt_Reporte.Formulas(8) = "RutCausante = '" & vlRptRutBen & "'"
   Rpt_Reporte.Formulas(9) = "Direccion = '" & vlRptGlsDirBen & "'"
   Rpt_Reporte.Formulas(10) = "Fono = '" & vlRptFonoBen & "'"
   Rpt_Reporte.Formulas(11) = "Comuna = '" & vlRptComuna & "'"
   Rpt_Reporte.Formulas(12) = "Provincia = '" & vlRptProvincia & "'"
   Rpt_Reporte.Formulas(13) = "Region = '" & vlRptRegion & "'"
   Rpt_Reporte.Formulas(14) = "NombreInter = '" & vlRptNomInter & "'"
   Rpt_Reporte.Formulas(15) = "RutInter = '" & vlRptRutInter & "'"
   Rpt_Reporte.Formulas(16) = "ComisionInter = '" & vlRptComInter & "'"

   Rpt_Reporte.Formulas(19) = "MotivoEndoso = '" & vlRptGlsCauEndoso & "'"
   Rpt_Reporte.Formulas(20) = "GlsFactorEndoso = '" & vlRptGlsFactorEndoso & "'"
   Rpt_Reporte.Formulas(21) = "MtoRtaMod = '" & (Format(vlRptMtoRtaMod, "###,###,##0.00")) & "'"
   Rpt_Reporte.Formulas(22) = "MtoPension = '" & (Format(vlRptMtoPension, "###,###,##0.00")) & "'"
   Rpt_Reporte.Formulas(23) = "FechaVigEndoso = '" & (vlRptFechaVigEndoso) & "'"
   Rpt_Reporte.Formulas(24) = "MtoRtaOri = '" & (Format(vlRptMtoPensionOri, "###,###,##0.00")) & "'"
   Rpt_Reporte.Formulas(18) = "CodMoneda = '" & vlImpNomMoneda & "'"
   'Rpt_Reporte.Formulas(25) = "CodMonedaCor = '" & vlImpCodMoneda & "'"

   Rpt_Reporte.SubreportToChange = ""
   Rpt_Reporte.Destination = crptToWindow
   Rpt_Reporte.WindowState = crptMaximized
   Rpt_Reporte.WindowTitle = "Póliza de Endoso Preliminar"
   'Rpt_Reporte.SelectionFormula = ""
   Rpt_Reporte.Action = 1
   Screen.MousePointer = 0

Exit Function
Err_flImprimirEndosoPrev:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
    
End Function

Private Sub txt_dirben_Change()
    Call UcaseText(txt_dirben)
End Sub

Private Sub Txt_PenPoliza_KeyPress(KeyAscii As Integer)
On Error GoTo Err_TxtPenPolizaKeyPress

    If KeyAscii = 13 Then
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
Sub UcaseText(txt As TextBox)
txt.Text = UCase(txt.Text)
txt.SelStart = Len(txt)
End Sub

Private Sub txtApematBen_Change()
    Call UcaseText(txtApematBen)
End Sub

Private Sub txtApepatBen_Change()
    Call UcaseText(txtApepatBen)
End Sub

Private Sub txtmto_pensioncal_LostFocus()
    txtmto_pensioncal = sacapuntos(txtmto_pensioncal.Text)
End Sub

Private Sub txtNomben_Change()
    Call UcaseText(txtNomben)
End Sub

Private Sub txtNomsegBen_Change()
    Call UcaseText(txtNomsegBen)
End Sub

Public Function sacapuntos(cadena As String) As String
Dim Comilla As String
 
    puntos = ""
    While InStr(cadena, ",") > 0
        cadena = Left(cadena, InStr(cadena, ",") - 1) & Comilla & Right(cadena, Len(cadena) - InStr(cadena, ","))
    Wend
    sacapuntos = cadena
End Function


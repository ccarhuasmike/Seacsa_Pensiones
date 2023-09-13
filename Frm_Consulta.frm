VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Frm_Consulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta General"
   ClientHeight    =   8505
   ClientLeft      =   675
   ClientTop       =   1335
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   10110
   Begin VB.Frame Fra_PenPoliza 
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
      TabIndex        =   110
      Top             =   0
      Width           =   9855
      Begin VB.TextBox txt_End 
         Height          =   285
         Left            =   8200
         MaxLength       =   10
         TabIndex        =   155
         Top             =   240
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.TextBox Txt_PenNumIdent 
         Height          =   285
         Left            =   5400
         MaxLength       =   16
         TabIndex        =   124
         Top             =   240
         Width           =   1875
      End
      Begin VB.ComboBox Cmb_PenNumIdent 
         BackColor       =   &H00E0FFFF&
         Height          =   315
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   123
         Top             =   240
         Width           =   2235
      End
      Begin VB.CommandButton Cmd_Buscar 
         BackColor       =   &H00000000&
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
         Left            =   9120
         Picture         =   "Frm_Consulta.frx":0000
         TabIndex        =   3
         ToolTipText     =   "Buscar"
         Top             =   520
         Width           =   615
      End
      Begin VB.TextBox Txt_PenPoliza 
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   1185
      End
      Begin VB.CommandButton Cmd_BuscarPol 
         Height          =   375
         Left            =   9120
         Picture         =   "Frm_Consulta.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Buscar Póliza"
         Top             =   150
         Width           =   615
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Ident."
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   125
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Lbl_PenNombre 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   116
         Top             =   585
         Width           =   7695
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Póliza"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   115
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   114
         Top             =   580
         Width           =   855
      End
      Begin VB.Label Lbl_End 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   8265
         TabIndex        =   113
         Top             =   240
         Width           =   525
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nº Endoso"
         Height          =   240
         Index           =   0
         Left            =   7320
         TabIndex        =   112
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label8 
         Caption         =   "  Póliza / Pensionado"
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
         TabIndex        =   111
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.Frame Fra_Botones 
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   7440
      Width           =   9855
      Begin VB.CommandButton cmdPrintSel 
         Height          =   675
         Left            =   9000
         Picture         =   "Frm_Consulta.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   151
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   4800
         Picture         =   "Frm_Consulta.frx":55E6
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Cancelar2 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   3840
         Picture         =   "Frm_Consulta.frx":56E0
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   200
         Width           =   720
      End
      Begin Crystal.CrystalReport Rpt_General 
         Left            =   6480
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label Label1 
         Caption         =   "Imprime Polizas por Archivo"
         Height          =   375
         Left            =   7800
         TabIndex        =   152
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Fra_Grupo 
      Caption         =   "Grupo familiar de la póliza"
      Height          =   1380
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   9900
      Begin MSFlexGridLib.MSFlexGrid Msf_GrillaGrupo 
         Height          =   1110
         Left            =   60
         TabIndex        =   6
         Top             =   195
         Width           =   9705
         _ExtentX        =   17119
         _ExtentY        =   1958
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         BackColor       =   14745599
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
   Begin TabDlg.SSTab SSTab_Consulta 
      Height          =   5055
      Left            =   105
      TabIndex        =   7
      Top             =   2385
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   8916
      _Version        =   393216
      Tabs            =   5
      Tab             =   3
      TabsPerRow      =   5
      TabHeight       =   520
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Póliza"
      TabPicture(0)   =   "Frm_Consulta.frx":5CBA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Fra_Poliza"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Grupo Familiar"
      TabPicture(1)   =   "Frm_Consulta.frx":5CD6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Fra_GrupFam"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Pago de Pensiones"
      TabPicture(2)   =   "Frm_Consulta.frx":5CF2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Fra_RetJud"
      Tab(2).Control(1)=   "Fra_Receptor"
      Tab(2).Control(2)=   "Fra_PlanSalud"
      Tab(2).Control(3)=   "Fra_FormaPago"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Liquidaciones (H/D)"
      TabPicture(3)   =   "Frm_Consulta.frx":5D0E
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Msf_GrillaHabDes"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Fra_LiqSeleccion"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Fra_LiqFecha"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Emisión Certificado"
      TabPicture(4)   =   "Frm_Consulta.frx":5D2A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Cmd_Imprimir"
      Tab(4).Control(1)=   "Fra_Certificados"
      Tab(4).Control(2)=   "cmdNuevo"
      Tab(4).ControlCount=   3
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   675
         Left            =   -66705
         Picture         =   "Frm_Consulta.frx":5D46
         Style           =   1  'Graphical
         TabIndex        =   150
         Top             =   1860
         Width           =   720
      End
      Begin VB.Frame Fra_Certificados 
         Caption         =   "Emisión Certificados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2850
         Left            =   -73320
         TabIndex        =   143
         Top             =   1080
         Width           =   6270
         Begin VB.OptionButton Opt_CertSupervivencia 
            Caption         =   "Certificado de Supervivencia"
            Height          =   195
            Left            =   840
            TabIndex        =   149
            Top             =   1875
            Width           =   3015
         End
         Begin VB.OptionButton Opt_ConPensiones 
            Caption         =   "Constancia de Pensiones"
            Height          =   195
            Left            =   840
            TabIndex        =   147
            Top             =   1320
            Value           =   -1  'True
            Width           =   2970
         End
         Begin VB.OptionButton Opt_CerCarFam 
            Caption         =   "Certificado de Cargas Familiares"
            Height          =   195
            Left            =   240
            TabIndex        =   146
            Top             =   2160
            Visible         =   0   'False
            Width           =   2970
         End
         Begin VB.OptionButton Opt_CerDecRta 
            Caption         =   "Certificado de Declaración de Renta"
            Height          =   195
            Left            =   240
            TabIndex        =   145
            Top             =   2400
            Visible         =   0   'False
            Width           =   2970
         End
         Begin VB.OptionButton Opt_CerPensiones 
            Caption         =   "Certificado de Pensiones"
            Height          =   195
            Left            =   840
            TabIndex        =   144
            Top             =   720
            Width           =   2970
         End
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   -66705
         Picture         =   "Frm_Consulta.frx":6188
         Style           =   1  'Graphical
         TabIndex        =   142
         Top             =   2640
         Width           =   720
      End
      Begin VB.Frame Fra_LiqFecha 
         Caption         =   " Fecha de Pago de Pensiones "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1935
         Left            =   120
         TabIndex        =   133
         Top             =   480
         Width           =   3120
         Begin VB.TextBox Txt_LiqFecIni 
            Height          =   285
            Left            =   120
            MaxLength       =   10
            TabIndex        =   138
            Top             =   540
            Width           =   1005
         End
         Begin VB.TextBox Txt_LiqFecTer 
            Height          =   285
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   137
            Top             =   540
            Width           =   1005
         End
         Begin VB.CommandButton Cmd_LiqConsultaHD 
            Caption         =   "Generar Consulta"
            Height          =   375
            Left            =   120
            TabIndex        =   136
            Top             =   840
            Width           =   1890
         End
         Begin VB.CommandButton Cmd_LiqHDImprimir 
            Caption         =   "Imprimir Liquidación"
            Height          =   495
            Left            =   120
            TabIndex        =   135
            Top             =   1320
            Width           =   1290
         End
         Begin VB.CommandButton Cmd_LiqHDLimpiar 
            Caption         =   "Limpiar"
            Height          =   375
            Left            =   2040
            TabIndex        =   134
            Top             =   840
            Width           =   690
         End
         Begin Crystal.CrystalReport Rpt_Reporte 
            Left            =   2640
            Top             =   1560
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowState     =   2
            PrintFileLinesPerPage=   60
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Ingrese Fecha a Procesar"
            Height          =   255
            Index           =   72
            Left            =   240
            TabIndex        =   140
            Top             =   240
            Width           =   2475
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
            Index           =   71
            Left            =   1080
            TabIndex        =   139
            Top             =   510
            Width           =   255
         End
      End
      Begin VB.Frame Fra_LiqSeleccion 
         Caption         =   "Selección de Concepto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1935
         Left            =   3360
         TabIndex        =   131
         Top             =   480
         Width           =   6325
         Begin VB.ListBox Lst_LiqSeleccion 
            Height          =   1410
            ItemData        =   "Frm_Consulta.frx":6842
            Left            =   150
            List            =   "Frm_Consulta.frx":6844
            Style           =   1  'Checkbox
            TabIndex        =   132
            Top             =   285
            Width           =   6045
         End
      End
      Begin VB.Frame Fra_Poliza 
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
         Height          =   4320
         Left            =   -74760
         TabIndex        =   75
         Top             =   480
         Width           =   9375
         Begin VB.Label lblFecSolDev 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   7800
            TabIndex        =   159
            Top             =   2400
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Solicitud del dev."
            Height          =   255
            Left            =   6480
            TabIndex        =   158
            Top             =   2400
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Asesor"
            Height          =   255
            Left            =   3600
            TabIndex        =   157
            Top             =   3840
            Width           =   615
         End
         Begin VB.Label lblNomAsesor 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   4320
            TabIndex        =   156
            Top             =   3800
            Width           =   4575
         End
         Begin VB.Label Lbl_Moneda 
            Caption         =   "(TM)"
            Enabled         =   0   'False
            Height          =   300
            Index           =   0
            Left            =   8880
            TabIndex        =   148
            Top             =   3480
            Width           =   375
         End
         Begin VB.Label Lbl_PolFecDev 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   2400
            TabIndex        =   130
            Top             =   650
            Width           =   975
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Fecha Devengue"
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   129
            Top             =   650
            Width           =   1695
         End
         Begin VB.Label Lbl_PolEstado 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   7320
            TabIndex        =   78
            Top             =   300
            Width           =   1575
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Monto Prima"
            Height          =   255
            Index           =   33
            Left            =   225
            TabIndex        =   109
            Top             =   3800
            Width           =   1335
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Monto Pensión"
            Height          =   255
            Index           =   34
            Left            =   6480
            TabIndex        =   108
            Top             =   3450
            Width           =   1215
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Tasa Cto. Equivalente"
            Height          =   255
            Index           =   35
            Left            =   225
            TabIndex        =   107
            Top             =   3100
            Width           =   1815
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Tasa de Venta"
            Height          =   255
            Index           =   36
            Left            =   6495
            TabIndex        =   106
            Top             =   2745
            Width           =   1245
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Tasa Cto. Reaseguro"
            Height          =   255
            Index           =   37
            Left            =   210
            TabIndex        =   105
            Top             =   3450
            Width           =   1815
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Tasa de Int. Período Garantizado"
            Height          =   270
            Index           =   38
            Left            =   5205
            TabIndex        =   104
            Top             =   3105
            Width           =   2520
         End
         Begin VB.Label Lbl_PolMtoPri 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   2400
            TabIndex        =   103
            Top             =   3800
            Width           =   1095
         End
         Begin VB.Label Lbl_PolMtoPen 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   7800
            TabIndex        =   102
            Top             =   3450
            Width           =   1095
         End
         Begin VB.Label Lbl_PolTasaCto 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   2400
            TabIndex        =   101
            Top             =   3100
            Width           =   1095
         End
         Begin VB.Label Lbl_PolTasaVta 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   7800
            TabIndex        =   100
            Top             =   2745
            Width           =   1095
         End
         Begin VB.Label Lbl_PolTasaRea 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   2400
            TabIndex        =   99
            Top             =   3450
            Width           =   1095
         End
         Begin VB.Label Lbl_PolTasaPerGar 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   7800
            TabIndex        =   98
            Top             =   3105
            Width           =   1095
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Número de Endoso"
            Height          =   255
            Index           =   22
            Left            =   195
            TabIndex        =   97
            Top             =   1000
            Width           =   2055
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "AFP"
            Height          =   255
            Index           =   23
            Left            =   240
            TabIndex        =   96
            Top             =   2750
            Width           =   975
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Modalidad"
            Height          =   255
            Index           =   28
            Left            =   225
            TabIndex        =   95
            Top             =   2050
            Width           =   900
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Tipo de Pensión "
            Height          =   255
            Index           =   24
            Left            =   180
            TabIndex        =   94
            Top             =   1350
            Width           =   1635
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Estado"
            Height          =   255
            Index           =   25
            Left            =   6480
            TabIndex        =   93
            Top             =   300
            Width           =   855
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Nº de Beneficiarios de Póliza"
            Height          =   255
            Index           =   30
            Left            =   225
            TabIndex        =   92
            Top             =   2400
            Width           =   2175
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Periodo de Vigencia"
            Height          =   255
            Index           =   31
            Left            =   180
            TabIndex        =   91
            Top             =   300
            Width           =   1695
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Meses Diferidos"
            Height          =   255
            Index           =   27
            Left            =   6480
            TabIndex        =   90
            Top             =   1695
            Width           =   1455
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Meses Garantizados"
            Height          =   255
            Index           =   29
            Left            =   6480
            TabIndex        =   89
            Top             =   2055
            Width           =   1440
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Tipo de Renta"
            Height          =   255
            Index           =   26
            Left            =   195
            TabIndex        =   88
            Top             =   1700
            Width           =   1215
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
            Index           =   32
            Left            =   3480
            TabIndex        =   87
            Top             =   300
            Width           =   255
         End
         Begin VB.Label Lbl_PolNumEndoso 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   2400
            TabIndex        =   86
            Top             =   1000
            Width           =   735
         End
         Begin VB.Label Lbl_PolMesDif 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   8040
            TabIndex        =   85
            Top             =   1695
            Width           =   855
         End
         Begin VB.Label Lbl_PolMesGar 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   8040
            TabIndex        =   84
            Top             =   2055
            Width           =   855
         End
         Begin VB.Label Lbl_PolNumCar 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   2400
            TabIndex        =   83
            Top             =   2400
            Width           =   735
         End
         Begin VB.Label Lbl_PolIniVig 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   2400
            TabIndex        =   82
            Top             =   300
            Width           =   975
         End
         Begin VB.Label Lbl_PolTerVig 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   3840
            TabIndex        =   81
            Top             =   300
            Width           =   975
         End
         Begin VB.Label Lbl_PolAfp 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   2400
            TabIndex        =   80
            Top             =   2750
            Width           =   1815
         End
         Begin VB.Label Lbl_PolTipPen 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   2400
            TabIndex        =   79
            Top             =   1350
            Width           =   6615
         End
         Begin VB.Label Lbl_PolTipRta 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   2400
            TabIndex        =   77
            Top             =   1700
            Width           =   3255
         End
         Begin VB.Label Lbl_PolMod 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   2400
            TabIndex        =   76
            Top             =   2050
            Width           =   3255
         End
      End
      Begin VB.Frame Fra_GrupFam 
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
         Height          =   4365
         Left            =   -74760
         TabIndex        =   40
         Top             =   480
         Width           =   9375
         Begin VB.CheckBox chkBolElec 
            Caption         =   "Envía Boletas Electronicas"
            Enabled         =   0   'False
            Height          =   300
            Left            =   6225
            TabIndex        =   160
            Top             =   1995
            Width           =   2430
         End
         Begin VB.Label lbl_GrupoFono2 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   3120
            TabIndex        =   154
            Top             =   3450
            Width           =   915
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono 2"
            Height          =   255
            Index           =   8
            Left            =   2160
            TabIndex        =   153
            Top             =   3450
            Width           =   795
         End
         Begin VB.Label Lbl_GrupNombreSeg 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            TabIndex        =   128
            Top             =   1700
            Width           =   3015
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "2do.Nombre"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   127
            Top             =   1700
            Width           =   975
         End
         Begin VB.Label Lbl_Moneda 
            Caption         =   "(TM)"
            Enabled         =   0   'False
            Height          =   300
            Index           =   1
            Left            =   2640
            TabIndex        =   126
            Top             =   3800
            Width           =   495
         End
         Begin VB.Label Lbl_MtoPQ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pensión Quiebra"
            Height          =   195
            Left            =   5640
            TabIndex        =   122
            Top             =   3800
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.Label Lbl_MtoPensionQui 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   6840
            TabIndex        =   121
            Top             =   3800
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Ingreso"
            Height          =   195
            Index           =   4
            Left            =   6240
            TabIndex        =   118
            Top             =   1000
            Width           =   1020
         End
         Begin VB.Label Lbl_GrupFecIngreso 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   7680
            TabIndex        =   117
            Top             =   1000
            Width           =   1335
         End
         Begin VB.Label Lbl_GrupPension 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   1080
            TabIndex        =   74
            Top             =   3800
            Width           =   1455
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pensión "
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   73
            Top             =   3800
            Width           =   615
         End
         Begin VB.Label Lbl_GrupSitInv 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   7680
            TabIndex        =   72
            Top             =   300
            Width           =   975
         End
         Begin VB.Label Lbl_GrupFecNac 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   7680
            TabIndex        =   71
            Top             =   1350
            Width           =   1335
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sit. de Inv."
            Height          =   195
            Index           =   15
            Left            =   6240
            TabIndex        =   70
            Top             =   300
            Width           =   765
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Nacimiento"
            Height          =   195
            Index           =   16
            Left            =   6225
            TabIndex        =   69
            Top             =   1350
            Width           =   1290
         End
         Begin VB.Label Lbl_GrupNumIdent 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   3360
            TabIndex        =   68
            Top             =   1000
            Width           =   2175
         End
         Begin VB.Label Lbl_GrupTipoIdent 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   1080
            TabIndex        =   67
            Top             =   1000
            Width           =   1935
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono"
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   66
            Top             =   3450
            Width           =   795
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Email"
            Height          =   255
            Index           =   18
            Left            =   4200
            TabIndex        =   65
            Top             =   3450
            Width           =   495
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Ap. Materno"
            Height          =   255
            Index           =   19
            Left            =   120
            TabIndex        =   64
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Ap. Paterno"
            Height          =   255
            Index           =   21
            Left            =   120
            TabIndex        =   63
            Top             =   2050
            Width           =   855
         End
         Begin VB.Label Lbl_GrupNumOrden 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   1080
            TabIndex        =   62
            Top             =   300
            Width           =   495
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Orden"
            Height          =   255
            Index           =   44
            Left            =   120
            TabIndex        =   61
            Top             =   300
            Width           =   735
         End
         Begin VB.Label Lbl_GrupRegion 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   6360
            TabIndex        =   60
            Top             =   3100
            Width           =   2655
         End
         Begin VB.Label Lbl_GrupProvincia 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   3750
            TabIndex        =   59
            Top             =   3100
            Width           =   2535
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "1er.Nombre"
            Height          =   255
            Index           =   45
            Left            =   120
            TabIndex        =   58
            Top             =   1350
            Width           =   975
         End
         Begin VB.Label Lbl_Nombre 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
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
            Index           =   47
            Left            =   3120
            TabIndex        =   57
            Top             =   1000
            Width           =   135
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Identificación"
            Height          =   255
            Index           =   48
            Left            =   120
            TabIndex        =   56
            Top             =   1000
            Width           =   975
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Domicilio"
            Height          =   255
            Index           =   49
            Left            =   120
            TabIndex        =   55
            Top             =   2760
            Width           =   810
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Ubicación"
            Height          =   255
            Index           =   50
            Left            =   120
            TabIndex        =   54
            Top             =   3100
            Width           =   825
         End
         Begin VB.Label Lbl_GrupPar 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   1080
            TabIndex        =   53
            Top             =   650
            Width           =   5055
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Parentesco"
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   52
            Top             =   650
            Width           =   855
         End
         Begin VB.Label Lbl_GrupEstado 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   7680
            TabIndex        =   51
            Top             =   650
            Width           =   975
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Est. Vigente"
            Height          =   195
            Index           =   51
            Left            =   6240
            TabIndex        =   50
            Top             =   650
            Width           =   855
         End
         Begin VB.Label Lbl_GrupFecFall 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   7680
            TabIndex        =   49
            Top             =   1700
            Width           =   1335
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Fall."
            Height          =   195
            Index           =   52
            Left            =   6240
            TabIndex        =   48
            Top             =   1700
            Width           =   780
         End
         Begin VB.Label Lbl_GrupNombre 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            TabIndex        =   47
            Top             =   1350
            Width           =   3015
         End
         Begin VB.Label Lbl_GrupPaterno 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            TabIndex        =   46
            Top             =   2050
            Width           =   3015
         End
         Begin VB.Label Lbl_GrupDomicilio 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            TabIndex        =   45
            Top             =   2750
            Width           =   7935
         End
         Begin VB.Label Lbl_GrupComuna 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            TabIndex        =   44
            Top             =   3100
            Width           =   2595
         End
         Begin VB.Label Lbl_GrupFono 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            TabIndex        =   43
            Top             =   3450
            Width           =   945
         End
         Begin VB.Label Lbl_GrupMail 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   4680
            TabIndex        =   42
            Top             =   3450
            Width           =   4305
         End
         Begin VB.Label Lbl_GrupMaterno 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            TabIndex        =   41
            Top             =   2400
            Width           =   3015
         End
      End
      Begin VB.Frame Fra_FormaPago 
         Caption         =   "Forma de Pago de Pensión"
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
         Height          =   2910
         Left            =   -74850
         TabIndex        =   29
         Top             =   380
         Width           =   4575
         Begin VB.Label Lbl_FPNumCtaCCI 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   164
            Top             =   2250
            Width           =   3195
         End
         Begin VB.Label Label6 
            Caption         =   "N°Cuenta CCI"
            Height          =   255
            Left            =   120
            TabIndex        =   163
            Top             =   2250
            Width           =   1095
         End
         Begin VB.Label Lbl_MonedaCta 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   162
            Top             =   1580
            Width           =   3195
         End
         Begin VB.Label Label4 
            Caption         =   "Moneda"
            Height          =   255
            Left            =   120
            TabIndex        =   161
            Top             =   1580
            Width           =   975
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Tipo Cta."
            Height          =   255
            Index           =   59
            Left            =   120
            TabIndex        =   39
            Top             =   1250
            Width           =   825
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Vía Pago"
            Height          =   255
            Index           =   60
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   810
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Banco"
            Height          =   255
            Index           =   61
            Left            =   120
            TabIndex        =   37
            Top             =   900
            Width           =   825
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "N°Cuenta"
            Height          =   255
            Index           =   62
            Left            =   120
            TabIndex        =   36
            Top             =   1910
            Width           =   795
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Sucursal"
            Height          =   255
            Index           =   63
            Left            =   120
            TabIndex        =   35
            Top             =   570
            Width           =   810
         End
         Begin VB.Label Lbl_FPViaPago 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   34
            Top             =   240
            Width           =   3195
         End
         Begin VB.Label Lbl_FPSucursal 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   33
            Top             =   570
            Width           =   3195
         End
         Begin VB.Label Lbl_FPTipoCta 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   32
            Top             =   1250
            Width           =   3195
         End
         Begin VB.Label Lbl_FPBanco 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   31
            Top             =   900
            Width           =   3195
         End
         Begin VB.Label Lbl_FPNumCta 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   30
            Top             =   1910
            Width           =   3195
         End
      End
      Begin VB.Frame Fra_PlanSalud 
         Caption         =   "Plan de Salud"
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
         Height          =   1515
         Left            =   -70200
         TabIndex        =   22
         Top             =   380
         Width           =   4935
         Begin VB.Label Lbl_PSFechaEfecto 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            TabIndex        =   120
            Top             =   1200
            Width           =   1260
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Periodo Efecto"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   119
            Top             =   1200
            Width           =   1395
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Institución"
            Height          =   255
            Index           =   65
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   810
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Modalidad de Pago"
            Height          =   255
            Index           =   66
            Left            =   120
            TabIndex        =   27
            Top             =   560
            Width           =   1425
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Monto de Pago"
            Height          =   255
            Index           =   67
            Left            =   120
            TabIndex        =   26
            Top             =   870
            Width           =   1395
         End
         Begin VB.Label Lbl_PSInstitucion 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            TabIndex        =   25
            Top             =   240
            Width           =   3135
         End
         Begin VB.Label Lbl_PSModPago 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            TabIndex        =   24
            Top             =   555
            Width           =   3135
         End
         Begin VB.Label Lbl_PSMtoPago 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            TabIndex        =   23
            Top             =   870
            Width           =   1260
         End
      End
      Begin VB.Frame Fra_Receptor 
         Caption         =   "Receptor de Pensión"
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
         Height          =   1365
         Left            =   -70200
         TabIndex        =   11
         Top             =   1920
         Width           =   4935
         Begin VB.Label Lbl_RecFecIniVig 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   1800
            TabIndex        =   15
            Top             =   620
            Width           =   975
         End
         Begin VB.Label Lbl_RecFecTerVig 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   3120
            TabIndex        =   14
            Top             =   620
            Width           =   975
         End
         Begin VB.Label Lbl_RecRut 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   1095
            TabIndex        =   13
            Top             =   300
            Width           =   1695
         End
         Begin VB.Label Label12 
            Caption         =   "Identificación"
            Height          =   300
            Left            =   120
            TabIndex        =   21
            Top             =   300
            Width           =   1035
         End
         Begin VB.Label Label13 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2880
            TabIndex        =   20
            Top             =   285
            Width           =   165
         End
         Begin VB.Label Lbl_RecNombre 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   720
            TabIndex        =   19
            Top             =   945
            Width           =   3945
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Nombre"
            Height          =   255
            Index           =   53
            Left            =   90
            TabIndex        =   18
            Top             =   950
            Width           =   855
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Periodo de Vigencia"
            Height          =   255
            Index           =   54
            Left            =   120
            TabIndex        =   17
            Top             =   620
            Width           =   1695
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
            Index           =   55
            Left            =   2805
            TabIndex        =   16
            Top             =   620
            Width           =   255
         End
         Begin VB.Label Lbl_RecDgv 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   3120
            TabIndex        =   12
            Top             =   300
            Width           =   1485
         End
      End
      Begin VB.Frame Fra_RetJud 
         Caption         =   "Retención Judicial"
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
         Height          =   1545
         Left            =   -74880
         TabIndex        =   9
         Top             =   3360
         Width           =   9580
         Begin MSFlexGridLib.MSFlexGrid Msf_GrillaRetJud 
            Height          =   1215
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   9345
            _ExtentX        =   16484
            _ExtentY        =   2143
            _Version        =   393216
            Rows            =   1
            Cols            =   1
            FixedCols       =   0
            BackColor       =   14745599
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
      Begin MSFlexGridLib.MSFlexGrid Msf_GrillaHabDes 
         Height          =   2385
         Left            =   180
         TabIndex        =   141
         Top             =   2520
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   4207
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "Frm_Consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vlFecha As String
Dim vlFechaIni As String
Dim vlFechaTer As String
Dim vlCodPar As String
Dim vlNumEndoso As Integer
Dim vlNumOrden As Integer
Dim vlCodConceptos As String
Dim vlNombreComuna As String
Dim vlNombreSucursal As String
Dim vlPeriodo As String
Dim vlCodTipReceptor As String
Dim vlTipoIden As String
Dim vlNumIden As String
Dim vlIdenEmpresa As String
Dim vlCodTp As String
Dim vlCodTr As String
Dim vlCodAl As String
Dim vlCodPa As String
Dim vlFecNacTitular As String

Dim vlRutReceptor As Integer
Dim vlRutRec As Double
Dim vlDgvRec As String
Dim vlNomRec As String
Dim vlNumCargas As Integer
Dim vlNombreEstado As String
Dim vlCodHabDesCCAF As String
Dim vlNombreConcepto As String
Dim vlNombreCCAF As String
Dim vlTipoRet As String
Dim vlModRet As String
Dim vlRutAux As Double
Dim vlMtoCarga As Double
Dim vlEstCerEst As String * 1
Dim vlFechaActual As String
Dim vlCodCauSus As String
Dim vlCodEstado As String
Dim vlArchivo As String
Dim vlCodInsSalud As String
Dim vlFechaEfectoSalud As String
Dim vlRut As String
Dim vlOrden As Integer
Dim vlRutCia As String
Dim vlTipoIdenCia As String
Dim vlNumIdenCia As String
Dim vlNombreUsuario As String
Dim vlCiuCia As String
Dim vlFechaEfectoAsigFam As String
Dim vlPago As String
Dim vlFechaDesde As String
Dim vlFechaHasta As String
Dim vlRutCliente As String
Dim vltipoidenCompania As String
Dim vlcobertura As String

Dim vlOpcionPago As String

Dim vlNumPerPago As String
Dim vlNumCargasPagadas As Integer
Dim vlNombreBenef As String
Dim vlRutBenef As String
Dim vlTipoRenta As String
Dim vlTipoPension As String
Dim vlTipoModalidad As String
Dim vlFechaVigenciaRta As String
Dim vlMtoPensionBruta As Double

Dim vlNumEndosoNoBen As Integer

Dim vlRegistro2 As ADODB.Recordset
Dim vlRegistro3 As ADODB.Recordset

Dim vlNombreCompania As String
Dim vlDirCompania As String
Dim vlFonoCompania As String

Const clFechaTopeTer As String * 8 = "99991231"
Const clRecTutor As String * 1 = "T"
Const clRecPensionado As String * 1 = "P"
Const clModOrigenCCAF As String * 4 = "CCAF"
Const clCargaActiva As String * 1 = "A"
Const clCodSinDerPen As String * 2 = "10"
Const clDeptoUsuario As String = "Servicio al Cliente"
Const clCodEstPen99 As String * 2 = 99
Const clOpcionDEF As String * 3 = "DEF"

Const clCodTipReceptorR As String * 1 = "R"

'CMV-20061102 I
Dim vlNumOrd As Integer
Dim vlUltimoPerPago As String
Dim vlPrcCastigoQui As Double
Dim vlTopeMaxQui As Double

Const clCodEstadoC As String * 1 = "C"
'CMV-20061102 F

Dim vlLargoTipoIden    As Integer 'sirve para llenar la grilla
Dim vlPosicionTipoIden As Integer 'sirve para llenar la grilla

Dim vlCodTipoIdenBenCau As String
Dim vlNumIdenBenCau As String

Dim vlCodTipoIdenBenTut As String
Dim vlNumIdenBenTut As String

Dim vlNombreSeg As String, vlApMaterno As String

Dim vlNombreApoderado As String
Dim vlCargoApoderado As String

Private Function flLlenaTemporal_masivo(iFecDesde, iFecHasta, iPoliza, iTipoIden, iNumIden, iGlosaOpcion, iPago) As Boolean

    Dim vlSql As String, vlTB As ADODB.Recordset
    Dim vlNumConceptosHab As Long, vlNumConceptosDesc As Long
    Dim vlItem As Long, vlPoliza As String, vlOrden As String
    Dim vlIndImponible As String, vlIndTributable As String
    Dim vlTipReceptor As String
    Dim vlNumIdenReceptor As String, vlCodTipoIdenReceptor As Long
    Dim vlPerPago As String
    Dim vlTB2 As ADODB.Recordset
    Dim vlTipPension As String
    Dim vlViaPago As String
    Dim vlCajaComp As String
    Dim vlInsSalud As String
    Dim vlCodDireccion As Double
    Dim vlAfp As String
    Dim vlFecPago As String
    Dim vlDescViaPago As String
    Dim vlSucursal As String
    Dim vlDescSucursal As String
    
    flLlenaTemporal_masivo = False
    vlItem = 1
    vlNumConceptosHab = 0
    vlNumConceptosDesc = 0
    vlPoliza = ""
    vlOrden = 0
    vlPerPago = ""
    vlFecPago = ""
    vlTipPension = ""
    vlViaPago = ""
    vlCajaComp = ""
    vlInsSalud = ""
    vlCodDireccion = 0
    vlAfp = ""
    vlSucursal = ""
    vlDescSucursal = ""
    'VARIABLES GENERALES
    stTTMPLiquidacion.Cod_Usuario = vgUsuario
    
    'Elimina Datos de la Tabla Temporal
    vlSql = "DELETE FROM PP_TTMP_LIQUIDACION WHERE COD_USUARIO = '" & vgUsuario & "'"
    vgConexionBD.Execute (vlSql)
    
    vlSql = "SELECT P.NUM_POLIZA, L.NUM_ENDOSO, P.NUM_ORDEN, P.COD_CONHABDES, P.MTO_CONHABDES, C.COD_TIPMOV, P.NUM_PERPAGO, P.COD_TIPOIDENRECEPTOR,P.NUM_IDENRECEPTOR, P.COD_TIPRECEPTOR,"
    vlSql = vlSql & " L.GLS_DIRECCION, L.FEC_PAGO, L.GLS_NOMRECEPTOR, L.GLS_NOMSEGRECEPTOR, L.GLS_PATRECEPTOR,"
    vlSql = vlSql & " L.GLS_MATRECEPTOR, L.MTO_LIQPAGAR, L.COD_DIRECCION, "
    vlSql = vlSql & " L.COD_TIPPENSION, L.COD_VIAPAGO, L.COD_SUCURSAL, L.COD_INSSALUD," ''*L.COD_CAJACOMPEN,
    vlSql = vlSql & " L.MTO_PENSION, L.NUM_CARGAS, L.MTO_HABER, L.MTO_DESCUENTO," ''*L.DGV_RECEPTOR,
    vlSql = vlSql & " B.NUM_IDENBEN, B.COD_TIPOIDENBEN, B.GLS_NOMBEN, B.GLS_NOMSEGBEN, B.GLS_PATBEN, B.GLS_MATBEN, "
    vlSql = vlSql & " C.GLS_CONHABDES, M.COD_SCOMP, POL.COD_AFP, L.COD_MONEDA, M.GLS_ELEMENTO AS MONEDA"
    vlSql = vlSql & " FROM PP_TMAE_PAGOPEN" & iGlosaOpcion & " P, PP_TMAE_LIQPAGOPEN" & iGlosaOpcion & " L, MA_TPAR_CONHABDES C"
    vlSql = vlSql & ", PP_TMAE_POLIZA POL, PP_TMAE_BEN B, MA_TPAR_TABCOD M  WHERE"
    vlSql = vlSql & " L.NUM_POLIZA = B.NUM_POLIZA AND"
    vlSql = vlSql & " L.NUM_ENDOSO = B.NUM_ENDOSO AND"
    vlSql = vlSql & " L.NUM_ORDEN = B.NUM_ORDEN AND"
    vlSql = vlSql & " L.NUM_POLIZA = POL.NUM_POLIZA AND"
    vlSql = vlSql & " L.NUM_ENDOSO = POL.NUM_ENDOSO AND"
    vlSql = vlSql & " L.COD_MONEDA = M.COD_ELEMENTO AND"
    vlSql = vlSql & " M.COD_TABLA = 'TM' AND" 'Tabla de Monedas
    'If Chk_Pensionado.Value = 1 Then
        'If Txt_Poliza <> "" Then
        '************ numeros de polizas a imprimir ********************************
        vlSql = vlSql & " L.NUM_POLIZA IN (002,003,007,008,009,011,012,014,020,021,023,024,027,028,"
        vlSql = vlSql & " 029,031,034,035,037,041,042,044,046,050,052,053,054,"
        vlSql = vlSql & " 056,057,060,064,066,067,068,069,071,076,079,080,081,083,"
        vlSql = vlSql & " 085,087,092,094,095,099,106,107,108,109,114,118,124,"
        vlSql = vlSql & " 143,187,194,197,203,207,212,220,221,233,247,256,295) AND "
        '***************************************************************************
        'If Txt_Rut <> "" Then
        If iTipoIden <> "" Then
            vlSql = vlSql & " B.COD_TIPOIDENBEN = " & Trim(iTipoIden) & " AND"
        End If
        If iNumIden <> "" Then
            vlSql = vlSql & " B.NUM_IDENBEN = '" & Trim(iNumIden) & "' AND"
        End If
'    End If
    vlSql = vlSql & " L.NUM_POLIZA = P.NUM_POLIZA"
    vlSql = vlSql & " AND L.NUM_ORDEN = P.NUM_ORDEN"
    ''*vlSql = vlSql & " AND L.RUT_RECEPTOR = P.RUT_RECEPTOR"
    vlSql = vlSql & " AND L.COD_TIPOIDENRECEPTOR=P.COD_TIPOIDENRECEPTOR"
    vlSql = vlSql & " AND L.NUM_IDENRECEPTOR=P.NUM_IDENRECEPTOR"
    vlSql = vlSql & " AND L.COD_TIPRECEPTOR = P.COD_TIPRECEPTOR"
    vlSql = vlSql & " AND L.NUM_PERPAGO = P.NUM_PERPAGO"
    If iPago = "P" Then 'PRIMER PAGO
        vlSql = vlSql & " AND L.COD_TIPOPAGO = 'P'"
    ElseIf iPago = "R" Then 'PAGO EN REGIMEN
        vlSql = vlSql & " AND L.COD_TIPOPAGO = 'R'"
    End If
    'I--- ABV 12/03/2005 ---
    If (iPago = "T") Then
        vlSql = vlSql & " AND L.COD_TIPOPAGO in ('R','P')"
    End If
    'F--- ABV 12/03/2005 ---
    vlSql = vlSql & " AND L.FEC_PAGO >= '" & iFecDesde & "' AND L.FEC_PAGO <= '" & iFecHasta & "'"
    vlSql = vlSql & " AND P.COD_CONHABDES  = C.COD_CONHABDES"
    'vlSql = vlSql & " ORDER BY P.NUM_POLIZA, P.NUM_ORDEN, P.RUT_RECEPTOR, P.COD_TIPRECEPTOR,"
    vlSql = vlSql & " ORDER BY P.NUM_PERPAGO, P.NUM_POLIZA, P.NUM_ORDEN,P.NUM_IDENRECEPTOR, P.COD_TIPOIDENRECEPTOR, P.COD_TIPRECEPTOR," 'hqr 17/03/2005 Se agrega número de periodo
    vlSql = vlSql & " C.COD_IMPONIBLE DESC, C.COD_TRIBUTABLE DESC, C.COD_TIPMOV DESC"
    Set vlTB = vgConexionBD.Execute(vlSql)
    If Not vlTB.EOF Then
        Do While Not vlTB.EOF
            If vlPoliza <> vlTB!num_poliza Or vlOrden <> vlTB!Num_Orden Or vlNumIdenReceptor <> vlTB!Num_IdenReceptor Or vlCodTipoIdenReceptor <> vlTB!Cod_TipoIdenReceptor Or vlTipReceptor <> vlTB!Cod_TipReceptor Or (vlPerPago <> vlTB!Num_PerPago And vlPago = "R") Or (vlFecPago <> vlTB!fec_pago And vlPago = "P") Then 'hqr 17/03/2006 Se agrega número de periodo
                'Reinicia el Contador
                vlItem = 1
                vlPoliza = vlTB!num_poliza
                vlOrden = vlTB!Num_Orden
                vlNumConceptosHab = 0
                vlNumConceptosDesc = 0
                vlNumIdenReceptor = vlTB!Num_IdenReceptor
                vlCodTipoIdenReceptor = vlTB!Cod_TipoIdenReceptor
                vlTipReceptor = vlTB!Cod_TipReceptor
                stTTMPLiquidacion.num_poliza = vlPoliza
                stTTMPLiquidacion.Num_IdenReceptor = vlTB!Num_IdenReceptor
                stTTMPLiquidacion.Cod_TipoIdenReceptor = vlTB!Cod_TipoIdenReceptor
                stTTMPLiquidacion.Num_Orden = vlOrden
                stTTMPLiquidacion.num_endoso = IIf(IsNull(vlTB!num_endoso), 0, vlTB!num_endoso)
                stTTMPLiquidacion.Cod_TipReceptor = vlTB!Cod_TipReceptor
                If stTTMPLiquidacion.Cod_TipReceptor <> "R" Then
                    stTTMPLiquidacion.Gls_Direccion = IIf(IsNull(vlTB!Gls_Direccion), "", vlTB!Gls_Direccion)
                Else
                    stTTMPLiquidacion.Gls_Direccion = " "
                End If
                stTTMPLiquidacion.Gls_NomReceptor = vlTB!Gls_NomReceptor & " " & IIf(IsNull(vlTB!Gls_NomSegReceptor), "", vlTB!Gls_NomSegReceptor & " ") & vlTB!Gls_PatReceptor & IIf(IsNull(vlTB!Gls_MatReceptor), "", " " + vlTB!Gls_MatReceptor)
                If stTTMPLiquidacion.Cod_TipReceptor <> "R" Then
                    stTTMPLiquidacion.Cod_Direccion = vlTB!Cod_Direccion
                Else
                    stTTMPLiquidacion.Cod_Direccion = "0"
                End If
                stTTMPLiquidacion.Cod_TipoIdenBen = vlTB!Cod_TipoIdenBen
                stTTMPLiquidacion.Num_IdenBen = vlTB!Num_IdenBen
                stTTMPLiquidacion.Gls_NomBen = vlTB!Gls_NomBen & " " & IIf(IsNull(vlTB!Gls_NomSegBen), "", vlTB!Gls_NomSegBen & " ") & vlTB!Gls_PatBen & IIf(IsNull(vlTB!Gls_MatBen), "", " " & vlTB!Gls_MatBen)
                'para primeros pagos
                stTTMPLiquidacion.Mto_LiqPagar = 0
                stTTMPLiquidacion.Num_Cargas = 0 'vlTB!Num_Cargas
                stTTMPLiquidacion.Mto_LiqHaber = 0
                stTTMPLiquidacion.Mto_LiqDescuento = 0
                'fin primeros pagos
                'Obtiene Fecha de Término del Poder Notarial
                If stTTMPLiquidacion.Cod_TipReceptor <> "R" Then
                    vlSql = "SELECT tut.fec_terpodnot FROM pp_tmae_tutor tut"
                    vlSql = vlSql & " WHERE tut.num_poliza = '" & stTTMPLiquidacion.num_poliza & "'"
                    vlSql = vlSql & " AND tut.num_orden = " & stTTMPLiquidacion.Num_Orden
                    vlSql = vlSql & " AND tut.cod_tipoidentut = " & vlTB!Cod_TipoIdenReceptor & " "
                    vlSql = vlSql & " AND tut.num_identut = '" & vlTB!Num_IdenReceptor & "' "
                    
                    Set vlTB2 = vgConexionBD.Execute(vlSql)
                    If Not vlTB2.EOF Then
                        stTTMPLiquidacion.Fec_TerPodNot = vlTB2!Fec_TerPodNot
                        stTTMPLiquidacion.Fec_TerPodNot = DateSerial(Mid(stTTMPLiquidacion.Fec_TerPodNot, 1, 4), Mid(stTTMPLiquidacion.Fec_TerPodNot, 5, 2), Mid(stTTMPLiquidacion.Fec_TerPodNot, 7, 2))
                    Else
                        stTTMPLiquidacion.Fec_TerPodNot = ""
                    End If
                    'Obtiene Mensajes al Beneficiario
                    stTTMPLiquidacion.Gls_Mensaje = ""
                    vlSql = "SELECT par.gls_mensaje FROM pp_tmae_menpoliza men, pp_tpar_mensaje par"
                    vlSql = vlSql & " WHERE par.cod_mensaje = men.cod_mensaje"
                    vlSql = vlSql & " AND men.num_poliza = '" & stTTMPLiquidacion.num_poliza & "'"
                    vlSql = vlSql & " AND men.num_orden = " & stTTMPLiquidacion.Num_Orden
                    vlSql = vlSql & " AND men.num_perpago = '" & stTTMPLiquidacion.Num_PerPago & "'"
                    Set vlTB2 = vgConexionBD.Execute(vlSql)
                    If Not vlTB2.EOF Then
                        Do While Not vlTB2.EOF
                            stTTMPLiquidacion.Gls_Mensaje = stTTMPLiquidacion.Gls_Mensaje & vlTB2!Gls_Mensaje & Chr(13)
                            vlTB2.MoveNext
                        Loop
                    Else
                        stTTMPLiquidacion.Gls_Mensaje = ""
                    End If
                Else
                    stTTMPLiquidacion.Fec_TerPodNot = ""
                    stTTMPLiquidacion.Gls_Mensaje = ""
                End If
            End If

            stTTMPLiquidacion.Num_PerPago = vlTB!Num_PerPago
            'stTTMPLiquidacion.Mto_LiqPagar = vlTB!Mto_LiqPagar
            stTTMPLiquidacion.Mto_Pension = vlTB!Mto_Pension
            stTTMPLiquidacion.Num_Cargas = 0 'vlTB!Num_Cargas
            'stTTMPLiquidacion.Mto_LiqHaber = vlTB!Mto_Haber
            'stTTMPLiquidacion.Mto_LiqDescuento = vlTB!Mto_Descuento
            If vlPerPago <> stTTMPLiquidacion.Num_PerPago Then
                stTTMPLiquidacion.fec_pago = vlTB!fec_pago
                'Obtiene Fecha del Próximo Pago
'                vlSql = "SELECT pro.fec_pagoproxreg"
'                vlSql = vlSql & " FROM pp_tmae_propagopen pro"
'                vlSql = vlSql & " WHERE pro.num_perpago = '" & stTTMPLiquidacion.Num_PerPago & "'"
'                Set vlTB2 = vgConexionBD.Execute(vlSql)
'                If Not vlTB2.EOF Then
'                    stTTMPLiquidacion.Fec_PagoProxReg = vlTB2!Fec_PagoProxReg
'                    stTTMPLiquidacion.Fec_PagoProxReg = DateSerial(Mid(stTTMPLiquidacion.Fec_PagoProxReg, 1, 4), Mid(stTTMPLiquidacion.Fec_PagoProxReg, 5, 2), Mid(stTTMPLiquidacion.Fec_PagoProxReg, 7, 2))
'                Else
                    stTTMPLiquidacion.Fec_PagoProxReg = ""
'                End If
                'Obtiene Valor UF
'                vlSql = "SELECT val.mto_moneda"
'                vlSql = vlSql & " FROM ma_tval_moneda val"
'                vlSql = vlSql & " WHERE val.cod_moneda = 'UF'"
'                vlSql = vlSql & " AND val.fec_moneda = '" & stTTMPLiquidacion.Fec_Pago & "'"
'                Set vlTB2 = vgConexionBD.Execute(vlSql)
'                If Not vlTB2.EOF Then
'                    stTTMPLiquidacion.Mto_Moneda = vlTB2!Mto_Moneda
'                Else
                    stTTMPLiquidacion.Mto_Moneda = 0
'                End If
                stTTMPLiquidacion.fec_pago = DateSerial(Mid(stTTMPLiquidacion.fec_pago, 1, 4), Mid(stTTMPLiquidacion.fec_pago, 5, 2), Mid(stTTMPLiquidacion.fec_pago, 7, 2))
                vlPerPago = stTTMPLiquidacion.Num_PerPago
            End If
            'Obtiene Tipo de Pensión
            If vlTipPension <> vlTB!Cod_TipPension Then
                vlSql = "SELECT tab.gls_elemento FROM ma_tpar_tabcod tab"
                vlSql = vlSql & " WHERE tab.cod_tabla = 'TP'"
                vlSql = vlSql & " AND tab.cod_elemento = '" & vlTB!Cod_TipPension & "'"
                Set vlTB2 = vgConexionBD.Execute(vlSql)
                If Not vlTB2.EOF Then
                    stTTMPLiquidacion.Gls_TipPension = vlTB2!GLS_ELEMENTO
                Else
                    stTTMPLiquidacion.Gls_TipPension = ""
                End If
                vlTipPension = vlTB!Cod_TipPension
            End If
            
            'Obtiene Via de Pago
            If vlViaPago <> vlTB!Cod_ViaPago Then
                vlSql = "SELECT tab.gls_elemento FROM ma_tpar_tabcod tab"
                vlSql = vlSql & " WHERE tab.cod_tabla = 'VPG'"
                vlSql = vlSql & " AND tab.cod_elemento = '" & vlTB!Cod_ViaPago & "'"
                Set vlTB2 = vgConexionBD.Execute(vlSql)
                If Not vlTB2.EOF Then
                    vlDescViaPago = vlTB2!GLS_ELEMENTO
                Else
                    vlDescViaPago = ""
                End If
                vlViaPago = vlTB!Cod_ViaPago
            End If
            
            'hqr 13/10/2007 Obtiene Sucursal de la Via de Pago
            If vlTB!Cod_ViaPago = "04" Then
                If vlTB!Cod_Sucursal <> vlSucursal Then
                    'Obtiene Sucursal
                    stTTMPLiquidacion.Gls_ViaPago = vlDescViaPago
                    vlSql = "SELECT a.gls_sucursal FROM ma_tpar_sucursal a"
                    vlSql = vlSql & " WHERE a.cod_sucursal = '" & vlTB!Cod_Sucursal & "'"
                    vlSql = vlSql & " AND a.cod_tipo = 'A'" 'AFP
                    Set vlTB2 = vgConexionBD.Execute(vlSql)
                    If Not vlTB2.EOF Then
                        vlDescSucursal = vlTB2!gls_sucursal
                    End If
                    vlSucursal = vlTB!Cod_Sucursal
                End If
                stTTMPLiquidacion.Gls_ViaPago = Mid(vlDescViaPago & " - " & vlDescSucursal, 1, 50)
            Else
                stTTMPLiquidacion.Gls_ViaPago = vlDescViaPago
            End If
                        
            stTTMPLiquidacion.Gls_CajaComp = ""
            
            'Obtiene Institución de Salud
            If Not IsNull(vlTB!Cod_InsSalud) Then
                If vlTB!Cod_InsSalud <> "NULL" Then
                    If vlInsSalud <> vlTB!Cod_InsSalud Then
                        vlSql = "SELECT tab.gls_elemento FROM ma_tpar_tabcod tab"
                        vlSql = vlSql & " WHERE tab.cod_tabla = 'IS'"
                        vlSql = vlSql & " AND tab.cod_elemento = '" & vlTB!Cod_InsSalud & "'"
                        Set vlTB2 = vgConexionBD.Execute(vlSql)
                        If Not vlTB2.EOF Then
                            stTTMPLiquidacion.Gls_InsSalud = vlTB2!GLS_ELEMENTO
                        Else
                            stTTMPLiquidacion.Gls_InsSalud = ""
                        End If
                        vlInsSalud = vlTB!Cod_InsSalud
                    End If
                Else
                    vlInsSalud = ""
                    stTTMPLiquidacion.Gls_InsSalud = ""
                End If
            Else
                vlInsSalud = ""
                stTTMPLiquidacion.Gls_InsSalud = ""
            End If
            
            'Obtiene AFP
            If vlAfp <> vlTB!cod_afp Then
                vlSql = "SELECT tab.gls_elemento FROM ma_tpar_tabcod tab"
                vlSql = vlSql & " WHERE tab.cod_tabla = 'AF'"
                vlSql = vlSql & " AND tab.cod_elemento = '" & vlTB!cod_afp & "'"
                Set vlTB2 = vgConexionBD.Execute(vlSql)
                If Not vlTB2.EOF Then
                    stTTMPLiquidacion.Gls_Afp = "AFP " & vlTB2!GLS_ELEMENTO
                Else
                    stTTMPLiquidacion.Gls_Afp = "AFP"
                End If
                vlAfp = vlTB!cod_afp
            End If
                        
            'Obtiene Direccion
            If stTTMPLiquidacion.Cod_Direccion <> vlCodDireccion Then
                If stTTMPLiquidacion.Cod_Direccion <> 0 Then
                    vlSql = "SELECT com.gls_comuna, prov.gls_provincia, reg.gls_region"
                    vlSql = vlSql & " FROM ma_tpar_comuna com, ma_tpar_provincia prov, ma_tpar_region reg"
                    vlSql = vlSql & " WHERE reg.cod_region = prov.cod_region"
                    vlSql = vlSql & " AND prov.cod_region = com.cod_region"
                    vlSql = vlSql & " AND prov.cod_provincia = com.cod_provincia"
                    vlSql = vlSql & " AND com.cod_direccion = '" & stTTMPLiquidacion.Cod_Direccion & "'"
                    Set vlTB2 = vgConexionBD.Execute(vlSql)
                    If Not vlTB2.EOF Then
                        stTTMPLiquidacion.Gls_Direccion2 = vlTB2!gls_region & " - " & vlTB2!gls_provincia & " - " & vlTB2!gls_comuna
                    Else
                        stTTMPLiquidacion.Gls_Direccion2 = ""
                    End If
                Else
                    stTTMPLiquidacion.Gls_Direccion2 = ""
                End If
                vlCodDireccion = stTTMPLiquidacion.Cod_Direccion
            End If
            
            'Obtiene Datos
            stTTMPLiquidacion.Cod_Moneda = vlTB!COD_SCOMP
            If vlPago = "R" Then
                stTTMPLiquidacion.Mto_LiqPagar = vlTB!Mto_LiqPagar
                stTTMPLiquidacion.Mto_LiqHaber = vlTB!Mto_Haber
                stTTMPLiquidacion.Mto_LiqDescuento = vlTB!Mto_Descuento
                stTTMPLiquidacion.Gls_MontoPension = fgConvierteNumeroLetras(stTTMPLiquidacion.Mto_LiqPagar, vlTB!Moneda)
            End If
            If vlTB!cod_tipmov = "H" Then 'haber
                If vlPago <> "R" Then
                    stTTMPLiquidacion.Mto_LiqHaber = stTTMPLiquidacion.Mto_LiqHaber + vlTB!Mto_ConHabDes
                    stTTMPLiquidacion.Mto_LiqPagar = stTTMPLiquidacion.Mto_LiqPagar + vlTB!Mto_ConHabDes
                    stTTMPLiquidacion.Gls_MontoPension = fgConvierteNumeroLetras(stTTMPLiquidacion.Mto_LiqPagar, vlTB!Moneda)
                End If
                stTTMPLiquidacion.Cod_ConDescto = "NULL"
                stTTMPLiquidacion.Mto_Descuento = 0
                stTTMPLiquidacion.Cod_ConHaber = vlTB!gls_ConHabDes
                stTTMPLiquidacion.Mto_Haber = vlTB!Mto_ConHabDes
                vlNumConceptosHab = vlNumConceptosHab + 1
                stTTMPLiquidacion.Num_Item = vlNumConceptosHab
                If vlNumConceptosHab > vlNumConceptosDesc Then
                    Call fgInsertaTTMPLiquidacion(stTTMPLiquidacion)
                Else
                    Call fgActualizaTTMPLiquidacionHab(stTTMPLiquidacion)
                End If
            ElseIf vlTB!cod_tipmov = "D" Then 'descuento
                If vlPago <> "R" Then
                    stTTMPLiquidacion.Mto_LiqDescuento = stTTMPLiquidacion.Mto_LiqDescuento + vlTB!Mto_ConHabDes
                    stTTMPLiquidacion.Mto_LiqPagar = stTTMPLiquidacion.Mto_LiqPagar - vlTB!Mto_ConHabDes
                    stTTMPLiquidacion.Gls_MontoPension = fgConvierteNumeroLetras(stTTMPLiquidacion.Mto_LiqPagar, vlTB!Moneda)
                End If
                stTTMPLiquidacion.Cod_ConDescto = vlTB!gls_ConHabDes
                stTTMPLiquidacion.Mto_Descuento = vlTB!Mto_ConHabDes
                stTTMPLiquidacion.Cod_ConHaber = "NULL"
                stTTMPLiquidacion.Mto_Haber = 0
                vlNumConceptosDesc = vlNumConceptosDesc + 1
                stTTMPLiquidacion.Num_Item = vlNumConceptosDesc
                If vlNumConceptosDesc > vlNumConceptosHab Then
                    Call fgInsertaTTMPLiquidacion(stTTMPLiquidacion)
                Else
                    Call fgActualizaTTMPLiquidacionDesc(stTTMPLiquidacion)
                End If
            Else 'OTROS
                stTTMPLiquidacion.Cod_ConDescto = vlTB!Cod_ConHabDes
                stTTMPLiquidacion.Mto_Haber = vlTB!Mto_ConHabDes
                stTTMPLiquidacion.Num_Item = 0
                'Call fgInsertaTTMPLiquidacion(stTTMPLiquidacion)
            End If
            vlFecPago = vlTB!fec_pago
            vlItem = vlItem + 1
            vlTB.MoveNext
        Loop
    Else
        MsgBox "No existe Información para este rango de Fechas", vbInformation, "Operacion Cancelada"
        Exit Function
    End If
    flLlenaTemporal_masivo = True

End Function

Private Function flCargarListaConHabDes() As Boolean
    
On Error GoTo Err_flCargarListaConHabDes

    flCargarListaConHabDes = False
    
    Screen.MousePointer = vbHourglass
    
    Lst_LiqSeleccion.Clear
    
    vgSql = ""
    vgSql = "SELECT cod_conhabdes,gls_conhabdes,cod_tipmov "
    vgSql = vgSql & "FROM MA_TPAR_CONHABDES "
    vgSql = vgSql & "ORDER BY cod_conhabdes ASC "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       Do While Not vgRegistro.EOF
          If Lst_LiqSeleccion.ListCount = 1 Then
            Lst_LiqSeleccion.AddItem (" TODOS "), 0
          End If
          Lst_LiqSeleccion.AddItem (" " & Trim(vgRegistro!Cod_ConHabDes) & " - " & Trim(vgRegistro!gls_ConHabDes) & "   (" & Trim(vgRegistro!cod_tipmov) & ")")
            
          vgRegistro.MoveNext
       Loop
    End If
    
    Set vgRegistro = Nothing
    Screen.MousePointer = vbDefault
    
    flCargarListaConHabDes = True

Exit Function
Err_flCargarListaConHabDes:
    Screen.MousePointer = vbDefault
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flRecibe(vlNumPoliza, vlCodTipoIden, vlNumIden, vlNumEndoso)
    Txt_PenPoliza = vlNumPoliza
    Call fgBuscarPosicionCodigoCombo(vlCodTipoIden, Cmb_PenNumIdent)
    Txt_PenNumIdent = vlNumIden
    ''Txt_PenDigito = vlDigito
    Lbl_End = vlNumEndoso
    Cmd_BuscarPol_Click
End Function

Function flInicializaGrillaGrupo()

On Error GoTo Err_flInicializaGrillaGrupo
    
    Msf_GrillaGrupo.Clear
    Msf_GrillaGrupo.Cols = 6
    Msf_GrillaGrupo.rows = 1
    'Msf_GrillaGrupo.RowHeight(0) = 250
    Msf_GrillaGrupo.row = 0
    
    Msf_GrillaGrupo.Col = 0
    Msf_GrillaGrupo.Text = "Nº Orden"
    Msf_GrillaGrupo.ColWidth(0) = 700
    
    Msf_GrillaGrupo.Col = 1
    Msf_GrillaGrupo.Text = "Parentesco"
    Msf_GrillaGrupo.ColWidth(1) = 3500
    
    Msf_GrillaGrupo.Col = 2
    Msf_GrillaGrupo.Text = "Tipo Ident."
    Msf_GrillaGrupo.ColWidth(2) = 1000
    
    Msf_GrillaGrupo.Col = 3
    Msf_GrillaGrupo.Text = "Nº Ident."
    Msf_GrillaGrupo.ColWidth(3) = 1000

    Msf_GrillaGrupo.Col = 4
    Msf_GrillaGrupo.Text = "Nombre"
    Msf_GrillaGrupo.ColWidth(4) = 2500
    
    Msf_GrillaGrupo.Col = 5
    Msf_GrillaGrupo.Text = "Estado"
    Msf_GrillaGrupo.ColWidth(5) = 1000
    
Exit Function
Err_flInicializaGrillaGrupo:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function flInicializaGrillaRetJud()

On Error GoTo Err_flInicializaGrillaRetJud

    Msf_GrillaRetJud.Clear
    Msf_GrillaRetJud.Cols = 6
    Msf_GrillaRetJud.rows = 1
    Msf_GrillaRetJud.RowHeight(0) = 250
    Msf_GrillaRetJud.row = 0
    
    Msf_GrillaRetJud.Col = 0
    Msf_GrillaRetJud.Text = "P. Vigencia"
    Msf_GrillaRetJud.ColWidth(0) = 1700
    
    Msf_GrillaRetJud.Col = 1
    Msf_GrillaRetJud.Text = "Tipo Ret."
    Msf_GrillaRetJud.ColWidth(1) = 1800
    
    Msf_GrillaRetJud.Col = 2
    Msf_GrillaRetJud.Text = "Rut"
    Msf_GrillaRetJud.ColWidth(2) = 1000

    Msf_GrillaRetJud.Col = 3
    Msf_GrillaRetJud.Text = "Nombre"
    Msf_GrillaRetJud.ColWidth(3) = 3000
    
    Msf_GrillaRetJud.Col = 4
    Msf_GrillaRetJud.Text = "Modalidad"
    Msf_GrillaRetJud.ColWidth(4) = 1500

    Msf_GrillaRetJud.Col = 5
    Msf_GrillaRetJud.Text = "Fecha Efecto"
    Msf_GrillaRetJud.ColWidth(5) = 1000


Exit Function
Err_flInicializaGrillaRetJud:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function flInicializaGrillaHabDes()

On Error GoTo Err_flInicializaGrillaHabDes

    Msf_GrillaHabDes.Clear
    Msf_GrillaHabDes.Cols = 4
    Msf_GrillaHabDes.rows = 1
    Msf_GrillaHabDes.RowHeight(0) = 250
    Msf_GrillaHabDes.row = 0
    
    Msf_GrillaHabDes.Col = 0
    Msf_GrillaHabDes.Text = "Periodo"
    Msf_GrillaHabDes.ColWidth(0) = 2100
    
    Msf_GrillaHabDes.Col = 1
    Msf_GrillaHabDes.Text = "Código"
    Msf_GrillaHabDes.ColWidth(1) = 800
    
    Msf_GrillaHabDes.Col = 2
    Msf_GrillaHabDes.Text = "Concepto"
    Msf_GrillaHabDes.ColWidth(2) = 4000

    Msf_GrillaHabDes.Col = 3
    Msf_GrillaHabDes.Text = "Monto"
    Msf_GrillaHabDes.ColWidth(3) = 1500

Exit Function
Err_flInicializaGrillaHabDes:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function flInicializaGrillaCertificados()

On Error GoTo Err_flInicializaGrillaCertificados

Exit Function
Err_flInicializaGrillaCertificados:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function flCargaGrillaGrupo()
Dim vlTipoI As String
On Error GoTo Err_flCargaGrillaGrupo

    vgSql = ""
    vgSql = "SELECT num_orden,cod_par,cod_tipoidenben,num_idenben, "
    vgSql = vgSql & "gls_nomben,gls_nomsegben,gls_patben,gls_matben,cod_estpension "
    vgSql = vgSql & "FROM PP_TMAE_BEN "
    vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
    vgSql = vgSql & "num_endoso = " & str(vlNumEndoso) & " "
    vgSql = vgSql & "ORDER by num_orden "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       Call flInicializaGrillaGrupo
       
       While Not vgRegistro.EOF
          
          vlTipoI = Trim(vgRegistro!Cod_TipoIdenBen) & " - " & fgBuscarNombreTipoIden(Trim(vgRegistro!Cod_TipoIdenBen), False)
          vlCodPar = " " & Trim(vgRegistro!Cod_Par) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_Par, Trim(vgRegistro!Cod_Par)))
          
          Msf_GrillaGrupo.AddItem (vgRegistro!Num_Orden) & vbTab _
          & (vlCodPar) & vbTab _
          & vlTipoI & vbTab & (Trim(vgRegistro!Num_IdenBen)) & vbTab _
          & (Trim(vgRegistro!Gls_NomBen)) & " " & (Trim(vgRegistro!Gls_PatBen)) & " " & (Trim(vgRegistro!Gls_MatBen)) & vbTab _
          & (Trim(vgRegistro!Cod_EstPension))
          
          vgRegistro.MoveNext
       Wend
    End If
    vgRegistro.Close

Exit Function
Err_flCargaGrillaGrupo:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
    
End Function

Function flCargaGrillaRetJud(numero As Integer)
Dim vlTipoI As String
On Error GoTo Err_flCargaGrillaRetJud

    vgSql = ""
    vgSql = "SELECT fec_iniret,fec_terret,cod_tipret,cod_modret, "
    vgSql = vgSql & "cod_tipoidenreceptor,num_idenreceptor,gls_nomreceptor, "
    vgSql = vgSql & "gls_patreceptor,gls_matreceptor,fec_efecto "
    vgSql = vgSql & "FROM PP_TMAE_RETJUDICIAL "
    vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
    vgSql = vgSql & "num_orden = " & numero & " AND "
    vgSql = vgSql & "fec_terret = '" & clFechaTopeTer & "' "
    vgSql = vgSql & "ORDER by fec_iniret "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       Call flInicializaGrillaRetJud
           
       While Not vgRegistro.EOF
       
             vlTipoI = Trim(vgRegistro!Cod_TipoIdenReceptor) & " - " & fgBuscarNombreTipoIden(Trim(vgRegistro!Cod_TipoIdenReceptor), False)
             vlTipoRet = Trim(vgRegistro!cod_tipret) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_TipRetJud, Trim(vgRegistro!cod_tipret)))
             vlModRet = Trim(vgRegistro!cod_modret) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_ModPagoRetJud, Trim(vgRegistro!cod_modret)))
             
             vlFechaIni = DateSerial(Mid((vgRegistro!FEC_INIRET), 1, 4), Mid((vgRegistro!FEC_INIRET), 5, 2), Mid((vgRegistro!FEC_INIRET), 7, 2))
             vlFechaTer = DateSerial(Mid((vgRegistro!FEC_TERRET), 1, 4), Mid((vgRegistro!FEC_TERRET), 5, 2), Mid((vgRegistro!FEC_TERRET), 7, 2))
             vlFecha = Trim(vlFechaIni) & " - " & Trim(vlFechaTer)
       
             Msf_GrillaRetJud.AddItem Trim(vlFecha) & vbTab _
             & (vlTipoRet) & vbTab _
             & (vlTipoI) & vbTab & (Trim(vgRegistro!Num_IdenReceptor)) & vbTab _
             & (Trim(vgRegistro!Gls_NomReceptor)) & " " & (Trim(vgRegistro!Gls_PatReceptor)) & " " & (Trim(vgRegistro!Gls_MatReceptor)) & vbTab _
             & (vlModRet) & vbTab _
             & (DateSerial(Mid((vgRegistro!FEC_EFECTO), 1, 4), Mid((vgRegistro!FEC_EFECTO), 5, 2), Mid((vgRegistro!FEC_EFECTO), 7, 2)))
                          
             vgRegistro.MoveNext
       Wend
    End If
    vgRegistro.Close

Exit Function
Err_flCargaGrillaRetJud:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
    
End Function
'''
'''Function flCargaGrillaAsigFam(numero As Integer)
'''
'''On Error GoTo Err_flCargaGrillaAsigFam
'''    'Obtener Valor de Monto de la Carga
'''
'''
'''    vlMtoCarga = 0
'''    vgSql = ""
'''    vgSql = "SELECT mto_carga "
'''    vgSql = vgSql & "FROM PP_TMAE_VALCARFAM "
'''    vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
'''    vgSql = vgSql & "num_orden = " & Str(numero) & " "
'''    vgSql = vgSql & "ORDER BY num_annodecing DESC "
'''    Set vgRegistro = vgConexionBD.Execute(vgSql)
'''    If Not vgRegistro.EOF Then
'''       vlMtoCarga = (vgRegistro!Mto_Carga)
'''    End If
'''
''''''    Lbl_AFMontoCarga = Format(vlMtoCarga, "###,###,##0.00")
'''
'''    'Obtener Número de Cargas ACTIVAS
'''    vlNumCargas = 0
'''    vgSql = ""
'''    vgSql = "SELECT COUNT (DISTINCT(num_orden)) as numero "
'''    vgSql = vgSql & "FROM PP_TMAE_ASIGFAM "
'''    vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
'''    vgSql = vgSql & "num_ordenrec = " & vlNumOrden & " AND "
'''    vgSql = vgSql & "cod_estvigencia = '" & clCargaActiva & "' "
'''    vgSql = vgSql & "ORDER by fec_iniactiva DESC "
'''    Set vlRegistro2 = vgConexionBD.Execute(vgSql)
'''    If Not vlRegistro2.EOF Then
'''        vlNumCargas = (vlRegistro2!numero)
'''    End If
'''
'''    'Obtener todos los beneficiarios que en alguna oportunidad han sido
'''    'cargas familiares activas
'''    vlEstCerEst = "N"
'''    vgSql = ""
'''    vgSql = "SELECT DISTINCT (num_orden) "
'''    vgSql = vgSql & "FROM PP_TMAE_ASIGFAM "
'''    vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
'''    vgSql = vgSql & "num_ordenrec = " & vlNumOrden & " "
'''    vgSql = vgSql & "ORDER by num_orden ASC "
'''    Set vlRegistro2 = vgConexionBD.Execute(vgSql)
'''    If Not vlRegistro2.EOF Then
''''''       Call flInicializaGrillaAsigFam
'''       While Not vlRegistro2.EOF
'''
'''           'Obtener datos de la carga ben o no ben
'''               If (vlRegistro2!Num_Orden) >= 50 Then
'''
'''                    vgSql = ""
'''                    vgSql = "SELECT cod_ascdes,rut_ben,dgv_ben,num_orden, "
'''                    vgSql = vgSql & "gls_nomben,gls_patben,gls_matben "
'''                    vgSql = vgSql & "FROM PP_TMAE_NOBEN "
'''                    vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
''''                    vgSql = vgSql & "num_endoso = " & Str(vlNumEndoso) & " AND "
'''                    vgSql = vgSql & "num_orden = " & (vlRegistro2!Num_Orden) & " "
'''                    Set vgRegistro = vgConexionBD.Execute(vgSql)
'''                    If Not vgRegistro.EOF Then
'''                       vlCodPar = " " & Trim(vgRegistro!COD_ASCDES) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_ParNoBen, Trim(vgRegistro!COD_ASCDES)))
'''                    End If
'''               Else
'''                   vgSql = ""
'''                   vgSql = "SELECT cod_par,rut_ben,dgv_ben,num_orden, "
'''                   vgSql = vgSql & "gls_nomben,gls_patben,gls_matben "
'''                   vgSql = vgSql & "FROM PP_TMAE_BEN "
'''                   vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
'''                   vgSql = vgSql & "num_endoso = " & Str(vlNumEndoso) & " AND "
'''                   vgSql = vgSql & "num_orden = " & (vlRegistro2!Num_Orden) & " "
'''                   Set vgRegistro = vgConexionBD.Execute(vgSql)
'''                   If Not vgRegistro.EOF Then
'''                      vlCodPar = " " & Trim(vgRegistro!Cod_Par) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_Par, Trim(vgRegistro!Cod_Par)))
'''                   End If
'''               End If
'''
'''          'Obtener Estado de Certificado de Estudio a la fecha
'''            'vlEstCerEst = ""
'''            vgSql = ""
'''            vgSql = "SELECT fec_inicerest,fec_tercerest "
'''            vgSql = vgSql & "FROM PP_TMAE_CERESTUDIO "
'''            vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
'''            vgSql = vgSql & "num_orden = " & Str(vlRegistro2!Num_Orden) & " "
'''            vgSql = vgSql & "ORDER BY fec_inicerest DESC "
'''            Set vgRs = vgConexionBD.Execute(vgSql)
'''            If Not vgRs.EOF Then
'''               vlFechaActual = fgBuscaFecServ
'''               vlFechaActual = Format(CDate(Trim(vlFechaActual)), "yyyymmdd")
'''               If vlFechaActual >= (vgRs!Fec_IniCerEst) And _
'''                  vlFechaActual <= (vgRs!FEC_TERCEREST) Then
'''                  vlEstCerEst = "S"
'''               Else
'''                   vlEstCerEst = "N"
'''               End If
'''            End If
'''
'''            'Obtener los datos de la última vez que la carga
'''            'estuvo activa
'''            vgSql = ""
'''            vgSql = "SELECT MAX(fec_iniactiva) ,fec_iniactiva,fec_teractiva, "
'''            vgSql = vgSql & "cod_caususpension,cod_estvigencia "
'''            vgSql = vgSql & "FROM PP_TMAE_ASIGFAM "
'''            vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
'''            vgSql = vgSql & "num_orden = " & Str(vlRegistro2!Num_Orden) & " "
'''            vgSql = vgSql & "GROUP BY fec_iniactiva,fec_teractiva, "
'''            vgSql = vgSql & "cod_caususpension,cod_estvigencia "
'''            Set vlRegistro3 = vgConexionBD.Execute(vgSql)
'''            If Not vlRegistro3.EOF Then
'''                vlCodCauSus = " " & Trim(vlRegistro3!COD_CAUSUSPENSION) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_CauSupAsiFam, Trim(vlRegistro3!COD_CAUSUSPENSION)))
'''                vlCodEstado = " " & Trim(vlRegistro3!COD_ESTVIGENCIA) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_EstVigAsiFam, Trim(vlRegistro3!COD_ESTVIGENCIA)))
'''                vlFechaIni = DateSerial(Mid((vlRegistro3!FEC_INIACTIVA), 1, 4), Mid((vlRegistro3!FEC_INIACTIVA), 5, 2), Mid((vlRegistro3!FEC_INIACTIVA), 7, 2))
'''                vlFechaTer = DateSerial(Mid((vlRegistro3!FEC_TERACTIVA), 1, 4), Mid((vlRegistro3!FEC_TERACTIVA), 5, 2), Mid((vlRegistro3!FEC_TERACTIVA), 7, 2))
'''                vlFecha = Trim(vlFechaIni) & " - " & Trim(vlFechaTer)
'''            End If
'''
'''          'vlCodPar = " " & Trim(vgRs2!COD_PAR) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_Par, Trim(vgRs2!COD_PAR)))
'''
'''          'vlFechaEfectoAsigFam = DateSerial(Mid((vgRs2!FEC_INIACTIVA), 1, 4), Mid((vgRs2!FEC_INIACTIVA), 5, 2), Mid((vgRs2!FEC_INIACTIVA), 7, 2))
'''          vlFechaEfectoAsigFam = fgValidaFechaEfecto(vlFechaIni, Trim(Txt_PenPoliza.Text), vlNumOrden)
'''
'''
'''          Msf_GrillaAsigFam.AddItem (vlCodPar) & vbTab _
'''          & (" " & Format((Trim(vgRegistro!Rut_Ben)), "##,###,##0") & " - " & (Trim(vgRegistro!Dgv_Ben))) & vbTab _
'''          & (Trim(vgRegistro!Gls_NomBen)) & " " & (Trim(vgRegistro!Gls_PatBen)) & " " & (Trim(vgRegistro!Gls_MatBen)) & vbTab _
'''          & (vlFecha) & vbTab _
'''          & (vlCodCauSus) & vbTab _
'''          & (vlCodEstado) & vbTab _
'''          & (vlEstCerEst) & vbTab _
'''          & (vlFechaEfectoAsigFam)
'''          'Falta Fecha de Efecto
'''
'''
''''          If Trim(vlRegistro3!COD_ESTVIGENCIA) = clCargaActiva Then
''''             vlNumCargas = vlNumCargas + 1
''''          End If
'''          vlRegistro2.MoveNext
'''       Wend
'''    End If
'''    vgRegistro.Close
'''
'''Exit Function
'''Err_flCargaGrillaAsigFam:
'''    Screen.MousePointer = 0
'''    Select Case Err
'''        Case Else
'''        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'''    End Select
'''
'''End Function

Function flCargaGrillaHabDes(numero As Integer)

On Error GoTo Err_flCargaGrillaHabDes
        
    vgSql = ""
    'mv 20190625
'    vgSql = vgSql & "SELECT p.cod_conhabdes,p.mto_conhabdes,p.fec_inipago,p.fec_terpago "
    vgSql = vgSql & "SELECT distinct p.cod_conhabdes,p.mto_conhabdes,p.fec_inipago,p.fec_terpago "
    vgSql = vgSql & "FROM PP_TMAE_LIQPAGOPENDEF l,PP_TMAE_PAGOPENDEF p "
    vgSql = vgSql & "WHERE l.num_poliza = '" & Trim(Txt_PenPoliza) & "' AND "
    vgSql = vgSql & "l.num_orden = " & numero & " AND "
    vgSql = vgSql & "l.fec_pago >= '" & vlFechaIni & "' AND "
    vgSql = vgSql & "l.fec_pago <= '" & vlFechaTer & "' AND "
    vgSql = vgSql & "l.num_perpago = p.num_perpago AND "
    vgSql = vgSql & "l.num_poliza = p.num_poliza AND "
    vgSql = vgSql & "l.num_orden = p.num_orden AND "
    ''*vgSql = vgSql & "l.rut_receptor = p.rut_receptor AND "
    vgSql = vgSql & "l.cod_tipoidenreceptor=p.cod_tipoidenreceptor AND "
    vgSql = vgSql & "l.num_idenreceptor = p.num_idenreceptor AND "
    vgSql = vgSql & "l.cod_tipreceptor = p.cod_tipreceptor "
'    vgSql = vgSql & "AND p.cod_conhabdes IN " & vlCodConceptos & " "
    vgSql = vgSql & vgQuery
    vgSql = vgSql & "ORDER BY p.fec_inipago DESC, p.cod_conhabdes ASC "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
        
       Call flInicializaGrillaHabDes
       
       While Not vgRegistro.EOF
                         
          vlNombreConcepto = ""
          Call flBuscaNombreConcepto(vgRegistro!Cod_ConHabDes)
          
          vlFechaIni = DateSerial(Mid((vgRegistro!Fec_IniPago), 1, 4), Mid((vgRegistro!Fec_IniPago), 5, 2), Mid((vgRegistro!Fec_IniPago), 7, 2))
          vlFechaTer = DateSerial(Mid((vgRegistro!Fec_TerPago), 1, 4), Mid((vgRegistro!Fec_TerPago), 5, 2), Mid((vgRegistro!Fec_TerPago), 7, 2))
          vlFecha = Trim(vlFechaIni) & " - " & Trim(vlFechaTer)
                 
          Msf_GrillaHabDes.AddItem (vlFecha) & vbTab _
          & (vgRegistro!Cod_ConHabDes) & vbTab _
          & (vlNombreConcepto) & vbTab _
          & (Format((vgRegistro!Mto_ConHabDes), "###,###,##0.00"))
          
          vgRegistro.MoveNext
       Wend
    Else
        Call flInicializaGrillaHabDes
        MsgBox "No Existen Detalles para los Conceptos Seleccionados", vbInformation, "Información"
    End If
    vgRegistro.Close
        
        
Exit Function
Err_flCargaGrillaHabDes:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
    
End Function

Function flCargaGrillaCertificados()

On Error GoTo Err_flCargaGrillaCertificados

Exit Function
Err_flCargaGrillaCertificados:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
    
End Function
'''
'''Function flCargaGrillaCCAF(numero As Integer)
'''Dim iFecha As String
'''Dim iAnno As Integer
'''Dim iMes As Integer
'''Dim iDia As Integer
'''Dim iFechaInicio As String
'''Dim iFechaTermino As String
'''
'''On Error GoTo Err_flCargaGrillaCCAF
'''
'''    vgSql = ""
'''    vgSql = "SELECT num_perpago "
'''    vgSql = vgSql & "FROM pp_tmae_propagopen "
'''    vgSql = vgSql & "WHERE "
'''    vgSql = vgSql & "cod_estadoreg = 'C' or "
'''    vgSql = vgSql & "cod_estadopri = 'C' "
'''    vgSql = vgSql & "ORDER BY num_perpago desc"
'''    Set vgRs2 = vgConexionBD.Execute(vgSql)
'''    If Not vgRs2.EOF Then
'''        iFecha = DateSerial(CInt(Mid(vgRs2!Num_PerPago, 1, 4)), CInt(Mid(vgRs2!Num_PerPago, 5, 2)), 1)
'''    Else
'''        iFecha = fgBuscaFecServ
'''    End If
'''    vgRs2.Close
'''
'''    vlFecha = fgValidaFechaEfecto(iFecha, Trim(Txt_PenPoliza), numero)
'''    vlFecha = Format(vlFecha, "yyyymmdd")
'''    vlFecha = Mid(vlFecha, 1, 6)
'''
'''    vlCodHabDesCCAF = ""
'''    Call flBuscarCodHabDesCCAF
'''
'''    vgSql = ""
'''    vgSql = "SELECT cod_conhabdes,cod_cajacompen,fec_inihabdes, "
'''    vgSql = vgSql & "fec_terhabdes,num_cuotas,mto_cuota, "
'''    vgSql = vgSql & "mto_total,cod_moneda "
'''    vgSql = vgSql & "FROM PP_TMAE_HABDESCCAF "
'''    vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
'''    vgSql = vgSql & "num_orden = " & numero & " AND "
'''    vgSql = vgSql & "cod_conhabdes IN " & vlCodHabDesCCAF & " AND "
'''    'vgSql = vgSql & "fec_inihabdes <= '" & Trim(vlFecha) & "' AND "
'''    If vgTipoBase = "ORACLE" Then
'''       vgSql = vgSql & "substr(fec_inihabdes,1,6) <= '" & Trim(vlFecha) & "' "
'''       vgSql = vgSql & " AND substr(fec_terhabdes,1,6) >= '" & Trim(vlFecha) & "' "
'''    Else
'''        vgSql = vgSql & "substring(fec_inihabdes,1,6) <= '" & Trim(vlFecha) & "' "
'''        vgSql = vgSql & " AND substring(fec_terhabdes,1,6) >= '" & Trim(vlFecha) & "' "
'''    End If
'''    'vgSql = vgSql & "fec_terhabdes >= '" & Trim(vlFecha) & "' "
'''    vgSql = vgSql & "ORDER BY cod_conhabdes,fec_inihabdes DESC "
'''    Set vgRegistro = vgConexionBD.Execute(vgSql)
'''    If Not vgRegistro.EOF Then
'''       Call flInicializaGrillaCCAF
'''       While Not vgRegistro.EOF
'''             vlNombreConcepto = ""
'''             Call flBuscaNombreConcepto(vgRegistro!Cod_ConHabDes)
'''             vlNombreCCAF = " " & Trim(vgRegistro!Cod_CajaCompen) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_CCAF, Trim(vgRegistro!Cod_CajaCompen)))
'''             vlFechaIni = DateSerial(Mid((vgRegistro!Fec_IniHabDes), 1, 4), Mid((vgRegistro!Fec_IniHabDes), 5, 2), Mid((vgRegistro!Fec_IniHabDes), 7, 2))
'''             vlFechaTer = DateSerial(Mid((vgRegistro!FEC_TERHabDes), 1, 4), Mid((vgRegistro!FEC_TERHabDes), 5, 2), Mid((vgRegistro!FEC_TERHabDes), 7, 2))
'''             vlFecha = Trim(vlFechaIni) & " - " & Trim(vlFechaTer)
'''
'''             Msf_GrillaCCAF.AddItem (" " & Trim(vgRegistro!Cod_ConHabDes) & " - " & Trim(vlNombreConcepto)) & vbTab _
'''             & ((vlNombreCCAF)) & vbTab _
'''             & (Trim(vlFecha)) & vbTab _
'''             & (Format((vgRegistro!Num_Cuotas), "###,###,##0.00")) & vbTab _
'''             & (Format((vgRegistro!MTO_CUOTA), "###,###,##0.00")) & vbTab _
'''             & (Format((vgRegistro!mto_total), "###,###,##0.00")) & vbTab _
'''             & (Trim(vgRegistro!Cod_Moneda))
'''
'''             vgRegistro.MoveNext
'''       Wend
'''    End If
'''    vgRegistro.Close
'''
'''Exit Function
'''Err_flCargaGrillaCCAF:
'''    Screen.MousePointer = 0
'''    Select Case Err
'''        Case Else
'''        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'''    End Select
'''
'''End Function


'------------------------------------------------
Function flMostrarDatosPoliza()


On Error GoTo Err_flMostrarDatosPoliza
    Dim vlTB2 As ADODB.Recordset
    vgSql = ""
    vgSql = "SELECT fec_vigencia,fec_tervigencia,cod_tippension, "
    vgSql = vgSql & "cod_tipren,cod_modalidad,num_cargas,cod_afp,prc_tasace, "
    vgSql = vgSql & "prc_tasactorea,mto_prima,cod_estado,num_mesdif, "
    vgSql = vgSql & "num_mesgar,prc_tasavta,prc_tasaintpergar,mto_pension,fec_dev, cod_moneda, fec_devsol "
    vgSql = vgSql & "FROM PP_TMAE_POLIZA "
    vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
    vgSql = vgSql & "num_endoso = '" & vlNumEndoso & "' "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       If IsNull(vgRegistro!Fec_Vigencia) Then
          Lbl_PolIniVig = ""
       Else
           Lbl_PolIniVig = DateSerial(Mid((vgRegistro!Fec_Vigencia), 1, 4), Mid((vgRegistro!Fec_Vigencia), 5, 2), Mid((vgRegistro!Fec_Vigencia), 7, 2))
       End If
       If IsNull(vgRegistro!Fec_TerVigencia) Then
          Lbl_PolTerVig = ""
       Else
           Lbl_PolTerVig = DateSerial(Mid((vgRegistro!Fec_TerVigencia), 1, 4), Mid((vgRegistro!Fec_TerVigencia), 5, 2), Mid((vgRegistro!Fec_TerVigencia), 7, 2))
       End If
       If IsNull(vgRegistro!fec_dev) Then
          Lbl_PolFecDev = ""
       Else
           Lbl_PolFecDev = DateSerial(Mid((vgRegistro!fec_dev), 1, 4), Mid((vgRegistro!fec_dev), 5, 2), Mid((vgRegistro!fec_dev), 7, 2))
       End If
       
       If IsNull(vgRegistro!FEC_DEVSOL) Then
           lblFecSolDev = ""
       Else
           lblFecSolDev = DateSerial(Mid((vgRegistro!FEC_DEVSOL), 1, 4), Mid((vgRegistro!FEC_DEVSOL), 5, 2), Mid((vgRegistro!FEC_DEVSOL), 7, 2))
       End If
       
       Lbl_PolNumEndoso = vlNumEndoso
       If IsNull(vgRegistro!Cod_TipPension) Then
          Lbl_PolTipPen = ""
       Else
           
           Lbl_PolTipPen = Trim(vgRegistro!Cod_TipPension) + " - " + fgBuscarGlosaElemento(vgCodTabla_TipPen, Trim(vgRegistro!Cod_TipPension))
       End If
       If IsNull(vgRegistro!Cod_TipRen) Then
          Lbl_PolTipRta = ""
       Else
           Lbl_PolTipRta = Trim(vgRegistro!Cod_TipRen) + " - " + fgBuscarGlosaElemento(vgCodTabla_TipRen, Trim(vgRegistro!Cod_TipRen))
       End If
       If IsNull(vgRegistro!Cod_Modalidad) Then
          Lbl_PolMod = ""
       Else
           Lbl_PolMod = Trim(vgRegistro!Cod_Modalidad) + " - " + fgBuscarGlosaElemento(vgCodTabla_AltPen, Trim(vgRegistro!Cod_Modalidad))
       End If
       If IsNull(vgRegistro!Num_Cargas) Then
          Lbl_PolNumCar = ""
       Else
           Lbl_PolNumCar = (vgRegistro!Num_Cargas)
       End If
       If IsNull(vgRegistro!cod_afp) Then
          Lbl_PolAfp = ""
       Else
           Lbl_PolAfp = Trim(vgRegistro!cod_afp) + " - " + fgBuscarGlosaElemento(vgCodTabla_AFP, Trim(vgRegistro!cod_afp))
       End If
       If IsNull(vgRegistro!Prc_TasaCe) Then
          Lbl_PolTasaCto = ""
       Else
           Lbl_PolTasaCto = Format((vgRegistro!Prc_TasaCe), "#,###,##0.00")
       End If
       If IsNull(vgRegistro!prc_tasactorea) Then
          Lbl_PolTasaRea = ""
       Else
           Lbl_PolTasaRea = Format((vgRegistro!prc_tasactorea), "#,###,##0.00")
       End If
       If IsNull(vgRegistro!Mto_Prima) Then
          Lbl_PolMtoPri = ""
       Else
           Lbl_PolMtoPri = Format((vgRegistro!Mto_Prima), "###,###,##0.00")
       End If
       If IsNull(vgRegistro!Cod_Estado) Then
          Lbl_PolEstado = ""
       Else
           Lbl_PolEstado = Trim(vgRegistro!Cod_Estado) + " - " + fgBuscarGlosaElemento(vgCodTabla_TipVigPol, Trim(vgRegistro!Cod_Estado))
       End If
       If IsNull(vgRegistro!Num_MesDif) Then
          Lbl_PolMesDif = ""
       Else
           Lbl_PolMesDif = (vgRegistro!Num_MesDif)
       End If
       If IsNull(vgRegistro!Num_MesGar) Then
          Lbl_PolMesGar = ""
       Else
           Lbl_PolMesGar = (vgRegistro!Num_MesGar)
       End If
       If IsNull(vgRegistro!Prc_TasaVta) Then
          Lbl_PolTasaVta = ""
       Else
           Lbl_PolTasaVta = Format((vgRegistro!Prc_TasaVta), "#,###,##0.00")
       End If
       If IsNull(vgRegistro!Prc_TasaIntPerGar) Then
          Lbl_PolTasaPerGar = ""
       Else
           Lbl_PolTasaPerGar = Format((vgRegistro!Prc_TasaIntPerGar), "#,###,##0.00")
       End If
       If IsNull(vgRegistro!Mto_Pension) Then
          Lbl_PolMtoPen = ""
       Else
           Lbl_PolMtoPen = Format((vgRegistro!Mto_Pension), "###,###,##0.00")
       End If
       
        'hqr 20/10/2007 Obtiene Monto de la Pensión Actualizada
        Lbl_Moneda(0) = vgRegistro!Cod_Moneda
        Lbl_Moneda(1) = vgRegistro!Cod_Moneda
        
        'Obtiene Pension Actualizada
        vgSql = "SELECT mto_pension FROM pp_tmae_pensionact a"
        vgSql = vgSql & " WHERE a.num_poliza = '" & Trim(Txt_PenPoliza.Text) & "'"
        vgSql = vgSql & " AND a.num_endoso = " & vlNumEndoso
        vgSql = vgSql & " AND a.fec_desde = "
            vgSql = vgSql & " (SELECT max(fec_desde) FROM pp_tmae_pensionact b"
            vgSql = vgSql & " WHERE b.num_poliza = a.num_poliza"
            vgSql = vgSql & " AND b.num_endoso = a.num_endoso"
            vgSql = vgSql & " AND b.fec_desde <= TO_CHAR(SYSDATE,'yyyymmdd'))"
        Set vlTB2 = vgConexionBD.Execute(vgSql)
        If Not vlTB2.EOF Then
            If Not IsNull(vlTB2!Mto_Pension) Then
                Lbl_PolMtoPen = Format((vlTB2!Mto_Pension), "###,###,##0.00")
            End If
        End If
    End If

Exit Function
Err_flMostrarDatosPoliza:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
    
End Function

Function flMostrarDatosGrupo(numero As Integer)
Dim vlTipoI As String
On Error GoTo Err_flMostrarDatosGrupo

    vgSql = ""
    vgSql = "SELECT cod_par,cod_tipoidenben,num_idenben,gls_nomben,gls_nomsegben,gls_patben, "
    vgSql = vgSql & "gls_matben,gls_dirben,cod_direccion,gls_fonoben,gls_correoben, "
    vgSql = vgSql & "mto_pension,num_orden,cod_sitinv,fec_nacben,fec_fallben, cod_derpen,"
    vgSql = vgSql & "fec_ingreso,cod_estpension, prc_pension, prc_pensiongar, fec_terpagopengar, mto_pensiongar, TO_CHAR(SYSDATE,'yyyymmdd') as Fecha_Actual "
    'RRR
    vgSql = vgSql & ", gls_telben2 "
    'mvg 20170904
    vgSql = vgSql & ", ind_bolelec "
    vgSql = vgSql & "FROM PP_TMAE_BEN "
    vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
    vgSql = vgSql & "num_endoso = (select max(num_endoso) from PP_TMAE_BEN  where num_poliza='" & Trim(Txt_PenPoliza.Text) & "') AND "
    vgSql = vgSql & "num_orden = '" & numero & "' "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       If IsNull(vgRegistro!Cod_Par) Then
          Lbl_GrupPar = ""
       Else
           Lbl_GrupPar = Trim(vgRegistro!Cod_Par) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_Par, Trim(vgRegistro!Cod_Par)))
       End If
       If IsNull(vgRegistro!Cod_TipoIdenBen) Then
          Lbl_GrupTipoIdent = ""
       Else
           vlTipoI = Trim(vgRegistro!Cod_TipoIdenBen) & " - " & fgBuscarNombreTipoIden(Trim(vgRegistro!Cod_TipoIdenBen), False)
           Lbl_GrupTipoIdent = vlTipoI
       End If
       If IsNull(vgRegistro!Num_IdenBen) Then
          Lbl_GrupNumIdent = ""
       Else
           Lbl_GrupNumIdent = Trim(vgRegistro!Num_IdenBen)
       End If
       If IsNull(vgRegistro!Gls_NomBen) Then
          Lbl_GrupNombre = ""
       Else
           Lbl_GrupNombre = Trim(vgRegistro!Gls_NomBen)
       End If
       If IsNull(vgRegistro!Gls_NomSegBen) Then
          Lbl_GrupNombreSeg = ""
       Else
           Lbl_GrupNombreSeg = Trim(vgRegistro!Gls_NomSegBen)
       End If
       If IsNull(vgRegistro!Gls_PatBen) Then
          Lbl_GrupPaterno = ""
       Else
           Lbl_GrupPaterno = Trim(vgRegistro!Gls_PatBen)
       End If
       If IsNull(vgRegistro!Gls_MatBen) Then
          Lbl_GrupMaterno = ""
       Else
           Lbl_GrupMaterno = Trim(vgRegistro!Gls_MatBen)
       End If
       If IsNull(vgRegistro!Gls_DirBen) Then
          Lbl_GrupDomicilio = ""
       Else
           Lbl_GrupDomicilio = Trim(vgRegistro!Gls_DirBen)
       End If
       vlNombreComuna = ""
       Call flBuscaNombreComuna(vgRegistro!Cod_Direccion)
       Lbl_GrupComuna = vlNombreComuna
       Call fgBuscarNombreProvinciaRegion(vgRegistro!Cod_Direccion)
       Lbl_GrupProvincia = vgNombreProvincia
       Lbl_GrupRegion = vgNombreRegion
       If IsNull(vgRegistro!Gls_FonoBen) Then
          Lbl_GrupFono = ""
       Else
          Lbl_GrupFono = (vgRegistro!Gls_FonoBen)
       End If
       'RRR 08/05/213
       If IsNull(vgRegistro!Gls_Telben2) Then
          lbl_GrupoFono2 = ""
       Else
          lbl_GrupoFono2 = (vgRegistro!Gls_Telben2)
       End If
       
       If IsNull(vgRegistro!Mto_Pension) Then
          Lbl_GrupPension = ""
       Else
          Lbl_GrupPension = Format((vgRegistro!Mto_Pension), "###,###,##0.00")
       End If
       
       'hqr 20/10/2007 Se muestra la pensión actualizada
       If vgRegistro!Cod_DerPen <> clCodSinDerPen Then
           Lbl_GrupPension = Format(Lbl_PolMtoPen * (vgRegistro!Prc_Pension / 100), "#,#0.00")
           If Not IsNull(vgRegistro!Fec_TerPagoPenGar) Then
               If vgRegistro!Fec_TerPagoPenGar >= vgRegistro!Fecha_Actual Then
                   Lbl_GrupPension = Format(Lbl_PolMtoPen * (vgRegistro!Prc_PensionGar / 100), "#,#0.00")
               End If
           End If
       Else
           Lbl_GrupPension = Format(vgRegistro!Mto_Pension, "#,#0.00")
       End If
       'fin hqr 20/10/2007
       
       Lbl_GrupNumOrden = (vgRegistro!Num_Orden)
       If IsNull(vgRegistro!Cod_SitInv) Then
          Lbl_GrupSitInv = ""
       Else
           Lbl_GrupSitInv = (vgRegistro!Cod_SitInv)
       End If
       If IsNull(vgRegistro!Fec_NacBen) Then
          Lbl_GrupFecNac = ""
       Else
           Lbl_GrupFecNac = DateSerial(Mid((vgRegistro!Fec_NacBen), 1, 4), Mid((vgRegistro!Fec_NacBen), 5, 2), Mid((vgRegistro!Fec_NacBen), 7, 2))
       End If
       If IsNull(vgRegistro!Fec_Ingreso) Then
          Lbl_GrupFecIngreso = ""
       Else
           Lbl_GrupFecIngreso = DateSerial(Mid((vgRegistro!Fec_Ingreso), 1, 4), Mid((vgRegistro!Fec_Ingreso), 5, 2), Mid((vgRegistro!Fec_Ingreso), 7, 2))
       End If
       If IsNull(vgRegistro!Gls_CorreoBen) Then
          Lbl_GrupMail = ""
       Else
           Lbl_GrupMail = (vgRegistro!Gls_CorreoBen)
       End If
       If IsNull(vgRegistro!Fec_FallBen) Then
          Lbl_GrupFecFall = ""
       Else
           Lbl_GrupFecFall = DateSerial(Mid((vgRegistro!Fec_FallBen), 1, 4), Mid((vgRegistro!Fec_FallBen), 5, 2), Mid((vgRegistro!Fec_FallBen), 7, 2))
       End If
       If IsNull(vgRegistro!Cod_EstPension) Then
          Lbl_GrupEstado = ""
       Else
           Lbl_GrupEstado = (vgRegistro!Cod_EstPension)
       End If
       
       'mvg 20170904
       chkBolElec.Value = IIf(vgRegistro!ind_bolelec = "S", 1, 0)
       'CMV-20061031 I
       'Mostrar Monto de Pensión en Quiebra
       vlNumOrd = numero
       vlUltimoPerPago = flUltimoPeriodoCerrado(Trim(Txt_PenPoliza.Text))
       If fgObtieneParametrosQuiebra(vlUltimoPerPago, vlPrcCastigoQui, vlTopeMaxQui) Then
           Lbl_MtoPQ.Visible = True
           Lbl_MtoPensionQui.Visible = True
           Lbl_MtoPensionQui = flBuscarPensionQuiebra(vlUltimoPerPago, Trim(Txt_PenPoliza.Text), vlNumOrd, clCodTipReceptorR)
           Lbl_MtoPensionQui = Format(Lbl_MtoPensionQui, "#,#0.00")
       End If
       'CMV-20061031 F
       
       
    End If

Exit Function
Err_flMostrarDatosGrupo:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
    
End Function

Function flMostrarDatosPensiones(numero As Integer)

On Error GoTo Err_flMostrarDatosPensiones

    vgSql = ""
    vgSql = "SELECT cod_viapago,cod_sucursal,cod_tipcuenta,cod_banco, "
    vgSql = vgSql & "num_cuenta,cod_inssalud,cod_modsalud,mto_plansalud, cod_monbco, num_cuenta_cci, num_ctabco, COD_TIPCTA "
    ''*vgSql = vgSql & ",cod_modsalud2,mto_plansalud2 " no existen estos campos
    vgSql = vgSql & "FROM PP_TMAE_BEN "
    vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
    vgSql = vgSql & "num_endoso = " & vlNumEndoso & " AND "
    vgSql = vgSql & "num_orden = " & numero & " "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       vlCodInsSalud = ""
       If IsNull(vgRegistro!Cod_ViaPago) Then
          Lbl_FPViaPago = ""
       Else
           Lbl_FPViaPago = Trim(vgRegistro!Cod_ViaPago) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_ViaPago, Trim(vgRegistro!Cod_ViaPago)))
       End If
       If IsNull(vgRegistro!Cod_Sucursal) Then
          Lbl_FPSucursal = ""
       Else
           vlNombreSucursal = ""
           Call flBuscaNombreSucursal(vgRegistro!Cod_Sucursal)
           Lbl_FPSucursal = Trim(vlNombreSucursal)
       End If
       If IsNull(vgRegistro!cod_tipcta) Then
          Lbl_FPTipoCta = "00 - SIN INFORMACION"
       Else
           Lbl_FPTipoCta = Trim(vgRegistro!cod_tipcta) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_TipCta, Trim(vgRegistro!cod_tipcta)))
       End If
       If IsNull(vgRegistro!Cod_Banco) Then
          Lbl_FPBanco = "00 - SIN INFORMACION"
       Else
           Lbl_FPBanco = Trim(vgRegistro!Cod_Banco) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_Bco, Trim(vgRegistro!Cod_Banco)))
       End If
       
       If IsNull(vgRegistro!cod_monbco) Then
          Lbl_MonedaCta = "00 - SIN INFORMACION"
       Else
          Lbl_MonedaCta = Trim(vgRegistro!cod_monbco) & " - " & Trim(fgBuscarGlosaElemento("MP", Trim(vgRegistro!cod_monbco))) '(vgRegistro!cod_monbco)
       End If
       If IsNull(vgRegistro!num_ctabco) Then
          Lbl_FPNumCta = ""
       Else
           Lbl_FPNumCta = (vgRegistro!num_ctabco)
       End If
       If IsNull(vgRegistro!NUM_CUENTA_CCI) Then
          Lbl_FPNumCtaCCI = ""
       Else
          Lbl_FPNumCtaCCI = (vgRegistro!NUM_CUENTA_CCI)
       End If

       If IsNull(vgRegistro!Cod_InsSalud) Then
          Lbl_PSInstitucion = ""
       Else
           Lbl_PSInstitucion = Trim(vgRegistro!Cod_InsSalud) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_InsSal, Trim(vgRegistro!Cod_InsSalud)))
           vlCodInsSalud = Trim(vgRegistro!Cod_InsSalud)
       End If
       If IsNull(vgRegistro!Cod_ModSalud) Then
          Lbl_PSModPago = ""
       Else
           Lbl_PSModPago = Trim(vgRegistro!Cod_ModSalud) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_ModPago, Trim(vgRegistro!Cod_ModSalud)))
       End If
       
       
       
       
''*       If IsNull(vgRegistro!Cod_ModSalud2) Then
''          Lbl_PSModPago2 = ""
''       Else
''           Lbl_PSModPago2 = Trim(vgRegistro!Cod_ModSalud2) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_ModPago, Trim(vgRegistro!Cod_ModSalud2)))
''       End If
       If IsNull(vgRegistro!Mto_PlanSalud) Then
          Lbl_PSMtoPago = ""
       Else
           Lbl_PSMtoPago = Format((vgRegistro!Mto_PlanSalud), "###,###,##0.00")
       End If
''*       If IsNull(vgRegistro!Mto_PlanSalud2) Then
''          Lbl_PSMtoPago2 = ""
''       Else
''           Lbl_PSMtoPago2 = Format((vgRegistro!Mto_PlanSalud2), "###,###,##0.00")
''       End If
       Call flBuscaReceptor(numero)
       Call flBuscaFecEfectoInsSalud(numero)
       If vlFechaEfectoSalud = "" Then
       'Si no encuentra ningún registro de pago en liquidaciones, se calcula fecha
       'efecto según fecha actual de consulta
          vlFechaEfectoSalud = fgBuscaFecServ
          vlFechaEfectoSalud = fgValidaFechaEfecto(Trim(vlFechaEfectoSalud), Txt_PenPoliza.Text, numero)
       End If
       Lbl_PSFechaEfecto = Trim(vlFechaEfectoSalud)
       Lbl_RecRut = Trim(vlCodTipoIdenBenTut) & " - " & fgBuscarNombreTipoIden(vlCodTipoIdenBenTut, False)
        ''*vlRutRec
       Lbl_RecDgv = vlNumIdenBenTut ''*vlDgvRec
       Lbl_RecFecIniVig = vlFechaIni
       Lbl_RecFecTerVig = vlFechaTer
       Lbl_RecNombre = vlNomRec
       
       Call flCargaGrillaRetJud(numero)
       
    End If
    
Exit Function
Err_flMostrarDatosPensiones:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
    
End Function

'''Function flMostrarDatosOtrosBeneficios(numero As Integer)
'''
'''On Error GoTo Err_flMostrarDatosOtrosBeneficios
'''
''''Obtener datos de la última resolución de GE
'''    vgSql = ""
'''    vgSql = "SELECT MAX(fec_inires) as fec_inires "
'''    vgSql = vgSql & "FROM PP_TMAE_GARESTRES "
'''    vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
'''    vgSql = vgSql & "num_orden = " & numero & " "
'''    Set vgRegistro = vgConexionBD.Execute(vgSql)
'''    If Not vgRegistro.EOF Then
'''        If Not IsNull(vgRegistro!FEC_INIRES) Then
'''            vlFecha = (vgRegistro!FEC_INIRES)
'''            vgSql = ""
'''            vgSql = "SELECT num_resgarest,num_annores,cod_tipres,prc_deduccion "
'''            vgSql = vgSql & "FROM PP_TMAE_GARESTRES "
'''            vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
'''            vgSql = vgSql & "num_orden = " & numero & " AND "
'''            vgSql = vgSql & "fec_inires = '" & Trim(vlFecha) & "' "
'''            Set vgRegistro = vgConexionBD.Execute(vgSql)
'''            If Not vgRegistro.EOF Then
'''               'mostrar datos de resolucion GE
'''               If IsNull(vgRegistro!COD_TIPRES) Then
'''                  Lbl_GETipoRes = ""
'''               Else
'''                   Lbl_GETipoRes = (vgRegistro!COD_TIPRES)
'''               End If
'''               If IsNull(vgRegistro!NUM_RESGAREST) Then
'''                  Lbl_GENumRes = ""
'''               Else
'''                   Lbl_GENumRes = (vgRegistro!NUM_RESGAREST)
'''               End If
'''               If IsNull(vgRegistro!NUM_ANNORES) Then
'''                  Lbl_GEAnnoRes = ""
'''               Else
'''                   Lbl_GEAnnoRes = (vgRegistro!NUM_ANNORES)
'''               End If
'''               If IsNull(vgRegistro!PRC_DEDUCCION) Then
'''                  Lbl_GEPorcDed = ""
'''               Else
'''                   Lbl_GEPorcDed = (vgRegistro!PRC_DEDUCCION)
'''               End If
'''
'''            End If
'''        End If
'''    End If
'''
''''Obtener datos del último estado de GE
'''    vgSql = ""
'''    vgSql = "SELECT MAX (fec_iniestgarest) as fec_iniestgarest "
'''    vgSql = vgSql & "FROM PP_TMAE_GARESTESTADO "
'''    vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
'''    vgSql = vgSql & "num_orden = " & numero & " "
'''    Set vgRegistro = vgConexionBD.Execute(vgSql)
'''    If Not vgRegistro.EOF Then
'''       If Not IsNull(vgRegistro!fec_iniestgarest) Then
'''
'''            vlFecha = (vgRegistro!fec_iniestgarest)
'''            vgSql = ""
'''            vgSql = "SELECT fec_iniestgarest,cod_dergarest,cod_caususestgarest "
'''            vgSql = vgSql & "FROM PP_TMAE_GARESTESTADO "
'''            vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
'''            vgSql = vgSql & "num_orden = " & numero & " AND "
'''            vgSql = vgSql & "fec_iniestgarest = '" & Trim(vlFecha) & "' "
'''            Set vgRegistro = vgConexionBD.Execute(vgSql)
'''            If Not vgRegistro.EOF Then
'''               'mostrar datos de resolucion GE
'''               If IsNull(vgRegistro!COD_DERGAREST) Then
'''                  Lbl_GEEstado = ""
'''               Else
'''                   vlNombreEstado = ""
'''                   Call flBuscaNombreEstado(vgRegistro!COD_DERGAREST)
'''                   Lbl_GEEstado = Trim(vlNombreEstado)
'''               End If
'''               If IsNull(vgRegistro!cod_caususestgarest) Then
'''                  Lbl_GEEstSuspension = ""
'''               Else
'''                   Lbl_GEEstSuspension = Trim(vgRegistro!cod_caususestgarest) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_CauSusGarEst, Trim(vgRegistro!cod_caususestgarest)))
'''               End If
'''               If IsNull(vgRegistro!fec_iniestgarest) Then
'''                  Lbl_GEFecEst = ""
'''               Else
'''                   Lbl_GEFecEst = DateSerial(Mid((vgRegistro!fec_iniestgarest), 1, 4), Mid((vgRegistro!fec_iniestgarest), 5, 2), Mid((vgRegistro!fec_iniestgarest), 7, 2))
'''               End If
'''
'''            End If
'''       End If
'''    End If
'''
'''    Call flCargaGrillaAsigFam(numero)
'''
'''    Lbl_AFCargas = vlNumCargas
'''
'''    Call flCargaGrillaCCAF(numero)
'''
'''Exit Function
'''Err_flMostrarDatosOtrosBeneficios:
'''    Screen.MousePointer = 0
'''    Select Case Err
'''        Case Else
'''        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'''    End Select
'''
'''End Function
'''
Function flObtenerConceptos()

On Error GoTo Err_flObtenerConceptos

    vlCodConceptos = "("
    For vgI = 0 To Lst_LiqSeleccion.ListCount - 1
        If Lst_LiqSeleccion.Selected(vgI) Then
           'If vgI <> (Lst_LiqSeleccion.ListCount - 1) Then 'HQR 22/10/2005 Se corrige porque producía error al marcar el último concepto de la lista
              If vlCodConceptos <> "(" Then
                 vlCodConceptos = (vlCodConceptos & ",")
              End If
           'End If
           'If vgI <> 0 Then
           vlCodConceptos = (vlCodConceptos & "'" & (Trim(Mid(Lst_LiqSeleccion.List(vgI), 1, (InStr(1, Lst_LiqSeleccion.List(vgI), "-") - 1)))) & "'")
'              vgSql = vgSql & "'" & (Trim(Mid(Lst_LiqSeleccion.List(vgI), 1, (InStr(1, Lst_LiqSeleccion.List(vgI), "-") - 1)))) & "', "
           'End If
           
        End If
        If vgI = (Lst_LiqSeleccion.ListCount - 1) Then
           vlCodConceptos = (vlCodConceptos & ")")
'        Else
'            If vlCodConceptos <> "(" Then
'               vlCodConceptos = (vlCodConceptos & ",")
'            End If
        End If
    Next
    If vlCodConceptos = "()" Then
       vlCodConceptos = "(' ')"
    End If
    
Exit Function
Err_flObtenerConceptos:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function flBuscarCodHabDesCCAF()

On Error GoTo Err_flBuscarCodHabDesCCAF

    vlCodHabDesCCAF = ""
    vgSql = ""
    vgSql = vgSql & "SELECT cod_conhabdes "
    vgSql = vgSql & "FROM MA_TPAR_CONHABDES "
    vgSql = vgSql & "WHERE cod_modorigen = '" & clModOrigenCCAF & "' "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       While Not vgRegistro.EOF
             If vlCodHabDesCCAF = "" Then
                vlCodHabDesCCAF = "("
             End If
             vlCodHabDesCCAF = (vlCodHabDesCCAF & "'" & (vgRegistro!Cod_ConHabDes) & "'")
             vgRegistro.MoveNext
             If Not vgRegistro.EOF Then
                vlCodHabDesCCAF = (vlCodHabDesCCAF & ",")
             End If
       Wend
       vlCodHabDesCCAF = (vlCodHabDesCCAF & ")")
    End If

Exit Function
Err_flBuscarCodHabDesCCAF:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function flBuscaNombreConcepto(Codigo As String)

    vgSql = ""
    vgSql = vgSql & "SELECT gls_conhabdes "
    vgSql = vgSql & "FROM MA_TPAR_CONHABDES "
    vgSql = vgSql & "WHERE cod_conhabdes = " & Trim(Codigo) & " "
    Set vgRs4 = vgConexionBD.Execute(vgSql)
    If Not vgRs4.EOF Then
       vlNombreConcepto = (vgRs4!gls_ConHabDes)
    End If

End Function


Function flBuscaNombreComuna(Codigo As Integer)

    vgSql = ""
    vgSql = vgSql & "SELECT gls_comuna "
    vgSql = vgSql & "FROM MA_TPAR_COMUNA "
    vgSql = vgSql & "WHERE cod_direccion = " & Trim(Codigo) & " "
    Set vgRs4 = vgConexionBD.Execute(vgSql)
    If Not vgRs4.EOF Then
       vlNombreComuna = (vgRs4!gls_comuna)
    End If

End Function

Function flBuscaNombreSucursal(Codigo As String)

    vgSql = ""
    vgSql = "SELECT gls_sucursal "
    vgSql = vgSql & "FROM MA_TPAR_SUCURSAL "
    vgSql = vgSql & "WHERE cod_sucursal = '" & Codigo & "' "
    Set vgRs4 = vgConexionBD.Execute(vgSql)
    If Not vgRs4.EOF Then
       vlNombreSucursal = (vgRs4!gls_sucursal)
    End If

End Function

Function flBuscaNombreEstado(Codigo As String)
'Código de estado de derecho a garantia estatal
    vgSql = ""
    vgSql = "SELECT gls_dergarest "
    vgSql = vgSql & "FROM MA_TPAR_ESTDERGAREST "
    vgSql = vgSql & "WHERE cod_dergarest = '" & Codigo & "' "
    Set vgRs4 = vgConexionBD.Execute(vgSql)
    If Not vgRs4.EOF Then
       vlNombreEstado = (vgRs4!GLS_DERGAREST)
    End If

End Function

Function flBuscaReceptor(numero As Integer) 'numero de orden
    
    'Buscar Tutor asociado a beneficiario
    vgSql = ""
    vgSql = "SELECT gls_nomtut,gls_nomsegtut,gls_pattut,gls_mattut, "
    vgSql = vgSql & "fec_inipodnot,fec_terpodnot,cod_tipoidentut,num_identut "
    vgSql = vgSql & "FROM PP_TMAE_TUTOR "
    vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza) & "' AND "
    vgSql = vgSql & "num_orden = " & numero & " "
    Set vgRs3 = vgConexionBD.Execute(vgSql)
    If Not vgRs3.EOF Then
        vlCodTipoIdenBenTut = (vgRs3!cod_tipoidentut)
        vlNumIdenBenTut = (Trim(vgRs3!num_identut))
        
''       vlRutRec = Format((Trim(vgRs3!rut_tut)), "##,###,##0")
''       vlDgvRec = (vgRs3!dgv_tut)
       vlFechaIni = DateSerial(Mid((vgRs3!fec_inipodnot), 1, 4), Mid((vgRs3!fec_inipodnot), 5, 2), Mid((vgRs3!fec_inipodnot), 7, 2))
       vlFechaTer = DateSerial(Mid((vgRs3!Fec_TerPodNot), 1, 4), Mid((vgRs3!Fec_TerPodNot), 5, 2), Mid((vgRs3!Fec_TerPodNot), 7, 2))
       vlNomRec = Trim(vgRs3!gls_nomtut) + " " + IIf(IsNull(vgRs3!gls_mattut), "", Trim(vgRs3!gls_mattut)) + " " + Trim(vgRs3!gls_pattut) + " " + IIf(IsNull(vgRs3!gls_mattut), "", Trim(vgRs3!gls_mattut))
    Else
    'Asignar Datos del Pensionado o Beneficiario
        vlCodTipoIdenBenTut = (vlCodTipoIdenBenCau)
        vlNumIdenBenTut = (vlNumIdenBenCau)
''        vlRutRec = Trim(Lbl_GrupRut)
''        vlDgvRec = Trim(Lbl_GrupDgv)
        vlFechaIni = Trim(Lbl_PolIniVig)
        vlFechaTer = DateSerial(Mid(Trim(clFechaTopeTer), 1, 4), Mid(Trim(clFechaTopeTer), 5, 2), Mid(Trim(clFechaTopeTer), 7, 2))
        vlNomRec = Trim(Lbl_GrupNombre) + " " + Trim(Lbl_GrupPaterno) + " " + Trim(Lbl_GrupMaterno)
    End If

End Function

Function flLimpiar()

On Error GoTo Err_flLimpiar

'Limpiar frame consulta
    Txt_PenPoliza = ""
    If (Cmb_PenNumIdent.ListCount <> 0) Then
        Cmb_PenNumIdent.ListIndex = 0
    End If
    Txt_PenNumIdent = ""
    ''*Txt_PenDigito = ""
    Lbl_End = ""
    Lbl_PenNombre = ""
    
'Limpiar grilla grupo familiar
  
    Call flInicializaGrillaGrupo

'Limpiar ficha Poliza
    Lbl_PolIniVig = ""
    Lbl_PolTerVig = ""
    Lbl_PolFecDev.Caption = ""
    Lbl_PolNumEndoso = ""
    Lbl_PolTipPen = ""
    Lbl_PolTipRta = ""
    Lbl_PolMod = ""
    Lbl_PolNumCar = ""
    Lbl_PolAfp = ""
    Lbl_PolTasaCto = ""
    Lbl_PolTasaRea = ""
    Lbl_PolMtoPri = ""
    Lbl_PolEstado = ""
    Lbl_PolMesDif = ""
    Lbl_PolMesGar = ""
    Lbl_PolTasaVta = ""
    Lbl_PolTasaPerGar = ""
    Lbl_PolMtoPen = ""
    Lbl_Moneda(0) = "(TM)" 'Moneda Pension
    Lbl_Moneda(1) = "(TM)" 'Moneda Pension Beneficiario
    lblNomAsesor = ""
    Call flLimpiarTab
    
'''Limpiar ficha Grupo Familiar
''    Lbl_GrupPar = ""
''    Lbl_GrupRut = ""
''    Lbl_GrupDgv = ""
''    Lbl_GrupNombre = ""
''    Lbl_GrupPaterno = ""
''    Lbl_GrupDomicilio = ""
''    Lbl_GrupComuna = ""
''    Lbl_GrupFono = ""
''    Lbl_GrupPension = ""
''    Lbl_GrupNumOrden = ""
''    Lbl_GrupSitInv = ""
''    Lbl_GrupFecNac = ""
''    Lbl_GrupFecIngreso = ""
''    Lbl_GrupMaterno = ""
''    Lbl_GrupProvincia = ""
''    Lbl_GrupRegion = ""
''    Lbl_GrupMail = ""
''    Lbl_GrupFecFall = ""
''    Lbl_GrupEstado = ""
''
'''Limpiar ficha Pago de Pensiones
''    Lbl_FPViaPago = ""
''    Lbl_FPSucursal = ""
''    Lbl_FPTipoCta = ""
''    Lbl_FPBanco = ""
''    Lbl_FPNumCta = ""
''    Lbl_PSInstitucion = ""
''    Lbl_PSModPago = ""
''    Lbl_PSMtoPago = ""
''    Lbl_RecRut = ""
''    Lbl_RecDgv = ""
''    Lbl_RecNombre = ""
''    Lbl_RecFecIniVig = ""
''    Lbl_RecFecTerVig = ""
''
''    Call flInicializaGrillaRetJud
''
'''Limpiar ficha Otros Beneficios
''    Lbl_GEEstado = ""
''    Lbl_GEEstSuspension = ""
''    Lbl_GEFecEst = ""
''    Lbl_GETipoRes = ""
''    Lbl_GENumRes = ""
''    Lbl_GEAnnoRes = ""
''    Lbl_GEPorcDed = ""
''    Lbl_AFCargas = ""
''
''    Call flInicializaGrillaAsigFam
''
''    Call flInicializaGrillaCCAF
''
'''Limpiar ficha Liquidaciones (H/D)
''    Txt_LiqFecIni = ""
''    Txt_LiqFecTer = ""
''
''    Call flInicializaGrillaHabDes
''
'''Limpia ficha Emisión de Certificados
''    Opt_CerPensiones.Value = True
''    Opt_CerDecRta.Value = False
''    Opt_CerCarFam.Value = False
''
''    Call flInicializaGrillaCertificados
    
Exit Function
Err_flLimpiar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function flLimpiarTab()

On Error GoTo Err_flLimpiarTab

'Limpiar ficha Grupo Familiar
    Lbl_GrupPar = ""
    Lbl_GrupTipoIdent = ""
    Lbl_GrupNumIdent = ""
    Lbl_GrupNombre = ""
    Lbl_GrupPaterno = ""
    Lbl_GrupDomicilio = ""
    Lbl_GrupComuna = ""
    Lbl_GrupFono = ""
    Lbl_GrupPension = ""
    Lbl_GrupNumOrden = ""
    Lbl_GrupSitInv = ""
    Lbl_GrupFecNac = ""
    Lbl_GrupFecIngreso = ""
    Lbl_GrupMaterno = ""
    Lbl_GrupProvincia = ""
    Lbl_GrupRegion = ""
    Lbl_GrupMail = ""
    Lbl_GrupFecFall = ""
    Lbl_GrupEstado = ""

'Limpiar ficha Pago de Pensiones

    Call flCargarListaConHabDes

    Lbl_FPViaPago = ""
    Lbl_FPSucursal = ""
    Lbl_FPTipoCta = ""
    Lbl_FPBanco = ""
    Lbl_FPNumCta = ""
    Lbl_PSInstitucion = ""
    Lbl_PSModPago = ""
''*    Lbl_PSModPago2 = ""
    Lbl_PSMtoPago = ""
''*    Lbl_PSMtoPago2 = ""
    Lbl_RecRut = ""
    Lbl_RecDgv = ""
    Lbl_RecNombre = ""
    Lbl_RecFecIniVig = ""
    Lbl_RecFecTerVig = ""

    Call flInicializaGrillaRetJud

'Limpiar ficha Otros Beneficios
    ''*Lbl_GEEstado = ""
    ''*Lbl_GEEstSuspension = ""
    ''*Lbl_GEFecEst = ""
    ''*Lbl_GETipoRes = ""
    ''*Lbl_GENumRes = ""
    ''*Lbl_GEAnnoRes = ""
    ''*Lbl_GEPorcDed = ""
    ''*Lbl_AFCargas = ""

    ''*Call flInicializaGrillaAsigFam

    ''*Call flInicializaGrillaCCAF

'Limpiar ficha Liquidaciones (H/D)
    Txt_LiqFecIni = ""
    Txt_LiqFecTer = ""

    Call flInicializaGrillaHabDes

'Limpia ficha Emisión de Certificados
    Opt_CerPensiones.Value = True
    Opt_CerDecRta.Value = False
    Opt_CerCarFam.Value = False

    Call flInicializaGrillaCertificados

Exit Function
Err_flLimpiarTab:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function flDeshabilitarConsulta()

    Fra_PenPoliza.Enabled = True

    SSTab_Consulta.Tab = 0
    SSTab_Consulta.Enabled = False
    
    Fra_Grupo.Enabled = False
    
End Function

Function flHabilitarConsulta()

    Fra_PenPoliza.Enabled = False
    
    SSTab_Consulta.Tab = 0
    SSTab_Consulta.Enabled = True
    
    Fra_Grupo.Enabled = True

End Function

Function flImprimirCerPen()

On Error GoTo Err_flImprimirCerPen
            
   Screen.MousePointer = 11

   vlArchivo = strRpt & "PP_Rpt_CONCertificadoPension.rpt" '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
      MsgBox "Archivo de Reporte de Certificado de Pensión no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
      Screen.MousePointer = 0
      Exit Function
   End If
   
   'Verificar si el beneficiario seleccionado tiene derecho a pension
   vgSql = ""
   vgSql = "SELECT num_orden,cod_estpension "
   vgSql = vgSql & "FROM PP_TMAE_BEN "
   vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
   vgSql = vgSql & "num_endoso = " & vlNumEndoso & " AND "
   vgSql = vgSql & "num_orden = " & Trim(Lbl_GrupNumOrden.Caption) & " AND "
   vgSql = vgSql & "cod_estpension = '" & clCodEstPen99 & "' "
   Set vgRegistro = vgConexionBD.Execute(vgSql)
   If vgRegistro.EOF Then
      MsgBox "El Beneficiario Seleccionado No tiene Derecho a Pensión", vbInformation, "Información"
      Exit Function
   End If
   
      'Buscar Rut del cliente
   vlTipoIdenCia = ""
   vlNumIdenCia = ""
   vgSql = ""
   vgSql = "SELECT cod_tipoidencli,num_idencli,gls_ciucli,gls_nomlarcli "
   vgSql = vgSql & "FROM MA_TMAE_CLIENTE "
   Set vgRegistro = vgConexionBD.Execute(vgSql)
   If Not vgRegistro.EOF Then
      ''*vlTipoIdenCia = Trim(vgRegistro!cod_tipoidencli)
      vlTipoIdenCia = fgBuscarNombreTipoIden(Trim(vgRegistro!cod_tipoidencli), True)
      vlNumIdenCia = Trim(vgRegistro!num_idencli)
      If Not IsNull(vgRegistro!gls_ciucli) Then
        vlCiuCia = Trim(vgRegistro!gls_ciucli)
      Else
        vlCiuCia = ""
      End If
      If Not IsNull(vgRegistro!gls_nomlarcli) Then
          vlNombreCompania = Trim(vgRegistro!gls_nomlarcli)
      Else
          vlNombreCompania = ""
      End If
   End If
 
 
 Dim objRep As New ClsReporte
 Dim MontoTexto As String
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   rs.CursorLocation = adUseClient

   rs.Open "PP_LISTA_CERTIFICADO_PENSIONES.LISTAR('" & Trim(Txt_PenPoliza.Text) & "'," & vlNumEndoso & "," & Trim(Lbl_GrupNumOrden.Caption) & ")", vgConexionBD, adOpenStatic, adLockReadOnly
   
   If Not rs.EOF Then
    MontoTexto = fgConvierteNumeroLetras(rs!Mto_Pension, rs!Cod_Moneda)
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PP_Rpt_CONCertificadoPension.rpt"), ".RPT", ".TTX"), 1)
         
    If objRep.CargaReporte(strRpt & "", "PP_Rpt_CONCertificadoPension.rpt", "Reporte de Certificado de Pensión", rs, True, _
                            ArrFormulas("NombreCompania", vlNombreCompania), _
                            ArrFormulas("NombreSistema", vgNombreSistema), _
                            ArrFormulas("NombreSubSistema", vgNombreSubSistema), _
                            ArrFormulas("TipoIdenCompania", vlTipoIdenCia), _
                            ArrFormulas("NumIdenCompania", vlNumIdenCia), _
                            ArrFormulas("NombreUsuario", vlNombreCompania), _
                            ArrFormulas("DeptoUsuario", clDeptoUsuario), _
                            ArrFormulas("CiudadCia", vlCiuCia), _
                            ArrFormulas("MontoPalabras", MontoTexto)) = False Then
                            
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Function
    End If
  End If
'
'   'CMV-20060925 I
'   Rpt_General.Formulas(0) = "NombreCompania = '" & vlNombreCompania & "'"
''   Rpt_General.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
'   'CMV-20060925 F
'   Rpt_General.Formulas(1) = "NombreSistema= '" & vgNombreSistema & "'"
'   Rpt_General.Formulas(2) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
'   Rpt_General.Formulas(3) = "TipoIdenCompania = '" & vlTipoIdenCia & "'"
'   Rpt_General.Formulas(4) = "NumIdenCompania = '" & vlNumIdenCia & "'"
'
'   Rpt_General.Formulas(5) = "NombreUsuario = '" & vlNombreCompania & "'"
'   Rpt_General.Formulas(6) = "DeptoUsuario = '" & clDeptoUsuario & "'"
'   Rpt_General.Formulas(7) = "CiudadCia = '" & vlCiuCia & "'"
'   Rpt_General.Formulas(8) = "MtoPensionBruta = " & Str(Lbl_GrupPension)
'
'   Rpt_General.SubreportToChange = ""
'   Rpt_General.Destination = crptToWindow
'   Rpt_General.WindowState = crptMaximized
'   Rpt_General.WindowTitle = ""
'   'Rpt_Reporte.SelectionFormula = ""
'   Rpt_General.Action = 1
   Screen.MousePointer = 0

Exit Function
Err_flImprimirCerPen:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function
Private Sub pBuscaRepresentante(TP As String)
On Error GoTo Err_Cargarep
Dim vlSql As String
Dim vlRepresentante, vlDocum As String

TP = Mid(TP, 1, 2)

If TP = "08" Then

    vlSql = ""
    vlSql = "SELECT * FROM pd_tmae_polrep a, ma_tpar_tipoiden b WHERE "
    vlSql = vlSql & "num_poliza = '" & Txt_PenPoliza & "' and a.cod_tipoidenrep = b.cod_tipoiden"
    Set vgRs = vgConexionBD.Execute(vlSql)
    If Not vgRs.EOF Then
        vlRepresentante = IIf(Not IsNull(vgRs!Gls_NombresRep), vgRs!Gls_NombresRep, "")
        vlRepresentante = vlRepresentante & " " & IIf(Not IsNull(vgRs!Gls_ApepatRep), vgRs!Gls_ApepatRep, "")
        vlRepresentante = vlRepresentante & " " & IIf(Not IsNull(vgRs!Gls_ApematRep), vgRs!Gls_ApematRep, "")
        vlDocum = IIf(Not IsNull(vgRs!gls_tipoidencor), vgRs!gls_tipoidencor, "") & " " & IIf(Not IsNull(vgRs!Num_Idenrep), vgRs!Num_Idenrep, "")
    End If
    vgRs.Close
Else
    vlSql = ""
    vlSql = "SELECT * FROM pd_tmae_polben WHERE "
    vlSql = vlSql & "num_poliza = '" & Txt_PenPoliza & "' and cod_par='99'"
    Set vgRs = vgConexionBD.Execute(vlSql)
    If Not vgRs.EOF Then
        vlRepresentante = IIf(Not IsNull(vgRs!Gls_NomBen), vgRs!Gls_NomBen, "")
        vlRepresentante = vlRepresentante & " " & IIf(Not IsNull(vgRs!Gls_NomSegBen), vgRs!Gls_NomSegBen, "")
        vlRepresentante = vlRepresentante & " " & IIf(Not IsNull(vgRs!Gls_PatBen), vgRs!Gls_PatBen, "")
        vlRepresentante = vlRepresentante & " " & IIf(Not IsNull(vgRs!Gls_MatBen), vgRs!Gls_MatBen, "")
        'vlDocum = IIf(Not IsNull(vgRs!gls_Tipoidencor), vgRs!gls_Tipoidencor, "") & " " & IIf(Not IsNull(vgRs!Num_Idenrep), vgRs!Num_Idenrep, "")
    End If
    vgRs.Close
End If
 
Exit Sub
Err_Cargarep:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub
Public Function fgObtenerNombreSuc_Usuario(iCodigo As String) As String
Dim vlRegBuscar As ADODB.Recordset
    
    fgObtenerNombreSuc_Usuario = ""
    
    vgQuery = "SELECT gls_sucursal as nombre "
    vgQuery = vgQuery & "FROM pd_tpar_sucursal "
    vgQuery = vgQuery & "WHERE "
    vgQuery = vgQuery & "cod_sucursal = '" & iCodigo & "'"
    Set vlRegBuscar = vgConexionBD.Execute(vgQuery)
    If Not (vlRegBuscar.EOF) Then
        If Not IsNull(vlRegBuscar!Nombre) Then
            fgObtenerNombreSuc_Usuario = Trim(vlRegBuscar!Nombre)
        End If
    End If
    vlRegBuscar.Close

End Function

Public Function fgObtenerNombre_TextoCompuesto(iTexto As String) As String
'Función: Permite obtener el Nombre o Descripción de un Texto que tiene el Código y la
'Descripción separados por un Guión
'Parámetros de Entrada :
'- iTexto     => Texto que contiene el Código y Descripción
'Parámetros de Salida :
'- Devuelve la descripción del Texto
    
    If (InStr(1, iTexto, "-") <> 0) Then
        fgObtenerNombre_TextoCompuesto = Trim(Mid(iTexto, InStr(1, iTexto, "-") + 1, Len(iTexto)))
    Else
        fgObtenerNombre_TextoCompuesto = UCase(Trim(iTexto))
    End If

End Function

Function FlImprimeConstan()

Dim vlCodPar As String
Dim vlCodDerpen As String
Dim vlNombreSucursal, vlNombreTipoPension As String
Dim rs As ADODB.Recordset
Dim vlFecTras As String
Dim objRep As New ClsReporte
 'Validar el Ingreso de la Póliza
    If Txt_PenPoliza = "" Then
        MsgBox "Debe ingresar Póliza a Consultar.", vbCritical, "Error de Datos"
        Txt_PenPoliza.SetFocus
        Exit Function
    End If
    
    'Valida que exista la Póliza
    If Trim(Txt_PenNumIdent) = "" Then
        MsgBox "Debe Buscar Datos de la Póliza", vbCritical, "Error de Datos"
        Cmd_BuscarPol.SetFocus
        Exit Function
    End If
    
    vlCodPar = "99"   'Causante
    vlCodDerpen = "10" 'Sin Derecho a Pension
    vlNombreSucursal = "" 'fgObtenerNombreSuc_Usuario(vgUsuarioSuc)
    vlNombreTipoPension = fgObtenerNombre_TextoCompuesto(Lbl_PolTipPen)
    Call pBuscaRepresentante(Lbl_PolTipPen)
  
    
    vgSql = "SELECT  num_poliza, fec_traspaso "
    vgSql = vgSql & "FROM "
    vgSql = vgSql & "pd_tmae_polprirec"
    vgSql = vgSql & " WHERE num_poliza = '" & Txt_PenPoliza.Text & "'"
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not (vgRegistro.EOF) Then
   
        vlFecTras = vgRegistro!fec_traspaso
        'vlFecTras = Mid(vlFecTras, 7, 4) & Mid(vlFecTras, 4, 2) & Mid(vlFecTras, 1, 2)
   End If
    
On Error GoTo mierror


    '''******************************************CONDTSNCIA DE POLIZA*********************************************************************
    
    If Mid(Lbl_PolTipPen.Caption, 1, 2) = "08" Then
        Exit Function
    End If
    
    Dim NomReporte As String
    
    'If CInt(Lbl_Diferidos.Caption) > 0 Then
        NomReporte = "PP_Rpt_PolizaConstaDif.rpt"
    'Else
    '    NomReporte = "PD_Rpt_PolizaConstaDif.rpt"
    'End If
    
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "PD_LISTA_POLIZA.LISTAR('" & Txt_PenPoliza.Text & "', '" & txt_End & "')", vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PP_Rpt_PolizaConsta.rpt"), ".RPT", ".TTX"), 1)
    
    If objRep.CargaReporte(strRpt & "", NomReporte, "Póliza Constancia", rs, True, _
                            ArrFormulas("NombreAfi", Lbl_PenNombre.Caption), _
                            ArrFormulas("TipoPension", vlNombreTipoPension), _
                            ArrFormulas("MesGar", Lbl_PolMesGar.Caption), _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("Concatenar", vlcobertura), _
                            ArrFormulas("Sucursal", "Surquillo"), _
                            ArrFormulas("RepresentanteNom", ""), _
                            ArrFormulas("RepresentanteDoc", ""), _
                            ArrFormulas("CodTipPen", Left(Trim(Lbl_PolTipPen), 2)), _
                            ArrFormulas("TipoDocTit", Left(Trim(Cmb_PenNumIdent), 2)), _
                            ArrFormulas("NumDocTit", Trim(Txt_PenNumIdent)), _
                            ArrFormulas("fec_trasp", vlFecTras)) = False Then
            
            
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Function
    End If
    
    
Exit Function
mierror:
    MsgBox "No pudo cargar el reporte " & Err.Description, vbInformation
    
End Function




Function flImprimirConPen()

On Error GoTo Err_flImprimirConPen
            
            
            
   Screen.MousePointer = 11

   vlArchivo = strRpt & "PP_Rpt_CONConstanciaPension.rpt" '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
      MsgBox "Archivo de Reporte de Certificado de Pensión no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
      Screen.MousePointer = 0
      Exit Function
   End If
   
   'Verificar si el beneficiario seleccionado tiene derecho a pension
   vgSql = ""
   vgSql = "SELECT num_orden,cod_estpension "
   vgSql = vgSql & "FROM PP_TMAE_BEN "
   vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
   vgSql = vgSql & "num_endoso = " & vlNumEndoso & " AND "
   vgSql = vgSql & "num_orden = " & Trim(Lbl_GrupNumOrden.Caption) & " AND "
   vgSql = vgSql & "cod_estpension = '" & clCodEstPen99 & "' "
   Set vgRegistro = vgConexionBD.Execute(vgSql)
   If vgRegistro.EOF Then
      MsgBox "El Beneficiario Seleccionado No tiene Derecho a Pensión", vbInformation, "Información"
      Exit Function
   End If
   
      'Buscar Rut del cliente
   vlTipoIdenCia = ""
   vlNumIdenCia = ""
   vgSql = ""
   vgSql = "SELECT cod_tipoidencli,num_idencli,gls_ciucli,gls_nomlarcli "
   vgSql = vgSql & "FROM MA_TMAE_CLIENTE "
   Set vgRegistro = vgConexionBD.Execute(vgSql)
   If Not vgRegistro.EOF Then
      ''*vlTipoIdenCia = Trim(vgRegistro!cod_tipoidencli)
      vlTipoIdenCia = fgBuscarNombreTipoIden(Trim(vgRegistro!cod_tipoidencli), True)
      vlNumIdenCia = Trim(vgRegistro!num_idencli)
      If Not IsNull(vgRegistro!gls_ciucli) Then
        vlCiuCia = Trim(vgRegistro!gls_ciucli)
      Else
        vlCiuCia = ""
      End If
      If Not IsNull(vgRegistro!gls_nomlarcli) Then
          vlNombreCompania = UCase(Trim(vgRegistro!gls_nomlarcli))
      Else
          vlNombreCompania = ""
      End If
    End If
    
    'Determinar Nombre y Cargo del Auditor
    vgSql = "Select * from PD_TMAE_APODERADO "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
        If Not IsNull(vgRegistro!gls_nomApo) Then vlNombreApoderado = Trim(vgRegistro!gls_nomApo)
        If Not IsNull(vgRegistro!gls_carApo) Then vlCargoApoderado = Trim(vgRegistro!gls_carApo)
    End If
    vgRegistro.Close
        
   'Identificación de la Empresa
   vgSql = ""
   vgSql = "SELECT gls_tipoiden,num_idencli,gls_ciucli "
   vgSql = vgSql & "FROM MA_TPAR_TIPOIDEN,MA_TMAE_CLIENTE "
   vgSql = vgSql & "where cod_tipoiden=cod_tipoidencli "
   Set vgRegistro = vgConexionBD.Execute(vgSql)
   If Not vgRegistro.EOF Then
     
      If Not IsNull(vgRegistro!gls_tipoiden) Then
        vlTipoIden = UCase(Trim(vgRegistro!gls_tipoiden))
      Else
        vlTipoIden = ""
      End If
      If Not IsNull(vgRegistro!num_idencli) Then
          vlNumIden = (Trim(vgRegistro!num_idencli))
      Else
          vlNumIden = ""
      End If
      If Not IsNull(vgRegistro!gls_ciucli) Then
        vlCiuCia = Trim(vgRegistro!gls_ciucli)
      Else
        vlCiuCia = ""
      End If
    End If
    
    Dim objRep As New ClsReporte

   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   rs.CursorLocation = adUseClient

 
         
    'If objRep.CargaReporte(strRpt & "", "PP_Rpt_CONConstanciaPension.rpt", "Constancia de Pensión", RS, True, _
    '                        ArrFormulas("NombreCompaniaLargo", vlNombreCompania), _
    '                        ArrFormulas("IdenCompania", vlTipoIden), _
    '                        ArrFormulas("NumCompania", vlNumIden), _
    '                        ArrFormulas("Periodos", vlcobertura), _
    '                        ArrFormulas("Nombre", vlNombreApoderado), _
    '                        ArrFormulas("CiudadCia", vlCiuCia), _
    '                        ArrFormulas("Cargo", vlCargoApoderado)) = False Then
    '
    '    MsgBox "No se pudo abrir el reporte", vbInformation
    '    Exit Function
    'End If
  'End If
   
'
'    vgSql = ""
'    vlCodTp = "TP"
'    vlCodTr = "TR"
'    vlCodAl = "AL"
'    vlCodPa = "99"
'
'    vgSql = "SELECT  p.num_poliza,p.cod_tippension,p.num_idenafi,a.gls_tipoidencor,p.cod_cuspp,"
'    vgSql = vgSql & "p.cod_tipren,p.num_mesdif,p.cod_modalidad,p.num_mesgar,"
'    vgSql = vgSql & "p.mto_priuni,p.mto_pension,p.mto_pensiongar,"
'    vgSql = vgSql & "t.gls_elemento as gls_pension,"
'    vgSql = vgSql & "r.gls_elemento as gls_renta,"
'    vgSql = vgSql & "m.gls_elemento as gls_modalidad,"
'    vgSql = vgSql & "be.gls_nomben,be.gls_patben,be.gls_matben, p.cod_afp,"
'    vgSql = vgSql & "p.cod_moneda, p.mto_valmoneda, cod_liquidacion, pr.fec_traspaso, "
'    vgSql = vgSql & "p.cod_cobercon, b.gls_cobercon,p.cod_dercre, p.cod_dergra, be.fec_nacben "
'    vgSql = vgSql & "FROM "
'    vgSql = vgSql & "pd_tmae_poliza p, pd_tmae_polprirec pr, ma_tpar_tabcod t, ma_tpar_tabcod r, "
'    vgSql = vgSql & "ma_tpar_tabcod m, pd_tmae_polben be, ma_tpar_tipoiden a, ma_tpar_cobercon b "
'    vgSql = vgSql & "WHERE "
'    vgSql = vgSql & "p.num_poliza = '" & Trim(Txt_PenPoliza) & "' AND "
'    vgSql = vgSql & "p.num_poliza = pr.num_poliza AND "
'    vgSql = vgSql & "p.num_poliza = be.num_poliza AND "
'    vgSql = vgSql & "p.num_endoso = be.num_endoso AND "
'    vgSql = vgSql & "be.cod_par = '" & Trim(vlCodPa) & "' AND "
'    vgSql = vgSql & "t.cod_tabla = '" & Trim(vlCodTp) & "' AND "
'    vgSql = vgSql & "t.cod_elemento = p.cod_tippension AND "
'    vgSql = vgSql & "r.cod_tabla = '" & Trim(vlCodTr) & "' AND "
'    vgSql = vgSql & "r.cod_elemento = p.cod_tipren AND "
'    vgSql = vgSql & "m.cod_tabla = '" & Trim(vlCodAl) & "' AND "
'    vgSql = vgSql & "m.cod_elemento = p.cod_modalidad AND "
'    vgSql = vgSql & "p.cod_tipoidenafi = a.cod_tipoiden AND "
'    vgSql = vgSql & "p.cod_cobercon = b.cod_cobercon"
'    Set vgRegistro = vgConexionBD.Execute(vgSql)
'    If Not (vgRegistro.EOF) Then
'        Fra_Poliza.Enabled = False
'        vlFecNacTitular = DateSerial(Mid(vgRegistro!Fec_NacBen, 1, 4), Mid(vgRegistro!Fec_NacBen, 5, 2), Mid(vgRegistro!Fec_NacBen, 7, 2))
'        vlcobertura = vgRegistro!Gls_Renta
'        If vgRegistro!Cod_Modalidad = 1 Then
'            If Not IsNull(vgRegistro!gls_modalidad) Then
'                vlcobertura = vlcobertura & " " & vgRegistro!gls_modalidad
'            End If
'        Else
'            If Not IsNull(vgRegistro!gls_modalidad) Then
'                vlcobertura = vlcobertura & " CON P. " & vgRegistro!gls_modalidad
'            End If
'        End If
'        If vgRegistro!Cod_CoberCon <> 0 Then
'            If Not IsNull(vgRegistro!GLS_COBERCON) Then
'                vlcobertura = vlcobertura & " CON " & vgRegistro!GLS_COBERCON
'            End If
'        End If
'        If vgRegistro!Cod_DerCre = "S" Then
'            vlcobertura = vlcobertura & " CON D.CRECER"
'        End If
'
'        If vgRegistro!Cod_DerGra = "S" Then
'            vlcobertura = vlcobertura & " Y CON GRATIFICACIÓN"
'        End If
'   End If
'   vgQuery = ""
'   'vgQuery = vgQuery & "{PP_TMAE_BEN.rut_ben} = " & (Str(Trim(Lbl_GrupRut.Caption))) & " "
'   vgQuery = vgQuery & "{PP_TMAE_POLIZA.num_poliza} = '" & Trim(Txt_PenPoliza) & "' AND "
'   vgQuery = vgQuery & "{PP_TMAE_POLIZA.num_endoso} = " & (Lbl_End) & " AND "
'   vgQuery = vgQuery & "{PP_TMAE_BEN.num_orden} = " & (Lbl_GrupNumOrden) & " "
'
'
'   Rpt_General.Reset
'   Rpt_General.WindowState = crptMaximized
'   Rpt_General.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
'   Rpt_General.Connect = vgRutaDataBase       ' o App.Path & "\Nestle.mdb"
'   Rpt_General.SelectionFormula = vgQuery
'
''   vgPalabra = ""
''   vgPalabra = "12345678-K"
'
'   Rpt_General.Formulas(0) = ""
'   Rpt_General.Formulas(1) = ""
'   Rpt_General.Formulas(2) = ""
'   Rpt_General.Formulas(3) = ""
'   Rpt_General.Formulas(4) = ""
'
'   Rpt_General.Formulas(5) = ""
'   Rpt_General.Formulas(6) = ""
'   Rpt_General.Formulas(7) = ""
'
   'CMV-20060925 I
'   Rpt_General.Formulas(0) = "NombreCompaniaLargo = '" & vlNombreCompania & "'"
'   Rpt_General.Formulas(1) = "IdenCompania='" & vlTipoIden & "'"
'   Rpt_General.Formulas(2) = "NumCompania='" & vlNumIden & "'"
'   Rpt_General.Formulas(3) = "Periodos='" & vlcobertura & "'"
'   Rpt_General.Formulas(4) = "Nombre = '" & vlNombreApoderado & "'"
'   Rpt_General.Formulas(5) = "Cargo= '" & vlCargoApoderado & "'"
'   Rpt_General.SubreportToChange = ""
'   Rpt_General.Destination = crptToWindow
'   Rpt_General.WindowState = crptMaximized
'   Rpt_General.WindowTitle = ""
   'Rpt_Reporte.SelectionFormula = ""
   'Rpt_General.Action = 1
   Screen.MousePointer = 0

Exit Function
Err_flImprimirConPen:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function flImprimirCerDecRta()

On Error GoTo Err_flImprimirCerDecRta

    MsgBox "La Opción Seleccionada No se Encuentra Disponible en esta Versión de la Aplicación", 16, "Archivo no encontrado"
    
Exit Function
Err_flImprimirCerDecRta:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

'''Function flImprimirCerCarFam()
'''
'''On Error GoTo Err_flImprimirCerCarFam
'''
'''   Screen.MousePointer = 11
'''
'''   vlArchivo = strRpt & "PP_Rpt_CONCertificadoCargas.rpt" '\Reportes
'''   If Not fgExiste(vlArchivo) Then     ', vbNormal
'''      MsgBox "Archivo de Reporte de Certificado de Cargas Familiares no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
'''      Screen.MousePointer = 0
'''      Exit Function
'''   End If
'''
''''   'Carga Tabla Temporal de cargas familiares activas (independientemente de si son factibles o no de pago)
''''   Call flCargaTemporalCargas
'''
'''   'Verificar si el beneficiario seleccionado tiene derecho a pension
''''   vgSql = ""
''''   vgSql = "SELECT num_orden,cod_estpension "
''''   vgSql = vgSql & "FROM PP_TMAE_BEN "
''''   vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
''''   vgSql = vgSql & "num_endoso = " & vlNumEndoso & " AND "
''''   vgSql = vgSql & "num_orden = " & Trim(Lbl_GrupNumOrden.Caption) & " AND "
''''   vgSql = vgSql & "cod_estpension = '" & clCodEstPen99 & "' "
''''   vgSql = vgSql & "ORDER BY num_orden "
''''   Set vgRegistro = vgConexionBD.Execute(vgSql)
''''   If Not vgRegistro.EOF Then
''''      If Trim(vlNumOrden) = Trim(vgRegistro!Num_Orden) Then
''''         vlRut = Trim(Str(Txt_PenRut.Text))
''''         vlOrden = vlNumOrden
''''      Else
''''          vlRut = Trim(Str(Lbl_GrupRut.Caption))
''''          vlOrden = Trim(Lbl_GrupNumOrden.Caption)
''''      End If
''''   End If
'''   vlNombreBenef = ""
'''   vlRutBenef = ""
'''   vlTipoRenta = ""
'''   vlTipoPension = ""
'''   vlTipoModalidad = ""
'''   vlFechaVigenciaRta = ""
'''   vlMtoPensionBruta = 0
'''
'''   vgSql = ""
'''   vgSql = "SELECT b.num_orden,b.rut_ben,b.dgv_ben,b.cod_estpension, "
'''   vgSql = vgSql & "b.gls_nomben,b.gls_patben,b.gls_matben,b.mto_pension, "
'''   vgSql = vgSql & "p.cod_tipren,p.cod_tippension,p.cod_modalidad,p.fec_vigencia  "
'''   vgSql = vgSql & "FROM pp_tmae_ben b, pp_tmae_poliza p "
'''   vgSql = vgSql & "WHERE b.num_poliza = p.num_poliza AND "
'''   vgSql = vgSql & "b.num_endoso = p.num_endoso AND "
'''   vgSql = vgSql & "b.num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
'''   vgSql = vgSql & "b.num_endoso = " & vlNumEndoso & " AND "
'''   vgSql = vgSql & "b.num_orden = '" & Trim(Lbl_GrupNumOrden.Caption) & "' AND "
'''   vgSql = vgSql & "b.cod_estpension = '" & clCodEstPen99 & "' "
'''   Set vgRegistro = vgConexionBD.Execute(vgSql)
'''   If Not vgRegistro.EOF Then
'''      vlRut = Trim(Str(Lbl_GrupRut.Caption))
'''      vlOrden = Trim(Lbl_GrupNumOrden.Caption)
'''      vlNombreBenef = (vgRegistro!Gls_NomBen) & " " & (vgRegistro!Gls_PatBen) & " " & (vgRegistro!Gls_MatBen)
'''      vlRutBenef = (vgRegistro!Rut_Ben) & " - " & (vgRegistro!Dgv_Ben)
'''      vlTipoRenta = (vgRegistro!Cod_TipRen)
'''      vlTipoPension = (vgRegistro!Cod_TipPension)
'''      vlTipoModalidad = (vgRegistro!Cod_Modalidad)
'''      vlFechaVigenciaRta = DateSerial(Mid((vgRegistro!Fec_Vigencia), 1, 4), Mid((vgRegistro!Fec_Vigencia), 5, 2), Mid((vgRegistro!Fec_Vigencia), 7, 2))
'''      vlMtoPensionBruta = (vgRegistro!Mto_Pension)
'''   Else
'''       MsgBox "El Beneficiario Seleccionado No tiene Derecho a Pensión", vbInformation, "Información"
'''       Exit Function
'''   End If
'''
'''   If vlTipoRenta <> "" Then
'''      vlTipoRenta = fgBuscarGlosaElemento(vgCodTabla_TipRen, vlTipoRenta)
'''   End If
'''   If vlTipoPension <> "" Then
'''      vlTipoPension = fgBuscarGlosaElemento(vgCodTabla_TipPen, vlTipoPension)
'''   End If
'''   If vlTipoModalidad <> "" Then
'''      vlTipoModalidad = fgBuscarGlosaElemento(vgCodTabla_AltPen, vlTipoModalidad)
'''   End If
'''
'''    'Obtener último periodo de pago
'''    vgSql = ""
'''    vgSql = "SELECT num_perpago "
'''    vgSql = vgSql & "FROM pp_tmae_liqpagopendef "
'''    vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
'''    vgSql = vgSql & "num_orden = " & Trim(Lbl_GrupNumOrden.Caption) & " "
'''    vgSql = vgSql & "ORDER BY num_perpago DESC "
'''    Set vgRegistro = vgConexionBD.Execute(vgSql)
'''    If Not vgRegistro.EOF Then
'''        vlNumPerPago = (vgRegistro!Num_PerPago)
'''    End If
'''
'''    'Verificar existencia de cargas familiares para el beneficiario seleccionado
'''    vlNumCargasPagadas = 0
'''    vgSql = ""
'''    vgSql = "SELECT num_ordencar "
'''    vgSql = vgSql & "FROM pp_tmae_pagoasigdef "
'''    vgSql = vgSql & "WHERE num_perpago = '" & Trim(vlNumPerPago) & "' AND "
'''    vgSql = vgSql & "num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
'''    vgSql = vgSql & "num_orden = " & Trim(Lbl_GrupNumOrden.Caption) & " AND "
'''    vgSql = vgSql & "cod_tipreceptor = '" & clRecPensionado & "' "
'''    Set vgRegistro = vgConexionBD.Execute(vgSql)
'''    If Not vgRegistro.EOF Then
'''        While Not vgRegistro.EOF
'''            vlNumCargasPagadas = vlNumCargasPagadas + 1
'''            vgRegistro.MoveNext
'''        Wend
'''    End If
'''    If vlNumCargasPagadas = 0 Then
'''        MsgBox "El Beneficiario Seleccionado No tiene Cargas Familiares Activas", vbInformation, "Información"
'''        Exit Function
'''    End If
'''
''''   'Verificar existencia de cargas familiares para el beneficiario seleccionado
''''   vgPalabra = ""
''''   vlNumCargas = 0
''''   vgSql = ""
''''   vgSql = "SELECT num_ordenrec,cod_estvigencia "
''''   vgSql = vgSql & "FROM PP_TMAE_ASIGFAM "
''''   vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
'''''   vgSql = vgSql & "num_endoso = " & vlNumEndoso & " AND "
''''   vgSql = vgSql & "num_ordenrec = " & Str(vlOrden) & " AND "
''''   vgSql = vgSql & "cod_estvigencia = '" & clCargaActiva & "' "
''''   vgSql = vgSql & "ORDER BY num_orden "
''''   Set vgRegistro = vgConexionBD.Execute(vgSql)
''''   If Not vgRegistro.EOF Then
''''      While Not vgRegistro.EOF
''''            vlNumCargas = vlNumCargas + 1
''''            vgRegistro.MoveNext
''''      Wend
''''      vgPalabra = "AND {PP_TMAE_ASIGFAM.cod_estvigencia} = '" & (Trim(clCargaActiva)) & "' "
''''   Else
''''       MsgBox "El Beneficiario Seleccionado No tiene Cargas Familiares Activas", vbInformation, "Información"
''''       Exit Function
''''   End If
'''
'''   'Buscar Rut del cliente
'''   vlRutCia = ""
'''   vgSql = ""
'''   vgSql = "SELECT rut_cliente,dgv_cliente,gls_ciucli,gls_nomlarcli "
'''   vgSql = vgSql & "FROM MA_TMAE_CLIENTE "
'''   Set vgRegistro = vgConexionBD.Execute(vgSql)
'''   If Not vgRegistro.EOF Then
'''      vlRutCia = Format((Trim(vgRegistro!rut_cliente)), "##,###,##0") & " - " & (Trim(vgRegistro!dgv_cliente))
'''      If Not IsNull(vgRegistro!gls_ciucli) Then
'''        vlCiuCia = (vgRegistro!gls_ciucli)
'''      Else
'''        vlCiuCia = ""
'''      End If
'''      If Not IsNull(vgRegistro!gls_nomlarcli) Then
'''          vlNombreCompania = Trim(vgRegistro!gls_nomlarcli)
'''      Else
'''          vlNombreCompania = ""
'''      End If
'''   End If
'''
''''CMV-20060925 I
''''   'Buscar Nombre de Usuario
''''   vlNombreUsuario = ""
''''   vgSql = ""
''''   vgSql = "SELECT gls_nombre,gls_paterno,gls_materno "
''''   vgSql = vgSql & "FROM MA_TMAE_USUARIO "
''''   vgSql = vgSql & "WHERE cod_usuario = '" & vgUsuario & "' AND "
''''   vgSql = vgSql & "cod_sistema = '" & vgTipoSistema & "' "
''''   Set vgRegistro = vgConexionBD.Execute(vgSql)
''''   If Not vgRegistro.EOF Then
''''      vlNombreUsuario = (vgRegistro!gls_nombre) & " " & (vgRegistro!gls_paterno) & " " & (vgRegistro!gls_materno)
''''   End If
''''CMV-20060925 F
'''
'''   'vlNumEndosoNoBen = 3
'''
'''   vgQuery = ""
'''   vgQuery = vgQuery & "{PP_TMAE_PAGOASIGDEF.num_perpago} = '" & Trim(vlNumPerPago) & "' AND "
'''   vgQuery = vgQuery & "{PP_TMAE_PAGOASIGDEF.num_poliza} = '" & Trim(Txt_PenPoliza) & "' AND "
'''   vgQuery = vgQuery & "{PP_TMAE_PAGOASIGDEF.num_orden} = " & vlOrden & " AND "
'''   vgQuery = vgQuery & "{PP_TMAE_PAGOASIGDEF.cod_tipreceptor} <> '" & clCodTipReceptorR & "' AND "
'''   vgQuery = vgQuery & "{PP_TMAE_BEN.num_endoso} = " & vlNumEndoso & " "
'''   'vgQuery = vgQuery & "{MA_TPAR_TABCOD.cod_tabla} = '" & Trim(vgCodTabla_ParNoBen) & "' "
''''   vgQuery = vgQuery & "{PP_TMAE_NOBEN.num_endoso} = " & vlNumEndosoNoBen & " )"
'''
'''   Rpt_General.Reset
'''   Rpt_General.WindowState = crptMaximized
'''   Rpt_General.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
'''   Rpt_General.Connect = vgRutaDataBase       ' o App.Path & "\Nestle.mdb"
'''   Rpt_General.SelectionFormula = vgQuery
'''
'''   Rpt_General.Formulas(0) = ""
'''   Rpt_General.Formulas(1) = ""
'''   Rpt_General.Formulas(2) = ""
'''   Rpt_General.Formulas(3) = ""
'''
'''   Rpt_General.Formulas(4) = ""
'''   Rpt_General.Formulas(5) = ""
'''   Rpt_General.Formulas(6) = ""
'''   Rpt_General.Formulas(7) = ""
'''   Rpt_General.Formulas(8) = ""
'''   Rpt_General.Formulas(9) = ""
'''   Rpt_General.Formulas(10) = ""
'''   Rpt_General.Formulas(11) = ""
'''   Rpt_General.Formulas(12) = ""
'''   Rpt_General.Formulas(13) = ""
'''   Rpt_General.Formulas(14) = ""
'''
'''   Rpt_General.Formulas(0) = "NombreCompania = '" & vlNombreCompania & "'"
'''   Rpt_General.Formulas(1) = "NombreSistema= '" & vgNombreSistema & "'"
'''   Rpt_General.Formulas(2) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
'''   Rpt_General.Formulas(3) = "RutCompania = '" & vlRutCia & "'"
'''
'''   Rpt_General.Formulas(4) = "NombreUsuario = '" & vlNombreCompania & "'"
'''   Rpt_General.Formulas(5) = "DeptoUsuario = '" & clDeptoUsuario & "'"
'''   Rpt_General.Formulas(6) = "NumeroCargas = '" & vlNumCargasPagadas & "'"
'''   Rpt_General.Formulas(7) = "CiudadCia = '" & vlCiuCia & "'"
'''
'''   Rpt_General.Formulas(8) = "Nombre Benef = '" & vlNombreBenef & "'"
'''   Rpt_General.Formulas(9) = "Rut Benef = '" & vlRutBenef & "'"
'''   Rpt_General.Formulas(10) = "TipoRenta = '" & vlTipoRenta & "'"
'''   Rpt_General.Formulas(11) = "TipoPension = '" & vlTipoPension & "'"
'''   Rpt_General.Formulas(12) = "TipoModalidad = '" & vlTipoModalidad & "'"
'''   Rpt_General.Formulas(13) = "FechaVigenciaRta = '" & vlFechaVigenciaRta & "'"
'''   Rpt_General.Formulas(14) = "MtoPensionBruta = " & Str(vlMtoPensionBruta) & ""
'''
''''   vgQuery = ""
''''   vgQuery = vgQuery & "{PP_TMAE_PAGOASIGDEF.num_perpago} = '" & Trim(vlNumPerPago) & "' AND "
''''   vgQuery = vgQuery & "{PP_TMAE_PAGOASIGDEF.num_poliza} = '" & Trim(Txt_PenPoliza) & "' AND "
''''   vgQuery = vgQuery & "{PP_TMAE_PAGOASIGDEF.num_orden} = " & vlOrden & " "
'''
'''   Rpt_General.Destination = crptToWindow
'''   Rpt_General.WindowState = crptMaximized
'''   'HQR 17/01/2007 Se agrega SubreportToChange
'''   Rpt_General.SubreportToChange = "SUBCerCarFam"
'''   Rpt_General.SelectionFormula = ""
'''   Rpt_General.Connect = vgRutaDataBase
'''   Rpt_General.WindowTitle = ""
'''   'Rpt_Reporte.SelectionFormula = ""
'''   Rpt_General.Action = 1
'''   Rpt_General.SubreportToChange = ""
'''   Screen.MousePointer = 0
'''
'''Exit Function
'''Err_flImprimirCerCarFam:
'''    Screen.MousePointer = 0
'''    Select Case Err
'''        Case Else
'''        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'''    End Select
'''
'''End Function
'''
''''Function flCargaTemporalCargas(numero As Integer)
''''
''''On Error GoTo Err_flCargaTemporalCargas
''''
''''    'Obtener todos los beneficiarios que en alguna oportunidad han sido
''''    'cargas familiares activas
''''    vlEstCerEst = "N"
''''    vgSql = ""
''''    vgSql = "SELECT DISTINCT (num_orden) "
''''    vgSql = vgSql & "FROM PP_TMAE_ASIGFAM "
''''    vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
''''    vgSql = vgSql & "num_ordenrec = " & vlNumOrden & " "
''''    vgSql = vgSql & "ORDER by num_orden ASC "
''''    Set vlRegistro2 = vgConexionBD.Execute(vgSql)
''''    If Not vlRegistro2.EOF Then
''''       Call flInicializaGrillaAsigFam
''''       While Not vlRegistro2.EOF
''''
''''           'Obtener datos de la carga ben o no ben
''''               If (vlRegistro2!Num_Orden) >= 50 Then
''''
''''                    vgSql = ""
''''                    vgSql = "SELECT cod_ascdes,rut_ben,dgv_ben,num_orden, "
''''                    vgSql = vgSql & "gls_nomben,gls_patben,gls_matben "
''''                    vgSql = vgSql & "FROM PP_TMAE_NOBEN "
''''                    vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
'''''                    vgSql = vgSql & "num_endoso = " & Str(vlNumEndoso) & " AND "
''''                    vgSql = vgSql & "num_orden = " & (vlRegistro2!Num_Orden) & " "
''''                    Set vgRegistro = vgConexionBD.Execute(vgSql)
''''                    If Not vgRegistro.EOF Then
''''                       vlCodPar = " " & Trim(vgRegistro!COD_ASCDES) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_ParNoBen, Trim(vgRegistro!COD_ASCDES)))
''''                    End If
''''               Else
''''                   vgSql = ""
''''                   vgSql = "SELECT cod_par,rut_ben,dgv_ben,num_orden, "
''''                   vgSql = vgSql & "gls_nomben,gls_patben,gls_matben "
''''                   vgSql = vgSql & "FROM PP_TMAE_BEN "
''''                   vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
''''                   vgSql = vgSql & "num_endoso = " & Str(vlNumEndoso) & " AND "
''''                   vgSql = vgSql & "num_orden = " & (vlRegistro2!Num_Orden) & " "
''''                   Set vgRegistro = vgConexionBD.Execute(vgSql)
''''                   If Not vgRegistro.EOF Then
''''                      vlCodPar = " " & Trim(vgRegistro!Cod_Par) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_Par, Trim(vgRegistro!Cod_Par)))
''''                   End If
''''               End If
''''
''''          'Obtener Estado de Certificado de Estudio a la fecha
''''            'vlEstCerEst = ""
''''            vgSql = ""
''''            vgSql = "SELECT fec_inicerest,fec_tercerest "
''''            vgSql = vgSql & "FROM PP_TMAE_CERESTUDIO "
''''            vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
''''            vgSql = vgSql & "num_orden = " & Str(vlRegistro2!Num_Orden) & " "
''''            vgSql = vgSql & "ORDER BY fec_inicerest DESC "
''''            Set vgRs = vgConexionBD.Execute(vgSql)
''''            If Not vgRs.EOF Then
''''               vlFechaActual = fgBuscaFecServ
''''               vlFechaActual = Format(CDate(Trim(vlFechaActual)), "yyyymmdd")
''''               If vlFechaActual >= (vgRs!Fec_IniCerEst) And _
''''                  vlFechaActual <= (vgRs!FEC_TERCEREST) Then
''''                  vlEstCerEst = "S"
''''               Else
''''                   vlEstCerEst = "N"
''''               End If
''''            End If
''''
''''            'Obtener los datos de la última vez que la carga
''''            'estuvo activa
''''            vgSql = ""
''''            vgSql = "SELECT MAX(fec_iniactiva) ,fec_iniactiva,fec_teractiva, "
''''            vgSql = vgSql & "cod_caususpension,cod_estvigencia "
''''            vgSql = vgSql & "FROM PP_TMAE_ASIGFAM "
''''            vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
''''            vgSql = vgSql & "num_orden = " & Str(vlRegistro2!Num_Orden) & " "
''''            vgSql = vgSql & "GROUP BY fec_iniactiva,fec_teractiva, "
''''            vgSql = vgSql & "cod_caususpension,cod_estvigencia "
''''            Set vlRegistro3 = vgConexionBD.Execute(vgSql)
''''            If Not vlRegistro3.EOF Then
''''                vlCodCauSus = " " & Trim(vlRegistro3!COD_CAUSUSPENSION) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_CauSupAsiFam, Trim(vlRegistro3!COD_CAUSUSPENSION)))
''''                vlCodEstado = " " & Trim(vlRegistro3!COD_ESTVIGENCIA) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_EstVigAsiFam, Trim(vlRegistro3!COD_ESTVIGENCIA)))
''''                vlFechaIni = DateSerial(Mid((vlRegistro3!FEC_INIACTIVA), 1, 4), Mid((vlRegistro3!FEC_INIACTIVA), 5, 2), Mid((vlRegistro3!FEC_INIACTIVA), 7, 2))
''''                vlFechaTer = DateSerial(Mid((vlRegistro3!FEC_TERACTIVA), 1, 4), Mid((vlRegistro3!FEC_TERACTIVA), 5, 2), Mid((vlRegistro3!FEC_TERACTIVA), 7, 2))
''''                vlFecha = Trim(vlFechaIni) & " - " & Trim(vlFechaTer)
''''            End If
''''
''''          'vlCodPar = " " & Trim(vgRs2!COD_PAR) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_Par, Trim(vgRs2!COD_PAR)))
''''
''''          'vlFechaEfectoAsigFam = DateSerial(Mid((vgRs2!FEC_INIACTIVA), 1, 4), Mid((vgRs2!FEC_INIACTIVA), 5, 2), Mid((vgRs2!FEC_INIACTIVA), 7, 2))
''''          vlFechaEfectoAsigFam = fgValidaFechaEfecto(vlFechaIni, Trim(Txt_PenPoliza.Text), vlNumOrden)
''''
''''
''''          Msf_GrillaAsigFam.AddItem (vlCodPar) & vbTab _
''''          & (" " & Format((Trim(vgRegistro!Rut_Ben)), "##,###,##0") & " - " & (Trim(vgRegistro!Dgv_Ben))) & vbTab _
''''          & (Trim(vgRegistro!Gls_NomBen)) & " " & (Trim(vgRegistro!Gls_PatBen)) & " " & (Trim(vgRegistro!Gls_MatBen)) & vbTab _
''''          & (vlFecha) & vbTab _
''''          & (vlCodCauSus) & vbTab _
''''          & (vlCodEstado) & vbTab _
''''          & (vlEstCerEst) & vbTab _
''''          & (vlFechaEfectoAsigFam)
''          'Falta Fecha de Efecto
''
''
'''          If Trim(vlRegistro3!COD_ESTVIGENCIA) = clCargaActiva Then
'''             vlNumCargas = vlNumCargas + 1
'''          End If
''          vlRegistro2.MoveNext
''       Wend
''    End If
''    vgRegistro.Close
''
''Exit Function
''Err_flCargaTemporalCargas:
''    Screen.MousePointer = 0
''    Select Case Err
''        Case Else
''        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
''    End Select
''
''End Function
'
Function flBuscaFecEfectoInsSalud(numero As Integer)

On Error GoTo Err_flBuscaFecEfectoInsSalud

    vlFechaEfectoSalud = ""
    vgSql = ""
    vgSql = "SELECT fec_pago,cod_inssalud "
    vgSql = vgSql & "FROM PP_TMAE_LIQPAGOPENDEF "
    vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
    vgSql = vgSql & "num_endoso = " & vlNumEndoso & " AND "
    vgSql = vgSql & "num_orden = " & numero & " "
    vgSql = vgSql & "ORDER BY fec_pago DESC "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       If (vgRegistro!Cod_InsSalud) <> vlCodInsSalud Then
          Exit Function
       End If
       While Not vgRegistro.EOF
             If (vgRegistro!Cod_InsSalud) = vlCodInsSalud Then
                vlFechaEfectoSalud = DateSerial(Mid((vgRegistro!fec_pago), 1, 4), Mid((vgRegistro!fec_pago), 5, 2), Mid((vgRegistro!fec_pago), 7, 2))
'                vlFechaEfectoSalud = fgValidaFechaEfecto(Trim(vlFechaEfectoSalud), Txt_PenPoliza.Text, numero)
             Else
                 Exit Function
             End If
             vgRegistro.MoveNext
       Wend
    End If

Exit Function
Err_flBuscaFecEfectoInsSalud:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

'*********************Funciones Imprimir Liquidacion

Function flInformeLiqPago(iGlosaOpcion As String)
'Imprime Liquidaciones de Pension
On Error GoTo Err_VenTut
    Dim vlSQLQuery As String
    Dim vgRutCliente As String
    Dim vgDgvCliente As String
    Dim vlTipoId As String
    ''''''''''''''''''''''''
    Dim rs As ADODB.Recordset
    Dim rsLiq As ADODB.Recordset
    Dim objRep As New ClsReporte
    Dim cadena As String
    Dim nombres, Direccion, fec_finter As String
    Dim LNGa As Long
    
    vgRutCliente = ""
    vgDgvCliente = ""
    
    vlFechaDesde = Format(CDate(Trim(Txt_LiqFecIni.Text)), "yyyymmdd")
    vlFechaHasta = Format(CDate(Trim(Txt_LiqFecTer.Text)), "yyyymmdd")
    
    'I--- ABV 12/03/2005 ---
    'vlPago = Trim(fgBuscaTipoPago(Trim(Txt_PenPoliza.Text)))
    vlPago = "R"
    'F--- ABV 12/03/2005 ---
    vlTipoId = (Trim(Mid(Lbl_GrupTipoIdent.Caption, 1, (InStr(1, Lbl_GrupTipoIdent.Caption, "-") - 1))))
    
    If Not flLlenaTemporal(vlFechaDesde, vlFechaHasta, Trim(Txt_PenPoliza.Text), (vlTipoId), (Lbl_GrupNumIdent.Caption), clOpcionDEF, vlPago) Then
        Exit Function
    End If
    
    'esta función se utiliza cuando Peter quiere sacar masivamente las boletas de pago de un grupo de polizas determinado
    'cuando se utiliza esta función, la anterior se pone en comentario o se salta con el debug
    'RVF 03/02/2011
    'If Not flLlenaTemporal_masivo(vlFechaDesde, vlFechaHasta, Trim(Txt_PenPoliza.Text), (vlTipoId), (Lbl_GrupNumIdent.Caption), clOpcionDEF, vlPago) Then
    '    Exit Function
    'End If
    
    Screen.MousePointer = 11
    
    vlArchivo = strRpt & "PP_Rpt_LiquidacionRV.rpt"
    If Not fgExiste(vlArchivo) Then
        MsgBox "Archivo de Reporte no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
        Screen.MousePointer = 0
        Exit Function
    End If

    '' Reporte de Carta de Renovacion


    'cadena = "select a.*, gls_comuna, gls_provincia, gls_region, g.gls_tipoiden, g.gls_tipoidencor,"
    'cadena = cadena & " c.num_idenben, c.gls_nomben as gls_nomben1, c.gls_nomsegben, c.gls_patben, c.gls_matben, gls_dirben"
    'cadena = cadena & " from pp_ttmp_liquidacion a"
    'cadena = cadena & " join pp_tmae_poliza b on a.num_poliza=b.num_poliza and a.num_endoso=b.num_endoso"
    'cadena = cadena & " join pp_tmae_ben c on a.num_poliza=c.num_poliza and a.num_endoso=c.num_endoso and a.num_orden=c.num_orden"
    'cadena = cadena & " join ma_tpar_comuna d on c.cod_direccion=d.cod_direccion"
    'cadena = cadena & " join ma_tpar_provincia e on d.cod_provincia=e.cod_provincia"
    'cadena = cadena & " join ma_tpar_region f on e.cod_region=f.cod_region"
    'cadena = cadena & " join ma_tpar_tipoiden g on c.cod_tipoidenben=g.cod_tipoiden"
    'cadena = cadena & " where cod_usuario='" & vgUsuario & "'"
    
    cadena = "select a.*, gls_comuna, gls_provincia, gls_region, g.gls_tipoiden, g.gls_tipoidencor,"
    cadena = cadena & " c.num_idenben, c.gls_nomben as gls_nomben1, c.gls_nomsegben, c.gls_patben, c.gls_matben, gls_dirben, h.gls_tipoiden as tipoRec, i.gls_tipoiden as tipoTit"
    cadena = cadena & " from pp_ttmp_liquidacion a"
    cadena = cadena & " join pp_tmae_poliza b on a.num_poliza=b.num_poliza and a.num_endoso=b.num_endoso"
    cadena = cadena & " join pp_tmae_ben c on a.num_poliza=c.num_poliza and a.num_endoso=c.num_endoso and a.num_orden=c.num_orden"
    cadena = cadena & " join ma_tpar_comuna d on c.cod_direccion=d.cod_direccion"
    cadena = cadena & " join ma_tpar_provincia e on d.cod_provincia=e.cod_provincia"
    cadena = cadena & " join ma_tpar_region f on e.cod_region=f.cod_region"
    cadena = cadena & " join ma_tpar_tipoiden g on c.cod_tipoidenben=g.cod_tipoiden"
    'RRR 22/07/2015
    cadena = cadena & " join ma_tpar_tipoiden h on a.cod_tipoidenreceptor=h.cod_tipoiden"
    cadena = cadena & " join ma_tpar_tipoiden i on a.cod_tipoidentit=i.cod_tipoiden"
    'RRR 22/07/2015
    cadena = cadena & " where cod_usuario='" & vgUsuario & "'"
    cadena = cadena & " and a.num_poliza in (select num_poliza from pp_tmae_liqpagopendef where num_perpago=a.num_perpago and cod_tipopago<>'P')"
    cadena = cadena & " order by 3"
    Set rsLiq = New ADODB.Recordset
    rsLiq.CursorLocation = adUseClient
    rsLiq.Open cadena, vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
    vlRutCliente = flRutCliente 'vgRutCliente + " - " + vgDgvCliente
    
    LNGa = CreateFieldDefFile(rsLiq, Replace(UCase(strRpt & "Estructura\PP_Rpt_LiquidacionRV.rpt"), ".RPT", ".TTX"), 1)

    If objRep.CargaReporte(strRpt & "", "PP_Rpt_LiquidacionRV_ind.rpt", "Informe de Liquidación de Rentas Vitalicias", rsLiq, True, _
                            ArrFormulas("NombreCompania", UCase("Protecta SA compañia de Seguros")), _
                            ArrFormulas("rutcliente", vlRutCliente)) = False Then

        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Function
    End If
    
'descomentar
    'If objRep.CargaReporte_toPdf(strRpt & "", "PP_Rpt_LiquidacionRV.rpt", "Informe de Liquidación de Rentas Vitalicias",  rsLiq, True, _
    '                        ArrFormulas("NombreCompania", UCase("Protecta SA compañia de Seguros")), _
    '                        ArrFormulas("rutcliente", vlRutCliente)) = False Then

    '    MsgBox "No se pudo abrir el reporte", vbInformation
    '    Exit Function
    'End If
'/descomentar




'    vlRutCliente = flRutCliente 'vgRutCliente + " - " + vgDgvCliente
'    vgQuery = "{PP_TTMP_LIQUIDACION2.COD_USUARIO} = '" & vgUsuario & "'"
'    Rpt_Reporte.Reset
'    Rpt_Reporte.WindowState = crptMaximized
'    Rpt_Reporte.ReportFileName = vlArchivo
'    Rpt_Reporte.SelectionFormula = vgQuery
'    Rpt_Reporte.Connect = vgRutaDataBase
'    Rpt_Reporte.Formulas(0) = "NombreCompania='" & UCase(vgNombreCompania) & "'"
''    Rpt_Reporte.Formulas(1) = "NombreSistema= '" & vgNombreSistema & "'"
''    Rpt_Reporte.Formulas(2) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
'    Rpt_Reporte.Formulas(3) = "rutcliente= '" & vlRutCliente & "'"
'    Rpt_Reporte.Destination = crptToWindow
'    Rpt_Reporte.WindowTitle = "Informe de Liquidación de Rentas Vitalicias"
''    Rpt_Reporte.SubreportToChange = "PP_Rpt_MensajesPoliza.rpt"
''    Rpt_Reporte.Connect = vgRutaDataBase
'    Rpt_Reporte.Action = 1
'    Rpt_Reporte.SubreportToChange = ""
'    Screen.MousePointer = 0
    
    '' Reporte de Carta de Renovacion
    
    cadena = "select L.num_poliza, L.GLS_DIRECCION, L.GLS_NOMRECEPTOR, L.GLS_NOMSEGRECEPTOR, L.GLS_PATRECEPTOR, L.GLS_MATRECEPTOR,fec_tercer"
    cadena = cadena & " from pp_tmae_liqpagopendef L, pp_tmae_certificado C"
    cadena = cadena & " where L.FEC_PAGO >= '" & vlFechaDesde & "' AND L.FEC_PAGO <= '" & vlFechaHasta & "'"
    cadena = cadena & " and L.num_poliza=C.num_poliza and L.num_orden=C.num_orden"
    cadena = cadena & " and L.num_poliza='" & Trim(Txt_PenPoliza.Text) & "'"
    cadena = cadena & " and fec_tercer=(select max(fec_tercer) from pp_tmae_certificado where num_poliza=L.num_poliza)"
    cadena = cadena & " and fec_tercer < to_char(to_date('" & vlFechaHasta & "', 'YYYYMMDD') + to_number(to_char(LAST_DAY(to_date('" & vlFechaHasta & "', 'YYYYMMDD')), 'DD')), 'YYYYMMDD')"
    cadena = cadena & " order by 1"


    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

    rs.Open cadena, vgConexionBD, adOpenForwardOnly, adLockReadOnly

    If Not rs.EOF Then
        'MsgBox "No hay datos para mostrar", vbExclamation, "Pagos Pendientes Acumulados"
        
        nombres = rs!Gls_NomReceptor & " " & rs!Gls_NomSegReceptor & " " & rs!Gls_PatReceptor & " " & rs!Gls_MatReceptor
        Direccion = rs!Gls_Direccion
        fec_finter = rs!FEC_TERCER
    Else
        Exit Function
    End If

    'Dim LNGa As Long
    LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PP_Rpt_CartaRenovacion.rpt"), ".RPT", ".TTX"), 1)

    If objRep.CargaReporte(strRpt & "", "PP_Rpt_CartaRenovacion.rpt", "Informe Detallado de pagos pendientes Acumulado", rs, True, _
                            ArrFormulas("NombreCompania", "Protecta SA compañia de Seguros"), _
                            ArrFormulas("NombreTitular", " & nombres &"), _
                            ArrFormulas("Direccion", " & direccion & "), _
                            ArrFormulas("Poliza", " & Trim(Txt_PenPoliza.Text) & "), _
                            ArrFormulas("FechaFin", " & fec_finter & ")) = False Then

        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Function
    End If
    
Exit Function
Err_VenTut:
    Screen.MousePointer = 0
    MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
End Function

Private Function flRutCliente() As String
   'Buscar Rut del cliente
   flRutCliente = ""
   vgSql = "SELECT cli.num_idencli, "
   vgSql = vgSql & "cli.gls_dircli,cli.gls_comcli,cli.gls_ciucli, "
   vgSql = vgSql & "cli.gls_fonocli,cli.gls_faxcli, "
   vgSql = vgSql & "cli.gls_nomlarcli,cli.gls_correocli,cli.gls_paiscli "
   vgSql = vgSql & "FROM ma_tmae_cliente cli "
   Set vgRegistro = vgConexionBD.Execute(vgSql)
   If Not vgRegistro.EOF Then
        If Not IsNull(vgRegistro!gls_nomlarcli) Then
            vlNombreCompania = Trim(vgRegistro!gls_nomlarcli)
        Else
            vlNombreCompania = ""
        End If
        If Not IsNull(vgRegistro!gls_dircli) Then
            vlDirCompania = Trim(vgRegistro!gls_dircli)
        Else
            vlDirCompania = ""
        End If
        If Not IsNull(vgRegistro!gls_comcli) Then
            vlDirCompania = vlDirCompania & ", " & Trim(vgRegistro!gls_comcli)
        End If
        If Not IsNull(vgRegistro!gls_ciucli) Then
            vlDirCompania = vlDirCompania & ", " & Trim(vgRegistro!gls_ciucli)
        End If
        If Not IsNull(vgRegistro!gls_paiscli) Then
            vlDirCompania = vlDirCompania & ", " & Trim(vgRegistro!gls_paiscli)
        End If
        If Not IsNull(vgRegistro!gls_fonocli) Then
            vlFonoCompania = "Teléfono: " & Trim(vgRegistro!gls_fonocli)
        Else
            vlFonoCompania = ""
        End If
        If Not IsNull(vgRegistro!gls_faxcli) Then
            vlFonoCompania = vlFonoCompania & ", Fax: " & Trim(vgRegistro!gls_faxcli)
        End If
        If Not IsNull(vgRegistro!gls_correocli) Then
            vlFonoCompania = vlFonoCompania & ". e-mail: " & Trim(vgRegistro!gls_correocli)
        End If
      'vlDirCompania = Trim(vgRegistro!gls_dircli) & ", " & Trim(vgRegistro!gls_comcli) & ", " & Trim(vgRegistro!gls_ciucli) & ", " & Trim(vgRegistro!gls_paiscli)
      'vlFonoCompania = "Teléfono: " & Trim(vgRegistro!gls_fonocli) & ", Fax: " & Trim(vgRegistro!gls_faxcli) & ". e-mail: " & Trim(vgRegistro!gls_correocli)
      flRutCliente = (Trim(vgRegistro!num_idencli))
   End If
'   vlNombreCompania = "Le Mans Desarrollo Compania de Seguros de Vida S.A. (en Quiebra)"
'   vlDirCompania = "Encomenderos N° 113, piso 2, Las Condes, Santiago, Chile"
'   vlFonoCompania = "Teléfono: 378 70 14, Fax: 246 08 55. e-mail: lemansvida@lemans.cl"
   
End Function


Function flLlenaTemporal(iFecDesde, iFecHasta, iPoliza, iTipoIden, iNumIden, iGlosaOpcion, iPago) As Boolean

    Dim vlSql As String, vlTB As ADODB.Recordset
    Dim vlNumConceptosHab As Long, vlNumConceptosDesc As Long
    Dim vlItem As Long, vlPoliza As String, vlOrden As String
    Dim vlIndImponible As String, vlIndTributable As String
    Dim vlTipReceptor As String
    Dim vlNumIdenReceptor As String, vlCodTipoIdenReceptor As Long
    Dim vlPerPago As String
    Dim vlTB2 As ADODB.Recordset
    Dim vlTipPension As String
    Dim vlViaPago As String
    Dim vlCajaComp As String
    Dim vlInsSalud As String
    Dim vlCodDireccion As Double
    Dim vlAfp As String
    Dim vlFecPago As String
    Dim vlDescViaPago As String
    Dim vlSucursal As String
    Dim vlDescSucursal As String
    
    flLlenaTemporal = False
    vlItem = 1
    vlNumConceptosHab = 0
    vlNumConceptosDesc = 0
    vlPoliza = ""
    vlOrden = 0
    vlPerPago = ""
    vlFecPago = ""
    vlTipPension = ""
    vlViaPago = ""
    vlCajaComp = ""
    vlInsSalud = ""
    vlCodDireccion = 0
    vlAfp = ""
    vlSucursal = ""
    vlDescSucursal = ""
    'VARIABLES GENERALES
    stTTMPLiquidacion.Cod_Usuario = vgUsuario
    
    'Elimina Datos de la Tabla Temporal
    vlSql = "DELETE FROM PP_TTMP_LIQUIDACION WHERE COD_USUARIO = '" & vgUsuario & "'"
    vgConexionBD.Execute (vlSql)
    
   'mv 20190625
    'cuando seleccione algun concepro que tenga que ver con regularizaciones (MEMOS) solo imprimir eso
    Dim intUbica, intUbica2 As Integer
    Dim concatenaString As String
    concatenaString = ""
    
    intUbica = InStr(1, vgQuery, "84")
    If intUbica > 0 Then
        concatenaString = "'84'"
    End If
    
    'mv 20190626
    intUbica2 = InStr(1, vgQuery, "85")
    If intUbica2 > 0 Then
        If concatenaString = "" Then
            concatenaString = concatenaString & "'85'"
        Else
            concatenaString = concatenaString & ",'85'"
        End If
    End If
    
    vlSql = "SELECT DISTINCT Z.* FROM (SELECT P.NUM_POLIZA, L.NUM_ENDOSO, P.NUM_ORDEN, P.COD_CONHABDES, P.MTO_CONHABDES, C.COD_TIPMOV, P.NUM_PERPAGO, P.COD_TIPOIDENRECEPTOR,P.NUM_IDENRECEPTOR, P.COD_TIPRECEPTOR,"
    vlSql = vlSql & " L.COD_DIRECCION, L.FEC_PAGO, L.GLS_NOMRECEPTOR, L.GLS_NOMSEGRECEPTOR, L.GLS_PATRECEPTOR, L.GLS_MATRECEPTOR, L.MTO_LIQPAGAR,"
    'vlSql = vlSql & " (SELECT gls_dirben FROM PP_TMAE_BEN WHERE NUM_POLIZA=P.NUM_POLIZA and num_endoso=(select max(num_endoso) from pp_tmae_poliza where num_poliza=P.NUM_POLIZA) and num_orden=P.NUM_ORDEN) as GLS_DIRECCION,"
    vlSql = vlSql & " PD.Gls_Direccion,"
    vlSql = vlSql & " L.COD_TIPPENSION, L.COD_VIAPAGO, L.COD_SUCURSAL, L.COD_INSSALUD, L.MTO_PENSION, L.NUM_CARGAS, L.MTO_HABER, L.MTO_DESCUENTO, B.NUM_IDENBEN, B.COD_TIPOIDENBEN, B.GLS_NOMBEN, B.GLS_NOMSEGBEN, B.GLS_PATBEN, B.GLS_MATBEN, "
    vlSql = vlSql & " C.GLS_CONHABDES, M.COD_SCOMP, POL.COD_AFP, L.COD_MONEDA, M.GLS_ELEMENTO AS MONEDA, COD_VEJEZ,"
    vlSql = vlSql & " TV.GLS_ELEMENTO AS GLS_TP2, TP.GLS_ELEMENTO AS GLS_TP, PD.COD_DIRECCION AS COD_DIRENUE"
    'RRR 22/07/2015
    vlSql = vlSql & " , (SELECT  GLS_NOMBEN || ' ' || GLS_NOMSEGBEN || ' ' || GLS_PATBEN || ' ' || GLS_MATBEN FROM PP_TMAE_BEN WHERE NUM_POLIZA=P.NUM_POLIZA AND NUM_ORDEN=1 AND NUM_ENDOSO=1) AS NOMTIT"
    vlSql = vlSql & " , (SELECT NUM_IDENBEN FROM PP_TMAE_BEN WHERE NUM_POLIZA=P.NUM_POLIZA AND NUM_ORDEN=1 AND NUM_ENDOSO=1) AS NUMIDENTIT"
    vlSql = vlSql & " , (SELECT COD_TIPOIDENBEN FROM PP_TMAE_BEN WHERE NUM_POLIZA=P.NUM_POLIZA AND NUM_ORDEN=1 AND NUM_ENDOSO=1) AS TIPIDENTIT,COD_TRIBUTABLE,COD_IMPONIBLE"
    'RRR 22/07/2015
    vlSql = vlSql & " FROM PP_TMAE_PAGOPENDEF P JOIN"
    vlSql = vlSql & " PP_TMAE_LIQPAGOPENDEF L ON P.NUM_POLIZA=L.NUM_POLIZA AND L.NUM_PERPAGO = P.NUM_PERPAGO AND  L.COD_TIPOIDENRECEPTOR=P.COD_TIPOIDENRECEPTOR AND L.NUM_IDENRECEPTOR=P.NUM_IDENRECEPTOR AND L.COD_TIPRECEPTOR = P.COD_TIPRECEPTOR  AND L.NUM_ORDEN=P.NUM_ORDEN JOIN"
    vlSql = vlSql & " MA_TPAR_CONHABDES C ON P.COD_CONHABDES  = C.COD_CONHABDES JOIN"
    
    'MV 280616
    vlSql = vlSql & " PP_TMAE_POLIZA POL ON L.NUM_POLIZA = POL.NUM_POLIZA  and L.NUM_ENDOSO=POL.NUM_ENDOSO JOIN"
    vlSql = vlSql & " PP_TMAE_BEN B ON L.NUM_POLIZA = B.NUM_POLIZA AND L.NUM_ORDEN = B.NUM_ORDEN AND B.NUM_ENDOSO=L.NUM_ENDOSO JOIN"
'    vlSql = vlSql & " PP_TMAE_POLIZA POL ON L.NUM_POLIZA = POL.NUM_POLIZA AND L.NUM_ENDOSO=POL.NUM_ENDOSO JOIN"
'    vlSql = vlSql & " PP_TMAE_BEN B ON L.NUM_POLIZA = B.NUM_POLIZA AND L.NUM_ORDEN = B.NUM_ORDEN  AND L.NUM_ENDOSO=B.NUM_ENDOSO JOIN"
    
    vlSql = vlSql & " PD_TMAE_POLIZA PD ON PD.NUM_POLIZA = POL.NUM_POLIZA  JOIN"
    vlSql = vlSql & " MA_TPAR_TABCOD M ON L.COD_MONEDA = M.COD_ELEMENTO AND  M.COD_TABLA = 'TM' JOIN"
    vlSql = vlSql & " MA_TPAR_TABCOD TV ON PD.COD_VEJEZ=TV.COD_ELEMENTO AND TV.COD_TABLA = 'TV'  JOIN"
    vlSql = vlSql & " MA_TPAR_TABCOD TP ON PD.COD_TIPPENSION=TP.COD_ELEMENTO AND TP.COD_TABLA = 'TP'"
    vlSql = vlSql & " WHERE L.NUM_POLIZA = '" & Trim(iPoliza) & "'"
    'vlSql = vlSql & " WHERE "
    vlSql = vlSql & " AND B.COD_TIPOIDENBEN = " & Trim(iTipoIden) & ""
    'vlSql = vlSql & " B.COD_TIPOIDENBEN = " & Trim(iTipoIden) & ""
    vlSql = vlSql & " AND B.NUM_IDENBEN = " & Trim(iNumIden) & ""
    If iPago = "P" Then 'PRIMER PAGO
        vlSql = vlSql & " AND L.COD_TIPOPAGO in ('P')"
    ElseIf iPago = "R" Then 'PAGO EN REGIMEN
        vlSql = vlSql & " AND L.COD_TIPOPAGO in ('R')"
    End If
    If (iPago = "T") Then 'TODO TIPO PAGO
        vlSql = vlSql & " AND L.COD_TIPOPAGO in ('R','P')"
    End If
    vlSql = vlSql & " AND L.FEC_PAGO >= '" & iFecDesde & "'"
    vlSql = vlSql & " AND L.FEC_PAGO <= '" & iFecHasta & "'"
    vlSql = vlSql & " AND PD.NUM_ENDOSO=(select max(num_endoso) from PD_TMAE_POLIZA where num_poliza=P.NUM_POLIZA)"
    'vlSql = vlSql & " AND POL.NUM_ENDOSO=(select max(num_endoso) from PP_TMAE_LIQPAGOPENDEF where num_poliza=P.NUM_POLIZA)"
    vlSql = vlSql & " AND P.COD_CONHABDES NOT IN (60)"
    
    'mv 20190625  regularizaciones (memos manuales)
    If concatenaString <> "" Then
        vlSql = vlSql & " AND P.cod_conhabdes IN(" & concatenaString & ")"
    Else
        vlSql = vlSql & " AND P.cod_conhabdes not IN('84','85')"
    End If
    
    vlSql = vlSql & " ORDER BY P.NUM_PERPAGO, P.NUM_POLIZA, P.NUM_ORDEN,P.NUM_IDENRECEPTOR, P.COD_TIPOIDENRECEPTOR, P.COD_TIPRECEPTOR, C.COD_IMPONIBLE DESC, C.COD_TRIBUTABLE DESC, C.COD_TIPMOV DESC"
    'mv
    vlSql = vlSql & " )Z ORDER BY Z.NUM_PERPAGO, Z.NUM_POLIZA, Z.NUM_ORDEN,Z.NUM_IDENRECEPTOR, Z.COD_TIPOIDENRECEPTOR, Z.COD_TIPRECEPTOR, Z.COD_IMPONIBLE DESC, Z.COD_TRIBUTABLE DESC, Z.COD_TIPMOV DESC"
    Set vlTB = vgConexionBD.Execute(vlSql)
    If Not vlTB.EOF Then
        Do While Not vlTB.EOF
            If vlPoliza <> vlTB!num_poliza Or vlOrden <> vlTB!Num_Orden Or vlNumIdenReceptor <> vlTB!Num_IdenReceptor Or vlCodTipoIdenReceptor <> vlTB!Cod_TipoIdenReceptor Or vlTipReceptor <> vlTB!Cod_TipReceptor Or (vlPerPago <> vlTB!Num_PerPago And vlPago = "R") Or (vlFecPago <> vlTB!fec_pago And vlPago = "P") Then 'hqr 17/03/2006 Se agrega número de periodo
                'Reinicia el Contador
                vlItem = 1
                vlPoliza = vlTB!num_poliza
                vlOrden = vlTB!Num_Orden
                vlNumConceptosHab = 0
                vlNumConceptosDesc = 0
                vlNumIdenReceptor = vlTB!Num_IdenReceptor
                vlCodTipoIdenReceptor = vlTB!Cod_TipoIdenReceptor
                vlTipReceptor = vlTB!Cod_TipReceptor
                stTTMPLiquidacion.num_poliza = vlPoliza
                stTTMPLiquidacion.Num_IdenReceptor = vlTB!Num_IdenReceptor
                stTTMPLiquidacion.Cod_TipoIdenReceptor = vlTB!Cod_TipoIdenReceptor
                stTTMPLiquidacion.Num_Orden = vlOrden
                stTTMPLiquidacion.num_endoso = IIf(IsNull(vlTB!num_endoso), 0, vlTB!num_endoso)
                stTTMPLiquidacion.Cod_TipReceptor = vlTB!Cod_TipReceptor
                If stTTMPLiquidacion.Cod_TipReceptor <> "R" Then
                    stTTMPLiquidacion.Gls_Direccion = IIf(IsNull(vlTB!Gls_Direccion), "", vlTB!Gls_Direccion)
                Else
                    stTTMPLiquidacion.Gls_Direccion = " "
                End If
                stTTMPLiquidacion.Gls_NomReceptor = vlTB!Gls_NomReceptor & " " & IIf(IsNull(vlTB!Gls_NomSegReceptor), "", vlTB!Gls_NomSegReceptor & " ") & vlTB!Gls_PatReceptor & IIf(IsNull(vlTB!Gls_MatReceptor), "", " " + vlTB!Gls_MatReceptor)
                If stTTMPLiquidacion.Cod_TipReceptor <> "R" Then
                    stTTMPLiquidacion.Cod_Direccion = vlTB!Cod_Direccion
                Else
                    stTTMPLiquidacion.Cod_Direccion = "0"
                End If
                stTTMPLiquidacion.Cod_TipoIdenBen = vlTB!Cod_TipoIdenBen
                stTTMPLiquidacion.Num_IdenBen = vlTB!Num_IdenBen
                stTTMPLiquidacion.Gls_NomBen = vlTB!Gls_NomBen & " " & IIf(IsNull(vlTB!Gls_NomSegBen), "", vlTB!Gls_NomSegBen & " ") & vlTB!Gls_PatBen & IIf(IsNull(vlTB!Gls_MatBen), "", " " & vlTB!Gls_MatBen)
                'para primeros pagos
                stTTMPLiquidacion.Mto_LiqPagar = 0
                stTTMPLiquidacion.Num_Cargas = 0 'vlTB!Num_Cargas
                stTTMPLiquidacion.Mto_LiqHaber = 0
                stTTMPLiquidacion.Mto_LiqDescuento = 0
                stTTMPLiquidacion.gls_vejez = vlTB!COD_VEJEZ
                stTTMPLiquidacion.Cod_TipoIdenTit = vlTB!TIPIDENTIT
                stTTMPLiquidacion.Num_IdenTit = vlTB!NUMIDENTIT
                stTTMPLiquidacion.gls_nomTit = vlTB!NOMTIT
                
            
                'fin primeros pagos
                'Obtiene Fecha de Término del Poder Notarial
                If stTTMPLiquidacion.Cod_TipReceptor <> "R" Then
                    vlSql = "SELECT tut.fec_terpodnot FROM pp_tmae_tutor tut"
                    vlSql = vlSql & " WHERE tut.num_poliza = '" & stTTMPLiquidacion.num_poliza & "'"
                    vlSql = vlSql & " AND tut.num_orden = " & stTTMPLiquidacion.Num_Orden
                    vlSql = vlSql & " AND tut.cod_tipoidentut = " & vlTB!Cod_TipoIdenReceptor & " "
                    vlSql = vlSql & " AND tut.num_identut = '" & vlTB!Num_IdenReceptor & "' "
                    
                    Set vlTB2 = vgConexionBD.Execute(vlSql)
                    If Not vlTB2.EOF Then
                        stTTMPLiquidacion.Fec_TerPodNot = vlTB2!Fec_TerPodNot
                        stTTMPLiquidacion.Fec_TerPodNot = DateSerial(Mid(stTTMPLiquidacion.Fec_TerPodNot, 1, 4), Mid(stTTMPLiquidacion.Fec_TerPodNot, 5, 2), Mid(stTTMPLiquidacion.Fec_TerPodNot, 7, 2))
                    Else
                        stTTMPLiquidacion.Fec_TerPodNot = ""
                    End If
                    'Obtiene Mensajes al Beneficiario
                    stTTMPLiquidacion.Gls_Mensaje = ""
                    vlSql = "SELECT par.gls_mensaje FROM pp_tmae_menpoliza men, pp_tpar_mensaje par"
                    vlSql = vlSql & " WHERE par.cod_mensaje = men.cod_mensaje"
                    vlSql = vlSql & " AND men.num_poliza = '" & stTTMPLiquidacion.num_poliza & "'"
                    vlSql = vlSql & " AND men.num_orden = " & stTTMPLiquidacion.Num_Orden
                    vlSql = vlSql & " AND men.num_perpago = '" & stTTMPLiquidacion.Num_PerPago & "'"
                    Set vlTB2 = vgConexionBD.Execute(vlSql)
                    If Not vlTB2.EOF Then
                        Do While Not vlTB2.EOF
                            stTTMPLiquidacion.Gls_Mensaje = stTTMPLiquidacion.Gls_Mensaje & vlTB2!Gls_Mensaje & Chr(13)
                            vlTB2.MoveNext
                        Loop
                    Else
                        stTTMPLiquidacion.Gls_Mensaje = ""
                    End If
                Else
                    stTTMPLiquidacion.Fec_TerPodNot = ""
                    stTTMPLiquidacion.Gls_Mensaje = ""
                End If
            End If

            stTTMPLiquidacion.Num_PerPago = vlTB!Num_PerPago
            'stTTMPLiquidacion.Mto_LiqPagar = vlTB!Mto_LiqPagar
            stTTMPLiquidacion.Mto_Pension = vlTB!Mto_Pension
            stTTMPLiquidacion.Num_Cargas = 0 'vlTB!Num_Cargas
            'stTTMPLiquidacion.Mto_LiqHaber = vlTB!Mto_Haber
            'stTTMPLiquidacion.Mto_LiqDescuento = vlTB!Mto_Descuento
            If vlPerPago <> stTTMPLiquidacion.Num_PerPago Then
                stTTMPLiquidacion.fec_pago = vlTB!fec_pago
                'Obtiene Fecha del Próximo Pago
'                vlSql = "SELECT pro.fec_pagoproxreg"
'                vlSql = vlSql & " FROM pp_tmae_propagopen pro"
'                vlSql = vlSql & " WHERE pro.num_perpago = '" & stTTMPLiquidacion.Num_PerPago & "'"
'                Set vlTB2 = vgConexionBD.Execute(vlSql)
'                If Not vlTB2.EOF Then
'                    stTTMPLiquidacion.Fec_PagoProxReg = vlTB2!Fec_PagoProxReg
'                    stTTMPLiquidacion.Fec_PagoProxReg = DateSerial(Mid(stTTMPLiquidacion.Fec_PagoProxReg, 1, 4), Mid(stTTMPLiquidacion.Fec_PagoProxReg, 5, 2), Mid(stTTMPLiquidacion.Fec_PagoProxReg, 7, 2))
'                Else
                    stTTMPLiquidacion.Fec_PagoProxReg = ""
'                End If
                'Obtiene Valor UF
'                vlSql = "SELECT val.mto_moneda"
'                vlSql = vlSql & " FROM ma_tval_moneda val"
'                vlSql = vlSql & " WHERE val.cod_moneda = 'UF'"
'                vlSql = vlSql & " AND val.fec_moneda = '" & stTTMPLiquidacion.Fec_Pago & "'"
'                Set vlTB2 = vgConexionBD.Execute(vlSql)
'                If Not vlTB2.EOF Then
'                    stTTMPLiquidacion.Mto_Moneda = vlTB2!Mto_Moneda
'                Else
                    stTTMPLiquidacion.Mto_Moneda = 0
'                End If
                stTTMPLiquidacion.fec_pago = DateSerial(Mid(stTTMPLiquidacion.fec_pago, 1, 4), Mid(stTTMPLiquidacion.fec_pago, 5, 2), Mid(stTTMPLiquidacion.fec_pago, 7, 2))
                vlPerPago = stTTMPLiquidacion.Num_PerPago
            End If
            'Obtiene Tipo de Pensión
            If vlTipPension <> vlTB!Cod_TipPension Then
                vlSql = "SELECT tab.gls_elemento FROM ma_tpar_tabcod tab"
                vlSql = vlSql & " WHERE tab.cod_tabla = 'TP'"
                vlSql = vlSql & " AND tab.cod_elemento = '" & vlTB!Cod_TipPension & "'"
                Set vlTB2 = vgConexionBD.Execute(vlSql)
                If Not vlTB2.EOF Then
                    stTTMPLiquidacion.Gls_TipPension = vlTB2!GLS_ELEMENTO & " " & IIf(vlTB!COD_VEJEZ = "S", "", vlTB!GLS_TP2)
                Else
                    stTTMPLiquidacion.Gls_TipPension = ""
                End If
                vlTipPension = vlTB!Cod_TipPension
            End If
            
            'Obtiene Via de Pago
            If vlViaPago <> vlTB!Cod_ViaPago Then
                vlSql = "SELECT tab.gls_elemento FROM ma_tpar_tabcod tab"
                vlSql = vlSql & " WHERE tab.cod_tabla = 'VPG'"
                vlSql = vlSql & " AND tab.cod_elemento = '" & vlTB!Cod_ViaPago & "'"
                Set vlTB2 = vgConexionBD.Execute(vlSql)
                If Not vlTB2.EOF Then
                    vlDescViaPago = vlTB2!GLS_ELEMENTO
                Else
                    vlDescViaPago = ""
                End If
                vlViaPago = vlTB!Cod_ViaPago
            End If
            
            'hqr 13/10/2007 Obtiene Sucursal de la Via de Pago
            If vlTB!Cod_ViaPago = "04" Then
                If vlTB!Cod_Sucursal <> vlSucursal Then
                    'Obtiene Sucursal
                    stTTMPLiquidacion.Gls_ViaPago = vlDescViaPago
                    vlSql = "SELECT a.gls_sucursal FROM ma_tpar_sucursal a"
                    vlSql = vlSql & " WHERE a.cod_sucursal = '" & vlTB!Cod_Sucursal & "'"
                    vlSql = vlSql & " AND a.cod_tipo = 'A'" 'AFP
                    Set vlTB2 = vgConexionBD.Execute(vlSql)
                    If Not vlTB2.EOF Then
                        vlDescSucursal = vlTB2!gls_sucursal
                    End If
                    vlSucursal = vlTB!Cod_Sucursal
                End If
                stTTMPLiquidacion.Gls_ViaPago = Mid(vlDescViaPago & " - " & vlDescSucursal, 1, 50)
            Else
                stTTMPLiquidacion.Gls_ViaPago = vlDescViaPago
            End If
                        
            stTTMPLiquidacion.Gls_CajaComp = ""
            
            'Obtiene Institución de Salud
            If Not IsNull(vlTB!Cod_InsSalud) Then
                If vlTB!Cod_InsSalud <> "NULL" Then
                    If vlInsSalud <> vlTB!Cod_InsSalud Then
                        vlSql = "SELECT tab.gls_elemento FROM ma_tpar_tabcod tab"
                        vlSql = vlSql & " WHERE tab.cod_tabla = 'IS'"
                        vlSql = vlSql & " AND tab.cod_elemento = '" & vlTB!Cod_InsSalud & "'"
                        Set vlTB2 = vgConexionBD.Execute(vlSql)
                        If Not vlTB2.EOF Then
                            stTTMPLiquidacion.Gls_InsSalud = vlTB2!GLS_ELEMENTO
                        Else
                            stTTMPLiquidacion.Gls_InsSalud = ""
                        End If
                        vlInsSalud = vlTB!Cod_InsSalud
                    End If
                Else
                    vlInsSalud = ""
                    stTTMPLiquidacion.Gls_InsSalud = ""
                End If
            Else
                vlInsSalud = ""
                stTTMPLiquidacion.Gls_InsSalud = ""
            End If
            
            'Obtiene AFP
            If vlAfp <> vlTB!cod_afp Then
                vlSql = "SELECT tab.gls_elemento FROM ma_tpar_tabcod tab"
                vlSql = vlSql & " WHERE tab.cod_tabla = 'AF'"
                vlSql = vlSql & " AND tab.cod_elemento = '" & vlTB!cod_afp & "'"
                Set vlTB2 = vgConexionBD.Execute(vlSql)
                If Not vlTB2.EOF Then
                    stTTMPLiquidacion.Gls_Afp = "AFP " & vlTB2!GLS_ELEMENTO
                Else
                    stTTMPLiquidacion.Gls_Afp = "AFP"
                End If
                vlAfp = vlTB!cod_afp
            End If
                        
            'Obtiene Direccion
            If stTTMPLiquidacion.Cod_Direccion <> vlCodDireccion Then
                If stTTMPLiquidacion.Cod_Direccion <> 0 Then
                    vlSql = "SELECT com.gls_comuna, prov.gls_provincia, reg.gls_region"
                    vlSql = vlSql & " FROM ma_tpar_comuna com, ma_tpar_provincia prov, ma_tpar_region reg"
                    vlSql = vlSql & " WHERE reg.cod_region = prov.cod_region"
                    vlSql = vlSql & " AND prov.cod_region = com.cod_region"
                    vlSql = vlSql & " AND prov.cod_provincia = com.cod_provincia"
                    vlSql = vlSql & " AND com.cod_direccion = '" & vlTB!COD_DIRENUE & "'"
                    Set vlTB2 = vgConexionBD.Execute(vlSql)
                    If Not vlTB2.EOF Then
                        stTTMPLiquidacion.Gls_Direccion2 = vlTB2!gls_region & " - " & vlTB2!gls_provincia & " - " & vlTB2!gls_comuna
                    Else
                        stTTMPLiquidacion.Gls_Direccion2 = ""
                    End If
                Else
                    stTTMPLiquidacion.Gls_Direccion2 = ""
                End If
                vlCodDireccion = stTTMPLiquidacion.Cod_Direccion
            End If
            
            'obtiene vejes rrr 02/11/2012
                        
            If Not IsNull(stTTMPLiquidacion.gls_vejez) Then
                vlSql = "SELECT tab.gls_elemento FROM ma_tpar_tabcod tab"
                vlSql = vlSql & " WHERE tab.cod_tabla = 'TV'"
                vlSql = vlSql & " AND tab.cod_elemento = '" & vlTB!COD_VEJEZ & "'"
                Set vlTB2 = vgConexionBD.Execute(vlSql)
                If Not vlTB2.EOF Then
                    stTTMPLiquidacion.gls_vejez = vlTB2!GLS_ELEMENTO
                Else
                    stTTMPLiquidacion.gls_vejez = ""
                End If
            End If
            
            'stTTMPLiquidacion.Mto_LiqHaber = 0
            'stTTMPLiquidacion.Mto_LiqPagar = 0
            'Obtiene Datos
            stTTMPLiquidacion.Cod_Moneda = vlTB!COD_SCOMP
            If vlPago = "R" Then
                stTTMPLiquidacion.Mto_LiqPagar = vlTB!Mto_LiqPagar
                stTTMPLiquidacion.Mto_LiqHaber = vlTB!Mto_Haber
                stTTMPLiquidacion.Mto_LiqDescuento = vlTB!Mto_Descuento
                stTTMPLiquidacion.Gls_MontoPension = fgConvierteNumeroLetras(vlTB!Mto_LiqPagar, vlTB!Moneda)
            End If
            If vlTB!cod_tipmov = "H" Then 'haber
                If vlPago <> "R" Then
                    stTTMPLiquidacion.Mto_LiqHaber = stTTMPLiquidacion.Mto_LiqHaber + vlTB!Mto_ConHabDes
                    stTTMPLiquidacion.Mto_LiqPagar = stTTMPLiquidacion.Mto_LiqPagar + vlTB!Mto_ConHabDes 'RRR 21/11/2013
                    stTTMPLiquidacion.Gls_MontoPension = fgConvierteNumeroLetras(stTTMPLiquidacion.Mto_LiqPagar, vlTB!Moneda)
                End If
                stTTMPLiquidacion.Cod_ConDescto = "NULL"
                stTTMPLiquidacion.Mto_Descuento = 0
                stTTMPLiquidacion.Cod_ConHaber = vlTB!gls_ConHabDes
                stTTMPLiquidacion.Mto_Haber = vlTB!Mto_ConHabDes
                vlNumConceptosHab = vlNumConceptosHab + 1
                stTTMPLiquidacion.Num_Item = vlNumConceptosHab
                If vlNumConceptosHab > vlNumConceptosDesc Then
                    Call fgInsertaTTMPLiquidacion(stTTMPLiquidacion)
                Else
                    Call fgActualizaTTMPLiquidacionHab(stTTMPLiquidacion)
                End If
            ElseIf vlTB!cod_tipmov = "D" Then 'descuento
                If vlPago <> "R" Then
                    stTTMPLiquidacion.Mto_LiqDescuento = stTTMPLiquidacion.Mto_LiqDescuento + vlTB!Mto_ConHabDes
                    stTTMPLiquidacion.Mto_LiqPagar = stTTMPLiquidacion.Mto_LiqPagar - vlTB!Mto_ConHabDes 'RRR 21/11/2013
                    stTTMPLiquidacion.Gls_MontoPension = fgConvierteNumeroLetras(stTTMPLiquidacion.Mto_LiqPagar, vlTB!Moneda)
                End If
                stTTMPLiquidacion.Cod_ConDescto = vlTB!gls_ConHabDes
                stTTMPLiquidacion.Mto_Descuento = vlTB!Mto_ConHabDes
                stTTMPLiquidacion.Cod_ConHaber = "NULL"
                stTTMPLiquidacion.Mto_Haber = 0
                vlNumConceptosDesc = vlNumConceptosDesc + 1
                stTTMPLiquidacion.Num_Item = vlNumConceptosDesc
                If vlNumConceptosDesc > vlNumConceptosHab Then
                    Call fgInsertaTTMPLiquidacion(stTTMPLiquidacion)
                Else
                    Call fgActualizaTTMPLiquidacionDesc(stTTMPLiquidacion)
                End If
            Else 'OTROS
                stTTMPLiquidacion.Cod_ConDescto = vlTB!Cod_ConHabDes
                stTTMPLiquidacion.Mto_Haber = vlTB!Mto_ConHabDes
                stTTMPLiquidacion.Num_Item = 0
                'Call fgInsertaTTMPLiquidacion(stTTMPLiquidacion)
            End If
            vlFecPago = vlTB!fec_pago
            vlItem = vlItem + 1
            vlTB.MoveNext
        Loop
    Else
        MsgBox "No existe Información para este rango de Fechas", vbInformation, "Operacion Cancelada"
        Exit Function
    End If
    flLlenaTemporal = True
End Function

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

Private Sub Cmb_PenNumIdent_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        If (Txt_PenNumIdent.Enabled = True) Then
            Txt_PenNumIdent.SetFocus
        Else
            Cmd_BuscarPol.SetFocus
        End If
    End If
End Sub

'**************************************************************

Private Sub Cmd_Buscar_Click()
On Error GoTo Err_CmdBuscarClick

    Frm_Busqueda.flInicio ("Frm_Consulta")
    
Exit Sub
Err_CmdBuscarClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_BuscarPol_Click()
On Error GoTo Err_CmdBuscarPolClick
        
    If Txt_PenPoliza.Text = "" Then
       If ((Trim(Cmb_PenNumIdent.Text)) = "") Or (Txt_PenNumIdent.Text = "") Then
       ''*Or _(Not ValiRut(Txt_PenRut.Text, Txt_PenDigito.Text))
           MsgBox "Debe Ingresar el Número de Póliza o la Identificación del Pensionado.", vbCritical, "Error de Datos"
           Txt_PenPoliza.SetFocus
           Exit Sub
       Else
           ''Txt_PenRut = Format(Txt_PenRut, "##,###,##0")
           Txt_PenNumIdent = UCase(Trim(Txt_PenNumIdent))
           ''Txt_PenDigito.SetFocus
           ''vlRutAux = Format(Txt_PenRut, "#0")
       End If
    Else
        Txt_PenPoliza.Text = Trim(Txt_PenPoliza.Text)
    End If
    
    vlCodTipoIdenBenCau = fgObtenerCodigo_TextoCompuesto(Cmb_PenNumIdent)
    vlNumIdenBenCau = Txt_PenNumIdent
    
    vgPalabra = ""
    'Seleccionar beneficiario, según número de póliza y rut de beneficiario.
    If (Txt_PenPoliza.Text <> "") And (Cmb_PenNumIdent.Text <> "") And (Txt_PenNumIdent.Text <> "") Then
        vgPalabra = "num_poliza = '" & Txt_PenPoliza.Text & "' AND "
        vgPalabra = vgPalabra & "cod_tipoidenBEN = " & (vlCodTipoIdenBenCau) & " and "
        vgPalabra = vgPalabra & "num_idenben = '" & (vlNumIdenBenCau) & "' "
    Else
        'Seleccionar, según número de póliza, el primer beneficiario con derecho a pensión.
        'En caso de no existir, seleccionar sólo el primer beneficiario sin derecho.
        If Txt_PenPoliza.Text <> "" Then
           vgSql = ""
           vgSql = "SELECT COUNT(num_orden) as NumeroBen "
           vgSql = vgSql & "FROM PP_TMAE_BEN WHERE "
           vgSql = vgSql & "num_poliza = '" & Txt_PenPoliza.Text & "' AND "
           vgSql = vgSql & "cod_estpension <> '" & clCodSinDerPen & "' "
           vgSql = vgSql & "ORDER BY num_endoso DESC, num_orden ASC "
           Set vgRegistro = vgConexionBD.Execute(vgSql)
           If (vgRegistro!numeroben) <> 0 Then
              vgPalabra = "num_poliza = '" & Txt_PenPoliza.Text & "' AND "
              vgPalabra = vgPalabra & "cod_estpension <> '" & clCodSinDerPen & "' "
           Else
               vgPalabra = "num_poliza = '" & Txt_PenPoliza.Text & "' "
           End If
        Else
            'Seleccionar beneficiario, según rut beneficiario. (Datos de primera póliza encontrada.)
            If Txt_PenNumIdent.Text <> "" Then
               ''vlRutAux = Format(Txt_PenNumIdent, "#0")
               vgPalabra = "cod_tipoidenBEN = " & (vlCodTipoIdenBenCau) & " "
               vgPalabra = vgPalabra & "AND num_idenben = '" & (vlNumIdenBenCau) & "' "
            End If
        End If
    End If
    
    'Ejecutar selección según los parámetros correspondientes, contenidos en
    'variable vgpalabra
    vgSql = ""
    vgSql = "SELECT num_endoso,num_orden,gls_nomben,gls_nomsegben,gls_patben,gls_matben, "
    vgSql = vgSql & "cod_estpension,cod_tipoidenben,num_idenben,num_poliza "
    vgSql = vgSql & "FROM PP_TMAE_BEN WHERE "
    vgSql = vgSql & vgPalabra
    vgSql = vgSql & " ORDER BY num_endoso DESC,num_orden ASC "
    Set vgRs2 = vgConexionBD.Execute(vgSql)
    If Not vgRs2.EOF Then
     
       If Trim(vgRs2!Cod_EstPension) = Trim(clCodSinDerPen) Then
          MsgBox " El Beneficiario Seleccionado No Tiene Derecho a Pensión " & Chr(13) & _
                 "          Sólo podrá Consultar los Datos del Registro", vbInformation, "Información"

          'Desactivar Todos los Controles del Formulario
          

       Else
        
            If Txt_PenPoliza.Text <> "" Then
                vlCodTipoIdenBenCau = vgRs2!Cod_TipoIdenBen
                vlNumIdenBenCau = Trim(vgRs2!Num_IdenBen)
                     
                Txt_PenPoliza.Text = Trim(vgRs2!num_poliza)
                Call fgBuscarPosicionCodigoCombo(vlCodTipoIdenBenCau, Cmb_PenNumIdent)
                Txt_PenNumIdent.Text = vlNumIdenBenCau
            Else
                Txt_PenPoliza.Text = Trim(vgRs2!num_poliza)
            End If
                 
           If Not fgValidaVigenciaPoliza(Trim(Txt_PenPoliza.Text), fgBuscaFecServ) Then
              MsgBox " La Póliza Ingresada no se Encuentra Vigente en el Sistema " & Chr(13) & _
                     "      Sólo podrá Consultar los Datos del Registro", vbInformation, "Información"
              
              'Desactivar Todos los Controles del Formulario
              

           Else
'               Fra_Grupo.Enabled = True
'               SSTab_Consulta.Enabled = True
'               SSTab_Consulta.Tab = 0
           End If
          
       End If
              
       
       If Txt_PenPoliza.Text <> "" Then
            vlCodTipoIdenBenCau = vgRs2!Cod_TipoIdenBen
            vlNumIdenBenCau = Trim(vgRs2!Num_IdenBen)
                 
            Txt_PenPoliza.Text = Trim(vgRs2!num_poliza)
            Call fgBuscarPosicionCodigoCombo(vlCodTipoIdenBenCau, Cmb_PenNumIdent)
            Txt_PenNumIdent.Text = vlNumIdenBenCau
       Else
           Txt_PenPoliza.Text = Trim(vgRs2!num_poliza)
       End If
              
        vlNombreSeg = IIf(IsNull(vgRs2!Gls_NomSegBen), "", (vgRs2!Gls_NomSegBen))
        vlApMaterno = IIf(IsNull(vgRs2!Gls_MatBen), "", (vgRs2!Gls_MatBen))
       
       Lbl_PenNombre.Caption = fgFormarNombreCompleto(Trim(vgRs2!Gls_NomBen), vlNombreSeg, Trim(vgRs2!Gls_PatBen), vlApMaterno)
       ''*Lbl_PenNombre.Caption = Trim(vgRs2!Gls_NomBen) + " " + Trim(vgRs2!Gls_PatBen) + " " + Trim(vgRs2!Gls_MatBen)
       Lbl_End.Caption = (vgRs2!num_endoso)
       txt_End.Text = (vgRs2!num_endoso)
       'GCP TOMA EL NUMERO DE ENDOSO
       vlNumEndoso = (vgRs2!num_endoso)
       vlNumOrden = (vgRs2!Num_Orden)
       
       
       vgSql = ""
       vgSql = " select a.num_idencor, gls_nomcor || ' ' || gls_patcor || ' ' || gls_matcor as nombres"
       vgSql = vgSql & " from pd_tmae_poliza a"
       vgSql = vgSql & " join pt_tmae_corredor b on a.num_idencor=b.num_idencor"
       vgSql = vgSql & " where num_poliza='" & Txt_PenPoliza & "'"
       Set vgRs2 = vgConexionBD.Execute(vgSql)
       If Not vgRs2.EOF Then
             lblNomAsesor.Caption = vgRs2!num_idencor & " - " & vgRs2!nombres
       End If
      
'       Call flCargaGrillaAsigFam(vlNumOrden)
       Call flCargaGrillaGrupo
       Call flMostrarDatosPoliza
       Call flMostrarDatosGrupo(1)
       Call flMostrarDatosPensiones(1)
       ''*Call flMostrarDatosOtrosBeneficios(vlNumOrden)
       
       Call flHabilitarConsulta
       
       SSTab_Consulta.SetFocus
       
    Else
        MsgBox "El Beneficiario o la Póliza Ingresados, No Existen en la Base de Datos", vbInformation, "Información"
        Txt_PenPoliza.SetFocus
        Exit Sub
    End If
'    vgRs2.Close
       
Exit Sub
Err_CmdBuscarPolClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub

Private Sub Cmd_Cancelar2_Click()

    Call flLimpiar
    Call flCargarListaConHabDes
    Call flDeshabilitarConsulta
    Txt_PenPoliza.SetFocus

End Sub

Private Sub Cmd_Imprimir_Click()
On Error GoTo errImprimir
    
    Screen.MousePointer = 11
    
    'Permite imprimir la Opción Indicada a través del Menú
    
    If Opt_CerPensiones.Value = True Then
        Call flImprimirCerPen
    ElseIf Opt_ConPensiones.Value = True Then
        If Opt_ConPensiones.Value = True Then
            Call FlImprimeConstan
        End If
    'MARCO-----09/03/2010
    ElseIf Opt_CertSupervivencia.Value = True Then
        Call Imprime_Cert_Sobrevivencia
    End If
    
    Screen.MousePointer = 0

Exit Sub
errImprimir:
    Screen.MousePointer = 0
    If Err.Number <> 0 Then
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End If
End Sub
Private Sub Imprime_Cert_Sobrevivencia()

Dim fechini As String
Dim fechfin As String
Dim fechcre As String

Dim cadena As String
Dim d As String
Dim m As String
Dim y As String
Dim x As Integer

Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

Dim vlArchivo As String
Dim r_temp As ADODB.Recordset

On Error GoTo Errores1
   
   
Screen.MousePointer = 11

'lleno variables necesarias para enviarlo al reporte marco 09/03/2010

cadena = "SELECT c.NUM_POLIZA,c.NUM_ENDOSO,c.NUM_ORDEN,c.fec_inicer,c.fec_tercer, c.COD_FRECUENCIA,c.GLS_NOMINSTITUCION," & _
            "c.FEC_RECCIA,c.FEC_INGCIA, c.FEC_EFECTO  FROM pp_tmae_certificado C " & _
            "Where  c.NUM_POLIZA = '" & Trim(Txt_PenPoliza.Text) & "' AND " & _
            "fec_inicer=(select max(x.fec_inicer) from pp_tmae_certificado x where x.NUM_POLIZA=c.NUM_POLIZA) " & _
            "ORDER BY C.fec_inicer DESC"
            
Set rs = vgConexionBD.Execute(cadena)

If Not rs.EOF Then
    fechini = rs("fec_inicer")
    fechfin = rs("fec_tercer")
    fechcre = rs("FEC_RECCIA")
    
    y = Mid(fechini, 1, 4)
    m = Mid(fechini, 5, 2)
    d = Mid(fechini, 7, 2)
    fechini = DateSerial(y, m, d)
    y = Mid(fechfin, 1, 4)
    m = Mid(fechfin, 5, 2)
    d = Mid(fechfin, 7, 2)
    fechfin = DateSerial(y, m, d)
    y = Mid(fechcre, 1, 4)
    m = Mid(fechcre, 5, 2)
    d = Mid(fechcre, 7, 2)
    fechcre = DateSerial(y, m, d)
End If


If Mid(Trim(Msf_GrillaGrupo.TextMatrix(Msf_GrillaGrupo.row, 1)), 1, 2) = "99" Then
    Nombre_Afiliado = Msf_GrillaGrupo.TextMatrix(Msf_GrillaGrupo.row, 4)
    Tipo_num_documento_afiliado = Trim(Mid(Lbl_GrupTipoIdent, InStr(1, Lbl_GrupTipoIdent, "-") + 1, Len(Lbl_GrupTipoIdent))) & " - " & Lbl_GrupNumIdent

    If Msf_GrillaGrupo.TextMatrix(Msf_GrillaGrupo.row, 5) = "99" And Mid(Trim(Msf_GrillaGrupo.TextMatrix(Msf_GrillaGrupo.row, 1)), 1, 2) <> "99" Then
        Nombre_Beneficiario = Msf_GrillaGrupo.TextMatrix(Msf_GrillaGrupo.row, 4)
    Else
        Nombre_Beneficiario = ""
    End If
    
Else
    For x = 0 To Msf_GrillaGrupo.rows - 1
        If Mid(Trim(Msf_GrillaGrupo.TextMatrix(x, 1)), 1, 2) = "99" Then
            Nombre_Afiliado = Msf_GrillaGrupo.TextMatrix(x, 4)
            Tipo_num_documento_afiliado = Trim(Mid(Msf_GrillaGrupo.TextMatrix(x, 2), InStr(1, Msf_GrillaGrupo.TextMatrix(x, 2), "-") + 1, Len(Msf_GrillaGrupo.TextMatrix(x, 2)))) & " - " & Msf_GrillaGrupo.TextMatrix(x, 3)
        End If
    Next
    
    If Msf_GrillaGrupo.TextMatrix(Msf_GrillaGrupo.row, 5) = "99" And Mid(Trim(Msf_GrillaGrupo.TextMatrix(Msf_GrillaGrupo.row, 1)), 1, 2) <> "99" Then
        Nombre_Beneficiario = Msf_GrillaGrupo.TextMatrix(Msf_GrillaGrupo.row, 4)
    Else
        Nombre_Beneficiario = ""
    End If
    
End If

Fecha_Creacion = fechcre
If Mid(Trim(Msf_GrillaGrupo.TextMatrix(Msf_GrillaGrupo.row, 1)), 1, 2) = "99" Then
    Ident_1 = "x"
    Ident_2 = ""
Else
    Ident_1 = ""
    Ident_2 = "x"
End If

num_poliza = Trim(Txt_PenPoliza.Text)
RangoFecha = fechini & " al " & fechfin

If Mid(Trim(Lbl_PolTipPen.Caption), 1, 2) = "04" Or Mid(Trim(Lbl_PolTipPen.Caption), 1, 2) = "05" Then
    Tipo_J = "x"
    Tipo_S = ""
    Tipo_I = ""
ElseIf Mid(Trim(Lbl_PolTipPen.Caption), 1, 2) = "08" Or Mid(Trim(Lbl_PolTipPen.Caption), 1, 2) = "09" Or Mid(Trim(Lbl_PolTipPen.Caption), 1, 2) = "10" Or Mid(Trim(Lbl_PolTipPen.Caption), 1, 2) = "11" Or Mid(Trim(Lbl_PolTipPen.Caption), 1, 2) = "12" Then
    Tipo_S = "x"
    Tipo_I = ""
    Tipo_J = ""
ElseIf Mid(Trim(Lbl_PolTipPen.Caption), 1, 2) = "06" Or Mid(Trim(Lbl_PolTipPen.Caption), 1, 2) = "07" Then
    Tipo_I = "x"
    Tipo_J = ""
    Tipo_S = ""
End If


If Len(Nombre_Beneficiario) > 0 Then
    Tipo_num_documento_beneficiario = Trim(Mid(Cmb_PenNumIdent.Text, InStr(1, Cmb_PenNumIdent.Text, "-") + 1, Len(Cmb_PenNumIdent.Text))) & " - " & Txt_PenNumIdent.Text
Else
    Tipo_num_documento_beneficiario = ""
End If


'llamo reporte

   vlArchivo = strRpt & "PP_Rpt_CertificadoSuperv.rpt"   '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
       MsgBox "El reporte no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
       Screen.MousePointer = 0
   End If
   
   Screen.MousePointer = 0
     
rs.Close
sName_Reporte = "PP_Rpt_CertificadoSuperv.rpt"
frm_plantilla.Show 1
   
Errores1:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " " & "Numero de Error : " & Err.Number & Err.HelpFile & Err.HelpContext, vbExclamation, "Mensaje de error ..."
        Screen.MousePointer = 0
    End If
    
    
End Sub


Private Sub Cmd_LiqConsultaHD_Click()
    
    If (Trim(Txt_LiqFecIni) = "") Then
       MsgBox "Debe ingresar una Fecha de Inicio de Vigencia", vbCritical, "Error de Datos"
       Txt_LiqFecIni.SetFocus
       Exit Sub
    End If
    If Not IsDate(Txt_LiqFecIni.Text) Then
       MsgBox "La Fecha Ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
       Txt_LiqFecIni.SetFocus
       Exit Sub
    End If
    If (CDate(Txt_LiqFecIni) > CDate(Date)) Then
       MsgBox "La Fecha Ingresada es Mayor a la Fecha Actual", vbCritical, "Error de Datos"
       Txt_LiqFecIni.SetFocus
       Exit Sub
    End If
    If (Year(Txt_LiqFecIni) < 1900) Then
       MsgBox "La Fecha Ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
       Txt_LiqFecIni.SetFocus
       Exit Sub
    End If
    
    Txt_LiqFecIni.Text = Format(CDate(Trim(Txt_LiqFecIni)), "yyyymmdd")
    Txt_LiqFecIni.Text = DateSerial(Mid((Txt_LiqFecIni.Text), 1, 4), Mid((Txt_LiqFecIni.Text), 5, 2), Mid((Txt_LiqFecIni.Text), 7, 2))
    
    'Valida Vigencia de la poliza según fecha ingresafda en Inicio de Vigencia
    If Not fgValidaVigenciaPoliza(Trim(Txt_PenPoliza.Text), Txt_LiqFecIni.Text) Then
       MsgBox " La Fecha Ingresada es Anterior a la Fecha de Vigencia de la Póliza ", vbInformation, "Información"
       Txt_LiqFecIni.SetFocus
       Exit Sub
    End If
    
'Valida Fecha de Término de Periodo Ingresada
    If (Trim(Txt_LiqFecTer) = "") Then
       MsgBox "Debe ingresar una Fecha de Término de Vigencia", vbCritical, "Error de Datos"
       Txt_LiqFecIni.SetFocus
       Exit Sub
    End If
    If Not IsDate(Txt_LiqFecTer.Text) Then
       MsgBox "La Fecha Ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
       Txt_LiqFecTer.SetFocus
       Exit Sub
    End If
    If (CDate(Txt_LiqFecTer) > CDate(Date)) Then
       MsgBox "La Fecha Ingresada es Mayor a la Fecha Actual", vbCritical, "Error de Datos"
       Txt_LiqFecTer.SetFocus
       Exit Sub
    End If
    If (Year(Txt_LiqFecTer) < 1900) Then
       MsgBox "La Fecha Ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
       Txt_LiqFecTer.SetFocus
       Exit Sub
    End If
        
    Txt_LiqFecTer.Text = Format(CDate(Trim(Txt_LiqFecTer)), "yyyymmdd")
    Txt_LiqFecTer.Text = DateSerial(Mid((Txt_LiqFecTer.Text), 1, 4), Mid((Txt_LiqFecTer.Text), 5, 2), Mid((Txt_LiqFecTer.Text), 7, 2))
        
'    If CDate(Trim(Txt_LiqFecIni)) < CDate(Trim(Txt_LiqFecTer)) Then
'       MsgBox "La Fecha de Término es Inferior a la Fecha de Inicio", vbCritical, "Error de Datos"
'       Txt_LiqFecTer.SetFocus
'       Exit Sub
'    End If
        
        'Valida Vigencia de la poliza según fecha ingresada en termino de Vigencia
    If Not fgValidaVigenciaPoliza(Trim(Txt_PenPoliza.Text), Txt_LiqFecTer.Text) Then
       MsgBox " La Fecha Ingresada es Anterior a la Fecha de Vigencia de la Póliza ", vbInformation, "Información"
       Txt_LiqFecTer.SetFocus
       Exit Sub
    End If
    
'-------------------------------------------
    
    vlFechaIni = Format(CDate(Trim(Txt_LiqFecIni)), "yyyymmdd")
    vlFechaTer = Format(CDate(Trim(Txt_LiqFecTer)), "yyyymmdd")
    
    vlCodConceptos = ""
    vgQuery = ""
    If Trim(Lst_LiqSeleccion.Text) = "TODOS" Then
        vgQuery = ""
    Else
        Call flObtenerConceptos
        vgQuery = "AND p.cod_conhabdes IN " & vlCodConceptos & " "
    End If
    
    Call flCargaGrillaHabDes(Lbl_GrupNumOrden)
    

End Sub

Private Sub Cmd_LiqHDImprimir_Click()

    If (Trim(Txt_LiqFecIni) = "") Then
       MsgBox "Debe ingresar una Fecha de Inicio de Vigencia", vbCritical, "Error de Datos"
       Txt_LiqFecIni.SetFocus
       Exit Sub
    End If
    If Not IsDate(Txt_LiqFecIni.Text) Then
       MsgBox "La Fecha Ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
       Txt_LiqFecIni.SetFocus
       Exit Sub
    End If
    If (CDate(Txt_LiqFecIni) > CDate(Date)) Then
       MsgBox "La Fecha Ingresada es Mayor a la Fecha Actual", vbCritical, "Error de Datos"
       Txt_LiqFecIni.SetFocus
       Exit Sub
    End If
    If (Year(Txt_LiqFecIni) < 1900) Then
       MsgBox "La Fecha Ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
       Txt_LiqFecIni.SetFocus
       Exit Sub
    End If
    
    Txt_LiqFecIni.Text = Format(CDate(Trim(Txt_LiqFecIni)), "yyyymmdd")
    Txt_LiqFecIni.Text = DateSerial(Mid((Txt_LiqFecIni.Text), 1, 4), Mid((Txt_LiqFecIni.Text), 5, 2), Mid((Txt_LiqFecIni.Text), 7, 2))
    
    'Valida Vigencia de la poliza según fecha ingresafda en Inicio de Vigencia
    If Not fgValidaVigenciaPoliza(Trim(Txt_PenPoliza.Text), Txt_LiqFecIni.Text) Then
       MsgBox " La Fecha Ingresada es Anterior a la Fecha de Vigencia de la Póliza ", vbInformation, "Información"
       Txt_LiqFecIni.SetFocus
       Exit Sub
    End If
    
'Valida Fecha de Término de Periodo Ingresada
    If (Trim(Txt_LiqFecTer) = "") Then
       MsgBox "Debe ingresar una Fecha de Término de Vigencia", vbCritical, "Error de Datos"
       Txt_LiqFecIni.SetFocus
       Exit Sub
    End If
    If Not IsDate(Txt_LiqFecTer.Text) Then
       MsgBox "La Fecha Ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
       Txt_LiqFecTer.SetFocus
       Exit Sub
    End If
    If (CDate(Txt_LiqFecTer) > CDate(Date)) Then
       MsgBox "La Fecha Ingresada es Mayor a la Fecha Actual", vbCritical, "Error de Datos"
       Txt_LiqFecTer.SetFocus
       Exit Sub
    End If
    If (Year(Txt_LiqFecTer) < 1900) Then
       MsgBox "La Fecha Ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
       Txt_LiqFecTer.SetFocus
       Exit Sub
    End If
        
    Txt_LiqFecTer.Text = Format(CDate(Trim(Txt_LiqFecTer)), "yyyymmdd")
    Txt_LiqFecTer.Text = DateSerial(Mid((Txt_LiqFecTer.Text), 1, 4), Mid((Txt_LiqFecTer.Text), 5, 2), Mid((Txt_LiqFecTer.Text), 7, 2))
        
'    If CDate(Trim(Txt_LiqFecIni)) < CDate(Trim(Txt_LiqFecTer)) Then
'       MsgBox "La Fecha de Término es Inferior a la Fecha de Inicio", vbCritical, "Error de Datos"
'       Txt_LiqFecTer.SetFocus
'       Exit Sub
'    End If
        
        'Valida Vigencia de la poliza según fecha ingresada en termino de Vigencia
    If Not fgValidaVigenciaPoliza(Trim(Txt_PenPoliza.Text), Txt_LiqFecTer.Text) Then
       MsgBox " La Fecha Ingresada es Anterior a la Fecha de Vigencia de la Póliza ", vbInformation, "Información"
       Txt_LiqFecTer.SetFocus
       Exit Sub
    End If
    
'-------------------------------------------
    
    vlFechaIni = Format(CDate(Trim(Txt_LiqFecIni)), "yyyymmdd")
    vlFechaTer = Format(CDate(Trim(Txt_LiqFecTer)), "yyyymmdd")



     'Verificar si el beneficiario seleccionado tiene derecho a pension
    vgSql = ""
    vgSql = "SELECT num_orden,cod_estpension "
    vgSql = vgSql & "FROM PP_TMAE_BEN "
    vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
    vgSql = vgSql & "num_endoso = " & vlNumEndoso & " AND "
    vgSql = vgSql & "num_orden = " & Trim(Lbl_GrupNumOrden.Caption) & " "
'    vgSql = vgSql & "AND cod_estpension = '" & clCodEstPen99 & "' "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If vgRegistro.EOF Then
       MsgBox "El Beneficiario Seleccionado No Existe", vbInformation, "Información"
       Exit Sub
    End If

    Call flInformeLiqPago(clOpcionDEF)

''    Screen.MousePointer = vbHourglass
''    'Informe de Liquidaciones de Pago
''    vgNombreInformeSeleccionadoInd = "InfCONLiqPago"
''    Frm_PlanillaPensionado.Show
''    Frm_PlanillaPensionado.Caption = "Informe de Liquidación de Pensiones."
''    Screen.MousePointer = vbDefault
End Sub

Private Sub Cmd_LiqHDLimpiar_Click()
Dim i As Integer

    'Limpiar campos de Fechas
    Txt_LiqFecIni = ""
    Txt_LiqFecTer = ""
    'Limpiar Lista de Conceptos de Haberes y Descuentos
    For i = 0 To Lst_LiqSeleccion.ListCount - 1
        Lst_LiqSeleccion.Selected(i) = False
    Next
    'Limpiar Grilla de Conceptos de Haberes y Descuentos
    Call flInicializaGrillaHabDes
    
    Txt_LiqFecIni.SetFocus
    
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

Private Sub cmdNuevo_Click()
'MARCO ----- 09/03/2010

Dim x As Integer
On Error GoTo mierror

frmCertSuperv.Orden = Msf_GrillaGrupo.TextMatrix(Msf_GrillaGrupo.row, 0)
frmCertSuperv.Identificacion = Msf_GrillaGrupo.TextMatrix(Msf_GrillaGrupo.row, 1)
frmCertSuperv.Poliza = Txt_PenPoliza.Text

If Mid(Trim(Msf_GrillaGrupo.TextMatrix(Msf_GrillaGrupo.row, 1)), 1, 2) = "99" Then
    frmCertSuperv.Afiliado = Msf_GrillaGrupo.TextMatrix(Msf_GrillaGrupo.row, 4)
    frmCertSuperv.Doc_afiliado = Lbl_GrupTipoIdent
    frmCertSuperv.NumDoc_afiliado = Lbl_GrupNumIdent
    frmCertSuperv.Doc = Lbl_GrupTipoIdent
    frmCertSuperv.NumDoc = Lbl_GrupNumIdent
Else
    frmCertSuperv.Doc = Lbl_GrupTipoIdent
    frmCertSuperv.NumDoc = Lbl_GrupNumIdent

    For x = 0 To Msf_GrillaGrupo.rows - 1
        If Mid(Trim(Msf_GrillaGrupo.TextMatrix(x, 1)), 1, 2) = "99" Then
            frmCertSuperv.Afiliado = Msf_GrillaGrupo.TextMatrix(x, 4)
            frmCertSuperv.Doc_afiliado = Trim(Msf_GrillaGrupo.TextMatrix(x, 2))
            frmCertSuperv.NumDoc_afiliado = Trim(Msf_GrillaGrupo.TextMatrix(x, 3))
        End If
    Next
End If

frmCertSuperv.Endoso = Lbl_End
frmCertSuperv.TipoPension = Lbl_PolTipPen
If Msf_GrillaGrupo.TextMatrix(Msf_GrillaGrupo.row, 5) = "99" And Mid(Trim(Msf_GrillaGrupo.TextMatrix(Msf_GrillaGrupo.row, 1)), 1, 2) <> "99" Then
    frmCertSuperv.Beneficiado = Msf_GrillaGrupo.TextMatrix(Msf_GrillaGrupo.row, 4)
Else
    frmCertSuperv.Beneficiado = ""
End If

frmCertSuperv.Show 1

Exit Sub
mierror:
    MsgBox "No pudo Cargar", vbInformation
    
End Sub

Private Sub Command1_Click()
Frm_ImprimeTodosPolizas.Show

End Sub

Private Sub cmdPrintSel_Click()
    'Screen.MousePointer = 11
    Frm_ImprimeTodosPolizas.Show
    'Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
On Error GoTo Err_Cargar

    Frm_Consulta.Top = 0
    Frm_Consulta.Left = 0
    
    fgComboTipoIdentificacion Cmb_PenNumIdent
    
    'CMV-20061102 I
    Lbl_MtoPQ.Visible = False
    Lbl_MtoPensionQui.Visible = False
    'CMV-20061102 F
    
    Call flLimpiar
        
    Call flDeshabilitarConsulta
         
    Call flCargarListaConHabDes
    
    SSTab_Consulta.Tab = 0
            
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub

Private Sub Lst_LiqSeleccion_ItemCheck(Item As Integer)

'    If Lst_LiqSeleccion.ListCount = 1 Then
'      Lst_LiqSeleccion.AddItem (" TODOS "), 0
'    End If
'    Dim i As Integer
'    If Item = 0 Then
'        If Lst_LiqSeleccion.Selected(0) Then
'            For i = 0 To Lst_LiqSeleccion.ListCount - 1
'                Lst_LiqSeleccion.Selected(i) = True
'            Next
'        Else
'            For i = 0 To Lst_LiqSeleccion.ListCount - 1
'                Lst_LiqSeleccion.Selected(i) = False
'            Next
'        End If
'    Else
'        Lst_LiqSeleccion.Selected(0) = False
'    End If

    Dim i As Integer
    If Item = 0 Then
        If Lst_LiqSeleccion.Selected(0) Then
            For i = 1 To Lst_LiqSeleccion.ListCount - 1
                Lst_LiqSeleccion.Selected(i) = False
            Next
'        Else
'            For i = 1 To Lst_LiqSeleccion.ListCount - 1
'                Lst_LiqSeleccion.Selected(i) = True
'            Next
        End If
    Else
        Lst_LiqSeleccion.Selected(0) = False
    End If
End Sub

Private Sub Msf_GrillaGrupo_Click()

On Error GoTo Err_MsfGrillaGrupoClick
    
    Msf_GrillaGrupo.Col = 0
    If (Msf_GrillaGrupo.Text = "") Or (Msf_GrillaGrupo.row = 0) Then
        MsgBox "No existen Detalles", vbExclamation, "Información"
        Exit Sub
    Else
        Msf_GrillaGrupo.Col = 0
        'Lbl_NumRetencion.Caption = Msf_GrillaRetJud.Text
        
'        Call flCargaGrillaAsigFam(Msf_GrillaGrupo.Text)
        Call flLimpiarTab
        Call flMostrarDatosPoliza
        Call flMostrarDatosGrupo(Msf_GrillaGrupo.Text)
        Call flMostrarDatosPensiones(Msf_GrillaGrupo.Text)
        ''*Call flMostrarDatosOtrosBeneficios(Msf_GrillaGrupo.Text)
        
        If Msf_GrillaGrupo.TextMatrix(Msf_GrillaGrupo.row, 5) = "99" Then
            Opt_CertSupervivencia.Visible = True
        Else
            Opt_CertSupervivencia.Visible = False
        End If
        
        
        SSTab_Consulta.Tab = 1
        
    End If

Exit Sub
Err_MsfGrillaGrupoClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub


Private Sub Opt_CerPensiones_Click()
    cmdNuevo.Visible = False
End Sub

Private Sub Opt_CertSupervivencia_Click()
    cmdNuevo.Visible = True
End Sub

Private Sub Opt_ConPensiones_Click()
    cmdNuevo.Visible = False
End Sub

Private Sub SSTab_Consulta_Click(PreviousTab As Integer)
    If SSTab_Consulta.Tab = 4 Then
       Txt_LiqFecIni.SetFocus
    End If
End Sub

Private Sub Txt_LiqFecIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If (Trim(Txt_LiqFecIni) = "") Then
           MsgBox "Debe ingresar una Fecha de Inicio de Vigencia", vbCritical, "Error de Datos"
           Txt_LiqFecIni.SetFocus
           Exit Sub
        End If
        If Not IsDate(Txt_LiqFecIni.Text) Then
           MsgBox "La Fecha Ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
           Txt_LiqFecIni.SetFocus
           Exit Sub
        End If
        If (CDate(Txt_LiqFecIni) > CDate(Date)) Then
           MsgBox "La Fecha Ingresada es Mayor a la Fecha Actual", vbCritical, "Error de Datos"
           Txt_LiqFecIni.SetFocus
           Exit Sub
        End If
        If (Year(Txt_LiqFecIni) < 1900) Then
           MsgBox "La Fecha Ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
           Txt_LiqFecIni.SetFocus
           Exit Sub
        End If
        
        Txt_LiqFecIni.Text = Format(CDate(Trim(Txt_LiqFecIni)), "yyyymmdd")
        Txt_LiqFecIni.Text = DateSerial(Mid((Txt_LiqFecIni.Text), 1, 4), Mid((Txt_LiqFecIni.Text), 5, 2), Mid((Txt_LiqFecIni.Text), 7, 2))
        
        'Valida Vigencia de la poliza según fecha ingresafda en Inicio de Vigencia
        If Not fgValidaVigenciaPoliza(Trim(Txt_PenPoliza.Text), Txt_LiqFecIni.Text) Then
           MsgBox " La Fecha Ingresada es Anterior a la Fecha de Vigencia de la Póliza ", vbInformation, "Información"
           Txt_LiqFecIni.SetFocus
           Exit Sub
        End If
         
'        If Not fgValidaPagoPension(Txt_LiqFecIni, Trim(Txt_PenPoliza), vlNumOrden) Then
'           MsgBox " Ya se ha Realizado el Proceso de Cálculo de Pensión para ésta Fecha ", vbCritical, "Operación Cancelada"
'           Txt_FechaIniVig.SetFocus
'           Exit Sub
'        End If
        
        Txt_LiqFecTer.SetFocus
        
    End If

End Sub

Private Sub Txt_LiqFecIni_LostFocus()

    If (Trim(Txt_LiqFecIni) <> "") Then
        If IsDate(Txt_LiqFecIni) Then
            Txt_LiqFecIni.Text = Format(CDate(Trim(Txt_LiqFecIni)), "yyyymmdd")
            Txt_LiqFecIni.Text = DateSerial(Mid((Txt_LiqFecIni.Text), 1, 4), Mid((Txt_LiqFecIni.Text), 5, 2), Mid((Txt_LiqFecIni.Text), 7, 2))
        End If
    End If

End Sub

Private Sub Txt_LiqFecTer_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
        If (Trim(Txt_LiqFecTer) = "") Then
           MsgBox "Debe ingresar una Fecha de Término de Vigencia", vbCritical, "Error de Datos"
           Txt_LiqFecIni.SetFocus
           Exit Sub
        End If
        If Not IsDate(Txt_LiqFecTer.Text) Then
           MsgBox "La Fecha Ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
           Txt_LiqFecTer.SetFocus
           Exit Sub
        End If
        If (CDate(Txt_LiqFecTer) > CDate(Date)) Then
           MsgBox "La Fecha Ingresada es Mayor a la Fecha Actual", vbCritical, "Error de Datos"
           Txt_LiqFecTer.SetFocus
           Exit Sub
        End If
        If (Year(Txt_LiqFecTer) < 1900) Then
           MsgBox "La Fecha Ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
           Txt_LiqFecTer.SetFocus
           Exit Sub
        End If
        
        Txt_LiqFecTer.Text = Format(CDate(Trim(Txt_LiqFecTer)), "yyyymmdd")
        Txt_LiqFecTer.Text = DateSerial(Mid((Txt_LiqFecTer.Text), 1, 4), Mid((Txt_LiqFecTer.Text), 5, 2), Mid((Txt_LiqFecTer.Text), 7, 2))
        
'        If CDate(Trim(Txt_LiqFecIni)) < CDate(Trim(Txt_LiqFecTer)) Then
'           MsgBox "La Fecha de Término es Inferior a la Fecha de Inicio", vbCritical, "Error de Datos"
'           Txt_LiqFecTer.SetFocus
'           Exit Sub
'        End If
        
        'Valida Vigencia de la poliza según fecha ingresada en termino de Vigencia
        If Not fgValidaVigenciaPoliza(Trim(Txt_PenPoliza.Text), Txt_LiqFecTer.Text) Then
           MsgBox " La Fecha Ingresada es Anterior a la Fecha de Vigencia de la Póliza ", vbInformation, "Información"
           Txt_LiqFecTer.SetFocus
           Exit Sub
        End If
         
'        If Not fgValidaPagoPension(Txt_LiqFecIni, Trim(Txt_PenPoliza), vlNumOrden) Then
'           MsgBox " Ya se ha Realizado el Proceso de Cálculo de Pensión para ésta Fecha ", vbCritical, "Operación Cancelada"
'           Txt_FechaIniVig.SetFocus
'           Exit Sub
'        End If
        
        Cmd_LiqConsultaHD.SetFocus
        
    End If

End Sub

Private Sub Txt_LiqFecTer_LostFocus()
    
    If (Trim(Txt_LiqFecTer) <> "") Then
        Txt_LiqFecTer.Text = Format(CDate(Trim(Txt_LiqFecTer)), "yyyymmdd")
        Txt_LiqFecTer.Text = DateSerial(Mid((Txt_LiqFecTer.Text), 1, 4), Mid((Txt_LiqFecTer.Text), 5, 2), Mid((Txt_LiqFecTer.Text), 7, 2))
        
    End If
    
End Sub

Private Sub Txt_PenNumIdent_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If (Trim(Txt_PenNumIdent) <> "") Then
            Txt_PenNumIdent = UCase(Trim(Txt_PenNumIdent))
        End If
        Cmd_BuscarPol.SetFocus
    End If
End Sub

Private Sub txt_pennumident_lostfocus()
    Txt_PenNumIdent = UCase(Trim(Txt_PenNumIdent))
End Sub
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

Private Sub Txt_PenPoliza_LostFocus()
    Txt_PenPoliza = UCase(Trim(Txt_PenPoliza))
    Txt_PenPoliza = Format(Txt_PenPoliza, "0000000000")
End Sub

'CMV-20061102 I
'Funciones para  mostrar el dato de monto de Pensión en Quiebra
Function flUltimoPeriodoCerrado(iNumPol As String) As String
On Error GoTo Err_flUltimoPeriodoCerrado
Dim vlTipoPago As String
'Permite obtener el último periodo que se encuentra cerrado, es decir,
'que tiene calculo definitivo
'Estados del Pago de Pensión
'PP: Primer Pago
'PR: Pago en Regimen
    
    flUltimoPeriodoCerrado = ""
    
    'Determinar si el Caso es Primer Pago o Pago en Regimen
    vgSql = ""
    vgSql = "SELECT num_poliza,num_endoso,num_orden "
    vgSql = vgSql & " FROM pp_tmae_liqpagopendef WHERE "
    vgSql = vgSql & " num_poliza = '" & iNumPol & "' "
    Set vgRs3 = vgConexionBD.Execute(vgSql)
    If vgRs3.EOF Then
        'Pago Régimen
        vlTipoPago = "R"
    Else
        'Primer Pago
        vlTipoPago = "P"
    End If
    vgRs3.Close

    vlTipoPago = "R"

    'Permite obtener el último periodo que se encuentra cerrado
    vgSql = ""
    vgSql = "SELECT p.num_perpago "
    vgSql = vgSql & "FROM pp_tmae_propagopen p "
    vgSql = vgSql & "WHERE "
    If vlTipoPago = "P" Then
        vgSql = vgSql & "p.cod_estadopri = '" & clCodEstadoC & "' "
    Else
        vgSql = vgSql & "p.cod_estadoreg = '" & clCodEstadoC & "' "
    End If
    vgSql = vgSql & "ORDER BY num_perpago DESC"
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
        flUltimoPeriodoCerrado = Trim(vgRegistro!Num_PerPago)
    End If

Exit Function
Err_flUltimoPeriodoCerrado:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flBuscarPensionQuiebra(iNumPerPago As String, iNumPoliza As String, iNumOrden As Integer, _
                            iCodTipReceptor As String) As Double
On Error GoTo Err_flBuscarPensionQuiebra

    flBuscarPensionQuiebra = 0

    vgSql = ""
    vgSql = "SELECT mto_pension "
    vgSql = vgSql & " FROM pp_tmae_liqpagopendef WHERE "
    vgSql = vgSql & " num_perpago = '" & iNumPerPago & "' AND "
    vgSql = vgSql & " num_poliza = '" & iNumPoliza & "' AND "
    vgSql = vgSql & " num_orden = " & iNumOrden & " AND "
    vgSql = vgSql & " cod_tipreceptor <> '" & iCodTipReceptor & "' "
    Set vgRs3 = vgConexionBD.Execute(vgSql)
    If Not vgRs3.EOF Then
        flBuscarPensionQuiebra = (vgRs3!Mto_Pension)
    Else
        flBuscarPensionQuiebra = 0
    End If

Exit Function
Err_flBuscarPensionQuiebra:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function
'CMV-20061102 F


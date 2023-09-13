VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Frm_PensRegistroPagosGar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantención de Pagos a Terceros - Periodo Garantizado."
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   10785
   Begin VB.Frame Fra_DetallePago 
      Height          =   1095
      Left            =   120
      TabIndex        =   96
      Top             =   1080
      Width           =   10575
      Begin VB.TextBox Txt_TasaInteres 
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   10
         Top             =   720
         Width           =   1200
      End
      Begin VB.TextBox Txt_FecPago 
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   9
         Top             =   480
         Width           =   1200
      End
      Begin VB.TextBox Txt_FecRecepcion 
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   8
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Fin Per. Gar."
         Height          =   195
         Index           =   29
         Left            =   6960
         TabIndex        =   118
         Top             =   480
         Width           =   1380
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Ini. Pago"
         Height          =   195
         Index           =   21
         Left            =   3720
         TabIndex        =   117
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cambio"
         Height          =   195
         Index           =   10
         Left            =   3720
         TabIndex        =   116
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Lbl_FecIniPago 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5280
         TabIndex        =   115
         Top             =   722
         Width           =   1215
      End
      Begin VB.Label Lbl_TipoCambio 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5280
         TabIndex        =   114
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Lbl_MtoPensionRef 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   8760
         TabIndex        =   13
         Top             =   722
         Width           =   1215
      End
      Begin VB.Label Lbl_FecIniPerGar 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   8760
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Lbl_FecFinPerGar 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   8760
         TabIndex        =   14
         Top             =   488
         Width           =   1215
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicio Per. Gar."
         Height          =   195
         Index           =   28
         Left            =   6960
         TabIndex        =   104
         Top             =   240
         Width           =   1545
      End
      Begin VB.Label Lbl_FecDevengue 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5280
         TabIndex        =   11
         Top             =   488
         Width           =   1215
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Devengue"
         Height          =   195
         Index           =   27
         Left            =   3720
         TabIndex        =   103
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Pensión Referencia"
         Height          =   195
         Index           =   26
         Left            =   6960
         TabIndex        =   102
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label Lbl_MonPension 
         Caption         =   "(TM)"
         Height          =   285
         Index           =   0
         Left            =   10080
         TabIndex        =   101
         Top             =   720
         Width           =   375
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
         Index           =   22
         Left            =   240
         TabIndex        =   100
         Top             =   0
         Width           =   1560
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Fecha Recepción"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   99
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Tasa Interés Anual"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   98
         Top             =   720
         Width           =   1620
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Fecha Pago"
         Height          =   255
         Index           =   25
         Left            =   240
         TabIndex        =   97
         Top             =   480
         Width           =   1695
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   120
      TabIndex        =   64
      Top             =   3720
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   6588
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Antecedentes Generales"
      TabPicture(0)   =   "Frm_PensRegistroPagosGar.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Fra_BM"
      Tab(0).Control(1)=   "Fra_AntRecep"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Antecedentes del Cálculo"
      TabPicture(1)   =   "Frm_PensRegistroPagosGar.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Lbl_Nombre(20)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Lbl_EstPago"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Fra_DetPgo"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Cmd_BMSumar"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Cmd_Calcular"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Fra_FormPgo"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
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
         Height          =   2175
         Left            =   240
         TabIndex        =   106
         Top             =   1080
         Width           =   4215
         Begin VB.ComboBox Cmb_Sucursal 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   600
            Width           =   2955
         End
         Begin VB.TextBox Txt_NumCta 
            Height          =   285
            Left            =   960
            MaxLength       =   15
            TabIndex        =   42
            Top             =   1680
            Width           =   2940
         End
         Begin VB.ComboBox Cmb_TipCuenta 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   960
            Width           =   2940
         End
         Begin VB.ComboBox Cmb_ViaPago 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   240
            Width           =   2955
         End
         Begin VB.ComboBox Cmb_Banco 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   1320
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
            Index           =   23
            Left            =   240
            TabIndex        =   112
            Top             =   0
            Width           =   1305
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Sucursal"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   111
            Top             =   600
            Width           =   810
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "N°Cuenta"
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   110
            Top             =   1695
            Width           =   795
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Banco"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   109
            Top             =   1320
            Width           =   825
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Vía Pago"
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   108
            Top             =   240
            Width           =   810
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Tipo Cta."
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   107
            Top             =   960
            Width           =   825
         End
      End
      Begin VB.CommandButton Cmd_Calcular 
         Caption         =   "&Calcular"
         Height          =   675
         Left            =   9480
         Picture         =   "Frm_PensRegistroPagosGar.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Cálculo Periodo Garantizado"
         Top             =   1200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_BMSumar 
         Height          =   690
         Left            =   9480
         Picture         =   "Frm_PensRegistroPagosGar.frx":04DA
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Modificar Beneficiario"
         Top             =   2040
         Width           =   735
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
         Height          =   2745
         Left            =   4800
         TabIndex        =   88
         Top             =   600
         Width           =   4395
         Begin VB.Label Lbl_FechaEfecto 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   480
            TabIndex        =   119
            Top             =   2400
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Meses No Devengados"
            Height          =   255
            Index           =   19
            Left            =   120
            TabIndex        =   95
            Top             =   1680
            Width           =   1860
         End
         Begin VB.Label Lbl_MesNoDev 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   2880
            TabIndex        =   47
            Top             =   1680
            Width           =   1080
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Meses Transcurridos"
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   94
            Top             =   1320
            Width           =   1860
         End
         Begin VB.Label Lbl_MesTrans 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   2880
            TabIndex        =   46
            Top             =   1320
            Width           =   1080
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Inicio Periodo Garantizado a Pagar"
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   93
            Top             =   600
            Width           =   2700
         End
         Begin VB.Label Lbl_FecIniPer 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   2880
            TabIndex        =   44
            Top             =   600
            Width           =   1080
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Pensión No Percibida a Valor Pte."
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   92
            Top             =   2040
            Width           =   2700
         End
         Begin VB.Label Lbl_MtoPenValPte 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   2880
            TabIndex        =   48
            Top             =   2040
            Width           =   1080
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Tipo Cambio"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   91
            Top             =   300
            Width           =   1860
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
            Index           =   24
            Left            =   240
            TabIndex        =   90
            Top             =   0
            Width           =   1560
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Fin Periodo Garantizado a Pagar"
            Height          =   255
            Index           =   8
            Left            =   135
            TabIndex        =   89
            Top             =   960
            Width           =   2700
         End
         Begin VB.Label Lbl_TipCambio 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   2880
            TabIndex        =   43
            Top             =   300
            Width           =   1080
         End
         Begin VB.Label Lbl_FecFinPer 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   2880
            TabIndex        =   45
            Top             =   960
            Width           =   1080
         End
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
         Height          =   1095
         Left            =   -74760
         TabIndex        =   82
         Top             =   2520
         Width           =   10080
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
            Left            =   9360
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Efectuar Busqueda de Dirección"
            Top             =   500
            Width           =   300
         End
         Begin VB.TextBox Txt_EmailRecep 
            Height          =   285
            Left            =   3840
            MaxLength       =   40
            TabIndex        =   36
            Top             =   780
            Width           =   5400
         End
         Begin VB.TextBox Txt_DomicRecep 
            Height          =   255
            Left            =   1095
            MaxLength       =   50
            TabIndex        =   30
            Top             =   240
            Width           =   8625
         End
         Begin VB.TextBox Txt_TelefRecep 
            Height          =   285
            Left            =   1080
            MaxLength       =   15
            TabIndex        =   35
            Top             =   780
            Width           =   1800
         End
         Begin VB.Label Lbl_Distrito 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   6600
            TabIndex        =   33
            Top             =   495
            Width           =   2655
         End
         Begin VB.Label Lbl_Provincia 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   3840
            TabIndex        =   32
            Top             =   495
            Width           =   2655
         End
         Begin VB.Label Label16 
            Caption         =   "Domicilio"
            Height          =   285
            Left            =   240
            TabIndex        =   86
            Top             =   240
            Width           =   810
         End
         Begin VB.Label Lbl_Departamento 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            TabIndex        =   31
            Top             =   495
            Width           =   2655
         End
         Begin VB.Label Label2 
            Caption         =   "Antecedentes de Dirección"
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
            Left            =   240
            TabIndex        =   87
            Top             =   0
            Width           =   2445
         End
         Begin VB.Label Label18 
            Caption         =   "Ubicación"
            Height          =   270
            Left            =   240
            TabIndex        =   85
            Top             =   500
            Width           =   825
         End
         Begin VB.Label Label20 
            Caption         =   "Teléfono"
            Height          =   270
            Left            =   240
            TabIndex        =   84
            Top             =   780
            Width           =   840
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Email"
            Height          =   255
            Index           =   13
            Left            =   3240
            TabIndex        =   83
            Top             =   780
            Width           =   465
         End
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
         Height          =   2145
         Left            =   -74760
         TabIndex        =   65
         Top             =   360
         Width           =   10095
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Est. Pago Pensión"
            Height          =   255
            Index           =   9
            Left            =   7920
            TabIndex        =   113
            Top             =   1125
            Width           =   1335
         End
         Begin VB.Label Lbl_BMEstPension 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   9360
            TabIndex        =   29
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Lbl_BMApMaterno 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   1080
            TabIndex        =   21
            Top             =   1680
            Width           =   3375
         End
         Begin VB.Label Lbl_BMApPaterno 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   1080
            TabIndex        =   20
            Top             =   1380
            Width           =   3375
         End
         Begin VB.Label Lbl_BMNombreSeg 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   1080
            TabIndex        =   19
            Top             =   1080
            Width           =   3375
         End
         Begin VB.Label Lbl_BMNombre 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   1080
            TabIndex        =   18
            Top             =   795
            Width           =   3375
         End
         Begin VB.Label Lbl_BMPrcLegal 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   5760
            TabIndex        =   25
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Lbl_BMFecNac 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   5760
            TabIndex        =   24
            Top             =   795
            Width           =   1215
         End
         Begin VB.Label Lbl_BMGrupoFam 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   9360
            TabIndex        =   28
            Top             =   795
            Width           =   495
         End
         Begin VB.Label Lbl_BMSexo 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   5760
            TabIndex        =   23
            Top             =   495
            Width           =   4095
         End
         Begin VB.Label Lbl_BMParentesco 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   5760
            TabIndex        =   22
            Top             =   210
            Width           =   4095
         End
         Begin VB.Label Lbl_BMNumIdent 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   2640
            TabIndex        =   17
            Top             =   495
            Width           =   1815
         End
         Begin VB.Label Lbl_BMTipoIdent 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   1080
            TabIndex        =   16
            Top             =   495
            Width           =   1575
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2do. Nombre"
            Height          =   195
            Index           =   69
            Left            =   120
            TabIndex        =   81
            Top             =   1080
            Width           =   915
         End
         Begin VB.Label Lbl_MonPension 
            Caption         =   "(TM)"
            Height          =   285
            Index           =   2
            Left            =   7005
            TabIndex        =   80
            Top             =   1700
            Width           =   495
         End
         Begin VB.Label Lbl_MonPension 
            Caption         =   "(TM)"
            Height          =   285
            Index           =   1
            Left            =   7005
            TabIndex        =   79
            Top             =   1400
            Width           =   495
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pensión Gar."
            Height          =   195
            Index           =   91
            Left            =   4800
            TabIndex        =   78
            Top             =   1680
            Width           =   1035
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Ap. Materno"
            Height          =   255
            Index           =   86
            Left            =   120
            TabIndex        =   77
            Top             =   1680
            Width           =   915
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Ap. Paterno"
            Height          =   255
            Index           =   85
            Left            =   120
            TabIndex        =   76
            Top             =   1380
            Width           =   915
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1er. Nombre"
            Height          =   195
            Index           =   84
            Left            =   120
            TabIndex        =   75
            Top             =   800
            Width           =   870
         End
         Begin VB.Label Lbl_BMMtoPension 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   5760
            TabIndex        =   26
            Top             =   1380
            Width           =   1215
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pensión"
            Height          =   195
            Index           =   76
            Left            =   4800
            TabIndex        =   74
            Top             =   1380
            Width           =   930
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Nac."
            Height          =   195
            Index           =   74
            Left            =   4800
            TabIndex        =   73
            Top             =   795
            Width           =   840
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Grupo Fam."
            Height          =   255
            Index           =   73
            Left            =   7920
            TabIndex        =   72
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Parentesco"
            Height          =   255
            Index           =   72
            Left            =   4800
            TabIndex        =   71
            Top             =   210
            Width           =   975
         End
         Begin VB.Label Lbl_BMNumOrd 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            TabIndex        =   15
            Top             =   210
            Width           =   495
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Orden"
            Height          =   255
            Index           =   71
            Left            =   120
            TabIndex        =   70
            Top             =   210
            Width           =   975
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Ident."
            Height          =   255
            Index           =   68
            Left            =   120
            TabIndex        =   69
            Top             =   490
            Width           =   735
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sexo"
            Height          =   195
            Index           =   67
            Left            =   4800
            TabIndex        =   68
            Top             =   495
            Width           =   960
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Prc. Pensión"
            Height          =   195
            Index           =   62
            Left            =   4800
            TabIndex        =   67
            Top             =   1080
            Width           =   900
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "%"
            Height          =   210
            Index           =   61
            Left            =   7080
            TabIndex        =   66
            Top             =   1120
            Width           =   255
         End
         Begin VB.Label Lbl_BMMtoPensionGar 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   5760
            TabIndex        =   27
            Top             =   1680
            Width           =   1215
         End
      End
      Begin VB.Label Lbl_EstPago 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         TabIndex        =   37
         Top             =   600
         Width           =   1560
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Estado de Pago  Per. Garantizado"
         Height          =   255
         Index           =   20
         Left            =   240
         TabIndex        =   105
         Top             =   600
         Width           =   2415
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
      TabIndex        =   56
      Top             =   0
      Width           =   10575
      Begin VB.ComboBox Cmb_PenNumIdent 
         BackColor       =   &H00E0FFFF&
         Height          =   315
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   2235
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
         Left            =   9840
         Picture         =   "Frm_PensRegistroPagosGar.frx":0664
         TabIndex        =   7
         ToolTipText     =   "Buscar"
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton Cmd_BuscarPol 
         Height          =   375
         Left            =   9840
         Picture         =   "Frm_PensRegistroPagosGar.frx":0766
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Buscar Póliza"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Txt_PenNumIdent 
         Height          =   285
         Left            =   6000
         MaxLength       =   16
         TabIndex        =   3
         Top             =   360
         Width           =   1875
      End
      Begin VB.TextBox Txt_PenPoliza 
         Height          =   285
         Left            =   840
         MaxLength       =   10
         TabIndex        =   1
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label Lbl_PenNombre 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   720
         Width           =   8775
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   62
         Top             =   720
         Width           =   855
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
         TabIndex        =   61
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label Lbl_End 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   9120
         TabIndex        =   4
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "N° End"
         Height          =   195
         Index           =   42
         Left            =   8520
         TabIndex        =   60
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Ident."
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   59
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Póliza"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   58
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Fra_Botones 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   7440
      Width           =   10575
      Begin VB.CommandButton Cmd_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   6705
         Picture         =   "Frm_PensRegistroPagosGar.frx":0868
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   200
         Width           =   730
      End
      Begin VB.CommandButton Cmd_Eliminar 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   3480
         Picture         =   "Frm_PensRegistroPagosGar.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Eliminar Pago"
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   4560
         Picture         =   "Frm_PensRegistroPagosGar.frx":1184
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   7800
         Picture         =   "Frm_PensRegistroPagosGar.frx":183E
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   5655
         Picture         =   "Frm_PensRegistroPagosGar.frx":1938
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Grabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   675
         Left            =   2415
         Picture         =   "Frm_PensRegistroPagosGar.frx":1FF2
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   165
         Width           =   720
      End
      Begin Crystal.CrystalReport Rpt_Imprimir 
         Left            =   9360
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Msf_BMGrilla 
      Height          =   1410
      Left            =   120
      TabIndex        =   63
      Top             =   2235
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   2487
      _Version        =   393216
      Rows            =   1
      Cols            =   12
      FixedCols       =   0
      BackColor       =   14745599
      FormatString    =   $"Frm_PensRegistroPagosGar.frx":26AC
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
Attribute VB_Name = "Frm_PensRegistroPagosGar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vlCodViaPag      As String
Dim vlAfp            As String
Dim vlSwClic         As Boolean
Dim vlCodViaPgo      As String, vlCodSucursal As String
Dim vlCodTipCuenta   As String, vlCodBco      As String
Const clCodSinDerPen As String = "10"
'--- Estados del periodo garantizado --------
Const clEstPagado    As String = "Pagado"
Const clEstNoPagado  As String = "No Pagado"
Const clEstCalculado As String = "Calculado"

'--- Casos de Sobrevivencia ---
Const clTipPenSob As String = "01,03,08,09,10,11,12,13,15"

'--- Columnas de la grilla ----------------
Dim vlNumOrd     As String, vlCodPar     As String
Dim vlTipoIdent  As String, vlNumIdent   As String
Dim vlNom        As String, vlNomSeg     As String
Dim vlApPat      As String, vlApMat      As String
Dim vlEstPension As String, vlEstPago    As String
Dim vlPrcPension As String, vlMtoPension As String
Dim vlMesesPag   As String, vlMesesNoDev As String
Dim vlPenNoPer   As String, vlFecPago    As String
Dim vlSexo       As String, vlFecNac     As String
Dim vlPrcLeg     As String, vlPrcGar     As String
Dim vlMtoPenGar  As String, vlGruFam     As String
Dim vlDirec      As String, vlCodDir     As String
Dim vlFono       As String, vlEmail      As String
Dim vlViaPago    As String, vlSucursal   As String
Dim vlTipcta     As String, vlBanco      As String
Dim vlNumCta     As String, vlFecIni     As String
Dim vlFecFin     As String, vlTipCambio  As String
Dim vlFecRecep   As String, vlTasaInt    As String
Dim vlDerpen     As String, vlCodMoneda  As String
Dim vlCodMonedaScomp As String

'------------------------------------------------------------------------
'VARIABLES DEL ENDOSO
'------------------------------------------------------------------------
'1. Variables para la Tabla Endoso
Dim vlNumEndosoNew As Integer

Dim vlNumPoliza As String, vlNumEndoso As String
Dim vlFecSolEndoso As String, vlFecEndoso As String
Dim vlCodTipEndoso As String, vlCodCauEndoso As String
Dim vlMtoDiferencia As String, vlMtoPensionOri As String
Dim vlMtoPensionCal As String, vlPrcFactor As String
Dim vlFecEfectoEnd As String, vlObsEndoso As String
Dim vlFecFinEfecto As String, vlCodEstadoEndoso As String
Dim vlGlsUsuarioCrea As String, vlFecCrea As String, vlHorCrea As String
Dim vlGlsUsuarioModi As String, vlFecModi As String, vlHorModi As String
'2. Variables para la Póliza
Dim vlCodAFP As String
Dim vlPrcTasaCtoRea As Double, vlPrcTasaIntPerGar As Double
Dim vlPrcTasaTir As Double
Dim vlCodCuspp As String, vlIndCobertura As String
Dim vlValMoneda As Double
Dim vlCobCoberCon As String
Dim vlMtoFacPenElla As Double, vlPrcFacPenElla As Double
Dim vlCodDerCrecer As String, vlCodDerGratificacion As String
Dim vlFecIniPagoPen As String, vlFecEmision As String
Dim vlFecDevengue As String, vlFecIniPenCia As String
Dim vlFecPriPago As String
Dim vlCodTipPension As String, vlCodEstado As String
Dim vlCodTipRen As String, vlCodModalidad As String
Dim vlNumCargas As Long
Dim vlFecVigencia As String, vlFecTerVigencia As String
Dim vlMtoPrima As Double
Dim vlNumMesDif As Integer, vlNumMesGar As Integer
Dim vlPrcTasaCe As Double, vlPrcTasaVta As Double
Dim vlFecTerPagoPenDif As String, vlFecTerPagoPenGar As String
Dim vlMtoPensionGar As Double
'3. Variables faltantes para Beneficiarios
Dim vlCodSitInv As String, vlCodDerCre As String
Dim vlCodCauInv As String
Dim vlFecNacHM As String, vlFecInvBen As String
Dim vlFecFallBen As String
Dim vlCodMotReqPen As String, vlCodCauSusBen As String
Dim vlFecSusBen As String
Dim vlFecMatrimonio As String
Dim vlCodInsSalud As String, vlCodModSalud As String
Dim vlMtoPlanSalud As Double
Dim vlFecIngreso As String
'RRR 31/10/2019
Dim vlTelben2, vlTipctaB, vlMonbco, vlBolelec, vlNumCCI, vlTrainfo, vlDatcomer, vlCtabco As String




Dim vlTotalFilas As Integer
'--RRR
Dim vlTtpreaj As String
Dim vlMtoajTri, vlMtoajMen As Double
Dim vlFecSol As String
Dim vlBendes As String
Dim vlBolelect As String
Dim vlIndHerencia As String
Dim vlIndEstSub As String


'--------------------- Número de Póliza ---------------------
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

'--------------------- Tipo Identificación ---------------------
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

'--------------------- Número Identificación ---------------------
Private Sub Txt_PenNumIdent_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If (Trim(Txt_PenNumIdent) <> "") Then
            Txt_PenNumIdent = UCase(Trim(Txt_PenNumIdent))
        End If
        Cmd_BuscarPol.SetFocus
    End If
End Sub
Private Sub txt_pennumident_lostfocus()
        Txt_PenNumIdent = Trim(UCase(Txt_PenNumIdent))
End Sub

'--------------------- Fecha de Recepción ---------------------
Private Sub Txt_FecRecepcion_GotFocus()
    Txt_FecRecepcion.SelStart = 0
    Txt_FecRecepcion.SelLength = Len(Txt_FecRecepcion)
End Sub
Private Sub Txt_FecRecepcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(Txt_FecRecepcion <> "") Then
       If (flValidaFecha(Txt_FecRecepcion) = False) Then
          Txt_FecRecepcion = ""
          Exit Sub
       End If
       Txt_FecRecepcion.Text = Format(CDate(Trim(Txt_FecRecepcion)), "yyyymmdd")
       Txt_FecRecepcion.Text = DateSerial(Mid((Txt_FecRecepcion.Text), 1, 4), Mid((Txt_FecRecepcion.Text), 5, 2), Mid((Txt_FecRecepcion.Text), 7, 2))
       Txt_FecPago.SetFocus
    End If
End Sub
Private Sub Txt_FecRecepcion_LostFocus()
   If (Trim(Txt_FecRecepcion)) <> "" Then
       If (flValidaFecha(Txt_FecRecepcion) = False) Then
          Txt_FecRecepcion = ""
          Exit Sub
      End If
      Txt_FecRecepcion.Text = Format(CDate(Trim(Txt_FecRecepcion)), "yyyymmdd")
      Txt_FecRecepcion.Text = DateSerial(Mid((Txt_FecRecepcion.Text), 1, 4), Mid((Txt_FecRecepcion.Text), 5, 2), Mid((Txt_FecRecepcion.Text), 7, 2))
   End If
End Sub

'--------------------- Fecha de Pago ---------------------
Private Sub Txt_FecPago_GotFocus()
    Txt_FecPago.SelStart = 0
    Txt_FecPago.SelLength = Len(Txt_FecPago)
End Sub
Private Sub Txt_FecPago_KeyPress(KeyAscii As Integer)
Dim vlCambio As Double
    If KeyAscii = 13 And Trim(Txt_FecPago <> "") Then
       If (flValidaFecha(Txt_FecPago) = False) Then
          Txt_FecPago = ""
          Exit Sub
       End If
        'Obtiene Valor Tipo Cambio a la fecha de pago
        If Not fgObtieneConversion(Format(Trim(Txt_FecPago), "yyyymmdd"), vlCodMoneda, vlCambio) Then
            MsgBox "Debe ingresar Tipo de Cambio para la Moneda : '" & vlCodMoneda & "'", vbCritical, "Falta Tipo de Cambio"
            Exit Sub
        Else
            Lbl_TipoCambio = Format(vlCambio, "#,#0.000")
        End If
       Txt_FecPago.Text = Format(CDate(Trim(Txt_FecPago)), "yyyymmdd")
       Txt_FecPago.Text = DateSerial(Mid((Txt_FecPago.Text), 1, 4), Mid((Txt_FecPago.Text), 5, 2), Mid((Txt_FecPago.Text), 7, 2))
       Txt_TasaInteres.SetFocus
    End If
End Sub
Private Sub Txt_FecPago_LostFocus()
   If (Trim(Txt_FecPago)) <> "" Then
       If (flValidaFecha(Txt_FecPago) = False) Then
          Txt_FecPago = ""
          Exit Sub
      End If
      'Obtiene Valor Tipo Cambio a la fecha de pago
      If Not fgObtieneConversion(Format(Trim(Txt_FecPago), "yyyymmdd"), vlCodMoneda, vlCambio) Then
          MsgBox "Debe ingresar Tipo de Cambio para la Moneda : '" & vlCodMoneda & "'", vbCritical, "Falta Tipo de Cambio"
          Exit Sub
      Else
          Lbl_TipoCambio = Format(vlCambio, "#,#0.000")
      End If
      Txt_FecPago.Text = Format(CDate(Trim(Txt_FecPago)), "yyyymmdd")
      Txt_FecPago.Text = DateSerial(Mid((Txt_FecPago.Text), 1, 4), Mid((Txt_FecPago.Text), 5, 2), Mid((Txt_FecPago.Text), 7, 2))
   Else
    Lbl_TipoCambio = ""
   End If
End Sub

'--------------------- Tasa de Interés Anual ---------------------
Private Sub Txt_TasaInteres_GotFocus()
    Txt_TasaInteres.SelStart = 0
    Txt_TasaInteres.SelLength = Len(Txt_TasaInteres)
End Sub
Private Sub Txt_TasaInteres_Change()
If Not IsNumeric(Txt_TasaInteres) Then
    Txt_TasaInteres = ""
End If
End Sub
Private Sub Txt_TasaInteres_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If IsNumeric(Txt_TasaInteres) Then
        Txt_TasaInteres = Format(Txt_TasaInteres, "#,#0.00")
        Cmd_Calcular.SetFocus
    End If
End If
End Sub
Private Sub Txt_TasaInteres_LostFocus()
    If IsNumeric(Txt_TasaInteres) Then
        Txt_TasaInteres = Format(Txt_TasaInteres, "#,#0.00")
    End If
End Sub

'------------------------ Dirección --------------------------------
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

'------------------------ Telefono --------------------------------
Private Sub Txt_TelefRecep_GotFocus()
    Txt_TelefRecep.SelStart = 0
    Txt_TelefRecep.SelLength = Len(Txt_TelefRecep)
End Sub
Private Sub Txt_TelefRecep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If (Trim(Txt_TelefRecep)) <> "" Then
          Txt_TelefRecep = UCase(Trim(Txt_TelefRecep))
          Txt_EmailRecep.SetFocus
      End If
   End If
End Sub
Private Sub Txt_TelefRecep_LostFocus()
   If (Trim(Txt_TelefRecep)) <> "" Then
       Txt_TelefRecep = UCase(Trim(Txt_TelefRecep))
   End If
End Sub
'------------------------ Email --------------------------------
Private Sub Txt_EmailRecep_GotFocus()
    Txt_EmailRecep.SelStart = 0
    Txt_EmailRecep.SelLength = Len(Txt_EmailRecep)
End Sub
Private Sub Txt_EmailRecep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If (Trim(Txt_EmailRecep)) <> "" Then
          Txt_EmailRecep = UCase(Trim(Txt_EmailRecep))
          SSTab1 = 1
          Cmb_ViaPago.SetFocus
      End If
   End If
End Sub
Private Sub Txt_EmailRecep_LostFocus()
   If (Trim(Txt_EmailRecep)) <> "" Then
       Txt_EmailRecep = UCase(Trim(Txt_EmailRecep))
   End If
End Sub
'--------------------- Combo Vía Pago ---------------------------------
Private Sub Cmb_ViaPago_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If (Cmb_Sucursal.Enabled = True) Then
            Cmb_Sucursal.SetFocus
        Else
            Cmd_Calcular.SetFocus
        End If
    End If
End Sub
'--------------------- Combo Sucursal ---------------------------------
Private Sub Cmb_Sucursal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If vlCodViaPgo = "02" Or vlCodViaPgo = "03" Then
           Cmb_TipCuenta.SetFocus
        Else
           Cmd_Calcular.SetFocus
        End If
    End If
End Sub
'--------------------- Combo Tipo de Cuenta ---------------------------------
Private Sub Cmb_TipCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If vlCodViaPgo = "02" Or vlCodViaPgo = "03" Then
           Cmb_Banco.SetFocus
        Else
           Cmd_Calcular.SetFocus
        End If
    End If
End Sub
'--------------------- Combo Banco ---------------------------------
Private Sub Cmb_Banco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If vlCodViaPgo = "02" Or vlCodViaPgo = "03" Then
           Txt_NumCta.SetFocus
        Else
           Cmd_Calcular.SetFocus
        End If
    End If
End Sub
'--------------------- Número de Cuenta ---------------------------------
Private Sub Txt_NumCta_GotFocus()
    Txt_NumCta.SelStart = 0
    Txt_NumCta.SelLength = Len(Txt_NumCta)
End Sub
Private Sub Txt_NumCta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If (Trim(Txt_NumCta)) <> "" Then
          Txt_NumCta = UCase(Trim(Txt_NumCta))
          Cmd_Calcular.SetFocus
      End If
   End If
End Sub
Private Sub Txt_NumCta_LostFocus()
   If (Trim(Txt_NumCta)) <> "" Then
       Txt_NumCta = UCase(Trim(Txt_NumCta))
   End If
End Sub


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


Private Sub Cmd_BMSumar_Click()
Dim vlFila As Integer
On Error GoTo Err_CmdSumar
    
    Screen.MousePointer = 11
    
    If (flValidaDatos = False) Then
        Screen.MousePointer = 0
        Exit Sub
    End If
   
    fila = 0
    Msf_BMGrilla.Col = 0
    For i = 1 To Msf_BMGrilla.rows
        Msf_BMGrilla.row = i
        If (Lbl_BMNumOrd = Msf_BMGrilla.Text) Then
            vlFila = i
            Exit For
        End If
    Next
    
    If (vlFila > 0) Then
    
       Msf_BMGrilla.row = vlFila
       
       vlNumOrd = Trim(Lbl_BMNumOrd)
       vlCodPar = Trim(Mid(Lbl_BMParentesco, 1, InStr(Lbl_BMParentesco, "-") - 1))
       vlTipoIdent = Trim(Lbl_BMTipoIdent)
       ''vlTipoIdent = (vgRs2!Cod_TipoIdenBen & " - " & fgBuscarNombreTipoIden(vgRs2!Cod_TipoIdenBen))
       vlNumIdent = Trim(Lbl_BMNumIdent)
       vlNom = Trim(Lbl_BMNombre)
       vlNomSeg = Trim(Lbl_BMNombreSeg)
       vlApPat = Trim(Lbl_BMApPaterno)
       vlApMat = Trim(Lbl_BMApMaterno)
       vlEstPension = Trim(Lbl_BMEstPension)
       vlEstPago = Trim(Lbl_EstPago)
       vlPrcPension = Trim(Lbl_BMPrcLegal)
       vlMtoPension = Trim(Lbl_BMMtoPension)
       vlMesesPag = Trim(Lbl_MesTrans)
       vlMesesNoDev = Trim(Lbl_MesNoDev)
       vlPenNoPer = Trim(Lbl_MtoPenValPte)
       vlFecPago = Trim(Txt_FecPago)
       vlSexo = Trim(Mid(Lbl_BMSexo, 1, InStr(Lbl_BMSexo, "-") - 1))
       vlFecNac = Format(Trim(Lbl_BMFecNac), "yyyyMMdd")
       Msf_BMGrilla.Col = 18
       vlPrcLeg = Trim(Msf_BMGrilla.Text)
       Msf_BMGrilla.Col = 19
       vlPrcGar = Trim(Msf_BMGrilla.Text)
       vlMtoPenGar = Trim(Lbl_BMMtoPensionGar)
       vlGruFam = Trim(Lbl_BMGrupoFam)
       vlDirec = Trim(Txt_DomicRecep)
       'vlCodDir =
       vlFono = Trim(Txt_TelefRecep)
       vlEmail = Trim(Txt_EmailRecep)
       vlViaPago = Trim(Mid(Cmb_ViaPago, 1, InStr(Cmb_ViaPago, "-") - 1))
       vlSucursal = Trim(Mid(Cmb_Sucursal, 1, InStr(Cmb_Sucursal, "-") - 1))
       vlTipcta = Trim(Mid(Cmb_TipCuenta, 1, InStr(Cmb_TipCuenta, "-") - 1))
       vlBanco = Trim(Mid(Cmb_Banco, 1, InStr(Cmb_Banco, "-") - 1))
       vlNumCta = Trim(Txt_NumCta)
       vlFecIni = Format(Trim(Lbl_FecIniPer), "yyyyMMdd")
       vlFecFin = Format(Trim(Lbl_FecFinPer), "yyyyMMdd")
       vlTipCambio = Trim(Lbl_TipCambio)
       vlFecRecep = Format(Trim(Txt_FecRecepcion), "yyyyMMdd")
       vlTasaInt = Trim(Txt_TasaInteres)
       Msf_BMGrilla.Col = 36
       vlDerpen = Trim(Msf_BMGrilla.Text)
        
       If (Msf_BMGrilla.rows = 2) Then
        Call flInicializaGrilla
       Else
        Msf_BMGrilla.RemoveItem vlFila
       End If
       
       Msf_BMGrilla.AddItem (vlNumOrd) & vbTab & (vlCodPar) & vbTab & _
                       " " & vlTipoIdent & vbTab & (vlNumIdent) & vbTab & _
                       (vlNom) & vbTab & (vlNomSeg) & vbTab & _
                       (vlApPat) & vbTab & (vlApMat) & vbTab & _
                       (vlEstPension) & vbTab & (vlEstPago) & vbTab & _
                       (vlPrcPension) & vbTab & (vlMtoPension) & vbTab & _
                       (vlMesesPag) & vbTab & (vlMesesNoDev) & vbTab & _
                       (vlPenNoPer) & vbTab & (vlFecPago) & vbTab & _
                       (vlSexo) & vbTab & (vlFecNac) & vbTab & _
                       (vlPrcLeg) & vbTab & (vlPrcGar) & vbTab & _
                       (vlMtoPenGar) & vbTab & (vlGruFam) & vbTab & _
                       (vlDirec) & vbTab & (vlCodDir) & vbTab & _
                       (vlFono) & vbTab & (vlEmail) & vbTab & _
                       (vlViaPago) & vbTab & (vlSucursal) & vbTab & _
                       (vlTipcta) & vbTab & (vlBanco) & vbTab & _
                       (vlNumCta) & vbTab & (vlFecIni) & vbTab & _
                       (vlFecFin) & vbTab & (vlTipCambio) & vbTab & _
                       (vlFecRecep) & vbTab & (vlTasaInt) & vbTab & _
                       (vlDerpen), vlFila
    
    End If
    
    Screen.MousePointer = 0
    
Exit Sub
Err_CmdSumar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Buscar_Click()
On Error GoTo Err_CmdBuscarClick

    Frm_Busqueda.flInicio ("Frm_PensRegistroPagosGar")
    
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

    Frm_BusDireccion.flInicio ("Frm_PensRegistroPagosGar")
    
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

Private Sub Cmd_BuscarPol_Click()
Dim vlNombreSeg As String, vlApMaterno As String
Dim vlNumMesGar As Integer, vlNumero As Integer
Dim vlTipoPension As String
Dim vlNumOrden As Integer
Dim vlFecVigencia As String

On Error GoTo Err_CmdBuscarPolClick
        
    vlSwSeleccionado = False

    If Txt_PenPoliza.Text = "" Then
       If ((Trim(Cmb_PenNumIdent.Text)) = "") Or (Txt_PenNumIdent.Text = "") Then
           MsgBox "Debe Ingresar el Número de Póliza o la Identificación del Pensionado.", vbCritical, "Error de Datos"
           Txt_PenPoliza.SetFocus
           Exit Sub
       Else
           Txt_PenNumIdent = UCase(Trim(Txt_PenNumIdent))
           Txt_PenNumIdent.SetFocus
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
           vgSql = vgSql & "num_poliza = '" & Txt_PenPoliza.Text & "' "
           ''**vgSql = vgSql & "AND cod_estpension <> '" & clCodSinDerPen & "' "
           vgSql = vgSql & "ORDER BY num_endoso DESC, num_orden ASC "
           Set vgRegistro = vgConexionBD.Execute(vgSql)
           If (vgRegistro!numeroben) <> 0 Then
              vgPalabra = "num_poliza = '" & Txt_PenPoliza.Text & "' "
              ''**vgPalabra = vgPalabra & "AND cod_estpension <> '" & clCodSinDerPen & "' "
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
    vgSql = "SELECT num_endoso,num_orden,gls_nomben,gls_nomsegben,gls_patben,gls_matben,"
    vgSql = vgSql & "cod_estpension,cod_tipoidenben,num_idenben,num_poliza "
    vgSql = vgSql & "FROM PP_TMAE_BEN WHERE "
    vgSql = vgSql & vgPalabra
    vgSql = vgSql & " ORDER BY num_endoso DESC,num_orden ASC "
    Set vgRs2 = vgConexionBD.Execute(vgSql)
    If Not vgRs2.EOF Then
        
        'Valida que la póliza sea garantizada
        Sql = "select cod_tippension,num_mesgar,fec_vigencia from pp_tmae_poliza "
        Sql = Sql & "where num_poliza='" & Trim(vgRs2!num_poliza) & "' "
        Sql = Sql & "and num_endoso=" & Trim(vgRs2!num_endoso) & " "
        Set vgRs = vgConexionBD.Execute(Sql)
        If Not vgRs.EOF Then
            vlNumMesGar = vgRs!Num_MesGar
            vlTipoPension = vgRs!Cod_TipPension
            vlFecVigencia = vgRs!Fec_Vigencia
        End If
        vgRs.Close
        
        'Determina si la pensión proviene de una Sobrevivencia de Origen o de ...
        vgI = InStr(1, clTipPenSob, vlTipoPension)
        
'        If (vlNumMesGar = 0) Or (vgI = 0) Then
'            MsgBox " La Póliza Seleccionada No tiene Derecho al Pago de una Pensión Garantizada ", vbCritical, "Proceso Cancelado"
'
''            'Desactivar Todos los Controles del Formulario
''            Call flDeshabilitarIngreso
''            Fra_Poliza.Enabled = False
'            Txt_PenPoliza.SetFocus
'            Exit Sub
'        End If
        
''       If Trim(vgRs2!Cod_EstPension) = Trim(clCodSinDerPen) Then '* debe preg esto
''          MsgBox " El Beneficiario Seleccionado No Tiene Derecho a Pensión " & Chr(13) & _
''          "          Sólo podrá Consultar los Datos del Registro", vbInformation, "Información"
''
''          'Desactivar Todos los Controles del Formulario
''            Fra_Poliza.Enabled = False
''            Fra_DetallePago.Enabled = False
''            SSTab1.Enabled = False
''
''       Else
''           If Not fgValidaVigenciaPoliza(Trim(Txt_PenPoliza.Text), fgBuscaFecServ) Then
''              MsgBox " La Póliza Ingresada no se Encuentra Vigente en el Sistema " & Chr(13) & _
''                     "      Sólo podrá Consultar los Datos del Registro", vbInformation, "Información"
''
''              'Desactivar Todos los Controles del Formulario
''                Fra_Poliza.Enabled = False
''                Fra_DetallePago.Enabled = False
''                SSTab1.Enabled = False
''           Else
''               flHabilitarIngreso
''           End If
''       End If
              
        vlAfp = fgObtenerPolizaCod_AFP(vgRs2!num_poliza, CStr(vgRs2!num_endoso))
        
        vlCodTipoIdenBenCau = vgRs2!Cod_TipoIdenBen
        vlNumIdenBenCau = Trim(vgRs2!Num_IdenBen)
             
        Txt_PenPoliza.Text = Trim(vgRs2!num_poliza)
        Call fgBuscarPosicionCodigoCombo(vlCodTipoIdenBenCau, Cmb_PenNumIdent)
        Txt_PenNumIdent.Text = vlNumIdenBenCau
        vlNombreSeg = IIf(IsNull(vgRs2!Gls_NomSegBen), "", Trim(vgRs2!Gls_NomSegBen))
        vlApMaterno = IIf(IsNull(vgRs2!Gls_MatBen), "", Trim(vgRs2!Gls_MatBen))
        Lbl_PenNombre.Caption = fgFormarNombreCompleto(Trim(vgRs2!Gls_NomBen), vlNombreSeg, Trim(vgRs2!Gls_PatBen), vlApMaterno)
        Lbl_End.Caption = (vgRs2!num_endoso)
        vlNumEndoso = (vgRs2!num_endoso)
        vlNumOrden = (vgRs2!Num_Orden)
       
        Lbl_FechaEfecto = fgValidaFechaEfecto(DateSerial(Mid(vlFecVigencia, 1, 4), Mid(vlFecVigencia, 5, 2), Mid(vlFecVigencia, 7, 2)), Trim(Txt_PenPoliza), vlNumOrden)
       
       'Busca la Moneda de la póliza seleccionada
        ''**Lbl_Moneda.Caption = flBuscaMoneda(Trim(Txt_PenPoliza.Text), Lbl_End)
           
        'se cargan los datos de la póliza
        Call flCargarDatosPol(Trim(Txt_PenPoliza), Trim(Lbl_End))
        
        'se carga la grilla
        Call flCargaGrilla(Trim(Txt_PenPoliza), Trim(Lbl_End))
        
        If (vlNumMesGar > 0) Then
            ''SSTab1.Enabled = True
            Call flHabilitarIngreso
            SSTab1.Tab = 0
        End If
        
    Else
        MsgBox "El Beneficiario o la Póliza Ingresados, No Existen en la Base de Datos", vbInformation, "Información"
        Txt_PenPoliza.SetFocus
        Exit Sub
    End If
    vgRs2.Close
        
    SSTab1.Tab = 0
    vlSwMostrar = False

''**    If Fra_AntecedentesRet.Enabled = True Then
''       Txt_FechaIniVig.SetFocus
''    End If
       
Exit Sub
Err_CmdBuscarPolClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Calcular_Click()
Dim Pension As Double, vlTasa As Double
Dim FInicioPG As String
Dim FFinPG As String
Dim FCalculo As String
Dim iab As Integer, imb As Integer, idb As Integer
Dim iaIpG As Integer, imIpG As Integer, idIpG As Integer
Dim iaFpG As Integer, imFpG As Integer, idFpG As Integer
Dim MesesTranscurridos As Long
Dim MesesNodevengados As Long
Dim TasaAnual   As Double, TasaMensual  As Double
Dim FactorVA    As Double, MtoVA        As Double
On Error GoTo Err_CmdCalcular

    If Trim(Lbl_BMNumOrd) = "" Then
        MsgBox "Debe Seleccionar un beneficiario a calcular.", vbCritical, "Operación Cancelada"
        Exit Sub
    End If

    If (Trim(Lbl_EstPago) = clEstPagado) Then
        MsgBox "No se Puede Calcular ya que el Periodo Garantizado se Encuentra Pagado. ", vbCritical, "Proceso Cancelado"
        Exit Sub
    End If

    If Trim(Txt_FecPago) = "" Then
       MsgBox "Debe Ingresar la Fecha de Pago.", vbCritical, "Operación Cancelada"
       Txt_FecPago.SetFocus
       Exit Sub
    Else
      If (flValidaFecha(Txt_FecPago) = False) Then
          Txt_FecPago.SetFocus
          Exit Sub
      End If
    End If
    
    If Trim(Txt_TasaInteres) = "" Then
       MsgBox "Debe Ingresar la Tasa de Interés Anual del Pago.", vbCritical, "Error de Datos"
       Txt_TasaInteres.SetFocus
       Exit Sub
    Else
      If Not IsNumeric(Txt_TasaInteres) Then
          Txt_TasaInteres.SetFocus
          Exit Sub
      End If
    End If
    If CDbl(Txt_TasaInteres) = 0 Then
        MsgBox "Debe ingresar un valor distinto de Cero para la Tasa de Interés.", vbCritical, "Error de Datos"
        Txt_TasaInteres.SetFocus
        Exit Sub
    End If

    'Validar que tenga Derecho al Pago de Pensiones
    If (Lbl_BMEstPension = "99") Then
        MsgBox "Este Beneficiario cuenta con derecho de pensión.", vbCritical, "Error de Datos"
        Txt_FecRecepcion.SetFocus
        Exit Sub
    End If
    
    'Validar que la Fecha de Pago no sea Menor a la Fecha de Inicio del Pago
    If (Format(Lbl_FecIniPer, "yyyymmdd") > Format(Txt_FecPago, "yyyymmdd")) Then
        MsgBox "La Fecha de Pago es anterior a la Fecha de Inicio del Pago Garantizado", vbExclamation, "Precaución"
    End If
    
    Cmd_Grabar.Enabled = True
    vlTasa = CDbl(Txt_TasaInteres)
    Lbl_FecFinPer = Lbl_FecFinPerGar
    FactorVA = 0
    MtoVA = 0
    MesesTranscurridos = 0
    MesesNodevengados = 0
    
    'Rutina de Cálculo
'I--- ABV 11/11/2007 ---
'    Pension = Lbl_BMMtoPension
    Pension = Lbl_BMMtoPensionGar
'F--- ABV 11/11/2007 ---
'I--- ABV 17/11/2007 ---
'    FInicioPG = Format(Lbl_FecIniPer, "yyyymmdd")
    FInicioPG = Format(Lbl_FecIniPerGar, "yyyymmdd")
'F--- ABV 17/11/2007 ---
    
    FFinPG = Format(Lbl_FecFinPer, "yyyymmdd")
    FCalculo = Format(Txt_FecPago, "yyyymmdd")
    
    iab = Mid(FCalculo, 1, 4)
    imb = Mid(FCalculo, 5, 2)
    idb = Mid(FCalculo, 7, 2)
    
    iaIpG = Mid(FInicioPG, 1, 4)
    imIpG = Mid(FInicioPG, 5, 2)
    idIpG = Mid(FInicioPG, 7, 2)
    
    iaFpG = Mid(FFinPG, 1, 4)
    imFpG = Mid(FFinPG, 5, 2)
    idFpG = Mid(FFinPG, 7, 2)
    
'I--- ABV 17/11/2007 ---
'    MesesTranscurridos = ((iab * 12 + imb) - (iaIpG * 12 + imIpG))
    MesesTranscurridos = ((iab * 12 + imb) - (iaIpG * 12 + imIpG)) + 1
'F--- ABV 17/11/2007 ---
    
    MesesNodevengados = ((iaFpG * 12 + imFpG) - (iab * 12 + imb))
    
    TasaAnual = vlTasa / 100
    
    TasaMensual = (1 + TasaAnual) ^ (1 / 12) - 1
    
    'VPMonto = PV(TasaMensual, MesesNodevengados, -pension, 0, 0)
    'VPFactor = PV(TasaMensual, MesesNodevengados, -1, 0, 0)
    
    FactorVA = (1 - (1 / (1 + TasaMensual) ^ MesesNodevengados)) / TasaMensual
    MtoVA = Pension * (1 - (1 / (1 + TasaMensual) ^ MesesNodevengados)) / TasaMensual

    'Imprimir los valores calculados
    Lbl_TipCambio = Lbl_TipoCambio '?*borrar despues
    'Lbl_FecIniPer = "01/09/2007"
    'Lbl_FecFinPer = "01/09/2008"
    Lbl_MesTrans = Format(MesesTranscurridos, "#0")
    Lbl_MesNoDev = Format(MesesNodevengados, "#0")
    Lbl_MtoPenValPte = Format(MtoVA, "#0.00")
    Lbl_EstPago = clEstCalculado
    'Lbl_FecFinPerGar = "12/09/2007"

Exit Sub
Err_CmdCalcular:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Cancelar_Click()
    Lbl_End.Caption = ""

    Call flDeshabilitarIngreso
    Call Cmd_Limpiar_Click
    
    Txt_PenPoliza.Text = ""
    Txt_PenNumIdent.Text = ""
    Cmb_PenNumIdent.ListIndex = 0
    Lbl_PenNombre.Caption = ""
    
    Lbl_TipoCambio = ""
    Lbl_FecDevengue = ""
    Lbl_FecIniPago = ""
    Lbl_FecIniPerGar = ""
    Lbl_FecFinPerGar = ""
    Lbl_MtoPensionRef = ""
    
    Txt_FecRecepcion = ""
    Txt_FecPago = ""
    Txt_TasaInteres = ""
    
    Msf_BMGrilla.rows = 1
        
    vlSw = True
    
    Call flLimpiar
    vlAfp = ""
    
    vlSw = False
    
    Call flDeshabilitarIngreso
    SSTab1.Tab = 0
    
    Lbl_MonPension(0) = ""
    Lbl_MonPension(1) = ""
    Lbl_MonPension(2) = ""
    
    Txt_PenPoliza.SetFocus
    Cmd_Grabar.Enabled = False
End Sub

Private Sub cmd_grabar_Click()
Dim vlFila As Integer, i As Integer
Dim vlident As String
Dim vlFPago As String
On Error GoTo Err_CmdGrabar
    
    'Valida si existen beneficiarios
    If (Msf_BMGrilla.rows = 1) Then
        MsgBox "No Existen Beneficiarios a Grabar. ", vbCritical, "Proceso Cancelado"
        Exit Sub
    End If
    
    'Valida que al menos 1 beneficiario este calculado
    vgSw = False
    Msf_BMGrilla.Col = 9
    vlFila = Msf_BMGrilla.rows
    For i = 1 To vlFila - 1
        Msf_BMGrilla.row = i
        If (Trim(Msf_BMGrilla.Text) = clEstCalculado) Then
            vgSw = True
            Exit For
        End If
    Next
    If (vgSw = False) Then
        MsgBox "No Existe Ningún Beneficiario Calculado para Grabar. ", vbCritical, "Proceso Cancelado"
        Exit Sub
    End If

    'Valida la información del Detalle de Pago
    If Trim(Txt_FecRecepcion) = "" Then
       MsgBox "Debe Ingresar la Fecha de Recepción del Pago.", vbCritical, "Operación Cancelada"
       Txt_FecRecepcion.SetFocus
       Exit Sub
    Else
      If (flValidaFecha(Txt_FecRecepcion) = False) Then
          Txt_FecRecepcion.SetFocus
          Exit Sub
      End If
    End If
    
    If Trim(Txt_FecPago) = "" Then
       MsgBox "Debe Ingresar la Fecha de Pago.", vbCritical, "Operación Cancelada"
       Txt_FecPago.SetFocus
       Exit Sub
    Else
      If (flValidaFecha(Txt_FecPago) = False) Then
          Txt_FecPago.SetFocus
          Exit Sub
      End If
    End If
    
    If (CDate(Txt_FecRecepcion) > CDate(Txt_FecPago)) Then
       MsgBox "La Fecha de Pago debe ser mayor a la Fecha de Recepción", vbCritical, "Operación Cancelada"
       Txt_FechaPgo.SetFocus
       Exit Sub
    End If
    
    If Trim(Txt_TasaInteres) = "" Then
       MsgBox "Debe Ingresar la Tasa de Interés Anual del Pago.", vbCritical, "Operación Cancelada"
       Txt_TasaInteres.SetFocus
       Exit Sub
    Else
      If Not IsNumeric(Txt_TasaInteres) Then
          Txt_TasaInteres.SetFocus
          Exit Sub
      End If
    End If

    'Abrir Conexión de BD
    If Not fgConexionBaseDatos(vgConectarBD) Then
        MsgBox "Error en Conexión a la Base de Datos", vbCritical, Me.Caption
        Exit Sub
    End If
    
    'Iniciar Transacción
    vgConexionBD.BeginTrans
    
'1. Registrar los antecedentes de Pago en las Tablas del Periodo Garantizado

    'Valida que al menos 1 beneficiario este calculado
    vgSw = False
    vlFila = Msf_BMGrilla.rows
    For i = 1 To vlFila - 1
        Msf_BMGrilla.row = i
        Msf_BMGrilla.Col = 9
        If (Trim(Msf_BMGrilla.Text) = clEstCalculado) Then
                    
            Msf_BMGrilla.Col = 15
            vlFPago = Format(Trim(Msf_BMGrilla.Text), "yyyyMMdd")
            
            vgSw = False
            'Valida que exista la póliza
            Sql = "select 1 from pp_tmae_pagtergar "
            Sql = Sql & "where num_poliza='" & Trim(Txt_PenPoliza) & "' "
            Sql = Sql & "and fec_pago= '" & Trim(vlFPago) & "' "
            ''Sql = Sql & "and num_endoso=" & Trim(Lbl_End) & " "
            ''Set vgRs = vgConectarBD.Execute(Sql)
            Set vgRs = vgConexionBD.Execute(Sql)
            If Not vgRs.EOF Then
                vgSw = True
            End If
            vgRs.Close
            
            If (vgSw = False) Then
                Sql = ""
                Sql = "INSERT INTO PP_TMAE_PAGTERGAR (num_poliza,num_endoso,fec_pago,"
                Sql = Sql & "fec_solpago,prc_tasaint,cod_conpago,cod_moneda,"
                Sql = Sql & "mto_valmoneda,mto_pension,mto_pensiongar,num_endosocrear,"
                Sql = Sql & "fec_inipencia,fec_finpergar,fec_inipergarpag,"
                Sql = Sql & "cod_usuariocrea,fec_crea,hor_crea ) VALUES ("
                Sql = Sql & "'" & Trim(Txt_PenPoliza) & "',"
                Sql = Sql & " " & CInt(Trim(Lbl_End)) & ","
                Sql = Sql & "'" & Trim(vlFPago) & "',"
                Msf_BMGrilla.Col = 34
                Sql = Sql & "'" & Trim(Msf_BMGrilla.Text) & "',"
                ''Sql = Sql & "'" & Format(Trim(Txt_FecRecepcion), "yyyyMMdd") & "',"
                Msf_BMGrilla.Col = 35
                Sql = Sql & " " & str(Trim(Msf_BMGrilla.Text)) & ","
                ''Sql = Sql & " " & Str(Txt_TasaInteres) & ","
                Sql = Sql & "'" & cgPagoTerceroPerGar & "',"
                Sql = Sql & "'" & Trim(vlCodMoneda) & "',"
                Msf_BMGrilla.Col = 33
                ''Sql = Sql & " " & Str(Trim(Lbl_TipoCambio)) & ","
                Sql = Sql & " " & str(Trim(Msf_BMGrilla.Text)) & ","
                Sql = Sql & " " & str(Trim(Lbl_MtoPensionRef)) & ","
                Sql = Sql & " " & str(Trim(Lbl_MtoPensionRef)) & ","
                Sql = Sql & " " & CInt(Trim(Lbl_End)) & ","
                Sql = Sql & "'" & Format(Trim(Lbl_FecIniPago), "yyyyMMdd") & "',"
                Sql = Sql & "'" & Format(Trim(Lbl_FecFinPerGar), "yyyyMMdd") & "',"
                Sql = Sql & "'" & Format(Trim(Lbl_FecIniPerGar), "yyyyMMdd") & "',"
                Sql = Sql & "'" & vgUsuario & "',"
                Sql = Sql & "'" & Format(Now(), "yyyyMMdd") & "',"
                Sql = Sql & "'" & Format(Now(), "hhmmss") & "' ) "
                vgConexionBD.Execute Sql
                
                
                    'CORPTEC
                vlLog_tabla = "PP_TMAE_PAGTERGAR"
                vlLog_idtabla = "NUM_POLIZA.FEC_PAGO"
                vlLog_valtabla = "" & Trim(Txt_PenPoliza) & "." & Trim(vlFPago) & ""
                vlLog_trans = "INS"
                'Call flLog_Tabla
    
            End If

            vgSw = False
            Msf_BMGrilla.Col = 0
            'Verifica si existe el beneficiario
            Sql = "select 1 from pp_tmae_pagtergarben "
            Sql = Sql & "where num_poliza='" & Trim(Txt_PenPoliza) & "' "
            Sql = Sql & "and fec_pago= '" & Trim(vlFPago) & "' "
            ''Sql = Sql & "and num_endoso=" & Trim(Lbl_End) & " "
            Sql = Sql & "and num_orden=" & Trim(Msf_BMGrilla.Text) & " "
            ''Set vgRs = vgConectarBD.Execute(Sql)
            Set vgRs = vgConexionBD.Execute(Sql)
            If Not vgRs.EOF Then
                vgSw = True
            End If
            vgRs.Close
            
            'Si no existe el beneficiario se inserta
            If (vgSw = False) Then
                
                Msf_BMGrilla.Col = 0
                
                Sql = ""
                Sql = "INSERT INTO PP_TMAE_PAGTERGARBEN (num_poliza,num_endoso,fec_pago,"
                Sql = Sql & "num_orden,cod_tipoidenben,num_idenben,gls_dirben,cod_direccion,"
                Sql = Sql & "gls_fonoben,gls_correoben,cod_derpen,cod_estpension,cod_viapago,"
                Sql = Sql & "cod_banco,cod_tipcuenta,num_cuenta,cod_sucursal,mto_pension,"
                Sql = Sql & "prc_pension,num_mesespag,num_mesesnodev,mto_pago "
                Sql = Sql & ") VALUES ("
                Sql = Sql & "'" & Trim(Txt_PenPoliza) & "',"
                Sql = Sql & " " & CInt(Trim(Lbl_End)) & ","
                ''Sql = Sql & "'" & Format(Trim(Txt_FecPago), "yyyyMMdd") & "',"
                Sql = Sql & "'" & Trim(vlFPago) & "',"
                Msf_BMGrilla.Col = 0
                Sql = Sql & " " & Trim(Msf_BMGrilla.Text) & ","     'nº orden
                Msf_BMGrilla.Col = 2
                vlident = Trim(Mid(Msf_BMGrilla.Text, 1, InStr(Msf_BMGrilla.Text, "-") - 1))
                Sql = Sql & " " & Trim(vlident) & ","               'tipo ident
                Msf_BMGrilla.Col = 3
                Sql = Sql & "'" & Trim(Msf_BMGrilla.Text) & "',"    'nº ident
                Msf_BMGrilla.Col = 22
                Sql = Sql & "'" & Trim(Msf_BMGrilla.Text) & "',"    'dirección
                Msf_BMGrilla.Col = 23
                Sql = Sql & " " & Trim(Msf_BMGrilla.Text) & ","     'cod dirección
                Msf_BMGrilla.Col = 24
                Sql = Sql & "'" & Trim(Msf_BMGrilla.Text) & "',"    'telefono
                Msf_BMGrilla.Col = 25
                Sql = Sql & "'" & Trim(Msf_BMGrilla.Text) & "',"    'email
                Msf_BMGrilla.Col = 36
                Sql = Sql & "'" & Trim(Msf_BMGrilla.Text) & "',"    'derecho pensión
                Msf_BMGrilla.Col = 8
                Sql = Sql & "'" & Trim(Msf_BMGrilla.Text) & "',"    'estado pensión
                Msf_BMGrilla.Col = 26
                Sql = Sql & "'" & Trim(Msf_BMGrilla.Text) & "',"    'vía pago
                Msf_BMGrilla.Col = 29
                Sql = Sql & "'" & Trim(Msf_BMGrilla.Text) & "',"    'banco
                Msf_BMGrilla.Col = 28
                Sql = Sql & "'" & Trim(Msf_BMGrilla.Text) & "',"    'tipo cuenta
                Msf_BMGrilla.Col = 30
                Sql = Sql & "'" & Trim(Msf_BMGrilla.Text) & "',"    'num cuenta
                Msf_BMGrilla.Col = 27
                Sql = Sql & "'" & Trim(Msf_BMGrilla.Text) & "',"    'sucursal
                Msf_BMGrilla.Col = 11
                Sql = Sql & " " & str(Trim(Msf_BMGrilla.Text)) & "," 'mto pensión
                Msf_BMGrilla.Col = 10
                Sql = Sql & " " & str(Trim(Msf_BMGrilla.Text)) & "," 'prc pensión
                Msf_BMGrilla.Col = 12
                Sql = Sql & " " & CInt(Trim(Msf_BMGrilla.Text)) & "," 'meses trans
                Msf_BMGrilla.Col = 13
                Sql = Sql & " " & CInt(Trim(Msf_BMGrilla.Text)) & "," 'meses no devengados
                Msf_BMGrilla.Col = 14
                Sql = Sql & " " & str(Trim(Msf_BMGrilla.Text)) & ")"  'mto pago '*debe ir esto
                vgConexionBD.Execute Sql
                 
                     'CORPTEC
                     
                      Msf_BMGrilla.Col = 0
                vlLog_tabla = "PP_TMAE_PAGTERGARBEN"
                vlLog_idtabla = "NUM_POLIZA.FEC_PAGO.NUM_ORDEN"
                vlLog_valtabla = "" & Trim(Txt_PenPoliza) & "." & Trim(vlFPago) & "." & Trim(Msf_BMGrilla.Text) & ""
                vlLog_trans = "INS"
                'Call flLog_Tabla
                
            End If
        End If
    Next
    
'2. Generar Endoso desde el cual se registre la nueva estructura de la Póliza

    'Obtener el Máximo Endoso de la Póliza
    vgSql = "SELECT max(num_endoso) as numero FROM PP_TMAE_POLIZA "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "num_poliza = '" & Trim(Txt_PenPoliza) & "' "
    Set vgRegistro = vgConectarBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
        vlNumEndosoNew = vgRegistro!numero + 1
    Else
        'Deshacer la Transacción
        vgConexionBD.RollbackTrans

        'Cerrar Conexión
        vgConectarBD.Close

        MsgBox "Problemas al tratar de generar el Nuevo Número de Endoso de la Póliza seleccionada.", vbCritical, "Proceso Cancelado"
        Cmd_Salir.SetFocus
        Screen.MousePointer = 0

        Exit Sub
    End If
    vgRegistro.Close

'1.1 Insertar los Datos de la Modificación de la Póliza en la Tabla Endoso
    Call flFormatearDatosEndoso
    Call flInsertarEndoso

'1.2 Insertar los Nuevos Datos en la Póliza
    If (flFormatearDatosPolizaDef = False) Then
        'Deshacer la Transacción
        vgConexionBD.RollbackTrans

        'Cerrar Conexión
        vgConectarBD.Close

        Cmd_Salir.SetFocus
        Screen.MousePointer = 0
        Exit Sub
    End If
    Call flInsertarPolizaDef

'1.3 Insertar los Nuevos Datos en los Beneficiarios
    vlTotalFilas = Msf_BMGrilla.rows
    vgI = 1
    While vgI < vlTotalFilas

        If (flFormatearDatosBeneficiarioDef(vgI) = False) Then
            'Deshacer la Transacción
            vgConexionBD.RollbackTrans

            'Cerrar Conexión
            vgConectarBD.Close

            Cmd_Salir.SetFocus
            Screen.MousePointer = 0
            Exit Sub
        End If

        Call flInsertarBeneficiarioDef
        vgI = vgI + 1
    Wend

    'Ejecutar la Transacción
    vgConexionBD.CommitTrans
    
    'Cerrar la Conexión
    vgConectarBD.Close
    
    Cmd_Limpiar_Click
    Cmd_BuscarPol_Click
    
    MsgBox "Se ha registrado correctamente el Pago Garantizado.", vbInformation, "Proceso Finalizado"
    
Exit Sub
Err_CmdGrabar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        If (vgConectarBD.State = adStateOpen) Then
            'Cerrar la Conexión
            vgConectarBD.Close
        End If
        
        'Deshacer la Transacción
        vgConexionBD.RollbackTrans
        
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Limpiar_Click()
 
    Call flLimpiar
    
    vlSw = False
    
    vlSwSeleccionado = False

    ''Call flDeshabilitarIngreso
    SSTab1.Tab = 0

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

    Call fgComboGeneral(vgCodTabla_TipCta, Cmb_TipCuenta)
    Call fgComboGeneral(vgCodTabla_Bco, Cmb_Banco)
    Call fgComboGeneral(vgCodTabla_ViaPago, Cmb_ViaPago)
    Call fgComboSucursal(Cmb_Sucursal, "S")
    
    fgComboTipoIdentificacion Cmb_PenNumIdent
    
    Call fgCargarTablaMoneda(vgCodTabla_TipMon, egTablaMoneda(), vgNumeroTotalTablasMoneda)

Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Msf_BMGrilla_DblClick()
On Error GoTo Err_Grilla

    vlSwMostrar = True
    
    If Msf_BMGrilla.rows = 1 Then
        Exit Sub
    End If
    
    vlPos = Msf_BMGrilla.RowSel
    Msf_BMGrilla.row = vlPos
    Msf_BMGrilla.Col = 0
    If (Msf_BMGrilla.Text = "") And vlPos = 0 Then
        Exit Sub
    End If
        
    Screen.MousePointer = 11
    
    'CARGA DATOS DE LA GRILLA A LOS TEXT
    
    Msf_BMGrilla.Col = 0
    Lbl_BMNumOrd = Msf_BMGrilla.Text
    
    Msf_BMGrilla.Col = 1
    Lbl_BMParentesco = Msf_BMGrilla.Text + " - " + Trim(fgBuscarGlosaElemento(vgCodTabla_Par, Trim(Msf_BMGrilla.Text)))
    
    Msf_BMGrilla.Col = 2
    vlCodTipoIden = fgObtenerCodigo_TextoCompuesto(Msf_BMGrilla.Text)
    If Msf_BMGrilla.Text <> "" Then
        Lbl_BMTipoIdent = Msf_BMGrilla.Text
    End If
    
    Msf_BMGrilla.Col = 3
    vlNumIden = Msf_BMGrilla.Text
    Lbl_BMNumIdent = vlNumIden
    
    Msf_BMGrilla.Col = 4
    Lbl_BMNombre = Msf_BMGrilla.Text
    
    Msf_BMGrilla.Col = 5
    Lbl_BMNombreSeg = Msf_BMGrilla.Text
    
    Msf_BMGrilla.Col = 6
    Lbl_BMApPaterno = Msf_BMGrilla.Text
    
    Msf_BMGrilla.Col = 7
    Lbl_BMApMaterno = Msf_BMGrilla.Text
    
    Msf_BMGrilla.Col = 8
    Lbl_BMEstPension = Msf_BMGrilla.Text
    
    Msf_BMGrilla.Col = 9
    Lbl_EstPago = Msf_BMGrilla.Text
    
    Msf_BMGrilla.Col = 11
    Lbl_BMMtoPension = Format(Msf_BMGrilla.Text, "#,#0.00")
    
    Msf_BMGrilla.Col = 12
    Lbl_MesTrans = Msf_BMGrilla.Text
    
    Msf_BMGrilla.Col = 13
    Lbl_MesNoDev = Msf_BMGrilla.Text
    
    Msf_BMGrilla.Col = 14
    Lbl_MtoPenValPte = Format(Msf_BMGrilla.Text, "#,#0.00") '*revisar

    Msf_BMGrilla.Col = 15
''    If (Msf_BMGrilla.Text <> "") Then
''        Txt_FecPago = DateSerial(Mid((Msf_BMGrilla.Text), 1, 4), Mid((Msf_BMGrilla.Text), 5, 2), Mid((Msf_BMGrilla.Text), 7, 2))
''    Else
        Txt_FecPago = Msf_BMGrilla.Text
''    End If
    
    Msf_BMGrilla.Col = 16
    Lbl_BMSexo = Msf_BMGrilla.Text + " - " + Trim(fgBuscarGlosaElemento(vgCodTabla_Sexo, Trim(Msf_BMGrilla.Text)))
    
    Msf_BMGrilla.Col = 17
    Lbl_BMFecNac = DateSerial(Mid((Msf_BMGrilla.Text), 1, 4), Mid((Msf_BMGrilla.Text), 5, 2), Mid((Msf_BMGrilla.Text), 7, 2))
    
    Msf_BMGrilla.Col = 18 '*debe ir este o prc_pension
    Lbl_BMPrcLegal = Format(Msf_BMGrilla.Text, "#,#0.00")
    
    Msf_BMGrilla.Col = 20
    Lbl_BMMtoPensionGar = Format(Msf_BMGrilla.Text, "#,#0.00")
    
    Msf_BMGrilla.Col = 21
    Lbl_BMGrupoFam = Msf_BMGrilla.Text
    
    Msf_BMGrilla.Col = 22
    Txt_DomicRecep = Msf_BMGrilla.Text

    Msf_BMGrilla.Col = 23
    fgBuscarNombreProvinciaRegion (Msf_BMGrilla.Text)
    Lbl_Departamento = vgNombreRegion
    Lbl_Provincia = vgNombreProvincia
    Lbl_Distrito = vgNombreComuna

    Msf_BMGrilla.Col = 24
    Txt_TelefRecep = Msf_BMGrilla.Text
    
    Msf_BMGrilla.Col = 25
    Txt_EmailRecep = Msf_BMGrilla.Text
    
    Msf_BMGrilla.Col = 26
    Cmb_ViaPago.ListIndex = fgBuscarPosicionCodigoCombo(Msf_BMGrilla.Text, Cmb_ViaPago)
    
    Msf_BMGrilla.Col = 27
    Cmb_Sucursal.ListIndex = fgBuscarPosicionCodigoCombo(Msf_BMGrilla.Text, Cmb_Sucursal)
    
    Msf_BMGrilla.Col = 28
    Cmb_TipCuenta.ListIndex = fgBuscarPosicionCodigoCombo(Msf_BMGrilla.Text, Cmb_TipCuenta)
    
    Msf_BMGrilla.Col = 29
    Cmb_Banco.ListIndex = fgBuscarPosicionCodigoCombo(Msf_BMGrilla.Text, Cmb_Banco)
    
    Msf_BMGrilla.Col = 30
    Txt_NumCta = Msf_BMGrilla.Text
    
    Msf_BMGrilla.Col = 31
    If (Msf_BMGrilla.Text <> "") Then
        Lbl_FecIniPer = DateSerial(Mid((Msf_BMGrilla.Text), 1, 4), Mid((Msf_BMGrilla.Text), 5, 2), Mid((Msf_BMGrilla.Text), 7, 2))
    Else
        Lbl_FecIniPer = Msf_BMGrilla.Text
    End If
    
    Msf_BMGrilla.Col = 32
    If (Msf_BMGrilla.Text <> "") Then
        Lbl_FecFinPer = DateSerial(Mid((Msf_BMGrilla.Text), 1, 4), Mid((Msf_BMGrilla.Text), 5, 2), Mid((Msf_BMGrilla.Text), 7, 2))
    Else
        Lbl_FecFinPer = Msf_BMGrilla.Text
    End If
    
    Msf_BMGrilla.Col = 33
    Lbl_TipoCambio = IIf(IsNull(Msf_BMGrilla.Text), "", Format(Msf_BMGrilla.Text, "#,#0.000"))

    If (Lbl_Estado = clEstPagado) Then
        Lbl_TipCambio = IIf(IsNull(Msf_BMGrilla.Text), "", Format(Msf_BMGrilla.Text, "#,#0.000"))
    Else
        Lbl_TipCambio = Lbl_TipoCambio
    End If
    
    Msf_BMGrilla.Col = 34
    If (Msf_BMGrilla.Text <> "") Then
        Txt_FecRecepcion = DateSerial(Mid((Msf_BMGrilla.Text), 1, 4), Mid((Msf_BMGrilla.Text), 5, 2), Mid((Msf_BMGrilla.Text), 7, 2))
    Else
        Txt_FecRecepcion = Msf_BMGrilla.Text
    End If
    
    Msf_BMGrilla.Col = 35
    Txt_TasaInteres = IIf(IsNull(Msf_BMGrilla.Text), "", Format(Msf_BMGrilla.Text, "#,#0.00"))
    
    
'''    vlSwClic = True
'''    Call flCargarDatosBen(Trim(Txt_PenPoliza), vlCodTipoIden, vlNumIden, (Lbl_End))
'''    vlSwClic = False
'''    If vgCodEstado <> 9 Then 'Póliza No Vigente
'''        Fra_AntRecep.Enabled = True
'''        Txt_NomBen.SetFocus
'''    Else
'''        Call flDeshabilitarIngreso ''**flDesHab
'''    End If
'''    'If vgRs2!Cod_EstPension = "10" Then
'''    If vgRs2!Cod_DerPen = "10" Then 'HQR 03/05/2005
'''        'Fra_Personales.Enabled = False
'''        ''*Fra_Pago.Enabled = False
'''        Fra_FormPgo.Enabled = False
'''    Else
'''        ''*Fra_AntRecep.Enabled = True
'''        ''*Fra_Pago.Enabled = True
'''        Fra_FormPgo.Enabled = True
'''    End If
    
    'vlSwMostrar = True
    vlSwMostrar = False
    
    Screen.MousePointer = 0
Exit Sub
Err_Grilla:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub flInicializaGrilla()

    Msf_BMGrilla.Clear
    Msf_BMGrilla.Cols = 37
    Msf_BMGrilla.rows = 1
    
    Msf_BMGrilla.row = 0
        
    Msf_BMGrilla.Col = 0
    Msf_BMGrilla.Text = "NºOrden"
    Msf_BMGrilla.ColWidth(0) = 800
    Msf_BMGrilla.ColAlignment(0) = 3
    
    Msf_BMGrilla.Col = 1
    Msf_BMGrilla.Text = "Parentesco"
    Msf_BMGrilla.ColWidth(1) = 1200
    
    Msf_BMGrilla.Col = 2
    Msf_BMGrilla.Text = "Tipo Ident."
    Msf_BMGrilla.ColWidth(2) = 1500
    
    Msf_BMGrilla.Col = 3
    Msf_BMGrilla.Text = "NºIdent"
    Msf_BMGrilla.ColWidth(3) = 1000
    
    Msf_BMGrilla.Col = 4
    Msf_BMGrilla.Text = "Nombre"
    Msf_BMGrilla.ColWidth(4) = 1000
    
    Msf_BMGrilla.Col = 5
    Msf_BMGrilla.Text = "Nombre Seg."
    Msf_BMGrilla.ColWidth(5) = 1000
    
    Msf_BMGrilla.Col = 6
    Msf_BMGrilla.Text = "Ap. Paterno"
    Msf_BMGrilla.ColWidth(6) = 1000
    
    Msf_BMGrilla.Col = 7
    Msf_BMGrilla.Text = "Ap. Materno"
    Msf_BMGrilla.ColWidth(7) = 1000
    
    Msf_BMGrilla.Col = 8
    Msf_BMGrilla.Text = "Est. Pensión"
    Msf_BMGrilla.ColWidth(8) = 900
    
    Msf_BMGrilla.Col = 9
    Msf_BMGrilla.Text = "Est. Pago Gar."
    Msf_BMGrilla.ColWidth(9) = 1000
   
    Msf_BMGrilla.Col = 10
    Msf_BMGrilla.Text = "% Pensión"
    Msf_BMGrilla.ColWidth(10) = 900
    
    Msf_BMGrilla.Col = 11
    Msf_BMGrilla.Text = "Mto. Pensión"
    Msf_BMGrilla.ColWidth(11) = 900
    
    Msf_BMGrilla.Col = 12
    Msf_BMGrilla.Text = "Meses Pag."
    Msf_BMGrilla.ColWidth(12) = 900
        
    Msf_BMGrilla.Col = 13
    Msf_BMGrilla.Text = "Meses No Dev."
    Msf_BMGrilla.ColWidth(13) = 900
    
    Msf_BMGrilla.Col = 14
    Msf_BMGrilla.Text = "Pensión No Percibida"
    Msf_BMGrilla.ColWidth(14) = 900
    
    Msf_BMGrilla.Col = 15
    Msf_BMGrilla.Text = "Fecha Pago"
    Msf_BMGrilla.ColWidth(15) = 900
    
    'Columnas Invisibles
    Msf_BMGrilla.Col = 16
    Msf_BMGrilla.Text = "Sexo"
    Msf_BMGrilla.ColWidth(16) = 0
    
    Msf_BMGrilla.Col = 17
    Msf_BMGrilla.Text = "Fec. Nac."
    Msf_BMGrilla.ColWidth(17) = 0
    
    Msf_BMGrilla.Col = 18
    Msf_BMGrilla.Text = "Prc. legal"
    Msf_BMGrilla.ColWidth(18) = 0
    
    Msf_BMGrilla.Col = 19
    Msf_BMGrilla.Text = "Prc. Gar."
    Msf_BMGrilla.ColWidth(19) = 0
    
    Msf_BMGrilla.Col = 20
    Msf_BMGrilla.Text = "Pension Gar."
    Msf_BMGrilla.ColWidth(20) = 0
    
    Msf_BMGrilla.Col = 21
    Msf_BMGrilla.Text = "Grupo Fam."
    Msf_BMGrilla.ColWidth(21) = 0
    
    Msf_BMGrilla.Col = 22
    Msf_BMGrilla.Text = "Domicilio"
    Msf_BMGrilla.ColWidth(22) = 0
    
    Msf_BMGrilla.Col = 23
    Msf_BMGrilla.Text = "Cod Dir"
    Msf_BMGrilla.ColWidth(23) = 0
    
    Msf_BMGrilla.Col = 24
    Msf_BMGrilla.Text = "Telefono"
    Msf_BMGrilla.ColWidth(24) = 0
    
    Msf_BMGrilla.Col = 25
    Msf_BMGrilla.Text = "Email"
    Msf_BMGrilla.ColWidth(25) = 0
    
    Msf_BMGrilla.Col = 26
    Msf_BMGrilla.Text = "Via Pago"
    Msf_BMGrilla.ColWidth(26) = 0
    
    Msf_BMGrilla.Col = 27
    Msf_BMGrilla.Text = "Sucursal"
    Msf_BMGrilla.ColWidth(27) = 0
    
    Msf_BMGrilla.Col = 28
    Msf_BMGrilla.Text = "Tipo Cta."
    Msf_BMGrilla.ColWidth(28) = 0
    
    Msf_BMGrilla.Col = 29
    Msf_BMGrilla.Text = "Banco"
    Msf_BMGrilla.ColWidth(29) = 0
    
    Msf_BMGrilla.Col = 30
    Msf_BMGrilla.Text = "Nº Cta."
    Msf_BMGrilla.ColWidth(30) = 0
    
    Msf_BMGrilla.Col = 31
    Msf_BMGrilla.Text = "Inicio Per"
    Msf_BMGrilla.ColWidth(31) = 0
    
    Msf_BMGrilla.Col = 32
    Msf_BMGrilla.Text = "Fin Per"
    Msf_BMGrilla.ColWidth(32) = 0
    
    Msf_BMGrilla.Col = 33
    Msf_BMGrilla.Text = "Tipo Cambio"
    Msf_BMGrilla.ColWidth(33) = 0
    
    Msf_BMGrilla.Col = 34
    Msf_BMGrilla.Text = "Fec Recep"
    Msf_BMGrilla.ColWidth(34) = 0
    
    Msf_BMGrilla.Col = 35
    Msf_BMGrilla.Text = "Tasa Int."
    Msf_BMGrilla.ColWidth(35) = 0
    
    Msf_BMGrilla.Col = 36
    Msf_BMGrilla.Text = "Der Pen"
    Msf_BMGrilla.ColWidth(36) = 0
    
End Sub

'FUNCION QUE CARGA LA GRILLA CON LOS DATOS DEL BENEFICIARIO
Function flCargaGrilla(iPoliza, iEndoso)
Dim vlRegBen As ADODB.Recordset
On Error GoTo Err_Cargar
    
    Call flInicializaGrilla
    
    vgSql = ""
    vgSql = "SELECT num_orden,cod_par,cod_tipoidenben,num_idenben,gls_nomben,"
    vgSql = vgSql & "gls_nomsegben,gls_patben,gls_matben,cod_estpension,num_poliza,"
    vgSql = vgSql & "cod_sexo,fec_nacben,prc_pensionleg,mto_pension,mto_pensiongar,"
    vgSql = vgSql & "cod_grufam,gls_dirben,cod_direccion,gls_fonoben,gls_correoben,"
    vgSql = vgSql & "prc_pension,prc_pensiongar,cod_derpen "
    vgSql = vgSql & ",cod_viapago,cod_sucursal,cod_tipcuenta,cod_banco,num_cuenta"
    vgSql = vgSql & ",fec_inipagopen,fec_terpagopengar "
    vgSql = vgSql & "FROM PP_TMAE_BEN "
    vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza) & "' and "
    vgSql = vgSql & "num_endoso = '" & Trim(Lbl_End) & "' "
    Set vlRegBen = vgConexionBD.Execute(vgSql)
    Do While Not vlRegBen.EOF
        vlNumOrd = Trim(vlRegBen!Num_Orden)
        vlCodPar = Trim(vlRegBen!Cod_Par)
        vlTipoIdent = (vlRegBen!Cod_TipoIdenBen & " - " & fgBuscarNombreTipoIden(vlRegBen!Cod_TipoIdenBen))
        vlNumIdent = Trim(vlRegBen!Num_IdenBen)
        vlNom = Trim(vlRegBen!Gls_NomBen)
        vlNomSeg = IIf(IsNull(vlRegBen!Gls_NomSegBen), "", (vlRegBen!Gls_NomSegBen))
        vlApPat = Trim(vlRegBen!Gls_PatBen)
        vlApMat = IIf(IsNull(vlRegBen!Gls_MatBen), "", (vlRegBen!Gls_MatBen))
        vlEstPension = Trim(vlRegBen!Cod_EstPension)
        vlPrcPension = Format(vlRegBen!Prc_Pension, "#,#0.00")
        vlMtoPension = Format(vlRegBen!Mto_Pension, "#,#0.00")
        vlMesesPag = 0  '2da tabla
        vlMesesNoDev = 0 '2da tabla
        vlPenNoPer = 0  'No se se donde se obtiene
        vlFecPago = ""  '2da tabla
        vlSexo = vlRegBen!Cod_Sexo
        vlFecNac = Trim(vlRegBen!Fec_NacBen)
        vlPrcLeg = vlRegBen!Prc_PensionLeg
        vlPrcGar = vlRegBen!Prc_PensionGar
        vlMtoPenGar = vlRegBen!Mto_PensionGar
        vlGruFam = vlRegBen!Cod_GruFam
        vlDirec = Trim(vlRegBen!Gls_DirBen)
        vlCodDir = vlRegBen!Cod_Direccion
        vlFono = IIf(IsNull(vlRegBen!Gls_FonoBen), "", Trim(vlRegBen!Gls_FonoBen))
        vlEmail = IIf(IsNull(vlRegBen!Gls_CorreoBen), "", Trim(vlRegBen!Gls_CorreoBen))
        vlViaPago = vlRegBen!Cod_ViaPago
        vlSucursal = vlRegBen!Cod_Sucursal
        vlTipcta = vlRegBen!Cod_TipCuenta
        vlBanco = vlRegBen!Cod_Banco
        vlNumCta = IIf(IsNull(vlRegBen!Num_Cuenta), "", Trim(vlRegBen!Num_Cuenta))
'I--- ABV 28/09/2007 ---
        vlFecIni = vlRegBen!Fec_IniPagoPen '* VA ESTO?
        'Definir como Fecha de Inicio, la Fecha de Efecto correspondiente
        vgPalabraAux = fgValidaFechaEfecto(DateSerial(Mid(vlFecIni, 1, 4), Mid(vlFecIni, 5, 2), Mid(vlFecIni, 7, 4)), Trim(Txt_PenPoliza), CInt(vlNumOrd))
        If (vgPalabraAux <> "") Then
            vlFecIni = Format(vgPalabraAux, "yyyymmdd")
        End If
'F--- ABV 28/09/2007 ---
        vlFecFin = IIf(IsNull(vlRegBen!Fec_TerPagoPenGar), "", Trim(vlRegBen!Fec_TerPagoPenGar)) '*va esto?
        vlEstPago = clEstNoPagado
        vlTipCambio = ""
        vlFecRecep = ""
        vlTasaInt = 0
        vlDerpen = vlRegBen!Cod_DerPen
        
    ''    Sql = "SELECT cod_tipoidenben,num_idenben,cod_estpension,prc_pension,mto_pension,"
    ''    Sql = Sql & "num_mesespag,num_mesesnodev,fec_pago,gls_dirben,cod_direccion,"
    ''    Sql = Sql & "gls_fonoben,gls_correoben,cod_viapago,cod_sucursal,cod_tipcuenta,"
    ''    Sql = Sql & "cod_banco,num_cuenta "
    ''    Sql = Sql & "FROM PP_TMAE_PAGTERGARBEN "
    ''    Sql = Sql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza) & "' and "
    ''    Sql = Sql & "num_orden = " & VlNumOrd & " and "
    ''    Sql = Sql & "num_endoso = '" & Trim(Lbl_End) & "' "'11-09-2007
    
        Sql = ""
        Sql = "SELECT cod_tipoidenben,num_idenben,cod_estpension,prc_pension,b.mto_pension,"
        Sql = Sql & "num_mesespag,num_mesesnodev,b.fec_pago,gls_dirben,cod_direccion,"
        Sql = Sql & "gls_fonoben,gls_correoben,cod_viapago,cod_sucursal,cod_tipcuenta,"
        Sql = Sql & "cod_banco,num_cuenta,mto_pago,fec_inipergarpag,fec_finpergar,"
        Sql = Sql & "mto_valmoneda,fec_solpago,prc_tasaint,cod_derpen,fec_inipencia,"
        Sql = Sql & "cod_moneda " '',cod_scomp
        Sql = Sql & "FROM PP_TMAE_PAGTERGARBEN B, PP_TMAE_PAGTERGAR P " '',MA_TPAR_TABCOD m
        Sql = Sql & "WHERE p.num_poliza=b.num_poliza and p.fec_pago=b.fec_pago and "
        Sql = Sql & "b.num_poliza = '" & Trim(Txt_PenPoliza) & "' and "
        Sql = Sql & "num_orden = " & vlNumOrd & " "
''        Sql = Sql & "and b.num_endoso = '" & Trim(Lbl_End) & "' "
        ''Sql = Sql & "and p.cod_moneda = m.cod_elemento and "
        ''Sql = Sql & "m.cod_tabla='" & vgCodTabla_TipMon & "' "
        Set vgRs = vgConexionBD.Execute(Sql)
        If Not vgRs.EOF Then
           vlTipoIdent = (vgRs!Cod_TipoIdenBen & " - " & fgBuscarNombreTipoIden(vgRs!Cod_TipoIdenBen))
           vlNumIdent = Trim(vgRs!Num_IdenBen)
           vlEstPension = Trim(vgRs!Cod_EstPension)
           vlPrcPension = Format(vgRs!Prc_Pension, "#,#0.00")
           vlMtoPension = Format(vgRs!Mto_Pension, "#,#0.00")
           If (vlMtoPension <> "") Then Lbl_MtoPensionRef = vlMtoPension
           vlMesesPag = vgRs!num_mesespag
           vlMesesNoDev = vgRs!num_mesesnodev
           vlPenNoPer = vgRs!mto_pago
           If Not IsNull(vgRs!Fec_Pago) Then
            vlFecPago = DateSerial(Mid((vgRs!Fec_Pago), 1, 4), Mid((vgRs!Fec_Pago), 5, 2), Mid((vgRs!Fec_Pago), 7, 2))
            'Txt_FecPago = vlFecPago
           End If
           vlDirec = Trim(vgRs!Gls_DirBen)
           vlCodDir = vgRs!Cod_Direccion
           vlFono = IIf(IsNull(vgRs!Gls_FonoBen), "", Trim(vgRs!Gls_FonoBen))
           vlEmail = IIf(IsNull(vgRs!Gls_CorreoBen), "", Trim(vgRs!Gls_CorreoBen))
           vlViaPago = vgRs!Cod_ViaPago
           vlSucursal = vgRs!Cod_Sucursal
           vlTipcta = vgRs!Cod_TipCuenta
           vlBanco = vgRs!Cod_Banco
           vlNumCta = IIf(IsNull(vgRs!Num_Cuenta), "", Trim(vgRs!Num_Cuenta))
           vlFecIni = vgRs!fec_inipergarpag
           If (vlFecIni <> "") Then Lbl_FecIniPerGar = DateSerial(Mid(vlFecIni, 1, 4), Mid(vlFecIni, 5, 2), Mid(vlFecIni, 7, 2))
           vlFecFin = vgRs!fec_finpergar
           If (vlFecFin <> "") Then Lbl_FecFinPerGar = DateSerial(Mid(vlFecFin, 1, 4), Mid(vlFecFin, 5, 2), Mid(vlFecFin, 7, 2))
           vlEstPago = clEstPagado
           vlTipCambio = vgRs!Mto_ValMoneda
           'If (vlTipCambio <> "") Then Lbl_TipoCambio = Format(vlTipCambio, "#,#0.000")
           vlFecRecep = vgRs!fec_solpago
           ''If (vlFecRecep <> "") Then
            ''Txt_FecRecepcion = DateSerial(Mid(vlFecRecep, 1, 4), Mid(vlFecRecep, 5, 2), Mid(vlFecRecep, 7, 2))
           ''End If
           vlTasaInt = vgRs!prc_tasaint
           ''If (vlTasaInt <> "") Then
            ''Txt_TasaInteres = Format(vlTasaInt, "#,#0.00")
           ''End If
           vlDerpen = vgRs!Cod_DerPen
           'If (vgRs!Fec_IniPenCia <> "") Then Lbl_FecIniPago = DateSerial(Mid((vgRs!Fec_IniPenCia), 1, 4), Mid((vgRs!Fec_IniPenCia), 5, 2), Mid((vgRs!Fec_IniPenCia), 7, 2))
     
            vlCodMoneda = vgRs!Cod_Moneda
            
            vlCodMonedaScomp = fgObtenerCodMonedaScomp(egTablaMoneda(), vgNumeroTotalTablasMoneda, vlCodMoneda)
            Lbl_MonPension(0) = vlCodMonedaScomp
            Lbl_MonPension(1) = vlCodMonedaScomp
            Lbl_MonPension(2) = vlCodMonedaScomp

        End If
        
        
        Msf_BMGrilla.AddItem (vlNumOrd) & vbTab & (vlCodPar) & vbTab & _
                        " " & vlTipoIdent & vbTab & (vlNumIdent) & vbTab & _
                        (vlNom) & vbTab & (vlNomSeg) & vbTab & _
                        (vlApPat) & vbTab & (vlApMat) & vbTab & _
                        (vlEstPension) & vbTab & (vlEstPago) & vbTab & _
                        (vlPrcPension) & vbTab & (vlMtoPension) & vbTab & _
                        (vlMesesPag) & vbTab & (vlMesesNoDev) & vbTab & _
                        (vlPenNoPer) & vbTab & (vlFecPago) & vbTab & _
                        (vlSexo) & vbTab & (vlFecNac) & vbTab & _
                        (vlPrcLeg) & vbTab & (vlPrcGar) & vbTab & _
                        (vlMtoPenGar) & vbTab & (vlGruFam) & vbTab & _
                        (vlDirec) & vbTab & (vlCodDir) & vbTab & _
                        (vlFono) & vbTab & (vlEmail) & vbTab & _
                        (vlViaPago) & vbTab & (vlSucursal) & vbTab & _
                        (vlTipcta) & vbTab & (vlBanco) & vbTab & _
                        (vlNumCta) & vbTab & (vlFecIni) & vbTab & _
                        (vlFecFin) & vbTab & (vlTipCambio) & vbTab & _
                        (vlFecRecep) & vbTab & (vlTasaInt) & vbTab & _
                        (vlDerpen)
        vlRegBen.MoveNext
    Loop

Exit Function
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'FUNCION QUE CARGA LOS DATOS DEL BENEFICIARIO A LOS TEXT
Function flCargarDatosBen(iPoliza, iCodTipoIden, iNumIden, iEndoso)
On Error GoTo Err_CargarDatos
    flCargarDatosBen = True
      
    'busca el beneficiario seleccionado en la bd
    If vlSwClic = False Then
        vgSql = "select * from PP_TMAE_BEN where "
        vgSql = vgSql & "num_poliza = '" & Trim(iPoliza) & "' AND "
        vgSql = vgSql & "cod_tipoidenben = " & Trim(iCodTipoIden) & " AND "
        vgSql = vgSql & "num_idenben = '" & Trim(iNumIden) & "' "
        vgSql = vgSql & "order by num_endoso desc "
    Else
        vgSql = "select * from PP_TMAE_BEN where "
        vgSql = vgSql & "num_poliza = '" & Trim(Txt_PenPoliza) & "' and "
        vgSql = vgSql & "num_endoso= " & Trim(Lbl_End) & " AND "
        vgSql = vgSql & "cod_tipoidenben = " & Trim(iCodTipoIden) & " AND "
        vgSql = vgSql & "num_idenben = '" & Trim(iNumIden) & "' "
    End If
      
    ' si existe el beneficiario llena los text
    Set vgRs2 = vgConexionBD.Execute(vgSql)
    If Not vgRs2.EOF Then
        'si no ha hecho clic en la grilla
        If vlSwClic = False Then
            If Not IsNull(vgRs2!Gls_NomSegBen) Then
                Txt_NomSegBen = (vgRs2!Gls_NomSegBen)
                Lbl_PenNombre = Trim(vgRs2!Gls_NomBen) + " " + (vgRs2!Gls_NomSegBen) + " " + Trim(vgRs2!Gls_PatBen) + " " + IIf(IsNull(vgRs2!Gls_MatBen), "", Trim(vgRs2!Gls_MatBen))
            Else
                Txt_NomSegBen = ""
                vlNomSegBen = ""
                Lbl_PenNombre = Trim(vgRs2!Gls_NomBen) + " " + (vlNomSegBen) + " " + Trim(vgRs2!Gls_PatBen) + " " + IIf(IsNull(vgRs2!Gls_MatBen), "", Trim(vgRs2!Gls_MatBen))
            End If
            Lbl_End = vgRs2!num_endoso
            
            'Llena el combo de tipo identificación beneficiario
            vgPalabra = Trim(vgRs2!Cod_TipoIdenBen)
            vgI = fgBuscarPosicionCodigoCombo(vgPalabra, Cmb_PenNumIdent)
            If (Cmb_PenNumIdent.ListCount > 0) Then
                Cmb_PenNumIdent.ListIndex = vgI
            End If
    
            Txt_PenNumIdent = Trim(vgRs2!Num_IdenBen)
        End If
        'Antecedentes Personales
        'vlCodTp = vgRs2!Cod_TipPension
        Lbl_BMTipoIdent = (vgRs2!Cod_TipoIdenBen & " - " & fgBuscarNombreTipoIden(vgRs2!Cod_TipoIdenBen))
        Lbl_BMNumIdent = vgRs2!Num_IdenBen
        Lbl_BMNombre = vgRs2!Gls_NomBen
        
        If Not IsNull(vgRs2!Gls_NomSegBen) Then
            Lbl_BMNombreSeg = (vgRs2!Gls_NomSegBen)
        Else
            Lbl_BMNombreBen = ""
            vlNomSegBen = ""
        End If
        
        Lbl_BMApPaterno = vgRs2!Gls_PatBen
        If Not IsNull(vgRs2!Gls_MatBen) Then
            Lbl_BMApMaterno = vgRs2!Gls_MatBen
        Else
            Lbl_BMApMaterno = ""
        End If
        
        Lbl_BMParentesco = vgRs2!Cod_Par
        
        vlAnno = Mid(vgRs2!Fec_NacBen, 1, 4)
        vlMes = Mid(vgRs2!Fec_NacBen, 5, 2)
        vlDia = Mid(vgRs2!Fec_NacBen, 7, 2)
        Lbl_BMFecNac = DateSerial(vlAnno, vlMes, vlDia)
''        If Not IsNull(vgRs2!Fec_FallBen) Then
''            Lbl_FecFall = DateSerial(Mid((vgRs2!Fec_FallBen), 1, 4), Mid((vgRs2!Fec_FallBen), 5, 2), Mid((vgRs2!Fec_FallBen), 7, 2))
''        Else
''            Lbl_FecFall = ""
''        End If
        Lbl_BMNumOrd = Trim(vgRs2!Num_Orden)
        vlNumOrd = Lbl_NumOrd
'''        Lbl_SitInv = Trim(vgRs2!Cod_SitInv)
        Txt_DomBen = IIf(IsNull(vgRs2!Gls_DirBen), "", vgRs2!Gls_DirBen)
        'vlCodDir = IIf(IsNull(vgRs2!Cod_Direccion), 0, vgRs2!Cod_Direccion)
        vlCodDir = (vgRs2!Cod_Direccion)
   
        If Not IsNull(vgRs2!Cod_Direccion) Then
            vlCodDir = vgRs2!Cod_Direccion
            Call fgBuscarNombreComunaProvinciaRegion(CStr(vlCodDir))
            Lbl_Departamento = vgNombreRegion
            Lbl_Provincia = vgNombreProvincia
            Lbl_Distrito = vgNombreComuna
        End If

        If IsNull(vgRs2!Gls_FonoBen) Then
            Txt_FonoBen = ""
        Else
            Txt_FonoBen = Trim(vgRs2!Gls_FonoBen)
        End If
        If IsNull(vgRs2!Gls_CorreoBen) Then
            Txt_CorreoBen = ""
        Else
            Txt_CorreoBen = Trim(vgRs2!Gls_CorreoBen)
        End If
        vlCodViaPag = vgRs2!Cod_ViaPago
        vlCodSuc = vgRs2!Cod_Sucursal
        vlCodTipcta = vgRs2!Cod_TipCuenta
        vlCodBco = vgRs2!Cod_Banco
        If Not IsNull(vgRs2!Num_Cuenta) Then
            Txt_NumCta = vgRs2!Num_Cuenta
            vlSwNumCta = True
        Else
            vlSwNumCta = False
            Txt_NumCta = ""
        End If
        vlCodIns = IIf(IsNull(vgRs2!Cod_InsSalud), "00", vgRs2!Cod_InsSalud)
        vlCodModPago = IIf(IsNull(vgRs2!Cod_ModSalud), "", vgRs2!Cod_ModSalud)
        Txt_MtoPago = IIf(IsNull(vgRs2!Mto_PlanSalud), "", Format(vgRs2!Mto_PlanSalud, "#,#0.000"))

        Lbl_BMMtoPension = Format(vgRs2!Mto_Pension, "#,#0.00")
        Lbl_BMMtoPensionGar = Format(vgRs2!Mto_PensionGar, "#,#0.00")
        
'''        'CMV-20061031 I
'''        'Mostrar Monto de Pensión en Quiebra
'''        vlUltimoPerPago = flUltimoPeriodoCerrado(iPoliza)
'''        If fgObtieneParametrosQuiebra(vlUltimoPerPago, vlPrcCastigoQui, vlTopeMaxQui) Then
'''            Lbl_MtoPQ.Visible = True
'''            Lbl_MtoPensionQui.Visible = True
'''            Lbl_MtoPensionQui = flBuscarPensionQuiebra(vlUltimoPerPago, iPoliza, VlNumOrd, clCodTipReceptorR)
'''            Lbl_MtoPensionQui = Format(Lbl_MtoPensionQui, "#,#0.00")
'''        End If
'''        'CMV-20061031 F
        
        'se llena el combo comuna, los campos provincia y region
'        If Cmb_Comuna.Text <> "" Then
'            For vlI = 0 To Cmb_Comuna.ListCount - 1
'                If Cmb_Comuna.ItemData(vlI) = vlCodDir Then
'                    Cmb_Comuna.ListIndex = vlI
'                    Exit For
'                End If
'            Next vlI
'        End If
  
''        'se carga el combo de via de pago
''        Call flBusEle(vgCodTabla_ViaPago, vlCodViaPag)
''        vlSwMostrar = False
''        Call flBusPos(vlCodViaPag, vlElemento, Cmb_ViaPago)
''        vlSwMostrar = True

'''        'se carga la sucursal
'''        Call flBusSuc(vlCodSuc)
'''        Call flBusPos(vlCodSuc, vlElemento, Cmb_Suc)
'''
'''        'se carga el combo de tipo de cuenta
'''        Call flBusEle(vgCodTabla_TipCta, vlCodTipcta)
'''        Call flBusPos(vlCodTipcta, vlElemento, Cmb_TipCta)
'''
'''        'se carga el combo banco
'''        Call flBusEle(vgCodTabla_Bco, vlCodBco)
'''        Call flBusPos(vlCodBco, vlElemento, Cmb_Banco)
'''
'''        'se carga el combo de institucion de salud
'''        Call flBusEle(vgCodTabla_InsSal, vlCodIns)
'''        Call flBusPos(vlCodIns, vlElemento, Cmb_Inst)
'''
'''        'se carga el combo de modalidad de pago
'''        Call flBusEle(vgCodTabla_ModPago, vlCodModPago)
'''        Call flBusPos(vlCodModPago, vlElemento, Cmb_ModPago)
'''        'Call flBusEle(vgCodTabla_ModPago, vlCodModPago2)
'''        'Call flBusPos(vlCodModPago2, vlElemento, Cmb_ModPago2)

        flCargarDatosBen = False
        Call Cmb_ViaPago_Click
        'Call Cmb_Inst_Click 'HQR 05/04/2005
        
'''        If vgRs2!Cod_EstPension = "10" Then
'''            Call flDesHabSinDer
'''        Else
'''            Fra_Pago.Enabled = True
'''            Fra_Salud.Enabled = True
'''        End If
        vlSwNumCta = False
    'si no existe el beneficiario
    Else
        MsgBox "El Número de Póliza Ingresado no está Registrado", vbCritical, "Operación Cancelada"
        Txt_PenPoliza.SetFocus
        Exit Function
    End If
    
Exit Function
Err_CargarDatos:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flLimpiar()
    
    Txt_FecRecepcion.Text = ""
    Txt_FecPago.Text = ""
    Txt_TasaInteres.Text = ""
    Lbl_TipoCambio.Caption = ""
    
    Lbl_BMNumOrd.Caption = ""
    Lbl_BMTipoIdent.Caption = ""
    Lbl_BMNumIdent.Caption = ""
    Lbl_BMNombre.Caption = ""
    Lbl_BMNombreSeg.Caption = ""
    Lbl_BMApPaterno.Caption = ""
    Lbl_BMApMaterno.Caption = ""
    Lbl_BMParentesco.Caption = ""
    Lbl_BMSexo.Caption = ""
    Lbl_BMFecNac.Caption = ""
    Lbl_BMPrcLegal.Caption = ""
    Lbl_BMMtoPension.Caption = ""
    Lbl_BMMtoPensionGar.Caption = ""
    Lbl_BMGrupoFam.Caption = ""
    Lbl_BMEstPension.Caption = ""
    
    Txt_DomicRecep = ""
    Lbl_Departamento.Caption = ""
    Lbl_Provincia.Caption = ""
    Lbl_Distrito.Caption = ""
    Txt_TelefRecep = ""
    Txt_EmailRecep.Text = ""
    
    Lbl_EstPago.Caption = ""
    Lbl_TipCambio.Caption = ""
    Lbl_FecIniPer.Caption = ""
    Lbl_FecFinPer.Caption = ""
    Lbl_MesTrans.Caption = ""
    Lbl_MesNoDev.Caption = ""
    Lbl_MtoPenValPte.Caption = ""
    
    If (Cmb_ViaPago.ListCount > 0) Then Cmb_ViaPago.ListIndex = 0
    If (Cmb_Sucursal.ListCount > 0) Then Cmb_Sucursal.ListIndex = 0
    ''*Call fgComboSucursal(Cmb_Sucursal, "S")
    ''*Cmb_Sucursal.Enabled = False
    If (Cmb_TipCuenta.ListCount > 0) Then Cmb_TipCuenta.ListIndex = 0
    If (Cmb_Banco.ListCount > 0) Then Cmb_Banco.ListIndex = 0
    Txt_NumCta.Text = ""
    
    Call Cmb_ViaPago_Click
    
    vlSwSeleccionado = False
    

End Function

Function flDeshabilitarIngreso()

    Fra_Poliza.Enabled = True
    Fra_DetallePago.Enabled = False
    ''SSTab1.Enabled = False
    Fra_AntRecep.Enabled = False
    Fra_FormPgo.Enabled = False
    Cmd_Calcular.Enabled = False
    Cmd_BMSumar.Enabled = False

End Function

Function flHabilitarIngreso()

    Fra_Poliza.Enabled = False
    Fra_DetallePago.Enabled = True
    ''SSTab1.Enabled = True
    Fra_AntRecep.Enabled = True
    Fra_FormPgo.Enabled = True
    Cmd_Calcular.Enabled = True
    Cmd_BMSumar.Enabled = True

End Function

Function flCargarDatosPol(iPoliza, iNumEnd)
On Error GoTo Err_CDP
        
    vgSql = ""
    vgSql = "select fec_dev,fec_inipencia,fec_inipagopen,fec_finpergar,mto_pension,"
    vgSql = vgSql & "cod_moneda " ',m.cod_scomp
    vgSql = vgSql & "from PP_TMAE_POLIZA p " ',MA_TPAR_TABCOD m
    ''vgSql = vgSql & "p.cod_moneda = m.cod_elemento and "
    ''vgSql = vgSql & "m.cod_tabla='" & vgCodTabla_TipMon & "' and "
    vgSql = vgSql & "where num_poliza = '" & Trim(iPoliza) & "' and "
    vgSql = vgSql & "num_endoso = '" & Trim(iNumEnd) & "'"
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
        If IsNull(vgRs!fec_dev) Then
            Lbl_FecDevengue = ""
        Else
            Lbl_FecDevengue = DateSerial(Mid((vgRs!fec_dev), 1, 4), Mid((vgRs!fec_dev), 5, 2), Mid((vgRs!fec_dev), 7, 2))
        End If
        
        If IsNull(vgRs!Fec_IniPenCia) Then
            Lbl_FecIniPago = ""
        Else
            Lbl_FecIniPago = DateSerial(Mid((vgRs!Fec_IniPenCia), 1, 4), Mid((vgRs!Fec_IniPenCia), 5, 2), Mid((vgRs!Fec_IniPenCia), 7, 2))
        End If
        
        If IsNull(vgRs!Fec_IniPagoPen) Then
            Lbl_FecIniPerGar = ""
        Else
            Lbl_FecIniPerGar = DateSerial(Mid((vgRs!Fec_IniPagoPen), 1, 4), Mid((vgRs!Fec_IniPagoPen), 5, 2), Mid((vgRs!Fec_IniPagoPen), 7, 2))
        End If
        
        If IsNull(vgRs!fec_finpergar) Then
            Lbl_FecFinPerGar = ""
        Else
            Lbl_FecFinPerGar = DateSerial(Mid((vgRs!fec_finpergar), 1, 4), Mid((vgRs!fec_finpergar), 5, 2), Mid((vgRs!fec_finpergar), 7, 2))
        End If
        
        Lbl_MtoPensionRef = vgRs!Mto_Pension
    
        vlCodMoneda = vgRs!Cod_Moneda
        vlCodMonedaScomp = fgObtenerCodMonedaScomp(egTablaMoneda(), vgNumeroTotalTablasMoneda, vlCodMoneda)
        Lbl_MonPension(0) = vlCodMonedaScomp
        Lbl_MonPension(1) = vlCodMonedaScomp
        Lbl_MonPension(2) = vlCodMonedaScomp
        
    End If
    vgRs.Close
            
Exit Function
Err_CDP:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
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

Function flValidaDatos() As Boolean
On Error GoTo Err_valdatos
    flValidaDatos = False
     
    If (Trim(Lbl_BMNumOrd.Caption) = "") Then
        MsgBox "Debe Seleccionar el Beneficiario.", vbCritical, "Operación Cancelada"
        Exit Function
    End If
    
    'Valida la información del Detalle de Pago
    If Trim(Txt_FecRecepcion) = "" Then
       MsgBox "Debe Ingresar la Fecha de Recepción del Pago.", vbCritical, "Operación Cancelada"
       Txt_FecRecepcion.SetFocus
       Exit Function
    Else
      If (flValidaFecha(Txt_FecRecepcion) = False) Then
          Txt_FecRecepcion.SetFocus
          Exit Function
      End If
    End If
    
    If Trim(Txt_FecPago) = "" Then
       MsgBox "Debe Ingresar la Fecha de Pago.", vbCritical, "Operación Cancelada"
       Txt_FecPago.SetFocus
       Exit Function
    Else
      If (flValidaFecha(Txt_FecPago) = False) Then
          Txt_FecPago.SetFocus
          Exit Function
      End If
    End If
    
    If (CDate(Txt_FecRecepcion) > CDate(Txt_FecPago)) Then
       MsgBox "La Fecha de Pago debe ser mayor a la Fecha de Recepción", vbCritical, "Operación Cancelada"
       Txt_FecPago.SetFocus
       Exit Function
    End If
    
    If Trim(Txt_TasaInteres) = "" Then
       MsgBox "Debe Ingresar la Tasa de Interés Anual del Pago.", vbCritical, "Operación Cancelada"
       Txt_TasaInteres.SetFocus
       Exit Function
    Else
      If Not IsNumeric(Txt_TasaInteres) Then
          Txt_TasaInteres.SetFocus
          Exit Function
      End If
    End If
    
    'Valida información del Beneficiario
    If (Trim(Txt_DomicRecep.Text) = "") Then
        MsgBox "Debe Ingresar la Dirección del Beneficiario.", vbCritical, "Operación Cancelada"
        Txt_DomicRecep.SetFocus
        Exit Function
    End If
    If (Trim(Lbl_Departamento) = "") Then
        MsgBox "Debe Ingresar la Ubicación del Beneficiario.", vbCritical, "Operación Cancelada"
        Cmd_BuscarDir.SetFocus
        Exit Function
    End If
    If (Trim(Txt_TelefRecep) = "") Then
'        MsgBox "Debe Ingresar el Teléfono del Beneficiario.", vbCritical, "Operación Cancelada"
'        Txt_TelefRecep.SetFocus
'        Exit Function
    End If
'    If (Trim(Txt_EmailRecep) = "") Then
'        MsgBox "Debe Ingresar el Email del Beneficiario.", vbCritical, "Operación Cancelada"
'        Txt_EmailRecep.SetFocus
'        Exit Function
'    End If

    'Valida la Forma de Pago
    vlCodViaPgo = Trim(Mid(Cmb_ViaPago.Text, 1, (InStr(1, Cmb_ViaPago.Text, "-") - 1)))
    vlCodSucursal = Trim(Mid(Cmb_Sucursal.Text, 1, (InStr(1, Cmb_Sucursal.Text, "-") - 1)))
    vlCodTipCuenta = Trim(Mid(Cmb_TipCuenta.Text, 1, (InStr(1, Cmb_TipCuenta.Text, "-") - 1)))
    vlCodBco = Trim(Mid(Cmb_Banco.Text, 1, (InStr(1, Cmb_Banco.Text, "-") - 1)))
    
    If vlCodViaPgo = "00" Then
       MsgBox "Debe Seleccionar Forma de Pago", vbCritical, "Operación Cancelada"
       Cmb_ViaPago.SetFocus
       Screen.MousePointer = 0
       Exit Function
    End If
    If vlCodViaPgo = "01" Then
       If vlCodSucursal = "0000" Then
            MsgBox "Debe seleccionar la Sucursal de la Vía de Pago", vbCritical, "Falta Información"
            Cmb_Sucursal.SetFocus
            Screen.MousePointer = 0
            Exit Function
       End If
    End If
    If vlCodViaPgo = "02" Or vlCodViaPgo = "03" Then
       If vlCodTipCuenta = "00" Then
          MsgBox "Debe seleccionar el tipo de Cuenta", vbCritical, "Falta Información"
          Cmb_TipCuenta.SetFocus
          Screen.MousePointer = 0
          Exit Function
       End If
       If vlCodBco = "00" Then
          MsgBox "Debe seleccionar el Banco", vbCritical, "Falta Información"
          Screen.MousePointer = 0
          Exit Function
       End If
       If Trim(Txt_NumCta) = "" Then
          MsgBox "Debe ingresar el número de cuenta", vbCritical, "Falta Información"
          Txt_NumCta.SetFocus
          Screen.MousePointer = 0
          Exit Function
       End If
    End If

    flValidaDatos = True
Exit Function
Err_valdatos:
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
   
   vlArchivo = strRpt & "PP_Rpt_ConPagoPG.rpt"   '\Reportes
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
        vlNom = vlRegistro!Gls_NomBen + " " + IIf(IsNull(vlRegistro!Gls_NomSegBen), "", (vlRegistro!Gls_NomSegBen)) + " " + vlRegistro!Gls_PatBen + " " + IIf(IsNull(vlRegistro!Gls_MatBen), "", (vlRegistro!Gls_MatBen))
    End If
    vlRegistro.Close
  
   vgQuery = "{PP_TMAE_PAGTERGARBEN.NUM_POLIZA} = '" & Trim(Txt_PenPoliza) & "' AND "
   vgQuery = vgQuery & "{MA_TPAR_TABCODViaPago.COD_TABLA} = '" & vgCodTabla_ViaPago & "' "
   
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
   Rpt_Imprimir.Formulas(6) = "Nombre = '" & Trim(vlNom) & "'"
   
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

Private Sub Cmd_Imprimir_Click()
 If Fra_Poliza.Enabled = False Then
    flImpresion
 End If
End Sub

Private Sub Cmd_Eliminar_Click()
On Error GoTo Err_Eliminar

    If (Fra_Poliza.Enabled = True) Then
        Exit Sub
    End If
    
    vgSw = False
    'Valida que exista la póliza
    Sql = "select 1 from pp_tmae_pagtergarben "
    Sql = Sql & "where num_poliza='" & Trim(Txt_PenPoliza) & "' "
    'Sql = Sql & "and num_endoso=" & Trim(Lbl_End) & " "
    Sql = Sql & "AND NUM_ORDEN = " & Trim(Lbl_BMNumOrd) & " "
    Set vgRs = vgConexionBD.Execute(Sql)
    If Not vgRs.EOF Then
        vgSw = True
    End If
    vgRs.Close

    If (vgSw) Then
        vlResp = MsgBox(" ¿ Está seguro que desea Eliminar los Datos ?", 4 + 32 + 256, "Proceso de Eliminación de Datos")
        If vlResp <> 6 Then
            Screen.MousePointer = 0
            Exit Sub
        End If
        
        Sql = "delete from pp_tmae_pagtergarben "
        Sql = Sql & "where num_poliza='" & Trim(Txt_PenPoliza) & "' "
        Sql = Sql & "AND NUM_ORDEN = " & Trim(Lbl_BMNumOrd) & " "
''        Sql = Sql & "and num_endoso=" & Trim(Lbl_End) & " "
        vgConexionBD.Execute (Sql)
        
            'CORPTEC
        vlLog_tabla = "PP_TMAE_PAGTERGARBEN"
        vlLog_idtabla = "NUM_POLIZA.NUM_ORDEN"
        vlLog_valtabla = "" & Trim(Txt_PenPoliza) & "." & Trim(Lbl_BMNumOrd) & ""
        vlLog_trans = "DLT"
        Call flLog_Tabla
        
        vgSw = False
        Sql = "select 1 from pp_tmae_pagtergarben "
        Sql = Sql & "where num_poliza='" & Trim(Txt_PenPoliza) & "' "
        Set vgRs = vgConexionBD.Execute(Sql)
        If vgRs.EOF Then
            vgSw = True
        End If
        vgRs.Close
        
        If (vgSw = True) Then
            Sql = "delete from pp_tmae_pagtergar "
            Sql = Sql & "where num_poliza='" & Trim(Txt_PenPoliza) & "' "
            'Sql = Sql & "and num_endoso=" & Trim(Lbl_End) & " "
            vgConexionBD.Execute (Sql)
            
                        'CORPTEC
            vlLog_tabla = "PP_TMAE_PAGTERGAR"
            vlLog_idtabla = "NUM_POLIZA"
            vlLog_valtabla = "" & Trim(Txt_PenPoliza) & ""
            vlLog_trans = "DLT"
            Call flLog_Tabla
        
        End If

        MsgBox "La Eliminacion del Pago Garantizado ha finalizado correctamente.", vbInformation, "Proceso Finalizado"

        Call Cmd_Limpiar_Click
        'Call Cmd_Cancelar_Click
        Call Cmd_BuscarPol_Click
    Else
        MsgBox "La Póliza no tiene registrado ningún Pago Garantizado.", vbCritical, "Proceso Cancelado"
    End If

Exit Sub
Err_Eliminar:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " " & "Numero de Error : " & Err.Number & Err.HelpFile & Err.HelpContext, vbExclamation, "Mensaje de error ..."
        Screen.MousePointer = 0
    End If
End Sub

Function flFormatearDatosEndoso()

    vlNumPoliza = Trim(Txt_PenPoliza)
    vlNumEndoso = Trim(Lbl_End)
    vlFecEndoso = fgBuscaFecServ
    vlFecEndoso = Format(CDate(Trim(vlFecEndoso)), "yyyymmdd")
    vlFecSolEndoso = vlFecEndoso
    vlCodTipEndoso = "S"
    vlCodCauEndoso = "23"
    vlMtoDiferencia = 0
    vlMtoPensionOri = Lbl_MtoPensionRef
    vlMtoPensionCal = Lbl_MtoPensionRef
    vlPrcFactor = Format(1, "#0.000000") 'vlPorcentaje
    vlFecEfectoEnd = Format(CDate(Trim(Lbl_FechaEfecto)), "yyyymmdd")
    vlObsEndoso = ""
    vlFecFinEfecto = Trim(vgTopeFecFin)
    'Estado para PreEndoso
    vlCodEstadoEndoso = "P"
    vlGlsUsuarioCrea = vgUsuario
    vlFecCrea = Format(Date, "yyyymmdd")
    vlHorCrea = Format(Time, "hhmmss")
    vlGlsUsuarioModi = vgUsuario
    vlFecModi = Format(Date, "yyyymmdd")
    vlHorModi = Format(Time, "hhmmss")

End Function

Function flInsertarEndoso()

    vgSql = "INSERT INTO PP_TMAE_ENDOSO "
    vgSql = vgSql & "(num_poliza,num_endoso,fec_solendoso,fec_endoso, "
    vgSql = vgSql & "cod_cauendoso,cod_tipendoso,mto_diferencia, "
    vgSql = vgSql & "mto_pensionori,mto_pensioncal,fec_efecto, "
    vgSql = vgSql & "prc_factor,"
    If (vlObsEndoso <> "") Then
        vgSql = vgSql & "gls_observacion, "
    End If
    vgSql = vgSql & "fec_finefecto,cod_moneda,"
    vgSql = vgSql & "cod_estado,"
    vgSql = vgSql & "cod_usuariocrea,fec_crea,hor_crea "
    vgSql = vgSql & " ) VALUES ( "
    vgSql = vgSql & "'" & vlNumPoliza & "', "
    vgSql = vgSql & " " & vlNumEndoso & ", "
    vgSql = vgSql & "'" & vlFecSolEndoso & "', "
    vgSql = vgSql & "'" & vlFecEndoso & "', "
    vgSql = vgSql & "'" & vlCodCauEndoso & "', "
    vgSql = vgSql & "'" & vlCodTipEndoso & "', "
    vgSql = vgSql & " " & str(vlMtoDiferencia) & ", "
    vgSql = vgSql & " " & str(vlMtoPensionOri) & ", "
    vgSql = vgSql & " " & str(vlMtoPensionCal) & ", "
    vgSql = vgSql & "'" & vlFecEfectoEnd & "', "
    vgSql = vgSql & " " & str(vlPrcFactor) & ", "
    If (vlObsEndoso <> "") Then
        vgSql = vgSql & "'" & vlObsEndoso & "', "
    End If
    vgSql = vgSql & "'" & Trim(vlFecFinEfecto) & "',"
    vgSql = vgSql & "'" & Trim(vlCodMoneda) & "',"
    vgSql = vgSql & "'" & Trim(vlCodEstadoEndoso) & "',"
    vgSql = vgSql & "'" & vlGlsUsuarioCrea & "', "
    vgSql = vgSql & "'" & vlFecCrea & "', "
    vgSql = vgSql & "'" & vlHorCrea & "' ) "
    vgConexionBD.Execute vgSql
    
End Function

Function flFormatearDatosPolizaDef() As Boolean
    
    flFormatearDatosPolizaDef = False
    
    'Obtener los Datos de la Póliza Anterior
    vgSql = "SELECT cod_afp,mto_pension, "
    vgSql = vgSql & "prc_tasactorea,fec_inipagopen "
    vgSql = vgSql & ",cod_cuspp,ind_cob,cod_moneda,mto_valmoneda,"
    vgSql = vgSql & "cod_cobercon,mto_facpenella,prc_facpenella, "
    vgSql = vgSql & "cod_dercre,cod_dergra,prc_tasatir,fec_emision,"
    vgSql = vgSql & "fec_dev,fec_inipencia,fec_pripago,Fec_finperdif,Fec_finpergar "
    vgSql = vgSql & " ,COD_TIPPENSION,COD_ESTADO,COD_TIPREN,COD_MODALIDAD,"
    vgSql = vgSql & " NUM_CARGAS,FEC_VIGENCIA,FEC_TERVIGENCIA,MTO_PRIMA,mto_pension,"
    vgSql = vgSql & " MTO_PENSIONGAR,NUM_MESDIF,NUM_MESGAR,PRC_TASACE,"
    vgSql = vgSql & " PRC_TASAVTA,PRC_TASAINTPERGAR,COD_TIPORIGEN,"
    vgSql = vgSql & " NUM_INDQUIEBRA,FEC_EFECTO,COD_TIPREAJUSTE, MTO_VALREAJUSTETRI, MTO_VALREAJUSTEMEN, FEC_DEVSOL, IND_BENDES, IND_BOLELEC, IND_HEREN, IND_ESTSUN "
    vgSql = vgSql & "FROM PP_TMAE_POLIZA "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "num_poliza = '" & vlNumPoliza & "' AND "
    vgSql = vgSql & "num_endoso = " & vlNumEndoso & " "
    vgSql = vgSql & "ORDER BY num_endoso DESC "
    Set vgRegistro = vgConectarBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
    
        vlCodAFP = (vgRegistro!cod_afp)
        vlFecIniPagoPen = (vgRegistro!Fec_IniPagoPen)
        vlCodCuspp = Trim(vgRegistro!Cod_Cuspp)
        vlIndCobertura = Trim(vgRegistro!Ind_Cob)
        'vlCodMoneda = vgRegistro!Cod_Moneda
        vlValMoneda = vgRegistro!Mto_ValMoneda
        vlCobCoberCon = vgRegistro!Cod_CoberCon
        vlMtoFacPenElla = vgRegistro!Mto_FacPenElla
        vlPrcFacPenElla = vgRegistro!Prc_FacPenElla
        vlCodDerCrecer = vgRegistro!Cod_DerCre
        vlCodDerGratificacion = vgRegistro!Cod_DerGra
        vlFecEmision = vgRegistro!Fec_Emision
        vlFecDevengue = vgRegistro!fec_dev
        vlFecIniPenCia = vgRegistro!Fec_IniPenCia
        vlFecPriPago = vgRegistro!Fec_PriPago
        vlPrcTasaTir = vgRegistro!prc_tasatir
        vlFecTerPagoPenDif = IIf(IsNull(vgRegistro!FEC_FINPERDIF), "", vgRegistro!FEC_FINPERDIF)
        vlFecTerPagoPenGar = IIf(IsNull(vgRegistro!fec_finpergar), "", vgRegistro!fec_finpergar)
        vlCodTipPension = vgRegistro!Cod_TipPension
        vlCodEstado = "9" 'vgRegistro!Cod_Estado
        vlCodTipRen = vgRegistro!Cod_TipRen
        vlCodModalidad = vgRegistro!Cod_Modalidad
        vlNumCargas = vgRegistro!Num_Cargas
        vlFecVigencia = vgRegistro!Fec_Vigencia
        vlFecTerVigencia = vgRegistro!Fec_TerVigencia
        vlMtoPrima = vgRegistro!Mto_Prima
        vlMtoPension = vgRegistro!Mto_Pension
        vlMtoPensionGar = vgRegistro!Mto_PensionGar
        vlNumMesDif = vgRegistro!Num_MesDif
        vlNumMesGar = vgRegistro!Num_MesGar
        vlPrcTasaCe = vgRegistro!Prc_TasaCe
        vlPrcTasaVta = vgRegistro!Prc_TasaVta
        vlPrcTasaCtoRea = IIf(IsNull(vgRegistro!prc_tasactorea), 0, (vgRegistro!prc_tasactorea))
        vlPrcTasaIntPerGar = IIf(IsNull(vgRegistro!Prc_TasaIntPerGar), 0, (vgRegistro!Prc_TasaIntPerGar))
        'RRR 31/10/2019
        vlTtpreaj = vgRegistro!Cod_TipReajuste
        vlMtoajTri = vgRegistro!Mto_ValReajusteTri
        vlMtoajMen = vgRegistro!Mto_ValReajusteMen
        vlFecSol = IIf(IsNull(vgRegistro!FEC_DEVSOL), "", (vgRegistro!FEC_DEVSOL))
        vlBendes = IIf(IsNull(vgRegistro!IND_BENDES), "", (vgRegistro!IND_BENDES))
        vlBolelect = IIf(IsNull(vgRegistro!ind_bolelec), "", (vgRegistro!ind_bolelec))
        vlIndHerencia = "2"
        vlIndEstSub = IIf(IsNull(vgRegistro!IND_ESTSUN), "", (vgRegistro!IND_ESTSUN))
        
        flFormatearDatosPolizaDef = True
    
    End If
    vgRegistro.Close
    
End Function

Function flInsertarPolizaDef()
    
    vgSql = "INSERT INTO PP_TMAE_POLIZA "
    vgSql = vgSql & "(num_poliza,num_endoso,cod_afp,cod_tippension, "
    vgSql = vgSql & "cod_estado,cod_tipren,cod_modalidad,num_cargas, "
    vgSql = vgSql & "fec_vigencia,fec_tervigencia,mto_prima,mto_pension, "
    vgSql = vgSql & "mto_pensiongar, "
    vgSql = vgSql & "num_mesdif,num_mesgar,prc_tasace,prc_tasavta, "
    vgSql = vgSql & "prc_tasactorea,prc_tasaintpergar,fec_inipagopen, "
    vgSql = vgSql & "cod_usuariocrea,fec_crea,hor_crea "
    vgSql = vgSql & ",Cod_Cuspp,Ind_Cob,Cod_Moneda,Mto_ValMoneda,"
    vgSql = vgSql & "Cod_CoberCon,Mto_FacPenElla,Prc_FacPenElla,"
    vgSql = vgSql & "Cod_DerCre,Cod_DerGra,Fec_Emision,fec_dev,"
    vgSql = vgSql & "Fec_IniPenCia,Fec_PriPago,prc_tasatir "
    If (vlFecTerPagoPenGar <> "") Then vgSql = vgSql & ",Fec_finpergar"
    If (vlFecTerPagoPenDif <> "") Then vgSql = vgSql & ",Fec_finperdif"
    vgSql = vgSql & ",Fec_efecto, cod_tipreajuste, mto_valreajustetri, mto_valreajustemen"
    If (vlFecSol <> "") Then vgSql = vgSql & ",fec_devsol"
    If (vlBendes <> "") Then vgSql = vgSql & ",ind_bendes"
    If (vlBolelect <> "") Then vgSql = vgSql & ",ind_bolelec"
    If (vlIndHerencia <> "") Then vgSql = vgSql & ",ind_heren"
    If (vlIndEstSub <> "") Then vgSql = vgSql & ",ind_estsun"
    vgSql = vgSql & " ) VALUES ( "
    vgSql = vgSql & "'" & vlNumPoliza & "', "
    vgSql = vgSql & " " & vlNumEndosoNew & ", "
    vgSql = vgSql & "'" & vlCodAFP & "', "
    vgSql = vgSql & "'" & vlCodTipPension & "', "
    vgSql = vgSql & "'" & vlCodEstado & "', "
    vgSql = vgSql & "'" & vlCodTipRen & "', "
    vgSql = vgSql & "'" & vlCodModalidad & "', "
    vgSql = vgSql & " " & vlNumCargas & ", "
    vgSql = vgSql & "'" & vlFecVigencia & "', "
    vgSql = vgSql & "'" & vlFecTerVigencia & "', "
    vgSql = vgSql & " " & str(vlMtoPrima) & ", "
    vgSql = vgSql & " " & str(vlMtoPension) & ", "
    vgSql = vgSql & " " & str(vlMtoPensionGar) & ", "
    vgSql = vgSql & " " & vlNumMesDif & ", "
    vgSql = vgSql & " " & vlNumMesGar & ", "
    vgSql = vgSql & " " & str(vlPrcTasaCe) & ", "
    vgSql = vgSql & " " & str(vlPrcTasaVta) & ", "
    vgSql = vgSql & " " & str(vlPrcTasaCtoRea) & ", "
    vgSql = vgSql & " " & str(vlPrcTasaIntPerGar) & ", "
    vgSql = vgSql & "'" & Trim(vlFecIniPagoPen) & "', "
    vgSql = vgSql & "'" & Trim(vlGlsUsuarioCrea) & "', "
    vgSql = vgSql & "'" & Trim(vlFecCrea) & "', "
    vgSql = vgSql & "'" & Trim(vlHorCrea) & "'"
    vgSql = vgSql & ",'" & Trim(vlCodCuspp) & "',"
    vgSql = vgSql & "'" & Trim(vlIndCobertura) & "', "
    vgSql = vgSql & "'" & Trim(vlCodMoneda) & "', "
    vgSql = vgSql & " " & str(vlValMoneda) & ", "
    vgSql = vgSql & "'" & Trim(vlCobCoberCon) & "', "
    vgSql = vgSql & " " & str(vlMtoFacPenElla) & ", "
    vgSql = vgSql & " " & str(vlPrcFacPenElla) & ", "
    vgSql = vgSql & "'" & Trim(vlCodDerCrecer) & "', "
    vgSql = vgSql & "'" & Trim(vlCodDerGratificacion) & "', "
    vgSql = vgSql & "'" & Trim(vlFecEmision) & "', "
    vgSql = vgSql & "'" & Trim(vlFecDevengue) & "', "
    vgSql = vgSql & "'" & Trim(vlFecIniPenCia) & "', "
    vgSql = vgSql & "'" & Trim(vlFecPriPago) & "', "
    vgSql = vgSql & " " & str(vlPrcTasaTir) & " "
    If (vlFecTerPagoPenGar <> "") Then vgSql = vgSql & ",'" & Trim(vlFecTerPagoPenGar) & "' "
    If (vlFecTerPagoPenDif <> "") Then vgSql = vgSql & ",'" & Trim(vlFecTerPagoPenDif) & "' "
    vgSql = vgSql & ",'" & Trim(vlFecEfectoEnd) & "' "
    vgSql = vgSql & "," & str(vlTtpreaj) & " "
    vgSql = vgSql & "," & str(vlMtoajTri) & " "
    vgSql = vgSql & "," & str(vlMtoajMen) & " "
    'RRR 31/10/2019
    If (vlFecSol <> "") Then vgSql = vgSql & ",'" & Trim(vlFecSol) & "' "
    If (vlBendes <> "") Then vgSql = vgSql & ",'" & Trim(vlBendes) & "' "
    If (vlBolelect <> "") Then vgSql = vgSql & ",'" & Trim(vlBolelect) & "' "
    If (vlIndHerencia <> "") Then vgSql = vgSql & ",'" & Trim(vlIndHerencia) & "' "
    If (vlIndEstSub <> "") Then vgSql = vgSql & ",'" & Trim(vlIndEstSub) & "' "
    vgSql = vgSql & ") "
    
    vgConexionBD.Execute vgSql
    
End Function

Function flFormatearDatosBeneficiarioDef(iFila As Integer) As Boolean

    flFormatearDatosBeneficiarioDef = False
    
    Msf_BMGrilla.row = iFila
    
    vlNumOrd = Msf_BMGrilla.TextMatrix(iFila, 0)
    vlCodPar = Msf_BMGrilla.TextMatrix(iFila, 1)
    vlTipoIdent = fgObtenerCodigo_TextoCompuesto(Msf_BMGrilla.TextMatrix(iFila, 2))
    vlNumIdent = Msf_BMGrilla.TextMatrix(iFila, 3)
    vlNom = Msf_BMGrilla.TextMatrix(iFila, 4)
    vlNomSeg = Msf_BMGrilla.TextMatrix(iFila, 5)
    vlApPat = Msf_BMGrilla.TextMatrix(iFila, 6)
    vlApMat = Msf_BMGrilla.TextMatrix(iFila, 7)
    vlEstPension = Msf_BMGrilla.TextMatrix(iFila, 8)
    vlEstPago = Msf_BMGrilla.TextMatrix(iFila, 9)
    vlPrcPension = Msf_BMGrilla.TextMatrix(iFila, 10)
    vlMtoPension = Msf_BMGrilla.TextMatrix(iFila, 11)
    vlMesesPag = Msf_BMGrilla.TextMatrix(iFila, 12)
    vlMesesNoDev = Msf_BMGrilla.TextMatrix(iFila, 13)
    vlPenNoPer = Msf_BMGrilla.TextMatrix(iFila, 14)
    vlFecPago = Format(Msf_BMGrilla.TextMatrix(iFila, 15), "yyyymmdd")
    vlSexo = Msf_BMGrilla.TextMatrix(iFila, 16)
    vlFecNac = Msf_BMGrilla.TextMatrix(iFila, 17)
    vlPrcLeg = Msf_BMGrilla.TextMatrix(iFila, 18)
    vlPrcGar = Msf_BMGrilla.TextMatrix(iFila, 19)
    vlMtoPenGar = Msf_BMGrilla.TextMatrix(iFila, 20)
    vlGruFam = Msf_BMGrilla.TextMatrix(iFila, 21)
    vlDirec = Msf_BMGrilla.TextMatrix(iFila, 22)
    vlCodDir = Msf_BMGrilla.TextMatrix(iFila, 23)
    vlFono = Msf_BMGrilla.TextMatrix(iFila, 24)
    vlEmail = Msf_BMGrilla.TextMatrix(iFila, 25)
    
    vlViaPago = Msf_BMGrilla.TextMatrix(iFila, 26)
    vlSucursal = Msf_BMGrilla.TextMatrix(iFila, 27)
    vlTipcta = Msf_BMGrilla.TextMatrix(iFila, 28)
    vlBanco = Msf_BMGrilla.TextMatrix(iFila, 29)
    vlNumCta = Msf_BMGrilla.TextMatrix(iFila, 30)
    vlFecIni = Msf_BMGrilla.TextMatrix(iFila, 31)
    vlFecFin = Msf_BMGrilla.TextMatrix(iFila, 32)
    vlTipCambio = Msf_BMGrilla.TextMatrix(iFila, 33)
    vlFecRecep = Msf_BMGrilla.TextMatrix(iFila, 34)
    vlTasaInt = Msf_BMGrilla.TextMatrix(iFila, 35)
    vlDerpen = Msf_BMGrilla.TextMatrix(iFila, 36)
    
    'Consulta para obtener los Datos faltantes del Beneficiario
    'Obtener los datos de la Tabla de Beneficiarios
    vgSql = "SELECT "
    vgSql = vgSql & "Cod_InsSalud,COD_MODSALUD,MTO_PLANSALUD,"
    vgSql = vgSql & "FEC_INGRESO,COD_SITINV,COD_DERCRE,"
    vgSql = vgSql & "COD_CAUINV,FEC_NACHM,"
    vgSql = vgSql & "FEC_INVBEN,COD_MOTREQPEN,FEC_FALLBEN,"
    vgSql = vgSql & "FEC_MATRIMONIO,COD_CAUSUSBEN,FEC_SUSBEN,"
    vgSql = vgSql & "FEC_INIPAGOPEN,FEC_TERPAGOPENGAR, GLS_TELBEN2,COD_TIPCTA, COD_MONBCO, NUM_CTABCO, IND_BOLELEC, NUM_CUENTA_CCI, CONS_TRAINFO, CONS_DATCOMER "
    vgSql = vgSql & "FROM PP_TMAE_BEN WHERE "
    vgSql = vgSql & "num_poliza = '" & vlNumPoliza & "' "
    vgSql = vgSql & "and num_endoso = " & vlNumEndoso & " "
    vgSql = vgSql & "and num_orden = " & vlNumOrd & " "
    Set vgRs4 = vgConectarBD.Execute(vgSql)
    If Not vgRs4.EOF Then
        vlCodInsSalud = vgRs4!Cod_InsSalud
        vlCodModSalud = vgRs4!Cod_ModSalud
        vlMtoPlanSalud = vgRs4!Mto_PlanSalud
        vlFecIngreso = vgRs4!Fec_Ingreso
        vlCodSitInv = vgRs4!Cod_SitInv
        vlCodDerCre = vgRs4!Cod_DerCre
        vlCodCauInv = vgRs4!Cod_CauInv
        If (Not IsNull(vgRs4!Fec_NacHM)) Then vlFecNacHM = vgRs4!Fec_NacHM Else vlFecNacHM = ""
        If (Not IsNull(vgRs4!Fec_InvBen)) Then vlFecInvBen = vgRs4!Fec_InvBen Else vlFecInvBen = ""
        vlCodMotReqPen = vgRs4!Cod_MotReqPen
        If (Not IsNull(vgRs4!Fec_FallBen)) Then vlFecFallBen = vgRs4!Fec_FallBen Else vlFecFallBen = ""
        If (Not IsNull(vgRs4!Fec_Matrimonio)) Then vlFecMatrimonio = vgRs4!Fec_Matrimonio Else vlFecMatrimonio = ""
        vlCodCauSusBen = vgRs4!Cod_CauSusBen
        If (vgRs4!Fec_SusBen <> "") Then vlFecSusBen = vgRs4!Fec_SusBen Else vlFecSusBen = ""
        vlFecIniPagoPen = vgRs4!Fec_IniPagoPen
        If (vgRs4!Fec_TerPagoPenGar <> "") Then vlFecTerPagoPenGar = vgRs4!Fec_TerPagoPenGar Else vlFecTerPagoPenGar = ""
        'RRR 31/10/2019
        If (Not IsNull(vgRs4!Gls_Telben2)) Then vlTelben2 = vgRs4!Gls_Telben2 Else vlTelben2 = ""
        If (Not IsNull(vgRs4!cod_tipcta)) Then vlTipcta = vgRs4!cod_tipcta Else vlTipcta = ""
        If (Not IsNull(vgRs4!cod_monbco)) Then vlMonbco = vgRs4!cod_monbco Else vlMonbco = ""
        If (Not IsNull(vgRs4!num_ctabco)) Then vlCtabco = vgRs4!num_ctabco Else vlCtabco = ""
        If (Not IsNull(vgRs4!ind_bolelec)) Then vlBolelec = vgRs4!ind_bolelec Else vlBolelec = ""
        If (Not IsNull(vgRs4!NUM_CUENTA_CCI)) Then vlNumCCI = vgRs4!NUM_CUENTA_CCI Else vlNumCCI = ""
        If (Not IsNull(vgRs4!CONS_TRAINFO)) Then vlTrainfo = vgRs4!CONS_TRAINFO Else vlTrainfo = ""
        If (Not IsNull(vgRs4!CONS_DATCOMER)) Then vlDatcomer = vgRs4!CONS_DATCOMER Else vlDatcomer = ""
    
    End If
    
    If (Trim(Msf_BMGrilla.TextMatrix(vgI, 9)) = clEstCalculado) Then
        If (vlCodPar = "77") Then
            vlDerpen = "10"
            vlEstPension = "10"
            vlFecIniPagoPen = Format(DateSerial(Mid(vlFecEfectoEnd, 1, 4), Mid(vlFecEfectoEnd, 5, 2), 1 - 1), "yyyymmdd")
            'vlFecTerPagoPenGar = vlFecIniPagoPen
        Else
            'vlFecIniPagoPen = Format(DateSerial(Mid(vgRs4!Fec_TerPagoPenGar, 1, 4), Mid(vgRs4!Fec_TerPagoPenGar, 5, 2), Mid(vgRs4!Fec_TerPagoPenGar, 7, 2) + 1), "yyyymmdd")
            'vlFecTerPagoPenGar = vgRs4!Fec_TerPagoPenGar
        End If
    End If
    vgRs4.Close
    
    flFormatearDatosBeneficiarioDef = True
    
End Function

Function flInsertarBeneficiarioDef()

    vgSql = "INSERT INTO PP_TMAE_BEN "
    vgSql = vgSql & "(num_poliza,num_endoso,num_orden,fec_ingreso,Cod_TipoIdenben,Num_Idenben, "
    vgSql = vgSql & "gls_nomben,gls_nomsegben,gls_patben,gls_matben,gls_dirben,cod_direccion,"
    If (Trim(vlFono) <> "") Then vgSql = vgSql & "gls_fonoben,"
    If (Trim(vlEmail) <> "") Then vgSql = vgSql & "gls_correoben,"
    vgSql = vgSql & "cod_grufam, "
    vgSql = vgSql & "cod_par,cod_sexo,cod_sitinv,cod_dercre,cod_derpen, "
    vgSql = vgSql & "cod_cauinv,fec_nacben,"
    If (Trim(vlFecNacHM) <> "") Then vgSql = vgSql & "fec_nachm,"
    If (Trim(vlFecInvBen) <> "") Then vgSql = vgSql & "fec_invben, "
    vgSql = vgSql & "cod_motreqpen,mto_pension,mto_pensiongar,prc_pension, "
    vgSql = vgSql & "cod_inssalud,cod_modsalud,mto_plansalud, "
    vgSql = vgSql & "cod_estpension,"
    'vgSql = vgSql & "cod_cajacompen,"
    vgSql = vgSql & "cod_viapago,cod_banco,"
    vgSql = vgSql & "cod_tipcuenta,"
    If (Trim(vlNumCta) <> "") Then vgSql = vgSql & "num_cuenta,"
    vgSql = vgSql & "cod_sucursal,"
    If (Trim(vlFecFallBen) <> "") Then vgSql = vgSql & "fec_fallben,"
    If (Trim(vlFecMatrimonio) <> "") Then vgSql = vgSql & "fec_matrimonio,"
    vgSql = vgSql & "cod_caususben,"
    If (Trim(vlFecSusBen) <> "") Then vgSql = vgSql & "fec_susben, "
    vgSql = vgSql & "fec_inipagopen,"
    If (Trim(vlFecTerPagoPenGar) <> "") Then vgSql = vgSql & "fec_terpagopengar, "
    vgSql = vgSql & "cod_usuariocrea,fec_crea,hor_crea, "
    vgSql = vgSql & "prc_pensiongar, prc_pensionleg, "
    'RRR 31/10/2019
    If (Trim(vlTelben2) <> "") Then vgSql = vgSql & "Gls_Telben2, "
    If (Trim(vlTipctaB) <> "") Then vgSql = vgSql & "cod_tipcta, "
    If (Trim(vlMonbco) <> "") Then vgSql = vgSql & "cod_monbco, "
    If (Trim(vlCtabco) <> "") Then vgSql = vgSql & "cod_ctabco, "
    If (Trim(vlBolelec) <> "") Then vgSql = vgSql & "ind_bolelec, "
    If (Trim(vlNumCCI) <> "") Then vgSql = vgSql & "NUM_CUENTA_CCI, "
    If (Trim(vlTrainfo) <> "") Then vgSql = vgSql & "CONS_TRAINFO, "
    If (Trim(vlDatcomer) <> "") Then vgSql = vgSql & "CONS_DATCOMER "
    
    vgSql = vgSql & " ) VALUES ( "
    vgSql = vgSql & "'" & vlNumPoliza & "', "
    vgSql = vgSql & " " & vlNumEndosoNew & ", "
    vgSql = vgSql & " " & vlNumOrd & ", "
    vgSql = vgSql & "'" & Trim(vlFecIngreso) & "', "
    vgSql = vgSql & " " & (vlTipoIdent) & ", "
    vgSql = vgSql & "'" & vlNumIdent & "', "
    vgSql = vgSql & "'" & vlNom & "', "
    If (vlNomSeg <> "") Then vgSql = vgSql & "'" & vlNomSeg & "', " Else vgSql = vgSql & "NULL,"
    vgSql = vgSql & "'" & vlApPat & "', "
    If (vlApMat <> "") Then vgSql = vgSql & "'" & vlApMat & "', " Else vgSql = vgSql & "NULL,"
    vgSql = vgSql & "'" & Trim(vlDirec) & "', "
    vgSql = vgSql & " " & str(vlCodDir) & ", "
    If (Trim(vlFono) <> "") Then vgSql = vgSql & "'" & Trim(vlFono) & "', "
    If (Trim(vlEmail) <> "") Then vgSql = vgSql & "'" & Trim(vlEmail) & "', "
    vgSql = vgSql & "'" & vlGruFam & "', "
    vgSql = vgSql & "'" & vlCodPar & "', "
    vgSql = vgSql & "'" & vlSexo & "', "
    vgSql = vgSql & "'" & vlCodSitInv & "', "
    vgSql = vgSql & "'" & vlCodDerCre & "', "
    vgSql = vgSql & "'" & vlDerpen & "', "
    vgSql = vgSql & "'" & vlCodCauInv & "', "
    vgSql = vgSql & "'" & vlFecNac & "', "
    If (Trim(vlFecNacHM) <> "") Then vgSql = vgSql & "'" & vlFecNacHM & "', "
    If (Trim(vlFecInvBen) <> "") Then vgSql = vgSql & "'" & vlFecInvBen & "', "
    vgSql = vgSql & "'" & vlCodMotReqPen & "', "
    vgSql = vgSql & " " & str(vlMtoPension) & ", "
    vgSql = vgSql & " " & str(vlMtoPenGar) & ", "
    vgSql = vgSql & " " & str(vlPrcPension) & ", "
    vgSql = vgSql & "'" & Trim(vlCodInsSalud) & "', "
    vgSql = vgSql & "'" & Trim(vlCodModSalud) & "', "
    vgSql = vgSql & " " & str(vlMtoPlanSalud) & ", "
    vgSql = vgSql & "'" & vlEstPension & "', "
    'vgSql = vgSql & "'" & Trim(vlCodCajaCompen) & "', "
    vgSql = vgSql & "'" & Trim(vlViaPago) & "', "
    vgSql = vgSql & "'" & Trim(vlBanco) & "', "
    vgSql = vgSql & "'" & Trim(vlTipcta) & "', "
    If (Trim(vlNumCta) <> "") Then vgSql = vgSql & "'" & Trim(vlNumCta) & "', "
    vgSql = vgSql & "'" & Trim(vlSucursal) & "', "
    If (Trim(vlFecFallBen) <> "") Then vgSql = vgSql & "'" & vlFecFallBen & "', "
    If (Trim(vlFecMatrimonio) <> "") Then vgSql = vgSql & "'" & Trim(vlFecMatrimonio) & "', "
    vgSql = vgSql & "'" & vlCodCauSusBen & "', "
    If (Trim(vlFecSusBen) <> "") Then vgSql = vgSql & "'" & vlFecSusBen & "', "
    If (vlCodPar = "77") Then
        vgSql = vgSql & "'" & vlFecIniPagoPen & "', "
        If (Trim(vlFecTerPagoPenGar) <> "") Then vgSql = vgSql & "'" & vlFecTerPagoPenGar & "', "
    Else
'        vgSql = vgSql & "'" & Format(DateSerial(Mid(vlFecTerPagoPenGar, 1, 4), Mid(vlFecTerPagoPenGar, 5, 2), Mid(vlFecTerPagoPenGar, 7, 2) + 1), "yyyymmdd") & "', "
        vgSql = vgSql & "'" & vlFecIniPagoPen & "', "
        If (Trim(vlFecTerPagoPenGar) <> "") Then vgSql = vgSql & "'" & vlFecTerPagoPenGar & "', "
    End If
    vgSql = vgSql & "'" & Trim(vlGlsUsuarioCrea) & "', "
    vgSql = vgSql & "'" & Trim(vlFecCrea) & "', "
    vgSql = vgSql & "'" & Trim(vlHorCrea) & "'"
    vgSql = vgSql & "," & str(vlPrcGar) & ","
    vgSql = vgSql & " " & str(vlPrcLeg) & ", "
    'RRR 31/10/2019
    If (Trim(vlTelben2) <> "") Then vgSql = vgSql & "'" & vlTelben2 & "', "
    If (Trim(vlTipctaB) <> "") Then vgSql = vgSql & "'" & vlTipctaB & "', "
    If (Trim(vlMonbco) <> "") Then vgSql = vgSql & "'" & vlMonbco & "', "
    If (Trim(vlCtabco) <> "") Then vgSql = vgSql & "'" & str(vlCtabco) & "', "
    If (Trim(vlBolelec) <> "") Then vgSql = vgSql & "'" & vlBolelec & "', "
    If (Trim(vlNumCCI) <> "") Then vgSql = vgSql & "'" & vlNumCCI & "', "
    If (Trim(vlTrainfo) <> "") Then vgSql = vgSql & "'" & vlTrainfo & "', "
    If (Trim(vlDatcomer) <> "") Then vgSql = vgSql & "'" & vlDatcomer & "' "
    vgSql = vgSql & ") "
    vgConexionBD.Execute vgSql

End Function
'CORPTEC
'FUNCION PARA REGISTRAR EL INICIO DEL LOG DE PROCESO :CARGA DE ARCHIVO 11-08-2017

Function flLog_Tabla() As Boolean
  Dim com As ADODB.Command
    Dim sistema, modulo, opcion, origen, tipo As String
    sistema = "SEACSA"
    modulo = "PENSIONES"
    opcion = "PAGO TERCEROS-PER.GAR."
    origen = "A"
    tipo = "T"
    
    Set com = New ADODB.Command
    
    vgConexionBD.BeginTrans
    com.ActiveConnection = vgConexionBD
    com.CommandText = "SP_LOG_TABLAS"
    com.CommandType = adCmdStoredProc
    
    com.Parameters.Append com.CreateParameter("TRANS", adChar, adParamInput, 3, vlLog_trans)
    com.Parameters.Append com.CreateParameter("TABLA", adVarChar, adParamInput, 50, vlLog_tabla)
    com.Parameters.Append com.CreateParameter("IDTABLA", adVarChar, adParamInput, 50, vlLog_idtabla)
    com.Parameters.Append com.CreateParameter("VALTABLA", adVarChar, adParamInput, 50, vlLog_valtabla)
    com.Parameters.Append com.CreateParameter("USUARIO", adVarChar, adParamInput, 10, vgLogin)
    com.Parameters.Append com.CreateParameter("SISTEMA", adVarChar, adParamInput, 50, sistema)
    com.Parameters.Append com.CreateParameter("MODULO", adVarChar, adParamInput, 50, modulo)
    com.Parameters.Append com.CreateParameter("OPCION", adVarChar, adParamInput, 50, opcion)
    com.Parameters.Append com.CreateParameter("ORIGEN", adChar, adParamInput, 1, origen)
    com.Parameters.Append com.CreateParameter("TIPO", adChar, adParamInput, 1, tipo)
    com.Parameters.Append com.CreateParameter("IDLOG", adDouble, adParamInput, 2, 0)
    com.Parameters.Append com.CreateParameter("Retorno", adDouble, adParamReturnValue)
    com.Execute
    vgConexionBD.CommitTrans
    num_log_Tabla = com("Retorno")
End Function






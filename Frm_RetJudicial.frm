VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Frm_RetJudicial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenedor de Retenciones Judiciales."
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   9000
   Begin VB.Frame Fra_Poliza 
      Caption         =   "c"
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
      TabIndex        =   40
      Top             =   0
      Width           =   8805
      Begin VB.ComboBox Cmb_PenNumIdent 
         BackColor       =   &H00E0FFFF&
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   1875
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
         Left            =   8040
         Picture         =   "Frm_RetJudicial.frx":0000
         TabIndex        =   4
         ToolTipText     =   "Buscar"
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton Cmd_BuscarPol 
         Height          =   375
         Left            =   8040
         Picture         =   "Frm_RetJudicial.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Buscar Póliza"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Txt_PenNumIdent 
         Height          =   285
         Left            =   4800
         MaxLength       =   16
         TabIndex        =   2
         Top             =   360
         Width           =   1875
      End
      Begin VB.TextBox Txt_PenPoliza 
         Height          =   285
         Left            =   840
         MaxLength       =   10
         TabIndex        =   0
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label Lbl_PenNombre 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   255
         Left            =   840
         TabIndex        =   42
         Top             =   720
         Width           =   7095
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   47
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   43
         Left            =   120
         TabIndex        =   46
         Top             =   0
         Width           =   1725
      End
      Begin VB.Label Lbl_End 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   7440
         TabIndex        =   45
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "N° End"
         Height          =   195
         Index           =   42
         Left            =   6840
         TabIndex        =   44
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Ident."
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   43
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Póliza"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   41
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Fra_Botones 
      Height          =   975
      Left            =   120
      TabIndex        =   39
      Top             =   8400
      Width           =   8775
      Begin VB.CommandButton Cmd_Cancelar2 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   5550
         Picture         =   "Frm_RetJudicial.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Eliminar 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   2300
         Picture         =   "Frm_RetJudicial.frx":07DE
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Eliminar Retención"
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   3380
         Picture         =   "Frm_RetJudicial.frx":0B20
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   6650
         Picture         =   "Frm_RetJudicial.frx":11DA
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   4450
         Picture         =   "Frm_RetJudicial.frx":12D4
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Grabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   1220
         Picture         =   "Frm_RetJudicial.frx":198E
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   200
         Width           =   720
      End
      Begin Crystal.CrystalReport Rpt_General 
         Left            =   8280
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   7155
      Left            =   120
      TabIndex        =   48
      Top             =   1200
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   12621
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Retención Judicial"
      TabPicture(0)   =   "Frm_RetJudicial.frx":2048
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Fra_Suspension"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Fra_Receptor"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Fra_Pago"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Fra_AntecedentesRet"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "&Historia de Retención Judicial"
      TabPicture(1)   =   "Frm_RetJudicial.frx":2064
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Msf_GrillaRetJud"
      Tab(1).ControlCount=   1
      Begin VB.Frame Fra_AntecedentesRet 
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
         Height          =   2235
         Left            =   120
         TabIndex        =   72
         Top             =   360
         Width           =   8535
         Begin VB.TextBox txtPensionAct 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            MaxLength       =   14
            TabIndex        =   86
            Top             =   1440
            Width           =   1140
         End
         Begin VB.ComboBox Cmb_ModRetencion 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1080
            Width           =   2985
         End
         Begin VB.ComboBox Cmb_TipoRetencion 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   360
            Width           =   2985
         End
         Begin VB.TextBox Txt_Juzgado 
            Height          =   285
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   13
            Top             =   1800
            Width           =   6300
         End
         Begin VB.TextBox Txt_MontoRetenido 
            Height          =   285
            Left            =   6720
            MaxLength       =   14
            TabIndex        =   9
            Text            =   "0"
            Top             =   1440
            Width           =   1275
         End
         Begin VB.TextBox Txt_FechaRecepcion 
            Height          =   285
            Left            =   6720
            MaxLength       =   10
            TabIndex        =   11
            Top             =   720
            Width           =   1260
         End
         Begin VB.TextBox Txt_FechaIniVig 
            Height          =   285
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   5
            Top             =   720
            Width           =   1260
         End
         Begin VB.TextBox Txt_FechaTerVig 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   3360
            MaxLength       =   10
            TabIndex        =   6
            Top             =   720
            Width           =   1260
         End
         Begin VB.TextBox Txt_MtoMaxRet 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Left            =   6720
            MaxLength       =   14
            TabIndex        =   12
            Text            =   "0"
            Top             =   1080
            Width           =   1275
         End
         Begin VB.Label lblMontoRetAct 
            Caption         =   "Label1"
            Height          =   255
            Left            =   3720
            TabIndex        =   87
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            Caption         =   "Pension Actual"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   85
            Top             =   1440
            Width           =   1065
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            Caption         =   "Entidad Judicial"
            Height          =   195
            Index           =   30
            Left            =   120
            TabIndex        =   83
            Top             =   1800
            Width           =   1110
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            Caption         =   "Monto Retenido"
            Height          =   195
            Index           =   31
            Left            =   5160
            TabIndex        =   82
            Top             =   1440
            Width           =   1140
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Retención"
            Height          =   195
            Index           =   29
            Left            =   120
            TabIndex        =   81
            Top             =   360
            Width           =   1320
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            Caption         =   "Modalidad Retención"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   80
            Top             =   1080
            Width           =   1515
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Recepción"
            Height          =   195
            Index           =   4
            Left            =   5160
            TabIndex        =   79
            Top             =   720
            Width           =   1275
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            Caption         =   "Periodo de Vigencia"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   78
            Top             =   720
            Width           =   1425
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
            Index           =   21
            Left            =   3000
            TabIndex        =   77
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label9 
            Caption         =   "  Antecedentes Retención  "
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
            TabIndex        =   76
            Top             =   0
            Width           =   2415
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            Caption         =   "Periodo de Efecto"
            Height          =   195
            Index           =   37
            Left            =   5160
            TabIndex        =   75
            Top             =   360
            Width           =   1275
         End
         Begin VB.Label Lbl_FechaEfecto 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6720
            TabIndex        =   10
            Top             =   360
            Width           =   1260
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            Caption         =   "Máxima Retención"
            Height          =   195
            Index           =   20
            Left            =   5160
            TabIndex        =   74
            Top             =   1080
            Width           =   1320
         End
         Begin VB.Label Lbl_Moneda 
            Caption         =   "NS"
            Height          =   255
            Left            =   2880
            TabIndex        =   73
            Top             =   1440
            Width           =   255
         End
      End
      Begin VB.Frame Fra_Pago 
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
         Height          =   1185
         Left            =   120
         TabIndex        =   65
         Top             =   5040
         Width           =   8535
         Begin VB.ComboBox Cmb_TipoCta 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   5280
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   180
            Width           =   2940
         End
         Begin VB.ComboBox Cmb_ViaPago 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   315
            Width           =   2955
         End
         Begin VB.ComboBox Cmb_Banco 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   5280
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   510
            Width           =   2940
         End
         Begin VB.TextBox Txt_NumCuenta 
            Height          =   285
            Left            =   5280
            MaxLength       =   15
            TabIndex        =   31
            Top             =   840
            Width           =   2910
         End
         Begin VB.ComboBox Cmb_Sucursal 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   660
            Width           =   2955
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Tipo Cta."
            Height          =   255
            Index           =   17
            Left            =   4320
            TabIndex        =   71
            Top             =   200
            Width           =   945
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Vía Pago"
            Height          =   255
            Index           =   15
            Left            =   165
            TabIndex        =   70
            Top             =   360
            Width           =   930
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Banco"
            Height          =   255
            Index           =   18
            Left            =   4320
            TabIndex        =   69
            Top             =   510
            Width           =   945
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "N°Cuenta"
            Height          =   255
            Index           =   19
            Left            =   4320
            TabIndex        =   68
            Top             =   840
            Width           =   915
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Sucursal"
            Height          =   255
            Index           =   16
            Left            =   165
            TabIndex        =   67
            Top             =   680
            Width           =   900
         End
         Begin VB.Label Label11 
            Caption         =   "  Forma de Pago de Pensión  "
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
            TabIndex        =   66
            Top             =   0
            Width           =   2535
         End
      End
      Begin VB.Frame Fra_Receptor 
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
         Left            =   120
         TabIndex        =   52
         Top             =   2640
         Width           =   8535
         Begin VB.TextBox Txt_CorreoRec 
            Height          =   285
            Left            =   4080
            MaxLength       =   20
            TabIndex        =   26
            Top             =   1900
            Width           =   4330
         End
         Begin VB.TextBox Txt_MaternoRec 
            Height          =   285
            Left            =   5520
            MaxLength       =   20
            TabIndex        =   19
            Top             =   930
            Width           =   2895
         End
         Begin VB.TextBox Txt_PaternoRec 
            Height          =   285
            Left            =   1200
            MaxLength       =   20
            TabIndex        =   18
            Top             =   930
            Width           =   2775
         End
         Begin VB.TextBox Txt_DireccionRec 
            Height          =   285
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   20
            Top             =   1250
            Width           =   7215
         End
         Begin VB.TextBox Txt_NombreRec 
            Height          =   285
            Left            =   1200
            MaxLength       =   25
            TabIndex        =   16
            Top             =   600
            Width           =   2775
         End
         Begin VB.TextBox Txt_TelefonoRec 
            Height          =   285
            Left            =   1200
            MaxLength       =   20
            TabIndex        =   25
            Top             =   1900
            Width           =   2025
         End
         Begin VB.TextBox Txt_NumIdent 
            Height          =   285
            Left            =   3120
            MaxLength       =   10
            TabIndex        =   15
            Top             =   240
            Width           =   1740
         End
         Begin VB.TextBox Txt_NombreSegRec 
            Height          =   285
            Left            =   5520
            MaxLength       =   25
            TabIndex        =   17
            Top             =   600
            Width           =   2895
         End
         Begin VB.ComboBox Cmb_NumIdent 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   240
            Width           =   1875
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
            Left            =   8160
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Efectuar Busqueda de Dirección"
            Top             =   1560
            Width           =   300
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Ap. Materno"
            Height          =   255
            Index           =   10
            Left            =   4200
            TabIndex        =   64
            Top             =   930
            Width           =   915
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Ap. Paterno"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   63
            Top             =   930
            Width           =   915
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Email"
            Height          =   255
            Index           =   13
            Left            =   3480
            TabIndex        =   62
            Top             =   1900
            Width           =   615
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Ubicación"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   61
            Top             =   1560
            Width           =   855
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Dirección"
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   60
            Top             =   1250
            Width           =   735
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Primer Nombre"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   59
            Top             =   615
            Width           =   1035
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Teléfono"
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   58
            Top             =   1900
            Width           =   795
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "N° Ident."
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   57
            Top             =   300
            Width           =   810
         End
         Begin VB.Label Label10 
            Caption         =   "  Antecedenes Personales del Receptor  "
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
            TabIndex        =   56
            Top             =   0
            Width           =   3615
         End
         Begin VB.Label Lbl_NumRetencion 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6840
            TabIndex        =   55
            Top             =   240
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Num. Retención"
            Height          =   255
            Index           =   32
            Left            =   5520
            TabIndex        =   54
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Segundo Nombre"
            Height          =   255
            Index           =   33
            Left            =   4200
            TabIndex        =   53
            Top             =   600
            Width           =   1275
         End
         Begin VB.Label Lbl_Distrito 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   5880
            TabIndex        =   23
            Top             =   1560
            Width           =   2220
         End
         Begin VB.Label Lbl_Provincia 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   3600
            TabIndex        =   22
            Top             =   1560
            Width           =   2205
         End
         Begin VB.Label Lbl_Departamento 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   21
            Top             =   1560
            Width           =   2325
         End
      End
      Begin VB.Frame Fra_Suspension 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   705
         Left            =   120
         TabIndex        =   49
         Top             =   6240
         Width           =   8535
         Begin VB.CheckBox chkViegencia 
            Caption         =   "Vigente"
            Height          =   255
            Left            =   7200
            TabIndex        =   88
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox Txt_FecSuspension 
            Height          =   285
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   32
            Top             =   300
            Width           =   1260
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Fecha de Recepción de la Carta"
            Height          =   255
            Index           =   35
            Left            =   240
            TabIndex        =   51
            Top             =   300
            Width           =   2415
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "  Termino de la Retención  "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   36
            Left            =   120
            TabIndex        =   50
            Top             =   0
            Width           =   2415
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Msf_GrillaRetJud 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   84
         Top             =   480
         Width           =   8505
         _ExtentX        =   15002
         _ExtentY        =   6588
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   14745599
         AllowBigSelection=   0   'False
         FocusRect       =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "Frm_RetJudicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vlRegistroBen As ADODB.Recordset

Dim vlSw As Boolean
Dim vlPos As Integer
Dim vlPosAux As Integer
Dim vlCodDir As Integer
Dim vlNombreRegion As String
Dim vlNombreProvincia As String
Dim vlNombreComuna As String
Dim vlNumero As Integer
Dim vlOpcion As String
Dim vlRutAux As String
Dim vlNumEndoso As Integer
Dim vlNumOrden As Integer
Dim vlcont As Integer
Dim vlNumPoliza As String
Dim vlRut As String
Dim vlDigito As String
Dim vlTipoRet As String
Dim vlModRet As String
Dim vlNumRetencion As Long
Dim vlViaPago As String
Dim vlRutPen As String
Dim vlNombrePen As String
Dim vlNumRetAux As Long
Dim vlCargasRetenidas As String
Dim vlRutRecAux As String
Dim vlFecha As String
Dim vlTipoIdenRec As Integer
Dim vlMoneda As String

Dim vlGlsUsuarioCrea As Variant
Dim vlFecCrea As Variant
Dim vlHorCrea As Variant
Dim vlGlsUsuarioModi As Variant
Dim vlFecModi As Variant
Dim vlHorModi As Variant

Dim vlArchivo As String

Dim vlSwSeleccionado As Boolean

Const clTipoRetencion As String * 2 = "RJ"
Const clTipoAsigFam As String * 3 = "RAF"
Const clCodSinDerPen As String * 2 = "10"
Const clCodParConyugeSH As String * 5 = "10"
Const clCodParConyugeCH As String * 5 = "11"
Const clModRetAsigFam As String * 5 = "PESOS"
Const clTopeMtoRet As Double = 999999999.99
Const clFechaTopeTer As String * 8 = "99991231"

Const clGlsFrmDir As String = "FrmDir"

Dim vlCodTipoIdenBenCau As String
Dim vlNumIdenBenCau As String

Dim vlLargoTipoIden    As Integer 'sirve para llenar la grilla
Dim vlPosicionTipoIden As Integer 'sirve para llenar la grilla

Dim vlNombreSeg As String, vlApMaterno As String
Dim vlAfp As String

Function flRecibe(vlNumPoliza, vlCodTipoIden, vlNumIden, vlNumEndoso)
    Txt_PenPoliza = vlNumPoliza
    Call fgBuscarPosicionCodigoCombo(vlCodTipoIden, Cmb_PenNumIdent)
    Txt_PenNumIdent = vlNumIden
    ''Txt_PenDigito = vlDigito
    Lbl_End = vlNumEndoso
    Cmd_BuscarPol_Click
End Function

Function flLimpiar()
    
    Lbl_NumRetencion.Caption = ""
    Txt_FechaIniVig.Text = ""
    Txt_FechaTerVig.Text = ""
    Lbl_FechaEfecto.Caption = ""
    
    Txt_FecSuspension.Text = ""
    Txt_FechaRecepcion.Text = ""
    Cmb_TipoRetencion.ListIndex = 0
    Cmb_ModRetencion.ListIndex = 0
    Txt_MontoRetenido.Text = ""
    Txt_MtoMaxRet.Text = ""
    Txt_Juzgado.Text = ""
    
    Cmb_NumIdent.ListIndex = 0
    Txt_NumIdent.Text = ""
    ''*Txt_RutRec.Text = ""
''*    Txt_DigitoRec.Text = ""
    Txt_NombreRec.Text = ""
    Txt_NombreSegRec.Text = ""
    Txt_PaternoRec.Text = ""
    Txt_MaternoRec.Text = ""
    Txt_DireccionRec.Text = ""
    
''    Cmb_ComunaRec.ListIndex = 0
''    Call fgBuscarNombreProvinciaRegion(vlCodDir)
''    vlNombreRegion = vgNombreRegion
''    vlNombreProvincia = vgNombreProvincia
''    Lbl_Region.Caption = vlNombreRegion
''    Lbl_Provincia.Caption = vlNombreProvincia
    
    Lbl_Distrito.Caption = ""
    Lbl_Departamento.Caption = ""
    Lbl_Provincia.Caption = ""
    
    Txt_TelefonoRec.Text = ""
    Txt_CorreoRec.Text = ""
    
    Cmb_ViaPago.ListIndex = 0
    'Cmb_Sucursal.ListIndex = 0
    Call fgComboSucursal(Cmb_Sucursal, "S")
    Cmb_Sucursal.Enabled = False

    Cmb_TipoCta.ListIndex = 0
    Cmb_Banco.ListIndex = 0
    Txt_NumCuenta.Text = ""
    
    Call Cmb_ViaPago_Click
    
    ''Lbl_NumOrden.Caption = ""
''    Lbl_RutCarga.Caption = ""
''*    Lbl_DigitoCarga.Caption = ""
''    Lbl_NombreCarga.Caption = ""
''    Lbl_PaternoCarga.Caption = ""
''    Lbl_MaternoCarga.Caption = ""
    
    'inicializa grilla cargas
    ''Call flInicializaGrillaCargas
   
    
    If Txt_PenPoliza.Text = "" Then
       'inicializa grilla de Retencion Judicial
       Call flInicializaGrillaRetJud
       'inicializa grilla de beneficiarios
       Call flInicializaGrillaBeneficiarios
    End If
    
    'hqr 06/10/2007 Para que al limpiar vuelva a poner valor en fecha de termino de vigencia
    If Txt_FecSuspension.Text = "" Then
       Txt_FecSuspension.Text = DateSerial(Mid((clFechaTopeTer), 1, 4), Mid((clFechaTopeTer), 5, 2), Mid((clFechaTopeTer), 7, 2))
       Txt_FechaTerVig.Text = Txt_FecSuspension.Text
    End If
    
'    vlSwSeleccionado = False
    

End Function

Function flInicializaGrillaRetJud()
    
    Msf_GrillaRetJud.Clear
    Msf_GrillaRetJud.Cols = 11
    Msf_GrillaRetJud.rows = 1
    Msf_GrillaRetJud.RowHeight(0) = 250
    Msf_GrillaRetJud.row = 0
    
    Msf_GrillaRetJud.Col = 0
    Msf_GrillaRetJud.Text = "Vig. Desde"
    Msf_GrillaRetJud.ColWidth(0) = 1200
'    Msf_Grilla.ColAlignment(0) = 1  'centrado
    
    Msf_GrillaRetJud.Col = 1
    Msf_GrillaRetJud.Text = "Vig. Hasta"
    Msf_GrillaRetJud.ColWidth(1) = 1200
    
    Msf_GrillaRetJud.Col = 2
    Msf_GrillaRetJud.Text = "Tipo Ret."
    Msf_GrillaRetJud.ColWidth(2) = 2500

    Msf_GrillaRetJud.Col = 3
    Msf_GrillaRetJud.Text = "Modalidad"
    Msf_GrillaRetJud.ColWidth(3) = 1500
    
    Msf_GrillaRetJud.Col = 4
    Msf_GrillaRetJud.Text = "Monto"
    Msf_GrillaRetJud.ColWidth(4) = 1500
    
    Msf_GrillaRetJud.Col = 5
    Msf_GrillaRetJud.Text = "Tipo Ident. Receptor"
    Msf_GrillaRetJud.ColWidth(5) = 1600
    
    Msf_GrillaRetJud.Col = 6
    Msf_GrillaRetJud.Text = "Nº Ident. Receptor"
    Msf_GrillaRetJud.ColWidth(6) = 1500
    
    Msf_GrillaRetJud.Col = 7
    Msf_GrillaRetJud.Text = "Nombre Receptor"
    Msf_GrillaRetJud.ColWidth(7) = 3000
    
    Msf_GrillaRetJud.Col = 8
    Msf_GrillaRetJud.Text = "Ap. Paterno"
    Msf_GrillaRetJud.ColWidth(8) = 3000
    
    Msf_GrillaRetJud.Col = 9
    Msf_GrillaRetJud.Text = "Ap. Materno"
    Msf_GrillaRetJud.ColWidth(9) = 3000
    
    Msf_GrillaRetJud.Col = 10
    Msf_GrillaRetJud.Text = "Num. Retencion"
    Msf_GrillaRetJud.ColWidth(10) = 0
        
End Function

''Function flInicializaGrillaCargas()
''
''    Msf_GrillaCargas.Clear
''    Msf_GrillaCargas.Cols = 2
''    Msf_GrillaCargas.Rows = 1
''    Msf_GrillaCargas.RowHeight(0) = 250
''    Msf_GrillaCargas.Row = 0
''
''    Msf_GrillaCargas.Col = 0
''    Msf_GrillaCargas.Text = "Nº Orden"
''    Msf_GrillaCargas.ColWidth(0) = 1000
'''    Msf_Grilla.ColAlignment(0) = 1  'centrado
''
''    Msf_GrillaCargas.Col = 1
''    Msf_GrillaCargas.Text = "Rut"
''    Msf_GrillaCargas.ColWidth(1) = 1300
''
''End Function

Function flInicializaGrillaBeneficiarios()
    
''    Msf_GrillaBenef.Clear
''    Msf_GrillaBenef.Cols = 5
''    Msf_GrillaBenef.Rows = 1
''    Msf_GrillaBenef.RowHeight(0) = 250
''    Msf_GrillaBenef.Row = 0
''
''    Msf_GrillaBenef.Col = 0
''    Msf_GrillaBenef.Text = "Nº Orden"
''    Msf_GrillaBenef.ColWidth(0) = 1000
'''    Msf_Grilla.ColAlignment(0) = 1  'centrado
''
''    Msf_GrillaBenef.Col = 1
''    Msf_GrillaBenef.Text = "Rut"
''    Msf_GrillaBenef.ColWidth(1) = 1300
''
''    Msf_GrillaBenef.Col = 2
''    Msf_GrillaBenef.Text = "Nombre"
''    Msf_GrillaBenef.ColWidth(2) = 2500
''
''    Msf_GrillaBenef.Col = 3
''    Msf_GrillaBenef.Text = "Ap. Paterno"
''    Msf_GrillaBenef.ColWidth(3) = 1000
''
''    Msf_GrillaBenef.Col = 4
''    Msf_GrillaBenef.Text = "Ap. Materno"
''    Msf_GrillaBenef.ColWidth(4) = 1000
    
End Function

Function flHabilitarIngreso()

    Fra_Poliza.Enabled = False
    SSTab.Enabled = True
              
    Fra_AntecedentesRet.Enabled = True
    Fra_Receptor.Enabled = True
    Fra_Pago.Enabled = True
    Fra_Suspension.Enabled = True
    
    ''Fra_Cargas.Enabled = True
    ''Msf_GrillaBenef.Enabled = True

End Function

Function flDeshabilitarIngreso()

    Fra_Poliza.Enabled = True
    SSTab.Enabled = False

End Function

Function flCargaGrillaRetJud()
Dim vlTipoI As String
On Error GoTo Err_CargaGrillaRetJud
    
    vgSql = ""
    vgSql = "SELECT num_retencion,fec_iniret,fec_terret,cod_tipret,cod_modret, "
    vgSql = vgSql & "mto_ret,cod_tipoidenreceptor,num_idenreceptor, "
    vgSql = vgSql & "gls_nomreceptor,gls_patreceptor,gls_matreceptor "
    vgSql = vgSql & "FROM PP_TMAE_RETJUDICIAL "
    vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' "
    vgSql = vgSql & "ORDER by fec_iniret "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
       Call flInicializaGrillaRetJud
       
       While Not vgRs.EOF
       
          vlTipoI = Trim(vgRs!Cod_TipoIdenReceptor) & " - " & fgBuscarNombreTipoIden(Trim(vgRs!Cod_TipoIdenReceptor), False)
          vlTipoRet = Trim(vgRs!cod_tipret) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_TipRetJud, Trim(vgRs!cod_tipret)))
          vlModRet = Trim(vgRs!cod_modret) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_ModPagoRetJud, Trim(vgRs!cod_modret)))
                 
          Msf_GrillaRetJud.AddItem (DateSerial(Mid((vgRs!FEC_INIRET), 1, 4), Mid((vgRs!FEC_INIRET), 5, 2), Mid((vgRs!FEC_INIRET), 7, 2))) & vbTab _
          & (DateSerial(Mid((vgRs!FEC_TERRET), 1, 4), Mid((vgRs!FEC_TERRET), 5, 2), Mid((vgRs!FEC_TERRET), 7, 2))) & vbTab _
          & (Trim(vlTipoRet)) & vbTab _
          & (Trim(vlModRet)) & vbTab _
          & (Format((vgRs!mto_ret), "###,###,##0.00")) & vbTab _
          & vlTipoI & vbTab & (Trim(vgRs!Num_IdenReceptor)) & vbTab _
          & (Trim(vgRs!Gls_NomReceptor)) & vbTab _
          & (Trim(vgRs!Gls_PatReceptor)) & vbTab _
          & (Trim(vgRs!Gls_MatReceptor)) & vbTab _
          & (Trim(vgRs!num_retencion))
          
'          & (DateSerial(Mid((vgRs!fec_terret), 1, 4), Mid((vgRs!fec_terret), 5, 2), Mid((vgRs!fec_terret), 7, 2))) & vbTab _

          vgRs.MoveNext
       Wend
    End If
    vgRs.Close

Exit Function
Err_CargaGrillaRetJud:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function flCargaGrillaBeneficiarios()

On Error GoTo Err_CargaGrillaBeneficiarios
    
    

''    vlNumero = InStr(Cmb_TipoRetencion.Text, "-")
''    vlOpcion = Trim(Mid(Cmb_TipoRetencion.Text, 1, vlNumero - 1))
''
''    If vlOpcion = clTipoRetencion Then
''
''       vgSql = ""
''       vgSql = "SELECT b.num_orden,b.cod_tipoidenben,b.num_idenben, "
''       vgSql = vgSql & "b.gls_nomben,b.gls_patben,b.gls_matben,b.cod_par "
''       vgSql = vgSql & "FROM PP_TMAE_BEN b "
''       vgSql = vgSql & "WHERE b.num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
''       vgSql = vgSql & "b.num_idenben <> '" & Trim(Txt_PenNumIdent.Text) & "' AND "
''       vgSql = vgSql & " b.num_endoso = "
''       vgSql = vgSql & " (SELECT MAX(p.num_endoso) FROM pp_tmae_poliza p WHERE "
''       vgSql = vgSql & " p.num_poliza = b.num_poliza) "
''       vgSql = vgSql & "ORDER by b.num_orden "
''       Set vgRs = vgConexionBD.Execute(vgSql)
''       If Not vgRs.EOF Then
''          SSTab.TabEnabled(2) = True
''          SSTab.Tab = 0
''          Call flInicializaGrillaBeneficiarios
''
''          While Not vgRs.EOF
''                If (vgRs!Cod_Par) <> (clCodParConyugeSH) And _
''                   (vgRs!Cod_Par) <> (clCodParConyugeCH) Then
''                    Msf_GrillaBenef.AddItem (vgRs!Num_Orden) & vbTab _
''                    & (Trim(vgRs!num_idenben)) & vbTab _
''                    & (Trim(vgRs!Gls_NomBen)) & vbTab _
''                    & (Trim(vgRs!Gls_PatBen)) & vbTab _
''                    & (Trim(vgRs!Gls_MatBen))
''                    vgRs.MoveNext
''                End If
''          Wend
''       Else
''           SSTab.TabEnabled(2) = False
''           SSTab.Tab = 0
''
''       End If
''       vgRs.Close
''''*
''''       vgSql = ""
''''       vgSql = "SELECT n.num_orden,n.rut_ben,n.dgv_ben,"
''''       vgSql = vgSql & "n.gls_nomben,n.gls_patben,n.gls_matben "
''''       vgSql = vgSql & "FROM PP_TMAE_NOBEN n "
''''       vgSql = vgSql & "WHERE n.num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
''''       vgSql = vgSql & " n.num_endoso = "
''''       vgSql = vgSql & " (SELECT MAX(p.num_endoso) FROM pp_tmae_poliza p WHERE "
''''       vgSql = vgSql & " p.num_poliza = n.num_poliza) "
''''       vgSql = vgSql & "ORDER by n.num_orden "
''''       Set vgRs = vgConexionBD.Execute(vgSql)
''''       If Not vgRs.EOF Then
''''          SSTab.TabEnabled(2) = True
''''          SSTab.Tab = 0
''''
''''          While Not vgRs.EOF
''''             Msf_GrillaBenef.AddItem (vgRs!Num_Orden) & vbTab _
''''             & (" " & Format((Trim(vgRs!Rut_Ben)), "##,###,##0") & " - " & (Trim(vgRs!Dgv_Ben))) & vbTab _
''''             & (Trim(vgRs!Gls_NomBen)) & vbTab _
''''             & (Trim(vgRs!Gls_PatBen)) & vbTab _
''''             & (Trim(vgRs!Gls_MatBen))
''''             vgRs.MoveNext
''''          Wend
'''I-CMV -- 20050121
'''Modificado por que no debe desactivar la ficha 2 si no existen NOBEN, ya que
'''podria haber sido justamente activada por si tener BEN. No es necesario, ya que si no
'''tiene BEN ya fue desactivada dicha ficha
'''       Else
'''           SSTab.TabEnabled(2) = False
'''           SSTab.Tab = 0
'''F-CMV -- 20050121
''''       End If
''''       vgRs.Close
''
''    Else
''        If vlOpcion = clTipoAsigFam Then
''        'Solamente mostrar Conyuge
''           vgSql = ""
''           vgSql = "SELECT b.num_orden,b.rut_ben,b.dgv_ben,"
''           vgSql = vgSql & "b.gls_nomben,b.gls_patben,b.gls_matben "
''           vgSql = vgSql & "FROM PP_TMAE_BEN b "
''           vgSql = vgSql & "WHERE b.num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
''           vgSql = vgSql & "(b.cod_par = '" & Trim(clCodParConyugeSH) & "' OR "
''           vgSql = vgSql & "b.cod_par = '" & Trim(clCodParConyugeCH) & "') AND "
''           vgSql = vgSql & " b.num_endoso = "
''           vgSql = vgSql & " (SELECT MAX(p.num_endoso) FROM pp_tmae_poliza p WHERE "
''           vgSql = vgSql & " p.num_poliza = b.num_poliza) AND "
''           vgSql = vgSql & "rut_ben <> '" & Trim(Str(Txt_PenNumIdent.Text)) & "' "
''           vgSql = vgSql & "ORDER by num_orden "
''           Set vgRs = vgConexionBD.Execute(vgSql)
''           If Not vgRs.EOF Then
''              Call flInicializaGrillaBeneficiarios
''              SSTab.TabEnabled(2) = True
''              SSTab.Tab = 2
''              While Not vgRs.EOF
''
''                 Msf_GrillaBenef.AddItem (vgRs!Num_Orden) & vbTab _
''                 & (" " & Format((Trim(vgRs!Rut_Ben)), "##,###,##0") & " - " & (Trim(vgRs!Dgv_Ben))) & vbTab _
''                 & (Trim(vgRs!Gls_NomBen)) & vbTab _
''                 & (Trim(vgRs!Gls_PatBen)) & vbTab _
''                 & (Trim(vgRs!Gls_MatBen))
''                 vgRs.MoveNext
''              Wend
''           Else
''               SSTab.TabEnabled(2) = False
''               SSTab.Tab = 0
''           End If
''           vgRs.Close
''
''        End If
''    End If
    
Exit Function
Err_CargaGrillaBeneficiarios:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function flGenerarCorrelativoRet()

On Error GoTo Err_GenerarCorrelativoRet

    vgSql = ""
    vgSql = "SELECT num_retencion "
    vgSql = vgSql & "FROM PP_TMAE_RETJUDICIAL "
    vgSql = vgSql & "ORDER BY num_retencion DESC"
    Set vgRs2 = vgConexionBD.Execute(vgSql)
    If Not vgRs2.EOF Then
       vlNumRetencion = ((vgRs2!num_retencion) + 1)
    Else
        vlNumRetencion = 1
    End If

Exit Function
Err_GenerarCorrelativoRet:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function flMostrarDatosRetencion()

On Error GoTo Err_flMostrarDatosRetencion
        
    vgSql = ""
    vgSql = "SELECT *  FROM PP_TMAE_RETJUDICIAL "
    vgSql = vgSql & "WHERE num_retencion = '" & Trim(Lbl_NumRetencion.Caption) & "' "
'    vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
'    vgSql = vgSql & "num_orden = " & Trim(vlNumOrden) & " "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
    
       Lbl_NumRetencion.Caption = (vgRegistro!num_retencion)
       Txt_FechaIniVig.Text = DateSerial(Mid((vgRegistro!FEC_INIRET), 1, 4), Mid((vgRegistro!FEC_INIRET), 5, 2), Mid((vgRegistro!FEC_INIRET), 7, 2))
       
'       If IsNull(vgRs2!fec_terret) Then
'          Txt_FechaTerVig.Text = ""
'       Else
'           Txt_FechaTerVig.Text = DateSerial(Mid((vgRs2!fec_terret), 1, 4), Mid((vgRs2!fec_terret), 5, 2), Mid((vgRs2!fec_terret), 7, 2))
'       End If

       Txt_FecSuspension.Text = DateSerial(Mid((vgRegistro!FEC_TERRET), 1, 4), Mid((vgRegistro!FEC_TERRET), 5, 2), Mid((vgRegistro!FEC_TERRET), 7, 2))
       Txt_FechaTerVig.Text = Txt_FecSuspension.Text
       
       If CDate(Txt_FechaTerVig.Text) > Now Then
            chkViegencia.Value = 1
       Else
            chkViegencia.Value = 0
       End If
       
'       Txt_FechaTerVig.Text = (vgRs2!fec_terret)
       vlSw = True
       Txt_FechaRecepcion.Text = DateSerial(Mid((vgRegistro!fec_resdoc), 1, 4), Mid((vgRegistro!fec_resdoc), 5, 2), Mid((vgRegistro!fec_resdoc), 7, 2))
       Cmb_TipoRetencion.ListIndex = fgBuscarPosicionCodigoCombo(Trim(vgRegistro!cod_tipret), Cmb_TipoRetencion)
       Cmb_ModRetencion.ListIndex = fgBuscarPosicionCodigoCombo(Trim(vgRegistro!cod_modret), Cmb_ModRetencion)
       Txt_MontoRetenido.Text = Format((vgRegistro!mto_ret), "###,###,##0.00")
       Txt_MtoMaxRet.Text = Format((vgRegistro!mto_retmax), "###,###,##0.00")
       vlSw = False
              
       If IsNull(vgRegistro!FEC_EFECTO) Then
          Lbl_FechaEfecto.Caption = ""
       Else
           Lbl_FechaEfecto.Caption = (vgRegistro!FEC_EFECTO)
           Lbl_FechaEfecto.Caption = DateSerial((Mid(Lbl_FechaEfecto.Caption, 1, 4)), (Mid(Lbl_FechaEfecto.Caption, 5, 2)), (Mid(Lbl_FechaEfecto.Caption, 7, 2)))
       End If
              
       If IsNull(vgRegistro!gls_nomjuzgado) Then
          Txt_Juzgado.Text = ""
       Else
           Txt_Juzgado.Text = (vgRegistro!gls_nomjuzgado)
       End If

'       Txt_Juzgado.Text = (vgRs2!gls_nomjuzgado)
       vlCodTipoIdenBenCau = (vgRegistro!Cod_TipoIdenReceptor)
       Call fgBuscarPosicionCodigoCombo(vlCodTipoIdenBenCau, Cmb_NumIdent)
       Txt_NumIdent.Text = (vgRegistro!Num_IdenReceptor)
''       Txt_DigitoRec.Text = (vgRegistro!DGV_Receptor)
       Txt_NombreRec.Text = (vgRegistro!Gls_NomReceptor)
       If IsNull(vgRegistro!Gls_NomSegReceptor) Then
           Txt_NombreSegRec.Text = ""
       Else
           Txt_NombreSegRec.Text = (vgRegistro!Gls_NomSegReceptor)
       End If
       Txt_PaternoRec.Text = (vgRegistro!Gls_PatReceptor)
       If IsNull(vgRegistro!Gls_MatReceptor) Then
            Txt_MaternoRec.Text = ""
       Else
            Txt_MaternoRec.Text = (vgRegistro!Gls_MatReceptor)
        End If
       Txt_DireccionRec.Text = (vgRegistro!gls_dirreceptor)
              
''       vlCont = 0
''       Do While vlCont <= Cmb_ComunaRec.ListCount
''             If Cmb_ComunaRec.ItemData(vlCont) = (vgRegistro!Cod_Direccion) Then
''                Cmb_ComunaRec.ListIndex = vlCont
''                vlCont = Cmb_ComunaRec.ListCount + 1
''                Exit Do
''             End If
''             vlCont = vlCont + 1
''       Loop
          
       vlCodDir = (vgRegistro!Cod_Direccion)
       
       Call fgBuscarNombreProvinciaRegion(vlCodDir)
       vlNombreRegion = vgNombreRegion
       vlNombreProvincia = vgNombreProvincia
       vlNombreComuna = vgNombreComuna
       
        Lbl_Departamento.Caption = vlNombreRegion
        Lbl_Provincia.Caption = vlNombreProvincia
        Lbl_Distrito.Caption = vlNombreComuna

''       Lbl_Region.Caption = vlNombreRegion
''       Lbl_Provincia.Caption = vlNombreProvincia
       
       If IsNull(vgRegistro!gls_fonoreceptor) Then
          Txt_TelefonoRec.Text = ""
       Else
           Txt_TelefonoRec.Text = Trim(vgRegistro!gls_fonoreceptor)
       End If
       
       If IsNull(vgRegistro!gls_emailreceptor) Then
          Txt_CorreoRec.Text = ""
       Else
           Txt_CorreoRec.Text = (vgRegistro!gls_emailreceptor)
       End If
       
       Cmb_ViaPago.ListIndex = fgBuscarPosicionCodigoCombo(Trim(vgRegistro!Cod_ViaPago), Cmb_ViaPago)
       Cmb_TipoCta.ListIndex = fgBuscarPosicionCodigoCombo(Trim(vgRegistro!Cod_TipCuenta), Cmb_TipoCta)
       Cmb_Banco.ListIndex = fgBuscarPosicionCodigoCombo(Trim(vgRegistro!Cod_Banco), Cmb_Banco)
                
       vlcont = 0
'       Cmb_Sucursal.ListIndex = 0
'       Do While vlCont < Cmb_Sucursal.ListCount
'          If (Trim(Cmb_Sucursal) <> "") Then
'             If (vgRegistro!Cod_Sucursal = Trim(Mid(Cmb_Sucursal.Text, 1, (InStr(1, Cmb_Sucursal, "-") - 1)))) Then
'                Exit Do
'             End If
'          End If
'          vlCont = vlCont + 1
'          'I---- ABV 23/08/2004 ---
'          If (vlCont < Cmb_Sucursal.ListCount) Then
'            Cmb_Sucursal.ListIndex = vlCont
'          End If
'          'F---- ABV 23/08/2004 ---
'       Loop
'
'       'I---- ABV 23/08/2004 ---
'       'Cmb_Sucursal.ListIndex = vlCont
'        If (vlCont >= Cmb_Sucursal.ListCount) Then
'            Cmb_Sucursal.ListIndex = 0
'        End If
'       'F---- ABV 23/08/2004 ---
       
       If IsNull(vgRegistro!Num_Cuenta) Then
            Txt_NumCuenta.Text = ""
       Else
            Txt_NumCuenta.Text = Trim(vgRegistro!Num_Cuenta)
       End If
'       vgRs2.Close
               
         'hqr 06/10/2007 Se comenta para que se habilite frame si la fecha de termino no es igual a la fecha de tope
''       If (vgRegistro!fec_terret) = clFechaTopeTer Then
          Call flHabilitarIngreso
''       Else
''
''           Fra_Poliza.Enabled = False
''           SSTab.Enabled = True
''
''           Fra_AntecedentesRet.Enabled = False
''           Fra_Receptor.Enabled = False
''           Fra_Pago.Enabled = False
''           Fra_Suspension.Enabled = False
''
''           ''Fra_Cargas.Enabled = False
''           ''Msf_GrillaBenef.Enabled = False
''
''       End If
                    
       Call Cmb_ViaPago_Click
       
       Call Cmb_TipoRetencion_Click
       
    Else
        vgRegistro.Close
        MsgBox "El Beneficiario ingresado No tiene Retenciones Judiciales", vbInformation, "Información"
        Call flHabilitarIngreso
        Call Cmb_ViaPago_Click
    End If
    
Exit Function
Err_flMostrarDatosRetencion:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function flEliminarDetCargas()

    vgSql = ""
    vgSql = "DELETE PP_TMAE_DETRETENCION WHERE "
    vgSql = vgSql & "num_retencion = '" & vlNumRetencion & "' "
    vgConexionBD.Execute (vgSql)
    
End Function

Function flValidaReceptor()

On Error GoTo Err_flValidaReceptor
    
    vlNumRetAux = 0
    vlRutRecAux = ""
    
    vlNumero = InStr(Cmb_NumIdent.Text, "-")
    vlTipoIdenRec = Trim(Mid(Cmb_NumIdent.Text, 1, vlNumero - 1))
            
    vgSql = ""
    vgSql = "SELECT num_retencion,num_idenreceptor "
    vgSql = vgSql & "FROM PP_TMAE_RETJUDICIAL "
    vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza) & "' AND "
    vgSql = vgSql & "num_orden = " & str(vlNumOrden) & " AND "
    vgSql = vgSql & "cod_tipret = '" & Trim(clTipoRetencion) & "' AND "
    vgSql = vgSql & "fec_terret = '" & Trim(clFechaTopeTer) & "' AND "
    vgSql = vgSql & "cod_tipoidenreceptor = " & vlTipoIdenRec & " AND "
    vgSql = vgSql & "num_idenreceptor = " & str(Txt_NumIdent.Text) & " "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       vlNumRetAux = (vgRegistro!num_retencion)
       vlRutRecAux = (vgRegistro!Num_IdenReceptor)
     
'       If vlNumRetAux <> CLng(Lbl_NumRetencion.Caption) Then
'          Exit Function
'       End If
        If vlNumRetAux <> vlNumRetencion Then
           Exit Function
        End If
        
'       vgRegistro.MoveNext
    End If

Exit Function
Err_flValidaReceptor:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function flValidaRetAsigFam()

On Error GoTo Err_flValidaRetAsigFam
    
    vlNumRetAux = 0
    
    vgSql = ""
    vgSql = "SELECT num_retencion "
    vgSql = vgSql & "FROM PP_TMAE_RETJUDICIAL "
    vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza) & "' AND "
    vgSql = vgSql & "num_orden = " & str(vlNumOrden) & " AND "
    vgSql = vgSql & "cod_tipret = '" & Trim(clTipoAsigFam) & "' AND "
    vgSql = vgSql & "fec_terret = '" & Trim(clFechaTopeTer) & "' "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       vlNumRetAux = (vgRegistro!num_retencion)
       If vlNumRetAux <> CLng(Lbl_NumRetencion.Caption) Then
          Exit Function
       End If
       vgRegistro.MoveNext
    End If

Exit Function
Err_flValidaRetAsigFam:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function


Private Sub chkViegencia_Click()
    If chkViegencia.Value = 1 Then
        Txt_FechaTerVig = "31/12/9999"
    Else
       Txt_FechaTerVig = Format(CStr(Now), "DD/MM/yyyy")
    End If
End Sub

Private Sub Cmb_Banco_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
       Txt_NumCuenta.SetFocus
    End If
End Sub

Private Sub Cmb_ModRetencion_Click()
    Select Case Mid(Cmb_ModRetencion.Text, 1, 5)
        Case "PRCIM"
            Lbl_Nombre(31).Caption = "% Retenido"
            Txt_MtoMaxRet.Text = 60
        Case "MTOIM"
            Lbl_Nombre(31).Caption = "Monto Retenido"
            If txtPensionAct.Text <> "" Then
                 Txt_MtoMaxRet.Text = Format(CDbl(txtPensionAct.Text) * 0.6, "##0.00")
            End If
           
    End Select
    'Txt_MontoRetenido.SetFocus
   
End Sub

Private Sub Cmb_ModRetencion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Txt_MontoRetenido.SetFocus
    End If
End Sub

Private Sub Cmb_NumIdent_Click()
If (Cmb_NumIdent <> "") Then
    vlPosicionTipoIden = Cmb_NumIdent.ListIndex
    vlLargoTipoIden = Cmb_NumIdent.ItemData(vlPosicionTipoIden)
    If (vlLargoTipoIden = 0) Then
        Txt_NumIdent.Text = "0"
        Txt_NumIdent.Enabled = False
    Else
        Txt_NumIdent = ""
        Txt_NumIdent.Enabled = True
        Txt_NumIdent.MaxLength = vlLargoTipoIden
        If (Txt_NumIdent <> "") Then Txt_NumIdent.Text = Mid(Txt_NumIdent, 1, vlLargoTipoIden)
    End If
End If
End Sub

Private Sub Cmb_NumIdent_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
        If (Txt_NumIdent.Enabled = True) Then
            Txt_NumIdent.SetFocus
        Else
            Txt_NombreRec.SetFocus
        End If
    End If
End Sub

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

Private Sub Txt_DireccionRec_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txt_DireccionRec.Text = Trim(UCase(Txt_DireccionRec.Text))
        Cmd_BuscarDir.SetFocus
    End If
End Sub

Private Sub Txt_DireccionRec_LostFocus()
    Txt_DireccionRec.Text = Trim(UCase(Txt_DireccionRec.Text))
End Sub

Private Sub Txt_MtoMaxRet_Change()
On Error GoTo Err_TxtMtoMaxRetChange

    If Not IsNumeric(Txt_MtoMaxRet) Then
       Txt_MtoMaxRet = ""
    End If
    
Exit Sub
Err_TxtMtoMaxRetChange:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Txt_MtoMaxRet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Txt_MtoMaxRet.Text = "" Then
          MsgBox "Debe Ingresar Máxima Retención.", vbCritical, "Error de Datos"
          Txt_MtoMaxRet.SetFocus
          Exit Sub
       Else
           Txt_MtoMaxRet.Text = Format(Txt_MtoMaxRet.Text, "###,###,##0.00")
       End If
    
       If CDbl(Txt_MtoMaxRet.Text) > clTopeMtoRet Then
          MsgBox "El Valor Ingresado en Máxima Retención Excede el Máximo Permitido de " & Format(clTopeMtoRet, "###,###,##0.00") & "  .", vbExclamation, "Información"
          Txt_MtoMaxRet.Text = Format(clTopeMtoRet, "###,###,##0.00")
          Txt_MtoMaxRet.SetFocus
          Exit Sub
       End If
       Txt_Juzgado.SetFocus
    End If
End Sub

Private Sub Txt_MtoMaxRet_LostFocus()
   If Txt_MtoMaxRet.Text = "" Then
       Exit Sub
    Else
        Txt_MtoMaxRet.Text = Format(Txt_MtoMaxRet.Text, "###,###,##0.00")
    End If
    If CDbl(Txt_MtoMaxRet.Text) = 0 Then
       Exit Sub
    End If
    If CDbl(Txt_MtoMaxRet.Text) > clTopeMtoRet Then
       Txt_MtoMaxRet.Text = Format(clTopeMtoRet, "###,###,##0.00")
       Exit Sub
    End If
End Sub

Private Sub Txt_NombreSegRec_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txt_NombreSegRec.Text = Trim(UCase(Txt_NombreSegRec.Text))
        Txt_PaternoRec.SetFocus
    End If
End Sub

Private Sub Txt_NombreSegRec_LostFocus()
 Txt_NombreSegRec.Text = Trim(UCase(Txt_NombreSegRec.Text))
    If Txt_NombreSegRec.Text = "" Then
       Exit Sub
    End If
End Sub

Private Sub Txt_NumIdent_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
        If (Trim(Txt_PenNumIdent) <> "") Then
            Txt_PenNumIdent = UCase(Trim(Txt_PenNumIdent))
        End If
        Txt_NombreRec.SetFocus
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

Private Sub Cmb_Sucursal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Cmb_TipoCta.Enabled = True Then
          Cmb_TipoCta.SetFocus
       Else
           vlNumero = InStr(Cmb_TipoRetencion.Text, "-")
           vlOpcion = Trim(Mid(Cmb_TipoRetencion.Text, 1, vlNumero - 1))
            
           If vlOpcion = clTipoRetencion Then
              ''SSTab.Tab = 2
              Txt_FecSuspension.SetFocus
           Else
               If vlOpcion = clTipoAsigFam Then
                  Cmd_Grabar.SetFocus
               End If
           End If
           
'           If SSTab.TabEnabled(2) = True Then
'              SSTab.Tab = 2
'           Else
'               Cmd_Grabar.SetFocus
'           End If
       End If
    End If
End Sub

Private Sub Cmb_TipoCta_KeyPress(KeyAscii As Integer)
On Error GoTo Err_CmbTipoCtaKeyPress

    If KeyAscii = 13 Then
       Cmb_Banco.SetFocus
    End If
    
Exit Sub
Err_CmbTipoCtaKeyPress:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmb_TipoRetencion_Click()

On Error GoTo Err_CmbTipoRetencion

If vlSw = False Then
    vlNumero = InStr(Cmb_TipoRetencion.Text, "-")
    vlOpcion = Trim(Mid(Cmb_TipoRetencion.Text, 1, vlNumero - 1))
    
    If vlOpcion = clTipoRetencion Then
       
       If Txt_MontoRetenido.Text = "0" Then
          Cmb_ModRetencion.ListIndex = 0
          Txt_MontoRetenido.Text = ""
          Txt_Juzgado.Text = ""
       End If
       
       Cmb_ModRetencion.Enabled = True
       Txt_MontoRetenido.Enabled = True
       Txt_Juzgado.Enabled = True
       
       Call flCargaGrillaBeneficiarios
       
'       SSTab.TabEnabled(2) = True
       
        
    Else
        If vlOpcion = clTipoAsigFam Then
           
'           If Txt_MontoRetenido.Text = "" Then
              'Busca posicion Modalidad de Retencion en PESOS
              If Cmb_ModRetencion.ListIndex <> -1 Then
                 Cmb_ModRetencion.ListIndex = fgBuscarPosicionCodigoCombo(Trim(clModRetAsigFam), Cmb_ModRetencion)
              End If
              Txt_MontoRetenido.Text = "0"
              Txt_Juzgado.Text = ""
'           End If
           
           Cmb_ModRetencion.Enabled = False
           Txt_MontoRetenido.Enabled = False
           Txt_Juzgado.Enabled = False
           
       
           ''Call flInicializaGrillaCargas
           Call flCargaGrillaBeneficiarios
                       
           '''''''SSTab.Tab = 2
       
'            SSTab.TabEnabled(2) = True
        End If
    End If
    
''    '*Busca las Cargas
''    If Lbl_NumRetencion.Caption <> "" Then
''       vgSql = ""
''       vgSql = "SELECT cod_tipret  FROM PP_TMAE_RETJUDICIAL "
''       vgSql = vgSql & "WHERE num_retencion = '" & Trim(Lbl_NumRetencion.Caption) & "' "
''       Set vgRs = vgConexionBD.Execute(vgSql)
''       If Not vgRs.EOF Then
''          If Trim(vgRs!cod_tipret) = clTipoRetencion Then
''             If vlOpcion = clTipoRetencion Then
''                Call flCargaGrillaCargas
''             Else
''                 If vlOpcion = clTipoAsigFam Then
''                    Call flInicializaGrillaCargas
''                 End If
''             End If
''          Else
''              If Trim(vgRs!cod_tipret) = clTipoAsigFam Then
''                 If vlOpcion = clTipoRetencion Then
''                    Call flInicializaGrillaCargas
''                 Else
''                     If vlOpcion = clTipoAsigFam Then
''                        Call flCargaGrillaCargas
''                     End If
''                 End If
''              End If
''          End If
''       End If
''    End If
            
            
            
End If
            

Exit Sub
Err_CmbTipoRetencion:
Screen.MousePointer = 0
Select Case Err
    Case Else
    MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
End Select


End Sub

Private Sub Cmb_TipoRetencion_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       If Cmb_ModRetencion.Enabled = True Then
          Cmb_ModRetencion.SetFocus
       Else
           Txt_NumIdent.SetFocus
       End If
       Call Cmb_TipoRetencion_Click
       
    End If

End Sub

Private Sub Cmb_ViaPago_Click()
On Error GoTo Err_CmbViaPagoClick

If Cmb_ViaPago.Enabled = True Then

    vlNumero = InStr(Cmb_ViaPago.Text, "-")
    vlOpcion = Trim(Mid(Cmb_ViaPago.Text, 1, vlNumero - 1))
'Opción 01 = Via de Pago CAJA
    If vlSw = False Then
       
        If vlOpcion = "00" Or vlOpcion = "05" Then
            Cmb_Sucursal.Enabled = False
            Cmb_TipoCta.Enabled = False
            Cmb_Banco.Enabled = False
            Txt_NumCuenta.Enabled = False
            Cmb_Sucursal.ListIndex = 0
            Cmb_TipoCta.ListIndex = 0
            Cmb_Banco.ListIndex = 0
            Txt_NumCuenta.Text = ""
        Else
            If vlOpcion = "01" Or vlOpcion = "04" Then
                If (vlOpcion = "04") Then
                    vgTipoSucursal = cgTipoSucursalAfp
                Else
                    vgTipoSucursal = cgTipoSucursalSuc
                End If
                fgComboSucursal Cmb_Sucursal, vgTipoSucursal
                
                Cmb_Sucursal.Enabled = True
                Cmb_TipoCta.Enabled = False
                Cmb_Banco.Enabled = False
                Txt_NumCuenta.Enabled = False
                Cmb_TipoCta.ListIndex = 0
                Cmb_Banco.ListIndex = 0
                If (vlOpcion = "04") Then
                    vgPalabra = fgObtenerCodigo_TextoCompuesto(vlAfp)
                    Call fgBuscarPosicionCodigoCombo(vgPalabra, Cmb_Sucursal)
                End If
                Txt_NumCuenta.Text = ""
            Else
                If (vlOpcion = "02") Or (vlOpcion = "03") Then
                   Cmb_Sucursal.Enabled = False
                   Cmb_TipoCta.Enabled = True
                   Cmb_Banco.Enabled = True
                   Txt_NumCuenta.Enabled = True
                   Cmb_Sucursal.ListIndex = 0
'                   Txt_NumCuenta.Text = ""
                Else
                    Cmb_Sucursal.Enabled = True
                    Cmb_TipoCta.Enabled = True
                    Cmb_Banco.Enabled = True
                    Txt_NumCuenta.Enabled = True
                    Cmb_Sucursal.ListIndex = 0
                    Cmb_TipoCta.ListIndex = 0
                    Cmb_Banco.ListIndex = 0
                    Txt_NumCuenta.Text = ""
                End If
            End If
        End If
    End If
End If
    
Exit Sub
Err_CmbViaPagoClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmb_ViaPago_KeyPress(KeyAscii As Integer)
On Error GoTo Err_CmbViaPagoKeyPress

    If KeyAscii = 13 Then
       vlNumero = InStr(Cmb_ViaPago.Text, "-")
       vlOpcion = Trim(Mid(Cmb_ViaPago.Text, 1, vlNumero - 1))
        If vlOpcion = "00" Or vlOpcion = "05" Then
            Cmb_Sucursal.Enabled = False
            Cmb_TipoCta.Enabled = False
            Cmb_Banco.Enabled = False
            Txt_NumCuenta.Enabled = False
            Cmb_Sucursal.ListIndex = 0
            Cmb_TipoCta.ListIndex = 0
            Cmb_Banco.ListIndex = 0
            
            vlNumero = InStr(Cmb_TipoRetencion.Text, "-")
            vlOpcion = Trim(Mid(Cmb_TipoRetencion.Text, 1, vlNumero - 1))
            
            If vlOpcion = clTipoRetencion Then
               ''SSTab.Tab = 2
               Cmd_Grabar.SetFocus
            Else
                If vlOpcion = clTipoAsigFam Then
                   Cmd_Grabar.SetFocus
                End If
            End If
            
'            If SSTab.TabEnabled(2) = True Then
'               SSTab.Tab = 2
'            Else
'                Cmd_Grabar.SetFocus
'            End If
        Else
'Opción 01 = Via de Pago CAJA
            If vlOpcion = "01" Or vlOpcion = "04" Then
                If (vlOpcion = "04") Then
                    vgTipoSucursal = cgTipoSucursalAfp
                Else
                    vgTipoSucursal = cgTipoSucursalSuc
                End If
                fgComboSucursal Cmb_Sucursal, vgTipoSucursal
                
                Cmb_Sucursal.Enabled = True
                Cmb_TipoCta.ListIndex = 0
                Cmb_TipoCta.Enabled = False
                Cmb_Banco.ListIndex = 0
                Cmb_Banco.Enabled = False
                Txt_NumCuenta.Enabled = False
                If (vlOpcion = "04") Then
                    vgPalabra = fgObtenerCodigo_TextoCompuesto(vlAfp)
                    Call fgBuscarPosicionCodigoCombo(vgPalabra, Cmb_Sucursal)
                End If
                Txt_NumCuenta.Text = ""
            Else
                If (vlOpcion = "02") Or (vlOpcion = "03") Then
                   Cmb_Sucursal.Enabled = False
                   Cmb_TipoCta.Enabled = True
                   Cmb_Banco.Enabled = True
                   Txt_NumCuenta.Enabled = True
                   Cmb_Sucursal.ListIndex = 0
                   Txt_NumCuenta.Text = ""
                Else
                    If (vlOpcion = "02") Or (vlOpcion = "03") Then
                       Cmb_Sucursal.Enabled = True
                       Cmb_TipoCta.Enabled = True
                       Cmb_Banco.Enabled = True
                       Txt_NumCuenta.Enabled = True
                       Cmb_Sucursal.ListIndex = 0
                       Cmb_TipoCta.ListIndex = 0
                       Cmb_Banco.ListIndex = 0
                       Txt_NumCuenta.Text = ""
                    Else
                        Cmb_Sucursal.Enabled = True
                        Cmb_TipoCta.Enabled = True
                        Cmb_Banco.Enabled = True
                        Txt_NumCuenta.Enabled = True
                        Cmb_Sucursal.ListIndex = 0
                        Cmb_TipoCta.ListIndex = 0
                        Cmb_Banco.ListIndex = 0
                        Txt_NumCuenta.Text = ""
                    End If
                End If
                
''                Cmb_Sucursal.ListIndex = 0
''                Cmb_Sucursal.Enabled = False
''                Cmb_TipoCta.Enabled = True
''                Cmb_Banco.Enabled = True
''                Txt_NumCuenta.Enabled = True
            End If
            
            
            If Cmb_Sucursal.Enabled = True Then
               Cmb_Sucursal.SetFocus
            Else
                Cmb_TipoCta.SetFocus
            End If
            
        End If

        
    End If
    
Exit Sub
Err_CmbViaPagoKeyPress:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Buscar_Click()
On Error GoTo Err_CmdBuscarClick

    Frm_Busqueda.flInicio ("Frm_RetJudicial")
    
Exit Sub
Err_CmdBuscarClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_BuscarDir_Click()
On Error GoTo Err_Buscar

    Frm_BusDireccion.flInicio ("Frm_RetJudicial")
    
Exit Sub
Err_Buscar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub


Private Sub Cmd_BuscarPol_Click()
On Error GoTo Err_CmdBuscarPolClick
        
    vlSwSeleccionado = False
        
    If Txt_PenPoliza.Text = "" Then
       If ((Trim(Cmb_PenNumIdent.Text)) = "") Or (Txt_PenNumIdent.Text = "") Then
       ''*Or _ (Not ValiRut(Txt_PenRut.Text, Txt_PenDigito.Text))
           MsgBox "Debe Ingresar el Número de Póliza o la Identificación del Pensionado.", vbCritical, "Error de Datos"
           Txt_PenPoliza.SetFocus
           Exit Sub
       Else
           ''Txt_PenRut = Format(Txt_PenRut, "##,###,##0")
           Txt_PenNumIdent = UCase(Trim(Txt_PenNumIdent))
           Txt_PenNumIdent.SetFocus
           ''vlRutAux = Format(Txt_PenRut, "#0")
       End If
    Else
        Txt_PenPoliza.Text = Trim(Txt_PenPoliza.Text)
    End If
    
    vlCodTipoIdenBenCau = fgObtenerCodigo_TextoCompuesto(Cmb_PenNumIdent)
    vlNumIdenBenCau = Txt_PenNumIdent
    
    vgPalabra = ""
    'Seleccionar beneficiario, según número de póliza y rut de beneficiario.
    If (Txt_PenPoliza.Text) And (Cmb_PenNumIdent.Text <> "") And (Txt_PenNumIdent.Text <> "") Then
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
        
        vlAfp = fgObtenerPolizaCod_AFP(vgRs2!num_poliza, CStr(vgRs2!num_endoso))
        
       If Trim(vgRs2!Cod_EstPension) = Trim(clCodSinDerPen) Then
          MsgBox " El Beneficiario Seleccionado No Tiene Derecho a Pensión " & Chr(13) & _
          "          Sólo podrá Consultar los Datos del Registro", vbInformation, "Información"

          'Desactivar Todos los Controles del Formulario
          Fra_Poliza.Enabled = False
          SSTab.Enabled = True
          Fra_AntecedentesRet.Enabled = False
          Fra_Receptor.Enabled = False
          Fra_Pago.Enabled = False
          Fra_Suspension.Enabled = False
          Msf_GrillaRetJud.Enabled = True
          ''Fra_Cargas.Enabled = False
          ''Msf_GrillaBenef.Enabled = False
          SSTab.Tab = 1

       Else
                             
           If Not fgValidaVigenciaPoliza(Trim(Txt_PenPoliza.Text), fgBuscaFecServ) Then
              MsgBox " La Póliza Ingresada no se Encuentra Vigente en el Sistema " & Chr(13) & _
                     "      Sólo podrá Consultar los Datos del Registro", vbInformation, "Información"
              
              'Desactivar Todos los Controles del Formulario
              Fra_Poliza.Enabled = False
              SSTab.Enabled = True
              Fra_AntecedentesRet.Enabled = False
              Fra_Receptor.Enabled = False
              Fra_Pago.Enabled = False
              Fra_Suspension.Enabled = False
              Msf_GrillaRetJud.Enabled = True
              ''Fra_Cargas.Enabled = False
              ''Msf_GrillaBenef.Enabled = False
              SSTab.Tab = 1
           Else
               Call flHabilitarIngreso
           End If
             
       End If
              
       
        vlCodTipoIdenBenCau = vgRs2!Cod_TipoIdenBen
        vlNumIdenBenCau = Trim(vgRs2!Num_IdenBen)
             
        Txt_PenPoliza.Text = Trim(vgRs2!num_poliza)
        Call fgBuscarPosicionCodigoCombo(vlCodTipoIdenBenCau, Cmb_PenNumIdent)
        Txt_PenNumIdent.Text = vlNumIdenBenCau
        
        If IsNull(vgRs2!Gls_NomSegBen) Then
          vlNombreSeg = ""
       Else
           vlNombreSeg = Trim(vgRs2!Gls_NomSegBen)
       End If
        If IsNull(vgRs2!Gls_MatBen) Then
          vlApMaterno = ""
       Else
           vlApMaterno = Trim(vgRs2!Gls_MatBen)
       End If
       Lbl_PenNombre.Caption = fgFormarNombreCompleto(Trim(vgRs2!Gls_NomBen), vlNombreSeg, Trim(vgRs2!Gls_PatBen), vlApMaterno)
       ''*Lbl_PenNombre.Caption = Trim(vgRs2!Gls_NomBen) + " " + Trim(vgRs2!Gls_NomSegBen) + " " + Trim(vgRs2!Gls_PatBen) + " " + Trim(vgRs2!Gls_MatBen)
       Lbl_End.Caption = (vgRs2!num_endoso)
       vlNumEndoso = (vgRs2!num_endoso)
       vlNumOrden = (vgRs2!Num_Orden)
       
       'Busca la Moneda de la póliza seleccionada
       Lbl_Moneda.Caption = flBuscaMoneda(Trim(Txt_PenPoliza.Text), Lbl_End)
       
           
           
       Call flCargaGrillaRetJud
       Call Cmb_TipoRetencion_Click
       SSTab.Tab = 0
'       Call flCargaGrillaBeneficiarios
       'Call flCargaGrillaCargas
       
    Else
        MsgBox "El Beneficiario o la Póliza Ingresados, No Existen en la Base de Datos", vbInformation, "Información"
        Txt_PenPoliza.SetFocus
        Exit Sub
    End If
    vgRs2.Close
    
          
    If Fra_AntecedentesRet.Enabled = True Then
       Txt_FechaIniVig.SetFocus
    End If
       
    Dim valPensionSal, valPensionNet As Double
       
    'trae la pension actual del pensionista
    vgSql = "select p.num_poliza, num_orden, "
    vgSql = vgSql & "case when a.mto_pension is null then "
    vgSql = vgSql & "(case when num_mesgar = 0 then (p.mto_pension *  prc_pension)/100 else (p.mto_pension *  prc_pensiongar)/100 end) "
    vgSql = vgSql & "Else "
    vgSql = vgSql & "(case when num_mesgar = 0 then (a.mto_pension *  prc_pension)/100 else (a.mto_pension *  prc_pensiongar)/100 end) "
    vgSql = vgSql & "end As Pension "
    vgSql = vgSql & "from pp_tmae_poliza p "
    vgSql = vgSql & "Left Join "
    vgSql = vgSql & "(select num_poliza, num_endoso, fec_desde, mto_pension, mto_pensiongar from pp_tmae_pensionact "
    vgSql = vgSql & "where fec_desde='20140401'"
    vgSql = vgSql & ") a on p.num_poliza=a.num_poliza and p.num_endoso=a.num_endoso "
    vgSql = vgSql & "join pp_tmae_ben b on p.num_poliza=b.num_poliza and p.num_endoso=b.num_endoso "
    vgSql = vgSql & "where p.num_endoso=(select max(num_endoso) from pp_tmae_poliza where num_poliza=p.num_poliza) "
    vgSql = vgSql & "and p.num_poliza='" & Txt_PenPoliza & "' and num_orden=1 "
    vgSql = vgSql & "order by 1 "

    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
        valPensionSal = Format(CDbl(vgRs!Pension) * 0.04, "##0.00")
        txtPensionAct.Text = Format(CDbl(vgRs!Pension) - valPensionSal, "##0.00")
        valPensionNet = txtPensionAct.Text
        Txt_MtoMaxRet.Text = Format(CDbl(txtPensionAct.Text) * 0.6, "##0.00")
    End If
    
    Dim SumRetenciones As Double
    SumRetenciones = 0
    vgSql = "select cod_modret, mto_ret from PP_TMAE_RETJUDICIAL where num_poliza='" & Txt_PenPoliza & "' and num_endoso=2"
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
       While Not vgRs.EOF
            Select Case vgRs!cod_modret
                Case "PRCIM"
                    SumRetenciones = SumRetenciones + Format(valPensionNet * (CDbl(vgRs!mto_ret) / 100), "##0.00")
                Case "MTOIM"
                    SumRetenciones = SumRetenciones + Format(CDbl(vgRs!mto_ret), "##0.00")
            End Select
            vgRs.MoveNext
       Wend
       lblMontoRetAct.Caption = CDbl(Txt_MtoMaxRet) - SumRetenciones
    Else
       lblMontoRetAct.Caption = txtPensionAct.Text
    End If
       
vgRs.Close
       
Exit Sub
Err_CmdBuscarPolClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Cancelar2_Click()

''''    Txt_PenPoliza.Text = ""
''''    Txt_PenRut.Text = ""
''''    Txt_PenDigito.Text = ""
''''    Lbl_PenNombre.Caption = ""
    Lbl_End.Caption = ""

    Call flDeshabilitarIngreso
    Call Cmd_Limpiar_Click
    
    Txt_PenPoliza.Text = ""
    Txt_PenNumIdent.Text = ""
''    Txt_PenDigito.Text = ""
''    Lbl_PenEndoso.Caption = ""
    Cmb_PenNumIdent.ListIndex = 0
    Lbl_PenNombre.Caption = ""
    
    vlSw = True
    
    Call flLimpiar
    vlAfp = ""
    
    vlSw = False
    
    Call flDeshabilitarIngreso
    SSTab.Tab = 0
    
    Txt_PenPoliza.SetFocus
    
    

End Sub

Private Sub Cmd_Eliminar_Click()

On Error GoTo Err_CmdEliminarClick

If Fra_AntecedentesRet.Enabled = True Then


    If Trim(Lbl_NumRetencion.Caption) <> "" Then
        'hqr 06/10/2007 Se comenta para que no se asuma que la fecha de termino debe ser siempre la fecha tope
       'If Format(CDate(Trim(Txt_FechaTerVig.Text)), "yyyymmdd") = clFechaTopeTer Then
          
          vgSql = ""
          vgSql = "SELECT num_orden "
          vgSql = vgSql & "FROM PP_TMAE_RETJUDICIAL "
          vgSql = vgSql & "WHERE num_retencion = '" & Trim(Lbl_NumRetencion.Caption) & "' "
          vgSql = vgSql & "ORDER by num_orden "
          Set vgRs = vgConexionBD.Execute(vgSql)
          If Not vgRs.EOF Then
                 vgRes = MsgBox("¿ Está seguro que desea Eliminar los Datos ?", 4 + 32 + 256, "Operación de Actualización")
                 If vgRes <> 6 Then
                    vgRs.Close
                    Screen.MousePointer = 0
                    Exit Sub
                 End If
                 
                'Valida si la fecha de inicio es mayor o menor que la fecha del último proceso de cálculo de pensiones.
                
                 vlFecha = Txt_FechaIniVig.Text
                                 
                 vgSql = ""
                 vgSql = "SELECT fec_iniret "
                 vgSql = vgSql & "FROM PP_TMAE_RETJUDICIAL "
                 vgSql = vgSql & "WHERE num_retencion = '" & Trim(Lbl_NumRetencion.Caption) & "' "
                 Set vgRegistro = vgConexionBD.Execute(vgSql)
                 If Not vgRegistro.EOF Then
                    vlFecha = DateSerial(Mid((vgRegistro!FEC_INIRET), 1, 4), Mid((vgRegistro!FEC_INIRET), 5, 2), Mid((vgRegistro!FEC_INIRET), 7, 2))
                 End If
                
                 If Not fgValidaPagoPension(vlFecha, Trim(Txt_PenPoliza), vlNumOrden) Then
                        MsgBox " Ya se ha Realizado el Proceso de Cálculo de Pensión para ésta Fecha " & Chr(13) & _
                               "                    El Registro No Será Eliminado                    ", vbCritical, "Operación Cancelada"
                       Exit Sub
                 End If
                 
                
                
'                Else
                    'MsgBox "Eliminar retenciòn ingresada por el usuario", vbCritical, "Error de Datos"
                ''*Call flEliminarDetCargas
             
          End If
          
          vgSql = ""
          vgSql = "DELETE PP_TMAE_RETJUDICIAL WHERE "
          vgSql = vgSql & "num_retencion = '" & Lbl_NumRetencion.Caption & "' "
          Set vgRs = vgConexionBD.Execute(vgSql)
          
          Call flLimpiar
          Call flInicializaGrillaRetJud
          Call flCargaGrillaRetJud
          ''*Call flCargaGrillaBeneficiarios
          SSTab.Tab = 0
          
          vlSwSeleccionado = False
                                                    
          MsgBox "La Eliminación de la Retención Judicial Fue Realizada Correctamente", vbInformation, "Información"
          Exit Sub
'       Else
'           MsgBox "No puede Eliminar una Retención Judicial Suspendida", vbInformation, "Información"
'           Exit Sub
'       End If
       
    Else
        MsgBox "Debe Seleccionar la Retención Judicial que Desea Eliminar", vbInformation, "Información"
        Exit Sub
    End If

Else
    'MsgBox "No puede Modificar una Retención Judicial Suspendida", vbInformation, "Información"
    If Txt_PenPoliza.Enabled = False Then
       MsgBox "Sólo puede Consultar los Datos que se encuentran en Pantalla", vbInformation, "Información"
    End If
End If



Exit Sub
Err_CmdEliminarClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub cmd_grabar_Click()

On Error GoTo Err_CmdGrabarClick

'Si se encuentra habilitado, significa que puede ser modificado, de los contrario,
'indica que la Retención Judicial seleccionada ya se encuentra suspendida.
If Fra_AntecedentesRet.Enabled = True Then

    If vlSwSeleccionado = False Then
        Call flGenerarCorrelativoRet
        Lbl_NumRetencion.Caption = vlNumRetencion
    End If

    
'Valida Fecha de Inicio de Vigencia de la Retención Judicial

    If (Trim(Txt_FechaIniVig) = "") Then
       MsgBox "Debe ingresar una Fecha de Inicio de Vigencia", vbCritical, "Error de Datos"
       Txt_FechaIniVig.SetFocus
       Exit Sub
    End If
    If Not IsDate(Txt_FechaIniVig.Text) Then
       MsgBox "La Fecha Ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
       Txt_FechaIniVig.SetFocus
       Exit Sub
    End If
    'hqr 06/10/2007 Se quita validación para que pueda ingresar retenciones con vigencia futura (Pagos Diferidos)
'    If (CDate(Txt_FechaIniVig) > CDate(Date)) Then
'       MsgBox "La Fecha Ingresada es Mayor a la Fecha Actual", vbCritical, "Error de Datos"
'       Txt_FechaIniVig.SetFocus
'       Exit Sub
'    End If
    If (Year(Txt_FechaIniVig) < 1900) Then
       MsgBox "La Fecha Ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
       Txt_FechaIniVig.SetFocus
       Exit Sub
    End If
    
    Txt_FechaIniVig.Text = Format(CDate(Trim(Txt_FechaIniVig)), "yyyymmdd")
    Txt_FechaIniVig.Text = DateSerial(Mid((Txt_FechaIniVig.Text), 1, 4), Mid((Txt_FechaIniVig.Text), 5, 2), Mid((Txt_FechaIniVig.Text), 7, 2))
        
    If Not fgValidaVigenciaPoliza(Trim(Txt_PenPoliza.Text), Txt_FechaIniVig.Text) Then
       MsgBox " La Fecha Ingresada es Anterior a la Fecha de Vigencia de la Póliza ", vbInformation, "Información"
       Txt_FechaIniVig.SetFocus
       Exit Sub
    End If
    
    ''hqr 06/10/2007 Se descomenta validacion, para que no se permita ingresar Retenciones con fechas pasadas
    'If Not fgValidaPagoPension(Txt_FechaIniVig, Trim(Txt_PenPoliza), vlNumOrden) Then
    '   MsgBox " Ya se ha Realizado el Proceso de Cálculo de Pensión para ésta Fecha " & Chr(13) & _
    '          "                       Ingrese una Nueva Fecha                       ", vbCritical, "Operación Cancelada"
    '   Txt_FechaIniVig.SetFocus
    '   Exit Sub
    'End If

    Lbl_FechaEfecto.Caption = fgValidaFechaEfecto(Txt_FechaIniVig.Text, Trim(Txt_PenPoliza), vlNumOrden)
        
'Valida Fecha de Termino de Vigencia de la Retención Judicial
    'hqr 06/10/2007 se descomenta validacion de la  fecha de Termino
    If (Trim(Txt_FechaTerVig) = "") Then
       MsgBox "Debe ingresar una Fecha de Término de Vigencia", vbCritical, "Error de Datos"
       Txt_FechaTerVig.SetFocus
       Exit Sub
    End If
    If Not IsDate(Txt_FechaTerVig.Text) Then
       MsgBox "La Fecha Ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
       Txt_FechaTerVig.SetFocus
       Exit Sub
    End If
'    If (CDate(Txt_FechaTerVig) > CDate(Date)) Then
'       MsgBox "La Fecha Ingresada es Mayor a la Fecha Actual", vbCritical, "Error de Datos"
'       Txt_FechaTerVig.SetFocus
'       Exit Sub
'    End If
'    If (Year(Txt_FechaTerVig) < 1900) Then
'       MsgBox "La Fecha Ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
'       Txt_FechaTerVig.SetFocus
'       Exit Sub
'    End If

    Txt_FechaTerVig.Text = Format(CDate(Trim(Txt_FechaTerVig)), "yyyymmdd")
    Txt_FechaTerVig.Text = DateSerial(Mid((Txt_FechaTerVig.Text), 1, 4), Mid((Txt_FechaTerVig.Text), 5, 2), Mid((Txt_FechaTerVig.Text), 7, 2))
        
    'Valida que la Fecha de Termino sea Mayor a la Fecha de Inicio de Vigencia
    If Format(CDate(Trim(Txt_FechaTerVig)), "yyyymmdd") < Format(CDate(Trim(Txt_FechaIniVig)), "yyyymmdd") Then
        MsgBox "La Fecha Ingresada es Mayor a la Fecha de Inicio de Vigencia.", vbCritical, "Error de Datos"
        If Txt_FechaTerVig.Enabled Then
            Txt_FechaTerVig.SetFocus
        End If
        Exit Sub
    End If
'Valida Fecha de Recepción de la Retención Judicial
    
    If (Trim(Txt_FechaRecepcion) = "") Then
       MsgBox "Debe ingresar una Fecha de Recepción", vbCritical, "Error de Datos"
       Txt_FechaRecepcion.SetFocus
       Exit Sub
    End If
    If Not IsDate(Txt_FechaRecepcion.Text) Then
       MsgBox "La Fecha Ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
       Txt_FechaRecepcion.SetFocus
       Exit Sub
    End If
    If (CDate(Txt_FechaRecepcion) > CDate(Date)) Then
       MsgBox "La Fecha Ingresada es Mayor a la Fecha Actual", vbCritical, "Error de Datos"
       Txt_FechaRecepcion.SetFocus
       Exit Sub
    End If
    If (Year(CDate(Txt_FechaRecepcion)) < 1900) Then
       MsgBox "La Fecha Ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
       Txt_FechaRecepcion.SetFocus
       Exit Sub
    End If
    
    Txt_FechaRecepcion.Text = Format(CDate(Trim(Txt_FechaRecepcion)), "yyyymmdd")
    Txt_FechaRecepcion.Text = DateSerial(Mid((Txt_FechaRecepcion.Text), 1, 4), Mid((Txt_FechaRecepcion.Text), 5, 2), Mid((Txt_FechaRecepcion.Text), 7, 2))
        
        
    If Txt_MontoRetenido.Enabled = True Then
      
       'Valida el Monto de Retención Ingresado

        If Txt_MontoRetenido.Text = "" Then
           MsgBox "Debe Ingresar Monto de la Retención.", vbCritical, "Error de Datos"
           Txt_MontoRetenido.SetFocus
           Exit Sub
        Else
            Txt_MontoRetenido.Text = Format(Txt_MontoRetenido.Text, "###,###,##0.00")
        End If
'CMV-I 20050311
'        If CDbl(Txt_MontoRetenido.Text) = 0 Then
'           MsgBox "Debe Ingresar un Valor Mayor que Cero para Monto Retención.", vbCritical, "Error de Datos"
'           Txt_MontoRetenido.SetFocus
'           Exit Sub
'        End If
'CMV-F 20050311
        If CDbl(Txt_MontoRetenido.Text) > clTopeMtoRet Then
           MsgBox "El Valor Ingresado en Monto Retención Excede el Máximo Permitido de " & Format(clTopeMtoRet, "###,###,##0.00") & "  .", vbExclamation, "Información"
           Txt_MontoRetenido.Text = Format(clTopeMtoRet, "###,###,##0.00")
           Txt_MontoRetenido.SetFocus
           Exit Sub
        End If
    
        
        Select Case Mid(Cmb_ModRetencion.Text, 1, 5)
        Case "PRCIM"
            If CDbl(Txt_MontoRetenido.Text) > 60 Then
                MsgBox "El Valor Ingresado exede el 60. Ingresar un valor menor del 60%", vbExclamation, "Información"
                Txt_MontoRetenido.Text = 60
                Txt_MontoRetenido.SetFocus
            Exit Sub
        End If
        Case "MTOIM"
            If CDbl(Txt_MontoRetenido.Text) > lblMontoRetAct Then
                MsgBox "El Valor Ingresado en Monto Retención Excede el Máximo Permitido de " & Format(lblMontoRetAct, "###,###,##0.00") & "  .", vbExclamation, "Información"
                Txt_MontoRetenido.Text = Format(lblMontoRetAct, "###,###,##0.00")
                Txt_MontoRetenido.SetFocus
            Exit Sub
            End If
        End Select
        
        
       
        
    End If
        
    If Txt_Juzgado.Enabled = True Then
       'Valida la Descripción de Juzgado Ingresada

        Txt_Juzgado.Text = Trim(UCase(Txt_Juzgado.Text))
        If Txt_Juzgado.Text = "" Then
           MsgBox "Debe Ingresar Descripción del Juzgado", vbCritical, "Error de Datos"
           Txt_Juzgado.SetFocus
           Exit Sub
        End If
        
    End If
    
'Valida el Rut del Receptor Ingresado

    If (Trim(Txt_NumIdent.Text)) = "" Then
       MsgBox "Debe ingresar el Número de Identificación de Receptor.", vbCritical, "Error de Datos"
       Txt_NumIdent.SetFocus
       Exit Sub
    Else
        Txt_NumIdent = Trim(UCase(Txt_NumIdent))
        ''*Txt_DigitoRec = UCase(Trim(Txt_DigitoRec))
        'Txt_DigitoRec.SetFocus
    End If
    
    'Valida la Existencia de otra retención ingresada con el mismo receptor
    Call flValidaReceptor
    If vlRutRecAux <> "" Then
       If vlNumRetAux <> CLng(Lbl_NumRetencion.Caption) Then
'       If vlNumRetAux <> vlNumRetencion Then
          MsgBox " El Receptor con Identificación: " & Cmb_NumIdent & " - " & Txt_NumIdent & " " & Chr(13) & _
         "  Se encuentra Registrado en otra Retención Judicial Activa.", vbExclamation, "Información"
         Exit Sub
       End If
    End If
    
'Valida el Digito Verificador del Rut de Receptor ingresado

''    Txt_DigitoRec = Trim(UCase(Txt_DigitoRec))
''    If Txt_DigitoRec.Text = "" Then
''        MsgBox "Debe ingresar Dígito Verificador del Rut.", vbCritical, "Error de Datos"
''        Txt_DigitoRec.SetFocus
''        Exit Sub
''    End If
''    If Not ValiRut(Txt_RutRec.Text, Txt_DigitoRec.Text) Then
''        MsgBox "El Dígito Verificador del Rut ingresado es incorrecto.", vbCritical, "Error de Datos"
''        Txt_DigitoRec.SetFocus
''        Exit Sub
''    End If
    
'Valida el Nombre de Receptor

    Txt_NombreRec.Text = Trim(UCase(Txt_NombreRec.Text))
    If Txt_NombreRec.Text = "" Then
       MsgBox "Debe Ingresar Nombre.", vbCritical, "Error de Datos"
       Txt_NombreRec.SetFocus
       Exit Sub
    End If

'Valida el Apellido Paterno del Receptor

    Txt_PaternoRec.Text = Trim(UCase(Txt_PaternoRec.Text))
    If Txt_PaternoRec.Text = "" Then
       MsgBox "Debe Ingresar Apellido Paterno.", vbCritical, "Error de Datos"
       Txt_PaternoRec.SetFocus
       Exit Sub
    End If
    
'    'Valida el Apellido Materno del Receptor
'    Txt_MaternoRec.Text = Trim(UCase(Txt_MaternoRec.Text))
'    If Txt_MaternoRec.Text = "" Then
'       MsgBox "Debe Ingresar Apellido Materno.", vbCritical, "Error de Datos"
'       Txt_MaternoRec.SetFocus
'       Exit Sub
'    End If
    
'Valida el ingreso de la Dirección del Receptor

    Txt_DireccionRec = Trim(UCase(Txt_DireccionRec))
    If Txt_DireccionRec.Text = "" Then
        MsgBox "Debe Ingresar Dirección.", vbCritical, "Error de Datos"
        Txt_DireccionRec.SetFocus
        Exit Sub
    End If
    
    'Asignación de fecha por defecto a fecha de término de vigencia

    If Txt_FecSuspension.Text = "" Then
'       Txt_FechaTerVig.Text = clFechaTopeTer
    
'       Txt_FechaIniVig.Text = Format(CDate(Trim(Txt_FechaTerVig)), "yyyymmdd")
       Txt_FecSuspension.Text = DateSerial(Mid((clFechaTopeTer), 1, 4), Mid((clFechaTopeTer), 5, 2), Mid((clFechaTopeTer), 7, 2))
       Txt_FechaTerVig.Text = Txt_FecSuspension.Text
    Else
'Valida si la fecha de suspensión es mayor o menor que la fecha del último proceso de cálculo de pensiones.
       ' If Not fgValidaPagoPension(Trim(Txt_FecSuspension.Text), Trim(Txt_PenPoliza.Text), vlNumOrden) Then
       '    MsgBox "Ya se ha Realizado el Proceso de Cálculo de Pensión para la Fecha Ingresada", vbCritical, "Operación Cancelada"
       '    Exit Sub
       ' 'Else
       '     'MsgBox "INGRESAR fecha ingresada por el usuario", vbCritical, "Error de Datos"
       ' End If
    End If
    
    
'Validar Ingreso de Datos para Pago, según Via de Pago Seleccionada
    vlNumero = InStr(Cmb_ViaPago.Text, "-")
    vlOpcion = Trim(Mid(Cmb_ViaPago.Text, 1, vlNumero - 1))
'Opción 01 = Via de Pago CAJA
           
'Opción 00 = Via de Pago - Sin Información
    If vlOpcion <> "00" Then
    
       If vlOpcion = "01" Then
          vlNumero = InStr(Cmb_Sucursal.Text, "-")
          If Trim(Mid(Cmb_Sucursal.Text, 1, vlNumero - 1)) = "0000" Then
             MsgBox "Debe Seleccionar Sucursal de Pago", vbCritical, "Error de Datos"
             Cmb_Sucursal.SetFocus
             Exit Sub
          End If
       Else
'Opción 02 = Via de Pago - Deposito en Cuenta
'Opción 03 = Via de Pago - Convenio
           If (vlOpcion = "02") Or (vlOpcion = "03") Then
              vlNumero = InStr(Cmb_TipoCta.Text, "-")
              If Trim(Mid(Cmb_TipoCta.Text, 1, vlNumero - 1)) = "00" Then
                 MsgBox "Debe Seleccionar Tipo de Cuenta para Pago", vbCritical, "Error de Datos"
                 Cmb_TipoCta.SetFocus
                 Exit Sub
              End If
              vlNumero = InStr(Cmb_Banco.Text, "-")
              If Trim(Mid(Cmb_Banco.Text, 1, vlNumero - 1)) = "00" Then
                 MsgBox "Debe Seleccionar Banco para Pago", vbCritical, "Error de Datos"
                 Cmb_Banco.SetFocus
                 Exit Sub
              End If
              If Txt_NumCuenta.Text = "" Then
                 MsgBox "Debe Ingresar Número de Cuenta para Pago", vbCritical, "Error de Datos"
                 Txt_NumCuenta.SetFocus
                 Exit Sub
              End If
           Else
'Via de Pago <> = Via de Pago <>
               'ACTIVAR TODO e ingresar todo sin validar
               Cmb_Sucursal.Enabled = True
               Cmb_TipoCta.Enabled = True
               Cmb_Banco.Enabled = True
               Txt_NumCuenta.Enabled = True
               Cmb_Sucursal.ListIndex = 0
               Cmb_TipoCta.ListIndex = 0
               Cmb_Banco.ListIndex = 0
               Txt_NumCuenta.Text = ""
           End If
       End If
    Else
        MsgBox "Debe Seleccionar Forma de Pago.", vbCritical, "Error de Datos"
        If (Cmb_ViaPago.Enabled = True) Then
            Cmb_ViaPago.SetFocus
        End If
        Exit Sub
    End If
    
    'Validar que no se encuentre ingresada otra retención por
    'asignación familiar activa.
    vlNumero = InStr(Cmb_TipoRetencion.Text, "-")
    vlOpcion = Trim(Mid(Cmb_TipoRetencion.Text, 1, vlNumero - 1))
    If vlOpcion = clTipoAsigFam Then
       Call flValidaRetAsigFam
       If vlNumRetAux <> 0 Then
          If vlNumRetAux <> CLng(Lbl_NumRetencion.Caption) Then
          'If vlNumRetAux <> vlNumRetencion Then
             MsgBox " La Retención por Asignación Familiar de Conyuge " & Chr(13) & _
             "  Se encuentra Ingresada en otra Retención Judicial Activa.", vbExclamation, "Información"
             Exit Sub
          End If
       End If
    End If
    
''    'Validar que las cargas seleccionadas no se encuentren ingresadas
''    'en otra retencion judicial activa.
''    If SSTab.TabEnabled(2) = True Then
''       Call flValidaCargas
''       If vlCargasRetenidas <> "" Then
''          MsgBox " Las Cargas con Nº de Orden: '" & vlCargasRetenidas & "' " & Chr(13) & _
''                 "  Se encuentran Ingresadas en otra Retención Judicial Activa.", vbExclamation, "Información"
''          SSTab.Tab = 2
''          Exit Sub
''       End If
''    End If

    If Trim(Lbl_NumRetencion.Caption) <> "" And vlSwSeleccionado = True Then
     
            '''    vlNumero = InStr(Cmb_TipoRetencion.Text, "-")
            '''    vlOpcion = Trim(Mid(Cmb_TipoRetencion.Text, 1, vlNumero - 1))
            '''
            '''    vgSql = ""
            '''    vgSql = "SELECT num_poliza "
            '''    vgSql = vgSql & "FROM PP_TMAE_RETJUDICIAL WHERE "
            '''    vgSql = vgSql & " num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
            '''    vgSql = vgSql & " num_endoso = " & Trim(vlNumEndoso) & " AND "
            '''    vgSql = vgSql & " num_orden = " & Trim(vlNumOrden) & " AND "
            '''    vgSql = vgSql & " cod_tipret = " & Trim(vlOpcion) & " AND "
            '''    vgSql = vgSql & " rut_receptor = '" & Str(Trim(Txt_RutRec.Text)) & "' "
            
       vgSql = ""
       vgSql = "SELECT num_poliza "
       vgSql = vgSql & "FROM PP_TMAE_RETJUDICIAL WHERE "
       vgSql = vgSql & " num_retencion = '" & Trim(Lbl_NumRetencion.Caption) & "' "
       
       Set vgRs = vgConexionBD.Execute(vgSql)
       If Not vgRs.EOF Then

          vgRes = MsgBox("¿ Está seguro que desea Modificar los Datos ?", 4 + 32 + 256, "Operación de Actualización")
          If vgRes <> 6 Then
             vgRs.Close
             Screen.MousePointer = 0
             Exit Sub
          End If
       
          vlNumRetencion = Lbl_NumRetencion.Caption
       
''          If SSTab.TabEnabled(2) = True Then
''             Call flEliminarDetCargas
''             Call flAgregarDetCargas
''          End If
       
          vlGlsUsuarioModi = vgUsuario
          vlFecModi = Format(Date, "yyyymmdd")
          vlHorModi = Format(Time, "hhmmss")
          
            vlNumero = InStr(Cmb_NumIdent.Text, "-")
            vlTipoIdenRec = Trim(Mid(Cmb_NumIdent.Text, 1, vlNumero - 1))
        
          vlNumero = InStr(Cmb_TipoRetencion.Text, "-")
          vlTipoRet = Trim(Mid(Cmb_TipoRetencion.Text, 1, vlNumero - 1))
       
          vlNumero = InStr(Cmb_ModRetencion.Text, "-")
          vlModRet = Trim(Mid(Cmb_ModRetencion.Text, 1, vlNumero - 1))
        
          vgSql = ""
          vgSql = " UPDATE PP_TMAE_RETJUDICIAL SET "
          vgSql = vgSql & " fec_resdoc = " & Format(CDate(Trim(Txt_FechaRecepcion.Text)), "yyyymmdd") & ", "
          vgSql = vgSql & " fec_iniret = " & Format(CDate(Trim(Txt_FechaIniVig.Text)), "yyyymmdd") & ", "
          
          vgSql = vgSql & " fec_terret = '" & Format(CDate(Trim(Txt_FechaTerVig.Text)), "yyyymmdd") & "', "
          vgSql = vgSql & " cod_tipret = '" & Trim(vlTipoRet) & "', "
          vgSql = vgSql & " cod_modret = '" & Trim(vlModRet) & "', "
          vgSql = vgSql & " mto_ret = " & str(Trim(Txt_MontoRetenido.Text)) & ", "
          
          If (Trim(Txt_Juzgado.Text) = "") Then
             vgSql = vgSql & " gls_nomjuzgado = NULL , "
          Else
              vgSql = vgSql & " gls_nomjuzgado = '" & Trim(Txt_Juzgado.Text) & "', "
          End If
          
          vgSql = vgSql & " cod_tipoidenreceptor = " & vlTipoIdenRec & ", "
          vgSql = vgSql & " num_idenreceptor = '" & Trim(Txt_NumIdent.Text) & "', "
          vgSql = vgSql & " gls_nomreceptor = '" & Trim(Txt_NombreRec.Text) & "', "
           If (Trim(Txt_Juzgado.Text) = "") Then
             vgSql = vgSql & " gls_nomsegjuzgado = NULL , "
          Else
             vgSql = vgSql & " gls_nomsegreceptor = '" & Trim(Txt_NombreSegRec.Text) & "', "
          End If
          vgSql = vgSql & " gls_patreceptor = '" & Trim(Txt_PaternoRec.Text) & "', "
          vgSql = vgSql & " gls_matreceptor = '" & Trim(Txt_MaternoRec.Text) & "', "
          vgSql = vgSql & " gls_dirreceptor = '" & Trim(Txt_DireccionRec.Text) & "', "
          vgSql = vgSql & " cod_direccion = " & vlCodDir & ", "
               
          If (Trim(Txt_TelefonoRec.Text) = "") Then
             vgSql = vgSql & " gls_fonoreceptor = NULL , "
          Else
              vgSql = vgSql & " gls_fonoreceptor = '" & Trim(Txt_TelefonoRec.Text) & "', "
          End If
          If (Trim(Txt_CorreoRec.Text) = "") Then
             vgSql = vgSql & " gls_emailreceptor = NULL, "
          Else
              vgSql = vgSql & " gls_emailreceptor = '" & Trim(Txt_CorreoRec.Text) & "', "
          End If
       
          vlNumero = InStr(Cmb_ViaPago.Text, "-")
          vlViaPago = Trim(Mid(Cmb_ViaPago.Text, 1, vlNumero - 1))
         
          vgSql = vgSql & " cod_viapago = '" & Trim(vlViaPago) & "', "
       
          vlNumero = InStr(Cmb_TipoCta.Text, "-")
          vgSql = vgSql & " cod_tipcuenta = '" & Trim(Mid(Cmb_TipoCta.Text, 1, vlNumero - 1)) & "', "
      
          vlNumero = InStr(Cmb_Banco.Text, "-")
          vgSql = vgSql & " cod_banco = '" & Trim(Mid(Cmb_Banco.Text, 1, vlNumero - 1)) & "', "
       
          If Txt_NumCuenta.Enabled = True Then
             vgSql = vgSql & " num_cuenta = '" & Trim(Txt_NumCuenta.Text) & "', "
          Else
              vgSql = vgSql & " num_cuenta = NULL, "
          End If
        
          vlNumero = InStr(Cmb_Sucursal.Text, "-")
          vgSql = vgSql & " cod_sucursal = '" & Trim(Mid(Cmb_Sucursal.Text, 1, vlNumero - 1)) & "', "
          
          vgSql = vgSql & " cod_monedaretmax = '" & Trim(Lbl_Moneda.Caption) & "', "
          vgSql = vgSql & " mto_retmax = " & str(Trim(Txt_MtoMaxRet.Text)) & ", "
          
          vgSql = vgSql & "fec_efecto = '" & Trim(Format(Lbl_FechaEfecto, "yyyymmdd")) & "', "
       
          vgSql = vgSql & " cod_usuariomodi = '" & vlGlsUsuarioModi & "', "
          vgSql = vgSql & " fec_modi = '" & vlFecModi & "', "
          vgSql = vgSql & " hor_modi = '" & vlHorModi & "' "

          vgSql = vgSql & " WHERE "
          vgSql = vgSql & "num_retencion = '" & Trim(Lbl_NumRetencion.Caption) & "' "
       
          vgConexionBD.Execute vgSql
       
          MsgBox "Los Datos han sido actualizados Satisfactoriamente", vbInformation, "Información"

       End If
    Else
        
        Call flGenerarCorrelativoRet
        'Lbl_NumRetencion.Caption = vlNumRetencion
                vlGlsUsuarioCrea = vgUsuario
        vlFecCrea = Format(Date, "yyyymmdd")
        vlHorCrea = Format(Time, "hhmmss")
        

        vlGlsUsuarioModi = Null
        vlFecModi = Null
        vlHorModi = Null
        
        vlNumero = InStr(Cmb_NumIdent.Text, "-")
        vlTipoIdenRec = Trim(Mid(Cmb_NumIdent.Text, 1, vlNumero - 1))
        
        vlNumero = InStr(Cmb_TipoRetencion.Text, "-")
        vlTipoRet = Trim(Mid(Cmb_TipoRetencion.Text, 1, vlNumero - 1))
       
        vlNumero = InStr(Cmb_ModRetencion.Text, "-")
        vlModRet = Trim(Mid(Cmb_ModRetencion.Text, 1, vlNumero - 1))
        
        vgSql = ""
        vgSql = "INSERT INTO PP_TMAE_RETJUDICIAL "
        vgSql = vgSql & "(num_retencion,num_poliza,num_endoso,num_orden,fec_resdoc,fec_iniret, "
'        vgSql = vgSql & " fec_terret, "
        vgSql = vgSql & " fec_terret,cod_tipret,cod_modret,mto_ret,gls_nomjuzgado, "
        vgSql = vgSql & " cod_tipoidenreceptor,num_idenreceptor,gls_nomreceptor,gls_nomsegreceptor,gls_patreceptor,gls_matreceptor,"
        vgSql = vgSql & " gls_dirreceptor,cod_direccion,gls_fonoreceptor,gls_emailreceptor,cod_viapago, "
        vgSql = vgSql & " cod_tipcuenta,cod_banco,num_cuenta,cod_sucursal,cod_monedaretmax,mto_retmax,fec_efecto, "
        vgSql = vgSql & " cod_usuariocrea,fec_crea,hor_crea,cod_usuariomodi,fec_modi,hor_modi "
        vgSql = vgSql & " ) VALUES ( "
        vgSql = vgSql & " " & vlNumRetencion & ", "
        vgSql = vgSql & "'" & Trim(Txt_PenPoliza) & "' , "
        vgSql = vgSql & " " & vlNumEndoso & ", "
        vgSql = vgSql & " " & vlNumOrden & ", "
        vgSql = vgSql & "'" & Format(CDate(Trim(Txt_FechaRecepcion.Text)), "yyyymmdd") & "', "
        vgSql = vgSql & "'" & Format(CDate(Trim(Txt_FechaIniVig.Text)), "yyyymmdd") & "', "
        
        vgSql = vgSql & "'" & Format(CDate(Trim(Txt_FechaTerVig.Text)), "yyyymmdd") & "', "
        vgSql = vgSql & "'" & vlTipoRet & "', "
        vgSql = vgSql & "'" & vlModRet & "', "
        vgSql = vgSql & " " & (str(Txt_MontoRetenido.Text)) & ", "
        
        If (Trim(Txt_Juzgado.Text) = "") Then
           vgSql = vgSql & " NULL, "
        Else
            vgSql = vgSql & " '" & Trim(Txt_Juzgado.Text) & "', "
        End If
                  
        vgSql = vgSql & " " & vlTipoIdenRec & ", "
        vgSql = vgSql & "'" & Trim(Txt_NumIdent.Text) & "', "
        vgSql = vgSql & "'" & Trim(Txt_NombreRec.Text) & "', "
        vgSql = vgSql & "'" & Trim(Txt_NombreSegRec.Text) & "', "
        vgSql = vgSql & "'" & Trim(Txt_PaternoRec.Text) & "', "
        vgSql = vgSql & "'" & Trim(Txt_MaternoRec.Text) & "', "
        vgSql = vgSql & "'" & Trim(Txt_DireccionRec.Text) & "', "
        vgSql = vgSql & " " & vlCodDir & ", "
    
        If (Trim(Txt_TelefonoRec.Text) = "") Then
           vgSql = vgSql & " NULL, "
        Else
            vgSql = vgSql & " '" & Txt_TelefonoRec.Text & "', "
        End If
        If (Trim(Txt_CorreoRec.Text) = "") Then
           vgSql = vgSql & " NULL, "
        Else
            vgSql = vgSql & " '" & Txt_CorreoRec.Text & "', "
        End If
        
        vlNumero = InStr(Cmb_ViaPago.Text, "-")
        vlViaPago = Trim(Mid(Cmb_ViaPago.Text, 1, vlNumero - 1))
        vgSql = vgSql & " '" & Trim(Mid(Cmb_ViaPago.Text, 1, vlNumero - 1)) & "', "
   
        vlNumero = InStr(Cmb_TipoCta.Text, "-")
        vgSql = vgSql & " '" & Trim(Mid(Cmb_TipoCta.Text, 1, vlNumero - 1)) & "', "
        
        vlNumero = InStr(Cmb_Banco.Text, "-")
        vgSql = vgSql & " '" & Trim(Mid(Cmb_Banco.Text, 1, vlNumero - 1)) & "', "

        If Txt_NumCuenta.Enabled = True Then
           vgSql = vgSql & " '" & Trim(Txt_NumCuenta.Text) & "', "
        Else
            vgSql = vgSql & " NULL, "
        End If

        vlNumero = InStr(Cmb_Sucursal.Text, "-")
        vgSql = vgSql & "'" & Trim(Mid(Cmb_Sucursal.Text, 1, vlNumero - 1)) & "', "
        
        vgSql = vgSql & "'" & Trim(Lbl_Moneda.Caption) & "', "
        vgSql = vgSql & " " & (str(Txt_MtoMaxRet.Text)) & ", "

        vgSql = vgSql & "'" & Trim(Format(Lbl_FechaEfecto, "yyyymmdd")) & "', "

        vgSql = vgSql & "'" & vlGlsUsuarioCrea & "', "
        vgSql = vgSql & "'" & vlFecCrea & "', "
        vgSql = vgSql & "'" & vlHorCrea & "', "
        If (Not IsNull(vlGlsUsuarioModi)) Then
            vgSql = vgSql & "'" & vlGlsUsuarioModi & "', "
            vgSql = vgSql & "'" & vlFecModi & "', "
            vgSql = vgSql & "'" & vlHorModi & "' "
        Else
            vgSql = vgSql & "NULL, "
            vgSql = vgSql & "NULL, "
            vgSql = vgSql & "NULL "
        End If
        vgSql = vgSql & ") "

        vgConexionBD.Execute vgSql
        
''        If SSTab.TabEnabled(2) = True Then
''           Call flAgregarDetCargas
''        End If

        
        Lbl_NumRetencion.Caption = vlNumRetencion

        MsgBox "Los Datos han sido Ingresados Satisfactoriamente.", vbInformation, "Información"
       End If
'    End If
'CMV 20041102
''''    Call flLimpiar
'CMV 20041102
    Call flCargaGrillaRetJud
    Call flCargaGrillaBeneficiarios
        
    SSTab.Tab = 0
    
    If Fra_AntecedentesRet.Enabled = True Then
       Txt_FechaIniVig.SetFocus
    End If
        
Else
    'MsgBox "No puede Modificar una Retención Judicial Suspendida", vbInformation, "Información"
    If Txt_PenPoliza.Enabled = False Then
       MsgBox "Sólo puede Consultar los Datos que se encuentran en Pantalla", vbInformation, "Información"
    End If
End If
    
Exit Sub
Err_CmdGrabarClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
  
End Sub

Private Sub Cmd_Imprimir_Click()

On Error GoTo Err_CmdImprimir

  'Validar que se encuentran los datos que identifican al Pensionado
   If (Trim(Txt_PenPoliza) = "") Then
       MsgBox "Debe ingresar el Nº de Póliza.", vbCritical, "Error de Datos"
       Exit Sub
   End If
   If (Trim(Txt_PenNumIdent) = "") Then
       MsgBox "Debe ingresar el Rut del Pensionado.", vbCritical, "Error de Datos"
       Exit Sub
   End If
''   If (Trim(Txt_PenDigito) = "") Then
''       MsgBox "Debe ingresar el Dígito del Rut del Pensionado.", vbCritical, "Error de Datos"
''       Exit Sub
''   End If
''
   Screen.MousePointer = 11
   
   vlArchivo = strRpt & "PP_Rpt_RetJudicial.rpt"   '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
      MsgBox "Archivo de Reporte de Retención Judicial no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
      Screen.MousePointer = 0
      Exit Sub
   End If
   
'   vgQuery = "{PP_TMAE_RETJUDICIAL.NUM_POLIZA} = '" & Trim(Txt_PenPoliza.Text) & "' "
   

''   vgSql = ""
''   vgSql = "SELECT rut_ben,dgv_ben,gls_nomben,gls_patben,gls_matben "
''   vgSql = vgSql & "FROM PP_TMAE_BEN "
''   vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
''   vgSql = vgSql & "rut_ben = '" & Trim(Str(Txt_PenRut.Text)) & "' "
''   vgSql = vgSql & "ORDER by num_orden "
''   Set vgRs = vgConexionBD.Execute(vgSql)
''   If Not vgRs.EOF Then
''      vlRutPen = (Trim(vgRs!rut_ben)) & " - " & (Trim(vgRs!dgv_ben))
''      vlNombrePen = Trim(vgRs!gls_nomben) + "  " + Trim(vgRs!gls_patben) + " " + Trim(vgRs!gls_matben)
''   End If


   
'   Rpt_General.Reset
'   Rpt_General.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
'   Rpt_General.Connect = vgRutaDataBase       ' o App.Path & "\Nestle.mdb"
'   Rpt_General.SelectionFormula = vgQuery
'   Rpt_General.Formulas(0) = "TipoIdenPensionado = '" & (Trim(Cmb_PenNumIdent.Text)) & "' "
'   Rpt_General.Formulas(1) = "NumIdenPensionado = '" & (Trim(Txt_PenNumIdent.Text)) & "' "
'   Rpt_General.Formulas(2) = "NombrePensionado = '" & Trim(Lbl_PenNombre.Caption) & "' "
'   'Rpt_General.Formulas(2) = ""
'
'   Rpt_General.Formulas(3) = "NombreCompania = '" & vgNombreCompania & "'"
'   Rpt_General.Formulas(4) = "NombreSistema= '" & vgNombreSistema & "'"
'   Rpt_General.Formulas(5) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
'
'   Rpt_General.SubreportToChange = ""
'   Rpt_General.Destination = crptToWindow
'   Rpt_General.WindowState = crptMaximized
'   Rpt_General.WindowTitle = "Informe de Retención Judicial por Pensionado"
''   Rpt_General.SubreportToChange = "PP_Rpt_RetJudicialSUB.rpt"
'''   Rpt_General.SelectionFormula = ""
'''   Rpt_General.Connect = vgRutaDataBase
'   Rpt_General.Action = 1
'   Screen.MousePointer = 0
     
    'Roger 11/03/2014
   'Dim cadena As String
   Dim objRep As New ClsReporte
   Dim vlFechaPago As String
   Dim rs As New ADODB.Recordset
   
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    
    vgQuery = "SELECT a.*, b.gls_elemento as cod_tp, c.gls_elemento as cod_mod, d.gls_tipoiden as cod_tiprec, "
    vgQuery = vgQuery & "e.gls_nomben, e.gls_nomsegben, e.gls_patben, e.gls_matben, e.num_idenben, f.gls_tipoiden as cod_tippen "
    vgQuery = vgQuery & "FROM PP_TMAE_RETJUDICIAL a "
    vgQuery = vgQuery & "join ma_tpar_tabcod b on a.cod_tipret=b.cod_elemento and b.cod_tabla='TRT' "
    vgQuery = vgQuery & "join ma_tpar_tabcod c on a.cod_modret=c.cod_elemento and c.cod_tabla='MPR' "
    vgQuery = vgQuery & "join ma_tpar_tipoiden d on a.cod_tipoidenreceptor=d.cod_tipoiden "
    vgQuery = vgQuery & "join pp_tmae_ben e on a.num_poliza=e.num_poliza and a.num_endoso=e.num_endoso and a.num_orden=e.num_orden "
    vgQuery = vgQuery & "join ma_tpar_tipoiden f on e.cod_tipoidenben=f.cod_tipoiden "
    vgQuery = vgQuery & "where a.num_poliza='" & Trim(Txt_PenPoliza) & "' order by a.num_retencion"
    rs.Open vgQuery, vgConexionBD, adOpenForwardOnly, adLockReadOnly

    Dim LNGa As Long
    LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PP_Rpt_RetJudicial.rpt"), ".RPT", ".TTX"), 1)
    
    If objRep.CargaReporte(strRpt & "", "PP_Rpt_RetJudicial.rpt", "Informe de Retención Judicial por Pensionado", rs, True, _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("NombreSistema", vgNombreSistema), _
                            ArrFormulas("NombreSubSistema", vgNombreSubSistema)) = False Then
                            
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
    End If
   
Exit Sub
Err_CmdImprimir:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub

Private Sub Cmd_Limpiar_Click()

''    Txt_PenPoliza.Text = ""
''    Txt_PenNumIdent.Text = ""
''    Txt_PenDigito.Text = ""
''    Lbl_End.Caption = ""
''    Lbl_PenNombre.Caption = ""
    
'    vlSw = True
    
    Call flLimpiar
    
    vlSw = False
    
    vlSwSeleccionado = False

    ''Call flDeshabilitarIngreso
    SSTab.Tab = 0

    ''Txt_PenPoliza.SetFocus

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

    Frm_RetJudicial.Top = 0
    Frm_RetJudicial.Left = 0
    
    ''Call fgComboComuna(Cmb_ComunaRec)
    'Call fgComboSucursal(Cmb_Sucursal)
    Call fgComboSucursal(Cmb_Sucursal, "S")
    
    vlSw = True
    
    fgComboGeneral vgCodTabla_TipRetJud, Cmb_TipoRetencion
    fgComboGeneral vgCodTabla_ModPagoRetJud, Cmb_ModRetencion
    fgComboGeneral vgCodTabla_ViaPago, Cmb_ViaPago
    fgComboGeneral vgCodTabla_TipCta, Cmb_TipoCta
    fgComboGeneral vgCodTabla_Bco, Cmb_Banco
    fgComboTipoIdentificacion Cmb_PenNumIdent
    fgComboTipoIdentificacion Cmb_NumIdent
    
'    Call flInicializaGrillaRetJud
'    Call flInicializaGrillaCargas
'    Call flInicializaGrillaBeneficiarios
    
    vlSw = False
    vlSwSeleccionado = False
    Call flLimpiar
    
    Call Cmb_ViaPago_Click
    
     'Asignación de fecha por defecto a fecha de término de vigencia

    If Txt_FecSuspension.Text = "" Then
'       Txt_FechaTerVig.Text = clFechaTopeTer
    
'       Txt_FechaIniVig.Text = Format(CDate(Trim(Txt_FechaTerVig)), "yyyymmdd")
       Txt_FecSuspension.Text = DateSerial(Mid((clFechaTopeTer), 1, 4), Mid((clFechaTopeTer), 5, 2), Mid((clFechaTopeTer), 7, 2))
       Txt_FechaTerVig.Text = Txt_FecSuspension.Text
    End If
    
    
    
    
    
    Call flDeshabilitarIngreso
    
    SSTab.Tab = 0
    
    Lbl_FechaEfecto = Lbl_FechaEfecto
    Lbl_FechaEfecto = flCalculaFechaEfecto(Lbl_FechaEfecto)
  
'    Call flIniciaForm
            
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Msf_GrillaRetJud_DblClick()

    On Error GoTo Err_GrillaRetJudClick
    
    ''Msf_GrillaBenef.Col = 0
    If (Msf_GrillaRetJud.Text = "") Or (Msf_GrillaRetJud.row = 0) Then
        MsgBox "No existen Detalles", vbExclamation, "Información"
        Exit Sub
    Else
        Msf_GrillaRetJud.Col = 10
        Lbl_NumRetencion.Caption = Msf_GrillaRetJud.Text
        
        vlSwSeleccionado = True
        
        Call flMostrarDatosRetencion
        
        SSTab.Tab = 0
        
        If Fra_AntecedentesRet.Enabled = True Then
           Txt_FechaIniVig.SetFocus
        End If
                
    End If

Exit Sub
Err_GrillaRetJudClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub

Private Sub Txt_correorec_KeyPress(KeyAscii As Integer)
On Error GoTo Err_TxtCorreoRecKeyPress

    If KeyAscii = 13 Then
        Txt_CorreoRec = Trim(Txt_CorreoRec)
        Cmb_ViaPago.SetFocus
    End If
    
Exit Sub
Err_TxtCorreoRecKeyPress:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Txt_correorec_LostFocus()
    Txt_CorreoRec = Trim(Txt_CorreoRec)
End Sub

Private Sub Txt_FechaIniVig_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        
        If (Trim(Txt_FechaIniVig) = "") Then
           MsgBox "Debe ingresar una Fecha de Inicio de Vigencia", vbCritical, "Error de Datos"
           Txt_FechaIniVig.SetFocus
           Exit Sub
        End If
        If Not IsDate(Txt_FechaIniVig.Text) Then
           MsgBox "La Fecha Ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
           Txt_FechaIniVig.SetFocus
           Exit Sub
        End If
        'hqr 06/10/2007 Se comenta para que permita ingresar Retenciones Futuras (Polizas Diferidas)
''        If (CDate(Txt_FechaIniVig) > CDate(Date)) Then
''           MsgBox "La Fecha Ingresada es Mayor a la Fecha Actual", vbCritical, "Error de Datos"
''           Txt_FechaIniVig.SetFocus
''           Exit Sub
''        End If
        If (Year(Txt_FechaIniVig) < 1900) Then
           MsgBox "La Fecha Ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
           Txt_FechaIniVig.SetFocus
           Exit Sub
        End If
        
        Txt_FechaIniVig.Text = Format(CDate(Trim(Txt_FechaIniVig)), "yyyymmdd")
        Txt_FechaIniVig.Text = DateSerial(Mid((Txt_FechaIniVig.Text), 1, 4), Mid((Txt_FechaIniVig.Text), 5, 2), Mid((Txt_FechaIniVig.Text), 7, 2))
        
        'Valida Vigencia de la poliza según fecha ingresada en Inicio de Vigencia
        If Not fgValidaVigenciaPoliza(Trim(Txt_PenPoliza.Text), Txt_FechaIniVig.Text) Then
           MsgBox " La Fecha Ingresada es Anterior a la Fecha de Vigencia de la Póliza ", vbInformation, "Información"
           Txt_FechaIniVig.SetFocus
           Exit Sub
        End If
         
        If Not fgValidaPagoPension(Txt_FechaIniVig, Trim(Txt_PenPoliza), vlNumOrden) Then
           MsgBox " Ya se ha Realizado el Proceso de Cálculo de Pensión para ésta Fecha ", vbCritical, "Operación Cancelada"
           Txt_FechaIniVig.SetFocus
           Exit Sub
        End If
        
        Call fgValidaFechaEfecto(Txt_FechaIniVig.Text, Trim(Txt_PenPoliza), vlNumOrden)
        
'        vlFecha = Format(CDate(Txt_FechaIniVig.Text), "yyyymmdd")
'        vlAnno = Mid(Trim(vlFecha), 1, 4)
'        vlMes = Mid(Trim(vlFecha), 5, 2)
'        vlDia = Mid(Trim(vlFecha), 7, 2)
'        Lbl_Hasta.Caption = DateSerial(vlAnno, vlMes + CDbl(Txt_Meses), vlDia)

        'Cmb_TipoRetencion.SetFocus
        Txt_FechaRecepcion.SetFocus
     End If

End Sub

Private Sub Txt_FechaIniVig_LostFocus()

    If (Trim(Txt_FechaIniVig) <> "") Then
        Txt_FechaIniVig.Text = Format(CDate(Trim(Txt_FechaIniVig)), "yyyymmdd")
        Txt_FechaIniVig.Text = DateSerial(Mid((Txt_FechaIniVig.Text), 1, 4), Mid((Txt_FechaIniVig.Text), 5, 2), Mid((Txt_FechaIniVig.Text), 7, 2))
        
        Call fgValidaFechaEfecto(Txt_FechaIniVig.Text, Trim(Txt_PenPoliza), vlNumOrden)
        Lbl_FechaEfecto.Caption = vgFechaEfecto
        
    End If
    
End Sub

Private Sub Txt_FechaRecepcion_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If (Trim(Txt_FechaRecepcion) = "") Then
           MsgBox "Debe ingresar una Fecha de Recepción", vbCritical, "Error de Datos"
           Txt_FechaRecepcion.SetFocus
           Exit Sub
        End If
        If Not IsDate(Txt_FechaRecepcion.Text) Then
           MsgBox "La Fecha Ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
           Txt_FechaRecepcion.SetFocus
           Exit Sub
        End If
        If (CDate(Txt_FechaRecepcion) > CDate(Date)) Then
           MsgBox "La Fecha Ingresada es Mayor a la Fecha Actual", vbCritical, "Error de Datos"
           Txt_FechaRecepcion.SetFocus
           Exit Sub
        End If
        If (Year(Txt_FechaRecepcion) < 1900) Then
           MsgBox "La Fecha Ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
           Txt_FechaRecepcion.SetFocus
           Exit Sub
        End If
        
        Txt_FechaRecepcion.Text = Format(CDate(Trim(Txt_FechaRecepcion)), "yyyymmdd")
        Txt_FechaRecepcion.Text = DateSerial(Mid((Txt_FechaRecepcion.Text), 1, 4), Mid((Txt_FechaRecepcion.Text), 5, 2), Mid((Txt_FechaRecepcion.Text), 7, 2))
        

        'Txt_MtoMaxRet.SetFocus
        Cmb_ModRetencion.SetFocus
     End If


End Sub

Private Sub Txt_FechaRecepcion_LostFocus()

    If (Trim(Txt_FechaRecepcion) <> "") Then
       Txt_FechaRecepcion.Text = Format(CDate(Trim(Txt_FechaRecepcion)), "yyyymmdd")
       Txt_FechaRecepcion.Text = DateSerial(Mid((Txt_FechaRecepcion.Text), 1, 4), Mid((Txt_FechaRecepcion.Text), 5, 2), Mid((Txt_FechaRecepcion.Text), 7, 2))
    End If
    
End Sub

Private Sub Txt_FechaTerVig_KeyPress(KeyAscii As Integer)

'    If KeyAscii = 13 Then
'        If (Trim(Txt_FechaTerVig) = "") Then
'           MsgBox "Debe ingresar una Fecha de Término de Vigencia", vbCritical, "Error de Datos"
'           Txt_FechaTerVig.SetFocus
'           Exit Sub
'        End If
'        If Not IsDate(Txt_FechaTerVig.Text) Then
'           MsgBox "La Fecha Ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
'           Txt_FechaTerVig.SetFocus
'           Exit Sub
'        End If
'        If (CDate(Txt_FechaTerVig) > CDate(Date)) Then
'           MsgBox "La Fecha Ingresada es Mayor a la Fecha Actual", vbCritical, "Error de Datos"
'           Txt_FechaTerVig.SetFocus
'           Exit Sub
'        End If
'        If (Year(Txt_FechaTerVig) < 1900) Then
'           MsgBox "La Fecha Ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
'           Txt_FechaTerVig.SetFocus
'           Exit Sub
'        End If
'
'        Txt_FechaTerVig.Text = Format(CDate(Trim(Txt_FechaTerVig)), "yyyymmdd")
'        Txt_FechaTerVig.Text = DateSerial(Mid((Txt_FechaTerVig.Text), 1, 4), Mid((Txt_FechaTerVig.Text), 5, 2), Mid((Txt_FechaTerVig.Text), 7, 2))
'
''        vlFecha = Format(CDate(Txt_FechaIniVig.Text), "yyyymmdd")
''        vlAnno = Mid(Trim(vlFecha), 1, 4)
''        vlMes = Mid(Trim(vlFecha), 5, 2)
''        vlDia = Mid(Trim(vlFecha), 7, 2)
''        Lbl_Hasta.Caption = DateSerial(vlAnno, vlMes + CDbl(Txt_Meses), vlDia)
'
'        Txt_FechaRecepcion.SetFocus
'     End If


End Sub

Private Sub Txt_FechaTerVig_LostFocus()

    'If (Trim(Txt_FechaTerVig) <> "") Then
    '   Txt_FechaTerVig.Text = Format(CDate(Trim(Txt_FechaTerVig)), "yyyymmdd")
    '   Txt_FechaTerVig.Text = DateSerial(Mid((Txt_FechaTerVig.Text), 1, 4), Mid((Txt_FechaTerVig.Text), 5, 2), Mid((Txt_FechaTerVig.Text), 7, 2))
    'End If
    
End Sub

Private Sub Txt_FecSuspension_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Not IsDate(Txt_FecSuspension.Text) Then
           MsgBox "La Fecha Ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
           Txt_FecSuspension.SetFocus
           Exit Sub
        End If
        If (Year(Txt_FecSuspension) < 1900) Then
           MsgBox "La Fecha Ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
           Txt_FecSuspension.SetFocus
           Exit Sub
        End If
        
        Txt_FecSuspension.Text = Format(CDate(Trim(Txt_FecSuspension)), "yyyymmdd")
        Txt_FecSuspension.Text = DateSerial(Mid((Txt_FecSuspension.Text), 1, 4), Mid((Txt_FecSuspension.Text), 5, 2), Mid((Txt_FecSuspension.Text), 7, 2))
        
        Txt_FechaTerVig.Text = Txt_FecSuspension.Text
        
        Cmd_Grabar.SetFocus
     End If

End Sub

Private Sub Txt_FecSuspension_LostFocus()

    If (Trim(Txt_FecSuspension) <> "") Then
       Txt_FecSuspension.Text = Format(CDate(Trim(Txt_FecSuspension)), "yyyymmdd")
       Txt_FecSuspension.Text = DateSerial(Mid((Txt_FecSuspension.Text), 1, 4), Mid((Txt_FecSuspension.Text), 5, 2), Mid((Txt_FecSuspension.Text), 7, 2))
       
       Txt_FechaTerVig.Text = Txt_FecSuspension.Text
       
    End If

End Sub

Private Sub Txt_Juzgado_KeyPress(KeyAscii As Integer)
On Error GoTo Err_TxtJuzgadoKeyPress

    If KeyAscii = 13 Then
       Txt_Juzgado.Text = Trim(UCase(Txt_Juzgado.Text))
       If Txt_Juzgado.Text = "" Then
          MsgBox "Debe Ingresar Descripción del Juzgado", vbCritical, "Error de Datos"
          Txt_Juzgado.SetFocus
          Exit Sub
       End If
       Cmb_NumIdent.SetFocus
    End If
    
Exit Sub
Err_TxtJuzgadoKeyPress:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Txt_Juzgado_LostFocus()

    Txt_Juzgado.Text = Trim(UCase(Txt_Juzgado.Text))
    If Txt_Juzgado.Text = "" Then
       Exit Sub
    End If

End Sub

Private Sub Txt_Maternorec_KeyPress(KeyAscii As Integer)
On Error GoTo Err_TxtMaternoRecKeyPress

    If KeyAscii = 13 Then
        Txt_MaternoRec.Text = Trim(UCase(Txt_MaternoRec.Text))
        Txt_DireccionRec.SetFocus
    End If
    
Exit Sub
Err_TxtMaternoRecKeyPress:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Txt_Maternorec_LostFocus()
    Txt_MaternoRec.Text = Trim(UCase(Txt_MaternoRec.Text))
    If Txt_MaternoRec.Text = "" Then
       Exit Sub
    End If
End Sub

Private Sub Txt_MontoRetenido_Change()

On Error GoTo Err_TxtMontoRetenidoChange

    If Not IsNumeric(Txt_MontoRetenido) Then
       Txt_MontoRetenido = ""
    End If
    
Exit Sub
Err_TxtMontoRetenidoChange:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Txt_MontoRetenido_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
       If Txt_MontoRetenido.Text = "" Then
          MsgBox "Debe Ingresar Monto de la Retención.", vbCritical, "Error de Datos"
          Txt_MontoRetenido.SetFocus
          Exit Sub
       Else
           Txt_MontoRetenido.Text = Format(Txt_MontoRetenido.Text, "###,###,##0.00")
       End If
    
       If CDbl(Txt_MontoRetenido.Text) = 0 Then
          MsgBox "Debe Ingresar un Valor Mayor que Cero para Monto Retención.", vbCritical, "Error de Datos"
          Txt_MontoRetenido.SetFocus
          Exit Sub
      End If
    
      
      'If Mid(Cmb_ModRetencion, 1, 5) = "PRCIM" Then
     '
     '   lblMontoRetAct = ""
     '
     ' End If
      
    
      If CDbl(Txt_MontoRetenido.Text) > CDbl(lblMontoRetAct) Then
          MsgBox "Valor es mayor al disponible. " & CStr(lblMontoRetAct), vbCritical, "Error de Datos"
          Txt_MontoRetenido.SetFocus
          Exit Sub
      End If
    
    
       If CDbl(Txt_MontoRetenido.Text) > clTopeMtoRet Then
          MsgBox "El Valor Ingresado en Monto Retención Excede el Máximo Permitido de " & Format(clTopeMtoRet, "###,###,##0.00") & "  .", vbExclamation, "Información"
          Txt_MontoRetenido.Text = Format(clTopeMtoRet, "###,###,##0.00")
          Txt_MontoRetenido.SetFocus
          Exit Sub
       End If
       Txt_FechaRecepcion.SetFocus
    End If

End Sub

Private Sub Txt_MontoRetenido_LostFocus()
    
    If Txt_MontoRetenido.Text = "" Then
       Exit Sub
    Else
        Txt_MontoRetenido.Text = Format(Txt_MontoRetenido.Text, "###,###,##0.00")
    End If
    If CDbl(Txt_MontoRetenido.Text) = 0 Then
       Exit Sub
    End If
    If CDbl(Txt_MontoRetenido.Text) > clTopeMtoRet Then
       Txt_MontoRetenido.Text = Format(clTopeMtoRet, "###,###,##0.00")
       Exit Sub
    End If

End Sub

Private Sub Txt_Nombrerec_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Txt_NombreRec.Text = Trim(UCase(Txt_NombreRec.Text))
       If Txt_NombreRec.Text = "" Then
          MsgBox "Debe Ingresar Nombre.", vbCritical, "Error de Datos"
          Txt_NombreSegRec.SetFocus
          Exit Sub
       End If
       Txt_NombreSegRec.SetFocus
    End If
End Sub

Private Sub Txt_Nombrerec_LostFocus()
    Txt_NombreRec.Text = Trim(UCase(Txt_NombreRec.Text))
    If Txt_NombreRec.Text = "" Then
       Exit Sub
    End If
End Sub

Private Sub Txt_NumCuenta_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       vlNumero = InStr(Cmb_TipoRetencion.Text, "-")
       vlOpcion = Trim(Mid(Cmb_TipoRetencion.Text, 1, vlNumero - 1))
       
''       If vlOpcion = clTipoRetencion Then
''          SSTab.Tab = 2
''       Else
''           If vlOpcion = clTipoAsigFam Then
''              Cmd_Grabar.SetFocus
''           End If
''       End If
'       If SSTab.TabEnabled(2) = True Then
'          SSTab.Tab = 2
'       Else
'           Cmd_Grabar.SetFocus
'       End If
        Cmd_Grabar.SetFocus
    End If

End Sub

Private Sub Txt_Paternorec_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Txt_PaternoRec.Text = Trim(UCase(Txt_PaternoRec.Text))
       If Txt_PaternoRec.Text = "" Then
          MsgBox "Debe Ingresar Apellido Paterno.", vbCritical, "Error de Datos"
          Txt_PaternoRec.SetFocus
          Exit Sub
       End If
       Txt_MaternoRec.SetFocus
    End If
End Sub

Private Sub Txt_Paternorec_LostFocus()
    Txt_PaternoRec.Text = Trim(UCase(Txt_PaternoRec.Text))
    If Txt_PaternoRec.Text = "" Then
       Exit Sub
    End If
End Sub

Private Sub txt_pennumident_lostfocus()
    Txt_PenNumIdent = Trim(UCase(Txt_PenNumIdent))
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

Private Sub Txt_Telefonorec_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txt_TelefonoRec = Trim(Txt_TelefonoRec)
        Txt_CorreoRec.SetFocus
    End If
End Sub

Private Sub Txt_Telefonorec_LostFocus()
    Txt_TelefonoRec = Trim(Txt_TelefonoRec)
End Sub

Function flRecibeDireccion(iNomDepartamento As String, iNomProvincia As String, iNomDistrito As String, iCodDir As String)
'FUNCION QUE RECIBE LOS DATOS DEL FORMULARIO DE BUSQUEDA de Dirección
    
    Lbl_Departamento = Trim(iNomDepartamento)
    Lbl_Provincia = Trim(iNomProvincia)
    Lbl_Distrito = Trim(iNomDistrito)
    vlCodDir = iCodDir
    Txt_TelefonoRec.SetFocus

End Function

Function flBuscaMoneda(iNumPoliza As String, iNumEnd As Integer) As String
        
    flBuscaMoneda = ""
        
    vgSql = "select cod_moneda from PP_TMAE_POLIZA "
    vgSql = vgSql & "where num_poliza = '" & iNumPoliza & "' "
    vgSql = vgSql & "and num_endoso = " & iNumEnd & " "
    vgSql = vgSql & "order by num_endoso desc"
    Set vgRs4 = vgConexionBD.Execute(vgSql)
    If Not vgRs4.EOF Then
        flBuscaMoneda = vgRs4!Cod_Moneda
    End If
    vgRs4.Close
    
End Function

Function flCalculaFechaEfecto(iFechaEfectoIngresada As String) As String
Dim iFecha As String
Dim iFechaCierre As String
Dim iFechaEfecto As String

On Error GoTo Err_ValidaFechaEfecto
    
    flCalculaFechaEfecto = ""
    If (Trim(iFechaEfectoIngresada) <> "") Then
        If Not IsDate(iFechaEfectoIngresada) Then
            Exit Function
        End If
        
        iFechaCierre = Format(fgValidaFechaEfecto(Trim(iFechaEfectoIngresada), vlNumPoliza, vlNumOrden), "yyyymmdd")
        iFechaEfecto = Format(CDate(iFechaEfectoIngresada), "yyyymmdd")
        If (iFechaCierre > iFechaEfecto) Then
            'MsgBox "La Fecha de Efecto es anterior a la Fecha del último cierre, el cual corresponde a :" & _
            DateSerial(CInt(Mid(iFechaCierre, 1, 4)), CInt(Mid(iFechaCierre, 5, 2)), CInt(Mid(iFechaCierre, 7, 2))), vbInformation, "Fecha Errónea"
            Exit Function
        Else
            flCalculaFechaEfecto = DateSerial(CInt(Mid(iFechaEfecto, 1, 4)), CInt(Mid(iFechaEfecto, 5, 2)), CInt(Mid(iFechaEfecto, 7, 2)))
        End If
    Else
        
        'Determinar el menor periodo de Proceso que se encuentre Abierto
        
        'Determinar si el periodo a registrar es posterior al que se desea ingresar
        vgSql = "SELECT NUM_PERPAGO,COD_ESTADOREG " ',COD_ESTADOPRI
        vgSql = vgSql & "FROM PP_TMAE_PROPAGOPEN "
        vgSql = vgSql & "WHERE "
        'vgSql = vgSql & "num_perpago >= '" & iFecha & "' AND "
        vgSql = vgSql & "cod_estadoreg <> 'C' "
        'vgSql = vgSql & "or cod_estadopri <> 'C' "
        vgSql = vgSql & "ORDER BY num_perpago ASC"
        Set vgRs2 = vgConexionBD.Execute(vgSql)
        If Not vgRs2.EOF Then
            iFecha = DateSerial(CInt(Mid(vgRs2!Num_PerPago, 1, 4)), CInt(Mid(vgRs2!Num_PerPago, 5, 2)), 1)
        Else
            iFecha = fgBuscaFecServ
        End If
        vgRs2.Close
        
        flCalculaFechaEfecto = fgValidaFechaEfecto(Trim(iFecha), vlNumPoliza, vlNumOrden)
    End If

Exit Function
Err_ValidaFechaEfecto:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function


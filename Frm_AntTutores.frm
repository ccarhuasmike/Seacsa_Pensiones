VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Frm_AntTutores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenedor de Tutores/Apoderados"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   9315
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1155
      Left            =   90
      TabIndex        =   75
      Top             =   825
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   2037
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
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
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox Cmb_PenNumIdent 
      BackColor       =   &H00E0FFFF&
      Height          =   315
      Left            =   855
      Style           =   2  'Dropdown List
      TabIndex        =   70
      Top             =   2010
      Width           =   2235
   End
   Begin VB.TextBox Txt_PenNumIdent 
      Height          =   285
      Left            =   3105
      MaxLength       =   16
      TabIndex        =   69
      Top             =   2010
      Width           =   1875
   End
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
      Height          =   600
      Left            =   90
      TabIndex        =   63
      Top             =   0
      Width           =   9165
      Begin VB.TextBox Txt_PenPoliza 
         Height          =   285
         Left            =   900
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1185
      End
      Begin VB.CommandButton Cmd_BuscarPol 
         Height          =   375
         Left            =   7605
         Picture         =   "Frm_AntTutores.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Buscar Póliza"
         Top             =   165
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
         Left            =   8250
         Picture         =   "Frm_AntTutores.frx":0102
         TabIndex        =   2
         ToolTipText     =   "Buscar"
         Top             =   165
         Width           =   615
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Póliza"
         Height          =   255
         Index           =   0
         Left            =   165
         TabIndex        =   66
         Top             =   255
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "N° End"
         Height          =   195
         Index           =   42
         Left            =   2370
         TabIndex        =   65
         Top             =   270
         Width           =   510
      End
      Begin VB.Label Lbl_End 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   2985
         TabIndex        =   3
         Top             =   240
         Width           =   480
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
         TabIndex        =   64
         Top             =   0
         Width           =   1725
      End
   End
   Begin TabDlg.SSTab SSTab_Tutor 
      Height          =   4995
      Left            =   120
      TabIndex        =   4
      Top             =   2655
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   8811
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   " Antecedentes del Tutor/Apoderado"
      TabPicture(0)   =   "Frm_AntTutores.frx":0204
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Fra_Vigencia"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Fra_Antecedentes"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Fra_Pago"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   " Historia de Tutores/Apoderados"
      TabPicture(1)   =   "Frm_AntTutores.frx":0220
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Cmd_VerHistoria"
      Tab(1).Control(1)=   "Msf_GrillaTutor"
      Tab(1).Control(2)=   "Lbl_Nombre(23)"
      Tab(1).ControlCount=   3
      Begin VB.CommandButton Cmd_VerHistoria 
         Caption         =   "?"
         Height          =   375
         Left            =   -66840
         Picture         =   "Frm_AntTutores.frx":023C
         TabIndex        =   61
         ToolTipText     =   "Buscar Póliza"
         Top             =   1440
         Width           =   615
      End
      Begin MSFlexGridLib.MSFlexGrid Msf_GrillaTutor 
         Height          =   4410
         Left            =   -74775
         TabIndex        =   60
         Top             =   465
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   7779
         _Version        =   393216
         BackColor       =   -2147483624
         ForeColorSel    =   -2147483643
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
         Height          =   1290
         Left            =   120
         TabIndex        =   53
         Top             =   3645
         Width           =   8775
         Begin VB.ComboBox Cmb_Sucursal 
            BackColor       =   &H80000018&
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   615
            Width           =   2955
         End
         Begin VB.TextBox Txt_Cuenta 
            Height          =   285
            Left            =   5280
            MaxLength       =   15
            TabIndex        =   25
            Top             =   930
            Width           =   2910
         End
         Begin VB.ComboBox Cmb_Banco 
            BackColor       =   &H80000018&
            Height          =   315
            Left            =   5280
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   600
            Width           =   2940
         End
         Begin VB.ComboBox Cmb_ViaPago 
            BackColor       =   &H80000018&
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   270
            Width           =   2955
         End
         Begin VB.ComboBox Cmb_TipoCta 
            BackColor       =   &H80000018&
            Height          =   315
            Left            =   5280
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   270
            Width           =   2940
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Sucursal"
            Height          =   255
            Index           =   16
            Left            =   165
            TabIndex        =   59
            Top             =   675
            Width           =   900
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "N°Cuenta"
            Height          =   255
            Index           =   19
            Left            =   4320
            TabIndex        =   58
            Top             =   990
            Width           =   915
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Banco"
            Height          =   255
            Index           =   18
            Left            =   4320
            TabIndex        =   57
            Top             =   660
            Width           =   945
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Vía Pago"
            Height          =   255
            Index           =   15
            Left            =   165
            TabIndex        =   56
            Top             =   330
            Width           =   930
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Tipo Cta."
            Height          =   255
            Index           =   17
            Left            =   4320
            TabIndex        =   55
            Top             =   330
            Width           =   945
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   " Forma de Pago de Pensión"
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
            Index           =   26
            Left            =   120
            TabIndex        =   54
            Top             =   0
            Width           =   2400
         End
      End
      Begin VB.Frame Fra_Antecedentes 
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
         Height          =   2205
         Left            =   120
         TabIndex        =   43
         Top             =   1410
         Width           =   8775
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
            Left            =   8300
            Style           =   1  'Graphical
            TabIndex        =   68
            ToolTipText     =   "Efectuar Busqueda de Dirección"
            Top             =   1515
            Width           =   300
         End
         Begin VB.TextBox Txt_SegNombre 
            Height          =   285
            Left            =   5640
            MaxLength       =   25
            TabIndex        =   12
            Top             =   570
            Width           =   3015
         End
         Begin VB.ComboBox Cmb_NumIdent 
            BackColor       =   &H80000018&
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   225
            Width           =   2955
         End
         Begin VB.CommandButton Cmd_CargaTutor 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6240
            TabIndex        =   10
            ToolTipText     =   "Cargar Datos Tutor"
            Top             =   255
            Width           =   285
         End
         Begin VB.TextBox Txt_NumIdent 
            Height          =   285
            Left            =   4200
            MaxLength       =   10
            TabIndex        =   9
            Top             =   255
            Width           =   1980
         End
         Begin VB.TextBox Txt_Telefono 
            Height          =   285
            Left            =   1200
            MaxLength       =   15
            TabIndex        =   19
            Top             =   1830
            Width           =   2025
         End
         Begin VB.TextBox Txt_Nombre 
            Height          =   285
            Left            =   1200
            MaxLength       =   25
            TabIndex        =   11
            Top             =   570
            Width           =   2895
         End
         Begin VB.TextBox Txt_Direccion 
            Height          =   285
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   15
            Top             =   1200
            Width           =   7455
         End
         Begin VB.TextBox Txt_Paterno 
            Height          =   285
            Left            =   1200
            MaxLength       =   20
            TabIndex        =   13
            Top             =   885
            Width           =   2895
         End
         Begin VB.TextBox Txt_Materno 
            Height          =   285
            Left            =   5640
            MaxLength       =   20
            TabIndex        =   14
            Top             =   885
            Width           =   3015
         End
         Begin VB.TextBox Txt_Email 
            Height          =   285
            Left            =   4080
            MaxLength       =   40
            TabIndex        =   20
            Top             =   1830
            Width           =   4570
         End
         Begin VB.Label Lbl_Distrito 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   5895
            TabIndex        =   18
            Top             =   1515
            Width           =   2295
         End
         Begin VB.Label Lbl_Provincia 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   3540
            TabIndex        =   17
            Top             =   1515
            Width           =   2295
         End
         Begin VB.Label Lbl_Departamento 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   16
            Top             =   1515
            Width           =   2295
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Segundo Nombre"
            Height          =   255
            Index           =   7
            Left            =   4320
            TabIndex        =   67
            Top             =   630
            Width           =   1395
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "N° Ident."
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   52
            Top             =   285
            Width           =   810
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Teléfono"
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   51
            Top             =   1875
            Width           =   795
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Primer Nombre"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   50
            Top             =   600
            Width           =   1155
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Dirección"
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   49
            Top             =   1230
            Width           =   735
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Ubicación"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   48
            Top             =   1560
            Width           =   855
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Email"
            Height          =   255
            Index           =   13
            Left            =   3480
            TabIndex        =   47
            Top             =   1890
            Width           =   615
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Ap. Paterno"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   46
            Top             =   900
            Width           =   915
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Ap. Materno"
            Height          =   255
            Index           =   10
            Left            =   4320
            TabIndex        =   45
            Top             =   930
            Width           =   915
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   " Antecedenes Personales del Tutor/Apoderado"
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
            Index           =   25
            Left            =   120
            TabIndex        =   44
            Top             =   0
            Width           =   4155
         End
      End
      Begin VB.Frame Fra_Vigencia 
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
         Height          =   1080
         Left            =   120
         TabIndex        =   33
         Top             =   330
         Width           =   8775
         Begin VB.TextBox Txt_FecRec 
            Height          =   285
            Left            =   6000
            TabIndex        =   7
            Top             =   700
            Width           =   1185
         End
         Begin VB.TextBox Txt_Desde 
            Height          =   285
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   6
            Top             =   700
            Width           =   1200
         End
         Begin VB.TextBox Txt_Meses 
            Height          =   285
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   5
            Top             =   360
            Width           =   690
         End
         Begin Crystal.CrystalReport Rpt_General 
            Left            =   8130
            Top             =   195
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowState     =   2
            PrintFileLinesPerPage=   60
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Duración (meses)"
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   42
            Top             =   240
            Width           =   780
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            Caption         =   "Periodo de Vigencia"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   41
            Top             =   720
            Width           =   1425
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Recepción"
            Height          =   195
            Index           =   5
            Left            =   4560
            TabIndex        =   40
            Top             =   720
            Width           =   1275
         End
         Begin VB.Label Lbl_Hasta 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3240
            TabIndex        =   39
            Top             =   705
            Width           =   1185
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   " Vigencia del Poder"
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
            Left            =   120
            TabIndex        =   38
            Top             =   0
            Width           =   1750
         End
         Begin VB.Label Lbl_FecRec 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7320
            TabIndex        =   37
            Top             =   720
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.Label Lbl_FecEfecto 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6000
            TabIndex        =   36
            Top             =   360
            Width           =   1185
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            Caption         =   "Periodo de Efecto"
            Height          =   195
            Index           =   27
            Left            =   4560
            TabIndex        =   35
            Top             =   360
            Width           =   1275
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
            Index           =   28
            Left            =   3000
            TabIndex        =   34
            Top             =   720
            Width           =   255
         End
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Ver Historial"
         Height          =   255
         Index           =   23
         Left            =   -66960
         TabIndex        =   62
         Top             =   1920
         Width           =   855
      End
   End
   Begin VB.Frame Fra_Operacion 
      Height          =   885
      Left            =   120
      TabIndex        =   32
      Top             =   7590
      Width           =   9075
      Begin VB.CommandButton Cmd_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   5880
         Picture         =   "Frm_AntTutores.frx":033E
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   150
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   3555
         Picture         =   "Frm_AntTutores.frx":0918
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   150
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Eliminar 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   2415
         Picture         =   "Frm_AntTutores.frx":0FD2
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   150
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   4695
         Picture         =   "Frm_AntTutores.frx":1314
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Limpiar Formulario"
         Top             =   150
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   7080
         Picture         =   "Frm_AntTutores.frx":19CE
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Salir del Formulario"
         Top             =   150
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Grabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   1275
         Picture         =   "Frm_AntTutores.frx":1AC8
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Grabar Datos"
         Top             =   150
         Width           =   720
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Beneficiarios"
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
      Height          =   210
      Left            =   210
      TabIndex        =   74
      Top             =   615
      Width           =   1515
   End
   Begin VB.Label Lbl_Nombre 
      Caption         =   "Nombre"
      Height          =   255
      Index           =   2
      Left            =   135
      TabIndex        =   73
      Top             =   2355
      Width           =   735
   End
   Begin VB.Label Lbl_Nombre 
      Caption         =   "N° Ident."
      Height          =   255
      Index           =   1
      Left            =   135
      TabIndex        =   72
      Top             =   2055
      Width           =   855
   End
   Begin VB.Label Lbl_PenNombre 
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   255
      Left            =   855
      TabIndex        =   71
      Top             =   2340
      Width           =   8205
   End
End
Attribute VB_Name = "Frm_AntTutores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vlRegistroBen As ADODB.Recordset

Dim Sql As String

Dim vlFecha As String
Dim vlFechaInicio As String
Dim vlFechaFin As String
Dim vlDia As String
Dim vlMes As String
Dim vlAnno As String
Dim vlPos As Integer
Dim vlNumEndoso As Integer
Dim vlNumOrden As Integer
Dim vlcont As Integer
Dim vlCodDir As Integer
Dim vlArchivo As String
Dim vlNumero As Integer
Dim vlOpcion As String
Dim vlNombreRegion As String
Dim vlNombreProvincia As String
Dim vlNombreComuna As String
Dim vlViaPago As String
Dim vlSw As Boolean
Dim vlNumPoliza As String
Dim vlRut As String
Dim vlDigito As String
Dim vlRutAux As String
Dim vlSwSinDerPen As String * 1
Dim vlSwModificar As String * 1
Dim vlPagoReg As String
Dim vlPagoProxReg As String
Dim vlPagoProxPri As String
Dim vlPagoPri As String
'Dim vlNumEndoso As String


Dim vlGlsUsuarioCrea As Variant
Dim vlFecCrea As Variant
Dim vlHorCrea As Variant
Dim vlGlsUsuarioModi As Variant
Dim vlFecModi As Variant
Dim vlHorModi As Variant

Dim vlRegistro As ADODB.Recordset
Dim vlRegistro1 As ADODB.Recordset

Const clMesesMax As Integer = 1200
Const clCodSinDerPen As String * 2 = "10"
Const clCodTipReceptor As String * 1 = "T"

Dim vlCodTipoIdenBenCau As String
Dim vlNumIdenBenCau As String

Dim vlCodTipoIdenBenTut As String
Dim vlNumIdenBenTut As String

Dim vlLargoTipoIden    As Integer 'sirve para llenar la grilla
Dim vlPosicionTipoIden As Integer 'sirve para llenar la grilla

Dim vlNombreSeg As String, vlApMaterno As String
Dim vlAfp As String
Dim RS As ADODB.Recordset

Function flCalcularFechaHasta(iFechaDesde, iMesesVigencia) As String
Dim iFecha As String
Dim iAnno As Integer, iMes As Integer, iDia As Integer
    
    flCalcularFechaHasta = ""
    
    iFecha = Format(CDate(iFechaDesde), "yyyymmdd")
    iAnno = Mid(Trim(iFecha), 1, 4)
    iMes = Mid(Trim(iFecha), 5, 2)
    iDia = Mid(Trim(iFecha), 7, 2)
    'Debería restar un día para completar un Mes Entero
    flCalcularFechaHasta = DateSerial(iAnno, iMes + iMesesVigencia, iDia - 1)

End Function

Function flHabilitarIngreso()
On Error GoTo Err_FlHabilitarIngreso

    Fra_Poliza.Enabled = False
    
'    Txt_PenPoliza.Enabled = False
'    Txt_PenRut.Enabled = False
'    Txt_PenDigito.Enabled = False
'    Lbl_PenNombre.Enabled = False
'    Lbl_PenEndoso.Enabled = False
'    Cmd_Buscar.Enabled = False
'    Cmd_BuscarPol.Enabled = False
    
    
    Fra_Vigencia.Enabled = True
    
'    Txt_Meses.Enabled = True
'    Txt_Desde.Enabled = True
'    Lbl_Hasta.Enabled = True

    Fra_Antecedentes.Enabled = True

'    Txt_Rut.Enabled = True
'    Txt_Digito.Enabled = True
'    Txt_Nombre.Enabled = True
'    Txt_Paterno.Enabled = True
'    Txt_Materno.Enabled = True
'    Txt_Direccion.Enabled = True
'    Cmb_Comuna.Enabled = True
'    Lbl_Provincia.Enabled = True
'    Lbl_Region.Enabled = True
'    Txt_Telefono.Enabled = True
'    Txt_Email.Enabled = True
    
    Cmb_ViaPago.Enabled = True
    Cmb_Sucursal.Enabled = False
    Cmb_TipoCta.Enabled = False
    Cmb_Banco.Enabled = False
    Txt_Cuenta.Enabled = False
    
'''    Lbl_FecRec.Caption = fgBuscaFecServ
    
    SSTab_Tutor.Enabled = True
    SSTab_Tutor.TabEnabled(0) = True
    SSTab_Tutor.TabEnabled(1) = True
    
Exit Function
Err_FlHabilitarIngreso:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flIniciaForm()
On Error GoTo Err_FlIniciaForm

    Txt_PenPoliza.Text = ""
    If (Cmb_PenNumIdent.ListCount <> 0) Then
        Cmb_PenNumIdent.ListIndex = 0
    End If
    Txt_PenNumIdent.Text = ""
'    Txt_PenDigito.Text = ""
    Lbl_PenNombre.Caption = ""
    Lbl_End.Caption = ""
    
    Txt_Meses.Text = ""
    Txt_Desde.Text = ""
    Lbl_Hasta.Caption = ""
    
    If (Cmb_NumIdent.ListCount <> 0) Then
        Cmb_NumIdent.ListIndex = 0
    End If
    Txt_NumIdent.Text = ""
''*    Txt_Digito.Text = ""
    Txt_Nombre.Text = ""
    Txt_SegNombre.Text = ""
    Txt_Paterno.Text = ""
    Txt_Materno.Text = ""
    Txt_Direccion.Text = ""
    
''*    If (Cmb_Comuna.ListCount > 0) Then
''        Cmb_Comuna.ListIndex = 0
''    End If
    Lbl_Departamento.Caption = ""
    Lbl_Provincia.Caption = ""
    Lbl_Distrito.Caption = ""
    
    Txt_Telefono.Text = ""
    Txt_Email.Text = ""
    
    Cmb_ViaPago.ListIndex = 0
    'Cmb_Sucursal.ListIndex = 0
    Call fgComboSucursal(Cmb_Sucursal, "S")
    Cmb_Sucursal.Enabled = False
    
    Cmb_TipoCta.ListIndex = 0
    Cmb_Banco.ListIndex = 0
    Txt_Cuenta.Text = ""

    Fra_Poliza.Enabled = True
    
'    Txt_PenPoliza.Enabled = True
'    Txt_PenRut.Enabled = True
'    Txt_PenDigito.Enabled = True
'    Lbl_PenNombre.Enabled = True
'    Lbl_PenEndoso.Enabled = True
'    Cmd_Buscar.Enabled = True
'    Cmd_BuscarPol.Enabled = True
    
    Fra_Vigencia.Enabled = False
    
'    Txt_Meses.Enabled = False
'    Txt_Desde.Enabled = False
'    Lbl_Hasta.Enabled = False
    
    Fra_Antecedentes.Enabled = False
    
'    Txt_Rut.Enabled = False
'    Txt_Digito.Enabled = False
'    Txt_Nombre.Enabled = False
'    Txt_Paterno.Enabled = False
'    Txt_Materno.Enabled = False
'    Txt_Direccion.Enabled = False
'    Cmb_Comuna.Enabled = False
'    Lbl_Provincia.Enabled = False
'    Lbl_Region.Enabled = False
'    Txt_Telefono.Enabled = False
'    Txt_Email.Enabled = False
    
'    Fra_Pago.Enabled = False
    
    Cmb_ViaPago.Enabled = False
    Cmb_Sucursal.Enabled = False
    Cmb_TipoCta.Enabled = False
    Cmb_Banco.Enabled = False
    Txt_Cuenta.Enabled = False
    
    Txt_FecRec.Text = ""
    Lbl_FecEfecto.Caption = ""
    
    SSTab_Tutor.Enabled = False
    
    Call flInicializaGrillaTutores
    SSTab_Tutor.Tab = 0
    Set DataGrid1.DataSource = Nothing
    Set RS = Nothing
    
Exit Function
Err_FlIniciaForm:
    Screen.MousePointer = 0
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

''''Function flCalculaFecEfecto() As String
''''On Error GoTo Err_flCalculaFecEfecto
''''
''''     flCalculaFecEfecto = ""
''''    'Verifica Último Periodo
''''     vgSql = ""
''''     vgSql = "SELECT NUM_PERPAGO,COD_ESTADOPRI,COD_ESTADOREG,"
''''     vgSql = vgSql & "FEC_PRIPAGO,FEC_PAGOPROXREG,FEC_PAGOREG FROM PP_TMAE_PROPAGOPEN ORDER BY num_perpago DESC"
''''     Set vlRegistro1 = vgConexionBD.Execute(vgSql)
''''     If Not vlRegistro1.EOF Then
''''           'Verifica si es Primer Pago o Pago Régimen
''''            vgSql = ""
''''            vgSql = "SELECT NUM_POLIZA,NUM_ENDOSO,NUM_ORDEN FROM PP_TMAE_LIQPAGOPENDEF"
''''            vgSql = vgSql & " Where "
''''            vgSql = vgSql & " NUM_POLIZA = '" & Trim(Txt_PenPoliza) & "' AND "
''''            vgSql = vgSql & " NUM_ENDOSO = " & vlNumEndoso & " AND "
''''            vgSql = vgSql & " NUM_ORDEN = " & vlNumOrden & " "
''''            Set vlRegistro = vgConexionBD.Execute(vgSql)
''''            If Not vlRegistro.EOF Then
''''                  'Pago Régimen
''''                   If (vlRegistro1!cod_estadoreg) = "A" Or (vlRegistro1!cod_estadoreg) = "P" Then
''''                       vlPagoReg = (vlRegistro1!Num_PerPago)
''''                       vlAnno = Mid(vlPagoReg, 1, 4)
''''                       vlMes = Mid(vlPagoReg, 5, 2)
''''                       vlDia = "01"
''''                       flCalculaFecEfecto = DateSerial(vlAnno, vlMes, vlDia)
''''                   Else
''''                       If (vlRegistro1!cod_estadoreg) = "C" Then
''''                           vlPagoProxReg = (vlRegistro1!Num_PerPago)
''''                           vlAnno = Mid(vlPagoProxReg, 1, 4)
''''                           vlMes = Mid(vlPagoProxReg, 5, 2)
''''                           vlMes = (vlMes) + 1
''''                           vlDia = "01"
''''                           flCalculaFecEfecto = DateSerial(vlAnno, vlMes, vlDia)
''''                       End If
''''                   End If
''''            Else
''''                  'Primer Pago
''''                   If (vlRegistro1!cod_estadopri) = "A" Or (vlRegistro1!cod_estadopri) = "P" Then
''''                       vlPagoPri = (vlRegistro1!Num_PerPago)
''''                       vlAnno = Mid(vlPagoPri, 1, 4)
''''                       vlMes = Mid(vlPagoPri, 5, 2)
''''                       vlDia = "01"
''''                       flCalculaFecEfecto = DateSerial(vlAnno, vlMes, vlDia)
''''                   Else
''''                       If (vlRegistro1!cod_estadopri) = "C" Then
''''                           vlPagoProxPri = (vlRegistro1!Num_PerPago)
''''                           vlAnno = Mid(vlPagoProxPri, 1, 4)
''''                           vlMes = Mid(vlPagoProxPri, 5, 2)
''''                           vlMes = (vlMes) + 1
''''                           vlDia = "01"
''''                           flCalculaFecEfecto = DateSerial(vlAnno, vlMes, vlDia)
''''                       End If
''''                   End If
''''            End If
''''     End If
''''
''''Exit Function
''''Err_flCalculaFecEfecto:
''''    If Err.Number <> 0 Then
''''        MsgBox Err.Description & " " & "Numero de Error : " & Err.Number & Err.HelpFile & Err.HelpContext, vbExclamation, "Mensaje de error ..."
''''        Screen.MousePointer = 0
''''    End If
''''End Function

Function flInicializaGrillaTutores()
    
    Msf_GrillaTutor.Clear
    Msf_GrillaTutor.Cols = 5
    Msf_GrillaTutor.rows = 1
    Msf_GrillaTutor.RowHeight(0) = 250
    Msf_GrillaTutor.row = 0
    
    Msf_GrillaTutor.Col = 0
    Msf_GrillaTutor.Text = "Periodo"
    Msf_GrillaTutor.ColWidth(0) = 1000
    
    Msf_GrillaTutor.Col = 1
    Msf_GrillaTutor.Text = "Fecha Pago"
    Msf_GrillaTutor.ColWidth(1) = 1200
    
    Msf_GrillaTutor.Col = 2
    Msf_GrillaTutor.Text = "Tipo Ident. Tutor"
    Msf_GrillaTutor.ColWidth(2) = 1500
    
    Msf_GrillaTutor.Col = 3
    Msf_GrillaTutor.Text = "Nº Ident. Tutor"
    Msf_GrillaTutor.ColWidth(3) = 1500

    Msf_GrillaTutor.Col = 4
    Msf_GrillaTutor.Text = "Nombre Tutor"
    Msf_GrillaTutor.ColWidth(4) = 3600
    
End Function

Function flCargaGrillaTutor()

On Error GoTo Err_flCargaGrillaTutor

Dim cont As Integer
Dim vlTipoI As String

    vgSql = ""
    vgSql = "SELECT num_perpago,fec_pago,cod_tipoidenreceptor,num_idenreceptor, "
    vgSql = vgSql & "gls_nomreceptor,gls_nomsegreceptor,gls_matreceptor,gls_patreceptor "
    vgSql = vgSql & "FROM PP_TMAE_LIQPAGOPENDEF "
    vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
    vgSql = vgSql & "num_orden = " & Trim(vlNumOrden) & " AND "
    vgSql = vgSql & "cod_tipreceptor = '" & Trim(clCodTipReceptor) & "' "
    vgSql = vgSql & "ORDER by num_perpago,num_endoso DESC "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
       Call flInicializaGrillaTutores
       cont = 1
       For cont = 1 To 36
            If Not vgRs.EOF Then
            
               vlTipoI = Trim(vgRs!Cod_TipoIdenReceptor) & " - " & fgBuscarNombreTipoIden(Trim(vgRs!Cod_TipoIdenReceptor), False)
            
               Msf_GrillaTutor.AddItem (Mid(Trim(vgRs!Num_PerPago), 5, 2)) & " - " & Mid(Trim(vgRs!Num_PerPago), 1, 4) & vbTab _
               & (DateSerial(Mid((vgRs!Fec_Pago), 1, 4), Mid((vgRs!Fec_Pago), 5, 2), Mid((vgRs!Fec_Pago), 7, 2))) & vbTab _
               & vlTipoI & vbTab & (Trim(vgRs!Num_IdenReceptor)) & vbTab _
               & (Trim(vgRs!Gls_NomReceptor)) & (Trim(vgRs!Gls_PatReceptor)) & IIf(IsNull(Trim(vgRs!Gls_MatReceptor)), "", Trim(vgRs!Gls_MatReceptor))
               
               vgRs.MoveNext
            Else
                Exit For
            End If
       Next cont
    Else
        MsgBox "El Beneficiario Ingresado No tiene Historial de Tutores.", vbInformation, "Información"
    End If
    vgRs.Close

Exit Function
Err_flCargaGrillaTutor:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " " & "Numero de Error : " & Err.Number & Err.HelpFile & Err.HelpContext, vbExclamation, "Mensaje de error ..."
        Screen.MousePointer = 0
    End If
End Function

Private Sub Cmb_Banco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Txt_Cuenta.SetFocus
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
            Txt_Nombre.SetFocus
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

Private Sub Cmb_Sucursal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Cmb_TipoCta.Enabled = True Then
          Cmb_TipoCta.SetFocus
       Else
           Cmd_Grabar.SetFocus
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

Private Sub Cmb_ViaPago_Click()
On Error GoTo Err_CmbViaPagoClick

If Cmb_ViaPago.Enabled = True Then

    vlNumero = InStr(Cmb_ViaPago.Text, "-")
    vlOpcion = Trim(Mid(Cmb_ViaPago.Text, 1, vlNumero - 1))
'Opción 01 = Via de Pago CAJA
    If vlSw = False Then
       
        If vlOpcion = "00" Or vlOpcion = "05" Then
            vgTipoSucursal = cgTipoSucursalSuc
            fgComboSucursal Cmb_Sucursal, vgTipoSucursal
            
            Cmb_Sucursal.Enabled = False
            Cmb_TipoCta.Enabled = False
            Cmb_Banco.Enabled = False
            Txt_Cuenta.Enabled = False
            Cmb_Sucursal.ListIndex = 0
            Cmb_TipoCta.ListIndex = 0
            Cmb_Banco.ListIndex = 0
            Txt_Cuenta.Text = ""
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
                Txt_Cuenta.Enabled = False
                Cmb_TipoCta.ListIndex = 0
                Cmb_Banco.ListIndex = 0
                If (vlOpcion = "04") Then
                    vgPalabra = fgObtenerCodigo_TextoCompuesto(vlAfp)
                    Call fgBuscarPosicionCodigoCombo(vgPalabra, Cmb_Sucursal)
                End If
                Txt_Cuenta.Text = ""
            Else
                If (vlOpcion = "02") Or (vlOpcion = "03") Then
                   Cmb_Sucursal.Enabled = False
                   Cmb_TipoCta.Enabled = True
                   Cmb_Banco.Enabled = True
                   Txt_Cuenta.Enabled = True
                   Cmb_Sucursal.ListIndex = 0
'                   Txt_Cuenta.Text = ""
                Else
                    Cmb_Sucursal.Enabled = True
                    Cmb_TipoCta.Enabled = True
                    Cmb_Banco.Enabled = True
                    Txt_Cuenta.Enabled = True
                    Cmb_Sucursal.ListIndex = 0
                    Cmb_TipoCta.ListIndex = 0
                    Cmb_Banco.ListIndex = 0
                    Txt_Cuenta.Text = ""
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
'Opción 01 = Via de Pago CAJA
        If vlOpcion = "00" Or vlOpcion = "05" Then
            vgTipoSucursal = cgTipoSucursalSuc
            fgComboSucursal Cmb_Sucursal, vgTipoSucursal
            
            Cmb_Sucursal.Enabled = False
            Cmb_TipoCta.Enabled = False
            Cmb_Banco.Enabled = False
            Txt_Cuenta.Enabled = False
            Cmb_Sucursal.ListIndex = 0
            Cmb_TipoCta.ListIndex = 0
            Cmb_Banco.ListIndex = 0
            Txt_Cuenta.Text = ""
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
                Txt_Cuenta.Enabled = False
                Cmb_TipoCta.ListIndex = 0
                Cmb_Banco.ListIndex = 0
                If (vlOpcion = "04") Then
                    vgPalabra = fgObtenerCodigo_TextoCompuesto(vlAfp)
                    Call fgBuscarPosicionCodigoCombo(vgPalabra, Cmb_Sucursal)
                End If
                Txt_Cuenta.Text = ""
            Else
                If (vlOpcion = "02") Or (vlOpcion = "03") Then
                   Cmb_Sucursal.Enabled = False
                   Cmb_TipoCta.Enabled = True
                   Cmb_Banco.Enabled = True
                   Txt_Cuenta.Enabled = True
                   Cmb_Sucursal.ListIndex = 0
'                   Txt_Cuenta.Text = ""
                Else
                    Cmb_Sucursal.Enabled = True
                    Cmb_TipoCta.Enabled = True
                    Cmb_Banco.Enabled = True
                    Txt_Cuenta.Enabled = True
                    Cmb_Sucursal.ListIndex = 0
                    Cmb_TipoCta.ListIndex = 0
                    Cmb_Banco.ListIndex = 0
                    Txt_Cuenta.Text = ""
                End If
            End If
        End If
            
        If Cmb_Sucursal.Enabled = True Then
           Cmb_Sucursal.SetFocus
        Else
            Cmb_TipoCta.SetFocus
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

    Frm_Busqueda.flInicio ("Frm_AntTutores")
    
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

    Frm_BusDireccion.flInicio ("Frm_AntTutores")
    
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
        
''''    If Txt_PenPoliza.Text = "" Then
''''       MsgBox "Debe Ingresar Número de Póliza.", vbCritical, "Error de Datos"
''''       Txt_PenPoliza.SetFocus
''''       Exit Sub
''''    End If
''''    If (Trim(Txt_PenRut.Text)) = "" Then
''''       MsgBox "Debe ingresar el Rut del Pensionado.", vbCritical, "Error de Datos"
''''       Txt_PenRut.SetFocus
''''       Exit Sub
''''    Else
''''        Txt_PenRut = Format(Txt_PenRut, "##,###,##0")
''''        Txt_PenDigito = UCase(Trim(Txt_PenDigito))
''''        Txt_PenDigito.SetFocus
''''    End If
''''    If Txt_PenDigito.Text = "" Then
''''       MsgBox "Debe Ingresar Dígito Verificador de Rut de Pensionado.", vbCritical, "Error de Datos"
''''       Txt_PenDigito.SetFocus
''''       Exit Sub
''''    End If
''''    If Not ValiRut(Txt_PenRut.Text, Txt_PenDigito.Text) Then
''''       MsgBox "El Rut Ingresado es incorrecto.", vbCritical, "Error de Datos"
''''       Txt_PenRut.SetFocus
''''       Exit Sub
''''    End If
''''
''''    Txt_PenPoliza.Text = Trim(Txt_PenPoliza.Text)
''''    vlRutAux = Format(Txt_PenRut, "#0")
''''
''''    vgSql = "SELECT num_endoso,num_orden,gls_nomben,gls_patben,gls_matben "
''''    vgSql = vgSql & "FROM PP_TMAE_BEN WHERE "
''''    vgSql = vgSql & "num_poliza = '" & Txt_PenPoliza.Text & "' and "
''''    vgSql = vgSql & "rut_ben = " & vlRutAux & " "
''''    vgSql = vgSql & "ORDER BY num_endoso DESC "
''''    Set vgRs2 = vgConexionBD.Execute(vgSql)

vlRutAux = ""
vlSwModificar = "N"

  If Txt_PenPoliza.Text = "" Then
       If ((Trim(Cmb_PenNumIdent.Text)) = "") Or (Txt_PenNumIdent.Text = "") Then
       ''*Or _ (Not ValiRut(Txt_PenRut.Text, Txt_PenDigito.Text))
           MsgBox "Debe Ingresar el Número de Póliza o la Identificación del Pensionado.", vbCritical, "Error de Datos"
           Txt_PenPoliza.SetFocus
           Exit Sub
       Else
           Txt_PenNumIdent = UCase(Trim(Txt_PenNumIdent))
           ''Txt_PenDigito = UCase(Trim(Txt_PenDigito))
           Txt_PenNumIdent.SetFocus
           ''vlRutAux = Format(Txt_PenRut, "#0")
       End If
    Else
        Txt_PenPoliza.Text = Trim(Txt_PenPoliza.Text)
    End If
        
'''''    vgSql = ""
'''''    vgSql = "SELECT num_endoso,num_orden,gls_nomben,gls_patben,gls_matben, "
'''''    vgSql = vgSql & "cod_estpension,rut_ben,dgv_ben,num_poliza "
'''''    vgSql = vgSql & "FROM PP_TMAE_BEN WHERE "
'''''    If Txt_PenPoliza.Text <> "" Then
'''''       vgSql = vgSql & "num_poliza = '" & Txt_PenPoliza.Text & "' AND "
'''''    End If
'''''    If Txt_PenRut.Text <> "" Then
'''''        vgSql = vgSql & "rut_ben = " & vlRutAux & " AND "
'''''    End If
'''''    vgSql = vgSql & "cod_estpension <> '" & clCodSinDerPen & "' "
'''''    vgSql = vgSql & "ORDER BY num_endoso DESC "
'''''    Set vgRs2 = vgConexionBD.Execute(vgSql)
    
'CMV/20041102-I

    vlCodTipoIdenBenCau = fgObtenerCodigo_TextoCompuesto(Cmb_PenNumIdent)
    vlNumIdenBenCau = Txt_PenNumIdent
    
    vgPalabra = ""
    'Seleccionar beneficiario, según número de póliza y rut de beneficiario.
    If (Txt_PenPoliza.Text <> "") And (Cmb_PenNumIdent.Text <> "") And (Txt_PenNumIdent.Text <> "") Then ''*
        ''vlRutAux = Format(Txt_PenRut, "#0")
        vgPalabra = "num_poliza = '" & Txt_PenPoliza.Text & "' AND "
        vgPalabra = vgPalabra & "cod_tipoidenben = " & vlCodTipoIdenBenCau & " AND "
        vgPalabra = vgPalabra & "num_idenben = '" & vlNumIdenBenCau & "' "
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
               vgPalabra = "cod_tipoidenben = " & vlCodTipoIdenBenCau & " AND "
               vgPalabra = vgPalabra & "num_idenben = '" & vlNumIdenBenCau & "' "
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
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
        
        vlAfp = fgObtenerPolizaCod_AFP(vgRs!num_poliza, CStr(vgRs!num_endoso))

       If Txt_PenPoliza.Text <> "" Then
          vlCodTipoIdenBenCau = vgRs!Cod_TipoIdenBen
          vlNumIdenBenCau = Trim(vgRs!Num_IdenBen)
           Call fgBuscarPosicionCodigoCombo(vlCodTipoIdenBenCau, Cmb_PenNumIdent)
          Txt_PenNumIdent.Text = vlNumIdenBenCau
          ''Txt_PenDigito.Text = Trim(vgRs!Dgv_Ben)
          ''Txt_PenRut = Format(Txt_PenRut, "##,###,##0")
          ''Txt_PenDigito = UCase(Trim(Txt_PenDigito))
       Else
           Txt_PenPoliza.Text = Trim(vgRs!num_poliza)
       End If
    
       vlSwSinDerPen = ""
       If Trim(vgRs!Cod_EstPension) = Trim(clCodSinDerPen) Then
          MsgBox " El Beneficiario Seleccionado No Tiene Derecho a Pensión " & Chr(13) & _
                 "          Sólo podrá Consultar los Datos del Registro", vbInformation, "Información"
                 
           vlSwSinDerPen = "S"
           vlSwModificar = "N"
          'Desactivar Todos los Controles del Formulario
          '(Se realiza la desactivafción al final del Sub)
       Else
           If Not fgValidaVigenciaPoliza(Trim(Txt_PenPoliza.Text), fgBuscaFecServ) Then
              MsgBox " La Póliza Ingresada no se Encuentra Vigente en el Sistema " & Chr(13) & _
                     "      Sólo podrá Consultar los Datos del Registro", vbInformation, "Información"
                     
              vlSwSinDerPen = "S"
              vlSwModificar = "N"
              'Desactivar Todos los Controles del Formulario
              '(Se realiza la desactivafción al final del Sub)
           Else
               vlSwSinDerPen = "N"
               vlSwModificar = "S"
               Call flHabilitarIngreso
           End If
       End If
       
       If IsNull(vgRs!Gls_NomSegBen) Then
          vlNombreSeg = ""
       Else
           vlNombreSeg = Trim(vgRs!Gls_NomSegBen)
       End If
        If IsNull(vgRs!Gls_MatBen) Then
          vlApMaterno = ""
       Else
           vlApMaterno = Trim(vgRs!Gls_MatBen)
       End If
                      
        Lbl_PenNombre.Caption = fgFormarNombreCompleto(Trim(vgRs!Gls_NomBen), vlNombreSeg, Trim(vgRs!Gls_PatBen), vlApMaterno)
        
       'Lbl_PenNombre.Caption = Trim(vgRs!Gls_NomBen) + " " + Trim(vgRs!Gls_NomSegBen) + " " + Trim(vgRs!Gls_PatBen) + " " + Trim(vgRs!Gls_MatBen)
       Lbl_End.Caption = (vgRs!num_endoso)
       vlNumEndoso = (vgRs!num_endoso)
       vlNumOrden = (vgRs!Num_Orden)
    Else
        MsgBox "El Beneficiario o la Póliza Ingresados, No Existen en la Base de Datos", vbInformation, "Información"
        Exit Sub
    End If
'    vgRs2.Close
'CMV/20041102-F
              
    'marco----25/03/2010
    If RS.State = 0 Then
        
            Set RS = New ADODB.Recordset
            RS.CursorLocation = adUseClient
            RS.Open "PP_LISTA_MENORES_INVALIDOS.LISTAR('" & Trim(Txt_PenPoliza.Text) & "')", vgConexionBD, adOpenForwardOnly, adLockReadOnly
            If Not RS.EOF Then
                Set DataGrid1.DataSource = RS
                DataGrid1.Refresh
            End If
            Call FormateaGrilla

    End If
    
    vgSql = ""
    vgSql = "SELECT *  FROM PP_TMAE_TUTOR "
    vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
    'I---- ABV 21/08/2004 ---
    'vgSql = vgSql & "num_endoso = " & Trim(vlNumEndoso) & " AND "
    vgSql = vgSql & "num_orden = " & Trim(vlNumOrden) & " "
    'F---- ABV 21/08/2004 ---
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
       Txt_Meses.Text = (vgRs!num_mespodnot)
       Txt_Desde.Text = (vgRs!fec_inipodnot)
       Txt_Desde.Text = DateSerial((Mid(Txt_Desde.Text, 1, 4)), (Mid(Txt_Desde.Text, 5, 2)), (Mid(Txt_Desde.Text, 7, 2)))
       Lbl_Hasta.Caption = (vgRs!Fec_TerPodNot)
       Lbl_Hasta.Caption = DateSerial((Mid(Lbl_Hasta.Caption, 1, 4)), (Mid(Lbl_Hasta.Caption, 5, 2)), (Mid(Lbl_Hasta.Caption, 7, 2)))
              
       If IsNull(vgRs!FEC_EFECTO) Then
          Lbl_FecEfecto.Caption = ""
       Else
           Lbl_FecEfecto.Caption = (vgRs!FEC_EFECTO)
           Lbl_FecEfecto.Caption = DateSerial((Mid(Lbl_FecEfecto.Caption, 1, 4)), (Mid(Lbl_FecEfecto.Caption, 5, 2)), (Mid(Lbl_FecEfecto.Caption, 7, 2)))
       End If
       
       If IsNull(vgRs!FEC_RECCIA) Then
          Txt_FecRec.Text = ""
       Else
           Txt_FecRec.Text = (vgRs!FEC_RECCIA)
           Txt_FecRec.Text = DateSerial((Mid(Txt_FecRec.Text, 1, 4)), (Mid(Txt_FecRec.Text, 5, 2)), (Mid(Txt_FecRec.Text, 7, 2)))
       End If
       
       Call fgBuscarPosicionCodigoCombo(Trim(vgRs!cod_tipoidentut), Cmb_NumIdent)
       Txt_NumIdent.Text = Trim(vgRs!num_identut)
''       Txt_Digito.Text = Trim(vgRs!dgv_tut)
       Txt_Nombre.Text = Trim(vgRs!gls_nomtut)
       If IsNull(vgRs!gls_nomsegtut) Then
           Txt_SegNombre.Text = ""
       Else
           Txt_SegNombre.Text = Trim(vgRs!gls_nomsegtut)
       End If
       Txt_Paterno.Text = Trim(vgRs!gls_pattut)
       If IsNull(vgRs!gls_mattut) Then
           Txt_Materno.Text = ""
       Else
           Txt_Materno.Text = Trim(vgRs!gls_mattut)
       End If
       
       Txt_Direccion.Text = Trim(vgRs!gls_dirtut)
       
       If IsNull(vgRs!gls_fonotut) Then
          Txt_Telefono.Text = ""
       Else
           Txt_Telefono.Text = Trim(vgRs!gls_fonotut)
       End If
       
       If IsNull(vgRs!gls_correotut) Then
            Txt_Email.Text = ""
       Else
            Txt_Email.Text = Trim(vgRs!gls_correotut)
       End If
                 
''       vlCont = 0
''       Do While vlCont <= Cmb_Comuna.ListCount
''             If Cmb_Comuna.ItemData(vlCont) = (vgRs!Cod_Direccion) Then
''                Cmb_Comuna.ListIndex = vlCont
''                vlCont = Cmb_Comuna.ListCount + 1
''                Exit Do
''             End If
''             vlCont = vlCont + 1
''       Loop
          
       vlCodDir = (vgRs!Cod_Direccion)
       Call fgBuscarNombreProvinciaRegion(vlCodDir)
       vlNombreRegion = vgNombreRegion
       vlNombreProvincia = vgNombreProvincia
       vlNombreComuna = vgNombreComuna
       
       Lbl_Departamento.Caption = vlNombreRegion
       Lbl_Provincia.Caption = vlNombreProvincia
       Lbl_Distrito.Caption = vlNombreComuna
       
       Cmb_ViaPago.ListIndex = fgBuscarPosicionCodigoCombo(Trim(vgRs!Cod_ViaPago), Cmb_ViaPago)
       Cmb_TipoCta.ListIndex = fgBuscarPosicionCodigoCombo(Trim(vgRs!Cod_TipCuenta), Cmb_TipoCta)
       Cmb_Banco.ListIndex = fgBuscarPosicionCodigoCombo(Trim(vgRs!Cod_Banco), Cmb_Banco)
                
       vlcont = 0
       Cmb_Sucursal.ListIndex = 0
       Do While vlcont < Cmb_Sucursal.ListCount
          If (Trim(Cmb_Sucursal) <> "") Then
             If (vgRs!Cod_Sucursal = Trim(Mid(Cmb_Sucursal.Text, 1, (InStr(1, Cmb_Sucursal, "-") - 1)))) Then
                Exit Do
             End If
          End If
          vlcont = vlcont + 1
          'I---- ABV 23/08/2004 ---
          If (vlcont < Cmb_Sucursal.ListCount) Then
            Cmb_Sucursal.ListIndex = vlcont
          End If
          'F---- ABV 23/08/2004 ---
       Loop
       
       'I---- ABV 23/08/2004 ---
       'Cmb_Sucursal.ListIndex = vlCont
        If (vlcont >= Cmb_Sucursal.ListCount) Then
            Cmb_Sucursal.ListIndex = 0
        End If
       'F---- ABV 23/08/2004 ---
       
       If IsNull(vgRs!Num_Cuenta) Then
          Txt_Cuenta.Text = ""
       Else
           Txt_Cuenta.Text = Trim(vgRs!Num_Cuenta)
       End If
'       vgRs2.Close
'CMV/20041102-I
       If (vlSwSinDerPen = "S") Then
       'Consultar Datos - Deshabilitar Controles
          Fra_Poliza.Enabled = False
          
          Fra_Vigencia.Enabled = False
          Fra_Antecedentes.Enabled = False
    
          Cmb_ViaPago.Enabled = False
          Cmb_Sucursal.Enabled = False
          Cmb_TipoCta.Enabled = False
          Cmb_Banco.Enabled = False
          Txt_Cuenta.Enabled = False
       Else
       'Modificar Datos
           Call flHabilitarIngreso
       End If
       
'CMV/20041102-F
       Call Cmb_ViaPago_Click
       
'''''       Call flCargaGrillaTutor
       
    Else
'        vgRs2.Close
        MsgBox "El Beneficiario ingresado No tiene ningún Tutor Asociado", vbInformation, "Información"
        If (vlSwSinDerPen = "S") Then
           Fra_Poliza.Enabled = False
        
           Fra_Vigencia.Enabled = False
           Fra_Antecedentes.Enabled = False
       
           Cmb_ViaPago.Enabled = False
           Cmb_Sucursal.Enabled = False
           Cmb_TipoCta.Enabled = False
           Cmb_Banco.Enabled = False
           Txt_Cuenta.Enabled = False
           SSTab_Tutor.Tab = 1
           SSTab_Tutor.Enabled = False
        Else
            Txt_FecRec.Text = fgBuscaFecServ
            Call flHabilitarIngreso
            Call Cmb_ViaPago_Click
        End If
           
    End If
          
    If Txt_Meses.Enabled = True Then
       Txt_Meses.SetFocus
    End If
    
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

Private Sub Cmd_Cancelar_Click()
On Error GoTo Err_CmdCancelarClick
    
    Call flIniciaForm
    SSTab_Tutor.Tab = 0
    Txt_PenPoliza.SetFocus
    Set RS = New ADODB.Recordset 'no borrar
    Cmb_PenNumIdent.Enabled = True
    Txt_PenNumIdent.Enabled = True

Exit Sub
Err_CmdCancelarClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_CargaTutor_Click()

On Error GoTo Err_CmdCargaTutor

    Cmd_BuscarPol_Click
    
Exit Sub
Err_CmdCargaTutor:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub

Private Sub Cmd_Eliminar_Click()
On Error GoTo Err_CmdEliminarClick
    
    
If vlSwModificar = "S" Then
    
    'Validar que se encuentran los datos que identifican al Pensionado
    If (Trim(Txt_PenPoliza) = "") Then
        MsgBox "Debe indicar la Póliza a la cual corresponde el Tutor", vbCritical, "Error de Datos"
        Exit Sub
    End If
    If (Trim(Cmb_PenNumIdent) = "") Then
        MsgBox "Debe indicar el Tipo de Identificación del Pensionado al cual corresponde el Tutor", vbCritical, "Error de Datos"
        Exit Sub
    End If
    If (Trim(Txt_PenNumIdent) = "") Then
        MsgBox "Debe indicar el Número de Identificación del Pensionado al cual corresponde el Tutor", vbCritical, "Error de Datos"
        Exit Sub
    End If
        
    vgSql = "SELECT num_poliza,num_endoso,num_orden "
    vgSql = vgSql & "FROM PP_TMAE_TUTOR WHERE "
    vgSql = vgSql & "num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
    'I---- ABV 21/08/2004 ---
    'vgSql = vgSql & "num_endoso = " & Trim(vlNumEndoso) & " AND "
    vgSql = vgSql & "num_orden = " & Trim(vlNumOrden) & " "
    'F---- ABV 21/08/2004 ---
    Set vgRs2 = vgConexionBD.Execute(vgSql)
    If Not vgRs2.EOF Then
       Screen.MousePointer = 11
       vlNumero = MsgBox("¿Desea Eliminar el Registro Seleccionado ?", vbQuestion + vbYesNo + 256, "Confirmación")
       If vlNumero <> 6 Then
          'Call Cmd_Limpiar_Click
       Else
           Sql = "DELETE PP_TMAE_TUTOR WHERE "
           Sql = Sql & "num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
           'I---- ABV 21/08/2004 ---
           'Sql = Sql & "num_endoso = " & Trim(vlNumEndoso) & "' AND "
           Sql = Sql & "num_orden = '" & Trim(vlNumOrden) & "'"
           'I---- ABV 21/08/2004 ---
           vgConexionBD.Execute Sql
           
           MsgBox "La Eliminación de Datos fue realizada Satisfactoriamente.", vbInformation, "Información"
           Call Cmd_Limpiar_Click
       End If
    Else
        MsgBox "El Registro No Existe en la Base de Datos.", vbInformation, "Información"
    End If
    
    'Call Cmd_Cancelar_Click
    Screen.MousePointer = 0
    
Else
    MsgBox "Sólo puede Consultar los Datos que se encuentran en Pantalla.", vbInformation, "Información"
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


If vlSwModificar = "S" Then

    'Validar que se encuentran los datos que identifican al Pensionado
    If (Trim(Txt_PenPoliza) = "") Then
        MsgBox "Debe indicar la Póliza a la cual corresponde el Tutor", vbCritical, "Error de Datos"
        Exit Sub
    End If
    If (Trim(Cmb_PenNumIdent) = "") Then
        MsgBox "Debe indicar el Tipo de Identificación del Pensionado al cual corresponde el Tutor", vbCritical, "Error de Datos"
        Exit Sub
    End If
    If (Trim(Txt_PenNumIdent) = "") Then
        MsgBox "Debe indicar el Número de Identificación del Pensionado al cual corresponde el Tutor", vbCritical, "Error de Datos"
        Exit Sub
    End If

    'Valida Duración en meses
    If Trim(Txt_Meses.Text) = "" Then
       MsgBox "Debe Ingresar Número de Meses.", vbCritical, "Error de Datos"
       Lbl_Hasta.Caption = ""
       Txt_Meses.SetFocus
       Exit Sub
    End If
    If CDbl(Txt_Meses.Text) > clMesesMax Then
       MsgBox "El Número de Meses Ingresado Excede el Máximo Permitido de 1200.", vbCritical, "Error de Datos"
       Lbl_Hasta.Caption = ""
       Txt_Meses.SetFocus
       Exit Sub
    End If
'Valida Fecha inicial (Fecha Desde)
    If (Trim(Txt_Desde) = "") Then
       MsgBox "Debe ingresar una Fecha para el Valor Desde", vbCritical, "Error de Datos"
       Lbl_Hasta.Caption = ""
       Txt_Desde.SetFocus
       Exit Sub
    End If
    If Not IsDate(Txt_Desde.Text) Then
       MsgBox "La Fecha ingresada no es una fecha válida.", vbCritical, "Error de Datos"
       Lbl_Hasta.Caption = ""
       Txt_Desde.SetFocus
       Exit Sub
    End If
    If (CDate(Txt_Desde) > CDate(Date)) Then
       MsgBox "La Fecha ingresada es mayor a la fecha actual", vbCritical, "Error de Datos"
       Lbl_Hasta.Caption = ""
       Txt_Desde.SetFocus
       Exit Sub
    End If
    If (Year(Txt_Desde) < 1900) Then
       MsgBox "La Fecha ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
       Lbl_Hasta.Caption = ""
       Txt_Desde.SetFocus
       Exit Sub
    End If
    
    If Not fgValidaVigenciaPoliza(Trim(Txt_PenPoliza.Text), Txt_Desde.Text) Then
       MsgBox " La Fecha Desde Ingresada es Anterior a la Fecha de Vigencia de la Póliza ", vbInformation, "Información"
       Lbl_Hasta.Caption = ""
       Txt_Desde.SetFocus
       Exit Sub
    End If
    
    
    'I---- ABV 28/08/2004 ---
    'vlFecha = Format(CDate(Txt_Desde.Text), "yyyymmdd")
    'vlAnno = Mid(Trim(vlFecha), 1, 4)
    'vlMes = Mid(Trim(vlFecha), 5, 2)
    'vlDia = Mid(Trim(vlFecha), 7, 2)
    'Lbl_Hasta.Caption = DateSerial(vlAnno, vlMes + CDbl(Txt_Meses), vlDia)
    Lbl_Hasta.Caption = flCalcularFechaHasta(Txt_Desde, CDbl(Txt_Meses))
    Lbl_FecEfecto.Caption = fgValidaFechaEfecto(Txt_Desde.Text, Trim(Txt_PenPoliza), vlNumOrden)
    'I---- ABV 28/08/2004 ---

    
    
'Valida Fecha Recepción
    If (Trim(Txt_FecRec) = "") Then
       MsgBox "Debe Ingresar Fecha Recepción", vbCritical, "Error de Datos"
       Txt_FecRec.SetFocus
       Exit Sub
    End If
    If Not IsDate(Txt_FecRec.Text) Then
       MsgBox "La Fecha Ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
       Txt_FecRec.SetFocus
       Exit Sub
    End If
    If (CDate(Txt_FecRec) > CDate(Date)) Then
       MsgBox "La Fecha Ingresada es Mayor a la Fecha Actual", vbCritical, "Error de Datos"
       Txt_FecRec.SetFocus
       Exit Sub
    End If
    If (Year(Txt_FecRec) < 1900) Then
       MsgBox "La Fecha Ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
       Txt_FecRec.SetFocus
       Exit Sub
    End If
    
    If Not fgValidaVigenciaPoliza(Trim(Txt_PenPoliza.Text), Txt_FecRec.Text) Then
       MsgBox " La Fecha Recepción Ingresada es Anterior a la Fecha de Vigencia de la Póliza ", vbInformation, "Información"
       Txt_FecRec.SetFocus
       Exit Sub
    End If
    
    
    'Valida la Identificación del Tutor
    If (Trim(Cmb_NumIdent) = "") Then
        MsgBox "Debe indicar el Tipo de Identificación del Tutor", vbCritical, "Error de Datos"
        Cmb_NumIdent.SetFocus
        Exit Sub
    End If
    If (Trim(Txt_NumIdent.Text)) = "" Then
       MsgBox "Debe ingresar el Número de Identificación del Tutor.", vbCritical, "Error de Datos"
       Txt_NumIdent.SetFocus
       Exit Sub
    Else
        'Txt_Rut = Format(Txt_Rut, "##,###,##0")
        Txt_NumIdent = UCase(Trim(Txt_NumIdent))
        'Txt_Digito.SetFocus
    End If
''    If Txt_Digito.Text = "" Then
''       MsgBox "Debe Ingresar Dígito Verificador de Rut.", vbCritical, "Error de Datos"
''       Txt_Digito.SetFocus
''       Exit Sub
''    End If
''    If Not ValiRut(Txt_Rut.Text, Txt_Digito.Text) Then
''       MsgBox "El Dígito Verificador del Rut ingresado es incorrecto.", vbCritical, "Error de Datos"
''       Txt_Digito.SetFocus
''       Exit Sub
''    End If
    'Valida Nombre Tutor
    If Txt_Nombre.Text = "" Then
       MsgBox "Debe Ingresar Nombre.", vbCritical, "Error de Datos"
       Txt_Nombre.SetFocus
       Exit Sub
    Else
        Txt_Nombre.Text = UCase(Txt_Nombre.Text)
    End If
    'Valida Apellido Paterno Tutor
    If Txt_Paterno.Text = "" Then
       MsgBox "Debe Ingresar Apellido Paterno.", vbCritical, "Error de Datos"
       Txt_Paterno.SetFocus
       Exit Sub
    Else
        Txt_Paterno.Text = UCase(Txt_Paterno.Text)
    End If
    'Valida Apellido Materno Tutor
    If Txt_Materno.Text = "" Then
    'Ya no es obligatorio
''       MsgBox "Debe Ingresar Apellido Materno.", vbCritical, "Error de Datos"
''       Txt_Materno.SetFocus
''       Exit Sub
    Else
        Txt_Materno.Text = UCase(Txt_Materno.Text)
    End If
    'Valida Dirección Tutor
    If Txt_Direccion.Text = "" Then
       MsgBox "Debe Ingresar Dirección.", vbCritical, "Error de Datos"
       Txt_Direccion.SetFocus
       Exit Sub
    Else
        Txt_Direccion.Text = UCase(Txt_Direccion.Text)
    End If
    
    
    'Validar Ingreso de Datos para Pago, según Via de Pago Seleccionada
    vlNumero = InStr(Cmb_ViaPago.Text, "-")
    vlOpcion = Trim(Mid(Cmb_ViaPago.Text, 1, vlNumero - 1))
    'Opción 01 = Via de Pago CAJA
           
    'Opción 00 = Via de Pago - Sin Información
    If vlOpcion <> "00" Then
    
       If vlOpcion = "01" Or vlOpcion = "04" Then
          vlNumero = InStr(Cmb_Sucursal.Text, "-")
          If Trim(Mid(Cmb_Sucursal.Text, 1, vlNumero - 1)) = "0000" Then
             MsgBox "Debe Seleccionar Sucursal/AFP de Pago", vbCritical, "Error de Datos"
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
              If Txt_Cuenta.Text = "" Then
                 MsgBox "Debe Ingresar Número de Cuenta para Pago", vbCritical, "Error de Datos"
                 Txt_Cuenta.SetFocus
                 Exit Sub
              End If
           Else
              If (vlOpcion = "05") Then
              Else
            'Via de Pago <> = Via de Pago <>
               'ACTIVAR TODO e ingresar todo sin validar
               Cmb_Sucursal.Enabled = True
               Cmb_TipoCta.Enabled = True
               Cmb_Banco.Enabled = True
               Txt_Cuenta.Enabled = True
'               Cmb_Sucursal.ListIndex = 0
'               Cmb_TipoCta.ListIndex = 0
'               Cmb_Banco.ListIndex = 0
'               Txt_Cuenta.Text = ""
              End If
           End If
       End If
    Else
        MsgBox "Debe Seleccionar Forma de Pago.", vbCritical, "Error de Datos"
        Exit Sub
    End If
    
   
    vlFechaInicio = Format(CDate(Trim(Txt_Desde.Text)), "yyyymmdd")
    vlFechaFin = Format(CDate(Trim(Lbl_Hasta.Caption)), "yyyymmdd")
    ''vlRut = Format(Txt_Rut, "#0")
    vlCodTipoIdenBenTut = fgObtenerCodigo_TextoCompuesto(Cmb_NumIdent)
    vlNumIdenBenTut = Trim(UCase(Txt_NumIdent))
    
    vgSql = ""
    vgSql = "SELECT num_poliza,num_endoso,num_orden "
    vgSql = vgSql & "FROM PP_TMAE_TUTOR WHERE "
    vgSql = vgSql & " num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
    'I---- ABV 21/08/2004 ---
    'vgSql = vgSql & " num_endoso = " & Trim(vlNumEndoso) & " AND "
    vgSql = vgSql & " num_orden = " & Trim(vlNumOrden) & " "
    'F---- ABV 21/08/2004 ---
    Set vgRs2 = vgConexionBD.Execute(vgSql)
    If Not vgRs2.EOF Then

        vgRes = MsgBox("¿ Está seguro que desea Modificar los Datos ?", 4 + 32 + 256, "Operación de Actualización")
        If vgRes <> 6 Then
            vgRs2.Close
            Cmd_Grabar.SetFocus
            Screen.MousePointer = 0
            Exit Sub
        End If
       
       vlGlsUsuarioModi = vgUsuario
       vlFecModi = Format(Date, "yyyymmdd")
       vlHorModi = Format(Time, "hhmmss")
       
       Sql = ""
       Sql = " UPDATE PP_TMAE_TUTOR SET num_identut = '" & str(Txt_NumIdent.Text) & "', "
       Sql = Sql & " cod_tipoidentut = " & vlCodTipoIdenBenTut & ", "
       Sql = Sql & " gls_nomtut = '" & Trim(Txt_Nombre.Text) & "', "
       If (Trim(Txt_SegNombre.Text) = "") Then
          Sql = Sql & " gls_nomsegtut = NULL , "
       Else
           Sql = Sql & " gls_nomsegtut = '" & Txt_SegNombre.Text & "', "
       End If
       Sql = Sql & " gls_pattut = '" & Trim(Txt_Paterno.Text) & "', "
       If (Trim(Txt_Materno.Text) = "") Then
           Sql = Sql & " gls_mattut = NULL , "
       Else
           Sql = Sql & " gls_mattut = '" & Trim(Txt_Materno.Text) & "', "
       End If
       Sql = Sql & " gls_dirtut = '" & Trim(Txt_Direccion.Text) & "', "
       Sql = Sql & " cod_direccion = " & vlCodDir & ", "
       
       If (Trim(Txt_Telefono.Text) = "") Then
          Sql = Sql & " gls_fonotut = NULL , "
       Else
           Sql = Sql & " gls_fonotut = '" & Txt_Telefono.Text & "', "
       End If
       If (Trim(Txt_Email.Text) = "") Then
          Sql = Sql & " gls_correotut = NULL, "
       Else
           Sql = Sql & " gls_correotut = '" & Txt_Email.Text & "', "
       End If
       
       Sql = Sql & " num_mespodnot = " & Txt_Meses.Text & ", "
       Sql = Sql & " fec_inipodnot = '" & Trim(vlFechaInicio) & "', "
       Sql = Sql & " fec_terpodnot = '" & Trim(vlFechaFin) & "', "
       
       vlNumero = InStr(Cmb_ViaPago.Text, "-")
       vlViaPago = Trim(Mid(Cmb_ViaPago.Text, 1, vlNumero - 1))
       
       Sql = Sql & " cod_viapago = '" & Trim(vlViaPago) & "', "
       
       vlNumero = InStr(Cmb_TipoCta.Text, "-")
       Sql = Sql & " cod_tipcuenta = '" & Trim(Mid(Cmb_TipoCta.Text, 1, vlNumero - 1)) & "', "
      
       vlNumero = InStr(Cmb_Banco.Text, "-")
       Sql = Sql & " cod_banco = '" & Trim(Mid(Cmb_Banco.Text, 1, vlNumero - 1)) & "', "
       
       If Txt_Cuenta.Enabled = True Then
          Sql = Sql & " num_cuenta = '" & Trim(Txt_Cuenta.Text) & "', "
       Else
           Sql = Sql & " num_cuenta = NULL, "
       End If
       
       vlNumero = InStr(Cmb_Sucursal.Text, "-")
       Sql = Sql & " cod_sucursal = '" & Trim(Mid(Cmb_Sucursal.Text, 1, vlNumero - 1)) & "', "
       
       Sql = Sql & "fec_efecto = '" & Trim(Format(Lbl_FecEfecto, "yyyymmdd")) & "', "
       Sql = Sql & "fec_reccia = '" & Trim(Format(Txt_FecRec, "yyyymmdd")) & "', "
               
       Sql = Sql & " cod_usuariomodi = '" & vlGlsUsuarioModi & "', "
       Sql = Sql & " fec_modi = '" & vlFecModi & "', "
       Sql = Sql & " hor_modi = '" & vlHorModi & "' "

       Sql = Sql & " WHERE "
       Sql = Sql & "num_poliza = '" & Trim(Txt_PenPoliza.Text) & "' AND "
       'Sql = Sql & " num_endoso = " & Trim(vlNumEndoso) & " AND "
       Sql = Sql & " num_orden = " & Trim(vlNumOrden) & " "
       
       
       vgConexionBD.Execute Sql
       
       MsgBox "Los Datos han sido actualizados Satisfactoriamente", vbInformation, "Información"
       
       Txt_Meses.SetFocus

    Else
        
        vlGlsUsuarioCrea = vgUsuario
        vlFecCrea = Format(Date, "yyyymmdd")
        vlHorCrea = Format(Time, "hhmmss")
 
        vlGlsUsuarioModi = Null
        vlFecModi = Null
        vlHorModi = Null
        
        vgSql = ""
        vgSql = "INSERT INTO PP_TMAE_TUTOR "
        vgSql = vgSql & "(num_poliza,num_endoso,num_orden,cod_tipoidentut,num_identut, "
        vgSql = vgSql & " gls_nomtut,gls_nomsegtut,gls_pattut,gls_mattut,gls_dirtut,cod_direccion, "
        vgSql = vgSql & " gls_fonotut,gls_correotut,num_mespodnot,fec_inipodnot,fec_terpodnot,"
        vgSql = vgSql & " cod_viapago,cod_tipcuenta,cod_banco,num_cuenta,cod_sucursal, "
        vgSql = vgSql & " fec_efecto,fec_reccia, "
        vgSql = vgSql & " cod_usuariocrea,fec_crea,hor_crea,cod_usuariomodi,fec_modi,hor_modi "
        vgSql = vgSql & " ) VALUES ( "
        vgSql = vgSql & " '" & Trim(Txt_PenPoliza) & "' , "
        vgSql = vgSql & " " & vlNumEndoso & ", "
        vgSql = vgSql & " " & vlNumOrden & ", "
        vgSql = vgSql & " " & vlCodTipoIdenBenTut & ", "
        vgSql = vgSql & " '" & vlNumIdenBenTut & "', "
        vgSql = vgSql & " '" & Trim(Txt_Nombre.Text) & "', "
        If (Trim(Txt_SegNombre.Text) = "") Then
           vgSql = vgSql & " NULL, "
        Else
            vgSql = vgSql & " '" & Trim(Txt_SegNombre.Text) & "', "
        End If
        vgSql = vgSql & " '" & Trim(Txt_Paterno.Text) & "', "
        If (Trim(Txt_Materno.Text) = "") Then
           vgSql = vgSql & " NULL, "
        Else
            vgSql = vgSql & " '" & Trim(Txt_Materno.Text) & "', "
        End If
        vgSql = vgSql & " '" & Trim(Txt_Direccion.Text) & "', "
        vgSql = vgSql & " " & vlCodDir & ", "
        
        If (Trim(Txt_Telefono.Text) = "") Then
           vgSql = vgSql & " NULL, "
        Else
            vgSql = vgSql & " '" & Txt_Telefono.Text & "', "
        End If
        If (Trim(Txt_Email.Text) = "") Then
           vgSql = vgSql & " NULL, "
        Else
            vgSql = vgSql & " '" & Txt_Email.Text & "', "
        End If
        
        vgSql = vgSql & " '" & Trim(str(Txt_Meses.Text)) & "', "
        vgSql = vgSql & " '" & Trim(str(vlFechaInicio)) & "', "
        vgSql = vgSql & " '" & Trim(str(vlFechaFin)) & "', "
                  
        vlNumero = InStr(Cmb_ViaPago.Text, "-")
        vlViaPago = Trim(Mid(Cmb_ViaPago.Text, 1, vlNumero - 1))
    
        vgSql = vgSql & " '" & Trim(Mid(Cmb_ViaPago.Text, 1, vlNumero - 1)) & "', "
        
        vlNumero = InStr(Cmb_TipoCta.Text, "-")
        vgSql = vgSql & " '" & Trim(Mid(Cmb_TipoCta.Text, 1, vlNumero - 1)) & "', "
           
        vlNumero = InStr(Cmb_Banco.Text, "-")
        vgSql = vgSql & " '" & Trim(Mid(Cmb_Banco.Text, 1, vlNumero - 1)) & "', "
        
        If Txt_Cuenta.Enabled = True Then
           vgSql = vgSql & " '" & Trim(Txt_Cuenta.Text) & "', "
        Else
            vgSql = vgSql & " NULL, "
        End If
        
        vlNumero = InStr(Cmb_Sucursal.Text, "-")
        vgSql = vgSql & "'" & Trim(Mid(Cmb_Sucursal.Text, 1, vlNumero - 1)) & "', "
        
        vgSql = vgSql & "'" & Trim(Format(Lbl_FecEfecto, "yyyymmdd")) & "', "
        vgSql = vgSql & "'" & Trim(Format(Txt_FecRec, "yyyymmdd")) & "', "
        
        vgSql = vgSql & "'" & vlGlsUsuarioCrea & "', "
        vgSql = vgSql & "'" & vlFecCrea & "', "
        vgSql = vgSql & "'" & vlHorCrea & "', "
        vgSql = vgSql & "'" & vlGlsUsuarioModi & "', "
        vgSql = vgSql & "'" & vlFecModi & "', "
        vgSql = vgSql & "'" & vlHorModi & "' ) "
                        
        vgConexionBD.Execute vgSql
        
        MsgBox "Los Datos han Sido Ingresados Satisfactoriamente.", vbInformation, "Información"
        
        Txt_Meses.SetFocus
        
    End If
    
'''''    Call flCargaGrillaTutor
    
    
Else
    MsgBox "Sólo puede Consultar los Datos que se encuentran en Pantalla.", vbInformation, "Información"
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
Dim vlArchivo As String

On Error GoTo Err_CmdImprimir
   
    'Validar que se encuentran los datos que identifican al Pensionado
    If (Trim(Txt_PenPoliza) = "") Then
        MsgBox "Debe ingresar el Nº de Póliza.", vbCritical, "Error de Datos"
        Exit Sub
    End If
    If (Trim(Cmb_PenNumIdent) = "") Then
        MsgBox "Debe ingresar el Tipo de Identificación del Pensionado.", vbCritical, "Error de Datos"
        Exit Sub
    End If
    If (Trim(Txt_PenNumIdent) = "") Then
        MsgBox "Debe ingresar el Número de Identificación del Pensionado.", vbCritical, "Error de Datos"
        Exit Sub
    End If
       
   Screen.MousePointer = 11
   
   vlArchivo = strRpt & "PP_Rpt_AntTutor.rpt"   '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
       MsgBox "Archivo de Reportes no se encuentra en el Directorio de la Aplicación.", 16, "Archivo no encontrado"
       Screen.MousePointer = 0
       Exit Sub
   End If

   vgQuery = " {ado.NUM_POLIZA} = '" & Trim(Txt_PenPoliza.Text) & "' AND "
   'vgQuery = vgQuery & "{PP_TMAE_BEN.NUM_ENDOSO} = " & Trim(vlNumEndoso) & " AND "
   vgQuery = vgQuery & " {ado.NUM_ORDEN} = " & Trim(vlNumOrden) & ""
   
   Rpt_General.Reset
   Rpt_General.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
   Rpt_General.Connect = vgRutaDataBase       ' o App.Path & "\Nestle.mdb"
   Rpt_General.SelectionFormula = vgQuery
   Rpt_General.Formulas(0) = ""
   Rpt_General.Formulas(1) = ""
   Rpt_General.Formulas(2) = ""
   Rpt_General.Formulas(3) = ""
   
   Rpt_General.Formulas(1) = "NombreCompania = '" & vgNombreCompania & "'"
   Rpt_General.Formulas(2) = "NombreSistema = '" & vgNombreSistema & "'"
   Rpt_General.Formulas(3) = "NombreSubSistema = '" & vgNombreSubSistema & "'"
   
   Rpt_General.WindowState = crptMaximized
   Rpt_General.Destination = crptToWindow
   Rpt_General.WindowTitle = "Informe de Tutores"
   Rpt_General.Action = 1
   
   Screen.MousePointer = 0
   
Exit Sub
Err_CmdImprimir:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Limpiar_Click()
On Error GoTo Err_CmdLimpiarClick

 If Fra_Poliza.Enabled = False Then
 
    Call flHabilitarIngreso

    Txt_Meses.Text = ""
    Txt_Desde.Text = ""
    Lbl_Hasta.Caption = ""
    Txt_FecRec.Text = fgBuscaFecServ
    
    If (Cmb_NumIdent.ListCount > 0) Then
        Cmb_NumIdent.ListIndex = 0
    End If
    Txt_NumIdent.Text = ""
    ''Txt_Digito.Text = ""
    Txt_Nombre.Text = ""
    Txt_SegNombre.Text = ""
    Txt_Paterno.Text = ""
    Txt_Materno.Text = ""
    Txt_Direccion.Text = ""
    
''    If (Cmb_Comuna.ListCount > 0) Then
''        Cmb_Comuna.ListIndex = 0
''    End If
    Lbl_Departamento.Caption = ""
    Lbl_Provincia.Caption = ""
    Lbl_Distrito.Caption = ""
    
    Txt_Telefono.Text = ""
    Txt_Email.Text = ""
   
    Cmb_ViaPago.ListIndex = 0
    Cmb_Sucursal.ListIndex = 0
    Cmb_TipoCta.ListIndex = 0
    Cmb_Banco.ListIndex = 0
    
    vlNumero = InStr(Cmb_ViaPago.Text, "-")
    vlOpcion = Trim(Mid(Cmb_ViaPago.Text, 1, vlNumero - 1))
    'Opción 01 = Via de Pago CAJA
    If Cmb_ViaPago.Enabled = True Then
        If vlOpcion = "01" Then
            Cmb_Sucursal.Enabled = True
            Cmb_TipoCta.Enabled = False
            Cmb_Banco.Enabled = False
            Txt_Cuenta.Enabled = False
        Else
            Cmb_Sucursal.Enabled = False
            Cmb_TipoCta.Enabled = True
            Cmb_Banco.Enabled = True
            Txt_Cuenta.Enabled = True
        End If
    End If
    
    Txt_Cuenta.Text = ""
    
    Txt_FecRec.Text = ""
    Lbl_FecEfecto.Caption = ""
    
    'Lbl_FecRec.Caption = fgBuscaFecServ
    
    If Txt_Meses.Enabled = True Then
       Txt_Meses.SetFocus
    Else
        Txt_PenPoliza.SetFocus
    End If
    
    Call flInicializaGrillaTutores
    SSTab_Tutor.Tab = 0
    
 End If
    
Exit Sub
Err_CmdLimpiarClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub cmd_salir_Click()
On Error GoTo Err_CmdSalirClick

    Unload Me
    
Exit Sub
Err_CmdSalirClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_VerHistoria_Click()

On Error GoTo Err_Cmd_verHistoria

    Call flCargaGrillaTutor

Exit Sub
Err_Cmd_verHistoria:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub

Private Sub DataGrid1_DblClick()
Dim dato As String
On Error GoTo mierror
    dato = DataGrid1.Columns(1).Value
    Call fgBuscarPosicionCodigoCombo(dato, Cmb_PenNumIdent)
    Txt_PenNumIdent.Text = DataGrid1.Columns(2).Value
    Lbl_PenNombre.Caption = DataGrid1.Columns(3).Value
    Call Cmd_BuscarPol_Click
Exit Sub
mierror:
    MsgBox Err.Description, vbInformation
    
End Sub

Private Sub Form_Load()
On Error GoTo Err_Cargar

    Frm_AntTutores.Top = 0
    Frm_AntTutores.Left = 0
    
    'Cmb_Comuna.Clear
    'Call fgComboComuna(Cmb_Comuna)
    fgComboTipoIdentificacion Cmb_PenNumIdent
    fgComboTipoIdentificacion Cmb_NumIdent
    
    vlSw = True
    
    fgComboGeneral vgCodTabla_ViaPago, Cmb_ViaPago
    'Call fgComboSucursal(Cmb_Sucursal)
    Call fgComboSucursal(Cmb_Sucursal, "S")
    
    fgComboGeneral vgCodTabla_TipCta, Cmb_TipoCta
    fgComboGeneral vgCodTabla_Bco, Cmb_Banco
    
    vlSw = False
  
    Call flIniciaForm
    Set RS = New ADODB.Recordset
    
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Msf_GrillaTutor_Click()

On Error GoTo Err_MsfGrillaTutorClick

    Msf_GrillaTutor.Col = 0
    If (Msf_GrillaTutor.Text = "") Or (Msf_GrillaTutor.row = 0) Then
        MsgBox "No existen Detalles", vbExclamation, "Información"
        Exit Sub
    End If
    
Exit Sub
Err_MsfGrillaTutorClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub

Private Sub Txt_Cuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Cmd_Grabar.SetFocus
    End If
End Sub

Private Sub Txt_Cuenta_LostFocus()
Txt_Cuenta = Trim(Txt_Cuenta)
End Sub

Private Sub Txt_Desde_KeyPress(KeyAscii As Integer)
On Error GoTo Err_TxtDesdeKeyPress

     If KeyAscii = 13 Then
        Txt_Desde = Trim(Txt_Desde)
        If (Trim(Txt_Desde) = "") Then
           MsgBox "Debe ingresar una Fecha para el Valor Desde", vbCritical, "Error de Datos"
           Lbl_Hasta.Caption = ""
           Txt_Desde.SetFocus
           Exit Sub
        End If
        If Not IsDate(Txt_Desde.Text) Then
           MsgBox "La Fecha ingresada no es una fecha válida.", vbCritical, "Error de Datos"
           Lbl_Hasta.Caption = ""
           Txt_Desde.SetFocus
           Exit Sub
        End If
        If (CDate(Txt_Desde) > CDate(Date)) Then
           MsgBox "La Fecha ingresada es mayor a la fecha actual", vbCritical, "Error de Datos"
           Lbl_Hasta.Caption = ""
           Txt_Desde.SetFocus
           Exit Sub
        End If
        If (Year(Txt_Desde) < 1900) Then
           MsgBox "La Fecha ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
           Lbl_Hasta.Caption = ""
           Txt_Desde.SetFocus
           Exit Sub
        End If
        
        Txt_Desde.Text = Format(CDate(Trim(Txt_Desde)), "yyyymmdd")
        Txt_Desde.Text = DateSerial(Mid((Txt_Desde.Text), 1, 4), Mid((Txt_Desde.Text), 5, 2), Mid((Txt_Desde.Text), 7, 2))
        
        'Valida Vigencia de la poliza según fecha ingresada en periodo de vigencia
        If Not fgValidaVigenciaPoliza(Trim(Txt_PenPoliza.Text), Txt_Desde.Text) Then
           MsgBox " La Fecha Ingresada es Anterior a la Fecha de Vigencia de la Póliza ", vbInformation, "Información"
           Lbl_Hasta.Caption = ""
           Txt_Desde.SetFocus
           Exit Sub
        End If
        
'        If Not fgValidaPagoPension(Txt_Desde.Text, Trim(Txt_PenPoliza), vlNumOrden) Then
'           MsgBox " Ya se ha Realizado el Proceso de Cálculo de Pensión para ésta Fecha ", vbCritical, "Operación Cancelada"
'           Lbl_Hasta.Caption = ""
'           Txt_Desde.SetFocus
'           Exit Sub
'        End If

        
        Lbl_Hasta.Caption = flCalcularFechaHasta(Txt_Desde, CInt(Txt_Meses))
        Call fgValidaFechaEfecto(Txt_Desde.Text, Trim(Txt_PenPoliza), vlNumOrden)
        Lbl_FecEfecto.Caption = vgFechaEfecto
'        Txt_FecRec.Text = fgBuscaFecServ
        
        
        
        'I---- ABV 28/08/2004 ---
        'vlFecha = Format(CDate(Txt_Desde.Text), "yyyymmdd")
        'vlAnno = Mid(Trim(vlFecha), 1, 4)
        'vlMes = Mid(Trim(vlFecha), 5, 2)
        'vlDia = Mid(Trim(vlFecha), 7, 2)
        'Lbl_Hasta.Caption = DateSerial(vlAnno, vlMes + CDbl(Txt_Meses), vlDia)
        'F---- ABV 28/08/2004 ---
        
        Txt_FecRec.SetFocus
     End If
     
Exit Sub
Err_TxtDesdeKeyPress:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Txt_Desde_LostFocus()
On Error GoTo Err_TxtDesdeLostFocus
        
    Txt_Desde = Trim(Txt_Desde)
    If (Txt_Meses.Text = "") Then
       Lbl_Hasta.Caption = ""
       Exit Sub
    End If
    If (Trim(Txt_Desde) = "") Then
       Lbl_Hasta.Caption = ""
       Exit Sub
    End If
    If Not IsDate(Txt_Desde.Text) Then
       Lbl_Hasta.Caption = ""
       Exit Sub
    End If
    If CDate(Txt_Desde) > CDate(Date) Then
       Lbl_Hasta.Caption = ""
       Exit Sub
    End If
    If Year(Txt_Desde) < 1900 Then
       Lbl_Hasta.Caption = ""
       Exit Sub
    End If
    
    Txt_Desde.Text = Format(CDate(Trim(Txt_Desde)), "yyyymmdd")
    Txt_Desde.Text = DateSerial(Mid((Txt_Desde.Text), 1, 4), Mid((Txt_Desde.Text), 5, 2), Mid((Txt_Desde.Text), 7, 2))
        
    If Not fgValidaVigenciaPoliza(Trim(Txt_PenPoliza.Text), Txt_Desde.Text) Then
       Lbl_Hasta.Caption = ""
       Exit Sub
    End If
    
    'I---- ABV 28/08/2004 ---
    'vlFecha = Format(CDate(Txt_Desde.Text), "yyyymmdd")
    'vlAnno = Mid(Trim(vlFecha), 1, 4)
    'vlMes = Mid(Trim(vlFecha), 5, 2)
    'vlDia = Mid(Trim(vlFecha), 7, 2)
    'Lbl_Hasta.Caption = DateSerial(vlAnno, vlMes + CDbl(Txt_Meses), vlDia)
    Lbl_Hasta.Caption = flCalcularFechaHasta(Txt_Desde, CInt(Txt_Meses))
    Call fgValidaFechaEfecto(Txt_Desde.Text, Trim(Txt_PenPoliza), vlNumOrden)
    Lbl_FecEfecto.Caption = vgFechaEfecto

'    Txt_FecRec.Text = fgBuscaFecServ
            
    
    'F---- ABV 28/08/2004 ---
    
Exit Sub
Err_TxtDesdeLostFocus:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Txt_Direccion_KeyPress(KeyAscii As Integer)
On Error GoTo Err_TxtDireccionKeyPress

    If KeyAscii = 13 Then
        Txt_Direccion = Trim(UCase(Txt_Direccion))
        If Txt_Direccion.Text = "" Then
            MsgBox "Debe Ingresar Dirección.", vbCritical, "Error de Datos"
            Txt_Direccion.SetFocus
            Exit Sub
        End If
        Cmd_BuscarDir.SetFocus
    End If
    
Exit Sub
Err_TxtDireccionKeyPress:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Txt_Direccion_LostFocus()
    Txt_Direccion.Text = Trim(UCase(Txt_Direccion.Text))
    If Txt_Direccion.Text = "" Then
       Exit Sub
    End If
End Sub

Private Sub Txt_Email_KeyPress(KeyAscii As Integer)
On Error GoTo Err_TxtEmailKeyPress

    If KeyAscii = 13 Then
        Txt_Email = Trim(Txt_Email)
        Cmb_ViaPago.SetFocus
    End If
    
Exit Sub
Err_TxtEmailKeyPress:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Txt_Email_LostFocus()
    Txt_Email = Trim(Txt_Email)
End Sub

Private Sub Txt_FecRec_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Txt_FecRec_KeyPress

     If KeyAscii = 13 Then
        Txt_FecRec = Trim(Txt_FecRec)
        If (Trim(Txt_FecRec) = "") Then
           MsgBox "Debe ingresar Fecha Recepción", vbCritical, "Error de Datos"
           Txt_FecRec.SetFocus
           Exit Sub
        End If
        If Not IsDate(Txt_FecRec.Text) Then
           MsgBox "La Fecha Ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
           Txt_FecRec.SetFocus
           Exit Sub
        End If
        If (CDate(Txt_FecRec) > CDate(Date)) Then
           MsgBox "La Fecha Ingresada es Mayor a la Fecha Actual", vbCritical, "Error de Datos"
           Txt_FecRec.SetFocus
           Exit Sub
        End If
        If (Year(Txt_FecRec) < 1900) Then
           MsgBox "La Fecha Ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
           Txt_FecRec.SetFocus
           Exit Sub
        End If
        
        Txt_FecRec.Text = Format(CDate(Trim(Txt_FecRec)), "yyyymmdd")
        Txt_FecRec.Text = DateSerial(Mid((Txt_FecRec.Text), 1, 4), Mid((Txt_FecRec.Text), 5, 2), Mid((Txt_FecRec.Text), 7, 2))
        
        If Not fgValidaVigenciaPoliza(Trim(Txt_PenPoliza.Text), Txt_FecRec.Text) Then
           MsgBox " La Fecha Ingresada es Anterior a la Fecha de Vigencia de la Póliza ", vbInformation, "Información"
           Txt_FecRec.SetFocus
           Exit Sub
        End If
        
        Cmb_NumIdent.SetFocus
     End If
     
Exit Sub
Err_Txt_FecRec_KeyPress:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Txt_FecRec_LostFocus()
On Error GoTo Err_Txt_FecRec_LostFocus
        
    Txt_FecRec = Trim(Txt_FecRec)
    If (Trim(Txt_FecRec) = "") Then
       Exit Sub
    End If
    If Not IsDate(Txt_FecRec.Text) Then
       Exit Sub
    End If
    If CDate(Txt_FecRec) > CDate(Date) Then
       Exit Sub
    End If
    If Year(Txt_FecRec) < 1900 Then
       Exit Sub
    End If
    
    Txt_FecRec.Text = Format(CDate(Trim(Txt_FecRec)), "yyyymmdd")
    Txt_FecRec.Text = DateSerial(Mid((Txt_FecRec.Text), 1, 4), Mid((Txt_FecRec.Text), 5, 2), Mid((Txt_FecRec.Text), 7, 2))
        
    If Not fgValidaVigenciaPoliza(Trim(Txt_PenPoliza.Text), Txt_FecRec.Text) Then
       Exit Sub
    End If
    
Exit Sub
Err_Txt_FecRec_LostFocus:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub

Private Sub Txt_Materno_KeyPress(KeyAscii As Integer)
On Error GoTo Err_TxtMaternoKeyPress

    If KeyAscii = 13 Then
       Txt_Materno.Text = Trim(UCase(Txt_Materno.Text))
       Txt_Direccion.SetFocus
    End If
    
Exit Sub
Err_TxtMaternoKeyPress:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Txt_Materno_LostFocus()
    Txt_Materno.Text = Trim(UCase(Txt_Materno.Text))
    If Txt_Materno.Text = "" Then
       Exit Sub
    End If
End Sub

Private Sub Txt_Meses_Change()
    If Not IsNumeric(Txt_Meses) Then
       Txt_Meses = ""
    End If
End Sub

Private Sub Txt_Meses_KeyPress(KeyAscii As Integer)
On Error GoTo Err_TxtMesesKeyPress

    If KeyAscii = 13 Then
       If Txt_Meses.Text = "" Then
          MsgBox "Debe Ingresar Número de Meses.", vbCritical, "Error de Datos"
          Lbl_Hasta.Caption = ""
          Txt_Meses.SetFocus
          Exit Sub
       End If
       If CDbl(Txt_Meses.Text) > clMesesMax Then
          MsgBox "El Número de Meses Ingresado Excede el Máximo Permitido de 1200.", vbCritical, "Error de Datos"
          Lbl_Hasta.Caption = ""
          Txt_Meses.SetFocus
          Exit Sub
       End If
       Txt_Desde.SetFocus
    End If
    
Exit Sub
Err_TxtMesesKeyPress:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Txt_Meses_LostFocus()
On Error GoTo Err_TxtMesesLostFocus
      
    Txt_Meses.Text = Trim(Txt_Meses.Text)
    If (Txt_Meses.Text = "") Then
       Lbl_Hasta.Caption = ""
       Exit Sub
    End If
'    If CDbl(Txt_Meses.Text > clMesesMax) Then
'       Lbl_Hasta.Caption = ""
'       Exit Sub
'    End If
    If (Trim(Txt_Desde) = "") Then
       Lbl_Hasta.Caption = ""
       Exit Sub
    End If
    If Not IsDate(Txt_Desde.Text) Then
       Lbl_Hasta.Caption = ""
       Exit Sub
    End If
    If CDate(Txt_Desde) > CDate(Date) Then
       Lbl_Hasta.Caption = ""
       Exit Sub
    End If
    If Year(Txt_Desde) < 1900 Then
       Lbl_Hasta.Caption = ""
       Exit Sub
    End If
    
    'I---- ABV 28/08/2004 ---
    'vlFecha = Format(CDate(Txt_Desde.Text), "yyyymmdd")
    'vlAnno = Mid(Trim(vlFecha), 1, 4)
    'vlMes = Mid(Trim(vlFecha), 5, 2)
    'vlDia = Mid(Trim(vlFecha), 7, 2)
    'Lbl_Hasta.Caption = DateSerial(vlAnno, vlMes + CDbl(Txt_Meses), vlDia)
    Lbl_Hasta.Caption = flCalcularFechaHasta(Txt_Desde, CInt(Txt_Meses))
    Lbl_FecEfecto.Caption = fgValidaFechaEfecto(Txt_Desde.Text, Trim(Txt_PenPoliza), vlNumOrden)
    'F---- ABV 28/08/2004 ---
        
Exit Sub
Err_TxtMesesLostFocus:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Txt_Nombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Txt_Nombre.Text = Trim(UCase(Txt_Nombre.Text))
       If Txt_Nombre.Text = "" Then
          MsgBox "Debe Ingresar Nombre.", vbCritical, "Error de Datos"
          Txt_Nombre.SetFocus
          Exit Sub
       End If
       Txt_SegNombre.SetFocus
    End If
End Sub

Private Sub Txt_Nombre_LostFocus()
    Txt_Nombre.Text = Trim(UCase(Txt_Nombre.Text))
    If Txt_Nombre.Text = "" Then
       Exit Sub
    End If
End Sub

Private Sub Txt_NumIdent_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If (Txt_NumIdent.Text <> "") Then
            Txt_NumIdent.Text = Trim(UCase(Txt_NumIdent.Text))
            Txt_Nombre.SetFocus
        Else
            Cmd_CargaTutor.SetFocus
        End If
    End If
End Sub

Private Sub Txt_NumIdent_LostFocus()
    Txt_NumIdent.Text = Trim(UCase(Txt_NumIdent.Text))
    If Txt_NumIdent.Text = "" Then
       Exit Sub
    End If
End Sub

Private Sub Txt_Paterno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Txt_Paterno.Text = Trim(UCase(Txt_Paterno.Text))
       If Txt_Paterno.Text = "" Then
          MsgBox "Debe Ingresar Apellido Paterno.", vbCritical, "Error de Datos"
          Txt_Paterno.SetFocus
          Exit Sub
       End If
       Txt_Materno.SetFocus
    End If
End Sub

Private Sub Txt_Paterno_LostFocus()
    Txt_Paterno.Text = Trim(UCase(Txt_Paterno.Text))
    If Txt_Paterno.Text = "" Then
       Exit Sub
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
        'Txt_PenNumIdent.SetFocus
        Cmd_BuscarPol.SetFocus
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

Private Sub FormateaGrilla()

On Error Resume Next
DataGrid1.Columns(0).Width = 689.9528
DataGrid1.Columns(1).Width = 945.0709
DataGrid1.Columns(2).Width = 1590.236
DataGrid1.Columns(3).Width = 3509.858
DataGrid1.Columns(4).Width = 1214.929
DataGrid1.Columns(5).Width = 540.2835
DataGrid1.Columns(6).Width = 989.8583
DataGrid1.Columns(7).Width = 1184.882

End Sub

Private Sub Txt_SegNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Txt_SegNombre.Text = Trim(UCase(Txt_SegNombre.Text))
       Txt_Paterno.SetFocus
    End If
End Sub

Private Sub Txt_SegNombre_LostFocus()
    Txt_SegNombre.Text = Trim(UCase(Txt_SegNombre.Text))
    If Txt_SegNombre.Text = "" Then
       Exit Sub
    End If
End Sub

Private Sub Txt_Telefono_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txt_Telefono = Trim(Txt_Telefono)
        Txt_Email.SetFocus
    End If
End Sub

Private Sub Txt_Telefono_LostFocus()
    Txt_Telefono = Trim(Txt_Telefono)
End Sub

Function flRecibeDireccion(iNomDepartamento As String, iNomProvincia As String, iNomDistrito As String, iCodDir As String)
'FUNCION QUE RECIBE LOS DATOS DEL FORMULARIO DE BUSQUEDA de Dirección
    
    Lbl_Departamento = Trim(iNomDepartamento)
    Lbl_Provincia = Trim(iNomProvincia)
    Lbl_Distrito = Trim(iNomDistrito)
    vlCodDir = iCodDir
    Txt_Telefono.SetFocus

End Function

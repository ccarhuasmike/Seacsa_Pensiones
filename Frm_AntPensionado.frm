VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "Crystl32.OCX"
Begin VB.Form Frm_AntPensionado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantención de Antecedentes del Pensionado."
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   9420
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
      Height          =   1095
      Left            =   120
      TabIndex        =   72
      Top             =   0
      Width           =   9165
      Begin VB.ComboBox Cmb_PenNumIdent 
         BackColor       =   &H00E0FFFF&
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   1
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
         Left            =   8400
         Picture         =   "Frm_AntPensionado.frx":0000
         TabIndex        =   5
         ToolTipText     =   "Buscar"
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton Cmd_BuscarPol 
         Height          =   375
         Left            =   8400
         Picture         =   "Frm_AntPensionado.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Buscar Póliza"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Txt_PenNumIdent 
         Height          =   285
         Left            =   5160
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
         TabIndex        =   109
         Top             =   0
         Width           =   1725
      End
      Begin VB.Label Lbl_End 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   7800
         TabIndex        =   3
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "N° End"
         Height          =   195
         Index           =   42
         Left            =   7200
         TabIndex        =   107
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Ident."
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   105
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Lbl_PenNombre 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   720
         Width           =   7455
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   74
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Póliza"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   73
         Top             =   360
         Width           =   855
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5985
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   10557
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Póliza"
      TabPicture(0)   =   "Frm_AntPensionado.frx":0204
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Lbl_Nombre(22)"
      Tab(0).Control(1)=   "Lbl_Nombre(23)"
      Tab(0).Control(2)=   "Lbl_Nombre(28)"
      Tab(0).Control(3)=   "Lbl_Nombre(24)"
      Tab(0).Control(4)=   "Lbl_Nombre(25)"
      Tab(0).Control(5)=   "Lbl_Nombre(30)"
      Tab(0).Control(6)=   "Lbl_Nombre(31)"
      Tab(0).Control(7)=   "Lbl_Nombre(33)"
      Tab(0).Control(8)=   "Lbl_Nombre(34)"
      Tab(0).Control(9)=   "Lbl_Nombre(27)"
      Tab(0).Control(10)=   "Lbl_Nombre(29)"
      Tab(0).Control(11)=   "Lbl_Nombre(26)"
      Tab(0).Control(12)=   "Lbl_Nombre(32)"
      Tab(0).Control(13)=   "Lbl_Nombre(35)"
      Tab(0).Control(14)=   "Lbl_Nombre(36)"
      Tab(0).Control(15)=   "Lbl_NumEndoso"
      Tab(0).Control(16)=   "Lbl_MesDif"
      Tab(0).Control(17)=   "Lbl_MesGar"
      Tab(0).Control(18)=   "Lbl_NumCar"
      Tab(0).Control(19)=   "Lbl_IniVig"
      Tab(0).Control(20)=   "Lbl_TerVig"
      Tab(0).Control(21)=   "Lbl_MtoPri"
      Tab(0).Control(22)=   "Lbl_MtoPen"
      Tab(0).Control(23)=   "Lbl_TasaCto"
      Tab(0).Control(24)=   "Lbl_TasaVta"
      Tab(0).Control(25)=   "Lbl_TasaRea"
      Tab(0).Control(26)=   "Lbl_TasaPerGar"
      Tab(0).Control(27)=   "Lbl_Afp"
      Tab(0).Control(28)=   "Lbl_TipPen"
      Tab(0).Control(29)=   "Lbl_Estado"
      Tab(0).Control(30)=   "Lbl_TipRta"
      Tab(0).Control(31)=   "Lbl_Mod"
      Tab(0).Control(32)=   "Lbl_Moneda(0)"
      Tab(0).Control(33)=   "Lbl_Moneda(1)"
      Tab(0).Control(34)=   "Lbl_Nombre(20)"
      Tab(0).Control(35)=   "Lbl_Nombre(52)"
      Tab(0).Control(36)=   "Lbl_FecEmision"
      Tab(0).Control(37)=   "Lbl_FecDevengue"
      Tab(0).Control(38)=   "Lbl_CUSPP"
      Tab(0).Control(39)=   "Lbl_Nombre(50)"
      Tab(0).Control(40)=   "Lbl_Nombre(37)"
      Tab(0).ControlCount=   41
      TabCaption(1)   =   "Beneficiarios"
      TabPicture(1)   =   "Frm_AntPensionado.frx":0220
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Msf_Grilla"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Fra_Personales"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Fra_Pago"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Fra_Salud"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.Frame Fra_Salud 
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
         Height          =   1875
         Left            =   4680
         TabIndex        =   51
         Top             =   4020
         Width           =   4335
         Begin VB.CommandButton Cmd_BenSalud 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3840
            TabIndex        =   35
            ToolTipText     =   "Ver Historial de Planes de Salud"
            Top             =   900
            Width           =   285
         End
         Begin VB.TextBox Txt_MtoPago 
            Height          =   285
            Left            =   1320
            MaxLength       =   16
            TabIndex        =   34
            Top             =   840
            Width           =   1500
         End
         Begin VB.ComboBox Cmb_ModPago 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   560
            Width           =   2835
         End
         Begin VB.ComboBox Cmb_Inst 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   240
            Width           =   2835
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   44
            Left            =   120
            TabIndex        =   112
            Top             =   0
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Monto de Pago"
            Height          =   255
            Index           =   21
            Left            =   120
            TabIndex        =   54
            Top             =   870
            Width           =   1155
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Mod. de Pago"
            Height          =   255
            Index           =   19
            Left            =   120
            TabIndex        =   53
            Top             =   560
            Width           =   1185
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Institución"
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   810
         End
      End
      Begin VB.Frame Fra_Pago 
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
         Height          =   1875
         Left            =   120
         TabIndex        =   46
         Top             =   4020
         Width           =   4335
         Begin VB.CommandButton Cmd_BenViaPago 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3960
            TabIndex        =   31
            ToolTipText     =   "Ver Historial de Vías de Pago"
            Top             =   1510
            Width           =   285
         End
         Begin VB.ComboBox Cmb_Suc 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   560
            Width           =   2955
         End
         Begin VB.TextBox Txt_NumCta 
            Height          =   285
            Left            =   960
            MaxLength       =   15
            TabIndex        =   30
            Top             =   1515
            Width           =   2940
         End
         Begin VB.ComboBox Cmb_TipCta 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   870
            Width           =   2940
         End
         Begin VB.ComboBox Cmb_ViaPago 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   240
            Width           =   2955
         End
         Begin VB.ComboBox Cmb_Banco 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   1185
            Width           =   2940
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   45
            Left            =   120
            TabIndex        =   111
            Top             =   0
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Sucursal"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   77
            Top             =   560
            Width           =   810
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "N°Cuenta"
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   50
            Top             =   1515
            Width           =   795
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Banco"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   49
            Top             =   1185
            Width           =   825
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Vía Pago"
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Width           =   810
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Tipo Cta."
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   47
            Top             =   870
            Width           =   825
         End
      End
      Begin VB.Frame Fra_Personales 
         Caption         =   "Antecedentes Personales"
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
         Height          =   2610
         Left            =   120
         TabIndex        =   43
         Top             =   1400
         Width           =   8895
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
            TabIndex        =   20
            ToolTipText     =   "Efectuar Busqueda de Dirección"
            Top             =   1680
            Width           =   300
         End
         Begin VB.TextBox Txt_NomSegBen 
            Height          =   285
            Left            =   5640
            MaxLength       =   25
            TabIndex        =   11
            Top             =   810
            Width           =   3015
         End
         Begin VB.CommandButton Cmd_BenDir 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   8520
            TabIndex        =   21
            ToolTipText     =   "Ver Historial de Direcciones"
            Top             =   1665
            Width           =   285
         End
         Begin VB.TextBox Txt_PatBen 
            Height          =   285
            Left            =   1200
            MaxLength       =   20
            TabIndex        =   12
            Top             =   1095
            Width           =   2895
         End
         Begin VB.TextBox Txt_MatBen 
            Height          =   285
            Left            =   5640
            MaxLength       =   20
            TabIndex        =   13
            Top             =   1095
            Width           =   3135
         End
         Begin VB.TextBox Txt_FonoBen 
            Height          =   285
            Left            =   1200
            MaxLength       =   15
            TabIndex        =   22
            Top             =   1980
            Width           =   2025
         End
         Begin VB.TextBox Txt_CorreoBen 
            Height          =   285
            Left            =   4215
            MaxLength       =   40
            TabIndex        =   23
            Top             =   1980
            Width           =   4560
         End
         Begin VB.TextBox Txt_NomBen 
            Height          =   285
            Left            =   1200
            MaxLength       =   25
            TabIndex        =   10
            Top             =   810
            Width           =   2895
         End
         Begin VB.TextBox Txt_DomBen 
            Height          =   285
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   16
            Top             =   1380
            Width           =   7575
         End
         Begin VB.Label Lbl_Distrito 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   6000
            TabIndex        =   19
            Top             =   1680
            Width           =   2175
         End
         Begin VB.Label Lbl_Provincia 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   3600
            TabIndex        =   18
            Top             =   1680
            Width           =   2295
         End
         Begin VB.Label Lbl_Departamento 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   17
            Top             =   1680
            Width           =   2295
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Segundo Nombre"
            Height          =   255
            Index           =   49
            Left            =   4320
            TabIndex        =   121
            Top             =   810
            Width           =   1455
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Ident."
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   120
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Lbl_TipoIdent 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   8
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label Lbl_MtoPQ 
            BackStyle       =   0  'Transparent
            Caption         =   "Mto. Pensión Quiebra"
            Height          =   255
            Left            =   5760
            TabIndex        =   115
            Top             =   2265
            Width           =   1635
         End
         Begin VB.Label Lbl_MtoPensionQui 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7440
            TabIndex        =   41
            Top             =   2265
            Width           =   1335
         End
         Begin VB.Label Lbl_FecFall 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   5640
            TabIndex        =   15
            Top             =   530
            Width           =   1335
         End
         Begin VB.Label Lbl_MtoPensionGar 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   4200
            TabIndex        =   25
            Top             =   2265
            Width           =   1335
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Fall."
            Height          =   255
            Index           =   48
            Left            =   4320
            TabIndex        =   114
            Top             =   530
            Width           =   795
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Mto. Pensión Gar."
            Height          =   255
            Index           =   47
            Left            =   2880
            TabIndex        =   113
            Top             =   2265
            Width           =   1275
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            Caption         =   "Antecedenes Personales"
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
            Index           =   46
            Left            =   120
            TabIndex        =   110
            Top             =   0
            Visible         =   0   'False
            Width           =   2235
         End
         Begin VB.Label Lbl_MtoPension 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   24
            Top             =   2265
            Width           =   1335
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mto. Pensión"
            Height          =   195
            Index           =   41
            Left            =   120
            TabIndex        =   108
            Top             =   2265
            Width           =   930
         End
         Begin VB.Label Lbl_SitInv 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   8280
            TabIndex        =   106
            Top             =   530
            Width           =   495
         End
         Begin VB.Label Lbl_FecNac 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   5640
            TabIndex        =   14
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sit. de Inv."
            Height          =   195
            Index           =   40
            Left            =   7320
            TabIndex        =   104
            Top             =   560
            Width           =   765
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Nac."
            Height          =   195
            Index           =   39
            Left            =   4320
            TabIndex        =   103
            Top             =   240
            Width           =   840
         End
         Begin VB.Label Lbl_NumIdent 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   9
            Top             =   525
            Width           =   2175
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono"
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   82
            Top             =   1980
            Width           =   795
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Email"
            Height          =   255
            Index           =   13
            Left            =   3720
            TabIndex        =   81
            Top             =   1980
            Width           =   495
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Ap. Materno"
            Height          =   255
            Index           =   8
            Left            =   4320
            TabIndex        =   80
            Top             =   1095
            Width           =   1095
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Ap. Paterno"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   79
            Top             =   1095
            Width           =   855
         End
         Begin VB.Label Lbl_NumOrd 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   8280
            TabIndex        =   76
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Orden"
            Height          =   255
            Index           =   5
            Left            =   7320
            TabIndex        =   75
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Primer Nombre"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   71
            Top             =   810
            Width           =   1095
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "N° Ident."
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   70
            Top             =   530
            Width           =   855
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Domicilio"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   45
            Top             =   1380
            Width           =   810
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Ubicación"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   44
            Top             =   1665
            Width           =   945
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Msf_Grilla 
         Height          =   930
         Left            =   120
         TabIndex        =   42
         Top             =   435
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   1640
         _Version        =   393216
         Rows            =   1
         Cols            =   6
         FixedCols       =   0
         BackColor       =   14745599
         FormatString    =   $"Frm_AntPensionado.frx":023C
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
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Tasa Cto. Reaseguro"
         Height          =   255
         Index           =   37
         Left            =   -74520
         TabIndex        =   124
         Top             =   4590
         Width           =   1815
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "CUSPP"
         Height          =   255
         Index           =   50
         Left            =   -74520
         TabIndex        =   123
         Top             =   4920
         Width           =   2535
      End
      Begin VB.Label Lbl_CUSPP 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   -71760
         TabIndex        =   122
         Top             =   4920
         Width           =   1695
      End
      Begin VB.Label Lbl_FecDevengue 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   -71760
         TabIndex        =   96
         Top             =   5535
         Width           =   1095
      End
      Begin VB.Label Lbl_FecEmision 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   -71760
         TabIndex        =   95
         Top             =   5220
         Width           =   1095
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Fecha de Devengue"
         Height          =   255
         Index           =   52
         Left            =   -74520
         TabIndex        =   119
         Top             =   5535
         Width           =   2535
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Fecha de Emisión"
         Height          =   255
         Index           =   20
         Left            =   -74520
         TabIndex        =   118
         Top             =   5220
         Width           =   2535
      End
      Begin VB.Label Lbl_Moneda 
         Caption         =   "(TM)"
         Enabled         =   0   'False
         Height          =   270
         Index           =   1
         Left            =   -70320
         TabIndex        =   117
         Top             =   3660
         Width           =   495
      End
      Begin VB.Label Lbl_Moneda 
         Caption         =   "(TM)"
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   -70320
         TabIndex        =   116
         Top             =   3360
         Width           =   495
      End
      Begin VB.Label Lbl_Mod 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   -71760
         TabIndex        =   101
         Top             =   2100
         Width           =   5295
      End
      Begin VB.Label Lbl_TipRta 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   -71760
         TabIndex        =   100
         Top             =   1485
         Width           =   5295
      End
      Begin VB.Label Lbl_Estado 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   -71760
         TabIndex        =   99
         Top             =   1170
         Width           =   5295
      End
      Begin VB.Label Lbl_TipPen 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   -71760
         TabIndex        =   98
         Top             =   870
         Width           =   5295
      End
      Begin VB.Label Lbl_Afp 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   -71760
         TabIndex        =   97
         Top             =   555
         Width           =   5295
      End
      Begin VB.Label Lbl_TasaPerGar 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   -67920
         TabIndex        =   94
         Top             =   4905
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Lbl_TasaRea 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   -71760
         TabIndex        =   93
         Top             =   4590
         Width           =   1095
      End
      Begin VB.Label Lbl_TasaVta 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   -71760
         TabIndex        =   92
         Top             =   4275
         Width           =   1095
      End
      Begin VB.Label Lbl_TasaCto 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   -71760
         TabIndex        =   91
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label Lbl_MtoPen 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   -71760
         TabIndex        =   90
         Top             =   3645
         Width           =   1455
      End
      Begin VB.Label Lbl_MtoPri 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   -71760
         TabIndex        =   89
         Top             =   3345
         Width           =   1455
      End
      Begin VB.Label Lbl_TerVig 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   -70080
         TabIndex        =   88
         Top             =   3030
         Width           =   1215
      End
      Begin VB.Label Lbl_IniVig 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   -71760
         TabIndex        =   87
         Top             =   3030
         Width           =   1215
      End
      Begin VB.Label Lbl_NumCar 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   -71760
         TabIndex        =   86
         Top             =   2715
         Width           =   735
      End
      Begin VB.Label Lbl_MesGar 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   -71760
         TabIndex        =   85
         Top             =   2415
         Width           =   1095
      End
      Begin VB.Label Lbl_MesDif 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   -71760
         TabIndex        =   84
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Lbl_NumEndoso 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   -71760
         TabIndex        =   83
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Tasa de Venta"
         Height          =   255
         Index           =   36
         Left            =   -74520
         TabIndex        =   69
         Top             =   4275
         Width           =   2295
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Tasa Cto. Equivalente"
         Height          =   255
         Index           =   35
         Left            =   -74520
         TabIndex        =   68
         Top             =   3960
         Width           =   1815
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
         Left            =   -70440
         TabIndex        =   67
         Top             =   3030
         Width           =   255
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Tipo de Renta"
         Height          =   255
         Index           =   26
         Left            =   -74520
         TabIndex        =   66
         Top             =   1485
         Width           =   1215
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Meses Garantizados"
         Height          =   255
         Index           =   29
         Left            =   -74520
         TabIndex        =   65
         Top             =   2415
         Width           =   1815
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Meses Diferidos"
         Height          =   255
         Index           =   27
         Left            =   -74520
         TabIndex        =   64
         Top             =   1785
         Width           =   1455
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Monto Pensión"
         Height          =   255
         Index           =   34
         Left            =   -74520
         TabIndex        =   63
         Top             =   3645
         Width           =   1215
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Monto Prima                           "
         Height          =   255
         Index           =   33
         Left            =   -74520
         TabIndex        =   62
         Top             =   3345
         Width           =   2775
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Periodo de Vigencia"
         Height          =   255
         Index           =   31
         Left            =   -74520
         TabIndex        =   61
         Top             =   3030
         Width           =   1695
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nº Beneficiario Póliza"
         Height          =   255
         Index           =   30
         Left            =   -74520
         TabIndex        =   60
         Top             =   2715
         Width           =   2055
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Estado"
         Height          =   255
         Index           =   25
         Left            =   -74520
         TabIndex        =   59
         Top             =   1170
         Width           =   1215
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Tipo de Pensión "
         Height          =   255
         Index           =   24
         Left            =   -74520
         TabIndex        =   58
         Top             =   870
         Width           =   1635
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Modalidad"
         Height          =   255
         Index           =   28
         Left            =   -74520
         TabIndex        =   57
         Top             =   2100
         Width           =   1260
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "AFP"
         Height          =   255
         Index           =   23
         Left            =   -74520
         TabIndex        =   56
         Top             =   555
         Width           =   975
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Número de Endoso"
         Height          =   255
         Index           =   22
         Left            =   -74520
         TabIndex        =   55
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Fra_Botones 
      Height          =   975
      Left            =   120
      TabIndex        =   102
      Top             =   7200
      Width           =   9165
      Begin VB.CommandButton Cmd_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   5400
         Picture         =   "Frm_AntPensionado.frx":02E5
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   200
         Width           =   730
      End
      Begin VB.CommandButton Cmd_Eliminar 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   360
         Picture         =   "Frm_AntPensionado.frx":08BF
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   "Eliminar Año"
         Top             =   240
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   3240
         Picture         =   "Frm_AntPensionado.frx":0C01
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   6480
         Picture         =   "Frm_AntPensionado.frx":12BB
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   4320
         Picture         =   "Frm_AntPensionado.frx":13B5
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Grabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   2160
         Picture         =   "Frm_AntPensionado.frx":1A6F
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   200
         Width           =   720
      End
      Begin Crystal.CrystalReport Rpt_General 
         Left            =   7320
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
End
Attribute VB_Name = "Frm_AntPensionado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------- INICIO VARIABLES -----------------------------
Dim vlNumEnd As Integer, vlcod As String, icod As String
Dim vlCodAFP As String, vlElemento As String, vlCodTp As String
Dim vlCodEst As String, vlCodRta As String, vlcodMod As String
Dim vlAnno As Integer, vlMes As Integer, vlDia As Integer
Dim vlRut As String, vlNombre As String, vlCodDir As Long
Dim vlSw As Boolean, vlI As Integer, vlCodViaPag As String
Dim vlCodSuc As String, vlCodTipcta As String, vlCodBco As String
Dim vlCodIns As String, vlCodModPago As String, vlNomSegBen As String
'CMV-20060804 I
Dim vlNumFun As String
'CMV-20060804 F
Dim vlCodSit As String, vlSwClic As Boolean

Dim vlArchivo As String
Dim vlStr As Variant
Dim i As Integer, vlLen As Integer, n As Integer
Dim vlNom As String
Dim iCodTabla As String, icanlin As Integer, inom As String

Public vgCodEstado As String
Dim vlFechaCrea As String, vlHoraCrea As String
Const clNumOrden = "1"
Const clCodSinDerPen = "10"
Dim vlSwNumCta As Boolean
Dim vlSwIS As Boolean

Dim vlSwMostrar As Boolean

'CMV-200650623 I
'Variables para Registros de Historia de Via de Pago, Direccion y
'Plan de Salud de Beneficiarios

Dim vlFecIniVig As String


'Direccion
Dim vlSwGrabarDir As Boolean

'Plan de Salud
Dim vlSwGrabarPlan As Boolean

'Via de Pago
Dim vlSwGrabarVia As Boolean

Dim vlRegHistoria As ADODB.Recordset

'CMV-200650623 F

'CMV 20060630 I
Const clGlsFrmDir As String = "FrmDir"
Const clGlsFrmVia As String = "FrmVia"
Const clGlsFrmSalud As String = "FrmSalud"
Const clTasaCtoRea0 As Integer = 0
Const clFechaTopeTer As String * 8 = "99991231"
Const clCodEstado6 As String * 1 = "6"

Dim vlNumPoliza As String
Dim vlNumOrden As Integer

'CMV-20061031 I
Dim vlUltimoPerPago As String
Dim vlPrcCastigoQui As Double
Dim vlTopeMaxQui As Double
Dim vlNumOrd As Integer

Dim vlCodTipoIden As String, vlNumIden As String

Const clCodTipReceptorR As String = "R"
Const clCodEstadoC As String * 1 = "C"
'CMV-200610331 F

'Constantes para la moneda
Const clMonedaMontoPrima As Integer = 0
Const clMonedaMontoPension As Integer = 1

Dim vlAfp As String
Dim vlMtoPensionActual As Double

'-------------------------------------------------------
'FUNCION QUE RECIBE LOS DATOS DEL FORMULARIO DE BUSQUEDA
'-------------------------------------------------------
'Function flRecibe(iPoliza, iRut, idig, iNumEnd)
'
'    Txt_PenPoliza = Trim(iPoliza)
'    Txt_PenRut = Trim(iRut)
'    Txt_PenDigito = Trim(idig)
'    Lbl_End = Trim(iNumEnd)
'    Call flCargarDatosBen(iRut, iPoliza, iNumEnd)
'    Call flCargaGrilla(iPoliza, iNumEnd)
'    Call flCargarDatosPol(iPoliza, iNumEnd)
'    SSTab1.Enabled = True
'    SSTab1.Tab = 0
'    Fra_Poliza.Enabled = False
'End Function

'-------------------------------------------------------
'FUNCION QUE RECIBE LOS DATOS DEL FORMULARIO DE BUSQUEDA
'-------------------------------------------------------
Function flRecibe(iPoliza, iCodTipoIden, iNumIden, iEndoso)
    Txt_PenPoliza = Trim(iPoliza)
    Lbl_End = Trim(iEndoso)
    Lbl_TipoIdent = Trim(iCodTipoIden)
    Lbl_NumIdent = Trim(iNumIden)
    Dim vlMontoPension As Double
    Call flCargarDatosPol(iPoliza, iEndoso, vlMontoPension)
    vlAfp = fgObtenerPolizaCod_AFP(CStr(iPoliza), CStr(iEndoso))
    Call flCargarDatosBen(iPoliza, iCodTipoIden, iNumIden, iEndoso, vlMontoPension)
    Call flCargaGrilla(iPoliza, iEndoso)
    SSTab1.Enabled = True
    SSTab1.Tab = 0
    Fra_Poliza.Enabled = False
End Function

'---------------------------------------
'DESHABILITA LOS CAMPOS DEL BENEFICIARIO
'---------------------------------------
Function flDesHab()
On Error GoTo Err_Habilita
    'Fra_Personales.Enabled = False
    Fra_Pago.Enabled = False
    Fra_Salud.Enabled = False
    
'    Txt_NomBen.Enabled = False
'    Txt_PatBen.Enabled = False
'    Txt_MatBen.Enabled = False
'    Txt_DomBen.Enabled = False
'    Cmb_Comuna.Enabled = False
'    Txt_FonoBen.Enabled = False
'    Txt_CorreoBen.Enabled = False
'    Cmb_ViaPago.Enabled = False
'    Cmb_Suc.Enabled = False
'    Cmb_TipCta.Enabled = False
'    Cmb_Banco.Enabled = False
'    Txt_NumCta.Enabled = False
'    Cmb_Inst.Enabled = False
'    Cmb_ModPago.Enabled = False
'    Txt_MtoPago.Enabled = False
    
Exit Function
Err_Habilita:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'Function flHabBen()
'    Fra_Personales.Enabled = True
''    Txt_NomBen.Enabled = True
''    Txt_PatBen.Enabled = True
''    Txt_MatBen.Enabled = True
''    Txt_DomBen.Enabled = True
''    Cmb_Comuna.Enabled = True
''    Txt_FonoBen.Enabled = True
''    Txt_CorreoBen.Enabled = True
'End Function
'------------------------------------
'HABILITA LOS CAMPOS DEL BENEFICIARIO
'------------------------------------
Function flHabilita()
On Error GoTo Err_DesHabilita
    Fra_Personales.Enabled = True
    Fra_Pago.Enabled = True
    Fra_Salud.Enabled = True
'    Txt_NomBen.Enabled = True
'    Txt_PatBen.Enabled = True
'    Txt_MatBen.Enabled = True
'    Txt_DomBen.Enabled = True
'    Cmb_Comuna.Enabled = True
'    Txt_FonoBen.Enabled = True
'    Txt_CorreoBen.Enabled = True
'    Cmb_ViaPago.Enabled = True
'    Cmb_Suc.Enabled = True
'    Cmb_TipCta.Enabled = True
'    Cmb_Banco.Enabled = True
'    Txt_NumCta.Enabled = True
'    Cmb_Inst.Enabled = True
'    Cmb_ModPago.Enabled = True
'    Txt_MtoPago.Enabled = True
Exit Function
Err_DesHabilita:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function
'----------------------------------------------------------
'DESHABILITA CAMPOS CUANDO EL BENEFICIARIO NO TIENE DERECHO
'----------------------------------------------------------
Function flDesHabSinDer()
    Fra_Pago.Enabled = False
    Fra_Salud.Enabled = False
'    Cmb_ViaPago.Enabled = False
'    Cmb_Banco.Enabled = False
'    Cmb_Suc.Enabled = False
'    Cmb_TipCta.Enabled = False
'    Txt_NumCta.Enabled = False
'    Cmb_Inst.Enabled = False
'    Cmb_ModPago.Enabled = False
'    Txt_MtoPago.Enabled = False
End Function

'VALIDA DATOS
Function flValida()
On Error GoTo Err_Valida
    
    flValida = False
    Txt_PenPoliza = Trim(UCase(Txt_PenPoliza))
    Txt_PenPoliza = Format(Txt_PenPoliza, "0000000000")
'    Txt_PenRut = Trim(Txt_PenRut)
'    Txt_PenDigito = Trim(UCase(Txt_PenDigito))
    
    If Trim(Txt_PenPoliza) = "" Then
        MsgBox "Debe Ingresar Número de Póliza", vbCritical, "Falta Información"
        Txt_PenPoliza.SetFocus
        flValida = True
        Exit Function
    End If
    
'    If Trim(Txt_PenRut) = "" Then
'        MsgBox "Debe ingresar el Rut del Pensionado.", vbCritical, "Falta Información"
'        Txt_PenRut.SetFocus
'        flValida = True
'        Exit Function
'    End If
'
'    If Trim(Txt_PenDigito) = "" Then
'        MsgBox "Debe ingresar el Dígito Verificador del Rut del Pensionado.", vbCritical, "Falta Información"
'        Txt_PenDigito.SetFocus
'        flValida = True
'        Exit Function
'    End If
    
Exit Function
Err_Valida:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'VALIDA DATOS DEL BENEFICIARIO
Function flValBen()
    flValBen = True
    If Txt_NomBen = "" Then
        MsgBox "El nombre del Beneficiario no puede estar en Blanco", vbCritical, "Falta Información"
        Txt_NomBen.SetFocus
        flValBen = False
        Exit Function
    End If
    If Txt_PatBen = "" Then
        MsgBox "El Apellido del Beneficiario no puede estar en Blanco", vbCritical, "Falta Información"
        Txt_PatBen.SetFocus
        flValBen = False
        Exit Function
    End If
    If Txt_MatBen = "" Then
''        MsgBox "El Apellido del Beneficiario no puede estar en Blanco", vbCritical, "Falta Información"
''        Txt_MatBen.SetFocus
''        flValBen = False
''        Exit Function
    End If
    If Txt_DomBen = "" Then
        MsgBox "El Domicilio del Beneficiario no puede estar en Blanco", vbCritical, "Falta Información"
        Txt_DomBen.SetFocus
        flValBen = False
        Exit Function
    End If
    If Txt_NumCta.Enabled = True And Txt_NumCta = "" Then
        MsgBox "Debe Ingresar el Número de Cuenta para este tipo de Vía de Pago", vbCritical, "Falta Información"
        Txt_NumCta.SetFocus
        flValBen = False
        Exit Function
    End If
    If Txt_MtoPago = "" Then
        MsgBox "El Monto de Pago no puede estar en Blanco", vbCritical, "Falta Información"
        Txt_MtoPago.SetFocus
        flValBen = False
        Exit Function
    End If
End Function

'FUNCIÓN QUE BUSCA LA GLOSA DEL ELEMENTO EN LA TABLA TABCOD
Function flBusEle(vlcod, icod)
On Error GoTo Err_BusDat
    vlElemento = ""
    flBusDat = False
    vgSql = ""
    vgSql = "select gls_elemento from MA_TPAR_TABCOD where "
    vgSql = vgSql & "cod_tabla= '" & vlcod & "' and "
    vgSql = vgSql & "cod_elemento= '" & icod & "'"
    Set vgRs4 = vgConexionBD.Execute(vgSql)
    If Not vgRs4.EOF Then
        vlElemento = vgRs4!GLS_ELEMENTO
        flBusEle = True
    End If
    vgRs4.Close
Exit Function
Err_BusDat:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function
'FUNCION QUE POSISIONA EL INDICE EN EL REGISTRO QUE
'CORRESPONDE
Function flBusPos(icod, iElemento, iCombo As ComboBox)
On Error GoTo Err_Gls

    If iCombo.Text <> "" Then
        For vlI = 0 To iCombo.ListCount - 1
        vlSwIS = True
            iCombo.ListIndex = vlI
            If iCombo.Text = icod + " - " + iElemento Then
                Exit For
            End If
        Next vlI
        vlSwIS = False 'hqr 04/04/2005
        'vlSw = False
    End If
Exit Function
Err_Gls:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'Permite Limpiar la Grilla de Beneficiarios
Function flLmpGrilla()

    Msf_Grilla.Clear
    Msf_Grilla.Cols = 6
    Msf_Grilla.Rows = 1
    
    Msf_Grilla.Row = 0
        
    Msf_Grilla.Col = 0
    Msf_Grilla.Text = "NºOrden"
    Msf_Grilla.ColWidth(0) = 800
    Msf_Grilla.ColAlignment(0) = 3
    
    Msf_Grilla.Col = 1
    Msf_Grilla.Text = "Cód.Parentesco"
    Msf_Grilla.ColWidth(1) = 1000
    
    Msf_Grilla.Col = 2
    Msf_Grilla.Text = "Tipo Ident."
    Msf_Grilla.ColWidth(2) = 800
    
    Msf_Grilla.Col = 3
    Msf_Grilla.Text = "NºIdent."
    Msf_Grilla.ColWidth(3) = 400
    
    Msf_Grilla.Col = 4
    Msf_Grilla.Text = "Nombre"
    Msf_Grilla.ColWidth(4) = 1500
    
    Msf_Grilla.Col = 5
    Msf_Grilla.Text = "Estado"
    Msf_Grilla.ColWidth(4) = 680
    
End Function

'FUNCION QUE CARGA LA GRILLA CON LOS DATOS DEL BENEFICIARIO
Function flCargaGrilla(iPoliza, iEndoso)
On Error GoTo Err_Cargar
    Msf_Grilla.Clear
    Msf_Grilla.Cols = 6
    Msf_Grilla.Rows = 1
    
    Msf_Grilla.Row = 0
        
    Msf_Grilla.Col = 0
    Msf_Grilla.Text = "NºOrden"
    Msf_Grilla.ColWidth(0) = 800
    Msf_Grilla.ColAlignment(0) = 3
    
    Msf_Grilla.Col = 1
    Msf_Grilla.Text = "Cód.Parentesco"
    Msf_Grilla.ColWidth(1) = 1200
    
    Msf_Grilla.Col = 2
    Msf_Grilla.Text = "Tipo Ident."
    Msf_Grilla.ColWidth(2) = 1500
    
    Msf_Grilla.Col = 3
    Msf_Grilla.Text = "NºIdent"
    Msf_Grilla.ColWidth(3) = 1000
    
    Msf_Grilla.Col = 4
    Msf_Grilla.Text = "Nombre"
    Msf_Grilla.ColWidth(4) = 4500
    
    Msf_Grilla.Col = 5
    Msf_Grilla.Text = "Estado"
    Msf_Grilla.ColWidth(5) = 600
   
    vgSql = "select "
    vgSql = vgSql & "num_orden,cod_par,cod_tipoidenben,num_idenben,gls_nomben, "
    vgSql = vgSql & "gls_nomsegben,gls_patben,gls_matben,cod_estado "
    vgSql = vgSql & "from PP_TMAE_BEN B, PP_TMAE_POLIZA P where "
    vgSql = vgSql & "B.num_poliza=P.num_poliza "
    vgSql = vgSql & "and B.num_endoso=P.num_endoso "
    vgSql = vgSql & "and B.num_poliza = '" & Trim(Txt_PenPoliza) & "' and "
    vgSql = vgSql & "B.num_endoso = '" & Trim(Lbl_End) & "' "
    Set vgRs2 = vgConexionBD.Execute(vgSql)
    Do While Not vgRs2.EOF
    If Not IsNull(vgRs2!Gls_NomSegBen) Then
        vlNomSegBen = (vgRs2!Gls_NomSegBen)
    Else
        vlNomSegBen = ""
    End If
    vlTipoIdent = (vgRs2!Cod_TipoIdenBen & " - " & fgBuscarNombreTipoIden(vgRs2!Cod_TipoIdenBen))
    Msf_Grilla.AddItem vgRs2!Num_Orden & vbTab & _
                           vgRs2!Cod_Par & vbTab & _
                           " " & vlTipoIdent & vbTab & _
                           vgRs2!Num_IdenBen & vbTab & _
                           ((vgRs2!Gls_NomBen) + " " + (vlNomSegBen) + " " + (vgRs2!Gls_PatBen) + " " + (vgRs2!Gls_MatBen)) & vbTab & _
                           vgRs2!Cod_Estado
                            'I---- ABV 21/08/2004 ---
                            '(vgRs2!Cod_Derpen)
                            'F---- ABV 21/08/2004 ---
        vgRs2.MoveNext
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
Function flCargarDatosBen(iPoliza, iCodTipoIden, iNumIden, iEndoso, iMontoPension)
On Error GoTo Err_CargarDatos
    flCargarDatosBen = True
      
    'busca el beneficiario seleccionado en la bd
    If vlSwClic = False Then
        vgSql = "select a.*, TO_CHAR(SYSDATE, 'yyyymmdd') as Fecha_Actual from PP_TMAE_BEN a where "
        vgSql = vgSql & "num_poliza = '" & Trim(iPoliza) & "' AND "
        vgSql = vgSql & "cod_tipoidenben = " & Trim(iCodTipoIden) & " AND "
        vgSql = vgSql & "num_idenben = '" & Trim(iNumIden) & "' "
        vgSql = vgSql & "order by num_endoso desc "
    Else
        vgSql = "select a.*, TO_CHAR(SYSDATE, 'yyyymmdd') as Fecha_Actual from PP_TMAE_BEN a where "
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
        Lbl_TipoIdent = (vgRs2!Cod_TipoIdenBen & " - " & fgBuscarNombreTipoIden(vgRs2!Cod_TipoIdenBen))
        Lbl_NumIdent = vgRs2!Num_IdenBen
        Txt_NomBen = vgRs2!Gls_NomBen
        
        If Not IsNull(vgRs2!Gls_NomSegBen) Then
            Txt_NomSegBen = (vgRs2!Gls_NomSegBen)
        Else
            Txt_NomSegBen = ""
            vlNomSegBen = ""
        End If
        
        Txt_PatBen = vgRs2!Gls_PatBen
        If Not IsNull(vgRs2!Gls_MatBen) Then
            Txt_MatBen = vgRs2!Gls_MatBen
        Else
            Txt_MatBen = ""
        End If
        
        vlAnno = Mid(vgRs2!Fec_NacBen, 1, 4)
        vlMes = Mid(vgRs2!Fec_NacBen, 5, 2)
        vlDia = Mid(vgRs2!Fec_NacBen, 7, 2)
        Lbl_FecNac = DateSerial(vlAnno, vlMes, vlDia)
        If Not IsNull(vgRs2!Fec_FallBen) Then
            Lbl_FecFall = DateSerial(Mid((vgRs2!Fec_FallBen), 1, 4), Mid((vgRs2!Fec_FallBen), 5, 2), Mid((vgRs2!Fec_FallBen), 7, 2))
        Else
            Lbl_FecFall = ""
        End If
        Lbl_NumOrd = Trim(vgRs2!Num_Orden)
        'CMV-20061102 I
        vlNumOrd = Lbl_NumOrd
        'CMV-20061102 F
        Lbl_SitInv = Trim(vgRs2!Cod_SitInv)
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
        'CMV-20060804 I
        'CMV-20060804 F
        
        'hqr 20/10/2007 Se despliega pension actualizada
        'Lbl_MtoPension = Format(vgRs2!Mto_Pension, "#,#0.00")
        'Lbl_MtoPensionGar = Format(vgRs2!Mto_PensionGar, "#,#0.00")
        
'        If vgRs2!Mto_Pension > 0 Then
'            Lbl_MtoPension = Format(iMontoPension * (vgRs2!Prc_Pension / 100), "#,#0.00")
'        Else
'            Lbl_MtoPension = Format(vgRs2!Mto_Pension, "#,#0.00")
'        End If
'
'        If vgRs2!Mto_PensionGar > 0 Then
'            Lbl_MtoPensionGar = Format(iMontoPension * (vgRs2!Prc_PensionGar / 100), "#,#0.00")
'        Else
'            Lbl_MtoPensionGar = Format(vgRs2!Mto_PensionGar, "#,#0.00")
'        End If

        If vgRs2!Cod_DerPen <> clCodSinDerPen Then
            Lbl_MtoPension = Format(iMontoPension * (vgRs2!Prc_Pension / 100), "#,#0.00")
            If Not IsNull(vgRs2!Fec_TerPagoPenGar) Then
                If vgRs2!Fec_TerPagoPenGar >= vgRs2!Fecha_Actual Then
                    Lbl_MtoPensionGar = Format(iMontoPension * (vgRs2!Prc_PensionGar / 100), "#,#0.00")
                Else
                    Lbl_MtoPensionGar = Format(0, "#,#0.00")
                End If
            Else
                Lbl_MtoPensionGar = Format(0, "#,#0.00")
            End If
        Else
            Lbl_MtoPension = Format(vgRs2!Mto_Pension, "#,#0.00")
            Lbl_MtoPensionGar = Format(vgRs2!Mto_PensionGar, "#,#0.00")
        End If
        'fin hqr 20/10/2007
        
        'CMV-20061031 I
        'Mostrar Monto de Pensión en Quiebra
        vlUltimoPerPago = flUltimoPeriodoCerrado(iPoliza)
        If fgObtieneParametrosQuiebra(vlUltimoPerPago, vlPrcCastigoQui, vlTopeMaxQui) Then
            Lbl_MtoPQ.Visible = True
            Lbl_MtoPensionQui.Visible = True
            Lbl_MtoPensionQui = flBuscarPensionQuiebra(vlUltimoPerPago, iPoliza, vlNumOrd, clCodTipReceptorR)
            Lbl_MtoPensionQui = Format(Lbl_MtoPensionQui, "#,#0.00")
        End If
        'CMV-20061031 F
        
        'se llena el combo comuna, los campos provincia y region
'        If Cmb_Comuna.Text <> "" Then
'            For vlI = 0 To Cmb_Comuna.ListCount - 1
'                If Cmb_Comuna.ItemData(vlI) = vlCodDir Then
'                    Cmb_Comuna.ListIndex = vlI
'                    Exit For
'                End If
'            Next vlI
'        End If
  
        'se carga el combo de via de pago
        Call flBusEle(vgCodTabla_ViaPago, vlCodViaPag)
        vlSwMostrar = False
        Call flBusPos(vlCodViaPag, vlElemento, Cmb_ViaPago)
        vlSwMostrar = True

        'se carga la sucursal
        Call flBusSuc(vlCodSuc)
        Call flBusPos(vlCodSuc, vlElemento, Cmb_Suc)

        'se carga el combo de tipo de cuenta
        Call flBusEle(vgCodTabla_TipCta, vlCodTipcta)
        Call flBusPos(vlCodTipcta, vlElemento, Cmb_TipCta)

        'se carga el combo banco
        Call flBusEle(vgCodTabla_Bco, vlCodBco)
        Call flBusPos(vlCodBco, vlElemento, Cmb_Banco)
    
        'se carga el combo de institucion de salud
        Call flBusEle(vgCodTabla_InsSal, vlCodIns)
        Call flBusPos(vlCodIns, vlElemento, Cmb_Inst)
        
        'se carga el combo de modalidad de pago
        Call flBusEle(vgCodTabla_ModPago, vlCodModPago)
        Call flBusPos(vlCodModPago, vlElemento, Cmb_ModPago)
        'Call flBusEle(vgCodTabla_ModPago, vlCodModPago2)
        'Call flBusPos(vlCodModPago2, vlElemento, Cmb_ModPago2)

        flCargarDatosBen = False
        Call Cmb_ViaPago_Click
        'Call Cmb_Inst_Click 'HQR 05/04/2005
        
        If vgRs2!Cod_EstPension = "10" Then
            Call flDesHabSinDer
        Else
            Fra_Pago.Enabled = True
            Fra_Salud.Enabled = True
        End If
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

Function flCargarDatosPol(iPoliza, iEndoso, oMontoPension) As Boolean
On Error GoTo Err_CDP
        flCargarDatosPol = False
        vgSql = ""
        vgSql = "select * from PP_TMAE_POLIZA where "
        vgSql = vgSql & "num_poliza = '" & Trim(iPoliza) & "' and "
        vgSql = vgSql & "num_endoso = '" & Trim(iEndoso) & "'"
        Set vgRs = vgConexionBD.Execute(vgSql)
        If Not vgRs.EOF Then
            Lbl_NumEndoso = vgRs!num_endoso
            
            '-------- busca la glosa de la afp segun el codigo
            vlCodAFP = vgRs!cod_afp
            If flBusEle(vgCodTabla_AFP, vlCodAFP) = True Then
                Lbl_Afp = (vlCodAFP) + " - " + (vlElemento)
            Else
                Lbl_Afp = ""
            End If
            '-------- busca la glosa del tipo de pension segun el codigo
            vlCodTp = vgRs!Cod_TipPension
            If flBusEle(vgCodTabla_TipPen, vlCodTp) = True Then
                Lbl_TipPen = (vlCodTp) + " - " + (vlElemento)
            Else
                Lbl_TipPen = ""
            End If
            
            '-------- busca la glosa del estado segun el codigo
             vlCodEst = vgRs!Cod_Estado
            vgCodEstado = vgRs!Cod_Estado
            If flBusEle(vgCodTabla_TipVigPol, vlCodEst) = True Then
                Lbl_Estado = (vlCodEst) + " - " + (vlElemento)
            Else
                Lbl_Estado = ""
            End If

            '-------- busca la glosa del tipo de renta segun el codigo
            vlCodRta = vgRs!Cod_TipRen
            If flBusEle(vgCodTabla_TipRen, vlCodRta) = True Then
                Lbl_TipRta = (vlCodRta) + " - " + (vlElemento)
            Else
                Lbl_TipRta = ""
            End If
            
            '-------- busca la glosa de la modalidad segun el codigo
            vlcodMod = vgRs!Cod_Modalidad
            If flBusEle(vgCodTabla_AltPen, vlcodMod) = True Then
                Lbl_Mod = (vlcodMod) + " - " + (vlElemento)
            Else
                Lbl_Mod = ""
            End If
            
            Lbl_MesDif = vgRs!Num_MesDif
            Lbl_MesGar = vgRs!Num_MesGar
            Lbl_NumCar = vgRs!Num_Cargas
            
            'se cambia el formato de la fecha de inicio de vigencia a dd/mm/aaaa
            vlAnno = Mid(vgRs!Fec_Vigencia, 1, 4)
            vlMes = Mid(vgRs!Fec_Vigencia, 5, 2)
            vlDia = Mid(vgRs!Fec_Vigencia, 7, 2)
            Lbl_IniVig = DateSerial(vlAnno, vlMes, vlDia)
            
            'se cambia el formato de la fecha de termino de vigencia a dd/mm/aaaa
            vlAnno = Mid(vgRs!Fec_TerVigencia, 1, 4)
            vlMes = Mid(vgRs!Fec_TerVigencia, 5, 2)
            vlDia = Mid(vgRs!Fec_TerVigencia, 7, 2)
            Lbl_TerVig = DateSerial(vlAnno, vlMes, vlDia)
            
            'Obtiene los scomps
            Lbl_Moneda(clMonedaMontoPrima) = fgObtenerCodMonedaScomp(egTablaMoneda(), vgNumeroTotalTablasMoneda, cgCodTipMonedaUF)
            Lbl_Moneda(clMonedaMontoPension) = fgObtenerCodMonedaScomp(egTablaMoneda(), vgNumeroTotalTablasMoneda, vgRs!Cod_Moneda)
       
            Lbl_MtoPri = Format(vgRs!Mto_Prima, "#,#0.00")
            
            Lbl_MtoPen = Format(vgRs!Mto_Pension, "#,#0.00")
            oMontoPension = vgRs!Mto_Pension
            'hqr 20/10/2007 Obtiene Monto de la Pensión Actualizada
            'Obtiene Pension Actualizada
            vlSql = "SELECT mto_pension FROM pp_tmae_pensionact a"
            vlSql = vlSql & " WHERE a.num_poliza = '" & Trim(iPoliza) & "'"
            vlSql = vlSql & " AND a.num_endoso = " & Trim(iEndoso)
            vlSql = vlSql & " AND a.fec_desde = "
                vlSql = vlSql & " (SELECT max(fec_desde) FROM pp_tmae_pensionact b"
                vlSql = vlSql & " WHERE b.num_poliza = a.num_poliza"
                vlSql = vlSql & " AND b.num_endoso = a.num_endoso"
                vlSql = vlSql & " AND b.fec_desde <= TO_CHAR(SYSDATE,'yyyymmdd'))"
            Set vlTB2 = vgConexionBD.Execute(vlSql)
            If Not vlTB2.EOF Then
                If Not IsNull(vlTB2!Mto_Pension) Then
                    Lbl_MtoPen = Format(vlTB2!Mto_Pension, "#,#0.00")
                    oMontoPension = vlTB2!Mto_Pension
                End If
            End If
            
            Lbl_TasaCto = Format(vgRs!Prc_TasaCe, "#,#0.00")
            Lbl_TasaVta = Format(vgRs!Prc_TasaVta, "#,#0.00")
            Lbl_TasaRea = Format(vgRs!prc_tasactorea, "#,#0.00")
            Lbl_CUSPP = (vgRs!Cod_Cuspp)
            Lbl_TasaPerGar = Format(vgRs!Prc_TasaIntPerGar, "#,#0.00")
            
            'se cambia el formato de la fecha de vigencia a dd/mm/aaaa
            vlAnno = Mid(vgRs!Fec_Emision, 1, 4)
            vlMes = Mid(vgRs!Fec_Emision, 5, 2)
            vlDia = Mid(vgRs!Fec_Emision, 7, 2)
            Lbl_FecEmision = DateSerial(vlAnno, vlMes, vlDia)
            
            'se cambia el formato de la fecha de devengue a dd/mm/aaaa
            vlAnno = Mid(vgRs!fec_dev, 1, 4)
            vlMes = Mid(vgRs!fec_dev, 5, 2)
            vlDia = Mid(vgRs!fec_dev, 7, 2)
            Lbl_FecDevengue = DateSerial(vlAnno, vlMes, vlDia)
'
'            Fra_Poliza.Enabled = False
'            Txt_PenPoliza.Enabled = True
'            Cmb_PenNumIdent.Enabled = True
'            Txt_PenNumIdent.Enabled = True
'            Lbl_End.Enabled = True
            Cmd_Buscar.Enabled = False
            Cmd_BuscarPol.Enabled = False
'
'            If vgCodEstado = "9" Then
'                Call flDesHab
'            Else
'                Fra_Personales.Enabled = True
''            End If
'            Fra_Personales.Enabled = True

        Else
            MsgBox "La Póliza Ingresada no contiene Información", vbCritical, "Operación Cancelada"
            Cmd_Cancelar_Click
            Exit Function
        End If
        
        SSTab1.Tab = 0
        Fra_Poliza.Enabled = False
        flCargarDatosPol = True
Exit Function
Err_CDP:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'BUSCA LA GLOSA DEL CODIGO DE LA SUCURSAL
Function flBusSuc(icodsuc)
    vgSql = ""
    vgSql = "select * from MA_TPAR_SUCURSAL where "
    vgSql = vgSql & "cod_sucursal = '" & icodsuc & "'"
    Set vgRs3 = vgConexionBD.Execute(vgSql)
    If Not vgRs3.EOF Then
        vlElemento = vgRs3!gls_sucursal
        Exit Function
    End If
End Function

Function flImprimir(vlNumPol, vlNumEnd)
On Error GoTo Err_Reporte

    vlArchivo = strRpt & "PP_Rpt_AntPensionado.rpt"
    If Not fgExiste(vlArchivo) Then
        MsgBox "Archivo de Reporte de Antecedentes de Pensionado no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
        Screen.MousePointer = 0
        Exit Function
    End If
    
    vgQuery = "{PP_TMAE_POLIZA.NUM_POLIZA} = '" & Trim(vlNumPol) & "' and "
    vgQuery = vgQuery & "{PP_TMAE_POLIZA.NUM_ENDOSO} = " & Trim(vlNumEnd) & ""
    
    Rpt_General.Reset
    Rpt_General.ReportFileName = vlArchivo
    Rpt_General.Connect = vgRutaDataBase
    Rpt_General.SelectionFormula = ""
    Rpt_General.SelectionFormula = vgQuery
    Rpt_General.Formulas(0) = ""
    Rpt_General.Formulas(1) = ""
    Rpt_General.Formulas(2) = ""
    Rpt_General.Formulas(3) = ""
    Rpt_General.Formulas(4) = ""
    Rpt_General.Formulas(5) = ""
    Rpt_General.Formulas(6) = ""
    Rpt_General.Formulas(7) = ""
    Rpt_General.Formulas(8) = ""
    Rpt_General.Formulas(9) = ""
    Rpt_General.Formulas(10) = ""
    Rpt_General.Formulas(11) = ""
    Rpt_General.Formulas(12) = ""
    Rpt_General.Formulas(13) = ""
    Rpt_General.Formulas(14) = ""
    Rpt_General.Formulas(15) = ""
    Rpt_General.Formulas(16) = ""
    Rpt_General.Formulas(17) = ""
    
    Rpt_General.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
    Rpt_General.Formulas(1) = "NombreSistema = '" & vgNombreSistema & "'"
    Rpt_General.Formulas(2) = "NombreSubSistema = '" & vgNombreSubSistema & "'"
    Rpt_General.Formulas(3) = "MtoPension = " & Replace(Format(Lbl_MtoPen, "#0.00"), ",", ".")
    
    i = 4
    vlLen = 0
        
    'Busca los códigos de parentesco y su glosa
    Call flCarCodRep(vgCodTabla_Par, 1, "Parentesco", "cod_par")
    'Busca los códigos de la situación de invalidez
    Call flCarCodRep(vgCodTabla_SitInv, 1, "SitInv", "cod_sitinv")
    'Busca los códigos de la via de pago
    Call flCarCodRep(vgCodTabla_ViaPago, 1, "ViaPago", "cod_viapago")
    'Busca los codigos del tipo de cuenta
    Call flCarCodRep(vgCodTabla_TipCta, 1, "TipCta", "cod_tipcuenta")
    'Busca los codigos del Banco
    Call flCarCodRep(vgCodTabla_Bco, 1, "Banco", "cod_banco")
    'Busca los codigos de la inst. de salud
    Call flCarCodRep(vgCodTabla_InsSal, 1, "InsSal", "cod_inssalud")
    'Busca los codigos de la modalidad de pago
    Call flCarCodRep(vgCodTabla_ModPago, 1, "ModPago", "cod_modsalud")
'    'Busca los codigos de la modalidad de pago 2
'    Call flCarCodRep(vgCodTabla_ModPago, 1, "ModPago2", "cod_modsalud2")
    
    'Busca los codigos de la sucursal
    vgSql = ""
    vgSql = "select * from MA_TPAR_SUCURSAL where "
    vgSql = vgSql & "cod_sucursal in (select distinct cod_sucursal "
    vgSql = vgSql & "from PP_TMAE_BEN where num_poliza= '" & Trim(Txt_PenPoliza) & "' and "
    vgSql = vgSql & "num_endoso = '" & Trim(Lbl_End) & "')"
    vgSql = vgSql & " order by cod_sucursal asc"
    Set vgRs = vgConexionBD.Execute(vgSql)
    n = 1
    For n = 1 To 2
        vlLen = 0
        vlStr = ""
        Do While Not vgRs.EOF And vlLen <= 200
                vlStr = vlStr + (vgRs!Cod_Sucursal) + " - " + (vgRs!gls_sucursal) + " / "
                vlLen = Len(vlStr)
                vgRs.MoveNext
        Loop
        If vlLen <> 0 Then
            vlStr = Mid(vlStr, 1, vlLen - 3)
            vlNom = ("Sucursal" + Str(n))
            Rpt_General.Formulas(i) = vlNom & "= '" & Trim(vlStr) & "'"
            i = i + 1
        End If
    Next n

    
    
    Rpt_General.Destination = crptToWindow
    Rpt_General.WindowState = crptMaximized
    Rpt_General.WindowTitle = "Informe de Antecedentes del Pensionado"
    Rpt_General.Action = 1
    
Exit Function
Err_Reporte:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flCarCodRep(iCodTabla, icanlin, inom, icodele)
On Error GoTo Err_Codigos
    
    Sql = ""
    Sql = Sql & "select cod_elemento, gls_elemento from MA_TPAR_TABCOD "
    Sql = Sql & "where cod_tabla= '" & iCodTabla & "'and "
    Sql = Sql & "cod_elemento in "
    Sql = Sql & "(select distinct " & icodele & " from PP_TMAE_BEN WHERE "
    Sql = Sql & "num_poliza = '" & Trim(Txt_PenPoliza) & "' and "
    Sql = Sql & "num_endoso = '" & Trim(Lbl_End) & "')"
    Sql = Sql & " order by cod_elemento asc"
    Set vgRs = vgConexionBD.Execute(Sql)
    For n = 1 To icanlin
        vlLen = 0
        vlStr = ""
        Do While Not vgRs.EOF And vlLen <= 200
            vlStr = vlStr + (vgRs!cod_elemento) + " - " + (vgRs!GLS_ELEMENTO) + " / "
                vlLen = Len(vlStr)
                vgRs.MoveNext
        Loop
        If vlLen <> 0 Then
            vlStr = Mid(vlStr, 1, vlLen - 3)
            vlNom = (inom + Str(n))
            Rpt_General.Formulas(i) = vlNom & "= '" & Trim(vlStr) & "'"
            i = i + 1
        End If
    Next n
Exit Function
Err_Codigos:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Private Sub Cmb_Banco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txt_NumCta.SetFocus
    End If
End Sub

'Private Sub Cmb_Comuna_Click()
'    If Cmb_Comuna.Text <> "" Then
'        vlI = Cmb_Comuna.ListIndex
'        vlCodDir = Cmb_Comuna.ItemData(vlI)
'        fgBuscarNombreProvinciaRegion (vlCodDir)
'        Lbl_Provincia = vgNombreProvincia
'        Lbl_Region = vgNombreRegion
'    End If
'End Sub

Private Sub Cmb_Comuna_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txt_FonoBen.SetFocus
    End If
End Sub

Private Sub Cmb_Inst_Click()
    vlCodIns = Trim(Mid(Cmb_Inst.Text, 1, (InStr(1, Cmb_Inst.Text, "-") - 1)))
    If vlCodIns = "00" Then
        If vlSwIS = False Then 'hqr 05/04/2005 Se agrega, para que no cambie el monto
            Txt_MtoPago = "0"
            'Txt_MtoPago2 = "0"
        End If
        Txt_MtoPago = Format(Txt_MtoPago, "#,#0.000")
        Txt_MtoPago.Enabled = False
'        Txt_MtoPago2 = Format(Txt_MtoPago2, "#,#0.000")
'        Txt_MtoPago2.Enabled = False
    Else
        If vlSwIS = False Then 'hqr 05/04/2005 Se agrega, para que no cambie el monto
            Txt_MtoPago.Text = ""
            'Txt_MtoPago2.Text = ""
        End If
        Txt_MtoPago.Enabled = True
        'Txt_MtoPago2.Enabled = True
    End If
End Sub

Private Sub Cmb_Inst_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cmb_ModPago.SetFocus
    End If
End Sub

Private Sub Cmb_ModPago_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txt_MtoPago.SetFocus
    End If
End Sub

Private Sub Cmb_ModPago2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txt_MtoPago2.SetFocus
    End If
End Sub

Private Sub Cmb_Suc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If (Cmb_TipCta.Enabled = True) Then
            Cmb_TipCta.SetFocus
        Else
            Cmb_Inst.SetFocus
        End If
    End If
End Sub

Private Sub Cmb_TipCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cmb_Banco.SetFocus
    End If
End Sub

Private Sub Cmb_ViaPago_Click()

    If vlSwMostrar = True Then
        Exit Sub
    End If

    vlCodViaPag = Trim(Mid(Cmb_ViaPago.Text, 1, (InStr(1, Cmb_ViaPago.Text, "-") - 1)))
    If vlCodViaPag = "01" Or vlCodViaPag = "04" Then 'caja
        If vlSw = False Then
            If (vlCodViaPag = "04") Then
                vgTipoSucursal = cgTipoSucursalAfp
            Else
                vgTipoSucursal = cgTipoSucursalSuc
            End If
            fgComboSucursal Cmb_Suc, vgTipoSucursal
        
            Cmb_TipCta.ListIndex = 0
            Cmb_Banco.ListIndex = 0
            If (vlCodViaPag = "04") Then
                vgPalabra = fgObtenerCodigo_TextoCompuesto(vlAfp)
                Call fgBuscarPosicionCodigoCombo(vgPalabra, Cmb_Suc)
            End If
            Cmb_TipCta.Enabled = False
            Cmb_Banco.Enabled = False
            If vlSwNumCta = False Then
                Txt_NumCta = ""
            End If
            Txt_NumCta.Enabled = False
            Cmb_Suc.Enabled = True
        End If
    Else
        If (vlCodViaPag = "00" Or vlCodViaPag = "05") And vlSw = False Then 'sin información
            vgTipoSucursal = cgTipoSucursalSuc
            fgComboSucursal Cmb_Suc, vgTipoSucursal
            
            Cmb_TipCta.ListIndex = 0
            Cmb_Banco.ListIndex = 0
            Cmb_Suc.ListIndex = 0
            Cmb_TipCta.Enabled = False
            Cmb_Banco.Enabled = False
            Cmb_Suc.Enabled = False
            If vlSwNumCta = False Then
                Txt_NumCta = ""
            End If
            Txt_NumCta.Enabled = False
        Else
            If vlSw = False Then
                
                vgTipoSucursal = cgTipoSucursalSuc
                fgComboSucursal Cmb_Suc, vgTipoSucursal
                
                If vlCodViaPag = "02" Or vlCodViaPag = "03" Then
                    Cmb_Suc.ListIndex = 0
                    Cmb_Suc.Enabled = False
                    Cmb_TipCta.Enabled = True
                    Cmb_Banco.Enabled = True
                    Txt_NumCta.Enabled = True
                Else
                    Cmb_TipCta.ListIndex = 0
                    Cmb_Banco.ListIndex = 0
                    Cmb_Suc.ListIndex = 0
                    Cmb_TipCta.Enabled = True
                    Cmb_Banco.Enabled = True
                    Cmb_Suc.Enabled = True
                    Txt_NumCta = ""
                    Txt_NumCta.Enabled = True
                End If
            End If
        End If
    End If
        
End Sub

Private Sub Cmb_ViaPago_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If (Cmb_Suc.Enabled = True) Then
            Cmb_Suc.SetFocus
        Else
            Cmb_Inst.SetFocus
        End If
    End If
End Sub

Private Sub Cmd_BenDir_Click()
    vgNumPol = Txt_PenPoliza
    vgNumOrden = CInt(Lbl_NumOrd)
    vgRutBen = Lbl_TipoIdent
    vgDgvBen = Lbl_NumIdent
    vgNomBen = Trim(Txt_NomBen) & " " & Trim(Txt_NomSegBen) & " " & Trim(Txt_PatBen) & " " & Trim(Txt_MatBen)
    vgGlsTipoForm = clGlsFrmDir
    Screen.MousePointer = vbHourglass
    Frm_ConsultaHistBen.Show
    Frm_ConsultaHistBen.Caption = "Consulta Datos Históricos - Dirección"
    Screen.MousePointer = vbDefault
End Sub

Private Sub Cmd_BenSalud_Click()
    vgNumPol = Txt_PenPoliza
    vgNumOrden = CInt(Lbl_NumOrd)
    vgRutBen = Lbl_TipoIdent
    vgDgvBen = Lbl_NumIdent
    vgNomBen = Trim(Txt_NomBen) & " " & Trim(Txt_NomSegBen) & " " & Trim(Txt_PatBen) & " " & Trim(Txt_MatBen)
    vgGlsTipoForm = clGlsFrmSalud
    Screen.MousePointer = vbHourglass
    Frm_ConsultaHistBen.Caption = "Consulta Datos Históricos - Plan de Salud"
    Frm_ConsultaHistBen.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub Cmd_BenViaPago_Click()
    vgNumPol = Txt_PenPoliza
    vgNumOrden = CInt(Lbl_NumOrd)
    vgRutBen = Lbl_TipoIdent
    vgDgvBen = Lbl_NumIdent
    vgNomBen = Trim(Txt_NomBen) & " " & Trim(Txt_NomSegBen) & " " & Trim(Txt_PatBen) & " " & Trim(Txt_MatBen)
    vgGlsTipoForm = clGlsFrmVia
    Screen.MousePointer = vbHourglass
    Frm_ConsultaHistBen.Caption = "Consulta Datos Históricos - Forma de Pago"
    Frm_ConsultaHistBen.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub Cmd_Buscar_Click()
On Error GoTo Err_Buscar

    Frm_Busqueda.flInicio ("Frm_AntPensionado")
    
Exit Sub
Err_Buscar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_BuscarDir_Click()
On Error GoTo Err_Buscar

    Frm_BusDireccion.flInicio ("Frm_AntPensionado")
    
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
    Txt_FonoBen.SetFocus

End Function

Private Sub Cmd_BuscarPol_Click()
On Error GoTo Err_Buscar
    
    Screen.MousePointer = 11
    
    vlSwMostrar = True
   
   If Txt_PenPoliza = "" Then
           
'        If ((Trim(Txt_PenRut.Text)) = "") Or (Txt_PenDigito.Text = "") Or _
'          (Not ValiRut(Txt_PenRut.Text, Txt_PenDigito.Text)) Then
           MsgBox "Debe Ingresar el Número de Póliza del Pensionado.", vbCritical, "Error de Datos"
           Txt_PenPoliza.SetFocus
           Screen.MousePointer = 0
           Exit Sub
        Else
           'Txt_PenRut = Format(Txt_PenRut, "##,###,##0")
           Txt_PenPoliza = Trim(UCase(Txt_PenPoliza))
           Txt_PenPoliza = Format(Txt_PenPoliza, "0000000000")
           'Txt_PenDigito = UCase(Trim(Txt_PenDigito))
           'Txt_PenDigito.SetFocus
           'vlRutAux = Format(Txt_PenRut, "#0")
    End If
    
'    vgPalabra = ""
'    'If (Txt_PenPoliza.Text <> "") And (Txt_PenRut.Text <> "") Then
'    If (Txt_PenPoliza.Text <> "") And (Cmb_PenNumIdent.Text <> "") Then
'        'vlRutAux = Format(Txt_PenRut, "#0")
'        vgPalabra = "num_poliza = '" & Txt_PenPoliza.Text & "' AND "
'        'vgPalabra = vgPalabra & "rut_ben = " & vlRutAux & " "
'    Else
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
'        Else
'            If Txt_PenRut.Text <> "" Then
'               vgPalabra = "rut_ben = " & vlRutAux & " "
'            End If
        'End If
    End If
    vgSql = ""
    vgSql = "SELECT num_endoso,num_orden,gls_nomben,gls_nomsegben,gls_patben,gls_matben, "
    vgSql = vgSql & "cod_estpension,cod_tipoidenben,num_idenben,num_poliza "
    vgSql = vgSql & "FROM PP_TMAE_BEN WHERE "
    vgSql = vgSql & vgPalabra
    vgSql = vgSql & " ORDER BY num_endoso DESC,num_orden ASC "
    Set vgRs2 = vgConexionBD.Execute(vgSql)
    
    If Not vgRs2.EOF Then
    
        vlNumOrden = (vgRs2!Num_Orden)
                
        vlAfp = fgObtenerPolizaCod_AFP(vgRs2!num_poliza, CStr(vgRs2!num_endoso))
        
        If Trim(vgRs2!Cod_EstPension) = Trim(clCodSinDerPen) Then
        MsgBox " El Beneficiario Seleccionado No Tiene Derecho a Pensión " & Chr(13) & _
        "          Sólo podrá Consultar los Datos del Registro", vbInformation, "Información"
            vlDerecho = False
            Lbl_End = vgRs2!num_endoso
            
            If Not IsNull(vgRs2!Gls_NomSegBen) Then
                Txt_NomSegBen = Trim(vgRs2!Gls_NomSegBen)
                Lbl_PenNombre = Trim(vgRs2!Gls_NomBen) + " " + (vgRs2!Gls_NomSegBen) + " " + Trim(vgRs2!Gls_PatBen) + " " + IIf(IsNull(vgRs2!Gls_MatBen), "", Trim(vgRs2!Gls_MatBen))
            Else
                Txt_NomSegBen = ""
                vlNomSegBen = ""
                Lbl_PenNombre = Trim(vgRs2!Gls_NomBen) + " " + (vlNomSegBen) + " " + Trim(vgRs2!Gls_PatBen) + " " + IIf(IsNull(vgRs2!Gls_MatBen), "", Trim(vgRs2!Gls_MatBen))
            End If
            
            'Txt_PenRut = Format(vgRs2!Rut_Ben, "##,###,##0")
            Txt_PenNumIdent = UCase(vgRs2!Num_IdenBen)
            Txt_PenPoliza = vgRs2!num_poliza
        Else
            vlDerecho = True
            Lbl_End = vgRs2!num_endoso
            
            If Not IsNull(vgRs2!Gls_NomSegBen) Then
                Txt_NomSegBen = Trim(vgRs2!Gls_NomSegBen)
                Lbl_PenNombre = Trim(vgRs2!Gls_NomBen) + " " + (vgRs2!Gls_NomSegBen) + " " + Trim(vgRs2!Gls_PatBen) + " " + Trim(vgRs2!Gls_MatBen)
            Else
                Txt_NomSegBen = ""
                vlNomSegBen = ""
                Lbl_PenNombre = Trim(vgRs2!Gls_NomBen) + " " + (vlNomSegBen) + " " + Trim(vgRs2!Gls_PatBen) + " " + IIf(IsNull(vgRs2!Gls_MatBen), "", Trim(vgRs2!Gls_MatBen))
            End If
           
            'Txt_PenRut = Format(vgRs2!Rut_Ben, "##,###,##0")
            Txt_PenNumIdent = UCase(vgRs2!Num_IdenBen)
            Txt_PenPoliza = vgRs2!num_poliza
        End If
    Else
        MsgBox " La Póliza consultada no existe en la Base de Datos ", vbInformation, "Información"
        Screen.MousePointer = 0
        Exit Sub
    End If
 
'    If flValida = False Then
'        If ValiRut(Trim(Txt_PenRut), (Trim(Txt_PenDigito))) = False Then
'            MsgBox "El Dígito Verificador no es Válido para el Rut Ingresado", vbCritical, "Operación Cancelada"
'            Screen.MousePointer = 0
'            Txt_PenDigito.SetFocus
'            Exit Sub
'        End If
'        'se busca al beneficiario y se cargan los text del beneficiario
        Dim vlMontoPension As Double
        If flCargarDatosPol(Trim(Txt_PenPoliza), Trim(Lbl_End), vlMontoPension) Then
            If flCargarDatosBen(Trim(Txt_PenPoliza), vgRs2!Cod_TipoIdenBen, vgRs2!Num_IdenBen, Trim(Lbl_End), vlMontoPension) = False Then
                'se carga la grilla
                Call flCargaGrilla(Trim(Txt_PenPoliza), Trim(Lbl_End))
                'se cargan los datos de la póliza
                'Call flCargarDatosPol(Trim(Txt_PenPoliza), Trim(Lbl_End))
                SSTab1.Enabled = True
                SSTab1.Tab = 0
                Fra_Poliza.Enabled = False
                
                vlSwMostrar = False
            End If
        End If
    Screen.MousePointer = 0

Exit Sub
Err_Buscar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Cancelar_Click()
On Error GoTo Err_Canc
    
    Screen.MousePointer = 11
    Txt_PenPoliza = ""
    Cmb_PenNumIdent.ListIndex = 0
    Txt_PenNumIdent = ""
    Lbl_End = ""
    Lbl_PenNombre = ""
    Lbl_NumEndoso = ""
    Lbl_Afp = ""
    Lbl_TipPen = ""
    Lbl_Estado = ""
    Lbl_TipRta = ""
    Lbl_MesDif = ""
    Lbl_Mod = ""
    Lbl_MesGar = ""
    Lbl_NumCar = ""
    Lbl_IniVig = ""
    Lbl_TerVig = ""
    Lbl_MtoPri = ""
    Lbl_MtoPen = ""
    Lbl_TasaCto = ""
    Lbl_TasaVta = ""
    Lbl_TasaRea = ""
    Lbl_CUSPP = ""
    Lbl_FecEmision = ""
    Lbl_FecDevengue = ""
    
    Lbl_Moneda(clMonedaMontoPrima) = "S/."
    Lbl_Moneda(clMonedaMontoPension) = ""
    
    'I---- ABV 23/08/2004 ---
    Call flLmpGrilla
    Call flDesHab
    'Msf_Grilla.Clear
    'Call Cmd_Limpiar_Click
    'F---- ABV 23/08/2004 ---
    
    Fra_Poliza.Enabled = True
    'Txt_PenPoliza.Enabled = True
    Txt_PenPoliza.SetFocus
    Cmd_Buscar.Enabled = True
    Cmd_BuscarPol.Enabled = True
    SSTab1.Tab = 0
    SSTab1.Enabled = False
    'Txt_PenRut.Enabled = True
    'Txt_PenDigito.Enabled = True

    Screen.MousePointer = 0
    
Exit Sub
Err_Canc:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Eliminar_Click()
On Error GoTo Err_Elimina
    If Lbl_End = "" Then
        MsgBox "Debe seleccionar el Registro que Desea Eliminar", vbCritical, "Proceso de Eliminación Cancelado"
        Exit Sub
    End If

    Screen.MousePointer = 11

    vgSql = ""
    vgSql = "select * from PP_TMAE_BEN where "
    vgSql = vgSql & "num_poliza= '" & Txt_PenPoliza & "' and "
    vgSql = vgSql & "rut_ben= '" & Str(Lbl_RutBen) & "' and "
    vgSql = vgSql & "num_orden= '" & Lbl_NumOrd & "'"
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
        vgRes = MsgBox(" ¿ Está seguro que desea Eliminar los Datos ?", 4 + 32 + 256, "Proceso de Eliminación")
        If vgRes <> 6 Then
            Cmd_Salir.SetFocus
            Screen.MousePointer = 0
            Exit Sub
        End If
        vgSql = "delete from PP_TMAE_BEN where "
        vgSql = vgSql & "num_poliza= '" & Txt_PenPoliza & "' and "
        vgSql = vgSql & "num_orden= '" & Lbl_NumOrd & "' and "
        vgSql = vgSql & "rut_ben= '" & Lbl_RutBen & "'"
        vgConexionBD.Execute vgSql
        
        MsgBox "Los Datos han sido Eliminados Satisfactoriamente", vbCritical, "Proceso de eliminación"
        Call Cmd_Limpiar_Click
        Call flCargaGrilla(Txt_PenPoliza, Lbl_End)
    Else
        MsgBox "El Registro que Intenta Eliminar no se encuentra en la Base de Datos", vbCritical, "Proceso de Eliminación Cancelado"
    End If
    Screen.MousePointer = 0
    

Exit Sub
Err_Elimina:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub cmd_grabar_Click()
On Error GoTo Err_Graba
iFecha = Date

If Fra_Personales.Enabled = False Then
    MsgBox "Debe escoger una Póliza", vbCritical, "Datos Incompletos"
    Exit Sub
End If
    
If fgValidaVigenciaPoliza(Trim(Txt_PenPoliza), Trim(iFecha)) = False Then
    MsgBox " La Póliza Consultada no se encuentra vigente " & Chr(13) & _
            "  No puede ingresar ni Modificar Información ", vbExclamation, "Operación Cancelada"
    Screen.MousePointer = 0
    Exit Sub
End If
'        If fgValidaPagoPension(Trim(Txt_FecInicio), Trim(iFecha), clNumOrden) = False Then
'            MsgBox " Ya se ha realizado el proceso de Cálculo de Pensión para ésta fecha ", vbExclamation, "Operación Cancelada"
'            Screen.MousePointer = 0
'            Exit Sub
'        End If
'End If

If flValida = False Then
    
    'validacion via de pago
    vlCodViaPag = Trim(Mid(Cmb_ViaPago.Text, 1, (InStr(1, Cmb_ViaPago.Text, "-") - 1)))
    vlCodSuc = Trim(Mid(Cmb_Suc.Text, 1, (InStr(1, Cmb_Suc.Text, "-") - 1)))
    vlCodTipcta = Trim(Mid(Cmb_TipCta.Text, 1, (InStr(1, Cmb_TipCta.Text, "-") - 1)))
    vlCodBco = Trim(Mid(Cmb_Banco.Text, 1, (InStr(1, Cmb_Banco.Text, "-") - 1)))
        
    If vlCodViaPag = "01" Then
        If vlCodSuc = "0000" Then
            MsgBox "Debe seleccionar la Sucursal de la Vía de Pago", vbInformation, "Falta Información"
            Exit Sub
        End If
    Else
        If vlCodViaPag = "02" Or vlCodViaPag = "03" Then
            If vlCodTipcta = "00" Then
                MsgBox "Debe seleccionar el tipo de Cuenta", vbInformation, "Falta Información"
                Exit Sub
            End If
            If vlCodBco = "00" Then
                MsgBox "Debe seleccionar el Banco", vbInformation, "Falta Información"
                Exit Sub
            End If
            If Txt_NumCta = "" Then
                MsgBox "Debe ingresar el número de cuenta", vbInformation, "Falta Información"
                Exit Sub
            End If
        End If
    End If
    
    'validacion plan de salud
    vlCodIns = Trim(Mid(Cmb_Inst.Text, 1, (InStr(1, Cmb_Inst.Text, "-") - 1)))
    If vlCodIns <> "00" Then
        If Txt_MtoPago.Text = "0.000" Or Txt_MtoPago.Text = "" Then
            MsgBox "Debe ingresar el Monto de Pago", vbInformation, "Falta Información"
            Txt_MtoPago.SetFocus
            Exit Sub
        End If
    Else
        Txt_MtoPago.Text = Format("0", "#0.000")
    End If

    Screen.MousePointer = 11
    vgSql = "select * from PP_TMAE_BEN where "
    vgSql = vgSql & "num_poliza= '" & Txt_PenPoliza & "' and "
    vgSql = vgSql & "num_endoso= '" & Lbl_End & "' and "
    vgSql = vgSql & "num_orden= '" & Lbl_NumOrd & "'"
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
        If flValBen = True Then
        vgRes = MsgBox(" ¿ Está seguro que desea Modificar los Datos ?", 4 + 32 + 256, "Proceso de Actualización")
        If vgRes <> 6 Then
            Cmd_Salir.SetFocus
            Screen.MousePointer = 0
            Exit Sub
        End If
        vlCodViaPag = Trim(Mid(Cmb_ViaPago.Text, 1, (InStr(1, Cmb_ViaPago.Text, "-") - 1)))
        vlCodSuc = Trim(Mid(Cmb_Suc.Text, 1, (InStr(1, Cmb_Suc.Text, "-") - 1)))
        vlCodTipcta = Trim(Mid(Cmb_TipCta.Text, 1, (InStr(1, Cmb_TipCta.Text, "-") - 1)))
        vlCodBco = Trim(Mid(Cmb_Banco.Text, 1, (InStr(1, Cmb_Banco.Text, "-") - 1)))
        vlCodIns = Trim(Mid(Cmb_Inst.Text, 1, (InStr(1, Cmb_Inst.Text, "-") - 1)))
        vlCodModPago = Trim(Mid(Cmb_ModPago.Text, 1, (InStr(1, Cmb_ModPago.Text, "-") - 1)))
        vgSql = ""
        vgSql = "update PP_TMAE_BEN set "
        vgSql = vgSql & "gls_nomben = '" & Trim(Txt_NomBen) & "', "
        vgSql = vgSql & "gls_nomsegben='" & Trim(Txt_NomSegBen) & "', "
        vgSql = vgSql & "gls_patben = '" & Trim(Txt_PatBen) & "', "
        vgSql = vgSql & "gls_matben = '" & Trim(Txt_MatBen) & "', "
        vgSql = vgSql & "gls_dirben = '" & Trim(Txt_DomBen) & "', "
        vgSql = vgSql & "cod_direccion = '" & vlCodDir & "', "
        If Txt_FonoBen <> "" Then
            vgSql = vgSql & "gls_fonoben = '" & Trim(Txt_FonoBen) & "', "
        Else
            vgSql = vgSql & "gls_fonoben = NULL, "
        End If
        If Txt_CorreoBen <> "" Then
            vgSql = vgSql & "gls_correoben = '" & Trim(Txt_CorreoBen) & "', "
        Else
            vgSql = vgSql & "gls_correoben = NULL, "
        End If
        vgSql = vgSql & "cod_viapago= '" & Trim(vlCodViaPag) & "', "
        vgSql = vgSql & "cod_sucursal = '" & vlCodSuc & "', "
        vgSql = vgSql & "cod_tipcuenta = '" & vlCodTipcta & "', "
        vgSql = vgSql & "cod_banco = '" & vlCodBco & "', "
        If Txt_NumCta.Enabled = True Then
            vgSql = vgSql & "num_cuenta = '" & Trim(Txt_NumCta) & "', "
        Else
            vgSql = vgSql & "num_cuenta = null, "
        End If
        vgSql = vgSql & "cod_inssalud = '" & vlCodIns & "', "
        vgSql = vgSql & "cod_modsalud = '" & vlCodModPago & "', "
        vgSql = vgSql & "mto_plansalud = " & Str(Txt_MtoPago) & ", "
        vgSql = vgSql & "cod_usuariomodi= '" & vgUsuario & "', "
        vgSql = vgSql & "fec_modi= '" & Format(Date, "yyyymmdd") & "', "
        vgSql = vgSql & "hor_modi= '" & Format(Time, "hhmmss") & "' "
        vgSql = vgSql & "where "
        vgSql = vgSql & "num_poliza= '" & Trim(Txt_PenPoliza) & "' and "
        'CMV-20060623 I
        vgSql = vgSql & "num_endoso= '" & Lbl_End & "' and "
        'CMV-20060623 F
        vgSql = vgSql & "num_orden = " & Lbl_NumOrd & " "
        vgConexionBD.Execute (vgSql)
        
        MsgBox "Los Datos han sido actualizados Satisfactoriamente", vbInformation, "Operación de Actualización"
        'Call Cmd_Limpiar_Click
        Call flCargaGrilla(Txt_PenPoliza, Lbl_End)
        End If
        
        vlFecIniVig = ""
        vlFecIniVig = fgBuscaFecServ
        vlFecIniVig = Format(vlFecIniVig, "yyyymmdd")
        
        vlNumPoliza = Trim(Txt_PenPoliza)
        vlNumOrden = Lbl_End
        
        Call flGuardarHistViaPago
        Call flGuardarHistPlanSalud
        Call flGuardarHistDireccion
        
    Else
        MsgBox "El Registro que intenta actualizar no se encuentra en la Base de Datos", vbCritical, "Operación Cancelada"
    
    End If
End If
Screen.MousePointer = 0
    
    
Exit Sub
Err_Graba:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Imprimir_Click()
On Error GoTo Err_Imp

    If Trim(Txt_PenPoliza = "") Or (Lbl_End = "") Then
        MsgBox "Debe escoger una Póliza para Imprimir", vbCritical, "Datos Incompletos"
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    vgSql = ""
    vgSql = "select num_poliza, num_endoso from PP_TMAE_BEN where "
    vgSql = vgSql & "num_poliza = '" & Trim(Txt_PenPoliza) & "' and "
    vgSql = vgSql & "num_endoso = '" & Trim(Lbl_End) & "'"
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
        Call flImprimir(Trim(Txt_PenPoliza), Trim(Lbl_End))
    Else
        MsgBox "El Registro No se encuentra en la Base de Datos", vbCritical, "Proceso de Impresión Cancelada"
        Screen.MousePointer = 0
        Exit Sub
    End If
    Screen.MousePointer = 0

Exit Sub
Err_Imp:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Limpiar_Click()
On Error GoTo Err_Limp

    Screen.MousePointer = 11
    If Txt_PenPoliza <> "" And Lbl_End <> "" Then
        Call flCargaGrilla(Txt_PenPoliza, Lbl_End)
    End If
    Cmb_ViaPago.ListIndex = 0
    Cmb_Suc.ListIndex = 0
    Cmb_TipCta.ListIndex = 0
    Cmb_Banco.ListIndex = 0
    Cmb_Inst.ListIndex = 0
    Cmb_ModPago.ListIndex = 0
    Lbl_TipoIdent = ""
    Lbl_NumIdent = ""
    Lbl_NumOrd = ""
    Lbl_FecNac = ""
    Lbl_FecFall = ""
    Lbl_SitInv = ""
    Txt_NomSegBen = ""
    Txt_NomBen = ""
    Txt_PatBen = ""
    Txt_MatBen = ""
    Txt_DomBen = ""
    Lbl_Departamento = ""
    Lbl_Provincia = ""
    Lbl_Distrito = ""
    Txt_FonoBen = ""
    Txt_CorreoBen = ""
    Lbl_MtoPension = ""
    Lbl_MtoPensionGar = ""
    Lbl_MtoPensionQui = ""
    Lbl_Moneda(clMonedaMontoPrima) = "S/."
    Lbl_Moneda(clMonedaMontoPension) = ""

    'CMV-20050420 I
    Txt_NumCta = ""
    Txt_MtoPago = ""

    Call flDesHab
    Fra_Personales.Enabled = True
    Fra_Pago.Enabled = True
    Fra_Salud.Enabled = True
    'CMV-20050420 F
    Screen.MousePointer = 0
    
Exit Sub
Err_Limp:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
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

    Frm_AntPensionado.Top = 0
    Frm_AntPensionado.Left = 0
    
    'CMV-20061102 I
    Lbl_MtoPQ.Visible = False
    Lbl_MtoPensionQui.Visible = False
    'CMV-20061102 F
    
    Lbl_Moneda(clMonedaMontoPrima) = ""
    Lbl_Moneda(clMonedaMontoPension) = ""
    
    SSTab1.Tab = 0
    vlSw = True
    Call flDesHab
    'Call fgComboComuna(Cmb_Comuna)
    'vlSwMostrar = False: vlSw = False
    Call fgComboGeneral(vgCodTabla_ViaPago, Cmb_ViaPago)
    'vlSwMostrar = True: vlSw = True
    Call fgComboSucursal(Cmb_Suc, "S")
    Call fgComboGeneral(vgCodTabla_TipCta, Cmb_TipCta)
    Call fgComboGeneral(vgCodTabla_Bco, Cmb_Banco)
    Call fgComboGeneral(vgCodTabla_InsSal, Cmb_Inst)
    Call fgComboGeneral(vgCodTabla_ModPago, Cmb_ModPago)
    'Call fgComboGeneral(vgCodTabla_ModPago, Cmb_ModPago2)
    fgComboTipoIdentificacion Cmb_PenNumIdent
'    Cmb_ModPago2.ListIndex = fgBuscarPosicionCodigoCombo(cgCodTipMonedaUF, Cmb_ModPago2)
    'Call flPosicionaModPago2UF
    vlSw = False
    SSTab1.Enabled = False
    
    Call fgCargarTablaMoneda(vgCodTabla_TipMon, egTablaMoneda(), vgNumeroTotalTablasMoneda)

Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Msf_Grilla_DblClick()
On Error GoTo Err_Grilla

    vlSwMostrar = True
    
    If Msf_Grilla.Rows = 1 Then
        Exit Sub
    End If
    
    vlPos = Msf_Grilla.RowSel
    Msf_Grilla.Row = vlPos
    Msf_Grilla.Col = 0
    If (Msf_Grilla.Text = "") And vlPos = 0 Then
        Exit Sub
    End If
        
    Screen.MousePointer = 11
    
    'CARGA DATOS DE LA GRILLA A LOS TEXT
    Msf_Grilla.Col = 2
    'Lbl_RutBen = Trim(Mid(Msf_Grilla.Text, 1, (InStr(1, Msf_Grilla.Text, "-") - 1)))
    vlCodTipoIden = fgObtenerCodigo_TextoCompuesto(Msf_Grilla.Text)
    If Msf_Grilla.Text <> "" Then
        Lbl_TipoIdent = Msf_Grilla.Text
    End If
    
    Msf_Grilla.Col = 3
    vlNumIden = Msf_Grilla.Text
    
    vlSwClic = True
    Call flCargarDatosBen(Trim(Txt_PenPoliza), vlCodTipoIden, vlNumIden, (Lbl_End), Lbl_MtoPen)
    vlSwClic = False
    If vgCodEstado <> 9 Then 'Póliza No Vigente
        Fra_Personales.Enabled = True
        Txt_NomBen.SetFocus
    Else
        Call flDesHab
    End If
    'If vgRs2!Cod_EstPension = "10" Then
    If vgRs2!Cod_DerPen = "10" Then 'HQR 03/05/2005
        'Fra_Personales.Enabled = False
        Fra_Pago.Enabled = False
        Fra_Salud.Enabled = False
    Else
        Fra_Personales.Enabled = True
        Fra_Pago.Enabled = True
        Fra_Salud.Enabled = True
    End If
    
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

Private Sub Txt_CorreoBen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If (Fra_Pago.Enabled = True) Then
            Cmb_ViaPago.SetFocus
        Else
            Cmd_Grabar.SetFocus
        End If
    End If
End Sub

Private Sub Txt_DomBen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cmd_BuscarDir.SetFocus
    End If
End Sub

Private Sub Txt_DomBen_LostFocus()
    Txt_DomBen = UCase(Txt_DomBen)
End Sub

Private Sub Txt_FonoBen_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        Txt_CorreoBen.SetFocus
    End If
End Sub

Private Sub Txt_MatBen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txt_DomBen.SetFocus
    End If
End Sub

Private Sub Txt_MatBen_LostFocus()
    Txt_MatBen = UCase(Txt_MatBen)
End Sub

Private Sub Txt_MtoPago_Change()
    If Not IsNumeric(Txt_MtoPago) Then
        Txt_MtoPago = ""
    End If
End Sub

Private Sub Txt_MtoPago_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cmd_Grabar.SetFocus
    End If
End Sub

Private Sub Txt_MtoPago_LostFocus()
    If Txt_MtoPago <> "" Then
        Txt_MtoPago = Format(Txt_MtoPago, "#,#0.000")
    End If
End Sub

Private Sub Txt_MtoPago2_Change()
    If Not IsNumeric(Txt_MtoPago2) Then
        Txt_MtoPago2 = ""
    End If
End Sub

Private Sub Txt_MtoPago2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cmd_Grabar.SetFocus
    End If
End Sub

Private Sub Txt_MtoPago2_LostFocus()
    If Txt_MtoPago2 <> "" Then
        Txt_MtoPago2 = Format(Txt_MtoPago2, "#,#0.000")
    End If
End Sub

Private Sub Txt_NomBen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txt_NomSegBen.SetFocus
    End If
End Sub

Private Sub Txt_NomBen_LostFocus()
    Txt_NomBen = Trim(UCase(Txt_NomBen))
End Sub

Private Sub Txt_NomSegBen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txt_PatBen.SetFocus
    End If
End Sub

Private Sub Txt_NomSegBen_LostFocus()
    Txt_NomSegBen = Trim(UCase(Txt_NomSegBen))
End Sub

Private Sub Txt_NumCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cmb_Inst.SetFocus
    End If
End Sub

Private Sub Txt_NumCta_LostFocus()
    Txt_NumCta = Trim(UCase(Txt_NumCta))
End Sub

Private Sub Txt_NumFun_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If (Trim(Txt_NumFun) <> "") Then
            Txt_NumFun = Trim(UCase(Txt_NumFun))
            Txt_NumFun = Format(Txt_NumFun, "0000000000")
            Cmd_BenSalud.SetFocus
        Else
            Cmd_BenSalud.SetFocus
        End If
    End If
End Sub

Private Sub Txt_NumFun_LostFocus()
    Txt_NumFun = Trim(UCase(Txt_NumFun))
    Txt_NumFun = Format(Txt_NumFun, "0000000000")
End Sub

Private Sub Txt_PatBen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txt_MatBen.SetFocus
    End If
End Sub

Private Sub Txt_PatBen_LostFocus()
    Txt_PatBen = Trim(UCase(Txt_PatBen))
End Sub

Private Sub Txt_PenDigito_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If (Txt_PenDigito <> "") Then
            Txt_PenDigito = Trim(UCase(Txt_PenDigito))
            If (Txt_PenRut <> "") And (Txt_PenDigito <> "") Then
                If ValiRut(Trim(Txt_PenRut), (Trim(Txt_PenDigito))) = False Then
                    MsgBox "El Dígito Verificador no es Válido para el Rut Ingresado", vbCritical, "Operación Cancelada"
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            End If
            
        End If
        Cmd_BuscarPol.SetFocus
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

'Private Sub Txt_PenDigito_LostFocus()
'    Txt_PenDigito = Trim(UCase(Txt_PenDigito))
'End Sub

Private Sub Txt_PenPoliza_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If (Trim(Txt_PenPoliza) <> "") Then
            Txt_PenPoliza = Trim(UCase(Txt_PenPoliza))
            Txt_PenPoliza = Format(Txt_PenPoliza, "0000000000")
'        Else
'            MsgBox "Debe Ingresar el Número de Póliza del Pensionado.", vbCritical, "Error de Datos"
'            Cmd_BuscarPol.SetFocus
        End If
        Cmb_PenNumIdent.SetFocus
    End If
End Sub

Private Sub Txt_PenPoliza_LostFocus()
    Txt_PenPoliza = Trim(UCase(Txt_PenPoliza))
    Txt_PenPoliza = Format(Txt_PenPoliza, "0000000000")
End Sub

Private Sub Cmb_PenNumIdent_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Txt_PenNumIdent.SetFocus
End If
End Sub

'Private Sub Txt_PenRut_Change()
'    If Not IsNumeric(Txt_PenRut) Then
'        Txt_PenRut = ""
'    End If
'End Sub

'Private Sub Txt_PenRut_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If (Txt_PenRut <> "") Then
'            Txt_PenRut = Format(Txt_PenRut, "#,#0")
'            If (Txt_PenRut <> "") And (Txt_PenDigito <> "") Then
'                If ValiRut(Trim(Txt_PenRut), (Trim(Txt_PenDigito))) = False Then
'                    MsgBox "El Dígito Verificador no es Válido para el Rut Ingresado", vbCritical, "Operación Cancelada"
'                    Screen.MousePointer = 0
'                    'Txt_PenDigito.SetFocus
'                    Exit Sub
'                End If
'            End If
'
'        End If
'        Txt_PenDigito.SetFocus
'    End If
'End Sub

'Private Sub Txt_PenRut_LostFocus()
'    If Txt_PenRut <> "" Then
'        Txt_PenRut = Format(Txt_PenRut, "#,#0")
'        If (Txt_PenRut <> "") And (Txt_PenDigito <> "") Then
'            If ValiRut(Trim(Txt_PenRut), (Trim(Txt_PenDigito))) = False Then
'                MsgBox "El Dígito Verificador no es Válido para el Rut Ingresado", vbCritical, "Operación Cancelada"
'                Screen.MousePointer = 0
'                Exit Sub
'            End If
'        End If
'    End If
'End Sub

Function flGuardarHistViaPago()
On Error GoTo Err_flGuardarHistViaPago

    vlSwGrabarVia = False

    vgSql = ""
    vgSql = "SELECT v.cod_viapago,v.cod_banco,v.cod_tipcuenta, "
    vgSql = vgSql & "v.num_cuenta,v.cod_sucursal "
    vgSql = vgSql & "FROM pp_this_benviapago v WHERE "
    vgSql = vgSql & "num_poliza = '" & vlNumPoliza & "' AND "
    vgSql = vgSql & "num_orden = " & vlNumOrden & " AND "
    vgSql = vgSql & "v.fec_inivig = "
    vgSql = vgSql & "(SELECT MAX(v.fec_inivig) "
    vgSql = vgSql & "FROM pp_this_benviapago v WHERE "
    vgSql = vgSql & "num_poliza = '" & vlNumPoliza & "' AND "
    vgSql = vgSql & "num_orden = " & vlNumOrden & ") "
    Set vlRegHistoria = vgConexionBD.Execute(vgSql)
    If Not vlRegHistoria.EOF Then
        If (vlRegHistoria!Cod_ViaPago) <> Trim(vlCodViaPag) Then
            vlSwGrabarVia = True
        End If
        If (vlRegHistoria!Cod_Banco) <> Trim(vlCodBanco) Then
            vlSwGrabarVia = True
        End If
        If (vlRegHistoria!Cod_TipCuenta) <> Trim(vlCodTipcta) Then
            vlSwGrabarVia = True
        End If
        If (vlRegHistoria!Num_Cuenta) <> Trim(Txt_NumCta) Then
            vlSwGrabarVia = True
        End If
        If (vlRegHistoria!Cod_Sucursal) <> Trim(vlCodSuc) Then
            vlSwGrabarVia = True
        End If
    End If
    
    If vlSwGrabarVia = True Then
    
        vgSql = ""
        vgSql = "SELECT v.num_poliza FROM pp_this_benviapago v "
        vgSql = vgSql & "WHERE "
        vgSql = vgSql & "v.num_poliza = '" & Trim(Txt_PenPoliza) & "' AND "
        vgSql = vgSql & "v.num_orden = " & Lbl_NumOrd & " AND "
        vgSql = vgSql & "v.fec_inivig = '" & Trim(vlFecIniVig) & "' "
        Set vgRegistro = vgConexionBD.Execute(vgSql)
        If Not vgRegistro.EOF Then
            'Eliminar el registro ya que corresponde al mismo dia
            vgSql = ""
            vgSql = "DELETE pp_this_benviapago "
            vgSql = vgSql & "WHERE "
            vgSql = vgSql & "num_poliza = '" & Trim(Txt_PenPoliza) & "' AND "
            vgSql = vgSql & "num_orden = " & Lbl_NumOrd & " AND "
            vgSql = vgSql & "fec_inivig = '" & Trim(vlFecIniVig) & "' "
            vgConexionBD.Execute (vgSql)
        End If
        vgRegistro.Close
    
        vgSql = ""
        vgSql = "INSERT INTO pp_this_benviapago "
        vgSql = vgSql & "(num_poliza,num_endoso,num_orden,fec_inivig, "
        vgSql = vgSql & "cod_viapago,cod_banco,cod_tipcuenta,num_cuenta, "
        vgSql = vgSql & "cod_sucursal,cod_usuariocrea,fec_crea,hor_crea "
        vgSql = vgSql & " ) VALUES ( "
        vgSql = vgSql & "'" & Trim(Txt_PenPoliza) & "', "
        vgSql = vgSql & " " & Lbl_End & ", "
        vgSql = vgSql & " " & Lbl_NumOrd & ", "
        vgSql = vgSql & "'" & Trim(vlFecIniVig) & "', "
        vgSql = vgSql & "'" & Trim(vlCodViaPag) & "', "
        vgSql = vgSql & "'" & Trim(vlCodBco) & "', "
        vgSql = vgSql & "'" & Trim(vlCodTipcta) & "', "
        vgSql = vgSql & "'" & Trim(Txt_NumCta) & "', "
        vgSql = vgSql & "'" & Trim(vlCodSuc) & "', "
        vgSql = vgSql & "'" & Trim(vgUsuario) & "', "
        vgSql = vgSql & "'" & Format(Date, "yyyymmdd") & "', "
        vgSql = vgSql & "'" & Format(Time, "hhmmss") & "' ) "
        vgConexionBD.Execute vgSql
        
    End If
    
Exit Function
Err_flGuardarHistViaPago:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flGuardarHistPlanSalud()
On Error GoTo Err_flGuardarHistPlanSalud

    vgSql = ""
    vgSql = "SELECT s.cod_inssalud,s.cod_modsalud,s.mto_plansalud "
    vgSql = vgSql & "FROM pp_this_bensalud s WHERE "
    vgSql = vgSql & "num_poliza = '" & vlNumPoliza & "' AND "
    vgSql = vgSql & "num_orden = " & vlNumOrden & " AND "
    vgSql = vgSql & "s.fec_inivig = "
    vgSql = vgSql & "(SELECT MAX(s.fec_inivig) "
    vgSql = vgSql & "FROM pp_this_bensalud s WHERE "
    vgSql = vgSql & "num_poliza = '" & vlNumPoliza & "' AND "
    vgSql = vgSql & "num_orden = " & vlNumOrden & ") "
    Set vlRegHistoria = vgConexionBD.Execute(vgSql)
    If Not vlRegHistoria.EOF Then
        If (vlRegHistoria!Cod_InsSalud) <> Trim(vlCodIns) Then
            vlSwGrabarPlan = True
        End If
        If (vlRegHistoria!Cod_ModSalud) <> Trim(vlCodModPago) Then
            vlSwGrabarPlan = True
        End If
        If (vlRegHistoria!Mto_PlanSalud) <> CDbl(Txt_MtoPago) Then
            vlSwGrabarPlan = True
        End If
    End If
    
    If vlSwGrabarPlan = True Then
    
        vgSql = ""
        vgSql = "SELECT s.num_poliza FROM pp_this_bensalud s "
        vgSql = vgSql & "WHERE "
        vgSql = vgSql & "s.num_poliza = '" & Trim(Txt_PenPoliza) & "' AND "
        vgSql = vgSql & "s.num_orden = " & Lbl_NumOrd & " AND "
        vgSql = vgSql & "s.fec_inivig = '" & Trim(vlFecIniVig) & "' "
        Set vgRegistro = vgConexionBD.Execute(vgSql)
        If Not vgRegistro.EOF Then
            'Eliminar el registro que ya corresponde al mismo dia
            vgSql = ""
            vgSql = "DELETE pp_this_bensalud "
            vgSql = vgSql & "WHERE "
            vgSql = vgSql & "num_poliza = '" & Trim(Txt_PenPoliza) & "' AND "
            vgSql = vgSql & "num_orden = " & Lbl_NumOrd & " AND "
            vgSql = vgSql & "fec_inivig = '" & Trim(vlFecIniVig) & "' "
            vgConexionBD.Execute (vgSql)
        End If
        vgRegistro.Close
    
        vgSql = ""
        vgSql = "INSERT INTO pp_this_bensalud "
        vgSql = vgSql & "(num_poliza,num_endoso,num_orden,fec_inivig, "
        vgSql = vgSql & "cod_inssalud,cod_modsalud,mto_plansalud, "
        vgSql = vgSql & "cod_usuariocrea,fec_crea,hor_crea "
        vgSql = vgSql & " ) VALUES ( "
        vgSql = vgSql & "'" & Trim(Txt_PenPoliza) & "', "
        vgSql = vgSql & " " & Lbl_End & ", "
        vgSql = vgSql & " " & Lbl_NumOrd & ", "
        vgSql = vgSql & "'" & Trim(vlFecIniVig) & "', "
        vgSql = vgSql & "'" & Trim(vlCodIns) & "', "
        vgSql = vgSql & "'" & Trim(vlCodModPago) & "', "
        vgSql = vgSql & " " & Str(Txt_MtoPago) & ", "
        vgSql = vgSql & "'" & Trim(vgUsuario) & "', "
        vgSql = vgSql & "'" & Format(Date, "yyyymmdd") & "', "
        vgSql = vgSql & "'" & Format(Time, "hhmmss") & "' ) "
        vgConexionBD.Execute vgSql
        
    End If

Exit Function
Err_flGuardarHistPlanSalud:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flGuardarHistDireccion()
On Error GoTo Err_flGuardarHistDireccion

    vgSql = ""
    vgSql = "SELECT d.gls_dirben, d.cod_direccion "
    vgSql = vgSql & "FROM pp_this_bendir d WHERE "
    vgSql = vgSql & "num_poliza = '" & vlNumPoliza & "' AND "
    vgSql = vgSql & "num_orden = " & vlNumOrden & " AND "
    vgSql = vgSql & "d.fec_inivig = "
    vgSql = vgSql & "(SELECT MAX(d.fec_inivig) "
    vgSql = vgSql & "FROM pp_this_bendir d WHERE "
    vgSql = vgSql & "num_poliza = '" & vlNumPoliza & "' AND "
    vgSql = vgSql & "num_orden = " & vlNumOrden & ") "
    Set vlRegHistoria = vgConexionBD.Execute(vgSql)
    If Not vlRegHistoria.EOF Then
        If (vlRegHistoria!Gls_DirBen) <> Trim(Txt_DomBen) Then
            vlSwGrabarDir = True
        End If
        If (vlRegHistoria!Cod_Direccion) <> Trim(vlCodDir) Then
            vlSwGrabarDir = True
        End If
    Else
        vlSwGrabarDir = True
    End If
    
    If vlSwGrabarDir = True Then
    
        vgSql = ""
        vgSql = "SELECT d.num_poliza FROM pp_this_bendir d "
        vgSql = vgSql & "WHERE "
        vgSql = vgSql & "d.num_poliza = '" & Trim(Txt_PenPoliza) & "' AND "
        vgSql = vgSql & "d.num_orden = " & Lbl_NumOrd & " AND "
        vgSql = vgSql & "d.fec_inivig = '" & Trim(vlFecIniVig) & "' "
        Set vgRegistro = vgConexionBD.Execute(vgSql)
        If Not vgRegistro.EOF Then
            'Eliminar el registro ya que corresponde al mismo dia
            vgSql = ""
            vgSql = "DELETE pp_this_bendir "
            vgSql = vgSql & "WHERE "
            vgSql = vgSql & "num_poliza = '" & Trim(Txt_PenPoliza) & "' AND "
            vgSql = vgSql & "num_orden = " & Lbl_NumOrd & " AND "
            vgSql = vgSql & "fec_inivig = '" & Trim(vlFecIniVig) & "' "
            vgConexionBD.Execute (vgSql)
        End If
        vgRegistro.Close
    
        vgSql = ""
        vgSql = "INSERT INTO pp_this_bendir "
        vgSql = vgSql & "(num_poliza,num_endoso,num_orden,fec_inivig, "
        vgSql = vgSql & "gls_dirben,cod_direccion, "
        vgSql = vgSql & "cod_usuariocrea,fec_crea,hor_crea "
        vgSql = vgSql & " ) VALUES ( "
        vgSql = vgSql & "'" & Trim(Txt_PenPoliza) & "', "
        vgSql = vgSql & " " & Lbl_End & ", "
        vgSql = vgSql & " " & Lbl_NumOrd & ", "
        vgSql = vgSql & "'" & Trim(vlFecIniVig) & "', "
        vgSql = vgSql & "'" & Trim(Txt_DomBen) & "', "
        vgSql = vgSql & " " & vlCodDir & ", "
        vgSql = vgSql & "'" & Trim(vgUsuario) & "', "
        vgSql = vgSql & "'" & Format(Date, "yyyymmdd") & "', "
        vgSql = vgSql & "'" & Format(Time, "hhmmss") & "' ) "
        vgConexionBD.Execute vgSql
        
    End If

Exit Function
Err_flGuardarHistDireccion:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'CMV-20061031 I
'Funciones para  mostrar el dato de monto de Pensión en Quiebra
Function flUltimoPeriodoCerrado(iNumPol) As String
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
    vgSql = vgSql & " num_poliza = " & iNumPol & " "
    Set vgRs3 = vgConexionBD.Execute(vgSql)
    If vgRs3.EOF Then
        'Pago Régimen
        vlTipoPago = "R"
    Else
        'Primer Pago
        vlTipoPago = "P"
    End If
    vgRs3.Close

    'Permite obtener el último periodo que se encuentra cerrado ****
'    vgSql = ""
'    vgSql = "SELECT p.num_perpago "
'    vgSql = vgSql & "FROM pp_tmae_propagopen p "
'    vgSql = vgSql & "WHERE "
'    If vlTipoPago = "P" Then
'        vgSql = vgSql & "p.cod_estadopri = '" & clCodEstadoC & "' "
'    Else
'        vgSql = vgSql & "p.cod_estadoreg = '" & clCodEstadoC & "' "
'    End If
'    vgSql = vgSql & "ORDER BY num_perpago DESC"
'    Set vgRegistro = vgConexionBD.Execute(vgSql)
'    If Not vgRegistro.EOF Then
'        flUltimoPeriodoCerrado = Trim(vgRegistro!Num_PerPago)
'    End If



Exit Function
Err_flUltimoPeriodoCerrado:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flBuscarPensionQuiebra(iNumPerPago As String, iNumPoliza, iNumOrden As Integer, _
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

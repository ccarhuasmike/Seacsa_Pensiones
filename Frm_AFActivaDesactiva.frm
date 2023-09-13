VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Frm_AFActivaDesactiva 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Activación y Desactivación de Cargas"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   10875
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
      TabIndex        =   33
      Top             =   0
      Width           =   10575
      Begin VB.CommandButton Cmd_BuscarPol 
         Height          =   375
         Left            =   8880
         Picture         =   "Frm_AFActivaDesactiva.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Buscar Póliza"
         Top             =   225
         Width           =   615
      End
      Begin VB.TextBox Txt_PenRut 
         Height          =   285
         Left            =   4560
         MaxLength       =   10
         TabIndex        =   2
         Top             =   360
         Width           =   1755
      End
      Begin VB.TextBox Txt_PenDigito 
         Height          =   285
         Left            =   6720
         MaxLength       =   1
         TabIndex        =   3
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox Txt_PenPoliza 
         Height          =   285
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   1
         Top             =   360
         Width           =   1185
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
         Left            =   8880
         Picture         =   "Frm_AFActivaDesactiva.frx":0102
         TabIndex        =   5
         ToolTipText     =   "Buscar"
         Top             =   600
         Width           =   615
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   22
         Left            =   120
         TabIndex        =   68
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nº Endoso"
         Height          =   255
         Index           =   19
         Left            =   7200
         TabIndex        =   51
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Lbl_Endoso 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   8160
         TabIndex        =   50
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Lbl_PenNombre 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   34
         Top             =   720
         Width           =   6855
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
         Index           =   20
         Left            =   6360
         TabIndex        =   38
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Rut Pensionado"
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   37
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   36
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Póliza"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   35
         Top             =   360
         Width           =   1095
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5445
      Left            =   120
      TabIndex        =   0
      Top             =   1185
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   9604
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   12632256
      TabCaption(0)   =   "Beneficiarios Pensión Sobrev."
      TabPicture(0)   =   "Frm_AFActivaDesactiva.frx":0204
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Msf_GrillaBenfPen"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Fra_Activacion"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Ascendientes / Descendientes"
      TabPicture(1)   =   "Frm_AFActivaDesactiva.frx":0220
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Fra_NoBenef"
      Tab(1).Control(1)=   "Msf_GrillaAscDesc"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Histórico de Activaciones"
      TabPicture(2)   =   "Frm_AFActivaDesactiva.frx":023C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Msf_GrillaHist"
      Tab(2).ControlCount=   1
      Begin VB.Frame Fra_Activacion 
         Caption         =   "  Activación o Desactivación de Carga Familiar  "
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
         Top             =   2280
         Width           =   10290
         Begin VB.Frame Fra_Duplo 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   5880
            TabIndex        =   78
            Top             =   600
            Width           =   3615
            Begin VB.OptionButton Opt_NO 
               Caption         =   "No"
               Height          =   255
               Left            =   2280
               TabIndex        =   10
               Top             =   120
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.OptionButton Opt_SI 
               Caption         =   "Si"
               Height          =   255
               Left            =   1560
               TabIndex        =   9
               Top             =   120
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "Pago Duplo"
               Height          =   255
               Left            =   240
               TabIndex        =   79
               Top             =   120
               Width           =   1095
            End
         End
         Begin VB.CommandButton Cmd_Reliquidar 
            Caption         =   "&Reliquidar ..."
            Height          =   375
            Left            =   8880
            TabIndex        =   72
            Top             =   1455
            Width           =   1095
         End
         Begin VB.TextBox Txt_Digito 
            Height          =   300
            Left            =   3840
            MaxLength       =   1
            TabIndex        =   7
            Top             =   285
            Width           =   255
         End
         Begin VB.TextBox Txt_Rut 
            Height          =   300
            Left            =   2130
            MaxLength       =   10
            TabIndex        =   6
            Top             =   285
            Width           =   1425
         End
         Begin VB.Frame Fra_Suspension 
            Caption         =   "  Suspensión de Carga  "
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
            Height          =   1095
            Left            =   255
            TabIndex        =   53
            Top             =   1815
            Width           =   9735
            Begin VB.ComboBox Cmb_CauSus 
               BackColor       =   &H00E0FFFF&
               Height          =   315
               Left            =   2055
               Style           =   2  'Dropdown List
               TabIndex        =   14
               Top             =   345
               Width           =   5595
            End
            Begin VB.TextBox Txt_FecSuspension 
               Height          =   285
               Left            =   2055
               MaxLength       =   10
               TabIndex        =   15
               Top             =   720
               Width           =   1140
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "Suspensión de Carga"
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
               Index           =   25
               Left            =   240
               TabIndex        =   71
               Top             =   0
               Width           =   1920
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "Motivo de Suspensión"
               Height          =   255
               Index           =   7
               Left            =   240
               TabIndex        =   55
               Top             =   360
               Width           =   1800
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "Fecha de Suspensión"
               Height          =   255
               Index           =   8
               Left            =   240
               TabIndex        =   54
               Top             =   720
               Width           =   1710
            End
         End
         Begin VB.OptionButton Opt_NoActiva 
            Caption         =   "&No Activa"
            Height          =   255
            Left            =   3120
            TabIndex        =   12
            Top             =   1080
            Width           =   1095
         End
         Begin VB.OptionButton Opt_Activa 
            Caption         =   "Activa"
            Height          =   255
            Left            =   2145
            TabIndex        =   11
            Top             =   1080
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.TextBox Txt_FecInicio 
            Height          =   285
            Left            =   2130
            MaxLength       =   10
            TabIndex        =   13
            Top             =   1455
            Width           =   1305
         End
         Begin VB.ComboBox Cmb_SitInv 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            ItemData        =   "Frm_AFActivaDesactiva.frx":0258
            Left            =   2130
            List            =   "Frm_AFActivaDesactiva.frx":025A
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   705
            Width           =   3105
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Fecha de Efecto"
            Height          =   255
            Index           =   26
            Left            =   6165
            TabIndex        =   74
            Top             =   1470
            Width           =   1260
         End
         Begin VB.Label Lbl_Efecto 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7440
            TabIndex        =   73
            Top             =   1455
            Width           =   1305
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Activación o Desactivación de Carga Familiar"
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
            Height          =   255
            Index           =   23
            Left            =   240
            TabIndex        =   69
            Top             =   0
            Width           =   3975
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
            Left            =   3495
            TabIndex        =   67
            Top             =   1470
            Width           =   210
         End
         Begin VB.Label Lbl_NombBen 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4305
            TabIndex        =   62
            Top             =   285
            Width           =   4455
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Rut"
            Height          =   255
            Index           =   16
            Left            =   315
            TabIndex        =   61
            Top             =   360
            Width           =   615
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
            Index           =   15
            Left            =   3630
            TabIndex        =   60
            Top             =   315
            Width           =   135
         End
         Begin VB.Label Lbl_FecTermino 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3870
            TabIndex        =   59
            Top             =   1455
            Width           =   1305
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Situación de Inválidez"
            Height          =   255
            Index           =   6
            Left            =   315
            TabIndex        =   58
            Top             =   705
            Width           =   1785
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Estado de la Carga"
            Height          =   255
            Index           =   3
            Left            =   315
            TabIndex        =   57
            Top             =   1080
            Width           =   1560
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Período de Vigencia"
            Height          =   255
            Index           =   5
            Left            =   315
            TabIndex        =   56
            Top             =   1470
            Width           =   1575
         End
      End
      Begin VB.Frame Fra_NoBenef 
         Caption         =   "  Activación o Desactivación de Carga Familiar  "
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
         Left            =   -74880
         TabIndex        =   40
         Top             =   2265
         Width           =   10290
         Begin VB.Frame Fra_DuploAscDes 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   6000
            TabIndex        =   80
            Top             =   720
            Width           =   3615
            Begin VB.OptionButton Opt_SIAscDes 
               Caption         =   "Si"
               Height          =   255
               Left            =   1560
               TabIndex        =   19
               Top             =   120
               Width           =   735
            End
            Begin VB.OptionButton Opt_NOAscDes 
               Caption         =   "No"
               Height          =   255
               Left            =   2280
               TabIndex        =   20
               Top             =   120
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.Label Label2 
               Caption         =   "Pago Duplo"
               Height          =   255
               Left            =   240
               TabIndex        =   81
               Top             =   120
               Width           =   1095
            End
         End
         Begin VB.CommandButton Cmd_ReliquidarAscDes 
            Caption         =   "&Reliquidar ..."
            Height          =   375
            Left            =   9000
            TabIndex        =   75
            Top             =   1440
            Width           =   1095
         End
         Begin VB.TextBox Txt_RutAscDes 
            Height          =   300
            Left            =   2130
            MaxLength       =   10
            TabIndex        =   16
            Top             =   285
            Width           =   1425
         End
         Begin VB.TextBox Txt_DigitoAscDes 
            Height          =   300
            Left            =   3840
            MaxLength       =   1
            TabIndex        =   17
            Top             =   285
            Width           =   255
         End
         Begin VB.ComboBox Cmb_Situacion 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   2130
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   720
            Width           =   3105
         End
         Begin VB.TextBox Txt_FecIni 
            Height          =   285
            Left            =   2115
            MaxLength       =   10
            TabIndex        =   23
            Top             =   1455
            Width           =   1305
         End
         Begin VB.OptionButton Opt_Act 
            Caption         =   "Activa"
            Height          =   255
            Left            =   2145
            TabIndex        =   21
            Top             =   1080
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton Opt_NoAct 
            Caption         =   "&No Activa"
            Height          =   255
            Left            =   3120
            TabIndex        =   22
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Frame Fra_Sus 
            Caption         =   "  Suspensión de Carga  "
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
            Height          =   1095
            Left            =   255
            TabIndex        =   41
            Top             =   1815
            Width           =   9735
            Begin VB.TextBox Txt_Suspension 
               Height          =   285
               Left            =   2055
               MaxLength       =   10
               TabIndex        =   25
               Top             =   720
               Width           =   1140
            End
            Begin VB.ComboBox Cmb_Suspension 
               BackColor       =   &H00E0FFFF&
               Height          =   315
               Left            =   2055
               Style           =   2  'Dropdown List
               TabIndex        =   24
               Top             =   345
               Width           =   5595
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "Suspensión de Carga"
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
               Index           =   24
               Left            =   240
               TabIndex        =   70
               Top             =   0
               Width           =   1920
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "Fecha de Suspensión"
               Height          =   255
               Index           =   10
               Left            =   240
               TabIndex        =   43
               Top             =   720
               Width           =   1710
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "Motivo de Suspensión"
               Height          =   255
               Index           =   9
               Left            =   240
               TabIndex        =   42
               Top             =   360
               Width           =   1800
            End
         End
         Begin VB.Label Lbl_EfectoAscDes 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7455
            TabIndex        =   77
            Top             =   1455
            Width           =   1305
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Fecha de Efecto"
            Height          =   255
            Index           =   27
            Left            =   6165
            TabIndex        =   76
            Top             =   1470
            Width           =   1260
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
            Index           =   11
            Left            =   3495
            TabIndex        =   66
            Top             =   1470
            Width           =   210
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
            Index           =   18
            Left            =   3630
            TabIndex        =   65
            Top             =   315
            Width           =   135
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Rut"
            Height          =   255
            Index           =   17
            Left            =   315
            TabIndex        =   64
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Lbl_NombAsc 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4305
            TabIndex        =   63
            Top             =   285
            Width           =   4455
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Período de Vigencia"
            Height          =   255
            Index           =   14
            Left            =   315
            TabIndex        =   47
            Top             =   1470
            Width           =   1575
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Estado de la Carga"
            Height          =   255
            Index           =   13
            Left            =   315
            TabIndex        =   46
            Top             =   1080
            Width           =   1560
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Situación de Inválidez"
            Height          =   255
            Index           =   12
            Left            =   315
            TabIndex        =   45
            Top             =   705
            Width           =   1785
         End
         Begin VB.Label Lbl_FecTer 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3840
            TabIndex        =   44
            Top             =   1455
            Width           =   1305
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Msf_GrillaBenfPen 
         Height          =   1755
         Left            =   105
         TabIndex        =   32
         Top             =   465
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   3096
         _Version        =   393216
         Cols            =   17
         BackColor       =   14745599
         FormatString    =   $"Frm_AFActivaDesactiva.frx":025C
      End
      Begin MSFlexGridLib.MSFlexGrid Msf_GrillaAscDesc 
         Height          =   1755
         Left            =   -74880
         TabIndex        =   39
         Top             =   480
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   3096
         _Version        =   393216
         Cols            =   16
         BackColor       =   14745599
         FormatString    =   $"Frm_AFActivaDesactiva.frx":0320
      End
      Begin MSFlexGridLib.MSFlexGrid Msf_GrillaHist 
         Height          =   4755
         Left            =   -74895
         TabIndex        =   48
         Top             =   465
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   8387
         _Version        =   393216
         Cols            =   16
         BackColor       =   14745599
         FormatString    =   $"Frm_AFActivaDesactiva.frx":03E3
      End
   End
   Begin VB.Frame Fra_Botones 
      Height          =   975
      Left            =   120
      TabIndex        =   49
      Top             =   6600
      Width           =   10605
      Begin VB.CommandButton Cmd_Grabar 
         Caption         =   "&Grabar"
         Height          =   660
         Left            =   1845
         Picture         =   "Frm_AFActivaDesactiva.frx":04A8
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   210
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   660
         Left            =   5040
         Picture         =   "Frm_AFActivaDesactiva.frx":0B62
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   210
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   7200
         Picture         =   "Frm_AFActivaDesactiva.frx":121C
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   660
         Left            =   3960
         Picture         =   "Frm_AFActivaDesactiva.frx":1316
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Eliminar 
         Caption         =   "&Eliminar"
         Height          =   660
         Left            =   2880
         Picture         =   "Frm_AFActivaDesactiva.frx":19D0
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Eliminar Año"
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   6120
         Picture         =   "Frm_AFActivaDesactiva.frx":1D12
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   200
         Width           =   730
      End
      Begin Crystal.CrystalReport Rpt_AsigFam 
         Left            =   9480
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
End
Attribute VB_Name = "Frm_AFActivaDesactiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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

    Frm_AFActivaDesactiva.Top = 0
    Frm_AFActivaDesactiva.Left = 0
            
    vlCodIndReliquidar = ""
            
    SSTab1.Tab = 0
    SSTab1.Enabled = True
   
    Opt_NoActiva.Value = True
    Opt_NoAct.Value = True
    Cmd_Reliquidar.Visible = True
    
    Cmd_ReliquidarAscDes.Visible = True
    
    
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub


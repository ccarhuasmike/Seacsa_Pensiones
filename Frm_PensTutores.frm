VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Frm_PensTutores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de Tutores"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   9030
   Begin VB.Frame Fra_Operacion 
      Height          =   990
      Left            =   120
      TabIndex        =   52
      Top             =   6840
      Width           =   8775
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   5280
         Picture         =   "Frm_PensTutores.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Eliminar 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   4080
         Picture         =   "Frm_PensTutores.frx":06BA
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   2760
         Picture         =   "Frm_PensTutores.frx":09FC
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   6600
         Picture         =   "Frm_PensTutores.frx":10B6
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Salir del Formulario"
         Top             =   215
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Grabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   1440
         Picture         =   "Frm_PensTutores.frx":11B0
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Grabar Datos"
         Top             =   240
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Antecedentes del Tutor"
      TabPicture(0)   =   "Frm_PensTutores.frx":186A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Fra_Vigencia"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Fra_Antecedentes"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Fra_Pago"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Historia de Tutores"
      TabPicture(1)   =   "Frm_PensTutores.frx":1886
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame6 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   46
         Top             =   480
         Width           =   8535
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
            Height          =   4575
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   8070
            _Version        =   393216
            Cols            =   4
            FixedCols       =   0
            BackColor       =   14745599
            FormatString    =   $"Frm_PensTutores.frx":18A2
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
         Height          =   1455
         Left            =   120
         TabIndex        =   35
         Top             =   3960
         Width           =   8535
         Begin VB.ComboBox Combo5 
            Height          =   315
            Left            =   5280
            TabIndex        =   40
            Top             =   240
            Width           =   2940
         End
         Begin VB.ComboBox Combo4 
            Height          =   315
            Left            =   1080
            TabIndex        =   39
            Top             =   315
            Width           =   2955
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            Left            =   5280
            TabIndex        =   38
            Top             =   645
            Width           =   2940
         End
         Begin VB.TextBox Text12 
            Height          =   285
            Left            =   5280
            TabIndex        =   37
            Top             =   1035
            Width           =   2910
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   1080
            TabIndex        =   36
            Top             =   720
            Width           =   2955
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Tipo Cta."
            Height          =   255
            Index           =   17
            Left            =   4320
            TabIndex        =   45
            Top             =   240
            Width           =   945
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Vía Pago"
            Height          =   255
            Index           =   15
            Left            =   165
            TabIndex        =   44
            Top             =   360
            Width           =   930
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Banco"
            Height          =   255
            Index           =   18
            Left            =   4320
            TabIndex        =   43
            Top             =   675
            Width           =   945
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "N°Cuenta"
            Height          =   255
            Index           =   19
            Left            =   4320
            TabIndex        =   42
            Top             =   1035
            Width           =   915
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Sucursal"
            Height          =   255
            Index           =   16
            Left            =   165
            TabIndex        =   41
            Top             =   720
            Width           =   900
         End
      End
      Begin VB.Frame Fra_Antecedentes 
         Caption         =   "Antecedenes Personales del Tutor"
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
         Height          =   2655
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   8535
         Begin VB.TextBox Text2 
            BackColor       =   &H00E0FFFF&
            Height          =   285
            Left            =   3500
            TabIndex        =   49
            Top             =   1800
            Width           =   2385
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00E0FFFF&
            Height          =   285
            Left            =   6000
            TabIndex        =   48
            Top             =   1800
            Width           =   2385
         End
         Begin VB.TextBox Text21 
            Height          =   285
            Left            =   1080
            TabIndex        =   27
            Top             =   2160
            Width           =   4380
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   1080
            TabIndex        =   26
            Top             =   1800
            Width           =   2355
         End
         Begin VB.TextBox Text13 
            Height          =   285
            Left            =   5400
            TabIndex        =   25
            Top             =   1080
            Width           =   3015
         End
         Begin VB.TextBox Text11 
            Height          =   285
            Left            =   1080
            TabIndex        =   24
            Top             =   1080
            Width           =   3015
         End
         Begin VB.TextBox Text20 
            Height          =   285
            Left            =   1080
            TabIndex        =   23
            Top             =   1440
            Width           =   7335
         End
         Begin VB.TextBox Text19 
            Height          =   285
            Left            =   1080
            TabIndex        =   22
            Top             =   720
            Width           =   3015
         End
         Begin VB.TextBox Text18 
            Height          =   285
            Left            =   3120
            TabIndex        =   21
            Top             =   360
            Width           =   405
         End
         Begin VB.TextBox Text17 
            Height          =   285
            Left            =   6360
            TabIndex        =   20
            Top             =   2160
            Width           =   2025
         End
         Begin VB.TextBox Text16 
            Height          =   285
            Left            =   1080
            TabIndex        =   19
            Top             =   360
            Width           =   1620
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Ap. Materno"
            Height          =   255
            Index           =   10
            Left            =   4440
            TabIndex        =   51
            Top             =   1080
            Width           =   915
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Ap. Paterno"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   50
            Top             =   1080
            Width           =   915
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Email"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   34
            Top             =   2160
            Width           =   795
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Comuna"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   33
            Top             =   1800
            Width           =   855
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
            Index           =   7
            Left            =   2760
            TabIndex        =   32
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Dirección"
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   31
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Nombre"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   30
            Top             =   720
            Width           =   795
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Teléfono"
            Height          =   255
            Index           =   14
            Left            =   5640
            TabIndex        =   29
            Top             =   2160
            Width           =   795
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Rut"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   28
            Top             =   405
            Width           =   810
         End
      End
      Begin VB.Frame Fra_Vigencia 
         Caption         =   "Vigencia del Poder"
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
         Height          =   765
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   8535
         Begin VB.TextBox Text10 
            Height          =   285
            Left            =   1680
            TabIndex        =   14
            Top             =   360
            Width           =   690
         End
         Begin VB.TextBox Text9 
            BackColor       =   &H00E0FFFF&
            Height          =   285
            Left            =   6240
            TabIndex        =   13
            Top             =   360
            Width           =   1185
         End
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   3720
            TabIndex        =   12
            Top             =   360
            Width           =   1200
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Hasta"
            Height          =   255
            Index           =   5
            Left            =   5520
            TabIndex        =   17
            Top             =   360
            Width           =   465
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Desde"
            Height          =   255
            Index           =   4
            Left            =   3000
            TabIndex        =   16
            Top             =   360
            Width           =   555
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Duración (meses)"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   15
            Top             =   360
            Width           =   1380
         End
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
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      Begin VB.CommandButton Cmd_Buscar 
         Height          =   375
         Left            =   7920
         Picture         =   "Frm_PensTutores.frx":193A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Buscar Póliza"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   3960
         TabIndex        =   5
         Top             =   360
         Width           =   1755
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00E0FFFF&
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   720
         Width           =   6825
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   6100
         TabIndex        =   3
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label Label2 
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
         Left            =   5760
         TabIndex        =   8
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Rut Pensionado"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Póliza"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "Frm_PensTutores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Btn_Salir_Click()
    Unload Me
End Sub

Private Sub Cmd_Buscar_Click()
    Frm_Busqueda.Show 1
End Sub

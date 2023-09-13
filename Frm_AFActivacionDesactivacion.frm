VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form Frm_AFActivacionDesactivacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Activación y Desactivación de Cargas"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   10140
   Begin VB.CommandButton Command3 
      Caption         =   "Cerrar"
      Height          =   420
      Left            =   8865
      TabIndex        =   26
      Top             =   8235
      Width           =   945
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Eliminar"
      Height          =   465
      Left            =   1845
      TabIndex        =   25
      Top             =   8205
      Width           =   1050
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Grabar"
      Height          =   450
      Left            =   345
      TabIndex        =   24
      Top             =   8205
      Width           =   1200
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   3255
      TabIndex        =   19
      Top             =   390
      Width           =   270
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0FFFF&
      Height          =   375
      Left            =   4230
      TabIndex        =   18
      Top             =   390
      Width           =   2805
   End
   Begin VB.TextBox Text4 
      Height          =   330
      Left            =   915
      TabIndex        =   17
      Top             =   405
      Width           =   900
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00E0FFFF&
      Height          =   345
      Left            =   8415
      TabIndex        =   16
      Top             =   405
      Width           =   1410
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   2235
      TabIndex        =   15
      Top             =   405
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6765
      Left            =   300
      TabIndex        =   0
      Top             =   1200
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   11933
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Beneficiarios Pensión Sobrev."
      TabPicture(0)   =   "Frm_AFActivacionDesactivacion.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "MSFlexGrid2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Ascendientes / Descendientes"
      TabPicture(1)   =   "Frm_AFActivacionDesactivacion.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MSFlexGrid1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Histórico de Activaciones"
      TabPicture(2)   =   "Frm_AFActivacionDesactivacion.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   2235
         Left            =   510
         TabIndex        =   14
         Top             =   735
         Width           =   8310
         _ExtentX        =   14658
         _ExtentY        =   3942
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         FormatString    =   "Parentesco|   Rut           | DV |    Nombre                                                 |  Sit.Invalidez | Fecha Nacimiento  "
      End
      Begin VB.Frame Frame1 
         Caption         =   "Activación o Desactivación de Cargas Familiares"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2790
         Left            =   480
         TabIndex        =   1
         Top             =   3300
         Width           =   8370
         Begin VB.TextBox Text6 
            Height          =   315
            Left            =   1920
            TabIndex        =   29
            Top             =   2280
            Width           =   1140
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   1920
            TabIndex        =   27
            Top             =   1425
            Width           =   1980
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   1935
            TabIndex        =   6
            Top             =   1890
            Width           =   4875
         End
         Begin VB.CheckBox Check2 
            Caption         =   "No"
            Height          =   210
            Left            =   2370
            TabIndex        =   5
            Top             =   615
            Width           =   765
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Sí"
            Height          =   195
            Left            =   1455
            TabIndex        =   4
            Top             =   585
            Width           =   750
         End
         Begin VB.TextBox Text9 
            Height          =   315
            Left            =   3825
            TabIndex        =   3
            Top             =   975
            Width           =   900
         End
         Begin VB.TextBox Text8 
            Height          =   315
            Left            =   1950
            TabIndex        =   2
            Top             =   975
            Width           =   900
         End
         Begin VB.Label Label10 
            Caption         =   "Fecha de Suspensión :"
            Height          =   270
            Left            =   225
            TabIndex        =   28
            Top             =   2385
            Width           =   1710
         End
         Begin VB.Label Label4 
            Caption         =   "Motivo de Suspensión:"
            Height          =   300
            Left            =   210
            TabIndex        =   12
            Top             =   1965
            Width           =   1920
         End
         Begin VB.Label Label6 
            Caption         =   "Tipo de Inválidez:"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   1500
            Width           =   1305
         End
         Begin VB.Label Label9 
            Caption         =   "Activa carga:"
            Height          =   270
            Left            =   240
            TabIndex        =   10
            Top             =   585
            Width           =   1200
         End
         Begin VB.Label Label8 
            Caption         =   "Hasta"
            Height          =   285
            Left            =   3255
            TabIndex        =   9
            Top             =   1020
            Width           =   435
         End
         Begin VB.Label Label7 
            Caption         =   "Desde"
            Height          =   285
            Left            =   1425
            TabIndex        =   8
            Top             =   1020
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Período :"
            Height          =   270
            Left            =   225
            TabIndex        =   7
            Top             =   1005
            Width           =   765
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   1365
         Left            =   -74745
         TabIndex        =   13
         Top             =   1290
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   2408
         _Version        =   393216
      End
   End
   Begin VB.Line Line1 
      X1              =   315
      X2              =   9840
      Y1              =   960
      Y2              =   990
   End
   Begin VB.Label Label5 
      Caption         =   "Nombre"
      Height          =   285
      Left            =   3585
      TabIndex        =   23
      Top             =   405
      Width           =   750
   End
   Begin VB.Label Label3 
      Caption         =   "N° Póliza"
      Height          =   315
      Left            =   225
      TabIndex        =   22
      Top             =   450
      Width           =   750
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo de Pensión"
      Height          =   300
      Left            =   7185
      TabIndex        =   21
      Top             =   450
      Width           =   1200
   End
   Begin VB.Label Label17 
      Caption         =   "Rut"
      Height          =   300
      Left            =   1905
      TabIndex        =   20
      Top             =   405
      Width           =   660
   End
End
Attribute VB_Name = "Frm_AFActivacionDesactivacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command3_Click()
    Unload Me
    
End Sub

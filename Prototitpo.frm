VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form Form1 
   Caption         =   "Atencion de Benefciarios Hijos"
   ClientHeight    =   10110
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12585
   LinkTopic       =   "Form1"
   ScaleHeight     =   10110
   ScaleWidth      =   12585
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Cartas"
      Height          =   4935
      Left            =   120
      TabIndex        =   10
      Top             =   4080
      Width           =   12375
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   4080
         TabIndex        =   17
         Text            =   "Selecione Mes"
         Top             =   360
         Width           =   1455
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4095
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   7223
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Vigentes"
         TabPicture(0)   =   "Prototitpo.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Label2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "MSFlexGrid1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Vencidas"
         TabPicture(1)   =   "Prototitpo.frx":001C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label3"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "MSFlexGrid2"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Command3"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Check2"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).ControlCount=   4
         Begin VB.CheckBox Check2 
            Caption         =   "Todas"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Carta 1 Meses"
            Height          =   375
            Left            =   120
            TabIndex        =   24
            Top             =   3600
            Width           =   1335
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
            Height          =   2775
            Left            =   -74880
            TabIndex        =   12
            Top             =   720
            Width           =   11895
            _ExtentX        =   20981
            _ExtentY        =   4895
            _Version        =   393216
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
            Height          =   2775
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   11895
            _ExtentX        =   20981
            _ExtentY        =   4895
            _Version        =   393216
         End
         Begin VB.Label Label3 
            Caption         =   "Son un total de "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9840
            TabIndex        =   15
            Top             =   3480
            Width           =   2175
         End
         Begin VB.Label Label2 
            Caption         =   "Son un total de "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -65160
            TabIndex        =   14
            Top             =   3480
            Width           =   2175
         End
      End
      Begin VB.Label Label7 
         Caption         =   "MAYO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   23
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Mes Actual"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Cartas Emitidas en Periodo:"
         Height          =   255
         Left            =   2040
         TabIndex        =   16
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Carga Masiva"
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   9000
      Width           =   12375
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         Height          =   255
         Left            =   11400
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3600
         TabIndex        =   7
         Top             =   360
         Width           =   7695
      End
      Begin VB.Label Label5 
         Caption         =   "Carga el archivo de cartas repecionadas "
         Height          =   495
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Busqueda Beneficiarios"
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12375
      Begin VB.CheckBox Check1 
         Caption         =   "Todas"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Imprime Carta 6 Meses"
         Height          =   375
         Left            =   7920
         TabIndex        =   19
         Top             =   3480
         Width           =   2775
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Exportar Excel"
         Height          =   375
         Left            =   10680
         TabIndex        =   18
         Top             =   3480
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
         Height          =   2535
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   4471
         _Version        =   393216
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Acreditado"
         Height          =   375
         Left            =   7440
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Sin Acreditacion"
         Height          =   375
         Left            =   5520
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   255
         Left            =   4680
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3000
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Proximos a cumplir 18 años  hasta el "
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2775
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_PensPrimerosPagos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cálculo de Primeros Pagos"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   5160
   Begin VB.Frame Frame1 
      Caption         =   "Parámetros del Cálculo"
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
      TabIndex        =   5
      Top             =   120
      Width           =   4935
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0FFFF&
         Height          =   285
         Left            =   3720
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txt_FecPago 
         BackColor       =   &H00E0FFFF&
         Height          =   285
         Left            =   2040
         TabIndex        =   11
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txt_FecCalculo 
         BackColor       =   &H00E0FFFF&
         Height          =   285
         Left            =   2040
         TabIndex        =   10
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txt_UF 
         BackColor       =   &H00E0FFFF&
         Height          =   285
         Left            =   2040
         TabIndex        =   9
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txt_TipoCalculo 
         BackColor       =   &H00E0FFFF&
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
         Left            =   2040
         TabIndex        =   8
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txt_Periodo 
         BackColor       =   &H00E0FFFF&
         Height          =   285
         Left            =   2040
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txt_FecProxPago 
         BackColor       =   &H00E0FFFF&
         Height          =   285
         Left            =   2040
         TabIndex        =   6
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Hasta"
         Height          =   255
         Left            =   3240
         TabIndex        =   19
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de Cálculo"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha de Pago"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de Cambio"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo de Cálculo"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Pólizas recibidas Desde"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha Próximo Pago"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   4455
      Begin VB.CommandButton cmd_calcular 
         Caption         =   "&Calcular"
         Height          =   675
         Left            =   840
         Picture         =   "Frm_PensPrimerosPagos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Realizar Cálculo de Primeros Pagos de Pensiones"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   2640
         Picture         =   "Frm_PensPrimerosPagos.frx":04A2
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Progreso del Cálculo"
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
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   4455
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
   End
End
Attribute VB_Name = "Frm_PensPrimerosPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_salir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Frm_Menu.Tag = "0"
End Sub



VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_BarraProg 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Progreso de Carga"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6675
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   Icon            =   "Frm_BarraProg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   6675
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Lbl_Texto 
      Caption         =   "Label1"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   6135
   End
   Begin VB.Label Lbl_Estado 
      Caption         =   "Cargando     :"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   2655
   End
End
Attribute VB_Name = "Frm_BarraProg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
    Frm_BarraProg.Left = 800
    Frm_BarraProg.Top = 800
    Lbl_Texto = ""
End Sub

Private Sub Form_GotFocus()
    Frm_BarraProg.Left = 800
    Frm_BarraProg.Top = 800
    Lbl_Texto = ""
End Sub

Private Sub Form_Initialize()
    Frm_BarraProg.Left = 800
    Frm_BarraProg.Top = 800
    Lbl_Texto = ""
    'Timer1.Enabled = True
End Sub

Private Sub Form_Load()
    'Frm_BarraProg.Refresh Me
    Frm_BarraProg.Left = 800
    Frm_BarraProg.Top = 800
    Lbl_Texto = ""
    'Timer1.Enabled = True
End Sub


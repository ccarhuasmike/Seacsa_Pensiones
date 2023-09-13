VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Frm_CCAFImportar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carga Automática de Descuentos CCAF."
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   8910
   Begin MSComDlg.CommonDialog ComDialogo 
      Left            =   480
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Fra_Fondo 
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
      Height          =   855
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   8655
      Begin VB.ComboBox Cmb_Ccaf 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox Txt_Fecha 
         Height          =   315
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   0
         ToolTipText     =   "Fecha utilizada para validar los datos de Carga"
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Caja de Compensación :"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   14
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Fecha de Pago :"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
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
      TabIndex        =   9
      Top             =   960
      Width           =   8655
      Begin VB.CommandButton CmdPolizas 
         Height          =   375
         Left            =   7680
         Picture         =   "Frm_CCAFImportar.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         Alignment       =   2  'Center
         Caption         =   "Selección de Archivos"
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
         Index           =   2
         Left            =   2760
         TabIndex        =   12
         Top             =   0
         Width           =   2415
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Archivo"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   855
      End
      Begin VB.Label LblPolizas 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   10
         Top             =   480
         Width           =   6135
      End
   End
   Begin VB.Frame Fra_Operacion 
      Height          =   1095
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   8655
      Begin VB.CommandButton Cmd_ImpEstadistica 
         Caption         =   "&Resumen"
         Height          =   675
         Left            =   2880
         Picture         =   "Frm_CCAFImportar.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Imprimir Estadísticas"
         Top             =   240
         Width           =   790
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Errores"
         Height          =   675
         Left            =   4200
         Picture         =   "Frm_CCAFImportar.frx":07BC
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Imprimir Errores"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5400
         Picture         =   "Frm_CCAFImportar.frx":0E76
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Cargar 
         Caption         =   "&Cargar"
         Height          =   675
         Left            =   1680
         Picture         =   "Frm_CCAFImportar.frx":0F70
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Carga de Datos"
         Top             =   240
         Width           =   720
      End
      Begin Crystal.CrystalReport Rpt_Calculo 
         Left            =   7440
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
End
Attribute VB_Name = "Frm_CCAFImportar"
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

    Frm_CCAFImportar.Top = 0
    Frm_CCAFImportar.Left = 0

    
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub


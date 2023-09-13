VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Frm_GECargaRecup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carga de Archivo de Montos Recuperados al Estado."
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   8895
   Begin VB.Frame Fra_Operacion 
      Height          =   1095
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   8655
      Begin VB.CommandButton Cmd_Cargar 
         Caption         =   "&Cargar"
         Height          =   675
         Left            =   1680
         Picture         =   "Frm_GECargaRecup.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Carga de Datos"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5400
         Picture         =   "Frm_GECargaRecup.frx":0822
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_ImpErrores 
         Caption         =   "&Errores"
         Height          =   675
         Left            =   4200
         Picture         =   "Frm_GECargaRecup.frx":091C
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Imprimir Errores"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_ImpResumen 
         Caption         =   "&Resumen"
         Height          =   675
         Left            =   2880
         Picture         =   "Frm_GECargaRecup.frx":0FD6
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Imprimir Estadísticas"
         Top             =   240
         Width           =   790
      End
      Begin Crystal.CrystalReport Rpt_General 
         Left            =   7080
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
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
      ForeColor       =   &H00800000&
      Height          =   1095
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   8655
      Begin VB.CommandButton Cmd_Buscar 
         Height          =   375
         Left            =   7680
         Picture         =   "Frm_GECargaRecup.frx":1690
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Lbl_Archivo 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         Top             =   480
         Width           =   6135
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Archivo"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   480
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
         TabIndex        =   5
         Top             =   0
         Width           =   2415
      End
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
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.TextBox Txt_Periodo 
         Height          =   315
         Left            =   3600
         MaxLength       =   10
         TabIndex        =   1
         ToolTipText     =   "Fecha utilizada para validar los datos de Carga"
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Fecha de Pago            :"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   2175
      End
   End
   Begin MSComDlg.CommonDialog ComDialogo 
      Left            =   360
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Frm_GECargaRecup"
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

   Frm_GECargaRecup.Top = 0
   Frm_GECargaRecup.Left = 0
   
   vlNumArchivo = ""
   

Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub


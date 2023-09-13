VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Frm_PlanillaSVS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planilla SBS por Periodo"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   6165
   Begin MSComDlg.CommonDialog ComDialogo 
      Left            =   5400
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Fra_Busqueda 
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   5925
      Begin VB.TextBox Txt_Mes 
         Height          =   285
         Left            =   2880
         MaxLength       =   2
         TabIndex        =   2
         Top             =   780
         Width           =   360
      End
      Begin VB.TextBox Txt_Anno 
         Height          =   285
         Left            =   3600
         MaxLength       =   4
         TabIndex        =   3
         Top             =   780
         Width           =   795
      End
      Begin VB.ComboBox Cmb_Tipo 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label lbl_nombre 
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
         Left            =   3315
         TabIndex        =   13
         Top             =   795
         Width           =   195
      End
      Begin VB.Label lbl_nombre 
         Caption         =   "(Mes - Año)"
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
         Height          =   270
         Index           =   10
         Left            =   1140
         TabIndex        =   12
         Top             =   795
         Width           =   1125
      End
      Begin VB.Label lbl_nombre 
         Caption         =   "Período "
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
         Index           =   8
         Left            =   285
         TabIndex        =   11
         Top             =   780
         Width           =   705
      End
      Begin VB.Label lbl_nombre 
         Caption         =   "Tipo de Proceso             :"
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
         Left            =   285
         TabIndex        =   10
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lbl_nombre 
         Caption         =   " :"
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
         Height          =   240
         Index           =   0
         Left            =   2400
         TabIndex        =   9
         Top             =   825
         Width           =   165
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1035
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   5895
      Begin VB.CommandButton Cmd_Cargar 
         Caption         =   "&Archivo"
         Height          =   675
         Left            =   3000
         Picture         =   "Frm_PlanillaSVS.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Exportar Datos a Archivo"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   1680
         Picture         =   "Frm_PlanillaSVS.frx":0822
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   4320
         Picture         =   "Frm_PlanillaSVS.frx":0EDC
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   270
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   720
         Picture         =   "Frm_PlanillaSVS.frx":0FD6
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   720
      End
      Begin Crystal.CrystalReport Rpt_Reporte 
         Left            =   5280
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Line Lin_Separar 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   3960
         X2              =   3960
         Y1              =   240
         Y2              =   900
      End
      Begin VB.Line Lin_Separar 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   2760
         X2              =   2760
         Y1              =   240
         Y2              =   900
      End
   End
End
Attribute VB_Name = "Frm_PlanillaSVS"
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

    Frm_PlanillaSVS.Left = 0
    Frm_PlanillaSVS.Top = 0
    
   

Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub


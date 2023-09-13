VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Frm_PlanillaEstadoBenGE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planilla por Estado de Beneficiario G.E."
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   6120
   Begin VB.Frame Fra_Operaciones 
      Height          =   1095
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   5895
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   3840
         Picture         =   "Frm_PlanillaEstadoBenGE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   2640
         Picture         =   "Frm_PlanillaEstadoBenGE.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   1440
         Picture         =   "Frm_PlanillaEstadoBenGE.frx":07B4
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   720
      End
      Begin Crystal.CrystalReport Rpt_Reporte 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowState     =   2
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.Frame Fra_Datos 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.TextBox Txt_Hasta 
         Height          =   285
         Left            =   4320
         MaxLength       =   10
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Txt_Desde 
         Height          =   285
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox Cmb_Estado 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Lbl_Contrato 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   4080
         TabIndex        =   10
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Lbl_Contrato 
         Caption         =   "Período (Desde - Hasta)  :"
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
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Lbl_Contrato 
         Caption         =   "Tipo Estado                   :"
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
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   2415
      End
   End
End
Attribute VB_Name = "Frm_PlanillaEstadoBenGE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'CMV / 20050401

Dim vlArchivo As String

Dim vlFechaInicio As String
Dim vlFechaTermino As String
Dim vlFechaDesde As String
Dim vlFechaHasta As String
Dim vlQuery As String
Dim vlEstado As String

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

    Frm_PlanillaEstadoBenGE.Left = 0
    Frm_PlanillaEstadoBenGE.Top = 0
    
    'fgComboEstadoTodo Cmb_Estado

Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

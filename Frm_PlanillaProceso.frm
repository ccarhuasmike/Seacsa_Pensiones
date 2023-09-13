VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Frm_PlanillaProceso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planilla Masiva y por Pensionado"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   6285
   Begin VB.Frame Fra_Operaciones 
      Height          =   1095
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   3120
      Width           =   6015
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   3840
         Picture         =   "Frm_PlanillaProceso.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   2640
         Picture         =   "Frm_PlanillaProceso.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   1440
         Picture         =   "Frm_PlanillaProceso.frx":07B4
         Style           =   1  'Graphical
         TabIndex        =   8
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
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.TextBox Txt_Hasta 
         Height          =   285
         Left            =   4320
         MaxLength       =   10
         TabIndex        =   3
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Txt_Desde 
         Height          =   285
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox Cmb_Tipo 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   2775
      End
      Begin VB.Frame Frame1 
         Height          =   1455
         Left            =   360
         TabIndex        =   11
         Top             =   1320
         Width           =   5175
         Begin VB.CheckBox Chk_Pensionado 
            Caption         =   "Datos del Pensionado :"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   240
            Width           =   2415
         End
         Begin VB.TextBox Txt_Digito 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4320
            MaxLength       =   1
            TabIndex        =   7
            Top             =   960
            Width           =   375
         End
         Begin VB.TextBox Txt_Rut 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2520
            MaxLength       =   10
            TabIndex        =   6
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox Txt_Poliza 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2520
            MaxLength       =   10
            TabIndex        =   5
            Top             =   600
            Width           =   1215
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
            Index           =   5
            Left            =   3960
            TabIndex        =   14
            Top             =   960
            Width           =   135
         End
         Begin VB.Label Lbl_Contrato 
            Caption         =   "Rut                        :"
            Height          =   255
            Index           =   3
            Left            =   840
            TabIndex        =   13
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Lbl_Contrato 
            Caption         =   "Nº de Póliza          :"
            Height          =   255
            Index           =   2
            Left            =   840
            TabIndex        =   12
            Top             =   600
            Width           =   1455
         End
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
         TabIndex        =   17
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Lbl_Contrato 
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
         Index           =   6
         Left            =   360
         TabIndex        =   16
         Top             =   360
         Width           =   2415
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
         Left            =   360
         TabIndex        =   15
         Top             =   840
         Width           =   2295
      End
   End
End
Attribute VB_Name = "Frm_PlanillaProceso"
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

    Frm_PlanillaProceso.Left = 0
    Frm_PlanillaProceso.Top = 0
    
 
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub



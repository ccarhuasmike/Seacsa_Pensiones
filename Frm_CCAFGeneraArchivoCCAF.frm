VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Frm_CCAFGeneraArchivoCCAF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carga Archivo CCAF"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   7470
   Begin VB.Frame Fra_Datos 
      Height          =   1935
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   7215
      Begin VB.ComboBox Cmb_TipoConcepto 
         BackColor       =   &H00E0FFFF&
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1110
         Width           =   4635
      End
      Begin VB.ComboBox Cmb_CCAF 
         BackColor       =   &H00E0FFFF&
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   700
         Width           =   4635
      End
      Begin VB.TextBox Txt_Desde 
         Height          =   285
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1520
         Width           =   1215
      End
      Begin VB.TextBox Txt_Hasta 
         Height          =   285
         Left            =   3960
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1520
         Width           =   1215
      End
      Begin VB.ComboBox Cmb_Tipo 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   300
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Concepto"
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
         Left            =   120
         TabIndex        =   14
         Top             =   1160
         Width           =   1755
      End
      Begin VB.Label Label2 
         Caption         =   "Caja de Compensación"
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
         Left            =   120
         TabIndex        =   13
         Top             =   750
         Width           =   2115
      End
      Begin VB.Label Lbl_Nombre 
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
         Index           =   1
         Left            =   3720
         TabIndex        =   12
         Top             =   1520
         Width           =   135
      End
      Begin VB.Label Lbl_Contrato 
         Caption         =   "Período (Desde - Hasta)"
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
         Left            =   120
         TabIndex        =   11
         Top             =   1570
         Width           =   2295
      End
      Begin VB.Label Lbl_Contrato 
         Caption         =   "Tipo de Proceso "
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
         Left            =   120
         TabIndex        =   10
         Top             =   350
         Width           =   1815
      End
   End
   Begin VB.Frame Fra_Operaciones 
      Height          =   1095
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   2000
      Width           =   7215
      Begin VB.CommandButton Cmd_Cargar 
         Caption         =   "&Archivo"
         Height          =   675
         Left            =   2160
         Picture         =   "Frm_CCAFGeneraArchivoCCAF.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Exportar Datos a Archivo"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   3360
         Picture         =   "Frm_CCAFGeneraArchivoCCAF.frx":0822
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   4560
         Picture         =   "Frm_CCAFGeneraArchivoCCAF.frx":0EDC
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
      Begin MSComDlg.CommonDialog ComDialogo 
         Left            =   960
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "Frm_CCAFGeneraArchivoCCAF"
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
On Error GoTo Err_Carga
    
    Frm_CCAFGeneraArchivoCCAF.Left = 0
    Frm_CCAFGeneraArchivoCCAF.Top = 0
    
Exit Sub
Err_Carga:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub




















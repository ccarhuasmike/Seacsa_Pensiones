VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Frm_CargaArchivoAnno 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carga de Archivo"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   4980
   Begin VB.Frame Fra_Operaciones 
      Height          =   1095
      Index           =   1
      Left            =   80
      TabIndex        =   10
      Top             =   2760
      Width           =   4815
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   2040
         Picture         =   "Frm_CargaArchivoAnno.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   3000
         Picture         =   "Frm_CargaArchivoAnno.frx":06BA
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Cargar 
         Caption         =   "&Archivo"
         Height          =   675
         Left            =   1080
         Picture         =   "Frm_CargaArchivoAnno.frx":07B4
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Exportar Datos a Archivo"
         Top             =   240
         Width           =   720
      End
      Begin MSComDlg.CommonDialog ComDialogo 
         Left            =   120
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Frame Fra_Datos 
      Height          =   2655
      Left            =   80
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.Frame Frame1 
         Height          =   1455
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   4575
         Begin VB.CheckBox Chk_Pensionado 
            Caption         =   "Datos del Pensionado :"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   2
            Top             =   240
            Width           =   2415
         End
         Begin VB.TextBox Txt_Digito 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3840
            MaxLength       =   1
            TabIndex        =   5
            Top             =   960
            Width           =   375
         End
         Begin VB.TextBox Txt_Rut 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2040
            MaxLength       =   10
            TabIndex        =   4
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox Txt_Poliza 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2040
            MaxLength       =   10
            TabIndex        =   3
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
            Left            =   3480
            TabIndex        =   14
            Top             =   960
            Width           =   135
         End
         Begin VB.Label Lbl_Contrato 
            Caption         =   "Rut                        :"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   13
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Lbl_Contrato 
            Caption         =   "Nº de Póliza          :"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   12
            Top             =   600
            Width           =   1455
         End
      End
      Begin VB.TextBox Txt_Anno 
         Height          =   285
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Lbl_FecProceso 
         Caption         =   "Año Proceso        :"
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
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Frm_CargaArchivoAnno"
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
    
    Frm_CargaArchivoAnno.Left = 0
    Frm_CargaArchivoAnno.Top = 0
    
Exit Sub
Err_Carga:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub


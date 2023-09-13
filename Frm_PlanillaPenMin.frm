VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Frm_PlanillaPenMin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planilla Informes Garantía Estatal"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   7665
   Begin VB.Frame Fra_Datos 
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   7455
      Begin VB.ComboBox Cmb_TipoInforme 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   840
         Width           =   4695
      End
      Begin VB.TextBox Txt_Desde 
         Height          =   285
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Txt_Hasta 
         Height          =   285
         Left            =   4200
         MaxLength       =   10
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Lbl_Contrato 
         Caption         =   "Tipo de Informe              :"
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
         TabIndex        =   10
         Top             =   840
         Width           =   2415
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
         Left            =   3960
         TabIndex        =   9
         Top             =   360
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
         TabIndex        =   8
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Fra_Botones 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   7455
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   2400
         Picture         =   "Frm_PlanillaPenMin.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   4560
         Picture         =   "Frm_PlanillaPenMin.frx":06BA
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   3480
         Picture         =   "Frm_PlanillaPenMin.frx":07B4
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   200
         Width           =   720
      End
      Begin Crystal.CrystalReport Rpt_Calculo 
         Left            =   120
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
End
Attribute VB_Name = "Frm_PlanillaPenMin"
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

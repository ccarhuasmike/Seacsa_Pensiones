VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Frm_PorcCastigo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Porcentajes de Disminución de la Pensión por Quiebra"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   6720
   Begin VB.Frame Fra_Botones 
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   4680
      Width           =   6495
      Begin VB.CommandButton Cmd_Eliminar 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   1440
         Picture         =   "Frm_PorcCastigo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Eliminar Año"
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   2640
         Picture         =   "Frm_PorcCastigo.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5040
         Picture         =   "Frm_PorcCastigo.frx":09FC
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   3840
         Picture         =   "Frm_PorcCastigo.frx":0AF6
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Grabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   240
         Picture         =   "Frm_PorcCastigo.frx":11B0
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   200
         Width           =   720
      End
      Begin Crystal.CrystalReport Rpt_Calculo 
         Left            =   4680
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "  Parámetros Definidos  "
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
      Height          =   2055
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   6495
      Begin VB.TextBox txt_Titulo 
         Height          =   495
         Left            =   2640
         MaxLength       =   70
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   1440
         Width           =   3735
      End
      Begin VB.TextBox txt_Año 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3240
         MaxLength       =   4
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txt_TopeMax 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2640
         MaxLength       =   6
         TabIndex        =   3
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txt_PorCastigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2640
         MaxLength       =   6
         TabIndex        =   2
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txt_Mes 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   0
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lbl_nombre 
         Caption         =   "Titulo a Agregar en Reportes"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   19
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label lbl_nombre 
         Alignment       =   2  'Center
         Caption         =   "-"
         Height          =   270
         Index           =   3
         Left            =   3000
         TabIndex        =   17
         Top             =   360
         Width           =   195
      End
      Begin VB.Label lbl_nombre 
         Caption         =   "(Mes - Año)"
         Height          =   270
         Index           =   10
         Left            =   3960
         TabIndex        =   16
         Top             =   360
         Width           =   930
      End
      Begin VB.Label lbl_nombre 
         Caption         =   "%"
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   15
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lbl_nombre 
         Caption         =   "Tope Máximo Pensión  UF"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label lbl_nombre 
         Caption         =   "Porcentaje de Castigo"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label lbl_nombre 
         Caption         =   "Período de Inicio de Aplicación"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   2415
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Msf_Grilla 
      Height          =   2490
      Left            =   120
      TabIndex        =   18
      Top             =   2160
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   4392
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColor       =   14745599
      FormatString    =   $"Frm_PorcCastigo.frx":186A
   End
   Begin Crystal.CrystalReport Rpt_Reportes 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "Frm_PorcCastigo"
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

    Frm_PorcCastigo.Top = 0
    Frm_PorcCastigo.Left = 0

            
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

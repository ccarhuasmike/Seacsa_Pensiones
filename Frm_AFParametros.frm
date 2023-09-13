VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Frm_AFParametros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenedor de Valores de Cargas Familiares."
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8880
   Begin VB.Frame Fra_Vigencia 
      Caption         =   " "
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
      Height          =   855
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   6255
      Begin VB.TextBox Txt_InicioVig 
         Height          =   315
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   1
         Top             =   360
         Width           =   1005
      End
      Begin VB.CommandButton Cmd_Buscar 
         Height          =   375
         Left            =   4680
         Picture         =   "Frm_AFParametros.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Efectuar Busqueda de Vigencia"
         Top             =   340
         Width           =   855
      End
      Begin Crystal.CrystalReport Rpt_General 
         Left            =   5760
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowState     =   2
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   " -"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   13
         Left            =   2520
         TabIndex        =   24
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Lbl_TerminoVig 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3000
         TabIndex        =   23
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "  Ingrese Vigencia :"
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
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   0
         Width           =   1680
      End
   End
   Begin VB.Frame Fra_Tasas 
      Enabled         =   0   'False
      Height          =   1815
      Left            =   120
      TabIndex        =   16
      Top             =   960
      Width           =   6255
      Begin VB.CommandButton Cmd_Sumar 
         Height          =   450
         Left            =   5280
         Picture         =   "Frm_AFParametros.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Agregar Montos"
         Top             =   250
         Width           =   495
      End
      Begin VB.CommandButton Cmd_Restar 
         Height          =   450
         Left            =   5280
         Picture         =   "Frm_AFParametros.frx":028C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Quitar Montos"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox Txt_Desde 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2280
         MaxLength       =   11
         TabIndex        =   3
         Top             =   555
         Width           =   1095
      End
      Begin VB.TextBox Txt_Hasta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2280
         MaxLength       =   11
         TabIndex        =   4
         Top             =   870
         Width           =   1095
      End
      Begin VB.TextBox Txt_Valor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2280
         MaxLength       =   11
         TabIndex        =   5
         Top             =   1185
         Width           =   1095
      End
      Begin VB.CommandButton Cmd_LimpiarRango 
         Height          =   450
         Left            =   5280
         Picture         =   "Frm_AFParametros.frx":0416
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpiar"
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Hasta                          :"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   4
         Left            =   360
         TabIndex        =   20
         Top             =   870
         Width           =   1815
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Desde                         :"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   360
         TabIndex        =   19
         Top             =   555
         Width           =   1815
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Valor Carga                 :"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   360
         TabIndex        =   18
         Top             =   1185
         Width           =   1815
      End
      Begin VB.Label Lbl_NumOrden 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4320
         TabIndex        =   17
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Frame Fra_Periodos 
      Caption         =   "  Períodos  "
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
      Height          =   6255
      Left            =   6480
      TabIndex        =   15
      Top             =   0
      Width           =   2295
      Begin VB.ListBox Lst_Annos 
         BackColor       =   &H00E8FFFF&
         Height          =   5910
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Fra_Operacion 
      Height          =   1095
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   5160
      Width           =   6255
      Begin VB.CommandButton Cmd_Grabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   600
         Picture         =   "Frm_AFParametros.frx":0A88
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Grabar datos de Vigencia"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Eliminar 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   1680
         Picture         =   "Frm_AFParametros.frx":1142
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Eliminar Vigencia"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   3840
         Picture         =   "Frm_AFParametros.frx":1484
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   2760
         Picture         =   "Frm_AFParametros.frx":1B3E
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Imprimir Reporte Montos Cargas Familiares"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   4920
         Picture         =   "Frm_AFParametros.frx":21F8
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Msf_Grilla 
      Height          =   2295
      Left            =   120
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2880
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   4048
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      BackColor       =   -2147483624
      BackColorFixed  =   12632256
      BackColorBkg    =   -2147483632
      GridColor       =   0
      AllowBigSelection=   0   'False
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "Frm_AFParametros"
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

    Frm_AFParametros.Left = 0
    Frm_AFParametros.Top = 0
  
  
Exit Sub
Err_Carga:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub


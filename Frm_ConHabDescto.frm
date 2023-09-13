VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Frm_ConHabDescto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta histórica de Haberes y Descuentos"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   10215
   Begin VB.Frame Fra_Botones 
      Height          =   975
      Left            =   195
      TabIndex        =   18
      Top             =   6645
      Width           =   9855
      Begin VB.CommandButton Cmd_Grabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   1320
         Picture         =   "Frm_ConHabDescto.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   4560
         Picture         =   "Frm_ConHabDescto.frx":06BA
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   6720
         Picture         =   "Frm_ConHabDescto.frx":0D74
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   3480
         Picture         =   "Frm_ConHabDescto.frx":0E6E
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Eliminar 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   2400
         Picture         =   "Frm_ConHabDescto.frx":1528
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Eliminar Año"
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   5640
         Picture         =   "Frm_ConHabDescto.frx":186A
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   200
         Width           =   730
      End
      Begin Crystal.CrystalReport Rpt_Calculo 
         Left            =   7680
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generar Consulta"
      Height          =   390
      Left            =   4515
      TabIndex        =   17
      Top             =   2595
      Width           =   1710
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3300
      Left            =   4470
      TabIndex        =   16
      Top             =   3075
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   5821
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      FormatString    =   "Fecha de Pensión|Código|Concepto                                    |Monto             "
   End
   Begin VB.Frame Fra_Poliza 
      Caption         =   "Pensionado"
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
      Height          =   1590
      Left            =   4470
      TabIndex        =   5
      Top             =   810
      Width           =   5655
      Begin VB.TextBox Txt_PenPoliza 
         Height          =   285
         Left            =   1470
         TabIndex        =   10
         Top             =   360
         Width           =   1185
      End
      Begin VB.TextBox Txt_PenDigito 
         Height          =   285
         Left            =   2865
         TabIndex        =   9
         Top             =   690
         Width           =   375
      End
      Begin VB.TextBox Txt_PenRut 
         Height          =   285
         Left            =   1470
         TabIndex        =   8
         Top             =   705
         Width           =   1170
      End
      Begin VB.CommandButton Cmd_BuscarPol 
         Height          =   375
         Left            =   4905
         Picture         =   "Frm_ConHabDescto.frx":1E44
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Buscar Póliza"
         Top             =   225
         Width           =   615
      End
      Begin VB.CommandButton Cmd_Buscar 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         Picture         =   "Frm_ConHabDescto.frx":1F46
         TabIndex        =   6
         ToolTipText     =   "Buscar"
         Top             =   630
         Width           =   615
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Póliza"
         Height          =   255
         Index           =   0
         Left            =   195
         TabIndex        =   15
         Top             =   390
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   2
         Left            =   165
         TabIndex        =   14
         Top             =   1050
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Rut Pensionado"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   13
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Lbl_Nombre 
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
         Index           =   20
         Left            =   2655
         TabIndex        =   12
         Top             =   735
         Width           =   255
      End
      Begin VB.Label Lbl_PenNombre 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   1470
         TabIndex        =   11
         Top             =   1035
         Width           =   3405
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   600
      Left            =   4440
      TabIndex        =   2
      Top             =   120
      Width           =   5655
      Begin VB.OptionButton Option2 
         Caption         =   "Por Pensionado"
         Height          =   330
         Left            =   150
         TabIndex        =   4
         Top             =   255
         Width           =   1620
      End
      Begin VB.OptionButton Option1 
         Caption         =   "General"
         Height          =   195
         Left            =   2775
         TabIndex        =   3
         Top             =   300
         Width           =   1020
      End
   End
   Begin VB.Frame Framselecciòn 
      Caption         =   "Selección de Concepto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   6315
      Left            =   180
      TabIndex        =   0
      Top             =   135
      Width           =   4125
      Begin VB.ListBox List1 
         Height          =   5685
         Left            =   195
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   390
         Width           =   3645
      End
   End
End
Attribute VB_Name = "Frm_ConHabDescto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Buscar_Click()
On Error GoTo Err_CmdBuscarClick

    Frm_Busqueda.flInicio ("Frm_ConHabDescto")
    
Exit Sub
Err_CmdBuscarClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Salir_Click()
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

    Frm_ConHabDescto.Top = 0
    Frm_ConHabDescto.Left = 0
            
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub


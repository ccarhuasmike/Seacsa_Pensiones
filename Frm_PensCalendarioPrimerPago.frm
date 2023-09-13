VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Frm_PensCalendarioPrimerPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantención de Calendario de Pagos - Primeros Pagos"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   6240
   Begin VB.Frame Fra_Operacion 
      Height          =   1095
      Left            =   120
      TabIndex        =   12
      Top             =   5760
      Width           =   6015
      Begin VB.CommandButton Cmd_Grabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   465
         Picture         =   "Frm_PensCalendarioPrimerPago.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Grabar Datos"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   4800
         Picture         =   "Frm_PensCalendarioPrimerPago.frx":06BA
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   3720
         Picture         =   "Frm_PensCalendarioPrimerPago.frx":07B4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Eliminar 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   1560
         Picture         =   "Frm_PensCalendarioPrimerPago.frx":0E6E
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   2640
         Picture         =   "Frm_PensCalendarioPrimerPago.frx":11B0
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Fra_Gral 
      Caption         =   "  Especificación de Períodos de Pago  "
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
      Height          =   2775
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   6045
      Begin VB.Frame Fra_Datos 
         Height          =   1455
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   5805
         Begin VB.TextBox Txt_ProxPago 
            Height          =   285
            Left            =   3150
            MaxLength       =   10
            TabIndex        =   6
            Top             =   975
            Width           =   1155
         End
         Begin Crystal.CrystalReport Rpt_General 
            Left            =   5280
            Top             =   600
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowState     =   2
            PrintFileLinesPerPage=   60
         End
         Begin VB.TextBox Txt_CalPagoReg 
            Height          =   285
            Left            =   3150
            MaxLength       =   10
            TabIndex        =   5
            Top             =   735
            Width           =   1155
         End
         Begin VB.TextBox Txt_PagoReg 
            Height          =   285
            Left            =   3150
            MaxLength       =   10
            TabIndex        =   4
            Top             =   495
            Width           =   1155
         End
         Begin VB.Label lbl_nombre 
            Caption         =   "Fecha Pago"
            Height          =   255
            Index           =   1
            Left            =   630
            TabIndex        =   21
            Top             =   495
            Width           =   1875
         End
         Begin VB.Label lbl_nombre 
            Caption         =   "Fecha Cálculo"
            Height          =   255
            Index           =   2
            Left            =   630
            TabIndex        =   20
            Top             =   735
            Width           =   2355
         End
         Begin VB.Label lbl_nombre 
            Caption         =   "Fecha Próximo Pago"
            Height          =   255
            Index           =   4
            Left            =   630
            TabIndex        =   19
            Top             =   975
            Width           =   1875
         End
         Begin VB.Label lbl_nombre 
            Caption         =   "Definición Primeros Pagos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   240
            Index           =   6
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   4155
         End
      End
      Begin VB.Frame Fra_Busqueda 
         Height          =   900
         Left            =   120
         TabIndex        =   14
         Top             =   255
         Width           =   5790
         Begin VB.CommandButton Cmd_Buscar 
            Height          =   375
            Left            =   4800
            Picture         =   "Frm_PensCalendarioPrimerPago.frx":186A
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Efectuar Busqueda de IPC"
            Top             =   320
            Width           =   855
         End
         Begin VB.TextBox Txt_Mes 
            Height          =   285
            Left            =   1920
            MaxLength       =   10
            TabIndex        =   1
            Top             =   360
            Width           =   1035
         End
         Begin VB.TextBox Txt_Anno 
            Height          =   285
            Left            =   3600
            MaxLength       =   10
            TabIndex        =   2
            Top             =   360
            Width           =   1035
         End
         Begin VB.Label lbl_nombre 
            Alignment       =   2  'Center
            Caption         =   "Hasta"
            Height          =   255
            Index           =   11
            Left            =   3000
            TabIndex        =   22
            Top             =   360
            Width           =   555
         End
         Begin VB.Label lbl_nombre 
            Caption         =   "Polizas recibidas Desde"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   1755
         End
         Begin VB.Label lbl_nombre 
            Caption         =   " Definición Período de Pago"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   7
            Left            =   135
            TabIndex        =   15
            Top             =   15
            Width           =   2475
         End
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Msf_Grilla 
      Height          =   2775
      Left            =   90
      TabIndex        =   13
      Top             =   3000
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   4895
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   14745599
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Frm_PensCalendarioPrimerPago"
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

    Frm_PensCalendarioPrimerPago.Top = 0
    Frm_PensCalendarioPrimerPago.Left = 0
    
   
    Fra_Busqueda.Enabled = True
 
    Fra_Datos.Enabled = False

            
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Frm_PensParametrosPrimerPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros del Cálculo de Primeros Pagos"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   6885
   Begin VB.Frame Frame1 
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
      Height          =   3135
      Left            =   80
      TabIndex        =   16
      Top             =   80
      Width           =   6765
      Begin VB.Frame Fra_PagoReg 
         Height          =   1815
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   6525
         Begin VB.TextBox Txt_PagoReg 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3150
            MaxLength       =   10
            TabIndex        =   3
            Top             =   480
            Width           =   1155
         End
         Begin VB.TextBox Txt_CalPagoReg 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3150
            MaxLength       =   10
            TabIndex        =   4
            Top             =   720
            Width           =   1155
         End
         Begin VB.TextBox Txt_ProxPago 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3150
            MaxLength       =   10
            TabIndex        =   5
            Top             =   960
            Width           =   1155
         End
         Begin VB.OptionButton Opt_AbiertoReg 
            Caption         =   "Abierto"
            Height          =   375
            Left            =   2280
            TabIndex        =   7
            Top             =   1305
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton Opt_ProvisorioReg 
            Caption         =   "Provisorio"
            Height          =   375
            Left            =   3360
            TabIndex        =   8
            Top             =   1305
            Width           =   1215
         End
         Begin VB.OptionButton Opt_CerradoReg 
            Caption         =   "Cerrado"
            Height          =   375
            Left            =   4560
            TabIndex        =   9
            Top             =   1305
            Width           =   975
         End
         Begin VB.Label lbl_nombre 
            Caption         =   "Fecha Pago"
            Height          =   255
            Index           =   1
            Left            =   630
            TabIndex        =   24
            Top             =   480
            Width           =   1875
         End
         Begin VB.Label lbl_nombre 
            Caption         =   "Fecha Cálculo"
            Height          =   255
            Index           =   2
            Left            =   630
            TabIndex        =   23
            Top             =   720
            Width           =   2355
         End
         Begin VB.Label lbl_nombre 
            Caption         =   "Fecha Próximo Pago"
            Height          =   255
            Index           =   4
            Left            =   630
            TabIndex        =   22
            Top             =   960
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
            Height          =   255
            Index           =   5
            Left            =   135
            TabIndex        =   21
            Top             =   240
            Width           =   3915
         End
         Begin VB.Label lbl_nombre 
            Caption         =   "Estado del Período"
            Height          =   255
            Index           =   12
            Left            =   630
            TabIndex        =   20
            Top             =   1305
            Width           =   1695
         End
      End
      Begin VB.Frame Fra_Busqueda 
         Height          =   900
         Left            =   120
         TabIndex        =   17
         Top             =   255
         Width           =   6525
         Begin VB.TextBox txt_Desde 
            Height          =   285
            Left            =   1920
            MaxLength       =   10
            TabIndex        =   0
            Top             =   360
            Width           =   1155
         End
         Begin VB.TextBox txt_Hasta 
            Height          =   285
            Left            =   3960
            MaxLength       =   10
            TabIndex        =   1
            Top             =   360
            Width           =   1155
         End
         Begin VB.CommandButton Cmd_Buscar 
            Height          =   375
            Left            =   5520
            Picture         =   "Frm_PensParametrosPrimerPago.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Efectuar Busqueda de IPC"
            Top             =   255
            Width           =   855
         End
         Begin VB.Label lbl_nombre 
            Caption         =   "Polizas recibidas Desde"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   1875
         End
         Begin VB.Label Label1 
            Caption         =   "Hasta"
            Height          =   255
            Left            =   3240
            TabIndex        =   25
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lbl_nombre 
            Caption         =   " Definición Periodo de Pago"
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
            Index           =   11
            Left            =   135
            TabIndex        =   18
            Top             =   15
            Width           =   2475
         End
      End
   End
   Begin VB.Frame Fra_Operacion 
      Height          =   1095
      Left            =   80
      TabIndex        =   6
      Top             =   5535
      Width           =   6735
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   3000
         Picture         =   "Frm_PensParametrosPrimerPago.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Eliminar 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   1680
         Picture         =   "Frm_PensParametrosPrimerPago.frx":07BC
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   4320
         Picture         =   "Frm_PensParametrosPrimerPago.frx":0AFE
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5640
         Picture         =   "Frm_PensParametrosPrimerPago.frx":11B8
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Grabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   400
         Picture         =   "Frm_PensParametrosPrimerPago.frx":12B2
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Grabar Datos"
         Top             =   240
         Width           =   720
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Msf_Grilla 
      Height          =   2295
      Left            =   75
      TabIndex        =   15
      Top             =   3285
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   4048
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
Attribute VB_Name = "Frm_PensParametrosPrimerPago"
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

    Frm_PensParametrosPrimerPago.Top = 0
    Frm_PensParametrosPrimerPago.Left = 0
    
    Fra_Busqueda.Enabled = True

    Opt_AbiertoReg.Value = True

    Fra_PagoReg.Enabled = False
    Call flInicializaGrilla
            
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
    
End Sub



Function flInicializaGrilla()
'Permite limpiar e inicializar la grilla.
'----------------------------------------------------------------------

On Error GoTo Err_flInicializaGrilla

    Msf_Grilla.Clear
    Msf_Grilla.Cols = 6
    Msf_Grilla.Rows = 1
    Msf_Grilla.RowHeight(0) = 250
    Msf_Grilla.Row = 0
        
    Msf_Grilla.Col = 0
    Msf_Grilla.Text = "Fecha Recep. Desde"
    Msf_Grilla.ColWidth(0) = 1700
    Msf_Grilla.ColAlignment(0) = 1  'centrado
    
    Msf_Grilla.Col = 1
    Msf_Grilla.Text = "Fecha Recep. Hasta"
    Msf_Grilla.ColWidth(1) = 1700
    Msf_Grilla.ColAlignment(1) = 1  'centrado
        
    Msf_Grilla.Col = 2
    Msf_Grilla.Text = "Fecha Pago 1º Pago"
    Msf_Grilla.ColWidth(2) = 1700
    
    Msf_Grilla.Col = 3
    Msf_Grilla.Text = "Cálculo Pago 1º Pago"
    Msf_Grilla.ColWidth(3) = 1700
    
    Msf_Grilla.Col = 4
    Msf_Grilla.Text = "Proximo Pago"
    Msf_Grilla.ColWidth(4) = 1500
    
    Msf_Grilla.Col = 5
    Msf_Grilla.Text = "Estado 1º P."
    Msf_Grilla.ColWidth(5) = 1000
    
Exit Function
Err_flInicializaGrilla:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
    

End Function



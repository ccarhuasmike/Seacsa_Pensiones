VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Frm_GEPensionesMinimas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenedor de Pensiones Mínimas."
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   10935
   Begin VB.Frame Fra_Operacion 
      Height          =   1095
      Index           =   0
      Left            =   120
      TabIndex        =   69
      Top             =   5760
      Width           =   8295
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   4800
         Picture         =   "Frm_GEPensionesMinimas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5880
         Picture         =   "Frm_GEPensionesMinimas.frx":05DA
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   3720
         Picture         =   "Frm_GEPensionesMinimas.frx":06D4
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Imprimir Reporte"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Eliminar 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   2640
         Picture         =   "Frm_GEPensionesMinimas.frx":0D8E
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Eliminar Año"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Grabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   1560
         Picture         =   "Frm_GEPensionesMinimas.frx":10D0
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Grabar datos de IPC"
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Fra_Vigencia 
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
      Height          =   975
      Left            =   120
      TabIndex        =   65
      Top             =   120
      Width           =   8295
      Begin VB.CommandButton Cmd_Buscar 
         Height          =   375
         Left            =   5160
         Picture         =   "Frm_GEPensionesMinimas.frx":178A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Efectuar Busqueda de Vigencia"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Txt_InicioVig 
         Height          =   315
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
      Begin Crystal.CrystalReport Rpt_General 
         Left            =   6480
         Top             =   360
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
         Left            =   3240
         TabIndex        =   68
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Lbl_TerminoVig 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   315
         Left            =   3600
         TabIndex        =   67
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "  Ingrese Vigencia :  "
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
         Index           =   0
         Left            =   240
         TabIndex        =   66
         Top             =   0
         Width           =   1800
      End
   End
   Begin VB.Frame Fra_Detalle 
      Height          =   4695
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   8295
      Begin VB.TextBox Txt_PenFin1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6840
         MaxLength       =   14
         TabIndex        =   8
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Txt_PenMin1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4080
         MaxLength       =   14
         TabIndex        =   7
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton Cmd_LimpiarRango 
         Height          =   450
         Left            =   5655
         Picture         =   "Frm_GEPensionesMinimas.frx":188C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Limpiar"
         Top             =   525
         Width           =   495
      End
      Begin VB.CommandButton Cmd_Calcular 
         Height          =   450
         Left            =   5160
         Picture         =   "Frm_GEPensionesMinimas.frx":1EFE
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Calcular el Monto de las Pensiones."
         Top             =   525
         Width           =   495
      End
      Begin VB.Frame Fra_Rango 
         Caption         =   "Rangos de Edad"
         ForeColor       =   &H00800000&
         Height          =   1215
         Left            =   6720
         TabIndex        =   36
         Top             =   120
         Width           =   1455
         Begin VB.ListBox Lst_Edad 
            BackColor       =   &H00E8FFFF&
            Height          =   840
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.TextBox Txt_Desde 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2160
         MaxLength       =   3
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Txt_Hasta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3720
         MaxLength       =   3
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Txt_PenMin2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4080
         MaxLength       =   14
         TabIndex        =   35
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox Txt_PenMin3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4080
         MaxLength       =   14
         TabIndex        =   34
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox Txt_PenMin4 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4080
         MaxLength       =   14
         TabIndex        =   33
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox Txt_PenMin5 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4080
         MaxLength       =   14
         TabIndex        =   32
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox Txt_PenMin6 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4080
         MaxLength       =   14
         TabIndex        =   31
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox Txt_PenMin7 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4080
         MaxLength       =   14
         TabIndex        =   30
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox Txt_PenMin8 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4080
         MaxLength       =   14
         TabIndex        =   29
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox Txt_PenMin9 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4080
         MaxLength       =   14
         TabIndex        =   28
         Top             =   3600
         Width           =   1335
      End
      Begin VB.TextBox Txt_PenMin10 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4080
         MaxLength       =   14
         TabIndex        =   27
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox Txt_PenMin11 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4080
         MaxLength       =   14
         TabIndex        =   26
         Top             =   4080
         Width           =   1335
      End
      Begin VB.TextBox Txt_PenFin2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6840
         MaxLength       =   14
         TabIndex        =   25
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox Txt_PenFin3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6840
         MaxLength       =   14
         TabIndex        =   24
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox Txt_PenFin4 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6840
         MaxLength       =   14
         TabIndex        =   23
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox Txt_PenFin5 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6840
         MaxLength       =   14
         TabIndex        =   22
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox Txt_PenFin6 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6840
         MaxLength       =   14
         TabIndex        =   21
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox Txt_PenFin7 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6840
         MaxLength       =   14
         TabIndex        =   20
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox Txt_PenFin8 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6840
         MaxLength       =   14
         TabIndex        =   19
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox Txt_PenFin9 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6840
         MaxLength       =   14
         TabIndex        =   18
         Top             =   3600
         Width           =   1335
      End
      Begin VB.TextBox Txt_PenFin10 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6840
         MaxLength       =   14
         TabIndex        =   17
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox Txt_PenFin11 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6840
         MaxLength       =   14
         TabIndex        =   16
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Lbl_Bon1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5400
         TabIndex        =   57
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Lbl_Nombre 
         Alignment       =   2  'Center
         Caption         =   "Pensión Final"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   7
         Left            =   6840
         TabIndex        =   64
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Lbl_Nombre 
         Alignment       =   2  'Center
         Caption         =   "Bonificación"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   6
         Left            =   5520
         TabIndex        =   63
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Lbl_Nombre 
         Alignment       =   2  'Center
         Caption         =   "Pensión Mínima"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   5
         Left            =   4200
         TabIndex        =   62
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Lbl_Item 
         AutoSize        =   -1  'True
         Caption         =   "Vejez o Invalidez"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   61
         Top             =   1800
         Width           =   4080
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Beneficiario/Beneficio"
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   60
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Rangos de Edad :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   59
         Top             =   600
         Width           =   1455
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
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   4
         Left            =   3360
         TabIndex        =   58
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Lbl_Bon2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5400
         TabIndex        =   56
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Lbl_Bon3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5400
         TabIndex        =   55
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Lbl_Bon4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5400
         TabIndex        =   54
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Lbl_Bon5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5400
         TabIndex        =   53
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Lbl_Bon6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5400
         TabIndex        =   52
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Lbl_Bon7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5400
         TabIndex        =   51
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Lbl_Bon8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5400
         TabIndex        =   50
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Lbl_Bon9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5400
         TabIndex        =   49
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label Lbl_Bon10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5400
         TabIndex        =   48
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label Lbl_Bon11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5400
         TabIndex        =   47
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Lbl_Item 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Viuda(o) Inv. Total, s/hij c/Derecho a Pensión"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   46
         Top             =   2040
         Width           =   4095
      End
      Begin VB.Label Lbl_Item 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Viudo Inv. Parcial, s/hij c/Derecho a Pensión"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   45
         Top             =   2280
         Width           =   4170
      End
      Begin VB.Label Lbl_Item 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Viuda(o) Inv. Total, c/hij c/Derecho a Pensión"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   44
         Top             =   2520
         Width           =   4125
      End
      Begin VB.Label Lbl_Item 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Viudo Inv. Parcial, c/hij c/Derecho a Pensión"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   43
         Top             =   2760
         Width           =   4065
      End
      Begin VB.Label Lbl_Item 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Madre Hijo Natural s/hij c/Derecho a Pensión"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   42
         Top             =   3000
         Width           =   4080
      End
      Begin VB.Label Lbl_Item 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Madre Hijo Natural c/hij c/Derecho a Pensión"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   41
         Top             =   3240
         Width           =   4095
      End
      Begin VB.Label Lbl_Item 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hijos Menores de 18 años de Edad"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   40
         Top             =   3480
         Width           =   2490
      End
      Begin VB.Label Lbl_Item 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hijo Inv. Total Mayor de 18 años de Edad"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   39
         Top             =   3720
         Width           =   2955
      End
      Begin VB.Label Lbl_Item 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hijo Inv. Parcial Mayor de 18 años de Edad"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   38
         Top             =   3960
         Width           =   3075
      End
      Begin VB.Label Lbl_Item 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Madre o Padre"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   37
         Top             =   4200
         Width           =   4050
      End
   End
   Begin VB.Frame Fra_Periodo 
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
      Height          =   6735
      Left            =   8520
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.ListBox Lst_Annos 
         BackColor       =   &H00E8FFFF&
         Height          =   6300
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   2175
      End
   End
End
Attribute VB_Name = "Frm_GEPensionesMinimas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_salir_Click()
On Error GoTo Err_Salir

    Unload Me

Exit Sub
Err_Salir:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(2) & Err.Description & " ]"
    End Select
End Sub

Private Sub Form_Load()
On Error GoTo Err_Cargar

    Frm_GEPensionesMinimas.Left = 0
    Frm_GEPensionesMinimas.Top = 0
    Lst_Edad.Enabled = False
    Fra_Detalle.Enabled = False
    vlClicEnLista = False
   
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(2) & Err.Description & " ]"
    End Select
End Sub






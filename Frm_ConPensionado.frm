VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Frm_ConPensionado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta por Pensionado"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   9210
   Begin VB.Frame Fra_Botones 
      Height          =   975
      Left            =   120
      TabIndex        =   46
      Top             =   6720
      Width           =   8970
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5175
         Picture         =   "Frm_ConPensionado.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   165
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   3915
         Picture         =   "Frm_ConPensionado.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   195
         Width           =   720
      End
      Begin Crystal.CrystalReport Rpt_Calculo 
         Left            =   8415
         Top             =   345
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalle de Selección"
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
      Height          =   3735
      Left            =   3555
      TabIndex        =   44
      Top             =   2880
      Width           =   5610
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3360
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   5130
         _ExtentX        =   9049
         _ExtentY        =   5927
         _Version        =   393216
         Rows            =   18
         FormatString    =   $"Frm_ConPensionado.frx":07B4
      End
   End
   Begin VB.Frame Framselecciòn 
      Caption         =   "Selección de Consulta"
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
      Height          =   3780
      Left            =   90
      TabIndex        =   5
      Top             =   2880
      Width           =   3390
      Begin VB.OptionButton Option11 
         Caption         =   "Antecedentes Plan de Salud"
         Height          =   300
         Left            =   120
         TabIndex        =   40
         Top             =   1065
         Width           =   3105
      End
      Begin VB.OptionButton Option10 
         Caption         =   "Retención Judicial"
         Height          =   300
         Left            =   120
         TabIndex        =   39
         Top             =   2325
         Width           =   2085
      End
      Begin VB.OptionButton Option9 
         Caption         =   "Liquidación de Pensión"
         Height          =   300
         Left            =   120
         TabIndex        =   38
         Top             =   2580
         Width           =   2085
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Certificado de Estudio y Solteria"
         Height          =   300
         Left            =   120
         TabIndex        =   37
         Top             =   2850
         Width           =   2850
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Tutores"
         Height          =   300
         Left            =   120
         TabIndex        =   36
         Top             =   2070
         Width           =   2085
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Antecedentes para el Pago de Pensión"
         Height          =   300
         Left            =   120
         TabIndex        =   11
         Top             =   780
         Width           =   3180
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Cajas de Compensación "
         Height          =   300
         Left            =   120
         TabIndex        =   10
         Top             =   1815
         Width           =   2085
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Beneficiarios"
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   540
         Width           =   1320
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Garantía Estatal"
         Height          =   300
         Left            =   120
         TabIndex        =   8
         Top             =   1305
         Width           =   2085
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Asignación Familiar"
         Height          =   300
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   2115
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Póliza"
         Height          =   300
         Left            =   120
         TabIndex        =   6
         Top             =   270
         Width           =   930
      End
   End
   Begin VB.Frame Fra_Poliza 
      Caption         =   "Póliza / Pensionado"
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
      Height          =   2715
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00E0FFFF&
         Height          =   315
         Left            =   1020
         TabIndex        =   20
         Top             =   1920
         Width           =   2310
      End
      Begin VB.TextBox Text15 
         Height          =   300
         Left            =   1020
         TabIndex        =   19
         Top             =   1590
         Width           =   6780
      End
      Begin VB.TextBox Text14 
         BackColor       =   &H00E0FFFF&
         Height          =   285
         Left            =   1020
         TabIndex        =   18
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox Text13 
         BackColor       =   &H00E0FFFF&
         Height          =   285
         Left            =   3660
         TabIndex        =   17
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   1020
         TabIndex        =   16
         Top             =   930
         Width           =   3015
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1020
         TabIndex        =   15
         Top             =   2250
         Width           =   2025
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   5250
         TabIndex        =   14
         Top             =   1290
         Width           =   2550
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1020
         TabIndex        =   13
         Top             =   1260
         Width           =   3015
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0FFFF&
         Height          =   285
         Left            =   5250
         TabIndex        =   12
         Top             =   960
         Width           =   2550
      End
      Begin VB.TextBox Txt_PenPoliza 
         Height          =   285
         Left            =   1020
         TabIndex        =   3
         Top             =   270
         Width           =   1185
      End
      Begin VB.CommandButton Cmd_BuscarPol 
         Height          =   375
         Left            =   8265
         Picture         =   "Frm_ConPensionado.frx":092F
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Buscar Póliza"
         Top             =   1095
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
         Left            =   8265
         Picture         =   "Frm_ConPensionado.frx":0A31
         TabIndex        =   1
         ToolTipText     =   "Buscar"
         Top             =   1455
         Width           =   615
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Comuna"
         Height          =   255
         Index           =   22
         Left            =   60
         TabIndex        =   35
         Top             =   1920
         Width           =   825
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Domicilio"
         Height          =   255
         Index           =   21
         Left            =   60
         TabIndex        =   34
         Top             =   1590
         Width           =   810
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Rut"
         Height          =   255
         Index           =   20
         Left            =   60
         TabIndex        =   33
         Top             =   660
         Width           =   615
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
         Index           =   19
         Left            =   3300
         TabIndex        =   32
         Top             =   600
         Width           =   135
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   18
         Left            =   60
         TabIndex        =   31
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   3390
         TabIndex        =   30
         Top             =   1920
         Width           =   1995
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5445
         TabIndex        =   29
         Top             =   1935
         Width           =   2355
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nº Orden"
         Height          =   255
         Index           =   17
         Left            =   7275
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   8085
         TabIndex        =   27
         Top             =   195
         Width           =   615
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Ap. Paterno"
         Height          =   255
         Index           =   16
         Left            =   60
         TabIndex        =   26
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Ap. Materno"
         Height          =   255
         Index           =   15
         Left            =   4170
         TabIndex        =   25
         Top             =   1290
         Width           =   1095
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Teléfono"
         Height          =   255
         Index           =   11
         Left            =   75
         TabIndex        =   24
         Top             =   2280
         Width           =   795
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5250
         TabIndex        =   23
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Fecha Nac."
         Height          =   255
         Index           =   2
         Left            =   4170
         TabIndex        =   22
         Top             =   630
         Width           =   975
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Sit. Inv."
         Height          =   255
         Index           =   1
         Left            =   4170
         TabIndex        =   21
         Top             =   960
         Width           =   825
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Póliza"
         Height          =   255
         Index           =   0
         Left            =   75
         TabIndex        =   4
         Top             =   315
         Width           =   855
      End
   End
   Begin VB.Label Lbl_Nombre 
      Caption         =   "Tasa de Venta"
      Height          =   255
      Index           =   36
      Left            =   990
      TabIndex        =   43
      Top             =   8610
      Width           =   2295
   End
   Begin VB.Label Lbl_Nombre 
      Caption         =   "Tasa Cto. Reaseguro"
      Height          =   255
      Index           =   37
      Left            =   990
      TabIndex        =   42
      Top             =   8910
      Width           =   1815
   End
   Begin VB.Label Lbl_Nombre 
      Caption         =   "Tasa de Int. Período Garantizado"
      Height          =   255
      Index           =   38
      Left            =   990
      TabIndex        =   41
      Top             =   9225
      Width           =   2535
   End
End
Attribute VB_Name = "Frm_ConPensionado"
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

Private Sub Option1_Click()
    Frm_AntPensionado.Show
    
End Sub

VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Frm_AFMantNoBenef 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenedor de Carga Familiar No Beneficiaria de Pensión."
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   9120
   Begin VB.Frame Fra_Antecedente 
      Caption         =   "Antecedenes Personales"
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
      Height          =   3420
      Left            =   120
      TabIndex        =   34
      Top             =   1200
      Width           =   8895
      Begin VB.TextBox Txt_FecNac 
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   15
         Top             =   2280
         Width           =   1425
      End
      Begin VB.TextBox Txt_FecMatri 
         Height          =   285
         Left            =   6000
         MaxLength       =   10
         TabIndex        =   20
         Top             =   3000
         Width           =   1425
      End
      Begin VB.TextBox Txt_FecFall 
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   19
         Top             =   3000
         Width           =   1515
      End
      Begin VB.ComboBox Cmb_SitInv 
         BackColor       =   &H00E0FFFF&
         Height          =   315
         Left            =   6000
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2640
         Width           =   2595
      End
      Begin VB.ComboBox Cmb_Sexo 
         BackColor       =   &H00E0FFFF&
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2640
         Width           =   2595
      End
      Begin VB.ComboBox Cmb_CodAscDesc 
         BackColor       =   &H00E0FFFF&
         Height          =   315
         Left            =   6000
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   2280
         Width           =   2595
      End
      Begin VB.ComboBox Cmb_Comuna 
         BackColor       =   &H00E0FFFF&
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1560
         Width           =   2595
      End
      Begin VB.TextBox Txt_Domicilio 
         Height          =   300
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   11
         Top             =   1230
         Width           =   7575
      End
      Begin VB.TextBox Txt_Rut 
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   6
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox Txt_Digito 
         Height          =   285
         Left            =   3840
         MaxLength       =   1
         TabIndex        =   7
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox Txt_Nombre 
         Height          =   285
         Left            =   1080
         MaxLength       =   25
         TabIndex        =   8
         Top             =   570
         Width           =   3015
      End
      Begin VB.TextBox Txt_Email 
         Height          =   285
         Left            =   3840
         MaxLength       =   40
         TabIndex        =   14
         Top             =   1920
         Width           =   4785
      End
      Begin VB.TextBox Txt_Telefono 
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   13
         Top             =   1920
         Width           =   1425
      End
      Begin VB.TextBox Txt_ApMaterno 
         Height          =   285
         Left            =   5640
         MaxLength       =   25
         TabIndex        =   10
         Top             =   900
         Width           =   3015
      End
      Begin VB.TextBox Txt_ApPaterno 
         Height          =   285
         Left            =   1080
         MaxLength       =   25
         TabIndex        =   9
         Top             =   900
         Width           =   3015
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Antecedentes Personales"
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
         Height          =   255
         Index           =   22
         Left            =   120
         TabIndex        =   57
         Top             =   0
         Width           =   2175
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Fecha Nac."
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   53
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Fecha Matrimonio"
         Height          =   255
         Index           =   17
         Left            =   4440
         TabIndex        =   52
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Fecha  Fallecimiento"
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   51
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Sit. Invalidez"
         Height          =   255
         Index           =   15
         Left            =   4440
         TabIndex        =   50
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Sexo"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   49
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Código Asc./Desc."
         Height          =   255
         Index           =   11
         Left            =   4440
         TabIndex        =   48
         Top             =   2280
         Width           =   1395
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Comuna"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   47
         Top             =   1560
         Width           =   825
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Domicilio"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   46
         Top             =   1230
         Width           =   810
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Rut"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   45
         Top             =   240
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
         Index           =   4
         Left            =   3480
         TabIndex        =   44
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   43
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Lbl_Provincia 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   3840
         TabIndex        =   42
         Top             =   1560
         Width           =   2325
      End
      Begin VB.Label Lbl_Region 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   6240
         TabIndex        =   41
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nº Orden"
         Height          =   255
         Index           =   5
         Left            =   4560
         TabIndex        =   40
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Lbl_NumOrden 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5640
         TabIndex        =   39
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Ap. Paterno"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   38
         Top             =   900
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Ap. Materno"
         Height          =   255
         Index           =   8
         Left            =   4560
         TabIndex        =   37
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Email"
         Height          =   255
         Index           =   13
         Left            =   3240
         TabIndex        =   36
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Teléfono"
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   35
         Top             =   1920
         Width           =   795
      End
   End
   Begin VB.Frame Fra_Botones 
      Height          =   1095
      Left            =   120
      TabIndex        =   32
      Top             =   6120
      Width           =   8925
      Begin VB.CommandButton Cmd_Grabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   1320
         Picture         =   "Frm_AFMantNoBenef.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   4560
         Picture         =   "Frm_AFMantNoBenef.frx":06BA
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   6720
         Picture         =   "Frm_AFMantNoBenef.frx":0D74
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   3480
         Picture         =   "Frm_AFMantNoBenef.frx":0E6E
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Eliminar 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   2400
         Picture         =   "Frm_AFMantNoBenef.frx":1528
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Eliminar Año"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   5640
         Picture         =   "Frm_AFMantNoBenef.frx":186A
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Width           =   730
      End
      Begin Crystal.CrystalReport Rpt_NoBenef 
         Left            =   7320
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
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
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8895
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
         Left            =   8160
         Picture         =   "Frm_AFMantNoBenef.frx":1E44
         TabIndex        =   5
         ToolTipText     =   "Buscar"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Txt_PenPoliza 
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   1
         Top             =   360
         Width           =   1185
      End
      Begin VB.TextBox Txt_PenDigito 
         Height          =   285
         Left            =   6000
         MaxLength       =   1
         TabIndex        =   3
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox Txt_PenRut 
         Height          =   285
         Left            =   3960
         MaxLength       =   10
         TabIndex        =   2
         Top             =   360
         Width           =   1755
      End
      Begin VB.CommandButton Cmd_BuscarPol 
         Height          =   375
         Left            =   8160
         Picture         =   "Frm_AFMantNoBenef.frx":1F46
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Buscar Póliza"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Lbl_Nombre 
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   56
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label Lbl_Endoso 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7440
         TabIndex        =   55
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nº Endoso"
         Height          =   255
         Index           =   19
         Left            =   6480
         TabIndex        =   54
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Lbl_PenNombre 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   27
         Top             =   720
         Width           =   6855
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Póliza"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   30
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Rut Pensionado"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   29
         Top             =   360
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
         Left            =   5760
         TabIndex        =   28
         Top             =   360
         Width           =   255
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Msf_Grilla 
      Height          =   1410
      Left            =   120
      TabIndex        =   33
      Top             =   4680
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   2487
      _Version        =   393216
      Rows            =   1
      Cols            =   17
      BackColor       =   14745599
      FormatString    =   $"Frm_AFMantNoBenef.frx":2048
   End
End
Attribute VB_Name = "Frm_AFMantNoBenef"
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

    Frm_AFMantNoBenef.Top = 0
    Frm_AFMantNoBenef.Left = 0
    
   
    Fra_Antecedente.Enabled = False
    
   
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub


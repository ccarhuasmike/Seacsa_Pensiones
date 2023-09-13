VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Frm_AntCertificado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenedor de Certificados de Estudios y Soltería."
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   8910
   Begin VB.Frame Fra_Ben 
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
      Height          =   735
      Left            =   120
      TabIndex        =   42
      Top             =   2550
      Width           =   8655
      Begin VB.Label Lbl_DgvBen 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2880
         TabIndex        =   49
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Lbl_RutBen 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   48
         Top             =   240
         Width           =   1215
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
         Left            =   2640
         TabIndex        =   47
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Rut Beneficiario"
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   17
         Left            =   3360
         TabIndex        =   45
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Lbl_NomBen 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4080
         TabIndex        =   44
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   " Beneficiario  "
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
         Index           =   14
         Left            =   120
         TabIndex        =   43
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Frame Fra_Beneficiarios 
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
      Height          =   1335
      Left            =   120
      TabIndex        =   40
      Top             =   1150
      Width           =   8655
      Begin MSFlexGridLib.MSFlexGrid Msf_GrillaBenef 
         Height          =   1005
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   8370
         _ExtentX        =   14764
         _ExtentY        =   1773
         _Version        =   393216
         Rows            =   1
         Cols            =   7
         FixedCols       =   0
         BackColor       =   14745599
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   " Beneficiarios / No Beneficiarios "
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
         Index           =   13
         Left            =   120
         TabIndex        =   41
         Top             =   0
         Width           =   2895
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
      TabIndex        =   28
      Top             =   0
      Width           =   8655
      Begin VB.TextBox Txt_PenRut 
         Height          =   285
         Left            =   4920
         MaxLength       =   11
         TabIndex        =   1
         Top             =   360
         Width           =   1275
      End
      Begin VB.TextBox Txt_PenPoliza 
         Height          =   285
         Left            =   960
         MaxLength       =   10
         TabIndex        =   0
         Top             =   360
         Width           =   1185
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
         Left            =   7920
         Picture         =   "Frm_AntCertificado.frx":0000
         TabIndex        =   4
         ToolTipText     =   "Buscar"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Txt_PenDigito 
         Height          =   285
         Left            =   5520
         MaxLength       =   1
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton Cmd_BuscarPol 
         Height          =   375
         Left            =   7920
         Picture         =   "Frm_AntCertificado.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Buscar Póliza"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Póliza / Beneficiario  "
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
         Index           =   10
         Left            =   120
         TabIndex        =   38
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label Lbl_Endoso 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7320
         TabIndex        =   35
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nº Endoso"
         Height          =   255
         Index           =   12
         Left            =   6360
         TabIndex        =   34
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Lbl_PenNombre 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   960
         TabIndex        =   33
         Top             =   720
         Width           =   6855
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Póliza"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   31
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Dni Beneficiario"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   30
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
         Left            =   5280
         TabIndex        =   29
         Top             =   360
         Visible         =   0   'False
         Width           =   135
      End
   End
   Begin VB.Frame Fra_Botones 
      Height          =   1095
      Left            =   120
      TabIndex        =   27
      Top             =   6720
      Width           =   8685
      Begin VB.CommandButton Cmd_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   5400
         Picture         =   "Frm_AntCertificado.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   730
      End
      Begin VB.CommandButton Cmd_Grabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   1080
         Picture         =   "Frm_AntCertificado.frx":07DE
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   4320
         Picture         =   "Frm_AntCertificado.frx":0E98
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   6480
         Picture         =   "Frm_AntCertificado.frx":1552
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   3240
         Picture         =   "Frm_AntCertificado.frx":164C
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Eliminar 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   2160
         Picture         =   "Frm_AntCertificado.frx":1D06
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Eliminar Año"
         Top             =   240
         Width           =   720
      End
      Begin Crystal.CrystalReport Rpt_CertEstudio 
         Left            =   7920
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Msf_Grilla 
      Height          =   1605
      Left            =   120
      TabIndex        =   12
      Top             =   5160
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   2831
      _Version        =   393216
      Rows            =   1
      Cols            =   7
      FixedCols       =   0
      BackColor       =   14745599
   End
   Begin VB.Frame Fra_CertEst 
      Caption         =   "Antecedentes de Certificado de Estudios"
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
      Height          =   1720
      Left            =   120
      TabIndex        =   18
      Top             =   3360
      Width           =   8685
      Begin VB.CommandButton Cmd_Reliquidar 
         Caption         =   "&Reliquidar Pensión..."
         Height          =   375
         Left            =   4800
         TabIndex        =   11
         Top             =   1230
         Width           =   1815
      End
      Begin VB.TextBox Txt_FecRecep 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   8
         Top             =   600
         Width           =   1155
      End
      Begin VB.TextBox Txt_FecIniVig 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   6
         Top             =   280
         Width           =   1155
      End
      Begin VB.TextBox Txt_FecTerVig 
         Height          =   285
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   7
         Top             =   280
         Width           =   1155
      End
      Begin VB.ComboBox Cmb_RegEst 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1230
         Width           =   2820
      End
      Begin VB.TextBox Txt_Institucion 
         Height          =   285
         Left            =   1680
         TabIndex        =   9
         Top             =   920
         Width           =   6705
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Antecedentes de Certificado de Estudios"
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
         Index           =   11
         Left            =   120
         TabIndex        =   39
         Top             =   0
         Width           =   3555
      End
      Begin VB.Label Lbl_Efecto 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   6840
         TabIndex        =   37
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Lbl_FecIngreso 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   6840
         TabIndex        =   36
         Top             =   280
         Width           =   1575
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Fecha de Ingreso"
         Height          =   255
         Index           =   8
         Left            =   5280
         TabIndex        =   26
         Top             =   280
         Width           =   1455
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Fecha de Efecto"
         Height          =   255
         Index           =   9
         Left            =   5280
         TabIndex        =   25
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Fecha Recepción  "
         Height          =   255
         Index           =   5
         Left            =   165
         TabIndex        =   24
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Fecha Vigencia"
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   23
         Top             =   280
         Width           =   1395
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
         Left            =   3000
         TabIndex        =   22
         Top             =   280
         Width           =   210
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Régimen de Estudio"
         Height          =   255
         Index           =   7
         Left            =   165
         TabIndex        =   21
         Top             =   1230
         Width           =   1530
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Institución"
         Height          =   255
         Index           =   6
         Left            =   165
         TabIndex        =   20
         Top             =   920
         Width           =   885
      End
   End
End
Attribute VB_Name = "Frm_AntCertificado"
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

    Frm_AntCertificado.Top = 0
    Frm_AntCertificado.Left = 0
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Msf_Grilla_Click()
On Error GoTo Err_Grilla
Dim vlI As Integer


    Msf_Grilla.Col = 0
    vlPos = Msf_Grilla.RowSel
    Msf_Grilla.row = vlPos
    If (Msf_Grilla.Text = "") Or (Msf_Grilla.row = 0) Then
        Exit Sub
    End If
    Screen.MousePointer = 11

    Txt_FecIniVig = (Msf_Grilla.Text)

    Msf_Grilla.Col = 1
    Txt_FecTerVig = (Msf_Grilla.Text)

    Msf_Grilla.Col = 2
    Txt_Institucion = (Msf_Grilla.Text)

    Msf_Grilla.Col = 3
    vlCodFre = Trim(Msf_Grilla.Text)
    If Cmb_RegEst.Text <> "" Then
       For vlI = 0 To Cmb_RegEst.ListCount - 1
           Cmb_RegEst.ListIndex = vlI
           If Cmb_RegEst.Text = vlCodFre Then
              Exit For
           End If
       Next vlI
    End If

    Msf_Grilla.Col = 4
    Txt_FecRecep = (Msf_Grilla.Text)

    Msf_Grilla.Col = 5
    Lbl_FecIngreso = (Msf_Grilla.Text)

    Msf_Grilla.Col = 6
    Lbl_Efecto = (Msf_Grilla.Text)

    'Deshabilitar la Fecha de Inicio de Vigencia de la Póliza
    Txt_FecIniVig.Enabled = False

    Screen.MousePointer = 0

Exit Sub
Err_Grilla:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub


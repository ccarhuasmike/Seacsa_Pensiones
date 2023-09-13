VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Frm_AntCertificadoSup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenedor de Certificados de Supervivencia."
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   9765
   Begin VB.CommandButton Command1 
      Caption         =   "Hijos mayores de 18 con estudios "
      Height          =   255
      Left            =   120
      TabIndex        =   56
      Top             =   6240
      Width           =   7095
   End
   Begin VB.Frame frmOpciones 
      Caption         =   "Elegir la opcion de Certificado"
      Height          =   1575
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   3615
      Begin VB.OptionButton optCertifEst 
         Caption         =   "Certificado de Estudios"
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
         Height          =   195
         Left            =   480
         TabIndex        =   9
         Top             =   960
         Width           =   2295
      End
      Begin VB.OptionButton optCertifSup 
         Caption         =   "Certificado de Supervivencia"
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
         Left            =   480
         TabIndex        =   8
         Top             =   480
         Width           =   2895
      End
   End
   Begin VB.Frame Fra_CertEst 
      Caption         =   "  Antecedentes de Certificado de Supervivencia  "
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
      Height          =   1275
      Left            =   120
      TabIndex        =   41
      Top             =   4920
      Width           =   7125
      Begin VB.TextBox Txt_Institucion 
         Height          =   285
         Left            =   1680
         TabIndex        =   46
         Top             =   920
         Width           =   3225
      End
      Begin VB.TextBox Txt_FecTerVig 
         Height          =   285
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   45
         Top             =   280
         Width           =   1155
      End
      Begin VB.TextBox Txt_FecIniVig 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   44
         Top             =   280
         Width           =   1155
      End
      Begin VB.TextBox Txt_FecRecep 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   43
         Top             =   600
         Width           =   1155
      End
      Begin VB.CheckBox chkEst 
         Caption         =   "Activo"
         Height          =   255
         Left            =   6000
         TabIndex        =   42
         Top             =   240
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Institución"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   55
         Top             =   920
         Width           =   885
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
         TabIndex        =   54
         Top             =   280
         Width           =   210
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Fecha Vigencia"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   53
         Top             =   285
         Width           =   1395
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Fecha Recepción  "
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   52
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Efecto"
         Height          =   255
         Index           =   9
         Left            =   5160
         TabIndex        =   51
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Ingreso"
         Height          =   255
         Index           =   8
         Left            =   3000
         TabIndex        =   50
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Lbl_FecIngreso 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3720
         TabIndex        =   49
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label Lbl_Efecto 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5760
         TabIndex        =   48
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Antecedentes de Certificado"
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
         Left            =   240
         TabIndex        =   47
         Top             =   0
         Width           =   4155
      End
   End
   Begin VB.Frame Fra_Poliza 
      Caption         =   "  Póliza / Causante  "
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
      TabIndex        =   27
      Top             =   120
      Width           =   9495
      Begin VB.CommandButton Cmd_BuscarPol 
         Height          =   375
         Left            =   8760
         Picture         =   "Frm_AntCertificadoSup.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Buscar Póliza"
         Top             =   120
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
         Left            =   8760
         Picture         =   "Frm_AntCertificadoSup.frx":0102
         TabIndex        =   31
         ToolTipText     =   "Buscar"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox Txt_PenPoliza 
         Height          =   285
         Left            =   960
         MaxLength       =   10
         TabIndex        =   30
         Top             =   240
         Width           =   1185
      End
      Begin VB.TextBox Txt_PenNumIdent 
         Height          =   285
         Left            =   5280
         MaxLength       =   16
         TabIndex        =   29
         Top             =   240
         Width           =   1875
      End
      Begin VB.ComboBox Cmb_PenNumIdent 
         BackColor       =   &H00E0FFFF&
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   240
         Width           =   2235
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   39
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Póliza"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Lbl_PenNombre 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   960
         TabIndex        =   37
         Top             =   600
         Width           =   7695
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nº Endoso"
         Height          =   255
         Index           =   12
         Left            =   7320
         TabIndex        =   36
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Lbl_End 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   8160
         TabIndex        =   35
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Ident."
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   34
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Póliza / Causante"
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
         Left            =   240
         TabIndex        =   33
         Top             =   0
         Width           =   1815
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
      TabIndex        =   23
      Top             =   1080
      Width           =   9495
      Begin MSFlexGridLib.MSFlexGrid Msf_GrillaBenef 
         Height          =   1005
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   9210
         _ExtentX        =   16245
         _ExtentY        =   1773
         _Version        =   393216
         Rows            =   1
         Cols            =   7
         FixedCols       =   0
         BackColor       =   14745599
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   " Beneficiarios"
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
         TabIndex        =   26
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   " Beneficiarios"
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
         Index           =   7
         Left            =   120
         TabIndex        =   25
         Top             =   0
         Width           =   1335
      End
   End
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
      TabIndex        =   16
      Top             =   2400
      Width           =   9495
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
         TabIndex        =   22
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Lbl_NomBen 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4920
         TabIndex        =   21
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   17
         Left            =   4200
         TabIndex        =   20
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nº Ident."
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Lbl_CodTipoIdenBen 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   840
         TabIndex        =   18
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Lbl_NumIdenBen 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2520
         TabIndex        =   17
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame frmCErtEstudios 
      Caption         =   "Requisitos de Estudios"
      Enabled         =   0   'False
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
      Height          =   1635
      Left            =   7320
      TabIndex        =   10
      Top             =   4920
      Width           =   2295
      Begin VB.CheckBox chk_bno 
         Caption         =   "Boleta de Notas"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CheckBox chk_pes 
         Caption         =   "Plan de Estudios"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   1815
      End
      Begin VB.CheckBox chk_dju 
         Caption         =   "Declaración Jurada"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   1815
      End
      Begin VB.CheckBox chk_dni 
         Caption         =   "DNI"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Fra_Botones 
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   6600
      Width           =   9525
      Begin VB.CommandButton Cmd_Reliquidar 
         Caption         =   "&Reliquidar Pensión..."
         Height          =   375
         Left            =   7560
         TabIndex        =   15
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton Cmd_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   5400
         Picture         =   "Frm_AntCertificadoSup.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   730
      End
      Begin VB.CommandButton Cmd_Grabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   1080
         Picture         =   "Frm_AntCertificadoSup.frx":07DE
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   4320
         Picture         =   "Frm_AntCertificadoSup.frx":0E98
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   6480
         Picture         =   "Frm_AntCertificadoSup.frx":1552
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   3240
         Picture         =   "Frm_AntCertificadoSup.frx":164C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Eliminar 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   2160
         Picture         =   "Frm_AntCertificadoSup.frx":1D06
         Style           =   1  'Graphical
         TabIndex        =   1
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
      TabIndex        =   40
      Top             =   3240
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   2831
      _Version        =   393216
      Rows            =   1
      Cols            =   11
      FixedCols       =   0
      BackColor       =   14745599
   End
End
Attribute VB_Name = "Frm_AntCertificadoSup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Sql As String
Dim vlRegistro   As ADODB.Recordset
Dim vlRegistro1  As ADODB.Recordset
Dim vlRegistro2  As ADODB.Recordset
Dim vlNumEndoso  As String
Dim vlNumOrden   As Integer
Dim vlNombre     As String
Dim vlCodTab     As String
Dim vlCodFre     As String
Dim vlSwVigPol   As Boolean
Dim Habilita     As Boolean
Dim vlPasa       As Boolean
Dim vlPos        As Double
Dim vlSw         As Boolean
Dim flBuscaCer   As Boolean
Dim vlFechaNac   As String
Dim vlIniVig     As String
Dim vlSwGrilla   As Boolean
Dim vlFechaMatrimonio As String, vlFechaFallecimiento As String
Dim vlFechaIni  As String, vlFechaTer As String, vlFechaRec As String
Dim vlOp        As String, vlOperacion As String

Dim vlInicio As String
Dim vlTermino As String
Dim vlRecepcion As String
Dim vlIngreso As String
Dim vlEfecto As String
Dim vlAnno As String
Dim vlMes As String
Dim vlDia As String

Dim vlDNI, vlDCJ, vlPes, vlBon, vlEST As String 'RRR 19/9/13



Dim vlNumPoliza As String
Dim vlNumOrdenCau As Integer
Dim vlCodTipoIdenBenCau As String
Dim vlNomTipoIdenBenCau As String
Dim vlNumIdenBenCau As String
Dim vlGlsNomBenCau As String
Dim vlGlsNomSegBenCau As String
Dim vlGlsPatBenCau As String
Dim vlGlsMatBenCau As String
Dim vlCodPar As String
Dim vlNombreCompleto As String

Const clCodParCausante As String * 2 = "99"
Const clCodParDes As String * 1 = "D" 'Parentesco Descendiente

Dim vlLargoTipoIden    As Integer 'sirve para llenar la grilla
Dim vlPosicionTipoIden As Integer 'sirve para llenar la grilla
Dim vlTipoCertif As String 'RRR 19/9/13

Private Sub Cmb_RegEst_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Cmd_Grabar.SetFocus
End If
End Sub

Private Sub Cmb_PenNumIdent_Click()
If (Cmb_PenNumIdent <> "") Then
    vlPosicionTipoIden = Cmb_PenNumIdent.ListIndex
    vlLargoTipoIden = Cmb_PenNumIdent.ItemData(vlPosicionTipoIden)
    If (vlLargoTipoIden = 0) Then
        Txt_PenNumIdent.Text = "0"
        Txt_PenNumIdent.Enabled = False
    Else
        Txt_PenNumIdent = ""
        Txt_PenNumIdent.Enabled = True
        Txt_PenNumIdent.MaxLength = vlLargoTipoIden
        If (Txt_PenNumIdent <> "") Then Txt_PenNumIdent.Text = Mid(Txt_PenNumIdent, 1, vlLargoTipoIden)
    End If
End If
End Sub

Private Sub Cmb_PenNumIdent_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        If (Txt_PenNumIdent.Enabled = True) Then
            Txt_PenNumIdent.SetFocus
        Else
            Cmd_BuscarPol.SetFocus
        End If
    End If
End Sub

Private Sub Cmd_Buscar_Click()
On Error GoTo Err_Buscar

    Frm_Busqueda.flInicio ("Frm_AntCertificadoSup")

Exit Sub
Err_Buscar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_BuscarPol_Click()
On Error GoTo Err_Buscar

    If Trim(Txt_PenPoliza) <> "" Or (Trim(Cmb_PenNumIdent) <> "" And Trim(Txt_PenNumIdent) <> "") Then
'       If ((Trim(Cmb_PenNumIdent.Text)) <> "") Then
          If (Trim(Txt_PenNumIdent.Text) = "") Then
'             MsgBox "Debe ingresar el Número de Identificación.", vbCritical, "Error de Datos"
'             Txt_PenNumIdent.SetFocus
'             Exit Sub
'          End If
          'Txt_PenRut = Format(Txt_PenRut, "#,#0")
          Txt_PenNumIdent = Trim(UCase(Txt_PenNumIdent))
       End If
       'Permite Buscar los Datos del Beneficiario
       Call flValidarBen
   Else
     MsgBox "Debe ingresar el NºPóliza o la Identificación del Pensionado", vbCritical, "Error de Datos"
     Txt_PenPoliza.SetFocus
   End If

Exit Sub
Err_Buscar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Cancelar_Click()
On Error GoTo Err_Cancelar

    Call flLmpCert
    
    Fra_Beneficiarios.Enabled = False
    Msf_GrillaBenef.Enabled = False
    Fra_CertEst.Enabled = False
    Msf_Grilla.Enabled = False
    Fra_Poliza.Enabled = True
    
    
    Txt_PenPoliza = ""
    If (Cmb_PenNumIdent.ListCount > 0) Then
        Cmb_PenNumIdent.ListIndex = 0
    End If
    Txt_PenNumIdent = ""
    Lbl_End = ""
    Lbl_PenNombre = ""
    Lbl_Efecto = ""
    
    'CMV-20060616 I
    Lbl_CodTipoIdenBen = ""
    Lbl_NumIdenBen = ""
    Lbl_NomBen = ""
    Lbl_FecIngreso = ""
    'CMV-20060616 F
    
    Call flLmpGrilla
    Call flInicializaGrillaBenef
    Txt_PenPoliza.SetFocus

Exit Sub
Err_Cancelar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Eliminar_Click()
On Error GoTo Err_Eliminar

'    If vlSwVigPol = False Then
'       MsgBox "Póliza Ingresada No se Encuentra vigente para el Sistema", vbCritical, "Operación Cancelada"
'       Exit Sub
'    End If

    If Fra_Poliza.Enabled = True Then
       Exit Sub
    End If

    Txt_FecIniVig = Trim(Txt_FecIniVig)

    If Txt_FecIniVig = "" Then
       MsgBox "Debe Ingresar Fecha de Inicio de Vigencia", vbInformation, "Operación Cancelada"
       Txt_FecIniVig.SetFocus
       Exit Sub
    End If

    vlSwVigPol = True

    If fgValidaVigenciaPoliza(Trim(Txt_PenPoliza), Trim(Txt_FecIniVig)) = False Then
       MsgBox "La Fecha Ingresada no se Encuentra dentro del Rango de Vigencia de la Póliza" & Chr(13) & _
              "O esta No Vigente. No se Ingresara Ni Modificara Información. ", vbCritical, "Operación Cancelada"
       Screen.MousePointer = 0
       vlSwVigPol = False
       Exit Sub
    End If

    If fgValidaPagoPension(Lbl_Efecto, Txt_PenPoliza, vlNumOrden) = False Then
       MsgBox " Ya se ha realizado el proceso de Cálculo de Pensión para ésta fecha " & Chr(13) & _
              "                El Registro No se puede Eliminar", vbCritical, "Operación Cancelada"
       Screen.MousePointer = 0
       Cmd_Salir.SetFocus
       Exit Sub
    End If

    vlOperacion = ""
    Screen.MousePointer = 11

    vlIniVig = Txt_FecIniVig
    vlIniVig = Format(CDate(Trim(vlIniVig)), "yyyymmdd")

    vgQuery = "SELECT NUM_POLIZA,NUM_ENDOSO,NUM_ORDEN,fec_inicer FROM pp_tmae_certificado WHERE "
    vgQuery = vgQuery & "NUM_POLIZA = '" & Txt_PenPoliza & "' And "
    'vgQuery = vgQuery & "NUM_ENDOSO = " & vlNumEndoso & " AND "
    vgQuery = vgQuery & "NUM_ORDEN = " & vlNumOrden & " AND "
    vgQuery = vgQuery & "fec_inicer = '" & vlIniVig & "' "
    Set vgRs = vgConexionBD.Execute(vgQuery)
    If Not (vgRs.EOF) Then
        vlOperacion = "E"
    End If
    vgRs.Close

    If (vlOperacion = "E") Then
        vgRes = MsgBox(" ¿ Esta seguro que desea Eliminar los Datos ? ", vbQuestion + vbYesNo + 256, "Operación de Eliminación")
        If vgRes <> 6 Then
            Screen.MousePointer = 0
            Exit Sub
        End If

        vgQuery = "DELETE FROM pp_tmae_certificado WHERE "
        vgQuery = vgQuery & "NUM_POLIZA = '" & Txt_PenPoliza & "' And "
        'vgQuery = vgQuery & "NUM_ENDOSO = " & vlNumEndoso & " AND "
        vgQuery = vgQuery & "NUM_ORDEN = " & vlNumOrden & " AND "
        vgQuery = vgQuery & "fec_inicer = '" & (vlIniVig) & "' "
        vgConexionBD.Execute (vgQuery)

        Call flLmpGrilla
        Call flCargaGrilla
        Call flLmpCert
        
        Txt_FecIniVig.Enabled = True
        Txt_FecIniVig.SetFocus
    End If
    Screen.MousePointer = 0

Exit Sub
Err_Eliminar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Grabar_Click()
Dim vlResp As String
Dim vlFecFin As String
Dim vlFecEfe As String
Dim vlFecRecep As String
Dim vlFecIng As String
Dim vlAnnoFin As String
Dim vlEdad As String
Dim vlAnnoBen As String
Dim vlDiferencia As String
Dim Mto_Edad As Double
On Error GoTo Err_Grabar

'    If vlSwVigPol = False Then
'       MsgBox "Póliza Ingresada No se Encuentra vigente para el Sistema", vbCritical, "Operación Cancelada"
'       Exit Sub
'    End If

    If Fra_Poliza.Enabled = True Then
       Exit Sub
    End If

    If Txt_FecIniVig = "" Then
       MsgBox "Debe Ingresar Fecha de Inicio de Vigencia", vbInformation, "Operación Cancelada"
       Txt_FecIniVig.SetFocus
       Exit Sub
    Else
       vlFechaIni = Trim(Txt_FecIniVig)
       If (flValidaFecha(vlFechaIni) = False) Then
           Txt_FecIniVig = ""
           Txt_FecIniVig.SetFocus
           Exit Sub
       End If
    End If

    vlSwVigPol = True

    If fgValidaVigenciaPoliza(Trim(Txt_PenPoliza), Trim(Txt_FecIniVig)) = False Then
       MsgBox "La Fecha Ingresada no se Encuentra dentro del Rango de Vigencia de la Póliza" & Chr(13) & _
              "O esta No Vigente. No se Ingresara Ni Modificara Información. ", vbCritical, "Operación Cancelada"
       Screen.MousePointer = 0
       vlSwVigPol = False
       Exit Sub
    End If

    If Txt_FecTerVig = "" Then
       MsgBox "Debe Ingresar Fecha de Término de Vigencia", vbInformation, "Operación Cancelada"
       Txt_FecTerVig.SetFocus
       Exit Sub
    Else
        vlFechaTer = Trim(Txt_FecTerVig)
        If (flValidaFecha(vlFechaTer) = False) Then
            Txt_FecTerVig = ""
            Txt_FecTerVig.SetFocus
            Exit Sub
        End If
    End If

    If (CDate(Format(Txt_FecTerVig, "dd/mm/yyyy")) < CDate(Format(Txt_FecIniVig, "dd/mm/yyyy"))) Then
        MsgBox "La Fecha Término es menor a la Fecha de Inicio de Vigencia.", vbCritical, "Dato Incorrecto"
        Exit Sub
    End If

    If Txt_FecRecep = "" Then
       MsgBox "Debe Ingresar Fecha de Recepción ", vbInformation, "Operación Cancelada"
       Exit Sub
    Else
       vlFechaRec = Trim(Txt_FecRecep)
       If (flValidaFecha(vlFechaRec) = False) Then
           Txt_FecRecep = ""
           Txt_FecRecep.SetFocus
           Exit Sub
       End If
    End If
    
'No es Obligatoria ABV
'    If Txt_Institucion = "" Then
'       MsgBox "Debe Ingresar la Institución de Educación Superior.", vbInformation, "Operación Cancelada"
'       Txt_Institucion.SetFocus
'       Exit Sub
'    End If
    Txt_Institucion = Trim(UCase(Txt_Institucion))

    If Lbl_FecIngreso = "" Then
       MsgBox " Falta Fecha de Ingreso", vbCritical, "Operación Cancelada"
       Exit Sub
    End If

    If Lbl_Efecto = "" Then
       MsgBox " Falta Fecha de Efecto", vbCritical, "Operación Cancelada"
       Exit Sub
    End If

    Screen.MousePointer = 11

    vlOp = ""
'No existe Frecuencia para este Certificado ABV
'    vlCodFre = Trim(Mid(Cmb_RegEst, 1, (InStr(1, Cmb_RegEst, "-") - 1)))
    vlFecFin = Txt_FecTerVig
    vlFecFin = Format(CDate(Trim(vlFecFin)), "yyyymmdd")
    vlAnnoFin = Mid(vlFecFin, 1, 4)
    vlFecRecep = Txt_FecRecep
    vlFecRecep = Format(CDate(Trim(vlFecRecep)), "yyyymmdd")
    vlFecIng = Lbl_FecIngreso
    vlFecIng = Format(CDate(Trim(vlFecIng)), "yyyymmdd")
    vlIniVig = Txt_FecIniVig
    vlIniVig = Format(CDate(Trim(vlIniVig)), "yyyymmdd")
    vlFecEfe = Lbl_Efecto
    vlFecEfe = Format(CDate(Trim(vlFecEfe)), "yyyymmdd")

    'Verificar Si el Beneficiario Existe y Esta Vivo
    vlOperacion = ""
    vlFechaMatrimonio = ""
    vlCodTipoIdenBenCau = fgObtenerCodigo_TextoCompuesto(Lbl_CodTipoIdenBen)
    vlNumIdenBenCau = Trim(UCase(Lbl_NumIdenBen))

    vgSql = ""
    vgSql = "SELECT NUM_POLIZA,FEC_fallben,fec_nacben "
    vgSql = vgSql & " FROM PP_TMAE_BEN "
    vgSql = vgSql & " Where "
    vgSql = vgSql & " NUM_POLIZA =  '" & (Txt_PenPoliza) & "' AND "
    vgSql = vgSql & " cod_tipoidenben = " & vlCodTipoIdenBenCau & " AND "
    vgSql = vgSql & " num_idenben = '" & vlNumIdenBenCau & "' "
    vgSql = vgSql & " ORDER BY NUM_ENDOSO DESC "
    Set vlRegistro = vgConexionBD.Execute(vgSql)
    If Not vlRegistro.EOF Then
        If IsNull(vlRegistro!Fec_FallBen) Or (vlRegistro!Fec_FallBen = "") Then
            vlOperacion = "S"
        Else
            vlFechaFallecimiento = vlRegistro!Fec_FallBen
        End If
    End If
    vlRegistro.Close

    'Validar que el Beneficiario se encuentre Vivo
    'Ya que si lo está le corresponde el ingreso de Certificado
    If (vlFechaFallecimiento <> "") Then
        If (CLng(Mid(vlFechaFallecimiento, 1, 6)) < CLng(Mid(vlIniVig, 1, 6))) Then
            MsgBox "No es posible registrar el Certificado, ya que el Pensionado/Beneficiario se encuentra Fallecido.", vbCritical, "Operación Cancelado"
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If

    If (vlOperacion = "") Then
        MsgBox "No es posible registrar los datos del Certificado de " & Chr(13) & _
        "este Beneficiario, ya que no se encuentra Registrado.", vbCritical, "Operación Cancelada"
        Screen.MousePointer = 0
        Exit Sub
    End If

    vlOp = ""
    'Verifica la existencia de la vigencia
    vgSql = "select num_poliza,fec_tercer from pp_tmae_certificado where "
    vgSql = vgSql & " NUM_POLIZA = '" & Trim(Txt_PenPoliza) & "' and "
    vgSql = vgSql & " NUM_ORDEN = " & Trim(vlNumOrden) & " and "
    vgSql = vgSql & " fec_inicer = '" & vlIniVig & "' and"
    vgSql = vgSql & " cod_tipo = '" & vlTipoCertif & "'"
    Set vlRegistro = vgConexionBD.Execute(vgSql)
    If vlRegistro.EOF Then
        vlOp = "I"
    Else
        vlOp = "A"
    End If
'    vlRegistro.Close

    If (vlOp = "A") Then
        vgRes = MsgBox("¿ Está seguro que desea Modificar los Datos ?", 4 + 32 + 256, "Operación de Actualización")
        If vgRes <> 6 Then
            Cmd_Salir.SetFocus
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If
    If (vlOp = "I") Then
        vlResp = MsgBox(" ¿ Está seguro que desea ingresar los Datos ?", 4 + 32 + 256, "Proceso de Ingreso de Datos")
        If vlResp <> 6 Then
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If

   'Si la accion es modificar
    If (vlOp = "A") Then
        If (vlRegistro!FEC_TERCER) <> (vlFecFin) Then
            Call flComparaFechaMod
            If vlSw = True Then
               Sql = ""
               Sql = "update pp_tmae_certificado set"
               Sql = Sql & " fec_tercer = '" & (vlFecFin) & "',"
               Sql = Sql & " NUM_ENDOSO = " & (vlNumEndoso) & ","
               'Sql = Sql & " COD_FRECUENCIA = '" & (vlCodFre) & "',"
               If (Txt_Institucion <> "") Then
                    Sql = Sql & " GLS_NOMINSTITUCION = '" & Trim(Txt_Institucion) & "',"
               Else
                    Sql = Sql & " GLS_NOMINSTITUCION = NULL,"
               End If
               Sql = Sql & " FEC_RECCIA = '" & (vlFecRecep) & "',"
               Sql = Sql & " COD_USUARIOMODI = '" & (vgUsuario) & "',"
               Sql = Sql & " FEC_MODI = '" & Format(Date, "yyyymmdd") & "',"
               Sql = Sql & " HOR_MODI = '" & Format(Time, "hhmmss") & "',"
               'If Txt_FecIniVig <> Lbl_Efecto Then 'Corresponde Reliquidar
                    Sql = Sql & " COD_INDRELIQUIDAR = 'S',"
'               Else
'                    Sql = Sql & " COD_INDRELIQUIDAR = 'N'"
'               End If
               If vlTipoCertif = "EST" Then
                    Sql = Sql & " IND_DNI = '" & IIf(chk_dni.Value, "S", "") & "',"
                    Sql = Sql & " IND_DJU = '" & IIf(chk_dju.Value, "S", "") & "',"
                    Sql = Sql & " IND_PES = '" & IIf(chk_pes.Value, "S", "") & "',"
                    Sql = Sql & " IND_BNO = '" & IIf(chk_bno.Value, "S", "") & "',"
                    Sql = Sql & " EST_ACT = '" & IIf(chkEst.Value, "1", "0") & "'"
               End If
               Sql = Sql & " Where "
               Sql = Sql & " NUM_POLIZA = '" & Trim(Txt_PenPoliza) & "' and "
               Sql = Sql & " NUM_ORDEN = " & Trim(vlNumOrden) & " and "
               Sql = Sql & " COD_TIPO = '" & Trim(vlTipoCertif) & "' and "
               Sql = Sql & " fec_inicer = '" & (vlIniVig) & "'"
               vgConexionBD.Execute (Sql)
            Else
               vlRegistro.Close
               Screen.MousePointer = 0
               Call flLmpCert
               Txt_FecIniVig.Enabled = True
               Txt_FecIniVig.SetFocus
               Exit Sub
            End If
        Else
            Sql = ""
            Sql = "update pp_tmae_certificado set"
            Sql = Sql & " NUM_ENDOSO = " & (vlNumEndoso) & ","
            'Sql = Sql & " COD_FRECUENCIA = '" & (vlCodFre) & "',"
            If (Txt_Institucion <> "") Then
                Sql = Sql & " GLS_NOMINSTITUCION = '" & Trim(Txt_Institucion) & "',"
            Else
                 Sql = Sql & " GLS_NOMINSTITUCION = NULL,"
            End If
            Sql = Sql & " FEC_RECCIA = '" & (vlFecRecep) & "',"
            Sql = Sql & " COD_USUARIOMODI = '" & (vgUsuario) & "',"
            Sql = Sql & " FEC_MODI = '" & Format(Date, "yyyymmdd") & "',"
            Sql = Sql & " HOR_MODI = '" & Format(Time, "hhmmss") & "',"
            'If Txt_FecIniVig <> Lbl_Efecto Then 'Corresponde Reliquidar
                 Sql = Sql & " COD_INDRELIQUIDAR = 'S'" 'EISEI RICHARD
'            Else
'                 Sql = Sql & " COD_INDRELIQUIDAR = 'N'"
'            End If
            If vlTipoCertif = "EST" Then
                Sql = Sql & "," 'EISEI RICHARD
                Sql = Sql & " IND_DNI = '" & IIf(chk_dni.Value, "S", "") & "',"
                Sql = Sql & " IND_DJU = '" & IIf(chk_dju.Value, "S", "") & "',"
                Sql = Sql & " IND_PES = '" & IIf(chk_pes.Value, "S", "") & "',"
                Sql = Sql & " IND_BNO = '" & IIf(chk_bno.Value, "S", "") & "',"
                Sql = Sql & " EST_ACT = '" & IIf(chkEst.Value, "1", "0") & "'"
            End If
            Sql = Sql & " Where "
            Sql = Sql & " NUM_POLIZA = '" & Trim(Txt_PenPoliza) & "' and "
            Sql = Sql & " NUM_ORDEN = " & Trim(vlNumOrden) & " and "
            Sql = Sql & " COD_TIPO = '" & Trim(vlTipoCertif) & "' and "
            Sql = Sql & " fec_inicer = '" & (vlIniVig) & "'"
            vgConexionBD.Execute (Sql)
        End If
    Else
       'Inserta los Datos en la Tabla pp_tmae_certificado
        Call flComparaFechaIngresada
        If vlSw = False Then
           Call flLmpCert
           Txt_FecIniVig.Enabled = True
           Txt_FecIniVig.SetFocus
           Exit Sub
        Else
           Call flComparaFechaExistente
           If vlSw = False Then
              Call flLmpCert
              Txt_FecIniVig.Enabled = True
              Txt_FecIniVig.SetFocus
              Exit Sub
           End If
        End If

        Sql = ""
        Sql = "insert into pp_tmae_certificado ("
        Sql = Sql & "NUM_POLIZA,NUM_ORDEN,fec_inicer,"
        Sql = Sql & "fec_tercer,NUM_ENDOSO,COD_TIPO,"
        'Sql = Sql & "COD_FRECUENCIA,"
        Sql = Sql & "GLS_NOMINSTITUCION,"
        Sql = Sql & "FEC_RECCIA,FEC_INGCIA,FEC_EFECTO,COD_USUARIOCREA,"
        Sql = Sql & "FEC_CREA,HOR_CREA,COD_INDRELIQUIDAR, IND_DNI, IND_DJU, IND_PES, IND_BNO, EST_ACT"
        Sql = Sql & " "
        Sql = Sql & ") values ("
        Sql = Sql & "'" & Trim(Txt_PenPoliza) & "',"
        Sql = Sql & "" & Trim(vlNumOrden) & ","
        Sql = Sql & "'" & (vlIniVig) & "',"
        Sql = Sql & "'" & (vlFecFin) & "',"
        Sql = Sql & "" & (vlNumEndoso) & ","
        'Sql = Sql & "'" & (vlCodFre) & "',"
        Sql = Sql & "'" & (vlTipoCertif) & "',"
        If (Txt_Institucion <> "") Then
            Sql = Sql & "'" & Trim(Txt_Institucion) & "',"
        Else
            Sql = Sql & "NULL,"
        End If
        Sql = Sql & "'" & (vlFecRecep) & "',"
        Sql = Sql & "'" & (vlFecIng) & "',"
        Sql = Sql & "'" & (vlFecEfe) & "',"
        Sql = Sql & "'" & (vgUsuario) & "',"
        Sql = Sql & "'" & Format(Date, "yyyymmdd") & "',"
        Sql = Sql & "'" & Format(Time, "hhmmss") & "','S',"
        'Sql = Sql & "'" & IIf(Txt_FecIniVig <> Lbl_Efecto, "S", "N") & "'"
        Sql = Sql & "'" & IIf(chk_dni.Value, "S", "") & "',"
        Sql = Sql & "'" & IIf(chk_dju.Value, "S", "") & "',"
        Sql = Sql & "'" & IIf(chk_pes.Value, "S", "") & "',"
        Sql = Sql & "'" & IIf(chk_bno.Value, "S", "") & "',"
        Sql = Sql & "'" & IIf(chkEst.Value, "1", "0") & "'"
        Sql = Sql & ")"
        vgConexionBD.Execute (Sql)
    End If
    vlRegistro.Close
    
    Dim cont As Integer
    
    cont = 0
    
    If vlTipoCertif = "EST" Then
    
        If chk_dni.Value = 1 Then cont = cont + 1
        If chk_dju.Value = 1 Then cont = cont + 1
        If chk_pes.Value = 1 Then cont = cont + 1
        If chk_bno.Value = 1 Then cont = cont + 1
        
        If cont = 4 Then
            Sql = ""
            Sql = "update pp_tmae_ben set cod_estpension='99' where "
            Sql = Sql & " NUM_POLIZA = '" & Trim(Txt_PenPoliza) & "' and  NUM_ORDEN = " & Trim(vlNumOrden) & " and num_endoso=" & vlNumEndoso & ""
            vgConexionBD.Execute (Sql)
        Else
            vgRes = MsgBox("El joven no a cumplido con completar los requerimientos obligatorios. No recibira pension.", vbExclamation, "Importante!")
        End If
    
    End If
    
   
    
    
    If (vlOp <> "") Then
        'Limpia los Datos de la Pantalla

        Call flLmpGrilla
        Call flCargaGrilla
'        flLmpCert
        Txt_FecIniVig.Enabled = True
        Txt_FecIniVig.SetFocus
    End If


    MsgBox "Proceso de certificado Terminado.", vbExclamation, "Operación Certificado"
    Screen.MousePointer = 0

Exit Sub
Err_Grabar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(2) & Err.Description & " ]"
    End Select
End Sub

Private Sub Cmd_Imprimir_Click()
On Error GoTo Err_Imprimir

    If Fra_Poliza.Enabled = True Then
       Exit Sub
    End If

    'Imprime el Reporte de Variables
    flImpresion

Exit Sub
Err_Imprimir:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Limpiar_Click()
On Error GoTo Err_Limpiar

   If Txt_FecIniVig <> "" Then
      Call flLmpCert
      Lbl_Efecto = ""
'     Call flBuscaFecServ

      Lbl_FecIngreso = fgBuscaFecServ

      'Lbl_FecIngreso = DateSerial(Year(Date), Month(Date), Day(Date))
      Txt_FecIniVig.Enabled = True
      Txt_FecIniVig.SetFocus
   End If

Exit Sub
Err_Limpiar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Reliquidar_Click()
Dim vlSql As String
Dim vlNumReliq As Double

If Frm_AFReliquidacion.vpEstado = "A" Then
    MsgBox "Debe cerrar Formulario de Reliquidaciones para volver a Reliquidar", vbCritical, Me.Caption
    Exit Sub
End If
If Trim(Txt_PenPoliza) = "" Then
    MsgBox "Debe Ingresar Póliza", vbCritical
    Exit Sub
End If
''If vlNumOrdenPar = 0 Then
''    MsgBox "Debe seleccionar Causante de Asignación Familiar", vbCritical
''    Exit Sub
''End If
If Not IsDate(Txt_FecIniVig) Then
    MsgBox "Debe Ingresar Fecha de Inicio de Vigencia Válida", vbCritical
    Exit Sub
End If
'Lbl_Efecto = DateAdd("d", 1, Lbl_FecTermino)

'Validar si tiene Derecho a Pensión

If flReliquidar(Txt_PenPoliza, Lbl_End, vlNumOrden, Format(Txt_FecIniVig, "yyyymmdd"), vlNumReliq) Then
    Frm_AFReliquidacion.Txt_PenPoliza = Txt_PenPoliza
    'CMV-20051124 I
    'Modificado por que el encabezado del formulario de cer.est. ahora contine
    'los datos del Causante y no de la carga
    'Frm_AFReliquidacion.Txt_PenRut = Txt_PenRut
    Frm_AFReliquidacion.Cmb_PenNumIdent = " " & Lbl_CodTipoIdenBen
    Frm_AFReliquidacion.Txt_PenNumIdent = Lbl_NumIdenBen
    'CMV-20051124 F
    Frm_AFReliquidacion.Lbl_Endoso = Lbl_End
    Frm_AFReliquidacion.vpEfecto = Lbl_Efecto
    If vlNumReliq > 0 Then 'Actualización
        Frm_AFReliquidacion.txt_NumReliq = vlNumReliq
        Frm_AFReliquidacion.vpAccion = "M" 'Modificar
        'marco
        Frm_AFReliquidacion.vpNumOrden = vlNumOrden
    Else 'Reliquidación Nueva
        Frm_AFReliquidacion.vpAccion = "N" 'Nueva
        Frm_AFReliquidacion.txt_NumReliq = ""
        Frm_AFReliquidacion.vpIndAF = 0
        Frm_AFReliquidacion.vpIndGE = 0
        Frm_AFReliquidacion.vpIndPension = 1
        Frm_AFReliquidacion.vpNumOrden = vlNumOrden
        Frm_AFReliquidacion.vpNumOrdenRec = -1
        Frm_AFReliquidacion.txt_PerDesde = Format(Txt_FecIniVig, "mm/yyyy")
        If CDate(Txt_FecTerVig) < CDate(Lbl_Efecto) Then
            Frm_AFReliquidacion.txt_PerHasta = Format(Txt_FecTerVig, "mm/yyyy")
        Else
            Frm_AFReliquidacion.txt_PerHasta = Format(DateAdd("d", -1, CDate(Lbl_Efecto)), "mm/yyyy")
        End If
   End If
    'Construye Query que actualiza Tabla con número de Reliquidación
    vlSql = "UPDATE pp_tmae_certificado"
    Frm_AFReliquidacion.vpSQLUpdate = vlSql
    vlSql = " WHERE num_poliza = '" & Txt_PenPoliza & "'"
    vlSql = vlSql & " AND num_orden = " & vlNumOrden
    vlSql = vlSql & " AND fec_inicer = '" & Format(Txt_FecIniVig, "yyyymmdd") & "'"
    Frm_AFReliquidacion.vpSQLWhere = vlSql

    'Trae Datos de la Póliza
    Frm_AFReliquidacion.Cmd_BuscarPol_Click

End If
End Sub

Private Function flReliquidar(iPoliza As String, iEndoso As Integer, iOrden As Integer, iFecActiva As String, oNumReliq As Double) As Boolean
    Dim vlTB As ADODB.Recordset
    Dim vlEndoso As Integer
    flReliquidar = False
    vgSql = "SELECT cer.cod_indreliquidar AS ind, cer.num_reliq AS reliq"
    vgSql = vgSql & " FROM pp_tmae_certificado cer"
    vgSql = vgSql & " WHERE cer.num_poliza = '" & iPoliza & "'"
    vgSql = vgSql & " AND cer.num_orden = " & iOrden
    vgSql = vgSql & " AND cer.fec_inicer = '" & iFecActiva & "'"
    Set vlTB = vgConexionBD.Execute(vgSql)
    If vlTB.EOF Then
        MsgBox "Debe grabar el Registro antes de Reliquidar", vbCritical
        Exit Function
    End If
    If vlTB!ind = "N" Then
        MsgBox "La información ingresada no requiere una Reliquidación", vbCritical
        Exit Function
    End If
    oNumReliq = vlTB!reliq
'    If vlTB!reliq > 0 Then
'        MsgBox "La información seleccionada ya fue Reliquidada", vbCritical
'        Exit Function
'    End If

    'Valida si tiene derecho a Pensión
    vgSql = ""
    If iEndoso > 1 Then
        vgSql = "SELECT ben.cod_estpension"
        vgSql = vgSql & " FROM pp_tmae_ben ben, pp_tmae_endoso en"
        vgSql = vgSql & " WHERE en.num_poliza = ben.num_poliza"
        'vgSql = vgSql & " AND en.num_endoso = ben.num_endoso"
        vgSql = vgSql & " AND en.num_endoso = (ben.num_endoso - 1)" 'hqr 08/06/2006 El endoso del ben es 1 mayor a la tabla del Endoso
        vgSql = vgSql & " AND en.num_poliza = '" & iPoliza & "'"
        vgSql = vgSql & " AND ben.num_orden = " & iOrden
        vgSql = vgSql & " AND en.fec_efecto <= '" & iFecActiva & "'"
        'HQR 26/08/2005 Descomentar cuando se agregue el Campo
        vgSql = vgSql & " AND en.fec_finefecto >= '" & iFecActiva & "'"
        vgSql = vgSql & " UNION "
    End If
    vgSql = vgSql & "SELECT ben.cod_estpension"  ' Obtiene Estado de la Pension del Primer Endoso, si hay más de un Endoso se pregunta además que la Fecha de Activaciòn del Certificado sea menor a la fecha de efecto del segundo endoso
    vgSql = vgSql & " FROM pp_tmae_ben ben "
    vgSql = vgSql & " WHERE ben.num_poliza = '" & iPoliza & "'"
    vgSql = vgSql & " AND ben.num_endoso = 1" 'El primer Endoso
    vgSql = vgSql & " AND ben.num_orden = " & iOrden
    If iEndoso > 1 Then 'Verificar solo para las fechas en que estaba activo el Endoso 1
        'ABV-20051124 I
        'vgSql = vgSql & " AND ben.fec_inipension <= '" & iFecActiva & "'"
        'ABV-20051124 F
        vgSql = vgSql & " AND ben.fec_inipagopen <= '" & iFecActiva & "'"
        vgSql = vgSql & " AND '" & iFecActiva & "' < ("
        vgSql = vgSql & " SELECT end2.fec_efecto FROM pp_tmae_endoso end2"
        vgSql = vgSql & " WHERE end2.num_poliza = '" & iPoliza & "'"
        'vgSql = vgSql & " AND end2.num_endoso = 2)"
        vgSql = vgSql & " AND end2.num_endoso = 1)" 'HQR 08/06/2006 El endoso 1 de la tabla endoso es equivalente al endoso 2 de la tabla de beneficiarios
    End If
    vgSql = vgSql & " ORDER BY cod_estpension DESC" 'Para que salga último el 10
    Set vlRegistro = vgConexionBD.Execute(vgSql)
    If Not vlRegistro.EOF Then
        If vlRegistro!Cod_EstPension = "10" Then
            MsgBox "Beneficiario no tiene Derecho a Pensión, no se puede Reliquidar." & Chr(13) & "Para Reliquidar Asignación Familiar lo debe hacer desde la Activación de Cargas.", vbCritical, "Error"
            Exit Function
        End If
    Else
        MsgBox "Beneficiario no existe, no se puede Reliquidar", vbCritical, "Error"
        Exit Function
    End If
    flReliquidar = True
End Function


Private Sub cmd_salir_Click()
On Error GoTo Err_Salir

    Screen.MousePointer = 11
    flLmpCert
    Lbl_Efecto = ""
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

Private Sub Command1_Click()
On Error GoTo Err_flInformeCerEst
'Certificados de Supervivencia
   Screen.MousePointer = 11
   
   Dim cadena, vlSql, vlFechaTermino As String
   Dim objRep As New ClsReporte
   Dim vlRS18 As New ADODB.Recordset
   
   vlFechaTermino = Mid(CStr(Now), 1, 10) 'Format(Txt_FecIniVig, "mm/yyyy") 'Mid(FecCalculo, 7, 8) & "/" & Mid(FecCalculo, 5, 2) & "/" & Mid(FecCalculo, 1, 4)
   'vgPalabra = ""
   vgPalabra = "Beneficiarios de Pensión Garantizada que cumplieron 18 años con certificiados de estudios "
   
    vlSql = " select b.num_poliza, a.cod_tippension,  b.num_orden, b.gls_nomben || ' ' || b.gls_nomsegben || ' ' || b.gls_patben || ' ' || b.gls_matben nomben,"
    vlSql = vlSql & " c.fec_inicer, c.fec_tercer, fec_nacben, substr((months_between(sysdate,to_date(fec_nacben,'YYYYMMDD'))/12),1,2) edad,"
    vlSql = vlSql & " est_act , ind_dni, ind_dju, ind_pes, ind_bno, cod_tipo"
    vlSql = vlSql & " from pp_tmae_ben b"
    vlSql = vlSql & " left join pp_tmae_certificado c on b.num_poliza=c.num_poliza and b.num_orden=c.num_orden"
    vlSql = vlSql & " join pp_tmae_poliza a on a.num_poliza=b.num_poliza and a.num_endoso=b.num_endoso"
    vlSql = vlSql & " where b.num_poliza in (select num_poliza from pp_tmae_poliza where num_endoso=1 and fec_devsol>'20130801')"
    vlSql = vlSql & " and b.num_endoso=1"
    vlSql = vlSql & " and cod_par='30'"
    vlSql = vlSql & " and (months_between(sysdate,to_date(fec_nacben,'YYYYMMDD'))/12) >=18"
    vlSql = vlSql & " and cod_sitinv='N'"
    vlSql = vlSql & " and (cod_tipo='EST' or cod_tipo is null)"
    vlSql = vlSql & " and (fec_inicer is null or fec_inicer =(select max(fec_inicer) from pp_tmae_certificado where num_poliza=b.num_poliza and num_orden=b.num_orden) and cod_tipo='EST')"
    vlSql = vlSql & " and a.cod_tippension>='08'"
    vlSql = vlSql & " order by 1"

   Set vlRS18 = vgConexionBD.Execute(vlSql)
   Dim LNGa As Long
   LNGa = CreateFieldDefFile(vlRS18, Replace(UCase(strRpt & "Estructura\PP_Rpt_BenGarMay18Est.rpt"), ".RPT", ".TTX"), 1)
    
        
    If objRep.CargaReporte(strRpt & "", "PP_Rpt_BenGarMay18Est.rpt", "Informe de Beneficiarios que cumplieron 18 años y tienen Estudios", vlRS18, True, _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("NombreSubSistema", vgNombreSubSistema), _
                            ArrFormulas("NombreSistema", vgNombreSistema), _
                            ArrFormulas("Fecha", vlFechaTermino)) = False Then
                            
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
    End If
    
    Screen.MousePointer = 0
    'fin marco

Exit Sub
Err_flInformeCerEst:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Form_Load()
On Error GoTo Err_Cargar

    Me.Top = 0
    Me.Left = 0

    Me.Height = 1920
    Me.Width = 3690

   'Carga Combo de Tipo de Identificación Causante
    fgComboTipoIdentificacion Cmb_PenNumIdent

    Fra_CertEst.Enabled = False
    Txt_FecIniVig.Enabled = False
    
    Call flInicializaGrillaBenef
    Msf_GrillaBenef.Enabled = False
    
    Call flLmpGrilla
    Msf_Grilla.Enabled = False

    Fra_Poliza.Enabled = True
    Fra_Beneficiarios.Enabled = False
    Msf_GrillaBenef.Enabled = False
    vlTipoCertif = "SUP"
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
    Txt_FecRecep = (Msf_Grilla.Text)

    Msf_Grilla.Col = 4
    Lbl_FecIngreso = (Msf_Grilla.Text)

    Msf_Grilla.Col = 5
    Lbl_Efecto = (Msf_Grilla.Text)
    
    Msf_Grilla.Col = 6
    chk_dni.Value = IIf((Msf_Grilla.Text) = "S", 1, 0)
    
    Msf_Grilla.Col = 7
    chk_dju.Value = IIf((Msf_Grilla.Text) = "S", 1, 0)
    
    Msf_Grilla.Col = 8
    chk_pes.Value = IIf((Msf_Grilla.Text) = "S", 1, 0)
    
    Msf_Grilla.Col = 9
    chk_bno.Value = IIf((Msf_Grilla.Text) = "S", 1, 0)

    Msf_Grilla.Col = 10
    chkEst.Value = IIf((Msf_Grilla.Text) = "1", 1, 0)


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

Private Sub Msf_GrillaBenef_Click()
On Error GoTo Err_Msf_GrillaBenef_Click
    
    Msf_GrillaBenef.Col = 0
    If (Msf_GrillaBenef.Text = "") Or (Msf_GrillaBenef.row = 0) Then
        MsgBox "No existen Detalles de Beneficiarios", vbExclamation, "Información"
        Exit Sub
    Else
        Msf_GrillaBenef.Col = 2
        Lbl_CodTipoIdenBen = Trim(Msf_GrillaBenef.Text)
        Msf_GrillaBenef.Col = 3
        Lbl_NumIdenBen = Trim(Msf_GrillaBenef.Text)
        Msf_GrillaBenef.Col = 4
        Lbl_NomBen = Msf_GrillaBenef.Text
    
    
        Call flLmpGrilla
        
        Msf_GrillaBenef.Col = 0
        vlNumOrden = Msf_GrillaBenef.Text
        
        Call flCargaGrilla
        Call Mostrar_Institucion(Trim(Txt_PenPoliza), Lbl_CodTipoIdenBen, Lbl_NumIdenBen)
        Fra_Ben.Enabled = True
        Fra_CertEst.Enabled = True
        Msf_Grilla.Enabled = True
        
        'Saca la fecha Actual del Servidor
        Lbl_FecIngreso = fgBuscaFecServ
        Txt_FecIniVig.Enabled = True
        Txt_FecIniVig.SetFocus
        
        
        Txt_FecIniVig.Text = ""
        Txt_FecTerVig.Text = ""
        Txt_FecRecep.Text = ""
        Txt_Institucion.Text = ""
        chk_dni.Value = False
        chk_dju.Value = False
        chk_pes.Value = False
        chk_bno.Value = False
        
    End If

Exit Sub
Err_Msf_GrillaBenef_Click:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Mostrar_Institucion(ByVal Poliza As String, ByVal tipodoc As String, ByVal num As String)
Dim cadena As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
    
    cadena = "select DISTINCT T.GLS_ELEMENTO from PP_TMAE_POLIZA p inner join ma_tpar_tabcod t on P.COD_AFP=T.COD_ELEMENTO WHERE t.cod_tabla='AF' AND p.NUM_POLIZA='" & Poliza & "'"
    On Error GoTo mierror
        rs.Open cadena, vgConexionBD, adOpenStatic, adLockReadOnly
        If Not rs.EOF Then
            Txt_Institucion.Text = "AFP " & rs!GLS_ELEMENTO
        End If
        
        rs.Close
        Set rs = Nothing
    Exit Sub
mierror:
    MsgBox "Error al mostrar Institucion", vbInformation, "Pensiones"
End Sub



Private Sub optCertifEst_Click()
    frmCErtEstudios.Enabled = True
    vlTipoCertif = "EST"
    frmOpciones.Visible = False
    Me.Height = 8040
    Me.Width = 9810
    Me.Caption = "Mantenedor de Certificados de Estudios"
    Command1.Visible = True
End Sub

Private Sub optCertifSup_Click()
    frmCErtEstudios.Enabled = False
    vlTipoCertif = "SUP"
    frmOpciones.Visible = False
    Me.Height = 8040
    Me.Width = 9810
    Me.Caption = "Mantenedor de Certificados de Supervivencia."
    Command1.Visible = False
End Sub

Private Sub Txt_FecRecep_GotFocus()
    vlSw = False
    Txt_FecRecep.SelStart = 0
    Txt_FecRecep.SelLength = Len(Txt_FecRecep)

End Sub

Private Sub Txt_FecRecep_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And Trim(Txt_FecRecep <> "") Then
       Txt_Institucion.SetFocus
    End If

End Sub

Private Sub Txt_FecRecep_LostFocus()
    If Txt_FecRecep = "" Then
       Exit Sub
    End If
    If Not IsDate(Txt_FecRecep) Then
       Txt_FecRecep = ""
       Exit Sub
    End If
    If Txt_FecRecep <> "" Then
       Txt_FecRecep = Format(CDate(Trim(Txt_FecRecep)), "yyyymmdd")
       Txt_FecRecep = DateSerial(Mid((Txt_FecRecep), 1, 4), Mid((Txt_FecRecep), 5, 2), Mid((Txt_FecRecep), 7, 2))
    End If
End Sub

Private Sub Txt_FecTerVig_GotFocus()
    Txt_FecTerVig.SelStart = 0
    Txt_FecTerVig.SelLength = Len(Txt_FecTerVig)
End Sub

Private Sub Txt_FecTerVig_KeyPress(KeyAscii As Integer)
Dim iFecha As String

    If KeyAscii = 13 And Trim(Txt_FecTerVig > "") Then
        'se guarda en una variable lo que hay en el text de la fecha
        vlFechaTer = ""
        If IsDate(Txt_FecTerVig) Then
            Txt_FecTerVig = Format(CDate(Trim(Txt_FecTerVig)), "yyyymmdd")
            Txt_FecTerVig = DateSerial(Mid((Txt_FecTerVig), 1, 4), Mid((Txt_FecTerVig), 5, 2), Mid((Txt_FecTerVig), 7, 2))
            'valida que el texto no este en blanco
            vlFechaTer = Trim(Txt_FecTerVig)
            iFecha = Format(CDate(Trim(vlFechaTer)), "yyyymmdd")
        Else
            Txt_FecTerVig = ""
        End If

        If (Txt_FecIniVig <> "") And (Txt_FecTerVig <> "") Then
            If (Format(CDate(Txt_FecTerVig), "yyyymmdd") < Format(CDate(Txt_FecIniVig), "yyyymmdd")) Then
                MsgBox "La Fecha de Término es menor a la Fecha de Inicio de Vigencia.", vbCritical, "Dato Incorrecto"
                Exit Sub
            End If
        End If

        Txt_FecRecep.SetFocus
    End If

End Sub

Private Sub Txt_FecTerVig_LostFocus()
Dim iFecha As String
Dim vlFechaEfecto As Date

  If vlPasa = False Then
    If Txt_FecTerVig = "" Then
       Exit Sub
    End If
    If Not IsDate(Txt_FecTerVig) Then
       Txt_FecTerVig = ""
       Exit Sub
    End If
    If IsDate(Txt_FecTerVig) Then
       Txt_FecTerVig = Format(CDate(Trim(Txt_FecTerVig)), "yyyymmdd")
       Txt_FecTerVig = DateSerial(Mid((Txt_FecTerVig), 1, 4), Mid((Txt_FecTerVig), 5, 2), Mid((Txt_FecTerVig), 7, 2))
       iFecha = Format(CDate(Trim(Txt_FecTerVig)), "yyyymmdd")
       'Solo si no se trata de un registro grabado anteriormente
       If Txt_FecIniVig.Enabled Then
            vlFechaEfecto = fgFechaEfectoReliq(Txt_FecIniVig, Txt_PenPoliza, vlNumOrden, Txt_FecTerVig)
            Lbl_Efecto = vlFechaEfecto
       End If
    End If

    If (Txt_FecIniVig <> "") And (Txt_FecTerVig <> "") Then
        If (Format(CDate(Txt_FecTerVig), "yyyymmdd") < Format(CDate(Txt_FecIniVig), "yyyymmdd")) Then
            MsgBox "La Fecha de Término es menor a la Fecha de Inicio de Vigencia.", vbCritical, "Dato Incorrecto"
            Exit Sub
        End If
    End If
  End If
End Sub

Private Sub Txt_Institucion_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Inst

   If KeyAscii = 13 Then
      If Txt_Institucion <> "" Then
         Txt_Institucion = UCase(Trim(Txt_Institucion))
      End If
      Cmd_Grabar.SetFocus
   End If

Exit Sub
Err_Inst:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub

Private Sub Txt_Institucion_LostFocus()
On Error GoTo Err_Inst

   If Txt_Institucion <> "" Then
      Txt_Institucion = UCase(Trim(Txt_Institucion))
   End If

Exit Sub
Err_Inst:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Txt_PenNumIdent_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If (Trim(Txt_PenNumIdent) <> "") Then
            Txt_PenNumIdent = UCase(Trim(Txt_PenNumIdent))
        End If
        Cmd_BuscarPol.SetFocus
    End If
End Sub


Private Sub Txt_PenPoliza_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Poliza

   If KeyAscii = 13 Then
      If Trim(Txt_PenPoliza) <> "" Then
        Txt_PenPoliza = Trim(UCase(Txt_PenPoliza))
        Txt_PenPoliza = Format(Txt_PenPoliza, "0000000000")
      End If
      Cmb_PenNumIdent.SetFocus
   End If

Exit Sub
Err_Poliza:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub


Private Sub Txt_PenPoliza_LostFocus()
    Txt_PenPoliza = Trim(UCase(Txt_PenPoliza))
    Txt_PenPoliza = Format(Txt_PenPoliza, "0000000000")
End Sub

Function flValidarBen()
Dim vlFechaActual As String
On Error GoTo Err_Validar

   Screen.MousePointer = 11
   'vlFechaActual = fgBuscaFecServ
   Txt_PenPoliza = Trim(UCase(Txt_PenPoliza))
   Txt_PenNumIdent = Trim(UCase(Txt_PenNumIdent))
   
   vlCodTipoIdenBenCau = fgObtenerCodigo_TextoCompuesto(Cmb_PenNumIdent)
   vlNumIdenBenCau = Txt_PenNumIdent

   'Verificar Número de Póliza, y saca el último Endoso
   vgPalabra = ""
   vgSql = ""
   If Txt_PenPoliza <> "" And Cmb_PenNumIdent <> "" And Txt_PenNumIdent <> "" Then
      vgPalabra = "NUM_POLIZA = '" & Txt_PenPoliza & "' AND "
      vgPalabra = vgPalabra & "num_idenBEN = '" & (vlNumIdenBenCau) & "' "
      vgPalabra = vgPalabra & "AND cod_tipoidenBEN = " & (vlCodTipoIdenBenCau) & " "
   Else
     If Txt_PenPoliza <> "" Then
        vgPalabra = "NUM_POLIZA = '" & Txt_PenPoliza & "'"
     Else
        If Cmb_PenNumIdent <> "" And Txt_PenNumIdent <> "" Then
            vgPalabra = vgPalabra & "num_idenBEN = '" & (vlNumIdenBenCau) & "' "
            vgPalabra = vgPalabra & "AND cod_tipoidenBEN = " & (vlCodTipoIdenBenCau) & " "
        End If
     End If
   End If

    vlSwGrilla = True

    vgSql = ""
    
    
    vgSql = "SELECT b.num_poliza,b.num_endoso "
    vgSql = vgSql & "FROM pp_tmae_ben b "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & vgPalabra
    vgSql = vgSql & " AND b.num_endoso = "
    vgSql = vgSql & "(SELECT MAX (num_endoso) as numero FROM pp_tmae_poliza "
    vgSql = vgSql & "WHERE Num_Poliza = b.Num_Poliza) "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
        vlNumPoliza = Trim(vgRegistro!num_poliza)
        vlNumEndoso = (vgRegistro!num_endoso)
        Txt_PenPoliza = Trim(vlNumPoliza)
        Lbl_End = vlNumEndoso
        
        Call flBuscarPensionado
        
        Call fgBuscarPosicionCodigoCombo(vlCodTipoIdenBenCau, Cmb_PenNumIdent)
        'Cmb_PenNumIdent = vlCodTipoIdenBenCau
        Txt_PenNumIdent = vlNumIdenBenCau
        vlNombreCompleto = fgFormarNombreCompleto(vlGlsNomBenCau, vlGlsNomSegBenCau, vlGlsPatBenCau, vlGlsMatBenCau)
        Lbl_PenNombre = vlNombreCompleto
        
        Call flCargaGrillaBenef

        Fra_Poliza.Enabled = False
        Fra_Beneficiarios.Enabled = True
        Msf_GrillaBenef.Enabled = True
    End If
    

'    If (Lbl_End <> "") And (Lbl_PenNombre <> "") And (Txt_PenPoliza <> "") And (Txt_PenRut <> "") Then
'        flLmpGrilla
'        flCargaGrilla
'        Fra_CertEst.Enabled = True
'        Msf_Grilla.Enabled = True
'        'Saca la fecha Actual del Servidor
'         Lbl_FecIngreso = fgBuscaFecServ
'        Txt_FecIniVig.Enabled = True
'        Txt_FecIniVig.SetFocus
'    End If

    Screen.MousePointer = 0

Exit Function
Err_Validar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flBuscarPensionado()
On Error GoTo Err_flBuscarPensionado

    'Buscar los datos del Causante de la póliza
    vgSql = ""
    vgSql = "SELECT b.num_orden,b.cod_tipoidenben,b.num_idenben,"
    vgSql = vgSql & "b.gls_nomben,b.gls_nomsegben,b.gls_patben,b.gls_matben "
    vgSql = vgSql & "FROM pp_tmae_ben b "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "b.num_poliza = '" & Trim(vlNumPoliza) & "' AND "
    vgSql = vgSql & "b.num_endoso = " & vlNumEndoso & " AND "
    vgSql = vgSql & "b.cod_par = '" & clCodParCausante & "' "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
        vlNumOrdenCau = (vgRegistro!Num_Orden)
        vlCodTipoIdenBenCau = (vgRegistro!Cod_TipoIdenBen)
        vlNumIdenBenCau = (vgRegistro!Num_IdenBen)
        vlGlsNomBenCau = (vgRegistro!Gls_NomBen)
        vlGlsNomSegBenCau = IIf(IsNull(vgRegistro!Gls_NomSegBen), "", vgRegistro!Gls_NomSegBen)
        vlGlsPatBenCau = (vgRegistro!Gls_PatBen)
        vlGlsMatBenCau = IIf(IsNull(vgRegistro!Gls_MatBen), "", vgRegistro!Gls_MatBen)
    End If

Exit Function
Err_flBuscarPensionado:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

'--------------------------------------
'Permite cargar los Datos en la Grilla para desplegarlos por pantalla
'--------------------------------------
Function flCargaGrilla()
On Error GoTo Err_Carga

     vlCodTab = "FP"

     vgSql = ""
     vgSql = "SELECT c.NUM_POLIZA,c.NUM_ENDOSO,c.NUM_ORDEN,c.fec_inicer,c.fec_tercer,"
     vgSql = vgSql & " c.COD_FRECUENCIA,c.GLS_NOMINSTITUCION,c.FEC_RECCIA,c.FEC_INGCIA,"
     vgSql = vgSql & " c.FEC_EFECTO, IND_DNI, IND_DJU, IND_PES, IND_BNO, EST_ACT "
     vgSql = vgSql & " FROM pp_tmae_certificado C "
     vgSql = vgSql & " Where "
     vgSql = vgSql & " c.NUM_POLIZA = '" & Trim(Txt_PenPoliza) & "' AND "
     'CMV-20051124 I
     'No se debe utilizar el num_endoso para buscar los certificados de estudio
     'de una carga, ya que debe mostrar todos los certificados que la carga
     'haya tenido, no sólo los registros del endoso actual
     'vgSql = vgSql & " c.NUM_ENDOSO = " & Trim(vlNumEndoso) & " AND "
     'CMV-20051124 F
     vgSql = vgSql & " c.NUM_ORDEN = " & Trim(vlNumOrden) & " AND COD_TIPO='" & vlTipoCertif & "' "  '  AND COD_TIPO='SUP'
     vgSql = vgSql & " ORDER BY C.fec_inicer DESC"
     Set vlRegistro = vgConexionBD.Execute(vgSql)
     If Not vlRegistro.EOF Then
        While Not vlRegistro.EOF

              vlInicio = (vlRegistro!fec_inicer)
              vlAnno = Mid(vlInicio, 1, 4)
              vlMes = Mid(vlInicio, 5, 2)
              vlDia = Mid(vlInicio, 7, 2)
              vlInicio = DateSerial((vlAnno), (vlMes), (vlDia))

              vlTermino = (vlRegistro!FEC_TERCER)
              vlAnno = Mid(vlTermino, 1, 4)
              vlMes = Mid(vlTermino, 5, 2)
              vlDia = Mid(vlTermino, 7, 2)
              vlTermino = DateSerial((vlAnno), (vlMes), (vlDia))

              vlRecepcion = (vlRegistro!FEC_RECCIA)
              vlAnno = Mid(vlRecepcion, 1, 4)
              vlMes = Mid(vlRecepcion, 5, 2)
              vlDia = Mid(vlRecepcion, 7, 2)
              vlRecepcion = DateSerial((vlAnno), (vlMes), (vlDia))

              vlIngreso = (vlRegistro!FEC_INGCIA)
              vlAnno = Mid(vlIngreso, 1, 4)
              vlMes = Mid(vlIngreso, 5, 2)
              vlDia = Mid(vlIngreso, 7, 2)
              vlIngreso = DateSerial((vlAnno), (vlMes), (vlDia))

              vlEfecto = (vlRegistro!FEC_EFECTO)
              vlAnno = Mid(vlEfecto, 1, 4)
              vlMes = Mid(vlEfecto, 5, 2)
              vlDia = Mid(vlEfecto, 7, 2)
              vlEfecto = DateSerial((vlAnno), (vlMes), (vlDia))

              vlDNI = IIf(IsNull(vlRegistro!ind_dni), "", vlRegistro!ind_dni)
              vlDCJ = IIf(IsNull(vlRegistro!ind_dju), "", vlRegistro!ind_dju)
              vlPes = IIf(IsNull(vlRegistro!ind_pes), "", vlRegistro!ind_pes)
              vlBon = IIf(IsNull(vlRegistro!ind_bno), "", vlRegistro!ind_bno)
              vlEST = IIf(IsNull(vlRegistro!est_act), "", vlRegistro!est_act)
              Msf_Grilla.AddItem ((vlInicio) & vbTab & (vlTermino)) & vbTab & _
                                 (vlRegistro!GLS_NOMINSTITUCION) & vbTab & _
                                 (vlRecepcion) & vbTab & _
                                 (vlIngreso) & vbTab & _
                                 (vlEfecto) & vbTab & (vlDNI) & vbTab & (vlDCJ) & vbTab & (vlPes) & vbTab & (vlBon) & vbTab & vlEST
              vlRegistro.MoveNext
        Wend
     End If
     vlRegistro.Close

     Screen.MousePointer = 0

Exit Function
Err_Carga:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flLmpGrilla()
    Msf_Grilla.Clear
    Msf_Grilla.rows = 1
    Msf_Grilla.RowHeight(0) = 250
    Msf_Grilla.row = 0

    Msf_Grilla.Col = 0
    Msf_Grilla.Text = "Desde"
    Msf_Grilla.ColWidth(0) = 1500
    Msf_Grilla.ColAlignment(0) = 1

    Msf_Grilla.Col = 1
    Msf_Grilla.Text = "Hasta"
    Msf_Grilla.ColWidth(1) = 1500
    Msf_Grilla.ColAlignment(1) = 1

    Msf_Grilla.Col = 2
    Msf_Grilla.Text = "Institución"
    Msf_Grilla.ColWidth(2) = 4500

    Msf_Grilla.Col = 3
    Msf_Grilla.Text = "Fecha Recepción"
    Msf_Grilla.ColWidth(3) = 0

    Msf_Grilla.Col = 4
    Msf_Grilla.Text = "Fecha de Ingreso"
    Msf_Grilla.ColWidth(4) = 0

    Msf_Grilla.Col = 5
    Msf_Grilla.Text = "Fecha de Efecto"
    Msf_Grilla.ColWidth(5) = 0

    Msf_Grilla.Col = 6
    Msf_Grilla.Text = "SI_DNI"
    Msf_Grilla.ColWidth(6) = 0
    Msf_Grilla.ColAlignment(6) = 1

    Msf_Grilla.Col = 7
    Msf_Grilla.Text = "SI_Dec.Jurada"
    Msf_Grilla.ColWidth(7) = 0
    Msf_Grilla.ColAlignment(7) = 1
    
    Msf_Grilla.Col = 8
    Msf_Grilla.Text = "SI_Plan.Estudio"
    Msf_Grilla.ColWidth(8) = 0
    Msf_Grilla.ColAlignment(8) = 1
    
    Msf_Grilla.Col = 9
    Msf_Grilla.Text = "SI_Bol.Notas"
    Msf_Grilla.ColWidth(9) = 0
    Msf_Grilla.ColAlignment(9) = 1

    Msf_Grilla.Col = 10
    Msf_Grilla.Text = "Estado"
    Msf_Grilla.ColWidth(10) = 0
    Msf_Grilla.ColAlignment(10) = 1

End Function

Private Sub Txt_FecIniVig_GotFocus()
    vlPasa = False
    Txt_FecIniVig.SelStart = 0
    Txt_FecIniVig.SelLength = Len(Txt_FecIniVig)
End Sub

Private Sub Txt_FecIniVig_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And (Trim(Txt_FecIniVig <> "")) Then
    vlPasa = True
    If IsDate(Txt_FecIniVig) Then
        Txt_FecIniVig = Format(CDate(Trim(Txt_FecIniVig)), "yyyymmdd")
        Txt_FecIniVig = DateSerial(Mid((Txt_FecIniVig), 1, 4), Mid((Txt_FecIniVig), 5, 2), Mid((Txt_FecIniVig), 7, 2))
        flCertEstudio
    Else
        Txt_FecIniVig = ""
    End If
    If IsDate(Txt_FecTerVig) Then
        Txt_FecTerVig = Format(CDate(Trim(Txt_FecTerVig)), "yyyymmdd")
        Txt_FecTerVig = DateSerial(Mid((Txt_FecTerVig), 1, 4), Mid((Txt_FecTerVig), 5, 2), Mid((Txt_FecTerVig), 7, 2))
        flCertEstudio
    Else
        Txt_FecTerVig = ""
    End If
    If (Txt_FecIniVig <> "") And (Txt_FecTerVig <> "") Then
        If (Format(CDate(Txt_FecTerVig), "yyyymmdd") < Format(CDate(Txt_FecIniVig), "yyyymmdd")) Then
            MsgBox "La Fecha de Término es menor a la Fecha de Inicio de Vigencia.", vbCritical, "Dato Incorrecto"
            Exit Sub
        End If
    End If
End If
End Sub

Private Sub Txt_FecIniVig_LostFocus()
Dim fecha As Date
If Txt_FecIniVig <> "" Then
    If vlPasa = False Then
        If IsDate(Txt_FecIniVig) Then
            Txt_FecIniVig = Format(CDate(Trim(Txt_FecIniVig)), "yyyymmdd")
            Txt_FecIniVig = DateSerial(Mid((Txt_FecIniVig), 1, 4), Mid((Txt_FecIniVig), 5, 2), Mid((Txt_FecIniVig), 7, 2))
            flCertEstudio
        Else
            Txt_FecIniVig = ""
        End If
        If IsDate(Txt_FecTerVig) Then
            Txt_FecTerVig = Format(CDate(Trim(Txt_FecTerVig)), "yyyymmdd")
            Txt_FecTerVig = DateSerial(Mid((Txt_FecTerVig), 1, 4), Mid((Txt_FecTerVig), 5, 2), Mid((Txt_FecTerVig), 7, 2))
            flCertEstudio
        Else
            Txt_FecTerVig = ""
        End If
        If (Txt_FecIniVig <> "") And (Txt_FecTerVig <> "") Then
            If (Format(CDate(Txt_FecTerVig), "yyyymmdd") < Format(CDate(Txt_FecIniVig), "yyyymmdd")) Then
                MsgBox "La Fecha de Término es menor a la Fecha de Inicio de Vigencia.", vbCritical, "Dato Incorrecto"
                Exit Sub
            End If
        End If
    End If
    fecha = Txt_FecIniVig.Text
    If vlTipoCertif = "EST" Then
        Txt_FecTerVig.Text = DateAdd("d", -1, DateAdd("m", 6, fecha))
    Else
        Txt_FecTerVig.Text = DateAdd("d", -1, DateAdd("m", 12, fecha))
    End If
    Txt_FecTerVig = DateSerial(Mid((Txt_FecTerVig), 7, 4), Mid((Txt_FecTerVig), 4, 2) + 1, 0)
End If
End Sub

Function flValidaFecha(iFecha)
On Error GoTo Err_valfecha

      flValidaFecha = False

     'valida que la fecha este correcta
      If Trim(iFecha <> "") Then
         If Not IsDate(iFecha) Then
                MsgBox "La Fecha ingresada no es una fecha válida.", vbCritical, "Dato Incorrecto"
                Exit Function
         End If

         If (Year(iFecha) < 1900) Then
             MsgBox "La Fecha ingresada es menor a la mínima que se puede ingresar (1900).", vbCritical, "Dato Incorrecto"
             Exit Function
         End If
         flValidaFecha = True
     End If

Exit Function
Err_valfecha:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(2) & Err.Description & " ]"
    End Select
End Function

Function flCertEstudio()
On Error GoTo Err_Est

     vlFechaIni = Trim(Txt_FecIniVig)

     If (Txt_FecIniVig) <> "" Then
          vlSwVigPol = True

        If fgValidaVigenciaPoliza(Trim(Txt_PenPoliza), Trim(Txt_FecIniVig)) = False Then
           MsgBox "La Fecha Ingresada no se Encuentra dentro del Rango de Vigencia de la Póliza" & Chr(13) & _
                  "O esta No Vigente. No se Ingresara Ni Modificara Información. ", vbCritical, "Operación Cancelada"
           Screen.MousePointer = 0
           vlSwVigPol = False
           Exit Function
        End If

         If (flValidaFecha(vlFechaIni) = True) Then
            'transforma la fecha al formato yyyymmdd
             Screen.MousePointer = 11
             vlFechaIni = Format(CDate(Trim(vlFechaIni)), "yyyymmdd")
            'se valida que exista información para esa fecha en la BD
             Call flBuscaCert(vlFechaIni)
         Else
             Txt_FecRecep = ""
         End If
     Else
        MsgBox "Debe Ingresar la Fecha de Inicio de la Vigencia ", vbCritical, "Falta Información"
        Screen.MousePointer = 0
        Txt_FecIniVig.SetFocus
        Exit Function
     End If
     Screen.MousePointer = 0

Exit Function
Err_Est:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(2) & Err.Description & " ]"
    End Select
End Function

Function flBuscaCert(iFecha)
Dim vlFechaEfecto As Date
Dim fila As Integer
Dim vls As Integer
Dim vlI As Integer
Dim vlFecgrilla As String
On Error GoTo Err_buscavig

    If (Txt_FecIniVig <> "") Then
       fila = Msf_Grilla.rows - 1
       vls = 0
       For vlI = 1 To fila
           Msf_Grilla.row = vlI
           Msf_Grilla.Col = 0
           vlFecgrilla = Format(CDate(Msf_Grilla), "yyyymmdd")
           If vlFecgrilla = (iFecha) Then
              vls = 1
             ' JANM ----10/09/04
               vlPos = vlI
              Call flEncontroInf
              Exit For
           End If
       Next vlI

       If vls = 0 Then
          vlFechaEfecto = fgFechaEfectoReliq(Txt_FecIniVig, Txt_PenPoliza, vlNumOrden, Txt_FecTerVig)
          Lbl_Efecto = vlFechaEfecto
          'Call flBFechaEfecto
          'Lbl_Efecto = fgValidaFechaEfecto(Txt_FecIniVig, Txt_PenPoliza, vlNumOrden)
          Txt_FecTerVig.SetFocus
       End If
    End If

Exit Function
Err_buscavig:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flComparaFechaIngresada()
On Error GoTo Err_buscavig
Dim vlTxtInicio As String
Dim vlTxtTermino As String
Dim fila As Integer
Dim vlI As Integer
Dim vlGrInicio As String
Dim vlGrTermino As String

    If (Txt_FecIniVig <> "") Then
       vlSw = True
       vlTxtInicio = Format(CDate(Txt_FecIniVig), "yyyymmdd")
       vlTxtTermino = Format(CDate(Txt_FecTerVig), "yyyymmdd")

       fila = Msf_Grilla.rows - 1

       For vlI = 1 To fila

           Msf_Grilla.row = fila

           Msf_Grilla.Col = 0
           vlGrInicio = Format(CDate(Msf_Grilla.Text), "yyyymmdd")

           Msf_Grilla.Col = 1
           vlGrTermino = Format(CDate(Msf_Grilla.Text), "yyyymmdd")

           If (vlTxtInicio) >= (vlGrInicio) And (vlTxtInicio) <= (vlGrTermino) Then
               MsgBox "Rango de Fechas ya se encuentran en un período de Vigencia", vbCritical, "Error de Datos"
               Screen.MousePointer = 0
               vlSw = False
               Exit For
           Else
               If (vlTxtTermino) >= (vlGrInicio) And (vlTxtTermino) <= (vlGrTermino) Then
                   MsgBox "Rango de Fechas ya se encuentran en un período de Vigencia", vbCritical, "Error de Datos"
                   Screen.MousePointer = 0
                   vlSw = False
                   Exit For
               End If
           End If
        fila = fila - 1
       Next
    End If

Exit Function
Err_buscavig:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flComparaFechaExistente()
On Error GoTo Err_buscavig
Dim vlTxtInicio As String
Dim vlTxtTermino As String
Dim fila As Integer
Dim vlI As Integer
Dim vlGrInicio As String
Dim vlGrTermino As String

    If (Txt_FecIniVig <> "") Then
       vlSw = True
       vlTxtInicio = Format(CDate(Txt_FecIniVig), "yyyymmdd")
       vlTxtTermino = Format(CDate(Txt_FecTerVig), "yyyymmdd")
       fila = Msf_Grilla.rows - 1
       For vlI = 1 To fila

           Msf_Grilla.row = fila

           Msf_Grilla.Col = 0
           vlGrInicio = Format(CDate(Msf_Grilla.Text), "yyyymmdd")

           Msf_Grilla.Col = 1
           vlGrTermino = Format(CDate(Msf_Grilla.Text), "yyyymmdd")

           If (vlGrInicio) >= (vlTxtInicio) And (vlGrInicio) <= (vlTxtTermino) Then
               MsgBox "Rango de Fechas ya se encuentran en un período de Vigencia", vbCritical, "Error de Datos"
               Screen.MousePointer = 0
               vlSw = False
               Exit For
           Else
               If (vlGrTermino) >= (vlTxtInicio) And (vlGrTermino) <= (vlTxtTermino) Then
                   MsgBox "Rango de Fechas ya se encuentran en un período de Vigencia", vbCritical, "Error de Datos"
                   vlSw = False
                   Screen.MousePointer = 0
                   Exit For
               End If
           End If
        fila = fila - 1
       Next
    End If

Exit Function
Err_buscavig:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flEncontroInf()
Dim vlI As Integer

    Msf_Grilla.Col = 1
    Txt_FecTerVig = (Msf_Grilla.Text)

    Msf_Grilla.Col = 2
    Txt_Institucion = (Msf_Grilla.Text)

    Msf_Grilla.Col = 3
    Txt_FecRecep = (Msf_Grilla.Text)

    Msf_Grilla.Col = 4
    Lbl_FecIngreso = (Msf_Grilla.Text)

    Msf_Grilla.Col = 5
    Lbl_Efecto = (Msf_Grilla.Text)

    'Deshabilitar la Fecha de Inicio de Vigencia de la Póliza
    Txt_FecIniVig.Enabled = False

End Function

Function flComparaFechaMod()
On Error GoTo Err_buscavig
Dim vlTxtTermino As String
Dim vlAnt As Integer
Dim vlI As Integer
Dim vlGrInicio As String
Dim vlGrTermino As String
Dim vlTxtInicio As String

       vlSw = True
       vlTxtTermino = Format(CDate(Txt_FecTerVig), "yyyymmdd")
       vlAnt = vlPos
       vlPos = vlPos - 1
       For vlI = 1 To vlPos
           Msf_Grilla.row = vlI
           Msf_Grilla.Col = 0
           vlGrInicio = Format(CDate(Msf_Grilla.Text), "yyyymmdd")
           Msf_Grilla.Col = 1
           vlGrTermino = Format(CDate(Msf_Grilla.Text), "yyyymmdd")
           If (vlTxtTermino) >= (vlGrInicio) And _
              (vlTxtTermino) <= (vlGrTermino) Then
               MsgBox "Rango de Fechas ya se encuentran en un período de Vigencia", vbCritical, "Error de Datos"
               vlSw = False
               Screen.MousePointer = 0
               Exit For
           End If
       Next

       If vlSw = True Then
          vlAnt = vlAnt - 1
          For vlI = 1 To vlAnt
              Msf_Grilla.row = vlI
              Msf_Grilla.Col = 0
              vlGrInicio = Format(CDate(Msf_Grilla.Text), "yyyymmdd")
              Msf_Grilla.Col = 1
              vlGrTermino = Format(CDate(Msf_Grilla.Text), "yyyymmdd")
              If (vlGrInicio) >= (vlTxtInicio) And (vlGrInicio) <= (vlTxtTermino) Then
                  MsgBox "Rango de Fechas ya se encuentran en un período de Vigencia", vbCritical, "Error de Datos"
                  Screen.MousePointer = 0
                  vlSw = False
                  Exit For
              Else
                  If (vlGrTermino) >= (vlTxtInicio) And (vlGrTermino) <= (vlTxtTermino) Then
                      MsgBox "Rango de Fechas ya se encuentran en un período de Vigencia", vbCritical, "Error de Datos"
                      vlSw = False
                      Screen.MousePointer = 0
                      Exit For
                  End If
              End If
          Next
       End If

Exit Function
Err_buscavig:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flLmpCert()
On Error GoTo Err_lmp

    Txt_FecIniVig = ""
    Txt_FecTerVig = ""
    Txt_FecRecep = ""
    Txt_Institucion = ""
    Lbl_FecIngreso = ""
    Lbl_Efecto = ""

Exit Function
Err_lmp:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Sub flImpresion()
Dim vlArchivo As String
Dim vlFecNac As String
Dim vlRut As String

Err.Clear
On Error GoTo Errores1

   Screen.MousePointer = 11

   vlArchivo = strRpt & "PP_Rpt_AntCertificado.rpt"   '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
       MsgBox "Archivo de Reporte de Listados no se encuentra en el directorio de la Aplicación.", 16, "Archivo no encontrado"
       Screen.MousePointer = 0
       Exit Sub
   End If

   If Txt_PenPoliza = "" Then
      Screen.MousePointer = 0
      Exit Sub
   End If

   vgQuery = ""
   vgQuery = "{pp_tmae_certificado.NUM_POLIZA} = '" & Trim(Txt_PenPoliza) & "' and "
   'vgQuery = vgQuery & " {pp_tmae_certificado.NUM_ENDOSO} = " & vlNumEndoso & " AND "
   vgQuery = vgQuery & " {pp_tmae_certificado.NUM_ORDEN} = " & vlNumOrden & " "
   'I---- ABV 23/08/2004 ---
   '   vgQuery = vgQuery & " AND {PP_TMAE_BEN.NUM_ENDOSO} = " & vlNumEndoso & " "
   'F---- ABV 23/08/2004 ---

   Rpt_CertEstudio.Reset
   Rpt_CertEstudio.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
   'Rpt_General.DataFiles(0) = vgRutaBasedeDatos       ' o App.Path & "\Nestle.mdb"
   'Rpt_General.Connect = "ODBC;DATABASE= " & vgNombreBaseDatos & ";DSN=" & vgDsn
   Rpt_CertEstudio.Connect = vgRutaDataBase
'   Rpt_General.SelectionFormula = ""
   Rpt_CertEstudio.SelectionFormula = vgQuery
   Rpt_CertEstudio.Formulas(0) = ""
   Rpt_CertEstudio.Formulas(1) = ""
   Rpt_CertEstudio.Formulas(2) = ""

   Rpt_CertEstudio.Formulas(3) = ""
   Rpt_CertEstudio.Formulas(4) = ""
   Rpt_CertEstudio.Formulas(5) = ""
   Rpt_CertEstudio.Formulas(6) = ""
   Rpt_CertEstudio.Formulas(7) = ""
   Rpt_CertEstudio.Formulas(8) = ""

   Rpt_CertEstudio.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
   Rpt_CertEstudio.Formulas(1) = "NombreSistema= '" & vgNombreSistema & "'"
   Rpt_CertEstudio.Formulas(2) = "NombreSubSistema= '" & vgNombreSubSistema & "'"

   'vlRut = Trim(Txt_PenRut) + "-" + Trim(Txt_PenDigito)

   Rpt_CertEstudio.Formulas(3) = "Poliza = '" & Trim(Txt_PenPoliza) & "'"
   Rpt_CertEstudio.Formulas(4) = "Endoso = '" & Trim(Lbl_End) & "'"
   Rpt_CertEstudio.Formulas(5) = "CodTipoIden = '" & Lbl_CodTipoIdenBen & "'"
   Rpt_CertEstudio.Formulas(6) = "NumIden = '" & Lbl_NumIdenBen & "'"
   Rpt_CertEstudio.Formulas(7) = "Nombre_bene = '" & Trim(Lbl_NomBen) & "'"
   'Rpt_CertEstudio.Formulas(8) = "Fec_Nac = '" & (vlFechaNac) & "'"

   Rpt_CertEstudio.Destination = crptToWindow
   Rpt_CertEstudio.WindowState = crptMaximized
   Rpt_CertEstudio.WindowTitle = "Informe Certificado de Supervivencia"
   Rpt_CertEstudio.Action = 1

   Screen.MousePointer = 0

Exit Sub
Errores1:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " " & "Numero de Error : " & Err.Number & Err.HelpFile & Err.HelpContext, vbExclamation, "Mensaje de error ..."
        Screen.MousePointer = 0
    End If
End Sub

Function flBuscaFechaEfecto()
'On Error GoTo Err_Buscar
'
'    'Verifica Último Periodo
'     vgSql = ""
'     vgSql = "SELECT NUM_PERPAGO,COD_ESTADOPRI,COD_ESTADOREG,"
'     vgSql = vgSql & "FEC_PRIPAGO,FEC_PAGOPROXREG,FEC_PAGOREG FROM PP_TMAE_PROPAGOPEN ORDER BY num_perpago DESC"
'     Set vlRegistro1 = vgConexionBD.Execute(vgSql)
'     If Not vlRegistro1.EOF Then
'           'Verifica si es Primer Pago o Pago Régimen
'            vgSql = ""
'            vgSql = "SELECT NUM_POLIZA,NUM_ENDOSO,NUM_ORDEN FROM PP_TMAE_LIQPAGOPENPRO"
'            vgSql = vgSql & " Where "
'            vgSql = vgSql & " NUM_POLIZA = '" & Trim(Txt_PenPoliza) & "' AND "
'            vgSql = vgSql & " NUM_ENDOSO = " & vlNumEndoso & " AND "
'            vgSql = vgSql & " NUM_ORDEN = " & vlNumOrden & " "
'            Set vlRegistro = vgConexionBD.Execute(vgSql)
'            If Not vlRegistro.EOF Then
'                  'Pago Régimen
'                   If (vlRegistro1!COD_ESTADOREG) = "A" Or (vlRegistro1!COD_ESTADOREG) = "P" Then
'                       vlPagoReg = (vlRegistro1!FEC_PAGOREG)
'                       vlAnno = Mid(vlPagoReg, 1, 4)
'                       vlMes = Mid(vlPagoReg, 5, 2)
'                       vldia = Mid(vlPagoReg, 7, 2)
'                       Lbl_Efecto = DateSerial(vlAnno, vlMes, vldia)
'                   Else
'                       If (vlRegistro1!COD_ESTADOREG) = "C" Then
'                           vlPagoProxReg = (vlRegistro1!FEC_PAGOPROXREG)
'                           vlAnno = Mid(vlPagoProxReg, 1, 4)
'                           vlMes = Mid(vlPagoProxReg, 5, 2)
'                           vldia = Mid(vlPagoProxReg, 7, 2)
'                           Lbl_Efecto = DateSerial(vlAnno, vlMes, vldia)
'                       End If
'                   End If
'            Else
'                  'Primer Pago
'                   If (vlRegistro1!COD_ESTADOPRI) = "A" Or (vlRegistro1!COD_ESTADOPRI) = "P" Then
'                       vlPagoPri = (vlRegistro1!FEC_PRIPAGO)
'                       vlAnno = Mid(vlPagoPri, 1, 4)
'                       vlMes = Mid(vlPagoPri, 5, 2)
'                       vldia = Mid(vlPagoPri, 7, 2)
'                       Lbl_Efecto = DateSerial(vlAnno, vlMes, vldia)
'                   Else
'                       If (vlRegistro1!COD_ESTADOPRI) = "C" Then
'                           vlPagoProxPri = (vlRegistro1!FEC_PAGOPROXREG)
'                           vlAnno = Mid(vlPagoProxPri, 1, 4)
'                           vlMes = Mid(vlPagoProxPri, 5, 2)
'                           vldia = Mid(vlPagoProxPri, 7, 2)
'                           Lbl_Efecto = DateSerial(vlAnno, vlMes, vldia)
'                       End If
'                   End If
'            End If
'     End If
'Exit Function
'Err_Buscar:
'    If Err.Number <> 0 Then
'        MsgBox Err.Description & " " & "Numero de Error : " & Err.Number & Err.HelpFile & Err.HelpContext, vbExclamation, "Mensaje de error ..."
'        Screen.MousePointer = 0
'    End If
End Function

Function flRecibe(NPoliza, NCodTipoIden, NNumIden, NEndoso)
On Error GoTo Err_Buscar

    Txt_PenPoliza = NPoliza
    Call fgBuscarPosicionCodigoCombo(NCodTipoIden, Cmb_PenNumIdent)
    Txt_PenNumIdent = NNumIden
    Lbl_End = NEndoso
    
    Cmd_BuscarPol_Click

Exit Function
Err_Buscar:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " " & "Numero de Error : " & Err.Number & Err.HelpFile & Err.HelpContext, vbExclamation, "Mensaje de error ..."
        Screen.MousePointer = 0
    End If
End Function

Function flLimiteEdad(vlIniVig, Mto_Edad) As Boolean
'On Error GoTo Err_Lim
'
'    vgSql = ""
'    flLimiteEdad = False
'    vgSql = "SELECT FEC_INIVIG,FEC_TERVIG,MTO_ELEMENTO FROM MA_TPAR_TABCODVIG WHERE COD_TABLA = 'LI' AND"
'    vgSql = vgSql & " COD_ELEMENTO = 'L24'"
'    Set vlRegistro1 = vgConexionBD.Execute(vgSql)
'    If Not vlRegistro1.EOF Then
'       While Not vlRegistro1.EOF
'          If vlIniVig >= (vlRegistro1!fec_inivig) And _
'             vlIniVig <= (vlRegistro1!fec_tervig) Then
'             flLimiteEdad = True
'             Mto_Edad = (vlRegistro1!MTO_ELEMENTO)
'             Exit Function
'          End If
'          vlRegistro1.MoveNext
'       Wend
'    End If
'    vlRegistro1.Close
'
'Exit Function
'Err_Lim:
'    If Err.Number <> 0 Then
'        MsgBox Err.Description & " " & "Numero de Error : " & Err.Number & Err.HelpFile & Err.HelpContext, vbExclamation, "Mensaje de error ..."
'        Screen.MousePointer = 0
'    End If
End Function

Function flBuscaFecServ()
'On Error GoTo Err_FecSer
'
'    If vgTipoBase = "ORACLE" Then
'       vgSql = ""
'       vgSql = "SELECT SYSDATE AS FEC_ACTUAL FROM MA_TCOD_GENERAL"
'       Set vlRegistro2 = vgConexionBD.Execute(vgSql)
'       If Not vlRegistro2.EOF Then
'          Lbl_FecIngreso = Mid((vlRegistro2!FEC_ACTUAL), 1, 10)
'       End If
'    Else
'      If vgTipoBase = "SQL" Then
'         vgSql = ""
'         vgSql = "SELECT GETDATE()AS FEC_ACTUAL FROM MA_TCOD_GENERAL"
'         Set vlRegistro2 = vgConexionBD.Execute(vgSql)
'         If Not vlRegistro2.EOF Then
'            Lbl_FecIngreso = Mid((vlRegistro2!FEC_ACTUAL), 1, 10)
'         End If
'      End If
'    End If
'
'Exit Function
'Err_FecSer:
'    If Err.Number <> 0 Then
'        MsgBox Err.Description & " " & "Numero de Error : " & Err.Number & Err.HelpFile & Err.HelpContext, vbExclamation, "Mensaje de error ..."
'        Screen.MousePointer = 0
'    End If
End Function

Function flInicializaGrillaBenef()
On Error GoTo Err_flInicializaGrillaBenef

    Msf_GrillaBenef.Clear
    Msf_GrillaBenef.Cols = 6
    Msf_GrillaBenef.rows = 1

    Msf_GrillaBenef.row = 0

    Msf_GrillaBenef.Col = 0
    Msf_GrillaBenef.Text = "NºOrden"
    Msf_GrillaBenef.ColWidth(0) = 800
    Msf_GrillaBenef.ColAlignment(0) = 3

    Msf_GrillaBenef.Col = 1
    Msf_GrillaBenef.Text = "Parentesco"
    Msf_GrillaBenef.ColWidth(1) = 900

    Msf_GrillaBenef.Col = 2
    Msf_GrillaBenef.Text = "Tipo Ident."
    Msf_GrillaBenef.ColWidth(2) = 1200

    Msf_GrillaBenef.Col = 3
    Msf_GrillaBenef.Text = "N° Ident."
    Msf_GrillaBenef.ColWidth(3) = 1200

    Msf_GrillaBenef.Col = 4
    Msf_GrillaBenef.Text = "Nombre"
    Msf_GrillaBenef.ColWidth(4) = 4500

    Msf_GrillaBenef.Col = 5
    Msf_GrillaBenef.Text = "Fecha Nac."
    Msf_GrillaBenef.ColWidth(5) = 900

Exit Function
Err_flInicializaGrillaBenef:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flCargaGrillaBenef()
On Error GoTo Err_flCargaGrillaBenef
Dim vlFechaNac As String
Dim vlFechaMat As String

    Call flInicializaGrillaBenef
    Fra_Beneficiarios.Enabled = True
    Msf_GrillaBenef.Enabled = True

    'Beneficiarios
    vgSql = ""
    vgSql = "SELECT "
    vgSql = vgSql & "b.num_orden,b.cod_par,b.cod_tipoidenben,b.num_idenben,b.gls_nomben, "
    vgSql = vgSql & "b.gls_patben,b.gls_matben,b.fec_nacben "
    vgSql = vgSql & ",b.gls_nomsegben "
    vgSql = vgSql & "FROM pp_tmae_ben b WHERE "
    vgSql = vgSql & "b.num_poliza = '" & Trim(Txt_PenPoliza) & "' AND "
    vgSql = vgSql & "b.num_endoso = " & vlNumEndoso & " "
    If vlTipoCertif = "EST" Then
        vgSql = vgSql & "AND b.cod_par BETWEEN 30 AND 35 "
    End If
    'vgSql = vgSql & "AND b.cod_par BETWEEN 30 AND 35 "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    While Not vgRegistro.EOF
        'vlCodPar = " " & Trim(vgRegistro!Cod_Par) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_Par, Trim(vgRegistro!Cod_Par)))
        vlCodPar = Trim(vgRegistro!Cod_Par)
        vlFechaNac = (vgRegistro!Fec_NacBen)
        vlFechaNac = DateSerial(Mid((vlFechaNac), 1, 4), Mid((vlFechaNac), 5, 2), Mid((vlFechaNac), 7, 2))
        vlCodTipoIdenBenCau = vgRegistro!Cod_TipoIdenBen
        vlNomTipoIdenBenCau = fgBuscarNombreTipoIden(vlCodTipoIdenBenCau, False)
        vlNumIdenBenCau = vgRegistro!Num_IdenBen
        vlGlsNomBenCau = vgRegistro!Gls_NomBen
        vlGlsNomSegBenCau = IIf(IsNull(vgRegistro!Gls_NomSegBen), "", vgRegistro!Gls_NomSegBen)
        vlGlsPatBenCau = vgRegistro!Gls_PatBen
        vlGlsMatBenCau = IIf(IsNull(vgRegistro!Gls_MatBen), "", vgRegistro!Gls_MatBen)

        vlNombreCompleto = fgFormarNombreCompleto(vlGlsNomBenCau, vlGlsNomSegBenCau, vlGlsPatBenCau, vlGlsMatBenCau)
        
        Msf_GrillaBenef.AddItem vgRegistro!Num_Orden & vbTab & _
                           Trim(vlCodPar) & vbTab & _
                           " " & vlCodTipoIdenBenCau + " - " + vlNomTipoIdenBenCau & vbTab & _
                           vlNumIdenBenCau & vbTab & _
                           vlNombreCompleto & vbTab & _
                            Trim(vlFechaNac)

        vgRegistro.MoveNext
    Wend

Exit Function
Err_flCargaGrillaBenef:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fgFormarNombreCompleto(iNombre As String, iNombreSeg As String, iPaterno As String, iMaterno As String) As String

fgFormarNombreCompleto = ""

If (iNombre = "") Then iNombre = "" Else iNombre = iNombre & " "
If (iNombreSeg = "") Then iNombreSeg = "" Else iNombreSeg = iNombreSeg & " "
If (iPaterno = "") Then iPaterno = "" Else iPaterno = iPaterno & " "
If (iMaterno = "") Then iMaterno = "" Else iMaterno = iMaterno & " "

fgFormarNombreCompleto = Trim(iNombre & iNombreSeg & iPaterno & iMaterno)
End Function

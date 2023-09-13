VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Frm_ConsultaHistBen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Datos Históricos"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   6645
   Begin VB.Frame Fra_Poliza 
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
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6615
      Begin VB.Label Lbl_DgvBen 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   3480
         TabIndex        =   10
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Lbl_RutBen 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Top             =   240
         Width           =   2115
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   550
         Width           =   855
      End
      Begin VB.Label Lbl_NomBen 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   555
         Width           =   5175
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Identificación"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   " Beneficiario "
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
         Index           =   43
         Left            =   120
         TabIndex        =   5
         Top             =   0
         Width           =   1140
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   3230
      Width           =   6615
      Begin VB.CommandButton Btn_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5160
         Picture         =   "Frm_ConsultaHistBen.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Msf_Grilla 
      Height          =   1875
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1320
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3307
      _Version        =   393216
      BackColor       =   14745599
   End
   Begin VB.Label Lbl_Buscador 
      Caption         =   "Resultado Búsqueda"
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
      Index           =   5
      Left            =   0
      TabIndex        =   3
      Top             =   1090
      Width           =   1815
   End
End
Attribute VB_Name = "Frm_ConsultaHistBen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Variables para Direccion
Dim vlFecIniDir As String
Dim vlGlsDir As String
Dim vlComuna As String
Dim vlProvincia As String
Dim vlRegion As String

'Variables para Via de Pago
Dim vlFecIniVia As String
Dim vlViaPago As String
Dim vlBanco As String
Dim vlTipoCta As String
Dim vlNumCuenta As String
Dim vlSucursal As String

'Variables para Plan de Salud
Dim vlFecIniSalud As String
Dim vlInstitucion As String
Dim vlModalidad As String
Dim vlMtoPlan As Double
Dim vlModalidad2 As String
Dim vlMtoPlan2 As Double
Dim vlNumFun As Double

Dim vlCodTipoIden, vlNumIden As String

Const clViaPago As String * 3 = "VPG"
Const clBanco As String * 3 = "BCO"
Const clTipoCta As String * 3 = "TCT"
Const clInsSalud As String * 2 = "IS"
Const clModalidad As String * 3 = "MPS"

Function flIniGrillaDir()
On Error GoTo Err_flIniGrillaDir

    Msf_Grilla.Clear
    Msf_Grilla.Cols = 5
    Msf_Grilla.Rows = 1
    
    Msf_Grilla.Row = 0
        
    Msf_Grilla.Col = 0
    Msf_Grilla.Text = "Fecha Inicio."
    Msf_Grilla.ColWidth(0) = 1200
    Msf_Grilla.ColAlignment(0) = 3
    
    Msf_Grilla.Col = 1
    Msf_Grilla.Text = "Domicilio"
    Msf_Grilla.ColWidth(1) = 3000
    
    Msf_Grilla.Col = 2
    Msf_Grilla.Text = "Distrito"
    Msf_Grilla.ColWidth(2) = 2000
    
    Msf_Grilla.Col = 3
    Msf_Grilla.Text = "Provincia"
    Msf_Grilla.ColWidth(3) = 2000

    Msf_Grilla.Col = 4
    Msf_Grilla.Text = "Departamento"
    Msf_Grilla.ColWidth(4) = 2000
    
Exit Function
Err_flIniGrillaDir:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flIniGrillaVia()
On Error GoTo Err_flIniGrillaVia

    Msf_Grilla.Clear
    Msf_Grilla.Cols = 6
    Msf_Grilla.Rows = 1
    
    Msf_Grilla.Row = 0
        
    Msf_Grilla.Col = 0
    Msf_Grilla.Text = "Fecha Inicio."
    Msf_Grilla.ColWidth(0) = 1200
    Msf_Grilla.ColAlignment(0) = 3
    
    Msf_Grilla.Col = 1
    Msf_Grilla.Text = "Vía Pago"
    Msf_Grilla.ColWidth(1) = 2000
    
    Msf_Grilla.Col = 2
    Msf_Grilla.Text = "Banco"
    Msf_Grilla.ColWidth(2) = 2000
    
    Msf_Grilla.Col = 3
    Msf_Grilla.Text = "Tipo Cta."
    Msf_Grilla.ColWidth(3) = 2000
    
    Msf_Grilla.Col = 4
    Msf_Grilla.Text = "Nº.Cta."
    Msf_Grilla.ColWidth(4) = 2000
    
    Msf_Grilla.Col = 5
    Msf_Grilla.Text = "Sucursal"
    Msf_Grilla.ColWidth(5) = 2200
    
Exit Function
Err_flIniGrillaVia:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flIniGrillaSalud()
On Error GoTo Err_flIniGrillaSalud

    Msf_Grilla.Clear
    Msf_Grilla.Cols = 4
    Msf_Grilla.Rows = 1
    
    Msf_Grilla.Row = 0
        
    Msf_Grilla.Col = 0
    Msf_Grilla.Text = "Fecha Inicio."
    Msf_Grilla.ColWidth(0) = 1200
    Msf_Grilla.ColAlignment(0) = 3
    
    Msf_Grilla.Col = 1
    Msf_Grilla.Text = "Inst. Salud"
    Msf_Grilla.ColWidth(1) = 2800
    
    Msf_Grilla.Col = 2
    Msf_Grilla.Text = "Modalidad"
    Msf_Grilla.ColWidth(2) = 2000
    
    Msf_Grilla.Col = 3
    Msf_Grilla.Text = "Mto. Plan"
    Msf_Grilla.ColWidth(3) = 1500
    
Exit Function
Err_flIniGrillaSalud:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flCargaGrillaDir(iNumPoliza As String, iNumOrden As Integer)
On Error GoTo Err_flCargaGrillaDir

    Call flIniGrillaDir
    vgSql = ""
    vgSql = "SELECT d.fec_inivig,d.gls_dirben,d.cod_direccion, "
    vgSql = vgSql & "c.cod_comuna,c.gls_comuna, "
    vgSql = vgSql & "p.cod_provincia,p.gls_provincia, "
    vgSql = vgSql & "r.cod_region,r.gls_region "
    vgSql = vgSql & "FROM pp_this_bendir d, ma_tpar_comuna c, "
    vgSql = vgSql & "ma_tpar_provincia p, ma_tpar_region r "
    vgSql = vgSql & "WHERE d.num_poliza = '" & Trim(iNumPoliza) & "' AND "
    vgSql = vgSql & "d.num_orden = " & iNumOrden & " AND "
    vgSql = vgSql & "d.cod_direccion = c.cod_direccion AND "
    vgSql = vgSql & "c.cod_provincia = p.cod_provincia AND "
    vgSql = vgSql & "c.cod_region = p.cod_region AND "
    vgSql = vgSql & "c.cod_region = r.cod_region "
    vgSql = vgSql & "ORDER by d.fec_inivig ASC "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
       While Not vgRs.EOF
            Msf_Grilla.AddItem (DateSerial(Mid((vgRs!fec_inivig), 1, 4), Mid((vgRs!fec_inivig), 5, 2), Mid((vgRs!fec_inivig), 7, 2))) & vbTab _
            & (Trim(vgRs!Gls_DirBen)) & vbTab _
            & (Trim(vgRs!cod_comuna)) & " - " & (Trim(vgRs!gls_comuna)) & vbTab _
            & (Trim(vgRs!cod_provincia)) & " - " & (Trim(vgRs!gls_provincia)) & vbTab _
            & (Trim(vgRs!cod_region)) & " - " & (Trim(vgRs!gls_region))
             vgRs.MoveNext
       Wend
    End If
Exit Function
Err_flCargaGrillaDir:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flCargaGrillaVia(iNumPoliza As String, iNumOrden As Integer)
On Error GoTo Err_flCargaGrillaVia

    Call flIniGrillaVia
    vgSql = ""
    vgSql = "SELECT v.fec_inivig,v.cod_viapago,v.cod_banco, "
    vgSql = vgSql & "v.cod_tipcuenta,v.num_cuenta,v.cod_sucursal, "
    vgSql = vgSql & "p.gls_elemento as gls_viapago, "
    vgSql = vgSql & "b.gls_elemento as gls_banco, "
    vgSql = vgSql & "t.gls_elemento as gls_tipocta, "
    vgSql = vgSql & "s.gls_sucursal "
    vgSql = vgSql & "FROM pp_this_benviapago v, ma_tpar_tabcod p, "
    vgSql = vgSql & "ma_tpar_tabcod b,ma_tpar_tabcod t,ma_tpar_sucursal s "
    vgSql = vgSql & "WHERE v.num_poliza = '" & Trim(iNumPoliza) & "' AND "
    vgSql = vgSql & "v.num_orden = " & iNumOrden & " AND "
    vgSql = vgSql & "p.cod_tabla = '" & clViaPago & "' AND "
    vgSql = vgSql & "b.cod_tabla = '" & clBanco & "' AND "
    vgSql = vgSql & "t.cod_tabla = '" & clTipoCta & "' AND "
    vgSql = vgSql & "v.cod_viapago = p.cod_elemento AND "
    vgSql = vgSql & "v.cod_banco = b.cod_elemento AND "
    vgSql = vgSql & "v.cod_tipcuenta = t.cod_elemento AND "
    vgSql = vgSql & "v.cod_sucursal = s.cod_sucursal "
    vgSql = vgSql & "ORDER by v.fec_inivig ASC "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
       While Not vgRs.EOF
            Msf_Grilla.AddItem (DateSerial(Mid((vgRs!fec_inivig), 1, 4), Mid((vgRs!fec_inivig), 5, 2), Mid((vgRs!fec_inivig), 7, 2))) & vbTab _
            & (Trim(vgRs!Cod_ViaPago)) & " - " & (Trim(vgRs!Gls_ViaPago)) & vbTab _
            & (Trim(vgRs!Cod_Banco)) & " - " & (Trim(vgRs!gls_banco)) & vbTab _
            & (Trim(vgRs!Cod_TipCuenta)) & " - " & (Trim(vgRs!gls_tipocta)) & vbTab _
            & (vgRs!Num_Cuenta) & vbTab _
            & (Trim(vgRs!Cod_Sucursal)) & " - " & (Trim(vgRs!gls_sucursal))
             vgRs.MoveNext
       Wend
    End If

Exit Function
Err_flCargaGrillaVia:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flCargaGrillaSalud(iNumPoliza As String, iNumOrden As Integer)
On Error GoTo Err_flCargaGrillaSalud

    Call flIniGrillaSalud
    vgSql = ""
    vgSql = "SELECT s.fec_inivig,s.cod_inssalud,s.cod_modsalud,s.mto_plansalud, "
    vgSql = vgSql & "i.gls_elemento as gls_inssalud, "
    vgSql = vgSql & "m.gls_elemento as gls_modalidad "
    vgSql = vgSql & "FROM pp_this_bensalud s, ma_tpar_tabcod i,  "
    vgSql = vgSql & "ma_tpar_tabcod m "
    vgSql = vgSql & "WHERE s.num_poliza = '" & Trim(iNumPoliza) & "' AND "
    vgSql = vgSql & "s.num_orden = " & iNumOrden & " AND "
    vgSql = vgSql & "i.cod_tabla = '" & clInsSalud & "' AND "
    vgSql = vgSql & "m.cod_tabla = '" & clModalidad & "' AND "
    vgSql = vgSql & "s.cod_inssalud = i.cod_elemento AND "
    vgSql = vgSql & "s.cod_modsalud = m.cod_elemento "
    vgSql = vgSql & "ORDER by s.fec_inivig ASC "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
        While Not vgRs.EOF
            Msf_Grilla.AddItem (DateSerial(Mid((vgRs!fec_inivig), 1, 4), Mid((vgRs!fec_inivig), 5, 2), Mid((vgRs!fec_inivig), 7, 2))) & vbTab _
                & (Trim(vgRs!Cod_InsSalud)) & " - " & (Trim(vgRs!Gls_InsSalud)) & vbTab _
                & (Trim(vgRs!Cod_ModSalud)) & " - " & (Trim(vgRs!gls_modalidad)) & vbTab _
                & (Format(vgRs!Mto_PlanSalud, "#,#0.000"))
            vgRs.MoveNext
       Wend
    End If

Exit Function
Err_flCargaGrillaSalud:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Private Sub Btn_Salir_Click()
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

    Frm_ConsultaHistBen.Top = 0
    Frm_ConsultaHistBen.Left = 0

    Lbl_RutBen = vgRutBen
    Lbl_DgvBen = vgDgvBen
    Lbl_NomBen = vgNomBen
       
    Select Case vgGlsTipoForm
        Case "FrmDir"
           Call flCargaGrillaDir(vgNumPol, vgNumOrden)
            'Call flIniGrillaDir
        Case "FrmVia"
            Call flCargaGrillaVia(vgNumPol, vgNumOrden)
            'Call flIniGrillaVia
        Case "FrmSalud"
            Call flCargaGrillaSalud(vgNumPol, vgNumOrden)
            'Call flIniGrillaSalud
    End Select
End Sub

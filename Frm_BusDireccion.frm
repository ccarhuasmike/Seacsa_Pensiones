VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.Ocx"
Begin VB.Form Frm_BusDireccion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Búsqueda de Direcciones."
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   10200
   Begin VB.Frame Frame1 
      Caption         =   "  Búsqueda  "
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
      TabIndex        =   9
      Top             =   0
      Width           =   9975
      Begin VB.TextBox Txt_Distrito 
         Height          =   285
         Left            =   7320
         MaxLength       =   50
         TabIndex        =   5
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox Txt_CodDistrito 
         Height          =   285
         Left            =   6600
         MaxLength       =   6
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Txt_Provincia 
         Height          =   285
         Left            =   4080
         MaxLength       =   50
         TabIndex        =   3
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox Txt_CodProvincia 
         Height          =   285
         Left            =   3360
         MaxLength       =   6
         TabIndex        =   2
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Txt_Departamento 
         Height          =   285
         Left            =   840
         MaxLength       =   50
         TabIndex        =   1
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox Txt_CodDepartamento 
         Height          =   285
         Left            =   120
         MaxLength       =   6
         TabIndex        =   0
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Lbl_Buscador 
         Caption         =   "C. Dist."
         Height          =   255
         Index           =   5
         Left            =   6600
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Lbl_Buscador 
         Caption         =   "Nombre Distrito"
         Height          =   255
         Index           =   4
         Left            =   7320
         TabIndex        =   15
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Lbl_Buscador 
         Caption         =   "Nombre Provincia"
         Height          =   255
         Index           =   3
         Left            =   4080
         TabIndex        =   14
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Lbl_Buscador 
         Caption         =   "C. Prov."
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Lbl_Buscador 
         Caption         =   "Nombre Departamento"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   12
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Lbl_Buscador 
         Caption         =   "C. Depto."
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   4800
      Width           =   9975
      Begin VB.CommandButton Btn_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   3960
         Picture         =   "Frm_BusDireccion.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Btn_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5400
         Picture         =   "Frm_BusDireccion.frx":06BA
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Msf_Grilla 
      Height          =   3735
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   6588
      _Version        =   393216
      Rows            =   1
      Cols            =   7
      FixedCols       =   0
      BackColor       =   -2147483624
      Enabled         =   -1  'True
      FocusRect       =   2
      SelectionMode   =   1
   End
End
Attribute VB_Name = "Frm_BusDireccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vlNomDep As String
Dim vlNomProv As String
Dim vlNomDis As String
Dim vlCodDep As String
Dim vlCodProv As String
Dim vlCodDis As String
Dim vlDep As String
Dim vlProv As String
Dim vlDis As String
Dim vlCodDir As String

'FUNCION QUE GUARDA EL NOMBRE DEL FORMULARIO DEL QUE SE LLAMO A LA FUNCION
Function flInicio(iNomForm)
    vgNomForm = iNomForm
    Call Form_Load
End Function

Function flAsignaFormulario(vgNomForm As String, iNomDep As String, iNomProv As String, iNomDiS As String, iCodDir As String)

    If vgNomForm = "Frm_AntPensionado" Then
        Call Frm_AntPensionado.flRecibeDireccion(iNomDep, iNomProv, iNomDiS, iCodDir)
    End If
    If vgNomForm = "Frm_RetJudicial" Then
        Call Frm_RetJudicial.flRecibeDireccion(iNomDep, iNomProv, iNomDiS, iCodDir)
    End If
    If vgNomForm = "Frm_AntTutores" Then
        Call Frm_AntTutores.flRecibeDireccion(iNomDep, iNomProv, iNomDiS, iCodDir)
    End If
    If vgNomForm = "Frm_PensRegistroPagosGar" Then
        Call Frm_PensRegistroPagosGar.flRecibeDireccion(iNomDep, iNomProv, iNomDiS, iCodDir)
    End If
    If vgNomForm = "Frm_PensRegistroPagos" Then
        Call Frm_PensRegistroPagos.flRecibeDireccion(iNomDep, iNomProv, iNomDiS, iCodDir)
    End If
    If vgNomForm = "Frm_EndosoPol" Then
        Call Frm_EndosoPol.flRecibeDireccion(iNomDep, iNomProv, iNomDiS, iCodDir)
    End If
    'DEXTRE
    If vgNomForm = "Frm_EndosoHerencia" Then
        Call Frm_EndosoHerencia.flRecibeDireccion(iNomDep, iNomProv, iNomDiS, iCodDir)
    End If
    
End Function


'FUNCION QUE INICIA LA GRILLA
Function flIniciaGrilla()
On Error GoTo Err_IniGrilla
    
    Msf_Grilla.Clear
    Msf_Grilla.Rows = 1
    Msf_Grilla.Cols = 7
    
    Msf_Grilla.Row = 0
    
    Msf_Grilla.Col = 0
    Msf_Grilla.ColWidth(0) = 700
    Msf_Grilla.Text = "C.Depto."
    
    Msf_Grilla.Col = 1
    Msf_Grilla.ColWidth(1) = 2500
    Msf_Grilla.Text = "Departamento"
    
    Msf_Grilla.Col = 2
    Msf_Grilla.ColWidth(2) = 700
    Msf_Grilla.Text = "C.Prov."
    
    Msf_Grilla.Col = 3
    Msf_Grilla.ColWidth(3) = 2500
    Msf_Grilla.Text = "Provincia"
    
    Msf_Grilla.Col = 4
    Msf_Grilla.ColWidth(4) = 700
    Msf_Grilla.Text = "C.Dist."
    
    Msf_Grilla.Col = 5
    Msf_Grilla.ColWidth(5) = 2500
    Msf_Grilla.Text = "Distrito"
    
    Msf_Grilla.Col = 6
    Msf_Grilla.ColWidth(6) = 0
    Msf_Grilla.Text = "Cod.Direccion"
    
Exit Function
Err_IniGrilla:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Function

Private Sub Btn_Limpiar_Click()
On Error GoTo Err_Limp
    Screen.MousePointer = 11
    
    Txt_Departamento = ""
    Txt_Provincia = ""
    Txt_Distrito = ""
    
    Call flIniciaGrilla
    
    Screen.MousePointer = 0
    
    Txt_CodDepartamento.SetFocus
Exit Sub
Err_Limp:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Private Sub Btn_Salir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo Err_Load

    Screen.MousePointer = 11
    Frm_BusDireccion.Top = 0
    Frm_BusDireccion.Left = 0
    
    Call flIniciaGrilla
    
    Txt_Departamento = ""
    Txt_Provincia = ""
    Txt_Distrito = ""
    
    Screen.MousePointer = 0
    
Exit Sub
Err_Load:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Private Sub Msf_Grilla_DblClick()
On Error GoTo Err_Grilla

    Msf_Grilla.Col = 0
    Msf_Grilla.Row = Msf_Grilla.RowSel
    If Msf_Grilla.Text = "" Or (Msf_Grilla.Row = 0) Then
        Exit Sub
    End If

    Msf_Grilla.Col = 1
    vlNomDep = Trim(Msf_Grilla.Text)
    
    Msf_Grilla.Col = 3
    vlNomProv = Trim(Msf_Grilla.Text)
    
    Msf_Grilla.Col = 5
    vlNomDis = Trim(Msf_Grilla.Text)
    
    Msf_Grilla.Col = 6
    vlCodDir = Trim(Msf_Grilla.Text)
    
    Call flAsignaFormulario(vgNomForm, vlNomDep, vlNomProv, vlNomDis, vlCodDir)
        
    Unload Me
    
    ''Frm_AntPensionado.Txt_FonoBen.SetFocus
Exit Sub
Err_Grilla:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Private Sub Txt_CodDepartamento_KeyPress(KeyAscii As Integer)
Dim vlCodDir As String
On Error GoTo Err_CodDepartamento
If KeyAscii = 13 Then

    Screen.MousePointer = 11
    
    If (Trim(Txt_Departamento <> "") Or Trim(Txt_Provincia) <> "" Or _
        Trim(Txt_Distrito <> "") Or Trim(Txt_CodDepartamento <> "") Or _
        Trim(Txt_CodProvincia) <> "" Or Trim(Txt_CodDistrito <> "")) Then
            
        vgSql = ""
        vgSql = "SELECT r.Cod_Region,r.Gls_Region,p.Cod_Provincia,p.Gls_Provincia,"
        vgSql = vgSql & "c.Cod_Comuna,c.Gls_Comuna,c.cod_direccion  "
        vgSql = vgSql & " FROM "
        vgSql = vgSql & " MA_TPAR_COMUNA c, MA_TPAR_PROVINCIA p, MA_TPAR_REGION r "
        vgSql = vgSql & " Where "
        vgSql = vgSql & " c.cod_region = p.cod_region and"
        vgSql = vgSql & " c.cod_provincia = p.cod_provincia and"
        vgSql = vgSql & " p.cod_region = r.cod_region"
          
        If Trim(Txt_CodDepartamento <> "") Then vgSql = vgSql & " and r.Cod_Region like '" & Txt_CodDepartamento & "%' "
        If Trim(Txt_Departamento <> "") Then vgSql = vgSql & " and r.Gls_Region like '" & Txt_Departamento & "%' "
        If Trim(Txt_CodProvincia <> "") Then vgSql = vgSql & " and p.Cod_Provincia like '" & Txt_CodProvincia & "%' "
        If Trim(Txt_Provincia <> "") Then vgSql = vgSql & " and p.Gls_Provincia like '" & Txt_Provincia & "%' "
        If Trim(Txt_CodDistrito <> "") Then vgSql = vgSql & " and c.Cod_Comuna like '" & Txt_CodDistrito & "%' "
        If Trim(Txt_Distrito <> "") Then vgSql = vgSql & " and c.Gls_Comuna like '" & Txt_Distrito & "%' "
        
        'vgSql = Mid(vgSql, 1, Len(vgSql) - 4)
        vgSql = vgSql & " ORDER BY r.Cod_Region,p.Cod_Provincia,c.Cod_Comuna "
        Set vgRs = vgConexionBD.Execute(vgSql)
        Msf_Grilla.Rows = 1
        
        Do While Not vgRs.EOF
            If Not IsNull(vgRs!gls_region) Then vlNomDep = Trim(vgRs!gls_region)
            If Not IsNull(vgRs!gls_provincia) Then vlNomProv = Trim(vgRs!gls_provincia)
            If Not IsNull(vgRs!gls_comuna) Then vlNomDis = Trim(vgRs!gls_comuna)
            If Not IsNull(vgRs!cod_region) Then vlCodDep = Trim(vgRs!cod_region)
            If Not IsNull(vgRs!cod_provincia) Then vlCodProv = Trim(vgRs!cod_provincia)
            If Not IsNull(vgRs!cod_comuna) Then vlCodDis = Trim(vgRs!cod_comuna)
            If Not IsNull(vgRs!Cod_Direccion) Then vlCodDir = Trim(vgRs!Cod_Direccion)
           
            Msf_Grilla.AddItem vlCodDep & vbTab & vlNomDep & vbTab & _
                        vlCodProv & vbTab & vlNomProv & vbTab & _
                        vlCodDis & vbTab & vlNomDis & vbTab & vlCodDir
                    
            vgRs.MoveNext
        Loop
    Else
        Call flIniciaGrilla
    End If

   Txt_Departamento.SetFocus
End If
Screen.MousePointer = 0

Exit Sub
Err_CodDepartamento:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Private Sub Txt_Departamento_KeyPress(KeyAscii As Integer)
Dim vlCodDir As String
On Error GoTo Err_MatPen
If KeyAscii = 13 Then

    Txt_Departamento = UCase(Trim(Txt_Departamento))
    Screen.MousePointer = 11
    
    If (Trim(Txt_Departamento <> "") Or Trim(Txt_Provincia) <> "" Or _
        Trim(Txt_Distrito <> "") Or Trim(Txt_CodDepartamento <> "") Or _
        Trim(Txt_CodProvincia) <> "" Or Trim(Txt_CodDistrito <> "")) Then
            
        vgSql = ""
        vgSql = "SELECT r.Cod_Region,r.Gls_Region,p.Cod_Provincia,p.Gls_Provincia,"
        vgSql = vgSql & "c.Cod_Comuna,c.Gls_Comuna,c.cod_direccion  "
        vgSql = vgSql & " FROM "
        vgSql = vgSql & " MA_TPAR_COMUNA c, MA_TPAR_PROVINCIA p, MA_TPAR_REGION r "
        vgSql = vgSql & " Where "
        vgSql = vgSql & " c.cod_region = p.cod_region and"
        vgSql = vgSql & " c.cod_provincia = p.cod_provincia and"
        vgSql = vgSql & " p.cod_region = r.cod_region"
          
        If Trim(Txt_CodDepartamento <> "") Then vgSql = vgSql & " and r.Cod_Region like '" & Txt_CodDepartamento & "%' "
        If Trim(Txt_Departamento <> "") Then vgSql = vgSql & " and r.Gls_Region like '" & Txt_Departamento & "%' "
        If Trim(Txt_CodProvincia <> "") Then vgSql = vgSql & " and p.Cod_Provincia like '" & Txt_CodProvincia & "%' "
        If Trim(Txt_Provincia <> "") Then vgSql = vgSql & " and p.Gls_Provincia like '" & Txt_Provincia & "%' "
        If Trim(Txt_CodDistrito <> "") Then vgSql = vgSql & " and c.Cod_Comuna like '" & Txt_CodDistrito & "%' "
        If Trim(Txt_Distrito <> "") Then vgSql = vgSql & " and c.Gls_Comuna like '" & Txt_Distrito & "%' "
        
        'vgSql = Mid(vgSql, 1, Len(vgSql) - 4)
        vgSql = vgSql & " ORDER BY r.Cod_Region,p.Cod_Provincia,c.Cod_Comuna "
        Set vgRs = vgConexionBD.Execute(vgSql)
        Msf_Grilla.Rows = 1
        
        Do While Not vgRs.EOF
            If Not IsNull(vgRs!gls_region) Then vlNomDep = Trim(vgRs!gls_region)
            If Not IsNull(vgRs!gls_provincia) Then vlNomProv = Trim(vgRs!gls_provincia)
            If Not IsNull(vgRs!gls_comuna) Then vlNomDis = Trim(vgRs!gls_comuna)
            If Not IsNull(vgRs!cod_region) Then vlCodDep = Trim(vgRs!cod_region)
            If Not IsNull(vgRs!cod_provincia) Then vlCodProv = Trim(vgRs!cod_provincia)
            If Not IsNull(vgRs!cod_comuna) Then vlCodDis = Trim(vgRs!cod_comuna)
            If Not IsNull(vgRs!Cod_Direccion) Then vlCodDir = Trim(vgRs!Cod_Direccion)
           
            Msf_Grilla.AddItem vlCodDep & vbTab & vlNomDep & vbTab & _
                        vlCodProv & vbTab & vlNomProv & vbTab & _
                        vlCodDis & vbTab & vlNomDis & vbTab & vlCodDir
                    
            vgRs.MoveNext
        Loop
    Else
        Call flIniciaGrilla
    End If

   Txt_CodProvincia.SetFocus
End If
Screen.MousePointer = 0

Exit Sub
Err_MatPen:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Private Sub Txt_CodProvincia_KeyPress(KeyAscii As Integer)
Dim vlCodDir As String
On Error GoTo Err_CodProvincia
If KeyAscii = 13 Then

    Screen.MousePointer = 11
    
    If (Trim(Txt_Departamento <> "") Or Trim(Txt_Provincia) <> "" Or _
        Trim(Txt_Distrito <> "") Or Trim(Txt_CodDepartamento <> "") Or _
        Trim(Txt_CodProvincia) <> "" Or Trim(Txt_CodDistrito <> "")) Then
            
        vgSql = ""
        vgSql = "SELECT r.Cod_Region,r.Gls_Region,p.Cod_Provincia,p.Gls_Provincia,"
        vgSql = vgSql & "c.Cod_Comuna,c.Gls_Comuna,c.cod_direccion  "
        vgSql = vgSql & " FROM "
        vgSql = vgSql & " MA_TPAR_COMUNA c, MA_TPAR_PROVINCIA p, MA_TPAR_REGION r "
        vgSql = vgSql & " Where "
        vgSql = vgSql & " c.cod_region = p.cod_region and"
        vgSql = vgSql & " c.cod_provincia = p.cod_provincia and"
        vgSql = vgSql & " p.cod_region = r.cod_region"
          
        If Trim(Txt_CodDepartamento <> "") Then vgSql = vgSql & " and r.Cod_Region like '" & Txt_CodDepartamento & "%' "
        If Trim(Txt_Departamento <> "") Then vgSql = vgSql & " and r.Gls_Region like '" & Txt_Departamento & "%' "
        If Trim(Txt_CodProvincia <> "") Then vgSql = vgSql & " and p.Cod_Provincia like '" & Txt_CodProvincia & "%' "
        If Trim(Txt_Provincia <> "") Then vgSql = vgSql & " and p.Gls_Provincia like '" & Txt_Provincia & "%' "
        If Trim(Txt_CodDistrito <> "") Then vgSql = vgSql & " and c.Cod_Comuna like '" & Txt_CodDistrito & "%' "
        If Trim(Txt_Distrito <> "") Then vgSql = vgSql & " and c.Gls_Comuna like '" & Txt_Distrito & "%' "
        
        'vgSql = Mid(vgSql, 1, Len(vgSql) - 4)
        vgSql = vgSql & " ORDER BY r.Cod_Region,p.Cod_Provincia,c.Cod_Comuna "
        Set vgRs = vgConexionBD.Execute(vgSql)
        Msf_Grilla.Rows = 1
        
        Do While Not vgRs.EOF
            If Not IsNull(vgRs!gls_region) Then vlNomDep = Trim(vgRs!gls_region)
            If Not IsNull(vgRs!gls_provincia) Then vlNomProv = Trim(vgRs!gls_provincia)
            If Not IsNull(vgRs!gls_comuna) Then vlNomDis = Trim(vgRs!gls_comuna)
            If Not IsNull(vgRs!cod_region) Then vlCodDep = Trim(vgRs!cod_region)
            If Not IsNull(vgRs!cod_provincia) Then vlCodProv = Trim(vgRs!cod_provincia)
            If Not IsNull(vgRs!cod_comuna) Then vlCodDis = Trim(vgRs!cod_comuna)
            If Not IsNull(vgRs!Cod_Direccion) Then vlCodDir = Trim(vgRs!Cod_Direccion)
           
              Msf_Grilla.AddItem vlCodDep & vbTab & vlNomDep & vbTab & _
                        vlCodProv & vbTab & vlNomProv & vbTab & _
                        vlCodDis & vbTab & vlNomDis & vbTab & vlCodDir
                    
            vgRs.MoveNext
       Loop
    Else
        Call flIniciaGrilla
    End If
    Txt_Provincia.SetFocus
End If
Screen.MousePointer = 0
Exit Sub
Err_CodProvincia:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select

End Sub

Private Sub Txt_Provincia_KeyPress(KeyAscii As Integer)
Dim vlCodDir As String
On Error GoTo Err_NumPol
    If KeyAscii = 13 Then
    
    Txt_Provincia = UCase(Trim(Txt_Provincia))

    Screen.MousePointer = 11

   If (Trim(Txt_Departamento <> "") Or Trim(Txt_Provincia) <> "" Or _
        Trim(Txt_Distrito <> "") Or Trim(Txt_CodDepartamento <> "") Or _
        Trim(Txt_CodProvincia) <> "" Or Trim(Txt_CodDistrito <> "")) Then
    
        vgSql = ""
        vgSql = "SELECT r.Cod_Region,r.Gls_Region,p.Cod_Provincia,p.Gls_Provincia,"
        vgSql = vgSql & "c.Cod_Comuna,c.Gls_Comuna,c.cod_direccion  "
        vgSql = vgSql & " FROM MA_TPAR_COMUNA c, MA_TPAR_PROVINCIA p, MA_TPAR_REGION r "
        vgSql = vgSql & " Where c.cod_region = p.cod_region and"
        vgSql = vgSql & " c.cod_provincia = p.cod_provincia and"
        vgSql = vgSql & " p.cod_region = r.cod_region"
        
        If Trim(Txt_CodDepartamento <> "") Then vgSql = vgSql & " and r.Cod_Region like '" & Txt_CodDepartamento & "%' "
        If Trim(Txt_Departamento <> "") Then vgSql = vgSql & " and r.Gls_Region like '" & Txt_Departamento & "%' "
        If Trim(Txt_CodProvincia <> "") Then vgSql = vgSql & " and p.Cod_Provincia like '" & Txt_CodProvincia & "%' "
        If Trim(Txt_Provincia <> "") Then vgSql = vgSql & " and p.Gls_Provincia like '" & Txt_Provincia & "%' "
        If Trim(Txt_CodDistrito <> "") Then vgSql = vgSql & " and c.Cod_Comuna like '" & Txt_CodDistrito & "%' "
        If Trim(Txt_Distrito <> "") Then vgSql = vgSql & " and c.Gls_Comuna like '" & Txt_Distrito & "%' "

        Set vgRs = vgConexionBD.Execute(vgSql)
        Msf_Grilla.Rows = 1

        Do While Not vgRs.EOF
            If Not IsNull(vgRs!gls_region) Then vlNomDep = Trim(vgRs!gls_region)
            If Not IsNull(vgRs!gls_provincia) Then vlNomProv = Trim(vgRs!gls_provincia)
            If Not IsNull(vgRs!gls_comuna) Then vlNomDis = Trim(vgRs!gls_comuna)
            If Not IsNull(vgRs!cod_region) Then vlCodDep = Trim(vgRs!cod_region)
            If Not IsNull(vgRs!cod_provincia) Then vlCodProv = Trim(vgRs!cod_provincia)
            If Not IsNull(vgRs!cod_comuna) Then vlCodDis = Trim(vgRs!cod_comuna)
            If Not IsNull(vgRs!Cod_Direccion) Then vlCodDir = Trim(vgRs!Cod_Direccion)
           
              Msf_Grilla.AddItem vlCodDep & vbTab & vlNomDep & vbTab & _
                        vlCodProv & vbTab & vlNomProv & vbTab & _
                        vlCodDis & vbTab & vlNomDis & vbTab & vlCodDir

            vgRs.MoveNext
        Loop

    Else
        Call flIniciaGrilla
    End If
    Txt_CodDistrito.SetFocus
End If
Screen.MousePointer = 0
Exit Sub
Err_NumPol:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select

End Sub
Private Sub Txt_CodDistrito_KeyPress(KeyAscii As Integer)
Dim vlCodDir As String
On Error GoTo Err_MatPen
If KeyAscii = 13 Then

    Screen.MousePointer = 11
    
    If (Trim(Txt_Departamento <> "") Or Trim(Txt_Provincia) <> "" Or _
        Trim(Txt_Distrito <> "") Or Trim(Txt_CodDepartamento <> "") Or _
        Trim(Txt_CodProvincia) <> "" Or Trim(Txt_CodDistrito <> "")) Then
            
        vgSql = ""
        vgSql = "SELECT r.Cod_Region,r.Gls_Region,p.Cod_Provincia,p.Gls_Provincia,"
        vgSql = vgSql & "c.Cod_Comuna,c.Gls_Comuna,c.cod_direccion  "
        vgSql = vgSql & " FROM "
        vgSql = vgSql & " MA_TPAR_COMUNA c, MA_TPAR_PROVINCIA p, MA_TPAR_REGION r "
        vgSql = vgSql & " Where "
        vgSql = vgSql & " c.cod_region = p.cod_region and"
        vgSql = vgSql & " c.cod_provincia = p.cod_provincia and"
        vgSql = vgSql & " p.cod_region = r.cod_region"
          
        If Trim(Txt_CodDepartamento <> "") Then vgSql = vgSql & " and r.Cod_Region like '" & Txt_CodDepartamento & "%' "
        If Trim(Txt_Departamento <> "") Then vgSql = vgSql & " and r.Gls_Region like '" & Txt_Departamento & "%' "
        If Trim(Txt_CodProvincia <> "") Then vgSql = vgSql & " and p.Cod_Provincia like '" & Txt_CodProvincia & "%' "
        If Trim(Txt_Provincia <> "") Then vgSql = vgSql & " and p.Gls_Provincia like '" & Txt_Provincia & "%' "
        If Trim(Txt_CodDistrito <> "") Then vgSql = vgSql & " and c.Cod_Comuna like '" & Txt_CodDistrito & "%' "
        If Trim(Txt_Distrito <> "") Then vgSql = vgSql & " and c.Gls_Comuna like '" & Txt_Distrito & "%' "
                
        'vgSql = Mid(vgSql, 1, Len(vgSql) - 4)
        vgSql = vgSql & " ORDER BY r.Cod_Region,p.Cod_Provincia,c.Cod_Comuna "
        Set vgRs = vgConexionBD.Execute(vgSql)
        Msf_Grilla.Rows = 1
        
        Do While Not vgRs.EOF
            If Not IsNull(vgRs!gls_region) Then vlNomDep = Trim(vgRs!gls_region)
            If Not IsNull(vgRs!gls_provincia) Then vlNomProv = Trim(vgRs!gls_provincia)
            If Not IsNull(vgRs!gls_comuna) Then vlNomDis = Trim(vgRs!gls_comuna)
            If Not IsNull(vgRs!cod_region) Then vlCodDep = Trim(vgRs!cod_region)
            If Not IsNull(vgRs!cod_provincia) Then vlCodProv = Trim(vgRs!cod_provincia)
            If Not IsNull(vgRs!cod_comuna) Then vlCodDis = Trim(vgRs!cod_comuna)
            If Not IsNull(vgRs!Cod_Direccion) Then vlCodDir = Trim(vgRs!Cod_Direccion)
           
              Msf_Grilla.AddItem vlCodDep & vbTab & vlNomDep & vbTab & _
                        vlCodProv & vbTab & vlNomProv & vbTab & _
                        vlCodDis & vbTab & vlNomDis & vbTab & vlCodDir
                    
            vgRs.MoveNext
        Loop
    Else
        Call flIniciaGrilla
    End If

   Txt_Distrito.SetFocus
End If
Screen.MousePointer = 0

Exit Sub
Err_MatPen:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Private Sub Txt_Distrito_KeyPress(KeyAscii As Integer)
Dim vlCodDir As String
On Error GoTo Err_RutPen
    If KeyAscii = 13 Then
    
    Txt_Distrito = UCase(Trim(Txt_Distrito))
    Screen.MousePointer = 11
If (Trim(Txt_Departamento <> "") Or Trim(Txt_Provincia) <> "" Or _
        Trim(Txt_Distrito <> "") Or Trim(Txt_CodDepartamento <> "") Or _
        Trim(Txt_CodProvincia) <> "" Or Trim(Txt_CodDistrito <> "")) Then
    
        vgSql = ""
        vgSql = "SELECT r.Cod_Region,r.Gls_Region,p.Cod_Provincia,p.Gls_Provincia,"
        vgSql = vgSql & "c.Cod_Comuna,c.Gls_Comuna,c.cod_direccion  "
        vgSql = vgSql & " FROM MA_TPAR_COMUNA c, MA_TPAR_PROVINCIA p, MA_TPAR_REGION r "
        vgSql = vgSql & " Where c.cod_region = p.cod_region and"
        vgSql = vgSql & " c.cod_provincia = p.cod_provincia and"
        vgSql = vgSql & " p.cod_region = r.cod_region"

        If Trim(Txt_CodDepartamento <> "") Then vgSql = vgSql & " and r.Cod_Region like '" & Txt_CodDepartamento & "%' "
        If Trim(Txt_Departamento <> "") Then vgSql = vgSql & " and r.Gls_Region like '" & Txt_Departamento & "%' "
        If Trim(Txt_CodProvincia <> "") Then vgSql = vgSql & " and p.Cod_Provincia like '" & Txt_CodProvincia & "%' "
        If Trim(Txt_Provincia <> "") Then vgSql = vgSql & " and p.Gls_Provincia like '" & Txt_Provincia & "%' "
        If Trim(Txt_CodDistrito <> "") Then vgSql = vgSql & " and c.Cod_Comuna like '" & Txt_CodDistrito & "%' "
        If Trim(Txt_Distrito <> "") Then vgSql = vgSql & " and c.Gls_Comuna like '" & Txt_Distrito & "%' "

        Set vgRs = vgConexionBD.Execute(vgSql)
        Msf_Grilla.Rows = 1

        Do While Not vgRs.EOF
            If Not IsNull(vgRs!gls_region) Then vlNomDep = Trim(vgRs!gls_region)
            If Not IsNull(vgRs!gls_provincia) Then vlNomProv = Trim(vgRs!gls_provincia)
            If Not IsNull(vgRs!gls_comuna) Then vlNomDis = Trim(vgRs!gls_comuna)
            If Not IsNull(vgRs!cod_region) Then vlCodDep = Trim(vgRs!cod_region)
            If Not IsNull(vgRs!cod_provincia) Then vlCodProv = Trim(vgRs!cod_provincia)
            If Not IsNull(vgRs!cod_comuna) Then vlCodDis = Trim(vgRs!cod_comuna)
            If Not IsNull(vgRs!Cod_Direccion) Then vlCodDir = Trim(vgRs!Cod_Direccion)
           
              Msf_Grilla.AddItem vlCodDep & vbTab & vlNomDep & vbTab & _
                        vlCodProv & vbTab & vlNomProv & vbTab & _
                        vlCodDis & vbTab & vlNomDis & vbTab & vlCodDir
            vgRs.MoveNext
        Loop
    Else
        Call flIniciaGrilla
    End If
    Btn_Limpiar.SetFocus
End If
Screen.MousePointer = 0
Exit Sub
Err_RutPen:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub


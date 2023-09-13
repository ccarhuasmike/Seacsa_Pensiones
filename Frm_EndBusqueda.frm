VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Frm_EndBusqueda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Búsqueda de Pólizas / Pensionados por Endoso"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   11370
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   120
      TabIndex        =   12
      Top             =   5640
      Width           =   11175
      Begin VB.CommandButton Btn_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   4680
         Picture         =   "Frm_EndBusqueda.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Btn_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   6120
         Picture         =   "Frm_EndBusqueda.frx":06BA
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      Begin VB.TextBox Txt_TipIdent 
         Height          =   285
         Left            =   2400
         MaxLength       =   3
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Txt_NumPol 
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Txt_Cuspp 
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Txt_NumIdent 
         Height          =   285
         Left            =   3480
         MaxLength       =   3
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Txt_NumEndoso 
         Height          =   285
         Left            =   5160
         MaxLength       =   50
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Txt_PatPen 
         Height          =   285
         Left            =   5880
         MaxLength       =   20
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Txt_MatPen 
         Height          =   285
         Left            =   7560
         MaxLength       =   20
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Txt_NomPen 
         Height          =   285
         Left            =   9120
         MaxLength       =   25
         TabIndex        =   8
         Top             =   240
         Width           =   1935
      End
      Begin MSFlexGridLib.MSFlexGrid Msf_Grilla 
         Height          =   4935
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   8705
         _Version        =   393216
         Rows            =   1
         Cols            =   8
         FixedCols       =   0
         BackColor       =   -2147483624
         Enabled         =   -1  'True
         FocusRect       =   2
         SelectionMode   =   1
         FormatString    =   $"Frm_EndBusqueda.frx":07B4
      End
   End
End
Attribute VB_Name = "Frm_EndBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vlCodTipoIden As Long, vlNumIden As String

'FUNCION QUE GUARDA EL NOMBRE DEL FORMULARIO DEL QUE SE LLAMO A LA FUNCION
Function flInicio(iNomForm)
    vgNomForm = iNomForm
    Call Form_Load
End Function

Function flAsignaFormulario(vgNomForm, vlNumPol, vlCodTipoIden, vlNumIden, vlNumEnd)
    If vgNomForm = "Frm_ConsultaEndoso" Then
        Call Frm_ConsultaEndoso.flRecibe(vlNumPol, vlCodTipoIden, vlNumIden, vlNumEnd)
    End If
End Function

'FUNCION QUE INICIA LA GRILLA
Function flIniciaGrilla()
On Error GoTo Err_IniGrilla
    
    Msf_Grilla.Clear
    Msf_Grilla.Rows = 1
    Msf_Grilla.Cols = 9
    
    Msf_Grilla.Row = 0
    
    Msf_Grilla.Col = 0
    Msf_Grilla.ColWidth(0) = 1000
    Msf_Grilla.Text = "Póliza"

    Msf_Grilla.Col = 1
    Msf_Grilla.ColWidth(1) = 1250
    Msf_Grilla.Text = "CUSPP"
    
    Msf_Grilla.Col = 2
    Msf_Grilla.ColWidth(2) = 1100
    Msf_Grilla.Text = "Tipo Ident."
    
    Msf_Grilla.Col = 3
    Msf_Grilla.ColWidth(3) = 1700
    Msf_Grilla.Text = "Nº. Ident. Pensionado"
    
    Msf_Grilla.Col = 4
    Msf_Grilla.ColWidth(4) = 700
    Msf_Grilla.Text = "NºEnd."
      
    Msf_Grilla.Col = 5
    Msf_Grilla.ColWidth(5) = 1700
    Msf_Grilla.Text = "Apellido Paterno"
    
    Msf_Grilla.Col = 6
    Msf_Grilla.ColWidth(6) = 1600
    Msf_Grilla.Text = "Apellido Materno"
    
    Msf_Grilla.Col = 7
    Msf_Grilla.ColWidth(7) = 1600
    Msf_Grilla.Text = "Primer Nombre"
    
    Msf_Grilla.Col = 8
    Msf_Grilla.ColWidth(8) = 0
    
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
    
    Txt_NumPol = ""
    Txt_RutPen = ""
    Txt_NumOrd = ""
    Txt_Par = ""
    Txt_PatPen = ""
    Txt_MatPen = ""
    Txt_NomPen = ""
    Txt_NumEndoso = ""
    Call flIniciaGrilla
    Txt_NumPol.SetFocus
    Screen.MousePointer = 0
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
    Frm_EndBusqueda.Top = 0
    Frm_EndBusqueda.Left = 0
    Call flIniciaGrilla
    Txt_NumPol = ""
    Txt_Cuspp = ""
    Txt_TipIdent = ""
    Txt_NumIdent = ""
    Txt_NumEndoso = ""
    Txt_PatBen = ""
    Txt_MatBen = ""
    Txt_NomPen = ""
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

    Msf_Grilla.Col = 0
    vlNumPol = Msf_Grilla.Text
   
    Msf_Grilla.Col = 2
    If (fgObtenerCod_Identificacion(Msf_Grilla.Text, vlCodTipoIden) = False) Then
        MsgBox "Tipo de Identificación no se encuentra registrada.", vbCritical, "Error de Dato"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    Msf_Grilla.Col = 3
    vlNumIden = Msf_Grilla.Text
    
    Msf_Grilla.Col = 4
    vlNumEnd = Msf_Grilla.Text
    
    Call flAsignaFormulario(vgNomForm, vlNumPol, vlCodTipoIden, vlNumIden, vlNumEnd)
    Unload Me
    
Exit Sub
Err_Grilla:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select

End Sub

'-------------------------- KeyPress ------------------------------------
Private Sub Txt_NumPol_KeyPress(KeyAscii As Integer)
    Screen.MousePointer = 11
    If KeyAscii = 13 Then
        If (Trim(Txt_NumPol) <> "" Or Trim(Txt_Cuspp <> "") Or _
            Trim(Txt_TipIdent <> "") Or Trim(Txt_NumIdent <> "") Or _
            Trim(Txt_NumEndoso <> "") Or Trim(Txt_PatPen <> "") Or _
            Trim(Txt_MatPen <> "") Or Trim(Txt_NomPen <> "")) Then
            Call flLlenaGrilla
        Else
            Call flIniciaGrilla
        End If
        Txt_Cuspp.SetFocus
    End If
    Screen.MousePointer = 0
End Sub
Private Sub Txt_Cuspp_KeyPress(KeyAscii As Integer)
    Screen.MousePointer = 11
    If KeyAscii = 13 Then
        If (Trim(Txt_NumPol) <> "" Or Trim(Txt_Cuspp <> "") Or _
            Trim(Txt_TipIdent <> "") Or Trim(Txt_NumIdent <> "") Or _
            Trim(Txt_NumEndoso <> "") Or Trim(Txt_PatPen <> "") Or _
            Trim(Txt_MatPen <> "") Or Trim(Txt_NomPen <> "")) Then
            Call flLlenaGrilla
        Else
            Call flIniciaGrilla
        End If
        Txt_TipIdent.SetFocus
    End If
    Screen.MousePointer = 0
End Sub
Private Sub Txt_TipIdent_KeyPress(KeyAscii As Integer)
    Screen.MousePointer = 11
    If KeyAscii = 13 Then
        If (Trim(Txt_NumPol) <> "" Or Trim(Txt_Cuspp <> "") Or _
            Trim(Txt_TipIdent <> "") Or Trim(Txt_NumIdent <> "") Or _
            Trim(Txt_NumEndoso <> "") Or Trim(Txt_PatPen <> "") Or _
            Trim(Txt_MatPen <> "") Or Trim(Txt_NomPen <> "")) Then
            Call flLlenaGrilla
        Else
            Call flIniciaGrilla
        End If
        Txt_NumIdent.SetFocus
    End If
    Screen.MousePointer = 0
End Sub
Private Sub Txt_NumIdent_KeyPress(KeyAscii As Integer)
    Screen.MousePointer = 11
    If KeyAscii = 13 Then
        If (Trim(Txt_NumPol) <> "" Or Trim(Txt_Cuspp <> "") Or _
            Trim(Txt_TipIdent <> "") Or Trim(Txt_NumIdent <> "") Or _
            Trim(Txt_NumEndoso <> "") Or Trim(Txt_PatPen <> "") Or _
            Trim(Txt_MatPen <> "") Or Trim(Txt_NomPen <> "")) Then
            Call flLlenaGrilla
        Else
            Call flIniciaGrilla
        End If
        Txt_NumEndoso.SetFocus
    End If
    Screen.MousePointer = 0
End Sub
Private Sub Txt_NumEndoso_KeyPress(KeyAscii As Integer)
    Screen.MousePointer = 11
    If KeyAscii = 13 Then
        If (Trim(Txt_NumPol) <> "" Or Trim(Txt_Cuspp <> "") Or _
            Trim(Txt_TipIdent <> "") Or Trim(Txt_NumIdent <> "") Or _
            Trim(Txt_NumEndoso <> "") Or Trim(Txt_PatPen <> "") Or _
            Trim(Txt_MatPen <> "") Or Trim(Txt_NomPen <> "")) Then
            Call flLlenaGrilla
        Else
            Call flIniciaGrilla
        End If
        Txt_PatPen.SetFocus
    End If
    Screen.MousePointer = 0
End Sub
Private Sub Txt_PatPen_KeyPress(KeyAscii As Integer)
    Screen.MousePointer = 11
    If KeyAscii = 13 Then
        If (Trim(Txt_NumPol) <> "" Or Trim(Txt_Cuspp <> "") Or _
            Trim(Txt_TipIdent <> "") Or Trim(Txt_NumIdent <> "") Or _
            Trim(Txt_NumEndoso <> "") Or Trim(Txt_PatPen <> "") Or _
            Trim(Txt_MatPen <> "") Or Trim(Txt_NomPen <> "")) Then
            Call flLlenaGrilla
        Else
            Call flIniciaGrilla
        End If
        Txt_MatPen.SetFocus
    End If
    Screen.MousePointer = 0
End Sub
Private Sub Txt_MatPen_KeyPress(KeyAscii As Integer)
    Screen.MousePointer = 11
    If KeyAscii = 13 Then
        If (Trim(Txt_NumPol) <> "" Or Trim(Txt_Cuspp <> "") Or _
            Trim(Txt_TipIdent <> "") Or Trim(Txt_NumIdent <> "") Or _
            Trim(Txt_NumEndoso <> "") Or Trim(Txt_PatPen <> "") Or _
            Trim(Txt_MatPen <> "") Or Trim(Txt_NomPen <> "")) Then
            Call flLlenaGrilla
        Else
            Call flIniciaGrilla
        End If
        Txt_NomPen.SetFocus
    End If
    Screen.MousePointer = 0
End Sub
Private Sub Txt_NomPen_KeyPress(KeyAscii As Integer)
    Screen.MousePointer = 11
    If KeyAscii = 13 Then
        If (Trim(Txt_NumPol) <> "" Or Trim(Txt_Cuspp <> "") Or _
            Trim(Txt_TipIdent <> "") Or Trim(Txt_NumIdent <> "") Or _
            Trim(Txt_NumEndoso <> "") Or Trim(Txt_PatPen <> "") Or _
            Trim(Txt_MatPen <> "") Or Trim(Txt_NomPen <> "")) Then
            Call flLlenaGrilla
        Else
            Call flIniciaGrilla
        End If
        Txt_NumPol.SetFocus
    End If
    Screen.MousePointer = 0
End Sub
'-------------------- Fin KeyPress --------------------------------------

Private Sub flLlenaGrilla()
On Error GoTo Err_LlenaGrilla

    Txt_NumPol = Format(Trim(Txt_NumPol), "0000000000")
    Txt_Cuspp = Trim(UCase(Txt_Cuspp))
    Txt_TipIdent = Trim(UCase(Txt_TipIdent))
    Txt_NumIdent = Trim(UCase(Txt_NumIdent))
    Txt_NumEndoso = Trim(UCase(Txt_NumEndoso))
    Txt_PatPen = Trim(UCase(Txt_PatPen))
    Txt_MatPen = Trim(UCase(Txt_MatPen))
    Txt_NomPen = Trim(UCase(Txt_NomPen))
    
    vgSql = ""
    vgSql = "select B.num_poliza,P.cod_cuspp, B.cod_tipoidenben, B.num_idenben,"
    vgSql = vgSql & " i.gls_tipoidencor as tipoiden,"
    vgSql = vgSql & " B.gls_patben, B.gls_matben, B.gls_nomben, B.num_endoso "
    vgSql = vgSql & " FROM PP_TMAE_BEN B, PP_TMAE_POLIZA P,ma_tpar_tipoiden i WHERE "
    vgSql = vgSql & "P.num_poliza = B.num_poliza "
    vgSql = vgSql & "and P.num_endoso=B.num_endoso "
    vgSql = vgSql & "AND B.cod_tipoidenben = i.cod_tipoiden "
    If Trim(Txt_NumPol <> "") Then vgSql = vgSql & " and B.num_poliza like '" & Txt_NumPol & "%'"
    If Trim(Txt_Cuspp <> "") Then vgSql = vgSql & " and P.cod_cuspp like '" & Txt_Cuspp & "%'"
    If Trim(Txt_TipIdent <> "") Then vgSql = vgSql & " and i.gls_tipoidencor like '" & Txt_TipIdent & "%'"
    If Trim(Txt_NumIdent <> "") Then vgSql = vgSql & " and B.num_idenben like '" & Txt_NumIdent & "%'"
    If Trim(Txt_NumEndoso <> "") Then vgSql = vgSql & " and B.num_endoso like '" & Txt_NumEndoso & "%'"
    If Trim(Txt_PatPen <> "") Then vgSql = vgSql & " and B.gls_patben like '" & Txt_PatPen & "%'"
    If Trim(Txt_MatPen <> "") Then vgSql = vgSql & " and B.gls_matben like '" & Txt_MatPen & "%'"
    If Trim(Txt_NomPen <> "") Then vgSql = vgSql & " and B.gls_nomben like '" & Txt_NomPen & "%'"
    vgSql = vgSql & "order by B.num_poliza, B.num_endoso"
    Set vgRs = vgConexionBD.Execute(vgSql)
    
    Msf_Grilla.Rows = 1
    
    Do While Not vgRs.EOF
        If Not IsNull(vgRs!Num_Poliza) Then vlNumPol = Trim(vgRs!Num_Poliza)
        If Not IsNull(vgRs!Cod_Cuspp) Then vlCuspp = Trim(vgRs!Cod_Cuspp)
        If Not IsNull(vgRs!tipoiden) Then vlTipoIdent = Trim(vgRs!tipoiden)
        If Not IsNull(vgRs!Num_IdenBen) Then vlIdenBen = Trim(vgRs!Num_IdenBen)
        If Not IsNull(vgRs!num_endoso) Then vlNumEnd = Trim(vgRs!num_endoso)
        If Not IsNull(vgRs!Gls_PatBen) Then vlPatPen = Trim(vgRs!Gls_PatBen)
        If Not IsNull(vgRs!Gls_MatBen) Then vlMatPen = Trim(vgRs!Gls_MatBen)
        If Not IsNull(vgRs!Gls_NomBen) Then vlNomPen = Trim(vgRs!Gls_NomBen)
        
        Msf_Grilla.AddItem vlNumPol & vbTab & vlCuspp & vbTab & _
                        vlTipoIdent & vbTab & vlIdenBen & vbTab & _
                        vlNumEnd & vbTab & vlPatPen & vbTab & _
                        vlMatPen & vbTab & vlNomPen
                                    
        vgRs.MoveNext
    Loop

Exit Sub
Err_LlenaGrilla:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

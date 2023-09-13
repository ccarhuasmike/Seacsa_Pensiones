VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "Crystl32.OCX"
Begin VB.Form Frm_MenIndividual 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de Mensajes por Pólizas - Pensionados."
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   9870
   Begin VB.Frame Fra_Seleccion 
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
      Height          =   1200
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   9615
      Begin VB.CommandButton Cmd_Buscar 
         Height          =   375
         Left            =   4560
         Picture         =   "Frm_MenIndividual.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Buscar Póliza"
         Top             =   255
         Width           =   615
      End
      Begin VB.TextBox Txt_Anno 
         Height          =   285
         Left            =   2490
         MaxLength       =   4
         TabIndex        =   2
         Top             =   330
         Width           =   540
      End
      Begin VB.TextBox Txt_Mes 
         Height          =   285
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   1
         Top             =   330
         Width           =   420
      End
      Begin VB.ComboBox Cmb_CodMensaje 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   690
         Width           =   7785
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "  Selección de Mensaje  "
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
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   25
         Top             =   0
         Width           =   2100
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "(Mes - Año)"
         Height          =   300
         Index           =   2
         Left            =   3225
         TabIndex        =   23
         Top             =   330
         Width           =   930
      End
      Begin VB.Label Lbl_Nombre 
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
         Height          =   210
         Index           =   1
         Left            =   2265
         TabIndex        =   22
         Top             =   345
         Width           =   150
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Periodo"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   21
         Top             =   330
         Width           =   780
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Código Mensaje"
         Height          =   270
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   690
         Width           =   1335
      End
   End
   Begin VB.Frame Fra_Botones 
      Height          =   975
      Left            =   120
      TabIndex        =   18
      Top             =   5400
      Width           =   9645
      Begin VB.CommandButton Cmd_Cancelar2 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   5400
         Picture         =   "Frm_MenIndividual.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Eliminar 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   3240
         Picture         =   "Frm_MenIndividual.frx":06DC
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Eliminar Año"
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   4320
         Picture         =   "Frm_MenIndividual.frx":0A1E
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   6480
         Picture         =   "Frm_MenIndividual.frx":10D8
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Grabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   2160
         Picture         =   "Frm_MenIndividual.frx":11D2
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   200
         Width           =   720
      End
   End
   Begin VB.Frame Fra_Asignacion 
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
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   9615
      Begin VB.ComboBox Cmb_PenNumIdent 
         BackColor       =   &H00E0FFFF&
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   600
         Width           =   2235
      End
      Begin VB.CommandButton Cmd_Cancelar 
         Height          =   450
         Left            =   7455
         Picture         =   "Frm_MenIndividual.frx":188C
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Quitar Beneficiario"
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton Cmd_Restar 
         Height          =   450
         Left            =   6960
         Picture         =   "Frm_MenIndividual.frx":1EFE
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Quitar Beneficiario"
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton Cmd_Sumar 
         Height          =   450
         Left            =   6465
         Picture         =   "Frm_MenIndividual.frx":2088
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Agregar Beneficiario"
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton Opt_Rut 
         Caption         =   "Por Identificación"
         Height          =   255
         Left            =   960
         TabIndex        =   6
         Top             =   675
         Width           =   1695
      End
      Begin VB.OptionButton Opt_Poliza 
         Caption         =   "Por Pólizas"
         Height          =   255
         Left            =   960
         TabIndex        =   5
         Top             =   315
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox Txt_PenNumIdent 
         Height          =   285
         Left            =   4800
         MaxLength       =   10
         TabIndex        =   9
         Top             =   600
         Width           =   1620
      End
      Begin VB.TextBox Txt_Poliza 
         Height          =   285
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   7
         Top             =   315
         Width           =   1605
      End
      Begin MSFlexGridLib.MSFlexGrid Msf_Grilla 
         Height          =   2625
         Left            =   195
         TabIndex        =   24
         Top             =   1080
         Width           =   9180
         _ExtentX        =   16193
         _ExtentY        =   4630
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
      Begin Crystal.CrystalReport Rpt_General 
         Left            =   8805
         Top             =   210
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowState     =   2
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "  Asignación de Mensaje  "
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
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   26
         Top             =   0
         Width           =   2200
      End
   End
End
Attribute VB_Name = "Frm_MenIndividual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Sql As String

Dim vlPeriodo As String
Dim vlFecha As String
Dim vlAnno As String
Dim vlPos As Integer
Dim vlUltEndoso As Integer
Dim vlPunt As Integer
Dim vlRut As String
Dim vlTipoIden As String
Dim vlNumIden As String
Dim vlModificar As Boolean
Dim vlRutAux As String
Dim vlNumerosPolizas As String
Dim vlcont As Integer
Dim vlFechaPeriodo As String
Dim vlNumPoliza As String


Dim vlGlsUsuarioCrea As Variant
Dim vlFecCrea As Variant
Dim vlHorCrea As Variant
Dim vlGlsUsuarioModi As Variant
Dim vlFecModi As Variant
Dim vlHorModi As Variant


Dim vlCodMensaje As Integer
Dim vlNumero As Integer

Dim vlSwModificar As String * 1
Dim vlSwSinDerPen As String * 1
Dim vlSwEstPeriodo As String * 1

Dim vlOperacion As String

Const clCodTipoIng As String * 1 = "I"
Const clCodSinDerPen As String * 2 = "10"

Dim vlCodTipoIdenBenCau As String
Dim vlNumIdenBenCau As String

Dim vlLargoTipoIden    As Integer 'sirve para llenar la grilla
Dim vlPosicionTipoIden As Integer 'sirve para llenar la grilla

Function flLimpiar()
    
    Opt_Poliza.Enabled = True
    Opt_Rut.Enabled = True
    Txt_Poliza.Enabled = True
    Cmb_PenNumIdent.Enabled = True
    Txt_PenNumIdent.Enabled = True
    Cmd_Sumar.Enabled = True
    Cmd_Restar.Enabled = True
    Cmd_Cancelar.Enabled = True
    
    
    Opt_Poliza.Enabled = True
    Txt_Poliza.Text = ""
    If (Cmb_PenNumIdent.ListCount <> 0) Then
        Cmb_PenNumIdent.ListIndex = 0
    End If
    Txt_PenNumIdent.Text = ""
    Call flInicializaGrilla
    Call flDeshabilitaIngreso
    Txt_Mes.SetFocus
    
'''    Txt_GlsMensaje.Text = ""
    Txt_Mes.Text = ""
    Txt_Anno.Text = ""
    Cmb_CodMensaje.ListIndex = 0
    
    vlModificar = False
    
End Function

Function flInicializaGrilla()

    Msf_Grilla.Clear
    Msf_Grilla.Cols = 7
    Msf_Grilla.Rows = 1
    Msf_Grilla.RowHeight(0) = 250
    Msf_Grilla.Row = 0
    
    Msf_Grilla.Col = 0
    Msf_Grilla.Text = "Nº Póliza"
    Msf_Grilla.ColWidth(0) = 1100
    Msf_Grilla.ColAlignment(0) = 1  'centrado
    
    Msf_Grilla.Col = 1
    Msf_Grilla.Text = "Nº Endoso"
    Msf_Grilla.ColWidth(1) = 900
    Msf_Grilla.ColAlignment(0) = 1  'centrado
    
    Msf_Grilla.Col = 2
    Msf_Grilla.Text = "Nº Orden"
    Msf_Grilla.ColWidth(2) = 900
    Msf_Grilla.ColAlignment(0) = 1  'centrado
    
    Msf_Grilla.Col = 3
    Msf_Grilla.Text = "Tipo Ident."
    Msf_Grilla.ColWidth(3) = 1300
    
    Msf_Grilla.Col = 4
    Msf_Grilla.Text = "Nº Ident."
    Msf_Grilla.ColWidth(4) = 1300
    Msf_Grilla.ColAlignment(0) = 1  'centrado
    
    Msf_Grilla.Col = 5
    Msf_Grilla.Text = "Nombre"
    Msf_Grilla.ColWidth(5) = 3600
    
    Msf_Grilla.Col = 6
    Msf_Grilla.Text = "Parent."
    Msf_Grilla.ColWidth(6) = 800
        
End Function

Function flCargaGrilla()
Dim vlTipoI As String
On Error GoTo Err_Carga
    
    vlModificar = False
'    vlPeriodo = ""
    'vlPeriodo = Trim((Txt_Anno.Text) + (Txt_Mes.Text))
    
    vlPeriodo = ""
    vlPeriodo = Trim((Txt_Anno.Text) + (Txt_Mes.Text))
    vlPeriodo = Format(Txt_Anno.Text, "0000") & Format(Txt_Mes.Text, "00")
       
    vgSql = ""
    vgSql = " SELECT num_poliza,num_endoso,num_orden, "
    vgSql = vgSql & " cod_usuariocrea,fec_crea,hor_crea "
    vgSql = vgSql & "FROM PP_TMAE_MENPOLIZA "
    vgSql = vgSql & " WHERE "
    vgSql = vgSql & " num_perpago = '" & Trim(vlPeriodo) & "' AND "
    vgSql = vgSql & " cod_tipoing = '" & Trim(clCodTipoIng) & "' AND "
    
    vlNumero = InStr(Cmb_CodMensaje.Text, "-")
    vlCodMensaje = Trim(Mid(Cmb_CodMensaje.Text, 1, vlNumero - 1))
    
    vgSql = vgSql & " cod_mensaje = '" & Trim(vlCodMensaje) & "' "
    
    vgSql = vgSql & "ORDER by num_poliza ASC "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
    
       Call flInicializaGrilla
       
       vlGlsUsuarioCrea = (vgRs!Cod_UsuarioCrea)
       vlFecCrea = (vgRs!Fec_Crea)
       vlHorCrea = (vgRs!Hor_Crea)
       While Not vgRs.EOF
'***** BUSCAR DATOS DE BENEFICIARIO *****
          Sql = "SELECT cod_tipoidenben,num_idenben,gls_nomben,gls_patben,"
          Sql = Sql & "gls_matben,cod_par FROM PP_TMAE_BEN "
          Sql = Sql & "WHERE "
          Sql = Sql & " num_poliza = '" & (vgRs!num_poliza) & "' AND "
          'Sql = Sql & " num_endoso = '" & (vgRs!num_endoso) & "' AND "
          Sql = Sql & " num_orden = '" & (vgRs!Num_Orden) & "' "
          Set vgRs2 = vgConexionBD.Execute(Sql)
          If Not vgRs2.EOF Then
            
            vlTipoI = " " & Trim(vgRs2!Cod_TipoIdenBen) & " - " & fgBuscarNombreTipoIden(vgRs2!Cod_TipoIdenBen, False)

             Msf_Grilla.AddItem CStr(Trim(vgRs!num_poliza)) & vbTab _
             & Trim(vgRs!num_endoso) & vbTab _
             & Trim(vgRs!Num_Orden) & vbTab _
             & (vlTipoI) & vbTab & (Trim(vgRs2!Num_IdenBen)) & vbTab _
             & ((Trim(vgRs2!Gls_NomBen)) & " " & (Trim(vgRs2!Gls_PatBen)) & " " & (Trim(vgRs2!Gls_MatBen))) & vbTab _
             & Trim(vgRs2!Cod_Par)
          End If
          
          vlModificar = True
          
          vgRs.MoveNext
             
       Wend
    End If
    vgRs.Close

Exit Function
Err_Carga:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function flCargaCombo(vlCombo As ComboBox)

Dim vlcont As Integer
On Error GoTo Err_Combo
    
    vlCombo.Clear
    vgQuery = "SELECT cod_mensaje,gls_mensaje "
    vgQuery = vgQuery & "FROM PP_TPAR_MENSAJE "
    vgQuery = vgQuery & "ORDER BY cod_mensaje "
    Set vgCmb = vgConexionBD.Execute(vgQuery)
    If Not (vgCmb.EOF) Then
        While Not (vgCmb.EOF)
            vlCombo.AddItem ((Trim(vgCmb!cod_mensaje) & " - " & Trim(vgCmb!Gls_Mensaje)))
            'vlCombo.AddItem ((Trim(vgCmb!cod_mensaje)))
            vgCmb.MoveNext
        Wend
        
        If (vlCombo.ListCount <> 0) Then
            vlCombo.ListIndex = 0
        End If
    End If
    vgCmb.Close
    
Exit Function
Err_Combo:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

''''Function flBuscarGlosaMensaje(iNumeroMensaje) As String
''''On Error GoTo Err_BuscarGlosa
''''
''''    If IsNumeric(iNumeroMensaje) Then
''''        vgQuery = "SELECT gls_mensaje "
''''        vgQuery = vgQuery & "FROM PP_TPAR_MENSAJE "
''''        vgQuery = vgQuery & "WHERE cod_mensaje = " & iNumeroMensaje & " "
''''        Set vgRs = vgConexionBD.Execute(vgQuery)
''''        If Not (vgRs.EOF) Then
''''            flBuscarGlosaMensaje = (Trim(vgRs!gls_mensaje))
''''        Else
''''            flBuscarGlosaMensaje = ""
''''        End If
''''        vgRs.Close
''''    End If
''''
''''Exit Function
''''Err_BuscarGlosa:
''''    Screen.MousePointer = 0
''''    Select Case Err
''''        Case Else
''''        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
''''    End Select
''''End Function

Function flHabilitaIngreso()

    Fra_Seleccion.Enabled = False

'    Txt_Mes.Enabled = False
'    Txt_Anno.Enabled = False
'    Cmb_CodMensaje.Enabled = False
'    Cmd_Buscar.Enabled = False
    
    Fra_Asignacion.Enabled = True
    
    Opt_Poliza.Enabled = True
    Opt_Rut.Enabled = True
    Txt_Poliza.Enabled = True
    Cmb_PenNumIdent.Enabled = True
    Txt_PenNumIdent.Enabled = True
    Cmd_Sumar.Enabled = True
    Cmd_Restar.Enabled = True
    Cmd_Cancelar.Enabled = True
    
    Txt_Poliza.SetFocus
    Call Opt_Poliza_Click
    

End Function

Function flDeshabilitaIngreso()

    Fra_Seleccion.Enabled = True
    
'    Txt_Mes.Enabled = True
'    Txt_Anno.Enabled = True
'    Cmb_CodMensaje.Enabled = True
'    Cmd_Buscar.Enabled = True
    
    Fra_Asignacion.Enabled = False
    
'    Opt_Poliza.Enabled = False
'    Opt_Poliza.Enabled = False
'    Txt_Poliza.Enabled = False
'    Txt_Rut.Enabled = False
'    Txt_Digito.Enabled = False
'    Cmd_Sumar.Enabled = False
'    Cmd_Restar.Enabled = False
'    Cmd_Cancelar.Enabled = False
    
    
    

End Function

Function flMostrarMensajeCantidad()

On Error GoTo Err_flMostrarMensajeCantidad

    vgSql = ""
    vgSql = "SELECT DISTINCT num_poliza "
    vgSql = vgSql & "FROM PP_TMAE_BEN WHERE "
    vgSql = vgSql & "cod_tipoidenben = " & vlCodTipoIdenBenCau & " AND "
    vgSql = vgSql & "num_idenben = '" & vlNumIdenBenCau & "' AND "
    vgSql = vgSql & "cod_estpension <> '" & clCodSinDerPen & "' "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       vlcont = 0
       vlNumerosPolizas = ""
       While Not vgRegistro.EOF
             If vlNumerosPolizas = "" Then
                vlNumerosPolizas = Trim(vgRegistro!num_poliza)
             Else
                 vlNumerosPolizas = vlNumerosPolizas & " - " & Trim(vgRegistro!num_poliza)
             End If
             vlcont = vlcont + 1
             vgRegistro.MoveNext
       Wend
       vlNumerosPolizas = vlNumerosPolizas & " ."

       MsgBox " La Identificación del Beneficiario Ingresado, se encuentra Asociado a las Siguientes Pólizas : " & Chr(13) & _
               vlNumerosPolizas & "  Ingrese Número de Póliza.", vbExclamation, "Información"
                        
    End If
    
Exit Function
Err_flMostrarMensajeCantidad:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

''Function flValidaEstadoPeriodo(ifecha, iPoliza, iOrden)
''
''On Error GoTo Err_flValidaEstadoPeriodo
''
''    If Not fgValidaPagoPension(ifecha, iPoliza, iOrden) Then
''       MsgBox " Ya se ha Realizado el Proceso de Cálculo de Pensión para ésta Fecha ", vbCritical, "Operación Cancelada"
''
''    End If
''
''Exit Function
''Err_flMostrarMensajeCantidad:
''    Screen.MousePointer = 0
''    Select Case Err
''        Case Else
''        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
''    End Select
''
''
''End Function

Private Sub Cmb_CodMensaje_Click()

'''    If Trim(Cmb_CodMensaje.Text) <> "" Then
'''       Txt_GlsMensaje.Text = flBuscarGlosaMensaje(Cmb_CodMensaje.Text)
'''    End If
    
End Sub

Private Sub Cmb_CodMensaje_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Cmd_Buscar.SetFocus
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
            Cmd_Sumar.SetFocus
        End If
    End If
End Sub

Private Sub Cmd_Buscar_Click()
On Error GoTo Err_Buscar
    
    Txt_Mes = Format(Txt_Mes, "00")
    Txt_Anno = Format(Txt_Anno, "0000")
    
'Valida dato ingresado en TEXT correspondiente a mes.
    If Txt_Mes.Text = "" Then
       MsgBox "Debe Ingresar Mes del Período.", vbCritical, "Error de Datos"
       Txt_Mes.SetFocus
       Exit Sub
    End If
         
    If CLng(Txt_Mes.Text) <= 0 Or CLng(Txt_Mes.Text) > 12 Then
       MsgBox "El Mes Ingresado no es un Valor Válido.", vbCritical, "Error de Datos"
       Txt_Mes.SetFocus
       Exit Sub
    End If
    
'Valida dato ingresado en TEXT correspondiente a año.
    If Txt_Anno.Text = "" Then
       MsgBox "Debe Ingresar el Año del Período.", vbCritical, "Error de Datos"
       Txt_Anno.SetFocus
       Exit Sub
    End If
    
    vlFecha = Date
    vlFecha = Format(CDate(Trim(vlFecha)), "yyyymmdd")
    vlAnno = (Mid(vlFecha, 1, 4))
         
    If CLng(Txt_Anno.Text) < 1900 Or CLng(Txt_Anno.Text) > vlAnno Then
       MsgBox "Debe ingresar un Año Mayor a 1900 o Menor Igual al Actual.", vbCritical, "Error de Datos"
       Txt_Anno.SetFocus
       Exit Sub
    End If
    
    'Varificación de existencia de Período
    vlPeriodo = Format(Txt_Anno, "0000") & Format(Txt_Mes, "00")
    vlOperacion = ""
    
    
    'Verifica Existencia del Periodo Ingresado
    vgSql = ""
    vgSql = "SELECT num_perpago "
    vgSql = vgSql & "FROM PP_TMAE_MENPOLIZA "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "num_perpago = '" & Trim(vlPeriodo) & "' AND "
    vgSql = vgSql & "cod_tipoing = '" & Trim(clCodTipoIng) & "' AND "
    
    vlNumero = InStr(Cmb_CodMensaje.Text, "-")
    vlCodMensaje = Trim(Mid(Cmb_CodMensaje.Text, 1, vlNumero - 1))
    
    vgSql = vgSql & "cod_mensaje = " & Trim(Str(vlCodMensaje)) & " "
    vgSql = vgSql & "ORDER by num_poliza ASC "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If vgRs.EOF Then
       vlOperacion = "S"
    End If
    
    'Validar si Periodo se encuentra abierto
    'vlSwEstPeriodo = C    Periodo Cerrado
    'vlSwEstPeriodo = A    Periodo Abierto
    vlSwEstPeriodo = "C"
    
    vgSql = ""
    vgSql = "SELECT num_perpago "
    vgSql = vgSql & "FROM PP_TMAE_PROPAGOPEN "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "num_perpago = '" & Trim(vlPeriodo) & "' AND "
    'vgSql = vgSql & "(cod_estadoreg = 'A' OR cod_estadopri = 'A') "
    vgSql = vgSql & "(cod_estadoreg <> 'C' ) " ''*OR cod_estadopri <> 'C'
    Set vgRs = vgConexionBD.Execute(vgSql)
    If vgRs.EOF Then
       vlSwEstPeriodo = "C"
    Else
        vlSwEstPeriodo = "A"
    End If
        
 
    'Si el periodo ingresado no existe
    If (vlOperacion = "S") Then
        'El periodo no existe
        If vlSwEstPeriodo = "A" Then
           'El periodo no existe y se encuentra abierto
           Call flHabilitaIngreso
        Else
            'El periodo no existe y se encuentra cerrado
            MsgBox "El Período Ingresado se encuentra Cerrado, Debe Ingresar un Nuevo Periodo.", vbCritical, "Error de Datos"
            Txt_Mes.SetFocus
            Exit Sub
        End If
        
        vgRes = MsgBox("El Período indicado no se encuentra registrado." & Chr(13) & _
                       "   ¿ Desea ingresar este nuevo Periodo ?", 4 + 32 + 256, "Operación de Ingreso")
        If vgRes <> 6 Then
            Fra_Seleccion.Enabled = True
            Txt_Mes.SetFocus
            Screen.MousePointer = 0
            Exit Sub
        End If
                
    Else
        'Si el periodo ingresado Existe
        If vlSwEstPeriodo = "A" Then
           'El periodo ingresado existe, y se encuentra abierto
           Call flHabilitaIngreso
           Call flCargaGrilla
        Else
            'El periodo ingresado existe pero se encuentra cerrado
            MsgBox "El Periodo Ingresado se encuentra Cerrado, Sólo podrá Consultar los Datos.", vbCritical, "Error de Datos"
            Opt_Poliza.Enabled = False
            Opt_Rut.Enabled = False
            Txt_Poliza.Enabled = False
            Cmb_PenNumIdent.Enabled = False
            Txt_PenNumIdent.Enabled = False
            Cmd_Sumar.Enabled = False
            Cmd_Restar.Enabled = False
            Cmd_Cancelar.Enabled = False
            
            Call flCargaGrilla
            
        End If
    End If
        
'    Call flCargaGrilla
'    Call flHabilitaIngreso
    
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

    Call Opt_Poliza_Click
    
Exit Sub
Err_Cancelar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Cancelar2_Click()

On Error GoTo Err_Limpiar

    Call flLimpiar

Exit Sub
Err_Limpiar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub

Private Sub Cmd_Eliminar_Click()
On Error GoTo Err_Eliminar

    If vlSwEstPeriodo = "C" Then
       MsgBox "El Periodo Ingresado se encuentra Cerrado, Sólo podra Consultar los Datos.", vbCritical, "Error de Datos"
       Exit Sub
    End If

    If Txt_Mes.Enabled = False Then
              
       vlPeriodo = ""
       vlPeriodo = Trim((Txt_Anno.Text) + (Txt_Mes.Text))
       vlPeriodo = Format(Txt_Anno.Text, "0000") & Format(Txt_Mes.Text, "00")
              
       vgSql = "SELECT num_perpago "
       vgSql = vgSql & " FROM PP_TMAE_MENPOLIZA "
       vgSql = vgSql & " WHERE (num_perpago = '" & Trim(Str(vlPeriodo)) & "') AND "
       vgSql = vgSql & " (cod_tipoing = '" & Trim(clCodTipoIng) & "' ) AND "
       
       vlNumero = InStr(Cmb_CodMensaje.Text, "-")
       vlCodMensaje = Trim(Mid(Cmb_CodMensaje.Text, 1, vlNumero - 1))
       
       vgSql = vgSql & " (cod_mensaje = '" & Trim(vlCodMensaje) & "') "
       Set vgRs = vgConexionBD.Execute(vgSql)
       If Not vgRs.EOF Then
              
          vgSql = " DELETE PP_TMAE_MENPOLIZA "
          vgSql = vgSql & " WHERE (num_perpago = '" & Trim(Str(vlPeriodo)) & "') AND "
          vgSql = vgSql & " (cod_tipoing = '" & Trim(clCodTipoIng) & "' ) AND "
          
          vlNumero = InStr(Cmb_CodMensaje.Text, "-")
          vlCodMensaje = Trim(Mid(Cmb_CodMensaje.Text, 1, vlNumero - 1))
          
          vgSql = vgSql & " (cod_mensaje = '" & Trim(vlCodMensaje) & "') "
          vgConexionBD.Execute vgSql
          MsgBox "La Eliminación de Datos fue realizada Correctamente.", vbInformation, "Información"
          Call flLimpiar
          
       End If
    Else
        MsgBox "No Existen Datos a Eliminar.", vbCritical, "Error de Datos"
        
    End If
                    
Exit Sub
Err_Eliminar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub cmd_grabar_Click()
On Error GoTo Err_Grabar

    If vlSwEstPeriodo = "C" Then
       MsgBox "El Periodo Ingresado se encuentra Cerrado, Sólo podra Consultar los Datos.", vbCritical, "Error de Datos"
       Exit Sub
    End If

    If Txt_Mes.Enabled = False Then
              
       vlPeriodo = ""
       vlPeriodo = Trim((Txt_Anno.Text) + (Txt_Mes.Text))
       vlPeriodo = Format(Txt_Anno.Text, "0000") & Format(Txt_Mes.Text, "00")
              
       vgSql = "SELECT num_perpago "
       vgSql = vgSql & " FROM PP_TMAE_MENPOLIZA "
       vgSql = vgSql & " WHERE (num_perpago = '" & Trim(Str(vlPeriodo)) & "') AND "
       vgSql = vgSql & " (cod_tipoing = '" & Trim(clCodTipoIng) & "' ) AND "
       
       vlNumero = InStr(Cmb_CodMensaje.Text, "-")
       vlCodMensaje = Trim(Mid(Cmb_CodMensaje.Text, 1, vlNumero - 1))
       
       vgSql = vgSql & " (cod_mensaje = '" & Trim(vlCodMensaje) & "') "
       
       Set vgRs = vgConexionBD.Execute(vgSql)
       If Not vgRs.EOF Then
                 
          vlModificar = True
          vgSql = " DELETE PP_TMAE_MENPOLIZA "
          vgSql = vgSql & " WHERE (num_perpago = '" & Trim(Str(vlPeriodo)) & "') AND "
          vgSql = vgSql & " (cod_tipoing = '" & Trim(clCodTipoIng) & "' ) AND "
          
          vlNumero = InStr(Cmb_CodMensaje.Text, "-")
          vlCodMensaje = Trim(Mid(Cmb_CodMensaje.Text, 1, vlNumero - 1))
          
          vgSql = vgSql & " (cod_mensaje = '" & Trim(vlCodMensaje) & "') "
          vgConexionBD.Execute vgSql
'          MsgBox "La Eliminación de Datos fue realizada Correctamente.", vbInformation, "Información"

       End If
       
       If vlModificar = True Then
          vlGlsUsuarioModi = vgUsuario
          vlFecModi = Format(Date, "yyyymmdd")
          vlHorModi = Format(Time, "hhmmss")
       Else
           vlGlsUsuarioCrea = vgUsuario
           vlFecCrea = Format(Date, "yyyymmdd")
           vlHorCrea = Format(Time, "hhmmss")
           vlGlsUsuarioModi = Null
           vlFecModi = Null
           vlHorModi = Null
       End If
                        
       vlPunt = 1
       Msf_Grilla.Col = 0
       While vlPunt < Msf_Grilla.Rows
             Msf_Grilla.Row = vlPunt
             
                   
             vgSql = "SELECT num_perpago "
             vgSql = vgSql & " FROM PP_TMAE_MENPOLIZA "
             vgSql = vgSql & " WHERE (num_perpago = '" & Trim(Str(vlPeriodo)) & "') AND "
             
             Msf_Grilla.Col = 0
             vgSql = vgSql & " (num_poliza = '" & Trim(Msf_Grilla.Text) & "' ) AND "
                
             Msf_Grilla.Col = 2
             vgSql = vgSql & " (num_orden = '" & Trim(Msf_Grilla.Text) & "' ) AND "
                          
             vlNumero = InStr(Cmb_CodMensaje.Text, "-")
             vlCodMensaje = Trim(Mid(Cmb_CodMensaje.Text, 1, vlNumero - 1))
                
             vgSql = vgSql & " (cod_mensaje = '" & Trim(vlCodMensaje) & "') "
                
             Set vgRegistro = vgConexionBD.Execute(vgSql)
             If vgRegistro.EOF Then
                                   
                vgSql = ""
                vgSql = "INSERT INTO PP_TMAE_MENPOLIZA "
                vgSql = vgSql & "(num_poliza,num_endoso,num_orden,cod_mensaje, "
                vgSql = vgSql & " num_perpago,cod_tipoing, "
                vgSql = vgSql & " cod_usuariocrea,fec_crea,hor_crea, "
                vgSql = vgSql & " cod_usuariomodi,fec_modi,hor_modi "
                vgSql = vgSql & " ) VALUES ( "
                Msf_Grilla.Col = 0
                vgSql = vgSql & " '" & Trim(Msf_Grilla.Text) & "' , "
                Msf_Grilla.Col = 1
                vgSql = vgSql & " " & Trim(Str(Msf_Grilla.Text)) & ", "
                Msf_Grilla.Col = 2
                vgSql = vgSql & " " & Trim(Str(Msf_Grilla.Text)) & ", "
                
                vlNumero = InStr(Cmb_CodMensaje.Text, "-")
                vlCodMensaje = Trim(Mid(Cmb_CodMensaje.Text, 1, vlNumero - 1))
                
                vgSql = vgSql & " '" & Trim(vlCodMensaje) & "', "
                vgSql = vgSql & " " & Trim(Str(vlPeriodo)) & ", "
                vgSql = vgSql & " '" & Trim(clCodTipoIng) & "', "
                vgSql = vgSql & " '" & vlGlsUsuarioCrea & "', "
                vgSql = vgSql & " '" & vlFecCrea & "', "
                vgSql = vgSql & " '" & vlHorCrea & "', "
                vgSql = vgSql & " '" & vlGlsUsuarioModi & "', "
                vgSql = vgSql & " '" & vlFecModi & "', "
                vgSql = vgSql & " '" & vlHorModi & "' ) "
                
                vgConexionBD.Execute vgSql
                
             End If
                   
             vlPunt = (vlPunt + 1)
        Wend
        MsgBox "El Proceso de Actualización ha sido llevado a cabo Exitosamente.", vbInformation, "Información"
    Else
        MsgBox "No se han ingresado Pólizas a registrar.", vbCritical, "Error de Datos"
    End If

Exit Sub
Err_Grabar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Imprimir_Click()
Dim vlArchivo As String
On Error GoTo Err_CmdImprimir
   
   'Validar que se encuentren ingresados los datos del periodo a imprimir
   'Valida dato ingresado en TEXT correspondiente a mes.
    If Txt_Mes.Text = "" Then
       MsgBox "Debe Ingresar Mes del Período.", vbCritical, "Error de Datos"
       Txt_Mes.SetFocus
       Exit Sub
    End If
         
    If CLng(Txt_Mes.Text) <= 0 Or CLng(Txt_Mes.Text) > 12 Then
       MsgBox "El Mes Ingresado no es un Valor Válido.", vbCritical, "Error de Datos"
       Txt_Mes.SetFocus
       Exit Sub
    End If
    
    'Valida dato ingresado en TEXT correspondiente a año.
    If Txt_Anno.Text = "" Then
       MsgBox "Debe Ingresar el Año del Período.", vbCritical, "Error de Datos"
       Txt_Anno.SetFocus
       Exit Sub
    End If
    
    vlFecha = Date
    vlFecha = Format(CDate(Trim(vlFecha)), "yyyymmdd")
    vlAnno = (Mid(vlFecha, 1, 4))
         
    If CLng(Txt_Anno.Text) < 1900 Or CLng(Txt_Anno.Text) > vlAnno Then
       MsgBox "Debe ingresar un Año Mayor a 1900 o Menor Igual al Actual.", vbCritical, "Error de Datos"
       Txt_Anno.SetFocus
       Exit Sub
    End If
    
    'Varificación de existencia de Período
    vlPeriodo = Format(Txt_Anno, "0000") & Format(Txt_Mes, "00")
    vlOperacion = ""
    
    vgSql = "SELECT num_perpago "
    vgSql = vgSql & "FROM PP_TMAE_MENPOLIZA "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "num_perpago = '" & Trim(vlPeriodo) & "' AND "
    vgSql = vgSql & "cod_tipoing = '" & Trim(clCodTipoIng) & "' AND "
    
    vlNumero = InStr(Cmb_CodMensaje.Text, "-")
    vlCodMensaje = Trim(Mid(Cmb_CodMensaje.Text, 1, vlNumero - 1))
    
    vgSql = vgSql & "cod_mensaje = " & Trim(Str(vlCodMensaje)) & " "
    vgSql = vgSql & "ORDER by num_poliza ASC "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If vgRs.EOF Then
        MsgBox "El Periodo indicado no se encuentra registrado.", vbInformation, "Operación de Ingreso"
        Exit Sub
    End If
   
   Screen.MousePointer = 11
   
   vlArchivo = strRpt & "PP_Rpt_MenIndividuales.rpt"   '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
       MsgBox "Archivo de Reporte Mensajes Individuales no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
       Screen.MousePointer = 0
       Exit Sub
   End If

   vgQuery = "{PP_TMAE_MENPOLIZA.COD_TIPOING} = '" & (clCodTipoIng) & "' AND "
   vgQuery = vgQuery & "{PP_TMAE_MENPOLIZA.NUM_PERPAGO} = '" & Trim(vlPeriodo) & "'"
   
   Rpt_General.Reset
   Rpt_General.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
   Rpt_General.Connect = vgRutaDataBase       ' o App.Path & "\Nestle.mdb"
   Rpt_General.SelectionFormula = vgQuery
   Rpt_General.Formulas(0) = ""
   Rpt_General.Formulas(1) = ""
   Rpt_General.Formulas(2) = ""
   Rpt_General.Formulas(3) = ""
   
   Rpt_General.Formulas(1) = "NombreCompania = '" & vgNombreCompania & "'"
   Rpt_General.Formulas(2) = "NombreSistema = '" & vgNombreSistema & "'"
   Rpt_General.Formulas(3) = "NombreSubSistema = '" & vgNombreSubSistema & "'"
   
   Rpt_General.WindowState = crptMaximized
   Rpt_General.Destination = crptToWindow
   Rpt_General.WindowTitle = "Informe de Mensajes Individuales"
   Rpt_General.Action = 1
   Screen.MousePointer = 0
   
Exit Sub
Err_CmdImprimir:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Restar_Click()
On Error GoTo Err_Restar

    If Txt_Poliza.Text = "" Then
       MsgBox "Debe Ingresar el Nº de Póliza.", vbCritical, "Error de Datos"
       Txt_Poliza.SetFocus
       Exit Sub
    End If
         
    If CDbl(Txt_Poliza.Text) <= 0 Then
       MsgBox "El Valor Ingresado no es un Valor Válido para Póliza.", vbCritical, "Error de Datos"
       Txt_Poliza.SetFocus
       Exit Sub
    End If

    If Opt_Poliza.Value = True Then
        
       vlPunt = 1
       Msf_Grilla.Col = 0
       While vlPunt < Msf_Grilla.Rows
             Msf_Grilla.Row = vlPunt
             If Trim(Msf_Grilla.Text) = Trim(Txt_Poliza.Text) Then
                If Msf_Grilla.Rows = 2 Then
                   Call flInicializaGrilla
                Else
                    Msf_Grilla.RemoveItem vlPunt
                End If
             Else
                 vlPunt = (vlPunt + 1)
             End If
       Wend
    Else
    
        If Cmb_PenNumIdent.Text = "" Then
           MsgBox "Debe Ingresar el Tipo de identificación.", vbCritical, "Error de Datos"
           Cmb_PenNumIdent.SetFocus
           Exit Sub
        End If
        
        If (Trim(Txt_PenNumIdent.Text)) = "" Then
           MsgBox "Debe Ingresar Identificación.", vbCritical, "Error de Datos"
           Txt_PenNumIdent.SetFocus
           Exit Sub
        Else
            ''*Txt_Rut = Format(Txt_Rut, "##,###,##0")
            Txt_PenNumIdent = UCase(Trim(Txt_PenNumIdent))
            Txt_PenNumIdent.SetFocus
        End If
     
''        If Not ValiRut(Txt_Rut.Text, Txt_Digito.Text) Then
''           MsgBox "El Rut Ingresado es incorrecto.", vbCritical, "Error de Datos"
''           Txt_Rut.SetFocus
''           Exit Sub
''        End If
        
        vlPunt = 1
        
        While vlPunt < Msf_Grilla.Rows
              Msf_Grilla.Row = vlPunt
              Msf_Grilla.Col = 0
              If (Trim(Msf_Grilla.Text) = Trim(Txt_Poliza.Text)) Then
                Msf_Grilla.Col = 3
                vlPos = InStr(Msf_Grilla.Text, "-")
                vlTipoIden = Trim(Msf_Grilla.Text) ''*Trim(Mid(Msf_Grilla.Text, 1, vlPos - 1))
                ''*vlRut = Format((Trim(vlRut)), "#######0")
                Msf_Grilla.Col = 4
                vlNumIden = Trim(Msf_Grilla.Text)
                
                If Trim(vlTipoIden) = Trim(Cmb_PenNumIdent.Text) And _
                   Trim(vlNumIden) = Trim(Txt_PenNumIdent.Text) Then
                    If Msf_Grilla.Rows = 2 Then
                       Call flInicializaGrilla
                    Else
                        Msf_Grilla.RemoveItem vlPunt
                    End If
                 Else
                     vlPunt = (vlPunt + 1)
                 End If
              Else
                  vlPunt = (vlPunt + 1)
              End If
        Wend
    End If
    
    Call Cmd_Cancelar_Click
    Call Opt_Poliza_Click
        
Exit Sub
Err_Restar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

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

Private Sub Cmd_Sumar_Click()
On Error GoTo Err_Sumar
Dim vlTipoI As String

  vlRutAux = ""
  vlSwModificar = "N"
  
  
  'Validación de Número de póliza y/o rut ingresados
  If Opt_Poliza.Value = True Then
     Txt_Poliza = Trim(UCase(Txt_Poliza))
     If Txt_Poliza.Text = "" Then
        MsgBox "Debe ingresar el Nº de Póliza.", vbCritical, "Error de Datos"
        Txt_Poliza.SetFocus
        Exit Sub
     Else
         Txt_Poliza.Text = Trim(Txt_Poliza.Text)
     End If
  Else
      If ((Trim(Cmb_PenNumIdent.Text)) = "") Or (Txt_PenNumIdent.Text = "") Then
      ''*Or _ (Not ValiRut(Txt_Rut.Text, Txt_Digito.Text))
         MsgBox "Debe Ingresar la Identificación del Pensionado.", vbCritical, "Error de Datos"
         Txt_Poliza.SetFocus
         Exit Sub
     Else
         ''*Txt_Rut = Format(Txt_Rut, "##,###,##0")
         Txt_PenNumIdent = UCase(Trim(Txt_PenNumIdent))
         Txt_PenNumIdent.SetFocus
           
     End If
  End If
    
vlCodTipoIdenBenCau = fgObtenerCodigo_TextoCompuesto(Cmb_PenNumIdent)
vlNumIdenBenCau = Trim(UCase(Txt_PenNumIdent))

'--------------------------------------------------------------------
'Validación de Existencia de Número de Póliza y/o Rut, según corresponda
  vgPalabra = ""
  ''*vlRutAux = Format(Txt_Rut, "#0")
  If (Txt_Poliza.Text <> "") And (Cmb_PenNumIdent.Text <> "") And (Txt_PenNumIdent.Text <> "") Then
      vgPalabra = "num_poliza = '" & Trim(Txt_Poliza.Text) & "' AND "
      vgPalabra = vgPalabra & "cod_tipoidenBEN = " & (vlCodTipoIdenBenCau) & " and "
      vgPalabra = vgPalabra & "num_idenben = '" & (vlNumIdenBenCau) & "' "
  Else
      If (Txt_Poliza.Text <> "") Then
         vgPalabra = "num_poliza = " & Trim(Txt_Poliza.Text) & " "
      Else
          vgPalabra = "cod_tipoidenBEN = " & (vlCodTipoIdenBenCau) & " "
          vgPalabra = vgPalabra & "AND num_idenben = '" & (vlNumIdenBenCau) & "' "
      End If
  End If
    
  'Validar Existencia de Poliza - Identificación
  vgSql = ""
  vgSql = "SELECT num_poliza "
  vgSql = vgSql & "FROM PP_TMAE_BEN WHERE "
  vgSql = vgSql & vgPalabra
  Set vgRegistro = vgConexionBD.Execute(vgSql)
  If vgRegistro.EOF Then
     MsgBox "La Póliza o Identificación Ingresados No Existen en la Base de Datos.", vbInformation, "Información"
     Exit Sub
  End If
'--------------------------------------------------------------------
        
    vgPalabra = ""
    
    If Opt_Poliza.Value = True Then
       If Txt_Poliza.Text <> "" Then
           vlNumPoliza = Trim(Txt_Poliza.Text)
           vgPalabra = "num_poliza = '" & Txt_Poliza.Text & "' AND "
           vgPalabra = vgPalabra & "cod_estpension <> '" & clCodSinDerPen & "' "
       End If
    Else
        ''*vlRutAux = Format(Txt_Rut, "#0")
        If (Txt_Poliza.Text <> "") And (Cmb_PenNumIdent.Text <> "") And (Txt_PenNumIdent.Text <> "") Then
            vlNumPoliza = Trim(Txt_Poliza.Text)
            vgPalabra = "num_poliza = '" & Txt_Poliza.Text & "' AND "
            vgPalabra = vgPalabra & "cod_tipoidenben = " & vlCodTipoIdenBenCau & " AND "
            vgPalabra = vgPalabra & "num_idenben = '" & vlNumIdenBenCau & "' "
        Else
            If (Cmb_PenNumIdent.Text <> "") And (Txt_PenNumIdent.Text <> "") Then
               vlcont = 0
               vgSql = ""
               vgSql = "SELECT DISTINCT num_poliza "
               vgSql = vgSql & "FROM PP_TMAE_BEN WHERE "
               vgSql = vgSql & "cod_tipoidenben = " & (vlCodTipoIdenBenCau) & " AND "
               vgSql = vgSql & "num_idenben = '" & (vlNumIdenBenCau) & "' AND "
               vgSql = vgSql & "cod_estpension <> '" & clCodSinDerPen & "' "
               Set vgRegistro = vgConexionBD.Execute(vgSql)
               If Not vgRegistro.EOF Then
                  While Not vgRegistro.EOF
                        vlcont = vlcont + 1
                        vgRegistro.MoveNext
                  Wend
               End If
               If vlcont > 0 Then
                   If vlcont = 1 Then
                       ''*vlRutAux = Format(Txt_Rut, "#0")
                       vgPalabra = "cod_tipoidenben = " & vlCodTipoIdenBenCau & " AND "
                       vgPalabra = vgPalabra & "num_idenben = '" & vlNumIdenBenCau & "' AND "
                       vgPalabra = vgPalabra & "cod_estpension <> '" & clCodSinDerPen & "' "
                       vgRegistro.MoveFirst
                       vlNumPoliza = (vgRegistro!num_poliza)
                   Else
                       Call flMostrarMensajeCantidad
                       Txt_Poliza.SetFocus
                       Exit Sub
                   End If
               Else
                   vgSql = ""
                   vgSql = "SELECT DISTINCT num_poliza "
                   vgSql = vgSql & "FROM PP_TMAE_BEN WHERE "
                   vgSql = vgSql & "cod_tipoidenben = " & vlCodTipoIdenBenCau & " AND "
                   vgSql = vgSql & "num_idenben = '" & vlNumIdenBenCau & "' "
                   vgSql = vgSql & "ORDER BY num_poliza "
                   Set vgRegistro = vgConexionBD.Execute(vgSql)
                   If Not vgRegistro.EOF Then
                      'MsgBox "El Rut Ingresado No tiene Derecho a Pensión, en la(s) Póliza(s) a la(s) que se encuentra Asociado.", vbInformation, "Información"
                      vlNumPoliza = Trim(vgRegistro!num_poliza)
                      vgPalabra = "num_poliza = '" & Trim(vgRegistro!num_poliza) & "' AND "
                   End If
                   
                    ''*vgPalabra = vgPalabra & "rut_ben = " & vlRutAux & " "
                    vgPalabra = vgPalabra & "cod_tipoidenben = " & vlCodTipoIdenBenCau & " AND "
                    vgPalabra = "num_idenben = '" & vlNumIdenBenCau & "' "
               End If
            End If
        End If
        
    End If
    
'-----------------------------------------------------------------------
'Ejecutar selección según los parámetros correspondientes, contenidos en
'variable vgpalabra
    
    vlFechaPeriodo = ""
    vlFechaPeriodo = DateSerial(Trim(Txt_Anno.Text), Trim(Txt_Mes.Text), "1")
    
    vgSql = ""
    vgSql = "SELECT num_poliza,num_endoso,num_orden,cod_tipoidenben,num_idenben, "
    vgSql = vgSql & " gls_nomben,gls_patben,gls_matben,cod_par,cod_estpension "
    vgSql = vgSql & " FROM PP_TMAE_BEN "
    vgSql = vgSql & " WHERE "
    vgSql = vgSql & vgPalabra
    vgSql = vgSql & " AND (num_endoso = "
    vgSql = vgSql & " (SELECT MAX(num_endoso) FROM PP_TMAE_POLIZA where num_poliza = '" & vlNumPoliza & "')) "
    vgSql = vgSql & " ORDER BY num_orden ASC "
    Set vgRs = vgConexionBD.Execute(vgSql)
    
    If vgRs.EOF Then
       MsgBox "El Número de Póliza o Rut Ingresados no tienen Derecho a Pensión.", vbInformation, "Información"
       Exit Sub
    End If

    
    'Valida Vigencia de la Póliza a la fecha ingresada en periodo
    If Not fgValidaVigenciaPoliza(Trim(vgRs!num_poliza), Trim(vlFechaPeriodo)) Then
       MsgBox "El Periodo Ingresado no se Encuentra dentro del Rango de Vigencia de la Póliza. " & Chr(13) & _
              "                       No se Agregará el Registro de Mensaje.", vbInformation, "Información"
       If Opt_Poliza.Value = True Then
          Txt_Poliza.SetFocus
       Else
           Cmb_PenNumIdent.SetFocus
       End If
       Exit Sub
    End If
    
    If Not fgValidaPagoPension(vlFechaPeriodo, Trim(vgRs!num_poliza), (vgRs!Num_Orden)) Then
       MsgBox " Ya se ha Realizado el Proceso de Cálculo de Pensión para ésta Fecha ", vbCritical, "Operación Cancelada"
       Exit Sub
    End If

    
    While Not vgRs.EOF

          If Not vgRs.EOF Then
             'Valida Derecho a Pensión del Beneficiario de la Póliza
             If Trim(vgRs!Cod_EstPension) = Trim(clCodSinDerPen) Then
                MsgBox " El Beneficiario de la Póliza No Tiene Derecho a Pensión " & Chr(13) & _
                       "          No se Agregará el Registro de Mensaje.", vbInformation, "Información"
                
        '          If Opt_Poliza.Value = True Then
        '             Txt_Poliza.SetFocus
        '          Else
        '              Txt_Rut.SetFocus
        '          End If
        '          Exit Sub
               vgRs.MoveNext
             Else
        
                 If Msf_Grilla.Rows > 1 Then
    
                'Verificar Existencia del Detalle a Ingresar en la Grilla
                    vlPunt = 1
                    Msf_Grilla.Row = vlPunt 'Msf_Grilla.Rows - 1
                    While vlPunt < Msf_Grilla.Rows
                          Msf_Grilla.Row = vlPunt
                          Msf_Grilla.Col = 0
                          If Trim(Msf_Grilla.Text) = Trim(vgRs!num_poliza) Then
                             Msf_Grilla.Col = 3
                             vlPos = InStr(Msf_Grilla.Text, "-")
                             vlTipoIden = Trim(Mid(Msf_Grilla.Text, 1, vlPos - 1))
                             ''*vlRut = Format((Trim(vlRut)), "#######0")
                             Msf_Grilla.Col = 4
                             vlNumIden = Trim(Msf_Grilla.Text)
                             
                             If Trim(vlTipoIden) = Trim(vgRs!Cod_TipoIdenBen) And _
                                Trim(vlNumIden) = Trim(vgRs!Num_IdenBen) Then
                                If Msf_Grilla.Row = (Msf_Grilla.Rows - 1) Then
                                   If Msf_Grilla.Row = 1 Then
                                      Call flInicializaGrilla
                                      vlPunt = Msf_Grilla.Rows
                                   Else
                                       Msf_Grilla.RemoveItem vlPunt
                                   End If
                                Else
                                    Msf_Grilla.RemoveItem vlPunt
                                End If
                             Else
                                 vlPunt = (vlPunt + 1)
                             End If
                          Else
                              vlPunt = (vlPunt + 1)
                          End If
                    Wend
                 End If
                   
                vlTipoI = " " & Trim(vgRs!Cod_TipoIdenBen) & " - " & fgBuscarNombreTipoIden(vgRs!Cod_TipoIdenBen, False)
                   
                'Agregar detalle en la grilla
                 Msf_Grilla.AddItem CStr(Trim(vgRs!num_poliza)) & vbTab _
                 & Trim(vgRs!num_endoso) & vbTab _
                 & Trim(vgRs!Num_Orden) & vbTab _
                 & (vlTipoI) & vbTab & (Trim(vgRs!Num_IdenBen)) & vbTab _
                 & ((Trim(vgRs!Gls_NomBen)) & " " & (Trim(vgRs!Gls_PatBen)) & " " & (Trim(vgRs!Gls_MatBen))) & vbTab _
                 & Trim(vgRs!Cod_Par)
    
                 vgRs.MoveNext
            
             End If
               
          Else
              MsgBox "El Número de Póliza o la Identificación Ingresados no tienen Derecho a Pensión.", vbInformation, "Información"
              Exit Sub
          End If
          
    Wend
    
    Call Cmd_Cancelar_Click
    Call Opt_Poliza_Click

Exit Sub
Err_Sumar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Form_Load()
On Error GoTo Err_Cargar

    Frm_MenIndividual.Top = 0
    Frm_MenIndividual.Left = 0
    
    Call flInicializaGrilla
    Call flCargaCombo(Cmb_CodMensaje)
    fgComboTipoIdentificacion Cmb_PenNumIdent
    Call flDeshabilitaIngreso
            
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Msf_Grilla_DblClick()
Dim vlTipoI As String
On Error GoTo Err_GrillaDblClick
    
    If vlSwEstPeriodo = "A" Then
    
        Msf_Grilla.Col = 0
        If (Msf_Grilla.Text = "") Or (Msf_Grilla.Row = 0) Then
            MsgBox "No existen Detalles", vbExclamation, "Información"
            Exit Sub
        Else
            If Opt_Poliza.Value = True Then
               Msf_Grilla.Col = 0
               Txt_Poliza.Text = Msf_Grilla.Text
               Txt_Poliza.SetFocus
               
            Else
                Msf_Grilla.Col = 0
                Txt_Poliza.Text = Msf_Grilla.Text
                Msf_Grilla.Col = 3
                vlPos = InStr(Msf_Grilla.Text, "-")
                vlTipoI = Trim(Mid(Msf_Grilla.Text, 1, vlPos - 1))
                Call fgBuscarPosicionCodigoCombo(vlTipoI, Cmb_PenNumIdent)
                Msf_Grilla.Col = 4
                Txt_PenNumIdent.Text = Trim(Msf_Grilla.Text)
''                Txt_Digito.Text = Trim(Mid(Msf_Grilla.Text, vlPos + 1, 2))
                Txt_PenNumIdent.SetFocus
            End If
            
        End If
    End If

Exit Sub
Err_GrillaDblClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Opt_Poliza_Click()

    Opt_Poliza.Value = True
    If (Cmb_PenNumIdent.ListCount <> 0) Then
        Cmb_PenNumIdent.ListIndex = 0
    End If
    Txt_PenNumIdent.Text = ""
    Txt_Poliza.Text = ""
    Txt_Poliza.Enabled = True
    Cmb_PenNumIdent.Enabled = False
    Cmb_PenNumIdent.Enabled = False
    Txt_Poliza.SetFocus

End Sub

Private Sub Opt_Rut_Click()

    Opt_Rut.Value = True
    If (Cmb_PenNumIdent.ListCount <> 0) Then
        Cmb_PenNumIdent.ListIndex = 0
    End If
    Txt_PenNumIdent.Text = ""
    ''Txt_Digito.Text = ""
    Txt_Poliza.Text = ""
'    Txt_Poliza.Enabled = False
    Cmb_PenNumIdent.Enabled = True
    Txt_PenNumIdent.Enabled = True
    'Txt_Rut.SetFocus
    Cmb_PenNumIdent.SetFocus

End Sub

Private Sub Txt_Anno_Change()

On Error GoTo Err_AnnoChange

    If Not IsNumeric(Txt_Anno) Then
       Txt_Anno = ""
    End If
    
Exit Sub
Err_AnnoChange:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub

Private Sub Txt_Anno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

    If Txt_Anno.Text = "" Then
       MsgBox "Debe Ingresar Año del Periodo.", vbCritical, "Error de Datos"
       Txt_Anno.SetFocus
       Exit Sub
    End If
    
    vlFecha = Date
    vlFecha = Format(CDate(Trim(vlFecha)), "yyyymmdd")
    vlAnno = (Mid(vlFecha, 1, 4))
    Txt_Anno.Text = Format(Txt_Anno.Text, "0000")
        
    If CDbl(Txt_Anno.Text) < 1900 Or CDbl(Txt_Anno.Text) > vlAnno Then
       MsgBox "Debe Ingresar un Año Mayor a 1900 o Menor Igual al Actual.", vbCritical, "Error de Datos"
       Txt_Anno.SetFocus
       Exit Sub
    End If
        
    Cmb_CodMensaje.SetFocus
End If
End Sub

Private Sub Txt_Mes_Change()

On Error GoTo Err_MesChange

    If Not IsNumeric(Txt_Mes) Then
       Txt_Mes = ""
    End If
    
Exit Sub
Err_MesChange:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
    
End Sub

Private Sub Txt_Mes_KeyPress(KeyAscii As Integer)

   If KeyAscii = 13 Then
       If Txt_Mes.Text = "" Then
          MsgBox "Debe Ingresar Mes del Periodo.", vbCritical, "Error de Datos"
          Txt_Mes.SetFocus
          Exit Sub
       End If
         
       If CDbl(Txt_Mes.Text) <= 0 Or CDbl(Txt_Mes.Text) > 12 Then
          MsgBox "El Mes Ingresado no es un Valor Válido.", vbCritical, "Error de Datos"
          Txt_Mes.SetFocus
          Exit Sub
       End If
       
       Txt_Mes.Text = Format(Txt_Mes.Text, "00")
       Txt_Anno.SetFocus
       
        
    End If

End Sub

Private Sub Txt_PenNumIdent_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txt_PenNumIdent = UCase(Trim(Txt_PenNumIdent))
        If Txt_PenNumIdent.Text = "" Then
           MsgBox "Debe ingresar el Nº de Identificación.", vbCritical, "Error de Datos"
           Txt_PenNumIdent.SetFocus
           Exit Sub
        End If
        Cmd_Sumar.SetFocus
    End If
End Sub

Private Sub txt_pennumident_lostfocus()
    Txt_PenNumIdent = Trim(UCase(Txt_PenNumIdent))
End Sub

Private Sub Txt_Poliza_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Txt_Poliza = UCase(Trim(Txt_Poliza))
    If Txt_Poliza.Text = "" Then
       MsgBox "Debe ingresar el Nº de Póliza.", vbCritical, "Error de Datos"
       Txt_Poliza.SetFocus
       Exit Sub
    End If
    Txt_Poliza = Format(Txt_Poliza, "0000000000")
    Cmd_Sumar.SetFocus
    If Cmb_PenNumIdent.Enabled = True Then
       Cmb_PenNumIdent.SetFocus
    End If
    
End If
End Sub

Private Sub Txt_Poliza_LostFocus()
If (Trim(Txt_Poliza) <> "") Then
    Txt_Poliza = UCase(Trim(Txt_Poliza))
    Txt_Poliza = Format(Txt_Poliza, "0000000000")
End If
End Sub

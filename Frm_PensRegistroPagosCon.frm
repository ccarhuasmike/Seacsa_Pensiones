VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "Crystl32.OCX"
Begin VB.Form Frm_PensRegistroPagosCon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Pagos a Terceros"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   9975
   Begin VB.Frame Fra_Botones 
      Height          =   975
      Left            =   120
      TabIndex        =   18
      Top             =   5760
      Width           =   9735
      Begin VB.CommandButton Cmd_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   4305
         Picture         =   "Frm_PensRegistroPagosCon.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   200
         Width           =   730
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   3120
         Picture         =   "Frm_PensRegistroPagosCon.frx":05DA
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   195
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5520
         Picture         =   "Frm_PensRegistroPagosCon.frx":0C94
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   200
         Width           =   720
      End
      Begin Crystal.CrystalReport Rpt_Imprimir 
         Left            =   8280
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.Frame Fra_RangoFec 
      Caption         =   "5337,57"
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
      Height          =   4335
      Left            =   120
      TabIndex        =   16
      Top             =   1320
      Width           =   9735
      Begin MSFlexGridLib.MSFlexGrid Msf_GrillaCM 
         Height          =   1095
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   1931
         _Version        =   393216
         Rows            =   1
         Cols            =   11
         BackColor       =   14745599
         AllowUserResizing=   1
      End
      Begin MSFlexGridLib.MSFlexGrid Msf_GrillaPG 
         Height          =   1575
         Left            =   120
         TabIndex        =   22
         Top             =   2520
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   2778
         _Version        =   393216
         Rows            =   1
         Cols            =   11
         BackColor       =   14745599
         AllowUserResizing=   1
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         Index           =   1
         X1              =   120
         X2              =   9600
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Pago Garantizado"
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
         Index           =   7
         Left            =   120
         TabIndex        =   21
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Cuota Mortuoria"
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
         Index           =   9
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Lbl_Nombre 
         Alignment       =   2  'Center
         Caption         =   "Detalle de Pagos"
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
         Index           =   6
         Left            =   120
         TabIndex        =   17
         Top             =   0
         Width           =   1815
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
      Top             =   120
      Width           =   9735
      Begin VB.TextBox Txt_PenPoliza 
         Height          =   285
         Left            =   840
         MaxLength       =   10
         TabIndex        =   1
         Top             =   360
         Width           =   1185
      End
      Begin VB.TextBox Txt_PenNumIdent 
         Height          =   285
         Left            =   5400
         MaxLength       =   16
         TabIndex        =   3
         Top             =   360
         Width           =   1875
      End
      Begin VB.CommandButton Cmd_BuscarPol 
         Height          =   375
         Left            =   9000
         Picture         =   "Frm_PensRegistroPagosCon.frx":0D8E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Buscar Póliza"
         Top             =   240
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
         Left            =   9000
         Picture         =   "Frm_PensRegistroPagosCon.frx":0E90
         TabIndex        =   7
         ToolTipText     =   "Buscar"
         Top             =   600
         Width           =   615
      End
      Begin VB.ComboBox Cmb_PenNumIdent 
         BackColor       =   &H00E0FFFF&
         Height          =   315
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   2235
      End
      Begin VB.Label Lbl_PenNombre 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   720
         Width           =   7935
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Póliza"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Ident."
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "N° End"
         Height          =   195
         Index           =   42
         Left            =   7680
         TabIndex        =   13
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Lbl_End 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   8280
         TabIndex        =   4
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Lbl_Nombre 
         Alignment       =   2  'Center
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
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   0
         Width           =   1935
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   855
      End
   End
End
Attribute VB_Name = "Frm_PensRegistroPagosCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim vlCodTipoIden As String
Dim vlNombre As String, vlNombreSeg As String
Dim vlApMaterno As String
Dim vlDep As String, vlProv As String, vlDist As String
Dim vlViaPago As String, vlFecDev As String
Dim vlFecRecep As String, vlFecPago As String
Dim vlValCob As Double, vlValPag As Double
Dim vlTipDcto As String
Dim vlParent As String
Dim vlTasaInt As Double
Dim vlFecIni As String, vlFecFin As String
Dim vlMesTran As String, vlMesNoDev As String
Dim vlPenNoPer As Double
                             
Private Sub Cmd_Imprimir_Click()
    If (Fra_Poliza.Enabled = False) Then
        If (Msf_GrillaCM.Rows > 1) Or (Msf_GrillaPG.Rows > 1) Then
            flImpresion
        Else
           MsgBox "No Existen Información de Pagos para la Póliza Seleccionada.", vbCritical, "Operación Cancelada"
           'Txt_PenPoliza.SetFocus
           Exit Sub
        End If
    End If
End Sub

'--------------------- Número de Póliza ---------------------
Private Sub Txt_PenPoliza_KeyPress(KeyAscii As Integer)
On Error GoTo Err_TxtPenPolizaKeyPress

    If KeyAscii = 13 Then
        If Trim(Txt_PenPoliza.Text) = "" Then
          'MsgBox "Debe Ingresar Número de Póliza.", vbCritical, "Error de Datos"
          'Txt_PenPoliza.SetFocus
          'Exit Sub
        End If
        Txt_PenPoliza = UCase(Trim(Txt_PenPoliza))
        Txt_PenPoliza = Format(Txt_PenPoliza, "0000000000")
        Cmb_PenNumIdent.SetFocus
    End If
    
Exit Sub
Err_TxtPenPolizaKeyPress:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub
Private Sub Txt_PenPoliza_LostFocus()
    Txt_PenPoliza = UCase(Trim(Txt_PenPoliza))
    Txt_PenPoliza = Format(Txt_PenPoliza, "0000000000")
End Sub

'--------------------- Tipo Identificación ---------------------
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

'--------------------- Número Identificación ---------------------
Private Sub Txt_PenNumIdent_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If (Trim(Txt_PenNumIdent) <> "") Then
            Txt_PenNumIdent = UCase(Trim(Txt_PenNumIdent))
        End If
        Cmd_BuscarPol.SetFocus
    End If
End Sub
Private Sub txt_pennumident_lostfocus()
        Txt_PenNumIdent = Trim(UCase(Txt_PenNumIdent))
End Sub

Private Sub Form_Load()
On Error GoTo Err_Cargar

    Me.Top = 0
    Me.Left = 0

    fgComboTipoIdentificacion Cmb_PenNumIdent
    
    Call flInicializaGrillaCM
    Call flInicializaGrillaPG
    
    Call flDeshabilitarIngreso
    
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_BuscarPol_Click()
Dim vlNombreSeg As String, vlApMaterno As String
On Error GoTo Err_CmdBuscarPolClick
        
    vlSwSeleccionado = False
            
    If Txt_PenPoliza.Text = "" Then
       If ((Trim(Cmb_PenNumIdent.Text)) = "") Or (Txt_PenNumIdent.Text = "") Then
           MsgBox "Debe Ingresar el Número de Póliza o la Identificación del Pensionado.", vbCritical, "Error de Datos"
           Txt_PenPoliza.SetFocus
           Exit Sub
       Else
           Txt_PenNumIdent = UCase(Trim(Txt_PenNumIdent))
           Txt_PenNumIdent.SetFocus
       End If
    Else
        Txt_PenPoliza.Text = Trim(Txt_PenPoliza.Text)
    End If
    
    vlCodTipoIdenBenCau = fgObtenerCodigo_TextoCompuesto(Cmb_PenNumIdent)
    vlNumIdenBenCau = Txt_PenNumIdent
    
    vgPalabra = ""
    'Seleccionar beneficiario, según número de póliza y rut de beneficiario.
    If (Txt_PenPoliza.Text) And (Cmb_PenNumIdent.Text <> "") And (Txt_PenNumIdent.Text <> "") Then
        vgPalabra = "num_poliza = '" & Txt_PenPoliza.Text & "' AND "
        vgPalabra = vgPalabra & "cod_tipoidenBEN = " & (vlCodTipoIdenBenCau) & " and "
        vgPalabra = vgPalabra & "num_idenben = '" & (vlNumIdenBenCau) & "' "
    Else
        'Seleccionar, según número de póliza, el primer beneficiario con derecho a pensión.
        'En caso de no existir, seleccionar sólo el primer beneficiario sin derecho.
        If Txt_PenPoliza.Text <> "" Then
           vgSql = ""
           vgSql = "SELECT COUNT(num_orden) as NumeroBen "
           vgSql = vgSql & "FROM PP_TMAE_BEN WHERE "
           vgSql = vgSql & "num_poliza = '" & Txt_PenPoliza.Text & "' "
           ''**vgSql = vgSql & "AND cod_estpension <> '" & clCodSinDerPen & "' "
           vgSql = vgSql & "ORDER BY num_endoso DESC, num_orden ASC "
           Set vgRegistro = vgConexionBD.Execute(vgSql)
           If (vgRegistro!numeroben) <> 0 Then
              vgPalabra = "num_poliza = '" & Txt_PenPoliza.Text & "' "
              ''**vgPalabra = vgPalabra & "AND cod_estpension <> '" & clCodSinDerPen & "' "
           Else
               vgPalabra = "num_poliza = '" & Txt_PenPoliza.Text & "' "
           End If
        Else
            'Seleccionar beneficiario, según rut beneficiario. (Datos de primera póliza encontrada.)
            If Txt_PenNumIdent.Text <> "" Then
               ''vlRutAux = Format(Txt_PenNumIdent, "#0")
               vgPalabra = "cod_tipoidenBEN = " & (vlCodTipoIdenBenCau) & " "
               vgPalabra = vgPalabra & "AND num_idenben = '" & (vlNumIdenBenCau) & "' "
            End If
        End If
    End If
    
    'Ejecutar selección según los parámetros correspondientes, contenidos en
    'variable vgpalabra
    vgSql = ""
    vgSql = "SELECT num_endoso,num_orden,gls_nomben,gls_nomsegben,gls_patben,gls_matben,"
    vgSql = vgSql & "cod_estpension,cod_tipoidenben,num_idenben,num_poliza "
    vgSql = vgSql & "FROM PP_TMAE_BEN WHERE "
    vgSql = vgSql & vgPalabra
    vgSql = vgSql & " ORDER BY num_endoso DESC,num_orden ASC "
    Set vgRs2 = vgConexionBD.Execute(vgSql)
    If Not vgRs2.EOF Then
        
        vlAfp = fgObtenerPolizaCod_AFP(vgRs2!num_poliza, CStr(vgRs2!num_endoso))
        
       If Trim(vgRs2!Cod_EstPension) = Trim(clCodSinDerPen) Then '* debe preg esto
          MsgBox " El Beneficiario Seleccionado No Tiene Derecho a Pensión " & Chr(13) & _
          "          Sólo podrá Consultar los Datos del Registro", vbInformation, "Información"

          'Desactivar Todos los Controles del Formulario
            Call flDeshabilitarIngreso
            Fra_Poliza.Enabled = False

       Else
           If Not fgValidaVigenciaPoliza(Trim(Txt_PenPoliza.Text), fgBuscaFecServ) Then
              MsgBox " La Póliza Ingresada no se Encuentra Vigente en el Sistema " & Chr(13) & _
                     "      Sólo podrá Consultar los Datos del Registro", vbInformation, "Información"
              
              'Desactivar Todos los Controles del Formulario
                Call flDeshabilitarIngreso
                Fra_Poliza.Enabled = False
           Else
               Call flHabilitarIngreso
           End If
       End If
              
        vlCodTipoIdenBenCau = vgRs2!Cod_TipoIdenBen
        vlNumIdenBenCau = Trim(vgRs2!Num_IdenBen)
             
        Txt_PenPoliza.Text = Trim(vgRs2!num_poliza)
        Call fgBuscarPosicionCodigoCombo(vlCodTipoIdenBenCau, Cmb_PenNumIdent)
        Txt_PenNumIdent.Text = vlNumIdenBenCau
        vlNombreSeg = IIf(IsNull(vgRs2!Gls_NomSegBen), "", Trim(vgRs2!Gls_NomSegBen))
        vlApMaterno = IIf(IsNull(vgRs2!Gls_MatBen), "", Trim(vgRs2!Gls_MatBen))
        Lbl_PenNombre.Caption = fgFormarNombreCompleto(Trim(vgRs2!Gls_NomBen), vlNombreSeg, Trim(vgRs2!Gls_PatBen), vlApMaterno)
        Lbl_End.Caption = (vgRs2!num_endoso)
        vlNumEndoso = (vgRs2!num_endoso)
        vlNumOrden = (vgRs2!Num_Orden)
        
        Call flCargaGrillaCM
        Call flCargaGrillaPG
        
        If (Msf_GrillaCM.Rows = 1) And (Msf_GrillaPG.Rows = 1) Then
           MsgBox "No Existen Información de Pagos para la Póliza Seleccionada.", vbCritical, "Operación Cancelada"
           Exit Sub
        End If
    Else
        MsgBox "El Beneficiario o la Póliza Ingresados, No Existen en la Base de Datos", vbInformation, "Información"
        Txt_PenPoliza.SetFocus
        Exit Sub
    End If
    vgRs2.Close
    
    Cmd_Imprimir.SetFocus
    vlSwMostrar = False

       
Exit Sub
Err_CmdBuscarPolClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Cancelar_Click()
    
    Txt_PenPoliza.Text = ""
    If (Cmb_PenNumIdent.ListCount > 0) Then
        Cmb_PenNumIdent.ListIndex = 0
    End If
    Txt_PenNumIdent.Text = ""
    Lbl_End.Caption = ""
    Lbl_PenNombre.Caption = ""
    Msf_GrillaCM.Rows = 1
    Msf_GrillaPG.Rows = 1
    
    Call flDeshabilitarIngreso
    
    Txt_PenPoliza.SetFocus
    
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

Private Sub Cmd_Buscar_Click()
On Error GoTo Err_CmdBuscarClick

    Frm_Busqueda.flInicio ("Frm_PensRegistroPagosCon")
    
Exit Sub
Err_CmdBuscarClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Function flRecibe(vlNumPoliza, vlCodTipoIden, vlNumIden, vlNumEndoso)
    Txt_PenPoliza = vlNumPoliza
    Call fgBuscarPosicionCodigoCombo(vlCodTipoIden, Cmb_PenNumIdent)
    Txt_PenNumIdent = vlNumIden
    Lbl_End = vlNumEndoso
    Cmd_BuscarPol_Click
End Function

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

Private Sub flInicializaGrillaCM()
    
    Msf_GrillaCM.Clear
    Msf_GrillaCM.Cols = 17
    Msf_GrillaCM.Rows = 1
    
    Msf_GrillaCM.Row = 0
        
    Msf_GrillaCM.Col = 0
    Msf_GrillaCM.Text = "Nº Póliza"
    Msf_GrillaCM.ColWidth(0) = 1200
    Msf_GrillaCM.ColAlignment(0) = 3
    
    Msf_GrillaCM.Col = 1
    Msf_GrillaCM.Text = "Nº Endoso"
    Msf_GrillaCM.ColWidth(1) = 900
    
    Msf_GrillaCM.Col = 2
    Msf_GrillaCM.Text = "Tipo Ident."
    Msf_GrillaCM.ColWidth(2) = 1200
    
    Msf_GrillaCM.Col = 3
    Msf_GrillaCM.Text = "Nº Ident."
    Msf_GrillaCM.ColWidth(3) = 1200
    
    Msf_GrillaCM.Col = 4
    Msf_GrillaCM.Text = "Nombre"
    Msf_GrillaCM.ColWidth(4) = 3000
    
    Msf_GrillaCM.Col = 5
    Msf_GrillaCM.Text = "Domicilio"
    Msf_GrillaCM.ColWidth(5) = 2000
      
    Msf_GrillaCM.Col = 6
    Msf_GrillaCM.Text = "Departamento"
    Msf_GrillaCM.ColWidth(6) = 1800
    
    Msf_GrillaCM.Col = 7
    Msf_GrillaCM.Text = "Provincia"
    Msf_GrillaCM.ColWidth(7) = 1800
    
    Msf_GrillaCM.Col = 8
    Msf_GrillaCM.Text = "Distrito"
    Msf_GrillaCM.ColWidth(8) = 1800
    
    Msf_GrillaCM.Col = 9
    Msf_GrillaCM.Text = "Vía de Pago"
    Msf_GrillaCM.ColWidth(9) = 2300
    
    Msf_GrillaCM.Col = 10
    Msf_GrillaCM.Text = "Fec. Recep."
    Msf_GrillaCM.ColWidth(10) = 1100
    
    Msf_GrillaCM.Col = 11
    Msf_GrillaCM.Text = "Fec. Pago"
    Msf_GrillaCM.ColWidth(11) = 1100
    
    Msf_GrillaCM.Col = 12
    Msf_GrillaCM.Text = "Val. Cobrado"
    Msf_GrillaCM.ColWidth(12) = 1200
    
    Msf_GrillaCM.Col = 13
    Msf_GrillaCM.Text = "Val. Pagado"
    Msf_GrillaCM.ColWidth(13) = 1200
    
    Msf_GrillaCM.Col = 14
    Msf_GrillaCM.Text = "Tipo Dcto. Pago"
    Msf_GrillaCM.ColWidth(14) = 2000
    
    Msf_GrillaCM.Col = 15
    Msf_GrillaCM.Text = "Nº Dcto. Pago"
    Msf_GrillaCM.ColWidth(15) = 1200
    
    Msf_GrillaCM.Col = 16
    Msf_GrillaCM.Text = "RUC Funeraria"
    Msf_GrillaCM.ColWidth(16) = 1200
    
End Sub

Private Sub flInicializaGrillaPG()
    
    Msf_GrillaPG.Clear
    Msf_GrillaPG.Cols = 20
    Msf_GrillaPG.Rows = 1
    
    Msf_GrillaPG.Row = 0
        
    Msf_GrillaPG.Col = 0
    Msf_GrillaPG.Text = "Nº Póliza"
    Msf_GrillaPG.ColWidth(0) = 1200
    Msf_GrillaPG.ColAlignment(0) = 3
    
    Msf_GrillaPG.Col = 1
    Msf_GrillaPG.Text = "Nº Endoso"
    Msf_GrillaPG.ColWidth(1) = 900
    
    Msf_GrillaPG.Col = 2
    Msf_GrillaPG.Text = "Tipo Ident."
    Msf_GrillaPG.ColWidth(2) = 1200
    
    Msf_GrillaPG.Col = 3
    Msf_GrillaPG.Text = "Nº Ident."
    Msf_GrillaPG.ColWidth(3) = 1200
    
    Msf_GrillaPG.Col = 4
    Msf_GrillaPG.Text = "Nombre"
    Msf_GrillaPG.ColWidth(4) = 3000
    
    Msf_GrillaPG.Col = 5
    Msf_GrillaPG.Text = "Domicilio"
    Msf_GrillaPG.ColWidth(5) = 2000
    
    Msf_GrillaPG.Col = 6
    Msf_GrillaPG.Text = "Departamento"
    Msf_GrillaPG.ColWidth(6) = 1800
    
    Msf_GrillaPG.Col = 7
    Msf_GrillaPG.Text = "Provincia"
    Msf_GrillaPG.ColWidth(7) = 1800
    
    Msf_GrillaPG.Col = 8
    Msf_GrillaPG.Text = "Distrito"
    Msf_GrillaPG.ColWidth(8) = 1800
    
    Msf_GrillaPG.Col = 9
    Msf_GrillaPG.Text = "Vía de Pago"
    Msf_GrillaPG.ColWidth(9) = 2300
    
    Msf_GrillaPG.Col = 10
    Msf_GrillaPG.Text = "Parentesco"
    Msf_GrillaPG.ColWidth(10) = 1200
    
    Msf_GrillaPG.Col = 11
    Msf_GrillaPG.Text = "Fec. Recep."
    Msf_GrillaPG.ColWidth(11) = 1200
    
    Msf_GrillaPG.Col = 12
    Msf_GrillaPG.Text = "Fec. Pago"
    Msf_GrillaPG.ColWidth(12) = 1200
    
    Msf_GrillaPG.Col = 13
    Msf_GrillaPG.Text = "Tasa Interes"
    Msf_GrillaPG.ColWidth(13) = 1200
    
    Msf_GrillaPG.Col = 14
    Msf_GrillaPG.Text = "Fec. Dev."
    Msf_GrillaPG.ColWidth(14) = 1200
    
    Msf_GrillaPG.Col = 15
    Msf_GrillaPG.Text = "Ini. Per. Gar."
    Msf_GrillaPG.ColWidth(15) = 1200
    
    Msf_GrillaPG.Col = 16
    Msf_GrillaPG.Text = "Fin Per. Gar."
    Msf_GrillaPG.ColWidth(16) = 1200
    
    Msf_GrillaPG.Col = 17
    Msf_GrillaPG.Text = "Meses Transc."
    Msf_GrillaPG.ColWidth(17) = 1200
    
    Msf_GrillaPG.Col = 18
    Msf_GrillaPG.Text = "Meses No Dev."
    Msf_GrillaPG.ColWidth(18) = 1200
    
    Msf_GrillaPG.Col = 19
    Msf_GrillaPG.Text = "Pen. No Percib."
    Msf_GrillaPG.ColWidth(19) = 1200
    
End Sub


Private Sub flHabilitarIngreso()

    Fra_Poliza.Enabled = False
    Fra_RangoFec.Enabled = True
    Msf_GrillaCM.Enabled = True
    Msf_GrillaPG.Enabled = True
    
End Sub

Private Sub flDeshabilitarIngreso()

    Fra_Poliza.Enabled = True
    Fra_RangoFec.Enabled = False
    Msf_GrillaCM.Enabled = False
    Msf_GrillaPG.Enabled = False
    
End Sub

Private Function flCargaGrillaCM()

    Call flInicializaGrillaCM
    
    vgSql = ""
    vgSql = "SELECT num_poliza,num_endoso,cod_tipoidensolicita,num_idensolicita,"
    vgSql = vgSql & "gls_nomsolicita,gls_nomsegsolicita,gls_patsolicita,"
    vgSql = vgSql & "gls_matsolicita,gls_dirsolicita,cod_direccion,cod_viapago,"
    vgSql = vgSql & "fec_solpago,fec_pago,mto_cobra,mto_pago,cod_tipodctopago,"
    vgSql = vgSql & "num_dctopago,cod_tipoidenfun,num_idenfun "
    vgSql = vgSql & "FROM PP_TMAE_PAGTERCUOMOR WHERE "
    vgSql = vgSql & "cod_conpago ='" & cgPagoTerceroCuoMor & "' and "
    vgSql = vgSql & "num_poliza = '" & Trim(Txt_PenPoliza) & "' "
    vgSql = vgSql & "ORDER BY num_endoso DESC "
    Set vgRs3 = vgConexionBD.Execute(vgSql)
    If Not vgRs3.EOF Then
    
        vlCodTipoIden = (vgRs3!cod_tipoidensolicita & "-" & fgBuscarNombreTipoIden(vgRs3!cod_tipoidensolicita, False))
        vlNombreSeg = IIf(IsNull(vgRs3!gls_nomsegsolicita), "", Trim(vgRs3!gls_nomsegsolicita))
        vlApMaterno = IIf(IsNull(vgRs3!gls_matsolicita), "", Trim(vgRs3!gls_matsolicita))
        vlNombre = fgFormarNombreCompleto(Trim(vgRs3!gls_nomsolicita), vlNombreSeg, Trim(vgRs3!gls_patsolicita), vlApMaterno)
        fgBuscarNombreProvinciaRegion (vgRs3!Cod_Direccion)
        vlDep = vgNombreRegion
        vlProv = vgNombreProvincia
        vlDist = vgNombreComuna
        vlViaPago = (vgRs3!Cod_ViaPago & "-" & fgBuscarGlosaElemento(vgCodTabla_ViaPago, vgRs3!Cod_ViaPago))
        vlFecRecep = DateSerial(Mid((vgRs3!fec_solpago), 1, 4), Mid((vgRs3!fec_solpago), 5, 2), Mid((vgRs3!fec_solpago), 7, 2))
        vlFecPago = DateSerial(Mid((vgRs3!Fec_Pago), 1, 4), Mid((vgRs3!Fec_Pago), 5, 2), Mid((vgRs3!Fec_Pago), 7, 2))
        vlValCob = Format(vgRs3!mto_cobra, "#,#0.00")
        vlValPag = Format(vgRs3!mto_pago, "#,#0.00")
        vlTipDcto = (vgRs3!cod_tipodctopago & "-" & fgBuscarGlosaElemento(vgCodTabla_TipDoc, vgRs3!cod_tipodctopago))
        
        Msf_GrillaCM.AddItem Trim(vgRs3!num_poliza) & vbTab & _
                             Trim(vgRs3!num_endoso) & vbTab & _
                             " " & Trim(vlCodTipoIden) & vbTab & _
                             Trim(vgRs3!num_idensolicita) & vbTab & _
                             Trim(vlNombre) & vbTab & _
                             Trim(vgRs3!gls_dirsolicita) & vbTab & _
                             Trim(vlDep) & vbTab & Trim(vlProv) & vbTab & Trim(vlDist) & vbTab & _
                             " " & Trim(vlViaPago) & vbTab & _
                             Trim(vlFecRecep) & vbTab & Trim(vlFecPago) & vbTab & _
                             Trim(vlValCob) & vbTab & Trim(vlValPag) & vbTab & _
                             Trim(vlTipDcto) & vbTab & _
                             Trim(vgRs3!num_dctopago) & vbTab & _
                             Trim(vgRs3!num_idenfun)
    End If
    vgRs3.Close
        
End Function


Private Function flCargaGrillaPG()
    
    Call flInicializaGrillaPG
    
    vgSql = ""
    vgSql = "SELECT d.num_poliza,d.num_endoso,d.cod_tipoidenben,d.num_idenben,"
    vgSql = vgSql & "b.gls_nomben,b.gls_nomsegben,b.gls_patben,b.gls_matben,"
    vgSql = vgSql & "d.gls_dirben,d.cod_direccion,d.cod_viapago,b.cod_par,"
    vgSql = vgSql & "p.fec_solpago,d.fec_pago,p.prc_tasaint,po.fec_dev,"
    vgSql = vgSql & "p.fec_inipergarpag,p.fec_finpergar,"
    vgSql = vgSql & "d.num_mesespag,d.num_mesesnodev,d.mto_pago "
    vgSql = vgSql & "FROM PP_TMAE_PAGTERGAR P,PP_TMAE_PAGTERGARBEN D,PP_TMAE_BEN B,PP_TMAE_POLIZA PO "
    vgSql = vgSql & "WHERE d.num_poliza=b.num_poliza and d.num_orden=b.num_orden and "
    vgSql = vgSql & "d.num_endoso=b.num_endoso and p.num_poliza=d.num_poliza and "
    vgSql = vgSql & "p.num_endoso=d.num_endoso and p.fec_pago=d.fec_pago and "
    vgSql = vgSql & "po.num_poliza=p.num_poliza and po.num_endoso=p.num_endoso and "
    vgSql = vgSql & "cod_conpago ='" & cgPagoTerceroPerGar & "' and "
    vgSql = vgSql & "d.num_poliza = '" & Trim(Txt_PenPoliza) & "' "
    vgSql = vgSql & "ORDER BY num_endoso DESC "
    Set vgRs3 = vgConexionBD.Execute(vgSql)
    While Not vgRs3.EOF
    
        vlCodTipoIden = (vgRs3!Cod_TipoIdenBen & "-" & fgBuscarNombreTipoIden(vgRs3!Cod_TipoIdenBen, False))
        vlNombreSeg = IIf(IsNull(vgRs3!Gls_NomSegBen), "", Trim(vgRs3!Gls_NomSegBen))
        vlApMaterno = IIf(IsNull(vgRs3!Gls_MatBen), "", Trim(vgRs3!Gls_MatBen))
        vlNombre = fgFormarNombreCompleto(Trim(vgRs3!Gls_NomBen), vlNombreSeg, Trim(vgRs3!Gls_PatBen), vlApMaterno)
        fgBuscarNombreProvinciaRegion (vgRs3!Cod_Direccion)
        vlDep = vgNombreRegion
        vlProv = vgNombreProvincia
        vlDist = vgNombreComuna
        vlViaPago = (vgRs3!Cod_ViaPago & "-" & fgBuscarGlosaElemento(vgCodTabla_ViaPago, vgRs3!Cod_ViaPago))
        vlParent = (vgRs3!Cod_Par & "-" & fgBuscarGlosaElemento(vgCodTabla_Par, vgRs3!Cod_Par))
        vlFecRecep = DateSerial(Mid((vgRs3!fec_solpago), 1, 4), Mid((vgRs3!fec_solpago), 5, 2), Mid((vgRs3!fec_solpago), 7, 2))
        vlFecPago = DateSerial(Mid((vgRs3!Fec_Pago), 1, 4), Mid((vgRs3!Fec_Pago), 5, 2), Mid((vgRs3!Fec_Pago), 7, 2))
        vlTasaInt = Format(vgRs3!prc_tasaint, "#,#0.00")
        vlFecDev = DateSerial(Mid((vgRs3!fec_dev), 1, 4), Mid((vgRs3!fec_dev), 5, 2), Mid((vgRs3!fec_dev), 7, 2))
        vlFecIni = DateSerial(Mid((vgRs3!fec_inipergarpag), 1, 4), Mid((vgRs3!fec_inipergarpag), 5, 2), Mid((vgRs3!fec_inipergarpag), 7, 2))
        vlFecFin = DateSerial(Mid((vgRs3!fec_finpergar), 1, 4), Mid((vgRs3!fec_finpergar), 5, 2), Mid((vgRs3!fec_finpergar), 7, 2))
        vlMesTran = Format(vgRs3!num_mesespag, "#,#0")
        vlMesNoDev = Format(vgRs3!num_mesesnodev, "#,#0")
        vlPenNoPer = Format(vgRs3!mto_pago, "#,#0.00")
        
        Msf_GrillaPG.AddItem Trim(vgRs3!num_poliza) & vbTab & _
                             Trim(vgRs3!num_endoso) & vbTab & _
                             " " & Trim(vlCodTipoIden) & vbTab & _
                             Trim(vgRs3!Num_IdenBen) & vbTab & _
                             Trim(vlNombre) & vbTab & _
                             Trim(vgRs3!Gls_DirBen) & vbTab & _
                             Trim(vlDep) & vbTab & Trim(vlProv) & vbTab & Trim(vlDist) & vbTab & _
                             " " & Trim(vlViaPago) & vbTab & _
                             Trim(vlParent) & vbTab & _
                             Trim(vlFecRecep) & vbTab & Trim(vlFecPago) & vbTab & _
                             Trim(vlTasaInt) & vbTab & Trim(vlFecDev) & vbTab & _
                             Trim(vlFecIni) & vbTab & Trim(vlFecFin) & vbTab & _
                             Trim(vlMesTran) & vbTab & Trim(vlMesNoDev) & vbTab & _
                             Trim(vlPenNoPer)
        vgRs3.MoveNext

    Wend
    vgRs3.Close
        
End Function

Sub flImpresion()
Dim vlArchivo As String
Dim vlTIdent, vlNIdent, vlNom As String
Err.Clear
On Error GoTo Errores1
   
   Screen.MousePointer = 11
   
   vlArchivo = strRpt & "PP_Rpt_ConPagoTer.rpt"   '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
       MsgBox "Archivo de Reporte de Listados no se encuentra en el directorio de la Aplicación.", 16, "Archivo no encontrado"
       Screen.MousePointer = 0
       Exit Sub
   End If
  
    'Busca Información del causante
    vgSql = "Select cod_tipoidenben,num_idenben,gls_nomben,gls_nomsegben,gls_patben,"
    vgSql = vgSql & "gls_matben from pp_tmae_ben Where cod_par='99' and "
    vgSql = vgSql & "num_poliza = '" & Trim(Txt_PenPoliza) & "' "
    vgSql = vgSql & "order by num_endoso desc"
    Set vlRegistro = vgConexionBD.Execute(vgSql)
    If Not vlRegistro.EOF Then
        vlTIdent = fgBuscarNombreTipoIden(vlRegistro!Cod_TipoIdenBen, False)
        vlNIdent = vlRegistro!Num_IdenBen
        vlNom = vlRegistro!Gls_NomBen + " " + IIf(IsNull(vlRegistro!Gls_NomSegBen), "", (vlRegistro!Gls_NomSegBen)) + " " + vlRegistro!Gls_PatBen + " " + IIf(IsNull(vlRegistro!Gls_MatBen), "", (vlRegistro!Gls_MatBen))
    End If
    vlRegistro.Close
  
   vgQuery = ""
   vgQuery = "{PP_TMAE_PAGTERGARBEN.NUM_POLIZA} = '" & Trim(Txt_PenPoliza) & "' AND "
   vgQuery = vgQuery & "{MA_TPAR_TABCODViaPago.COD_TABLA} = '" & vgCodTabla_ViaPago & "' "
 
   Rpt_Imprimir.Reset
   Rpt_Imprimir.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
   Rpt_Imprimir.Connect = vgRutaDataBase
   Rpt_Imprimir.SelectionFormula = vgQuery
   Rpt_Imprimir.Formulas(0) = ""
   Rpt_Imprimir.Formulas(1) = ""
   Rpt_Imprimir.Formulas(2) = ""
   Rpt_Imprimir.Formulas(3) = ""
   Rpt_Imprimir.Formulas(4) = ""
   Rpt_Imprimir.Formulas(5) = ""
   Rpt_Imprimir.Formulas(6) = ""
   
   Rpt_Imprimir.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
   Rpt_Imprimir.Formulas(1) = "NombreSistema= '" & vgNombreSistema & "'"
   Rpt_Imprimir.Formulas(2) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
   Ident = Trim(vlTIdent) + "-" + Trim(vlNIdent)
   Rpt_Imprimir.Formulas(3) = "Poliza = '" & Trim(Txt_PenPoliza) & "'"
   Rpt_Imprimir.Formulas(4) = "Endoso = '" & Trim(Lbl_End) & "'"
   Rpt_Imprimir.Formulas(5) = "Identificacion = '" & Trim(Ident) & "'"
   Rpt_Imprimir.Formulas(6) = "Nombre = '" & Trim(vlNom) & "'"

   Rpt_Imprimir.SubreportToChange = ""
   Rpt_Imprimir.SubreportToChange = "PP_Rpt_SubConPagoCM.rpt"
   Rpt_Imprimir.Connect = vgRutaDataBase

   vgQuery = ""
   vgQuery = "{PP_TMAE_PAGTERCUOMOR.NUM_POLIZA} = '" & Trim(Txt_PenPoliza) & "' "
   vgQuery = vgQuery & " AND {MA_TPAR_TABCODViaPago.COD_TABLA} = '" & vgCodTabla_ViaPago & "' AND "
   vgQuery = vgQuery & "{MA_TPAR_TABCODTipDcto.COD_TABLA} = '" & vgCodTabla_TipDoc & "' "

   Rpt_Imprimir.SelectionFormula = vgQuery

   Rpt_Imprimir.Destination = crptToWindow
   Rpt_Imprimir.WindowState = crptMaximized
   Rpt_Imprimir.WindowTitle = "Informe de Pagos a Terceros"
   Rpt_Imprimir.Action = 1
   
   Rpt_Imprimir.SubreportToChange = ""
   
   Screen.MousePointer = 0
   
Exit Sub
Errores1:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " " & "Numero de Error : " & Err.Number & Err.HelpFile & Err.HelpContext, vbExclamation, "Mensaje de error ..."
        Screen.MousePointer = 0
    End If
End Sub


Sub flImpresionCM()
Dim vlArchivo As String
Dim vlTIdent, vlNIdent, vlNom As String
Err.Clear
On Error GoTo Errores1
   
   Screen.MousePointer = 11
   
   vlArchivo = strRpt & "PP_Rpt_ConPagoCM.rpt"   '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
       MsgBox "Archivo de Reporte de Listados no se encuentra en el directorio de la Aplicación.", 16, "Archivo no encontrado"
       Screen.MousePointer = 0
       Exit Sub
   End If
  
''    'Busca Información del causante
''    vgSql = "Select cod_tipoidenben,num_idenben,gls_nomben,gls_nomsegben,gls_patben,"
''    vgSql = vgSql & "gls_matben from pp_tmae_ben Where cod_par='99' and "
''    vgSql = vgSql & "num_poliza = '" & Trim(Txt_PenPoliza) & "' "
''    vgSql = vgSql & "order by num_endoso desc"
''    Set vlRegistro = vgConexionBD.Execute(vgSql)
''    If Not vlRegistro.EOF Then
''        vlTIdent = fgBuscarNombreTipoIden(vlRegistro!Cod_TipoIdenBen, False)
''        vlNIdent = vlRegistro!Num_IdenBen
''        vlNom = vlRegistro!Gls_NomBen + " " + IIf(IsNull(vlRegistro!Gls_NomBen), "", (vlRegistro!Gls_NomBen)) + " " + vlRegistro!Gls_PatBen + " " + IIf(IsNull(vlRegistro!Gls_MatBen), "", (vlRegistro!Gls_MatBen))
''    End If
''    vlRegistro.Close
  
   vgQuery = ""
   vgQuery = "{PP_TMAE_PAGTERCUOMOR.NUM_POLIZA} = '" & Trim(Txt_PenPoliza) & "' AND "
   vgQuery = vgQuery & "{MA_TPAR_TABCODViaPago.COD_TABLA} = '" & vgCodTabla_ViaPago & "' AND "
   vgQuery = vgQuery & "{MA_TPAR_TABCODTipDcto.COD_TABLA} = '" & vgCodTabla_TipDoc & "' "

   Rpt_Imprimir.Reset
   Rpt_Imprimir.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
   Rpt_Imprimir.Connect = vgRutaDataBase
   Rpt_Imprimir.SelectionFormula = vgQuery
   Rpt_Imprimir.Formulas(0) = ""
   Rpt_Imprimir.Formulas(1) = ""
   Rpt_Imprimir.Formulas(2) = ""
   Rpt_Imprimir.Formulas(3) = ""
   Rpt_Imprimir.Formulas(4) = ""
   Rpt_Imprimir.Formulas(5) = ""
   Rpt_Imprimir.Formulas(6) = ""
   
   Rpt_Imprimir.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
   Rpt_Imprimir.Formulas(1) = "NombreSistema= '" & vgNombreSistema & "'"
   Rpt_Imprimir.Formulas(2) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
   Ident = Trim(Cmb_PenNumIdent) + "-" + Trim(Txt_PenNumIdent)
   Rpt_Imprimir.Formulas(3) = "Poliza = '" & Trim(Txt_PenPoliza) & "'"
   Rpt_Imprimir.Formulas(4) = "Endoso = '" & Trim(Lbl_End) & "'"
   Rpt_Imprimir.Formulas(5) = "Identificacion = '" & Trim(Ident) & "'"
   Rpt_Imprimir.Formulas(6) = "Nombre = '" & Trim(Lbl_PenNombre) & "'"
   
   Rpt_Imprimir.Destination = crptToWindow
   Rpt_Imprimir.WindowState = crptMaximized
   Rpt_Imprimir.WindowTitle = "Informe Pagos a Terceros Gastos de Sepelio"
   Rpt_Imprimir.Action = 1
   
   Screen.MousePointer = 0
   
Exit Sub
Errores1:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " " & "Numero de Error : " & Err.Number & Err.HelpFile & Err.HelpContext, vbExclamation, "Mensaje de error ..."
        Screen.MousePointer = 0
    End If
End Sub


VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmCertSuperv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Certificado de Supervivencia"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6150
   Icon            =   "frmCertSuperv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmd_Salir 
      Caption         =   "&Salir"
      Height          =   765
      Left            =   3945
      Picture         =   "frmCertSuperv.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5985
      Width           =   690
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos a Insertar o Actualizar"
      Height          =   1920
      Left            =   150
      TabIndex        =   13
      Top             =   2010
      Width           =   5880
      Begin VB.TextBox txtInstitución 
         Height          =   285
         Left            =   1920
         TabIndex        =   25
         Top             =   1485
         Width           =   3615
      End
      Begin VB.TextBox txtFechaPeriodoEfect 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1170
         Width           =   1275
      End
      Begin VB.TextBox txtFechaIngreso 
         BackColor       =   &H8000000F&
         Height          =   300
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   855
         Width           =   1275
      End
      Begin VB.TextBox txtFechaRecepcion 
         BackColor       =   &H8000000F&
         Height          =   300
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   525
         Width           =   1275
      End
      Begin VB.TextBox txtFechaTermino 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3795
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   210
         Width           =   1320
      End
      Begin VB.TextBox txtFechaVigencia 
         Height          =   285
         Left            =   1920
         TabIndex        =   19
         Top             =   225
         Width           =   1275
      End
      Begin VB.Label Label10 
         Caption         =   "----"
         Height          =   165
         Left            =   3390
         TabIndex        =   21
         Top             =   285
         Width           =   210
      End
      Begin VB.Label Label9 
         Caption         =   "Institución :"
         Height          =   225
         Left            =   195
         TabIndex        =   18
         Top             =   1545
         Width           =   930
      End
      Begin VB.Label Label8 
         Caption         =   "Periodo de Efecto :"
         Height          =   210
         Left            =   195
         TabIndex        =   17
         Top             =   1230
         Width           =   1425
      End
      Begin VB.Label Label7 
         Caption         =   "Fecha de Ingreso :"
         Height          =   195
         Left            =   195
         TabIndex        =   16
         Top             =   915
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha de Recepción :"
         Height          =   180
         Left            =   195
         TabIndex        =   15
         Top             =   585
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha de Vigencia :"
         Height          =   225
         Left            =   210
         TabIndex        =   14
         Top             =   255
         Width           =   1515
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Imprimir"
      Height          =   765
      Left            =   2685
      Picture         =   "frmCertSuperv.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5985
      Width           =   750
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Grabar"
      Height          =   780
      Left            =   1440
      Picture         =   "frmCertSuperv.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5985
      Width           =   750
   End
   Begin MSFlexGridLib.MSFlexGrid Msf_Grilla 
      Height          =   1950
      Left            =   105
      TabIndex        =   10
      Top             =   3990
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   3440
      _Version        =   393216
      Rows            =   1
      Cols            =   6
      FixedCols       =   3
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Afiliado"
      Height          =   1965
      Left            =   150
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.TextBox txtBeneficiario 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1665
         TabIndex        =   9
         Top             =   1470
         Width           =   3825
      End
      Begin VB.TextBox txtAfiliado 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1665
         TabIndex        =   7
         Top             =   1095
         Width           =   3825
      End
      Begin VB.ComboBox cboTipoPension 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   660
         Width           =   3810
      End
      Begin VB.OptionButton opt_beneficiario 
         Caption         =   "Beneficiario"
         Enabled         =   0   'False
         Height          =   240
         Left            =   3090
         TabIndex        =   3
         Top             =   315
         Width           =   1185
      End
      Begin VB.OptionButton opt_Afiliado 
         Caption         =   "Afiliado"
         Enabled         =   0   'False
         Height          =   210
         Left            =   1680
         TabIndex        =   2
         Top             =   330
         Width           =   1035
      End
      Begin VB.Label Label4 
         Caption         =   "Beneficiario :"
         Height          =   210
         Left            =   210
         TabIndex        =   8
         Top             =   1530
         Width           =   1065
      End
      Begin VB.Label Label3 
         Caption         =   "Afiliado :"
         Height          =   240
         Left            =   210
         TabIndex        =   6
         Top             =   1125
         Width           =   900
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo de Pensión :"
         Height          =   210
         Left            =   210
         TabIndex        =   4
         Top             =   720
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Identificación :"
         Height          =   225
         Left            =   240
         TabIndex        =   1
         Top             =   315
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmCertSuperv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim str_identificacion As String
Dim str_Afiliado As String
Dim str_Beneficiado As String
Dim str_Poliza As String
Dim str_Orden As String
Dim str_doc As String
Dim str_numdoc As String
Dim str_Endoso As String
Dim str_TipoPension As String
Dim str_doc_afiliado As String
Dim str_numdoc_afiliado As String


Dim vlResp As String
Dim vlFecFin As String
Dim vlFecEfe As String
Dim vlFecRecep As String
Dim vlFecIng As String
Dim vlAnnoFin As String
Dim vlFechaIni As String
Dim vlFechaTer As String
Dim vl_Institucion As String
Dim vlSwVigPol   As Boolean
Dim vlOp        As String
Dim vlIniVig As String
Dim vlFechaEfecto As String
Dim vlOperacion As String
Dim vlFechaMatrimonio As String
Dim vlCodTipoIdenBenCau As String
Dim vlNumIdenBenCau As String
Dim vlFechaFallecimiento As String
Dim vlSw As Boolean
Dim Sql As String
Dim vlPos        As Double

Dim vlRegistro As New ADODB.Recordset

Dim vlInicio As String
Dim vlTermino As String
Dim vlRecepcion As String
Dim vlIngreso As String
Dim vlEfecto As String
Dim vlAnno As String
Dim vlMes As String
Dim vlDia As String

Public Property Let Identificacion(ByVal vNewValue As String)
str_identificacion = vNewValue
End Property

Public Property Let Afiliado(ByVal vNewValue As String)
str_Afiliado = vNewValue
End Property

Public Property Let Beneficiado(ByVal vNewValue As String)
str_Beneficiado = vNewValue
End Property

Public Property Let Poliza(ByVal vNewValue As String)
str_Poliza = vNewValue
End Property

Public Property Let Orden(ByVal vNewValue As String)
str_Orden = vNewValue
End Property

Public Property Let Doc_afiliado(ByVal vNewValue As String)
str_doc_afiliado = vNewValue
End Property

Public Property Let NumDoc_afiliado(ByVal vNewValue As String)
str_numdoc_afiliado = vNewValue
End Property

Public Property Let Doc(ByVal vNewValue As String)
str_doc = vNewValue
End Property

Public Property Let NumDoc(ByVal vNewValue As String)
str_numdoc = vNewValue
End Property

Public Property Let Endoso(ByVal vNewValue As String)
str_Endoso = vNewValue
End Property

Public Property Let TipoPension(ByVal vNewValue As String)
str_TipoPension = vNewValue
End Property


Function flBuscaCert(iFecha)
'Dim vlFechaEfecto As Date
Dim fila As Integer
Dim vls As Integer
Dim vlI As Integer
Dim vlFecgrilla As String
On Error GoTo Err_buscavig

    If (iFecha <> "") Then
       fila = Msf_Grilla.Rows - 1
       vls = 0
       For vlI = 1 To fila
           Msf_Grilla.Row = vlI
           Msf_Grilla.Col = 0
           vlFecgrilla = Format(CDate(Msf_Grilla), "yyyymmdd")
           If vlFecgrilla = (iFecha) Then
              vls = 1
             ' JANM ----10/09/04
             
             'MAVG---COMENTO 08/04/2010
'               vlPos = vlI
'              Call flEncontroInf
              Exit For
           End If
       Next vlI

       If vls = 0 Then
          vlFechaEfecto = fgFechaEfectoReliq(txtFechaIngreso.Text, str_Poliza, CInt(str_Orden), vlFechaTer)
          'Lbl_Efecto = vlFechaEfecto
          'Txt_FecTerVig.SetFocus
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

Private Sub Cargar_nuevo()

On Error GoTo Err_Grabar
    
    
    Dim anio As String
    Dim ames As String
    Dim adia As String
    Dim newFec As String
    

    txtFechaIngreso.Text = fgBuscaFecServ
    anio = Mid(txtFechaVigencia.Text, 7, 4)
    ames = Mid(txtFechaVigencia.Text, 4, 2)
    adia = Mid(txtFechaVigencia.Text, 1, 2) '"01" 'f_dia_ultimo(anio, ames)
    newFec = adia & "/" & ames & "/" & anio
    vlFechaTer = DateAdd("m", 12, newFec)
    
    anio = Mid(vlFechaTer, 7, 4)
    ames = Mid(vlFechaTer, 4, 2)
    adia = Mid(vlFechaTer, 1, 2) 'f_dia_ultimo(anio, ames)
    vlFechaTer = adia & "/" & ames & "/" & anio
    
    vl_Institucion = "Certificado de Supervivencia Actualizada"
    
    If (flValidaFecha(txtFechaIngreso.Text) = True) Then
       'transforma la fecha al formato yyyymmdd
        'Screen.MousePointer = 11
        Call flBuscaCert(Trim(txtFechaIngreso.Text))
        vlFechaIni = Format(CDate(Trim(txtFechaIngreso.Text)), "yyyymmdd")
       'se valida que exista información para esa fecha en la BD
        
    End If
    
    'txtFechaVigencia.Text =
    txtFechaIngreso.Text = txtFechaVigencia.Text
    txtFechaTermino.Text = vlFechaTer
    txtFechaRecepcion.Text = fgBuscaFecServ
    txtFechaPeriodoEfect.Text = vlFechaEfecto
    txtInstitución.Text = vl_Institucion
    
    
Exit Sub
Err_Grabar:
    MsgBox "No pudo Cargar datos necesarios", vbInformation

End Sub

Private Sub Cargar()

On Error GoTo Err_Grabar
    
    
    Dim anio As String
    Dim ames As String
    Dim adia As String
    Dim newFec As String
     
    'txtFechaVigencia.SetFocus
     
    txtFechaIngreso.Text = fgBuscaFecServ
    anio = Mid(fgBuscaFecServ, 7, 4)
    ames = Mid(fgBuscaFecServ, 4, 2)
    adia = "01" 'f_dia_ultimo(anio, ames)
    newFec = adia & "/" & ames & "/" & anio
    vlFechaTer = DateAdd("m", 5, newFec)
    
    anio = Mid(vlFechaTer, 7, 4)
    ames = Mid(vlFechaTer, 4, 2)
    adia = f_dia_ultimo(anio, ames)
    vlFechaTer = adia & "/" & ames & "/" & anio
    
    vl_Institucion = "Certificado de Supervivencia Actualizada"
    
    If (flValidaFecha(txtFechaIngreso.Text) = True) Then
       'transforma la fecha al formato yyyymmdd
        'Screen.MousePointer = 11
        Call flBuscaCert(Trim(txtFechaIngreso.Text))
        vlFechaIni = Format(CDate(Trim(txtFechaIngreso.Text)), "yyyymmdd")
       'se valida que exista información para esa fecha en la BD
        
    End If
    
    txtFechaVigencia.Text = txtFechaIngreso.Text
    txtFechaTermino.Text = vlFechaTer
    txtFechaRecepcion.Text = txtFechaIngreso.Text
    txtFechaPeriodoEfect.Text = vlFechaEfecto
    txtInstitución.Text = vl_Institucion
    
   'Dim Search, Where   ' Declare variables.
   ' Get search string from user.
   'Search = InputBox("Enter text to be found:")
   'Where = InStr(txtFechaVigencia.Text ,  Search)   ' Find string in text.
   'If Where Then   ' If found,
      
      'txtFechaVigencia.SelStart = 1  ' set selection start and
      'txtFechaVigencia.SelLength = Len(txtFechaVigencia.Text)     ' set selection length.
   'Else
   '   MsgBox "String not found."   ' Notify user.
   'End If
    
Exit Sub
Err_Grabar:
    MsgBox "No pudo Cargar datos necesarios", vbInformation

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

Private Sub cmd_salir_Click()
Unload Me
End Sub

Private Sub Command1_Click()

Dim vlEdad As String
Dim vlAnnoBen As String
Dim vlDiferencia As String
Dim Mto_Edad As Double

Dim vl_Institucion As String
On Error GoTo Err_Grabar

If Command1.Caption = "Nuevo" Then
    Call Cargar
    Command1.Picture = LoadPicture(App.Path & "\Drive.ICO")
    Command1.Caption = "&Grabar"
    Exit Sub
End If

    vlSwVigPol = True

    If fgValidaVigenciaPoliza(Trim(str_Poliza), Trim(txtFechaIngreso.Text)) = False Then
       MsgBox "La fecha ingresada no se encuentra dentro del rango de vigencia de la póliza" & Chr(13) & _
              "o esta no vigente. No se ingresara ni modificara información. ", vbCritical, "Operación Cancelada"
       Screen.MousePointer = 0
       vlSwVigPol = False
       Exit Sub
    End If

    If (CDate(Format(vlFechaTer, "dd/mm/yyyy")) < CDate(Format(txtFechaIngreso.Text, "dd/mm/yyyy"))) Then
        MsgBox "La fecha de término es menor a la fecha de inicio de vigencia.", vbCritical, "Dato Incorrecto"
        Exit Sub
    End If


    vlOp = ""
'No existe Frecuencia para este Certificado ABV


         
    vlFecFin = vlFechaTer
    vlFecFin = Format(CDate(Trim(vlFecFin)), "yyyymmdd")
    vlAnnoFin = Mid(vlFecFin, 1, 4)
    vlFecRecep = txtFechaIngreso.Text
    vlFecRecep = Format(CDate(Trim(vlFecRecep)), "yyyymmdd")
    vlFecIng = txtFechaIngreso.Text
    vlFecIng = Format(CDate(Trim(vlFecIng)), "yyyymmdd")
    vlIniVig = txtFechaIngreso.Text
    vlIniVig = Format(CDate(Trim(vlIniVig)), "yyyymmdd")
    vlFecEfe = vlFechaEfecto
    vlFecEfe = Format(CDate(Trim(vlFecEfe)), "yyyymmdd")


    'Verificar Si el Beneficiario Existe y Esta Vivo
    vlOperacion = ""
    vlFechaMatrimonio = ""
    vlCodTipoIdenBenCau = fgObtenerCodigo_TextoCompuesto(str_doc)
    vlNumIdenBenCau = Trim(UCase(str_numdoc))
    

    vgSql = ""
    vgSql = "SELECT NUM_POLIZA,FEC_fallben,fec_nacben "
    vgSql = vgSql & " FROM PP_TMAE_BEN "
    vgSql = vgSql & " Where "
    vgSql = vgSql & " NUM_POLIZA =  '" & (str_Poliza) & "' AND "
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
    vgSql = vgSql & " NUM_POLIZA = '" & Trim(str_Poliza) & "' and "
    vgSql = vgSql & " NUM_ORDEN = " & Trim(str_Orden) & " and "
    vgSql = vgSql & " fec_inicer = '" & vlIniVig & "'"
    Set vlRegistro = vgConexionBD.Execute(vgSql)
    If vlRegistro.EOF Then
        vlOp = "I"
    Else
        vlOp = "A"
    End If
'    vlRegistro.Close

    If (vlOp = "A") Then
        vgRes = MsgBox("¿ Está seguro que desea modificar los datos ?", 4 + 32 + 256, "Operación de Actualización")
        If vgRes <> 6 Then
            Cmd_Salir.SetFocus
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If
    If (vlOp = "I") Then
        vlResp = 6 'MsgBox(" ¿ Está seguro que desea ingresar los Datos ?", 4 + 32 + 256, "Proceso de Ingreso de Datos")
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
               Sql = Sql & " NUM_ENDOSO = " & (str_Endoso) & ","
               'Sql = Sql & " COD_FRECUENCIA = '" & (vlCodFre) & "',"
               If (vl_Institucion <> "") Then
                    Sql = Sql & " GLS_NOMINSTITUCION = '" & Trim(vl_Institucion) & "',"
               Else
                    Sql = Sql & " GLS_NOMINSTITUCION = NULL,"
               End If
               Sql = Sql & " FEC_RECCIA = '" & (vlFecRecep) & "',"
               Sql = Sql & " COD_USUARIOMODI = '" & (vgUsuario) & "',"
               Sql = Sql & " FEC_MODI = '" & Format(Date, "yyyymmdd") & "',"
               Sql = Sql & " HOR_MODI = '" & Format(Time, "hhmmss") & "',"
               If vlIniVig <> vlFecEfe Then 'Corresponde Reliquidar
                    Sql = Sql & " COD_INDRELIQUIDAR = 'S'"
               Else
                    Sql = Sql & " COD_INDRELIQUIDAR = 'N'"
               End If
               Sql = Sql & " Where "
               Sql = Sql & " NUM_POLIZA = '" & Trim(str_Poliza) & "' and "
               Sql = Sql & " NUM_ORDEN = " & Trim(str_Orden) & " and "
               Sql = Sql & " fec_inicer = '" & (vlIniVig) & "'"
               vgConexionBD.Execute (Sql)
            Else
               vlRegistro.Close
               Screen.MousePointer = 0
               Call flLmpCert
'               Txt_FecIniVig.Enabled = True
'               Txt_FecIniVig.SetFocus
               Exit Sub
            End If
        Else
            
            Sql = ""
            Sql = "update pp_tmae_certificado set"
            Sql = Sql & " NUM_ENDOSO = " & (str_Endoso) & ","
            'Sql = Sql & " COD_FRECUENCIA = '" & (vlCodFre) & "',"
            If (txtInstitución.Text <> "") Then
                Sql = Sql & " GLS_NOMINSTITUCION = '" & Trim(txtInstitución.Text) & "',"
            Else
                 Sql = Sql & " GLS_NOMINSTITUCION = NULL,"
            End If
            Sql = Sql & " FEC_RECCIA = '" & (vlFecRecep) & "',"
            Sql = Sql & " COD_USUARIOMODI = '" & (vgUsuario) & "',"
            Sql = Sql & " FEC_MODI = '" & Format(Date, "yyyymmdd") & "',"
            Sql = Sql & " HOR_MODI = '" & Format(Time, "hhmmss") & "',"
            If vlIniVig <> vlFecEfe Then 'Corresponde Reliquidar
                 Sql = Sql & " COD_INDRELIQUIDAR = 'S'"
            Else
                 Sql = Sql & " COD_INDRELIQUIDAR = 'N'"
            End If
            Sql = Sql & " Where "
            Sql = Sql & " NUM_POLIZA = '" & Trim(str_Poliza) & "' and "
            Sql = Sql & " NUM_ORDEN = " & Trim(str_Orden) & " and "
            Sql = Sql & " fec_inicer = '" & (vlIniVig) & "'"
            vgConexionBD.Execute (Sql)
        End If
    Else
       'Inserta los Datos en la Tabla pp_tmae_certificado
        Call flComparaFechaIngresada
        If vlSw = False Then
           Call flLmpCert
'           Txt_FecIniVig.Enabled = True
'           Txt_FecIniVig.SetFocus
           Exit Sub
        Else
           Call flComparaFechaExistente
           If vlSw = False Then
              Call flLmpCert
'              Txt_FecIniVig.Enabled = True
'              Txt_FecIniVig.SetFocus
              Exit Sub
           End If
        End If

        Sql = ""
        Sql = "insert into pp_tmae_certificado ("
        Sql = Sql & "NUM_POLIZA,NUM_ORDEN,fec_inicer,"
        Sql = Sql & "fec_tercer,NUM_ENDOSO,"
        'Sql = Sql & "COD_FRECUENCIA,"
        Sql = Sql & "GLS_NOMINSTITUCION,"
        Sql = Sql & "FEC_RECCIA,FEC_INGCIA,FEC_EFECTO,COD_USUARIOCREA,"
        Sql = Sql & "FEC_CREA,HOR_CREA,COD_INDRELIQUIDAR"
        Sql = Sql & " "
        Sql = Sql & ") values ("
        Sql = Sql & "'" & Trim(str_Poliza) & "',"
        Sql = Sql & "" & Trim(str_Orden) & ","
        Sql = Sql & "'" & (vlIniVig) & "',"
        Sql = Sql & "'" & (vlFecFin) & "',"
        Sql = Sql & "" & (str_Endoso) & ","
        'Sql = Sql & "'" & (vlCodFre) & "',"
        If (txtInstitución.Text <> "") Then
            Sql = Sql & "'" & Trim(txtInstitución.Text) & "',"
        Else
            Sql = Sql & "NULL,"
        End If
        Sql = Sql & "'" & (vlFecRecep) & "',"
        Sql = Sql & "'" & (vlFecIng) & "',"
        Sql = Sql & "'" & (vlFecEfe) & "',"
        Sql = Sql & "'" & (vgUsuario) & "',"
        Sql = Sql & "'" & Format(Date, "yyyymmdd") & "',"
        Sql = Sql & "'" & Format(Time, "hhmmss") & "',"
        Sql = Sql & "'" & IIf(vlFecIng <> vlFecEfe, "S", "N") & "'"
        Sql = Sql & ")"
        vgConexionBD.Execute (Sql)
    End If
    vlRegistro.Close
    
    MsgBox "Proceso Terminado", vbInformation
    
    If (vlOp <> "") Then
        'Limpia los Datos de la Pantalla

        Call flLmpGrilla
        Call flCargaGrilla
'        flLmpCert
'        Txt_FecIniVig.Enabled = True
'        Txt_FecIniVig.SetFocus
    End If
    Command2.Enabled = True
    Screen.MousePointer = 0

Exit Sub
Err_Grabar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(2) & Err.Description & " ]"
    End Select
End Sub

Private Sub Command2_Click()


Nombre_Afiliado = txtAfiliado.Text
Nombre_Beneficiario = txtBeneficiario.Text
Fecha_Creacion = txtFechaRecepcion.Text
If Mid(Trim(str_identificacion), 1, 2) = "99" Then
    Ident_1 = "x"
    Ident_2 = ""
Else
    Ident_1 = ""
    Ident_2 = "x"
End If

num_poliza = str_Poliza
RangoFecha = txtFechaVigencia.Text & " al " & txtFechaTermino.Text

If Mid(Trim(str_TipoPension), 1, 2) = "04" Or Mid(Trim(str_TipoPension), 1, 2) = "05" Then
    Tipo_J = "x"
    Tipo_S = ""
    Tipo_I = ""
ElseIf Mid(Trim(str_identificacion), 1, 2) = "08" Or Mid(Trim(str_identificacion), 1, 2) = "09" Or Mid(Trim(str_identificacion), 1, 2) = "10" Or Mid(Trim(str_identificacion), 1, 2) = "11" Or Mid(Trim(str_identificacion), 1, 2) = "12" Then
    Tipo_S = "x"
    Tipo_I = ""
    Tipo_J = ""
ElseIf Mid(Trim(str_identificacion), 1, 2) = "06" Or Mid(Trim(str_identificacion), 1, 2) = "07" Then
    Tipo_I = "x"
    Tipo_J = ""
    Tipo_S = ""
End If

Tipo_num_documento_afiliado = Trim(Mid(str_doc_afiliado, InStr(1, str_doc_afiliado, "-") + 1, Len(str_doc_afiliado))) & " - " & str_numdoc_afiliado
If Len(txtBeneficiario.Text) > 0 Then
    Tipo_num_documento_beneficiario = Trim(Mid(str_doc, InStr(1, str_doc, "-") + 1, Len(str_doc))) & " - " & str_numdoc
Else
    Tipo_num_documento_beneficiario = ""
End If

Dim vlArchivo As String
Dim r_temp As ADODB.Recordset

On Error GoTo Errores1
   
   
   Screen.MousePointer = 11
   
   vlArchivo = strRpt & "PP_Rpt_CertificadoSuperv.rpt"   '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
       MsgBox "El reporte no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
       Screen.MousePointer = 0
   End If
   
   Screen.MousePointer = 0
     
sName_Reporte = "PP_Rpt_CertificadoSuperv.rpt"
frm_plantilla.Show 1
   
Errores1:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " " & "Numero de Error : " & Err.Number & Err.HelpFile & Err.HelpContext, vbExclamation, "Mensaje de error ..."
        Screen.MousePointer = 0
    End If
    
End Sub

Private Sub Form_Load()

    
 vlResp = ""
 vlFecFin = ""
 vlFecEfe = ""
 vlFecRecep = ""
 vlFecIng = ""
 vlAnnoFin = ""
 vlFechaIni = ""
 vlFechaTer = ""

If Mid(Trim(str_identificacion), 1, 2) = "99" Then
    opt_Afiliado.Value = True
Else
    opt_Afiliado.Value = False
End If
txtAfiliado.Text = str_Afiliado
txtBeneficiario.Text = str_Beneficiado
cboTipoPension.AddItem str_TipoPension
cboTipoPension.ListIndex = 0
cboTipoPension.Enabled = False
Command2.Enabled = False

Call flLmpGrilla
Call flCargaGrilla
Call Cargar
End Sub

Function flLmpGrilla()
    Msf_Grilla.Clear
    Msf_Grilla.Rows = 1
    Msf_Grilla.RowHeight(0) = 250
    Msf_Grilla.Row = 0

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

End Function
Function flCargaGrilla()
On Error GoTo Err_Carga
Dim vlCodTab As String
     vlCodTab = "FP"

     vgSql = ""
     vgSql = "SELECT c.NUM_POLIZA,c.NUM_ENDOSO,c.NUM_ORDEN,c.fec_inicer,c.fec_tercer,"
     vgSql = vgSql & " c.COD_FRECUENCIA,c.GLS_NOMINSTITUCION,c.FEC_RECCIA,c.FEC_INGCIA,"
     vgSql = vgSql & " c.FEC_EFECTO "
     vgSql = vgSql & " FROM pp_tmae_certificado C "
     vgSql = vgSql & " Where "
     vgSql = vgSql & " c.NUM_POLIZA = '" & Trim(str_Poliza) & "' AND "
     'CMV-20051124 I
     'No se debe utilizar el num_endoso para buscar los certificados de estudio
     'de una carga, ya que debe mostrar todos los certificados que la carga
     'haya tenido, no sólo los registros del endoso actual
'     vgSql = vgSql & " c.NUM_ENDOSO = " & Trim(vlNumEndoso) & " AND "
     'CMV-20051124 F
     vgSql = vgSql & " c.NUM_ORDEN = " & Trim(str_Orden) & " "
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

              Msf_Grilla.AddItem ((vlInicio) & vbTab & (vlTermino)) & vbTab & _
                                 (vlRegistro!GLS_NOMINSTITUCION) & vbTab & _
                                 (vlRecepcion) & vbTab & _
                                 (vlIngreso) & vbTab & _
                                 (vlEfecto)
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
Function flComparaFechaMod()
On Error GoTo Err_buscavig
Dim vlTxtTermino As String
Dim vlAnt As Integer
Dim vlI As Integer
Dim vlGrInicio As String
Dim vlGrTermino As String
Dim vlTxtInicio As String

       vlSw = True
       vlTxtTermino = Format(CDate(vlFechaTer), "yyyymmdd")
       vlAnt = vlPos
       vlPos = vlPos - 1
       For vlI = 1 To vlPos
           Msf_Grilla.Row = vlI
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
              Msf_Grilla.Row = vlI
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

'    txtFechaVigencia = ""
'    txtFechaTermino = ""
'    txtFechaRecepcion = ""
'    txtInstitución = ""
'    txtFechaIngreso = ""
'    txtFechaPeriodoEfect = ""

Exit Function
Err_lmp:
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

'MARCO ---09/03/2010
Dim vlGrInicioMes As String
Dim vlGrTerminoMes As String
'FIN MARCO

    If (vlFechaIni <> "") Then
       vlSw = True
       vlTxtInicio = Format(CDate(txtFechaVigencia.Text), "yyyymmdd")
       vlTxtTermino = Format(CDate(txtFechaTermino.Text), "yyyymmdd")

        'MARCO---09/03/2010
        Dim FECHA_I As Date
        Dim FECHA_F As Date
        FECHA_I = "01/" & Format(CDate(txtFechaVigencia.Text), "MM") & "/" & Format(CDate(txtFechaVigencia.Text), "YYYY")
        FECHA_F = DateAdd("M", 1, FECHA_I)
        FECHA_F = DateAdd("D", -1, FECHA_F)
        vlGrInicioMes = Format(CDate(txtFechaVigencia.Text), "yyyymm")
        vlGrTerminoMes = Format(CDate(txtFechaTermino.Text), "yyyymm")
        'FIN MARCO---09/03/2010
        
       fila = Msf_Grilla.Rows - 1

       For vlI = 1 To fila

           Msf_Grilla.Row = fila

           Msf_Grilla.Col = 0
           vlGrInicio = Format(CDate(Msf_Grilla.Text), "yyyymm") 'marco quito el dd

           Msf_Grilla.Col = 1
           vlGrTermino = Format(CDate(Msf_Grilla.Text), "yyyymm") 'marco quito el dd

            'marco ----09/03/2010----solo puede ver un certificado por mes
            If (vlGrInicioMes) = (vlGrInicio) Then
                MsgBox "Ya tiene registrado para este mes un Certificado de Supervivencia", vbInformation, "Error de Datos"
                Screen.MousePointer = 0
                vlSw = False
                Exit For
            End If
            'fin marco
            
            'marco comento lo siguiente
'           If (vlTxtInicio) >= (vlGrInicio) And (vlTxtInicio) <= (vlGrTermino) Then
'               MsgBox "Rango de Fechas ya se encuentran en un período de Vigencia", vbCritical, "Error de Datos"
'               Screen.MousePointer = 0
'               vlSw = False
'               Exit For
'           Else
'               If (vlTxtTermino) >= (vlGrInicio) And (vlTxtTermino) <= (vlGrTermino) Then
'                   MsgBox "Rango de Fechas ya se encuentran en un período de Vigencia", vbCritical, "Error de Datos"
'                   Screen.MousePointer = 0
'                   vlSw = False
'                   Exit For
'               End If
'           End If
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

'MARCO ---09/03/2010
Dim vlGrInicioMes As String
Dim vlGrTerminoMes As String
'FIN MARCO

    If (vlFechaIni <> "") Then
       vlSw = True
       vlTxtInicio = Format(CDate(txtFechaVigencia.Text), "yyyymmdd")
       vlTxtTermino = Format(CDate(txtFechaTermino.Text), "yyyymmdd")
        'MARCO---09/03/2010
        Dim FECHA_I As Date
        Dim FECHA_F As Date
        FECHA_I = "01/" & Format(CDate(txtFechaVigencia.Text), "MM") & "/" & Format(CDate(txtFechaVigencia.Text), "YYYY")
        FECHA_F = DateAdd("M", 1, FECHA_I)
        FECHA_F = DateAdd("D", -1, FECHA_F)
        vlGrInicioMes = Format(CDate(txtFechaVigencia.Text), "yyyymm")
        vlGrTerminoMes = Format(CDate(txtFechaTermino.Text), "yyyymm")
        'FIN MARCO---09/03/2010
        
       fila = Msf_Grilla.Rows - 1
       For vlI = 1 To fila

           Msf_Grilla.Row = fila

           Msf_Grilla.Col = 0
           vlGrInicio = Format(CDate(Msf_Grilla.Text), "yyyymmdd")

           Msf_Grilla.Col = 1
           vlGrTermino = Format(CDate(Msf_Grilla.Text), "yyyymmdd")
            'marco ----09/03/2010----solo puede ver un certificado por mes
            If (vlGrInicioMes) = (vlGrInicio) Then
                MsgBox "Ya tiene registrado para este mes un Certificado de Supervivencia", vbInformation, "Error de Datos"
                Screen.MousePointer = 0
                vlSw = False
                Exit For
            End If
            'fin marco
            
'           If (vlGrInicio) >= (vlTxtInicio) And (vlGrInicio) <= (vlTxtTermino) Then
'               MsgBox "Rango de Fechas ya se encuentran en un período de Vigencia", vbCritical, "Error de Datos"
'               Screen.MousePointer = 0
'               vlSw = False
'               Exit For
'           Else
'               If (vlGrTermino) >= (vlTxtInicio) And (vlGrTermino) <= (vlTxtTermino) Then
'                   MsgBox "Rango de Fechas ya se encuentran en un período de Vigencia", vbCritical, "Error de Datos"
'                   vlSw = False
'                   Screen.MousePointer = 0
'                   Exit For
'               End If
'           End If
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
Private Sub Msf_Grilla_Click()
On Error GoTo Err_Grilla
Dim vlI As Integer


    Msf_Grilla.Col = 0
    vlPos = Msf_Grilla.RowSel
    Msf_Grilla.Row = vlPos
    If (Msf_Grilla.Text = "") Or (Msf_Grilla.Row = 0) Then
        Exit Sub
    End If
    Screen.MousePointer = 11

    txtFechaVigencia.Text = (Msf_Grilla.Text)

    Msf_Grilla.Col = 1
    txtFechaTermino.Text = (Msf_Grilla.Text)

    Msf_Grilla.Col = 2
    txtInstitución.Text = (Msf_Grilla.Text)

    Msf_Grilla.Col = 3
    txtFechaRecepcion.Text = (Msf_Grilla.Text)

    Msf_Grilla.Col = 4
    txtFechaIngreso.Text = (Msf_Grilla.Text)

    Msf_Grilla.Col = 5
    txtFechaPeriodoEfect.Text = (Msf_Grilla.Text)

    'Deshabilitar la Fecha de Inicio de Vigencia de la Póliza
    'Txt_FecIniVig.Enabled = False
    Command1.Picture = LoadPicture(App.Path & "\DRAG3PG.ICO")
    Command1.Caption = "Nuevo"
    
    
    Screen.MousePointer = 0

Exit Sub
Err_Grilla:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub txtFechaVigencia_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then

    Call Cargar_nuevo
    
End If

End Sub

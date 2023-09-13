VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "Crystl32.OCX"
Begin VB.Form Frm_HabDesImportar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carga de Haberes y Descuentos desde Archivo"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   8985
   Begin VB.Frame Fra_Operacion 
      Height          =   1095
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   8655
      Begin VB.CommandButton Cmd_Cargar 
         Caption         =   "&Cargar"
         Height          =   675
         Left            =   1680
         Picture         =   "Frm_HabDesImportar.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Carga de Datos"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5400
         Picture         =   "Frm_HabDesImportar.frx":0822
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_ImpErrores 
         Caption         =   "&Errores"
         Height          =   675
         Left            =   4200
         Picture         =   "Frm_HabDesImportar.frx":091C
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Imprimir Errores"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_ImpResumen 
         Caption         =   "&Resumen"
         Height          =   675
         Left            =   2880
         Picture         =   "Frm_HabDesImportar.frx":0FD6
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Imprimir Estadísticas"
         Top             =   240
         Width           =   790
      End
      Begin Crystal.CrystalReport Rpt_General 
         Left            =   7800
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.Frame Fra_Archivo 
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
      TabIndex        =   3
      Top             =   960
      Width           =   8655
      Begin VB.CommandButton Cmd_Buscar 
         Height          =   375
         Left            =   7680
         Picture         =   "Frm_HabDesImportar.frx":1690
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Lbl_Archivo 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         Top             =   480
         Width           =   6135
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Archivo"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         Alignment       =   2  'Center
         Caption         =   "Selección de Archivos"
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
         Index           =   2
         Left            =   2760
         TabIndex        =   5
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.Frame Fra_Periodo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   855
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      Begin VB.TextBox Txt_FecCalculo 
         Height          =   315
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   1
         ToolTipText     =   "Fecha utilizada para validar los datos de Carga"
         Top             =   360
         Width           =   1095
      End
      Begin MSComDlg.CommonDialog ComDialogo 
         Left            =   360
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Mes de Pago     :"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Frm_HabDesImportar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Dim vlArchivo As String, linea As String, vlLargoArchivo As Long
'Dim vlCont As Long, vlLargoRegistro As Long, vlAumento As Double
'
'Dim vlSwEstPeriodo As String
'Dim vlSql As String
'
''Variables para lectura de linea desde archivo
'Dim vlNumPoliza As String
'Dim vlRutBen As String
'Dim vlDgvBen As String
'Dim vlCodConHabDes As String
'Dim vlFecIniHabDes As String
'Dim vlFecTerHabDes As String
'Dim vlNumCuotas As String
'Dim vlMtoCuota As String
'Dim vlMtoTotal As String
'Dim vlCodMoneda As String
'Dim vlCodMotSusHabDes As String
'
'Dim vlNumEndoso As String
'Dim vlNumOrden As String
'Dim vlNumPerPago As String
'Dim vlNumArchivo As String
'Dim vlError As String
'Dim vlNumHabDesErr As Integer
'Dim vlNumHabDesOK As Integer
'
'Const clCodMotSusHabDes00 As String * 2 = "00"
'
'Dim vlFechaPeriodoIni As String
'Dim vlFechaPeriodoTer As String
'
'Const clCodErrCero As Integer = 0
'
''------------------------ F U N C I O N E S --------------------------
'
'Function flCargaArchivo()
'Dim ilargo As Long
'
'On Error GoTo Err_Cargaarchivo
'
'    Screen.MousePointer = 11
'
'    'abre el archivo
'    Open vlArchivo For Input As #1
'
'    vlLargoArchivo = LOF(1)
'    vlCont = 0
'    vlLargoRegistro = 75   '79  '190 '70
'    vlAumento = CDbl((100 / vlLargoArchivo) * vlLargoRegistro)
'
'    'Saca el ultimo nro de archivo
'    vlSql = ""
'    vlSql = "SELECT num_archivo FROM pp_tmae_estcarhabdes "
'    vlSql = vlSql & "ORDER BY num_archivo DESC"
'    Set vgRs = vgConexionBD.Execute(vlSql)
'    If Not vgRs.EOF Then
'        vlNumArchivo = CDbl(vgRs!num_archivo) + 1
'    Else
'        vlNumArchivo = "1"
'    End If
'
'    'Obtiene Periodo de Pago Ingresado
'    If Trim(Txt_FecCalculo.Text) <> "" Then
'        vlNumPerPago = Format(Trim(Txt_FecCalculo.Text), "yyyymmdd")
'        vlNumPerPago = Trim(Mid(vlNumPerPago, 1, 6))
'        vlFechaPeriodoIni = vlNumPerPago & "01"
'        vlFechaPeriodoTer = Format(DateSerial(CInt(Mid(vlNumPerPago, 1, 4)), CInt(Mid(vlNumPerPago, 5, 2)) + 1, 1 - 1), "yyyymmdd")
'    End If
'
'    'Consulta si existe el periodo de pago de la tabla
'    vlSql = ""
'    vlSql = "SELECT num_archivo FROM pp_tmae_estcarhabdes "
'    vlSql = vlSql & " WHERE "
'    vlSql = vlSql & " num_perpago = '" & Trim(vlNumPerPago) & "' "
'    Set vgRegistro = vgConexionBD.Execute(vlSql)
'    If Not vgRegistro.EOF Then
'       vlNumArchivo = (vgRegistro!num_archivo)
'       vlSql = ""
'       vlSql = "DELETE FROM pp_ttmp_carhabdes WHERE "
'       vlSql = vlSql & "num_archivo = " & vlNumArchivo & " "
'       vgConexionBD.Execute vlSql
'    End If
'
'    If Not EOF(1) Then
'
'        'Validar si Periodo se encuentra abierto
'        'vlSwEstPeriodo = C    Periodo Cerrado
'        'vlSwEstPeriodo = A    Periodo Abierto
'        vlSwEstPeriodo = "C"
'
'        vgSql = ""
'        vgSql = "SELECT * "
'        vgSql = vgSql & "FROM pp_tmae_propagopen "
'        vgSql = vgSql & "WHERE "
'        vgSql = vgSql & "num_perpago = '" & Trim(vlNumPerPago) & "' AND "
'        'vgSql = vgSql & "(cod_estadoreg = 'A' OR cod_estadopri = 'A') "
'        vgSql = vgSql & "(cod_estadoreg <> 'C' OR cod_estadopri <> 'C') "
'        Set vgRs = vgConexionBD.Execute(vgSql)
'        If vgRs.EOF Then
'           'Periodo Cerrado - "C"
'           MsgBox "El Periodo Ingresado se encuentra Cerrado, Debe Ingresar una Nueva Fecha de Cálculo.", vbCritical, "Error de Datos"
'           Txt_FecCalculo.SetFocus
'           Exit Function
'        End If
'
'        Frm_BarraProg.Show
'        Frm_BarraProg.Refresh
'        Frm_BarraProg.ProgressBar1.Value = 0
'        Frm_BarraProg.Lbl_Texto = "Cargando Archivo de Haberes y Descuentos Masivos " & vlArchivo
'        Frm_BarraProg.Refresh
'        Frm_BarraProg.ProgressBar1.Visible = True
'        Frm_BarraProg.Refresh
'
'    Else
'        MsgBox "El Archivo Seleccionado se encuentra Vacio.", vbCritical, "Error de Datos"
'        Exit Function
'    End If
'
'    Do While Not EOF(1)
'
'        Line Input #1, linea
'
'        linea = Replace(linea, "'", " ")
'        linea = Replace(linea, ",", ".")
'        linea = Replace(linea, "¥", "Ñ")
'        linea = Replace(linea, "#", "Ñ")
'        ilargo = Len(linea)
'        vlError = 0
'
'        'Inicializar variables
'        vlNumPoliza = ""
'        vlRutBen = ""
'        vlDgvBen = ""
'        vlCodConHabDes = ""
'        vlFecIniHabDes = ""
'        vlFecTerHabDes = ""
'        vlNumCuotas = ""
'        vlMtoCuota = ""
'        vlMtoTotal = ""
'        vlCodMoneda = ""
'        vlCodMotSusHabDes = ""
'
'        vlNumPoliza = Trim(Mid(linea, 1, 10))
'        vlRutBen = Trim(Mid(linea, 11, 9))
'        vlDgvBen = UCase(Trim(Mid(linea, 20, 1)))
'        vlCodConHabDes = Trim(Mid(linea, 21, 5))
'        vlFecIniHabDes = Trim(Mid(linea, 26, 8))
'        vlFecTerHabDes = Trim(Mid(linea, 34, 8))
'        vlCodMoneda = UCase(Trim(Mid(linea, 42, 5)))
'        vlNumCuotas = Trim(Mid(linea, 47, 5))
'        vlMtoCuota = Trim(Mid(linea, 52, 12))
'        vlMtoTotal = Trim(Mid(linea, 64, 12))
'        vlCodMotSusHabDes = clCodMotSusHabDes00
'
'        Call flValidarDatos
'        Call flGrabaDatos
'        Call flGrabarDatosDefinitivos
'
'        If Frm_BarraProg.ProgressBar1.Value + vlAumento < 100 Then
'           Frm_BarraProg.ProgressBar1.Value = Frm_BarraProg.ProgressBar1.Value + vlAumento
'           Frm_BarraProg.ProgressBar1.Refresh
'        End If
'
'    Loop
'    MsgBox "El Archivo ha Sido Cargado con Exito ", vbInformation, "Proceso De Carga Finalizado"
'
'    Close #1
'
'    Call flGrabarEstadistica
'    Call flEliminarRegSinError
'
'    Unload Frm_BarraProg
'    Screen.MousePointer = 0
'
'Exit Function
'Err_Cargaarchivo:
'    Screen.MousePointer = 0
'    Close #1
'    Unload Frm_BarraProg
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
'End Function
'
'Function flValidarDatos()
'On Error GoTo Err_flValidarDatos
'Dim vgI As Integer
'Dim vlFechaIniValida As String
'Dim vlFechaTerValida As String
'Dim vlMontoTotalValida As Double
'
'    If vlNumPoliza = "" Then
'       vlError = 304
'       'Número de póliza no enviado
'       Exit Function
'    End If
'    If vlRutBen = "" Then
'       'Rut de asegurado no enviado
'       vlError = 301
'       Exit Function
'    Else
'        If Not IsNumeric(vlRutBen) Then
'           'Rut de asegurado no es numerico
'           vlError = 300
'           Exit Function
'        End If
'    End If
'    If vlDgvBen = "" Then
'       'Dígito Verificador del Asegurado No Enviado
'       vlError = 302
'       Exit Function
'    End If
'    If Not ValiRut(vlRutBen, vlDgvBen) Then
'       'El Dígito Verificador no corresponde al número de Rut
'       vlError = 307
'       Exit Function
'    End If
'    If vlCodConHabDes = "" Then
'       'Código de Concepto H/D no Enviado
'       vlError = 50
'       Exit Function
'    End If
'    If vlFecIniHabDes = "" Then
'       'Fecha de Inicio del H/D no enviada
'       vlError = 52
'       Exit Function
'    End If
'    If vlFecTerHabDes = "" Then
'       'Fecha de Termino del H/D no enviada
'       vlError = 53
'       Exit Function
'    End If
'    If vlNumCuotas = "" Then
'       'Número de Cuotas no enviado
'       vlError = 54
'       Exit Function
'    End If
'    If vlMtoCuota = "" Then
'       'Monto de Cuota no enviado
'       vlError = 55
'       Exit Function
'    End If
'    If vlMtoTotal = "" Then
'       'Monto Total del H/D no enviado
'       vlError = 56
'       Exit Function
'    End If
'
'    'Validar Fecha de Inicio y de Término del Haber/Descuento
'
'    vlFechaIniValida = DateSerial(Mid((vlFecIniHabDes), 1, 4), Mid((vlFecIniHabDes), 5, 2), Mid((vlFecIniHabDes), 7, 2))
'    If Not IsDate(Trim(vlFechaIniValida)) Then
'        'Fecha de inicio no válida
'        vlError = 18
'        Exit Function
'    End If
'    'Validar Fecha de Termino de Haberes y Descuentos
'    vlFechaTerValida = DateSerial(Mid((vlFecTerHabDes), 1, 4), Mid((vlFecTerHabDes), 5, 2), Mid((vlFecTerHabDes), 7, 2))
'    If Not IsDate(Trim(vlFechaTerValida)) Then
'        'Fecha de Termino no válida
'        vlError = 20
'        Exit Function
'    End If
'    If (CDate(Trim(vlFechaIniValida)) > CDate(Trim(vlFechaTerValida))) Then
'        'Fecha de Inicio del H/D es mayor a la Fecha de Termino del H/D
'        vlError = 57
'        Exit Function
'    End If
'    If (Year(CDate(Trim(vlFechaIniValida))) < 1900) Then
'        'Año de Fecha de Inicio del H/D es inferior al mínimo
'        vlError = 58
'        Exit Function
'    End If
'    If (Year(CDate(Trim(vlFechaTerValida))) < 1900) Then
'        'Año de Fecha de Término del H/D es inferior al mínimo
'        vlError = 59
'        Exit Function
'    End If
'
'    'Validar Número de Cuotas y Montos del Haber/Descuento
'    vlMontoTotalValida = (CDbl(vlNumCuotas) * CDbl(vlMtoCuota))
'    If CDbl(vlMtoTotal) <> vlMontoTotalValida Then
'        'Número de cuotas y monto cuota no corresponden con Monto Total
'        vlError = 60
'        Exit Function
'    End If
'
'
''    'Validar fechas
''    If (Txt_Desde <> "") Then
''        If Not IsDate(Trim(Txt_Desde)) Then
''            Txt_Desde = ""
''            Exit Sub
''        End If
''        If Txt_Hasta <> "" Then
''            If (CDate(Trim(Txt_Desde)) > CDate(Trim(Txt_Hasta))) Then
''                Txt_Desde = ""
''                Exit Sub
''            End If
''        End If
''        If (Year(CDate(Trim(Txt_Desde))) < 1900) Then
''            Txt_Desde = ""
''            Exit Sub
''        End If
''        Txt_Desde.Text = Format(CDate(Trim(Txt_Desde)), "yyyymmdd")
''        vlFechaDesde = Trim(Txt_Desde)
''        Txt_Desde.Text = DateSerial(Mid((Txt_Desde.Text), 1, 4), Mid((Txt_Desde.Text), 5, 2), Mid((Txt_Desde.Text), 7, 2))
''    End If
'
'    'Valida que num_poliza se encuentra registrado en Tabla de Polizas
'    vgSql = ""
'    vgSql = "SELECT num_poliza "
'    vgSql = vgSql & " FROM pp_tmae_poliza "
'    vgSql = vgSql & " WHERE num_poliza = '" & Trim(vlNumPoliza) & "' "
'    Set vgRs = vgConexionBD.Execute(vgSql)
'    If vgRs.EOF Then
'        vgError = 10
'    End If
'
'    'Valida que el num_poliza y rut_ben se encuentren registrados en la Tabla
'    'de Beneficiarios
'    vgSql = ""
'    vgSql = "SELECT num_poliza "
'    vgSql = vgSql & " FROM pp_tmae_ben "
'    vgSql = vgSql & " WHERE "
'    vgSql = vgSql & " num_poliza = '" & Trim(vlNumPoliza) & "' AND "
'    vgSql = vgSql & " rut_ben = " & vlRutBen & " "
'    Set vgRs = vgConexionBD.Execute(vgSql)
'    If vgRs.EOF Then
'        vgError = 310
'    End If
'
'    'Valida que el rut del beneficiario se encuentre vigente sólo en
'    'una póliza
'    vgI = 0
''    vlNumPoliza = ""
'    vlNumOrden = ""
'    vgSql = "SELECT distinct num_poliza as numero, num_orden as orden,num_endoso as endoso "
'    vgSql = vgSql & " FROM pp_tmae_ben b WHERE "
'    vgSql = vgSql & " rut_ben = " & vlRutBen & " "
'    vgSql = vgSql & " and cod_estpension <> '10' "
'    vgSql = vgSql & " and num_endoso = "
'    vgSql = vgSql & " (select max(num_endoso) from pp_tmae_poliza where "
'    vgSql = vgSql & " num_poliza = b.num_poliza)"
'    Set vgRs = vgConexionBD.Execute(vgSql)
'    If Not (vgRs.EOF) Then
'        While Not vgRs.EOF
'            'If Not IsNull(vgRs!numero) Then vlNumeroPoliza = vgRs!numero
'            If Not IsNull(vgRs!orden) Then vlNumOrden = CStr(vgRs!orden)
'            If Not IsNull(vgRs!endoso) Then vlNumEndoso = CStr(vgRs!endoso)
'            vgI = vgI + 1
'            vgRs.MoveNext
'        Wend
'    End If
'    vgRs.Close
'
'    If vgI = 0 Then
'        'El Beneficiario se encuentra registrado, pero no tiene derecho a pension
'        vgError = 48
'        'Buscar poliza,endoso y orden para grabar registro
''        vlNumPoliza = ""
'        vlNumOrden = ""
'        vgSql = "SELECT distinct num_poliza as numero, num_orden as orden,num_endoso as endoso "
'        vgSql = vgSql & " FROM pp_tmae_ben b WHERE "
'        vgSql = vgSql & " rut_ben = " & vlRutBen & " "
'        vgSql = vgSql & " and num_endoso = "
'        vgSql = vgSql & " (select max(num_endoso) from pp_tmae_poliza where "
'        vgSql = vgSql & " num_poliza = b.num_poliza)"
'        Set vgRs = vgConexionBD.Execute(vgSql)
'        If Not (vgRs.EOF) Then
'            While Not vgRs.EOF
''                If Not IsNull(vgRs!numero) Then vlNumeroPoliza = vgRs!numero
'                If Not IsNull(vgRs!orden) Then vlNumOrden = CStr(vgRs!orden)
'                If Not IsNull(vgRs!endoso) Then vlNumEndoso = CStr(vgRs!endoso)
'                vgRs.MoveNext
'            Wend
'        End If
'        Exit Function
'    End If
'    If (vgI > 1) Then
'    'El rut del beneficiario se encuentre vigente en más de una póliza
'        vgError = 36
'        Exit Function
'    End If
'
'    'Validar que Código de Concepto de Haber/Descuento Exista
'    vgSql = ""
'    vgSql = "SELECT cod_conhabdes "
'    vgSql = vgSql & " FROM ma_tpar_conhabdes "
'    vgSql = vgSql & " WHERE "
'    vgSql = vgSql & " cod_conhabdes = " & Format(vlCodConHabDes, "00") & " "
'    Set vgRs = vgConexionBD.Execute(vgSql)
'    If vgRs.EOF Then
'        vgError = 49
'    End If
'
'    'Validar que Código de Moneda exista
'    vgSql = ""
'    vgSql = "SELECT cod_elemento "
'    vgSql = vgSql & " FROM ma_tpar_tabcod "
'    vgSql = vgSql & " WHERE "
'    vgSql = vgSql & " cod_tabla = '" & Trim(vgCodTabla_TipMon) & "' AND "
'    vgSql = vgSql & " cod_elemento = '" & Trim(vlCodMoneda) & "' "
'    Set vgRs = vgConexionBD.Execute(vgSql)
'    If vgRs.EOF Then
'        vgError = 51
'    End If
'
'Exit Function
'Err_flValidarDatos:
'    Screen.MousePointer = 0
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
'
'End Function
'
'Function flGrabaDatos()
'On Error GoTo Err_flGrabaDatos
'
'    vlSql = ""
'    vlSql = "INSERT INTO pp_ttmp_carhabdes ("
'
'    If vlNumArchivo <> "" Then vlSql = vlSql & "num_archivo, "
'    If vlNumPerPago <> "" Then vlSql = vlSql & "num_perpago, "
'    If vlNumPoliza <> "" Then vlSql = vlSql & "num_poliza, "
'    If vlNumEndoso <> "" Then vlSql = vlSql & "num_endoso, "
'    If vlNumOrden <> "" Then vlSql = vlSql & "num_orden, "
'    If vlCodConHabDes <> "" Then vlSql = vlSql & "cod_conhabdes, "
'    If vlFecIniHabDes <> "" Then vlSql = vlSql & "fec_inihabdes, "
'    If vlFecTerHabDes <> "" Then vlSql = vlSql & "fec_terhabdes, "
'    If vlNumCuotas <> "" Then vlSql = vlSql & "num_cuotas, "
'    If vlMtoCuota <> "" Then vlSql = vlSql & "mto_cuota, "
'    If vlMtoTotal <> "" Then vlSql = vlSql & "mto_total, "
'    If vlCodMoneda <> "" Then vlSql = vlSql & "cod_moneda, "
''    If vlRutBen <> "" Then vlSql = vlSql & "rut_ben, "
''    If vlDgvBen <> "" Then vlSql = vlSql & "dgv_ben, "
'
'    vlSql = vlSql & "cod_motsushabdes, "
'    vlSql = vlSql & "fec_sushabdes, "
'    vlSql = vlSql & "cod_error) VALUES ("
'    vlSql = vlSql & " " & vlNumArchivo & ","
'    vlSql = vlSql & "'" & Trim(vlNumPerPago) & "',"
'
'    If vlNumPoliza <> "" Then vlSql = vlSql & "'" & vlNumPoliza & "', "
'    If vlNumEndoso <> "" Then vlSql = vlSql & "" & vlNumEndoso & ", "
'    If vlNumOrden <> "" Then vlSql = vlSql & "" & vlNumOrden & ", "
'    If vlCodConHabDes <> "" Then vlSql = vlSql & "'" & vlCodConHabDes & "', "
'    If vlFecIniHabDes <> "" Then vlSql = vlSql & "'" & Trim(vlFecIniHabDes) & "', "
'    If vlFecTerHabDes <> "" Then vlSql = vlSql & "'" & Trim(vlFecTerHabDes) & "', "
'    If vlNumCuotas <> "" Then vlSql = vlSql & "'" & vlNumCuotas & "', "
'    If vlMtoCuota <> "" Then vlSql = vlSql & "'" & vlMtoCuota & "', "
'    If vlMtoTotal <> "" Then vlSql = vlSql & "'" & vlMtoTotal & "', "
'    If vlCodMoneda <> "" Then vlSql = vlSql & "'" & Trim(vlCodMoneda) & "', "
''    If vlRutBen <> "" Then vlSql = vlSql & "" & Str(vlRutBen) & ", "
''    If vlDgvBen <> "" Then vlSql = vlSql & "'" & Trim(vlDgvBen) & "', "
'
'    If vlCodMotSusHabDes <> "" Then vlSql = vlSql & "'" & Trim(vlCodMotSusHabDes) & "', "
'    vlSql = vlSql & " NULL , "
'    vlSql = vlSql & "" & vlError & ")"
'    vgConexionBD.Execute (vlSql)
'
'Exit Function
'Err_flGrabaDatos:
'    Screen.MousePointer = 0
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
'
'End Function
'
'Function flGrabarDatosDefinitivos()
'
'On Error GoTo Err_GrabarDatosDefinitivos
'
'    If vlError = 0 Then
'
'        vlCodMotSusHabDes = Trim(clCodMotSusHabDes00)
'
'        vlSql = ""
'        vlSql = " SELECT num_poliza "
'        vlSql = vlSql & "FROM pp_tmae_habdes "
'        vlSql = vlSql & "WHERE "
'        vlSql = vlSql & "num_poliza = '" & Trim(vlNumPoliza) & "' AND "
'        vlSql = vlSql & "num_orden = " & Str(vlNumOrden) & " AND "
'        vlSql = vlSql & "cod_conhabdes = '" & Trim(Format(vlCodConHabDes, "00")) & "' AND "
'        vlSql = vlSql & "fec_inihabdes = '" & Trim(vlFecIniHabDes) & "' "
'        Set vgRegistro = vgConexionBD.Execute(vlSql)
'        If Not vgRegistro.EOF Then
'
'            vgSql = ""
'            vgSql = "UPDATE pp_ttmp_carhabdes "
'            vgSql = vgSql & "SET cod_error = 311 "
'            vgSql = vgSql & "WHERE cod_error = " & Str(clCodErrCero) & " AND "
'            vgSql = vgSql & "num_archivo = " & Str(vlNumArchivo) & " AND "
'            vgSql = vgSql & "num_perpago = '" & Trim(vlNumPerPago) & "' AND "
'            vgSql = vgSql & "num_poliza = '" & Trim(vlNumPoliza) & "' AND "
'            vgSql = vgSql & "num_orden = " & Trim(Str(vlNumOrden)) & " "
'            vgConexionBD.Execute (vgSql)
'        Else
'
'            vlSql = ""
'            vlSql = "INSERT INTO pp_tmae_habdes ("
'            vlSql = vlSql & "num_poliza,"
'            vlSql = vlSql & "num_endoso,"
'            vlSql = vlSql & "num_orden, "
'            vlSql = vlSql & "cod_conhabdes, "
'            vlSql = vlSql & "fec_inihabdes, "
'            vlSql = vlSql & "fec_terhabdes, "
'            vlSql = vlSql & "num_cuotas, "
'            vlSql = vlSql & "mto_cuota, "
'            vlSql = vlSql & "mto_total, "
'            vlSql = vlSql & "cod_moneda, "
'            vlSql = vlSql & "cod_motsushabdes, "
'            vlSql = vlSql & "fec_sushabdes, "
'            vlSql = vlSql & "gls_obshabdes, "
'            vlSql = vlSql & "cod_usuariocrea, "
'            vlSql = vlSql & "fec_crea, "
'            vlSql = vlSql & "hor_crea, "
'            vlSql = vlSql & "num_reliq "
'            vlSql = vlSql & ") VALUES ("
'            vlSql = vlSql & " '" & Trim(vlNumPoliza) & "',"
'            vlSql = vlSql & "" & Str(vlNumEndoso) & ","
'            vlSql = vlSql & "" & Str(vlNumOrden) & ", "
'            vlSql = vlSql & "'" & Trim(Format(vlCodConHabDes, "00")) & "', "
'            vlSql = vlSql & "'" & Trim(vlFecIniHabDes) & "', "
'            vlSql = vlSql & "'" & Trim(vlFecTerHabDes) & "', "
'            vlSql = vlSql & "" & Str(vlNumCuotas) & ", "
'            vlSql = vlSql & "" & Str(vlMtoCuota) & ", "
'            vlSql = vlSql & "" & Str(vlMtoTotal) & ", "
'            vlSql = vlSql & "'" & Trim(vlCodMoneda) & "', "
'            vlSql = vlSql & "'" & Trim(vlCodMotSusHabDes) & "', "
'            vlSql = vlSql & " NULL , "
'            vlSql = vlSql & " NULL , "
'
'            vlSql = vlSql & "'" & Trim(vgUsuario) & "', "
'            vlSql = vlSql & "'" & Trim(Format(Date, "yyyymmdd")) & "', "
'            vlSql = vlSql & "'" & Trim(Format(Time, "hhmmss")) & "', "
'            vlSql = vlSql & " NULL ) "
'            vgConexionBD.Execute (vlSql)
'        End If
'    End If
'
'Exit Function
'Err_GrabarDatosDefinitivos:
'    Screen.MousePointer = 0
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
'
'
'End Function
'
'Function flGrabarEstadistica()
'On Error GoTo Err_flGrabarEstadistica
'
'    'Inicializa Variables
'    vlNumHabDesErr = 0
'    vlNumHabDesOK = 0
'
'    'Obtiene el total de registros ingresados CON Error en la tabla de temporales
'    vlSql = ""
'    vlSql = "SELECT count (num_archivo) as TotalRegistrosErr "
'    vlSql = vlSql & " FROM pp_ttmp_carhabdes "
'    vlSql = vlSql & " WHERE cod_error <> " & clCodErrCero & " AND "
'    vlSql = vlSql & " num_archivo = " & vlNumArchivo & " "
'    Set vgRegistro = vgConexionBD.Execute(vlSql)
'    If Not vgRegistro.EOF Then
'       vlNumHabDesErr = (vgRegistro!totalregistrosErr)
'    End If
'
'    'Obtiene el total de registros ingresados SIN Error en la tabla de temporales
'    vlSql = ""
'    vlSql = "SELECT count (num_archivo) as TotalRegistrosOK "
'    vlSql = vlSql & " FROM pp_ttmp_carhabdes "
'    vlSql = vlSql & " WHERE cod_error = " & clCodErrCero & " AND "
'    vlSql = vlSql & " num_archivo = " & vlNumArchivo & " "
'    Set vgRegistro = vgConexionBD.Execute(vlSql)
'    If Not vgRegistro.EOF Then
'       vlNumHabDesOK = (vgRegistro!totalregistrosOK)
'    End If
'
'    'Grabar Valores en tabla de estadísticas
'    vlSql = "SELECT num_archivo "
'    vlSql = vlSql & " FROM pp_tmae_estcarhabdes "
'    vlSql = vlSql & " WHERE num_archivo = " & vlNumArchivo & " AND "
'    vlSql = vlSql & " num_perpago = '" & Trim(vlNumPerPago) & "' "
'    Set vgRegistro = vgConexionBD.Execute(vlSql)
'    If Not vgRegistro.EOF Then
'       'Actualizar registros de estadística
'       vgSql = ""
'       vgSql = "UPDATE pp_tmae_estcarhabdes "
'       vgSql = vgSql & " SET num_habdeserr = " & vlNumHabDesErr & ", "
'       vgSql = vgSql & " num_habdesok = " & vlNumHabDesOK & ", "
'       vgSql = vgSql & " cod_usuariomodi= '" & Trim(vgUsuario) & "', "
'       vgSql = vgSql & " fec_modi = '" & Trim(Format(Date, "yyyymmdd")) & "', "
'       vgSql = vgSql & " hor_modi = '" & Trim(Format(Time, "hhmmss")) & "' "
'       vgSql = vgSql & " WHERE num_archivo = " & vlNumArchivo & " AND "
'       vgSql = vgSql & " num_perpago = '" & Trim(vlNumPerPago) & "' "
'       vgConexionBD.Execute (vgSql)
'
'    Else
'        'Ingresar registros de estadísticas
'        vlSql = ""
'        vlSql = "INSERT INTO pp_tmae_estcarhabdes ("
'        vlSql = vlSql & "num_archivo,"
'        vlSql = vlSql & "num_perpago,"
'        vlSql = vlSql & "num_habdeserr, "
'        vlSql = vlSql & "num_habdesok, "
'        vlSql = vlSql & "cod_usuariocrea, "
'        vlSql = vlSql & "fec_crea, "
'        vlSql = vlSql & "hor_crea "
'        vlSql = vlSql & ") VALUES ("
'        vlSql = vlSql & " " & Str(vlNumArchivo) & ","
'        vlSql = vlSql & "'" & (vlNumPerPago) & "',"
'        vlSql = vlSql & "" & Str(vlNumHabDesErr) & ", "
'        vlSql = vlSql & "" & Str(vlNumHabDesOK) & ", "
'        vlSql = vlSql & "'" & vgUsuario & "', "
'        vlSql = vlSql & "'" & Trim(Format(Date, "yyyymmdd")) & "', "
'        vlSql = vlSql & "'" & Trim(Format(Time, "hhmmss")) & "' ) "
'        vgConexionBD.Execute (vlSql)
'
'    End If
'
'Exit Function
'Err_flGrabarEstadistica:
'    Screen.MousePointer = 0
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
'End Function
'
'Function flEliminarRegSinError()
'
'On Error GoTo Err_flEliminarRegSinError
'
'    vlSql = ""
'    vlSql = "DELETE FROM pp_ttmp_carhabdes "
'    vlSql = vlSql & " WHERE cod_error = " & clCodErrCero & " AND "
'    vlSql = vlSql & " num_archivo = " & vlNumArchivo & " "
'    vgConexionBD.Execute vlSql
'
'Exit Function
'Err_flEliminarRegSinError:
'    Screen.MousePointer = 0
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
'
'End Function
'
'
'Private Sub Cmd_Cargar_Click()
'
'    'validacion fecha
'    If Not IsDate(Txt_FecCalculo) Then
'        MsgBox "La Fecha de Proceso ingresada no es válida.", vbCritical, "Operación Cancelada"
'        Screen.MousePointer = 0
'        Txt_FecCalculo.SetFocus
'        Exit Sub
'    End If
'    'Validar Directorios
'    If Lbl_Archivo = "" Then
'        MsgBox "Debe seleccionar Archivo de Pólizas a ser cargado.", vbCritical, "Operación Cancelada"
'        Cmd_Buscar.SetFocus
'        Exit Sub
'    End If
''    vlArchivo = Trim(Lbl_Archivo.Caption)
'    If Not fgExiste(vlArchivo) Then
'        MsgBox "Ruta Inválida o Archivo Inexistente.", vbCritical, "Operación Cancelada"
'        Exit Sub
'    End If
'
'    Call flCargaArchivo
'
'    Screen.MousePointer = 0
'End Sub
'
'Private Sub Cmd_Buscar_Click()
'Dim ilargo As Long
'On Error GoTo Err_Cmd
'
'    vlArchivo = ""
'    ComDialogo.CancelError = True
'    ComDialogo.FileName = "*.txt"
'    ComDialogo.DialogTitle = "Archivo de Haberes y Descuentos Masivos"
'    ComDialogo.Filter = "*.txt"
'    ComDialogo.FilterIndex = 1
'    ComDialogo.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
'    ComDialogo.ShowOpen
'
'    vlArchivo = ComDialogo.FileName
'    Lbl_Archivo.Caption = vlArchivo
'    If (Len(vlArchivo) > 60) Then
'        While Len(Lbl_Archivo) > 60
'            ilargo = InStr(1, Lbl_Archivo, "\")
'            Lbl_Archivo = Mid(Lbl_Archivo, ilargo + 1, Len(Lbl_Archivo))
'        Wend
'        Lbl_Archivo.Caption = "\\" & Lbl_Archivo
'    End If
'Exit Sub
'Err_Cmd:
'    If Err.Number = 32755 Then
'       Exit Sub
'    End If
'    Screen.MousePointer = 0
'    MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'
'End Sub
'
'Private Sub Cmd_ImpErrores_Click()
'
'On Error GoTo Err_CmdImpErrores
'
'    'Validación de la Fecha ingresada en Periodo.
'    If (Trim(Txt_FecCalculo.Text) = "") Then
'      MsgBox "Debe ingresar una Fecha de Periodo", vbCritical, "Error de Datos"
'      Exit Sub
'    End If
'    If Not IsDate(Txt_FecCalculo.Text) Then
'      MsgBox "La Fecha ingresada no es una fecha válida.", vbCritical, "Error de Datos"
'      Exit Sub
'    End If
'    If (Year(Txt_FecCalculo.Text) < 1900) Then
'      MsgBox "La Fecha ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
'      Exit Sub
'    End If
'
'    Txt_FecCalculo.Text = Format(CDate(Trim(Txt_FecCalculo.Text)), "yyyymmdd")
'    Txt_FecCalculo.Text = DateSerial(Mid((Txt_FecCalculo.Text), 1, 4), Mid((Txt_FecCalculo.Text), 5, 2), Mid((Txt_FecCalculo.Text), 7, 2))
'
'    Screen.MousePointer = 11
'
'    vlArchivo = strRpt & "PP_Rpt_HabDesImportarErr.rpt"   '\Reportes
'    If Not fgExiste(vlArchivo) Then     ', vbNormal
'      MsgBox "Archivo de Reporte de Listado de Errores de Carga de Haberes y Descuentos desde Archivo no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
'      Screen.MousePointer = 0
'      Exit Sub
'    End If
'
'    vlNumArchivo = ""
'    vlNumPerPago = ""
'    vlFechaPeriodoIni = ""
'    vlFechaPeriodoTer = ""
'
'    'SACA NUM PERIODO DE PAGO
'    If Trim(Txt_FecCalculo) <> "" Then
'        vlNumPerPago = Format(Trim(Txt_FecCalculo), "yyyymmdd")
'        vlNumPerPago = Mid(vlNumPerPago, 1, 6)
'        vlFechaPeriodoIni = vlNumPerPago & "01"
'        vlFechaPeriodoTer = Format(DateSerial(CInt(Mid(vlNumPerPago, 1, 4)), CInt(Mid(vlNumPerPago, 5, 2)) + 1, 1 - 1), "yyyymmdd")
'    End If
'
'    'CONSULTA SI EXISTE EL PERIODO DE PAGO EN LA TABLA
'    vlSql = ""
'    vlSql = "SELECT num_archivo FROM pp_tmae_estcarhabdes "
'    vlSql = vlSql & "WHERE "
'    vlSql = vlSql & "num_perpago = '" & Trim(vlNumPerPago) & "' "
'    Set vgRs = vgConexionBD.Execute(vlSql)
'    If Not vgRs.EOF Then
'        vlNumArchivo = vgRs!num_archivo
'    Else
'        MsgBox "No Existen Registros a Imprimir", vbCritical, "Operación Cancelada"
'        Screen.MousePointer = 0
'        Exit Sub
'     End If
'    vgRs.Close
'
'    If vlNumArchivo <> "" Then
'         vlSql = ""
'         vlSql = "SELECT num_archivo FROM pp_ttmp_carhabdes "
'         vlSql = vlSql & "WHERE "
'         vlSql = vlSql & "num_archivo = '" & vlNumArchivo & "' "
'         Set vgRs = vgConexionBD.Execute(vlSql)
'         If vgRs.EOF Then
'            MsgBox "No Existen Registros a Imprimir", vbCritical, "Operación Cancelada"
'            Screen.MousePointer = 0
'            Exit Sub
'         End If
'    End If
'
'    vgQuery = "{PP_TTMP_CARHABDES.NUM_ARCHIVO} = " & Trim(vlNumArchivo) & " "
'
'    Rpt_General.Reset
'    Rpt_General.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
'    Rpt_General.Connect = vgRutaDataBase       ' o App.Path & "\Nestle.mdb"
'    Rpt_General.SelectionFormula = vgQuery
'    Rpt_General.Formulas(0) = ""
'    Rpt_General.Formulas(1) = ""
'    Rpt_General.Formulas(2) = ""
'    Rpt_General.Formulas(3) = "NombreCompania = '" & vgNombreCompania & "'"
'    Rpt_General.Formulas(4) = "NombreSistema= '" & vgNombreSistema & "'"
'    Rpt_General.Formulas(5) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
'
'    Rpt_General.WindowState = crptMaximized
'    Rpt_General.Destination = crptToWindow
'    Rpt_General.WindowTitle = "Informe de Errores de Carga de Haberes Y Descuentos desde Archivo"
'    Rpt_General.Action = 1
'    Screen.MousePointer = 0
'
'Exit Sub
'Err_CmdImpErrores:
'    Screen.MousePointer = 0
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
'
'End Sub
'
'Private Sub Cmd_ImpResumen_Click()
'
'On Error GoTo Err_CmdImpResumen
'
'   'Validación de la Fecha ingresada en Periodo.
'   If (Trim(Txt_FecCalculo.Text) = "") Then
'      MsgBox "Debe ingresar una Fecha de Periodo", vbCritical, "Error de Datos"
'      Exit Sub
'   End If
'   If Not IsDate(Txt_FecCalculo.Text) Then
'      MsgBox "La Fecha ingresada no es una fecha válida.", vbCritical, "Error de Datos"
'      Exit Sub
'   End If
'   If (Year(Txt_FecCalculo.Text) < 1900) Then
'      MsgBox "La Fecha ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
'      Exit Sub
'   End If
'   Txt_FecCalculo.Text = Format(CDate(Trim(Txt_FecCalculo.Text)), "yyyymmdd")
'   Txt_FecCalculo.Text = DateSerial(Mid((Txt_FecCalculo.Text), 1, 4), Mid((Txt_FecCalculo.Text), 5, 2), Mid((Txt_FecCalculo.Text), 7, 2))
'
'   Screen.MousePointer = 11
'
'   vlArchivo = strRpt & "PP_Rpt_HabDesImportarEstadistica.rpt"   '\Reportes
'   If Not fgExiste(vlArchivo) Then     ', vbNormal
'      MsgBox "Archivo de Reporte Estadísticas de Carga de Haberes y Descuentos desde Archivo no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
'      Screen.MousePointer = 0
'      Exit Sub
'   End If
'
'    vlNumArchivo = ""
'    vlNumPerPago = ""
'    vlFechaPeriodoIni = ""
'    vlFechaPeriodoTer = ""
'
'    'SACA NUM PERIODO DE PAGO
'    If Trim(Txt_FecCalculo) <> "" Then
'        vlNumPerPago = Format(Trim(Txt_FecCalculo), "yyyymmdd")
'        vlNumPerPago = Mid(vlNumPerPago, 1, 6)
'        vlFechaPeriodoIni = vlNumPerPago & "01"
'        vlFechaPeriodoTer = Format(DateSerial(CInt(Mid(vlNumPerPago, 1, 4)), CInt(Mid(vlNumPerPago, 5, 2)) + 1, 1 - 1), "yyyymmdd")
'    End If
'
'    'CONSULTA SI EXISTE EL PERIODO DE PAGO EN LA TABLA
'    vlSql = ""
'    vlSql = "SELECT num_archivo FROM pp_tmae_estcarhabdes "
'    vlSql = vlSql & "WHERE "
'    vlSql = vlSql & "num_perpago = '" & vlNumPerPago & "' "
'    Set vgRs = vgConexionBD.Execute(vlSql)
'    If Not vgRs.EOF Then
'        vlNumArchivo = vgRs!num_archivo
'    Else
'        MsgBox "No Existen Registros a Imprimir", vbCritical, "Operación Cancelada"
'        Screen.MousePointer = 0
'        Exit Sub
'    End If
'    vgRs.Close
'
'   vgQuery = "{PP_TMAE_ESTCARHABDES.NUM_ARCHIVO} = " & Trim(vlNumArchivo) & " "
'
'   Rpt_General.Reset
'   Rpt_General.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
'   Rpt_General.Connect = vgRutaDataBase       ' o App.Path & "\Nestle.mdb"
'   Rpt_General.SelectionFormula = vgQuery
'   Rpt_General.Formulas(0) = ""
'   Rpt_General.Formulas(1) = ""
'   Rpt_General.Formulas(2) = ""
'
'   Rpt_General.Formulas(3) = "NombreCompania = '" & vgNombreCompania & "'"
'   Rpt_General.Formulas(4) = "NombreSistema= '" & vgNombreSistema & "'"
'   Rpt_General.Formulas(5) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
'
'   Rpt_General.WindowState = crptMaximized
'   Rpt_General.Destination = crptToWindow
'   Rpt_General.WindowTitle = "Informe de Estadísticas de Carga de Haberes y Descuentos desde Archivo"
'   Rpt_General.Action = 1
'   Screen.MousePointer = 0
'
'Exit Sub
'Err_CmdImpResumen:
'    Screen.MousePointer = 0
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
'End Sub
'
'Private Sub cmd_salir_Click()
'On Error GoTo Err_Salir
'
'    Screen.MousePointer = 11
'    Unload Me
'    Screen.MousePointer = 0
'
'Exit Sub
'Err_Salir:
'    Screen.MousePointer = 0
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
'End Sub
'
'Private Sub CmdPolizas_Click()
'Dim ilargo As Long
'On Error GoTo Err_Cmd
'
'    vlArchivo = ""
'    ComDialogo.CancelError = True
'    ComDialogo.FileName = "*.txt"
'    ComDialogo.DialogTitle = "Abrir Archivo de Haberes y Descuentos Masivos"
'    ComDialogo.Filter = "*.txt"
'    ComDialogo.FilterIndex = 1
'    ComDialogo.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
'    ComDialogo.ShowOpen
'    vlArchivo = ComDialogo.FileName
'    Lbl_Archivo.Caption = vlArchivo
'    If (Len(vlArchivo) > 60) Then
'        While Len(Lbl_Archivo) > 60
'            ilargo = InStr(1, Lbl_Archivo, "\")
'            Lbl_Archivo = Mid(Lbl_Archivo, ilargo + 1, Len(Lbl_Archivo))
'        Wend
'        Lbl_Archivo.Caption = "\\" & Lbl_Archivo
'    End If
'Exit Sub
'Err_Cmd:
'    Screen.MousePointer = 0
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
'End Sub
'
'Private Sub Form_Load()
'On Error GoTo Err_Cargar
'
'    Frm_HabDesImportar.Top = 0
'    Frm_HabDesImportar.Left = 0
'
'    vlNumArchivo = ""
'
'Exit Sub
'Err_Cargar:
'    Screen.MousePointer = 0
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
'End Sub
'
'
'Private Sub Txt_feccalculo_KeyPress(KeyAscii As Integer)
'
'If KeyAscii = 13 Then
'
'    If (Trim(Txt_FecCalculo) = "") Then
'       MsgBox "Debe Ingresar una Fecha de Cálculo", vbCritical, "Error de Datos"
'       Txt_FecCalculo.SetFocus
'       Exit Sub
'    End If
'    If Not IsDate(Txt_FecCalculo.Text) Then
'       MsgBox "La Fecha ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
'       Txt_FecCalculo.SetFocus
'       Exit Sub
'    End If
'    If (CDate(Txt_FecCalculo) > CDate(Date)) Then
'       MsgBox "La Fecha Ingresada es Mayor a la Fecha Actual", vbCritical, "Error de Datos"
'       Txt_FecCalculo.SetFocus
'       Exit Sub
'    End If
'    If (Year(Txt_FecCalculo) < 1900) Then
'       MsgBox "La Fecha Ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
'       Txt_FecCalculo.SetFocus
'       Exit Sub
'    End If
'
'    Txt_FecCalculo = Format(CDate(Trim(Txt_FecCalculo)), "yyyymmdd")
'    Txt_FecCalculo = DateSerial(Mid((Txt_FecCalculo), 1, 4), Mid((Txt_FecCalculo), 5, 2), Mid((Txt_FecCalculo), 7, 2))
'
'    Cmd_Buscar.SetFocus
'
'End If
'
'End Sub
'
'Private Sub Txt_feccalculo_LostFocus()
'
'    If (Trim(Txt_FecCalculo) = "") Then
'       Exit Sub
'    End If
'    If Not IsDate(Txt_FecCalculo.Text) Then
'       Exit Sub
'    End If
'    If (CDate(Txt_FecCalculo) > CDate(Date)) Then
'       Exit Sub
'    End If
'    If (Year(Txt_FecCalculo) < 1900) Then
'       Exit Sub
'    End If
'
'    Txt_FecCalculo = Format(CDate(Trim(Txt_FecCalculo)), "yyyymmdd")
'    Txt_FecCalculo = DateSerial(Mid((Txt_FecCalculo), 1, 4), Mid((Txt_FecCalculo), 5, 2), Mid((Txt_FecCalculo), 7, 2))
'
'End Sub

Option Explicit

Dim vlArchivo As String, linea As String, vlLargoArchivo As Long
Dim vlcont As Long, vlLargoRegistro As Long, vlAumento As Double

Dim vlSwEstPeriodo As String
Dim vlSql As String

'Variables para lectura de linea desde archivo
Dim vlNumPoliza As String
''Dim vlRutBen As String
''Dim vlDgvBen As String
Dim vlTipoIden  As String
Dim vlNumIden   As String
Dim vlCodConHabDes As String
Dim vlFecIniHabDes As String
Dim vlFecTerHabDes As String
Dim vlNumCuotas As String
Dim vlMtoCuota As String
Dim vlMtoTotal As String
Dim vlCodMoneda As String
Dim vlCodMotSusHabDes As String

Dim vlNumEndoso As String
Dim vlNumOrden As String
Dim vlNumPerPago As String
Dim vlNumArchivo As String
Dim vlError As String
Dim vlNumHabDesErr As Integer
Dim vlNumHabDesOK As Integer

Const clCodMotSusHabDes00 As String * 2 = "00"

Dim vlFechaPeriodoIni As String
Dim vlFechaPeriodoTer As String

Const clCodErrCero As Integer = 0

'------------------------ F U N C I O N E S --------------------------
        
Function flCargaArchivo()
Dim ilargo As Long

On Error GoTo Err_Cargaarchivo

    Screen.MousePointer = 11
    
    'abre el archivo
    Open vlArchivo For Input As #1
    
    vlLargoArchivo = LOF(1)
    vlcont = 0
    vlLargoRegistro = 75   '79  '190 '70
    vlAumento = CDbl((100 / vlLargoArchivo) * vlLargoRegistro)
    
    'Saca el ultimo nro de archivo
    vlSql = ""
    vlSql = "SELECT num_archivo FROM pp_tmae_estcarhabdes "
    vlSql = vlSql & "ORDER BY num_archivo DESC"
    Set vgRs = vgConexionBD.Execute(vlSql)
    If Not vgRs.EOF Then
        vlNumArchivo = CDbl(vgRs!num_archivo) + 1
    Else
        vlNumArchivo = "1"
    End If
    
    'Obtiene Periodo de Pago Ingresado
    If Trim(Txt_FecCalculo.Text) <> "" Then
        vlNumPerPago = Format(Trim(Txt_FecCalculo.Text), "yyyymmdd")
        vlNumPerPago = Trim(Mid(vlNumPerPago, 1, 6))
        vlFechaPeriodoIni = vlNumPerPago & "01"
        vlFechaPeriodoTer = Format(DateSerial(CInt(Mid(vlNumPerPago, 1, 4)), CInt(Mid(vlNumPerPago, 5, 2)) + 1, 1 - 1), "yyyymmdd")
    End If
    
    'Consulta si existe el periodo de pago de la tabla
    vlSql = ""
    vlSql = "SELECT num_archivo FROM pp_tmae_estcarhabdes "
    vlSql = vlSql & " WHERE "
    vlSql = vlSql & " num_perpago = '" & Trim(vlNumPerPago) & "' "
    Set vgRegistro = vgConexionBD.Execute(vlSql)
    If Not vgRegistro.EOF Then
       vlNumArchivo = (vgRegistro!num_archivo)
       vlSql = ""
       vlSql = "DELETE FROM pp_ttmp_carhabdes WHERE "
       vlSql = vlSql & "num_archivo = " & vlNumArchivo & " "
       vgConexionBD.Execute vlSql
    End If
    
    If Not EOF(1) Then
        
        'Validar si Periodo se encuentra abierto
        'vlSwEstPeriodo = C    Periodo Cerrado
        'vlSwEstPeriodo = A    Periodo Abierto
        vlSwEstPeriodo = "C"
        
        vgSql = ""
        vgSql = "SELECT * "
        vgSql = vgSql & "FROM pp_tmae_propagopen "
        vgSql = vgSql & "WHERE "
        vgSql = vgSql & "num_perpago = '" & Trim(vlNumPerPago) & "' AND "
        ''vgSql = vgSql & "(cod_estadoreg = 'A' OR cod_estadopri = 'A') "
        'vgSql = vgSql & "(cod_estadoreg <> 'C' OR cod_estadopri <> 'C') "ABV 15-09-2007
        vgSql = vgSql & "(cod_estadoreg <> 'C')" 'OR cod_estadopri <> 'C') "
        Set vgRs = vgConexionBD.Execute(vgSql)
        If vgRs.EOF Then
           'Periodo Cerrado - "C"
           MsgBox "El Periodo Ingresado se encuentra Cerrado, Debe Ingresar una Nueva Fecha de Cálculo.", vbCritical, "Error de Datos"
           Txt_FecCalculo.SetFocus
           Exit Function
        End If
        
        Frm_BarraProg.Show
        Frm_BarraProg.Refresh
        Frm_BarraProg.ProgressBar1.Value = 0
        Frm_BarraProg.Lbl_Texto = "Cargando Archivo de Haberes y Descuentos Masivos " & vlArchivo
        Frm_BarraProg.Refresh
        Frm_BarraProg.ProgressBar1.Visible = True
        Frm_BarraProg.Refresh
            
    Else
        MsgBox "El Archivo Seleccionado se encuentra Vacio.", vbCritical, "Error de Datos"
        Exit Function
    End If
    
    Do While Not EOF(1)
        
        Line Input #1, linea
              
        linea = Replace(linea, "'", " ")
        linea = Replace(linea, ",", ".")
        linea = Replace(linea, "¥", "Ñ")
        linea = Replace(linea, "#", "Ñ")
        ilargo = Len(linea)
        vlError = 0
        
'        vgError = 0
                        
        'Inicializar variables
        vlNumPoliza = ""
        vlTipoIden = ""
        vlNumIden = ""
        vlCodConHabDes = ""
        vlFecIniHabDes = ""
        vlFecTerHabDes = ""
        vlNumCuotas = ""
        vlMtoCuota = ""
        vlMtoTotal = ""
        vlCodMoneda = ""
        vlCodMotSusHabDes = ""
        
        vlNumPoliza = Trim(Mid(linea, 1, 10))
        vlTipoIden = Trim(Mid(linea, 11, 3))
        vlNumIden = UCase(Trim(Mid(linea, 14, 16)))
        vlCodConHabDes = Trim(Mid(linea, 30, 5))
        vlFecIniHabDes = Trim(Mid(linea, 35, 8))
        vlFecTerHabDes = Trim(Mid(linea, 43, 8))
        vlCodMoneda = UCase(Trim(Mid(linea, 51, 5)))
        vlNumCuotas = Trim(Mid(linea, 56, 5))
        vlMtoCuota = Trim(Mid(linea, 61, 12))
        vlMtoTotal = Trim(Mid(linea, 73, 12))
        vlCodMotSusHabDes = clCodMotSusHabDes00
        
        Call flValidarDatos
        Call flGrabaDatos
        Call flGrabarDatosDefinitivos
                
        If Frm_BarraProg.ProgressBar1.Value + vlAumento < 100 Then
           Frm_BarraProg.ProgressBar1.Value = Frm_BarraProg.ProgressBar1.Value + vlAumento
           Frm_BarraProg.ProgressBar1.Refresh
        End If
            
    Loop
    MsgBox "El Archivo ha Sido Cargado con Exito ", vbInformation, "Proceso De Carga Finalizado"
    
    Close #1

    Call flGrabarEstadistica
    Call flEliminarRegSinError
         
    Unload Frm_BarraProg
    Screen.MousePointer = 0
    
Exit Function
Err_Cargaarchivo:
    Screen.MousePointer = 0
    Close #1
    Unload Frm_BarraProg
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flValidarDatos()
On Error GoTo Err_flValidarDatos
Dim vgI As Integer
Dim vlFechaIniValida As String
Dim vlFechaTerValida As String
Dim vlMontoTotalValida As Double

    If vlNumPoliza = "" Then
       vlError = 304
       'Número de póliza no enviado
       Exit Function
    End If
    If vlTipoIden = "" Then
       'Tipo Identificación del asegurado no enviado
       vlError = 301
       Exit Function
    Else
        If Not IsNumeric(vlTipoIden) Then
           'Tipo Identificación del asegurado no es numerico
           vlError = 300
           Exit Function
        End If
    End If
    If vlNumIden = "" Then
       'Número de Identificación del Asegurado No Enviado
       vlError = 302
       Exit Function
    End If
''    If Not ValiRut(vlRutBen, vlDgvBen) Then
''       'El Dígito Verificador no corresponde al número de Rut
''       vlError = 307
''       Exit Function
''    End If
    If vlCodConHabDes = "" Then
       'Código de Concepto H/D no Enviado
       vlError = 50
       Exit Function
    End If
    If vlFecIniHabDes = "" Then
       'Fecha de Inicio del H/D no enviada
       vlError = 52
       Exit Function
    End If
    If vlFecTerHabDes = "" Then
       'Fecha de Termino del H/D no enviada
       vlError = 53
       Exit Function
    End If
    If vlNumCuotas = "" Then
       'Número de Cuotas no enviado
       vlError = 54
       Exit Function
    Else
        If Not IsNumeric(vlNumCuotas) Then
           'Número de Cuotas no es numerico
           vlError = 54
           Exit Function
        End If
    End If
    If vlMtoCuota = "" Then
       'Monto de Cuota no enviado
       vlError = 55
       Exit Function
        If Not IsNumeric(vlMtoCuota) Then
           'Monto de Cuota no es numerico
           vlError = 55
           Exit Function
        End If
    End If
    If vlMtoTotal = "" Then
       'Monto Total del H/D no enviado
       vlError = 56
       Exit Function
       If Not IsNumeric(vlMtoTotal) Then
          'Monto Total del H/D no es numerico
          vlError = 56
          Exit Function
       End If
    End If
    
    'Validar Fecha de Inicio y de Término del Haber/Descuento
    
    vlFechaIniValida = DateSerial(Mid((vlFecIniHabDes), 1, 4), Mid((vlFecIniHabDes), 5, 2), Mid((vlFecIniHabDes), 7, 2))
    If Not IsDate(Trim(vlFechaIniValida)) Then
        'Fecha de inicio no válida
        vlError = 18
        Exit Function
    End If
    'Validar Fecha de Termino de Haberes y Descuentos
    vlFechaTerValida = DateSerial(Mid((vlFecTerHabDes), 1, 4), Mid((vlFecTerHabDes), 5, 2), Mid((vlFecTerHabDes), 7, 2))
    If Not IsDate(Trim(vlFechaTerValida)) Then
        'Fecha de Termino no válida
        vlError = 20
        Exit Function
    End If
    If (CDate(Trim(vlFechaIniValida)) > CDate(Trim(vlFechaTerValida))) Then
        'Fecha de Inicio del H/D es mayor a la Fehca de Termino del H/D
        vlError = 57
        Exit Function
    End If
    If (Year(CDate(Trim(vlFechaIniValida))) < 1900) Then
        'Año de Fecha de Inicio del H/D es inferior al mínimo
        vlError = 58
        Exit Function
    End If
    If (Year(CDate(Trim(vlFechaTerValida))) < 1900) Then
        'Año de Fecha de Término del H/D es inferior al mínimo
        vlError = 59
        Exit Function
    End If
    
    'Validar Número de Cuotas y Montos del Haber/Descuento
    vlMontoTotalValida = (CDbl(vlNumCuotas) * CDbl(vlMtoCuota))
    If CDbl(vlMtoTotal) <> vlMontoTotalValida Then
        'Número de cuotas y monto cuota no corresponden con Monto Total
        vlError = 60
        Exit Function
    End If
    
    
'    'Validar fechas
'    If (Txt_Desde <> "") Then
'        If Not IsDate(Trim(Txt_Desde)) Then
'            Txt_Desde = ""
'            Exit Sub
'        End If
'        If Txt_Hasta <> "" Then
'            If (CDate(Trim(Txt_Desde)) > CDate(Trim(Txt_Hasta))) Then
'                Txt_Desde = ""
'                Exit Sub
'            End If
'        End If
'        If (Year(CDate(Trim(Txt_Desde))) < 1900) Then
'            Txt_Desde = ""
'            Exit Sub
'        End If
'        Txt_Desde.Text = Format(CDate(Trim(Txt_Desde)), "yyyymmdd")
'        vlFechaDesde = Trim(Txt_Desde)
'        Txt_Desde.Text = DateSerial(Mid((Txt_Desde.Text), 1, 4), Mid((Txt_Desde.Text), 5, 2), Mid((Txt_Desde.Text), 7, 2))
'    End If
    
    'Valida que num_poliza se encuentra registrado en Tabla de Polizas
    vgSql = ""
    vgSql = "SELECT num_poliza "
    vgSql = vgSql & " FROM pp_tmae_poliza "
    vgSql = vgSql & " WHERE num_poliza = '" & Trim(vlNumPoliza) & "' "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If vgRs.EOF Then
        vlError = 10
        Exit Function
    End If
    
    'Valida que el num_poliza y la identificación se encuentren registrados en la Tabla
    'de Beneficiarios
    vgSql = ""
    vgSql = "SELECT num_poliza "
    vgSql = vgSql & " FROM pp_tmae_ben "
    vgSql = vgSql & " WHERE "
    vgSql = vgSql & " num_poliza = '" & Trim(vlNumPoliza) & "' AND "
    vgSql = vgSql & " cod_tipoidenben = " & vlTipoIden & " AND "
    vgSql = vgSql & " num_idenben = '" & vlNumIden & "' "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If vgRs.EOF Then
        vlError = 310
        Exit Function
    End If
    
    'Valida que la identificación del beneficiario se encuentre vigente sólo en
    'una póliza
    vgI = 0
'    vlNumPoliza = ""
    vlNumOrden = ""
        
    vgSql = "SELECT distinct num_poliza as numero, num_orden as orden,num_endoso as endoso "
    vgSql = vgSql & " FROM pp_tmae_ben b WHERE "
    vgSql = vgSql & " cod_tipoidenben = " & vlTipoIden & " "
    vgSql = vgSql & " and num_idenben='" & vlNumIden & "' "
    vgSql = vgSql & " and cod_estpension <> '10' "
    vgSql = vgSql & " and num_endoso = "
    vgSql = vgSql & " (select max(num_endoso) from pp_tmae_poliza where "
    vgSql = vgSql & " num_poliza = b.num_poliza)"
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not (vgRs.EOF) Then
        While Not vgRs.EOF
            'If Not IsNull(vgRs!numero) Then vlNumeroPoliza = vgRs!numero
            If Not IsNull(vgRs!Orden) Then vlNumOrden = (vgRs!Orden)
            If Not IsNull(vgRs!Endoso) Then vlNumEndoso = (vgRs!Endoso)
            vgI = vgI + 1
            vgRs.MoveNext
        Wend
    End If
    vgRs.Close
    
    If vgI = 0 Then
        'El Beneficiario se encuentra registrado, pero no tiene derecho a pension
        vlError = 48
        'Buscar poliza,endoso y orden para grabar registro
'        vlNumPoliza = ""
        vlNumOrden = ""
        vgSql = "SELECT distinct num_poliza as numero, num_orden as orden,num_endoso as endoso "
        vgSql = vgSql & " FROM pp_tmae_ben b WHERE "
        vgSql = vgSql & " cod_tipoidenben = " & vlTipoIden & " "
        vgSql = vgSql & " and num_idenben='" & vlNumIden & "' "
        vgSql = vgSql & " and num_endoso = "
        vgSql = vgSql & " (select max(num_endoso) from pp_tmae_poliza where "
        vgSql = vgSql & " num_poliza = b.num_poliza)"
        Set vgRs = vgConexionBD.Execute(vgSql)
        If Not (vgRs.EOF) Then
            While Not vgRs.EOF
'                If Not IsNull(vgRs!numero) Then vlNumeroPoliza = vgRs!numero
                If Not IsNull(vgRs!Orden) Then vlNumOrden = vgRs!Orden
                If Not IsNull(vgRs!Endoso) Then vlNumEndoso = vgRs!Endoso
                vgRs.MoveNext
            Wend
        End If
        Exit Function
    End If
    If (vgI > 1) Then
    'La Identificación del beneficiario se encuentre vigente en más de una póliza
        vlError = 36
        Exit Function
    End If
    
    'Validar que Código de Concepto de Haber/Descuento Exista
    vgSql = ""
    vgSql = "SELECT cod_conhabdes "
    vgSql = vgSql & "FROM ma_tpar_conhabdes "
    vgSql = vgSql & "WHERE "
    'vgSql = vgSql & "cod_conhabdes = '" & Format(vlCodConHabDes, "00") & "' "
    vgSql = vgSql & "cod_conhabdes = '" & Trim(Format(vlCodConHabDes, "00")) & "' "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If vgRs.EOF Then
        vlError = 49
    End If

    'Validar que Código de Moneda exista
    vgSql = ""
    vgSql = "SELECT cod_elemento "
    vgSql = vgSql & " FROM ma_tpar_tabcod "
    vgSql = vgSql & " WHERE "
    vgSql = vgSql & " cod_tabla = '" & Trim(vgCodTabla_TipMon) & "' AND "
    vgSql = vgSql & " cod_elemento = '" & Trim(vlCodMoneda) & "' "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If vgRs.EOF Then
        vlError = 51
    End If
    
    'vgError = vlError

Exit Function
Err_flValidarDatos:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function flGrabaDatos()
On Error GoTo Err_flGrabaDatos

    vlSql = ""
    vlSql = "INSERT INTO pp_ttmp_carhabdes ("
    
    If vlNumArchivo <> "" Then vlSql = vlSql & "num_archivo, "
    If vlNumPerPago <> "" Then vlSql = vlSql & "num_perpago, "
    If vlNumPoliza <> "" Then vlSql = vlSql & "num_poliza, "
    If vlNumEndoso <> "" Then vlSql = vlSql & "num_endoso, "
    If vlNumOrden <> "" Then vlSql = vlSql & "num_orden, "
    If vlCodConHabDes <> "" Then vlSql = vlSql & "cod_conhabdes, "
    If vlFecIniHabDes <> "" Then vlSql = vlSql & "fec_inihabdes, "
    If vlFecTerHabDes <> "" Then vlSql = vlSql & "fec_terhabdes, "
    If vlNumCuotas <> "" Then vlSql = vlSql & "num_cuotas, "
    If vlMtoCuota <> "" Then vlSql = vlSql & "mto_cuota, "
    If vlMtoTotal <> "" Then vlSql = vlSql & "mto_total, "
    If vlCodMoneda <> "" Then vlSql = vlSql & "cod_moneda, "
    If vlTipoIden <> "" Then vlSql = vlSql & "cod_tipoidenben, "
    If vlNumIden <> "" Then vlSql = vlSql & "num_idenben, "
    
    vlSql = vlSql & "cod_motsushabdes, "
    vlSql = vlSql & "fec_sushabdes, "
    vlSql = vlSql & "cod_error) VALUES ("
    vlSql = vlSql & " " & vlNumArchivo & ","
    vlSql = vlSql & "'" & Trim(vlNumPerPago) & "',"
    
    If vlNumPoliza <> "" Then vlSql = vlSql & "'" & vlNumPoliza & "', "
    If vlNumEndoso <> "" Then vlSql = vlSql & "" & vlNumEndoso & ", "
    If vlNumOrden <> "" Then vlSql = vlSql & "" & vlNumOrden & ", "
    If vlCodConHabDes <> "" Then vlSql = vlSql & "" & Str(vlCodConHabDes) & ", "
    If vlFecIniHabDes <> "" Then vlSql = vlSql & "'" & Trim(vlFecIniHabDes) & "', "
    If vlFecTerHabDes <> "" Then vlSql = vlSql & "'" & Trim(vlFecTerHabDes) & "', "
    If vlNumCuotas <> "" Then vlSql = vlSql & "" & Str(vlNumCuotas) & ", "
    If vlMtoCuota <> "" Then vlSql = vlSql & "" & Str(vlMtoCuota) & ", "
    If vlMtoTotal <> "" Then vlSql = vlSql & "" & Str(vlMtoTotal) & ", "
    If vlCodMoneda <> "" Then vlSql = vlSql & "'" & Trim(vlCodMoneda) & "', "
    If vlTipoIden <> "" Then vlSql = vlSql & "'" & Trim(vlTipoIden) & "', "
    If vlNumIden <> "" Then vlSql = vlSql & "'" & Trim(vlNumIden) & "', "
    
    If vlCodMotSusHabDes <> "" Then vlSql = vlSql & "'" & Trim(vlCodMotSusHabDes) & "', "
    vlSql = vlSql & " NULL , "
'    vlSql = vlSql & "" & vgError & ")"
    vlSql = vlSql & "" & vlError & ")"
    vgConexionBD.Execute (vlSql)
    
Exit Function
Err_flGrabaDatos:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
    
End Function

Function flGrabarDatosDefinitivos()

On Error GoTo Err_GrabarDatosDefinitivos

'    If vgError = 0 Then
    If vlError = 0 Then
    
        vlCodMotSusHabDes = Trim(clCodMotSusHabDes00)
    
        vlSql = ""
        vlSql = " SELECT num_poliza "
        vlSql = vlSql & "FROM pp_tmae_habdes "
        vlSql = vlSql & "WHERE "
        vlSql = vlSql & "num_poliza = '" & Trim(vlNumPoliza) & "' AND "
        vlSql = vlSql & "num_orden = " & Str(vlNumOrden) & " AND "
        vlSql = vlSql & "cod_conhabdes = '" & Trim(Format(vlCodConHabDes, "00")) & "' AND "
        vlSql = vlSql & "fec_inihabdes = '" & Trim(vlFecIniHabDes) & "' "
        Set vgRegistro = vgConexionBD.Execute(vlSql)
        If Not vgRegistro.EOF Then
                
            vgSql = ""
            vgSql = "UPDATE pp_ttmp_carhabdes "
            vgSql = vgSql & "SET cod_error = 311 "
            vgSql = vgSql & "WHERE cod_error = " & Str(clCodErrCero) & " AND "
            vgSql = vgSql & "num_archivo = " & Str(vlNumArchivo) & " AND "
            vgSql = vgSql & "num_perpago = '" & Trim(vlNumPerPago) & "' AND "
            vgSql = vgSql & "num_poliza = '" & Trim(vlNumPoliza) & "' AND "
            vgSql = vgSql & "num_orden = " & Trim(Str(vlNumOrden)) & " "
            vgConexionBD.Execute (vgSql)
        Else

            vlSql = ""
            vlSql = "INSERT INTO pp_tmae_habdes ("
            vlSql = vlSql & "num_poliza,"
            vlSql = vlSql & "num_endoso,"
            vlSql = vlSql & "num_orden, "
            vlSql = vlSql & "cod_conhabdes, "
            vlSql = vlSql & "fec_inihabdes, "
            vlSql = vlSql & "fec_terhabdes, "
            vlSql = vlSql & "num_cuotas, "
            vlSql = vlSql & "mto_cuota, "
            vlSql = vlSql & "mto_total, "
            vlSql = vlSql & "cod_moneda, "
            vlSql = vlSql & "cod_motsushabdes, "
            vlSql = vlSql & "fec_sushabdes, "
            vlSql = vlSql & "gls_obshabdes, "
            vlSql = vlSql & "cod_usuariocrea, "
            vlSql = vlSql & "fec_crea, "
            vlSql = vlSql & "hor_crea, "
            vlSql = vlSql & "num_reliq "
            vlSql = vlSql & ") VALUES ("
            vlSql = vlSql & " '" & Trim(vlNumPoliza) & "',"
            vlSql = vlSql & "" & Str(vlNumEndoso) & ","
            vlSql = vlSql & "" & Str(vlNumOrden) & ", "
            vlSql = vlSql & "'" & Trim(Format(vlCodConHabDes, "00")) & "', "
            vlSql = vlSql & "'" & Trim(vlFecIniHabDes) & "', "
            vlSql = vlSql & "'" & Trim(vlFecTerHabDes) & "', "
            vlSql = vlSql & "" & Str(vlNumCuotas) & ", "
            vlSql = vlSql & "" & Str(vlMtoCuota) & ", "
            vlSql = vlSql & "" & Str(vlMtoTotal) & ", "
            vlSql = vlSql & "'" & Trim(vlCodMoneda) & "', "
            vlSql = vlSql & "'" & Trim(vlCodMotSusHabDes) & "', "
            vlSql = vlSql & " NULL , "
            vlSql = vlSql & " NULL , "
            
            vlSql = vlSql & "'" & Trim(vgUsuario) & "', "
            vlSql = vlSql & "'" & Trim(Format(Date, "yyyymmdd")) & "', "
            vlSql = vlSql & "'" & Trim(Format(Time, "hhmmss")) & "', "
            vlSql = vlSql & " NULL ) "
            vgConexionBD.Execute (vlSql)
        End If
    End If

Exit Function
Err_GrabarDatosDefinitivos:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flGrabarEstadistica()
On Error GoTo Err_flGrabarEstadistica

    'Inicializa Variables
    vlNumHabDesErr = 0
    vlNumHabDesOK = 0
    
    'Obtiene el total de registros ingresados CON Error en la tabla de temporales
    vlSql = ""
    vlSql = "SELECT count (num_archivo) as TotalRegistrosErr "
    vlSql = vlSql & " FROM pp_ttmp_carhabdes "
    vlSql = vlSql & " WHERE cod_error <> " & clCodErrCero & " AND "
    vlSql = vlSql & " num_archivo = " & vlNumArchivo & " "
    Set vgRegistro = vgConexionBD.Execute(vlSql)
    If Not vgRegistro.EOF Then
       vlNumHabDesErr = (vgRegistro!totalregistrosErr)
    End If

    'Obtiene el total de registros ingresados SIN Error en la tabla de temporales
    vlSql = ""
    vlSql = "SELECT count (num_archivo) as TotalRegistrosOK "
    vlSql = vlSql & " FROM pp_ttmp_carhabdes "
    vlSql = vlSql & " WHERE cod_error = " & clCodErrCero & " AND "
    vlSql = vlSql & " num_archivo = " & vlNumArchivo & " "
    Set vgRegistro = vgConexionBD.Execute(vlSql)
    If Not vgRegistro.EOF Then
       vlNumHabDesOK = (vgRegistro!totalregistrosOK)
    End If
    
    'Grabar Valores en tabla de estadísticas
    vlSql = "SELECT num_archivo "
    vlSql = vlSql & " FROM pp_tmae_estcarhabdes "
    vlSql = vlSql & " WHERE num_archivo = " & vlNumArchivo & " AND "
    vlSql = vlSql & " num_perpago = '" & Trim(vlNumPerPago) & "' "
    Set vgRegistro = vgConexionBD.Execute(vlSql)
    If Not vgRegistro.EOF Then
       'Actualizar registros de estadística
       vgSql = ""
       vgSql = "UPDATE pp_tmae_estcarhabdes "
       vgSql = vgSql & " SET num_habdeserr = " & vlNumHabDesErr & ", "
       vgSql = vgSql & " num_habdesok = " & vlNumHabDesOK & ", "
       vgSql = vgSql & " cod_usuariomodi= '" & Trim(vgUsuario) & "', "
       vgSql = vgSql & " fec_modi = '" & Trim(Format(Date, "yyyymmdd")) & "', "
       vgSql = vgSql & " hor_modi = '" & Trim(Format(Time, "hhmmss")) & "' "
       vgSql = vgSql & " WHERE num_archivo = " & vlNumArchivo & " AND "
       vgSql = vgSql & " num_perpago = '" & Trim(vlNumPerPago) & "' "
       vgConexionBD.Execute (vgSql)
       
    Else
        'Ingresar registros de estadísticas
        vlSql = ""
        vlSql = "INSERT INTO pp_tmae_estcarhabdes ("
        vlSql = vlSql & "num_archivo,"
        vlSql = vlSql & "num_perpago,"
        vlSql = vlSql & "num_habdeserr, "
        vlSql = vlSql & "num_habdesok, "
        vlSql = vlSql & "cod_usuariocrea, "
        vlSql = vlSql & "fec_crea, "
        vlSql = vlSql & "hor_crea "
        vlSql = vlSql & ") VALUES ("
        vlSql = vlSql & " " & Str(vlNumArchivo) & ","
        vlSql = vlSql & "'" & (vlNumPerPago) & "',"
        vlSql = vlSql & "" & Str(vlNumHabDesErr) & ", "
        vlSql = vlSql & "" & Str(vlNumHabDesOK) & ", "
        vlSql = vlSql & "'" & vgUsuario & "', "
        vlSql = vlSql & "'" & Trim(Format(Date, "yyyymmdd")) & "', "
        vlSql = vlSql & "'" & Trim(Format(Time, "hhmmss")) & "' ) "
        vgConexionBD.Execute (vlSql)
        
    End If

Exit Function
Err_flGrabarEstadistica:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flEliminarRegSinError()

On Error GoTo Err_flEliminarRegSinError

    vlSql = ""
    vlSql = "DELETE FROM pp_ttmp_carhabdes "
    vlSql = vlSql & " WHERE cod_error = " & clCodErrCero & " AND "
    vlSql = vlSql & " num_archivo = " & vlNumArchivo & " "
    vgConexionBD.Execute vlSql
    
Exit Function
Err_flEliminarRegSinError:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function


Private Sub Cmd_Cargar_Click()

    'validacion fecha
    If Not IsDate(Txt_FecCalculo) Then
        MsgBox "La Fecha de Proceso ingresada no es válida.", vbCritical, "Operación Cancelada"
        Screen.MousePointer = 0
        Txt_FecCalculo.SetFocus
        Exit Sub
    End If
    'Validar Directorios
    If Lbl_Archivo = "" Then
        MsgBox "Debe seleccionar Archivo de Pólizas a ser cargado.", vbCritical, "Operación Cancelada"
        Cmd_Buscar.SetFocus
        Exit Sub
    End If
'    vlArchivo = Trim(Lbl_Archivo.Caption)
    If Not fgExiste(vlArchivo) Then
        MsgBox "Ruta Inválida o Archivo Inexistente.", vbCritical, "Operación Cancelada"
        Exit Sub
    End If
    
    Call flCargaArchivo
    
    Screen.MousePointer = 0
End Sub

Private Sub Cmd_Buscar_Click()
Dim ilargo As Long
On Error GoTo Err_Cmd

    vlArchivo = ""
    ComDialogo.CancelError = True
    ComDialogo.FileName = "*.txt"
    ComDialogo.DialogTitle = "Archivo de Haberes y Descuentos Masivos"
    ComDialogo.Filter = "*.txt"
    ComDialogo.FilterIndex = 1
    ComDialogo.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    ComDialogo.ShowOpen
    
    vlArchivo = ComDialogo.FileName
    Lbl_Archivo.Caption = vlArchivo
    If (Len(vlArchivo) > 60) Then
        While Len(Lbl_Archivo) > 60
            ilargo = InStr(1, Lbl_Archivo, "\")
            Lbl_Archivo = Mid(Lbl_Archivo, ilargo + 1, Len(Lbl_Archivo))
        Wend
        Lbl_Archivo.Caption = "\\" & Lbl_Archivo
    End If
Exit Sub
Err_Cmd:
    If Err.Number = 32755 Then
       Exit Sub
    End If
    Screen.MousePointer = 0
    MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    
End Sub

Private Sub Cmd_ImpErrores_Click()

On Error GoTo Err_CmdImpErrores

    'Validación de la Fecha ingresada en Periodo.
    If (Trim(Txt_FecCalculo.Text) = "") Then
      MsgBox "Debe ingresar una Fecha de Periodo", vbCritical, "Error de Datos"
      Exit Sub
    End If
    If Not IsDate(Txt_FecCalculo.Text) Then
      MsgBox "La Fecha ingresada no es una fecha válida.", vbCritical, "Error de Datos"
      Exit Sub
    End If
    If (Year(Txt_FecCalculo.Text) < 1900) Then
      MsgBox "La Fecha ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
      Exit Sub
    End If
    
    Txt_FecCalculo.Text = Format(CDate(Trim(Txt_FecCalculo.Text)), "yyyymmdd")
    Txt_FecCalculo.Text = DateSerial(Mid((Txt_FecCalculo.Text), 1, 4), Mid((Txt_FecCalculo.Text), 5, 2), Mid((Txt_FecCalculo.Text), 7, 2))
        
    Screen.MousePointer = 11
    
    vlArchivo = strRpt & "PP_Rpt_HabDesImportarErr.rpt"   '\Reportes
    If Not fgExiste(vlArchivo) Then     ', vbNormal
      MsgBox "Archivo de Reporte de Listado de Errores de Carga de Haberes y Descuentos desde Archivo no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
      Screen.MousePointer = 0
      Exit Sub
    End If
    
    vlNumArchivo = ""
    vlNumPerPago = ""
    vlFechaPeriodoIni = ""
    vlFechaPeriodoTer = ""
    
    'SACA NUM PERIODO DE PAGO
    If Trim(Txt_FecCalculo) <> "" Then
        vlNumPerPago = Format(Trim(Txt_FecCalculo), "yyyymmdd")
        vlNumPerPago = Mid(vlNumPerPago, 1, 6)
        vlFechaPeriodoIni = vlNumPerPago & "01"
        vlFechaPeriodoTer = Format(DateSerial(CInt(Mid(vlNumPerPago, 1, 4)), CInt(Mid(vlNumPerPago, 5, 2)) + 1, 1 - 1), "yyyymmdd")
    End If
    
    'CONSULTA SI EXISTE EL PERIODO DE PAGO EN LA TABLA
    vlSql = ""
    vlSql = "SELECT num_archivo FROM pp_tmae_estcarhabdes "
    vlSql = vlSql & "WHERE "
    vlSql = vlSql & "num_perpago = '" & Trim(vlNumPerPago) & "' "
    Set vgRs = vgConexionBD.Execute(vlSql)
    If Not vgRs.EOF Then
        vlNumArchivo = vgRs!num_archivo
    Else
        MsgBox "No Existen Registros a Imprimir", vbCritical, "Operación Cancelada"
        Screen.MousePointer = 0
        Exit Sub
     End If
    vgRs.Close
    
    If vlNumArchivo <> "" Then
         vlSql = ""
         vlSql = "SELECT num_archivo FROM pp_ttmp_carhabdes "
         vlSql = vlSql & "WHERE "
         vlSql = vlSql & "num_archivo = '" & vlNumArchivo & "' "
         Set vgRs = vgConexionBD.Execute(vlSql)
         If vgRs.EOF Then
            MsgBox "No Existen Registros a Imprimir", vbCritical, "Operación Cancelada"
            Screen.MousePointer = 0
            Exit Sub
         End If
    End If

    vgQuery = "{PP_TTMP_CARHABDES.NUM_ARCHIVO} = " & Trim(vlNumArchivo) & " "
     
    Rpt_General.Reset
    Rpt_General.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
    Rpt_General.Connect = vgRutaDataBase       ' o App.Path & "\Nestle.mdb"
    Rpt_General.SelectionFormula = vgQuery
    Rpt_General.Formulas(0) = ""
    Rpt_General.Formulas(1) = ""
    Rpt_General.Formulas(2) = ""
    Rpt_General.Formulas(3) = "NombreCompania = '" & vgNombreCompania & "'"
    Rpt_General.Formulas(4) = "NombreSistema= '" & vgNombreSistema & "'"
    Rpt_General.Formulas(5) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
       
    Rpt_General.WindowState = crptMaximized
    Rpt_General.Destination = crptToWindow
    Rpt_General.WindowTitle = "Informe de Errores de Carga de Haberes Y Descuentos desde Archivo"
    Rpt_General.Action = 1
    Screen.MousePointer = 0
   
Exit Sub
Err_CmdImpErrores:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub

Private Sub Cmd_ImpResumen_Click()

On Error GoTo Err_CmdImpResumen

   'Validación de la Fecha ingresada en Periodo.
   If (Trim(Txt_FecCalculo.Text) = "") Then
      MsgBox "Debe ingresar una Fecha de Periodo", vbCritical, "Error de Datos"
      Exit Sub
   End If
   If Not IsDate(Txt_FecCalculo.Text) Then
      MsgBox "La Fecha ingresada no es una fecha válida.", vbCritical, "Error de Datos"
      Exit Sub
   End If
   If (Year(Txt_FecCalculo.Text) < 1900) Then
      MsgBox "La Fecha ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
      Exit Sub
   End If
   Txt_FecCalculo.Text = Format(CDate(Trim(Txt_FecCalculo.Text)), "yyyymmdd")
   Txt_FecCalculo.Text = DateSerial(Mid((Txt_FecCalculo.Text), 1, 4), Mid((Txt_FecCalculo.Text), 5, 2), Mid((Txt_FecCalculo.Text), 7, 2))
         
   Screen.MousePointer = 11
   
   vlArchivo = strRpt & "PP_Rpt_HabDesImportarEstadistica.rpt"   '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
      MsgBox "Archivo de Reporte Estadísticas de Carga de Haberes y Descuentos desde Archivo no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
      Screen.MousePointer = 0
      Exit Sub
   End If
   
    vlNumArchivo = ""
    vlNumPerPago = ""
    vlFechaPeriodoIni = ""
    vlFechaPeriodoTer = ""
    
    'SACA NUM PERIODO DE PAGO
    If Trim(Txt_FecCalculo) <> "" Then
        vlNumPerPago = Format(Trim(Txt_FecCalculo), "yyyymmdd")
        vlNumPerPago = Mid(vlNumPerPago, 1, 6)
        vlFechaPeriodoIni = vlNumPerPago & "01"
        vlFechaPeriodoTer = Format(DateSerial(CInt(Mid(vlNumPerPago, 1, 4)), CInt(Mid(vlNumPerPago, 5, 2)) + 1, 1 - 1), "yyyymmdd")
    End If
      
    'CONSULTA SI EXISTE EL PERIODO DE PAGO EN LA TABLA
    vlSql = ""
    vlSql = "SELECT num_archivo FROM pp_tmae_estcarhabdes "
    vlSql = vlSql & "WHERE "
    vlSql = vlSql & "num_perpago = '" & vlNumPerPago & "' "
    Set vgRs = vgConexionBD.Execute(vlSql)
    If Not vgRs.EOF Then
        vlNumArchivo = vgRs!num_archivo
    Else
        MsgBox "No Existen Registros a Imprimir", vbCritical, "Operación Cancelada"
        Screen.MousePointer = 0
        Exit Sub
    End If
    vgRs.Close

   vgQuery = "{PP_TMAE_ESTCARHABDES.NUM_ARCHIVO} = " & Trim(vlNumArchivo) & " "
     
   Rpt_General.Reset
   Rpt_General.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
   Rpt_General.Connect = vgRutaDataBase       ' o App.Path & "\Nestle.mdb"
   Rpt_General.SelectionFormula = vgQuery
   Rpt_General.Formulas(0) = ""
   Rpt_General.Formulas(1) = ""
   Rpt_General.Formulas(2) = ""
   
   Rpt_General.Formulas(3) = "NombreCompania = '" & vgNombreCompania & "'"
   Rpt_General.Formulas(4) = "NombreSistema= '" & vgNombreSistema & "'"
   Rpt_General.Formulas(5) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
      
   Rpt_General.WindowState = crptMaximized
   Rpt_General.Destination = crptToWindow
   Rpt_General.WindowTitle = "Informe de Estadísticas de Carga de Haberes y Descuentos desde Archivo"
   Rpt_General.Action = 1
   Screen.MousePointer = 0
   
Exit Sub
Err_CmdImpResumen:
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

Private Sub CmdPolizas_Click()
Dim ilargo As Long
On Error GoTo Err_Cmd

    vlArchivo = ""
    ComDialogo.CancelError = True
    ComDialogo.FileName = "*.txt"
    ComDialogo.DialogTitle = "Abrir Archivo de Haberes y Descuentos Masivos"
    ComDialogo.Filter = "*.txt"
    ComDialogo.FilterIndex = 1
    ComDialogo.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    ComDialogo.ShowOpen
    vlArchivo = ComDialogo.FileName
    Lbl_Archivo.Caption = vlArchivo
    If (Len(vlArchivo) > 60) Then
        While Len(Lbl_Archivo) > 60
            ilargo = InStr(1, Lbl_Archivo, "\")
            Lbl_Archivo = Mid(Lbl_Archivo, ilargo + 1, Len(Lbl_Archivo))
        Wend
        Lbl_Archivo.Caption = "\\" & Lbl_Archivo
    End If
Exit Sub
Err_Cmd:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Form_Load()
On Error GoTo Err_Cargar

    Frm_HabDesImportar.Top = 0
    Frm_HabDesImportar.Left = 0
    
    vlNumArchivo = ""
            
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub


Private Sub Txt_feccalculo_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then

    If (Trim(Txt_FecCalculo) = "") Then
       MsgBox "Debe Ingresar una Fecha de Cálculo", vbCritical, "Error de Datos"
       Txt_FecCalculo.SetFocus
       Exit Sub
    End If
    If Not IsDate(Txt_FecCalculo.Text) Then
       MsgBox "La Fecha ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
       Txt_FecCalculo.SetFocus
       Exit Sub
    End If
    If (CDate(Txt_FecCalculo) > CDate(Date)) Then
       MsgBox "La Fecha Ingresada es Mayor a la Fecha Actual", vbCritical, "Error de Datos"
       Txt_FecCalculo.SetFocus
       Exit Sub
    End If
    If (Year(Txt_FecCalculo) < 1900) Then
       MsgBox "La Fecha Ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
       Txt_FecCalculo.SetFocus
       Exit Sub
    End If
        
    Txt_FecCalculo = Format(CDate(Trim(Txt_FecCalculo)), "yyyymmdd")
    Txt_FecCalculo = DateSerial(Mid((Txt_FecCalculo), 1, 4), Mid((Txt_FecCalculo), 5, 2), Mid((Txt_FecCalculo), 7, 2))
    
    Cmd_Buscar.SetFocus
        
End If

End Sub

Private Sub Txt_feccalculo_LostFocus()

    If (Trim(Txt_FecCalculo) = "") Then
       Exit Sub
    End If
    If Not IsDate(Txt_FecCalculo.Text) Then
       Exit Sub
    End If
    If (CDate(Txt_FecCalculo) > CDate(Date)) Then
       Exit Sub
    End If
    If (Year(Txt_FecCalculo) < 1900) Then
       Exit Sub
    End If
        
    Txt_FecCalculo = Format(CDate(Trim(Txt_FecCalculo)), "yyyymmdd")
    Txt_FecCalculo = DateSerial(Mid((Txt_FecCalculo), 1, 4), Mid((Txt_FecCalculo), 5, 2), Mid((Txt_FecCalculo), 7, 2))

End Sub


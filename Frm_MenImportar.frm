VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "Crystl32.OCX"
Begin VB.Form Frm_MenImportar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carga de Mensajes desde Archivo."
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   8895
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
      Left            =   1200
      TabIndex        =   11
      Top             =   0
      Width           =   6375
      Begin VB.TextBox Txt_Periodo 
         Height          =   315
         Left            =   3600
         MaxLength       =   10
         TabIndex        =   1
         ToolTipText     =   "Fecha utilizada para validar los datos de Carga"
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Fecha de Cálculo            :"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   12
         Top             =   480
         Width           =   2175
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
      TabIndex        =   8
      Top             =   960
      Width           =   8655
      Begin VB.CommandButton Cmd_Buscar 
         Height          =   375
         Left            =   7680
         Picture         =   "Frm_MenImportar.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
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
         TabIndex        =   10
         Top             =   0
         Width           =   2415
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Archivo"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   480
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
   End
   Begin VB.Frame Fra_Operacion 
      Height          =   1095
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   8655
      Begin VB.CommandButton Cmd_ImpResumen 
         Caption         =   "&Resumen"
         Height          =   675
         Left            =   2880
         Picture         =   "Frm_MenImportar.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Imprimir Estadísticas"
         Top             =   240
         Width           =   790
      End
      Begin VB.CommandButton Cmd_ImpErrores 
         Caption         =   "&Errores"
         Height          =   675
         Left            =   4200
         Picture         =   "Frm_MenImportar.frx":07BC
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Imprimir Errores"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5400
         Picture         =   "Frm_MenImportar.frx":0E76
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Cargar 
         Caption         =   "&Cargar"
         Height          =   675
         Left            =   1680
         Picture         =   "Frm_MenImportar.frx":0F70
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Carga de Datos"
         Top             =   240
         Width           =   720
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
   Begin MSComDlg.CommonDialog ComDialogo 
      Left            =   360
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Frm_MenImportar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vlArchivo As String, linea As String, vlLargoArchivo As Long
Dim vlcont As Long, vlLargoRegistro As Long, vlAumento As Double

Dim vlSwEstPeriodo As String

Dim vlSql As String

Dim vlNumPoliza As String
Dim vlRutBen As String
Dim vlDgvBen As String
Dim vlCodMensaje As String
Dim vlFiller As String

Dim vlNumPerPago As String
Dim vlNumArchivo As String
Dim vlError As String
Dim vlNumMensajeErr As Integer
Dim vlNumMensajeOK As Integer
Dim vlNumEndoso As Integer
Dim vlNumOrden As Integer
Dim vlNumPerPagoAux As String

Dim vlFechaPeriodoIni As String
Dim vlFechaPeriodoTer As String

Dim vlGlsUsuarioCrea As Variant
Dim vlFecCrea As Variant
Dim vlHorCrea As Variant
Dim vlGlsUsuarioModi As Variant
Dim vlFecModi As Variant
Dim vlHorModi As Variant

Const clCodErrCero As Integer = 0
Const clCodTipoIng As String * 5 = "C"

'Mensajes de Error Utilizados dentro del formulario
'300 Rut del Asegurado no es numerico
'301 Rut del Asegurado no enviado
'302 Dígito Verificador del Asegurado no enviado
'303 Número de Póliza no válido
'304 Número de Póliza no enviado
'305 Código de Mensaje no es numérico
'306 Código de Mensaje no enviado
'307 El Dígito Verificador no corresponde al número de Rut
'308 El código de Mensaje no se encuentra registrado
'309 El número de Póliza no se encuentra registrado
'310 El número de Póliza y Rut de Beneficiario no se encuentran registrados
'311 El Registro ya se encuentra grabado en la Base de Datos


'------------------------ F U N C I O N E S --------------------------
        
        
Function flCargaArchivo()

On Error GoTo Err_Cargaarchivo

    Screen.MousePointer = 11
    
    'abre el archivo
    Open vlArchivo For Input As #1
    
    vlLargoArchivo = LOF(1)
    vlcont = 0
    vlLargoRegistro = 35   '79  '190 '70
    vlAumento = CDbl((98 / vlLargoArchivo) * vlLargoRegistro)
    
    
    
    'Saca el ultimo nro de archivo
    vlSql = "select NUM_ARCHIVO from PP_TMAE_ESTCARMENSAJE "
    vlSql = vlSql & "ORDER BY NUM_ARCHIVO DESC"
    Set vgRs = vgConexionBD.Execute(vlSql)
    If Not vgRs.EOF Then
        vlNumArchivo = CDbl(vgRs!num_archivo) + 1
    Else
        vlNumArchivo = "1"
    End If
    
    'SACA NUM PERIODO DE PAGO
    If Trim(Txt_Periodo.Text) <> "" Then
        vlNumPerPago = Format(Trim(Txt_Periodo.Text), "yyyymmdd")
        vlNumPerPagoAux = vlNumPerPago
        vlNumPerPago = Trim(Mid(vlNumPerPago, 1, 6))
        vlFechaPeriodoIni = vlNumPerPago & "01"
        vlFechaPeriodoTer = Format(DateSerial(CInt(Mid(vlNumPerPago, 1, 4)), CInt(Mid(vlNumPerPago, 5, 2)) + 1, 1 - 1), "yyyymmdd")

    End If
      
    'CONSULTA SI EXISTE EL PERIODO DE PAGO EN LA TABLA
    vlSql = "SELECT num_archivo FROM PP_TMAE_ESTCARMENSAJE "
    'vlSQL = "SELECT num_archivo,num_mensajeok FROM PP_TTMP_CARMENpoliza "
    vlSql = vlSql & " WHERE  "
    vlSql = vlSql & " num_perpago >= '" & vlFechaPeriodoIni & "' AND "
    vlSql = vlSql & " num_perpago <= '" & vlFechaPeriodoTer & "' "
    
'    If vgTipoBase = "ORACLE" Then
'       vlSQL = vlSQL & " substr(num_perpago,1,6)= '" & vlNumPerPago & "' "
'    Else
'        vlSQL = vlSQL & " substring(num_perpago,1,6)= '" & vlNumPerPago & "' "
'    End If
    Set vgRegistro = vgConexionBD.Execute(vlSql)
    If Not vgRegistro.EOF Then
       vlNumArchivo = (vgRegistro!num_archivo)
       vlSql = "DELETE FROM PP_TTMP_CARMENPOLIZA WHERE "
'       If vgTipoBase = "ORACLE" Then
'          vlSQL = vlSQL & "substr(num_perpago,1,6) = '" & vlNumPerPago & "' AND "
'       Else
'           vlSQL = vlSQL & "substring(num_perpago,1,6) = '" & vlNumPerPago & "' AND "
'       End If
       vlSql = vlSql & "num_archivo = " & vlNumArchivo & " "
       'vlSQL = vlSQL & "substr(num_perpago,1,6) = '" & vlNumPerPago & "' "
       vgConexionBD.Execute vlSql
    End If
    
If Not EOF(1) Then
    
       'Validar si Periodo se encuentra abierto
       'vlSwEstPeriodo = C    Periodo Cerrado
       'vlSwEstPeriodo = A    Periodo Abierto
       vlSwEstPeriodo = "C"
       
       vgSql = ""
       vgSql = "SELECT * "
       vgSql = vgSql & "FROM PP_TMAE_PROPAGOPEN "
       vgSql = vgSql & "WHERE "
       vgSql = vgSql & "num_perpago = '" & Trim(vlNumPerPago) & "' AND "
       'vgSql = vgSql & "(cod_estadoreg = 'A' OR cod_estadopri = 'A') "
       vgSql = vgSql & "(cod_estadoreg <> 'C' OR cod_estadopri <> 'C') "
       Set vgRs = vgConexionBD.Execute(vgSql)
       If vgRs.EOF Then
          'Periodo Cerrado - "C"
          MsgBox "El Periodo Ingresado se encuentra Cerrado, Debe Ingresar una Nueva Fecha de Cálculo.", vbCritical, "Error de Datos"
          Txt_Periodo.SetFocus
          Exit Function
       End If
    
    
        Frm_BarraProg.Show
        Frm_BarraProg.Refresh
        Frm_BarraProg.ProgressBar1.Value = 0
        Frm_BarraProg.Lbl_Texto = "Cargando Archivo de Mensajes Masivos" & vlArchivo
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
                
        vlNumPoliza = ""
        vlRutBen = ""
        vlDgvBen = ""
        vlCodMensaje = ""
        vlFiller = ""
        
        vlNumPoliza = Trim(Mid(linea, 1, 10))
        vlRutBen = Trim(Mid(linea, 11, 9))
        vlDgvBen = UCase(Trim(Mid(linea, 20, 1)))
        vlCodMensaje = Trim(Mid(linea, 21, 10))
        
        vgFiller = UCase(Trim(Mid(linea, 32, 10)))
        
        Call flValidarDatos
              
        Call flGrabaDatos(vlNumArchivo, vlFechaPeriodoIni)
                
        If Frm_BarraProg.ProgressBar1.Value + vlAumento < 100 Then
           Frm_BarraProg.ProgressBar1.Value = Frm_BarraProg.ProgressBar1.Value + vlAumento
           Frm_BarraProg.ProgressBar1.Refresh
        End If
            
    Loop
    MsgBox "El Archivo ha Sido Cargado con Exito ", vbInformation, "Proceso De Carga Finalizado"
    
    Close #1
    
    'Traspaso de datos a tablas definitivas
    Call flValidacionesConsistencia
    Call flGrabarDatosDefinitivos
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
Function flGrabaDatos(iarchivo As String, iPerPago As String)

On Error GoTo Err_flGrabaDatos

    vlSql = "INSERT INTO PP_TTMP_CARMENPOLIZA ("
    vlSql = vlSql & "num_archivo,"
    vlSql = vlSql & "num_perpago,"
    
    If vlNumPoliza <> "" Then vlSql = vlSql & "num_poliza, "
    If vlRutBen <> "" Then vlSql = vlSql & "rut_ben, "
    If vlDgvBen <> "" Then vlSql = vlSql & "dgv_ben, "
    If vlCodMensaje <> "" Then vlSql = vlSql & "cod_mensaje, "
    If vgFiller <> "" Then vlSql = vlSql & "gls_filler, "
    
    vlSql = vlSql & "cod_error) VALUES ("
    vlSql = vlSql & " " & iarchivo & ","
    vlSql = vlSql & "'" & iPerPago & "',"
    
    If vlNumPoliza <> "" Then vlSql = vlSql & "'" & vlNumPoliza & "', "
    If vlRutBen <> "" Then vlSql = vlSql & "'" & vlRutBen & "', "
    If vlDgvBen <> "" Then vlSql = vlSql & "'" & vlDgvBen & "', "
    If vlCodMensaje <> "" Then vlSql = vlSql & "'" & vlCodMensaje & "', "
    If vgFiller <> "" Then vlSql = vlSql & "'" & vgFiller & "', "
    
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

Function flValidarDatos()

On Error GoTo Err_flValidarDatos

    If vlNumPoliza = "" Then
       vlError = 304
       Exit Function
    End If
    If vlRutBen = "" Then
       vlError = 301
       Exit Function
    Else
        If Not IsNumeric(vlRutBen) Then
           vlError = 300
           Exit Function
        End If
    End If
    If vlDgvBen = "" Then
       vlError = 302
       Exit Function
    End If
    If Not ValiRut(vlRutBen, vlDgvBen) Then
       vlError = 307
       Exit Function
    End If
    If vlCodMensaje = "" Then
       vlError = 306
       Exit Function
    Else
        If Not IsNumeric(vlCodMensaje) Then
           vlError = 305
           Exit Function
        End If
    End If

Exit Function
Err_flValidarDatos:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function flGrabarDatosDefinitivos()

On Error GoTo Err_GrabarDatosDefinitivos

'Selecciona todos los registros de la tabla temporal de mensajes que se encuentren
'sin errores, para generar los registros para la tabla definitiva. Se seleccionan
'además datos de otras tablas para completar el registro definivo.
    vlSql = ""
    vlSql = " SELECT c.num_archivo,c.num_perpago,c.num_poliza,c.rut_ben,c.cod_mensaje, "
    vlSql = vlSql & " p.num_endoso, "
    vlSql = vlSql & " b.Num_Orden "
    vlSql = vlSql & " FROM PP_TTMP_CARMENPOLIZA c,PP_TMAE_POLIZA p, "
    vlSql = vlSql & " PP_TMAE_BEN b "
    vlSql = vlSql & " WHERE "
    vlSql = vlSql & " (c.cod_error = 0) AND "
    vlSql = vlSql & " (c.num_poliza = p.num_poliza) AND "
    vlSql = vlSql & " (c.num_poliza = b.num_poliza) AND "
    vlSql = vlSql & " (p.num_poliza = b.num_poliza) AND "
    vlSql = vlSql & " (c.rut_ben = b.rut_ben) AND "
    vlSql = vlSql & " p.num_endoso = "
    vlSql = vlSql & " (SELECT MAX(num_endoso) FROM PP_TMAE_POLIZA WHERE num_poliza = p.num_poliza ) AND "
    vlSql = vlSql & " p.num_endoso = b.num_endoso "
    
'''    vlSQL = vlSQL & " b.num_orden = "
'''    vlSQL = vlSQL & " (SELECT num_orden FROM PP_TMAE_BEN WHERE (num_poliza = p.num_poliza) AND "
'''    vlSQL = vlSQL & " (num_endoso = p.num_endoso) AND (rut_ben = b.rut_ben)) "
    
    Set vgRegistro = vgConexionBD.Execute(vlSql)
    If Not vgRegistro.EOF Then
       
       While Not vgRegistro.EOF
       
'Asignando los datos seleccionados a las variables correspondientes
             vlNumArchivo = (vgRegistro!num_archivo)
             vlNumPoliza = (vgRegistro!num_poliza)
             vlNumEndoso = (vgRegistro!num_endoso)
             vlNumOrden = (vgRegistro!Num_Orden)
             vlCodMensaje = (vgRegistro!cod_mensaje)
             'vlNumPerPago = (vgRegistro!Num_PerPago)
             vlNumPerPago = vlFechaPeriodoIni
             vlNumPerPagoAux = vlNumPerPago
             'vlNumPerPago = Mid((vlNumPerPago), 1, 6)
             
             vlRutBen = (vgRegistro!Rut_Ben)
             
'Valida que el registro a grabar, no se encuentra registrado en la tabla definitiva
             vgSql = ""
             vgSql = "SELECT num_poliza FROM PP_TMAE_MENPOLIZA "
             vgSql = vgSql & "WHERE num_poliza = '" & (vlNumPoliza) & "' AND "
             vgSql = vgSql & "num_endoso = " & Trim(vlNumEndoso) & " AND "
             vgSql = vgSql & "num_orden = " & Trim(vlNumOrden) & " AND "
             vgSql = vgSql & "num_perpago = '" & Mid(vlFechaPeriodoIni, 1, 6) & "' AND "
             vgSql = vgSql & "cod_mensaje = " & CInt(vlCodMensaje) & " "
             Set vgRs = vgConexionBD.Execute(vgSql)
             If Not vgRs.EOF Then
                
                vgSql = ""
                vgSql = "UPDATE PP_TTMP_CARMENPOLIZA t "
                vgSql = vgSql & "SET cod_error = 311 "
                vgSql = vgSql & "WHERE cod_error = " & clCodErrCero & " AND "
                vgSql = vgSql & "num_archivo = " & (vgRegistro!num_archivo) & " AND "
                vgSql = vgSql & "num_perpago = '" & Trim(vlFechaPeriodoIni) & "' AND "
                vgSql = vgSql & "num_poliza = '" & (vlNumPoliza) & "' AND "
                vgSql = vgSql & "rut_ben = " & Trim(vlRutBen) & " "
                vgConexionBD.Execute (vgSql)
                
             Else
             
                 vlGlsUsuarioCrea = vgUsuario
                 vlFecCrea = Format(Date, "yyyymmdd")
                 vlHorCrea = Format(Time, "hhmmss")
                              
'Se realiza la grabación del registro en la tabla de mensajes definitiva.
                 vlSql = ""
                 vlSql = "INSERT INTO PP_TMAE_MENPOLIZA ("
                 vlSql = vlSql & "num_poliza,"
                 vlSql = vlSql & "num_endoso,"
                 vlSql = vlSql & "num_orden, "
                 vlSql = vlSql & "cod_mensaje, "
                 vlSql = vlSql & "num_perpago, "
                 vlSql = vlSql & "cod_tipoing, "
                 vlSql = vlSql & "cod_usuariocrea, "
                 vlSql = vlSql & "fec_crea, "
                 vlSql = vlSql & "hor_crea "
                 
                 vlSql = vlSql & ") VALUES ("
                 vlSql = vlSql & " '" & Trim(vlNumPoliza) & "',"
                 vlSql = vlSql & "" & Str(vlNumEndoso) & ","
                 vlSql = vlSql & "" & Str(vlNumOrden) & ", "
                 vlSql = vlSql & "" & Str(vlCodMensaje) & ", "
                 vlSql = vlSql & "'" & Mid(vlFechaPeriodoIni, 1, 6) & "', "
                 vlSql = vlSql & "'" & Trim(clCodTipoIng) & "', "
                 vlSql = vlSql & "'" & vlGlsUsuarioCrea & "', "
                 vlSql = vlSql & "'" & vlFecCrea & "', "
                 vlSql = vlSql & "'" & vlHorCrea & "' ) "
                 
                 vgConexionBD.Execute (vlSql)
                                  
             End If
             
             vgRegistro.MoveNext
       Wend
    End If

Exit Function
Err_GrabarDatosDefinitivos:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flValidacionesConsistencia()

On Error GoTo Err_ValidacionesConsistencia

'Validar que cod_mensaje se encuentra registrado en Tabla de Mensajes
    vgSql = ""
    vgSql = "UPDATE PP_TTMP_CARMENPOLIZA "
    vgSql = vgSql & " SET cod_error = 308 "
    vgSql = vgSql & " WHERE cod_error = " & clCodErrCero & " AND "
    vgSql = vgSql & " cod_mensaje NOT IN "
    vgSql = vgSql & " (SELECT cod_mensaje FROM PP_TPAR_MENSAJE )"
    vgConexionBD.Execute (vgSql)

'Valida que num_poliza se encuentra registrado en Tabla de Polizas
    vgSql = ""
    vgSql = "UPDATE PP_TTMP_CARMENPOLIZA "
    vgSql = vgSql & " SET cod_error = 309 "
    vgSql = vgSql & " WHERE cod_error = " & clCodErrCero & " AND "
    vgSql = vgSql & " num_poliza NOT IN "
    vgSql = vgSql & " (SELECT num_poliza FROM PP_TMAE_POLIZA )"
    vgConexionBD.Execute (vgSql)
    
    
'Valida que el num_poliza y rut_ben se encuentren registrados en la Tabla
'de Beneficiarios
    vgSql = ""
    vgSql = "UPDATE PP_TTMP_CARMENPOLIZA t "
    vgSql = vgSql & " SET t.cod_error = 310 "
    vgSql = vgSql & " WHERE t.cod_error = " & clCodErrCero & " AND "
    vgSql = vgSql & " t.num_poliza NOT IN "
    vgSql = vgSql & " (SELECT b.num_poliza FROM PP_TMAE_BEN b "
    vgSql = vgSql & " WHERE b.num_poliza = t.num_poliza AND "
    vgSql = vgSql & " b.rut_ben = t.rut_ben) "
    vgConexionBD.Execute (vgSql)

'''''''Valida que el registro a traspasar a la Tabla Definitiva, no se encuentra ya
'''''''registrado
''''''    vgSql = ""
''''''    vgSql = "UPDATE PP_TTMP_CARMENPOLIZA t"
''''''    vgSql = vgSql & " SET cod_error = 310 "
''''''    vgSql = vgSql & " WHERE cod_error = " & clCodErrCero & " AND "
''''''    vgSql = vgSql & " t.num_poliza IN "
''''''    vgSql = vgSql & " (SELECT num_poliza FROM PP_TMAE_MENPOLIZA "
''''''    vgSql = vgSql & " WHERE num_poliza = t.num_poliza AND "
''''''    vgSql = vgSql & " rut_ben = t.rut_ben AND "
''''''    vgSql = vgSql & " num_perpago = t.num_perpago AND "
''''''    vgSql = vgSql & " cod_mensaje = t.cod_mensaje) "
''''''    vgConexionBD.Execute (vgSql)

Exit Function
Err_ValidacionesConsistencia:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function flGrabarEstadistica()

On Error GoTo Err_flGrabarEstadistica


    vlNumMensajeErr = 0
    vlNumMensajeOK = 0
    
'Obtiene el total de registros ingresados CON Error en la tabla de temporales
    vlSql = ""
    vlSql = "SELECT count (num_archivo) as TotalRegistrosErr "
    vlSql = vlSql & " FROM PP_TTMP_CARMENPOLIZA "
    vlSql = vlSql & " WHERE cod_error <> " & clCodErrCero & " AND "
    vlSql = vlSql & " num_archivo = " & vlNumArchivo & " "
    Set vgRegistro = vgConexionBD.Execute(vlSql)
    If Not vgRegistro.EOF Then
       vlNumMensajeErr = (vgRegistro!totalregistrosErr)
    End If


'Obtiene el total de registros ingresados SIN Error en la tabla de temporales
    vlSql = ""
    vlSql = "SELECT count (num_archivo) as TotalRegistrosOK "
    vlSql = vlSql & " FROM PP_TTMP_CARMENPOLIZA "
    vlSql = vlSql & " WHERE cod_error = " & clCodErrCero & " AND "
    vlSql = vlSql & " num_archivo = " & vlNumArchivo & " "
    Set vgRegistro = vgConexionBD.Execute(vlSql)
    If Not vgRegistro.EOF Then
       vlNumMensajeOK = (vgRegistro!totalregistrosOK)
    End If
    
    
    vlSql = "SELECT num_archivo "
    vlSql = vlSql & " FROM PP_TMAE_ESTCARMENSAJE "
    vlSql = vlSql & " WHERE num_archivo = " & vlNumArchivo & " AND "
    vlSql = vlSql & " num_perpago = " & vlFechaPeriodoIni & " "
    Set vgRegistro = vgConexionBD.Execute(vlSql)
    If Not vgRegistro.EOF Then
           
       vlGlsUsuarioModi = vgUsuario
       vlFecModi = Format(Date, "yyyymmdd")
       vlHorModi = Format(Time, "hhmmss")
       
       vgSql = ""
       vgSql = "UPDATE PP_TMAE_ESTCARMENSAJE "
       vgSql = vgSql & " SET num_mensajeerr = " & vlNumMensajeErr & ", "
       vgSql = vgSql & " num_mensajeok = " & vlNumMensajeOK & ", "
       vgSql = vgSql & " cod_usuariomodi= '" & vlGlsUsuarioModi & "', "
       vgSql = vgSql & " fec_modi = '" & vlFecModi & "', "
       vgSql = vgSql & " hor_modi = '" & vlHorModi & "' "
       vgSql = vgSql & " WHERE num_archivo = " & vlNumArchivo & " AND "
       vgSql = vgSql & " num_perpago = " & vlFechaPeriodoIni & " "
       vgConexionBD.Execute (vgSql)

    
    Else
        
        vlGlsUsuarioCrea = vgUsuario
        vlFecCrea = Format(Date, "yyyymmdd")
        vlHorCrea = Format(Time, "hhmmss")
       
        vlSql = ""
        vlSql = "INSERT INTO PP_TMAE_ESTCARMENSAJE ("
        vlSql = vlSql & "num_archivo,"
        vlSql = vlSql & "num_perpago,"
        vlSql = vlSql & "num_mensajeerr, "
        vlSql = vlSql & "num_mensajeok, "
        vlSql = vlSql & "cod_usuariocrea, "
        vlSql = vlSql & "fec_crea, "
        vlSql = vlSql & "hor_crea, "
        vlSql = vlSql & "cod_usuariomodi, "
        vlSql = vlSql & "fec_modi, "
        vlSql = vlSql & "hor_modi "
       
        vlSql = vlSql & ") VALUES ("
        vlSql = vlSql & " " & Str(vlNumArchivo) & ","
        vlSql = vlSql & "'" & (vlFechaPeriodoIni) & "',"
        vlSql = vlSql & "" & Str(vlNumMensajeErr) & ", "
        vlSql = vlSql & "" & Str(vlNumMensajeOK) & ", "
        vlSql = vlSql & "'" & vlGlsUsuarioCrea & "', "
        vlSql = vlSql & "'" & vlFecCrea & "', "
        vlSql = vlSql & "'" & vlHorCrea & "', "
        vlSql = vlSql & "'" & vlGlsUsuarioModi & "', "
        vlSql = vlSql & "'" & vlFecModi & "', "
        vlSql = vlSql & "'" & vlHorModi & "' ) "
                
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

    vlSql = "DELETE FROM PP_TTMP_CARMENPOLIZA "
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
    If Not IsDate(Txt_Periodo) Then
        MsgBox "La Fecha de Proceso ingresada no es válida.", vbCritical, "Operación Cancelada"
        Screen.MousePointer = 0
        Txt_Periodo.SetFocus
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
    ComDialogo.DialogTitle = "Archivo de Mensajes Masivos"
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
   
   If (Trim(Txt_Periodo.Text) = "") Then
      MsgBox "Debe ingresar una Fecha de Periodo", vbCritical, "Error de Datos"
      Exit Sub
   End If
   If Not IsDate(Txt_Periodo.Text) Then
      MsgBox "La Fecha ingresada no es una fecha válida.", vbCritical, "Error de Datos"
      Exit Sub
   End If
'    If (CDate(Txt_Periodo.Text) > CDate(Date)) Then
'       MsgBox "La Fecha ingresada es mayor a la fecha actual", vbCritical, "Error de Datos"
'       Lbl_Hasta.Caption = ""
'       Txt_Desde.SetFocus
'       Exit Sub
'    End If
   If (Year(Txt_Periodo.Text) < 1900) Then
      MsgBox "La Fecha ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
      Exit Sub
   End If
    
   Txt_Periodo.Text = Format(CDate(Trim(Txt_Periodo.Text)), "yyyymmdd")
   Txt_Periodo.Text = DateSerial(Mid((Txt_Periodo.Text), 1, 4), Mid((Txt_Periodo.Text), 5, 2), Mid((Txt_Periodo.Text), 7, 2))
        



   Screen.MousePointer = 11
   
   vlArchivo = strRpt & "PP_Rpt_MenImportarErr.rpt"   '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
      MsgBox "Archivo de Reporte de Listado de Errores de Carga de Mensajes desde Archivo no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   '-----------------------------
'CMV-20050321 I

    vlNumArchivo = ""
    vlNumPerPago = ""
    vlFechaPeriodoIni = ""
    vlFechaPeriodoTer = ""
    
    'SACA NUM PERIODO DE PAGO
    If Trim(Txt_Periodo) <> "" Then
        vlNumPerPago = Format(Trim(Txt_Periodo), "yyyymmdd")
        vlNumPerPago = Mid(vlNumPerPago, 1, 6)
        vlFechaPeriodoIni = vlNumPerPago & "01"
        vlFechaPeriodoTer = Format(DateSerial(CInt(Mid(vlNumPerPago, 1, 4)), CInt(Mid(vlNumPerPago, 5, 2)) + 1, 1 - 1), "yyyymmdd")
    End If
      
    'CONSULTA SI EXISTE EL PERIODO DE PAGO EN LA TABLA
    vlSql = "select NUM_ARCHIVO from PP_TMAE_ESTCARMENSAJE "
    vlSql = vlSql & "where "
    vlSql = vlSql & "num_perpago >= '" & vlFechaPeriodoIni & "' and "
    vlSql = vlSql & "num_perpago <= '" & vlFechaPeriodoTer & "' "
    Set vgRs = vgConexionBD.Execute(vlSql)
    If Not vgRs.EOF Then
        vlNumArchivo = vgRs!num_archivo
    Else
        MsgBox "No Existen Registros a Imprimir", vbCritical, "Operación Cancelada"
        Exit Sub
     End If
    vgRs.Close
    
    If vlNumArchivo <> "" Then
         vlSql = "SELECT num_archivo FROM PP_TTMP_CARMENPOLIZA "
         vlSql = vlSql & "WHERE "
         vlSql = vlSql & "num_archivo = '" & vlNumArchivo & "' "
         Set vgRs = vgConexionBD.Execute(vlSql)
         If vgRs.EOF Then
            MsgBox "No Existen Registros a Imprimir", vbCritical, "Operación Cancelada"
            Exit Sub
         End If
    End If

'CMV-20050321 F
'-----------------------------
    
   vgQuery = "{PP_TTMP_CARMENPOLIZA.NUM_ARCHIVO} = " & Trim(vlNumArchivo) & " "
     
   Rpt_General.Reset
   Rpt_General.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
   Rpt_General.Connect = vgRutaDataBase       ' o App.Path & "\Nestle.mdb"
   Rpt_General.SelectionFormula = vgQuery
   Rpt_General.Formulas(0) = ""
   '"RutPensionado = '" & (Trim(Txt_PenRut.Text)) & " - " & (Trim(Txt_PenDigito.Text)) & "' "
   Rpt_General.Formulas(1) = ""
   '"NombrePensionado = '" & Trim(Lbl_PenNombre.Caption) & "' "
   Rpt_General.Formulas(2) = ""
   
   Rpt_General.Formulas(3) = "NombreCompania = '" & vgNombreCompania & "'"
   Rpt_General.Formulas(4) = "NombreSistema= '" & vgNombreSistema & "'"
   Rpt_General.Formulas(5) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
      
   Rpt_General.WindowState = crptMaximized
   Rpt_General.Destination = crptToWindow
   Rpt_General.WindowTitle = "Informe de Errores de Carga de Mensajes desde Archivo"
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
   
   If (Trim(Txt_Periodo.Text) = "") Then
      MsgBox "Debe ingresar una Fecha de Periodo", vbCritical, "Error de Datos"
      Exit Sub
   End If
   If Not IsDate(Txt_Periodo.Text) Then
      MsgBox "La Fecha ingresada no es una fecha válida.", vbCritical, "Error de Datos"
      Exit Sub
   End If
'    If (CDate(Txt_Periodo.Text) > CDate(Date)) Then
'       MsgBox "La Fecha ingresada es mayor a la fecha actual", vbCritical, "Error de Datos"
'       Lbl_Hasta.Caption = ""
'       Txt_Desde.SetFocus
'       Exit Sub
'    End If
   If (Year(Txt_Periodo.Text) < 1900) Then
      MsgBox "La Fecha ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
      Exit Sub
   End If
    
   Txt_Periodo.Text = Format(CDate(Trim(Txt_Periodo.Text)), "yyyymmdd")
   Txt_Periodo.Text = DateSerial(Mid((Txt_Periodo.Text), 1, 4), Mid((Txt_Periodo.Text), 5, 2), Mid((Txt_Periodo.Text), 7, 2))
    
    
    
    
   Screen.MousePointer = 11
   
   vlArchivo = strRpt & "PP_Rpt_MenImportarEstadistica.rpt"   '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
      MsgBox "Archivo de Reporte Estadísticas de Carga de Mensajes desde Archivo no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
      Screen.MousePointer = 0
      Exit Sub
   End If
   
'-----------------------------
'CMV-20050321 I

    vlNumArchivo = ""
    vlNumPerPago = ""
    vlFechaPeriodoIni = ""
    vlFechaPeriodoTer = ""
    
    'SACA NUM PERIODO DE PAGO
    If Trim(Txt_Periodo) <> "" Then
        vlNumPerPago = Format(Trim(Txt_Periodo), "yyyymmdd")
        vlNumPerPago = Mid(vlNumPerPago, 1, 6)
        vlFechaPeriodoIni = vlNumPerPago & "01"
        vlFechaPeriodoTer = Format(DateSerial(CInt(Mid(vlNumPerPago, 1, 4)), CInt(Mid(vlNumPerPago, 5, 2)) + 1, 1 - 1), "yyyymmdd")
    End If
      
    'CONSULTA SI EXISTE EL PERIODO DE PAGO EN LA TABLA
    vlSql = "select NUM_ARCHIVO from PP_TMAE_ESTCARMENSAJE "
    vlSql = vlSql & "where "
    vlSql = vlSql & "num_perpago >= '" & vlFechaPeriodoIni & "' and "
    vlSql = vlSql & "num_perpago <= '" & vlFechaPeriodoTer & "' "
    Set vgRs = vgConexionBD.Execute(vlSql)
    If Not vgRs.EOF Then
        vlNumArchivo = vgRs!num_archivo
    Else
        MsgBox "No Existen Registros a Imprimir", vbCritical, "Operación Cancelada"
        Exit Sub
    End If
    vgRs.Close

'CMV-20050321 F
'-----------------------------
   
   vgQuery = "{PP_TMAE_ESTCARMENSAJE.NUM_ARCHIVO} = " & Trim(vlNumArchivo) & " "
     
   Rpt_General.Reset
   Rpt_General.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
   Rpt_General.Connect = vgRutaDataBase       ' o App.Path & "\Nestle.mdb"
   Rpt_General.SelectionFormula = vgQuery
   Rpt_General.Formulas(0) = ""
   '"RutPensionado = '" & (Trim(Txt_PenRut.Text)) & " - " & (Trim(Txt_PenDigito.Text)) & "' "
   Rpt_General.Formulas(1) = ""
   '"NombrePensionado = '" & Trim(Lbl_PenNombre.Caption) & "' "
   Rpt_General.Formulas(2) = ""
   
   Rpt_General.Formulas(3) = "NombreCompania = '" & vgNombreCompania & "'"
   Rpt_General.Formulas(4) = "NombreSistema= '" & vgNombreSistema & "'"
   Rpt_General.Formulas(5) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
      
   Rpt_General.WindowState = crptMaximized
   Rpt_General.Destination = crptToWindow
   Rpt_General.WindowTitle = "Informe de Estadísticas de Carga de Mensajes desde Archivo"
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
    ComDialogo.DialogTitle = "Abrir Archivo de Mensajes Masivos"
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

    Frm_MenImportar.Top = 0
    Frm_MenImportar.Left = 0
    
    vlNumArchivo = ""
            
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub


Private Sub Txt_periodo_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then

    If (Trim(Txt_Periodo) = "") Then
       MsgBox "Debe Ingresar una Fecha de Cálculo", vbCritical, "Error de Datos"
       Txt_Periodo.SetFocus
       Exit Sub
    End If
    If Not IsDate(Txt_Periodo.Text) Then
       MsgBox "La Fecha ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
       Txt_Periodo.SetFocus
       Exit Sub
    End If
    If (CDate(Txt_Periodo) > CDate(Date)) Then
       MsgBox "La Fecha Ingresada es Mayor a la Fecha Actual", vbCritical, "Error de Datos"
       Txt_Periodo.SetFocus
       Exit Sub
    End If
    If (Year(Txt_Periodo) < 1900) Then
       MsgBox "La Fecha Ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
       Txt_Periodo.SetFocus
       Exit Sub
    End If
        
    Txt_Periodo.Text = Format(CDate(Trim(Txt_Periodo)), "yyyymmdd")
    Txt_Periodo.Text = DateSerial(Mid((Txt_Periodo.Text), 1, 4), Mid((Txt_Periodo.Text), 5, 2), Mid((Txt_Periodo.Text), 7, 2))
    
    Cmd_Buscar.SetFocus
        
End If

End Sub

Private Sub Txt_periodo_LostFocus()

    If (Trim(Txt_Periodo) = "") Then
       Exit Sub
    End If
    If Not IsDate(Txt_Periodo.Text) Then
       Exit Sub
    End If
    If (CDate(Txt_Periodo) > CDate(Date)) Then
       Exit Sub
    End If
    If (Year(Txt_Periodo) < 1900) Then
       Exit Sub
    End If
        
    Txt_Periodo.Text = Format(CDate(Trim(Txt_Periodo)), "yyyymmdd")
    Txt_Periodo.Text = DateSerial(Mid((Txt_Periodo.Text), 1, 4), Mid((Txt_Periodo.Text), 5, 2), Mid((Txt_Periodo.Text), 7, 2))


'Txt_Periodo = Trim(Txt_Periodo)
'If IsDate(Txt_Periodo) Then
'    Txt_Periodo.Text = Format(CDate(Trim(Txt_Periodo)), "yyyymmdd")
'    Txt_Periodo.Text = DateSerial(Mid((Txt_Periodo.Text), 1, 4), Mid((Txt_Periodo.Text), 5, 2), Mid((Txt_Periodo.Text), 7, 2))
'Else
'    Txt_Periodo = ""
'End If
End Sub


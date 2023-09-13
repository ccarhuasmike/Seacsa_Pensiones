VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_InfConAsigFam 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Asignación Familiar"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   4950
   Begin VB.Frame Frame3 
      Caption         =   "Progreso del Cálculo"
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
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   4695
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.Frame Fra_Operacion 
      Height          =   1095
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   4695
      Begin VB.CommandButton cmd_calcular 
         Caption         =   "&Calcular"
         Height          =   675
         Left            =   960
         Picture         =   "Frm_InfConAsigFam.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Generar Tabla de Mortalidad"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   2040
         Picture         =   "Frm_InfConAsigFam.frx":04A2
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Imprimir Reporte"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   3120
         Picture         =   "Frm_InfConAsigFam.frx":0B5C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
      Begin Crystal.CrystalReport Rpt_Calculo 
         Left            =   120
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.TextBox Txt_Anno 
         Height          =   285
         Left            =   2400
         MaxLength       =   4
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Txt_Mes 
         Height          =   285
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   5
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "( Mes - Año )"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3120
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   " -"
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de Pago "
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   2295
      End
   End
End
Attribute VB_Name = "Frm_InfConAsigFam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vlRegistro As ADODB.Recordset

Dim stConLiquidacion As TyLiquidacion
Dim stConDetPension As TyDetPension
Dim stConPagoAsig As TyPagoAsignacion
Dim stConTutor As TyTutor
Dim stDatosInformeControl As TyDatosInformeControl
Private Sub cmd_calcular_Click()
If Txt_Mes = "" Then
    Exit Sub
End If
If Txt_Anno = "" Then
    Exit Sub
End If

'Asignación Familiar
If Not flLlenaTablaTemporal(Txt_Mes, Txt_Anno, "AF") Then
    Exit Sub
End If

'Retención Judicial
If Not flLlenaTablaTemporal(Txt_Mes, Txt_Anno, "RJ") Then
    Exit Sub
End If

'Cajas de Compensación
If Not flLlenaTablaTemporal(Txt_Mes, Txt_Anno, "CCAF") Then
    Exit Sub
End If

MsgBox "Calculo para Informe de Control realizado exitosamente", vbInformation, Me.Caption

End Sub

Private Sub Cmd_Imprimir_Click()
On Error GoTo Err_Imprimir

   vlArchivo = App.Path & "\Reportes\PP_Rpt_ConAsigFam.rpt"   '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
       MsgBox "Archivo de Reporte de Póliza no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
       Screen.MousePointer = 0
       Exit Sub
   End If
      
   Call flCargaTemporal
    
            
   ' vgQuery = "{PP_TMAE_ASIGFAM.FEC_INIACTIVA}<= '" & vlFecha & "' AND "
   ' vgQuery = vgQuery & "{PP_TMAE_ASIGFAM.FEC_TERACTIVA}>= '" & vlFecha & "'"

    Rpt_Calculo.Reset
    Rpt_Calculo.WindowState = crptMaximized
    Rpt_Calculo.ReportFileName = vlArchivo
    Rpt_Calculo.Connect = vgRutaDataBase
    Rpt_Calculo.SelectionFormula = ""
    'Rpt_Calculo.SelectionFormula = vgQuery
    
    Rpt_Calculo.Formulas(0) = "NombreCompania='" & vgNombreCompania & "'"
    Rpt_Calculo.Formulas(1) = "NombreSistema= '" & vgNombreSistema & "'"
    Rpt_Calculo.Formulas(2) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
    Rpt_Calculo.Formulas(3) = "Fecha= '" & vlFecha & "'"
    Rpt_Calculo.SubreportToChange = ""
    Rpt_Calculo.Destination = crptToWindow
    Rpt_Calculo.WindowTitle = "Informe de Control Asignación Familiar"
    Rpt_Calculo.Action = 1
    Screen.MousePointer = 0
    


Exit Sub
Err_Imprimir:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select


End Sub

Private Sub Cmd_Salir_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Frm_InfConAsigFam.Top = 0
    Frm_InfConAsigFam.Left = 0
End Sub

Private Sub Txt_Anno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Txt_Anno <> "" Then
       If Txt_Anno < 1900 Then
           MsgBox "Año ingresado es menor a la mínima que se puede ingresar (1900).", vbCritical, "Dato Incorrecto"
       Else
         cmd_calcular.SetFocus
       End If
    End If
End If
End Sub

Function flCargaTemporal()
On Error GoTo Err_Carga

   'vlInicio = DateSerial(Txt_Anno, Txt_Mes, Format("01", "00"))
   'vltermino = DateSerial(Txt_Anno, Txt_Mes + 1, 1 - 1)
   
   vlNumPerPago = Txt_Anno + Txt_Mes
   
   vgSql = ""
   vgSql = "SELECT NUM_PERPAGO,NUM_POLIZA,NUM_ORDEN,COD_CONHABDES, RUT_RECEPTOR"
   vgSql = vgSql & " FROM PP_TTMP_CONPAGOPEN"
   vgSql = vgSql & " WHERE COD_TIPOMOD = 'AF' AND "
   vgSql = vgSql & " NUM_PERPAGO = '" & vlNumPerPago & "' AND"
   vgSql = vgSql & " COD_CONHABDES = '08'"
   Set vlRegistro = vgConexionBD.Execute(vgSql)
   If Not vlRegistro.EOF Then
      While Not vlRegistro.EOF
            vlNumPoliza = (vlRegistro!Num_Poliza)
            vlNumOrden = (vlRegistro!Num_Orden)
            vlRutReceptor = (vlRegistro!Rut_Receptor)
            vgSql = ""
            vgSql = "SELECT MTO_CARGA FROM PP_TTMP_CONPAGOASIG WHERE"
            vgSql = vgSql & " COD_TIPOMOD = 'AF' AND "
            vgSql = vgSql & " NUM_PERPAGO = '" & vlNumPerPago & "' AND"
            vgSql = vgSql & " NUM_POLIZA = '" & vlNumPoliza & "' AND"
            vgSql = vgSql & " NUM_ORDEN = '" & vlNumOrden & "' AND"
            vgSql = vgSql & " RUT_RECEPTOR = '" & vlRutReceptor & "' AND"
            Set vlRegistro1 = vgConexionBD.Execute(vgSql)
            If Not vlRegistro1.EOF Then
               While Not vlRegistro1.EOF
                     vlNumOrdenCar = (vlRegistro1!Num_OrdenCar)
                     vlMontoCarga = (vlRegistro1!Mto_Carga)
                     vgSql = ""
                     vgSql = "SELECT COD_SITINV FEC_TERACTIVA FROM PP_TMAE_ASIGFAM WHERE "
                     vgSql = vgSql & "NUM_POLIZA = '" & vlNumPoliza & "' AND"
                     vgSql = vgSql & "NUM_ORDEN = '" & vlNumOrdenCar & "' ORDER BY FEC_TERACTIVA DESC"
                     Set vlregistro3 = vgConexionBD.Execute(vgSql)
                     If Not vlregistro3.EOF Then
                        vlCodSitInv = (vlregistro3!COD_SITINV)
                        vlFecVencimiento = (vlregistro3!FEC_TERACTIVA)
                     End If 'HQR 23/10/2004 lo agregué solo para que compile
                Wend 'HQR 23/10/2004 lo agregué solo para que compile
            End If 'HQR 23/10/2004 lo agregué solo para que compile
            
            'Insert
           Sql = ""
           Sql = "insert into PP_TTMP_CONASIGFAM ("
           Sql = Sql & "COD_USUARIO,NUM_POLIZA,NUM_ORDENREC,NUM_ORDEN,"
           Sql = Sql & "RUT_BEN,DGV_BEN,MTO_CARGA,COD_PAR,COD_SITINV,"
           Sql = Sql & "MTO_RETRO,MTO_REINTEGRO,FEC_VENCARGA,MTO_TOTAL"
           Sql = Sql & " "
           Sql = Sql & ") values ("
           Sql = Sql & "'" & (vgUsuario) & "',"
           Sql = Sql & "" & (vlRegistro!Num_Poliza) & ","
           Sql = Sql & "" & (vlRegistro!NUM_ORDENREC) & ","
           Sql = Sql & "'" & (vlRegistro!Num_Orden) & "',"
           Sql = Sql & "'" & (vlRegistro!RUT_BEN) & "',"
           Sql = Sql & "'" & (vlRegistro!Num_Poliza) & "',"
           Sql = Sql & "'" & (vlRegistro!Num_Poliza) & "',"
           Sql = Sql & "'" & (vlRegistro!Num_Poliza) & "',"
           Sql = Sql & "'" & (vlRegistro!Num_Poliza) & "',"
           Sql = Sql & "'" & (vlRegistro!Num_Poliza) & "',"
           Sql = Sql & "'" & (vlRegistro!Num_Poliza) & "',"
           Sql = Sql & "'" & (vlRegistro!Num_Poliza) & "',"
           Sql = Sql & "'" & (vlRegistro!Num_Poliza) & "'"
           Sql = Sql & ")"
           vgConexionBD.Execute (Sql)
           
      Wend
           
           
           
           
           
           
           
   End If
   
Exit Function
Err_Carga:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Private Sub Txt_Mes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Txt_Mes <> "" Then
       If Txt_Mes > 1 And Txt_Mes < 12 Then
          Txt_Mes = Trim(Format(Txt_Mes, "00"))
       Else
         Txt_Mes = ""
       End If
    End If
End If
End Sub

Function flLlenaTablaTemporal(iMes, iAño, iModulo) As Boolean
'Función que llena la Tabla de Control de Asignación Familiar

On Error GoTo Errores

Dim vlPerPago As String
Dim vlFecIniPag As Date
Dim vlFecTerPag As Date
Dim vlFecPago As String
Dim vlUF As Double

flLlenaTablaTemporal = False
If Not fgConexionBaseDatos(vgConexionTransac) Then
    MsgBox "Error en Conexion a la Base de Datos", vbCritical, Me.Caption
    Exit Function
End If
vgConexionTransac.BeginTrans

'0.- Elimina Registros Anteriores
If Not fgBorraCalculosControlAnteriores(iModulo, Me.Caption) Then
    Exit Function
End If

vlPerPago = iAño * 100 + iMes

'Obtiene Datos del Pago en Régimen
If Not fgObtieneDatosPagoRegimen(vlPerPago, stDatosInformeControl.Fec_PagoReg, stDatosInformeControl.Val_UFReg, vlFecIniPag, vlFecTerPag, Me.Caption) Then
    Exit Function
End If

'Obtiene Datos del Primer Pago
If Not fgObtieneDatosPrimerPago(vlPerPago, stDatosInformeControl.Fec_PriPago, stDatosInformeControl.Val_UFPri, vlFecIniPag, vlFecTerPag, Me.Caption) Then
    Exit Function
End If

'Llena Parte Invariable de la Estructura
stConLiquidacion.Num_PerPago = vlPerPago

'Cuenta Nº de Pólizas a Procesar

vlSQL = flObtieneQuery(iModulo, vlFecIniPag, vlFecTerPag, 2)

'vlSQL = "SELECT COUNT(1) AS CONTADOR FROM PP_TMAE_POLIZA A"
'vlSQL = vlSQL & " WHERE A.NUM_ENDOSO ="
'vlSQL = vlSQL & " (SELECT MAX(NUM_ENDOSO) FROM PP_TMAE_POLIZA " 'Último Endoso
'vlSQL = vlSQL & " WHERE NUM_POLIZA = A.NUM_POLIZA)"
'vlSQL = vlSQL & " AND COD_ESTADO IN (6, 7, 8) " 'solo Pólizas Vigentes, con Beneficiarios Designados y Pendiente (Pago Diferido)
''Se deben identificar las Pólizas que tengan su Primer Pago en este periodo (Inmediatas y Diferidas)
'vlSQL = vlSQL & " AND FEC_INIPAGOPEN BETWEEN '" & Format(vlFecIniPag, "yyyymmdd") & "'"
'vlSQL = vlSQL & " AND '" & Format(vgFecTerPag, "yyyymmdd") & "'"
Set vlTB = vgConexionBD.Execute(vlSQL)
If Not vlTB.EOF Then
    vlCont = vlTB!CONTADOR
    If vlCont = 0 Then
        MsgBox "No existen Pólizas con Primer Pago para este periodo", vbCritical, Me.Caption
        Exit Function
    End If
Else
    MsgBox "No existen Pólizas con Primer Pago para este periodo", vbCritical, Me.Caption
    Exit Function
End If

vlSQL = flObtieneQuery(iModulo, vlFecIniPag, vlFecTerPag, 1) 'Obtiene Query para filtrar las Pólizas

Set vlTB = vgConexionBD.Execute(vlSQL)
If Not vlTB.EOF Then
    vlAumento = (100 / vlCont)
    ProgressBar.Refresh
    stConTutor.Cod_GruFami = "-1"
    Do While Not vlTB.EOF
        'Registra Datos de la Poliza en la Estructura
        stConLiquidacion.Num_Poliza = vlTB!Num_Poliza
        stConLiquidacion.num_endoso = vlTB!num_endoso
        stConLiquidacion.Cod_TipPension = vlTB!Cod_TipPension
        If vlTB!TIPPAGO = "R" Then
            vlUF = stDatosInformeControl.Val_UFReg
            vlFecPago = stDatosInformeControl.Fec_PagoReg
        Else
            vlUF = stDatosInformeControl.Val_UFPri
            vlFecPago = stDatosInformeControl.Fec_PriPago
        End If
        stConLiquidacion.Cod_TipoPago = vlTB!TIPPAGO
        stConLiquidacion.Fec_Pago = vlFecPago
        'Obtiene Beneficiarios de la Póliza que tengan Derecho a Pensión
        vlSQL = "SELECT * FROM PP_TMAE_BEN"
        vlSQL = vlSQL & " WHERE NUM_POLIZA = '" & vlTB!Num_Poliza & "'"
        vlSQL = vlSQL & " AND NUM_ENDOSO = " & vlTB!num_endoso
        vlSQL = vlSQL & " AND NUM_ORDEN = " & vlTB!Num_Orden
        vlSQL = vlSQL & " AND COD_DERPEN = 99" 'Solo los que tienen Derecho a Pensión
        vlSQL = vlSQL & " AND FEC_INIPAGOPEN <= '" & Format(vlFecIniPag, "yyyymmdd") & "'" 'Solo los que ya iniciaron su pago de pensión o lo inician en este periodo
        vlSQL = vlSQL & " ORDER BY NUM_POLIZA,COD_GRUFAM, COD_PAR"
        Set vlTB2 = vgConexionBD.Execute(vlSQL)
        If Not vlTB2.EOF Then
            Do While Not vlTB2.EOF
                'Llena Datos de la Estructura que no cambian
                stConDetPension.Fec_IniPago = Format(vlFecIniPag, "yyyymmdd")
                stConDetPension.Fec_TerPago = Format(vlFecTerPag, "yyyymmdd")
                stConDetPension.num_endoso = vlTB!num_endoso
                stConDetPension.Num_Orden = vlTB2!Num_Orden
                stConDetPension.Num_Poliza = vlTB!Num_Poliza
                stConDetPension.Num_PerPago = vlPerPago
                stConDetPension.Edad = 0
                
                If vlTB2!COD_PAR < 30 Then 'Padres quedan registrados para Tutores
                    stConTutor.Cod_GruFami = vlTB2!COD_GRUFAM
                    stConTutor.Cod_TipReceptor = "M"
                    stConTutor.DGV_Receptor = vlTB2!DGV_BEN
                    stConTutor.Gls_MatReceptor = vlTB2!gls_matben
                    stConTutor.Gls_PatReceptor = vlTB2!gls_patben
                    stConTutor.Gls_NomReceptor = vlTB2!gls_nomben
                    stConTutor.Rut_Receptor = vlTB2!RUT_BEN
                    stConTutor.Gls_Direccion = vlTB2!GLS_DIRBEN
                    stConTutor.Cod_Direccion = vlTB2!Cod_Direccion
                    stConTutor.Cod_ViaPago = vlTB2!Cod_ViaPago
                    stConTutor.Cod_Banco = IIf(IsNull(vlTB2!Cod_Banco), "NULL", vlTB2!Cod_Banco)
                    stConTutor.Cod_TipCuenta = IIf(IsNull(vlTB2!Cod_TipCuenta), "NULL", vlTB2!Cod_TipCuenta)
                    stConTutor.Num_Cuenta = IIf(IsNull(vlTB2!Num_Cuenta), "NULL", vlTB2!Num_Cuenta)
                    stConTutor.Cod_Sucursal = IIf(IsNull(vlTB2!Cod_Sucursal), "NULL", vlTB2!Cod_Sucursal)
                End If
        
                'Calcula Edad de Todos los Beneficiarios
                bResp = fgCalculaEdad(vlTB2!FEC_NACBEN, vlFecIniPag)
                If bResp = "-1" Then 'Error
                    Exit Function
                End If
                stConDetPension.Edad = bResp
                stConDetPension.EdadAños = fgConvierteEdadAños(stConDetPension.Edad)
                'Si son Hijos se Calcula la Edad y se Verifica Certificado de Estudios
                If vlTB2!COD_PAR >= 30 And vlTB2!COD_PAR <= 35 Then 'Hijos
                    If stConDetPension.Edad >= stDatGenerales.MesesEdad18 And stConDetPension.Edad <= stDatGenerales.MesesEdad24 And vlTB2!COD_SITINV = "N" Then 'Hijos Sanos
                        'Verifica Certificados de Estudio
                        bResp = fgVerificaCertEstudios(stConLiquidacion.Num_Poliza, stConDetPension.Num_Orden)
                        If bResp = "-1" Then 'Error
                            Exit Function
                        Else
                            If bResp = "0" Then 'No tiene Certificado de Estudios
                                GoTo Siguiente 'Va al Siguiente Beneficiario, ya que éste no tiene Derecho
                            End If
                        End If
                    Else
                        If stConDetPension.Edad >= stDatGenerales.MesesEdad24 And vlTB2!COD_SITINV = "N" Then 'Hijo Mayor de 24 No Invalido
                            GoTo Siguiente
                        End If
                    End If
                End If
                
                'Inicializa Monto Haber y Descuento
                stConLiquidacion.Mto_Haber = 0
                stConLiquidacion.Mto_Descuento = 0
                stConLiquidacion.Num_Orden = vlTB2!Num_Orden
                
                '11.-  Obtener Tutores (1a. Etapa) (Se deja acá porque se necesita el Rut del Receptor)
                bResp = fgObtieneTutor(stConLiquidacion.Num_Poliza, stConLiquidacion.num_endoso, stConLiquidacion.Num_Orden, vlFecIniPag, vlFecTerPag, stConLiquidacion)
                If bResp = "-1" Then 'Error
                    Exit Function
                Else
                    If bResp = "0" Then 'No Encontró Tutor
                        If vlTB2!COD_PAR >= 30 And vlTB2!COD_PAR <= 35 And stConDetPension.Edad <= stDatGenerales.MesesEdad18 And stConTutor.Cod_GruFami = vlTB2!COD_GRUFAM Then 'El Tutor debe ser la Madre
                            stConLiquidacion.Cod_TipReceptor = stConTutor.Cod_TipReceptor 'MADRE
                            stConLiquidacion.Rut_Receptor = stConTutor.Rut_Receptor
                            stConLiquidacion.DGV_Receptor = stConTutor.DGV_Receptor
                            stConLiquidacion.Gls_NomReceptor = stConTutor.Gls_NomReceptor
                            stConLiquidacion.Gls_PatReceptor = stConTutor.Gls_PatReceptor
                            stConLiquidacion.Gls_MatReceptor = stConTutor.Gls_MatReceptor
                            stConLiquidacion.Gls_Direccion = stConTutor.Gls_Direccion
                            stConLiquidacion.Cod_Direccion = stConTutor.Cod_Direccion
                            stConLiquidacion.Cod_ViaPago = stConTutor.Cod_ViaPago
                            stConLiquidacion.Cod_Banco = stConTutor.Cod_Banco
                            stConLiquidacion.Cod_TipCuenta = stConTutor.Cod_TipCuenta
                            stConLiquidacion.Num_Cuenta = stConTutor.Num_Cuenta
                            stConLiquidacion.Cod_Sucursal = stConTutor.Cod_Sucursal
                            
                        Else 'Else se le Pagará a El Mismo
                            stConLiquidacion.Cod_TipReceptor = "P" 'Causante
                            stConLiquidacion.Rut_Receptor = vlTB2!RUT_BEN
                            stConLiquidacion.DGV_Receptor = vlTB2!DGV_BEN
                            stConLiquidacion.Gls_NomReceptor = vlTB2!gls_nomben
                            stConLiquidacion.Gls_PatReceptor = vlTB2!gls_patben
                            stConLiquidacion.Gls_MatReceptor = vlTB2!gls_matben
                            stConLiquidacion.Gls_Direccion = vlTB2!GLS_DIRBEN
                            stConLiquidacion.Cod_Direccion = vlTB2!Cod_Direccion
                            stConLiquidacion.Cod_ViaPago = vlTB2!Cod_ViaPago
                            stConLiquidacion.Cod_Banco = IIf(IsNull(vlTB2!Cod_Banco), "NULL", vlTB2!Cod_Banco)
                            stConLiquidacion.Cod_TipCuenta = IIf(IsNull(vlTB2!Cod_TipCuenta), "NULL", vlTB2!Cod_TipCuenta)
                            stConLiquidacion.Num_Cuenta = IIf(IsNull(vlTB2!Num_Cuenta), "NULL", vlTB2!Num_Cuenta)
                            stConLiquidacion.Cod_Sucursal = IIf(IsNull(vlTB2!Cod_Sucursal), "NULL", vlTB2!Cod_Sucursal)
                        End If
                    'Else 'Encontró Tutor
                    End If
                End If
                
                stConDetPension.Rut_Receptor = stConLiquidacion.Rut_Receptor
                
                If iModulo <> "AF" Then
                    '2.- Obtiene Monto de la Pensión
                    'Verifica si está en Periodo Garantizado
                        If Not IsNull(vlTB2!FEC_TERPAGOPENGAR) Then
                            If vlTB2!FEC_TERPAGOPENGAR >= Format(vlFecIniPag, "yyyymmdd") Then
                                vlPension = vlTB2!MTO_PENSIONGAR
                            Else
                                vlPension = vlTB2!MTO_PENSION
                            End If
                        Else
                            vlPension = vlTB2!MTO_PENSION
                        End If
                        vlPension = Format(vlPension * vlUF, "###,##0") 'Transforma a Pesos
                        
                    stConDetPension.Cod_ConHabDes = stDatGenerales.Cod_ConceptoPension
                    stConDetPension.Mto_ConHabDes = vlPension
                    stConLiquidacion.Mto_Haber = stConLiquidacion.Mto_Haber + vlPension
                    'Graba Monto de la Pensión
                    If Not fgInsertaDetallePensionProv(stConDetPension, "C", iModulo) Then
                        MsgBox "Se ha producido un Error al Grabar Monto de la Pensión" & Chr(13) & Err.Description, vbCritical, Me.Caption
                        Exit Function
                    End If
                    
                    '3.- Obtener Garantía Estatal (3a. Etapa)
                    vlGarantia = 0
                    If Not fgCalculaGarantiaEstatal(stConLiquidacion.Cod_TipPension, stConLiquidacion.Num_Poliza, stConLiquidacion.num_endoso, stConLiquidacion.Num_Orden, vlTB2!COD_PAR, vlTB2!COD_SITINV, vlTB2!cod_sexo, vlPension, stConDetPension.EdadAños, vlFecPago, vlGarantia) Then
                        Exit Function
                    End If
                    If vlGarantia > 0 Then
                        stConDetPension.Cod_ConHabDes = stDatGenerales.Cod_ConceptoGarantiaEstatal
                        stConDetPension.Mto_ConHabDes = vlGarantia
                        stConLiquidacion.Mto_Haber = stConLiquidacion.Mto_Haber + vlGarantia
                        'Graba Monto de la Pensión
                        If Not fgInsertaDetallePensionProv(stConDetPension, "C", iModulo) Then
                            MsgBox "Se ha producido un Error al Grabar Monto de la Garantia Estatal" & Chr(13) & Err.Description, vbCritical, Me.Caption
                            Exit Function
                        End If
                    End If
                End If
                
                '4.- Obtener Haberes y Descuentos Imponibles (1a. Etapa)
                If Not fgObtieneHaberesDescuentos(vlTB!Num_Poliza, vlTB!num_endoso, vlTB2!Num_Orden, "S", "S", vlMonto, 0, 0, vlFecPago, stConLiquidacion, stConDetPension, vlUF, vlFecIniPag, vlFecTerPag, "C", iModulo) Then
                    Exit Function
                End If
                
                If iModulo <> "AF" Then
                    vlBaseImp = vlPension + vlMonto 'Base Imponible
                    stConLiquidacion.Mto_BaseImp = vlBaseImp
                    
                    '5.- Calcular Descto. Salud (1a. Etapa)
                    vlDesSalud = 0
                    If Not fgObtienePrcSalud(vlTB2!Cod_InsSalud, vlTB2!COD_MODSALUD, vlTB2!MTO_PLANSALUD, vlBaseImp, vlDesSalud, vlUF, vlFecPago) Then
                        Exit Function
                    End If
                    stConDetPension.Cod_ConHabDes = stDatGenerales.Cod_ConceptoDesctoSalud
                    stConDetPension.Mto_ConHabDes = vlDesSalud
                    stConLiquidacion.Mto_Descuento = stConLiquidacion.Mto_Descuento + vlDesSalud
                    'Graba Monto del Descuento de Salud
                    If Not fgInsertaDetallePensionProv(stConDetPension, "C", iModulo) Then
                        MsgBox "Se ha producido un Error al Grabar Descuento de Salud" & Chr(13) & Err.Description, vbCritical, Me.Caption
                        Exit Function
                    End If
                End If
                
                '6.- Agregar Haberes y Descuentos No Imponibles y Tributables (1a. Etapa)
                If Not fgObtieneHaberesDescuentos(vlTB!Num_Poliza, vlTB!num_endoso, vlTB2!Num_Orden, "N", "S", vlMonto, vlBaseImp, 0, vlFecPago, stConLiquidacion, stConDetPension, vlUF, vlFecIniPag, vlFecTerPag, "C", iModulo) Then
                    Exit Function
                End If
                
                If iModulo <> "AF" Then
                    vlBaseTrib = (vlBaseImp - vlDesSalud) + vlMonto 'Base Imponible
                    stConLiquidacion.Mto_BaseTri = vlBaseTrib
                    
                    '7.- Calcular Impto. Único (1a. Etapa)
                    vlImpuesto = 0
                    If Not fgObtieneImpuestoUnico(vlBaseTrib, vlImpuesto, vlPerPago) Then
                        Exit Function
                    End If
                    vlBaseLiq = vlBaseTrib - vlImpuesto
                    stConDetPension.Cod_ConHabDes = stDatGenerales.Cod_ConceptoImpuesto
                    stConDetPension.Mto_ConHabDes = vlImpuesto
                    stConLiquidacion.Mto_Descuento = stConLiquidacion.Mto_Descuento + vlImpuesto
                
                    'Graba Monto del Impuesto Único
                    If Not fgInsertaDetallePensionProv(stConDetPension, "C", iModulo) Then
                        MsgBox "Se ha producido un Error al Grabar Descuento de Salud" & Chr(13) & Err.Description, vbCritical, Me.Caption
                        Exit Function
                    End If
                End If
                
                '8.- Calcular Asign. Familiar (2a. Etapa)
                
                'Calcula Valor de la Carga Familiar
                vlVarCarga = 0
                vlNumCargas = 0
                If Not fgCalculaValorCargaFamiliar(vlTB!Num_Poliza, vlTB!num_endoso, vlTB2!Num_Orden, vlFecIniPag, vlVarCarga) Then
                    Exit Function
                End If
                If vlVarCarga >= 0 Then
                    vlNumCargas = 0 'Número de Cargas Familiares
                    If Not fgCalculaNumCargasFamiliares(vlTB!Num_Poliza, vlTB!num_endoso, vlTB2!Num_Orden, -1, stConDetPension.Rut_Receptor, vlVarCarga, vlFecIniPag, vlPerPago, vlNumCargas, stConDetPension, "C", iModulo) Then
                        Exit Function
                    End If
                    If vlNumCargas > 0 Then
                        vlAsigFamiliar = Format(vlVarCarga * vlNumCargas, "##0")
                        stConDetPension.Cod_ConHabDes = stDatGenerales.Cod_ConceptoAsigFami
                        stConDetPension.Mto_ConHabDes = vlAsigFamiliar
                        stConLiquidacion.Mto_Haber = stConLiquidacion.Mto_Haber + vlAsigFamiliar
                        'Graba Monto de la Asignación Familiar
                        If Not fgInsertaDetallePensionProv(stConDetPension, "C", iModulo) Then
                            MsgBox "Se ha producido un Error al Grabar la Asignación Familiar" & Chr(13) & Err.Description, vbCritical, Me.Caption
                            Exit Function
                        End If
                    End If
                End If
                stConLiquidacion.Num_Cargas = vlNumCargas
                
                If iModulo <> "AF" Then
                    '9.- Calcular Retencion Judicial (2a. Etapa)
                    If Not fgCalculaRetencion(vlTB!Num_Poliza, vlTB!num_endoso, vlTB2!Num_Orden, vlPerPago, vlFecPago, vlBaseImp, vlBaseTrib, vlVarCarga, stConLiquidacion, stConDetPension, vlFecIniPag, Me.Caption, "C", iModulo) Then
                        Exit Function
                    End If
                End If

                '10.- Agregar Haberes y Descuentos No Imponibles y No Tributables (1a. Etapa)
                If Not fgObtieneHaberesDescuentos(vlTB!Num_Poliza, vlTB!num_endoso, vlTB2!Num_Orden, "N", "N", vlMonto, vlBaseImp, vlBaseTrib, vlFecPago, stConLiquidacion, stConDetPension, vlUF, vlFecIniPag, vlFecTerPag, "C", iModulo) Then
                    Exit Function
                End If
                
                '12.- Obtener Mensajes (1a. Etapa)
                
                '13.-  Generar Liquidación (1a. Etapa)
                stConLiquidacion.Cod_CajaCompen = IIf(IsNull(vlTB2!Cod_CajaCompen), "NULL", vlTB2!Cod_CajaCompen)
                stConLiquidacion.Cod_InsSalud = vlTB2!Cod_InsSalud
                stConLiquidacion.Mto_LiqPagar = stConLiquidacion.Mto_Haber - stConLiquidacion.Mto_Descuento
                
                If stConLiquidacion.Mto_Haber > 0 Or stConLiquidacion.Mto_Descuento > 0 Then
                    'Inserta Liquidacion
                    If Not fgInsertaLiquidacion(stConLiquidacion, "C", iModulo) Then
                        Exit Function
                    End If
                End If
Siguiente:
                vlTB2.MoveNext
            Loop
        End If
        'Refresca Barra de Progreso
        If (ProgressBar.Value + vlAumento) <= 100 Then
            ProgressBar.Value = (ProgressBar.Value + vlAumento)
        End If
        ProgressBar.Refresh
        vlTB.MoveNext
    
    Loop
Else
    MsgBox "No existen Pólizas con Pago para este Periodo", vbCritical, Me.Caption
    Exit Function
End If

'''Traspasa Datos a Histórico
''If Me.Tag = "D" Then
''    If Not flTraspasaDatosADefinitivos Then
''        Exit Function
''    End If
''End If
vgConexionTransac.CommitTrans
vgConexionTransac.Close
        
ProgressBar.Value = 0
flLlenaTablaTemporal = True
Errores:
    Screen.MousePointer = 0
    If Err.Number <> 0 Then
        vgConexionTransac.RollbackTrans
        vgConexionTransac.Close
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End If


End Function
Function flObtieneQuery(iModulo, iFecIniPag, iFecTerPag, iModalidad) As String
'Función que obtiene la Query que obtendrá las Pólizas a Evaluar
'Devuelve el String que obtiene los Datos
'iModalidad: 1 => Query, 2 => Contador
Dim vlSQL As String

If iModulo = "AF" Then 'Asignación Familiar
    vlSQL = ""
    If iModalidad = 2 Then 'REVISAR PARA SQL
        vlSQL = "SELECT COUNT(1) AS CONTADOR FROM ("
    End If
    vlSQL = vlSQL & " SELECT A.NUM_POLIZA, A.NUM_ENDOSO, B.NUM_ORDENREC AS NUM_ORDEN, A.COD_TIPPENSION, A.NUM_MESGAR, 'P' AS TIPPAGO "
    vlSQL = vlSQL & " FROM PP_TMAE_POLIZA A, PP_TMAE_ASIGFAM B"
    vlSQL = vlSQL & " WHERE A.NUM_ENDOSO ="
    vlSQL = vlSQL & " (SELECT MAX(NUM_ENDOSO) FROM PP_TMAE_POLIZA " 'Último Endoso
    vlSQL = vlSQL & " WHERE NUM_POLIZA = A.NUM_POLIZA)"
    vlSQL = vlSQL & " AND A.COD_ESTADO IN (6, 7, 8) " 'solo Pólizas Vigentes, con Beneficiarios Designados y Pendiente (Pago Diferido)
    
    'Se deben identificar las Pólizas que tengan su Primer Pago en este periodo (Inmediatas y Diferidas)
    vlSQL = vlSQL & " AND A.FEC_INIPAGOPEN BETWEEN '" & Format(iFecIniPag, "yyyymmdd") & "'"
    vlSQL = vlSQL & " AND '" & Format(iFecTerPag, "yyyymmdd") & "'"
    
    vlSQL = vlSQL & " AND A.NUM_POLIZA = B.NUM_POLIZA"
    vlSQL = vlSQL & " AND '" & Format(iFecIniPag, "yyyymmdd") & "'"
    vlSQL = vlSQL & " BETWEEN B.FEC_INIACTIVA AND B.FEC_TERACTIVA"
    vlSQL = vlSQL & " AND B.COD_ESTVIGENCIA = 'A'" 'Activa
    vlSQL = vlSQL & " UNION" 'HABERES Y DESCUENTOS CON PRIMER PAGO
    vlSQL = vlSQL & " SELECT A.NUM_POLIZA, A.NUM_ENDOSO, B.NUM_ORDEN, A.COD_TIPPENSION, A.NUM_MESGAR, 'R' AS TIPPAGO "
    vlSQL = vlSQL & " FROM PP_TMAE_POLIZA A, PP_TMAE_HABDES B, MA_TPAR_CONHABDES C"
    vlSQL = vlSQL & " WHERE A.NUM_ENDOSO ="
    vlSQL = vlSQL & " (SELECT MAX(NUM_ENDOSO) FROM PP_TMAE_POLIZA " 'Último Endoso
    vlSQL = vlSQL & " WHERE NUM_POLIZA = A.NUM_POLIZA)"
    vlSQL = vlSQL & " AND COD_ESTADO IN (6, 7, 8) " 'solo Pólizas Vigentes, con Beneficiarios Designados y Pendiente (Pago Diferido)
    vlSQL = vlSQL & " AND A.FEC_INIPAGOPEN BETWEEN '" & Format(iFecIniPag, "yyyymmdd") & "'"
    vlSQL = vlSQL & " AND '" & Format(iFecTerPag, "yyyymmdd") & "'"
    vlSQL = vlSQL & " AND A.NUM_POLIZA = B.NUM_POLIZA"
    vlSQL = vlSQL & " AND '" & Format(iFecIniPag, "yyyymmdd") & "'"
    vlSQL = vlSQL & " BETWEEN B.FEC_INIHABDES AND B.FEC_TERHABDES"
    vlSQL = vlSQL & " AND B.COD_CONHABDES = C.COD_CONHABDES"
    vlSQL = vlSQL & " AND C.COD_MODORIGEN = '" & iModulo & "'"
    
    
    'UNION los Pagos en Régimen
    vlSQL = vlSQL & " UNION "
    vlSQL = vlSQL & "SELECT A.NUM_POLIZA, A.NUM_ENDOSO, B.NUM_ORDENREC AS NUM_ORDEN, A.COD_TIPPENSION, A.NUM_MESGAR, 'R' AS TIPPAGO "
    vlSQL = vlSQL & " FROM PP_TMAE_POLIZA A, PP_TMAE_ASIGFAM B"
    vlSQL = vlSQL & " WHERE A.NUM_ENDOSO ="
    vlSQL = vlSQL & " (SELECT MAX(NUM_ENDOSO) FROM PP_TMAE_POLIZA " 'Último Endoso
    vlSQL = vlSQL & " WHERE NUM_POLIZA = A.NUM_POLIZA)"
    vlSQL = vlSQL & " AND COD_ESTADO IN (6, 7, 8) " 'solo Pólizas Vigentes, con Beneficiarios Designados y Pendiente (Pago Diferido)
    'Se deben identificar las Pólizas que tengan su Primer Pago en un Periodo Anterior (Inmediatas y Diferidas)
    vlSQL = vlSQL & " AND A.FEC_INIPAGOPEN < '" & Format(iFecIniPag, "yyyymmdd") & "'"
    
    vlSQL = vlSQL & " AND A.NUM_POLIZA = B.NUM_POLIZA"
    vlSQL = vlSQL & " AND '" & Format(iFecIniPag, "yyyymmdd") & "'"
    vlSQL = vlSQL & " BETWEEN B.FEC_INIACTIVA AND B.FEC_TERACTIVA"
    vlSQL = vlSQL & " AND B.COD_ESTVIGENCIA = 'A'" 'Activa
    
    vlSQL = vlSQL & " UNION" 'HABERES Y DESCUENTOS
    vlSQL = vlSQL & " SELECT A.NUM_POLIZA, A.NUM_ENDOSO, B.NUM_ORDEN, A.COD_TIPPENSION, A.NUM_MESGAR, 'R' AS TIPPAGO "
    vlSQL = vlSQL & " FROM PP_TMAE_POLIZA A, PP_TMAE_HABDES B, MA_TPAR_CONHABDES C"
    vlSQL = vlSQL & " WHERE A.NUM_ENDOSO ="
    vlSQL = vlSQL & " (SELECT MAX(NUM_ENDOSO) FROM PP_TMAE_POLIZA " 'Último Endoso
    vlSQL = vlSQL & " WHERE NUM_POLIZA = A.NUM_POLIZA)"
    vlSQL = vlSQL & " AND COD_ESTADO IN (6, 7, 8) " 'solo Pólizas Vigentes, con Beneficiarios Designados y Pendiente (Pago Diferido)
    'Se deben identificar las Pólizas que tengan su Primer Pago en un Periodo Anterior (Inmediatas y Diferidas)
    vlSQL = vlSQL & " AND A.FEC_INIPAGOPEN < '" & Format(iFecIniPag, "yyyymmdd") & "'"
    
    vlSQL = vlSQL & " AND A.NUM_POLIZA = B.NUM_POLIZA"
    vlSQL = vlSQL & " AND '" & Format(iFecIniPag, "yyyymmdd") & "'"
    vlSQL = vlSQL & " BETWEEN B.FEC_INIHABDES AND B.FEC_TERHABDES"
    vlSQL = vlSQL & " AND B.COD_CONHABDES = C.COD_CONHABDES"
    vlSQL = vlSQL & " AND C.COD_MODORIGEN = '" & iModulo & "'"
    If iModalidad = 2 Then
        vlSQL = vlSQL & ")"
    Else
        If vgTipoBase = "ORACLE" Then
            vlSQL = vlSQL & " ORDER BY NUM_POLIZA, NUM_ENDOSO"
        Else
            vlSQL = vlSQL & " ORDER BY A.NUM_POLIZA, A.NUM_ENDOSO"
        End If
    End If
ElseIf iModulo = "CCAF" Then 'Caja de Compensación
    vlSQL = ""
    If iModalidad = 2 Then
        vlSQL = "SELECT COUNT(1) AS CONTADOR FROM ("
    End If
    'HABERES Y DESCUENTOS CON PRIMER PAGO
    vlSQL = vlSQL & " SELECT A.NUM_POLIZA, A.NUM_ENDOSO, B.NUM_ORDEN, A.COD_TIPPENSION, A.NUM_MESGAR, 'R' AS TIPPAGO "
    vlSQL = vlSQL & " FROM PP_TMAE_POLIZA A, PP_TMAE_HABDES B, MA_TPAR_CONHABDES C"
    vlSQL = vlSQL & " WHERE A.NUM_ENDOSO ="
    vlSQL = vlSQL & " (SELECT MAX(NUM_ENDOSO) FROM PP_TMAE_POLIZA " 'Último Endoso
    vlSQL = vlSQL & " WHERE NUM_POLIZA = A.NUM_POLIZA)"
    vlSQL = vlSQL & " AND COD_ESTADO IN (6, 7, 8) " 'solo Pólizas Vigentes, con Beneficiarios Designados y Pendiente (Pago Diferido)
    vlSQL = vlSQL & " AND A.FEC_INIPAGOPEN BETWEEN '" & Format(iFecIniPag, "yyyymmdd") & "'"
    vlSQL = vlSQL & " AND '" & Format(iFecTerPag, "yyyymmdd") & "'"
    vlSQL = vlSQL & " AND A.NUM_POLIZA = B.NUM_POLIZA"
    vlSQL = vlSQL & " AND '" & Format(iFecIniPag, "yyyymmdd") & "'"
    vlSQL = vlSQL & " BETWEEN B.FEC_INIHABDES AND B.FEC_TERHABDES"
    vlSQL = vlSQL & " AND B.COD_CONHABDES = C.COD_CONHABDES"
    vlSQL = vlSQL & " AND C.COD_MODORIGEN = '" & iModulo & "'"
    vlSQL = vlSQL & " UNION"
    vlSQL = vlSQL & " SELECT A.NUM_POLIZA, A.NUM_ENDOSO, B.NUM_ORDEN, A.COD_TIPPENSION, A.NUM_MESGAR, 'R' AS TIPPAGO "
    vlSQL = vlSQL & " FROM PP_TMAE_POLIZA A, PP_TMAE_HABDES B, MA_TPAR_CONHABDES C"
    vlSQL = vlSQL & " WHERE A.NUM_ENDOSO ="
    vlSQL = vlSQL & " (SELECT MAX(NUM_ENDOSO) FROM PP_TMAE_POLIZA " 'Último Endoso
    vlSQL = vlSQL & " WHERE NUM_POLIZA = A.NUM_POLIZA)"
    vlSQL = vlSQL & " AND COD_ESTADO IN (6, 7, 8) " 'solo Pólizas Vigentes, con Beneficiarios Designados y Pendiente (Pago Diferido)
    'Se deben identificar las Pólizas que tengan su Primer Pago en un Periodo Anterior (Inmediatas y Diferidas)
    vlSQL = vlSQL & " AND A.FEC_INIPAGOPEN < '" & Format(iFecIniPag, "yyyymmdd") & "'"
    
    vlSQL = vlSQL & " AND A.NUM_POLIZA = B.NUM_POLIZA"
    vlSQL = vlSQL & " AND '" & Format(iFecIniPag, "yyyymmdd") & "'"
    vlSQL = vlSQL & " BETWEEN B.FEC_INIHABDES AND B.FEC_TERHABDES"
    vlSQL = vlSQL & " AND B.COD_CONHABDES = C.COD_CONHABDES"
    vlSQL = vlSQL & " AND C.COD_MODORIGEN = '" & iModulo & "'"

    If iModalidad = 2 Then
        vlSQL = vlSQL & ")"
    Else
        If vgTipoBase = "ORACLE" Then
            vlSQL = vlSQL & " ORDER BY NUM_POLIZA, NUM_ENDOSO"
        Else
            vlSQL = vlSQL & " ORDER BY A.NUM_POLIZA, A.NUM_ENDOSO"
        End If
    End If
ElseIf iModulo = "RJ" Then 'Retención Judicial
    vlSQL = ""
    If iModalidad = 2 Then
        vlSQL = "SELECT COUNT(1) AS CONTADOR FROM ("
    End If
    vlSQL = vlSQL & " SELECT A.NUM_POLIZA, A.NUM_ENDOSO, B.NUM_ORDEN, A.COD_TIPPENSION, A.NUM_MESGAR, 'P' AS TIPPAGO "
    vlSQL = vlSQL & " FROM PP_TMAE_POLIZA A, PP_TMAE_RETJUDICIAL B"
    vlSQL = vlSQL & " WHERE A.NUM_ENDOSO ="
    vlSQL = vlSQL & " (SELECT MAX(NUM_ENDOSO) FROM PP_TMAE_POLIZA " 'Último Endoso
    vlSQL = vlSQL & " WHERE NUM_POLIZA = A.NUM_POLIZA)"
    vlSQL = vlSQL & " AND A.COD_ESTADO IN (6, 7, 8) " 'solo Pólizas Vigentes, con Beneficiarios Designados y Pendiente (Pago Diferido)
    
    'Se deben identificar las Pólizas que tengan su Primer Pago en este periodo (Inmediatas y Diferidas)
    vlSQL = vlSQL & " AND A.FEC_INIPAGOPEN BETWEEN '" & Format(iFecIniPag, "yyyymmdd") & "'"
    vlSQL = vlSQL & " AND '" & Format(iFecTerPag, "yyyymmdd") & "'"
    
    vlSQL = vlSQL & " AND A.NUM_POLIZA = B.NUM_POLIZA"
    vlSQL = vlSQL & " AND '" & Format(iFecIniPag, "yyyymmdd") & "'"
    vlSQL = vlSQL & " BETWEEN B.FEC_INIRET AND B.FEC_TERRET"
    vlSQL = vlSQL & " UNION" 'HABERES Y DESCUENTOS CON PRIMER PAGO
    vlSQL = vlSQL & " SELECT A.NUM_POLIZA, A.NUM_ENDOSO, B.NUM_ORDEN, A.COD_TIPPENSION, A.NUM_MESGAR, 'R' AS TIPPAGO "
    vlSQL = vlSQL & " FROM PP_TMAE_POLIZA A, PP_TMAE_HABDES B, MA_TPAR_CONHABDES C"
    vlSQL = vlSQL & " WHERE A.NUM_ENDOSO ="
    vlSQL = vlSQL & " (SELECT MAX(NUM_ENDOSO) FROM PP_TMAE_POLIZA " 'Último Endoso
    vlSQL = vlSQL & " WHERE NUM_POLIZA = A.NUM_POLIZA)"
    vlSQL = vlSQL & " AND COD_ESTADO IN (6, 7, 8) " 'solo Pólizas Vigentes, con Beneficiarios Designados y Pendiente (Pago Diferido)
    vlSQL = vlSQL & " AND A.FEC_INIPAGOPEN BETWEEN '" & Format(iFecIniPag, "yyyymmdd") & "'"
    vlSQL = vlSQL & " AND '" & Format(iFecTerPag, "yyyymmdd") & "'"
    vlSQL = vlSQL & " AND A.NUM_POLIZA = B.NUM_POLIZA"
    vlSQL = vlSQL & " AND '" & Format(iFecIniPag, "yyyymmdd") & "'"
    vlSQL = vlSQL & " BETWEEN B.FEC_INIHABDES AND B.FEC_TERHABDES"
    vlSQL = vlSQL & " AND B.COD_CONHABDES = C.COD_CONHABDES"
    vlSQL = vlSQL & " AND C.COD_MODORIGEN = '" & iModulo & "'"

    'UNION los Pagos en Régimen
    vlSQL = vlSQL & " UNION "
    vlSQL = vlSQL & "SELECT A.NUM_POLIZA, A.NUM_ENDOSO, B.NUM_ORDEN, A.COD_TIPPENSION, A.NUM_MESGAR, 'R' AS TIPPAGO "
    vlSQL = vlSQL & " FROM PP_TMAE_POLIZA A, PP_TMAE_RETJUDICIAL B"
    vlSQL = vlSQL & " WHERE A.NUM_ENDOSO ="
    vlSQL = vlSQL & " (SELECT MAX(NUM_ENDOSO) FROM PP_TMAE_POLIZA " 'Último Endoso
    vlSQL = vlSQL & " WHERE NUM_POLIZA = A.NUM_POLIZA)"
    vlSQL = vlSQL & " AND COD_ESTADO IN (6, 7, 8) " 'solo Pólizas Vigentes, con Beneficiarios Designados y Pendiente (Pago Diferido)
    'Se deben identificar las Pólizas que tengan su Primer Pago en un Periodo Anterior (Inmediatas y Diferidas)
    vlSQL = vlSQL & " AND FEC_INIPAGOPEN < '" & Format(iFecIniPag, "yyyymmdd") & "'"
    
    vlSQL = vlSQL & " AND A.NUM_POLIZA = B.NUM_POLIZA"
    vlSQL = vlSQL & " AND '" & Format(iFecIniPag, "yyyymmdd") & "'"
    vlSQL = vlSQL & " BETWEEN B.FEC_INIRET AND B.FEC_TERRET"
    
    vlSQL = vlSQL & " UNION" 'HABERES Y DESCUENTOS
    vlSQL = vlSQL & " SELECT A.NUM_POLIZA, A.NUM_ENDOSO, B.NUM_ORDEN, A.COD_TIPPENSION, A.NUM_MESGAR, 'R' AS TIPPAGO "
    vlSQL = vlSQL & " FROM PP_TMAE_POLIZA A, PP_TMAE_HABDES B, MA_TPAR_CONHABDES C"
    vlSQL = vlSQL & " WHERE A.NUM_ENDOSO ="
    vlSQL = vlSQL & " (SELECT MAX(NUM_ENDOSO) FROM PP_TMAE_POLIZA " 'Último Endoso
    vlSQL = vlSQL & " WHERE NUM_POLIZA = A.NUM_POLIZA)"
    vlSQL = vlSQL & " AND COD_ESTADO IN (6, 7, 8) " 'solo Pólizas Vigentes, con Beneficiarios Designados y Pendiente (Pago Diferido)
    'Se deben identificar las Pólizas que tengan su Primer Pago en un Periodo Anterior (Inmediatas y Diferidas)
    vlSQL = vlSQL & " AND A.FEC_INIPAGOPEN < '" & Format(iFecIniPag, "yyyymmdd") & "'"
    
    vlSQL = vlSQL & " AND A.NUM_POLIZA = B.NUM_POLIZA"
    vlSQL = vlSQL & " AND '" & Format(iFecIniPag, "yyyymmdd") & "'"
    vlSQL = vlSQL & " BETWEEN B.FEC_INIHABDES AND B.FEC_TERHABDES"
    vlSQL = vlSQL & " AND B.COD_CONHABDES = C.COD_CONHABDES"
    vlSQL = vlSQL & " AND C.COD_MODORIGEN = '" & iModulo & "'"
    
    If iModalidad = 2 Then
        vlSQL = vlSQL & ")"
    Else
        If vgTipoBase = "ORACLE" Then
            vlSQL = vlSQL & " ORDER BY NUM_POLIZA, NUM_ENDOSO"
        Else
            vlSQL = vlSQL & " ORDER BY A.NUM_POLIZA, A.NUM_ENDOSO"
        End If
    End If

End If

flObtieneQuery = vlSQL
End Function

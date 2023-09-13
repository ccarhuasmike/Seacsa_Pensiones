Attribute VB_Name = "Mod_Calculo"
Option Explicit

'CMV-20060907 I
'Se agrega constante para código causal de suspensión de estado de G.E.
'Con valor 0 = Sin Información (Para guardar valor en Calculo G.E.)
Const vgCodCauSusEstGarEst0 As String * 1 = "0"
'CMV-20060907 F

'Datos Generales del Pago en Curso
Global vgFecIniPag As Date 'Fecha de Inicio del Pago de Pensiones (Primer Día del Mes)
Global vgFecTerPag As Date 'Fecha de Término del Pago de Pensiones (Primer Día del Mes)
Global vgPerPago As String 'Periodo de Pago
Global vgFecPago As String 'Fecha de Pago

'Estructura de Beneficiarios
Type TyBeneficiariosPension
    num_poliza As String
    num_endoso As String
    Num_Orden As Long
    Cod_Par As String
    Cod_GruFam As String
    Cod_TipoIdenBen As Long
    Gls_MatBen As String
    Gls_PatBen As String
    Gls_NomBen As String
    Gls_NomSegBen As String
    Num_IdenBen As String
    Gls_DirBen As String
    Cod_Direccion As String
    Cod_ViaPago As String
    Cod_Banco As String
    Cod_TipCuenta As String
    Num_Cuenta As String
    Cod_Sucursal As String
    Fec_NacBen As String
    Cod_SitInv As String
    Fec_TerPagoPenGar As String
    Prc_PensionGar As Double
    Prc_Pension As Double
    Cod_InsSalud As String
    Cod_ModSalud As String
    Mto_PlanSalud As Double
    Mto_Pension As Double
End Type

'Estructura del Pago de Pensiones
Type TyDetPension
    Num_PerPago As String
    num_poliza As String
    num_endoso As String
    Num_Orden As String
    Cod_ConHabDes As String
    Fec_IniPago As String
    Fec_TerPago As String
    Mto_ConHabDes As Double
    Edad As Integer
    EdadAños As Integer
    Num_IdenReceptor As String
    Cod_TipoIdenReceptor As Long
    Cod_Modulo As String
    Cod_TipReceptor As String
End Type

Type TyDatosGenerales
    Prc_SaludMin As Double
    Mto_MaxSalud As Double
    Mto_TopeBaseImponible As Double
    Val_UF As Double
    Val_UFUltDiaMes As Double
    MesesEdad18 As Long
    MesesEdad24 As Long
    Cod_ConceptoPension As String
    Cod_ConceptoDesctoSalud As String
    Cod_ConceptoGratificacion As String
    
    'Cod_ConceptoImpuesto As String
    'Cod_ConceptoAsigFami As String
    Cod_ConceptoRetencion As String
    'Cod_ConceptoRetAsig As String
    'Cod_ConceptoRetJudAsig As String
    Cod_ConceptoCobroRetencion As String
    'Cod_ConceptoGarantiaEstatal As String
    'Cod_ConceptoCobroAsigRet As String
    'Cod_ConceptoCobroRetJudAsig As String
    'Cod_ConceptoCajaCompensacion As String
    'Cod_ConceptoAsigFamPago As String 'Concepto de Pago de A.F. al Pensionado
    'Cod_ConceptoAsigFamCobro As String 'Concepto de Cobro de A.F. al Pensionado
    Cod_PensionAnticipada As String
    Cod_ConceptoPensionPago As String
    Cod_ConceptoPensionPagoRetro As String
    Cod_ConceptoPensionCobro As String
    
    'Cod_ConceptoGEPago As String
    'Cod_ConceptoGeCobro As String
    MesesEdad28 As Long
    MesesEdad65 As Long
    MesesEdad60 As Long
    
    'Cod_ConceptoDesctoSalud2 As String 'Concepto de Descuento por Salud
    
    Prc_Castigo As Double 'Porcentaje de Castigo para Garantía en Quiebra
    Mto_TopMaxQuiebra As Double 'Tope Máximo que se puede Pagar por Quiebra
    Ind_AplicaQuiebra As Boolean 'Indica si se debe aplicar quiebra o no
    Cod_ConceptoPensionQuiebra As String 'Concepto de Pensión en Quiebra
End Type

Type TyLiquidacion
    Num_PerPago As String
    num_poliza As String
    num_endoso As String
    Num_Orden As String
    Fec_Pago As String
    Gls_Direccion As String
    Cod_Direccion As String
    Cod_TipPension As String
    Cod_ViaPago As String
    Cod_Banco As String
    Cod_TipCuenta As String
    Num_Cuenta As String
    Cod_Sucursal As String
    Cod_InsSalud As String
    'Cod_CajaCompen As String
    Num_IdenReceptor As String
    Cod_TipoIdenReceptor As Long
    Gls_NomReceptor As String
    Gls_NomSegReceptor As String
    Gls_PatReceptor As String
    Gls_MatReceptor As String
    Cod_TipReceptor As String
    'Num_Cargas As Long
    Mto_Haber As Double
    Mto_Descuento As Double
    Mto_LiqPagar As Double
    Cod_TipoPago As String
    Mto_BaseImp As Double
    Mto_BaseTri As Double
    Cod_Modulo As String
    Mto_Pension As Double
    Cod_ModSalud As String
    Mto_PlanSalud As Double
    Cod_Moneda As String
    gls_vejez As String 'RRR 02/11/2012
End Type

Type TyTutor
    Num_IdenReceptor As String
    Cod_TipoIdenReceptor As Long
    Gls_NomReceptor As String
    Gls_NomSegReceptor As String
    Gls_PatReceptor As String
    Gls_MatReceptor As String
    Cod_TipReceptor As String
    Cod_GruFami As String
    Gls_Direccion As String
    Cod_Direccion As String
    Cod_TipPension As String
    Cod_ViaPago As String
    Cod_Banco As String
    Cod_TipCuenta As String
    Num_Cuenta As String
    Cod_Sucursal As String
End Type

Type TyTTMPLiquidacion
    Cod_Usuario  As String
    Num_PerPago As String
    num_poliza As String
    num_endoso As Integer
    Num_Orden As Integer
    Num_Item As Integer
    Cod_ConHaber As String
    Mto_Haber As Double
    Cod_ConDescto As String
    Mto_Descuento As Double
    Num_IdenReceptor As String
    Cod_TipoIdenReceptor As Long
    Cod_Modulo As String
    Cod_TipReceptor As String
    Gls_Direccion As String
    Fec_Pago As String
    Gls_NomReceptor As String
'    Gls_PatReceptor As String
'    Gls_MatReceptor As String
    Mto_LiqPagar As Double
    Fec_TerPodNot As String
    Fec_PagoProxReg As String
    Gls_TipPension As String
    Gls_ViaPago As String
    Gls_CajaComp As String
    Gls_InsSalud As String
    Mto_Moneda As Double
    Cod_Direccion As Double
'    DGV_Receptor As String
    Mto_Pension As Double
    Num_Cargas As Integer
    Mto_LiqHaber As Double
    Mto_LiqDescuento As Double
    Num_IdenBen As String
    Cod_TipoIdenBen As Long
    Gls_NomBen As String
    Gls_Direccion2 As String
    Gls_Mensaje As String
    Cod_Moneda As String
    Gls_MontoPension As String
    Gls_Afp As String
    gls_vejez As String 'RRR 02/11/2012
    Num_IdenTit As String 'RRR 22/07/2012
    Cod_TipoIdenTit As Long 'RRR 22/07/2012
    gls_nomTit As String 'RRR 22/07/2012
End Type

'Estructura para Datos de Informes de Control
Type TyDatosInformeControl
    Fec_PagoReg As String 'Fecha de Pago en Régimen
    Fec_PriPago As String 'Fecha del Primer Pago
    Val_UFReg As Double 'UF del Pago en Régimen
    Val_UFPri As Double 'UF del Primer Pago
    Val_UFRegUltDia As Double 'UF del Último Día del Mes de Pago en Régimen
    Val_UFPriUltDia As Double 'UF del Último Día del Mes de Primer Pago
    Prc_SaludMinimoPri As Double 'Porcentaje Mínimo de Cobro de Salud para Primeros Pagos
    Prc_SaludMinimoReg As Double 'Porcentaje Mínimo de Cobro de Salud para Pagos en Régimen
    Mto_MaximoSaludPri As Double 'Monto Máximo a Descontar por Salud para Primeros Pagos
    Mto_MaximoSaludReg As Double 'Monto Máximo a Descontar por Salud para Pagos en Régimen
    Mto_TopeBaseImponPri As Double 'Tope Base Imponible para Primeros Pagos
    Mto_TopeBaseImponReg As Double 'Tope Base Imponible para Pagos en Régimen
End Type

'Estructura para el Control de la Grilla de Pagos de la Reliquidacion
Type stPagos
    Ind_Acumulado As Integer 'Indica si la Fila ya se utilizó para calcular los Haberes y Descuentos
    monto As Double 'Monto de la Grilla
    Moneda As String 'Moneda del Concepto
    Concepto As String 'Concepto
    Num_Orden As Long 'Número de Orden del Beneficiario
End Type

Global stDetPension As TyDetPension
Global stDetPensionCCAF As TyDetPension
Global stDatGenerales As TyDatosGenerales
Global stLiquidacion As TyLiquidacion
Global stTutor As TyTutor
Global stTTMPLiquidacion As TyTTMPLiquidacion
Global vgConexionTransac As ADODB.Connection
Global vgTipoPago As String '"R:Regimen, P:Primeros Pagos
Global stBeneficiarios() As TyBeneficiariosPension
'marco
Global Const vgMesesExtension As Integer = -1 'Hasta 3 meses de gracia para certificado de supervivencia
Global Const clAjusteDesdeFechaDevengamiento As Boolean = True 'hqr 03/03/2011 Indica si se debe ajustar desde la fecha de devengamiento, TRUE: Se ajusta desde el devengamiento, FALSE: Se ajusta desde el inicio de vigencia de la Póliza

Public Function fgCalculaValorCargaFamiliar(iPoliza As String, iEndoso As Long, iOrden As Long, iFecCalculo As Date, oValAsignacion As Double, Optional iModulo As String) As Boolean
    'Función Encargada de Obtener el Valor de la Carga Familiar
    'Parametros de Entrada:
    ' iPoliza           => Nº de Poliza
    ' iEndoso           => Nº de Endoso (no se utiliza)
    ' iOrden            => Nº de Orden de quien recibe la Asignación Familiar (Beneficiario de Asignación Familiar)
    ' iFecCalculo       => Fecha a la cual se obtiene el Valor de la Carga
    'Parametros de Salida:
    ' oValAsignacion    => Valor de la Carga Familiar
    On Error GoTo Errores
    Dim vlSql As String
    Dim vlTB As ADODB.Recordset
    
    fgCalculaValorCargaFamiliar = False
    
    oValAsignacion = 0
    'Obtiene Valor de la Carga Familiar
    vlSql = "SELECT v.mto_carga FROM pp_tmae_valcarfam v"
    vlSql = vlSql & " WHERE v.num_poliza = '" & iPoliza & "'"
    'vlSQL = vlSQL & " AND v.num_endoso = " & iEndoso
    vlSql = vlSql & " AND v.num_orden = " & iOrden
    'vlSQL = vlSQL & " AND v.num_annodecing"
    If Not IsMissing(iModulo) Then
        If iModulo = "AFREL" Then 'Si es reliquidación
            vlSql = vlSql & " AND v.fec_inidecing <='" & Format(iFecCalculo, "yyyymmdd") & "'"
            vlSql = vlSql & " AND v.fec_terdecing >='" & Format(iFecCalculo, "yyyymmdd") & "'"
            'vlSQL = vlSQL & " AND v.cod_estvigencia = 'V'" 'para que obtenga el último ingresado
        Else
            vlSql = vlSql & " AND v.fec_efecto <='" & Format(iFecCalculo, "yyyymmdd") & "'"
            vlSql = vlSql & " AND v.fec_suspension >='" & Format(iFecCalculo, "yyyymmdd") & "'"
        End If
    Else
        vlSql = vlSql & " AND v.fec_efecto <='" & Format(iFecCalculo, "yyyymmdd") & "'"
        vlSql = vlSql & " AND v.fec_suspension >='" & Format(iFecCalculo, "yyyymmdd") & "'"
    End If
    Set vlTB = vgConexionBD.Execute(vlSql)
    If Not vlTB.EOF Then
        oValAsignacion = vlTB!Mto_Carga
    End If
    
    fgCalculaValorCargaFamiliar = True
        
Errores:
    If Err.Number <> 0 Then
        MsgBox "Se ha producido un error al Insertar el Pago" & Chr(13) & Err.Description, vbCritical, "Error de Datos"
    End If
End Function

Function fgCalculaEdad(iFecNac, iFecIniPag) As Integer
'Calcula Edad del Pensionado a la Fecha de Pago
On Error GoTo Errores
Dim vlEdad As Integer

fgCalculaEdad = "-1"

vlEdad = DateDiff("m", DateSerial(Mid(iFecNac, 1, 4), Mid(iFecNac, 5, 2), Mid(iFecNac, 7, 2)), iFecIniPag) - 1 'Edad del Beneficiario
If vlEdad >= 0 Then
    fgCalculaEdad = vlEdad
End If

Errores:
    If Err.Number <> 0 Then
        MsgBox "Error al Calcular Edad del Pensionado" & Chr(13) & Err.Description, vbCritical, "Error"
    End If
End Function

Function fgCargarVariablesGlobales()
    'Carga Constantes para Conceptos Fijos
    stDatGenerales.Cod_ConceptoDesctoSalud = "24"
    'stDatGenerales.Cod_ConceptoImpuesto = "23"
    stDatGenerales.Cod_ConceptoPension = "01"
    'stDatGenerales.Cod_ConceptoAsigFami = "08"
    stDatGenerales.Cod_ConceptoRetencion = "31"
    'stDatGenerales.Cod_ConceptoRetJudAsig = "51" 'Retencion de los Hijos
    'stDatGenerales.Cod_ConceptoRetAsig = "52" 'Retención de la Conyuge
    stDatGenerales.Cod_ConceptoCobroRetencion = "60"
    'stDatGenerales.Cod_ConceptoCobroRetJudAsig = "61" 'Cobro de Retencion de los Hijos
    'stDatGenerales.Cod_ConceptoCobroAsigRet = "62" 'Cobro de Retención de la Conyuge
    'stDatGenerales.Cod_ConceptoGarantiaEstatal = "03"
    'stDatGenerales.Cod_ConceptoCajaCompensacion = "25" 'Cotización a Cajas de Compensación
    
    'Reliquidaciones de Asignación Familiar
    'stDatGenerales.Cod_ConceptoAsigFamCobro = "30" 'Reintegro a la Cia
    'stDatGenerales.Cod_ConceptoAsigFamPago = "09" 'Retroactiva
    
    'Reliquidaciones de Pensión
    stDatGenerales.Cod_ConceptoPensionCobro = "20"
    stDatGenerales.Cod_ConceptoPensionPago = "05"
    stDatGenerales.Cod_ConceptoPensionPagoRetro = "06"
    
    'Reliquidaciones de Garantía Estatal
    'stDatGenerales.Cod_ConceptoGeCobro = "21"
    'stDatGenerales.Cod_ConceptoGEPago = "04"
    
    stDatGenerales.MesesEdad18 = 18 * 12
    stDatGenerales.MesesEdad24 = 24 * 12
    stDatGenerales.MesesEdad28 = 28 * 12
    stDatGenerales.MesesEdad65 = 65 * 12
    stDatGenerales.MesesEdad60 = 60 * 12
    stDatGenerales.Cod_ConceptoGratificacion = "80"
    
    stDatGenerales.Cod_PensionAnticipada = "05"
    
    'stDatGenerales.Cod_ConceptoDesctoSalud2 = "34" 'Concepto de Descuento de Salud 2
    stDatGenerales.Cod_ConceptoPensionQuiebra = "02" 'Concepto de Garantía Estatal por Quiebra de la Compañía
End Function

Function fgConvierteEdadAños(iEdad)
'Convierte la Edad Calculada en Meses a Años

fgConvierteEdadAños = (iEdad \ 12)

End Function

Function fgInsertaDetallePensionProv(iDetalle As TyDetPension, Optional iTipoTabla, Optional iModulo) As Boolean
Dim vlSql As String
On Error GoTo Errores
fgInsertaDetallePensionProv = False
'Función encargada de Insertar un registro en la Tabla de
'Detalle de las Pensiones Provisorias (TP_TMAE_PAGOPENPROV)
If Not IsMissing(iTipoTabla) Then
    If iTipoTabla = "C" Then 'Tabla de Control
        vlSql = "INSERT INTO PP_TTMP_CONPAGOPEN "
    Else
        vlSql = "INSERT INTO PP_TMAE_PAGOPENPROV "
    End If
Else
    vlSql = "INSERT INTO PP_TMAE_PAGOPENPROV "
End If
vlSql = vlSql & "(NUM_PERPAGO, NUM_POLIZA, " 'NUM_ENDOSO,"
vlSql = vlSql & "NUM_ORDEN, COD_CONHABDES, "
vlSql = vlSql & "FEC_INIPAGO, FEC_TERPAGO, MTO_CONHABDES, NUM_IDENRECEPTOR, COD_TIPOIDENRECEPTOR, COD_TIPRECEPTOR"
If Not IsMissing(iModulo) Then
    vlSql = vlSql & ", COD_TIPOMOD"
End If
vlSql = vlSql & ") VALUES ("
vlSql = vlSql & "" & iDetalle.Num_PerPago & ","
vlSql = vlSql & "'" & iDetalle.num_poliza & "',"
'vlSQL = vlSQL & "" & iDetalle.num_endoso & ","
vlSql = vlSql & "" & iDetalle.Num_Orden & ","
vlSql = vlSql & "'" & iDetalle.Cod_ConHabDes & "',"
vlSql = vlSql & "'" & iDetalle.Fec_IniPago & "',"
vlSql = vlSql & "'" & iDetalle.Fec_TerPago & "',"
vlSql = vlSql & "" & str(iDetalle.Mto_ConHabDes) & ","
vlSql = vlSql & "'" & (iDetalle.Num_IdenReceptor) & "',"
vlSql = vlSql & "" & str(iDetalle.Cod_TipoIdenReceptor) & ","
vlSql = vlSql & "'" & iDetalle.Cod_TipReceptor & "'"
If Not IsMissing(iModulo) Then
    vlSql = vlSql & ", '" & iModulo & "'"
End If
vlSql = vlSql & ")"
vgConexionTransac.Execute (vlSql)
'Call FgGuardaLog(vlSql, vgUsuario, "6086") 'RRR 25/07/2014
fgInsertaDetallePensionProv = True
Errores:
    If Err.Number <> 0 Then
        MsgBox "Se ha producido un error al Insertar el Pago." & Chr(13) & "Póliza: " & iDetalle.num_poliza & Chr(13) & "Nº Orden: " & iDetalle.Num_Orden & Chr(13) & "Concepto: " & iDetalle.Cod_ConHabDes & Chr(13) & Err.Description, vbCritical, "Error de Datos"
    End If
End Function

Function fgInsertaLiquidacion(iLiquidacion As TyLiquidacion, Optional iTipoTabla, Optional iModulo) As Boolean
Dim vlSql As String
On Error GoTo Errores
fgInsertaLiquidacion = False

'Función encargada de Insertar un registro en la Tabla de
'Liquidacion de las Pensiones Provisorias (PP_TMAE_LIQPAGOPENPROV)
If Not IsMissing(iTipoTabla) Then
    If iTipoTabla = "C" Then 'Tabla de Control
        vlSql = "INSERT INTO PP_TTMP_CONLIQPAGOPEN "
    Else
        vlSql = "INSERT INTO PP_TMAE_LIQPAGOPENPROV "
    End If
Else
    vlSql = "INSERT INTO PP_TMAE_LIQPAGOPENPROV "
End If
vlSql = vlSql & "(NUM_PERPAGO, NUM_POLIZA, NUM_ENDOSO, "
vlSql = vlSql & "NUM_ORDEN, COD_TIPOPAGO, FEC_PAGO, GLS_DIRECCION, "
vlSql = vlSql & "COD_DIRECCION, COD_TIPPENSION, COD_VIAPAGO, "
vlSql = vlSql & "COD_BANCO, COD_TIPCUENTA, NUM_CUENTA, "
vlSql = vlSql & "COD_SUCURSAL, COD_INSSALUD, "
vlSql = vlSql & "NUM_IDENRECEPTOR, COD_TIPOIDENRECEPTOR, "
vlSql = vlSql & "GLS_NOMRECEPTOR, GLS_NOMSEGRECEPTOR, GLS_PATRECEPTOR, "
vlSql = vlSql & "GLS_MATRECEPTOR, COD_TIPRECEPTOR, " 'NUM_CARGAS, "
vlSql = vlSql & "MTO_HABER, MTO_DESCUENTO, MTO_LIQPAGAR, "
vlSql = vlSql & "MTO_BASEIMP, MTO_BASETRI"
If Not IsMissing(iModulo) Then
    vlSql = vlSql & ", COD_TIPOMOD, COD_USUARIO"
End If
vlSql = vlSql & ", MTO_PENSION, COD_MODSALUD, MTO_PLANSALUD, COD_MONEDA" ' HQR 31/03/2005
vlSql = vlSql & " )VALUES ("
vlSql = vlSql & "'" & iLiquidacion.Num_PerPago & "',"
vlSql = vlSql & "'" & iLiquidacion.num_poliza & "',"
vlSql = vlSql & "" & iLiquidacion.num_endoso & ","
vlSql = vlSql & "" & iLiquidacion.Num_Orden & ","
vlSql = vlSql & "'" & iLiquidacion.Cod_TipoPago & "',"
vlSql = vlSql & "'" & iLiquidacion.Fec_Pago & "',"
vlSql = vlSql & "'" & iLiquidacion.Gls_Direccion & "',"
If iLiquidacion.Cod_Direccion <> "NULL" Then
    vlSql = vlSql & "'" & iLiquidacion.Cod_Direccion & "',"
Else
    vlSql = vlSql & "NULL,"
End If
vlSql = vlSql & "'" & iLiquidacion.Cod_TipPension & "',"
vlSql = vlSql & "'" & iLiquidacion.Cod_ViaPago & "',"
If iLiquidacion.Cod_Banco <> "NULL" Then
    vlSql = vlSql & "'" & iLiquidacion.Cod_Banco & "',"
Else
    vlSql = vlSql & "NULL,"
End If
If iLiquidacion.Cod_TipCuenta <> "NULL" Then
    vlSql = vlSql & "'" & iLiquidacion.Cod_TipCuenta & "',"
Else
    vlSql = vlSql & "NULL,"
End If
If iLiquidacion.Num_Cuenta <> "NULL" Then
    vlSql = vlSql & "'" & iLiquidacion.Num_Cuenta & "',"
Else
    vlSql = vlSql & "NULL,"
End If
If iLiquidacion.Cod_Sucursal <> "NULL" Then
    vlSql = vlSql & "'" & iLiquidacion.Cod_Sucursal & "',"
Else
    vlSql = vlSql & "NULL,"
End If
vlSql = vlSql & "'" & iLiquidacion.Cod_InsSalud & "',"

vlSql = vlSql & "'" & iLiquidacion.Num_IdenReceptor & "',"
vlSql = vlSql & "" & iLiquidacion.Cod_TipoIdenReceptor & ","
vlSql = vlSql & "'" & iLiquidacion.Gls_NomReceptor & "',"
vlSql = vlSql & "'" & iLiquidacion.Gls_NomSegReceptor & "',"
vlSql = vlSql & "'" & iLiquidacion.Gls_PatReceptor & "',"
vlSql = vlSql & "'" & iLiquidacion.Gls_MatReceptor & "',"
vlSql = vlSql & "'" & iLiquidacion.Cod_TipReceptor & "',"
'vlSql = vlSql & "" & iLiquidacion.Num_Cargas & ","
vlSql = vlSql & "" & str(iLiquidacion.Mto_Haber) & ","
vlSql = vlSql & "" & str(iLiquidacion.Mto_Descuento) & ","
vlSql = vlSql & "" & str(iLiquidacion.Mto_LiqPagar) & ","
vlSql = vlSql & "" & str(iLiquidacion.Mto_BaseImp) & ","
vlSql = vlSql & "" & str(iLiquidacion.Mto_BaseTri) & ""
If Not IsMissing(iModulo) Then
    vlSql = vlSql & ",'" & iModulo & "','" & vgUsuario & "'"
End If
vlSql = vlSql & "," & str(iLiquidacion.Mto_Pension) & ""
vlSql = vlSql & ",'" & (iLiquidacion.Cod_ModSalud) & "'"
vlSql = vlSql & "," & str(iLiquidacion.Mto_PlanSalud) & ""
vlSql = vlSql & ",'" & iLiquidacion.Cod_Moneda & "'"
vlSql = vlSql & ")"
vgConexionTransac.Execute (vlSql)
'Call FgGuardaLog(vlSql, vgUsuario, "6086") 'RRR 25/07/2014
fgInsertaLiquidacion = True
Errores:
    If Err.Number <> 0 Then
        MsgBox "Se ha producido un error al Insertar la Liquidación del Pago." & Chr(13) & "Póliza: " & iLiquidacion.num_poliza & Chr(13) & "Nº Orden: " & iLiquidacion.Num_Orden & Chr(13) & Err.Description, vbCritical, "Error de Datos"
    End If
End Function
Function fgObtieneHaberesDescuentos(iTipoCal As String, iPeriodo As String, iPoliza As String, iEndosos, iOrden As Long, iImponible As String, iTributable As String, oMonto As Double, iBaseImp As Double, iBaseTrib As Double, iFechaPago As String, iLiquidacion As TyLiquidacion, istDetPension As TyDetPension, iUf As Double, iFecIniPag As Date, iFecTerPag As Date, iMonedaPension As String, iTipoCambioPension As Double, Optional iTipoTabla, Optional iModulo) As Boolean
'Obtiene los Haberes y Descuentos de Acuerdo
Dim vlSql As String
Dim vlTb3 As ADODB.Recordset
Dim vlTb4 As ADODB.Recordset
Dim vlMonto As Double
Dim vlMoneda As String 'Tipo de Moneda
Dim vlValMoneda As Double 'Tipo de Cambio
Dim vlMontoMaximoCCAF As Double 'Monto Máximo a Pagar por Aporte de CCAF
Dim sCCAF As String
Dim vlMontoAnt As Double 'Monto Pagado Anteriormente de un concepto
Dim vlNumCuotasCon As Long 'Número de Cuotas del Concepto
On Error GoTo Errores
        
fgObtieneHaberesDescuentos = False
oMonto = 0 'Monto de la Suma y Resta de los Haberes y Descuentos
vlSql = "SELECT a.mto_cuota, a.mto_total, a.cod_conhabdes, a.cod_moneda, a.fec_inihabdes, b.cod_tipmov, b.cod_modorigen"
vlSql = vlSql & " FROM PP_TMAE_HABDES A, MA_TPAR_CONHABDES B"
vlSql = vlSql & " WHERE A.COD_CONHABDES = B.COD_CONHABDES"
vlSql = vlSql & " AND A.NUM_POLIZA = '" & iPoliza & "'"
'vlSQL = vlSQL & " AND A.NUM_ENDOSO = " & iEndosos
vlSql = vlSql & " AND A.NUM_ORDEN = " & iOrden
vlSql = vlSql & " AND FEC_INIHABDES <= '" & Format(iFecIniPag, "yyyymmdd") & "'"
vlSql = vlSql & " AND FEC_TERHABDES >= '" & Format(iFecTerPag, "yyyymmdd") & "'"
vlSql = vlSql & " AND B.COD_IMPONIBLE = '" & iImponible & "'"
vlSql = vlSql & " AND B.COD_TRIBUTABLE = '" & iTributable & "'"
vlSql = vlSql & " AND (A.FEC_SUSHABDES IS NULL OR A.FEC_SUSHABDES >= '" & Format(iFecTerPag, "yyyymmdd") & "')" 'Validar si está suspendido
If Not IsMissing(iModulo) Then
    If iModulo = "AF" Or iModulo = "GE" Then  'Solo para Asignación Familiar se filtra, ya que en los demás se necesita el sueldo detallado
        vlSql = vlSql & " AND B.COD_MODORIGEN = '" & iModulo & "'"
    End If
End If
vlSql = vlSql & " ORDER BY COD_MONEDA"

Set vlTb3 = vgConexionBD.Execute(vlSql)
If Not vlTb3.EOF Then
    Do While Not vlTb3.EOF
        If vlTb3!Cod_Moneda = "PRCIM" Then '% sobre la Renta Imponible
            vlMonto = iBaseImp * (vlTb3!MTO_CUOTA / 100)
        Else
            If vlTb3!Cod_Moneda = "PRCTR" Then '% sobre la Renta Tributable
                vlMonto = iBaseTrib * (vlTb3!MTO_CUOTA / 100)
            'HQR 04/03/2006 Se agrega tratamiento de Porcentajes, con monto final a cobrar Fijo
            ElseIf vlTb3!Cod_Moneda = "PRIMF" Or vlTb3!Cod_Moneda = "PRTMF" Then
                If vlTb3!Cod_Moneda = "PRIMF" Then '% Sobre la Renta Imponible, Monto Total
                    vlMonto = iBaseImp * (vlTb3!MTO_CUOTA / 100)
                Else '% Sobre la Renta Tributable, Monto Total
                    vlMonto = iBaseTrib * (vlTb3!MTO_CUOTA / 100)
                End If
                vlMontoAnt = fgObtieneMontoTotalConcepto(iPeriodo, iPoliza, iOrden, vlTb3!Cod_ConHabDes, vlTb3!Fec_IniHabDes, vlNumCuotasCon) 'Obtiene Monto Pagado Anteriormente por este Concepto
                If vlMontoAnt >= vlTb3!mto_total Then 'Si ya se pagó completo, el monto es Cero (no debería pasar este caso)
                    vlMonto = 0
                Else
                    If vlMontoAnt + vlMonto >= vlTb3!mto_total Then 'Es la última Cuota, se paga el saldo
                        vlMonto = vlTb3!mto_total - vlMontoAnt
                        'Se debe Actualizar Fecha de Término del Pago para que no esté siempre considerando este registro (Solo si es Cálculo Definitivo)
                        If iTipoCal = "D" Then
                            vlSql = "UPDATE pp_tmae_habdes SET fec_terhabdes = '" & Format(iFecTerPag, "yyyymmdd") & "', "
                            vlSql = vlSql & " num_cuotas = " & (vlNumCuotasCon + 1)
                            vlSql = vlSql & " WHERE num_poliza = '" & iPoliza & "'"
                            vlSql = vlSql & " AND num_orden = " & iOrden
                            vlSql = vlSql & " AND cod_conhabdes = '" & vlTb3!Cod_ConHabDes & "'"
                            vlSql = vlSql & " AND fec_inihabdes = '" & Format(iFecIniPag, "yyyymmdd") & "'"
                            vgConexionTransac.Execute (vlSql)
                            'Call FgGuardaLog(vlSql, vgUsuario, "6087") 'RRR 25/07/2014
                        End If
                    End If
                End If
            'FIN HQR 04/03/2006
            Else
                If vlTb3!Cod_Moneda = iMonedaPension Then 'Monto en la misma moneda de la pensión, no se debe convertir
                    vlMonto = vlTb3!MTO_CUOTA
                Else
                    If vlTb3!Cod_Moneda = "US" Then 'Monto en US, se obtiene el Cambio de Variable General
                        vlMonto = Format((vlTb3!MTO_CUOTA * iUf) / iTipoCambioPension, "##0.00")
                    Else
                        'Obtiene Monto en Moneda de la Pensión
                        If vlMoneda <> vlTb3!Cod_Moneda Then 'Si no ha cambiado la Moneda, no es necesario volver a Buscar el Cambio
                            vlMoneda = vlTb3!Cod_Moneda
                            If Not fgObtieneConversion(iFechaPago, vlMoneda, vlValMoneda) Then
                                fgObtieneHaberesDescuentos = False
                                MsgBox "Debe Ingresar Tipo de Cambio para la Moneda: " & vlMoneda, vbCritical, "Faltan Datos"
                                Exit Function
                            End If
                        End If
                        vlMonto = Format((vlTb3!MTO_CUOTA * vlValMoneda) / iTipoCambioPension, "##0.00")
                         
                    End If
                End If
            End If
        End If
        vlMonto = Format(vlMonto, "###,##0.00")
        If vlTb3!cod_tipmov = "H" Then 'Haber
            oMonto = oMonto + vlMonto
            iLiquidacion.Mto_Haber = iLiquidacion.Mto_Haber + vlMonto
        Else
            If vlTb3!cod_tipmov = "D" Then 'Descuento
                oMonto = oMonto - vlMonto
                iLiquidacion.Mto_Descuento = iLiquidacion.Mto_Descuento + vlMonto
            End If
            'Si no, es "OTRO" y no se hace nada
        End If
        
        'Llena Datos Faltantes de la Estructura
        istDetPension.Cod_ConHabDes = vlTb3!Cod_ConHabDes
        istDetPension.Mto_ConHabDes = Abs(vlMonto)
        
        'Graba Haber o Descuento
        If Not fgInsertaDetallePensionProv(istDetPension, iTipoTabla, iModulo) Then
            fgObtieneHaberesDescuentos = False
            MsgBox "Se ha producido un Error al Grabar un Haber o Descuento" & Chr(13) & Err.Description, vbCritical, "Error al Grabar"
            Exit Function
        End If
        vlTb3.MoveNext
    Loop
End If

fgObtieneHaberesDescuentos = True
Errores:
    If Err.Number <> 0 Then
        MsgBox "Error al Obtener Haberes y Descuentos.  Póliza: " & iPoliza & "Numero de Orden :" & iOrden
        Exit Function
    End If
End Function


Function fgObtieneImpuestoUnico(iBaseTrib, oImpuesto, iPerPago) As Boolean
'Cálculo del Impuesto Único
On Error GoTo Errores
Dim vlSql As String
Dim vlTB As ADODB.Recordset
Dim vlImpuesto As Double

    fgObtieneImpuestoUnico = False
    vlImpuesto = 0
    vlSql = "SELECT MTO_REBAJA, PRC_FACTOR"
    vlSql = vlSql & " FROM MA_TVAL_IMPUNICOPESOS"
    vlSql = vlSql & " WHERE NUM_PERIODO = '" & iPerPago & "'"
    vlSql = vlSql & " AND MTO_INITRAMO <= " & str(iBaseTrib)
    vlSql = vlSql & " AND MTO_TERTRAMO >= " & str(iBaseTrib)
    Set vlTB = vgConexionBD.Execute(vlSql)
    If Not vlTB.EOF Then
        vlImpuesto = (iBaseTrib * (vlTB!PRC_FACTOR)) - vlTB!MTO_REBAJA
    End If
    If vlImpuesto < 0 Then vlImpuesto = 0
    
    oImpuesto = Format(vlImpuesto, "###,##0")
    
    fgObtieneImpuestoUnico = True
Errores:
If Err.Number <> 0 Then
    MsgBox "Se ha producido un Error al Calcular el Impuesto Único" & Chr(13) & Err.Description, vbCritical, "Error al Calcular Impuesto Único"
End If
End Function

Function fgObtienePensionMinima(iFecPago As String, iParentesco As String, iSitInv As String, iSexo As String, iEdadAños As Integer, iPension As Double, oPensionMinima As Double) As Boolean
    Dim vlSql As String
    Dim vlTB As ADODB.Recordset
    'Obtiene Monto de la Pensión Mínima
    vlSql = "SELECT MTO_PENMINFIN FROM PP_TVAL_PENMINIMA"
    vlSql = vlSql & " WHERE FEC_INIPENMIN <= '" & iFecPago & "'"
    vlSql = vlSql & " AND COD_PAR = " & iParentesco
    vlSql = vlSql & " AND COD_SITINV = '" & iSitInv & "'"
    vlSql = vlSql & " AND COD_SEXO = '" & iSexo & "'"
    vlSql = vlSql & " AND NUM_EDADINI <= " & iEdadAños
    vlSql = vlSql & " AND NUM_EDADFIN >= " & iEdadAños
    vlSql = vlSql & " AND FEC_TERPENMIN >= '" & iFecPago & "'"
    'vlSQL = vlSQL & " AND MTO_PENMINFIN > " & Str(iPension)
    'vlSQL = vlSQL & " ORDER BY MTO_PENMINFIN"
    Set vlTB = vgConexionBD.Execute(vlSql)
    If Not vlTB.EOF Then
        oPensionMinima = vlTB!mto_penminfin 'Siempre se obtiene el Monto de la Pensión que debería ganar
        If vlTB!mto_penminfin > iPension Then
            fgObtienePensionMinima = True
        End If
    Else
        oPensionMinima = 0
        Exit Function
    End If
End Function

Function fgObtienePensionQuiebra(iFechago As String, iParentesco As String, iSitInv As String, iSexo As String, iEdadAños As Integer, iPension As Double, iUf As Double, oPensionQuiebra As Double) As Boolean
'Obtiene Monto de la Pensión en Quiebra
Dim vlPensionMinima As Double
Dim vlPensionMaxima As Double

    If fgObtienePensionMinima(iFechago, iParentesco, iSitInv, iSexo, iEdadAños, iPension, vlPensionMinima) Then
        'Está bajo la mínima
        oPensionQuiebra = iPension 'Se paga 100% de la Pensión
    Else
        'No está bajo la mínima
        If vlPensionMinima > 0 Then
            'Si pudo Obtener Pensión Mínima
            oPensionQuiebra = Format(vlPensionMinima + ((stDatGenerales.Prc_Castigo / 100) * (iPension - vlPensionMinima)), "##0") 'Pensión Mìnima + 75% del diferencial entre la pensión minima y la pensión bruta
            vlPensionMaxima = Format(iUf * stDatGenerales.Mto_TopMaxQuiebra, "##0")
            If oPensionQuiebra > vlPensionMaxima Then
                oPensionQuiebra = vlPensionMaxima
            End If
        Else
            fgObtienePensionQuiebra = False
        End If
    End If
    fgObtienePensionQuiebra = True
End Function


Function fgObtieneTutor(iPoliza, iEndoso, iOrden, iFecIniPag, iFecTerPag, ostLiquidacion As TyLiquidacion) As Integer
'Obtiene Tutor del Beneficiario
'Devuelve : -1, Si es Error
'            0, Si No Encontró Datos
'            1, Si Encontró Titular
On Error GoTo Errores
Dim vlSql As String
Dim vlTB As ADODB.Recordset
Dim vlImpuesto As Double

    fgObtieneTutor = -1
    vlImpuesto = 0
    vlSql = "SELECT *"
    vlSql = vlSql & " FROM PP_TMAE_TUTOR"
    vlSql = vlSql & " WHERE NUM_POLIZA = '" & iPoliza & "'"
    'vlSQL = vlSQL & " AND NUM_ENDOSO = " & iEndoso
    vlSql = vlSql & " AND NUM_ORDEN = " & iOrden
    vlSql = vlSql & " AND FEC_INIPODNOT <= '" & Format(iFecIniPag, "yyyymmdd") & "'"
    vlSql = vlSql & " AND FEC_TERPODNOT >= '" & Format(iFecTerPag, "yyyymmdd") & "'"
    Set vlTB = vgConexionBD.Execute(vlSql)
    If Not vlTB.EOF Then
        ostLiquidacion.Num_IdenReceptor = vlTB!num_identut
        ostLiquidacion.Cod_TipoIdenReceptor = vlTB!cod_tipoidentut
        ostLiquidacion.Gls_NomReceptor = vlTB!gls_nomtut
        ostLiquidacion.Gls_NomSegReceptor = IIf(IsNull(vlTB!gls_nomsegtut), "", vlTB!gls_nomsegtut)
        ostLiquidacion.Gls_PatReceptor = vlTB!gls_pattut
        ostLiquidacion.Gls_MatReceptor = IIf(IsNull(vlTB!gls_mattut), "", vlTB!gls_mattut)
        ostLiquidacion.Cod_TipReceptor = "T" 'Tutor
        ostLiquidacion.Gls_Direccion = vlTB!gls_dirtut
        ostLiquidacion.Cod_Direccion = vlTB!Cod_Direccion
        ostLiquidacion.Cod_ViaPago = vlTB!Cod_ViaPago
        ostLiquidacion.Cod_Banco = IIf(IsNull(vlTB!Cod_Banco), "NULL", vlTB!Cod_Banco)
        ostLiquidacion.Cod_TipCuenta = IIf(IsNull(vlTB!Cod_TipCuenta), "NULL", vlTB!Cod_TipCuenta)
        ostLiquidacion.Num_Cuenta = IIf(IsNull(vlTB!Num_Cuenta), "NULL", vlTB!Num_Cuenta)
        ostLiquidacion.Cod_Sucursal = IIf(IsNull(vlTB!Cod_Sucursal), "NULL", vlTB!Cod_Sucursal)
        fgObtieneTutor = 1
    Else
        fgObtieneTutor = 0
    End If
    
Errores:
If Err.Number <> 0 Then
    MsgBox "Se ha producido un Error al obtener el Tutor" & Chr(13) & Err.Description, vbCritical, "Error al Obtener Tutor"
End If
End Function

Function fgObtieneParametroVigencia(iTabla, iElemento, iVigencia, oValor) As Boolean
'Obtiene Parámetros Generales de Tabla de Vigencias MA_TPAR_TABCODVIG
On Error GoTo Errores
Dim vlSql As String
Dim vlTB As ADODB.Recordset

fgObtieneParametroVigencia = False
vlSql = "SELECT MTO_ELEMENTO"
vlSql = vlSql & " FROM MA_TPAR_TABCODVIG"
vlSql = vlSql & " WHERE COD_TABLA = '" & iTabla & "'"
vlSql = vlSql & " AND COD_ELEMENTO = '" & iElemento & "'"
vlSql = vlSql & " AND FEC_INIVIG <= '" & iVigencia & "'"
vlSql = vlSql & " AND FEC_TERVIG >= '" & iVigencia & "'"
Set vlTB = vgConexionBD.Execute(vlSql)
If Not vlTB.EOF Then
    oValor = vlTB!mto_elemento
Else
    vlTB.Close
    Exit Function
End If
vlTB.Close
fgObtieneParametroVigencia = True

Errores:
If Err.Number <> 0 Then
    MsgBox "Se ha producido un Error al Obtener Parámetros Generales de Vigencia" & Chr(13) & Err.Description, vbCritical, "Error al Obtener Datos Generales"
End If

End Function

Function fgObtienePrcSalud(iInstitucion As String, iModalidad As String, iMonto As Double, iBaseImp As Double, oValor As Double, iUf As Double, iFecPago As String, iMonedaPension As String, iTipoCambioPension As Double) As Boolean
'Obtiene el Monto a Descontar por Salud
On Error GoTo Errores
Dim vlValor As Double
Dim vlValor2 As Double
Dim vlConversion As Double
Dim vlTopeMinimo As Double
Dim vlTopeMaximo As Double
'Código de elemento, en tabla ma_tpar_tabcod de Institución Recaudadora Fonasa
Const clCodInstRecauda As String * 2 = "13"
Dim vlMaximoSalud As Double

fgObtienePrcSalud = False
vlValor = 0
vlValor2 = 0

If iModalidad = "PORCE" Then 'PORCENTAJE
    vlValor = iBaseImp * (iMonto / 100)
Else
    If iModalidad = iMonedaPension Then 'Moneda de la Pensión
        vlValor = iMonto 'No se realiza Conversion
    Else
        If iModalidad = "US" Then 'Monto en UF, se obtiene Conversión de Variable General
            vlValor = Format((iUf * iMonto) / iTipoCambioPension, "#0.00")
        Else
            'OTRA MONEDA
            If Not fgObtieneConversion(iFecPago, iModalidad, vlConversion) Then
                MsgBox "Debe ingresar el Tipo de Cambio de la Moneda '" & iModalidad & "'", vbCritical, "Falta Tipo de Cambio"
                Exit Function
            End If
            vlValor = Format((vlConversion * iMonto) / iTipoCambioPension, "#0.00")
        End If
    End If
End If

vlMaximoSalud = 0
If Not fgObtieneMontoMaximoSalud(iMonedaPension, iFecPago, vlMaximoSalud) Then
    Exit Function
End If
vlTopeMaximo = vlMaximoSalud

If vlValor > vlTopeMaximo Then 'Si es Mas de 4.2 UF debe cobrarse 4.2 UF
    vlValor = vlTopeMaximo
End If

oValor = Format(vlValor, "###,##0.00")
fgObtienePrcSalud = True
Errores:
    If Err.Number <> 0 Then
        MsgBox "Se ha producido un Error al Calcular descuento por Salud" & Chr(13) & Err.Description, vbCritical, "Error en Cálculo de Salud"
        Exit Function
    End If
End Function

Function fgObtieneMontoMaximoSalud(iMoneda, iFecha, oValor) As Boolean
'Obtiene Monto Máximo de Salud

On Error GoTo Errores
Dim vlSql As String
Dim vlTB As ADODB.Recordset

fgObtieneMontoMaximoSalud = False
vlSql = "SELECT a.mto_elemento"
vlSql = vlSql & " FROM ma_tval_saludmax a"
vlSql = vlSql & " WHERE cod_moneda = '" & iMoneda & "'"
vlSql = vlSql & " AND fec_inivig <= '" & iFecha & "'"
vlSql = vlSql & " AND fec_tervig >= '" & iFecha & "'"
Set vlTB = vgConexionBD.Execute(vlSql)
If Not vlTB.EOF Then
    oValor = vlTB!mto_elemento
Else
    vlTB.Close
    MsgBox "No está definido Monto Máximo de Salud para la Moneda de la Pensión", vbCritical
    Exit Function
End If
vlTB.Close
fgObtieneMontoMaximoSalud = True

Errores:
If Err.Number <> 0 Then
    MsgBox "Se ha producido un Error al Obtener Monto Máximo de Salud" & Chr(13) & Err.Description, vbCritical, "Error al Obtener Datos Generales"
End If

End Function



Function fgTraspasaDatosaEstructuraRetencion(istDetPago As TyDetPension, ostDetPagoRet As TyDetPension, istLiqPago As TyLiquidacion, ostLiqPagRet As TyLiquidacion) As Boolean
    On Error GoTo Errores
    fgTraspasaDatosaEstructuraRetencion = False
    
    'Traspasa Datos de Estructura de Detalle del Pago
    ostDetPagoRet.Fec_IniPago = istDetPago.Fec_IniPago
    ostDetPagoRet.Fec_TerPago = istDetPago.Fec_TerPago
    ostDetPagoRet.Num_PerPago = istDetPago.Num_PerPago
    ostDetPagoRet.num_poliza = istDetPago.num_poliza
    ostDetPagoRet.num_endoso = istDetPago.num_endoso
    ostDetPagoRet.Num_Orden = istDetPago.Num_Orden
    ostDetPagoRet.Cod_TipReceptor = "R"
    'Traspasa Datos de Estructura de la Liquidación del Pago
    
    ostLiqPagRet.Num_PerPago = istLiqPago.Num_PerPago
    ostLiqPagRet.num_poliza = istLiqPago.num_poliza
    ostLiqPagRet.num_endoso = istLiqPago.num_endoso
    ostLiqPagRet.Num_Orden = istLiqPago.Num_Orden
    ostLiqPagRet.Cod_TipoPago = istLiqPago.Cod_TipoPago
    ostLiqPagRet.Cod_TipPension = istLiqPago.Cod_TipPension
    ostLiqPagRet.Fec_Pago = istLiqPago.Fec_Pago
    ostLiqPagRet.Cod_Moneda = istLiqPago.Cod_Moneda
    ostLiqPagRet.Cod_InsSalud = "NULL"
    'ostLiqPagRet.Cod_CajaCompen = "NULL"
    ostLiqPagRet.Cod_TipReceptor = "R"
    ostLiqPagRet.Mto_BaseImp = 0
    ostLiqPagRet.Mto_BaseTri = 0
    
    fgTraspasaDatosaEstructuraRetencion = True
    
Errores:
If Err.Number <> 0 Then
    MsgBox "Se ha producido un Error al Calcular el Impuesto Único" & Chr(13) & Err.Description, vbCritical, "Error al Calcular Impuesto Único"
End If
End Function

Function fgVerificaCertificado(iPoliza, iOrden, iFecIniPag, iFecTerPag, iTipoCert As String, Optional iModulo) As Integer
'Verifica si el Pensionado tiene Certificado de Estudio Vigente
'Devuelve : -1 Si hay Error
'           0 Si no Encuentra Certificado
'           1 Si Encuentra Certificado
On Error GoTo Errores
    Dim vlSql As String
    Dim vlTB As ADODB.Recordset
    Dim vlFechaTermino As String
    'If iPoliza = "0000000011" Then MsgBox iPoliza
    vlFechaTermino = DateAdd("d", -1, DateAdd("m", (vgMesesExtension + 1), DateSerial(Year(iFecTerPag), Month(iFecTerPag), 1)))
        
    fgVerificaCertificado = "0" 'No Encuentra Certificado
    vlSql = "SELECT COUNT(1) AS CONTADOR FROM PP_TMAE_CERTIFICADO "
    vlSql = vlSql & "WHERE NUM_POLIZA = '" & iPoliza & "'"
    vlSql = vlSql & " AND COD_TIPO = '" & iTipoCert & "'"
    vlSql = vlSql & " AND NUM_ORDEN = " & iOrden
    If Not IsMissing(iModulo) Then
        If iModulo = "REL" Or iModulo = "AFREL" Then 'Reliquidacion de Pensiones
            vlSql = vlSql & " AND FEC_INICER <= '" & Format(iFecIniPag, "yyyymmdd") & "'"
            vlSql = vlSql & " AND FEC_TERCER >= '" & Format(iFecIniPag, "yyyymmdd") & "'"
        Else
            vlSql = vlSql & " AND FEC_EFECTO <= '" & Format(iFecIniPag, "yyyymmdd") & "'"
            vlSql = vlSql & " AND FEC_TERCER >= '" & Format(iFecIniPag, "yyyymmdd") & "'"
        End If
    Else
        vlSql = vlSql & " AND FEC_EFECTO <= '" & Format(iFecIniPag, "yyyymmdd") & "'"
        vlSql = vlSql & " AND FEC_TERCER >= '" & Format(iFecIniPag, "yyyymmdd") & "'"
    End If
    Set vlTB = vgConexionBD.Execute(vlSql)
    If Not vlTB.EOF Then
        If vlTB!contador > 0 Then
            fgVerificaCertificado = "1"
        End If
    End If

Errores:
    If Err.Number <> 0 Then
        fgVerificaCertificado = "-1"
        MsgBox "Se ha producido un Error al Validar Certificado de Estudios" & Chr(13) & Err.Description, vbCritical, "Error al Validar Certificado de Estudios"
    End If
End Function

Function fgVerificaCertificadoEst(iPoliza, iOrden, iEndoso, iFecIniPag) As Integer
'Verifica si el Pensionado tiene Certificado de Estudio Vigente
'Devuelve : -1 Si hay Error
'           0 Si no Encuentra Certificado
'           1 Si Encuentra Certificado
On Error GoTo Errores
    Dim vlSql As String
    Dim vlTB As ADODB.Recordset
  
    fgVerificaCertificadoEst = "0" 'No Encuentra Certificado
    
    vlSql = " select nvl(case when IND_DNI='S' then 1  else 0 end  +  case when IND_DJU='S' then 1  else 0  end  +  case when IND_PES='S' then 1 else 0  end  + case when IND_BNO='S' then 1  else 0  end, 0) Valor"
    'vlSql = vlSql & " from pp_tmae_certificado where num_poliza='" & iPoliza & "' and num_endoso=" & iEndoso & " and num_orden=" & iOrden & " and cod_tipo='EST' AND EST_ACT=1"
    vlSql = vlSql & " from pp_tmae_certificado where num_poliza='" & iPoliza & "' and num_orden=" & iOrden & " and cod_tipo='EST' AND EST_ACT=1"
    vlSql = vlSql & " and (FEC_INICER<='" & Format(iFecIniPag, "yyyymmdd") & "' and FEC_TERCER>='" & Format(iFecIniPag, "yyyymmdd") & "')"
    
    Set vlTB = vgConexionBD.Execute(vlSql)
    If Not vlTB.EOF Then
        fgVerificaCertificadoEst = vlTB!valor
    Else
        fgVerificaCertificadoEst = 0
    End If

Errores:
    If Err.Number <> 0 Then
        fgVerificaCertificadoEst = "-1"
        MsgBox "Se ha producido un Error al Validar Certificado de Estudios" & Chr(13) & Err.Description, vbCritical, "Error al Validar Certificado de Estudios"
    End If
End Function


Function fgObtieneConversion(iFecha, iMoneda, oValor) As Boolean
Dim vlFechaAnterior As String
Dim vlFechaMaxima As String
Dim vlFeriadosMaximo As Long

    'Obtiene Valor de la Moneda a la Fecha ingresada como Parámetro
    Dim vlSql As String
    Dim vlTB As ADODB.Recordset
    
    fgObtieneConversion = False
    vlFeriadosMaximo = 7
    
    If (iMoneda <> vgMonedaCodOfi) Then
        vlSql = "SELECT MTO_MONEDA FROM MA_TVAL_MONEDA"
        vlSql = vlSql & " WHERE COD_MONEDA = '" & iMoneda & "'"
        vlSql = vlSql & " AND FEC_MONEDA = '" & iFecha & "'"
        Set vlTB = vgConexionBD.Execute(vlSql)
        If Not vlTB.EOF Then
            oValor = vlTB!Mto_Moneda
            fgObtieneConversion = True
        End If
        vlTB.Close
    
        If (fgObtieneConversion = False) Then
            vlFechaAnterior = ""
            vlFechaMaxima = Format(DateSerial(Mid(iFecha, 1, 4), Mid(iFecha, 5, 2), CInt(Mid(iFecha, 7, 2)) - vlFeriadosMaximo), "yyyymmdd")
            
            'Debe obtener el primero que encuentre anterior a la fecha consultada
            vlSql = "SELECT FEC_MONEDA, MTO_MONEDA FROM MA_TVAL_MONEDA"
            vlSql = vlSql & " WHERE COD_MONEDA = '" & iMoneda & "'"
            vlSql = vlSql & " AND FEC_MONEDA = "
                vlSql = vlSql & " (SELECT MAX(FEC_MONEDA) FROM MA_TVAL_MONEDA"
                vlSql = vlSql & " WHERE COD_MONEDA = '" & iMoneda & "'"
                vlSql = vlSql & " AND FEC_MONEDA < '" & iFecha & "')"
            Set vlTB = vgConexionBD.Execute(vlSql)
            If Not vlTB.EOF Then
                vlFechaAnterior = vlTB!FEC_MONEDA
                If (vlFechaAnterior > vlFechaMaxima) Then
                    oValor = vlTB!Mto_Moneda
                    fgObtieneConversion = True
                End If
            End If
            vlTB.Close
        End If
    
    Else
        fgObtieneConversion = True
        oValor = 1
    End If
    
End Function

Function fgGeneraTablaImpuestoUnico() As Boolean
'Genera Tabla de Impuesto Único utilizando el Valor de la UTM a la Fecha de Pago
On Error GoTo Errores

    Dim vlSql As String
    Dim vlTB As ADODB.Recordset
    Dim vlConversion As Double
    Dim vlIniTramo As Double
    Dim vlTerTramo As Double
    Dim vlRebaja As Double
    
    fgGeneraTablaImpuestoUnico = False
    
    'Obtiene Conversión de la UTM
    If Not fgObtieneConversion(Format(vgFecIniPag, "yyyymmdd"), "UTM", vlConversion) Then
        MsgBox "Debe ingresar Tipo de Cambio para la Moneda : 'UTM'", vbCritical, "Falta Tipo de Cambio"
        Exit Function
    End If
    
    'Elimina Tabla Generada Anteriormente para el Mismo Periodo
    vlSql = "DELETE FROM MA_TVAL_IMPUNICOPESOS"
    vlSql = vlSql & " WHERE NUM_PERIODO = '" & vgPerPago & "'"
    vgConexionBD.Execute (vlSql)
    
    'Obtiene Tabla en UTM
    vlSql = "SELECT * FROM MA_TVAL_IMPUNICO"
    vlSql = vlSql & " WHERE FEC_INIIMPUNI <= '" & vgFecPago & "'"
    vlSql = vlSql & " AND FEC_TERIMPUNI >= '" & vgFecPago & "'"
    Set vlTB = vgConexionBD.Execute(vlSql)
    If Not vlTB.EOF Then
        Do While Not vlTB.EOF
            vlIniTramo = vlTB!MTO_INITRAMO
            vlTerTramo = vlTB!MTO_TERTRAMO
            vlRebaja = vlTB!MTO_REBAJA
            vlSql = "INSERT INTO MA_TVAL_IMPUNICOPESOS"
            vlSql = vlSql & "(NUM_PERIODO, NUM_TRAMO, MTO_INITRAMO, "
            vlSql = vlSql & "MTO_TERTRAMO, MTO_REBAJA, PRC_FACTOR)"
            vlSql = vlSql & " VALUES ('"
            vlSql = vlSql & vgPerPago & "','" & vlTB!NUM_TRAMO & "',"
            vlSql = vlSql & str(Format(vlTB!MTO_INITRAMO * vlConversion, "###,##0")) & ","
            vlSql = vlSql & str(Format(vlTB!MTO_TERTRAMO * vlConversion, "###,##0")) & ","
            vlSql = vlSql & str(Format(vlTB!MTO_REBAJA * vlConversion, "###,##0")) & ","
            vlSql = vlSql & str(vlTB!PRC_FACTOR) & ")"
            vgConexionBD.Execute (vlSql)
            vlTB.MoveNext
        Loop
    End If
    vlTB.Close
    fgGeneraTablaImpuestoUnico = True

Errores:
    If Err.Number <> 0 Then
        MsgBox "Se ha producido un Error al Generar Tabla de Impuesto Único" & Chr(13) & Err.Description, vbCritical, "Error al Calcular Tabla de Impuesto Único"
    End If
End Function

Function fgInsertaTTMPLiquidacion(iLiquidacion As TyTTMPLiquidacion) As Boolean
Dim vlSql As String
On Error GoTo Errores
fgInsertaTTMPLiquidacion = False

'Función encargada de Insertar un registro en la Tabla Temporal de
'Liquidacion de las Pensiones (PP_TTMP_LIQUIDACION)
vlSql = "INSERT INTO PP_TTMP_LIQUIDACION "
vlSql = vlSql & "( COD_USUARIO, NUM_PERPAGO, NUM_POLIZA,"
vlSql = vlSql & " NUM_ENDOSO, NUM_ORDEN, NUM_ITEM, COD_CONHABER,"
vlSql = vlSql & " MTO_HABER, COD_CONDESCTO, MTO_DESCUENTO, NUM_IDENRECEPTOR, COD_TIPOIDENRECEPTOR, "
vlSql = vlSql & " COD_TIPRECEPTOR, GLS_DIRECCION, FEC_PAGO, GLS_NOMRECEPTOR,"
vlSql = vlSql & " MTO_LIQPAGAR, FEC_TERPODNOT, FEC_PAGOPROXREG, GLS_TIPPENSION, GLS_VIAPAGO,"
vlSql = vlSql & " GLS_INSSALUD, MTO_MONEDA, COD_DIRECCION, MTO_PENSION, NUM_CARGAS, "
vlSql = vlSql & " MTO_LIQHABER, MTO_LIQDESCUENTO, NUM_IDENBEN, COD_TIPOIDENBEN, GLS_NOMBEN,"
vlSql = vlSql & " GLS_DIRECCION2, GLS_MENSAJE, GLS_MONTOPENSION, COD_MONEDA, GLS_AFP, gls_vejez,COD_TIPOIDENTIT,NUM_IDENTIT,GLS_NOMTIT)"
vlSql = vlSql & " VALUES ("
vlSql = vlSql & "'" & iLiquidacion.Cod_Usuario & "',"
vlSql = vlSql & "'" & iLiquidacion.Num_PerPago & "',"
vlSql = vlSql & "'" & iLiquidacion.num_poliza & "',"
vlSql = vlSql & "" & iLiquidacion.num_endoso & ","
vlSql = vlSql & "" & iLiquidacion.Num_Orden & ","
vlSql = vlSql & "" & iLiquidacion.Num_Item & ","
If iLiquidacion.Cod_ConHaber <> "NULL" Then
    vlSql = vlSql & "'" & iLiquidacion.Cod_ConHaber & "',"
    vlSql = vlSql & "" & str(iLiquidacion.Mto_Haber) & ","
Else
    vlSql = vlSql & "NULL, NULL,"
End If
If iLiquidacion.Cod_ConDescto <> "NULL" Then
    vlSql = vlSql & "'" & iLiquidacion.Cod_ConDescto & "',"
    vlSql = vlSql & "" & str(iLiquidacion.Mto_Descuento) & ",'"
Else
    vlSql = vlSql & "NULL, NULL,'"
End If
vlSql = vlSql & iLiquidacion.Num_IdenReceptor & "','"
vlSql = vlSql & iLiquidacion.Cod_TipoIdenReceptor & "','"
vlSql = vlSql & iLiquidacion.Cod_TipReceptor & "','"
vlSql = vlSql & iLiquidacion.Gls_Direccion & "','"
vlSql = vlSql & iLiquidacion.Fec_Pago & "','"
vlSql = vlSql & iLiquidacion.Gls_NomReceptor & "',"
vlSql = vlSql & str(iLiquidacion.Mto_LiqPagar) & ",'"
vlSql = vlSql & iLiquidacion.Fec_TerPodNot & "','"
vlSql = vlSql & iLiquidacion.Fec_PagoProxReg & "','"
vlSql = vlSql & iLiquidacion.Gls_TipPension & "','"
vlSql = vlSql & iLiquidacion.Gls_ViaPago & "','"
vlSql = vlSql & iLiquidacion.Gls_InsSalud & "',"
vlSql = vlSql & str(iLiquidacion.Mto_Moneda) & ","
vlSql = vlSql & iLiquidacion.Cod_Direccion & ","
vlSql = vlSql & str(iLiquidacion.Mto_Pension) & ","
vlSql = vlSql & iLiquidacion.Num_Cargas & ","
vlSql = vlSql & str(iLiquidacion.Mto_LiqHaber) & ","
vlSql = vlSql & str(iLiquidacion.Mto_LiqDescuento) & ",'"
vlSql = vlSql & iLiquidacion.Num_IdenBen & "','"
vlSql = vlSql & iLiquidacion.Cod_TipoIdenBen & "','"
vlSql = vlSql & iLiquidacion.Gls_NomBen & "','"
vlSql = vlSql & iLiquidacion.Gls_Direccion2 & "','"
vlSql = vlSql & iLiquidacion.Gls_Mensaje & "','"
vlSql = vlSql & iLiquidacion.Gls_MontoPension & "','"
vlSql = vlSql & iLiquidacion.Cod_Moneda & "','"
vlSql = vlSql & iLiquidacion.Gls_Afp & "','"
vlSql = vlSql & iLiquidacion.gls_vejez & "',"
vlSql = vlSql & iLiquidacion.Cod_TipoIdenTit & ",'"
vlSql = vlSql & iLiquidacion.Num_IdenTit & "','"
vlSql = vlSql & iLiquidacion.gls_nomTit & "')"

vgConexionBD.Execute (vlSql)
fgInsertaTTMPLiquidacion = True
Errores:
    If Err.Number <> 0 Then
        MsgBox "Se ha producido un error al Insertar el Pago" & Chr(13) & Err.Description, vbCritical, "Error de Datos"
    End If
End Function

Function fgActualizaTTMPLiquidacionDesc(iLiquidacion As TyTTMPLiquidacion) As Boolean
Dim vlSql As String
On Error GoTo Errores
fgActualizaTTMPLiquidacionDesc = False

'Función encargada de Insertar un registro en la Tabla de
'Liquidacion de las Pensiones Provisorias (PP_TMAE_LIQPAGOPENPROV)
vlSql = "UPDATE pp_ttmp_liquidacion"
vlSql = vlSql & " SET cod_condescto = '" & iLiquidacion.Cod_ConDescto & "',"
vlSql = vlSql & " mto_descuento = " & str(iLiquidacion.Mto_Descuento) & ","
vlSql = vlSql & " mto_liqpagar = " & str(iLiquidacion.Mto_LiqPagar) & ","
vlSql = vlSql & " mto_liqdescuento = " & str(iLiquidacion.Mto_LiqDescuento) & ","
vlSql = vlSql & " gls_montopension = '" & Trim(iLiquidacion.Gls_MontoPension) & "'"
vlSql = vlSql & " WHERE num_poliza = '" & iLiquidacion.num_poliza & "'"
vlSql = vlSql & " AND num_endoso = " & iLiquidacion.num_endoso
vlSql = vlSql & " AND num_orden = " & iLiquidacion.Num_Orden
vlSql = vlSql & " AND num_perpago = '" & iLiquidacion.Num_PerPago & "'"
vlSql = vlSql & " AND num_idenreceptor = '" & iLiquidacion.Num_IdenReceptor & "'"
vlSql = vlSql & " AND cod_tipoidenreceptor = '" & iLiquidacion.Cod_TipoIdenReceptor & "'"
vlSql = vlSql & " AND cod_tipreceptor = '" & iLiquidacion.Cod_TipReceptor & "'"
vlSql = vlSql & " AND cod_usuario = '" & iLiquidacion.Cod_Usuario & "'"
vlSql = vlSql & " AND num_item = " & iLiquidacion.Num_Item

vgConexionBD.Execute (vlSql)
fgActualizaTTMPLiquidacionDesc = True
Errores:
    If Err.Number <> 0 Then
        MsgBox "Se ha producido un error al Insertar el Pago" & Chr(13) & Err.Description, vbCritical, "Error de Datos"
    End If
End Function

Function fgActualizaTTMPLiquidacionHab(iLiquidacion As TyTTMPLiquidacion) As Boolean
Dim vlSql As String
On Error GoTo Errores
fgActualizaTTMPLiquidacionHab = False

'Función encargada de Insertar un registro en la Tabla de
'Liquidacion de las Pensiones Provisorias (PP_TMAE_LIQPAGOPENPROV)
vlSql = "UPDATE pp_ttmp_liquidacion"
vlSql = vlSql & " SET cod_conhaber = '" & iLiquidacion.Cod_ConHaber & "',"
vlSql = vlSql & " mto_haber = " & str(iLiquidacion.Mto_Haber) & ","
vlSql = vlSql & " mto_liqpagar = " & str(iLiquidacion.Mto_LiqPagar) & ","
vlSql = vlSql & " mto_liqhaber = " & str(iLiquidacion.Mto_LiqHaber) & ","
vlSql = vlSql & " gls_montopension = '" & Trim(iLiquidacion.Gls_MontoPension) & "'"
vlSql = vlSql & " WHERE num_poliza = '" & iLiquidacion.num_poliza & "'"
vlSql = vlSql & " AND num_endoso = " & iLiquidacion.num_endoso
vlSql = vlSql & " AND num_orden = " & iLiquidacion.Num_Orden
vlSql = vlSql & " AND num_perpago = '" & iLiquidacion.Num_PerPago & "'"
vlSql = vlSql & " AND num_idenreceptor = '" & iLiquidacion.Num_IdenReceptor & "'"
vlSql = vlSql & " AND cod_tipoidenreceptor = '" & iLiquidacion.Cod_TipoIdenReceptor & "'"
vlSql = vlSql & " AND cod_tipreceptor = '" & iLiquidacion.Cod_TipReceptor & "'"
vlSql = vlSql & " AND cod_usuario = '" & iLiquidacion.Cod_Usuario & "'"
vlSql = vlSql & " AND num_item = " & iLiquidacion.Num_Item

vgConexionBD.Execute (vlSql)
fgActualizaTTMPLiquidacionHab = True
Errores:
    If Err.Number <> 0 Then
        MsgBox "Se ha producido un error al Insertar el Pago" & Chr(13) & Err.Description, vbCritical, "Error de Datos"
    End If
End Function

Function fgCalculaRetencion(iPoliza, iEndoso, iOrden, iPerPago, iFechaPago, iBaseImp, iBaseTrib, istLiquidacion As TyLiquidacion, istDetPension As TyDetPension, iFecCalculo, iCaption, iMonedaPension As String, iTipoCambioPension As Double, Optional iTipoTabla, Optional iModulo) As Boolean

Dim stDetPensionRet As TyDetPension
Dim stLiquidacionRet As TyLiquidacion

'Función que calcula el Monto correspondiente a Retención Judicial
'Parametros de Entrada:
' iPoliza           => Nº de Poliza
' iEndoso           => Nº de Endoso (no se utiliza)
' iOrden            => Nº de Orden de a quién se le retiene
' iPerPago          => Periodo de Pago
' iFechaPago        => Fecha a la cual se realiza la Conversión
' iBaseImp          => Base Imponible del Pensionado al que se le retiene parte de la Pensión
' iBaseTrib         => Base Tributable del Pensionado al que se le retiene parte de la Pensión
' istLiquidacionRet => Estructura de la Liquidación del Retenedor
' iFecCalculo       => Fecha a la cual se realizaran los cálculos (Periodo de Pago con día 01)
'Variables Globales:
' stDatGenerales.Val_UF => Valor de UF a la Fecha de Pago

Dim vlSql As String
Dim vlTB As ADODB.Recordset
Dim vlTB2 As ADODB.Recordset
Dim vlRetencion As Double 'Monto de la Retención Judicial
Dim vlMoneda As String 'Tipo de Moneda
Dim vlValMoneda As Double 'Tipo de Cambio
Dim vlAsignacionRet As Double 'Asignación Familiar Retenida
Dim vlNumCargasRet As Long 'Número de Cargas Retenidas
Dim vlcont As Integer 'Variable de Paso para que retorne de la función de Asignación Familiar el número de cargas retenidas
Dim vlValCarga As Double 'Valor de Cada Carga Familiar
Dim vlNumIdenReceptor As String, vlCodTipoIdenReceptor As Long 'Para guardar el anterior
Dim vlTotalRetencion As Double 'Total Retenido de la Pensión
Dim vlTotalAsigFam As Double 'Total Retenido por Asignación Familiar
Dim vlTipRetencion As String 'Tipo de Retención Judicial

On Error GoTo Errores
fgCalculaRetencion = False

vlNumIdenReceptor = ""
vlCodTipoIdenReceptor = -1

vlTotalRetencion = 0
vlTipRetencion = ""

stLiquidacionRet.Mto_Haber = 0
stLiquidacionRet.Mto_LiqPagar = 0
'stLiquidacionRet.Num_Cargas = 0

vlSql = "SELECT * FROM PP_TMAE_RETJUDICIAL"
vlSql = vlSql & " WHERE NUM_POLIZA = '" & iPoliza & "'"
vlSql = vlSql & " AND NUM_ORDEN = " & iOrden
vlSql = vlSql & " AND '" & Format(iFecCalculo, "yyyymmdd") & "'"
'hqr 06/10/2007 Se considera el pago desde la Fecha de Efecto
'vlSql = vlSql & " BETWEEN FEC_INIRET AND FEC_TERRET"
vlSql = vlSql & " BETWEEN FEC_INIRET AND FEC_TERRET"
vlSql = vlSql & " ORDER BY NUM_IDENRECEPTOR"
'vlSQL = vlSQL & " AND (FEC_CAUSUSPENSION IS NULL OR FEC_CAUSUSPENSION >= '" & Format(vgFecIniPag, "yyyymmdd") & "')" 'Validar si está suspendido
Set vlTB = vgConexionBD.Execute(vlSql)
If Not vlTB.EOF Then
    'vlTB.MoveFirst
    vlNumIdenReceptor = vlTB!Num_IdenReceptor
    vlCodTipoIdenReceptor = vlTB!Cod_TipoIdenReceptor
    Do While Not vlTB.EOF
        If vlNumIdenReceptor <> vlTB!Num_IdenReceptor Or vlCodTipoIdenReceptor <> vlTB!Cod_TipoIdenReceptor Then
            If stLiquidacionRet.Mto_LiqPagar <> 0 Then
                'Inserta Liquidación Anterior
                If Not fgInsertaLiquidacion(stLiquidacionRet, iTipoTabla, iModulo) Then
                    Exit Function
                End If
            End If
            'Si se trata de una nueva liquidación, se limpian las variables acumulativas
            stLiquidacionRet.Mto_Haber = 0
            stLiquidacionRet.Mto_LiqPagar = 0
            'stLiquidacionRet.Num_Cargas = 0
        End If
        
        vlRetencion = 0
        vlNumIdenReceptor = vlTB!Num_IdenReceptor
        vlCodTipoIdenReceptor = vlTB!Cod_TipoIdenReceptor
        vlTipRetencion = vlTB!cod_tipret
        If IsNull(vlTB!cod_modret) Then
            vlRetencion = 0
        Else
            If vlTB!cod_modret = "PRCIM" Then '% sobre la Renta Imponible
                vlRetencion = Format(iBaseImp * (vlTB!mto_ret / 100), "##0.00")
            Else
                'If vlTB!cod_modret = "PRCTR" Then '% sobre la Renta Tributable
                If vlTB!cod_modret = "MTOIM" Then 'Cuando es monto
                    vlRetencion = Format(vlTB!mto_ret, "##0.00")
                Else
                    If vlTB!cod_modret = iMonedaPension Then 'Monto en Moneda de la Pensión, no se debe convertir
                        vlRetencion = vlTB!mto_ret
                    Else
                        If vlTB!cod_modret = "US" Then 'Monto en US, se obtiene el Cambio de Variable General
                            vlRetencion = Format((vlTB!mto_ret * stDatGenerales.Val_UF) / iTipoCambioPension, "##0.00")
                        Else
                            'Obtiene Monto en Pesos
                            If vlMoneda <> vlTB!cod_modret Then 'Si no ha cambiado la Moneda, no es necesario volver a Buscar el Cambio
                                vlMoneda = vlTB!cod_modret
                                If Not fgObtieneConversion(iFechaPago, vlMoneda, vlValMoneda) Then
                                    fgCalculaRetencion = False
                                    MsgBox "Debe Ingresar Tipo de Cambio para la Moneda: " & vlMoneda, vbCritical, "Faltan Datos"
                                    Exit Function
                                End If
                            End If
                            vlRetencion = Format((vlTB!mto_ret * vlValMoneda) / iTipoCambioPension, "##0.00")
                        End If
                    End If
                End If
            End If
        End If
        'Retención no puede ser mayor al máximo permitido (Se asume que el máximo está en la moneda de la pensión)
        'If vlRetencion > vlTB!mto_retmax Then
        '    vlRetencion = vlTB!mto_retmax
        'End If
        
        vlTotalRetencion = vlTotalRetencion + vlRetencion
        
        If vlRetencion > 0 Then
            'Traspasa Datos a Estructura de la Liquidacion para el Retenedor
            stLiquidacionRet.Num_IdenReceptor = vlTB!Num_IdenReceptor
            stLiquidacionRet.Cod_TipoIdenReceptor = vlTB!Cod_TipoIdenReceptor
            stLiquidacionRet.Gls_NomReceptor = vlTB!Gls_NomReceptor
            stLiquidacionRet.Gls_NomSegReceptor = IIf(IsNull(vlTB!Gls_NomSegReceptor), "", vlTB!Gls_NomSegReceptor)
            stLiquidacionRet.Gls_PatReceptor = vlTB!Gls_PatReceptor
            stLiquidacionRet.Gls_MatReceptor = IIf(IsNull(vlTB!Gls_MatReceptor), "", vlTB!Gls_MatReceptor)
            stLiquidacionRet.Gls_Direccion = vlTB!gls_dirreceptor
            stLiquidacionRet.Cod_Direccion = vlTB!Cod_Direccion
            stLiquidacionRet.Cod_ViaPago = IIf(IsNull(vlTB!Cod_ViaPago), "NULL", vlTB!Cod_ViaPago)
            stLiquidacionRet.Cod_TipCuenta = IIf(IsNull(vlTB!Cod_TipCuenta), "NULL", vlTB!Cod_TipCuenta)
            stLiquidacionRet.Cod_Banco = IIf(IsNull(vlTB!Cod_Banco), "NULL", vlTB!Cod_Banco)
            stLiquidacionRet.Num_Cuenta = IIf(IsNull(vlTB!Num_Cuenta), "NULL", vlTB!Num_Cuenta)
            stLiquidacionRet.Cod_Sucursal = IIf(IsNull(vlTB!Cod_Sucursal), "NULL", vlTB!Cod_Sucursal)
            stLiquidacionRet.Mto_Descuento = 0
            stLiquidacionRet.Mto_Haber = stLiquidacionRet.Mto_Haber + vlRetencion
            stLiquidacionRet.Mto_LiqPagar = stLiquidacionRet.Mto_LiqPagar '+ vlRetencion
            'stLiquidacionRet.Num_Cargas = 0 'stLiquidacionRet.Num_Cargas + vlNumCargasRet
            
            istLiquidacion.Mto_Descuento = istLiquidacion.Mto_Descuento + vlRetencion
                            
            'Traspasa Datos a la Estructura del Receptor
            If Not fgTraspasaDatosaEstructuraRetencion(istDetPension, stDetPensionRet, istLiquidacion, stLiquidacionRet) Then
                Exit Function
            End If
                                                
            'If stLiquidacionRet.Mto_LiqPagar <> 0 Then
                'Graba la última Liquidación
                'If stDatGenerales.Cod_ConceptoCobroRetencion <> "60" Then
                    If Not fgInsertaLiquidacion(stLiquidacionRet, iTipoTabla, iModulo) Then
                        Exit Function
                    End If
                'End If
            'End If
                                                                            
            'Graba Detalle del Receptor de la Retención Judicial
            If vlRetencion > 0 Then
                stDetPensionRet.Cod_ConHabDes = stDatGenerales.Cod_ConceptoCobroRetencion
                stDetPensionRet.Mto_ConHabDes = vlRetencion
                stDetPensionRet.Num_IdenReceptor = stLiquidacionRet.Num_IdenReceptor
                stDetPensionRet.Cod_TipoIdenReceptor = stLiquidacionRet.Cod_TipoIdenReceptor
                stDetPensionRet.Cod_TipReceptor = stLiquidacionRet.Cod_TipReceptor
                'If stDatGenerales.Cod_ConceptoCobroRetencion <> "60" Then
                    If Not fgInsertaDetallePensionProv(stDetPensionRet, iTipoTabla, iModulo) Then
                        MsgBox "Se ha producido un Error al Grabar la Retención Judicial" & Chr(13) & Err.Description, vbCritical, iCaption
                        Exit Function
                    End If
                'End If
            End If
        End If
        vlTB.MoveNext
    Loop
    'If stLiquidacionRet.Mto_LiqPagar <> 0 Then
    '    'Graba la última Liquidación
    '    If Not fgInsertaLiquidacion(stLiquidacionRet, iTipoTabla, iModulo) Then
    '        Exit Function
    '    End If
    'End If
    
    If vlTotalRetencion > 0 Then 'Graba Retención de la Renta
        istDetPension.Cod_ConHabDes = stDatGenerales.Cod_ConceptoRetencion
        istDetPension.Mto_ConHabDes = vlTotalRetencion
 
        'Graba Monto de la Retención Judicial
        If Not fgInsertaDetallePensionProv(istDetPension, iTipoTabla, iModulo) Then
            MsgBox "Se ha producido un Error al Grabar la Retención Judicial" & Chr(13) & Err.Description, vbCritical, iCaption
            Exit Function
        End If
    End If
End If
    
fgCalculaRetencion = True

Errores:
    If Err.Number <> 0 Then
        MsgBox "Se ha producido un error al Calcular la Retención Judicial" & Chr(13) & Err.Description, vbCritical, "Error de Datos"
    End If

End Function
Function fgBorraCalculosControlAnteriores(iModulo, iCaption) As Boolean
On Error GoTo Errores
Dim vlSql As String

fgBorraCalculosControlAnteriores = False

'Elimina Tabla de Detalle de Pagos
vlSql = "DELETE FROM PP_TTMP_CONPAGOPEN"
vlSql = vlSql & " WHERE COD_TIPOMOD = '" & iModulo & "'"
vgConexionTransac.Execute (vlSql)

'Elimina Tabla de Liquidacion
vlSql = "DELETE FROM PP_TTMP_CONLIQPAGOPEN"
vlSql = vlSql & " WHERE COD_TIPOMOD = '" & iModulo & "'"
vgConexionTransac.Execute (vlSql)

fgBorraCalculosControlAnteriores = True
Errores:
    If Err.Number <> 0 Then
        MsgBox "Se han producido errores al Eliminar los Datos de Cálculos Anteriores" & Chr(13) & Err.Description, vbCritical, iCaption
    End If
End Function
Function fgCalculaGarantiaEstatal(iTipPension As String, iPoliza As String, iEndoso As String, iOrden, iParentesco As String, iSitInv As String, iSexo As String, iPension As Double, iEdad As Integer, iEdadAños As Integer, iFecNac As String, iFecPago As String, iFecIniPag As Date, iFecTerPag As Date, oGarantia As Double, Optional iModulo As String, Optional iFecEfecto As Date) As Boolean
'Función que Calcula el Monto de la Garantía Estatal
Dim vlSql As String
Dim vlTB As ADODB.Recordset
Dim vlTB2 As ADODB.Recordset
Dim vlPensionMinima As Double
Dim vlPorDeduccion As Double
Dim vlFecTerGarEst As String
Dim vlPrimeraVez As Boolean 'Indica si es la Primera Vez que cae bajo la Mínima
Dim vlNumDias As Long 'Dias que corresponde Pagar
Dim vlDiasMes As Long 'Dias del Mes
Dim vlModulo As String
On Error GoTo Errores

'    If iPoliza = "0000020236" Then
'        MsgBox "POLIZA : " & iPoliza
'    End If
    
    vlModulo = ""
    If Not IsMissing(iModulo) Then
        vlModulo = iModulo
    End If
    fgCalculaGarantiaEstatal = True
    vlPensionMinima = 0
    vlPorDeduccion = 0
    vlPrimeraVez = False
    vlNumDias = 1 'Si no es proporcional la división daría 1
    vlDiasMes = 1
    If iTipPension = stDatGenerales.Cod_PensionAnticipada Then
        'Si es Vejez Anticipada solo se paga después de la Edad Legal
        If iSexo = "M" Then
            'Si es hombre desde los 65
            If iEdad < stDatGenerales.MesesEdad65 Then
                Exit Function
            ElseIf iEdad = stDatGenerales.MesesEdad65 Then
                'Si acaba de cumplir la Edad, se  paga Proporcional
                vlDiasMes = DateDiff("d", iFecIniPag, iFecTerPag) + 1
                vlNumDias = vlDiasMes - CLng(Mid(iFecNac, 7, 2)) + 1
            End If
        Else
            'Si es mujer desde los 60
            If iEdad < stDatGenerales.MesesEdad60 Then
                Exit Function
            ElseIf iEdad = stDatGenerales.MesesEdad60 Then
                'Si acaba de cumplir la Edad, se  paga Proporcional
                vlDiasMes = DateDiff("d", iFecIniPag, iFecTerPag) + 1
                vlNumDias = vlDiasMes - CLng(Mid(iFecNac, 7, 2)) + 1
            End If
        End If
    End If
     
    'Obtiene Monto de la Pensión Mínima
    If Not fgObtienePensionMinima(iFecPago, iParentesco, iSitInv, iSexo, iEdadAños, iPension, vlPensionMinima) Then
        Exit Function
    End If
'    vlSQL = "SELECT MTO_PENMINFIN FROM PP_TVAL_PENMINIMA"
'    vlSQL = vlSQL & " WHERE FEC_INIPENMIN <= '" & iFecPago & "'"
'    vlSQL = vlSQL & " AND COD_PAR = " & iParentesco
'    vlSQL = vlSQL & " AND COD_SITINV = '" & iSitInv & "'"
'    vlSQL = vlSQL & " AND COD_SEXO = '" & iSexo & "'"
'    vlSQL = vlSQL & " AND NUM_EDADINI <= " & iEdadAños
'    vlSQL = vlSQL & " AND NUM_EDADFIN >= " & iEdadAños
'    vlSQL = vlSQL & " AND FEC_TERPENMIN >= '" & iFecPago & "'"
'    vlSQL = vlSQL & " AND MTO_PENMINFIN > " & Str(iPension)
'    vlSQL = vlSQL & " ORDER BY MTO_PENMINFIN"
'    Set vlTB = vgConexionBD.Execute(vlSQL)
'    If Not vlTB.EOF Then
'        vlPensionMinima = vlTB!mto_penminfin 'Monto de la Pensión que debería ganar
'    Else
'        Exit Function
'    End If
    
    'Verificar si tiene Derecho a Garantía Estatal
    vlSql = "SELECT a.cod_dergarest, a.fec_iniestgarest, b.cod_indpago"
    vlSql = vlSql & " FROM pp_tmae_garestestado a, ma_tpar_estdergarest b"
    vlSql = vlSql & " WHERE a.cod_dergarest = b.cod_dergarest"
    vlSql = vlSql & " AND a.num_poliza = '" & iPoliza & "'"
    'vlSQL = vlSQL & " AND a.num_endoso = " & iEndoso
    vlSql = vlSql & " AND a.num_orden = " & iOrden
    If vlModulo <> "GEREL" Then 'Reliquidacion de Garantía Estatal
        'vlSQL = vlSQL & " AND a.fec_ingestgarest <= '" & Format(iFecIniPag, "yyyymmdd") & "'"
        vlSql = vlSql & " AND a.fec_efecto <= '" & Format(iFecIniPag, "yyyymmdd") & "'"
        vlSql = vlSql & " AND a.fec_terestgarest >= '" & Format(iFecIniPag, "yyyymmdd") & "'"
    Else
        vlSql = vlSql & " AND a.fec_iniestgarest <= '" & Format(iFecIniPag, "yyyymmdd") & "'"
        vlSql = vlSql & " AND a.fec_terestgarest >= '" & Format(iFecIniPag, "yyyymmdd") & "'"
    End If
    'vlSQL = vlSQL & " AND a.cod_dergarest <> 0"  '0 = Sin Derecho
    Set vlTB = vgConexionBD.Execute(vlSql)
    If Not vlTB.EOF Then
        If vlTB!cod_indpago = "N" Then 'No se Paga
            Exit Function
        End If
        If vlTB!COD_DERGAREST = 3 Then  'Con Documentación Pendiente
            'Verificar si lleva los 3 meses, si es así se deja sin Derecho
            If DateDiff("m", DateSerial(Mid(vlTB!fec_iniestgarest, 1, 4), Mid(vlTB!fec_iniestgarest, 5, 2), Mid(vlTB!fec_iniestgarest, 7, 2)), iFecIniPag) >= 3 Then
                'Se debe revisar si es la primera vez que está en estado Con Documentación Pendiente
                vlSql = "SELECT 1 FROM pp_tmae_garestestado"
                vlSql = vlSql & " WHERE num_poliza = '" & iPoliza & "'"
                vlSql = vlSql & " AND num_orden = " & iOrden
                vlSql = vlSql & " AND fec_iniestgarest < '" & vlTB!fec_iniestgarest & "'"
                vlSql = vlSql & " AND cod_dergarest = '" & vlTB!COD_DERGAREST & "'"
                If vgTipoBase = "ORACLE" Then
                    vlSql = vlSql & " AND ROWNUM = 1"
                End If
                Set vlTB2 = vgConexionBD.Execute(vlSql)
                If vlTB2.EOF Then
                    'Se debe revisar si se ingresó la Documentación (Solicitud y Certificados Civiles)
                    vlSql = "SELECT COUNT(1) AS CONT FROM pp_tmae_garestdoc a"
                    vlSql = vlSql & " WHERE a.num_poliza = '" & iPoliza & "'"
                    vlSql = vlSql & " AND a.num_orden = " & iOrden
                    vlSql = vlSql & " AND a.cod_docgarest IN ('01', '04')"
                    vlSql = vlSql & " AND a.fec_recdocgarest <= '" & Format(iFecIniPag, "yyyymmdd") & "'"
                    Set vlTB2 = vgConexionBD.Execute(vlSql)
                    If vlTB2.EOF Then
                        Exit Function 'Si no está la documentación no tiene derecho
                    End If
                    If vlTB2!cont <> 2 Then
                        Exit Function 'No están los dos documentos con fecha de recepción
                    End If
                'Else 'Se debe pagar
                End If
            'Else 'Se debe pagar
            End If
        End If
    Else
        'Cae por Primera vez bajo Garantía Estatal?
        vlSql = "SELECT 1 FROM PP_TMAE_GARESTESTADO"
        vlSql = vlSql & " WHERE NUM_POLIZA = '" & iPoliza & "'"
        vlSql = vlSql & " AND NUM_ORDEN = " & iOrden
        vlSql = vlSql & " AND FEC_INIESTGAREST <= '" & Format(iFecIniPag, "yyyymmdd") & "'"
        If vgTipoBase = "ORACLE" Then
            vlSql = vlSql & " AND ROWNUM = 1"
        End If
        Set vlTB = vgConexionBD.Execute(vlSql)
        If Not vlTB.EOF Then
            'Existen registros anteriores, por lo que no corresponde pagar GE
            Exit Function
        Else
            'Cae por Primera vez bajo la Mínima (Se deja con estado Pendiente hasta Fecha Indefinida y se Paga)
            vlSql = "INSERT INTO PP_TMAE_GARESTESTADO "
            vlSql = vlSql & " ( NUM_POLIZA, NUM_ENDOSO, NUM_ORDEN, FEC_INIESTGAREST, COD_DERGAREST,"
            vlSql = vlSql & " FEC_TERESTGAREST, COD_CAUSUSESTGAREST, FEC_INGESTGAREST, FEC_EFECTO ) VALUES ("
            vlSql = vlSql & " '" & iPoliza & "'," & iEndoso & "," & iOrden & ",'" & Format(iFecIniPag, "yyyymmdd") & "', 3, "
            'CMV-20060907 I
            'Se cambia valor Null por constante con valor 0 = Sin Información
            'vlSql = vlSql & "'" & vgTopeFecFin & "', NULL,"
            vlSql = vlSql & "'" & vgTopeFecFin & "', "
            vlSql = vlSql & "'" & vgCodCauSusEstGarEst0 & "', "
            'CMV-20060907 F
            vlSql = vlSql & "'" & iFecPago & "',"
            If vlModulo <> "GEREL" Then 'Reliquidacion de Garantía Estatal
                vlSql = vlSql & "'" & Format(iFecIniPag, "yyyymmdd") & "')" 'F
            Else
                vlSql = vlSql & "'" & Format(iFecEfecto, "yyyymmdd") & "')" 'F
            End If
            vgConexionBD.Execute (vlSql)
            vlPrimeraVez = True
        End If
    End If
''    HQR 27/11/2004 a petición de Verónica, no se valida resolución
''    If Not vlPrimeraVez Then
''        'Verificar si tiene Resolución Vigente
''        vlSQL = "SELECT 1 FROM PP_TMAE_GARESTRES"
''        vlSQL = vlSQL & " WHERE NUM_POLIZA = '" & iPoliza & "'"
''        'vlSQL = vlSQL & " AND NUM_ENDOSO = " & iEndoso
''        vlSQL = vlSQL & " AND NUM_ORDEN = " & iOrden
''        vlSQL = vlSQL & " AND FEC_INIRES <= '" & Format(vgFecIniPag, "yyyymmdd") & "'"
''        vlSQL = vlSQL & " AND FEC_TERRES >= '" & Format(vgFecIniPag, "yyyymmdd") & "'"
''        Set vlTB = vgConexionBD.Execute(vlSQL)
''        If vlTB.EOF Then 'Si no tiene Resolución Vigente, no se paga Garantía Estatal
''            Exit Function
''        End If
''    End If
    
    'Obtiene Porcentaje de Deducción
    vlSql = "SELECT PRC_DEDTOTAL FROM PP_TMAE_CALPORDED"
    vlSql = vlSql & " WHERE NUM_POLIZA = '" & iPoliza & "'"
    'vlSQL = vlSQL & " AND NUM_ENDOSO = " & iEndoso
    vlSql = vlSql & " AND NUM_ORDEN = " & iOrden
    'vlSQL = vlSQL & " AND FEC_PENSION <= '" & Format(iFecIniPag, "yyyymmdd") & "'"
    vlSql = vlSql & " AND FEC_INIPORDED <= '" & Format(iFecIniPag, "yyyymmdd") & "'"
    vlSql = vlSql & " AND FEC_TERPORDED >= '" & Format(iFecIniPag, "yyyymmdd") & "'"
    Set vlTB = vgConexionBD.Execute(vlSql)
    If Not vlTB.EOF Then
        vlPorDeduccion = vlTB!PRC_DEDTOTAL
    End If
    
    'Obtiene Número de Beneficiarios entre los que se debe repartir el porcentaje
    'Esto es solo para las Conyuges y Madres de Hijo Natural
    Dim vlParPenMin As Integer 'Parentesco para el que se debe obtener la pensión mínima
    Dim vlNumConyuges As Integer 'Número de Conyuges entre las que se debe repartir la Garantía Estatal
    
    If iParentesco >= 10 And iParentesco < 30 Then 'Conyuge o Madre de Hijo Natural
        vlParPenMin = iParentesco 'El mismo de entrada
        vlNumConyuges = 1
        vlSql = "SELECT COUNT(1) AS cont, ben.cod_par FROM pp_tmae_ben ben"
        vlSql = vlSql & " WHERE ben.num_poliza = '" & iPoliza & "'"
        vlSql = vlSql & " AND ben.num_endoso = " & iEndoso
        vlSql = vlSql & " AND ben.num_orden <> " & iOrden 'Que no sea el mismo beneficiario
        If iParentesco >= 10 And iParentesco < 20 Then 'Conyuge
            vlSql = vlSql & " AND ben.cod_par >= 10 and ben.cod_par < 20"
        Else
            vlSql = vlSql & " AND ben.cod_par >= 20 and ben.cod_par < 30" 'Madre de Hijo Natural
        End If
        'HQR 18/02/2006 Se agrega para que no divida la pensión cuando algún beneficiario no tiene derecho
        vlSql = vlSql & " AND ben.cod_estpension = 99" 'Solo los que tienen Derecho a Pensión
        vlSql = vlSql & " AND ben.fec_inipagopen <= '" & Format(iFecIniPag, "yyyymmdd") & "'" 'Solo los que ya iniciaron su pago de pensión o lo inician en este periodo
        'FIN HQR 18/02/2006
        vlSql = vlSql & " GROUP BY ben.cod_par"
        Set vlTB = vgConexionBD.Execute(vlSql)
        If Not vlTB.EOF Then
            vlTB.MoveFirst
            Do While Not vlTB.EOF
                vlNumConyuges = vlNumConyuges + vlTB!cont
                If iParentesco = 11 Then 'Ver si existe alguna conyuge sin hijos para buscar la pensión mínima para este parentesco
                    If vlTB!Cod_Par = 10 Then
                        vlParPenMin = vlTB!Cod_Par
                    End If
                ElseIf iParentesco = 21 Then 'Ver si existe alguna conyuge con hijos
                    If vlTB!Cod_Par = 20 Then
                        vlParPenMin = vlTB!Cod_Par
                    End If
                End If
                vlTB.MoveNext
            Loop
            'Si el parentesco es distinto se debe volver a obtener el Monto de la Pensión Mínima
            If vlParPenMin <> iParentesco Then
                'Obtiene Monto de la Pensión Mínima
                vlSql = "SELECT MTO_PENMINFIN FROM PP_TVAL_PENMINIMA"
                vlSql = vlSql & " WHERE FEC_INIPENMIN <= '" & iFecPago & "'"
                vlSql = vlSql & " AND COD_PAR = " & vlParPenMin
                vlSql = vlSql & " AND COD_SITINV = '" & iSitInv & "'"
                vlSql = vlSql & " AND COD_SEXO = '" & iSexo & "'"
                vlSql = vlSql & " AND NUM_EDADINI <= " & iEdadAños
                vlSql = vlSql & " AND NUM_EDADFIN >= " & iEdadAños
                vlSql = vlSql & " AND FEC_TERPENMIN >= '" & iFecPago & "'"
                vlSql = vlSql & " AND MTO_PENMINFIN > " & str(iPension)
                vlSql = vlSql & " ORDER BY MTO_PENMINFIN"
                Set vlTB = vgConexionBD.Execute(vlSql)
                If Not vlTB.EOF Then
                    vlPensionMinima = vlTB!mto_penminfin 'Monto de la Pensión que debería ganar
                Else
                    Exit Function
                End If
            End If
            'La pensión Mínima se divide por el Número de Conyuges
            vlPensionMinima = Format(vlPensionMinima / vlNumConyuges, "##0.00") 'Con 2 decimales
        End If
    End If
    If vlPrimeraVez Then
        'Debe calcular monto en Forma Proporcional
        vlDiasMes = DateDiff("d", iFecIniPag, iFecTerPag) + 1
        vlNumDias = DateDiff("d", iFecIniPag, iFecTerPag) + 1
        oGarantia = Format((vlPensionMinima - iPension) * (vlNumDias / vlDiasMes) * (1 - vlPorDeduccion), "##0")
    Else
        'Paga el Mes completo
        oGarantia = Format((vlPensionMinima - iPension) * (1 - (vlPorDeduccion / 100)), "##0")
    End If
    
Errores:
    If Err.Number <> 0 Then
        MsgBox "Se ha producido un error al Calcular la Retención Judicial o de Asignación Familiar" & Chr(13) & Err.Description, vbCritical, "Error de Datos"
        fgCalculaGarantiaEstatal = False
    End If
End Function


Function fgObtieneDatosPrimerPago(iPeriodo, oFecPago, oValUF, oFecIniPag, oFecTerPag, iCaption, oValUFUltDia As Double, oPrcMinSalud As Double, oMtoMaxSalud As Double, oMtoTopeImp As Double) As Boolean
'Obtiene Datos del Primer Pago
Dim vlSql As String
Dim vlUF As Double
Dim vlFecPago As String

fgObtieneDatosPrimerPago = False
vlUF = 0
vlSql = "SELECT * FROM PP_TMAE_PROPAGOPEN A"
vlSql = vlSql & " WHERE A.NUM_PERPAGO = '" & iPeriodo & "'"
Set vgRs = vgConexionBD.Execute(vlSql)
If Not vgRs.EOF Then
    vlFecPago = DateSerial(Mid(vgRs!fec_calpripago, 1, 4), Mid(vgRs!fec_calpripago, 5, 2), Mid(vgRs!fec_calpripago, 7, 2))
    If Not fgObtieneConversion(vgRs!fec_calpripago, "UF", oValUF) Then
        MsgBox "Debe ingresar Valor UF a la Fecha de Pago : " & vlFecPago, vbCritical, iCaption
        Exit Function
    End If
    oFecIniPag = DateSerial(Mid(vgRs!Num_PerPago, 1, 4), Mid(vgRs!Num_PerPago, 5, 2), 1)
    oFecTerPag = DateAdd("d", -1, DateAdd("m", 1, oFecIniPag))
    oFecPago = vgRs!Fec_PriPago
Else
    MsgBox "Debe Ingresar en Calendario el Periodo de los Primeros Pagos que se va a Calcular: " & Mid(iPeriodo, 5, 2) & "-" & Mid(iPeriodo, 1, 4), vbCritical, iCaption
    Exit Function
End If
'Valor UF al último día del mes
Dim vlUltDiaMesFecPago As Date
vlUltDiaMesFecPago = DateSerial(Mid(oFecPago, 1, 4), Mid(oFecPago, 5, 2) + 1, 0)
If Not fgObtieneConversion(Format(vlUltDiaMesFecPago, "yyyymmdd"), "UF", oValUFUltDia) Then
    MsgBox "Debe ingresar Valor UF al Último Día del Mes : " & vlUltDiaMesFecPago, vbCritical, iCaption
    Exit Function
End If

'Obtiene Porcentaje Minimo de Salud
If Not fgObtieneParametroVigencia("PS", "PSM", vgRs!fec_calpripago, oPrcMinSalud) Then
    MsgBox "Debe ingresar Porcentaje Mínimo de Salud", vbCritical, iCaption
    Exit Function
End If

'Obtiene Monto Máximo de Salud a Cobrar
If Not fgObtieneParametroVigencia("MS", "MSM", vgRs!fec_calpripago, oMtoMaxSalud) Then
    MsgBox "Debe ingresar Monto Máximo de Salud en UF", vbCritical, iCaption
    Exit Function
End If

'Obtiene Tope Base Imponible (Para calculo de Aporte por CCAF)
If Not fgObtieneParametroVigencia("TBI", "MBM", vgRs!fec_calpripago, oMtoTopeImp) Then
    MsgBox "Debe ingresar Monto Tope Base Imponible en UF", vbCritical, iCaption
    Exit Function
End If

fgObtieneDatosPrimerPago = True
End Function
Function fgObtieneDatosPagoRegimen(iPeriodo, oFecPago, oValUF, oFecIniPag, oFecTerPag, iCaption, oValUFUltDia As Double, oPrcMinSalud As Double, oMtoMaxSalud As Double, oMtoTopeImp As Double) As Boolean
'Obtiene Datos del Pago en Régimen
Dim vlSql As String
Dim vlUF As Double
Dim vlFecPago As String

fgObtieneDatosPagoRegimen = False
vlUF = 0
vlSql = "SELECT * FROM PP_TMAE_PROPAGOPEN A"
vlSql = vlSql & " WHERE A.NUM_PERPAGO = '" & iPeriodo & "'"
Set vgRs = vgConexionBD.Execute(vlSql)
If Not vgRs.EOF Then
    vlFecPago = DateSerial(Mid(vgRs!fec_calpagoreg, 1, 4), Mid(vgRs!fec_calpagoreg, 5, 2), Mid(vgRs!fec_calpagoreg, 7, 2) - 1)
    'hqr 12/12/2007 Se comenta a peticion de MCHirinos
    oValUF = 1
'    If Not fgObtieneConversion(Format(vlFecPago, "yyyymmdd"), "US", oValUF) Then
'        MsgBox "Debe ingresar Valor US a la Fecha de Pago : " & vlFecPago, vbCritical, iCaption
'        Exit Function
'    End If
    'fin hqr 12/12/2007 Se comenta a peticion de MCHirinos
    oFecIniPag = DateSerial(Mid(vgRs!Num_PerPago, 1, 4), Mid(vgRs!Num_PerPago, 5, 2), 1)
    oFecTerPag = DateAdd("d", -1, DateAdd("m", 1, oFecIniPag))
    oFecPago = vgRs!fec_calpagoreg
Else
    MsgBox "Debe Ingresar en Calendario el Periodo de los Pagos Recurrentes que se va a Calcular: " & Mid(iPeriodo, 5, 2) & "-" & Mid(iPeriodo, 1, 4), vbCritical, iCaption
    Exit Function
End If

'Valor UF al último día del mes
''Dim vlUltDiaMesFecPago As Date
''vlUltDiaMesFecPago = DateSerial(Mid(oFecPago, 1, 4), Mid(oFecPago, 5, 2) + 1, 0)
''If Not fgObtieneConversion(Format(vlUltDiaMesFecPago, "yyyymmdd"), "US", oValUFUltDia) Then
''    MsgBox "Debe ingresar Valor US al Último Día del Mes : " & vlUltDiaMesFecPago, vbCritical, iCaption
''    Exit Function
''End If

'Obtiene Porcentaje Minimo de Salud
If Not fgObtieneParametroVigencia("PS", "PSM", vgRs!fec_calpagoreg, oPrcMinSalud) Then
    MsgBox "Debe ingresar Porcentaje Mínimo de Salud", vbCritical, iCaption
    Exit Function
End If

''Obtiene Monto Máximo de Salud a Cobrar
'If Not fgObtieneParametroVigencia("MS", "MSM", vgRs!fec_calpagoreg, oMtoMaxSalud) Then
'    MsgBox "Debe ingresar Monto Máximo de Salud en UF", vbCritical, iCaption
'    Exit Function
'End If
    
fgObtieneDatosPagoRegimen = True
End Function

Function fgObtieneDescripcionConcepto(iCodConcepto) As String
    'Obtiene la Descripción de un Concepto Específico
    
    Dim vlSql As String
    Dim vlTB As ADODB.Recordset
    
    fgObtieneDescripcionConcepto = ""
    vlSql = "SELECT a.cod_conhabdes, a.gls_conhabdes FROM ma_tpar_conhabdes a"
    vlSql = vlSql & " WHERE a.cod_conhabdes = '" & iCodConcepto & "'"
    Set vlTB = vgConexionBD.Execute(vlSql)
    If Not vlTB.EOF Then
        fgObtieneDescripcionConcepto = " " & vlTB!Cod_ConHabDes & " - " & vlTB!gls_ConHabDes
    End If

End Function

Function fgObtieneMontoConcepto(iPoliza As String, iOrden As Long, iOrdenCar As Long, iConcepto As String, iPeriodo As String, iModulo) As Double
'Obtiene el Monto pagado por un Concepto y Periodo
Dim vlSql As String
Dim vlTB As ADODB.Recordset

    fgObtieneMontoConcepto = 0
    If iModulo = "AF" Then
        'Monto de las Cargas Familiares
        vlSql = "SELECT (p.mto_carga * p.num_cargas) as monto FROM pp_tmae_pagoasigdef p"
        vlSql = vlSql & " WHERE p.num_perpago = '" & iPeriodo & "'"
        vlSql = vlSql & " AND p.num_poliza = '" & iPoliza & "'"
        vlSql = vlSql & " AND p.num_orden = " & iOrden
        vlSql = vlSql & " AND p.cod_tipreceptor <> 'R'"
        vlSql = vlSql & " AND p.num_ordencar = " & iOrdenCar
    Else
        vlSql = "SELECT p.mto_conhabdes as monto FROM pp_tmae_pagopendef p"
        vlSql = vlSql & " WHERE p.num_perpago = '" & iPeriodo & "'"
        vlSql = vlSql & " AND p.num_poliza = '" & iPoliza & "'"
        vlSql = vlSql & " AND p.num_orden = " & iOrden
        vlSql = vlSql & " AND p.cod_tipreceptor <> 'R'"
        vlSql = vlSql & " AND p.cod_conhabdes = '" & iConcepto & "'"
    End If
    Set vlTB = vgConexionBD.Execute(vlSql)
    If Not vlTB.EOF Then
        fgObtieneMontoConcepto = vlTB!monto
    End If
End Function
Function fgObtieneMontoTotalConcepto(iPeriodo As String, iPoliza As String, iOrden As Long, iConcepto As String, iFecIniPag, oNumCuotas As Long) As Double
'Obtiene el Monto pagado por un Concepto y Periodo
Dim vlSql As String
Dim vlTB As ADODB.Recordset

    fgObtieneMontoTotalConcepto = 0
    oNumCuotas = 0
    vlSql = "SELECT SUM(p.mto_conhabdes) as monto, COUNT(1) as num_cuotas FROM pp_tmae_pagopendef p"
    vlSql = vlSql & " WHERE p.num_perpago <> '" & iPeriodo & "'" 'Para que busque por la llave y si se está repitiendo el cálculo de un periodo no se considere el monto
    vlSql = vlSql & " AND p.num_poliza = '" & iPoliza & "'"
    vlSql = vlSql & " AND p.num_orden = " & iOrden
    'vlSql = vlSql & " AND p.rut_receptor > 0" 'Para que busque por la llave
    vlSql = vlSql & " AND p.cod_tipreceptor <> 'R'"
    vlSql = vlSql & " AND p.cod_conhabdes = '" & iConcepto & "'"
    vlSql = vlSql & " AND p.fec_inipago >= '" & iFecIniPag & "'" 'Para que se trate del mismo concepto y no de un pago ingresado anteriormente
    Set vlTB = vgConexionBD.Execute(vlSql)
    If Not vlTB.EOF Then
        oNumCuotas = IIf(IsNull(vlTB!Num_Cuotas), 0, vlTB!Num_Cuotas)
        fgObtieneMontoTotalConcepto = IIf(IsNull(vlTB!monto), 0, vlTB!monto)
    End If
End Function
Function fgObtieneParametrosQuiebra(iPeriodo As String, oPrcCastigo As Double, oTopeMax As Double) As Boolean
    'Obtiene Parámetros de Quiebra
    Dim vlSql As String
    Dim vlTB As ADODB.Recordset
    On Error Resume Next 'Si se produce un Error se asume que no hay quiebra
    
    fgObtieneParametrosQuiebra = False
    vlSql = "SELECT quie.prc_castigo, quie.mto_topemax"
    vlSql = vlSql & " FROM pp_tpar_quiebra quie"
    vlSql = vlSql & " WHERE quie.num_perini <= '" & iPeriodo & "'"
    vlSql = vlSql & " AND quie.num_perter >= '" & iPeriodo & "'"
    Set vlTB = vgConexionBD.Execute(vlSql)
    If Not vlTB.EOF Then
        oPrcCastigo = vlTB!Prc_Castigo
        oTopeMax = vlTB!mto_topemax
        fgObtieneParametrosQuiebra = True
    End If
    vlTB.Close
End Function
Public Function fgFechaEfectoReliq(iFecha As String, iNumPoliza As String, iNumOrden As Integer, iFecTermino As String) As String
    Dim vlFechaEfecto As String
    
    fgFechaEfectoReliq = ""
    vlFechaEfecto = fgValidaFechaEfecto(iFecha, iNumPoliza, iNumOrden)
    If IsDate(iFecTermino) Then
        If CDate(vlFechaEfecto) > CDate(iFecTermino) Then
            vlFechaEfecto = DateAdd("d", 1, iFecTermino)
        End If
    End If
    fgFechaEfectoReliq = vlFechaEfecto
End Function
Function fgObtieneFactorAjuste(iFecha, oValor) As Boolean
'Obtiene Factor IPC
On Error GoTo Errores
Dim vlSql As String
Dim vlTB As ADODB.Recordset

fgObtieneFactorAjuste = False
vlSql = "SELECT a.mto_ipc"
vlSql = vlSql & " FROM ma_tval_ipc a"
vlSql = vlSql & " WHERE a.fec_ipc = '" & iFecha & "'"
Set vlTB = vgConexionBD.Execute(vlSql)
If Not vlTB.EOF Then
    oValor = vlTB!mto_ipc
Else
    vlTB.Close
    MsgBox "No existe Factor de Variación para la Fecha : " & DateSerial(Mid(iFecha, 1, 4), Mid(iFecha, 5, 2), Mid(iFecha, 7, 2)), vbCritical, "Faltan Datos"
    Exit Function
End If
vlTB.Close
fgObtieneFactorAjuste = True

Errores:
If Err.Number <> 0 Then
    MsgBox "Se ha producido un Error al Obtener Factor de Variación Pensión" & Chr(13) & Err.Description, vbCritical, "Error al Obtener Datos Generales"
End If

End Function

Function fgObtieneFactorAjusteIPCSolInd(iFecha, bValQ, oValor) As Boolean
'Obtiene Factor IPC
On Error GoTo Errores
Dim vlSql As String
Dim vlTB As ADODB.Recordset

fgObtieneFactorAjusteIPCSolInd = False

If bValQ = False Then
    vlSql = "SELECT MTO_FACTOR FROM MA_TVAL_VALVAC WHERE '" & iFecha & "' BETWEEN FEC_INICUOMOR AND FEC_TERCUOMOR"
    Set vlTB = vgConexionBD.Execute(vlSql)
    If Not vlTB.EOF Then
            oValor = vlTB!MTO_FACTOR
    Else
        vlTB.Close
        MsgBox "No existe Factor de Variación para la Fecha : " & DateSerial(Mid(iFecha, 1, 4), Mid(iFecha, 5, 2), Mid(iFecha, 7, 2)), vbCritical, "Faltan Datos"
        Exit Function
    End If
Else
    vlSql = "SELECT MTO_IPCMEN FROM MA_TVAL_IPCVAC WHERE FEC_VIGIPC='" & iFecha & "'"
    Set vlTB = vgConexionBD.Execute(vlSql)
    If Not vlTB.EOF Then
            oValor = vlTB!MTO_IPCMEN
    Else
        vlTB.Close
        MsgBox "No existe Factor de Variación para la Fecha : " & DateSerial(Mid(iFecha, 1, 4), Mid(iFecha, 5, 2), Mid(iFecha, 7, 2)), vbCritical, "Faltan Datos"
        Exit Function
    End If
End If

'vlTb.Close
fgObtieneFactorAjusteIPCSolInd = True

Errores:
If Err.Number <> 0 Then
    MsgBox "Se ha producido un Error al Obtener Factor de Variación Pensión" & Chr(13) & Err.Description, vbCritical, "Error al Obtener Datos Generales"
End If

End Function



Function fgConvierteDigito(iDigito As Double) As String
'Convierte Monto en palabras
Dim vlMontoPalabras As String
Select Case iDigito
    Case 1
        vlMontoPalabras = "uno"
    Case 2
        vlMontoPalabras = "dos"
    Case 3
        vlMontoPalabras = "tres"
    Case 4
        vlMontoPalabras = "cuatro"
    Case 5
        vlMontoPalabras = "cinco"
    Case 6
        vlMontoPalabras = "seis"
    Case 7
        vlMontoPalabras = "siete"
    Case 8
        vlMontoPalabras = "ocho"
    Case 9
        vlMontoPalabras = "nueve"
    Case 10
        vlMontoPalabras = "diez"
    Case 11
        vlMontoPalabras = "once"
    Case 12
        vlMontoPalabras = "doce"
    Case 13
        vlMontoPalabras = "trece"
    Case 14
        vlMontoPalabras = "catorce"
    Case 15
        vlMontoPalabras = "quince"
    Case 16
        vlMontoPalabras = "dieciseis"
    Case 17
        vlMontoPalabras = "diecisiete"
    Case 18
        vlMontoPalabras = "dieciocho"
    Case 19
        vlMontoPalabras = "diecinueve"
    Case 20
        vlMontoPalabras = "veinte"
End Select
fgConvierteDigito = vlMontoPalabras
End Function


Function fgConvierteDecenas(iMonto As Double) As String
'Convierte Monto en palabras
Dim vlMontoPalabras As String
Dim vlDigito As Double
vlDigito = iMonto Mod 10
Select Case iMonto
    Case Is > 90
        vlMontoPalabras = "noventa y " + fgConvierteDigito(vlDigito)
    Case 90
        vlMontoPalabras = "noventa"
    Case Is > 80
        vlMontoPalabras = "ochenta y " + fgConvierteDigito(vlDigito)
    Case 80
        vlMontoPalabras = "ochenta"
    Case Is > 70
        vlMontoPalabras = "setenta y " + fgConvierteDigito(vlDigito)
    Case 70
        vlMontoPalabras = "setenta"
    Case Is > 60
        vlMontoPalabras = "sesenta y " + fgConvierteDigito(vlDigito)
    Case 60
        vlMontoPalabras = "sesenta"
    Case Is > 50
        vlMontoPalabras = "cincuenta y " + fgConvierteDigito(vlDigito)
    Case 50
        vlMontoPalabras = "cincuenta"
    Case Is > 40
        vlMontoPalabras = "cuarenta y " + fgConvierteDigito(vlDigito)
    Case 40
        vlMontoPalabras = "cuarenta"
    Case Is > 30
        vlMontoPalabras = "treinta y " + fgConvierteDigito(vlDigito)
    Case 30
        vlMontoPalabras = "treinta"
    Case Is > 20
        vlMontoPalabras = "veinti" + fgConvierteDigito(vlDigito)
    Case Is <= 20
        vlMontoPalabras = fgConvierteDigito(iMonto)
End Select
fgConvierteDecenas = vlMontoPalabras
End Function

Function fgConvierteCentenas(iMonto As Double) As String
'Convierte Monto en palabras
Dim vlMontoPalabras As String
Dim vlDecena As Double
vlDecena = iMonto Mod 100
Select Case iMonto
    Case Is > 900
        vlMontoPalabras = "novecientos " + fgConvierteDecenas(vlDecena)
    Case 900
        vlMontoPalabras = "novecientos"
    Case Is > 800
        vlMontoPalabras = "ochocientos " + fgConvierteDecenas(vlDecena)
    Case 800
        vlMontoPalabras = "ochocientos"
    Case Is > 700
        vlMontoPalabras = "setecientos " + fgConvierteDecenas(vlDecena)
    Case 700
        vlMontoPalabras = "setecientos"
    Case Is > 600
        vlMontoPalabras = "seiscientos " + fgConvierteDecenas(vlDecena)
    Case 600
        vlMontoPalabras = "seiscientos"
    Case Is > 500
        vlMontoPalabras = "quinientos " + fgConvierteDecenas(vlDecena)
    Case 500
        vlMontoPalabras = "quinientos"
    Case Is > 400
        vlMontoPalabras = "cuatrocientos " + fgConvierteDecenas(vlDecena)
    Case 400
        vlMontoPalabras = "cuatrocientos"
    Case Is > 300
        vlMontoPalabras = "trescientos " + fgConvierteDecenas(vlDecena)
    Case 300
        vlMontoPalabras = "trescientos"
    Case Is > 200
        vlMontoPalabras = "doscientos " + fgConvierteDecenas(vlDecena)
    Case 200
        vlMontoPalabras = "doscientos"
    Case Is > 100
        vlMontoPalabras = "ciento " + fgConvierteDecenas(vlDecena)
    Case 100
        vlMontoPalabras = "cien"
    Case Is < 100
        vlMontoPalabras = fgConvierteDecenas(iMonto)
End Select
fgConvierteCentenas = vlMontoPalabras
End Function


Function fgConvierteMiles(iMonto As Double) As String
'Convierte Monto en palabras
Dim vlMontoPalabras As String
Dim vlFraccion As Double
Dim vlCantidad As Double
vlFraccion = iMonto Mod 1000
vlCantidad = Int(iMonto / 1000) 'Parte Entera
Select Case iMonto
    Case 1000
        vlMontoPalabras = "mil"
    Case Is > 999
        vlMontoPalabras = Trim(fgConvierteCentenas(vlCantidad) + " mil " + fgConvierteCentenas(vlFraccion))
    Case Else
        vlMontoPalabras = fgConvierteCentenas(iMonto)
End Select
fgConvierteMiles = vlMontoPalabras
End Function

Function fgConvierteMillones(iMonto As Double) As String
'Convierte Monto en palabras
Dim vlMontoPalabras As String
Dim vlFraccion As Double
Dim vlCantidad As Double
vlFraccion = iMonto Mod 1000000
vlCantidad = Int(iMonto / 1000000) 'Parte Entera
Select Case iMonto
    Case 1000000
        vlMontoPalabras = "un millón"
    Case Is > 999999
        vlMontoPalabras = Trim(fgConvierteMiles(vlCantidad) + " millones " + fgConvierteMiles(vlFraccion))
    Case Else
        vlMontoPalabras = fgConvierteMiles(iMonto)
End Select
fgConvierteMillones = vlMontoPalabras
End Function

Function fgConvierteNumeroLetras(iMonto As Double, Optional iMoneda As String) As String
'Convierte Monto en palabras
Dim vlMontoPalabras As String
Dim vlDecimales As Double
Dim vlEntero As Double
vlEntero = Fix(iMonto)
vlDecimales = Format((iMonto - vlEntero) * 100, "#0.00")
vlMontoPalabras = fgConvierteMillones(vlEntero)
If iMonto > 2 Then
    If Mid(vlMontoPalabras, 1, 3) = "uno" Then
        vlMontoPalabras = Mid(vlMontoPalabras, 1, 2) + Mid(vlMontoPalabras, 4)
    End If
End If
If vlDecimales > 0 Then
    vlMontoPalabras = vlMontoPalabras + " con " & vlDecimales & "/100"
End If
If Not IsMissing(iMoneda) Then
    vlMontoPalabras = vlMontoPalabras + " " & iMoneda
End If
fgConvierteNumeroLetras = UCase(vlMontoPalabras)

End Function

Function fgObtieneFactorAjusteTasaFija(iFecVigencia As String, iFecIniPag As Date, iFactorAjusteTri As Double, iFactorAjusteMen As Double, iFactorAjusteTasaFija As Double, iFecDevengamiento As String) As Double
    Dim vlFactorAjuste As Double
    'hqr  03/03/2011
    Dim vlFecDesdeAjustePension As String

    If clAjusteDesdeFechaDevengamiento Then 'hqr  03/03/2011
        vlFecDesdeAjustePension = iFecDevengamiento 'Fecha de Devengamiento
    Else
        vlFecDesdeAjustePension = iFecVigencia 'Fecha de Inicio de Vigencia de la Póliza
    End If
    'fin hqr  03/03/2011
    
    vlFactorAjuste = 1
    'Calcula Factor de Ajuste
    If Mid(vlFecDesdeAjustePension, 1, 6) <= Format(DateAdd("m", -3, iFecIniPag), "yyyymm") Then 'Si se pagaron 3 meses se debe ajustar con la tasa trimestral
        vlFactorAjuste = 1 + (iFactorAjusteTasaFija * (iFactorAjusteTri / 100))
    ElseIf Mid(vlFecDesdeAjustePension, 1, 6) = Format(iFecIniPag, "yyyymm") Then ' Primer mes, no se ajusta
        vlFactorAjuste = 1
    ElseIf Mid(vlFecDesdeAjustePension, 1, 6) = Format(DateAdd("m", -1, iFecIniPag), "yyyymm") Then ' Segundo mes, siempre se aplica tasa mensual
        vlFactorAjuste = 1 + (iFactorAjusteMen / 100)
    ElseIf Mid(vlFecDesdeAjustePension, 1, 6) = Format(DateAdd("m", -2, iFecIniPag), "yyyymm") And (Mid(vlFecDesdeAjustePension, 5, 2) Mod 3 <> 0) Then ' tercer mes, se aplica tasa mensual, solo si el inicio de vigencia no fue en marzo, junio, sept, diciembre
        vlFactorAjuste = 1 + (iFactorAjusteMen / 100)
    Else
        vlFactorAjuste = 1 + (iFactorAjusteTasaFija * (iFactorAjusteTri / 100)) 'acá no debería entrar
    End If
    fgObtieneFactorAjusteTasaFija = vlFactorAjuste
    
End Function




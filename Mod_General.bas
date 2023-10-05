Attribute VB_Name = "Mod_General"
Option Explicit

'Definición de Constantes
Global Const vgCodTabla_AFP = "AF"         'Administradora de Fondos de Pensiones
Global Const vgCodTabla_AltPen = "AL"      'Alternativas de Pensión
Global Const vgCodTabla_Bco = "BCO"        'Entidades Bancarias
Global Const vgCodTabla_CCAF = "CC"        'Caja de Compensación
Global Const vgCodTabla_CauEnd = "CE"      'Causa de Generación de Endosos
Global Const vgCodTabla_CobPol = "CO"      'Cobertura de Pólizas (Trad)
Global Const vgCodTabla_ConPagTer = "COP"  'Concepto de Pago a Terceros
Global Const vgCodTabla_ComRea = "CR"      'Compañías de Reaseguro
Global Const vgCodTabla_CauSupAsiFam = "CSA" 'Causa Suspensión de la Asignación Familiar
Global Const vgCodTabla_CauSusGarEst = "CSG" 'Causa Suspensión de la Garantía Estatal
Global Const vgCodTabla_DerAcre = "DC"     'Derecho a Acrecer
Global Const vgCodTabla_DerPen = "DE"      'Derecho a Pensión
Global Const vgCodTabla_DerGarEst = "DEG"  'Derecho a Garantía Estatal
Global Const vgCodTabla_EstCiv = "EC"      'Estado Civil
Global Const vgCodTabla_EstPol = "EP"      'Estado de la Póliza
Global Const vgCodTabla_EstVigAsiFam = "EV" 'Estado de Vigencia de la Asignación Familiar
Global Const vgCodTabla_FacCas = "FC"      'Porcentaje Disminución Pensión por Quiebra Cía.
Global Const vgCodTabla_FrePago = "FP"     'Forma o Frecuencia de Pago
Global Const vgCodTabla_FacQui = "FQ"      'Porc. Pensión Garantizada por el Estado
Global Const vgCodTabla_GarEst = "GE"      'Estado de Garantía Estatal
Global Const vgCodTabla_GruFam = "GF"      'Grupo Familiar
Global Const vgCodTabla_InsSal = "IS"      'Institución de Salud
Global Const vgCodTabla_LimEdad = "LI"     'Límite de Edad para la Tabla de Mortalidad
Global Const vgCodTabla_ModRen = "MO"      'Modalidad de Rentabilidad
Global Const vgCodTabla_ModOriHabDes = "MOR" 'Modalidad de Origen del Haber o Descuento
Global Const vgCodTabla_MovHabDes = "MOV"  'Tipo de Movimiento del Haber o Descuento
Global Const vgCodTabla_ModPago = "MP"     'Modalidad de Pago
Global Const vgCodTabla_ModPagoRetJud = "MPR"     'Modalidad de Pago de Retención Judicial
Global Const vgCodTabla_ModRea = "MR"      'Modalidad de Reaseguro
Global Const vgCodTabla_ModVejAnt = "MVA"  'Modalidad de Vejez Anticipada para Gar.Est.
Global Const vgCodTabla_OpeRea = "OR"      'Operación de Reaseguro
Global Const vgCodTabla_Par = "PA"         'Parentesco
Global Const vgCodTabla_ParNoBen = "PNB"   'Estado de Parentesco de los No Beneficiarios
Global Const vgCodTabla_Plan = "PL"        'Plan (Trad)
Global Const vgCodTabla_ReqPen = "RPE"     'Requisitos de Pensión
Global Const vgCodTabla_SitCor = "SC"      'Situación del Corredor
Global Const vgCodTabla_Sexo = "SE"        'Sexo
Global Const vgCodTabla_SitInv = "SI"      'Situación de Invalidez
Global Const vgCodTabla_SitHabDes = "SHD"  'Suspensión de Haberes y Descuentos
Global Const vgCodTabla_TipCor = "TC"      'Tipo de Corredor
Global Const vgCodTabla_TipCta = "TCT"     'Tipo de Cuenta de Depósito
Global Const vgCodTabla_ModTipCta = "TCC"     'Modalidad Tipo de Cuenta de Dep?sito
Global Const vgCodTabla_TipDoc = "TD"      'Tipo de Documento
Global Const vgCodTabla_TipPagTem = "TE"   'Tipo de Pago de Pensiones Temporales
Global Const vgCodTabla_TipEnd = "TEN"     'Tipo de Endoso
Global Const vgCodTabla_TipPri = "TI"      'Tipo de Prima
Global Const vgCodTabla_TipIngMen = "TIM"  'Tipo de Ingreso de Mensaje
Global Const vgCodTabla_TipMon = "TM"      'Tipo de Moneda
Global Const vgCodTabla_TipPen = "TP"      'Tipo de Pensión
Global Const vgCodTabla_TipPer = "TPE"     'Tipo de Persona
Global Const vgCodTabla_TipRen = "TR"      'Tipo de Rentabilidad
Global Const vgCodTabla_TipReso = "TRE"    'Tipo de Resolución
Global Const vgCodTabla_TipRetJud = "TRT"  'Tipo de Retención Judicial
Global Const vgCodTabla_TipVej = "TV"      'Tipo de Vejez
Global Const vgCodTabla_TipVig = "VI"      'Estado de Vigencia
Global Const vgCodTabla_TipVigPol = "VP"   'Estado de Vigencia de la Póliza
Global Const vgCodTabla_ViaPago = "VPG"    'Vía de Pago de la Pensión
'----- KVR CCAF -----
Global Const vgCodTabla_SusCCAF = "CSC"     'Causa de suspensión CCAF
Global Const vgCodTabla_ModPagoCC = "MPC"   'Modalidad de Pago de la CCAF
'CMV
Global Const vgCodTabla_PrcSal = "PS"      'Porcentaje de Salud Minima


Global Const vgTipoSistema = "PP"
Global Const vgNombreSistema = "Sistema Previsional"
Global Const vgTopeFecFin = "99991231"
Global Const vgTopeMtoCarFam = "999999999"
Global Const vgTopeEdadMaxima = "150"
Global Const vgTopeEdadMinima = "0"

Global vgConexionBD As ADODB.Connection
Global vgConectarBD As ADODB.Connection

Global vgRs         As ADODB.Recordset
Global vgRs2        As ADODB.Recordset
Global vgCmb        As ADODB.Recordset
Global vgRegistro   As ADODB.Recordset
Global vgRs3        As ADODB.Recordset
Global vgRs4        As ADODB.Recordset

Global vgSql    As String
Global vgQuery  As String
Global vgGra    As String

'Variables utilizadas en la Conexión al Sistema
Global vgMensaje As String
Global vgNombreServidor As String
Global vgNombreBaseDatos As String
Global vgNombreUsuario As String
Global vgPassWord As String
Global vgUsuario As String
Global vgRutUsuario As String
Global vgDsn As String
Global vgRutaDataBase As String
Global vgRutaBasedeDatos As String
Global vgRutaBasedeDatos_Aux As String
Global vgRutaArchivo  As String

'Nombres del Cliente o la Compañía adquisidora
'I--- ABV 04/12/2006 ---
'Global vgRutCompania         As String
'Global vgDgVCompania         As String
Global vgNumIdenCompania     As String
Global vgTipoIdenCompania    As Long
'F--- ABV 04/12/2006 ---
Global vgNombreCortoCompania As String
Global vgNombreCompania      As String
Global vgNombreSubSistema    As String

Global Const vgTipoPerJuridica = "J"
Global Const vgTipoPerNatural = "N"

Global lpAppName As String
Global lpKeyName As String
Global lpDefault As String
Global lpReturnString As String
Global Size As Integer
Global lpFileName As String

'Datos correspondientes al Usuario que accesa el programa
Global vgLogin   As String
Global vgContraseña As String
Global vgNivel      As Integer

Global vgPalabra    As String
Global vgPalabraAux As String
Global vgPalabraAux2 As String
Global vgRes        As Long

Global vgSw         As Boolean
Global vgI          As Integer
Global vgX          As Integer
Global vgError      As Long

'Nuevas Variables
Global vgTipoTablaGeneral  As String
Global vgGlosaTablaGeneral As String
Global vgTituloGeneral     As String
Global vgIndAjuste         As Boolean
Global vgTipoInforme       As String
Global vgTituloInfControl  As String
Global vgFechaEfecto       As String

Global vgHoraActual        As String
    
Global vgNombreRegion     As String
Global vgNombreProvincia  As String
Global vgNombreComuna     As String
Global vgCodigoRegion   As String
Global vgCodigoProvincia As String
Global vgCodigoComuna As String
Global vgNomForm          As String

Global vgTipoCauEnd As String


Public DatabaseName As String
Public ServerName   As String
Public ProviderName As String
Public UserName     As String
Public PasswordName As String

'CMV - 20050708 I
'Código Indicador de Calculo de Declaración de Ingresos
Global vgCodIndCalDecIng As String
'CMV - 20050708 F

'CMV-20060302 I
Global Const vgCodInsSaludFonasa = "13"
'CMV-20060302 F

'******** F I N    V A R    A G R E G A D A S  *******************
Global Const vgTipoBase = "ORACLE" 'Para indicar el Tipo de Base de Datos
'Global Const vgTipoBase = "SQL" 'SQL Server Tipo de Base de Datos

Global vgNombreInformeSeleccionado As String
Global vgNombreInformePeriodoSeleccionado As String
Global vgNombreInformeSeleccionadoInd As String
Global vgPlanillaPeriodoSeleccionada As String
Global vgNombreInfSeleccionadoProceso As String 'Utilizado en Frm_PlanillaProceso
Global vgNombreInformePlanillaCont As String 'Utilizado en Frm_PlanillaCont
'Para Procesos de Frm_CargaArchivo
Global vgNomInfSeleccionado As String
Global vgFechaIni As String
Global vgFechaTer As String
'Global Const cgCodTipMonedaUF As String * 2 = "UF"
Global Const cgCodTipMonedaPESOS As String * 5 = "PESOS"

'Borrar
Global Const cgCodTipMonedaUF As String * 2 = "NS"
Global Const cgNomTipMonedaUF As String * 2 = "S/."
Global vgMonedaCodOfi  As String
Global vgMonedaCodTran As String
'Borrar

'hqr 03/12/2010
Global Const cgSINAJUSTE = 0
Global Const cgAJUSTESOLES = 1
Global Const cgAJUSTETASAFIJA = 2
'fin hqr 03/12/2010

Global vgPagoCCaf As String
Global vgNomInfCargaArchivoAnno As String 'Utilizado en Frm_Carga Archivo Anno
'CMV-20060316 I
Global vgApeManualSel As String 'Utilizado en Frm_AperturaManual
'CMV-20060316 F
Global vgGlosaQuiebra As String

'Estructura para los Tipos de Moneda
Public Type TypeTablaMoneda
    Codigo      As String
    Descripcion As String
    Scomp       As String
End Type
Global egTablaMoneda()           As TypeTablaMoneda
Global vgNumeroTotalTablasMoneda As Long

'Variables Frm_EndosoPol CMV
'Estructura de Poliza
Public Type TyPoliza
    num_poliza As String
    num_endoso As Integer
    Cod_TipPension As String
    Cod_Estado As String
    Cod_TipRen As String
    Cod_Modalidad As String
    Num_Cargas As Integer
    Fec_Vigencia As String
    Fec_TerVigencia As String
    Mto_Prima As Double
    Mto_Pension As Double
    Num_MesDif As Integer
    Num_MesGar As Integer
    Prc_TasaCe As Double
    Prc_TasaVta As Double
    Prc_TasaIntPerGar As Double
    Fec_Emision As String
    Fec_Devengue As String
    Fec_IniPenCia As String
    Cod_Moneda As String
    Mto_ValMoneda As Double
    Mto_PensionGar As Double
    Ind_Cob As String
    Cod_CoberCon As String
    Mto_FacPenElla As Double
    Prc_FacPenElla As Double
    Cod_DerCre As String
    Cod_DerGra As String
    Cod_Cuspp As String
    Cod_TipReajuste As String 'hqr 13/01/2011
    Mto_ValReajusteTri As Double 'hqr 13/01/2011
    Mto_ValReajusteMen As Double 'hqr 26/02/2011
    Fec_TerPerGar As String
End Type

Public Type TyBeneficiarios
    num_poliza As String
    num_endoso As Integer
    Num_Orden As Integer
    Fec_Ingreso As String
    Cod_TipoIdenBen As String
    Num_IdenBen As String
    Gls_NomBen As String
    Gls_NomSegBen As String
    Gls_PatBen As String
    Gls_MatBen As String
    Gls_DirBen As String
    Cod_Direccion As Integer
    Gls_FonoBen As String
    Gls_CorreoBen As String
    Cod_GruFam As String
    Cod_Par As String
    Cod_Sexo As String
    Cod_SitInv As String
    Cod_DerCre As String
    Cod_DerPen As String
    Cod_CauInv As String
    Fec_NacBen  As String
    Fec_NacHM As String
    Fec_InvBen As String
    Cod_MotReqPen As String
    Mto_Pension As Double
    Mto_PensionGar As Double
    Prc_Pension As Double
    Cod_InsSalud As String
    Cod_ModSalud As String
    Mto_PlanSalud As Double
    Cod_EstPension As String
    Cod_CajaCompen As String
    Cod_ViaPago As String
    Cod_Banco As String
    Cod_TipCuenta As String
    Num_Cuenta As String
    Cod_Sucursal As String
    Fec_FallBen As String
    Fec_Matrimonio As String
    Cod_CauSusBen As String
    Fec_SusBen As String
    Fec_IniPagoPen As String
    Fec_TerPagoPenGar As String
    Cod_UsuarioCrea As String
    Fec_Crea As String
    Hor_Crea As String
    Prc_PensionLeg As Double
    Prc_PensionGar As Double
    Gls_Telben2 As String
    cod_tipcta As String
    cod_monbco As String
    num_ctabco As String
    'mvg 20170904
    ind_bolelec As String
    
    'INICIO GCP-FRACTAL 20190104
    NUM_CUENTA_CCI As String
    CONS_TRAINFO As String
    CONS_DATCOMER As String
    'FIN GCP-FRACTAL 20190104

    COD_MODTIPOCUENTA_MANC As String
    COD_TIPODOC_MANC As String
    NUM_DOC_MANC As String
    NOMBRE_MANC As String
    APELLIDO_MANC As String
    
   
End Type

'Public Type TyBeneficiarios
'    Num_Poliza As String
'    num_endoso As Integer
'    Num_Orden As Integer
'    Rut_Ben As Double
'    Dgv_Ben As String
'    Gls_NomBen As String
'    Gls_PatBen As String
'    Gls_MatBen As String
'    Cod_GruFam As String
'    Cod_Par As String
'    Cod_Sexo As String
'    Cod_SitInv As String
'    Cod_DerCre As String
'    Cod_EstPension As String
'    Cod_CauInv As String
'    Fec_NacBen As String
'    Fec_NacHM As String
'    Fec_InvBen As String
'    Mto_Pension As Double
'    Prc_Pension As Double
'    Fec_FallBen As String
'    Cod_DerPen As String
'    Cod_MotReqPen As String
'    Mto_PensionGar As Double
'    Cod_CauSusBen As String
'    Fec_SusBen As String
'    Fec_IniPagoPen As String
'    Fec_TerPagoPenGar As String
'    Fec_Matrimonio As String
'End Type

''CMV-20060607 I
'Public Type TyBenEndAuto
'    Num_Poliza As String
'    num_endoso As Integer
'    Num_Orden As Integer
'    Fec_Ingreso As String
'    Rut_Ben As Double
'    Dgv_Ben As String
'    Gls_NomBen As String
'    Gls_PatBen As String
'    Gls_MatBen As String
'    Gls_DirBen As String
'    Cod_Direccion As Integer
'    Gls_FonoBen As String
'    Gls_CorreoBen As String
'    Cod_GruFam As String
'    Cod_Par As String
'    Cod_Sexo As String
'    Cod_SitInv As String
'    Cod_DerCre As String
'    Cod_DerPen As String
'    Cod_CauInv As String
'    Fec_NacBen  As String
'    Fec_NacHM As String
'    Fec_InvBen As String
'    Cod_MotReqPen As String
'    Mto_Pension As Double
'    Mto_PensionGar As Double
'    Prc_Pension As Double
'    Cod_InsSalud As String
'    Cod_ModSalud As String
'    Mto_PlanSalud As Double
'    Cod_EstPension As String
'    Cod_CajaCompen As String
'    Cod_ViaPago As String
'    Cod_Banco As String
'    Cod_TipCuenta As String
'    Num_Cuenta As String
'    Cod_Sucursal As String
'    Fec_FallBen As String
'    Fec_Matrimonio As String
'    Cod_CauSusBen As String
'    Fec_SusBen As String
'    Fec_IniPagoPen As String
'    Fec_TerPagoPenGar As String
'    Cod_UsuarioCrea As String
'    Fec_Crea As String
'    Hor_Crea As String
'    Cod_ModSalud2 As String
'    Mto_PlanSalud2 As Double
'End Type
'Global stBenEndAuto() As TyBenEndAuto 'Registro de Tabla Beneficiarios para endosos automáticos
''CMV-2006007 F

Global stPolizaOri        As TyPoliza 'Registro de Poliza Original
Global stPolizaMod        As TyPoliza 'Registro de Poliza Modificada
Global stBeneficiariosOri() As TyBeneficiarios 'Registro de Beneficiarios Originales
Global stBeneficiariosMod() As TyBeneficiarios 'Registro de Beneficiarios Modificados
Global stBenEndAuto() As TyBeneficiarios 'Registro de Tabla Beneficiarios para endosos automáticos


'CMV
Global vgNumBen As Integer
'ANI 24/02/2005
Global vgValorParametro As Double
Global vgValorPorcentaje As Double

'CMV 20060630 I
Global vgGlsTipoForm As String
Global vgNumPol As String
Global vgNumOrden As Integer
Global vgRutBen As String
Global vgDgvBen As String
Global vgNomBen As String
'CMV 20060630 F

Global vgTipoSucursal As String
Global Const cgIndicadorSi As String * 2 = "Si"
Global Const cgIndicadorNo As String * 2 = "No"
Global Const cgTipoSucursalSuc As String * 1 = "S"
Global Const cgTipoSucursalAfp As String * 1 = "A"

Global Const cgPagoTerceroCuoMor As String * 2 = "CM"
Global Const cgPagoTerceroPerGar As String * 2 = "PG"

Global Const cgTipoIdenRuc As String * 1 = "9"

'marco---08/03/2010

  Public Ident_1 As String
  Public Ident_2 As String
  Public Tipo_J As String
  Public Tipo_S As String
  Public Tipo_I As String
  Public Nombre_Afiliado As String
  Public Nombre_Beneficiario As String
  Public Tipo_num_documento_afiliado As String
  Public Tipo_num_documento_beneficiario As String
  Public num_poliza As String
  Public RangoFecha As String
  Public Fecha_Creacion As String
  Public sName_Reporte As String



'RRR18/01/2012

Public tammin As Integer
Public cantclvant As Integer
Public canmincaralf As Integer
Public freccambio As Integer
Public canantclv As Integer
Public FechaIni As String
Public FechaFin As String
Public fechaant As String
Public fecfinDa As Date
Public vlPassword As String
Global vgChkdiaant As Integer
Global vgValorAr As Integer
Dim balfanum As Integer
Global vgDiasFaltan As Integer
Global vgIntentos As Integer
Global vgOptTE As String
Global strRpt As String

''CORPTEC
Global num_session_pension  As Double

'Implementacion GobiernoDeDatos()_
Public Type TyBeneficiariosEst
    num_poliza As String
    num_endoso As Integer
    Num_Orden As Integer
    Fec_Ingreso As String
    Cod_TipoIdenBen As String
    Num_IdenBen As String
    Cod_Direccion As String
    cod_tip_fonoben As String
    cod_area_fonoben As String
    Gls_FonoBen As String
    cod_tipo_telben2 As String
    cod_area_telben2 As String
    Gls_Telben2 As String
    pTipoVia As String
    pDireccion As String
    pNumero As String
    pTipoPref As String
    pInterior As String
    pManzana As String
    pLote As String
    pEtapa As String
    pTipoConj As String
    pConjHabit As String
    pTipoBlock As String
    pNumBlock As String
    pReferencia As String
    pConcatDirec As String
    pGlsCorreo As String
    pvalEndosoGS As String
End Type
Global stPolizaBenDirec() As TyBeneficiariosEst ' registro de direccion estructurada
Global stPolizaBenDirecMod() As TyBeneficiariosEst
Global stPolizaBenDirecOri() As TyBeneficiariosEst
'fin Implementacion GobiernoDeDatos()_
Function fgLogIn_Pen() As Boolean
    Dim com As ADODB.Command
   Dim sistema, modulo, Estado As String
    sistema = "SEACSA"
    modulo = "PENSIONES"
    Estado = "A"
 
    Set com = New ADODB.Command
    
    vgConexionBD.BeginTrans
    com.ActiveConnection = vgConexionBD
    com.CommandText = "SP_LOG_SESSION"
   com.CommandType = adCmdStoredProc
    
    com.Parameters.Append com.CreateParameter("ESTADO", adChar, adParamInput, 1, Estado)
    com.Parameters.Append com.CreateParameter("USUARIO", adVarChar, adParamInput, 10, vgLogin)
    com.Parameters.Append com.CreateParameter("SISTEMA", adVarChar, adParamInput, 50, sistema)
    com.Parameters.Append com.CreateParameter("MODULO", adVarChar, adParamInput, 50, modulo)
    com.Parameters.Append com.CreateParameter("IDLOG", adDouble, adParamInput, 2, 0)
    com.Parameters.Append com.CreateParameter("Retorno", adDouble, adParamReturnValue)
    com.Execute
    
    vgConexionBD.CommitTrans
    num_session_pension = com("Retorno")

End Function

Function fgLogOut_Pen() As Boolean

    Dim com As ADODB.Command
    Dim sistema, modulo, Estado As String
    sistema = "SEACSA"
    modulo = "PENSIONES"
    Estado = "I"
    Set com = New ADODB.Command
    
    vgConexionBD.BeginTrans
    com.ActiveConnection = vgConexionBD
    com.CommandText = "SP_LOG_SESSION"
    com.CommandType = adCmdStoredProc
    
    com.Parameters.Append com.CreateParameter("ESTADO", adChar, adParamInput, 1, Estado)
    com.Parameters.Append com.CreateParameter("USUARIO", adVarChar, adParamInput, 10, vgLogin)
    com.Parameters.Append com.CreateParameter("SISTEMA", adVarChar, adParamInput, 50, sistema)
    com.Parameters.Append com.CreateParameter("MODULO", adVarChar, adParamInput, 50, modulo)
    com.Parameters.Append com.CreateParameter("IDLOG", adDouble, adParamInput, 2, num_session_pension)
    com.Parameters.Append com.CreateParameter("Retorno", adDouble, adParamReturnValue)
    com.Execute
    vgConexionBD.CommitTrans
    num_session_pension = com("Retorno")
End Function
'rrr
'************************ F U N C I O N E S RRR******************
'RRR 18/01/2012
Public Function fIaplicavalidacion(usuario As String, txt_password As TextBox, txt_passwordcomfir As TextBox) As Integer

    vgSql = "SELECT * FROM MA_TMAE_ADMINCUENTAS WHERE "
    vgSql = vgSql & "cod_cliente = '1' "
    Set vgRs = vgConexionBD.Execute(vgSql)

    If Not vgRs.EOF Then
        tammin = vgRs!ntamañomin
        cantclvant = vgRs!ncanclvant
        canmincaralf = vgRs!ncaracmin
        freccambio = vgRs!nfrecuencia
        canantclv = vgRs!ncanclvant
        balfanum = vgRs!balfanum
    End If
    
    If (txt_password <> txt_passwordcomfir) Then
            MsgBox "Las Contraseñas registradas son distintas, vuelva a registrarlas.", vbExclamation, "Error de Contraseña"
            txt_password = ""
            txt_passwordcomfir = ""
            txt_password.SetFocus
            fIaplicavalidacion = 0
        Exit Function
    End If
    
    If Len(txt_password) < tammin Then
        MsgBox "Password debe ser minimo de " & CStr(tammin) & " caracteres ", vbCritical, "Error de Datos"
        txt_password.SetFocus
        fIaplicavalidacion = 0
        Exit Function
    End If
   
    If Len(txt_passwordcomfir) < tammin Then
        MsgBox "Password debe ser minimo de " & CStr(tammin) & " caracteres ", vbCritical, "Error de Datos"
        txt_passwordcomfir.SetFocus
        fIaplicavalidacion = 0
        Exit Function
    End If
    
    vgRs.Close
    
       
    vgSql = " select nro_usupass, gls_password from MA_TMAE_USUPASSWORD "
    vgSql = vgSql & " where cod_usuario='" & usuario & "' "
    vgSql = vgSql & " and nro_usupass > (select count(*) from MA_TMAE_USUPASSWORD where cod_usuario='" & usuario & "') - " & cantclvant
    vgSql = vgSql & " and nro_usupass <= (select count(*) from MA_TMAE_USUPASSWORD where cod_usuario='" & usuario & "')"
    vgSql = vgSql & " order by 1 desc"


    Set vgRs = vgConexionBD.Execute(vgSql)

    Dim strclave As String
    
    If Not vgRs.EOF Then
        Do While Not vgRs.EOF
            strclave = fgDesPassword(vgRs!gls_password)
            
            If UCase(Trim(txt_password)) = strclave Then
                MsgBox "No puede utilizar un password anterior, Por favor elegir otro.", vbCritical, "Error de Datos"
                txt_password.SetFocus
                fIaplicavalidacion = 0
                Exit Function
            End If
            vgRs.MoveNext
        Loop
    End If
    
   Dim i, l, n, a As Integer
   Dim car As String

    
    For i = 1 To Len(txt_password)
    
        car = Mid(txt_password, i, 1)
    
        If VLetras(Asc(car)) <> 0 Then l = l + 1
        If Numeros(Asc(car)) <> 0 Then n = n + 1
        If VAlfanumerico(Asc(car)) <> 0 Then a = a + 1
        
    Next
    
    If balfanum = 1 Then
        If a < canmincaralf Then
            MsgBox "La clave debe contener como minimo " & canmincaralf & " caracteres alfanumericos.", vbCritical, "Error de Datos"
            txt_password.SetFocus
            fIaplicavalidacion = 0
            Exit Function
        End If
    Else
        If a > 0 Then
            MsgBox "La clave no debe contener caracteres alfanumericos.", vbCritical, "Error de Datos"
            fIaplicavalidacion = 0
            Exit Function
        End If
    End If
    
    FechaIni = Mid(CStr(Now), 7, 4) & Mid(CStr(Now), 4, 2) & Mid(CStr(Now), 1, 2)
    fecfinDa = DateAdd("d", freccambio, Now)
    FechaFin = Mid(CStr(DateAdd("d", freccambio, Now)), 7, 4) & Mid(CStr(DateAdd("d", freccambio, Now)), 4, 2) & Mid(CStr(DateAdd("d", freccambio, Now)), 1, 2)
    fechaant = Mid(CStr(DateAdd("d", -CInt(canantclv), fecfinDa)), 7, 4) & Mid(CStr(DateAdd("d", -CInt(canantclv), fecfinDa)), 4, 2) & Mid(CStr(DateAdd("d", -CInt(canantclv), fecfinDa)), 1, 2)
    
    fIaplicavalidacion = 1
End Function

Public Function VLetras(Tecla As Integer) As Integer
Dim strValido As String
'letras no validas: .*-}¿'!%&/()=?¡]¨*[Ññ;:_ áéíó
strValido = "qwertyuioplkjhgfdsazxcvbnmQWERTYUIOPASDFGHJKLZXCV BNM, "
If Tecla > 26 Then
If InStr(strValido, Chr(Tecla)) = 0 Then
Tecla = 0
End If
End If
VLetras = Tecla
End Function
Public Function Numeros(Tecla As Integer) As Integer
Dim strValido As String
strValido = "0123456789"
If Tecla > 26 Then
'compara los numeros ke hay en la variable strValido _
con el numero ingresado(Tecla)
'si el numero ingresado(Tecla) no esta en la variable strValido entonces _
Tecla = 0, la funcion Chr convierte el numero a ascii
If InStr(strValido, Chr(Tecla)) = 0 Then
Tecla = 0
End If
End If
Numeros = Tecla
End Function

Public Function VAlfanumerico(Tecla As Integer) As Integer
Dim strValido As String
'letras no validas: .*-}¿'!%&/()=?¡]¨*[Ññ;:_ áéíó
strValido = "!#$%&/()=?¡'¿{}^`[]*\-+.,;:_ "
If Tecla > 26 Then
If InStr(strValido, Chr(Tecla)) = 0 Then
Tecla = 0
End If
End If
VAlfanumerico = Tecla
End Function

'RRR 18/01/2012


'Function flCargaEstructuraPoliza(iNombreTabla As String, iPoliza As String, iEndoso As Integer, istPoliza As TyPoliza)
'On Error GoTo Err_flCargaEstructuraPoliza
'
'    vgSql = ""
'    vgSql = "SELECT num_poliza,num_endoso,cod_tippension,cod_estado, "
'    vgSql = vgSql & "cod_tipren,cod_modalidad,num_cargas,fec_vigencia, "
'    vgSql = vgSql & "fec_tervigencia,mto_prima,mto_pension,num_mesdif, "
'    vgSql = vgSql & "num_mesgar,prc_tasace,prc_tasavta,prc_tasaintpergar "
'    vgSql = vgSql & "FROM " & iNombreTabla & " WHERE "
'    vgSql = vgSql & "num_poliza = '" & Trim(iPoliza) & "' AND "
'    vgSql = vgSql & "num_endoso = " & iEndoso & " "
'    vgSql = vgSql & " ORDER BY num_endoso DESC"
'    Set vgRegistro = vgConexionBD.Execute(vgSql)
'    If Not vgRegistro.EOF Then
'       With istPoliza
'            If IsNull(vgRegistro!Num_Poliza) Then
'               .Num_Poliza = ""
'            Else
'                .Num_Poliza = (vgRegistro!Num_Poliza)
'            End If
'            If IsNull(vgRegistro!num_endoso) Then
'               .num_endoso = ""
'            Else
'                .num_endoso = (vgRegistro!num_endoso)
'            End If
'            If IsNull(vgRegistro!Cod_TipPension) Then
'               .Cod_TipPension = ""
'            Else
'                .Cod_TipPension = (vgRegistro!Cod_TipPension)
'            End If
'            If IsNull(vgRegistro!Cod_Estado) Then
'               .Cod_Estado = ""
'            Else
'                .Cod_Estado = (vgRegistro!Cod_Estado)
'            End If
'            If IsNull(vgRegistro!Cod_TipRen) Then
'               .Cod_TipRen = ""
'            Else
'                .Cod_TipRen = (vgRegistro!Cod_TipRen)
'            End If
'            If IsNull(vgRegistro!Cod_Modalidad) Then
'               .Cod_Modalidad = ""
'            Else
'                .Cod_Modalidad = (vgRegistro!Cod_Modalidad)
'            End If
'            If IsNull(vgRegistro!Num_Cargas) Then
'               .Num_Cargas = ""
'            Else
'                .Num_Cargas = (vgRegistro!Num_Cargas)
'            End If
'            If IsNull(vgRegistro!Fec_Vigencia) Then
'               .Fec_Vigencia = ""
'            Else
'                .Fec_Vigencia = (vgRegistro!Fec_Vigencia)
'            End If
'            .Fec_TerVigencia = (vgRegistro!Fec_TerVigencia)
'            If IsNull(vgRegistro!Fec_TerVigencia) Then
'               .Mto_Prima = ""
'            Else
'                .Mto_Prima = (vgRegistro!Mto_Prima)
'            End If
'            If IsNull(vgRegistro!Mto_Pension) Then
'               .Mto_Pension = ""
'            Else
'                .Mto_Pension = (vgRegistro!Mto_Pension)
'            End If
'            If IsNull(vgRegistro!Num_MesDif) Then
'               .Num_MesDif = ""
'            Else
'                .Num_MesDif = (vgRegistro!Num_MesDif)
'            End If
'            If IsNull(vgRegistro!Num_MesGar) Then
'               .Num_MesGar = ""
'            Else
'                .Num_MesGar = (vgRegistro!Num_MesGar)
'            End If
'            If IsNull(vgRegistro!Prc_TasaCe) Then
'               .Prc_TasaCe = ""
'            Else
'                .Prc_TasaCe = (vgRegistro!Prc_TasaCe)
'            End If
'            If IsNull(vgRegistro!Prc_TasaVta) Then
'               .Prc_TasaVta = ""
'            Else
'                .Prc_TasaVta = (vgRegistro!Prc_TasaVta)
'            End If
'            If IsNull(vgRegistro!Prc_TasaIntPerGar) Then
'               .Prc_TasaIntPerGar = ""
'            Else
'                .Prc_TasaIntPerGar = (vgRegistro!Prc_TasaIntPerGar)
'            End If
'       End With
'    End If
'
'Exit Function
'Err_flCargaEstructuraPoliza:
'    Screen.MousePointer = 0
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
'
'End Function
'
'Function flCargaEstructuraBeneficiarios(iNombreTabla As String, iPoliza As String, iEndoso As Integer, istBeneficiarios() As TyBeneficiarios)
'On Error GoTo Err_flCargaEstructuraBeneficiarios
'
'    vgSql = ""
'    vgSql = "SELECT COUNT (num_orden) as numero "
'    vgSql = vgSql & "FROM " & iNombreTabla & " WHERE "
'    vgSql = vgSql & "num_poliza = '" & Trim(iPoliza) & "' AND "
'    vgSql = vgSql & "num_endoso = " & iEndoso & " "
'    Set vgRegistro = vgConexionBD.Execute(vgSql)
'    If Not vgRegistro.EOF Then
'       vgNumBen = (vgRegistro!numero)
'    End If
'
'    ReDim istBeneficiarios(vgNumBen) As TyBeneficiarios
'
'    vgSql = ""
'    vgSql = "SELECT num_poliza,num_endoso,num_orden,rut_ben,dgv_ben, "
'    vgSql = vgSql & "gls_nomben,gls_patben,gls_matben,cod_grufam, "
'    vgSql = vgSql & "cod_par,cod_sexo,cod_sitinv,cod_dercre,cod_derpen, "
'    vgSql = vgSql & "cod_cauinv,fec_nacben,fec_nachm,fec_invben, "
'    vgSql = vgSql & "mto_pension,prc_pension,fec_fallben,cod_estpension, "
'    vgSql = vgSql & "cod_motreqpen,mto_pensiongar,cod_caususben,fec_susben, "
'    vgSql = vgSql & "fec_inipagopen,fec_terpagopengar,fec_matrimonio "
'    vgSql = vgSql & "FROM " & iNombreTabla & " WHERE "
'    vgSql = vgSql & "num_poliza = '" & Trim(iPoliza) & "' AND "
'    vgSql = vgSql & "num_endoso = " & iEndoso & " "
'    vgSql = vgSql & " ORDER BY num_orden ASC"
'    Set vgRegistro = vgConexionBD.Execute(vgSql)
'    If Not vgRegistro.EOF Then
'       vgX = 0
'       While Not vgRegistro.EOF
'             vgX = vgX + 1
'             With istBeneficiarios(vgX)
'                  If IsNull(vgRegistro!Num_Poliza) Then
'                     .Num_Poliza = ""
'                  Else
'                      .Num_Poliza = (vgRegistro!Num_Poliza)
'                  End If
'                  If IsNull(vgRegistro!num_endoso) Then
'                     .num_endoso = ""
'                  Else
'                      .num_endoso = (vgRegistro!num_endoso)
'                  End If
'                  If IsNull(vgRegistro!Num_Orden) Then
'                     .Num_Orden = ""
'                  Else
'                      .Num_Orden = (vgRegistro!Num_Orden)
'                  End If
'                  If IsNull(vgRegistro!Rut_Ben) Then
'                     .Rut_Ben = ""
'                  Else
'                      .Rut_Ben = (vgRegistro!Rut_Ben)
'                  End If
'                  If IsNull(vgRegistro!Dgv_Ben) Then
'                     .Dgv_Ben = ""
'                  Else
'                      .Dgv_Ben = (vgRegistro!Dgv_Ben)
'                  End If
'                  If IsNull(vgRegistro!Gls_NomBen) Then
'                     .Gls_NomBen = ""
'                  Else
'                      .Gls_NomBen = (vgRegistro!Gls_NomBen)
'                  End If
'                  If IsNull(vgRegistro!Gls_PatBen) Then
'                     .Gls_PatBen = ""
'                  Else
'                      .Gls_PatBen = (vgRegistro!Gls_PatBen)
'                  End If
'                  If IsNull(vgRegistro!Gls_MatBen) Then
'                     .Gls_MatBen = ""
'                  Else
'                      .Gls_MatBen = (vgRegistro!Gls_MatBen)
'                  End If
'                  If IsNull(vgRegistro!Cod_GruFam) Then
'                     .Cod_GruFam = ""
'                  Else
'                      .Cod_GruFam = (vgRegistro!Cod_GruFam)
'                  End If
'                  If IsNull(vgRegistro!Cod_Par) Then
'                     .Cod_Par = ""
'                  Else
'                      .Cod_Par = (vgRegistro!Cod_Par)
'                  End If
'                  If IsNull(vgRegistro!Cod_Sexo) Then
'                     .Cod_Sexo = ""
'                  Else
'                      .Cod_Sexo = (vgRegistro!Cod_Sexo)
'                  End If
'                  If IsNull(vgRegistro!Cod_SitInv) Then
'                     .Cod_SitInv = ""
'                  Else
'                      .Cod_SitInv = (vgRegistro!Cod_SitInv)
'                  End If
'                  If IsNull(vgRegistro!Cod_DerCre) Then
'                     .Cod_DerCre = ""
'                  Else
'                      .Cod_DerCre = (vgRegistro!Cod_DerCre)
'                  End If
'                  If IsNull(vgRegistro!Cod_EstPension) Then
'                     .Cod_EstPension = ""
'                  Else
'                      .Cod_EstPension = (vgRegistro!Cod_EstPension)
'                  End If
'                  If IsNull(vgRegistro!Cod_CauInv) Then
'                     .Cod_CauInv = ""
'                  Else
'                      .Cod_CauInv = (vgRegistro!Cod_CauInv)
'                  End If
'                  If IsNull(vgRegistro!Fec_NacBen) Then
'                     .Fec_NacBen = ""
'                  Else
'                      .Fec_NacBen = (vgRegistro!Fec_NacBen)
'                  End If
'                  If IsNull(vgRegistro!Fec_NacHM) Then
'                     .Fec_NacHM = ""
'                  Else
'                      .Fec_NacHM = (vgRegistro!Fec_NacHM)
'                  End If
'                  If IsNull(vgRegistro!Fec_InvBen) Then
'                     .Fec_InvBen = ""
'                  Else
'                      .Fec_InvBen = (vgRegistro!Fec_InvBen)
'                  End If
'                  If IsNull(vgRegistro!Mto_Pension) Then
'                     .Mto_Pension = ""
'                  Else
'                      .Mto_Pension = (vgRegistro!Mto_Pension)
'                  End If
'                  If IsNull(vgRegistro!Prc_Pension) Then
'                     .Prc_Pension = ""
'                  Else
'                      .Prc_Pension = (vgRegistro!Prc_Pension)
'                  End If
'                  If IsNull(vgRegistro!Fec_FallBen) Then
'                     .Fec_FallBen = ""
'                  Else
'                      .Fec_FallBen = (vgRegistro!Fec_FallBen)
'                  End If
'                  If IsNull(vgRegistro!Cod_DerPen) Then
'                     .Cod_DerPen = ""
'                  Else
'                      .Cod_DerPen = (vgRegistro!Cod_DerPen)
'                  End If
'                  If IsNull(vgRegistro!Cod_MotReqPen) Then
'                     .Cod_MotReqPen = ""
'                  Else
'                      .Cod_MotReqPen = (vgRegistro!Cod_MotReqPen)
'                  End If
'                  If IsNull(vgRegistro!Mto_PensionGar) Then
'                     .Mto_PensionGar = ""
'                  Else
'                      .Mto_PensionGar = (vgRegistro!Mto_PensionGar)
'                  End If
'                  If IsNull(vgRegistro!Cod_CauSusBen) Then
'                     .Cod_CauSusBen = ""
'                  Else
'                      .Cod_CauSusBen = (vgRegistro!Cod_CauSusBen)
'                  End If
'                  If IsNull(vgRegistro!Fec_SusBen) Then
'                     .Fec_SusBen = ""
'                  Else
'                      .Fec_SusBen = (vgRegistro!Fec_SusBen)
'                  End If
'                  If IsNull(vgRegistro!Fec_IniPagoPen) Then
'                     .Fec_IniPagoPen = ""
'                  Else
'                      .Fec_IniPagoPen = (vgRegistro!Fec_IniPagoPen)
'                  End If
'                  If IsNull(vgRegistro!Fec_TerPagoPenGar) Then
'                     .Fec_TerPagoPenGar = ""
'                  Else
'                      .Fec_TerPagoPenGar = (vgRegistro!Fec_TerPagoPenGar)
'                  End If
'                  If IsNull(vgRegistro!Fec_Matrimonio) Then
'                     .Fec_Matrimonio = ""
'                  Else
'                      .Fec_Matrimonio = (vgRegistro!Fec_Matrimonio)
'                  End If
'             End With
'             vgRegistro.MoveNext
'       Wend
'    End If
'
'Exit Function
'Err_flCargaEstructuraBeneficiarios:
'    Screen.MousePointer = 0
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
'
'End Function
'
'Function flCargaGrillaBeneficiarios(iGrilla As MSFlexGrid, istBeneficiarios() As TyBeneficiarios)
'On Error GoTo Err_flCargaGrillaBeneficiarios
'Dim vgCodPar As String
'
'    vgX = 0
'    'Call flInicializaGrillaBenef(iGrilla)
'    While vgX < vgNumBen
'          vgX = vgX + 1
'          With istBeneficiarios(vgX)
'               'vgCodPar = " " & Trim(.Cod_Par) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_Par, Trim(.Cod_Par)))
'               iGrilla.AddItem (.Num_Orden) & vbTab _
'               & (" " & Format((Trim(.Rut_Ben)), "##,###,##0") & " - " & (Trim(.Dgv_Ben))) & vbTab _
'               & (Trim(.Gls_NomBen)) & vbTab _
'               & (Trim(.Gls_PatBen)) & vbTab _
'               & (Trim(.Gls_MatBen)) & vbTab _
'               & (Trim(.Cod_Par)) & vbTab _
'               & (Trim(.Cod_GruFam)) & vbTab _
'               & (Trim(.Cod_Sexo)) & vbTab _
'               & (Trim(.Cod_SitInv)) & vbTab _
'               & (Trim(.Cod_EstPension)) & vbTab _
'               & (Trim(.Cod_DerCre)) & vbTab _
'               & (Trim(.Num_Poliza)) & vbTab _
'               & (Trim(.num_endoso)) & vbTab _
'               & (Trim(.Cod_CauInv)) & vbTab _
'               & (Trim(.Fec_NacBen)) & vbTab _
'               & (Trim(.Fec_NacHM)) & vbTab _
'               & (Trim(.Fec_InvBen)) & vbTab _
'               & (Trim(.Mto_Pension)) & vbTab _
'               & (Trim(.Prc_Pension)) & vbTab _
'               & (Trim(.Fec_FallBen)) & vbTab _
'               & (Trim(.Cod_DerPen)) & vbTab _
'               & (Trim(.Cod_MotReqPen)) & vbTab _
'               & (Trim(.Mto_PensionGar)) & vbTab _
'               & (Trim(.Cod_CauSusBen)) & vbTab _
'               & (Trim(.Fec_SusBen)) & vbTab & (Trim(.Fec_IniPagoPen)) & vbTab & (Trim(.Fec_TerPagoPenGar)) & vbTab & (Trim(.Fec_Matrimonio))
'
'          End With
'    Wend
'
'Exit Function
'Err_flCargaGrillaBeneficiarios:
'    Screen.MousePointer = 0
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
'
'End Function
'
'Function fgCargaEstBenGrilla(iGrilla As MSFlexGrid, istBeneficiarios() As TyBeneficiarios)
'On Error GoTo Err_fgCargaEstBenGrilla
'Dim vlPos, vlNumero As Integer
'
'    If iGrilla.Rows > 1 Then
'    vlPos = 1
'    iGrilla.Col = 0
'    vgX = 0
'    vgNumBen = (iGrilla.Rows - 1)
'    ReDim istBeneficiarios(vgNumBen) As TyBeneficiarios
'    While vlPos <= (iGrilla.Rows - 1)
'            iGrilla.Row = vlPos
'            iGrilla.Col = 0
'
'            vgX = vgX + 1
'            With istBeneficiarios(vgX)
'                 iGrilla.Col = 11
'                 .Num_Poliza = (iGrilla.Text)
'                 iGrilla.Col = 12
'                 .num_endoso = (iGrilla.Text)
'                 iGrilla.Col = 0
'                 .Num_Orden = (iGrilla.Text)
'                 iGrilla.Col = 1
'                 vlNumero = InStr(iGrilla.Text, "-")
'                 .Rut_Ben = ((Str(Trim(Mid(iGrilla.Text, 1, vlNumero - 1)))))
'                 .Dgv_Ben = ((Trim(Mid(iGrilla.Text, vlNumero + 1, 2))))
'                 iGrilla.Col = 2
'                 .Gls_NomBen = (iGrilla.Text)
'                 iGrilla.Col = 3
'                 .Gls_PatBen = (iGrilla.Text)
'                 iGrilla.Col = 4
'                 .Gls_MatBen = (iGrilla.Text)
'                 iGrilla.Col = 6
'                 .Cod_GruFam = (iGrilla.Text)
'                 iGrilla.Col = 5
'                 .Cod_Par = (iGrilla.Text)
'                 iGrilla.Col = 7
'                 .Cod_Sexo = (iGrilla.Text)
'                 iGrilla.Col = 8
'                 .Cod_SitInv = (iGrilla.Text)
'                 iGrilla.Col = 10
'                 .Cod_DerCre = (iGrilla.Text)
'                 iGrilla.Col = 9
'                 .Cod_EstPension = (iGrilla.Text)
'                 iGrilla.Col = 13
'                 .Cod_CauInv = (iGrilla.Text)
'                 iGrilla.Col = 14
'                 .Fec_NacBen = (iGrilla.Text)
'                 iGrilla.Col = 15
'                 .Fec_NacHM = (iGrilla.Text)
'                 iGrilla.Col = 16
'                 .Fec_InvBen = (iGrilla.Text)
'                 iGrilla.Col = 17
'                 .Mto_Pension = (iGrilla.Text)
'                 iGrilla.Col = 18
'                 .Prc_Pension = (iGrilla.Text)
'                 iGrilla.Col = 19
'                 .Fec_FallBen = (iGrilla.Text)
'                 iGrilla.Col = 20
'                 .Cod_DerPen = (iGrilla.Text)
'                 iGrilla.Col = 21
'                 .Cod_MotReqPen = (iGrilla.Text)
'                 iGrilla.Col = 22
'                 .Mto_PensionGar = (iGrilla.Text)
'                 iGrilla.Col = 23
'                 .Cod_CauSusBen = (iGrilla.Text)
'                 iGrilla.Col = 24
'                 .Fec_SusBen = (iGrilla.Text)
'                 iGrilla.Col = 25
'                 .Fec_IniPagoPen = (iGrilla.Text)
'                 iGrilla.Col = 26
'                 .Fec_TerPagoPenGar = (iGrilla.Text)
'                 iGrilla.Col = 27
'                 .Fec_Matrimonio = (iGrilla.Text)
'
'            End With
'
'            vlPos = vlPos + 1
'       Wend
'    End If
'
'Exit Function
'Err_fgCargaEstBenGrilla:
'    Screen.MousePointer = 0
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
'
'End Function
'
'
'Function fgCalcularPorcentajeBenef(iFechaIniVig As String, iNumBenef As Integer, ostBeneficiarios() As TyBeneficiarios) As Boolean
''Función: Permite actualizar los Porcentajes de Pensión de los Beneficiarios,
''         su Derecho a Acrecer y la Fecha de Nacimiento del Hijo Menor
''Parámetros de Entrada/Salida:
''iFechaIniVig     => Fecha de Inicio de Vigencia de la Póliza
''iNumBenef        => Número de Beneficiarios
''ostBeneficiarios => Estructura desde la cual se obtienen los datos de los
''                    Beneficiarios y al mismo tiempo se calcula el Porcentaje
''                    de Pensión al cual tienen Dº
'
'Dim vlValor  As Double
'Dim Numor()  As Integer
'Dim Ncorbe() As Integer
'Dim Codrel() As Integer
'Dim Cod_Grfam() As Integer
'Dim Sexobe() As String, Inv()   As String, Coinbe()   As String
'Dim Derpen() As Integer
'Dim Nanbe()  As Integer, Nmnbe() As Integer, Ndnbe() As Integer
'Dim Porben() As Double
'Dim Codcbe() As String
'Dim Iaap     As Integer, Immp   As Integer, Iddp    As Integer
'Dim vlNum_Ben As Integer
'Dim Hijos()  As Integer, Hijos_Inv() As Integer
'Dim Hijo_Menor() As Date
'Dim Hijo_Menor_Ant() As Date
'Dim Fec_NacHM() As Date
'
'Dim cont_mhn() As Integer
'Dim cont_causante As Integer
'Dim cont_esposa As Integer
'Dim cont_mhn_tot As Integer
'Dim cont_hijo As Integer
'Dim cont_padres As Integer
'
'Dim L24 As Long, i As Long, edad_mes_ben As Long
'Dim fecha_sin As Long, vlContBen As Long
'Dim sexo_cau As String
'Dim g As Long, Q As Long, X As Long, j As Long, u As String, k As Long
'Dim v_hijo As Double
'
'On Error GoTo Err_fgCalcularPorcentaje
'
''    Call flAsignaPorcentajesLegales
''    Call flValidaBeneficiarios
''    Call flDerechoAcrecer
''    Call flMadreHijoMenor
''    Call flVariosConyuges
''    Call flHijosSolos
'
'    fgCalcularPorcentajeBenef = False
'    L24 = 0
'
'    'Debiera tomar la Fecha de Devengue
'
'    If (fgCarga_Param("LI", "L24", iFechaIniVig) = True) Then
'        L24 = vgValorParametro
'    Else
'        MsgBox "No existe Edad de tope para los 24 años.", vbCritical, "Proceso Cancelado"
'        Exit Function
'    End If
'
'    'Mensualizar la Edad de 24 Años
'    L24 = L24 * 12
'
'    'If Not IsDate(txt_fecha_devengo) Then
'    '   X = MsgBox("Debe ingresar la Fecha de Inicio de la Pensión", 16)
'    '   txt_fecha_devengo.SetFocus
'    '   Exit Function
'    'End If
'    'If CDate(txt_fecha_devengo) > CDate(lbl_cotizacion) Then
'    '    X = MsgBox("Error", "La fecha de devengamiento no puede ser mayor que la fecha de cotización", 16)
'    '    Exit Function
'    'End If
'    'Iaap = CInt(Year(CDate(txt_fecha_devengo))) 'a Fecha de siniestro
'    'Immp = CInt(Month(CDate(txt_fecha_devengo))) 'm Fecha de siniestro
'    'Iddp = CInt(Day(CDate(txt_fecha_devengo))) 'd Fecha de siniestro
'
'    Iaap = CInt(Mid(iFechaIniVig, 1, 4)) 'a Fecha de siniestro
'    Immp = CInt(Mid(iFechaIniVig, 5, 2)) 'm Fecha de siniestro
'    Iddp = CInt(Mid(iFechaIniVig, 7, 2)) 'd Fecha de siniestro
'
'    fecha_sin = Iaap * 12 + Immp
'    'sexo_cau = Trim(Mid(cbo_sexo, 1, (InStr(1, cbo_sexo, "-") - 1)))
'    'vlNum_Ben = txt_n_orden - 1 '.Rows - 1
'    vlNum_Ben = iNumBenef 'grd_beneficiarios.Rows - 1 '.Rows - 1
'
'
'    ReDim Numor(vlNum_Ben) As Integer
'    ReDim Ncorbe(vlNum_Ben) As Integer
'    ReDim Codrel(vlNum_Ben) As Integer
'    ReDim Cod_Grfam(vlNum_Ben) As Integer
'    ReDim Sexobe(vlNum_Ben) As String
'    ReDim Inv(vlNum_Ben) As String
'    ReDim Coinbe(vlNum_Ben) As String
'    ReDim Derpen(vlNum_Ben) As Integer
'    ReDim Nanbe(vlNum_Ben) As Integer
'    ReDim Nmnbe(vlNum_Ben) As Integer
'    ReDim Ndnbe(vlNum_Ben) As Integer
'    ReDim Hijos(vlNum_Ben) As Integer
'    ReDim Hijos_Inv(vlNum_Ben) As Integer
'    ReDim Hijo_Menor(vlNum_Ben) As Date
'    ReDim Hijo_Menor_Ant(vlNum_Ben) As Date
'    ReDim Porben(vlNum_Ben) As Double
'    ReDim Codcbe(vlNum_Ben) As String
'    ReDim Fec_NacHM(vlNum_Ben) As Date
'    ReDim cont_mhn(vlNum_Ben) As Integer
'
'    'tb_cau!cod_sexo
'    vlContBen = 1 '0
'    i = 1
'    Do While i <= vlNum_Ben
'
'        vlContBen = vlContBen + 1
'        'Nº Orden
'
'        'Msf_GrillaBenef.Row = i
'        'Msf_GrillaBenef.Col = 0
'        'Numor(i) = Msf_GrillaBenef.Text ''tb_ben!cod_numordben 'N° de orden  NUMOR(I)
'
'        'If Trim(grd_beneficiarios.TextMatrix(i, 0)) = "" Then
'        If Trim(ostBeneficiarios(i).Num_Orden) = "" Then
'            Exit Do
'        End If
'
'        'Número de Orden
'        Numor(i) = ostBeneficiarios(i).Num_Orden
'
'        'Parentesco
'        Ncorbe(i) = ostBeneficiarios(i).Cod_Par ''tb_ben!cod_par
'        Codrel(i) = ostBeneficiarios(i).Cod_Par ''tb_ben!cod_par
'
'        'Grupo Familiar
'        Cod_Grfam(i) = ostBeneficiarios(i).Cod_GruFam
'
'        'Sexo
'        Sexobe(i) = ostBeneficiarios(i).Cod_Sexo
'        If (Ncorbe(i) = "99") Then
'            sexo_cau = Sexobe(i)
'        End If
'
'        'Situación de Invalidez
'        Inv(i) = ostBeneficiarios(i).Cod_SitInv
'
'        'Derecho Pensión
'        Derpen(i) = ostBeneficiarios(i).Cod_EstPension
'
'        'Fecha de Nacimiento
'        Nanbe(i) = CInt(Mid(ostBeneficiarios(i).Fec_NacBen, 1, 4)) 'a Fecha de nacimiento
'        Nmnbe(i) = CInt(Mid(ostBeneficiarios(i).Fec_NacBen, 5, 2)) 'm Fecha de nacimiento
'        Ndnbe(i) = CInt(Mid(ostBeneficiarios(i).Fec_NacBen, 7, 2)) 'd Fecha de nacimiento
'
'        'Fecha nacimiento hijo menor =IJAM(I),IJMN(I),IJDN(I)
'
'        'Codificación de Situación de Invalidez
'        If Inv(i) = "P" Then Coinbe(i) = "P"
'        If Inv(i) = "T" Then Coinbe(i) = "T"
'        If Inv(i) = "N" Then Coinbe(i) = "N"
'
'        '*********
'        edad_mes_ben = fecha_sin - (Nanbe(i) * 12 + Nmnbe(i))
'        'If edad_mes_ben > L24 And Coinbe(i) = "N" And _
'            (Codrel(i) >= 30 And Codrel(i) < 40) Then
'        '    Derpen(i) = 10
'        '    ostBeneficiarios(i).Cod_EstPension = 10
'        'Else
'            Derpen(i) = 99
'            'ostBeneficiarios.Cod_EstPension = 99
'        'End If
'        i = i + 1
'    Loop
'
'    cont_causante = 0
'    cont_esposa = 0
'    'cont_mhn = 0
'    cont_mhn_tot = 0
'    cont_hijo = 0
'    cont_padres = 0
'
'    'Primer Ciclo
'    For g = 1 To vlNum_Ben
'        If Derpen(g) <> 10 Then '
'            '99: con derecho a pension,20: con Derecho Pendiente
'            '10: sin derecho a pension
'            If Ncorbe(g) = 99 Then
'                cont_causante = cont_causante + 1
'            End If
'            If Ncorbe(g) <> 99 Then
'                Select Case Ncorbe(g)
'                    Case 10, 11
'                        cont_esposa = cont_esposa + 1
'                    Case 20, 21
'                        Q = Cod_Grfam(g)
'                        cont_mhn(Q) = cont_mhn(Q) + 1
'                        cont_mhn_tot = cont_mhn_tot + 1
'                    Case 30
'                        'edad = fgEdadBen(vg_fecsin, vgBen(g).fec_nacben)
'                        Q = Cod_Grfam(g)
'                        'q = ncorbe(g)
'                        Hijos(Q) = Hijos(Q) + 1
'                        If Coinbe(g) <> "N" Then Hijos_Inv(Q) = Hijos_Inv(Q) + 1
'                        Hijo_Menor(Q) = DateSerial(Nanbe(g), Nmnbe(g), Ndnbe(g))
'                        'hijo_menor_ant = hijo_menor(q)
'                        If Hijos(Q) > 1 Then
'                            If Hijo_Menor(Q) > Hijo_Menor_Ant(Q) Then
'                                Hijo_Menor_Ant(Q) = Hijo_Menor(Q)
'                            End If
'                        Else
'                            Hijo_Menor_Ant(Q) = Hijo_Menor(Q)
'                        End If
'                    Case 35
'                        edad_mes_ben = fecha_sin - (Nanbe(g) * 12 + Nmnbe(g))
'                        If Coinbe(g) = "P" And edad_mes_ben <= L24 Then
'                            cont_hijo = cont_hijo + 1
'                        Else
'                            If Coinbe(g) = "T" Or Coinbe(g) = "N" Then
'                                cont_hijo = cont_hijo + 1
'                            End If
'                        End If
'
'                    Case 41, 42
'                        cont_padres = cont_padres + 1
'                    Case Else
'                        X = MsgBox("Error en codificación de codigo de relación", vbCritical)
'                        Exit Function
'                End Select
'            End If
'        End If
'    Next g
'
'    j = 1
'    For j = 1 To vlNum_Ben
'        '99: con derecho a pension,20: con Derecho Pendiente
'        '10: sin derecho a pension
'        If Derpen(j) <> 10 Then
'            edad_mes_ben = fecha_sin - (Nanbe(j) * 12 + Nmnbe(j))
'            Select Case Ncorbe(j)
'                Case 99
'                    If cont_causante > 1 Then
'                        X = MsgBox("Error en codificación de codigo de relación, No puede ingresar otro causante", vbCritical)
'                        Exit Function
'                    End If
'                    'I--- ABV 25/02/2005 ---
'                    'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
'                    If (fgObtenerPorcentaje(CStr(Ncorbe(j)), Coinbe(j), Sexobe(j), iFechaIniVig) = True) Then
'                        vlValor = vgValorPorcentaje
'                    Else
'                        MsgBox "No existe valor de Porcentaje para el parámetro de :" _
'                        & Ncorbe(j) & " - " & Coinbe(j) & " - " & Sexobe(j) & " - " & iFechaIniVig & "."
'                        Exit Function
'                    End If
'                    'F--- ABV 25/02/2005 ---
'                    If (vlValor < 0) Then
'                        MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
'                        'error
'                        Exit Function
'                    Else
'                        Porben(j) = vlValor
'                        'Porben(j) = 100
'                        Codcbe(j) = "N"
'                    End If
'                Case 10, 11
'                    'I--- ABV 25/02/2005 ---
'                    'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
'                    If (fgObtenerPorcentaje(CStr(Ncorbe(j)), Coinbe(j), Sexobe(j), iFechaIniVig) = True) Then
'                        vlValor = vgValorPorcentaje
'                    Else
'                        MsgBox "No existe valor de Porcentaje para el parámetro de :" _
'                        & Ncorbe(j) & " - " & Coinbe(j) & " - " & Sexobe(j) & " - " & iFechaIniVig & "."
'                        Exit Function
'                    End If
'                    'F--- ABV 25/02/2005 ---
'                    If (vlValor < 0) Then
'                        MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
'                        'error
'                        Exit Function
'                    Else
'                        If sexo_cau = "M" Then
'                            If Sexobe(j) <> "F" Then
'                                X = MsgBox("Error de código de sexo, el Sexo de la Cónyuge debe ser Femenino.", vbCritical)
'                                Exit Function
'                            End If
'                        Else
'                            If Sexobe(j) <> "M" Then
'                                X = MsgBox("Error de codigo de sexo, el Sexo del Cónyuge debe ser Masculino.", vbCritical)
'                                Exit Function
'                            End If
'                        End If
'                        If sexo_cau = "M" Then
'                            'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
'                            'If (vlValor < 0) Then
'                            '    Exit Sub
'                            'Else
'                                Porben(j) = CDbl(Format(vlValor / cont_esposa, "#0.00"))
'                            'End If
'                            u = Cod_Grfam(j)
'                            If Hijos(u) > 0 Then
'                                Codcbe(j) = "S"
'                                If Hijos_Inv(u) > 0 Then Codcbe(j) = "N"
'                            End If
'
'                            If Hijos(u) > 0 And Ncorbe(j) = 10 Then
'                                X = MsgBox("Error de código de relación, 'Cónyuge Sin Hijos con Dº Pensión', tiene Hijos.", vbCritical)
'                                Exit Function
'                            End If
'                            If Hijos(u) = 0 And Ncorbe(j) = 11 Then
'                                X = MsgBox("Error de código de relación, 'Cónyuge Con Hijos con Dº Pensión', no tiene Hijos.", vbCritical)
'                                Exit Function
'                            End If
'                        Else
'                            u = Cod_Grfam(j)
'                            If Hijos(u) > 0 Then
'                                'I--- ABV 25/02/2005 ---
'                                'If Coinbe(j) = "T" Then
'                                '    Porben(j) = vlValor     '50
'                                'Else
'                                '    'HQR 16-06-2004
'                                '    'Porben(j) = 36
'                                '    If Coinbe(j) = "P" Then
'                                '        Porben(j) = 36
'                                '    Else
'                                '        Porben(j) = 0
'                                '        cont_esposa = cont_esposa - 1
'                                '    End If
'                                '    'FIN HQR 16-06-2004
'                                'End If
'
'                                Porben(j) = vlValor
'                                If (Coinbe(j) = "N") Then
'                                    cont_esposa = cont_esposa - 1
'                                End If
'                                'F--- ABV 25/02/2005 ---
'
'                                Codcbe(j) = "S"
'                                If Hijos_Inv(u) > 0 Then Codcbe(j) = "N"
'                            Else
'                                'I--- ABV 25/02/2005 ---
'                                'If Coinbe(j) = "T" Then
'                                '    Porben(j) = vlValor     '60
'                                'Else
'                                '    'HQR 16-06-2004
'                                '    'Porben(j) = 43
'                                '    If Coinbe(j) = "P" Then
'                                '        Porben(j) = 43
'                                '    Else
'                                '        Porben(j) = 0
'                                '        cont_esposa = cont_esposa - 1
'                                '    End If
'                                '    'FIN HQR 16-06-2004
'                                'End If
'
'                                Porben(j) = vlValor
'                                If (Coinbe(j) = "N") Then
'                                    cont_esposa = cont_esposa - 1
'                                End If
'                                'F--- ABV 25/02/2005 ---
'
'                                Codcbe(j) = "S"
'                                If Hijos_Inv(u) > 0 Then Codcbe(j) = "N"
'                            End If
'                        End If
'                    End If
'                Case 20, 21
'                    If sexo_cau = "M" Then
'                        If Sexobe(j) <> "F" Then
'                            X = MsgBox("Error de código de sexo, el Sexo de la Cónyuge debe ser Femenino.", vbCritical)
'                            Exit Function
'                        End If
'                    Else
'                        X = MsgBox("Error de codigo de sexo, el Sexo del Cónyuge debe ser Masculino.", vbCritical)
'                        Exit Function
'                    End If
'
'                    'I--- ABV 25/02/2005 ---
'                    'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
'                    If (fgObtenerPorcentaje(CStr(Ncorbe(j)), Coinbe(j), Sexobe(j), iFechaIniVig) = True) Then
'                        vlValor = vgValorPorcentaje
'                    Else
'                        MsgBox "No existe valor de Porcentaje para el parámetro de :" _
'                        & Ncorbe(j) & " - " & Coinbe(j) & " - " & Sexobe(j) & " - " & iFechaIniVig & "."
'                        Exit Function
'                    End If
'                    'F--- ABV 25/02/2005 ---
'                    If (vlValor < 0) Then
'                        MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
'                        Exit Function
'                    Else
'                        Porben(j) = vlValor / cont_mhn_tot
'                    End If
'
'                    u = Cod_Grfam(j)
'                    If Hijos(u) > 0 Then
'                        Codcbe(j) = "S"
'                    Else
'                        Codcbe(j) = "N"
'                    End If
'                    If Hijos_Inv(u) > 0 Then
'                        Codcbe(j) = "N"
'                    End If
'                    If Hijos(u) > 0 And Ncorbe(j) = 20 Then
'                        X = MsgBox("Error en código de relación 'Madre Sin Hijos con Dº Pensión', tiene Hijos.", vbCritical)
'                        Exit Function
'                    End If
'                    If Hijos(u) = 0 And Ncorbe(j) = 21 Then
'                        X = MsgBox("Error en código de relación 'Madre Con Hijos con Dº Pensión, no tiene Hijos.", vbCritical)
'                        Exit Function
'                    End If
'
'                Case 30
'                    Codcbe(j) = "N"
'                    Q = Cod_Grfam(j)
'
'                    If cont_esposa > 0 Or cont_mhn(Q) > 0 Then
'                        If Coinbe(j) = "N" And edad_mes_ben > L24 Then
'                            Porben(j) = 0
'                        Else
'                            'I--- ABV 25/02/2005 ---
'                            'If Coinbe(j) = "P" And edad_mes_ben > L24 Then
'                            If (Coinbe(j) = "P" Or Coinbe(j) = "T") And edad_mes_ben > L24 Then
'                            'F--- ABV 25/02/2005 ---
'
'                                'Porben(j) = 11
'                                'I--- ABV 25/02/2005 ---
'                                'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
'                                If (fgObtenerPorcentaje(CStr(Ncorbe(j)), Coinbe(j), Sexobe(j), iFechaIniVig) = True) Then
'                                    vlValor = vgValorPorcentaje
'                                Else
'                                    MsgBox "No existe valor de Porcentaje para el parámetro de :" _
'                                    & Ncorbe(j) & " - " & Coinbe(j) & " - " & Sexobe(j) & " - " & iFechaIniVig & "."
'                                    Exit Function
'                                End If
'
'                                If (vlValor < 0) Then
'                                    MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
'                                    'X = MsgBox("Error: Código de cónyuge no esta incorporado", vbCritical)
'                                    Exit Function
'                                Else
'                                    Porben(j) = vlValor
'                                End If
'                                'F--- ABV 25/02/2005 ---
'
'                            Else
'
'                                'I--- ABV 25/02/2005 ---
'                                If (fgObtenerPorcentaje(CStr(Ncorbe(j)), "N", Sexobe(j), iFechaIniVig) = True) Then
'                                    'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
'                                    vlValor = vgValorPorcentaje
'                                Else
'                                    MsgBox "No existe valor de Porcentaje para el parámetro de :" _
'                                    & Ncorbe(j) & " - N - " & Sexobe(j) & " - " & iFechaIniVig & "."
'                                    Exit Function
'                                End If
'                                'F--- ABV 25/02/2005 ---
'
'                                If (vlValor < 0) Then
'                                    'X = MsgBox("Error: Código de cónyuge no esta incorporado", vbCritical)
'                                    MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
'                                    Exit Function
'                                Else
'                                    'Porben(j) = 15
'                                    Porben(j) = vlValor
'                                End If
'                            End If
'                        End If
'                        'If cont_esposa = 0 Or cont_mhn > 0 Then
'                    Else
'                        X = MsgBox("Error: Los códigos de beneficiarios de hijos estan mal ingresados", vbCritical)
'                        Exit Function
'                    End If
'                Case 35
'                    Q = Cod_Grfam(j)
'                    Codcbe(j) = "N"
'                    If cont_esposa = 0 And cont_mhn(Q) = 0 Then
'                        'I--- ABV 25/02/2005 ---
'                        'If Coinbe(j) = "P" And edad_mes_ben > L24 Then
'                        If (Coinbe(j) = "P" Or Coinbe(j) = "T") And edad_mes_ben > L24 Then
'                        'F--- ABV 25/02/2005 ---
'                            'Porben(j) = 11
'                            'I--- ABV 25/02/2005 ---
'                            'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
'                            If (fgObtenerPorcentaje(CStr(Ncorbe(j)), Coinbe(j), Sexobe(j), iFechaIniVig) = True) Then
'                                vlValor = vgValorPorcentaje
'                            Else
'                                MsgBox "No existe valor de Porcentaje para el parámetro de :" _
'                                & Ncorbe(j) & " - " & Coinbe(j) & " - " & Sexobe(j) & " - " & iFechaIniVig & "."
'                                Exit Function
'                            End If
'
'                            If (vlValor < 0) Then
'                                'X = MsgBox("Error: Código de cónyuge no esta incorporado", vbCritical)
'                                MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
'                                Exit Function
'                            Else
'                                Porben(j) = vlValor
'                            End If
'                            'F--- ABV 25/02/2005 ---
'                        Else
'                            'Porben(j) = 15
'                            'A = L24
'
'                            'I--- ABV 25/02/2005 ---
'                            If (fgObtenerPorcentaje(CStr(Ncorbe(j)), "N", Sexobe(j), iFechaIniVig) = True) Then
'                                'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
'                                vlValor = vgValorPorcentaje
'                            Else
'                                MsgBox "No existe valor de Porcentaje para el parámetro de :" _
'                                & Ncorbe(j) & " - N - " & Sexobe(j) & " - " & iFechaIniVig & "."
'                                Exit Function
'                            End If
'                            'F--- ABV 25/02/2005 ---
'
'                            If (vlValor < 0) Then
'                                MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
'                                'X = MsgBox("Error: Código de cónyuge no esta incorporado", vbCritical)
'                                Exit Function
'                            Else
'                                'Porben(j) = 15
'                                Porben(j) = vlValor
'                            End If
'                        End If
'
'                        'Sql = "select prc_par from porpar where cod_par = 11"
'                        'Set tb_por = vlConBD.Execute(Sql)
'                        'If Not tb_por.EOF Then
'                        '    v_hijo = tb_por!prc_par
'                        '    If coinbe(J) = "P" And edad_mes_ben <= l24 Then
'                        '        porben(J) = v_hijo / cont_hijo + 15
'                        '    Else
'                        '        If coinbe(J) = "T" Or coinbe(J) = "N" Then
'                        '            porben(J) = v_hijo / cont_hijo + 15
'                        '        End If
'                        '    End If
'                        'Else
'                        '    x = MsgBox("Error: Código de cónyuge no esta incorporado", vbCritical)
'                        'End If
'
'                        'Obtener el Porcentaje de la Cónyuge
'                        'I--- ABV 25/02/2005 ---
'                        If (fgObtenerPorcentaje("11", "N", "F", iFechaIniVig) = True) Then
'                            'vlValor = fgValorPorcentaje(1, j, 11)
'                            vlValor = vgValorPorcentaje
'                        Else
'                            MsgBox "No existe valor de Porcentaje para el parámetro de :" _
'                            & "11 - N - F  - " & iFechaIniVig & "."
'                            Exit Function
'                        End If
'                        'F--- ABV 25/02/2005 ---
'
'                        If (vlValor < 0) Then
'                            'X = MsgBox("Error, el porcentaje para la 'Cónyuge Con Hijos' no se encuentra.", vbCritical)
'                            MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & "11" & ".", vbCritical, "Error de Datos"
'                            Exit Function
'                        Else
'                            v_hijo = vlValor
'                            If Coinbe(j) = "P" And edad_mes_ben <= L24 Then
'                                Porben(j) = v_hijo / cont_hijo + 15
'                            Else
'                                If Coinbe(j) = "T" Or Coinbe(j) = "N" Then
'                                    Porben(j) = v_hijo / cont_hijo + 15
'                                End If
'                            End If
'                        End If
'
'                    Else
'                        X = MsgBox("Error: Los códigos de beneficiarios de hijos estan mal ingresados", vbCritical)
'                        Exit Function
'                    End If
'
'                Case 41, 42
'                    'I--- ABV 25/02/2005 ---
'                    If (fgObtenerPorcentaje(CStr(Ncorbe(j)), Coinbe(j), Sexobe(j), iFechaIniVig) = True) Then
'                        'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
'                        vlValor = vgValorPorcentaje
'                    Else
'                        MsgBox "No existe valor de Porcentaje para el parámetro de :" _
'                        & Ncorbe(j) & " - " & Coinbe(j) & " - " & Sexobe(j) & " - " & iFechaIniVig & "."
'                        Exit Function
'                    End If
'                    'F--- ABV 25/02/2005 ---
'
'                    If (vlValor < 0) Then
'                        MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
'                        Exit Function
'                    Else
'                        Codcbe(j) = "N"
'                        Porben(j) = vlValor
'                    End If
'                'Case 42
'                '    Codcbe(j) = "N"
'                '    Porben(j) = 50
'            End Select
'        End If
'    Next j
'
'    For k = 1 To vlNum_Ben '60
'        If Derpen(k) <> 10 Then
'            Select Case Ncorbe(k)
'                Case 11
'                    'q = codrel(k)
'                    Q = Cod_Grfam(k)
'                    If Codcbe(k) = "S" Then
'                        Fec_NacHM(k) = Hijo_Menor_Ant(Q)  '******Agruegué 05/11/2000 HILDA
'                    End If
'                Case 21
'                    'q = codrel(k)
'                    Q = Cod_Grfam(k)
'                    If Codcbe(k) = "S" Then
'                        Fec_NacHM(k) = Hijo_Menor_Ant(Q)  '*******Agruegué 05/11/2000 HILDA
'                    End If
'            End Select
'        End If
'    Next k
'
'    For j = 1 To (vlContBen - 1)
'        'Guardar el Valor del Porcentaje Calculado
'        'If IsNumeric(Porben(j)) Then
'        '    grd_beneficiarios.Text = Format(Porben(j), "##0.00")
'        'Else
'        '    grd_beneficiarios.Text = Format("0", "0.00")
'        'End If
'        If IsNumeric(Porben(j)) Then
'            ostBeneficiarios(j).Prc_Pension = Format(Porben(j), "#0.00")
'        Else
'            ostBeneficiarios(j).Prc_Pension = 0
'        End If
'
'        'Guardar el Derecho a Acrecer de los Beneficiarios
'        'Inicio
'        'If (Codcbe(j) <> Empty And Not IsNull(Codcbe(j))) Then
'        '    grd_beneficiarios.Text = Codcbe(j)
'        'Else
'        '    'Por Defecto Negar el Derecho a Acrecer de los Beneficiarios
'        '    grd_beneficiarios.Text = "N"
'        'End If
'        If (Codcbe(j) <> Empty And Not IsNull(Codcbe(j))) Then
'            ostBeneficiarios(j).Cod_DerCre = Codcbe(j)
'        Else
'            ostBeneficiarios(j).Cod_DerCre = "N"
'        End If
'        'Fin
'
'        'Guardar la Fecha de Nacimiento del Hijo Menor de la Cónyuge
'        'Inicio
'        ''If Format((Fec_NacHM(J)), "yyyy/mm/dd") > "1899/12/30" Then
'        'If Format(CDate(Fec_NacHM(j)), "yyyymmdd") > "18991230" Then
'        '    grd_beneficiarios.Text = CDate(Fec_NacHM(j))
'        'Else
'        '    'Guardar la Fecha de Nacimiento del Hijo Menor de la Cónyuge
'        '    grd_beneficiarios.Text = ""
'        'End If
'        If Format(CDate(Fec_NacHM(j)), "yyyymmdd") > "18991230" Then
'            ostBeneficiarios(j).Fec_NacHM = Fec_NacHM(j)
'        Else
'            ostBeneficiarios(j).Fec_NacHM = ""
'        End If
'        'Fin
'    Next j
'
'    'vll_numorden = grd_beneficiarios.Rows
'    'txt_n_orden.Caption = vll_numorden
'
'    fgCalcularPorcentajeBenef = True
'
'Exit Function
'Err_fgCalcularPorcentaje:
'    Screen.MousePointer = 0
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
'End Function
'
'Function fgCarga_Param(iTabla As String, iElemento As String, iFecha As String) As Boolean
'Dim vlRegistro As ADODB.Recordset
'
'    'Call AbrirBaseDeDatos_Aux(vgRutaBasedeDatos)
''''    If Not AbrirBaseDeDatos(vgConexionParam) Then
''''        MsgBox "Falló la conexión con la Base de Datos.", vbCritical, "Error de Conexión"
''''        Exit Function
''''    End If
'
'    fgCarga_Param = False
'    vgValorParametro = 0
'
'    vgSql = "SELECT mto_elemento FROM MA_TPAR_TABCODVIG WHERE "
'    vgSql = vgSql & "COD_TABLA = '" & iTabla & "' and "
'    vgSql = vgSql & "COD_ELEMENTO = '" & iElemento & "' "
'    vgSql = vgSql & "AND (FEC_INIVIG <= '" & iFecha & "' "
'    vgSql = vgSql & "AND FEC_TERVIG >= '" & iFecha & "') "
'    Set vlRegistro = vgConexionBD.Execute(vgSql)
'    'If Not (vlRegistro.EOF) Then
'    If Not (vlRegistro.EOF) Then
'        If Not IsNull(vlRegistro!Mto_Elemento) Then
'            vgValorParametro = Trim(vlRegistro!Mto_Elemento)
'
'            fgCarga_Param = True
'        End If
'    End If
'    vlRegistro.Close
'
''''    Call CerrarBaseDeDatos(vgConexionParam)
'End Function
'
'Function fgObtenerPorcentaje(iParentesco As String, iInvalidez As String, iSexo As String, iFecha As String) As Boolean
''Función : Permite validar la existencia del valor del Porcentaje de Pensión a buscar
''Parámetros de Entrada:
''       - iParentesco => Código del Parentesco del Beneficiario
''       - iInvalidez  => Código de la Situación de Invalidez del Beneficiario
''       - iSexo       => Código del Sexo del Beneficiario
''       - iFecha      => Fecha con la cual se compara la Vigencia del Porcentaje (Vigencia Póliza)
''Parámetros de Salida:
''       - Retorna un Falso o True de acuerdo a su existencia
''Variables de Salida:
''       - vgValorPorcentaje => Permite guardar el Porcentaje buscado
'Dim Tb_Por As ADODB.Recordset
'Dim Sql    As String
'
'    fgObtenerPorcentaje = False
'
'    Sql = "select prc_pension as valor_porcentaje "
'    Sql = Sql & "from MA_TVAL_PORPAR where "
'    Sql = Sql & "Cod_par = '" & iParentesco & "' AND "
'    Sql = Sql & "Cod_sitinv = '" & iInvalidez & "' AND "
'    Sql = Sql & "Cod_sexo = '" & iSexo & "' AND "
'    Sql = Sql & "fec_inivigpor <= '" & iFecha & "' AND "
'    Sql = Sql & "fec_tervigpor >= '" & iFecha & "' "
'
'    Set Tb_Por = vgConexionBD.Execute(Sql)
'    If Not Tb_Por.EOF Then
'
'        If Not IsNull(Tb_Por!Valor_Porcentaje) Then
'            vgValorPorcentaje = Tb_Por!Valor_Porcentaje
'
'            fgObtenerPorcentaje = True
'        End If
'    End If
'    Tb_Por.Close
'
'End Function


Public Function fgBuscarMonedaOfiTran(oMonedaOfi As String, oMonedaTran As String)
'Función: Permite buscar el Código de la Moneda a expresar los valores, y
'el Código de la Moneda en que se deben transformar los valores
'Parámetros de Entrada:
'Parámetros de Salida:
'- oMonedaOfi   => Código de la Moneda Oficial
'- oMonedaTran  => Código de la Moneda a Transformar
'----------------------------------------------------
'Fecha Creación     : 07/07/2007
'Fecha Modificación :
'----------------------------------------------------
Dim vlRegistroMon As ADODB.Recordset
On Error GoTo Err_BuscarMoneda

'    fgBuscarMonedaOfiTran = False
    oMonedaOfi = cgCodTipMonedaUF
    oMonedaTran = cgCodTipMonedaUF
    
    vgSql = "SELECT cod_monedaofi,cod_monedatrans "
    vgSql = vgSql & "FROM ma_tcod_moneda "
    Set vlRegistroMon = vgConexionBD.Execute(vgSql)
    If Not vlRegistroMon.EOF Then
        If Not IsNull(vlRegistroMon!cod_monedaofi) Then oMonedaOfi = vlRegistroMon!cod_monedaofi
        If Not IsNull(vlRegistroMon!cod_monedatrans) Then oMonedaTran = vlRegistroMon!cod_monedatrans
'        fgBuscarMonedaOfiTran = True
    End If
    vlRegistroMon.Close
    
Exit Function
Err_BuscarMoneda:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function


'Function flCargaEstructuraPoliza(iNombreTabla As String, ipoliza As String, iEndoso As Integer, istPoliza As TyPoliza)
'On Error GoTo Err_flCargaEstructuraPoliza
'
'    vgSql = ""
'    vgSql = "SELECT num_poliza,num_endoso,cod_tippension,cod_estado, "
'    vgSql = vgSql & "cod_tipren,cod_modalidad,num_cargas,fec_vigencia, "
'    vgSql = vgSql & "fec_tervigencia,mto_prima,mto_pension,num_mesdif, "
'    vgSql = vgSql & "num_mesgar,prc_tasace,prc_tasavta,prc_tasaintpergar "
'    vgSql = vgSql & "FROM " & iNombreTabla & " WHERE "
'    vgSql = vgSql & "num_poliza = '" & Trim(ipoliza) & "' AND "
'    vgSql = vgSql & "num_endoso = " & iEndoso & " "
'    vgSql = vgSql & " ORDER BY num_endoso DESC"
'    Set vgRegistro = vgConexionBD.Execute(vgSql)
'    If Not vgRegistro.EOF Then
'       With istPoliza
'            .Num_Poliza = (vgRegistro!Num_Poliza)
'            .num_endoso = (vgRegistro!num_endoso)
'            .Cod_TipPension = (vgRegistro!Cod_TipPension)
'            .Cod_Estado = (vgRegistro!Cod_Estado)
'            .Cod_TipRen = (vgRegistro!Cod_TipRen)
'            .Cod_Modalidad = (vgRegistro!Cod_Modalidad)
'            .Num_Cargas = (vgRegistro!Num_Cargas)
'            .Fec_Vigencia = (vgRegistro!Fec_Vigencia)
'            .Fec_TerVigencia = (vgRegistro!Fec_TerVigencia)
'            .Mto_Prima = (vgRegistro!Mto_Prima)
'            .Mto_Pension = (vgRegistro!Mto_Pension)
'            .Num_MesDif = (vgRegistro!Num_MesDif)
'            .Num_MesGar = (vgRegistro!Num_MesGar)
'            .Prc_TasaCe = (vgRegistro!Prc_TasaCe)
'            .Prc_TasaVta = (vgRegistro!Prc_TasaVta)
'            .Prc_TasaIntPerGar = (vgRegistro!Prc_TasaIntPerGar)
'       End With
'    End If
'
'Exit Function
'Err_flCargaEstructuraPoliza:
'    Screen.MousePointer = 0
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
'
'End Function

'F--CMV

'------------------------------------------------------
' Función de Inicio del Sistema
'------------------------------------------------------
Public Sub p_Actualiza_Version(ByVal g_strPathExe As String, ByVal sExeUpgade As String)
    
    Dim FechaAppS As Date
    Dim FechaAppM As Date
    Dim r_FlagOk As Variant
    'Hora y fecha del exe en el servidor
    FechaAppS = Format$(FileDateTime(g_strPathExe & "Pensiones.exe"), "dd/mm/yyyy hh:mm")
    'Hora y fecha del exe local
    FechaAppM = Format$(FileDateTime(App.Path & "\Pensiones.exe"), "dd/mm/yyyy hh:mm")
        
    If FechaAppS <> FechaAppM Then
        If MsgBox("La versión del sistema en el servidor, difiere a la del computador. Desea Actualizar la versión ?", vbQuestion + vbYesNo, "") = vbYes Then
           r_FlagOk = Shell(sExeUpgade & " RutaO=" & g_strPathExe & "Pensiones.exe" & ",RutaD=" & App.Path & ",NomEx=Pensiones.exe", vbNormalFocus)
           End
        Else
           End
        End If
    End If
        
End Sub

Sub Main()
Dim inises As String
Dim usua As String
Dim pas As String

Dim strExes As String
Dim strUpgrade As String

On Error GoTo Err_Main

    inises = "sincn"
    Do Until inises = "concn"
        usua = LeeArchivoIni("Conexion", "Usuario", "", App.Path & "\AdmPrevBD.Ini")
        pas = LeeArchivoIni("Conexion", "Password", "", App.Path & "\AdmPrevBD.Ini")
        strRpt = LeeArchivoIni("REPORTES", "strPath", "", App.Path & "\Rutas.ini")
        strExes = LeeArchivoIni("EXES", "strPath", "", App.Path & "\Rutas.ini")
        strUpgrade = LeeArchivoIni("EXES", "strUpgrade", "", App.Path & "\Rutas.ini")
     
        If (usua = "" Or usua = Empty) Or (pas = "" Or pas = Empty) Then
            Frm_InicioSesion.Show 1
            If inises = "cancelar" Then End
        Else
            inises = "concn"
        End If
    Loop
   
   If strExes <> "" Then
       ' Call p_Actualiza_Version(strExes, strUpgrade)
   End If
    
    vgDsn = ""
    vgNombreServidor = ""
    vgNombreBaseDatos = ""
    vgNombreUsuario = ""
    vgPassWord = ""
    vgMensaje = ""
    vgRutaArchivo = ""
    
    'Valida Si Existe Archivo de AdmBasDat.Inicio
    vgRutaArchivo = App.Path & "\AdmPrevBD_Prod.Ini"
    'vgRutaArchivo = App.Path & "\AdmPrevBD.Ini"
    If Not fgExiste(vgRutaArchivo) Then
        MsgBox "No existe el Archivo de Parámetros para ejecutar la Aplicación.", vbCritical, "Ejecución Cancelada"
        End
    End If

    lpFileName = vgRutaArchivo
    lpAppName = "Conexion"
    lpDefault = ""
    lpReturnString = Space$(128)
    Size = Len(lpReturnString)
    lpKeyName = ""

    'Valida Si Existe Nombre de la entrada para definir al Proveedor
    ProviderName = fgGetPrivateIni(lpAppName, "Proveedor", lpFileName)
    If (ProviderName = "DESCONOCIDO") Then
        vgMensaje = "La Entrada 'Proveedor', no está definida en el Archivo AdmPrevBD.Ini" & vbCrLf
    End If
    
    'Valida Si Existe Nombre de la entrada para definir al Servidor
    vgNombreServidor = fgGetPrivateIni(lpAppName, "Servidor", lpFileName)
    If (vgNombreServidor = "DESCONOCIDO") Then
        vgMensaje = "La Entrada 'Servidor', no está definida en el Archivo AdmPrevBD.Ini" & vbCrLf
    End If

    'Valida Si Existe Nombre de la entrada para definir Base de Datos SisSin
    vgNombreBaseDatos = fgGetPrivateIni(lpAppName, "BaseDatos", lpFileName)
    If (vgNombreBaseDatos = "DESCONOCIDO") Then
        vgMensaje = "La Entrada 'Base de Datos', no está definida en el Archivo AdmPrevBD.Ini" & vbCrLf
    End If
    
    'Valida Si Existe Nombre de la entrada para definir Usuario de SisSin
    vgNombreUsuario = fgGetPrivateIni(lpAppName, "Usuario", lpFileName)
    If (vgNombreUsuario = "DESCONOCIDO") Then
        vgMensaje = "La Entrada 'Usuario', no está definida en el Archivo AdmPrevBD.Ini" & vbCrLf
    End If
    vgNombreUsuario = fgDesPassword(vgNombreUsuario)
    
    'Valida Si Existe Nombre de la entrada para definir Password de SisSin
    vgPassWord = fgGetPrivateIni(lpAppName, "Password", lpFileName)
    'vgPassWord = ""
    If (vgPassWord = "DESCONOCIDO") Then
        vgMensaje = "La Entrada 'PassWord', no está definida en el Archivo AdmPrevBD.Ini" & vbCrLf
    Else
        If (vgPassWord = "DESCONOCIDO") And (UCase(vgNombreUsuario) = "SA") Then
            vgPassWord = ""
        End If
    End If
    
    vgPassWord = "rentcalidad64" 'fgDesPassword(vgPassWord)
    'vgPassWord = "rentcalidad64"

    'Valida Si Existe Nombre de la entrada para definir DSN de SisSin
    vgDsn = fgGetPrivateIni(lpAppName, "DSN", lpFileName)
    If (vgDsn = "DESCONOCIDO") Then
        vgMensaje = "La Entrada 'DSN', no está definida en el Archivo AdmPrevBD.Ini" & vbCrLf
    End If

    If (vgMensaje <> "") Then
        MsgBox "Status de los Datos de Inicio" & vbCrLf & vbCrLf & vgMensaje & vbCrLf & vbCrLf & "Proceso Cancelado." & vbCrLf & "Se deben Ingresar todos los datos Básicos."
        'Exit Sub
        End
    End If

    'vgRutaBasedeDatos = LeeArchivoIni("Conexion", "Ruta", "", App.Path & "\AdmPrevBD.Ini")
    'vgRutaBasedeDatos = vgRutaBasedeDatos & LeeArchivoIni("Conexion", "BasedeDatos", "", App.Path & "\AdmPrevBD.Ini")
    ''AbrirBaseDeDatos (vgRutaBasedeDatos)

    If Not fgConexionBaseDatos(vgConexionBD) Then
        MsgBox "Falló la conexión con la Base de Datos.", vbCritical, "Error de Conexión"
        'Exit Sub
        End
    End If
    'vgConBD.Close
    
    'Call CerrarBaseDeDatos

    'Ruta de la Base de Datos
    vgRutaDataBase = "ODBC;DSN=" & vgDsn & ";UID=" & vgNombreUsuario & ";PWD=" & vgPassWord & ";DATABASE=" & vgNombreBaseDatos & ";"
    vgRutaBasedeDatos = vgRutaDataBase
    
    '*******************************************************
    'Sacar luego cuando sean ingresados los datos que faltan
    '*******************************************************
    'Determinar los Datos del Cliente
    vgQuery = "SELECT * "
    vgQuery = vgQuery & "FROM ma_tmae_cliente "
    Set vgRs = vgConexionBD.Execute(vgQuery)
    If Not vgRs.EOF Then
        If Not IsNull(vgRs!GLS_NOMCLI) Then vgNombreCompania = Trim(vgRs!GLS_NOMCLI)
        If Not IsNull(vgRs!gls_nomcorcli) Then vgNombreCortoCompania = Trim(vgRs!gls_nomcorcli)
        If Not IsNull(vgRs!num_idencli) Then vgNumIdenCompania = vgRs!num_idencli
        If Not IsNull(vgRs!cod_tipoidencli) Then vgTipoIdenCompania = Trim(vgRs!cod_tipoidencli)
    Else
        vgNumIdenCompania = ""
        vgTipoIdenCompania = ""
        vgNombreCompania = ""
        vgNombreCortoCompania = ""
    End If
    vgRs.Close
    
    'Determinar los Datos del Sistema
    vgQuery = "select * from ma_tpar_sistema where "
    vgQuery = vgQuery & "cod_sistema = '" & vgTipoSistema & "'"
    Set vgRs = vgConexionBD.Execute(vgQuery)
    If Not vgRs.EOF Then
        If Not IsNull(vgRs!gls_sistema) Then vgNombreSubSistema = Trim(vgRs!gls_sistema)
    Else
        vgNombreSubSistema = ""
    End If
    vgRs.Close
    '*******************************************************
    
     'RRR
        vgSql = "SELECT * FROM MA_TMAE_ADMINCUENTAS WHERE "
        vgSql = vgSql & "cod_cliente = '1' "
        Set vgRs = vgConexionBD.Execute(vgSql)

                    If Not vgRs.EOF Then
                        vgIntentos = vgRs!nintentos
                        vgChkdiaant = vgRs!bdiasvence
                        vgDiasFaltan = vgRs!ndiasvence
                    End If
        vgRs.Close
    'RRR
    
    
    'Cerrar la Conexión
'    vgConexionBD.Close
    
    'Ruta de la Base de Datos
    'vgRutaDataBase = "ODBC;DSN=" & vgDsn & ";UID=" & vgNombreUsuario & ";PWD=" & vgPassWord & ";DATABASE=" & vgNombreBaseDatos & ";"
    'vgRutaDataBase = "ODBC;DSN=" & vgDsn & ";DATABASE=" & vgNombreBaseDatos & ";"
    'vgRutaDataBase = "ODBC;UID="";PWD="";DATABASE=" & vgNombreBaseDatos & ";DSN=" & vgDsn & ";"
    'vgRutaBasedeDatos = vgRutaDataBase
    
    '****************************
    'Inicio: Sacar Comentarios
    '****************************
    Screen.MousePointer = 11
    Frm_Menu.Show
    'I----- Debe validarse ABV 21/01/2004 ---
    'Activarlo cuando se encuentre el Menú corregido y activado
    'I------ CMV 07/09/2004 -----
    Frm_Menu.Mnu_AdmSistema.Enabled = False
    Frm_Menu.mnuMantenciondeInformacion.Enabled = False
    Frm_Menu.mnuCalculoPension.Enabled = False
    Frm_Menu.mnuRegistroPagosATerceros.Enabled = False
    Frm_Menu.mnuConsultas.Enabled = False 'HQR 05/11/2005
    Frm_Menu.mnuAcerca.Enabled = False
    Frm_Menu.mnuSalir.Enabled = False
    'F------ CMV 07/09/2004 -----
    'I----- Debe validarse ABV 21/01/2004 ---
    Screen.MousePointer = 0

    Screen.MousePointer = 11
    Frm_SisPassword.Show
    Screen.MousePointer = 0
    '****************************
    'Fin  : Sacar Comentarios
    '****************************
    
    fgCargarVariablesGlobales
Exit Sub
Err_Main:
    Screen.MousePointer = 0
    Select Case Err
    Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
        End
    End Select
End Sub

'----------------------------------------------------------
'Función que permite centralizar el Formulario o Pantalla
'----------------------------------------------------------
Sub Center(Frm As Form)
    Frm.Left = (Screen.Width - Frm.Width) \ 2
    Frm.Top = (Screen.Height - (Frm.Height + 2000)) \ 2
End Sub



Function fgComboNivelGlosa(vlCombo As ComboBox)
Dim vlGlosa As String
On Error GoTo Err_Combo

    'Call AbrirBaseDeDatos_Aux(vgRutaBasedeDatos)
'''    If Not AbrirBaseDeDatos(vgConectarBD) Then
'''        MsgBox "Falló la conexión con la Base de Datos.", vbCritical, "Error de Conexión"
'''        Exit Function
'''    End If
    
    vlCombo.Clear
    vgQuery = "SELECT cod_nivel as codigo "
    vgQuery = vgQuery & ",gls_nivel as glosa "
    vgQuery = vgQuery & "from "
    vgQuery = vgQuery & "MA_TPAR_NIVEL WHERE "
    vgQuery = vgQuery & "cod_sistema = '" & vgTipoSistema & "' "
    vgQuery = vgQuery & "ORDER BY cod_nivel "
    Set vgCmb = vgConexionBD.Execute(vgQuery)
    If Not (vgCmb.EOF) Then
        'vgCmb.MoveFirst
        While Not (vgCmb.EOF)
            vlGlosa = ""
            If Not IsNull(vgCmb!glosa) Then vlGlosa = Trim(vgCmb!glosa)
            vlCombo.AddItem (Trim(vgCmb!Codigo) & " - " & vlGlosa)
            vgCmb.MoveNext
        Wend
        If (vlCombo.ListCount <> 0) Then
            vlCombo.ListIndex = 0
        End If
    End If
    vgCmb.Close
    
    'Call CerrarBaseDeDatos_Aux
'''     Call CerrarBaseDeDatos(vgConectarBD)
Exit Function
Err_Combo:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'--------------------------------------------------------
'Permite Encriptar la Password del Usuario que es registrada
'en la Base de Datos
'--------------------------------------------------------
Function fgEncPassword(iContraseña) As String
Dim iPassword As String
Dim iOculto As String
On Error GoTo Err_Encriptar
    
    iPassword = ""
    iOculto = ""
    For vgI = 1 To Len(UCase(iContraseña))
        iOculto = Chr(255 - Asc(UCase(Mid(iContraseña, vgI, 1))))
        Asc (iOculto)
        iPassword = iPassword + iOculto
        Chr (15)
    Next
    fgEncPassword = iPassword

Exit Function
Err_Encriptar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'--------------------------------------------------------
'Permite Desencriptar la Password del Usuario registrada
'en la Base de Datos
'--------------------------------------------------------
Function fgDesPassword(iContraseña) As String
Dim iPassword As String
Dim iOculto As String
On Error GoTo Err_Desencriptar
    
    iPassword = ""
    iOculto = ""
    For vgI = 1 To Len(UCase(iContraseña))
        iOculto = Chr(255 - Asc(UCase(Mid(iContraseña, vgI, 1))))
        Asc (iOculto)
        iPassword = iPassword + iOculto
        Chr (15)
    Next
    fgDesPassword = iPassword

Exit Function
Err_Desencriptar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function


'--------------------------------------------------------
'Permite determinar si el Archivo Existe
'--------------------------------------------------------
Function fgExiste(iarchivo) As Boolean
On Local Error GoTo Err_NoExiste
    
    If Dir$(iarchivo) = "" Then
        fgExiste = False
    Else
        fgExiste = True
    End If
    Exit Function
    
Err_NoExiste:
    fgExiste = False
    Exit Function
End Function

Function ValiRut(Rut As String, DgV As String) As Boolean
Dim vlDigitos As String
Dim vlDiv As Long, vlMul As Double
Dim vlDgv As String
Dim Resultado, y As Integer
    
    Rut = Format(Rut, "#0")
    vlDigitos = "32765432"
    Resultado = 0
    Select Case Len(Rut)
           Case 7: Rut = "0" + Rut
           Case 6: Rut = "00" + Rut
           Case 5: Rut = "000" + Rut
           Case 4: Rut = "0000" + Rut
           Case 3: Rut = "00000" + Rut
           Case 2: Rut = "000000" + Rut
           Case 1: Rut = "0000000" + Rut
    End Select
    For y = 1 To Len(Trim(Rut))
        Resultado = Resultado + Val(Mid(Rut, Len(Trim(Rut)) - y + 1, 1)) * Val(Mid(vlDigitos, Len(Trim(Rut)) - y + 1, 1))
    Next y
    vlDiv = Int(Resultado / 11)
    vlMul = vlDiv * 11
    vlDiv = Resultado - vlMul
    vlDgv = 11 - vlDiv
    If vlDgv = 11 Then
       vlDgv = 0
    End If
    If vlDgv = 10 Then
       vlDgv = "K"
    End If
    If CStr(vlDgv) = DgV Then
       ValiRut = True
    Else
       ValiRut = False
    End If
End Function

'------------------------------------------------------------'
'Permite Cargar los Distintos Combos existentes en el sistema'
'De acuerdo a los parámetros de llamada                      '
'------------------------------------------------------------'
Function fgComboGeneral(iCodigo, iCombo As ComboBox)
Dim vlRsCombo As ADODB.Recordset
On Error GoTo Err_ComboGeneral

    iCombo.Clear
    
    
        vgSql = "select cod_elemento AS CODIGO, gls_elemento AS GLOSA "
        vgSql = vgSql & "from MA_TPAR_TABCOD where "
        vgSql = vgSql & "cod_tabla = '" & iCodigo & "'"
        vgSql = vgSql & "order by cod_elemento "
    Set vlRsCombo = vgConexionBD.Execute(vgSql)
    Do While Not vlRsCombo.EOF
        If (Trim(iCodigo) = vgCodTabla_TipMon) Or (Trim(iCodigo) = vgCodTabla_ModPagoRetJud) Or _
           (Trim(iCodigo) = vgCodTabla_ModPago) Or (Trim(iCodigo) = vgCodTabla_ModPagoCC) Then
            If Trim(vlRsCombo!Codigo) = cgCodTipMonedaPESOS Then
                iCombo.AddItem (Trim(vlRsCombo!Codigo) & " - " & Trim(vlRsCombo!glosa)), 0
            Else
                iCombo.AddItem ((Trim(vlRsCombo!Codigo) & " - " & Trim(vlRsCombo!glosa)))
            End If
        Else
            iCombo.AddItem ((Trim(vlRsCombo!Codigo) & " - " & Trim(vlRsCombo!glosa)))
        End If
        vlRsCombo.MoveNext
    Loop
    vlRsCombo.Close

    If iCombo.ListCount <> 0 Then
        iCombo.ListIndex = 0
    End If

Exit Function
Err_ComboGeneral:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function
Function fgComboCausaEndosos(iCodigo, iCombo As ComboBox, opt As String)
Dim vlRsCombo As ADODB.Recordset
On Error GoTo Err_ComboGeneral

    iCombo.Clear
    vgSql = "select cod_elemento AS CODIGO, gls_elemento AS GLOSA "
    vgSql = vgSql & "from MA_TPAR_TABCOD where "
    vgSql = vgSql & "cod_tabla = '" & iCodigo & "' and cod_scomp='" & opt & "'"
    vgSql = vgSql & "order by cod_elemento "
    Set vlRsCombo = vgConexionBD.Execute(vgSql)
    Do While Not vlRsCombo.EOF
        iCombo.AddItem ((Trim(vlRsCombo!Codigo) & " - " & Trim(vlRsCombo!glosa)))
        vlRsCombo.MoveNext
    Loop
    vlRsCombo.Close

    If iCombo.ListCount <> 0 Then
        iCombo.ListIndex = 0
    End If

Exit Function
Err_ComboGeneral:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fgComboEstadoTodo(iCombo As ComboBox)
Dim vlRsCombo As ADODB.Recordset
On Error GoTo Err_fgComboEstadoTodo

    iCombo.Clear

    iCombo.AddItem ("TODOS")
    
    vgSql = "SELECT cod_dergarest AS codigo, gls_dergarest AS glosa "
    vgSql = vgSql & "FROM ma_tpar_estdergarest "
    vgSql = vgSql & "ORDER BY cod_dergarest "
    Set vlRsCombo = vgConexionBD.Execute(vgSql)
    Do While Not vlRsCombo.EOF
        iCombo.AddItem ((Trim(vlRsCombo!Codigo) & " - " & Trim(vlRsCombo!glosa)))
        vlRsCombo.MoveNext
    Loop
    vlRsCombo.Close
    
    If iCombo.ListCount <> 0 Then
        iCombo.ListIndex = 0
    End If

Exit Function
Err_fgComboEstadoTodo:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function


'------------------------------------------------------------
'Permite Cargar el Combo de Sucursales del Sistema
'------------------------------------------------------------
Function fgComboSucursal(iCombo As ComboBox, iTipo As String)
Dim vlRsCombo As ADODB.Recordset
On Error GoTo Err_ComboSucursal
    
    iCombo.Clear
    vgSql = "SELECT cod_sucursal,gls_sucursal "
    vgSql = vgSql & "FROM MA_TPAR_SUCURSAL "
    vgSql = vgSql & "WHERE cod_tipo = '" & iTipo & "' "
    vgSql = vgSql & "ORDER BY cod_sucursal "
    Set vlRsCombo = vgConexionBD.Execute(vgSql)
    Do While Not vlRsCombo.EOF
        iCombo.AddItem ((Trim(vlRsCombo!Cod_Sucursal) & " - " & Trim(vlRsCombo!gls_sucursal)))
        vlRsCombo.MoveNext
    Loop
    vlRsCombo.Close
    If iCombo.ListCount <> 0 Then
        iCombo.ListIndex = 0
    End If
    
Exit Function
Err_ComboSucursal:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'------------------------------------------------------------
'Permite Obtener la Glosa de un Código especifico de la Tabla
'de Parámetros Generales (TABCOD)
'iTabla = código de la Tabla, iElemento = código de Elemento
'------------------------------------------------------------
Function fgBuscarGlosaElemento(iTabla, iElemento) As String
Dim vlRsDescripcion As ADODB.Recordset
On Error GoTo Err_BuscarGlosa

    fgBuscarGlosaElemento = ""
    
    vgSql = ""
    vgSql = "select gls_elemento "
    vgSql = vgSql & "from MA_TPAR_TABCOD where "
    vgSql = vgSql & "cod_tabla = '" & iTabla & "' and "
    vgSql = vgSql & "cod_elemento = '" & iElemento & "' "
    Set vlRsDescripcion = vgConexionBD.Execute(vgSql)
    If Not vlRsDescripcion.EOF Then
        fgBuscarGlosaElemento = vlRsDescripcion!GLS_ELEMENTO
    End If
    vlRsDescripcion.Close
    
Exit Function
Err_BuscarGlosa:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'------------------------------------------------------------
'Permite Obtener la Posición, dentro de un Combo, del Código
'del Elemento indicado
'iElemento = código de Elemento
'------------------------------------------------------------
Function fgBuscarPosicionCodigoCombo(iElemento, iCombo As ComboBox)
Dim iContador As Long
On Error GoTo Err_BuscarPosicion

    fgBuscarPosicionCodigoCombo = -1
    
    iContador = 0
    iCombo.ListIndex = 0
    Do While iContador < iCombo.ListCount
        If (Trim(iCombo) <> "") Then
            If Trim(iElemento) = Trim(Mid(iCombo.Text, 1, (InStr(1, iCombo, "-") - 1))) Then
                fgBuscarPosicionCodigoCombo = iContador
                Exit Do
            End If
        End If
        iContador = iContador + 1
        iCombo.ListIndex = iContador
    Loop

Exit Function
Err_BuscarPosicion:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fgComboNivel(vlCombo As ComboBox)
On Error GoTo Err_Combo

    'Call AbrirBaseDeDatos_Aux(vgRutaBasedeDatos)
'''    If Not AbrirBaseDeDatos(vgConectarBD) Then
'''        MsgBox "Falló la conexión con la Base de Datos.", vbCritical, "Error de Conexión"
'''        Exit Function
'''    End If
    
    vlCombo.Clear
    vgQuery = "SELECT cod_nivel as codigo from "
    vgQuery = vgQuery & "MA_TPAR_NIVEL WHERE "
    vgQuery = vgQuery & "cod_sistema = '" & vgTipoSistema & "' "
    vgQuery = vgQuery & "ORDER BY cod_nivel "
    Set vgCmb = vgConexionBD.Execute(vgQuery)
    If Not (vgCmb.EOF) Then
        'vgCmb.MoveFirst
        While Not (vgCmb.EOF)
            vlCombo.AddItem (Trim(vgCmb!Codigo))
            vgCmb.MoveNext
        Wend
        If (vlCombo.ListCount <> 0) Then
            vlCombo.ListIndex = 0
        End If
    End If
    vgCmb.Close
    
    'Call CerrarBaseDeDatos_Aux
'''     Call CerrarBaseDeDatos(vgConectarBD)
Exit Function
Err_Combo:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'PERMITE CARGAR TODAS LAS COMUNAS
Function fgComboComuna(vlCombo As ComboBox)
Dim vlcont As Long
On Error GoTo Err_Carga
     
     vlCombo.Clear
     vgSql = ""
     vlcont = 0
     vgSql = "Select Gls_Comuna,Cod_Direccion from MA_TPAR_COMUNA ORDER BY GLS_Comuna"
     Set vgCmb = vgConexionBD.Execute(vgSql)
         If Not (vgCmb.EOF) Then
            While Not (vgCmb.EOF)
                  vlCombo.AddItem (Trim(vgCmb!gls_comuna))
                  vlcont = vlCombo.ListCount - 1
                  vlCombo.ItemData(vlcont) = (vgCmb!Cod_Direccion)
                  vgCmb.MoveNext
            Wend
         End If
         vgCmb.Close
        
     If vlCombo.ListCount <> 0 Then
        vlCombo.ListIndex = 0
     End If

Exit Function
Err_Carga:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fgBuscarNombreProvinciaRegion(vlCodDir)
Dim vlRegistroNombre As ADODB.Recordset
On Error GoTo Err_Buscar

     vgSql = "SELECT r.Cod_Region,r.Gls_Region,p.Cod_Provincia,p.Gls_Provincia,c.Cod_Comuna,c.Gls_Comuna"
     vgSql = vgSql & " FROM MA_TPAR_COMUNA c, MA_TPAR_PROVINCIA p, MA_TPAR_REGION r"
     vgSql = vgSql & " Where c.Cod_Direccion = '" & vlCodDir & "' and  "
     vgSql = vgSql & " c.cod_region = p.cod_region and"
     vgSql = vgSql & " c.cod_provincia = p.cod_provincia and"
     vgSql = vgSql & " p.cod_region = r.cod_region"
     Set vlRegistroNombre = vgConexionBD.Execute(vgSql)
     If Not vlRegistroNombre.EOF Then
        vgNombreRegion = IIf(IsNull(vlRegistroNombre!gls_region), "", vlRegistroNombre!gls_region)
        vgNombreProvincia = IIf(IsNull(vlRegistroNombre!gls_provincia), "", vlRegistroNombre!gls_provincia)
        vgNombreComuna = IIf(IsNull(vlRegistroNombre!gls_comuna), "", vlRegistroNombre!gls_comuna)
     End If
     vlRegistroNombre.Close

Exit Function
Err_Buscar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function fgCargaCodHabDesc(iCombo As ComboBox)
Dim vlRsCombo As ADODB.Recordset
On Error GoTo Err_ComboHabDes

    vgSql = "select Cod_ConHabDes,Gls_ConHabDes "
    vgSql = vgSql & "from MA_TPAR_CONHABDES order by Cod_ConHabDes"
    Set vlRsCombo = vgConexionBD.Execute(vgSql)
    Do While Not vlRsCombo.EOF
        iCombo.AddItem ((Trim(vlRsCombo!Cod_ConHabDes) & " - " & Trim(vlRsCombo!gls_ConHabDes)))
        vlRsCombo.MoveNext
    Loop
    vlRsCombo.Close
    
    If iCombo.ListCount <> 0 Then
        iCombo.ListIndex = 0
    End If

Exit Function
Err_ComboHabDes:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fgComboTipoCalculo(iCombo As ComboBox)
iCombo.Clear
iCombo.AddItem ("D - Definitivo")
iCombo.AddItem ("P - Provisorio")
iCombo.ListIndex = 0
End Function

Function fgComboTipoPension(iCombo As ComboBox)
iCombo.Clear
iCombo.AddItem ("P - Primeros Pagos")
iCombo.AddItem ("R - Pagos Recurrentes")
iCombo.ListIndex = 0
End Function

Function fgLimiteEdad(vgIniVig As String, vgEdad As String, Mto_Edad As Double) As Boolean
On Error GoTo Err_Lim

    vgSql = ""
    fgLimiteEdad = False
    vgSql = "SELECT FEC_INIVIG,FEC_TERVIG,MTO_ELEMENTO FROM MA_TPAR_TABCODVIG"
    vgSql = vgSql & " WHERE COD_TABLA = 'LI' AND"
    vgSql = vgSql & " COD_ELEMENTO = '" & vgEdad & "'"
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
       While Not vgRs.EOF
          If vgIniVig >= (vgRs!fec_inivig) And _
             vgIniVig <= (vgRs!fec_tervig) Then
             fgLimiteEdad = True
             Mto_Edad = (vgRs!mto_elemento)
             Exit Function
          End If
          vgRs.MoveNext
       Wend
    End If
    vgRs.Close
    
Exit Function
Err_Lim:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " " & "Numero de Error : " & Err.Number & Err.HelpFile & Err.HelpContext, vbExclamation, "Mensaje de error ..."
        Screen.MousePointer = 0
    End If
End Function

''Function fgValidaPagoPension(iFecha As String, iNumPoliza As String, iNumOrden As Integer) As Boolean
''iFecha = Format(iFecha, "yyyymmdd")
''iFecha = Mid(iFecha, 1, 6)
''fgValidaPagoPension = False
'''Verifica Último Periodo
''    vgSql = ""
''    vgSql = "SELECT NUM_PERPAGO,COD_ESTADOPRI,COD_ESTADOREG"
''    vgSql = vgSql & " FROM PP_TMAE_PROPAGOPEN  "
''    vgSql = vgSql & "ORDER BY num_perpago DESC"
''    Set vgRs2 = vgConexionBD.Execute(vgSql)
''    If Not vgRs2.EOF Then
''        If vgRs2!Num_PerPago <= iFecha Then
''            'Verifica si es Primer Pago o Pago Régimen
''            vgSql = ""
''            vgSql = "SELECT NUM_POLIZA,NUM_ENDOSO,NUM_ORDEN FROM PP_TMAE_LIQPAGOPENDEF"
''            vgSql = vgSql & " Where "
''            vgSql = vgSql & " NUM_POLIZA = '" & iNumPoliza & "' AND "
''            vgSql = vgSql & " NUM_ORDEN = " & iNumOrden & ""
''           'vgSql = vgSql & " NUM_ENDOSO = " & inumend & " "
''            Set vgRs3 = vgConexionBD.Execute(vgSql)
''            If Not vgRs3.EOF Then
''                'Pago Régimen
''                If (vgRs2!cod_estadoreg) = "A" Or (vgRs2!cod_estadoreg) = "P" Then
''                    fgValidaPagoPension = True
''                Else
''                    If (vgRs2!cod_estadoreg) = "C" Then
''                        fgValidaPagoPension = False
''                    End If
''                End If
''            Else
''                'Primer Pago
''                If (vgRs2!cod_estadopri) = "A" Or (vgRs2!cod_estadopri) = "P" Then
''                    fgValidaPagoPension = True
''                Else
''                    If (vgRs2!cod_estadopri) = "C" Then
''                        fgValidaPagoPension = False
''                    End If
''                End If
''            End If
''        Else
''        fgValidaPagoPension = False
''        End If
''    End If
''
''End Function

Function fgBuscaFecServ() As String
    
    If vgTipoBase = "ORACLE" Then
       vgSql = ""
       'vgSql = "SELECT SYSDATE AS FEC_ACTUAL FROM MA_TCOD_GENERAL"
       vgSql = "SELECT TO_CHAR(SYSDATE,'DD/MM/YYYY HH24:MI:SS') AS FEC_ACTUAL FROM MA_TCOD_GENERAL"
       Set vgRs4 = vgConexionBD.Execute(vgSql)
       If Not vgRs4.EOF Then
          fgBuscaFecServ = Mid((vgRs4!FEC_ACTUAL), 1, 10)
          vgHoraActual = Mid((vgRs4!FEC_ACTUAL), 12, 8)
       End If
    Else
      If vgTipoBase = "SQL" Then
         vgSql = ""
         vgSql = "SELECT GETDATE()AS FEC_ACTUAL FROM MA_TCOD_GENERAL"
         Set vgRs4 = vgConexionBD.Execute(vgSql)
         If Not vgRs4.EOF Then
            fgBuscaFecServ = Mid((vgRs4!FEC_ACTUAL), 1, 10)
            vgHoraActual = Mid((vgRs4!FEC_ACTUAL), 12, 8)
         End If
      End If
    End If
    
End Function

''Function fgValidaVigenciaPoliza(iNumPoliza As String, iFecha As String)
''
''    fgValidaVigenciaPoliza = True
''    Const clEstado = "9"
''    Dim iNumEnd As Integer
''    iFecha = Format(iFecha, "yyyymmdd")
''
''    vgSql = "select num_endoso from PP_TMAE_POLIZA "
''    vgSql = vgSql & "where num_poliza='" & iNumPoliza & "' "
''    vgSql = vgSql & "order by num_endoso desc"
''    Set vgRs4 = vgConexionBD.Execute(vgSql)
''    If Not vgRs4.EOF Then
''        iNumEnd = vgRs4!num_endoso
''    Else
''        fgValidaVigenciaPoliza = False
''        Exit Function
''    End If
''    vgRs4.Close
''
''    vgSql = "select cod_estado,FEC_VIGENCIA from PP_TMAE_POLIZA "
''    vgSql = vgSql & "where num_poliza= '" & iNumPoliza & "' and "
''    vgSql = vgSql & "num_endoso= " & iNumEnd & " and "
''    vgSql = vgSql & "cod_estado<> '" & clEstado & "' and "
''    vgSql = vgSql & "fec_vigencia<= '" & iFecha & "'"
''    Set vgRs4 = vgConexionBD.Execute(vgSql)
''    If Not vgRs4.EOF Then
''        fgValidaVigenciaPoliza = True
''    Else
''        fgValidaVigenciaPoliza = False
''    End If
''    vgRs4.Close
''
''End Function
''

Function fgValidaPagoPension(iFecha As String, iNumPoliza As String, iNumOrden As Integer) As Boolean
Dim iFechaInicio As String
Dim iFechaTermino As String
Dim iAnno As Integer
Dim iMes As Integer
Dim iDia As Integer
Dim vlOpcionPago As String

    fgValidaPagoPension = False

    iFecha = Format(iFecha, "yyyymmdd")
   
    
'Calcula último día del mes
    iAnno = CInt(Mid(iFecha, 1, 4))
    iMes = CInt(Mid(iFecha, 5, 2))
    iDia = 1
    
    'hqr 17/03/2005
    'iFechaTermino = Format(DateSerial(iAnno, iMes + 1, iDia - 1), "yyyymmdd")
    If iFecha = vgTopeFecFin Then
        iFechaTermino = Format(DateSerial(2001, iMes + 1, iDia - 1), "yyyymmdd")
        iFechaTermino = Replace(iFechaTermino, "2001", iAnno)
    Else
        iFechaTermino = Format(DateSerial(iAnno, iMes + 1, iDia - 1), "yyyymmdd")
    End If
    'fin hqr 17/03/2005
    
'Calcula Primer día del mes
    iAnno = CInt(Mid(iFecha, 1, 4))
    iMes = CInt(Mid(iFecha, 5, 2))
    iDia = 1
    iFechaInicio = Format(DateSerial(iAnno, iMes, iDia), "yyyymmdd")
            
    iFecha = Mid(iFecha, 1, 6)
            
    'Estados del Pago de Pensión
    'PP: Primer Pago
    'PR: Pago en Regimen
    vlOpcionPago = ""
    
    'Estados que puede tener el Periodo
    'A : Abierto
    'P : Provisorio
    'C : Cerrado

    'Determinar si el Caso es Primer Pago o Pago en Regimen
    'vgSql = "SELECT NUM_POLIZA,NUM_ENDOSO,NUM_ORDEN "
    'vgSql = vgSql & " FROM PP_TMAE_LIQPAGOPENDEF WHERE "
    'vgSql = vgSql & " NUM_POLIZA = '" & iNumPoliza & "' "
    ''vgSql = vgSql & " AND NUM_ORDEN = " & iNumOrden & ""
    '''vgSql = vgSql & " NUM_ENDOSO = " & inumend & " "
    vgSql = "SELECT num_poliza,num_endoso "
    vgSql = vgSql & " FROM pp_tmae_poliza A WHERE "
    vgSql = vgSql & " num_poliza = '" & iNumPoliza & "' "
    'vgSql = vgSql & " AND NUM_ORDEN = " & iNumOrden & ""
    vgSql = vgSql & " AND NUM_ENDOSO = "
    vgSql = vgSql & " (SELECT MAX(NUM_ENDOSO) FROM PP_TMAE_POLIZA " 'Último Endoso
    vgSql = vgSql & " WHERE NUM_POLIZA = A.NUM_POLIZA)"
    vgSql = vgSql & " AND FEC_INIPAGOPEN BETWEEN '" & iFechaInicio & "'"
    vgSql = vgSql & " AND '" & iFechaTermino & "'"
    Set vgRs3 = vgConexionBD.Execute(vgSql)
    If vgRs3.EOF Then
        'Pago Régimen
        vlOpcionPago = "PR"
    Else
        'Primer Pago
        vlOpcionPago = "PP"
    End If
    vgRs3.Close

    vlOpcionPago = "PR"

    'Determinar la Existencia del Período Ingresado
    vgSql = "SELECT NUM_PERPAGO,COD_ESTADOREG " ',COD_ESTADOPRI
    vgSql = vgSql & "FROM PP_TMAE_PROPAGOPEN "
    vgSql = vgSql & "WHERE num_perpago = '" & iFecha & "'"
    'vgSql = vgSql & "ORDER BY num_perpago DESC"
    Set vgRs2 = vgConexionBD.Execute(vgSql)
    If Not vgRs2.EOF Then
        If vlOpcionPago = "PR" Then
            'Pago Régimen
            If (vgRs2!cod_estadoreg) <> "C" Then
                fgValidaPagoPension = True
            End If
        Else
            'Primer Pago
            If (vgRs2!cod_estadopri) <> "C" Then
                fgValidaPagoPension = True
            End If
        End If
    Else
        vgSw = False
        'Determinar si el periodo a registrar es posterior al que se desea ingresar
        vgSql = "SELECT NUM_PERPAGO,COD_ESTADOREG " ',COD_ESTADOPRI
        vgSql = vgSql & "FROM PP_TMAE_PROPAGOPEN "
        vgSql = vgSql & "WHERE "
        vgSql = vgSql & "num_perpago <= '" & iFecha & "' AND "
        If (vlOpcionPago = "PR") Then
            vgSql = vgSql & "cod_estadoreg <> 'C' "
        Else
            vgSql = vgSql & "cod_estadopri <> 'C' "
        End If
        vgSql = vgSql & "ORDER BY num_perpago ASC"
        Set vgRs3 = vgConexionBD.Execute(vgSql)
        If Not vgRs3.EOF Then
            fgValidaPagoPension = True
            vgSw = True
        End If
        vgRs3.Close
        
        If (vgSw = False) Then
            'Verificar que el periodo a registrar sea mayor que el
            'último que se encuentra Cerrado
            vgSql = "SELECT NUM_PERPAGO,COD_ESTADOREG " ',COD_ESTADOPRI
            vgSql = vgSql & "FROM PP_TMAE_PROPAGOPEN "
            vgSql = vgSql & "WHERE "
            vgSql = vgSql & "num_perpago <= '" & iFecha & "' AND "
            If (vlOpcionPago = "PR") Then
                vgSql = vgSql & "cod_estadoreg = 'C' "
            Else
                vgSql = vgSql & "cod_estadopri = 'C' "
            End If
            vgSql = vgSql & "ORDER BY num_perpago DESC"
            Set vgRs3 = vgConexionBD.Execute(vgSql)
            If Not vgRs3.EOF Then
                fgValidaPagoPension = True
            End If
            vgRs3.Close
            
        End If
    End If
    vgRs2.Close
    
End Function

Function fgValidaVigenciaPoliza(iNumPoliza As String, iFecha As String) As Boolean
Const clEstado = "9"
Dim iNumEnd As Integer

    fgValidaVigenciaPoliza = False
    iFecha = Format(iFecha, "yyyymmdd")
    
    vgSql = "select max(num_endoso) as num_endoso from PP_TMAE_POLIZA "
    vgSql = vgSql & "where num_poliza = '" & iNumPoliza & "' "
    vgSql = vgSql & "order by num_endoso desc"
    Set vgRs4 = vgConexionBD.Execute(vgSql)
    If Not vgRs4.EOF Then
        iNumEnd = vgRs4!num_endoso
    Else
        vgRs4.Close
        Exit Function
    End If
    vgRs4.Close
    
    vgSql = "select cod_estado,FEC_VIGENCIA "
    vgSql = vgSql & "from PP_TMAE_POLIZA where "
    vgSql = vgSql & "num_poliza = '" & iNumPoliza & "' and "
    vgSql = vgSql & "num_endoso = " & iNumEnd & " and "
    vgSql = vgSql & "cod_estado <> '" & clEstado & "' and "
    vgSql = vgSql & "fec_vigencia <= '" & iFecha & "'"
    Set vgRs4 = vgConexionBD.Execute(vgSql)
    If Not vgRs4.EOF Then
        fgValidaVigenciaPoliza = True
    End If
    vgRs4.Close
   
End Function

Function fgValidaFechaEfecto(iFecha As String, iNumPoliza As String, iNumOrden As Integer) As String
Dim vlOpcionPago As String
Dim vlPagoReg    As String
Dim vlAnno       As String, vlMes As String, vlDia As String
Dim iFechaInicio As String
Dim iFechaTermino As String
Dim iAnno As Integer
Dim iMes As Integer
Dim iDia As Integer
Dim vlRs3 As ADODB.Recordset
Dim vlRs2 As ADODB.Recordset

    fgValidaFechaEfecto = ""

    iFecha = Format(iFecha, "yyyymmdd")
    
'Calcula último día del mes
    iAnno = CInt(Mid(iFecha, 1, 4))
    iMes = CInt(Mid(iFecha, 5, 2))
    iDia = 1
    iFechaTermino = Format(DateSerial(iAnno, iMes + 1, iDia - 1), "yyyymmdd")
    
'Calcula Primer día del mes
    iAnno = CInt(Mid(iFecha, 1, 4))
    iMes = CInt(Mid(iFecha, 5, 2))
    iDia = 1
    iFechaInicio = Format(DateSerial(iAnno, iMes, iDia), "yyyymmdd")
            
    iFecha = Mid(iFecha, 1, 6)

    vlPagoReg = ""
    vlAnno = ""
    vlMes = ""
    vlDia = ""
    vgI = 0
    vgFechaEfecto = ""
    
    'Estados del Pago de Pensión
    'PP: Primer Pago
    'PR: Pago en Regimen
    vlOpcionPago = ""
    
    'Estados que puede tener el Periodo
    'A : Abierto
    'P : Provisorio
    'C : Cerrado

    'Determinar si el Caso es Primer Pago o Pago en Regimen
    'vgSql = "SELECT NUM_POLIZA,NUM_ENDOSO,NUM_ORDEN "
    'vgSql = vgSql & " FROM PP_TMAE_LIQPAGOPENDEF WHERE "
    'vgSql = vgSql & " NUM_POLIZA = '" & iNumPoliza & "' "
    ''vgSql = vgSql & " AND NUM_ORDEN = " & iNumOrden & ""
    '''vgSql = vgSql & " NUM_ENDOSO = " & inumend & " "
    vgSql = "SELECT num_poliza,num_endoso "
    vgSql = vgSql & " FROM pp_tmae_poliza A WHERE "
    vgSql = vgSql & " num_poliza = '" & iNumPoliza & "' "
    'vgSql = vgSql & " AND NUM_ORDEN = " & iNumOrden & ""
    vgSql = vgSql & " AND NUM_ENDOSO = "
    vgSql = vgSql & " (SELECT MAX(NUM_ENDOSO) FROM PP_TMAE_POLIZA " 'Último Endoso
    vgSql = vgSql & " WHERE NUM_POLIZA = A.NUM_POLIZA)"
    vgSql = vgSql & " AND FEC_INIPAGOPEN BETWEEN '" & iFechaInicio & "'"
    vgSql = vgSql & " AND '" & iFechaTermino & "'"
    Set vlRs3 = vgConexionBD.Execute(vgSql)
    If vlRs3.EOF Then
        'Pago Régimen
        vlOpcionPago = "PR"
    Else
        'Primer Pago
        vlOpcionPago = "PP"
    End If
    vlRs3.Close

    vlOpcionPago = "PR"
    
    'Determinar si el periodo a registrar es posterior al que se desea ingresar
    vgSql = ""
    vgSql = "SELECT NUM_PERPAGO,COD_ESTADOREG " ',COD_ESTADOPRI
    vgSql = vgSql & "FROM PP_TMAE_PROPAGOPEN "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "num_perpago >= '" & iFecha & "' AND "
    If (vlOpcionPago = "PR") Then
        vgSql = vgSql & "cod_estadoreg <> 'C' "
    Else
        vgSql = vgSql & "cod_estadopri <> 'C' "
    End If
    vgSql = vgSql & "ORDER BY num_perpago ASC"
    Set vlRs2 = vgConexionBD.Execute(vgSql)
    If Not vlRs2.EOF Then
        'If vlOpcionPago = "PR" Then
            vlPagoReg = vlRs2!Num_PerPago
        '    'Pago Régimen
        '    If (vlRs2!cod_estadoreg) <> "C" Then
        '        fgValidaPagoPension = True
        '    Else
        '        vgI = 1
        '    End If
        'Else
        '    'Primer Pago
        '    vlPagoReg = vlRs2!Num_PerPago
        '    If (vlRs2!cod_estadopri) <> "C" Then
        '        fgValidaPagoPension = True
        '    Else
        '        vgI = 1
        '    End If
        'End If
    Else
        'Determinar si el periodo a registrar es posterior al que se desea ingresar
        vgSql = ""
        vgSql = "SELECT NUM_PERPAGO,COD_ESTADOREG " ',COD_ESTADOPRI
        vgSql = vgSql & "FROM PP_TMAE_PROPAGOPEN "
        vgSql = vgSql & "WHERE "
        vgSql = vgSql & "num_perpago >= '" & iFecha & "' AND "
        If (vlOpcionPago = "PR") Then
            vgSql = vgSql & "cod_estadoreg = 'C' "
        Else
            vgSql = vgSql & "cod_estadopri = 'C' "
        End If
        vgSql = vgSql & "ORDER BY num_perpago DESC"
        Set vlRs3 = vgConexionBD.Execute(vgSql)
        If Not vlRs3.EOF Then
            vlPagoReg = vlRs3!Num_PerPago
            vgI = 1
        Else
            vlPagoReg = iFecha
        End If
        vlRs3.Close
    End If
    vlRs2.Close
    
    If (vlPagoReg <> "") Then
        vlAnno = Mid(vlPagoReg, 1, 4)
        vlMes = Mid(vlPagoReg, 5, 2) + vgI
        vlDia = "01"
        vgFechaEfecto = DateSerial(vlAnno, vlMes, vlDia)
    End If
    
    fgValidaFechaEfecto = vgFechaEfecto
End Function

Function fgDevuelveDigito(Rut As String, vlDgv As String)
Dim vlDigitos As String
Dim vlDiv As Long, vlMul As Double
Dim Resultado, y As Integer
    
    Rut = Format(Rut, "#0")
    vlDigitos = "32765432"
    Resultado = 0
    Select Case Len(Rut)
           Case 7: Rut = "0" + Rut
           Case 6: Rut = "00" + Rut
           Case 5: Rut = "000" + Rut
           Case 4: Rut = "0000" + Rut
           Case 3: Rut = "00000" + Rut
           Case 2: Rut = "000000" + Rut
           Case 1: Rut = "0000000" + Rut
    End Select
    For y = 1 To Len(Trim(Rut))
        Resultado = Resultado + Val(Mid(Rut, Len(Trim(Rut)) - y + 1, 1)) * Val(Mid(vlDigitos, Len(Trim(Rut)) - y + 1, 1))
    Next y
    vlDiv = Int(Resultado / 11)
    vlMul = vlDiv * 11
    vlDiv = Resultado - vlMul
    vlDgv = 11 - vlDiv
    If vlDgv = 11 Then
       vlDgv = 0
    End If
    If vlDgv = 10 Then
       vlDgv = "K"
    End If
End Function


Function fgBuscaTipoPago(iFecha As String, iNumPoliza As String) As String
On Error GoTo Err_fgBuscaTipoPago
Dim vlOpcionPago As String

    'Estados del Pago de Pensión
    'PP: Primer Pago
    'PR: Pago en Regimen
    vlOpcionPago = ""
    
    'Determinar si el Caso es Primer Pago o Pago en Regimen
    vgSql = "SELECT num_poliza,num_endoso,num_orden "
    vgSql = vgSql & " FROM PP_TMAE_LIQPAGOPENDEF WHERE "
    vgSql = vgSql & " num_poliza = '" & iNumPoliza & "' "
    Set vgRs3 = vgConexionBD.Execute(vgSql)
    If vgRs3.EOF Then
        'Pago Régimen
        fgBuscaTipoPago = "R"
    Else
        'Primer Pago
        fgBuscaTipoPago = "P"
    End If
    
    vgRs3.Close
        
Exit Function
Err_fgBuscaTipoPago:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " " & "Numero de Error : " & Err.Number & Err.HelpFile & Err.HelpContext, vbExclamation, "Mensaje de error ..."
        Screen.MousePointer = 0
    End If
End Function

Function fgVigenciaQuiebra(iFecha As String)
On Error GoTo Err_fgVigenciaQuiebra

    iFecha = Format(iFecha, "yyyymmdd")
    iFecha = Mid(iFecha, 1, 6)
    
    vgGlosaQuiebra = ""

    vgSql = ""
    vgSql = "SELECT num_perini,num_perter,gls_titulo "
    vgSql = vgSql & "FROM pp_tpar_quiebra "
    vgSql = vgSql & "ORDER BY num_perter DESC "
    Set vgRs3 = vgConexionBD.Execute(vgSql)
    If Not vgRs3.EOF Then
        If Not IsNull(vgRs3!num_perini) And Not IsNull(vgRs3!num_perter) Then
            If Trim(iFecha) >= Trim(vgRs3!num_perini) And _
               Trim(iFecha) <= Trim(vgRs3!num_perter) Then
               If Not IsNull(vgRs3!gls_titulo) Then
                   vgGlosaQuiebra = Trim(vgRs3!gls_titulo)
               End If
            Else
                vgGlosaQuiebra = ""
            End If
        End If
    End If

Exit Function
Err_fgVigenciaQuiebra:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function


Function fgComboTipoIdentificacion(vlCombo As ComboBox)
Dim vlRegCombo As ADODB.Recordset
Dim vlcont As Long
On Error GoTo Err_Combo

    vlCombo.Clear
    vgSql = ""
    vlcont = 0
    vgQuery = "SELECT cod_tipoiden as codigo, gls_tipoidencor as Nombre, "
    vgQuery = vgQuery & "num_lartipoiden as largo "
    vgQuery = vgQuery & "FROM MA_TPAR_TIPOIDEN "
    vgQuery = vgQuery & "WHERE cod_tipoiden <> 0 "
    vgQuery = vgQuery & "ORDER BY codigo "
    Set vlRegCombo = vgConexionBD.Execute(vgQuery)
    If Not (vlRegCombo.EOF) Then
        While Not (vlRegCombo.EOF)
            vlCombo.AddItem Space(2 - Len(vlRegCombo!Codigo)) & vlRegCombo!Codigo & " - " & (Trim(vlRegCombo!Nombre))
            vlcont = vlCombo.ListCount - 1
            vlCombo.ItemData(vlcont) = (vlRegCombo!largo)
            vlRegCombo.MoveNext
        Wend
    End If
    vlRegCombo.Close
        
    'Colocar el Tipo de Identificación "Sin Información" al final
    vgQuery = "SELECT cod_tipoiden as codigo, gls_tipoidencor as Nombre, "
    vgQuery = vgQuery & "num_lartipoiden as largo "
    vgQuery = vgQuery & "FROM MA_TPAR_TIPOIDEN "
    vgQuery = vgQuery & "WHERE cod_tipoiden = 0 "
    vgQuery = vgQuery & "ORDER BY codigo "
    Set vlRegCombo = vgConexionBD.Execute(vgQuery)
    If Not (vlRegCombo.EOF) Then
        While Not (vlRegCombo.EOF)
            vlCombo.AddItem Space(2 - Len(vlRegCombo!Codigo)) & vlRegCombo!Codigo & " - " & (Trim(vlRegCombo!Nombre))
            vlcont = vlCombo.ListCount - 1
            vlCombo.ItemData(vlcont) = (vlRegCombo!largo)
            vlRegCombo.MoveNext
        Wend
    End If
    vlRegCombo.Close
        
    If vlCombo.ListCount <> 0 Then
        vlCombo.ListIndex = 0
    End If

Exit Function
Err_Combo:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Public Function fgObtenerCodMonedaScomp(iEstructura() As TypeTablaMoneda, iNumTotal As Long, iTipoCodigo As String) As String
'Función: Permite obtener el Código de Scomp de una Moneda específica
'Parámetros de Entrada :
'- iEstructura => Estructura que contiene los Tipos de Moneda
'- iNumTotal   => Número Total de Filas de la Estructura
'- iTipoCodigo => Código de la Moneda a buscar
'Parámetros de Salida :
'- Devuelve el código de la Moneda definida para SCOMP
    
    fgObtenerCodMonedaScomp = ""
    
    If (iNumTotal <> 0) Then
        For vgX = 1 To iNumTotal
            If (iEstructura(vgX).Codigo = iTipoCodigo) Then
                fgObtenerCodMonedaScomp = iEstructura(vgX).Scomp
                Exit For
            End If
        Next vgX
    End If
    
End Function
Public Function fgCargarTablaMoneda(iCodTabla As String, oEstructura() As TypeTablaMoneda, oNumTotal As Long)
'Función : Llenar la Estructura de los Tipos de Moneda que se encuentran registrados en la BD
'Parámetros de Entrada:
'Parámetros de Salida:
'- Llenar la Estructura de Tipos de Monedas
'------------------------------------------------------
'Fecha de Creación     : 05/07/2007 - ABV
'Fecha de Modificación :
'------------------------------------------------------
Dim iRegistro As ADODB.Recordset
Dim iSql As String
On Error GoTo Err_Tabla

    oNumTotal = 0
    
    'Selecciona el Número Máximo de los Códigos de Parámetros
    vgQuery = "SELECT count(cod_elemento) as numero "
    vgQuery = vgQuery & "from ma_tpar_tabcod "
    vgQuery = vgQuery & "WHERE cod_tabla = '" & iCodTabla & "' "
    vgQuery = vgQuery & "AND (cod_sistema <> 'PP' OR cod_sistema is null) "
    Set iRegistro = vgConexionBD.Execute(vgQuery)
    If Not (iRegistro.EOF) Then
        If Not IsNull(iRegistro!numero) Then
            oNumTotal = iRegistro!numero
        End If
    End If
    iRegistro.Close
    
    'Llena la Estructura con los códigos de Parámetros
    If (oNumTotal <> 0) Then
        ReDim oEstructura(oNumTotal) As TypeTablaMoneda
        
        iSql = "SELECT cod_elemento as codigo,gls_elemento as descripcion,"
        iSql = iSql & "cod_scomp as cod_asociado "
        iSql = iSql & "FROM ma_tpar_tabcod "
        iSql = iSql & "WHERE cod_tabla = '" & iCodTabla & "' "
        iSql = iSql & "AND (cod_sistema <> 'PP' OR cod_sistema is null) "
        iSql = iSql & "ORDER BY cod_elemento "
        Set iRegistro = vgConexionBD.Execute(iSql)
        If Not (iRegistro.EOF) Then
            vgX = 1
            While Not (iRegistro.EOF)
                oEstructura(vgX).Codigo = iRegistro!Codigo
                oEstructura(vgX).Descripcion = IIf(IsNull(iRegistro!Descripcion), "", Trim(iRegistro!Descripcion))
                oEstructura(vgX).Scomp = IIf(IsNull(iRegistro!Cod_Asociado), "", iRegistro!Cod_Asociado)
                
                iRegistro.MoveNext
                vgX = vgX + 1
            Wend
        End If
        iRegistro.Close
    End If
    
Exit Function
Err_Tabla:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fgBuscarNombreTipoIden(iCodTipoIden As String, Optional iNombreLargo As Boolean) As String
Dim vlRegNombre As ADODB.Recordset
On Error GoTo Err_BuscarNombreCorto
    
    fgBuscarNombreTipoIden = ""
    
    If (iNombreLargo = True) Then
        vgSql = "SELECT gls_tipoiden as gls_tipoidencor "
    Else
        vgSql = "SELECT gls_tipoidencor "
    End If
    vgSql = vgSql & "FROM ma_tpar_tipoiden "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "cod_tipoiden = " & iCodTipoIden & " "
    Set vlRegNombre = vgConexionBD.Execute(vgSql)
    If Not vlRegNombre.EOF Then
         If Not IsNull(vlRegNombre!gls_tipoidencor) Then fgBuscarNombreTipoIden = vlRegNombre!gls_tipoidencor
    End If
    vlRegNombre.Close
  
Exit Function
Err_BuscarNombreCorto:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Public Function fgObtenerCodigo_TextoCompuesto(iTexto As String) As String
'Función: Permite obtener el Código de un Texto que tiene el Código y la
'Descripción separados por un Guión
'Parámetros de Entrada :
'- iCodigoDescripcion => Estructura que contiene los Tipos de Moneda
'Parámetros de Salida :
'- Devuelve el código del Texto
    
    If (InStr(1, iTexto, "-") <> 0) Then
        fgObtenerCodigo_TextoCompuesto = Trim(Mid(iTexto, 1, InStr(1, iTexto, "-") - 1))
    Else
        fgObtenerCodigo_TextoCompuesto = UCase(Trim(iTexto))
    End If

End Function

Function fgBuscarNombreComunaProvinciaRegion(iCodDir As String)
Dim vlRegistroDir As ADODB.Recordset
On Error GoTo Err_Buscar

     vgSql = "SELECT r.Cod_Region,r.Gls_Region,p.Cod_Provincia,p.Gls_Provincia,c.Cod_Comuna,c.Gls_Comuna"
     vgSql = vgSql & " FROM MA_TPAR_COMUNA c, MA_TPAR_PROVINCIA p, MA_TPAR_REGION r"
     vgSql = vgSql & " Where c.Cod_Direccion = '" & iCodDir & "' and  "
     vgSql = vgSql & " c.cod_region = p.cod_region and"
     vgSql = vgSql & " c.cod_provincia = p.cod_provincia and"
     vgSql = vgSql & " p.cod_region = r.cod_region"
     Set vlRegistroDir = vgConexionBD.Execute(vgSql)
     If Not vlRegistroDir.EOF Then
        vgNombreRegion = (vlRegistroDir!gls_region)
        vgNombreProvincia = (vlRegistroDir!gls_provincia)
        vgNombreComuna = (vlRegistroDir!gls_comuna)
        vgCodigoRegion = (vlRegistroDir!cod_region)
        vgCodigoProvincia = (vlRegistroDir!cod_provincia)
        vgCodigoComuna = (vlRegistroDir!cod_comuna)
     End If
     vlRegistroDir.Close

Exit Function
Err_Buscar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Public Function fgObtenerCod_Identificacion(iNombreCorto As String, oCodigo As Long) As Boolean
Dim vlRegBuscar As ADODB.Recordset
    
    fgObtenerCod_Identificacion = False
    oCodigo = 0
    
    vgQuery = "SELECT cod_tipoiden as codigo "
    vgQuery = vgQuery & "FROM ma_tpar_tipoiden "
    vgQuery = vgQuery & "WHERE "
    vgQuery = vgQuery & "gls_tipoidencor = '" & iNombreCorto & "'"
    Set vlRegBuscar = vgConexionBD.Execute(vgQuery)
    If Not (vlRegBuscar.EOF) Then
        If Not IsNull(vlRegBuscar!Codigo) Then
            oCodigo = vlRegBuscar!Codigo
            fgObtenerCod_Identificacion = True
        End If
    End If
    vlRegBuscar.Close

End Function

Function fgFormarNombreCompleto(iNombre As String, iNombreSeg As String, iPaterno As String, iMaterno As String) As String

fgFormarNombreCompleto = ""

If (iNombre = "") Then iNombre = "" Else iNombre = iNombre & " "
If (iNombreSeg = "") Then iNombreSeg = "" Else iNombreSeg = iNombreSeg & " "
If (iPaterno = "") Then iPaterno = "" Else iPaterno = iPaterno & " "
If (iMaterno = "") Then iMaterno = "" Else iMaterno = iMaterno & " "

fgFormarNombreCompleto = Trim(iNombre & iNombreSeg & iPaterno & iMaterno)
End Function

Public Function fgObtenerPolizaCod_AFP(iNumPoliza As String, inumendoso As String) As String
Dim vlRegBuscar As ADODB.Recordset
    
    fgObtenerPolizaCod_AFP = ""
    
    vgQuery = "SELECT cod_afp as codigo "
    vgQuery = vgQuery & "FROM pp_tmae_poliza "
    vgQuery = vgQuery & "WHERE "
    vgQuery = vgQuery & "num_poliza = '" & iNumPoliza & "' AND "
    vgQuery = vgQuery & "num_endoso = " & inumendoso & " "
    Set vlRegBuscar = vgConexionBD.Execute(vgQuery)
    If Not (vlRegBuscar.EOF) Then
        If Not IsNull(vlRegBuscar!Codigo) Then
            fgObtenerPolizaCod_AFP = vlRegBuscar!Codigo
        End If
    End If
    vlRegBuscar.Close

End Function

Function fgComboTipoPago(iCombo As ComboBox)
iCombo.Clear
iCombo.AddItem ("B - Boleta")
iCombo.AddItem ("F - Factura")
iCombo.ListIndex = 0
End Function

Function fgComboPDT(iCombo As ComboBox)
iCombo.Clear
iCombo.AddItem ("T00 - Datos de Pensionistas - Detallados")
iCombo.AddItem ("T02 - Datos de Pensionistas - Especificos")
iCombo.AddItem ("P00 - Datos de Periodos")
iCombo.AddItem ("PEN - Datos de Remuneración")
iCombo.AddItem ("DER - Datos de Derechohabientes")
iCombo.AddItem ("PER - Otras Condiciones")
iCombo.ListIndex = 0
End Function

Function f_dia_ultimo(ByVal ianio As Integer, ByVal iMes As Integer)

   Select Case iMes
      Case 1, 3, 5, 7, 8, 10, 12
         f_dia_ultimo = 31
      Case 2
            If ianio Mod 4 = 0 Then
                f_dia_ultimo = 29
            Else
                f_dia_ultimo = 28
            End If
      Case 4, 6, 9, 11
         f_dia_ultimo = 30
   End Select
   
End Function

Function fTipoCambio(ByVal sCodMon As String, ByVal dFecha As String) As Double

Dim rs_Temp As ADODB.Recordset
Dim Sql As String
    'Para tipo de cambio
    Sql = "SELECT mto_moneda FROM MA_TVAL_MONEDA WHERE "
    Sql = Sql & "cod_moneda = '" & sCodMon & "' AND "
    Sql = Sql & "fec_moneda = '" & dFecha & "'"
    Set rs_Temp = New ADODB.Recordset
    Set rs_Temp = vgConexionBD.Execute(Sql)
    If rs_Temp.EOF Then
        fTipoCambio = 0
    Else
        fTipoCambio = rs_Temp!Mto_Moneda
    End If

End Function

Function f_amd_dma(ByVal lfecha As Long)
    
    Dim sFecha As String
    
    f_amd_dma = ""
    If lfecha <> 0 Then
        sFecha = Trim$(lfecha)
        f_amd_dma = Right(sFecha, 2) & "/" & Mid(sFecha, 5, 2) & "/" & Left(sFecha, 4)
    End If
    
End Function

Sub p_sombrea_texto(ByVal objcontrol As Control)

  objcontrol.SelStart = 0
  objcontrol.SelLength = Len(objcontrol.Text)
  
End Sub

Function f_convierte_mayuscula(ByVal icaracter As Integer)

   If (Chr$(icaracter) >= "a" And Chr$(icaracter) <= "z") Or Chr$(icaracter) = "ñ" Then
      f_convierte_mayuscula = Asc(UCase$(Chr(icaracter)))
   Else
      f_convierte_mayuscula = icaracter
   End If
   
End Function
Function f_valida_numeros(ByVal icaracter As Integer)
   
   If (Chr$(icaracter) >= "0" And Chr$(icaracter) <= "9") Or icaracter = 8 Or Chr$(icaracter) = "-" Then
      f_valida_numeros = icaracter
   Else
      f_valida_numeros = 0
   End If

End Function
Function FechaServidor() As Date

Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset

RS.Open "select to_char(sysdate,'dd/mm/yyyy')as fecha from dual", vgConexionBD, adOpenStatic, adLockReadOnly
FechaServidor = RS!fecha

End Function

Function flObtieneEdadNormativa(StrNumCot As String) As Double

    Dim FecDevSol As String
    Dim dateFecdevsol As Date
    Dim dateFecCortee As Date
    Dim EdadRes As Double
    Dim vlRegistro As ADODB.Recordset

    flObtieneEdadNormativa = 216

    vgSql = ""
    vgSql = "select fec_devsol from PP_TMAE_POLIZA where "
    vgSql = vgSql & " NUM_POLIZA ='" & StrNumCot & "'"
    Set vlRegistro = vgConexionBD.Execute(vgSql)
    
    If Not (vlRegistro.EOF) <> 0 Then
        FecDevSol = IIf(IsNull(vlRegistro!FEC_DEVSOL), "", vlRegistro!FEC_DEVSOL)
    End If
    
    If Len(FecDevSol) <> 0 Then
        dateFecdevsol = CDate(Mid(FecDevSol, 7, 2) & "/" & Mid(FecDevSol, 5, 2) & "/" & Mid(FecDevSol, 1, 4))
        dateFecCortee = CDate("01/08/2013")
        If dateFecdevsol >= dateFecCortee Then
            flObtieneEdadNormativa = fgCarga_ParamV("LI", "L24") * 12
        Else
            flObtieneEdadNormativa = fgCarga_ParamV("LI", "L18") * 12
        End If
    End If
End Function

Function flCompletaRequisitos(StrNumCot As String, StrNorden As Integer) As Boolean

    Dim valor As Integer
    Dim vlRegistro As ADODB.Recordset
    flCompletaRequisitos = False

    vgSql = ""
    vgSql = "SELECT C.NUM_POLIZA, C.NUM_ORDEN,"
    vgSql = vgSql & " sum(case when nvl(ind_dni, 0)='S' then 1 else 0 end +"
    vgSql = vgSql & " case when nvl(ind_dju, 0)='S' then 1 else 0 end +"
    vgSql = vgSql & " case when nvl(ind_pes, 0)='S' then 1 else 0 end +"
    vgSql = vgSql & " case when nvl(ind_bno, 0)='S' then 1 else 0 end ) as val"
    vgSql = vgSql & " FROM PP_TMAE_CERTIFICADO C"
    vgSql = vgSql & " JOIN pp_tmae_ben B ON C.num_poliza=B.num_poliza and C.num_orden=B.num_orden and C.num_endoso=B.num_endoso"
    vgSql = vgSql & " WHERE B.COD_PAR IN ('30')"
    vgSql = vgSql & " AND cod_derpen='99'"
    vgSql = vgSql & " AND cod_estpension='99'"
    vgSql = vgSql & " AND C.NUM_ENDOSO=(select max(num_endoso) from pp_tmae_poliza where num_poliza=C.num_poliza)"
    vgSql = vgSql & " AND B.COD_SITINV NOT IN ('T', 'P')"
    vgSql = vgSql & " AND C.COD_TIPO='EST'"
    vgSql = vgSql & " AND C.NUM_POLIZA='" & StrNumCot & "' AND C.NUM_ORDEN=" & StrNorden & ""
    vgSql = vgSql & " GROUP BY C.NUM_POLIZA, C.NUM_ORDEN"
    Set vlRegistro = vgConexionBD.Execute(vgSql)
    
    If Not (vlRegistro.EOF) <> 0 Then
        valor = IIf(IsNull(vlRegistro!Val), 0, vlRegistro!Val)
    End If

    If valor = 4 Then
        flCompletaRequisitos = True
    End If
    
End Function

Function flSiDebeCambiarParentesco(StrNumCot As String, strFecPagoAct As String) As Integer

    Dim FecPriPago As String
    Dim dateFecpripago As Date
    Dim dateFecCortee As Date
    Dim EdadRes As Double
    Dim vlRegistro As ADODB.Recordset

    'flObtieneEdadNormativa = 216

    vgSql = ""
    vgSql = "select fec_pripago from PP_TMAE_POLIZA where "
    vgSql = vgSql & " NUM_POLIZA ='" & StrNumCot & "' and num_endoso=1"
    Set vlRegistro = vgConexionBD.Execute(vgSql)
    
    If Not (vlRegistro.EOF) <> 0 Then
        FecPriPago = IIf(IsNull(vlRegistro!Fec_PriPago), "", vlRegistro!Fec_PriPago)
    End If
    
    If Len(FecPriPago) <> 0 Then
        dateFecpripago = CDate(Mid(FecPriPago, 7, 2) & "/" & Mid(FecPriPago, 5, 2) & "/" & Mid(FecPriPago, 1, 4))
        dateFecCortee = CDate(Mid(strFecPagoAct, 7, 2) & "/" & Mid(strFecPagoAct, 5, 2) & "/" & Mid(strFecPagoAct, 1, 4))
        If dateFecpripago > dateFecCortee Then
            flSiDebeCambiarParentesco = 1
        Else
            flSiDebeCambiarParentesco = 0
        End If
    End If
End Function

Function FgGuardaLog(sQuery As String, sUser As String, nError As String) As Boolean
    Dim vlSql As String
    Dim oValor As Long
    Dim vlRegistro As ADODB.Recordset
    
    vlSql = "SELECT count(*) + 1 as Cor"
    vlSql = vlSql & " FROM PP_TMAE_LOGACTUAL a"
    Set vlRegistro = vgConexionBD.Execute(vlSql)
    If Not vlRegistro.EOF Then
        oValor = vlRegistro!Cor
    End If
    'vlRegistro.Close
    sQuery = Replace(sQuery, "'", "|")
    vlSql = ""
    vlSql = "INSERT INTO PP_TMAE_LOGACTUAL (LOG_NNUMREG, LOG_CCODERR, LOG_SQUERY ,LOG_DFECREG ,LOG_CUSUCRE) VALUES ("
    vlSql = vlSql & oValor & ",'" & Trim(nError) & "','" & Trim(sQuery) & "','" & Format(Date, "yyyymmdd") & "','" & sUser & "')"
    vgConexionBD.Execute (vlSql)

End Function

'------------------------------------------------------------
'Permite Cargar el Combo de Sucursales del Sistema
'------------------------------------------------------------
Function fgComboTipoRegularizacion(iCombo As ComboBox, optTodos As Boolean)
Dim vlRsCombo As ADODB.Recordset
On Error GoTo Err_ComboSucursal
    
    iCombo.Clear
    vgSql = "select Cod_Elemento, Gls_Elemento from MA_TPAR_TABCOD where Cod_Tabla='TRG'"
    Set vlRsCombo = vgConexionBD.Execute(vgSql)
    
    If optTodos Then
        iCombo.AddItem (("0 - [TODOS]"))
    End If
    
    Do While Not vlRsCombo.EOF
        iCombo.AddItem ((Trim(vlRsCombo!COD_ELEMENTO) & " - " & Trim(vlRsCombo!GLS_ELEMENTO)))
        vlRsCombo.MoveNext
    Loop
    vlRsCombo.Close
    If iCombo.ListCount <> 0 Then
        iCombo.ListIndex = 0
    End If
    
Exit Function
Err_ComboSucursal:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Public Sub Foco(ctactu As Control)
    ctactu.SelStart = 0
    ctactu.SelLength = Len(ctactu.Text)
End Sub


Attribute VB_Name = "Mod_Endoso"
'I--- ABV 21/06/2006 ---
Const cgPensionInvVejez  As String = "02,04,05,06,07,14"
Const cgPensionSobOrigen As String = "08,13"
Const cgPensionSobTransf As String = "01,03,09,10,11,12,15"
Const cgParConyugeMadre  As String = "10,11,20,21"

'Código de la Causal con la cual se genera un Endoso adelantado para un periodo Diferido - 26/09/2007
Global Const cgCodCausalEndPagoDif As String * 2 = "13"

'Copiar en Módulo General para el Tema de Endoso
'Definición de la Estructura que guardará a las Tablas de Mortalidad
Public Type TypeTabla
    Correlativo As Double
    TipoTabla   As String
    Nombre      As String
    Sexo        As String
    FechaIni    As Long
    FechaFin    As Long
    TipoPeriodo As String
    TipoGenerar As String
    IniTab      As Long
    Fintab      As Long
    tasa        As Double
    Oficial     As String
    Estado      As String
    TipoMovimiento As String
    AñoBase     As Integer
'I--- ABV 17/11/2005 ---
    Descripcion As String
'F--- ABV 17/11/2005 ---
End Type

Global egTablaMortal() As TypeTabla
Global vgNumeroTotalTablas As Long

Global Lx() As Double
Global Ly() As Double
Global vgFechaAnterior     As String
Global vgTina              As Double
Global Flupen()            As Double
Global Flucm()             As Double
Global FluBenef(8, 1332)   As Double
Global vgCargarTasas       As Boolean
Global vgCargarTasasRea    As Boolean
Global vgMatrizTasas()     As Double
Global vgMatrizTasasRea()    As Double

Global vgs_Sexo  As String
Global vgs_Tipo  As String
Global vgs_Nro   As String
Global vgs_Error As Long
Global Tb2       As ADODB.Recordset
Global Fintab        As Long
Global vgFinTabVit_F As Long
Global vgFinTabTot_F As Long
Global vgFinTabPar_F As Long
Global vgFinTabBen_F As Long
Global vgFinTabVit_M As Long
Global vgFinTabTot_M As Long
Global vgFinTabPar_M As Long
Global vgFinTabBen_M As Long

Global L24      As Double
Global L21      As Double
Global L18      As Double

Global vgMortalVit_M As Long
Global vgMortalTot_M As Long
Global vgMortalPar_M As Long
Global vgMortalBen_M As Long
Global vgMortalVit_F As Long
Global vgMortalTot_F As Long
Global vgMortalPar_F As Long
Global vgMortalBen_F As Long

Global vgPalabra_MortalVit_M As String
Global vgPalabra_MortalTot_M As String
Global vgPalabra_MortalPar_M As String
Global vgPalabra_MortalBen_M As String
Global vgPalabra_MortalVit_F As String
Global vgPalabra_MortalTot_F As String
Global vgPalabra_MortalPar_F As String
Global vgPalabra_MortalBen_F As String

Global vgGtoFun            As Double
Global vgTipoReserva       As String
Global vgCodTipoReserva        As String

Global vgBuscarMortalVit_F As String
Global vgBuscarMortalTot_F As String
Global vgBuscarMortalPar_F As String
Global vgBuscarMortalBen_F As String
Global vgBuscarMortalVit_M As String
Global vgBuscarMortalTot_M As String
Global vgBuscarMortalPar_M As String
Global vgBuscarMortalBen_M As String

Global vgFechaIniMortalVit_F As String
Global vgFechaFinMortalVit_F As String
Global vgFechaIniMortalTot_F As String
Global vgFechaFinMortalTot_F As String
Global vgFechaIniMortalPar_F As String
Global vgFechaFinMortalPar_F As String
Global vgFechaIniMortalBen_F As String
Global vgFechaFinMortalBen_F As String
Global vgFechaIniMortalVit_M As String
Global vgFechaFinMortalVit_M As String
Global vgFechaIniMortalTot_M As String
Global vgFechaFinMortalTot_M As String
Global vgFechaIniMortalPar_M As String
Global vgFechaFinMortalPar_M As String
Global vgFechaIniMortalBen_M As String
Global vgFechaFinMortalBen_M As String

'I--- ABV 15/03/2005 ---
Global vgIndicadorTipoMovimiento_F As String
Global vgIndicadorTipoMovimiento_M As String
Global vgDinamicaAñoBase_F         As Integer
Global vgDinamicaAñoBase_M         As Integer

Global Const vgTipoPeriodo = "M"
Global Const vgTipoTablaRentista = "RV"
Global Const vgTipoTablaTotal = "MIT"
Global Const vgTipoTablaParcial = "MIP"
Global Const vgTipoTablaBeneficiario = "B"
Global Const vgEdadJubilacionHombre = 65
Global Const vgEdadJubilacionMujer = 60
Global Const vgTasaCapitalBono = 4
Global Const vgMonedaOficial = "NS"

Global vgFechaTexto   As String

'F--- ABV 15/03/2005 ---

'I--- ABV 17/11/2005
Global vgUtilizarNormativa As String
Global vgReservaUtilizada  As String
'F--- ABV 17/11/2005
Global vgCausaEndoso As String

'************* Variables Endosos Automaticos ***********************

Dim vlNumPoliza As String
Dim vlNumEndoso As Integer
Dim vlNumOrden As Integer

'Variables Tabla Beneficiario
Dim vlNumPolizaB As String
Dim vlNumEndosoB As Integer
Dim vlNumOrdenB As Integer
Dim vlFecIngreso As String
Dim vlCodTipoIdenBen As String
Dim vlNumIdenBen As String
Dim vlGlsNomBen As String
Dim vlGlsNomSegBen As String
Dim vlGlsPatBen As String
Dim vlGlsMatBen As String
Dim vlGlsDirBen As String
Dim vlCodDireccion As Integer
Dim vlGlsFonoBen As String
Dim vlGlsCorreoBen As String
Dim vlCodGruFam As String
Dim vlCodPar As String
Dim vlCodSexo As String
Dim vlCodSitInv As String
Dim vlCodDerCre As String
Dim vlCodDerpen As String
Dim vlCodCauInv As String
Dim vlFecNacBen As String
Dim vlFecNacHM As String
Dim vlFecInvBen As String
Dim vlCodMotReqPen As String
Dim vlMtoPension As Double
Dim vlMtoPensionGar As Double
Dim vlPrcPension As Double
Dim vlCodInsSalud As String
Dim vlCodModSalud As String
Dim vlMtoPlanSalud As Double
Dim vlCodEstPension As String
Dim vlCodViaPago As String
Dim vlCodBanco As String
Dim vlCodTipCuenta As String
Dim vlNumCuenta As String
Dim vlCodSucursal As String
Dim vlFecFallBen As String
Dim vlFecMatrimonio As String
Dim vlCodCauSusBen As String
Dim vlFecSusBen As String
Dim vlFecIniPagoPen As String
Dim vlFecTerPagoPenGar As String
Dim vlCodUsuarioCreaB As String
Dim vlFecCreaB As String
Dim vlHorCreaB As String
Dim vlCodUsuarioModiB As String
Dim vlFecModiB As String
Dim vlHorModiB As String
Dim vlPrcPensionLeg As Double
Dim vlPrcPensionGar As Double
Dim vlMtoPensionAct As Double


'Variables Tabla Poliza
Dim vlNumPolizaP As String
Dim vlNumEndosoP As Integer
Dim vlCodAFP As String
Dim vlCodTipPension As String
Dim vlCodEstadoP As String
Dim vlCodTipRen As String
Dim vlCodModalidad As String
Dim vlNumCargas As Integer
Dim vlFecVigencia As String
Dim vlFecTerVigencia As String
Dim vlMtoPrima As Double
Dim vlMtoPensionPol As Double
Dim vlNumMesDif As Double
Dim vlNumMesGar As Long
Dim vlPrcTasaCe As Double
Dim vlPrcTasaVta As Double
Dim vlPrcTasaCtoRea As Double
Dim vlPrcTasaIntPerGar As Double
Dim vlFecIniPagoPenPol As String
Dim vlCodUsuarioCreaP As String
Dim vlFecCreaP As String
Dim vlHorCreaP As String
Dim vlCodUsuarioModiP As String
Dim vlFecModiP As String
Dim vlHorModiP As String
Dim vlCodTipOrigen As String
Dim vlNumIndQuiebra As Integer
'hqr 27/08/2007 campos agregados
Dim vlMtoPensionGarP As Double
Dim vlCodCuspp As String
Dim vlIndCob As String
Dim vlCodMoneda As String
Dim vlMtoValMoneda As Double
Dim vlCodCobercon As String
Dim vlMtoFacPenElla As Double
Dim vlPrcFacPenElla As Double
Dim vlCodDercreP As String
Dim vlCodDerGra As String
Dim vlPrcTasaTir As Double
Dim vlFecEmision As String
Dim vlFecDev As String
Dim vlFecIniPenCia As String
Dim vlFecPriPago As String
Dim vlFecFinPerdif As String
Dim vlFecFinPerGar As String
Dim vlCodTipReajuste As String 'hqr 03/12/2010
Dim vlMtoValReajusteTri As Double 'hqr 03/12/2010
Dim vlMtoValReajusteMen As Double 'hqr 19/02/2011
Dim vlFecDevSol As String 'RRR
'Variables Tabla Endoso
Dim vlNumPolizaE As String
Dim vlNumEndosoE As String
Dim vlFecSolEndoso As String
Dim vlFecEndoso As String
Dim vlCodCauEndoso As String
Dim vlCodTipEndoso As String
Dim vlCodMonedaEndoso As String
Dim vlMtoDiferencia As Double
Dim vlMtoPensionOri As Double
Dim vlMtoPensionCal As Double
Dim vlFecEfecto As String
Dim vlPrcFactor As Double
Dim vlGlsObservacion As String
Dim vlCodUsuarioCreaE As String
Dim vlFecCreaE As String
Dim vlHorCreaE As String
Dim vlCodUsuarioModiE As String
Dim vlFecModiE As String
Dim vlHorModiE As String
Dim vlFecFinEfecto As String
Dim vlCodEstadoE As String
Dim vlCodTipReajusteE As String 'hqr 03/12/2010
Dim vlMtoValReajusteETri As Double 'hqr 03/12/2010
Dim vlMtoValReajusteEMen As Double 'hqr 19/02/2011
'Dim vlFecFinPerGar As String
'Dim vlGlsUsuarioCrea As String
'Dim vlFecCrea As String
'Dim vlHorCrea As String
'Dim vlGlsUsuarioModi As String
'Dim vlFecModi As String
'Dim vlHorModi As String

Dim vlCamposPoliza As String
Dim vlCamposBen As String

Dim vlNumPolAnterior As String

Dim vlNumEndosoNuevo As Integer
Dim vlNumUltEndoso As Integer

Dim vlFechaActual As String
Dim vlFecha18 As String
Dim vlFecha24 As String

Dim vlLargoArchivo As Integer
Dim vlLargoRegistro As Integer
Dim vlAumento As Integer

Global vgRegCerEst As ADODB.Recordset
Global vgRegPolizas As ADODB.Recordset
Global vgRegBen As ADODB.Recordset

Const clCodEstPension10 As String * 2 = "10"
Const clCodEstPension20 As String * 2 = "20"
Const clCodEstPension99 As String * 2 = "99"
Const clCodSitInvN As String * 1 = "N"
Const clCodCauEndoso14 As String * 2 = "14"
Const clCodTipEndosoS As String * 1 = "S"
Const clFechaTope As String = "99991231"
Const clCodEstadoP As String * 1 = "P"
Const clCodEstadoE As String * 1 = "E"
Const clCodParHijo30 As String * 2 = "30"
Const clCodParHijo35 As String * 2 = "35"
Const clPrcFactor1 As Double = 1


Global vgGeneraEndosos As Boolean
Dim vlFactorAjusteEndoso As Double
Dim vlFactorAjusteTasaFijaEndoso As Double 'hqr 03/12/2010



 ' La Propiedad Nombre
Private num_pol As String
Global vgNum_pol As String

Public Property Get NumPoliza() As String
    NumPoliza = num_pol
End Property

Public Property Let NumPoliza(ByVal Value As String)
    num_pol = Value
End Property


Function amax0(param1, param2) As Long
    If param1 > param2 Then
        amax0 = param1
    Else
        amax0 = param2
    End If
End Function

Function amin0(param1, param2) As Long
    If param1 > param2 Then
        amin0 = param2
    Else
        amin0 = param1
    End If
End Function

Function amax1(param1, param2) As Double
    If param1 > param2 Then
        amax1 = param1
    Else
        amax1 = param2
    End If
End Function

Function amin1(param1, param2) As Double
    If param1 < param2 Then
        amin1 = param1
    Else
        amin1 = param2
    End If
End Function

Function fgCarga_ParamV(iTabla As String, iElemento As String) As Double
    
    'Call AbrirBaseDeDatos_Aux(vgRutaBasedeDatos)
'''    If Not AbrirBaseDeDatos(vgConexionParam) Then
'''        MsgBox "Falló la conexión con la Base de Datos.", vbCritical, "Error de Conexión"
'''        Exit Function
'''    End If
    
    fgCarga_ParamV = -1000
    
    'vgSql = "SELECT * FROM MA_TPAR_TABCOD WHERE "

    vgSql = "SELECT * FROM MA_TPAR_TABCODVIG WHERE "
    vgSql = vgSql & "COD_TABLA = '" & iTabla & "' and "
    vgSql = vgSql & "COD_ELEMENTO = '" & iElemento & "'  "
    Set vlRegistro = vgConexionBD.Execute(vgSql)
    If Not (vlRegistro.EOF) Then
        fgCarga_ParamV = Trim(vlRegistro!mto_elemento)
    End If
    vlRegistro.Close

'''    Call CerrarBaseDeDatos(vgConexionParam)

End Function

Function fgCarga_Param(iTabla As String, iElemento As String, iFecha As String) As Boolean
Dim vlRegistro As ADODB.Recordset

    'Call AbrirBaseDeDatos_Aux(vgRutaBasedeDatos)
'''    If Not AbrirBaseDeDatos(vgConexionParam) Then
'''        MsgBox "Falló la conexión con la Base de Datos.", vbCritical, "Error de Conexión"
'''        Exit Function
'''    End If

    fgCarga_Param = False
    vgValorParametro = 0

    vgSql = "SELECT mto_elemento FROM MA_TPAR_TABCODVIG WHERE "
    vgSql = vgSql & "COD_TABLA = '" & iTabla & "' and "
    vgSql = vgSql & "COD_ELEMENTO = '" & iElemento & "' "
    vgSql = vgSql & "AND (FEC_INIVIG <= '" & iFecha & "' "
    vgSql = vgSql & "AND FEC_TERVIG >= '" & iFecha & "') "
    Set vlRegistro = vgConexionBD.Execute(vgSql)
    'If Not (vlRegistro.EOF) Then
    If Not (vlRegistro.EOF) Then
        If Not IsNull(vlRegistro!mto_elemento) Then
            vgValorParametro = Trim(vlRegistro!mto_elemento)

            fgCarga_Param = True
        End If
    End If
    vlRegistro.Close

'''    Call CerrarBaseDeDatos(vgConexionParam)
End Function

Function fgCarga_CuoMor(iFecha As String, iMoneda As String) As Double
Dim vlRegistro As ADODB.Recordset

    vgSql = "SELECT mto_CUOMOR FROM MA_TVAL_CUOMOR WHERE "
    vgSql = vgSql & "COD_MONEDA = '" & iMoneda & "' and "
    vgSql = vgSql & "FEC_INICUOMOR <= '" & iFecha & "' and "
    vgSql = vgSql & "FEC_TERCUOMOR >= '" & iFecha & "' "
    Set vlRegistro = vgConexionBD.Execute(vgSql)
    If Not (vlRegistro.EOF) Then
        fgCarga_CuoMor = Trim(vlRegistro!mto_CUOMOR)
    Else
        fgCarga_CuoMor = -1000
    End If
    vlRegistro.Close
    
End Function

Function fgCargarTablaMortalidad(ioPeriodo)
Dim vgCmb As ADODB.Recordset
Dim iRegistro As ADODB.Recordset
Dim iSql As String
On Error GoTo Err_Tabla

    'Call AbrirBaseDeDatos_Aux(vgRutaBasedeDatos)
'''    If Not AbrirBaseDeDatos(vgConectarBD) Then
'''        MsgBox "Falló la conexión con la Base de Datos.", vbCritical, "Error de Conexión"
'''        Exit Function
'''    End If
    
    'If (iFecha <> "") And (iPeriodo <> "") And (iSexo <> "") And (iTipoTabla <> "") Then
    If (vgTipoPeriodo <> "") Then
        
        vgNumeroTotalTablas = 0
        
        vgQuery = "SELECT count(num_correlativo) as numero "
        vgQuery = vgQuery & "from MA_TVAL_MORTAL WHERE "
        vgQuery = vgQuery & "cod_tipoper = '" & vgTipoPeriodo & "' "
        vgQuery = vgQuery & "and "
        vgQuery = vgQuery & "cod_tiptabmor in ("
        vgQuery = vgQuery & "'" & vgTipoTablaParcial & "',"
        vgQuery = vgQuery & "'" & vgTipoTablaTotal & "',"
        vgQuery = vgQuery & "'" & vgTipoTablaBeneficiario & "',"
        vgQuery = vgQuery & "'" & vgTipoTablaRentista & "') "
        ''vgQuery = vgQuery & "cod_sexo = '" & iSexo & "' and "
        ''vgQuery = vgQuery & "#" & Format(iFecha, "yyyy/mm/dd") & "# beetwen "
        ''vgQuery = vgQuery & "fec_ini and fec_ter and "
        'I--- ABV 14/04/2005 ---
        vgQuery = vgQuery & "and cod_estado = 'A' "
        vgQuery = vgQuery & "and cod_oficial = 'S' "
        'F--- ABV 14/04/2005 ---
        Set vgCmb = vgConexionBD.Execute(vgQuery)
        If Not (vgCmb.EOF) Then
            If Not IsNull(vgCmb!numero) Then
                vgNumeroTotalTablas = vgCmb!numero
            End If
        End If
        vgCmb.Close
        
        If (vgNumeroTotalTablas <> 0) Then
            ReDim egTablaMortal(vgNumeroTotalTablas) As TypeTabla
            
            iSql = "SELECT num_correlativo,cod_tiptabmor,cod_sexo,fec_ini,fec_ter,gls_nombre,"
            iSql = iSql & "cod_tipogen,cod_tipoper,num_initabmor,num_tertabmor,prc_tasaint,cod_estado "
            iSql = iSql & ",cod_oficial "
            'I--- ABV 15/03/2005 ---
            iSql = iSql & ",cod_tipotabla,num_annobase,gls_descripcion "
            'F--- ABV 15/03/2005 ---
            'iSql = "SELECT * "
            iSql = iSql & "from MA_TVAL_MORTAL WHERE "
            iSql = iSql & "cod_tipoper = '" & vgTipoPeriodo & "' "
            iSql = iSql & "and "
            iSql = iSql & "cod_tiptabmor in ("
            iSql = iSql & "'" & vgTipoTablaParcial & "',"
            iSql = iSql & "'" & vgTipoTablaTotal & "',"
            iSql = iSql & "'" & vgTipoTablaBeneficiario & "',"
            iSql = iSql & "'" & vgTipoTablaRentista & "') "
            ''vgQuery = vgQuery & "cod_sexo = '" & iSexo & "' and "
            ''vgQuery = vgQuery & "#" & Format(iFecha, "yyyy/mm/dd") & "# beetwen "
            ''vgQuery = vgQuery & "fec_ini and fec_ter and "
            'vgQuery = vgQuery & "and cod_estado = 'A' "
            'iSql = iSql & "ORDER BY gls_nombre,fec_ini "
            'Debug.Print iSql
'I--- ABV 17/11/2005 ---
            iSql = iSql & "and cod_estado = 'A' "
            iSql = iSql & "and cod_oficial = 'S' "
'F--- ABV 17/11/2005 ---
            Set iRegistro = vgConexionBD.Execute(iSql)
            If Not (iRegistro.EOF) Then
                'iRegistro.MoveFirst
                'ReDim TablaMortal(vgCmb!Numero)
                vgX = 1
                While Not (iRegistro.EOF)
                    egTablaMortal(vgX).Correlativo = iRegistro!num_correlativo
                    egTablaMortal(vgX).TipoTabla = iRegistro!cod_tiptabmor
                    egTablaMortal(vgX).Sexo = iRegistro!Cod_Sexo
                    egTablaMortal(vgX).FechaIni = iRegistro!fec_ini
                    egTablaMortal(vgX).FechaFin = iRegistro!fec_ter
                    egTablaMortal(vgX).Nombre = Trim(iRegistro!gls_nombre)
                    egTablaMortal(vgX).TipoGenerar = iRegistro!cod_tipogen
                    egTablaMortal(vgX).TipoPeriodo = iRegistro!cod_tipoper
                    egTablaMortal(vgX).IniTab = iRegistro!num_initabmor
                    egTablaMortal(vgX).Fintab = iRegistro!num_tertabmor
                    egTablaMortal(vgX).tasa = iRegistro!prc_tasaint
                    egTablaMortal(vgX).Oficial = iRegistro!cod_oficial
                    egTablaMortal(vgX).Estado = iRegistro!Cod_Estado
                    egTablaMortal(vgX).TipoMovimiento = IIf(IsNull(iRegistro!cod_tipotabla), "E", iRegistro!cod_tipotabla)
                    egTablaMortal(vgX).AñoBase = IIf(IsNull(iRegistro!num_annobase), "0", iRegistro!num_annobase)
                    egTablaMortal(vgX).Descripcion = Trim(iRegistro!gls_descripcion)
                    vgX = vgX + 1
                    iRegistro.MoveNext
                Wend
            End If
            iRegistro.Close
        End If
    End If
    
    'Call CerrarBaseDeDatos_Aux
'''     Call CerrarBaseDeDatos(vgConectarBD)
Exit Function
Err_Tabla:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fgCargarTablaMortalidad_Old(ioPeriodo)
Dim vgCmb As ADODB.Recordset
Dim iRegistro As ADODB.Recordset
Dim iSql As String
On Error GoTo Err_Tabla

    'Call AbrirBaseDeDatos_Aux(vgRutaBasedeDatos)
'''    If Not AbrirBaseDeDatos(vgConectarBD) Then
'''        MsgBox "Falló la conexión con la Base de Datos.", vbCritical, "Error de Conexión"
'''        Exit Function
'''    End If
    
    'If (iFecha <> "") And (iPeriodo <> "") And (iSexo <> "") And (iTipoTabla <> "") Then
    If (vgTipoPeriodo <> "") Then
        
        vgNumeroTotalTablas = 0
        
        vgQuery = "SELECT count(num_correlativo) as numero "
        vgQuery = vgQuery & "from MA_TVAL_MORTAL WHERE "
        vgQuery = vgQuery & "cod_tipoper = '" & vgTipoPeriodo & "' "
        vgQuery = vgQuery & "and "
        vgQuery = vgQuery & "cod_tiptabmor in ("
        vgQuery = vgQuery & "'" & vgTipoTablaParcial & "',"
        vgQuery = vgQuery & "'" & vgTipoTablaTotal & "',"
        vgQuery = vgQuery & "'" & vgTipoTablaBeneficiario & "',"
        vgQuery = vgQuery & "'" & vgTipoTablaRentista & "') "
        ''vgQuery = vgQuery & "cod_sexo = '" & iSexo & "' and "
        ''vgQuery = vgQuery & "#" & Format(iFecha, "yyyy/mm/dd") & "# beetwen "
        ''vgQuery = vgQuery & "fec_ini and fec_ter and "
        'I--- ABV 14/04/2005 ---
        vgQuery = vgQuery & "and cod_estado = 'A' "
        vgQuery = vgQuery & "and cod_oficial = 'S' "
        'F--- ABV 14/04/2005 ---
        Set vgCmb = vgConexionBD.Execute(vgQuery)
        If Not (vgCmb.EOF) Then
            If Not IsNull(vgCmb!numero) Then
                vgNumeroTotalTablas = vgCmb!numero
            End If
        End If
        vgCmb.Close
        
        If (vgNumeroTotalTablas <> 0) Then
            ReDim egTablaMortal(vgNumeroTotalTablas) As TypeTabla
            
            iSql = "SELECT num_correlativo,cod_tiptabmor,cod_sexo,fec_ini,fec_ter,gls_nombre,"
            iSql = iSql & "cod_tipogen,cod_tipoper,num_initabmor,num_tertabmor,prc_tasaint,cod_estado "
            iSql = iSql & ",cod_oficial "
            'I--- ABV 15/03/2005 ---
            iSql = iSql & ",cod_tipotabla,num_annobase "
            'F--- ABV 15/03/2005 ---
            'iSql = "SELECT * "
            iSql = iSql & "from MA_TVAL_MORTAL WHERE "
            iSql = iSql & "cod_tipoper = '" & vgTipoPeriodo & "' "
            iSql = iSql & "and "
            iSql = iSql & "cod_tiptabmor in ("
            iSql = iSql & "'" & vgTipoTablaParcial & "',"
            iSql = iSql & "'" & vgTipoTablaTotal & "',"
            iSql = iSql & "'" & vgTipoTablaBeneficiario & "',"
            iSql = iSql & "'" & vgTipoTablaRentista & "') "
            ''vgQuery = vgQuery & "cod_sexo = '" & iSexo & "' and "
            ''vgQuery = vgQuery & "#" & Format(iFecha, "yyyy/mm/dd") & "# beetwen "
            ''vgQuery = vgQuery & "fec_ini and fec_ter and "
            'vgQuery = vgQuery & "and cod_estado = 'A' "
            'iSql = iSql & "ORDER BY gls_nombre,fec_ini "
            'Debug.Print iSql
            Set iRegistro = vgConexionBD.Execute(iSql)
            If Not (iRegistro.EOF) Then
                'iRegistro.MoveFirst
                'ReDim TablaMortal(vgCmb!Numero)
                vgX = 1
                While Not (iRegistro.EOF)
                    egTablaMortal(vgX).Correlativo = iRegistro!num_correlativo
                    egTablaMortal(vgX).TipoTabla = iRegistro!cod_tiptabmor
                    egTablaMortal(vgX).Sexo = iRegistro!Cod_Sexo
                    egTablaMortal(vgX).FechaIni = iRegistro!fec_ini
                    egTablaMortal(vgX).FechaFin = iRegistro!fec_ter
                    egTablaMortal(vgX).Nombre = Trim(iRegistro!gls_nombre)
                    egTablaMortal(vgX).TipoGenerar = iRegistro!cod_tipogen
                    egTablaMortal(vgX).TipoPeriodo = iRegistro!cod_tipoper
                    egTablaMortal(vgX).IniTab = iRegistro!num_initabmor
                    egTablaMortal(vgX).Fintab = iRegistro!num_tertabmor
                    egTablaMortal(vgX).tasa = iRegistro!prc_tasaint
                    egTablaMortal(vgX).Oficial = iRegistro!cod_oficial
                    egTablaMortal(vgX).Estado = iRegistro!Cod_Estado
                    egTablaMortal(vgX).TipoMovimiento = IIf(IsNull(iRegistro!cod_tipotabla), "E", iRegistro!cod_tipotabla)
                    egTablaMortal(vgX).AñoBase = IIf(IsNull(iRegistro!num_annobase), "0", iRegistro!num_annobase)
                    vgX = vgX + 1
                    iRegistro.MoveNext
                Wend
            End If
            iRegistro.Close
        End If
    End If
    
    'Call CerrarBaseDeDatos_Aux
'''     Call CerrarBaseDeDatos(vgConectarBD)
Exit Function
Err_Tabla:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fgCrearMortalidadDinamica_Old(iNavig, iNmvig, iNdvig, _
    iab, imb, idb, iSexo, iInval, iCorrelativo, iFinTab, iAñoBase, iFNac) As Boolean
'Función: permite generar la Tabla de Mortalidad Dinámica desde la descripción de una
'         Tabla Anual
'Parámetros de Entrada:
'   - iNavig = Año de Vigencia de la Tabla de Mortalidad
'   - iNmvig = Mes de Vigencia de la Tabla de Mortalidad
'   - iNdvig = Día de Vigencia de la Tabla de Mortalidad
'   - iab    = Año de Proceso
'   - imb    = Mes de Proceso
'   - idb    = Día de Proceso
'   - iSexo  = Sexo a generar
'   - iInval = Invalidez a generar
'   - iCorrelativo = Número de la Tabla de Mortalidad a Leer
'   - iFinTab = Valor que indica el termino de la Tabla de Mortalidad
'   - iAñoBase = Indica el Año Base desde el cual se genera la Tabla Dinámica
'Valor de Salida:
'   Retorna un valor True o False si pudo realizar la Actualización de la Matriz

'Dim tledad(0 To 90), qx(0 To 90), facmejor(0 To 90)
Dim Lxm(1 To 2, 1 To 3, 1 To 1332) As Double
Dim QxModif(0 To 120) As Double
Dim Ajuste(0 To 120) As Double, TqEdad(0 To 120) As Double
Dim AñoBase As Long, AñoProceso As Long
Dim QxMensModif As Double, Parte1 As Double, Parte2 As Double
Dim vliI As Long, vliJ As Long, vliK As Long
Dim Tb2 As ADODB.Recordset

Dim Fechap   As Long, Fechan As Long
Dim Edad     As Long, difdia As Long
Dim edaca    As Long, factor1 As Long
Dim vls_sexo As String

'On Error GoTo Err_Mortal

    fgCrearMortalidadDinamica = False
    AñoBase = iAñoBase
    AñoProceso = iNavig
    Fechap = iab * 12 + imb
    'Fechap = iNavig * 12 + iNmvig
    Fechan = CInt(Mid(iFNac, 1, 4)) * 12 + CInt(Mid(iFNac, 5, 2))
    Edad = Fechap - Fechan
    difdia = idb - CInt(Mid(iFNac, 7, 2))
    'difdia = iNdvig - CInt(Mid(iFNac, 7, 2))
    If difdia > 15 Then Edad = Edad + 1
    If Edad <= 240 Then Edad = 240
    If Edad > (110 * 12) Then
        vgError = 1023
        Exit Function
    End If
    edaca = Fix(Edad / 12)
    vls_sexo = iSexo
    vliI = 0
    
    'Lectura de tabla de mortalidad
    vgSql = "SELECT num_edad AS edad, mto_qx AS qx, prc_factor AS factor "
    vgSql = vgSql & "FROM ma_tval_mordet "
    vgSql = vgSql & "WHERE num_correlativo = " & vgs_Nro & " "
    vgSql = vgSql & "ORDER BY num_edad "
    Set Tb2 = vgConexionBD.Execute(vgSql)
    If Not (Tb2.EOF) Then
        Do While Not Tb2.EOF
            vliK = Tb2!Edad
            'If (vliK >= 20) Then
                TqEdad(vliK) = Tb2!Qx
                Ajuste(vliK) = Tb2!Factor
                vliI = vliI + 1
            'End If
            Tb2.MoveNext
        Loop
    Else
        vgError = 1061
        Tb2.Close
        Exit Function
    End If
    Tb2.Close
   vliJ = -1
   vliI = edaca
    For vliI = edaca To 110 '- edaca)
        vliJ = vliJ + 1
        QxModif(vliI) = TqEdad(vliI) * (1 - Ajuste(vliI)) ^ (vliJ + (AñoProceso - AñoBase))
    Next vliI
    
    QxMensModif = 0
    factor1 = 0 'Edad - 1
    'For vliI= 20 To (iFinTab / 12)
    For vliI = edaca To 110 '- edaca)
        For vliJ = 1 To 12
            vliK = factor1 + ((vliI - 1) * 12) + vliJ
            If vliK >= Edad Then
                If vliK = Edad Then
                    Lxm(iSexo, iInval, vliK) = 100000
                Else
                    Lxm(iSexo, iInval, vliK) = Lxm(iSexo, iInval, vliK - 1) - (Lxm(iSexo, iInval, vliK - 1) * QxMensModif)
                End If
                Parte1 = ((1 / 12) * QxModif(vliI - 1))
                Parte2 = (vliK / 12 - Fix(vliK / 12))
                If ((1 - Parte2 * QxModif(vliI - 1)) = 0) Then
                    QxMensModif = 0
                Else
                    QxMensModif = Parte1 / (1 - Parte2 * QxModif(vliI - 1))
                End If
                Lx(iSexo, iInval, vliK) = Lxm(iSexo, iInval, vliK)
                If vliK > (110 * 12) Then Exit For
            
''''            'Borra - Daniela
''''''            vgQuery = "INSERT into DANI (agno ,edad,qx ,lx,sexo) values( "
''''''            vgQuery = vgQuery & Str(Format(iNavig, "#0")) & ", "
''''''            vgQuery = vgQuery & Str(Format(vliK, "#000")) & ", "
''''''            vgQuery = vgQuery & Str(Format(QxMensModif, "#0.0000000000000")) & ", "
''''''            vgQuery = vgQuery & Str(Format(Lxm(iSexo, iInval, vliK), "#000.000000000")) & ", "
''''''            vgQuery = vgQuery & Str(Format(iSexo, "#0")) & ") "
''''''            vgConexionBD.Execute (vgQuery)

            'vgQuery = "update DANI set  "
            'vgQuery = vgQuery & "agno = " & Str(Format(iNavig, "#0")) & ", "
            'vgQuery = vgQuery & "edad = " & Str(Format(vliK, "#000")) & ", "
            'vgQuery = vgQuery & "qx = " & Str(Format(QxMensModif, "#0.0000000")) & ", "
            'vgQuery = vgQuery & "lx = " & Str(Format(Lxm(iSexo, iInval, vliK), "#000.000")) & " "
            'vgConexionBD.Execute (vgQuery)
            End If
            
            
            
        Next vliJ
    Next vliI
    
    fgCrearMortalidadDinamica = True
    
Exit Function   'Buscar otra Póliza a calcular
Err_Mortal:
    'Screen.MousePointer = 0
    Select Case Err
        Case Else
        'ProgressBar.Value = 0
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fgBuscarMortalidad_Old(iNavig, iNmvig, iNdvig, iNap, iNmp, iNdp, iSexoCau, iFechaNacCau) As Boolean
Dim h As Integer, i As Integer, j As Integer, k As Integer
Dim Sql As String

'    fgBuscarMortalidad = False
'
'    'If (iNavig = 2003) Then
'    '    iNdvig = iNdvig
'    'End If
'
'    '1. Leer Tabla de Mortalidad de Rtas. Vitalicias Mujer
'    vgBuscarMortalVit_F = ""
'    If (vgFechaIniMortalVit_F <> "") And (vgFechaFinMortalVit_F <> "") Then
'        If ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) >= (vgFechaIniMortalVit_F)) And _
'        ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) <= (vgFechaFinMortalVit_F)) Then
'            If (vgIndicadorTipoMovimiento_F = "E") Or (vgIndicadorTipoMovimiento_F <> "E" And iSexoCau = "M") Then
'                vgBuscarMortalVit_F = "N"
'            Else
'                vgBuscarMortalVit_F = "S"
'            End If
'        Else
'            vgBuscarMortalVit_F = "S"
'        End If
'    Else
'        vgBuscarMortalVit_F = "S"
'    End If
'
'    If (vgBuscarMortalVit_F = "S") Then
'        'For h = 1 To 2  '1=Causante '2=Beneficiario
'        h = 1
'            vgs_Sexo = ""
'            vgs_Tipo = ""
'            vgs_Nro = ""
'            'For i = 1 To 2
'            '    If i = 1 Then vgs_Sexo = "M"    ' hombre  (M)asculino
'            '    If i = 2 Then vgs_Sexo = "F"    ' mujer   (F)emenino
'                i = 2
'                vgs_Sexo = "F"
'
'                'For j = 1 To 3
'                j = 2
'                    'If j = 1 Then vgs_Tipo = "MIT"    '1  = Invalido Total
'                    If j = 2 Then vgs_Tipo = "RV"    '2  = No Invalido
'                    'If h = 2 And j = 2 Then vgs_Tipo = "B" ' Tabla de beneficiarios no invalidos
'                    'If j = 3 Then vgs_Tipo = "MIP"    '3  = invalido Parcial
'
'                vgSw = False
'
'                'Buscar Número Correlativo desde la Estructura
'                For vgI = 1 To vgNumeroTotalTablas
'
'                    'Sql = " SELECT * "
'                    'Sql = Sql & " from PR_TVAL_MORTAL "
'                    'Sql = Sql & " where cod_sexo = '" & vgs_Sexo & "' and "
'                    'Sql = Sql & " cod_tiptabmor ='" & vgs_Tipo & "' and "
'                    'Sql = Sql & " fec_ini <= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "# AND "
'                    'Sql = Sql & " fec_ter >= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "#"
'                    'Set tb = vgConexionBD.Execute(Sql)
'                    'If Not (tb.EOF) Then
'
'                    If (egTablaMortal(vgI).Sexo = vgs_Sexo) And _
'                    (egTablaMortal(vgI).TipoTabla = vgs_Tipo) And _
'                    (egTablaMortal(vgI).Nombre = vgPalabra_MortalVit_F) And _
'                    (egTablaMortal(vgI).FechaIni <= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) And _
'                    (egTablaMortal(vgI).FechaFin >= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) Then
'
'                        vgs_Error = 0
'                        vgs_Nro = egTablaMortal(vgI).Correlativo
'                        vgFechaIniMortalVit_F = egTablaMortal(vgI).FechaIni
'                        vgFechaFinMortalVit_F = egTablaMortal(vgI).FechaFin
'                        vgIndicadorTipoMovimiento_F = egTablaMortal(vgI).TipoMovimiento
'                        vgDinamicaAñoBase_F = egTablaMortal(vgI).AñoBase
'
'                        'Limpiar columna de Datos
'                        For vgX = 1 To Fintab
'                            Lx(i, j, vgX) = 0
'                        Next vgX
'
'                        'vgs_Nro = tb!num_correlativo
'                        'vgFechaInicioMortal = Format(tb!fec_ini, "yyyymmdd")
'                        'vgFechaFinMortal = Format(tb!fec_ter, "yyyymmdd")
'                        ''vgFechaInicioMortal = DateSerial(Year(tb!fec_ini), Month(tb!fec_ini), Day(tb!fec_ini))
'                        ''vgFechaFinMortal = DateSerial(Year(tb!fec_ter), Month(tb!fec_ter), Day(tb!fec_ter))
'                        If vgs_Nro <> 0 Then
'                            If (vgIndicadorTipoMovimiento_F <> "D") Then
'                                Sql = "Select num_edad AS edad,mto_lx from MA_TVAL_MORDET "
'                                Sql = Sql & " where num_correlativo = " & vgs_Nro
'                                Sql = Sql & " order by num_edad "
'                                Set Tb2 = vgConexionBD.Execute(Sql)
'                                If Not (Tb2.EOF) Then
'                                    vgSw = True
'                                    'tb2.MoveFirst
'                                    'k = 1
'                                    Do While Not Tb2.EOF
'                                        k = Tb2!Edad
'                                        'If h = 1 Then   'Causante
'                                            Lx(i, j, k) = Tb2!mto_lx
'                                        'Else    'Beneficiario
'                                        '    ly(i, j, k) = tb2!mto_lx
'                                        'End If
'                                        'k = k + 1
'                                        Tb2.MoveNext
'                                    Loop
'                                Else
'                                    vgError = 1061
'                                    Exit Function
'                                End If
'                                Tb2.Close
'                            Else
'                                'Obtener la Tabla Temporal Dinánica
'                                If (fgCrearMortalidadDinamica(iNavig, iNmvig, iNdvig, _
'                                iNap, iNmp, iNdp, i, j, vgs_Nro, Fintab, vgDinamicaAñoBase_F, iFechaNacCau) = False) Then
'                                    vgError = 1061
'                                    Exit Function
'                                Else
'                                    vgSw = True
'                                End If
'                            End If
'                        Else
'                            vgError = 1061
'                            Exit Function
'                        End If
'                    'Else
'                    '    vgError = 1061
'                    '    Exit Function
'                    End If
'                    'tb.Close
'                Next vgI
'                If (vgSw = False) Then
'                    vgError = 1061
'                    Exit Function
'                End If
'                'Next j
'            'Next i
'        'Next h
'    End If
'
'    '2. Leer Tabla de Mortalidad de Inv. Totales Mujer
'    vgBuscarMortalTot_F = ""
'    If (vgFechaIniMortalTot_F <> "") And (vgFechaFinMortalTot_F <> "") Then
'        If ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) >= (vgFechaIniMortalTot_F)) And _
'        ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) <= (vgFechaFinMortalTot_F)) Then
'            vgBuscarMortalTot_F = "N"
'        Else
'            vgBuscarMortalTot_F = "S"
'        End If
'    Else
'        vgBuscarMortalTot_F = "S"
'    End If
'
'    If (vgBuscarMortalTot_F = "S") Then
'        For h = 1 To 2  '1=Causante '2=Beneficiario
'            vgs_Sexo = ""
'            vgs_Tipo = ""
'            vgs_Nro = ""
'            'For i = 1 To 2
'            '    If i = 1 Then vgs_Sexo = "M"    ' hombre  (M)asculino
'            '    If i = 2 Then vgs_Sexo = "F"    ' mujer   (F)emenino
'                i = 2
'                vgs_Sexo = "F"
'
'                'For j = 1 To 3
'                j = 1
'                    If j = 1 Then vgs_Tipo = "MIT"    '1  = Invalido Total
'                    'If j = 2 Then vgs_Tipo = "RV"    '2  = No Invalido
'                    'If h = 2 And j = 2 Then vgs_Tipo = "B" ' Tabla de beneficiarios no invalidos
'                    'If j = 3 Then vgs_Tipo = "MIP"    '3  = invalido Parcial
'
'                vgSw = False
'
'                'Buscar Número Correlativo desde la Estructura
'                For vgI = 1 To vgNumeroTotalTablas
'
'                    'Sql = " SELECT * "
'                    'Sql = Sql & " from PR_TVAL_MORTAL "
'                    'Sql = Sql & " where cod_sexo = '" & vgs_Sexo & "' and "
'                    'Sql = Sql & " cod_tiptabmor ='" & vgs_Tipo & "' and "
'                    'Sql = Sql & " fec_ini <= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "# AND "
'                    'Sql = Sql & " fec_ter >= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "#"
'                    'Set tb = vgConexionBD.Execute(Sql)
'                    'If Not (tb.EOF) Then
'
'                    If (egTablaMortal(vgI).Sexo = vgs_Sexo) And _
'                    (egTablaMortal(vgI).TipoTabla = vgs_Tipo) And _
'                    (egTablaMortal(vgI).Nombre = vgPalabra_MortalTot_F) And _
'                    (egTablaMortal(vgI).FechaIni <= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) And _
'                    (egTablaMortal(vgI).FechaFin >= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) Then
'
'                        vgs_Error = 0
'                        vgs_Nro = egTablaMortal(vgI).Correlativo
'                        vgFechaIniMortalTot_F = egTablaMortal(vgI).FechaIni
'                        vgFechaFinMortalTot_F = egTablaMortal(vgI).FechaFin
'
'                        'Limpiar columna de Datos
'                        If (h = 1) Then
'                            For vgX = 1 To Fintab
'                                Lx(i, j, vgX) = 0
'                            Next vgX
'                        Else
'                            For vgX = 1 To Fintab
'                                Ly(i, j, vgX) = 0
'                            Next vgX
'                        End If
'
'                        'vgs_Nro = tb!num_correlativo
'                        'vgFechaInicioMortal = Format(tb!fec_ini, "yyyymmdd")
'                        'vgFechaFinMortal = Format(tb!fec_ter, "yyyymmdd")
'                        ''vgFechaInicioMortal = DateSerial(Year(tb!fec_ini), Month(tb!fec_ini), Day(tb!fec_ini))
'                        ''vgFechaFinMortal = DateSerial(Year(tb!fec_ter), Month(tb!fec_ter), Day(tb!fec_ter))
'                        If vgs_Nro <> 0 Then
'                            Sql = "Select num_edad AS edad,mto_lx from MA_TVAL_MORDET "
'                            Sql = Sql & " where num_correlativo = " & vgs_Nro
'                            Sql = Sql & " order by num_edad "
'                            Set Tb2 = vgConexionBD.Execute(Sql)
'                            If Not (Tb2.EOF) Then
'                                vgSw = True
'                                'tb2.MoveFirst
'                                'k = 1
'                                Do While Not Tb2.EOF
'                                    k = Tb2!Edad
'                                    If h = 1 Then   'Causante
'                                        Lx(i, j, k) = Tb2!mto_lx
'                                    Else    'Beneficiario
'                                        Ly(i, j, k) = Tb2!mto_lx
'                                    End If
'                                    'k = k + 1
'                                    Tb2.MoveNext
'                                Loop
'                            Else
'                                vgError = 1062
'                                Exit Function
'                            End If
'                            Tb2.Close
'                        Else
'                            vgError = 1062
'                            Exit Function
'                        End If
'                    'Else
'                    '    vgError = 1062
'                    '    Exit Function
'                    End If
'                    'tb.Close
'                Next vgI
'                If (vgSw = False) Then
'                    vgError = 1062
'                    Exit Function
'                End If
'                'Next j
'            'Next i
'        Next h
'    End If
'
'    '3. Leer Tabla de Mortalidad de Inv. Parciales Mujer
'    vgBuscarMortalPar_F = ""
'    If (vgFechaIniMortalPar_F <> "") And (vgFechaFinMortalPar_F <> "") Then
'        If ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) >= (vgFechaIniMortalPar_F)) And _
'        ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) <= (vgFechaFinMortalPar_F)) Then
'            vgBuscarMortalPar_F = "N"
'        Else
'            vgBuscarMortalPar_F = "S"
'        End If
'    Else
'        vgBuscarMortalPar_F = "S"
'    End If
'
'    If (vgBuscarMortalPar_F = "S") Then
'        For h = 1 To 2  '1=Causante '2=Beneficiario
'            vgs_Sexo = ""
'            vgs_Tipo = ""
'            vgs_Nro = ""
'            'For i = 1 To 2
'            '    If i = 1 Then vgs_Sexo = "M"    ' hombre  (M)asculino
'            '    If i = 2 Then vgs_Sexo = "F"    ' mujer   (F)emenino
'                i = 2
'                vgs_Sexo = "F"
'
'                'For j = 1 To 3
'                j = 3
'                    'If j = 1 Then vgs_Tipo = "MIT"    '1  = Invalido Total
'                    'If j = 2 Then vgs_Tipo = "RV"    '2  = No Invalido
'                    'If h = 2 And j = 2 Then vgs_Tipo = "B" ' Tabla de beneficiarios no invalidos
'                    If j = 3 Then vgs_Tipo = "MIP"    '3  = invalido Parcial
'
'                vgSw = False
'
'                'Buscar Número Correlativo desde la Estructura
'                For vgI = 1 To vgNumeroTotalTablas
'
'                    'Sql = " SELECT * "
'                    'Sql = Sql & " from PR_TVAL_MORTAL "
'                    'Sql = Sql & " where cod_sexo = '" & vgs_Sexo & "' and "
'                    'Sql = Sql & " cod_tiptabmor ='" & vgs_Tipo & "' and "
'                    'Sql = Sql & " fec_ini <= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "# AND "
'                    'Sql = Sql & " fec_ter >= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "#"
'                    'Set tb = vgConexionBD.Execute(Sql)
'                    'If Not (tb.EOF) Then
'
'                    If (egTablaMortal(vgI).Sexo = vgs_Sexo) And _
'                    (egTablaMortal(vgI).TipoTabla = vgs_Tipo) And _
'                    (egTablaMortal(vgI).Nombre = vgPalabra_MortalPar_F) And _
'                    (egTablaMortal(vgI).FechaIni <= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) And _
'                    (egTablaMortal(vgI).FechaFin >= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) Then
'
'                        vgs_Error = 0
'                        vgs_Nro = egTablaMortal(vgI).Correlativo
'                        vgFechaIniMortalPar_F = egTablaMortal(vgI).FechaIni
'                        vgFechaFinMortalPar_F = egTablaMortal(vgI).FechaFin
'
'                        'Limpiar columna de Datos
'                        If (h = 1) Then
'                            For vgX = 1 To Fintab
'                                Lx(i, j, vgX) = 0
'                            Next vgX
'                        Else
'                            For vgX = 1 To Fintab
'                                Ly(i, j, vgX) = 0
'                            Next vgX
'                        End If
'
'                        'vgs_Nro = tb!num_correlativo
'                        'vgFechaInicioMortal = Format(tb!fec_ini, "yyyymmdd")
'                        'vgFechaFinMortal = Format(tb!fec_ter, "yyyymmdd")
'                        ''vgFechaInicioMortal = DateSerial(Year(tb!fec_ini), Month(tb!fec_ini), Day(tb!fec_ini))
'                        ''vgFechaFinMortal = DateSerial(Year(tb!fec_ter), Month(tb!fec_ter), Day(tb!fec_ter))
'                        If vgs_Nro <> 0 Then
'                            Sql = "Select num_edad AS edad,mto_lx from MA_TVAL_MORDET "
'                            Sql = Sql & " where num_correlativo = " & vgs_Nro
'                            Sql = Sql & " order by num_edad "
'                            Set Tb2 = vgConexionBD.Execute(Sql)
'                            If Not (Tb2.EOF) Then
'                                vgSw = True
'                                'tb2.MoveFirst
'                                'k = 1
'                                Do While Not Tb2.EOF
'                                    k = Tb2!Edad
'                                    If h = 1 Then   'Causante
'                                        Lx(i, j, k) = Tb2!mto_lx
'                                    Else    'Beneficiario
'                                        Ly(i, j, k) = Tb2!mto_lx
'                                    End If
'                                    'k = k + 1
'                                    Tb2.MoveNext
'                                Loop
'                            Else
'                                vgError = 1063
'                                Exit Function
'                            End If
'                            Tb2.Close
'                        Else
'                            vgError = 1063
'                            Exit Function
'                        End If
'                    'Else
'                    '    vgError = 1063
'                    '    Exit Function
'                    End If
'                    'tb.Close
'                Next vgI
'                If (vgSw = False) Then
'                    vgError = 1063
'                    Exit Function
'                End If
'                'Next j
'            'Next i
'        Next h
'    End If
'
'    '4. Leer Tabla de Mortalidad de Beneficiarios Mujer
'    vgBuscarMortalBen_F = ""
'    If (vgFechaIniMortalBen_F <> "") And (vgFechaFinMortalBen_F <> "") Then
'        If ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) >= (vgFechaIniMortalBen_F)) And _
'        ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) <= (vgFechaFinMortalBen_F)) Then
'            vgBuscarMortalBen_F = "N"
'        Else
'            vgBuscarMortalBen_F = "S"
'        End If
'    Else
'        vgBuscarMortalBen_F = "S"
'    End If
'
'    If (vgBuscarMortalBen_F = "S") Then
'        'For h = 1 To 2  '1=Causante '2=Beneficiario
'        h = 2
'            vgs_Sexo = ""
'            vgs_Tipo = ""
'            vgs_Nro = ""
'            'For i = 1 To 2
'            '    If i = 1 Then vgs_Sexo = "M"    ' hombre  (M)asculino
'            '    If i = 2 Then vgs_Sexo = "F"    ' mujer   (F)emenino
'                i = 2
'                vgs_Sexo = "F"
'
'                'For j = 1 To 3
'                j = 2
'                    'If j = 1 Then vgs_Tipo = "MIT"    '1  = Invalido Total
'                    'If j = 2 Then vgs_Tipo = "RV"    '2  = No Invalido
'                    If h = 2 And j = 2 Then vgs_Tipo = "B" ' Tabla de beneficiarios no invalidos
'                    'If j = 3 Then vgs_Tipo = "MIP"    '3  = invalido Parcial
'
'                vgSw = False
'
'                'Buscar Número Correlativo desde la Estructura
'                For vgI = 1 To vgNumeroTotalTablas
'
'                    'Sql = " SELECT * "
'                    'Sql = Sql & " from PR_TVAL_MORTAL "
'                    'Sql = Sql & " where cod_sexo = '" & vgs_Sexo & "' and "
'                    'Sql = Sql & " cod_tiptabmor ='" & vgs_Tipo & "' and "
'                    'Sql = Sql & " fec_ini <= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "# AND "
'                    'Sql = Sql & " fec_ter >= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "#"
'                    'Set tb = vgConexionBD.Execute(Sql)
'                    'If Not (tb.EOF) Then
'
'                    If (egTablaMortal(vgI).Sexo = vgs_Sexo) And _
'                    (egTablaMortal(vgI).TipoTabla = vgs_Tipo) And _
'                    (egTablaMortal(vgI).Nombre = vgPalabra_MortalBen_F) And _
'                    (egTablaMortal(vgI).FechaIni <= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) And _
'                    (egTablaMortal(vgI).FechaFin >= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) Then
'
'                        vgs_Error = 0
'                        vgs_Nro = egTablaMortal(vgI).Correlativo
'                        vgFechaIniMortalBen_F = egTablaMortal(vgI).FechaIni
'                        vgFechaFinMortalBen_F = egTablaMortal(vgI).FechaFin
'
'                        'Limpiar columna de Datos
'                        For vgX = 1 To Fintab
'                            Ly(i, j, vgX) = 0
'                        Next vgX
'
'                        'vgs_Nro = tb!num_correlativo
'                        'vgFechaInicioMortal = Format(tb!fec_ini, "yyyymmdd")
'                        'vgFechaFinMortal = Format(tb!fec_ter, "yyyymmdd")
'                        ''vgFechaInicioMortal = DateSerial(Year(tb!fec_ini), Month(tb!fec_ini), Day(tb!fec_ini))
'                        ''vgFechaFinMortal = DateSerial(Year(tb!fec_ter), Month(tb!fec_ter), Day(tb!fec_ter))
'                        If vgs_Nro <> 0 Then
'                            Sql = "Select num_edad AS edad,mto_lx from MA_TVAL_MORDET "
'                            Sql = Sql & " where num_correlativo = " & vgs_Nro
'                            Sql = Sql & " order by num_edad "
'                            Set Tb2 = vgConexionBD.Execute(Sql)
'                            If Not (Tb2.EOF) Then
'                                vgSw = True
'                                'tb2.MoveFirst
'                                'k = 1
'                                Do While Not Tb2.EOF
'                                    k = Tb2!Edad
'                                    'If h = 1 Then   'Causante
'                                    '    lx(i, j, k) = tb2!mto_lx
'                                    'Else    'Beneficiario
'                                        Ly(i, j, k) = Tb2!mto_lx
'                                    'End If
'                                    'k = k + 1
'                                    Tb2.MoveNext
'                                Loop
'                            Else
'                                vgError = 1064
'                                Exit Function
'                            End If
'                            Tb2.Close
'                        Else
'                            vgError = 1064
'                            Exit Function
'                        End If
'                    'Else
'                    '    vgError = 1064
'                    '    Exit Function
'                    End If
'                    'tb.Close
'                Next vgI
'                If (vgSw = False) Then
'                    vgError = 1064
'                    Exit Function
'                End If
'                'Next j
'            'Next i
'        'Next h
'    End If
'
''--------------------------------------------------------------------
'    '5. Leer Tabla de Mortalidad de Rtas. Vitalicias Hombre
'    vgBuscarMortalVit_M = ""
'    If (vgFechaIniMortalVit_M <> "") And (vgFechaFinMortalVit_M <> "") Then
'        If ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) >= (vgFechaIniMortalVit_M)) And _
'        ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) <= (vgFechaFinMortalVit_M)) Then
'            If (vgIndicadorTipoMovimiento_M = "E") Or (vgIndicadorTipoMovimiento_M <> "E" And iSexoCau = "F") Then
'                vgBuscarMortalVit_M = "N"
'            Else
'                vgBuscarMortalVit_M = "S"
'            End If
'        Else
'            vgBuscarMortalVit_M = "S"
'        End If
'    Else
'        vgBuscarMortalVit_M = "S"
'    End If
'
'    If (vgBuscarMortalVit_M = "S") Then
'        'For h = 1 To 2  '1=Causante '2=Beneficiario
'        h = 1
'            vgs_Sexo = ""
'            vgs_Tipo = ""
'            vgs_Nro = ""
'            'For i = 1 To 2
'            '    If i = 1 Then vgs_Sexo = "M"    ' hombre  (M)asculino
'            '    If i = 2 Then vgs_Sexo = "F"    ' mujer   (F)emenino
'                i = 1
'                vgs_Sexo = "M"
'
'                'For j = 1 To 3
'                j = 2
'                    'If j = 1 Then vgs_Tipo = "MIT"    '1  = Invalido Total
'                    If j = 2 Then vgs_Tipo = "RV"    '2  = No Invalido
'                    'If h = 2 And j = 2 Then vgs_Tipo = "B" ' Tabla de beneficiarios no invalidos
'                    'If j = 3 Then vgs_Tipo = "MIP"    '3  = invalido Parcial
'
'                vgSw = False
'
'                'Buscar Número Correlativo desde la Estructura
'                For vgI = 1 To vgNumeroTotalTablas
'
'                    'Sql = " SELECT * "
'                    'Sql = Sql & " from PR_TVAL_MORTAL "
'                    'Sql = Sql & " where cod_sexo = '" & vgs_Sexo & "' and "
'                    'Sql = Sql & " cod_tiptabmor ='" & vgs_Tipo & "' and "
'                    'Sql = Sql & " fec_ini <= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "# AND "
'                    'Sql = Sql & " fec_ter >= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "#"
'                    'Set tb = vgConexionBD.Execute(Sql)
'                    'If Not (tb.EOF) Then
'
'                    If (egTablaMortal(vgI).Sexo = vgs_Sexo) And _
'                    (egTablaMortal(vgI).TipoTabla = vgs_Tipo) And _
'                    (egTablaMortal(vgI).Nombre = vgPalabra_MortalVit_M) And _
'                    (egTablaMortal(vgI).FechaIni <= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) And _
'                    (egTablaMortal(vgI).FechaFin >= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) Then
'
'                        vgs_Error = 0
'                        vgs_Nro = egTablaMortal(vgI).Correlativo
'                        vgFechaIniMortalVit_M = egTablaMortal(vgI).FechaIni
'                        vgFechaFinMortalVit_M = egTablaMortal(vgI).FechaFin
'                        vgIndicadorTipoMovimiento_M = egTablaMortal(vgI).TipoMovimiento
'                        vgDinamicaAñoBase_M = egTablaMortal(vgI).AñoBase
'
'                        'Limpiar columna de Datos
'                        For vgX = 1 To Fintab
'                            Lx(i, j, vgX) = 0
'                        Next vgX
'
'                        'vgs_Nro = tb!num_correlativo
'                        'vgFechaInicioMortal = Format(tb!fec_ini, "yyyymmdd")
'                        'vgFechaFinMortal = Format(tb!fec_ter, "yyyymmdd")
'                        ''vgFechaInicioMortal = DateSerial(Year(tb!fec_ini), Month(tb!fec_ini), Day(tb!fec_ini))
'                        ''vgFechaFinMortal = DateSerial(Year(tb!fec_ter), Month(tb!fec_ter), Day(tb!fec_ter))
'                        If vgs_Nro <> 0 Then
'                            If (vgIndicadorTipoMovimiento_M <> "D") Then
'                                Sql = "Select num_edad AS edad,mto_lx from MA_TVAL_MORDET "
'                                Sql = Sql & " where num_correlativo = " & vgs_Nro
'                                Sql = Sql & " order by num_edad "
'                                Set Tb2 = vgConexionBD.Execute(Sql)
'                                If Not (Tb2.EOF) Then
'                                    vgSw = True
'                                    'tb2.MoveFirst
'                                    'k = 1
'                                    Do While Not Tb2.EOF
'                                        k = Tb2!Edad
'                                        'If h = 1 Then   'Causante
'                                            Lx(i, j, k) = Tb2!mto_lx
'                                        'Else    'Beneficiario
'                                        '    ly(i, j, k) = tb2!mto_lx
'                                        'End If
'                                        'k = k + 1
'                                        Tb2.MoveNext
'                                    Loop
'                                Else
'                                    vgError = 1065
'                                    Exit Function
'                                End If
'                                Tb2.Close
'                            Else
'                                'Obtener la Tabla Temporal Dinánica
'                                If (fgCrearMortalidadDinamica(iNavig, iNmvig, iNdvig, _
'                                iNap, iNmp, iNdp, i, j, vgs_Nro, Fintab, vgDinamicaAñoBase_M, iFechaNacCau) = False) Then
'                                    vgError = 1061
'                                    Exit Function
'                                Else
'                                    vgSw = True
'                                End If
'                            End If
'                        Else
'                            vgError = 1065
'                            Exit Function
'                        End If
'                    'Else
'                    '    vgError = 1065
'                    '    Exit Function
'                    End If
'                    'tb.Close
'                Next vgI
'                If (vgSw = False) Then
'                    vgError = 1065
'                    Exit Function
'                End If
'                'Next j
'            'Next i
'        'Next h
'    End If
'
'    '6. Leer Tabla de Mortalidad de Inv. Totales Hombre
'    vgBuscarMortalTot_M = ""
'    If (vgFechaIniMortalTot_M <> "") And (vgFechaFinMortalTot_M <> "") Then
'        If ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) >= (vgFechaIniMortalTot_M)) And _
'        ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) <= (vgFechaFinMortalTot_M)) Then
'            vgBuscarMortalTot_M = "N"
'        Else
'            vgBuscarMortalTot_M = "S"
'        End If
'    Else
'        vgBuscarMortalTot_M = "S"
'    End If
'
'    If (vgBuscarMortalTot_M = "S") Then
'        For h = 1 To 2  '1=Causante '2=Beneficiario
'            vgs_Sexo = ""
'            vgs_Tipo = ""
'            vgs_Nro = ""
'            'For i = 1 To 2
'            '    If i = 1 Then vgs_Sexo = "M"    ' hombre  (M)asculino
'            '    If i = 2 Then vgs_Sexo = "F"    ' mujer   (F)emenino
'                i = 1
'                vgs_Sexo = "M"
'
'                'For j = 1 To 3
'                j = 1
'                    If j = 1 Then vgs_Tipo = "MIT"    '1  = Invalido Total
'                    'If j = 2 Then vgs_Tipo = "RV"    '2  = No Invalido
'                    'If h = 2 And j = 2 Then vgs_Tipo = "B" ' Tabla de beneficiarios no invalidos
'                    'If j = 3 Then vgs_Tipo = "MIP"    '3  = invalido Parcial
'
'                vgSw = False
'
'                'Buscar Número Correlativo desde la Estructura
'                For vgI = 1 To vgNumeroTotalTablas
'
'                    'Sql = " SELECT * "
'                    'Sql = Sql & " from PR_TVAL_MORTAL "
'                    'Sql = Sql & " where cod_sexo = '" & vgs_Sexo & "' and "
'                    'Sql = Sql & " cod_tiptabmor ='" & vgs_Tipo & "' and "
'                    'Sql = Sql & " fec_ini <= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "# AND "
'                    'Sql = Sql & " fec_ter >= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "#"
'                    'Set tb = vgConexionBD.Execute(Sql)
'                    'If Not (tb.EOF) Then
'
'                    If (egTablaMortal(vgI).Sexo = vgs_Sexo) And _
'                    (egTablaMortal(vgI).TipoTabla = vgs_Tipo) And _
'                    (egTablaMortal(vgI).Nombre = vgPalabra_MortalTot_M) And _
'                    (egTablaMortal(vgI).FechaIni <= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) And _
'                    (egTablaMortal(vgI).FechaFin >= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) Then
'
'                        vgs_Error = 0
'                        vgs_Nro = egTablaMortal(vgI).Correlativo
'                        vgFechaIniMortalTot_M = egTablaMortal(vgI).FechaIni
'                        vgFechaFinMortalTot_M = egTablaMortal(vgI).FechaFin
'
'                        'Limpiar columna de Datos
'                        If (h = 1) Then
'                            For vgX = 1 To Fintab
'                                Lx(i, j, vgX) = 0
'                            Next vgX
'                        Else
'                            For vgX = 1 To Fintab
'                                Ly(i, j, vgX) = 0
'                            Next vgX
'                        End If
'
'                        'vgs_Nro = tb!num_correlativo
'                        'vgFechaInicioMortal = Format(tb!fec_ini, "yyyymmdd")
'                        'vgFechaFinMortal = Format(tb!fec_ter, "yyyymmdd")
'                        ''vgFechaInicioMortal = DateSerial(Year(tb!fec_ini), Month(tb!fec_ini), Day(tb!fec_ini))
'                        ''vgFechaFinMortal = DateSerial(Year(tb!fec_ter), Month(tb!fec_ter), Day(tb!fec_ter))
'                        If vgs_Nro <> 0 Then
'                            Sql = "Select num_edad AS edad,mto_lx from MA_TVAL_MORDET "
'                            Sql = Sql & " where num_correlativo = " & vgs_Nro
'                            Sql = Sql & " order by num_edad "
'                            Set Tb2 = vgConexionBD.Execute(Sql)
'                            If Not (Tb2.EOF) Then
'                                vgSw = True
'                                'tb2.MoveFirst
'                                'k = 1
'                                Do While Not Tb2.EOF
'                                    k = Tb2!Edad
'                                    If h = 1 Then   'Causante
'                                        Lx(i, j, k) = Tb2!mto_lx
'                                    Else    'Beneficiario
'                                        Ly(i, j, k) = Tb2!mto_lx
'                                    End If
'                                    'k = k + 1
'                                    Tb2.MoveNext
'                                Loop
'                            Else
'                                vgError = 1066
'                                Exit Function
'                            End If
'                            Tb2.Close
'                        Else
'                            vgError = 1066
'                            Exit Function
'                        End If
'                    'Else
'                    '    vgError = 1066
'                    '    Exit Function
'                    End If
'                    'tb.Close
'                Next vgI
'                If (vgSw = False) Then
'                    vgError = 1066
'                    Exit Function
'                End If
'                'Next j
'            'Next i
'        Next h
'    End If
'
'    '7. Leer Tabla de Mortalidad de Inv. Parciales Hombre
'    vgBuscarMortalPar_M = ""
'    If (vgFechaIniMortalPar_M <> "") And (vgFechaFinMortalPar_M <> "") Then
'        If ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) >= (vgFechaIniMortalPar_M)) And _
'        ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) <= (vgFechaFinMortalPar_M)) Then
'            vgBuscarMortalPar_M = "N"
'        Else
'            vgBuscarMortalPar_M = "S"
'        End If
'    Else
'        vgBuscarMortalPar_M = "S"
'    End If
'
'    If (vgBuscarMortalPar_M = "S") Then
'        For h = 1 To 2  '1=Causante '2=Beneficiario
'            vgs_Sexo = ""
'            vgs_Tipo = ""
'            vgs_Nro = ""
'            'For i = 1 To 2
'            '    If i = 1 Then vgs_Sexo = "M"    ' hombre  (M)asculino
'            '    If i = 2 Then vgs_Sexo = "F"    ' mujer   (F)emenino
'                i = 1
'                vgs_Sexo = "M"
'
'                'For j = 1 To 3
'                j = 3
'                    'If j = 1 Then vgs_Tipo = "MIT"    '1  = Invalido Total
'                    'If j = 2 Then vgs_Tipo = "RV"    '2  = No Invalido
'                    'If h = 2 And j = 2 Then vgs_Tipo = "B" ' Tabla de beneficiarios no invalidos
'                    If j = 3 Then vgs_Tipo = "MIP"    '3  = invalido Parcial
'
'                vgSw = False
'
'                'Buscar Número Correlativo desde la Estructura
'                For vgI = 1 To vgNumeroTotalTablas
'
'                    'Sql = " SELECT * "
'                    'Sql = Sql & " from PR_TVAL_MORTAL "
'                    'Sql = Sql & " where cod_sexo = '" & vgs_Sexo & "' and "
'                    'Sql = Sql & " cod_tiptabmor ='" & vgs_Tipo & "' and "
'                    'Sql = Sql & " fec_ini <= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "# AND "
'                    'Sql = Sql & " fec_ter >= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "#"
'                    'Set tb = vgConexionBD.Execute(Sql)
'                    'If Not (tb.EOF) Then
'
'                    If (egTablaMortal(vgI).Sexo = vgs_Sexo) And _
'                    (egTablaMortal(vgI).TipoTabla = vgs_Tipo) And _
'                    (egTablaMortal(vgI).Nombre = vgPalabra_MortalPar_M) And _
'                    (egTablaMortal(vgI).FechaIni <= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) And _
'                    (egTablaMortal(vgI).FechaFin >= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) Then
'
'                        vgs_Error = 0
'                        vgs_Nro = egTablaMortal(vgI).Correlativo
'                        vgFechaIniMortalPar_M = egTablaMortal(vgI).FechaIni
'                        vgFechaFinMortalPar_M = egTablaMortal(vgI).FechaFin
'
'                        'Limpiar columna de Datos
'                        If (h = 1) Then
'                            For vgX = 1 To Fintab
'                                Lx(i, j, vgX) = 0
'                            Next vgX
'                        Else
'                            For vgX = 1 To Fintab
'                                Ly(i, j, vgX) = 0
'                            Next vgX
'                        End If
'
'                        'vgs_Nro = tb!num_correlativo
'                        'vgFechaInicioMortal = Format(tb!fec_ini, "yyyymmdd")
'                        'vgFechaFinMortal = Format(tb!fec_ter, "yyyymmdd")
'                        ''vgFechaInicioMortal = DateSerial(Year(tb!fec_ini), Month(tb!fec_ini), Day(tb!fec_ini))
'                        ''vgFechaFinMortal = DateSerial(Year(tb!fec_ter), Month(tb!fec_ter), Day(tb!fec_ter))
'                        If vgs_Nro <> 0 Then
'                            Sql = "Select num_edad AS edad,mto_lx from MA_TVAL_MORDET "
'                            Sql = Sql & " where num_correlativo = " & vgs_Nro
'                            Sql = Sql & " order by num_edad "
'                            Set Tb2 = vgConexionBD.Execute(Sql)
'                            If Not (Tb2.EOF) Then
'                                vgSw = True
'                                'tb2.MoveFirst
'                                'k = 1
'                                Do While Not Tb2.EOF
'                                    k = Tb2!Edad
'                                    If h = 1 Then   'Causante
'                                        Lx(i, j, k) = Tb2!mto_lx
'                                    Else    'Beneficiario
'                                        Ly(i, j, k) = Tb2!mto_lx
'                                    End If
'                                    'k = k + 1
'                                    Tb2.MoveNext
'                                Loop
'                            Else
'                                vgError = 1067
'                                Exit Function
'                            End If
'                            Tb2.Close
'                        Else
'                            vgError = 1067
'                            Exit Function
'                        End If
'                    'Else
'                    '    vgError = 1067
'                    '    Exit Function
'                    End If
'                    'tb.Close
'                Next vgI
'                If (vgSw = False) Then
'                    vgError = 1067
'                    Exit Function
'                End If
'                'Next j
'            'Next i
'        Next h
'    End If
'
'    '8. Leer Tabla de Mortalidad de Beneficiarios Hombre
'    vgBuscarMortalBen_M = ""
'    If (vgFechaIniMortalBen_M <> "") And (vgFechaFinMortalBen_M <> "") Then
'        If ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) >= (vgFechaIniMortalBen_M)) And _
'        ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) <= (vgFechaFinMortalBen_M)) Then
'            vgBuscarMortalBen_M = "N"
'        Else
'            vgBuscarMortalBen_M = "S"
'        End If
'    Else
'        vgBuscarMortalBen_M = "S"
'    End If
'
'    If (vgBuscarMortalBen_M = "S") Then
'        'For h = 1 To 2  '1=Causante '2=Beneficiario
'        h = 2
'            vgs_Sexo = ""
'            vgs_Tipo = ""
'            vgs_Nro = ""
'            'For i = 1 To 2
'            '    If i = 1 Then vgs_Sexo = "M"    ' hombre  (M)asculino
'            '    If i = 2 Then vgs_Sexo = "F"    ' mujer   (F)emenino
'                i = 1
'                vgs_Sexo = "M"
'
'                'For j = 1 To 3
'                j = 2
'                    'If j = 1 Then vgs_Tipo = "MIT"    '1  = Invalido Total
'                    'If j = 2 Then vgs_Tipo = "RV"    '2  = No Invalido
'                    If h = 2 And j = 2 Then vgs_Tipo = "B" ' Tabla de beneficiarios no invalidos
'                    'If j = 3 Then vgs_Tipo = "MIP"    '3  = invalido Parcial
'
'                vgSw = False
'
'                'Buscar Número Correlativo desde la Estructura
'                For vgI = 1 To vgNumeroTotalTablas
'
'                    'Sql = " SELECT * "
'                    'Sql = Sql & " from PR_TVAL_MORTAL "
'                    'Sql = Sql & " where cod_sexo = '" & vgs_Sexo & "' and "
'                    'Sql = Sql & " cod_tiptabmor ='" & vgs_Tipo & "' and "
'                    'Sql = Sql & " fec_ini <= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "# AND "
'                    'Sql = Sql & " fec_ter >= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "#"
'                    'Set tb = vgConexionBD.Execute(Sql)
'                    'If Not (tb.EOF) Then
'
'                    If (egTablaMortal(vgI).Sexo = vgs_Sexo) And _
'                    (egTablaMortal(vgI).TipoTabla = vgs_Tipo) And _
'                    (egTablaMortal(vgI).Nombre = vgPalabra_MortalBen_M) And _
'                    (egTablaMortal(vgI).FechaIni <= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) And _
'                    (egTablaMortal(vgI).FechaFin >= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) Then
'
'                        vgs_Error = 0
'                        vgs_Nro = egTablaMortal(vgI).Correlativo
'                        vgFechaIniMortalBen_M = egTablaMortal(vgI).FechaIni
'                        vgFechaFinMortalBen_M = egTablaMortal(vgI).FechaFin
'
'                        'Limpiar columna de Datos
'                        For vgX = 1 To Fintab
'                            Ly(i, j, vgX) = 0
'                        Next vgX
'
'                        'vgs_Nro = tb!num_correlativo
'                        'vgFechaInicioMortal = Format(tb!fec_ini, "yyyymmdd")
'                        'vgFechaFinMortal = Format(tb!fec_ter, "yyyymmdd")
'                        ''vgFechaInicioMortal = DateSerial(Year(tb!fec_ini), Month(tb!fec_ini), Day(tb!fec_ini))
'                        ''vgFechaFinMortal = DateSerial(Year(tb!fec_ter), Month(tb!fec_ter), Day(tb!fec_ter))
'                        If vgs_Nro <> 0 Then
'                            Sql = "Select num_edad AS edad,mto_lx from MA_TVAL_MORDET "
'                            Sql = Sql & " where num_correlativo = " & vgs_Nro
'                            Sql = Sql & " order by num_edad "
'                            Set Tb2 = vgConexionBD.Execute(Sql)
'                            If Not (Tb2.EOF) Then
'                                vgSw = True
'                                'tb2.MoveFirst
'                                'k = 1
'                                Do While Not Tb2.EOF
'                                    k = Tb2!Edad
'                                    'If h = 1 Then   'Causante
'                                    '    lx(i, j, k) = tb2!mto_lx
'                                    'Else    'Beneficiario
'                                        Ly(i, j, k) = Tb2!mto_lx
'                                    'End If
'                                    'k = k + 1
'                                    Tb2.MoveNext
'                                Loop
'                            Else
'                                vgError = 1068
'                                Exit Function
'                            End If
'                            Tb2.Close
'                        Else
'                            vgError = 1068
'                            Exit Function
'                        End If
'                    'Else
'                    '    vgError = 1068
'                    '    Exit Function
'                    End If
'                    'tb.Close
'                Next vgI
'                If (vgSw = False) Then
'                    vgError = 1068
'                    Exit Function
'                End If
'                'Next j
'            'Next i
'        'Next h
'    End If
'
'
'
''    '-------------------------------------------------
''    'Leer Tabla de Mortalidad
''    '-------------------------------------------------
''    vgBuscarMortal = ""
''    If (vgFechaInicioMortal <> "") And (vgFechaFinMortal <> "") Then
''        If ((Format(Navig, "0000") & Format(Nmvig, "00") & Format(Ndvig, "00")) >= (vgFechaInicioMortal)) And _
''        ((Format(Navig, "0000") & Format(Nmvig, "00") & Format(Ndvig, "00")) <= (vgFechaFinMortal)) Then
''            vgBuscarMortal = "N"
''        Else
''            vgBuscarMortal = "S"
''        End If
''    Else
''        vgBuscarMortal = "S"
''    End If
''
''    If (vgBuscarMortal = "S") Then
''        For h = 1 To 2  '1=Causante '2=Beneficiario
''            vgs_Sexo = ""
''            vgs_Tipo = ""
''            vgs_Nro = ""
''            For i = 1 To 2
''                If i = 1 Then vgs_Sexo = "M"    ' hombre  (M)asculino
''                If i = 2 Then vgs_Sexo = "F"    ' mujer   (F)emenino
''                For j = 1 To 3
''                    If j = 1 Then vgs_Tipo = "MIT"    '1  = Invalido Total
''                    If j = 2 Then vgs_Tipo = "RV"    '2  = No Invalido
''                    If h = 2 And j = 2 Then vgs_Tipo = "B" ' Tabla de beneficiarios no invalidos
''                    If j = 3 Then vgs_Tipo = "MIP"    '3  = invalido Parcial
''                    Sql = " SELECT * "
''                    Sql = Sql & " from PR_TVAL_MORTAL "
''                    Sql = Sql & " where cod_sexo = '" & vgs_Sexo & "' and "
''                    Sql = Sql & " cod_tiptabmor ='" & vgs_Tipo & "' and "
''                    Sql = Sql & " fec_ini <= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "# AND "
''                    Sql = Sql & " fec_ter >= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "#"
''                    Set tb = vgConexionBD.Execute(Sql)
''                    If Not (tb.EOF) Then
''                        vgs_error = 0
''                        vgs_Nro = tb!num_correlativo
''                        vgFechaInicioMortal = Format(tb!fec_ini, "yyyymmdd")
''                        vgFechaFinMortal = Format(tb!fec_ter, "yyyymmdd")
''                        'vgFechaInicioMortal = DateSerial(Year(tb!fec_ini), Month(tb!fec_ini), Day(tb!fec_ini))
''                        'vgFechaFinMortal = DateSerial(Year(tb!fec_ter), Month(tb!fec_ter), Day(tb!fec_ter))
''                        If vgs_Nro <> 0 Then
''                            Sql = "Select num_edad AS edad,mto_lx from PR_TVAL_MORDET "
''                            Sql = Sql & " where num_correlativo = " & vgs_Nro
''                            Sql = Sql & " order by num_edad "
''                            Set tb2 = vgConexionBD.Execute(Sql)
''                            If Not (tb2.EOF) Then
''                                'tb2.MoveFirst
''                                k = 1
''                                Do While Not tb2.EOF
''                                    If h = 1 Then   'Causante
''                                        lx(i, j, k) = tb2!mto_lx
''                                    Else    'Beneficiario
''                                        ly(i, j, k) = tb2!mto_lx
''                                    End If
''                                    k = k + 1
''                                    tb2.MoveNext
''                                Loop
''                            End If
''                            tb2.Close
''                        Else
''                            vgError = 1001
''                            Exit Function
''                        End If
''                    Else
''                        vgError = 1001
''                        Exit Function
''                    End If
''                    tb.Close
''                Next j
''            Next i
''        Next h
''    End If
'
'    fgBuscarMortalidad = True
End Function

Function fgComboMortalNombre(iFecha, iTipoTabla, iPeriodo, iSexo) As String
Dim iFechaCot As Long
On Error GoTo Err_Combo

    'vlCombo.Clear
    fgComboMortalNombre = ""
    
    If (iFecha <> "") And (iPeriodo <> "") And (iSexo <> "") And (iTipoTabla <> "") Then
        
        vgI = vgNumeroTotalTablas
        vgX = 1
        vgJ = 1
        iFechaCot = Format(iFecha, "yyyymmdd")
        
        Do While vgX <= vgI
        
            If (egTablaMortal(vgX).FechaIni <= iFechaCot) And _
               (egTablaMortal(vgX).FechaFin >= iFechaCot) And _
               (egTablaMortal(vgX).Sexo = iSexo) And _
               (egTablaMortal(vgX).TipoTabla = iTipoTabla) And _
               (egTablaMortal(vgX).TipoPeriodo = iPeriodo) _
               And (egTablaMortal(vgX).Estado = "A") Then
                
                If (egTablaMortal(vgX).Oficial = "S") Then
                    fgComboMortalNombre = egTablaMortal(vgX).Nombre
                    Exit Do
                '    vlCombo.AddItem egTablaMortal(vgX).Nombre, 0
                '    vlCombo.ItemData(0) = egTablaMortal(vgX).Correlativo
                'Else
                '    vlCombo.AddItem egTablaMortal(vgX).Nombre
                '    vgJ = vlCombo.ListCount - 1
                '    vlCombo.ItemData(vgJ) = egTablaMortal(vgX).Correlativo
                End If
                'vgJ = vgJ + 1
                'CInt(egTablaMortal(vgX).Correlativo)
                'vlCombo.List = egTablaMortal(vgX).Correlativo
            End If
            
            vgX = vgX + 1
        Loop
        
        'If (vlCombo.ListCount <> 0) Then
        '    vlCombo.ListIndex = 0
        'End If
    End If
    
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

Function flCargaEstructuraPoliza(iNombreTabla As String, iPoliza As String, iEndoso As Integer, istPoliza As TyPoliza)
On Error GoTo Err_flCargaEstructuraPoliza

    vgSql = ""
    vgSql = "SELECT num_poliza,num_endoso,cod_tippension,cod_estado, "
    vgSql = vgSql & "cod_tipren,cod_modalidad,num_cargas,fec_vigencia, "
    vgSql = vgSql & "fec_tervigencia,mto_prima,mto_pension,num_mesdif, "
    vgSql = vgSql & "num_mesgar,prc_tasace,prc_tasavta,prc_tasaintpergar "
    'vgSql = vgSql & ",fec_emision,mto_pensiongar,cod_cuspp,"
    vgSql = vgSql & ",fec_dev,fec_inipencia,"
    vgSql = vgSql & "cod_moneda,mto_valmoneda,"
    vgSql = vgSql & "ind_cob,cod_cobercon,mto_facpenella,prc_facpenella, "
    vgSql = vgSql & "cod_dercre,cod_dergra "
    vgSql = vgSql & ", cod_tipreajuste, mto_valreajustetri, mto_valreajustemen, '0' as mto_pensiongar, fec_finpergar " 'hqr 13/01/2011
    vgSql = vgSql & "FROM " & iNombreTabla & " WHERE "
    vgSql = vgSql & "num_poliza = '" & Trim(iPoliza) & "' AND "
    vgSql = vgSql & "num_endoso = " & iEndoso & " "
    vgSql = vgSql & " ORDER BY num_endoso DESC"
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       With istPoliza
            .num_poliza = (vgRegistro!num_poliza)
            .num_endoso = (vgRegistro!num_endoso)
            .Cod_TipPension = (vgRegistro!Cod_TipPension)
            .Cod_Estado = (vgRegistro!Cod_Estado)
            .Cod_TipRen = (vgRegistro!Cod_TipRen)
            .Cod_Modalidad = (vgRegistro!Cod_Modalidad)
            .Num_Cargas = (vgRegistro!Num_Cargas)
            .Fec_Vigencia = (vgRegistro!Fec_Vigencia)
            .Fec_TerVigencia = (vgRegistro!Fec_TerVigencia)
            .Mto_Prima = (vgRegistro!Mto_Prima)
'            If (iNombreTabla = "PP_TMAE_ENDPOLIZA") Then
                .Mto_Pension = (vgRegistro!Mto_Pension)
'            Else
'                .Mto_Pension = iMtoPensionRef '*ABV 26/09/2007
'            End If
            .Num_MesDif = (vgRegistro!Num_MesDif)
            .Num_MesGar = (vgRegistro!Num_MesGar)
            .Prc_TasaCe = (vgRegistro!Prc_TasaCe)
            .Prc_TasaVta = (vgRegistro!Prc_TasaVta)
            If IsNull(vgRegistro!Prc_TasaIntPerGar) Then
               .Prc_TasaIntPerGar = ""
            Else
                .Prc_TasaIntPerGar = (vgRegistro!Prc_TasaIntPerGar)
            End If
'            .Fec_Emision = (vgRegistro!Fec_Emision)
            .Fec_Devengue = (vgRegistro!fec_dev)
            .Fec_IniPenCia = (vgRegistro!Fec_IniPenCia)
            .Cod_Moneda = (vgRegistro!Cod_Moneda)
            .Mto_ValMoneda = (vgRegistro!Mto_ValMoneda)
            .Mto_PensionGar = (vgRegistro!Mto_PensionGar)
            .Ind_Cob = (vgRegistro!Ind_Cob)
            .Cod_CoberCon = (vgRegistro!Cod_CoberCon)
            .Mto_FacPenElla = (vgRegistro!Mto_FacPenElla)
            .Prc_FacPenElla = (vgRegistro!Prc_FacPenElla)
            .Cod_DerCre = (vgRegistro!Cod_DerCre)
            .Cod_DerGra = (vgRegistro!Cod_DerGra)
            '.Cod_Cuspp = (vgRegistro!Cod_Cuspp)
            .Cod_TipReajuste = vgRegistro!Cod_TipReajuste 'hqr 13/11/2011
            .Mto_ValReajusteTri = IIf(IsNull(vgRegistro!Mto_ValReajusteTri), 0, vgRegistro!Mto_ValReajusteTri) 'hqr 13/11/2011
            .Mto_ValReajusteMen = IIf(IsNull(vgRegistro!Mto_ValReajusteMen), 0, vgRegistro!Mto_ValReajusteMen) 'hqr 26/02/2012
            .Fec_TerPerGar = IIf(IsNull(vgRegistro!fec_finpergar), "", vgRegistro!fec_finpergar)
       End With
    End If

Exit Function
Err_flCargaEstructuraPoliza:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function
'Implementacion GobiernoDeDatos(Se estructura y concatena la nueva direccion)_
Function flCargaEstructuraBeneficiariosDirec(iNombreTabla As String, iPoliza As String, iEndoso As Integer, stPolizaBenDirec() As TyBeneficiariosEst, vTabla As String, texto As String)
On Error GoTo Err_flCargaEstructuraBeneficiarios
Dim tablaDirec As String
Dim tablaTelef As String
Dim tablaBen As String

 vgSql = ""
    vgSql = "SELECT COUNT (num_orden) as numero "
    vgSql = vgSql & "FROM " & iNombreTabla & " WHERE "
    vgSql = vgSql & "num_poliza = '" & Trim(iPoliza) & "' AND "
    vgSql = vgSql & "num_endoso = " & iEndoso & " "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       vgNumBen = (vgRegistro!numero)
    End If
    
    ReDim stPolizaBenDirec(vgNumBen) As TyBeneficiariosEst
    
    If vTabla = "ENDBEN" Then
    tablaDirec = "PP_TMAE_ENDBEN_DIRECCION"
    tablaTelef = "PP_TMAE_ENDBEN_TELEFONO"
    Else
    tablaDirec = "PP_TMAE_BEN_DIRECCION"
    tablaTelef = "PP_TMAE_BEN_TELEFONO"
    End If


    vgSql = ""
    vgSql = "SELECT d.num_poliza,d.num_endoso,d.num_orden,d.cod_tipoidenben,d.num_idenben, "
    vgSql = vgSql & "d.cod_dire_via,d.gls_direccion,d.num_direccion,d.cod_blockchalet,d.gls_blockchalet, "
    vgSql = vgSql & "d.cod_interior,d.num_interior,d.cod_cjht,d.gls_nom_cjht,d.gls_etapa, "
    vgSql = vgSql & "d.gls_manzana,d.gls_lote,d.gls_referencia,d.cod_pais, "
    vgSql = vgSql & "d.cod_departamento,d.cod_provincia,d.cod_distrito,d.gls_desdirebusq,"
    vgSql = vgSql & "t.cod_tipo_fonoben,t.cod_area_fonoben,t.gls_fonoben,t.cod_tipo_telben2,cod_area_telben2,gls_telben2"
    vgSql = vgSql & " FROM " & tablaDirec & " D INNER JOIN " & tablaTelef & " T ON D.NUM_POLIZA = T.NUM_POLIZA AND D.NUM_ENDOSO = T.NUM_ENDOSO and d.num_orden = t.num_orden" & " WHERE "
    vgSql = vgSql & "d.num_poliza = '" & Trim(iPoliza) & "' AND "
    vgSql = vgSql & "d.num_endoso = " & iEndoso & " "
    vgSql = vgSql & " ORDER BY d.num_orden ASC"
    
    texto = vgSql
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       vgX = 0
       While Not vgRegistro.EOF
             vgX = vgX + 1
             With stPolizaBenDirec(vgX)
                 .num_poliza = (vgRegistro!num_poliza)
                .num_endoso = (vgRegistro!num_endoso)
                .Num_Orden = (vgRegistro!Num_Orden)
                .Cod_TipoIdenBen = (vgRegistro!Cod_TipoIdenBen)
                .Num_IdenBen = (vgRegistro!Num_IdenBen)
                .Cod_Direccion = IIf(IsNull(vgRegistro!cod_distrito), "", vgRegistro!cod_distrito)
                .cod_tip_fonoben = IIf(IsNull(vgRegistro!cod_tipo_fonoben), "", vgRegistro!cod_tipo_fonoben)
                .cod_area_fonoben = IIf(IsNull(vgRegistro!cod_area_fonoben), "", vgRegistro!cod_area_fonoben)
                .Gls_FonoBen = IIf(IsNull(vgRegistro!Gls_FonoBen), "", vgRegistro!Gls_FonoBen)
                .cod_tipo_telben2 = IIf(IsNull(vgRegistro!cod_tipo_telben2), "", vgRegistro!cod_tipo_telben2)
                .cod_area_telben2 = IIf(IsNull(vgRegistro!cod_area_telben2), "", vgRegistro!cod_area_telben2)
                .Gls_Telben2 = IIf(IsNull(vgRegistro!Gls_Telben2), "", vgRegistro!Gls_Telben2)
                .pTipoVia = IIf(IsNull(vgRegistro!cod_dire_via), "", vgRegistro!cod_dire_via)
                .pDireccion = IIf(IsNull(vgRegistro!Gls_Direccion), "", vgRegistro!Gls_Direccion)
                .pNumero = IIf(IsNull(vgRegistro!num_direccion), "", vgRegistro!num_direccion)
                .pTipoPref = IIf(IsNull(vgRegistro!COD_INTERIOR), "", vgRegistro!COD_INTERIOR)
                .pInterior = IIf(IsNull(vgRegistro!num_interior), "", vgRegistro!num_interior)
                .pManzana = IIf(IsNull(vgRegistro!gls_manzana), "", vgRegistro!gls_manzana)
                .pLote = IIf(IsNull(vgRegistro!gls_lote), "", vgRegistro!gls_lote)
                .pEtapa = IIf(IsNull(vgRegistro!gls_etapa), "", vgRegistro!gls_etapa)
                .pTipoConj = IIf(IsNull(vgRegistro!cod_cjht), "", vgRegistro!cod_cjht)
                .pConjHabit = IIf(IsNull(vgRegistro!gls_nom_cjht), "", vgRegistro!gls_nom_cjht)
                .pTipoBlock = IIf(IsNull(vgRegistro!cod_blockchalet), "", vgRegistro!cod_blockchalet)
                .pNumBlock = IIf(IsNull(vgRegistro!gls_blockchalet), "", vgRegistro!gls_blockchalet)
                .pReferencia = IIf(IsNull(vgRegistro!gls_referencia), "", vgRegistro!gls_referencia)
                .pvalEndosoGS = "1"
             End With
             vgRegistro.MoveNext
       Wend
    End If
    
    If vTabla = "ENDBEN" Then
    tablaBen = "PP_TMAE_ENDBEN"
    Else
    tablaBen = "PP_TMAE_BEN"
    End If
    
    vgSql = ""
    vgSql = "SELECT num_poliza,num_endoso,num_orden"
    vgSql = vgSql & ",Gls_DirBen, Cod_Direccion, Gls_FonoBen, Gls_CorreoBen, Gls_telben2"
    vgSql = vgSql & " FROM " & tablaBen & " WHERE "
    vgSql = vgSql & "num_poliza = '" & Trim(iPoliza) & "' AND "
    vgSql = vgSql & "num_endoso = " & iEndoso & " "
    vgSql = vgSql & " ORDER BY num_orden ASC"
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       vgX = 0
       While Not vgRegistro.EOF
        vgX = vgX + 1
            With stPolizaBenDirec(vgX)
            .Gls_FonoBen = IIf(IsNull(vgRegistro!Gls_FonoBen), "", vgRegistro!Gls_FonoBen)
            .Gls_Telben2 = IIf(IsNull(vgRegistro!Gls_Telben2), "", vgRegistro!Gls_Telben2)
            .pGlsCorreo = IIf(IsNull(vgRegistro!Gls_CorreoBen), "", vgRegistro!Gls_CorreoBen)
            .Cod_Direccion = IIf(IsNull(vgRegistro!Cod_Direccion), "", vgRegistro!Cod_Direccion)
            .pConcatDirec = IIf(IsNull(vgRegistro!Gls_DirBen), "", vgRegistro!Gls_DirBen)
             End With
             vgRegistro.MoveNext
       Wend
     End If
     
Exit Function
Err_flCargaEstructuraBeneficiarios:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
    
End Function
'fin Implementacion GobiernoDeDatos()__

Function flCargaEstructuraBeneficiarios(iNombreTabla As String, iPoliza As String, iEndoso As Integer, istBeneficiarios() As TyBeneficiarios)
On Error GoTo Err_flCargaEstructuraBeneficiarios

    vgSql = ""
    vgSql = "SELECT COUNT (num_orden) as numero "
    vgSql = vgSql & "FROM " & iNombreTabla & " WHERE "
    vgSql = vgSql & "num_poliza = '" & Trim(iPoliza) & "' AND "
    vgSql = vgSql & "num_endoso = " & iEndoso & " "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       vgNumBen = (vgRegistro!numero)
    End If

    vgSql = ""
    vgSql = "SELECT fec_finpergar "
    vgSql = vgSql & "FROM PP_TMAE_POLIZA WHERE "
    vgSql = vgSql & "num_poliza = '" & Trim(iPoliza) & "' AND "
    vgSql = vgSql & "num_endoso = " & iEndoso & " "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       vlFecFinPerGar = ("" & vgRegistro!fec_finpergar)
    End If
    
    
    ReDim istBeneficiarios(vgNumBen) As TyBeneficiarios

    vgSql = ""
    vgSql = "SELECT num_poliza,num_endoso,num_orden,cod_tipoidenben,num_idenben, "
    vgSql = vgSql & "gls_nomben,gls_nomsegben,gls_patben,gls_matben,cod_grufam, "
    vgSql = vgSql & "cod_par,cod_sexo,cod_sitinv,cod_dercre,cod_derpen, "
    vgSql = vgSql & "cod_cauinv,fec_nacben,fec_nachm,fec_invben, "
    vgSql = vgSql & "mto_pension,prc_pension,fec_fallben,cod_estpension, "
    vgSql = vgSql & "cod_motreqpen,mto_pensiongar,cod_caususben,fec_susben, "
    vgSql = vgSql & "fec_inipagopen,fec_terpagopengar,fec_matrimonio "
    vgSql = vgSql & ",prc_pensionleg,prc_pensiongar "
    'RRR 22/04/2013
    'If iNombreTabla = "PP_TMAE_BEN" Then
        vgSql = vgSql & ", Gls_DirBen, Cod_Direccion, Gls_FonoBen, Gls_CorreoBen, Gls_telben2, cod_banco, cod_tipcta, cod_monbco, num_ctabco"
    'End If
    '''''''''''''''''''''
    'mvg 20170904
    vgSql = vgSql & ",ind_bolelec "
    vgSql = vgSql & ",NUM_CUENTA_CCI "
    vgSql = vgSql & ",Cod_ViaPago, CONS_TRAINFO, CONS_DATCOMER "
    vgSql = vgSql & ",COD_SUCURSAL "
    vgSql = vgSql & " FROM " & iNombreTabla & " WHERE "
    vgSql = vgSql & "num_poliza = '" & Trim(iPoliza) & "' AND "
    vgSql = vgSql & "num_endoso = " & iEndoso & " "
    vgSql = vgSql & " ORDER BY num_orden ASC"
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       vgX = 0
       While Not vgRegistro.EOF
             vgX = vgX + 1
             With istBeneficiarios(vgX)
                .num_poliza = (vgRegistro!num_poliza)
                .num_endoso = (vgRegistro!num_endoso)
                .Num_Orden = (vgRegistro!Num_Orden)
                .Cod_TipoIdenBen = (vgRegistro!Cod_TipoIdenBen)
                .Num_IdenBen = (vgRegistro!Num_IdenBen)
                .Gls_NomBen = (vgRegistro!Gls_NomBen)
                If IsNull(vgRegistro!Gls_NomSegBen) Then
                   .Gls_NomSegBen = ""
                Else
                    .Gls_NomSegBen = (vgRegistro!Gls_NomSegBen)
                End If
                .Gls_PatBen = (vgRegistro!Gls_PatBen)
                If IsNull(vgRegistro!Gls_MatBen) Then
                   .Gls_MatBen = ""
                Else
                    .Gls_MatBen = (vgRegistro!Gls_MatBen)
                End If
                .Cod_GruFam = (vgRegistro!Cod_GruFam)
                .Cod_Par = (vgRegistro!Cod_Par)
                .Cod_Sexo = (vgRegistro!Cod_Sexo)
                .Cod_SitInv = (vgRegistro!Cod_SitInv)
                .Cod_DerCre = (vgRegistro!Cod_DerCre)
                If IsNull(vgRegistro!Cod_EstPension) Then
                   .Cod_EstPension = ""
                Else
                    .Cod_EstPension = (vgRegistro!Cod_EstPension)
                End If
                If IsNull(vgRegistro!Cod_CauInv) Then
                   .Cod_CauInv = ""
                Else
                    .Cod_CauInv = (vgRegistro!Cod_CauInv)
                End If
                .Fec_NacBen = (vgRegistro!Fec_NacBen)
                If IsNull(vgRegistro!Fec_NacHM) Then
                   .Fec_NacHM = ""
                Else
                    .Fec_NacHM = (vgRegistro!Fec_NacHM)
                End If
                If IsNull(vgRegistro!Fec_InvBen) Then
                   .Fec_InvBen = ""
                Else
                    .Fec_InvBen = (vgRegistro!Fec_InvBen)
                End If
'                If (iNombreTabla = "PP_TMAE_ENDBEN") Then
                    .Mto_Pension = (vgRegistro!Mto_Pension)
'                Else
'                    .Mto_Pension = Format((vgRegistro!Prc_Pension / 100) * (iMtoPensionAct), "#0.00")
'                End If
                
                .Prc_Pension = (vgRegistro!Prc_Pension)
                If IsNull(vgRegistro!Fec_FallBen) Then
                   .Fec_FallBen = ""
                Else
                    .Fec_FallBen = (vgRegistro!Fec_FallBen)
                End If
                .Cod_DerPen = (vgRegistro!Cod_DerPen)
                If IsNull(vgRegistro!Cod_MotReqPen) Then
                   .Cod_MotReqPen = ""
                Else
                    .Cod_MotReqPen = (vgRegistro!Cod_MotReqPen)
                End If
                .Mto_PensionGar = (vgRegistro!Mto_PensionGar)
                '.Mto_PensionGar = Format((vgRegistro!Prc_PensionGar / 100) * (vlMtoPensionAct), "#0.00")
                
                If IsNull(vgRegistro!Cod_CauSusBen) Then
                   .Cod_CauSusBen = ""
                Else
                    .Cod_CauSusBen = (vgRegistro!Cod_CauSusBen)
                End If
                If IsNull(vgRegistro!Fec_SusBen) Then
                   .Fec_SusBen = ""
                Else
                    .Fec_SusBen = (vgRegistro!Fec_SusBen)
                End If
                .Fec_IniPagoPen = (vgRegistro!Fec_IniPagoPen)
                If IsNull(vlFecFinPerGar) Then
                   .Fec_TerPagoPenGar = ""
                Else
                    .Fec_TerPagoPenGar = (vlFecFinPerGar)
                End If
                If IsNull(vgRegistro!Fec_Matrimonio) Then
                   .Fec_Matrimonio = ""
                Else
                    .Fec_Matrimonio = (vgRegistro!Fec_Matrimonio)
                End If
                .Prc_PensionLeg = (vgRegistro!Prc_PensionLeg)
                .Prc_PensionGar = (vgRegistro!Prc_PensionGar)
                'RRR 22/04/2013
                'If iNombreTabla = "PP_TMAE_BEN" Then
                .Cod_Direccion = IIf(IsNull(vgRegistro!Cod_Direccion) = True, 0, vgRegistro!Cod_Direccion)
                .Gls_DirBen = IIf(IsNull(vgRegistro!Gls_DirBen) = True, "", vgRegistro!Gls_DirBen)
                .Gls_CorreoBen = IIf(IsNull(vgRegistro!Gls_CorreoBen), "", vgRegistro!Gls_CorreoBen)
                .Gls_FonoBen = IIf(IsNull(vgRegistro!Gls_FonoBen), "", vgRegistro!Gls_FonoBen)
                .Gls_Telben2 = IIf(IsNull(vgRegistro!Gls_Telben2), "", vgRegistro!Gls_Telben2)
                .cod_tipcta = IIf(IsNull(vgRegistro!cod_tipcta), "00", vgRegistro!cod_tipcta)
                .cod_monbco = IIf(IsNull(vgRegistro!cod_monbco), "NS", vgRegistro!cod_monbco)
                .num_ctabco = IIf(IsNull(vgRegistro!num_ctabco), "", vgRegistro!num_ctabco)
                .Cod_Banco = IIf(IsNull(vgRegistro!Cod_Banco), "00", vgRegistro!Cod_Banco)
               
                'mvg 20170904
                .ind_bolelec = IIf(IsNull(vgRegistro!ind_bolelec), "N", vgRegistro!ind_bolelec)
                'End If
                
                '.Cod_Direccion = (vgRegistro!Cod_Direccion)
                '.Gls_DirBen = (vgRegistro!Gls_DirBen)
                '.Gls_CorreoBen = IIf(IsNull(vgRegistro!Gls_CorreoBen), "", vgRegistro!Gls_CorreoBen)
                '.Gls_FonoBen = IIf(IsNull(vgRegistro!Gls_FonoBen), "", vgRegistro!Gls_FonoBen)
                'INICIO GCP - FRACTAL 01042019
                .NUM_CUENTA_CCI = IIf(IsNull(vgRegistro!NUM_CUENTA_CCI), "", vgRegistro!NUM_CUENTA_CCI)
                .Cod_ViaPago = IIf(IsNull(vgRegistro!Cod_ViaPago), "00", vgRegistro!Cod_ViaPago)
                .Cod_Sucursal = IIf(IsNull(vgRegistro!Cod_Sucursal), "00", vgRegistro!Cod_Sucursal)
                .CONS_TRAINFO = IIf(IsNull(vgRegistro!CONS_TRAINFO), "0", vgRegistro!CONS_TRAINFO)
                .CONS_DATCOMER = IIf(IsNull(vgRegistro!CONS_DATCOMER), "0", vgRegistro!CONS_DATCOMER)
                
                'FIN GCP - FRACTAL 01042019
          End With
             vgRegistro.MoveNext
       Wend
    End If

Exit Function
Err_flCargaEstructuraBeneficiarios:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flCargaGrillaBeneficiarios(iGrilla As MSFlexGrid, istBeneficiarios() As TyBeneficiarios)
On Error GoTo Err_flCargaGrillaBeneficiarios
Dim vgCodPar As String
Dim vlCargaFecNac   As String
Dim vlCargaFecNacHM As String
Dim vlCargaFecFall  As String
Dim vlCargaFecMat   As String
Dim vlCargaFecInv   As String
Dim vlCargaFecSus   As String
Dim vlCargaFecIniPagoPen    As String
Dim vlCargaFecTerpagoPenGar As String
Dim vlCargaNomTipoIden As String
Dim vlMtoPension As Double
Dim vlMtoPensionGar As Double
''RRR 22/04/2013
Dim vlCodDireccion As String
Dim vlGlsDirb As String
Dim vlGlsFonob As String
Dim vlGlsCorreo As String
Dim vlGlsTelben2 As String
Dim vlPrcPensionGar As Double


    vgX = 0
    
    'I--- ABV 25/04/2005 ---
    'vgNumBen = iGrilla.Rows
    'F--- ABV 25/04/2005 ---
    
    'Call flInicializaGrillaBenef(iGrilla)
    While vgX < vgNumBen
          vgX = vgX + 1
          With istBeneficiarios(vgX)
                
                'Formatear la Fecha de Nacimiento
                vlCargaFecNac = DateSerial(CInt(Mid(.Fec_NacBen, 1, 4)), CInt(Mid(.Fec_NacBen, 5, 2)), CInt(Mid(.Fec_NacBen, 7, 2)))
                'Formatear la Fecha de Nacimiento del Hijo Menor
                If (Trim(.Fec_NacHM) <> "") Then
                    vlCargaFecNacHM = DateSerial(CInt(Mid(.Fec_NacHM, 1, 4)), CInt(Mid(.Fec_NacHM, 5, 2)), CInt(Mid(.Fec_NacHM, 7, 2)))
                Else
                    vlCargaFecNacHM = ""
                End If
                'Formatear la Fecha de Invalidez
                If (Trim(.Fec_InvBen) <> "") Then
                    vlCargaFecInv = DateSerial(CInt(Mid(.Fec_InvBen, 1, 4)), CInt(Mid(.Fec_InvBen, 5, 2)), CInt(Mid(.Fec_InvBen, 7, 2)))
                Else
                    vlCargaFecInv = ""
                End If
                
                'Formatear la Fecha de Fallecimiento
                If (Trim(.Fec_FallBen) <> "") Then
                    vlCargaFecFall = DateSerial(CInt(Mid(.Fec_FallBen, 1, 4)), CInt(Mid(.Fec_FallBen, 5, 2)), CInt(Mid(.Fec_FallBen, 7, 2)))
                Else
                    vlCargaFecFall = ""
                End If
                
                'Formatear la Fecha de Suspención del Beneficiario
                If (Trim(.Fec_SusBen) <> "") Then
                    vlCargaFecSus = DateSerial(CInt(Mid(.Fec_SusBen, 1, 4)), CInt(Mid(.Fec_SusBen, 5, 2)), CInt(Mid(.Fec_SusBen, 7, 2)))
                Else
                    vlCargaFecSus = ""
                End If
                
                'Formatear la Fecha de Inicio de Pago de Pensiones
                If (Trim(.Fec_IniPagoPen) <> "") Then
                    vlCargaFecIniPagoPen = DateSerial(CInt(Mid(.Fec_IniPagoPen, 1, 4)), CInt(Mid(.Fec_IniPagoPen, 5, 2)), CInt(Mid(.Fec_IniPagoPen, 7, 2)))
                Else
                    vlCargaFecIniPagoPen = ""
                End If
                
                'Formatear la Fecha de Termino de Pago del Periodo Garantizado
                If (Trim(.Fec_TerPagoPenGar) <> "") Then
                    vlCargaFecTerpagoPenGar = DateSerial(CInt(Mid(.Fec_TerPagoPenGar, 1, 4)), CInt(Mid(.Fec_TerPagoPenGar, 5, 2)), CInt(Mid(.Fec_TerPagoPenGar, 7, 2)))
                Else
                    vlCargaFecTerpagoPenGar = ""
                End If
                
                'Formatear la Fecha de Termino de Pago del Periodo Garantizado
                If (Trim(.Fec_Matrimonio) <> "") Then
                    vlCargaFecMat = DateSerial(CInt(Mid(.Fec_Matrimonio, 1, 4)), CInt(Mid(.Fec_Matrimonio, 5, 2)), CInt(Mid(.Fec_Matrimonio, 7, 2)))
                Else
                    vlCargaFecMat = ""
                End If
                
                vlCargaNomTipoIden = fgBuscarNombreTipoIden(.Cod_TipoIdenBen, False)
                
               'vgCodPar = " " & Trim(.Cod_Par) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_Par, Trim(.Cod_Par)))
                vlMtoPension = .Mto_Pension 'iMtoPensionAct * (.Prc_Pension / 100)
                vlMtoPensionGar = .Mto_PensionGar 'iMtoPensionAct * (.Prc_PensionGar / 100)
                vlCodDireccion = .Cod_Direccion '--RRR22 / 4 / 2013
                vlGlsDirb = .Gls_DirBen '--RRR22 / 4 / 2013
                vlGlsFonob = IIf(IsNull(.Gls_FonoBen) = True, "", .Gls_FonoBen) '--RRR22 / 4 / 2013
                vlGlsCorreo = IIf(IsNull(.Gls_CorreoBen) = True, "", .Gls_CorreoBen) '--RRR22 / 4 / 2013
                vlGlsTelben2 = IIf(IsNull(.Gls_Telben2) = True, "", .Gls_Telben2)
                vlPrcPensionGar = .Prc_PensionGar

               iGrilla.AddItem (.Num_Orden) & vbTab _
               & (" " & (Trim(.Cod_TipoIdenBen) & " - " & vlCargaNomTipoIden) & vbTab _
               & (Trim(.Num_IdenBen))) & vbTab _
               & (Trim(.Gls_NomBen)) & vbTab & (Trim(.Gls_NomSegBen)) & vbTab & (Trim(.Gls_PatBen)) & vbTab & (Trim(.Gls_MatBen)) & vbTab _
               & (Trim(.Cod_Par)) & vbTab _
               & (Trim(.Cod_GruFam)) & vbTab _
               & (Trim(.Cod_Sexo)) & vbTab _
               & (Trim(.Cod_SitInv)) & vbTab _
               & (Trim(.Cod_DerPen)) & vbTab _
               & (Trim(.Cod_DerCre)) & vbTab _
               & (Trim(.num_poliza)) & vbTab & (Trim(.num_endoso)) & vbTab _
               & (Trim(.Cod_CauInv)) & vbTab _
               & (vlCargaFecNac) & vbTab & (vlCargaFecNacHM) & vbTab _
               & (vlCargaFecInv) & vbTab _
               & vlMtoPension & vbTab & (Trim(.Prc_Pension)) & vbTab _
               & (vlCargaFecFall) & vbTab _
               & (Trim(.Cod_EstPension)) & vbTab _
               & (Trim(.Cod_MotReqPen)) & vbTab _
               & (vlMtoPensionGar) & vbTab _
               & (Trim(.Cod_CauSusBen)) & vbTab & (vlCargaFecSus) & vbTab _
               & (vlCargaFecIniPagoPen) & vbTab _
               & (vlCargaFecTerpagoPenGar) & vbTab _
               & (vlCargaFecMat) & vbTab _
               & (vlPrcPensionGar) & vbTab & (.Prc_PensionLeg) & vbTab _
               & (.Cod_Direccion) & vbTab & (.Gls_DirBen) & vbTab & (.Gls_FonoBen) & vbTab & (.Gls_CorreoBen) & vbTab & (.Gls_Telben2) & vbTab & (.Cod_Banco) & vbTab & (.cod_tipcta) & vbTab & (.cod_monbco) & vbTab & (.num_ctabco) & vbTab & (.ind_bolelec) & vbTab & (.NUM_CUENTA_CCI) & vbTab & (.Cod_ViaPago) & vbTab & (.Cod_Sucursal) & vbTab & (.CONS_TRAINFO) & vbTab & (.CONS_DATCOMER)  'mvg 20170904
    
          End With
    Wend

Exit Function
Err_flCargaGrillaBeneficiarios:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fgCargaEstBenGrilla(iGrilla As MSFlexGrid, istBeneficiarios() As TyBeneficiarios)
On Error GoTo Err_fgCargaEstBenGrilla
Dim vlPos, vlNumero As Integer
    
    If iGrilla.rows > 1 Then
    vlPos = 1
    iGrilla.Col = 0
    vgX = 0
    vgNumBen = (iGrilla.rows - 1)
    ReDim istBeneficiarios(vgNumBen) As TyBeneficiarios
    While vlPos <= (iGrilla.rows - 1)
            iGrilla.row = vlPos
            iGrilla.Col = 0
    
            vgX = vgX + 1
            With istBeneficiarios(vgX)
                 iGrilla.Col = 13
                 .num_poliza = (iGrilla.Text)
                 iGrilla.Col = 14
                 .num_endoso = (iGrilla.Text)
                 iGrilla.Col = 0
                 .Num_Orden = (iGrilla.Text)
                 iGrilla.Col = 1
                 vlCodTipoIdenBen = fgObtenerCodigo_TextoCompuesto(iGrilla.Text)
                 .Cod_TipoIdenBen = vlCodTipoIdenBen
                 iGrilla.Col = 2
                 .Num_IdenBen = (iGrilla.Text)
                 iGrilla.Col = 3
                 .Gls_NomBen = (iGrilla.Text)
                 iGrilla.Col = 4
                 .Gls_NomSegBen = (iGrilla.Text)
                 iGrilla.Col = 5
                 .Gls_PatBen = (iGrilla.Text)
                 iGrilla.Col = 6
                 .Gls_MatBen = (iGrilla.Text)
                 iGrilla.Col = 7
                 .Cod_Par = (iGrilla.Text)
                 iGrilla.Col = 8
                 .Cod_GruFam = (iGrilla.Text)
                 iGrilla.Col = 9
                 .Cod_Sexo = (iGrilla.Text)
                 iGrilla.Col = 10
                 .Cod_SitInv = (iGrilla.Text)
                 'iGrilla.Col = 11
                 '.Cod_EstPension = (iGrilla.Text)
                 iGrilla.Col = 12
                 .Cod_DerCre = (iGrilla.Text)
                 iGrilla.Col = 15
                 .Cod_CauInv = (iGrilla.Text)
                 iGrilla.Col = 16
                 '.Fec_NacBen = (iGrilla.Text)
                 .Fec_NacBen = Format(iGrilla.Text, "yyyymmdd")
                 iGrilla.Col = 17
                 '.Fec_NacHM = (iGrilla.Text)
                 .Fec_NacHM = Format(iGrilla.Text, "yyyymmdd")
                 iGrilla.Col = 18
                 '.Fec_InvBen = (iGrilla.Text)
                 .Fec_InvBen = Format(iGrilla.Text, "yyyymmdd")
                 iGrilla.Col = 19
                 .Mto_Pension = (iGrilla.Text)
                 iGrilla.Col = 20
                 .Prc_Pension = (iGrilla.Text)
                 iGrilla.Col = 21
                 '.Fec_FallBen = (iGrilla.Text)
                 .Fec_FallBen = Format(iGrilla.Text, "yyyymmdd")
                 'RRR
                 iGrilla.Col = 11
                 .Cod_DerPen = (iGrilla.Text)
                 
                 iGrilla.Col = 22
                 .Cod_EstPension = (iGrilla.Text)
                 'RRRR
                 iGrilla.Col = 23
                 .Cod_MotReqPen = (iGrilla.Text)
                 iGrilla.Col = 24
                 .Mto_PensionGar = (iGrilla.Text)
                 iGrilla.Col = 25
                 .Cod_CauSusBen = (iGrilla.Text)
                 iGrilla.Col = 26
                 '.Fec_SusBen = (iGrilla.Text)
                 .Fec_SusBen = Format(iGrilla.Text, "yyyymmdd")
                 iGrilla.Col = 27
                 '.Fec_IniPagoPen = (iGrilla.Text)
                 .Fec_IniPagoPen = Format(iGrilla.Text, "yyyymmdd")
                 iGrilla.Col = 28
                 '.Fec_TerPagoPenGar = (iGrilla.Text)
                 .Fec_TerPagoPenGar = Format(iGrilla.Text, "yyyymmdd")
                 iGrilla.Col = 29
                 '.Fec_Matrimonio = (iGrilla.Text)
                 .Fec_Matrimonio = Format(iGrilla.Text, "yyyymmdd")
                 iGrilla.Col = 30
                 .Prc_PensionGar = (iGrilla.Text)
                 iGrilla.Col = 31
                 .Prc_PensionLeg = (iGrilla.Text)
                 'RRR 21/05/2013
                 iGrilla.Col = 32
                 .Cod_Direccion = (iGrilla.Text)
                 iGrilla.Col = 33
                 .Gls_DirBen = (iGrilla.Text)
                 iGrilla.Col = 34
                 .Gls_FonoBen = (iGrilla.Text)
                 iGrilla.Col = 35
                 .Gls_CorreoBen = (iGrilla.Text)
                 iGrilla.Col = 36
                 .Gls_Telben2 = (iGrilla.Text)
                 'RRR 30/09/2019
                 iGrilla.Col = 37
                 .Cod_Banco = (iGrilla.Text)
                 iGrilla.Col = 38
                 .cod_tipcta = (iGrilla.Text)
                 iGrilla.Col = 39
                 .cod_monbco = (iGrilla.Text)
                 iGrilla.Col = 40
                 .num_ctabco = (iGrilla.Text)
                 iGrilla.Col = 41
                 .ind_bolelec = (iGrilla.Text)
                 iGrilla.Col = 42
                 .NUM_CUENTA_CCI = (iGrilla.Text)
                 iGrilla.Col = 43
                 .Cod_ViaPago = (iGrilla.Text)
                 iGrilla.Col = 44
                 .Cod_Sucursal = (iGrilla.Text)
                 iGrilla.Col = 45
                 .CONS_TRAINFO = (iGrilla.Text)
                 iGrilla.Col = 46
                 .CONS_DATCOMER = (iGrilla.Text)
            End With
            
            vlPos = vlPos + 1
       Wend
    End If

Exit Function
Err_fgCargaEstBenGrilla:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fgCalcularPorcentajeBenef(iFechaIniVig As String, iNumBenef As Integer, _
ostBeneficiarios() As TyBeneficiarios, Optional iTipoPension As String, _
Optional iPensionRef As Double, Optional iCalcularPension As Boolean, _
Optional iDerCrecerCotizacion As String, Optional iCobCobertura As String, _
Optional iCalcularPorcentaje As Boolean, Optional iPerGar As Long, _
Optional iPensionRefGar As Double) As Boolean
'Función: Permite actualizar los Porcentajes de Pensión de los Beneficiarios,
'         su Derecho a Acrecer y la Fecha de Nacimiento del Hijo Menor
'Parámetros de Entrada/Salida:
'iFechaIniVig     => Fecha de Inicio de Vigencia de la Póliza
'iNumBenef        => Número de Beneficiarios
'ostBeneficiarios => Estructura desde la cual se obtienen los datos de los
'                    Beneficiarios y al mismo tiempo se calcula el Porcentaje
'                    de Pensión al cual tienen Dº
'iCalcularPension => Permite indicar si se debe realizar el cálculo del Monto de la Pensión que le corresponde a cada Beneficiario
'iTipoPension     => Tipo de Pensión de la Póliza
'iPensionRef      => Monto de la Pensión de Referencia utilizada para el Calculo de la Pensión si el campo anterior esta en Verdadero
'iDerCrecerCotizacion => Indicador de Derecho a Crecer definido en la Cotización (S o N)
'iCobCobertura    => Indicador de Cobertura de la Cotización (S o N)

Dim vlValor  As Double
Dim vlSumPrcPns  As Double 'RRR 03/05/2013
Dim Numor()  As Integer
Dim Ncorbe() As Integer
Dim Codrel() As Integer
Dim Cod_Grfam() As Integer
Dim Sexobe() As String, Inv()   As String, Coinbe()   As String
Dim derpen() As Integer
Dim Nanbe()  As Integer, Nmnbe() As Integer, Ndnbe() As Integer
Dim Porben() As Double
Dim PorbenGar() As Double
Dim Codcbe() As String
Dim Iaap     As Integer, Immp   As Integer, Iddp    As Integer
Dim vlNum_Ben As Integer
Dim Hijos()  As Integer, Hijos_Inv() As Integer
Dim Hijos_SinDerechoPension()  As Integer 'hqr 08/07/2006
Dim Hijo_Menor() As Date
Dim Hijo_Menor_Ant() As Date
Dim Fec_NacHM() As Date
Dim estpen() As Integer

'I--- ABV 18/07/2006 ---
Dim Hijos_SinConyugeMadre()  As Integer
Dim Fec_Fall() As String
Dim vlFechaFallCau As String
Dim cont_esp_Totales As Integer
Dim cont_esp_Tot_GF() As Integer
Dim cont_mhn_Totales As Integer
Dim cont_mhn_Tot_GF() As Integer
Dim vlValorHijo  As Double
'F--- ABV 18/07/2006 ---

Dim cont_mhn() As Integer
Dim cont_causante As Integer
Dim cont_esposa As Integer
Dim cont_mhn_tot As Integer
Dim cont_hijo As Integer
Dim cont_padres As Integer

Dim L24 As Long, i As Long, edad_mes_ben As Long
Dim fecha_sin As Long, vlContBen As Long
Dim sexo_cau As String
Dim g As Long, Q As Long, x As Long, j As Long, u As String, k As Long
Dim v_hijo As Double

Dim vlFechaFallecimiento As String
Dim vlFechaMatrimonio    As String

Dim vlPorcBenef As Double
Dim vlPorcGarBenef As Double
Dim vlPenBenef As Double
Dim vlPenGarBenef As Double

Dim vlSumarTotalPorcentajePension As Double, vlSumaDef As Double, vlDif As Double
Dim vlPorcentajeRecal As Double

Dim vlRemuneracion As Double, vlRemuneracionProm As Double
Dim vlRemuneracionBase As Double
Dim vlPrcCobertura As Double
Dim vlFecTerPerGar As String
Dim vlSumaPjePenPadres As Double
Dim vlEsEdadLeg As Integer
Dim vbEsEdadLeg As Boolean
Dim ctoActivos As Integer

Dim cont_fall() As Integer
'I--- ABV 21/06/2007 ---
'On Error GoTo Err_fgCalcularPorcentaje
'F--- ABV 21/06/2007 ---

'    Call flAsignaPorcentajesLegales
'    Call flValidaBeneficiarios
'    Call flDerechoAcrecer
'    Call flMadreHijoMenor
'    Call flVariosConyuges
'    Call flHijosSolos
    vbEsEdadLeg = False
    fgCalcularPorcentajeBenef = False
    L24 = 0
    'Debiera tomar la Fecha de Devengue
    
    If (fgCarga_Param("LI", "L24", iFechaIniVig) = True) Then
        L24 = vgValorParametro
    Else
        vgError = 1000
        MsgBox "No existe Edad de tope para los 24 años.", vbCritical, "Proceso Cancelado"
        Exit Function
    End If
    
    'Mensualizar la Edad de 24 Años
    L24 = L24 * 12
    'I--- ABV 21/07/2006 ---
    'Edad de 18 años
    Dim L18 As Long
    If (fgCarga_Param("LI", "L18", iFechaIniVig) = True) Then
        L18 = vgValorParametro
    Else
        vgError = 1000
        MsgBox "No existe Edad de tope para los 18 años.", vbCritical, "Proceso Cancelado"
        Exit Function
    End If
    L18 = L18 * 12
    'Mensualizar la Edad de 24 Años
    vlEsEdadLeg = flObtieneEdadNormativa(vgNum_pol)
    
    If vlEsEdadLeg > 216 Then
        vbEsEdadLeg = True
    End If


'F--- ABV 21/07/2006 ---

    'If Not IsDate(txt_fecha_devengo) Then
    '   X = MsgBox("Debe ingresar la Fecha de Inicio de la Pensión", 16)
    '   txt_fecha_devengo.SetFocus
    '   Exit Function
    'End If
    'If CDate(txt_fecha_devengo) > CDate(lbl_cotizacion) Then
    '    X = MsgBox("Error", "La fecha de devengamiento no puede ser mayor que la fecha de cotización", 16)
    '    Exit Function
    'End If
    'Iaap = CInt(Year(CDate(txt_fecha_devengo))) 'a Fecha de siniestro
    'Immp = CInt(Month(CDate(txt_fecha_devengo))) 'm Fecha de siniestro
    'Iddp = CInt(Day(CDate(txt_fecha_devengo))) 'd Fecha de siniestro
    
    Iaap = CInt(Mid(iFechaIniVig, 1, 4)) 'a Fecha de siniestro
    Immp = CInt(Mid(iFechaIniVig, 5, 2)) 'm Fecha de siniestro
    Iddp = CInt(Mid(iFechaIniVig, 7, 2)) 'd Fecha de siniestro
    
    fecha_sin = Iaap * 12 + Immp
    vlNum_Ben = iNumBenef 'grd_beneficiarios.Rows - 1 '.Rows - 1
    
    ReDim Numor(vlNum_Ben) As Integer
    ReDim Ncorbe(vlNum_Ben) As Integer
    ReDim Codrel(vlNum_Ben) As Integer
    ReDim Cod_Grfam(vlNum_Ben) As Integer
    ReDim Sexobe(vlNum_Ben) As String
    ReDim Inv(vlNum_Ben) As String
    ReDim Coinbe(vlNum_Ben) As String
    ReDim derpen(vlNum_Ben) As Integer
    ReDim Nanbe(vlNum_Ben) As Integer
    ReDim Nmnbe(vlNum_Ben) As Integer
    ReDim Ndnbe(vlNum_Ben) As Integer
    ReDim Hijos(vlNum_Ben) As Integer
    ReDim Hijos_Inv(vlNum_Ben) As Integer
    ReDim Hijos_SinDerechoPension(vlNum_Ben) As Integer
    ReDim Hijo_Menor(vlNum_Ben) As Date
    ReDim Hijo_Menor_Ant(vlNum_Ben) As Date
    ReDim Porben(vlNum_Ben) As Double
    ReDim Codcbe(vlNum_Ben) As String
    ReDim Fec_NacHM(vlNum_Ben) As Date
    ReDim cont_mhn(vlNum_Ben) As Integer
    ReDim PorbenGar(vlNum_Ben) As Double 'RRR 03/06/2013
    ReDim estpen(vlNum_Ben) As Integer
    
    
'I--- ABV 18/07/2006 ---
    ReDim Hijos_SinConyugeMadre(vlNum_Ben) As Integer
    ReDim Fec_Fall(vlNum_Ben) As String
    ReDim cont_mhn_Tot_GF(vlNum_Ben) As Integer
    ReDim cont_esp_Tot_GF(vlNum_Ben) As Integer
    ReDim cont_fall(vlNum_Ben) As Integer
    ReDim Der_pen(vlNum_Ben) As String
    
    Dim cont_hijosS As Integer
    Dim SitInv As String
    Dim vf As Boolean
    
    vf = False
    
    vlFechaFallCau = ""
    vlValorHijo = 0
    cont_hijosS = 0
'F--- ABV 18/07/2006 ---
    
    num_pol = vgNum_pol
    
    ''cuenta la cantidad de beneficiarios activos
    ctoActivos = 0
    i = 1
    Do While i <= vlNum_Ben
        derpen(i) = ostBeneficiarios(i).Cod_DerPen
        If derpen(i) = "99" Then
            ctoActivos = ctoActivos + 1
        End If
    i = i + 1
    Loop
    
      
If (iCalcularPorcentaje = True) Then
    vlContBen = 1 '0
    i = 1
    Do While i <= vlNum_Ben
    
        vlContBen = vlContBen + 1
        'Nº Orden
        
        'Msf_GrillaBenef.Row = i
        'Msf_GrillaBenef.Col = 0
        'Numor(i) = Msf_GrillaBenef.Text ''tb_ben!cod_numordben 'N° de orden  NUMOR(I)
        
        'If Trim(grd_beneficiarios.TextMatrix(i, 0)) = "" Then
        If Trim(ostBeneficiarios(i).Num_Orden) = "" Then
            vgError = 1000
            MsgBox "No existe Número de Orden de Beneficiario.", vbCritical, "Error de Datos"
            Exit Do
        End If
        
        'Número de Orden
        Numor(i) = ostBeneficiarios(i).Num_Orden
        
        'Parentesco
        Ncorbe(i) = ostBeneficiarios(i).Cod_Par ''tb_ben!cod_par
        Codrel(i) = ostBeneficiarios(i).Cod_Par ''tb_ben!cod_par
        
        'Grupo Familiar
        Cod_Grfam(i) = ostBeneficiarios(i).Cod_GruFam
        
        'Sexo y Situación de Invalidez
        Sexobe(i) = ostBeneficiarios(i).Cod_Sexo
        Inv(i) = ostBeneficiarios(i).Cod_SitInv
        
        If (Ncorbe(i) = "99") Then
            sexo_cau = Sexobe(i)
            SitInv = Inv(i)
        End If
        
        'If CInt(iTipoPension) >= 8 Then
            'Derecho Pensión
        Der_pen(i) = ostBeneficiarios(i).Cod_DerPen
        derpen(i) = ostBeneficiarios(i).Cod_DerPen
        
        If Ncorbe(i) <> "99" Then
            derpen(i) = "99"
        End If
            
        estpen(i) = ostBeneficiarios(i).Cod_EstPension
        'Else
        '    derpen(i) = ostBeneficiarios(i).Cod_EstPension
        '    derpen(i) = ostBeneficiarios(i).Cod_EstPension
        'End If
        
        
        'Fecha de Nacimiento
        Nanbe(i) = CInt(Mid(ostBeneficiarios(i).Fec_NacBen, 1, 4)) 'a Fecha de nacimiento
        Nmnbe(i) = CInt(Mid(ostBeneficiarios(i).Fec_NacBen, 5, 2)) 'm Fecha de nacimiento
        Ndnbe(i) = CInt(Mid(ostBeneficiarios(i).Fec_NacBen, 7, 2)) 'd Fecha de nacimiento
                    
        'Fecha nacimiento hijo menor =IJAM(I),IJMN(I),IJDN(I)
        
        'Codificación de Situación de Invalidez
        If Inv(i) = "P" Then Coinbe(i) = "P"
        If Inv(i) = "T" Then Coinbe(i) = "T"
        If Inv(i) = "N" Then Coinbe(i) = "N"
        
        '*********
        edad_mes_ben = fecha_sin - (Nanbe(i) * 12 + Nmnbe(i))
        vlFecTerPerGar = ostBeneficiarios(i).Fec_TerPagoPenGar
        vlFechaFallecimiento = ostBeneficiarios(i).Fec_FallBen
'I--- ABV 18/07/2006 ---
        Fec_Fall(i) = vlFechaFallecimiento
        If vlFechaFallecimiento <> "" Then
            Der_pen(i) = "10"
        End If
        If Codrel(i) = 99 Then
            vlFechaFallCau = vlFechaFallecimiento
        End If
'F--- ABV 18/07/2006 ---

        vlFechaMatrimonio = ostBeneficiarios(i).Fec_Matrimonio
    'RRR 11/06/2013
        'I--- ABV 16/04/2005 ---
        'If derpen(i) <> "10" Then
        
            If (Codrel(i) >= 30 And Codrel(i) < 40) Then
                If vbEsEdadLeg = True Then
                    L18 = IIf(flCompletaRequisitos(num_pol, CInt(i)) = True, L24, L18)
                End If
                If ctoActivos = 0 Then
                    If vgCausaEndoso = "05" Or vgCausaEndoso = "09" Then
                        derpen(i) = 10
                        PorbenGar(i) = 0
                    Else
                        If vlFecTerPerGar > iFechaIniVig Then
                            derpen(i) = 99
                            PorbenGar(i) = ostBeneficiarios(i).Prc_PensionGar
                            vf = True
                        End If
                    End If
                Else
                    
                End If
                If (vlFechaFallecimiento <> "") Or (vlFechaMatrimonio <> "") Then
                    'derpen(i) = 10
                     
                    If ostBeneficiarios(i).Cod_EstPension <> "10" Then
                        ostBeneficiarios(i).Cod_EstPension = "10"
                        Hijos_SinDerechoPension(ostBeneficiarios(i).Cod_GruFam) = Hijos_SinDerechoPension(ostBeneficiarios(i).Cod_GruFam) + 1
                    End If
                Else
                    If edad_mes_ben > vlEsEdadLeg And Coinbe(i) = "N" Then
                        derpen(i) = 10
                        'hqr 08/07/2006 cambiar Estado de pension a No Vigente.  Se debe cambiar estado de madre si es hijo unico
                        If ostBeneficiarios(i).Cod_EstPension <> "10" Then
                            ostBeneficiarios(i).Cod_EstPension = "10"
                            Hijos_SinDerechoPension(ostBeneficiarios(i).Cod_GruFam) = Hijos_SinDerechoPension(ostBeneficiarios(i).Cod_GruFam) + 1
                        Else
                            Hijos_SinDerechoPension(ostBeneficiarios(i).Cod_GruFam) = Hijos_SinDerechoPension(ostBeneficiarios(i).Cod_GruFam) + 1
                        End If
                    Else
                        If derpen(i) = 10 Then
                            ostBeneficiarios(i).Cod_EstPension = "10"
                            Hijos_SinDerechoPension(ostBeneficiarios(i).Cod_GruFam) = Hijos_SinDerechoPension(ostBeneficiarios(i).Cod_GruFam) + 1
                        Else
                            derpen(i) = 99
                            If edad_mes_ben > 216 Then
                                If flCompletaRequisitos(num_pol, CInt(i)) = False Then
                                    ostBeneficiarios(i).Cod_EstPension = "10"
                                End If
                            End If
                            
                        End If
                    End If
                End If
                cont_hijosS = cont_hijosS + 1
            Else
                'El resto de los Beneficiarios que no sean Hijos, solo se dejan como
                'Sin Derecho a Pensión cuando están fallecidos
                If (vlFechaFallecimiento <> "") Then
                    If (iTipoPension <> "08" And iTipoPension <> "09" And iTipoPension <> "10" And iTipoPension <> "11" And iTipoPension <> "12") Then
                        'derpen(i) = 10
                        If vlFecTerPerGar > iFechaIniVig Then
                            derpen(i) = 99
                            PorbenGar(i) = ostBeneficiarios(i).Prc_PensionGar
                        End If
                    Else
                        If Codrel(i) <> "99" Then
                            'derpen(i) = 10
                            If vlFechaFallecimiento <> "" Then
                            
                            Else
                                If vlFecTerPerGar > iFechaIniVig Then
                                    derpen(i) = 99
                                    PorbenGar(i) = ostBeneficiarios(i).Prc_PensionGar
                                    vf = True
                                End If
                            End If
                            
                        Else
                            derpen(i) = 10
                        End If
                    End If
                   
                   
'                    If vgCausaEndoso = "09" Then
'                        derpen(i) = 99
'                    End If
    'I--- ABV 18/07/2006 ---
    '                If (Codrel(i) = "11") Or (Codrel(i) = "21") Then
    '                    If ostBeneficiarios(i).Cod_EstPension <> "10" Then
    '                        ostBeneficiarios(i).Cod_EstPension = "10"
    '                        Hijos_SinConyugeMadre(ostBeneficiarios(i).Cod_GruFam) = Hijos_SinConyugeMadre(ostBeneficiarios(i).Cod_GruFam) + 1
    '                    End If
    '                End If
    'F--- ABV 18/07/2006 ---
                Else
                    
                    'derpen(i) = 99 'ostBeneficiarios(i).Cod_EstPension
'                    If CInt(iTipoPension) >= 8 Then
'                        derpen(i) = 99
'                        ostBeneficiarios(i).Cod_DerPen = derpen(i)
'                    Else
'                        ostBeneficiarios(i).Cod_DerPen = 99
'                        derpen(i) = ostBeneficiarios(i).Cod_EstPension
'                        'ostBeneficiarios(i).Cod_EstPension = derpen(i)
'                    End If
                    
                End If
            End If
        'End If
        ostBeneficiarios(i).Cod_DerPen = derpen(i)
        'ostBeneficiarios(i).Cod_EstPension = derpen(i)
        'F--- ABV 16/04/2005 ---
        
        i = i + 1
    Loop
    
    cont_causante = 0
    cont_esposa = 0
    'cont_mhn = 0
    cont_mhn_tot = 0
    cont_hijo = 0
    cont_padres = 0

'I--- ABV 18/07/2006 ---
    cont_esp_Totales = 0
    cont_mhn_Totales = 0
'F--- ABV 18/07/2006 ---
    
    'Primer Ciclo
    For g = 1 To vlNum_Ben
        If derpen(g) <> 10 Then
            '99: con derecho a pension,20: con Derecho Pendiente
            '10: sin derecho a pension
            If Ncorbe(g) = 99 Then
                cont_causante = cont_causante + 1
            Else
                Select Case Ncorbe(g)
                    Case 10, 11
                        cont_esposa = cont_esposa + 1
                    Case 20, 21
                        Q = Cod_Grfam(g)
                        cont_mhn(Q) = cont_mhn(Q) + 1
                        cont_mhn_tot = cont_mhn_tot + 1
                    Case 30
                        'edad = fgEdadBen(vg_fecsin, vgBen(g).fec_nacben)
                        Q = Cod_Grfam(g)
                        Hijos(Q) = Hijos(Q) + 1
                        If Coinbe(g) <> "N" Then Hijos_Inv(Q) = Hijos_Inv(Q) + 1
                        Hijo_Menor(Q) = DateSerial(Nanbe(g), Nmnbe(g), Ndnbe(g))
                        If Hijos(Q) > 1 Then
                            If Hijo_Menor(Q) > Hijo_Menor_Ant(Q) Then
                                Hijo_Menor_Ant(Q) = Hijo_Menor(Q)
                            End If
                        Else
                            Hijo_Menor_Ant(Q) = Hijo_Menor(Q)
                        End If
                        'Case 35
                        If vbEsEdadLeg = True Then
                             L18 = IIf(flCompletaRequisitos(num_pol, CInt(i)) = True, L24, L18)
                        End If

                        edad_mes_ben = fecha_sin - (Nanbe(g) * 12 + Nmnbe(g))
                        If Coinbe(g) = "N" And edad_mes_ben <= vlEsEdadLeg Then
                            cont_hijo = cont_hijo + 1
                        Else
                            If Coinbe(g) = "T" Or Coinbe(g) = "P" Then
                                cont_hijo = cont_hijo + 1
                            End If
                        End If
                    Case 41, 42
                        cont_padres = cont_padres + 1
                    Case Else
                        vgError = 1000
                        x = MsgBox("Error en codificación de codigo de relación", vbCritical)
                        Exit Function
                End Select
            End If
        Else
            'Verificar si la cónyuge o la Madre Falleció antes que el Causante para contarla
            Select Case Ncorbe(g)
                Case 11:
                    If (vlFechaFallCau > Fec_Fall(g)) Then
                        cont_esposa = cont_esposa + 1
                        Q = Cod_Grfam(g)
                        cont_esp_Tot_GF(Q) = cont_esp_Tot_GF(Q) + 1
                        cont_esp_Totales = cont_esp_Totales + 1
                    End If
                Case 21:
'                    If (vlFechaFallCau > Fec_Fall(g)) Then
'                        Q = Cod_Grfam(g)
'                        cont_mhn_Tot_GF(Q) = cont_mhn_Tot_GF(Q) + 1
'                        cont_mhn_Totales = cont_mhn_Totales + 1
''                    Else
''                        Q = Cod_Grfam(g)
''                        cont_mhn_Tot_GF(Q) = cont_mhn_Tot_GF(Q) - 1
''                        cont_mhn_Totales = cont_mhn_Totales - 1
'                    End If
            End Select
        End If
    Next g
                
''I--- ABV 18/07/2006 ---
''Corregir el Parentesco de los Hijos cuando la Cónyuge o la Madre haya Muerto antes que el Causante
'    For j = 1 To vlNum_Ben
'        '99: con derecho a pension,20: con Derecho Pendiente
'        '10: sin derecho a pension
'        If derpen(j) <> 10 Then
''            edad_mes_ben = fecha_sin - (Nanbe(j) * 12 + Nmnbe(j))
'            Select Case Ncorbe(j)
'
'                Case 30
'                    Q = Cod_Grfam(j)
'                    'Cuando No existe Conyuge y tampoco no Existe MHN para el Hijo a Analizar
'                    'se deben modificar los Códigos de los Hijos a 35
''I--- ABV 20/07/2006 ---
''                    If cont_esp_Tot_GF(Q) <= 0 And cont_mhn_Tot_GF(Q) <= 0 Then
'                    If cont_esposa <= 0 And cont_mhn(Q) <= 0 Then
''F--- ABV 20/07/2006 ---
'                        Ncorbe(j) = 35
'                        Hijos(Q) = Hijos(Q) - 1
'                        cont_hijo = cont_hijo + 1
'                        If ostBeneficiarios(j).Cod_Par <> "35" Then
'                            ostBeneficiarios(j).Cod_Par = "35"
'                            'Hijos_SinConyugeMadre(ostBeneficiarios(j).Cod_GruFam) = Hijos_SinConyugeMadre(ostBeneficiarios(j).Cod_GruFam) + 1
'                        End If
'                    End If
'                Case 35
'                    Q = Cod_Grfam(j)
'                    'Cuando existe Conyuge o MHN, considerándola como existente
'                    'cuando ésta muere antes del Causante de la Póliza,
'                    'se deben modificar los Códigos de los Hijos a 30
''I--- ABV 20/07/2006 ---
''                    If cont_esp_Tot_GF(Q) > 0 Or cont_mhn_Tot_GF(Q) > 0 Then
'                    If cont_esposa > 0 Or cont_mhn(Q) > 0 Then
''F--- ABV 20/07/2006 ---
'                        Ncorbe(j) = 30
'                        Hijos(Q) = Hijos(Q) + 1
'                        cont_hijo = cont_hijo - 1
'                        If ostBeneficiarios(j).Cod_Par <> "30" Then
'                            ostBeneficiarios(j).Cod_Par = "30"
'                            'Hijos_SinConyugeMadre(ostBeneficiarios(j).Cod_GruFam) = Hijos_SinConyugeMadre(ostBeneficiarios(j).Cod_GruFam) + 1
'                        End If
'                    End If
'            End Select
'        End If
'    Next j
''F--- ABV 18/07/2006 ---
    If ctoActivos = 0 Then
        GoTo NoValora
    End If
    j = 1
    For j = 1 To vlNum_Ben
        '99: con derecho a pension,20: con Derecho Pendiente
        '10: sin derecho a pension
        If derpen(j) <> 10 Then
            edad_mes_ben = fecha_sin - (Nanbe(j) * 12 + Nmnbe(j))
            Select Case Ncorbe(j)
                Case 99
                    If cont_causante > 1 Then
                        vgError = 1000
                        x = MsgBox("Error en codificación de codigo de relación, No puede ingresar otro causante", vbCritical)
                        Exit Function
                    End If
                    'I--- ABV 25/02/2005 ---
                    'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
                    If (fgObtenerPorcentaje(CStr(Ncorbe(j)), Coinbe(j), Sexobe(j), iFechaIniVig, SitInv) = True) Then
                        vlValor = vgValorPorcentaje
                    Else
                        vgError = 1000
                        MsgBox "No existe valor de Porcentaje para el parámetro de :" _
                        & Ncorbe(j) & " - " & Coinbe(j) & " - " & Sexobe(j) & " - " & iFechaIniVig & "."
                        Exit Function
                    End If
                    'F--- ABV 25/02/2005 ---
                    If (vlValor < 0) Then
                        vgError = 1000
                        MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
                        'error
                        Exit Function
                    Else
                        Porben(j) = vlValor
                        'Porben(j) = 100
                        Codcbe(j) = "N"
                    End If
                Case 10, 11
                    If sexo_cau = "M" Then
                        If Sexobe(j) <> "F" Then
                            vgError = 1000
                            x = MsgBox("Error de código de sexo, el Sexo de la Cónyuge debe ser Femenino.", vbCritical)
                            Exit Function
                        End If
                    Else
                        If Sexobe(j) <> "M" Then
                            vgError = 1000
                            x = MsgBox("Error de codigo de sexo, el Sexo del Cónyuge debe ser Masculino.", vbCritical)
                            Exit Function
                        End If
                    End If
                    'hqr 08/07/2006 validacion y cambio parentesco
                    u = Cod_Grfam(j)
                    If Hijos(u) = 0 And Ncorbe(j) = 11 Then
                        If iPerGar > 0 Then
                            'If flSiDebeCambiarParentesco(num_pol, iFechaIniVig) = 1 Then
                                If Hijos_SinDerechoPension(u) > cont_hijo And (vgCausaEndoso = "05" Or vgCausaEndoso = "09") Then
                                    Ncorbe(j) = 11
                                    ostBeneficiarios(j).Cod_Par = Ncorbe(j)
                                End If
                            'End If
                        Else
'                            If Hijos_SinDerechoPension(u) = 0 Then
'                                'vgError = 1000
'                                X = MsgBox("Error de código de relación, 'Cónyuge Con Hijos con Dº Pensión', no tiene Hijos.", vbCritical)
'                                Exit Function
'                            Else
'                                Ncorbe(j) = 10
'                                ostBeneficiarios(j).Cod_Par = 10
'                            End If
                        End If
                    End If
                    
                    'HQR 08/07/2006 se deja al final porque se cambia el tipo de parentesco
                    'I--- ABV 25/02/2005 ---
                     'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
                     If (fgObtenerPorcentaje(CStr(Ncorbe(j)), Coinbe(j), Sexobe(j), iFechaIniVig, SitInv) = True) Then
                         vlValor = vgValorPorcentaje
                     Else
                         vgError = 1000
                         MsgBox "No existe valor de Porcentaje para el parámetro de :" _
                         & Ncorbe(j) & " - " & Coinbe(j) & " - " & Sexobe(j) & " - " & iFechaIniVig & "."
                         Exit Function
                     End If
                     'F--- ABV 25/02/2005 ---
                     If (vlValor < 0) Then
                         vgError = 1000
                         MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
                         'error
                         Exit Function
                     End If

                    'fin hqr 08/07/2006
                    If sexo_cau = "M" Or sexo_cau = "F" Then

                        'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
                        'If (vlValor < 0) Then
                        '    Exit Sub
                        'Else
                            Porben(j) = CDbl(Format(vlValor / cont_esposa, "#0.00"))
                        'End If
                                         
                        If (iDerCrecerCotizacion = "S") Then 'I--- ABV 14/07/2007
                            If Hijos(u) > 0 Then
                                Codcbe(j) = "S"
                                If Hijos_Inv(u) > 0 Then Codcbe(j) = "N"
                            End If
                        Else
                            Codcbe(j) = "N"
                        End If
                            
                        If Hijos(u) > 0 And Ncorbe(j) = 10 Then
                            vgError = 1000
                            x = MsgBox("Error de código de relación, 'Cónyuge Sin Hijos con Dº Pensión', tiene Hijos.", vbCritical)
                            Exit Function
                        End If
'                    Else
'                        u = Cod_Grfam(j)
'                        If Hijos(u) > 0 Then
'                            'I--- ABV 25/02/2005 ---
'                            'If Coinbe(j) = "T" Then
'                            '    Porben(j) = vlValor     '50
'                            'Else
'                            '    'HQR 16-06-2004
'                            '    'Porben(j) = 36
'                            '    If Coinbe(j) = "P" Then
'                            '        Porben(j) = 36
'                            '    Else
'                            '        Porben(j) = 0
'                            '        cont_esposa = cont_esposa - 1
'                            '    End If
'                            '    'FIN HQR 16-06-2004
'                            'End If
'
'                            Porben(j) = vlValor
'                            If (Coinbe(j) = "N") Then
'                                cont_esposa = cont_esposa - 1
'                            End If
'                            'F--- ABV 25/02/2005 ---
'
'                            If (iDerCrecerCotizacion = "S") Then 'I--- ABV 14/07/2007
'                                Codcbe(j) = "S"
'                                If Hijos_Inv(u) > 0 Then Codcbe(j) = "N"
'                            Else
'                                Codcbe(j) = "N"
'                            End If
'                        Else
'                            'I--- ABV 25/02/2005 ---
'                            'If Coinbe(j) = "T" Then
'                            '    Porben(j) = vlValor     '60
'                            'Else
'                            '    'HQR 16-06-2004
'                            '    'Porben(j) = 43
'                            '    If Coinbe(j) = "P" Then
'                            '        Porben(j) = 43
'                            '    Else
'                            '        Porben(j) = 0
'                            '        cont_esposa = cont_esposa - 1
'                            '    End If
'                            '    'FIN HQR 16-06-2004
'                            'End If
'
'                            Porben(j) = vlValor
'                            If (Coinbe(j) = "N") Then
'                                cont_esposa = cont_esposa - 1
'                            End If
'                            'F--- ABV 25/02/2005 ---
'
'                            If (iDerCrecerCotizacion = "S") Then 'I--- ABV 14/07/2007
'                                Codcbe(j) = "S"
'                                If Hijos_Inv(u) > 0 Then Codcbe(j) = "N"
'                            Else
'                                Codcbe(j) = "N"
'                            End If
'                        End If
                    End If
                Case 20, 21

                    If sexo_cau = "M" Then
                        If Sexobe(j) <> "F" Then
                            vgError = 1000
                            x = MsgBox("Error de código de sexo, el Sexo de la Cónyuge debe ser Femenino.", vbCritical)
                            Exit Function
                        End If
                    Else
                        If Sexobe(j) <> "M" Then
                            vgError = 1000
                            x = MsgBox("Error de codigo de sexo, el Sexo del Cónyuge debe ser Masculino.", vbCritical)
                            Exit Function
                        End If
                    End If
                    
                    u = Cod_Grfam(j)
                    If Hijos(u) = 0 And Ncorbe(j) = 21 Then
                        If iPerGar > 0 Then
                            'If flSiDebeCambiarParentesco(num_pol, iFechaIniVig) = 1 Then
                                If Hijos_SinDerechoPension(u) > cont_hijo And (vgCausaEndoso = "05" Or vgCausaEndoso = "09") Then
                                    Ncorbe(j) = 20
                                    ostBeneficiarios(j).Cod_Par = Ncorbe(j)
                                End If
                            'End If
                        Else
'                            If Hijos_SinDerechoPension(u) = 0 Then
'                                vgError = 1000
'                                X = MsgBox("Error en código de relación 'Madre Con Hijos con Dº Pensión, no tiene Hijos.", vbCritical)
'                                Exit Function
'                            Else
'                                Ncorbe(j) = 20
'                                ostBeneficiarios(j).Cod_Par = 20
'                            End If
                        End If
                    End If
                    
                    'I--- ABV 25/02/2005 ---
                    'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
                    If (fgObtenerPorcentaje(CStr(Ncorbe(j)), Coinbe(j), Sexobe(j), iFechaIniVig, SitInv) = True) Then
                        vlValor = vgValorPorcentaje
                    Else
                        vgError = 1000
                        MsgBox "No existe valor de Porcentaje para el parámetro de :" _
                        & Ncorbe(j) & " - " & Coinbe(j) & " - " & Sexobe(j) & " - " & iFechaIniVig & "."
                        Exit Function
                    End If
                    'F--- ABV 25/02/2005 ---
                    If (vlValor < 0) Then
                        vgError = 1000
                        MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
                        Exit Function
                    Else
                        Porben(j) = vlValor / cont_mhn_tot
                    End If

                    If (iDerCrecerCotizacion = "S") Then 'I--- ABV 14/07/2007
                        If Hijos(u) > 0 Then
                            Codcbe(j) = "S"
                        Else
                            Codcbe(j) = "N"
                        End If
                        If Hijos_Inv(u) > 0 Then
                            Codcbe(j) = "N"
                        End If
                    Else
                        Codcbe(j) = "N"
                    End If

                    If Hijos(u) > 0 And Ncorbe(j) = 20 Then
                        vgError = 1000
                        x = MsgBox("Error en código de relación 'Madre Sin Hijos con Dº Pensión', tiene Hijos.", vbCritical)
                        Exit Function
                    End If
                    
                Case 30
                    Codcbe(j) = "N"
                    Q = Cod_Grfam(j)
                    If cont_esposa > 0 Or cont_mhn(Q) > 0 Then
                        If Coinbe(j) = "N" And edad_mes_ben > vlEsEdadLeg Then
                            Porben(j) = 0
                        Else
                            If vbEsEdadLeg = True Then
                                L18 = IIf(flCompletaRequisitos(num_pol, CInt(j)) = True, L24, L18)
                            End If
                            
                            If (Coinbe(j) = "P" Or Coinbe(j) = "T") Then
                                'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
                                If (fgObtenerPorcentaje(CStr(Ncorbe(j)), Coinbe(j), Sexobe(j), iFechaIniVig, SitInv) = True) Then
                                    vlValor = vgValorPorcentaje
                                Else
                                    vgError = 1000
                                    MsgBox "No existe valor de Porcentaje para el parámetro de :" _
                                    & Ncorbe(j) & " - " & Coinbe(j) & " - " & Sexobe(j) & " - " & iFechaIniVig & "."
                                    Exit Function
                                End If
                                
                                If (vlValor < 0) Then
                                    vgError = 1000
                                    MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
                                    'X = MsgBox("Error: Código de cónyuge no esta incorporado", vbCritical)
                                    Exit Function
                                Else
                                    Porben(j) = vlValor
                                End If
                                'F--- ABV 25/02/2005 ---
                            
                            Else
                                If Der_pen(j) <> "10" Then
                                    If edad_mes_ben <= vlEsEdadLeg Then
                                        'I--- ABV 25/02/2005 ---
                                        If (fgObtenerPorcentaje(CStr(Ncorbe(j)), Coinbe(j), Sexobe(j), iFechaIniVig, SitInv) = True) Then
                                            'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
                                            vlValor = vgValorPorcentaje
                                        Else
                                            vgError = 1000
                                            MsgBox "No existe valor de Porcentaje para el parámetro de :" _
                                            & Ncorbe(j) & " - N - " & Sexobe(j) & " - " & iFechaIniVig & "."
                                            Exit Function
                                        End If
                                        'F--- ABV 25/02/2005 ---
                                        
                                        If (vlValor < 0) Then
                                            vgError = 1000
                                            'X = MsgBox("Error: Código de cónyuge no esta incorporado", vbCritical)
                                            MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
                                            Exit Function
                                        Else
                                            'Porben(j) = 15
                                            Porben(j) = vlValor
                                        End If
                                    End If
                                Else
'                                    If edad_mes_ben <= 216 Then
'                                        MsgBox "No se puede excluir un hijo que aun es menor de edad.", vbCritical, "Error de Datos"
'                                        Exit Function
'                                    End If
'                                    If L18 > 216 Then
'                                        MsgBox "No se puede excluir un hijo que aun esta estudiando.", vbCritical, "Error de Datos"
'                                        Exit Function
'                                    End If
                                    'If vgCausaEndoso <> "26" Then
                                    '
                                    'End If
                                    vlValor = 0
                                    Porben(j) = vlValor
                                    'Der_pen(j) = "10"
                                End If
                            End If
                        End If
                    Else
                        Q = Cod_Grfam(j)
                        Codcbe(j) = "N"

                        If cont_esposa = 0 And cont_mhn(Q) = 0 Then
                            If vbEsEdadLeg = True Then
                                L18 = IIf(flCompletaRequisitos(num_pol, CInt(j)) = True, L24, L18)
                            End If
                            If (Coinbe(j) = "P" Or Coinbe(j) = "T") Then
                                If (fgObtenerPorcentaje(CStr(Ncorbe(j)), Coinbe(j), Sexobe(j), iFechaIniVig, SitInv) = True) Then
                                    'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
                                    vlValor = vgValorPorcentaje
                                Else
                                    vgError = 1000
                                    MsgBox "No existe valor de Porcentaje para el parámetro de :" _
                                    & Ncorbe(j) & " - " & Coinbe(j) & " - " & Sexobe(j) & " - " & iFechaIniVig & "."
                                    Exit Function
                                End If
                                
                                If vlValor < 0 Then
                                    MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
                                    Exit Function
                                Else
                                    vlValorHijo = vlValor
                                    Porben(j) = vlValor
                                End If

                            Else
                                If Der_pen(j) <> "10" Then
                                    If edad_mes_ben <= vlEsEdadLeg Then
                                        If (fgObtenerPorcentaje(CStr(Ncorbe(j)), Coinbe(j), Sexobe(j), iFechaIniVig, SitInv) = True) Then
                                            'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
                                            vlValor = vgValorPorcentaje
                                        Else
                                            vgError = 1000
                                            MsgBox "No existe valor de Porcentaje para el parámetro de :" _
                                            & Ncorbe(j) & " - " & Coinbe(j) & " - " & Sexobe(j) & " - " & iFechaIniVig & "."
                                            Exit Function
                                        End If
                                        
                                        If vlValor < 0 Then
                                            MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
                                            Exit Function
                                        Else
                                            vlValorHijo = vlValor
                                            Porben(j) = vlValor
                                        End If
                                    End If
                                Else
'                                    If L18 > 216 Then
'                                        MsgBox "No se puede excluir un hijo que aun esta estudiando.", vbCritical, "Error de Datos"
'                                        Exit Function
'                                    End If
                                    vlValor = 0
                                    Porben(j) = vlValor
                                    'Der_pen(j) = "10"
                                End If
                            End If
                            '--modificacion RRR 22072020
                            'If vgCausaEndoso = "02" Or vgCausaEndoso = "05" Or vgCausaEndoso = "07" Or vgCausaEndoso = "09" Then
                                If cont_esposa = 0 And cont_hijo = 1 Then
                                    'Buscar el Porcentaje de una Cónyuge
                                    If (fgObtenerPorcentaje("10", "N", "F", iFechaIniVig, SitInv) = True) Then
                                        vlValor = vgValorPorcentaje
                                    Else
                                        vgError = 1000
                                        MsgBox "No existe valor de Porcentaje para el parámetro de :" _
                                        & "10" & " - " & "N" & " - " & "F" & " - " & iFechaIniVig & "."
                                        Exit Function
                                    End If
                                ElseIf cont_esposa = 0 And cont_hijo > 1 Then 'MVG 20170303
                                    If (fgObtenerPorcentaje("10", "N", "F", iFechaIniVig, SitInv) = True) Then
                                        vlValor = vgValorPorcentaje
                                    End If
                                    
                                End If
    
                                If vlValor < 0 Then
                                    MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
                                    Exit Function
                                Else
                                    v_hijo = vlValor
                                    If vbEsEdadLeg = True Then
                                         L18 = IIf(flCompletaRequisitos(num_pol, CInt(i)) = True, L24, L18)
                                    End If
                                    If Coinbe(j) = "N" And edad_mes_ben <= vlEsEdadLeg Then
                                        If cont_hijo = 1 Then
                                            Porben(j) = v_hijo
                                        Else
                                            Porben(j) = v_hijo / cont_hijo + vlValorHijo
                                        End If
                                    Else
                                        If Coinbe(j) = "T" Or Coinbe(j) = "P" Then
                                            If cont_hijo = 1 Then
                                                Porben(j) = v_hijo
                                            Else
                                                Porben(j) = v_hijo / cont_hijo + vlValorHijo
                                            End If
                                        End If
                                    End If
                                End If
                            'End If
                            '--fin modificacion RRR 22072020
                            
                        End If
                    End If
                Case 41, 42
                
                    If Ncorbe(j) = 41 Then
                        If Sexobe(j) <> "M" Then
                            vgError = 1000
                            x = MsgBox("Error de código de sexo, el Sexo de la Padre debe ser Masculino.", vbCritical)
                            Exit Function
                        End If
                    Else
                        If Sexobe(j) <> "F" Then
                            vgError = 1000
                            x = MsgBox("Error de codigo de sexo, el Sexo del Madre debe ser Femenino.", vbCritical)
                            Exit Function
                        End If
                    End If

                
                
                    'I--- ABV 25/02/2005 ---
                    If (fgObtenerPorcentaje(CStr(Ncorbe(j)), Coinbe(j), Sexobe(j), iFechaIniVig, SitInv) = True) Then
                        'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
                        vlValor = vgValorPorcentaje
                    Else
                        vgError = 1000
                        MsgBox "No existe valor de Porcentaje para el parámetro de :" _
                        & Ncorbe(j) & " - " & Coinbe(j) & " - " & Sexobe(j) & " - " & iFechaIniVig & "."
                        Exit Function
                    End If
                    'F--- ABV 25/02/2005 ---
                    
                    If (vlValor < 0) Then
                        vgError = 1000
                        MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
                        Exit Function
                    Else
                        Codcbe(j) = "N"
                        Porben(j) = vlValor
                    End If

            End Select
        End If
        vlSumPrcPns = vlSumPrcPns + vlValor
    Next j
NoValora:
    For k = 1 To vlNum_Ben '60
        If derpen(k) <> 10 Then
            Select Case Ncorbe(k)
                Case 11
                    Q = Cod_Grfam(k)
                    'If codcbe(k) = "S" Then DAJ 10/08/2007
                        Fec_NacHM(k) = Hijo_Menor_Ant(Q)  '******Agruegué 05/11/2000 HILDA
                    'End If DAJ 10/08/2007
                Case 21
                    Q = Cod_Grfam(k)
                    'If codcbe(k) = "S" Then DAJ 10/08/2007
                        Fec_NacHM(k) = Hijo_Menor_Ant(Q)  '*******Agruegué 05/11/2000 HILDA
                    'End If DAJ 10/08/2007
            End Select
        End If
    Next k
    
'''    If (iPerGar > 0) Then
'''
'''        For o = 1 To vlNum_Ben
'''                PorbenGar(o) = Round((Porben(o) / vlSumPrcPns) * 100, 2)
'''            Next o
'''        If vlNum_Ben <= 2 Then
'''            PorbenGar(1) = 0
'''            PorbenGar(2) = 100
'''        End If
'''    End If
    
    
    vlSumarTotalPorcentajePension = 0
    vlSumaPjePenPadres = 0
    j = 0
    For j = 1 To vlNum_Ben
        'Guardar el Valor del Porcentaje Calculado
        'If IsNumeric(Porben(j)) Then
        '    grd_beneficiarios.Text = Format(Porben(j), "##0.00")
        'Else
        '    grd_beneficiarios.Text = Format("0", "0.00")
        'End If
        If IsNumeric(Porben(j)) Then
            'I--- ABV 22/06/2005 ---
            'ostBeneficiarios(j).Prc_Pension = Format(Porben(j), "#0.00")
            If (ostBeneficiarios(j).Cod_Par = "99" Or ostBeneficiarios(j).Cod_Par = "0") Then
                If (derpen(j) <> 10) Then
                    If (vlFechaFallCau <> "" And cont_causante = 1) And vlFecTerPerGar > iFechaIniVig Then
                        ostBeneficiarios(j).Prc_Pension = 0
                        ostBeneficiarios(j).Prc_PensionLeg = 0
                        ostBeneficiarios(j).Prc_PensionGar = 100 'Format(PorbenGar(0), "#0.00")
                    Else
                        ostBeneficiarios(j).Prc_Pension = 100
                        ostBeneficiarios(j).Prc_PensionLeg = 100 '*-+
                        ostBeneficiarios(j).Prc_PensionGar = 100
                    End If
                Else
                    ostBeneficiarios(j).Prc_Pension = 0
                    ostBeneficiarios(j).Prc_PensionLeg = 0
                    ostBeneficiarios(j).Prc_PensionGar = 0
                End If
            Else
                '**********marcocomenta if dejando abierto siempre las tasas 30/11/2016
'                If vf = True Then
'                    ostBeneficiarios(j).Prc_Pension = 0
'                    ostBeneficiarios(j).Prc_PensionLeg = 0
'                Else
                    ostBeneficiarios(j).Prc_Pension = Format(Porben(j), "#0.00")
                    ostBeneficiarios(j).Prc_PensionLeg = Format(Porben(j), "#0.00") '*-+
'                End If
                ostBeneficiarios(j).Prc_PensionGar = Format(PorbenGar(j), "#0.00")
            End If
            'F--- ABV 22/06/2005 ---
        Else
            ostBeneficiarios(j).Prc_Pension = 0
            ostBeneficiarios(j).Prc_PensionLeg = 0 '*-+
            ostBeneficiarios(j).Prc_PensionGar = 0
        End If
        
        'Guardar el Derecho a Acrecer de los Beneficiarios
        'Inicio
        'If (Codcbe(j) <> Empty And Not IsNull(Codcbe(j))) Then
        '    grd_beneficiarios.Text = Codcbe(j)
        'Else
        '    'Por Defecto Negar el Derecho a Acrecer de los Beneficiarios
        '    grd_beneficiarios.Text = "N"
        'End If
        If (Codcbe(j) <> Empty And Not IsNull(Codcbe(j))) Then
            ostBeneficiarios(j).Cod_DerCre = Codcbe(j)
        Else
            ostBeneficiarios(j).Cod_DerCre = "N"
        End If
        'Fin
        
        'Guardar la Fecha de Nacimiento del Hijo Menor de la Cónyuge
        'Inicio
        If Format(CDate(Fec_NacHM(j)), "yyyymmdd") > "18991230" Then
            'ostBeneficiarios(j).Fec_NacHM = CDate(Fec_NacHM(j))
            ostBeneficiarios(j).Fec_NacHM = Format(CDate(Fec_NacHM(j)), "yyyymmdd")
        Else
            ostBeneficiarios(j).Fec_NacHM = ""
        End If
        'Fin

'I--- ABV 10/08/2007 ---
        'Inicializar el Monto de Pensión a Cero
        '--------------------------------------
        'Actualizar el Monto de la Pensión
        ostBeneficiarios(j).Mto_Pension = Format(0, "#0.00")
        
         
        'Actualizar el Monto de la Pensión Garantizada
        If (derpen(j) = 10) Then
            'If (vlFechaFallecimiento <> "" And cont_causante = 0) And vlFecTerPerGar > iFechaIniVig Then
               
            'Else
                ostBeneficiarios(j).Mto_PensionGar = Format(0, "#0.00")
            'End If
            
        End If
        ostBeneficiarios(j).Mto_PensionGar = Format(0, "#0.00")
'F--- ABV 10/08/2007 ---
        
        If (ostBeneficiarios(j).Cod_DerPen <> "10") And (ostBeneficiarios(j).Cod_Par <> "99" And ostBeneficiarios(j).Cod_Par <> "0") Then
            vlSumarTotalPorcentajePension = vlSumarTotalPorcentajePension + ostBeneficiarios(j).Prc_Pension
        End If
'F--- ABV 21/06/2006 ---
    'RRR 04/06/2013
        If (ostBeneficiarios(j).Cod_DerPen <> "10") And ((ostBeneficiarios(j).Cod_Par = "41" Or ostBeneficiarios(j).Cod_Par = "42") And ostBeneficiarios(j).Cod_Par <> "0") Then
            vlSumaPjePenPadres = vlSumaPjePenPadres + ostBeneficiarios(j).Prc_Pension
        End If

       ' If vgTipoCauEnd = "S" Then ostBeneficiarios(j).Cod_EstPension = derpen(j)
    Next j
    
    'guardarlo para luego
    'RRR 03/06/2013
    Dim fx As Double
    Dim SumaTotalDef, SumaTotalDefGar As Double

    If iCobCobertura = "S" Then
        Select Case SitInv
            Case "N": fx = 1
            Case "T": fx = 0.7
            Case "P": fx = 0.5
        End Select
    Else
        fx = 1
    End If
    
    
    SumaTotalDef = 0
    vlPorcentajeRecal = 0
    If cont_padres > 0 And vlSumarTotalPorcentajePension > 100 Then
        If vlSumarTotalPorcentajePension - vlSumaPjePenPadres <= 100 Then
            SumaTotalDef = 0
            For j = 1 To vlNum_Ben
                If (ostBeneficiarios(j).Cod_DerPen <> "10") And ((ostBeneficiarios(j).Cod_Par = "41" Or ostBeneficiarios(j).Cod_Par = "42") And ostBeneficiarios(j).Cod_Par <> "0") Then
                    vlPorcentajeRecal = (100 - (vlSumarTotalPorcentajePension - vlSumaPjePenPadres)) / cont_padres
                    ostBeneficiarios(j).Prc_Pension = Format(vlPorcentajeRecal, "#0.00")
                End If
                If ostBeneficiarios(j).Cod_Par <> "99" Then
'                    If Fec_Fall(j) <> "" Then
'                        ostBeneficiarios(j).Prc_Pension = 0
'                    End If
                    SumaTotalDef = SumaTotalDef + ostBeneficiarios(j).Prc_Pension
                End If
            Next j
            SumaTotalDefGar = 0
            For j = 1 To vlNum_Ben
                If (ostBeneficiarios(j).Cod_DerPen <> "10") And (ostBeneficiarios(j).Cod_Par <> "99" And ostBeneficiarios(j).Cod_Par <> "0") Then
                    vlPorcentajeRecal = (ostBeneficiarios(j).Prc_Pension / (SumaTotalDef)) * 100
                    ostBeneficiarios(j).Prc_Pension = Format((vlPorcentajeRecal / fx), "#0.00")
                    ostBeneficiarios(j).Prc_PensionLeg = Format(vlPorcentajeRecal, "#0.00")
                    'ostBeneficiarios(j).Prc_PensionGar = Format((ostBeneficiarios(j).Prc_Pension / SumaTotalDef) * 100, "#0.00")
                    If ostBeneficiarios(j).Cod_Par <> "99" Then
                        If Fec_Fall(j) <> "" Or Der_pen(j) = "10" Then
                                    ostBeneficiarios(j).Prc_Pension = 0
                                    ostBeneficiarios(j).Prc_PensionLeg = 0
                                    ostBeneficiarios(j).Prc_PensionGar = 0
                                    ostBeneficiarios(j).Cod_DerPen = "10"
                                    ostBeneficiarios(j).Cod_EstPension = "10"
                            End If
                        SumaTotalDefGar = SumaTotalDefGar + ostBeneficiarios(j).Prc_Pension
                    End If
                End If
            Next j
            'OBTIENE EL PORCENTAJE GARANTIZADO
            If iTipoPension = "08" Or iTipoPension = "09" Or iTipoPension = "10" Or iTipoPension = "11" Or iTipoPension = "12" Then
                For j = 1 To vlNum_Ben
                    If (ostBeneficiarios(j).Cod_DerPen <> "10") And (ostBeneficiarios(j).Cod_Par <> "99" And ostBeneficiarios(j).Cod_Par <> "0") Then
                        'ostBeneficiarios(j).Prc_Pension = Format(((ostBeneficiarios(j).Prc_Pension / (SumaTotalDef)) * 100), "#0.00")
                        'ostBeneficiarios(j).Prc_PensionLeg = Format(((ostBeneficiarios(j).Prc_Pension / (SumaTotalDef)) * 100), "#0.00")
                        ostBeneficiarios(j).Prc_PensionGar = Format((ostBeneficiarios(j).Prc_Pension / SumaTotalDefGar) * 100, "#0.00")
                    End If
                Next j
            End If
            
            
        Else
        
        End If
    Else
        'Saca el valor por el factor de la invalidez
        
        For j = 1 To vlNum_Ben
                If (ostBeneficiarios(j).Cod_DerPen <> "10") And (ostBeneficiarios(j).Cod_Par <> "99" And ostBeneficiarios(j).Cod_Par <> "0") Then
                    ostBeneficiarios(j).Prc_Pension = (ostBeneficiarios(j).Prc_PensionLeg / fx)
                    vlPorcentajeRecal = ostBeneficiarios(j).Prc_Pension
                    SumaTotalDef = SumaTotalDef + vlPorcentajeRecal
                End If
        Next j
        If iTipoPension = "08" Or iTipoPension = "09" Or iTipoPension = "10" Or iTipoPension = "11" Or iTipoPension = "12" Then
'            For j = 1 To vlNum_Ben
'                If (ostBeneficiarios(j).Cod_DerPen <> "10") And (ostBeneficiarios(j).Cod_Par <> "99" And ostBeneficiarios(j).Cod_Par <> "0") Then
'                    ostBeneficiarios(j).Prc_Pension = (ostBeneficiarios(j).Prc_PensionLeg)
'                    vlPorcentajeRecal = ostBeneficiarios(j).Prc_Pension
'                    SumaTotalDef = SumaTotalDef + vlPorcentajeRecal
'                End If
'            Next j
            SumaTotalDefGar = 0
            If (SumaTotalDef > 100) Then
                For j = 1 To vlNum_Ben
                    If (ostBeneficiarios(j).Cod_DerPen <> "10") And (ostBeneficiarios(j).Cod_Par <> "99" And ostBeneficiarios(j).Cod_Par <> "0") Then
                        'vlPorcentajeRecal = Format((ostBeneficiarios(j).Prc_Pension / SumaTotalDef) * 100, "#0.00")
                        'ostBeneficiarios(j).Prc_Pension = Format(vlPorcentajeRecal, "#0.00")
                        'ostBeneficiarios(j).Prc_PensionLeg = Format(ostBeneficiarios(j).Prc_PensionLeg, "#0.00")
                        If ostBeneficiarios(j).Cod_Par <> "99" Then
                            If Fec_Fall(j) <> "" Or Der_pen(j) = "10" Then
                                ostBeneficiarios(j).Prc_Pension = 0
                                ostBeneficiarios(j).Prc_PensionLeg = 0
                                ostBeneficiarios(j).Cod_DerPen = "10"
                                ostBeneficiarios(j).Cod_EstPension = "10"
                            End If
                            SumaTotalDefGar = SumaTotalDefGar + ostBeneficiarios(j).Prc_Pension
                        End If
                    End If
                Next j
                ''periodo garatizado
                For j = 1 To vlNum_Ben
                    If (ostBeneficiarios(j).Cod_DerPen <> "10") And (ostBeneficiarios(j).Cod_Par <> "99" And ostBeneficiarios(j).Cod_Par <> "0") Then
                        ostBeneficiarios(j).Prc_PensionGar = Format((ostBeneficiarios(j).Prc_Pension / SumaTotalDefGar) * 100, "#0.00")
                    End If
                Next j
            ElseIf (SumaTotalDef > 0) Then
                For j = 1 To vlNum_Ben
                    If (ostBeneficiarios(j).Cod_DerPen <> "10") And (ostBeneficiarios(j).Cod_Par <> "99" And ostBeneficiarios(j).Cod_Par <> "0") Then
                        'ostBeneficiarios(j).Prc_Pension = Format(ostBeneficiarios(j).Prc_Pension, "#0.00")
                        'ostBeneficiarios(j).Prc_PensionLeg = Format(ostBeneficiarios(j).Prc_PensionLeg, "#0.00")
                        'ostBeneficiarios(j).Prc_PensionGar = Format((ostBeneficiarios(j).Prc_Pension / SumaTotalDefGar) * 100, "#0.00") ' mvg comenta para revisar
                        If ostBeneficiarios(j).Cod_Par <> "99" Then
                            If Fec_Fall(j) <> "" Or Der_pen(j) = "10" Then
                                ostBeneficiarios(j).Prc_Pension = 0
                                ostBeneficiarios(j).Prc_PensionLeg = 0
                                ostBeneficiarios(j).Cod_DerPen = "10"
                                ostBeneficiarios(j).Cod_EstPension = "10"
                            End If
                            SumaTotalDefGar = SumaTotalDefGar + ostBeneficiarios(j).Prc_Pension
                        End If
                    End If
                Next j
                ''periodo garatizado
                For j = 1 To vlNum_Ben
                    If (ostBeneficiarios(j).Cod_DerPen <> "10") And (ostBeneficiarios(j).Cod_Par <> "99" And ostBeneficiarios(j).Cod_Par <> "0") Then
                        ostBeneficiarios(j).Prc_PensionGar = Format((ostBeneficiarios(j).Prc_Pension / IIf(SumaTotalDefGar = 0, 1, SumaTotalDefGar)) * 100, "#0.00")
                    End If
                Next j
            End If
        Else
            If (SumaTotalDef > 100) Then
                For j = 1 To vlNum_Ben
                        If (ostBeneficiarios(j).Cod_DerPen <> "10") And (ostBeneficiarios(j).Cod_Par <> "99" And ostBeneficiarios(j).Cod_Par <> "0") Then
                            vlPorcentajeRecal = Format((ostBeneficiarios(j).Prc_Pension / 100) * 100, "#0.00")
                            ostBeneficiarios(j).Prc_Pension = Format(vlPorcentajeRecal, "#0.00")
                            ostBeneficiarios(j).Prc_PensionLeg = Format(ostBeneficiarios(j).Prc_PensionLeg, "#0.00")
                            ostBeneficiarios(j).Prc_PensionGar = Format(0, "#0.00")
                            If Fec_Fall(j) <> "" Or Der_pen(j) = "10" Then
                                    ostBeneficiarios(j).Prc_Pension = 0
                                    ostBeneficiarios(j).Prc_PensionLeg = 0
                                    ostBeneficiarios(j).Prc_PensionGar = 0
                                    ostBeneficiarios(j).Cod_DerPen = "10"
                                    ostBeneficiarios(j).Cod_EstPension = "10"
                            End If
                        End If
                Next j
            Else
                For j = 1 To vlNum_Ben
                        If (ostBeneficiarios(j).Cod_DerPen <> "10") And (ostBeneficiarios(j).Cod_Par <> "99" And ostBeneficiarios(j).Cod_Par <> "0" And ostBeneficiarios(j).Cod_EstPension <> "10" And ostBeneficiarios(j).Cod_EstPension <> "20") Then 'mvg agrego Cod_EstPension and else
                            ostBeneficiarios(j).Prc_PensionLeg = Format(ostBeneficiarios(j).Prc_PensionLeg, "#0.00")
                            ostBeneficiarios(j).Prc_PensionGar = Format(0, "#0.00")
                            If Fec_Fall(j) <> "" Or Der_pen(j) = "10" Then
                                    ostBeneficiarios(j).Prc_Pension = 0
                                    ostBeneficiarios(j).Prc_PensionLeg = 0
                                    ostBeneficiarios(j).Prc_PensionGar = 0
                                    ostBeneficiarios(j).Cod_DerPen = "10"
                                    ostBeneficiarios(j).Cod_EstPension = "10"
                            End If
                            'SumaTotalDefGar = SumaTotalDefGar + ostBeneficiarios(j).Prc_Pension
                            
'                            'ostBeneficiarios(j).Prc_PensionGar = Format((ostBeneficiarios(j).Prc_Pension / SumaTotalDef) * 100, "#0.00")
'                            If derpen(1) = "99" Then
'                                ostBeneficiarios(j).Prc_PensionGar = 0
'                            End If
                        Else
                            ostBeneficiarios(j).Prc_PensionLeg = Format(ostBeneficiarios(j).Prc_Pension, "#0.00")
                            'ostBeneficiarios(j).Prc_PensionGar = Format((ostBeneficiarios(j).Prc_PensionGar / SumaTotalDef) * 100, "#0.00")  mvg comenta y reemplaza
                            ostBeneficiarios(j).Prc_PensionGar = Format((ostBeneficiarios(j).Prc_PensionGar), "#0.00")
                            If Fec_Fall(j) <> "" Or Der_pen(j) = "10" Then
                                    ostBeneficiarios(j).Prc_Pension = 0
                                    ostBeneficiarios(j).Prc_PensionLeg = 0
                                    ostBeneficiarios(j).Prc_PensionGar = 0
                                    ostBeneficiarios(j).Cod_DerPen = "10"
                                    ostBeneficiarios(j).Cod_EstPension = "10"
                            End If
                        End If
                Next j
            End If
        End If
        
        
    End If
    
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'**********************Validar si los Porcentajes de Pensión Suman más del 100%**************************************************************************
'**********************Validacion de existencia de padres y eliminar o disminuir su % segun corresponda**************************************************
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     
'''''    If cont_padres > 0 And vlSumarTotalPorcentajePension > 100 Then
'''''        If vlSumarTotalPorcentajePension - vlSumaPjePenPadres <= 100 Then
'''''            For j = 1 To vlNum_Ben
'''''                If (ostBeneficiarios(j).Cod_DerPen <> "10") And ((ostBeneficiarios(j).Cod_Par = "41" Or ostBeneficiarios(j).Cod_Par = "42") And ostBeneficiarios(j).Cod_Par <> "0") Then
'''''                    vlPorcentajeRecal = (100 - (vlSumarTotalPorcentajePension - vlSumaPjePenPadres)) / cont_padres
'''''                    ostBeneficiarios(j).Prc_Pension = Format(vlPorcentajeRecal, "#0.00")
'''''                End If
'''''            Next j
'''''        Else
'''''            For j = 1 To vlNum_Ben
'''''                If (ostBeneficiarios(j).Cod_DerPen <> "10") And ((ostBeneficiarios(j).Cod_Par = "41" Or ostBeneficiarios(j).Cod_Par = "42") And ostBeneficiarios(j).Cod_Par <> "0") Then
'''''                    vlPorcentajeRecal = 0
'''''                    ostBeneficiarios(j).Prc_Pension = Format(vlPorcentajeRecal, "#0.00")
'''''                    ostBeneficiarios(j).Cod_DerPen = "10"
'''''                End If
'''''            Next j
'''''
'''''            vlSumaDef = 0
'''''            For j = 1 To vlNum_Ben
'''''                If (ostBeneficiarios(j).Cod_DerPen <> "10") And (ostBeneficiarios(j).Cod_Par <> "99" And ostBeneficiarios(j).Cod_Par <> "0") Then
'''''                    vlPorcentajeRecal = (ostBeneficiarios(j).Prc_PensionLeg / (vlSumarTotalPorcentajePension - vlSumaPjePenPadres)) * 100
'''''                    ostBeneficiarios(j).Prc_Pension = Format(vlPorcentajeRecal, "#0.00")
'''''                    vlSumaDef = vlSumaDef + ostBeneficiarios(j).Prc_Pension
'''''                End If
'''''            Next j
'''''            vlDif = Format(100 - vlSumaDef, "#0.00")
'''''            If (vlDif < 0) Or (vlDif > 0) Then
'''''                For j = 1 To vlNum_Ben
'''''                    If (ostBeneficiarios(j).Cod_DerPen <> "10") And (ostBeneficiarios(j).Cod_Par <> "99" And ostBeneficiarios(j).Cod_Par <> "0") Then
'''''                        vlPorcentajeRecal = ostBeneficiarios(j).Prc_Pension + vlDif
'''''                        ostBeneficiarios(j).Prc_Pension = vlPorcentajeRecal
'''''                        Exit For
'''''                    End If
'''''                Next
'''''            End If
'''''        End If
'''''    Else
'''''        If (vlSumarTotalPorcentajePension > 100) Then
'''''            vlSumaDef = 0
'''''            For j = 1 To vlNum_Ben
'''''                If (ostBeneficiarios(j).Cod_DerPen <> "10") And (ostBeneficiarios(j).Cod_Par <> "99" And ostBeneficiarios(j).Cod_Par <> "0") Then
'''''                    vlPorcentajeRecal = (ostBeneficiarios(j).Prc_Pension / vlSumarTotalPorcentajePension) * 100
'''''                    ostBeneficiarios(j).Prc_Pension = Format(vlPorcentajeRecal, "#0.00")
'''''                    vlSumaDef = vlSumaDef + ostBeneficiarios(j).Prc_Pension
'''''                End If
'''''            Next j
'''''            vlDif = Format(100 - vlSumaDef, "#0.00")
'''''            If (vlDif < 0) Or (vlDif > 0) Then
'''''                'Asignar la diferencia de Porcentaje al Primer Beneficiario con Derecho distinto del Causante
'''''                For j = 1 To vlNum_Ben
'''''                    If (ostBeneficiarios(j).Cod_DerPen <> "10") And (ostBeneficiarios(j).Cod_Par <> "99" And ostBeneficiarios(j).Cod_Par <> "0") Then
'''''                        vlPorcentajeRecal = ostBeneficiarios(j).Prc_Pension + vlDif
'''''                        ostBeneficiarios(j).Prc_Pension = vlPorcentajeRecal
'''''                        Exit For
'''''                    End If
'''''                Next j
'''''            End If
'''''        End If
'''''    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'*********************************************************FIN********************************************************************************************
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



'*-+
'I--- ABV 04/08/2007 ---
'Recalcular Porcentajes si se trata de un Caso de Invalidez - Con Cobertura
'    If (iTipoPension = "06" Or iTipoPension = "07") And (iCobCobertura = "S") Then
'        If (fgObtenerPorcCobertura(iTipoPension, iFechaIniVig, vlRemuneracion, vlPrcCobertura) = True) Then
'            'Determinar la Remuneración Promedio para el Causante
'            vlRemuneracionBase = vlRemuneracion * (vlPrcCobertura / 100)
'
'            'Determinar para cada Beneficiario el Nuevo Porcentaje
'            For j = 1 To (vlContBen - 1)
'                If (ostBeneficiarios(j).Cod_DerPen <> "10") Then
'                    If (ostBeneficiarios(j).Cod_Par = "99") Then
'                        vlRemuneracionProm = vlRemuneracion * (vlPrcCobertura / 100)
'                    Else
'                        vlRemuneracionProm = vlRemuneracion * (ostBeneficiarios(j).Prc_Pension / 100)
'                    End If
'                    vlPorcentajeRecal = (vlRemuneracionProm / vlRemuneracionBase) * 100
'                    ostBeneficiarios(j).Prc_Pension = Format(vlPorcentajeRecal, "#0.00")
'                End If
'            Next j
'        Else
'            vgError = 1000
'            MsgBox "No existen valores de Porcentaje de Cobertura para el Recalculo de Pensiones.", vbCritical, "Inexistencia de Datos"
'            Exit Function
'        End If
'    End If
    
'PARA LOS COSOS DE JUBILACION
    If (iTipoPension = "09" Or iTipoPension = "10") And (iCobCobertura = "N") Then
        'If (fgObtenerPorcCobertura(iTipoPension, iFechaIniVig, vlRemuneracion, vlPrcCobertura) = True) Then
            'Determinar la Remuneración Promedio para el Causante
            'vlRemuneracionBase = vlRemuneracion * (vlPrcCobertura / 100)
            '
            'Determinar para cada Beneficiario el Nuevo Porcentaje
            For j = 1 To (vlContBen - 1)
                If (ostBeneficiarios(j).Cod_DerPen <> "10") Then
                    If (ostBeneficiarios(j).Cod_Par = "99") Then
                        'vlRemuneracionProm = vlRemuneracion * (vlPrcCobertura / 100)
                    Else
                        ostBeneficiarios(j).Prc_Pension = (ostBeneficiarios(j).Prc_Pension * fx)
                        ostBeneficiarios(j).Prc_PensionLeg = ostBeneficiarios(j).Prc_Pension
                    End If
                End If
            Next j
        'Else
            'vgError = 1000
            'MsgBox "No existen valores de Porcentaje de Cobertura para el Recalculo de Pensiones.", vbCritical, "Inexistencia de Datos"
            'Exit Function
        'End If
    End If
    
'F--- ABV 04/08/2007 ---
'*-+
End If


If (vlFechaFallecimiento <> "" And cont_causante = 0) And vlFecTerPerGar > iFechaIniVig Then
    iCalcularPension = True
End If

'Recalcular los Montos de Pensión
If (iCalcularPension = True) Then
    For j = 1 To (vlNum_Ben)
            
        vlPorcBenef = 0
        vlPenBenef = 0
        vlPenGarBenef = 0
                
        'Verificar el Estado de Derecho a Pensión (Cod_EstPension)
'            If (ostBeneficiarios(j).Cod_EstPension <> "10") Then
        If (ostBeneficiarios(j).Cod_DerPen <> "10" Or (vlFechaFallecimiento <> "" And vlFecTerPerGar > iFechaIniVig)) Then
            vlPorcBenef = ostBeneficiarios(j).Prc_Pension
            vlPorcGarBenef = ostBeneficiarios(j).Prc_PensionGar
            
            'Determinar el monto de la Pensión del Beneficiario
            'Penben(i) = IIf(Msf_BMGrilla.TextMatrix(vlI, 17) = "", 0, Msf_BMGrilla.TextMatrix(vlI, 17))
            vlPenBenef = iPensionRef * (vlPorcBenef / 100)
        
            'Definir la Pensión del Garantizado
            'vlPenGarBenef = ostBeneficiarios(j).Mto_PensionGar
            vlPenGarBenef = iPensionRefGar * (vlPorcGarBenef / 100)
        
        
            'I--- ABV 20/04/2005 ---
            vgPalabra = iTipoPension
            vgI = InStr(1, cgPensionInvVejez, vgPalabra)
            If (vgI <> 0) Then
                'Valida que sea un caso de Invalidez o Vejez
                vlPenGarBenef = vlPenBenef
            End If
            
            vgI = InStr(1, cgPensionSobOrigen, vgPalabra)
            If (vgI <> 0) Then
'I--- ABV 17/11/2007 ---
'                'Valida que sea un caso de Sobrevivencia por Origen
'                'Se supone que la Cónyuge o Madre es la única que Garantiza su pensión
'                'vlNumero = InStr(Cmb_BMPar.Text, "-")
'                vgPalabraAux = ostBeneficiarios(j).Cod_Par
'
'                vgX = InStr(1, cgParConyugeMadre, vgPalabraAux)
'                If (vgX <> 0) Then
'                    vlPenGarBenef = vlPenBenef
'                'Else
'                '    Txt_BMMtoPensionGar.Enabled = True
'                End If
                'vlPenGarBenef = vlPenBenef
'F--- ABV 17/11/2007 ---
            End If
            
            vgI = InStr(1, cgPensionSobTransf, vgPalabra)
            If (vgI <> 0) Then
                ''Valida que sea un caso de Invalidez o Vejez
                'Txt_BMMtoPensionGar.Enabled = True
            End If

            'F--- ABV 19/04/2005 ---
        End If
        
        'Actualizar el Monto de la Pensión
        ostBeneficiarios(j).Mto_Pension = Format(vlPenBenef, "#0.00")
        
        'Actualizar el Monto de la Pensión Garantizada
        If (iPerGar > 0) Then
            ostBeneficiarios(j).Mto_PensionGar = Format(vlPenGarBenef, "#0.00")
            'ostBeneficiarios(j).Prc_PensionGar = vlPorcBenef
        Else
            ostBeneficiarios(j).Mto_PensionGar = 0
            ostBeneficiarios(j).Prc_PensionGar = 0
        End If
        
    Next j
End If

    fgCalcularPorcentajeBenef = True
    
Exit Function
Err_fgCalcularPorcentaje:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        vgError = 1000
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fgCalcularPorcentajeBenef_20070925(iFechaIniVig As String, iNumBenef As Integer, ostBeneficiarios() As TyBeneficiarios, Optional iTipoPension As String, Optional iPensionRef As Double, Optional iCalcularPension As Boolean, Optional iDerCrecerCotizacion As String, Optional iCobCobertura As String, Optional iCalcularPorcentaje As Boolean, Optional iPerGar As Long) As Boolean
'Función: Permite actualizar los Porcentajes de Pensión de los Beneficiarios,
'         su Derecho a Acrecer y la Fecha de Nacimiento del Hijo Menor
'Parámetros de Entrada/Salida:
'iFechaIniVig     => Fecha de Inicio de Vigencia de la Póliza
'iNumBenef        => Número de Beneficiarios
'ostBeneficiarios => Estructura desde la cual se obtienen los datos de los
'                    Beneficiarios y al mismo tiempo se calcula el Porcentaje
'                    de Pensión al cual tienen Dº
'iCalcularPension => Permite indicar si se debe realizar el cálculo del Monto de la Pensión que le corresponde a cada Beneficiario
'iTipoPension     => Tipo de Pensión de la Póliza
'iPensionRef      => Monto de la Pensión de Referencia utilizada para el Calculo de la Pensión si el campo anterior esta en Verdadero
'iDerCrecerCotizacion => Indicador de Derecho a Crecer definido en la Cotización (S o N)
'iCobCobertura    => Indicador de Cobertura de la Cotización (S o N)

'Dim vlValor  As Double
'Dim Numor()  As Integer
'Dim Ncorbe() As Integer
'Dim Codrel() As Integer
'Dim Cod_Grfam() As Integer
'Dim Sexobe() As String, Inv()   As String, Coinbe()   As String
'Dim derpen() As Integer
'Dim Nanbe()  As Integer, Nmnbe() As Integer, Ndnbe() As Integer
'Dim Porben() As Double
'Dim Codcbe() As String
'Dim Iaap     As Integer, Immp   As Integer, Iddp    As Integer
'Dim vlNum_Ben As Integer
'Dim Hijos()  As Integer, Hijos_Inv() As Integer
'Dim Hijos_SinDerechoPension()  As Integer 'hqr 08/07/2006
'Dim Hijo_Menor() As Date
'Dim Hijo_Menor_Ant() As Date
'Dim Fec_NacHM() As Date
'
''I--- ABV 18/07/2006 ---
'Dim Hijos_SinConyugeMadre()  As Integer
'Dim Fec_Fall() As String
'Dim vlFechaFallCau As String
'Dim cont_esp_Totales As Integer
'Dim cont_esp_Tot_GF() As Integer
'Dim cont_mhn_Totales As Integer
'Dim cont_mhn_Tot_GF() As Integer
'Dim vlValorHijo  As Double
''F--- ABV 18/07/2006 ---
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
'Dim vlFechaFallecimiento As String
'Dim vlFechaMatrimonio    As String
'
'Dim vlPorcBenef As Double
'Dim vlPenBenef As Double
'Dim vlPenGarBenef As Double
'
'Dim vlSumarTotalPorcentajePension As Double, vlSumaDef As Double, vlDif As Double
'Dim vlPorcentajeRecal As Double
'Dim vlRemuneracion As Double, vlRemuneracionProm As Double
'Dim vlRemuneracionBase As Double
'Dim vlPrcCobertura As Double
'
''I--- ABV 21/06/2007 ---
''On Error GoTo Err_fgCalcularPorcentaje
''F--- ABV 21/06/2007 ---
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
'    'Debiera tomar la Fecha de Devengue
'
'    If (fgCarga_Param("LI", "L24", iFechaIniVig) = True) Then
'        L24 = vgValorParametro
'    Else
'        vgError = 1000
'        MsgBox "No existe Edad de tope para los 24 años.", vbCritical, "Proceso Cancelado"
'        Exit Function
'    End If
'
'    'Mensualizar la Edad de 24 Años
'    L24 = L24 * 12
''I--- ABV 21/07/2006 ---
''Edad de 18 años
'Dim L18 As Long
'If (fgCarga_Param("LI", "L18", iFechaIniVig) = True) Then
'    L18 = vgValorParametro
'Else
'    vgError = 1000
'    MsgBox "No existe Edad de tope para los 18 años.", vbCritical, "Proceso Cancelado"
'    Exit Function
'End If
''Mensualizar la Edad de 24 Años
'L18 = L18 * 12
''F--- ABV 21/07/2006 ---
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
'    ReDim Numor(vlNum_Ben) As Integer
'    ReDim Ncorbe(vlNum_Ben) As Integer
'    ReDim Codrel(vlNum_Ben) As Integer
'    ReDim Cod_Grfam(vlNum_Ben) As Integer
'    ReDim Sexobe(vlNum_Ben) As String
'    ReDim Inv(vlNum_Ben) As String
'    ReDim Coinbe(vlNum_Ben) As String
'    ReDim derpen(vlNum_Ben) As Integer
'    ReDim Nanbe(vlNum_Ben) As Integer
'    ReDim Nmnbe(vlNum_Ben) As Integer
'    ReDim Ndnbe(vlNum_Ben) As Integer
'    ReDim Hijos(vlNum_Ben) As Integer
'    ReDim Hijos_Inv(vlNum_Ben) As Integer
'    ReDim Hijos_SinDerechoPension(vlNum_Ben) As Integer
'    ReDim Hijo_Menor(vlNum_Ben) As Date
'    ReDim Hijo_Menor_Ant(vlNum_Ben) As Date
'    ReDim Porben(vlNum_Ben) As Double
'    ReDim Codcbe(vlNum_Ben) As String
'    ReDim Fec_NacHM(vlNum_Ben) As Date
'    ReDim cont_mhn(vlNum_Ben) As Integer
'
'    ReDim Hijos_SinConyugeMadre(vlNum_Ben) As Integer
'    ReDim Fec_Fall(vlNum_Ben) As String
'    ReDim cont_mhn_Tot_GF(vlNum_Ben) As Integer
'    ReDim cont_esp_Tot_GF(vlNum_Ben) As Integer
'    vlFechaFallCau = ""
'    vlValorHijo = 0
'
'If (iCalcularPorcentaje = True) Then
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
'            vgError = 1000
'            MsgBox "No existe Número de Orden de Beneficiario.", vbCritical, "Error de Datos"
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
'        derpen(i) = ostBeneficiarios(i).Cod_EstPension
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
'        edad_mes_ben = fecha_sin - (Nanbe(i) * 12 + Nmnbe(i))
'
'        vlFechaFallecimiento = ostBeneficiarios(i).Fec_FallBen
'
'        Fec_Fall(i) = vlFechaFallecimiento
'        If Codrel(i) = 99 Then
'            vlFechaFallCau = vlFechaFallecimiento
'        End If
'
'        vlFechaMatrimonio = ostBeneficiarios(i).Fec_Matrimonio
'
'        If (Codrel(i) >= 30 And Codrel(i) < 40) Then
'            If (vlFechaFallecimiento <> "") Or (vlFechaMatrimonio <> "") Then
'                derpen(i) = 10
'                If ostBeneficiarios(i).Cod_EstPension <> "10" Then
'                    ostBeneficiarios(i).Cod_EstPension = "10"
'                    Hijos_SinDerechoPension(ostBeneficiarios(i).Cod_GruFam) = Hijos_SinDerechoPension(ostBeneficiarios(i).Cod_GruFam) + 1
'                End If
'            Else
'                If edad_mes_ben > L24 And Coinbe(i) = "N" Then
'                    derpen(i) = 10
'                    'hqr 08/07/2006 cambiar Estado de pension a No Vigente.  Se debe cambiar estado de madre si es hijo unico
'                    If ostBeneficiarios(i).Cod_EstPension <> "10" Then
'                        ostBeneficiarios(i).Cod_EstPension = "10"
'                        Hijos_SinDerechoPension(ostBeneficiarios(i).Cod_GruFam) = Hijos_SinDerechoPension(ostBeneficiarios(i).Cod_GruFam) + 1
'                    End If
'                Else
'                    derpen(i) = 99
'                End If
'            End If
'        Else
'            'El resto de los Beneficiarios que no sean Hijos, solo se dejan como
'            'Sin Derecho a Pensión cuando están fallecidos
'            If (vlFechaFallecimiento <> "") Then
'                derpen(i) = 10
'            Else
'                derpen(i) = 99
'            End If
'        End If
'        ostBeneficiarios(i).Cod_DerPen = derpen(i)
'
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
'    cont_esp_Totales = 0
'    cont_mhn_Totales = 0
'
'    'Primer Ciclo
'    For g = 1 To vlNum_Ben
'        If derpen(g) <> 10 Then
'            '99: con derecho a pension,20: con Derecho Pendiente
'            '10: sin derecho a pension
'            If Ncorbe(g) = 99 Then
'                cont_causante = cont_causante + 1
'            Else
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
'                        Hijos(Q) = Hijos(Q) + 1
'                        If Coinbe(g) <> "N" Then Hijos_Inv(Q) = Hijos_Inv(Q) + 1
'                        Hijo_Menor(Q) = DateSerial(Nanbe(g), Nmnbe(g), Ndnbe(g))
'                        If Hijos(Q) > 1 Then
'                            If Hijo_Menor(Q) > Hijo_Menor_Ant(Q) Then
'                                Hijo_Menor_Ant(Q) = Hijo_Menor(Q)
'                            End If
'                        Else
'                            Hijo_Menor_Ant(Q) = Hijo_Menor(Q)
'                        End If
''                    Case 35
'                        edad_mes_ben = fecha_sin - (Nanbe(g) * 12 + Nmnbe(g))
'                        If Coinbe(g) = "N" And edad_mes_ben <= L24 Then
'                            cont_hijo = cont_hijo + 1
'                        Else
'                            If Coinbe(g) = "T" Or Coinbe(g) = "P" Then
'                                cont_hijo = cont_hijo + 1
'                            End If
'                        End If
'
'                    Case 41, 42
'                        cont_padres = cont_padres + 1
'                    Case Else
'                        vgError = 1000
'                        X = MsgBox("Error en codificación de codigo de relación", vbCritical)
'                        Exit Function
'                End Select
'            End If
'        Else
'            'Verificar si la cónyuge o la Madre Falleció antes que el Causante para contarla
'            Select Case Ncorbe(g)
'                Case 11:
'                    If (vlFechaFallCau > Fec_Fall(g)) Then
'                        cont_esposa = cont_esposa + 1
'                        Q = Cod_Grfam(g)
'                        cont_esp_Tot_GF(Q) = cont_esp_Tot_GF(Q) + 1
'                        cont_esp_Totales = cont_esp_Totales + 1
'                    End If
'                Case 21:
'
'            End Select
'        End If
'    Next g
'
'
'    j = 1
'    For j = 1 To vlNum_Ben
'        '99: con derecho a pension,20: con Derecho Pendiente
'        '10: sin derecho a pension
'        If derpen(j) <> 10 Then
'            edad_mes_ben = fecha_sin - (Nanbe(j) * 12 + Nmnbe(j))
'            Select Case Ncorbe(j)
'                Case 99
'                    If cont_causante > 1 Then
'                        vgError = 1000
'                        X = MsgBox("Error en codificación de codigo de relación, No puede ingresar otro causante", vbCritical)
'                        Exit Function
'                    End If
'                    'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
'                    If (fgObtenerPorcentaje(CStr(Ncorbe(j)), Coinbe(j), Sexobe(j), iFechaIniVig) = True) Then
'                        vlValor = vgValorPorcentaje
'                    Else
'                        vgError = 1000
'                        MsgBox "No existe valor de Porcentaje para el parámetro de :" _
'                        & Ncorbe(j) & " - " & Coinbe(j) & " - " & Sexobe(j) & " - " & iFechaIniVig & "."
'                        Exit Function
'                    End If
'                    If (vlValor < 0) Then
'                        vgError = 1000
'                        MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
'                        'error
'                        Exit Function
'                    Else
'                        Porben(j) = vlValor
'                        Codcbe(j) = "N"
'                    End If
'                Case 10, 11
'                    If sexo_cau = "M" Then
'                        If Sexobe(j) <> "F" Then
'                            vgError = 1000
'                            X = MsgBox("Error de código de sexo, el Sexo de la Cónyuge debe ser Femenino.", vbCritical)
'                            Exit Function
'                        End If
'                    Else
'                        If Sexobe(j) <> "M" Then
'                            vgError = 1000
'                            X = MsgBox("Error de codigo de sexo, el Sexo del Cónyuge debe ser Masculino.", vbCritical)
'                            Exit Function
'                        End If
'                    End If
'                    'hqr 08/07/2006 validacion y cambio parentesco
'                    u = Cod_Grfam(j)
'                    If Hijos(u) = 0 And Ncorbe(j) = 11 Then
'                        If Hijos_SinDerechoPension(u) = 0 Then
'                            vgError = 1000
'                            X = MsgBox("Error de código de relación, 'Cónyuge Con Hijos con Dº Pensión', no tiene Hijos.", vbCritical)
'                            Exit Function
'                        Else
'                            Ncorbe(j) = 10
'                            ostBeneficiarios(j).Cod_Par = 10
'                        End If
'                    End If
'
'                    'HQR 08/07/2006 se deja al final porque se cambia el tipo de parentesco
'                    'I--- ABV 25/02/2005 ---
'                     'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
'                     If (fgObtenerPorcentaje(CStr(Ncorbe(j)), Coinbe(j), Sexobe(j), iFechaIniVig) = True) Then
'                         vlValor = vgValorPorcentaje
'                     Else
'                         vgError = 1000
'                         MsgBox "No existe valor de Porcentaje para el parámetro de :" _
'                         & Ncorbe(j) & " - " & Coinbe(j) & " - " & Sexobe(j) & " - " & iFechaIniVig & "."
'                         Exit Function
'                     End If
'                     'F--- ABV 25/02/2005 ---
'                     If (vlValor < 0) Then
'                         vgError = 1000
'                         MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
'                         'error
'                         Exit Function
'                     End If
'
'                    'fin hqr 08/07/2006
'                    If sexo_cau = "M" Or sexo_cau = "F" Then
'                        Porben(j) = CDbl(Format(vlValor / cont_esposa, "#0.00"))
'
'                        'If (iDerCrecerCotizacion = "S") Then 'I--- ABV 14/07/2007
'                        If Hijos(u) > 0 Then
'                            Codcbe(j) = "S"
'                            If Hijos_Inv(u) > 0 Then Codcbe(j) = "N"
'                        End If
'                        'Else
'                        '    Codcbe(j) = "N"
'                        'End If
'
'                        If Hijos(u) > 0 And Ncorbe(j) = 10 Then
'                            vgError = 1000
'                            X = MsgBox("Error de código de relación, 'Cónyuge Sin Hijos con Dº Pensión', tiene Hijos.", vbCritical)
'                            Exit Function
'                        End If
''                       Else
''                       u = Cod_Grfam(j)
''                       If Hijos(u) > 0 Then
'                        'I--- ABV 25/02/2005 ---
'                        'If Coinbe(j) = "T" Then
'                        '    Porben(j) = vlValor     '50
'                        'Else
'                        '    'HQR 16-06-2004
'                        '    'Porben(j) = 36
'                        '    If Coinbe(j) = "P" Then
'                        '        Porben(j) = 36
'                        '    Else
'                        '        Porben(j) = 0
'                        '        cont_esposa = cont_esposa - 1
'                        '    End If
'                        '    'FIN HQR 16-06-2004
'                        'End If
''                       Porben(j) = vlValor
''                       If (Coinbe(j) = "N") Then
''                       cont_esposa = cont_esposa - 1
'                        'End If
'
'                        If (iDerCrecerCotizacion = "S") Then 'I--- ABV 14/07/2007
'                            Codcbe(j) = "S"
'                            If Hijos_Inv(u) > 0 Then Codcbe(j) = "N"
'                        Else
'                            Codcbe(j) = "N"
'                        End If
'                        'Else
'                        'I--- ABV 25/02/2005 ---
'                        'If Coinbe(j) = "T" Then
'                        '    Porben(j) = vlValor     '60
'                        'Else
'                        '    'HQR 16-06-2004
'                        '    'Porben(j) = 43
'                        '    If Coinbe(j) = "P" Then
'                        '        Porben(j) = 43
'                        '    Else
'                        '        Porben(j) = 0
'                        '        cont_esposa = cont_esposa - 1
'                        '    End If
'                        '    'FIN HQR 16-06-2004
'                        'End If
'
'                        Porben(j) = vlValor
'                        If (Coinbe(j) = "N") Then
'                            cont_esposa = cont_esposa - 1
'                        End If
'
'                        If (iDerCrecerCotizacion = "S") Then 'I--- ABV 14/07/2007
'                            Codcbe(j) = "S"
'                            If Hijos_Inv(u) > 0 Then Codcbe(j) = "N"
'                        Else
'                            Codcbe(j) = "N"
'                        End If
'                        'End If
'                    End If
'                Case 20, 21
'                    If sexo_cau = "M" Then
'                        If Sexobe(j) <> "F" Then
'                            vgError = 1000
'                            X = MsgBox("Error de código de sexo, el Sexo de la Cónyuge debe ser Femenino.", vbCritical)
'                            Exit Function
'                        End If
'                    Else
'                            If Sexobe(j) <> "M" Then
'                                X = MsgBox("Error de codigo de sexo, el Sexo del Cónyuge debe ser Masculino.", vbCritical)
'                                Exit Function
'                            End If
'                    End If
'
'                    u = Cod_Grfam(j)
'                    If Hijos(u) = 0 And Ncorbe(j) = 21 Then
'                        If Hijos_SinDerechoPension(u) = 0 Then
'                            vgError = 1000
'                            X = MsgBox("Error en código de relación 'Madre Con Hijos con Dº Pensión, no tiene Hijos.", vbCritical)
'                            Exit Function
'                        Else
'                            Ncorbe(j) = 20
'                            ostBeneficiarios(j).Cod_Par = 20
'                        End If
'                    End If
'
'                    'I--- ABV 25/02/2005 ---
'                    'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
'                    If (fgObtenerPorcentaje(CStr(Ncorbe(j)), Coinbe(j), Sexobe(j), iFechaIniVig) = True) Then
'                        vlValor = vgValorPorcentaje
'                    Else
'                        vgError = 1000
'                        MsgBox "No existe valor de Porcentaje para el parámetro de :" _
'                        & Ncorbe(j) & " - " & Coinbe(j) & " - " & Sexobe(j) & " - " & iFechaIniVig & "."
'                        Exit Function
'                    End If
'                    'F--- ABV 25/02/2005 ---
'                    If (vlValor < 0) Then
'                        vgError = 1000
'                        MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
'                        Exit Function
'                    Else
'                        Porben(j) = vlValor / cont_mhn_tot
'                    End If
'
'                    If (iDerCrecerCotizacion = "S") Then 'I--- ABV 14/07/2007
'                        If Hijos(u) > 0 Then
'                            Codcbe(j) = "S"
'                        Else
'                            Codcbe(j) = "N"
'                        End If
'                        If Hijos_Inv(u) > 0 Then
'                            Codcbe(j) = "N"
'                        End If
'                    Else
'                        Codcbe(j) = "N"
'                    End If
'
'                    If Hijos(u) > 0 And Ncorbe(j) = 20 Then
'                        vgError = 1000
'                        X = MsgBox("Error en código de relación 'Madre Sin Hijos con Dº Pensión', tiene Hijos.", vbCritical)
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
'                            If (Coinbe(j) = "P" Or Coinbe(j) = "T") And edad_mes_ben > L24 Then
'                                'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
'                                If (fgObtenerPorcentaje(CStr(Ncorbe(j)), Coinbe(j), Sexobe(j), iFechaIniVig) = True) Then
'                                    vlValor = vgValorPorcentaje
'                                Else
'                                    vgError = 1000
'                                    MsgBox "No existe valor de Porcentaje para el parámetro de :" _
'                                    & Ncorbe(j) & " - " & Coinbe(j) & " - " & Sexobe(j) & " - " & iFechaIniVig & "."
'                                    Exit Function
'                                End If
'
'                                If (vlValor < 0) Then
'                                    vgError = 1000
'                                    MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
'                                    'X = MsgBox("Error: Código de cónyuge no esta incorporado", vbCritical)
'                                    Exit Function
'                                Else
'                                    Porben(j) = vlValor
'                                End If
'                            Else
'                                If (fgObtenerPorcentaje(CStr(Ncorbe(j)), "N", Sexobe(j), iFechaIniVig) = True) Then
'                                    'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
'                                    vlValor = vgValorPorcentaje
'                                Else
'                                    vgError = 1000
'                                    MsgBox "No existe valor de Porcentaje para el parámetro de :" _
'                                    & Ncorbe(j) & " - N - " & Sexobe(j) & " - " & iFechaIniVig & "."
'                                    Exit Function
'                                End If
'                                If (vlValor < 0) Then
'                                    vgError = 1000
'                                    'X = MsgBox("Error: Código de cónyuge no esta incorporado", vbCritical)
'                                    MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
'                                    Exit Function
'                                Else
'                                    Porben(j) = vlValor
'                                End If
'                            End If
'                        End If
'                        'If cont_esposa = 0 Or cont_mhn > 0 Then
'                    Else
''                       vgError = 1000
''                       X = MsgBox("Error: Los códigos de beneficiarios de hijos estan mal ingresados", vbCritical)
''                       Exit Function
''                       End If
'                        Q = Cod_Grfam(j)
'                        Codcbe(j) = "N"
''                       vlValorHijo = 0
'    '                    If cont_esp_Tot_GF(Q) > 0 Or cont_mhn_Tot_GF(Q) > 0 Then
'    '                        If Hijos_SinConyugeMadre(Q) = 0 Then
'    '                            vgError = 1000
'    '                            X = MsgBox("Error: Los códigos de beneficiarios de hijos estan mal ingresados", vbCritical)
'    '                            Exit Function
'    '                        End If
'    '                    End If
'                        If cont_esposa = 0 And cont_mhn(Q) = 0 Then
'                            If (Coinbe(j) = "P" Or Coinbe(j) = "T") And edad_mes_ben > L24 Then
'''                            'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
'                                If (fgObtenerPorcentaje(CStr(Ncorbe(j)), Coinbe(j), Sexobe(j), iFechaIniVig) = True) Then
'                                    vlValor = vgValorPorcentaje
'                                Else
'                                    vgError = 1000
'                                    MsgBox "No existe valor de Porcentaje para el parámetro de :" _
'                                    & Ncorbe(j) & " - " & Coinbe(j) & " - " & Sexobe(j) & " - " & iFechaIniVig & "."
'                                    Exit Function
'                                End If
'''
'                                If (vlValor < 0) Then
'                                    vgError = 1000
'    ''                                'X = MsgBox("Error: Código de cónyuge no esta incorporado", vbCritical)
'                                    MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
'                                    Exit Function
'                                Else
'                                    Porben(j) = vlValor
'                                End If
'                            Else
'                                If (fgObtenerPorcentaje(CStr(Ncorbe(j)), "N", Sexobe(j), iFechaIniVig) = True) Then
'    ''                                'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
'                                    vlValor = vgValorPorcentaje
'                                    vlValorHijo = vlValor
'                                Else
'                                    vgError = 1000
'                                    MsgBox "No existe valor de Porcentaje para el parámetro de :" _
'                                    & Ncorbe(j) & " - N - " & Sexobe(j) & " - " & iFechaIniVig & "."
'                                    Exit Function
'                                End If
'                                If (vlValor < 0) Then
'                                    vgError = 1000
'                                    MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
'    ''                                'X = MsgBox("Error: Código de cónyuge no esta incorporado", vbCritical)
'                                    Exit Function
'                                Else
'                                    Porben(j) = vlValor
'                                End If
'                            End If
'
'                            If (fgObtenerPorcentaje("10", "N", "F", iFechaIniVig) = True) Then
'    ''                            'vlValor = fgValorPorcentaje(1, j, 11)
'                                vlValor = vgValorPorcentaje
'                            Else
'                                vgError = 1000
'                                MsgBox "No existe valor de Porcentaje para el parámetro de :" _
'                                & "10 - N - F  - " & iFechaIniVig & "."
'                                Exit Function
'                            End If
'                            If (vlValor < 0) Then
'                                vgError = 1000
'                                'X = MsgBox("Error, el porcentaje para la 'Cónyuge Con Hijos' no se encuentra.", vbCritical)
'                                MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & "11" & ".", vbCritical, "Error de Datos"
'                                Exit Function
'                            Else
'                                v_hijo = vlValor
'                                If Coinbe(j) = "N" And edad_mes_ben <= L24 Then
'                                    Porben(j) = v_hijo / cont_hijo + vlValorHijo
'                                Else
'                                    If Coinbe(j) = "T" Or Coinbe(j) = "P" Then
'                                        Porben(j) = v_hijo / cont_hijo + vlValorHijo
'                                    End If
'                                End If
'                            End If
'
'                        Else
'                            vgError = 1000
'                            X = MsgBox("Error: Los códigos de beneficiarios de hijos estan mal ingresados", vbCritical)
'                            Exit Function
'                        End If
'                    End If
'                Case 41, 42
'                    'I--- ABV 25/02/2005 ---
'                    If (fgObtenerPorcentaje(CStr(Ncorbe(j)), Coinbe(j), Sexobe(j), iFechaIniVig) = True) Then
'                        'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
'                        vlValor = vgValorPorcentaje
'                    Else
'                        vgError = 1000
'                        MsgBox "No existe valor de Porcentaje para el parámetro de :" _
'                        & Ncorbe(j) & " - " & Coinbe(j) & " - " & Sexobe(j) & " - " & iFechaIniVig & "."
'                        Exit Function
'                    End If
'                    'F--- ABV 25/02/2005 ---
'
'                    If (vlValor < 0) Then
'                        vgError = 1000
'                        MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
'                        Exit Function
'                    Else
'                        Codcbe(j) = "N"
'                        Porben(j) = vlValor
'                    End If
'            End Select
'        End If
'    Next j
'
'    For k = 1 To vlNum_Ben '60
'        If derpen(k) <> 10 Then
'            Select Case Ncorbe(k)
'                Case 11
'                    Q = Cod_Grfam(k)
'                    'If codcbe(k) = "S" Then DAJ 10/08/2007
'                        Fec_NacHM(k) = Hijo_Menor_Ant(Q)  '******Agruegué 05/11/2000 HILDA
'                    'End If DAJ 10/08/2007
'                Case 21
'                    Q = Cod_Grfam(k)
'                    'If codcbe(k) = "S" Then DAJ 10/08/2007
'                        Fec_NacHM(k) = Hijo_Menor_Ant(Q)  '*******Agruegué 05/11/2000 HILDA
'                    'End If DAJ 10/08/2007
'            End Select
'        End If
'    Next k
'
'    vlSumarTotalPorcentajePension = 0
'
'    For j = 1 To (vlContBen - 1)
'        'Guardar el Valor del Porcentaje Calculado
'        'If IsNumeric(Porben(j)) Then
'        '    grd_beneficiarios.Text = Format(Porben(j), "##0.00")
'        'Else
'        '    grd_beneficiarios.Text = Format("0", "0.00")
'        'End If
'        If IsNumeric(Porben(j)) Then
'            'I--- ABV 22/06/2005 ---
'            'ostBeneficiarios(j).Prc_Pension = Format(Porben(j), "#0.00")
'            If (ostBeneficiarios(j).Cod_Par = "99" Or ostBeneficiarios(j).Cod_Par = "0") Then
'                If (derpen(j) <> 10) Then
'                    ostBeneficiarios(j).Prc_Pension = 100
'                    ostBeneficiarios(j).Prc_PensionLeg = 100 '*-+
'                Else
'                    ostBeneficiarios(j).Prc_Pension = 0
'                    ostBeneficiarios(j).Prc_PensionLeg = 0 '*-+
'                End If
'            Else
'                ostBeneficiarios(j).Prc_Pension = Format(Porben(j), "#0.00")
'                ostBeneficiarios(j).Prc_PensionLeg = Format(Porben(j), "#0.00") '*-+
'            End If
'            'F--- ABV 22/06/2005 ---
'        Else
'            ostBeneficiarios(j).Prc_Pension = 0
'            ostBeneficiarios(j).Prc_PensionLeg = 0 '*-+
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
'            'ostBeneficiarios(j).Fec_NacHM = CDate(Fec_NacHM(j))
'            ostBeneficiarios(j).Fec_NacHM = Format(CDate(Fec_NacHM(j)), "yyyymmdd")
'        Else
'            ostBeneficiarios(j).Fec_NacHM = ""
'        End If
'        'Fin
'
''I--- ABV 10/08/2007 ---
'        'Inicializar el Monto de Pensión a Cero
'        '--------------------------------------
'        'Actualizar el Monto de la Pensión
'        ostBeneficiarios(j).Mto_Pension = Format(0, "#0.00")
'
'        'Actualizar el Monto de la Pensión Garantizada
'        ostBeneficiarios(j).Mto_PensionGar = Format(0, "#0.00")
''F--- ABV 10/08/2007 ---
'
''I--- ABV 21/06/2006 ---
''        If (iCalcularPension = True) Then
''
''            vlPorcBenef = 0
''            vlPenBenef = 0
''            vlPenGarBenef = 0
''
''            'Verificar el Estado de Derecho a Pensión (Cod_EstPension)
''            If (ostBeneficiarios(j).Cod_EstPension <> "10") Then
''                vlPorcBenef = ostBeneficiarios(j).Prc_Pension
''
''                'Determinar el monto de la Pensión del Beneficiario
''                'Penben(i) = IIf(Msf_BMGrilla.TextMatrix(vlI, 17) = "", 0, Msf_BMGrilla.TextMatrix(vlI, 17))
''                vlPenBenef = iPensionRef * (vlPorcBenef / 100)
''
''                'Definir la Pensión del Garantizado
''                vlPenGarBenef = ostBeneficiarios(j).Mto_PensionGar
''
''                'I--- ABV 20/04/2005 ---
''                vgPalabra = iTipoPension
''                vgI = InStr(1, cgPensionInvVejez, vgPalabra)
''                If (vgI <> 0) Then
''                    'Valida que sea un caso de Invalidez o Vejez
''                    vlPenGarBenef = vlPenBenef
''                End If
''
''                vgI = InStr(1, cgPensionSobOrigen, vgPalabra)
''                If (vgI <> 0) Then
''                    'Valida que sea un caso de Sobrevivencia por Origen
''                    'Se supone que la Cónyuge o Madre es la única que Garantiza su pensión
''                    'vlNumero = InStr(Cmb_BMPar.Text, "-")
''                    vgPalabraAux = ostBeneficiarios(j).Cod_Par
''
''                    vgX = InStr(1, cgParConyugeMadre, vgPalabraAux)
''                    If (vgX <> 0) Then
''                        vlPenGarBenef = vlPenBenef
''                    'Else
''                    '    Txt_BMMtoPensionGar.Enabled = True
''                    End If
''                End If
''
''                vgI = InStr(1, cgPensionSobTransf, vgPalabra)
''                If (vgI <> 0) Then
''                    ''Valida que sea un caso de Invalidez o Vejez
''                    'Txt_BMMtoPensionGar.Enabled = True
''                End If
''
''                'F--- ABV 19/04/2005 ---
''            End If
''
''            'Actualizar el Monto de la Pensión
''            ostBeneficiarios(j).Mto_Pension = Format(vlPenBenef, "#0.00")
''
''            'Actualizar el Monto de la Pensión Garantizada
''            ostBeneficiarios(j).Mto_PensionGar = Format(vlPenGarBenef, "#0.00")
''
''        End If
'
'        If (ostBeneficiarios(j).Cod_DerPen <> "10") And (ostBeneficiarios(j).Cod_Par <> "99" And ostBeneficiarios(j).Cod_Par <> "0") Then
'            vlSumarTotalPorcentajePension = vlSumarTotalPorcentajePension + ostBeneficiarios(j).Prc_Pension
'        End If
''F--- ABV 21/06/2006 ---
'    Next j
'
''Validar si los Porcentajes de Pensión Suman más del 100%
'    If (vlSumarTotalPorcentajePension > 100) Then
'        vlSumaDef = 0
'        For j = 1 To (vlContBen - 1)
'            If (ostBeneficiarios(j).Cod_DerPen <> "10") And (ostBeneficiarios(j).Cod_Par <> "99" And ostBeneficiarios(j).Cod_Par <> "0") Then
'                vlPorcentajeRecal = (ostBeneficiarios(j).Prc_Pension / vlSumarTotalPorcentajePension) * 100
'                ostBeneficiarios(j).Prc_Pension = Format(vlPorcentajeRecal, "#0.00")
'                vlSumaDef = vlSumaDef + ostBeneficiarios(j).Prc_Pension
'            End If
'        Next j
'        vlDif = Format(100 - vlSumaDef, "#0.00")
'        If (vlDif < 0) Or (vlDif > 0) Then
'            'Asignar la diferencia de Porcentaje al Primer Beneficiario con Derecho distinto del Causante
'            For j = 1 To (vlContBen - 1)
'                If (ostBeneficiarios(j).Cod_DerPen <> "10") And (ostBeneficiarios(j).Cod_Par <> "99" And ostBeneficiarios(j).Cod_Par <> "0") Then
'                    vlPorcentajeRecal = ostBeneficiarios(j).Prc_Pension + vlDif
'                    ostBeneficiarios(j).Prc_Pension = vlPorcentajeRecal
'                    Exit For
'                End If
'            Next j
'        End If
'    End If
'
''*-+
''I--- ABV 04/08/2007 ---
''Recalcular Porcentajes si se trata de un Caso de Invalidez - Con Cobertura
'    If (iTipoPension = "06" Or iTipoPension = "07") And (iCobCobertura = "S") Then
'        If (fgObtenerPorcCobertura(iTipoPension, iFechaIniVig, vlRemuneracion, vlPrcCobertura) = True) Then
'            'Determinar la Remuneración Promedio para el Causante
'            vlRemuneracionBase = vlRemuneracion * (vlPrcCobertura / 100)
'
'            'Determinar para cada Beneficiario el Nuevo Porcentaje
'            For j = 1 To (vlContBen - 1)
'                If (ostBeneficiarios(j).Cod_DerPen <> "10") Then
'                    If (ostBeneficiarios(j).Cod_Par = "99") Then
'                        vlRemuneracionProm = vlRemuneracion * (vlPrcCobertura / 100)
'                    Else
'                        vlRemuneracionProm = vlRemuneracion * (ostBeneficiarios(j).Prc_Pension / 100)
'                    End If
'                    vlPorcentajeRecal = (vlRemuneracionProm / vlRemuneracionBase) * 100
'                    ostBeneficiarios(j).Prc_Pension = Format(vlPorcentajeRecal, "#0.00")
'                End If
'            Next j
'        Else
'            vgError = 1000
'            MsgBox "No existen valores de Porcentaje de Cobertura para el Recalculo de Pensiones.", vbCritical, "Inexistencia de Datos"
'            Exit Function
'        End If
'    End If
''F--- ABV 04/08/2007 ---
''*-+
'End If
'
''Recalcular los Montos de Pensión
'If (iCalcularPension = True) Then
'    For j = 1 To (vlNum_Ben)
'
'        vlPorcBenef = 0
'        vlPenBenef = 0
'        vlPenGarBenef = 0
'
'        'Verificar el Estado de Derecho a Pensión (Cod_EstPension)
''            If (ostBeneficiarios(j).Cod_EstPension <> "10") Then
'        If (ostBeneficiarios(j).Cod_DerPen <> "10") Then
'            vlPorcBenef = ostBeneficiarios(j).Prc_Pension
'
'            'Determinar el monto de la Pensión del Beneficiario
'            'Penben(i) = IIf(Msf_BMGrilla.TextMatrix(vlI, 17) = "", 0, Msf_BMGrilla.TextMatrix(vlI, 17))
'            vlPenBenef = iPensionRef * (vlPorcBenef / 100)
'
'            'Definir la Pensión del Garantizado
'            vlPenGarBenef = ostBeneficiarios(j).Mto_PensionGar
'
'            'I--- ABV 20/04/2005 ---
'            vgPalabra = iTipoPension
'            vgI = InStr(1, cgPensionInvVejez, vgPalabra)
'            If (vgI <> 0) Then
'                'Valida que sea un caso de Invalidez o Vejez
'                vlPenGarBenef = vlPenBenef
'            End If
'
'            vgI = InStr(1, cgPensionSobOrigen, vgPalabra)
'            If (vgI <> 0) Then
'                'Valida que sea un caso de Sobrevivencia por Origen
'                'Se supone que la Cónyuge o Madre es la única que Garantiza su pensión
'                'vlNumero = InStr(Cmb_BMPar.Text, "-")
'                vgPalabraAux = ostBeneficiarios(j).Cod_Par
'
'                vgX = InStr(1, cgParConyugeMadre, vgPalabraAux)
'                If (vgX <> 0) Then
'                    vlPenGarBenef = vlPenBenef
'                'Else
'                '    Txt_BMMtoPensionGar.Enabled = True
'                End If
'            End If
'
'            vgI = InStr(1, cgPensionSobTransf, vgPalabra)
'            If (vgI <> 0) Then
'                ''Valida que sea un caso de Invalidez o Vejez
'                'Txt_BMMtoPensionGar.Enabled = True
'            End If
'
'            'F--- ABV 19/04/2005 ---
'        End If
'
'        'Actualizar el Monto de la Pensión
'        ostBeneficiarios(j).Mto_Pension = Format(vlPenBenef, "#0.00")
'
'        'Actualizar el Monto de la Pensión Garantizada
'        If (iPerGar > 0) Then
'            ostBeneficiarios(j).Mto_PensionGar = Format(vlPenGarBenef, "#0.00")
'            ostBeneficiarios(j).Prc_PensionGar = vlPorcBenef
'        Else
'            ostBeneficiarios(j).Mto_PensionGar = 0
'            ostBeneficiarios(j).Prc_PensionGar = 0
'        End If
'
'    Next j
'End If
'
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
'        vgError = 1000
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
End Function

Function fgObtenerPorcentaje(iParentesco As String, iInvalidez As String, iSexo As String, iFecha As String, sInvS As String) As Boolean
'Función : Permite validar la existencia del valor del Porcentaje de Pensión a buscar
'Parámetros de Entrada:
'       - iParentesco => Código del Parentesco del Beneficiario
'       - iInvalidez  => Código de la Situación de Invalidez del Beneficiario
'       - iSexo       => Código del Sexo del Beneficiario
'       - iFecha      => Fecha con la cual se compara la Vigencia del Porcentaje (Vigencia Póliza)
'Parámetros de Salida:
'       - Retorna un Falso o True de acuerdo a su existencia
'Variables de Salida:
'       - vgValorPorcentaje => Permite guardar el Porcentaje buscado
Dim Tb_Por As ADODB.Recordset
Dim Sql    As String
Dim Tb_cob As ADODB.Recordset
Dim val_cob As Integer

    fgObtenerPorcentaje = False
    
    Sql = "select prc_pension as valor_porcentaje "
    Sql = Sql & "from MA_TVAL_PORPAR where "
    Sql = Sql & "Cod_par = '" & iParentesco & "' AND "
    Sql = Sql & "Cod_sitinv = '" & iInvalidez & "' AND "
    Sql = Sql & "Cod_sexo = '" & iSexo & "' AND "
    Sql = Sql & "fec_inivigpor <= '" & iFecha & "' AND "
    Sql = Sql & "fec_tervigpor >= '" & iFecha & "' "
    
    Set Tb_Por = vgConexionBD.Execute(Sql)
    If Not Tb_Por.EOF Then
        
        If Not IsNull(Tb_Por!valor_porcentaje) Then
            
            Sql = "select cod_cobercon from pp_tmae_poliza where num_poliza=" & vgNum_pol & " and num_endoso=1"
            Set Tb_cob = vgConexionBD.Execute(Sql)
            If Not Tb_cob.EOF Then
                val_cob = Tb_cob!Cod_CoberCon
            End If
            If val_cob <> 0 Then
                vgValorPorcentaje = val_cob
            Else
                vgValorPorcentaje = Tb_Por!valor_porcentaje
            End If
            
            
'            Select Case sInvS
'                Case "N"
'                    vgValorPorcentaje = Tb_Por!Valor_Porcentaje
'                Case "P"
'                    vgValorPorcentaje = Tb_Por!Valor_Porcentaje / 0.5
'                Case "T"
'                    vgValorPorcentaje = Tb_Por!Valor_Porcentaje / 0.7
'            End Select
            
            
            fgObtenerPorcentaje = True
        End If
    End If
    Tb_Por.Close
    
End Function

'I---- ABV 17/11/2005 ---
'***********************************************************
'Inicio : Funciones agregadas para el manejo de los Endosos
'***********************************************************
Function fgCrearMortalidadDinamica_DaniMensual(iNavig, iNmvig, iNdvig, _
    iab, imb, idb, iSexo, iInval, iCorrelativo, iFinTab, iAñoBase, iFNac) As Boolean
'Función: permite generar la Tabla de Mortalidad Dinámica desde la descripción de una
'         Tabla Anual
'Parámetros de Entrada:
'   - iNavig = Año de Vigencia de la Tabla de Mortalidad
'   - iNmvig = Mes de Vigencia de la Tabla de Mortalidad
'   - iNdvig = Día de Vigencia de la Tabla de Mortalidad
'   - iab    = Año de Proceso
'   - imb    = Mes de Proceso
'   - idb    = Día de Proceso
'   - iSexo  = Sexo a generar
'   - iInval = Invalidez a generar
'   - iCorrelativo = Número de la Tabla de Mortalidad a Leer
'   - iFinTab = Valor que indica el termino de la Tabla de Mortalidad
'   - iAñoBase = Indica el Año Base desde el cual se genera la Tabla Dinámica
'Valor de Salida:
'   Retorna un valor True o False si pudo realizar la Actualización de la Matriz

'Dim tledad(0 To 90), qx(0 To 90), facmejor(0 To 90)
'Dim Lxm(1 To 2, 1 To 3, 1 To 1332) As Double
Dim Lxm(1 To 1332) As Double

Dim QxModif(0 To 120) As Double
Dim FacMejor(0 To 120) As Double, TlEdad(0 To 120) As Double
Dim Qx(0 To 120) As Double
Dim AñoBase As Long
Dim AñoProceso As Integer, MesProceso As Integer, DiaProceso As Integer
Dim AñoNac     As Integer, MesNac     As Integer, DiaNac     As Integer
Dim QxMensModif As Double, Parte1 As Double, Parte2 As Double
Dim i As Long, j As Long, k As Long
Dim Tb2 As ADODB.Recordset
Dim FProceso As String
Dim Edad As Long, edaca As Long
Dim difdia As Integer

'On Error GoTo Err_Mortal

    fgCrearMortalidadDinamica_DaniMensual = False
    
    AñoBase = iAñoBase
    
    FProceso = Format(iab, "0000") & Format(imb, "00") & Format(idb, "00")
    AñoProceso = iab
    MesProceso = imb
    DiaProceso = idb
    Fechap = AñoProceso * 12 + MesProceso
    
    AñoNac = Mid(iFNac, 1, 4)
    MesNac = Mid(iFNac, 5, 2)
    DiaNac = Mid(iFNac, 7, 2)
    Fechan = AñoNac * 12 + MesNac
    
    Edad = Fechap - Fechan
    
    difdia = idb - DiaNac
    If difdia > 15 Then Edad = Edad + 1
    If Edad <= 240 Then Edad = 240
    If Edad > (110 * 12) Then
        vgError = 1023
        Exit Function
    End If
    edaca = Fix(Edad / 12)
    
    'Lectura de tabla de mortalidad
    vgSql = "SELECT num_edad AS edad, mto_qx AS qx, prc_factor AS factor "
    vgSql = vgSql & "FROM ma_tval_mordet "
    vgSql = vgSql & "WHERE num_correlativo = " & vgs_Nro & " "
    vgSql = vgSql & "ORDER BY num_edad "
    Set Tb2 = vgConexionBD.Execute(vgSql)
    If Not (Tb2.EOF) Then
        Do While Not Tb2.EOF
            k = Tb2!Edad
            TlEdad(k) = Tb2!Edad
            Qx(k) = Tb2!Qx
            FacMejor(k) = Tb2!Factor
            Tb2.MoveNext
        Loop
    Else
        vgError = 1061
        Tb2.Close
        Exit Function
    End If
    Tb2.Close
    
    j = -1
    For i = edaca To 110 '- edaca)
        j = j + 1
        QxModif(i) = Qx(i) * (1 - FacMejor(i)) ^ (j + (AñoProceso - AñoBase))
    Next i
    
    QxMensModif = 0
    For i = edaca To 110 '- edaca)
        For j = 1 To 12
            k = (i * 12) + j - 1
            If k < Edad Then
            Else
                If k = Edad Then
                    Lxm(k) = 100000
                Else
                    Lxm(k) = Lxm(k - 1) - (Lxm(k - 1) * QxMensModif)
                End If
                
                Parte1 = ((1 / 12) * QxModif(i))
                Parte2 = (k / 12 - Fix(k / 12))
                If ((1 - Parte2 * QxModif(i)) = 0) Then
                    QxMensModif = 0
                Else
                    QxMensModif = Parte1 / (1 - Parte2 * QxModif(i))
                End If
                
                If k > (110 * 12) Then Exit For
                
                Lx(iSexo, iInval, k) = Lxm(k)
            
'                'Borra - Daniela
'                vgQuery = "INSERT into DANI (poliza,agno ,edad,qx ,lx,sexo) values( "
'                vgQuery = vgQuery & "'1', "
'                vgQuery = vgQuery & Str(Format(AñoProceso, "#0")) & ", "
'                vgQuery = vgQuery & Str(Format(k, "#000")) & ", "
'                vgQuery = vgQuery & Str(Format(QxMensModif, "#0.0000000000000")) & ", "
'                vgQuery = vgQuery & Str(Format(Lxm(k), "#000.000000000")) & ", "
'                vgQuery = vgQuery & Str(Format(iSexo, "#0")) & ") "
'                vgConexionBD.Execute (vgQuery)
'
'                'vgQuery = "update DANI set  "
'                'vgQuery = vgQuery & "agno = " & Str(Format(iNavig, "#0")) & ", "
'                'vgQuery = vgQuery & "edad = " & Str(Format(k, "#000")) & ", "
'                'vgQuery = vgQuery & "qx = " & Str(Format(QxMensModif, "#0.0000000")) & ", "
'                'vgQuery = vgQuery & "lx = " & Str(Format(Lxm(iSexo, iInval, k), "#000.000")) & " "
'                'vgConexionBD.Execute (vgQuery)
'                'End If
            End If
        Next j
    Next i
    
    fgCrearMortalidadDinamica_DaniMensual = True
    
Exit Function   'Buscar otra Póliza a calcular
Err_Mortal:
    'Screen.MousePointer = 0
    Select Case Err
        Case Else
        'ProgressBar.Value = 0
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function


Function fgBuscarMortalidadNormativa(iNavig, iNmvig, iNdvig, iNap, iNmp, iNdp, iSexoCau, iFechaNacCau) As Boolean
Dim vlResPregunta As Boolean

    fgBuscarMortalidadNormativa = False
    
    '1. Leer Tabla de Mortalidad de Rtas. Vitalicias Mujer
    vgBuscarMortalVit_F = ""
    If (vgUtilizarNormativa = "S") Then
        If (vgFechaIniMortalVit_F = "") And (vgFechaFinMortalVit_F = "") Then
            vgBuscarMortalVit_F = "S"
        Else
            If (vgIndicadorTipoMovimiento_F = "E") Or (vgIndicadorTipoMovimiento_F <> "E" And iSexoCau = "M") Then
                vgBuscarMortalVit_F = "N"
            Else
                vgBuscarMortalVit_F = "S"
            End If
        End If
    Else
        If (vgFechaIniMortalVit_F <> "") And (vgFechaFinMortalVit_F <> "") Then
            If ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) >= (vgFechaIniMortalVit_F)) And _
            ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) <= (vgFechaFinMortalVit_F)) Then
                If (vgIndicadorTipoMovimiento_F = "E") Or (vgIndicadorTipoMovimiento_F <> "E" And iSexoCau = "M") Then
                    vgBuscarMortalVit_F = "N"
                Else
                    vgBuscarMortalVit_F = "S"
                End If
            Else
                vgBuscarMortalVit_F = "S"
            End If
        Else
            vgBuscarMortalVit_F = "S"
        End If
    End If
    
    If (vgBuscarMortalVit_F = "S") Then
        'For h = 1 To 2  '1=Causante '2=Beneficiario
        h = 1
            vgs_Sexo = ""
            vgs_Tipo = ""
            vgs_Nro = ""
            'For i = 1 To 2
            '    If i = 1 Then vgs_Sexo = "M"    ' hombre  (M)asculino
            '    If i = 2 Then vgs_Sexo = "F"    ' mujer   (F)emenino
                i = 2
                vgs_Sexo = "F"
                
                'For j = 1 To 3
                j = 2
                    'If j = 1 Then vgs_Tipo = "MIT"    '1  = Invalido Total
                    If j = 2 Then vgs_Tipo = "RV"    '2  = No Invalido
                    'If h = 2 And j = 2 Then vgs_Tipo = "B" ' Tabla de beneficiarios no invalidos
                    'If j = 3 Then vgs_Tipo = "MIP"    '3  = invalido Parcial
                
                vgSw = False
                
                'Buscar Número Correlativo desde la Estructura
                For vgI = 1 To vgNumeroTotalTablas
                                      
                    'Sql = " SELECT * "
                    'Sql = Sql & " from PR_TVAL_MORTAL "
                    'Sql = Sql & " where cod_sexo = '" & vgs_Sexo & "' and "
                    'Sql = Sql & " cod_tiptabmor ='" & vgs_Tipo & "' and "
                    'Sql = Sql & " fec_ini <= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "# AND "
                    'Sql = Sql & " fec_ter >= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "#"
                    'Set tb = vgConexionBD.Execute(Sql)
                    'If Not (tb.EOF) Then
                    
                    vlResPregunta = False
                    If (vgUtilizarNormativa = "S") Then
                        If (egTablaMortal(vgI).Correlativo = vgMortalVit_F_Normativa) Then
                            vlResPregunta = True
                        End If
                    Else
                        If (egTablaMortal(vgI).Sexo = vgs_Sexo) And _
                        (egTablaMortal(vgI).TipoTabla = vgs_Tipo) And _
                        (egTablaMortal(vgI).Nombre = vgPalabra_MortalVit_F) And _
                        (egTablaMortal(vgI).FechaIni <= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) And _
                        (egTablaMortal(vgI).FechaFin >= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) Then
                            vlResPregunta = True
                        End If
                    End If
                    
                    If (vlResPregunta = True) Then
                    
                        vgs_Error = 0
                        vgs_Nro = egTablaMortal(vgI).Correlativo
                        vgFechaIniMortalVit_F = egTablaMortal(vgI).FechaIni
                        vgFechaFinMortalVit_F = egTablaMortal(vgI).FechaFin
                        vgIndicadorTipoMovimiento_F = egTablaMortal(vgI).TipoMovimiento
                        vgDinamicaAñoBase_F = egTablaMortal(vgI).AñoBase
                        
                        'Limpiar columna de Datos
                        For vgX = 1 To Fintab
                            Lx(i, j, vgX) = 0
                        Next vgX
                        
                        'vgs_Nro = tb!num_correlativo
                        'vgFechaInicioMortal = Format(tb!fec_ini, "yyyymmdd")
                        'vgFechaFinMortal = Format(tb!fec_ter, "yyyymmdd")
                        ''vgFechaInicioMortal = DateSerial(Year(tb!fec_ini), Month(tb!fec_ini), Day(tb!fec_ini))
                        ''vgFechaFinMortal = DateSerial(Year(tb!fec_ter), Month(tb!fec_ter), Day(tb!fec_ter))
                        If vgs_Nro <> 0 Then
                            If (vgIndicadorTipoMovimiento_F <> "D") Then
                                Sql = "Select num_edad AS edad,mto_lx from MA_TVAL_MORDET "
                                Sql = Sql & " where num_correlativo = " & vgs_Nro
                                Sql = Sql & " order by num_edad "
                                Set Tb2 = vgConexionBD.Execute(Sql)
                                If Not (Tb2.EOF) Then
                                    vgSw = True
                                    'tb2.MoveFirst
                                    'k = 1
                                    Do While Not Tb2.EOF
                                        k = Tb2!Edad
                                        'If h = 1 Then   'Causante
                                            Lx(i, j, k) = Tb2!mto_lx
                                        'Else    'Beneficiario
                                        '    ly(i, j, k) = tb2!mto_lx
                                        'End If
                                        'k = k + 1
                                        Tb2.MoveNext
                                    Loop
                                Else
                                    vgError = 1061
                                    Exit Function
                                End If
                                Tb2.Close
                                
                                Exit For
                            Else
                                'Obtener la Tabla Temporal Dinánica
                                'If (fgCrearMortalidadDinamica(iNavig, iNmvig, iNdvig, _
                                iNap, iNmp, iNdp, i, j, vgs_Nro, Fintab, vgDinamicaAñoBase_F, iFechaNacCau) = False) Then
                                    
                                If (fgCrearMortalidadDinamica_DaniMensual(iNavig, iNmvig, iNdvig, _
                                iNap, iNmp, iNdp, i, j, vgs_Nro, Fintab, vgDinamicaAñoBase_F, iFechaNacCau) = False) Then
                                    vgError = 1061
                                    Exit Function
                                Else
                                    vgSw = True
                                End If
                            End If
                        Else
                            vgError = 1061
                            Exit Function
                        End If
                    'Else
                    '    vgError = 1061
                    '    Exit Function
                    End If
                    'tb.Close
                Next vgI
                If (vgSw = False) Then
                    vgError = 1061
                    Exit Function
                End If
                'Next j
            'Next i
        'Next h
    End If

    '2. Leer Tabla de Mortalidad de Inv. Totales Mujer
    vgBuscarMortalTot_F = ""
    If (vgUtilizarNormativa = "S") Then
        If (vgFechaIniMortalTot_F = "") And (vgFechaFinMortalTot_F = "") Then
            vgBuscarMortalTot_F = "S"
        End If
    Else
        If (vgFechaIniMortalTot_F <> "") And (vgFechaFinMortalTot_F <> "") Then
            If ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) >= (vgFechaIniMortalTot_F)) And _
            ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) <= (vgFechaFinMortalTot_F)) Then
                vgBuscarMortalTot_F = "N"
            Else
                vgBuscarMortalTot_F = "S"
            End If
        Else
            vgBuscarMortalTot_F = "S"
        End If
    End If
    
    If (vgBuscarMortalTot_F = "S") Then
        For h = 1 To 2  '1=Causante '2=Beneficiario
            vgs_Sexo = ""
            vgs_Tipo = ""
            vgs_Nro = ""
            'For i = 1 To 2
            '    If i = 1 Then vgs_Sexo = "M"    ' hombre  (M)asculino
            '    If i = 2 Then vgs_Sexo = "F"    ' mujer   (F)emenino
                i = 2
                vgs_Sexo = "F"
                
                'For j = 1 To 3
                j = 1
                    If j = 1 Then vgs_Tipo = "MIT"    '1  = Invalido Total
                    'If j = 2 Then vgs_Tipo = "RV"    '2  = No Invalido
                    'If h = 2 And j = 2 Then vgs_Tipo = "B" ' Tabla de beneficiarios no invalidos
                    'If j = 3 Then vgs_Tipo = "MIP"    '3  = invalido Parcial
                
                vgSw = False
                
                'Buscar Número Correlativo desde la Estructura
                For vgI = 1 To vgNumeroTotalTablas
                                      
                    'Sql = " SELECT * "
                    'Sql = Sql & " from PR_TVAL_MORTAL "
                    'Sql = Sql & " where cod_sexo = '" & vgs_Sexo & "' and "
                    'Sql = Sql & " cod_tiptabmor ='" & vgs_Tipo & "' and "
                    'Sql = Sql & " fec_ini <= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "# AND "
                    'Sql = Sql & " fec_ter >= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "#"
                    'Set tb = vgConexionBD.Execute(Sql)
                    'If Not (tb.EOF) Then
                    
                    vlResPregunta = False
                    If (vgUtilizarNormativa = "S") Then
                        If (egTablaMortal(vgI).Correlativo = vgMortalTot_F_Normativa) Then
                            vlResPregunta = True
                        End If
                    Else
                        If (egTablaMortal(vgI).Sexo = vgs_Sexo) And _
                        (egTablaMortal(vgI).TipoTabla = vgs_Tipo) And _
                        (egTablaMortal(vgI).Nombre = vgPalabra_MortalTot_F) And _
                        (egTablaMortal(vgI).FechaIni <= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) And _
                        (egTablaMortal(vgI).FechaFin >= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) Then
                            vlResPregunta = True
                        End If
                    End If
                    
                    If (vlResPregunta = True) Then
                        vgs_Error = 0
                        vgs_Nro = egTablaMortal(vgI).Correlativo
                        vgFechaIniMortalTot_F = egTablaMortal(vgI).FechaIni
                        vgFechaFinMortalTot_F = egTablaMortal(vgI).FechaFin
                        
                        'Limpiar columna de Datos
                        If (h = 1) Then
                            For vgX = 1 To Fintab
                                Lx(i, j, vgX) = 0
                            Next vgX
                        Else
                            For vgX = 1 To Fintab
                                Ly(i, j, vgX) = 0
                            Next vgX
                        End If
                        
                        'vgs_Nro = tb!num_correlativo
                        'vgFechaInicioMortal = Format(tb!fec_ini, "yyyymmdd")
                        'vgFechaFinMortal = Format(tb!fec_ter, "yyyymmdd")
                        ''vgFechaInicioMortal = DateSerial(Year(tb!fec_ini), Month(tb!fec_ini), Day(tb!fec_ini))
                        ''vgFechaFinMortal = DateSerial(Year(tb!fec_ter), Month(tb!fec_ter), Day(tb!fec_ter))
                        If vgs_Nro <> 0 Then
                            Sql = "Select num_edad AS edad,mto_lx from MA_TVAL_MORDET "
                            Sql = Sql & " where num_correlativo = " & vgs_Nro
                            Sql = Sql & " order by num_edad "
                            Set Tb2 = vgConexionBD.Execute(Sql)
                            If Not (Tb2.EOF) Then
                                vgSw = True
                                'tb2.MoveFirst
                                'k = 1
                                Do While Not Tb2.EOF
                                    k = Tb2!Edad
                                    If h = 1 Then   'Causante
                                        Lx(i, j, k) = Tb2!mto_lx
                                    Else    'Beneficiario
                                        Ly(i, j, k) = Tb2!mto_lx
                                    End If
                                    'k = k + 1
                                    Tb2.MoveNext
                                Loop
                            Else
                                vgError = 1062
                                Exit Function
                            End If
                            Tb2.Close
                            
                            Exit For
                        Else
                            vgError = 1062
                            Exit Function
                        End If
                    'Else
                    '    vgError = 1062
                    '    Exit Function
                    End If
                    'tb.Close
                Next vgI
                If (vgSw = False) Then
                    vgError = 1062
                    Exit Function
                End If
                'Next j
            'Next i
        Next h
    End If

    '3. Leer Tabla de Mortalidad de Inv. Parciales Mujer
    vgBuscarMortalPar_F = ""
    If (vgUtilizarNormativa = "S") Then
        If (vgFechaIniMortalPar_F = "") And (vgFechaFinMortalPar_F = "") Then
            vgBuscarMortalPar_F = "S"
        End If
    Else
        If (vgFechaIniMortalPar_F <> "") And (vgFechaFinMortalPar_F <> "") Then
            If ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) >= (vgFechaIniMortalPar_F)) And _
            ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) <= (vgFechaFinMortalPar_F)) Then
                vgBuscarMortalPar_F = "N"
            Else
                vgBuscarMortalPar_F = "S"
            End If
        Else
            vgBuscarMortalPar_F = "S"
        End If
    End If
    
    If (vgBuscarMortalPar_F = "S") Then
        For h = 1 To 2  '1=Causante '2=Beneficiario
            vgs_Sexo = ""
            vgs_Tipo = ""
            vgs_Nro = ""
            'For i = 1 To 2
            '    If i = 1 Then vgs_Sexo = "M"    ' hombre  (M)asculino
            '    If i = 2 Then vgs_Sexo = "F"    ' mujer   (F)emenino
                i = 2
                vgs_Sexo = "F"
                
                'For j = 1 To 3
                j = 3
                    'If j = 1 Then vgs_Tipo = "MIT"    '1  = Invalido Total
                    'If j = 2 Then vgs_Tipo = "RV"    '2  = No Invalido
                    'If h = 2 And j = 2 Then vgs_Tipo = "B" ' Tabla de beneficiarios no invalidos
                    If j = 3 Then vgs_Tipo = "MIP"    '3  = invalido Parcial
                
                vgSw = False
                
                'Buscar Número Correlativo desde la Estructura
                For vgI = 1 To vgNumeroTotalTablas
                                      
                    'Sql = " SELECT * "
                    'Sql = Sql & " from PR_TVAL_MORTAL "
                    'Sql = Sql & " where cod_sexo = '" & vgs_Sexo & "' and "
                    'Sql = Sql & " cod_tiptabmor ='" & vgs_Tipo & "' and "
                    'Sql = Sql & " fec_ini <= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "# AND "
                    'Sql = Sql & " fec_ter >= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "#"
                    'Set tb = vgConexionBD.Execute(Sql)
                    'If Not (tb.EOF) Then
                    
                    vlResPregunta = False
                    If (vgUtilizarNormativa = "S") Then
                        If (egTablaMortal(vgI).Correlativo = vgMortalPar_F_Normativa) Then
                            vlResPregunta = True
                        End If
                    Else
                        If (egTablaMortal(vgI).Sexo = vgs_Sexo) And _
                        (egTablaMortal(vgI).TipoTabla = vgs_Tipo) And _
                        (egTablaMortal(vgI).Nombre = vgPalabra_MortalPar_F) And _
                        (egTablaMortal(vgI).FechaIni <= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) And _
                        (egTablaMortal(vgI).FechaFin >= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) Then
                            vlResPregunta = True
                        End If
                    End If
                    
                    If (vlResPregunta = True) Then
                        vgs_Error = 0
                        vgs_Nro = egTablaMortal(vgI).Correlativo
                        vgFechaIniMortalPar_F = egTablaMortal(vgI).FechaIni
                        vgFechaFinMortalPar_F = egTablaMortal(vgI).FechaFin
                        
                        'Limpiar columna de Datos
                        If (h = 1) Then
                            For vgX = 1 To Fintab
                                Lx(i, j, vgX) = 0
                            Next vgX
                        Else
                            For vgX = 1 To Fintab
                                Ly(i, j, vgX) = 0
                            Next vgX
                        End If
                        
                        'vgs_Nro = tb!num_correlativo
                        'vgFechaInicioMortal = Format(tb!fec_ini, "yyyymmdd")
                        'vgFechaFinMortal = Format(tb!fec_ter, "yyyymmdd")
                        ''vgFechaInicioMortal = DateSerial(Year(tb!fec_ini), Month(tb!fec_ini), Day(tb!fec_ini))
                        ''vgFechaFinMortal = DateSerial(Year(tb!fec_ter), Month(tb!fec_ter), Day(tb!fec_ter))
                        If vgs_Nro <> 0 Then
                            Sql = "Select num_edad AS edad,mto_lx from MA_TVAL_MORDET "
                            Sql = Sql & " where num_correlativo = " & vgs_Nro
                            Sql = Sql & " order by num_edad "
                            Set Tb2 = vgConexionBD.Execute(Sql)
                            If Not (Tb2.EOF) Then
                                vgSw = True
                                'tb2.MoveFirst
                                'k = 1
                                Do While Not Tb2.EOF
                                    k = Tb2!Edad
                                    If h = 1 Then   'Causante
                                        Lx(i, j, k) = Tb2!mto_lx
                                    Else    'Beneficiario
                                        Ly(i, j, k) = Tb2!mto_lx
                                    End If
                                    'k = k + 1
                                    Tb2.MoveNext
                                Loop
                            Else
                                vgError = 1063
                                Exit Function
                            End If
                            Tb2.Close
                            
                            Exit For
                        Else
                            vgError = 1063
                            Exit Function
                        End If
                    'Else
                    '    vgError = 1063
                    '    Exit Function
                    End If
                    'tb.Close
                Next vgI
                If (vgSw = False) Then
                    vgError = 1063
                    Exit Function
                End If
                'Next j
            'Next i
        Next h
    End If

    '4. Leer Tabla de Mortalidad de Beneficiarios Mujer
    vgBuscarMortalBen_F = ""
    If (vgUtilizarNormativa = "S") Then
        If (vgFechaIniMortalBen_F = "") And (vgFechaFinMortalBen_F = "") Then
            vgBuscarMortalBen_F = "S"
        End If
    Else
        If (vgFechaIniMortalBen_F <> "") And (vgFechaFinMortalBen_F <> "") Then
            If ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) >= (vgFechaIniMortalBen_F)) And _
            ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) <= (vgFechaFinMortalBen_F)) Then
                vgBuscarMortalBen_F = "N"
            Else
                vgBuscarMortalBen_F = "S"
            End If
        Else
            vgBuscarMortalBen_F = "S"
        End If
    End If
    
    If (vgBuscarMortalBen_F = "S") Then
        'For h = 1 To 2  '1=Causante '2=Beneficiario
        h = 2
            vgs_Sexo = ""
            vgs_Tipo = ""
            vgs_Nro = ""
            'For i = 1 To 2
            '    If i = 1 Then vgs_Sexo = "M"    ' hombre  (M)asculino
            '    If i = 2 Then vgs_Sexo = "F"    ' mujer   (F)emenino
                i = 2
                vgs_Sexo = "F"
                
                'For j = 1 To 3
                j = 2
                    'If j = 1 Then vgs_Tipo = "MIT"    '1  = Invalido Total
                    'If j = 2 Then vgs_Tipo = "RV"    '2  = No Invalido
                    If h = 2 And j = 2 Then vgs_Tipo = "B" ' Tabla de beneficiarios no invalidos
                    'If j = 3 Then vgs_Tipo = "MIP"    '3  = invalido Parcial
                
                vgSw = False
                
                'Buscar Número Correlativo desde la Estructura
                For vgI = 1 To vgNumeroTotalTablas
                                      
                    'Sql = " SELECT * "
                    'Sql = Sql & " from PR_TVAL_MORTAL "
                    'Sql = Sql & " where cod_sexo = '" & vgs_Sexo & "' and "
                    'Sql = Sql & " cod_tiptabmor ='" & vgs_Tipo & "' and "
                    'Sql = Sql & " fec_ini <= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "# AND "
                    'Sql = Sql & " fec_ter >= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "#"
                    'Set tb = vgConexionBD.Execute(Sql)
                    'If Not (tb.EOF) Then
                    
                    vlResPregunta = False
                    If (vgUtilizarNormativa = "S") Then
                        If (egTablaMortal(vgI).Correlativo = vgMortalBen_F_Normativa) Then
                            vlResPregunta = True
                        End If
                    Else
                        If (egTablaMortal(vgI).Sexo = vgs_Sexo) And _
                        (egTablaMortal(vgI).TipoTabla = vgs_Tipo) And _
                        (egTablaMortal(vgI).Nombre = vgPalabra_MortalBen_F) And _
                        (egTablaMortal(vgI).FechaIni <= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) And _
                        (egTablaMortal(vgI).FechaFin >= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) Then
                            vlResPregunta = True
                        End If
                    End If
                    
                    If (vlResPregunta = True) Then
                        vgs_Error = 0
                        vgs_Nro = egTablaMortal(vgI).Correlativo
                        vgFechaIniMortalBen_F = egTablaMortal(vgI).FechaIni
                        vgFechaFinMortalBen_F = egTablaMortal(vgI).FechaFin
                        
                        'Limpiar columna de Datos
                        For vgX = 1 To Fintab
                            Ly(i, j, vgX) = 0
                        Next vgX
                        
                        'vgs_Nro = tb!num_correlativo
                        'vgFechaInicioMortal = Format(tb!fec_ini, "yyyymmdd")
                        'vgFechaFinMortal = Format(tb!fec_ter, "yyyymmdd")
                        ''vgFechaInicioMortal = DateSerial(Year(tb!fec_ini), Month(tb!fec_ini), Day(tb!fec_ini))
                        ''vgFechaFinMortal = DateSerial(Year(tb!fec_ter), Month(tb!fec_ter), Day(tb!fec_ter))
                        If vgs_Nro <> 0 Then
                            Sql = "Select num_edad AS edad,mto_lx from MA_TVAL_MORDET "
                            Sql = Sql & " where num_correlativo = " & vgs_Nro
                            Sql = Sql & " order by num_edad "
                            Set Tb2 = vgConexionBD.Execute(Sql)
                            If Not (Tb2.EOF) Then
                                vgSw = True
                                'tb2.MoveFirst
                                'k = 1
                                Do While Not Tb2.EOF
                                    k = Tb2!Edad
                                    'If h = 1 Then   'Causante
                                    '    lx(i, j, k) = tb2!mto_lx
                                    'Else    'Beneficiario
                                        Ly(i, j, k) = Tb2!mto_lx
                                    'End If
                                    'k = k + 1
                                    Tb2.MoveNext
                                Loop
                            Else
                                vgError = 1064
                                Exit Function
                            End If
                            Tb2.Close
                            
                            Exit For
                        Else
                            vgError = 1064
                            Exit Function
                        End If
                    'Else
                    '    vgError = 1064
                    '    Exit Function
                    End If
                    'tb.Close
                Next vgI
                If (vgSw = False) Then
                    vgError = 1064
                    Exit Function
                End If
                'Next j
            'Next i
        'Next h
    End If

'--------------------------------------------------------------------
    '5. Leer Tabla de Mortalidad de Rtas. Vitalicias Hombre
    vgBuscarMortalVit_M = ""
    If (vgUtilizarNormativa = "S") Then
        If (vgFechaIniMortalVit_M = "") And (vgFechaFinMortalVit_M = "") Then
            vgBuscarMortalVit_M = "S"
        Else
            If (vgIndicadorTipoMovimiento_M = "E") Or (vgIndicadorTipoMovimiento_M <> "E" And iSexoCau = "F") Then
                vgBuscarMortalVit_M = "N"
            Else
                vgBuscarMortalVit_M = "S"
            End If
        End If
    Else
        If (vgFechaIniMortalVit_M <> "") And (vgFechaFinMortalVit_M <> "") Then
            If ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) >= (vgFechaIniMortalVit_M)) And _
            ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) <= (vgFechaFinMortalVit_M)) Then
                If (vgIndicadorTipoMovimiento_M = "E") Or (vgIndicadorTipoMovimiento_M <> "E" And iSexoCau = "F") Then
                    vgBuscarMortalVit_M = "N"
                Else
                    vgBuscarMortalVit_M = "S"
                End If
            Else
                vgBuscarMortalVit_M = "S"
            End If
        Else
            vgBuscarMortalVit_M = "S"
        End If
    End If
    
    If (vgBuscarMortalVit_M = "S") Then
        'For h = 1 To 2  '1=Causante '2=Beneficiario
        h = 1
            vgs_Sexo = ""
            vgs_Tipo = ""
            vgs_Nro = ""
            'For i = 1 To 2
            '    If i = 1 Then vgs_Sexo = "M"    ' hombre  (M)asculino
            '    If i = 2 Then vgs_Sexo = "F"    ' mujer   (F)emenino
                i = 1
                vgs_Sexo = "M"
                
                'For j = 1 To 3
                j = 2
                    'If j = 1 Then vgs_Tipo = "MIT"    '1  = Invalido Total
                    If j = 2 Then vgs_Tipo = "RV"    '2  = No Invalido
                    'If h = 2 And j = 2 Then vgs_Tipo = "B" ' Tabla de beneficiarios no invalidos
                    'If j = 3 Then vgs_Tipo = "MIP"    '3  = invalido Parcial
                
                vgSw = False
                
                'Buscar Número Correlativo desde la Estructura
                For vgI = 1 To vgNumeroTotalTablas
                                      
                    'Sql = " SELECT * "
                    'Sql = Sql & " from PR_TVAL_MORTAL "
                    'Sql = Sql & " where cod_sexo = '" & vgs_Sexo & "' and "
                    'Sql = Sql & " cod_tiptabmor ='" & vgs_Tipo & "' and "
                    'Sql = Sql & " fec_ini <= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "# AND "
                    'Sql = Sql & " fec_ter >= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "#"
                    'Set tb = vgConexionBD.Execute(Sql)
                    'If Not (tb.EOF) Then
                    
                    vlResPregunta = False
                    If (vgUtilizarNormativa = "S") Then
                        If (egTablaMortal(vgI).Correlativo = vgMortalVit_M_Normativa) Then
                            vlResPregunta = True
                        End If
                    Else
                        If (egTablaMortal(vgI).Sexo = vgs_Sexo) And _
                        (egTablaMortal(vgI).TipoTabla = vgs_Tipo) And _
                        (egTablaMortal(vgI).Nombre = vgPalabra_MortalVit_M) And _
                        (egTablaMortal(vgI).FechaIni <= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) And _
                        (egTablaMortal(vgI).FechaFin >= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) Then
                            vlResPregunta = True
                        End If
                    End If
                    
                    If (vlResPregunta = True) Then
                        vgs_Error = 0
                        vgs_Nro = egTablaMortal(vgI).Correlativo
                        vgFechaIniMortalVit_M = egTablaMortal(vgI).FechaIni
                        vgFechaFinMortalVit_M = egTablaMortal(vgI).FechaFin
                        vgIndicadorTipoMovimiento_M = egTablaMortal(vgI).TipoMovimiento
                        vgDinamicaAñoBase_M = egTablaMortal(vgI).AñoBase
                        
                        'Limpiar columna de Datos
                        For vgX = 1 To Fintab
                            Lx(i, j, vgX) = 0
                        Next vgX
                        
                        'vgs_Nro = tb!num_correlativo
                        'vgFechaInicioMortal = Format(tb!fec_ini, "yyyymmdd")
                        'vgFechaFinMortal = Format(tb!fec_ter, "yyyymmdd")
                        ''vgFechaInicioMortal = DateSerial(Year(tb!fec_ini), Month(tb!fec_ini), Day(tb!fec_ini))
                        ''vgFechaFinMortal = DateSerial(Year(tb!fec_ter), Month(tb!fec_ter), Day(tb!fec_ter))
                        If vgs_Nro <> 0 Then
                            If (vgIndicadorTipoMovimiento_M <> "D") Then
                                Sql = "Select num_edad AS edad,mto_lx from MA_TVAL_MORDET "
                                Sql = Sql & " where num_correlativo = " & vgs_Nro
                                Sql = Sql & " order by num_edad "
                                Set Tb2 = vgConexionBD.Execute(Sql)
                                If Not (Tb2.EOF) Then
                                    vgSw = True
                                    'tb2.MoveFirst
                                    'k = 1
                                    Do While Not Tb2.EOF
                                        k = Tb2!Edad
                                        'If h = 1 Then   'Causante
                                            Lx(i, j, k) = Tb2!mto_lx
                                        'Else    'Beneficiario
                                        '    ly(i, j, k) = tb2!mto_lx
                                        'End If
                                        'k = k + 1
                                        Tb2.MoveNext
                                    Loop
                                Else
                                    vgError = 1065
                                    Exit Function
                                End If
                                Tb2.Close
                                
                                Exit For
                            Else
                                'Obtener la Tabla Temporal Dinánica
                                'If (fgCrearMortalidadDinamica(iNavig, iNmvig, iNdvig, _
                                iNap, iNmp, iNdp, i, j, vgs_Nro, Fintab, vgDinamicaAñoBase_M, iFechaNacCau) = False) Then
                                    
                                If (fgCrearMortalidadDinamica_DaniMensual(iNavig, iNmvig, iNdvig, _
                                iNap, iNmp, iNdp, i, j, vgs_Nro, Fintab, vgDinamicaAñoBase_M, iFechaNacCau) = False) Then
                                    vgError = 1061
                                    Exit Function
                                Else
                                    vgSw = True
                                End If
                            End If
                        Else
                            vgError = 1065
                            Exit Function
                        End If
                    'Else
                    '    vgError = 1065
                    '    Exit Function
                    End If
                    'tb.Close
                Next vgI
                If (vgSw = False) Then
                    vgError = 1065
                    Exit Function
                End If
                'Next j
            'Next i
        'Next h
    End If

    '6. Leer Tabla de Mortalidad de Inv. Totales Hombre
    vgBuscarMortalTot_M = ""
    If (vgUtilizarNormativa = "S") Then
        If (vgFechaIniMortalTot_M = "") And (vgFechaFinMortalTot_M = "") Then
            vgBuscarMortalTot_M = "S"
        End If
    Else
        If (vgFechaIniMortalTot_M <> "") And (vgFechaFinMortalTot_M <> "") Then
            If ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) >= (vgFechaIniMortalTot_M)) And _
            ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) <= (vgFechaFinMortalTot_M)) Then
                vgBuscarMortalTot_M = "N"
            Else
                vgBuscarMortalTot_M = "S"
            End If
        Else
            vgBuscarMortalTot_M = "S"
        End If
    End If
    
    If (vgBuscarMortalTot_M = "S") Then
        For h = 1 To 2  '1=Causante '2=Beneficiario
            vgs_Sexo = ""
            vgs_Tipo = ""
            vgs_Nro = ""
            'For i = 1 To 2
            '    If i = 1 Then vgs_Sexo = "M"    ' hombre  (M)asculino
            '    If i = 2 Then vgs_Sexo = "F"    ' mujer   (F)emenino
                i = 1
                vgs_Sexo = "M"
                
                'For j = 1 To 3
                j = 1
                    If j = 1 Then vgs_Tipo = "MIT"    '1  = Invalido Total
                    'If j = 2 Then vgs_Tipo = "RV"    '2  = No Invalido
                    'If h = 2 And j = 2 Then vgs_Tipo = "B" ' Tabla de beneficiarios no invalidos
                    'If j = 3 Then vgs_Tipo = "MIP"    '3  = invalido Parcial
                
                vgSw = False
                
                'Buscar Número Correlativo desde la Estructura
                For vgI = 1 To vgNumeroTotalTablas
                                      
                    'Sql = " SELECT * "
                    'Sql = Sql & " from PR_TVAL_MORTAL "
                    'Sql = Sql & " where cod_sexo = '" & vgs_Sexo & "' and "
                    'Sql = Sql & " cod_tiptabmor ='" & vgs_Tipo & "' and "
                    'Sql = Sql & " fec_ini <= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "# AND "
                    'Sql = Sql & " fec_ter >= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "#"
                    'Set tb = vgConexionBD.Execute(Sql)
                    'If Not (tb.EOF) Then
                    
                    vlResPregunta = False
                    If (vgUtilizarNormativa = "S") Then
                        If (egTablaMortal(vgI).Correlativo = vgMortalTot_M_Normativa) Then
                            vlResPregunta = True
                        End If
                    Else
                        If (egTablaMortal(vgI).Sexo = vgs_Sexo) And _
                        (egTablaMortal(vgI).TipoTabla = vgs_Tipo) And _
                        (egTablaMortal(vgI).Nombre = vgPalabra_MortalTot_M) And _
                        (egTablaMortal(vgI).FechaIni <= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) And _
                        (egTablaMortal(vgI).FechaFin >= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) Then
                            vlResPregunta = True
                        End If
                    End If
                    
                    If (vlResPregunta = True) Then
                        vgs_Error = 0
                        vgs_Nro = egTablaMortal(vgI).Correlativo
                        vgFechaIniMortalTot_M = egTablaMortal(vgI).FechaIni
                        vgFechaFinMortalTot_M = egTablaMortal(vgI).FechaFin
                        
                        'Limpiar columna de Datos
                        If (h = 1) Then
                            For vgX = 1 To Fintab
                                Lx(i, j, vgX) = 0
                            Next vgX
                        Else
                            For vgX = 1 To Fintab
                                Ly(i, j, vgX) = 0
                            Next vgX
                        End If
                        
                        'vgs_Nro = tb!num_correlativo
                        'vgFechaInicioMortal = Format(tb!fec_ini, "yyyymmdd")
                        'vgFechaFinMortal = Format(tb!fec_ter, "yyyymmdd")
                        ''vgFechaInicioMortal = DateSerial(Year(tb!fec_ini), Month(tb!fec_ini), Day(tb!fec_ini))
                        ''vgFechaFinMortal = DateSerial(Year(tb!fec_ter), Month(tb!fec_ter), Day(tb!fec_ter))
                        If vgs_Nro <> 0 Then
                            Sql = "Select num_edad AS edad,mto_lx from MA_TVAL_MORDET "
                            Sql = Sql & " where num_correlativo = " & vgs_Nro
                            Sql = Sql & " order by num_edad "
                            Set Tb2 = vgConexionBD.Execute(Sql)
                            If Not (Tb2.EOF) Then
                                vgSw = True
                                'tb2.MoveFirst
                                'k = 1
                                Do While Not Tb2.EOF
                                    k = Tb2!Edad
                                    If h = 1 Then   'Causante
                                        Lx(i, j, k) = Tb2!mto_lx
                                    Else    'Beneficiario
                                        Ly(i, j, k) = Tb2!mto_lx
                                    End If
                                    'k = k + 1
                                    Tb2.MoveNext
                                Loop
                            Else
                                vgError = 1066
                                Exit Function
                            End If
                            Tb2.Close
                            
                            Exit For
                        Else
                            vgError = 1066
                            Exit Function
                        End If
                    'Else
                    '    vgError = 1066
                    '    Exit Function
                    End If
                    'tb.Close
                Next vgI
                If (vgSw = False) Then
                    vgError = 1066
                    Exit Function
                End If
                'Next j
            'Next i
        Next h
    End If

    '7. Leer Tabla de Mortalidad de Inv. Parciales Hombre
    vgBuscarMortalPar_M = ""
    If (vgUtilizarNormativa = "S") Then
        If (vgFechaIniMortalPar_M = "") And (vgFechaFinMortalPar_M = "") Then
            vgBuscarMortalPar_M = "S"
        End If
    Else
        If (vgFechaIniMortalPar_M <> "") And (vgFechaFinMortalPar_M <> "") Then
            If ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) >= (vgFechaIniMortalPar_M)) And _
            ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) <= (vgFechaFinMortalPar_M)) Then
                vgBuscarMortalPar_M = "N"
            Else
                vgBuscarMortalPar_M = "S"
            End If
        Else
            vgBuscarMortalPar_M = "S"
        End If
    End If
    
    If (vgBuscarMortalPar_M = "S") Then
        For h = 1 To 2  '1=Causante '2=Beneficiario
            vgs_Sexo = ""
            vgs_Tipo = ""
            vgs_Nro = ""
            'For i = 1 To 2
            '    If i = 1 Then vgs_Sexo = "M"    ' hombre  (M)asculino
            '    If i = 2 Then vgs_Sexo = "F"    ' mujer   (F)emenino
                i = 1
                vgs_Sexo = "M"
                
                'For j = 1 To 3
                j = 3
                    'If j = 1 Then vgs_Tipo = "MIT"    '1  = Invalido Total
                    'If j = 2 Then vgs_Tipo = "RV"    '2  = No Invalido
                    'If h = 2 And j = 2 Then vgs_Tipo = "B" ' Tabla de beneficiarios no invalidos
                    If j = 3 Then vgs_Tipo = "MIP"    '3  = invalido Parcial
                
                vgSw = False
                
                'Buscar Número Correlativo desde la Estructura
                For vgI = 1 To vgNumeroTotalTablas
                                      
                    'Sql = " SELECT * "
                    'Sql = Sql & " from PR_TVAL_MORTAL "
                    'Sql = Sql & " where cod_sexo = '" & vgs_Sexo & "' and "
                    'Sql = Sql & " cod_tiptabmor ='" & vgs_Tipo & "' and "
                    'Sql = Sql & " fec_ini <= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "# AND "
                    'Sql = Sql & " fec_ter >= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "#"
                    'Set tb = vgConexionBD.Execute(Sql)
                    'If Not (tb.EOF) Then
                    
                    vlResPregunta = False
                    If (vgUtilizarNormativa = "S") Then
                        If (egTablaMortal(vgI).Correlativo = vgMortalPar_M_Normativa) Then
                            vlResPregunta = True
                        End If
                    Else
                        If (egTablaMortal(vgI).Sexo = vgs_Sexo) And _
                        (egTablaMortal(vgI).TipoTabla = vgs_Tipo) And _
                        (egTablaMortal(vgI).Nombre = vgPalabra_MortalPar_M) And _
                        (egTablaMortal(vgI).FechaIni <= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) And _
                        (egTablaMortal(vgI).FechaFin >= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) Then
                            vlResPregunta = True
                        End If
                    End If
                    
                    If (vlResPregunta = True) Then
                        vgs_Error = 0
                        vgs_Nro = egTablaMortal(vgI).Correlativo
                        vgFechaIniMortalPar_M = egTablaMortal(vgI).FechaIni
                        vgFechaFinMortalPar_M = egTablaMortal(vgI).FechaFin
                        
                        'Limpiar columna de Datos
                        If (h = 1) Then
                            For vgX = 1 To Fintab
                                Lx(i, j, vgX) = 0
                            Next vgX
                        Else
                            For vgX = 1 To Fintab
                                Ly(i, j, vgX) = 0
                            Next vgX
                        End If
                        
                        'vgs_Nro = tb!num_correlativo
                        'vgFechaInicioMortal = Format(tb!fec_ini, "yyyymmdd")
                        'vgFechaFinMortal = Format(tb!fec_ter, "yyyymmdd")
                        ''vgFechaInicioMortal = DateSerial(Year(tb!fec_ini), Month(tb!fec_ini), Day(tb!fec_ini))
                        ''vgFechaFinMortal = DateSerial(Year(tb!fec_ter), Month(tb!fec_ter), Day(tb!fec_ter))
                        If vgs_Nro <> 0 Then
                            Sql = "Select num_edad AS edad,mto_lx from MA_TVAL_MORDET "
                            Sql = Sql & " where num_correlativo = " & vgs_Nro
                            Sql = Sql & " order by num_edad "
                            Set Tb2 = vgConexionBD.Execute(Sql)
                            If Not (Tb2.EOF) Then
                                vgSw = True
                                'tb2.MoveFirst
                                'k = 1
                                Do While Not Tb2.EOF
                                    k = Tb2!Edad
                                    If h = 1 Then   'Causante
                                        Lx(i, j, k) = Tb2!mto_lx
                                    Else    'Beneficiario
                                        Ly(i, j, k) = Tb2!mto_lx
                                    End If
                                    'k = k + 1
                                    Tb2.MoveNext
                                Loop
                            Else
                                vgError = 1067
                                Exit Function
                            End If
                            Tb2.Close
                            
                            Exit For
                        Else
                            vgError = 1067
                            Exit Function
                        End If
                    'Else
                    '    vgError = 1067
                    '    Exit Function
                    End If
                    'tb.Close
                Next vgI
                If (vgSw = False) Then
                    vgError = 1067
                    Exit Function
                End If
                'Next j
            'Next i
        Next h
    End If

    '8. Leer Tabla de Mortalidad de Beneficiarios Hombre
    vgBuscarMortalBen_M = ""
    If (vgUtilizarNormativa = "S") Then
        If (vgFechaIniMortalBen_M = "") And (vgFechaFinMortalBen_M = "") Then
            vgBuscarMortalBen_M = "S"
        End If
    Else
        If (vgFechaIniMortalBen_M <> "") And (vgFechaFinMortalBen_M <> "") Then
            If ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) >= (vgFechaIniMortalBen_M)) And _
            ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) <= (vgFechaFinMortalBen_M)) Then
                vgBuscarMortalBen_M = "N"
            Else
                vgBuscarMortalBen_M = "S"
            End If
        Else
            vgBuscarMortalBen_M = "S"
        End If
    End If
    
    If (vgBuscarMortalBen_M = "S") Then
        'For h = 1 To 2  '1=Causante '2=Beneficiario
        h = 2
            vgs_Sexo = ""
            vgs_Tipo = ""
            vgs_Nro = ""
            'For i = 1 To 2
            '    If i = 1 Then vgs_Sexo = "M"    ' hombre  (M)asculino
            '    If i = 2 Then vgs_Sexo = "F"    ' mujer   (F)emenino
                i = 1
                vgs_Sexo = "M"
                
                'For j = 1 To 3
                j = 2
                    'If j = 1 Then vgs_Tipo = "MIT"    '1  = Invalido Total
                    'If j = 2 Then vgs_Tipo = "RV"    '2  = No Invalido
                    If h = 2 And j = 2 Then vgs_Tipo = "B" ' Tabla de beneficiarios no invalidos
                    'If j = 3 Then vgs_Tipo = "MIP"    '3  = invalido Parcial
                
                vgSw = False
                
                'Buscar Número Correlativo desde la Estructura
                For vgI = 1 To vgNumeroTotalTablas
                                      
                    'Sql = " SELECT * "
                    'Sql = Sql & " from PR_TVAL_MORTAL "
                    'Sql = Sql & " where cod_sexo = '" & vgs_Sexo & "' and "
                    'Sql = Sql & " cod_tiptabmor ='" & vgs_Tipo & "' and "
                    'Sql = Sql & " fec_ini <= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "# AND "
                    'Sql = Sql & " fec_ter >= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "#"
                    'Set tb = vgConexionBD.Execute(Sql)
                    'If Not (tb.EOF) Then
                    
                    vlResPregunta = False
                    If (vgUtilizarNormativa = "S") Then
                        If (egTablaMortal(vgI).Correlativo = vgMortalBen_M_Normativa) Then
                            vlResPregunta = True
                        End If
                    Else
                        If (egTablaMortal(vgI).Sexo = vgs_Sexo) And _
                        (egTablaMortal(vgI).TipoTabla = vgs_Tipo) And _
                        (egTablaMortal(vgI).Nombre = vgPalabra_MortalBen_M) And _
                        (egTablaMortal(vgI).FechaIni <= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) And _
                        (egTablaMortal(vgI).FechaFin >= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) Then
                            vlResPregunta = True
                        End If
                    End If
                    
                    If (vlResPregunta = True) Then
                        vgs_Error = 0
                        vgs_Nro = egTablaMortal(vgI).Correlativo
                        vgFechaIniMortalBen_M = egTablaMortal(vgI).FechaIni
                        vgFechaFinMortalBen_M = egTablaMortal(vgI).FechaFin
                        
                        'Limpiar columna de Datos
                        For vgX = 1 To Fintab
                            Ly(i, j, vgX) = 0
                        Next vgX
                        
                        'vgs_Nro = tb!num_correlativo
                        'vgFechaInicioMortal = Format(tb!fec_ini, "yyyymmdd")
                        'vgFechaFinMortal = Format(tb!fec_ter, "yyyymmdd")
                        ''vgFechaInicioMortal = DateSerial(Year(tb!fec_ini), Month(tb!fec_ini), Day(tb!fec_ini))
                        ''vgFechaFinMortal = DateSerial(Year(tb!fec_ter), Month(tb!fec_ter), Day(tb!fec_ter))
                        If vgs_Nro <> 0 Then
                            Sql = "Select num_edad AS edad,mto_lx from MA_TVAL_MORDET "
                            Sql = Sql & " where num_correlativo = " & vgs_Nro
                            Sql = Sql & " order by num_edad "
                            Set Tb2 = vgConexionBD.Execute(Sql)
                            If Not (Tb2.EOF) Then
                                vgSw = True
                                'tb2.MoveFirst
                                'k = 1
                                Do While Not Tb2.EOF
                                    k = Tb2!Edad
                                    'If h = 1 Then   'Causante
                                    '    lx(i, j, k) = tb2!mto_lx
                                    'Else    'Beneficiario
                                        Ly(i, j, k) = Tb2!mto_lx
                                    'End If
                                    'k = k + 1
                                    Tb2.MoveNext
                                Loop
                            Else
                                vgError = 1068
                                Exit Function
                            End If
                            Tb2.Close
                            
                            Exit For
                        Else
                            vgError = 1068
                            Exit Function
                        End If
                    'Else
                    '    vgError = 1068
                    '    Exit Function
                    End If
                    'tb.Close
                Next vgI
                If (vgSw = False) Then
                    vgError = 1068
                    Exit Function
                End If
                'Next j
            'Next i
        'Next h
    End If

    fgBuscarMortalidadNormativa = True
End Function

Function fgEdad(Fechap As Long, idp As Integer, Nanbe As Integer, Nmnbe As Integer, Ndnbe As Integer, Ncorbe As Integer) As Integer
Dim Fechameses1 As Long, fechameses2 As Long
    
    edabe = Fechap - (Nanbe * 12 + Nmnbe)
    difdia = idp - Ndnbe
    If difdia > 15 Then edabe = edabe + 1
    If (Ncorbe = 30 Or Ncorbe = 35) And (edabe = L24 And difdia > 15) Then edabe = edabe - 1
    fgEdad = edabe
End Function

Function fgCodInvalidez(codinval As String) As Integer
    Select Case (codinval)
        Case "S", "T": fgCodInvalidez = 1
        Case "N": fgCodInvalidez = 2
        Case "P": fgCodInvalidez = 3
        Case Else
            fgCodInvalidez = 0
    End Select
End Function

Function fgBisiestro(Nap, Nmp As Variant) As Integer
    Dim iagno As Integer
    iagno = Nap Mod 4
    If iagno > 0 Then idbis = 365 Else idbis = 366
    Select Case (Nmp)
        Case (2)
            If idbis = 365 Then idp = 28 Else idp = 29
        Case 1, 3, 5, 7, 8, 10, 12
            idp = 31
        Case Else
            idp = 30
    End Select
    fgBisiestro = idp
End Function

Function fgMaximo(param1, param2) As Variant
    If param1 > param2 Then
        fgMaximo = param1
    Else
        fgMaximo = param2
    End If
End Function


Function fgMinimo(param1, param2) As Variant
    If param1 > param2 Then
        fgMinimo = param2
    Else
        fgMinimo = param1
    End If
End Function

Function fgNCobe(Cober) As Integer
'Function fgNCobe(Cober As Integer) As Integer
Dim ncobe As Integer
Select Case (Cober)
    Case 1: ncobe = 3
    Case 2, 3: ncobe = 2
    Case 4, 9: ncobe = 1
    Case 5, 10: ncobe = 1
    Case 6, 11: ncobe = 2
    Case 14, 15: ncobe = 2
    Case 8, 13: ncobe = 3
    Case 7, 12: ncobe = 4
    Case Else
        ncobe = 0
End Select
fgNCobe = ncobe
End Function

Function fgMesNoC(Iare, Imre, Navig, Nmvig, Nap, Nmp) As Long
'Function fgMesNoC(Iare, Imre, Navig, Nmvig, Nap, Nmp As Integer) As Long
Dim diftem As Long
Dim mescon As Long
    diftem = ((Iare * 12) + Imre) - ((Navig * 12) + Nmvig)
    mescon = ((Nap * 12) + Nmp) - ((Navig * 12) + Nmvig)
    mesnoc = diftem - mescon - 1
    mesnoc = fgMaximo(0, mesnoc)
    fgMesNoC = mesnoc
End Function

Function fgSexo(codsexo As String) As Integer
    Select Case (codsexo)
        Case "M": fgSexo = 1
        Case "F": fgSexo = 2
        Case Else
            fgSexo = 0
    End Select
End Function



'***********************************************************
'Fin : Funciones agregadas para el manejo de los Endosos
'***********************************************************
'F---- ABV 17/11/2005 ---




'CMV-20060601 I
'***********************************************************
'INICIO:      Procedimientos para Endosos Automáticos
'***********************************************************

'Función Principal Generación Endosos Automáticos
Function fgGenerarEndosoAutomatico(iFactorAjuste As Double, iFactorAjusteTasaFija As Double) As Boolean
On Error GoTo Err_fgGenerarEndosoAutomatico

'    If flValidaEstadoProceso(Mid(Trim(vlFechaInicio), 1, 6), vlCodEstado) = False Then
'        MsgBox "El Tipo de Proceso Seleccionado no se encuentra Realizado.", vbCritical, "Error de Datos"
'        Exit Function
'    End If
    
'    MsgBox "Este Proceso realizará Endosos Automáticos Definitivos." & Chr(13) & _
'           "¿ Está seguro que desea ejecutar el", vbCritical, "Operación Cancelada"
'    vgRes = MsgBox("¿ Está seguro que desea Modificar los Datos ?", 4 + 32 + 256, "Operación de Actualización")
        
        
    Screen.MousePointer = 11
        
'    vgRes = MsgBox("Este Proceso realizará Endosos Automáticos Definitivos." & Chr(13) & " ¿ Está seguro que desea ejecutar este proceso ?", 4 + 32 + 256, "Operación Cálculo")
'    If vgRes <> 6 Then
'       Screen.MousePointer = 0
'       Exit Function
'    End If
    fgGenerarEndosoAutomatico = True
    vlFactorAjusteEndoso = iFactorAjuste
    vlFactorAjusteTasaFijaEndoso = iFactorAjusteTasaFija 'hqr 03/12/2010
    vlCamposPoliza = ""
    vlCamposPoliza = "p.num_poliza,p.num_endoso, "
    vlCamposPoliza = vlCamposPoliza & "p.cod_afp,p.cod_tippension,p.cod_estado,p.cod_tipren, "
    vlCamposPoliza = vlCamposPoliza & "p.cod_modalidad,p.num_cargas,p.fec_vigencia, "
    vlCamposPoliza = vlCamposPoliza & "p.fec_tervigencia,p.mto_prima,p.mto_pension,p.num_mesdif, "
    vlCamposPoliza = vlCamposPoliza & "p.num_mesgar,p.prc_tasace,p.prc_tasavta,p.prc_tasactorea, "
    vlCamposPoliza = vlCamposPoliza & "p.prc_tasaintpergar,p.fec_inipagopen, "
    vlCamposPoliza = vlCamposPoliza & "p.cod_usuariocrea,p.fec_crea,p.hor_crea, "
    vlCamposPoliza = vlCamposPoliza & "p.cod_tiporigen , p.num_indquiebra, "
    vlCamposPoliza = vlCamposPoliza & "p.mto_pensiongar as mto_pensiongarpol,p.cod_cuspp,p.ind_cob, p.cod_moneda as cod_monedapol,"
    vlCamposPoliza = vlCamposPoliza & "p.mto_valmoneda,p.cod_cobercon,p.mto_facpenella,p.prc_facpenella,"
    vlCamposPoliza = vlCamposPoliza & "p.cod_dercre as cod_dercrepol,p.cod_dergra,p.prc_tasatir,p.fec_emision,p.fec_dev,"
    vlCamposPoliza = vlCamposPoliza & "p.fec_inipencia,p.fec_pripago,p.fec_finperdif,p.fec_finpergar"
    vlCamposPoliza = vlCamposPoliza & ",p.cod_tipreajuste,p.mto_valreajustetri, mto_valreajustemen, p.fec_devsol" 'hqr 03/12/2010
    
    vlCamposBen = ""
    vlCamposBen = "b.num_poliza,b.num_endoso, "
    vlCamposBen = vlCamposBen & "b.num_orden,b.fec_ingreso,b.num_idenben,b.cod_tipoidenben, "
    vlCamposBen = vlCamposBen & "b.gls_nomben,b.gls_nomsegben,b.gls_patben,b.gls_matben, "
    vlCamposBen = vlCamposBen & "b.gls_dirben,b.cod_direccion,b.gls_fonoben, "
    vlCamposBen = vlCamposBen & "b.gls_correoben,b.cod_grufam,b.cod_par, "
    vlCamposBen = vlCamposBen & "b.cod_sexo,b.cod_sitinv,b.cod_dercre, "
    vlCamposBen = vlCamposBen & "b.cod_derpen,b.cod_cauinv,b.fec_nacben, "
    vlCamposBen = vlCamposBen & "b.fec_nachm,b.fec_invben,b.cod_motreqpen, "
    'vlCamposBen = vlCamposBen & "b.mto_pension,b.mto_pensiongar,b.prc_pension, "
    vlCamposBen = vlCamposBen & "b.mto_pension as pensionben,b.mto_pensiongar,b.prc_pension, " 'hqr 06/07/2006 se cambia nombre de la pension
    vlCamposBen = vlCamposBen & "b.cod_inssalud,b.cod_modsalud,b.mto_plansalud, "
    vlCamposBen = vlCamposBen & "b.cod_estpension,b.cod_viapago, "
    vlCamposBen = vlCamposBen & "b.cod_banco,b.cod_tipcuenta,b.num_cuenta, "
    vlCamposBen = vlCamposBen & "b.cod_sucursal,b.fec_fallben,b.fec_matrimonio, "
    vlCamposBen = vlCamposBen & "b.cod_caususben,b.fec_susben,b.fec_inipagopen as fecinipagopen, "
    vlCamposBen = vlCamposBen & "b.fec_terpagopengar,b.cod_usuariocrea as cod_usuariocreaben, "
    vlCamposBen = vlCamposBen & "b.fec_crea as fec_creaben,b.hor_crea as hor_creaben,b.cod_usuariomodi, "
    vlCamposBen = vlCamposBen & "b.fec_modi,b.hor_modi, "
    'hqr 27/08/2007 campos agregados
    'vlCamposBen = vlCamposBen & "b.cod_modsalud2,b.mto_plansalud2, b.num_fun, "
    vlCamposBen = vlCamposBen & "b.prc_pensionleg,b.prc_pensiongar"

    'Cargar Barra de Progreso en Pantalla
    vlLargoArchivo = 190
    vlLargoRegistro = 185
    vlAumento = CDbl((90 / vlLargoArchivo) * vlLargoRegistro)
    Frm_PensPagosRegimen.Refresh
    Frm_PensPagosRegimen.ProgressBar.Value = 0
    Frm_PensPagosRegimen.Frame3 = "Progreso del Cálculo de Endosos Automáticos"
    Frm_PensPagosRegimen.Refresh
        
    vgGeneraEndosos = False 'No se hicieron Endosos
    'RRR 13/11/2013
'    If Not (fgBenConCerEst = False) Then
'        fgGenerarEndosoAutomatico = 1
'        Exit Function
'    End If

'    If Not (fgBenSinCerEst = False) Then
'        fgGenerarEndosoAutomatico = 2
'        MsgBox "Existen Beneficiarios que no cuentan con un certificado de Supervivencia. Revisar el reporte y crear los certificados.", vbCritical, "Pago de Pensiones"
'        Exit Function
'    End If
    
    If Not fgBenMayor18 Then
        fgGenerarEndosoAutomatico = True
        Exit Function
    End If
    
    'Unload Frm_BarraProg
    
    Screen.MousePointer = 0

Exit Function
Err_fgGenerarEndosoAutomatico:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        fgGenerarEndosoAutomatico = 0
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function fgBenConCerEst() As Boolean
On Error GoTo Err_fgBenConCerEst
'Función que permite modificar el estado de derecho a pensión
'de todos aquellos beneficiarios hijos, de entre 18 y 24 años
'de edad, que tienen certificado de estudio vigente
Dim vlAnno As String
Dim vlMes As String
Dim vlDia As String
Dim vlFechaCalculo As String
Dim vlFechaTermino As String


    'SE COMENATRA LA CREACION DE ENDOSOS EN ESTE PROCESO YA QUE NO SON NECESARIOS CREAR ENDOSO FICTISIOS RRR


    fgBenConCerEst = False
    vlFechaCalculo = vgFecPago '"20050720"

    'vlFechaActual = fgBuscaFecServ
    vlFechaActual = vlFechaCalculo
    vlFechaTermino = Format(DateAdd("d", -1, DateAdd("m", (vgMesesExtension + 1), DateSerial(Mid(vlFechaCalculo, 1, 4), Mid(vlFechaCalculo, 5, 2), 1))), "yyyymmdd")
    'vlFechaActual = Format(vlFechaActual, "yyyymmdd")
    'Fecha Con último dia del mes
    'vlFecha18 = DateSerial(Mid((vlFechaActual), 1, 4) - 18, Mid((vlFechaActual), 5, 2) + 1, 1 - 1)
    vlFecha18 = DateSerial(Mid((vlFechaActual), 1, 4) - 18, Mid((vlFechaActual), 5, 2), 0)  'hqr 08/07/2006
    vlFecha18 = Format(vlFecha18, "yyyymmdd")
    'Fecha Con primer dia del mes
    vlFecha24 = DateSerial(Mid((vlFechaActual), 1, 4) - 24, Mid((vlFechaActual), 5, 2), 1)
    vlFecha24 = Format(vlFecha24, "yyyymmdd")

    'Seleccionar todos aquellos beneficiarios con parentesco hijo
    'con edad entre 18 y 24 años que SI tengan certificado de
    'estudios vigente
    
    'Obtiene los beneficiarios que tengan certificado de Supervivencia
    vgSql = ""
    vgSql = "SELECT " & vlCamposPoliza & " , " & vlCamposBen & " "
    vgSql = vgSql & "FROM pp_tmae_ben b, pp_tmae_poliza p, pp_tmae_certificado c "
    vgSql = vgSql & "WHERE " 'b.fec_nacben >= '" & vlFecha24 & "' AND "
    'vgSql = vgSql & "b.fec_nacben <= '" & vlFecha18 & "' AND "
    'vgSql = vgSql & "(b.cod_par = '" & clCodParHijo30 & "' OR "
    'vgSql = vgSql & "b.cod_par = '" & clCodParHijo35 & "') AND "
    'vgSql = vgSql & "b.cod_sitinv = '" & clCodSitInvN & "' AND "
    vgSql = vgSql & "c.fec_inicer <= '" & vlFechaCalculo & "' AND "
    vgSql = vgSql & "c.fec_tercer >= '" & vlFechaTermino & "' AND "
    vgSql = vgSql & "c.cod_tipo = 'SUP' AND " 'Certificado de Supervivencia
    vgSql = vgSql & "b.cod_estpension = '" & clCodEstPension20 & "' AND "
    vgSql = vgSql & "b.num_poliza = p.num_poliza AND "
    vgSql = vgSql & "b.num_endoso = p.num_endoso AND "
    vgSql = vgSql & "p.num_endoso = (SELECT MAX (num_endoso) as numero "
    vgSql = vgSql & "FROM pp_tmae_poliza "
    vgSql = vgSql & "WHERE num_poliza = p.num_poliza) AND "
    vgSql = vgSql & "b.num_poliza = c.num_poliza AND "
    vgSql = vgSql & "b.num_orden = c.num_orden "
    'hqr 06/10/2007 Se agrega para que no le genere Endoso a las pólizas que no están pagando pensiones
    vgSql = vgSql & " AND p.cod_estado IN (6, 7, 8) " 'solo Pólizas Vigentes, con Beneficiarios Designados y Pendiente (Pago Diferido)
    'Se deben identificar las Pólizas que tengan su Primer Pago en un Periodo Anterior (Inmediatas y Diferidas)
    vgSql = vgSql & " AND (p.fec_pripago < '" & Format(vgFecIniPag, "yyyymmdd") & "'"
    vgSql = vgSql & " OR (p.num_mesdif > 0 AND p.fec_pripago = '" & Format(vgFecIniPag, "yyyymmdd") & "')) order by 1"
    'fin hqr 06/10/2007
    Set vgRegPolizas = vgConexionBD.Execute(vgSql)
    If Not vgRegPolizas.EOF Then
        
        fgBenConCerEst = True 'RRR 13/11/2013
    
        vlNumPoliza = ""
        vlNumPolAnterior = ""
        vlNumEndoso = 0
        vlNumOrden = 0
'        While Not vgRegPolizas.EOF
'            vgGeneraEndosos = True
'            If Frm_PensPagosRegimen.ProgressBar.Value + vlAumento < 100 Then
'                Frm_PensPagosRegimen.ProgressBar.Value = Frm_PensPagosRegimen.ProgressBar.Value + vlAumento
'                Frm_PensPagosRegimen.ProgressBar.Refresh
'            End If
'            vlNumPoliza = (vgRegPolizas!num_poliza)
'            vlNumEndoso = (vgRegPolizas!num_endoso)
'            vlNumOrden = (vgRegPolizas!Num_Orden)
'
'
'            If Trim(vlNumPolAnterior) <> Trim(vgRegPolizas!num_poliza) Then
'
'            'MsgBox vlNumPoliza
'
'                'Obtener numero de proximo endoso
'                vlNumUltEndoso = vlNumEndoso
'                'Asignar Número de Nuevo Endoso
'                vlNumEndosoNuevo = vlNumUltEndoso + 1
'
'                Call fgInicializaVarBen
'                Call fgInicializaVarPol
'                Call fgInicializaVarEnd
'                Call fgAsignarValoresPol
'
'                'Confirmar la No Existencia de Pre-Endoso
'                '(Si existe un pre-endoso, este se elimina para crear el nuevo endoso sin problemas)
'                Call fgVerificarPreEndoso
'
'                'Call fgInicializaVarEnd
'                Call fgAsignarValoresEnd
'                Call fgGeneraEndoso
'
'                Call fgActualizaEndAnterior(vlNumPoliza)
'
'                Call fgGeneraEndPol
'                'Insert Select de todos los beneficiarios de la poliza
'                Call fgGeneraEndBeneficiarios
'
'                'Actualiza el Estado de pension del Beneficiario seleccionado
'                Call fgActualizaEndBen(vlNumPoliza, vlNumEndosoNuevo, vlNumOrden, clCodEstPension99)
'
'                vlNumPolAnterior = vlNumPoliza
'
'            Else
'
'                'Actualiza el Estado de pension del Beneficiario seleccionado
'                Call fgActualizaEndBen(vlNumPoliza, vlNumEndosoNuevo, vlNumOrden, clCodEstPension99)
'
'                vlNumPolAnterior = vlNumPoliza
'
'            End If
'            vgRegPolizas.MoveNext
'        Wend
        
    Else
        If vgTipoPago = "P" Then
            If Frm_PensPrimerosPagos.ProgressBar.Value + vlAumento < 100 Then
                Frm_PensPrimerosPagos.ProgressBar.Value = Frm_PensPrimerosPagos.ProgressBar.Value + vlAumento
                Frm_PensPrimerosPagos.ProgressBar.Refresh
            End If
        Else
            If Frm_PensPagosRegimen.ProgressBar.Value + vlAumento < 100 Then
                Frm_PensPagosRegimen.ProgressBar.Value = Frm_PensPagosRegimen.ProgressBar.Value + vlAumento
                Frm_PensPagosRegimen.ProgressBar.Refresh
            End If
        End If
        
    End If

Exit Function
Err_fgBenConCerEst:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
        fgBenConCerEst = False
    End Select

End Function

Function fgInicializaVarPol()
On Error GoTo Err_fgInicializaVarPol

    vlNumPolizaP = ""
    vlNumEndosoP = 0
    vlCodAFP = ""
    vlCodTipPension = ""
    vlCodEstadoP = ""
    vlCodTipRen = ""
    vlCodModalidad = ""
    vlNumCargas = 0
    vlFecVigencia = ""
    vlFecTerVigencia = ""
    vlMtoPrima = 0
    vlMtoPensionPol = 0
    vlNumMesDif = 0
    vlNumMesGar = 0
    vlPrcTasaCe = 0
    vlPrcTasaVta = 0
    vlPrcTasaCtoRea = 0
    vlPrcTasaIntPerGar = 0
    vlFecIniPagoPenPol = ""
    vlCodUsuarioCreaP = ""
    vlFecCreaP = ""
    vlHorCreaP = ""
    vlCodUsuarioModiP = ""
    vlFecModiP = ""
    vlHorModiP = ""
    vlCodTipOrigen = ""
    vlNumIndQuiebra = 0
    vlMtoPensionGarP = 0
    vlCodCuspp = ""
    vlIndCob = ""
    vlCodMoneda = ""
    vlMtoValMoneda = 0
    vlCodCobercon = ""
    vlMtoFacPenElla = 0
    vlPrcFacPenElla = 0
    vlCodDercreP = ""
    vlCodDerGra = ""
    vlPrcTasaTir = 0
    vlFecEmision = ""
    vlFecDev = ""
    vlFecIniPenCia = ""
    vlFecPriPago = ""
    vlFecFinPerdif = ""
    vlFecFinPerGar = ""
    vlCodTipReajuste = "" 'hqr 03/12/2010
    vlMtoValReajusteTri = 0 'hqr 03/12/2010
    vlMtoValReajusteMen = 0 'hqr 19/02/2011
    vlFecDevSol = ""
Exit Function
Err_fgInicializaVarPol:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function fgAsignarValoresPol()
On Error GoTo Err_fgAsignarValoresPol

    vlNumPolizaP = (vgRegPolizas!num_poliza)
    vlNumEndosoP = vlNumEndosoNuevo
    vlCodAFP = (vgRegPolizas!cod_afp)
    vlCodTipPension = (vgRegPolizas!Cod_TipPension)
    vlCodEstadoP = (vgRegPolizas!Cod_Estado)
    vlCodTipRen = (vgRegPolizas!Cod_TipRen)
    vlCodModalidad = (vgRegPolizas!Cod_Modalidad)
    vlNumCargas = (vgRegPolizas!Num_Cargas)
    vlFecVigencia = (vgRegPolizas!Fec_Vigencia)
    If Not IsNull(vgRegPolizas!Fec_TerVigencia) Then
        vlFecTerVigencia = (vgRegPolizas!Fec_TerVigencia)
    Else
        vlFecTerVigencia = ""
    End If
    vlMtoPrima = (vgRegPolizas!Mto_Prima)
    vlFecDev = vgRegPolizas!fec_dev
                        
    'hqr 15/10/2007 Se debe obtener el valor actualizado de la pension
    vlMtoPensionPol = (vgRegPolizas!Mto_Pension)
    ''If vgRegPolizas!Cod_Monedapol = vgMonedaCodOfi Then 'Solo para nuevos soles
'    If vgRegPolizas!Cod_TipReajuste = cgAJUSTESOLES Or vgRegPolizas!Cod_TipReajuste = cgAJUSTETASAFIJA Then 'hqr 03/12/2010
'        'Obtiene Pension Actualizada
'        vlSql = "SELECT mto_pension FROM pp_tmae_pensionact a"
'        vlSql = vlSql & " WHERE a.num_poliza = '" & vgRegPolizas!num_poliza & "'"
'        vlSql = vlSql & " AND a.num_endoso = " & vgRegPolizas!num_endoso
'        vlSql = vlSql & " AND a.fec_desde = "
'            vlSql = vlSql & " (SELECT max(fec_desde) FROM pp_tmae_pensionact b"
'            vlSql = vlSql & " WHERE b.num_poliza = a.num_poliza"
'            vlSql = vlSql & " AND b.num_endoso = a.num_endoso"
'            vlSql = vlSql & " AND b.fec_desde < '" & Format(vgFecIniPag, "yyyymmdd") & "')"
'        Set vlTB2 = vgConexionBD.Execute(vlSql)
'        If Not vlTB2.EOF Then
'            If Not IsNull(vlTB2!Mto_Pension) Then
'                If vgRegPolizas!Cod_TipReajuste = cgAJUSTESOLES Then 'hqr 03/12/2010
'                    vlMtoPensionPol = Format(vlTB2!Mto_Pension * vlFactorAjusteEndoso, "##0.00")
'                'Ajuste Tasa Fija hqr 03/12/2010
'                Else
'                    'vlMtoPensionPol = Format(vlTB2!Mto_Pension + (vlTB2!Mto_Pension * vlFactorAjusteTasaFijaEndoso * (vgRegPolizas!Mto_ValReajuste / 100)), "##0.00")
'                    vlMtoPensionPol = Format(vlTB2!Mto_Pension * fgObtieneFactorAjusteTasaFija(vlFecVigencia, vgFecIniPag, vgRegPolizas!Mto_ValReajusteTri, vgRegPolizas!Mto_ValReajusteMen, vlFactorAjusteTasaFijaEndoso, vlFecDev), "##0.00")
'                End If
'                'Fin Ajuste Tasa Fija hqr 03/12/2010
'            End If
'        End If
'    End If
    'fin hqr 15/10/2007
    
    vlNumMesDif = (vgRegPolizas!Num_MesDif)
    vlNumMesGar = (vgRegPolizas!Num_MesGar)
    vlPrcTasaCe = (vgRegPolizas!Prc_TasaCe)
    vlPrcTasaVta = (vgRegPolizas!Prc_TasaVta)
    vlPrcTasaCtoRea = (vgRegPolizas!prc_tasactorea)
    If Not IsNull(vgRegPolizas!Prc_TasaIntPerGar) Then
        vlPrcTasaIntPerGar = (vgRegPolizas!Prc_TasaIntPerGar)
    Else
        vlPrcTasaIntPerGar = 0
    End If
    If Not IsNull(vgRegPolizas!Fec_IniPagoPen) Then
        vlFecIniPagoPenPol = (vgRegPolizas!Fec_IniPagoPen)
    Else
        vlFecIniPagoPenPol = ""
    End If
    If Not IsNull(vgRegPolizas!Cod_UsuarioCrea) Then
        vlCodUsuarioCreaP = (vgRegPolizas!Cod_UsuarioCrea)
    Else
        vlCodUsuarioCreaP = ""
    End If
    If Not IsNull(vgRegPolizas!Fec_Crea) Then
        vlFecCreaP = (vgRegPolizas!Fec_Crea)
    Else
        vlFecCreaP = ""
    End If
    If Not IsNull(vgRegPolizas!Hor_Crea) Then
        vlHorCreaP = (vgRegPolizas!Hor_Crea)
    Else
        vlHorCreaP = ""
    End If
    vlCodUsuarioModiP = vgUsuario
    vlFecModiP = Format(Date, "yyyymmdd")
    vlHorModiP = Format(Time, "hhmmss")
    vlCodTipOrigen = (vgRegPolizas!cod_tiporigen)
    If Not IsNull(vgRegPolizas!num_indquiebra) Then
        vlNumIndQuiebra = (vgRegPolizas!num_indquiebra)
    Else
        vlNumIndQuiebra = 0
    End If
    
    'RRR 21/01/2014
    vlMtoPensionPol = vgRegPolizas!Mto_Pension
    
    If vgRegPolizas!mto_pensiongarpol > 0 Then
        vlMtoPensionGarP = (vgRegPolizas!mto_pensiongarpol) 'vlMtoPensionPol
    Else
        vlMtoPensionGarP = 0
    End If
    vlCodCuspp = vgRegPolizas!Cod_Cuspp
    vlIndCob = vgRegPolizas!Ind_Cob
    vlCodMoneda = IIf(IsNull(vgRegPolizas!Cod_Monedapol), "", vgRegPolizas!Cod_Monedapol)
    vlMtoValMoneda = vgRegPolizas!Mto_ValMoneda
    vlCodCobercon = vgRegPolizas!Cod_CoberCon
    vlMtoFacPenElla = vgRegPolizas!Mto_FacPenElla
    vlPrcFacPenElla = vgRegPolizas!Prc_FacPenElla
    vlCodDercreP = vgRegPolizas!cod_dercrepol
    vlCodDerGra = vgRegPolizas!Cod_DerGra
    vlPrcTasaTir = vgRegPolizas!prc_tasatir
    vlFecEmision = vgRegPolizas!Fec_Emision
    vlFecIniPenCia = vgRegPolizas!Fec_IniPenCia
    vlFecPriPago = vgRegPolizas!Fec_PriPago
    vlFecFinPerdif = IIf(IsNull(vgRegPolizas!FEC_FINPERDIF), "", vgRegPolizas!FEC_FINPERDIF)
    vlFecFinPerGar = IIf(IsNull(vgRegPolizas!fec_finpergar), "", vgRegPolizas!fec_finpergar)
    vlCodTipReajuste = IIf(IsNull(vgRegPolizas!Cod_TipReajuste), "", vgRegPolizas!Cod_TipReajuste) 'hqr 03/12/2010
    vlMtoValReajusteTri = IIf(IsNull(vgRegPolizas!Mto_ValReajusteTri), "", vgRegPolizas!Mto_ValReajusteTri) 'hqr 03/12/2010
    vlMtoValReajusteMen = IIf(IsNull(vgRegPolizas!Mto_ValReajusteMen), "", vgRegPolizas!Mto_ValReajusteMen) 'hqr 03/12/2010
    vlFecDevSol = IIf(IsNull(vgRegPolizas!FEC_DEVSOL), "", vgRegPolizas!FEC_DEVSOL)
Exit Function
Err_fgAsignarValoresPol:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function fgInicializaVarBen()
On Error GoTo Err_fgInicializaVarBen

    vlNumPolizaB = ""
    vlNumEndosoB = 0
    vlNumOrdenB = 0
    vlFecIngreso = ""
    vlCodTipoIdenBen = 0
    vlNumIdenBen = ""
    vlGlsNomBen = ""
    vlGlsNomSegBen = ""
    vlGlsPatBen = ""
    vlGlsMatBen = ""
    vlGlsDirBen = ""
    vlCodDireccion = 0
    vlGlsFonoBen = ""
    vlGlsCorreoBen = ""
    vlCodGruFam = ""
    vlCodPar = ""
    vlCodSexo = ""
    vlCodSitInv = ""
    vlCodDerCre = ""
    vlCodDerpen = ""
    vlCodCauInv = ""
    vlFecNacBen = ""
    vlFecNacHM = ""
    vlFecInvBen = ""
    vlCodMotReqPen = ""
    vlMtoPension = 0
    vlMtoPensionGar = 0
    vlPrcPension = 0
    vlCodInsSalud = ""
    vlCodModSalud = ""
    vlMtoPlanSalud = 0
    vlCodEstPension = ""
    vlCodViaPago = ""
    vlCodBanco = ""
    vlCodTipCuenta = ""
    vlNumCuenta = ""
    vlCodSucursal = ""
    vlFecFallBen = ""
    vlFecMatrimonio = ""
    vlCodCauSusBen = ""
    vlFecSusBen = ""
    vlFecIniPagoPen = ""
    vlFecTerPagoPenGar = ""
    vlCodUsuarioCrea = ""
    vlFecCrea = ""
    vlHorCrea = ""
    vlGlsUsuarioModi = ""
    vlFecModi = ""
    vlHorModi = ""
    vlPrcPensionLeg = 0
    vlPrcPensionGar = 0
    
Exit Function
Err_fgInicializaVarBen:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function fgInicializaVarEnd()
On Error GoTo Err_fgInicializaVarEnd

    vlNumPolizaE = ""
    vlNumEndosoE = 0
    vlFecSolEndoso = ""
    vlFecEndoso = ""
    vlCodCauEndoso = ""
    vlCodTipEndoso = ""
    vlMtoDiferencia = 0
    vlCodMonedaEndoso = ""
    vlMtoPensionOri = 0
    vlMtoPensionCal = 0
    vlFecEfecto = ""
    vlPrcFactor = 0
    vlGlsObservacion = ""
    vlCodUsuarioCreaE = ""
    vlFecCreaE = ""
    vlHorCreaE = ""
    vlFecFinEfecto = ""
    vlCodEstadoE = ""
    vlCodTipReajusteE = ""
    vlMtoValReajusteETri = 0
    vlMtoValReajusteEMen = 0
Exit Function
Err_fgInicializaVarEnd:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function fgAsignarValoresEnd()
On Error GoTo Err_fgAsignarValoresEnd

    vlNumPolizaE = vlNumPoliza
    'Se asigna el endoso actual de la poliza ya que
    'Ej.: el endoso 2 de la poliza es el endoso 1 de Endosos
    vlNumEndosoE = vlNumEndoso
    vlFecSolEndoso = fgBuscaFecServ
    vlFecSolEndoso = Format(vlFecSolEndoso, "yyyymmdd")
    vlFecEndoso = fgBuscaFecServ
    vlFecEndoso = Format(vlFecEndoso, "yyyymmdd")
    vlCodCauEndoso = "26"
    vlCodTipEndoso = clCodTipEndosoS
    'Se asignará los valores originales a todas las variables, ya que los casos
    'actuales no afectan los montos de pensión
    vlMtoDiferencia = 0 'hqr 07/07/2006 la Diferencia es cero(vgRegPolizas!Mto_Pension)
    vlCodMonedaEndoso = vgRegPolizas!Cod_Monedapol
    vlMtoPensionOri = vlMtoPensionPol '(vgRegPolizas!Mto_Pension) 'hqr 15/10/2007 Pension Actualizada
    vlMtoPensionCal = vlMtoPensionPol '(vgRegPolizas!Mto_Pension) 'hqr 15/10/2007 Pension Actualizada
    'vlFecEfecto = fgCalculaFechaEfecto(vlFecEfecto)
    vlFecEfecto = vgFecIniPag 'hqr 06/07/2006 Tiene Fecha de Efecto en el mes en que se genera
    vlFecEfecto = Format(vlFecEfecto, "yyyymmdd")
    vlPrcFactor = clPrcFactor1
    vlGlsObservacion = ""
    vlCodUsuarioCreaE = vgUsuario
    vlFecCreaE = Format(Date, "yyyymmdd")
    vlHorCreaE = Format(Time, "hhmmss")
    vlFecFinEfecto = clFechaTope
    vlCodEstadoE = clCodEstadoE
    vlCodTipReajusteE = vlCodTipReajuste 'hqr 03/12/2010
    vlMtoValReajusteETri = vlMtoValReajusteTri 'hqr 03/12/2010
    vlMtoValReajusteEMen = vlMtoValReajusteMen 'hqr 19/02/2011
Exit Function
Err_fgAsignarValoresEnd:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function fgGeneraEndoso()
On Error GoTo Err_fgGeneraEndoso

    Dim a As Integer
    
    If vlNumPolizaE = "0000000397" Then
        a = 1
    End If

    vgSql = ""
    vgSql = "INSERT INTO PP_TMAE_ENDOSO "
    vgSql = vgSql & "(num_poliza,num_endoso,fec_solendoso,fec_endoso, "
    vgSql = vgSql & "cod_cauendoso,cod_tipendoso,mto_diferencia, "
    vgSql = vgSql & "cod_moneda, "
    vgSql = vgSql & "mto_pensionori,mto_pensioncal,fec_efecto, "
    vgSql = vgSql & "prc_factor,gls_observacion, "
    vgSql = vgSql & "cod_usuariocrea,fec_crea,hor_crea, "
    vgSql = vgSql & "cod_usuariomodi,fec_modi,hor_modi, "
    vgSql = vgSql & "fec_finefecto,cod_estado "
    vgSql = vgSql & ",cod_tipreajuste, mto_valreajustetri, mto_valreajustemen" 'hqr 03/12/2010
    vgSql = vgSql & " ) VALUES ( "
    vgSql = vgSql & "'" & vlNumPolizaE & "', "
    vgSql = vgSql & " " & str(vlNumEndosoE) & ", "
    vgSql = vgSql & "'" & vlFecSolEndoso & "', "
    vgSql = vgSql & "'" & vlFecEndoso & "', "
    vgSql = vgSql & "'" & vlCodCauEndoso & "', "
    vgSql = vgSql & "'" & vlCodTipEndoso & "', "
    vgSql = vgSql & " " & str(vlMtoDiferencia) & ", "
    vgSql = vgSql & "'" & vlCodMonedaEndoso & "', "
    vgSql = vgSql & " " & str(vlMtoPensionOri) & ", "
    vgSql = vgSql & " " & str(vlMtoPensionCal) & ", "
    vgSql = vgSql & "'" & vlFecEfecto & "', "
    vgSql = vgSql & " " & str(vlPrcFactor) & ", "
    vgSql = vgSql & "'" & vlObsEndoso & "', "
    vgSql = vgSql & "'" & vlCodUsuarioCreaE & "', "
    vgSql = vgSql & "'" & vlFecCreaE & "', "
    vgSql = vgSql & "'" & vlHorCreaE & "', "
    vgSql = vgSql & " NULL, "
    vgSql = vgSql & " NULL, "
    vgSql = vgSql & " NULL, "
    vgSql = vgSql & "'" & Trim(vlFecFinEfecto) & "',"
    vgSql = vgSql & "'" & Trim(vlCodEstadoE) & "', "
    'hqr 03/12/2010
    vgSql = vgSql & "'" & Trim(vlCodTipReajusteE) & "', "
    vgSql = vgSql & str(vlMtoValReajusteETri) & ", "
    vgSql = vgSql & str(vlMtoValReajusteEMen)
    'fin hqr 03/12/2010
    vgSql = vgSql & ")"
    vgConexionTransac.Execute vgSql

    Call FgGuardaLog(vgSql, vgUsuario, "6086")

Exit Function
Err_fgGeneraEndoso:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ] en poliza" & vlNumPolizaP, vbCritical, "¡ERROR!..."
    End Select
End Function

Function fgGeneraEndPol()
On Error GoTo Err_fgGeneraEndPol

    'Verificar existencia de regitro de poliza
    vgSql = ""
    vgSql = "SELECT num_poliza  FROM PP_TMAE_POLIZA "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "num_poliza = '" & vlNumPoliza & "' AND "
    vgSql = vgSql & "num_endoso = " & vlNumEndosoNuevo & " "
    'Set vgRegistro = vgConexionBD.Execute(vgSql)
    Set vgRegistro = vgConexionTransac.Execute(vgSql)
    If Not vgRegistro.EOF Then
        'Deshacer la Transacción
'        vgConexionTransac.RollbackTrans
        
         'Cerrar Conexión
        'vgConectarBD.Close
        
        MsgBox "El Nuevo Endoso a crear ya se encuentra Registrado en la Base de Datos." & Chr(13) & "Vuelva a comenzar nuevamente el proceso, realizando la búsqueda del caso para obtener su Nuevo Endoso." & vlNumPolizaP, vbCritical, "Proceso Cancelado"
        Cmd_Salir.SetFocus
        Screen.MousePointer = 0
        
        Exit Function
    End If
    vgRegistro.Close

    If vlNumPolizaP = "0000001017" Then
        a = 1
    End If

    vgSql = ""
    vgSql = "INSERT INTO pp_tmae_poliza "
    vgSql = vgSql & "(num_poliza,num_endoso,cod_afp,cod_tippension, "
    vgSql = vgSql & "cod_estado,cod_tipren,cod_modalidad,num_cargas, "
    vgSql = vgSql & "fec_vigencia,fec_tervigencia,mto_prima,mto_pension, "
    vgSql = vgSql & "num_mesdif,num_mesgar,prc_tasace,prc_tasavta, "
    vgSql = vgSql & "prc_tasactorea,prc_tasaintpergar,fec_inipagopen, "
    vgSql = vgSql & "cod_usuariocrea,fec_crea,hor_crea, "
    vgSql = vgSql & "cod_usuariomodi,fec_modi,hor_modi, "
    vgSql = vgSql & "cod_tiporigen,num_indquiebra, "
    'hqr 27/08/2007 campos agregados
    vgSql = vgSql & "mto_pensiongar,cod_cuspp,ind_cob, cod_moneda,"
    vgSql = vgSql & "mto_valmoneda,cod_cobercon,mto_facpenella,prc_facpenella,"
    vgSql = vgSql & "cod_dercre,cod_dergra,prc_tasatir,fec_emision,fec_dev,"
    vgSql = vgSql & "fec_inipencia,fec_pripago,fec_finperdif,fec_finpergar"
    vgSql = vgSql & ",fec_efecto" 'hqr 03/10/2007
    vgSql = vgSql & ",cod_tipreajuste,mto_valreajustetri, mto_valreajustemen" 'hqr 03/12/2010
    If vlFecDevSol <> "" Then
         vgSql = vgSql & ",fec_devsol" 'hqr 03/10/2007
    End If
    vgSql = vgSql & " ) VALUES ( "
    vgSql = vgSql & "'" & vlNumPolizaP & "', "
    vgSql = vgSql & " " & vlNumEndosoP & ", "
    vgSql = vgSql & "'" & vlCodAFP & "', "
    vgSql = vgSql & "'" & vlCodTipPension & "', "
    vgSql = vgSql & "'" & vlCodEstadoP & "', "
    vgSql = vgSql & "'" & vlCodTipRen & "', "
    vgSql = vgSql & "'" & vlCodModalidad & "', "
    vgSql = vgSql & " " & vlNumCargas & ", "
    vgSql = vgSql & "'" & vlFecVigencia & "', "
    vgSql = vgSql & "'" & vlFecTerVigencia & "', "
    vgSql = vgSql & " " & str(vlMtoPrima) & ", "
    vgSql = vgSql & " " & str(vlMtoPensionPol) & ", "
    vgSql = vgSql & " " & vlNumMesDif & ", "
    vgSql = vgSql & " " & vlNumMesGar & ", "
    vgSql = vgSql & " " & str(vlPrcTasaCe) & ", "
    vgSql = vgSql & " " & str(vlPrcTasaVta) & ", "
    vgSql = vgSql & " " & str(vlPrcTasaCtoRea) & ", "
    vgSql = vgSql & " " & str(vlPrcTasaIntPerGar) & ", "
    vgSql = vgSql & "'" & Trim(vlFecIniPagoPenPol) & "', "
    vgSql = vgSql & "'" & Trim(vlCodUsuarioCreaP) & "', "
    vgSql = vgSql & "'" & Trim(vlFecCreaP) & "', "
    vgSql = vgSql & "'" & Trim(vlHorCreaP) & "', "
    vgSql = vgSql & "'" & Trim(vgUsuario) & "', "
    vgSql = vgSql & "'" & Format(Date, "yyyymmdd") & "', "
    vgSql = vgSql & "'" & Format(Time, "hhmmss") & "', "
    vgSql = vgSql & "'" & Trim(vlCodTipOrigen) & "', "
    vgSql = vgSql & "" & (vlNumIndQuiebra) & " , "
    vgSql = vgSql & "" & str(vlMtoPensionGarP) & ", "
    vgSql = vgSql & "'" & Trim(vlCodCuspp) & "', "
    vgSql = vgSql & "'" & Trim(vlIndCob) & "', "
    vgSql = vgSql & "'" & Trim(vlCodMoneda) & "', "
    vgSql = vgSql & "" & str(vlMtoValMoneda) & ", "
    vgSql = vgSql & "'" & Trim(vlCodCobercon) & "', "
    vgSql = vgSql & "" & str(vlMtoFacPenElla) & ", "
    vgSql = vgSql & "" & str(vlPrcFacPenElla) & ", "
    vgSql = vgSql & "'" & Trim(vlCodDercreP) & "', "
    vgSql = vgSql & "'" & Trim(vlCodDerGra) & "', "
    vgSql = vgSql & "" & str(vlPrcTasaTir) & ", "
    vgSql = vgSql & "'" & Trim(vlFecEmision) & "', "
    vgSql = vgSql & "'" & Trim(vlFecDev) & "', "
    vgSql = vgSql & "'" & Trim(vlFecIniPenCia) & "', "
    vgSql = vgSql & "'" & Trim(vlFecPriPago) & "', "
    vgSql = vgSql & "'" & Trim(vlFecFinPerdif) & "', "
    vgSql = vgSql & "'" & Trim(vlFecFinPerGar) & "',"
    vgSql = vgSql & "'" & Trim(vlFecEfecto) & "'," 'hqr 03/10/2007
    'hqr 03/12/2010
    vgSql = vgSql & "'" & Trim(vlCodTipReajuste) & "',"
    vgSql = vgSql & str(vlMtoValReajusteTri) & ","
    vgSql = vgSql & str(vlMtoValReajusteMen)
    If vlFecDevSol <> "" Then
         vgSql = vgSql & "," & str(vlFecDevSol)
    End If
    'fin hqr 03/12/2010
    vgSql = vgSql & ")"
    vgConexionTransac.Execute vgSql
    
     Call FgGuardaLog(vgSql, vgUsuario, "6086")

    
    'Traspasa pensiones actualizadas
    vgSql = "INSERT INTO pp_tmae_pensionact"
    vgSql = vgSql & "(num_poliza, num_endoso, fec_desde,mto_pension,cod_tipopago,mto_pensiongar )"
    vgSql = vgSql & "SELECT a.num_poliza," & str(vlNumEndosoP) & ","
    vgSql = vgSql & "a.fec_desde,a.mto_pension,a.cod_tipopago,a.mto_pensiongar "
    vgSql = vgSql & "FROM pp_tmae_pensionact a "
    vgSql = vgSql & "WHERE a.num_poliza = '" & vlNumPolizaP & "' "
    vgSql = vgSql & "AND a.num_endoso = " & str(vlNumEndoso)
    vgConexionTransac.Execute (vgSql)
    
     Call FgGuardaLog(vgSql, vgUsuario, "6086")

Exit Function
Err_fgGeneraEndPol:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]" & " " & vlNumPolizaP, vbCritical, "¡ERROR!..."
    End Select
End Function

Function fgGeneraEndBeneficiarios()
On Error GoTo Err_fgGeneraEndBeneficiarios

    vgSql = ""
    vgSql = "INSERT INTO pp_tmae_ben "
    vgSql = vgSql & " SELECT "
    vgSql = vgSql & "num_poliza, " & vlNumEndosoNuevo & " ,num_orden,fec_ingreso,cod_tipoidenben, "
    vgSql = vgSql & "num_idenben,gls_nomben,gls_nomsegben, gls_patben,gls_matben,gls_dirben, "
    vgSql = vgSql & "cod_direccion,gls_fonoben,gls_correoben,cod_grufam, "
    vgSql = vgSql & "cod_par,cod_sexo,cod_sitinv,cod_dercre,cod_derpen, "
    vgSql = vgSql & "cod_cauinv,fec_nacben,fec_nachm,fec_invben, "
    vgSql = vgSql & "cod_motreqpen,"
    vgSql = vgSql & "mto_pension,mto_pensiongar,"
    vgSql = vgSql & "prc_pension, prc_pensionleg,"
    vgSql = vgSql & "cod_inssalud,cod_modsalud,mto_plansalud,cod_estpension, "
    vgSql = vgSql & "cod_viapago,cod_banco,cod_tipcuenta, "
    vgSql = vgSql & "num_cuenta,cod_sucursal,fec_fallben,fec_matrimonio, "
    vgSql = vgSql & "cod_caususben,fec_susben,fec_inipagopen,fec_terpagopengar, "
    vgSql = vgSql & "cod_usuariocrea,fec_crea,hor_crea, "
    vgSql = vgSql & "cod_usuariomodi,fec_modi,hor_modi, "
    vgSql = vgSql & "prc_pensiongar, gls_telben2 "
    vgSql = vgSql & "FROM pp_tmae_ben "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "num_poliza = '" & vlNumPoliza & "' AND "
    vgSql = vgSql & "num_endoso = " & vlNumEndoso & " "
    vgConexionTransac.Execute vgSql

Exit Function
Err_fgGeneraEndBeneficiarios:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fgGeneraEndBen()
On Error GoTo Err_fgGeneraEndBen
    
    vgSql = ""
    vgSql = "INSERT INTO pp_tmae_ben "
    vgSql = vgSql & "(num_poliza,num_endoso,num_orden,fec_ingreso,num_idenben, "
    vgSql = vgSql & "cod_tipoidenben,gls_nomben,gls_nomsegben,gls_patben,gls_matben,gls_dirben, "
    vgSql = vgSql & "cod_direccion,gls_fonoben,gls_correoben,cod_grufam, "
    vgSql = vgSql & "cod_par,cod_sexo,cod_sitinv,cod_dercre,cod_derpen, "
    vgSql = vgSql & "cod_cauinv,fec_nacben,fec_nachm,fec_invben, "
    vgSql = vgSql & "cod_motreqpen,mto_pension,mto_pensiongar,prc_pension, "
    vgSql = vgSql & "cod_inssalud,cod_modsalud,mto_plansalud,cod_estpension, "
    vgSql = vgSql & "cod_viapago,cod_banco,cod_tipcuenta, "
    vgSql = vgSql & "num_cuenta,cod_sucursal,fec_fallben,fec_matrimonio, "
    vgSql = vgSql & "cod_caususben,fec_susben,fec_inipagopen,fec_terpagopengar, "
    vgSql = vgSql & "cod_usuariocrea,fec_crea,hor_crea, "
    vgSql = vgSql & "cod_usuariomodi,fec_modi,hor_modi, "
    'vgSql = vgSql & "mto_plansalud2,cod_modsalud2,num_fun "
    vgSql = vgSql & "prc_pensionleg,prc_pensiongar "
    vgSql = vgSql & " ) VALUES ( "
    vgSql = vgSql & "'" & vlNumPolizaB & "', "
    vgSql = vgSql & " " & vlNumEndosoB & ", "
    vgSql = vgSql & " " & vlNumOrdenB & ", "
    vgSql = vgSql & "'" & Trim(vlFecIngreso) & "', "
    vgSql = vgSql & "'" & vlNumIdenBen & "', "
    vgSql = vgSql & " " & str(vlCodTipoIdenBen) & ", "
    vgSql = vgSql & "'" & vlGlsNomBen & "', "
    vgSql = vgSql & "'" & vlGlsNomSegBen & "', "
    vgSql = vgSql & "'" & vlGlsPatBen & "', "
    vgSql = vgSql & "'" & vlGlsMatBen & "', "
    vgSql = vgSql & "'" & Trim(vlGlsDirBen) & "', "
    vgSql = vgSql & " " & str(vlCodDireccion) & ", "
    vgSql = vgSql & "'" & Trim(vlGlsFonoBen) & "', "
    vgSql = vgSql & "'" & Trim(vlGlsCorreoBen) & "', "
    vgSql = vgSql & "'" & vlCodGruFam & "', "
    vgSql = vgSql & "'" & vlCodPar & "', "
    vgSql = vgSql & "'" & vlCodSexo & "', "
    vgSql = vgSql & "'" & vlCodSitInv & "', "
    vgSql = vgSql & "'" & vlCodDerCre & "', "
    vgSql = vgSql & "'" & vlCodDerpen & "', "
    vgSql = vgSql & "'" & vlCodCauInv & "', "
    vgSql = vgSql & "'" & vlFecNacBen & "', "
    vgSql = vgSql & "'" & vlFecNacHM & "', "
    vgSql = vgSql & "'" & vlFecInvBen & "', "
    vgSql = vgSql & "'" & vlCodMotReqPen & "', "
    vgSql = vgSql & " " & str(vlMtoPension) & ", "
    vgSql = vgSql & " " & str(vlMtoPensionGar) & ", "
    vgSql = vgSql & " " & str(vlPrcPension) & ", "
    vgSql = vgSql & "'" & Trim(vlCodInsSalud) & "', "
    vgSql = vgSql & "'" & Trim(vlCodModSalud) & "', "
    vgSql = vgSql & " " & str(vlMtoPlanSalud) & ", "
    vgSql = vgSql & "'" & vlCodEstPension & "', "
    vgSql = vgSql & "'" & Trim(vlCodViaPago) & "', "
    vgSql = vgSql & "'" & Trim(vlCodBanco) & "', "
    vgSql = vgSql & "'" & Trim(vlCodTipCuenta) & "', "
    vgSql = vgSql & "'" & Trim(vlNumCuenta) & "', "
    vgSql = vgSql & "'" & Trim(vlCodSucursal) & "', "
    vgSql = vgSql & "'" & vlFecFallBen & "', "
    vgSql = vgSql & "'" & Trim(vlFecMatrimonio) & "', "
    vgSql = vgSql & "'" & vlCodCauSusBen & "', "
    vgSql = vgSql & "'" & vlFecSusBen & "', "
    vgSql = vgSql & "'" & vlFecIniPagoPen & "', "
    vgSql = vgSql & "'" & vlFecTerPagoPenGar & "', "
    vgSql = vgSql & "'" & Trim(vlCodUsuarioCreaB) & "', "
    vgSql = vgSql & "'" & Trim(vlFecCreaB) & "', "
    vgSql = vgSql & "'" & Trim(vlHorCreaB) & "', "
    vgSql = vgSql & "'" & Trim(vgUsuario) & "', "
    vgSql = vgSql & "'" & Format(Date, "yyyymmdd") & "', "
    vgSql = vgSql & "'" & Format(Time, "hhmmss") & "', "
    vgSql = vgSql & " " & str(vlPrcPensionLeg) & ", "
    vgSql = vgSql & "" & str(vlPrcPensionGar) & ") "
    vgConexionTransac.Execute vgSql

Exit Function
Err_fgGeneraEndBen:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fgActualizaEndBen(iPoliza As String, iEndoso As Integer, iOrden As Integer, iEstPension As String)
On Error GoTo Err_fgActualizaEndBen

    vgSql = ""
    vgSql = "UPDATE pp_tmae_ben "
    vgSql = vgSql & " SET cod_estpension = '" & iEstPension & "', "
    'hqr 20/10/2007 Se deja la pensión actualizada en la tabla de beneficiarios
    'vgSql = vgSql & "mto_pension,mto_pensiongar,"
    vgSql = vgSql & "mto_pension = " & str(Format(vlMtoPensionPol * (vgRegPolizas!Prc_Pension / 100), "#0.00")) & ","
    If vgRegPolizas!Mto_PensionGar > 0 Then
        vgSql = vgSql & "mto_pensiongar = " & str(Format(vlMtoPensionPol * (vgRegPolizas!Prc_PensionGar / 100), "#0.00")) & " "
    Else
       vgSql = vgSql & "mto_pensiongar = 0 "
    End If
    'fin hqr 20/10/2007
    vgSql = vgSql & " WHERE "
    vgSql = vgSql & "num_poliza = '" & iPoliza & "' AND "
    vgSql = vgSql & "num_endoso = " & iEndoso & " AND "
    vgSql = vgSql & "num_orden = " & iOrden & " "
    vgConexionTransac.Execute vgSql

Exit Function
Err_fgActualizaEndBen:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function fgVerificarPreEndoso()
On Error GoTo Err_fgVerificarPreEndoso

    vgSql = ""
    vgSql = "SELECT num_endoso "
    vgSql = vgSql & "FROM pp_tmae_endoso "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "num_poliza = '" & Trim(vlNumPoliza) & "' AND "
    vgSql = vgSql & "cod_estado = '" & clCodEstadoP & "' "
    vgSql = vgSql & "ORDER BY num_endoso DESC "
    Set vgRegistro = vgConexionTransac.Execute(vgSql)
    If Not vgRegistro.EOF Then
       Call fgEliminarPreEndoso(Trim(vlNumPoliza), (vgRegistro!num_endoso))
    End If
    vgRegistro.Close
    
Exit Function
Err_fgVerificarPreEndoso:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fgEliminarPreEndoso(iNumPoliza As String, inumendoso As Integer)
On Error GoTo Err_fgEliminarPreEndoso

    'Eliminar Registros de Beneficiarios Pre-Endoso
    vgSql = ""
    vgSql = "DELETE pp_tmae_endben "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "num_poliza = '" & iNumPoliza & "' AND "
    vgSql = vgSql & "num_endoso = " & inumendoso & " "
    vgConexionBD.Execute (vgSql)

    'Eliminar Registro de Póliza de Pre-Endoso
    vgSql = ""
    vgSql = "DELETE pp_tmae_endpoliza "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "num_poliza = '" & iNumPoliza & "' AND "
    vgSql = vgSql & "num_endoso = " & inumendoso & " "
    vgConexionBD.Execute (vgSql)

    'Eliminar Pre-Endoso
    vgSql = ""
    vgSql = "DELETE FROM pp_tmae_endoso "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "num_poliza = '" & iNumPoliza & "' AND "
    vgSql = vgSql & "num_endoso = " & inumendoso & " AND "
    vgSql = vgSql & "cod_estado = '" & clCodEstadoP & "' "
    vgConexionBD.Execute vgSql

Exit Function
Err_fgEliminarPreEndoso:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fgActualizaEndAnterior(iPoliza As String)
On Error GoTo Err_fgActualizaEndAnterior
Dim vlFecFinEfectoAnt As String

    'Actualizar Datos de Fecha Fin Efecto del Endoso Anterior
    vgSql = ""
    vgSql = "SELECT num_poliza,num_endoso  FROM pp_tmae_endoso "
    vgSql = vgSql & "WHERE num_poliza = '" & iPoliza & "' AND "
    vgSql = vgSql & "num_endoso = " & (vlNumEndosoE - 1) & " "
   'Set vgRegistro = vgConexionBD.Execute(vgSql)
    Set vgRegistro = vgConexionTransac.Execute(vgSql)
    If Not vgRegistro.EOF Then
        If (vgRegistro!num_endoso) >= 1 Then
            'Fecha de Efecto del Ultimo Endoso, es decir, el Endoso actual
            vlFecEfecto = Trim(vlFecEfecto)
            vlAnno = Mid(vlFecEfecto, 1, 4)
            vlMes = Mid(vlFecEfecto, 5, 2)
            vlDia = Mid(vlFecEfecto, 7, 2)
            'Obtener fecha de fin de efecto del endoso anterior (ya grabado)
            vlFecFinEfectoAnt = DateSerial(vlAnno, vlMes, vlDia - 1)
            vlFecFinEfectoAnt = Format(CDate(Trim(vlFecFinEfectoAnt)), "yyyymmdd")
            'Actualizar fecha de fin de efecto del endoso anterior, ya grabado
            vgSql = ""
            vgSql = "UPDATE pp_tmae_endoso SET "
            vgSql = vgSql & "fec_finefecto = " & Trim(vlFecFinEfectoAnt) & " "
            vgSql = vgSql & "WHERE "
            vgSql = vgSql & "num_poliza = '" & Trim(iPoliza) & "' AND "
            vgSql = vgSql & "num_endoso = " & (vlNumEndosoE - 1) & " "
            vgConexionTransac.Execute vgSql
            
             Call FgGuardaLog(vgSql, vgUsuario, "6087")
        End If
    End If
    vgRegistro.Close

Exit Function
Err_fgActualizaEndAnterior:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fgCalculaFechaEfecto(iFechaEfectoIngresada As String) As String
Dim iFecha As String
Dim iFechaCierre As String
Dim iFechaEfecto As String

On Error GoTo Err_fgCalculaFechaEfecto
    
    fgCalculaFechaEfecto = ""
    If (Trim(iFechaEfectoIngresada) <> "") Then
    
        If Not IsDate(iFechaEfectoIngresada) Then
            Exit Function
        End If
        iFechaCierre = Format(fgValidaFechaEfecto(Trim(iFechaEfectoIngresada), vlNumPoliza, vlNumOrden), "yyyymmdd")
        iFechaEfecto = Format(CDate(iFechaEfectoIngresada), "yyyymmdd")
        If (iFechaCierre > iFechaEfecto) Then
            'MsgBox "La Fecha de Efecto es anterior a la Fecha del último cierre, el cual corresponde a :" & _
            DateSerial(CInt(Mid(iFechaCierre, 1, 4)), CInt(Mid(iFechaCierre, 5, 2)), CInt(Mid(iFechaCierre, 7, 2))), vbInformation, "Fecha Errónea"
            Exit Function
        Else
            fgCalculaFechaEfecto = DateSerial(CInt(Mid(iFechaEfecto, 1, 4)), CInt(Mid(iFechaEfecto, 5, 2)), CInt(Mid(iFechaEfecto, 7, 2)))
        End If
        
    Else
        
        'Determinar el menor periodo de Proceso que se encuentre Abierto
        'Determinar si el periodo a registrar es posterior al que se desea ingresar
        vgSql = ""
        vgSql = "SELECT num_perpago,cod_estadopri,cod_estadoreg "
        vgSql = vgSql & "FROM pp_tmae_propagopen "
        vgSql = vgSql & "WHERE "
        vgSql = vgSql & "cod_estadoreg <> 'C' or "
        vgSql = vgSql & "cod_estadopri <> 'C' "
        vgSql = vgSql & "ORDER BY num_perpago ASC"
        Set vgRs2 = vgConexionBD.Execute(vgSql)
        If Not vgRs2.EOF Then
            iFecha = DateSerial(CInt(Mid(vgRs2!Num_PerPago, 1, 4)), CInt(Mid(vgRs2!Num_PerPago, 5, 2)), 1)
        Else
            iFecha = fgBuscaFecServ
        End If
        vgRs2.Close
        fgCalculaFechaEfecto = fgValidaFechaEfecto(Trim(iFecha), vlNumPoliza, 1)
        
    End If

Exit Function
Err_fgCalculaFechaEfecto:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'----------------------------------------------------------------
Function fgBenSinCerEst() As Boolean
On Error GoTo Err_fgBenSinCerEst
'Función que permite modificar el estado de derecho a pensión
'de todos aquellos beneficiarios hijos, de entre 18 y 24 años
'de edad, que tienen certificado de estudio vigente
Dim vlAnno As String
Dim vlMes As String
Dim vlDia As String
Dim vlFechaCalculo As String
Dim vlFechaTermino As String
    


    vlFechaCalculo = vgFecPago '"20050720"
    fgBenSinCerEst = False
    vlFechaTermino = Format(DateAdd("d", -1, DateAdd("m", (vgMesesExtension + 1), DateSerial(Mid(vlFechaCalculo, 1, 4), Mid(vlFechaCalculo, 5, 2), 1))), "yyyymmdd")
    'vlFechaActual = fgBuscaFecServ
    vlFechaActual = vlFechaCalculo
    'vlFechaActual = Format(vlFechaActual, "yyyymmdd")
    'Fecha Con último dia del mes
    'vlFecha18 = DateSerial(Mid((vlFechaActual), 1, 4) - 18, Mid((vlFechaActual), 5, 2) + 1, 1 - 1)
    vlFecha18 = DateSerial(Mid((vlFechaActual), 1, 4) - 18, Mid((vlFechaActual), 5, 2), 0)
    vlFecha18 = Format(vlFecha18, "yyyymmdd")
    'Fecha Con primer dia del mes
    vlFecha24 = DateSerial(Mid((vlFechaActual), 1, 4) - 24, Mid((vlFechaActual), 5, 2), 1)
    vlFecha24 = Format(vlFecha24, "yyyymmdd")

    'Seleccionar todos aquellos beneficiarios con parentesco hijo
    'con edad entre 18 y 24 años que no tengan certificado de
    'estudios vigente
    
    'Selecciona los beneficiarios que no tengan certificado de supervivencia
    vgSql = ""
    vgSql = "SELECT " & vlCamposPoliza & " , " & vlCamposBen & " "
    vgSql = vgSql & "FROM pp_tmae_ben b, pp_tmae_poliza p "
    vgSql = vgSql & "WHERE " 'b.fec_nacben >= '" & vlFecha24 & "' AND "
    'vgSql = vgSql & "b.fec_nacben <= '" & vlFecha18 & "' AND "
    'vgSql = vgSql & "(b.cod_par = '" & clCodParHijo30 & "' OR "
    'vgSql = vgSql & "b.cod_par = '" & clCodParHijo35 & "') AND "
    'vgSql = vgSql & "b.cod_sitinv = '" & clCodSitInvN & "' AND "
    vgSql = vgSql & "NOT EXISTS "
    vgSql = vgSql & "(SELECT num_poliza FROM pp_tmae_certificado c "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "c.num_poliza = b.num_poliza AND "
    vgSql = vgSql & "c.num_orden = b.num_orden AND "
    vgSql = vgSql & "c.fec_inicer <= '" & vlFechaCalculo & "' AND "
    vgSql = vgSql & "c.fec_tercer >= '" & vlFechaTermino & "') AND "
    vgSql = vgSql & "b.cod_estpension = '" & clCodEstPension99 & "' AND "
    vgSql = vgSql & "b.num_poliza = p.num_poliza AND "
    vgSql = vgSql & "b.num_endoso = p.num_endoso AND "
    vgSql = vgSql & "p.num_endoso = (SELECT MAX (num_endoso) as numero "
    vgSql = vgSql & "FROM pp_tmae_poliza "
    vgSql = vgSql & "WHERE num_poliza = p.num_poliza) "
    'hqr 06/10/2007 Se agrega para que no le genere Endoso a las pólizas que no están pagando pensiones
    vgSql = vgSql & " AND p.cod_estado IN (6, 7, 8) " 'solo Pólizas Vigentes, con Beneficiarios Designados y Pendiente (Pago Diferido)
    'Se deben identificar las Pólizas que tengan su Primer Pago en un Periodo Anterior (Inmediatas y Diferidas)
    vgSql = vgSql & " AND (p.fec_pripago < '" & Format(vgFecIniPag, "yyyymmdd") & "'"
    vgSql = vgSql & " OR (p.num_mesdif > 0 AND p.fec_pripago = '" & Format(vgFecIniPag, "yyyymmdd") & "'))"
    'fin hqr 06/10/2007
    
    'Set vgRegPolizas = vgConexionBD.Execute(vgSql)
    Set vgRegPolizas = vgConexionTransac.Execute(vgSql) 'Se cambia conexion para que lea los endosos agregados en la función anterior
    
    If Not vgRegPolizas.EOF Then
        fgBenSinCerEst = True
    
        vlNumPoliza = ""
        vlNumPolAnterior = ""
        vlNumEndoso = 0
        vlNumOrden = 0
'        While Not vgRegPolizas.EOF
'            vgGeneraEndosos = True
'            If vgTipoPago = "P" Then
'                If Frm_PensPrimerosPagos.ProgressBar.Value + vlAumento < 100 Then
'                    Frm_PensPrimerosPagos.ProgressBar.Value = Frm_PensPrimerosPagos.ProgressBar.Value + vlAumento
'                    Frm_PensPrimerosPagos.ProgressBar.Refresh
'                End If
'            Else
'                If Frm_PensPagosRegimen.ProgressBar.Value + vlAumento < 100 Then
'                    Frm_PensPagosRegimen.ProgressBar.Value = Frm_PensPagosRegimen.ProgressBar.Value + vlAumento
'                    Frm_PensPagosRegimen.ProgressBar.Refresh
'                End If
'            End If
'
'            vlNumPoliza = (vgRegPolizas!num_poliza)
'            vlNumEndoso = (vgRegPolizas!num_endoso)
'            vlNumOrden = (vgRegPolizas!Num_Orden)
'
'            If Trim(vlNumPolAnterior) <> Trim(vgRegPolizas!num_poliza) Then
'
'                'Obtener numero de proximo endoso
'                vlNumUltEndoso = vlNumEndoso
'                'Asignar Número de Nuevo Endoso
'                vlNumEndosoNuevo = vlNumUltEndoso + 1
'
'                Call fgInicializaVarBen
'                Call fgInicializaVarPol
'                Call fgInicializaVarEnd
'                Call fgAsignarValoresPol
'
'                'Confirmar la No Existencia de Pre-Endoso
'                '(Si existe un pre-endoso, este se elimina para crear el nuevo endoso sin problemas)
'                Call fgVerificarPreEndoso
'
'                'Call fgInicializaVarEnd
'                Call fgAsignarValoresEnd
'                Call fgGeneraEndoso
'
'                Call fgActualizaEndAnterior(vlNumPoliza)
'
'                Call fgGeneraEndPol
'                'Insert Select de todos los beneficiarios de la poliza
'                Call fgGeneraEndBeneficiarios
'
'                'Actualiza el Estado de pension del Beneficiario seleccionado
'                Call fgActualizaEndBen(vlNumPoliza, vlNumEndosoNuevo, vlNumOrden, clCodEstPension20)
'
'                vlNumPolAnterior = vlNumPoliza
'
'            Else
'
'                'Actualiza el Estado de pension del Beneficiario seleccionado
'                Call fgActualizaEndBen(vlNumPoliza, vlNumEndosoNuevo, vlNumOrden, clCodEstPension20)
'
'                vlNumPolAnterior = vlNumPoliza
'
'            End If
'            vgRegPolizas.MoveNext
'        Wend
    Else
        If vgTipoPago = "P" Then
            If Frm_PensPrimerosPagos.ProgressBar.Value + vlAumento < 100 Then
                Frm_PensPrimerosPagos.ProgressBar.Value = Frm_PensPrimerosPagos.ProgressBar.Value + vlAumento
                Frm_PensPrimerosPagos.ProgressBar.Refresh
            End If
        Else
            If Frm_PensPagosRegimen.ProgressBar.Value + vlAumento < 100 Then
                Frm_PensPagosRegimen.ProgressBar.Value = Frm_PensPagosRegimen.ProgressBar.Value + vlAumento
                Frm_PensPagosRegimen.ProgressBar.Refresh
            End If
        End If
        
    End If

Exit Function
Err_fgBenSinCerEst:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
        fgBenSinCerEst = False
    End Select

End Function

'----------------------------------------------------------------
Function fgCargaEstructuraBenEndAut(istBenEndAuto() As TyBeneficiarios) As Boolean
'Carga de estructura de beneficiarios para endosos automaticos
On Error GoTo Err_fgCargaEstructuraBenEndAut

    fgCargaEstructuraBenEndAut = True
    vgSql = ""
    vgSql = "SELECT COUNT (num_orden) as numero "
    vgSql = vgSql & "FROM pp_tmae_ben WHERE "
    vgSql = vgSql & "num_poliza = '" & Trim(vlNumPoliza) & "' AND "
    vgSql = vgSql & "num_endoso = " & vlNumEndoso & " "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       vgNumBen = (vgRegistro!numero)
    End If

    ReDim istBenEndAuto(vgNumBen) ' As TyBeneficiarios

    vgSql = ""
    vgSql = "SELECT  "
    vgSql = vgSql & vlCamposBen & " "
    vgSql = vgSql & "FROM pp_tmae_ben b WHERE "
    vgSql = vgSql & "num_poliza = '" & Trim(vlNumPoliza) & "' AND "
    vgSql = vgSql & "num_endoso = " & vlNumEndoso & " "
    vgSql = vgSql & "ORDER BY num_orden ASC"
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       vgX = 0
       While Not vgRegistro.EOF
             vgX = vgX + 1
             With istBenEndAuto(vgX)
                  If IsNull(vgRegistro!num_poliza) Then
                     .num_poliza = ""
                  Else
                      .num_poliza = (vgRegistro!num_poliza)
                  End If
                  If IsNull(vgRegistro!num_endoso) Then
                     .num_endoso = ""
                  Else
                      .num_endoso = (vgRegistro!num_endoso)
                  End If
                  If IsNull(vgRegistro!Num_Orden) Then
                     .Num_Orden = ""
                  Else
                      .Num_Orden = (vgRegistro!Num_Orden)
                  End If
                  If IsNull(vgRegistro!Fec_Ingreso) Then
                     .Fec_Ingreso = ""
                  Else
                      .Fec_Ingreso = (vgRegistro!Fec_Ingreso)
                  End If
                  If IsNull(vgRegistro!Num_IdenBen) Then
                     .Num_IdenBen = ""
                  Else
                      .Num_IdenBen = (vgRegistro!Num_IdenBen)
                  End If
                  If IsNull(vgRegistro!Cod_TipoIdenBen) Then
                     .Cod_TipoIdenBen = ""
                  Else
                      .Cod_TipoIdenBen = (vgRegistro!Cod_TipoIdenBen)
                  End If
                  If IsNull(vgRegistro!Gls_NomBen) Then
                     .Gls_NomBen = ""
                  Else
                      .Gls_NomBen = (vgRegistro!Gls_NomBen)
                  End If
                  If IsNull(vgRegistro!Gls_NomSegBen) Then
                     .Gls_NomSegBen = ""
                  Else
                      .Gls_NomSegBen = (vgRegistro!Gls_NomSegBen)
                  End If
                  If IsNull(vgRegistro!Gls_PatBen) Then
                     .Gls_PatBen = ""
                  Else
                      .Gls_PatBen = (vgRegistro!Gls_PatBen)
                  End If
                  If IsNull(vgRegistro!Gls_MatBen) Then
                     .Gls_MatBen = ""
                  Else
                      .Gls_MatBen = (vgRegistro!Gls_MatBen)
                  End If
                  If IsNull(vgRegistro!Gls_DirBen) Then
                     .Gls_DirBen = ""
                  Else
                      .Gls_DirBen = (vgRegistro!Gls_DirBen)
                  End If
                  If IsNull(vgRegistro!Cod_Direccion) Then
                     .Cod_Direccion = ""
                  Else
                      .Cod_Direccion = (vgRegistro!Cod_Direccion)
                  End If
                  If IsNull(vgRegistro!Gls_FonoBen) Then
                     .Gls_FonoBen = ""
                  Else
                      .Gls_FonoBen = (vgRegistro!Gls_FonoBen)
                  End If
                  If IsNull(vgRegistro!Gls_CorreoBen) Then
                     .Gls_CorreoBen = ""
                  Else
                      .Gls_CorreoBen = (vgRegistro!Gls_CorreoBen)
                  End If
                  If IsNull(vgRegistro!Cod_GruFam) Then
                     .Cod_GruFam = ""
                  Else
                      .Cod_GruFam = (vgRegistro!Cod_GruFam)
                  End If
                  If IsNull(vgRegistro!Cod_Par) Then
                     .Cod_Par = ""
                  Else
                      .Cod_Par = (vgRegistro!Cod_Par)
                  End If
                  If IsNull(vgRegistro!Cod_Sexo) Then
                     .Cod_Sexo = ""
                  Else
                      .Cod_Sexo = (vgRegistro!Cod_Sexo)
                  End If
                  If IsNull(vgRegistro!Cod_SitInv) Then
                     .Cod_SitInv = ""
                  Else
                      .Cod_SitInv = (vgRegistro!Cod_SitInv)
                  End If
                  If IsNull(vgRegistro!Cod_DerCre) Then
                     .Cod_DerCre = ""
                  Else
                      .Cod_DerCre = (vgRegistro!Cod_DerCre)
                  End If
                  If IsNull(vgRegistro!Cod_DerPen) Then
                     .Cod_DerPen = ""
                  Else
                      .Cod_DerPen = (vgRegistro!Cod_DerPen)
                  End If
                  If IsNull(vgRegistro!Cod_CauInv) Then
                     .Cod_CauInv = ""
                  Else
                      .Cod_CauInv = (vgRegistro!Cod_CauInv)
                  End If
                  If IsNull(vgRegistro!Fec_NacBen) Then
                     .Fec_NacBen = ""
                  Else
                      .Fec_NacBen = (vgRegistro!Fec_NacBen)
                  End If
                  If IsNull(vgRegistro!Fec_NacHM) Then
                     .Fec_NacHM = ""
                  Else
                      .Fec_NacHM = (vgRegistro!Fec_NacHM)
                  End If
                  If IsNull(vgRegistro!Fec_InvBen) Then
                     .Fec_InvBen = ""
                  Else
                      .Fec_InvBen = (vgRegistro!Fec_InvBen)
                  End If
                  If IsNull(vgRegistro!Cod_MotReqPen) Then
                     .Cod_MotReqPen = ""
                  Else
                      .Cod_MotReqPen = (vgRegistro!Cod_MotReqPen)
                  End If
                  If IsNull(vgRegistro!pensionben) Then
                     .Mto_Pension = ""
                  Else
                      .Mto_Pension = (vgRegistro!pensionben)
                  End If
                  If IsNull(vgRegistro!Mto_PensionGar) Then
                     .Mto_PensionGar = ""
                  Else
                      .Mto_PensionGar = (vgRegistro!Mto_PensionGar)
                  End If
                  If IsNull(vgRegistro!Prc_Pension) Then
                     .Prc_Pension = ""
                  Else
                      .Prc_Pension = (vgRegistro!Prc_Pension)
                  End If
                  If IsNull(vgRegistro!Cod_InsSalud) Then
                     .Cod_InsSalud = ""
                  Else
                      .Cod_InsSalud = (vgRegistro!Cod_InsSalud)
                  End If
                  If IsNull(vgRegistro!Cod_ModSalud) Then
                     .Cod_ModSalud = ""
                  Else
                      .Cod_ModSalud = (vgRegistro!Cod_ModSalud)
                  End If
                  If IsNull(vgRegistro!Mto_PlanSalud) Then
                     .Mto_PlanSalud = ""
                  Else
                      .Mto_PlanSalud = (vgRegistro!Mto_PlanSalud)
                  End If
                  If IsNull(vgRegistro!Cod_EstPension) Then
                     .Cod_EstPension = ""
                  Else
                      .Cod_EstPension = (vgRegistro!Cod_EstPension)
                  End If
                  If IsNull(vgRegistro!Cod_ViaPago) Then
                     .Cod_ViaPago = ""
                  Else
                      .Cod_ViaPago = (vgRegistro!Cod_ViaPago)
                  End If
                  If IsNull(vgRegistro!Cod_Banco) Then
                     .Cod_Banco = ""
                  Else
                      .Cod_Banco = (vgRegistro!Cod_Banco)
                  End If
                  If IsNull(vgRegistro!Cod_TipCuenta) Then
                     .Cod_TipCuenta = ""
                  Else
                      .Cod_TipCuenta = (vgRegistro!Cod_TipCuenta)
                  End If
                  If IsNull(vgRegistro!Num_Cuenta) Then
                     .Num_Cuenta = ""
                  Else
                      .Num_Cuenta = (vgRegistro!Num_Cuenta)
                  End If
                  If IsNull(vgRegistro!Cod_Sucursal) Then
                     .Cod_Sucursal = ""
                  Else
                      .Cod_Sucursal = (vgRegistro!Cod_Sucursal)
                  End If
                  If IsNull(vgRegistro!Fec_FallBen) Then
                     .Fec_FallBen = ""
                  Else
                      .Fec_FallBen = (vgRegistro!Fec_FallBen)
                  End If
                  If IsNull(vgRegistro!Fec_Matrimonio) Then
                     .Fec_Matrimonio = ""
                  Else
                      .Fec_Matrimonio = (vgRegistro!Fec_Matrimonio)
                  End If
                  If IsNull(vgRegistro!Cod_CauSusBen) Then
                     .Cod_CauSusBen = ""
                  Else
                      .Cod_CauSusBen = (vgRegistro!Cod_CauSusBen)
                  End If
                  If IsNull(vgRegistro!Fec_SusBen) Then
                     .Fec_SusBen = ""
                  Else
                      .Fec_SusBen = (vgRegistro!Fec_SusBen)
                  End If
                  If IsNull(vgRegistro!fecinipagopen) Then
                     .Fec_IniPagoPen = ""
                  Else
                      .Fec_IniPagoPen = (vgRegistro!fecinipagopen)
                  End If
                  If IsNull(vgRegistro!Fec_TerPagoPenGar) Then
                     .Fec_TerPagoPenGar = ""
                  Else
                      .Fec_TerPagoPenGar = (vgRegistro!Fec_TerPagoPenGar)
                  End If
                  If IsNull(vgRegistro!Cod_UsuarioCreaben) Then
                     .Cod_UsuarioCrea = ""
                  Else
                      .Cod_UsuarioCrea = (vgRegistro!Cod_UsuarioCreaben)
                  End If
                  If IsNull(vgRegistro!Fec_Creaben) Then
                     .Fec_Crea = ""
                  Else
                      .Fec_Crea = (vgRegistro!Fec_Creaben)
                  End If
                  If IsNull(vgRegistro!Hor_Creaben) Then
                     .Hor_Crea = ""
                  Else
                      .Hor_Crea = (vgRegistro!Hor_Creaben)
                  End If
                  If IsNull(vgRegistro!Prc_PensionLeg) Then
                     .Prc_PensionLeg = 0
                  Else
                      .Prc_PensionLeg = (vgRegistro!Prc_PensionLeg)
                  End If
                  If IsNull(vgRegistro!Prc_PensionGar) Then
                     .Prc_PensionGar = 0
                  Else
                      .Prc_PensionGar = (vgRegistro!Prc_PensionGar)
                  End If
             End With
             vgRegistro.MoveNext
       Wend
    End If

Exit Function
Err_fgCargaEstructuraBenEndAut:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        fgCargaEstructuraBenEndAut = False
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fgBenMayor18() As Boolean
On Error GoTo Err_fgBenMayor18
Dim vlAnno As String
Dim vlMes As String
Dim vlDia As String
Dim vlFechaCalculo As String

    vlFechaCalculo = vgFecPago '"20050720"
    fgBenMayor18 = True
    vlFechaActual = vlFechaCalculo
    'Fecha Con primer dia del mes
    vlFecha18 = DateSerial(Mid((vlFechaActual), 1, 4) - 18, Mid((vlFechaActual), 5, 2), 1)
    vlFecha18 = Format(vlFecha18, "yyyymmdd")
    
    'Seleccionar todos aquellos beneficiarios con parentesco hijo
    'mayores de 18 años.
    '20130801
    L18 = 216
    
    vlSql = ""
    'vgSql = "SELECT " & vlCamposPoliza & " , " & vlCamposBen & " "
    'vgSql = vgSql & " FROM pp_tmae_poliza p JOIN pp_tmae_ben b ON p.num_poliza=b.num_poliza and p.num_endoso=b.num_endoso"
    'vgSql = vgSql & " WHERE p.num_endoso = (select max(num_endoso) from pp_tmae_poliza where num_poliza=p.num_poliza)"
    'vgSql = vgSql & " AND b.cod_estpension not in ('10','20') and b.cod_derpen not in (10,20) AND p.num_mesgar > 0 AND p.fec_finpergar >= '" & vgFecPago & "'"
    'vgSql = vgSql & " AND b.fec_nacben < to_Char(add_months(to_date('" & vgFecPago & "','yyyymmdd'),-"
    'vgSql = vgSql & " ("
    'vgSql = vgSql & "                 select case when (fec_devsol is null) or (to_date(fec_devsol, 'YYYYMMDD') < to_date('20130801', 'YYYYMMDD')) then " & CStr(L18) & " else case when ("
    'vgSql = vgSql & "                 select distinct case when IND_DNI='S' then 1 else 0 end  + case when IND_DJU='S' then 1 else 0 end  + case when IND_PES='S' then 1 else 0 end  + case when IND_BNO='S' then 1 else 0 end"
    'vgSql = vgSql & "                 from pp_tmae_certificado where num_poliza=a.num_poliza and num_endoso=a.num_endoso and num_orden=ben.num_orden"
    'vgSql = vgSql & "                 ) = 4 then " & CStr(L24) & " else " & CStr(L18) & " end end as edad_l"
    'vgSql = vgSql & "                 from pp_tmae_poliza a"
    'vgSql = vgSql & "                 join pp_tmae_ben ben on a.num_poliza=ben.num_poliza and a.num_endoso=ben.num_endoso"
    'vgSql = vgSql & "                 Where a.num_poliza = p.num_poliza And a.num_endoso = p.num_endoso And ben.Num_Orden = b.Num_Orden"
    'vgSql = vgSql & " )"
    'vgSql = vgSql & " ),'yyyymmdd') AND b.cod_par in ('30','35')"
    'vgSql = vgSql & " AND b.fec_inipagopen <= '" & vgFecPago & "' AND p.cod_estado IN (6, 7, 8) and cod_sitinv not in ('T', 'P')"
    'vgSql = vgSql & " AND (p.fec_pripago < '" & vgFecPago & "' OR (p.num_mesdif > 0 AND p.fec_pripago = '" & vgFecPago & "'))"
    
    vlSql = " SELECT " & vlCamposPoliza & " , " & vlCamposBen & " FROM pp_tmae_poliza p JOIN pp_tmae_ben b ON p.num_poliza=b.num_poliza and p.num_endoso=b.num_endoso"
    vlSql = vlSql & " WHERE p.num_endoso = (select max(num_endoso) from pp_tmae_poliza where num_poliza=p.num_poliza)"
    vlSql = vlSql & " AND b.cod_estpension not in ('10','20') and b.cod_derpen not in (10,20) AND p.num_mesgar > 0 AND p.fec_finpergar >= '" & vlFechaCalculo & "' " 'AND p.num_poliza=904"
    vlSql = vlSql & " AND b.fec_nacben < to_Char(add_months(to_date('" & vlFechaCalculo & "','yyyymmdd'),-"
    vlSql = vlSql & " ("
    vlSql = vlSql & "                 select case when (fec_devsol is null) or (to_date(fec_devsol, 'YYYYMMDD') < to_date('20130801', 'YYYYMMDD')) then " & CStr(L18) & " else case when ("
    vlSql = vlSql & "                 select distinct case when IND_DNI='S' then 1 else 0 end  + case when IND_DJU='S' then 1 else 0 end  + case when IND_PES='S' then 1 else 0 end  + case when IND_BNO='S' then 1 else 0 end"
    vlSql = vlSql & "                 from pp_tmae_certificado where num_poliza=a.num_poliza and num_endoso=a.num_endoso and num_orden=ben.num_orden and cod_tipo='EST' AND EST_ACT=1 AND (FEC_INICER<='" & vlFechaCalculo & "' AND FEC_TERCER>='" & vlFechaCalculo & "')"
    vlSql = vlSql & "                 ) = 4 then " & CStr(L24) & " else case when (to_date(fec_devsol, 'YYYYMMDD') < to_date('20130801', 'YYYYMMDD')) then " & CStr(L18) & " else " & CStr(L24) & " end end end as edad_l"
    vlSql = vlSql & "                 from pp_tmae_poliza a"
    vlSql = vlSql & "                 join pp_tmae_ben ben on a.num_poliza=ben.num_poliza and a.num_endoso=ben.num_endoso"
    vlSql = vlSql & "                 Where a.num_poliza = p.num_poliza And a.num_endoso = p.num_endoso And ben.Num_Orden = b.Num_Orden"
    vlSql = vlSql & " )"
    vlSql = vlSql & " ),'yyyymmdd') AND b.cod_par in ('30','35') "
    vlSql = vlSql & " AND b.fec_inipagopen <= '" & vlFechaCalculo & "' AND p.cod_estado IN (6, 7, 8) and cod_sitinv not in ('T', 'P')"
    vlSql = vlSql & " AND (p.fec_pripago < '" & vlFechaCalculo & "' OR (p.num_mesdif > 0 AND p.fec_pripago = '" & vlFechaCalculo & "'))"
    
    
    
'    vgSql = vgSql & " FROM pp_tmae_poliza p"
'    vgSql = vgSql & " JOIN pp_tmae_ben b ON p.num_poliza=b.num_poliza and p.num_endoso=b.num_endoso"
'    vgSql = vgSql & " join ma_tpar_tipoiden c on b.cod_tipoidenben=cod_tipoiden"
'    vgSql = vgSql & " WHERE p.num_endoso = (select max(num_endoso) from pp_tmae_poliza"
'    vgSql = vgSql & " where num_poliza=p.num_poliza) AND b.cod_estpension not in ('10','20') and b.cod_derpen not in ('10','20')"
'    vgSql = vgSql & " AND p.num_mesgar > 0 AND p.fec_finpergar >= '" & Format(vgFecIniPag, "yyyymmdd") & "' AND b.fec_nacben < to_Char(add_months(to_date('20131101','yyyymmdd'),-216),'yyyymmdd')"
'    vgSql = vgSql & " AND b.cod_par in ('30','35') AND b.fec_inipagopen <= '" & Format(vgFecIniPag, "yyyymmdd") & "'"
'    vgSql = vgSql & " AND p.cod_estado IN (6, 7, 8) and cod_sitinv not in ('T', 'P')"
'    vgSql = vgSql & " AND (p.fec_pripago < '" & Format(vgFecIniPag, "yyyymmdd") & "' OR (p.num_mesdif > 0 AND p.fec_pripago = '" & Format(vgFecIniPag, "yyyymmdd") & "'))"
'
'    vgSql = vgSql & "FROM pp_tmae_ben b, pp_tmae_poliza p "
'    vgSql = vgSql & "WHERE b.fec_nacben < '" & vlFecha18 & "' AND "
'    vgSql = vgSql & "(b.cod_par = '" & clCodParHijo30 & "' OR "
'    vgSql = vgSql & "b.cod_par = '" & clCodParHijo35 & "') AND "
'    vgSql = vgSql & "b.cod_sitinv = '" & clCodSitInvN & "' AND "
'    vgSql = vgSql & "(b.cod_estpension = '" & clCodEstPension20 & "' OR "
'    vgSql = vgSql & "b.cod_estpension = '" & clCodEstPension99 & "') AND "
'    vgSql = vgSql & "b.num_poliza = p.num_poliza AND "
'    vgSql = vgSql & "b.num_endoso = p.num_endoso AND "
'    vgSql = vgSql & "p.num_endoso = (SELECT MAX (num_endoso) as numero "
'    vgSql = vgSql & "FROM pp_tmae_poliza "
'    vgSql = vgSql & "WHERE num_poliza = p.num_poliza) "
'    'hqr 06/10/2007 Se agrega para que no le genere Endoso a las pólizas que no están pagando pensiones
'    vgSql = vgSql & " AND p.cod_estado IN (6, 7, 8) " 'solo Pólizas Vigentes, con Beneficiarios Designados y Pendiente (Pago Diferido)
'    'Se deben identificar las Pólizas que tengan su Primer Pago en un Periodo Anterior (Inmediatas y Diferidas)
'    vgSql = vgSql & " AND (p.fec_pripago < '" & Format(vgFecIniPag, "yyyymmdd") & "'"
'    vgSql = vgSql & " OR (p.num_mesdif > 0 AND p.fec_pripago = '" & Format(vgFecIniPag, "yyyymmdd") & "'))"
    'fin hqr 06/10/2007
    Set vgRegPolizas = vgConexionTransac.Execute(vlSql)
    If Not vgRegPolizas.EOF Then
        vlNumPoliza = ""
        vlNumPolAnterior = ""
        vlNumEndoso = 0
        vlNumOrden = 0
        While Not vgRegPolizas.EOF
            vgGeneraEndosos = True
            If Frm_PensPagosRegimen.ProgressBar.Value + vlAumento < 100 Then
                Frm_PensPagosRegimen.ProgressBar.Value = Frm_PensPagosRegimen.ProgressBar.Value + vlAumento
                Frm_PensPagosRegimen.ProgressBar.Refresh
            End If
    
            vlNumPoliza = (vgRegPolizas!num_poliza)
            vlNumEndoso = (vgRegPolizas!num_endoso)
            vlNumOrden = (vgRegPolizas!Num_Orden)
            
            If Trim(vlNumPolAnterior) <> Trim(vgRegPolizas!num_poliza) Then

                'Obtener numero de proximo endoso
                vlNumUltEndoso = vlNumEndoso
                'Asignar Número de Nuevo Endoso
                vlNumEndosoNuevo = vlNumUltEndoso + 1
                
                Call fgInicializaVarPol
                Call fgInicializaVarEnd
                Call fgAsignarValoresPol
                
                'Confirmar la No Existencia de Pre-Endoso
                '(Si existe un pre-endoso, este se elimina para crear el nuevo endoso sin problemas)
                Call fgVerificarPreEndoso
                
                'Call fgInicializaVarEnd
                Call fgAsignarValoresEnd
                Call fgGeneraEndoso
                
                Call fgActualizaEndAnterior(vlNumPoliza)
                
                Call fgGeneraEndPol
                
                
                If Not fgCargaEstructuraBenEndAut(stBenEndAuto) Then
                    fgBenMayor18 = False
                    Exit Function
                End If
                
                vgTipoCauEnd = "S"
                vgNum_pol = Trim(vlNumPoliza)
                If Not fgCalcularPorcentajeBenef(vlFecEfecto, vgNumBen, stBenEndAuto, vlCodTipPension, vlMtoPensionPol, True, vlCodDercreP, vlIndCob, True, vlNumMesGar, vlMtoPensionGarP) Then
                    'Call fgCalcularPorcentajeBenef(vlFecVigencia, vlNumCargas, stBeneficiariosMod, vlCodTipPension, vlMtoPensionRef, False, vlCodDerCrecerCot, vlIndCobertura, True, vlIsGar)
                    fgBenMayor18 = False
                    Exit Function
                End If
                Call fgCargaDatosBenEnd(stBenEndAuto)
                
                'Actualiza el Estado de pension del Beneficiario seleccionado
                'Call fgActualizaEndBen(vlNumPoliza, vlNumEndosoNuevo, vlNumOrden, clCodEstPension10)
                
                vlNumPolAnterior = vlNumPoliza
                
            Else
            
                'Actualiza el Estado de pension del Beneficiario seleccionado
                'Call fgActualizaEndBen(vlNumPoliza, vlNumEndosoNuevo, vlNumOrden, clCodEstPension10)
                
                vlNumPolAnterior = vlNumPoliza
                
            End If
            vgRegPolizas.MoveNext
        Wend
    Else
        If Frm_PensPagosRegimen.ProgressBar.Value + vlAumento < 100 Then
            Frm_PensPagosRegimen.ProgressBar.Value = Frm_PensPagosRegimen.ProgressBar.Value + vlAumento
            Frm_PensPagosRegimen.ProgressBar.Refresh
        End If
    End If

Exit Function
Err_fgBenMayor18:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
            fgBenMayor18 = False
            MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function fgCargaDatosBenEnd(istBenEndAuto() As TyBeneficiarios)
On Error GoTo Err_fgCargaDatosBenEnd
'Funcion que permite cargar los datos de los beneficiarios en variables,
'para asi finalmente generar el nuevo registro de endoso para el beneficiario
    vgX = 0
    While vgX < vgNumBen
          vgX = vgX + 1
          With istBenEndAuto(vgX)
          
                Call fgInicializaVarBen
          
                vlNumPolizaB = Trim(.num_poliza)
                'vlNumEndoso = .num_endoso
                vlNumEndosoB = vlNumEndosoNuevo
                vlNumOrdenB = .Num_Orden
                vlFecIngreso = .Fec_Ingreso
                vlCodTipoIdenBen = .Cod_TipoIdenBen
                vlNumIdenBen = .Num_IdenBen
                vlGlsNomBen = .Gls_NomBen
                vlGlsNomSegBen = .Gls_NomSegBen
                vlGlsPatBen = .Gls_PatBen
                vlGlsMatBen = .Gls_MatBen
                vlGlsDirBen = .Gls_DirBen
                vlCodDireccion = .Cod_Direccion
                vlGlsFonoBen = .Gls_FonoBen
                vlGlsCorreoBen = .Gls_CorreoBen
                vlCodGruFam = .Cod_GruFam
                vlCodPar = .Cod_Par
                vlCodSexo = .Cod_Sexo
                vlCodSitInv = .Cod_SitInv
                vlCodDerCre = .Cod_DerCre
                vlCodDerpen = .Cod_DerPen
                vlCodCauInv = .Cod_CauInv
                vlFecNacBen = .Fec_NacBen
                vlFecNacHM = .Fec_NacHM
                vlFecInvBen = .Fec_InvBen
                vlCodMotReqPen = .Cod_MotReqPen
                vlMtoPension = .Mto_Pension
                vlMtoPensionGar = .Mto_PensionGar
                vlPrcPension = .Prc_Pension
                vlCodInsSalud = .Cod_InsSalud
                vlCodModSalud = .Cod_ModSalud
                vlMtoPlanSalud = .Mto_PlanSalud
                vlCodEstPension = .Cod_EstPension
                vlCodViaPago = .Cod_ViaPago
                vlCodBanco = .Cod_Banco
                vlCodTipCuenta = .Cod_TipCuenta
                vlNumCuenta = .Num_Cuenta
                vlCodSucursal = .Cod_Sucursal
                vlFecFallBen = .Fec_FallBen
                vlFecMatrimonio = .Fec_Matrimonio
                vlCodCauSusBen = .Cod_CauSusBen
                vlFecSusBen = .Fec_SusBen
                vlFecIniPagoPen = .Fec_IniPagoPen
                vlFecTerPagoPenGar = .Fec_TerPagoPenGar
                vlCodUsuarioCreaB = .Cod_UsuarioCrea
                vlFecCreaB = .Fec_Crea
                vlHorCreaB = .Hor_Crea
                vlPrcPensionLeg = .Prc_PensionLeg
                vlPrcPensionGar = .Prc_PensionGar
                

                'Crea el nuevo Endoso para el beneficiario
                Call fgGeneraEndBen
               
          End With
    Wend

Exit Function
Err_fgCargaDatosBenEnd:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'CMV-20060601 F
Function fgObtenerPorcCobertura(iTipoPension As String, iFecha As String, oRemuneracion As Double, oPrcCobertura As Double) As Boolean
Dim vlRegCober As ADODB.Recordset

    fgObtenerPorcCobertura = False
    oRemuneracion = 0
    oPrcCobertura = 0

    vgSql = "SELECT mto_rem as remuneracion, prc_cober as porcentaje "
    vgSql = vgSql & "FROM ma_tcod_cober WHERE "
    vgSql = vgSql & "cod_tippension = '" & iTipoPension & "' "
    vgSql = vgSql & "AND (fec_inivig <= '" & iFecha & "' "
    vgSql = vgSql & "AND fec_tervig >= '" & iFecha & "') "
    Set vlRegCober = vgConexionBD.Execute(vgSql)
    If Not (vlRegCober.EOF) Then
        If Not IsNull(vlRegCober!remuneracion) Then oRemuneracion = (vlRegCober!remuneracion)
        If Not IsNull(vlRegCober!porcentaje) Then oPrcCobertura = (vlRegCober!porcentaje)
        
        fgObtenerPorcCobertura = True
    End If
    vlRegCober.Close

End Function

Public Function fgCalcularFechaFinPerDiferido(iFecDev As String, iMesDif As Long) As String
'Permite determinar la Fecha de Termino del Periodo Diferido
'Parámetros de Entrada:
'- iMesDif      => Número de Meses Diferidos
'- iFechaDev    => Fecha de Devengue
'Parámetros de Salida:
'- Retorna      => Fecha Fin Periodo Diferido "yyyymmdd"
'------------------------------------------------------
'Fecha de Creación     : 07/07/2007
'Fecha de Modificación : 09/08/2007
'------------------------------------------------------
    
    fgCalcularFechaFinPerDiferido = ""
    
    If (iMesDif > 0) Then
        fgCalcularFechaFinPerDiferido = Format(DateSerial(Mid(iFecDev, 1, 4), Mid(iFecDev, 5, 2) + iMesDif, 1 - 1), "yyyymmdd")
    End If
    
End Function

Public Function fgCalcularFechaFinPerGarantizado(iFecDev As String, iMesDif As Long, iMesGar As Long) As String
'Permite determinar la Fecha de Termino del Periodo Garantizado
'Parámetros de Entrada:
'- iMesDif      => Número de Meses Diferidos
'- iMesGar      => Número de Meses Garantizados
'- iFechaDev    => Fecha de Devengue
'Parámetros de Salida:
'- Retorna      => Fecha Fin Periodo Garantizado "yyyymmdd"
'------------------------------------------------------
'Fecha de Creación     : 07/07/2007
'Fecha de Modificación : 09/08/2007
'------------------------------------------------------
Dim vlFecha As String

    fgCalcularFechaFinPerGarantizado = ""
    
    If (iMesGar > 0) Then
        If (iMesDif > 0) Then
            vlFecha = Format(DateSerial(Mid(iFecDev, 1, 4), Mid(iFecDev, 5, 2) + iMesDif, 1), "yyyymmdd")
        Else
            vlFecha = iFecDev
        End If
        fgCalcularFechaFinPerGarantizado = Format(DateSerial(Mid(vlFecha, 1, 4), Mid(vlFecha, 5, 2) + iMesGar, 1 - 1), "yyyymmdd")
    End If

End Function

Function fgCalcularFechaIniPagoPensiones(iFecDev As String, iMesDif As Long) As String
'Permite determinar la Fecha de Inicio de Pago de Pensiones
'Parámetros de Entrada:
'- iMesDif      => Número de Meses Diferidos
'- iFechaDev    => Fecha de Devengue
'Parámetros de Salida:
'- Retorna      => Fecha de Inicio de Pago de Pensiones "yyyymmdd"
'-----------------------------------------------------------------
'Fecha de Creación     : 09/08/2007
'Fecha de Modificación :
'------------------------------------------------------
Dim vlFecha As String

    fgCalcularFechaIniPagoPensiones = ""
    
    If (iMesDif > 0) Then
        fgCalcularFechaIniPagoPensiones = Format(DateSerial(Mid(iFecDev, 1, 4), Mid(iFecDev, 5, 2) + iMesDif, 1), "yyyymmdd")
    Else
        fgCalcularFechaIniPagoPensiones = iFecDev
    End If

End Function

Function fgBuscarRegistroPensionActualizada(iNumPoliza As String, iFecIniPag As String, oMtoPensionAct As Double, oFecEfePol As String) As Boolean

    fgBuscarRegistroPensionActualizada = False
    
    vlMonedaPensionAnt = ""
    
''''    vlSql = "SELECT NVL(c.mto_pension, a.mto_pension) AS pensionact"
''''    vlSql = vlSql & ",c.fec_desde "
''''    vlSql = vlSql & " FROM pp_tmae_poliza a, pp_tmae_pensionact c "
''''    vlSql = vlSql & " WHERE a.num_poliza = c.num_poliza (+)"
''''    vlSql = vlSql & " AND a.num_endoso = c.num_endoso (+)"
''''    vlSql = vlSql & " AND a.num_poliza = '" & iNumPoliza & "'"
''''    vlSql = vlSql & " AND (c.fec_desde = "
''''        vlSql = vlSql & " (SELECT max(fec_desde)"
''''        vlSql = vlSql & " FROM pp_tmae_pensionact"
''''        vlSql = vlSql & " WHERE num_poliza = a.num_poliza"
''''        vlSql = vlSql & " AND num_endoso = a.num_endoso"
''''        vlSql = vlSql & " AND fec_desde < '" & iFecIniPag & "')"
''''    vlSql = vlSql & " OR c.fec_desde IS NULL)    "
''''    vlSql = vlSql & " AND a.num_endoso ="
''''        vlSql = vlSql & " (SELECT NVL(MAX(b.num_endoso + 1),1) AS num_endoso" 'Último Endoso
''''        vlSql = vlSql & " FROM pp_tmae_endoso b"
''''        vlSql = vlSql & " WHERE b.num_poliza = a.num_poliza"
''''        vlSql = vlSql & " AND b.fec_efecto <= '" & iFecIniPag & "'"
''''        vlSql = vlSql & " AND b.fec_finefecto >= '" & iFecIniPag & "'"
''''        vlSql = vlSql & " AND b.cod_estado = 'E')"
''    '*ABV vlSql = vlSql & " AND a.cod_estado IN (6, 7, 8) " 'solo Pólizas Vigentes, con Beneficiarios Designados y Pendiente (Pago Diferido)
''    'Se deben identificar las Pólizas que tengan su Primer Pago en un Periodo Anterior (Inmediatas y Diferidas)
''    ''vlSql = vlSql & " AND a.fec_inipagopen < '" & Format(iFecIniPag, "yyyymmdd") & "'"
''    '*ABV vlSql = vlSql & " AND (a.fec_pripago < '" & Format(iFecIniPag, "yyyymmdd") & "'"
''    '*ABV vlSql = vlSql & " OR (a.num_mesdif > 0 AND a.fec_pripago = '" & Format(iFecIniPag, "yyyymmdd") & "'))"
'''    vlSql = vlSql & " ORDER BY a.cod_moneda,c.num_endoso "

    'hqr 11/10/2007 se saca tabla pp_tmae_poliza de la consulta
    vlSql = "SELECT a.mto_pension AS pensionact"
    vlSql = vlSql & ",a.fec_desde "
    vlSql = vlSql & " FROM pp_tmae_pensionact a "
    vlSql = vlSql & " WHERE a.num_poliza = '" & iNumPoliza & "'"
    vlSql = vlSql & " AND a.fec_desde = "
        vlSql = vlSql & " (SELECT max(fec_desde)"
        vlSql = vlSql & " FROM pp_tmae_pensionact b"
        vlSql = vlSql & " WHERE b.num_poliza = a.num_poliza"
        vlSql = vlSql & " AND b.num_endoso = a.num_endoso"
        vlSql = vlSql & " AND b.fec_desde < '" & iFecIniPag & "')"
    vlSql = vlSql & " AND a.num_endoso ="
        vlSql = vlSql & " (SELECT NVL(MAX(c.num_endoso + 1),1) AS num_endoso" 'Último Endoso
        vlSql = vlSql & " FROM pp_tmae_endoso c"
        vlSql = vlSql & " WHERE c.num_poliza = a.num_poliza"
        vlSql = vlSql & " AND c.fec_efecto <= '" & iFecIniPag & "'"
        vlSql = vlSql & " AND c.fec_finefecto >= '" & iFecIniPag & "'"
        vlSql = vlSql & " AND c.cod_estado = 'E')"
    Set vlTB = vgConexionBD.Execute(vlSql)
    If Not vlTB.EOF Then
        oMtoPensionAct = vlTB!Pensionact
        If Not IsNull(vlTB!fec_desde) Then
            oFecEfePol = vlTB!fec_desde
        End If
        fgBuscarRegistroPensionActualizada = True
    End If
    vlTB.Close
    
End Function

Function fgCalcularPensionActualizada(iCodMoneda As String, iNumMesDif As Integer, iFecPriPago As String, iFecIniPenCia As String, iFecEfectoPol As String, iFecDevengue As String, iMtoPensionAct As Double, iFecEfecto As String, iTipoAjuste, iMontoAjusteTri, iMontoAjusteMen, iFecVigencia) As Double
Dim vlFactorAjuste As Double, vlFactorAjusteDif As Double, vlFactorAjusteDif2 As Double
Dim vlPrimerMesAjuste As Double
Dim vlFechaIniDif As Date, vlFechaFinDif As Date
Dim vlMontoPensionAct As Double
Dim vlMesDif As Integer, vlAñoDif As Integer
Dim vlNumPension As Double 'hqr 26/02/2011
Dim vlPasaTrimestre As Boolean 'hqr 26/02/2011

    vlFactorAjuste = 1
    vgFecIniPag = DateSerial(Mid(iFecEfecto, 1, 4), Mid(iFecEfecto, 5, 2), Mid(iFecEfecto, 7, 2))
    'vgFecIniPag = DateSerial(Mid(iFecIniPenCia, 1, 4), Mid(iFecIniPenCia, 5, 2), Mid(iFecIniPenCia, 7, 2))

    '********************************************************************
    'Actualizacion Pension Diferida
    '********************************************************************
    'Obtiene Monto de la Pensión Actualizada
    'Si es primer pago de una diferida, se debe obtener la pensión actualizada
    
    vlPrimerMesAjuste = True
    'If iCodMoneda = vgMonedaCodOfi Then 'Nuevos Soles
    If iTipoAjuste <> cgSINAJUSTE Then 'hqr 12/01/2011
        'Solo si está en Nuevos Soles se hace la actualización
        'desde la Fecha de Devengamiento
        'If (iNumMesDif > 0 And iFecPriPago = Format(vgFecIniPag, "yyyymmdd")) Then
        'If (iNumMesDif > 0 And iFecPriPago = iFecIniPenCia And iFecPriPago = Format(vgFecIniPag, "yyyymmdd")) Then
'        If (iNumMesDif > 0) Then
            'Obtiene Factores de actualización
            vlFactorAjusteDif = 1 'Sin Ajuste
            
            'hqr  05/03/2011
            If clAjusteDesdeFechaDevengamiento Then
                vlFecDesdeAjustePension = Mid(iFecDevengue, 1, 6) & "01" 'Fecha de Devengamiento, se deja en dia 01 para que el while de las vueltas correctas
            Else
                vlFecDesdeAjustePension = iFecVigencia 'Fecha de Inicio de Vigencia de la Póliza
            End If
            'fin hqr  05/03/2011
            
            If (iFecEfectoPol <> "") Then
                vlFechaIniDif = DateSerial(Mid(iFecEfectoPol, 1, 4), Mid(iFecEfectoPol, 5, 2) + 1, Mid(iFecEfectoPol, 7, 2))
            Else
                'vlFechaIniDif = DateSerial(Mid(iFecDevengue, 1, 4), Mid(iFecDevengue, 5, 2) + 1, Mid(iFecDevengue, 7, 2))
                'vlFechaIniDif = DateSerial(Mid(iFecVigencia, 1, 4), Mid(iFecVigencia, 5, 2) + 1, Mid(iFecVigencia, 7, 2)) 'hqr 25/02/2011
                vlFechaIniDif = DateSerial(Mid(vlFecDesdeAjustePension, 1, 4), Mid(vlFecDesdeAjustePension, 5, 2) + 1, Mid(vlFecDesdeAjustePension, 7, 2)) 'hqr 25/02/2011
            End If
            'vlFechaFinDif = DateAdd("m", -1, vgFecIniPag)
            vlFechaFinDif = vgFecIniPag
            vlMontoPensionAct = iMtoPensionAct
            vlNumPension = 2 'hqr 26/02/2011
            vlPasaTrimestre = False 'hqr 26/02/2011
            Do While vlFechaFinDif >= vlFechaIniDif
                vlMesDif = Month(vlFechaIniDif)
                vlAñoDif = Year(vlFechaIniDif)
                If (((vlMesDif Mod 3) - 1) = 0) Then
                    'Obtiene Factor de Ajuste Anterior
                    'inicio hqr 03/12/2010
                    If ((vlAñoDif <> Mid(vlFecDesdeAjustePension, 1, 4) Or vlMesDif <> Mid(vlFecDesdeAjustePension, 5, 2))) Then 'No se ajusta el primer mes
                        If iTipoAjuste = cgAJUSTETASAFIJA Then
                            If (vlNumPension >= 2 And vlNumPension <= 3) And (Not (vlPasaTrimestre)) Then
                                vlFactorAjusteDif = (1 + (iMontoAjusteMen / 100))
                            Else
                                vlFactorAjusteDif = (1 + (iMontoAjusteTri / 100))
                            End If
                        Else
                            If vlPrimerMesAjuste Then
                                 vlFechaAjuste = Format(DateSerial(Year(vlFechaIniDif), Month(vlFechaIniDif) - 4, 1), "yyyymmdd")
                                 'Obtener Factor anterior
                                 If Not fgObtieneFactorAjuste(vlFechaAjuste, vlFactorAjusteDif) Then
                                     MsgBox "No se encuentra Factor de Ajuste del Periodo: " & DateSerial(Mid(vlFechaAjuste, 1, 4), Mid(vlFechaAjuste, 5, 2), Mid(vlFechaAjuste, 7, 2)), vbCritical, "Error de Datos"
                                     Exit Function
                                 End If
                             End If
                             vlPrimerMesAjuste = False
                        
                            'Obtiene Factor de Ajuste Actual
                            vlFechaAjuste = Format(DateSerial(Year(vlFechaIniDif), Month(vlFechaIniDif) - 1, 1), "yyyymmdd")
                            'Factor de Ajuste Mes actual
                            If Not fgObtieneFactorAjuste(vlFechaAjuste, vlFactorAjusteDif2) Then
                                MsgBox "No se encuentra Factor de Ajuste del Periodo: " & DateSerial(Mid(vlFechaAjuste, 1, 4), Mid(vlFechaAjuste, 5, 2), Mid(vlFechaAjuste, 7, 2)), vbCritical, "Error de Datos"
                                Exit Function
                            End If
                            vlFactorAjusteDif = Format(vlFactorAjusteDif2 / vlFactorAjusteDif, "##0.0000")
                            vlFecDesdeAjuste = vlFechaIniDif
                        End If
                        vlMontoPensionAct = Format(vlFactorAjusteDif * vlMontoPensionAct, "#0.00") 'La última Actualización
                        If iTipoAjuste <> cgAJUSTETASAFIJA Then
                            vlFactorAjusteDif = vlFactorAjusteDif2
                        End If
                        vlPasaTrimestre = True 'hqr 14/02/2011
                    End If
                Else 'No está en mes del trimestre
                    If (iTipoAjuste = cgAJUSTETASAFIJA) And (vlNumPension >= 2 And vlNumPension <= 3) And (Not (vlPasaTrimestre)) Then
                        vlFactorAjusteDif = (1 + (iMontoAjusteMen / 100))
                        vlMontoPensionAct = Format(vlFactorAjusteDif * vlMontoPensionAct, "#0.00") 'La última Actualización
                    End If
                End If
                vlFechaIniDif = DateAdd("m", 1, vlFechaIniDif)
                vlNumPension = vlNumPension + 1 'hqr 26/02/2011
            Loop
            
'            'Graba último Ajuste
'            vlSql = "INSERT INTO pp_tmae_pensionact "
'            vlSql = vlSql & "(num_poliza, num_endoso, fec_desde, mto_pension, cod_tipopago) "
'            vlSql = vlSql & "VALUES ("
'            vlSql = vlSql & "'" & stLiquidacion.Num_Poliza & "', " & stLiquidacion.num_endoso & ", "
'            vlSql = vlSql & "'" & Format(vlFecDesdeAjuste, "yyyymmdd") & "', " & Str(vlMontoPensionAct) & ",'" & stLiquidacion.Cod_TipoPago & "')"
'            vgConexionTransac.Execute (vlSql)
            
            'Se actualiza con factor del mes actual
            ''vlMontoPensionAct = Format(vlFactorAjuste * vlMontoPensionAct, "#0.00") 'La última Actualización
'        Else
'            If (iFecEfectoPol <> "") Then 'Para que el primer mes del endoso no actualice pension
'                If iFecEfectoPol = Format(vgFecIniPag, "yyyymmdd") Then
'                    vlMontoPensionAct = Format(iMtoPensionAct, "#0.00")
'                Else
'                    vlMontoPensionAct = Format(vlFactorAjuste * iMtoPensionAct, "#0.00")
'                End If
'            Else
'                vlMontoPensionAct = Format(vlFactorAjuste * iMtoPensionAct, "#0.00")
'            End If
'        End If
    Else
        vlMontoPensionAct = Format(iMtoPensionAct, "#0.00")
    End If
    '********************************************************************
    
    fgCalcularPensionActualizada = vlMontoPensionAct

End Function

Function fgObtenerFechaCotizacion(iNumPoliza As String, oFecCotizacion As String) As Boolean
'Permite obtener la Fecha de Cotización o de Cálculo de la Póliza, la cual se obtiene
'desde las Tablas correspondientes al Módulo de Producción
Dim vlRegFecha As ADODB.Recordset

    fgObtenerFechaCotizacion = False
    oFecCotizacion = ""

    vgSql = "SELECT fec_calculo as fec_cot "
    vgSql = vgSql & "FROM pd_tmae_poliza WHERE "
    vgSql = vgSql & "num_poliza = '" & iNumPoliza & "' "
    Set vlRegFecha = vgConexionBD.Execute(vgSql)
    If Not (vlRegFecha.EOF) Then
        If Not IsNull(vlRegFecha!fec_cot) Then oFecCotizacion = (vlRegFecha!fec_cot)
        
        fgObtenerFechaCotizacion = True
'I - Borrar
'    Else
'        oFecCotizacion = "20070801"
'        fgObtenerFechaCotizacion = True
'F - Borrar
    End If
    vlRegFecha.Close

End Function







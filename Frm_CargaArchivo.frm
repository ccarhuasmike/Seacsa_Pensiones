VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Frm_CargaArchivo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carga de Archivo"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   6045
   Begin VB.Frame Fra_Datos 
      Height          =   2295
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   5775
      Begin VB.TextBox Txt_FecProceso 
         Height          =   285
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox Txt_Desde 
         Height          =   285
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox Txt_Hasta 
         Height          =   285
         Left            =   4320
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1440
         Width           =   1215
      End
      Begin VB.ComboBox Cmb_Tipo 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   2775
      End
      Begin VB.ComboBox Cmb_Pago 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   880
         Width           =   2775
      End
      Begin VB.Label Lbl_FecProceso 
         Caption         =   "Fecha Proceso               :"
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
         Left            =   360
         TabIndex        =   14
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   4080
         TabIndex        =   13
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Lbl_Contrato 
         Caption         =   "Período (Desde - Hasta)  :"
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
         Index           =   1
         Left            =   360
         TabIndex        =   12
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Lbl_Contrato 
         Caption         =   "Tipo de Proceso             :"
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
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Lbl_Contrato 
         Caption         =   "Tipo de Pago                 :"
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
         Left            =   360
         TabIndex        =   10
         Top             =   880
         Width           =   2415
      End
   End
   Begin VB.Frame Fra_Operaciones 
      Height          =   1095
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   2300
      Width           =   5775
      Begin VB.CommandButton Cmd_Cargar 
         Caption         =   "&Archivo"
         Height          =   675
         Left            =   1200
         Picture         =   "Frm_CargaArchivo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Exportar Datos a Archivo"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   3840
         Picture         =   "Frm_CargaArchivo.frx":0822
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   2520
         Picture         =   "Frm_CargaArchivo.frx":091C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Width           =   720
      End
      Begin MSComDlg.CommonDialog ComDialogo 
         Left            =   480
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "Frm_CargaArchivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vlSql As String

Dim vlDiasCotizados As String
Dim vlRutCliente As String
Dim vlDgvCliente As String
Dim vlRutRecauda As String
Dim vlDgvRecauda As String
Dim vlFechaDesde As String
Dim vlFechaHasta As String
Dim vlOpcion As String
Dim vlPago As String
Dim vlGlosaOpcion As String
Dim vlFechaInicio As String
Dim vlFechaTermino As String
Dim vlRutReceptor As String
Dim vlDgvReceptor As String

'Dim vlNombre As String
Dim vlPaterno As String
Dim vlMaterno As String

Dim vlCodEstado As String

'Variables Banco
Dim vlRutEmpresa As String
Dim vlDgvEmpresa As String
Dim vlCodConvenio As String
Dim vlFillerEncabezado1 As String
Dim vlFechaGeneracion As String
Dim vlFechaProceso As String
Dim vlHoraGeneracion As String
Dim vlFillerEncabezado2 As String

Dim vlTipoReg As String
Dim vlRut As String
Dim vlDigito As String
Dim vlNombre As String
Dim vlFiller1 As String
Dim vlIndicadorCta As String
Dim vlNumCta11 As String
Dim vlFiller2 As String
Dim vlFiller3 As String
Dim vlTipoMov As String
Dim vlCodBco As String
Dim vlNumCta15 As String
Dim vlMonto As String
Dim vlFiller4 As String

Dim vlNum As Integer
Dim vlMto As Double

Dim vlMtoProcesar As String
Dim vlNumReg As String
Dim vlFillerControl As String

'Declaración de variables para proceso de Cheque Contable
Dim vlCodEmpresa As String
Dim vlCodSistema As String
Dim vlNumSolicitud As String
Dim vlNumCorrelativo As String
Dim vlCodSucursal As String
Dim vlFechaPago As String
Dim vlCentroCosto As String
Dim vlConPago As String
Dim vlDoctoPago As String
Dim vlNumDocto As String
Dim vlFechaDocto As String
Dim vlCodLugar As String
Dim vlRutRec As String
Dim vlCodMonedaOri As String
Dim vlCodMonedaPago As String
Dim vlCodImpto As String
Dim vlMtoNetoMO As String
Dim vlMtoDesctoMO As String
Dim vlMtoAfectoMO As String
Dim vlMtoExentoMO As String
Dim vlMtoImptoMO As String
Dim vlMtoNotaMO As String
Dim vlMtoTotalPagarMO As String
Dim vlMtoNetoMP As String
Dim vlMtoDesctoMP As String
Dim vlMtoAfectoMP As String
Dim vlMtoExentoMP As String
Dim vlMtoImptoMP As String
Dim vlMtoNotaMP As String
Dim vlMtoTotalPagarMP As String
Dim vlCuentaConAC As String
Dim vlCentroCostoAC As String
Dim vlCodProdAC As String
Dim vlCodAuxAC As String
Dim vlDoctoConAC As String
Dim vlGlosaAC As String
Dim vlTipoMovAC As String
Dim vlMtoAsigMO As String
Dim vlMtoAsigMN As String

'Variables para valores provenientes de tabla ma_tcod_general
Dim vlCodConEmpresa As String
Dim vlCodConSistema As String
Dim vlCodConSucursal As String
Dim vlCodConCenCosto As String
Dim vlCodConConPago As String
Dim vlCodConDocPago As String
Dim vlCodConMonOri As String
Dim vlCodConMonPago As String
Dim vlCodConImpuesto As String
Dim vlCodConCtaCon As String
Dim vlCodConDocCon As String
Dim vlCodConSenMov As String

Dim vlContCorr As Integer
Dim vlNumPerPago As String
Dim vlSwSinReg As Boolean

'***********************************************************************
'Variables para generación de archivo de centralizacion contable
'Variables de Centralización Contable

Dim vlRegistroConcepto As ADODB.Recordset
Dim vlRegistroMtoConcepto As ADODB.Recordset
Dim vlRegistroFechas As ADODB.Recordset
Dim vlRegistroAux As ADODB.Recordset

Dim vlLinea As String
Dim vllinea2 As String
Dim vlArchivo As String, vlOpen As Boolean
Dim vlarchivo2 As String, vlopen2 As Boolean
Dim vlContador As Long
Dim vlAumento As Integer

'Montos totales por tipo de pensión y Concepto de Hab./Des.
Dim vlMtoTotalVN1 As String
Dim vlMtoTotalVA2 As String
Dim vlMtoTotalIT3 As String
Dim vlMtoTotalSO4 As String
Dim vlMtoTotalSO1005 As String
Dim vlMtoTotalSOVN6 As String
Dim vlMtoTotalSOVA7 As String
Dim vlMtoTotalSOIT8 As String
Dim vlMtoTotalSOIP9 As String
Dim vlMtoTotalIP10 As String
Dim vlMtoTotalOTRO11 As String

'Montos totales por tipo de pensión para concepto H/D 01 = Pensión Renta Vitalicia
Dim vlMtoVNRV As String
Dim vlMtoVARV As String
Dim vlMtoITRV As String
Dim vlMtoSORV As String
Dim vlMtoSO100RV As String
Dim vlMtoSOVNRV As String
Dim vlMtoSOVARV As String
Dim vlMtoSOITRV As String
Dim vlMtoSOIPRV As String
Dim vlMtoIPRV As String
Dim vlMtoOTRORV As String

'Montos totales por tipo de pensión para concepto H/D 05 = Pensión RetroActiva
Dim vlMtoVNRetro As String
Dim vlMtoVARetro As String
Dim vlMtoITRetro As String
Dim vlMtoSORetro As String
Dim vlMtoSO100Retro As String
Dim vlMtoSOVNRetro As String
Dim vlMtoSOVARetro As String
Dim vlMtoSOITRetro As String
Dim vlMtoSOIPRetro As String
Dim vlMtoIPRetro As String
Dim vlMtoOTRORetro As String

'Estructura para guardar conceptos de haberes y descuentos
Private Type TyCodConHabDes
    CodConHabDes    As String
    GlsCtaCon       As String
    CodTipMov       As String
    CodCtaCon       As String
    GlsConHabDes    As String
End Type

'Registro de Códigos de Conceptos de Haberes y Descuentos
Private stCodConHabDes() As TyCodConHabDes

Dim vlFecPago As String
Dim vlFechaActual As String
Dim vlTipoPago As String

Dim vlCodConHabDes As String
Dim vlGlsCtaCon As String
Dim vlCodTipMov As String
Dim vlCodCtaCon As String
Dim vlGlsConHabDes As String
Dim vlMtoTotalGral As String

Dim vlNumRegistros As String
Dim vlMtoMonedaUF As String

'Variables para linea de registro de pensión de Renta Vitalicia
Dim vlGlsCtaConRV As String
Dim vlGlsConHabDesRV As String
Dim vlCodConHabDesRV As String
Dim vlMtoTotalGralRV As String

'Variables para linea de registro de pensión retroactiva
Dim vlGlsCtaConRetro As String
Dim vlGlsConHabDesRetro As String
Dim vlCodConHabDesRetro As String
Dim vlMtoTotalGralRetro As String

Const clCodConHabDes01 As String * 2 = "01"
Const clCodConHabDes05 As String * 2 = "05"

Const clCodViaPagoConve As String * 2 = "03"
Const clGlosaAC As String = "Pago Pensiones "

Const clTipPenVN As String = "('04')"
Const clTipPenVA As String = "('05')"
Const clTipPenIT As String = "('06')"
Const clTipPenIP As String = "('07')"
Const clTipPenSO As String = "('01','03','08','13','15')"
Const clTipPenSO100 As String = "('')"
Const clTipPenSOVN As String = "('09')"
Const clTipPenSOVA As String = "('10')"
Const clTipPenSOIT As String = "('11')"
Const clTipPenSOIP As String = "('12')"
Const clTipPenOTRO As String = "('02','14')"

'***********************************************************************

'CONSTANTES
'-------------------------------------------------------
'Calidad del Trabajador 1: Activo 2: Pasivo
Const clCalTrab As String * 1 = "2"
'Tipo de Institución A: Afp C: Cía. de Seguros M: Mutual S: Servicio de Salud P:Capredena
Const clCodIns As String * 1 = "C"
'Cógigo de tabla, en tabla ma_tpar_tabcod de Instituciones de Salud
Const clCodInsSalud As String * 2 = "IS"
'Código de elemento, en tabla ma_tpar_tabcod de Institución Recaudadora Fonasa
Const clCodInstRecauda As String * 2 = "13"
'Código de Concepto de Haber y Descuento de Cotización de Salud
Const clCodConHabDes24 As String * 2 = "24"
'Código de Via de Pago Depósito en Cuenta
Const clCodViaPago02 As String * 2 = "02"

'CMV-20050131 I
'Códigos de Vias de Pago Depositos en Banco
Const clCodViaPagoBco As String = "('02','03')"
'CMV-20050131 F

'Código de Tipo de Movimiento, 22: Chequera Electronica
Const clCodTipoMov As String * 2 = "22"
'Código de Tipo de Registro (Dato enviado en archivo)
Const clTipoReg1 As String * 2 = "1"
'Código de Tipo de Registro (Dato enviado en archivo)
Const clTipoReg2 As String * 2 = "2"
'Código de Tipo de Registro (Dato enviado en archivo)
Const clTipoReg3 As String * 2 = "3"
'Código de elemento, en tabla ma_tpar_tabcod de Código de Convenio de Banco
Const clCodConBco12 As String * 2 = "12"

Const clCodMonedaNS As String * 2 = "NS"

''''Códigos de Tipo de Receptor
''''T: Tutor
'''Const clCodReceptorT As String * 1 = "T"
''''P: Pensionado
'''Const clCodReceptorP As String * 1 = "P"
''''R: Retenedor
'''Const clCodReceptorR As String * 1 = "R"
''''M: Madre de Hijo (En Caso de Sobrevivencia)
'''Const clCodReceptorM As String * 1 = "M"

'------------------- VARIABLES GENERACION ARCHIVO CONTABLE ---------------
'Variables
'Dim vlCodEmpresa As String
Dim vlCodComprobante As String
Dim vlCtaContable As String
Dim vlCodCentroCosto As String
Dim vlCodRamo As String
Dim vlCodAux As String
Dim vlCodDoctoCont As String
Dim vlCodAux2 As String
'Dim vlNumDocto As String
Dim vlFecDocto As String
Dim vlGlsDetalle As String
Dim vlCodMoneda As String
Dim vlFecOperacion As String
Dim vlMetroCubico As String
Dim vlDebeMO As String
Dim vlHaberMO As String
Dim vlDebeMN As String
Dim vlHaberMN As String
Dim vlFecIniDocto As String
Dim vlFecTerDocto As String
Dim vlEspacios As String

Dim vlCtaContableGral As String
Dim vlCodMonedaPESOS As String
Dim vlCodMonedaUF As String
Dim vlValorUF As Double
Dim vlFechaUF As String
'Dim vlFechaPago As String

Dim vlMtoEnt As String
Dim vlMtoDec As String

'Constantes
Const clConHabDes0105 As String = "('01','05')"
Const clConHabDes01 As String * 2 = "01"
Const clConHabDes05 As String * 2 = "05"
Const clCodComprobanteT As String * 1 = "T"
Const clCodTipMovD As String * 1 = "D"
Const clCodTipMovH As String * 1 = "H"
Const clGlsDetTodos As String = "Pago Pensiones"

'----------- Generacion archivo previred

'Dim vlRut As String
Dim vlDgv As String
'Dim vlPaterno As String
'Dim vlMaterno As String
Dim vlNombres As String
Dim vlTipoReceptor As String
'Dim vlTipoPago As String
Dim vlPeriodo As String
Dim vlRentaImp As String
Dim vlCotFonasa As String
Dim vlCodInsSalud As String
Dim vlMonedaPlanIsapre As String
Dim vlCotIsapre As String
Dim vlCotAdicional As String
Dim vlOtrosAportes As String
Dim vlCotPactada As String
Dim vlMtoHabDes24 As String
Dim vlTotalPagoIsapre As String
Dim vlFUN As String
Dim vlCodCCAF As String
Dim vlAporteCCAF As String
Dim vlAporteAdicionalCCAF As String
Dim vlCreditoCCAF As String
Dim vlConvDentalCCAF As String
Dim vlLeassingCCAF As String
Dim vlSeguroCCAF As String
Dim vlOtrosCCAF As String

Dim vlCodConHabDesSA As String

Dim vlCotPactada1 As String
Dim vlCotPactada2 As String

'Dim vlValorUF As Double

Dim vlNumPoliza As String
Dim vlNumOrden As String
Dim vlCodOtrosHabDes As String

Dim vlCodCCAFORI As String

Dim vlCodTipReceptor As String

Dim vlDia As String
Dim vlMes As String
Dim vlAnno As String

Const clCodParCau As String * 2 = "99"
Const clCodCausante As String * 1 = "1"
Const clCodBen As String * 1 = "2"

Const clModSaludNS As String * 5 = "NS"
Const clModSaludUS As String * 2 = "US"
Const clModSaludPORCE As String * 5 = "PORCE"

Const clCodInsSalud07 As String * 2 = "07"

Const clCodTipReceptorR As String * 1 = "R"

Const clCodModOrigenSA As String * 2 = "SA"
Const clCodModOrigenCCAF As String * 4 = "CCAF"
'CMV-20060302 I
'Modificacion solicitada por Sr.: Isaias Pizarro
'No se debe considerar código correspondiente a Préstamos Médicos
'Const clCodHabDesSA As String = "('24','28','29','34')"
Const clCodHabDesSA As String = "('24','28','29')"
'CMV-20060302 F
Const clCodHabDesSAAux As String = "('24','28','34')"
Const clCodHabDesCCAF As String = "('25','26')"
Const clCodHabDes24 As String * 2 = "24"
Const clCodHabDes25 As String * 2 = "25"
Const clCodHabDes26 As String * 2 = "26"
'Const clCodHabDes27 As String * 2 = "27"
Const clCodConHabDes29 As String * 2 = "29"

Const clCodPrcSalud As String = "PS"
Dim vlPrcSalud As Double

'****************** VARIABLES ARCHIVO PRESTMOS MEDICOS ***********

Dim vlMtoConHabDes As String
Dim vlNomCompleto As String
Dim vlIndNoPago As String

Const clIndNoPago0 As String * 1 = "0"
Const clIndNoPago3 As String * 1 = "3"

'*********************VARIABLES ARCHIVO A EXCEL
Dim vlTipoPension As String 'columna 1
Dim vlNumPagoAño As Integer 'Columna 2
Dim vlNumPagoMes As String 'Columna 3
Dim vlNumOrden2 As Integer 'Columna4
Dim vlTipoIdentif As Integer 'columna5
Dim vlNumIdenBenef As String 'columna6
Dim vlMatBen As String
Dim vlPatBen As String
Dim vlFecNac As String
Dim vlSexo As Integer
Dim vlFecIngreso As String
Dim vlTelefono As String
Dim vlFchBaj As String
Dim vlRuceps As String
Dim vldireccion As String
Dim vlInteri As String
Dim vlNomZon As String
Dim vlRefEre As String
Dim vlTipVia As String
Dim vlTipZon As String
Dim vlUbiGeo As String
Dim vlMtoConHab As Double
Dim vlNomCompania As String
Dim vlCuspp As String
Dim vlNumPoliza2 As String
Dim vlNomCausante As String
Dim vlMoneda As String
''Dim vlNomBen As String
Dim vlFchAfp As String
Dim vlNumDireccion As String
Dim vlNombreComuna As String

'**********************CONSTANTES ARCHIVO A EXCEL
Const clArcFonasaSittra As Integer = "11" 'Columna 14
Const clArcFonasaTiptra As Integer = "24" 'Columna 15
Const clArcFonasaEssvid As Integer = "0" 'Columna 18
Const clArcFonasaRegpen As Integer = "1" 'Columna 19
Const clArcFonasaSctr_1 As Integer = "0" 'Columna 20
Const clArcFonasaDiatra As Integer = "30" 'Columna 29
Const clColumna1 As String = "CODPLA"
Const clColumna2 As String = "ANOPRO"
Const clColumna3 As String = "MESPRO"
Const clColumna4 As String = "CODPER"
Const clColumna5 As String = "TIPDOC"
Const clColumna6 As String = "NRODOC"
Const clColumna7 As String = "APPPER"
Const clColumna8 As String = "APMPER"
Const clColumna9 As String = "NOMPER"
Const clColumna10 As String = "FCHNAC"
Const clColumna11 As String = "SEXPER"
Const clColumna12 As String = "TLFPER"
Const clColumna13 As String = "FCHING"
Const clColumna14 As String = "SITTRA"
Const clColumna15 As String = "TIPTRA"
Const clColumna16 As String = "FCHBAJ"
Const clColumna17 As String = "RUCEPS"
Const clColumna18 As String = "ESSVID"
Const clColumna19 As String = "REGPEN"
Const clColumna20 As String = "SCTR_1"
Const clColumna21 As String = "NOMVIA"
Const clColumna22 As String = "NUMERO"
Const clColumna23 As String = "INTERI"
Const clColumna24 As String = "NOMZON"
Const clColumna25 As String = "REFERE"
Const clColumna26 As String = "TIPVIA"
Const clColumna27 As String = "TIPZON"
Const clColumna28 As String = "UBIGEO"
Const clColumna29 As String = "DIATRA"
Const clColumna30 As String = "REMSAL"
Const clColumna31 As String = "FCHAFP"
Const clColumna32 As String = "NOMREN"
Const clColumna33 As String = "NOMCIA"
Const clColumna34 As String = "CUSPP"
Const clColumna35 As String = "POLIZA"
Const clColumna36 As String = "NOMCAU"

Private Sub Cmb_Pago_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Txt_Desde.SetFocus
End If
End Sub

Private Sub Cmb_Tipo_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Cmb_Pago.SetFocus
End If
End Sub

Private Sub Cmd_Cargar_Click()
On Error GoTo Err_ExportarDatos

    If Trim(Txt_Desde) = "" Then
        MsgBox "Falta Ingresar Fecha Desde ", vbCritical, "Falta Información"
        Txt_Desde.SetFocus
        Exit Sub
    Else
        If Not IsDate(Txt_Desde) Then
            MsgBox "Fecha Desde no es una Fecha válida", vbCritical, "Falta Información"
            Txt_Desde.SetFocus
            Exit Sub
        End If
    End If
    If Trim(Txt_Hasta) = "" Then
        MsgBox "Falta Ingresar Fecha Hasta", vbCritical, "Falta Información"
        Txt_Hasta.SetFocus
        Exit Sub
    Else
        If Not IsDate(Txt_Hasta) Then
            MsgBox "Fecha Hasta no es una fecha Válida", vbCritical, "Falta Información"
            Txt_Hasta = ""
            Exit Sub
        End If
    End If
    
    vlOpcion = ""
    vlPago = ""
    vlGlosaOpcion = ""
    vlOpcion = Trim(Mid(Cmb_Tipo, 1, InStr(1, Cmb_Tipo, " - ") - 1))
    vlPago = Trim(Mid(Cmb_Pago, 1, InStr(1, Cmb_Pago, " - ") - 1))
    Select Case vlOpcion
        Case "D": 'Definitivo
            vlGlosaOpcion = "DEF"
        Case "P": 'Provisorio
            vlGlosaOpcion = "PROV"
    End Select
    
'Permite imprimir la Opción Indicada a través del Menú
    Select Case vgNomInfSeleccionado
        Case "InfGeneraArchCotFonasa"    'Genera Archivo de Cotizaciones Fonasa
            Call flExportarCotFonasa
'        Case "InfGeneraArchPagoBco"    'Genera Archivo de Pagos a Bancos (Deposito Bancario)
'            Call flExportarPagosBanco 'BECH
'        Case "InfGeneraArchChequeCon"    'Genera Archivo de Cheque Contable
'            Call flExportarChequeContable
''        Case "InfGeneraArchCenCont"    'Genera Archivo de Centralización Contable
'            Call flExportarCenContable
'        Case "InfGeneraArchContable"    'Genera Archivo Contable
'            Call flExportarArchContable
'        Case "InfGeneraArchPrevired"    'Genera Archivo Pagos Previsionales
'            Call flExportarArchPrevired
'        Case "InfGeneraArchPresMed"    'Genera Archivo de Préstamos Médicos
'            Call flExportarArchPresMed
    End Select
    
Exit Sub
Err_ExportarDatos:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Limpiar_Click()
On Error GoTo Err_Limpiar

    Txt_Desde = ""
    Txt_Hasta = ""
    Txt_Desde.SetFocus
    If Cmb_Tipo.ListCount <> 0 Then
        Cmb_Tipo.ListIndex = 0
    End If
    If Cmb_Pago.ListCount <> 0 Then
        Cmb_Pago.ListIndex = 0
    End If

Exit Sub
Err_Limpiar:
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

Private Sub Form_Load()
On Error GoTo Err_Carga
    
    Frm_CargaArchivo.Left = 0
    Frm_CargaArchivo.Top = 0
    fgComboTipoCalculo Cmb_Tipo
    fgComboTipoPension Cmb_Pago
    
Exit Sub
Err_Carga:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Txt_Desde_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    If (Txt_Desde <> "") Then
        If Not IsDate(Trim(Txt_Desde)) Then
            MsgBox "La Fecha ingresada no es una fecha válida.", vbCritical, "Error de Datos"
            Txt_Desde.SetFocus
            Exit Sub
        End If
        If Txt_Hasta <> "" Then
            If (CDate(Trim(Txt_Desde)) > CDate(Trim(Txt_Hasta))) Then
                MsgBox "La Fecha de Término de Perido es mayor a la fecha de Inicio", vbCritical, "Error de Datos"
                Exit Sub
            End If
        End If
        If (Year(CDate(Trim(Txt_Desde))) < 1900) Then
            MsgBox "Error en la Fecha ingresada es menor a la mínima fecha que se puede ingresar (1900).", vbCritical, "Error de Datos"
            Txt_Desde.SetFocus
            Exit Sub
        End If
        Txt_Desde.Text = Format(CDate(Trim(Txt_Desde)), "yyyymmdd")
        vlFechaDesde = Trim(Txt_Desde)
        Txt_Desde.Text = DateSerial(Mid((Txt_Desde.Text), 1, 4), Mid((Txt_Desde.Text), 5, 2), Mid((Txt_Desde.Text), 7, 2))
    End If
Txt_Hasta.SetFocus
End If

End Sub

Private Sub Txt_Desde_LostFocus()

If (Txt_Desde <> "") Then
    If Not IsDate(Trim(Txt_Desde)) Then
        Exit Sub
    End If
    If Txt_Hasta <> "" Then
        If (CDate(Trim(Txt_Desde)) > CDate(Trim(Txt_Hasta))) Then
            Exit Sub
        End If
    End If
    If (Year(CDate(Trim(Txt_Desde))) < 1900) Then
        Exit Sub
    End If
    Txt_Desde.Text = Format(CDate(Trim(Txt_Desde)), "yyyymmdd")
    vlFechaDesde = Trim(Txt_Desde)
    Txt_Desde.Text = DateSerial(Mid((Txt_Desde.Text), 1, 4), Mid((Txt_Desde.Text), 5, 2), Mid((Txt_Desde.Text), 7, 2))
End If

End Sub

Private Sub Txt_FecProceso_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    If (Txt_FecProceso <> "") Then
        If Not IsDate(Trim(Txt_FecProceso)) Then
            MsgBox "La Fecha ingresada no es una fecha válida.", vbCritical, "Error de Datos"
            Txt_FecProceso.SetFocus
            Exit Sub
        End If
'        If Txt_Hasta <> "" Then
'            If (CDate(Trim(Txt_Desde)) > CDate(Trim(Txt_Hasta))) Then
'                MsgBox "La Fecha de Término de Perido es mayor a la fecha de Inicio", vbCritical, "Error de Datos"
'                Exit Sub
'            End If
'        End If
        If (Year(CDate(Trim(Txt_FecProceso))) < 1900) Then
            MsgBox "Error en la Fecha ingresada es menor a la mínima fecha que se puede ingresar (1900).", vbCritical, "Error de Datos"
            Txt_FecProceso.SetFocus
            Exit Sub
        End If
        Txt_FecProceso.Text = Format(CDate(Trim(Txt_FecProceso)), "yyyymmdd")
        vlFechaProceso = Trim(Txt_FecProceso)
        Txt_FecProceso.Text = DateSerial(Mid((Txt_FecProceso.Text), 1, 4), Mid((Txt_FecProceso.Text), 5, 2), Mid((Txt_FecProceso.Text), 7, 2))
    End If
    Cmd_Cargar.SetFocus
End If

End Sub

Private Sub Txt_FecProceso_LostFocus()
If (Txt_FecProceso <> "") Then
    If Not IsDate(Trim(Txt_FecProceso)) Then
        Exit Sub
    End If
'    If Txt_Hasta <> "" Then
'        If (CDate(Trim(Txt_Desde)) > CDate(Trim(Txt_Hasta))) Then
'            Exit Sub
'        End If
'    End If
    If (Year(CDate(Trim(Txt_FecProceso))) < 1900) Then
        Exit Sub
    End If
    Txt_FecProceso.Text = Format(CDate(Trim(Txt_FecProceso)), "yyyymmdd")
    vlFechaProceso = Trim(Txt_FecProceso)
    Txt_FecProceso.Text = DateSerial(Mid((Txt_FecProceso.Text), 1, 4), Mid((Txt_FecProceso.Text), 5, 2), Mid((Txt_FecProceso.Text), 7, 2))
End If
End Sub

Private Sub Txt_Hasta_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    If (Txt_Hasta <> "") Then
        If Not IsDate(Trim(Txt_Hasta)) Then
            MsgBox "La Fecha ingresada no es una fecha válida.", vbCritical, "Error de Datos"
            Txt_Hasta.SetFocus
            Exit Sub
        End If
        If Txt_Desde <> "" Then
            If (CDate(Trim(Txt_Desde)) > CDate(Trim(Txt_Hasta))) Then
                MsgBox "La Fecha de Término de Perido es mayor a la fecha de Inicio", vbCritical, "Error de Datos"
                Exit Sub
            End If
        End If
        If (Year(CDate(Trim(Txt_Hasta))) < 1900) Then
            MsgBox "Error en la Fecha ingresada es menor a la mínima fecha que se puede ingresar (1900).", vbCritical, "Error de Datos"
            Txt_Hasta.SetFocus
            Exit Sub
        End If
        Txt_Hasta.Text = Format(CDate(Trim(Txt_Hasta)), "yyyymmdd")
        vlFechaHasta = Trim(Txt_Hasta)
        Txt_Hasta.Text = DateSerial(Mid((Txt_Hasta.Text), 1, 4), Mid((Txt_Hasta.Text), 5, 2), Mid((Txt_Hasta.Text), 7, 2))
    End If
    If Txt_FecProceso.Visible = True Then
        Txt_FecProceso.SetFocus
    Else
        Cmd_Cargar.SetFocus
    End If
End If

End Sub

Private Sub Txt_Hasta_LostFocus()

If (Txt_Hasta <> "") Then
    If Not IsDate(Trim(Txt_Hasta)) Then
        Exit Sub
    End If
    If Txt_Desde <> "" Then
        If (CDate(Trim(Txt_Desde)) > CDate(Trim(Txt_Hasta))) Then
            Exit Sub
        End If
    End If
    If (Year(CDate(Trim(Txt_Hasta))) < 1900) Then
        Exit Sub
    End If
    Txt_Hasta.Text = Format(CDate(Trim(Txt_Hasta)), "yyyymmdd")
    vlFechaHasta = Trim(Txt_Hasta)
    Txt_Hasta.Text = DateSerial(Mid((Txt_Hasta.Text), 1, 4), Mid((Txt_Hasta.Text), 5, 2), Mid((Txt_Hasta.Text), 7, 2))
End If
End Sub

Function flValidaEstadoProceso(iPeriodo As String, iCodEstado As String) As Boolean
On Error GoTo Err_flValidaTipoProceso

    flValidaEstadoProceso = False

    vgSql = ""
    vgSql = "SELECT p.cod_estadoreg " ',p.cod_estadopri "
    vgSql = vgSql & "FROM pp_tmae_propagopen p "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "p.num_perpago = '" & Trim(iPeriodo) & "' "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
        If Trim(vgRegistro!cod_estadoreg) = Trim(iCodEstado) Then 'Or Trim(vgRegistro!cod_estadopri) = Trim(iCodEstado) Then
           flValidaEstadoProceso = True
        Else
            flValidaEstadoProceso = False
        End If
    Else
        flValidaEstadoProceso = False
    End If
    vgRegistro.Close

Exit Function
Err_flValidaTipoProceso:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function flInicializaVariables()
On Error GoTo Err_flInicializaVariables

    vlCodEmpresa = "0"
    vlCodSistema = "0"
    vlNumSolicitud = "0"
    vlNumCorrelativo = "0"
    vlCodSucursal = "0"
    vlFechaPago = "0"
    vlCentroCosto = "0"
    vlConPago = "0"
    vlDoctoPago = "0"
    vlNumDocto = "0"
    vlFechaDocto = "0"
    vlCodLugar = "0"
    vlRutRec = "0"
    vlCodMonedaOri = "0"
    vlCodMonedaPago = "0"
    vlCodImpto = "0"
    vlMtoNetoMO = "0"
    vlMtoDesctoMO = "0"
    vlMtoAfectoMO = "0"
    vlMtoExentoMO = "0"
    vlMtoImptoMO = "0"
    vlMtoNotaMO = "0"
    vlMtoTotalPagarMO = "0"
    vlMtoNetoMP = "0"
    vlMtoDesctoMP = "0"
    vlMtoAfectoMP = "0"
    vlMtoExentoMP = "0"
    vlMtoImptoMP = "0"
    vlMtoNotaMP = "0"
    vlMtoTotalPagarMP = "0"
    vlCuentaConAC = "0"
    vlCentroCostoAC = "0"
    vlCodProdAC = "0"
    vlCodAuxAC = "0"
    vlDoctoConAC = "0"
    vlGlosaAC = "0"
    vlTipoMovAC = "0"
    vlMtoAsigMO = "0"
    vlMtoAsigMN = "0"

Exit Function
Err_flInicializaVariables:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function flInsertarLinea()
On Error GoTo Err_flInsertarLinea

    vlAporteAdicionalCCAF = "0"
    vlConvDentalCCAF = "0"
    vlLeassingCCAF = "0"
    vlSeguroCCAF = "0"
    
    'Formatear datos para generar linea de registro de PREVIRED
    vlRut = Format(Trim(vlRut), "00000000000")
    vlDgv = Trim(vlDgv) & Space(1 - Len(Trim(vlDgv)))
    vlPaterno = Trim(vlPaterno) & Space(30 - Len(Trim(vlPaterno)))
    vlMaterno = Trim(vlMaterno) & Space(30 - Len(Trim(vlMaterno)))
    vlNombres = Trim(vlNombres) & Space(30 - Len(Trim(vlNombres)))
    vlTipoReceptor = Format(Trim(vlTipoReceptor), "0")
    vlTipoPago = Format(Trim(vlTipoPago), "0")
    vlPeriodo = Format(Trim(vlPeriodo), "000000")
    vlRentaImp = Replace(Format(Trim(vlRentaImp), "000000000000"), ",", ".")
    vlCotFonasa = Replace(Format(Trim(vlCotFonasa), "00000000"), ",", ".")
    
    If Trim(vlCodInsSalud) = "" Then
        vlCodInsSalud = "0"
    End If
    
    vlCodInsSalud = Format(Trim(vlCodInsSalud), "00")
    
    'vlMonedaPlanIsapre = Trim(Format(vlMonedaPlanIsapre, "0"))
    vlMonedaPlanIsapre = Trim(Format("1", "0"))
    
    vlCotIsapre = Replace(Format(Trim(vlCotIsapre), "00000000"), ",", ".")
    vlCotAdicional = Replace(Format(Trim(vlCotAdicional), "00000000"), ",", ".")
    vlOtrosAportes = Replace(Format(Trim(vlOtrosAportes), "00000000"), ",", ".")
    vlCotPactada = Replace(Format(Trim(vlCotPactada), "00000000"), ",", ".")
    vlTotalPagoIsapre = Replace(Format(Trim(vlTotalPagoIsapre), "00000000"), ",", ".")
    vlFUN = Format(Trim(vlFUN), "00000000")
    vlCodCCAF = Format(Trim(vlCodCCAF), "00")
    vlAporteCCAF = Replace(Format(Trim(vlAporteCCAF), "00000000"), ",", ".")
    vlAporteAdicionalCCAF = Replace(Format(Trim(vlAporteAdicionalCCAF), "00000000"), ",", ".")
    vlCreditoCCAF = Replace(Format(Trim(vlCreditoCCAF), "00000000"), ",", ".")
    vlConvDentalCCAF = Replace(Format(Trim(vlConvDentalCCAF), "00000000"), ",", ".")
    vlLeassingCCAF = Replace(Format(Trim(vlLeassingCCAF), "00000000"), ",", ".")
    vlSeguroCCAF = Replace(Format(Trim(vlSeguroCCAF), "00000000"), ",", ".")
    vlOtrosCCAF = Replace(Format(Trim(vlOtrosCCAF), "00000000"), ",", ".")
       
    'Generar Línea de Registro de Archivo PREVIRED
    vlLinea = (vlRut) & (vlDgv) & _
              (vlPaterno) & (vlMaterno) & _
              (vlNombres) & (vlTipoReceptor) & _
              (vlTipoPago) & (vlPeriodo) & _
              (vlRentaImp) & (vlCotFonasa) & _
              (vlCodInsSalud) & (vlMonedaPlanIsapre) & _
              (vlCotIsapre) & (vlCotAdicional) & _
              (vlOtrosAportes) & (vlCotPactada) & _
              (vlTotalPagoIsapre) & (vlFUN) & _
              (vlCodCCAF) & (vlAporteCCAF) & _
              (vlAporteAdicionalCCAF) & (vlCreditoCCAF) & _
              (vlConvDentalCCAF) & (vlLeassingCCAF) & _
              (vlSeguroCCAF) & (vlOtrosCCAF)

        Print #1, vlLinea

Exit Function
Err_flInsertarLinea:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function flIniciaVarLinea()
On Error GoTo Err_flIniciaVarLinea

    vlCotFonasa = "0"
    vlCodInsSalud = "0"
    vlCotIsapre = "0"
    vlCotAdicional = "0"
    vlOtrosAportes = "0"
    vlCotPactada = "0"
    vlTotalPagoIsapre = "0"
    vlFUN = "0"
    vlAporteCCAF = "0"
    vlAporteAdicionalCCAF = "0"
    vlConvDentalCCAF = "0"
    vlLeassingCCAF = "0"
    vlSeguroCCAF = "0"
    vlOtrosCCAF = "0"
   
Exit Function
Err_flIniciaVarLinea:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

'**************** FUNCIONES GENERACION ARCHIVO PRESTAMOS MEDICOS  *****
Function flExportarArchPresMed()
Dim vlLinea As String
Dim vlArchivo As String, vlOpen As Boolean
Dim vlContador As Long
Dim vlAumento As Integer
On Error GoTo Err_flExportarArchPresMed
    
    vlFechaInicio = Format(CDate(Trim(Txt_Desde)), "yyyymmdd")
    vlFechaTermino = Format(CDate(Trim(Txt_Hasta)), "yyyymmdd")

    If vlGlosaOpcion = "DEF" Then
        vlCodEstado = "C"
    Else
        vlCodEstado = "P"
    End If

    If flValidaEstadoProceso(Mid(Trim(vlFechaInicio), 1, 6), vlCodEstado) = False Then
        MsgBox "El Tipo de Proceso Seleccionado no se encuentra Realizado.", vbCritical, "Error de Datos"
        Exit Function
    End If
    
    vgSql = ""
    vgSql = "SELECT  b.num_poliza,b.num_orden,b.rut_ben,b.dgv_ben, "
    vgSql = vgSql & "b.gls_patben,b.gls_matben,b.gls_nomben,b.fec_fallben, "
    vgSql = vgSql & "p.num_perpago,p.mto_conhabdes "
    vgSql = vgSql & "FROM pp_tmae_liqpagopen" & vlGlosaOpcion & " l, pp_tmae_pagopen" & vlGlosaOpcion & " p, "
    vgSql = vgSql & "pp_tmae_ben b "
    vgSql = vgSql & "WHERE l.fec_pago >= '" & vlFechaInicio & "' AND "
    vgSql = vgSql & "l.fec_pago <= '" & vlFechaTermino & "' AND "
    vgSql = vgSql & "l.cod_tipopago = '" & vlPago & "' AND "
'    vgSql = vgSql & "l.cod_tipreceptor <> '" & Trim(clCodTipReceptorR) & "' AND "
    vgSql = vgSql & "p.cod_conhabdes = '" & Trim(clCodConHabDes29) & "' AND "
    vgSql = vgSql & "l.num_perpago = p.num_perpago AND "
    vgSql = vgSql & "l.num_poliza = p.num_poliza AND "
    vgSql = vgSql & "l.num_orden = p.num_orden AND "
    vgSql = vgSql & "l.rut_receptor = p.rut_receptor AND "
    vgSql = vgSql & "l.cod_tipreceptor = p.cod_tipreceptor AND "
    vgSql = vgSql & "l.num_poliza = b.num_poliza AND "
    vgSql = vgSql & "l.num_endoso = b.num_endoso AND "
    vgSql = vgSql & "l.num_orden = b.num_orden "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
      
        vlNumPerPago = ""
        vlNumPerPago = Trim(vgRegistro!Num_PerPago)
          
        'Selección del Archivo en el que se generarán los pagos
        ComDialogo.CancelError = True
        ComDialogo.FileName = "ArchivoPresMed" & vlNumPerPago & ".txt"
        ComDialogo.DialogTitle = "Guardar Pagos Préstamos Médicos como"
        ComDialogo.Filter = "*.txt"
        ComDialogo.FilterIndex = 1
        ComDialogo.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
        ComDialogo.ShowSave
        vlArchivo = ComDialogo.FileName
        If vlArchivo = "" Then
            Exit Function
        End If

        Screen.MousePointer = 11

        Open vlArchivo For Output As #1
        vlOpen = True

        Frm_BarraProg.Show
        Frm_BarraProg.ProgressBar1.Value = 0
        Frm_BarraProg.Refresh
        Frm_BarraProg.Lbl_Texto = "Generando Archivo " & vlArchivo 'vlNombre_Archivo
        Frm_BarraProg.Refresh
        Frm_BarraProg.ProgressBar1.Visible = True
        Frm_BarraProg.Refresh
        vlAumento = 100 / (1 + 2 + 3)
        
        While Not vgRegistro.EOF

            Call flInicializaVariablesPresMed
                
            vlNumPerPago = Trim(vgRegistro!Num_PerPago)
            vlNumPoliza = Trim(vgRegistro!Num_Poliza)
            vlNumOrden = (vgRegistro!Num_Orden)
            vlRut = (vgRegistro!Rut_Ben)
            vlDgv = (vgRegistro!Dgv_Ben)
            vlPaterno = (vgRegistro!Gls_PatBen)
            vlMaterno = (vgRegistro!Gls_MatBen)
            vlNombres = (vgRegistro!Gls_NomBen)
            vlNomCompleto = Trim(vlPaterno) & " " & Trim(vlMaterno) & " " & Trim(vlNombres)
            vlMtoConHabDes = (vgRegistro!Mto_ConHabDes)
                        
            If Len(Trim(vlNomCompleto)) > 40 Then
                vlNomCompleto = Mid(Trim(vlNomCompleto), 1, 40)
            End If
            
            If Not IsNull(vgRegistro!Fec_FallBen) Then
                vlIndNoPago = clIndNoPago3
            Else
                vlIndNoPago = clIndNoPago0
            End If
            
            'Formatear datos para generar linea de registro de PRESTAMOS MEDICOS
            vlRut = Format(Trim(vlRut), "00000000")
            vlDgv = Trim(vlDgv) & Space(1 - Len(Trim(vlDgv)))
            vlNomCompleto = Trim(vlNomCompleto) & Space(40 - Len(Trim(vlNomCompleto)))
            vlMtoConHabDes = Replace(Format(Trim(vlMtoConHabDes), "0000000"), ",", ".")
            vlIndNoPago = Format(Trim(vlIndNoPago), "0")
                
            'Generar Línea de Registro de Prestamos Medicos
            vlLinea = vlRut & _
                      vlDgv & _
                      vlNomCompleto & _
                      vlMtoConHabDes & _
                      vlIndNoPago
                      
            Print #1, vlLinea
            
            vgRegistro.MoveNext
            
            If Frm_BarraProg.ProgressBar1.Value + vlAumento < 100 Then
                Frm_BarraProg.ProgressBar1.Value = Frm_BarraProg.ProgressBar1.Value + vlAumento
            End If

        Wend

        Close #1

        Unload Frm_BarraProg
        Screen.MousePointer = 0
        vlOpen = False
        MsgBox "La Exportación de Datos al Archivo ha sido Finalizada Exitosamente.", vbInformation, "Estado de Exportación"
        Screen.MousePointer = vbDefault

    Else
        MsgBox "No existe Información para este Rango de Fechas", vbInformation, "Operacion Cancelada"
        Exit Function
    End If
    
Exit Function
Err_flExportarArchPresMed:
Screen.MousePointer = vbDefault
'Error por hacer click en boton cancelar de pantalla guardar como
If Err.Number = 32755 Then
    Exit Function
Else
    If vlOpen Then
        Close #1
    End If
    MsgBox "Se ha producido el siguiente error : " & Err.Description, vbCritical, "Error"
End If
End Function

Function flInicializaVariablesPresMed()
On Error GoTo Err_flInicializaVariablesPresMed

    vlNumPerPago = ""
    vlNumPoliza = ""
    vlNumOrden = ""
    vlRut = ""
    vlDgv = ""
    vlPaterno = ""
    vlMaterno = ""
    vlNombres = ""
    vlNomCompleto = ""
    vlMtoConHabDes = ""
    vlIndNoPago = ""

Exit Function
Err_flInicializaVariablesPresMed:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'*****************************************FUNCIONES***************
Function flExportarCotFonasa()
Dim vlLinea As String
Dim vlArchivo As String, vlOpen As Boolean
Dim vlContador As Long
Dim vlAumento As Integer
Dim vldireccion2 As String
On Error GoTo Err_flExportarCotFonasa


    vlFechaInicio = Format(CDate(Trim(Txt_Desde)), "yyyymmdd")
    vlFechaTermino = Format(CDate(Trim(Txt_Hasta)), "yyyymmdd")
    
    If vlGlosaOpcion = "DEF" Then
        vlCodEstado = "C"
    Else
        vlCodEstado = "P"
    End If
    
    If flValidaEstadoProceso(Mid(Trim(vlFechaInicio), 1, 6), vlCodEstado) = False Then
        MsgBox "El Tipo de Proceso Seleccionado no se encuentra Realizado.", vbCritical, "Error de Datos"
        Exit Function
    End If

    'Selección del Archivo en el que se generarán los pagos
    ComDialogo.CancelError = True
    ComDialogo.FileName = "ArchivoESSalud_" & vlFechaTermino & ".txt"
    ComDialogo.DialogTitle = "Guardar Cotización de EsSalud como"
    ComDialogo.Filter = "*.txt"
    ComDialogo.FilterIndex = 1
    ComDialogo.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    ComDialogo.ShowSave
    vlArchivo = ComDialogo.FileName
    If vlArchivo = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11

    Frm_BarraProg.Show
    Frm_BarraProg.ProgressBar1.Value = 0
    Frm_BarraProg.Refresh
    Frm_BarraProg.Lbl_Texto = "Generando Archivo " & vlArchivo 'vlNombre_Archivo
    Frm_BarraProg.Refresh
    Frm_BarraProg.ProgressBar1.Visible = True
    Frm_BarraProg.Refresh
    vlAumento = 100 / (1 + 2 + 3)

    Open vlArchivo For Output As #1
    vlOpen = True
    
    vlLinea = Trim(clColumna1) + ";" + Trim(clColumna2) + ";" + Trim(clColumna3) + ";" & _
    Trim(clColumna4) + ";" + Trim(clColumna5) + ";" + Trim(clColumna6) + ";" & _
    Trim(clColumna7) + ";" + Trim(clColumna8) + ";" + Trim(clColumna9) + ";" & _
    Trim(clColumna10) + ";" + Trim(clColumna11) + ";" + Trim(clColumna12) + ";" & _
    Trim(clColumna13) + ";" + Trim(clColumna14) + ";" + Trim(clColumna15) + ";" & _
    Trim(clColumna16) + ";" + Trim(clColumna17) + ";" + Trim(clColumna18) + ";" & _
    Trim(clColumna19) + ";" + Trim(clColumna20) + ";" + Trim(clColumna21) + ";" & _
    Trim(clColumna22) + ";" + Trim(clColumna23) + ";" + Trim(clColumna24) + ";" & _
    Trim(clColumna25) + ";" + Trim(clColumna26) + ";" + Trim(clColumna27) + ";" & _
    Trim(clColumna28) + ";" + Trim(clColumna29) + ";" + Trim(clColumna30) + ";" & _
    Trim(clColumna31) + ";" + Trim(clColumna32) + ";" + Trim(clColumna33) + ";" & _
    Trim(clColumna34) + ";" + Trim(clColumna35) + ";" + Trim(clColumna36)
    
    Print #1, vlLinea
    
    vgSql = "SELECT l.cod_tippension,l.num_perpago,l.num_orden," 'l.mto_baseimp,"
    vgSql = vgSql & "b.cod_tipoidenben,b.num_idenben,"
    vgSql = vgSql & "b.gls_patben,b.gls_matben,b.gls_nomben,b.gls_nomsegben,"
    vgSql = vgSql & "b.fec_nacben,b.cod_sexo,b.gls_fonoben,b.fec_ingreso,"
    vgSql = vgSql & "b.gls_dirben,b.cod_direccion,co.gls_comuna,b.cod_par,"
    vgSql = vgSql & "p.mto_conhabdes,l.cod_moneda,l.fec_pago, po.cod_cuspp,"
    vgSql = vgSql & "b.num_poliza,b.num_endoso "
    vgSql = vgSql & "FROM "
    vgSql = vgSql & "pp_tmae_liqpagopen" & vlGlosaOpcion & " l, pp_tmae_pagopen" & vlGlosaOpcion & " p,  "
    vgSql = vgSql & "pp_tmae_ben b , pp_tmae_poliza po "
    vgSql = vgSql & ",ma_tpar_comuna co "
    vgSql = vgSql & " WHERE "
    vgSql = vgSql & " (l.fec_pago >= '" & vlFechaInicio & "') AND "
    vgSql = vgSql & " (l.fec_pago <= '" & vlFechaTermino & "') AND "
    vgSql = vgSql & " (l.cod_tipopago = '" & vlPago & "') AND "
    vgSql = vgSql & " (l.num_perpago = p.num_perpago) AND "
    vgSql = vgSql & " (l.num_poliza = p.num_poliza ) AND "
    vgSql = vgSql & " (l.num_orden = p.num_orden ) AND "
    vgSql = vgSql & " (l.cod_tipreceptor = p.cod_tipreceptor) AND "
    vgSql = vgSql & " (l.num_idenreceptor = p.num_idenreceptor) AND "
    vgSql = vgSql & " (l.cod_tipoidenreceptor = p.cod_tipoidenreceptor) AND "
    vgSql = vgSql & " (p.cod_conhabdes = '" & clCodConHabDes24 & "') AND "
    vgSql = vgSql & " (l.cod_inssalud = '" & clCodInstRecauda & "') AND "
    vgSql = vgSql & " (l.num_poliza = b.num_poliza) AND "
    vgSql = vgSql & " (l.num_endoso = b.num_endoso) AND "
    vgSql = vgSql & " (l.num_orden = b.num_orden) "
    vgSql = vgSql & " AND (b.num_poliza = po.num_poliza) AND "
    vgSql = vgSql & " (b.num_endoso = po.num_endoso) "
    vgSql = vgSql & " AND b.cod_direccion = co.cod_direccion "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    While Not vgRegistro.EOF
               
        vlTipoPension = Format(Trim(vgRegistro!Cod_TipPension), "00")
        vlNumPagoAño = Trim(Mid((vgRegistro!Num_PerPago), 1, 4))
        vlNumPagoMes = Format((Mid((vgRegistro!Num_PerPago), 5, 6)), "00")
        vlNumOrden2 = Trim(vgRegistro!Num_Orden)
        vlTipoIdentif = Trim(vgRegistro!Cod_TipoIdenBen)
        vlNumIdenBenef = Trim(vgRegistro!Num_IdenBen)
        vlPaterno = Trim(vgRegistro!Gls_PatBen)
        
        If Not IsNull(vgRegistro!Gls_MatBen) Then
            vlMatBen = Trim(vgRegistro!Gls_MatBen)
        Else
            vlMatBen = ""
        End If
        
        If Not IsNull(vgRegistro!Gls_NomSegBen) Then
            vlNombre = Trim(vgRegistro!Gls_NomBen) + " " + Trim(vgRegistro!Gls_NomSegBen)
        Else
            vlNombre = Trim(vgRegistro!Gls_NomBen)
        End If
        
        vlFecNac = DateSerial(Mid((vgRegistro!Fec_NacBen), 1, 4), Mid((vgRegistro!Fec_NacBen), 5, 2), Mid((vgRegistro!Fec_NacBen), 7, 2))
        
        If (vgRegistro!Cod_Sexo) = "M" Then
            vlSexo = 1
        Else
            vlSexo = 2
        End If
        vlTelefono = IIf(IsNull(vgRegistro!Gls_FonoBen), "", Trim(vgRegistro!Gls_FonoBen))
        
        vlFecIngreso = DateSerial(Mid((vgRegistro!Fec_Ingreso), 1, 4), Mid((vgRegistro!Fec_Ingreso), 5, 2), Mid((vgRegistro!Fec_Ingreso), 7, 2))
        vlFchBaj = ""
        vlRuceps = ""
        vldireccion = Trim(vgRegistro!Gls_DirBen)
        vlNumDireccion = flBuscaNumDireccion(vgRegistro!Gls_DirBen)
        vlInteri = ""
        Call flBuscaNombreComuna(vgRegistro!Cod_Direccion)
        vlNomZon = vlNombreComuna
        vlRefEre = ""
        vlTipVia = ""
        vlTipZon = ""
        vlUbiGeo = ""
        vlMtoConHab = Format(Trim(vgRegistro!Mto_ConHabDes), "#0.00")
        vlNomCompania = vgNombreCompania
        vlCuspp = Trim(vgRegistro!Cod_Cuspp)

        vlNumPoliza2 = Format(Trim(vgRegistro!Num_Poliza), "0000000000")
                
        vlFchAfp = DateSerial(Mid((vgRegistro!Fec_Pago), 1, 4), Mid((vgRegistro!Fec_Pago), 5, 2), Mid((vgRegistro!Fec_Pago), 7, 2))
        
        If (vgRegistro!Cod_TipPension) <> "04" And (vgRegistro!Cod_TipPension) <> "05" And (vgRegistro!Cod_TipPension) <> "06" And (vgRegistro!Cod_TipPension) <> "07" Then
            vlNomCausante = flBuscarNombreCausante(vgRegistro!Num_Poliza, vgRegistro!num_endoso, vgRegistro!Cod_Par)
            ''vlNomBen = flBuscarNombreCausante(vgRegistro!Num_Poliza, vgRegistro!num_endoso, vgRegistro!Cod_Par)
        Else
            vlNomCausante = (vlPaterno) + " " + (vlMatBen) + " " + (vlNombre)
            ''vlNomBen = (vlPaterno) + " " + (vlMatBen) + " " + (vlNombre)
        End If
        
        If Trim(vgRegistro!Cod_Moneda) = "NS" Then
            vlMoneda = 1
        Else
            vlMoneda = 2
        End If
       
        vlLinea = Trim(vlTipoPension) + ";" + Trim(vlNumPagoAño) + ";" & _
                      Trim(vlNumPagoMes) + ";" + Trim(vlNumOrden2) + ";" & _
                      Trim(vlTipoIdentif) + ";" + Trim(vlNumIdenBenef) + ";" & _
                      Trim(vlPaterno) + ";" + Trim(vlMatBen) + ";" & _
                      Trim(vlNombre) + ";" + Trim(vlFecNac) + ";" + Trim(vlSexo) + ";" & _
                      Trim(vlTelefono) + ";" + Trim(vlFecIngreso) + ";" + Trim(clArcFonasaSittra) + ";" & _
                      Trim(clArcFonasaTiptra) + ";" + Trim(vlFchBaj) + ";" + Trim(vlRuceps) + ";" & _
                      Trim(clArcFonasaEssvid) + ";" + Trim(clArcFonasaRegpen) + ";" + Trim(clArcFonasaSctr_1) + ";" & _
                      Trim(vldireccion) + ";" + Trim(vlNumDireccion) + ";" + Trim(vlInteri) + ";" + Trim(vlNomZon) + ";" & _
                      Trim(vlRefEre) + ";" + Trim(vlTipVia) + ";" + Trim(vlTipZon) + ";" & _
                      Trim(vlUbiGeo) + ";" + Trim(clArcFonasaDiatra) + ";" + Trim(vlMtoConHab) + ";" & _
                      Trim(vlFchAfp) + ";" + Trim(vlMoneda) + ";" + Trim(vlNomCompania) + ";" & _
                      Trim(vlCuspp) + ";" + Trim(vlNumPoliza2) + ";" + Trim(vlNomCausante)
         
            Print #1, vlLinea
            vgRegistro.MoveNext
            If Frm_BarraProg.ProgressBar1.Value + vlAumento < 100 Then
                Frm_BarraProg.ProgressBar1.Value = Frm_BarraProg.ProgressBar1.Value + vlAumento
            End If
            
        Wend
        
        Close #1
        vgRegistro.Close
        
        Unload Frm_BarraProg
        Screen.MousePointer = 0
        vlOpen = False
        MsgBox "La Exportación de Datos al Archivo ha sido Finalizada Exitosamente.", vbInformation, "Estado de Exportación"
        Screen.MousePointer = vbDefault
        
'    Else
'        MsgBox "No existe Información para este Rango de Fechas", vbInformation, "Operacion Cancelada"
'        Exit Function
'    End If

Exit Function
Err_flExportarCotFonasa:
Screen.MousePointer = vbDefault
'Error por hacer click en boton cancelar de pantalla guardar como
If Err.Number = 32755 Then
    Exit Function
Else
    If vlOpen Then
        Close #1
    End If
    MsgBox "Se ha producido el siguiente error : " & Err.Description, vbCritical, "Error"
End If
End Function

Function flBuscarNombreCausante(iNumPoliza As String, inumendoso As Long, iCodPar As String) As String
Dim vlRegCau As ADODB.Recordset
Dim Sql As String
Dim vlAuxNomSeg As String, vlAuxApeMat As String

    flBuscarNombreCausante = ""

    Sql = " SELECT gls_nomben,gls_nomsegben,gls_patben,gls_matben "
    Sql = Sql & " FROM pp_tmae_ben "
    Sql = Sql & " WHERE num_poliza = '" & iNumPoliza & "' AND "
    Sql = Sql & " num_endoso = " & inumendoso & " "
    Set vlRegCau = vgConexionBD.Execute(Sql)
    If Not vlRegCau.EOF Then
        vlAuxNomSeg = IIf(IsNull(vlRegCau!Gls_NomSegBen), "", Trim(vlRegCau!Gls_NomSegBen))
        vlAuxApeMat = IIf(IsNull(vlRegCau!Gls_MatBen), "", Trim(vlRegCau!Gls_MatBen))
        flBuscarNombreCausante = fgFormarNombreCompleto(Trim(vlRegCau!Gls_NomBen), vlAuxNomSeg, Trim(vlRegCau!Gls_PatBen), vlAuxApeMat)
    End If
    vlRegCau.Close

End Function

Function flBuscaNumDireccion(idireccion As String) As String
Dim vlnumeros As String
Dim vldireccion As String
Dim vlPosicion As Integer
Dim vlPosicion2 As Integer
Dim i As Integer
vlnumeros = "0123456789"
For i = 1 To Len(idireccion)
    vldireccion = Mid(idireccion, i, 1)
    If (InStr(1, vlnumeros, vldireccion)) <> 0 Then
        vlPosicion = i
    End If
Next i
If (vlPosicion <> 0) Then
    vlPosicion2 = Trim(InStrRev(idireccion, " ", vlPosicion) + 1)
    flBuscaNumDireccion = Mid(idireccion, vlPosicion2, vlPosicion)
Else
    flBuscaNumDireccion = ""
End If
End Function

Function flBuscaNombreComuna(Codigo As Integer)

    vgSql = ""
    vgSql = vgSql & "SELECT gls_comuna "
    vgSql = vgSql & "FROM MA_TPAR_COMUNA "
    vgSql = vgSql & "WHERE cod_direccion = " & Trim(Codigo) & " "
    Set vgRs4 = vgConexionBD.Execute(vgSql)
    If Not vgRs4.EOF Then
       vlNombreComuna = (vgRs4!gls_comuna)
    End If

End Function

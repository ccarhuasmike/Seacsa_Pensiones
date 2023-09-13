VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Frm_PlanillaPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planilla de Pagos "
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   Icon            =   "Frm_PlanillaPago.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   6285
   Begin Crystal.CrystalReport Rpt_Reporte 
      Left            =   120
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Fra_Operaciones 
      Height          =   1095
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   6015
      Begin VB.CommandButton Cmd_Imprimir_Directo 
         Caption         =   "&Directo"
         Height          =   675
         Left            =   1320
         Picture         =   "Frm_PlanillaPago.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir_AFP 
         Caption         =   "&AFP"
         Enabled         =   0   'False
         Height          =   675
         Left            =   2280
         Picture         =   "Frm_PlanillaPago.frx":06C6
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   4200
         Picture         =   "Frm_PlanillaPago.frx":0D80
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   3240
         Picture         =   "Frm_PlanillaPago.frx":1092
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   1320
         Picture         =   "Frm_PlanillaPago.frx":174C
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   720
      End
   End
   Begin VB.Frame Fra_Datos 
      Height          =   2055
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   6015
      Begin VB.ComboBox Cmb_Pago 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   880
         Width           =   2775
      End
      Begin VB.ComboBox Cmb_Tipo 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox Txt_Hasta 
         Height          =   285
         Left            =   4320
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox Txt_Desde 
         Height          =   285
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1440
         Width           =   1215
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
         TabIndex        =   12
         Top             =   880
         Width           =   2415
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
         TabIndex        =   10
         Top             =   1440
         Width           =   2295
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
         TabIndex        =   8
         Top             =   1440
         Width           =   135
      End
   End
   Begin VB.Label lbl_Indicador 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3240
      Width           =   6015
   End
End
Attribute VB_Name = "Frm_PlanillaPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim vlCodEstado As String

Dim vlFechaDesde As String
Dim vlFechaHasta As String
Dim vlNombre As String

Dim vlArchivo       As String
Dim vlGlosaOpcion   As String
Dim vlOpcion        As String
Dim vlPago          As String
Dim vlSw            As Boolean

'Variables Retencion Judicial

Dim vlFechaInicio As String, vlFechaTermino As String, vlTipoRetJud As String
Dim vlNumRetencion As Integer, vlNumCargas As Integer, vlMtoCargas As Double
Dim vlMtoHaber As Double, vlMtoDescto As Double, vlMtoHabDes As Double
Dim vlPeriodo As String, vlCodHabDes As String, vlMtoRetJud As Double
Dim vlCodOtrosHabDes As String, vlTipoMov As String
Dim vlNumOrdenCar As Integer

Dim vlNumCargasAux As Double
Dim vlMtoCargasAux As Double
Dim vlMtoHaberAux As Double
Dim vlMtoDesctoAux As Double
Dim vlMtoRetNetaAux As Double

Const vlCodHabDesRJC As String = "('60','61','62')"
Const clTipoRJ As String * 2 = "RJ"
Const clTipoRAF As String * 3 = "RAF"
Const clCodHabDes61 As String * 2 = "61"
Const clCodHabDes62 As String * 2 = "62"
Const clCodRJ As String * 3 = "RJ"
Const clCodRJC As String * 3 = "RJC"
Const clCodHab As String * 1 = "H"
Const clCodDescto As String * 1 = "D"

Dim vlRegRetJud As ADODB.Recordset

'Variables de Asignación Familiar
Const vlCodHabDesAF As String = "('08','09','30')"

Dim vlNumCarRetro As Long
Dim vlNumReliq As Integer
Dim vlMtoRetroPagoPen As Double
 
Const clCodHabDes01 As String * 2 = "01"
Const clCodHabDes0102 As String = "('01','02')"

'Variables de Garantia Estatal

Dim vlFecPeriodo As String, vlMtoPensionPrc As Double, vlMtoPensionPesos As Double
Dim vlMtoGarEst As Double, vlMtoOtroHab As Double, vlMtoOtroDes As Double
Dim vlBono As Double, vlCodOtrosHabDesGE As String, vlCodigo As String
Dim vlCodUsuario As String, vlNumPerPago As String, vlNumPoliza As String
Dim vlNumEndoso As Integer, vlNumOrden As Integer, vlNumResGarEst As String
Dim vlNumAnnoRes As String, vlCodTipRes As String, vlPrcDeduccion As Double
Dim vlCodDerGarEst As String, vlMtoPension As Double 'vlMtoPensionPesos As Double
'Dim vlMtoOtroDes As Double 'vlMtoGarEst As Double 'vlMtoOtroHab As Double,
Dim vlMtoBono As Double, vlCodConHabDes As String, vlNumCuotas As Integer
Dim vlMtoCuota As Double, vlMtoTotal As Double, vlCodMoneda As String
Dim vlMtoTotalHabDes As Double, vlCodTipMov As String, vlFechaPago As String




Dim vlCodTipReceptor As String
Dim vlCodTipoIdenReceptor As String
Dim vlNumIdenReceptor As String
Dim vlCodTipoIdenRec As String
Dim vlNumIdenRec As String
Dim vlGlsNomRec As String
Dim vlGlsNomSegRec As String
Dim vlGlsPatRec As String
Dim vlGlsMatRec As String
Dim vlCodGruFam As String
Dim vlNumPol As String
'RRR 16/08/2012
Dim vlCuspp As String
Dim vlPerPago As String


'Inicio GCP 29032019
Dim VlCOD_TIPCTA As String
Dim VlCOD_SUCURSAL As String
Dim VlCOD_MONBCO As String
Dim VlNUM_CTABCO As String
Dim VlCOD_BANCO As String
Dim VlDES_BANCO As String
Dim VlNUM_CUENTA_CCI As String
'Fin GCP 29032019

Dim vlMtoSalud As Double
Dim vlMtoImpuesto As Double
'Dim vlMtoPensionPesos As Double

Dim vlMtoHabImp As Double
Dim vlMtoDesImp As Double
Dim vlMtoHabNoImp As Double
Dim vlMtoDesNoImp As Double

'Tipo de Receptor R = Retenedor
Const clTipRecR As String * 1 = "R" 'Retenedor
Const clTipRecT As String * 1 = "T" 'Tutor
Const clTipRecP As String * 1 = "P" 'Pensionado
Const clTipRecM As String * 1 = "M" 'Madre
'Códigos de Parentesco de Madres y Conyuges
Const clCodParMadres As String = "('10','11','20','21')"

Const clCodConHDMtoPen As String * 2 = "01"
Const clCodConHD0102 As String = "('01','02')"
Const clModOrigenPP As String * 2 = "PP"
Const clCodGE As String * 2 = "03"
Const clCodBonInv As String * 2 = "06"
Const clCodHabGE As String = "('02','04','07')"
Const clCodDesGE As String = "('21','22')"
Const clTipoGE As String * 2 = "GE"
Const clCodH As String * 1 = "H"
Const clCodD As String * 1 = "D"
Const clCodImpS As String * 1 = "S"
Const clCodImpN As String * 1 = "N"
'Codigo de Concepto de H/D de Monto de Pension en Pesos
Const clCodHD01 As String * 2 = "01"
'Codigo de Concepto de H/D de Monto de Garantía Estatal por Quiebra de la Cía.
Const clCodHD02 As String * 2 = "02"
'Codigo de Concepto de H/D de Monto de Impuesto
Const clCodHD23 As String * 2 = "23"
'Codigo de Concepto de H/D de Monto de Salud
Const clCodHD24 As String * 2 = "24"

'Código de Derecho a Garantia Estatal (0 = Sin Derecho/ codtabla = 'DEG')
Const clCodSinDer As String * 1 = "0"
'Código de Causa de Suspensión de G.E. (0 = Sin Información/codtabla = 'CSG')
Const clCodSinInfo As String * 1 = "0"

'Variables de Planilla de Pago de Salud de Pensiondes RV

Const clCodISExento As String * 2 = "00"

'Variables de Prestamos Medicos Fonasa

Dim vlPrcSalud As Integer

Const clFechaTopeTer As String * 8 = "99991231"
Const clCodConHabDes29 As String * 2 = "29"
Const clCodConHabDes24 As String * 2 = "24"
Const clCodConHabDes34 As String * 2 = "34"
Const clCodTabCodIS As String * 2 = "IS" 'Código de Institución de Salud
Const clCodPrcSalud As String * 2 = "PS" 'Código de Tabla = Porcentaje de Salud
Const clCodPSM As String * 3 = "PSM" 'Código Elemento = Porcentaje de Salud

'Variables de Centralización Contable

Dim vlRegistroConcepto As ADODB.Recordset
Dim vlRegistroMtoConcepto As ADODB.Recordset
Dim vlRegistroFechas As ADODB.Recordset

'Dim vlGlsLibro As String
'Dim vlCodConHabDes As String
'Dim vlCodTipMov As String
Dim vlMtoTotal1 As Double
Dim vlMtoTotal2 As Double
Dim vlMtoTotal3 As Double
Dim vlMtoTotal4 As Double
Dim vlMtoTotal5 As Double
Dim vlMtoTotal6 As Double
Dim vlMtoTotal7 As Double
Dim vlMtoTotal8 As Double
Dim vlMtoTotal9 As Double
Dim vlMtoTotal10 As Double

Dim vlCodCtaCon As String

Dim vlMonto As Double

Dim vlMtoTotalVN As Double
Dim vlMtoTotalVA As Double
Dim vlMtoTotalIT As Double
Dim vlMtoTotalIP As Double
Dim vlMtoTotalSO As Double
Dim vlMtoTotalSO100 As Double
Dim vlMtoTotalSOVN As Double
Dim vlMtoTotalSOVA As Double
Dim vlMtoTotalSOIT As Double
Dim vlMtoTotalSOIP As Double

Dim vgFormatoAFP As String    '200912 DCM

'20050706
Private Type TyCodConHabDes
    CodConHabDes    As String
    GlsCtaCon       As String
    CodTipMov       As String
    CodCtaCon       As String
End Type

'Registro de Códigos de Conceptos de Haberes y Descuentos
Private stCodConHabDes() As TyCodConHabDes

Dim vlGlsCtaCon As String
Dim vlNumRegistros As Integer

'Libro de Pensiones de Renta Vitalicia
Dim vlTipoReceptor As String

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
Const clCodConHabDes As String = "('01','23','24') "

Const clTipPenSob As String = "('08')"

Dim vlMtoTipPenVN As Double
Dim vlMtoTipPenVA As Double
Dim vlMtoTipPenIT As Double
Dim vlMtoTipPenSO As Double

Dim vlCodAFP As String, vlViaPago As String, vlFecPago As String  '21/06/2008
Dim vlFactorAju As Double
Dim vlDesAFP As String
Dim vlDesPension As String
Dim glNumpol As String

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
(ByVal lpBuffer As String, nSize As Long) As Long

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
        If Trim(vgRegistro!cod_estadoreg) = Trim(iCodEstado) Then
           'Or Trim(vgRegistro!cod_estadopri) = Trim(iCodEstado)
           flValidaEstadoProceso = True
        Else
            flValidaEstadoProceso = False
        End If
    Else
        flValidaEstadoProceso = False
    End If

Exit Function
Err_flValidaTipoProceso:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flInformeVenTut()
On Error GoTo Err_VenTut

    'HQR 01/09/2004 Se cambia query
    'vlSql = "select T.num_poliza,P.cod_tippension,B.rut_ben,B.dgv_ben,"
    'vlSQL = vlSQL & "T.num_orden,B.gls_nomben,B.gls_patben,B.gls_matben,"
    'vlSQL = vlSQL & "T.fec_inipodnot,T.fec_terpodnot,T.rut_tut,T.dgv_tut,"
    'vlSQL = vlSQL & "T.gls_nomtut,T.gls_pattut,T.gls_mattut "
    'vlSQL = vlSQL & "from PP_TMAE_POLIZA P, PP_TMAE_BEN B, PP_TMAE_TUTOR T "
    'vlSQL = vlSQL & "where P.num_poliza = T.num_poliza "
    'vlSQL = vlSQL & "and B.num_poliza= T.num_poliza "
    'vlSQL = vlSQL & "and P.num_endoso= (select max(num_endoso) from PP_tmae_tutor where num_poliza=T.num_poliza group by num_poliza)"
    'vlSQL = vlSQL & " and B.num_endoso = (select max(num_endoso) from PP_tmae_tutor where num_poliza=T.num_poliza group by num_poliza)"
    'vlSQL = vlSQL & " and B.cod_par= '99' "
    'vlSQL = vlSQL & "and T.fec_inipodnot>= '" & vlFechaDesde & "' "
    'vlSQL = vlSQL & "and T.fec_terpodnot<= '" & vlFechaHasta & "' "
    'vlSQL = vlSQL & "order by T.num_poliza, T.num_endoso, T.num_orden "
    
    vlSql = ""
    vlSql = "SELECT P.NUM_POLIZA, P.COD_TIPPENSION, B.COD_TIPOIDENBEN, B.NUM_IDENBEN,"
    vlSql = vlSql & " B.NUM_ORDEN, B.GLS_NOMBEN, B.GLS_NOMSEGBEN, B.GLS_PATBEN, B.GLS_MATBEN,"
    vlSql = vlSql & " T.FEC_INIPODNOT, T.FEC_TERPODNOT, T.COD_TIPOIDENTUT, T.NUM_IDENTUT,"
    vlSql = vlSql & " T.GLS_NOMTUT, T.GLS_NOMSEGTUT, T.GLS_PATTUT, T.GLS_MATTUT"
    vlSql = vlSql & " FROM PP_TMAE_POLIZA P, PP_TMAE_BEN B, PP_TMAE_TUTOR T"
    vlSql = vlSql & " WHERE "
    vlSql = vlSql & " T.FEC_TERPODNOT >= '" & vlFechaDesde & "' "
    vlSql = vlSql & " AND T.FEC_TERPODNOT <= '" & vlFechaHasta & "' "
    vlSql = vlSql & " AND P.NUM_POLIZA = T.NUM_POLIZA"
    vlSql = vlSql & " AND P.NUM_ENDOSO = (SELECT MAX(NUM_ENDOSO) FROM PP_TMAE_POLIZA WHERE NUM_POLIZA = T.NUM_POLIZA)"
    vlSql = vlSql & " AND P.num_poliza = B.num_poliza"
    vlSql = vlSql & " AND P.NUM_ENDOSO = B.NUM_ENDOSO"
    vlSql = vlSql & " AND B.NUM_ORDEN = T.NUM_ORDEN"
    vlSql = vlSql & " ORDER BY P.NUM_POLIZA, P.NUM_ENDOSO, T.NUM_ORDEN "
    'Fin HQR 01/09/2004
    Set vgRs = vgConexionBD.Execute(vlSql)
    If vgRs.EOF Then
        MsgBox "No existe Información para este rango de Fechas", vbInformation, "Operacion Cancelada"
        Exit Function
    Else
       vgRs.Close
       
       vlArchivo = strRpt & "PP_Rpt_CieVenTutor.rpt"   '\Reportes
       If Not fgExiste(vlArchivo) Then     ', vbNormal
           MsgBox "Archivo de Reporte de Vencimiento de Tutores no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
           Screen.MousePointer = 0
           Exit Function
       End If
       
       Call fgVigenciaQuiebra(Txt_Desde)
        
        'HQR 01/09/2004
        'vgQuery = " {PP_TMAE_BEN.cod_par}= '99' "
        'vgQuery = vgQuery & " AND {PP_TMAE_TUTOR.FEC_INIPODNOT}>= '" & vlFechaDesde & "' "
        vgQuery = "{PP_TMAE_TUTOR.FEC_TERPODNOT} >= '" & vlFechaDesde & "' "
        vgQuery = vgQuery & " AND {PP_TMAE_TUTOR.FEC_TERPODNOT} <= '" & vlFechaHasta & "' "
        
        Rpt_Reporte.Reset
        Rpt_Reporte.WindowState = crptMaximized
        Rpt_Reporte.ReportFileName = vlArchivo
        Rpt_Reporte.Connect = vgRutaDataBase
        Rpt_Reporte.SelectionFormula = ""
        Rpt_Reporte.SelectionFormula = vgQuery
        Rpt_Reporte.Formulas(7) = ""
        Rpt_Reporte.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
        Rpt_Reporte.Formulas(1) = "NombreSistema = '" & vgNombreSistema & "'"
        Rpt_Reporte.Formulas(2) = "NombreSubSistema = '" & vgNombreSubSistema & "'"
        Rpt_Reporte.Formulas(3) = "FecDes = '" & vlFechaDesde & "'"
        Rpt_Reporte.Formulas(4) = "FecHas = '" & vlFechaHasta & "'"
        Rpt_Reporte.Formulas(5) = "FechaTexto = '" & Txt_Desde & "  *  " & Txt_Hasta & "'"
        
        If Trim(vlGlosaOpcion) = "DEF" Then
            Rpt_Reporte.Formulas(6) = "TipoProceso= 'DEFINITIVO' "
        Else
            Rpt_Reporte.Formulas(6) = "TipoProceso= 'PROVISORIO' "
        End If
        
        Rpt_Reporte.Formulas(7) = "GlosaQuiebra = '" & vgGlosaQuiebra & "'"
        
        Rpt_Reporte.Destination = crptToWindow
        Rpt_Reporte.WindowTitle = "Informe de Vencimiento de Tutores"
        Rpt_Reporte.Action = 1
        Screen.MousePointer = 0
    End If

Exit Function
Err_VenTut:
    Screen.MousePointer = 0
    MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
End Function

Function Get_User_Name() As String
    'creamos variables
    Dim lpBuff As String * 25
    Dim ret As Long
    Dim UserName As String
    'Obtenemos el nombre de la api.
    ret = GetUserName(lpBuff, 25)
    UserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
    ' Devolvemos el nombre de usuario
    Get_User_Name = UserName
End Function

Function flInformePensiones()
On Error GoTo Err_pension
Dim objRep As New ClsReporte
Dim pol As String
Dim WL As Integer
Dim fecCierre As String
Dim NombreExcel As String
fecCierre = Mid(Format(Txt_Desde, "YYYYMMDD"), 1, 6)
NombreExcel = "INTERNO POR MONEDAS - " & fecCierre
WL = 1000


'    vlArchivo = strRpt & "PP_Rpt_CieLibroPensRtasVit.rpt"   '\Reportes
'    If Not fgExiste(vlArchivo) Then     ', vbNormal
'        MsgBox "Archivo de Reporte no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
'        Screen.MousePointer = 0
'        Exit Function
'    End If


    If Trim(vlGlosaOpcion) = "DEF" Then
        TipoProceso = "DEFINITIVO"
    Else
        TipoProceso = "PROVISORIO"
    End If

    Screen.MousePointer = 11
    
    'Función que permite determinar los Datos del Libro de Pensiones
    Call flLlenaTablaTemporalPensiones(vlGlosaOpcion, vlPago)
    Call fgVigenciaQuiebra(Txt_Desde)

    Dim vlTipoProceso As String
    Dim vlGlosa As String
    Dim vlSql As String

'    vlSql = " select a.*, c.gls_elemento from PP_TTMP_PENSIONES a "
'    vlSql = vlSql & " join ma_tpar_tabcod c on a.cod_moneda=cod_elemento and c.cod_tabla='MP' "
    
    vlSql = " SELECT SUBSTR(A.FEC_PAGO,1,6) NUM_PERIODO, NUM_POLIZA, NUM_ORDEN, DES_PENSION, B.GLS_TIPOIDEN  || '-' || NUM_IDENBEN GLS_DOCBEN, COD_CUSSP, GLS_NOMBEN || ' '|| GLS_NOMSEGBEN || ' ' || GLS_PATBEN || ' ' || GLS_MATREC GLS_BENEF,"
    vlSql = vlSql & " MTO_PENSION, MTO_BASEIMP, MTO_SALUD, MTO_BASETRI, MTO_DESNOIMP, MTO_PENLIQ,"
    vlSql = vlSql & " C.GLS_TIPOIDEN || '-' || NUM_IDENREC GLS_DOCREC,GLS_NOMREC || ' ' || GLS_NOMSEGREC || ' ' || GLS_PATREC || ' ' || GLS_MATREC GLS_RECEPTOR, DES_AFP, COD_MONEDA, DES_VIAPAGO, DES_BANCO, TIPO_CUENTA, COD_MONBCO, NUM_CTABCO, NUM_CUENTA_CCI"
    vlSql = vlSql & " FROM PP_TTMP_PENSIONES A"
    vlSql = vlSql & " LEFT JOIN MA_TPAR_TIPOIDEN B ON A.COD_TIPOIDENBEN=B.COD_TIPOIDEN"
    vlSql = vlSql & " LEFT JOIN MA_TPAR_TIPOIDEN C ON A.COD_TIPOIDENREC=C.COD_TIPOIDEN"
    vlSql = vlSql & " WHERE A.COD_USUARIO='" & vgUsuario & "'"
    vlSql = vlSql & " ORDER BY 2"
    
    'Set vgRs = vgConexionBD.Execute(vlSql)
    Set vgRs = New ADODB.Recordset
    vgRs.CursorLocation = adUseClient
    vgRs.Open vlSql, vgConexionBD, adOpenStatic, adLockBatchOptimistic
    
    
    If Not (vgRs.EOF) Then
            
            Dim totReg As Long
            totReg = vgRs.RecordCount
            
            Dim o_Excel     As Object
            Dim o_Libro     As Object
            Dim o_Hoja      As Object
            Dim Columna     As Long
            Dim Col     As Long
            
            Set o_Excel = CreateObject("Excel.Application")
            Set o_Libro = o_Excel.Workbooks.Add
            Set o_Hoja = o_Libro.Worksheets.Add
            
            Dim porc As Double
            Dim i As Long
            Dim ColTot As Integer
            i = 0
            
            ColTot = vgRs.Fields.Count - 1
            'Para colocar los titulos
            For Col = 0 To ColTot
                Select Case Col + 1
                    Case 0
                    'fechas
                    o_Hoja.Columns(Col + 1).NumberFormat = "d/m/yyyy;@"
                    Case 1, 2, 3, 4, 5, 6, 7, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23
                    'texto
                    o_Hoja.Columns(Col + 1).NumberFormat = "@"
                    Case 8, 9, 10, 11, 12, 13
                    'numeros
                End Select
                
                o_Hoja.Cells(1, Col + 1).Font.Bold = True
                o_Hoja.Cells(1, Col + 1).Value = vgRs.Fields(Col).Name
            Next
            
            Dim celda As String
            Dim celfec As Date
            Do While Not vgRs.EOF 'i > 50
                If totReg > 0 Then
                    porc = (100 * i) / totReg
                Else
                    porc = 0
                End If
                lbl_Indicador.Caption = "Cargando... " & CStr(Int(porc)) & "% (" & i & "/" & totReg & ")"
                
                i = i + 1
                
                For Col = 0 To ColTot
                    Select Case Col + 1
'                        Case 0
'                        fechas
'                            If CStr(vgRs.Fields(Col)) = "01/01/1900" Then
'                               'celda = vgRs.Fields(Col).Value
'                               celda = "01/01/1900"
'                               o_Hoja.Cells(i + 1, Col + 1).Value = celda
'                            Else
'                               celfec = CDate(vgRs.Fields(Col).Value)
'                               o_Hoja.Cells(i + 1, Col + 1).Value = celfec
'                            End If
                            
                        Case 1, 2, 3, 4, 5, 6, 7, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23
                        'texto
                        
                        
                            If IsNull(vgRs.Fields(2)) Then
                               pol = CStr(vgRs.Fields(2).Value)
                            End If
                            
                            If IsNull(vgRs.Fields(Col)) Then
                               celda = ""
                            Else
                               celda = CStr(vgRs.Fields(Col).Value)
                               'celda = vgRs.Fields(Col).Value
                            End If
                            o_Hoja.Cells(i + 1, Col + 1).Value = Trim(celda)
                        Case 8, 9, 10, 11, 12, 13
                        'numeros
                            If IsNull(vgRs.Fields(Col)) Then
                               celda = ""
                            Else
                               celda = CStr(vgRs.Fields(Col).Value)
                               'celda = vgRs.Fields(Col).Value
                            End If
                            o_Hoja.Cells(i + 1, Col + 1).Value = Trim(celda)
                    End Select
'                    'o_Hoja.Cells(i + 1, Columna + 1).NumberFormat = "@"
'                    Dim celda As String
'                    If IsNull(vgRs.Fields(Col)) Then
'                       celda = ""
'                    Else
'                       celda = CStr(vgRs.Fields(Col).Value)
'                    End If
'                    o_Hoja.Cells(i + 1, Col + 1).Value = celda
                Next
                vgRs.MoveNext
            Loop
            lbl_Indicador.Caption = "Archivo generado correctamente"
            'o_Excel.Visible = True 'para ver vista previa
            o_Excel.ActiveWorkbook.SaveAs "C:\Users\" & Get_User_Name & "\Documents\" & NombreExcel & "_" & TipoProceso & ".xlsx"
            vgRs.Close
            
    End If
    
    
    Set o_Hoja = Nothing
    Set o_Libro = Nothing
    o_Excel.Quit
    Set o_Excel = Nothing

    lbl_Indicador.Caption = "Se guardo en " & "C:\Users\" & Get_User_Name & "\Documents\"
    Screen.MousePointer = 0


    
'    Dim LNGa As Long
'    LNGa = CreateFieldDefFile(vgRs, Replace(UCase(strRpt & "Estructura\PP_Rpt_CieLibroPensRtasVit.rpt"), ".RPT", ".TTX"), 1)
'
'    If Trim(vlGlosaOpcion) = "DEF" Then
'        TipoProceso = "DEFINITIVO"
'    Else
'        TipoProceso = "PROVISORIO"
'    End If
'
'    If objRep.CargaReporte(strRpt & "", "PP_Rpt_CieLibroPensRtasVit.rpt", "Informe Certificados Pendientes de reliquidación", vgRs, True, _
'                            ArrFormulas("GlosaQuiebra", vgNombreCompania), _
'                            ArrFormulas("TipoProceso", TipoProceso), _
'                            ArrFormulas("NombreCompania", vgNombreCompania), _
'                            ArrFormulas("NombreSistema", vgNombreSistema), _
'                            ArrFormulas("NombreSubSistema", vgNombreSubSistema)) = False Then
'
'        MsgBox "No se pudo abrir el reporte", vbInformation
'        Exit Function
'    End If



   ' Call fgVigenciaQuiebra(Txt_Desde)

   ' If vlSw = True Then

   '     vgQuery = ""
   '     vgQuery = "{PP_TTMP_PENSIONES.COD_USUARIO} = '" & Trim(vgUsuario) & "' "

   '     Rpt_Reporte.Reset
   '     Rpt_Reporte.WindowState = crptMaximized
   '     Rpt_Reporte.ReportFileName = vlArchivo
   '     Rpt_Reporte.Connect = vgRutaDataBase
   '     Rpt_Reporte.SelectionFormula = vgQuery
   '     Rpt_Reporte.Formulas(4) = ""
   '     Rpt_Reporte.Formulas(0) = "NombreCompania='" & vgNombreCompania & "'"
   '     Rpt_Reporte.Formulas(1) = "NombreSistema= '" & vgNombreSistema & "'"
   '     Rpt_Reporte.Formulas(2) = "NombreSubSistema= '" & vgNombreSubSistema & "'"

   '     If Trim(vlGlosaOpcion) = "DEF" Then
   '         Rpt_Reporte.Formulas(3) = "TipoProceso= 'DEFINITIVO' "
   '     Else
   '         Rpt_Reporte.Formulas(3) = "TipoProceso= 'PROVISORIO' "
   '     End If
   '     Rpt_Reporte.Formulas(4) = "GlosaQuiebra = '" & vgGlosaQuiebra & "'"

   '     Rpt_Reporte.Destination = crptToWindow
   '    Rpt_Reporte.WindowTitle = "Libro de Pensiones de Rentas Vitalicias"
   '     Rpt_Reporte.Action = 1
   ' End If
    Screen.MousePointer = 0

Exit Function
Err_pension:
   Screen.MousePointer = 0
    If Err.Number <> 0 Then
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]" & glNumpol, vbCritical, "¡ERROR!..."
    End If
End Function

Function flInformePensionesInt()
On Error GoTo Err_pension
Dim objRep As New ClsReporte
Dim pol As String
Dim WL As Integer
Dim fecCierre As String
Dim NombreExcel As String
fecCierre = Mid(Format(Txt_Desde, "YYYYMMDD"), 1, 6)
NombreExcel = "INTERNO POR AFP - " & fecCierre
WL = 1000

'    '200912 DCM
'    If vgFormatoAFP = "" Then
'        vlArchivo = strRpt & "PP_Rpt_CieLibroPensRtasVitInt.rpt"      '\Reportes
'
'    Else
'        vlArchivo = strRpt & "PP_Rpt_CieLibroPensRtasVitIntAFP.rpt"   '\Reportes
'     End If
    
    
 
'    If Not fgExiste(vlArchivo) Then     ', vbNormal
'        MsgBox "Archivo de Reporte no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
'        Screen.MousePointer = 0
'        Exit Function
'    End If

    If Trim(vlGlosaOpcion) = "DEF" Then
        TipoProceso = "DEFINITIVO"
    Else
        TipoProceso = "PROVISORIO"
    End If

    Screen.MousePointer = 11
    lbl_Indicador.Caption = "Cargando Data de Liquidación..."

    'Función que permite determinar los Datos del Libro de Pensiones
    Call flLlenaTablaTemporalPensiones(vlGlosaOpcion, vlPago)
    Call fgVigenciaQuiebra(Txt_Desde)

    Dim vlTipoProceso As String
    Dim vlGlosa As String
    Dim vlSql As String

    vlSql = " SELECT SUBSTR(A.FEC_PAGO,1,6) NUM_PERIODO, DES_VIAPAGO, DES_AFP, NUM_POLIZA, NUM_ORDEN, DES_PENSION, B.GLS_TIPOIDEN  || '-' || NUM_IDENBEN GLS_DOCBEN, A.COD_CUSSP,GLS_NOMBEN || ' '|| GLS_NOMSEGBEN || ' ' || GLS_PATBEN || ' ' || GLS_MATREC GLS_BENEF,"
    vlSql = vlSql & " A.COD_MONEDA, MTO_PENSION, MTO_BASEIMP, MTO_SALUD,  MTO_BASETRI, MTO_DESNOIMP, MTO_PENLIQ,"
    vlSql = vlSql & " C.GLS_TIPOIDEN || '-' || NUM_IDENREC GLS_DOCREC,GLS_NOMREC || ' ' || GLS_NOMSEGREC || ' ' || GLS_PATREC || ' ' || GLS_MATREC GLS_RECEPTOR "
    vlSql = vlSql & " FROM PP_TTMP_PENSIONES A"
    vlSql = vlSql & " LEFT JOIN MA_TPAR_TIPOIDEN B ON A.COD_TIPOIDENBEN=B.COD_TIPOIDEN"
    vlSql = vlSql & " LEFT JOIN MA_TPAR_TIPOIDEN C ON A.COD_TIPOIDENREC=C.COD_TIPOIDEN"
    vlSql = vlSql & " WHERE A.COD_USUARIO='" & vgUsuario & "'"
    vlSql = vlSql & " ORDER BY 3,4"
    
    Set vgRs = New ADODB.Recordset
    vgRs.CursorLocation = adUseClient
    vgRs.Open vlSql, vgConexionBD, adOpenStatic, adLockBatchOptimistic
    
    
    If Not (vgRs.EOF) Then
            
            Dim totReg As Long
            totReg = vgRs.RecordCount
            
            Dim o_Excel     As Object
            Dim o_Libro     As Object
            Dim o_Hoja      As Object
            Dim Columna     As Long
            Dim Col     As Long
            
            Set o_Excel = CreateObject("Excel.Application")
            Set o_Libro = o_Excel.Workbooks.Add
            Set o_Hoja = o_Libro.Worksheets.Add
            
            Dim porc As Double
            Dim i As Long
            Dim ColTot As Integer
            i = 0
            
            ColTot = vgRs.Fields.Count - 1
            'Para colocar los titulos
            For Col = 0 To ColTot
                Select Case Col + 1
                    Case 0
                    'fechas
                    o_Hoja.Columns(Col + 1).NumberFormat = "d/m/yyyy;@"
                    Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 17, 18
                    'texto
                    o_Hoja.Columns(Col + 1).NumberFormat = "@"
                    Case 11, 12, 13, 14, 15, 16
                    'numeros
                End Select
                
                o_Hoja.Cells(1, Col + 1).Font.Bold = True
                o_Hoja.Cells(1, Col + 1).Value = vgRs.Fields(Col).Name
            Next
            
            Dim celda As String
            Dim celfec As Date
            Do While Not vgRs.EOF 'i > 50
                If totReg > 0 Then
                    porc = (100 * i) / totReg
                Else
                    porc = 0
                End If
                lbl_Indicador.Caption = "Cargando... " & CStr(Int(porc)) & "% (" & i & "/" & totReg & ")"
                
                i = i + 1
                
                For Col = 0 To ColTot
                    Select Case Col + 1
                        Case 0
'                        fechas
'                            If CStr(vgRs.Fields(Col)) = "01/01/1900" Then
'                               'celda = vgRs.Fields(Col).Value
'                               celda = "01/01/1900"
'                               o_Hoja.Cells(i + 1, Col + 1).Value = celda
'                            Else
'                               celfec = CDate(vgRs.Fields(Col).Value)
'                               o_Hoja.Cells(i + 1, Col + 1).Value = celfec
'                            End If
                            
                        Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 17, 18
                        'texto
                        
                        
                            If IsNull(vgRs.Fields(2)) Then
                               pol = CStr(vgRs.Fields(2).Value)
                            End If
                            
                            If IsNull(vgRs.Fields(Col)) Then
                               celda = ""
                            Else
                               celda = CStr(vgRs.Fields(Col).Value)
                               'celda = vgRs.Fields(Col).Value
                            End If
                            o_Hoja.Cells(i + 1, Col + 1).Value = Trim(celda)
                        Case 11, 12, 13, 14, 15, 16
                        'numeros
                            If IsNull(vgRs.Fields(Col)) Then
                               celda = ""
                            Else
                               celda = CStr(vgRs.Fields(Col).Value)
                               'celda = vgRs.Fields(Col).Value
                            End If
                            o_Hoja.Cells(i + 1, Col + 1).Value = Trim(celda)
                    End Select
'                    'o_Hoja.Cells(i + 1, Columna + 1).NumberFormat = "@"
'                    Dim celda As String
'                    If IsNull(vgRs.Fields(Col)) Then
'                       celda = ""
'                    Else
'                       celda = CStr(vgRs.Fields(Col).Value)
'                    End If
'                    o_Hoja.Cells(i + 1, Col + 1).Value = celda
                Next
                vgRs.MoveNext
            Loop
            lbl_Indicador.Caption = "Archivo generado correctamente"
            'o_Excel.Visible = True 'para ver vista previa
            o_Excel.ActiveWorkbook.SaveAs "C:\Users\" & Get_User_Name & "\Documents\" & NombreExcel & "_" & TipoProceso & ".xlsx"
            vgRs.Close
            
    End If
    
    
    Set o_Hoja = Nothing
    Set o_Libro = Nothing
    o_Excel.Quit
    Set o_Excel = Nothing

    lbl_Indicador.Caption = "Se guardo en " & "C:\Users\" & Get_User_Name & "\Documents\"
    Screen.MousePointer = 0

'    vlSql = " select a.*, gls_elemento from PP_TTMP_PENSIONES a "
'    vlSql = vlSql & " join ma_tpar_tabcod b on a.cod_viapago=cod_elemento and b.cod_tabla='VPG' "
'    vlSql = vlSql & " WHERE COD_USUARIO='" & vgUsuario & "'"
'    Set vgRs = vgConexionBD.Execute(vlSql)
    
'    Dim LNGa As Long
'    LNGa = CreateFieldDefFile(vgRs, Replace(UCase(strRpt & "Estructura\PP_Rpt_CieLibroPensRtasVitIntAFP.rpt"), ".RPT", ".TTX"), 1)
'
'    If Trim(vlGlosaOpcion) = "DEF" Then
'        TipoProceso = "DEFINITIVO"
'    Else
'        TipoProceso = "PROVISORIO"
'    End If
'
'    If objRep.CargaReporte(strRpt & "", "PP_Rpt_CieLibroPensRtasVitIntAFP_2.rpt", "Informe Certificados Pendientes de reliquidación", vgRs, True, _
'                            ArrFormulas("GlosaQuiebra", vgNombreCompania), _
'                            ArrFormulas("TipoProceso", vgNombreSistema), _
'                            ArrFormulas("NombreCompania", vgNombreCompania), _
'                            ArrFormulas("NombreSistema", vgNombreSistema), _
'                            ArrFormulas("NombreSubSistema", vgNombreSubSistema)) = False Then
'
'        MsgBox "No se pudo abrir el reporte", vbInformation
'        Exit Function
'    End If


'    If vlSw = True Then
'
'        vgQuery = ""
'        vgQuery = "{PP_TTMP_PENSIONES.COD_USUARIO} = '" & Trim(vgUsuario) & "' "
'
'        Rpt_Reporte.Reset
'        Rpt_Reporte.WindowState = crptMaximized
'        Rpt_Reporte.ReportFileName = vlArchivo
'        Rpt_Reporte.Connect = vgRutaDataBase
'        Rpt_Reporte.SelectionFormula = vgQuery
'        Rpt_Reporte.Formulas(4) = ""
'        Rpt_Reporte.Formulas(0) = "NombreCompania='" & vgNombreCompania & "'"
'        Rpt_Reporte.Formulas(1) = "NombreSistema= '" & vgNombreSistema & "'"
'        Rpt_Reporte.Formulas(2) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
'
'        If Trim(vlGlosaOpcion) = "DEF" Then
'            Rpt_Reporte.Formulas(3) = "TipoProceso= 'DEFINITIVO' "
'        Else
'            Rpt_Reporte.Formulas(3) = "TipoProceso= 'PROVISORIO' "
'        End If
'        Rpt_Reporte.Formulas(4) = "GlosaQuiebra = '" & vgGlosaQuiebra & "'"
'
'        Rpt_Reporte.Destination = crptToWindow
'        Rpt_Reporte.WindowTitle = "Libro de Pensiones de Rentas Vitalicias"
'        Rpt_Reporte.Action = 1
'    End If
    Screen.MousePointer = 0

Exit Function
Err_pension:
   Screen.MousePointer = 0
    If Err.Number <> 0 Then
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End If
End Function
Function flInformePenDirecto()
On Error GoTo Err_pension
Dim objRep As New ClsReporte
Dim pol As String
Dim WL As Integer
Dim fecCierre As String
Dim NombreExcel As String
fecCierre = Mid(Format(Txt_Desde, "YYYYMMDD"), 1, 6)
NombreExcel = "INTERNO (PAGO DIRECTO) - " & fecCierre
WL = 1000
   
'   vlArchivo = strRpt & "PP_Rpt_CieLibroPensRtasIntDirecto.rpt"   '\Reportes
'
'   If Not fgExiste(vlArchivo) Then     ', vbNormal
'        MsgBox "Archivo de Reporte no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
'        Screen.MousePointer = 0
'        Exit Function
'    End If

    If Trim(vlGlosaOpcion) = "DEF" Then
        TipoProceso = "DEFINITIVO"
    Else
        TipoProceso = "PROVISORIO"
    End If
    
    Screen.MousePointer = 11
    lbl_Indicador.Caption = "Cargando Data de Liquidación..."

    'Función que permite determinar los Datos del Libro de Pensiones
    Call flLlenaTablaTempRptDirecto(vlGlosaOpcion, vlPago)
    Call fgVigenciaQuiebra(Txt_Desde)

    Dim vlTipoProceso As String
    Dim vlGlosa As String
    Dim vlSql As String

    vlSql = " SELECT SUBSTR(A.FEC_PAGO,1,6) NUM_PERIODO, NUM_POLIZA, NUM_ORDEN, DES_PENSION, B.GLS_TIPOIDEN  || '-' || NUM_IDENBEN GLS_DOCBEN, GLS_NOMBEN || ' '|| GLS_NOMSEGBEN || ' ' || GLS_PATBEN || ' ' || GLS_MATREC GLS_BENEF,"
    vlSql = vlSql & " MTO_PENSION, MTO_BASEIMP, MTO_SALUD, MTO_BASETRI, MTO_DESNOIMP, MTO_PENLIQ,"
    vlSql = vlSql & " C.GLS_TIPOIDEN || '-' || NUM_IDENREC GLS_DOCREC,GLS_NOMREC || ' ' || GLS_NOMSEGREC || ' ' || GLS_PATREC || ' ' || GLS_MATREC GLS_RECEPTOR, DES_AFP, COD_MONEDA, DES_VIAPAGO, DES_BANCO, TIPO_CUENTA, COD_MONBCO, NUM_CTABCO, NUM_CUENTA_CCI"
    vlSql = vlSql & " FROM PP_TTMP_PENSIONES A"
    vlSql = vlSql & " JOIN MA_TPAR_TIPOIDEN B ON A.COD_TIPOIDENBEN=B.COD_TIPOIDEN"
    vlSql = vlSql & " JOIN MA_TPAR_TIPOIDEN C ON A.COD_TIPOIDENREC=C.COD_TIPOIDEN"
    vlSql = vlSql & " WHERE A.COD_USUARIO='" & vgUsuario & "'"
    vlSql = vlSql & " ORDER BY 2"
    
    'Set vgRs = vgConexionBD.Execute(vlSql)
    Set vgRs = New ADODB.Recordset
    vgRs.CursorLocation = adUseClient
    vgRs.Open vlSql, vgConexionBD, adOpenStatic, adLockBatchOptimistic
    
    If Not (vgRs.EOF) Then
            
            Dim totReg As Long
            totReg = vgRs.RecordCount
            
            Dim o_Excel     As Object
            Dim o_Libro     As Object
            Dim o_Hoja      As Object
            Dim Columna     As Long
            Dim Col     As Long
            
            Set o_Excel = CreateObject("Excel.Application")
            Set o_Libro = o_Excel.Workbooks.Add
            Set o_Hoja = o_Libro.Worksheets.Add
            
            Dim porc As Double
            Dim i As Long
            Dim ColTot As Integer
            i = 0
            
            ColTot = vgRs.Fields.Count - 1
            'Para colocar los titulos
            For Col = 0 To ColTot
                Select Case Col + 1
                    Case 0
                    'fechas
                    o_Hoja.Columns(Col + 1).NumberFormat = "d/m/yyyy;@"
                    Case 1, 2, 3, 4, 5, 6, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22
                    'texto
                    o_Hoja.Columns(Col + 1).NumberFormat = "@"
                    Case 7, 8, 9, 10, 11, 12
                    'numeros
                End Select
                
                o_Hoja.Cells(1, Col + 1).Font.Bold = True
                o_Hoja.Cells(1, Col + 1).Value = vgRs.Fields(Col).Name
            Next
            
            Dim celda As String
            Dim celfec As Date
            Do While Not vgRs.EOF 'i > 50
                If totReg > 0 Then
                    porc = (100 * i) / totReg
                Else
                    porc = 0
                End If
                lbl_Indicador.Caption = "Cargando... " & CStr(Int(porc)) & "% (" & i & "/" & totReg & ")"
                
                i = i + 1
                
                For Col = 0 To ColTot
                    Select Case Col + 1
'                        Case 0
'                        fechas
'                            If CStr(vgRs.Fields(Col)) = "01/01/1900" Then
'                               'celda = vgRs.Fields(Col).Value
'                               celda = "01/01/1900"
'                               o_Hoja.Cells(i + 1, Col + 1).Value = celda
'                            Else
'                               celfec = CDate(vgRs.Fields(Col).Value)
'                               o_Hoja.Cells(i + 1, Col + 1).Value = celfec
'                            End If
                            
                        Case 1, 2, 3, 4, 5, 6, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22
                        'texto
                        
                        
                            If IsNull(vgRs.Fields(2)) Then
                               pol = CStr(vgRs.Fields(2).Value)
                            End If
                            
                            If IsNull(vgRs.Fields(Col)) Then
                               celda = ""
                            Else
                               celda = CStr(vgRs.Fields(Col).Value)
                               'celda = vgRs.Fields(Col).Value
                            End If
                            o_Hoja.Cells(i + 1, Col + 1).Value = Trim(celda)
                        Case 7, 8, 9, 10, 11, 12
                        'numeros
                            If IsNull(vgRs.Fields(Col)) Then
                               celda = ""
                            Else
                               celda = CStr(vgRs.Fields(Col).Value)
                               'celda = vgRs.Fields(Col).Value
                            End If
                            o_Hoja.Cells(i + 1, Col + 1).Value = Trim(celda)
                    End Select
'                    'o_Hoja.Cells(i + 1, Columna + 1).NumberFormat = "@"
'                    Dim celda As String
'                    If IsNull(vgRs.Fields(Col)) Then
'                       celda = ""
'                    Else
'                       celda = CStr(vgRs.Fields(Col).Value)
'                    End If
'                    o_Hoja.Cells(i + 1, Col + 1).Value = celda
                Next
                vgRs.MoveNext
            Loop
            lbl_Indicador.Caption = "Archivo generado correctamente"
            'o_Excel.Visible = True 'para ver vista previa
            o_Excel.ActiveWorkbook.SaveAs "C:\Users\" & Get_User_Name & "\Documents\" & NombreExcel & "_" & TipoProceso & ".xlsx"
            vgRs.Close
            
    End If
    
    
    Set o_Hoja = Nothing
    Set o_Libro = Nothing
    o_Excel.Quit
    Set o_Excel = Nothing

    lbl_Indicador.Caption = "Se guardo en " & "C:\Users\" & Get_User_Name & "\Documents\"
    Screen.MousePointer = 0
'    vlSql = " select a.*, gls_elemento from PP_TTMP_PENSIONES a "
'    vlSql = vlSql & " join ma_tpar_tabcod b on a.cod_viapago=cod_elemento and b.cod_tabla='VPG' "
'    vlSql = vlSql & " WHERE COD_USUARIO='" & vgUsuario & "'"
'    Set vgRs = vgConexionBD.Execute(vlSql)
'
'    Dim LNGa As Long
'    LNGa = CreateFieldDefFile(vgRs, Replace(UCase(strRpt & "Estructura\PP_Rpt_CieLibroPensRtasIntDirecto.rpt"), ".RPT", ".TTX"), 1)
'
'    If Trim(vlGlosaOpcion) = "DEF" Then
'        TipoProceso = "DEFINITIVO"
'    Else
'        TipoProceso = "PROVISORIO"
'    End If
'
'    If objRep.CargaReporte(strRpt & "", "PP_Rpt_CieLibroPensRtasIntDirecto.rpt", "Informe Certificados Pendientes de reliquidación", vgRs, True, _
'                            ArrFormulas("GlosaQuiebra", vgNombreCompania), _
'                            ArrFormulas("TipoProceso", vgNombreSistema), _
'                            ArrFormulas("NombreCompania", vgNombreCompania), _
'                            ArrFormulas("NombreSistema", vgNombreSistema), _
'                            ArrFormulas("NombreSubSistema", vgNombreSubSistema)) = False Then
'
'        MsgBox "No se pudo abrir el reporte", vbInformation
'        Exit Function
'    End If
'

Exit Function
Err_pension:
   Screen.MousePointer = 0
    If Err.Number <> 0 Then
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End If
End Function

Function flInformeDefRetJud()

On Error GoTo Err_flInformeDefRetJud

'''    vlPeriodo = DateSerial(Mid(vlFecha, 1, 4), Mid(vlFecha, 5, 2), Mid(vlFecha, 7, 2))

    vlSql = "DELETE FROM PP_TTMP_CONRETJUD WHERE cod_usuario = '" & vgUsuario & "'"
    vgConexionBD.Execute (vlSql)

    Call flCargaTablaTemporal
    
   ''Buscar codigos para seleccionar Otros Haberes y Otros Descuentos
   'vlTipoMov = clCodHab
   'Call flBuscaCodOtrosHabDes(vlTipoMov)
   'If vlCodOtrosHabDes <> "" Then
   '   Call flAgregarOtrosHabDes
   'End If
   'vlTipoMov = clCodDescto
   'Call flBuscaCodOtrosHabDes(vlTipoMov)
   'If vlCodOtrosHabDes <> "" Then
   '   Call flAgregarOtrosHabDes
   'End If
                 
   Screen.MousePointer = 11
   
   vlArchivo = strRpt & "PP_Rpt_DefRetJud2.rpt"   '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
      MsgBox "Archivo de Reporte Definitivo de Retención Judicial no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
      Screen.MousePointer = 0
      Exit Function
   End If
   
   'Call fgVigenciaQuiebra(Txt_Desde)
   
   'vgQuery = ""
   'vgQuery = "{PP_TTMP_CONRETJUD.COD_USUARIO} = '" & Trim(vgUsuario) & "' AND "
   'vgQuery = vgQuery & "{MA_TVAL_MONEDA.cod_moneda} = '" & Trim(cgCodTipMonedaUF) & "' "
   '
   'Rpt_Reporte.Reset
   'Rpt_Reporte.WindowState = crptMaximized
   'Rpt_Reporte.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
   'Rpt_Reporte.Connect = vgRutaDataBase       ' o App.Path & "\Nestle.mdb"
   'Rpt_Reporte.SelectionFormula = vgQuery
  '
   'Rpt_Reporte.Formulas(0) = ""
   'Rpt_Reporte.Formulas(1) = ""
   'Rpt_Reporte.Formulas(2) = ""
   'Rpt_Reporte.Formulas(3) = ""
   'Rpt_Reporte.Formulas(4) = ""
  '
  ' Rpt_Reporte.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
   'Rpt_Reporte.Formulas(1) = "NombreSistema= '" & vgNombreSistema & "'"
   'Rpt_Reporte.Formulas(2) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
   'Rpt_Reporte.Formulas(3) = ""
   
   'If Trim(vlGlosaOpcion) = "DEF" Then
   '    Rpt_Reporte.Formulas(3) = "TipoProceso= 'DEFINITIVO' "
   'Else
   '    Rpt_Reporte.Formulas(3) = "TipoProceso= 'PROVISORIO' "
   'End If
   
   'Rpt_Reporte.Formulas(4) = "GlosaQuiebra = '" & vgGlosaQuiebra & "'"
  '
   'Rpt_Reporte.SubreportToChange = ""
   'Rpt_Reporte.Destination = crptToWindow
   'Rpt_Reporte.WindowState = crptMaximized
   'Rpt_Reporte.WindowTitle = "Informe Definitivo de Retención Judicial"
   'Rpt_Reporte.SelectionFormula = ""
   'Rpt_Reporte.Action = 1
   'Screen.MousePointer = 0
    
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
   Dim objRep As New ClsReporte


    Screen.MousePointer = 11
    
    'Dim vlSql As String

    vlSql = "select a.*, b.num_poliza, b.num_orden, b.mto_ret, b.num_idenreceptor, b.gls_nomreceptor, b.gls_nomsegreceptor, b.gls_patreceptor, b.gls_matreceptor, d.cod_moneda,"
    vlSql = vlSql & " c.gls_nomben, c.gls_nomsegben, c.gls_matben, c.gls_patben, d.mto_valmoneda, e.gls_tipoiden , f.gls_elemento as viapago, g.gls_elemento as modoret, h.gls_elemento as moneda, i.gls_elemento as tipocuenta,"
    vlSql = vlSql & " j.gls_elemento as banco , b.Num_Cuenta, d.cod_cuspp, b.cod_modret"
    vlSql = vlSql & " from pp_ttmp_conretjud a"
    vlSql = vlSql & " join pp_tmae_retjudicial b on a.num_retencion=b.num_retencion"
    vlSql = vlSql & " join pp_tmae_ben c on c.num_poliza=b.num_poliza and c.num_endoso=b.num_endoso and c.num_orden=b.num_orden"
    vlSql = vlSql & " join pp_tmae_poliza d on d.num_poliza=c.num_poliza and d.num_endoso=c.num_endoso"
    vlSql = vlSql & " join ma_tpar_tipoiden e on e.cod_tipoiden=c.cod_tipoidenben"
    vlSql = vlSql & " join ma_tpar_tabcod f on b.cod_viapago=f.cod_elemento and f.cod_tabla='VPG'"
    vlSql = vlSql & " join ma_tpar_tabcod g on b.cod_modret=g.cod_elemento and g.cod_tabla='MPR'"
    vlSql = vlSql & " join ma_tpar_tabcod h on d.cod_moneda=h.cod_elemento and h.cod_tabla='MP'"
    vlSql = vlSql & " join ma_tpar_tabcod i on b.cod_tipcuenta=i.cod_elemento and i.cod_tabla='TCT'"
    vlSql = vlSql & " join ma_tpar_tabcod j on b.cod_banco=j.cod_elemento and j.cod_tabla='BCO' order by b.num_poliza, a.num_retencion"
    Set vgRs = vgConexionBD.Execute(vlSql)
    
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(vgRs, Replace(UCase(strRpt & "Estructura\PP_Rpt_DefRetJud.rpt"), ".RPT", ".TTX"), 1)

    If Trim(vlGlosaOpcion) = "DEF" Then
        TipoProceso = "DEFINITIVO"
    Else
        TipoProceso = "PROVISORIO"
    End If

    If objRep.CargaReporte(strRpt & "", "PP_Rpt_DefRetJud2.rpt", "Informe de retenciones del periodo de pago.", vgRs, True, _
                            ArrFormulas("GlosaQuiebra", vlGlosaOpcion), _
                            ArrFormulas("TipoProceso", TipoProceso), _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("NombreSistema", vgNombreSistema), _
                            ArrFormulas("NombreSubSistema", vgNombreSubSistema)) = False Then

        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Function
    End If
    
    
    
Exit Function
Err_flInformeDefRetJud:
    Screen.MousePointer = 0
    If Err.Number <> 0 Then
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End If

End Function

Function flCargaTablaTemporal()

On Error GoTo Err_flCargaTablaTemporal
     
    vlNumCargas = 0
    vlMtoCargas = 0
    vlMtoHaber = 0
    vlMtoDescto = 0
    
    vgSql = ""
    vgSql = vgSql & "SELECT l.num_perpago,l.fec_pago,l.num_poliza,l.num_orden,"
    vgSql = vgSql & "l.cod_tipopago,p.cod_conhabdes,p.mto_conhabdes,"
    vgSql = vgSql & "p.cod_tipoidenreceptor,p.num_idenreceptor " 'MC 31-08-2007 p.rut_receptor
    vgSql = vgSql & "FROM PP_TMAE_LIQPAGOPEN" & vlGlosaOpcion & " l, "
    vgSql = vgSql & "PP_TMAE_PAGOPEN" & vlGlosaOpcion & " p "
    vgSql = vgSql & "WHERE l.fec_pago >= '" & vlFechaInicio & "' AND "
    vgSql = vgSql & "l.fec_pago <= '" & vlFechaTermino & "' AND "
    vgSql = vgSql & "l.cod_tipopago = '" & vlPago & "' AND "
    vgSql = vgSql & "p.cod_conhabdes IN " & vlCodHabDesRJC & " AND "
    vgSql = vgSql & "l.num_perpago = p.num_perpago AND "
    vgSql = vgSql & "l.num_poliza = p.num_poliza AND "
    vgSql = vgSql & "l.num_orden = p.num_orden AND "
    'I - MC 31-08-2007
    'vgSql = vgSql & "l.rut_receptor = p.rut_receptor AND "
    'vgSql = vgSql & "l.cod_tipoidenreceptor= p.cod_tipoidenreceptor AND "
    'vgSql = vgSql & "l.num_idenreceptor = p.num_idenreceptor AND "
    'F - MC 31-08-2007
    vgSql = vgSql & "P.cod_tipreceptor = 'R' "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
       
       
       While Not vgRs.EOF
       
             vlNumCargas = 0
             vlMtoCargas = 0
             vlMtoHaber = 0
             vlMtoDescto = 0
             vlMtoRetJud = 0
              
             If (vgRs!Cod_ConHabDes) = clCodHabDes62 Then
                vlTipoRetJud = clTipoRAF
             Else
                 vlTipoRetJud = clTipoRJ
             End If
             
             'Buscar el Nùmero de Retenciòn de cada uno de los detalles seleccionados.
             
             vgSql = ""
             vgSql = "SELECT num_retencion FROM PP_TMAE_RETJUDICIAL "
             vgSql = vgSql & "WHERE num_poliza = '" & (vgRs!num_poliza) & "' AND "
             'vgSql = vgSql & "num_endoso = " & (vgRs!num_endoso) & " AND "
             vgSql = vgSql & "num_orden = " & (vgRs!Num_Orden) & " AND "
             'I - MC 31-08-2007
             'vgSql = vgSql & "rut_receptor = " & (vgRs!Rut_Receptor) & " AND "
             vgSql = vgSql & "cod_tipoidenreceptor = " & (vgRs!Cod_TipoIdenReceptor) & " AND "
             vgSql = vgSql & "num_idenreceptor = '" & (vgRs!Num_IdenReceptor) & "' AND "
             'F - MC 31-08-2007
             vgSql = vgSql & "cod_tipret in ('RJ', 'RJ1', 'RJ2', 'RJ3') AND "
             If vgTipoBase = "ORACLE" Then
                vgSql = vgSql & "substr(fec_iniret,1,6) <= '" & (vgRs!Num_PerPago) & "' AND "
                vgSql = vgSql & "substr(fec_terret,1,6) >= '" & (vgRs!Num_PerPago) & "' "
             Else
                 vgSql = vgSql & "substring(fec_iniret,1,6) <= '" & (vgRs!Num_PerPago) & "' AND "
                 vgSql = vgSql & "substring(fec_terret,1,6) >= '" & (vgRs!Num_PerPago) & "' "
             End If
             Set vgRegistro = vgConexionBD.Execute(vgSql)
             If Not vgRegistro.EOF Then
                vlNumRetencion = (vgRegistro!num_retencion)
                vlNumOrdenCar = 0
                If vlTipoRetJud = clTipoRAF Then
                   'Buscar Nùmero de Orden correspondiente a Conyuge
                   vgSql = ""
                   vgSql = "SELECT num_orden FROM PP_TMAE_DETRETENCION "
                   vgSql = vgSql & "WHERE num_retencion = " & str(vlNumRetencion) & " "
                   Set vgRegistro = vgConexionBD.Execute(vgSql)
                   vlNumOrdenCar = (vgRegistro!Num_Orden)
                End If
                
                If (vgRs!Cod_ConHabDes) = clCodHabDes62 Then
                   vlNumCargas = 1
                Else
                    'Call flBuscaNumCargas
                    vlNumCargas = 1
                End If
                
                If ((vgRs!Cod_ConHabDes) = clCodHabDes61) Or ((vgRs!Cod_ConHabDes) = clCodHabDes62) Then
                   vlMtoCargas = (vgRs!Mto_ConHabDes)
                Else
                    vlMtoRetJud = (vgRs!Mto_ConHabDes)
                End If
                
                vlMtoHaber = 0
                vlMtoDescto = 0
               
                Call flAgregarTablaTemporal
                
             End If
             
             vgRs.MoveNext
       
       Wend
       
    End If

Exit Function
Err_flCargaTablaTemporal:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function


Function flBuscaCodOtrosHabDes(Codigo As String)

On Error GoTo flBuscaCodOtrosHabDes
'Busca códigos de otros haberes y otros descuentos por concepto de RJC
'Cobro de Retención Judicial
'Código = H(Haber) o D(Descuento)
    
    vlCodOtrosHabDes = ""
    vgSql = ""
    vgSql = vgSql & "SELECT cod_conhabdes "
    vgSql = vgSql & "FROM ma_tpar_conhabdes "
    vgSql = vgSql & "WHERE cod_modorigen = '" & clCodRJC & "' AND "
    vgSql = vgSql & "cod_tipmov = '" & Trim(Codigo) & "' AND "
    vgSql = vgSql & "cod_conhabdes NOT IN " & vlCodHabDesRJC & " "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       While Not vgRegistro.EOF
             If vlCodOtrosHabDes = "" Then
                vlCodOtrosHabDes = "("
             End If
             vlCodOtrosHabDes = (vlCodOtrosHabDes & "'" & (vgRegistro!Cod_ConHabDes) & "'")
             vgRegistro.MoveNext
             If Not vgRegistro.EOF Then
                vlCodOtrosHabDes = (vlCodOtrosHabDes & ",")
             End If
       Wend
       vlCodOtrosHabDes = (vlCodOtrosHabDes & ")")
    End If

Exit Function
flBuscaCodOtrosHabDes:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function flAgregarOtrosHabDes()

On Error GoTo Err_flAgregarOtrosHabDes

    vgSql = ""
    vgSql = "SELECT num_poliza,num_endoso,num_orden,cod_conhabdes, "
    vgSql = vgSql & "mto_conhabdes,"
    vgSql = vgSql & "cod_tipoidenreceptor,num_idenreceptor " 'MC - 31-08-2007 Rut_Receptor
    vgSql = vgSql & "FROM PP_TMAE_PAGOPEN" & vlGlosaOpcion & " "
'    vgSql = vgSql & "FROM PP_TTMP_CONPAGOPEN "
    vgSql = vgSql & "WHERE cod_conhabdes IN " & vlCodOtrosHabDes & " "
    
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       While Not vgRegistro.EOF
       
             vlNumCargas = 0
             vlMtoCargas = 0
             vlMtoHaber = 0
             vlMtoDescto = 0
             vlMtoHabDes = 0
              
             'Buscar el Nùmero de Retenciòn de cada uno de los detalles seleccionados.
             
             vgSql = ""
             vgSql = "SELECT num_retencion FROM PP_TMAE_RETJUDICIAL "
             vgSql = vgSql & "WHERE num_poliza = " & (vgRs!num_poliza) & " AND "
             vgSql = vgSql & "num_endoso = " & (vgRs!num_endoso) & " AND "
             vgSql = vgSql & "num_orden = " & (vgRs!Num_Orden) & " AND "
             'I - MC 31-08-2007
             'vgSql = vgSql & "rut_receptor = (vgrs!rut_receptor) AND "
             vgSql = vgSql & "cod_tipoidenreceptor = " & (vgRs!Cod_TipoIdenReceptor) & " AND "
             vgSql = vgSql & "num_idenreceptor = '" & (vgRs!Num_IdenReceptor) & "' AND "
             'F - MC 31-08-2007
             
             vgSql = vgSql & "cod_tipret = '" & Trim(vlTipoRetJud) & "' AND "
             If vgTipoBase = "ORACLE" Then
                vgSql = vgSql & "substr(fec_iniret,1,6) <= '" & (vgRs!Num_PerPago) & "' AND "
                vgSql = vgSql & "substr(fec_terret,1,6) >= '" & (vgRs!Num_PerPago) & "' "
             Else
                 vgSql = vgSql & "substring(fec_iniret,1,6) <= '" & (vgRs!Num_PerPago) & "' AND "
                 vgSql = vgSql & "substring(fec_terret,1,6) >= '" & (vgRs!Num_PerPago) & "' "
             End If
             
             Set vgRegistro = vgConexionBD.Execute(vgSql)
             If Not vgRegistro.EOF Then
                vlNumRetencion = (vgRegistro!num_retencion)
                vlMtoHabDes = (vgRs!Mto_ConHabDes)
                Call flAgregarTablaTemporal
             End If
             
             vgRs.MoveNext
       
       Wend
       
    End If

Exit Function
Err_flAgregarOtrosHabDes:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function



Function flBuscaNumCargas()

On Error GoTo Err_flBuscaNumCargas
'Busca el número de cargas pagadas por concepto de Retención
'Judicial de Asignacion Familiar (Solo Hijos, sin Conyuge)

    vgSql = ""
    vgSql = "SELECT count(num_ordencar) as numcargas "
    vgSql = vgSql & "FROM PP_TMAE_PAGOASIG" & vlGlosaOpcion & " "
    vgSql = vgSql & "WHERE num_perpago = '" & (vgRs!Num_PerPago) & "' AND "
    vgSql = vgSql & "num_poliza = '" & Trim(vgRs!num_poliza) & "' AND "
    vgSql = vgSql & "num_orden = '" & Trim(vgRs!Num_Orden) & "' AND "
    vgSql = vgSql & "rut_receptor = '" & Trim(vgRs!Rut_Receptor) & "' AND "
    vgSql = vgSql & "num_ordencar <> '" & vlNumOrdenCar & "' "
    
    Set vgRegistro = vgConexionBD.Execute(vgSql)

    If Not vgRegistro.EOF Then
       vlNumCargas = (vgRegistro!numcargas)
    End If

Exit Function
Err_flBuscaNumCargas:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function


Function flAgregarTablaTemporal()

On Error GoTo Err_flAgregarTablaTemporal

    vgSql = ""
    vgSql = "SELECT * "
    vgSql = vgSql & "FROM PP_TTMP_CONRETJUD "
    vgSql = vgSql & "WHERE num_retencion = '" & Trim(vlNumRetencion) & "' "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If vgRegistro.EOF Then
       
       vgSql = ""
       vgSql = "INSERT INTO PP_TTMP_CONRETJUD "
       vgSql = vgSql & "(cod_usuario,num_retencion,num_cargas, "
       vgSql = vgSql & " mto_carga,mto_haber,mto_descto,mto_retneta, "
       vgSql = vgSql & " fec_pago,num_perpago "
       vgSql = vgSql & " ) VALUES ( "
       vgSql = vgSql & " '" & vgUsuario & "', "
       vgSql = vgSql & " '" & Trim(vlNumRetencion) & "' , "
       vgSql = vgSql & " " & str(vlNumCargas) & ", "
       vgSql = vgSql & " " & str(vlMtoCargas) & ", "
       vgSql = vgSql & " " & str(vlMtoHaber) & ", "
       vgSql = vgSql & " " & str(vlMtoDescto) & ", "
       vgSql = vgSql & " " & str(vlMtoRetJud) & ", "
       vgSql = vgSql & " '" & (vgRs!Fec_Pago) & "', "
       vgSql = vgSql & " '" & (vgRs!Num_PerPago) & "') "
       vgConexionBD.Execute vgSql
       
    Else
        vlNumCargasAux = (vgRegistro!Num_Cargas)
        vlMtoCargasAux = (vgRegistro!Mto_Carga)
        vlMtoHaberAux = (vgRegistro!Mto_Haber)
        vlMtoDesctoAux = (vgRegistro!mto_descto)
        vlMtoRetNetaAux = (vgRegistro!mto_retneta)
        If vlNumCargasAux = 0 Then
            vlNumCargasAux = vlNumCargasAux + vlNumCargas
        End If
        If vlMtoCargasAux = 0 Then
            vlMtoCargasAux = vlMtoCargasAux + vlMtoCargas
        End If
        If vlMtoRetNetaAux = 0 Then
            vlMtoRetNetaAux = vlMtoRetNetaAux + vlMtoRetJud
        End If
        vgSql = ""
        vgSql = " UPDATE PP_TTMP_CONRETJUD SET "
        vgSql = vgSql & "num_cargas = " & str(vlNumCargasAux) & ", "
        vgSql = vgSql & "mto_carga = " & str(vlMtoCargasAux) & ", "
        vgSql = vgSql & "mto_retneta = " & str(vlMtoRetNetaAux) & " "
        vgSql = vgSql & "WHERE num_retencion = '" & Trim(vlNumRetencion) & "' "
        vgConexionBD.Execute vgSql
        
'        If vlTipoRetJud = clTipoRAF Then
'           vgSql = ""
'           vgSql = " UPDATE PP_TTMP_CONRETJUD SET "
'           vgSql = vgSql & "mto_carga = " & Str(vlMtoCargas) & " "
'           vgSql = vgSql & "WHERE num_retencion = '" & Trim(vlNumRetencion) & "' "
'           vgConexionBD.Execute vgSql
'        End If
'        If vlTipoRetJud = clTipoRJ Then
'           vgSql = ""
'           vgSql = " UPDATE PP_TTMP_CONRETJUD SET "
'           vgSql = vgSql & "mto_retneta = " & Str(vlMtoRetJud) & " "
'           vgSql = vgSql & "WHERE num_retencion = '" & Trim(vlNumRetencion) & "' "
'           vgConexionBD.Execute vgSql
'        End If
        If vlTipoMov = clCodHab Then
            If vlMtoHaberAux = 0 Then
                vlMtoHaberAux = vlMtoHaberAux + vlMtoHabDes
            End If
           vgSql = ""
           vgSql = " UPDATE PP_TTMP_CONRETJUD SET "
           vgSql = vgSql & "mto_haber = " & str(vlMtoHaberAux) & " "
           vgSql = vgSql & "WHERE num_retencion = '" & Trim(vlNumRetencion) & "' "
           vgConexionBD.Execute vgSql
        End If
        If vlTipoMov = clCodDescto Then
            If vlMtoDesctoAux = 0 Then
                vlMtoDesctoAux = vlMtoDesctoAux + vlMtoHabDes
            End If
            vgSql = ""
            vgSql = " UPDATE PP_TTMP_CONRETJUD SET "
            vgSql = vgSql & "mto_descto = " & str(vlMtoDesctoAux) & " "
            vgSql = vgSql & "WHERE num_retencion = '" & Trim(vlNumRetencion) & "' "
            vgConexionBD.Execute vgSql
        End If
             
    End If
    
Exit Function
Err_flAgregarTablaTemporal:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
    
End Function
Function flLlenaTablaTemporalPensiones(iTabla As String, iPago As String)
Dim vlMtoPension   As String
'dim vlMtoPensionPesos As String
Dim vlMtoPensionUF As String
'Dim vlMtoHabImp    As String, vlMtoDesImp   As String
'Dim vlMtoHabNoImp  As String, vlMtoDesNoImp As String
'Dim vlMtoImpuesto  As String, vlMtoSalud    As String
Dim vlBaseImp      As Double, vlBaseTri     As Double
Dim vlMtoPenLiq    As Double
'dim vlCodTipReceptor As String,
'Dim vlNumPol       As String, vlNumPerPago  As String, vlNumOrden As String
'On Error GoTo Err_LlenaTabla

    vlSw = False
    
'    'Limpiar variables de totales para resumen
'    Call flLimpiarVarTotTipPen
    
    'Borrar tabla temporal
    vlSql = "DELETE FROM PP_TTMP_PENSIONES "
    vlSql = vlSql & "WHERE COD_USUARIO = '" & vgUsuario & "'"
    vgConexionBD.Execute vlSql
    
'    'Consulta si existen polizas para ese rango de fechas
'    vlSql = "SELECT l.cod_tipreceptor,L.NUM_POLIZA,L.NUM_ORDEN,L.COD_TIPPENSION,"
'    vlSql = vlSql & "B.COD_TIPOIDENBEN,B.NUM_IDENBEN,"
'    vlSql = vlSql & "B.GLS_NOMBEN,B.GLS_NOMSEGBEN,B.GLS_PATBEN,B.GLS_MATBEN,"
'    vlSql = vlSql & "B.FEC_TERPAGOPENGAR,L.MTO_BASEIMP,"
'    vlSql = vlSql & "L.MTO_BASETRI,L.MTO_LIQPAGAR,L.FEC_PAGO, L.NUM_PERPAGO,L.COD_TIPRECEPTOR,"
'    vlSql = vlSql & "B.MTO_PENSION, B.MTO_PENSIONGAR,b.cod_grufam, "
'    vlSql = vlSql & "l.mto_HABER,l.COD_TIPOIDENRECEPTOR,l.NUM_IDENreceptor,l.num_endoso, "
'    vlSql = vlSql & "L.COD_MONEDA, P.FEC_CALPAGOREG "
'    vlSql = vlSql & ",POL.cod_afp,B.cod_viapago " '27/06/2008
'    vlSql = vlSql & ",A.GLS_ELEMENTO AS DES_AFP"  '200912 DCM
'    vlSql = vlSql & ",B.GLS_ELEMENTO AS DES_PENSION"  '200912 DCM
'    vlSql = vlSql & ",COD_CUSPP"
'     'Inicio GCP 29032019
'    vlSql = vlSql & ",B.COD_TIPCTA"
'    vlSql = vlSql & ",B.COD_SUCURSAL"
'    vlSql = vlSql & ",B.COD_BANCO"
'    vlSql = vlSql & ",BCO.gls_elemento AS DES_BANCO"
'    vlSql = vlSql & ",B.COD_MONBCO"
'    vlSql = vlSql & ",B.NUM_CTABCO"
'    vlSql = vlSql & ",B.NUM_CUENTA_CCI"
'    vlSql = vlSql & ",VP.GLS_ELEMENTO AS DES_VIAPAGO"
'    vlSql = vlSql & ",TC.GLS_ELEMENTO AS TIPO_CUENTA"
'    vlSql = vlSql & ",TRIM(L.GLS_NOMRECEPTOR)||' '||TRIM(L.GLS_NOMSEGRECEPTOR)||' '||TRIM(L.GLS_PATRECEPTOR)||' '||TRIM(L.GLS_MATRECEPTOR) AS NOMBRE_RECEPTOR"
'   'Fin GCP 29032019
'    vlSql = vlSql & " FROM "
'    vlSql = vlSql & "PP_TMAE_LIQPAGOPEN" & iTabla & " L, PP_TMAE_BEN B, "
'    vlSql = vlSql & "PP_TMAE_PROPAGOPEN P ,PP_TMAE_POLIZA POL, " '27/06/2008
'    vlSql = vlSql & "MA_TPAR_TABCOD A , MA_TPAR_TABCOD B "   '200912 DCM
'    vlSql = vlSql & ",MA_TPAR_TABCOD BCO, "  'GCP 29032019
'    vlSql = vlSql & "MA_TPAR_TABCOD VP,"
'    vlSql = vlSql & "MA_TPAR_TABCOD TC "
'    vlSql = vlSql & "WHERE "
'    vlSql = vlSql & "L.fec_pago >= '" & vlFechaDesde & "' and "
'    vlSql = vlSql & "L.fec_pago <= '" & vlFechaHasta & "' and "
'    vlSql = vlSql & "L.cod_tipopago = '" & iPago & "' and "
'    vlSql = vlSql & "L.num_poliza = B.num_poliza and "
'    vlSql = vlSql & "L.num_endoso = B.num_endoso and "
'    vlSql = vlSql & "L.num_orden = B.num_orden  and "
'    vlSql = vlSql & "L.num_perpago = P.num_perpago "
'    vlSql = vlSql & "and POL.num_poliza=B.num_poliza and POL.num_endoso=B.num_endoso " '27/06/2008
''    If vlPago = "R" Then
''        vlSql = vlSql & " and P.fec_pagoreg = M.fec_moneda  "
''    Else
''        vlSql = vlSql & " and P.fec_pripago = M.fec_moneda "
''    End If
'    If vgNombreInformeSeleccionado <> "InfLibPen" Then
'        vlSql = vlSql & " L.COD_TIPPENSION not IN ('04','05','09','10')  AND"
'    End If
'    vlSql = vlSql & " AND A.COD_TABLA='AF' AND A.COD_ELEMENTO=POL.COD_AFP" '200912 DCM
'    vlSql = vlSql & " AND B.COD_TABLA='TP' AND B.COD_ELEMENTO=L.COD_TIPPENSION" '200912 DCM
'    vlSql = vlSql & " AND L.MTO_PENSION<>0" '20140526 RRR
'     'Inicio GCP 29032019
'
'    vlSql = vlSql & " AND B.COD_BANCO  = BCO.COD_ELEMENTO "
'    vlSql = vlSql & " AND BCO.COD_TABLA = 'BCO' AND"
'    'vlSql = vlSql & " AND B.COD_VIAPAGO = '02' AND"
'    'vlSql = vlSql & " L.COD_TIPPENSION IN ('04','05','09','10')  AND"
'    vlSql = vlSql & " B.COD_VIAPAGO  = VP.COD_ELEMENTO AND"
'    vlSql = vlSql & " VP.COD_TABLA = 'VPG' AND"
'    vlSql = vlSql & " B.COD_TIPCTA  =  TC.COD_ELEMENTO AND"
'    vlSql = vlSql & " TC.COD_TABLA = 'TCT'"
'   'Fin GCP 29032019
    
    
    vlSql = "SELECT l.cod_tipreceptor,L.NUM_POLIZA,L.NUM_ORDEN,L.COD_TIPPENSION,B.COD_TIPOIDENBEN,B.NUM_IDENBEN,B.GLS_NOMBEN,B.GLS_NOMSEGBEN,B.GLS_PATBEN,B.GLS_MATBEN,B.FEC_TERPAGOPENGAR,L.MTO_BASEIMP,L.MTO_BASETRI,L.MTO_LIQPAGAR,L.FEC_PAGO,"
    vlSql = vlSql & " L.NUM_PERPAGO,L.COD_TIPRECEPTOR,B.MTO_PENSION, B.MTO_PENSIONGAR,b.cod_grufam, l.mto_HABER,l.COD_TIPOIDENRECEPTOR,l.NUM_IDENreceptor,l.num_endoso, L.COD_MONEDA, P.FEC_CALPAGOREG ,POL.cod_afp,B.cod_viapago ,"
    vlSql = vlSql & " A.GLS_ELEMENTO AS DES_AFP,TP.GLS_ELEMENTO AS DES_PENSION,COD_CUSPP,B.COD_TIPCTA,B.COD_SUCURSAL,B.COD_BANCO,BCO.gls_elemento AS DES_BANCO,B.COD_MONBCO,B.NUM_CTABCO,B.NUM_CUENTA_CCI,VP.GLS_ELEMENTO AS DES_VIAPAGO,TC.GLS_ELEMENTO AS TIPO_CUENTA,"
    vlSql = vlSql & " TRIM(L.GLS_NOMRECEPTOR)||' '||TRIM(L.GLS_NOMSEGRECEPTOR)||' '||TRIM(L.GLS_PATRECEPTOR)||' '||TRIM(L.GLS_MATRECEPTOR) AS NOMBRE_RECEPTOR"
    vlSql = vlSql & " FROM PP_TMAE_LIQPAGOPEN" & iTabla & " L"
    vlSql = vlSql & " JOIN PP_TMAE_BEN B ON L.num_poliza = B.num_poliza and L.num_orden = B.num_orden"
    vlSql = vlSql & " JOIN PP_TMAE_PROPAGOPEN P ON L.num_perpago = P.num_perpago"
    vlSql = vlSql & " JOIN PP_TMAE_POLIZA POL ON POL.num_poliza=B.num_poliza and POL.num_endoso=B.num_endoso"
    vlSql = vlSql & " LEFT JOIN MA_TPAR_TABCOD A ON A.COD_TABLA='AF' AND A.COD_ELEMENTO=POL.COD_AFP"
    vlSql = vlSql & " LEFT JOIN MA_TPAR_TABCOD TP ON TP.COD_TABLA='TP' AND TP.COD_ELEMENTO=L.COD_TIPPENSION"
    vlSql = vlSql & " LEFT JOIN MA_TPAR_TABCOD BCO ON BCO.COD_TABLA = 'BCO' AND  B.COD_BANCO= BCO.COD_ELEMENTO"
    vlSql = vlSql & " LEFT JOIN MA_TPAR_TABCOD VP ON VP.COD_TABLA = 'VPG' AND B.COD_VIAPAGO  = VP.COD_ELEMENTO"
    vlSql = vlSql & " LEFT JOIN MA_TPAR_TABCOD TC ON TC.COD_TABLA = 'TCT' AND B.COD_TIPCTA  =  TC.COD_ELEMENTO"
    vlSql = vlSql & " WHERE "
    vlSql = vlSql & " L.fec_pago >= '" & vlFechaDesde & "' and "
    vlSql = vlSql & " L.fec_pago <= '" & vlFechaHasta & "' and "
    vlSql = vlSql & " L.cod_tipopago = '" & iPago & "' "
    If vgNombreInformeSeleccionado <> "InfLibPen" Then
        vlSql = vlSql & " AND L.COD_TIPPENSION not IN ('04','05','09','10')"
    End If
    vlSql = vlSql & " AND POL.NUM_ENDOSO=(SELECT MAX(NUM_ENDOSO) FROM PP_TMAE_POLIZA WHERE NUM_POLIZA=L.NUM_POLIZA) "
    vlSql = vlSql & " AND COD_TIPRECEPTOR<>'R'"
    Set vgRs = vgConexionBD.Execute(vlSql)
    If Not vgRs.EOF Then
        Do While Not vgRs.EOF
            
            vlNumPol = (vgRs!num_poliza)
            vlNumEndoso = (vgRs!num_endoso)
            vlNumOrden = (vgRs!Num_Orden)
            vlNumPerPago = (vgRs!Num_PerPago)
            vlTipoReceptor = (vgRs!Cod_TipReceptor)
            
            vlCodGruFam = (vgRs!Cod_GruFam)
            vlCodTipReceptor = (vgRs!Cod_TipReceptor)
            vlCodTipoIdenReceptor = (vgRs!Cod_TipoIdenReceptor)
            vlNumIdenReceptor = (vgRs!Num_IdenReceptor)
            vlCodTipoIdenRec = ""
            vlNumIdenRec = ""
            vlGlsNomRec = ""
            vlGlsNomSegRec = ""
            vlGlsPatRec = ""
            vlGlsMatRec = ""
            vlDesAFP = vgRs!DES_AFP  '200912   DCM
            vlDesPension = vgRs!des_Pension '200912   DCM
            
            If vlNumPol = "0000000510" Then
                glNumpol = vlNumPol
            End If
            
            'Buscar Nombre del Receptor del pago de la póliza
            Call flBuscarNombreReceptor(vlCodTipoIdenRec, vlNumIdenRec, vlGlsNomRec, vlGlsNomSegRec, vlGlsPatRec, vlGlsMatRec, _
                    vlCodTipReceptor, vlNumPol, vlNumEndoso)
                    
            'CMV/20050705 I
            vlMtoHabImp = 0
            vlMtoDesImp = 0
            vlMtoHabNoImp = 0
            vlMtoDesNoImp = 0
            vlMtoSalud = 0
            vlMtoImpuesto = 0
            vlMtoPensionPesos = 0
            Call flBuscarMontosConceptos(iTabla)
          
            'CMV/20050705 F
            
            'vlMtoPensionUF = Format(CDbl(vlMtoPensionPesos) / (vgRs!mto_moneda), "#0.00")
            vlMtoPensionUF = Format((vgRs!Mto_Haber), "#0.00")
            
            'calculo de Pension Liquida
            'I---- ABV 04/09/2004 ---
            'vlBaseImp = CDbl(vlMtoPensionPesos) + CDbl(vlMtoHabImp) - CDbl(vlMtoDesImp)
            'vlBaseTri = vlBaseImp - CDbl(vlMtoSalud)
            'vlMtoPenLiq = Format(vlBaseTri - CDbl(vlMtoImpuesto) + CDbl(vlMtoHabNoImp) - CDbl(vlMtoDesNoImp), "#0.00")
            vlBaseImp = IIf(IsNull(vgRs!Mto_BaseImp), 0, vgRs!Mto_BaseImp)
            vlBaseTri = IIf(IsNull(vgRs!Mto_BaseTri), 0, vgRs!Mto_BaseTri)
            vlMtoPenLiq = IIf(IsNull(vgRs!Mto_LiqPagar), 0, vgRs!Mto_LiqPagar)
            'F---- ABV 04/09/2004 ---
            
            'I - MC 27/06/2008
            vlCodAFP = IIf(IsNull(vgRs!cod_afp), "", vgRs!cod_afp)
            vlViaPago = IIf(IsNull(vgRs!Cod_ViaPago), "", vgRs!Cod_ViaPago)
            vlFecPago = Mid(vgRs!Fec_Pago, 1, 6) & "01"
            vlFactorAju = flObtieneFactorAjuste(iPago)
            'F - MC 27/06/2008
            
            'RRR 16/08/2012
            vlCuspp = IIf(IsNull(vgRs!Cod_Cuspp), "", vgRs!Cod_Cuspp)
            vlPerPago = IIf(IsNull(vgRs!Num_PerPago), "", vgRs!Num_PerPago)
            
            'Inicio GCP 29032019
            VlCOD_TIPCTA = IIf(IsNull(vgRs!cod_tipcta), "", vgRs!cod_tipcta)
            VlCOD_SUCURSAL = IIf(IsNull(vgRs!Cod_Sucursal), "", vgRs!Cod_Sucursal)
            VlCOD_MONBCO = IIf(IsNull(vgRs!cod_monbco), "", vgRs!cod_monbco)
            VlNUM_CTABCO = IIf(IsNull(vgRs!num_ctabco), "", vgRs!num_ctabco)
            VlDES_BANCO = IIf(IsNull(vgRs!DES_BANCO), "", vgRs!DES_BANCO)
            VlNUM_CUENTA_CCI = IIf(IsNull(vgRs!NUM_CUENTA_CCI), "", vgRs!NUM_CUENTA_CCI)
            VlCOD_BANCO = IIf(IsNull(vgRs!Cod_Banco), "", vgRs!Cod_Banco)
            vlDES_VIAPAGO = IIf(IsNull(vgRs!DES_VIAPAGO), "", vgRs!DES_VIAPAGO)
            vlTIPO_CUENTA = IIf(IsNull(vgRs!TIPO_CUENTA), "", vgRs!TIPO_CUENTA)
            
            'Fin GCP 29032019
            
            'grabar en la tabla temporal
            vlSql = "INSERT INTO PP_TTMP_PENSIONES ("
            vlSql = vlSql & "COD_USUARIO,"
            vlSql = vlSql & "NUM_POLIZA,NUM_ORDEN,FEC_PAGO,cod_moneda,COD_TIPPENSION,"
            vlSql = vlSql & "cod_tipoidenBEN,num_idenBEN,GLS_NOMBEN,GLS_NOMSEGBEN,GLS_PATBEN,GLS_MATBEN,"
            vlSql = vlSql & "MTO_PENSION,MTO_PENSIONPESOS,MTO_HABIMP,"
            vlSql = vlSql & "MTO_DESIMP,MTO_BASEIMP,MTO_BASETRI,MTO_SALUD,"
            vlSql = vlSql & "MTO_IMPUESTO,MTO_HABNOIMP,MTO_DESNOIMP,MTO_PENLIQ, "
            vlSql = vlSql & "Cod_TipReceptor,gls_nomrec,gls_nomsegrec,gls_patrec,gls_matrec, "
            vlSql = vlSql & "cod_tipoidenrec,num_idenrec "
            vlSql = vlSql & ",cod_afp,cod_viapago,prc_factoraju " '21/06/2008
            vlSql = vlSql & ",DES_AFP, DES_PENSION, COD_CUSSP, NUM_PERPAGO"   '200912 DCM
            'Inicio GCP 29032019
            vlSql = vlSql & ",COD_SUCURSAL"
            vlSql = vlSql & ",Cod_Banco"
            vlSql = vlSql & ",COD_TIPCTA"
            vlSql = vlSql & ",COD_MONBCO"
            vlSql = vlSql & ",NUM_CTABCO"
            vlSql = vlSql & ",NUM_CUENTA_CCI"
            vlSql = vlSql & ",DES_BANCO"
            vlSql = vlSql & ",DES_VIAPAGO"
            vlSql = vlSql & ",TIPO_CUENTA"
           'Fin GCP 29032019
            vlSql = vlSql & ") VALUES ("
            vlSql = vlSql & "'" & vgUsuario & "',"
            vlSql = vlSql & "'" & vlNumPol & "',"
            vlSql = vlSql & " " & vlNumOrden & ","
            vlSql = vlSql & "'" & vgRs!Fec_Pago & "',"
            vlSql = vlSql & "'" & vgRs!Cod_Moneda & "',"
            vlSql = vlSql & "'" & vgRs!Cod_TipPension & "',"
            vlSql = vlSql & " " & (vgRs!Cod_TipoIdenBen) & ","
            vlSql = vlSql & "'" & vgRs!Num_IdenBen & "',"
            vlSql = vlSql & "'" & vgRs!Gls_NomBen & "',"
            If IsNull(vgRs!Gls_NomSegBen) Then
                vlSql = vlSql & "NULL,"
            Else
                vlSql = vlSql & "'" & vgRs!Gls_NomSegBen & "',"
            End If
            vlSql = vlSql & "'" & vgRs!Gls_PatBen & "',"
            If IsNull(vgRs!Gls_MatBen) Then
                vlSql = vlSql & "NULL,"
            Else
                vlSql = vlSql & "'" & vgRs!Gls_MatBen & "',"
            End If
            vlSql = vlSql & " " & str(Format(vlMtoPensionUF, "#0.00")) & " ,"
            vlSql = vlSql & " " & str(vlMtoPensionPesos) & " ,"
            vlSql = vlSql & " " & str(vlMtoHabImp) & " ,"
            vlSql = vlSql & " " & str(vlMtoDesImp) & " ,"
            vlSql = vlSql & " " & str(Format(vgRs!Mto_BaseImp, "#0.00")) & " ,"
            vlSql = vlSql & " " & str(Format(vgRs!Mto_BaseTri, "#0.00")) & " ,"
            vlSql = vlSql & " " & str(vlMtoSalud) & " ,"
            vlSql = vlSql & " " & str(vlMtoImpuesto) & " ,"
            vlSql = vlSql & " " & str(vlMtoHabNoImp) & " ,"
            vlSql = vlSql & " " & str(vlMtoDesNoImp) & " ,"
            vlSql = vlSql & " " & str(Format(vlMtoPenLiq, "#0.00")) & ","
            vlSql = vlSql & "'" & Trim(vlCodTipReceptor) & "',"
            vlSql = vlSql & "'" & Trim(vlGlsNomRec) & "',"
            If (vlGlsNomSegRec = "") Then
                vlSql = vlSql & "NULL,"
            Else
                vlSql = vlSql & "'" & Trim(vlGlsNomSegRec) & "',"
            End If
            vlSql = vlSql & "'" & Trim(vlGlsPatRec) & "',"
            If (vlGlsMatRec = "") Then
                vlSql = vlSql & "NULL,"
            Else
                vlSql = vlSql & "'" & Trim(vlGlsMatRec) & "',"
            End If
            vlSql = vlSql & "'" & Trim(vlCodTipoIdenRec) & "',"
            vlSql = vlSql & "'" & Trim(vlNumIdenRec) & "',"
            vlSql = vlSql & "'" & Trim(vlCodAFP) & "'," '27/06/2008
            vlSql = vlSql & "'" & Trim(vlViaPago) & "',"
            vlSql = vlSql & " " & str(vlFactorAju) & " "
            vlSql = vlSql & ", '" & Trim(vlDesAFP) & "' " '200912 DCM
            vlSql = vlSql & ", '" & Trim(vlDesPension) & "' " '200912 DCM
            vlSql = vlSql & ", '" & Trim(vlCuspp) & "' "
            vlSql = vlSql & ", '" & Trim(vlPerPago) & "' "
             'Inicio GCP 29032019
            vlSql = vlSql & ", '" & Trim(VlCOD_SUCURSAL) & "' "
            vlSql = vlSql & ", '" & Trim(VlCOD_BANCO) & "' "
            vlSql = vlSql & ", '" & Trim(VlCOD_TIPCTA) & "' "
            vlSql = vlSql & ", '" & Trim(VlCOD_MONBCO) & "' "
            vlSql = vlSql & ", '" & Trim(VlNUM_CTABCO) & "' "
            vlSql = vlSql & ", '" & Trim(VlNUM_CUENTA_CCI) & "' "
            vlSql = vlSql & ", '" & Trim(VlDES_BANCO) & "'"
            vlSql = vlSql & ", '" & Trim(vlDES_VIAPAGO) & "'"
            vlSql = vlSql & ", '" & Trim(vlTIPO_CUENTA) & "'"
            'Fin GCP 29032019
            vlSql = vlSql & " )"
            
 
            vgConexionBD.Execute vlSql
            
'            Call flAgregarTotalesTipPen
            
            vlSw = True
            vgRs.MoveNext
        Loop
    Else
        MsgBox "No existe Información para mostrar" & vlNumPol, vbExclamation, "Operación Cancelada"
        Screen.MousePointer = 0
        Exit Function
    End If
    vgRs.Close
    
'Exit Function
'Err_LlenaTabla:
'    Screen.MousePointer = 0
'    If Err.Number <> 0 Then
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End If
End Function

Function flLlenaTablaTempRptDirecto(iTabla As String, iPago As String)
Dim vlMtoPension   As String
'dim vlMtoPensionPesos As String
Dim vlMtoPensionUF As String
'Dim vlMtoHabImp    As String, vlMtoDesImp   As String
'Dim vlMtoHabNoImp  As String, vlMtoDesNoImp As String
'Dim vlMtoImpuesto  As String, vlMtoSalud    As String
Dim vlBaseImp      As Double, vlBaseTri     As Double
Dim vlMtoPenLiq    As Double
'dim vlCodTipReceptor As String,
'Dim vlNumPol       As String, vlNumPerPago  As String, vlNumOrden As String
'On Error GoTo Err_LlenaTabla

    vlSw = False
    
'    'Limpiar variables de totales para resumen
'    Call flLimpiarVarTotTipPen
    
    'Borrar tabla temporal
    vlSql = "DELETE FROM PP_TTMP_PENSIONES "
    vlSql = vlSql & "WHERE COD_USUARIO = '" & vgUsuario & "'"
    vgConexionBD.Execute vlSql
    
'    'Consulta si existen polizas para ese rango de fechas
'    vlSql = "SELECT l.cod_tipreceptor,L.NUM_POLIZA,L.NUM_ORDEN,L.COD_TIPPENSION,"
'    vlSql = vlSql & "B.COD_TIPOIDENBEN,B.NUM_IDENBEN,"
'    vlSql = vlSql & "B.GLS_NOMBEN,B.GLS_NOMSEGBEN,B.GLS_PATBEN,B.GLS_MATBEN,"
'    vlSql = vlSql & "B.FEC_TERPAGOPENGAR,L.MTO_BASEIMP,"
'    vlSql = vlSql & "L.MTO_BASETRI,L.MTO_LIQPAGAR,L.FEC_PAGO, L.NUM_PERPAGO,L.COD_TIPRECEPTOR,"
'    vlSql = vlSql & "B.MTO_PENSION, B.MTO_PENSIONGAR,b.cod_grufam, "
'    vlSql = vlSql & "l.mto_HABER,l.COD_TIPOIDENRECEPTOR,l.NUM_IDENreceptor,l.num_endoso, "
'    vlSql = vlSql & "L.COD_MONEDA, P.FEC_CALPAGOREG "
'    vlSql = vlSql & ",POL.cod_afp,B.cod_viapago " '27/06/2008
'    vlSql = vlSql & ",A.GLS_ELEMENTO AS DES_AFP"  '200912 DCM
'    vlSql = vlSql & ",B.GLS_ELEMENTO AS DES_PENSION"  '200912 DCM
'    vlSql = vlSql & ",COD_CUSPP"
'    'Inicio GCP 29032019
'    vlSql = vlSql & ",B.COD_TIPCTA"
'    vlSql = vlSql & ",B.COD_SUCURSAL"
'    vlSql = vlSql & ",B.COD_BANCO"
'    vlSql = vlSql & ",BCO.gls_elemento AS DES_BANCO"
'    vlSql = vlSql & ",B.COD_MONBCO"
'    vlSql = vlSql & ",B.NUM_CTABCO"
'    vlSql = vlSql & ",B.NUM_CUENTA_CCI"
'    vlSql = vlSql & ",VP.GLS_ELEMENTO AS DES_VIAPAGO"
'    vlSql = vlSql & ",TC.GLS_ELEMENTO AS TIPO_CUENTA"
'    vlSql = vlSql & ",TRIM(L.GLS_NOMRECEPTOR)||' '||TRIM(L.GLS_NOMSEGRECEPTOR)||' '||TRIM(L.GLS_PATRECEPTOR)||' '||TRIM(L.GLS_MATRECEPTOR) AS NOMBRE_RECEPTOR"
'   'Fin GCP 29032019
'    vlSql = vlSql & " FROM "
'    vlSql = vlSql & "PP_TMAE_LIQPAGOPEN" & iTabla & " L, PP_TMAE_BEN B, "
'    vlSql = vlSql & "PP_TMAE_PROPAGOPEN P ,PP_TMAE_POLIZA POL, " '27/06/2008
'    vlSql = vlSql & "MA_TPAR_TABCOD A , MA_TPAR_TABCOD B, "   '200912 DCM
'    vlSql = vlSql & "MA_TPAR_TABCOD BCO, "  'GCP 29032019
'    vlSql = vlSql & "MA_TPAR_TABCOD VP,"
'    vlSql = vlSql & "MA_TPAR_TABCOD TC "
'    vlSql = vlSql & "WHERE "
'    vlSql = vlSql & "L.fec_pago >= '" & vlFechaDesde & "' and "
'    vlSql = vlSql & "L.fec_pago <= '" & vlFechaHasta & "' and "
'    vlSql = vlSql & "L.cod_tipopago = '" & iPago & "' and "
'    vlSql = vlSql & "L.num_poliza = B.num_poliza and "
'    vlSql = vlSql & "L.num_endoso = B.num_endoso and "
'    vlSql = vlSql & "L.num_orden = B.num_orden  and "
'    vlSql = vlSql & "L.num_perpago = P.num_perpago "
'    vlSql = vlSql & "and POL.num_poliza=B.num_poliza and POL.num_endoso=B.num_endoso " '27/06/2008
'    vlSql = vlSql & " AND A.COD_TABLA='AF' AND A.COD_ELEMENTO=POL.COD_AFP" '200912 DCM
'    vlSql = vlSql & " AND B.COD_TABLA='TP' AND B.COD_ELEMENTO=L.COD_TIPPENSION" '200912 DCM
'    vlSql = vlSql & " AND L.MTO_PENSION<>0" '20140526 RRR
'    'Inicio GCP 29032019
'    vlSql = vlSql & " AND B.COD_BANCO = BCO.COD_ELEMENTO "
'    vlSql = vlSql & " AND BCO.COD_TABLA = 'BCO' AND"
'    'vlSql = vlSql & " AND B.COD_VIAPAGO = '02' AND"
'    vlSql = vlSql & " L.COD_TIPPENSION IN ('04','05','09','10')  AND"
'    vlSql = vlSql & " B.COD_VIAPAGO = VP.COD_ELEMENTO AND"
'    vlSql = vlSql & " VP.COD_TABLA = 'VPG' AND"
'    vlSql = vlSql & " B.COD_TIPCTA =  TC.COD_ELEMENTO AND"
'    vlSql = vlSql & " TC.COD_TABLA = 'TCT'"
'   'Fin GCP 29032019
    
    vlSql = "SELECT l.cod_tipreceptor,L.NUM_POLIZA,L.NUM_ORDEN,L.COD_TIPPENSION,B.COD_TIPOIDENBEN,B.NUM_IDENBEN,B.GLS_NOMBEN,B.GLS_NOMSEGBEN,B.GLS_PATBEN,B.GLS_MATBEN,B.FEC_TERPAGOPENGAR,L.MTO_BASEIMP,L.MTO_BASETRI,L.MTO_LIQPAGAR,L.FEC_PAGO,"
    vlSql = vlSql & " L.NUM_PERPAGO,L.COD_TIPRECEPTOR,B.MTO_PENSION, B.MTO_PENSIONGAR,b.cod_grufam, l.mto_HABER,l.COD_TIPOIDENRECEPTOR,l.NUM_IDENreceptor,l.num_endoso, L.COD_MONEDA, P.FEC_CALPAGOREG ,POL.cod_afp,B.cod_viapago ,"
    vlSql = vlSql & " A.GLS_ELEMENTO AS DES_AFP,TP.GLS_ELEMENTO AS DES_PENSION,COD_CUSPP,B.COD_TIPCTA,B.COD_SUCURSAL,B.COD_BANCO,BCO.gls_elemento AS DES_BANCO,B.COD_MONBCO,B.NUM_CTABCO,B.NUM_CUENTA_CCI,VP.GLS_ELEMENTO AS DES_VIAPAGO,TC.GLS_ELEMENTO AS TIPO_CUENTA,"
    vlSql = vlSql & " TRIM(L.GLS_NOMRECEPTOR)||' '||TRIM(L.GLS_NOMSEGRECEPTOR)||' '||TRIM(L.GLS_PATRECEPTOR)||' '||TRIM(L.GLS_MATRECEPTOR) AS NOMBRE_RECEPTOR"
    vlSql = vlSql & " FROM PP_TMAE_LIQPAGOPEN" & iTabla & " L"
    vlSql = vlSql & " JOIN PP_TMAE_BEN B ON L.num_poliza = B.num_poliza and L.num_orden = B.num_orden"
    vlSql = vlSql & " JOIN PP_TMAE_PROPAGOPEN P ON L.num_perpago = P.num_perpago"
    vlSql = vlSql & " JOIN PP_TMAE_POLIZA POL ON POL.num_poliza=B.num_poliza and POL.num_endoso=B.num_endoso"
    vlSql = vlSql & " LEFT JOIN MA_TPAR_TABCOD A ON A.COD_TABLA='AF' AND A.COD_ELEMENTO=POL.COD_AFP"
    vlSql = vlSql & " LEFT JOIN MA_TPAR_TABCOD TP ON TP.COD_TABLA='TP' AND TP.COD_ELEMENTO=L.COD_TIPPENSION"
    vlSql = vlSql & " LEFT JOIN MA_TPAR_TABCOD BCO ON BCO.COD_TABLA = 'BCO' AND  B.COD_BANCO= BCO.COD_ELEMENTO"
    vlSql = vlSql & " LEFT JOIN MA_TPAR_TABCOD VP ON VP.COD_TABLA = 'VPG' AND B.COD_VIAPAGO  = VP.COD_ELEMENTO"
    vlSql = vlSql & " LEFT JOIN MA_TPAR_TABCOD TC ON TC.COD_TABLA = 'TCT' AND B.COD_TIPCTA  =  TC.COD_ELEMENTO"
    vlSql = vlSql & " WHERE "
    vlSql = vlSql & " L.fec_pago >= '" & vlFechaDesde & "' and "
    vlSql = vlSql & " L.fec_pago <= '" & vlFechaHasta & "' and "
    vlSql = vlSql & " L.cod_tipopago = '" & iPago & "' and "
    vlSql = vlSql & " L.COD_TIPPENSION IN ('04','05','09','10')  and "
    vlSql = vlSql & " POL.NUM_ENDOSO=(SELECT MAX(NUM_ENDOSO) FROM PP_TMAE_POLIZA WHERE NUM_POLIZA=L.NUM_POLIZA) and "
    vlSql = vlSql & " L.COD_TIPRECEPTOR<>'R'"
    
    Set vgRs = vgConexionBD.Execute(vlSql)
    If Not vgRs.EOF Then
        Do While Not vgRs.EOF
            
            vlNumPol = (vgRs!num_poliza)
            vlNumEndoso = (vgRs!num_endoso)
            vlNumOrden = (vgRs!Num_Orden)
            vlNumPerPago = (vgRs!Num_PerPago)
            vlTipoReceptor = (vgRs!Cod_TipReceptor)
            
            vlCodGruFam = (vgRs!Cod_GruFam)
            vlCodTipReceptor = (vgRs!Cod_TipReceptor)
            vlCodTipoIdenReceptor = (vgRs!Cod_TipoIdenReceptor)
            vlNumIdenReceptor = (vgRs!Num_IdenReceptor)
            vlCodTipoIdenRec = ""
            vlNumIdenRec = ""
            vlGlsNomRec = ""
            vlGlsNomSegRec = ""
            vlGlsPatRec = ""
            vlGlsMatRec = ""
            vlDesAFP = vgRs!DES_AFP  '200912   DCM
            vlDesPension = vgRs!des_Pension '200912   DCM
            
            If vlNumPol = "0000000510" Then
                glNumpol = vlNumPol
            End If
            
            'Buscar Nombre del Receptor del pago de la póliza
            Call flBuscarNombreReceptor(vlCodTipoIdenRec, vlNumIdenRec, vlGlsNomRec, vlGlsNomSegRec, vlGlsPatRec, vlGlsMatRec, _
                    vlCodTipReceptor, vlNumPol, vlNumEndoso)
                    
            'CMV/20050705 I
            vlMtoHabImp = 0
            vlMtoDesImp = 0
            vlMtoHabNoImp = 0
            vlMtoDesNoImp = 0
            vlMtoSalud = 0
            vlMtoImpuesto = 0
            vlMtoPensionPesos = 0
            Call flBuscarMontosConceptos(iTabla)
          
            'CMV/20050705 F
            
            'vlMtoPensionUF = Format(CDbl(vlMtoPensionPesos) / (vgRs!mto_moneda), "#0.00")
            vlMtoPensionUF = Format((vgRs!Mto_Haber), "#0.00")
            
            'calculo de Pension Liquida
            'I---- ABV 04/09/2004 ---
            'vlBaseImp = CDbl(vlMtoPensionPesos) + CDbl(vlMtoHabImp) - CDbl(vlMtoDesImp)
            'vlBaseTri = vlBaseImp - CDbl(vlMtoSalud)
            'vlMtoPenLiq = Format(vlBaseTri - CDbl(vlMtoImpuesto) + CDbl(vlMtoHabNoImp) - CDbl(vlMtoDesNoImp), "#0.00")
            vlBaseImp = IIf(IsNull(vgRs!Mto_BaseImp), 0, vgRs!Mto_BaseImp)
            vlBaseTri = IIf(IsNull(vgRs!Mto_BaseTri), 0, vgRs!Mto_BaseTri)
            vlMtoPenLiq = IIf(IsNull(vgRs!Mto_LiqPagar), 0, vgRs!Mto_LiqPagar)
            'F---- ABV 04/09/2004 ---
            
            'I - MC 27/06/2008
            vlCodAFP = IIf(IsNull(vgRs!cod_afp), "", vgRs!cod_afp)
            vlViaPago = IIf(IsNull(vgRs!Cod_ViaPago), "", vgRs!Cod_ViaPago)
            vlFecPago = Mid(vgRs!Fec_Pago, 1, 6) & "01"
            vlFactorAju = flObtieneFactorAjuste(iPago)
            'F - MC 27/06/2008
            
            'RRR 16/08/2012
            vlCuspp = IIf(IsNull(vgRs!Cod_Cuspp), "", vgRs!Cod_Cuspp)
            vlPerPago = IIf(IsNull(vgRs!Num_PerPago), "", vgRs!Num_PerPago)
            
            
            'Inicio GCP 29032019
            VlCOD_TIPCTA = IIf(IsNull(vgRs!cod_tipcta), "", vgRs!cod_tipcta)
            VlCOD_SUCURSAL = IIf(IsNull(vgRs!Cod_Sucursal), "", vgRs!Cod_Sucursal)
            VlCOD_MONBCO = IIf(IsNull(vgRs!cod_monbco), "", vgRs!cod_monbco)
            VlNUM_CTABCO = IIf(IsNull(vgRs!num_ctabco), "", vgRs!num_ctabco)
            VlDES_BANCO = IIf(IsNull(vgRs!DES_BANCO), "", vgRs!DES_BANCO)
            VlNUM_CUENTA_CCI = IIf(IsNull(vgRs!NUM_CUENTA_CCI), "", vgRs!NUM_CUENTA_CCI)
            VlCOD_BANCO = IIf(IsNull(vgRs!Cod_Banco), "", vgRs!Cod_Banco)
            vlDES_VIAPAGO = IIf(IsNull(vgRs!DES_VIAPAGO), "", vgRs!DES_VIAPAGO)
            vlTIPO_CUENTA = IIf(IsNull(vgRs!TIPO_CUENTA), "", vgRs!TIPO_CUENTA)
            
            'Fin GCP 29032019
 
            
            'grabar en la tabla temporal
            vlSql = "INSERT INTO PP_TTMP_PENSIONES ("
            vlSql = vlSql & "COD_USUARIO,"
            vlSql = vlSql & "NUM_POLIZA,NUM_ORDEN,FEC_PAGO,cod_moneda,COD_TIPPENSION,"
            vlSql = vlSql & "cod_tipoidenBEN,num_idenBEN,GLS_NOMBEN,GLS_NOMSEGBEN,GLS_PATBEN,GLS_MATBEN,"
            vlSql = vlSql & "MTO_PENSION,MTO_PENSIONPESOS,MTO_HABIMP,"
            vlSql = vlSql & "MTO_DESIMP,MTO_BASEIMP,MTO_BASETRI,MTO_SALUD,"
            vlSql = vlSql & "MTO_IMPUESTO,MTO_HABNOIMP,MTO_DESNOIMP,MTO_PENLIQ, "
            vlSql = vlSql & "Cod_TipReceptor,gls_nomrec,gls_nomsegrec,gls_patrec,gls_matrec, "
            vlSql = vlSql & "cod_tipoidenrec,num_idenrec "
            vlSql = vlSql & ",cod_afp,cod_viapago,prc_factoraju " '21/06/2008
            vlSql = vlSql & ",DES_AFP, DES_PENSION, COD_CUSSP, NUM_PERPAGO"   '200912 DCM
            'Inicio GCP 29032019
            vlSql = vlSql & ",COD_SUCURSAL"
            vlSql = vlSql & ",Cod_Banco"
            vlSql = vlSql & ",COD_TIPCTA"
            vlSql = vlSql & ",COD_MONBCO"
            vlSql = vlSql & ",NUM_CTABCO"
            vlSql = vlSql & ",NUM_CUENTA_CCI"
            vlSql = vlSql & ",DES_BANCO"
            vlSql = vlSql & ",DES_VIAPAGO"
            vlSql = vlSql & ",TIPO_CUENTA"
           'Fin GCP 29032019
            vlSql = vlSql & ") VALUES ("
            vlSql = vlSql & "'" & vgUsuario & "',"
            vlSql = vlSql & "'" & vlNumPol & "',"
            vlSql = vlSql & " " & vlNumOrden & ","
            vlSql = vlSql & "'" & vgRs!Fec_Pago & "',"
            vlSql = vlSql & "'" & vgRs!Cod_Moneda & "',"
            vlSql = vlSql & "'" & vgRs!Cod_TipPension & "',"
            vlSql = vlSql & " " & (vgRs!Cod_TipoIdenBen) & ","
            vlSql = vlSql & "'" & vgRs!Num_IdenBen & "',"
            vlSql = vlSql & "'" & vgRs!Gls_NomBen & "',"
            If IsNull(vgRs!Gls_NomSegBen) Then
                vlSql = vlSql & "NULL,"
            Else
                vlSql = vlSql & "'" & vgRs!Gls_NomSegBen & "',"
            End If
            vlSql = vlSql & "'" & vgRs!Gls_PatBen & "',"
            If IsNull(vgRs!Gls_MatBen) Then
                vlSql = vlSql & "NULL,"
            Else
                vlSql = vlSql & "'" & vgRs!Gls_MatBen & "',"
            End If
            vlSql = vlSql & " " & str(Format(vlMtoPensionUF, "#0.00")) & " ,"
            vlSql = vlSql & " " & str(vlMtoPensionPesos) & " ,"
            vlSql = vlSql & " " & str(vlMtoHabImp) & " ,"
            vlSql = vlSql & " " & str(vlMtoDesImp) & " ,"
            vlSql = vlSql & " " & str(Format(vgRs!Mto_BaseImp, "#0.00")) & " ,"
            vlSql = vlSql & " " & str(Format(vgRs!Mto_BaseTri, "#0.00")) & " ,"
            vlSql = vlSql & " " & str(vlMtoSalud) & " ,"
            vlSql = vlSql & " " & str(vlMtoImpuesto) & " ,"
            vlSql = vlSql & " " & str(vlMtoHabNoImp) & " ,"
            vlSql = vlSql & " " & str(vlMtoDesNoImp) & " ,"
            vlSql = vlSql & " " & str(Format(vlMtoPenLiq, "#0.00")) & ","
            vlSql = vlSql & "'" & Trim(vlCodTipReceptor) & "',"
            vlSql = vlSql & "'" & Trim(vlGlsNomRec) & "',"
            If (vlGlsNomSegRec = "") Then
                vlSql = vlSql & "NULL,"
            Else
                vlSql = vlSql & "'" & Trim(vlGlsNomSegRec) & "',"
            End If
            vlSql = vlSql & "'" & Trim(vlGlsPatRec) & "',"
            If (vlGlsMatRec = "") Then
                vlSql = vlSql & "NULL,"
            Else
                vlSql = vlSql & "'" & Trim(vlGlsMatRec) & "',"
            End If
            vlSql = vlSql & "'" & Trim(vlCodTipoIdenRec) & "',"
            vlSql = vlSql & "'" & Trim(vlNumIdenRec) & "',"
            vlSql = vlSql & "'" & Trim(vlCodAFP) & "'," '27/06/2008
            vlSql = vlSql & "'" & Trim(vlViaPago) & "',"
            vlSql = vlSql & " " & str(vlFactorAju) & " "
            vlSql = vlSql & ", '" & Trim(vlDesAFP) & "' " '200912 DCM
            vlSql = vlSql & ", '" & Trim(vlDesPension) & "' " '200912 DCM
            vlSql = vlSql & ", '" & Trim(vlCuspp) & "' "
            vlSql = vlSql & ", '" & Trim(vlPerPago) & "' "
            'Inicio GCP 29032019
            vlSql = vlSql & ", '" & Trim(VlCOD_SUCURSAL) & "' "
            vlSql = vlSql & ", '" & Trim(VlCOD_BANCO) & "' "
            vlSql = vlSql & ", '" & Trim(VlCOD_TIPCTA) & "' "
            vlSql = vlSql & ", '" & Trim(VlCOD_MONBCO) & "' "
            vlSql = vlSql & ", '" & Trim(VlNUM_CTABCO) & "' "
            vlSql = vlSql & ", '" & Trim(VlNUM_CUENTA_CCI) & "' "
            vlSql = vlSql & ", '" & Trim(VlDES_BANCO) & "'"
             vlSql = vlSql & ", '" & Trim(vlDES_VIAPAGO) & "'"
              vlSql = vlSql & ", '" & Trim(vlTIPO_CUENTA) & "'"
            'Fin GCP 29032019
           vlSql = vlSql & " )"
              
            vgConexionBD.Execute vlSql
            
'            Call flAgregarTotalesTipPen
            
            vlSw = True
            vgRs.MoveNext
        Loop
    Else
        MsgBox "No existe Información para mostrar", vbExclamation, "Operación Cancelada"
        Screen.MousePointer = 0
        Exit Function
    End If
    vgRs.Close
    
'Exit Function
'Err_LlenaTabla:
'    Screen.MousePointer = 0
'    If Err.Number <> 0 Then
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End If
End Function


Function flBuscarNombreReceptor(iRutRec As String, iDgvRec As String, iNombre As String, iNombreSeg As String, _
                                iPaterno As String, iMaterno As String, iCodTipReceptor As String, _
                                iNumPol As String, inumendoso As Integer)
On Error GoTo Err_flBuscarNombreReceptor

    'Receptor de Tipo PENSIONADO
    If iCodTipReceptor = clTipRecP Then
        iRutRec = Trim(vlCodTipoIdenReceptor)
        iDgvRec = Trim(vlNumIdenReceptor)
        iNombre = (vgRs!Gls_NomBen)
        iNombreSeg = IIf(IsNull(vgRs!Gls_NomSegBen), "", (vgRs!Gls_NomSegBen))
        iPaterno = (vgRs!Gls_PatBen)
        iMaterno = IIf(IsNull(vgRs!Gls_MatBen), "", (vgRs!Gls_MatBen))
        Exit Function
    End If
    
    'Receptor de Tipo MADRE
    If iCodTipReceptor = clTipRecM Then
        vgSql = ""
        vgSql = "SELECT cod_tipoidenben,num_idenben,gls_nomben,gls_nomsegben,gls_patben,gls_matben "
        vgSql = vgSql & "FROM pp_tmae_ben "
        vgSql = vgSql & "WHERE cod_grufam = '" & Trim(vlCodGruFam) & "' AND "
        vgSql = vgSql & "cod_par IN " & clCodParMadres & " AND "
        vgSql = vgSql & "num_poliza = '" & Trim(iNumPol) & "' AND "
        vgSql = vgSql & "num_endoso = " & Trim(inumendoso) & " "
        Set vgRegistro = vgConexionBD.Execute(vgSql)
        If Not vgRegistro.EOF Then
            iRutRec = (vgRegistro!Cod_TipoIdenBen)
            iDgvRec = (vgRegistro!Num_IdenBen)
            iNombre = (vgRegistro!Gls_NomBen)
            iNombreSeg = IIf(IsNull(vgRegistro!Gls_NomSegBen), "", (vgRegistro!Gls_NomSegBen))
            iPaterno = (vgRegistro!Gls_PatBen)
            iMaterno = IIf(IsNull(vgRegistro!Gls_MatBen), "", (vgRegistro!Gls_MatBen))
        End If
        Exit Function
    End If

    'Receptor de Tipo TUTOR
    If iCodTipReceptor = clTipRecT Then
        vgSql = ""
        vgSql = "SELECT cod_tipoidentut,num_identut,gls_nomtut,gls_nomsegtut,gls_pattut,gls_mattut "
        vgSql = vgSql & "FROM pp_tmae_tutor "
        vgSql = vgSql & "WHERE cod_tipoidentut = " & Trim(vlCodTipoIdenReceptor) & " "
        vgSql = vgSql & "AND num_identut = '" & Trim(vlNumIdenReceptor) & "' "
        Set vgRegistro = vgConexionBD.Execute(vgSql)
        If Not vgRegistro.EOF Then
            iRutRec = (vgRegistro!cod_tipoidentut)
            iDgvRec = (vgRegistro!num_identut)
            iNombre = (vgRegistro!gls_nomtut)
            iNombreSeg = IIf(IsNull(vgRegistro!gls_nomsegtut), "", vgRegistro!gls_nomsegtut)
            iPaterno = (vgRegistro!gls_pattut)
            iMaterno = IIf(IsNull(vgRegistro!gls_mattut), "", vgRegistro!gls_mattut)
        End If
        Exit Function
    End If

    'Receptor de Tipo RETENEDOR
    If iCodTipReceptor = clTipRecR Then
        vgSql = ""
        vgSql = "SELECT cod_tipoidenreceptor,num_idenreceptor,gls_nomreceptor,gls_nomsegreceptor,gls_patreceptor,gls_matreceptor "
        vgSql = vgSql & "FROM pp_tmae_retjudicial "
        vgSql = vgSql & "WHERE cod_tipoidenreceptor = " & Trim(vlCodTipoIdenReceptor) & " "
        vgSql = vgSql & "AND num_idenreceptor = '" & Trim(vlNumIdenReceptor) & "' "
        Set vgRegistro = vgConexionBD.Execute(vgSql)
        If Not vgRegistro.EOF Then
            iRutRec = (vgRegistro!Cod_TipoIdenReceptor)
            iDgvRec = (vgRegistro!Num_IdenReceptor)
            iNombre = (vgRegistro!Gls_NomReceptor)
            iNombreSeg = IIf(IsNull(vgRegistro!Gls_NomSegReceptor), "", vgRegistro!Gls_NomSegReceptor)
            iPaterno = (vgRegistro!Gls_PatReceptor)
            iMaterno = IIf(IsNull(vgRegistro!Gls_MatReceptor), "", vgRegistro!Gls_MatReceptor)
        End If
        Exit Function
    End If

Exit Function
Err_flBuscarNombreReceptor:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flBuscarMontosConceptos(iTabla As String)
Dim vlTabla As String
    
    vlSql = ""
    vlSql = "SELECT p.cod_conhabdes,c.cod_imponible,c.cod_tributable,cod_tipmov,"
    If vgTipoBase = "ORACLE" Then
        vlSql = vlSql & " NVL(SUM(mto_conhabdes),0) as monto FROM "
    Else
        vlSql = vlSql & " isnull(SUM(mto_conhabdes),0) as monto FROM "
    End If
    vlSql = vlSql & "pp_tmae_pagopen" & iTabla & " p, "
    vlSql = vlSql & "ma_tpar_conhabdes c WHERE "
    vlSql = vlSql & "p.cod_conhabdes = c.cod_conhabdes AND "
    vlSql = vlSql & "p.num_perpago = '" & vlNumPerPago & "' AND "
    vlSql = vlSql & "p.num_poliza = '" & vlNumPol & "' AND "
    vlSql = vlSql & "p.num_orden = " & vlNumOrden & " AND "
'    vlSQL = vlSQL & "c.cod_conhabdes NOT IN ('01','23','24') AND "
    vlSql = vlSql & "p.cod_tipreceptor = '" & vlCodTipReceptor & "' "
    vlSql = vlSql & "GROUP BY p.cod_conhabdes,c.cod_imponible,c.cod_tributable,cod_tipmov "
    vlSql = vlSql & "ORDER BY p.cod_conhabdes,c.cod_imponible,c.cod_tributable,cod_tipmov "
    Set vgRs2 = vgConexionBD.Execute(vlSql)
    If Not vgRs2.EOF Then
        While Not vgRs2.EOF
            'Monto de Pension en Pesos
            If Trim(vgRs2!Cod_ConHabDes) = clCodHD01 Then
                vlMtoPensionPesos = vlMtoPensionPesos + (vgRs2!monto)
            Else
                If Trim(vgRs2!Cod_ConHabDes) = clCodHD02 Then
                    vlMtoPensionPesos = vlMtoPensionPesos + (vgRs2!monto)
                Else
                    If Trim(vgRs2!Cod_ConHabDes) = clCodHD23 Then
                        vlMtoImpuesto = (vgRs2!monto)
                    Else
                        If Trim(vgRs2!Cod_ConHabDes) = clCodHD24 Then
                            vlMtoSalud = (vgRs2!monto)
                        End If
                    End If
                End If
            End If
            If Trim(vgRs2!Cod_ConHabDes) <> clCodHD01 And Trim(vgRs2!Cod_ConHabDes) <> clCodHD02 And _
                Trim(vgRs2!Cod_ConHabDes) <> clCodHD23 And Trim(vgRs2!Cod_ConHabDes) <> clCodHD24 Then
                'Monto de Haberes
                If (vgRs2!cod_tipmov) = clCodH Then
                    'Monto Haber Imponible
                    If ((vgRs2!cod_imponible) = clCodImpS) And ((vgRs2!cod_tributable) = clCodImpS) Then
                        vlMtoHabImp = vlMtoHabImp + (vgRs2!monto)
                    Else
                        'Monto Haber No Imponible
                        If ((vgRs2!cod_imponible) = clCodImpN) And ((vgRs2!cod_tributable) = clCodImpN) Then
                            vlMtoHabNoImp = vlMtoHabNoImp + (vgRs2!monto)
                        End If
                    End If
                Else
                    'Montos de Descuentos
                    If (vgRs2!cod_tipmov) = clCodD Then
                        'Monto Descuento Imponible
                        If ((vgRs2!cod_imponible) = clCodImpS) And ((vgRs2!cod_tributable) = clCodImpS) Then
                            vlMtoDesImp = vlMtoDesImp + (vgRs2!monto)
                        Else
                            'Monto Descuento No Imponible
                            If ((vgRs2!cod_imponible) = clCodImpN) And ((vgRs2!cod_tributable) = clCodImpN) Then
                                vlMtoDesNoImp = vlMtoDesNoImp + (vgRs2!monto)
                            End If
                        End If
                    End If
                End If
            End If
            vgRs2.MoveNext
        Wend
    End If
    vgRs2.Close

End Function

Function flProcesoViaPago()
On Error GoTo Err_flProcesoViaPago

   Screen.MousePointer = 11

    If vgNombreInformeSeleccionado = "InfViaPago" Then
        vlArchivo = strRpt & "PP_Rpt_NominaViaPago" & vlGlosaOpcion & ".rpt"   '\Reportes
        If Not fgExiste(vlArchivo) Then     ', vbNormal
           MsgBox "Archivo de Reporte de Vías de Pago de Pensiones por Pensionado no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
           Screen.MousePointer = 0
           Exit Function
        End If
   Else
       If vgNombreInformeSeleccionado = "InfViaPagoRec" Then
            vlArchivo = strRpt & "PP_Rpt_NominaViaPagoReceptor" & vlGlosaOpcion & ".rpt"   '\Reportes
            If Not fgExiste(vlArchivo) Then     ', vbNormal
               MsgBox "Archivo de Reporte de Vías de Pago de Pensiones por Receptor no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
               Screen.MousePointer = 0
               Exit Function
            End If
       End If
    End If
    
    Call fgVigenciaQuiebra(Txt_Desde)
   
   vlFechaInicio = Format(CDate(Trim(Txt_Desde.Text)), "yyyymmdd")
   vlFechaTermino = Format(CDate(Trim(Txt_Hasta.Text)), "yyyymmdd")
   
'   vgSql = ""
'   vgSql = "SELECT mto_elemento "
'   vgSql = vgSql & "FROM MA_TPAR_TABCODVIG "
'   vgSql = vgSql & "WHERE "
'   vgSql = vgSql & "cod_tabla = '" & clCodPrcSalud & "' AND "
'   vgSql = vgSql & "cod_elemento = '" & clCodPSM & "' AND "
'   vgSql = vgSql & "fec_tervig = " & clFechaTopeTer & " "
'   vgSql = vgSql & "ORDER BY fec_inivig DESC "
'   Set vgRegistro = vgConexionBD.Execute(vgSql)
'   If Not vgRegistro.EOF Then
'      vlPrcSalud = (vgRegistro!Mto_Elemento)
'   End If
   
   vgQuery = ""
   vgQuery = vgQuery & "{PP_TMAE_LIQPAGOPEN" & vlGlosaOpcion & ".cod_tipopago} = '" & Trim(vlPago) & "' AND "
   vgQuery = vgQuery & "{PP_TMAE_LIQPAGOPEN" & vlGlosaOpcion & ".fec_pago} >= '" & Trim(vlFechaInicio) & "' AND "
   vgQuery = vgQuery & "{PP_TMAE_LIQPAGOPEN" & vlGlosaOpcion & ".fec_pago} <= '" & Trim(vlFechaTermino) & "'"
   'vgQuery = vgQuery & " AND {MA_TVAL_MONEDA.cod_moneda} = '" & Trim(cgCodTipMonedaUF) & "' "
'   vgQuery = vgQuery & "{PP_TMAE_PAGOPEN" & vlGlosaOpcion & ".cod_conhabdes} = '" & Trim(clCodConHabDes24) & "' AND "
'   vgQuery = vgQuery & "{MA_TPAR_TABCOD.cod_tabla} = '" & Trim(clCodTabCodIS) & "' AND "
   
   
        
   Rpt_Reporte.Reset
   Rpt_Reporte.WindowState = crptMaximized
   Rpt_Reporte.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
   Rpt_Reporte.Connect = vgRutaDataBase       ' o App.Path & "\Nestle.mdb"
   Rpt_Reporte.SelectionFormula = vgQuery

'   vgPalabra = ""
'   vgPalabra = Txt_Desde.Text & "   -   " & Txt_Hasta.Text
  

   Rpt_Reporte.Formulas(0) = ""
   Rpt_Reporte.Formulas(1) = ""
   Rpt_Reporte.Formulas(2) = ""
   Rpt_Reporte.Formulas(3) = ""
   Rpt_Reporte.Formulas(4) = ""
   Rpt_Reporte.Formulas(5) = ""

   Rpt_Reporte.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
   Rpt_Reporte.Formulas(1) = "NombreSistema= '" & vgNombreSistema & "'"
   Rpt_Reporte.Formulas(2) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
'   Rpt_Reporte.Formulas(3) = "Porcentaje = " & vlPrcSalud & ""
    If Trim(vlGlosaOpcion) = "DEF" Then
        Rpt_Reporte.Formulas(4) = "TipoProceso= 'DEFINITIVO' "
    Else
        Rpt_Reporte.Formulas(4) = "TipoProceso= 'PROVISORIO' "
    End If
    
    Rpt_Reporte.Formulas(5) = "GlosaQuiebra = '" & vgGlosaQuiebra & "'"

   Rpt_Reporte.SubreportToChange = ""
   Rpt_Reporte.Destination = crptToWindow
   Rpt_Reporte.WindowState = crptMaximized
   Rpt_Reporte.WindowTitle = "Informe de Vías de Pago de Pensiones"
   'Rpt_Reporte.SelectionFormula = ""
   Rpt_Reporte.Action = 1
   Screen.MousePointer = 0

Exit Function
Err_flProcesoViaPago:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flProcesoLiqconError()
On Error GoTo Err_flProcesoLiqconError

   Screen.MousePointer = 11

   vlArchivo = strRpt & "PP_Rpt_LiqconError" & vlGlosaOpcion & ".rpt"   '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
      MsgBox "Archivo de Reporte de Pensiondes con Error no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
      Screen.MousePointer = 0
      Exit Function
   End If
   
   Call fgVigenciaQuiebra(Txt_Desde)
   
   vgQuery = ""
   vgQuery = vgQuery & "{PP_TMAE_LIQPAGOPEN" & vlGlosaOpcion & ".cod_tipopago} = '" & Trim(vlPago) & "' AND " 'hqr 11/06/2005
   vgQuery = vgQuery & "{PP_TMAE_LIQPAGOPEN" & vlGlosaOpcion & ".FEC_PAGO} >= '" & Trim(vlFechaInicio) & "' AND "
   vgQuery = vgQuery & "{PP_TMAE_LIQPAGOPEN" & vlGlosaOpcion & ".FEC_PAGO} <= '" & Trim(vlFechaTermino) & "' AND "
   vgQuery = vgQuery & "{PP_TMAE_LIQPAGOPEN" & vlGlosaOpcion & ".MTO_LIQPAGAR} <= 0"
   'vgQuery = vgQuery & " AND {MA_TVAL_MONEDA.cod_moneda} = '" & Trim(cgCodTipMonedaUF) & "' "

        
   Rpt_Reporte.Reset
   Rpt_Reporte.WindowState = crptMaximized
   Rpt_Reporte.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
   Rpt_Reporte.Connect = vgRutaDataBase       ' o App.Path & "\Nestle.mdb"
   Rpt_Reporte.SelectionFormula = vgQuery

   Rpt_Reporte.Formulas(0) = ""
   Rpt_Reporte.Formulas(1) = ""
   Rpt_Reporte.Formulas(2) = ""
   Rpt_Reporte.Formulas(3) = ""
   Rpt_Reporte.Formulas(4) = ""

   Rpt_Reporte.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
   Rpt_Reporte.Formulas(1) = "NombreSistema= '" & vgNombreSistema & "'"
   Rpt_Reporte.Formulas(2) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
   If Trim(vlGlosaOpcion) = "DEF" Then
       Rpt_Reporte.Formulas(3) = "TipoProceso= 'DEFINITIVO' "
   Else
       Rpt_Reporte.Formulas(3) = "TipoProceso= 'PROVISORIO' "
   End If
   
   Rpt_Reporte.Formulas(4) = "GlosaQuiebra = '" & vgGlosaQuiebra & "'"
  
   Rpt_Reporte.SubreportToChange = ""
   Rpt_Reporte.Destination = crptToWindow
   Rpt_Reporte.WindowState = crptMaximized
   Rpt_Reporte.WindowTitle = "Informe de Pensiones con Error"
   Rpt_Reporte.Action = 1
   Screen.MousePointer = 0

Exit Function
Err_flProcesoLiqconError:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flProcesoPlanillaCotSalud()
On Error GoTo Err_flProcesoPlanillaCotSalud

   Screen.MousePointer = 11

   vlArchivo = strRpt & "PP_Rpt_PlanillaCotSalud" & vlGlosaOpcion & ".rpt"   '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
      MsgBox "Archivo de Reporte de Resumen de Planilla de Pagos de Cotizaciones no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
      Screen.MousePointer = 0
      Exit Function
   End If
   
   Call fgVigenciaQuiebra(Txt_Desde)
   
   vlFechaInicio = Format(CDate(Trim(Txt_Desde.Text)), "yyyymmdd")
   vlFechaTermino = Format(CDate(Trim(Txt_Hasta.Text)), "yyyymmdd")
   
'   vgSql = "SELECT mto_elemento "
'   vgSql = vgSql & "FROM MA_TPAR_TABCODVIG "
'   vgSql = vgSql & "WHERE "
'   vgSql = vgSql & "cod_tabla = '" & clCodPrcSalud & "' AND "
'   vgSql = vgSql & "cod_elemento = '" & clCodPSM & "' AND "
'   vgSql = vgSql & "fec_tervig = " & clFechaTopeTer & " "
'   vgSql = vgSql & "ORDER BY fec_inivig DESC "
'   Set vgRegistro = vgConexionBD.Execute(vgSql)
'   If Not vgRegistro.EOF Then
'      vlPrcSalud = (vgRegistro!Mto_Elemento)
'   End If
'   vgSql = ""
   
   vgQuery = ""
   vgQuery = vgQuery & "{PP_TMAE_LIQPAGOPEN" & vlGlosaOpcion & ".cod_tipopago} = '" & Trim(vlPago) & "' "
   vgQuery = vgQuery & "AND {PP_TMAE_LIQPAGOPEN" & vlGlosaOpcion & ".fec_pago} >= '" & Trim(vlFechaInicio) & "' "
   vgQuery = vgQuery & "AND {PP_TMAE_LIQPAGOPEN" & vlGlosaOpcion & ".fec_pago} <= '" & Trim(vlFechaTermino) & "' "
   vgQuery = vgQuery & "AND {PP_TMAE_LIQPAGOPEN" & vlGlosaOpcion & ".cod_inssalud} <> '" & Trim(clCodISExento) & "' "
   vgQuery = vgQuery & "AND {PP_TMAE_PAGOPEN" & vlGlosaOpcion & ".cod_conhabdes} = '" & Trim(clCodConHabDes24) & "' "
'   vgQuery = vgQuery & "AND {PP_TMAE_PAGOPEN" & vlGlosaOpcion & "34" & ".cod_conhabdes} = '" & Trim(clCodConHabDes34) & "' "
   vgQuery = vgQuery & "AND {MA_TPAR_TABCOD.cod_tabla} = '" & Trim(clCodTabCodIS) & "' "
'   vgQuery = vgQuery & "AND {MA_TVAL_MONEDA.cod_moneda} = '" & Trim(cgCodTipMonedaUF) & "' "
   vgQuery = vgQuery & "AND {MA_TPAR_TABCODMON.cod_tabla} = '" & Trim(vgCodTabla_TipMon) & "' "
   vgQuery = vgQuery & "AND {MA_TPAR_TABCODVIG.cod_tabla} = '" & Trim(clCodPrcSalud) & "' "
   vgQuery = vgQuery & "AND {MA_TPAR_TABCODVIG.fec_inivig} <= {PP_TMAE_LIQPAGOPEN" & vlGlosaOpcion & ".fec_pago} "
   vgQuery = vgQuery & "AND {MA_TPAR_TABCODVIG.fec_tervig} >= {PP_TMAE_LIQPAGOPEN" & vlGlosaOpcion & ".fec_pago} "

   Rpt_Reporte.Reset
   Rpt_Reporte.WindowState = crptMaximized
   Rpt_Reporte.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
   Rpt_Reporte.Connect = vgRutaDataBase       ' o App.Path & "\Nestle.mdb"
   Rpt_Reporte.SelectionFormula = vgQuery

'   vgPalabra = ""
'   vgPalabra = Txt_Desde.Text & "   -   " & Txt_Hasta.Text
  

   Rpt_Reporte.Formulas(0) = ""
   Rpt_Reporte.Formulas(1) = ""
   Rpt_Reporte.Formulas(2) = ""
   Rpt_Reporte.Formulas(3) = ""
   Rpt_Reporte.Formulas(4) = ""
   Rpt_Reporte.Formulas(5) = ""
   
   Rpt_Reporte.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
   Rpt_Reporte.Formulas(1) = "NombreSistema= '" & vgNombreSistema & "'"
   Rpt_Reporte.Formulas(2) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
'   Rpt_Reporte.Formulas(3) = "Porcentaje = " & vlPrcSalud & ""
   If Trim(vlGlosaOpcion) = "DEF" Then
       Rpt_Reporte.Formulas(4) = "TipoProceso= 'DEFINITIVO' "
   Else
       Rpt_Reporte.Formulas(4) = "TipoProceso= 'PROVISORIO' "
   End If
   
   Rpt_Reporte.Formulas(5) = "GlosaQuiebra = '" & vgGlosaQuiebra & "'"

   Rpt_Reporte.SubreportToChange = ""
   Rpt_Reporte.Destination = crptToWindow
   Rpt_Reporte.WindowState = crptMaximized
   Rpt_Reporte.WindowTitle = "Planilla de Cotizaciones de Salud"
   'Rpt_Reporte.SelectionFormula = ""
   Rpt_Reporte.Action = 1
   Screen.MousePointer = 0

Exit Function
Err_flProcesoPlanillaCotSalud:
    Screen.MousePointer = 0
    
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flProcesoPagoSaludPenRV()
On Error GoTo Err_flProcesoPagoSaludPenRV

   Screen.MousePointer = 11

   vlArchivo = strRpt & "PP_Rpt_PlanillaPagoSalud" & vlGlosaOpcion & ".rpt"   '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
      MsgBox "Archivo de Reporte de Planilla de Pago Salud Pensiones de Renta Vitalicia no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
      Screen.MousePointer = 0
      Exit Function
   End If
   
   Call fgVigenciaQuiebra(Txt_Desde)
   
   vlFechaInicio = Format(CDate(Trim(Txt_Desde.Text)), "yyyymmdd")
   vlFechaTermino = Format(CDate(Trim(Txt_Hasta.Text)), "yyyymmdd")

   vgQuery = ""
   vgQuery = vgQuery & "{PP_TMAE_LIQPAGOPEN" & vlGlosaOpcion & ".cod_tipopago} = '" & Trim(vlPago) & "' AND "
   vgQuery = vgQuery & "{PP_TMAE_LIQPAGOPEN" & vlGlosaOpcion & ".fec_pago} >= '" & Trim(vlFechaInicio) & "' AND "
   vgQuery = vgQuery & "{PP_TMAE_LIQPAGOPEN" & vlGlosaOpcion & ".fec_pago} <= '" & Trim(vlFechaTermino) & "' AND "
   vgQuery = vgQuery & "{PP_TMAE_LIQPAGOPEN" & vlGlosaOpcion & ".cod_inssalud} <> '" & Trim(clCodISExento) & "' AND "
   vgQuery = vgQuery & "{PP_TMAE_PAGOPEN" & vlGlosaOpcion & ".cod_conhabdes} = '" & Trim(clCodConHabDes24) & "' AND "
   vgQuery = vgQuery & "{MA_TPAR_TABCOD.cod_tabla} = '" & Trim(clCodTabCodIS) & "' AND "
   vgQuery = vgQuery & "{MA_TPAR_TABCODMON.cod_tabla} = '" & Trim(vgCodTabla_TipMon) & "' "
          
   Rpt_Reporte.Reset
   Rpt_Reporte.WindowState = crptMaximized
   Rpt_Reporte.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
   Rpt_Reporte.Connect = vgRutaDataBase       ' o App.Path & "\Nestle.mdb"
   Rpt_Reporte.SelectionFormula = vgQuery

'   vgPalabra = ""
'   vgPalabra = Txt_Desde.Text & "   -   " & Txt_Hasta.Text

   Rpt_Reporte.Formulas(0) = ""
   Rpt_Reporte.Formulas(1) = ""
   Rpt_Reporte.Formulas(2) = ""
   Rpt_Reporte.Formulas(3) = ""
   Rpt_Reporte.Formulas(4) = ""

   Rpt_Reporte.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
   Rpt_Reporte.Formulas(1) = "NombreSistema= '" & vgNombreSistema & "'"
   Rpt_Reporte.Formulas(2) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
'   Rpt_Reporte.Formulas(3) = "Periodo = '" & vgPalabra & "'"
   If Trim(vlGlosaOpcion) = "DEF" Then
       Rpt_Reporte.Formulas(3) = "TipoProceso= 'DEFINITIVO' "
   Else
       Rpt_Reporte.Formulas(3) = "TipoProceso= 'PROVISORIO' "
   End If
   
   Rpt_Reporte.Formulas(4) = "GlosaQuiebra = '" & vgGlosaQuiebra & "'"

   Rpt_Reporte.SubreportToChange = ""
   Rpt_Reporte.Destination = crptToWindow
   Rpt_Reporte.WindowState = crptMaximized
   Rpt_Reporte.WindowTitle = "Planilla de Pago Salud Pensiones Renta Vitalicia"
   'Rpt_Reporte.SelectionFormula = ""
   Rpt_Reporte.Action = 1
   Screen.MousePointer = 0

Exit Function
Err_flProcesoPagoSaludPenRV:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Private Sub Cmb_Pago_Click()
    If Cmb_Pago.ListIndex <> -1 Then
        If vgNombreInformeSeleccionado = "InfLibPenInt" Then
            Cmd_Imprimir_AFP.Enabled = True
        Else
            Cmd_Imprimir_AFP.Enabled = False
        End If
    End If
End Sub

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

Private Sub Cmd_Imprimir_AFP_Click()

'On Error GoTo errImprimir

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
    
    If (CDate(Trim(Txt_Desde)) > CDate(Trim(Txt_Hasta))) Then
        MsgBox "La Fecha de Término de Periodo es Mayor a la Fecha de Inicio", vbCritical, "Error de Datos"
        Exit Sub
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
    
    vlFechaInicio = Format(CDate(Trim(Txt_Desde)), "yyyymmdd")
    
    If vlGlosaOpcion = "DEF" Then
        vlCodEstado = "C"
    Else
        vlCodEstado = "P"
    End If
    
    If vgNombreInformeSeleccionado <> "InfVenTut" Then
        If vlPago = "R" Then 'Solo para los pagos en regimen
        If flValidaEstadoProceso(Mid(Trim(vlFechaInicio), 1, 6), vlCodEstado) = False Then
            MsgBox "El Tipo de Proceso Seleccionado no se encuentra Realizado.", vbCritical, "Error de Datos"
            Screen.MousePointer = 0
            Exit Sub
        End If
        End If
    End If
  
    Screen.MousePointer = 11
    
    vlFechaInicio = Format(CDate(Trim(Txt_Desde)), "yyyymmdd")
    vlFechaTermino = Format(CDate(Trim(Txt_Hasta)), "yyyymmdd")

    'Permite imprimir la Opción Indicada a través del Menú
    Select Case vgNombreInformeSeleccionado
        Case "InfLibPenInt"    'Informe de Libro de Pensiones Interno
            vgFormatoAFP = "AFP"
            Call flInformePensionesInt
    End Select
    Screen.MousePointer = 0

Exit Sub
errImprimir:
    Screen.MousePointer = 0
    If Err.Number <> 0 Then
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End If

End Sub

Private Sub Cmd_Imprimir_Click()
On Error GoTo errImprimir

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
    
    If (CDate(Trim(Txt_Desde)) > CDate(Trim(Txt_Hasta))) Then
        MsgBox "La Fecha de Término de Periodo es Mayor a la Fecha de Inicio", vbCritical, "Error de Datos"
        Exit Sub
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
    
    'CMV-20060222 I
    
    vlFechaInicio = Format(CDate(Trim(Txt_Desde)), "yyyymmdd")
    
    If vlGlosaOpcion = "DEF" Then
        vlCodEstado = "C"
    Else
        vlCodEstado = "P"
    End If
    'HQR 15/10/2007 Se quita validacion (¿para qué se usa?)
    If vgNombreInformeSeleccionado <> "InfVenTut" Then
        If vlPago = "R" Then 'Solo para los pagos en regimen
        If flValidaEstadoProceso(Mid(Trim(vlFechaInicio), 1, 6), vlCodEstado) = False Then
            MsgBox "El Tipo de Proceso Seleccionado no se encuentra Realizado.", vbCritical, "Error de Datos"
            Screen.MousePointer = 0
            Exit Sub
        End If
        End If
    End If
    'Fin hqr 15/10/2007
    'CMV-20060222 F
  
    Screen.MousePointer = 11
    
'    vgFechaIni = CDate(Txt_Desde)
'    vgFechaTer = CDate(Txt_Hasta)
'

    vlFechaInicio = Format(CDate(Trim(Txt_Desde)), "yyyymmdd")
    vlFechaTermino = Format(CDate(Trim(Txt_Hasta)), "yyyymmdd")

    'Permite imprimir la Opción Indicada a través del Menú
    Select Case vgNombreInformeSeleccionado
        Case "InfLibPen"    'Informe de Libro de Pensiones
             Call flInformePensiones
        Case "InfLibPenInt"    'Informe de Libro de Pensiones Interno
            vgFormatoAFP = ""  '200912 DCM
            Call flInformePensionesInt
'        Case "InfImpUni"    'Informe de Impuesto Único
'            Call flInformeImpUnico
        Case "InfVenTut"    'Informe de Vencimiento de Tutores
            Call flInformeVenTut
        Case "InfDefRetJud"    'Informe Definitivo de Retenciones Judiciales
            Call flInformeDefRetJud
'        Case "InfPagoMensual" 'Informe de Nómina de Pago Mensual Asignaciones Familiares
'            Call flProcesoAsigFam
'        Case "InfLibroPenMin"     'Informe de Libro de Pensiones Minimas
'            Call flProcesoGE
'        Case "InfPenMinSinDer"  'Informe de Pensiones Mínimas sin Derecho
'            Call flPenMinSinDer
'        Case "InfHabDesGE"    'Informe de Haberes y Descuentos de Garantía Estatal
'            Call flProcesoHabDesGE
'        Case "InfPreMedFon"    'Planilla de Pago de Prestamos Medicos Fonasa
'            Call flProcesoPreMedFonasa
        Case "InfPagSalPenRV"    'Planilla de Pago Salud Pensiones Renta Vitalicia (Resumen de Planilla de Pagos de Cotizaciones)
            Call flProcesoPagoSaludPenRV
        Case "InfPlaCotSalud"    'Planilla de Cotizaciones de Salud
            Call flProcesoPlanillaCotSalud
        Case "InfViaPago"    'Informe de Vias de Pago de Pensiones por Pensionado
            Call flProcesoViaPago
        Case "InfViaPagoRec"    'Informe de Vias de Pago de Pensiones por Receptor
            Call flProcesoViaPago
'        Case "infPagoCCaf" ' Informe de Pago CCAF Aporte,Creditos,Otras Prestaciones
'             Call flInfPagoCCaf
''        Case "InfPagoExceso" 'Informe de Pagos en Exceso detectados por la Compañia
''             Call flInfPagoExceso
''        Case "InfPagoMenos" 'Informe de Solicitud de Liquidacion por pagos efectuados de menos
''             Call flInfPagoMenos
'        Case "InfMontosPenCancelados" 'Informe de Montos de Pensiones Cancelados
'             Call flMontosPenCancelados
''        Case "InfAnexo2" 'Informe de Anexo 2
''             Call flInfAnexo2
''        Case "InfConciliacion" 'Informe de Conciliación de Garantía Estatal
''             Call flInfConciliacion
'        Case "InfCenContable"    'Informe de Centralización Contable
'            Call flProcesoCenContable
        Case "InfLiqconError"    'Informe de Pensiones con Error
            Call flProcesoLiqconError
'        Case "InfLibPenCont"    'Informe de Libro de Pensiones Contable (Ordenado por Tipo de Pensión)
'               Call flInformePensionesCont
             
    End Select
    Screen.MousePointer = 0

Exit Sub
errImprimir:
    Screen.MousePointer = 0
    If Err.Number <> 0 Then
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]" & glNumpol, vbCritical, "¡ERROR!..."
    End If
End Sub

Private Sub Cmd_Imprimir_Directo_Click()

'On Error GoTo errImprimir

    

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
    
    If (CDate(Trim(Txt_Desde)) > CDate(Trim(Txt_Hasta))) Then
        MsgBox "La Fecha de Término de Periodo es Mayor a la Fecha de Inicio", vbCritical, "Error de Datos"
        Exit Sub
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
    
    vlFechaInicio = Format(CDate(Trim(Txt_Desde)), "yyyymmdd")
    
    If vlGlosaOpcion = "DEF" Then
        vlCodEstado = "C"
    Else
        vlCodEstado = "P"
    End If
    
    If vgNombreInformeSeleccionado <> "InfVenTut" Then
        If vlPago = "R" Then 'Solo para los pagos en regimen
        If flValidaEstadoProceso(Mid(Trim(vlFechaInicio), 1, 6), vlCodEstado) = False Then
            MsgBox "El Tipo de Proceso Seleccionado no se encuentra Realizado.", vbCritical, "Error de Datos"
            Screen.MousePointer = 0
            Exit Sub
        End If
        End If
    End If
  
    Screen.MousePointer = 11
    
    vlFechaInicio = Format(CDate(Trim(Txt_Desde)), "yyyymmdd")
    vlFechaTermino = Format(CDate(Trim(Txt_Hasta)), "yyyymmdd")

    'Permite imprimir la Opción Indicada a través del Menú
    Select Case vgNombreInformeSeleccionado
        Case "InfLibPenInt"    'Informe de Libro de Pensiones Interno
            vgFormatoAFP = "DIR"
            lbl_Indicador.Caption = "Cargando Data de Liquidación..."
            Call flInformePenDirecto
    End Select
    Screen.MousePointer = 0

Exit Sub
errImprimir:
    Screen.MousePointer = 0
    If Err.Number <> 0 Then
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End If
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
    
    Frm_PlanillaPago.Left = 0
    Frm_PlanillaPago.Top = 0
    fgComboTipoCalculo Cmb_Tipo
    fgComboTipoPension Cmb_Pago
    
    Select Case vgNombreInformeSeleccionado
        Case "InfLibPen"    'Informe de Libro de Pensiones
            Cmd_Imprimir.Visible = True
        Case "InfLibPenInt"    'Informe de Libro de Pensiones Interno
            Cmd_Imprimir_Directo.Visible = True
        Case "InfVenTut"    'Informe de Vencimiento de Tutores
            Cmd_Imprimir.Visible = True
        Case "InfDefRetJud"    'Informe Definitivo de Retenciones Judiciales
            Cmd_Imprimir.Visible = True
'            Call flProcesoPreMedFonasa
        Case "InfPagSalPenRV"    'Planilla de Pago Salud Pensiones Renta Vitalicia (Resumen de Planilla de Pagos de Cotizaciones)
            Cmd_Imprimir.Visible = True
        Case "InfPlaCotSalud"    'Planilla de Cotizaciones de Salud
            Cmd_Imprimir.Visible = True
        Case "InfViaPago"    'Informe de Vias de Pago de Pensiones por Pensionado
            Cmd_Imprimir.Visible = True
        Case "InfViaPagoRec"    'Informe de Vias de Pago de Pensiones por Receptor
            Cmd_Imprimir.Visible = True
        Case "InfLiqconError"    'Informe de Pensiones con Error
            Cmd_Imprimir.Visible = True
        Case "InfLibPenCont"    'Informe de Libro de Pensiones Contable (Ordenado por Tipo de Pensión)
            Cmd_Imprimir_Directo.Visible = True
    End Select
    
    
    
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
                Txt_Desde = ""
                Exit Sub
            End If
            If Txt_Hasta <> "" Then
                If (CDate(Trim(Txt_Desde)) > CDate(Trim(Txt_Hasta))) Then
                    Txt_Desde = ""
                    Exit Sub
                End If
            End If
            If (Year(CDate(Trim(Txt_Desde))) < 1900) Then
                Txt_Desde = ""
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
        Txt_Desde = ""
        Exit Sub
    End If
    If Txt_Hasta <> "" Then
        If (CDate(Trim(Txt_Desde)) > CDate(Trim(Txt_Hasta))) Then
            Txt_Desde = ""
            Exit Sub
        End If
    End If
    If (Year(CDate(Trim(Txt_Desde))) < 1900) Then
        Txt_Desde = ""
        Exit Sub
    End If
    Txt_Desde.Text = Format(CDate(Trim(Txt_Desde)), "yyyymmdd")
    vlFechaDesde = Trim(Txt_Desde)
    Txt_Desde.Text = DateSerial(Mid((Txt_Desde.Text), 1, 4), Mid((Txt_Desde.Text), 5, 2), Mid((Txt_Desde.Text), 7, 2))
End If
End Sub

Private Sub Txt_Hasta_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    If (Txt_Hasta <> "") Then
        If Not IsDate(Trim(Txt_Hasta)) Then
           Txt_Hasta = ""
            Exit Sub
        End If
        If Txt_Desde <> "" Then
            If (CDate(Trim(Txt_Desde)) > CDate(Trim(Txt_Hasta))) Then
                Txt_Hasta = ""
                Exit Sub
            End If
        End If
        If (Year(CDate(Trim(Txt_Hasta))) < 1900) Then
            Txt_Hasta = ""
            Exit Sub
        End If
        Txt_Hasta.Text = Format(CDate(Trim(Txt_Hasta)), "yyyymmdd")
        vlFechaHasta = Trim(Txt_Hasta)
        Txt_Hasta.Text = DateSerial(Mid((Txt_Hasta.Text), 1, 4), Mid((Txt_Hasta.Text), 5, 2), Mid((Txt_Hasta.Text), 7, 2))
    End If
    Cmd_Imprimir.SetFocus
End If
End Sub

Private Sub Txt_Hasta_LostFocus()
If (Txt_Hasta <> "") Then
    If Not IsDate(Trim(Txt_Hasta)) Then
        Txt_Hasta = ""
        Exit Sub
    End If
    If Txt_Desde <> "" Then
        If (CDate(Trim(Txt_Desde)) > CDate(Trim(Txt_Hasta))) Then
            Txt_Hasta = ""
            Exit Sub
        End If
    End If
    If (Year(CDate(Trim(Txt_Hasta))) < 1900) Then
        Txt_Hasta = ""
        Exit Sub
    End If
    Txt_Hasta.Text = Format(CDate(Trim(Txt_Hasta)), "yyyymmdd")
    vlFechaHasta = Trim(Txt_Hasta)
    Txt_Hasta.Text = DateSerial(Mid((Txt_Hasta.Text), 1, 4), Mid((Txt_Hasta.Text), 5, 2), Mid((Txt_Hasta.Text), 7, 2))
End If
End Sub

Private Function flObtieneFactorAjuste(iTipoPago As String) As Double
'Función: Permite calcular el Factor de Ajuste de las Pólizas utilizando
'el monto de pensión actual y del periodo anterior
'Parámetros de Entrada:
'- iTipoPago  => Tipo de Pago (P-Primeros Pagos, R-Pagos Recurrentes)
'------------------------------------------------------------------------
'Fecha Creación     : 27/06/2008
'Fecha Modificación :
'------------------------------------------------------------------------
Dim vlMtoPenActual As Double, vlMtoPenAnterior As Double
On Error GoTo Err_FactorAjuste

    flObtieneFactorAjuste = 1
    
    'Obtiene los montos de pensión de cada póliza para los distintos periodos
    vlSql = "Select mto_pension From pp_tmae_pensionact "
    vlSql = vlSql & "Where cod_tipopago = '" & iTipoPago & "' "
    vlSql = vlSql & "and num_poliza = '" & vlNumPol & "' "
    vlSql = vlSql & "and fec_desde <= '" & vlFecPago & "' "
    vlSql = vlSql & "ORDER BY fec_desde desc "
    Set vgRs2 = vgConexionBD.Execute(vlSql)
    If Not vgRs2.EOF Then
        'Guarda el mto de pensión del mes actual
        vlMtoPenActual = (vgRs2!Mto_Pension)
        vgRs2.MoveNext
        Do While Not vgRs2.EOF
            'Guarda el mto de pensión del periodo encontrado, anterior al actual
            vlMtoPenAnterior = (vgRs2!Mto_Pension)
            Exit Do
        Loop
        'Calcula el Factor de Ajuste si la división no sera por 0
        If (vlMtoPenAnterior <> 0) Then flObtieneFactorAjuste = Format(vlMtoPenActual / vlMtoPenAnterior, "#0.000000")
    Else
        'En el caso de que no exista ningún periodo el Factor de Ajuste sera 1, al igual si existe solo el periodo actual y no uno anterior.
        flObtieneFactorAjuste = 1
    End If

Exit Function
Err_FactorAjuste:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "Crystl32.OCX"
Begin VB.Form Frm_InformeControl 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5700
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   5415
      Begin VB.TextBox Txt_Mes 
         Height          =   285
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   1
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Txt_Anno 
         Height          =   285
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   " -"
         Height          =   255
         Left            =   2280
         TabIndex        =   11
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de Pago "
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   480
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "( Mes - Año )"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3600
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Fra_Operacion 
      Height          =   1095
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   5415
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   3480
         Picture         =   "Frm_InformeControl.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   2280
         Picture         =   "Frm_InformeControl.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Imprimir Reporte"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton cmd_calcular 
         Caption         =   "&Calcular"
         Height          =   675
         Left            =   1080
         Picture         =   "Frm_InformeControl.frx":07B4
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Generar Tabla de Control"
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
      TabIndex        =   0
      Top             =   960
      Width           =   5415
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
   End
End
Attribute VB_Name = "Frm_InformeControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim stConLiquidacion      As TyLiquidacion
Dim stConDetPension       As TyDetPension
Dim stConTutor            As TyTutor
Dim stConBeneficiarios()    As TyBeneficiariosPension
Dim stDatosInformeControl As TyDatosInformeControl
Dim vlInfSeleccionado     As String

Dim vlRegistro   As ADODB.Recordset
Dim vlRegistro1  As ADODB.Recordset
Dim vlRegistro2  As ADODB.Recordset
Dim vlRegistro3  As ADODB.Recordset
Dim vlRegistro4  As ADODB.Recordset
Dim vlRegistro5  As ADODB.Recordset
Dim vlRegistro6  As ADODB.Recordset

'Variables Asignación Familiar

Dim vlNumPoliza  As String, vlNumOrden  As Integer, vlNumOrdenCar As Integer
Dim vlRutBen     As Long, vlDigito As String, vlMontoCarga As Double
Dim vlCodPar     As String, vlCodSitInv As String, vlMtoRetro As Double
Dim vlMtoReintegro   As Double, vlFecVencimiento As String, vlMto_total As Double
Dim vlDescripcion  As String, vlDgvReceptor  As String, vlRutReceptor As String
Dim vlMonto As Double
Dim vlNumReliq As Integer

Const clCodHabDesAF As String = "('08','09','30')"
Const clCodAF As String * 2 = "AF"
Const clCodCCAF As String * 4 = "CCAF"

'Variables Caja de Compensación

Dim vlCCAF As String, vlNumEndoso As String, vlRut As String, vlDgv As String
Dim vlMontoAporte As String, vlMontoCredito As String, vlMontoOtros As String
Dim vlNumPerPago As String, vlFecha As String

'Variables Retencion Judicial

Dim vlFechaInicio As String, vlFechaTermino As String, vlTipoRetJud As String
Dim vlNumRetencion As Integer, vlNumCargas As Integer, vlMtoCargas As Double
Dim vlVarCarga As Double
Dim vlMtoHaber As Double, vlMtoDescto As Double, vlMtoHabDes As Double
Dim vlMontoRetencion As Double
Dim vlPeriodo As String, vlCodHabDes As String
Dim vlCodOtrosHabDes As String, vlTipoMov As String, vlArchivo As String
Dim vlCodTipReceptor As String

Dim vlPerPagoRetencion As String 'Periodo de Pago de la Retención Judicial
Dim vlPerPagoAsigFam   As String 'Periodo de Pago de la Asignación Familiar
Dim vlPerPagoGEHabDes  As String 'Periodo de Pago de GE Haberes y Desctos.
Dim vlPerPagoGEPenMin  As String 'Periodo de Pago de GE Pensión Mínima

Const vlCodHabDesRJC As String = "('60','61','62')"
Const clTipoRJ As String * 2 = "RJ"
Const clTipoRAF As String * 3 = "RAF"
Const clCodHabDes61 As String * 2 = "61"
Const clCodHabDes62 As String * 2 = "62"
Const clCodRJ As String * 3 = "RJ"
Const clCodRJC As String * 3 = "RJC"
Const clCodHab As String * 1 = "H"
Const clCodDescto As String * 1 = "D"

'Variables de Garantia Estatal

Dim vlFecPeriodo As String
Dim vlMtoPensionPrc As Double
Dim vlMtoPensionPesos As Double
Dim vlMtoGarEst As Double
Dim vlMtoOtroHab As Double
Dim vlMtoOtroDes As Double
Dim vlBono As Double
Dim vlCodOtrosHabDesGE As String
Dim vlCodigo As String

Dim vlCodUsuario As String
'Dim vlNumPerPago As String
'Dim vlNumPoliza As String
'Dim vlNumEndoso As Integer
'Dim vlNumOrden As Integer
Dim vlNumResGarEst As String
Dim vlNumAnnoRes As String
Dim vlCodTipRes As String
Dim vlPrcDeduccion As Double
Dim vlCodDerGarEst As String
Dim vlMtoPension As Double
'Dim vlMtoPensionPesos As Double
'Dim vlMtoGarEst As Double
'Dim vlMtoOtroHab As Double
'Dim vlMtoOtroDes As Double
Dim vlMtoBono As Double

Dim vlCodConHabDes As String
Dim vlNumCuotas As Integer
Dim vlMtoCuota As Double
Dim vlMtoTotal As Double
Dim vlCodMoneda As String
Dim vlMtoTotalHabDes As Double
Dim vlCodTipMov As String

Const clCodConHDMtoPen As String * 2 = "01"
Const clCodConHDMtoPen0102 As String = "('01','02')"
Const clModOrigenPP As String * 2 = "PP"
Const clCodGE As String * 2 = "03"
Const clCodBonInv As String * 2 = "06"
'Const clCodHabGE As String = "('02','04','07')"
'Const clCodDesGE As String = "('21','22')"
Const clTipoGE As String * 2 = "GE"
Const clCodH As String * 1 = "H"
Const clCodD As String * 1 = "D"

'----------------------------------
'Conciliacion

Dim vlRegistroCia As ADODB.Recordset
Dim vlRegistroTes As ADODB.Recordset
Dim vlRegistroPeriodo As ADODB.Recordset

'Dim vlNumPoliza As String
'Dim vlNumOrden As Integer
'Dim vlNumEndoso As Integer
Dim vlCodTipoImp As String
Dim vlNumPeriodo As String
Dim vlFecPago As String
'Dim vlCodTipRes As String
'Dim vlNumResGarEst As Integer
'Dim vlNumAnnoRes As Integer
Dim vlNumDias As Integer
'Dim vlMtoPension As Double
Dim vlMtoPensionUF As Double
Dim vlMtoPensionMin As Double
'Dim vlPrcDeduccion As Double
Dim vlMtoDeduccion As Double
Dim vlMtoGarEstQui As Double
Dim vlMtoGarEstNor As Double
Dim vlMtoGarEstCia As Double
Dim vlMtoGarEstRec As Double
Dim vlMtoDiferencia As Double
'Dim vlCodDerGarEst As String
'Dim vlMtoHaber As Double
Dim vlMtoDescuento As Double
Dim vlCodEstado As String

'Dim vlCodPar As String
Dim vlCodSexo As String
'Dim vlCodSitInv As String
Dim vlFecNacBen As String
Dim vlEdadBen As Integer
Dim vlAnno As String
Dim vlMes As String
Dim vlDia As String
Dim vlMtoCalculado As Boolean

Dim vlPeriodoInicio As String
Dim vlPeriodoTer As String

Dim vlTipoPago As String

'----------------------------------
Const clCodConHDHab As String * 2 = "04" 'MA_TPAR_TABCOD / Garantía Estatal RetroActiva
Const clCodConHDDes As String * 2 = "33" 'MA_TPAR_TABCOD / Descuento Pensionado por Garantía Estatal
'Const clCodConHDMtoPen As String * 2 = "01" 'MA_TPAR_TABCOD / Pensión Renta Vitalicia
Const clCodConHDGECia As String * 2 = "03"
'Const clModOrigenPP As String * 2 = "PP" 'Pago de Pensiones
'Const clCodGE As String * 2 = "03"
'Const clCodBonInv As String * 2 = "06"
Const clCodHabGE As String = "('02','04','07')"
Const clCodDesGE As String = "('21','22')"
'Const clTipoGE As String * 2 = "GE" 'Garantía Estatal
'Const clCodH As String * 1 = "H" 'Haber
'Const clCodD As String * 1 = "D" 'Descuento
Const clTipoImpC As String * 1 = "C" 'Conciliación
Const clTipoImpE As String * 1 = "E" 'Exceso
Const clTipoImpD As String * 1 = "D" 'Deficit
Const clSinCodDerGE As String * 1 = "N" 'Sin Estado

Const clTipoPagoR As String * 1 = "R"
Const clTipoPagoP As String * 1 = "P"
'-----------------------------------------------------------------------FIN

Private Sub Cmd_Calcular_Click()

If Txt_Mes = "" Then
    Exit Sub
End If
If Txt_Anno = "" Then
    Exit Sub
End If

'Retencion Judicial
If Not flLlenaTablaTemporal(Txt_Mes, Txt_Anno, vlInfSeleccionado) Then
    Exit Sub
End If

MsgBox "Calculo para Informe de Control realizado exitosamente", vbInformation, Me.Caption

End Sub

Private Sub Cmd_Imprimir_Click()
On Error GoTo Err_Imprimir

   If Txt_Mes.Text = "" Then
      MsgBox "Debe Ingresar Mes del Período de Pago.", vbCritical, "Error de Datos"
      Txt_Mes.SetFocus
      Exit Sub
   End If
   If CDbl(Txt_Mes.Text) <= 0 Or CDbl(Txt_Mes.Text) > 12 Then
      MsgBox "El Mes Ingresado No es un Valor Válido.", vbCritical, "Error de Datos"
      Txt_Mes.SetFocus
      Exit Sub
   End If
   Txt_Mes.Text = Format(Txt_Mes.Text, "00")
   
   'Valida Año del Periodo
   If Txt_Anno.Text = "" Then
      MsgBox "Debe Ingresar Año del Período de Pago.", vbCritical, "Error de Datos"
      Txt_Anno.SetFocus
      Exit Sub
   End If
   If CDbl(Txt_Anno.Text) < 1900 Then
      MsgBox "Debe Ingresar un Año Mayor a 1900.", vbCritical, "Error de Datos"
      Txt_Anno.SetFocus
      Exit Sub
   End If
   Txt_Anno.Text = Format(Txt_Anno.Text, "0000")
      
   If vlInfSeleccionado = "RJ" Then
      Call flProcesoRetJudicial
   End If
   
Exit Sub
Err_Imprimir:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub cmd_salir_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Frm_InformeControl.Top = 0
    Frm_InformeControl.Left = 0
        
    vlInfSeleccionado = vgTipoInforme
    Frm_InformeControl.Caption = vgTituloInfControl
    
End Sub

Private Sub Txt_Anno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Txt_Anno <> "" Then
       If Txt_Anno < 1900 Then
           MsgBox "Año ingresado es menor a la mínima que se puede ingresar (1900).", vbCritical, "Dato Incorrecto"
       Else
         Cmd_Calcular.SetFocus
       End If
    End If
End If
End Sub

Private Sub Txt_Anno_LostFocus()
If Txt_Anno <> "" Then
   If Txt_Anno < 1900 Then
      Txt_Anno = ""
   End If
End If
End Sub

Private Sub Txt_Mes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Txt_Mes <> "" Then
       If Txt_Mes >= 1 And Txt_Mes <= 12 Then
          Txt_Mes = Trim(Format(Txt_Mes, "00"))
          Txt_Anno.SetFocus
       Else
         Txt_Mes = ""
       End If
    End If
End If
End Sub

Private Sub Txt_Mes_LostFocus()

If Txt_Mes <> "" Then
   If Txt_Mes >= 1 And Txt_Mes <= 12 Then
      Txt_Mes = Trim(Format(Txt_Mes, "00"))
   Else
      Txt_Mes = ""
   End If
End If

End Sub

Function flLlenaTablaTemporal(iMes As Long, iAño As Long, iModulo As String) As Boolean
'Función que llena la Tabla de Control de Asignación Familiar

On Error GoTo Errores

Dim vlPerPago As String, vlPerPriPago As String, vlPerPagoReg As String
Dim vlFecIniPag As Date, vlFecTerPag As Date
Dim vlFecIniPriPag As Date, vlFecTerPriPag As Date
Dim vlFecPago As String
Dim vlUF As Double
Dim vlModulo As String
'HQR 17/05/2005
Dim vlPension As Double 'Pension de RV o Garantizada
Dim vlBaseImp As Double, vlBaseTrib As Double 'Para calcular la Base Imponible y Base Tributable
Dim vlSql As String, vlTB As ADODB.Recordset
Dim vlcont As Double, vlAumento As Double
Dim vlTB2 As ADODB.Recordset
Dim bResp As Integer 'Retorno de las Funciones
Dim vlGarantia As Double 'Monto de la Garantía Estatal
Dim vlDesSalud As Double 'Descuento de Salud
Dim vlDesSalud2 As Double 'Descuento de Salud2
Dim vlImpuesto As Double 'Monto del Impuesto Único
Dim vlBaseLiq As Double 'Base Tributable - Impuesto Único
Dim vlAsigFamiliar As Double 'Monto Asignacion Familiar
'FIN HQR 17/05/2005
'HQR 09/07/2005
Dim vlPensionQuiebra As Double 'Monto de la Pensión en Quiebra
'Fin HQR 09/07/2005
Dim vlUFUltDia As Double 'UF al último día del Mes de Pago
Dim vlRecibePension As Boolean

Dim vlMontoPensionAct As Double 'Monto de la Pensión Actualizada
Dim vlMontoAcumPagado As Double 'Monto Acumulado por Poliza (para ver si cuadra con el Total por Beneficiarios)
Dim vlTCMonedaPension As Double 'Tipo de Cambio de la Moneda de la Pensión
Dim vlMonedaPension As String, vlMonedaPensionAnt As String 'Moneda de la Pensión (Para no buscar siempre el T.C., solo cuando cambia la moneda)
Dim vlEsGarantizado As Boolean, vlNumOrdenDiferencia As Long
Dim vlFactorAjuste As Double, vlFactorAjuste2 As Double
Dim vlMes As Long
Dim vlFechaAjuste As String
Dim i As Double

'hqr 05/09/2007 Agregados para Calculo de Primera Pensión Diferida
Dim vlPrimerMesAjuste As Boolean
Dim vlFactorAjusteDif As Double
Dim vlFactorAjusteDif2 As Double
Dim vlFechaIniDif As Date
Dim vlFechaFinDif As Date
Dim vlMesDif As Long, vlAñoDif As Long
Dim vlFecDesdeAjuste As String
'fin hqr 05/09/2007
Dim vlDiaAnteriorFecPago As String

'hqr 05/01/2008 para que limpie el grupo familiar, al cambiar de póliza o de Grupo Familiar
Dim vlPolizaAnt As Long
Dim vlGruFamAnt As Long
'fin hqr 05/01/2008

Dim vlFactorAjusteTasaFija As Double 'hqr 03/12/2010
Dim vlFactorAjusteporIPC As Double 'hqr 03/12/2010
Dim vlMesesDiferencia As Long 'hqr 18/02/2011

Dim vlNumPension As Double 'hqr 18/02/2011
Dim vlPasaTrimestre As Boolean 'hqr 18/02/2011
Dim vlFecDesdeAjustePension As String 'hqr 05/03/2011 Fecha desde la cual se realizan los ajustes


Screen.MousePointer = 11
flLlenaTablaTemporal = False

vlModulo = iModulo
If Not fgConexionBaseDatos(vgConexionTransac) Then
    MsgBox "Error en Conexion a la Base de Datos", vbCritical, Me.Caption
    Exit Function
End If
vgConexionTransac.BeginTrans

'0.- Elimina Registros Anteriores
If Not fgBorraCalculosControlAnteriores(vlModulo, Me.Caption) Then
    GoTo Deshacer
End If

vlPerPagoReg = iAño * 100 + iMes

'0.1.- Obtiene Parametros de Quiebra (REVISAR si es por Periodo de Pago o Fecha de Pago)
stDatGenerales.Ind_AplicaQuiebra = fgObtieneParametrosQuiebra(vlPerPagoReg, stDatGenerales.Prc_Castigo, stDatGenerales.Mto_TopMaxQuiebra)

'Obtiene Datos del Pago en Régimen
If Not fgObtieneDatosPagoRegimen(vlPerPagoReg, stDatosInformeControl.Fec_PagoReg, stDatosInformeControl.Val_UFReg, vlFecIniPag, vlFecTerPag, Me.Caption, stDatosInformeControl.Val_UFRegUltDia, stDatosInformeControl.Prc_SaludMinimoReg, stDatosInformeControl.Mto_MaximoSaludReg, stDatosInformeControl.Mto_TopeBaseImponReg) Then
    GoTo Deshacer
End If

'Cuenta Nº de Pólizas a Procesar
vlSql = flObtieneQuery(iModulo, vlFecIniPag, vlFecTerPag, 2)
Set vlTB = vgConexionBD.Execute(vlSql)
If Not vlTB.EOF Then
    vlcont = vlTB!contador
    If vlcont = 0 Then
        MsgBox "No existen Datos para este periodo", vbCritical, Me.Caption
        GoTo Deshacer
    End If
Else
    MsgBox "No existen Datos para este período", vbCritical, Me.Caption
    GoTo Deshacer
End If

'Obtiene Factor de Ajuste
'0 Obtiene Factor de Ajuste
vlFactorAjuste = 1 'Sin Ajuste
vlFactorAjusteTasaFija = 0 'hqr 03/12/2010
vlMes = Mid(vlPerPagoReg, 5, 2)
If (((vlMes Mod 3) - 1) = 0) Then
    'Obtiene Factor de Ajuste Anterior
    vlFechaAjuste = Format(DateSerial(Mid(vlPerPagoReg, 1, 4), Mid(vlPerPagoReg, 5, 2) - 4, 1), "yyyymmdd")
    'Obtener Factor anterior
    If Not fgObtieneFactorAjuste(vlFechaAjuste, vlFactorAjuste) Then
        vgConexionTransac.CommitTrans
        vgConexionTransac.Close
        MsgBox "No se encuentra Factor de Ajuste del Periodo: " & DateSerial(Mid(vlFechaAjuste, 1, 4), Mid(vlFechaAjuste, 5, 2), Mid(vlFechaAjuste, 7, 2)), vbCritical, Me.Caption
        Exit Function
    End If
    'Obtiene Factor de Ajuste Actual
    vlFechaAjuste = Format(DateSerial(Mid(vlPerPagoReg, 1, 4), Mid(vlPerPagoReg, 5, 2) - 1, 1), "yyyymmdd")
    'Factor de Ajuste Mes actual
    If Not fgObtieneFactorAjuste(vlFechaAjuste, vlFactorAjuste2) Then
        vgConexionTransac.CommitTrans
        vgConexionTransac.Close
        MsgBox "No se encuentra Factor de Ajuste del Periodo: " & DateSerial(Mid(vlFechaAjuste, 1, 4), Mid(vlFechaAjuste, 5, 2), Mid(vlFechaAjuste, 7, 2)), vbCritical, Me.Caption
        Exit Function
    End If
    vlFactorAjuste = Format(vlFactorAjuste2 / vlFactorAjuste, "##0.0000")
    vlFactorAjusteTasaFija = 1
End If

vlSql = flObtieneQuery(iModulo, vlFecIniPag, vlFecTerPag, 1) 'Obtiene Query para filtrar las Pólizas

'hqr 05/01/2008 para que limpie el grupo familiar, al cambiar de póliza o de Grupo Familiar
vlMonedaPensionAnt = ""
vlPolizaAnt = -1
vlGruFamAnt = -1
'fin hqr 05/01/2008

vlFactorAjusteporIPC = vlFactorAjuste 'hqr 03/12/2010
Set vlTB = vgConexionBD.Execute(vlSql)
If Not vlTB.EOF Then
    vlAumento = (100 / vlcont)
    ProgressBar.Refresh
    stConTutor.Cod_GruFami = "-1"
    Do While Not vlTB.EOF
        'Registra Datos de la Poliza en la Estructura
        stConLiquidacion.num_poliza = vlTB!num_poliza
        stConLiquidacion.num_endoso = vlTB!num_endoso
        stConLiquidacion.Cod_TipPension = vlTB!Cod_TipPension
        stConLiquidacion.Cod_Moneda = vlTB!Cod_Moneda
        vlEsGarantizado = False
        vlNumOrdenDiferencia = -1
        'hqr 03/12/2010
        If vlTB!Cod_TipReajuste = cgAJUSTESOLES Then
            vlFactorAjuste = vlFactorAjusteporIPC
        ElseIf vlTB!Cod_TipReajuste = cgAJUSTETASAFIJA Then
            'Buscar la tasa de Ajuste, dependiendo de la fecha de vigencia
            vlFactorAjuste = fgObtieneFactorAjusteTasaFija(vlTB!Fec_Vigencia, vlFecIniPag, vlTB!Mto_ValReajusteTri, vlTB!Mto_ValReajusteMen, vlFactorAjusteTasaFija, vlTB!fec_dev)
        Else
            vlFactorAjuste = 1 'sin ajuste
        End If
        'fin hqr 03/12/2010
        
        If vlTB!TIPPAGO = "R" Then
            vlUF = stDatosInformeControl.Val_UFReg
            vlFecPago = stDatosInformeControl.Fec_PagoReg
            vlPerPago = vlPerPagoReg
            vlUFUltDia = stDatosInformeControl.Val_UFRegUltDia
            stDatGenerales.Prc_SaludMin = stDatosInformeControl.Prc_SaludMinimoReg
            'stDatGenerales.Mto_MaxSalud = stDatosInformeControl.Mto_MaximoSaludReg
            'stDatGenerales.Mto_TopeBaseImponible = stDatosInformeControl.Mto_TopeBaseImponReg
        End If
        vlDiaAnteriorFecPago = Format(DateAdd("d", -1, DateSerial(Mid(vlFecPago, 1, 4), Mid(vlFecPago, 5, 2), Mid(vlFecPago, 7, 2))), "yyyymmdd")

        '********************************************************************
        'Actualizacion Pension Diferida
        '********************************************************************
        'Obtiene Monto de la Pensión Actualizada
        'Si es primer pago de una diferida, se debe obtener la pensión actualizada
        
        vlPrimerMesAjuste = True
        'If stConLiquidacion.Cod_Moneda = vgMonedaCodOfi Then 'Nuevos Soles
        If vlTB!Cod_TipReajuste <> cgSINAJUSTE Then 'hqr 03/12/2010
            'Solo si está en Nuevos Soles se hace la actualización
            'desde la Fecha de Devengamiento
            vlMontoPensionAct = vlTB!Pensionact 'Por Defecto es la Pensión de la Póliza
            
            'Obtiene Pension Actualizada
            vlSql = "SELECT mto_pension FROM pp_tmae_pensionact a"
            vlSql = vlSql & " WHERE a.num_poliza = '" & stConLiquidacion.num_poliza & "'"
            vlSql = vlSql & " AND a.num_endoso = " & stConLiquidacion.num_endoso
            vlSql = vlSql & " AND a.fec_desde = "
                vlSql = vlSql & " (SELECT max(fec_desde) FROM pp_tmae_pensionact b"
                vlSql = vlSql & " WHERE b.num_poliza = a.num_poliza"
                vlSql = vlSql & " AND b.num_endoso = a.num_endoso"
                vlSql = vlSql & " AND b.fec_desde < '" & Format(vlFecIniPag, "yyyymmdd") & "')"
            Set vlTB2 = vgConexionBD.Execute(vlSql)
            If Not vlTB2.EOF Then
                If Not IsNull(vlTB2!Mto_Pension) Then
                    vlMontoPensionAct = vlTB2!Mto_Pension
                End If
            End If
            
            'If (vlTB!Num_MesDif > 0 And vlTB!Fec_PriPago = Format(vgFecIniPag, "yyyymmdd")) Then
            If (vlTB!Num_MesDif > 0 And vlTB!Fec_PriPago = vlTB!Fec_IniPenCia And vlTB!Fec_PriPago = Format(vlFecIniPag, "yyyymmdd")) Then
                'Obtiene Factores de actualización
                
                vlFactorAjusteDif = 1 'Sin Ajuste
                
                'hqr 05/03/2011
                If clAjusteDesdeFechaDevengamiento Then
                    vlFecDesdeAjustePension = Mid(vlTB!fec_dev, 1, 6) & "01" 'Fecha de Devengamiento, primer dia del mes
                Else
                    vlFecDesdeAjustePension = vlTB!Fec_Vigencia 'Fecha de Inicio de Vigencia de la Póliza
                End If
                'fin hqr  05/03/2011
                
                If Not IsNull(vlTB!FEC_EFECTO) Then
                    If vlTB!FEC_EFECTO <> "" Then
                        vlFechaIniDif = DateSerial(Mid(vlTB!FEC_EFECTO, 1, 4), Mid(vlTB!FEC_EFECTO, 5, 2) + 1, Mid(vlTB!FEC_EFECTO, 7, 2))
                        vlMontoPensionAct = vlTB!Pensionpol 'Pension de la Póliza
                    Else
                        'vlFechaIniDif = DateSerial(Mid(vlTB!fec_dev, 1, 4), Mid(vlTB!fec_dev, 5, 2) + 1, Mid(vlTB!fec_dev, 7, 2))
                        'vlFechaIniDif = DateSerial(Mid(vlTB!Fec_Vigencia, 1, 4), Mid(vlTB!Fec_Vigencia, 5, 2) + 1, Mid(vlTB!Fec_Vigencia, 7, 2)) 'hqr 18/02/2011 Se actualiza desde la fecha de inicio de vigencia
                        vlFechaIniDif = DateSerial(Mid(vlFecDesdeAjustePension, 1, 4), Mid(vlFecDesdeAjustePension, 5, 2) + 1, Mid(vlFecDesdeAjustePension, 7, 2)) 'hqr 05/03/2011
                    End If
                Else
                    'vlFechaIniDif = DateSerial(Mid(vlTB!fec_dev, 1, 4), Mid(vlTB!fec_dev, 5, 2) + 1, Mid(vlTB!fec_dev, 7, 2))
                    'vlFechaIniDif = DateSerial(Mid(vlTB!Fec_Vigencia, 1, 4), Mid(vlTB!Fec_Vigencia, 5, 2) + 1, Mid(vlTB!Fec_Vigencia, 7, 2))
                    vlFechaIniDif = DateSerial(Mid(vlFecDesdeAjustePension, 1, 4), Mid(vlFecDesdeAjustePension, 5, 2) + 1, Mid(vlFecDesdeAjustePension, 7, 2)) 'hqr 05/03/2011
                End If
                vlFechaFinDif = DateAdd("m", -1, vlFecIniPag)
                vlNumPension = 2 'hqr 18/02/2011
                vlPasaTrimestre = False 'hqr 18/02/2011
                Do While vlFechaFinDif >= vlFechaIniDif
                    vlMesDif = Month(vlFechaIniDif)
                    vlAñoDif = Year(vlFechaIniDif)
                    If (((vlMesDif Mod 3) - 1) = 0) Then
                        'inicio hqr 03/12/2010
                        If ((vlAñoDif <> Mid(vlFecDesdeAjustePension, 1, 4) Or vlMesDif <> Mid(vlFecDesdeAjustePension, 5, 2))) Then 'No se ajusta el primer mes
                            'inicio hqr 03/12/2010
                            If vlTB!Cod_TipReajuste = cgAJUSTETASAFIJA Then
                                If (vlNumPension >= 2 And vlNumPension <= 3) And (Not (vlPasaTrimestre)) Then
                                    vlFactorAjusteDif = (1 + (vlTB!Mto_ValReajusteMen / 100))
                                Else
                                    vlFactorAjusteDif = (1 + (vlTB!Mto_ValReajusteTri / 100))
                                End If
                            Else
                            'fin hqr 03/12/2010
                                 'Obtiene Factor de Ajuste Anterior
                                If vlPrimerMesAjuste Then
                                     vlFechaAjuste = Format(DateSerial(Year(vlFechaIniDif), Month(vlFechaIniDif) - 4, 1), "yyyymmdd")
                                     'Obtener Factor anterior
                                     If Not fgObtieneFactorAjuste(vlFechaAjuste, vlFactorAjusteDif) Then
                                         MsgBox "No se encuentra Factor de Ajuste del Periodo: " & DateSerial(Mid(vlFechaAjuste, 1, 4), Mid(vlFechaAjuste, 5, 2), Mid(vlFechaAjuste, 7, 2)), vbCritical, Me.Caption
                                         GoTo Deshacer
                                     End If
                                 End If
                                 vlPrimerMesAjuste = False
                                 
                                 'Obtiene Factor de Ajuste Actual
                                 vlFechaAjuste = Format(DateSerial(Year(vlFechaIniDif), Month(vlFechaIniDif) - 1, 1), "yyyymmdd")
                                 'Factor de Ajuste Mes actual
                                 If Not fgObtieneFactorAjuste(vlFechaAjuste, vlFactorAjusteDif2) Then
                                     MsgBox "No se encuentra Factor de Ajuste del Periodo: " & DateSerial(Mid(vlFechaAjuste, 1, 4), Mid(vlFechaAjuste, 5, 2), Mid(vlFechaAjuste, 7, 2)), vbCritical, Me.Caption
                                     GoTo Deshacer
                                 End If
                                 vlFactorAjusteDif = Format(vlFactorAjusteDif2 / vlFactorAjusteDif, "##0.0000")
                            End If 'hqr 03/12/2010
                            vlMontoPensionAct = Format(vlFactorAjusteDif * vlMontoPensionAct, "#0.00") 'La última Actualización
                            If vlTB!Cod_TipReajuste <> cgAJUSTETASAFIJA Then
                                vlFactorAjusteDif = vlFactorAjusteDif2
                            End If
                            vlPasaTrimestre = True 'hqr 14/02/2011
                        End If
                        vlFecDesdeAjuste = vlFechaIniDif
                    Else 'No está en mes del trimestre
                        If (vlTB!Cod_TipReajuste = cgAJUSTETASAFIJA) And (vlNumPension >= 2 And vlNumPension <= 3) And (Not (vlPasaTrimestre)) Then
                            vlFactorAjusteDif = (1 + (vlTB!Mto_ValReajusteMen / 100))
                            vlMontoPensionAct = Format(vlFactorAjusteDif * vlMontoPensionAct, "#0.00") 'La última Actualización
                            vlFecDesdeAjuste = vlFechaIniDif
                        End If
                    End If
                    vlFechaIniDif = DateAdd("m", 1, vlFechaIniDif)
                    vlNumPension = vlNumPension + 1 'hqr 14/02/2011
                Loop
                
                If Not IsNull(vlTB!FEC_EFECTO) Then
                    If vlTB!FEC_EFECTO <> Format(vlFecIniPag, "yyyymmdd") Then
                        'Se actualiza con factor del mes actual
                        vlMontoPensionAct = Format(vlFactorAjuste * vlMontoPensionAct, "#0.00") 'La última Actualización
                    End If
                Else
                    'Se actualiza con factor del mes actual
                    vlMontoPensionAct = Format(vlFactorAjuste * vlMontoPensionAct, "#0.00") 'La última Actualización
                End If
            Else
                If Not IsNull(vlTB!FEC_EFECTO) Then 'Para que el primer mes del endoso no actualice pension
                    If vlTB!FEC_EFECTO = Format(vlFecIniPag, "yyyymmdd") Then
                        'vlMontoPensionAct = Format(vlTB!Pensionact, "#0.00")
                        vlMontoPensionAct = Format(vlFactorAjuste * vlMontoPensionAct, "#0.00")
                    Else
                        vlMontoPensionAct = Format(vlFactorAjuste * vlMontoPensionAct, "#0.00")
                    End If
                Else
                    vlMontoPensionAct = Format(vlFactorAjuste * vlMontoPensionAct, "#0.00")
                End If
            End If
        Else
            vlMontoPensionAct = Format(vlTB!Pensionact, "#0.00")
        End If
        '********************************************************************
        
        vlMontoAcumPagado = 0 'Reinica Monto
        vlMonedaPension = vlTB!Cod_Moneda
        If vlMonedaPensionAnt <> vlMonedaPension Then
            'Obtiene Tipo de Cambio
            'If Not fgObtieneConversion(vlFecPago, vlMonedaPension, vlTCMonedaPension) Then
            'hqr 12/12/2007 Comentado a peticion de MCHirinos
            vlTCMonedaPension = 1
'            If Not fgObtieneConversion(vlDiaAnteriorFecPago, vlMonedaPension, vlTCMonedaPension) Then
'                MsgBox "Debe ingresar el Tipo de Cambio de la Moneda '" & vlMonedaPension & "' a la Fecha de Pago", vbCritical, "Falta Tipo de Cambio"
'                GoTo Deshacer
'            End If
            'fin hqr 12/12/2007 Comentado a peticion de MCHirinos
        End If
        vlMonedaPensionAnt = vlMonedaPension
        
        stConLiquidacion.Num_PerPago = vlPerPago
        stConLiquidacion.Cod_TipoPago = vlTB!TIPPAGO
        stConLiquidacion.Fec_Pago = vlFecPago
        'Obtiene Beneficiarios de la Póliza que tengan Derecho a Pensión
        vlSql = "SELECT * FROM PP_TMAE_BEN"
        vlSql = vlSql & " WHERE NUM_POLIZA = '" & vlTB!num_poliza & "'"
        vlSql = vlSql & " AND NUM_ENDOSO = " & vlTB!num_endoso
        vlSql = vlSql & " AND NUM_ORDEN = " & vlTB!Num_Orden
        'vlSQL = vlSQL & " AND COD_DERPEN = 99" 'Solo los que tienen Derecho a Pensión
        vlSql = vlSql & " AND COD_ESTPENSION = 99" 'Solo los que tienen Derecho a Pensión
        vlSql = vlSql & " AND FEC_INIPAGOPEN <= '" & Format(vlFecIniPag, "yyyymmdd") & "'" 'Solo los que ya iniciaron su pago de pensión o lo inician en este periodo
        vlSql = vlSql & " ORDER BY NUM_POLIZA,COD_GRUFAM, COD_PAR"
        Set vlTB2 = vgConexionBD.Execute(vlSql)
        i = 0
        If Not vlTB2.EOF Then
            Do While Not vlTB2.EOF
                vlRecibePension = True
                 
                bResp = fgCalculaEdad(vlTB2!Fec_NacBen, vlFecIniPag)
                If bResp = "-1" Then 'Error
                    GoTo Deshacer
                End If
                stConDetPension.Edad = bResp
                stConDetPension.EdadAños = fgConvierteEdadAños(stConDetPension.Edad)
                'Si son Hijos se Calcula la Edad y se Verifica Certificado de Estudios
                If vlTB2!Cod_Par >= 30 And vlTB2!Cod_Par <= 35 Then 'Hijos
                    If stConDetPension.Edad >= stDatGenerales.MesesEdad18 And vlTB2!Cod_SitInv = "N" Then 'Hijos Sanos
                        'OBS: Se asume que el mes de los 18 años se paga completo
                        'Verifica Certificados de Estudio
                        vlRecibePension = False 'No recibe pensión, por lo que no se envía al arreglo de Beneficiarios
                    End If
                End If
                'Valida Certificado de Supervivencia
                bResp = fgVerificaCertificado(stConLiquidacion.num_poliza, vlTB2!Num_Orden, vlFecIniPag, vlFecTerPag, "SUP")
                If bResp = "-1" Then 'Error
                    GoTo Deshacer
                Else
                    If bResp = "0" Then 'No tiene Certificado de Supervivencia
                        vlRecibePension = False 'Va al Siguiente Beneficiario, ya que éste no tiene Derecho
                    End If
                End If
                
                If vlRecibePension Then 'Para los que tienen Derecho a Pensión Obtiene Monto de la Pensión
                    'Verifica si está en Periodo Garantizado
                    If Not IsNull(vlTB2!Fec_TerPagoPenGar) Then
                        If vlTB2!Fec_TerPagoPenGar >= Format(vlFecIniPag, "yyyymmdd") Then
                            vlPension = Format(vlMontoPensionAct * (vlTB2!Prc_PensionGar / 100), "#0.00")
                            vlEsGarantizado = True
                        Else
                            vlPension = Format(vlMontoPensionAct * (vlTB2!Prc_Pension / 100), "#0.00")
                        End If
                    Else
                        vlPension = Format(vlMontoPensionAct * (vlTB2!Prc_Pension / 100), "#0.00")
                    End If
                    
                    ReDim Preserve stConBeneficiarios(i)
                    stConBeneficiarios(i).Mto_Pension = vlPension
                    vlMontoAcumPagado = vlMontoAcumPagado + vlPension
                    stConBeneficiarios(i).num_poliza = vlTB!num_poliza
                    stConBeneficiarios(i).num_endoso = vlTB!num_endoso
                    stConBeneficiarios(i).Num_Orden = vlTB2!Num_Orden
                    stConBeneficiarios(i).Cod_Par = vlTB2!Cod_Par
                    stConBeneficiarios(i).Cod_GruFam = vlTB2!Cod_GruFam
                    stConBeneficiarios(i).Cod_TipoIdenBen = vlTB2!Cod_TipoIdenBen
                    stConBeneficiarios(i).Gls_MatBen = IIf(IsNull(vlTB2!Gls_MatBen), "", vlTB2!Gls_MatBen)
                    stConBeneficiarios(i).Gls_PatBen = vlTB2!Gls_PatBen
                    stConBeneficiarios(i).Gls_NomBen = vlTB2!Gls_NomBen
                    stConBeneficiarios(i).Gls_NomSegBen = IIf(IsNull(vlTB2!Gls_NomSegBen), "", vlTB2!Gls_NomSegBen)
                    stConBeneficiarios(i).Num_IdenBen = vlTB2!Num_IdenBen
                    stConBeneficiarios(i).Gls_DirBen = vlTB2!Gls_DirBen
                    stConBeneficiarios(i).Cod_Direccion = vlTB2!Cod_Direccion
                    stConBeneficiarios(i).Cod_ViaPago = vlTB2!Cod_ViaPago
                    stConBeneficiarios(i).Cod_Banco = IIf(IsNull(vlTB2!Cod_Banco), "NULL", vlTB2!Cod_Banco)
                    stConBeneficiarios(i).Cod_TipCuenta = IIf(IsNull(vlTB2!Cod_TipCuenta), "NULL", vlTB2!Cod_TipCuenta)
                    stConBeneficiarios(i).Num_Cuenta = IIf(IsNull(vlTB2!Num_Cuenta), "NULL", vlTB2!Num_Cuenta)
                    stConBeneficiarios(i).Cod_Sucursal = IIf(IsNull(vlTB2!Cod_Sucursal), "NULL", vlTB2!Cod_Sucursal)
                    stConBeneficiarios(i).Fec_NacBen = vlTB2!Fec_NacBen
                    stConBeneficiarios(i).Cod_SitInv = vlTB2!Cod_SitInv
                    stConBeneficiarios(i).Fec_TerPagoPenGar = IIf(IsNull(vlTB2!Fec_TerPagoPenGar), "NULL", vlTB2!Fec_TerPagoPenGar)
                    stConBeneficiarios(i).Prc_PensionGar = vlTB2!Prc_PensionGar
                    stConBeneficiarios(i).Prc_Pension = vlTB2!Prc_Pension
                    stConBeneficiarios(i).Cod_InsSalud = vlTB2!Cod_InsSalud
                    stConBeneficiarios(i).Cod_ModSalud = vlTB2!Cod_ModSalud
                    stConBeneficiarios(i).Mto_PlanSalud = vlTB2!Mto_PlanSalud
                    i = i + 1
                End If
                vlTB2.MoveNext
            Loop
            'Valida que el Monto de la Pensión sea el total garantizado
            If vlEsGarantizado Then
                'Validar si se pagó el 100%
                If vlMontoAcumPagado <> vlMontoPensionAct Then 'Se debe ajustar la diferencia
                    stConBeneficiarios(0).Mto_Pension = stConBeneficiarios(0).Mto_Pension + (vlMontoPensionAct - vlMontoAcumPagado) 'Suma o Resta al Primer Beneficiario calculado
                End If
            End If
        End If
        
        If i > 0 Then 'Si se le pagó a alguien
            i = 0
            Do While i <= UBound(stConBeneficiarios)
                'Llena Datos de la Estructura que no cambian
                stConDetPension.Fec_IniPago = Format(vlFecIniPag, "yyyymmdd")
                stConDetPension.Fec_TerPago = Format(vlFecTerPag, "yyyymmdd")
                stConDetPension.num_endoso = vlTB!num_endoso
                stConDetPension.Num_Orden = stConBeneficiarios(i).Num_Orden
                stConDetPension.num_poliza = vlTB!num_poliza
                stConDetPension.Num_PerPago = vlPerPago
                stConDetPension.Edad = 0
                
                'hqr 05/01/2008 para que limpie el grupo familiar, al cambiar de póliza o de Grupo Familiar
                If stConDetPension.num_poliza <> vlPolizaAnt Or vlGruFamAnt <> stConBeneficiarios(i).Cod_GruFam Then
                    stConTutor.Cod_GruFami = "-1"
                End If
                vlPolizaAnt = stConDetPension.num_poliza
                vlGruFamAnt = stConBeneficiarios(i).Cod_GruFam
                'fin hqr 05/01/2008
                
                If stConBeneficiarios(i).Cod_Par < 30 Then 'Padres quedan registrados para Tutores
                    stConTutor.Cod_GruFami = stConBeneficiarios(i).Cod_GruFam
                    stConTutor.Cod_TipReceptor = "M"
                    stConTutor.Cod_TipoIdenReceptor = stConBeneficiarios(i).Cod_TipoIdenBen
                    stConTutor.Gls_MatReceptor = stConBeneficiarios(i).Gls_MatBen
                    stConTutor.Gls_PatReceptor = stConBeneficiarios(i).Gls_PatBen
                    stConTutor.Gls_NomReceptor = stConBeneficiarios(i).Gls_NomBen
                    stConTutor.Gls_NomSegReceptor = stConBeneficiarios(i).Gls_NomSegBen
                    stConTutor.Num_IdenReceptor = stConBeneficiarios(i).Num_IdenBen
                    stConTutor.Gls_Direccion = stConBeneficiarios(i).Gls_DirBen
                    stConTutor.Cod_Direccion = stConBeneficiarios(i).Cod_Direccion
                    stConTutor.Cod_ViaPago = stConBeneficiarios(i).Cod_ViaPago
                    stConTutor.Cod_Banco = stConBeneficiarios(i).Cod_Banco
                    stConTutor.Cod_TipCuenta = stConBeneficiarios(i).Cod_TipCuenta
                    stConTutor.Num_Cuenta = stConBeneficiarios(i).Num_Cuenta
                    stConTutor.Cod_Sucursal = stConBeneficiarios(i).Cod_Sucursal
                End If
                
                'Inicializa Monto Haber y Descuento
                stConLiquidacion.Mto_Haber = 0
                stConLiquidacion.Mto_Descuento = 0
                stConLiquidacion.Num_Orden = stConBeneficiarios(i).Num_Orden
                
                '11.-  Obtener Tutores (1a. Etapa) (Se deja acá porque se necesita el Rut del Receptor)
                bResp = fgObtieneTutor(stConLiquidacion.num_poliza, stConLiquidacion.num_endoso, stConLiquidacion.Num_Orden, vlFecIniPag, vlFecTerPag, stConLiquidacion)
                If bResp = "-1" Then 'Error
                    GoTo Deshacer
                Else
                    If bResp = "0" Then 'No Encontró Tutor
                        If stConBeneficiarios(i).Cod_Par >= 30 And stConBeneficiarios(i).Cod_Par <= 35 And stConDetPension.Edad <= stDatGenerales.MesesEdad18 And stConTutor.Cod_GruFami = stConBeneficiarios(i).Cod_GruFam Then 'El Tutor debe ser la Madre
                            stConLiquidacion.Cod_TipReceptor = stConTutor.Cod_TipReceptor 'MADRE
                            stConLiquidacion.Num_IdenReceptor = stConTutor.Num_IdenReceptor
                            stConLiquidacion.Cod_TipoIdenReceptor = stConTutor.Cod_TipoIdenReceptor
                            stConLiquidacion.Gls_NomReceptor = stConTutor.Gls_NomReceptor
                            stConLiquidacion.Gls_NomSegReceptor = stConTutor.Gls_NomSegReceptor
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
                            stConLiquidacion.Num_IdenReceptor = stConBeneficiarios(i).Num_IdenBen
                            stConLiquidacion.Cod_TipoIdenReceptor = stConBeneficiarios(i).Cod_TipoIdenBen
                            stConLiquidacion.Gls_NomReceptor = stConBeneficiarios(i).Gls_NomBen
                            stConLiquidacion.Gls_NomSegReceptor = stConBeneficiarios(i).Gls_NomSegBen
                            stConLiquidacion.Gls_PatReceptor = stConBeneficiarios(i).Gls_PatBen
                            stConLiquidacion.Gls_MatReceptor = stConBeneficiarios(i).Gls_MatBen
                            stConLiquidacion.Gls_Direccion = stConBeneficiarios(i).Gls_DirBen
                            stConLiquidacion.Cod_Direccion = stConBeneficiarios(i).Cod_Direccion
                            stConLiquidacion.Cod_ViaPago = stConBeneficiarios(i).Cod_ViaPago
                            stConLiquidacion.Cod_Banco = stConBeneficiarios(i).Cod_Banco
                            stConLiquidacion.Cod_TipCuenta = stConBeneficiarios(i).Cod_TipCuenta
                            stConLiquidacion.Num_Cuenta = stConBeneficiarios(i).Num_Cuenta
                            stConLiquidacion.Cod_Sucursal = stConBeneficiarios(i).Cod_Sucursal
                        End If
                    'Else 'Encontró Tutor
                    End If
                End If
                
                stConDetPension.Num_IdenReceptor = stConLiquidacion.Num_IdenReceptor
                stConDetPension.Cod_TipoIdenReceptor = stConLiquidacion.Cod_TipoIdenReceptor
                stConDetPension.Cod_TipReceptor = stConLiquidacion.Cod_TipReceptor
                stConDetPension.Cod_ConHabDes = stDatGenerales.Cod_ConceptoPension
                stConLiquidacion.Mto_Pension = stConBeneficiarios(i).Mto_Pension
                stConDetPension.Mto_ConHabDes = stConLiquidacion.Mto_Pension
                stConLiquidacion.Mto_Haber = stConLiquidacion.Mto_Haber + stConLiquidacion.Mto_Pension
                'Graba Monto de la Pensión
                If Not fgInsertaDetallePensionProv(stConDetPension, "C", vlModulo) Then
                    MsgBox "Se ha producido un Error al Grabar Monto de la Pensión" & Chr(13) & Err.Description, vbCritical, Me.Caption
                    GoTo Deshacer
                End If
                vlBaseImp = stConBeneficiarios(i).Mto_Pension
                If vlTB!Cod_DerGra = "S" Then 'Gratificación se paga en Julio y Diciembre
                    If (vlMes = 7 Or vlMes = 12) Then
                        stConDetPension.Mto_ConHabDes = stConLiquidacion.Mto_Pension
                        stConLiquidacion.Mto_Haber = stConLiquidacion.Mto_Haber + stConLiquidacion.Mto_Pension
                        stConDetPension.Cod_ConHabDes = stDatGenerales.Cod_ConceptoGratificacion
                        'Graba Monto de la Pensión
                        If Not fgInsertaDetallePensionProv(stConDetPension, "C", vlModulo) Then
                            MsgBox "Se ha producido un Error al Grabar Monto de la Pensión" & Chr(13) & Err.Description, vbCritical, Me.Caption
                            GoTo Deshacer
                        End If
                        vlBaseImp = vlBaseImp + stConDetPension.Mto_ConHabDes
                    End If
                End If
                
                '4.- Obtener Haberes y Descuentos Imponibles (1a. Etapa)
                'If Not fgObtieneHaberesDescuentos("C", vlPerPago, vlTB!Num_Poliza, vlTB!num_endoso, stConBeneficiarios(i).Num_Orden, "S", "S", vlMonto, 0, 0, vlFecPago, stConLiquidacion, stConDetPension, vlUF, vlFecIniPag, vlFecTerPag, vlTB!Cod_Moneda, vlTCMonedaPension, "C", vlModulo) Then
                If Not fgObtieneHaberesDescuentos("C", vlPerPago, vlTB!num_poliza, vlTB!num_endoso, stConBeneficiarios(i).Num_Orden, "S", "S", vlMonto, 0, 0, vlDiaAnteriorFecPago, stConLiquidacion, stConDetPension, vlUF, vlFecIniPag, vlFecTerPag, vlTB!Cod_Moneda, vlTCMonedaPension, "C", vlModulo) Then
                    GoTo Deshacer
                End If
                
                vlBaseImp = vlBaseImp + vlMonto 'Base Imponible
                stConLiquidacion.Mto_BaseImp = vlBaseImp
                    
                '5.- Calcular Descto. Salud (1a. Etapa)
                vlDesSalud = 0
                'If Not fgObtienePrcSalud(stConBeneficiarios(i).Cod_InsSalud, stConBeneficiarios(i).Cod_ModSalud, stConBeneficiarios(i).Mto_PlanSalud, vlBaseImp, vlDesSalud, vlUFUltDia, vlFecPago, vlTB!Cod_Moneda, vlTCMonedaPension) Then
                If Not fgObtienePrcSalud(stConBeneficiarios(i).Cod_InsSalud, stConBeneficiarios(i).Cod_ModSalud, stConBeneficiarios(i).Mto_PlanSalud, vlBaseImp, vlDesSalud, vlUFUltDia, vlDiaAnteriorFecPago, vlTB!Cod_Moneda, vlTCMonedaPension) Then
                    GoTo Deshacer
                End If
                stConDetPension.Cod_ConHabDes = stDatGenerales.Cod_ConceptoDesctoSalud
                stConDetPension.Mto_ConHabDes = vlDesSalud
                stConLiquidacion.Mto_Descuento = stConLiquidacion.Mto_Descuento + vlDesSalud
                'Graba Monto del Descuento de Salud
                If Not fgInsertaDetallePensionProv(stConDetPension, "C", vlModulo) Then
                    MsgBox "Se ha producido un Error al Grabar Descuento de Salud" & Chr(13) & Err.Description, vbCritical, Me.Caption
                    GoTo Deshacer
                End If
                                    
                '6.- Agregar Haberes y Descuentos No Imponibles y Tributables (1a. Etapa)
                'If Not fgObtieneHaberesDescuentos("C", vlPerPago, vlTB!Num_Poliza, vlTB!num_endoso, stConBeneficiarios(i).Num_Orden, "N", "S", vlMonto, vlBaseImp, 0, vlFecPago, stConLiquidacion, stConDetPension, vlUF, vlFecIniPag, vlFecTerPag, vlTB!Cod_Moneda, vlTCMonedaPension, "C", vlModulo) Then
                If Not fgObtieneHaberesDescuentos("C", vlPerPago, vlTB!num_poliza, vlTB!num_endoso, stConBeneficiarios(i).Num_Orden, "N", "S", vlMonto, vlBaseImp, 0, vlDiaAnteriorFecPago, stConLiquidacion, stConDetPension, vlUF, vlFecIniPag, vlFecTerPag, vlTB!Cod_Moneda, vlTCMonedaPension, "C", vlModulo) Then
                    GoTo Deshacer
                End If
                
                vlBaseTrib = (vlBaseImp - vlDesSalud) + vlMonto 'Base Imponible
                stConLiquidacion.Mto_BaseTri = vlBaseTrib
                
                '9.- Calcular Retencion Judicial (2a. Etapa)
                'If Not fgCalculaRetencion(vlTB!Num_Poliza, vlTB!num_endoso, stConBeneficiarios(i).Num_Orden, vlPerPago, vlFecPago, vlBaseImp, vlBaseTrib, stConLiquidacion, stConDetPension, vlFecIniPag, Me.Caption, vlTB!Cod_Moneda, vlTCMonedaPension, "C", vlModulo) Then
                If Not fgCalculaRetencion(vlTB!num_poliza, vlTB!num_endoso, stConBeneficiarios(i).Num_Orden, vlPerPago, vlDiaAnteriorFecPago, vlBaseImp, vlBaseTrib, stConLiquidacion, stConDetPension, vlFecIniPag, Me.Caption, vlTB!Cod_Moneda, vlTCMonedaPension, "C", vlModulo) Then
                    GoTo Deshacer
                End If

                '10.- Agregar Haberes y Descuentos No Imponibles y No Tributables (1a. Etapa)
                'If Not fgObtieneHaberesDescuentos("C", vlPerPago, vlTB!Num_Poliza, vlTB!num_endoso, stConBeneficiarios(i).Num_Orden, "N", "N", vlMonto, vlBaseImp, vlBaseTrib, vlFecPago, stConLiquidacion, stConDetPension, vlUF, vlFecIniPag, vlFecTerPag, vlTB!Cod_Moneda, vlTCMonedaPension, "C", vlModulo) Then
                If Not fgObtieneHaberesDescuentos("C", vlPerPago, vlTB!num_poliza, vlTB!num_endoso, stConBeneficiarios(i).Num_Orden, "N", "N", vlMonto, vlBaseImp, vlBaseTrib, vlDiaAnteriorFecPago, stConLiquidacion, stConDetPension, vlUF, vlFecIniPag, vlFecTerPag, vlTB!Cod_Moneda, vlTCMonedaPension, "C", vlModulo) Then
                    GoTo Deshacer
                End If
                
                '13.-  Generar Liquidación (1a. Etapa)
                stConLiquidacion.Cod_InsSalud = stConBeneficiarios(i).Cod_InsSalud
                stConLiquidacion.Cod_ModSalud = stConBeneficiarios(i).Cod_ModSalud
                stConLiquidacion.Mto_PlanSalud = stConBeneficiarios(i).Mto_PlanSalud
                stConLiquidacion.Mto_LiqPagar = stConLiquidacion.Mto_Haber - stConLiquidacion.Mto_Descuento
                
                If stConLiquidacion.Mto_Haber > 0 Or stConLiquidacion.Mto_Descuento > 0 Then
                    'Inserta Liquidacion
                    If Not fgInsertaLiquidacion(stConLiquidacion, "C", vlModulo) Then
                        GoTo Deshacer
                    End If
                End If
Siguiente:
                i = i + 1
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
    GoTo Deshacer
End If

'''Traspasa Datos a Histórico
''If Me.Tag = "D" Then
''    If Not flTraspasaDatosADefinitivos Then
''        Exit Function
''    End If
''End If
vgConexionTransac.CommitTrans
vgConexionTransac.Close
flLlenaTablaTemporal = True

Screen.MousePointer = 0
Errores:

    If Err.Number <> 0 Then
Deshacer:
        vgConexionTransac.RollbackTrans
        vgConexionTransac.Close
        If Err.Number <> 0 Then
            MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
        End If
    End If
    ProgressBar.Value = 0
    Screen.MousePointer = 0

End Function
Private Function flObtieneQuery(iModulo As String, iFecIniPag As Date, iFecTerPag As Date, iModalidad As Integer) As String
'Función que obtiene la Query que obtendrá las Pólizas a Evaluar
'Devuelve el String que obtiene los Datos
'iModalidad: 1 => Query, 2 => Contador
Dim vlSql As String

Select Case iModulo
    Case "RJ" 'Retención Judicial
        vlSql = ""
        If iModalidad = 2 Then
            vlSql = "SELECT COUNT(1) AS CONTADOR FROM ("
        End If
        'los Pagos en Régimen
        vlSql = vlSql & "SELECT A.NUM_POLIZA, A.NUM_ENDOSO, B.NUM_ORDEN, A.COD_TIPPENSION, A.NUM_MESGAR, 'R' AS TIPPAGO, A.NUM_INDQUIEBRA,"
        vlSql = vlSql & " a.mto_pension AS pensionact, A.COD_MONEDA, a.cod_dergra, "
        vlSql = vlSql & " a.num_mesdif, a.fec_dev, a.fec_pripago, a.fec_inipencia, a.fec_efecto "
        vlSql = vlSql & ", a.cod_tipreajuste, a.mto_valreajustetri, a.mto_valreajustemen,fec_vigencia" 'hqr 03/12/2010
        vlSql = vlSql & " FROM PP_TMAE_POLIZA A, PP_TMAE_RETJUDICIAL B "
        vlSql = vlSql & " WHERE A.NUM_ENDOSO ="
        vlSql = vlSql & " (SELECT MAX(NUM_ENDOSO) FROM PP_TMAE_POLIZA " 'Último Endoso
        vlSql = vlSql & " WHERE NUM_POLIZA = A.NUM_POLIZA)"
        vlSql = vlSql & " AND COD_ESTADO IN (6, 7, 8) " 'solo Pólizas Vigentes, con Beneficiarios Designados y Pendiente (Pago Diferido)
        'Se deben identificar las Pólizas que tengan su Primer Pago en un Periodo Anterior (Inmediatas y Diferidas)
        'vlSql = vlSql & " AND FEC_INIPAGOPEN < '" & Format(iFecIniPag, "yyyymmdd") & "'"
        vlSql = vlSql & " AND (a.fec_pripago < '" & Format(iFecIniPag, "yyyymmdd") & "'"
        vlSql = vlSql & " OR (a.num_mesdif > 0 AND a.fec_pripago = '" & Format(iFecIniPag, "yyyymmdd") & "'))"
        vlSql = vlSql & " AND A.NUM_POLIZA = B.NUM_POLIZA"
        vlSql = vlSql & " AND '" & Format(iFecIniPag, "yyyymmdd") & "'"
        vlSql = vlSql & " BETWEEN B.FEC_INIRET AND B.FEC_TERRET"
        
        vlSql = vlSql & " UNION" 'HABERES Y DESCUENTOS
        vlSql = vlSql & " SELECT A.NUM_POLIZA, A.NUM_ENDOSO, B.NUM_ORDEN, A.COD_TIPPENSION, A.NUM_MESGAR, 'R' AS TIPPAGO, A.NUM_INDQUIEBRA,"
        vlSql = vlSql & " a.mto_pension AS pensionact, A.COD_MONEDA, a.cod_dergra,"
        vlSql = vlSql & " a.num_mesdif, a.fec_dev, a.fec_pripago, a.fec_inipencia, a.fec_efecto "
        vlSql = vlSql & ", a.cod_tipreajuste, a.mto_valreajustetri, a.mto_valreajustemen, a.fec_vigencia" 'hqr 03/12/2010
        vlSql = vlSql & " FROM PP_TMAE_POLIZA A, PP_TMAE_HABDES B, MA_TPAR_CONHABDES C"
        vlSql = vlSql & " WHERE A.NUM_ENDOSO ="
        vlSql = vlSql & " (SELECT MAX(NUM_ENDOSO) FROM PP_TMAE_POLIZA " 'Último Endoso
        vlSql = vlSql & " WHERE NUM_POLIZA = A.NUM_POLIZA)"
        vlSql = vlSql & " AND COD_ESTADO IN (6, 7, 8) " 'solo Pólizas Vigentes, con Beneficiarios Designados y Pendiente (Pago Diferido)
        'Se deben identificar las Pólizas que tengan su Primer Pago en un Periodo Anterior (Inmediatas y Diferidas)
        'vlSql = vlSql & " AND A.FEC_INIPAGOPEN < '" & Format(iFecIniPag, "yyyymmdd") & "'"
        vlSql = vlSql & " AND (a.fec_pripago < '" & Format(iFecIniPag, "yyyymmdd") & "'"
        vlSql = vlSql & " OR (a.num_mesdif > 0 AND a.fec_pripago = '" & Format(iFecIniPag, "yyyymmdd") & "'))"
        vlSql = vlSql & " AND A.NUM_POLIZA = B.NUM_POLIZA"
        vlSql = vlSql & " AND '" & Format(iFecIniPag, "yyyymmdd") & "'"
        vlSql = vlSql & " BETWEEN B.FEC_INIHABDES AND B.FEC_TERHABDES"
        vlSql = vlSql & " AND B.COD_CONHABDES = C.COD_CONHABDES"
        vlSql = vlSql & " AND C.COD_MODORIGEN = '" & iModulo & "'"
        
        If iModalidad = 2 Then
            vlSql = vlSql & ")"
        Else
            If vgTipoBase = "ORACLE" Then
                vlSql = vlSql & " ORDER BY NUM_POLIZA, NUM_ENDOSO"
            Else
                vlSql = vlSql & " ORDER BY A.NUM_POLIZA, A.NUM_ENDOSO"
            End If
        End If
    End Select

flObtieneQuery = vlSql
End Function

Function flProcesoAsigFam()
   'HQR 17/05/2005
   Dim vlSql As String
   Dim vlPago As String
   'FIN HQR 17/05/2005
   vlArchivo = strRpt & "PP_Rpt_ConAsigFam.rpt"   '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
       MsgBox "Archivo de Reporte de Póliza no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
       Screen.MousePointer = 0
       Exit Function
   End If
      
   Screen.MousePointer = 11
   
   vlSql = "delete from PP_TTMP_CONASIGFAM WHERE COD_USUARIO ='" & vgUsuario & "'"
   vgConexionBD.Execute vlSql
   
   flCargaTemporal
   
   Call fgVigenciaQuiebra(DateSerial(Txt_Anno, Txt_Mes, 1))
   
   vgQuery = ""
   vgQuery = "{PP_TTMP_CONASIGFAM.COD_USUARIO} = '" & Trim(vgUsuario) & "'"
    
   Rpt_Calculo.Reset
   Rpt_Calculo.WindowState = crptMaximized
   Rpt_Calculo.ReportFileName = vlArchivo
   Rpt_Calculo.Connect = vgRutaDataBase
   'Rpt_Calculo.SelectionFormula = ""
   Rpt_Calculo.SelectionFormula = vgQuery
   
   Rpt_Calculo.Formulas(0) = ""
   Rpt_Calculo.Formulas(1) = ""
   Rpt_Calculo.Formulas(2) = ""
   Rpt_Calculo.Formulas(3) = ""
   Rpt_Calculo.Formulas(4) = ""
   
   Rpt_Calculo.Formulas(0) = "NombreCompania='" & vgNombreCompania & "'"
   Rpt_Calculo.Formulas(1) = "NombreSistema= '" & vgNombreSistema & "'"
   Rpt_Calculo.Formulas(2) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
   vlPago = Txt_Mes.Text + "-" + Txt_Anno.Text
   Rpt_Calculo.Formulas(3) = "Periodo = '" & vlPago & "'"
   Rpt_Calculo.Formulas(4) = "GlosaQuiebra = '" & vgGlosaQuiebra & "'"
   
   Rpt_Calculo.SubreportToChange = ""
   Rpt_Calculo.Destination = crptToWindow
   Rpt_Calculo.WindowState = crptMaximized
   Rpt_Calculo.WindowTitle = "Informe de Control Asignación Familiar"
   Rpt_Calculo.SubreportToChange = "PP_Rpt_ConAsigFamTotal.rpt"
   Rpt_Calculo.SelectionFormula = ""
   Rpt_Calculo.Connect = vgRutaDataBase
   Rpt_Calculo.Action = 1

   
   Screen.MousePointer = 0


End Function

Function flCargaTemporal()
Dim vlMtoCodHabDes As Double 'HQR 17/05/2005
On Error GoTo Err_Carga

   vlNumPerPago = Txt_Anno + Txt_Mes
   
   vgSql = ""
   vgSql = "SELECT NUM_PERPAGO,NUM_POLIZA,NUM_ORDEN,COD_CONHABDES, RUT_RECEPTOR,MTO_CONHABDES,COD_TIPRECEPTOR"
   vgSql = vgSql & " FROM PP_TTMP_CONPAGOPEN"
   vgSql = vgSql & " WHERE COD_TIPOMOD = '" & clCodAF & "' AND "
   'I--- ABV 11/05/2005 ---
   'vgSql = vgSql & " NUM_PERPAGO = '" & vlNumPerPago & "' AND"
   'F--- ABV 11/05/2005 ---
   vgSql = vgSql & " COD_CONHABDES IN " & clCodHabDesAF & " AND "
   vgSql = vgSql & " COD_TIPRECEPTOR <> 'R' "
   'Ordenado por concepto, para primero insertar el 08 de A.F y despues
   'insertar o modificar los 09 de retroactivo
   vgSql = vgSql & " ORDER BY num_poliza,num_orden,cod_conhabdes "
   Set vlRegistro = vgConexionBD.Execute(vgSql)
   If Not vlRegistro.EOF Then
      While Not vlRegistro.EOF
        vlNumPoliza = (vlRegistro!num_poliza)
        vlNumOrden = (vlRegistro!Num_Orden)
        vlRutReceptor = (vlRegistro!Rut_Receptor)
        
        If vlNumPoliza = "0000020001" Then
            vlNumPoliza = Trim(vlNumPoliza)
        End If
        
        'I--- ABV 11/05/2005 ---
        vlPerPagoAsigFam = vlRegistro!Num_PerPago
        'F--- ABV 11/05/2005 ---
        
        'Calcula último día del mes
        vlAnno = CInt(Mid(vlPerPagoAsigFam, 1, 4))
        vlMes = CInt(Mid(vlPerPagoAsigFam, 5, 2))
        vlDia = 1
        vlFechaTermino = Format(DateSerial(vlAnno, vlMes + 1, vlDia - 1), "yyyymmdd")
        
        'Calcula Primer día del mes
        vlAnno = CInt(Mid(vlPerPagoAsigFam, 1, 4))
        vlMes = CInt(Mid(vlPerPagoAsigFam, 5, 2))
        vlDia = 1
        vlFechaInicio = Format(DateSerial(vlAnno, vlMes, vlDia), "yyyymmdd")
        
        Call fgDevuelveDigito(vlRutReceptor, vlDgvReceptor)
             
        vlCodHabDes = (vlRegistro!Cod_ConHabDes)
        vlMtoCodHabDes = (vlRegistro!Mto_ConHabDes)
        vlCodTipReceptor = (vlRegistro!Cod_TipReceptor)
            
        If vlCodHabDes = "09" Or vlCodHabDes = "30" Then
           vlNumOrdenCar = vlNumOrden
           vlMontoCarga = 0
           vlFecVencimiento = ""
           vgSql = ""
           vgSql = "SELECT b.COD_SITINV,b.COD_PAR,b.RUT_BEN,b.DGV_BEN,t.GLS_ELEMENTO,b.NUM_ENDOSO "
           vgSql = vgSql & "FROM PP_TMAE_BEN b,MA_TPAR_TABCOD t WHERE "
           vgSql = vgSql & "b.NUM_POLIZA = '" & vlNumPoliza & "' AND "
           vgSql = vgSql & "b.NUM_ORDEN = '" & vlNumOrden & "' AND "
           vgSql = vgSql & "t.COD_TABLA = 'PA' AND "
           vgSql = vgSql & "b.COD_PAR = t.COD_ELEMENTO "
           vgSql = vgSql & "ORDER BY b.NUM_ENDOSO DESC"
           Set vlRegistro5 = vgConexionBD.Execute(vgSql)
           If Not vlRegistro5.EOF Then
              vlCodPar = (vlRegistro5!Cod_Par)
              vlRutBen = (vlRegistro5!Rut_Ben)
              vlDigito = (vlRegistro5!Dgv_Ben)
              vlNumEndoso = (vlRegistro5!num_endoso)
              vlCodSitInv = (vlRegistro5!Cod_SitInv)
              vlDescripcion = (vlRegistro5!GLS_ELEMENTO)
           End If
           'Asignacion Familiar RetroActiva
           If vlCodHabDes = "09" Then
                'Buscar numero de reliquidacion del pago retroactivo
                vgSql = ""
                vgSql = "SELECT h.num_reliq, h.mto_cuota, h.num_cuotas "
                vgSql = vgSql & "FROM pp_tmae_habdes h "
                vgSql = vgSql & "WHERE h.num_poliza = '" & Trim(vlNumPoliza) & "' AND "
                vgSql = vgSql & "h.num_orden = " & Str(vlNumOrden) & " AND "
                vgSql = vgSql & "h.cod_conhabdes = '" & Trim(vlCodHabDes) & "' AND "
                vgSql = vgSql & "h.fec_inihabdes <= '" & Trim(vlFechaInicio) & "' AND "
                vgSql = vgSql & "h.fec_terhabdes >= '" & Trim(vlFechaTermino) & "' "
                Set vgRegistro = vgConexionBD.Execute(vgSql)
                If Not vgRegistro.EOF Then
                    If Not IsNull(vgRegistro!num_reliq) Then
                        vlNumReliq = (vgRegistro!num_reliq)
                        vlMonto = 0
                        vlMtoRetro = 0
                        'Seleccionar todas las cargas con pago retroactivo
                        'segun numero de reliquidacion
                        vgSql = ""
                        vgSql = "SELECT r.num_orden, SUM(r.mto_conhabdes)/" & Str(vgRegistro!Num_Cuotas) & " as monto "
                        vgSql = vgSql & "FROM pp_tmae_detcalcreliq r "
                        vgSql = vgSql & "WHERE r.num_reliq = " & Str(vlNumReliq) & " "
                        vgSql = vgSql & "GROUP BY r.num_orden,r.mto_conhabdes "
                        Set vgRegistro = vgConexionBD.Execute(vgSql)
                        If Not vgRegistro.EOF Then
                            While Not vgRegistro.EOF
                                vlNumOrdenCar = (vgRegistro!Num_Orden)
                                'PENDIENTE: Transformacion de monto de acuerdo a la moneda
                                'Obtener Parentesco de la carga con pago retroactivo
                                Call flBuscarParentescoRetro
                                'Obtener Situacion de Invalidez de la carga
                                Call flBuscarSitInvRetro
                                'Variable usada para modificacion
                                vlMonto = Format((vgRegistro!monto), "###,##0")
                                'Variable usada para insercion
                                vlMtoRetro = Format((vgRegistro!monto), "###,##0")
                                
                                'Confirmar existencia del registro a ingresar
                                vgSql = ""
                                vgSql = "SELECT cod_usuario FROM pp_ttmp_conasigfam "
                                vgSql = vgSql & "WHERE cod_usuario = '" & vgUsuario & "' AND "
                                vgSql = vgSql & "num_perpago = '" & vlPerPagoAsigFam & "' AND "
                                vgSql = vgSql & "num_poliza = '" & vlNumPoliza & "' AND "
'                                vgSql = vgSql & "num_ordenrec = '" & vlNumOrden & "' AND "
                                vgSql = vgSql & "num_orden = '" & vlNumOrdenCar & "' "
                                Set vlRegistro6 = vgConexionBD.Execute(vgSql)
                                If Not vlRegistro6.EOF Then
                                   Call flModificaTabla(vlCodHabDes, vlMonto)
                                Else
                                   Call flInsertaTabla
                                End If
                            
                                vgRegistro.MoveNext
                            Wend
                        End If
                    End If
                End If
              
           Else
             If vlCodHabDes = "30" Then
                vlMonto = vlMtoCodHabDes
                vlMtoReintegro = vlMtoCodHabDes
                
                vgSql = ""
                'vgSql = "SELECT cod_usuario FROM pp_tmae_conasigfam "
                vgSql = "SELECT cod_usuario FROM pp_ttmp_conasigfam "
                vgSql = vgSql & " WHERE "
                vgSql = vgSql & " cod_usuario = '" & vgUsuario & "' AND"
                vgSql = vgSql & " num_perpago = '" & vlPerPagoAsigFam & "' AND"
                vgSql = vgSql & " num_poliza = '" & vlNumPoliza & "' AND"
                vgSql = vgSql & " num_ordenrec = " & vlNumOrden & " AND"
                vgSql = vgSql & " num_orden = " & vlNumOrdenCar & ""
                Set vlRegistro6 = vgConexionBD.Execute(vgSql)
                If Not vlRegistro6.EOF Then
                   Call flModificaTabla(vlCodHabDes, vlMonto)
                Else
                   Call flInsertaTabla
                End If
                
             End If
           End If
       End If
           
       If vlCodHabDes = "08" Then
         If vlNumPoliza = "0000020114" Then
         vlNumPoliza = vlNumPoliza
        End If
              
          vgSql = ""
          vgSql = "SELECT l.NUM_ENDOSO,a.MTO_CARGA,a.NUM_ORDENCAR FROM PP_TTMP_CONLIQPAGOPEN l,"
          vgSql = vgSql & " PP_TTMP_CONPAGOASIG a WHERE "
          vgSql = vgSql & "l.COD_TIPOMOD = '" & clCodAF & "' AND "
          vgSql = vgSql & "l.COD_TIPOMOD = a.COD_TIPOMOD AND "
          'I--- ABV 11/05/2005 ---
          'vgSql = vgSql & "l.NUM_PERPAGO = '" & vlNumPerPago & "' AND "
          vgSql = vgSql & "l.NUM_PERPAGO = '" & vlPerPagoAsigFam & "' AND "
          'F--- ABV 11/05/2005 ---
          vgSql = vgSql & "l.NUM_PERPAGO = a.NUM_PERPAGO AND "
          vgSql = vgSql & "l.NUM_POLIZA = '" & vlNumPoliza & "' AND "
          vgSql = vgSql & "l.NUM_POLIZA = a.NUM_POLIZA AND "
          vgSql = vgSql & "l.NUM_ORDEN = '" & vlNumOrden & "' AND "
          vgSql = vgSql & "l.NUM_ORDEN = a.NUM_ORDEN AND "
          vgSql = vgSql & "l.RUT_RECEPTOR = '" & vlRutReceptor & "' and "
          vgSql = vgSql & "l.RUT_RECEPTOR = a.RUT_RECEPTOR and "
          'JANM
          vgSql = vgSql & "l.COD_TIPRECEPTOR = '" & vlCodTipReceptor & "' and "
          vgSql = vgSql & "l.COD_TIPRECEPTOR = a.COD_TIPRECEPTOR "
          'JANM
          Set vlRegistro1 = vgConexionBD.Execute(vgSql)
          If Not vlRegistro1.EOF Then
             While Not vlRegistro1.EOF
               vlNumOrdenCar = (vlRegistro1!Num_OrdenCar)
               vlMontoCarga = (vlRegistro1!Mto_Carga)
               vlNumEndoso = (vlRegistro1!num_endoso)
               'Tabla de los beneficiarios
               If vlNumOrdenCar < 50 Then
                  vgSql = ""
                  vgSql = "SELECT b.COD_PAR,b.RUT_BEN,b.DGV_BEN,t.GLS_ELEMENTO  "
                  vgSql = vgSql & "FROM PP_TMAE_BEN b,MA_TPAR_TABCOD t WHERE "
                  vgSql = vgSql & "b.NUM_POLIZA = '" & vlNumPoliza & "' AND "
                  vgSql = vgSql & "b.NUM_ORDEN = '" & vlNumOrdenCar & "' AND "
                  vgSql = vgSql & "b.NUM_ENDOSO = " & vlNumEndoso & " AND "
                  vgSql = vgSql & "t.COD_TABLA = 'PA' AND "
                  vgSql = vgSql & "b.COD_PAR = t.COD_ELEMENTO "
                  Set vlRegistro3 = vgConexionBD.Execute(vgSql)
                  If Not vlRegistro3.EOF Then
                     vlCodPar = (vlRegistro3!Cod_Par)
                     vlRutBen = (vlRegistro3!Rut_Ben)
                     vlDigito = (vlRegistro3!Dgv_Ben)
                     vlDescripcion = (vlRegistro3!GLS_ELEMENTO)
                  End If
               Else
                If vlNumOrdenCar >= 50 Then
                   vgSql = ""
                   vgSql = "SELECT b.COD_ASCDES,b.RUT_BEN,b.DGV_BEN,t.GLS_ELEMENTO "
                   vgSql = vgSql & "FROM PP_TMAE_NOBEN b,MA_TPAR_TABCOD t WHERE "
                   vgSql = vgSql & "b.NUM_POLIZA = '" & vlNumPoliza & "' AND "
                   vgSql = vgSql & "b.NUM_ORDEN = " & vlNumOrdenCar & " AND "
                   vgSql = vgSql & "t.COD_TABLA = 'PNB' AND "
                   vgSql = vgSql & "b.COD_ASCDES = t.COD_ELEMENTO"
                  ' AND "
                  ' vgSql = vgSql & "b.NUM_ORDENREC = " & vlNumOrden & ""
                   Set vlRegistro3 = vgConexionBD.Execute(vgSql)
                   If Not vlRegistro3.EOF Then
                      vlCodPar = (vlRegistro3!COD_ASCDES)
                      vlRutBen = (vlRegistro3!Rut_Ben)
                      vlDigito = (vlRegistro3!Dgv_Ben)
                      vlDescripcion = (vlRegistro3!GLS_ELEMENTO)
                   End If
                End If
               End If
               vgSql = ""
               vgSql = "SELECT COD_SITINV, FEC_TERACTIVA FROM PP_TMAE_ASIGFAM WHERE "
               vgSql = vgSql & "NUM_POLIZA = '" & vlNumPoliza & "' AND "
               vgSql = vgSql & "NUM_ORDEN = '" & vlNumOrdenCar & "' AND "
               vgSql = vgSql & "NUM_ORDENREC = '" & vlNumOrden & "' AND "
               If vgTipoBase = "ORACLE" Then
                    'abv 11/05/2005 existen dos periodos distintos (en regimen y primer pago)
                    'vgSql = vgSql & "SUBSTR(FEC_INIACTIVA,1,6) <= '" & vlNumPerPago & "' AND "
                    'vgSql = vgSql & "SUBSTR(FEC_TERACTIVA,1,6) >= '" & vlNumPerPago & "'"
                    'hqr 05/10/2005 por las reliquidaciones
                    'vgSql = vgSql & "SUBSTR(FEC_INIACTIVA,1,6) <= '" & vlPerPagoAsigFam & "' AND "
                    'vgSql = vgSql & "SUBSTR(FEC_TERACTIVA,1,6) >= '" & vlPerPagoAsigFam & "'"
                    vgSql = vgSql & "SUBSTR(FEC_EFECTO,1,6) <= '" & vlPerPagoAsigFam & "' AND "
                    vgSql = vgSql & "SUBSTR(FEC_SUSPENSION,1,6) >= '" & vlPerPagoAsigFam & "'"
               Else
                    'abv 11/05/2005 existen dos periodos distintos (en regimen y primer pago)
                    'vgSql = vgSql & "SUBSTRING(FEC_INIACTIVA,1,6) <= '" & vlNumPerPago & "' AND "
                    'vgSql = vgSql & "SUBSTRING(FEC_TERACTIVA,1,6) >= '" & vlNumPerPago & "'"
                    'hqr 05/10/2005 por las reliquidaciones
                    'vgSql = vgSql & "SUBSTRING(FEC_INIACTIVA,1,6) <= '" & vlPerPagoAsigFam & "' AND "
                    'vgSql = vgSql & "SUBSTRING(FEC_TERACTIVA,1,6) >= '" & vlPerPagoAsigFam & "'"
                    vgSql = vgSql & "SUBSTRING(FEC_EFECTO,1,6) <= '" & vlPerPagoAsigFam & "' AND "
                    vgSql = vgSql & "SUBSTRING(FEC_SUSPENSION,1,6) >= '" & vlPerPagoAsigFam & "'"
               End If
               Set vlRegistro4 = vgConexionBD.Execute(vgSql)
               If Not vlRegistro4.EOF Then
                  vlCodSitInv = (vlRegistro4!Cod_SitInv)
                  vlFecVencimiento = (vlRegistro4!FEC_TERACTIVA)
               End If
               vlMtoRetro = 0
               vlMtoReintegro = 0
               vlMto_total = 0
                        
               Call flInsertaTabla
               
              vlRegistro1.MoveNext
             Wend
           End If
       End If
       vlRegistro.MoveNext
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

Function flBuscarParentescoRetro()
On Error GoTo Err_flBuscarParentescoRetro

    'Si el Numero de Orden corresponde a un Beneficiario
    If vlNumOrdenCar < 50 Then
       vgSql = ""
       vgSql = "SELECT b.cod_par,b.rut_ben,b.dgv_ben,t.gls_elemento "
       vgSql = vgSql & "FROM pp_tmae_ben b,ma_tpar_tabcod t WHERE "
       vgSql = vgSql & "b.num_poliza = '" & Trim(vlNumPoliza) & "' AND "
       vgSql = vgSql & "b.num_orden = " & Str(vlNumOrdenCar) & " AND "
       vgSql = vgSql & "b.num_endoso = " & Str(vlNumEndoso) & " AND "
       vgSql = vgSql & "t.cod_tabla = 'PA' AND "
       vgSql = vgSql & "b.cod_par = t.cod_elemento "
       Set vlRegistro3 = vgConexionBD.Execute(vgSql)
       If Not vlRegistro3.EOF Then
          vlCodPar = (vlRegistro3!Cod_Par)
          vlRutBen = (vlRegistro3!Rut_Ben)
          vlDigito = (vlRegistro3!Dgv_Ben)
          vlDescripcion = (vlRegistro3!GLS_ELEMENTO)
       End If
    Else
    'Si el Numero de Orden corresponde a un NO Beneficiario
     If vlNumOrdenCar >= 50 Then
        vgSql = ""
        vgSql = "SELECT b.cod_ascdes,b.rut_ben,b.dgv_ben,t.gls_elemento "
        vgSql = vgSql & "FROM pp_tmae_noben b,ma_tpar_tabcod t "
        vgSql = vgSql & "WHERE b.num_poliza = '" & vlNumPoliza & "' AND "
        vgSql = vgSql & "b.num_orden = " & vlNumOrdenCar & " AND "
        vgSql = vgSql & "t.cod_tabla = 'PNB' AND "
        vgSql = vgSql & "b.cod_ascdes = t.cod_elemento "
        Set vlRegistro3 = vgConexionBD.Execute(vgSql)
        If Not vlRegistro3.EOF Then
           vlCodPar = (vlRegistro3!COD_ASCDES)
           vlRutBen = (vlRegistro3!Rut_Ben)
           vlDigito = (vlRegistro3!Dgv_Ben)
           vlDescripcion = (vlRegistro3!GLS_ELEMENTO)
        End If
     End If
    End If

Exit Function
Err_flBuscarParentescoRetro:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function flBuscarSitInvRetro()
On Error GoTo Err_flBuscarSitInvRetro

    vlCodSitInv = "N"
    vlFecVencimiento = ""
    vgSql = ""
    vgSql = "SELECT a.cod_sitinv, a.fec_teractiva "
    vgSql = vgSql & "FROM pp_tmae_asigfam a WHERE "
    vgSql = vgSql & "a.num_poliza = '" & Trim(vlNumPoliza) & "' AND "
    vgSql = vgSql & "a.num_orden = " & Str(vlNumOrdenCar) & " AND "
    vgSql = vgSql & "a.num_ordenrec = " & Str(vlNumOrden) & " AND "
    If vgTipoBase = "ORACLE" Then
         vgSql = vgSql & "SUBSTR(FEC_INIACTIVA,1,6) <= '" & Trim(vlPerPagoAsigFam) & "' AND "
         vgSql = vgSql & "SUBSTR(FEC_TERACTIVA,1,6) >= '" & Trim(vlPerPagoAsigFam) & "'"
    Else
         vgSql = vgSql & "SUBSTRING(FEC_INIACTIVA,1,6) <= '" & Trim(vlPerPagoAsigFam) & "' AND "
         vgSql = vgSql & "SUBSTRING(FEC_TERACTIVA,1,6) >= '" & Trim(vlPerPagoAsigFam) & "'"
    End If
    Set vlRegistro4 = vgConexionBD.Execute(vgSql)
    If Not vlRegistro4.EOF Then
       vlCodSitInv = (vlRegistro4!Cod_SitInv)
       vlFecVencimiento = (vlRegistro4!FEC_TERACTIVA)
    End If

Exit Function
Err_flBuscarSitInvRetro:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function flInsertaTabla()
    Dim Sql As String 'HQR 17/05/2005
On Error GoTo Errores

    Sql = ""
    Sql = "INSERT INTO PP_TTMP_CONASIGFAM ("
    Sql = Sql & "COD_USUARIO,NUM_PERPAGO,NUM_POLIZA,NUM_ORDENREC,NUM_ORDEN,"
    Sql = Sql & "RUT_BEN,DGV_BEN,MTO_CARGA,COD_PAR,"
    Sql = Sql & "COD_SITINV,MTO_RETRO,MTO_REINTEGRO,FEC_VENCARGA,MTO_TOTAL,GLS_ELEMENTO,"
    Sql = Sql & "RUT_BENREC,DGV_BENREC"
    Sql = Sql & " "
    Sql = Sql & ") values ("
    Sql = Sql & "'" & vgUsuario & "',"
    'I--- ABV 11/05/2005 ---
    'Sql = Sql & "'" & Trim(vlNumPerPago) & "',"
    Sql = Sql & "'" & Trim(vlPerPagoAsigFam) & "',"
    'F--- ABV 11/05/2005 ---
    Sql = Sql & "'" & vlNumPoliza & "',"
    Sql = Sql & "" & vlNumOrden & ","
    Sql = Sql & "" & vlNumOrdenCar & ","
    Sql = Sql & "" & vlRutBen & ","
    Sql = Sql & "'" & vlDigito & "',"
    Sql = Sql & "" & Str(vlMontoCarga) & ","
    Sql = Sql & "'" & vlCodPar & "',"
    Sql = Sql & "'" & vlCodSitInv & "',"
    Sql = Sql & "" & Str(vlMtoRetro) & ","
    Sql = Sql & "" & Str(vlMtoReintegro) & ","
    If vlFecVencimiento = "" Then
       Sql = Sql & " NULL,"
    Else
       Sql = Sql & "'" & vlFecVencimiento & "',"
    End If
    Sql = Sql & "" & Str(vlMto_total) & ","
    Sql = Sql & "'" & Trim(vlDescripcion) & "',"
    Sql = Sql & "" & vlRutReceptor & ","
    Sql = Sql & "'" & vlDgvReceptor & "'"
    Sql = Sql & ")"
    vgConexionBD.Execute (Sql)
    'Nota: Valor de monto total es calculado (sumado) en el reporte
   
Exit Function
Errores:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flProcesoCCAF()
    Dim vlSql As String 'HQR 17/05/2005
On Error GoTo Err_Imprimir
   
   Screen.MousePointer = 11
   
   vlArchivo = strRpt & "PP_Rpt_CCAFControl.rpt"   '\Reportes
   
   If Not fgExiste(vlArchivo) Then     ', vbNormal
       MsgBox "Archivo de Reporte de Póliza no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
       Screen.MousePointer = 0
       Exit Function
   End If
   
   vlFecha = Trim(Txt_Mes) + "-" + Trim(Txt_Anno)
   
   vlSql = "delete from PP_TTMP_CONCCAF WHERE COD_USUARIO ='" & vgUsuario & "'"
   vgConexionBD.Execute vlSql
   
   Call flCargaTemporalCCAF
   
   Call fgVigenciaQuiebra(DateSerial(Txt_Anno, Txt_Mes, 1))
   
    vgQuery = ""
'   vgQuery = "{PP_TTMP_CONCCAF.FEC_INIACTIVA}<= '" & vlFecha & "' AND "
'   vgQuery = vgQuery & "{PP_TMAE_ASIGFAM.FEC_TERACTIVA}>= '" & vlFecha & "'"

    Rpt_Calculo.Reset
    Rpt_Calculo.WindowState = crptMaximized
    Rpt_Calculo.ReportFileName = vlArchivo
    Rpt_Calculo.Connect = vgRutaDataBase
    Rpt_Calculo.SelectionFormula = vgQuery
    
    Rpt_Calculo.Formulas(0) = ""
    Rpt_Calculo.Formulas(1) = ""
    Rpt_Calculo.Formulas(2) = ""
    Rpt_Calculo.Formulas(3) = ""
    Rpt_Calculo.Formulas(4) = ""
   
    
    Rpt_Calculo.Formulas(0) = "NombreCompania ='" & vgNombreCompania & "'"
    Rpt_Calculo.Formulas(1) = "NombreSistema = '" & vgNombreSistema & "'"
    Rpt_Calculo.Formulas(2) = "NombreSubSistema = '" & vgNombreSubSistema & "'"
    Rpt_Calculo.Formulas(3) = "Fecha = '" & vlFecha & "'"
    Rpt_Calculo.Formulas(4) = "GlosaQuiebra = '" & vgGlosaQuiebra & "'"
   
    Rpt_Calculo.SubreportToChange = ""
    Rpt_Calculo.Destination = crptToWindow
    Rpt_Calculo.WindowState = crptMaximized
    Rpt_Calculo.WindowTitle = "Informe de Control Cajas de Compensación"
    Rpt_Calculo.SelectionFormula = ""
    Rpt_Calculo.Action = 1
    
    Screen.MousePointer = 0

Exit Function
Err_Imprimir:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flCargaTemporalCCAF()
    Dim vlSql As String 'HQR 17/05/2005
On Error GoTo Err_Carga

   vlNumPerPago = Txt_Anno + Txt_Mes
   vlSql = ""
   vlSql = "SELECT DISTINCT NUM_POLIZA,NUM_ORDEN "
   vlSql = vlSql & "FROM PP_TTMP_CONPAGOPENCCAF "
   vlSql = vlSql & "WHERE "
   vlSql = vlSql & "NUM_PERPAGO = '" & vlNumPerPago & "' AND "
   vlSql = vlSql & "COD_TIPOMOD = '" & clCodCCAF & "'"
   Set vgRs = vgConexionBD.Execute(vlSql)
   Do While Not vgRs.EOF
        vlNumPoliza = vgRs!num_poliza
'        If vlNumPoliza = "0000020010" Then
'            vlNumPoliza = vgRs!Num_Poliza
'        End If
        vlNumOrden = vgRs!Num_Orden
        
        vlSql = ""
        vlSql = "SELECT DISTINCT COD_CAJACOMPEN FROM PP_TTMP_CONPAGOPENCCAF "
        vlSql = vlSql & "WHERE "
        vlSql = vlSql & "NUM_POLIZA ='" & vlNumPoliza & "' AND "
        vlSql = vlSql & "NUM_PERPAGO = '" & vlNumPerPago & "' AND "
        vlSql = vlSql & "COD_TIPOMOD = '" & clCodCCAF & "' AND "
        vlSql = vlSql & "NUM_ORDEN = " & vlNumOrden & " "
        Set vgRs2 = vgConexionBD.Execute(vlSql)
        Do While Not vgRs2.EOF
            vlCCAF = Trim(vgRs2!Cod_CajaCompen)
            'saca suma aporte
            vlSql = ""
            vlSql = "SELECT SUM(MTO_CONHABDES) AS MONTO "
            vlSql = vlSql & "FROM PP_TTMP_CONPAGOPENCCAF "
            vlSql = vlSql & "WHERE NUM_PERPAGO = '" & Trim(vlNumPerPago) & "' AND "
            vlSql = vlSql & "NUM_POLIZA = '" & Trim(vlNumPoliza) & "' AND "
            vlSql = vlSql & "NUM_ORDEN = " & vlNumOrden & " AND "
            vlSql = vlSql & "COD_TIPOMOD = '" & clCodCCAF & "' AND "
            vlSql = vlSql & "COD_CONHABDES = '25' AND "
            vlSql = vlSql & "COD_CAJACOMPEN = '" & Trim(vlCCAF) & "' "
            Set vgRs3 = vgConexionBD.Execute(vlSql)
            If Not vgRs3.EOF Then
                If Not IsNull(vgRs3!monto) Then
                    vlMontoAporte = vgRs3!monto
                Else
                    vlMontoAporte = "0"
                End If
            End If
            vgRs3.Close
            
            'saca suma credito
            vlSql = "SELECT SUM(MTO_CONHABDES) AS MONTO"
            vlSql = vlSql & " FROM PP_TTMP_CONPAGOPENCCAF "
            vlSql = vlSql & "WHERE NUM_POLIZA='" & vlNumPoliza & "' AND "
            vlSql = vlSql & "NUM_PERPAGO= '" & vlNumPerPago & "' AND "
            vlSql = vlSql & "COD_TIPOMOD='" & clCodCCAF & "' AND "
            vlSql = vlSql & "COD_CONHABDES='26' AND "
            vlSql = vlSql & "COD_CAJACOMPEN='" & vlCCAF & "' AND "
            vlSql = vlSql & "NUM_ORDEN=" & vlNumOrden & ""
            Set vgRs3 = vgConexionBD.Execute(vlSql)
            If Not vgRs3.EOF Then
                If Not IsNull(vgRs3!monto) Then
                    vlMontoCredito = vgRs3!monto
                Else
                    vlMontoCredito = "0"
                End If
            End If
            vgRs3.Close
            
            'saca suma otros descuentos
            vlSql = "SELECT SUM(MTO_CONHABDES) AS MONTO "
            vlSql = vlSql & "FROM PP_TTMP_CONPAGOPENCCAF "
            vlSql = vlSql & "WHERE NUM_POLIZA='" & vlNumPoliza & "' AND "
            vlSql = vlSql & "NUM_PERPAGO= '" & vlNumPerPago & "' AND "
            vlSql = vlSql & "COD_TIPOMOD='" & clCodCCAF & "' AND "
            vlSql = vlSql & "COD_CONHABDES<>'25' AND "
            vlSql = vlSql & "COD_CONHABDES<>'26' AND "
            vlSql = vlSql & "COD_CAJACOMPEN='" & vlCCAF & "' AND "
            vlSql = vlSql & "NUM_ORDEN=" & vlNumOrden & ""
            Set vgRs3 = vgConexionBD.Execute(vlSql)
            If Not vgRs3.EOF Then
                If Not IsNull(vgRs3!monto) Then
                    vlMontoOtros = vgRs3!monto
                Else
                    vlMontoOtros = "0"
                End If
                
            End If
            vgRs3.Close
               
            'saca el ultimo endoso
            vlSql = ""
            vlSql = "SELECT NUM_ENDOSO FROM PP_TTMP_CONLIQPAGOPEN "
            vlSql = vlSql & "WHERE NUM_POLIZA='" & vlNumPoliza & "' AND "
            vlSql = vlSql & "NUM_PERPAGO='" & vlNumPerPago & "' AND "
            vlSql = vlSql & "NUM_ORDEN= " & vlNumOrden & " AND "
            vlSql = vlSql & "COD_TIPOMOD='" & clCodCCAF & "' "
            vlSql = vlSql & " ORDER BY NUM_ENDOSO DESC"
            Set vgRs3 = vgConexionBD.Execute(vlSql)
            If Not vgRs.EOF Then
                vlNumEndoso = vgRs3!num_endoso
            End If
            vgRs3.Close
            
            'saca el nro de rut del afiliado
            vlSql = ""
            vlSql = "SELECT RUT_BEN,DGV_BEN FROM PP_TMAE_BEN "
            vlSql = vlSql & "WHERE NUM_POLIZA='" & vlNumPoliza & "' AND "
            vlSql = vlSql & "NUM_ORDEN= " & vlNumOrden & " AND "
            vlSql = vlSql & "NUM_ENDOSO=" & vlNumEndoso & ""
            Set vgRs3 = vgConexionBD.Execute(vlSql)
            If Not vgRs3.EOF Then
                vlRut = vgRs3!Rut_Ben
                vlDgv = vgRs3!Dgv_Ben
            End If
            vgRs3.Close
            
            vlSql = "insert into PP_TTMP_CONCCAF ("
            vlSql = vlSql & "COD_USUARIO,NUM_POLIZA,NUM_ORDEN,RUT_BEN,"
            vlSql = vlSql & "DGV_BEN,COD_CAJACOMPEN,MTO_APORTE,MTO_CREDITO,"
            vlSql = vlSql & "MTO_OTRO)VALUES("
            vlSql = vlSql & "'" & vgUsuario & "',"
            vlSql = vlSql & "'" & vlNumPoliza & "',"
            vlSql = vlSql & " " & Str(vlNumOrden) & ","
            vlSql = vlSql & " " & Str(vlRut) & ","
            vlSql = vlSql & "'" & vlDgv & "',"
            vlSql = vlSql & "'" & vlCCAF & "',"
            vlSql = vlSql & " " & Str(vlMontoAporte) & ","
            vlSql = vlSql & " " & Str(vlMontoCredito) & ","
            vlSql = vlSql & " " & Str(vlMontoOtros) & ")"
            vgConexionBD.Execute (vlSql)
            vgRs2.MoveNext
        Loop
        vgRs.MoveNext
    Loop

Exit Function
Err_Carga:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function
Function flProcesoRetJudicial()
    Dim vlSql As String 'HQR 17/05/2005
On Error GoTo Err_CmdImprimir
    
   vlPeriodo = Trim(Txt_Anno.Text + Txt_Mes.Text)
   
   'Buscar codigos para seleccionar Retenciones
'   Call flBuscaCodHabDes("RJC")

   vlSql = "DELETE FROM PP_TTMP_CONRETJUD WHERE cod_usuario = '" & vgUsuario & "'"
   vgConexionBD.Execute (vlSql)
   

   Call flCargaTablaTemporal
   
   'Buscar codigos para seleccionar Otros Haberes y Otros Descuentos
   vlTipoMov = clCodHab
   Call flBuscaCodOtrosHabDes(vlTipoMov)
   If vlCodOtrosHabDes <> "" Then
      Call flAgregarOtrosHabDes
   End If
   vlTipoMov = clCodDescto
   Call flBuscaCodOtrosHabDes(vlTipoMov)
   If vlCodOtrosHabDes <> "" Then
      Call flAgregarOtrosHabDes
   End If
                 
   Screen.MousePointer = 11
   
   vlArchivo = strRpt & "PP_Rpt_ControlRetJud.rpt"   '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
      MsgBox "Archivo de Reporte de Control de Retención Judicial no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
      Screen.MousePointer = 0
      Exit Function
   End If
   
   Call fgVigenciaQuiebra(DateSerial(Txt_Anno, Txt_Mes, 1))
   
'''   vlFechaInicio = Format(CDate(Trim(Txt_FechaIni.Text)), "yyyymmdd")
'''   vlFechaTermino = Format(CDate(Trim(Txt_FechaTer.Text)), "yyyymmdd")
'''
'''   vgQuery = ""
'''   vgQuery = vgQuery & "{PP_TMAE_RETJUDICIAL.fec_iniret}>= '" & Trim(vlFechaInicio) & "' AND "
'''   vgQuery = vgQuery & "{PP_TMAE_RETJUDICIAL.fec_iniret}<= '" & Trim(vlFechaTermino) & "' "

   vgQuery = ""
   vgQuery = "{PP_TTMP_CONRETJUD.COD_USUARIO} = '" & Trim(vgUsuario) & "'"
       
   Rpt_Calculo.Reset
   Rpt_Calculo.WindowState = crptMaximized
   Rpt_Calculo.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
   Rpt_Calculo.Connect = vgRutaDataBase       ' o App.Path & "\Nestle.mdb"
   Rpt_Calculo.SelectionFormula = vgQuery
'''   Rpt_General.SelectionFormula = vgQuery
      
'''   vgPalabra = Txt_FechaIni.Text & " - " & Txt_FechaTer.Text
   
   vgPalabra = Txt_Mes.Text + "-" + Txt_Anno.Text
   
   Rpt_Calculo.Formulas(0) = ""
   Rpt_Calculo.Formulas(1) = ""
   Rpt_Calculo.Formulas(2) = ""
   Rpt_Calculo.Formulas(3) = ""
   Rpt_Calculo.Formulas(4) = ""
   
   Rpt_Calculo.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
   Rpt_Calculo.Formulas(1) = "NombreSistema= '" & vgNombreSistema & "'"
   Rpt_Calculo.Formulas(2) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
   Rpt_Calculo.Formulas(3) = "Periodo = '" & vgPalabra & "'"
   Rpt_Calculo.Formulas(4) = "GlosaQuiebra = '" & vgGlosaQuiebra & "'"
   
   Rpt_Calculo.SubreportToChange = ""
   Rpt_Calculo.Destination = crptToWindow
   Rpt_Calculo.WindowState = crptMaximized
   Rpt_Calculo.WindowTitle = "Informe Control Retención Judicial"
   Rpt_Calculo.SelectionFormula = ""
   Rpt_Calculo.Action = 1
   Screen.MousePointer = 0
   
Exit Function
Err_CmdImprimir:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flCargaTablaTemporal()

On Error GoTo Err_flCargaTablaTemporal
     
'    select num_poliza,num_endoso,num_orden,cod_conhabdes,rut_receptor
'    from pp_ttmp_conpagopen p
'    where cod_conhabdes IN ('60','61','62')
'Seleccionar todos los detalles por concepto de RJC Cobro de Ret. Jud.
    
    vlNumCargas = 0
    vlMtoCargas = 0
    vlMtoHaber = 0
    vlMtoDescto = 0

    vgSql = ""
    vgSql = "SELECT num_poliza,num_orden,cod_conhabdes,"
    vgSql = vgSql & "mto_conhabdes,num_idenreceptor, cod_tipoidenreceptor,cod_tipreceptor, num_perpago "
    vgSql = vgSql & "FROM PP_TTMP_CONPAGOPEN "
    vgSql = vgSql & "WHERE cod_conhabdes IN " & vlCodHabDesRJC & " AND "
    vgSql = vgSql & "cod_tipomod = '" & clTipoRJ & "' "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
       vlCodTipReceptor = (vgRs!Cod_TipReceptor)
       
       While Not vgRs.EOF
       
             vlNumCargas = 0
             vlMtoCargas = 0
             vlMtoHaber = 0
             vlMtoDescto = 0
             vlPerPagoRetencion = vgRs!Num_PerPago
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
             vgSql = vgSql & "num_idenreceptor = '" & (vgRs!Num_IdenReceptor) & "' AND "
             vgSql = vgSql & "cod_tipoidenreceptor = " & (vgRs!Cod_TipoIdenReceptor) & " AND "
             vgSql = vgSql & "cod_tipret = '" & Trim(vlTipoRetJud) & "' AND "
'             vgSql = vgSql & "Cod_TipReceptor = '" & (vgRs!Cod_TipReceptor) & "' AND "
             If vgTipoBase = "ORACLE" Then
                'hqr 28/04/2005 existen dos periodos distintos (en regimen y primer pago)
                'vgSql = vgSql & "substr(fec_iniret,1,6) <= '" & vlPeriodo & "' AND "
                'vgSql = vgSql & "substr(fec_terret,1,6) >= '" & vlPeriodo & "' "
                vgSql = vgSql & "substr(fec_iniret,1,6) <= '" & vlPerPagoRetencion & "' AND "
                vgSql = vgSql & "substr(fec_terret,1,6) >= '" & vlPerPagoRetencion & "' "
             Else
                 'vgSql = vgSql & "substring(fec_iniret,1,6) <= '" & vlPeriodo & "' AND "
                 'vgSql = vgSql & "substring(fec_terret,1,6) >= '" & vlPeriodo & "' "
                 vgSql = vgSql & "substring(fec_iniret,1,6) <= '" & vlPerPagoRetencion & "' AND "
                 vgSql = vgSql & "substring(fec_terret,1,6) >= '" & vlPerPagoRetencion & "' "
             End If
             Set vgRegistro = vgConexionBD.Execute(vgSql)
             If Not vgRegistro.EOF Then
                vlNumRetencion = (vgRegistro!num_retencion)
                
                vlNumOrdenCar = 0
'                If vlTipoRetJud = clTipoRAF Then
'                   'Buscar Nùmero de Orden correspondiente a Conyuge
'                   vgSql = ""
'                   vgSql = "SELECT num_orden FROM PP_TMAE_DETRETENCION "
'                   vgSql = vgSql & "WHERE num_retencion = " & Str(vlNumRetencion) & " "
'                   Set vgRegistro = vgConexionBD.Execute(vgSql)
'                   vlNumOrdenCar = (vgRegistro!Num_Orden)
'                End If
                
'                If (vgRs!Cod_ConHabDes) = clCodHabDes62 Then
'                   vlNumCargas = 1
'                Else
'                    'Call flBuscaNumCargas
'                    vlNumCargas = 0
'                End If
                
''                If ((vgRs!Cod_ConHabDes) = clCodHabDes61) Or ((vgRs!Cod_ConHabDes) = clCodHabDes62) Then
''                   vlMtoCargas = (vgRs!Mto_ConHabDes)
''                End If
                
                vlMtoHaber = 0
                vlMtoDescto = 0
                vlMontoRetencion = (vgRs!Mto_ConHabDes)
               
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

Function flAgregarTablaTemporal()
'HQR 17/05/2005
Dim vlNumCargasAux As Integer
Dim vlMtoCargasAux As Double, vlMtoHaberAux As Double
Dim vlMtoDesctoAux As Double, vlMtoRetNetaAux As Double
Dim vlMtoRetJud As Double 'Para qué se usa???
'FIN HQR 17/05/2005
On Error GoTo Err_flAgregarTablaTemporal

    vgSql = ""
    vgSql = "SELECT * "
    vgSql = vgSql & "FROM PP_TTMP_CONRETJUD "
    vgSql = vgSql & "WHERE num_retencion = '" & Trim(vlNumRetencion) & "' "
    vgSql = vgSql & "AND num_perpago = '" & vlPerPagoRetencion & "'"
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If vgRegistro.EOF Then
       
       vgSql = ""
       vgSql = "INSERT INTO PP_TTMP_CONRETJUD "
       vgSql = vgSql & "(cod_usuario,num_retencion,num_perpago,num_cargas, "
       vgSql = vgSql & " mto_carga,mto_haber,mto_descto,mto_retneta "
       vgSql = vgSql & " ) VALUES ( "
       vgSql = vgSql & " '" & vgUsuario & "', "
       vgSql = vgSql & " '" & Trim(vlNumRetencion) & "' , "
       vgSql = vgSql & " '" & Trim(vlPerPagoRetencion) & "' , "
       vgSql = vgSql & " " & Str(vlNumCargas) & ", "
       vgSql = vgSql & " " & Str(vlMtoCargas) & ", "
       vgSql = vgSql & " " & Str(vlMtoHaber) & ", "
       vgSql = vgSql & " " & Str(vlMtoDescto) & ", "
       vgSql = vgSql & " " & Str(vlMontoRetencion) & ") "
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
        vgSql = vgSql & "num_cargas = " & Str(vlNumCargasAux) & ", "
        vgSql = vgSql & "mto_carga = " & Str(vlMtoCargasAux) & ", "
        vgSql = vgSql & "mto_retneta = " & Str(vlMtoRetNetaAux) & " "
        vgSql = vgSql & "WHERE num_retencion = '" & Trim(vlNumRetencion) & "' "
        vgConexionBD.Execute vgSql
        
'''        If vlTipoRetJud = clTipoRAF Then
'''           vgSql = ""
'''           vgSql = " UPDATE PP_TTMP_CONRETJUD SET "
'''           vgSql = vgSql & "mto_carga = " & Str(vlMtoCargas) & " "
'''           vgSql = vgSql & "WHERE num_retencion = '" & Trim(vlNumRetencion) & "' "
'''           vgConexionBD.Execute vgSql
'''        End If
        If vlTipoMov = clCodHab Then
            If vlMtoHaberAux = 0 Then
                vlMtoHaberAux = vlMtoHaberAux + vlMtoHabDes
            End If
           vgSql = ""
           vgSql = " UPDATE PP_TTMP_CONRETJUD SET "
           vgSql = vgSql & "mto_haber = " & Str(vlMtoHabDes) & " "
           vgSql = vgSql & "WHERE num_retencion = '" & Trim(vlNumRetencion) & "' "
           vgConexionBD.Execute vgSql
        End If
        If vlTipoMov = clCodDescto Then
            If vlMtoDesctoAux = 0 Then
                vlMtoDesctoAux = vlMtoDesctoAux + vlMtoHabDes
            End If
            vgSql = ""
            vgSql = " UPDATE PP_TTMP_CONRETJUD SET "
            vgSql = vgSql & "mto_descto = " & Str(vlMtoHabDes) & " "
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

Function flBuscaNumCargas()

On Error GoTo Err_flBuscaNumCargas
'Busca el número de cargas pagadas por concepto de Retención
'Judicial de Asignacion Familiar (Solo Hijos, sin Conyuge)

    vgSql = ""
    vgSql = "SELECT count(num_ordencar) as numcargas "
    vgSql = vgSql & "FROM PP_TTMP_CONPAGOASIG "
    vgSql = vgSql & "WHERE cod_tipomod = '" & Trim(clCodRJ) & "' AND "
    'vgSql = vgSql & "num_perpago = '" & Trim(vlPeriodo) & "' AND "
    vgSql = vgSql & "num_perpago = '" & Trim(vlPerPagoRetencion) & "' AND "
    vgSql = vgSql & "num_poliza = '" & Trim(vgRs!num_poliza) & "' AND "
    vgSql = vgSql & "num_orden = '" & Trim(vgRs!Num_Orden) & "' AND "
    vgSql = vgSql & "rut_receptor = '" & Trim(vgRs!Rut_Receptor) & "' AND "
    vgSql = vgSql & "num_ordencar <> '" & vlNumOrdenCar & "' AND "
    vgSql = vgSql & "cod_tipreceptor = '" & vlCodTipReceptor & "' "
    
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

Function flBuscaMtoHaberes()

On Error GoTo Err_flBuscaMtoHaberes
    
    Call flBuscaCodOtrosHabDes("H")
    vgSql = ""
    vgSql = vgSql & "SELECT sum(mto_conhabdes) as mtohaberes "
    vgSql = vgSql & "FROM PP_TTMP_CONPAGOPEN "
    vgSql = vgSql & "WHERE num_perpago = '" & Trim(vlPeriodo) & "' AND "
    vgSql = vgSql & "num_poliza = '" & Trim(vgRs!num_poliza) & "' AND "
    vgSql = vgSql & "num_orden = '" & Trim(vgRs!Num_Orden) & "' AND "
    vgSql = vgSql & "rut_receptor = '" & Trim(vgRs!Rut_Receptor) & "' AND "
    vgSql = vgSql & "cod_conhabdes IN '" & vlCodOtrosHabDes & "' "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       vlMtoHaber = (vgRegistro!mtohaberes)
    End If
    
Exit Function
Err_flBuscaMtoHaberes:
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
    vgSql = vgSql & "mto_conhabdes,num_idenReceptor,cod_tipoidenreceptor,num_perpago "
    vgSql = vgSql & "FROM PP_TTMP_CONPAGOPEN "
    vgSql = vgSql & "WHERE cod_conhabdes IN " & vlCodOtrosHabDes & " "
    
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       While Not vgRegistro.EOF
       
             vlNumCargas = 0
             vlMtoCargas = 0
             vlMtoHaber = 0
             vlMtoDescto = 0
             vlMtoHabDes = 0
             vlPerPagoRetencion = vgRegistro!Num_PerPago 'HQR 28/04/2005
             'Buscar el Nùmero de Retenciòn de cada uno de los detalles seleccionados.
             
             vgSql = ""
             vgSql = "SELECT num_retencion FROM PP_TMAE_RETJUDICIAL "
             vgSql = vgSql & "WHERE num_poliza = '" & (vgRs!num_poliza) & "' AND "
             vgSql = vgSql & "num_endoso = " & (vgRs!num_endoso) & " AND "
             vgSql = vgSql & "num_orden = " & (vgRs!Num_Orden) & " AND "
             vgSql = vgSql & "num_idenreceptor = '" & (vgRs!Num_IdenReceptor) & "' AND "
             vgSql = vgSql & "cod_tipoidenreceptor = " & (vgRs!Cod_TipoIdenReceptor) & " AND "
             vgSql = vgSql & "cod_tipret = '" & Trim(vlTipoRetJud) & "' AND "
             If vgTipoBase = "ORACLE" Then
                'HQR 28/04/2005 Existen distintos periodos (primer pago y regimen)
                'vgSql = vgSql & "substr(fec_iniret,1,6) <= '" & vlPeriodo & "' AND "
                'vgSql = vgSql & "substr(fec_terret,1,6) >= '" & vlPeriodo & "' "
                vgSql = vgSql & "substr(fec_iniret,1,6) <= '" & vlPerPagoRetencion & "' AND "
                vgSql = vgSql & "substr(fec_terret,1,6) >= '" & vlPerPagoRetencion & "' "
             Else
                 'HQR 28/04/2005 Existen distintos periodos (primer pago y regimen)
                 'vgSql = vgSql & "substring(fec_iniret,1,6) <= '" & vlPeriodo & "' AND "
                 'vgSql = vgSql & "substring(fec_terret,1,6) >= '" & vlPeriodo & "' "
                 vgSql = vgSql & "substring(fec_iniret,1,6) <= '" & vlPerPagoRetencion & "' AND "
                 vgSql = vgSql & "substring(fec_terret,1,6) >= '" & vlPerPagoRetencion & "' "
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

Function flProcesoGE()
    Dim vlSql As String 'HQR 17/05/2005
On Error GoTo Err_flProcesoGE

   vlPeriodo = Trim(Txt_Anno.Text + Txt_Mes.Text)
      
   vlArchivo = strRpt & "PP_Rpt_GELibroPenMin.rpt"   '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
      MsgBox "Archivo de Reporte de Libro de Pensiones Mínimas no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
      Screen.MousePointer = 0
      Exit Function
   End If
      
   vlSql = "DELETE FROM PP_TTMP_CONPENMIN WHERE cod_usuario = '" & vgUsuario & "'"
   vgConexionBD.Execute (vlSql)
   
   Screen.MousePointer = 11
   
   Call flCargaTablaTemporalGE
   
   Call fgVigenciaQuiebra(DateSerial(Txt_Anno, Txt_Mes, 1))

   vgQuery = ""
   vgQuery = "{PP_TTMP_CONPENMIN.COD_USUARIO} = '" & Trim(vgUsuario) & "' "
   
       
   Rpt_Calculo.Reset
   Rpt_Calculo.WindowState = crptMaximized
   Rpt_Calculo.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
   Rpt_Calculo.Connect = vgRutaDataBase       ' o App.Path & "\Nestle.mdb"
   Rpt_Calculo.SelectionFormula = vgQuery
'''   Rpt_General.SelectionFormula = vgQuery
      
'''   vgPalabra = Txt_FechaIni.Text & " - " & Txt_FechaTer.Text
   
   vgPalabra = Txt_Mes.Text + "-" + Txt_Anno.Text
   
   Rpt_Calculo.Formulas(0) = ""
   Rpt_Calculo.Formulas(1) = ""
   Rpt_Calculo.Formulas(2) = ""
   Rpt_Calculo.Formulas(3) = ""
   Rpt_Calculo.Formulas(4) = ""
   
   Rpt_Calculo.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
   Rpt_Calculo.Formulas(1) = "NombreSistema= '" & vgNombreSistema & "'"
   Rpt_Calculo.Formulas(2) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
   Rpt_Calculo.Formulas(3) = "Periodo = '" & vgPalabra & "'"
   Rpt_Calculo.Formulas(4) = "GlosaQuiebra = '" & vgGlosaQuiebra & "'"
   
   
   Rpt_Calculo.SubreportToChange = ""
   Rpt_Calculo.Destination = crptToWindow
   Rpt_Calculo.WindowState = crptMaximized
   Rpt_Calculo.WindowTitle = ""
'   Rpt_Calculo.SelectionFormula = ""
   Rpt_Calculo.Action = 1
   Screen.MousePointer = 0

Exit Function
Err_flProcesoGE:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flCargaTablaTemporalGE()
On Error GoTo Err_flCargaTablaTemporalGE
    
    vlCodUsuario = ""
    vlNumPerPago = ""
    vlNumPoliza = ""
    vlNumEndoso = 0
    vlNumOrden = 0
    vlNumResGarEst = ""
    vlNumAnnoRes = ""
    vlCodTipRes = ""
    vlPrcDeduccion = 0
    vlCodDerGarEst = ""
    vlMtoPension = 0
    vlMtoPensionPesos = 0
    vlMtoGarEst = 0
    vlMtoOtroHab = 0
    vlMtoOtroDes = 0
    vlMtoBono = 0
 
    vlNumPerPago = vlPeriodo
    
    vgPalabra = ""
    Call flBuscaCodOtrosHabDesGE
    
    vgSql = ""
    'vgSql = "SELECT DISTINCT p.num_poliza,p.num_orden,l.num_endoso,p.cod_conhabdes, "
    'vgSql = vgSql & "p.fec_inipago,p.fec_terpago,p.mto_conhabdes "
    vgSql = "SELECT DISTINCT p.num_poliza,p.num_orden,l.num_endoso,p.num_perpago,l.mto_pension " ',p.cod_conhabdes, " 'HQR 03/05/2005 Se deben obtener las distintas pólizas pero sin llegar al nivel de concepto o salen duplicadas
    'vgSql = vgSql & "p.fec_inipago,p.fec_terpago,p.mto_conhabdes "
    vgSql = vgSql & "FROM PP_TTMP_CONPAGOPEN p, PP_TTMP_CONLIQPAGOPEN l "
    vgSql = vgSql & "WHERE p.cod_conhabdes IN " & vlCodOtrosHabDesGE & " AND "
    vgSql = vgSql & "p.cod_tipomod = '" & clTipoGE & "' AND "
'    vgSql = vgSql & "p.num_perpago = '" & Trim(vlNumPerPago) & "' AND "
    vgSql = vgSql & "p.cod_tipomod = l.cod_tipomod AND "
    vgSql = vgSql & "p.num_perpago = l.num_perpago AND "
    vgSql = vgSql & "p.num_poliza = l.num_poliza AND "
    vgSql = vgSql & "p.num_orden = l.num_orden AND "
    vgSql = vgSql & "p.rut_receptor = l.rut_receptor AND "
    vgSql = vgSql & "p.cod_tipreceptor = l.cod_tipreceptor "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
          
       While Not vgRs.EOF
       
            
       
            vlNumResGarEst = ""
            vlNumAnnoRes = ""
            vlCodTipRes = ""
            vlPrcDeduccion = 0
            vlCodDerGarEst = ""
            vlMtoPension = 0
            vlMtoPensionPesos = 0
            vlMtoGarEst = 0
            vlMtoOtroHab = 0
            vlMtoOtroDes = 0
            vlMtoBono = 0
             
             vlNumPoliza = (vgRs!num_poliza)
             
             If IsNull(vgRs!num_endoso) Then
                vlNumEndoso = ""
             Else
                 vlNumEndoso = (vgRs!num_endoso)
             End If
             vlNumOrden = (vgRs!Num_Orden)
             
             vlPerPagoGEPenMin = (vgRs!Num_PerPago)
             vlMtoPensionUF = (vgRs!Mto_Pension)
             vlMtoPension = vlMtoPensionUF
                  
'Insertar registro con valores de montos en 0
             Call flInsertarTablaTemporalGE
             
    'Actualizar registro con Valor de MtoPensionGE
             Call flAgregarMonto(clCodConHDMtoPen0102, clTipoGE, vlMtoPensionPesos)
    'Actualizar registro con Valor de MtoGarEstGE
             Call flAgregarMonto(clCodGE, clTipoGE, vlMtoGarEst)
    'Actualizar registro con Valor de MtoBono
             Call flAgregarMonto(clCodBonInv, clTipoGE, vlMtoBono)
    
             vlCodigo = clCodH
             vgPalabra = ""
             vgPalabra = " AND cod_tipmov = '" & Trim(vlCodigo) & "' AND "
             vgPalabra = vgPalabra & "cod_conhabdes <> " & clCodGE & " AND "
             vgPalabra = vgPalabra & "cod_conhabdes <> " & clCodBonInv & " "
                          
             Call flBuscaCodOtrosHabDesGE
    
    'Actualizar registro con Valor de MtoOtroHabGE
             Call flAgregarMonto(vlCodOtrosHabDesGE, clTipoGE, vlMtoOtroHab)
             
             vlCodigo = clCodD
             vgPalabra = ""
             vgPalabra = " AND cod_tipmov = '" & Trim(vlCodigo) & "' AND "
             vgPalabra = vgPalabra & "cod_conhabdes <> " & clCodGE & " AND "
             vgPalabra = vgPalabra & "cod_conhabdes <> " & clCodBonInv & " "
                          
             Call flBuscaCodOtrosHabDesGE
             
    'Actualizar registro con Valor de MtoOtroDesGE
             Call flAgregarMonto(vlCodOtrosHabDesGE, clTipoGE, vlMtoOtroDes)
             
             'abv El valor se obtiene de la Tabla ConLiqPagoPen
             'Call flAgregarMontoPension(vlMtoPension)
                          
             vgSql = ""
             vgSql = " UPDATE PP_TTMP_CONPENMIN SET "
             vgSql = vgSql & "mto_pension = " & Str(vlMtoPension) & ", "
             vgSql = vgSql & "mto_pensionpesos = " & Str(vlMtoPensionPesos) & ", "
             vgSql = vgSql & "mto_garest = " & Str(vlMtoGarEst) & ", "
             vgSql = vgSql & "mto_otrohab = " & Str(vlMtoOtroHab) & ", "
             vgSql = vgSql & "mto_otrodes = " & Str(vlMtoOtroDes) & ", "
             vgSql = vgSql & "mto_bono = " & Str(vlMtoBono) & " "
             vgSql = vgSql & "WHERE cod_usuario = '" & Trim(vgUsuario) & "' AND "
             vgSql = vgSql & "num_perpago = '" & Trim(vlPerPagoGEPenMin) & "' AND "
             vgSql = vgSql & "num_poliza = '" & Trim(vlNumPoliza) & "' AND "
             vgSql = vgSql & "num_orden = " & Str(vlNumOrden) & " "
             vgConexionBD.Execute vgSql
            
             vgRs.MoveNext
       Wend
    End If

Exit Function
Err_flCargaTablaTemporalGE:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flBuscaCodOtrosHabDesGE() '(Codigo As String)
On Error GoTo flBuscaCodOtrosHabDesGE
'Busca códigos de otros haberes y otros descuentos por concepto de GE
'Garantia Estatal
    
    vlCodOtrosHabDesGE = ""
    vgSql = ""
    vgSql = vgSql & "SELECT cod_conhabdes "
    vgSql = vgSql & "FROM ma_tpar_conhabdes "
    vgSql = vgSql & "WHERE cod_modorigen = '" & clTipoGE & "' " 'AND "
    vgSql = vgSql & vgPalabra
    'vgSql = vgSql & "cod_tipmov = '" & Trim(Codigo) & "' AND "
    'vgSql = vgSql & "cod_conhabdes <> " & clCodGE & " AND "
    'vgSql = vgSql & "cod_conhabdes <> " & clCodBonInv & " "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       While Not vgRegistro.EOF
             If vlCodOtrosHabDesGE = "" Then
                vlCodOtrosHabDesGE = "("
             End If
             vlCodOtrosHabDesGE = (vlCodOtrosHabDesGE & "'" & (vgRegistro!Cod_ConHabDes) & "'")
             vgRegistro.MoveNext
             If Not vgRegistro.EOF Then
                vlCodOtrosHabDesGE = (vlCodOtrosHabDesGE & ",")
             End If
       Wend
       vlCodOtrosHabDesGE = (vlCodOtrosHabDesGE & ")")
    End If

Exit Function
flBuscaCodOtrosHabDesGE:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flInsertarTablaTemporalGE()
On Error GoTo Err_flInsertarTablaTemporalGE

'Seleccionar Datos de Resolución para agregar al nuevo registro de la
'Tabla temporal
    vgSql = ""
    vgSql = "SELECT num_resgarest,num_annores,cod_tipres, "
    vgSql = vgSql & "prc_deduccion "
    vgSql = vgSql & "FROM PP_TMAE_GARESTRES "
    vgSql = vgSql & "WHERE num_poliza = '" & Trim(vlNumPoliza) & "' AND "
    vgSql = vgSql & "num_endoso = " & Str(vlNumEndoso) & " AND "
    vgSql = vgSql & "num_orden = " & Str(vlNumOrden) & " AND "
    If vgTipoBase = "ORACLE" Then
       vgSql = vgSql & "substr(fec_inires,1,6) <= '" & Trim(vlPerPagoGEPenMin) & "' AND "
       vgSql = vgSql & "substr(fec_terres,1,6) >= '" & Trim(vlPerPagoGEPenMin) & "' "
    Else
        vgSql = vgSql & "substring(fec_inires,1,6) <= '" & Trim(vlPerPagoGEPenMin) & "' AND "
        vgSql = vgSql & "substring(fec_terres,1,6) >= '" & Trim(vlPerPagoGEPenMin) & "' "
    End If
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       vlNumResGarEst = (vgRegistro!NUM_RESGAREST)
       vlNumAnnoRes = (vgRegistro!NUM_ANNORES)
       vlCodTipRes = (vgRegistro!COD_TIPRES)
       vlPrcDeduccion = (vgRegistro!PRC_DEDUCCION)
    End If
'Seleccionar datos de Estado de Garantía Estatal, para agregar al nuevo registro
'de la tabla temporal
    vgSql = ""
    vgSql = "SELECT cod_dergarest "
    vgSql = vgSql & "FROM PP_TMAE_GARESTESTADO "
    vgSql = vgSql & "WHERE num_poliza = '" & Trim(vlNumPoliza) & "' AND "
    vgSql = vgSql & "num_endoso = " & Str(vlNumEndoso) & " AND "
    vgSql = vgSql & "num_orden = " & Str(vlNumOrden) & " AND "
    If vgTipoBase = "ORACLE" Then
       vgSql = vgSql & "substr(fec_iniestgarest,1,6) <= '" & Trim(vlPerPagoGEPenMin) & "' AND "
       vgSql = vgSql & "substr(fec_terestgarest,1,6) >= '" & Trim(vlPerPagoGEPenMin) & "' "
    Else
        vgSql = vgSql & "substring(fec_iniestgarest,1,6) <= '" & Trim(vlPerPagoGEPenMin) & "' AND "
        vgSql = vgSql & "substring(fec_terestgarest,1,6) >= '" & Trim(vlPerPagoGEPenMin) & "' "
    End If
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       vlCodDerGarEst = (vgRegistro!COD_DERGAREST)
    End If
                    
'Insertar nuevo registro en la tabla temporal
    vgSql = ""
    vgSql = "INSERT INTO PP_TTMP_CONPENMIN "
    vgSql = vgSql & "(cod_usuario,num_perpago,num_poliza,num_endoso,num_orden, "
    'num_endoso ,num_orden, "
    vgSql = vgSql & " num_resgarest,num_annores,cod_tipres,prc_deduccion, "
    vgSql = vgSql & " cod_dergarest,mto_pension,mto_pensionpesos,mto_garest, "
    vgSql = vgSql & " mto_otrohab,mto_otrodes,mto_bono "
    vgSql = vgSql & " ) VALUES ( "
    vgSql = vgSql & " '" & vgUsuario & "', "
    vgSql = vgSql & " '" & Trim(vlPerPagoGEPenMin) & "' , "
    vgSql = vgSql & " '" & Trim(vlNumPoliza) & "' , "
    vgSql = vgSql & " " & Str(vlNumEndoso) & ", "
    vgSql = vgSql & " " & Str(vlNumOrden) & ", "
    vgSql = vgSql & " '" & Trim(vlNumResGarEst) & "', "
    vgSql = vgSql & " '" & Trim(vlNumAnnoRes) & "', "
    vgSql = vgSql & " '" & Trim(vlCodTipRes) & "', "
    vgSql = vgSql & " " & Str(vlPrcDeduccion) & ", "
    vgSql = vgSql & " '" & Trim(vlCodDerGarEst) & "', "
    vgSql = vgSql & " " & Str(vlMtoPension) & ", "
    vgSql = vgSql & " " & Str(vlMtoPensionPesos) & ", "
    vgSql = vgSql & " " & Str(vlMtoGarEst) & ", "
    vgSql = vgSql & " " & Str(vlMtoOtroHab) & ", "
    vgSql = vgSql & " " & Str(vlMtoOtroDes) & ", "
    vgSql = vgSql & " " & Str(vlMtoBono) & ") "
    vgConexionBD.Execute vgSql
    
Exit Function
Err_flInsertarTablaTemporalGE:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flAgregarMonto(CodHabDes, CodMod As String, monto As Double)
On Error GoTo Err_flAgregarMonto

'Seleccionar Montos de Pensión por concepto de Garantìa Estatal
    vgSql = ""
    vgSql = "SELECT SUM(mto_conhabdes) as monto "
    vgSql = vgSql & "FROM PP_TTMP_CONPAGOPEN "
    vgSql = vgSql & "WHERE cod_conhabdes IN " & CodHabDes & " AND "
    vgSql = vgSql & "cod_tipomod = '" & CodMod & "' AND "
    vgSql = vgSql & "num_perpago = '" & Trim(vlPerPagoGEPenMin) & "' AND "
    vgSql = vgSql & "num_poliza = '" & Trim(vlNumPoliza) & "' AND "
    vgSql = vgSql & "num_orden = " & Str(vlNumOrden) & " "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       If Not IsNull(vgRegistro!monto) Then
          monto = (vgRegistro!monto)
       Else
           monto = 0
       End If
    End If
    
Exit Function
Err_flAgregarMonto:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flAgregarMontoPension(vlMtoPensionUF As Double)
Dim vlNum_Poliza As String 'HQR 17/05/2005
On Error GoTo Err_flAgregarMontoPension
    
'Seleccionar Monto de Pension en UF desde tabla Beneficiarios
    vlMtoPensionUF = 0
    
    vgSql = ""
    vgSql = "SELECT mto_pension,mto_pensiongar,fec_terpagopengar "
    vgSql = vgSql & "FROM PP_TMAE_BEN "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "num_poliza = '" & Trim(vlNumPoliza) & "' AND "
    vgSql = vgSql & "num_endoso = '" & Trim(vlNumEndoso) & "' AND "
    vgSql = vgSql & "num_orden = " & Str(vlNumOrden) & " "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
        If Trim(vlNumPoliza) = "0000021766" Then
            vlNum_Poliza = Trim(vlNumPoliza)
        End If
        If Trim(vlNumPoliza) = "0000021778" Then
            vlNum_Poliza = Trim(vlNumPoliza)
        End If
        If Trim(vlNumPoliza) = "0000021803" Then
            vlNum_Poliza = Trim(vlNumPoliza)
        End If

       If IsNull(vgRegistro!Fec_TerPagoPenGar) Then
          vlMtoPensionUF = (vgRegistro!Mto_Pension)
       Else
           If vlPeriodo > Mid((vgRegistro!Fec_TerPagoPenGar), 1, 6) Then
              vlMtoPensionUF = (vgRegistro!Mto_Pension)
           Else
               If vlPeriodo <= Mid((vgRegistro!Fec_TerPagoPenGar), 1, 6) Then
                  vlMtoPensionUF = (vgRegistro!Mto_PensionGar)
               End If
           End If
       End If
    End If

Exit Function
Err_flAgregarMontoPension:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'*********************************************************

Function flProcesoHabDesGE()
Dim vlSql As String 'HQR 17/05/2005
On Error GoTo Err_flProcesoHabDesGE

   vlPeriodo = Trim(Txt_Anno.Text + Txt_Mes.Text)
      
   vlSql = "DELETE FROM PP_TTMP_CONHABDESGE WHERE cod_usuario = '" & vgUsuario & "'"
   vgConexionBD.Execute (vlSql)
   
   Call flCargaTablaTemporalHDGE
                    
   Screen.MousePointer = 11
   
   vlArchivo = strRpt & "PP_Rpt_GEHabDesGE.rpt"   '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
      MsgBox "Archivo de Reporte de Haberes y Descuentos de Garantía Estatal no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
      Screen.MousePointer = 0
      Exit Function
   End If
   
   Call fgVigenciaQuiebra(DateSerial(Txt_Anno, Txt_Mes, 1))

   vgQuery = ""
   vgQuery = "{PP_TTMP_CONHABDESGE.COD_USUARIO} = '" & Trim(vgUsuario) & "' "
   
       
   Rpt_Calculo.Reset
   Rpt_Calculo.WindowState = crptMaximized
   Rpt_Calculo.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
   Rpt_Calculo.Connect = vgRutaDataBase       ' o App.Path & "\Nestle.mdb"
   Rpt_Calculo.SelectionFormula = vgQuery
'''   Rpt_General.SelectionFormula = vgQuery
      
'''   vgPalabra = Txt_FechaIni.Text & " - " & Txt_FechaTer.Text
   
   vgPalabra = Txt_Mes.Text + "-" + Txt_Anno.Text
   
   Rpt_Calculo.Formulas(0) = ""
   Rpt_Calculo.Formulas(1) = ""
   Rpt_Calculo.Formulas(2) = ""
   Rpt_Calculo.Formulas(3) = ""
   Rpt_Calculo.Formulas(4) = ""
   
   Rpt_Calculo.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
   Rpt_Calculo.Formulas(1) = "NombreSistema= '" & vgNombreSistema & "'"
   Rpt_Calculo.Formulas(2) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
   Rpt_Calculo.Formulas(3) = "Periodo = '" & vgPalabra & "'"
   Rpt_Calculo.Formulas(4) = "GlosaQuiebra = '" & vgGlosaQuiebra & "'"
   
   Rpt_Calculo.SubreportToChange = ""
   Rpt_Calculo.Destination = crptToWindow
   Rpt_Calculo.WindowState = crptMaximized
   Rpt_Calculo.WindowTitle = ""
   Rpt_Calculo.SelectionFormula = ""
   Rpt_Calculo.Action = 1
   Screen.MousePointer = 0

Exit Function
Err_flProcesoHabDesGE:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flCargaTablaTemporalHDGE()
On Error GoTo Err_flCargaTablaTemporalHDGE
    
    vlCodUsuario = ""
    vlNumPerPago = ""
    vlNumPoliza = ""
    vlNumEndoso = 0
    vlNumOrden = 0
    vlCodDerGarEst = ""
    vlMtoPension = 0
    vlMtoPensionPesos = 0
    vlCodConHabDes = ""
    vlCodTipMov = ""
    vlNumCuotas = 0
    vlMtoCuota = 0
    vlMtoTotal = 0
    vlCodMoneda = ""
    vlMtoTotalHabDes = 0
 
    vlNumPerPago = vlPeriodo
    
    vgPalabra = ""
    Call flBuscaCodOtrosHabDesGE
    
    vgSql = ""
    vgSql = "SELECT DISTINCT p.num_poliza,l.num_endoso,p.num_orden,p.cod_conhabdes, "
    vgSql = vgSql & "p.fec_inipago,p.fec_terpago,p.mto_conhabdes,p.num_perpago,l.mto_pension "
    vgSql = vgSql & "FROM PP_TTMP_CONPAGOPEN p, PP_TTMP_CONLIQPAGOPEN l "
    vgSql = vgSql & "WHERE p.cod_conhabdes IN " & vlCodOtrosHabDesGE & " AND "
    vgSql = vgSql & "p.cod_tipomod = '" & clTipoGE & "' AND "
'    vgSql = vgSql & "p.num_perpago = '" & Trim(vlNumPerPago) & "' AND "
    vgSql = vgSql & "p.cod_tipomod = l.cod_tipomod AND "
    vgSql = vgSql & "p.num_perpago = l.num_perpago AND "
    vgSql = vgSql & "p.num_poliza = l.num_poliza AND "
    vgSql = vgSql & "p.num_orden = l.num_orden AND "
    vgSql = vgSql & "p.rut_receptor = l.rut_receptor AND "
    vgSql = vgSql & "p.cod_tipreceptor = l.cod_tipreceptor "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
          
       While Not vgRs.EOF
       
            vlCodDerGarEst = ""
            vlMtoPension = 0
            vlMtoPensionPesos = 0
            vlCodConHabDes = ""
            vlCodTipMov = ""
            vlNumCuotas = 0
            vlMtoCuota = 0
            vlMtoTotal = 0
            vlCodMoneda = ""
            vlMtoTotalHabDes = 0
             
             vlNumPoliza = (vgRs!num_poliza)
             vlNumEndoso = (vgRs!num_endoso) 'hqr
             vlNumOrden = (vgRs!Num_Orden)
             vlCodConHabDes = (vgRs!Cod_ConHabDes)
             vlMtoTotalHabDes = (vgRs!Mto_ConHabDes)
             
             vlPerPagoGEHabDes = (vgRs!Num_PerPago)
             
             vlMtoPensionUF = (vgRs!Mto_Pension)
             vlMtoPension = vlMtoPensionUF
             
    'Actualizar registro con Valor de Código de Tipo de Movimiento
             Call flAgregarTipMovHD
             
    'Actualizar registro con Valor de MtoPensionGE (En Pesos)
             Call flAgregarMontoHD(clCodConHDMtoPen0102, clTipoGE, vlMtoPensionPesos)
             
    'Actualizar registro con Valor de MtoPensionGe (En UF)
             'abv El valor se obtiene de la Tabla ConLiqPagoPen
             'Call flAgregarMontoPensionHD
             
    'Actualizar registro con Valores de Cuota
             Call flAgregarValoresCuota
                          
             'Insertar registro con valores de montos en 0
             Call flInsertarTablaTemporalHDGE
                          
                          
'             vgSql = ""
'             vgSql = " UPDATE PP_TTMP_CONHABDESGE SET "
'             vgSql = vgSql & "cod_dergarest = '" & Trim(vlCodDerGarEst) & "', "
'             vgSql = vgSql & "mto_pension = " & Str(vlMtoPension) & ", "
'             vgSql = vgSql & "mto_pensionpesos = " & Str(vlMtoPensionPesos) & ", "
'             vgSql = vgSql & "cod_tipmov = '" & Trim(vlCodTipMov) & "', "
'             vgSql = vgSql & "num_cuotas = " & Str(vlNumCuotas) & ", "
'             vgSql = vgSql & "mto_cuota = " & Str(vlMtoCuota) & ", "
'             vgSql = vgSql & "mto_total = " & Str(vlMtoTotal) & " "
'             vgSql = vgSql & "WHERE cod_usuario = '" & Trim(vlCodUsuario) & "' AND "
'             vgSql = vgSql & "num_poliza = '" & Trim(vlNumPoliza) & "' AND "
'             vgSql = vgSql & "num_endoso = " & Str(vlNumEndoso) & " AND "
'             vgSql = vgSql & "num_orden = " & Str(vlNumOrden) & "' AND "
'             vgSql = vgSql & "cod_conhabdes = '" & Trim(vlCodConHabDes) & "' AND "
'             If vgTipoBase = "ORACLE" Then
'                vgSql = vgSql & "substr(fec_inihabdes,1,6) <= '" & Trim(vlNumPerPago) & "' AND "
'                vgSql = vgSql & "substr(fec_terhabdes,1,6) >= '" & Trim(vlNumPerPago) & "' "
'             Else
'                 vgSql = vgSql & "substring(fec_inihabdes,1,6) <= '" & Trim(vlNumPerPago) & "' AND "
'                 vgSql = vgSql & "substring(fec_terhabdes,1,6) >= '" & Trim(vlNumPerPago) & "' "
'             End If
'             vgConexionBD.Execute vgSql
            
             vgRs.MoveNext
       
       Wend
       
    End If

Exit Function
Err_flCargaTablaTemporalHDGE:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flBuscaCodOtrosHabDesHDGE() '(Codigo As String)
On Error GoTo Err_flBuscaCodOtrosHabDesHDGE
'Busca códigos de otros haberes y otros descuentos por concepto de GE
'Garantia Estatal
    
    vlCodOtrosHabDesGE = ""
    vgSql = ""
    vgSql = vgSql & "SELECT cod_conhabdes "
    vgSql = vgSql & "FROM ma_tpar_conhabdes "
    vgSql = vgSql & "WHERE cod_modorigen = '" & clTipoGE & "' " 'AND "
    vgSql = vgSql & vgPalabra
    'vgSql = vgSql & "cod_tipmov = '" & Trim(Codigo) & "' AND "
    'vgSql = vgSql & "cod_conhabdes <> " & clCodGE & " AND "
    'vgSql = vgSql & "cod_conhabdes <> " & clCodBonInv & " "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       While Not vgRegistro.EOF
             If vlCodOtrosHabDesGE = "" Then
                vlCodOtrosHabDesGE = "("
             End If
             vlCodOtrosHabDesGE = (vlCodOtrosHabDesGE & "'" & (vgRegistro!Cod_ConHabDes) & "'")
             vgRegistro.MoveNext
             If Not vgRegistro.EOF Then
                vlCodOtrosHabDesGE = (vlCodOtrosHabDesGE & ",")
             End If
       Wend
       vlCodOtrosHabDesGE = (vlCodOtrosHabDesGE & ")")
    End If

Exit Function
Err_flBuscaCodOtrosHabDesHDGE:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flInsertarTablaTemporalHDGE()
On Error GoTo Err_flInsertarTablaTemporalHDGE

'Seleccionar datos de Estado de Garantía Estatal, para agregar al nuevo registro
'de la tabla temporal
    vgSql = ""
    vgSql = "SELECT cod_dergarest "
    vgSql = vgSql & "FROM PP_TMAE_GARESTESTADO "
    vgSql = vgSql & "WHERE num_poliza = '" & Trim(vlNumPoliza) & "' AND "
    vgSql = vgSql & "num_endoso = " & Str(vlNumEndoso) & " AND "
    vgSql = vgSql & "num_orden = " & Str(vlNumOrden) & " AND "
    If vgTipoBase = "ORACLE" Then
       vgSql = vgSql & "substr(fec_iniestgarest,1,6) <= '" & Trim(vlPerPagoGEHabDes) & "' AND "
       vgSql = vgSql & "substr(fec_terestgarest,1,6) >= '" & Trim(vlPerPagoGEHabDes) & "' "
    Else
        vgSql = vgSql & "substring(fec_iniestgarest,1,6) <= '" & Trim(vlPerPagoGEHabDes) & "' AND "
        vgSql = vgSql & "substring(fec_terestgarest,1,6) >= '" & Trim(vlPerPagoGEHabDes) & "' "
    End If
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       vlCodDerGarEst = (vgRegistro!COD_DERGAREST)
    End If
                    
                    
'Confirmar existencia del nuevo registro antes de insertar
    vgSql = ""
    vgSql = "SELECT num_poliza "
    vgSql = vgSql & "FROM PP_TTMP_CONHABDESGE "
    vgSql = vgSql & "WHERE cod_usuario = '" & Trim(vgUsuario) & "' AND "
    vgSql = vgSql & "num_perpago = '" & Trim(vlPerPagoGEHabDes) & "' AND "
    vgSql = vgSql & "num_poliza = '" & Trim(vlNumPoliza) & "' AND "
    vgSql = vgSql & "num_orden = " & Str(vlNumOrden) & " AND "
    vgSql = vgSql & "cod_tipmov = '" & Trim(vlCodTipMov) & "' AND "
    vgSql = vgSql & "cod_conhabdes = '" & Trim(vlCodConHabDes) & "' "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       Exit Function
    Else
'Insertar nuevo registro en la tabla temporal
        vgSql = ""
        vgSql = "INSERT INTO PP_TTMP_CONHABDESGE "
        vgSql = vgSql & "(cod_usuario,num_perpago,num_poliza,num_endoso,num_orden, "
        vgSql = vgSql & " cod_dergarest,mto_pension,mto_pensionpesos,cod_tipmov, "
        vgSql = vgSql & " cod_conhabdes,cod_moneda,num_cuotas,mto_cuota, "
        vgSql = vgSql & " mto_total,mto_totalhabdes "
        vgSql = vgSql & " ) VALUES ( "
        vgSql = vgSql & " '" & vgUsuario & "', "
        vgSql = vgSql & " '" & Trim(vlPerPagoGEHabDes) & "' , "
        vgSql = vgSql & " '" & Trim(vlNumPoliza) & "' , "
        vgSql = vgSql & " " & Str(vlNumEndoso) & ", "
        vgSql = vgSql & " " & Str(vlNumOrden) & ", "
        vgSql = vgSql & " '" & Trim(vlCodDerGarEst) & "', "
        vgSql = vgSql & " " & Str(vlMtoPension) & ", "
        vgSql = vgSql & " " & Str(vlMtoPensionPesos) & ", "
        vgSql = vgSql & " '" & Trim(vlCodTipMov) & "', "
        vgSql = vgSql & " '" & Trim(vlCodConHabDes) & "', "
        vgSql = vgSql & " '" & Trim(vlCodMoneda) & "', "
        vgSql = vgSql & " " & Str(vlNumCuotas) & ", "
        vgSql = vgSql & " " & Str(vlMtoCuota) & ", "
        vgSql = vgSql & " " & Str(vlMtoTotal) & ", "
        vgSql = vgSql & " " & Str(vlMtoTotalHabDes) & ") "
        vgConexionBD.Execute vgSql
    
    End If
    
Exit Function
Err_flInsertarTablaTemporalHDGE:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flAgregarMontoHD(CodHabDes, CodMod As String, monto As Double)

On Error GoTo Err_flAgregarMontoHD

'Seleccionar Montos de Pensión por concepto de Garantìa Estatal
    vgSql = ""
    vgSql = "SELECT SUM (mto_conhabdes) as montohabdes "
    vgSql = vgSql & "FROM PP_TTMP_CONPAGOPEN "
    vgSql = vgSql & "WHERE cod_conhabdes IN " & CodHabDes & " AND "
    vgSql = vgSql & "cod_tipomod = '" & CodMod & "' AND "
    vgSql = vgSql & "num_perpago = '" & Trim(vlPerPagoGEHabDes) & "' AND "
    vgSql = vgSql & "num_poliza = '" & Trim(vlNumPoliza) & "' AND "
    vgSql = vgSql & "num_orden = " & Str(vlNumOrden) & " "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       If Not IsNull(vgRegistro!montohabdes) Then
          monto = (vgRegistro!montohabdes)
       Else
           monto = 0
       End If
    End If
    
Exit Function
Err_flAgregarMontoHD:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flAgregarMontoPensionHD()
On Error GoTo Err_flAgregarMontoPensionHD
    
'Seleccionar Monto de Pension en UF desde tabla Beneficiarios
    vlMtoPensionUF = 0
    
    
    vgSql = ""
    vgSql = "SELECT mto_pension,mto_pensiongar,fec_terpagopengar "
    vgSql = vgSql & "FROM PP_TMAE_BEN "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "num_poliza = '" & Trim(vlNumPoliza) & "' AND "
    vgSql = vgSql & "num_endoso = " & Trim(vlNumEndoso) & " AND "
    vgSql = vgSql & "num_orden = " & Str(vlNumOrden) & " "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       If IsNull(vgRegistro!Fec_TerPagoPenGar) Then
          vlMtoPensionUF = (vgRegistro!Mto_Pension)
       Else
           If vlPeriodo > Mid((vgRegistro!Fec_TerPagoPenGar), 1, 6) Then
              vlMtoPensionUF = (vgRegistro!Mto_Pension)
           Else
               If vlPeriodo < Mid((vgRegistro!Fec_TerPagoPenGar), 1, 6) Then
                  vlMtoPensionUF = (vgRegistro!Mto_PensionGar)
               End If
           End If
       End If
    End If

Exit Function
Err_flAgregarMontoPensionHD:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flAgregarTipMovHD()
On Error GoTo Err_flAgregarTipMov

'Seleccionar Código de Tipo de Movimiento de Haber o Descuento
    vgSql = ""
    vgSql = "SELECT cod_tipmov "
    vgSql = vgSql & "FROM MA_TPAR_CONHABDES "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "cod_conhabdes = '" & vlCodConHabDes & "' AND "
    vgSql = vgSql & "cod_modorigen = '" & clTipoGE & "' "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       vlCodTipMov = (vgRegistro!cod_tipmov)
    End If
    
Exit Function
Err_flAgregarTipMov:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flAgregarValoresCuota()
On Error GoTo Err_flAgregarValoresCuota

'Seleccionar Valores de Cuotas para agregar a registro
    vgSql = ""
    vgSql = "SELECT num_cuotas,mto_cuota,mto_total,cod_moneda "
    vgSql = vgSql & "FROM PP_TMAE_HABDES "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "num_poliza = '" & vlNumPoliza & "' AND "
    vgSql = vgSql & "num_endoso = " & vlNumEndoso & " AND "
    vgSql = vgSql & "num_orden = " & vlNumOrden & " AND "
    vgSql = vgSql & "cod_conhabdes = '" & vlCodConHabDes & "' AND "
    If vgTipoBase = "ORACLE" Then
       vgSql = vgSql & "substr(fec_inihabdes,1,6) <= '" & Trim(vlNumPerPago) & "' AND "
       vgSql = vgSql & "substr(fec_terhabdes,1,6) >= '" & Trim(vlNumPerPago) & "' "
    Else
        vgSql = vgSql & "substring(fec_inihabdes,1,6) <= '" & Trim(vlNumPerPago) & "' AND "
        vgSql = vgSql & "substring(fec_terhabdes,1,6) >= '" & Trim(vlNumPerPago) & "' "
    End If
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       vlNumCuotas = (vgRegistro!Num_Cuotas)
       vlMtoCuota = (vgRegistro!MTO_CUOTA)
       vlMtoTotal = (vgRegistro!mto_total)
       vlCodMoneda = (vgRegistro!Cod_Moneda)
    End If

Exit Function
Err_flAgregarValoresCuota:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flModificaTabla(iCodigo As String, iMonto As Double)
 Dim Sql As String 'HQR 17/05/2005
'Permite modificar SOLO montos de reintegros y retroactivos

 Sql = ""
 Sql = " UPDATE PP_TTMP_CONASIGFAM SET "
 'Sql = Sql & " RUT_BEN = " & vlRutBen & ","
 'Sql = Sql & " DGV_BEN = '" & vlDigito & "',"
 'Sql = Sql & " MTO_CARGA = " & Str(vlMontoCarga) & ","
 'Sql = Sql & " COD_PAR = '" & vlCodPar & "',"
 'Sql = Sql & " COD_SITINV = '" & vlCodSitInv & "',"
 
 If iCodigo = "09" Then
    Sql = Sql & " MTO_RETRO = " & Str(iMonto) & ","
 End If
 If iCodigo = "30" Then
    Sql = Sql & " MTO_REINTEGRO = " & Str(iMonto) & ","
 End If
 
 'If vlFecVencimiento = "" Then
 '   Sql = Sql & " FEC_VENCARGA = NULL,"
 'Else
 '   Sql = Sql & " FEC_VENCARGA = '" & vlFecVencimiento & "',"
 'End If
 
 Sql = Sql & " MTO_TOTAL = " & Str(vlMto_total) & " "
 'Sql = Sql & " GLS_ELEMENTO = '" & Trim(vlDescripcion) & "',"
 'Sql = Sql & " RUT_BENREC = " & vlRutReceptor & ","
 'Sql = Sql & " DGV_BENREC = '" & vlDgvReceptor & "' "
 Sql = Sql & " WHERE COD_USUARIO = '" & vgUsuario & "' AND"
 'I--- ABV 11/05/2005 ---
 'Sql = Sql & " NUM_PERPAGO = '" & Trim(vlNumPerPago) & "' AND"
 Sql = Sql & " NUM_PERPAGO = '" & Trim(vlPerPagoAsigFam) & "' AND"
 'F--- ABV 11/05/2005 ---
 Sql = Sql & " NUM_POLIZA = '" & vlNumPoliza & "' AND"
 Sql = Sql & " NUM_ORDENREC = " & vlNumOrden & " AND"
 Sql = Sql & " NUM_ORDEN = " & vlNumOrdenCar & ""
 vgConexionBD.Execute Sql
 'Nota: Valor de monto total es calculado (sumado) en el reporte

End Function


''--------------CONCILIACION DE GARANTIA ESTATAL 20050330
''Informe de Control
''----GE CONCILIACION------------------------------------------
Function flProcesoConciliacionGE()
Dim vlSql As String 'HQR 17/05/2005

vlMtoCalculado = 0
On Error GoTo Err_flInfConciliacion

   vlCodTipoImp = clTipoImpC
   vlPeriodo = Trim(Txt_Anno.Text + Txt_Mes.Text)
   vlNumPerPago = vlPeriodo

   vlSql = "DELETE FROM pp_ttmp_conconcigarest WHERE cod_usuario = '" & vgUsuario & "' AND "
   vlSql = vlSql & "cod_tipoimp = '" & vlCodTipoImp & "' "
   vgConexionBD.Execute (vlSql)

   Call flCargaTablaTemporalConGE

   Screen.MousePointer = 11

   vlArchivo = strRpt & "PP_Rpt_GEConciliacionGECON.rpt"   '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
      MsgBox "Archivo de Reporte de Control de Conciliación de Garantía Estatal no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
      Screen.MousePointer = 0
      Exit Function
   End If

   vgQuery = ""
   vgQuery = "{PP_TTMP_CONCONCIGAREST.COD_USUARIO} = '" & Trim(vgUsuario) & "' AND "
   vgQuery = vgQuery & "{PP_TTMP_CONCONCIGAREST.COD_TIPOIMP} = '" & Trim(vlCodTipoImp) & "' "

   Rpt_Calculo.Reset
   Rpt_Calculo.WindowState = crptMaximized
   Rpt_Calculo.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
   Rpt_Calculo.Connect = vgRutaDataBase       ' o App.Path & "\Nestle.mdb"
   Rpt_Calculo.SelectionFormula = vgQuery

   Rpt_Calculo.Formulas(0) = ""
   Rpt_Calculo.Formulas(1) = ""
   Rpt_Calculo.Formulas(2) = ""
   Rpt_Calculo.Formulas(3) = ""

   vgPalabra = Txt_Mes.Text + "-" + Txt_Anno.Text

   Rpt_Calculo.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
   Rpt_Calculo.Formulas(1) = "NombreSistema= '" & vgNombreSistema & "'"
   Rpt_Calculo.Formulas(2) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
   Rpt_Calculo.Formulas(3) = "Periodo = '" & vgPalabra & "'"

'   Rpt_Reporte.SubreportToChange = ""
   Rpt_Calculo.Destination = crptToWindow
   Rpt_Calculo.WindowState = crptMaximized
   Rpt_Calculo.WindowTitle = ""
'   Rpt_Reporte.SelectionFormula = ""
   Rpt_Calculo.Action = 1
   Screen.MousePointer = 0

Exit Function
Err_flInfConciliacion:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

'----GE CONCILIACION------------------------------------------
Function flCargaTablaTemporalConGE()
Dim vlRegistroRes As ADODB.Recordset 'HQR 17/05/2005
On Error GoTo Err_flCargaTablaTemporalConGE

'    vlCodUsuario = ""
    vlCodTipoImp = ""
    vlNumPeriodo = ""
    vlFecPago = ""
'    vlNumPoliza = ""
'    vlNumEndoso = 0
'    vlNumOrden = 0
    vlCodTipRes = ""
    vlNumResGarEst = 0
    vlNumAnnoRes = 0
    vlNumDias = 0
    vlMtoPension = 0
    vlMtoPensionUF = 0
    vlMtoPensionMin = 0
    vlPrcDeduccion = 0
    vlMtoDeduccion = 0
    vlMtoGarEstQui = 0
    vlMtoGarEstNor = 0
    vlMtoGarEstCia = 0
    vlMtoGarEstRec = 0
    vlMtoDiferencia = 0
    vlCodDerGarEst = ""
    vlMtoHaber = 0
    vlMtoDescuento = 0
    vlCodEstado = ""
    
    vlCodUsuario = vgUsuario
    vlCodTipoImp = clTipoImpC
    vlCodDerGarEst = clSinCodDerGE

        
'    vgPalabra = ""
'    If vlPago = "P" Then
'        vgPalabra = "p.fec_calpagoreg "
'    Else
'        vgPalabra = "p.fec_calpripago "
'    End If
        
'    'Seleccionar periodos en los que se realizaron pagos
'    'según fechas ingresadas por pantalla
'    vgSql = ""
'    vgSql = "SELECT "
'    vgSql = vgSql & vgPalabra & "as fecha "
'    vgSql = vgSql & "FROM pp_tmae_propagopen p "
''    vgSql = vgSql & "WHERE "
''    vgSql = vgSql & "num_perpago >= '" & Trim(vlNumPerPago) & "' AND "
''    vgSql = vgSql & "num_perpago <= '" & Trim(vlNumPerPago) & "' "
'    Set vlRegistroPeriodo = vgConexionBD.Execute(vgSql)
'    If Not vlRegistroPeriodo.EOF Then
'
'        While Not vlRegistroPeriodo.EOF
'                vlFecPago = (vlRegistroPeriodo!FECHA)
'                vlNumPeriodo = Trim(Mid(vlFecPago, 1, 6))
                vlNumPeriodo = vlNumPerPago
                'Valores Nulos del registro, en la presente version 20050323
                'vlNumDias = 0
                vlMtoGarEstQui = 0
                vlMtoGarEstNor = 0
                
                'Seleccionar registros con Resolución de Garantía Estatal
                'registrada en la compañia
                
                vlCodTipRes = ""
                vlNumResGarEst = 0
                vlNumAnnoRes = 0
                
                vgSql = ""
                vgSql = "SELECT g.num_poliza,g.num_endoso,g.num_orden, "
                vgSql = vgSql & "g.num_resgarest , g.num_annores, g.cod_tipres "
                vgSql = vgSql & "FROM pp_tmae_garestres g "
                vgSql = vgSql & "WHERE "
                'Sacar Ini
'                vgSql = vgSql & "num_poliza in ('0000020436') AND "
                'Sacar Ini
                If vgTipoBase = "ORACLE" Then
                   vgSql = vgSql & "substr(fec_inires,1,6) <= '" & (vlNumPeriodo) & "' AND "
                   vgSql = vgSql & "substr(fec_terres,1,6) >= '" & (vlNumPeriodo) & "' "
                Else
                    vgSql = vgSql & "substring(fec_inires,1,6) <= '" & (vlNumPeriodo) & "' AND "
                    vgSql = vgSql & "substring(fec_terres,1,6) >= '" & (vlNumPeriodo) & "' "
                End If
'                vgSql = vgSql & "fec_inires <= '" & Trim(vlFecPago) & "' AND "
'                vgSql = vgSql & "fec_terres >= '" & Trim(vlFecPago) & "' "
                
                Set vlRegistroRes = vgConexionBD.Execute(vgSql)
                If Not vlRegistroRes.EOF Then
                    While Not vlRegistroRes.EOF
                
                        vlCodTipRes = (vlRegistroRes!COD_TIPRES)
                        vlNumResGarEst = (vlRegistroRes!NUM_RESGAREST)
                        vlNumAnnoRes = (vlRegistroRes!NUM_ANNORES)
                        vlNumPoliza = (vlRegistroRes!num_poliza)
                        vlNumOrden = (vlRegistroRes!Num_Orden)
                        vlNumEndoso = (vlRegistroRes!num_endoso)
                        
                        vlFecPago = ""
                        
                        vlMtoPension = 0
                        vlMtoPensionUF = 0
                        vlMtoPensionMin = 0
                        vlPrcDeduccion = 0
                        vlMtoGarEstCia = 0
                        vlMtoGarEstRec = 0
                        vlMtoDiferencia = 0
                        vlMtoHaber = 0
                        vlMtoDescuento = 0
                        vlCodDerGarEst = ""
                        vlCodEstado = ""
                        vlNumDias = 0
                 
                        'COMPAÑIA
                        'Seleccionar registros con pagos de garantia estatal por
                        'parte de la compañia
                        vgSql = ""
                        vgSql = "SELECT p.mto_conhabdes,p.num_perpago,l.fec_pago "
                        vgSql = vgSql & "FROM pp_ttmp_conpagopen p, pp_ttmp_conliqpagopen l "
                        vgSql = vgSql & "WHERE p.cod_conhabdes = '" & clCodConHDGECia & "' AND "
                        vgSql = vgSql & "p.num_poliza = '" & Trim(vlNumPoliza) & "' AND "
                        vgSql = vgSql & "p.num_orden = " & Str(vlNumOrden) & " AND "
                        vgSql = vgSql & "p.num_perpago = l.num_perpago AND "
                        vgSql = vgSql & "p.num_poliza = l.num_poliza AND "
                        vgSql = vgSql & "p.num_orden = l.num_orden AND "
                        vgSql = vgSql & "p.rut_receptor = l.rut_receptor AND "
                        vgSql = vgSql & "p.cod_tipreceptor = l.cod_tipreceptor AND "
                        'vgSql = vgSql & "l.fec_pago = '" & vlFecPago & "' AND "
                        vgSql = vgSql & "l.num_perpago = '" & Trim(vlNumPeriodo) & "' "
                        'vgSql = vgSql & "l.cod_tipopago = '" & vlPago & "' "
                        vgSql = vgSql & "ORDER BY p.num_poliza,p.num_orden "
                        Set vlRegistroCia = vgConexionBD.Execute(vgSql)
                        If Not vlRegistroCia.EOF Then
                            vlFecPago = (vlRegistroCia!Fec_Pago)
                            vlMtoGarEstCia = (vlRegistroCia!Mto_ConHabDes)
                        End If
                
                        'ESTADO
                        'Seleccionar datos de pago de garantia estatal
                        'informados por el estado
                        vgSql = ""
                        vgSql = "SELECT t.num_diapago,t.mto_garestpagtes,t.cod_estado "
                        vgSql = vgSql & "FROM pp_tmae_garestpagtesoro t "
                        vgSql = vgSql & "WHERE "
                        vgSql = vgSql & "t.num_perpago = '" & Trim(vlNumPeriodo) & "' AND "
                        vgSql = vgSql & "t.num_poliza = '" & Trim(vlNumPoliza) & "' AND "
                        vgSql = vgSql & "t.num_orden = " & Str(vlNumOrden) & " "
                        vgSql = vgSql & "ORDER BY t.num_poliza,t.num_orden "
                        Set vlRegistroTes = vgConexionBD.Execute(vgSql)
                        If Not vlRegistroTes.EOF Then
                            vlNumDias = (vlRegistroTes!num_diapago)
                            vlCodEstado = (vlRegistroTes!Cod_Estado)
                            vlMtoGarEstRec = (vlRegistroTes!mto_garestpagtes)
                        End If
                        
                         'Insertar o Actualizar registro en la tabla temporal
                        Call flIngresarRegistro
                       
                        vlRegistroRes.MoveNext
                        
                    Wend
                End If
'                vlRegistroPeriodo.MoveNext
'        Wend
'
'    End If
    
Exit Function
Err_flCargaTablaTemporalConGE:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

'----GE CONCILIACION------------------------------------------
Function flIngresarRegistro()
On Error GoTo Err_flIngresarRegistro

    vlMtoCalculado = False

    'Buscar datos de estado de garantia estatal cia
    Call flBuscarCodEst(vlNumPoliza, vlNumOrden)
    'Buscar datos de porcentaje de deduccion
    Call flBuscarPrcDed(vlNumPoliza, vlNumOrden)
    'Buscar Valor de MtoPensionGe (En UF)
    Call flAgregarMontoPensionUfCGE
    'Buscar Valor de MtoPensionGE (En Pesos)
    Call flAgregarConHabDesCGE(clCodConHDMtoPen, vlMtoPension)
    'Buscar Valor de MtoPensiónMinima
    Call flBuscarPenMin
    'Buscar otros Haberes (G.E. RetroActiva)
    Call flAgregarConHabDesCGE(clCodConHDHab, vlMtoHaber)
    'Buscar otros Decuentos (Descuentos al Pensionado por concepto de G.E.)
    Call flAgregarConHabDesCGE(clCodConHDDes, vlMtoDescuento)
    
    If vlPrcDeduccion > 0 Then
        vlMtoDeduccion = Format(((vlMtoPensionMin - vlMtoPension) * vlPrcDeduccion) / 100, "#0")
    Else
        vlMtoDeduccion = 0
    End If
    

'Confirmar si el registro a insertar se encuentra registrado
    vgSql = ""
    vgSql = "SELECT num_poliza,mto_garestcia,mto_garestrec "
    vgSql = vgSql & "FROM pp_ttmp_conconcigarest "
    vgSql = vgSql & "WHERE cod_usuario = '" & vgUsuario & "' AND "
    vgSql = vgSql & "cod_tipoimp = '" & vlCodTipoImp & "' AND "
    vgSql = vgSql & "num_periodo = '" & Trim(vlNumPeriodo) & "' AND "
    'vgSql = vgSql & "fec_pago = '" & vlFecPago & "' AND "
    vgSql = vgSql & "num_poliza = '" & Trim(vlNumPoliza) & "' AND "
    vgSql = vgSql & "num_orden = " & Str(vlNumOrden) & " "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If vgRegistro.EOF Then

        'Obtener monto de diferencia para insertar registro
        If (vlMtoGarEstCia <> 0) Or (vlMtoGarEstRec <> 0) Or _
           (vlMtoHaber <> 0) Or (vlMtoDescuento <> 0) Then
            vlMtoDiferencia = (((vlMtoGarEstCia + vlMtoHaber) - vlMtoDescuento) - vlMtoGarEstRec)
        Else
            vlMtoDiferencia = 0
        End If
        
        'Obtener monto total de G.E. pagada por la Cía. para insertar registro
        vlMtoGarEstCia = (vlMtoGarEstCia + vlMtoHaber) - vlMtoDescuento
        
        'Verificar valores de monto pension y monto deduccion segun derecho
        'Si el monto pension fue calculado por no tener derecho a GE
        If vlMtoGarEstCia = 0 Then
            vlMtoDeduccion = 0
        End If
        
        'Verificar Fecha de Pago
        'Si fecha de Pago = "" entonces se asignara por defecto el primer
        'dia del mes junto al periodo
        If vlFecPago = "" Then
            vlAnno = CInt(Mid(vlNumPeriodo, 1, 4))
            vlMes = CInt(Mid(vlNumPeriodo, 5, 2))
            vlDia = 1
            vlFecPago = Format(DateSerial(vlAnno, vlMes, vlDia), "yyyymmdd")
        End If
    
        'Insertar nuevo registro en la tabla temporal
        vgSql = ""
        vgSql = "INSERT INTO pp_ttmp_conconcigarest "
        vgSql = vgSql & "(cod_usuario,cod_tipoimp,fec_pago,num_periodo, "
        vgSql = vgSql & " num_poliza,num_orden,num_endoso,cod_tipres, "
        vgSql = vgSql & " num_resgarest,num_annores,num_dias,mto_pension, "
        vgSql = vgSql & " mto_pensionuf,mto_pensionmin,prc_deduccion, "
        vgSql = vgSql & " mto_deduccion,mto_garestqui,mto_garestnor, "
        vgSql = vgSql & " mto_garestcia,mto_garestrec,mto_diferencia, "
        vgSql = vgSql & " cod_estado,cod_dergarest,mto_haber,mto_descuento "
        vgSql = vgSql & " ) VALUES ( "
        vgSql = vgSql & " '" & vgUsuario & "', "
        vgSql = vgSql & " '" & Trim(vlCodTipoImp) & "' , "
        vgSql = vgSql & " '" & Trim(vlFecPago) & "' , "
        If vlNumPeriodo <> "" Then
            vgSql = vgSql & " '" & Trim(vlNumPeriodo) & "' , "
        Else
            vgSql = vgSql & " NULL , "
        End If
        vgSql = vgSql & " '" & Trim(vlNumPoliza) & "' , "
        vgSql = vgSql & " " & Str(vlNumOrden) & ", "
        If vlNumEndoso <> 0 Then
            vgSql = vgSql & " " & Str(vlNumEndoso) & ", "
        Else
            vgSql = vgSql & " NULL, "
        End If
        If vlCodTipRes <> "" Then
            vgSql = vgSql & " '" & Trim(vlCodTipRes) & "', "
        Else
            vgSql = vgSql & " NULL, "
        End If
        If vlNumResGarEst <> 0 Then
            vgSql = vgSql & " " & Str(vlNumResGarEst) & ", "
        Else
            vgSql = vgSql & " NULL, "
        End If
        If vlNumAnnoRes <> 0 Then
            vgSql = vgSql & " " & Str(vlNumAnnoRes) & ", "
        Else
            vgSql = vgSql & " NULL, "
        End If
        If vlNumDias <> 0 Then
            vgSql = vgSql & " " & Str(vlNumDias) & ", "
        Else
            vgSql = vgSql & " NULL, "
        End If
        If vlMtoPension <> 0 Then
            vgSql = vgSql & " " & Str(vlMtoPension) & ", "
        Else
            vgSql = vgSql & " NULL, "
        End If
        If vlMtoPensionUF <> 0 Then
            vgSql = vgSql & " " & Str(vlMtoPensionUF) & ", "
        Else
            vgSql = vgSql & " NULL, "
        End If
        If vlMtoPensionMin <> 0 Then
            vgSql = vgSql & " " & Str(vlMtoPensionMin) & ", "
        Else
            vgSql = vgSql & " NULL, "
        End If
        If vlPrcDeduccion <> 0 Then
            vgSql = vgSql & " " & Str(vlPrcDeduccion) & ", "
        Else
            vgSql = vgSql & " NULL, "
        End If
        If vlMtoDeduccion <> 0 Then
            vgSql = vgSql & " " & Str(vlMtoDeduccion) & ", "
        Else
            vgSql = vgSql & " NULL, "
        End If
        If vlMtoGarEstQui <> 0 Then
            vgSql = vgSql & " " & Str(vlMtoGarEstQui) & ", "
        Else
            vgSql = vgSql & " NULL, "
        End If
        If vlMtoGarEstNor <> 0 Then
            vgSql = vgSql & " " & Str(vlMtoGarEstNor) & ", "
        Else
            vgSql = vgSql & " NULL, "
        End If
        If vlMtoGarEstCia <> 0 Then
            vgSql = vgSql & " " & Str(vlMtoGarEstCia) & ", "
        Else
            vgSql = vgSql & " NULL, "
        End If
        If vlMtoGarEstRec <> 0 Then
            vgSql = vgSql & " " & Str(vlMtoGarEstRec) & ", "
        Else
            vgSql = vgSql & " NULL, "
        End If
        If vlMtoDiferencia <> 0 Then
            vgSql = vgSql & " " & Str(vlMtoDiferencia) & ", "
        Else
            vgSql = vgSql & " NULL, "
        End If
        If vlCodEstado <> "" Then
            vgSql = vgSql & " '" & Trim(vlCodEstado) & "', "
        Else
            vgSql = vgSql & " NULL, "
        End If
        If vlCodDerGarEst <> "" Then
            vgSql = vgSql & " '" & Trim(vlCodDerGarEst) & "', "
        Else
            vgSql = vgSql & " NULL, "
        End If

        If vlMtoHaber <> 0 Then
            vgSql = vgSql & " " & Str(vlMtoHaber) & ", "
        Else
            vgSql = vgSql & " NULL, "
        End If
        If vlMtoDescuento <> 0 Then
            vgSql = vgSql & " " & Str(vlMtoDescuento) & ") "
        Else
            vgSql = vgSql & " NULL) "
        End If
                
        vgConexionBD.Execute vgSql
    End If

Exit Function
Err_flIngresarRegistro:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

'----GE CONCILIACION------------------------------------------
Function flAgregarConHabDesCGE(iCodHabDes As String, monto As Double)
Dim vlValorUF As Double 'HQR 17/05/2005
On Error GoTo Err_flAgregarMontoPensionCGE

'Seleccionar Montos de Concepto de Garantìa Estatal
    vgSql = ""
    vgSql = "SELECT SUM (mto_conhabdes) as montohabdes "
    vgSql = vgSql & "FROM pp_ttmp_conpagopen "
    vgSql = vgSql & "WHERE cod_conhabdes IN " & iCodHabDes & " AND "
    vgSql = vgSql & "num_perpago = '" & Trim(vlNumPeriodo) & "' AND "
    vgSql = vgSql & "num_poliza = '" & Trim(vlNumPoliza) & "' AND "
    vgSql = vgSql & "num_orden = " & Str(vlNumOrden) & " "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       If Not IsNull(vgRegistro!montohabdes) Then
          monto = (vgRegistro!montohabdes)
       Else
           monto = 0
            If Trim(iCodHabDes) = clCodConHDMtoPen Then
                vlMtoCalculado = True
                'Obtener Tipo de Pago
                vgPalabra = ""
                vlTipoPago = flTipoPago(vlNumPeriodo, vlNumPoliza)
                If vlTipoPago = clTipoPagoR Then
                    vgPalabra = "p.fec_calpagoreg "
                Else
                    vgPalabra = "p.fec_calpripago "
                End If
                'Obtener Fecha de Calculo según tipo de pago
                vgSql = ""
                vgSql = "SELECT "
                vgSql = vgSql & vgPalabra & "as fecha "
                vgSql = vgSql & "FROM pp_tmae_propagopen p "
                vgSql = vgSql & "WHERE "
                vgSql = vgSql & "num_perpago >= '" & Trim(vlNumPeriodo) & "' AND "
                vgSql = vgSql & "num_perpago <= '" & Trim(vlNumPeriodo) & "' "
                Set vgRegistro = vgConexionBD.Execute(vgSql)
                If Not vgRegistro.EOF Then
                'Obtener Valor de UF según fecha de calculo
                    vlValorUF = flValorUF((vgRegistro!fecha))
                End If
                
                'Calcular Monto de Pensión en Pesos
                monto = (vlMtoPensionUF * vlValorUF)
            End If
        End If
    End If

Exit Function
Err_flAgregarMontoPensionCGE:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

'----GE CONCILIACION------------------------------------------
Function flTipoPago(ByVal iPeriodo As String, iNumPoliza As String) As String
On Error GoTo Err_flTipoPago
Dim iAnno As String
Dim iMes As String
Dim iDia As String
Dim iFechaInicio As String
Dim iFechaTermino As String

'    iFecha = Format(iFecha, "yyyymmdd")
    
'Calcula último día del mes
    iAnno = CInt(Mid(iPeriodo, 1, 4))
    iMes = CInt(Mid(iPeriodo, 5, 2))
    iDia = 1
    iFechaTermino = Format(DateSerial(iAnno, iMes + 1, iDia - 1), "yyyymmdd")
    
'Calcula Primer día del mes
    iAnno = CInt(Mid(iPeriodo, 1, 4))
    iMes = CInt(Mid(iPeriodo, 5, 2))
    iDia = 1
    iFechaInicio = Format(DateSerial(iAnno, iMes, iDia), "yyyymmdd")
            
    'iFecha = Mid(iFecha, 1, 6)
    
    'Estados del Pago de Pensión
    'PP: Primer Pago
    'PR: Pago en Regimen
    flTipoPago = ""
    
    'Determinar si el Caso es Primer Pago o Pago en Regimen
    vgSql = ""
    vgSql = "SELECT num_poliza,num_endoso "
    vgSql = vgSql & " FROM pp_tmae_poliza A WHERE "
    vgSql = vgSql & " num_poliza = '" & iNumPoliza & "' "
    vgSql = vgSql & " AND NUM_ENDOSO = "
    vgSql = vgSql & " (SELECT MAX(NUM_ENDOSO) FROM PP_TMAE_POLIZA " 'Último Endoso
    vgSql = vgSql & " WHERE NUM_POLIZA = A.NUM_POLIZA)"
    vgSql = vgSql & " AND FEC_INIPAGOPEN BETWEEN '" & iFechaInicio & "'"
    vgSql = vgSql & " AND '" & iFechaTermino & "'"
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If vgRegistro.EOF Then
        'Pago Régimen
        flTipoPago = "R"
    Else
        'Primer Pago
        flTipoPago = "P"
    End If
    vgRegistro.Close
    
Exit Function
Err_flTipoPago:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

'----GE CONCILIACION------------------------------------------
Function flValorUF(iFecha As String) As Double
On Error GoTo Err_flValorUF

    flValorUF = 0
    vgSql = ""
    vgSql = "SELECT m.mto_moneda "
    vgSql = vgSql & "FROM ma_tval_moneda m WHERE "
    vgSql = vgSql & "m.fec_moneda = '" & Trim(iFecha) & "' AND "
    vgSql = vgSql & "m.cod_moneda = '" & Trim(cgCodTipMonedaUF) & "' "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
        flValorUF = (vgRegistro!Mto_Moneda)
    End If

Exit Function
Err_flValorUF:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

'----GE CONCILIACION------------------------------------------
Function flAgregarMontoPensionUfCGE()
On Error GoTo Err_flAgregarMontoPensionUfCGE
    
'Seleccionar Monto de Pension en UF desde tabla Beneficiarios
    vlMtoPensionUF = 0
    
    vgSql = ""
    vgSql = "SELECT mto_pension,mto_pensiongar,fec_terpagopengar "
    vgSql = vgSql & "FROM PP_TMAE_BEN "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "num_poliza = '" & Trim(vlNumPoliza) & "' AND "
    vgSql = vgSql & "num_endoso = " & Trim(vlNumEndoso) & " AND "
    vgSql = vgSql & "num_orden = " & Str(vlNumOrden) & " "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       If IsNull(vgRegistro!Fec_TerPagoPenGar) Then
          vlMtoPensionUF = (vgRegistro!Mto_Pension)
       Else
           If vlPeriodo > Mid((vgRegistro!Fec_TerPagoPenGar), 1, 6) Then
              vlMtoPensionUF = (vgRegistro!Mto_Pension)
           Else
               If vlPeriodo <= Mid((vgRegistro!Fec_TerPagoPenGar), 1, 6) Then
                  vlMtoPensionUF = (vgRegistro!Mto_PensionGar)
               End If
           End If
       End If
    End If

Exit Function
Err_flAgregarMontoPensionUfCGE:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

'----GE CONCILIACION------------------------------------------
Function flBuscarPenMin()
On Error GoTo Err_flBuscarPenMin

    vgSql = ""
    vgSql = "SELECT cod_par,cod_sitinv,cod_sexo,fec_nacben "
    vgSql = vgSql & "FROM pp_tmae_ben "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "num_poliza = '" & Trim(vlNumPoliza) & "' AND "
    vgSql = vgSql & "num_endoso = " & Trim(vlNumEndoso) & " AND "
    vgSql = vgSql & "num_orden = " & Str(vlNumOrden) & " "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
        vlCodPar = (vgRegistro!Cod_Par)
        vlCodSexo = (vgRegistro!Cod_Sexo)
        vlCodSitInv = (vgRegistro!Cod_SitInv)
        vlFecNacBen = (vgRegistro!Fec_NacBen)
            
        vlEdadBen = fgCalculaEdad(vlFecNacBen, fgBuscaFecServ)
        vlEdadBen = fgConvierteEdadAños(vlEdadBen)
            
        vgSql = ""
        vgSql = "SELECT mto_penminfin "
        vgSql = vgSql & "FROM pp_tval_penminima "
        vgSql = vgSql & "WHERE "
        vgSql = vgSql & "cod_par = '" & Trim(vlCodPar) & "' AND "
        vgSql = vgSql & "cod_sitinv = '" & Trim(vlCodSitInv) & "' AND "
        vgSql = vgSql & "cod_sexo = '" & Trim(vlCodSexo) & "' AND "
        vgSql = vgSql & "num_edadini <= " & Str(vlEdadBen) & " AND "
        vgSql = vgSql & "num_edadfin >= " & Str(vlEdadBen) & " AND "
        If vgTipoBase = "ORACLE" Then
           vgSql = vgSql & "substr(fec_terpenmin,1,6) >= '" & (vlNumPeriodo) & "' AND "
        Else
            vgSql = vgSql & "substring(fec_terpenmin,1,6) >= '" & (vlNumPeriodo) & "' AND "
        End If
'        vgSql = vgSql & "fec_terpenmin >= '" & Trim(vlFecPago) & "' AND "
        vgSql = vgSql & "mto_penminfin > " & Str(vlMtoPension) & " "
        vgSql = vgSql & "ORDER BY mto_penminfin "
        Set vgRegistro = vgConexionBD.Execute(vgSql)
        If Not vgRegistro.EOF Then
            vlMtoPensionMin = (vgRegistro!mto_penminfin)
        Else
            vlMtoPensionMin = 0
            Exit Function
        End If
        
    End If
    
Exit Function
Err_flBuscarPenMin:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function
'----GE CONCILIACION------------------------------------------
Function flBuscarPrcDed(iPoliza As String, iOrden As Integer)
On Error GoTo Err_flBuscarPrcDed

'Seleccionar Datos de Porcentaje de Deducción
    vgSql = ""
    vgSql = "SELECT prc_dedtotal "
    vgSql = vgSql & "FROM pp_tmae_calporded "
    vgSql = vgSql & "WHERE num_poliza = '" & Trim(iPoliza) & "' AND "
    vgSql = vgSql & "num_orden = " & Str(iOrden) & " AND "
    If vgTipoBase = "ORACLE" Then
       vgSql = vgSql & "substr(fec_iniporded ,1,6) <= '" & (vlNumPeriodo) & "' AND "
       vgSql = vgSql & "substr(fec_terporded ,1,6) >= '" & (vlNumPeriodo) & "' "
    Else
        vgSql = vgSql & "substring(fec_iniporded ,1,6) <= '" & (vlNumPeriodo) & "' AND "
        vgSql = vgSql & "substring(fec_terporded ,1,6) >= '" & (vlNumPeriodo) & "' "
    End If

'    vgSql = vgSql & "fec_iniporded <= '" & Trim(vlFecPago) & "' AND "
'    vgSql = vgSql & "fec_terporded >= '" & Trim(vlFecPago) & "' "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
        vlPrcDeduccion = (vgRegistro!PRC_DEDTOTAL)
    End If

Exit Function
Err_flBuscarPrcDed:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

'----GE CONCILIACION------------------------------------------
Function flBuscarCodEst(iPoliza As String, iOrden As Integer)
On Error GoTo Err_flBuscarCodEst

'Seleccionar Datos de Porcentaje de Deducción
    vgSql = ""
    vgSql = "SELECT cod_dergarest "
    vgSql = vgSql & "FROM pp_tmae_garestestado "
    vgSql = vgSql & "WHERE num_poliza = '" & Trim(iPoliza) & "' AND "
    vgSql = vgSql & "num_orden = " & Str(iOrden) & " AND "
    If vgTipoBase = "ORACLE" Then
       vgSql = vgSql & "substr(fec_iniestgarest,1,6) <= '" & (vlNumPeriodo) & "' AND "
       vgSql = vgSql & "substr(fec_terestgarest,1,6) >= '" & (vlNumPeriodo) & "' "
    Else
        vgSql = vgSql & "substring(fec_iniestgarest,1,6) <= '" & (vlNumPeriodo) & "' AND "
        vgSql = vgSql & "substring(fec_terestgarest,1,6) >= '" & (vlNumPeriodo) & "' "
    End If
'    vgSql = vgSql & "fec_iniestgarest <= '" & Trim(vlFecPago) & "' AND "
'    vgSql = vgSql & "fec_terestgarest >= '" & Trim(vlFecPago) & "' "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
        vlCodDerGarEst = (vgRegistro!COD_DERGAREST)
    End If

Exit Function
Err_flBuscarCodEst:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function


'Function flProcesoConciliacionGE()
'vlmtocalculado = 0
'On Error GoTo Err_flInfConciliacion
'
'   vlCodTipoImp = clTipoImpC
'   vlPeriodo = Trim(Txt_Anno.Text + Txt_Mes.Text)
'   vlNumPerPago = vlPeriodo
'
'   vlSQL = "DELETE FROM pp_ttmp_garest WHERE cod_usuario = '" & vgUsuario & "' AND "
'   vlSQL = vlSQL & "cod_tipoimp = '" & vlCodTipoImp & "' "
'   vgConexionBD.Execute (vlSQL)
'
'   Call flCargaTablaTemporalConGE
'
'   Screen.MousePointer = 11
'
'   vlArchivo = strRpt & "PP_Rpt_GEConciliacionGECON.rpt"   '\Reportes
'   If Not fgExiste(vlArchivo) Then     ', vbNormal
'      MsgBox "Archivo de Reporte de Conciliación de Garantía Estatal no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
'      Screen.MousePointer = 0
'      Exit Function
'   End If
'
'   vgQuery = ""
'   vgQuery = "{PP_TTMP_CONCONCIGAREST.COD_USUARIO} = '" & Trim(vgUsuario) & "' AND "
'   vgQuery = vgQuery & "{PP_TTMP_CONCONCIGAREST.COD_TIPOIMP} = '" & Trim(vlCodTipoImp) & "' "
'
'   Rpt_Calculo.Reset
'   Rpt_Calculo.WindowState = crptMaximized
'   Rpt_Calculo.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
'   Rpt_Calculo.Connect = vgRutaDataBase       ' o App.Path & "\Nestle.mdb"
'   Rpt_Calculo.SelectionFormula = vgQuery
'
'   Rpt_Calculo.Formulas(0) = ""
'   Rpt_Calculo.Formulas(1) = ""
'   Rpt_Calculo.Formulas(2) = ""
'   Rpt_Calculo.Formulas(3) = ""
'
'   vgPalabra = Txt_Mes.Text + "-" + Txt_Anno.Text
'
'   Rpt_Calculo.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
'   Rpt_Calculo.Formulas(1) = "NombreSistema= '" & vgNombreSistema & "'"
'   Rpt_Calculo.Formulas(2) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
'   Rpt_Calculo.Formulas(3) = "Periodo = '" & vgPalabra & "'"
'
''   Rpt_Reporte.SubreportToChange = ""
'   Rpt_Calculo.Destination = crptToWindow
'   Rpt_Calculo.WindowState = crptMaximized
'   Rpt_Calculo.WindowTitle = ""
''   Rpt_Reporte.SelectionFormula = ""
'   Rpt_Calculo.Action = 1
'   Screen.MousePointer = 0
'
'Exit Function
'Err_flInfConciliacion:
'    Screen.MousePointer = 0
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
'
'End Function
'
''----GE CONCILIACION------------------------------------------
'Function flCargaTablaTemporalConGE()
'On Error GoTo Err_flCargaTablaTemporalConGE
'
'    vlCodTipoImp = ""
'    vlNumPeriodo = ""
'    vlFecPago = ""
'    vlCodTipRes = ""
'    vlNumResGarEst = 0
'    vlNumAnnoRes = 0
'    vlNumDias = 0
'    vlMtoPension = 0
'    vlMtoPensionUF = 0
'    vlMtoPensionMin = 0
'    vlPrcDeduccion = 0
'    vlMtoDeduccion = 0
'    vlMtoGarEstQui = 0
'    vlMtoGarEstNor = 0
'    vlMtoGarEstCia = 0
'    vlMtoGarEstRec = 0
'    vlMtoDiferencia = 0
'    vlCodDerGarEst = ""
'    vlMtoHaber = 0
'    vlMtoDescuento = 0
'
'    vlCodTipoImp = clTipoImpC
'    vlCodDerGarEst = clSinCodDerGE
'
'    vgPalabra = ""
'    If vlPago = "P" Then
'        vgPalabra = "p.fec_calpagoreg "
'    Else
'        vgPalabra = "p.fec_calpripago "
'    End If
'
'    'Seleccionar periodos en los que se realizaron pagos
'    'según fechas ingresadas por pantalla
'    vgSql = ""
'    vgSql = "SELECT "
'    vgSql = vgSql & vgPalabra & "as fecha "
'    vgSql = vgSql & "FROM pp_tmae_propagopen p "
'    vgSql = vgSql & "WHERE "
'    vgSql = vgSql & "num_perpago >= '" & Trim(vlNumPerPago) & "' AND "
'    vgSql = vgSql & "num_perpago <= '" & Trim(vlNumPerPago) & "' "
'    Set vlRegistroPeriodo = vgConexionBD.Execute(vgSql)
'    If Not vlRegistroPeriodo.EOF Then
'
'        While Not vlRegistroPeriodo.EOF
'                vlFecPago = (vlRegistroPeriodo!FECHA)
'                'Valores Nulos del registro, en la presente version 20050323
'                vlNumDias = 0
'                vlMtoGarEstQui = 0
'                vlMtoGarEstNor = 0
'
'                'COMPAÑIA
'                'Seleccionar registros con pagos de garantia estatal por
'                'parte de la compañia
'                vgSql = ""
'                vgSql = "SELECT p.num_poliza,l.num_endoso,p.num_orden, "
'                vgSql = vgSql & "p.mto_conhabdes,p.num_perpago,l.fec_pago "
'                vgSql = vgSql & "FROM pp_ttmp_conliqpagopen" & vlGlosaOpcion & " l, "
'                vgSql = vgSql & "pp_ttmp_conpagopen" & vlGlosaOpcion & " p "
'                vgSql = vgSql & "WHERE p.cod_conhabdes = '" & clCodConHDGECia & "' AND "
'                vgSql = vgSql & "p.num_perpago = l.num_perpago AND "
'                vgSql = vgSql & "p.num_poliza = l.num_poliza AND "
'                vgSql = vgSql & "p.num_orden = l.num_orden AND "
'                vgSql = vgSql & "p.rut_receptor = l.rut_receptor AND "
'                vgSql = vgSql & "p.cod_tipreceptor = l.cod_tipreceptor AND "
'                vgSql = vgSql & "l.num_perpago = '" & Trim(vlNumPerPago) & "' AND "
'                vgSql = vgSql & "l.cod_tipopago = '" & vlPago & "' "
'                vgSql = vgSql & "ORDER BY p.num_poliza,p.num_orden "
'                Set vlRegistroCia = vgConexionBD.Execute(vgSql)
'                If Not vlRegistroCia.EOF Then
'                    While Not vlRegistroCia.EOF
'
'                        vlFecPago = (vlRegistroCia!Fec_Pago)
'                        vlNumPeriodo = (vlRegistroCia!Num_PerPago)
'                        vlNumPoliza = (vlRegistroCia!Num_Poliza)
'                        vlNumEndoso = (vlRegistroCia!num_endoso)
'                        vlNumOrden = (vlRegistroCia!Num_Orden)
'                        vlMtoGarEstCia = (vlRegistroCia!Mto_ConHabDes)
'
'                        vlCodTipRes = ""
'                        vlNumResGarEst = 0
'                        vlNumAnnoRes = 0
'                        vlMtoPension = 0
'                        vlMtoPensionUF = 0
'                        vlMtoPensionMin = 0
'                        vlPrcDeduccion = 0
'                        vlMtoGarEstRec = 0
'                        vlMtoDiferencia = 0
'                        vlMtoHaber = 0
'                        vlMtoDescuento = 0
'                        vlCodDerGarEst = ""
'
'                        'Insertar o Actualizar registro en la tabla temporal
'                        vgPalabra = ""
'                        vgPalabra = "mto_garestcia = " & Str(vlMtoGarEstCia) & " "
'                        Call flIngresarRegistro
'
'                        vlRegistroCia.MoveNext
'                    Wend
'                End If
'
'                'ESTADO
'                'Seleccionar registros con pagos de garantia estatal
'                'informados por el estado
'                vgSql = ""
'                vgSql = "SELECT t.num_poliza,t.num_endoso,t.num_orden, "
'                vgSql = vgSql & "t.mto_garestpagtes,t.num_perpago,t.cod_estado,t.fec_pago "
'                vgSql = vgSql & "FROM pp_tmae_garestpagtesoro t "
'                vgSql = vgSql & "WHERE "
'                vgSql = vgSql & "t.num_perpago = '" & Trim(vlNumPerPago) & "' "
'                vgSql = vgSql & "ORDER BY t.num_poliza,t.num_orden "
'                Set vlRegistroTes = vgConexionBD.Execute(vgSql)
'                If Not vlRegistroTes.EOF Then
'                    While Not vlRegistroTes.EOF
'
'                        vlFecPago = (vlRegistroTes!Fec_Pago)
'                        vlNumPeriodo = (vlRegistroTes!Num_PerPago)
'                        vlNumPoliza = (vlRegistroTes!Num_Poliza)
'                        vlNumEndoso = (vlRegistroTes!num_endoso)
'                        vlNumOrden = (vlRegistroTes!Num_Orden)
'                        vlCodDerGarEst = (vlRegistroTes!Cod_Estado)
'                        vlMtoGarEstRec = (vlRegistroTes!mto_garestpagtes)
'
'                        vlCodTipRes = ""
'                        vlNumResGarEst = 0
'                        vlNumAnnoRes = 0
'                        vlMtoPension = 0
'                        vlMtoPensionUF = 0
'                        vlMtoPensionMin = 0
'                        vlPrcDeduccion = 0
'                        vlMtoDeduccion = 0
'                        vlMtoGarEstCia = 0
'                        vlMtoDiferencia = 0
'                        vlMtoHaber = 0
'                        vlMtoDescuento = 0
'
'                        'Insertar o Actualizar registro en la tabla temporal
'                        vgPalabra = ""
'                        vgPalabra = "mto_garestrec = " & Str(vlMtoGarEstRec) & " AND "
'                        vgPalabra = vgPalabra & "cod_estado = '" & Trim(vlCodDerGarEst) & "' "
'
'                        Call flIngresarRegistro
'
'                        vlRegistroTes.MoveNext
'                    Wend
'                End If
'
'            vlRegistroPeriodo.MoveNext
'
'        Wend
'
'    End If
'
'Exit Function
'Err_flCargaTablaTemporalConGE:
'    Screen.MousePointer = 0
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
'
'End Function
'
''----GE CONCILIACION------------------------------------------
'Function flIngresarRegistro()
'On Error GoTo Err_flIngresarRegistro
'
'    'Buscar datos de resolución
'    Call flBuscarResolucion(vlNumPoliza, vlNumOrden)
'    'Buscar datos de porcentaje de deduccion
'    Call flBuscarPrcDed(vlNumPoliza, vlNumOrden)
'    'Buscar Valor de MtoPensionGE (En Pesos)
'    Call flAgregarConHabDesCGE(clCodConHDMtoPen, vlMtoPension)
'    'Buscar Valor de MtoPensionGe (En UF)
'    Call flAgregarMontoPensionUfCGE
'    'Buscar Valor de MtoPensiónMinima
'    Call flBuscarPenMin
'    'Buscar otros Haberes (G.E. RetroActiva)
'    Call flAgregarConHabDesCGE(clCodConHDHab, vlMtoHaber)
'    'Buscar otros Decuentos (Descuentos al Pensionado por concepto de G.E.)
'    Call flAgregarConHabDesCGE(clCodConHDDes, vlMtoDescuento)
'
'    If vlPrcDeduccion > 0 Then
'        vlMtoDeduccion = (vlMtoPensionMin - vlMtoPension) * vlPrcDeduccion
'    Else
'        vlMtoDeduccion = 0
'    End If
'
''Confirmar si el registro se deberá insertar o modificar
'    vgSql = ""
'    vgSql = "SELECT num_poliza,mto_garestcia,mto_garestrec "
'    vgSql = vgSql & "FROM pp_ttmp_conconcigarest "
'    vgSql = vgSql & "WHERE cod_usuario = '" & vgUsuario & "' AND "
'    vgSql = vgSql & "cod_tipoimp = '" & vlCodTipoImp & "' AND "
'    vgSql = vgSql & "num_periodo = '" & Trim(vlNumPerPago) & "' AND "
'    vgSql = vgSql & "num_poliza = '" & Trim(vlNumPoliza) & "' AND "
'    vgSql = vgSql & "num_orden = " & Str(vlNumOrden) & " "
'    Set vgRegistro = vgConexionBD.Execute(vgSql)
'    If Not vgRegistro.EOF Then
'        'Obtener monto de diferencia para actualizacion
'        If ((vgRegistro!mto_garestcia) <> 0) Then
'            vlMtoGarEstCia = vlMtoGarEstCia + (vgRegistro!mto_garestcia)
'        End If
'        If ((vgRegistro!mto_garestrec) <> 0) Then
'            vlMtoGarEstRec = vlMtoGarEstRec + (vgRegistro!mto_garestrec)
'        End If
'        If ((vgRegistro!Mto_Haber) <> 0) Then
'            vlMtoHaber = vlMtoHaber + (vgRegistro!Mto_Haber)
'        End If
'        If ((vgRegistro!Mto_Descuento) <> 0) Then
'            vlMtoDescuento = vlMtoDescuento + (vgRegistro!Mto_Descuento)
'        End If
'
'        If ((vgRegistro!mto_garestcia) <> 0) Or ((vgRegistro!mto_garestrec) <> 0) Or _
'           ((vgRegistro!Mto_Haber) <> 0) Or ((vgRegistro!Mto_Descuento) <> 0) Then
'            vlMtoDiferencia = (((vlMtoGarEstCia + vlMtoHaber) - vlMtoDescuento) - vlMtoGarEstRec)
'        Else
'            vlMtoDiferencia = 0
'        End If
'
'
'        'Actualizar registro en la tabla temporal
'        vgSql = ""
'        vgSql = " UPDATE pp_ttmp_conconcigarest SET "
'        vgSql = vgSql & vgPalabra
'        If vlCodTipRes <> "" Then
'            vgSql = vgSql & "AND cod_tipres = '" & Trim(vlCodTipRes) & "' "
'        Else
'            vgSql = vgSql & "AND cod_tipres = NULL "
'        End If
'        If vlNumResGarEst <> 0 Then
'            vgSql = vgSql & "AND num_resgarest = " & Str(vlNumResGarEst) & " "
'        Else
'            vgSql = vgSql & "AND num_resgarest = NULL "
'        End If
'        If vlNumAnnoRes <> 0 Then
'            vgSql = vgSql & "AND num_annores = " & Str(vlNumAnnoRes) & " "
'        Else
'            vgSql = vgSql & "AND num_annores = NULL "
'        End If
'        If vlMtoPension <> 0 Then
'            vgSql = vgSql & "AND mto_pension = " & Str(vlMtoPension) & " "
'        Else
'            vgSql = vgSql & "AND mto_pension = NULL "
'        End If
'        If vlMtoPensionUF <> 0 Then
'            vgSql = vgSql & "AND mto_pensionuf = " & Str(vlMtoPensionUF) & " "
'        Else
'            vgSql = vgSql & "AND mto_pensionuf = NULL "
'        End If
'        If vlMtoPensionMin <> 0 Then
'            vgSql = vgSql & "AND mto_pensionmin = " & Str(vlMtoPensionMin) & " "
'        Else
'            vgSql = vgSql & "AND mto_pensionmin = NULL "
'        End If
'        If vlPrcDeduccion <> 0 Then
'            vgSql = vgSql & "AND prc_deduccion = " & Str(vlPrcDeduccion) & " "
'        Else
'            vgSql = vgSql & "AND prc_deduccion = NULL "
'        End If
'        If vlMtoDeduccion <> 0 Then
'            vgSql = vgSql & "AND mto_deduccion = " & Str(vlMtoDeduccion) & " "
'        Else
'            vgSql = vgSql & "AND mto_deduccion = NULL "
'        End If
'        If vlMtoDiferencia <> 0 Then
'            vgSql = vgSql & "AND mto_diferencia = " & Str(vlMtoDiferencia) & " "
'        Else
'            vgSql = vgSql & "AND mto_diferencia = NULL "
'        End If
'        If vlMtoGarEstCia <> 0 Then
'            vgSql = vgSql & "AND mto_garestcia = " & Str(vlMtoGarEstCia) & " "
'        Else
'            vgSql = vgSql & "AND mto_garestcia = NULL "
'        End If
'        If vlMtoGarEstRec <> 0 Then
'            vgSql = vgSql & "AND mto_garestrec = " & Str(vlMtoGarEstRec) & " "
'        Else
'            vgSql = vgSql & "AND mto_garestrec = NULL "
'        End If
'        If vlMtoHaber <> 0 Then
'            vgSql = vgSql & "AND mto_haber = " & Str(vlMtoHaber) & " "
'        Else
'            vgSql = vgSql & "AND mto_haber = NULL "
'        End If
'        If vlMtoDescuento <> 0 Then
'            vgSql = vgSql & "AND mto_descuento = " & Str(vlMtoDescuento) & " "
'        Else
'            vgSql = vgSql & "AND mto_descuento = NULL "
'        End If
'
'        vgSql = vgSql & "WHERE cod_usuario = '" & vgUsuario & "' AND "
'        vgSql = vgSql & "cod_tipoimp = '" & vlCodTipoImp & "' AND "
'        vgSql = vgSql & "num_periodo = '" & Trim(vlNumPerPago) & "' AND "
'        vgSql = vgSql & "num_poliza = '" & Trim(vlNumPoliza) & "' AND "
'        vgSql = vgSql & "num_orden = " & Str(vlNumOrden) & " "
'        vgConexionBD.Execute vgSql
'    Else
'        'Obtener monto de diferencia para insertar registro
'        If (vlMtoGarEstCia <> 0) Or (vlMtoGarEstRec <> 0) Or _
'           (vlMtoHaber <> 0) Or (vlMtoDescuento <> 0) Then
'            vlMtoDiferencia = (((vlMtoGarEstCia + vlMtoHaber) - vlMtoDescuento) - vlMtoGarEstRec)
'        Else
'            vlMtoDiferencia = 0
'        End If
'
'        'Insertar nuevo registro en la tabla temporal
'        vgSql = ""
'        vgSql = "INSERT INTO pp_ttmp_conconcigarest "
'        vgSql = vgSql & "(cod_usuario,cod_tipoimp,fec_pago,num_periodo, "
'        vgSql = vgSql & " num_poliza,num_orden,num_endoso,cod_tipres, "
'        vgSql = vgSql & " num_resgarest,num_annores,num_dias,mto_pension, "
'        vgSql = vgSql & " mto_pensionuf,mto_pensionmin,prc_deduccion, "
'        vgSql = vgSql & " mto_deduccion,mto_garestqui,mto_garestnor, "
'        vgSql = vgSql & " mto_garestcia,mto_garestrec,mto_diferencia, "
'        vgSql = vgSql & " cod_dergarest,mto_haber,mto_descuento "
'        vgSql = vgSql & " ) VALUES ( "
'        vgSql = vgSql & " '" & vgUsuario & "', "
'        vgSql = vgSql & " '" & Trim(vlCodTipoImp) & "' , "
'        vgSql = vgSql & " '" & Trim(vlFecPago) & "' , "
'        If vlNumPeriodo <> "" Then
'            vgSql = vgSql & " '" & Trim(vlNumPeriodo) & "' , "
'        Else
'            vgSql = vgSql & " NULL , "
'        End If
'        vgSql = vgSql & " '" & Trim(vlNumPoliza) & "' , "
'        vgSql = vgSql & " " & Str(vlNumOrden) & ", "
'        If vlNumEndoso <> 0 Then
'            vgSql = vgSql & " " & Str(vlNumEndoso) & ", "
'        Else
'            vgSql = vgSql & " NULL, "
'        End If
'        If vlCodTipRes <> "" Then
'            vgSql = vgSql & " '" & Trim(vlCodTipRes) & "', "
'        Else
'            vgSql = vgSql & " NULL, "
'        End If
'        If vlNumResGarEst <> 0 Then
'            vgSql = vgSql & " " & Str(vlNumResGarEst) & ", "
'        Else
'            vgSql = vgSql & " NULL, "
'        End If
'        If vlNumAnnoRes <> 0 Then
'            vgSql = vgSql & " " & Str(vlNumAnnoRes) & ", "
'        Else
'            vgSql = vgSql & " NULL, "
'        End If
'        If vlNumDias <> 0 Then
'            vgSql = vgSql & " " & Str(vlNumDias) & ", "
'        Else
'            vgSql = vgSql & " NULL, "
'        End If
'        If vlMtoPension <> 0 Then
'            vgSql = vgSql & " " & Str(vlMtoPension) & ", "
'        Else
'            vgSql = vgSql & " NULL, "
'        End If
'        If vlMtoPensionUF <> 0 Then
'            vgSql = vgSql & " " & Str(vlMtoPensionUF) & ", "
'        Else
'            vgSql = vgSql & " NULL, "
'        End If
'        If vlMtoPensionMin <> 0 Then
'            vgSql = vgSql & " " & Str(vlMtoPensionMin) & ", "
'        Else
'            vgSql = vgSql & " NULL, "
'        End If
'        If vlPrcDeduccion <> 0 Then
'            vgSql = vgSql & " " & Str(vlPrcDeduccion) & ", "
'        Else
'            vgSql = vgSql & " NULL, "
'        End If
'        If vlMtoDeduccion <> 0 Then
'            vgSql = vgSql & " " & Str(vlMtoDeduccion) & ", "
'        Else
'            vgSql = vgSql & " NULL, "
'        End If
'        If vlMtoGarEstQui <> 0 Then
'            vgSql = vgSql & " " & Str(vlMtoGarEstQui) & ", "
'        Else
'            vgSql = vgSql & " NULL, "
'        End If
'        If vlMtoGarEstNor <> 0 Then
'            vgSql = vgSql & " " & Str(vlMtoGarEstNor) & ", "
'        Else
'            vgSql = vgSql & " NULL, "
'        End If
'        If vlMtoGarEstCia <> 0 Then
'            vgSql = vgSql & " " & Str(vlMtoGarEstCia) & ", "
'        Else
'            vgSql = vgSql & " NULL, "
'        End If
'        If vlMtoGarEstRec <> 0 Then
'            vgSql = vgSql & " " & Str(vlMtoGarEstRec) & ", "
'        Else
'            vgSql = vgSql & " NULL, "
'        End If
'        If vlMtoDiferencia <> 0 Then
'            vgSql = vgSql & " " & Str(vlMtoDiferencia) & ", "
'        Else
'            vgSql = vgSql & " NULL, "
'        End If
'        If vlCodDerGarEst <> "" Then
'            vgSql = vgSql & " '" & Trim(vlCodDerGarEst) & "', "
'        Else
'            vgSql = vgSql & " NULL, "
'        End If
'        If vlMtoHaber <> 0 Then
'            vgSql = vgSql & " " & Str(vlMtoHaber) & ", "
'        Else
'            vgSql = vgSql & " NULL, "
'        End If
'        If vlMtoDescuento <> 0 Then
'            vgSql = vgSql & " " & Str(vlMtoDescuento) & ") "
'        Else
'            vgSql = vgSql & " NULL) "
'        End If
'
'        vgConexionBD.Execute vgSql
'    End If
'
'Exit Function
'Err_flIngresarRegistro:
'    Screen.MousePointer = 0
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
'
'End Function
'
''----GE CONCILIACION------------------------------------------
'Function flBuscarResolucion(iPoliza As String, iOrden As Integer)
'On Error GoTo Err_flBuscarResolucion
'
''Seleccionar Datos de Resolución
'    vgSql = ""
'    vgSql = "SELECT num_resgarest,num_annores,cod_tipres, "
'    vgSql = vgSql & "prc_deduccion "
'    vgSql = vgSql & "FROM pp_tmae_garestres "
'    vgSql = vgSql & "WHERE num_poliza = '" & Trim(iPoliza) & "' AND "
'    vgSql = vgSql & "num_orden = " & Str(iOrden) & " AND "
'    vgSql = vgSql & "fec_inires <= '" & Trim(vlFecPago) & "' AND "
'    vgSql = vgSql & "fec_terres >= '" & Trim(vlFecPago) & "' "
'    Set vgRegistro = vgConexionBD.Execute(vgSql)
'    If Not vgRegistro.EOF Then
'        vlCodTipRes = (vgRegistro!COD_TIPRES)
'        vlNumResGarEst = (vgRegistro!NUM_RESGAREST)
'        vlNumAnnoRes = (vgRegistro!NUM_ANNORES)
'    End If
'
'Exit Function
'Err_flBuscarResolucion:
'    Screen.MousePointer = 0
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
'
'End Function
'
''----GE CONCILIACION------------------------------------------
'Function flAgregarConHabDesCGE(iCodHabDes As String, monto As Double)
'On Error GoTo Err_flAgregarMontoPensionCGE
'
''Seleccionar Montos de Concepto de Garantìa Estatal
'    vgSql = ""
'    vgSql = "SELECT SUM (mto_conhabdes) as montohabdes "
'    vgSql = vgSql & "FROM pp_ttmp_pagopen" & vlGlosaOpcion & " "
'    vgSql = vgSql & "WHERE cod_conhabdes IN " & iCodHabDes & " AND "
'    vgSql = vgSql & "num_perpago = '" & Trim(vlNumPeriodo) & "' AND "
'    vgSql = vgSql & "num_poliza = '" & Trim(vlNumPoliza) & "' AND "
'    vgSql = vgSql & "num_orden = " & Str(vlNumOrden) & " "
'    Set vgRegistro = vgConexionBD.Execute(vgSql)
'    If Not vgRegistro.EOF Then
'       If Not IsNull(vgRegistro!montohabdes) Then
'          monto = (vgRegistro!montohabdes)
'       Else
'           monto = 0
'       End If
'    End If
'
'Exit Function
'Err_flAgregarMontoPensionCGE:
'    Screen.MousePointer = 0
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
'
'End Function
'
''----GE CONCILIACION------------------------------------------
'Function flAgregarMontoPensionUfCGE()
'On Error GoTo Err_flAgregarMontoPensionUfCGE
'
''Seleccionar Monto de Pension en UF desde tabla Beneficiarios
'    vlMtoPensionUF = 0
'
'    vgSql = ""
'    vgSql = "SELECT mto_pension,mto_pensiongar,fec_terpagopengar "
'    vgSql = vgSql & "FROM PP_TMAE_BEN "
'    vgSql = vgSql & "WHERE "
'    vgSql = vgSql & "num_poliza = '" & Trim(vlNumPoliza) & "' AND "
'    vgSql = vgSql & "num_endoso = " & Trim(vlNumEndoso) & " AND "
'    vgSql = vgSql & "num_orden = " & Str(vlNumOrden) & " "
'    Set vgRegistro = vgConexionBD.Execute(vgSql)
'    If Not vgRegistro.EOF Then
'       If IsNull(vgRegistro!Fec_TerPagoPenGar) Then
'          vlMtoPensionUF = (vgRegistro!Mto_Pension)
'       Else
'           If vlPeriodo > Mid((vgRegistro!Fec_TerPagoPenGar), 1, 6) Then
'              vlMtoPensionUF = (vgRegistro!Mto_Pension)
'           Else
'               If vlPeriodo < Mid((vgRegistro!Fec_TerPagoPenGar), 1, 6) Then
'                  vlMtoPensionUF = (vgRegistro!Mto_PensionGar)
'               End If
'           End If
'       End If
'    End If
'
'Exit Function
'Err_flAgregarMontoPensionUfCGE:
'    Screen.MousePointer = 0
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
'
'End Function
'
''----GE CONCILIACION------------------------------------------
'Function flBuscarPenMin()
'On Error GoTo Err_flBuscarPenMin
'
'    vgSql = ""
'    vgSql = "SELECT cod_par,cod_sitinv,cod_sexo,fec_nacben "
'    vgSql = vgSql & "FROM pp_tmae_ben "
'    vgSql = vgSql & "WHERE "
'    vgSql = vgSql & "num_poliza = '" & Trim(vlNumPoliza) & "' AND "
'    vgSql = vgSql & "num_endoso = " & Trim(vlNumEndoso) & " AND "
'    vgSql = vgSql & "num_orden = " & Str(vlNumOrden) & " "
'    Set vgRegistro = vgConexionBD.Execute(vgSql)
'    If Not vgRegistro.EOF Then
'        vlCodPar = (vgRegistro!Cod_Par)
'        vlCodSexo = (vgRegistro!Cod_Sexo)
'        vlCodSitInv = (vgRegistro!Cod_SitInv)
'        vlFecNacBen = (vgRegistro!Fec_NacBen)
'
'        vlEdadBen = fgCalculaEdad(vlFecNacBen, fgBuscaFecServ)
'        vlEdadBen = fgConvierteEdadAños(vlEdadBen)
'
'        vgSql = ""
'        vgSql = "SELECT mto_penminfin "
'        vgSql = vgSql & "FROM pp_tval_penminima "
'        vgSql = vgSql & "WHERE "
'        vgSql = vgSql & "cod_par = '" & Trim(vlCodPar) & "' AND "
'        vgSql = vgSql & "cod_sitinv = '" & Trim(vlCodSitInv) & "' AND "
'        vgSql = vgSql & "cod_sexo = '" & Trim(vlCodSexo) & "' AND "
'        vgSql = vgSql & "num_edadini <= " & Str(vlEdadBen) & " AND "
'        vgSql = vgSql & "num_edadfin >= " & Str(vlEdadBen) & " AND "
'        vgSql = vgSql & "fec_terpenmin >= '" & Trim(vlFecPago) & "' AND "
'        vgSql = vgSql & "mto_penminfin > " & Str(vlMtoPension) & " "
'        vgSql = vgSql & "ORDER BY mto_penminfin "
'        Set vgRegistro = vgConexionBD.Execute(vgSql)
'        If Not vgRegistro.EOF Then
'            vlMtoPensionMin = (vgRegistro!mto_penminfin)
'        Else
'            Exit Function
'        End If
'
'    End If
'
'Exit Function
'Err_flBuscarPenMin:
'    Screen.MousePointer = 0
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
'
'End Function
''----GE CONCILIACION------------------------------------------
'Function flBuscarPrcDed(iPoliza As String, iOrden As Integer)
'On Error GoTo Err_flBuscarPrcDed
'
''Seleccionar Datos de Porcentaje de Deducción
'    vgSql = ""
'    vgSql = "SELECT prc_dedtotal "
'    vgSql = vgSql & "FROM pp_tmae_calporded "
'    vgSql = vgSql & "WHERE num_poliza = '" & Trim(iPoliza) & "' AND "
'    vgSql = vgSql & "num_orden = " & Str(iOrden) & " AND "
'    vgSql = vgSql & "fec_iniporded <= '" & Trim(vlFecPago) & "' AND "
'    vgSql = vgSql & "fec_terporded >= '" & Trim(vlFecPago) & "' "
'    Set vgRegistro = vgConexionBD.Execute(vgSql)
'    If Not vgRegistro.EOF Then
'        vlPrcDeduccion = (vgRegistro!PRC_DEDTOTAL)
'    End If
'
'Exit Function
'Err_flBuscarPrcDed:
'    Screen.MousePointer = 0
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
'
'End Function



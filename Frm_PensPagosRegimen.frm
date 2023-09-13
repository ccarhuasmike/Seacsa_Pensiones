VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Frm_PensPagosRegimen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cálculo de Pagos Recurrentes"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   4725
   Begin VB.Frame Frame1 
      Caption         =   "Parámetros del Cálculo"
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
      Height          =   1935
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txt_FecPago 
         BackColor       =   &H00E0FFFF&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txt_FecCalculo 
         BackColor       =   &H00E0FFFF&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txt_UF 
         BackColor       =   &H00E0FFFF&
         Height          =   285
         Left            =   1920
         TabIndex        =   9
         Top             =   2160
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txt_TipoCalculo 
         BackColor       =   &H00E0FFFF&
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
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txt_Periodo 
         BackColor       =   &H00E0FFFF&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txt_FecProxPago 
         BackColor       =   &H00E0FFFF&
         Height          =   285
         Left            =   1920
         TabIndex        =   6
         Top             =   1800
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de Cálculo"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha de Pago"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de Cambio"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   2160
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo de Cálculo"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Período de Cálculo"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha Próximo Pago"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1800
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   4455
      Begin VB.CommandButton cmdErrTrans 
         Caption         =   "&Transac"
         Height          =   675
         Left            =   2880
         Picture         =   "Frm_PensPagosRegimen.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   1200
         Picture         =   "Frm_PensPagosRegimen.frx":06BA
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton cmd_calcular 
         Caption         =   "&Calcular"
         Height          =   675
         Left            =   360
         Picture         =   "Frm_PensPagosRegimen.frx":0D74
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Realizar Cálculo de Pensiones en Régimen"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   2040
         Picture         =   "Frm_PensPagosRegimen.frx":1216
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
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
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   4455
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   315
         Width           =   4170
         _ExtentX        =   7355
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblmensaje 
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   4095
      End
   End
End
Attribute VB_Name = "Frm_PensPagosRegimen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type TyBeneficiarios
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

End Type



Dim vlDiaAnteriorFecPago As String
Dim num_log_FlujoP As Long
Const clCodEstadoE As String * 1 = "E" 'Estado del Endoso (E:Endoso, P:PreEndoso)
'CORPTEC
Dim sTipoPro As String
Function flBorraPagosAnteriores(iPeriodo, iFecPago) As Boolean
On Error GoTo Errores
Dim vlSql As String

flBorraPagosAnteriores = False

'Elimina Detalle de Pagos Realizados anteriormente para el mismo Periodo
vlSql = "DELETE FROM PP_TMAE_PAGOPENPROV "
vlSql = vlSql & " WHERE NUM_PERPAGO = '" & iPeriodo & "'"
vlSql = vlSql & " AND NUM_POLIZA IN "
vlSql = vlSql & "(SELECT DISTINCT NUM_POLIZA FROM PP_TMAE_LIQPAGOPENPROV "
vlSql = vlSql & " WHERE NUM_PERPAGO = '" & iPeriodo & "'"
vlSql = vlSql & " AND COD_TIPOPAGO = 'R')" 'Pagos en Régimen
vgConexionTransac.Execute (vlSql)
'Call FgGuardaLog(vlSql, vgUsuario, "6088") 'RRR 25/07/2014
'Elimina Liquidaciones generadas anteriormente para el Mismo periodo

vlSql = "DELETE FROM PP_TMAE_LIQPAGOPENPROV"
vlSql = vlSql & " WHERE NUM_PERPAGO = '" & iPeriodo & "'"
vlSql = vlSql & " AND COD_TIPOPAGO = 'R'" 'Pagos en Régimen
vgConexionTransac.Execute (vlSql)
'Call FgGuardaLog(vlSql, vgUsuario, "6088") 'RRR 25/07/2014
'Elimina Datos Resumen
vlSql = "DELETE FROM pp_tmae_respagopenprov"
vlSql = vlSql & " WHERE num_perpago = '" & vgPerPago & "'"
vlSql = vlSql & " AND cod_tipopago = 'R'"
vgConexionTransac.Execute (vlSql)
'Call FgGuardaLog(vlSql, vgUsuario, "6088") 'RRR 25/07/2014
'If Me.Tag = "D" Then
    'Elimina Pensiones Actualizadas
    vlSql = "DELETE FROM pp_tmae_pensionact"
    vlSql = vlSql & " WHERE fec_desde = '" & vgPerPago & "01'" 'Primer Día del Mes
    vlSql = vlSql & " AND cod_tipopago = 'R'"
    vgConexionTransac.Execute (vlSql)
'End If
'Call FgGuardaLog(vlSql, vgUsuario, "6088") 'RRR 25/07/2014
'comente esta opcion por que no iba para el calculo
''Elimina Pensiones Actualizadas
'vlSql = "DELETE FROM pp_tmae_pensionact"
'vlSql = vlSql & " WHERE fec_desde = '" & vgPerPago & "01'" 'Primer Día del Mes
'vlSql = vlSql & " AND cod_tipopago = 'R'"
'vgConexionTransac.Execute (vlSql)


flBorraPagosAnteriores = True
Errores:
    If Err.Number <> 0 Then
        MsgBox "Se han producido errores al Eliminar los Datos de Cálculos Anteriores" & Chr(13) & Err.Description, vbCritical, Me.Caption
    End If
End Function

'RRR
Function flCreaTmpPagosProv(iFecPago) As Boolean
On Error GoTo Errores
Dim vlSql As String

flCreaTmpPagosProv = False

'Elimina Detalle de Pagos Realizados anteriormente para el mismo Periodo
vlSql = "DELETE FROM PP_TTMP_LIQPAGOPENPROV "
vgConexionTransac.Execute (vlSql)
 'Call FgGuardaLog(vlSql, vgUsuario, "6088") 'RRR 25/07/2014

vlSql = "INSERT INTO PP_TTMP_LIQPAGOPENPROV SELECT * FROM pp_tmae_liqpagopenprov"
vgConexionTransac.Execute (vlSql)
'Call FgGuardaLog(vlSql, vgUsuario, "6086") 'RRR 25/07/2014
 
flCreaTmpPagosProv = True
Errores:
    If Err.Number <> 0 Then
        MsgBox "Se han producido errores al Eliminar los Datos de Cálculos Anteriores" & Chr(13) & Err.Description, vbCritical, Me.Caption
    End If
End Function

Function flTraspasaDatosADefinitivos() As Boolean
On Error GoTo Errores
    Dim vlSql As String

    flTraspasaDatosADefinitivos = False
    'Verifica si existen Pensiones con monto Menor a 0
    vlSql = "SELECT COUNT(1) AS cont FROM pp_tmae_liqpagopenprov liq"
    vlSql = vlSql & " WHERE liq.num_perpago = '" & vgPerPago & "'"
    vlSql = vlSql & " AND liq.cod_tipopago = 'R'" 'Pago en Régimen
    vlSql = vlSql & " AND liq.mto_liqpagar < 0" 'Monto Pension Negativa
    Set vgRs = vgConexionTransac.Execute(vlSql)
    If Not vgRs.EOF Then
        'Pueden existir Pensiones con monto Menor a cero
        If vgRs!cont > 0 Then
            MsgBox "Cálculo de Pagos Recurrentes realizados exitosamente." & Chr(13) & Chr(13) & "Existen [" & vgRs!cont & "] Pensiones con Monto a Pagar Menor a Cero." & _
            Chr(13) & "Favor corregir estos casos y vuelva a realizar el Cálculo Definitivo.", vbCritical, Me.Caption
            flTraspasaDatosADefinitivos = True
            Exit Function
        End If
    End If
    
    'Elimina Datos Resumen
    vlSql = "DELETE FROM pp_tmae_respagopendef"
    vlSql = vlSql & " WHERE num_perpago = '" & vgPerPago & "'"
    vlSql = vlSql & " AND cod_tipopago = 'R'"
    vgConexionTransac.Execute (vlSql)
    'Call FgGuardaLog(vlSql, vgUsuario, "6088") 'RRR 25/07/2014
    'Elimina Datos anteriores Detalle
    vlSql = "DELETE FROM PP_TMAE_PAGOPENDEF"
    vlSql = vlSql & " WHERE NUM_PERPAGO = '" & vgPerPago & "'"
    vlSql = vlSql & " AND NUM_POLIZA IN"
    vlSql = vlSql & " (SELECT NUM_POLIZA FROM PP_TMAE_LIQPAGOPENDEF"
    vlSql = vlSql & " WHERE NUM_PERPAGO = '" & vgPerPago & "'"
    vlSql = vlSql & " AND COD_TIPOPAGO = 'R')" 'Pago en Régimen
    vgConexionTransac.Execute (vlSql)
    'Call FgGuardaLog(vlSql, vgUsuario, "6088") 'RRR 25/07/2014
    'Elimina Datos anteriores Liquidacion
    vlSql = "DELETE FROM PP_TMAE_LIQPAGOPENDEF"
    vlSql = vlSql & " WHERE NUM_PERPAGO = '" & vgPerPago & "'"
    vlSql = vlSql & " AND COD_TIPOPAGO = 'R'" 'Pago en Régimen
    vgConexionTransac.Execute (vlSql)
    'Call FgGuardaLog(vlSql, vgUsuario, "6088") 'RRR 25/07/2014
    'Traspasa Liquidacion
    vlSql = "INSERT INTO PP_TMAE_LIQPAGOPENDEF"
    vlSql = vlSql & " SELECT * FROM PP_TMAE_LIQPAGOPENPROV"
    vlSql = vlSql & " WHERE NUM_PERPAGO = '" & vgPerPago & "'"
    vlSql = vlSql & " AND COD_TIPOPAGO = 'R'" 'Pago en Régimen
    vgConexionTransac.Execute (vlSql)
    'Call FgGuardaLog(vlSql, vgUsuario, "6086") 'RRR 25/07/2014
    'Traspasa Detalle
    vlSql = "INSERT INTO PP_TMAE_PAGOPENDEF"
    vlSql = vlSql & " SELECT * FROM PP_TMAE_PAGOPENPROV"
    vlSql = vlSql & " WHERE NUM_PERPAGO = '" & vgPerPago & "'"
    vlSql = vlSql & " AND NUM_POLIZA IN "
    vlSql = vlSql & " (SELECT DISTINCT NUM_POLIZA FROM PP_TMAE_LIQPAGOPENPROV "
    vlSql = vlSql & " WHERE NUM_PERPAGO = '" & vgPerPago & "'"
    vlSql = vlSql & " AND COD_TIPOPAGO = 'R')" 'Pago en Régimen
    vgConexionTransac.Execute (vlSql)
    'Call FgGuardaLog(vlSql, vgUsuario, "6086") 'RRR 25/07/2014
    'Traspasa Resumen
    vlSql = "INSERT INTO pp_tmae_respagopendef" 'Pago en Régimen
    vlSql = vlSql & " SELECT * FROM pp_tmae_respagopenprov res"
    vlSql = vlSql & " WHERE res.num_perpago = '" & vgPerPago & "'"
    vlSql = vlSql & " AND res.cod_tipopago = 'R'" 'Pago en Régimen
    vgConexionTransac.Execute (vlSql)
    'Call FgGuardaLog(vlSql, vgUsuario, "6086") 'RRR 25/07/2014
    'Elimina Detalle Provisorio
    vlSql = " DELETE FROM PP_TMAE_PAGOPENPROV"
    vlSql = vlSql & " WHERE NUM_PERPAGO = '" & vgPerPago & "'"
    vlSql = vlSql & " AND NUM_POLIZA IN "
    vlSql = vlSql & " (SELECT DISTINCT NUM_POLIZA FROM PP_TMAE_LIQPAGOPENPROV "
    vlSql = vlSql & " WHERE NUM_PERPAGO = '" & vgPerPago & "'"
    vlSql = vlSql & " AND COD_TIPOPAGO = 'R')" 'Pago en Régimen
    vgConexionTransac.Execute (vlSql)
    'Call FgGuardaLog(vlSql, vgUsuario, "6088") 'RRR 25/07/2014
    'Elimina Liquidación Provisoria
    vlSql = " DELETE FROM PP_TMAE_LIQPAGOPENPROV"
    vlSql = vlSql & " WHERE NUM_PERPAGO = '" & vgPerPago & "'"
    vlSql = vlSql & " AND COD_TIPOPAGO = 'R'" 'Pago en Régimen
    vgConexionTransac.Execute (vlSql)
    'Call FgGuardaLog(vlSql, vgUsuario, "6088") 'RRR 25/07/2014
    'Elimina Resumen de Pagos
    vlSql = "DELETE FROM pp_tmae_respagopenprov"
    vlSql = vlSql & " WHERE NUM_PERPAGO = '" & vgPerPago & "'"
    vlSql = vlSql & " AND COD_TIPOPAGO = 'R'" 'Pago en Régimen
    vgConexionTransac.Execute (vlSql)
    'Call FgGuardaLog(vgSql, vgUsuario, "6088") 'RRR 25/07/2014
    'Cierra el Periodo de Cálculo
    If Not flCierraPeriodo Then
        flTraspasaDatosADefinitivos = False
    Else
        flTraspasaDatosADefinitivos = True
        MsgBox "Cálculo de Pagos en Régimen realizados exitosamente", vbInformation, Me.Caption
    End If

Errores:
    If Err.Number <> 0 Then
        MsgBox "Se ha producido un Error al Traspasar Datos a Histórico" & Chr(13) & Err.Description, vbCritical, Me.Caption
    End If
End Function

Private Function ValidaPagosPendientes() As Boolean
Dim rs As ADODB.Recordset
Dim cadena As String
    On Error GoTo mierror
    
    cadena = "select b.num_poliza,b.fec_inipagopen,b.num_orden"
    cadena = cadena & " FROM pp_tmae_certificado c inner join pp_tmae_ben b on c.num_poliza = b.num_poliza"
    cadena = cadena & " inner join pp_tmae_poliza a on a.num_poliza = b.num_poliza And a.num_endoso = b.num_endoso"
    cadena = cadena & " where  c.cod_indreliquidar<>'N' and c.num_reliq=0 and"
    cadena = cadena & " b.num_endoso = (SELECT MAX(NUM_ENDOSO) FROM pp_tmae_poliza WHERE NUM_POLIZA=b.num_poliza) and b.Cod_EstPension<>'10'"
    cadena = cadena & " and"
    cadena = cadena & " nvl(ceil(MONTHS_BETWEEN(TO_DATE(to_char(to_date('" & txt_FecCalculo.Text & "','dd/mm/yyyy'),'YYYYMM'),'YYYYMM'),TO_DATE(SUBSTR(b.fec_inipagopen,1,6),'YYYYMM'))) -"
    cadena = cadena & " nvl((select count(*) from pp_tmae_pagopendef x left join pp_tmae_ben c on x.num_poliza=c.num_poliza and x.num_orden=c.num_orden where x.num_poliza = b.num_poliza and"
    cadena = cadena & " x.cod_tipreceptor <> 'R' AND x.cod_conhabdes = '01' and c.num_endoso=b.num_endoso"
    cadena = cadena & " and SUBSTR(x.fec_inipago,1,6)>=SUBSTR(b.fec_inipagopen,1,6) and SUBSTR(x.fec_inipago,1,6)<=to_char(to_date('" & txt_FecCalculo.Text & "','dd/mm/yyyy'),'YYYYMM')"
    cadena = cadena & " group by c.fec_inipagopen),0)-"
    cadena = cadena & " nvl((select count(*) from pp_tmae_detcalcreliq where num_reliq=(select max(num_reliq) from pp_tmae_certificado where num_poliza=b.num_poliza)),0),-1)>0"
    cadena = cadena & " and c.fec_inicer<(select to_char(add_months(to_date(max(fec_calpagoreg), 'yyyymmdd'),1), 'yyyymmdd') from PP_TMAE_PROPAGOPEN where cod_estadoreg='C')"
    cadena = cadena & " and c.fec_tercer>(select to_char(add_months(to_date(max(fec_calpagoreg), 'yyyymmdd'),1), 'yyyymmdd') from PP_TMAE_PROPAGOPEN where cod_estadoreg='C') order by b.num_poliza"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open cadena, vgConexionBD, adOpenStatic, adLockReadOnly
            
    If Not rs.EOF Then
        ValidaPagosPendientes = True
    Else
        ValidaPagosPendientes = False
    End If
    
    Exit Function
mierror:
    MsgBox "No se pudo validar pagos pendientes"
End Function

Sub ImprimeMAy18(FecCalculo As String)
On Error GoTo Err_flInformeCerEst
'Certificados de Supervivencia
   Screen.MousePointer = 11
   
   Dim cadena, vlSql As String
   Dim objRep As New ClsReporte
   Dim vlRS18 As New ADODB.Recordset
   
   vlFechaTermino = FecCalculo 'Mid(FecCalculo, 7, 8) & "/" & Mid(FecCalculo, 5, 2) & "/" & Mid(FecCalculo, 1, 4)
   vgPalabra = ""
   vgPalabra = "Beneficiarios de Pensión Garantizada que cumplieron 18 años al " & vlFechaTermino
   
    If (fgCarga_Param("LI", "L18", FecCalculo) = True) Then
        L18 = vgValorParametro
    Else
        'vgError = 1000
        MsgBox "No existe Edad de tope para los 18 años.", vbCritical, "Proceso Cancelado"
        'Exit Function
    End If
    
    If (fgCarga_Param("LI", "L24", FecCalculo) = True) Then
        L24 = vgValorParametro
    Else
        'vgError = 1000
        MsgBox "No existe Edad de tope para los 18 años.", vbCritical, "Proceso Cancelado"
        'Exit Function
    End If
    
    'Mensualizar la Edad de 24 Años
    L18 = L18 * 12
    L24 = L24 * 12
   
   
'    vlSql = " SELECT p.num_poliza, num_orden, gls_tipoiden, num_idenben, gls_nomben , gls_nomsegben  , gls_patben  , gls_matben, fec_nacben  FROM pp_tmae_poliza p"
'    vlSql = vlSql & " JOIN pp_tmae_ben b ON p.num_poliza=b.num_poliza and p.num_endoso=b.num_endoso"
'    vlSql = vlSql & " join ma_tpar_tipoiden c on b.cod_tipoidenben=cod_tipoiden"
'    vlSql = vlSql & " WHERE p.num_endoso = (select max(num_endoso) from pp_tmae_poliza where num_poliza=p.num_poliza)"
'    vlSql = vlSql & " AND b.cod_estpension not in ('10','20') and b.cod_derpen not in (10,20)"
'    vlSql = vlSql & " AND p.num_mesgar > 0"
'    vlSql = vlSql & " AND p.fec_finpergar >= '" & FecCalculo & "'"
'    vlSql = vlSql & " AND b.fec_nacben < to_Char(add_months(to_date('" & FecCalculo & "','yyyymmdd'),-216),'yyyymmdd')"
'    vlSql = vlSql & " AND b.cod_par in ('30','35')"
'    vlSql = vlSql & " AND b.fec_inipagopen <= '" & FecCalculo & "'"
'    vlSql = vlSql & " AND p.cod_estado IN (6, 7, 8)"
'    vlSql = vlSql & " AND (p.fec_pripago < '" & FecCalculo & "' OR (p.num_mesdif > 0 AND p.fec_pripago = '" & FecCalculo & "'))"
   
'    vlSql = " SELECT p.num_poliza, num_orden, gls_tipoiden, num_idenben, gls_nomben , gls_nomsegben  , gls_patben  , gls_matben, fec_nacben"
'    vlSql = vlSql & " FROM pp_tmae_poliza p"
'    vlSql = vlSql & " JOIN pp_tmae_ben b ON p.num_poliza=b.num_poliza and p.num_endoso=b.num_endoso"
'    vlSql = vlSql & " join ma_tpar_tipoiden c on b.cod_tipoidenben=cod_tipoiden"
'    vlSql = vlSql & " WHERE p.num_endoso = (select max(num_endoso) from pp_tmae_poliza"
'    vlSql = vlSql & " where num_poliza=p.num_poliza) AND b.cod_estpension not in ('10','20') and b.cod_derpen not in ('10','20')"
'    vlSql = vlSql & " AND p.num_mesgar > 0 AND p.fec_finpergar >= '20131101' AND b.fec_nacben < to_Char(add_months(to_date('20131101','yyyymmdd'),-216),'yyyymmdd')"
'    vlSql = vlSql & " AND b.cod_par in ('30','35') AND b.fec_inipagopen <= '20131101'"
'    vlSql = vlSql & " AND p.cod_estado IN (6, 7, 8) and cod_sitinv not in ('T', 'P')"
'    vlSql = vlSql & " AND (p.fec_pripago < '20131101' OR (p.num_mesdif > 0 AND p.fec_pripago = '20131101'))"
'

    vlSql = " SELECT p.num_poliza, num_orden, gls_tipoiden, num_idenben, gls_nomben , gls_nomsegben  , gls_patben  , gls_matben, fec_nacben FROM pp_tmae_poliza p JOIN pp_tmae_ben b ON p.num_poliza=b.num_poliza and p.num_endoso=b.num_endoso join ma_tpar_tipoiden c on b.cod_tipoidenben=cod_tipoiden"
    vlSql = vlSql & " WHERE p.num_endoso = (select max(num_endoso) from pp_tmae_poliza where num_poliza=p.num_poliza)"
    vlSql = vlSql & " AND b.cod_estpension not in ('10','20') and b.cod_derpen not in (10,20) AND p.num_mesgar > 0 AND p.fec_finpergar >= '" & FecCalculo & "'"
    vlSql = vlSql & " AND b.fec_nacben < to_Char(add_months(to_date('" & FecCalculo & "','yyyymmdd'),-"
    vlSql = vlSql & " ("
    vlSql = vlSql & "                 select case when (fec_devsol is null) or (to_date(fec_devsol, 'YYYYMMDD') < to_date('20130801', 'YYYYMMDD')) then " & CStr(L18) & " else case when ("
    vlSql = vlSql & "                 select distinct case when IND_DNI='S' then 1 else 0 end  + case when IND_DJU='S' then 1 else 0 end  + case when IND_PES='S' then 1 else 0 end  + case when IND_BNO='S' then 1 else 0 end"
    vlSql = vlSql & "                 from pp_tmae_certificado where num_poliza=a.num_poliza and num_endoso=a.num_endoso and num_orden=ben.num_orden and cod_tipo='EST' AND EST_ACT=1"
    vlSql = vlSql & "                 ) = 4 then " & CStr(L24) & " else " & CStr(L18) & " end end as edad_l"
    vlSql = vlSql & "                 from pp_tmae_poliza a"
    vlSql = vlSql & "                 join pp_tmae_ben ben on a.num_poliza=ben.num_poliza and a.num_endoso=ben.num_endoso"
    vlSql = vlSql & "                 Where a.num_poliza = p.num_poliza And a.num_endoso = p.num_endoso And ben.Num_Orden = b.Num_Orden"
    vlSql = vlSql & " )"
    vlSql = vlSql & " ),'yyyymmdd') AND b.cod_par in ('30','35')"
    vlSql = vlSql & " AND b.fec_inipagopen <= '" & FecCalculo & "' AND p.cod_estado IN (6, 7, 8) and cod_sitinv not in ('T', 'P')"
    vlSql = vlSql & " AND (p.fec_pripago < '" & FecCalculo & "' OR (p.num_mesdif > 0 AND p.fec_pripago = '" & FecCalculo & "'))"


    'vlSql = " SELECT p.num_poliza, num_orden, gls_tipoiden, num_idenben, gls_nomben , gls_nomsegben  , gls_patben  , gls_matben, fec_nacben FROM pp_tmae_poliza p JOIN pp_tmae_ben b ON p.num_poliza=b.num_poliza and p.num_endoso=b.num_endoso join ma_tpar_tipoiden c on b.cod_tipoidenben=cod_tipoiden"
    'vlSql = vlSql & " WHERE p.num_endoso = (select max(num_endoso) from pp_tmae_poliza where num_poliza=p.num_poliza)"
    'vlSql = vlSql & " AND b.cod_estpension not in ('10','20') and b.cod_derpen not in (10,20) AND p.num_mesgar > 0 AND p.fec_finpergar >= '" & FecCalculo & "'"
    'vlSql = vlSql & " AND b.fec_nacben < to_Char(add_months(to_date('" & FecCalculo & "','yyyymmdd'),-"
    'vlSql = vlSql & " ("
    'vlSql = vlSql & "                 select case when (fec_devsol is null) or (to_date(fec_devsol, 'YYYYMMDD') < to_date('" & FecCalculo & "', 'YYYYMMDD')) then " & CStr(L18) & " else case when ("
    'vlSql = vlSql & "                 select case when IND_DNI='S' then 1 else 0 end  + case when IND_DJU='S' then 1 else 0 end  + case when IND_PES='S' then 1 else 0 end  + case when IND_BNO='S' then 1 else 0 end"
    'vlSql = vlSql & "                 from pp_tmae_certificado where num_poliza=a.num_poliza and num_endoso=a.num_endoso and num_orden=ben.num_orden"
    'vlSql = vlSql & "                 ) = 4 then " & CStr(L24) & " else " & CStr(L18) & " end end as edad_l"
    'vlSql = vlSql & "                 from pp_tmae_poliza a"
    'vlSql = vlSql & "                 join pp_tmae_ben ben on a.num_poliza=ben.num_poliza and a.num_endoso=ben.num_endoso"
    'vlSql = vlSql & "                 Where a.num_poliza = p.num_poliza And a.num_endoso = p.num_endoso And ben.Num_Orden = b.Num_Orden"
    'vlSql = vlSql & " )"
    'vlSql = vlSql & " ),'yyyymmdd') AND b.cod_par in ('30','35')"
    'vlSql = vlSql & " AND b.fec_inipagopen <= '" & FecCalculo & "' AND p.cod_estado IN (6, 7, 8) and cod_sitinv not in ('T', 'P')"
    'vlSql = vlSql & " AND (p.fec_pripago < '" & FecCalculo & "' OR (p.num_mesdif > 0 AND p.fec_pripago = '" & FecCalculo & "'))"
   
   ''reumnir solo un rato
'   Set vlRS18 = vgConexionBD.Execute(vlSql)
'   Dim LNGa As Long
'   LNGa = CreateFieldDefFile(vlRS18, Replace(UCase(strRpt & "Estructura\PP_Rpt_BenGarMay18.rpt"), ".RPT", ".TTX"), 1)
'
'
'    If objRep.CargaReporte(strRpt & "", "PP_Rpt_BenGarMay18.rpt", "Informe de Beneficiarios de Pensión Garantizada que cumplieron 18 años", vlRS18, True, _
'                            ArrFormulas("NombreCompania", vgNombreCompania), _
'                            ArrFormulas("NombreSistema", vgNombreSistema), _
'                            ArrFormulas("NombreSubSistema", vgNombreSubSistema), _
'                            ArrFormulas("fecha", vlFechaTermino)) = False Then
'
'        MsgBox "No se pudo abrir el reporte", vbInformation
'        Exit Sub
'    End If
    

    'fin marco

Exit Sub
Err_flInformeCerEst:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select


End Sub


Private Sub Cmd_Calcular_Click()
    Dim vlNumPensionesGarantizadas As Double
    'Realiza Cálculo de Pagos en Régimen
    On Error GoTo Errores
    
'    If ValidaPagosPendientes = True Then
'        MsgBox "Hay polizas pendientes de reliquidar", vbCritical, Me.Caption
'        Exit Sub
'    End If
    
    'por el momento
    'hqr 08/03/2011 Se agrega validacion de beneficiarios con Derecho a Pension Garantizada, que cumplieron los 18 años
    
    If Me.Tag <> "D" Then
        vlNumPensionesGarantizadas = flValidaPensionGarantizada(Format(vgFecIniPag, "yyyymmdd"))
        If vlNumPensionesGarantizadas > 0 Then
            MsgBox "Existen [" & vlNumPensionesGarantizadas & "] Beneficiarios con Pensión Garantizada, que cumplieron los 18 años." & Chr(13) & "Debe generar el endoso de estos beneficiarios.", vbCritical, Me.Caption
            ImprimeMAy18 (Format(vgFecIniPag, "yyyymmdd"))
            vgRes = MsgBox(" ¿ Desea Continuar con la generación de pago ?", 4 + 32 + 256, "Operación de pago de pensiones")
            If vgRes <> 6 Then
                Screen.MousePointer = 0
                Exit Sub
            End If
        ElseIf vlNumPensionesGarantizadas = -1 Then 'Error ya se mostró en la función
            Exit Sub
        End If
    End If
    
    'fin hqr 08/03/2011
    
    'Materia Gris - Jaime Rios 19/02/2018
    If flBuscarHijosMayoresPensionGarantizada Then
        'MsgBox ("CALCULAR PAGOS")
    End If
    
    'Obtiene Porcentaje Minimo de Salud
    If Not fgObtieneParametroVigencia("PS", "PSM", vgFecPago, stDatGenerales.Prc_SaludMin) Then
        MsgBox "Debe ingresar Porcentaje Mínimo de Salud", vbCritical, Me.Caption
        Exit Sub
    End If
    
    vlDiaAnteriorFecPago = Format(DateAdd("d", -1, DateSerial(Mid(vgFecPago, 1, 4), Mid(vgFecPago, 5, 2), Mid(vgFecPago, 7, 2))), "yyyymmdd")

    'Obtiene Valor US a Fecha de Pago (Para no estarla obteniendo por Cada Caso
    'hqr 12/12/2007 Se comenta a petición de MChirinos
    stDatGenerales.Val_UF = 1 'Se deja por Defecto Tipo de Cambio = 1
    'If Not fgObtieneConversion(vgFecPago, "US", stDatGenerales.Val_UF) Then
'    If Not fgObtieneConversion(vlDiaAnteriorFecPago, "US", stDatGenerales.Val_UF) Then
'        MsgBox "Debe ingresar el Tipo de Cambio de la Moneda 'US' a la Fecha de Pago", vbCritical, "Falta Tipo de Cambio"
'        Exit Sub
'    End If
    'fin hqr 12/12/2007 Se comenta a petición de MChirinos
    
    'CORPTEC
    'CORPTEC
    sTipoPro = "I"
    flEst_Proc
    
    If Me.Tag = "D" Then
        vgRes = MsgBox(" ¿ Desea Cerrar la Planilla Definitiva ?", 4 + 32 + 256, "Operación de pago de pensiones")
        If vgRes <> 6 Then
            Screen.MousePointer = 0
            Exit Sub
        End If
        flCierraPeriodoDefinitivo
    Else
        flCalcularPagoEnRegimen
    End If
    
    'CORPTEC
    sTipoPro = "F"
    flEst_Proc
    
Errores:
    Screen.MousePointer = 0
    If Err.Number <> 0 Then
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End If
End Sub
'CORPTEC
Function flEst_Proc() As Boolean
    Dim com As ADODB.Command
   Dim objRs As ADODB.Recordset
    Dim sistema, modulo, opcion, origen, tipo As String
    sistema = "SEACSA"
    modulo = "PENSIONES"
    opcion = "PAGOS"
    origen = "A"
    tipo = "A"
   ' Estado = "I"
    Set com = New ADODB.Command
    Set objRs = New ADODB.Recordset
    
    vgConexionBD.BeginTrans
    com.ActiveConnection = vgConexionBD
    com.CommandText = "SP_LOG_CARGAPROCESO"
    com.CommandType = adCmdStoredProc
    
    com.Parameters.Append com.CreateParameter("ESTADO", adChar, adParamInput, 1, sTipoPro)
    com.Parameters.Append com.CreateParameter("USUARIO", adVarChar, adParamInput, 10, vgLogin)
    com.Parameters.Append com.CreateParameter("SISTEMA", adVarChar, adParamInput, 50, sistema)
    com.Parameters.Append com.CreateParameter("MODULO", adVarChar, adParamInput, 50, modulo)
    com.Parameters.Append com.CreateParameter("OPCION", adVarChar, adParamInput, 50, opcion)
    com.Parameters.Append com.CreateParameter("ORIGEN", adChar, adParamInput, 1, origen)
    com.Parameters.Append com.CreateParameter("TIPO", adChar, adParamInput, 1, tipo)
    com.Parameters.Append com.CreateParameter("IDLOG", adDouble, adParamInput, 2, num_log_FlujoP)
    com.Parameters.Append com.CreateParameter("Retorno", adDouble, adParamReturnValue)
    com.Execute
    vgConexionBD.CommitTrans
    num_log_FlujoP = com("Retorno")
 
End Function

Private Sub Cmd_Imprimir_Click()
    If Trim(txt_FecPago) = "" Then
        MsgBox "Falta Ingresar Fecha Hasta", vbCritical, "Falta Información"
        Txt_Hasta.SetFocus
        Exit Sub
    End If
    
    On Error GoTo Err_flInformeCerEst
'Certificados de Supervivencia
   Screen.MousePointer = 11
   
   'marco 11/03/2010
   Dim cadena As String
   Dim objRep As New ClsReporte
   Dim vlFechaPago As String
   Dim rs As New ADODB.Recordset
   
   If Me.Tag = "P" Then
        vlFechaPago = Mid(Format(DateAdd("m", -1, CDate(Trim(txt_FecPago.Text))), "yyyymmdd"), 1, 6)
   Else
        vlFechaPago = Mid(Format(CDate(Trim(txt_FecPago.Text)), "yyyymmdd"), 1, 6)
   End If
   
   'vgPalabra = ""
   'vgPalabra = "Certificados vencidos al " & Txt_Hasta.Text
   
   
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    If Me.Tag = "P" Then
        rs.Open "PP_LISTA_PAGOS_PENSIONES_MXM.LISTAR(" & vlFechaPago & ")", vgConexionBD, adOpenForwardOnly, adLockReadOnly
    Else
        rs.Open "PP_LISTA_PAGOS_PENSIONES_MXM_D.LISTAR(" & vlFechaPago & ")", vgConexionBD, adOpenForwardOnly, adLockReadOnly
    End If
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PP_Rpt_PagosMxM.rpt"), ".RPT", ".TTX"), 1)
    
        
    If objRep.CargaReporte(strRpt & "", "PP_Rpt_PagosMxM.rpt", "Informe de Pagos Recurrentes (Diferencias con el mes pasado.) ", rs, True, _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("NombreSistema", vgNombreSistema), _
                            ArrFormulas("NombreSubSistema", vgNombreSubSistema)) = False Then
                            
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
    End If
    

    'fin marco

Exit Sub
Err_flInformeCerEst:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub

Private Sub cmd_salir_Click()
    Unload Me
End Sub

Private Sub cmdErrTrans_Click()
    frm_periodos.Show
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
End Sub

Function flCierraPeriodoDefinitivo()
On Error GoTo Errores

Dim com As New ADODB.Command  'Declaras un command
Dim tmpNoRF As String 'Por si necesitas que el sp retorne variable si no la quitas
Dim estadoCie As String
Screen.MousePointer = 11


vlSql = "SELECT COD_ESTADOREG FROM PP_TMAE_PROPAGOPEN WHERE NUM_PERPAGO='" & Mid(Format(vgFecIniPag, "yyyymmdd"), 1, 6) & "'"
Set vlTB = vgConexionBD.Execute(vlSql)
If Not vlTB.EOF Then
    estadoCie = vlTB!cod_estadoreg
    If estadoCie = "C" Then
        MsgBox "¡La planilla ya esta cerrada!.", vbInformation, Me.Caption
    End If
End If

Set com.ActiveConnection = vgConexionBD 'La asignas al command
com.CommandType = adCmdStoredProc 'Le dices que es de tipo Sp
com.CommandText = "PP_SPU_CIERREPROCESOPAGOPEN" 'Nombre del SP
com.Parameters.Append com.CreateParameter("vFECPROC", adVarChar, adParamInput, 8, Format(vgFecIniPag, "yyyymmdd"))
com.Execute  'Ejecutas el procedimiento

MsgBox "Cierre de Pagos en Régimen realizados exitosamente", vbInformation, Me.Caption
    
Errores:
    If Err.Number <> 0 Then
        If Err.Number <> 0 Then
            MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
        End If
    End If
    ProgressBar.Value = 0
    Frame3 = "Progreso del Cálculo"
    Screen.MousePointer = vbDefault
End Function


Function flCalcularPagoEnRegimen()
'Realiza el Cálculo de las Pensiones
On Error GoTo Errores
Dim vlSql As String
Dim vlTB As ADODB.Recordset 'para polizas
Dim vlcont As Double
Dim vlTB2 As ADODB.Recordset 'para beneficiarios
Dim vlTb3 As ADODB.Recordset 'para pension
Dim vlPension As Double 'Pension de RV o Garantizada
Dim vlBaseImp As Double, vlBaseTrib As Double 'Para calcular la Base Imponible y Base Tributable
Dim vlBaseLiq As Double 'Base Tributable - Impuesto Único
Dim vlMonto As Double
Dim vlDesSalud As Double 'Descuento de Salud
Dim bResp As Integer 'Retorno de las Funciones
Dim vlMontoPensionAct As Double 'Monto de la Pensión Actualizada
Dim vlMontoAcumPagado As Double 'Monto Acumulado por Poliza (para ver si cuadra con el Total por Beneficiarios)
Dim vlTCMonedaPension As Double 'Tipo de Cambio de la Moneda de la Pensión
Dim vlMonedaPension As String, vlMonedaPensionAnt As String 'Moneda de la Pensión (Para no buscar siempre el T.C., solo cuando cambia la moneda)
Dim vlEsGarantizado As Boolean, vlNumOrdenDiferencia As Long
Dim vlFactorAjuste As Double, vlFactorAjuste2 As Double
Dim vlMes As Long
Dim vlFechaAjuste As String
Dim vlRecibePension As Boolean
Dim vlMontoPensionActGar As Double 'RRR 25/10/2013
    
'hqr 05/09/2007 Agregados para Calculo de Primera Pensión Diferida
Dim vlPrimerMesAjuste As Boolean
Dim vlFactorAjusteDif As Double
Dim vlFactorAjusteDif2 As Double
Dim vlFechaIniDif As Date
Dim vlFechaFinDif As Date
Dim vlMesDif As Long, vlAñoDif As Long
Dim vlFecDesdeAjuste As String
'fin hqr 05/09/2007

Dim vlFactorAjusteDif_tmp As Double 'RRR 06/09/2019

'hqr 05/01/2008 para que limpie el grupo familiar, al cambiar de póliza o de Grupo Familiar
Dim vlPolizaAnt As Long
Dim vlGruFamAnt As Long
'fin hqr 05/01/2008

Dim vlFactorAjusteTasaFija As Double 'hqr 03/12/2010
Dim vlFactorAjusteporIPC As Double 'hqr 03/12/2010
Dim vlMesesDiferencia As Long 'hqr 18/02/2011

Dim vlNumPension As Double 'hqr 18/02/2011
Dim vlPasaTrimestre As Boolean 'hqr 18/02/2011

Dim vlExistePensionAct As Boolean ''RRR
Dim vlFecDev As String 'RRR 09/09/2016

Dim vlIndDes As String
Dim vlMontoPensionActTmp As Double

Dim vlMesProc As String
Dim vlAjuste As Integer

If Not fgConexionBaseDatos(vgConexionTransac) Then
    MsgBox "Error en Conexión a la Base de Datos", vbCritical, Me.Caption
    Exit Function
End If


If (fgCarga_Param("LI", "L18", vgFecPago) = True) Then
    L18 = vgValorParametro
Else
    'vgError = 1000
    MsgBox "No existe Edad de tope para los 18 años.", vbCritical, "Proceso Cancelado"
    'Exit Function
End If
    
If (fgCarga_Param("LI", "L24", vgFecPago) = True) Then
    L24 = vgValorParametro
Else
    'vgError = 1000
    MsgBox "No existe Edad de tope para los 18 años.", vbCritical, "Proceso Cancelado"
    'Exit Function
End If
    
'Mensualizar la Edad de 24 Años
L18 = L18 * 12
L24 = L24 * 12

Screen.MousePointer = 11
vgConexionTransac.BeginTrans



'------MG Jaime Rios 10/04/2018 inicio------
'-1.- Genera Endosos Automáticos para beneficiarios hijos
'If Not fgGenerarEndosoAutomatico(vlFactorAjuste, vlFactorAjusteTasaFija) Then
'    GoTo Deshacer
'Else
'    If vgGeneraEndosos = True Then 'Si se generó algun endoso se hace commit a la transacción
'        vgConexionTransac.CommitTrans
'        vgConexionTransac.BeginTrans
'    End If
'End If
'------MG Jaime Rios 10/04/2018 fin------

'Guarda los pagos anteriores
If Me.Tag = "D" Then
    If Not flCreaTmpPagosProv(vgFecPago) Then
        GoTo Deshacer
    Else
        vgConexionTransac.CommitTrans
        vgConexionTransac.BeginTrans
    End If
End If
'
'0.- Elimina Registros Anteriores
If Not flBorraPagosAnteriores(vgPerPago, vgFecPago) Then
    GoTo Deshacer
Else
    vgConexionTransac.CommitTrans
    vgConexionTransac.BeginTrans
End If




'0.1.- Obtiene Parametros de Quiebra (REVISAR si es por Periodo de Pago o Fecha de Pago)
stDatGenerales.Ind_AplicaQuiebra = fgObtieneParametrosQuiebra(vgPerPago, stDatGenerales.Prc_Castigo, stDatGenerales.Mto_TopMaxQuiebra)

'Llena Parte Invariable de la Estructura
stLiquidacion.fec_pago = vgFecPago
stLiquidacion.Num_PerPago = vgPerPago
stLiquidacion.Cod_TipoPago = "R"

'Cuenta Nº de Pólizas a Procesar
vlSql = "SELECT COUNT(1) AS contador FROM pp_tmae_poliza a"
vlSql = vlSql & " WHERE a.num_endoso ="
vlSql = vlSql & " (SELECT NVL(MAX(b.num_endoso + 1),1) AS num_endoso"
vlSql = vlSql & " FROM pp_tmae_endoso b"
vlSql = vlSql & " WHERE b.num_poliza = a.num_poliza"
vlSql = vlSql & " AND b.cod_estado = '" & clCodEstadoE & "'" 'Estado del Endoso = 'E'
vlSql = vlSql & " AND b.fec_efecto <= '" & Format(vgFecIniPag, "yyyymmdd") & "'"
vlSql = vlSql & " AND b.fec_finefecto >= '" & Format(vgFecIniPag, "yyyymmdd") & "')"
vlSql = vlSql & " AND a.cod_estado IN (6, 7, 8) " 'solo Pólizas Vigentes, con Beneficiarios Designados y Pendiente (Pago Diferido)
'Se deben identificar las Pólizas que tengan su Primer Pago en un Periodo Anterior (Inmediatas y Diferidas)
''vlSql = vlSql & " AND a.fec_inipagopen < '" & Format(vgFecIniPag, "yyyymmdd") & "'"
vlSql = vlSql & " AND (a.fec_pripago < '" & Format(vgFecIniPag, "yyyymmdd") & "'"
vlSql = vlSql & " OR (a.num_mesdif > 0 AND a.fec_pripago = '" & Format(vgFecIniPag, "yyyymmdd") & "'))"
Set vlTB = vgConexionBD.Execute(vlSql)
If Not vlTB.EOF Then
    vlcont = vlTB!contador
    If vlcont = 0 Then
        'Cierra el Periodo de Cálculo
        If Me.Tag = "D" Then
            If Not flCierraPeriodo Then
                GoTo Deshacer
            End If
        End If
        vgConexionTransac.CommitTrans
        vgConexionTransac.Close
        MsgBox "No existen Pólizas con Pago Recurrente para este período", vbCritical, Me.Caption
        Exit Function
    End If
Else
    MsgBox "No existen Pólizas con Pago Recurrente para este período", vbCritical, Me.Caption
    GoTo Deshacer
End If


'1.- Obtener último Endoso de Pólizas Vigentes
vlMonedaPensionAnt = ""

'hqr 05/01/2008 para que limpie el grupo familiar, al cambiar de póliza o de Grupo Familiar
vlPolizaAnt = -1
vlGruFamAnt = -1
'fin hqr 05/01/2008

vlMesProc = Mid(Format(vgFecIniPag, "yyyymmdd"), 5, 2)

Select Case vlMesProc
    Case "01"
        vlPasaTrimestre = True
    Case "04"
        vlPasaTrimestre = True
    Case "07"
        vlPasaTrimestre = True
    Case "10"
        vlPasaTrimestre = True
End Select

'/*Aqui debe ir el procedimiento*/


Dim com As New ADODB.Command  'Declaras un command
Dim tmpNoRF As String 'Por si necesitas que el sp retorne variable si no la quitas

Set com.ActiveConnection = vgConexionBD 'La asignas al command
com.CommandType = adCmdStoredProc 'Le dices que es de tipo Sp
com.CommandText = "PP_SPU_PROCESOPAGOPEN" 'Nombre del SP
com.Parameters.Append com.CreateParameter("vFECPROC", adVarChar, adParamInput, 8, Format(vgFecIniPag, "yyyymmdd"))
com.Execute  'Ejecutas el procedimiento

'--------------con la tabla resultante empiezo a pagar a todos los beneficiraios activos
vlSql = "select num_poliza, mto_pension, prc_factor, num_endoso, cod_moneda, fec_finpergar,cod_dergra, cod_isapre, cod_tippension "
vlSql = vlSql & " from pp_ttmp_pensionliq "
vlSql = vlSql & " where num_perpago='" & Mid(Format(vgFecIniPag, "yyyymmdd"), 1, 6) & "'"
'vlSql = vlSql & " and num_poliza in (6176,2127,2763,3264,5074,5418,6314,7104)"
'vlSql = vlSql & " and num_poliza in (576,769,1586,1726,1987,2457,2661,3054,3792,3825,4824,5689,6082,6933,7008,6662,6630,5857,2401,3191,2130)"
'vlSql = vlSql & " and num_poliza in (6604)"
Dim contador As Double

vlSql = vlSql & " order by 1"
Set vlTB = vgConexionBD.Execute(vlSql)
contador = 0
If Not vlTB.EOF Then
    vlAumento = (100 / vlcont)
    ProgressBar.Value = 0
    Frame3.Caption = "Progreso del Cálculo de Pensiones"
    ProgressBar.Refresh
    Me.Refresh
    stTutor.Cod_GruFami = "-1"
    Do While Not vlTB.EOF
        vlExistePensionAct = True
        'Registra Datos de la Poliza en la Estructura
        
        stLiquidacion.num_poliza = vlTB!num_poliza
        stLiquidacion.Cod_Moneda = vlTB!Cod_Moneda
        stLiquidacion.num_endoso = vlTB!num_endoso
        stLiquidacion.Cod_TipPension = vlTB!Cod_TipPension
        
        vlMontoPensionAct = Format(vlTB!Mto_Pension, "#0.00")
        vlMontoPensionActGar = Format(vlTB!Mto_Pension, "#0.00")
        
        vlMes = CInt(Mid(Format(vgFecIniPag, "yyyymmdd"), 5, 2))
        vlIndDes = vlTB!cod_isapre
        
        
        vlAjuste = IIf(vlTB!PRC_FACTOR <> 0, 2, 0)
        '***********************
    
        Dim Imod As String
        Dim prc_gar As Double
        'Obtiene la edad de la normativa menor a 28
        
        Imod = "SUP"
    
        'Obtiene Beneficiarios de la Póliza que tengan Derecho a Pensión
'        If vlTB!num_poliza = "0000004021" Then
'            a = 1
'        End If
        
        vlSql = "select distinct a.num_poliza, a.num_orden, mto_haber, b.*"
        vlSql = vlSql & " from pp_ttmp_pensionliq_ben a"
        vlSql = vlSql & " join pp_tmae_ben b on a.num_poliza=b.num_poliza and a.num_orden=b.num_orden"
        vlSql = vlSql & " where num_perpago='" & Mid(Format(vgFecIniPag, "yyyymmdd"), 1, 6) & "'"
        vlSql = vlSql & " and b.num_poliza='" & vlTB!num_poliza & "'"
        vlSql = vlSql & " and b.num_endoso=(select max(num_endoso) from pp_tmae_ben where num_poliza=b.num_poliza)"
        vlSql = vlSql & " order by 1,2"
        
        Set vlTB2 = vgConexionBD.Execute(vlSql)
        i = 0
        If Not vlTB2.EOF Then
            Do While Not vlTB2.EOF
                vlRecibePension = True
                bResp = fgCalculaEdad(vlTB2!Fec_NacBen, vgFecTerPag)
                If bResp = "-1" Then 'Error
                    GoTo Deshacer
                End If
             
                stDetPension.Edad = bResp
                stDetPension.EdadAños = fgConvierteEdadAños(stDetPension.Edad)
                'Si son Hijos se Calcula la Edad y se Verifica Certificado de Estudios RRR 20150218
                Imod = "SUP"
                If vlTB2!Cod_Par >= 30 And vlTB2!Cod_Par <= 35 And bResp >= L18 And vlTB2!Cod_SitInv = "N" Then 'Hijos
                    bResp = fgVerificaCertificadoEst(stLiquidacion.num_poliza, vlTB2!Num_Orden, stLiquidacion.num_endoso, vgFecIniPag)
                    If bResp = "-1" Then 'Error
                        GoTo Deshacer
                    Else
                        If bResp = "0" Then 'No tiene Certificado de estudios
                            vlRecibePension = False 'Va al Siguiente Beneficiario, ya que éste no tiene Derecho
                            
                            'ACTULIZA EL ESTADO DEL HIJO COMO PENDIENTE HASTA QUE REGULARICE
                            vlSql = "update pp_tmae_ben a set "
                            vlSql = vlSql & " Cod_EstPension = 10"
                            vlSql = vlSql & " where num_poliza='" & stLiquidacion.num_poliza & "'"
                            vlSql = vlSql & " and num_orden=" & vlTB2!Num_Orden
                            vlSql = vlSql & " and num_endoso=(select max(num_endoso) from pp_tmae_ben where num_poliza=a.num_poliza)"
                            vgConexionTransac.Execute (vlSql)

                        ElseIf bResp < 4 Then
                            vlRecibePension = False 'Va al Siguiente Beneficiario, no tiene completo los requeisitos para la excepcion de la mayoria de edad
                            'ACTULIZA EL ESTADO DEL HIJO COMO PENDIENTE HASTA QUE REGULARICE
                            vlSql = "update pp_tmae_ben a set "
                            vlSql = vlSql & " Cod_EstPension = 10"
                            vlSql = vlSql & " where num_poliza='" & stLiquidacion.num_poliza & "'"
                            vlSql = vlSql & " and num_orden=" & vlTB2!Num_Orden
                            vlSql = vlSql & " and num_endoso=(select max(num_endoso) from pp_tmae_ben where num_poliza=a.num_poliza)"
                            vgConexionTransac.Execute (vlSql)
                        End If
                    End If
                
                    'Imod = IIf(flObtieneEdadNormativa(vlTB!num_poliza) <= L18, "SUP", "EST")
                               
                    'If stDetPension.Edad >= flObtieneEdadNormativa(vlTB!num_poliza) And vlTB2!Cod_SitInv = "N" Then   'Hijos Sanos
                        'OBS: Se asume que el mes de los 18 años se paga completo
                        'Verifica Certificados de Estudio
                        'vlRecibePension = False 'No recibe pensión, por lo que no se envía al arreglo de Beneficiarios
                    'End If
                Else
                    'Valida Certificado de Supervivencia
                    bResp = fgVerificaCertificado(stLiquidacion.num_poliza, vlTB2!Num_Orden, vgFecIniPag, vgFecTerPag, Imod, "SUP")
                    If bResp = "-1" Then 'Error
                        GoTo Deshacer
                    Else
                        If bResp = "0" Then 'No tiene Certificado de Supervivencia
                            vlRecibePension = False 'Va al Siguiente Beneficiario, ya que éste no tiene Derecho
                        End If
                    End If
                End If
              
                If vlRecibePension Then 'Para los que tienen Derecho a Pensión Obtiene Monto de la Pensión
                    
                    vlPension = IIf(IsNull(vlTB2!Mto_Haber), 0, vlTB2!Mto_Haber)

                    ReDim Preserve stBeneficiarios(i)
                    stBeneficiarios(i).Mto_Pension = vlPension
                    vlMontoAcumPagado = vlMontoAcumPagado + vlPension
                    stBeneficiarios(i).num_poliza = vlTB!num_poliza
                    stBeneficiarios(i).num_endoso = vlTB!num_endoso
                    stBeneficiarios(i).Num_Orden = vlTB2!Num_Orden
                    stBeneficiarios(i).Cod_Par = vlTB2!Cod_Par
                    stBeneficiarios(i).Cod_GruFam = vlTB2!Cod_GruFam
                    stBeneficiarios(i).Cod_TipoIdenBen = vlTB2!Cod_TipoIdenBen
                    stBeneficiarios(i).Gls_MatBen = IIf(IsNull(vlTB2!Gls_MatBen), "", vlTB2!Gls_MatBen)
                    stBeneficiarios(i).Gls_PatBen = vlTB2!Gls_PatBen
                    stBeneficiarios(i).Gls_NomBen = vlTB2!Gls_NomBen
                    stBeneficiarios(i).Gls_NomSegBen = IIf(IsNull(vlTB2!Gls_NomSegBen), "", vlTB2!Gls_NomSegBen)
                    stBeneficiarios(i).Num_IdenBen = vlTB2!Num_IdenBen
                    stBeneficiarios(i).Gls_DirBen = vlTB2!Gls_DirBen
                    stBeneficiarios(i).Cod_Direccion = vlTB2!Cod_Direccion
                    stBeneficiarios(i).Cod_ViaPago = IIf(IsNull(vlTB2!Cod_ViaPago), "NULL", vlTB2!Cod_ViaPago)
                    stBeneficiarios(i).Cod_Banco = IIf(IsNull(vlTB2!Cod_Banco), "NULL", vlTB2!Cod_Banco)
                    stBeneficiarios(i).Cod_TipCuenta = IIf(IsNull(vlTB2!Cod_TipCuenta), "NULL", vlTB2!Cod_TipCuenta)
                    stBeneficiarios(i).Num_Cuenta = IIf(IsNull(vlTB2!Num_Cuenta), "NULL", vlTB2!Num_Cuenta)
                    stBeneficiarios(i).Cod_Sucursal = IIf(IsNull(vlTB2!Cod_Sucursal), "NULL", vlTB2!Cod_Sucursal)
                    stBeneficiarios(i).Fec_NacBen = vlTB2!Fec_NacBen
                    stBeneficiarios(i).Cod_SitInv = vlTB2!Cod_SitInv
                    stBeneficiarios(i).Fec_TerPagoPenGar = IIf(IsNull(vlTB2!Fec_TerPagoPenGar), "NULL", vlTB2!Fec_TerPagoPenGar)
                    stBeneficiarios(i).Prc_PensionGar = vlTB2!Prc_PensionGar
                    stBeneficiarios(i).Prc_Pension = vlTB2!Prc_Pension
                    stBeneficiarios(i).Cod_InsSalud = vlTB2!Cod_InsSalud
                    stBeneficiarios(i).Cod_ModSalud = vlTB2!Cod_ModSalud
                    stBeneficiarios(i).Mto_PlanSalud = vlTB2!Mto_PlanSalud
                    i = i + 1
                End If

                vlTB2.MoveNext
            Loop
            'Valida que el Monto de la Pensión sea el total garantizado
            If vlEsGarantizado Then
                'Validar si se pagó el 100%
                If vlMontoAcumPagado <> vlMontoPensionAct Then 'Se debe ajustar la diferencia
                    'stBeneficiarios(0).Mto_Pension = stBeneficiarios(0).Mto_Pension + (vlMontoPensionAct - vlMontoAcumPagado) 'Suma o Resta al Primer Beneficiario calculado
                End If
            End If
        End If
        If i > 0 Then 'Si se le pagó a alguien
            i = 0
            Do While i <= UBound(stBeneficiarios)
                'Llena Datos de la Estructura que no cambian
                stDetPension.Fec_IniPago = Format(vgFecIniPag, "yyyymmdd")
                stDetPension.Fec_TerPago = Format(vgFecTerPag, "yyyymmdd")
                stDetPension.num_endoso = vlTB!num_endoso
                stDetPension.Num_Orden = stBeneficiarios(i).Num_Orden
                stDetPension.num_poliza = vlTB!num_poliza
                stDetPension.Num_PerPago = vgPerPago
                'JVB 20210910
                bResp = fgCalculaEdad(stBeneficiarios(i).Fec_NacBen, vgFecTerPag)
                If bResp = "-1" Then 'Error
                    GoTo Deshacer
                End If
             
                stDetPension.Edad = bResp
                stDetPension.EdadAños = fgConvierteEdadAños(stDetPension.Edad)
                '
                    
                'hqr 05/01/2008 para que limpie el grupo familiar, al cambiar de póliza o de Grupo Familiar
                If stDetPension.num_poliza <> vlPolizaAnt Or vlGruFamAnt <> stBeneficiarios(i).Cod_GruFam Then
                    stTutor.Cod_GruFami = "-1"
                End If
                vlPolizaAnt = stDetPension.num_poliza
                vlGruFamAnt = stBeneficiarios(i).Cod_GruFam
                'fin hqr 05/01/2008
                
                'If vlPolizaAnt = "0000000513" Then MsgBox "0000000513"
                
                
                If stBeneficiarios(i).Cod_Par < 30 Then 'Padres quedan registrados para Tutores
                    stTutor.Cod_GruFami = stBeneficiarios(i).Cod_GruFam
                    stTutor.Cod_TipReceptor = "M"
                    stTutor.Cod_TipoIdenReceptor = stBeneficiarios(i).Cod_TipoIdenBen
                    stTutor.Gls_MatReceptor = stBeneficiarios(i).Gls_MatBen
                    stTutor.Gls_PatReceptor = stBeneficiarios(i).Gls_PatBen
                    stTutor.Gls_NomReceptor = stBeneficiarios(i).Gls_NomBen
                    stTutor.Gls_NomSegReceptor = stBeneficiarios(i).Gls_NomSegBen
                    stTutor.Num_IdenReceptor = stBeneficiarios(i).Num_IdenBen
                    stTutor.Gls_Direccion = stBeneficiarios(i).Gls_DirBen
                    stTutor.Cod_Direccion = stBeneficiarios(i).Cod_Direccion
                    stTutor.Cod_ViaPago = stBeneficiarios(i).Cod_ViaPago
                    stTutor.Cod_Banco = stBeneficiarios(i).Cod_Banco
                    stTutor.Cod_TipCuenta = stBeneficiarios(i).Cod_TipCuenta
                    stTutor.Num_Cuenta = stBeneficiarios(i).Num_Cuenta
                    stTutor.Cod_Sucursal = stBeneficiarios(i).Cod_Sucursal
                End If
                    
                'Inicializa Monto Haber y Descuento
                stLiquidacion.Mto_Haber = 0
                stLiquidacion.Mto_Descuento = 0
                stLiquidacion.Num_Orden = stBeneficiarios(i).Num_Orden
                
                '11.-  Obtener Tutores (1a. Etapa) (Se deja acá porque se necesita el Rut del Receptor)
                bResp = fgObtieneTutor(stLiquidacion.num_poliza, stLiquidacion.num_endoso, stLiquidacion.Num_Orden, vgFecIniPag, vgFecTerPag, stLiquidacion)
                If bResp = "-1" Then 'Error
                    GoTo Deshacer
                Else
                    If bResp = "0" Then 'No Encontró Tutor
                        If stBeneficiarios(i).Cod_Par >= 30 And stBeneficiarios(i).Cod_Par <= 35 And stDetPension.Edad <= stDatGenerales.MesesEdad18 And stTutor.Cod_GruFami = stBeneficiarios(i).Cod_GruFam Then 'El Tutor debe ser la Madre
                            stLiquidacion.Cod_TipReceptor = stTutor.Cod_TipReceptor 'MADRE
                            stLiquidacion.Num_IdenReceptor = stTutor.Num_IdenReceptor
                            stLiquidacion.Cod_TipoIdenReceptor = stTutor.Cod_TipoIdenReceptor
                            stLiquidacion.Gls_NomReceptor = stTutor.Gls_NomReceptor
                            stLiquidacion.Gls_NomSegReceptor = stTutor.Gls_NomSegReceptor
                            stLiquidacion.Gls_PatReceptor = stTutor.Gls_PatReceptor
                            stLiquidacion.Gls_MatReceptor = stTutor.Gls_MatReceptor
                            stLiquidacion.Gls_Direccion = stTutor.Gls_Direccion
                            stLiquidacion.Cod_Direccion = stTutor.Cod_Direccion
                            stLiquidacion.Cod_ViaPago = stTutor.Cod_ViaPago
                            stLiquidacion.Cod_Banco = stTutor.Cod_Banco
                            stLiquidacion.Cod_TipCuenta = stTutor.Cod_TipCuenta
                            stLiquidacion.Num_Cuenta = stTutor.Num_Cuenta
                            stLiquidacion.Cod_Sucursal = stTutor.Cod_Sucursal
                            
                        Else 'Else se le Pagará a El Mismo
                            stLiquidacion.Cod_TipReceptor = "P" 'Causante
                            stLiquidacion.Num_IdenReceptor = stBeneficiarios(i).Num_IdenBen
                            stLiquidacion.Cod_TipoIdenReceptor = stBeneficiarios(i).Cod_TipoIdenBen
                            stLiquidacion.Gls_NomReceptor = stBeneficiarios(i).Gls_NomBen
                            stLiquidacion.Gls_NomSegReceptor = stBeneficiarios(i).Gls_NomSegBen
                            stLiquidacion.Gls_PatReceptor = stBeneficiarios(i).Gls_PatBen
                            stLiquidacion.Gls_MatReceptor = stBeneficiarios(i).Gls_MatBen
                            stLiquidacion.Gls_Direccion = stBeneficiarios(i).Gls_DirBen
                            stLiquidacion.Cod_Direccion = stBeneficiarios(i).Cod_Direccion
                            stLiquidacion.Cod_ViaPago = stBeneficiarios(i).Cod_ViaPago
                            stLiquidacion.Cod_Banco = stBeneficiarios(i).Cod_Banco
                            stLiquidacion.Cod_TipCuenta = stBeneficiarios(i).Cod_TipCuenta
                            stLiquidacion.Num_Cuenta = stBeneficiarios(i).Num_Cuenta
                            stLiquidacion.Cod_Sucursal = stBeneficiarios(i).Cod_Sucursal
                        End If
                    'Else 'Encontró Tutor
                    End If
                End If
                    
                stDetPension.Num_IdenReceptor = stLiquidacion.Num_IdenReceptor
                stDetPension.Cod_TipoIdenReceptor = stLiquidacion.Cod_TipoIdenReceptor
                stDetPension.Cod_TipReceptor = stLiquidacion.Cod_TipReceptor
                stDetPension.Cod_ConHabDes = stDatGenerales.Cod_ConceptoPension
                
                stLiquidacion.Mto_Pension = stBeneficiarios(i).Mto_Pension
                stDetPension.Mto_ConHabDes = stLiquidacion.Mto_Pension
                stLiquidacion.Mto_Haber = stLiquidacion.Mto_Haber + stLiquidacion.Mto_Pension
                'Graba Monto de la Pensión
                If Not fgInsertaDetallePensionProv(stDetPension) Then
                    MsgBox "Se ha producido un Error al Grabar Monto de la Pensión" & Chr(13) & Err.Description, vbCritical, Me.Caption
                    GoTo Deshacer
                End If
                vlBaseImp = stBeneficiarios(i).Mto_Pension
                If vlTB!Cod_DerGra = "S" Then 'Gratificación se paga en Julio y Diciembre
                    If (vlMes = 7 Or vlMes = 12) Then
                        stDetPension.Mto_ConHabDes = stLiquidacion.Mto_Pension
                        stLiquidacion.Mto_Haber = stLiquidacion.Mto_Haber + stLiquidacion.Mto_Pension
                        stDetPension.Cod_ConHabDes = stDatGenerales.Cod_ConceptoGratificacion
                        'Graba Monto de la Pensión
                        If Not fgInsertaDetallePensionProv(stDetPension) Then
                            MsgBox "Se ha producido un Error al Grabar Monto de la Pensión" & Chr(13) & Err.Description, vbCritical, Me.Caption
                            GoTo Deshacer
                        End If
                        vlBaseImp = vlBaseImp + stDetPension.Mto_ConHabDes
                        stBeneficiarios(i).Mto_PlanSalud = stBeneficiarios(i).Mto_PlanSalud / 2
                    End If
                End If
                    
                '4.- Obtener Haberes y Descuentos Imponibles (1a. Etapa)
                If Not fgObtieneHaberesDescuentos(Me.Tag, vgPerPago, vlTB!num_poliza, vlTB!num_endoso, stBeneficiarios(i).Num_Orden, "S", "S", vlMonto, 0, 0, vlDiaAnteriorFecPago, stLiquidacion, stDetPension, stDatGenerales.Val_UF, vgFecIniPag, vgFecTerPag, vlTB!Cod_Moneda, vlTCMonedaPension) Then
                    GoTo Deshacer
                End If
                vlBaseImp = vlBaseImp + vlMonto 'Base Imponible
                stLiquidacion.Mto_BaseImp = vlBaseImp
                
                '5.- Calcular Descto. Salud (1a. Etapa)
                
                vlDesSalud = 0
                If vlIndDes = "S" Then
'                    If (vlMes = 7 Or vlMes = 12) Then
'
'                    End If
                     'If Not fgObtienePrcSalud(stBeneficiarios(i).Cod_InsSalud, stBeneficiarios(i).Cod_ModSalud, stBeneficiarios(i).Mto_PlanSalud, vlBaseImp, vlDesSalud, stDatGenerales.Val_UFUltDiaMes, vgFecPago, vlTB!Cod_Moneda, vlTCMonedaPension) Then
                    If Not fgObtienePrcSalud(stBeneficiarios(i).Cod_InsSalud, stBeneficiarios(i).Cod_ModSalud, stBeneficiarios(i).Mto_PlanSalud, vlBaseImp, vlDesSalud, stDatGenerales.Val_UFUltDiaMes, vlDiaAnteriorFecPago, vlTB!Cod_Moneda, vlTCMonedaPension) Then
                        GoTo Deshacer
                    End If
                End If
                stDetPension.Cod_ConHabDes = stDatGenerales.Cod_ConceptoDesctoSalud
                stDetPension.Mto_ConHabDes = vlDesSalud
                stLiquidacion.Mto_Descuento = stLiquidacion.Mto_Descuento + vlDesSalud
                'Graba Monto del Descuento de Salud
                If Not fgInsertaDetallePensionProv(stDetPension) Then
                    MsgBox "Se ha producido un Error al Grabar Descuento de Salud" & Chr(13) & Err.Description, vbCritical, Me.Caption
                    GoTo Deshacer
                End If
                                
                '6.- Agregar Haberes y Descuentos No Imponibles y Tributables (1a. Etapa)
                'If Not fgObtieneHaberesDescuentos(Me.Tag, vgPerPago, vlTB!Num_Poliza, vlTB!num_endoso, stBeneficiarios(i).Num_Orden, "N", "S", vlMonto, vlBaseImp, 0, vgFecPago, stLiquidacion, stDetPension, stDatGenerales.Val_UF, vgFecIniPag, vgFecTerPag, vlTB!Cod_Moneda, vlTCMonedaPension) Then
                If Not fgObtieneHaberesDescuentos(Me.Tag, vgPerPago, vlTB!num_poliza, vlTB!num_endoso, stBeneficiarios(i).Num_Orden, "N", "S", vlMonto, vlBaseImp, 0, vlDiaAnteriorFecPago, stLiquidacion, stDetPension, stDatGenerales.Val_UF, vgFecIniPag, vgFecTerPag, vlTB!Cod_Moneda, vlTCMonedaPension) Then
                    GoTo Deshacer
                End If
                vlBaseTrib = (vlBaseImp - vlDesSalud) + vlMonto 'Base Imponible
                stLiquidacion.Mto_BaseTri = vlBaseTrib
                '9.- Calcular Retencion Judicial (2a. Etapa)
                'If Not fgCalculaRetencion(vlTB!Num_Poliza, vlTB!num_endoso, stBeneficiarios(i).Num_Orden, vgPerPago, vgFecPago, vlBaseImp, vlBaseTrib, stLiquidacion, stDetPension, vgFecIniPag, Me.Caption, vlTB!Cod_Moneda, vlTCMonedaPension) Then
                If Not fgCalculaRetencion(vlTB!num_poliza, vlTB!num_endoso, stBeneficiarios(i).Num_Orden, vgPerPago, vlDiaAnteriorFecPago, stLiquidacion.Mto_BaseTri, vlBaseTrib, stLiquidacion, stDetPension, vgFecIniPag, Me.Caption, vlTB!Cod_Moneda, vlTCMonedaPension) Then
                    GoTo Deshacer
                End If
                
                '10.- Agregar Haberes y Descuentos No Imponibles y No Tributables (1a. Etapa)
                'If Not fgObtieneHaberesDescuentos(Me.Tag, vgPerPago, vlTB!Num_Poliza, vlTB!num_endoso, stBeneficiarios(i).Num_Orden, "N", "N", vlMonto, vlBaseImp, vlBaseTrib, vgFecPago, stLiquidacion, stDetPension, stDatGenerales.Val_UF, vgFecIniPag, vgFecTerPag, vlTB!Cod_Moneda, vlTCMonedaPension) Then
                If Not fgObtieneHaberesDescuentos(Me.Tag, vgPerPago, vlTB!num_poliza, vlTB!num_endoso, stBeneficiarios(i).Num_Orden, "N", "N", vlMonto, vlBaseImp, vlBaseTrib, vlDiaAnteriorFecPago, stLiquidacion, stDetPension, stDatGenerales.Val_UF, vgFecIniPag, vgFecTerPag, vlTB!Cod_Moneda, vlTCMonedaPension) Then
                    GoTo Deshacer
                End If
                
                '12.- Obtener Mensajes (1a. Etapa)
                
                '13.-  Generar Liquidación (1a. Etapa)
                stLiquidacion.Cod_InsSalud = stBeneficiarios(i).Cod_InsSalud
                stLiquidacion.Cod_ModSalud = stBeneficiarios(i).Cod_ModSalud
                stLiquidacion.Mto_PlanSalud = stBeneficiarios(i).Mto_PlanSalud
                stLiquidacion.Mto_LiqPagar = stLiquidacion.Mto_Haber - stLiquidacion.Mto_Descuento
                
                'Inserta Liquidacion
                If Not fgInsertaLiquidacion(stLiquidacion) Then
                    GoTo Deshacer
                End If
Siguiente:
                i = i + 1
           
                
            Loop
        End If 'hqr 13/01/2011 se cambia de lugar el end if, para que siempre guarde la pensión actualizada
        
        'Graba Pensión Actualizada
'        If vlPasaTrimestre = True Then
'            If vlAjuste <> cgSINAJUSTE Then 'hqr 11/01/2011
'                'If vlExistePensionAct = True Then
''                    vlSql = "UPDATE PP_TMAE_PENSIONACT SET"
''                    vlSql = vlSql & " mto_pension=" & Str(vlMontoPensionAct) & ","
''                    vlSql = vlSql & " cod_tipopago='" & stLiquidacion.Cod_TipoPago & "',"
''                    vlSql = vlSql & " Mto_PensionGar=" & Str(vlMontoPensionActGar) & ""
''                    vlSql = vlSql & " WHERE num_poliza='" & stLiquidacion.num_poliza & "' AND num_endoso=" & stLiquidacion.num_endoso & " AND fec_desde='" & Format(vgFecIniPag, "yyyymmdd") & "'"
''                Else
'                    vlSql = "INSERT INTO pp_tmae_pensionact "
'                    vlSql = vlSql & "(num_poliza, num_endoso, fec_desde, mto_pension, cod_tipopago, mto_pensiongar) "
'                    vlSql = vlSql & "VALUES ("
'                    vlSql = vlSql & "'" & stLiquidacion.num_poliza & "', " & stLiquidacion.num_endoso & ", "
'                    vlSql = vlSql & "'" & Format(vgFecIniPag, "yyyymmdd") & "', " & str(vlMontoPensionAct) & ",'" & stLiquidacion.Cod_TipoPago & "', " & str(vlMontoPensionActGar) & ")"
'                'End If
'                vgConexionTransac.Execute (vlSql)
'                'Call FgGuardaLog(vlSql, vgUsuario, "6086") 'RRR 25/07/2014
'            End If
'        End If

        'Refresca Barra de Progreso
        If (ProgressBar.Value + vlAumento) <= 100 Then
            ProgressBar.Value = (ProgressBar.Value + vlAumento)
        End If
        ProgressBar.Refresh
        Me.Refresh
        contador = contador + 1
           lblmensaje.Caption = "Registro " & contador & " de " & vlcont
              Me.Refresh
      DoEvents
      
Sigui:
        vlTB.MoveNext
    Loop
Else
    MsgBox "No existen Pólizas con Pago en Régimen para este Periodo", vbCritical, Me.Caption
    GoTo Deshacer
End If

''Traspas Datos a Tabla de Resumen
'vlSql = "INSERT INTO pp_tmae_respagopenprov" 'Pago en Régimen
'vlSql = vlSql & " (num_perpago, cod_tippension, cod_conhabdes,"
'vlSql = vlSql & " cod_tipopago, cod_moneda, mto_conhabdes) "
'vlSql = vlSql & " SELECT '" & vgPerPago & "' , liq.cod_tippension, pag.cod_conhabdes,"
'vlSql = vlSql & " liq.cod_tipopago, liq.cod_moneda, sum(pag.mto_conhabdes)"
'vlSql = vlSql & " FROM pp_tmae_pagopenprov pag, pp_tmae_liqpagopenprov liq"
'vlSql = vlSql & " WHERE liq.num_perpago = pag.num_perpago"
'vlSql = vlSql & " AND liq.num_poliza = pag.num_poliza"
'vlSql = vlSql & " AND liq.num_orden = pag.num_orden"
'vlSql = vlSql & " AND liq.num_idenreceptor = pag.num_idenreceptor"
'vlSql = vlSql & " AND liq.cod_tipoidenreceptor = pag.cod_tipoidenreceptor"
'vlSql = vlSql & " AND liq.cod_tipreceptor = pag.cod_tipreceptor"
'vlSql = vlSql & " AND liq.num_perpago = '" & vgPerPago & "'"
'vlSql = vlSql & " AND liq.cod_tipopago = 'R'" 'Pago en Régimen
'vlSql = vlSql & " GROUP BY liq.cod_tippension, pag.cod_conhabdes, liq.cod_tipopago, liq.cod_moneda"
'vgConexionTransac.Execute (vlSql)
'Call FgGuardaLog(vlSql, vgUsuario, "6086") 'RRR 25/07/2014

    
'Traspasa Datos a Histórico
If Me.Tag = "D" Then
    'vgConexionTransac.CommitTrans
    If Not flTraspasaDatosADefinitivos Then
        GoTo Deshacer
    End If
'CMV-20060222 I
'Indicado por Srta.: Hilda Quezada
Else
    'Cierra el Periodo de Cálculo Provisorio
    vlSql = "UPDATE PP_TMAE_PROPAGOPEN"
    vlSql = vlSql & " SET COD_ESTADOREG = 'P'"
    vlSql = vlSql & " WHERE NUM_PERPAGO = '" & vgPerPago & "'"
    vgConexionTransac.Execute (vlSql)
    'Call FgGuardaLog(vlSql, vgUsuario, "6087") 'RRR 25/07/2014
'CMV-20060222 F
End If

vgConexionTransac.CommitTrans
vgConexionTransac.Close




If Me.Tag <> "D" Then
    'Verifica si existen Pensiones con monto Menor a 0
    vlSql = "SELECT COUNT(1) AS cont FROM pp_tmae_liqpagopenprov liq"
    vlSql = vlSql & " WHERE liq.num_perpago = '" & vgPerPago & "'"
    vlSql = vlSql & " AND liq.cod_tipopago = 'P'" 'Primer Pago
    vlSql = vlSql & " AND liq.mto_liqpagar < 0" 'Monto Pension Negativa
    Set vgRs = vgConexionBD.Execute(vlSql)
    If Not vgRs.EOF Then
        'Pueden existir Pensiones con monto Menor a cero
        If vgRs!cont > 0 Then
            MsgBox "Cálculo de Pagos en Régimen realizados exitosamente." & Chr(13) & Chr(13) & "Existen [" & vgRs!cont & "] Pensiones con Monto a Pagar Menor a Cero." & _
            Chr(13) & "Favor corregir estos casos o no podrá Cerrar el Periodo al realizar el Cálculo Definitivo.", vbExclamation, Me.Caption
        Else
            MsgBox "Cálculo de Pagos en Régimen realizados exitosamente", vbInformation, Me.Caption
        End If
    Else
        MsgBox "Cálculo de Pagos en Régimen realizados exitosamente", vbInformation, Me.Caption
    End If
End If

Screen.MousePointer = 0

Errores:
    If Err.Number <> 0 Then
Deshacer:
        vgConexionTransac.RollbackTrans
        vgConexionTransac.Close
        If Err.Number <> 0 Then
            MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]/ poliza " & stLiquidacion.num_poliza, vbCritical, "¡ERROR!..."
        End If
    End If
    ProgressBar.Value = 0
    Frame3 = "Progreso del Cálculo"
    Screen.MousePointer = vbDefault
End Function

Private Sub Form_Unload(Cancel As Integer)
    Frm_Menu.Tag = "0"
End Sub

Private Function flCierraPeriodo() As Boolean
    On Error GoTo Errores
    'Cierra el Periodo de Cálculo
    vlSql = "UPDATE PP_TMAE_PROPAGOPEN"
    vlSql = vlSql & " SET COD_ESTADOREG = 'C'"
    vlSql = vlSql & " WHERE NUM_PERPAGO = '" & vgPerPago & "'"
    vgConexionTransac.Execute (vlSql)
    'Call FgGuardaLog(vlSql, vgUsuario, "6087")    'RRR 25/07/2014
    flCierraPeriodo = True
    Exit Function
Errores:
    flCierraPeriodo = False
End Function

Function flValidaPensionGarantizada(iFecIniPag As String) As Double
'Obtiene los beneficiarios que cumplieron 18 años antes del primer día del mes de pago
'y que tienen una pensión garantizada.
'Esto es para que generen el endoso manual antes de las pensiones recurrentes
'y de esta forma la pensión garantizada quede calculada correctamente

On Error GoTo Errores

    Dim vlSql As String
    Dim L18 As Long
    
    
    If (fgCarga_Param("LI", "L18", iFecIniPag) = True) Then
        L18 = vgValorParametro
    Else
        'vgError = 1000
        MsgBox "No existe Edad de tope para los 18 años.", vbCritical, "Proceso Cancelado"
        Exit Function
    End If
    
    If (fgCarga_Param("LI", "L24", iFecIniPag) = True) Then
        L24 = vgValorParametro
    Else
        'vgError = 1000
        MsgBox "No existe Edad de tope para los 18 años.", vbCritical, "Proceso Cancelado"
        Exit Function
    End If
    
    'Mensualizar la Edad de 24 Años
    L18 = L18 * 12
    L24 = L24 * 12


    flValidaPensionGarantizada = -1
    'Verifica si existen Pensiones con monto Menor a 0
'    vlSql = "SELECT COUNT(1) AS cont "
'    vlSql = vlSql & " FROM pp_tmae_poliza p, pp_tmae_ben b"
'    vlSql = vlSql & " WHERE p.num_poliza = b.num_poliza"
'    vlSql = vlSql & " AND p.num_endoso = b.num_endoso"
'    vlSql = vlSql & " AND p.num_endoso = " 'Ultimo Endoso generado
'        vlSql = vlSql & " (SELECT NVL(MAX(e.num_endoso + 1),1) "
'        vlSql = vlSql & " FROM pp_tmae_endoso e "
'        vlSql = vlSql & " WHERE e.num_poliza = p.num_poliza "
'        vlSql = vlSql & " AND e.cod_estado = 'E' "
'        vlSql = vlSql & " AND e.fec_efecto <= '" & iFecIniPag & "' "
'        vlSql = vlSql & " AND e.fec_finefecto >='" & iFecIniPag & "')"
'    vlSql = vlSql & " AND b.cod_estpension <> '10'" 'Beneficiarios que tengan derecho a pension
'    vlSql = vlSql & " AND p.num_mesgar > 0 " 'Poliza con pensión garantizada
'    vlSql = vlSql & " AND p.fec_finpergar >= '" & iFecIniPag & "'" 'El periodo garantizado aun no ha concluido
'    vlSql = vlSql & " AND b.fec_nacben < to_Char(add_months(to_date('" & iFecIniPag & "','yyyymmdd'),-216),'yyyymmdd')" 'El beneficiario cumplió los 18 años antes del inicio del mes de pago
'    vlSql = vlSql & " AND b.cod_par >= '30' and b.cod_par <= '35'" 'Parentesco de hijos
'    vlSql = vlSql & " AND b.fec_inipagopen <= '" & iFecIniPag & "'" 'Beneficiarios que ya han iniciado el periodo de pago
'    vlSql = vlSql & " AND p.cod_estado IN (6, 7, 8) " 'solo Pólizas Vigentes, con Beneficiarios Designados y Pendiente (Pago Diferido)
'    vlSql = vlSql & " AND (p.fec_pripago < '" & iFecIniPag & "'" 'Polizas que ya tuvieron su primer pago
'    vlSql = vlSql & " OR (p.num_mesdif > 0 AND p.fec_pripago = '" & iFecIniPag & "'))" 'O polizas que inician el pago en el mes actual
    
    
    
'    vlSql = " SELECT count(*) as cont  FROM pp_tmae_poliza p"
'    vlSql = vlSql & " JOIN pp_tmae_ben b ON p.num_poliza=b.num_poliza and p.num_endoso=b.num_endoso"
'    vlSql = vlSql & " WHERE p.num_endoso = (select max(num_endoso) from pp_tmae_poliza where num_poliza=p.num_poliza)"
'    vlSql = vlSql & " AND b.cod_estpension not in ('10','20') and b.cod_derpen not in (10,20)"
'    vlSql = vlSql & " AND p.num_mesgar > 0"
'    vlSql = vlSql & " AND p.fec_finpergar >= '" & iFecIniPag & "'"
'    vlSql = vlSql & " AND b.fec_nacben < to_Char(add_months(to_date('" & iFecIniPag & "','yyyymmdd'),-" & CStr(L18) & "),'yyyymmdd')"
'    vlSql = vlSql & " AND b.cod_par in ('30','35')"
'    vlSql = vlSql & " AND b.fec_inipagopen <= '" & iFecIniPag & "'"
'    vlSql = vlSql & " AND p.cod_estado IN (6, 7, 8) and cod_sitinv not in ('T', 'P')"
'    vlSql = vlSql & " AND (p.fec_pripago < '" & iFecIniPag & "' OR (p.num_mesdif > 0 AND p.fec_pripago = '" & iFecIniPag & "'))"
    
    
    
    vlSql = " SELECT count(*) as cont  FROM pp_tmae_poliza p JOIN pp_tmae_ben b ON p.num_poliza=b.num_poliza and p.num_endoso=b.num_endoso"
    vlSql = vlSql & " WHERE p.num_endoso = (select max(num_endoso) from pp_tmae_poliza where num_poliza=p.num_poliza)"
    vlSql = vlSql & " AND b.cod_estpension not in ('10','20') and b.cod_derpen not in (10,20) AND p.num_mesgar > 0 AND p.fec_finpergar >= '" & iFecIniPag & "'"
    vlSql = vlSql & " AND b.fec_nacben < to_Char(add_months(to_date('" & iFecIniPag & "','yyyymmdd'),-"
    vlSql = vlSql & " ("
    vlSql = vlSql & "                 select case when (fec_devsol is null) or (to_date(fec_devsol, 'YYYYMMDD') < to_date('20130801', 'YYYYMMDD')) then " & CStr(L18) & " else case when ("
    vlSql = vlSql & "                 select distinct case when IND_DNI='S' then 1 else 0 end  + case when IND_DJU='S' then 1 else 0 end  + case when IND_PES='S' then 1 else 0 end  + case when IND_BNO='S' then 1 else 0 end"
    vlSql = vlSql & "                 from pp_tmae_certificado where num_poliza=a.num_poliza and num_endoso=a.num_endoso and num_orden=ben.num_orden and cod_tipo='EST' AND EST_ACT=1 AND (FEC_INICER<='" & iFecIniPag & "' AND FEC_TERCER>='" & iFecIniPag & "')"
    vlSql = vlSql & "                 ) = 4 then " & CStr(L24) & " else case when (to_date(fec_devsol, 'YYYYMMDD') < to_date('20130801', 'YYYYMMDD')) then " & CStr(L18) & " else " & CStr(L24) & " end end end as edad_l"
    vlSql = vlSql & "                 from pp_tmae_poliza a"
    vlSql = vlSql & "                 join pp_tmae_ben ben on a.num_poliza=ben.num_poliza and a.num_endoso=ben.num_endoso"
    vlSql = vlSql & "                 Where a.num_poliza = p.num_poliza And a.num_endoso = p.num_endoso And ben.Num_Orden = b.Num_Orden"
    vlSql = vlSql & " )"
    vlSql = vlSql & " ),'yyyymmdd') AND b.cod_par in ('30','35')"
    vlSql = vlSql & " AND b.fec_inipagopen <= '" & iFecIniPag & "' AND p.cod_estado IN (6, 7, 8) and cod_sitinv not in ('T', 'P')"
    vlSql = vlSql & " AND (p.fec_pripago < '" & iFecIniPag & "' OR (p.num_mesdif > 0 AND p.fec_pripago = '" & iFecIniPag & "'))"
    Set vgRs = vgConexionBD.Execute(vlSql)
    If Not vgRs.EOF Then
        If vgRs!cont <> 0 Then
            flValidaPensionGarantizada = vgRs!cont
        Else
            flValidaPensionGarantizada = 0
        End If
    End If
    
Errores:
    If Err.Number <> 0 Then
        MsgBox "Se ha producido un Error al Validar Beneficiarios con Pensión Garantizada." & Chr(13) & Err.Description, vbCritical, Me.Caption
    End If
End Function



'Materia Gris - Jaime Rios 19/02/2018
Function flBuscarHijosMayoresPensionGarantizada() As Boolean
    On Error GoTo Errores
    flBuscarHijosMayoresPensionGarantizada = False

'    vlSql = " SELECT b.num_poliza, b.num_endoso, b.num_orden"
'    vlSql = vlSql & " FROM pp_tmae_ben b where"
'    vlSql = vlSql & " cod_par=30 and cod_derpen=99 and cod_estpension=99 and prc_pensiongar<>0 and cod_sitinv='N' "
'    'vlSql = vlSql & " and (months_between(TRUNC(sysdate),to_date(fec_nacben,'yyyymmdd'))/12)>=18 "
'    vlSql = vlSql & " and (months_between(TRUNC(sysdate),to_date(regexp_replace(fec_nacben,'(^.{6})(.{2})(.*)$','\128\3'),'yyyymmdd')))>216  "
'    vlSql = vlSql & " and b.num_endoso=(select max(p.num_endoso) from pp_tmae_poliza p where p.num_poliza=b.num_poliza and p.cod_tippension<>08) "
'    vlSql = vlSql & " and to_date((select fec_finpergar from pp_tmae_poliza p where p.num_poliza=b.num_poliza and p.num_endoso=b.num_endoso) ,'yyyymmdd')>=sysdate "
'    vlSql = vlSql & "   and (select count(*) from pp_tmae_certificado c where c.num_poliza=b.num_poliza and c.num_orden=b.num_orden and c.num_endoso=b.num_endoso "
'    vlSql = vlSql & "   and fec_tercer=(select max(fec_tercer) from pp_tmae_certificado where num_poliza=c.num_poliza and num_orden=c.num_orden and cod_tipo='EST') "
'    vlSql = vlSql & "   and to_date(fec_tercer,'yyyymmdd') > sysdate " 'certificado vigente
'    vlSql = vlSql & " and c.est_act=1 and c.ind_dni='S' and c.ind_dju='S' and c.ind_pes='S' and c.ind_bno='S' )=0 "
'    vlSql = vlSql & " order by 1,2,3 "
    
    vlSql = " SELECT b.num_poliza, b.num_endoso, b.num_orden, b.fec_nacben as cont  FROM pp_tmae_poliza p JOIN pp_tmae_ben b ON p.num_poliza=b.num_poliza and p.num_endoso=b.num_endoso"
    vlSql = vlSql & " WHERE p.num_endoso = (select max(num_endoso) from pp_tmae_poliza where num_poliza=p.num_poliza)"
    vlSql = vlSql & " AND b.cod_estpension not in ('10','20') and b.cod_derpen not in (10,20) AND p.num_mesgar > 0 AND p.fec_finpergar >= '" & iFecIniPag & "'"
    vlSql = vlSql & " AND b.fec_nacben < to_Char(add_months(to_date('" & iFecIniPag & "','yyyymmdd'),-"
    vlSql = vlSql & " ("
    vlSql = vlSql & "                 select case when (fec_devsol is null) or (to_date(fec_devsol, 'YYYYMMDD') < to_date('20130801', 'YYYYMMDD')) then " & CStr(L18) & " else case when ("
    vlSql = vlSql & "                 select distinct case when IND_DNI='S' then 1 else 0 end  + case when IND_DJU='S' then 1 else 0 end  + case when IND_PES='S' then 1 else 0 end  + case when IND_BNO='S' then 1 else 0 end"
    vlSql = vlSql & "                 from pp_tmae_certificado where num_poliza=a.num_poliza and num_endoso=a.num_endoso and num_orden=ben.num_orden and cod_tipo='EST' AND EST_ACT=1 AND (FEC_INICER<='" & iFecIniPag & "' AND FEC_TERCER>='" & iFecIniPag & "')"
    vlSql = vlSql & "                 ) = 4 then " & CStr(L24) & " else case when (to_date(fec_devsol, 'YYYYMMDD') < to_date('20130801', 'YYYYMMDD')) then " & CStr(L18) & " else " & CStr(L24) & " end end end as edad_l"
    vlSql = vlSql & "                 from pp_tmae_poliza a"
    vlSql = vlSql & "                 join pp_tmae_ben ben on a.num_poliza=ben.num_poliza and a.num_endoso=ben.num_endoso"
    vlSql = vlSql & "                 Where a.num_poliza = p.num_poliza And a.num_endoso = p.num_endoso And ben.Num_Orden = b.Num_Orden"
    vlSql = vlSql & " )"
    vlSql = vlSql & " ),'yyyymmdd') AND b.cod_par in ('30','35')"
    vlSql = vlSql & " AND b.fec_inipagopen <= '" & iFecIniPag & "' AND p.cod_estado IN (6, 7, 8) and cod_sitinv not in ('T', 'P')"
    vlSql = vlSql & " AND (p.fec_pripago < '" & iFecIniPag & "' OR (p.num_mesdif > 0 AND p.fec_pripago = '" & iFecIniPag & "'))"
    
    Set vgRs1 = vgConexionBD.Execute(vlSql)
    
    'vgRs1.MoveFirst
    If Not vgRs1.EOF Then
        Do While Not vgRs1.EOF
            'MsgBox (vgRs1!num_poliza & "-" & vgRs1!num_endoso & "-" & vgRs1!Num_Orden)
            Call flActualizarPorcentajesGarantizados(vgRs1!num_poliza, vgRs1!num_endoso, vgRs1!Num_Orden)
            vgRs1.MoveNext
        Loop
        flBuscarHijosMayoresPensionGarantizada = True
    Else
        flBuscarHijosMayoresPensionGarantizada = True
    End If

    Exit Function
Errores:
    MsgBox "Se ha producido un Error al Buscar Beneficiarios con Pensión Garantizada que sean hijos mayores de edad y no deban recibir pensión." & Chr(13) & Err.Description, vbCritical, Me.Caption
    flBuscarHijosMayoresPensionGarantizada = False
End Function

'Materia Gris - Jaime Rios 19/02/2018
Function flActualizarPorcentajesGarantizados(num_poliza As String, num_endoso As String, Num_Orden As String)
On Error GoTo Error_flActualizarPor
    
    'Obtener beneficiarios
    vlSql = " SELECT b.num_poliza, b.num_endoso, b.num_orden, b.Cod_TipoIdenBen, b.Num_IdenBen, b.Gls_NomBen, b.Gls_NomSegBen, b.Gls_PatBen, b.Gls_MatBen "
    vlSql = vlSql & " ,b.Cod_Par, b.Cod_GruFam, b.Cod_Sexo, b.Cod_SitInv, b.Cod_DerCre, b.Cod_CauInv, b.Fec_NacBen, b.Fec_NacHM, b.Fec_InvBen, b.Mto_Pension "
    vlSql = vlSql & " ,b.Prc_Pension, b.Fec_FallBen, b.Cod_DerPen, b.Cod_EstPension, b.Cod_MotReqPen, b.Mto_PensionGar, b.Cod_CauSusBen, b.Fec_SusBen, b.Fec_IniPagoPen "
    vlSql = vlSql & " ,b.Fec_TerPagoPenGar, b.Fec_Matrimonio, b.Prc_PensionGar, b.Prc_PensionLeg, b.Cod_Direccion, b.Gls_DirBen, b.Gls_FonoBen, b.Gls_CorreoBen, b.Gls_Telben2 "
    vlSql = vlSql & " FROM pp_tmae_ben b where "
    vlSql = vlSql & " b.num_poliza= " & num_poliza & " and b.num_endoso=" & num_endoso & ""
    Set vgRs = vgConexionBD.Execute(vlSql)
    Dim stBeneficiariosMod() As TyBeneficiarios
    Dim vlNumCargas As Integer
    vlNumCargas = 1
    Do While Not vgRs.EOF
        vlNumCargas = vlNumCargas + 1
        vgRs.MoveNext
    Loop
    ReDim stBeneficiariosMod(vlNumCargas - 1) As TyBeneficiarios
    vgRs.MoveFirst
    vlNumCargas = 1
    Do While Not vgRs.EOF
        With stBeneficiariosMod(vlNumCargas)
            .num_poliza = vgRs!num_poliza
            .num_endoso = vgRs!num_endoso
            .Num_Orden = vgRs!Num_Orden
            .Cod_TipoIdenBen = vgRs!Cod_TipoIdenBen
            .Num_IdenBen = vgRs!Num_IdenBen
            .Gls_NomBen = vgRs!Gls_NomBen
            .Gls_NomSegBen = IIf(IsNull(vgRs!Gls_NomSegBen), "", vgRs!Gls_NomSegBen)
            .Gls_PatBen = vgRs!Gls_PatBen
            .Gls_MatBen = vgRs!Gls_MatBen
            .Cod_Par = vgRs!Cod_Par
            .Cod_GruFam = vgRs!Cod_GruFam
            .Cod_Sexo = vgRs!Cod_Sexo
            .Cod_SitInv = vgRs!Cod_SitInv
            .Cod_DerCre = vgRs!Cod_DerCre
            .Cod_CauInv = vgRs!Cod_CauInv
            .Fec_NacBen = vgRs!Fec_NacBen
            .Fec_NacHM = IIf(IsNull(vgRs!Fec_NacHM), "", vgRs!Fec_NacHM)
            .Fec_InvBen = IIf(IsNull(vgRs!Fec_InvBen), "", vgRs!Fec_InvBen)
            .Mto_Pension = vgRs!Mto_Pension
            .Prc_Pension = vgRs!Prc_Pension
            .Fec_FallBen = IIf(IsNull(vgRs!Fec_FallBen), "", vgRs!Fec_FallBen)
            .Cod_DerPen = vgRs!Cod_DerPen
            .Cod_EstPension = vgRs!Cod_EstPension
            .Cod_MotReqPen = vgRs!Cod_MotReqPen
            .Mto_PensionGar = vgRs!Mto_PensionGar
            .Cod_CauSusBen = vgRs!Cod_CauSusBen
            .Fec_SusBen = IIf(IsNull(vgRs!Fec_SusBen), "", vgRs!Fec_SusBen)
            .Fec_IniPagoPen = vgRs!Fec_IniPagoPen
            .Fec_TerPagoPenGar = IIf(IsNull(vgRs!Fec_TerPagoPenGar), "", vgRs!Fec_TerPagoPenGar)
            .Fec_Matrimonio = IIf(IsNull(vgRs!Fec_Matrimonio), "", vgRs!Fec_Matrimonio)
            .Prc_PensionGar = vgRs!Prc_PensionGar
            .Prc_PensionLeg = vgRs!Prc_PensionLeg
            .Cod_Direccion = vgRs!Cod_Direccion
            .Gls_DirBen = vgRs!Gls_DirBen
            .Gls_FonoBen = IIf(IsNull(vgRs!Gls_FonoBen), "", vgRs!Gls_FonoBen)
            .Gls_CorreoBen = IIf(IsNull(vgRs!Gls_CorreoBen), "", vgRs!Gls_CorreoBen)
            .Gls_Telben2 = IIf(IsNull(vgRs!Gls_Telben2), "", vgRs!Gls_Telben2)
        End With
        If (vgRs!Num_Orden = Num_Orden) Then
        '    stBeneficiariosMod(vlNumCargas).Cod_DerPen = 10
        '    stBeneficiariosMod(vlNumCargas).Cod_EstPension = 10
            stBeneficiariosMod(vlNumCargas).Prc_PensionGar = 0
        End If
        vlNumCargas = vlNumCargas + 1
        vgRs.MoveNext
    Loop
    
    vlNumCargas = vlNumCargas - 1
    Dim cadena As String
    '    For i = 1 To vlNumCargas
    '        With stBeneficiariosMod(i)
    '            cadena = cadena & .Cod_DerPen & " - " & .Cod_EstPension & " - " & .Num_Orden & " - " & .Prc_Pension & " - " & .Prc_PensionGar & " " & vbNewLine
    '        End With
    '    Next
    '    MsgBox cadena
    
    'Obtener datos de la poliza
    vlSql = " SELECT * FROM PP_TMAE_POLIZA WHERE NUM_POLIZA=" & num_poliza & " AND NUM_ENDOSO=" & num_endoso & ""
    Set vgRs = vgConexionBD.Execute(vlSql)
    
    'Obtener montos actuales
    vlSql = " SELECT nvl(max(a.mto_pension), b.mto_pension) as mto_pension, nvl(Max(a.Mto_PensionGar), b.Mto_PensionGar) As mto_Pensiongar "
    vlSql = vlSql & " FROM pp_tmae_pensionact a full join pp_tmae_poliza b on a.num_poliza=b.num_poliza and a.num_endoso=b.num_endoso "
    vlSql = vlSql & " WHERE b.num_poliza = " & num_poliza & " and b.num_endoso=" & num_endoso & " group by b.mto_pension, b.mto_pensiongar "
    Set vgMonto = vgConexionBD.Execute(vlSql)
    
    'Call fgCalcularPorcentajeBenef(vlFecVigencia, vlNumCargas, stBeneficiariosMod, vlCodTipPension, vlMtoPensionRef, False, vlCodDerCrecerCot, vlIndCobertura, True, vlIsGar)
    'Call fgCalcularPorcentajeB(num_poliza, vgFecPago, vlNumCargas, stBeneficiariosMod, vgRs!Cod_TipPension, 0#, False, vgRs!Cod_DerCre, vgRs!Ind_Cob, True, CInt(vgRs!Num_MesGar), 0#)
    Call fgCalcularPorcentajeB(num_poliza, vgFecPago, vlNumCargas, stBeneficiariosMod, vgRs!Cod_TipPension, vgMonto!Mto_Pension, False, vgRs!Cod_DerCre, vgRs!Ind_Cob, True, CInt(vgRs!Num_MesGar), vgMonto!Mto_PensionGar)
    
    'cadena = ""
    '    For i = 1 To vlNumCargas
    '        With stBeneficiariosMod(i)
    '            cadena = cadena & .Cod_DerPen & " - " & .Cod_EstPension & " - " & .Num_Orden & " - " & .Mto_Pension & " - " & .Prc_Pension & " - " & .Mto_PensionGar & " - " & .Prc_PensionGar & " " & vbNewLine & ""
    '        End With
    '    Next
    '    MsgBox cadena
    'Obtener fecha y hora actuales
    vgSql = "SELECT TO_CHAR(SYSDATE,'DD/MM/YYYY HH24:MI:SS') AS FEC_ACTUAL FROM MA_TCOD_GENERAL"
       Set vgRs4 = vgConexionBD.Execute(vgSql)
       If Not vgRs4.EOF Then
          FecServ = Format(CDate(Trim(Mid((vgRs4!FEC_ACTUAL), 1, 10))), "yyyymmdd")
          HoraActual = Format(Mid((vgRs4!FEC_ACTUAL), 12, 8), "hhmmss")
       End If
    
    'Eliminar registros de endoso preliminar
    vgSql = "DELETE PP_TMAE_ENDOSO WHERE "
    vgSql = vgSql & "num_poliza = '" & num_poliza & "' AND "
    vgSql = vgSql & "num_endoso = " & num_endoso & " "
    vgConexionBD.Execute (vgSql)
    
    vgSql = "DELETE PP_TMAE_ENDBEN WHERE "
    vgSql = vgSql & "num_poliza = '" & num_poliza & "' AND "
    vgSql = vgSql & "num_endoso = " & num_endoso & " "
    vgConexionBD.Execute (vgSql)
    
    vgSql = "DELETE PP_TMAE_ENDPOLIZA WHERE "
    vgSql = vgSql & "num_poliza = '" & num_poliza & "' AND "
    vgSql = vgSql & "num_endoso = " & num_endoso & " "
    vgConexionBD.Execute (vgSql)
    
    'Buscar datos de endoso anterior
    vgSql = "SELECT * FROM PP_TMAE_ENDOSO "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & " num_poliza = '" & num_poliza & "' AND "
    vgSql = vgSql & " num_endoso = " & num_endoso - 1 & " "
    'Set vgReg = vgConectarBD.Execute(vgSql)
    Set vgRs4 = vgConexionBD.Execute(vgSql)
    
        'Insertar Endoso
        vgSql = "INSERT INTO PP_TMAE_ENDOSO "
        vgSql = vgSql & "(num_poliza,num_endoso,fec_solendoso,fec_endoso,cod_cauendoso,cod_tipendoso,mto_diferencia, "
        vgSql = vgSql & "mto_pensionori,mto_pensioncal,fec_efecto,prc_factor, " 'gls_observacion, "
        vgSql = vgSql & "fec_finefecto,cod_moneda,cod_estado,cod_usuariocrea,fec_crea,hor_crea "
        vgSql = vgSql & ", cod_tipreajuste, mto_valreajustetri, mto_valreajustemen "
        vgSql = vgSql & " ) VALUES ( "
        vgSql = vgSql & "'" & num_poliza & "', " & num_endoso & ", '" & FecServ & "', '" & FecServ & "', "
        vgSql = vgSql & "'26', 'S', " & str(vgRs4!MTO_DIFERENCIA) & ", "
        vgSql = vgSql & " " & str(vgRs4!mto_pensionori) & ", " & str(vgRs4!mto_pensioncal) & ", "
        vgSql = vgSql & "'" & vgRs4!FEC_EFECTO & "', " & str(vgRs4!PRC_FACTOR) & ", "
        'vgSql = vgSql & "'Endoso automàtico por mayorìa de edad', "
        vgSql = vgSql & "'" & vgRs4!fec_finefecto & "', '" & vgRs4!Cod_Moneda & "', '" & vgRs4!Cod_Estado & "',"
        vgSql = vgSql & "'SEACSA', '" & FecServ & "', '" & HoraActual & "', "
        vgSql = vgSql & "'" & vgRs4!Cod_TipReajuste & "', " & str(vgRs4!Mto_ValReajusteTri) & ", " & str(vgRs4!Mto_ValReajusteMen) & ") "
        vgConexionBD.Execute vgSql
    'Cargar Datos en pp_tmae_poliza
        vgSql = "INSERT INTO PP_TMAE_POLIZA "
        vgSql = vgSql & "(num_poliza,num_endoso,cod_afp,cod_tippension, cod_estado,cod_tipren,cod_modalidad,num_cargas, "
        vgSql = vgSql & "fec_vigencia,fec_tervigencia,mto_prima,mto_pension, mto_pensiongar, num_mesdif,num_mesgar, "
        vgSql = vgSql & "prc_tasace,prc_tasavta,prc_tasactorea,prc_tasaintpergar,fec_inipagopen, "
        vgSql = vgSql & "cod_usuariocrea,fec_crea,hor_crea ,Cod_Cuspp, Ind_Cob,Cod_Moneda,Mto_ValMoneda, "
        vgSql = vgSql & "Cod_CoberCon,Mto_FacPenElla,Prc_FacPenElla, Cod_DerCre,Cod_DerGra,Fec_Emision,fec_dev, "
        vgSql = vgSql & "Fec_IniPenCia,Fec_PriPago,prc_tasatir ,Fec_finpergar, Fec_finperdif ,Fec_efecto "
        vgSql = vgSql & ",cod_tipreajuste, mto_valreajustetri, mto_valreajustemen  ,fec_devsol , ind_bendes"
        vgSql = vgSql & " ) VALUES ( "
        vgSql = vgSql & "'" & num_poliza & "', " & num_endoso + 1 & ", '" & vgRs!cod_afp & "', '" & vgRs!Cod_TipPension & "', "
        vgSql = vgSql & "'" & vgRs!Cod_Estado & "', '" & vgRs!Cod_TipRen & "', '" & vgRs!Cod_Modalidad & "', " & vlNumCargas & ", "
        vgSql = vgSql & "'" & vgRs!Fec_Vigencia & "', '" & vgRs!Fec_TerVigencia & "',  " & str(vgRs!Mto_Prima) & ", "
        vgSql = vgSql & " " & str(vgRs!Mto_Pension) & ", " & str(vgRs!Mto_PensionGar) & ", " & vgRs!Num_MesDif & ", " & vgRs!Num_MesGar & ", "
        vgSql = vgSql & " " & str(vgRs!Prc_TasaCe) & ", " & str(vgRs!Prc_TasaVta) & ", " & str(vgRs!prc_tasactorea) & ", "
        vgSql = vgSql & " " & str(vgRs!Prc_TasaIntPerGar) & ", '" & Trim(vgRs!Fec_IniPagoPen) & "', "
        vgSql = vgSql & "'SEACSA', '" & FecServ & "', '" & HoraActual & "' "
        vgSql = vgSql & ",'" & Trim(vgRs!Cod_Cuspp) & "', '" & Trim(vgRs!Ind_Cob) & "', "
        vgSql = vgSql & "'" & Trim(vgRs!Cod_Moneda) & "', " & str(vgRs!Mto_ValMoneda) & ", "
        vgSql = vgSql & "'" & Trim(vgRs!Cod_CoberCon) & "', " & str(vgRs!Mto_FacPenElla) & ", " & str(vgRs!Prc_FacPenElla) & ", "
        vgSql = vgSql & "'" & Trim(vgRs!Cod_DerCre) & "', '" & Trim(vgRs!Cod_DerGra) & "', "
        vgSql = vgSql & "'" & Trim(vgRs!Fec_Emision) & "', '" & Trim(vgRs!fec_dev) & "', "
        vgSql = vgSql & "'" & Trim(vgRs!Fec_IniPenCia) & "', '" & Trim(vgRs!Fec_PriPago) & "', " & str(vgRs!prc_tasatir) & " "
        vgSql = vgSql & ",'" & Trim(vgRs!fec_finpergar) & "' ,'" & Trim(vgRs!FEC_FINPERDIF) & "','" & Trim(vgRs!FEC_EFECTO) & "' "
        vgSql = vgSql & ",'" & Trim(vgRs!Cod_TipReajuste) & "'," & str(vgRs!Mto_ValReajusteTri) & "," & str(vgRs!Mto_ValReajusteMen)
        vgSql = vgSql & "," & str(vgRs!FEC_DEVSOL) & ",'" & vgRs!IND_BENDES & "' )"
        vgConexionBD.Execute vgSql
    'Cargar Datos en pd_tmae_poliza(?)
         vgSql = "Insert into PD_TMAE_POLIZA"
         vgSql = vgSql & " (NUM_POLIZA, NUM_ENDOSO, NUM_COT, NUM_CORRELATIVO, NUM_OPERACION,"
         vgSql = vgSql & "  NUM_ARCHIVO, COD_TRAPAGOPEN, FEC_TRAPAGOPEN, COD_AFP, COD_ISAPRE,"
         vgSql = vgSql & "  COD_TIPPENSION, COD_VEJEZ, COD_ESTCIVIL, COD_CUSPP, COD_TIPOIDENAFI,"
         vgSql = vgSql & "  NUM_IDENAFI, GLS_DIRECCION, COD_DIRECCION, GLS_FONO, GLS_CORREO,"
         vgSql = vgSql & "  COD_VIAPAGO, COD_TIPCUENTA, COD_BANCO, NUM_CUENTA, COD_SUCURSAL,"
         vgSql = vgSql & "  FEC_SOLICITUD, FEC_INGVIGENCIA, FEC_VIGENCIA, FEC_DEV, FEC_ACEPTA,"
         vgSql = vgSql & "  FEC_PRIPAGO, COD_MONEDAFON, MTO_MONEDAFON, MTO_PRIUNIFON, MTO_CTAINDFON,"
         vgSql = vgSql & "  MTO_BONOFON, MTO_APOADI, MTO_PRIUNI, MTO_CTAIND, MTO_BONO,"
         vgSql = vgSql & "  PRC_TASARPRT, COD_TIPOIDENCOR, NUM_IDENCOR, PRC_CORCOM, MTO_CORCOM,"
         vgSql = vgSql & "  PRC_CORCOMREAL, NUM_ANNOJUB, NUM_CARGAS, IND_COB, COD_BENSOCIAL,"
         vgSql = vgSql & "  COD_MONEDA, MTO_VALMONEDA, MTO_PRIUNIMOD, MTO_CTAINDMOD, MTO_BONOMOD,"
         vgSql = vgSql & "  COD_TIPREN, NUM_MESDIF, COD_MODALIDAD, NUM_MESGAR, COD_COBERCON,"
         vgSql = vgSql & "  MTO_FACPENELLA, PRC_FACPENELLA, COD_DERCRE, COD_DERGRA, PRC_RENTAAFP,"
         vgSql = vgSql & "  PRC_RENTAAFPORI, PRC_RENTATMP, MTO_CUOMOR, PRC_TASACE, PRC_TASAVTA,"
         vgSql = vgSql & "  PRC_TASATIR, PRC_TASAPERGAR, MTO_CNU, MTO_PRIUNISIM, MTO_PRIUNIDIF,"
         vgSql = vgSql & "  MTO_PENSION, MTO_PENSIONGAR, MTO_CTAINDAFP, MTO_RENTATMPAFP, MTO_RESMAT,"
         vgSql = vgSql & "  MTO_VALPREPENTMP, MTO_PERCON, PRC_PERCON, MTO_SUMPENSION, MTO_PENANUAL,"
         vgSql = vgSql & "  MTO_RMPENSION, MTO_RMGTOSEP, COD_TIPCOT, COD_ESTCOT, COD_USUARIO,"
         vgSql = vgSql & "  COD_SUCURSALUSU, FEC_INIPAGOPEN, COD_USUARIOCREA, FEC_CREA, HOR_CREA,"
         vgSql = vgSql & "  COD_USUARIOMODI, FEC_MODI, HOR_MODI, COD_SUCCORREDOR, FEC_FINPERDIF,"
         vgSql = vgSql & "  FEC_FINPERGAR, GLS_NACIONALIDAD, IND_RECALCULO, FEC_EMISION, FEC_INIPENCIA,"
         vgSql = vgSql & "  MTO_RMGTOSEPRV, FEC_CALCULO, MTO_AJUSTEIPC, MTO_APOADIFON, MTO_APOADIMOD,"
         vgSql = vgSql & "  COD_TIPVIA, GLS_NOMVIA, GLS_NUMDMC, GLS_INTDMC, COD_TIPZON,"
         vgSql = vgSql & "  GLS_NOMZON, GLS_REFERENCIA, COD_TIPREAJUSTE, MTO_VALREAJUSTETRI, MTO_VALREAJUSTEMEN, GLS_TELBEN2, IND_BENDES)"
         vgSql = vgSql & " select  '" & num_poliza & "', "
         vgSql = vgSql & num_endoso + 1 & " , NUM_COT, NUM_CORRELATIVO, NUM_OPERACION, NUM_ARCHIVO,"
         vgSql = vgSql & " COD_TRAPAGOPEN, FEC_TRAPAGOPEN, COD_AFP, COD_ISAPRE, COD_TIPPENSION,"
         vgSql = vgSql & " COD_VEJEZ, COD_ESTCIVIL, COD_CUSPP, COD_TIPOIDENAFI, NUM_IDENAFI,"
         vgSql = vgSql & " GLS_DIRECCION, COD_DIRECCION, GLS_FONO, GLS_CORREO, COD_VIAPAGO,"
         vgSql = vgSql & " COD_TIPCUENTA, COD_BANCO, NUM_CUENTA, COD_SUCURSAL, FEC_SOLICITUD, FEC_INGVIGENCIA,"
         vgSql = vgSql & " FEC_VIGENCIA, FEC_DEV, FEC_ACEPTA, FEC_PRIPAGO, COD_MONEDAFON, MTO_MONEDAFON,"
         vgSql = vgSql & " MTO_PRIUNIFON, MTO_CTAINDFON, MTO_BONOFON, MTO_APOADI, MTO_PRIUNI, MTO_CTAIND,"
         vgSql = vgSql & " MTO_BONO, PRC_TASARPRT, COD_TIPOIDENCOR, NUM_IDENCOR, PRC_CORCOM, MTO_CORCOM, PRC_CORCOMREAL,"
         vgSql = vgSql & " NUM_ANNOJUB, NUM_CARGAS, IND_COB, COD_BENSOCIAL, COD_MONEDA, MTO_VALMONEDA, MTO_PRIUNIMOD,"
         vgSql = vgSql & " MTO_CTAINDMOD, MTO_BONOMOD, COD_TIPREN, NUM_MESDIF, COD_MODALIDAD, NUM_MESGAR,"
         vgSql = vgSql & " COD_COBERCON, MTO_FACPENELLA, PRC_FACPENELLA, COD_DERCRE, COD_DERGRA, PRC_RENTAAFP,"
         vgSql = vgSql & " PRC_RENTAAFPORI, PRC_RENTATMP, MTO_CUOMOR, PRC_TASACE, PRC_TASAVTA, PRC_TASATIR,"
         vgSql = vgSql & " PRC_TASAPERGAR, MTO_CNU,MTO_PRIUNISIM, MTO_PRIUNIDIF, MTO_PENSION, MTO_PENSIONGAR,"
         vgSql = vgSql & " MTO_CTAINDAFP, MTO_RENTATMPAFP, MTO_RESMAT, MTO_VALPREPENTMP, MTO_PERCON,"
         vgSql = vgSql & " PRC_PERCON, MTO_SUMPENSION, MTO_PENANUAL, MTO_RMPENSION, MTO_RMGTOSEP, COD_TIPCOT,"
         vgSql = vgSql & " COD_ESTCOT, COD_USUARIO, COD_SUCURSALUSU, FEC_INIPAGOPEN, COD_USUARIOCREA,FEC_CREA,"
         vgSql = vgSql & " HOR_CREA, 'SEACSA', '" & FecServ & "', '" & HoraActual & "', COD_SUCCORREDOR, FEC_FINPERDIF,"
         vgSql = vgSql & " FEC_FINPERGAR, GLS_NACIONALIDAD, IND_RECALCULO, FEC_EMISION, FEC_INIPENCIA,"
         vgSql = vgSql & " MTO_RMGTOSEPRV, FEC_CALCULO, MTO_AJUSTEIPC, MTO_APOADIFON, MTO_APOADIMOD, COD_TIPVIA,"
         vgSql = vgSql & " GLS_NOMVIA, GLS_NUMDMC, GLS_INTDMC, COD_TIPZON, GLS_NOMZON, GLS_REFERENCIA, COD_TIPREAJUSTE,"
         vgSql = vgSql & " Mto_ValReajusteTri , Mto_ValReajusteMen, GLS_TELBEN2, IND_BENDES "
         vgSql = vgSql & " From PD_TMAE_POLIZA"
         vgSql = vgSql & " Where num_poliza =  '" & num_poliza & "' and num_endoso='" & num_endoso & "' " '(select max(num_endoso) from pd_tmae_poliza where num_poliza='" & num_poliza & "')"
         vgConexionBD.Execute vgSql
    For i = 1 To vlNumCargas
        With stBeneficiariosMod(i)
            'Cargar Datos en pp_tmae_ben
            vgSql = "INSERT INTO PP_TMAE_BEN "
            vgSql = vgSql & "(num_poliza,num_endoso,num_orden,fec_ingreso,Cod_TipoIdenben,Num_Idenben, "
            vgSql = vgSql & "gls_nomben,gls_nomsegben,gls_patben,gls_matben,gls_dirben,cod_direccion, "
            vgSql = vgSql & "gls_fonoben, gls_correoben, cod_grufam, cod_par,cod_sexo,cod_sitinv,cod_dercre,cod_derpen, "
            vgSql = vgSql & "cod_cauinv,fec_nacben, fec_nachm, fec_invben,cod_motreqpen, "
            vgSql = vgSql & "mto_pension,mto_pensiongar,prc_pension, cod_inssalud,cod_modsalud,mto_plansalud, "
            vgSql = vgSql & "cod_estpension, cod_viapago,cod_banco, cod_tipcuenta, num_cuenta,"
            vgSql = vgSql & "cod_sucursal, fec_fallben, fec_matrimonio, cod_caususben, fec_susben, "
            vgSql = vgSql & "fec_inipagopen, fec_terpagopengar, cod_usuariocrea,fec_crea,hor_crea "
            vgSql = vgSql & ",prc_pensiongar, prc_pensionleg ,gls_telben2,cod_tipcta, cod_monbco, num_ctabco) "
            vgSql = vgSql & " SELECT "
            vgSql = vgSql & " num_poliza," & num_endoso + 1 & ",num_orden,fec_ingreso,Cod_TipoIdenben,Num_Idenben, "
            vgSql = vgSql & "gls_nomben,gls_nomsegben,gls_patben,gls_matben,gls_dirben,cod_direccion, "
            vgSql = vgSql & "gls_fonoben, gls_correoben, cod_grufam, cod_par,cod_sexo,cod_sitinv,'" & .Cod_DerCre & "'," & .Cod_DerPen & ", "
            vgSql = vgSql & "cod_cauinv,fec_nacben, fec_nachm, fec_invben,cod_motreqpen, "
            vgSql = vgSql & str(.Mto_Pension) & "," & str(.Mto_PensionGar) & "," & str(.Prc_Pension) & ", cod_inssalud,cod_modsalud,mto_plansalud, "
            vgSql = vgSql & "cod_estpension, cod_viapago,cod_banco, cod_tipcuenta, num_cuenta,"
            vgSql = vgSql & "cod_sucursal, fec_fallben, fec_matrimonio, cod_caususben, fec_susben, "
            vgSql = vgSql & "fec_inipagopen, fec_terpagopengar,  'SEACSA', '" & FecServ & "', '" & HoraActual & "', "
            vgSql = vgSql & str(.Prc_PensionGar) & ", " & str(.Prc_PensionLeg) & " ,gls_telben2,cod_tipcta, cod_monbco, num_ctabco "
            vgSql = vgSql & " from PP_TMAE_BEN "
            vgSql = vgSql & " WHERE NUM_POLIZA='" & num_poliza & "' AND NUM_ORDEN=" & .Num_Orden & ""
            vgSql = vgSql & " AND NUM_ENDOSO=" & num_endoso & " "
            vgConexionBD.Execute vgSql
            'Cargar Datos en pd_tmae_polben(?)
            vgSql = "Insert into PD_TMAE_POLBEN"
            vgSql = vgSql & " (NUM_POLIZA, NUM_ENDOSO, NUM_ORDEN, COD_PAR, COD_GRUFAM,"
            vgSql = vgSql & " COD_SEXO, COD_SITINV, FEC_INVBEN, COD_CAUINV, COD_DERPEN,"
            vgSql = vgSql & " COD_DERCRE, COD_TIPOIDENBEN, NUM_IDENBEN, GLS_NOMBEN, GLS_NOMSEGBEN,"
            vgSql = vgSql & " GLS_PATBEN, GLS_MATBEN, FEC_NACBEN, FEC_FALLBEN, FEC_NACHM,"
            vgSql = vgSql & " PRC_PENSION, PRC_PENSIONLEG, PRC_PENSIONREP, MTO_PENSION, MTO_PENSIONGAR,"
            vgSql = vgSql & " COD_USUARIOCREA, FEC_CREA, HOR_CREA, COD_USUARIOMODI, FEC_MODI,"
            vgSql = vgSql & " HOR_MODI, COD_ESTPENSION, PRC_PENSIONGAR)"
            vgSql = vgSql & " select '" & num_poliza & "', " & num_endoso + 1 & ", NUM_ORDEN, COD_PAR, COD_GRUFAM,"
            vgSql = vgSql & " COD_SEXO, COD_SITINV, FEC_INVBEN, COD_CAUINV, COD_DERPEN,"
            vgSql = vgSql & " COD_DERCRE, COD_TIPOIDENBEN, NUM_IDENBEN, "
            vgSql = vgSql & " GLS_NOMBEN, GLS_NOMSEGBEN, GLS_PATBEN,  GLS_MATBEN, "
            vgSql = vgSql & " Fec_NacBen , Fec_FallBen, Fec_NacHM, "
            vgSql = vgSql & " PRC_PENSION, PRC_PENSIONLEG, PRC_PENSIONREP, MTO_PENSION, MTO_PENSIONGAR,"
            vgSql = vgSql & " 'SEACSA', '" & FecServ & "', '" & HoraActual & "', COD_USUARIOMODI, FEC_MODI,"
            vgSql = vgSql & " HOR_MODI, COD_ESTPENSION, PRC_PENSIONGAR from PD_TMAE_POLBEN "
            vgSql = vgSql & " WHERE NUM_POLIZA='" & num_poliza & "' AND NUM_ORDEN=" & i & ""
            vgSql = vgSql & " AND NUM_ENDOSO=" & num_endoso & " " ' (select max(num_endoso) from pd_tmae_polben where num_poliza='" & vlGlobalNumPoliza & "' and num_orden=" & vlNumOrden & ") "
            vgConexionBD.Execute vgSql
        End With
    Next
Exit Function
Error_flActualizarPor:
    'Eliminar registros
    vgSql = "delete from pp_tmae_endoso where cod_usuariocrea='SEACSA' and fec_crea='" & FecServ & "';"
    vgConexionBD.Execute vgSql
    vgSql = "delete from pp_tmae_ben where cod_usuariocrea='SEACSA' and fec_crea='" & FecServ & "';"
    vgConexionBD.Execute vgSql
    vgSql = "delete from pd_tmae_polben where cod_usuariocrea='SEACSA' and fec_crea='" & FecServ & "';"
    vgConexionBD.Execute vgSql
    vgSql = "delete from pp_tmae_poliza where cod_usuariocrea='SEACSA' and fec_crea='" & FecServ & "';"
    vgConexionBD.Execute vgSql
    vgSql = "delete from pd_tmae_poliza where cod_usuariomodi='SEACSA' and FEC_MODI='" & FecServ & "';"
    vgConexionBD.Execute vgSql
End Function

'Materia Gris - Jaime Rios 19/02/2018
Private Function fgCalcularPorcentajeB(num_poliza As String, iFechaIniVig As String, iNumBenef As Integer, _
ostBeneficiarios() As TyBeneficiarios, Optional iTipoPension As String, _
Optional iPensionRef As Double, Optional iCalcularPension As Boolean, _
Optional iDerCrecerCotizacion As String, Optional iCobCobertura As String, _
Optional iCalcularPorcentaje As Boolean, Optional iPerGar As Long, _
Optional iPensionRefGar As Double) As Boolean

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
'I--- ABV 21/06/2007 ---
'On Error GoTo Err_fgCalcularPorcentaje
'F--- ABV 21/06/2007 ---

    vbEsEdadLeg = False
    fgCalcularPorcentajeB = False
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
    vlEsEdadLeg = flObtieneEdadNormativa(num_poliza) '(num_pol)
    
    If vlEsEdadLeg > L18 Then
        vbEsEdadLeg = True
    End If


'F--- ABV 21/07/2006 ---
    
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
    
    
    Dim cont_hijosS As Integer
    Dim SitInv As String
    Dim vf As Boolean
    
    vf = False
    
    vlFechaFallCau = ""
    vlValorHijo = 0
    cont_hijosS = 0
'F--- ABV 18/07/2006 ---
    
    Dim num_pol As String
    num_pol = num_poliza 'vgNum_pol
    
If (iCalcularPorcentaje = True) Then
    vlContBen = 1 '0
    i = 1
    Do While i <= vlNum_Ben
    
        vlContBen = vlContBen + 1

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
        
            'Derecho Pensión
            derpen(i) = ostBeneficiarios(i).Cod_DerPen
            estpen(i) = ostBeneficiarios(i).Cod_EstPension
        
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
        If Codrel(i) = 99 Then
            vlFechaFallCau = vlFechaFallecimiento
        End If
'F--- ABV 18/07/2006 ---

        vlFechaMatrimonio = ostBeneficiarios(i).Fec_Matrimonio
    'RRR 11/06/2013
        
            If (Codrel(i) >= 30 And Codrel(i) < 40) Then
                If vbEsEdadLeg = True Then
                    L18 = IIf(flCompletaRequisitos(num_pol, CInt(i)) = True, L24, L18)
                End If
                If vlFecTerPerGar > iFechaIniVig Then
                    derpen(i) = 99
                    PorbenGar(i) = ostBeneficiarios(i).Prc_PensionGar
                    vf = True
                End If
                If (vlFechaFallecimiento <> "") Or (vlFechaMatrimonio <> "") Then
                    derpen(i) = 10
                     
                    If ostBeneficiarios(i).Cod_EstPension <> "10" Then
                        ostBeneficiarios(i).Cod_EstPension = "10"
                        Hijos_SinDerechoPension(ostBeneficiarios(i).Cod_GruFam) = Hijos_SinDerechoPension(ostBeneficiarios(i).Cod_GruFam) + 1
                    End If
                Else
                    If edad_mes_ben > L18 And Coinbe(i) = "N" Then
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
                        End If
                    End If
                End If
                cont_hijosS = cont_hijosS + 1
            Else
                'El resto de los Beneficiarios que no sean Hijos, solo se dejan como
                'Sin Derecho a Pensión cuando están fallecidos
                If (vlFechaFallecimiento <> "") Then
                    If (iTipoPension <> "08" And iTipoPension <> "09" And iTipoPension <> "10" And iTipoPension <> "11" And iTipoPension <> "12") Then
                        derpen(i) = 10
                        If vlFecTerPerGar > iFechaIniVig Then
                            'derpen(i) = 99
                            PorbenGar(i) = ostBeneficiarios(i).Prc_PensionGar
                        End If
                    Else
                        If Codrel(i) <> "99" Then
                            derpen(i) = 10
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
                   
                Else
                    
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
                        If Coinbe(g) = "N" And edad_mes_ben <= L18 Then
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
            End Select
        End If
    Next g
                                
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
                            If flSiDebeCambiarParentesco(num_pol, iFechaIniVig) = 1 Then
                                If Hijos_SinDerechoPension(u) > cont_hijo Then
                                    Ncorbe(j) = 10
                                    ostBeneficiarios(j).Cod_Par = Ncorbe(j)
                                End If
                            End If
                        Else

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
                            Porben(j) = CDbl(Format(vlValor / cont_esposa, "#0.00"))
                                         
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
                            If flSiDebeCambiarParentesco(num_pol, iFechaIniVig) = 1 Then
                                If Hijos_SinDerechoPension(u) > cont_hijo Then
                                    Ncorbe(j) = 20
                                    ostBeneficiarios(j).Cod_Par = Ncorbe(j)
                                End If
                            End If
                        Else

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
                        If Coinbe(j) = "N" And edad_mes_ben > L18 Then
                            Porben(j) = 0
                        Else
                            If vbEsEdadLeg = True Then
                                L18 = IIf(flCompletaRequisitos(num_pol, CInt(i)) = True, L24, L18)
                            End If
                            
                            If (Coinbe(j) = "P" Or Coinbe(j) = "T") And edad_mes_ben > L18 Then
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
                        End If
                    Else
                        Q = Cod_Grfam(j)
                        Codcbe(j) = "N"

                        If cont_esposa = 0 And cont_mhn(Q) = 0 Then
                            If vbEsEdadLeg = True Then
                                L18 = IIf(flCompletaRequisitos(num_pol, CInt(i)) = True, L24, L18)
                            End If
                            If (Coinbe(j) = "P" Or Coinbe(j) = "T") And edad_mes_ben > L18 Then
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
                                If Coinbe(j) = "N" And edad_mes_ben <= L18 Then
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
    
    vlSumarTotalPorcentajePension = 0
    vlSumaPjePenPadres = 0
    j = 0
    For j = 1 To vlNum_Ben

        If IsNumeric(Porben(j)) Then
            'I--- ABV 22/06/2005 ---
            'ostBeneficiarios(j).Prc_Pension = Format(Porben(j), "#0.00")
            If (ostBeneficiarios(j).Cod_Par = "99" Or ostBeneficiarios(j).Cod_Par = "0") Then
                If (derpen(j) <> 10) Then
                    If (vlFechaFallecimiento <> "" And cont_causante = 1) And vlFecTerPerGar > iFechaIniVig Then
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
                    ostBeneficiarios(j).Prc_Pension = Format(Porben(j), "#0.00")
                    ostBeneficiarios(j).Prc_PensionLeg = Format(Porben(j), "#0.00") '*-+
                ostBeneficiarios(j).Prc_PensionGar = Format(PorbenGar(j), "#0.00")
            End If
            'F--- ABV 22/06/2005 ---
        Else
            ostBeneficiarios(j).Prc_Pension = 0
            ostBeneficiarios(j).Prc_PensionLeg = 0 '*-+
            ostBeneficiarios(j).Prc_PensionGar = 0
        End If
        
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
                ostBeneficiarios(j).Mto_PensionGar = Format(0, "#0.00")
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

        If vgTipoCauEnd = "S" Then ostBeneficiarios(j).Cod_EstPension = derpen(j)
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
                    SumaTotalDef = SumaTotalDef + ostBeneficiarios(j).Prc_Pension
                End If
            Next j
            SumaTotalDefGar = 0
            For j = 1 To vlNum_Ben
                If (ostBeneficiarios(j).Cod_DerPen <> "10") And (ostBeneficiarios(j).Cod_Par <> "99" And ostBeneficiarios(j).Cod_Par <> "0") Then
                    vlPorcentajeRecal = (ostBeneficiarios(j).Prc_Pension / (SumaTotalDef)) * 100
                    ostBeneficiarios(j).Prc_Pension = Format((vlPorcentajeRecal / fx), "#0.00")
                    ostBeneficiarios(j).Prc_PensionLeg = Format((vlPorcentajeRecal / fx), "#0.00")
                    'ostBeneficiarios(j).Prc_PensionGar = Format((ostBeneficiarios(j).Prc_Pension / SumaTotalDef) * 100, "#0.00")
                    If ostBeneficiarios(j).Cod_Par <> "99" Then
                        SumaTotalDefGar = SumaTotalDefGar + ostBeneficiarios(j).Prc_Pension
                    End If
                End If
            Next j
            'OBTIENE EL PORCENTAJE GARANTIZADO
            For j = 1 To vlNum_Ben
                If (ostBeneficiarios(j).Cod_DerPen <> "10") And (ostBeneficiarios(j).Cod_Par <> "99" And ostBeneficiarios(j).Cod_Par <> "0") Then
                    ostBeneficiarios(j).Prc_Pension = Format(((ostBeneficiarios(j).Prc_Pension / (SumaTotalDef)) * 100), "#0.00")
                    ostBeneficiarios(j).Prc_PensionLeg = Format(((ostBeneficiarios(j).Prc_Pension / (SumaTotalDef)) * 100), "#0.00")
                    ostBeneficiarios(j).Prc_PensionGar = Format((ostBeneficiarios(j).Prc_Pension / SumaTotalDefGar) * 100, "#0.00")
                End If
            Next j
            
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
        If iTipoPension = "08" Then
            If (SumaTotalDef > 100) Then
                For j = 1 To vlNum_Ben
                        If (ostBeneficiarios(j).Cod_DerPen <> "10") And (ostBeneficiarios(j).Cod_Par <> "99" And ostBeneficiarios(j).Cod_Par <> "0") Then
                            vlPorcentajeRecal = Format((ostBeneficiarios(j).Prc_Pension / SumaTotalDef) * 100, "#0.00")
                            ostBeneficiarios(j).Prc_Pension = Format(vlPorcentajeRecal, "#0.00")
                            ostBeneficiarios(j).Prc_PensionLeg = Format(vlPorcentajeRecal, "#0.00")
                            ostBeneficiarios(j).Prc_PensionGar = Format(vlPorcentajeRecal, "#0.00")
                        End If
                Next j
            ElseIf (SumaTotalDef > 0) Then
                For j = 1 To vlNum_Ben
                        If (ostBeneficiarios(j).Cod_DerPen <> "10") And (ostBeneficiarios(j).Cod_Par <> "99" And ostBeneficiarios(j).Cod_Par <> "0") Then
                            ostBeneficiarios(j).Prc_PensionLeg = Format(ostBeneficiarios(j).Prc_Pension, "#0.00")
                            ostBeneficiarios(j).Prc_PensionGar = Format((ostBeneficiarios(j).Prc_Pension / SumaTotalDef) * 100, "#0.00")
                        End If
                Next j
            End If
        Else
            If (SumaTotalDef > 100) Then
                For j = 1 To vlNum_Ben
                        If (ostBeneficiarios(j).Cod_DerPen <> "10") And (ostBeneficiarios(j).Cod_Par <> "99" And ostBeneficiarios(j).Cod_Par <> "0") Then
                            vlPorcentajeRecal = Format(ostBeneficiarios(j).Prc_Pension, "#0.00")
                            ostBeneficiarios(j).Prc_Pension = Format(vlPorcentajeRecal, "#0.00")
                            ostBeneficiarios(j).Prc_PensionLeg = Format(vlPorcentajeRecal, "#0.00")
                            ostBeneficiarios(j).Prc_PensionGar = Format((vlPorcentajeRecal / SumaTotalDef) * 100, "#0.00")
                        End If
                Next j
            Else
                For j = 1 To vlNum_Ben
                        If (ostBeneficiarios(j).Cod_DerPen <> "10") And (ostBeneficiarios(j).Cod_Par <> "99" And ostBeneficiarios(j).Cod_Par <> "0" And ostBeneficiarios(j).Cod_EstPension <> "10" And ostBeneficiarios(j).Cod_EstPension <> "20") Then 'mvg agrego Cod_EstPension and else
                            ostBeneficiarios(j).Prc_PensionLeg = Format(ostBeneficiarios(j).Prc_Pension, "#0.00")
                            ostBeneficiarios(j).Prc_PensionGar = Format((ostBeneficiarios(j).Prc_Pension / SumaTotalDef) * 100, "#0.00")
                        Else
                            ostBeneficiarios(j).Prc_PensionLeg = Format(ostBeneficiarios(j).Prc_Pension, "#0.00")
                            ostBeneficiarios(j).Prc_PensionGar = Format((ostBeneficiarios(j).Prc_PensionGar), "#0.00")
                        End If
                Next j
            End If
        End If
        
        
    End If

'PARA LOS COSOS DE JUBILACION
    If (iTipoPension = "09" Or iTipoPension = "10") And (iCobCobertura = "N") Then
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
    End If
    
'F--- ABV 04/08/2007 ---
'*-+
End If


'If (vlFechaFallecimiento <> "" And cont_causante = 0) And vlFecTerPerGar > iFechaIniVig Then ' Jaime Rios 21/02/2018
    iCalcularPension = True
'End If

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

    fgCalcularPorcentajeB = True
    
Exit Function
Err_fgCalcularPorcentaje:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        vgError = 1000
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'Materia Gris - Jaime Rios 19/02/2018
Function fgCarga_Param(iTabla As String, iElemento As String, iFecha As String) As Boolean
Dim vlRegistro As ADODB.Recordset
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

End Function





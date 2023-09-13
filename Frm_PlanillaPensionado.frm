VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "Crystl32.OCX"
Begin VB.Form Frm_PlanillaPensionado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planilla Masiva y por Pensionado."
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6225
   Begin VB.Frame Fra_Datos 
      Height          =   3495
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   6015
      Begin VB.Frame Frame1 
         Height          =   1455
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   5775
         Begin VB.TextBox Txt_NumIdent 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3720
            MaxLength       =   16
            TabIndex        =   18
            Top             =   960
            Width           =   1875
         End
         Begin VB.ComboBox Cmb_TipoIdent 
            BackColor       =   &H00E0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   960
            Width           =   2235
         End
         Begin VB.TextBox Txt_Poliza 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1440
            MaxLength       =   10
            TabIndex        =   5
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox Chk_Pensionado 
            Caption         =   "Datos del Pensionado :"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "N° Identificación :"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   19
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Lbl_Contrato 
            Caption         =   "Nº de Póliza        :"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Width           =   1455
         End
      End
      Begin VB.ComboBox Cmb_Pago 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   840
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
      Begin VB.TextBox Txt_Desde 
         Height          =   285
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Txt_Hasta 
         Height          =   285
         Left            =   4320
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1320
         Width           =   1215
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
         TabIndex        =   14
         Top             =   1320
         Width           =   2295
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
         Index           =   7
         Left            =   360
         TabIndex        =   13
         Top             =   840
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
         Index           =   6
         Left            =   360
         TabIndex        =   12
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Lbl_Contrato 
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
         Index           =   4
         Left            =   4080
         TabIndex        =   11
         Top             =   1320
         Width           =   135
      End
   End
   Begin VB.Frame Fra_Operaciones 
      Height          =   1095
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   3480
      Width           =   6015
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   1440
         Picture         =   "Frm_PlanillaPensionado.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   2640
         Picture         =   "Frm_PlanillaPensionado.frx":06BA
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   3840
         Picture         =   "Frm_PlanillaPensionado.frx":0D74
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
      Begin Crystal.CrystalReport Rpt_Reporte 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowState     =   2
         PrintFileLinesPerPage=   60
      End
   End
End
Attribute VB_Name = "Frm_PlanillaPensionado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'----------------------------------
'Conciliacion

Dim vlRegistroCia As ADODB.Recordset
Dim vlRegistroTes As ADODB.Recordset
Dim vlRegistroRes As ADODB.Recordset
Dim vlRegistroPeriodo As ADODB.Recordset

Dim vlDiaIni As String
Dim vlMesIni As String
Dim vlAnnoIni As String
Dim vlDiaTer As String
Dim vlMesTer As String
Dim vlAnnoTer As String

Dim vlNumPoliza As String
Dim vlNumOrden As Integer
Dim vlNumEndoso As Integer
Dim vlCodTipoImp As String
Dim vlNumPeriodo As String
Dim vlFecPago As String
Dim vlCodTipRes As String
Dim vlNumResGarEst As Integer
Dim vlNumAnnoRes As Integer
Dim vlNumDias As Integer
Dim vlMtoPension As Double
Dim vlMtoPensionUF As Double
Dim vlMtoPensionMin As Double
Dim vlPrcDeduccion As Double
Dim vlMtoDeduccion As Double
Dim vlMtoGarEstQui As Double
Dim vlMtoGarEstNor As Double
Dim vlMtoGarEstCia As Double
Dim vlMtoGarEstRec As Double
Dim vlMtoDiferencia As Double
Dim vlCodDerGarEst As String
Dim vlCodEstado As String
Dim vlMtoHaber As Double
Dim vlMtoDescuento As Double

Dim vlCodPar As String
Dim vlCodSexo As String
Dim vlCodSitInv As String
Dim vlFecNacBen As String
Dim vlEdadBen As Integer

Dim vlPeriodoInicio As String
Dim vlPeriodoTer As String
Dim vlNumPerPago As String

'----------------------------------
Const clCodConHDHab As String * 2 = "04" 'MA_TPAR_TABCOD / Garantía Estatal RetroActiva
Const clCodConHDDes As String * 2 = "33" 'MA_TPAR_TABCOD / Descuento Pensionado por Garantía Estatal
Const clCodConHDMtoPen As String * 2 = "01" 'MA_TPAR_TABCOD / Pensión Renta Vitalicia
Const clCodConHDGECia As String * 2 = "03"
Const clModOrigenPP As String * 2 = "PP" 'Pago de Pensiones
Const clCodGE As String * 2 = "03"
Const clCodBonInv As String * 2 = "06"
Const clCodHabGE As String = "('02','04','07')"
Const clCodDesGE As String = "('21','22')"
Const clTipoGE As String * 2 = "GE" 'Garantía Estatal
Const clCodH As String * 1 = "H" 'Haber
Const clCodD As String * 1 = "D" 'Descuento
Const clTipoImpC As String * 1 = "C" 'Conciliación
Const clTipoImpE As String * 1 = "E" 'Exceso
Const clTipoImpD As String * 1 = "D" 'Deficit
Const clSinCodDerGE As String * 1 = "N" 'Sin Estado


'CMV-20060803 I
Const clMesEnero As String = "ENERO"
Const clMesFebrero As String = "FEBRERO"
Const clMesMarzo As String = "MARZO"
Const clMesAbril As String = "ABRIL"
Const clMesMayo As String = "MAYO"
Const clMesJunio As String = "JUNIO"
Const clMesJulio As String = "JULIO"
Const clMesAgosto As String = "AGOSTO"
Const clMesSeptiembre As String = "SEPTIEMBRE"
Const clMesOctubre As String = "OCTUBRE"
Const clMesNoviembre As String = "NOVIEMBRE"
Const clMesDiciembre As String = "DICIEMBRE"

Dim vlPerPagoNum As String
Dim vlPerPagoTxt As String
'CMV-20060803 F
'-----------------------------------------------------------------------FIN

Dim vlFechaDesde As String, vlFechaHasta As String
Dim vlNumPol As String ', vlRut As String, vlDigito As String
Dim vlNumIdent As String, vlCodTipoIdent As Long
Dim vlRutCliente As String

Dim vlOpcion        As String
Dim vlPago          As String
Dim vlArchivo       As String
Dim vlGlosaOpcion   As String

Dim vlNombreCompania As String
Dim vlDirCompania As String
Dim vlFonoCompania As String

Dim vlLargoTipoIden    As Integer 'sirve para llenar la grilla
Dim vlPosicionTipoIden As Integer 'sirve para llenar la grilla
'Dim vlNumIdent         As Integer
'Dim vlCodTipoIden As Integer 'sirve para guardar el código

Private Function flRutCliente() As String
   'Buscar Rut del cliente
   flRutCliente = ""
   vgSql = "SELECT cli.num_idencli, "
   vgSql = vgSql & "cli.gls_dircli,cli.gls_comcli,cli.gls_ciucli, "
   vgSql = vgSql & "cli.gls_fonocli,cli.gls_faxcli, "
   vgSql = vgSql & "cli.gls_nomlarcli,cli.gls_correocli,cli.gls_paiscli "
   vgSql = vgSql & "FROM ma_tmae_cliente cli "
   Set vgRegistro = vgConexionBD.Execute(vgSql)
   If Not vgRegistro.EOF Then
        If Not IsNull(vgRegistro!gls_nomlarcli) Then
            vlNombreCompania = Trim(vgRegistro!gls_nomlarcli)
        Else
            vlNombreCompania = ""
        End If
        If Not IsNull(vgRegistro!gls_dircli) Then
            vlDirCompania = Trim(vgRegistro!gls_dircli)
        Else
            vlDirCompania = ""
        End If
        If Not IsNull(vgRegistro!gls_comcli) Then
            vlDirCompania = vlDirCompania & ", " & Trim(vgRegistro!gls_comcli)
        End If
        If Not IsNull(vgRegistro!gls_ciucli) Then
            vlDirCompania = vlDirCompania & ", " & Trim(vgRegistro!gls_ciucli)
        End If
        If Not IsNull(vgRegistro!gls_paiscli) Then
            vlDirCompania = vlDirCompania & ", " & Trim(vgRegistro!gls_paiscli)
        End If
        If Not IsNull(vgRegistro!gls_fonocli) Then
            vlFonoCompania = "Teléfono: " & Trim(vgRegistro!gls_fonocli)
        Else
            vlFonoCompania = ""
        End If
        If Not IsNull(vgRegistro!gls_faxcli) Then
            vlFonoCompania = vlFonoCompania & ", Fax: " & Trim(vgRegistro!gls_faxcli)
        End If
        If Not IsNull(vgRegistro!gls_correocli) Then
            vlFonoCompania = vlFonoCompania & ". e-mail: " & Trim(vgRegistro!gls_correocli)
        End If
      'vlDirCompania = Trim(vgRegistro!gls_dircli) & ", " & Trim(vgRegistro!gls_comcli) & ", " & Trim(vgRegistro!gls_ciucli) & ", " & Trim(vgRegistro!gls_paiscli)
      'vlFonoCompania = "Teléfono: " & Trim(vgRegistro!gls_fonocli) & ", Fax: " & Trim(vgRegistro!gls_faxcli) & ". e-mail: " & Trim(vgRegistro!gls_correocli)
      flRutCliente = (Trim(vgRegistro!num_idencli))
   End If
'   vlNombreCompania = "Le Mans Desarrollo Compania de Seguros de Vida S.A. (en Quiebra)"
'   vlDirCompania = "Encomenderos N° 113, piso 2, Las Condes, Santiago, Chile"
'   vlFonoCompania = "Teléfono: 378 70 14, Fax: 246 08 55. e-mail: lemansvida@lemans.cl"
   
End Function

Function flInformeLiqPagoCarta()
'Imprime Liquidaciones de Pension
On Error GoTo Err_flInformeLiqPagoCarta
    Dim vlSQLQuery As String
    Dim vlMesPerPago As String
    Dim vlAnnoPerPago As String
    
    vlFechaDesde = vlFechaDesde
    vlFechaHasta = vlFechaHasta
    
'    vlPerPagoNum = ""
    
    If Not flLlenaTemporal(vlFechaDesde, vlFechaHasta) Then
        Exit Function
    End If
    
    vlArchivo = strRpt & "PP_Rpt_LiquidacionRVLMD.rpt"
    If Not fgExiste(vlArchivo) Then
        MsgBox "Archivo de Reporte no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
        Screen.MousePointer = 0
        Exit Function
    End If
    
    Call fgVigenciaQuiebra(Txt_Desde)
        
    vlNombreCompania = ""
    vlDirCompania = ""
    vlFonoCompania = ""
    vlRutCliente = flRutCliente 'vgRutCliente + " - " + vgDgvCliente
    
'    vlMesPerPago = ""
'    vlAnnoPerPago = ""
'    vlPerPagoTxt = ""
'
'    vlMesPerPago = Mid(vlPerPagoNum, 5, 2)
'    vlAnnoPerPago = Mid(vlPerPagoNum, 1, 4)
'    vlPerPagoTxt = flPerPago(vlMesPerPago)
'    vlPerPagoTxt = "MES:  " & vlPerPagoTxt & " - " & vlAnnoPerPago
    
    vgQuery = "{PP_TTMP_LIQUIDACION2.COD_USUARIO} = '" & vgUsuario & "'"
    Rpt_Reporte.Reset
'    Rpt_Reporte.SQLQuery = vlSQLQuery
    Rpt_Reporte.SelectionFormula = vgQuery
    Rpt_Reporte.WindowState = crptMaximized
    Rpt_Reporte.ReportFileName = vlArchivo
    Rpt_Reporte.Connect = vgRutaDataBase
    Rpt_Reporte.Formulas(5) = ""
    Rpt_Reporte.Formulas(6) = ""
    Rpt_Reporte.Formulas(7) = ""
    Rpt_Reporte.Formulas(8) = ""
'    Rpt_Reporte.Formulas(9) = ""
    
    Rpt_Reporte.Formulas(0) = "NombreCompania='" & vgNombreCompania & "'"
    'Rpt_Reporte.Formulas(1) = "NombreSistema= '" & vgNombreSistema & "'"
    'Rpt_Reporte.Formulas(2) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
    Rpt_Reporte.Formulas(3) = "rutcliente= '" & vlRutCliente & "'"
    If Trim(vlGlosaOpcion) = "DEF" Then
        'Rpt_Reporte.Formulas(4) = "TipoProceso= 'DEFINITIVO' "
        Rpt_Reporte.Formulas(4) = "TipoProceso= '' "
    Else
        Rpt_Reporte.Formulas(4) = "TipoProceso= 'PROVISORIO' "
    End If
    Rpt_Reporte.Formulas(5) = "GlosaQuiebra = '" & vgGlosaQuiebra & "'"
    Rpt_Reporte.Formulas(6) = "NomCompania = '" & vlNombreCompania & "'"
    Rpt_Reporte.Formulas(7) = "DirCompania = '" & vlDirCompania & "'"
    Rpt_Reporte.Formulas(8) = "FonoCompania = '" & vlFonoCompania & "'"
'    Rpt_Reporte.Formulas(9) = "PerPagoTxt = '" & vlPerPagoTxt & "'"
    
    Rpt_Reporte.Destination = crptToWindow
    Rpt_Reporte.WindowTitle = "Informe de Liquidación de Rentas Vitalicias"
'    Rpt_Reporte.SubreportToChange = "PP_Rpt_MensajesPoliza.rpt"
'    Rpt_Reporte.Connect = vgRutaDataBase
    Rpt_Reporte.Action = 1
'    Rpt_Reporte.SubreportToChange = ""
    Screen.MousePointer = 0
Exit Function
Err_flInformeLiqPagoCarta:
    Screen.MousePointer = 0
    MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
End Function

Function flInformeLiqPago()
'Imprime Liquidaciones de Pension
On Error GoTo Err_VenTut
    Dim vlSQLQuery As String
    
    vlFechaDesde = vlFechaDesde
    vlFechaHasta = vlFechaHasta
    If Not flLlenaTemporal(vlFechaDesde, vlFechaHasta) Then
        Exit Function
    End If
    If vlPago = "R" Then
        vlArchivo = strRpt & "PP_Rpt_LiquidacionRV.rpt"
    Else
        vlArchivo = strRpt & "PP_Rpt_LiquidacionPrimerPagoRV.rpt"
    End If
    
    If Not fgExiste(vlArchivo) Then
        MsgBox "Archivo de Reporte no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
        Screen.MousePointer = 0
        Exit Function
    End If
    
    Call fgVigenciaQuiebra(Txt_Desde)

    
    vlRutCliente = flRutCliente 'vgRutCliente + " - " + vgDgvCliente
    vgQuery = "{PP_TTMP_LIQUIDACION2.COD_USUARIO} = '" & vgUsuario & "'"
    Rpt_Reporte.Reset
'    Rpt_Reporte.SQLQuery = vlSQLQuery
    Rpt_Reporte.SelectionFormula = vgQuery
    Rpt_Reporte.WindowState = crptMaximized
    Rpt_Reporte.ReportFileName = vlArchivo
    Rpt_Reporte.Connect = vgRutaDataBase
    Rpt_Reporte.Formulas(5) = ""
    Rpt_Reporte.Formulas(0) = "NombreCompania='" & UCase(vgNombreCompania) & "'"
    'Rpt_Reporte.Formulas(1) = "NombreSistema= '" & vgNombreSistema & "'"
    'Rpt_Reporte.Formulas(2) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
    Rpt_Reporte.Formulas(3) = "rutcliente= '" & vlRutCliente & "'"
    If Trim(vlGlosaOpcion) = "DEF" Then
        'Rpt_Reporte.Formulas(4) = "TipoProceso= 'DEFINITIVO' "
        Rpt_Reporte.Formulas(4) = "TipoProceso= '' "
    Else
        Rpt_Reporte.Formulas(4) = "TipoProceso= 'PROVISORIO' "
    End If
    Rpt_Reporte.Formulas(5) = "GlosaQuiebra = '" & vgGlosaQuiebra & "'"
    
    Rpt_Reporte.Destination = crptToWindow
    Rpt_Reporte.WindowTitle = "Informe de Liquidación de Rentas Vitalicias"
'    Rpt_Reporte.SubreportToChange = "PP_Rpt_MensajesPoliza.rpt"
'    Rpt_Reporte.Connect = vgRutaDataBase
    Rpt_Reporte.Action = 1
'    Rpt_Reporte.SubreportToChange = ""
    Screen.MousePointer = 0
Exit Function
Err_VenTut:
    Screen.MousePointer = 0
    MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
End Function


Function flLlenaTemporal(iFecDesde, iFecHasta) As Boolean

    Dim vlSql As String, vlTB As ADODB.Recordset
    Dim vlNumConceptosHab As Long, vlNumConceptosDesc As Long
    Dim vlItem As Long, vlPoliza As String, vlOrden As String
    Dim vlIndImponible As String, vlIndTributable As String
    Dim vlTipReceptor As String
    Dim vlNumIdenReceptor As String, vlCodTipoIdenReceptor As Long
    Dim vlPerPago As String
    Dim vlTB2 As ADODB.Recordset
    Dim vlTipPension As String
    Dim vlViaPago As String
    Dim vlCajaComp As String
    Dim vlInsSalud As String
    Dim vlCodDireccion As Double
    Dim vlAfp As String
    Dim vlMontoPesosPrimerPago As Double
    Dim vlFecPago As String
    Dim vlDescViaPago As String
    Dim vlSucursal As String
    Dim vlDescSucursal As String
    
    flLlenaTemporal = False
    vlItem = 1
    vlNumConceptosHab = 0
    vlNumConceptosDesc = 0
    vlPoliza = ""
    vlOrden = 0
    vlPerPago = ""
    vlFecPago = ""
    vlTipPension = ""
    vlViaPago = ""
    vlDescViaPago = ""
    vlCajaComp = ""
    vlInsSalud = ""
    vlCodDireccion = 0
    vlMontoPesosPrimerPago = 0
    vlAfp = ""
    vlSucursal = ""
    vlDescSucursal = ""
    'VARIABLES GENERALES
    stTTMPLiquidacion.Cod_Usuario = vgUsuario
    
    'Elimina Datos de la Tabla Temporal
    vlSql = "DELETE FROM PP_TTMP_LIQUIDACION WHERE COD_USUARIO = '" & vgUsuario & "'"
    vgConexionBD.Execute (vlSql)
    
    vlSql = "SELECT P.NUM_POLIZA, L.NUM_ENDOSO, P.NUM_ORDEN, P.COD_CONHABDES, P.MTO_CONHABDES, C.COD_TIPMOV, P.NUM_PERPAGO, P.NUM_IDENRECEPTOR, P.COD_TIPOIDENRECEPTOR, P.COD_TIPRECEPTOR,"
    vlSql = vlSql & " L.GLS_DIRECCION, L.FEC_PAGO, L.GLS_NOMRECEPTOR, L.GLS_NOMSEGRECEPTOR, L.GLS_PATRECEPTOR,"
    vlSql = vlSql & " L.GLS_MATRECEPTOR, L.MTO_LIQPAGAR, L.COD_DIRECCION, "
    vlSql = vlSql & " L.COD_TIPPENSION, L.COD_VIAPAGO, L.COD_SUCURSAL, L.COD_INSSALUD,"
    vlSql = vlSql & " L.MTO_PENSION, L.NUM_CARGAS, L.MTO_HABER, L.MTO_DESCUENTO,"
    vlSql = vlSql & " B.NUM_IDENBEN, B.COD_TIPOIDENBEN, B.GLS_NOMBEN, B.GLS_NOMSEGBEN, B.GLS_PATBEN, B.GLS_MATBEN, "
    vlSql = vlSql & " C.GLS_CONHABDES, M.COD_SCOMP, POL.COD_AFP, L.COD_MONEDA, M.GLS_ELEMENTO AS MONEDA"
    vlSql = vlSql & " FROM PP_TMAE_PAGOPEN" & vlGlosaOpcion & " P, PP_TMAE_LIQPAGOPEN" & vlGlosaOpcion & " L, MA_TPAR_CONHABDES C"
    vlSql = vlSql & ", PP_TMAE_POLIZA POL, PP_TMAE_BEN B, MA_TPAR_TABCOD M WHERE"
    vlSql = vlSql & " L.NUM_POLIZA = B.NUM_POLIZA AND"
    vlSql = vlSql & " L.NUM_ENDOSO = B.NUM_ENDOSO AND"
    vlSql = vlSql & " L.NUM_ORDEN = B.NUM_ORDEN AND"
    vlSql = vlSql & " L.NUM_POLIZA = POL.NUM_POLIZA AND"
    vlSql = vlSql & " L.NUM_ENDOSO = POL.NUM_ENDOSO AND"
    vlSql = vlSql & " L.COD_MONEDA = M.COD_ELEMENTO AND"
    vlSql = vlSql & " M.COD_TABLA = 'TM' AND" 'Tabla de Monedas
    If Chk_Pensionado.Value = 1 Then
        If Txt_Poliza <> "" Then
            vlSql = vlSql & " L.NUM_POLIZA = '" & Trim(Txt_Poliza) & "' AND"
        End If
        If Txt_NumIdent <> "" And Txt_NumIdent <> "0" Then
            vlSql = vlSql & " B.NUM_IDENBEN = '" & Trim(Txt_NumIdent) & "' AND"
            vlSql = vlSql & " B.COD_TIPOIDENBEN = " & Str(Mid(Cmb_TipoIdent.Text, 1, InStr(1, Cmb_TipoIdent.Text, "-") - 1)) & " AND"
        End If
'        If Cmb_TipoIdent.ListIndex <> -1 Then
'            If Str(Mid(Cmb_TipoIdent.Text, 1, InStr(1, Cmb_TipoIdent.Text, "-") - 1)) <> 0 Then
'            vlSql = vlSql & " B.COD_TIPOIDENBEN = " & Str(Mid(Cmb_TipoIdent.Text, 1, InStr(1, Cmb_TipoIdent.Text, "-") - 1)) & " AND"
'            End If
'        End If
    End If
    vlSql = vlSql & " L.NUM_POLIZA = P.NUM_POLIZA"
    vlSql = vlSql & " AND L.NUM_ORDEN = P.NUM_ORDEN"
    vlSql = vlSql & " AND L.NUM_IDENRECEPTOR = P.NUM_IDENRECEPTOR"
    vlSql = vlSql & " AND L.COD_TIPOIDENRECEPTOR = P.COD_TIPOIDENRECEPTOR"
    vlSql = vlSql & " AND L.COD_TIPRECEPTOR = P.COD_TIPRECEPTOR"
    vlSql = vlSql & " AND L.NUM_PERPAGO = P.NUM_PERPAGO"
    If vlPago = "P" Then 'PRIMER PAGO
        vlSql = vlSql & " AND L.COD_TIPOPAGO = 'P'"
    ElseIf vlPago = "R" Then 'PAGO EN REGIMEN
        vlSql = vlSql & " AND L.COD_TIPOPAGO = 'R'"
    End If
    vlSql = vlSql & " AND L.FEC_PAGO >= '" & iFecDesde & "' AND L.FEC_PAGO <= '" & iFecHasta & "'"
    vlSql = vlSql & " AND P.COD_CONHABDES  = C.COD_CONHABDES"
    'vlSql = vlSql & " ORDER BY P.NUM_POLIZA, P.NUM_ORDEN, P.RUT_RECEPTOR, P.COD_TIPRECEPTOR,"
    vlSql = vlSql & " ORDER BY L.FEC_PAGO, P.NUM_POLIZA, P.NUM_ORDEN, P.NUM_IDENRECEPTOR, P.COD_TIPOIDENRECEPTOR, P.COD_TIPRECEPTOR, P.NUM_PERPAGO," 'HQR 17/03/2006 Se agrega número de periodo
    vlSql = vlSql & " C.COD_IMPONIBLE DESC, C.COD_TRIBUTABLE DESC, C.COD_TIPMOV DESC"
    Set vlTB = vgConexionBD.Execute(vlSql)
    If Not vlTB.EOF Then
        Do While Not vlTB.EOF
            If vlPoliza <> vlTB!num_poliza Or vlOrden <> vlTB!Num_Orden Or vlNumIdenReceptor <> vlTB!Num_IdenReceptor Or vlCodTipoIdenReceptor <> vlTB!Cod_TipoIdenReceptor Or vlTipReceptor <> vlTB!Cod_TipReceptor Or (vlPerPago <> vlTB!Num_PerPago And vlPago = "R") Or (vlFecPago <> vlTB!Fec_Pago And vlPago = "P") Then 'hqr 17/03/2006 Se agrega Número de Periodo
                'Reinicia el Contador
                vlItem = 1
                vlPoliza = vlTB!num_poliza
                vlOrden = vlTB!Num_Orden
                vlMontoPesosPrimerPago = 0
                vlNumConceptosHab = 0
                vlNumConceptosDesc = 0
                vlNumIdenReceptor = vlTB!Num_IdenReceptor
                vlCodTipoIdenReceptor = vlTB!Cod_TipoIdenReceptor
                vlTipReceptor = vlTB!Cod_TipReceptor
                stTTMPLiquidacion.num_poliza = vlPoliza
                stTTMPLiquidacion.Num_IdenReceptor = vlTB!Num_IdenReceptor
                stTTMPLiquidacion.Cod_TipoIdenReceptor = vlTB!Cod_TipoIdenReceptor
                stTTMPLiquidacion.Num_Orden = vlOrden
                stTTMPLiquidacion.num_endoso = IIf(IsNull(vlTB!num_endoso), 0, vlTB!num_endoso)
                stTTMPLiquidacion.Cod_TipReceptor = vlTB!Cod_TipReceptor
                If stTTMPLiquidacion.Cod_TipReceptor <> "R" Then
                    stTTMPLiquidacion.Gls_Direccion = IIf(IsNull(vlTB!Gls_Direccion), "", vlTB!Gls_Direccion)
                Else
                    stTTMPLiquidacion.Gls_Direccion = ""
                End If
                stTTMPLiquidacion.Gls_NomReceptor = vlTB!Gls_NomReceptor & " " & IIf(IsNull(vlTB!Gls_NomSegReceptor), "", vlTB!Gls_NomSegReceptor & " ") & vlTB!Gls_PatReceptor & IIf(IsNull(vlTB!Gls_MatReceptor), "", " " + vlTB!Gls_MatReceptor)
                If stTTMPLiquidacion.Cod_TipReceptor <> "R" Then
                    stTTMPLiquidacion.Cod_Direccion = vlTB!Cod_Direccion
                Else
                    stTTMPLiquidacion.Cod_Direccion = "0"
                End If
                stTTMPLiquidacion.Num_IdenBen = vlTB!Num_IdenBen
                stTTMPLiquidacion.Cod_TipoIdenBen = vlTB!Cod_TipoIdenBen
                stTTMPLiquidacion.Gls_NomBen = vlTB!Gls_NomBen & " " & IIf(IsNull(vlTB!Gls_NomSegBen), "", vlTB!Gls_NomSegBen & " ") & vlTB!Gls_PatBen & IIf(IsNull(vlTB!Gls_MatBen), "", " " & vlTB!Gls_MatBen)
                'para primeros pagos
                stTTMPLiquidacion.Mto_LiqPagar = 0
                stTTMPLiquidacion.Num_Cargas = 0 'vlTB!Num_Cargas
                stTTMPLiquidacion.Mto_LiqHaber = 0
                stTTMPLiquidacion.Mto_LiqDescuento = 0
                'fin primeros pagos
                'Obtiene Fecha de Término del Poder Notarial
                If stTTMPLiquidacion.Cod_TipReceptor <> "R" Then
                    vlSql = "SELECT tut.fec_terpodnot FROM pp_tmae_tutor tut"
                    vlSql = vlSql & " WHERE tut.num_poliza = '" & stTTMPLiquidacion.num_poliza & "'"
                    vlSql = vlSql & " AND tut.num_orden = " & stTTMPLiquidacion.Num_Orden
                    vlSql = vlSql & " AND tut.num_identut = '" & vlTB!Num_IdenReceptor & "'"
                    vlSql = vlSql & " AND tut.cod_tipoidentut = " & Str(Mid(Cmb_TipoIdent.Text, 1, InStr(1, Cmb_TipoIdent.Text, "-") - 1))
                    Set vlTB2 = vgConexionBD.Execute(vlSql)
                    If Not vlTB2.EOF Then
                        stTTMPLiquidacion.Fec_TerPodNot = vlTB2!Fec_TerPodNot
                        stTTMPLiquidacion.Fec_TerPodNot = DateSerial(Mid(stTTMPLiquidacion.Fec_TerPodNot, 1, 4), Mid(stTTMPLiquidacion.Fec_TerPodNot, 5, 2), Mid(stTTMPLiquidacion.Fec_TerPodNot, 7, 2))
                    Else
                        stTTMPLiquidacion.Fec_TerPodNot = ""
                    End If
                    
                    'Obtiene Mensajes al Beneficiario
                    stTTMPLiquidacion.Gls_Mensaje = ""
                    vlSql = "SELECT par.gls_mensaje FROM pp_tmae_menpoliza men, pp_tpar_mensaje par"
                    vlSql = vlSql & " WHERE par.cod_mensaje = men.cod_mensaje"
                    vlSql = vlSql & " AND men.num_poliza = '" & stTTMPLiquidacion.num_poliza & "'"
                    vlSql = vlSql & " AND men.num_orden = " & stTTMPLiquidacion.Num_Orden
                    vlSql = vlSql & " AND men.num_perpago = '" & stTTMPLiquidacion.Num_PerPago & "'"
                    Set vlTB2 = vgConexionBD.Execute(vlSql)
                    If Not vlTB2.EOF Then
                        Do While Not vlTB2.EOF
                            stTTMPLiquidacion.Gls_Mensaje = stTTMPLiquidacion.Gls_Mensaje & vlTB2!Gls_Mensaje & Chr(13)
                            vlTB2.MoveNext
                        Loop
                    Else
                        stTTMPLiquidacion.Gls_Mensaje = ""
                    End If
                Else
                    stTTMPLiquidacion.Fec_TerPodNot = ""
                    stTTMPLiquidacion.Gls_Mensaje = ""
                End If
            End If
            
''            'CMV-20060803 I
''            vlPerPagoNum = vlTB!Num_PerPago
''            'CMV-20060803 F
            stTTMPLiquidacion.Num_PerPago = vlTB!Num_PerPago
            stTTMPLiquidacion.Mto_Pension = vlTB!Mto_Pension
            stTTMPLiquidacion.Num_Cargas = 0 'vlTB!Num_Cargas
            If vlPerPago <> stTTMPLiquidacion.Num_PerPago Then
                stTTMPLiquidacion.Fec_Pago = vlTB!Fec_Pago
                'hqr 25/08/2007 Se deja comentado porque no se muestra en la Liquidación
                'Obtiene Fecha del Próximo
'                vlSql = "SELECT pro.fec_pagoproxreg"
'                vlSql = vlSql & " FROM pp_tmae_propagopen pro"
'                vlSql = vlSql & " WHERE pro.num_perpago = '" & stTTMPLiquidacion.Num_PerPago & "'"
'                Set vlTB2 = vgConexionBD.Execute(vlSql)
'                If Not vlTB2.EOF Then
'                    stTTMPLiquidacion.Fec_PagoProxReg = vlTB2!Fec_PagoProxReg
'                    stTTMPLiquidacion.Fec_PagoProxReg = DateSerial(Mid(stTTMPLiquidacion.Fec_PagoProxReg, 1, 4), Mid(stTTMPLiquidacion.Fec_PagoProxReg, 5, 2), Mid(stTTMPLiquidacion.Fec_PagoProxReg, 7, 2))
'                Else
                    stTTMPLiquidacion.Fec_PagoProxReg = ""
'                End If
                'Obtiene Valor US (hqr 25/08/2007 se deja comentado porque no se u)
'                vlSql = "SELECT val.mto_moneda"
'                vlSql = vlSql & " FROM ma_tval_moneda val"
'                vlSql = vlSql & " WHERE val.cod_moneda = 'UF'"
'                vlSql = vlSql & " AND val.fec_moneda = '" & stTTMPLiquidacion.Fec_Pago & "'"
'                Set vlTB2 = vgConexionBD.Execute(vlSql)
'                If Not vlTB2.EOF Then
'                    stTTMPLiquidacion.Mto_Moneda = vlTB2!Mto_Moneda
'                Else
                    stTTMPLiquidacion.Mto_Moneda = 0
'                End If
                stTTMPLiquidacion.Fec_Pago = DateSerial(Mid(stTTMPLiquidacion.Fec_Pago, 1, 4), Mid(stTTMPLiquidacion.Fec_Pago, 5, 2), Mid(stTTMPLiquidacion.Fec_Pago, 7, 2))
                vlPerPago = stTTMPLiquidacion.Num_PerPago
            End If
            'Obtiene Tipo de Pensión
            If vlTipPension <> vlTB!Cod_TipPension Then
                vlSql = "SELECT tab.gls_elemento FROM ma_tpar_tabcod tab"
                vlSql = vlSql & " WHERE tab.cod_tabla = 'TP'"
                vlSql = vlSql & " AND tab.cod_elemento = '" & vlTB!Cod_TipPension & "'"
                Set vlTB2 = vgConexionBD.Execute(vlSql)
                If Not vlTB2.EOF Then
                    stTTMPLiquidacion.Gls_TipPension = vlTB2!GLS_ELEMENTO
                Else
                    stTTMPLiquidacion.Gls_TipPension = ""
                End If
                vlTipPension = vlTB!Cod_TipPension
            End If
            'Obtiene Via de Pago
            If vlViaPago <> vlTB!Cod_ViaPago Then
                vlSql = "SELECT tab.gls_elemento FROM ma_tpar_tabcod tab"
                vlSql = vlSql & " WHERE tab.cod_tabla = 'VPG'"
                vlSql = vlSql & " AND tab.cod_elemento = '" & vlTB!Cod_ViaPago & "'"
                Set vlTB2 = vgConexionBD.Execute(vlSql)
                If Not vlTB2.EOF Then
                    vlDescViaPago = vlTB2!GLS_ELEMENTO
                Else
                    vlDescViaPago = ""
                End If
                vlViaPago = vlTB!Cod_ViaPago
            End If
            
            'hqr 13/10/2007 Obtiene Sucursal de la Via de Pago
            If vlTB!Cod_ViaPago = "04" Then
                If vlTB!Cod_Sucursal <> vlSucursal Then
                    'Obtiene Sucursal
                    stTTMPLiquidacion.Gls_ViaPago = vlDescViaPago
                    vlSql = "SELECT a.gls_sucursal FROM ma_tpar_sucursal a"
                    vlSql = vlSql & " WHERE a.cod_sucursal = '" & vlTB!Cod_Sucursal & "'"
                    vlSql = vlSql & " AND a.cod_tipo = 'A'" 'AFP
                    Set vlTB2 = vgConexionBD.Execute(vlSql)
                    If Not vlTB2.EOF Then
                        vlDescSucursal = vlTB2!gls_sucursal
                    End If
                    vlSucursal = vlTB!Cod_Sucursal
                End If
                stTTMPLiquidacion.Gls_ViaPago = Mid(vlDescViaPago & " - " & vlDescSucursal, 1, 50)
            Else
                stTTMPLiquidacion.Gls_ViaPago = vlDescViaPago
            End If
            
            stTTMPLiquidacion.Gls_CajaComp = ""
            
            'Obtiene Institución de Salud
            If Not IsNull(vlTB!Cod_InsSalud) Then
                If vlTB!Cod_InsSalud <> "NULL" Then
                    If vlInsSalud <> vlTB!Cod_InsSalud Then
                        vlSql = "SELECT tab.gls_elemento FROM ma_tpar_tabcod tab"
                        vlSql = vlSql & " WHERE tab.cod_tabla = 'IS'"
                        vlSql = vlSql & " AND tab.cod_elemento = '" & vlTB!Cod_InsSalud & "'"
                        Set vlTB2 = vgConexionBD.Execute(vlSql)
                        If Not vlTB2.EOF Then
                            stTTMPLiquidacion.Gls_InsSalud = vlTB2!GLS_ELEMENTO
                        Else
                            stTTMPLiquidacion.Gls_InsSalud = ""
                        End If
                        vlInsSalud = vlTB!Cod_InsSalud
                    End If
                Else
                    vlInsSalud = ""
                    stTTMPLiquidacion.Gls_InsSalud = ""
                End If
            Else
                vlInsSalud = ""
                stTTMPLiquidacion.Gls_InsSalud = ""
            End If
            
            'Obtiene AFP
            If vlTB!Cod_ViaPago <> "04" Then 'hqr 13/10/2007 Para que se muestre el nombre del receptor cuando no es traspaso AFP
                stTTMPLiquidacion.Gls_Afp = stTTMPLiquidacion.Gls_NomReceptor 'Receptor
            Else
                If vlAfp <> vlTB!cod_afp Then
                    vlSql = "SELECT tab.gls_elemento FROM ma_tpar_tabcod tab"
                    vlSql = vlSql & " WHERE tab.cod_tabla = 'AF'"
                    vlSql = vlSql & " AND tab.cod_elemento = '" & vlTB!cod_afp & "'"
                    Set vlTB2 = vgConexionBD.Execute(vlSql)
                    If Not vlTB2.EOF Then
                        stTTMPLiquidacion.Gls_Afp = "AFP " & vlTB2!GLS_ELEMENTO
                    Else
                        stTTMPLiquidacion.Gls_Afp = "AFP"
                    End If
                    vlAfp = vlTB!cod_afp
                End If
            End If
            
            'Obtiene Direccion
            If stTTMPLiquidacion.Cod_Direccion <> vlCodDireccion Then
                If stTTMPLiquidacion.Cod_Direccion <> 0 Then
                    vlSql = "SELECT com.gls_comuna, prov.gls_provincia, reg.gls_region"
                    vlSql = vlSql & " FROM ma_tpar_comuna com, ma_tpar_provincia prov, ma_tpar_region reg"
                    vlSql = vlSql & " WHERE reg.cod_region = prov.cod_region"
                    vlSql = vlSql & " AND prov.cod_region = com.cod_region"
                    vlSql = vlSql & " AND prov.cod_provincia = com.cod_provincia"
                    vlSql = vlSql & " AND com.cod_direccion = '" & stTTMPLiquidacion.Cod_Direccion & "'"
                    Set vlTB2 = vgConexionBD.Execute(vlSql)
                    If Not vlTB2.EOF Then
                        stTTMPLiquidacion.Gls_Direccion2 = vlTB2!gls_region & " - " & vlTB2!gls_provincia & " - " & vlTB2!gls_comuna
                    Else
                        stTTMPLiquidacion.Gls_Direccion2 = ""
                    End If
                Else
                    stTTMPLiquidacion.Gls_Direccion2 = ""
                End If
                vlCodDireccion = stTTMPLiquidacion.Cod_Direccion
            End If
            
            'Obtiene Datos
            stTTMPLiquidacion.Cod_Moneda = vlTB!COD_SCOMP
            If vlPago = "R" Then
                stTTMPLiquidacion.Mto_LiqPagar = vlTB!Mto_LiqPagar
                stTTMPLiquidacion.Mto_LiqHaber = vlTB!Mto_Haber
                stTTMPLiquidacion.Mto_LiqDescuento = vlTB!Mto_Descuento
                stTTMPLiquidacion.Gls_MontoPension = fgConvierteNumeroLetras(stTTMPLiquidacion.Mto_LiqPagar, vlTB!Moneda)
            End If
            If vlTB!cod_tipmov = "H" Then 'haber
                If vlPago <> "R" Then
                    stTTMPLiquidacion.Mto_LiqHaber = stTTMPLiquidacion.Mto_LiqHaber + vlTB!Mto_ConHabDes
                    stTTMPLiquidacion.Mto_LiqPagar = stTTMPLiquidacion.Mto_LiqPagar + vlTB!Mto_ConHabDes
                    stTTMPLiquidacion.Gls_MontoPension = fgConvierteNumeroLetras(stTTMPLiquidacion.Mto_LiqPagar, vlTB!Moneda)
                End If
                stTTMPLiquidacion.Cod_ConDescto = "NULL"
                stTTMPLiquidacion.Mto_Descuento = 0
                stTTMPLiquidacion.Cod_ConHaber = vlTB!gls_ConHabDes
                stTTMPLiquidacion.Mto_Haber = vlTB!Mto_ConHabDes
                vlNumConceptosHab = vlNumConceptosHab + 1
                stTTMPLiquidacion.Num_Item = vlNumConceptosHab
                If vlNumConceptosHab > vlNumConceptosDesc Then
                    Call fgInsertaTTMPLiquidacion(stTTMPLiquidacion)
                Else
                    Call fgActualizaTTMPLiquidacionHab(stTTMPLiquidacion)
                End If
            ElseIf vlTB!cod_tipmov = "D" Then 'descuento
                If vlPago <> "R" Then
                    stTTMPLiquidacion.Mto_LiqDescuento = stTTMPLiquidacion.Mto_LiqDescuento + vlTB!Mto_ConHabDes
                    stTTMPLiquidacion.Mto_LiqPagar = stTTMPLiquidacion.Mto_LiqPagar - vlTB!Mto_ConHabDes
                    stTTMPLiquidacion.Gls_MontoPension = fgConvierteNumeroLetras(stTTMPLiquidacion.Mto_LiqPagar, vlTB!Moneda)
                End If
                stTTMPLiquidacion.Cod_ConDescto = vlTB!gls_ConHabDes
                stTTMPLiquidacion.Mto_Descuento = vlTB!Mto_ConHabDes
                stTTMPLiquidacion.Cod_ConHaber = "NULL"
                stTTMPLiquidacion.Mto_Haber = 0
                vlNumConceptosDesc = vlNumConceptosDesc + 1
                stTTMPLiquidacion.Num_Item = vlNumConceptosDesc
                If vlNumConceptosDesc > vlNumConceptosHab Then
                    Call fgInsertaTTMPLiquidacion(stTTMPLiquidacion)
                Else
                    Call fgActualizaTTMPLiquidacionDesc(stTTMPLiquidacion)
                End If
            Else 'OTROS
                stTTMPLiquidacion.Cod_ConDescto = vlTB!Cod_ConHabDes
                stTTMPLiquidacion.Mto_Haber = vlTB!Mto_ConHabDes
                stTTMPLiquidacion.Num_Item = 0
                'Call fgInsertaTTMPLiquidacion(stTTMPLiquidacion)
            End If
            vlFecPago = vlTB!Fec_Pago
            vlItem = vlItem + 1
            vlTB.MoveNext
        Loop
    Else
        MsgBox "No existe Información para este rango de Fechas", vbInformation, "Operacion Cancelada"
        Exit Function
    End If
    flLlenaTemporal = True
End Function
Function flCargaCombo()
Cmb_Tipo.Clear
Cmb_Tipo.AddItem ("DEFINITIVO")
Cmb_Tipo.AddItem ("PROVISORIO")
Cmb_Tipo.ListIndex = 0
End Function

Function flInformeHabDes()
Dim vlSql As String
On Error GoTo Err_Habdes
    
    Screen.MousePointer = 11
    
    vlArchivo = ""
    
    Select Case vlOpcion
        Case "D": 'Definitivo
            vlArchivo = strRpt & "PP_Rpt_CieLibroHaberDescDef.rpt"
        Case "P": 'Provisorio
            vlArchivo = strRpt & "PP_Rpt_CieLibroHaberDescPro.rpt"
    End Select
    
    If Not fgExiste(vlArchivo) Then     ', vbNormal
        MsgBox "Archivo de Reporte no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
        Screen.MousePointer = 0
        Exit Function
    End If
        
    If Chk_Pensionado.Value = 1 Then
        
        If (Txt_Poliza <> "") Then
            vlNumPol = Trim(Txt_Poliza)
        Else
            vlNumPol = ""
        End If
        If (Txt_NumIdent <> "") Then
            vlNumIdent = Trim(Txt_NumIdent)
            vlCodTipoIdent = Str(Mid(Cmb_TipoIdent.Text, 1, InStr(1, Cmb_TipoIdent.Text, "-") - 1))
        Else
            vlNumIdent = ""
            vlCodTipoIdent = ""
        End If
        
        vlSql = ""
        vlSql = "select L.NUM_POLIZA, L.NUM_ENDOSO, L.NUM_ORDEN from "
        vlSql = vlSql & "PP_TMAE_LIQPAGOPEN" & vlGlosaOpcion & " L,"
        vlSql = vlSql & "PP_TMAE_BEN B, "
        vlSql = vlSql & "PP_TMAE_PAGOPEN" & vlGlosaOpcion & " P "
        vlSql = vlSql & "where "
        If (vlNumPol <> "") Then
            vlSql = vlSql & "L.NUM_POLIZA = '" & vlNumPol & "' and "
        End If
        If (vlNumIdent <> "") Then
            vlSql = vlSql & "B.num_idenben = '" & Trim(vlNumIdent) & "' and "
            vlSql = vlSql & "B.cod_tipoidenben = " & Str(vlCodTipoIdent) & " and "
        End If
        vlSql = vlSql & "L.FEC_PAGO >= '" & vlFechaDesde & "' AND "
        vlSql = vlSql & "L.FEC_PAGO <= '" & vlFechaHasta & "' AND "
        vlSql = vlSql & "L.COD_TIPOPAGO = '" & vlPago & "' AND "
        'I---- ABV 03/09/2004 ---
        vlSql = vlSql & "L.NUM_POLIZA = B.NUM_POLIZA AND "
        vlSql = vlSql & "L.NUM_ENDOSO = B.NUM_ENDOSO AND "
        vlSql = vlSql & "L.NUM_ORDEN = B.NUM_ORDEN AND "
        'vlSQL = vlSQL & "B.num_poliza=L.num_poliza and "
        'vlSQL = vlSQL & "P.num_poliza=L.num_poliza "
        vlSql = vlSql & "L.num_perpago = P.num_perpago and "
        vlSql = vlSql & "L.NUM_POLIZA = P.NUM_POLIZA and "
        vlSql = vlSql & "L.NUM_ORDEN = P.NUM_ORDEN AND "
        vlSql = vlSql & "L.cod_tipoidenRECEPTOR = P.cod_tipoidenRECEPTOR AND "
        vlSql = vlSql & "L.num_idenRECEPTOR = P.num_idenRECEPTOR AND "
        vlSql = vlSql & "L.COD_TIPRECEPTOR = P.COD_TIPRECEPTOR "
        'F---- ABV 03/09/2004 ---
    Else
        vlSql = "select NUM_POLIZA, NUM_ENDOSO, NUM_ORDEN "
        vlSql = vlSql & "from "
        vlSql = vlSql & "PP_TMAE_LIQPAGOPEN" & vlGlosaOpcion & " "
        vlSql = vlSql & "WHERE "
        vlSql = vlSql & "FEC_PAGO >= '" & vlFechaDesde & "' AND "
        vlSql = vlSql & "FEC_PAGO <= '" & vlFechaHasta & "' AND "
        vlSql = vlSql & "COD_TIPOPAGO = '" & vlPago & "' "
    End If
    Set vgRs = vgConexionBD.Execute(vlSql)
    If vgRs.EOF Then
        vgRs.Close
        MsgBox "No existe Información para este Rango de Fechas", vbInformation, "Operacion Cancelada"
        Exit Function
    End If
    vgRs.Close
    
    Call fgVigenciaQuiebra(Txt_Desde)
        
    vgQuery = ""
    If Chk_Pensionado.Value = 1 Then
        If (vlNumPol <> "") Then
            vgQuery = vgQuery & "{PP_TMAE_LIQPAGOPEN" & vlGlosaOpcion & ".NUM_POLIZA} = '" & Trim(vlNumPol) & "' AND "
        End If
        If (vlNumIdent <> "") Then
            vgQuery = vgQuery & "{PP_TMAE_BEN.NUM_IDENBEN} = '" & Trim(vlNumIdent) & "' AND "
            vgQuery = vgQuery & "{PP_TMAE_BEN.COD_TIPOIDENBEN} = " & Str(vlCodTipoIdent) & " AND "
        End If
    End If
    vgQuery = vgQuery & " {PP_TMAE_LIQPAGOPEN" & vlGlosaOpcion & ".FEC_PAGO} >= '" & vlFechaDesde & "' AND "
    vgQuery = vgQuery & " {PP_TMAE_LIQPAGOPEN" & vlGlosaOpcion & ".FEC_PAGO} <= '" & vlFechaHasta & "' AND "
    vgQuery = vgQuery & " {PP_TMAE_LIQPAGOPEN" & vlGlosaOpcion & ".COD_TIPOPAGO}= '" & vlPago & "' AND "
    vgQuery = vgQuery & " {MA_TPAR_TABCOD.cod_tabla} = '" & Trim(vgCodTabla_TipMon) & "' "
    
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
    Rpt_Reporte.Formulas(3) = "FecDesde = '" & vlFechaDesde & "'"
    Rpt_Reporte.Formulas(4) = "FecHasta = '" & vlFechaHasta & "'"
    Rpt_Reporte.Formulas(5) = "CodPago = '" & vlPago & "'"
    
   If Trim(vlGlosaOpcion) = "DEF" Then
       Rpt_Reporte.Formulas(6) = "TipoProceso= 'DEFINITIVO' "
   Else
       Rpt_Reporte.Formulas(6) = "TipoProceso= 'PROVISORIO' "
   End If
   Rpt_Reporte.Formulas(7) = "GlosaQuiebra = '" & vgGlosaQuiebra & "'"
    
    Rpt_Reporte.Destination = crptToWindow
    Rpt_Reporte.WindowTitle = "Libro de Haberes y Descuentos"
    Rpt_Reporte.Action = 1
    Screen.MousePointer = 0
    
Exit Function
Err_Habdes:
    Screen.MousePointer = 0
    MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
End Function

Private Sub Chk_Pensionado_Click()
If Chk_Pensionado.Value = 1 Then
    Txt_Poliza.Enabled = True
    Cmb_TipoIdent.Enabled = True
    Txt_NumIdent.Enabled = True
    Txt_Poliza.SetFocus
Else
    Txt_Poliza.Enabled = False
    Cmb_TipoIdent.Enabled = False
    Txt_NumIdent.Enabled = False
    Cmd_Imprimir.SetFocus
End If
End Sub

Private Sub Chk_Pensionado_KeyPress(KeyAscii As Integer)
If (Chk_Pensionado.Value = 1) Then
    Txt_Poliza.SetFocus
Else
    Cmd_Imprimir.SetFocus
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


Private Sub Cmb_TipoIdent_Click()
If (Cmb_TipoIdent <> "") Then
    vlPosicionTipoIden = Cmb_TipoIdent.ListIndex
    vlLargoTipoIden = Cmb_TipoIdent.ItemData(vlPosicionTipoIden)
    If (vlLargoTipoIden = 0) Then
        Txt_NumIdent.Text = "0"
        Txt_NumIdent.Enabled = False
    Else
        Txt_NumIdent.MaxLength = vlLargoTipoIden
        Txt_NumIdent.Enabled = True
        Txt_NumIdent.Text = Mid(Txt_NumIdent, 1, vlLargoTipoIden) 'HQR 09/06/2007
    End If
End If
End Sub

Private Sub Cmb_TipoIdent_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Txt_NumIdent.Enabled Then
            Txt_NumIdent.SetFocus
        Else
            Cmd_Imprimir.SetFocus
        End If
    End If
End Sub


Private Sub Cmd_Imprimir_Click()
Dim vlFechaInicio As String
On Error GoTo errImprimir

'    Txt_Hasta.Text = Format(CDate(Trim(Txt_Hasta)), "yyyymmdd")
'    vlFechaHasta = Trim(Txt_Hasta)
'    Txt_Hasta.Text = DateSerial(Mid((Txt_Hasta.Text), 1, 4), Mid((Txt_Hasta.Text), 5, 2), Mid((Txt_Hasta.Text), 7, 2))
'
'    Txt_Desde.Text = Format(CDate(Trim(Txt_Desde)), "yyyymmdd")
'    vlFechaDesde = Trim(Txt_Desde)
'    Txt_Desde.Text = DateSerial(Mid((Txt_Desde.Text), 1, 4), Mid((Txt_Desde.Text), 5, 2), Mid((Txt_Desde.Text), 7, 2))
    
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
        MsgBox "La Fecha de Término de Perido es mayor a la fecha de Inicio", vbCritical, "Error de Datos"
        Exit Sub
    End If
    
    Txt_NumIdent = Trim(Txt_NumIdent)
        
    If Chk_Pensionado.Value = 1 Then
        If Txt_Poliza = "" Then
            MsgBox "Falta ingresar Número de Póliza", vbCritical, "Falta Información"
            Txt_Poliza.SetFocus
            Exit Sub
        End If
        If (Txt_NumIdent <> "" And Txt_NumIdent <> "0") Then
            If Cmb_TipoIdent.ListIndex < 0 Then
                MsgBox "Debe Seleccionar Tipo de Documento del Pensionado", vbCritical, "Falta Información"
                Cmb_TipoIdent.SetFocus
                Exit Sub
            End If
        Else
            If Cmb_TipoIdent.ListIndex >= 0 Then
                If Str(Mid(Cmb_TipoIdent.Text, 1, InStr(1, Cmb_TipoIdent.Text, "-") - 1)) <> 0 Then
                    MsgBox "Debe Ingresar Número de Identificación del Pensionado", vbCritical, "Falta Información"
                    Txt_NumIdent.SetFocus
                    Exit Sub
                End If
            End If
        End If
        'Valida que el rut del beneficiario , se encuentre registrado
        If Txt_NumIdent <> "" And Txt_NumIdent <> "0" And Txt_Poliza = "" Then
            vlNumPoliza = ""
            vlNumOrden = 0
            vlNumEndoso = 0
            vgSql = ""
            vgSql = " SELECT b.num_poliza,b.num_orden,b.num_endoso "
            vgSql = vgSql & " FROM pp_tmae_ben b WHERE "
            vgSql = vgSql & " b.num_idenben = '" & Trim(Txt_NumIdent) & "' "
            vgSql = vgSql & " b.cod_tipoidenben = " & Str(Mid(Cmb_TipoIdent.Text, 1, InStr(1, Cmb_TipoIdent.Text, "-") - 1))
            Set vgRs = vgConexionBD.Execute(vgSql)
            If Not (vgRs.EOF) Then
                If Not IsNull(vgRs!num_poliza) Then vlNumPoliza = (vgRs!num_poliza)
                If Not IsNull(vgRs!Num_Orden) Then vlNumOrden = (vgRs!Num_Orden)
                If Not IsNull(vgRs!num_endoso) Then vlNumEndoso = (vgRs!num_endoso)
            Else
                MsgBox "El Beneficiario Ingresado No se Encuentra Registrado", vbCritical, "Error de Datos"
                Exit Sub
            End If
        End If
        'Valida que el rut del beneficiario y poliza, se encuentren registrados
        If Txt_NumIdent <> "" And Txt_NumIdent <> "0" And Txt_Poliza <> "" Then
            vlNumPoliza = ""
            vlNumOrden = 0
            vlNumEndoso = 0
            vgSql = ""
            vgSql = " SELECT b.num_poliza,b.num_orden,b.num_endoso "
            vgSql = vgSql & " FROM pp_tmae_ben b WHERE "
            vgSql = vgSql & " b.num_poliza = '" & Trim(Txt_Poliza) & "' AND "
            vgSql = vgSql & " b.num_idenben = '" & Trim(Txt_NumIdent) & "' AND "
            vgSql = vgSql & " b.cod_tipoidenben = " & Str(Mid(Cmb_TipoIdent.Text, 1, InStr(1, Cmb_TipoIdent.Text, "-") - 1)) & " AND "
            vgSql = vgSql & " b.num_endoso = "
            vgSql = vgSql & " (SELECT MAX(p.num_endoso) FROM pp_tmae_poliza p WHERE "
            vgSql = vgSql & " p.num_poliza = b.num_poliza) "
            Set vgRs = vgConexionBD.Execute(vgSql)
            If Not (vgRs.EOF) Then
                If Not IsNull(vgRs!num_poliza) Then vlNumPoliza = (vgRs!num_poliza)
                If Not IsNull(vgRs!Num_Orden) Then vlNumOrden = (vgRs!Num_Orden)
                If Not IsNull(vgRs!num_endoso) Then vlNumEndoso = (vgRs!num_endoso)
            Else
                MsgBox "El Beneficiario Ingresado No se Encuentra Registrado", vbCritical, "Error de Datos"
                Exit Sub
            End If
        End If
'''        If Txt_Rut <> "" Then
'''            vlNumPoliza = ""
'''            vlNumOrden = ""
'''            vlNumEndoso = ""
'''            vgSql = ""
'''            vgSql = " SELECT Distinct (b.num_poliza),b.num_orden,b.num_endoso "
'''            vgSql = vgSql & " FROM pp_tmae_ben b WHERE "
'''            vgSql = vgSql & " b.num_poliza = '" & Trim(Txt_Poliza) & "' AND "
'''            vgSql = vgSql & " b.rut_ben = " & Str(Txt_Rut) & " AND "
'''            vgSql = vgSql & " b.num_endoso = "
'''            vgSql = vgSql & " (SELECT MAX(p.num_endoso) FROM pp_tmae_poliza p WHERE "
'''            vgSql = vgSql & " p.num_poliza = b.num_poliza) "
'''            vgSql = vgSql & " ORDER BY b.num_poliza,b.num_endoso,b.num_orden "
'''            Set vgRs = vgConexionBD.Execute(vgSql)
'''            If Not (vgRs.EOF) Then
'''                If Not IsNull(vgRs!Num_Poliza) Then vlNumPoliza = (vgRs!Num_Poliza)
'''                If Not IsNull(vgRs!Num_Orden) Then vlNumOrden = (vgRs!Num_Orden)
'''                If Not IsNull(vgRs!num_endoso) Then vlNumEndoso = (vgRs!num_endoso)
'''            Else
'''                MsgBox "El Beneficiario Ingresado No se Encuentra Registrado", vbCritical, "Error de Datos"
'''                Exit Sub
'''            End If
'''        End If
        'Valida que la poliza ingresada se encuentre registrada
        If Txt_Poliza <> "" And (Txt_NumIdent = "" Or Txt_NumIdent = "0") Then
            vlNumPoliza = ""
            vlNumEndoso = 0
            vgSql = ""
            vgSql = " SELECT p.num_poliza,p.num_endoso "
            vgSql = vgSql & " FROM pp_tmae_poliza p WHERE "
            vgSql = vgSql & " p.num_poliza = '" & Trim(Txt_Poliza) & "' AND "
            vgSql = vgSql & " p.num_endoso = "
            vgSql = vgSql & " (SELECT MAX(p.num_endoso) FROM pp_tmae_poliza z WHERE "
            vgSql = vgSql & " z.num_poliza = p.num_poliza) "
            Set vgRs = vgConexionBD.Execute(vgSql)
            If Not (vgRs.EOF) Then
                If Not IsNull(vgRs!num_poliza) Then vlNumPoliza = (vgRs!num_poliza)
                If Not IsNull(vgRs!num_endoso) Then vlNumEndoso = (vgRs!num_endoso)
            Else
                MsgBox "La Póliza Ingresada No se Encuentra Registrada", vbCritical, "Error de Datos"
                Exit Sub
            End If
        End If
        
    End If
    Screen.MousePointer = 11

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
    If vlPago = "R" Then 'En regimen
        If flValidaEstadoProceso(Mid(Trim(vlFechaInicio), 1, 6), vlCodEstado) = False Then
            MsgBox "El Tipo de Proceso Seleccionado no se encuentra Realizado.", vbCritical, "Error de Datos"
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If

    'Permite imprimir la Opción Indicada a través del Menú
    Select Case vgNombreInformeSeleccionadoInd
        Case "InfLiqPago" 'Liquidaciones de Pensiones
            flInformeLiqPago
        Case "InfLibHabDes" 'Informe de Libro de Haberes y Desctos.
            flInformeHabDes
        Case "InfCONLiqPago" 'Liquidación de Pensiones desde Frm_Consulta
            flInformeLiqPago
        Case "InfLiqPagoCarta" 'Liquidaciones de Pensiones de Tipo Carta 20060207
            flInformeLiqPagoCarta
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
    Txt_Poliza = ""
    Txt_NumIdent = ""
    Cmb_TipoIdent.ListIndex = 0
    Chk_Pensionado.Value = 0
    Txt_Poliza.Enabled = False
    Txt_NumIdent.Enabled = False
    Cmb_TipoIdent.Enabled = False
    Txt_Desde.SetFocus

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
On Error GoTo Err_Cargar

    Frm_PlanillaPensionado.Left = 0
    Frm_PlanillaPensionado.Top = 0
    
    fgComboTipoCalculo Cmb_Tipo
    fgComboTipoPension Cmb_Pago
    fgComboTipoIdentificacion Cmb_TipoIdent
Exit Sub
Err_Cargar:
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
Chk_Pensionado.SetFocus
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


Private Sub Txt_NumIdent_GotFocus()
    Txt_NumIdent.SelStart = 0
    Txt_NumIdent.SelLength = Len(Txt_NumIdent)
End Sub

Private Sub Txt_NumIdent_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Cmd_Imprimir.SetFocus
    End If
End Sub

Private Sub Txt_Poliza_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Cmb_TipoIdent.SetFocus
End If
End Sub

Private Sub Txt_Poliza_LostFocus()
If Txt_Poliza <> "" Then
    Txt_Poliza = Format(Trim(Txt_Poliza), "0000000000")
End If
End Sub

Function flValidaEstadoProceso(iPeriodo As String, iCodEstado As String) As Boolean
On Error GoTo Err_flValidaTipoProceso

    flValidaEstadoProceso = False

    vgSql = ""
    vgSql = "SELECT p.cod_estadoreg "
    vgSql = vgSql & "FROM pp_tmae_propagopen p "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "p.num_perpago = '" & Trim(iPeriodo) & "' "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
        If Trim(vgRegistro!cod_estadoreg) = Trim(iCodEstado) Then
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

Function flPerPago(iMesPerPago As String) As String
On Error GoTo Err_flPerPago
    
    flPerPago = ""
    
    Select Case iMesPerPago
        Case "01" 'ENERO
            flPerPago = clMesEnero
        Case "02" 'FEBRERO
            flPerPago = clMesFebrero
        Case "03" 'MARZO
            flPerPago = clMesMarzo
        Case "04" 'ABRIL
            flPerPago = clMesAbril
        Case "05" 'MAYO
            flPerPago = clMesMayo
        Case "06" 'JUNIO
            flPerPago = clMesJunio
        Case "07" 'JULIO
            flPerPago = clMesJulio
        Case "08" 'AGOSTO
            flPerPago = clMesAgosto
        Case "09" 'SEPTIEMBRE
            flPerPago = clMesSeptiembre
        Case "10" 'OCTUBRE
            flPerPago = clMesOctubre
        Case "11" 'NOVIEMBRE
            flPerPago = clMesNoviembre
        Case "12" 'DICIEMBRE
            flPerPago = clMesDiciembre
    End Select

Exit Function
Err_flPerPago:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function



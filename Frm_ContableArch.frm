VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Frm_ContableArch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generación de Archivo Contable "
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11250
   Icon            =   "Frm_ContableArch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   11250
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   25
      Top             =   4800
      Width           =   11055
      Begin VB.CheckBox chkPagoA 
         Caption         =   "Pago por AFP"
         Height          =   255
         Left            =   6480
         TabIndex        =   27
         Top             =   240
         Width           =   1695
      End
      Begin VB.CheckBox chkPagoD 
         Caption         =   "Pago Directo a Cliente"
         Height          =   255
         Left            =   2760
         TabIndex        =   26
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Fra_Datos 
      Caption         =   "Opciones de Impresión"
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
      Height          =   1005
      Index           =   2
      Left            =   120
      TabIndex        =   18
      Top             =   3705
      Width           =   11055
      Begin VB.OptionButton Opt_DetPendAcumulado 
         Caption         =   "Detalle Pendiente Acumulado (Sólo reporte)"
         Height          =   540
         Left            =   8505
         TabIndex        =   24
         Top             =   330
         Width           =   2475
      End
      Begin VB.OptionButton Opt_DetMtos 
         Caption         =   "Detalle Pensionado"
         Height          =   465
         Left            =   6720
         TabIndex        =   23
         Top             =   375
         Width           =   1695
      End
      Begin VB.OptionButton Opt_DetProdMon 
         Caption         =   "Detalle Prod - Moneda"
         Height          =   480
         Left            =   4680
         TabIndex        =   22
         Top             =   375
         Width           =   1920
      End
      Begin VB.OptionButton Opt_DetAfpMon 
         Caption         =   "Detalle AFP - Moneda"
         Height          =   465
         Left            =   2700
         TabIndex        =   21
         Top             =   375
         Width           =   1935
      End
      Begin VB.OptionButton Opt_Detalle 
         Caption         =   "Detalle Mov."
         Height          =   435
         Left            =   1290
         TabIndex        =   20
         Top             =   375
         Width           =   1335
      End
      Begin VB.OptionButton Opt_Resumen 
         Caption         =   "Resumen"
         Height          =   450
         Left            =   180
         TabIndex        =   19
         Top             =   360
         Value           =   -1  'True
         Width           =   1110
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Proceso de Carga"
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
      Height          =   975
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   3375
      Begin VB.Label Lbl_FecCierre 
         Alignment       =   2  'Center
         BackColor       =   &H00E8FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame Fra_Datos 
      Height          =   1095
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   5400
      Width           =   11070
      Begin VB.CommandButton cmd_salir 
         Caption         =   "&Salir"
         Height          =   755
         Left            =   6900
         Picture         =   "Frm_ContableArch.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton Cmd_Cargar 
         Caption         =   "&Generar"
         Height          =   755
         Left            =   3300
         Picture         =   "Frm_ContableArch.frx":053C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Generación de Archivo Contable"
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   755
         Left            =   4500
         Picture         =   "Frm_ContableArch.frx":0D5E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Imprimir Resumen de Carga"
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmd_limpiar 
         Caption         =   "&Limpiar"
         Height          =   755
         Left            =   5700
         Picture         =   "Frm_ContableArch.frx":1418
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Width           =   840
      End
   End
   Begin VB.Frame Fra_Datos 
      Height          =   735
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   11040
      Begin VB.CommandButton CmdContable 
         Height          =   375
         Left            =   10020
         Picture         =   "Frm_ContableArch.frx":1AD2
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.Label LblArchivo 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1875
         TabIndex        =   9
         Top             =   225
         Width           =   7965
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Destino de Datos       :"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Fra_Datos 
      Caption         =   " Fecha de Pago"
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
      Height          =   1575
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   3375
      Begin VB.TextBox Txt_Hasta 
         Height          =   285
         Left            =   1800
         TabIndex        =   13
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Txt_Desde 
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
      Begin Crystal.CrystalReport Rpt_General 
         Left            =   2880
         Top             =   1100
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label Lbl_NumArchivo 
         Alignment       =   2  'Center
         BackColor       =   &H00E8FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2400
         TabIndex        =   15
         Top             =   1150
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Hasta       :"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   14
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Desde      :"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Msf_Periodos 
      Height          =   2535
      Left            =   3600
      TabIndex        =   12
      Top             =   240
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   4471
      _Version        =   393216
      Rows            =   1
      Cols            =   10
      BackColor       =   14745599
   End
   Begin MSComDlg.CommonDialog ComDialogo 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Frm_ContableArch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vlFecDesde      As String, vlFecHasta As String
Dim vlFecCrea       As String, vlHorCrea As String
Dim vlArchivoCont   As String
Dim vlArchivoContProv   As String
Dim vlArchivo       As String
Dim vlArchPagReg    As Boolean
Dim vlArchPagRegProv    As Boolean
Dim vlArchGtoSep    As Boolean
Dim vlArchPerGar    As Boolean
Dim vlArchHabDes    As Boolean
Dim vlOpen          As Boolean
Dim vlSql           As String, vlMoneda As String
Dim vlLinea         As String

'-------- Constantes para generar el archivo de Pagos en Regimen -------------------
Const clTipRegPR     As String = "5"    'Tipo de Registro
Const clRamoContPR   As String = "76"   'RAMO CONTABLE
Const clFrecPagPR    As String = "0"    'FRECUENCIA DE PAGO
Const clReaPR        As String = "2"    '23-REASEGURADOR
Const clTipMovSinLPR As String = "38"   '61-TIPO MOVIMIENTO SINIESTRO (Liquido)
Const clTipMovSinSPR As String = "39"   '61-TIPO MOVIMIENTO SINIESTRO (Salud)
Const clTipMovSinRPR As String = "40"   '61-TIPO MOVIMIENTO SINIESTRO (Retención)
Const clTipMovSinRes2PR As String = "2"   '61-TIPO MOVIMIENTO SINIESTRO (Resumen)
Const clTipMovSinRes3PR As String = "3"   '61-TIPO MOVIMIENTO SINIESTRO (Resumen)
Const clTipMovSinRes4PR As String = "4"   '61-TIPO MOVIMIENTO SINIESTRO (Resumen)
Const clTipMovSinRes7PR As String = "7"   '61-TIPO MOVIMIENTO SINIESTRO (PROVISION)
Const clTipPerNatPR  As String = "N"    '67-TIPO DE PERSONA (jurídico / natural)
Const clTipPerJurPR  As String = "J"    '67-TIPO DE PERSONA (jurídico / natural)
Const clRetPR        As String = "R"    'Retenido

'-------- Constantes para generar el archivo de Gastos de Sepelio ------------------
Const clTipRegGS     As String = "5"    '1-Tipo de Registro
Const clRamoContGS   As String = "76"   'RAMO CONTABLE
Const clFrecPagGS    As String = "0"    '13-FRECUENCIA DE PAGO
Const clReaGS        As String = "2"    '23-REASEGURADOR
Const clTipMovSinGS  As String = "41"   '61-TIPO MOVIMIENTO SINIESTRO
Const clTipMovHabDes  As String = "99"   '61-TIPO MOVIMIENTO HABERES Y DESCUENTOS
Const clTipPerNatGS  As String = "N"    '67-Tipo de persona (jurídico / natural)

'-------- Constantes para generar el archivo de Periodo Garantizado ----------------
Const clTipRegPG     As String = "5"    'Tipo de Registro
Const clRamoContPG   As String = "76"   'RAMO CONTABLE
Const clFrecPagPG    As String = "0"    'FRECUENCIA DE PAGO
Const clReaPG        As String = "2"    '23-REASEGURADOR
Const clTipMovSinPG  As String = "42"   '61-TIPO MOVIMIENTO SINIESTRO
Const clTipPerNatPG  As String = "N"    '67-Tipo de persona (jurídico / natural)
'-----------------------------------------------------------------------------------

Dim vlVar1  As String       'Tipo de Registro
Dim vlVar2  As String       'POLIZA
Dim vlVar3  As String       'SUCURSAL
Dim vlVar3_Res  As String       'SUCURSAL
Dim vlVar4  As String       'VIGENCIA "DESDE" DE LA POLIZA
Dim vlVar5  As String       '--VIGENCIA "HASTA" DE LA POLIZA
Dim vlVar6  As String       '--VIGENCIA 'DESDE' ORIGINAL
Dim vlVar7  As String       'FECHA CONTABLE(MES / ANO)
Dim vlVar8  As String       'MONEDA DEL MOVIMIENTO
Dim vlVar9  As String       'COBERTURA
Dim vlVar10 As String       'RAMO CONTABLE
Dim vlVar11 As String       'CONTRATANTE RUT
Dim vlVar11_Res As String       'CONTRATANTE RUT
Dim vlVar12 As String       'CONTRATANTE NOMBRE
Dim vlVar12_Res As String       'CONTRATANTE NOMBRE
Dim vlVar13 As String       'FRECUENCIA DE PAGO (INMEDIATA)
Dim vlVar14 As String       '--INTERMERDIARIO GTO DE COBRANZA
Dim vlVar14_Res As String       '--INTERMERDIARIO GTO DE COBRANZA
Dim vlVar15 As String       '--REGISTRO EN TRATAMIENTO
Dim vlVar16 As String       '--FECHA EFECTO REGISTRO (MOV)
Dim vlVar17 As String       'NOMBRE
Dim vlVar18 As String       'RUT
Dim vlVar19 As String       'SUCURSAL
Dim vlVar19_Res As String       'SUCURSAL
Dim vlVar20 As String       '--NOMBRE INTERMEDIARIO
Dim vlVar21 As String       '--RUT INNTERMEDIARIO
Dim vlVar22 As String       '--TIPO DE INTERMEDIARIO
Dim vlVar23 As String       '--REASEGURADOR
Dim vlVar24 As String       'NACIONALIDAD REASEGURADOR
Dim vlVar25 As String       '--CONTRATO DE REASEGURO
Dim vlVar26 As String       'TIPO DE REASEGURO
Dim vlVar27 As String       '--NUMERO DE SINIESTRO
Dim vlVar28 As String       '--ESTADO DEL SINIESTRO
Dim vlVar29 As String       '--TIPO MOVIMIENTO
Dim vlVar30 As String       '--NUMERO DE MOVIMIENTO
Dim vlVar31 As String       '--MONTO EXENTO PRIMA
Dim vlVar32 As String       '--MONTO AFECTO PRIMA
Dim vlVar33 As String       '--MONTO IGV PRIMA
Dim vlVar34 As String       '--MONTO BRUTO PRIMA
Dim vlVar35 As String       '--MONTO NETO PRIMA DEVENGADA
Dim vlVar36 As String       '--CAPITALES ASEGURADOS
Dim vlVar37 As String       '--ORIGEN DEL RECIBO
Dim vlVar38 As String       '--TIPO DE MOVIMIENTO
Dim vlVar39 As String       '--MONTO PRIMA CEDIDA ANTES DSCTO
Dim vlVar40 As String       '--MONTO DESC. POR PRIMA CEDIDA
Dim vlVar41 As String       '--MONTO IMPUESTO 2%
Dim vlVar42 As String       '--MONTO EXCESO DE PERDIDA
Dim vlVar43 As String       '--CAPITALES CEDIDOS
Dim vlVar44 As String       '--TIPO DE RESERVA
Dim vlVar45 As String       '--MONTO RESERVA MATEMATICA
Dim vlVar47 As String       '-- % DE COMSION SOBRE LA PRIMA
Dim vlVar48 As String       '--TIPO DE COMISION
Dim vlVar49 As String       '--MONTO COMISION NETA
Dim vlVar50 As String       '--MONTO IGV COMISION
Dim vlVar51 As String       '--MONTO BRUTO COMISION
Dim vlVar52 As String       '--PERIODO DE GRACIA
Dim vlVar53 As String       '--MONTO NETO COMISION
Dim vlVar54 As String       '--ESQUEMA DE PAGO
Dim vlVar55 As String       '--fecha desde
Dim vlVar56 As String       '--fecha Hasta
Dim vlVar57 As String       '--RAMO
Dim vlVar58 As String       '--PRODUCTO
Dim vlVar59 As String       '--POLIZA
Dim vlVar60 As String       '--RUT DEL CLIENTE
Dim vlVar61 As String       'TIPO DE MOVIMIENTO SINIESTRO
Dim vlVar62 As String       'MONTO
Dim vlVar63 As String       '--MONTO CEDIDO EN EL MES
Dim vlVar63_Res As String       '--MONTO CEDIDO EN EL MES
Dim vlVar64 As String       '-- % DE COMISION DE GASTOS DE COB
Dim vlVar64_Res As String       '-- % DE COMISION DE GASTOS DE COB
Dim vlVar65 As String       '--MTO. GASTOS DE COB. PRIMA REC.
Dim vlVar65_Res As String       '--MTO. GASTOS DE COB. PRIMA REC.
Dim vlVar66 As String       '--MTO. GASTOS DE COB. PRIMA DEV.
Dim vlVar66_Res As String       '--MTO. GASTOS DE COB. PRIMA DEV.
Dim vlVar67 As String       'Tipo de persona (jurídico / natural)
Dim vlVar67_Res As String       'Tipo de persona (jurídico / natural)
Dim vlVar68 As String       'Num cuenta banco
Dim vlVar69 As String       'Num Cuenta CCI
'-------------------------------------------------------------------------------

'Numero de archivo creado
Dim vlNumArchivo As Integer

'Variables de Pagos Recurrentes
Dim vlNumCasosPRPension     As Long
Dim vlNumCasosPRSalud       As Long
Dim vlNumCasosPRRetencion   As Long
Dim vlMtoPRPension          As Double
Dim vlMtoPRSalud            As Double
Dim vlMtoPRRetencion        As Double

'Variables de Gasto de Sepelio
Dim vlNumCasosGtoSep        As Long
Dim vlNumCasosHabDes        As Long
Dim vlMtoGtoSep             As Double
Dim vlMtoHabDes             As Double
'Variables de Periodo Garantizado
Dim vlNumCasosPerGar        As Long
Dim vlMtoPerGar             As Double

'CORPTEC
Dim sTipoPro As String
Dim num_log_FlujoP As Long
'CORPTEC
Function flLog_Proc() As Boolean
    Dim com As ADODB.Command
    Dim sistema, modulo, opcion, origen, tipo As String
    sistema = "SEACSA"
    modulo = "PENSIONES"
    opcion = "PAGOS RECURRENTES"
    origen = "A"
    tipo = "A"

    Set com = New ADODB.Command
    
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
    com.Parameters.Append com.CreateParameter("IDLOG", adDouble, adParamInput, 2, 0)
    com.Parameters.Append com.CreateParameter("Retorno", adDouble, adParamReturnValue)
    com.Execute
    vgConexionBD.CommitTrans
    num_log_FlujoP = com("Retorno")

End Function

Function flLmpGrilla()

    Msf_Periodos.Clear
    Msf_Periodos.rows = 1
    Msf_Periodos.Cols = 8
    Msf_Periodos.RowHeight(0) = 250
    Msf_Periodos.row = 0
    Msf_Periodos.ColWidth(0) = 0
    Msf_Periodos.Col = 1
    Msf_Periodos.Text = "Desde"
    Msf_Periodos.ColWidth(1) = 1100
    Msf_Periodos.Col = 2
    Msf_Periodos.Text = "Hasta"
    Msf_Periodos.ColWidth(2) = 1100
    Msf_Periodos.Col = 3
    Msf_Periodos.Text = "Nº Casos"
    Msf_Periodos.ColWidth(3) = 1200
    Msf_Periodos.Col = 4
    Msf_Periodos.Text = "Usuario"
    Msf_Periodos.ColWidth(4) = 1500
    Msf_Periodos.Col = 5
    Msf_Periodos.Text = "Fecha"
    Msf_Periodos.ColWidth(5) = 1200
    Msf_Periodos.Col = 6
    Msf_Periodos.Text = "Hora"
    Msf_Periodos.ColWidth(6) = 1200
    Msf_Periodos.Col = 7
    Msf_Periodos.Text = "Nº Archivo"
    Msf_Periodos.ColWidth(7) = 1000

End Function

Function flLmpGrillaPagReg()

    Msf_Periodos.Clear
    Msf_Periodos.rows = 1
    Msf_Periodos.Cols = 10
    Msf_Periodos.RowHeight(0) = 250
    Msf_Periodos.row = 0
    Msf_Periodos.ColWidth(0) = 0
    Msf_Periodos.Col = 1
    Msf_Periodos.Text = "Desde"
    Msf_Periodos.ColWidth(1) = 1100
    Msf_Periodos.Col = 2
    Msf_Periodos.Text = "Hasta"
    Msf_Periodos.ColWidth(2) = 1100
    Msf_Periodos.Col = 3
    Msf_Periodos.Text = "Nº Casos Pensión"
    Msf_Periodos.ColWidth(3) = 1500
    Msf_Periodos.Col = 4
    Msf_Periodos.Text = "Nº Casos Salud"
    Msf_Periodos.ColWidth(4) = 1500
    Msf_Periodos.Col = 5
    Msf_Periodos.Text = "Nº Casos Retención"
    Msf_Periodos.ColWidth(5) = 1500
    Msf_Periodos.Col = 6
    Msf_Periodos.Text = "Usuario"
    Msf_Periodos.ColWidth(6) = 1500
    Msf_Periodos.Col = 7
    Msf_Periodos.Text = "Fecha"
    Msf_Periodos.ColWidth(7) = 1200
    Msf_Periodos.Col = 8
    Msf_Periodos.Text = "Hora"
    Msf_Periodos.ColWidth(8) = 1200
    Msf_Periodos.Col = 9
    Msf_Periodos.Text = "Nº Archivo"
    Msf_Periodos.ColWidth(9) = 1000
    
End Function

Function flActGrillaPagReg()

vgQuery = ""
vgQuery = "SELECT distinct num_archivo,fec_desde,fec_hasta,"
vgQuery = vgQuery & "(select sum(num_casos) from pp_tmae_contableregpago where "
vgQuery = vgQuery & "num_archivo=p.num_archivo and cod_tipmov='38')as num_casospension,"
vgQuery = vgQuery & "(select sum(num_casos) from pp_tmae_contableregpago where "
vgQuery = vgQuery & "num_archivo=p.num_archivo and cod_tipmov='39')as num_casossalud,"
vgQuery = vgQuery & "(select sum(num_casos) from pp_tmae_contableregpago where "
vgQuery = vgQuery & "num_archivo=p.num_archivo and cod_tipmov='40')as num_casosretencion,"
vgQuery = vgQuery & "cod_usuariocrea,fec_crea,hor_crea "
vgQuery = vgQuery & "FROM PP_TMAE_CONTABLEREGPAGO P "
vgQuery = vgQuery & "ORDER BY fec_desde desc,fec_hasta desc,fec_crea desc,hor_crea desc"
Set vgRs = vgConexionBD.Execute(vgQuery)
If Not (vgRs.EOF) Then
    vgI = 1
    While Not (vgRs.EOF)
        Msf_Periodos.AddItem (vgI)
        Msf_Periodos.row = vgI
        
        Msf_Periodos.Col = 1
        Msf_Periodos.Text = DateSerial(Mid(vgRs!fec_desde, 1, 4), Mid(vgRs!fec_desde, 5, 2), Mid(vgRs!fec_desde, 7, 2))
        Msf_Periodos.Col = 2
        Msf_Periodos.Text = DateSerial(Mid(vgRs!FEC_HASTA, 1, 4), Mid(vgRs!FEC_HASTA, 5, 2), Mid(vgRs!FEC_HASTA, 7, 2))
        Msf_Periodos.Col = 3
        Msf_Periodos.Text = Format(vgRs!num_casospension, "#,#0")
        Msf_Periodos.Col = 4
        Msf_Periodos.Text = Format(vgRs!num_casossalud, "#,#0")
        Msf_Periodos.Col = 5
        Msf_Periodos.Text = Format(vgRs!num_casosretencion, "#,#0")
        Msf_Periodos.Col = 6
        Msf_Periodos.Text = Trim(vgRs!Cod_UsuarioCrea)
        Msf_Periodos.Col = 7
        Msf_Periodos.Text = DateSerial(Mid(vgRs!Fec_Crea, 1, 4), Mid(vgRs!Fec_Crea, 5, 2), Mid(vgRs!Fec_Crea, 7, 2))
        Msf_Periodos.Col = 8
        Msf_Periodos.Text = Trim(Mid(vgRs!Hor_Crea, 1, 2) & ":" & Mid(vgRs!Hor_Crea, 3, 2) & ":" & Mid(vgRs!Hor_Crea, 5, 2))
        Msf_Periodos.Col = 9
        Msf_Periodos.Text = Trim(vgRs!num_archivo)
        
        vgI = vgI + 1
        vgRs.MoveNext
    Wend
End If
vgRs.Close

End Function

Function flActGrillaGtoSep()

vgQuery = ""
vgQuery = "SELECT num_archivo,fec_desde,fec_hasta,sum(num_casos) as num_casos,"
vgQuery = vgQuery & "cod_usuariocrea,fec_crea,hor_crea "
vgQuery = vgQuery & "FROM PP_TMAE_CONTABLEGTOSEP "
vgQuery = vgQuery & "GROUP BY num_archivo,fec_desde,fec_hasta,"
vgQuery = vgQuery & "cod_usuariocrea,fec_crea,hor_crea "
vgQuery = vgQuery & "ORDER BY fec_desde desc,fec_hasta desc,fec_crea desc,hor_crea desc"
Set vgRs = vgConexionBD.Execute(vgQuery)
If Not (vgRs.EOF) Then
    vgI = 1
    While Not (vgRs.EOF)
        Msf_Periodos.AddItem (vgI)
        Msf_Periodos.row = vgI
        
        Msf_Periodos.Col = 1
        Msf_Periodos.Text = DateSerial(Mid(vgRs!fec_desde, 1, 4), Mid(vgRs!fec_desde, 5, 2), Mid(vgRs!fec_desde, 7, 2))
        Msf_Periodos.Col = 2
        Msf_Periodos.Text = DateSerial(Mid(vgRs!FEC_HASTA, 1, 4), Mid(vgRs!FEC_HASTA, 5, 2), Mid(vgRs!FEC_HASTA, 7, 2))
        Msf_Periodos.Col = 3
        Msf_Periodos.Text = Format(vgRs!NUM_CASOS, "#,#0")
        Msf_Periodos.Col = 4
        Msf_Periodos.Text = Trim(vgRs!Cod_UsuarioCrea)
        Msf_Periodos.Col = 5
        Msf_Periodos.Text = DateSerial(Mid(vgRs!Fec_Crea, 1, 4), Mid(vgRs!Fec_Crea, 5, 2), Mid(vgRs!Fec_Crea, 7, 2))
        Msf_Periodos.Col = 6
        Msf_Periodos.Text = Trim(Mid(vgRs!Hor_Crea, 1, 2) & ":" & Mid(vgRs!Hor_Crea, 3, 2) & ":" & Mid(vgRs!Hor_Crea, 5, 2))
        Msf_Periodos.Col = 7
        Msf_Periodos.Text = Trim(vgRs!num_archivo)
        
        vgI = vgI + 1
        vgRs.MoveNext
    Wend
End If
vgRs.Close

End Function
Function flActGrillaHabDes()

vgQuery = ""
vgQuery = "SELECT num_archivo,fec_desde,fec_hasta,sum(num_casos) as num_casos,"
vgQuery = vgQuery & "cod_usuariocrea,fec_crea,hor_crea "
vgQuery = vgQuery & "FROM PP_TMAE_CONTABLEHABDES "
vgQuery = vgQuery & "GROUP BY num_archivo,fec_desde,fec_hasta,"
vgQuery = vgQuery & "cod_usuariocrea,fec_crea,hor_crea "
vgQuery = vgQuery & "ORDER BY fec_desde desc,fec_hasta desc,fec_crea desc,hor_crea desc"
Set vgRs = vgConexionBD.Execute(vgQuery)
If Not (vgRs.EOF) Then
    vgI = 1
    While Not (vgRs.EOF)
        Msf_Periodos.AddItem (vgI)
        Msf_Periodos.row = vgI
        
        Msf_Periodos.Col = 1
        Msf_Periodos.Text = DateSerial(Mid(vgRs!fec_desde, 1, 4), Mid(vgRs!fec_desde, 5, 2), Mid(vgRs!fec_desde, 7, 2))
        Msf_Periodos.Col = 2
        Msf_Periodos.Text = DateSerial(Mid(vgRs!FEC_HASTA, 1, 4), Mid(vgRs!FEC_HASTA, 5, 2), Mid(vgRs!FEC_HASTA, 7, 2))
        Msf_Periodos.Col = 3
        Msf_Periodos.Text = Format(vgRs!NUM_CASOS, "#,#0")
        Msf_Periodos.Col = 4
        Msf_Periodos.Text = Trim(vgRs!Cod_UsuarioCrea)
        Msf_Periodos.Col = 5
        Msf_Periodos.Text = DateSerial(Mid(vgRs!Fec_Crea, 1, 4), Mid(vgRs!Fec_Crea, 5, 2), Mid(vgRs!Fec_Crea, 7, 2))
        Msf_Periodos.Col = 6
        Msf_Periodos.Text = Trim(Mid(vgRs!Hor_Crea, 1, 2) & ":" & Mid(vgRs!Hor_Crea, 3, 2) & ":" & Mid(vgRs!Hor_Crea, 5, 2))
        Msf_Periodos.Col = 7
        Msf_Periodos.Text = Trim(vgRs!num_archivo)
        
        vgI = vgI + 1
        vgRs.MoveNext
    Wend
End If
vgRs.Close

End Function
Function flActGrillaPerGar()

vgQuery = ""
vgQuery = "SELECT num_archivo,fec_desde,fec_hasta,sum(num_casos) as num_casos,"
vgQuery = vgQuery & "cod_usuariocrea,fec_crea,hor_crea "
vgQuery = vgQuery & "FROM PP_TMAE_CONTABLEPERGAR "
vgQuery = vgQuery & "GROUP BY num_archivo,fec_desde,fec_hasta,"
vgQuery = vgQuery & "cod_usuariocrea,fec_crea,hor_crea "
vgQuery = vgQuery & "ORDER BY fec_desde desc,fec_hasta desc,fec_crea desc,hor_crea desc"
Set vgRs = vgConexionBD.Execute(vgQuery)
If Not (vgRs.EOF) Then
    vgI = 1
    While Not (vgRs.EOF)
        Msf_Periodos.AddItem (vgI)
        Msf_Periodos.row = vgI
        
        Msf_Periodos.Col = 1
        Msf_Periodos.Text = DateSerial(Mid(vgRs!fec_desde, 1, 4), Mid(vgRs!fec_desde, 5, 2), Mid(vgRs!fec_desde, 7, 2))
        Msf_Periodos.Col = 2
        Msf_Periodos.Text = DateSerial(Mid(vgRs!FEC_HASTA, 1, 4), Mid(vgRs!FEC_HASTA, 5, 2), Mid(vgRs!FEC_HASTA, 7, 2))
        Msf_Periodos.Col = 3
        Msf_Periodos.Text = Format(vgRs!NUM_CASOS, "#,#0")
        Msf_Periodos.Col = 4
        Msf_Periodos.Text = Trim(vgRs!Cod_UsuarioCrea)
        Msf_Periodos.Col = 5
        Msf_Periodos.Text = DateSerial(Mid(vgRs!Fec_Crea, 1, 4), Mid(vgRs!Fec_Crea, 5, 2), Mid(vgRs!Fec_Crea, 7, 2))
        Msf_Periodos.Col = 6
        Msf_Periodos.Text = Trim(Mid(vgRs!Hor_Crea, 1, 2) & ":" & Mid(vgRs!Hor_Crea, 3, 2) & ":" & Mid(vgRs!Hor_Crea, 5, 2))
        Msf_Periodos.Col = 7
        Msf_Periodos.Text = Trim(vgRs!num_archivo)
        
        vgI = vgI + 1
        vgRs.MoveNext
    Wend
End If
vgRs.Close

End Function

Private Sub flImpresion()
Dim vlArchivo As String
Err.Clear
On Error GoTo Errores1
   
    Screen.MousePointer = 11
    
    If (Trim(Lbl_NumArchivo) = "") Then
        MsgBox "Debe seleccionar un Periodo a Imprimir.", vbInformation, "Falta Información"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    If (vgNomForm = "PagRec") Then
        Call CargaReporteContable
        'Call CargaReporteContableProvision
        Exit Sub
'        vlArchivo = strRpt & "PP_Rpt_ContableResPagReg.rpt"   '\Reportes
'        vgQuery = "{PP_TMAE_CONTABLEREGPAGO.NUM_ARCHIVO} = " & Trim(Lbl_NumArchivo)
    ElseIf (vgNomForm = "GtoSep") Then
        vlArchivo = strRpt & "PP_Rpt_ContableResGtoSep.rpt"   '\Reportes
        vgQuery = "{PP_TMAE_CONTABLEGTOSEP.NUM_ARCHIVO} = " & Trim(Lbl_NumArchivo)
    ElseIf (vgNomForm = "HabDes") Then
        Call CargaReporte(Trim(Lbl_NumArchivo))
        Exit Sub
'        vlArchivo = strRpt & "PP_Rpt_ContableResGtoSep.rpt"   '\Reportes
'        vgQuery = "{PP_TMAE_CONTABLEGTOSEP.NUM_ARCHIVO} = " & Trim(Lbl_NumArchivo)
    Else
        vlArchivo = strRpt & "PP_Rpt_ContableResPerGar.rpt"   '\Reportes
        vgQuery = "{PP_TMAE_CONTABLEPERGAR.NUM_ARCHIVO} = " & Trim(Lbl_NumArchivo)
    End If
    
    If Not fgExiste(vlArchivo) Then     ', vbNormal
        MsgBox "Archivo de Reporte de Resumen de Archivo Contable no se encuentra en el directorio de la Aplicación.", 16, "Archivo no encontrado"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    Rpt_General.Reset
    Rpt_General.WindowState = crptMaximized
    Rpt_General.ReportFileName = vlArchivo
    Rpt_General.Connect = vgRutaDataBase
    Rpt_General.Destination = crptToWindow
    Rpt_General.SelectionFormula = ""
    Rpt_General.SelectionFormula = vgQuery
    
    Rpt_General.Formulas(0) = ""
    Rpt_General.Formulas(1) = ""
    Rpt_General.Formulas(2) = ""
    Rpt_General.Formulas(3) = ""
    
    Rpt_General.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
    Rpt_General.Formulas(1) = "NombreSistema = '" & vgNombreSistema & "'"
    Rpt_General.Formulas(2) = "NombreSubSistema = '" & vgNombreSubSistema & "'"

    Rpt_General.WindowTitle = "Informe Resumen Archivo Contable"
    Rpt_General.Action = 1
    
    Screen.MousePointer = 0
   
Exit Sub
Errores1:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " " & "Numero de Error : " & Err.Number & Err.HelpFile & Err.HelpContext, vbExclamation, "Mensaje de error ..."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub Cmd_Cargar_Click()
On Error GoTo Err_Cargar
    
'    If chkPagoD.Value = 0 Or chkPagoA.Value = 0 Then
'       MsgBox "Debe elegir una opcion de Generacion. Pago Directo o Pago AFP o Ambas", vbInformation, "Falta Información"
'       Exit Sub
'    End If

    
    Lbl_FecCierre = Format(Now, "dd/mm/yyyy Hh:Nn:Ss AMPM")
    
    'Validación de Datos
    'Periodo Desde
    If Txt_Desde = "" Then
       MsgBox "Debe Ingresar Fecha Desde.", vbInformation, "Falta Información"
       Txt_Desde.SetFocus
       Exit Sub
    End If
    If Not IsDate(Txt_Desde) Then
        MsgBox "La Fecha Desde ingresada no es válida.", vbInformation, "Error de Datos"
        Screen.MousePointer = 0
        Txt_Desde.Text = ""
        Exit Sub
    End If
    
    'Periodo Hasta
    If Txt_Hasta = "" Then
       MsgBox "Debe Ingresar Fecha Hasta.", vbInformation, "Falta Información"
       Txt_Hasta.SetFocus
       Exit Sub
    End If
    If Not IsDate(Txt_Hasta) Then
        MsgBox "La Fecha Hasta ingresada no es válida.", vbInformation, "Error de Datos"
        Screen.MousePointer = 0
        Txt_Hasta.Text = ""
        Exit Sub
    End If
        
    If CDate(Txt_Desde) > CDate(Txt_Hasta) Then
       MsgBox "La Fecha Desde debe ser menor o igual a la Fecha Hasta.", vbInformation, "Error de Datos"
       Txt_Hasta.SetFocus
       Exit Sub
    End If

    If LblArchivo = "" Then
        MsgBox "Debe seleccionar Archivo a generar.", vbInformation, "Falta Información"
        Screen.MousePointer = 0
        CmdContable.SetFocus
        Exit Sub
    End If
    
    vlFecDesde = Format(CDate(Trim(Txt_Desde)), "yyyymmdd")
    vlFecHasta = Format(CDate(Trim(Txt_Hasta)), "yyyymmdd")
    
    vgSw = True
    
    vgRes = MsgBox(" ¿ Está Seguro que desea 'Generar' el Archivo Contable pare el Período definido ? ", 4 + 32 + 256, "Archivo Contable")
    If vgRes = 6 Then
        vgSw = False
    Else
        Exit Sub
    End If
    
    Screen.MousePointer = 11
                    
    'CORPTEC
 
    sTipoPro = "I"
    Call flLog_Proc
    
                    
    If (vgNomForm = "PagRec") Then
        
        
        'Genera el archivo de Pagos en Regimen
        vlArchPagReg = flExportarPagReg(vlFecDesde, vlFecHasta)
        
        If (vlArchPagReg = False) Then
            MsgBox "Se Ha Producido un error durante el proceso de Generación del Archivo Contable", vbCritical, "Proceso Cancelado"
            Exit Sub
        Else
            'Genera el archivo de Provision
            'MARCO
            'vlArchPagRegProv = flExportarPagRegProvision(vlFecDesde, vlFecHasta)
            'If (vlArchPagRegProv = False) Then
            '    MsgBox "Se Ha Producido un error durante el proceso de Generación del Archivo Contable de Provision", vbCritical, "Proceso Cancelado"
            '    Exit Sub
            'End If
            'Actualizar la Grilla de Datos
            Call flLmpGrillaPagReg
            Call flActGrillaPagReg
            Screen.MousePointer = 0
            MsgBox "El Proceso ha finalizado Exitosamente.", vbInformation, "Proceso Generación."
        End If
        
        
    ElseIf (vgNomForm = "GtoSep") Then
        'Genera el archivo de Gastos de Sepelio
        vlArchGtoSep = flExportarGtoSep(vlFecDesde, vlFecHasta)
        If (vlArchGtoSep = False) Then
            MsgBox "Se Ha Producido un error durante el proceso de Generación del Archivo Contable", vbCritical, "Proceso Cancelado"
            Exit Sub
        Else
            'Actualizar la Grilla de Datos
            Call flLmpGrilla
            Call flActGrillaGtoSep
            Screen.MousePointer = 0
            MsgBox "El Proceso ha finalizado Exitosamente.", vbInformation, "Proceso Generación."
        End If
    ElseIf (vgNomForm = "HabDes") Then
        
        vlArchHabDes = flExportarHabDes(vlFecDesde, vlFecHasta)
        If (vlArchHabDes = False) Then
            MsgBox "Se Ha Producido un error durante el proceso de Generación del Archivo Contable", vbCritical, "Proceso Cancelado"
            Exit Sub
        Else
            'Actualizar la Grilla de Datos
            Call flLmpGrilla
            Call flActGrillaHabDes
            Screen.MousePointer = 0
            MsgBox "El Proceso ha finalizado Exitosamente.", vbInformation, "Proceso Generación."
        End If
        
    Else
        'Genera el archivo de Periodo Garantizado
        vlArchPerGar = flExportarPerGar(vlFecDesde, vlFecHasta)
        If (vlArchPerGar = False) Then
            MsgBox "Se Ha Producido un error durante el proceso de Generación del Archivo Contable", vbCritical, "Proceso Cancelado"
            Exit Sub
        Else
            'Actualizar la Grilla de Datos
            Call flLmpGrilla
            Call flActGrillaPerGar
            Screen.MousePointer = 0
            MsgBox "El Proceso ha finalizado Exitosamente.", vbInformation, "Proceso Generación."
        End If
    End If
    
    'CORPTEC
    sTipoPro = "F"
    Call flLog_Proc
    
    Screen.MousePointer = 0
        
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub
Private Sub CargaReporteContableDetalleAfpMon()
    Dim rs As ADODB.Recordset
    Dim objRep As New ClsReporte
    Dim cadena As String
    
    On Error GoTo mierror
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

    cadena = "select C.NUM_ARCHIVO,C.FEC_DESDE,C.FEC_HASTA,C.COD_USUARIOCREA,C.FEC_CREA,C.HOR_CREA," & _
            " d.COD_TIPREG , d.COD_TIPMOV,d.Cod_Moneda, d.COD_AFP, d.MTO_PAGO" & _
            " from PP_TMAE_CONTABLEREGPAGO C INNER JOIN PP_TMAE_CONTABLEDETREGPAGO D ON C.NUM_ARCHIVO=D.NUM_ARCHIVO and C.COD_TIPREG=D.COD_TIPREG AND" & _
            " C.COD_TIPMOV=D.COD_TIPMOV AND C.COD_MONEDA=D.COD_MONEDA where c.NUM_ARCHIVO='" & Trim(Lbl_NumArchivo) & "'"
         
    rs.Open cadena, vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF Then
        MsgBox "No hay datos para mostrar", vbExclamation, "Pago de pensiones"
        Exit Sub
    End If
    
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PP_Rpt_ContableDetPagRegAfpMon.rpt"), ".RPT", ".TTX"), 1)
    
    If objRep.CargaReporte(strRpt & "", "PP_Rpt_ContableDetPagRegAfpMon.rpt", "Informe Detallado de Archivo Contable", rs, True, _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("NombreSistema", vgNombreSistema), _
                            ArrFormulas("NombreSubSistema", vgNombreSubSistema)) = False Then
                            
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
    End If
    
    Exit Sub
mierror:
    MsgBox "No se pudo cargar el reporte", vbInformation, "Pensiones"
    
End Sub
Private Sub CargaReporteContableDetalleAfpMon_Provision()
    Dim rs As ADODB.Recordset
    Dim objRep As New ClsReporte
    Dim cadena As String
    
    On Error GoTo mierror
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

    cadena = "select (CASE WHEN B.num_poliza IS NULL THEN 'N' ELSE 'S' END)AS ESTADO,X.NUM_ARCHIVO,X.FEC_DESDE,X.FEC_HASTA,X.COD_USUARIOCREA,X.FEC_CREA,X.HOR_CREA, A.COD_TIPREG , A.COD_TIPMOV," & _
            " A.Cod_Moneda , A.COD_AFP, A.MTO_PAGO, A.GLS_NOMBRE" & _
            " FROM PP_TMAE_CONTABLE_PROVISION X INNER JOIN PP_TMAE_CONTABLEDET_PROVISION A ON X.NUM_ARCHIVO = A.NUM_ARCHIVO And" & _
            " X.COD_TIPREG = A.COD_TIPREG And X.COD_TIPMOV = A.COD_TIPMOV And X.Cod_Moneda = A.Cod_Moneda" & _
            " left join PP_TMAE_CONTABLEDETREGPAGO B ON A.NUM_ARCHIVO = B.NUM_ARCHIVO And A.COD_TIPREG = B.COD_TIPREG And" & _
            " A.COD_TIPMOV = B.COD_TIPMOV And A.Cod_Moneda = B.Cod_Moneda And A.num_poliza = B.num_poliza where X.NUM_ARCHIVO='" & Trim(Lbl_NumArchivo) & "'" & _
            " GROUP BY X.NUM_ARCHIVO,X.FEC_DESDE,X.FEC_HASTA,X.COD_USUARIOCREA,X.FEC_CREA,X.HOR_CREA, A.COD_TIPREG , A.COD_TIPMOV," & _
            " A.Cod_Moneda, A.COD_AFP, A.MTO_PAGO ,A.GLS_NOMBRE,B.num_poliza "
            
    rs.Open cadena, vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF Then
        MsgBox "No hay datos para mostrar", vbExclamation, "Pago de pensiones"
        Exit Sub
    End If
    
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PP_Rpt_ContableDetPagRegAfpMonProv.rpt"), ".RPT", ".TTX"), 1)
    
    If objRep.CargaReporte(strRpt & "", "PP_Rpt_ContableDetPagRegAfpMonProv.rpt", "Informe Detallado de Archivo Contable de Provision", rs, True, _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("NombreSistema", vgNombreSistema), _
                            ArrFormulas("NombreSubSistema", vgNombreSubSistema)) = False Then
                            
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
    End If
    
    Exit Sub
mierror:
    MsgBox "No se pudo cargar el reporte", vbInformation, "Pensiones"
    
End Sub
Private Sub flImpAfpMon()
Dim vlArchivo As String
Err.Clear
On Error GoTo Err_ImpAfpMon
   
    Screen.MousePointer = 11
    
    If (Trim(Lbl_NumArchivo) = "") Then
        MsgBox "Debe seleccionar un Periodo a Imprimir.", vbInformation, "Falta Información"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    If (vgNomForm = "PagRec") Then
        Call CargaReporteContableDetalleAfpMon
        Call CargaReporteContableDetalleAfpMon_Provision
        Exit Sub
'        vlArchivo = strRpt & "PP_Rpt_ContableDetPagRegAfpMon.rpt"   '\Reportes
'        vgQuery = "{PP_TMAE_CONTABLEDETREGPAGO.NUM_ARCHIVO} = " & Trim(Lbl_NumArchivo)
    End If
    
    If Not fgExiste(vlArchivo) Then     ', vbNormal
        MsgBox "Archivo de Reporte de Detalle de Archivo Contable no se encuentra en el directorio de la Aplicación.", 16, "Archivo no encontrado"
        Screen.MousePointer = 0
        Exit Sub
    End If
        
    Rpt_General.Reset
    Rpt_General.WindowState = crptMaximized
    Rpt_General.ReportFileName = vlArchivo
    Rpt_General.Connect = vgRutaDataBase
    Rpt_General.Destination = crptToWindow
    Rpt_General.SelectionFormula = ""
    Rpt_General.SelectionFormula = vgQuery
    
    Rpt_General.Formulas(0) = ""
    Rpt_General.Formulas(1) = ""
    Rpt_General.Formulas(2) = ""
    Rpt_General.Formulas(3) = ""
    
    Rpt_General.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
    Rpt_General.Formulas(1) = "NombreSistema = '" & vgNombreSistema & "'"
    Rpt_General.Formulas(2) = "NombreSubSistema = '" & vgNombreSubSistema & "'"
    
    Rpt_General.WindowTitle = "Informe Resumen Archivo Contable por AFP - Moneda"
    Rpt_General.Action = 1
    
    Screen.MousePointer = 0
   
Exit Sub
Err_ImpAfpMon:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " " & "Numero de Error : " & Err.Number & Err.HelpFile & Err.HelpContext, vbExclamation, "Mensaje de error ..."
        Screen.MousePointer = 0
    End If
End Sub
Private Sub CargaReporteContable()
    Dim rs As ADODB.Recordset
    Dim objRep As New ClsReporte
    Dim cadena As String
    
    On Error GoTo mierror
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

    cadena = "select NUM_ARCHIVO,COD_TIPREG,COD_TIPMOV,COD_MONEDA,FEC_DESDE,FEC_HASTA,NUM_CASOS,MTO_PAGO," & _
             "COD_USUARIOCREA,FEC_CREA,HOR_CREA from PP_TMAE_CONTABLEREGPAGO where NUM_ARCHIVO='" & Trim(Lbl_NumArchivo) & "'"
         
    rs.Open cadena, vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF Then
        MsgBox "No hay datos para mostrar", vbExclamation, "Pago de pensiones"
        Exit Sub
    End If
    
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PP_Rpt_ContableResPagReg.rpt"), ".RPT", ".TTX"), 1)
    
    If objRep.CargaReporte(strRpt & "", "PP_Rpt_ContableResPagReg.rpt", "Informe Resumen Archivo Contable", rs, True, _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("NombreSistema", vgNombreSistema), _
                            ArrFormulas("NombreSubSistema", vgNombreSubSistema)) = False Then
                            
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
    End If
    
    Exit Sub
mierror:
    MsgBox "No se pudo cargar el reporte", vbInformation, "Pensiones"
    
End Sub

Private Sub CargaReporteContableProvision()
    Dim rs As ADODB.Recordset
    Dim objRep As New ClsReporte
    Dim cadena As String
    
    On Error GoTo mierror
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

'    cadena = "select NUM_ARCHIVO,COD_TIPREG,COD_TIPMOV,COD_MONEDA,FEC_DESDE,FEC_HASTA,NUM_CASOS,MTO_PAGO," & _
'             "COD_USUARIOCREA,FEC_CREA,HOR_CREA from PP_TMAE_CONTABLE_PROVISION where NUM_ARCHIVO='" & Trim(Lbl_NumArchivo) & "'"
             
    cadena = "select (CASE WHEN B.num_poliza IS NULL THEN 'N' ELSE 'S' END)AS ESTADO,X.NUM_ARCHIVO,X.COD_TIPREG,X.COD_TIPMOV," & _
            " X.Cod_Moneda , X.FEC_DESDE, X.FEC_HASTA, X.NUM_CASOS, A.MTO_PAGO, X.CPP_TMAE_CONTABLEDET_PROVISIONod_UsuarioCrea, X.Fec_Crea, X.Hor_Crea, A.GLS_NOMBRE" & _
            " FROM PP_TMAE_CONTABLE_PROVISION X INNER JOIN  A ON X.NUM_ARCHIVO = A.NUM_ARCHIVO And" & _
            " X.COD_TIPREG = A.COD_TIPREG And X.COD_TIPMOV = A.COD_TIPMOV And X.Cod_Moneda = A.Cod_Moneda" & _
            " left join PP_TMAE_CONTABLEDETREGPAGO B ON A.NUM_ARCHIVO = B.NUM_ARCHIVO And A.COD_TIPREG = B.COD_TIPREG And" & _
            " A.COD_TIPMOV = B.COD_TIPMOV And A.Cod_Moneda = B.Cod_Moneda And A.num_poliza = B.num_poliza" & _
            " where X.NUM_ARCHIVO='" & Trim(Lbl_NumArchivo) & "'" & _
            " GROUP BY X.NUM_ARCHIVO,X.COD_TIPREG,X.COD_TIPMOV,X.COD_MONEDA,X.FEC_DESDE,X.FEC_HASTA,X.NUM_CASOS,A.MTO_PAGO," & _
            " X.COD_USUARIOCREA,X.FEC_CREA,X.HOR_CREA ,A.GLS_NOMBRE ,B.num_poliza"
            
    rs.Open cadena, vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF Then
        MsgBox "No hay datos para mostrar", vbExclamation, "Pago de pensiones"
        Exit Sub
    End If
    
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PP_Rpt_ContableResPagReg_Prov.rpt"), ".RPT", ".TTX"), 1)
    
    If objRep.CargaReporte(strRpt & "", "PP_Rpt_ContableResPagReg_Prov.rpt", "Informe Resumen Archivo Contable de Provisiones", rs, True, _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("NombreSistema", vgNombreSistema), _
                            ArrFormulas("NombreSubSistema", vgNombreSubSistema)) = False Then
                            
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
    End If
    
    Exit Sub
mierror:
    MsgBox "No se pudo cargar el reporte", vbInformation, "Pensiones"
    
End Sub
Private Sub CargaReporteContableDetalle()
    Dim rs As ADODB.Recordset
    Dim objRep As New ClsReporte
    Dim cadena As String
    
    On Error GoTo mierror
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

    cadena = "select C.NUM_ARCHIVO,C.FEC_DESDE,C.FEC_HASTA,C.COD_USUARIOCREA,C.FEC_CREA,C.HOR_CREA," & _
            " d.COD_TIPREG , d.COD_TIPMOV, d.num_poliza, d.Fec_Pago, d.Cod_Moneda, d.Cod_TipPension, d.COD_AFP, d.GLS_NOMBRE, d.MTO_PAGO," & _
            "(select cod_adicional from ma_tpar_tabcod where cod_tabla='TP' and cod_elemento=d.Cod_TipPension)as Abrv_Pension," & _
            "(select cod_adicional from ma_tpar_tabcod where cod_tabla='AF' and cod_elemento=d.COD_AFP)as Abrv_afp," & _
            "(select cod_adicional from ma_tpar_tabcod where cod_tabla='PA' and cod_elemento=(select cod_par from pp_tmae_ben x where x.num_poliza=d.num_poliza and x.num_orden=d.num_orden and x.num_endoso=(select max(num_endoso) from pp_tmae_ben where num_poliza=d.num_poliza and num_orden=d.num_orden)))as Abrv_Parentesco," & _
            "(select GLS_ELEMENTO from ma_tpar_tabcod where cod_tabla='TM' AND COD_ELEMENTO=d.Cod_Moneda)AS Moneda" & _
            " from PP_TMAE_CONTABLEREGPAGO C INNER JOIN PP_TMAE_CONTABLEDETREGPAGO D ON C.NUM_ARCHIVO=D.NUM_ARCHIVO and C.COD_TIPREG=D.COD_TIPREG AND" & _
            " C.COD_TIPMOV=D.COD_TIPMOV AND C.COD_MONEDA=D.COD_MONEDA where c.NUM_ARCHIVO='" & Trim(Lbl_NumArchivo) & "'"

    rs.Open cadena, vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF Then
        MsgBox "No hay datos para mostrar", vbExclamation, "Pago de pensiones"
        Exit Sub
    End If
    
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PP_Rpt_ContableDetPagReg.rpt"), ".RPT", ".TTX"), 1)
    
    If objRep.CargaReporte(strRpt & "", "PP_Rpt_ContableDetPagReg.rpt", "Informe Detallado de Archivo Contable", rs, True, _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("NombreSistema", vgNombreSistema), _
                            ArrFormulas("NombreSubSistema", vgNombreSubSistema)) = False Then
                            
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
    End If
    
    Exit Sub
mierror:
    MsgBox "No se pudo cargar el reporte", vbInformation, "Pensiones"
    
End Sub

Private Sub CargaReporteContablePendienteAcumulada()
    Dim rs As ADODB.Recordset
    Dim objRep As New ClsReporte
    Dim cadena As String
    Dim dtFechaDesde As Date
    Dim dtFechaHasta As Date
    
    On Error GoTo mierror
    
    dtFechaDesde = Txt_Desde.Text
    dtFechaHasta = Txt_Hasta.Text
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

    rs.Open "PP_LISTA_PEND_ACUMULADO.LISTAR('" & Format(dtFechaDesde, "yyyymmdd") & "','" & Format(dtFechaHasta, "yyyymmdd") & "')", vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF Then
        MsgBox "No hay datos para mostrar", vbExclamation, "Pagos Pendientes Acumulados"
        Exit Sub
    End If
    
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PP_Rpt_PendienteAcumulado.rpt"), ".RPT", ".TTX"), 1)
    
    If objRep.CargaReporte(strRpt & "", "PP_Rpt_PendienteAcumulado.rpt", "Informe Detallado de pagos pendientes Acumulado", rs, True, _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("NombreSistema", vgNombreSistema), _
                            ArrFormulas("NombreSubSistema", vgNombreSubSistema), _
                            ArrFormulas("desde", dtFechaDesde), _
                            ArrFormulas("hasta", dtFechaHasta)) = False Then
                            
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
    End If
    
    Exit Sub
mierror:
    MsgBox "No se pudo cargar el reporte", vbInformation, "Pensiones"
    
End Sub

Private Sub CargaReporteContableProvisionDetalle()
    Dim rs As ADODB.Recordset
    Dim objRep As New ClsReporte
    Dim cadena As String
    
    On Error GoTo mierror
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

'    cadena = "select C.NUM_ARCHIVO,C.FEC_DESDE,C.FEC_HASTA,C.COD_USUARIOCREA,C.FEC_CREA,C.HOR_CREA," & _
'            " d.COD_TIPREG , d.COD_TIPMOV, d.num_poliza, d.Fec_Pago, d.Cod_Moneda, d.Cod_TipPension, d.COD_AFP, d.GLS_NOMBRE, d.MTO_PAGO" & _
'            " from PP_TMAE_CONTABLE_PROVISION C INNER JOIN PP_TMAE_CONTABLEDET_PROVISION D ON C.NUM_ARCHIVO=D.NUM_ARCHIVO and C.COD_TIPREG=D.COD_TIPREG AND" & _
'            " C.COD_TIPMOV=D.COD_TIPMOV AND C.COD_MONEDA=D.COD_MONEDA where c.NUM_ARCHIVO='" & Trim(Lbl_NumArchivo) & "'"
    '(CASE WHEN B.num_poliza IS NULL THEN 'N' ELSE 'S' END)AS ESTADO,
    cadena = "SELECT X.NUM_ARCHIVO,X.FEC_DESDE,X.FEC_HASTA,X.COD_USUARIOCREA,X.FEC_CREA,X.HOR_CREA," & _
        " A.COD_TIPREG , A.COD_TIPMOV, A.num_poliza, B.Fec_Pago, A.Cod_Moneda, A.Cod_TipPension, A.COD_AFP, A.GLS_NOMBRE, A.MTO_PAGO," & _
        "(select cod_adicional from ma_tpar_tabcod where cod_tabla='TP' and cod_elemento=A.Cod_TipPension)as Abrv_Pension," & _
        "(select cod_adicional from ma_tpar_tabcod where cod_tabla='AF' and cod_elemento=A.COD_AFP)as Abrv_afp," & _
        "(select cod_adicional from ma_tpar_tabcod where cod_tabla='PA' and cod_elemento=(select cod_par from pp_tmae_ben x where x.num_poliza=A.num_poliza and x.num_orden=A.num_orden and x.num_endoso=(select max(num_endoso) from pp_tmae_ben where num_poliza=A.num_poliza and num_orden=A.num_orden)))as Abrv_Parentesco," & _
        "(select GLS_ELEMENTO from ma_tpar_tabcod where cod_tabla='TM' AND COD_ELEMENTO=A.Cod_Moneda)AS Moneda" & _
        " FROM PP_TMAE_CONTABLE_PROVISION X INNER JOIN PP_TMAE_CONTABLEDET_PROVISION A ON" & _
        " X.NUM_ARCHIVO = A.NUM_ARCHIVO And X.COD_TIPREG = A.COD_TIPREG And X.COD_TIPMOV = A.COD_TIPMOV And X.Cod_Moneda = A.Cod_Moneda" & _
        " left join PP_TMAE_CONTABLEDETREGPAGO B ON" & _
        " A.NUM_ARCHIVO = B.NUM_ARCHIVO And A.COD_TIPREG = B.COD_TIPREG And A.COD_TIPMOV = B.COD_TIPMOV And A.Cod_Moneda = B.Cod_Moneda" & _
        " AND A.NUM_POLIZA=B.NUM_POLIZA AND A.NUM_ORDEN=B.NUM_ORDEN WHERE A.NUM_ARCHIVO='" & Trim(Lbl_NumArchivo) & "'" & _
        " GROUP BY X.NUM_ARCHIVO,X.FEC_DESDE,X.FEC_HASTA,X.COD_USUARIOCREA,X.FEC_CREA,X.HOR_CREA, A.COD_TIPREG , A.COD_TIPMOV, A.num_poliza, B.Fec_Pago, A.Cod_Moneda," & _
        " A.Cod_TipPension, A.COD_AFP, A.GLS_NOMBRE, A.MTO_PAGO ,B.num_poliza,A.NUM_ORDEN"
        
    rs.Open cadena, vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF Then
        MsgBox "No hay datos para mostrar", vbExclamation, "Pagos Pendientes"
        Exit Sub
    End If
    
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PP_Rpt_ContableDetPagReg_Prov.rpt"), ".RPT", ".TTX"), 1)
    
    If objRep.CargaReporte(strRpt & "", "PP_Rpt_ContableDetPagReg_Prov.rpt", "Informe Detallado de Archivo Contable de Provisión", rs, True, _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("NombreSistema", vgNombreSistema), _
                            ArrFormulas("NombreSubSistema", vgNombreSubSistema)) = False Then
                            
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
    End If
    
    Exit Sub
mierror:
    MsgBox "No se pudo cargar el reporte", vbInformation, "Pensiones"
    
End Sub
Private Sub flImpDetalle()
Dim vlArchivo As String
Err.Clear
On Error GoTo Errores1
   
    Screen.MousePointer = 11
    
    If (Trim(Lbl_NumArchivo) = "") Then
        MsgBox "Debe seleccionar un Periodo a Imprimir.", vbInformation, "Falta Información"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    If (vgNomForm = "PagRec") Then
        Call CargaReporteContableDetalle
        Call CargaReporteContableProvisionDetalle
        Exit Sub
'        vlArchivo = strRpt & "PP_Rpt_ContableDetPagReg.rpt"   '\Reportes
'        vgQuery = "{PP_TMAE_CONTABLEDETREGPAGO.NUM_ARCHIVO} = " & Trim(Lbl_NumArchivo)
    ElseIf (vgNomForm = "GtoSep") Then
        Call CargaReporteContableGtoSepelio
        Exit Sub
'        vlArchivo = strRpt & "PP_Rpt_ContableDetGtoSep.rpt"   '\Reportes
'        vgQuery = "{PP_TMAE_CONTABLEDETGTOSEP.NUM_ARCHIVO} = " & Trim(Lbl_NumArchivo)
    ElseIf (vgNomForm = "PerGar") Then
        vlArchivo = strRpt & "PP_Rpt_ContableDetPerGar.rpt"   '\Reportes
        vgQuery = "{PP_TMAE_CONTABLEDETPERGAR.NUM_ARCHIVO} = " & Trim(Lbl_NumArchivo)
    ElseIf (vgNomForm = "HabDes") Then
        Call CargaReporteDetallado(Trim(Lbl_NumArchivo))
        Exit Sub
        'vlArchivo = strRpt & "PP_Rpt_ContableDetHabDes.rpt"   '\Reportes
        'vgQuery = "{PP_TMAE_CONTABLEDETHABDES.NUM_ARCHIVO} = " & Trim(Lbl_NumArchivo)
    End If
    
    If Not fgExiste(vlArchivo) Then     ', vbNormal
        MsgBox "Archivo de Reporte de Detalle de Archivo Contable no se encuentra en el directorio de la Aplicación.", 16, "Archivo no encontrado"
        Screen.MousePointer = 0
        Exit Sub
    End If
        
    Rpt_General.Reset
    Rpt_General.WindowState = crptMaximized
    Rpt_General.ReportFileName = vlArchivo
    Rpt_General.Connect = vgRutaDataBase
    Rpt_General.Destination = crptToWindow
    Rpt_General.SelectionFormula = ""
    Rpt_General.SelectionFormula = vgQuery
    
    Rpt_General.Formulas(0) = ""
    Rpt_General.Formulas(1) = ""
    Rpt_General.Formulas(2) = ""
    Rpt_General.Formulas(3) = ""
    
    Rpt_General.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
    Rpt_General.Formulas(1) = "NombreSistema = '" & vgNombreSistema & "'"
    Rpt_General.Formulas(2) = "NombreSubSistema = '" & vgNombreSubSistema & "'"
    
    Rpt_General.WindowTitle = "Informe Resumen Archivo Contable"
    Rpt_General.Action = 1
    
    Screen.MousePointer = 0
   
Exit Sub
Errores1:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " " & "Numero de Error : " & Err.Number & Err.HelpFile & Err.HelpContext, vbExclamation, "Mensaje de error ..."
        Screen.MousePointer = 0
    End If
End Sub
Private Sub CargaReporteContableGtoSepelio()
    Dim rs As ADODB.Recordset
    Dim objRep As New ClsReporte
    Dim cadena As String
    
    On Error GoTo mierror
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

    cadena = "select cg.num_archivo, cg.cod_tipreg, cg.cod_tipmov, cg.cod_moneda, cg.fec_desde, cg.fec_hasta, cg.num_casos, cg.mto_pago, cg.cod_usuariocrea, cg.fec_crea, cg.hor_crea," & _
             " cd.num_poliza, cd.fec_pago, cd.cod_tippension, tv.gls_elemento as gls_tippension, cd.cod_moneda as cod_moneda_det, cd.cod_afp, af.gls_elemento as gls_afp, cd.gls_nombre, cd.mto_pago as mto_pago_det, po.cod_cuspp" & _
             " from pp_tmae_contablegtosep cg" & _
             " left join pp_tmae_contabledetgtosep cd on cg.num_archivo = cd.num_archivo and cg.cod_tipreg = cd.cod_tipreg and cg.cod_tipmov = cd.cod_tipmov and cg.cod_moneda = cd.cod_moneda" & _
             " left join PD_TMAE_POLIZA po on po.num_poliza = cd.num_poliza" & _
             " left join MA_TPAR_TABCOD TV on TV.COD_TABLA = 'TP' AND TV.COD_ELEMENTO = cd.COD_TIPPENSION" & _
             " left join MA_TPAR_TABCOD AF on AF.COD_TABLA = 'AF' AND AF.COD_ELEMENTO = cd.COD_AFP" & _
             " where po.num_endoso=(select max(num_endoso) from PD_TMAE_POLIZA where num_poliza=po.num_poliza)" & _
             " and cd.num_archivo = '" & Trim(Lbl_NumArchivo) & "'"
         
    rs.Open cadena, vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF Then
        MsgBox "No hay datos para mostrar", vbExclamation, "Pago de pensiones"
        Exit Sub
    End If
    
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PP_Rpt_ContableDetGtoSep.rpt"), ".RPT", ".TTX"), 1)
    
    If objRep.CargaReporte(strRpt & "", "PP_Rpt_ContableDetGtoSep.rpt", "Informe Resumen Archivo Contable", rs, True, _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("NombreSistema", vgNombreSistema), _
                            ArrFormulas("NombreSubSistema", vgNombreSubSistema)) = False Then
                            
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
    End If
    
    Exit Sub
mierror:
    MsgBox "No se pudo cargar el reporte", vbInformation, "Pensiones"
    
End Sub
Private Sub CargaReporteDetallado(ByVal numero As String)
Dim rs As ADODB.Recordset
Dim objRep As New ClsReporte
On Error GoTo mierror

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

    rs.Open "PP_LISTA_CONTABLEHABDES.LISTAR('" & numero & "')", vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF Then
        MsgBox "No hay datos para mostrar", vbExclamation, "Pago de pensiones"
        Exit Sub
    End If
    
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PP_Rpt_ContableDetHabDes.rpt"), ".RPT", ".TTX"), 1)
    
        
    If objRep.CargaReporte(strRpt & "", "PP_Rpt_ContableDetHabDes.rpt", "Informe Resumen Archivo Contable", rs, True, _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("NombreSistema", vgNombreSistema), _
                            ArrFormulas("NombreSubSistema", vgNombreSubSistema)) = False Then
                            
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
    End If

Exit Sub
mierror:
    MsgBox "No se pudo imprimir", vbExclamation, "Pago de Pensiones"
    
End Sub
Private Sub CargaReporte(ByVal numero As String)
Dim rs As ADODB.Recordset
Dim objRep As New ClsReporte
On Error GoTo mierror

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

    rs.Open "PP_LISTA_CONTABLEHABDES_CAB.LISTAR('" & numero & "')", vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF Then
        MsgBox "No hay datos para mostrar", vbExclamation, "Pago de pensiones"
        Exit Sub
    End If
    
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PP_Rpt_ContableResHabDes.rpt"), ".RPT", ".TTX"), 1)
    
        
    If objRep.CargaReporte(strRpt & "", "PP_Rpt_ContableResHabDes.rpt", "Informe Resumen Archivo Contable", rs, True, _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("NombreSistema", vgNombreSistema), _
                            ArrFormulas("NombreSubSistema", vgNombreSubSistema)) = False Then
                            
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
    End If

Exit Sub
mierror:
    MsgBox "No se pudo imprimir", vbExclamation, "Pago de Pensiones"
    
End Sub
Private Sub Cmd_ImpEstadistica_Click()
On Error GoTo Err_Imprimir

    'Imprime el Reporte de Resumen
    flImpresion

Exit Sub
Err_Imprimir:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub
Private Sub CargaReporteContableDetalleProdMon()
    Dim rs As ADODB.Recordset
    Dim objRep As New ClsReporte
    Dim cadena As String
    
    On Error GoTo mierror
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

    cadena = "select C.NUM_ARCHIVO,C.FEC_DESDE,C.FEC_HASTA,C.COD_USUARIOCREA,C.FEC_CREA,C.HOR_CREA," & _
            " d.COD_TIPREG , d.COD_TIPMOV,d.Cod_Moneda, d.COD_AFP, d.COD_TIPPENSION, d.MTO_PAGO" & _
            " from PP_TMAE_CONTABLEREGPAGO C INNER JOIN PP_TMAE_CONTABLEDETREGPAGO D ON C.NUM_ARCHIVO=D.NUM_ARCHIVO and C.COD_TIPREG=D.COD_TIPREG AND" & _
            " C.COD_TIPMOV=D.COD_TIPMOV AND C.COD_MONEDA=D.COD_MONEDA where c.NUM_ARCHIVO='" & Trim(Lbl_NumArchivo) & "'"
         
    rs.Open cadena, vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF Then
        MsgBox "No hay datos para mostrar", vbExclamation, "Pago de pensiones"
        Exit Sub
    End If
    
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PP_Rpt_ContableDetPagRegProdMon.rpt"), ".RPT", ".TTX"), 1)
    
    If objRep.CargaReporte(strRpt & "", "PP_Rpt_ContableDetPagRegProdMon.rpt", "Informe Detallado de Archivo Contable", rs, True, _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("NombreSistema", vgNombreSistema), _
                            ArrFormulas("NombreSubSistema", vgNombreSubSistema)) = False Then
                            
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
    End If
    
    Exit Sub
mierror:
    MsgBox "No se pudo cargar el reporte", vbInformation, "Pensiones"
    
End Sub
Private Sub CargaReporteContableDetalleProdMon_Provision()
    Dim rs As ADODB.Recordset
    Dim objRep As New ClsReporte
    Dim cadena As String
    
    On Error GoTo mierror
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

'    cadena = "select C.NUM_ARCHIVO,C.FEC_DESDE,C.FEC_HASTA,C.COD_USUARIOCREA,C.FEC_CREA,C.HOR_CREA," & _
'            " d.COD_TIPREG , d.COD_TIPMOV,d.Cod_Moneda, d.COD_AFP, d.COD_TIPPENSION, d.MTO_PAGO" & _
'            " from PP_TMAE_CONTABLE_PROVISION C INNER JOIN PP_TMAE_CONTABLEDET_PROVISION D ON C.NUM_ARCHIVO=D.NUM_ARCHIVO and C.COD_TIPREG=D.COD_TIPREG AND" & _
'            " C.COD_TIPMOV=D.COD_TIPMOV AND C.COD_MONEDA=D.COD_MONEDA where c.NUM_ARCHIVO='" & Trim(Lbl_NumArchivo) & "'"
    cadena = "select (CASE WHEN B.num_poliza IS NULL THEN 'N' ELSE 'S' END)AS ESTADO,X.NUM_ARCHIVO,X.FEC_DESDE,X.FEC_HASTA," & _
            " X.Cod_UsuarioCrea , X.Fec_Crea, X.Hor_Crea, A.COD_TIPREG, A.COD_TIPMOV, A.Cod_Moneda, A.COD_AFP, A.Cod_TipPension, A.MTO_PAGO, A.GLS_NOMBRE" & _
            " FROM PP_TMAE_CONTABLE_PROVISION X INNER JOIN PP_TMAE_CONTABLEDET_PROVISION A ON X.NUM_ARCHIVO = A.NUM_ARCHIVO" & _
            " And X.COD_TIPREG = A.COD_TIPREG And X.COD_TIPMOV = A.COD_TIPMOV And X.Cod_Moneda = A.Cod_Moneda" & _
            " left join PP_TMAE_CONTABLEDETREGPAGO B ON A.NUM_ARCHIVO = B.NUM_ARCHIVO And A.COD_TIPREG = B.COD_TIPREG And" & _
            " A.COD_TIPMOV = B.COD_TIPMOV And A.Cod_Moneda = B.Cod_Moneda And A.num_poliza = B.num_poliza" & _
            " where X.NUM_ARCHIVO='" & Trim(Lbl_NumArchivo) & "'" & _
            " GROUP BY X.NUM_ARCHIVO,X.FEC_DESDE,X.FEC_HASTA,X.COD_USUARIOCREA,X.FEC_CREA,X.HOR_CREA, A.COD_TIPREG , A.COD_TIPMOV," & _
            " A.Cod_Moneda, A.COD_AFP, A.COD_TIPPENSION, A.MTO_PAGO ,B.num_poliza,A.GLS_NOMBRE"
            
    rs.Open cadena, vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF Then
        MsgBox "No hay datos para mostrar", vbExclamation, "Pago de pensiones"
        Exit Sub
    End If
    
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PP_Rpt_ContableDetPagRegProdMonProv.rpt"), ".RPT", ".TTX"), 1)
    
    If objRep.CargaReporte(strRpt & "", "PP_Rpt_ContableDetPagRegProdMonProv.rpt", "Informe Detallado de Archivo Contable Provision", rs, True, _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("NombreSistema", vgNombreSistema), _
                            ArrFormulas("NombreSubSistema", vgNombreSubSistema)) = False Then
                            
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
    End If
    
    Exit Sub
mierror:
    MsgBox "No se pudo cargar el reporte", vbInformation, "Pensiones"
    
End Sub
Private Sub flImpProdMon()
Dim vlArchivo As String
Err.Clear
On Error GoTo Err_ImpProdMon
   
    Screen.MousePointer = 11
    
    If (Trim(Lbl_NumArchivo) = "") Then
        MsgBox "Debe seleccionar un Periodo a Imprimir.", vbInformation, "Falta Información"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    If (vgNomForm = "PagRec") Then
        Call CargaReporteContableDetalleProdMon
        Call CargaReporteContableDetalleProdMon_Provision
        Exit Sub
'        vlArchivo = strRpt & "PP_Rpt_ContableDetPagRegProdMon.rpt"   '\Reportes
'        vgQuery = "{PP_TMAE_CONTABLEDETREGPAGO.NUM_ARCHIVO} = " & Trim(Lbl_NumArchivo)
    End If
    
    If Not fgExiste(vlArchivo) Then     ', vbNormal
        MsgBox "Archivo de Reporte de Detalle de Archivo Contable no se encuentra en el directorio de la Aplicación.", 16, "Archivo no encontrado"
        Screen.MousePointer = 0
        Exit Sub
    End If
        
    Rpt_General.Reset
    Rpt_General.WindowState = crptMaximized
    Rpt_General.ReportFileName = vlArchivo
    Rpt_General.Connect = vgRutaDataBase
    Rpt_General.Destination = crptToWindow
    Rpt_General.SelectionFormula = ""
    Rpt_General.SelectionFormula = vgQuery
    
    Rpt_General.Formulas(0) = ""
    Rpt_General.Formulas(1) = ""
    Rpt_General.Formulas(2) = ""
    Rpt_General.Formulas(3) = ""
    
    Rpt_General.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
    Rpt_General.Formulas(1) = "NombreSistema = '" & vgNombreSistema & "'"
    Rpt_General.Formulas(2) = "NombreSubSistema = '" & vgNombreSubSistema & "'"
    
    Rpt_General.WindowTitle = "Informe Resumen Archivo Contable por Producto - Moneda"
    Rpt_General.Action = 1
    
    Screen.MousePointer = 0
   
Exit Sub
Err_ImpProdMon:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " " & "Numero de Error : " & Err.Number & Err.HelpFile & Err.HelpContext, vbExclamation, "Mensaje de error ..."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub Cmd_Imprimir_Click()
On Error GoTo Err_Imprimir

    'Imprime el Reporte de Resumen
    If (Opt_Resumen) Then
        flImpresion
    End If
    'Imprime el Reporte de Detalle
    If (Opt_Detalle) Then
        flImpDetalle
    End If
    'Imprime el Reporte de Detalle AFP - Moneda
    If (Opt_DetAfpMon) Then
        flImpAfpMon
    End If
    'Imprime el Reporte de Detalle Producto - Moneda
    If (Opt_DetProdMon) Then
        flImpProdMon
    End If
    'Imprime el Reporte de Detalle Montos del Beneficiario
    If (Opt_DetMtos) Then
        flImpDetMontos
    End If
    'Imprime lo pendiente acumulado hasta la fecha
    If (Opt_DetPendAcumulado) Then
        Call CargaReporteContablePendienteAcumulada
    End If
    
Exit Sub
Err_Imprimir:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Limpiar_Click()
On Error GoTo Err_Limpiar
    
    Lbl_FecCierre = Format(Now, "dd/mm/yyyy Hh:Nn:Ss AMPM")
    Txt_Desde = ""
    Txt_Hasta = ""
    LblArchivo.Caption = ""
    Lbl_NumArchivo.Caption = ""
    
    If (vgNomForm = "PagRec") Then
        Call flLmpGrillaPagReg
        Call flActGrillaPagReg
    ElseIf (vgNomForm = "GtoSep") Then
        Call flLmpGrilla
        Call flActGrillaGtoSep
    Else
        Call flLmpGrilla
        Call flActGrillaPerGar
    End If
    
    Opt_Resumen.Value = True
    
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

Private Sub CmdContable_Click()
Dim ilargo As Long
Dim iTexto As Long
Dim fecdesde As String
Dim fechasta As String
On Error GoTo Err_Carga
    
    If Not IsDate(Txt_Desde) Then
        MsgBox "Debe Ingresar la Fecha Desde.", vbInformation, "Falta Información"
        Txt_Desde.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(Txt_Hasta) Then
        MsgBox "Debe Ingresar la Fecha Hasta.", vbInformation, "Falta Información"
        Txt_Hasta.SetFocus
        Exit Sub
    End If
    
    If CDate(Txt_Desde) > CDate(Txt_Hasta) Then
       MsgBox "La Fecha Desde debe ser menor o igual a la Fecha Hasta.", vbInformation, "Error de Datos"
       Txt_Hasta.SetFocus
       Exit Sub
    End If
    
    fecdesde = Format(CDate(Trim(Txt_Desde)), "yyyymmdd")
    fechasta = Format(CDate(Trim(Txt_Hasta)), "yyyymmdd")
    
    'Selección del Archivo del cual se generará el archivo contable

    vlArchivoCont = ""
    ComDialogo.CancelError = True
    ComDialogo.FileName = "SEA_" & vgNomForm & "_" & fecdesde & "_" & fechasta & ".txt" '17/06/2008
    ComDialogo.DialogTitle = "Archivo Contable"
    ComDialogo.Filter = "*.txt"
    ComDialogo.FilterIndex = 1
    ComDialogo.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    ComDialogo.ShowSave
    vlArchivoCont = ComDialogo.FileName
    LblArchivo.Caption = vlArchivoCont
    If (Len(vlArchivoCont) > 65) Then
        While Len(LblArchivo) > 65
            ilargo = InStr(1, LblArchivo, "\")
            LblArchivo = Mid(LblArchivo, ilargo + 1, Len(LblArchivo))
        Wend
        LblArchivo.Caption = "\\" & LblArchivo
    End If
    If vlArchivoCont = "" Then
        Exit Sub
    End If
    
    'Selección del Archivo del cual se generará el archivo contable
    
    vlArchivoContProv = ""
    ComDialogo.CancelError = True
    ComDialogo.FileName = "SEA_PROVISION" & vgNomForm & "_" & fecdesde & "_" & fechasta & ".txt" '17/06/2008
    ComDialogo.DialogTitle = "Archivo Contable"
    ComDialogo.Filter = "*.txt"
    ComDialogo.FilterIndex = 1
    ComDialogo.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    'ComDialogo.ShowSave
    Dim x As Integer
    x = InStr(1, vlArchivoCont, "SEA_")
    vlArchivoContProv = Mid(vlArchivoCont, 1, x - 1)

    vlArchivoContProv = vlArchivoContProv & ComDialogo.FileName
    LblArchivo.Caption = vlArchivoContProv
'    vlArchivoContProv = ComDialogo.FileName
'    LblArchivo.Caption = vlArchivoContProv
    If (Len(vlArchivoContProv) > 65) Then
        While Len(LblArchivo) > 65
            ilargo = InStr(1, LblArchivo, "\")
            LblArchivo = Mid(LblArchivo, ilargo + 1, Len(LblArchivo))
        Wend
        LblArchivo.Caption = "\\" & LblArchivo
    End If
    If vlArchivoContProv = "" Then
        Exit Sub
    End If
    
Exit Sub
Err_Carga:
    If (Err.Number = 32755) Then
        Exit Sub
    End If
    Screen.MousePointer = 0
    MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
End Sub

Private Sub Form_Load()
On Error GoTo Err_Cargar
  
    Me.Top = 0
    Me.Left = 0
    Lbl_FecCierre = Format(Now, "dd/mm/yyyy Hh:Nn:Ss AMPM")

    Frame2.Visible = True
    If (vgNomForm = "PagRec") Then
        Call flMostrarOption
        Call flLmpGrillaPagReg
        Call flActGrillaPagReg
        Me.Caption = "Generación de Archivo Contable de Pagos Recurrentes."
    ElseIf (vgNomForm = "GtoSep") Then
        Call flOcultarOption
        Call flLmpGrilla
        Call flActGrillaGtoSep
        Frame2.Visible = False
        Me.Caption = "Generación de Archivo Contable de Gastos de Sepelio."
    ElseIf (vgNomForm = "HabDes") Then
        Call flOcultarOption
        Call flLmpGrilla
        Call flActGrillaHabDes
        Me.Caption = "Generación de Archivo Contable de Haberes y Descuentos."
    Else
        Call flOcultarOption
        Call flLmpGrilla
        Call flActGrillaPerGar
        Me.Caption = "Generación de Archivo Contable de Periodo Garantizado."
    End If
    
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Msf_Periodos_Click()
On Error GoTo Err_Periodo

    If (Msf_Periodos.Text = "") Or (Msf_Periodos.row = 0) Then
        Exit Sub
    End If
    
    vgI = Msf_Periodos.row
    
    Txt_Desde = Msf_Periodos.TextMatrix(vgI, 1)
    Txt_Hasta = Msf_Periodos.TextMatrix(vgI, 2)
    
    If (vgNomForm = "PagRec") Then
        Lbl_NumArchivo.Caption = Msf_Periodos.TextMatrix(vgI, 9)
    Else
        Lbl_NumArchivo.Caption = Msf_Periodos.TextMatrix(vgI, 7)
    End If
        
Exit Sub
Err_Periodo:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

 '----- RESCATA EL NUMERO DE ENTRADA CORRESPONDIENTE AL ARCHIVO DE PRIMA UNICA -----
Private Function flNumArchivoPagReg() As Integer
    vgQuery = ""
    vgQuery = "SELECT NUM_ARCHIVO FROM PP_TMAE_CONTABLEREGPAGO "
    vgQuery = vgQuery & " ORDER BY NUM_ARCHIVO DESC"
    Set vgRs = vgConexionBD.Execute(vgQuery)
    If Not (vgRs.EOF) Then
        flNumArchivoPagReg = CInt(vgRs!num_archivo) + 1
    Else
        flNumArchivoPagReg = 1
    End If
End Function
Private Function flNumArchivoPagRegProvision() As Integer
    vgQuery = ""
    vgQuery = "SELECT NUM_ARCHIVO FROM PP_TMAE_CONTABLE_PROVISION "
    vgQuery = vgQuery & " ORDER BY NUM_ARCHIVO DESC"
    Set vgRs = vgConexionBD.Execute(vgQuery)
    If Not (vgRs.EOF) Then
        flNumArchivoPagRegProvision = CInt(vgRs!num_archivo) + 1
    Else
        flNumArchivoPagRegProvision = 1
    End If
End Function
 '----- RESCATA EL NUMERO DE ENTRADA CORRESPONDIENTE AL ARCHIVO DE GASTOS DE SEPELIO -----
Private Function flNumArchivoGtoSep() As Integer
    vgQuery = ""
    vgQuery = "SELECT NUM_ARCHIVO FROM PP_TMAE_CONTABLEGTOSEP "
    vgQuery = vgQuery & " ORDER BY NUM_ARCHIVO DESC"
    Set vgRs = vgConexionBD.Execute(vgQuery)
    If Not (vgRs.EOF) Then
        flNumArchivoGtoSep = CInt(vgRs!num_archivo) + 1
    Else
        flNumArchivoGtoSep = 1
    End If
End Function
Private Function flNumArchivoHabDes() As Integer
    vgQuery = ""
    vgQuery = "SELECT NUM_ARCHIVO FROM PP_TMAE_CONTABLEHABDES "
    vgQuery = vgQuery & " ORDER BY NUM_ARCHIVO DESC"
    Set vgRs = vgConexionBD.Execute(vgQuery)
    If Not (vgRs.EOF) Then
        flNumArchivoHabDes = CInt(vgRs!num_archivo) + 1
    Else
        flNumArchivoHabDes = 1
    End If
End Function

 '----- RESCATA EL NUMERO DE ENTRADA CORRESPONDIENTE AL ARCHIVO DE PERIODO GARANTIZADO -----
Private Function flNumArchivoPerGar() As Integer
    vgQuery = ""
    vgQuery = "SELECT NUM_ARCHIVO FROM PP_TMAE_CONTABLEPERGAR "
    vgQuery = vgQuery & " ORDER BY NUM_ARCHIVO DESC"
    Set vgRs = vgConexionBD.Execute(vgQuery)
    If Not (vgRs.EOF) Then
        flNumArchivoPerGar = CInt(vgRs!num_archivo) + 1
    Else
        flNumArchivoPerGar = 1
    End If
End Function

Private Function flExportarPagReg(iFecDesde As String, iFecHasta As String) As Boolean
Dim vlNumPol As String, vlFecPago As String
Dim vlMtoPen As Double, vlMtoSal As Double, vlMtoRet As Double, vlMtoPenTot As Double
Dim vlCodPen As String, vlAfp As String
Dim vlNumOrd As Integer
Dim vlCodViaPag As String
Dim vgRs1 As ADODB.Recordset
Dim identit As String
Dim nombreTit As String

On Error GoTo Errores

flExportarPagReg = False

    vlArchivo = vlArchivoCont 'LblArchivo
    
    Screen.MousePointer = 11
    
    Open vlArchivo For Output As #2

    Me.Refresh
    vlOpen = True
        
    Call flInicializaVar
        
    'Obtiene el nº de archivo a crear
    vlNumArchivo = flNumArchivoPagReg
    vlFecCrea = Format(Date, "yyyymmdd")
    vlHorCrea = Format(Time, "hhmmss")
    
    vlSql = "SELECT COD_ELEMENTO FROM MA_TPAR_TABCOD "
    vlSql = vlSql & "WHERE COD_TABLA='" & vgCodTabla_TipMon & "' "
    vlSql = vlSql & "AND (COD_SISTEMA IS NULL OR COD_SISTEMA<>'PP') "
    vlSql = vlSql & "ORDER BY COD_ELEMENTO "
    Set vgRs1 = vgConexionBD.Execute(vlSql)
    While Not (vgRs1.EOF)
        vlMoneda = Trim(vgRs1!COD_ELEMENTO)
        
        Call flEstadisticaPagReg(clTipRegPR, clTipMovSinLPR)
        Call flEstadisticaPagReg(clTipRegPR, clTipMovSinSPR)
        Call flEstadisticaPagReg(clTipRegPR, clTipMovSinRPR)

        vlNumCasosPRPension = 0
        vlNumCasosPRSalud = 0
        vlNumCasosPRRetencion = 0
        vlMtoPRPension = 0
        vlMtoPRSalud = 0
        vlMtoPRRetencion = 0
        
        Dim a As Integer
        
'        vgSql = "SELECT l.num_poliza,l.cod_viapago,l.fec_pago,l.cod_moneda,l.cod_tippension,"
'        vgSql = vgSql & "b.cod_tipoidenben,b.num_idenben,b.gls_nomben,b.gls_nomsegben,"
'        vgSql = vgSql & "b.gls_patben,b.gls_matben,p.cod_afp,l.cod_inssalud,"
'        vgSql = vgSql & "mto_liqpagar,l.mto_plansalud,l.num_orden,l.num_perpago,"
'        vgSql = vgSql & "l.num_idenreceptor,l.cod_tipoidenreceptor,l.cod_tipreceptor"
'        vgSql = vgSql & ",l.mto_haber as mto_pensiontot, b.NUM_CTABCO, b.NUM_CUENTA_CCI, b.cod_banco, m.gls_elemento gls_banco " '01/12/2007
'        vgSql = vgSql & "FROM pp_tmae_liqpagopendef l, pp_tmae_poliza p, pp_tmae_ben b, ma_tpar_tabcod m "
'        vgSql = vgSql & "WHERE l.fec_pago between '" & iFecDesde & "' and '" & iFecHasta & "' "
'        vgSql = vgSql & "and p.cod_moneda ='" & vlMoneda & "' "
'        vgSql = vgSql & "and l.num_poliza=p.num_poliza "
'        vgSql = vgSql & "and l.num_endoso=p.num_endoso "
'        vgSql = vgSql & "and l.num_poliza=b.num_poliza "
'        vgSql = vgSql & "and l.num_endoso=b.num_endoso "
'        vgSql = vgSql & "and l.num_orden=b.num_orden and b.cod_banco=m.cod_elemento and m.cod_tabla='BCO' "
'        vgSql = vgSql & "AND COD_TIPOPAGO='" & clRetPR & "' "
'        vgSql = vgSql & "AND COD_TIPRECEPTOR<>'" & clRetPR & "' "
'
'        If chkPagoD.Value = 1 Then
'            vgSql = vgSql & "AND p.COD_TIPPENSION IN ('04','05','09','10') "
'        End If
'
'        If chkPagoA.Value = 1 Then
'            vgSql = vgSql & "AND p.COD_TIPPENSION NOT IN ('04','05','09','10') "
'        End If
'
'        vgSql = vgSql & "group by  l.num_poliza,l.cod_viapago,l.fec_pago,l.cod_moneda,"
'        vgSql = vgSql & "l.cod_tippension,b.cod_tipoidenben,b.num_idenben,b.gls_nomben,"
'        vgSql = vgSql & "b.gls_nomsegben,b.gls_patben,b.gls_matben,p.cod_afp,"
'        vgSql = vgSql & "l.cod_inssalud,mto_liqpagar,l.mto_plansalud,l.num_orden,l.num_perpago,"
'        vgSql = vgSql & "l.num_idenreceptor,l.cod_tipoidenreceptor,l.cod_tipreceptor,l.mto_haber,b.NUM_CTABCO, b.NUM_CUENTA_CCI, b.cod_banco, m.gls_elemento "
'        vgSql = vgSql & "ORDER BY l.num_poliza,l.fec_pago,l.num_orden "
        
        
        
        
        vgSql = "SELECT l.num_poliza,l.cod_viapago,l.fec_pago,l.cod_moneda,l.cod_tippension,b.cod_tipoidenben,b.num_idenben,b.gls_nomben,b.gls_nomsegben,b.gls_patben,b.gls_matben,p.cod_afp,l.cod_inssalud,mto_liqpagar,l.mto_plansalud,"
        vgSql = vgSql & " l.num_orden,l.num_perpago,l.num_idenreceptor,l.cod_tipoidenreceptor,l.cod_tipreceptor,l.mto_haber as mto_pensiontot, b.NUM_CTABCO, b.NUM_CUENTA_CCI, nvl(b.cod_banco, '00') cod_banco, m.gls_elemento gls_banco, b.cod_tipcta"
        vgSql = vgSql & " FROM pp_tmae_liqpagopendef l"
        vgSql = vgSql & " join pp_tmae_ben b on l.num_poliza=b.num_poliza and l.num_orden=b.num_orden"
        vgSql = vgSql & " join pp_tmae_poliza p on b.num_poliza=p.num_poliza and b.num_endoso=p.num_endoso"
        vgSql = vgSql & " left join ma_tpar_tabcod m on b.cod_banco=m.cod_elemento and m.cod_tabla='BCO'"
        vgSql = vgSql & " WHERE l.fec_pago between '" & iFecDesde & "' and '" & iFecHasta & "'"
        vgSql = vgSql & " and p.cod_moneda ='" & vlMoneda & "'"
        vgSql = vgSql & " AND COD_TIPOPAGO='" & clRetPR & "' "
        vgSql = vgSql & " AND COD_TIPRECEPTOR<>'" & clRetPR & "'"
        If chkPagoD.Value = 1 Then
            vgSql = vgSql & "AND p.COD_TIPPENSION IN ('04','05','09','10') "
        End If
        
        If chkPagoA.Value = 1 Then
            vgSql = vgSql & " AND p.COD_TIPPENSION NOT IN ('04','05','09','10') "
        End If
        vgSql = vgSql & " AND p.NUM_ENDOSO=(SELECT MAX(NUM_ENDOSO) FROM PP_TMAE_POLIZA WHERE NUM_POLIZA=L.NUM_POLIZA) "
        'vgSql = vgSql & " AND L.NUM_POLIZA in (26,36,246,280,285,335,358,402,640,769,815,852,1006,1018,1087,1255,1300,1536,1544,1639,1731,1758,1840,1878,2048,2092,2411,2655,2875,2954,3117,3118,3142,3251,3397,3861,3865)"
        'vgSql = vgSql & " AND L.NUM_POLIZA=6792"
        vgSql = vgSql & " group by  l.num_poliza,l.cod_viapago,l.fec_pago,l.cod_moneda,l.cod_tippension,b.cod_tipoidenben,b.num_idenben,b.gls_nomben,b.gls_nomsegben,b.gls_patben,b.gls_matben,"
        vgSql = vgSql & " p.cod_afp,l.cod_inssalud,mto_liqpagar,l.mto_plansalud,l.num_orden,l.num_perpago,l.num_idenreceptor,l.cod_tipoidenreceptor,l.cod_tipreceptor,l.mto_haber,b.NUM_CTABCO,"
        vgSql = vgSql & " b.NUM_CUENTA_CCI , b.Cod_Banco, m.gls_elemento, b.cod_tipcta"
        vgSql = vgSql & " ORDER BY l.num_poliza,l.fec_pago,l.num_orden"
        Set vgRs = vgConexionBD.Execute(vgSql)
        While Not (vgRs.EOF)
        
            vlNumPol = Trim(vgRs!num_poliza)
            vlNumOrd = CInt(Trim(vgRs!Num_Orden))
            vlFecPago = Trim(vgRs!Fec_Pago)
            
'            If vlNumPol = "0000000640" Then
'                a = 1
'            End If
'
            
            
            '****************** Movimiento de Pensión ****************************
            vlVar1 = Format(Trim(clTipRegPR), "00000")
            vlVar2 = Format(Trim(vgRs!num_poliza), "0000000000")
            'vlVar3 = flObtieneViaPago(Trim(vgRs!Cod_ViaPago))
            vlVar3 = Trim(vgRs!Cod_ViaPago)
            vlVar3 = Format(Trim(vlVar3), "00000")
            vlVar3_Res = vlVar3
            If Trim(vgRs!Fec_Pago) <> "" Then
                vlVar4 = DateSerial(Mid(vgRs!Fec_Pago, 1, 4), Mid(vgRs!Fec_Pago, 5, 2), Mid(vgRs!Fec_Pago, 7, 2))
                vlVar7 = Mid(vgRs!Fec_Pago, 1, 6)
            Else
                vlVar4 = Space(10)
                vlVar7 = Space(6)
            End If
            If (Len(Trim(vgRs!Cod_Moneda)) <= 5) Then
                vlVar8 = Trim(vgRs!Cod_Moneda) & Space(5 - Len(Trim(vgRs!Cod_Moneda)))
            Else
                vlVar8 = Mid(Trim(vgRs!Cod_Moneda), 1, 5)
            End If
            
            vlVar9 = Format(Trim(vgRs!Cod_TipPension), "00000")
            vlCodPen = Trim(vgRs!Cod_TipPension)
            
            Select Case vgRs!Cod_TipPension
             Case "04"
                vlVar10 = Format(Trim("76"), "00000")
             Case "05"
                vlVar10 = Format(Trim("76"), "00000")
             Case "06"
                vlVar10 = Format(Trim("94"), "00000")
             Case "07"
                vlVar10 = Format(Trim("94"), "00000")
             Case "08"
                vlVar10 = Format(Trim("95"), "00000")
            End Select
            'vlVar10 = Format(Trim(clRamoContPR), "00000")
            
            If (Len(Trim(vgRs!Num_IdenBen)) <= 12) Then
                vlVar11 = Format(Trim(vgRs!Cod_TipoIdenBen), "00") & (Trim(vgRs!Num_IdenBen) & Space(12 - Len(Trim(vgRs!Num_IdenBen))))
            Else
                vlVar11 = Format(Trim(vgRs!Cod_TipoIdenBen), "00") & Mid(Trim(vgRs!Num_IdenBen), 1, 12)
            End If
            identit = vlVar11
            vlVar11_Res = vlVar11
            vlVar12 = fgFormarNombreCompleto(IIf(IsNull(vgRs!Gls_NomBen), "", Trim(vgRs!Gls_NomBen)), IIf(IsNull(vgRs!Gls_NomSegBen), "", Trim(vgRs!Gls_NomSegBen)), IIf(IsNull(vgRs!Gls_PatBen), "", Trim(vgRs!Gls_PatBen)), IIf(IsNull(vgRs!Gls_MatBen), "", Trim(vgRs!Gls_MatBen)))
            If (Len(Trim(vlVar12)) <= 60) Then
                vlVar12 = vlVar12 & Space(60 - Len(Trim(vlVar12)))
            Else
                vlVar12 = Mid(vlVar12, 1, 60)
            End If
            nombreTit = vlVar12
            vlVar12_Res = vlVar12
            vlVar13 = Format(Trim(clFrecPagPR), "00000")
            vlVar14 = Mid(vlFecPago, 3, 4) & Mid(vlNumPol, 6, 5) & "-" & Format(vlNumOrd, "00")
            If (Len(Trim(vlVar14)) <= 14) Then
                vlVar14 = Trim(vlVar14) & Space(14 - Len(Trim(vlVar14)))
            Else
                vlVar14 = Mid(Trim(vlVar14), 1, 14)
            End If
            vlVar14_Res = vlVar14
            
            If vgRs!Cod_ViaPago = "02" Then
                If Len(vgRs!num_ctabco) <> 0 Then
                    vlVar17 = vgRs!num_ctabco
                    vlVar17 = vlVar17 & Space(60 - Len(Trim(vlVar17)))
                Else
                    vlVar17 = Space(60)
                End If
            Else
                If Len(vgRs!NUM_CUENTA_CCI) <> 0 Then
                    vlVar17 = vgRs!NUM_CUENTA_CCI
                    vlVar17 = vlVar17 & Space(60 - Len(Trim(vlVar17)))
                Else
                    vlVar17 = Space(60)
                End If
            End If
            
            vlVar19 = Format(Trim(vgRs!cod_afp), "00000")
            vlVar19_Res = vlVar19
            If Len(vgRs!gls_banco) <> 0 Then
                vlVar20 = Trim(vgRs!gls_banco)
                vlVar20 = vlVar20 & Space(60 - Len(Trim(vlVar20)))
            Else
                vlVar20 = Space(60)
            End If
            vlAfp = Trim(vgRs!cod_afp)
            If Len(vgRs!Cod_Banco) <> 0 Then
                vlVar22 = Trim(vgRs!Cod_Moneda) & Format(Trim(vgRs!Cod_Banco), "000")
            Else
                vlVar22 = String(5, "0")
            End If
            
            vlVar23 = Format(Trim(clReaPR), "0000000000")
            vlVar25 = Format(Trim(vgRs!Cod_TipPension), "00000")
            If Len(Trim(vgRs!cod_tipcta)) <> 0 Then
            vlVar26 = vgRs!cod_tipcta
                vlVar26 = vlVar26 & Space(6 - Len(Trim(vlVar26)))
            Else
                vlVar26 = Space(6)
            End If
            
            vlVar29 = Format(Trim(vgRs!Cod_TipPension), "00000")
            vlVar37 = Format(Trim(vgRs!Cod_TipPension), "00000")
            vlVar61 = Format(Trim(clTipMovSinLPR), "00000")
            vlVar62 = IIf(IsNull(vgRs!Mto_LiqPagar), 0, Format(vgRs!Mto_LiqPagar, "#0.00"))
            vlVar62 = flFormatNum18_2(vlVar62)
            vlVar63_Res = vlVar62
            vlVar67 = flObtieneTipPer(Trim(vgRs!Cod_ViaPago))
            vlVar67_Res = vlVar67
            'vlVar67 = Trim(clTipPerJurPR) & Space(1 - Len(Trim(clTipPerJurPR)))
        
        
'            vlNumPol = Trim(vgRs!num_poliza)
'            vlNumOrd = CInt(Trim(vgRs!Num_Orden))
'            vlFecPago = Trim(vgRs!Fec_Pago)
'
'            If vlNumPol = "0000000640" Then
'                a = 1
'            End If
'
'
'            '****************** Movimiento de Pensión ****************************
'            vlVar1 = Format(Trim(clTipRegPR), "00000")
'            vlVar2 = Format(Trim(vgRs!num_poliza), "0000000000")
'            vlVar3 = flObtieneViaPago(Trim(vgRs!Cod_ViaPago))
'            vlVar3 = Format(Trim(vlVar3), "00000")
'            vlVar3_Res = vlVar3
'            If Trim(vgRs!Fec_Pago) <> "" Then
'                vlVar4 = DateSerial(Mid(vgRs!Fec_Pago, 1, 4), Mid(vgRs!Fec_Pago, 5, 2), Mid(vgRs!Fec_Pago, 7, 2))
'                vlVar7 = Mid(vgRs!Fec_Pago, 1, 6)
'            Else
'                vlVar4 = Space(10)
'                vlVar7 = Space(6)
'            End If
'            If (Len(Trim(vgRs!Cod_Moneda)) <= 5) Then
'                vlVar8 = Trim(vgRs!Cod_Moneda) & Space(5 - Len(Trim(vgRs!Cod_Moneda)))
'            Else
'                vlVar8 = Mid(Trim(vgRs!Cod_Moneda), 1, 5)
'            End If
'
'            vlVar9 = Format(Trim(vgRs!Cod_TipPension), "00000")
'            vlCodPen = Trim(vgRs!Cod_TipPension)
'
'            Select Case vgRs!Cod_TipPension
'             Case "04"
'                vlVar10 = Format(Trim("76"), "00000")
'             Case "05"
'                vlVar10 = Format(Trim("76"), "00000")
'             Case "06"
'                vlVar10 = Format(Trim("94"), "00000")
'             Case "07"
'                vlVar10 = Format(Trim("94"), "00000")
'             Case "08"
'                vlVar10 = Format(Trim("95"), "00000")
'            End Select
'            'vlVar10 = Format(Trim(clRamoContPR), "00000")
'
'            If (Len(Trim(vgRs!Num_IdenBen)) <= 12) Then
'                vlVar11 = Format(Trim(vgRs!Cod_TipoIdenBen), "00") & (Trim(vgRs!Num_IdenBen) & Space(12 - Len(Trim(vgRs!Num_IdenBen))))
'            Else
'                vlVar11 = Format(Trim(vgRs!Cod_TipoIdenBen), "00") & Mid(Trim(vgRs!Num_IdenBen), 1, 12)
'            End If
'            identit = vlVar11
'            vlVar11_Res = vlVar11
'            vlVar12 = fgFormarNombreCompleto(IIf(IsNull(vgRs!Gls_NomBen), "", Trim(vgRs!Gls_NomBen)), IIf(IsNull(vgRs!Gls_NomSegBen), "", Trim(vgRs!Gls_NomSegBen)), IIf(IsNull(vgRs!Gls_PatBen), "", Trim(vgRs!Gls_PatBen)), IIf(IsNull(vgRs!Gls_MatBen), "", Trim(vgRs!Gls_MatBen)))
'            If (Len(Trim(vlVar12)) <= 60) Then
'                vlVar12 = vlVar12 & Space(60 - Len(Trim(vlVar12)))
'            Else
'                vlVar12 = Mid(vlVar12, 1, 60)
'            End If
'            nombreTit = vlVar12
'            vlVar12_Res = vlVar12
'            vlVar13 = Format(Trim(clFrecPagPR), "00000")
'            vlVar14 = Mid(vlFecPago, 3, 4) & Mid(vlNumPol, 6, 5) & "-" & Format(vlNumOrd, "00")
'            If (Len(Trim(vlVar14)) <= 14) Then
'                vlVar14 = Trim(vlVar14) & Space(14 - Len(Trim(vlVar14)))
'            Else
'                vlVar14 = Mid(Trim(vlVar14), 1, 14)
'            End If
'            vlVar14_Res = vlVar14
'            vlVar14_Res = vlVar14
'
''            If vgRs!Cod_ViaPago = "02" Then
''                If Len(vgRs!num_ctabco) <> 0 Then
''                    vlVar17 = vgRs!num_ctabco
''                    vlVar17 = vlVar17 & Space(60 - Len(Trim(vlVar17)))
''                Else
''                    vlVar17 = Space(60)
''                End If
''            Else
''                If Len(vgRs!NUM_CUENTA_CCI) <> 0 Then
''                    vlVar17 = vgRs!NUM_CUENTA_CCI
''                    vlVar17 = vlVar17 & Space(60 - Len(Trim(vlVar17)))
''                Else
''                    vlVar17 = Space(60)
''                End If
''            End If
'
'            vlVar19_Res = vlVar19
'            vlAfp = Trim(vgRs!cod_afp)
'            'vlVar22 = Format(Trim(vgRs!Cod_Banco), "00000")
'            vlVar23 = Format(Trim(clReaPR), "0000000000")
'            vlVar25 = Format(Trim(vgRs!Cod_TipPension), "00000")
'            vlVar29 = Format(Trim(vgRs!Cod_TipPension), "00000")
'            vlVar37 = Format(Trim(vgRs!Cod_TipPension), "00000")
'            vlVar61 = Format(Trim(clTipMovSinLPR), "00000")
'            vlVar62 = IIf(IsNull(vgRs!Mto_LiqPagar), 0, Format(vgRs!Mto_LiqPagar, "#0.00"))
'            vlVar62 = flFormatNum18_2(vlVar62)
'            vlVar63_Res = vlVar62
'            vlVar67 = flObtieneTipPer(Trim(vgRs!Cod_ViaPago))
'            vlVar67_Res = vlVar67
'            'vlVar67 = Trim(clTipPerJurPR) & Space(1 - Len(Trim(clTipPerJurPR)))
'            'vlVar68 = Format(Trim(vgRs!num_ctabco), String(15, "0"))
'            'vlVar69 = Format(Trim(vgRs!NUM_CUENTA_CCI), String(50, "0"))
            
    
            'Imprime la linea 38 Pago de Pensiones - Recurrentes Liquidos
            vlLinea = (vlVar1) & (vlVar2) & (vlVar3) & (vlVar4) & (vlVar5) & _
                      (vlVar6) & (vlVar7) & (vlVar8) & (vlVar9) & (vlVar10) & _
                      (vlVar11) & (vlVar12) & (vlVar13) & (vlVar14) & (vlVar15) & _
                      (vlVar16) & (vlVar17) & (vlVar18) & (vlVar19) & (vlVar20) & _
                      (vlVar21) & (vlVar22) & (vlVar23) & (vlVar24) & (vlVar25) & _
                      (vlVar26) & (vlVar27) & (vlVar28) & (vlVar29) & (vlVar30) & _
                      (vlVar31) & (vlVar32) & (vlVar33) & (vlVar34) & (vlVar35) & _
                      (vlVar36) & (vlVar37) & (vlVar38) & (vlVar39) & (vlVar40) & _
                      (vlVar41) & (vlVar42) & (vlVar43) & (vlVar44) & (vlVar45) & _
                      (vlVar47) & (vlVar48) & (vlVar49) & (vlVar50) & _
                      (vlVar51) & (vlVar52) & (vlVar53) & (vlVar54) & (vlVar55) & _
                      (vlVar56) & (vlVar57) & (vlVar58) & (vlVar59) & (vlVar60) & _
                      (vlVar61) & (vlVar62) & (vlVar63) & (vlVar64) & (vlVar65) & _
                      (vlVar66) & (vlVar67) '& (vlVar68) & (vlVar69)
            
            vlLinea = Replace(vlLinea, ",", ".")
            Print #2, vlLinea
            
            'Guarda el detalle de Pensión de la póliza informada
            vlMtoPen = IIf(IsNull(vgRs!Mto_LiqPagar), 0, Format(vgRs!Mto_LiqPagar, "#0.00"))
            Call flGrabaDetPagReg(clTipRegPR, clTipMovSinLPR, vlVar2, vlNumOrd, vlFecPago, vlCodPen, vlAfp, vlVar12, vlMtoPen)

            'Contador de Pagos Recurrentes - Pensión
            vlNumCasosPRPension = vlNumCasosPRPension + 1
            vlMtoPRPension = vlMtoPRPension + vlMtoPen

            
            '****************** Movimiento de Salud ****************************
            vlMtoSal = 0
            vgQuery = "SELECT mto_conhabdes as mto_salud FROM PP_TMAE_PAGOPENDEF "
            vgQuery = vgQuery & "WHERE cod_conhabdes='24' "
            vgQuery = vgQuery & "and num_perpago ='" & Trim(vgRs!Num_PerPago) & "' "
            vgQuery = vgQuery & "and num_poliza='" & Trim(vgRs!num_poliza) & "' "
            vgQuery = vgQuery & "and num_orden=" & vlNumOrd & " "
            vgQuery = vgQuery & "and trim(num_idenreceptor) ='" & Trim(vgRs!Num_IdenReceptor) & "' "
            'mv 201906
            vgQuery = vgQuery & "and cod_tipoidenreceptor =" & CInt(Trim(vgRs!Cod_TipoIdenReceptor)) & " "
            vgQuery = vgQuery & "and cod_tipreceptor ='" & Trim(vgRs!Cod_TipReceptor) & "' "
            Set vgRs4 = vgConexionBD.Execute(vgQuery)
            If Not (vgRs4.EOF) Then
                
                vlVar14 = Mid(vlFecPago, 3, 4) & Mid(vlNumPol, 6, 5) & "-" & Format(vlNumOrd, "00")
                If (Len(Trim(vlVar14)) <= 14) Then
                    vlVar14 = Trim(vlVar14) & Space(14 - Len(Trim(vlVar14)))
                Else
                    vlVar14 = Mid(Trim(vlVar14), 1, 14)
                End If
                vlVar19 = flObtieneCodSalud(Trim(vgRs!Cod_InsSalud))
                'vlVar7 = Space(6)
                'vlVar19 = Trim(vgRs!Cod_InsSalud) & Space(5 - Len(Trim(vgRs!Cod_InsSalud)))
                vlVar19 = Format(Trim(vlVar19), "00000")
                vlVar29 = Format(Trim(vgRs!Cod_TipPension), "00000")
                vlVar37 = Format(Trim(vgRs!Cod_TipPension), "00000")
                vlVar61 = Format(Trim(clTipMovSinSPR), "00000")
                vlVar62 = IIf(IsNull(vgRs4!mto_salud), 0, Format(vgRs4!mto_salud, "#0.00"))
                vlVar62 = flFormatNum18_2(vlVar62)
'                vlVar64_Res = IIf(IsNull(vgRs4!mto_salud), 0, Format(vgRs4!mto_salud, "#0.00"))
'                vlVar64_Res = flFormatNum5_2(vlVar64_Res)
                vlVar66_Res = IIf(IsNull(vgRs4!mto_salud), 0, Format(vgRs4!mto_salud, "#0.00"))
                vlVar66_Res = flFormatNum18_2(vlVar66_Res)
                'vlVar67 = Trim(clTipPerJurPR) & Space(5 - Len(Trim(clTipPerJurPR)))
                vlVar67 = flObtieneTipPer(Trim(vgRs!Cod_InsSalud))
                
                'Imprime la linea 39 Pago de Pensiones - Recurrentes Salud
                vlLinea = (vlVar1) & (vlVar2) & (vlVar3) & (vlVar4) & (vlVar5) & _
                          (vlVar6) & (vlVar7) & (vlVar8) & (vlVar9) & (vlVar10) & _
                          (vlVar11) & (vlVar12) & (vlVar13) & (vlVar14) & (vlVar15) & _
                          (vlVar16) & (vlVar17) & (vlVar18) & (vlVar19) & (vlVar20) & _
                          (vlVar21) & (vlVar22) & (vlVar23) & (vlVar24) & (vlVar25) & _
                          (vlVar26) & (vlVar27) & (vlVar28) & (vlVar29) & (vlVar30) & _
                          (vlVar31) & (vlVar32) & (vlVar33) & (vlVar34) & (vlVar35) & _
                          (vlVar36) & (vlVar37) & (vlVar38) & (vlVar39) & (vlVar40) & _
                          (vlVar41) & (vlVar42) & (vlVar43) & (vlVar44) & (vlVar45) & _
                          (vlVar47) & (vlVar48) & (vlVar49) & (vlVar50) & _
                          (vlVar51) & (vlVar52) & (vlVar53) & (vlVar54) & (vlVar55) & _
                          (vlVar56) & (vlVar57) & (vlVar58) & (vlVar59) & (vlVar60) & _
                          (vlVar61) & (vlVar62) & (vlVar63) & (vlVar64) & (vlVar65) & _
                          (vlVar66) & (vlVar67)
                        
                vlLinea = Replace(vlLinea, ",", ".")
                Print #2, vlLinea
                
                'Guarda el detalle de Salud de la póliza informada
                vlMtoSal = IIf(IsNull(vgRs4!mto_salud), 0, Format(vgRs4!mto_salud, "#0.00"))
                Call flGrabaDetPagReg(clTipRegPR, clTipMovSinSPR, vlVar2, vlNumOrd, vlFecPago, vlCodPen, vlAfp, vlVar12, vlMtoSal)
        
                'Contador de Pagos Recurrentes - Salud
                vlNumCasosPRSalud = vlNumCasosPRSalud + 1
                vlMtoPRSalud = vlMtoPRSalud + vlMtoSal
            
                If vlNumCasosPRPension <> vlNumCasosPRSalud Then
                    a = 1
                End If
            
            End If
            vgRs4.Close
            
            '****************** Movimiento de Retención ****************************
            'Verifica si existe retención de pagos
            vlMtoRet = 0
            vgQuery = "SELECT "  'cod_tipoidenreceptor,num_idenreceptor,gls_nomreceptor,"
            'vgQuery = vgQuery & "gls_nomsegreceptor,gls_patreceptor,gls_matreceptor,"
            vgQuery = vgQuery & " sum(mto_haber) as mto_retencion "
            vgQuery = vgQuery & "FROM PP_TMAE_LIQPAGOPENDEF "
            vgQuery = vgQuery & "WHERE cod_tipopago='" & clRetPR & "' "
            vgQuery = vgQuery & "and cod_tipreceptor='" & clRetPR & "' "
            vgQuery = vgQuery & "and fec_pago='" & vlFecPago & "' "
            vgQuery = vgQuery & "and num_poliza='" & vlNumPol & "' "
            vgQuery = vgQuery & "and num_orden=" & vlNumOrd & " "
            Set vgRs4 = vgConexionBD.Execute(vgQuery)
            If Not (vgRs4.EOF) Then
                'While Not (vgRs4.EOF)
                    vlVar3 = flObtieneViaPago(Trim(vgRs!Cod_ViaPago))
                    vlCodViaPag = vlVar3
                    vlVar3 = Format(Trim(vlVar3), "00000")
                    'If (Len(Trim(vgRs4!Num_IdenReceptor)) <= 12) Then
                    '    vlVar11 = Format(Trim(vgRs4!Cod_TipoIdenReceptor), "00") & (Trim(vgRs4!Num_IdenReceptor) & Space(12 - Len(Trim(vgRs4!Num_IdenReceptor))))
                    'Else
                    '    vlVar11 = Format(Trim(vgRs4!Cod_TipoIdenReceptor), "00") & Mid(Trim(vgRs4!Num_IdenReceptor), 1, 12)
                    'End If
                    vlVar11 = identit
                    'vlVar12 = fgFormarNombreCompleto(IIf(IsNull(vgRs4!Gls_NomReceptor), "", Trim(vgRs4!Gls_NomReceptor)), IIf(IsNull(vgRs4!Gls_NomSegReceptor), "", Trim(vgRs4!Gls_NomSegReceptor)), IIf(IsNull(vgRs4!Gls_PatReceptor), "", Trim(vgRs4!Gls_PatReceptor)), IIf(IsNull(vgRs4!Gls_MatReceptor), "", Trim(vgRs4!Gls_MatReceptor)))
                    'If (Len(Trim(vlVar12)) <= 60) Then
                    '    vlVar12 = vlVar12 & Space(60 - Len(Trim(vlVar12)))
                    'Else
                    '    vlVar12 = Mid(vlVar12, 1, 60)
                    'End If
                    vlVar12 = nombreTit
                    vlVar14 = Mid(vlFecPago, 3, 4) & Mid(vlNumPol, 6, 5) & "-" & Format(vlNumOrd, "00")
                    If (Len(Trim(vlVar14)) <= 14) Then
                        vlVar14 = Trim(vlVar14) & Space(14 - Len(Trim(vlVar14)))
                    Else
                        vlVar14 = Mid(Trim(vlVar14), 1, 14)
                    End If
                    vlVar19 = Format(Trim(vgRs!cod_afp), "00000")
                    vlVar29 = Format(Trim(vgRs!Cod_TipPension), "00000")
                    vlVar37 = Format(Trim(vgRs!Cod_TipPension), "00000")
                    vlVar61 = Format(Trim(clTipMovSinRPR), "00000")
                    vlVar62 = IIf(IsNull(vgRs4!MTO_RETENCION), 0, Format(vgRs4!MTO_RETENCION, "#0.00"))
                    'vlVar65_Res = flFormatNum5_2(vlVar62)
                    vlVar62 = flFormatNum18_2(vlVar62)
                    vlVar65_Res = vlVar62
                    vlVar67 = flObtieneTipPer(Trim(vgRs!Cod_ViaPago))
                
                    'Imprime la linea 40 Pago de Pensiones - Recurrentes Retención
                    vlLinea = (vlVar1) & (vlVar2) & (vlVar3) & (vlVar4) & (vlVar5) & _
                              (vlVar6) & (vlVar7) & (vlVar8) & (vlVar9) & (vlVar10) & _
                              (vlVar11) & (vlVar12) & (vlVar13) & (vlVar14) & (vlVar15) & _
                              (vlVar16) & (vlVar17) & (vlVar18) & (vlVar19) & (vlVar20) & _
                              (vlVar21) & (vlVar22) & (vlVar23) & (vlVar24) & (vlVar25) & _
                              (vlVar26) & (vlVar27) & (vlVar28) & (vlVar29) & (vlVar30) & _
                              (vlVar31) & (vlVar32) & (vlVar33) & (vlVar34) & (vlVar35) & _
                              (vlVar36) & (vlVar37) & (vlVar38) & (vlVar39) & (vlVar40) & _
                              (vlVar41) & (vlVar42) & (vlVar43) & (vlVar44) & (vlVar45) & _
                              (vlVar47) & (vlVar48) & (vlVar49) & (vlVar50) & _
                              (vlVar51) & (vlVar52) & (vlVar53) & (vlVar54) & (vlVar55) & _
                              (vlVar56) & (vlVar57) & (vlVar58) & (vlVar59) & (vlVar60) & _
                              (vlVar61) & (vlVar62) & (vlVar63) & (vlVar64) & (vlVar65) & _
                              (vlVar66) & (vlVar67)
                            
                    vlLinea = Replace(vlLinea, ",", ".")
                    Print #2, vlLinea
                
                    'Guarda el detalle de Retención de la póliza informada
                    vlMtoRet = IIf(IsNull(vgRs4!MTO_RETENCION), 0, Format(vgRs4!MTO_RETENCION, "#0.00"))
                    Call flGrabaDetPagReg(clTipRegPR, clTipMovSinRPR, vlVar2, vlNumOrd, vlFecPago, vlCodPen, vlAfp, vlVar12, vlMtoRet, vlVar11, vlCodViaPag)
                
                    'Contador de Pagos Recurrentes - Retención
                    vlNumCasosPRRetencion = vlNumCasosPRRetencion + 1
                    vlMtoPRRetencion = vlMtoPRRetencion + vlMtoRet
                    'vgRs4.MoveNext
                'Wend
            End If
            vgRs4.Close
            
            
            
            '**************** RESUMEN DE MOVIMIENTOS (por persona) ******************
            vlVar29 = Format(Trim(vgRs!Cod_TipPension), "00000")
            vlVar37 = Format(Trim(vgRs!Cod_TipPension), "00000")
            vlVar61 = Format(Trim(clTipMovSinRes7PR), "00000")
            vlVar62 = IIf(IsNull(vgRs!mto_pensiontot), 0, Format(vgRs!mto_pensiontot, "#0.00"))
            vlMtoPenTot = vlVar62
            vlVar62 = flFormatNum18_2(vlVar62)
            'vlVar66 = "0000000000000"
            'Imprime la linea 2 Pago de Pensiones - Resumen por persona
            vlLinea = (vlVar1) & (vlVar2) & (vlVar3_Res) & (vlVar4) & (vlVar5) & _
                      (vlVar6) & (vlVar7) & (vlVar8) & (vlVar9) & (vlVar10) & _
                      (vlVar11_Res) & (vlVar12_Res) & (vlVar13) & (vlVar14_Res) & (vlVar15) & _
                      (vlVar16) & (vlVar17) & (vlVar18) & (vlVar19_Res) & (vlVar20) & _
                      (vlVar21) & (vlVar22) & (vlVar23) & (vlVar24) & (vlVar25) & _
                      (vlVar26) & (vlVar27) & (vlVar28) & (vlVar29) & (vlVar30) & _
                      (vlVar31) & (vlVar32) & (vlVar33) & (vlVar34) & (vlVar35) & _
                      (vlVar36) & (vlVar37) & (vlVar38) & (vlVar39) & (vlVar40) & _
                      (vlVar41) & (vlVar42) & (vlVar43) & (vlVar44) & (vlVar45) & _
                      (vlVar47) & (vlVar48) & (vlVar49) & (vlVar50) & _
                      (vlVar51) & (vlVar52) & (vlVar53) & (vlVar54) & (vlVar55) & _
                      (vlVar56) & (vlVar57) & (vlVar58) & (vlVar59) & (vlVar60) & _
                      (vlVar61) & (vlVar62) & (vlVar63_Res) & (vlVar64_Res) & (vlVar65_Res) & _
                      (vlVar66_Res) & (vlVar67_Res)
            
            vlLinea = Replace(vlLinea, ",", ".")
            Print #2, vlLinea
            
            'update al registro de Mov de Pensión para Crear el registro 2 (detalle montos)
            Call flActDetallePagReg(clTipRegPR, clTipMovSinLPR, vlNumPol, vlNumOrd, vlMtoPenTot, vlMtoSal, vlMtoRet)
            
            'Limpia las variables utilizadas para informar la linea 2
            vlVar63_Res = String(18, "0")
            vlVar66_Res = String(18, "0")
            vlVar65_Res = String(18, "0")
            
            vgRs.MoveNext
        Wend
        vgRs.Close
        
        'Actualiza la cantidad de casos y mtos informados - Pensión
        Call flActEstadisticaPagReg(clTipRegPR, clTipMovSinLPR, vlNumCasosPRPension, vlMtoPRPension)
        'Actualiza la cantidad de casos y mtos informados - Salud
        Call flActEstadisticaPagReg(clTipRegPR, clTipMovSinSPR, vlNumCasosPRSalud, vlMtoPRSalud)
        'Actualiza la cantidad de casos y mtos informados - Retención
        Call flActEstadisticaPagReg(clTipRegPR, clTipMovSinRPR, vlNumCasosPRRetencion, vlMtoPRRetencion)

        vgRs1.MoveNext
    Wend
    vgRs1.Close
    
    
    '**************** RESUMEN DE MOVIMIENTOS (pagos liquidos - retenciones) ******************
    vlMtoPen = 0: vlMtoRet = 0
    Call flInicializaVar
    vlVar1 = Format(Trim(clTipRegPR), "00000")

    vlSql = "select distinct num_archivo,cod_afp,cod_moneda,fec_pago,(select sum(mto_pago) from "
    vlSql = vlSql & "pp_tmae_contabledetregpago where num_archivo=p.num_archivo and cod_moneda=p.cod_moneda "
    vlSql = vlSql & "and fec_pago=p.fec_pago and cod_afp=p.cod_afp and cod_tipmov='38') as mto_liquido,"
    vlSql = vlSql & "(select sum(mto_pago) from pp_tmae_contabledetregpago where num_archivo=p.num_archivo "
    vlSql = vlSql & "and cod_moneda=p.cod_moneda and fec_pago=p.fec_pago and cod_afp=p.cod_afp and cod_tipmov='40') as mto_retencion "
    vlSql = vlSql & "from pp_tmae_contabledetregpago p "
    vlSql = vlSql & "where num_archivo= " & vlNumArchivo & " "
    vlSql = vlSql & "order by cod_afp asc,cod_moneda asc,fec_pago asc "
    Set vgRs = vgConexionBD.Execute(vlSql)
    While Not (vgRs.EOF)
    
        If Trim(vgRs!Fec_Pago) <> "" Then
            vlVar4 = DateSerial(Mid(vgRs!Fec_Pago, 1, 4), Mid(vgRs!Fec_Pago, 5, 2), Mid(vgRs!Fec_Pago, 7, 2))
            vlVar7 = Mid(vgRs!Fec_Pago, 1, 6)
        Else
            vlVar4 = Space(10)
            vlVar7 = Space(6)
        End If
        If (Len(Trim(vgRs!Cod_Moneda)) <= 5) Then
            vlVar8 = Trim(vgRs!Cod_Moneda) & Space(5 - Len(Trim(vgRs!Cod_Moneda)))
        Else
            vlVar8 = Mid(Trim(vgRs!Cod_Moneda), 1, 5)
        End If
        vlVar19 = Format(Trim(vgRs!cod_afp), "00000")
        vlVar10 = Format(Trim(clRamoContPR), "00000")
        vlVar13 = Format(Trim(clFrecPagPR), "00000")
        vlVar14 = Mid(vlFecPago, 3, 4) & Mid(vlVar19, 3, 3) & Mid(vlVar8, 1, 2)
        If (Len(Trim(vlVar14)) <= 14) Then
            vlVar14 = Trim(vlVar14) & Space(14 - Len(Trim(vlVar14)))
        Else
            vlVar14 = Mid(Trim(vlVar14), 1, 14)
        End If
        vlVar23 = Format(Trim(clReaPR), "0000000000")
        vlVar61 = Format(Trim(clTipMovSinRes3PR), "00000")
        
        vlMtoPen = IIf(IsNull(vgRs!mto_liquido), 0, Format(vgRs!mto_liquido, "#0.00"))
        vlMtoRet = IIf(IsNull(vgRs!MTO_RETENCION), 0, Format(vgRs!MTO_RETENCION, "#0.00"))
        If (Trim(vgRs!cod_afp) <> "242") Then
            'vlVar62 = Format((vlMtoPen + vlMtoRet), "#0.00")
            vlVar62 = Format(vlMtoPen, "#0.00")
        Else
            vlVar62 = Format(vlMtoPen, "#0.00")
        End If
        vlVar62 = flFormatNum18_2(vlVar62)
                
        'Imprime la linea 3 Pago de Pensiones - Resumen por persona
        vlLinea = (vlVar1) & (vlVar2) & (vlVar3) & (vlVar4) & (vlVar5) & _
                  (vlVar6) & (vlVar7) & (vlVar8) & (vlVar9) & (vlVar10) & _
                  (vlVar11) & (vlVar12) & (vlVar13) & (vlVar14) & (vlVar15) & _
                  (vlVar16) & (vlVar17) & (vlVar18) & (vlVar19) & (vlVar20) & _
                  (vlVar21) & (vlVar22) & (vlVar23) & (vlVar24) & (vlVar25) & _
                  (vlVar26) & (vlVar27) & (vlVar28) & (vlVar29) & (vlVar30) & _
                  (vlVar31) & (vlVar32) & (vlVar33) & (vlVar34) & (vlVar35) & _
                  (vlVar36) & (vlVar37) & (vlVar38) & (vlVar39) & (vlVar40) & _
                  (vlVar41) & (vlVar42) & (vlVar43) & (vlVar44) & (vlVar45) & _
                  (vlVar47) & (vlVar48) & (vlVar49) & (vlVar50) & _
                  (vlVar51) & (vlVar52) & (vlVar53) & (vlVar54) & (vlVar55) & _
                  (vlVar56) & (vlVar57) & (vlVar58) & (vlVar59) & (vlVar60) & _
                  (vlVar61) & (vlVar62) & (vlVar63) & (vlVar64) & (vlVar65) & _
                  (vlVar66) & (vlVar67)
                
        vlLinea = Replace(vlLinea, ",", ".")
        If chkPagoA.Value = 1 Then
            Print #2, vlLinea
        End If
        
        
        vgRs.MoveNext
    Wend
    vgRs.Close
    
    
    
    
    '**************** DETALLE DE MOVIMIENTO DE RETENCIÓN (integra) ************************************
    Call flInicializaVar
    vlVar1 = Format(Trim(clTipRegPR), "00000")

    vlSql = "select num_poliza,fec_pago,cod_moneda,cod_tippension,cod_afp,cod_viapago,"
    vlSql = vlSql & "num_idenreceptor,gls_nombre,mto_pago as mto_retencion "
    vlSql = vlSql & "from pp_tmae_contabledetregpago "
    vlSql = vlSql & "where num_archivo= " & vlNumArchivo & " "
    vlSql = vlSql & "and cod_tipmov='40' "
    vlSql = vlSql & "and cod_afp='242'"
    vlSql = vlSql & "order by num_poliza,fec_pago "
    Set vgRs = vgConexionBD.Execute(vlSql)
    While Not (vgRs.EOF)
    
        vlVar2 = Format(Trim(vgRs!num_poliza), "0000000000")
        vlVar3 = Trim(vgRs!Cod_ViaPago)
        vlVar3 = Format(Trim(vlVar3), "00000")
        If Trim(vgRs!Fec_Pago) <> "" Then
            vlVar4 = DateSerial(Mid(vgRs!Fec_Pago, 1, 4), Mid(vgRs!Fec_Pago, 5, 2), Mid(vgRs!Fec_Pago, 7, 2))
            vlVar7 = Mid(vgRs!Fec_Pago, 1, 6)
        Else
            vlVar4 = Space(10)
            vlVar7 = Space(6)
        End If
        If (Len(Trim(vgRs!Cod_Moneda)) <= 5) Then
            vlVar8 = Trim(vgRs!Cod_Moneda) & Space(5 - Len(Trim(vgRs!Cod_Moneda)))
        Else
            vlVar8 = Mid(Trim(vgRs!Cod_Moneda), 1, 5)
        End If
        vlVar9 = Format(Trim(vgRs!Cod_TipPension), "00000")
        vlVar10 = Format(Trim(clRamoContPR), "00000")
        If (Len(Trim(vgRs!Num_IdenReceptor)) <= 14) Then
            vlVar11 = (Trim(vgRs!Num_IdenReceptor) & Space(14 - Len(Trim(vgRs!Num_IdenReceptor))))
        Else
            vlVar11 = Mid(Trim(vgRs!Num_IdenReceptor), 1, 14)
        End If
        vlVar12 = IIf(IsNull(vgRs!gls_nombre), "", Trim(vgRs!gls_nombre))
        If (Len(Trim(vlVar12)) <= 60) Then
            vlVar12 = vlVar12 & Space(60 - Len(Trim(vlVar12)))
        Else
            vlVar12 = Mid(vlVar12, 1, 60)
        End If
        vlVar13 = Format(Trim(clFrecPagPR), "00000")
        vlVar14 = Mid(vlFecPago, 3, 4) & Mid(vlVar19, 3, 3) & Mid(vlVar8, 1, 2)
        If (Len(Trim(vlVar14)) <= 14) Then
            vlVar14 = Trim(vlVar14) & Space(14 - Len(Trim(vlVar14)))
        Else
            vlVar14 = Mid(Trim(vlVar14), 1, 14)
        End If
        vlVar19 = Format(Trim(vgRs!cod_afp), "00000")
        vlVar23 = Format(Trim(clReaPR), "0000000000")
        vlVar61 = Format(Trim(clTipMovSinRes4PR), "00000")
        vlVar62 = IIf(IsNull(vgRs!MTO_RETENCION), 0, Format(vgRs!MTO_RETENCION, "#0.00"))
        vlVar62 = flFormatNum18_2(vlVar62)
        vlVar67 = Trim(clTipPerNatPG) & Space(1 - Len(Trim(clTipPerNatPG)))
                
        'Imprime la linea 4 Pago de Pensiones - Detalle por persona (retenciones de integra)
        vlLinea = (vlVar1) & (vlVar2) & (vlVar3) & (vlVar4) & (vlVar5) & _
                  (vlVar6) & (vlVar7) & (vlVar8) & (vlVar9) & (vlVar10) & _
                  (vlVar11) & (vlVar12) & (vlVar13) & (vlVar14) & (vlVar15) & _
                  (vlVar16) & (vlVar17) & (vlVar18) & (vlVar19) & (vlVar20) & _
                  (vlVar21) & (vlVar22) & (vlVar23) & (vlVar24) & (vlVar25) & _
                  (vlVar26) & (vlVar27) & (vlVar28) & (vlVar29) & (vlVar30) & _
                  (vlVar31) & (vlVar32) & (vlVar33) & (vlVar34) & (vlVar35) & _
                  (vlVar36) & (vlVar37) & (vlVar38) & (vlVar39) & (vlVar40) & _
                  (vlVar41) & (vlVar42) & (vlVar43) & (vlVar44) & (vlVar45) & _
                  (vlVar47) & (vlVar48) & (vlVar49) & (vlVar50) & _
                  (vlVar51) & (vlVar52) & (vlVar53) & (vlVar54) & (vlVar55) & _
                  (vlVar56) & (vlVar57) & (vlVar58) & (vlVar59) & (vlVar60) & _
                  (vlVar61) & (vlVar62) & (vlVar63) & (vlVar64) & (vlVar65) & _
                  (vlVar66) & (vlVar67)
                
        vlLinea = Replace(vlLinea, ",", ".")
        Print #2, vlLinea
        
        vgRs.MoveNext
    Wend
    vgRs.Close
    
    
Close #2
vlOpen = False

flExportarPagReg = True

Exit Function
Errores:
Screen.MousePointer = vbDefault
If Err.Number <> 0 Then
    If vlOpen Then
        Close #2
    End If
    MsgBox "Se ha producido el siguiente error : " & Err.Description & " - " & vlNumPol, vbCritical, "Error"
End If
End Function
Private Function flExportarPagRegProvision(iFecDesde As String, iFecHasta As String) As Boolean
Dim vlNumPol As String, vlFecPago As String
Dim vlMtoPen As Double, vlMtoSal As Double, vlMtoRet As Double, vlMtoPenTot As Double
Dim vlCodPen As String, vlAfp As String
Dim vlNumOrd As Integer
Dim vlCodViaPag As String
Dim vgRs1 As ADODB.Recordset

On Error GoTo Errores

flExportarPagRegProvision = False

    vlArchivo = vlArchivoContProv 'LblArchivo
    
    Screen.MousePointer = 11
    
    Open vlArchivo For Output As #1

    Me.Refresh
    vlOpen = True
        
    Call flInicializaVar
        
    'Obtiene el nº de archivo a crear
    'vlNumArchivo = flNumArchivoPagRegProvision
    vlFecCrea = Format(Date, "yyyymmdd")
    vlHorCrea = Format(Time, "hhmmss")
    
    vlSql = "SELECT COD_ELEMENTO FROM MA_TPAR_TABCOD "
    vlSql = vlSql & "WHERE COD_TABLA='" & vgCodTabla_TipMon & "' "
    vlSql = vlSql & "AND (COD_SISTEMA IS NULL OR COD_SISTEMA<>'PP') "
    vlSql = vlSql & "ORDER BY COD_ELEMENTO "
    Set vgRs1 = vgConexionBD.Execute(vlSql)
    While Not (vgRs1.EOF)
        vlMoneda = Trim(vgRs1!COD_ELEMENTO)
        
        Call flEstadisticaPagRegProvision(clTipRegPR, clTipMovSinLPR)
        Call flEstadisticaPagRegProvision(clTipRegPR, clTipMovSinSPR)
        Call flEstadisticaPagRegProvision(clTipRegPR, clTipMovSinRPR)

        vlNumCasosPRPension = 0
        vlNumCasosPRSalud = 0
        vlNumCasosPRRetencion = 0
        vlMtoPRPension = 0
        vlMtoPRSalud = 0
        vlMtoPRRetencion = 0
        
'        vgSql = "SELECT distinct a.num_poliza as num_poliza,b.cod_viapago,'00000000' as fec_pago,a.cod_moneda AS cod_moneda,a.cod_tippension,b.cod_tipoidenben,"
'        vgSql = vgSql & " b.num_idenben, b.gls_nomben,b.gls_nomsegben,b.gls_patben,b.gls_matben,a.cod_afp,b.cod_inssalud,"
'        vgSql = vgSql & " b.mto_pension-(b.mto_pension*(b.mto_plansalud/100)) as mto_liqpagar,"
'        vgSql = vgSql & " b.mto_plansalud, b.num_orden,to_char(sysdate,'yyyymm') as num_perpago,b.num_idenben as num_idenreceptor,"
'        vgSql = vgSql & " b.cod_tipoidenben as cod_tipoidenreceptor,'N' AS cod_tipreceptor, b.mto_pension as mto_pensiontot"
'        vgSql = vgSql & " FROM pp_tmae_poliza a inner join pp_tmae_ben b on a.num_poliza = b.num_poliza And a.num_endoso = b.num_endoso"
'        vgSql = vgSql & " WHERE a.num_endoso = (SELECT NVL(MAX(e.num_endoso + 1),1) AS num_endoso FROM pp_tmae_endoso e"
'        vgSql = vgSql & " WHERE e.num_poliza = a.num_poliza AND e.cod_estado = 'E' AND e.fec_efecto <= '" & iFecDesde & "' AND"
'        vgSql = vgSql & " e.fec_finefecto >= '" & iFecHasta & "') AND a.cod_estado IN (6, 7, 8)  AND"
'        vgSql = vgSql & " (a.fec_pripago < '" & iFecHasta & "' OR (a.num_mesdif > 0 AND a.fec_pripago = '" & iFecHasta & "'))"
'        vgSql = vgSql & " and b.Cod_EstPension<>'10' and a.cod_moneda ='" & vlMoneda & "'"
'        vgSql = vgSql & " union all " ' el siguiente selec solo para endoso 07
'        vgSql = vgSql & " SELECT distinct a.num_poliza as num_poliza,b.cod_viapago,'00000000' as fec_pago,a.cod_moneda AS cod_moneda,a.cod_tippension,b.cod_tipoidenben,"
'        vgSql = vgSql & " b.num_idenben, b.gls_nomben,b.gls_nomsegben,b.gls_patben,b.gls_matben,a.cod_afp,b.cod_inssalud,"
'        vgSql = vgSql & " b.mto_pension-(b.mto_pension*(b.mto_plansalud/100)) as mto_liqpagar,b.mto_plansalud, b.num_orden, "
'        vgSql = vgSql & " to_char(sysdate,'yyyymm') as num_perpago,b.num_idenben as num_idenreceptor, b.cod_tipoidenben as cod_tipoidenreceptor,"
'        vgSql = vgSql & " 'N' AS cod_tipreceptor, b.mto_pension as mto_pensiontot"
'        vgSql = vgSql & " FROM pp_tmae_poliza a inner join pp_tmae_ben b on a.num_poliza = b.num_poliza And a.num_endoso = b.num_endoso"
'        vgSql = vgSql & " left JOIN (SELECT * FROM pp_tmae_endoso Z WHERE NUM_ENDOSO=(SELECT MAX(NUM_ENDOSO) FROM pp_tmae_endoso WHERE NUM_POLIZA=Z.NUM_POLIZA)) E"
'        vgSql = vgSql & " ON a.num_poliza=e.num_poliza"
'        vgSql = vgSql & " where a.num_endoso = (SELECT NVL(MAX(e.num_endoso + 1),1) AS num_endoso FROM pp_tmae_endoso e"
'        vgSql = vgSql & " WHERE e.num_poliza = a.num_poliza AND e.cod_estado = 'E' AND e.fec_efecto <= '" & iFecDesde & "' AND"
'        vgSql = vgSql & " e.fec_finefecto >= '" & iFecHasta & "') AND a.cod_estado IN (6, 7, 8)  AND (a.fec_pripago < '" & iFecHasta & "' OR (a.num_mesdif > 0"
'        vgSql = vgSql & " AND a.fec_pripago = '" & iFecHasta & "')) and a.cod_moneda ='" & vlMoneda & "'"
'        vgSql = vgSql & " AND (b.COD_DERPEN='99' and e.cod_cauendoso='07') " 'or (b.Cod_EstPension<>'10' and e.cod_cauendoso='08'))"
'        vgSql = vgSql & " and a.num_endoso=(select NVL(MAX(e.num_endoso + 1),1) from pp_tmae_endoso where num_poliza=a.num_poliza)"
'        vgSql = vgSql & " ORDER BY num_poliza,cod_moneda"
        
        
'        vgSql = "SELECT distinct a.num_poliza as num_poliza,b.cod_viapago,'00000000' as fec_pago,a.cod_moneda AS cod_moneda,a.cod_tippension,b.cod_tipoidenben,"
'        vgSql = vgSql & " b.num_idenben, b.gls_nomben,b.gls_nomsegben,b.gls_patben,b.gls_matben,a.cod_afp,b.cod_inssalud,"
'
'        vgSql = vgSql & " (case when b.num_orden=1 then"
'        vgSql = vgSql & " ((b.mto_pension-(b.mto_pension*(b.mto_plansalud/100)))*"
'        vgSql = vgSql & " round(MONTHS_BETWEEN(to_date('" & iFecHasta & "','yyyymmdd'),to_date(b.fec_inipagopen,'yyyymmdd')),0))-"
'        vgSql = vgSql & " (select count(*) from pp_tmae_liqpagopendef where num_poliza=b.num_poliza and num_endoso=b.num_endoso and num_orden=b.num_orden)"
'        vgSql = vgSql & " Else"
'        vgSql = vgSql & " ((b.mto_pension-(b.mto_pension*(b.mto_plansalud/100)))*"
'        vgSql = vgSql & " round(MONTHS_BETWEEN(to_date('" & iFecHasta & "','yyyymmdd'),to_date((select fec_fallben from pp_tmae_ben where num_poliza=b.num_poliza and num_endoso=b.num_endoso and num_orden=1),'yyyymmdd')),0))-"
'        vgSql = vgSql & " ((nvl((select sum(mto_liqpagar) from pp_tmae_liqpagopendef where num_poliza=b.num_poliza and num_endoso=b.num_endoso and num_orden=b.num_orden),0)+"
'        vgSql = vgSql & " nvl((select sum(mto_liqpagar) from pp_tmae_liqpagopendef where fec_pago>=(select fec_fallben from pp_tmae_ben where num_poliza=b.num_poliza and num_endoso=b.num_endoso and num_orden=1) and num_poliza=b.num_poliza and num_orden=1),0)))"
'        vgSql = vgSql & " end)as mto_liqpagar,"
'
'        vgSql = vgSql & " b.mto_plansalud, b.num_orden,to_char(sysdate,'yyyymm') as num_perpago,b.num_idenben as num_idenreceptor,"
'        vgSql = vgSql & " b.cod_tipoidenben as cod_tipoidenreceptor,'N' AS cod_tipreceptor, b.mto_pension as mto_pensiontot"
'        vgSql = vgSql & " FROM pp_tmae_poliza a inner join pp_tmae_ben b on a.num_poliza = b.num_poliza And a.num_endoso = b.num_endoso"
'        vgSql = vgSql & " LEFT JOIN pp_tmae_liqpagopendef p ON b.num_poliza = p.num_poliza And b.num_endoso = p.num_endoso and b.num_orden=p.num_endoso"
'        vgSql = vgSql & " WHERE a.num_endoso = (SELECT NVL(MAX(e.num_endoso + 1),1) AS num_endoso FROM pp_tmae_endoso e"
'        vgSql = vgSql & " WHERE e.num_poliza = a.num_poliza AND e.cod_estado = 'E' AND e.fec_efecto <= '" & iFecDesde & "' AND"
'        vgSql = vgSql & " e.fec_finefecto >= '" & iFecHasta & "') AND a.cod_estado IN (6, 7, 8)  AND"
'        vgSql = vgSql & " (a.fec_pripago < '" & iFecHasta & "' OR (a.num_mesdif > 0 AND a.fec_pripago = '" & iFecHasta & "'))"
'        vgSql = vgSql & " and b.Cod_EstPension<>'10' and a.cod_moneda ='" & vlMoneda & "' "
'        vgSql = vgSql & " and fec_pago>='" & iFecDesde & "' and fec_pago<='" & iFecHasta & "'  and p.num_poliza is null"
'
'        vgSql = vgSql & " union all " ' el siguiente selec solo para endoso 07
'
'        vgSql = vgSql & " SELECT distinct a.num_poliza as num_poliza,b.cod_viapago,'00000000' as fec_pago,a.cod_moneda AS cod_moneda,a.cod_tippension,b.cod_tipoidenben,"
'        vgSql = vgSql & " b.num_idenben, b.gls_nomben,b.gls_nomsegben,b.gls_patben,b.gls_matben,a.cod_afp,b.cod_inssalud,"
'
'        vgSql = vgSql & " (case when b.num_orden=1 then"
'        vgSql = vgSql & " ((b.mto_pension-(b.mto_pension*(b.mto_plansalud/100)))*"
'        vgSql = vgSql & " round(MONTHS_BETWEEN(to_date(b.fec_inipagopen,'yyyymmdd'),to_date('" & iFecHasta & "','yyyymmdd')),0))-"
'        vgSql = vgSql & " (select count(*) from pp_tmae_liqpagopendef where num_poliza=b.num_poliza and num_endoso=b.num_endoso and num_orden=b.num_orden)"
'        vgSql = vgSql & " Else"
'        vgSql = vgSql & " ((b.mto_pension-(b.mto_pension*(b.mto_plansalud/100)))*"
'        vgSql = vgSql & " round(MONTHS_BETWEEN(to_date('" & iFecHasta & "','yyyymmdd'),to_date((select fec_fallben from pp_tmae_ben where num_poliza=b.num_poliza and num_endoso=b.num_endoso and num_orden=1),'yyyymmdd')),0))-"
'        vgSql = vgSql & " ((nvl((select sum(mto_liqpagar) from pp_tmae_liqpagopendef where num_poliza=b.num_poliza and num_endoso=b.num_endoso and num_orden=b.num_orden),0)+"
'        vgSql = vgSql & " nvl((select sum(mto_liqpagar) from pp_tmae_liqpagopendef where fec_pago>=(select fec_fallben from pp_tmae_ben where num_poliza=b.num_poliza and num_endoso=b.num_endoso and num_orden=1) and num_poliza=b.num_poliza and num_orden=1),0)))"
'        vgSql = vgSql & " end)as mto_liqpagar,"
'
'        vgSql = vgSql & " b.Mto_PlanSalud , b.Num_Orden, "
'        vgSql = vgSql & " to_char(sysdate,'yyyymm') as num_perpago,b.num_idenben as num_idenreceptor, b.cod_tipoidenben as cod_tipoidenreceptor,"
'        vgSql = vgSql & " 'N' AS cod_tipreceptor, b.mto_pension as mto_pensiontot"
'        vgSql = vgSql & " FROM pp_tmae_poliza a inner join pp_tmae_ben b on a.num_poliza = b.num_poliza And a.num_endoso = b.num_endoso"
'        vgSql = vgSql & " left JOIN (SELECT * FROM pp_tmae_endoso Z WHERE NUM_ENDOSO=(SELECT MAX(NUM_ENDOSO) FROM pp_tmae_endoso WHERE NUM_POLIZA=Z.NUM_POLIZA)) E"
'        vgSql = vgSql & " ON a.num_poliza=e.num_poliza"
'        vgSql = vgSql & " where a.num_endoso = (SELECT NVL(MAX(e.num_endoso + 1),1) AS num_endoso FROM pp_tmae_endoso e"
'        vgSql = vgSql & " WHERE e.num_poliza = a.num_poliza AND e.cod_estado = 'E' AND e.fec_efecto <= '" & iFecDesde & "' AND"
'        vgSql = vgSql & " e.fec_finefecto >= '" & iFecHasta & "') AND a.cod_estado IN (6, 7, 8)  AND (a.fec_pripago < '" & iFecHasta & "' OR (a.num_mesdif > 0"
'        vgSql = vgSql & " AND a.fec_pripago = '" & iFecHasta & "')) and a.cod_moneda ='" & vlMoneda & "'"
'        vgSql = vgSql & " AND (b.COD_DERPEN='99' and e.cod_cauendoso='07') " 'or (b.Cod_EstPension<>'10' and e.cod_cauendoso='08'))"
'        vgSql = vgSql & " and a.num_endoso=(select NVL(MAX(e.num_endoso + 1),1) from pp_tmae_endoso where num_poliza=a.num_poliza)"
'        vgSql = vgSql & " ORDER BY num_poliza,cod_moneda"
        'solo Pólizas Vigentes, con Beneficiarios Designados y Pendiente (Pago Diferido)
        
        'saco lo pendiente de pago
        vgSql = "select distinct a.num_poliza as num_poliza,b.cod_viapago,'00000000' as fec_pago,a.cod_moneda AS cod_moneda,a.cod_tippension,b.cod_tipoidenben, b.num_idenben, b.gls_nomben,b.gls_nomsegben,b.gls_patben,b.gls_matben,a.cod_afp,b.cod_inssalud,"
'        vgSql = vgSql & " (b.mto_pension-(b.mto_pension*(b.mto_plansalud/100))) as mto_liqpagar,"
        'RRR
        vgSql = vgSql & " PP_FUNCION_AJUSTE_PENSION ('" & iFecDesde & "', b.num_poliza) - (PP_FUNCION_AJUSTE_PENSION ('" & iFecDesde & "', b.num_poliza)) * (mto_plansalud/100) as mto_liqpagar, "
        
        
'        vgSql = vgSql & " (case when b.num_orden=1 then"
'        vgSql = vgSql & " (b.mto_pension-(b.mto_pension*(b.mto_plansalud/100)))* (CEIL(MONTHS_BETWEEN(to_date('" & iFecHasta & "','yyyymmdd'),to_date(b.fec_inipagopen,'yyyymmdd')))-"
'        vgSql = vgSql & " (select count(*) from pp_tmae_liqpagopendef where num_poliza=b.num_poliza and num_orden=b.num_orden))"
'        vgSql = vgSql & " Else"
'        vgSql = vgSql & " ((b.mto_pension-(b.mto_pension*(b.mto_plansalud/100)))*"
'        vgSql = vgSql & " CEIL(MONTHS_BETWEEN(to_date('" & iFecHasta & "','yyyymmdd'),to_date((select (case when to_date(fec_fallben,'yyyymmdd')>to_date(fec_inipagopen,'yyyymmdd') then fec_fallben else fec_inipagopen end) from pp_tmae_ben where num_poliza=b.num_poliza and num_endoso=b.num_endoso and num_orden=1),'yyyymmdd'))))-"
'        vgSql = vgSql & " ((nvl((select sum(mto_liqpagar) from pp_tmae_liqpagopendef where num_poliza=b.num_poliza and num_orden=b.num_orden),0)+ nvl((select sum(mto_liqpagar) from pp_tmae_liqpagopendef where fec_pago>=(select (case when to_date(fec_fallben,'yyyymmdd')>to_date(fec_inipagopen,'yyyymmdd') then fec_fallben else fec_inipagopen end) from pp_tmae_ben where num_poliza=b.num_poliza and num_endoso=b.num_endoso and num_orden=1) and num_poliza=b.num_poliza and num_orden=1),0)))"
'        vgSql = vgSql & " end)as mto_liqpagar,"
        vgSql = vgSql & " b.mto_plansalud, b.num_orden,to_char(sysdate,'yyyymm') as num_perpago,b.num_idenben as num_idenreceptor, b.cod_tipoidenben as cod_tipoidenreceptor,'N' AS cod_tipreceptor, b.mto_pension as mto_pensiontot,"
        vgSql = vgSql & " (case when nvl((select max(num_perpago) from pp_tmae_liqpagopendef where num_poliza=b.num_poliza and num_orden=b.num_orden),0)>nvl((select max(num_perpago) from pp_tmae_reliq x inner join pp_tmae_detcalcreliq y on x.num_reliq=y.num_reliq where x.num_poliza=b.num_poliza and x.num_endoso=b.num_endoso and y.num_orden=b.num_orden),0) then"
        vgSql = vgSql & "       nvl((select max(num_perpago) from pp_tmae_liqpagopendef where num_poliza=b.num_poliza and num_orden=b.num_orden),0)"
        vgSql = vgSql & " Else"
        vgSql = vgSql & "       nvl((select max(num_perpago) from pp_tmae_reliq x inner join pp_tmae_detcalcreliq y on x.num_reliq=y.num_reliq where x.num_poliza=b.num_poliza and x.num_endoso=b.num_endoso and y.num_orden=b.num_orden),0)"
        vgSql = vgSql & " end)as UltimoPAgo"
        vgSql = vgSql & " FROM pp_tmae_certificado c inner join pp_tmae_ben b on c.num_poliza = b.num_poliza and c.num_orden=b.num_orden"
        vgSql = vgSql & " inner join pp_tmae_poliza a on a.num_poliza = b.num_poliza And a.num_endoso = b.num_endoso"
        vgSql = vgSql & " left JOIN (SELECT * FROM pp_tmae_endoso Z WHERE NUM_ENDOSO=(SELECT MAX(NUM_ENDOSO) FROM pp_tmae_endoso WHERE NUM_POLIZA=Z.NUM_POLIZA)) E ON a.num_poliza=e.num_poliza "
        vgSql = vgSql & " where a.cod_estado IN (6, 7, 8)  AND (a.fec_pripago < '" & iFecHasta & "' OR(a.num_mesdif > 0 AND a.fec_pripago = '" & iFecHasta & "'))"
        vgSql = vgSql & " and a.cod_moneda ='" & vlMoneda & "' and b.num_endoso=(select max(num_endoso) from pp_tmae_poliza where num_poliza=b.num_poliza )"
        vgSql = vgSql & " and b.fec_inipagopen<='" & iFecHasta & "' AND (b.COD_DERPEN='99' and e.cod_cauendoso='07' or (b.Cod_EstPension<>'10')) and c.fec_tercer<'" & iFecDesde & "'"
        ''vgSql = vgSql & " and (c.fec_inicer=(select max(fec_inicer) from pp_tmae_CERTIFICADO where num_poliza=b.num_poliza) and c.fec_inicer<='" & iFecHasta & "' and fec_tercer>='" & iFecHasta & "')"
'        vgSql = vgSql & " and"
'        vgSql = vgSql & " round((case when b.num_orden=1 then"
'        vgSql = vgSql & " (b.mto_pension-(b.mto_pension*(b.mto_plansalud/100)))* (CEIL(MONTHS_BETWEEN(to_date('" & iFecHasta & "','yyyymmdd'),to_date(b.fec_inipagopen,'yyyymmdd')))-"
'        vgSql = vgSql & " (select count(*) from pp_tmae_liqpagopendef where num_poliza=b.num_poliza and num_orden=b.num_orden))"
'        vgSql = vgSql & " Else"
'        vgSql = vgSql & " ((b.mto_pension-(b.mto_pension*(b.mto_plansalud/100)))*"
'        vgSql = vgSql & " CEIL(MONTHS_BETWEEN(to_date('" & iFecHasta & "','yyyymmdd'),to_date((select (case when to_date(fec_fallben,'yyyymmdd')>to_date(fec_inipagopen,'yyyymmdd') then fec_fallben else fec_inipagopen end) from pp_tmae_ben where num_poliza=b.num_poliza and num_endoso=b.num_endoso and num_orden=1),'yyyymmdd'))))-"
'        vgSql = vgSql & " ((nvl((select sum(mto_liqpagar) from pp_tmae_liqpagopendef where num_poliza=b.num_poliza and num_orden=b.num_orden),0)+ nvl((select sum(mto_liqpagar) from pp_tmae_liqpagopendef where fec_pago>=(select (case when to_date(fec_fallben,'yyyymmdd')>to_date(fec_inipagopen,'yyyymmdd') then fec_fallben else fec_inipagopen end) from pp_tmae_ben where num_poliza=b.num_poliza and num_endoso=b.num_endoso and num_orden=1) and num_poliza=b.num_poliza and num_orden=1),0)))"
'        vgSql = vgSql & " end),0)>0 "
        vgSql = vgSql & " AND"
        vgSql = vgSql & " (case when nvl((select max(num_perpago) from pp_tmae_liqpagopendef where num_poliza=b.num_poliza and num_orden=b.num_orden),0)>nvl((select max(num_perpago) from pp_tmae_reliq x inner join pp_tmae_detcalcreliq y on x.num_reliq=y.num_reliq where x.num_poliza=b.num_poliza and x.num_endoso=b.num_endoso and y.num_orden=b.num_orden),0) then"
        vgSql = vgSql & " nvl((select max(num_perpago) from pp_tmae_liqpagopendef where num_poliza=b.num_poliza and num_orden=b.num_orden),0)"
        vgSql = vgSql & " Else"
        vgSql = vgSql & " nvl((select max(num_perpago) from pp_tmae_reliq x inner join pp_tmae_detcalcreliq y on x.num_reliq=y.num_reliq where x.num_poliza=b.num_poliza and x.num_endoso=b.num_endoso and y.num_orden=b.num_orden),0)"
        vgSql = vgSql & " end)<SUBSTR('" & iFecHasta & "',1,6)"
        vgSql = vgSql & " order by a.num_poliza"
        
'        vgSql = vgSql & " union all"
'        'lo uno con lo pagado
'        vgSql = "SELECT l.num_poliza,l.cod_viapago,l.fec_pago,l.cod_moneda,l.cod_tippension,"
'        vgSql = vgSql & "b.cod_tipoidenben,b.num_idenben,b.gls_nomben,b.gls_nomsegben,"
'        vgSql = vgSql & "b.gls_patben,b.gls_matben,p.cod_afp,l.cod_inssalud,"
'        vgSql = vgSql & "mto_liqpagar,l.mto_plansalud,l.num_orden,l.num_perpago,"
'        vgSql = vgSql & "l.num_idenreceptor,l.cod_tipoidenreceptor,l.cod_tipreceptor"
'        vgSql = vgSql & ",l.mto_pension as mto_pensiontot ,'' as as UltimoPAgo" '01/12/2007
'        vgSql = vgSql & "FROM pp_tmae_liqpagopendef l, pp_tmae_poliza p, pp_tmae_ben b "
'        vgSql = vgSql & "WHERE l.fec_pago between '" & iFecDesde & "' and '" & iFecHasta & "' "
'        vgSql = vgSql & "and p.cod_moneda ='" & vlMoneda & "' "
'        vgSql = vgSql & "and l.num_poliza=p.num_poliza "
'        vgSql = vgSql & "and l.num_endoso=p.num_endoso "
'        vgSql = vgSql & "and l.num_poliza=b.num_poliza "
'        vgSql = vgSql & "and l.num_endoso=b.num_endoso "
'        vgSql = vgSql & "and l.num_orden=b.num_orden "
'        vgSql = vgSql & "AND COD_TIPOPAGO='" & clRetPR & "' "
'        vgSql = vgSql & "AND COD_TIPRECEPTOR<>'" & clRetPR & "' "
'        vgSql = vgSql & "group by  l.num_poliza,l.cod_viapago,l.fec_pago,l.cod_moneda,"
'        vgSql = vgSql & "l.cod_tippension,b.cod_tipoidenben,b.num_idenben,b.gls_nomben,"
'        vgSql = vgSql & "b.gls_nomsegben,b.gls_patben,b.gls_matben,p.cod_afp,"
'        vgSql = vgSql & "l.cod_inssalud,mto_liqpagar,l.mto_plansalud,l.num_orden,l.num_perpago,"
'        vgSql = vgSql & "l.num_idenreceptor,l.cod_tipoidenreceptor,l.cod_tipreceptor,l.mto_pension "
'        vgSql = vgSql & "ORDER BY l.num_poliza,l.fec_pago,l.num_orden "
        
        Set vgRs = vgConexionBD.Execute(vgSql)
        Dim a As Integer
        While Not (vgRs.EOF)
        'If vgRs!num_poliza = "0000000038" Then MsgBox "GFGGH"
            vlNumPol = Trim(vgRs!num_poliza)
            vlNumOrd = CInt(Trim(vgRs!Num_Orden))
            vlFecPago = Trim(vgRs!Fec_Pago)
            
            If vlNumPol = "0000000038" Then
                a = 1
            End If
            
            '****************** Movimiento de Pensión ****************************
            vlVar1 = Format(Trim(clTipRegPR), "00000")
            vlVar2 = Format(Trim(vgRs!num_poliza), "0000000000")
            vlVar3 = flObtieneViaPago(Trim(vgRs!Cod_ViaPago))
            vlVar3 = Format(Trim(vlVar3), "00000")
            vlVar3_Res = vlVar3
            If (Trim(vgRs!Fec_Pago) <> "" And Trim(vgRs!Fec_Pago) <> "00000000") Then
                vlVar4 = vgRs!Fec_Pago 'DateSerial(Mid(vgRs!Fec_Pago, 1, 4), Mid(vgRs!Fec_Pago, 5, 2), Mid(vgRs!Fec_Pago, 7, 2))
                vlVar7 = Mid(vgRs!Fec_Pago, 1, 6)
            Else
                vlVar4 = Space(10)
                vlVar7 = Space(6)
            End If
            If (Len(Trim(vgRs!Cod_Moneda)) <= 5) Then
                vlVar8 = Trim(vgRs!Cod_Moneda) & Space(5 - Len(Trim(vgRs!Cod_Moneda)))
            Else
                vlVar8 = Mid(Trim(vgRs!Cod_Moneda), 1, 5)
            End If
            vlVar9 = Format(Trim(vgRs!Cod_TipPension), "00000")
            vlCodPen = Trim(vgRs!Cod_TipPension)
            vlVar10 = Format(Trim(clRamoContPR), "00000")
            If (Len(Trim(vgRs!Num_IdenBen)) <= 12) Then
                vlVar11 = Format(Trim(vgRs!Cod_TipoIdenBen), "00") & (Trim(vgRs!Num_IdenBen) & Space(12 - Len(Trim(vgRs!Num_IdenBen))))
            Else
                vlVar11 = Format(Trim(vgRs!Cod_TipoIdenBen), "00") & Mid(Trim(vgRs!Num_IdenBen), 1, 12)
            End If
            vlVar11_Res = vlVar11
            vlVar12 = fgFormarNombreCompleto(IIf(IsNull(vgRs!Gls_NomBen), "", Trim(vgRs!Gls_NomBen)), IIf(IsNull(vgRs!Gls_NomSegBen), "", Trim(vgRs!Gls_NomSegBen)), IIf(IsNull(vgRs!Gls_PatBen), "", Trim(vgRs!Gls_PatBen)), IIf(IsNull(vgRs!Gls_MatBen), "", Trim(vgRs!Gls_MatBen)))
            If (Len(Trim(vlVar12)) <= 60) Then
                vlVar12 = vlVar12 & Space(60 - Len(Trim(vlVar12)))
            Else
                vlVar12 = Mid(vlVar12, 1, 60)
            End If
            vlVar12_Res = vlVar12
            vlVar13 = Format(Trim(clFrecPagPR), "00000")
            vlVar14 = Mid(vlFecPago, 3, 4) & Mid(vlNumPol, 6, 5) & "-" & Format(vlNumOrd, "00")
            If (Len(Trim(vlVar14)) <= 14) Then
                vlVar14 = Trim(vlVar14) & Space(14 - Len(Trim(vlVar14)))
            Else
                vlVar14 = Mid(Trim(vlVar14), 1, 14)
            End If
            vlVar14_Res = vlVar14
            vlVar19 = Format(Trim(vgRs!cod_afp), "00000")
            vlVar19_Res = vlVar19
            vlAfp = Trim(vgRs!cod_afp)
            vlVar23 = Format(Trim(clReaPR), "0000000000")
            vlVar61 = Format(Trim(clTipMovSinLPR), "00000")
            vlVar62 = IIf(IsNull(vgRs!Mto_LiqPagar), 0, Format(vgRs!Mto_LiqPagar, "#0.00"))
            vlVar62 = flFormatNum18_2(vlVar62)
            vlVar63_Res = vlVar62
            vlVar67 = flObtieneTipPer(Trim(vgRs!Cod_ViaPago))
            vlVar67_Res = vlVar67
            'vlVar67 = Trim(clTipPerJurPR) & Space(1 - Len(Trim(clTipPerJurPR)))
            
            'Imprime la linea 38 Pago de Pensiones - Recurrentes Liquidos
            vlLinea = (vlVar1) & (vlVar2) & (vlVar3) & (vlVar4) & (vlVar5) & _
                      (vlVar6) & (vlVar7) & (vlVar8) & (vlVar9) & (vlVar10) & _
                      (vlVar11) & (vlVar12) & (vlVar13) & (vlVar14) & (vlVar15) & _
                      (vlVar16) & (vlVar17) & (vlVar18) & (vlVar19) & (vlVar20) & _
                      (vlVar21) & (vlVar22) & (vlVar23) & (vlVar24) & (vlVar25) & _
                      (vlVar26) & (vlVar27) & (vlVar28) & (vlVar29) & (vlVar30) & _
                      (vlVar31) & (vlVar32) & (vlVar33) & (vlVar34) & (vlVar35) & _
                      (vlVar36) & (vlVar37) & (vlVar38) & (vlVar39) & (vlVar40) & _
                      (vlVar41) & (vlVar42) & (vlVar43) & (vlVar44) & (vlVar45) & _
                      (vlVar47) & (vlVar48) & (vlVar49) & (vlVar50) & _
                      (vlVar51) & (vlVar52) & (vlVar53) & (vlVar54) & (vlVar55) & _
                      (vlVar56) & (vlVar57) & (vlVar58) & (vlVar59) & (vlVar60) & _
                      (vlVar61) & (vlVar62) & (vlVar63) & (vlVar64) & (vlVar65) & _
                      (vlVar66) & (vlVar67)
            
            vlLinea = Replace(vlLinea, ",", ".")
            Print #1, vlLinea
            
            'Guarda el detalle de Pensión de la póliza informada
            vlMtoPen = IIf(IsNull(vgRs!Mto_LiqPagar), 0, Format(vgRs!Mto_LiqPagar, "#0.00"))
            Call flGrabaDetPagRegProvision(clTipRegPR, clTipMovSinLPR, vlVar2, vlNumOrd, vlFecPago, vlCodPen, vlAfp, vlVar12, vlMtoPen)

            'Contador de Pagos Recurrentes - Pensión
            vlNumCasosPRPension = vlNumCasosPRPension + 1
            vlMtoPRPension = vlMtoPRPension + vlMtoPen

            
            '****************** Movimiento de Salud ****************************
            vlMtoSal = 0
            vgSql = "select distinct a.num_poliza as num_poliza,b.cod_viapago,'00000000' as fec_pago,a.cod_moneda AS cod_moneda,a.cod_tippension,b.cod_tipoidenben, b.num_idenben, b.gls_nomben,b.gls_nomsegben,b.gls_patben,b.gls_matben,a.cod_afp,b.cod_inssalud,"
            vgSql = vgSql & " (b.mto_pension * (b.mto_plansalud/100)) as mto_salud,"
            vgSql = vgSql & " b.mto_plansalud, b.num_orden,to_char(sysdate,'yyyymm') as num_perpago,b.num_idenben as num_idenreceptor, b.cod_tipoidenben as cod_tipoidenreceptor,'N' AS cod_tipreceptor, b.mto_pension as mto_pensiontot,"
            vgSql = vgSql & " (case when nvl((select max(num_perpago) from pp_tmae_liqpagopendef where num_poliza=b.num_poliza and num_orden=b.num_orden),0)>nvl((select max(num_perpago) from pp_tmae_reliq x inner join pp_tmae_detcalcreliq y on x.num_reliq=y.num_reliq where x.num_poliza=b.num_poliza and x.num_endoso=b.num_endoso and y.num_orden=b.num_orden),0) then"
            vgSql = vgSql & "       nvl((select max(num_perpago) from pp_tmae_liqpagopendef where num_poliza=b.num_poliza and num_orden=b.num_orden),0)"
            vgSql = vgSql & " Else"
            vgSql = vgSql & "       nvl((select max(num_perpago) from pp_tmae_reliq x inner join pp_tmae_detcalcreliq y on x.num_reliq=y.num_reliq where x.num_poliza=b.num_poliza and x.num_endoso=b.num_endoso and y.num_orden=b.num_orden),0)"
            vgSql = vgSql & " end)as UltimoPAgo"
            vgSql = vgSql & " FROM pp_tmae_certificado c inner join pp_tmae_ben b on c.num_poliza = b.num_poliza and c.num_orden=b.num_orden"
            vgSql = vgSql & " inner join pp_tmae_poliza a on a.num_poliza = b.num_poliza And a.num_endoso = b.num_endoso"
            vgSql = vgSql & " left JOIN (SELECT * FROM pp_tmae_endoso Z WHERE NUM_ENDOSO=(SELECT MAX(NUM_ENDOSO) FROM pp_tmae_endoso WHERE NUM_POLIZA=Z.NUM_POLIZA)) E ON a.num_poliza=e.num_poliza "
            vgSql = vgSql & " where a.cod_estado IN (6, 7, 8)  AND (a.fec_pripago < '" & iFecHasta & "' OR(a.num_mesdif > 0 AND a.fec_pripago = '" & iFecHasta & "'))"
            vgSql = vgSql & " and a.cod_moneda ='" & vlMoneda & "' and b.num_endoso=(select max(num_endoso) from pp_tmae_poliza where num_poliza=b.num_poliza )"
            vgSql = vgSql & " and b.fec_inipagopen<='" & iFecHasta & "' AND (b.COD_DERPEN='99' and e.cod_cauendoso='07' or (b.Cod_EstPension<>'10')) and c.fec_tercer<'" & iFecDesde & "'"
            vgSql = vgSql & " AND"
            vgSql = vgSql & " (case when nvl((select max(num_perpago) from pp_tmae_liqpagopendef where num_poliza=b.num_poliza and num_orden=b.num_orden),0)>nvl((select max(num_perpago) from pp_tmae_reliq x inner join pp_tmae_detcalcreliq y on x.num_reliq=y.num_reliq where x.num_poliza=b.num_poliza and x.num_endoso=b.num_endoso and y.num_orden=b.num_orden),0) then"
            vgSql = vgSql & " nvl((select max(num_perpago) from pp_tmae_liqpagopendef where num_poliza=b.num_poliza and num_orden=b.num_orden),0)"
            vgSql = vgSql & " Else"
            vgSql = vgSql & " nvl((select max(num_perpago) from pp_tmae_reliq x inner join pp_tmae_detcalcreliq y on x.num_reliq=y.num_reliq where x.num_poliza=b.num_poliza and x.num_endoso=b.num_endoso and y.num_orden=b.num_orden),0)"
            vgSql = vgSql & " end)<SUBSTR('" & iFecHasta & "',1,6) and a.num_poliza='" & Trim(vgRs!num_poliza) & "' and b.num_orden=" & vlNumOrd & " "
            vgSql = vgSql & " order by a.num_poliza"
            
'            vgQuery = "SELECT mto_conhabdes as mto_salud FROM PP_TMAE_PAGOPENDEF "
'            vgQuery = vgQuery & "WHERE cod_conhabdes='24' "
'            vgQuery = vgQuery & "and num_perpago ='" & Trim(vgRs!Num_PerPago) & "' "
'            vgQuery = vgQuery & "and num_poliza='" & Trim(vgRs!num_poliza) & "' "
'            vgQuery = vgQuery & "and num_orden=" & vlNumOrd & " "
'            vgQuery = vgQuery & "and num_idenreceptor ='" & Trim(vgRs!Num_IdenReceptor) & "' "
'            vgQuery = vgQuery & "and cod_tipoidenreceptor =" & CInt(Trim(vgRs!Cod_TipoIdenReceptor)) & " "
'            vgQuery = vgQuery & "and cod_tipreceptor ='" & Trim(vgRs!Cod_TipReceptor) & "' "
            Set vgRs4 = vgConexionBD.Execute(vgSql)
            If Not (vgRs4.EOF) Then
                
                vlVar14 = Mid(vlFecPago, 3, 4) & Mid(vlNumPol, 6, 5) & "-" & Format(vlNumOrd, "00")
                If (Len(Trim(vlVar14)) <= 14) Then
                    vlVar14 = Trim(vlVar14) & Space(14 - Len(Trim(vlVar14)))
                Else
                    vlVar14 = Mid(Trim(vlVar14), 1, 14)
                End If
                vlVar19 = flObtieneCodSalud(Trim(vgRs!Cod_InsSalud))
                'vlVar19 = Trim(vgRs!Cod_InsSalud) & Space(5 - Len(Trim(vgRs!Cod_InsSalud)))
                vlVar19 = Format(Trim(vlVar19), "00000")
                vlVar61 = Format(Trim(clTipMovSinSPR), "00000")
                vlVar62 = IIf(IsNull(vgRs4!mto_salud), 0, Format(vgRs4!mto_salud, "#0.00"))
                vlVar62 = flFormatNum18_2(vlVar62)
'                vlVar64_Res = IIf(IsNull(vgRs4!mto_salud), 0, Format(vgRs4!mto_salud, "#0.00"))
'                vlVar64_Res = flFormatNum5_2(vlVar64_Res)
                vlVar66_Res = IIf(IsNull(vgRs4!mto_salud), 0, Format(vgRs4!mto_salud, "#0.00"))
                vlVar66_Res = flFormatNum18_2(vlVar66_Res)
                'vlVar67 = Trim(clTipPerJurPR) & Space(5 - Len(Trim(clTipPerJurPR)))
                vlVar67 = flObtieneTipPer(Trim(vgRs!Cod_InsSalud))
                
                'Imprime la linea 39 Pago de Pensiones - Recurrentes Salud
                vlLinea = (vlVar1) & (vlVar2) & (vlVar3) & (vlVar4) & (vlVar5) & _
                          (vlVar6) & (vlVar7) & (vlVar8) & (vlVar9) & (vlVar10) & _
                          (vlVar11) & (vlVar12) & (vlVar13) & (vlVar14) & (vlVar15) & _
                          (vlVar16) & (vlVar17) & (vlVar18) & (vlVar19) & (vlVar20) & _
                          (vlVar21) & (vlVar22) & (vlVar23) & (vlVar24) & (vlVar25) & _
                          (vlVar26) & (vlVar27) & (vlVar28) & (vlVar29) & (vlVar30) & _
                          (vlVar31) & (vlVar32) & (vlVar33) & (vlVar34) & (vlVar35) & _
                          (vlVar36) & (vlVar37) & (vlVar38) & (vlVar39) & (vlVar40) & _
                          (vlVar41) & (vlVar42) & (vlVar43) & (vlVar44) & (vlVar45) & _
                          (vlVar47) & (vlVar48) & (vlVar49) & (vlVar50) & _
                          (vlVar51) & (vlVar52) & (vlVar53) & (vlVar54) & (vlVar55) & _
                          (vlVar56) & (vlVar57) & (vlVar58) & (vlVar59) & (vlVar60) & _
                          (vlVar61) & (vlVar62) & (vlVar63) & (vlVar64) & (vlVar65) & _
                          (vlVar66) & (vlVar67)
                        
                vlLinea = Replace(vlLinea, ",", ".")
                Print #1, vlLinea
                
                'Guarda el detalle de Salud de la póliza informada
                vlMtoSal = IIf(IsNull(vgRs4!mto_salud), 0, Format(vgRs4!mto_salud, "#0.00"))
                Call flGrabaDetPagRegProvision(clTipRegPR, clTipMovSinSPR, vlVar2, vlNumOrd, vlFecPago, vlCodPen, vlAfp, vlVar12, vlMtoSal)
        
                'Contador de Pagos Recurrentes - Salud
                vlNumCasosPRSalud = vlNumCasosPRSalud + 1
                vlMtoPRSalud = vlMtoPRSalud + vlMtoSal

            End If
            vgRs4.Close
            
            '****************** Movimiento de Retención ****************************
            'Verifica si existe retención de pagos
            vlMtoRet = 0
            vgQuery = "SELECT cod_tipoidenreceptor,num_idenreceptor,gls_nomreceptor,"
            vgQuery = vgQuery & "gls_nomsegreceptor,gls_patreceptor,gls_matreceptor,"
            vgQuery = vgQuery & "mto_liqpagar as mto_retencion "
            vgQuery = vgQuery & "FROM PP_TMAE_LIQPAGOPENDEF "
            vgQuery = vgQuery & "WHERE cod_tipopago='" & clRetPR & "' "
            vgQuery = vgQuery & "and cod_tipreceptor='" & clRetPR & "' "
            vgQuery = vgQuery & "and fec_pago='" & vlFecPago & "' "
            vgQuery = vgQuery & "and num_poliza='" & vlNumPol & "' "
            vgQuery = vgQuery & "and num_orden=" & vlNumOrd & " "
            Set vgRs4 = vgConexionBD.Execute(vgQuery)
            If Not (vgRs4.EOF) Then
                vlVar3 = flObtieneViaPago(Trim(vgRs!Cod_ViaPago))
                vlCodViaPag = vlVar3
                vlVar3 = Format(Trim(vlVar3), "00000")
                If (Len(Trim(vgRs4!Num_IdenReceptor)) <= 12) Then
                    vlVar11 = Format(Trim(vgRs4!Cod_TipoIdenReceptor), "00") & (Trim(vgRs4!Num_IdenReceptor) & Space(12 - Len(Trim(vgRs4!Num_IdenReceptor))))
                Else
                    vlVar11 = Format(Trim(vgRs4!Cod_TipoIdenReceptor), "00") & Mid(Trim(vgRs4!Num_IdenReceptor), 1, 12)
                End If
                vlVar12 = fgFormarNombreCompleto(IIf(IsNull(vgRs4!Gls_NomReceptor), "", Trim(vgRs4!Gls_NomReceptor)), IIf(IsNull(vgRs4!Gls_NomSegReceptor), "", Trim(vgRs4!Gls_NomSegReceptor)), IIf(IsNull(vgRs4!Gls_PatReceptor), "", Trim(vgRs4!Gls_PatReceptor)), IIf(IsNull(vgRs4!Gls_MatReceptor), "", Trim(vgRs4!Gls_MatReceptor)))
                If (Len(Trim(vlVar12)) <= 60) Then
                    vlVar12 = vlVar12 & Space(60 - Len(Trim(vlVar12)))
                Else
                    vlVar12 = Mid(vlVar12, 1, 60)
                End If
                vlVar14 = Mid(vlFecPago, 3, 4) & Mid(vlNumPol, 6, 5) & "-" & Format(vlNumOrd, "00")
                If (Len(Trim(vlVar14)) <= 14) Then
                    vlVar14 = Trim(vlVar14) & Space(14 - Len(Trim(vlVar14)))
                Else
                    vlVar14 = Mid(Trim(vlVar14), 1, 14)
                End If
                vlVar19 = Format(Trim(vgRs!cod_afp), "00000")
                vlVar61 = Format(Trim(clTipMovSinRPR), "00000")
                vlVar62 = IIf(IsNull(vgRs4!MTO_RETENCION), 0, Format(vgRs4!MTO_RETENCION, "#0.00"))
                vlVar62 = flFormatNum18_2(vlVar62)
                vlVar65_Res = vlVar62
                vlVar67 = flObtieneTipPer(Trim(vgRs!Cod_ViaPago))
            
                'Imprime la linea 40 Pago de Pensiones - Recurrentes Retención
                vlLinea = (vlVar1) & (vlVar2) & (vlVar3) & (vlVar4) & (vlVar5) & _
                          (vlVar6) & (vlVar7) & (vlVar8) & (vlVar9) & (vlVar10) & _
                          (vlVar11) & (vlVar12) & (vlVar13) & (vlVar14) & (vlVar15) & _
                          (vlVar16) & (vlVar17) & (vlVar18) & (vlVar19) & (vlVar20) & _
                          (vlVar21) & (vlVar22) & (vlVar23) & (vlVar24) & (vlVar25) & _
                          (vlVar26) & (vlVar27) & (vlVar28) & (vlVar29) & (vlVar30) & _
                          (vlVar31) & (vlVar32) & (vlVar33) & (vlVar34) & (vlVar35) & _
                          (vlVar36) & (vlVar37) & (vlVar38) & (vlVar39) & (vlVar40) & _
                          (vlVar41) & (vlVar42) & (vlVar43) & (vlVar44) & (vlVar45) & _
                          (vlVar47) & (vlVar48) & (vlVar49) & (vlVar50) & _
                          (vlVar51) & (vlVar52) & (vlVar53) & (vlVar54) & (vlVar55) & _
                          (vlVar56) & (vlVar57) & (vlVar58) & (vlVar59) & (vlVar60) & _
                          (vlVar61) & (vlVar62) & (vlVar63) & (vlVar64) & (vlVar65) & _
                          (vlVar66) & (vlVar67)
                        
                vlLinea = Replace(vlLinea, ",", ".")
                Print #1, vlLinea
            
                'Guarda el detalle de Retención de la póliza informada
                vlMtoRet = IIf(IsNull(vgRs4!MTO_RETENCION), 0, Format(vgRs4!MTO_RETENCION, "#0.00"))
                Call flGrabaDetPagRegProvision(clTipRegPR, clTipMovSinRPR, vlVar2, vlNumOrd, vlFecPago, vlCodPen, vlAfp, vlVar12, vlMtoRet, vlVar11, vlCodViaPag)
            
                'Contador de Pagos Recurrentes - Retención
                vlNumCasosPRRetencion = vlNumCasosPRRetencion + 1
                vlMtoPRRetencion = vlMtoPRRetencion + vlMtoRet
                
            End If
            vgRs4.Close
            
            
            
            '**************** RESUMEN DE MOVIMIENTOS (por persona) ******************
            vlVar61 = Format(Trim(clTipMovSinRes2PR), "00000")
            vlVar62 = IIf(IsNull(vgRs!mto_pensiontot), 0, Format(vgRs!mto_pensiontot, "#0.00"))
            vlMtoPenTot = vlVar62
            vlVar62 = flFormatNum18_2(vlVar62)
            
            'Imprime la linea 2 Pago de Pensiones - Resumen por persona
            vlLinea = (vlVar1) & (vlVar2) & (vlVar3_Res) & (vlVar4) & (vlVar5) & _
                      (vlVar6) & (vlVar7) & (vlVar8) & (vlVar9) & (vlVar10) & _
                      (vlVar11_Res) & (vlVar12_Res) & (vlVar13) & (vlVar14_Res) & (vlVar15) & _
                      (vlVar16) & (vlVar17) & (vlVar18) & (vlVar19_Res) & (vlVar20) & _
                      (vlVar21) & (vlVar22) & (vlVar23) & (vlVar24) & (vlVar25) & _
                      (vlVar26) & (vlVar27) & (vlVar28) & (vlVar29) & (vlVar30) & _
                      (vlVar31) & (vlVar32) & (vlVar33) & (vlVar34) & (vlVar35) & _
                      (vlVar36) & (vlVar37) & (vlVar38) & (vlVar39) & (vlVar40) & _
                      (vlVar41) & (vlVar42) & (vlVar43) & (vlVar44) & (vlVar45) & _
                      (vlVar47) & (vlVar48) & (vlVar49) & (vlVar50) & _
                      (vlVar51) & (vlVar52) & (vlVar53) & (vlVar54) & (vlVar55) & _
                      (vlVar56) & (vlVar57) & (vlVar58) & (vlVar59) & (vlVar60) & _
                      (vlVar61) & (vlVar62) & (vlVar63_Res) & (vlVar64_Res) & (vlVar65_Res) & _
                      (vlVar66_Res) & (vlVar67_Res)
            
            vlLinea = Replace(vlLinea, ",", ".")
            Print #1, vlLinea
            
            'update al registro de Mov de Pensión para Crear el registro 2 (detalle montos)
            Call flActDetallePagRegProvision(clTipRegPR, clTipMovSinLPR, vlNumPol, vlNumOrd, vlMtoPenTot, vlMtoSal, vlMtoRet)
            
            'Limpia las variables utilizadas para informar la linea 2
            vlVar63_Res = String(18, "0")
            vlVar66_Res = String(18, "0")
            vlVar65_Res = String(18, "0")
            
            vgRs.MoveNext
        Wend
        vgRs.Close
        
        'Actualiza la cantidad de casos y mtos informados - Pensión
        Call flActEstadisticaPagRegProvision(clTipRegPR, clTipMovSinLPR, vlNumCasosPRPension, vlMtoPRPension)
        'Actualiza la cantidad de casos y mtos informados - Salud
        Call flActEstadisticaPagRegProvision(clTipRegPR, clTipMovSinSPR, vlNumCasosPRSalud, vlMtoPRSalud)
        'Actualiza la cantidad de casos y mtos informados - Retención
        Call flActEstadisticaPagRegProvision(clTipRegPR, clTipMovSinRPR, vlNumCasosPRRetencion, vlMtoPRRetencion)

        vgRs1.MoveNext
    Wend
    vgRs1.Close
    
    
    '**************** RESUMEN DE MOVIMIENTOS (pagos liquidos - retenciones) ******************
    vlMtoPen = 0: vlMtoRet = 0
    Call flInicializaVar
    vlVar1 = Format(Trim(clTipRegPR), "00000")

    vlSql = "select distinct num_archivo,cod_afp,cod_moneda,fec_pago,(select sum(mto_pago) from "
    vlSql = vlSql & "pp_tmae_contabledet_provision where num_archivo=p.num_archivo and cod_moneda=p.cod_moneda "
    vlSql = vlSql & "and fec_pago=p.fec_pago and cod_afp=p.cod_afp and cod_tipmov='38') as mto_liquido,"
    vlSql = vlSql & "(select sum(mto_pago) from pp_tmae_contabledet_provision where num_archivo=p.num_archivo "
    vlSql = vlSql & "and cod_moneda=p.cod_moneda and fec_pago=p.fec_pago and cod_afp=p.cod_afp and cod_tipmov='40') as mto_retencion "
    vlSql = vlSql & "from pp_tmae_contabledet_provision p "
    vlSql = vlSql & "where num_archivo= " & vlNumArchivo & " "
    vlSql = vlSql & "order by cod_afp asc,cod_moneda asc,fec_pago asc "
    Set vgRs = vgConexionBD.Execute(vlSql)
    While Not (vgRs.EOF)
    
        If Trim(vgRs!Fec_Pago) <> "" And Trim(vgRs!Fec_Pago) <> "00000000" Then
            vlVar4 = DateSerial(Mid(vgRs!Fec_Pago, 1, 4), Mid(vgRs!Fec_Pago, 5, 2), Mid(vgRs!Fec_Pago, 7, 2))
            vlVar7 = Mid(vgRs!Fec_Pago, 1, 6)
        Else
            vlVar4 = Space(10)
            vlVar7 = Space(6)
        End If
        If (Len(Trim(vgRs!Cod_Moneda)) <= 5) Then
            vlVar8 = Trim(vgRs!Cod_Moneda) & Space(5 - Len(Trim(vgRs!Cod_Moneda)))
        Else
            vlVar8 = Mid(Trim(vgRs!Cod_Moneda), 1, 5)
        End If
        vlVar19 = Format(Trim(vgRs!cod_afp), "00000")
        vlVar10 = Format(Trim(clRamoContPR), "00000")
        vlVar13 = Format(Trim(clFrecPagPR), "00000")
        vlVar14 = Mid(vlFecPago, 3, 4) & Mid(vlVar19, 3, 3) & Mid(vlVar8, 1, 2)
        If (Len(Trim(vlVar14)) <= 14) Then
            vlVar14 = Trim(vlVar14) & Space(14 - Len(Trim(vlVar14)))
        Else
            vlVar14 = Mid(Trim(vlVar14), 1, 14)
        End If
        vlVar23 = Format(Trim(clReaPR), "0000000000")
        'vlVar61 = Format(Trim(clTipMovSinRes3PR), "00000")
        vlVar61 = Format(Trim(55), "00000")
        
        vlMtoPen = IIf(IsNull(vgRs!mto_liquido), 0, Format(vgRs!mto_liquido, "#0.00"))
        vlMtoRet = IIf(IsNull(vgRs!MTO_RETENCION), 0, Format(vgRs!MTO_RETENCION, "#0.00"))
        If (Trim(vgRs!cod_afp) <> "242") Then
            vlVar62 = Format((vlMtoPen + vlMtoRet), "#0.00")
        Else
            vlVar62 = Format(vlMtoPen, "#0.00")
        End If
        vlVar62 = flFormatNum18_2(vlVar62)
                
        'Imprime la linea 3 Pago de Pensiones - Resumen por persona
        vlLinea = (vlVar1) & (vlVar2) & (vlVar3) & (vlVar4) & (vlVar5) & _
                  (vlVar6) & (vlVar7) & (vlVar8) & (vlVar9) & (vlVar10) & _
                  (vlVar11) & (vlVar12) & (vlVar13) & (vlVar14) & (vlVar15) & _
                  (vlVar16) & (vlVar17) & (vlVar18) & (vlVar19) & (vlVar20) & _
                  (vlVar21) & (vlVar22) & (vlVar23) & (vlVar24) & (vlVar25) & _
                  (vlVar26) & (vlVar27) & (vlVar28) & (vlVar29) & (vlVar30) & _
                  (vlVar31) & (vlVar32) & (vlVar33) & (vlVar34) & (vlVar35) & _
                  (vlVar36) & (vlVar37) & (vlVar38) & (vlVar39) & (vlVar40) & _
                  (vlVar41) & (vlVar42) & (vlVar43) & (vlVar44) & (vlVar45) & _
                  (vlVar47) & (vlVar48) & (vlVar49) & (vlVar50) & _
                  (vlVar51) & (vlVar52) & (vlVar53) & (vlVar54) & (vlVar55) & _
                  (vlVar56) & (vlVar57) & (vlVar58) & (vlVar59) & (vlVar60) & _
                  (vlVar61) & (vlVar62) & (vlVar63) & (vlVar64) & (vlVar65) & _
                  (vlVar66) & (vlVar67)
                
        vlLinea = Replace(vlLinea, ",", ".")
        Print #1, vlLinea
        
        vgRs.MoveNext
    Wend
    vgRs.Close
    
    
    
    
    '**************** DETALLE DE MOVIMIENTO DE RETENCIÓN (integra) ************************************
    Call flInicializaVar
    vlVar1 = Format(Trim(clTipRegPR), "00000")

    vlSql = "select num_poliza,fec_pago,cod_moneda,cod_tippension,cod_afp,cod_viapago,"
    vlSql = vlSql & "num_idenreceptor,gls_nombre,mto_pago as mto_retencion "
    vlSql = vlSql & "from pp_tmae_contabledet_provision "
    vlSql = vlSql & "where num_archivo= " & vlNumArchivo & " "
    vlSql = vlSql & "and cod_tipmov='40' "
    vlSql = vlSql & "and cod_afp='242'"
    vlSql = vlSql & "order by num_poliza,fec_pago "
    Set vgRs = vgConexionBD.Execute(vlSql)
    While Not (vgRs.EOF)
    
        vlVar2 = Format(Trim(vgRs!num_poliza), "0000000000")
        vlVar3 = Trim(vgRs!Cod_ViaPago)
        vlVar3 = Format(Trim(vlVar3), "00000")
        If Trim(vgRs!Fec_Pago) <> "" Then
            vlVar4 = DateSerial(Mid(vgRs!Fec_Pago, 1, 4), Mid(vgRs!Fec_Pago, 5, 2), Mid(vgRs!Fec_Pago, 7, 2))
            vlVar7 = Mid(vgRs!Fec_Pago, 1, 6)
        Else
            vlVar4 = Space(10)
            vlVar7 = Space(6)
        End If
        If (Len(Trim(vgRs!Cod_Moneda)) <= 5) Then
            vlVar8 = Trim(vgRs!Cod_Moneda) & Space(5 - Len(Trim(vgRs!Cod_Moneda)))
        Else
            vlVar8 = Mid(Trim(vgRs!Cod_Moneda), 1, 5)
        End If
        vlVar9 = Format(Trim(vgRs!Cod_TipPension), "00000")
        vlVar10 = Format(Trim(clRamoContPR), "00000")
        If (Len(Trim(vgRs!Num_IdenReceptor)) <= 14) Then
            vlVar11 = (Trim(vgRs!Num_IdenReceptor) & Space(14 - Len(Trim(vgRs!Num_IdenReceptor))))
        Else
            vlVar11 = Mid(Trim(vgRs!Num_IdenReceptor), 1, 14)
        End If
        vlVar12 = IIf(IsNull(vgRs!gls_nombre), "", Trim(vgRs!gls_nombre))
        If (Len(Trim(vlVar12)) <= 60) Then
            vlVar12 = vlVar12 & Space(60 - Len(Trim(vlVar12)))
        Else
            vlVar12 = Mid(vlVar12, 1, 60)
        End If
        vlVar13 = Format(Trim(clFrecPagPR), "00000")
        vlVar14 = Mid(vlFecPago, 3, 4) & Mid(vlVar19, 3, 3) & Mid(vlVar8, 1, 2)
        If (Len(Trim(vlVar14)) <= 14) Then
            vlVar14 = Trim(vlVar14) & Space(14 - Len(Trim(vlVar14)))
        Else
            vlVar14 = Mid(Trim(vlVar14), 1, 14)
        End If
        vlVar19 = Format(Trim(vgRs!cod_afp), "00000")
        vlVar23 = Format(Trim(clReaPR), "0000000000")
        vlVar61 = Format(Trim(clTipMovSinRes4PR), "00000")
        vlVar62 = IIf(IsNull(vgRs!MTO_RETENCION), 0, Format(vgRs!MTO_RETENCION, "#0.00"))
        vlVar62 = flFormatNum18_2(vlVar62)
        vlVar67 = Trim(clTipPerNatPG) & Space(1 - Len(Trim(clTipPerNatPG)))
                
        'Imprime la linea 4 Pago de Pensiones - Detalle por persona (retenciones de integra)
        vlLinea = (vlVar1) & (vlVar2) & (vlVar3) & (vlVar4) & (vlVar5) & _
                  (vlVar6) & (vlVar7) & (vlVar8) & (vlVar9) & (vlVar10) & _
                  (vlVar11) & (vlVar12) & (vlVar13) & (vlVar14) & (vlVar15) & _
                  (vlVar16) & (vlVar17) & (vlVar18) & (vlVar19) & (vlVar20) & _
                  (vlVar21) & (vlVar22) & (vlVar23) & (vlVar24) & (vlVar25) & _
                  (vlVar26) & (vlVar27) & (vlVar28) & (vlVar29) & (vlVar30) & _
                  (vlVar31) & (vlVar32) & (vlVar33) & (vlVar34) & (vlVar35) & _
                  (vlVar36) & (vlVar37) & (vlVar38) & (vlVar39) & (vlVar40) & _
                  (vlVar41) & (vlVar42) & (vlVar43) & (vlVar44) & (vlVar45) & _
                  (vlVar47) & (vlVar48) & (vlVar49) & (vlVar50) & _
                  (vlVar51) & (vlVar52) & (vlVar53) & (vlVar54) & (vlVar55) & _
                  (vlVar56) & (vlVar57) & (vlVar58) & (vlVar59) & (vlVar60) & _
                  (vlVar61) & (vlVar62) & (vlVar63) & (vlVar64) & (vlVar65) & _
                  (vlVar66) & (vlVar67)
                
        vlLinea = Replace(vlLinea, ",", ".")
        Print #1, vlLinea
        
        vgRs.MoveNext
    Wend
    vgRs.Close
    
    
Close #1
vlOpen = False

flExportarPagRegProvision = True

Exit Function
Errores:
Screen.MousePointer = vbDefault

If Err.Number <> 0 Then
    If vlOpen Then
        Close #1
    End If
    MsgBox "Se ha producido el siguiente error : " & Err.Description & " " & vlNumPol, vbCritical, "Error"
End If

End Function
Private Function flExportarGtoSep(iFecDesde As String, iFecHasta As String) As Boolean
Dim vlFecPago As String, vlMtoGto As Double
Dim vlCodPen As String, vlAfp As String
On Error GoTo Errores

flExportarGtoSep = False

    vlArchivo = vlArchivoCont 'LblArchivo
    
    Screen.MousePointer = 11
    
    Open vlArchivo For Output As #1
    
    vlOpen = True
    
    Call flInicializaVar
    
    'Obtiene el nº de archivo a crear
    vlNumArchivo = flNumArchivoGtoSep
    vlFecCrea = Format(Date, "yyyymmdd")
    vlHorCrea = Format(Time, "hhmmss")
    
''    vlSql = "SELECT COD_ELEMENTO FROM MA_TPAR_TABCOD "
''    vlSql = vlSql & "WHERE COD_TABLA='" & vgCodTabla_TipMon & "' "
''    vlSql = vlSql & "AND (COD_SISTEMA IS NULL OR COD_SISTEMA<>'PP') "
''    vlSql = vlSql & "ORDER BY COD_ELEMENTO "
''    Set vgRs1 = vgConexionBD.Execute(vlSql)
''    While Not (vgRs1.EOF)
''        vlMoneda = Trim(vgRs1!cod_elemento)
        
        Call flEstadisticaGtoSep(clTipRegGS, clTipMovSinGS)
        
        vlNumCasosGtoSep = 0
        vlMtoGtoSep = 0

        vgSql = ""
        vgSql = vgSql & "SELECT g.num_poliza,g.cod_viapago,g.fec_pago,p.cod_moneda,"
        vgSql = vgSql & "'NS' as cod_monedagtosep,"
        vgSql = vgSql & "p.cod_tippension,g.cod_tipoidensolicita,g.num_idensolicita,"
        vgSql = vgSql & "g.gls_nomsolicita,g.gls_nomsegsolicita,g.gls_patsolicita,g.gls_matsolicita,"
        vgSql = vgSql & "p.cod_afp,g.mto_pago,g.cod_tipopersona "
        vgSql = vgSql & "FROM PP_TMAE_PAGTERCUOMOR G,PP_TMAE_POLIZA P "
        vgSql = vgSql & "WHERE g.fec_pago between '" & iFecDesde & "' and '" & iFecHasta & "' "
        ''vgSql = vgSql & "and cod_moneda ='" & vlMoneda & "' "
        vgSql = vgSql & "and g.num_poliza=p.num_poliza "
        vgSql = vgSql & "and g.num_endoso=p.num_endoso "
        vgSql = vgSql & "ORDER BY g.num_poliza "
        Set vgRs = vgConexionBD.Execute(vgSql)
        While Not (vgRs.EOF)
            ''vlNumPol = Trim(vgRs!Num_Poliza)
            vlFecPago = Trim(vgRs!Fec_Pago)
            
            vlVar1 = Format(Trim(clTipRegGS), "00000")
            vlVar2 = Format(Trim(vgRs!num_poliza), "0000000000")
            vlVar3 = flObtieneViaPago(Trim(vgRs!Cod_ViaPago))
            vlVar3 = Format(Trim(vlVar3), "00000")
            If Trim(vgRs!Fec_Pago) <> "" Then
                vlVar4 = DateSerial(Mid(vgRs!Fec_Pago, 1, 4), Mid(vgRs!Fec_Pago, 5, 2), Mid(vgRs!Fec_Pago, 7, 2))
                vlVar7 = Mid(vgRs!Fec_Pago, 1, 6)
            Else
                vlVar4 = Space(10)
                vlVar7 = Space(6)
            End If
            If (Len(Trim(vgRs!cod_monedagtosep)) <= 5) Then
                vlVar8 = Trim(vgRs!cod_monedagtosep) & Space(5 - Len(Trim(vgRs!cod_monedagtosep)))
            Else
                vlVar8 = Mid(Trim(vgRs!cod_monedagtosep), 1, 5)
            End If
            vlMoneda = Trim(vlVar8)
            vlVar9 = Format(Trim(vgRs!Cod_TipPension), "00000")
            vlCodPen = Trim(vgRs!Cod_TipPension)
            'vlVar10 = Format(Trim(clRamoContGS), "00000")
            
            Select Case vgRs!Cod_TipPension
                Case "04"
                    vlVar10 = Format(Trim("76"), "00000")
                Case "05"
                    vlVar10 = Format(Trim("76"), "00000")
                Case "06"
                    vlVar10 = Format(Trim("94"), "00000")
                Case "07"
                    vlVar10 = Format(Trim("94"), "00000")
                Case "08"
                    vlVar10 = Format(Trim("95"), "00000")
                Case "09"
                    vlVar10 = Format(Trim("76"), "00000")
                Case "10"
                    vlVar10 = Format(Trim("76"), "00000")
                Case "11"
                    vlVar10 = Format(Trim("94"), "00000")
                Case "12"
                    vlVar10 = Format(Trim("94"), "00000")
            End Select
            
            
            If (Len(Trim(vgRs!num_idensolicita)) <= 12) Then
                vlVar11 = Format(Trim(vgRs!cod_tipoidensolicita), "00") & (Trim(vgRs!num_idensolicita) & Space(12 - Len(Trim(vgRs!num_idensolicita))))
            Else
                vlVar11 = Format(Trim(vgRs!cod_tipoidensolicita), "00") & Mid(Trim(vgRs!num_idensolicita), 1, 12)
            End If
            vlVar12 = fgFormarNombreCompleto(IIf(IsNull(vgRs!gls_nomsolicita), "", Trim(vgRs!gls_nomsolicita)), IIf(IsNull(vgRs!gls_nomsegsolicita), "", Trim(vgRs!gls_nomsegsolicita)), IIf(IsNull(vgRs!gls_patsolicita), "", Trim(vgRs!gls_patsolicita)), IIf(IsNull(vgRs!gls_matsolicita), "", Trim(vgRs!gls_matsolicita)))
            If (Len(Trim(vlVar12)) <= 60) Then
                vlVar12 = vlVar12 & Space(60 - Len(Trim(vlVar12)))
            Else
                vlVar12 = Mid(vlVar12, 1, 60)
            End If
            vlVar13 = Format(Trim(clFrecPagGS), "00000")
            vlVar19 = Format(Trim(vgRs!cod_afp), "00000")
            vlAfp = Trim(vgRs!cod_afp)
            vlVar23 = Format(Trim(clReaGS), "0000000000")
            vlVar25 = Format(Trim(vgRs!Cod_TipPension), "00000")
            vlVar61 = Format(Trim(clTipMovSinGS), "00000")
            vlVar62 = IIf(IsNull(vgRs!mto_pago), 0, Format(vgRs!mto_pago, "#0.00"))
            ''vlVar62 = String(18 - Len(Trim(vlVar62)), "0") & Trim(vlVar62)
            vlVar62 = flFormatNum18_2(vlVar62)
            'vlVar67 = Trim(clTipPerNatGS) & Space(1 - Len(Trim(clTipPerNatGS)))
            If (Len(Trim(vgRs!COD_TIPOPERSONA)) <= 1) Then
                vlVar67 = Trim(vgRs!COD_TIPOPERSONA) & Space(1 - Len(Trim(vgRs!COD_TIPOPERSONA)))
            Else
                vlVar67 = Mid(Trim(vgRs!COD_TIPOPERSONA), 1, 1)
            End If
            vlLinea = (vlVar1) & (vlVar2) & (vlVar3) & (vlVar4) & (vlVar5) & _
                      (vlVar6) & (vlVar7) & (vlVar8) & (vlVar9) & (vlVar10) & _
                      (vlVar11) & (vlVar12) & (vlVar13) & (vlVar14) & (vlVar15) & _
                      (vlVar16) & (vlVar17) & (vlVar18) & (vlVar19) & (vlVar20) & _
                      (vlVar21) & (vlVar22) & (vlVar23) & (vlVar24) & (vlVar25) & _
                      (vlVar26) & (vlVar27) & (vlVar28) & (vlVar29) & (vlVar30) & _
                      (vlVar31) & (vlVar32) & (vlVar33) & (vlVar34) & (vlVar35) & _
                      (vlVar36) & (vlVar37) & (vlVar38) & (vlVar39) & (vlVar40) & _
                      (vlVar41) & (vlVar42) & (vlVar43) & (vlVar44) & (vlVar45) & _
                      (vlVar47) & (vlVar48) & (vlVar49) & (vlVar50) & _
                      (vlVar51) & (vlVar52) & (vlVar53) & (vlVar54) & (vlVar55) & _
                      (vlVar56) & (vlVar57) & (vlVar58) & (vlVar59) & (vlVar60) & _
                      (vlVar61) & (vlVar62) & (vlVar63) & (vlVar64) & (vlVar65) & _
                      (vlVar66) & (vlVar67)

            vlLinea = Replace(vlLinea, ",", ".")
            Print #1, vlLinea

            'Guarda el detalle de Pensión de la póliza informada
            vlMtoGto = IIf(IsNull(vgRs!mto_pago), 0, Format(vgRs!mto_pago, "#0.00"))
            Call flGrabaDetGtoSep(clTipRegGS, clTipMovSinGS, vlVar2, vlFecPago, vlCodPen, vlAfp, vlVar12, vlMtoGto)
           
            vlNumCasosGtoSep = vlNumCasosGtoSep + 1
            vlMtoGtoSep = vlMtoGtoSep + vlMtoGto

            vgRs.MoveNext
        Wend
        vgRs.Close
        
        'Actualiza la cantidad de casos y mtos informados - Gto Sepelio
        Call flActEstadisticaGtoSep(clTipRegGS, clTipMovSinGS, vlNumCasosGtoSep, vlMtoGtoSep)
        
''        vgRs1.MoveNext
''    Wend
''    vgRs1.Close
    
Close #1
vlOpen = False

flExportarGtoSep = True

Exit Function
Errores:
Screen.MousePointer = vbDefault
If Err.Number <> 0 Then
    If vlOpen Then
        Close #1
    End If
    MsgBox "Se ha producido el siguiente error : " & Err.Description, vbCritical, "Error"
End If
End Function

Private Function flExportarHabDes(iFecDesde As String, iFecHasta As String) As Boolean
Dim vlFecPago As String, vlMtoHab As Double
Dim vlCodPen As String, vlAfp As String
On Error GoTo Errores
Dim vgRs1 As ADODB.Recordset
flExportarHabDes = False

    vlArchivo = vlArchivoCont 'LblArchivo
    
    Screen.MousePointer = 11
    
    Open vlArchivo For Output As #1
    
    vlOpen = True
    
    Call flInicializaVar
    
    'Obtiene el nº de archivo a crear
    vlNumArchivo = flNumArchivoHabDes
    vlFecCrea = Format(Date, "yyyymmdd")
    vlHorCrea = Format(Time, "hhmmss")
    vlMtoHabDes = 0
    
    vlSql = "SELECT COD_ELEMENTO FROM MA_TPAR_TABCOD "
    vlSql = vlSql & "WHERE COD_TABLA='" & vgCodTabla_TipMon & "' "
    vlSql = vlSql & "AND (COD_SISTEMA IS NULL OR COD_SISTEMA<>'PP') "
    vlSql = vlSql & "ORDER BY COD_ELEMENTO "
    Set vgRs1 = vgConexionBD.Execute(vlSql)
    While Not (vgRs1.EOF)
        vlMoneda = Trim(vgRs1!COD_ELEMENTO)
        vgMonedaCodOfi = vlMoneda
        Call flEstadisticaHabDes(clTipRegGS, clTipMovHabDes)
        
        vlNumCasosGtoSep = 0
        vlMtoGtoSep = 0
        vlNumCasosHabDes = 0

        vgSql = ""
        vgSql = "SELECT P.num_poliza,B.num_orden,B.cod_viapago,H.fec_CREA AS fec_pago,H.cod_moneda,'NS' as cod_monedagtosep,p.cod_tippension,"
        vgSql = vgSql & " b.cod_tipoidenben AS cod_tipoidensolicita,b.num_idenben AS num_idensolicita,b.gls_nomben AS gls_nomsolicita,"
        vgSql = vgSql & " b.gls_nomsegben AS gls_nomsegsolicita,b.gls_patben AS gls_patsolicita,b.gls_matben AS gls_matsolicita,p.cod_afp,h.mto_total AS mto_pago,'N' AS cod_tipopersona"
        vgSql = vgSql & " FROM PP_TMAE_HABDES H INNER JOIN PP_TMAE_BEN B ON H.NUM_POLIZA=B.NUM_POLIZA AND H.NUM_ENDOSO=B.NUM_ENDOSO AND H.NUM_ORDEN=B.NUM_ORDEN"
        vgSql = vgSql & " INNER JOIN PP_TMAE_POLIZA P ON H.num_poliza=p.num_poliza and H.num_endoso=p.num_endoso"
        vgSql = vgSql & " WHERE H.fec_INIHABDES >='" & iFecDesde & "' and H.fec_INIHABDES <= '" & iFecHasta & "' AND COD_CONHABDES='06' AND"
        vgSql = vgSql & " COD_DERPEN='99' AND COD_ESTPENSION<>'10' and H.cod_moneda ='" & vlMoneda & "' ORDER BY P.num_poliza"  'and cod_moneda ='" & vlMoneda & "'
'        vgSql = vgSql & "SELECT g.num_poliza,g.cod_viapago,g.fec_pago,p.cod_moneda,"
'        vgSql = vgSql & "'NS' as cod_monedagtosep,"
'        vgSql = vgSql & "p.cod_tippension,g.cod_tipoidensolicita,g.num_idensolicita,"
'        vgSql = vgSql & "g.gls_nomsolicita,g.gls_nomsegsolicita,g.gls_patsolicita,g.gls_matsolicita,"
'        vgSql = vgSql & "p.cod_afp,g.mto_pago,g.cod_tipopersona "
'        vgSql = vgSql & "FROM PP_TMAE_PAGTERCUOMOR G,PP_TMAE_POLIZA P "
'        vgSql = vgSql & "WHERE g.fec_pago between '" & iFecDesde & "' and '" & iFecHasta & "' "
'        ''vgSql = vgSql & "and cod_moneda ='" & vlMoneda & "' "
'        vgSql = vgSql & "and g.num_poliza=p.num_poliza "
'        vgSql = vgSql & "and g.num_endoso=p.num_endoso "
'        vgSql = vgSql & "ORDER BY g.num_poliza "
        Set vgRs = vgConexionBD.Execute(vgSql)
        While Not (vgRs.EOF)
            ''vlNumPol = Trim(vgRs!Num_Poliza)
            vlFecPago = Trim(vgRs!Fec_Pago)
            
            vlVar1 = Format(Trim(clTipRegGS), "00000")
            vlVar2 = Format(Trim(vgRs!num_poliza), "0000000000")
            vlVar3 = flObtieneViaPago(Trim(vgRs!Cod_ViaPago))
            vlVar3 = Format(Trim(vlVar3), "00000")
            If Trim(vgRs!Fec_Pago) <> "" Then
                vlVar4 = DateSerial(Mid(vgRs!Fec_Pago, 1, 4), Mid(vgRs!Fec_Pago, 5, 2), Mid(vgRs!Fec_Pago, 7, 2))
                vlVar7 = Mid(vgRs!Fec_Pago, 1, 6)
            Else
                vlVar4 = Space(10)
                vlVar7 = Space(6)
            End If
            If (Len(Trim(vgRs!cod_monedagtosep)) <= 5) Then
                vlVar8 = Trim(vgRs!cod_monedagtosep) & Space(5 - Len(Trim(vgRs!cod_monedagtosep)))
            Else
                vlVar8 = Mid(Trim(vgRs!cod_monedagtosep), 1, 5)
            End If
            'vlMoneda = Trim(vlVar8)
            vlVar9 = Format(Trim(vgRs!Cod_TipPension), "00000")
            vlCodPen = Trim(vgRs!Cod_TipPension)
            vlVar10 = Format(Trim(clRamoContGS), "00000")
            If (Len(Trim(vgRs!num_idensolicita)) <= 12) Then
                vlVar11 = Format(Trim(vgRs!cod_tipoidensolicita), "00") & (Trim(vgRs!num_idensolicita) & Space(12 - Len(Trim(vgRs!num_idensolicita))))
            Else
                vlVar11 = Format(Trim(vgRs!cod_tipoidensolicita), "00") & Mid(Trim(vgRs!num_idensolicita), 1, 12)
            End If
            vlVar12 = fgFormarNombreCompleto(IIf(IsNull(vgRs!gls_nomsolicita), "", Trim(vgRs!gls_nomsolicita)), IIf(IsNull(vgRs!gls_nomsegsolicita), "", Trim(vgRs!gls_nomsegsolicita)), IIf(IsNull(vgRs!gls_patsolicita), "", Trim(vgRs!gls_patsolicita)), IIf(IsNull(vgRs!gls_matsolicita), "", Trim(vgRs!gls_matsolicita)))
            If (Len(Trim(vlVar12)) <= 60) Then
                vlVar12 = vlVar12 & Space(60 - Len(Trim(vlVar12)))
            Else
                vlVar12 = Mid(vlVar12, 1, 60)
            End If
            vlVar13 = Format(Trim(clFrecPagGS), "00000")
            vlVar19 = Format(Trim(vgRs!cod_afp), "00000")
            vlAfp = Trim(vgRs!cod_afp)
            vlVar23 = Format(Trim(clReaGS), "0000000000")
            vlVar61 = Format(Trim(clTipMovHabDes), "00000")
            vlVar62 = IIf(IsNull(vgRs!mto_pago), 0, Format(vgRs!mto_pago, "#0.00"))
            ''vlVar62 = String(18 - Len(Trim(vlVar62)), "0") & Trim(vlVar62)
            vlVar63 = flFormatNum18_2(vlVar62)
            vlVar62 = flFormatNum18_2(vlVar62)
            
            'vlVar67 = Trim(clTipPerNatGS) & Space(1 - Len(Trim(clTipPerNatGS)))
            If (Len(Trim(vgRs!COD_TIPOPERSONA)) <= 1) Then
                vlVar67 = Trim(vgRs!COD_TIPOPERSONA) & Space(1 - Len(Trim(vgRs!COD_TIPOPERSONA)))
            Else
                vlVar67 = Mid(Trim(vgRs!COD_TIPOPERSONA), 1, 1)
            End If
            vlLinea = (vlVar1) & (vlVar2) & (vlVar3) & (vlVar4) & (vlVar5) & _
                      (vlVar6) & (vlVar7) & (vlVar8) & (vlVar9) & (vlVar10) & _
                      (vlVar11) & (vlVar12) & (vlVar13) & (vlVar14) & (vlVar15) & _
                      (vlVar16) & (vlVar17) & (vlVar18) & (vlVar19) & (vlVar20) & _
                      (vlVar21) & (vlVar22) & (vlVar23) & (vlVar24) & (vlVar25) & _
                      (vlVar26) & (vlVar27) & (vlVar28) & (vlVar29) & (vlVar30) & _
                      (vlVar31) & (vlVar32) & (vlVar33) & (vlVar34) & (vlVar35) & _
                      (vlVar36) & (vlVar37) & (vlVar38) & (vlVar39) & (vlVar40) & _
                      (vlVar41) & (vlVar42) & (vlVar43) & (vlVar44) & (vlVar45) & _
                      (vlVar47) & (vlVar48) & (vlVar49) & (vlVar50) & _
                      (vlVar51) & (vlVar52) & (vlVar53) & (vlVar54) & (vlVar55) & _
                      (vlVar56) & (vlVar57) & (vlVar58) & (vlVar59) & (vlVar60) & _
                      (vlVar61) & (vlVar62) & (vlVar63) & (vlVar64) & (vlVar65) & _
                      (vlVar66) & (vlVar67)

            vlLinea = Replace(vlLinea, ",", ".")
            Print #1, vlLinea

            'Guarda el detalle de Pensión de la póliza informada
            vlMtoHab = IIf(IsNull(vgRs!mto_pago), 0, Format(vgRs!mto_pago, "#0.00"))
            Call flGrabaDetHabDes(clTipRegGS, clTipMovHabDes, vlVar2, vgRs!Num_Orden, vlFecPago, vlCodPen, vlAfp, vlVar12, vlMtoHab)
           
            vlNumCasosHabDes = vlNumCasosHabDes + 1
            vlMtoHabDes = vlMtoHabDes + vlMtoHab

            vgRs.MoveNext
        Wend
        vgRs.Close
        
        'Actualiza la cantidad de casos y mtos informados - Gto Sepelio
        Call flActEstadisticaHabDes(clTipRegGS, clTipMovHabDes, vlNumCasosHabDes, vlMtoHabDes)
        
        vgRs1.MoveNext
    Wend
    vgRs1.Close
    
Close #1
vlOpen = False

flExportarHabDes = True

Exit Function
Errores:
Screen.MousePointer = vbDefault
If Err.Number <> 0 Then
    If vlOpen Then
        Close #1
    End If
    MsgBox "Se ha producido el siguiente error : " & Err.Description, vbCritical, "Error"
End If
Resume
End Function

Private Function flExportarPerGar(iFecDesde As String, iFecHasta As String) As Boolean
Dim vlFecPago As String, vlCodPen As String, vlAfp As String
Dim vlMtoGar As Double
Dim vlNumOrd As Integer
Dim vgRs1 As ADODB.Recordset
On Error GoTo Errores

flExportarPerGar = False

    vlArchivo = vlArchivoCont 'LblArchivo
    
    Screen.MousePointer = 11
    
    Open vlArchivo For Output As #1

    Me.Refresh
    vlOpen = True
        
    Call flInicializaVar
        
    'Obtiene el nº de archivo a crear
    vlNumArchivo = flNumArchivoPerGar
    vlFecCrea = Format(Date, "yyyymmdd")
    vlHorCrea = Format(Time, "hhmmss")
    
    vlSql = "SELECT COD_ELEMENTO FROM MA_TPAR_TABCOD "
    vlSql = vlSql & "WHERE COD_TABLA='" & vgCodTabla_TipMon & "' "
    vlSql = vlSql & "AND (COD_SISTEMA IS NULL OR COD_SISTEMA<>'PP') "
    vlSql = vlSql & "ORDER BY COD_ELEMENTO "
    Set vgRs1 = vgConexionBD.Execute(vlSql)
    While Not (vgRs1.EOF)
        vlMoneda = Trim(vgRs1!COD_ELEMENTO)
        
        Call flEstadisticaPerGar(clTipRegPG, clTipMovSinPG)
        
        vlNumCasosPerGar = 0
        vlMtoPerGar = 0
        
        vgSql = ""
        vgSql = vgSql & "SELECT pg.num_poliza,pb.cod_viapago,pb.fec_pago,pg.cod_moneda,"
        vgSql = vgSql & "p.cod_tippension,pb.cod_tipoidenben,pb.num_idenben,"
        vgSql = vgSql & "b.gls_nomben,b.gls_nomsegben,b.gls_patben,b.gls_matben,"
        vgSql = vgSql & "p.cod_afp,pb.mto_pago,b.num_orden "
        vgSql = vgSql & "FROM PP_TMAE_PAGTERGARBEN PB, PP_TMAE_PAGTERGAR PG,"
        vgSql = vgSql & "PP_TMAE_POLIZA P,PP_TMAE_BEN B "
        vgSql = vgSql & "WHERE pb.fec_pago between '" & iFecDesde & "' and '" & iFecHasta & "' "
        vgSql = vgSql & "and p.cod_moneda = '" & vlMoneda & "' "
        vgSql = vgSql & "and pb.num_poliza=pg.num_poliza "
        vgSql = vgSql & "and pb.fec_pago=pg.fec_pago "
        vgSql = vgSql & "and pb.num_poliza=b.num_poliza "
        vgSql = vgSql & "and pb.num_orden=b.num_orden "
        vgSql = vgSql & "and pb.num_endoso=b.num_endoso "
        vgSql = vgSql & "and b.num_poliza=p.num_poliza "
        vgSql = vgSql & "and b.num_endoso=p.num_endoso "
        vgSql = vgSql & "ORDER BY pg.num_poliza "
        Set vgRs = vgConexionBD.Execute(vgSql)
        While Not (vgRs.EOF)
            
            vlNumOrd = CInt(Trim(vgRs!Num_Orden))
            vlFecPago = Trim(vgRs!Fec_Pago)
            
            vlVar1 = Format(Trim(clTipRegPG), "00000")
            vlVar2 = Format(Trim(vgRs!num_poliza), "0000000000")
            ''vlVar3 = Trim(vgRs!Cod_ViaPago) & Space(5 - Len(Trim(vgRs!Cod_ViaPago)))
            vlVar3 = flObtieneViaPago(Trim(vgRs!Cod_ViaPago))
            vlVar3 = Format(Trim(vlVar3), "00000")
            If Trim(vgRs!Fec_Pago) <> "" Then
                vlVar4 = DateSerial(Mid(vgRs!Fec_Pago, 1, 4), Mid(vgRs!Fec_Pago, 5, 2), Mid(vgRs!Fec_Pago, 7, 2))
                vlVar7 = Mid(vgRs!Fec_Pago, 1, 6)
            Else
                vlVar4 = Space(10)
                vlVar7 = Space(6)
            End If
            If (Len(Trim(vgRs!Cod_Moneda)) <= 5) Then
                vlVar8 = Trim(vgRs!Cod_Moneda) & Space(5 - Len(Trim(vgRs!Cod_Moneda)))
            Else
                vlVar8 = Mid(Trim(vgRs!Cod_Moneda), 1, 5)
            End If
            vlVar9 = Format(Trim(vgRs!Cod_TipPension), "00000")
            vlCodPen = Trim(vgRs!Cod_TipPension)
            vlVar10 = Format(Trim(clRamoContPG), "00000")
            If (Len(Trim(vgRs!Num_IdenBen)) <= 12) Then
                vlVar11 = Format(Trim(vgRs!Cod_TipoIdenBen), "00") & (Trim(vgRs!Num_IdenBen) & Space(12 - Len(Trim(vgRs!Num_IdenBen))))
            Else
                vlVar11 = Format(Trim(vgRs!Cod_TipoIdenBen), "00") & Mid(Trim(vgRs!Num_IdenBen), 1, 12)
            End If
            ''vlVar12 = Trim(vgRs!Gls_NomBen) & Trim(vgRs!Gls_NomSegBen) & Trim(vgRs!Gls_PatBen) & Trim(vgRs!Gls_MatBen)
            vlVar12 = fgFormarNombreCompleto(IIf(IsNull(vgRs!Gls_NomBen), "", Trim(vgRs!Gls_NomBen)), IIf(IsNull(vgRs!Gls_NomSegBen), "", Trim(vgRs!Gls_NomSegBen)), IIf(IsNull(vgRs!Gls_PatBen), "", Trim(vgRs!Gls_PatBen)), IIf(IsNull(vgRs!Gls_MatBen), "", Trim(vgRs!Gls_MatBen)))
            If (Len(Trim(vlVar12)) <= 60) Then
                vlVar12 = vlVar12 & Space(60 - Len(Trim(vlVar12)))
            Else
                vlVar12 = Mid(vlVar12, 1, 60)
            End If
            vlVar13 = Format(Trim(clFrecPagPG), "00000")
            vlVar19 = Format(Trim(vgRs!cod_afp), "00000")
            vlAfp = Trim(vgRs!cod_afp)
            vlVar23 = Format(Trim(clReaPG), "0000000000")
            vlVar61 = Format(Trim(clTipMovSinPG), "00000")
            vlVar62 = IIf(IsNull(vgRs!mto_pago), 0, Format(vgRs!mto_pago, "#0.00"))
            ''vlVar62 = String(18 - Len(Trim(vlVar62)), "0") & Trim(vlVar62)
            vlVar62 = flFormatNum18_2(vlVar62)
            vlVar67 = Trim(clTipPerNatPG) & Space(1 - Len(Trim(clTipPerNatPG)))
         
            
            vlLinea = (vlVar1) & (vlVar2) & (vlVar3) & (vlVar4) & (vlVar5) & _
                      (vlVar6) & (vlVar7) & (vlVar8) & (vlVar9) & (vlVar10) & _
                      (vlVar11) & (vlVar12) & (vlVar13) & (vlVar14) & (vlVar15) & _
                      (vlVar16) & (vlVar17) & (vlVar18) & (vlVar19) & (vlVar20) & _
                      (vlVar21) & (vlVar22) & (vlVar23) & (vlVar24) & (vlVar25) & _
                      (vlVar26) & (vlVar27) & (vlVar28) & (vlVar29) & (vlVar30) & _
                      (vlVar31) & (vlVar32) & (vlVar33) & (vlVar34) & (vlVar35) & _
                      (vlVar36) & (vlVar37) & (vlVar38) & (vlVar39) & (vlVar40) & _
                      (vlVar41) & (vlVar42) & (vlVar43) & (vlVar44) & (vlVar45) & _
                      (vlVar47) & (vlVar48) & (vlVar49) & (vlVar50) & _
                      (vlVar51) & (vlVar52) & (vlVar53) & (vlVar54) & (vlVar55) & _
                      (vlVar56) & (vlVar57) & (vlVar58) & (vlVar59) & (vlVar60) & _
                      (vlVar61) & (vlVar62) & (vlVar63) & (vlVar64) & (vlVar65) & _
                      (vlVar66) & (vlVar67)
                    
            vlLinea = Replace(vlLinea, ",", ".")
            Print #1, vlLinea
            
            'Guarda el detalle de Pensión de la póliza informada
            vlMtoGar = IIf(IsNull(vgRs!mto_pago), 0, Format(vgRs!mto_pago, "#0.00"))
            Call flGrabaDetPerGar(clTipRegPG, clTipMovSinPG, vlVar2, vlNumOrd, vlFecPago, vlCodPen, vlAfp, vlVar12, vlMtoGar)
            
            vlNumCasosPerGar = vlNumCasosPerGar + 1
            vlMtoPerGar = vlMtoPerGar + vlMtoGar
            
            vgRs.MoveNext
        Wend
        vgRs.Close
        
        'Actualiza la cantidad de casos y mtos informados
        Call flActEstadisticaPerGar(clTipRegPG, clTipMovSinPG, vlNumCasosPerGar, vlMtoPerGar)
        
        vgRs1.MoveNext
    
    Wend
    vgRs1.Close
        
Close #1
vlOpen = False

flExportarPerGar = True

Exit Function
Errores:
Screen.MousePointer = vbDefault
If Err.Number <> 0 Then
    If vlOpen Then
        Close #1
    End If
    MsgBox "Se ha producido el siguiente error : " & Err.Description, vbCritical, "Error"
End If
End Function

Private Sub Txt_Desde_GotFocus()
    Txt_Desde.SelStart = 0
    Txt_Desde.SelLength = Len(Txt_Desde)
End Sub

Private Sub Txt_Desde_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Txt_Hasta.SetFocus
End If
End Sub

Private Sub Txt_Desde_LostFocus()
    If Txt_Desde <> "" Then
        If (flValidaFecha(Txt_Desde) = False) Then
            Txt_Desde = ""
            Exit Sub
        End If
        Txt_Desde.Text = Format(CDate(Trim(Txt_Desde)), "yyyymmdd")
        Txt_Desde.Text = DateSerial(Mid((Txt_Desde.Text), 1, 4), Mid((Txt_Desde.Text), 5, 2), Mid((Txt_Desde.Text), 7, 2))
    End If
End Sub

Private Sub Txt_Hasta_GotFocus()
    Txt_Hasta.SelStart = 0
    Txt_Hasta.SelLength = Len(Txt_Hasta)
End Sub

Private Sub Txt_Hasta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CmdContable.SetFocus
End If
End Sub

Private Sub Txt_Hasta_LostFocus()
    If Txt_Hasta <> "" Then
        If (flValidaFecha(Txt_Hasta) = False) Then
            Txt_Hasta = ""
            Exit Sub
        End If
        If CDate(Txt_Desde) > CDate(Txt_Hasta) Then
            MsgBox "La Fecha Desde debe ser menor a la Fecha Hasta.", vbCritical, "Error de Datos"
            Txt_Hasta.SetFocus
            Exit Sub
        End If
        Txt_Hasta.Text = Format(CDate(Trim(Txt_Hasta)), "yyyymmdd")
        Txt_Hasta.Text = DateSerial(Mid((Txt_Hasta.Text), 1, 4), Mid((Txt_Hasta.Text), 5, 2), Mid((Txt_Hasta.Text), 7, 2))
    End If
End Sub

Function flValidaFecha(iFecha As String) As Boolean

    flValidaFecha = False

    If (Trim(iFecha) = "") Then
        Exit Function
    End If
    If Not IsDate(iFecha) Then
        Exit Function
    End If
    If (Year(CDate(iFecha)) < 1890) Then
        Exit Function
    End If

    flValidaFecha = True

End Function

Private Function flEstadisticaPagReg(iTipReg As String, iTipMov As String)
    
    vgSql = "INSERT INTO PP_TMAE_CONTABLEREGPAGO ("
    vgSql = vgSql & "num_archivo,cod_tipreg,cod_tipmov,"
    vgSql = vgSql & "cod_moneda,fec_desde,fec_hasta,"
    vgSql = vgSql & "cod_usuariocrea,fec_crea,hor_crea "
    vgSql = vgSql & ") VALUES ("
    vgSql = vgSql & " " & vlNumArchivo & ","
    vgSql = vgSql & "'" & Trim(iTipReg) & "',"
    vgSql = vgSql & "'" & Trim(iTipMov) & "',"
    vgSql = vgSql & "'" & vlMoneda & "',"
    vgSql = vgSql & "'" & vlFecDesde & "',"
    vgSql = vgSql & "'" & vlFecHasta & "',"
    vgSql = vgSql & "'" & (vgUsuario) & "',"
    vgSql = vgSql & "'" & Trim(vlFecCrea) & "',"
    vgSql = vgSql & "'" & Trim(vlHorCrea) & "')"
    vgConexionBD.Execute (vgSql)
    
    Lbl_NumArchivo.Caption = vlNumArchivo

End Function
Private Function flEstadisticaPagRegProvision(iTipReg As String, iTipMov As String)
    
    vgSql = "INSERT INTO PP_TMAE_CONTABLE_PROVISION ("
    vgSql = vgSql & "num_archivo,cod_tipreg,cod_tipmov,"
    vgSql = vgSql & "cod_moneda,fec_desde,fec_hasta,"
    vgSql = vgSql & "cod_usuariocrea,fec_crea,hor_crea "
    vgSql = vgSql & ") VALUES ("
    vgSql = vgSql & " " & vlNumArchivo & ","
    vgSql = vgSql & "'" & Trim(iTipReg) & "',"
    vgSql = vgSql & "'" & Trim(iTipMov) & "',"
    vgSql = vgSql & "'" & vlMoneda & "',"
    vgSql = vgSql & "'" & vlFecDesde & "',"
    vgSql = vgSql & "'" & vlFecHasta & "',"
    vgSql = vgSql & "'" & (vgUsuario) & "',"
    vgSql = vgSql & "'" & Trim(vlFecCrea) & "',"
    vgSql = vgSql & "'" & Trim(vlHorCrea) & "')"
    vgConexionBD.Execute (vgSql)
    
    Lbl_NumArchivo.Caption = vlNumArchivo

End Function

Private Function flEstadisticaGtoSep(iTipReg As String, iTipMov As String)

    vgSql = "INSERT INTO PP_TMAE_CONTABLEGTOSEP ("
    vgSql = vgSql & "num_archivo,cod_tipreg,cod_tipmov,"
    vgSql = vgSql & "cod_moneda,fec_desde,fec_hasta,"
    vgSql = vgSql & "cod_usuariocrea,fec_crea,hor_crea "
    vgSql = vgSql & ") VALUES ("
    vgSql = vgSql & " " & vlNumArchivo & ","
    vgSql = vgSql & "'" & Trim(iTipReg) & "',"
    vgSql = vgSql & "'" & Trim(iTipMov) & "',"
    vgSql = vgSql & "'" & Trim(vgMonedaCodOfi) & "',"
    vgSql = vgSql & "'" & vlFecDesde & "',"
    vgSql = vgSql & "'" & vlFecHasta & "',"
    vgSql = vgSql & "'" & (vgUsuario) & "',"
    vgSql = vgSql & "'" & Trim(vlFecCrea) & "',"
    vgSql = vgSql & "'" & Trim(vlHorCrea) & "')"
    vgConexionBD.Execute (vgSql)
    
    Lbl_NumArchivo.Caption = vlNumArchivo

End Function
Private Function flEstadisticaHabDes(iTipReg As String, iTipMov As String)

    vgSql = "INSERT INTO PP_TMAE_CONTABLEHABDES ("
    vgSql = vgSql & "num_archivo,cod_tipreg,cod_tipmov,"
    vgSql = vgSql & "cod_moneda,fec_desde,fec_hasta,"
    vgSql = vgSql & "cod_usuariocrea,fec_crea,hor_crea "
    vgSql = vgSql & ") VALUES ("
    vgSql = vgSql & " " & vlNumArchivo & ","
    vgSql = vgSql & "'" & Trim(iTipReg) & "',"
    vgSql = vgSql & "'" & Trim(iTipMov) & "',"
    vgSql = vgSql & "'" & Trim(vlMoneda) & "',"
    vgSql = vgSql & "'" & vlFecDesde & "',"
    vgSql = vgSql & "'" & vlFecHasta & "',"
    vgSql = vgSql & "'" & (vgUsuario) & "',"
    vgSql = vgSql & "'" & Trim(vlFecCrea) & "',"
    vgSql = vgSql & "'" & Trim(vlHorCrea) & "')"
    vgConexionBD.Execute (vgSql)
    
    Lbl_NumArchivo.Caption = vlNumArchivo

End Function
Private Function flEstadisticaPerGar(iTipReg As String, iTipMov As String)

    vgSql = "INSERT INTO PP_TMAE_CONTABLEPERGAR ("
    vgSql = vgSql & "num_archivo,cod_tipreg,cod_tipmov,"
    vgSql = vgSql & "cod_moneda,fec_desde,fec_hasta,"
    vgSql = vgSql & "cod_usuariocrea,fec_crea,hor_crea "
    vgSql = vgSql & ") VALUES ("
    vgSql = vgSql & " " & vlNumArchivo & ","
    vgSql = vgSql & "'" & Trim(iTipReg) & "',"
    vgSql = vgSql & "'" & Trim(iTipMov) & "',"
    vgSql = vgSql & "'" & vlMoneda & "',"
    vgSql = vgSql & "'" & vlFecDesde & "',"
    vgSql = vgSql & "'" & vlFecHasta & "',"
    vgSql = vgSql & "'" & (vgUsuario) & "',"
    vgSql = vgSql & "'" & Trim(vlFecCrea) & "',"
    vgSql = vgSql & "'" & Trim(vlHorCrea) & "')"
    vgConexionBD.Execute (vgSql)
    
    Lbl_NumArchivo.Caption = vlNumArchivo
    
End Function

Private Function flActEstadisticaPagReg(iTipReg As String, iTipMov As String, iNumCasos As Long, iMtoPago As Double)

    vgSql = "UPDATE PP_TMAE_CONTABLEREGPAGO set "
    vgSql = vgSql & "num_casos = " & iNumCasos & ","
    vgSql = vgSql & "mto_pago = " & str(iMtoPago) & " "
    vgSql = vgSql & "WHERE num_archivo = " & vlNumArchivo & " "
    vgSql = vgSql & "and cod_tipreg ='" & Trim(iTipReg) & "' "
    vgSql = vgSql & "and cod_tipmov ='" & Trim(iTipMov) & "' "
    vgSql = vgSql & "and cod_moneda ='" & vlMoneda & "' "
    vgConexionBD.Execute (vgSql)
          
End Function
Private Function flActEstadisticaPagRegProvision(iTipReg As String, iTipMov As String, iNumCasos As Long, iMtoPago As Double)

    vgSql = "UPDATE PP_TMAE_CONTABLE_PROVISION set "
    vgSql = vgSql & "num_casos = " & iNumCasos & ","
    vgSql = vgSql & "mto_pago = " & str(iMtoPago) & " "
    vgSql = vgSql & "WHERE num_archivo = " & vlNumArchivo & " "
    vgSql = vgSql & "and cod_tipreg ='" & Trim(iTipReg) & "' "
    vgSql = vgSql & "and cod_tipmov ='" & Trim(iTipMov) & "' "
    vgSql = vgSql & "and cod_moneda ='" & vlMoneda & "' "
    vgConexionBD.Execute (vgSql)
          
End Function

Private Function flActEstadisticaGtoSep(iTipReg As String, iTipMov As String, iNumCasos As Long, iMtoPago As Double)

    vgSql = "UPDATE PP_TMAE_CONTABLEGTOSEP set "
    vgSql = vgSql & "num_casos = " & iNumCasos & ","
    vgSql = vgSql & "mto_pago = " & str(iMtoPago) & " "
    vgSql = vgSql & "WHERE num_archivo = " & vlNumArchivo & " "
    vgSql = vgSql & "and cod_tipreg ='" & Trim(iTipReg) & "' "
    vgSql = vgSql & "and cod_tipmov ='" & Trim(iTipMov) & "' "
    vgSql = vgSql & "and cod_moneda ='" & Trim(vgMonedaCodOfi) & "' "
    vgConexionBD.Execute (vgSql)
          
End Function
Private Function flActEstadisticaHabDes(iTipReg As String, iTipMov As String, iNumCasos As Long, iMtoPago As Double)

    vgSql = "UPDATE PP_TMAE_CONTABLEHABDES set "
    vgSql = vgSql & "num_casos = " & iNumCasos & ","
    vgSql = vgSql & "mto_pago = " & str(iMtoPago) & " "
    vgSql = vgSql & "WHERE num_archivo = " & vlNumArchivo & " "
    vgSql = vgSql & "and cod_tipreg ='" & Trim(iTipReg) & "' "
    vgSql = vgSql & "and cod_tipmov ='" & Trim(iTipMov) & "' "
    vgSql = vgSql & "and cod_moneda ='" & Trim(vgMonedaCodOfi) & "' "
    vgConexionBD.Execute (vgSql)
          
End Function
Private Function flActEstadisticaPerGar(iTipReg As String, iTipMov As String, iNumCasos As Long, iMtoPago As Double)

    vgSql = "UPDATE PP_TMAE_CONTABLEPERGAR set "
    vgSql = vgSql & "num_casos = " & iNumCasos & ","
    vgSql = vgSql & "mto_pago = " & str(iMtoPago) & " "
    vgSql = vgSql & "WHERE num_archivo = " & vlNumArchivo & " "
    vgSql = vgSql & "and cod_tipreg ='" & Trim(iTipReg) & "' "
    vgSql = vgSql & "and cod_tipmov ='" & Trim(iTipMov) & "' "
    vgSql = vgSql & "and cod_moneda ='" & vlMoneda & "' "
    vgConexionBD.Execute (vgSql)
          
End Function

'FUNCION QUE GUARDA EL NOMBRE DEL FORMULARIO DEL QUE SE LLAMO A LA FUNCION
Function flInicio(iNomForm)
    vgNomForm = iNomForm
    Call Form_Load
End Function

Private Function flInicializaVar()

    vlVar1 = String(5, "0")     'Tipo de Registro
    vlVar2 = String(10, "0")    'POLIZA
    vlVar3 = String(5, "0")     'SUCURSAL
    vlVar3_Res = String(5, "0")     'SUCURSAL
    vlVar4 = Space(10)          'VIGENCIA "DESDE" DE LA POLIZA
    vlVar5 = Space(10)          '--VIGENCIA "HASTA" DE LA POLIZA
    vlVar6 = Space(10)          '--VIGENCIA 'DESDE' ORIGINAL
    vlVar7 = String(6, "0")     'FECHA CONTABLE(MES / ANO)
    vlVar8 = Space(5)           'MONEDA DEL MOVIMIENTO
    vlVar9 = String(5, "0")     'COBERTURA
    vlVar10 = String(5, "0")    'RAMO CONTABLE
    vlVar11 = Space(14)         'CONTRATANTE RUT
    vlVar11_Res = Space(14)         'CONTRATANTE RUT
    vlVar12 = Space(60)         'CONTRATANTE NOMBRE
    vlVar12_Res = Space(60)         'CONTRATANTE NOMBRE
    vlVar13 = String(5, "0")    'FRECUENCIA DE PAGO (INMEDIATA)
    vlVar14 = Space(14)         '--INTERMERDIARIO GTO DE COBRANZA
    vlVar14_Res = Space(14)         '--INTERMERDIARIO GTO DE COBRANZA
    vlVar15 = String(5, "0")   '--REGISTRO EN TRATAMIENTO
    vlVar16 = Space(10)         '--FECHA EFECTO REGISTRO (MOV)
    vlVar17 = Space(60)         'NOMBRE
    vlVar18 = Space(14)         'RUT
    vlVar19 = String(5, "0")    'SUCURSAL
    vlVar19_Res = String(5, "0")    'SUCURSAL
    vlVar20 = Space(60)         '--NOMBRE INTERMEDIARIO
    vlVar21 = Space(14)         '--RUT INNTERMEDIARIO
    vlVar22 = String(5, "0")    'TIPO DE INTERMEDIARIO
    vlVar23 = String(10, "0")   '--REASEGURADOR
    vlVar24 = Space(40)         'NACIONALIDAD REASEGURADOR
    vlVar25 = String(5, "0")    'CONTRATO DE REASEGURO
    vlVar26 = Space(6)          'TIPO DE REASEGURO
    vlVar27 = String(10, "0")   'NUMERO DE SINIESTRO
    vlVar28 = Space(6)          'ESTADO DEL SINIESTRO
    vlVar29 = String(5, "0")    'TIPO MOVIMIENTO
    vlVar30 = String(10, "0")   'NUMERO DE MOVIMIENTO
    vlVar31 = String(18, "0")   'MONTO EXENTO PRIMA
    vlVar32 = String(18, "0")   '--MONTO AFECTO PRIMA
    vlVar33 = String(18, "0")   '--MONTO IGV PRIMA
    vlVar34 = String(18, "0")   '--MONTO BRUTO PRIMA
    vlVar35 = String(18, "0")   '--MONTO NETO PRIMA DEVENGADA
    vlVar36 = String(18, "0")   '--CAPITALES ASEGURADOS
    vlVar37 = String(5, "0")    '--ORIGEN DEL RECIBO
    vlVar38 = String(5, "0")    '--TIPO DE MOVIMIENTO
    vlVar39 = String(18, "0")   '--MONTO PRIMA CEDIDA ANTES DSCTO
    vlVar40 = String(18, "0")   '--MONTO DESC. POR PRIMA CEDIDA
    vlVar41 = String(4, "0")    '--MONTO IMPUESTO 2%
    vlVar42 = String(18, "0")   '--MONTO EXCESO DE PERDIDA
    vlVar43 = String(18, "0")   '--CAPITALES CEDIDOS
    vlVar44 = Space(6)          '--TIPO DE RESERVA
    vlVar45 = String(18, "0")   '--MONTO RESERVA MATEMATICA
    vlVar47 = String(5, "0")    '-- % DE COMSION SOBRE LA PRIMA
    vlVar48 = Space(6)          '--TIPO DE COMISION
    vlVar49 = String(18, "0")   '--MONTO COMISION NETA
    vlVar50 = String(18, "0")   '--MONTO IGV COMISION
    vlVar51 = String(18, "0")   '--MONTO BRUTO COMISION
    vlVar52 = String(5, "0")    '--PERIODO DE GRACIA
    vlVar53 = String(18, "0")   '--MONTO NETO COMISION
    vlVar54 = Space(6)          '--ESQUEMA DE PAGO
    vlVar55 = String(10, "0")   '--fecha desde
    vlVar56 = String(10, "0")   '--fecha Hasta
    vlVar57 = String(5, "0")    '--RAMO
    vlVar58 = String(5, "0")    '--PRODUCTO
    vlVar59 = String(10, "0")   '--POLIZA
    vlVar60 = Space(14)         '--RUT DEL CLIENTE
    vlVar61 = String(5, "0")    '--TIPO DE MOVIMIENTO SINIESTRO
    vlVar62 = String(18, "0")   '--MONTO
    vlVar63 = String(18, "0")   '--MONTO CEDIDO EN EL MES
    vlVar63_Res = String(18, "0")   '--MONTO CEDIDO EN EL MES
    vlVar64 = String(5, "0")    '-- % DE COMISION DE GASTOS DE COB
    vlVar64_Res = String(5, "0")    '-- % DE COMISION DE GASTOS DE COB
    vlVar65 = String(18, "0")   '--MTO. GASTOS DE COB. PRIMA REC.
    vlVar65_Res = String(18, "0")   '--MTO. GASTOS DE COB. PRIMA REC.
    vlVar66 = String(18, "0")   '--MTO. GASTOS DE COB. PRIMA DEV.
    vlVar66_Res = String(18, "0")   '--MTO. GASTOS DE COB. PRIMA DEV.
    vlVar67 = Space(1)          '--Tipo de persona (jurídico / natural)
    vlVar67_Res = Space(1)          '--Tipo de persona (jurídico / natural)

End Function

Private Function flObtieneViaPago(iViaPago As String) As String

    vgQuery = "SELECT COD_ADICIONAL "
    vgQuery = vgQuery & "FROM MA_TPAR_TABCOD "
    vgQuery = vgQuery & "WHERE COD_TABLA='VPG' "
    vgQuery = vgQuery & "AND COD_ELEMENTO = '" & iViaPago & "' "
    Set vgRs3 = vgConexionBD.Execute(vgQuery)
    If Not (vgRs3.EOF) Then
        flObtieneViaPago = IIf(IsNull(vgRs3!cod_adicional), 0, (vgRs3!cod_adicional))
    Else
        flObtieneViaPago = 0
    End If

End Function

Private Function flObtieneCodSalud(iSalud As String) As String

    vgQuery = "SELECT COD_SISTEMA "
    vgQuery = vgQuery & "FROM MA_TPAR_TABCOD "
    vgQuery = vgQuery & "WHERE COD_TABLA='IS' "
    vgQuery = vgQuery & "AND COD_ELEMENTO = '" & iSalud & "' "
    Set vgRs3 = vgConexionBD.Execute(vgQuery)
    If Not (vgRs3.EOF) Then
        flObtieneCodSalud = IIf(IsNull(vgRs3!COD_SISTEMA), 0, (vgRs3!COD_SISTEMA))
    Else
        flObtieneCodSalud = 0
    End If

End Function

Private Function flObtieneTipPer(iViaPago As String) As String
'Devuelve el tipo de persona (Juridico / Natural) a la cual se realiza el pago
'"04"= Tranferencia AFP

    If (iViaPago = "04") Then
        flObtieneTipPer = clTipPerJurPR
    Else
        flObtieneTipPer = clTipPerNatPR
    End If
    
End Function

Private Function flFormatNum18_2(iNumero As String) As String
Dim vlNum As String
    
    vlNum = Format(iNumero, "#00000000000000.00")

    If (CDbl(vlNum) < 0) Then
        vlNum = Mid(vlNum, 1, 1) & "0" & Mid(vlNum, 2, 14) & Mid(vlNum, 17, 2)
    Else
        vlNum = "00" & Mid(vlNum, 1, 14) & Mid(vlNum, 16, 2)
    End If
    
    flFormatNum18_2 = vlNum

End Function

Private Function flFormatNum5_2(iNumero As String) As String
Dim vlNum As String
    
    vlNum = Format(iNumero, "#000.00")

    If (CDbl(vlNum) < 0) Then
        ''vlNum = Mid(vlNum, 1, 1) & Mid(vlNum, 2, 3) & Mid(vlNum, 6, 2)
        vlNum = Mid(vlNum, 2, 3) & Mid(vlNum, 6, 2)
    Else
        ''vlNum = "00" & Mid(vlNum, 1, 3) & Mid(vlNum, 5, 2)
        vlNum = Mid(vlNum, 1, 3) & Mid(vlNum, 5, 2)
    End If
    
    flFormatNum5_2 = vlNum

End Function

Private Function flGrabaDetPagReg(iTipReg As String, iTipMov As String, iNumPol As String, iNumOrd As Integer, iFecPago As String, iPension As String, iAFP As String, iNombre As String, iMtoPago As Double, Optional iNumIden As String, Optional iViaPago As String)

    vgSql = "INSERT INTO PP_TMAE_CONTABLEDETREGPAGO ("
    vgSql = vgSql & "num_archivo,cod_tipreg,cod_tipmov,"
    vgSql = vgSql & "num_poliza,num_orden,fec_pago,cod_moneda,"
    vgSql = vgSql & "cod_tippension,cod_afp,"
    If (Trim(iNombre) <> "") Then vgSql = vgSql & "gls_nombre,"
    vgSql = vgSql & "mto_pago "
    If (Trim(iNumIden) <> "") Then vgSql = vgSql & ",num_idenreceptor"
    If (Trim(iViaPago) <> "") Then vgSql = vgSql & ",cod_viapago"
    vgSql = vgSql & ") VALUES ("
    vgSql = vgSql & " " & vlNumArchivo & ","
    vgSql = vgSql & "'" & Trim(iTipReg) & "',"
    vgSql = vgSql & "'" & Trim(iTipMov) & "',"
    vgSql = vgSql & "'" & iNumPol & "',"
    vgSql = vgSql & " " & iNumOrd & ","
    vgSql = vgSql & "'" & iFecPago & "',"
    vgSql = vgSql & "'" & Trim(vlMoneda) & "',"
    vgSql = vgSql & "'" & Trim(iPension) & "',"
    vgSql = vgSql & "'" & Trim(iAFP) & "',"
    If (Trim(iNombre) <> "") Then vgSql = vgSql & "'" & Trim(iNombre) & "',"
    vgSql = vgSql & " " & str(iMtoPago) & " "
    If (Trim(iNumIden) <> "") Then vgSql = vgSql & ",'" & Trim(iNumIden) & "'"
    If (Trim(iViaPago) <> "") Then vgSql = vgSql & ",'" & Trim(iViaPago) & "'"
    vgSql = vgSql & ")"
    vgConexionBD.Execute (vgSql)

End Function

Private Function flGrabaDetPagRegProvision(iTipReg As String, iTipMov As String, iNumPol As String, iNumOrd As Integer, iFecPago As String, iPension As String, iAFP As String, iNombre As String, iMtoPago As Double, Optional iNumIden As String, Optional iViaPago As String)

    vgSql = "INSERT INTO PP_TMAE_CONTABLEDET_PROVISION ("
    vgSql = vgSql & "num_archivo,cod_tipreg,cod_tipmov,"
    vgSql = vgSql & "num_poliza,num_orden,fec_pago,cod_moneda,"
    vgSql = vgSql & "cod_tippension,cod_afp,"
    If (Trim(iNombre) <> "") Then vgSql = vgSql & "gls_nombre,"
    vgSql = vgSql & "mto_pago "
    If (Trim(iNumIden) <> "") Then vgSql = vgSql & ",num_idenreceptor"
    If (Trim(iViaPago) <> "") Then vgSql = vgSql & ",cod_viapago"
    vgSql = vgSql & ") VALUES ("
    vgSql = vgSql & " " & vlNumArchivo & ","
    vgSql = vgSql & "'" & Trim(iTipReg) & "',"
    vgSql = vgSql & "'" & Trim(iTipMov) & "',"
    vgSql = vgSql & "'" & iNumPol & "',"
    vgSql = vgSql & " " & iNumOrd & ","
    vgSql = vgSql & "'" & iFecPago & "',"
    vgSql = vgSql & "'" & Trim(vlMoneda) & "',"
    vgSql = vgSql & "'" & Trim(iPension) & "',"
    vgSql = vgSql & "'" & Trim(iAFP) & "',"
    If (Trim(iNombre) <> "") Then vgSql = vgSql & "'" & Trim(iNombre) & "',"
    vgSql = vgSql & " " & str(iMtoPago) & " "
    If (Trim(iNumIden) <> "") Then vgSql = vgSql & ",'" & Trim(iNumIden) & "'"
    If (Trim(iViaPago) <> "") Then vgSql = vgSql & ",'" & Trim(iViaPago) & "'"
    vgSql = vgSql & ")"
    vgConexionBD.Execute (vgSql)

End Function

Private Function flGrabaDetGtoSep(iTipReg As String, iTipMov As String, iNumPol As String, iFecPago As String, iPension As String, iAFP As String, iNombre As String, iMtoPago As Double)

    vgSql = "INSERT INTO PP_TMAE_CONTABLEDETGTOSEP ("
    vgSql = vgSql & "num_archivo,cod_tipreg,cod_tipmov,"
    vgSql = vgSql & "num_poliza,fec_pago,cod_moneda,"
    vgSql = vgSql & "cod_tippension,cod_afp,"
    If (Trim(iNombre) <> "") Then vgSql = vgSql & "gls_nombre,"
    vgSql = vgSql & "mto_pago "
    vgSql = vgSql & ") VALUES ("
    vgSql = vgSql & " " & vlNumArchivo & ","
    vgSql = vgSql & "'" & Trim(iTipReg) & "',"
    vgSql = vgSql & "'" & Trim(iTipMov) & "',"
    vgSql = vgSql & "'" & iNumPol & "',"
    vgSql = vgSql & "'" & iFecPago & "',"
    vgSql = vgSql & "'" & Trim(vlMoneda) & "',"
    vgSql = vgSql & "'" & Trim(iPension) & "',"
    vgSql = vgSql & "'" & Trim(iAFP) & "',"
    If (Trim(iNombre) <> "") Then vgSql = vgSql & "'" & Trim(iNombre) & "',"
    vgSql = vgSql & " " & str(iMtoPago) & ")"
    vgConexionBD.Execute (vgSql)

End Function

Private Function flGrabaDetHabDes(iTipReg As String, iTipMov As String, iNumPol As String, iOrden As Long, iFecPago As String, iPension As String, iAFP As String, iNombre As String, iMtoPago As Double)

    vgSql = "INSERT INTO PP_TMAE_CONTABLEDETHABDES ("
    vgSql = vgSql & "num_archivo,cod_tipreg,cod_tipmov,"
    vgSql = vgSql & "num_poliza,orden,fec_pago,cod_moneda,"
    vgSql = vgSql & "cod_tippension,cod_afp,"
    If (Trim(iNombre) <> "") Then vgSql = vgSql & "gls_nombre,"
    vgSql = vgSql & "mto_pago "
    vgSql = vgSql & ") VALUES ("
    vgSql = vgSql & " " & vlNumArchivo & ","
    vgSql = vgSql & "'" & Trim(iTipReg) & "',"
    vgSql = vgSql & "'" & Trim(iTipMov) & "',"
    vgSql = vgSql & "'" & iNumPol & "',"
    vgSql = vgSql & "" & iOrden & ","
    vgSql = vgSql & "'" & iFecPago & "',"
    vgSql = vgSql & "'" & Trim(vlMoneda) & "',"
    vgSql = vgSql & "'" & Trim(iPension) & "',"
    vgSql = vgSql & "'" & Trim(iAFP) & "',"
    If (Trim(iNombre) <> "") Then vgSql = vgSql & "'" & Trim(iNombre) & "',"
    vgSql = vgSql & " " & str(iMtoPago) & ")"
    vgConexionBD.Execute (vgSql)

End Function

Private Function flGrabaDetPerGar(iTipReg As String, iTipMov As String, iNumPol As String, iNumOrd As Integer, iFecPago As String, iPension As String, iAFP As String, iNombre As String, iMtoPago As Double)

    vgSql = "INSERT INTO PP_TMAE_CONTABLEDETPERGAR ("
    vgSql = vgSql & "num_archivo,cod_tipreg,cod_tipmov,"
    vgSql = vgSql & "num_poliza,num_orden,fec_pago,cod_moneda,"
    vgSql = vgSql & "cod_tippension,cod_afp,"
    If (Trim(iNombre) <> "") Then vgSql = vgSql & "gls_nombre,"
    vgSql = vgSql & "mto_pago "
    vgSql = vgSql & ") VALUES ("
    vgSql = vgSql & " " & vlNumArchivo & ","
    vgSql = vgSql & "'" & Trim(iTipReg) & "',"
    vgSql = vgSql & "'" & Trim(iTipMov) & "',"
    vgSql = vgSql & "'" & iNumPol & "',"
    vgSql = vgSql & " " & iNumOrd & ","
    vgSql = vgSql & "'" & iFecPago & "',"
    vgSql = vgSql & "'" & Trim(vlMoneda) & "',"
    vgSql = vgSql & "'" & Trim(iPension) & "',"
    vgSql = vgSql & "'" & Trim(iAFP) & "',"
    If (Trim(iNombre) <> "") Then vgSql = vgSql & "'" & Trim(iNombre) & "',"
    vgSql = vgSql & " " & str(iMtoPago) & ")"
    vgConexionBD.Execute (vgSql)

End Function

Private Function flMostrarOption()
    Opt_Resumen.Left = 240
    Opt_Detalle.Left = 1560
    Opt_Detalle.Caption = "Detalle Mov."
    Opt_DetAfpMon.Visible = True
    Opt_DetProdMon.Visible = True
    Opt_DetMtos.Visible = True
End Function

Private Function flOcultarOption()
    Opt_Resumen.Left = 2640
    Opt_Detalle.Left = 5160
    Opt_Detalle.Caption = "Detalle por Movimiento"
    Opt_DetAfpMon.Visible = False
    Opt_DetProdMon.Visible = False
    Opt_DetMtos.Visible = False
    Opt_DetPendAcumulado.Visible = False
End Function

Private Function flActDetallePagReg(iTipReg As String, iTipMov As String, iNumPol As String, iNumOrd As Integer, iMtoPenTotal As Double, iMtoSalud As Double, iMtoRet As Double)

    vgSql = "UPDATE PP_TMAE_CONTABLEDETREGPAGO set "
    vgSql = vgSql & "mto_pentot = " & str(iMtoPenTotal) & ","
    vgSql = vgSql & "mto_salud = " & str(iMtoSalud) & ","
    vgSql = vgSql & "mto_retencion = " & str(iMtoRet) & " "
    vgSql = vgSql & "WHERE num_archivo = " & vlNumArchivo & " "
    vgSql = vgSql & "and cod_tipreg ='" & Trim(iTipReg) & "' "
    vgSql = vgSql & "and cod_tipmov ='" & Trim(iTipMov) & "' "
    vgSql = vgSql & "and num_poliza ='" & iNumPol & "' "
    vgSql = vgSql & "and num_orden = " & iNumOrd & " "
    vgConexionBD.Execute (vgSql)
          
End Function
Private Function flActDetallePagRegProvision(iTipReg As String, iTipMov As String, iNumPol As String, iNumOrd As Integer, iMtoPenTotal As Double, iMtoSalud As Double, iMtoRet As Double)

    vgSql = "UPDATE PP_TMAE_CONTABLEDET_PROVISION set "
    vgSql = vgSql & "mto_pentot = " & str(iMtoPenTotal) & ","
    vgSql = vgSql & "mto_salud = " & str(iMtoSalud) & ","
    vgSql = vgSql & "mto_retencion = " & str(iMtoRet) & " "
    vgSql = vgSql & "WHERE num_archivo = " & vlNumArchivo & " "
    vgSql = vgSql & "and cod_tipreg ='" & Trim(iTipReg) & "' "
    vgSql = vgSql & "and cod_tipmov ='" & Trim(iTipMov) & "' "
    vgSql = vgSql & "and num_poliza ='" & iNumPol & "' "
    vgSql = vgSql & "and num_orden = " & iNumOrd & " "
    vgConexionBD.Execute (vgSql)
          
End Function
Private Sub CargaReporteContableDetalleMontos()
    Dim rs As ADODB.Recordset
    Dim objRep As New ClsReporte
    Dim cadena As String
    
    On Error GoTo mierror
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

    cadena = "select C.NUM_ARCHIVO,C.FEC_DESDE,C.FEC_HASTA,C.COD_USUARIOCREA,C.FEC_CREA,C.HOR_CREA," & _
            " d.COD_TIPREG , d.COD_TIPMOV,d.NUM_POLIZA,D.FEC_PAGO,d.Cod_Moneda,D.COD_TIPPENSION, d.COD_AFP,GLS_NOMBRE, d.MTO_PAGO,D.MTO_PENTOT,D.MTO_SALUD,D.MTO_RETENCION" & _
            " from PP_TMAE_CONTABLEREGPAGO C INNER JOIN PP_TMAE_CONTABLEDETREGPAGO D ON C.NUM_ARCHIVO=D.NUM_ARCHIVO and C.COD_TIPREG=D.COD_TIPREG AND" & _
            " C.COD_TIPMOV=D.COD_TIPMOV AND C.COD_MONEDA=D.COD_MONEDA where D.NUM_ARCHIVO='" & Trim(Lbl_NumArchivo) & "' AND D.COD_TIPMOV = '" & Trim(clTipMovSinLPR) & "'"
         
    rs.Open cadena, vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF Then
        MsgBox "No hay datos para mostrar", vbExclamation, "Pago de pensiones"
        Exit Sub
    End If
    
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PP_Rpt_ContableDetPagRegMontos.rpt"), ".RPT", ".TTX"), 1)
    
    If objRep.CargaReporte(strRpt & "", "PP_Rpt_ContableDetPagRegMontos.rpt", "Informe Detallado de Archivo Contable", rs, True, _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("NombreSistema", vgNombreSistema), _
                            ArrFormulas("NombreSubSistema", vgNombreSubSistema)) = False Then
                            
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
    End If
    
    Exit Sub
mierror:
    MsgBox "No se pudo cargar el reporte", vbInformation, "Pensiones"
    
End Sub
Private Sub CargaReporteContableDetalleMontosProvision()
    Dim rs As ADODB.Recordset
    Dim objRep As New ClsReporte
    Dim cadena As String
    
    On Error GoTo mierror
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

    cadena = "select (CASE WHEN B.num_poliza IS NULL THEN 'N' ELSE 'S' END)AS ESTADO,X.NUM_ARCHIVO,X.FEC_DESDE,X.FEC_HASTA," & _
            " X.COD_USUARIOCREA,X.FEC_CREA,X.HOR_CREA, A.COD_TIPREG , A.COD_TIPMOV,A.NUM_POLIZA,B.FEC_PAGO,A.Cod_Moneda," & _
            " A.COD_TIPPENSION,A.COD_AFP,A.GLS_NOMBRE, A.MTO_PAGO,A.MTO_PENTOT,A.MTO_SALUD,A.MTO_RETENCION" & _
            " FROM PP_TMAE_CONTABLE_PROVISION X INNER JOIN PP_TMAE_CONTABLEDET_PROVISION A ON X.NUM_ARCHIVO = A.NUM_ARCHIVO " & _
            " And X.COD_TIPREG = A.COD_TIPREG And X.COD_TIPMOV = A.COD_TIPMOV And X.Cod_Moneda = A.Cod_Moneda" & _
            " left join PP_TMAE_CONTABLEDETREGPAGO B ON A.NUM_ARCHIVO = B.NUM_ARCHIVO And A.COD_TIPREG = B.COD_TIPREG And" & _
            " A.COD_TIPMOV = B.COD_TIPMOV And A.Cod_Moneda = B.Cod_Moneda AND A.NUM_POLIZA=B.NUM_POLIZA where A.NUM_ARCHIVO='" & Trim(Lbl_NumArchivo) & "' AND A.COD_TIPMOV = '" & Trim(clTipMovSinLPR) & "'" & _
            " GROUP BY X.NUM_ARCHIVO,X.FEC_DESDE,X.FEC_HASTA," & _
            " X.COD_USUARIOCREA,X.FEC_CREA,X.HOR_CREA, A.COD_TIPREG , A.COD_TIPMOV,A.NUM_POLIZA,B.FEC_PAGO,A.Cod_Moneda," & _
            " A.COD_TIPPENSION,A.COD_AFP,A.GLS_NOMBRE, A.MTO_PAGO,A.MTO_PENTOT,A.MTO_SALUD,A.MTO_RETENCION,B.num_poliza"
         
    rs.Open cadena, vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF Then
        MsgBox "No hay datos para mostrar", vbExclamation, "Pago de pensiones"
        Exit Sub
    End If
    
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PP_Rpt_ContableDetPagRegMontosProv.rpt"), ".RPT", ".TTX"), 1)
    
    If objRep.CargaReporte(strRpt & "", "PP_Rpt_ContableDetPagRegMontosProv.rpt", "Informe Detallado de Archivo Contable de Provisión", rs, True, _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("NombreSistema", vgNombreSistema), _
                            ArrFormulas("NombreSubSistema", vgNombreSubSistema)) = False Then
                            
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
    End If
    
    Exit Sub
mierror:
    MsgBox "No se pudo cargar el reporte", vbInformation, "Pensiones"
    
End Sub
Private Sub flImpDetMontos()
Dim vlArchivo As String
Err.Clear
On Error GoTo Errores1
   
    Screen.MousePointer = 11
    
    If (Trim(Lbl_NumArchivo) = "") Then
        MsgBox "Debe seleccionar un Periodo a Imprimir.", vbInformation, "Falta Información"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    If (vgNomForm = "PagRec") Then
        Call CargaReporteContableDetalleMontos
        Call CargaReporteContableDetalleMontosProvision
        Exit Sub
'        vlArchivo = strRpt & "PP_Rpt_ContableDetPagRegMontos.rpt"   '\Reportes
'        vgQuery = "{PP_TMAE_CONTABLEDETREGPAGO.NUM_ARCHIVO} = " & Trim(Lbl_NumArchivo) & "and "
'        vgQuery = vgQuery & "{PP_TMAE_CONTABLEDETREGPAGO.COD_TIPMOV} = '" & Trim(clTipMovSinLPR) & "'"

''    ElseIf (vgNomForm = "GtoSep") Then
''        vlArchivo = strRpt & "PP_Rpt_ContableResGtoSep.rpt"   '\Reportes
''        vgQuery = "{PP_TMAE_CONTABLEDETREGPAGO.NUM_ARCHIVO} = " & Trim(Lbl_NumArchivo)
''    Else
''        vlArchivo = strRpt & "PP_Rpt_ContableResPerGar.rpt"   '\Reportes
''        vgQuery = "{PP_TMAE_CONTABLEDETREGPAGO.NUM_ARCHIVO} = " & Trim(Lbl_NumArchivo)
    End If
    
    If Not fgExiste(vlArchivo) Then     ', vbNormal
        MsgBox "Archivo de Reporte de Resumen de Archivo Contable no se encuentra en el directorio de la Aplicación.", 16, "Archivo no encontrado"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    Rpt_General.Reset
    Rpt_General.WindowState = crptMaximized
    Rpt_General.ReportFileName = vlArchivo
    Rpt_General.Connect = vgRutaDataBase
    Rpt_General.Destination = crptToWindow
    Rpt_General.SelectionFormula = ""
    Rpt_General.SelectionFormula = vgQuery
    
    Rpt_General.Formulas(0) = ""
    Rpt_General.Formulas(1) = ""
    Rpt_General.Formulas(2) = ""
    Rpt_General.Formulas(3) = ""
    
    Rpt_General.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
    Rpt_General.Formulas(1) = "NombreSistema = '" & vgNombreSistema & "'"
    Rpt_General.Formulas(2) = "NombreSubSistema = '" & vgNombreSubSistema & "'"

    Rpt_General.WindowTitle = "Informe Detalle Archivo Contable"
    Rpt_General.Action = 1
    
    Screen.MousePointer = 0
   
Exit Sub
Errores1:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " " & "Numero de Error : " & Err.Number & Err.HelpFile & Err.HelpContext, vbExclamation, "Mensaje de error ..."
        Screen.MousePointer = 0
    End If
End Sub




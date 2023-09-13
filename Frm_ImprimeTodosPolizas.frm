VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_ImprimeTodosPolizas 
   Caption         =   "Seleccion Masiva de Polizas a Imprimir"
   ClientHeight    =   5115
   ClientLeft      =   3525
   ClientTop       =   3705
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5115
   ScaleWidth      =   8310
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   4680
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtHastav 
      Height          =   285
      Left            =   5640
      TabIndex        =   20
      Top             =   2040
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   4455
      Left            =   5520
      TabIndex        =   2
      Top             =   120
      Width           =   2655
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   2085
         Top             =   1320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdExportatxt 
         Caption         =   "Exporta .txt Boleta electronica"
         Height          =   660
         Left            =   1485
         TabIndex        =   22
         Top             =   3735
         Width           =   1110
      End
      Begin VB.TextBox txtDesdev 
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Text            =   "1"
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar"
         Height          =   615
         Left            =   1320
         Picture         =   "Frm_ImprimeTodosPolizas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtNroArch 
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   45
         Picture         =   "Frm_ImprimeTodosPolizas.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3720
         Width           =   720
      End
      Begin VB.CommandButton brnBuscarfec 
         Caption         =   "Exporta"
         Height          =   675
         Left            =   765
         TabIndex        =   10
         Top             =   3720
         Width           =   720
      End
      Begin VB.CommandButton cmd_Marcar 
         Height          =   495
         Left            =   1560
         Picture         =   "Frm_ImprimeTodosPolizas.frx":0AFC
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox txt_polizas 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   3240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Txt_LiqFecTer 
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   4
         Top             =   2880
         Width           =   1005
      End
      Begin VB.TextBox Txt_LiqFecIni 
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   3
         Top             =   2880
         Width           =   1005
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   2055
         Top             =   855
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label5 
         Caption         =   "-"
         Height          =   255
         Left            =   1200
         TabIndex        =   19
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "Marcar desde"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Rango de Impresion"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label lblConteo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   15
         Top             =   3240
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Nro de Archivo"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Ingreso por Polizas"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Lbl_Nombre 
         Alignment       =   2  'Center
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   71
         Left            =   1080
         TabIndex        =   6
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Ingrese Fecha a Procesar"
         Height          =   255
         Index           =   72
         Left            =   120
         TabIndex        =   5
         Top             =   2640
         Width           =   2475
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.ListBox Lst_LiqSeleccion 
         Height          =   3885
         ItemData        =   "Frm_ImprimeTodosPolizas.frx":19C6
         Left            =   120
         List            =   "Frm_ImprimeTodosPolizas.frx":19C8
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   240
         Width           =   5085
      End
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
Attribute VB_Name = "Frm_ImprimeTodosPolizas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Dim vlCodConceptos As String
Dim vlFecha As String
Dim vlFechaIni As String
Dim vlFechaTer As String
Dim vlCodPar As String
Dim vlNumEndoso As Integer
Dim vlNumOrden As Integer
Dim vlNombreComuna As String
Dim vlNombreSucursal As String
Dim vlPeriodo As String
Dim vlCodTipReceptor As String
Dim vlTipoIden As String
Dim vlNumIden As String
Dim vlIdenEmpresa As String
Dim vlCodTp As String
Dim vlCodTr As String
Dim vlCodAl As String
Dim vlCodPa As String
Dim vlFecNacTitular As String

Dim vlRutReceptor As Integer
Dim vlRutRec As Double
Dim vlDgvRec As String
Dim vlNomRec As String
Dim vlNumCargas As Integer
Dim vlNombreEstado As String
Dim vlCodHabDesCCAF As String
Dim vlNombreConcepto As String
Dim vlNombreCCAF As String
Dim vlTipoRet As String
Dim vlModRet As String
Dim vlRutAux As Double
Dim vlMtoCarga As Double
Dim vlEstCerEst As String * 1
Dim vlFechaActual As String
Dim vlCodCauSus As String
Dim vlCodEstado As String
Dim vlArchivo As String
Dim vlCodInsSalud As String
Dim vlFechaEfectoSalud As String
Dim vlRut As String
Dim vlOrden As Integer
Dim vlRutCia As String
Dim vlTipoIdenCia As String
Dim vlNumIdenCia As String
Dim vlNombreUsuario As String
Dim vlCiuCia As String
Dim vlFechaEfectoAsigFam As String
Dim vlPago As String
Dim vlFechaDesde As String
Dim vlFechaHasta As String
Dim vlRutCliente As String
Dim vltipoidenCompania As String
Dim vlcobertura As String

Dim vlOpcionPago As String

Dim vlNumPerPago As String
Dim vlNumCargasPagadas As Integer
Dim vlNombreBenef As String
Dim vlRutBenef As String
Dim vlTipoRenta As String
Dim vlTipoPension As String
Dim vlTipoModalidad As String
Dim vlFechaVigenciaRta As String
Dim vlMtoPensionBruta As Double

Dim vlNumEndosoNoBen As Integer

Dim vlRegistro2 As ADODB.Recordset
Dim vlRegistro3 As ADODB.Recordset

Dim vlNombreCompania As String
Dim vlDirCompania As String
Dim vlFonoCompania As String

Const clFechaTopeTer As String * 8 = "99991231"
Const clRecTutor As String * 1 = "T"
Const clRecPensionado As String * 1 = "P"
Const clModOrigenCCAF As String * 4 = "CCAF"
Const clCargaActiva As String * 1 = "A"
Const clCodSinDerPen As String * 2 = "10"
Const clDeptoUsuario As String = "Servicio al Cliente"
Const clCodEstPen99 As String * 2 = 99
Const clOpcionDEF As String * 3 = "DEF"

Const clCodTipReceptorR As String * 1 = "R"

'CMV-20061102 I
Dim vlNumOrd As Integer
Dim vlUltimoPerPago As String
Dim vlPrcCastigoQui As Double
Dim vlTopeMaxQui As Double

Const clCodEstadoC As String * 1 = "C"
'CMV-20061102 F

Dim vlLargoTipoIden    As Integer 'sirve para llenar la grilla
Dim vlPosicionTipoIden As Integer 'sirve para llenar la grilla

Dim vlCodTipoIdenBenCau As String
Dim vlNumIdenBenCau As String

Dim vlCodTipoIdenBenTut As String
Dim vlNumIdenBenTut As String

Dim vlNombreSeg As String, vlApMaterno As String

Dim vlNombreApoderado As String
Dim vlCargoApoderado As String

Dim objCrystal As CRAXDRT.Application
Dim objRpt As CRAXDRT.Report
Private rs As ADODB.Recordset
Dim RSListaDoc As New ADODB.Recordset
Dim boltodos As Boolean


Function cargaReporteFin()

    Dim objRep As New ClsReporte
    Dim LNGa As Long
    Dim vlSql As String
    Dim rs As New ADODB.Recordset
    
    'vlSql = " select d.num_poliza as num_poliza, d.GLS_NOMBRE as nombres, L.gls_direccion as direccion, L.GLS_DIRECCION2 as direccion2, gls_fono "
    'vlSql = vlSql & " from PP_TMAE_CONTABLEREGPAGO C"
    'vlSql = vlSql & " INNER JOIN PP_TMAE_CONTABLEDETREGPAGO D ON"
    'vlSql = vlSql & " C.NUM_ARCHIVO=D.NUM_ARCHIVO and C.COD_TIPREG=D.COD_TIPREG AND"
    'vlSql = vlSql & " c.COD_TIPMOV = d.COD_TIPMOV And c.Cod_Moneda = d.Cod_Moneda"
    'vlSql = vlSql & " join ("
    'vlSql = vlSql & "       select distinct num_poliza, gls_direccion, GLS_DIRECCION2 from PP_TTMP_LIQUIDACION WHERE COD_USUARIO='" & vgUsuario & "'"
    ' vlSql = vlSql & "      ) L on L.num_poliza=D.num_poliza"
   ' vlSql = vlSql & " join pd_tmae_poliza p on p.num_poliza=D.num_poliza and num_endoso=(select max(num_endoso) from pp_tmae_poliza where num_poliza=p.num_poliza)"
   ' vlSql = vlSql & " where c.NUM_ARCHIVO='" & Me.txtNroArch.Text & "' and d.cod_tipmov=38"
   ' vlSql = vlSql & " order by 1"
     
    vlSql = " select a.num_poliza as num_poliza, a.GLS_NOMRECEPTOR as nombres, gls_dirben as direccion,  gls_direccion2 as direccion2, c.gls_fonoben as gls_fono from pp_ttmp_liquidacion a"
    vlSql = vlSql & " join pp_tmae_poliza b on a.num_poliza=b.num_poliza and a.num_endoso=b.num_endoso"
    vlSql = vlSql & " join pp_tmae_ben c on a.num_poliza=c.num_poliza and a.num_endoso=c.num_endoso and a.num_orden=c.num_orden"
    vlSql = vlSql & " where cod_usuario='" & vgUsuario & "' and a.num_poliza in (select num_poliza from pp_tmae_liqpagopendef where num_perpago=a.num_perpago and cod_tipopago<>'P') order by 1"
    'vlSql = vlSql & " where cod_usuario='" & vgUsuario & "' and a.num_poliza in (select num_poliza from pp_tmae_liqpagopendef where num_perpago=a.num_perpago and COD_TIPPENSION IN ('08','09','10','11','12')) order by 1"
    
   
    
    Set rs = vgConexionBD.Execute(vlSql)
     
    LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PP_Rpt_ListadoImpresos.rpt"), ".RPT", ".TTX"), 1)
    
        
    If objRep.CargaReporte(strRpt & "", "PP_Rpt_ListadoImpresos.rpt", "Listado de Impresion de Boletas", rs, True) = False Then
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Function
    End If


End Function
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
Function flInformeLiqPago()
'Imprime Liquidaciones de Pension
On Error GoTo Err_VenTut
    Dim vlSQLQuery As String
    Dim vgRutCliente As String
    Dim vgDgvCliente As String
    Dim vlTipoId As String
    Dim cadenaPolizas As String
    ''''''''''''''''''''''''
    Dim rs As ADODB.Recordset
    Dim rsBen As ADODB.Recordset
    Dim rsLiq As ADODB.Recordset
    Dim objRep As New ClsReporte
    Dim cadena As String
    Dim nombres, Direccion, fec_finter As String
    Dim LNGa As Long
    
    Call flRutCliente
    
    vgRutCliente = ""
    vgDgvCliente = ""
    
    vlFechaDesde = Format(CDate(Trim(Txt_LiqFecIni.Text)), "yyyymmdd")
    vlFechaHasta = Format(CDate(Trim(Txt_LiqFecTer.Text)), "yyyymmdd")
    
    'I--- ABV 12/03/2005 ---
    'vlPago = Trim(fgBuscaTipoPago(Trim(Txt_PenPoliza.Text)))
    vlPago = "T"
    'F--- ABV 12/03/2005 ---
    
    'vlTipoId = (Trim(Mid(Lbl_GrupTipoIdent.Caption, 1, (InStr(1, Lbl_GrupTipoIdent.Caption, "-") - 1))))
    
    
    
    cadenaPolizas = flObtenerPolizas()
    
    If cadenaPolizas = "(' ')" Then
        MsgBox ("Debe elegir polizas para la impresiòn.")
        Exit Function
    End If
    
    If Not flLlenaTemporal_masivo(vlFechaDesde, vlFechaHasta, cadenaPolizas, clOpcionDEF, vlPago) Then
        Exit Function
    End If
   

    'esta función se utiliza cuando Peter quiere sacar masivamente las boletas de pago de un grupo de polizas determinado
    'cuando se utiliza esta función, la anterior se pone en comentario o se salta con el debug
    'RVF 03/02/2011
    'If Not flLlenaTemporal_masivo(vlFechaDesde, vlFechaHasta, Trim(Txt_PenPoliza.Text), (vlTipoId), (Lbl_GrupNumIdent.Caption), clOpcionDEF, vlPago) Then
    '    Exit Function
    'End If
    

    
'    Dim objRep As New ClsReporte
'    vlFechaTermino = Format(CDate(Trim(Txt_Hasta.Text)), "yyyymmdd")
'    vgPalabra = ""
'    vgPalabra = "Certificados vencidos al " & Txt_Hasta.Text
   
   
'    Set rs = New ADODB.Recordset
'    rs.CursorLocation = adUseClient
'    rs.Open "PK_LISTA_POLIZAS_BOLETAS.LISTAR", vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
'    Dim LNGa As Long
'    LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PP_Rpt_LiquidacionRV.rpt"), ".RPT", ".TTX"), 1)
    
        
'    If objRep.CargaReporte(vlArchivo, "Informe de Liquidación de Rentas Vitalicias", rs, True, _
'                            ArrFormulas("NombreCompania", vgNombreCompania), _
'                            ArrFormulas("rutcliente", vlRutCliente)) = False Then
'
'        MsgBox "No se pudo abrir el reporte", vbInformation
'        Exit Sub
'    End If

''comenta mientras
    Screen.MousePointer = 11
    vlArchivo = strRpt & "PP_Rpt_LiquidacionRV.rpt"

    If Not fgExiste(vlArchivo) Then
        MsgBox "Archivo de Reporte no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
        Screen.MousePointer = 0
        Exit Function
    End If


    cadena = "select a.*, gls_comuna, gls_provincia, gls_region, g.gls_tipoiden, g.gls_tipoidencor,"
    cadena = cadena & " c.num_idenben, c.gls_nomben as gls_nomben1, c.gls_nomsegben, c.gls_patben, c.gls_matben, gls_dirben, h.gls_tipoiden as tipoRec, i.gls_tipoiden as tipoTit"
    cadena = cadena & " from pp_ttmp_liquidacion a"
    cadena = cadena & " join pp_tmae_poliza b on a.num_poliza=b.num_poliza and a.num_endoso=b.num_endoso"
    cadena = cadena & " join pp_tmae_ben c on a.num_poliza=c.num_poliza and a.num_endoso=c.num_endoso and a.num_orden=c.num_orden"
    cadena = cadena & " join ma_tpar_comuna d on c.cod_direccion=d.cod_direccion"
    cadena = cadena & " join ma_tpar_provincia e on d.cod_provincia=e.cod_provincia"
    cadena = cadena & " join ma_tpar_region f on e.cod_region=f.cod_region"
    cadena = cadena & " join ma_tpar_tipoiden g on c.cod_tipoidenben=g.cod_tipoiden"
    'RRR 22/07/2015
    cadena = cadena & " join ma_tpar_tipoiden h on a.cod_tipoidenreceptor=h.cod_tipoiden"
    cadena = cadena & " join ma_tpar_tipoiden i on a.cod_tipoidentit=i.cod_tipoiden"
    'RRR 22/07/2015
    cadena = cadena & " where cod_usuario='" & vgUsuario & "'"
    cadena = cadena & " and b.num_endoso in(select max(num_endoso) from pp_tmae_poliza where num_poliza=b.num_poliza)"
    cadena = cadena & " and num_poliza in (select num_poliza from pp_tmae_liqpagopendef where num_perpago=a.num_perpago and cod_tipopago<>'P')"
    'mvg 20170904
    cadena = cadena & " and c.ind_bolelec='N'"
    cadena = cadena & " order by 3"
    
    
    Set rsLiq = New ADODB.Recordset
    rsLiq.CursorLocation = adUseClient
    rsLiq.Open cadena, vgConexionBD, adOpenForwardOnly, adLockReadOnly

    vlRutCliente = flRutCliente 'vgRutCliente + " - " + vgDgvCliente

    LNGa = CreateFieldDefFile(rsLiq, Replace(UCase(strRpt & "Estructura\PP_Rpt_LiquidacionRV.rpt"), ".RPT", ".TTX"), 1)

    If objRep.CargaReporte(strRpt & "", "PP_Rpt_LiquidacionRV.rpt", "Informe de Liquidación de Rentas Vitalicias", rsLiq, True, _
                            ArrFormulas("NombreCompania", UCase("Protecta SA compañia de Seguros")), _
                            ArrFormulas("rutcliente", vlRutCliente)) = False Then

        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Function
    End If


 
    'cargos
    
    'cadena = "select distinct L.num_poliza, num_perpago, L.GLS_DIRECCION, L.GLS_NOMRECEPTOR, L.GLS_NOMSEGRECEPTOR, L.GLS_PATRECEPTOR, L.GLS_MATRECEPTOR"
    'cadena = cadena & " , c.gls_comuna || ' - ' || p.gls_provincia || ' - ' || r.gls_region as Ubicacion, L.cod_tipreceptor"
    'cadena = cadena & " from pp_tmae_liqpagopendef L"
    'cadena = cadena & " join ma_tpar_comuna C on l.cod_direccion=C.cod_direccion"
    'cadena = cadena & " join ma_tpar_provincia P on c.cod_provincia=P.cod_provincia"
    'cadena = cadena & " join ma_tpar_region R on C.cod_region=R.COD_REGION"
    'cadena = cadena & " where L.FEC_PAGO >= '" & vlFechaDesde & "' AND L.FEC_PAGO <= '" & vlFechaHasta & "'"
    'cadena = cadena & " AND L.cod_tipopago='R'"
    '''Solo para algunas cartas, comentar cuando no sea necesario
    'If Not boltodos = True Then
    '    cadena = cadena & " AND L.NUM_POLIZA IN " & cadenaPolizas & ""
    'End If
    ''''''
    ''cadena = cadena & " AND COD_TIPRECEPTOR='P'"
    'cadena = cadena & " order by 1"

    'vlFechaDesde = Format(CDate(Trim(Txt_LiqFecIni.Text)), "yyyymmdd")
    'vlFechaHasta = Format(CDate(Trim(Txt_LiqFecTer.Text)), "yyyymmdd")

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "PD_LISTA_POLIZA.RPTCARTASCONSTAN('" & vlFechaDesde & "','" & vlFechaHasta & "')", vgConexionBD, adOpenForwardOnly, adLockReadOnly
    'rs.Open cadena, vgConexionBD, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then
        'MsgBox "No hay datos para mostrar", vbExclamation, "Pagos Pendientes Acumulados"
        
        nombres = rs!Gls_NomReceptor & " " & rs!Gls_NomSegReceptor & " " & rs!Gls_PatReceptor & " " & rs!Gls_MatReceptor
        Direccion = rs!Gls_Direccion
    Else
        Exit Function
    End If

    'Dim LNGa As Long
    LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PP_Rpt_LiquidacionCargo.rpt"), ".RPT", ".TTX"), 1)

    If objRep.CargaReporte(strRpt & "", "PP_Rpt_LiquidacionCargo.rpt", "Cargo de Boleta", rs, True, _
                            ArrFormulas("NombreCompania", "Protecta SA compañia de Seguros"), _
                            ArrFormulas("NombreTitular", nombres), _
                            ArrFormulas("Direccion", Direccion), _
                            ArrFormulas("NombresHijos", Direccion)) = False Then

        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Function
    End If
     
   Call cargaReporteFin
    
    
    
Exit Function
Err_VenTut:
    Screen.MousePointer = 0
    MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
End Function



Private Sub Crea_Estructura()

Set RSListaDoc = New ADODB.Recordset
    With RSListaDoc.Fields
        .Append "ITEM", adVarChar, 3
        .Append "POLIZA", adVarChar, 50
        .Append "NOMBRES", adVarChar, 50
        .Append "DIRECCION", adVarChar, 50
    End With
RSListaDoc.Open
End Sub
Private Function CreaTablaRes(rs As ADODB.Recordset) As ADODB.Recordset

    Dim i As Integer
    
    Call Crea_Estructura
    
    Do While Not rs.EOF
    
        RSListaDoc.Fields("ITEM") = CStr(i)
        RSListaDoc.Fields("POLIZA") = rs!num_poliza
        RSListaDoc.Fields("NOMBRES") = rs!Gls_NomBen & " " & rs!Gls_NomSegBen & " " & rs!Gls_PatBen & " " & rs!Gls_MatBen
        RSListaDoc.Fields("DIRECCION") = rs!Gls_Direccion
       rs.MoveNext
       i = i + 1
    Loop

    Set CreaTablaRes = RSListaDoc

End Function

Private Function CopyRecordset(rsSource As ADODB.Recordset) As ADODB.Recordset

Dim rs As ADODB.Recordset
Dim pb As New PropertyBag

    ' creamos la copia del recordset
    pb.WriteProperty "rs", rsSource
    Set rs = pb.ReadProperty("rs")
    
    Set pb = Nothing
    
    'regresamos el recordset copiado
    Set CopyRecordset = rs

    
End Function

Private Function flLlenaTemporal_masivo(sFecDesde As String, sFecHasta As String, scadPoliza As String, iGlosaOpcion As String, sPago As String) As Boolean

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
    Dim vlFecPago As String
    Dim vlDescViaPago As String
    Dim vlSucursal As String
    Dim vlDescSucursal As String
    
    flLlenaTemporal_masivo = False
    vlItem = 1
    vlNumConceptosHab = 0
    vlNumConceptosDesc = 0
    vlPoliza = ""
    vlOrden = 0
    vlPerPago = ""
    vlFecPago = ""
    vlTipPension = ""
    vlViaPago = ""
    vlCajaComp = ""
    vlInsSalud = ""
    vlCodDireccion = 0
    vlAfp = ""
    vlSucursal = ""
    vlDescSucursal = ""
    'VARIABLES GENERALES
    stTTMPLiquidacion.Cod_Usuario = vgUsuario
    
    
    ProgressBar1.Value = 0
    
    
    'Elimina Datos de la Tabla Temporal
    vlSql = "DELETE FROM PP_TTMP_LIQUIDACION WHERE COD_USUARIO = '" & vgUsuario & "'"
    vgConexionBD.Execute (vlSql)
    
    vlSql = "SELECT P.NUM_POLIZA, L.NUM_ENDOSO, P.NUM_ORDEN, P.COD_CONHABDES, P.MTO_CONHABDES, C.COD_TIPMOV, P.NUM_PERPAGO, P.COD_TIPOIDENRECEPTOR,P.NUM_IDENRECEPTOR, P.COD_TIPRECEPTOR,"
    vlSql = vlSql & " L.GLS_DIRECCION, L.FEC_PAGO, L.GLS_NOMRECEPTOR, L.GLS_NOMSEGRECEPTOR, L.GLS_PATRECEPTOR,"
    vlSql = vlSql & " L.GLS_MATRECEPTOR, L.MTO_LIQPAGAR, L.COD_DIRECCION, "
    vlSql = vlSql & " L.COD_TIPPENSION, L.COD_VIAPAGO, L.COD_SUCURSAL, L.COD_INSSALUD," ''*L.COD_CAJACOMPEN,
    vlSql = vlSql & " L.MTO_PENSION, L.NUM_CARGAS, L.MTO_HABER, L.MTO_DESCUENTO," ''*L.DGV_RECEPTOR,
    vlSql = vlSql & " B.NUM_IDENBEN, B.COD_TIPOIDENBEN, B.GLS_NOMBEN, B.GLS_NOMSEGBEN, B.GLS_PATBEN, B.GLS_MATBEN, "
    vlSql = vlSql & " C.GLS_CONHABDES, M.COD_SCOMP, POL.COD_AFP, L.COD_MONEDA, M.GLS_ELEMENTO AS MONEDA, COD_VEJEZ, TV.GLS_ELEMENTO AS GLS_TP2, TP.GLS_ELEMENTO AS GLS_TP "
    'RRR 22/07/2015
    vlSql = vlSql & " , (SELECT  GLS_NOMBEN || ' ' || GLS_NOMSEGBEN || ' ' || GLS_PATBEN || ' ' || GLS_MATBEN FROM PP_TMAE_BEN WHERE NUM_POLIZA=P.NUM_POLIZA AND NUM_ORDEN=1 AND NUM_ENDOSO=1) AS NOMTIT"
    vlSql = vlSql & " , (SELECT NUM_IDENBEN FROM PP_TMAE_BEN WHERE NUM_POLIZA=P.NUM_POLIZA AND NUM_ORDEN=1 AND NUM_ENDOSO=1) AS NUMIDENTIT"
    vlSql = vlSql & " , (SELECT COD_TIPOIDENBEN FROM PP_TMAE_BEN WHERE NUM_POLIZA=P.NUM_POLIZA AND NUM_ORDEN=1 AND NUM_ENDOSO=1) AS TIPIDENTIT"
    'RRR 22/07/2015
    vlSql = vlSql & " FROM PP_TMAE_PAGOPEN" & iGlosaOpcion & " P, PP_TMAE_LIQPAGOPEN" & iGlosaOpcion & " L, MA_TPAR_CONHABDES C"
    vlSql = vlSql & ", PP_TMAE_POLIZA POL, PP_TMAE_BEN B, MA_TPAR_TABCOD M, PD_TMAE_POLIZA PD, MA_TPAR_TABCOD TV, MA_TPAR_TABCOD TP WHERE"
    vlSql = vlSql & " L.NUM_POLIZA = B.NUM_POLIZA AND"
    vlSql = vlSql & " L.NUM_ENDOSO = B.NUM_ENDOSO AND"
    vlSql = vlSql & " L.NUM_ORDEN = B.NUM_ORDEN AND"
    vlSql = vlSql & " L.NUM_POLIZA = POL.NUM_POLIZA AND"
    vlSql = vlSql & " L.NUM_ENDOSO = POL.NUM_ENDOSO AND"
    vlSql = vlSql & " L.COD_MONEDA = M.COD_ELEMENTO AND"
    vlSql = vlSql & " M.COD_TABLA = 'TM' AND" 'Tabla de Monedas
    vlSql = vlSql & " PD.NUM_POLIZA = POL.NUM_POLIZA AND"
    'If Chk_Pensionado.Value = 1 Then
        'If Txt_Poliza <> "" Then
       
        
        'vlSql = vlSql & " L.NUM_POLIZA IN (5,11,47,97,132,142,211,240,287,415,473,486,496,510,523,541,550,556,595,602,603,625,630,631,634,640,641,649,650,713,720,737,739,740,747,750,754,862,887,889,911,918,926,937,941,948,956,962,965,981,990,1017,1025,1046,1063,1076,1077,1105,1106,1116,1134,1159,1186,1238,1239,1251,1281,1304,1356,1430,1481,1516,1630,1722,1734,1798,1874,1881,2087,2234,2243,2250,2251,2261) AND"
        If boltodos = True Then
           'RRR
            vlSql = vlSql & " L.NUM_POLIZA BETWEEN (SELECT min(num_poliza) FROM PP_TMAE_POLIZA WHERE NUM_ENDOSO=1 ) and (SELECT max(num_poliza) FROM PP_TMAE_POLIZA WHERE NUM_ENDOSO=1 ) AND"
        Else
            '************ numeros de polizas a imprimir ********************************
            vlSql = vlSql & " L.NUM_POLIZA IN " & scadPoliza & " AND "
            '***************************************************************************
        End If
        
        
        
        
        'If Txt_Rut <> "" Then
        'If iTipoIden <> "" Then
        '    vlSql = vlSql & " B.COD_TIPOIDENBEN = " & Trim(iTipoIden) & " AND"
        'End If
'        If iNumIden <> "" Then
'            vlSql = vlSql & " B.NUM_IDENBEN = '" & Trim(iNumIden) & "' AND"
'        End If
'    End If
    vlSql = vlSql & " L.NUM_POLIZA = P.NUM_POLIZA"
    vlSql = vlSql & " AND L.NUM_ORDEN = P.NUM_ORDEN"
    ''*vlSql = vlSql & " AND L.RUT_RECEPTOR = P.RUT_RECEPTOR"
    vlSql = vlSql & " AND L.COD_TIPOIDENRECEPTOR=P.COD_TIPOIDENRECEPTOR"
    vlSql = vlSql & " AND L.NUM_IDENRECEPTOR=P.NUM_IDENRECEPTOR"
    vlSql = vlSql & " AND L.COD_TIPRECEPTOR = P.COD_TIPRECEPTOR"
    vlSql = vlSql & " AND L.NUM_PERPAGO = P.NUM_PERPAGO"
    'iPago = "R"
    If iPago = "P" Then 'PRIMER PAGO
        vlSql = vlSql & " AND L.COD_TIPOPAGO = 'P'"
    ElseIf iPago = "R" Then 'PAGO EN REGIMEN
        vlSql = vlSql & " AND L.COD_TIPOPAGO = 'R'"
    End If
    'I--- ABV 12/03/2005 ---
    'If (iPago = "T") Then
        vlSql = vlSql & " AND L.COD_TIPOPAGO in ('R','P')"
    'End If
    'F--- ABV 12/03/2005 ---
 
    vlSql = vlSql & " AND L.FEC_PAGO >= '" & sFecDesde & "' AND L.FEC_PAGO <= '" & sFecHasta & "'"
    'vlSql = vlSql & " AND L.NUM_PERPAGO in ('201107','201108','201109','201110','201112')" 'between '" & sFecDesde & "' and '" & sFecHasta & "'"
    vlSql = vlSql & " AND P.COD_CONHABDES  = C.COD_CONHABDES"
    'RRR 27/06/2013
    vlSql = vlSql & " AND PD.COD_VEJEZ=TV.COD_ELEMENTO"
    vlSql = vlSql & " AND TV.COD_TABLA = 'TV'"
    vlSql = vlSql & " AND PD.COD_TIPPENSION=TP.COD_ELEMENTO"
    vlSql = vlSql & " AND TP.COD_TABLA = 'TP'"
    vlSql = vlSql & " AND PD.NUM_ENDOSO=1"
    vlSql = vlSql & " AND P.COD_CONHABDES NOT IN (60)"
    '''''''''''''''
    'vlSql = vlSql & " ORDER BY P.NUM_POLIZA, P.NUM_ORDEN, P.RUT_RECEPTOR, P.COD_TIPRECEPTOR, P.NUM_PERPAGO"
    vlSql = vlSql & " ORDER BY P.NUM_POLIZA, P.NUM_PERPAGO, num_orden  "
    'vlSql = vlSql & " ORDER BY P.NUM_PERPAGO, P.NUM_POLIZA, P.NUM_ORDEN,P.NUM_IDENRECEPTOR, P.COD_TIPOIDENRECEPTOR, P.COD_TIPRECEPTOR," 'hqr 17/03/2005 Se agrega número de periodo
    'vlSql = vlSql & " C.COD_IMPONIBLE DESC, C.COD_TRIBUTABLE DESC, C.COD_TIPMOV DESC"
    Set vlTB = vgConexionBD.Execute(vlSql)
    
    Set RSListaDoc = vlTB
   
    'MsgBox ("cuenta " & CStr(RSListaDoc.RecordCount))
    
    If Not vlTB.EOF Then
    
    ProgressBar1.Min = 0
    ProgressBar1.Max = 1600
    
        Do While Not vlTB.EOF
'            If vlNumIdenReceptor = "22088246" Then
'                a = 1
'            End If
                        
            If vlPoliza <> vlTB!num_poliza Or vlOrden <> vlTB!Num_Orden Or vlNumIdenReceptor <> vlTB!Num_IdenReceptor Or vlCodTipoIdenReceptor <> vlTB!Cod_TipoIdenReceptor Or vlTipReceptor <> vlTB!Cod_TipReceptor Or (vlPerPago <> vlTB!Num_PerPago And vlPago = "R") Or (vlFecPago <> vlTB!Fec_Pago And vlPago = "P") Then 'hqr 17/03/2006 Se agrega número de periodo
                'Reinicia el Contador
                
                Dim a As Integer
                
'                If vlPoliza = "0000000867" Then
'                    a = 1
'                End If
                
                vlItem = 1
                vlPoliza = vlTB!num_poliza
                vlOrden = vlTB!Num_Orden
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
                    stTTMPLiquidacion.Gls_Direccion = " "
                End If
                stTTMPLiquidacion.Gls_NomReceptor = vlTB!Gls_NomReceptor & " " & IIf(IsNull(vlTB!Gls_NomSegReceptor), "", vlTB!Gls_NomSegReceptor & " ") & vlTB!Gls_PatReceptor & IIf(IsNull(vlTB!Gls_MatReceptor), "", " " + vlTB!Gls_MatReceptor)
                If stTTMPLiquidacion.Cod_TipReceptor <> "R" Then
                    stTTMPLiquidacion.Cod_Direccion = vlTB!Cod_Direccion
                Else
                    stTTMPLiquidacion.Cod_Direccion = "0"
                End If
                stTTMPLiquidacion.Cod_TipoIdenBen = vlTB!Cod_TipoIdenBen
                stTTMPLiquidacion.Num_IdenBen = vlTB!Num_IdenBen
                stTTMPLiquidacion.Gls_NomBen = vlTB!Gls_NomBen & " " & IIf(IsNull(vlTB!Gls_NomSegBen), "", vlTB!Gls_NomSegBen & " ") & vlTB!Gls_PatBen & IIf(IsNull(vlTB!Gls_MatBen), "", " " & vlTB!Gls_MatBen)
                'para primeros pagos
                stTTMPLiquidacion.Mto_LiqPagar = 0
                stTTMPLiquidacion.Num_Cargas = 0 'vlTB!Num_Cargas
                stTTMPLiquidacion.Mto_LiqHaber = 0
                stTTMPLiquidacion.Mto_LiqDescuento = 0
                stTTMPLiquidacion.gls_vejez = vlTB!COD_VEJEZ 'RRR 02/11/2012
                stTTMPLiquidacion.Cod_TipoIdenTit = vlTB!TIPIDENTIT 'RRR 22/17/2015
                stTTMPLiquidacion.Num_IdenTit = vlTB!NUMIDENTIT 'RRR 22/17/2015
                stTTMPLiquidacion.gls_nomTit = vlTB!NOMTIT 'RRR 22/17/2015
             
                'fin primeros pagos
                'Obtiene Fecha de Término del Poder Notarial
                If stTTMPLiquidacion.Cod_TipReceptor <> "R" Then
                    vlSql = "SELECT tut.fec_terpodnot FROM pp_tmae_tutor tut"
                    vlSql = vlSql & " WHERE tut.num_poliza = '" & stTTMPLiquidacion.num_poliza & "'"
                    vlSql = vlSql & " AND tut.num_orden = " & stTTMPLiquidacion.Num_Orden
                    vlSql = vlSql & " AND tut.cod_tipoidentut = " & vlTB!Cod_TipoIdenReceptor & " "
                    vlSql = vlSql & " AND tut.num_identut = '" & vlTB!Num_IdenReceptor & "' "
                    
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

            stTTMPLiquidacion.Num_PerPago = vlTB!Num_PerPago
            'stTTMPLiquidacion.Mto_LiqPagar = vlTB!Mto_LiqPagar
            stTTMPLiquidacion.Mto_Pension = vlTB!Mto_Pension
            stTTMPLiquidacion.Num_Cargas = 0 'vlTB!Num_Cargas
            'stTTMPLiquidacion.Mto_LiqHaber = vlTB!Mto_Haber
            'stTTMPLiquidacion.Mto_LiqDescuento = vlTB!Mto_Descuento
            If vlPerPago <> stTTMPLiquidacion.Num_PerPago Then
                stTTMPLiquidacion.Fec_Pago = vlTB!Fec_Pago
                'Obtiene Fecha del Próximo Pago
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
                'Obtiene Valor UF
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
                    stTTMPLiquidacion.Gls_TipPension = vlTB2!GLS_ELEMENTO & " " & IIf(vlTB!COD_VEJEZ = "S", "", vlTB!GLS_TP2)
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
                        
'            'obtiene vejes rrr 02/11/2012
            stTTMPLiquidacion.gls_vejez = "" 'vlTB!GLS_TP & " " & vlTB!GLS_TP2
'            If Not IsNull(stTTMPLiquidacion.gls_vejez) Then
'                vlSql = "SELECT tab.gls_elemento FROM ma_tpar_tabcod tab"
'                vlSql = vlSql & " WHERE tab.cod_tabla = 'TV'"
'                vlSql = vlSql & " AND tab.cod_elemento = '" & vlTB!Cod_vejez & "'"
'                Set vlTB2 = vgConexionBD.Execute(vlSql)
'                If Not vlTB2.EOF Then
'                    stTTMPLiquidacion.gls_vejez = vlTB!GLS_TP & " " & vlTB!GLS_TP2
'                Else
'                    stTTMPLiquidacion.gls_vejez = ""
'                End If
'            End If
            
                        
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
                'stTTMPLiquidacion.Gls_MontoPension = fgConvierteNumeroLetras(stTTMPLiquidacion.Mto_LiqPagar, vlTB!Moneda)
                stTTMPLiquidacion.Gls_MontoPension = fgConvierteNumeroLetras((vlTB!Mto_Haber - vlTB!Mto_Descuento), vlTB!Moneda)
            End If
            If vlTB!cod_tipmov = "H" Then 'haber
                If vlPago <> "R" Then
                    stTTMPLiquidacion.Mto_LiqHaber = stTTMPLiquidacion.Mto_LiqHaber + vlTB!Mto_ConHabDes
                    stTTMPLiquidacion.Mto_LiqPagar = stTTMPLiquidacion.Mto_LiqPagar + vlTB!Mto_ConHabDes
                    'stTTMPLiquidacion.Gls_MontoPension = fgConvierteNumeroLetras(stTTMPLiquidacion.Mto_LiqPagar, vlTB!Moneda)
                    stTTMPLiquidacion.Gls_MontoPension = fgConvierteNumeroLetras((vlTB!Mto_Haber - vlTB!Mto_Descuento), vlTB!Moneda)
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
                    stTTMPLiquidacion.Gls_MontoPension = fgConvierteNumeroLetras(vlTB!Mto_LiqPagar, vlTB!Moneda)
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
    flLlenaTemporal_masivo = True

  
End Function

Private Sub brnBuscarfec_Click()
Dim vlNumPoliza As String
Dim Conteo As Long
Dim Periodo As String
Conteo = 0

vlFechaHasta = Format(CDate(Trim(Txt_LiqFecTer.Text)), "yyyymmdd")
    
Periodo = Mid(vlFechaHasta, 1, 6)

ProgressBar1.Min = 0
ProgressBar1.Max = Lst_LiqSeleccion.ListCount

'If Dir(App.Path & "\PDF\", Periodo) = "" Then
'    MkDir (App.Path & "\PDF\" & Periodo)
'End If


sName = Dir(App.Path & "\PDF\" & Periodo, vbDirectory)

If sName = "" Then
    MkDir (App.Path & "\PDF\" & Periodo)
Else
    
End If


For vgI = 0 To Lst_LiqSeleccion.ListCount - 1
        If Lst_LiqSeleccion.Selected(vgI) Then
         
           If Trim(Lst_LiqSeleccion.List(vgI)) <> "TODOS" Then
                vlNumPoliza = (Trim(Mid(Lst_LiqSeleccion.List(vgI), 1, (InStr(1, Lst_LiqSeleccion.List(vgI), "-") - 1))))
                flInformeExportaPdf (vlNumPoliza)
                
                
                
           End If
        End If
        Conteo = Conteo + 1
        'ProgressBar1.Value = Conteo
Next
If Conteo = Lst_LiqSeleccion.ListCount - 1 Then
        MsgBox ("Debe Seleccionar alguna poliza del Listado.")
        Exit Sub
End If
MsgBox "Boletas Exportadas con Exito.", vbInformation
End Sub
Function flInformeExportaPdf(numPol As String)
'Imprime Liquidaciones de Pension
On Error GoTo Err_VenTut
    Dim vlSQLQuery As String
    Dim vgRutCliente As String
    Dim vgDgvCliente As String
    Dim vlTipoId As String
    ''''''''''''''''''''''''
    Dim rs As ADODB.Recordset
    Dim rsLiq As ADODB.Recordset
    Dim objRep As New ClsReporte
    Dim cadena As String
    Dim nombres, Direccion, fec_finter As String
    Dim LNGa As Long
    Dim Periodo As String
    Dim desBoleta As String
    
    vgRutCliente = ""
    vgDgvCliente = ""
    
    vlFechaDesde = Format(CDate(Trim(Txt_LiqFecIni.Text)), "yyyymmdd")
    vlFechaHasta = Format(CDate(Trim(Txt_LiqFecTer.Text)), "yyyymmdd")
    
    Periodo = Mid(vlFechaHasta, 1, 6)
    
    'I--- ABV 12/03/2005 ---
    'vlPago = Trim(fgBuscaTipoPago(Trim(Txt_PenPoliza.Text)))
    vlPago = "R"
    'F--- ABV 12/03/2005 ---
    'vlTipoId = (Trim(Mid(Lbl_GrupTipoIdent.Caption, 1, (InStr(1, Lbl_GrupTipoIdent.Caption, "-") - 1))))
    
    'cadenaPolizas = numPol
    
    If Not flLlenaTemporal_masivo(vlFechaDesde, vlFechaHasta, numPol, clOpcionDEF, vlPago) Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
    
    vlArchivo = strRpt & "PP_Rpt_LiquidacionRV.rpt"
    If Not fgExiste(vlArchivo) Then
        MsgBox "Archivo de Reporte no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
        Screen.MousePointer = 0
        Exit Function
    End If
    
    cadena = "select a.*, gls_comuna, gls_provincia, gls_region, g.gls_tipoiden, g.gls_tipoidencor,"
    cadena = cadena & " c.num_idenben, c.gls_nomben as gls_nomben1, c.gls_nomsegben, c.gls_patben, c.gls_matben, gls_dirben"
    cadena = cadena & " from pp_ttmp_liquidacion a"
    cadena = cadena & " join pp_tmae_poliza b on a.num_poliza=b.num_poliza and a.num_endoso=b.num_endoso"
    cadena = cadena & " join pp_tmae_ben c on a.num_poliza=c.num_poliza and a.num_endoso=c.num_endoso and a.num_orden=c.num_orden"
    cadena = cadena & " join ma_tpar_comuna d on c.cod_direccion=d.cod_direccion"
    cadena = cadena & " join ma_tpar_provincia e on d.cod_provincia=e.cod_provincia"
    cadena = cadena & " join ma_tpar_region f on e.cod_region=f.cod_region"
    cadena = cadena & " join ma_tpar_tipoiden g on c.cod_tipoidenben=g.cod_tipoiden"
    cadena = cadena & " where cod_usuario='" & vgUsuario & "'"
    Set rsLiq = New ADODB.Recordset
    rsLiq.CursorLocation = adUseClient
    rsLiq.Open cadena, vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
    vlRutCliente = flRutCliente 'vgRutCliente + " - " + vgDgvCliente
    vlFechaHasta = Format(CDate(Trim(Txt_LiqFecTer.Text)), "yyyymmdd")
    Periodo = Mid(vlFechaHasta, 1, 6)

    ruta = App.Path & "\PDF\" & Periodo & "\Boleta_" & numPol & ".pdf"
    LNGa = CreateFieldDefFile(rsLiq, Replace(UCase(strRpt & "Estructura\PP_Rpt_LiquidacionRV.rpt"), ".RPT", ".TTX"), 1)
    If objRep.CargaReporte_toPdf(strRpt & "", "PP_Rpt_LiquidacionRV.rpt", "Informe de Liquidación de Rentas Vitalicias", Periodo, numPol, rsLiq, True, ruta, _
                            ArrFormulas("NombreCompania", UCase("Protecta SA compañia de Seguros")), _
                            ArrFormulas("rutcliente", vlRutCliente)) = False Then

        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Function
    End If
    
    
    
    desBoleta = "Boleta_" & numPol & ".pdf"
    
'    cadena = "INSERT INTO PP_TMAE_BOLETASPDF VALUES('" & numPol & "','" & Periodo & "', '" & vlFechaHasta & "','" & desBoleta & "')"
'    vgConexionBD.Execute (cadena)
    
    Screen.MousePointer = 0
    
Exit Function
Err_VenTut:
    Screen.MousePointer = 0
    MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
End Function
Private Sub Cmd_Imprimir_Click()
Call flInformeLiqPago
flCargarListaBox
End Sub


Private Sub cmdExportatxt_Click()
On Error GoTo mierror
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim cadena As String
Dim sNombreArchivo As String
Dim sNombreArchivo2 As String



    Dim vlSQLQuery As String
    Dim vgRutCliente As String
    Dim vgDgvCliente As String
    Dim vlTipoId As String
    Dim cadenaPolizas As String
    
 Screen.MousePointer = 11
    
    Call flRutCliente
    
    vgRutCliente = ""
    vgDgvCliente = ""
    
    vlFechaDesde = Format(CDate(Trim(Txt_LiqFecIni.Text)), "yyyymmdd")
    vlFechaHasta = Format(CDate(Trim(Txt_LiqFecTer.Text)), "yyyymmdd")
    vlPago = "T"
    
    cadenaPolizas = flObtenerPolizas()
    
    If cadenaPolizas = "(' ')" Then
        MsgBox ("Debe elegir polizas para la impresiòn.")
        Exit Sub
    End If
    
    If Not flLlenaTemporal_masivo(vlFechaDesde, vlFechaHasta, cadenaPolizas, clOpcionDEF, vlPago) Then
        Exit Sub
    End If

cadena = "select a.*, gls_comuna, gls_provincia, gls_region, g.gls_tipoiden, g.gls_tipoidencor,"
cadena = cadena & " c.num_idenben, c.gls_nomben as gls_nomben1, c.gls_nomsegben, c.gls_patben, c.gls_matben, gls_dirben, h.gls_tipoiden as tipoRec, i.gls_tipoiden as tipoTit"
'mv 20190627
cadena = cadena & " ,'20517207331'as Ruc_Compañia,'Protecta SA compañia de Seguros' AS Nom_Compañia,c.cod_sexo,nvl(c.gls_correoben,' ')as gls_correoben"
'cadena = cadena & " ,'20517207331'as Ruc_Compañia,'Protecta SA compañia de Seguros' AS Nom_Compañia,c.cod_sexo,decode(c.ind_bolelec,'S',nvl(c.gls_correoben,''),'')as gls_correoben"
cadena = cadena & " from pp_ttmp_liquidacion a"
cadena = cadena & " join pp_tmae_poliza b on a.num_poliza=b.num_poliza " 'and a.num_endoso=b.num_endoso"
cadena = cadena & " join pp_tmae_ben c on a.num_poliza=c.num_poliza and a.num_endoso=c.num_endoso and a.num_orden=c.num_orden"
cadena = cadena & " join ma_tpar_comuna d on c.cod_direccion=d.cod_direccion"
cadena = cadena & " join ma_tpar_provincia e on d.cod_provincia=e.cod_provincia"
cadena = cadena & " join ma_tpar_region f on e.cod_region=f.cod_region"
cadena = cadena & " join ma_tpar_tipoiden g on c.cod_tipoidenben=g.cod_tipoiden"
cadena = cadena & " join ma_tpar_tipoiden h on a.cod_tipoidenreceptor=h.cod_tipoiden"
cadena = cadena & " join ma_tpar_tipoiden i on a.cod_tipoidentit=i.cod_tipoiden"
cadena = cadena & " where cod_usuario='" & vgUsuario & "'"
cadena = cadena & " and a.num_poliza in (select num_poliza from pp_tmae_liqpagopendef where num_perpago=a.num_perpago and cod_tipopago<>'P')"
'mvg 20170904
cadena = cadena & " and b.num_endoso = (select max(num_endoso) from pp_tmae_poliza where num_poliza=b.num_poliza)"
'cadena = cadena & " and not a.num_poliza in(select num_poliza from pp_ttmp_liquidacion where cod_usuario='" & vgUsuario & "' and num_item>1 )"
cadena = cadena & " and a.num_item=1"
'mv 20190627
'cadena = cadena & " and c.ind_bolelec='S'"
cadena = cadena & " and NVL(c.ind_bolelec, 'N') in ('S','N')"
cadena = cadena & " order by 3"

Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open cadena, vgConexionBD, adOpenStatic, adLockReadOnly

'--***** boletas especiales con mas de un detalle

cadena = "select a.*, gls_comuna, gls_provincia, gls_region, g.gls_tipoiden, g.gls_tipoidencor,"
cadena = cadena & " c.num_idenben, c.gls_nomben as gls_nomben1, c.gls_nomsegben, c.gls_patben, c.gls_matben, gls_dirben, h.gls_tipoiden as tipoRec, i.gls_tipoiden as tipoTit"
'mv 20190627
'cadena = cadena & " ,'20517207331'as Ruc_Compañia,'Protecta SA compañia de Seguros' AS Nom_Compañia,c.cod_sexo,decode(c.ind_bolelec,'S',nvl(c.gls_correoben,''),'')as gls_correoben"
cadena = cadena & " ,'20517207331'as Ruc_Compañia,'Protecta SA compañia de Seguros' AS Nom_Compañia,c.cod_sexo,nvl(c.gls_correoben,' ')as gls_correoben"
cadena = cadena & " from pp_ttmp_liquidacion a"
cadena = cadena & " join pp_tmae_poliza b on a.num_poliza=b.num_poliza " 'and a.num_endoso=b.num_endoso"
cadena = cadena & " join pp_tmae_ben c on a.num_poliza=c.num_poliza and a.num_endoso=c.num_endoso and a.num_orden=c.num_orden"
cadena = cadena & " join ma_tpar_comuna d on c.cod_direccion=d.cod_direccion"
cadena = cadena & " join ma_tpar_provincia e on d.cod_provincia=e.cod_provincia"
cadena = cadena & " join ma_tpar_region f on e.cod_region=f.cod_region"
cadena = cadena & " join ma_tpar_tipoiden g on c.cod_tipoidenben=g.cod_tipoiden"
cadena = cadena & " join ma_tpar_tipoiden h on a.cod_tipoidenreceptor=h.cod_tipoiden"
cadena = cadena & " join ma_tpar_tipoiden i on a.cod_tipoidentit=i.cod_tipoiden"
cadena = cadena & " where cod_usuario='" & vgUsuario & "'"
cadena = cadena & " and a.num_poliza in (select num_poliza from pp_tmae_liqpagopendef where num_perpago=a.num_perpago and cod_tipopago<>'P')"
'mvg 20170904
'cadena = cadena & " and a.num_poliza in(select num_poliza from pp_ttmp_liquidacion where cod_usuario='" & vgUsuario & "' and num_item>1 )"
cadena = cadena & " and a.num_item=1"
'mv 20190627
'cadena = cadena & " and c.ind_bolelec='S'"
cadena = cadena & " and NVL(c.ind_bolelec, 'N') in ('S','N')"
cadena = cadena & " order by 3"

Set rs2 = New ADODB.Recordset
rs2.CursorLocation = adUseClient
rs2.Open cadena, vgConexionBD, adOpenForwardOnly, adLockReadOnly

If Not rs.EOF Then
    CommonDialog1.CancelError = True
    CommonDialog1.FileName = "Polizas Electronicas.txt"
    CommonDialog1.DialogTitle = "Boletas Electrónicas"
    CommonDialog1.Filter = "*.txt"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    CommonDialog1.ShowSave
    sNombreArchivo = CommonDialog1.FileName
End If
If Not rs2.EOF Then
    CommonDialog2.CancelError = True
    CommonDialog2.FileName = "Polizas Electronicas_VariosItem.txt"
    CommonDialog2.DialogTitle = "Boletas Electrónicas con mas de un detalle"
    CommonDialog2.Filter = "*.txt"
    CommonDialog2.FilterIndex = 1
    CommonDialog2.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    CommonDialog2.ShowSave
    sNombreArchivo2 = CommonDialog2.FileName
End If

'mvg 20170904
If Not rs2.EOF Then
    Call Exportar_Recordset(rs2, sNombreArchivo2, "|")
    rs2.MoveFirst
    Call crea_excell(rs2, sNombreArchivo2, "Polizas Electronicas_VariosItem.txt")
    rs2.MoveFirst
    Call crea_pdf(rs2, sNombreArchivo2, "Polizas Electronicas_VariosItem.txt")
    
End If

If Not rs.EOF Then
    Call Exportar_Recordset(rs, sNombreArchivo, "|")
End If

If Not rs.EOF And Not rs2.EOF Then
    MsgBox "No hay pólizas con marca a envio por correo.", vbInformation, ""
    Exit Sub
End If

Screen.MousePointer = 1
ProgressBar1.Visible = False
MsgBox "Generación exitosa", vbInformation, ""
    

Exit Sub
mierror:
    MsgBox "Nose pudo exportar consultar con sistemas", vbInformation, ""
    Screen.MousePointer = 1
End Sub


'Private Function PneContraseñaPDF()
'
'End Function
'On Error GoTo ErrorHandler
'Dim ObjWord As Word.Application
'Dim objWordDoc As Word.Document
'
'Set ObjWord = CreateObject("Word.Application")
'ObjWord.Visible = True
'Set objWordDoc = ObjWord.Documents.Open(DocPath)
'objWordDoc.ExportAsFixedFormat OutputFileName:= _
'sDests & sDestsPDFFile, ExportFormat:=wdExportFormatPDF, _
'OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, Range:= _
'wdExportAllDocument, From:=1, To:=1, Item:=wdExportDocumentContent, _
'IncludeDocProps:=True, KeepIRM:=True, CreateBookmarks:= _
'wdExportCreateNoBookmarks, DocStructureTags:=True, BitmapMissingFonts:= _
'True, UseISO19005_1:=False
''objWordDoc.Password = "ita@123"
'
'ObjWord.Quit
'Set ObjWord = Nothing
'Exit Function




Private Function CreaNombreMes(mes As String, anio As String) As String

Select Case mes
Case "01"
    CreaNombreMes = "Ene"
Case "02"
    CreaNombreMes = "Feb"
Case "03"
    CreaNombreMes = "Mar"
Case "04"
    CreaNombreMes = "Abr"
Case "05"
    CreaNombreMes = "May"
Case "06"
    CreaNombreMes = "Jun"
Case "07"
    CreaNombreMes = "Jul"
Case "08"
    CreaNombreMes = "Ago"
Case "09"
    CreaNombreMes = "Set"
Case "10"
    CreaNombreMes = "Oct"
Case "11"
    CreaNombreMes = "Nov"
Case "12"
    CreaNombreMes = "Dic"
End Select

CreaNombreMes = CreaNombreMes & anio
End Function

Private Sub crea_pdf(rs As ADODB.Recordset, ruta As String, NombreArchivo As String)
Dim objRep As New ClsReporte
Dim LNGa As Long
Dim sName As String
Dim RutaBol As String
Dim RutaArchivo As String
Dim Nombre As String

On Error GoTo mierror

RutaArchivo = Replace(ruta, NombreArchivo, "Mail_")
RutaBol = Replace(ruta, NombreArchivo, "Bol_")
Nombre = CreaNombreMes(Mid(Txt_LiqFecTer.Text, 4, 2), Mid(Txt_LiqFecTer.Text, 9, 4))

sName = Dir(RutaBol & Nombre, vbDirectory)

If sName = "" Then
    MkDir (RutaBol & Nombre)
End If

'genera archivo pdf
LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PP_Rpt_LiquidacionRV.rpt"), ".RPT", ".TTX"), 1)
If objRep.CargaReporte_toPdf(strRpt & "", "PP_Rpt_LiquidacionRV.rpt", "Informe de Liquidación de Rentas Vitalicias", rs.Fields(1), rs.Fields(2), rs, True, RutaBol & Nombre & "\protecta_" & rs.Fields(2).Value & "_" & rs.Fields(4).Value & ".pdf", _
                        ArrFormulas("NombreCompania", rs.Fields(54).Value), _
                        ArrFormulas("rutcliente", rs.Fields(53).Value)) = False Then

    MsgBox "No se pudo abrir el reporte", vbInformation
    Exit Sub
End If
            
Exit Sub
mierror:
MsgBox "Hay un problema con la creación del archivo excell de configuracion de envio de boletas. consulte con sistemas.", vbExclamation, ""
         
            
End Sub
Private Sub crea_excell(rs As ADODB.Recordset, ruta As String, NombreArchivo As String)

Dim iFreeFile   As Integer
Dim iField      As Long
Dim i           As Long
Dim obj_Field   As ADODB.Field
Dim xlapp As New Excel.Application
Dim vArchivo As String
Dim Nombre As String
Dim objRep As New ClsReporte
Dim LNGa As Long
Dim sName As String

On Error GoTo mierror

RutaArchivo = Replace(ruta, NombreArchivo, "Mail_")
Nombre = CreaNombreMes(Mid(Txt_LiqFecTer.Text, 4, 2), Mid(Txt_LiqFecTer.Text, 9, 4))

Set xlapp = CreateObject("excel.application")
xlapp.Application.Workbooks.Add
'xlapp.Visible = True 'para ver vista previa
'xlapp.WindowState = 1 ' minimiza excel
Set w = xlapp.Worksheets.Item(1)
w.Activate
  
rs.MoveFirst
ProgressBar1.Max = rs.RecordCount
ProgressBar1.Min = 0

For x = 0 To rs.RecordCount
    If x = 0 Then

        w.Cells(1, 1).Value = "ARCHIVO"
        w.Cells(1, 2).Value = "NOMBRE1"
        w.Cells(1, 3).Value = "APELLIDO1"
        w.Cells(1, 4).Value = "APELLIDO2"
        w.Cells(1, 5).Value = "DIRECCION"
        w.Cells(1, 6).Value = "TRATO"
        w.Cells(1, 7).Value = "CAMPO1"
        w.Cells(1, 8).Value = "CAMPO2"
        w.Cells(1, 9).Value = "CAMPO3"
        w.Cells(1, 10).Value = "CAMPO4"
        w.Cells(1, 11).Value = "CAMPO5"
        w.Cells(1, 12).Value = "CAMPO6"
        w.Cells(1, 13).Value = "CAMPO7"
        w.Cells(1, 14).Value = "CAMPO8"
        w.Cells(1, 15).Value = "CAMPO9"
        w.Cells(1, 16).Value = "CAMPO10"
    Else
        
        If rs.Fields(2).Value & "_" & rs.Fields(4).Value <> vArchivo Then
            w.Cells.NumberFormat = "@"
            vArchivo = rs.Fields(2).Value & "_" & rs.Fields(4).Value
            w.Cells(x + 1, 1) = "protecta_" & vArchivo & ".pdf"
            w.Cells(x + 1, 2) = rs.Fields(46).Value
            w.Cells(x + 1, 3) = rs.Fields(48).Value
            w.Cells(x + 1, 4) = rs.Fields(49).Value
            w.Cells(x + 1, 5) = rs.Fields(56).Value
            If rs.Fields(55).Value = "M" Then
                w.Cells(x + 1, 6) = "Estimado"
            Else
                w.Cells(x + 1, 6) = "Estimada"
            End If
    
            ProgressBar1.Value = ProgressBar1.Value + 1
        End If
    End If
    If x > 0 Then
        rs.MoveNext
    End If
    
Next x
 
'xlapp.WindowState = 1

'xlapp.ActiveWorkbook.SaveAs RutaBol & Nombre & "\" & "Mail_" & Nombre & ".xls"
xlapp.ActiveWorkbook.SaveAs RutaArchivo & Nombre & ".xls"
'xlapp.WindowState = 1
xlapp.Application.Quit
Set xlapp = Nothing
Set objExcel = Nothing
Set w = Nothing

Exit Sub
mierror:
MsgBox "Hay un problema con la creación del archivo excell de configuracion de envio de boletas. consulte con sistemas.", vbExclamation, ""

End Sub



Private Function isValidField(obj_Field As ADODB.Field) As Boolean
      
    With obj_Field
        On Error GoTo error_handler
        Select Case obj_Field.Type
            Case adBinary, adIDispatch, adIUnknown, adUserDefined
                isValidField = False
            ' -- Campo válido
            Case Else
                isValidField = True
        End Select
    End With
Exit Function
error_handler:
End Function

Public Function Exportar_Recordset( _
    rs As Object, _
    Optional sFileName As String, _
    Optional sDelimiter As String = " ", _
    Optional bPrintField As Boolean = False) As Boolean
  
    Dim iFreeFile   As Integer
    Dim iField      As Long
    Dim i           As Long
    Dim obj_Field   As ADODB.Field
  
    On Error GoTo error_handler:
      
    Screen.MousePointer = vbHourglass
    ' -- Otener número de archivo disponible
    iFreeFile = FreeFile
    ' -- Crear el archivo
    Open sFileName For Output As #iFreeFile
  
    With rs
        iField = .Fields.Count - 1
        On Error Resume Next
        ' -- Primer registro
        .MoveFirst
        On Error GoTo error_handler
        ' -- Recorremos campo por campo y los registros de cada uno
        Do While Not .EOF
            For i = 1 To iField
                  
                ' -- Asigna el objeto Field
                Set obj_Field = .Fields(i)
                ' -- Verificar que el campo no es de ipo bunario o  un tipo no válido para grabar en el archivo
                If isValidField(obj_Field) Then
                    If i < iField Then
                        If bPrintField Then
                            ' -- Escribir el campo y el valor
                            Print #iFreeFile, obj_Field.Name & ":" & obj_Field.Value & sDelimiter;
                        Else
                            ' -- Guardar solo el valor sin el campo
                            Print #iFreeFile, obj_Field.Value & sDelimiter;
                        End If
                    Else
                        If bPrintField Then
                            ' -- Escribir el nombre del campo y el valor de la última columna ( Sin delimitador y sin punto y coma para añadir nueva línea )
                            Print #iFreeFile, obj_Field.Name & ": " & obj_Field.Value
                        Else
                            ' -- Guardar solo el valor sin el campo
                            Print #iFreeFile, obj_Field.Value
                        End If
                    End If
                End If
            Next
            ' -- Mover el cursor al siguiente registro
            .MoveNext
        Loop
    End With
      
    ' -- Cerrar el recordset
    'rs.Close
    Exportar_Recordset = True
    Screen.MousePointer = vbDefault
    Close #iFreeFile
    Exit Function
error_handler:
 On Error Resume Next
 Close #iFreeFile
 rst.Close
 Screen.MousePointer = vbDefault
End Function

Private Sub Command1_Click()
If Me.txtNroArch.Text = "" Then
    MsgBox "Debe Ingresar el numero de Archivo de Pagos Recurrentes."
    Exit Sub
End If

Dim i As Long

Call flCargarListaBox

'For i = 1 To Lst_LiqSeleccion.ListCount - 1
'    Lst_LiqSeleccion.Selected(i) = True
'Next

End Sub

Private Sub Form_Load()
    Me.Width = 8430
    Me.Height = 5865

End Sub

Private Function flCargarListaBox() As Boolean
    
On Error GoTo Err_flCargarListaConHabDes

    flCargarListaConHabDes = False
    
    Screen.MousePointer = vbHourglass
    
    Lst_LiqSeleccion.Clear
    
    Dim vlFecIni, vlFecTer As String
    
    'vlFecIni = Format(CDate(Trim(Txt_LiqFecIni.Text)), "yyyymmdd")
    'vlFecTer = Format(CDate(Trim(Txt_LiqFecTer.Text)), "yyyymmdd")
    
    'vgSql = ""
    'vgSql = vgSql & " select c.NUM_POLIZA, b.GLS_PATBEN || ' / ' || b.GLS_MATBEN || ' / ' || b.GLS_NOMBEN || ' / ' || b.GLS_NOMSEGBEN as nombres,"
    'vgSql = vgSql & " c.FEC_INICER , c.fec_tercer"
    'vgSql = vgSql & " from pp_tmae_certificado c"
    'vgSql = vgSql & " left join pp_tmae_ben b on c.NUM_POLIZA=b.NUM_POLIZA  and c.NUM_ENDOSO=b.NUM_ENDOSO and c.NUM_ORDEN=b.NUM_ORDEN"
    'vgSql = vgSql & " left join ma_tpar_tipoiden t on  b.cod_tipoidenben=t.cod_tipoiden"
    'vgSql = vgSql & " left join ma_tpar_comuna m on b.cod_direccion=m.cod_direccion"
    'vgSql = vgSql & " where c.fec_tercer>='" & vlFecIni & "' and c.fec_tercer<='" & vlFecTer & "'"
    'vgSql = vgSql & " order by c.NUM_POLIZA"
    'cambio momentan eo
    'vgSql = ""
    'vgSql = vgSql & " select d.num_poliza as num_poliza, d.GLS_NOMBRE as nombres, d.MTO_PAGO, fec_desde, fec_hasta"
    'vgSql = vgSql & " from PP_TMAE_CONTABLEREGPAGO C"
    'vgSql = vgSql & " INNER JOIN PP_TMAE_CONTABLEDETREGPAGO D ON"
    'vgSql = vgSql & " C.NUM_ARCHIVO=D.NUM_ARCHIVO and C.COD_TIPREG=D.COD_TIPREG AND"
    'vgSql = vgSql & " c.COD_TIPMOV = d.COD_TIPMOV And c.Cod_Moneda = d.Cod_Moneda"
    'vgSql = vgSql & " where c.NUM_ARCHIVO ='" & Me.txtNroArch.Text & "' and d.cod_tipmov=38"
    'vgSql = vgSql & " order by 1"

    vgSql = ""
    vgSql = vgSql & " select d.num_poliza as num_poliza, d.GLS_NOMBRE as nombres, min(fec_desde) as fec_desde, max(fec_hasta) as fec_hasta"
    vgSql = vgSql & " from PP_TMAE_CONTABLEREGPAGO C"
    vgSql = vgSql & " INNER JOIN PP_TMAE_CONTABLEDETREGPAGO D ON"
    vgSql = vgSql & " C.NUM_ARCHIVO=D.NUM_ARCHIVO and C.COD_TIPREG=D.COD_TIPREG AND"
    vgSql = vgSql & " c.COD_TIPMOV = d.COD_TIPMOV And c.Cod_Moneda = d.Cod_Moneda"
    vgSql = vgSql & " where c.NUM_ARCHIVO in (" & Me.txtNroArch.Text & ") and d.cod_tipmov=38"
    vgSql = vgSql & " group by d.num_poliza , d.GLS_NOMBRE" ', d.MTO_PAGO "
    vgSql = vgSql & " order by 1"


    Dim Conteo As Integer
    
    Conteo = 0

    Set vgRegistro = vgConexionBD.Execute(vgSql)
    Txt_LiqFecIni = Mid(vgRegistro!fec_desde, 7, 2) & "/" & Mid(vgRegistro!fec_desde, 5, 2) & "/" & Mid(vgRegistro!fec_desde, 1, 4)
    Txt_LiqFecTer = Mid(vgRegistro!FEC_HASTA, 7, 2) & "/" & Mid(vgRegistro!FEC_HASTA, 5, 2) & "/" & Mid(vgRegistro!FEC_HASTA, 1, 4)
    If Not vgRegistro.EOF Then
       Do While Not vgRegistro.EOF
          If Lst_LiqSeleccion.ListCount = 1 Then
            Lst_LiqSeleccion.AddItem (" TODOS "), 0
          End If
          Lst_LiqSeleccion.AddItem (" " & Trim(vgRegistro!num_poliza) & " - " & Trim(vgRegistro!nombres) & " ")
            
          
            
          vgRegistro.MoveNext
          Conteo = Conteo + 1
        
       Loop
       
    End If
    boltodos = False
    'Call cargaReporteFin(vgRegistro)
    lblConteo.Caption = "Existen " & CStr(Conteo) & " para Impresion."
    
    Set vgRegistro = Nothing
    Screen.MousePointer = vbDefault
    
    flCargarListaConHabDes = True

Exit Function
Err_flCargarListaConHabDes:
    Screen.MousePointer = vbDefault
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Private Sub Lst_LiqSeleccion_ItemCheck(Item As Integer)
Dim i As Integer
    If Item = 0 Then
        If Lst_LiqSeleccion.Selected(0) Then
            boltodos = True
            For i = 1 To Lst_LiqSeleccion.ListCount - 1
                Lst_LiqSeleccion.Selected(i) = True
            Next
        Else
            boltodos = False
            For i = 1 To Lst_LiqSeleccion.ListCount - 1
                Lst_LiqSeleccion.Selected(i) = False
            Next
        End If
    Else
        Lst_LiqSeleccion.Selected(0) = False
    End If
End Sub

Function flObtenerPolizas() As String
Dim Cadenasalida As String

On Error GoTo Err_flObtenerConceptos

    vlCodConceptos = "("
    For vgI = 0 To Lst_LiqSeleccion.ListCount - 1
        If Lst_LiqSeleccion.Selected(vgI) Then
           'If vgI <> (Lst_LiqSeleccion.ListCount - 1) Then 'HQR 22/10/2005 Se corrige porque producía error al marcar el último concepto de la lista
              If vlCodConceptos <> "(" Then
                 vlCodConceptos = (vlCodConceptos & ",")
              End If
           'End If
           'If vgI <> 0 Then
           If Trim(Lst_LiqSeleccion.List(vgI)) <> "TODOS" Then
                vlCodConceptos = (vlCodConceptos & "'" & (Trim(Mid(Lst_LiqSeleccion.List(vgI), 1, (InStr(1, Lst_LiqSeleccion.List(vgI), "-") - 1)))) & "'")
           End If
           
'              vgSql = vgSql & "'" & (Trim(Mid(Lst_LiqSeleccion.List(vgI), 1, (InStr(1, Lst_LiqSeleccion.List(vgI), "-") - 1)))) & "', "
           'End If
           
        End If
        If vgI = (Lst_LiqSeleccion.ListCount - 1) Then
           vlCodConceptos = (vlCodConceptos & ")")
'        Else
'            If vlCodConceptos <> "(" Then
'               vlCodConceptos = (vlCodConceptos & ",")
'            End If
        End If
    Next
    If vlCodConceptos = "()" Then
       vlCodConceptos = "(' ')"
    End If
    
    flObtenerPolizas = vlCodConceptos
    
Exit Function
Err_flObtenerConceptos:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function
Private Sub cmd_Marcar_Click()

Call MarcarDesde

'Dim vgI As Integer

'For vgI = 0 To Lst_LiqSeleccion.ListCount - 1
'    If Trim(Lst_LiqSeleccion.List(vgI)) <> "TODOS" Then
'        If Trim(Mid(Lst_LiqSeleccion.List(vgI), 1, (InStr(1, Lst_LiqSeleccion.List(vgI), "-") - 1))) = Trim(txt_polizas.Text) Then
'            Lst_LiqSeleccion.Selected(vgI) = True
'        End If
'    End If
'Next

End Sub

Private Sub MarcarDesde()

Dim vgI As Integer
Dim valDes, valHas As Long

valDes = txtDesdev.Text
valHas = txtHastav.Text

If valDes = 1 Then
        For vgI = 0 To valHas
            
            Lst_LiqSeleccion.Selected(vgI) = True
          
        Next
Else
        For vgI = valDes To valHas
            
            Lst_LiqSeleccion.Selected(vgI) = True
          
        Next
End If





End Sub

Private Sub txtNroArch_KeyPress(KeyAscii As Integer)
On Error GoTo Err_TxtPenPolizaKeyPress

    If KeyAscii = 13 Then
       Command1.SetFocus
    End If
    
Exit Sub
Err_TxtPenPolizaKeyPress:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

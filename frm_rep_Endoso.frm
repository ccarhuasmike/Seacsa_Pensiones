VERSION 5.00
Begin VB.Form frm_rep_Endoso 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Endosos de Fallecimiento y Activación de Sobrevivencia"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Imprimir"
      Height          =   465
      Left            =   1995
      TabIndex        =   0
      Top             =   495
      Width           =   2205
   End
End
Attribute VB_Name = "frm_rep_Endoso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()

    Dim objRep As New ClsReporte
    Dim cadena As String
    Dim rs As ADODB.Recordset
    Dim vlReporte, vlTitulo  As String
    
    On Error GoTo mierror

    If Me.Tag = "A" Then
        cadena = " select distinct p.num_poliza,p.cod_afp,b.gls_nomben,nvl(b.gls_nomsegben,' ')as gls_nomsegben,b.gls_patben,b.gls_matben,b.num_orden,b.cod_estpension,"
        cadena = cadena & "  to_date(e.fec_endoso,'yyyymmdd')as fec_endoso,to_date(e.fec_efecto,'yyyymmdd')as fec_efecto,e.cod_cauendoso,e.num_endoso,to_date(e.fec_finefecto,'yyyymmdd')as fec_finefecto,"
        cadena = cadena & " (select gls_elemento from ma_tpar_tabcod where cod_tabla='CE' and cod_elemento=e.cod_cauendoso)as Desc_endoso,"
        cadena = cadena & " (select gls_elemento from ma_tpar_tabcod where cod_tabla='AF' and cod_elemento=p.cod_afp)as Desc_Afp,"
        cadena = cadena & " (select gls_elemento from ma_tpar_tabcod where cod_tabla='VI' and cod_elemento=b.cod_estpension)as Estado"
        cadena = cadena & " from pp_tmae_ben b inner join pp_tmae_endoso e on b.num_poliza=e.num_poliza"
        cadena = cadena & " inner join pp_tmae_poliza p on b.num_poliza=p.num_poliza"
        cadena = cadena & " where b.num_endoso=(select max(num_endoso) from pp_tmae_poliza where num_poliza=b.num_poliza)"
        cadena = cadena & " and e.cod_cauendoso in ('07','08') order by p.num_poliza,b.num_orden,e.num_endoso"
        vlReporte = "PP_Rpt_Endosos.rpt"
        vlTitulo = "Informe de Endosos por Fallecimiento y Activación de Sobrevivencia"
    Else
        cadena = " select distinct p.num_poliza, b.num_orden, b.gls_nomben,nvl(b.gls_nomsegben,' ')as gls_nomsegben,b.gls_patben,b.gls_matben, b.cod_estpension,b.num_idenben,"
        cadena = cadena & " q.gls_elemento as Desc_Afp,"
        cadena = cadena & " (select gls_elemento from ma_tpar_tabcod where cod_tabla='VI' and cod_elemento=b.cod_estpension)as Estado"
        cadena = cadena & " from pp_tmae_ben b"
        cadena = cadena & " join pp_tmae_poliza p on b.num_poliza=p.num_poliza and b.num_endoso=p.num_endoso"
        cadena = cadena & " join ma_tpar_tabcod q on cod_elemento=p.cod_afp and cod_tabla='AF'"
        cadena = cadena & " where b.num_endoso=(select max(num_endoso) from pp_tmae_poliza where num_poliza=p.num_poliza)"
        cadena = cadena & " and cod_par in (30) and b.cod_tipoidenben=6"
        cadena = cadena & " order by 1,2"
        vlReporte = "PP_Rpt_PartBen.rpt"
        vlTitulo = "Informe de Beneficiarios con Partida de Nacimiento"
    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open cadena, vgConexionBD, adOpenStatic, adLockReadOnly
    
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\" & vlReporte), ".RPT", ".TTX"), 1)
    
    If objRep.CargaReporte(strRpt & "", vlReporte, vlTitulo, rs, True, _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("NombreSistema", vgNombreSistema), _
                            ArrFormulas("NombreSubSistema", vgNombreSubSistema)) = False Then
                            
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
    End If
    
    Exit Sub
mierror:
    MsgBox "No se pudo cargar el reporte", vbInformation
    
    
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Frm_Menu.Tag = "0"
End Sub

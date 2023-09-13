VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_CargarExcel_Superv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Migración "
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15840
   Icon            =   "frm_CargarExcel_Superv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   15840
   Begin VB.CommandButton Command6 
      Caption         =   "<<"
      Height          =   735
      Left            =   8280
      TabIndex        =   17
      Top             =   3480
      Width           =   615
   End
   Begin VB.Frame Frame3 
      Caption         =   "Listado de Datos de la AFP"
      Height          =   8175
      Left            =   9000
      TabIndex        =   10
      Top             =   120
      Width           =   6735
      Begin VB.CommandButton Command3 
         Caption         =   "Reporte"
         DownPicture     =   "frm_CargarExcel_Superv.frx":0442
         Height          =   810
         Left            =   5040
         Picture         =   "frm_CargarExcel_Superv.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   7200
         Width           =   690
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Migrar"
         Height          =   810
         Left            =   5760
         Picture         =   "frm_CargarExcel_Superv.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   7200
         Width           =   825
      End
      Begin VB.Frame Frame1 
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   6435
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Height          =   285
            Left            =   5400
            TabIndex        =   13
            Top             =   120
            Width           =   390
         End
         Begin VB.TextBox txtRuta 
            Height          =   285
            Left            =   1410
            TabIndex        =   12
            Top             =   120
            Width           =   3975
         End
         Begin VB.Label Label1 
            Caption         =   "Buscar archivo:"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   120
            Width           =   1275
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   6375
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   11245
         _Version        =   393216
         AllowBigSelection=   -1  'True
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Listado Pensionados con Certifiados Vencidos"
      Height          =   8175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin VB.CheckBox chkTodos 
         Caption         =   "Todos"
         Height          =   195
         Left            =   4200
         TabIndex        =   19
         Top             =   600
         Width           =   975
      End
      Begin VB.PictureBox picChecked 
         Height          =   250
         Left            =   0
         Picture         =   "frm_CargarExcel_Superv.frx":1108
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   250
      End
      Begin VB.PictureBox picUnchecked 
         Height          =   250
         Left            =   0
         Picture         =   "frm_CargarExcel_Superv.frx":144A
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   250
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Buscar Certificados Vencidos este mes"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   3855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Renovar"
         Height          =   810
         Left            =   7080
         Picture         =   "frm_CargarExcel_Superv.frx":178C
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   7200
         Width           =   825
      End
      Begin VB.TextBox txtDniBus 
         Height          =   285
         Left            =   6360
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   6375
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   11245
         _Version        =   393216
         AllowBigSelection=   -1  'True
      End
      Begin VB.Label lblConteo 
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   7320
         Width           =   3495
      End
      Begin VB.Label Label2 
         Caption         =   "Buscar DNI"
         Height          =   255
         Left            =   5040
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.ListBox Lst_LiqSeleccion 
      Height          =   4335
      ItemData        =   "frm_CargarExcel_Superv.frx":1BCE
      Left            =   16440
      List            =   "frm_CargarExcel_Superv.frx":1BD0
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   3120
      Width           =   3045
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   15240
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frm_CargarExcel_Superv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ruta As String
Dim rsLoad As ADODB.Recordset
Dim cad As String
Dim rs As ADODB.Recordset


Private Sub Command1_Click()

    
CommonDialog1.ShowOpen
ruta = CommonDialog1.FileName
txtRuta.Text = ruta
If txtRuta.Text <> txtRuta.Tag Then
    Command2.Enabled = True
    Command3.Enabled = False
End If

End Sub


Private Function CargaExcel(ByVal ruta As String) As Boolean
 'Variable de tipo Aplicación de Excel
    Dim objExcel As Excel.Application
    'Una variable de tipo Libro de Excel
    Dim xLibro As Excel.Workbook
    Dim Col As Integer, fila As Integer
    Dim dato As String
    Dim xHoja As Excel.Worksheet
    Dim Columnas As Integer, Filas As Integer
    Dim Sql As String
    'creamos un nuevo objeto excel
    Set objExcel = New Excel.Application
  
    'Usamos el método open para abrir el archivo que está en el directorio del programa xxx.xls
    Set xLibro = objExcel.Workbooks.Open(ruta)
    Set xHoja = xLibro.Sheets(1)

    'determinamos la cantidad de columnas y filas de la hoja
    'Columnas = xHoja.Rows(3).Find("").Column
    'Filas = xHoja.Columns(3).Find("").Row
    Columnas = 2
    Filas = 2
    'Hacemos el Excel Visible SI LO DESEAMOS
    'objExcel.Visible = True
  
    With xLibro
        ' Hacemos referencia a la Hoja
        With .Sheets(1)
            vgConexionBD.BeginTrans
            'Recorremos la fila desde la 6 hasta la 7
            Do While .Cells(Filas, 1) <> ""
                
                On Error GoTo mierror
                'Sql = "insert into  PT_TMP_SUPV" & vgUsuario & "(AFP,NOMBRE,NUM_DOC,ESTADO,RUTA,FECHA_RECEP,F_DESDE,F_HASTA)" & _
                '        "VALUES('" & .Cells(Filas, 1) & "','" & .Cells(Filas, 2) & "','" & .Cells(Filas, 3) & "','','" & Trim(txtRuta.Text) & "'," & Format(.Cells(Filas, 4), "yyyymm") & "," & Format(.Cells(Filas, 5), "yyyymm") & "," & Format(.Cells(Filas, 6), "yyyymm") & ")"
                        
                Sql = "insert into  PT_TMP_SUPV" & vgUsuario & "(AFP,NOMBRE,NUM_DOC,ESTADO,RUTA,FECHA_RECEP,F_DESDE,F_HASTA, COD_PAR)" & _
                      "VALUES('" & .Cells(Filas, 1) & "','" & .Cells(Filas, 2) & "','" & .Cells(Filas, 3) & "','','" & Trim(txtRuta.Text) & "','" & .Cells(Filas, 6) & "','" & .Cells(Filas, 7) & "','" & .Cells(Filas, 8) & "' , '" & .Cells(Filas, 9) & "')"
                                
                vgConexionBD.Execute (Sql)

                Filas = Filas + 1
            Loop
           vgConexionBD.CommitTrans
        End With
    End With
  
    'Eliminamos los objetos si ya no los usamos
    Set objExcel = Nothing
    xLibro.Close
    Set xLibro = Nothing
    
    cad = "SELECT '>' sel, NUM_DOC, AFP, NOMBRE FROM PT_TMP_SUPVPARANAB"
 
    Set rs = vgConexionBD.Execute(cad)
    If Not rs.EOF Then
        Call Cargar_FlexGridExcel(MSFlexGrid2, rs)
    End If
    
    
    CargaExcel = True
Exit Function
mierror:
    vgConexionBD.RollbackTrans
    MsgBox "No se pudo insertar al temporal fila:" & CStr(Filas) & Err.Description, vbCritical
    
End Function

Private Function Crea_temporal() As Boolean
Dim cad As String
Dim rs As ADODB.Recordset
On Error GoTo mierror

cad = "select * from all_tables where TABLE_NAME='PT_TMP_SUPV" & vgUsuario & "'"
Set rs = vgConexionBD.Execute(cad)
If Not rs.EOF Then
    cad = "DELETE FROM PT_TMP_SUPV" & vgUsuario
    vgConexionBD.Execute (cad)
Else
    cad = "create table PT_TMP_SUPV" & vgUsuario & "(AFP VARCHAR2(50),NOMBRE VARCHAR2(100),NUM_DOC VARCHAR2(12),ESTADO VARCHAR2(1),RUTA VARCHAR2(2000),FECHA_RECEP varchar2(10),F_DESDE varchar2(10),F_HASTA varchar2(10), COD_PAR varchar2(2))"
    vgConexionBD.Execute (cad)
End If
Crea_temporal = True
Exit Function
mierror:
    MsgBox "Problemas al crear el temporal", vbCritical
End Function

Private Function INSERTA_DATOS() As Boolean

Dim CM As ADODB.Command
Set CM = New ADODB.Command
On Error GoTo mierror

CM.ActiveConnection = vgConexionBD
CM.CommandType = adCmdStoredProc
CM.CommandText = "PD_INSERT_RENT_VITAL"
CM.Parameters.Append CM.CreateParameter("USUARIO", adChar, adParamInput, 50, vgUsuario)
CM.Execute

If MsgBox("Proceso terminado. Desea ver el reporte de resultado", vbInformation + vbYesNo, "Migración de certificados") = vbYes Then
    Call Imprimir
End If

INSERTA_DATOS = True
Exit Function
mierror:
    MsgBox "No se pudo insertar, consulte con el area de sistemas" & Err.Description, vbInformation
    
End Function
Private Sub Imprimir()

Dim cadena As String
Dim rs As ADODB.Recordset
Dim objRep As New ClsReporte
On Error GoTo mierror

cadena = "SELECT B.NUM_POLIZA,B.NUM_ORDEN,B.NUM_IDENBEN,B.GLS_NOMBEN,B.GLS_PATBEN,B.GLS_MATBEN,T.ESTADO"
cadena = cadena & " FROM PT_TMP_SUPV" & vgUsuario & " T LEFT JOIN PP_TMAE_BEN B  ON T.NUM_DOC=B.NUM_IDENBEN "
cadena = cadena & " WHERE B.COD_TIPOIDENBEN='1'"
cadena = cadena & " group by B.NUM_POLIZA,B.NUM_ORDEN,B.NUM_IDENBEN,B.GLS_NOMBEN,B.GLS_PATBEN,B.GLS_MATBEN,T.ESTADO"
cadena = cadena & " UNION ALL "
cadena = cadena & " SELECT '' AS NUM_POLIZA,0 AS NUM_ORDEN,T.NUM_DOC AS NUM_IDENBEN,T.NOMBRE AS GLS_NOMBEN,"
cadena = cadena & " '' AS GLS_PATBEN,'' AS GLS_MATBEN,T.ESTADO FROM PT_TMP_SUPV" & vgUsuario & " T WHERE T.ESTADO='N'"

Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open cadena, vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PP_Rpt_Cert_Migracion.rpt"), ".RPT", ".TTX"), 1)
    
        
    If objRep.CargaReporte(strRpt & "", "PP_Rpt_Cert_Migracion.rpt", "Informe Certificados de Supervivencia Vencidos", rs, True, _
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

Private Sub Command2_Click()
Screen.MousePointer = 11
If Crea_temporal = False Then
    Exit Sub
End If

If CargaExcel(ruta) = False Then
    Exit Sub
End If

'If INSERTA_DATOS = False Then
'    Exit Sub
'End If

Screen.MousePointer = 1
'Unload Me

End Sub

Private Sub Command3_Click()
Call Imprimir
End Sub

Private Sub Command4_Click()
    Call flCargarListaBox
End Sub

Private Function CargaDatosTmp(ByVal ruta As String) As Boolean

Dim numPol As String
Dim vlAfp, vlNombres, vlDNI, vlFecIni, vlFecFin, vlCodPar As String




On Error GoTo Err_flObtenerConceptos

    'For vgI = 0 To Lst_LiqSeleccion.ListCount - 1
    '    If Lst_LiqSeleccion.Selected(vgI) Then
    '
    '       If Trim(Lst_LiqSeleccion.List(vgI)) <> "TODOS" Then
    '            numPol = (Trim(Mid(Lst_LiqSeleccion.List(vgI), 1, (InStr(1, Lst_LiqSeleccion.List(vgI), "-") - 1))))
    '
    '
    '            Sql = "insert into  PT_TMP_SUPV" & vgUsuario & "(AFP,NOMBRE,NUM_DOC,ESTADO,RUTA,FECHA_RECEP,F_DESDE,F_HASTA, COD_PAR)" & _
    '                  "VALUES('" & .Cells(Filas, 1) & "','" & .Cells(Filas, 2) & "','" & .Cells(Filas, 3) & "','','" & Trim(txtRuta.Text) & "','" & .Cells(Filas, 4) & "','" & .Cells(Filas, 5) & "','" & .Cells(Filas, 6) & "' , '" & .Cells(Filas, 7) & "')"
    '
    '
    '       End If
    '
    '    End If
    'Next
    
Exit Function
Err_flObtenerConceptos:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function


Private Function flCargarListaBox() As Boolean
Dim vlFecPago As String
On Error GoTo Err_flCargarListaConHabDes

    'flCargarListaConHabDes = False
    
    Screen.MousePointer = vbHourglass
    
    Lst_LiqSeleccion.Clear
    
    Dim vlFecIni, vlFecTer As String
    
    'vlFecIni = Format(CDate(Trim(Txt_LiqFecIni.Text)), "yyyymmdd")
    'vlFecTer = Format(CDate(Trim(Txt_LiqFecTer.Text)), "yyyymmdd")

    vgSql = " SELECT fec_pagoreg FROM PP_TMAE_PROPAGOPEN WHERE cod_estadoreg <> 'C' ORDER BY num_perpago ASC"
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
        vlFecPago = vgRegistro!Fec_PagoReg
    End If


    'vgSql = " select 0 sel, a.num_poliza as poliza, a.num_idenben as DNI, a.gls_nomben || ' ' || a.gls_patben || ' ' || gls_matben nombres , to_date(a.fec_tercer,'YYYYMMDD') fechaFin"
    'vgSql = vgSql & " From"
    'vgSql = vgSql & " (select a.*, b.gls_nomben, b.gls_patben, b.gls_matben, b.num_idenben from pp_tmae_certificado a"
    'vgSql = vgSql & " Join"
    'vgSql = vgSql & " ("
    'vgSql = vgSql & " select * from pp_tmae_ben a where num_endoso=(select max(num_endoso) from pp_tmae_poliza where num_poliza=a.num_poliza)"
    'vgSql = vgSql & " and a.cod_estpension=99 and a.cod_derpen=99"
    'vgSql = vgSql & " ) b on a.num_poliza=b.num_poliza and a.num_endoso=b.num_endoso and a.num_orden=b.num_orden"
    'vgSql = vgSql & " where fec_tercer=(select max(fec_tercer) from pp_tmae_certificado where num_poliza=a.num_poliza)"
    ''vgSql = vgSql & " ) a where fec_tercer <='" & vlFecPago & "' order by 2"
    'vgSql = vgSql & " ) a where fec_tercer <='" & vlFecPago & "' order by 2"
    '----------MateriaGris JaimeRios 01/02/2018----------'
    '-------------certificados sin endosos-------------
    vgSql = "select 0 sel, a.num_poliza as poliza, a.num_idenben as DNI,  a.gls_nomben || ' ' || a.gls_patben || ' ' || gls_matben nombres , to_date(a.fec_tercer,'YYYYMMDD') fechaFin, a.cod_tipo"
    vgSql = vgSql & " From (select a.*, b.gls_nomben, b.gls_patben, b.gls_matben, b.num_idenben"
    vgSql = vgSql & "       from pp_tmae_certificado a Join ( select * from"
    vgSql = vgSql & "                                         ( select a.*,b.cod_cauendoso"
    vgSql = vgSql & "                                           from pp_tmae_ben a  left join pp_tmae_endoso b"
    vgSql = vgSql & "                                                                on a.num_poliza=b.num_poliza"
    vgSql = vgSql & "                                                                and a.num_endoso=b.num_endoso+1"
    vgSql = vgSql & "                                           where a.num_endoso=(select max(num_endoso)"
    vgSql = vgSql & "                                                               From pp_tmae_poliza"
    vgSql = vgSql & "                                                               where num_poliza=a.num_poliza)"
    vgSql = vgSql & "                                             and a.cod_estpension = 99"
    vgSql = vgSql & "                                             and a.cod_derpen     = 99"
    vgSql = vgSql & "                                           ) a where a.cod_cauendoso is null"
    vgSql = vgSql & "                                        ) b on a.num_poliza=b.num_poliza"
    vgSql = vgSql & "                                        and a.num_orden=b.num_orden"
    vgSql = vgSql & "        where fec_tercer=(select max(fec_tercer) from pp_tmae_certificado where num_poliza=a.num_poliza and num_orden=a.num_orden)"
    vgSql = vgSql & "        and ("
    vgSql = vgSql & "         b.cod_par<>30 or" '-- familiares no hijos
    vgSql = vgSql & "         (b.cod_par=30 and (months_between(TRUNC(sysdate),to_date(b.fec_nacben,'yyyymmdd'))/12)<18 and a.cod_tipo='SUP') or" '--hijos menores de edad
    vgSql = vgSql & "         (b.cod_par=30 and (months_between(TRUNC(sysdate),to_date(b.fec_nacben,'yyyymmdd'))/12)>=18 and" '--hijos mayores de edad
    'vgSql = vgSql & "          a.cod_tipo='EST' and a.est_act=1 and a.ind_dni='S' and a.ind_dju='S' and a.ind_pes='S' and a.ind_bno='S')" '--certificado de estudios vigente
    vgSql = vgSql & "         ( ((select fec_dev from pp_tmae_poliza p where p.num_poliza=b.num_poliza and p.num_endoso=b.num_endoso) >= '20130801' and "
    vgSql = vgSql & "            a.cod_tipo='EST' and a.est_act=1 and a.ind_dni='S' and a.ind_dju='S' and a.ind_pes='S' and a.ind_bno='S') or (b.cod_sitinv<>'N' and a.cod_tipo='SUP'))) " '--certificado de estudios vigente  o situacion de invalides (materiagris 01/03/2018)
    vgSql = vgSql & "      )) a"
    vgSql = vgSql & " where fec_tercer <='" & vlFecPago & "' "
    vgSql = vgSql & " Union All"
    '-------------certificados cod_cauendoso <> '13'-------------
    vgSql = vgSql & " select 0 sel, a.num_poliza as poliza, a.num_idenben as DNI, a.gls_nomben || ' ' || a.gls_patben || ' ' || gls_matben nombres , to_date(a.fec_tercer,'YYYYMMDD') fechaFin, a.cod_tipo"
    vgSql = vgSql & " From (select a.*, b.gls_nomben, b.gls_patben, b.gls_matben, b.num_idenben"
    vgSql = vgSql & "       from pp_tmae_certificado a Join (  select a.*"
    vgSql = vgSql & "                                           from pp_tmae_ben a  inner join pp_tmae_endoso b"
    vgSql = vgSql & "                                                                on a.num_poliza=b.num_poliza"
    vgSql = vgSql & "                                                                and a.num_endoso=b.num_endoso+1"
    vgSql = vgSql & "                                           where a.num_endoso=(select max(num_endoso)"
    vgSql = vgSql & "                                                               From pp_tmae_poliza"
    vgSql = vgSql & "                                                               where num_poliza=a.num_poliza)"
    vgSql = vgSql & "                                             and a.cod_estpension = 99"
    vgSql = vgSql & "                                             and a.cod_derpen     = 99"
    vgSql = vgSql & "                                             and b.cod_cauendoso <> 13"
    vgSql = vgSql & "                                        ) b on a.num_poliza=b.num_poliza"
    vgSql = vgSql & "                                        and a.num_orden=b.num_orden"
    vgSql = vgSql & "      where fec_tercer=(select max(fec_tercer) from pp_tmae_certificado where num_poliza=a.num_poliza and num_orden=a.num_orden)"
    vgSql = vgSql & "        and ("
    vgSql = vgSql & "         b.cod_par<>30 or" '-- familiares no hijos
    vgSql = vgSql & "         (b.cod_par=30 and (months_between(TRUNC(sysdate),to_date(b.fec_nacben,'yyyymmdd'))/12)<18 and a.cod_tipo='SUP') or" '--hijos menores de edad
    vgSql = vgSql & "         (b.cod_par=30 and (months_between(TRUNC(sysdate),to_date(b.fec_nacben,'yyyymmdd'))/12)>=18 and" '--hijos mayores de edad"
    'vgSql = vgSql & "          a.cod_tipo='EST' and a.est_act=1 and a.ind_dni='S' and a.ind_dju='S' and a.ind_pes='S' and a.ind_bno='S')" '--certificado de estudios vigente
    vgSql = vgSql & "         ( ((select fec_dev from pp_tmae_poliza p where p.num_poliza=b.num_poliza and p.num_endoso=b.num_endoso) >= '20130801' and "
    vgSql = vgSql & "            a.cod_tipo='EST' and a.est_act=1 and a.ind_dni='S' and a.ind_dju='S' and a.ind_pes='S' and a.ind_bno='S') or (b.cod_sitinv<>'N' and a.cod_tipo='SUP'))) " '--certificado de estudios vigente  o situacion de invalides (materiagris 01/03/2018)
    vgSql = vgSql & "     )) a"
    vgSql = vgSql & " where fec_tercer <='" & vlFecPago & "' "
    vgSql = vgSql & " Union All"
    '-------------certificados cod_cauendoso = '13'-------------
    vgSql = vgSql & " select 0 sel, a.num_poliza as poliza, a.num_idenben as DNI, a.gls_nomben || ' ' || a.gls_patben || ' ' || gls_matben nombres , to_date(a.fec_tercer,'YYYYMMDD') fechaFin, a.cod_tipo"
    vgSql = vgSql & " From (select a.*, b.gls_nomben, b.gls_patben, b.gls_matben, b.num_idenben"
    vgSql = vgSql & "       from pp_tmae_certificado a Join (  select a.*"
    vgSql = vgSql & "                                           from pp_tmae_ben a  inner join pp_tmae_endoso b"
    vgSql = vgSql & "                                                                on a.num_poliza=b.num_poliza"
    vgSql = vgSql & "                                                                and a.num_endoso=b.num_endoso+1"
    vgSql = vgSql & "                                           where a.num_endoso=(select max(num_endoso)"
    vgSql = vgSql & "                                                               From pp_tmae_poliza"
    vgSql = vgSql & "                                                               where num_poliza=a.num_poliza)"
    vgSql = vgSql & "                                             and a.cod_estpension = 99"
    vgSql = vgSql & "                                             and a.cod_derpen     = 99"
    vgSql = vgSql & "                                             and b.cod_cauendoso = 13"
    vgSql = vgSql & "                                             AND b.fec_efecto <='" & vlFecPago & "' "
    vgSql = vgSql & "                                        ) b on a.num_poliza=b.num_poliza"
    vgSql = vgSql & "                                        and a.num_orden=b.num_orden"
    ' Materia Gris Jaime Rios 21/05/2018 - INICIO
    'vgSql = vgSql & "      where fec_tercer=(select max(fec_tercer) from pp_tmae_certificado where num_poliza=a.num_poliza and num_orden=a.num_orden)"
    vgSql = vgSql & "      where fec_tercer=(select max(fec_tercer) from pp_tmae_certificado "
    vgSql = vgSql & "                         Where num_poliza = a.num_poliza And Num_Orden = a.Num_Orden "
    vgSql = vgSql & "                           and num_endoso = (select max(c2.num_endoso) from pp_tmae_certificado c2 "
    vgSql = vgSql & "                                              where c2.num_poliza=a.num_poliza and c2.num_orden=a.num_orden)) "
    ' Materia Gris Jaime Rios 21/05/2018 - FIN
    vgSql = vgSql & "        and ("
    vgSql = vgSql & "         b.cod_par<>30 or" '-- familiares no hijos
    vgSql = vgSql & "         (b.cod_par=30 and (months_between(TRUNC(sysdate),to_date(b.fec_nacben,'yyyymmdd'))/12)<18 and a.cod_tipo='SUP') or" '--hijos menores de edad
    vgSql = vgSql & "         (b.cod_par=30 and (months_between(TRUNC(sysdate),to_date(b.fec_nacben,'yyyymmdd'))/12)>=18 and" '--hijos mayores de edad
    'vgSql = vgSql & "          a.cod_tipo='EST' and a.est_act=1 and a.ind_dni='S' and a.ind_dju='S' and a.ind_pes='S' and a.ind_bno='S')" '--certificado de estudios vigente
    vgSql = vgSql & "         ( ((select fec_dev from pp_tmae_poliza p where p.num_poliza=b.num_poliza and p.num_endoso=b.num_endoso) >= '20130801' and "
    vgSql = vgSql & "            a.cod_tipo='EST' and a.est_act=1 and a.ind_dni='S' and a.ind_dju='S' and a.ind_pes='S' and a.ind_bno='S') or (b.cod_sitinv<>'N' and a.cod_tipo='SUP'))) " '--certificado de estudios vigente  o situacion de invalides (materiagris 01/03/2018)
    vgSql = vgSql & "     )) a"
    vgSql = vgSql & " where fec_tercer <='" & vlFecPago & "' order by 6,2"
    '----------MateriaGris JaimeRios 01/02/2018----------'
    'vgSql = " select 0 sel, a.num_poliza as poliza, a.num_idenben as DNI, a.gls_nomben || ' ' || a.gls_patben || ' ' || gls_matben nombres ,"
    'vgSql = vgSql & " to_date(a.fec_tercer,'YYYYMMDD') fechaFin From"
    'vgSql = vgSql & " (select a.*, b.gls_nomben, b.gls_patben, b.gls_matben, b.num_idenben"
    'vgSql = vgSql & " from pp_tmae_certificado a"
    'vgSql = vgSql & " Join ( select * from pp_tmae_ben a"
    'vgSql = vgSql & "       where num_endoso=(select max(num_endoso) from pp_tmae_poliza where num_poliza=a.num_poliza)"
    'vgSql = vgSql & "       and a.cod_estpension=99 and a.cod_derpen=99 ) b on a.num_poliza=b.num_poliza and a.num_orden=b.num_orden"
    'vgSql = vgSql & " where fec_tercer=(select max(fec_tercer) from pp_tmae_certificado where num_poliza=a.num_poliza and num_orden=a.num_orden) ) a"
    'vgSql = vgSql & " where fec_tercer <='" & vlFecPago & "' order by 2"
    
    Dim Conteo As Integer
    
    Conteo = 0

    Set vgRegistro = vgConexionBD.Execute(vgSql)
  
    If Not vgRegistro.EOF Then
    
       Call Cargar_FlexGrid(MSFlexGrid1, 0, vgRegistro)
    
    
       'Do While Not vgRegistro.EOF
       '   If Lst_LiqSeleccion.ListCount = 1 Then
       '     Lst_LiqSeleccion.AddItem (" TODOS "), 0
       '   End If
       '   Lst_LiqSeleccion.AddItem (" " & Trim(vgRegistro!num_poliza) & " - DNI: " & Trim(vgRegistro!Num_IdenBen) & " - Nombres: " & Trim(vgRegistro!nombres) & " - Fin Cert." & Trim(vgRegistro!FechaFin))
             
       '   vgRegistro.MoveNext
       '   Conteo = Conteo + 1
        
       'Loop
       
    End If
    'boltodos = False
    'Call cargaReporteFin(vgRegistro)
    'lblConteo.Caption = "Existen " & CStr(Conteo) & " Certificados por Caducar."
    
    Set vgRegistro = Nothing
    Screen.MousePointer = vbDefault
    
    'flCargarListaConHabDes = True

Exit Function
Err_flCargarListaConHabDes:
    Screen.MousePointer = vbDefault
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Private Sub Command5_Click()

On Local Error GoTo errSub

Dim i, x As Integer
Dim numPol As String
Dim NumDoc As String
Dim CodTipo As String
Dim Ex As Integer

    Ex = 0
    For x = 1 To MSFlexGrid1.rows - 1
        If MSFlexGrid1.CellPicture = picChecked Then
            Ex = 1
        End If
    Next

    If Ex = 0 Then
        MsgBox "Debe elegir algun elemento de la lista."
        Exit Sub
    End If

    vgRes = MsgBox("¿ Está seguro que desea Procesar las polizas?", 4 + 32 + 256, "Actualización de Certificados")
    If vgRes <> 6 Then
        Command5.SetFocus
        Screen.MousePointer = 0
        Exit Sub
    End If


'If MSFlexGrid2.Rows >= 1 Then
    For x = 1 To MSFlexGrid1.rows - 1
        numPol = MSFlexGrid1.TextMatrix(x, 2)
        NumDoc = MSFlexGrid1.TextMatrix(x, 3)
        CodTipo = MSFlexGrid1.TextMatrix(x, 6)
        MSFlexGrid1.Col = 1
        MSFlexGrid1.row = x
        
        If MSFlexGrid1.CellPicture = picChecked Then
            Ex = 1
            If CodTipo <> "EST" Then
            vgSql = " INSERT INTO pp_tmae_certificado (NUM_POLIZA, NUM_ENDOSO, NUM_ORDEN, FEC_INICER, COD_TIPO, FEC_TERCER, COD_FRECUENCIA, GLS_NOMINSTITUCION, FEC_RECCIA, FEC_INGCIA, FEC_EFECTO, COD_USUARIOCREA, FEC_CREA, HOR_CREA, COD_INDRELIQUIDAR,"
                vgSql = vgSql & " est_act,ind_dni,ind_dju,ind_pes,ind_bno) " 'MateriaGris Jaime Rios
                vgSql = vgSql & " select A.NUM_POLIZA, B.NUM_ENDOSO, A.NUM_ORDEN,"
                vgSql = vgSql & " to_char(to_date(A.fec_tercer,'YYYYMMDD') + 1,'YYYYMMDD'),"
                vgSql = vgSql & " COD_TIPO,"
                vgSql = vgSql & " to_char(add_months( to_date(A.fec_tercer,'YYYYMMDD') + 1,12),'YYYYMMDD'),"
                vgSql = vgSql & " COD_FRECUENCIA, GLS_NOMINSTITUCION, FEC_RECCIA, FEC_INGCIA,"
                vgSql = vgSql & " A.FEC_EFECTO,"
                vgSql = vgSql & " '" & vgUsuario & "', '" & Format(Date, "yyyymmdd") & "', '" & Format(Time, "hhmmss") & "', COD_INDRELIQUIDAR,"
                vgSql = vgSql & " est_act,ind_dni,ind_dju,ind_pes,ind_bno " 'MateriaGris Jaime Rios
                vgSql = vgSql & " from pp_tmae_certificado A left JOIN"
                vgSql = vgSql & " PP_TMAE_BEN B ON A.NUM_POLIZA=B.NUM_POLIZA AND A.NUM_ORDEN=B.NUM_ORDEN"
                vgSql = vgSql & " WHERE A.NUM_POLIZA ='" & numPol & "' and B.NUM_IDENBEN='" & NumDoc & "'"
                'MateriaGris Jaime Rios 21/05/2018 - INICIO
                'vgSql = vgSql & " AND FEC_TERCER = (SELECT MAX(FEC_TERCER) FROM pp_tmae_certificado WHERE NUM_POLIZA=A.NUM_POLIZA and num_orden=a.num_orden)"
                vgSql = vgSql & " AND FEC_TERCER = (SELECT MAX(FEC_TERCER) FROM pp_tmae_certificado "
                vgSql = vgSql & "                    WHERE NUM_POLIZA=A.NUM_POLIZA and num_orden=a.num_orden "
                vgSql = vgSql & "                      AND NUM_ENDOSO = (SELECT max(c2.num_endoso) FROM pp_tmae_certificado c2 "
                vgSql = vgSql & "                                         WHERE c2.num_poliza=a.num_poliza AND c2.num_orden=a.num_orden and c2.cod_tipo=a.cod_tipo)) "
                'MateriaGris Jaime Rios 21/05/2018 - FIN
                vgSql = vgSql & " AND B.NUM_ENDOSO = (SELECT MAX(NUM_ENDOSO) FROM PP_TMAE_BEN WHERE NUM_POLIZA=B.NUM_POLIZA)"
                vgSql = vgSql & " AND COD_TIPO='" & CodTipo & "'"
                vgSql = vgSql & " AND B.COD_ESTPENSION=99"
            Else
                vgSql = " INSERT INTO pp_tmae_certificado (NUM_POLIZA, NUM_ENDOSO, NUM_ORDEN, FEC_INICER, COD_TIPO, FEC_TERCER, COD_FRECUENCIA, GLS_NOMINSTITUCION, FEC_RECCIA, FEC_INGCIA, FEC_EFECTO, COD_USUARIOCREA, FEC_CREA, HOR_CREA, COD_INDRELIQUIDAR,"
                vgSql = vgSql & " est_act,ind_dni,ind_dju,ind_pes,ind_bno) " 'MateriaGris Jaime Rios
                vgSql = vgSql & " select A.NUM_POLIZA, A.NUM_ENDOSO, A.NUM_ORDEN,"
                'vgSql = vgSql & " to_char(to_date(A.fec_tercer,'YYYYMMDD') + 1,'YYYYMMDD'),"
                vgSql = vgSql & " to_char(sysdate,'YYYYMMDD') fec_ini," 'MateriaGris Jaime Rios
                vgSql = vgSql & " COD_TIPO,"
                'vgSql = vgSql & " to_char(add_months( to_date(A.fec_tercer,'YYYYMMDD') + 1,12),'YYYYMMDD'),"
                vgSql = vgSql & " to_char(add_months(sysdate + 1,6),'YYYYMMDD') fec_fin," 'MateriaGris Jaime Rios
                vgSql = vgSql & " COD_FRECUENCIA, GLS_NOMINSTITUCION, FEC_RECCIA, FEC_INGCIA,"
                vgSql = vgSql & " A.FEC_EFECTO,"
                vgSql = vgSql & " '" & vgUsuario & "', '" & Format(Date, "yyyymmdd") & "', '" & Format(Time, "hhmmss") & "', COD_INDRELIQUIDAR,"
                vgSql = vgSql & " est_act,ind_dni,ind_dju,ind_pes,ind_bno " 'MateriaGris Jaime Rios
                vgSql = vgSql & " from pp_tmae_certificado A left JOIN"
                vgSql = vgSql & " PP_TMAE_BEN B ON A.NUM_POLIZA=B.NUM_POLIZA AND A.NUM_ORDEN=B.NUM_ORDEN"
                vgSql = vgSql & " WHERE A.NUM_POLIZA ='" & numPol & "' and B.NUM_IDENBEN='" & NumDoc & "'"
                'MateriaGris Jaime Rios 21/05/2018 - INICIO
                'vgSql = vgSql & " AND FEC_TERCER = (SELECT MAX(FEC_TERCER) FROM pp_tmae_certificado WHERE NUM_POLIZA=A.NUM_POLIZA and num_orden=a.num_orden)"
                vgSql = vgSql & " AND FEC_TERCER = (SELECT MAX(FEC_TERCER) FROM pp_tmae_certificado "
                vgSql = vgSql & "                    WHERE NUM_POLIZA=A.NUM_POLIZA and num_orden=a.num_orden "
                vgSql = vgSql & "                      AND NUM_ENDOSO = (SELECT max(c2.num_endoso) FROM pp_tmae_certificado c2 "
                vgSql = vgSql & "                                         WHERE c2.num_poliza=a.num_poliza AND c2.num_orden=a.num_orden)) "
                'MateriaGris Jaime Rios 21/05/2018 - FIN
                vgSql = vgSql & " AND B.NUM_ENDOSO = (SELECT MAX(NUM_ENDOSO) FROM PP_TMAE_BEN WHERE NUM_POLIZA=B.NUM_POLIZA)"
                vgSql = vgSql & " AND COD_TIPO='" & CodTipo & "'"
                vgSql = vgSql & " AND B.COD_ESTPENSION=99"
            End If
            'MsgBox (numPol & " - " & NumDoc)
            
            'vgSql = " INSERT INTO pp_tmae_certificado (NUM_POLIZA, NUM_ENDOSO, NUM_ORDEN, FEC_INICER, COD_TIPO, FEC_TERCER, COD_FRECUENCIA, GLS_NOMINSTITUCION, FEC_RECCIA, FEC_INGCIA, FEC_EFECTO, COD_USUARIOCREA, FEC_CREA, HOR_CREA, COD_INDRELIQUIDAR)"
            'vgSql = vgSql & " select A.NUM_POLIZA, B.NUM_ENDOSO, A.NUM_ORDEN,"
            'vgSql = vgSql & " to_char(to_date(A.fec_tercer,'YYYYMMDD') + 1,'YYYYMMDD'),"
            'vgSql = vgSql & " COD_TIPO,"
            'vgSql = vgSql & " to_char(add_months( to_date(A.fec_tercer,'YYYYMMDD') + 1,12),'YYYYMMDD'),"
            'vgSql = vgSql & " COD_FRECUENCIA, GLS_NOMINSTITUCION, FEC_RECCIA, FEC_INGCIA,"
            'vgSql = vgSql & " A.FEC_EFECTO,"
            'vgSql = vgSql & " '" & vgUsuario & "', '" & Format(Date, "yyyymmdd") & "', '" & Format(Time, "hhmmss") & "', COD_INDRELIQUIDAR"
            'vgSql = vgSql & " from pp_tmae_certificado A left JOIN"
            
            
            'MVG 20180111
            
            'vgSql = vgSql & " PP_TMAE_BEN B ON A.NUM_POLIZA=B.NUM_POLIZA AND A.NUM_ORDEN=B.NUM_ORDEN and  A.num_ENDOSO=B.num_ENDOSO "
            
            ''vgSql = vgSql & " PP_TMAE_BEN B ON A.NUM_POLIZA=B.NUM_POLIZA AND A.NUM_ORDEN=B.NUM_ORDEN AND A.NUM_ENDOSO=B.NUM_ENDOSO"
            'vgSql = vgSql & " WHERE A.NUM_POLIZA ='" & numPol & "' and B.NUM_IDENBEN='" & NumDoc & "'"
            ''MVG 20180111
            ''vgSql = vgSql & " AND FEC_TERCER = (SELECT MAX(FEC_TERCER) FROM pp_tmae_certificado WHERE NUM_POLIZA=A.NUM_POLIZA and num_orden=a.num_orden)"
            'vgSql = vgSql & " AND FEC_TERCER = (SELECT MAX(FEC_TERCER) FROM pp_tmae_certificado WHERE NUM_POLIZA=A.NUM_POLIZA and num_ENDOSO=a.num_ENDOSO and num_orden=a.num_orden  AND COD_TIPO='SUP')"
            'vgSql = vgSql & " AND B.NUM_ENDOSO = (SELECT MAX(NUM_ENDOSO) FROM PP_TMAE_BEN WHERE NUM_POLIZA=B.NUM_POLIZA)"
            'vgSql = vgSql & " AND COD_TIPO='SUP'"
            'vgSql = vgSql & " AND B.COD_ESTPENSION=99"

            vgConexionBD.Execute (vgSql)
       
        End If
    Next
    
    MsgBox "Se procesaron Correctamente."
    
    Exit Sub
'End If
'Error
errSub:
MsgBox Err.Description & " " & numPol, vbCritical
MSFlexGrid1.Redraw = True
End Sub

Private Sub Command6_Click()

Dim i, x As Integer
Dim NumDoc1 As String
Dim NumDoc2 As String

If MSFlexGrid2.rows >= 1 Then
    For i = 1 To MSFlexGrid2.rows - 1
        NumDoc1 = MSFlexGrid2.TextMatrix(i, 1)
        For x = 1 To MSFlexGrid1.rows - 1
            NumDoc2 = MSFlexGrid1.TextMatrix(x, 3)
            If NumDoc1 = NumDoc2 Then

                MSFlexGrid1.row = x
                MSFlexGrid1.Col = 1
                Set MSFlexGrid1.CellPicture = picChecked.Picture  ' Set the default checkbox picture to the empty box
                'MSFlexGrid1.TextMatrix(x, 1) = "1"
            End If
        Next
    Next
End If




End Sub

Private Sub Lst_LiqSeleccion_ItemCheck(Item As Integer)
Dim i As Integer
    If Item = 0 Then
        If Lst_LiqSeleccion.Selected(0) Then
            'boltodos = True
            For i = 1 To Lst_LiqSeleccion.ListCount - 1
                Lst_LiqSeleccion.Selected(i) = True
            Next
        Else
            'boltodos = False
            For i = 1 To Lst_LiqSeleccion.ListCount - 1
                Lst_LiqSeleccion.Selected(i) = False
            Next
        End If
    Else
        Lst_LiqSeleccion.Selected(0) = False
    End If
End Sub
Private Sub Form_Load()
Dim cadena As String

On Error Resume Next

 Me.Top = 0
 Me.Left = 0
 'Command2.Enabled = False
 
' Set rsLoad = New ADODB.Recordset
' rsLoad.CursorLocation = adUseClient
' cadena = "select RUTA from PT_TMP_SUPV" & vgUsuario
' rsLoad.Open cadena, vgConexionBD, adOpenStatic, adLockReadOnly
' If Not rsLoad.EOF Then
'    txtRuta.Text = Trim(rsLoad!ruta)
'    txtRuta.Tag = Trim(rsLoad!ruta)
' End If
 
 
End Sub

  'Sub que carga los registros en la Grilla
''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Cargar_FlexGrid(FlexGrid As Object, _
                            NumeroCampo As Integer, _
                            objRst As ADODB.Recordset)
                              
On Local Error GoTo errSub
                              
Dim c As Integer
Dim fila As Integer
Dim AnchoCol() As Single
Dim TempAnchoCol As Single
Dim valniv As Integer

    With FlexGrid
        ' deshabilita el repintado para que sea mas rápido
        .Redraw = False
        'Cantidad de filas y columnas
        .rows = 1
        .Cols = objRst.Fields.Count + 1
    End With
      
    'Redimensiona el Array a la cantidad de campos del recordset
    ReDim AnchoCol(0 To objRst.Fields.Count)
      
    'Recorre las columnas
    For c = 1 To objRst.Fields.Count
        'Añade el título del campo al encabezado de columna
        FlexGrid.TextMatrix(0, c) = objRst.Fields(c - 1).Name
          
        'Guarda el ancho del campo en la matriz
        AnchoCol(c) = TextWidth(objRst.Fields(c - 1).Name)
    Next c
    fila = 1
    
    'Recorre todos los registros del recordset
    Do While Not objRst.EOF
        ' Añade una nueva fila
        FlexGrid.rows = FlexGrid.rows + 1
        For c = 1 To objRst.Fields.Count
              
            'Si el valor no es nulo
            If Not IsNull(objRst.Fields(1).Value) Then
               ' si la columna es el campo de tipo CheckBox ...
              
                If c = 1 Then
                         With FlexGrid
                             .row = fila
                             .Col = c
                             .CellPictureAlignment = 4 ' Align the checkbox
                             .CellBackColor = &H80000018
                             If objRst.Fields(c).Value = "1" Then
                                Set .CellPicture = picChecked.Picture  ' Set the default checkbox picture to the empty box
                             Else
                                Set .CellPicture = picUnchecked.Picture  ' Set the default checkbox picture to the empty box
                             End If
                             
                             '.TextMatrix(Fila, c) = Fila
                             TempAnchoCol = TextWidth(objRst.Fields(c - 1).Value)
                        End With
                Else
                        'Agrega el registro en la fila y columna específica
                        FlexGrid.TextMatrix(fila, c) = objRst.Fields(c - 1).Value
                             '.TextMatrix(Fila, c) = Fila
                             TempAnchoCol = TextWidth(objRst.Fields(c - 1).Value)
                             'Materia Gris Jaime Rios 22/03/2018
                             If c = 6 And objRst.Fields(c - 1).Value = "EST" Then
                                FlexGrid.CellBackColor = vbRed
                             End If
                             'Materia Gris Jaime Rios 22/03/2018
                End If
               
                    ' Almacena el ancho
                TempAnchoCol = TextWidth(objRst.Fields(c - 1).Value)
               
                             
               If AnchoCol(c) < TempAnchoCol Then
                  AnchoCol(c) = TempAnchoCol ' nuevo ancho
               End If
            End If
        Next
        ' Siguiente registro
        objRst.MoveNext
        fila = fila + 1 'Incrementa la posición de la fila actual
    Loop
    lblConteo.Caption = "Existen " & CStr(fila - 1) & " Certificados por Caducar."
    ' Establece los ancho máximos de columna
    For c = 0 To FlexGrid.Cols - 1
        FlexGrid.ColWidth(c) = AnchoCol(c) + 250
    Next
    ' vuelve a habilitar el redraw
    FlexGrid.Redraw = True
Exit Sub
  
'Error
errSub:
MsgBox Err.Description, vbCritical
FlexGrid.Redraw = True
End Sub

Private Sub Cargar_FlexGridExcel(FlexGrid As Object, objRst As ADODB.Recordset)
                              
On Local Error GoTo errSub
                              
Dim c As Integer
Dim fila As Integer
Dim AnchoCol() As Single
Dim TempAnchoCol As Single
Dim valniv As Integer

    With FlexGrid
        ' deshabilita el repintado para que sea mas rápido
        .Redraw = False
        'Cantidad de filas y columnas
        .rows = 1
        .Cols = objRst.Fields.Count
    End With
      
    'Redimensiona el Array a la cantidad de campos del recordset
    ReDim AnchoCol(0 To objRst.Fields.Count - 1)
      
    'Recorre las columnas
    For c = 0 To objRst.Fields.Count - 1
        'Añade el título del campo al encabezado de columna
        FlexGrid.TextMatrix(0, c) = objRst.Fields(c).Name
          
        'Guarda el ancho del campo en la matriz
        AnchoCol(c) = TextWidth(objRst.Fields(c).Name)
    Next c
    fila = 1
      
    'Recorre todos los registros del recordset
    Do While Not objRst.EOF
        ' Añade una nueva fila
        FlexGrid.rows = FlexGrid.rows + 1
        For c = 0 To objRst.Fields.Count - 1
              
            'Si el valor no es nulo
            If Not IsNull(objRst.Fields(c).Value) Then
               ' si la columna es el campo de tipo CheckBox ...
               If c = 0 Then
                    With FlexGrid
                         .row = fila
                         .Col = c
                         .TextMatrix(fila, c) = objRst.Fields(c).Value
                         '.CellPictureAlignment = 4 ' Align the checkbox
                         '.CellBackColor = &H80000018
                         'If objRst.Fields(c).Value = "1" Then
                         '   Set .CellPicture = picChecked.Picture  ' Set the default checkbox picture to the empty box
                         'Else
                         '   Set .CellPicture = picUnchecked.Picture  ' Set the default checkbox picture to the empty box
                         'End If
                         
                         '.TextMatrix(Fila, c) = Fila
                         TempAnchoCol = TextWidth(objRst.Fields(c).Value)
                         
                    End With
                      
               Else
                   'Agrega el registro en la fila y columna específica
                   FlexGrid.TextMatrix(fila, c) = objRst.Fields(c).Value

                    ' Almacena el ancho
                   TempAnchoCol = TextWidth(objRst.Fields(c).Value)
               End If
                             
               If AnchoCol(c) < TempAnchoCol Then
                  AnchoCol(c) = TempAnchoCol ' nuevo ancho
               End If
            End If
        Next
        ' Siguiente registro
        objRst.MoveNext
        'Fila = Fila + 1 'Incrementa la posición de la fila actual
    Loop
    'lblConteo.Caption = "Existen " & CStr(Fila) & " Certificados por Caducar."
    ' Establece los ancho máximos de columna
    For c = 0 To FlexGrid.Cols - 1
        FlexGrid.ColWidth(c) = AnchoCol(c) + 250
    Next
    ' vuelve a habilitar el redraw
    FlexGrid.Redraw = True
Exit Sub
  
'Error
errSub:
MsgBox Err.Description, vbCritical
FlexGrid.Redraw = True
End Sub



Private Sub MSFlexGrid1_Click()
Dim oldx, oldy, cell2text As String, strTextCheck As String
Dim Sis As String
Dim Niv As String
Dim strChecked As String
' Check or uncheck the grid checkbox
With MSFlexGrid1

    oldx = .Col
    oldy = .row
    
    If MSFlexGrid1.Col = 1 Then
            If MSFlexGrid1.CellPicture = picChecked Then
                Set MSFlexGrid1.CellPicture = picUnchecked
                .Col = .Col - 1 ' I use data that is in column #1, usually an Index or ID #
                strTextCheck = .Text
                ' When you de-select a CheckBox, we need to strip out the #
                strChecked = Replace(strChecked, strTextCheck & ",", "")
                ' Don't forget to strip off the trailing , before passing the string
                Debug.Print strChecked
                'MSFlexGrid1.TextMatrix(.Row, 1) = "0"
            Else
                Set MSFlexGrid1.CellPicture = picChecked
                .Col = .Col - 1
                strTextCheck = .Text
                strChecked = strChecked & strTextCheck & ","
                Debug.Print strChecked
                'MSFlexGrid1.TextMatrix(.Row, 1) = "1"
            End If
            
    End If
        
    '.Col = oldx
    '.Row = oldy
    
   
    
End With
End Sub

Private Sub txtDniBus_Change()
Dim i, j As Integer
Dim filSel As Integer
'Damos al FlexiGrid el color de fondo por defecto
MSFlexGrid1.BackColor = &H80000005
'‘Si la caja de texto está vacía eliminamos el contenido del label y salimos
If txtDniBus.Text = "" Then
    'Label1.Caption = ""
    Exit Sub
End If

i = 1
j = 2
'‘Recorremos todas la filas del FlexiGrid columna a columna
For i = 1 To MSFlexGrid1.rows - 1
    For j = 1 To MSFlexGrid1.Cols - 1
        '‘comprobamos si coincide el contenido del Text1 con la celda
        If LCase(txtDniBus.Text) = MSFlexGrid1.TextMatrix(i, 3) Then
            '‘En caso afirmativo mostramos su contenido en un Label1
            'Label1.Caption = MSFlexGrid1.TextMatrix(i, 3)
            '‘Seleccionamos la celda para darle color de fondo
            filSel = i
            MSFlexGrid1.row = i
            MSFlexGrid1.Col = j - 1
            MSFlexGrid1.ColSel = 5
            MSFlexGrid1.BackColorSel = QBColor(1)
            '‘Damos unos valores a I y J para que salga de nol dos For y no continue buscando. Si no hiciéramos esto el label mostraría la última celda que coincida con el contenido del Text1
            i = MSFlexGrid1.rows + 1
            j = MSFlexGrid1.Cols + 1
        End If
    Next j
Next i
If filSel <> 0 Then
    MSFlexGrid1.TopRow = filSel
End If


End Sub

Private Sub chkTodos_Click()
Dim i As Integer
Dim oldx, oldy, cell2text As String, strTextCheck As String
Dim Sis As String
Dim Niv As String
Dim strChecked As String

    If chkTodos.Value = 1 Then
        For i = 1 To MSFlexGrid1.rows - 1
            With MSFlexGrid1
                .row = i
                If MSFlexGrid1.CellPicture = picUnchecked Then
                    Set MSFlexGrid1.CellPicture = picChecked
                    '.Col = .Col - 1 ' I use data that is in column #1, usually an Index or ID #
                    strTextCheck = .Text
                    ' When you de-select a CheckBox, we need to strip out the #
                    strChecked = Replace(strChecked, strTextCheck & ",", "")
                    ' Don't forget to strip off the trailing , before passing the string
                    Debug.Print strChecked
                    'MSFlexGrid1.TextMatrix(.Row, 1) = "0"
                Else
                    Set MSFlexGrid1.CellPicture = picUnchecked
                    '.Col = .Col - 1
                    strTextCheck = .Text
                    strChecked = strChecked & strTextCheck & ","
                    Debug.Print strChecked
                    'MSFlexGrid1.TextMatrix(.Row, 1) = "1"
                End If
            End With
        Next i
    Else
        For i = 1 To MSFlexGrid1.rows - 1
            With MSFlexGrid1
                .row = i
                If MSFlexGrid1.CellPicture = picUnchecked Then
                    Set MSFlexGrid1.CellPicture = picChecked
                    '.Col = .Col - 1 ' I use data that is in column #1, usually an Index or ID #
                    strTextCheck = .Text
                    ' When you de-select a CheckBox, we need to strip out the #
                    strChecked = Replace(strChecked, strTextCheck & ",", "")
                    ' Don't forget to strip off the trailing , before passing the string
                    Debug.Print strChecked
                    'MSFlexGrid1.TextMatrix(.Row, 1) = "0"
                Else
                    Set MSFlexGrid1.CellPicture = picUnchecked
                    '.Col = .Col - 1
                    strTextCheck = .Text
                    strChecked = strChecked & strTextCheck & ","
                    Debug.Print strChecked
                    'MSFlexGrid1.TextMatrix(.Row, 1) = "1"
                End If
            End With
        Next i
    End If
End Sub

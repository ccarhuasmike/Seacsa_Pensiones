VERSION 5.00
Begin VB.Form Frm_ExcelAnexo2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exportacion Anexo 2"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.CommandButton cmdExpExc 
         Caption         =   "Exportar"
         Height          =   615
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtPoliza 
         Height          =   285
         Left            =   2160
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Digite el numero de Poliza"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Frm_ExcelAnexo2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExpExc_Click()

    Dim sPathDB        As String
    Dim Consulta    As String
  
    ' -- Path de la base de datos
    sPathDB = "C:\Archivos de programa\Microsoft Visual Studio\VB98\BIBLIO.MDB"
  
    ' -- Cadena Sql
    'Consulta = "Select * From Authors"
  
  
    'Consulta = "select distinct a.num_poliza,  c.cod_cuspp, "
    'Consulta = Consulta & " e.gls_elemento as AFP, c.cod_tipoidenafi,b.num_idenben, b.gls_nomben || ' ' || gls_nomsegben || ' ' || gls_patben || ' ' || gls_matben as nombres,"
    'Consulta = Consulta & " g.gls_elemento as TIPOPEN, h.gls_elemento as modalidad,"
    'Consulta = Consulta & " num_mesdif / 12 AS añosdif, num_mesgar / 12 as añosgar, a.cod_moneda,"
    'Consulta = Consulta & " a.mto_valmoneda as TCTrasp, k.mto_pritotal as CIC, mto_prirec as PrimaTras,"
'    Consulta = Consulta & " to_date(b.fec_inipagopen,'YYYYMMDD') as fec_inipagopen, to_date(k.fec_traspaso,'YYYYMMDD') as fec_traspaso,"
'    Consulta = Consulta & " c.num_idencor, l.gls_patcor || ' ' || l.gls_matcor || ' ' || l.gls_nomcor as nomasesor, round(c.prc_corcomreal,0) Comision, case when length(c.prc_corcomreal)=1 then 0 else 0.40 end Benef, lc.mto_comision, substr(fec_pago,1,6) as fec_pago,"
'    Consulta = Consulta & " l.num_idenjefe, s.gls_patcor || ' ' || s.gls_matcor || ' ' ||  s.gls_nomcor as nomsupervusor"
'    Consulta = Consulta & " from pp_tmae_poliza a"
'    Consulta = Consulta & " join pp_tmae_ben b on a.num_poliza=b.num_poliza and a.num_endoso=b.num_endoso"
'    Consulta = Consulta & " join pd_tmae_poliza c on a.num_poliza=c.num_poliza"
'    Consulta = Consulta & " join pt_tmae_cotizacion d on c.num_cot=d.num_cot"
'    Consulta = Consulta & " join pt_tmae_detcotizacion j on d.num_cot= j.num_cot"
'    Consulta = Consulta & " join ma_tpar_tabcod e on a.cod_afp=e.cod_elemento and e.cod_tabla='AF'"
'    Consulta = Consulta & " join ma_tpar_tabcod f on b.cod_par=f.cod_elemento and f.cod_tabla='PA'"
'    Consulta = Consulta & " join ma_tpar_tabcod g on a.cod_tippension=g.cod_elemento and g.cod_tabla='TP'"
'    Consulta = Consulta & " join ma_tpar_tabcod h on a.cod_modalidad=h.cod_elemento and h.cod_tabla='AL'"
'    '--left join (select a.num_poliza, num_idenreceptor, max(fec_pago), cod_par from pp_tmae_liqpagopendef a
'    '--join pp_tmae_ben b on b.num_idenben=a.num_idenreceptor group by a.num_poliza, num_idenreceptor, cod_par ) i on b.num_idenben=i.num_idenreceptor
'    Consulta = Consulta & " join pd_tmae_polprirec k on a.num_poliza=k.num_poliza"
'    Consulta = Consulta & " join pt_tmae_corredor l on c.num_idencor=l.num_idencor"
'    Consulta = Consulta & " left join pr_tmae_calben1 cb1 on b.num_poliza=cb1.num_poliza and b.num_orden=cb1.num_orden"
'    Consulta = Consulta & " left join pr_tmae_calben2 cb2 on b.num_poliza=cb2.num_poliza and b.num_orden=cb2.num_orden"
'    Consulta = Consulta & " join pc_tmae_liqcomision lc on a.num_poliza=lc.num_poliza and c.num_idencor=lc.num_idencor"
'    Consulta = Consulta & " join pt_tmae_corredor s on l.num_idenjefe=s.num_idencor"
'    Consulta = Consulta & " where a.num_endoso=(select max(num_endoso) from pp_tmae_poliza where num_poliza=a.num_poliza)"
'    '--and a.num_poliza >=  1
'    '--and substr(k.fec_traspaso,1,6) >='201510' and substr(k.fec_traspaso,1,6) <='201510'
'    Consulta = Consulta & " and a.num_poliza>=" & Trim(txtPoliza.Text) & ""
'    Consulta = Consulta & " and b.cod_par=99"
'    Consulta = Consulta & " order by 1,2"
  
  
  
    Consulta = " select distinct a.num_poliza,  c.cod_cuspp,  e.gls_elemento as AFP, c.cod_tipoidenafi,b.num_idenben, b.gls_nomben || ' ' || gls_nomsegben || ' ' || gls_patben || ' ' || gls_matben as nombres, g.gls_elemento as TIPOPEN,"
    Consulta = Consulta & " h.gls_elemento as modalidad, num_mesdif / 12 AS añosdif, num_mesgar / 12 as añosgar, a.cod_moneda, a.mto_valmoneda as TCTrasp, k.mto_pritotal as CIC, mto_prirec as PrimaTras, to_date(b.fec_inipagopen,'YYYYMMDD') as fec_inipagopen,"
    Consulta = Consulta & " to_date(k.fec_traspaso,'YYYYMMDD') as fec_traspaso, c.num_idencor, l.gls_patcor || ' ' || l.gls_matcor || ' ' || l.gls_nomcor as nomasesor,"
    Consulta = Consulta & " round(c.prc_corcomreal,0) ComReal,round (mto_prirec * round((c.prc_corcomreal/100),2),2) as MtoComReal,"
    Consulta = Consulta & " case when length(c.prc_corcomreal)=1 then 0 else 0.40 end Benef,"
    Consulta = Consulta & " round(mto_prirec * case when length(c.prc_corcomreal)=1 then 0 else 0.004 end,2) as MtoComBenef,"
    Consulta = Consulta & " lc.mto_comision as ComTotal,"
    Consulta = Consulta & " substr(fec_pago,1,6) as fec_pago, l.num_idenjefe, s.gls_patcor || ' ' || s.gls_matcor || ' ' ||  s.gls_nomcor as nomsupervusor"
    Consulta = Consulta & " from pp_tmae_poliza a join pp_tmae_ben b on a.num_poliza=b.num_poliza and a.num_endoso=b.num_endoso join pd_tmae_poliza c on a.num_poliza=c.num_poliza join pt_tmae_cotizacion d on c.num_cot=d.num_cot"
    Consulta = Consulta & " join pt_tmae_detcotizacion j on d.num_cot= j.num_cot join ma_tpar_tabcod e on a.cod_afp=e.cod_elemento and e.cod_tabla='AF' join ma_tpar_tabcod f on b.cod_par=f.cod_elemento and f.cod_tabla='PA'"
    Consulta = Consulta & " join ma_tpar_tabcod g on a.cod_tippension=g.cod_elemento and g.cod_tabla='TP' join ma_tpar_tabcod h on a.cod_modalidad=h.cod_elemento and h.cod_tabla='AL' join pd_tmae_polprirec k on a.num_poliza=k.num_poliza"
    Consulta = Consulta & " join pt_tmae_corredor l on c.num_idencor=l.num_idencor left join pr_tmae_calben1 cb1 on b.num_poliza=cb1.num_poliza and b.num_orden=cb1.num_orden"
    Consulta = Consulta & " left join pr_tmae_calben2 cb2 on b.num_poliza=cb2.num_poliza and b.num_orden=cb2.num_orden"
    Consulta = Consulta & " join pc_tmae_liqcomision lc on a.num_poliza=lc.num_poliza and c.num_idencor=lc.num_idencor"
    Consulta = Consulta & " join pt_tmae_corredor s on l.num_idenjefe=s.num_idencor"
    Consulta = Consulta & " where a.num_endoso=(select max(num_endoso) from pp_tmae_poliza"
    Consulta = Consulta & " where num_poliza=a.num_poliza) and a.num_poliza>=" & Trim(txtPoliza.Text) & " and b.cod_par=99 order by 1,2"
  
  
  
    ' -- Enviar el Path de la base de datos y la consulta sql
    If Exportar_ADO_Excel(sPathDB, Consulta, "c:\libro.xLS") Then
       MsgBox "Ok..Exportados con Exito.", vbInformation
    End If
End Sub

' ------------------------------------------------------------------------------------
' \\ -- Función para exportar el recordset ADO a una hoja de Excel
' ------------------------------------------------------------------------------------
Private Function Exportar_ADO_Excel(sPathDB As String, Sql As String, sOutputPathXLS As String) As Boolean
      
    On Error GoTo errSub
      
    Dim cn          As New ADODB.Connection
    Dim rec         As New ADODB.Recordset
    Dim Excel       As Object
    Dim Libro       As Object
    Dim Hoja        As Object
    Dim arrData     As Variant
    Dim iRec        As Long
    Dim iCol        As Integer
    Dim iRow        As Integer
      
    Me.Enabled = False
      
'   ' -- Abrir la base
'    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sPathDB & ";"
'
'    ' -- Abrir el Recordset pasándole la cadena sql
'    rec.Open Sql, cn
    
    Set vgRs2 = vgConexionBD.Execute(Sql)
    If Not vgRs2.EOF Then
         ' -- Crear los objetos para utilizar el Excel
          Set Excel = CreateObject("Excel.Application")
          Set Libro = Excel.Workbooks.Add
            
          ' -- Hacer referencia a la hoja
          Set Hoja = Libro.Worksheets(1)
            
          Excel.Visible = True: Excel.UserControl = True
          iCol = vgRs2.Fields.Count
          For iCol = 1 To vgRs2.Fields.Count
              Hoja.Cells(1, iCol).Value = vgRs2.Fields(iCol - 1).Name
          Next
            
          If Val(Mid(Excel.Version, 1, InStr(1, Excel.Version, ".") - 1)) > 8 Then
              Hoja.Cells(2, 1).CopyFromRecordset vgRs2
          Else
        
              arrData = vgRs2.GetRows
        
              iRec = UBound(arrData, 2) + 1
                
              For iCol = 0 To vgRs2.Fields.Count - 1
                  For iRow = 0 To iRec - 1
        
                      If IsDate(arrData(iCol, iRow)) Then
                          arrData(iCol, iRow) = Format(arrData(iCol, iRow))
        
                      ElseIf IsArray(arrData(iCol, iRow)) Then
                          arrData(iCol, iRow) = "Array Field"
                      End If
                  Next iRow
              Next iCol
                    
              ' -- Traspasa los datos a la hoja de Excel
              Hoja.Cells(2, 1).Resize(iRec, vgRs2.Fields.Count).Value = GetData(arrData)
          End If
        
          Excel.Selection.CurrentRegion.Columns.AutoFit
          Excel.Selection.CurrentRegion.Rows.AutoFit
        
          ' -- Cierra el recordset y la base de datos y los objetos ADO
          vgRs2.Close
          'cn.Close
            
          Set vgRs2 = Nothing
          'Set cn = Nothing
          ' -- guardar el libro
          Libro.SaveAs sOutputPathXLS
          Libro.Close
          ' -- Elimina las referencias Xls
          Set Hoja = Nothing
          Set Libro = Nothing
          Excel.Quit
          Set Excel = Nothing
            
          Exportar_ADO_Excel = True
          Me.Enabled = True
    End If
    
    Exit Function
errSub:
    MsgBox Err.Description, vbCritical, "Error"
    'Exportar_ADO_Excel = False
    Me.Enabled = True
End Function
  
Private Function GetData(vValue As Variant) As Variant
    Dim x As Long, y As Long, xMax As Long, yMax As Long, T As Variant
      
    xMax = UBound(vValue, 2): yMax = UBound(vValue, 1)
      
    ReDim T(xMax, yMax)
    For x = 0 To xMax
        For y = 0 To yMax
            T(x, y) = vValue(y, x)
        Next y
    Next x
      
    GetData = T
End Function

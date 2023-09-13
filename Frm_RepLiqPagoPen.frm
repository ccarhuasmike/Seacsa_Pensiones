VERSION 5.00
Begin VB.Form Frm_RepLiqPagoPen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de pagos pendientes"
   ClientHeight    =   3150
   ClientLeft      =   3990
   ClientTop       =   4440
   ClientWidth     =   6060
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   6060
   Begin VB.ListBox Lst_TPi 
      Height          =   450
      ItemData        =   "Frm_RepLiqPagoPen.frx":0000
      Left            =   240
      List            =   "Frm_RepLiqPagoPen.frx":0002
      MultiSelect     =   1  'Simple
      TabIndex        =   8
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Btn_Generar 
      Caption         =   "Generar y descargar"
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Periodo (DD/MM/YYYY)"
      Height          =   1935
      Left            =   3600
      TabIndex        =   1
      Top             =   240
      Width           =   2175
      Begin VB.TextBox Txt_FechaHasta 
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Txt_FechaDesde 
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Desde:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de pago"
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2895
      Begin VB.ListBox Lst_TP 
         Height          =   1425
         ItemData        =   "Frm_RepLiqPagoPen.frx":0004
         Left            =   240
         List            =   "Frm_RepLiqPagoPen.frx":0006
         MultiSelect     =   1  'Simple
         TabIndex        =   7
         Top             =   360
         Width           =   2295
      End
   End
End
Attribute VB_Name = "Frm_RepLiqPagoPen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Reporte de pagos

Private Sub Btn_Generar_Click()
On Error GoTo Err_Btn_Generar_Click
    
    Dim vlAlgunCheck As Boolean
    Dim TipoPagos As String
    Dim vlFechaIni As String
    Dim vlFechaTer As String
    
    'Validacion de los checkbox
    vlAlgunCheck = False
    TipoPagos = "("
    
    For i = 0 To Lst_TP.ListCount - 1
        If Lst_TP.Selected(i) Then
            TipoPagos = TipoPagos & "'" & Lst_TPi.List(i) & "',"
            vlAlgunCheck = True
        End If
    Next i
    
    If vlAlgunCheck = False Then
        MsgBox "Debe de seleccionar algún tipo de pago", vbCritical, "Error de Datos"
        Exit Sub
    Else
        TipoPagos = Left(TipoPagos, Len(TipoPagos) - 1) & ") " 'Se quita la ultima coma (,)
    End If
    
    'Validacion de la fecha de inicio de reporte
    If (Trim(Txt_FechaDesde) = "") Then
       MsgBox "Debe ingresar una Fecha de Inicio de periodo", vbCritical, "Error de Datos"
       Txt_FechaDesde.SetFocus
       Exit Sub
    End If
    If Not IsDate(Txt_FechaDesde.Text) Then
       MsgBox "La Fecha Ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
       Txt_FechaDesde.SetFocus
       Exit Sub
    End If
    If (CDate(Txt_FechaDesde) > CDate(Date)) Then
       MsgBox "La Fecha Ingresada es Mayor a la Fecha Actual", vbCritical, "Error de Datos"
       Txt_FechaDesde.SetFocus
       Exit Sub
    End If
    If (Year(Txt_FechaDesde) < 1900) Then
       MsgBox "La Fecha Ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
       Txt_FechaDesde.SetFocus
       Exit Sub
    End If
    
    Txt_FechaDesde.Text = Format(CDate(Trim(Txt_FechaDesde)), "yyyymmdd")
    Txt_FechaDesde.Text = DateSerial(Mid((Txt_FechaDesde.Text), 1, 4), Mid((Txt_FechaDesde.Text), 5, 2), Mid((Txt_FechaDesde.Text), 7, 2))

    'Validacion de la fecha de término de reporte
    If (Trim(Txt_FechaHasta) = "") Then
       MsgBox "Debe ingresar una Fecha de término de periodo", vbCritical, "Error de Datos"
       Txt_FechaHasta.SetFocus
       Exit Sub
    End If
    If Not IsDate(Txt_FechaHasta.Text) Then
       MsgBox "La Fecha Ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
       Txt_FechaHasta.SetFocus
       Exit Sub
    End If
    'If (CDate(Txt_FechaHasta) > CDate(Date)) Then
    '   MsgBox "La Fecha Ingresada es Mayor a la Fecha Actual", vbCritical, "Error de Datos"
    '   Txt_FechaHasta.SetFocus
    '   Exit Sub
    'End If
    If (Year(Txt_FechaHasta) < 1900) Then
       MsgBox "La Fecha Ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
       Txt_FechaHasta.SetFocus
       Exit Sub
    End If
    
    Txt_FechaHasta.Text = Format(CDate(Trim(Txt_FechaHasta)), "yyyymmdd")
    Txt_FechaHasta.Text = DateSerial(Mid((Txt_FechaHasta.Text), 1, 4), Mid((Txt_FechaHasta.Text), 5, 2), Mid((Txt_FechaHasta.Text), 7, 2))
    
    vlFechaIni = Format(CDate(Trim(Txt_FechaDesde)), "yyyymmdd")
    vlFechaTer = Format(CDate(Trim(Txt_FechaHasta)), "yyyymmdd")
    
'    vgQuery = "SELECT "
'    vgQuery = vgQuery & "NUM_PERPAGO,NUM_POLIZA,NUM_ENDOSO,NUM_ORDEN,COD_TIPOPAGO,GLS_DIRECCION,COD_DIRECCION,COD_TIPPENSION,FEC_PAGO,COD_VIAPAGO,COD_BANCO,COD_TIPCUENTA,NUM_CUENTA,COD_SUCURSAL,COD_INSSALUD,NUM_IDENRECEPTOR,COD_TIPOIDENRECEPTOR,GLS_NOMRECEPTOR,GLS_NOMSEGRECEPTOR,GLS_PATRECEPTOR,GLS_MATRECEPTOR,COD_TIPRECEPTOR,NUM_CARGAS,MTO_HABER,MTO_DESCUENTO,MTO_LIQPAGAR,MTO_BASEIMP,MTO_BASETRI,MTO_PENSION,COD_MODSALUD,MTO_PLANSALUD,COD_MONEDA "
'    vgQuery = vgQuery & "FROM pp_tmae_liqpagopendef "
'    vgQuery = vgQuery & "WHERE COD_TIPOPAGO IN " & TipoPagos
'    vgQuery = vgQuery & "AND FEC_PAGO >= '" & vlFechaIni & "' "
'    vgQuery = vgQuery & "AND FEC_PAGO <= '" & vlFechaTer & "' "
'    'vgQuery = vgQuery & "AND ROWNUM < 100 " 'QUITAR ESTA LINEA DESPUES!!!
'    vgQuery = vgQuery & "ORDER BY COD_TIPOPAGO ASC, FEC_PAGO ASC "
    
    
    vgQuery = " SELECT "
    vgQuery = vgQuery & " l.NUM_PERPAGO,l.NUM_POLIZA,l.NUM_ENDOSO,l.NUM_ORDEN,l.COD_TIPOPAGO,l.GLS_DIRECCION,l.COD_DIRECCION,l.COD_TIPPENSION||' - '||tp.gls_elemento COD_TIPPENSION,"
    vgQuery = vgQuery & " l.FEC_PAGO,l.COD_VIAPAGO,l.COD_BANCO,l.COD_TIPCUENTA,l.NUM_CUENTA,l.COD_SUCURSAL,l.COD_INSSALUD,l.NUM_IDENRECEPTOR,"
    vgQuery = vgQuery & " l.COD_TIPOIDENRECEPTOR,l.GLS_NOMRECEPTOR,l.GLS_NOMSEGRECEPTOR,l.GLS_PATRECEPTOR,l.GLS_MATRECEPTOR,l.COD_TIPRECEPTOR, "
    vgQuery = vgQuery & " l.NUM_CARGAS,l.MTO_HABER,l.MTO_DESCUENTO,l.MTO_LIQPAGAR,l.MTO_BASEIMP,l.MTO_BASETRI,l.MTO_PENSION,"
    vgQuery = vgQuery & " l.COD_MODSALUD,l.MTO_PLANSALUD,l.COD_MONEDA, "
    vgQuery = vgQuery & "  p.cod_AFP||' - '||af.gls_elemento AFP, p.cod_cuspp CUSPP "
    vgQuery = vgQuery & " FROM pp_tmae_liqpagopendef l "
    vgQuery = vgQuery & " inner join pp_tmae_poliza p on "
    vgQuery = vgQuery & "  l.num_poliza= p.num_poliza and "
    vgQuery = vgQuery & "  p.num_endoso = (select max(num_endoso) from pp_tmae_poliza where num_poliza = p.num_poliza)"
    vgQuery = vgQuery & " left Join ma_tpar_tabcod TP on "
    vgQuery = vgQuery & "   tp.cod_tabla='TP' "
    vgQuery = vgQuery & "   and tp.cod_elemento = l.cod_tippension "
    vgQuery = vgQuery & "  left join ma_tpar_tabcod AF on "
    vgQuery = vgQuery & "   af.cod_tabla ='AF' "
    vgQuery = vgQuery & "   and af.cod_elemento = p.cod_afp "
    vgQuery = vgQuery & "  WHERE COD_TIPOPAGO IN " & TipoPagos
    vgQuery = vgQuery & " AND FEC_PAGO >= '" & vlFechaIni & "' "
    vgQuery = vgQuery & " AND FEC_PAGO <= '" & vlFechaTer & "' "
   ' vgQuery = vgQuery & " AND ROWNUM < 100   --" 'QUITAR ESTA LINEA DESPUES!!!
    vgQuery = vgQuery & " ORDER BY COD_TIPOPAGO ASC, FEC_PAGO ASC"
   
      
    Set vgRegistro = vgConexionBD.Execute(vgQuery)
    
    Dim xlapp As Excel.Application
    'Dim WBRes As Workbook
    'Dim vlArchivo As String
    Dim iFila As Long
    Dim iCol As Integer
    
    If Not vgRegistro.EOF Then
        
        'vlArchivo = strRpt & "Archivo_Pagos.xls"
        Set xlapp = CreateObject("excel.application")
        
        xlapp.Visible = True 'para ver vista previa
        xlapp.WindowState = 2 ' minimiza excel
        xlapp.Workbooks.Add (xlWBATWorksheet)
        xlapp.Worksheets(1).Activate
        xlapp.ActiveSheet.Name = "Reporte de pagos"
        
        xlapp.ActiveSheet.Cells(1, 1).Value = "NUM_PERPAGO"
        xlapp.ActiveSheet.Cells(1, 2).Value = "NUM_POLIZA"
        xlapp.ActiveSheet.Cells(1, 3).Value = "NUM_ENDOSO"
        xlapp.ActiveSheet.Cells(1, 4).Value = "NUM_ORDEN"
        xlapp.ActiveSheet.Cells(1, 5).Value = "COD_TIPOPAGO"
        xlapp.ActiveSheet.Cells(1, 6).Value = "GLS_DIRECCION"
        xlapp.ActiveSheet.Cells(1, 7).Value = "COD_DIRECCION"
        xlapp.ActiveSheet.Cells(1, 8).Value = "TIPO PENSION"
        xlapp.ActiveSheet.Cells(1, 9).Value = "FEC_PAGO"
        xlapp.ActiveSheet.Cells(1, 10).Value = "COD_VIAPAGO"
        xlapp.ActiveSheet.Cells(1, 11).Value = "COD_BANCO"
        xlapp.ActiveSheet.Cells(1, 12).Value = "COD_TIPCUENTA"
        xlapp.ActiveSheet.Cells(1, 13).Value = "NUM_CUENTA"
        xlapp.ActiveSheet.Cells(1, 14).Value = "COD_SUCURSAL"
        xlapp.ActiveSheet.Cells(1, 15).Value = "COD_INSSALUD"
        xlapp.ActiveSheet.Cells(1, 16).Value = "NUM_IDENRECEPTOR"
        xlapp.ActiveSheet.Cells(1, 17).Value = "COD_TIPOIDENRECEPTOR"
        xlapp.ActiveSheet.Cells(1, 18).Value = "GLS_NOMRECEPTOR"
        xlapp.ActiveSheet.Cells(1, 19).Value = "GLS_NOMSEGRECEPTOR"
        xlapp.ActiveSheet.Cells(1, 20).Value = "GLS_PATRECEPTOR"
        xlapp.ActiveSheet.Cells(1, 21).Value = "GLS_MATRECEPTOR"
        xlapp.ActiveSheet.Cells(1, 22).Value = "COD_TIPRECEPTOR"
        xlapp.ActiveSheet.Cells(1, 23).Value = "NUM_CARGAS"
        xlapp.ActiveSheet.Cells(1, 24).Value = "MTO_HABER"
        xlapp.ActiveSheet.Cells(1, 25).Value = "MTO_DESCUENTO"
        xlapp.ActiveSheet.Cells(1, 26).Value = "MTO_LIQPAGAR"
        xlapp.ActiveSheet.Cells(1, 27).Value = "MTO_BASEIMP"
        xlapp.ActiveSheet.Cells(1, 28).Value = "MTO_BASETRI"
        xlapp.ActiveSheet.Cells(1, 29).Value = "MTO_PENSION"
        xlapp.ActiveSheet.Cells(1, 30).Value = "COD_MODSALUD"
        xlapp.ActiveSheet.Cells(1, 31).Value = "MTO_PLANSALUD"
        xlapp.ActiveSheet.Cells(1, 32).Value = "COD_MONEDA"
        xlapp.ActiveSheet.Cells(1, 33).Value = "AFP"
        xlapp.ActiveSheet.Cells(1, 34).Value = "CUSPP"
       
        
        
        iFila = 2
        While Not vgRegistro.EOF
            xlapp.ActiveSheet.Cells(iFila, 1).Value = vgRegistro!Num_PerPago
            xlapp.ActiveSheet.Cells(iFila, 2).Value = vgRegistro!num_poliza
            xlapp.ActiveSheet.Cells(iFila, 3).Value = vgRegistro!num_endoso
            xlapp.ActiveSheet.Cells(iFila, 4).Value = vgRegistro!Num_Orden
            xlapp.ActiveSheet.Cells(iFila, 5).Value = vgRegistro!Cod_TipoPago
            xlapp.ActiveSheet.Cells(iFila, 6).Value = vgRegistro!Gls_Direccion
            xlapp.ActiveSheet.Cells(iFila, 7).Value = vgRegistro!Cod_Direccion
            xlapp.ActiveSheet.Cells(iFila, 8).Value = vgRegistro!Cod_TipPension
            xlapp.ActiveSheet.Cells(iFila, 9).Value = vgRegistro!fec_pago
            xlapp.ActiveSheet.Cells(iFila, 10).Value = vgRegistro!Cod_ViaPago
            xlapp.ActiveSheet.Cells(iFila, 11).Value = vgRegistro!Cod_Banco
            xlapp.ActiveSheet.Cells(iFila, 12).Value = vgRegistro!Cod_TipCuenta
            xlapp.ActiveSheet.Cells(iFila, 13).Value = vgRegistro!Num_Cuenta
            xlapp.ActiveSheet.Cells(iFila, 14).Value = vgRegistro!Cod_Sucursal
            xlapp.ActiveSheet.Cells(iFila, 15).Value = vgRegistro!Cod_InsSalud
            xlapp.ActiveSheet.Cells(iFila, 16).Value = vgRegistro!Num_IdenReceptor
            xlapp.ActiveSheet.Cells(iFila, 17).Value = vgRegistro!Cod_TipoIdenReceptor
            xlapp.ActiveSheet.Cells(iFila, 18).Value = vgRegistro!Gls_NomReceptor
            xlapp.ActiveSheet.Cells(iFila, 19).Value = vgRegistro!Gls_NomSegReceptor
            xlapp.ActiveSheet.Cells(iFila, 20).Value = vgRegistro!Gls_PatReceptor
            xlapp.ActiveSheet.Cells(iFila, 21).Value = vgRegistro!Gls_MatReceptor
            xlapp.ActiveSheet.Cells(iFila, 22).Value = vgRegistro!Cod_TipReceptor
            xlapp.ActiveSheet.Cells(iFila, 23).Value = vgRegistro!Num_Cargas
            xlapp.ActiveSheet.Cells(iFila, 24).Value = vgRegistro!Mto_Haber
            xlapp.ActiveSheet.Cells(iFila, 25).Value = vgRegistro!Mto_Descuento
            xlapp.ActiveSheet.Cells(iFila, 26).Value = vgRegistro!Mto_LiqPagar
            xlapp.ActiveSheet.Cells(iFila, 27).Value = vgRegistro!Mto_BaseImp
            xlapp.ActiveSheet.Cells(iFila, 28).Value = vgRegistro!Mto_BaseTri
            xlapp.ActiveSheet.Cells(iFila, 29).Value = vgRegistro!Mto_Pension
            xlapp.ActiveSheet.Cells(iFila, 30).Value = vgRegistro!Cod_ModSalud
            xlapp.ActiveSheet.Cells(iFila, 31).Value = vgRegistro!Mto_PlanSalud
            xlapp.ActiveSheet.Cells(iFila, 32).Value = vgRegistro!Cod_Moneda
            xlapp.ActiveSheet.Cells(iFila, 33).Value = vgRegistro!AFP
            xlapp.ActiveSheet.Cells(iFila, 34).Value = vgRegistro!CUSPP
            
            iFila = iFila + 1
            vgRegistro.MoveNext
        Wend
        xlapp.ActiveSheet.Cells(1, 1).EntireRow.Insert Shift:=xlDown
        xlapp.ActiveSheet.Range("A1:AH2").Font.Bold = True
        xlapp.ActiveSheet.Range("A1:AH2").Interior.ColorIndex = 15
        xlapp.ActiveSheet.UsedRange.Columns.AutoFit
        'Hasta aqui es el codigo de la lista normal
        
        'aqui inicia la colocación de multiples encabezados segun el tipo de pago
        xlapp.ActiveSheet.Cells(1, 1).Value = "<Aqui va el tipo de pago>"
        Dim UltimaFila As Long
        UltimaFila = iFila
        Dim TipoActual As String
        
        TipoActual = xlapp.ActiveSheet.Cells(3, 5) 'El Cod_TipoPago actual
        For i = 0 To Lst_TP.ListCount - 1
            If Lst_TPi.List(i) = TipoActual Then
                xlapp.ActiveSheet.Cells(1, 1).Value = "Tipo de pago: " & Lst_TP.List(i)
            End If
        Next i
        
        iFila = 3 'Reiniciando el conteo
        While iFila <= UltimaFila
            'Si el tipo de pago cambia, agregar nuevamente el encabezado
            If TipoActual <> xlapp.ActiveSheet.Cells(iFila, 5) Then
                xlapp.ActiveSheet.Cells(2, 1).EntireRow.Copy
                xlapp.ActiveSheet.Cells(iFila, 1).EntireRow.Insert Shift:=xlDown
                xlapp.ActiveSheet.Cells(1, 1).EntireRow.Copy
                xlapp.ActiveSheet.Cells(iFila, 1).EntireRow.Insert Shift:=xlDown
                xlapp.CutCopyMode = False
                
                iFila = iFila + 2
                UltimaFila = UltimaFila + 2
                
                TipoActual = xlapp.ActiveSheet.Cells(iFila, 5)
                For i = 0 To Lst_TP.ListCount - 1
                    If Lst_TPi.List(i) = TipoActual Then
                        xlapp.ActiveSheet.Cells(iFila - 2, 1).Value = "Tipo de pago: " & Lst_TP.List(i)
                    End If
                Next i
                
            End If
            iFila = iFila + 1
        Wend
        xlapp.CutCopyMode = False
        
        xlapp.WindowState = 1 'Maximiza excel
    Else
        MsgBox "No Existen Registros para los Conceptos Seleccionados", vbInformation, "Información"
    End If
    
    'MsgBox vgQuery, vbInformation, "Dbug"
    vgRegistro.Close
Exit Sub
Err_Btn_Generar_Click:
    Screen.MousePointer = 0
    MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    If vgRegistro.State <> 0 Then
        vgRegistro.Close
    End If
End Sub

Private Sub Form_Load()
On Error GoTo Err_Form_Load
    
    vgQuery = "SELECT COD_ELEMENTO,GLS_ELEMENTO FROM MA_TPAR_TABCOD WHERE COD_TABLA='TRG'"
    Set vgRegistro = vgConexionBD.Execute(vgQuery)
    Lst_TPi.Clear
    Lst_TP.Clear
    If Not vgRegistro.EOF Then
        While Not vgRegistro.EOF
            Lst_TPi.AddItem vgRegistro!COD_ELEMENTO
            Lst_TP.AddItem vgRegistro!GLS_ELEMENTO
            vgRegistro.MoveNext
        Wend
        
        For i = 0 To Lst_TP.ListCount - 1
            Lst_TP.Selected(i) = True
        Next i
    End If
    vgRegistro.Close
Exit Sub
Err_Form_Load:
    Screen.MousePointer = 0
    MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    If vgRegistro.State <> 0 Then
        vgRegistro.Close
    End If
End Sub
Private Sub Txt_FechaDesde_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Txt_FechaHasta.SetFocus
    End If
End Sub
Private Sub Txt_FechaHasta_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Btn_Generar.SetFocus
    End If
End Sub

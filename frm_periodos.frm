VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_periodos 
   Caption         =   "Reporte de Log"
   ClientHeight    =   1590
   ClientLeft      =   2745
   ClientTop       =   3645
   ClientWidth     =   6180
   Icon            =   "frm_periodos.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1590
   ScaleWidth      =   6180
   Begin VB.CommandButton cmd_salir 
      Height          =   645
      Left            =   4920
      Picture         =   "frm_periodos.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   510
      Width           =   615
   End
   Begin VB.CommandButton cmd_grabar 
      Height          =   645
      Left            =   4260
      Picture         =   "frm_periodos.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Grabar Movimiento"
      Top             =   510
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Contabilización : "
      Height          =   735
      Left            =   2370
      TabIndex        =   2
      Top             =   450
      Width           =   1635
      Begin MSComCtl2.DTPicker dtp_dfinal 
         Height          =   285
         Left            =   210
         TabIndex        =   3
         Top             =   300
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   503
         _Version        =   393216
         Format          =   50331649
         CurrentDate     =   39344
      End
   End
   Begin VB.Frame fra_fecha 
      Caption         =   "Contabilización : "
      Height          =   735
      Left            =   540
      TabIndex        =   0
      Top             =   450
      Width           =   1635
      Begin MSComCtl2.DTPicker dtp_dinicio 
         Height          =   285
         Left            =   210
         TabIndex        =   1
         Top             =   300
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   503
         _Version        =   393216
         Format          =   50331649
         CurrentDate     =   39344
      End
   End
End
Attribute VB_Name = "frm_periodos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_grabar_Click()

'Dim rs_temp As ADODB.Recordset

'    If Format(dtp_dfinal.Value, "yyyymmdd") < Format(dtp_dinicio.Value, "yyyymmdd") Then
'        MsgBox "Fecha Final no puede ser menor a la Inicial", vbCritical
'        Exit Sub
'    End If
'
'    sSql = "Select LOG_NNUMREG, LOG_CCODERR, TGR_CDESCRI , LOG_SQUERY, LOG_DFECREG, LOG_CUSUCRE "
'    sSql = sSql & " from PT_LOGACTUAL inner join tablagrlg on tgr_ccodtab=substring(LOG_CCODERR,1,3) and tgr_ccoddet=substring(LOG_CCODERR,4,3)"
'    sSql = sSql & " AND CONVERT(NVARCHAR(8), LOG_DFECREG,112)>=" & Format(dtp_dinicio.Value, "yyyymmdd") & " "
'    sSql = sSql & " AND CONVERT(NVARCHAR(8), LOG_DFECREG,112)<=" & Format(dtp_dfinal.Value, "yyyymmdd")
'
'    cnnConexion.Open
'    Set rs_temp = New ADODB.Recordset
'    Set rs_temp = cnnConexion.Execute(sSql)
'    If Not rs_temp.EOF Then
'        Call p_crear_rs
'        Do Until rs_temp.EOF
'            objRsRpt.AddNew
 '           objRsRpt.Fields("TMP_NNUMREG").Value = rs_temp!LOG_NNUMREG
 ''           objRsRpt.Fields("TMP_CCODERR").Value = rs_temp!LOG_CCODERR
 '           objRsRpt.Fields("TMP_CDESCRI").Value = Mid(rs_temp!TGR_CDESCRI, 1, 60)
 '           objRsRpt.Fields("TMP_SQUERY").Value = Mid(rs_temp!LOG_SQUERY, 1, 300)
 '           objRsRpt.Fields("TMP_DFECREG").Value = Trim(rs_temp!LOG_DFECREG)
 '           objRsRpt.Fields("TMP_NFECREG").Value = Format(rs_temp!LOG_DFECREG, "YYYYMMDD")
 '           objRsRpt.Fields("TMP_CUSUCRE").Value = Trim(rs_temp!LOG_CUSUCRE)
 '           objRsRpt.Update
 '           rs_temp.MoveNext
 '       Loop
 '       svrpt_sTitulo = "Del " & dtp_dinicio.Value & " Al " & dtp_dfinal.Value
 '       sName_Reporte = "Rpt_Log_Registros.rpt"
 '       frm_plantilla.Show
 '   End If
 '   cnnConexion.Close
    
  'If Trim(txt_FecPago) = "" Then
  '      MsgBox "Falta Ingresar Fecha Hasta", vbCritical, "Falta Información"
  '      Txt_Hasta.SetFocus
  '      Exit Sub
  '  End If
    
    On Error GoTo Err_flInformeCerEst
'Certificados de Supervivencia
   Screen.MousePointer = 11
   
   'marco 11/03/2010
   Dim cadena As String
   Dim objRep As New ClsReporte
   Dim vlFechaPago, vlFinFechaPago As String
   Dim rs As New ADODB.Recordset
   
   
    vlFechaPago = Mid(Format(CDate(Trim(dtp_dinicio.Value)), "yyyymmdd"), 1, 8)  ' "201408" 'Mid(Format(CDate(Trim(txt_FecPago.Text)), "yyyymmdd"), 1, 6)
    vlFinFechaPago = Mid(Format(CDate(Trim(dtp_dfinal.Value)), "yyyymmdd"), 1, 8)
   
    'cadena = "select * from PP_TMAE_LOGACTUAL a join pp_tpar_error b on a.log_ccoderr=b.cod_error where log_dfecreg between '" & vlFechaPago & "' and '" & vlFinFechaPago & "' order by 1"
    
    cadena = "select log_nnumreg, log_ccoderr, log_dfecreg, log_cusucre, gls_error,"
    cadena = cadena & " substr(log_squery, 1, 255) as log_squery1,"
    cadena = cadena & " substr(log_squery, 256, 255) as log_squery2,"
    cadena = cadena & " substr(log_squery, 511, 255) as log_squery3,"
    cadena = cadena & " substr(log_squery, 776, 255) As log_squery4"
    cadena = cadena & " from PP_TMAE_LOGACTUAL a"
    cadena = cadena & " join pp_tpar_error b on a.log_ccoderr=b.cod_error"
    cadena = cadena & " where log_dfecreg between '" & vlFechaPago & "' and '" & vlFinFechaPago & "' order by 1"
    

    Set rs = vgConexionBD.Execute(cadena)
    If Not rs.EOF Then
        LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PP_TMAE_LOGACTUAL.rpt"), ".RPT", ".TTX"), 1)
        
            
        If objRep.CargaReporte(strRpt & "", "PP_Rpt_Log.rpt", "Informe de log de transacciones", rs, True, _
                                ArrFormulas("NombreCompania", vgNombreCompania), _
                                ArrFormulas("NombreSistema", vgNombreSistema), _
                                ArrFormulas("NombreSubSistema", vgNombreSubSistema)) = False Then
                                
            MsgBox "No se pudo abrir el reporte", vbInformation
            Screen.MousePointer = 0
            Exit Sub
        End If
    Else
        MsgBox ("No existen datos en este periodo.")
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

Private Sub Form_Load()

    Me.Width = 6300
    Me.Height = 2100
    'Call p_centerForm(mdi_principal, Me)
    dtp_dfinal.Value = Format(Date, "dd/mm/yyyy")
    dtp_dinicio.Value = Format(Date, "dd/mm/yyyy")

End Sub

'Private Sub p_crear_rs()
'
'    Set objRsRpt = New ADODB.Recordset
'
'    objRsRpt.Fields.Append "TMP_NNUMREG", adInteger
'    objRsRpt.Fields.Append "TMP_CCODERR", adVarChar, 6
'    objRsRpt.Fields.Append "TMP_CDESCRI", adVarChar, 60
'    objRsRpt.Fields.Append "TMP_SQUERY", adVarChar, 300
'    objRsRpt.Fields.Append "TMP_DFECREG", adVarChar, 30
'    objRsRpt.Fields.Append "TMP_NFECREG", adVarChar, 8
'    objRsRpt.Fields.Append "TMP_CUSUCRE", adVarChar, 10
'
'    objRsRpt.Open
'
'End Sub


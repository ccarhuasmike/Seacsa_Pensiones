VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frm_plantilla 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes de Producción"
   ClientHeight    =   8445
   ClientLeft      =   240
   ClientTop       =   1680
   ClientWidth     =   13260
   Icon            =   "frm_plantilla.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   13260
   ShowInTaskbar   =   0   'False
   Begin CRVIEWERLibCtl.CRViewer crw_object 
      Height          =   8385
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   13140
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   0   'False
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frm_plantilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objApp As New CRAXDRT.Application
Private objReport As CRAXDRT.Report
Private objSubReport As CRAXDRT.Report
Private objTotReport As CRAXDRT.Report

Private Sub Form_Load()
Dim vgTipoValor As String

Dim vgMoneda As String
Dim objRsRpt As New ADODB.Recordset



  'Call p_centerForm(mdi_principal, Me)
  Set objReport = objApp.OpenReport(strRpt & "" & sName_Reporte)
  Select Case sName_Reporte
    Case "PP_Rpt_CertificadoSuperv.rpt"

      objReport.ParameterFields.GetItemByName("Ident_1").AddCurrentValue (Ident_1)
      objReport.ParameterFields.GetItemByName("Ident_2").AddCurrentValue (Ident_2)
      objReport.ParameterFields.GetItemByName("Tipo_J").AddCurrentValue (Tipo_J)
      objReport.ParameterFields.GetItemByName("Tipo_S").AddCurrentValue (Tipo_S)
      objReport.ParameterFields.GetItemByName("Tipo_I").AddCurrentValue (Tipo_I)
      objReport.ParameterFields.GetItemByName("Nombre_Afiliado").AddCurrentValue (Nombre_Afiliado)
      objReport.ParameterFields.GetItemByName("Nombre_Beneficiario").AddCurrentValue (Nombre_Beneficiario)
      objReport.ParameterFields.GetItemByName("Documento").AddCurrentValue (Tipo_num_documento_afiliado)
      objReport.ParameterFields.GetItemByName("Documento_Beneficiario").AddCurrentValue (Tipo_num_documento_beneficiario)
      objReport.ParameterFields.GetItemByName("Poliza").AddCurrentValue (num_poliza)
      objReport.ParameterFields.GetItemByName("Rango_Fechas").AddCurrentValue (RangoFecha)
      objReport.ParameterFields.GetItemByName("Fecha_Creacion").AddCurrentValue (CDate(Fecha_Creacion))
      
      objReport.Database.SetDataSource objRsRpt
      crw_object.ReportSource = objReport
  End Select
  crw_object.ViewReport
  Screen.MousePointer = 0
  Set objReport = Nothing
  DoEvents

End Sub

Private Sub Form_Resize()
'crw_object.Top = 10
'crw_object.Left = 10
'crw_object.Width = Me.ScaleWidth - 50
'crw_object.Height = Me.ScaleHeight - 50
End Sub

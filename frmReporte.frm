VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmReporte 
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6870
   Icon            =   "frmReporte.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSetup 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1185
      Picture         =   "frmReporte.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Configurar Impresora"
      Top             =   0
      Width           =   375
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer 
      Height          =   3255
      Left            =   255
      TabIndex        =   0
      Top             =   480
      Width           =   5280
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   0   'False
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Option Explicit
   Dim oReporte As New CRAXDRT.Report
   Dim vSave As Boolean
   Dim vSetup As Boolean


Private Sub cmdSetup_Click()
oReporte.PrinterSetup (0)
End Sub

Private Sub Form_Load()
   'Me.Icon = LoadResPicture("REPORTE", vbResIcon)
   Screen.MousePointer = vbDefault
   'ToolRepo.Left = CRViewer.Left + 5750
   'ToolRepo.Top = CRViewer.Top + 55
   
   cmdSetup.Left = 8190
   cmdSetup.Top = 60

End Sub

Private Sub Form_Resize()
'   If vSetup = True Or vSave = True Then
'       CRViewer.Top = 500
'   ElseIf vSetup = False Or vSave = False Then
'       CRViewer.Top = 0
'   End If
   CRViewer.Top = 0
   CRViewer.Left = 0
   CRViewer.Height = ScaleHeight
   CRViewer.Width = ScaleWidth
End Sub

Public Sub SetReporte(rptReporteCrystal As CRAXDRT.Report)
   Screen.MousePointer = vbHourglass
   Set oReporte = rptReporteCrystal
   CRViewer.ReportSource = oReporte
   CRViewer.DisplayGroupTree = False
   CRViewer.EnableExportButton = True
    If vSetup = True Then
        cmdSetup.Visible = True
    Else
        cmdSetup.Visible = False
    End If
       
   'Call oReporte.SaveAs("", cr70FileFormat)
   'oReporte.PrintOut False
   CRViewer.ViewReport
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Public Property Let Save(ByVal vNewValue As Boolean)
    vSave = vNewValue
End Property
Public Property Let Setup(ByVal vNewValue As Boolean)
    vSetup = vNewValue
End Property

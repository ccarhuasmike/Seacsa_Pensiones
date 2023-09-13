VERSION 5.00
Begin VB.Form frmMayoresSinCertEst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hijos Mayores Sin Certificado de Estudio"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6660
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   6660
   Begin VB.Frame Frame1 
      Height          =   1200
      Left            =   120
      TabIndex        =   2
      Top             =   105
      Width           =   6420
      Begin VB.TextBox txtAño 
         Height          =   285
         Left            =   1695
         TabIndex        =   4
         Top             =   450
         Width           =   855
      End
      Begin VB.ComboBox CmbMes 
         Height          =   315
         Left            =   3285
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   420
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Año:"
         Height          =   240
         Left            =   1275
         TabIndex        =   6
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Mes:"
         Height          =   210
         Left            =   2730
         TabIndex        =   5
         Top             =   465
         Width           =   495
      End
   End
   Begin VB.CommandButton Cmd_Salir 
      Caption         =   "&Salir"
      Height          =   675
      Left            =   4020
      Picture         =   "frmMayoresSinCertEst.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir del Formulario"
      Top             =   1365
      Width           =   720
   End
   Begin VB.CommandButton Cmd_Imprimir 
      Caption         =   "&Imprimir"
      Height          =   675
      Left            =   1860
      Picture         =   "frmMayoresSinCertEst.frx":00FA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1365
      Width           =   720
   End
End
Attribute VB_Name = "frmMayoresSinCertEst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset

Private Sub Cmd_Imprimir_Click()

Dim cadena As String
Dim Periodo As String
Set rs = New ADODB.Recordset
Dim objRep As New ClsReporte

On Error GoTo mierror

If txtAño.Text = "" Or Len(txtAño.Text) < 4 Then
    MsgBox "Tiene que ingresar un año valido.", vbCritical, ""
    Exit Sub
End If
If CmbMes.ListIndex = 0 Then
    MsgBox "Seleccione un mes del año.", vbCritical, ""
    Exit Sub
End If

Periodo = txtAño.Text & Format(CmbMes.ListIndex, "00")

Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "PP_LISTA_MAYORES_SIN_CERT.Lista(" & Periodo & ")", vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
If Not rs.EOF Then
    
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PP_Rpt_MayoresSinCert.rpt"), ".RPT", ".TTX"), 1)
    
        
    If objRep.CargaReporte(strRpt & "", "PP_Rpt_MayoresSinCert.rpt", "Informe de Polizas Con beneficiarios Hijos Sin Certificado de Estudio", rs, True, _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("NombreSistema", vgNombreSistema), _
                            ArrFormulas("NombreSubSistema", vgNombreSubSistema), _
                            ArrFormulas("Periodo", Periodo)) = False Then
                            
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
    End If
Else
    MsgBox "No existe información para el periodo establecido.", vbInformation, ""
    Exit Sub
End If


Exit Sub
mierror:
    MsgBox "Ocurrio un problema al cargar el reporte, consulte con sistemas", vbInformation
    

End Sub

Private Sub cmd_salir_Click()
Unload Me

End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0

CmbMes.AddItem ("--Seleccionar--")
CmbMes.ItemData(CmbMes.NewIndex) = 0
CmbMes.AddItem ("ENERO")
CmbMes.ItemData(CmbMes.NewIndex) = 1
CmbMes.AddItem ("FEBRERO")
CmbMes.ItemData(CmbMes.NewIndex) = 2
CmbMes.AddItem ("MARZO")
CmbMes.ItemData(CmbMes.NewIndex) = 3
CmbMes.AddItem ("ABRIL")
CmbMes.ItemData(CmbMes.NewIndex) = 4
CmbMes.AddItem ("MAYO")
CmbMes.ItemData(CmbMes.NewIndex) = 5
CmbMes.AddItem ("JUNIO")
CmbMes.ItemData(CmbMes.NewIndex) = 6
CmbMes.AddItem ("JULIO")
CmbMes.ItemData(CmbMes.NewIndex) = 7
CmbMes.AddItem ("AGOSTO")
CmbMes.ItemData(CmbMes.NewIndex) = 8
CmbMes.AddItem ("SETIEMBRE")
CmbMes.ItemData(CmbMes.NewIndex) = 9
CmbMes.AddItem ("OCTUBRE")
CmbMes.ItemData(CmbMes.NewIndex) = 10
CmbMes.AddItem ("NOVIEMBRE")
CmbMes.ItemData(CmbMes.NewIndex) = 11
CmbMes.AddItem ("DICIEMBRE")
CmbMes.ItemData(CmbMes.NewIndex) = 12

CmbMes.ListIndex = 0

End Sub

Private Sub txtAño_KeyPress(KeyAscii As Integer)

Select Case (KeyAscii)
Case 48 To 57
Case vbKeyBack
Case 46
Case Else
    KeyAscii = 0
End Select
End Sub


VERSION 5.00
Begin VB.Form frmCertificadosVencidos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Certificados Vencidos"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4935
   Icon            =   "frmCertificadosVencidos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4935
   Begin VB.TextBox Txt_Hasta 
      Height          =   300
      Left            =   3390
      TabIndex        =   0
      Top             =   330
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   855
      Width           =   4695
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   3030
         Picture         =   "frmCertificadosVencidos.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Salir del Formulario"
         Top             =   195
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   1950
         Picture         =   "frmCertificadosVencidos.frx":053C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpiar Formulario"
         Top             =   195
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   870
         Picture         =   "frmCertificadosVencidos.frx":0BF6
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   195
         Width           =   720
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Certificados Vencidos a la fecha del :"
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
      Height          =   255
      Left            =   105
      TabIndex        =   5
      Top             =   390
      Width           =   3390
   End
End
Attribute VB_Name = "frmCertificadosVencidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rs As ADODB.Recordset

Private Sub Cmd_Imprimir_Click()

    If Trim(Txt_Hasta) = "" Then
        MsgBox "Falta Ingresar Fecha Hasta", vbCritical, "Falta Información"
        Txt_Hasta.SetFocus
        Exit Sub
    Else
        If Not IsDate(Txt_Hasta) Then
            MsgBox "Fecha Hasta no es una fecha Válida", vbCritical, "Falta Información"
            Txt_Hasta = ""
            Exit Sub
        End If
    End If
    
    On Error GoTo Err_flInformeCerEst
'Certificados de Supervivencia
   Screen.MousePointer = 11
   
   'marco 11/03/2010
   Dim cadena As String
   Dim objRep As New ClsReporte
   vlFechaTermino = Format(CDate(Trim(Txt_Hasta.Text)), "yyyymmdd")
   vgPalabra = ""
   vgPalabra = "Certificados vencidos al " & Txt_Hasta.Text
   
   
   Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "PK_LISTA_CERT_VENCIDOS.LISTAR(" & vlFechaTermino & ")", vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PP_Rpt_CertVencidos.rpt"), ".RPT", ".TTX"), 1)
    
        
    If objRep.CargaReporte(strRpt & "", "PP_Rpt_CertVencidos.rpt", "Informe Certificados de Supervivencia Vencidos", rs, True, _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("NombreSistema", vgNombreSistema), _
                            ArrFormulas("NombreSubSistema", vgNombreSubSistema), _
                            ArrFormulas("fecha", Txt_Hasta.Text)) = False Then
                            
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
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

Private Sub Cmd_Limpiar_Click()
On Error GoTo Err_Limpiar

    Txt_Hasta = ""
    Txt_Hasta.SetFocus

Exit Sub
Err_Limpiar:
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
Me.Top = 0
Me.Left = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)


Set rs = New ADODB.Recordset
If rs.State = 1 Then
        rs.Close
        Set rs = Nothing
End If

End Sub


Private Sub Txt_Hasta_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Cmd_Imprimir.SetFocus
End If
End Sub

VERSION 5.00
Begin VB.Form Frm_CauSiniestro 
   Caption         =   "Causa de fallecimiento"
   ClientHeight    =   2235
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   ScaleHeight     =   2235
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Indique la causa de fallecimiento"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   5400
         TabIndex        =   2
         Top             =   960
         Width           =   1455
      End
      Begin VB.ComboBox cmbCausafac 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Text            =   "Seleccionar"
         Top             =   480
         Width           =   6735
      End
   End
End
Attribute VB_Name = "Frm_CauSiniestro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public pNUM_POLIZA As String
Public pNUM_ENDOSO As Integer
Public pCOD_CAUENDOSO As String
Public pCOD_CAUSINIESTRO As String

Private Sub cmdAceptar_Click()
  Dim sCodCausa As String
  
  sCodCausa = Mid(Me.cmbCausafac.Text, 1, 4)
  If sCodCausa = "Sele" Then
    sCodCausa = ""
  End If
  Call Frm_EndosoPol.RecibeCausaFac(sCodCausa)
  Unload Me
     
End Sub

Private Sub Form_Load()
    Listar_causas
    
    If pCOD_CAUSINIESTRO <> "" Then
    
      Call fgBuscarPosicionCodigoCombo(pCOD_CAUSINIESTRO, Me.cmbCausafac)
  
    
    End If
    

End Sub

Private Sub Listar_causas()
    Dim Mensaje As String
           
       Dim conn    As ADODB.Connection
       Set conn = New ADODB.Connection
       Set rs = New ADODB.Recordset ' CreateObject("ADODB.Recordset")
       Set objCmd = New ADODB.Command ' CreateObject("ADODB.Command")
       
       conn.Provider = "OraOLEDB.Oracle"
       conn.ConnectionString = "PLSQLRSet=1;Data Source=" & vgNombreBaseDatos & ";" & "User ID=" & vgNombreUsuario & ";Password=" & vgPassWord & ";"
       conn.CursorLocation = adUseClient
       conn.Open
       
       Set objCmd = CreateObject("ADODB.Command")
       Set objCmd.ActiveConnection = conn
       
       objCmd.CommandText = "PKG_CauSiniestro.Listar_causas_siniestro"
       objCmd.CommandType = adCmdStoredProc
  
       Set rs = objCmd.Execute
       
       
       Do While Not rs.EOF
           cmbCausafac.AddItem ((Trim(rs!Codigo) & " - " & Trim(rs!causa)))
           
           rs.MoveNext
           
       Loop

  conn.Close
  Set objCmd = Nothing
  Set rs = Nothing
  Set conn = Nothing

End Sub




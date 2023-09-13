VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Frm_Circular18 
   Caption         =   "Circular 18"
   ClientHeight    =   4590
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   Picture         =   "Frm_Circular18.frx":0000
   ScaleHeight     =   4590
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   855
         Left            =   3840
         Picture         =   "Frm_Circular18.frx":0ECA
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txtTipoCambio 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin MSComctlLib.ProgressBar BarraProgreso 
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   1440
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton CmdIniciarProceso 
         Caption         =   "Generar"
         Height          =   855
         Left            =   1920
         Picture         =   "Frm_Circular18.frx":130C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2880
         Width           =   1095
      End
      Begin VB.ComboBox cmbPeriodo 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   360
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         Height          =   1815
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   6615
         Begin VB.Label lblcuantos 
            Caption         =   "Label2"
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   1200
            Width           =   5295
         End
         Begin VB.Label LblAnexo 
            Caption         =   "Label2"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Label lblTipoCambio 
         Caption         =   "Tipo de Cambio:"
         Height          =   375
         Left            =   4440
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Periodo"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "Frm_Circular18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CmdIniciarProceso_Click()
     Dim conn As ADODB.Connection
     Dim RSresult As ADODB.Recordset
     Dim pNumError As Integer
     Dim pMsjError As String
     Dim vArchivoFinal As String
     
     Dim xlapp As Excel.Application
     Dim Obj_Excel As Object
     Dim Obj_Libro As Object
  On Error GoTo ErrorHandler
  
     If Len(Trim(Me.txtTipoCambio.Text)) = 0 Then
        MsgBox "Debe ingresar el tipo de cambio.", vbExclamation, "Circular 18"
        Me.txtTipoCambio.SetFocus
        Exit Sub
        
     End If
         
    Set Obj_Excel = CreateObject("Excel.Application")
 
    Dim PlantillaCircular As String
    Set conn = New ADODB.Connection
        
    conn.Provider = "OraOLEDB.Oracle"
    conn.ConnectionString = "PLSQLRSet=1;Data Source=" & vgNombreBaseDatos & ";" & "User ID=" & vgNombreUsuario & ";Password=" & vgPassWord & ";"
    conn.CursorLocation = adUseClient
   
   
   ' vArchivoFinal = App.Path & "\Circular18_" & Me.cmbPeriodo.Text & ".xls"
    vArchivoFinal = "C:\Circular18_" & Me.cmbPeriodo.Text & ".xls"
    
    FileCopy strRpt & "Circular_018.xlsx", vArchivoFinal
    Set Obj_Libro = Obj_Excel.Workbooks.Open(vArchivoFinal)
    
    conn.Open
     
    '*******Integra Anexo10**************
    'AFP INTEGRA -> ANEXO10
    Set RSresult = Nothing
    Me.LblAnexo = "PROCESANDO ANEXO10"
    Call GetData(conn, "1", "242", Me.cmbPeriodo.Text, Me.txtTipoCambio.Text, pNumError, pMsjError, RSresult)
    GeneraXlsModalidades vArchivoFinal, "ANEXO10", Obj_Libro, RSresult

    '*******Integra Anexo11**************
    'AFP PROFUTURO -> ANEXO 13
    Set RSresult = Nothing
    Me.LblAnexo = "PROCESANDO ANEXO13"
    Call GetData(conn, "1", "243", Me.cmbPeriodo.Text, Me.txtTipoCambio.Text, pNumError, pMsjError, RSresult)
    GeneraXlsModalidades vArchivoFinal, "ANEXO13", Obj_Libro, RSresult

    '*******Integra Anexo12**************
    'AFP PRIMA -> ANEXO 16
    Set RSresult = Nothing
    Me.LblAnexo = "PROCESANDO ANEXO16"
    Call GetData(conn, "1", "245", Me.cmbPeriodo.Text, Me.txtTipoCambio.Text, pNumError, pMsjError, RSresult)
    GeneraXlsModalidades vArchivoFinal, "ANEXO16", Obj_Libro, RSresult

    '*******Integra Anexo12**************
    'AFP HABITAT -> ANEXO 19
    Set RSresult = Nothing
    Me.LblAnexo = "PROCESANDO ANEXO19"
    Call GetData(conn, "1", "244", Me.cmbPeriodo.Text, Me.txtTipoCambio.Text, pNumError, pMsjError, RSresult)
    GeneraXlsModalidades vArchivoFinal, "ANEXO19", Obj_Libro, RSresult

    
    '*******INTEGRA -> ANEXO 11 *******
     'AFP INTEGRA -> ANEXO11
    Set RSresult = Nothing
    Me.LblAnexo = "PROCESANDO ANEXO11"
    Call GetData(conn, "2", "242", Me.cmbPeriodo.Text, Me.txtTipoCambio.Text, pNumError, pMsjError, RSresult)
    GeneraXlsPensionJubil vArchivoFinal, "ANEXO11", Obj_Libro, RSresult
    
   '********PROFUTURO -> ANEXO 14 ******************
    Set RSresult = Nothing
    Me.LblAnexo = "PROCESANDO ANEXO14"
    Call GetData(conn, "2", "243", Me.cmbPeriodo.Text, Me.txtTipoCambio.Text, pNumError, pMsjError, RSresult)
    GeneraXlsPensionJubil vArchivoFinal, "ANEXO14", Obj_Libro, RSresult
    
     
    '******PRIMA -> ANEXO 17 ******************
     Set RSresult = Nothing
    Me.LblAnexo = "PROCESANDO ANEXO17"
    Call GetData(conn, "2", "245", Me.cmbPeriodo.Text, Me.txtTipoCambio.Text, pNumError, pMsjError, RSresult)
    GeneraXlsPensionJubil vArchivoFinal, "ANEXO17", Obj_Libro, RSresult
    
   '*********HABITAT -> ANEXO 20 ******************
    Set RSresult = Nothing
    Me.LblAnexo = "PROCESANDO ANEXO20"
    Call GetData(conn, "2", "244", Me.cmbPeriodo.Text, Me.txtTipoCambio.Text, pNumError, pMsjError, RSresult)
    GeneraXlsPensionJubil vArchivoFinal, "ANEXO20", Obj_Libro, RSresult
    
    '******************AFP INTEGRA  -> ANEXO 12 ******************
    Set RSresult = Nothing
    Me.LblAnexo = "PROCESANDO ANEXO12"
    Call GetData(conn, "3", "242", Me.cmbPeriodo.Text, Me.txtTipoCambio.Text, pNumError, pMsjError, RSresult)
    GeneraXlsSobrevivencia vArchivoFinal, "ANEXO12", Obj_Libro, RSresult
    
    '******************PROFUTURO -> ANEXO 15 ******************
    Set RSresult = Nothing
    Me.LblAnexo = "PROCESANDO ANEXO15"
    Call GetData(conn, "3", "243", Me.cmbPeriodo.Text, Me.txtTipoCambio.Text, pNumError, pMsjError, RSresult)
    GeneraXlsSobrevivencia vArchivoFinal, "ANEXO15", Obj_Libro, RSresult
    
    '******************PRIMA -> ANEXO 18 ******************
    Set RSresult = Nothing
    Me.LblAnexo = "PROCESANDO ANEXO18"
    Call GetData(conn, "3", "245", Me.cmbPeriodo.Text, Me.txtTipoCambio.Text, pNumError, pMsjError, RSresult)
    GeneraXlsSobrevivencia vArchivoFinal, "ANEXO18", Obj_Libro, RSresult
    
    
    '******************HABITAT -> ANEXO 21 ******************
    Set RSresult = Nothing
    Me.LblAnexo = "PROCESANDO ANEXO21"
    Call GetData(conn, "3", "244", Me.cmbPeriodo.Text, Me.txtTipoCambio.Text, pNumError, pMsjError, RSresult)
    GeneraXlsSobrevivencia vArchivoFinal, "ANEXO21", Obj_Libro, RSresult
    
  
    
    RSresult.Close
    Set RSresult = Nothing
    conn.Close
    
    Obj_Excel.Visible = True 'para ver vista previa
    Obj_Excel.WindowState = xlMaximized ' minimiza excel
    
    
    
    Me.LblAnexo = ""
    Me.lblcuantos = ""
    Me.BarraProgreso.Value = 0
    
    Set Obj_Excel = Nothing
    Exit Sub
ErrorHandler:

   MsgBox Err.Number & "-" & Err.Description & Chr(13) & "Verifique el archivo Excel no esté en uso.", vbCritical, "Circular 18"
   
   
   
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
            
        
        Dim AnnioActual As Integer
        Dim MesActual As Integer
        Dim ItemAnioMes As String
        Dim ItemInicio As String
            
       Me.LblAnexo = ""
       Me.lblcuantos = ""
       
   
        ItemInicio = "200801"
        
        AnnioActual = Format(Now, "YYYY")
        MesActual = Format(Now, "MM")
        
        ItemAnioMes = AnnioActual & Right(String(2, "0") & MesActual, 2)
        
        
        
        Dim i As Long
        
        For i = ItemInicio To ItemAnioMes
        
            cmbPeriodo.AddItem (i)
            
            If Right(i, 2) = 12 Then
                i = i + 88
            End If
        
        Next
        
        If Mid(i, 6, 2) = "00" Then
            cmbPeriodo.Text = i
        Else
        
            cmbPeriodo.Text = i - 1
        End If
        
        

End Sub

Private Sub GetData(ByRef pConn As ADODB.Connection, _
                 ByVal pTipoReporte As String, _
                 ByVal pCodAFP As String, _
                 ByVal pPeriodo As String, _
                 ByVal pTipoCambio As Double, _
                 ByRef pNumError As Integer, _
                 ByRef pMensajeError As String, _
                 ByRef RS As ADODB.Recordset)

   
     
                Dim objCmd As ADODB.Command
                
                Dim param1 As ADODB.Parameter
                Dim param2 As ADODB.Parameter
                Dim param3 As ADODB.Parameter
                Dim param4 As ADODB.Parameter
                Dim param5 As ADODB.Parameter
                
                Dim pPackage As String
                
                If pTipoReporte = "1" Then
                    pPackage = "PKG_CIRCULAR18SEACSA.PensionesPorModalidades"
                Else
                 If pTipoReporte = "2" Then
                     pPackage = "PKG_CIRCULAR18SEACSA.PensionDeJubilacion"
                 Else
                    If pTipoReporte = "3" Then
                         pPackage = "PKG_CIRCULAR18SEACSA.PensionesDeSobrevivencia"
                    End If
                 End If
                End If
    
                

                Set RS = New ADODB.Recordset
                Set objCmd = New ADODB.Command
                

                                      
                Set objCmd = CreateObject("ADODB.Command")
                Set objCmd.ActiveConnection = pConn
                
                objCmd.CommandText = pPackage
                objCmd.CommandType = adCmdStoredProc
          
                Set param1 = objCmd.CreateParameter("pCod_AFP", adVarChar, adParamInput, 15, pCodAFP)
                objCmd.Parameters.Append param1
                
                Set param2 = objCmd.CreateParameter("pPeriodo", adVarChar, adParamInput, 10, pPeriodo)
                objCmd.Parameters.Append param2
                
                If pTipoReporte <> "1" Then
                    Set param3 = objCmd.CreateParameter("pTipoCambio", adDecimal, adParamInput, , pTipoCambio)
                    objCmd.Parameters.Append param3
               End If
                 
                Set param4 = objCmd.CreateParameter("p_outNumError", adInteger, adParamOutput, 1000)
                objCmd.Parameters.Append param4
                
        
                Set param5 = objCmd.CreateParameter("p_outMsgError", adVarChar, adParamOutput, 1000)
                objCmd.Parameters.Append param5
        
                Set RS = objCmd.Execute
                
                pNumError = objCmd.Parameters("p_outNumError")
                pMensajeError = objCmd.Parameters("p_outMsgError")
                
                        

        
        Exit Sub
        

 
   

End Sub

Private Sub GeneraXlsModalidades(ByVal sArchivoFinal As String, _
                                            ByVal sAnexo As String, _
                                            ByVal Obj_Libro As Object, _
                                            ByRef RS As ADODB.Recordset)
    
    Dim vlArchivo As String
    Dim iFila As Long
    Dim iCol As Integer
    Dim vTotaCampos As Integer
    
    
    Dim Obj_Hoja As Object
    vlArchivo = sArchivoFinal

    iFila = 2
    
    vTotaCampos = RS.Fields.Count + 6
    
    
    Set Obj_Hoja = Obj_Libro.Worksheets(sAnexo)
    BarraProgreso.Value = 0
    BarraProgreso.Max = RS.RecordCount
    
    
    
     Do While Not RS.EOF
               iFila = iFila + 1
              For iCol = 6 To vTotaCampos - 2
                Obj_Hoja.Cells(iFila, iCol) = RS.Fields(iCol - 5).Value
              Next
    
            BarraProgreso.Value = BarraProgreso.Value + 1
            Me.lblcuantos = "Procesando " & BarraProgreso.Value & " de " & BarraProgreso.Max & " Registros."
            Me.Refresh
            
        RS.MoveNext
        
      Loop
    

  

End Sub


Private Sub GeneraXlsPensionJubil(ByVal sArchivoFinal As String, _
                                            ByVal sAnexo As String, _
                                            ByVal Obj_Libro As Object, _
                                            ByRef RS As ADODB.Recordset)
                                            
                                            
    Dim vlArchivo As String
    Dim iFila As Long
    Dim iCol As Integer
    Dim vTotaCampos As Integer
    Dim ArSubtotal(12) As Double
    Dim ArTotalGeneral(12) As Double
    
    Dim i As Integer
    Dim FilaSubTot As Integer
    
    
    
    
    Dim Obj_Hoja As Object
    vlArchivo = sArchivoFinal

    iFila = 3
    vTotaCampos = RS.Fields.Count + 3
    FilaSubTot = 4
    
    Set Obj_Hoja = Obj_Libro.Worksheets(sAnexo)
    BarraProgreso.Value = 0
    BarraProgreso.Max = RS.RecordCount
    
    
    
     Do While Not RS.EOF
            
            iFila = iFila + 1
            i = 1
               
            If iFila > 4 And iFila - FilaSubTot <> 3 Then
                
                For iCol = 3 To vTotaCampos - 2
                 Obj_Hoja.Cells(iFila, iCol) = RS.Fields(i).Value
                  ArSubtotal(i) = ArSubtotal(i) + RS.Fields(i).Value
                  i = i + 1
                Next
            
            Else
                FilaSubTot = iFila
                If FilaSubTot > 4 Then
                    i = 1
                    For iCol = 3 To vTotaCampos - 2
                    
                       Obj_Hoja.Cells(iFila - 3, iCol) = ArSubtotal(i)
                       ArTotalGeneral(i) = ArTotalGeneral(i) + ArSubtotal(i)
                       
                       ArSubtotal(i) = 0
                       i = i + 1
                    Next
                End If
        
            End If
    
            BarraProgreso.Value = BarraProgreso.Value + 1
            Me.lblcuantos = "Procesando " & BarraProgreso.Value & " de " & BarraProgreso.Max & " Registros."
            Me.Refresh
            
        RS.MoveNext
        
     Loop
                                              
      For i = 1 To 11
        Obj_Hoja.Cells(40, i + 2) = ArTotalGeneral(i)
      Next
      
      Obj_Hoja.Cells(41, 14) = Me.txtTipoCambio.Text

End Sub

Private Sub GeneraXlsSobrevivencia(ByVal sArchivoFinal As String, _
                                            ByVal sAnexo As String, _
                                            ByVal Obj_Libro As Object, _
                                            ByRef RS As ADODB.Recordset)
                                            
                                            
    Dim vlArchivo As String
    Dim iFila As Long
    Dim iCol As Integer
    Dim vTotaCampos As Integer
    
    Dim ArTotalGeneral(17) As Double
    
    Dim i As Integer
    Dim FilaSubTot As Integer
   
    Dim Obj_Hoja As Object
    vlArchivo = sArchivoFinal

    iFila = 3
    vTotaCampos = RS.Fields.Count + 2
    FilaSubTot = 13
    
    Set Obj_Hoja = Obj_Libro.Worksheets(sAnexo)
    BarraProgreso.Value = 0
    BarraProgreso.Max = RS.RecordCount
    
    
    
     Do While Not RS.EOF
            
            iFila = iFila + 1
            i = 1
               
         
                
                For iCol = 3 To vTotaCampos - 1
                  Obj_Hoja.Cells(iFila, iCol) = RS.Fields(i).Value
                  ArTotalGeneral(i) = ArTotalGeneral(i) + RS.Fields(i).Value
                  
                  i = i + 1
                Next
        
    
            BarraProgreso.Value = BarraProgreso.Value + 1
            Me.lblcuantos = "Procesando " & BarraProgreso.Value & " de " & BarraProgreso.Max & " Registros."
            Me.Refresh
            
        RS.MoveNext
        
     Loop
                                              
      For i = 1 To 16
        Obj_Hoja.Cells(13, i + 2) = ArTotalGeneral(i)
        Obj_Hoja.Cells(17, i + 2) = ArTotalGeneral(i)
      Next
      
      

End Sub
Private Sub txtTipoCambio_KeyPress(KeyAscii As Integer)
Dim KeyChar As String
KeyChar = Chr(KeyAscii)

 If KeyAscii > 31 Then
 
   If Not IsNumeric(KeyChar) Then
     If KeyChar <> "." Then
         KeyAscii = 0
     End If
   End If
   
End If


End Sub

Private Sub txtTipoCambio_LostFocus()

        Dim oReg As RegExp
        Set oReg = New RegExp
        
     
        oReg.Pattern = "^\d{1}\.\d{3}?$"

        If oReg.Test(Me.txtTipoCambio.Text) = False Then
            MsgBox "El tipo de cambio debe tener 1 entero y 3 decimales.", vbExclamation, "Tipo de Cambio"
            Me.txtTipoCambio.SetFocus
                   
        End If
        
        
        Set oReg = Nothing
        
    

End Sub


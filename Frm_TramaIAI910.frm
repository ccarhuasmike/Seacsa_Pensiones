VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_TramaIAI910 
   Caption         =   "Trama IAI9 y IAI10"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   5895
      Begin VB.CommandButton cmd_salir 
         Caption         =   "Salir"
         Height          =   855
         Left            =   4080
         Picture         =   "Frm_TramaIAI910.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2160
         Width           =   855
      End
      Begin VB.ComboBox cmbTipoCarga 
         Height          =   315
         Left            =   3480
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   960
         Width           =   2055
      End
      Begin MSComDlg.CommonDialog VentSelecFichero 
         Left            =   2400
         Top             =   3120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ProgressBar BarraProgreso 
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton CmdTrama 
         Caption         =   "Trama"
         Height          =   855
         Left            =   2520
         Picture         =   "Frm_TramaIAI910.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2160
         Width           =   855
      End
      Begin VB.CommandButton cmdExcel 
         Caption         =   "Excel"
         Height          =   855
         Left            =   840
         Picture         =   "Frm_TramaIAI910.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2160
         Width           =   855
      End
      Begin VB.OptionButton OptTrama 
         Caption         =   "IAI10"
         Height          =   375
         Index           =   1
         Left            =   1920
         TabIndex        =   8
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton OptTrama 
         Caption         =   "IAI9"
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   7
         Top             =   960
         Value           =   -1  'True
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtFechaHasta 
         Height          =   375
         Left            =   3480
         TabIndex        =   5
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   83689473
         CurrentDate     =   44748
      End
      Begin MSComCtl2.DTPicker dtFechaDesde 
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   83689473
         CurrentDate     =   44748
      End
      Begin VB.Label lblprogreso 
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1920
         Width           =   4095
      End
      Begin VB.Label Label4 
         Caption         =   "Trama:"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   1030
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta:"
         Height          =   255
         Left            =   2760
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Desde:"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Trama IAI9 y IAI10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Frm_TramaIAI910"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type PolizaIAI09
    cod_fila As String
    Cod_entidad As String
    Tipo_siniestro As String
    num_solicitud As String
    Fec_solicitud As String
    Cod_afiliado As String
    doc_identidad As String
    Fec_Ocurrencia As String
    Cod_causiniestro As String
    Fec_aceptacion As String
    situacion_expediente As String

End Type


Private Type PolizaIAI10
    cod_fila As String
    cod_protecta As String
    Cod_Cuspp As String
    doc_identidad As String
    tipo_solicitud As String
    Fec_Ocurrencia As String
    num_solicitud As String
    fecha_solicitud As String
    situacion_cobertura As String
    causal_cobertura As String
    modalidad_pension As String
    fecha_endoso As String
    fecha_pago As String
    situacion_expediente As String
 End Type

Dim oPoliza9() As PolizaIAI09
Dim oPoliza10() As PolizaIAI10

Private Sub Lst_ListaPolizasIA09(ByVal pfecha_desde As String, _
                                ByVal pfecha_hasta As String, _
                                ByVal pnum_carga As Integer)

       On Error GoTo GestionError
  
       Dim conn    As ADODB.Connection
       Dim objCmd As ADODB.Command
       Dim rs As ADODB.Recordset
       Dim itotalReg As Double
       Dim i As Double
       Dim Mensaje As String
       
       
       Set conn = New ADODB.Connection
       Set rs = New ADODB.Recordset ' CreateObject("ADODB.Recordset")
       
       Dim param1 As ADODB.Parameter
       Dim param2 As ADODB.Parameter
       Dim param3 As ADODB.Parameter
       Dim param4 As ADODB.Parameter
       Dim param5 As ADODB.Parameter
    
       Set objCmd = New ADODB.Command ' CreateObject("ADODB.Command")

       conn.Provider = "OraOLEDB.Oracle"
       conn.ConnectionString = "PLSQLRSet=1;Data Source=" & vgNombreBaseDatos & ";" & "User ID=" & vgNombreUsuario & ";Password=" & vgPassWord & ";"
       conn.CursorLocation = adUseClient
       conn.Open
    
       Set objCmd = CreateObject("ADODB.Command")
       Set objCmd.ActiveConnection = conn
       
       objCmd.CommandText = "PKG_INFORMES_IAI.Lst_IAI09"
       objCmd.CommandType = adCmdStoredProc
       
        Set param1 = objCmd.CreateParameter("pfecha_desde", adVarChar, adParamInput, 8, pfecha_desde)
        objCmd.Parameters.Append param1
                
        Set param2 = objCmd.CreateParameter("pfecha_hasta", adVarChar, adParamInput, 8, pfecha_hasta)
        objCmd.Parameters.Append param2
             
        Set param3 = objCmd.CreateParameter("ptipocarga", adInteger, adParamInput)
        param3.Value = pnum_carga
        objCmd.Parameters.Append param3
             
        Set param4 = objCmd.CreateParameter("p_outNumError", adDouble, adParamOutput)
        objCmd.Parameters.Append param4
        
        Set param5 = objCmd.CreateParameter("p_outMsgError", adVarChar, adParamOutput, 200)
        objCmd.Parameters.Append param5
           
        Set rs = objCmd.Execute
        
        If Not IsNull(objCmd.Parameters.Item("p_outMsgError").Value) Then
            Mensaje = objCmd.Parameters.Item("p_outMsgError").Value
            Err.Raise -2222, "Ins_Causa", Mensaje
        Else
          Mensaje = ""
    
        If rs.RecordCount > 0 Then
          ReDim oPoliza9(rs.RecordCount - 1)
        End If
    
        
          itotalReg = rs.RecordCount
          i = -1
          
    
          While Not rs.EOF
            i = i + 1
            
            oPoliza9(i).cod_fila = IIf(IsNull(rs!cod_fila), "", rs!cod_fila)
            oPoliza9(i).Cod_entidad = IIf(IsNull(rs!entidad), "", rs!entidad)
            oPoliza9(i).Tipo_siniestro = IIf(IsNull(rs!tipo_solicitud), "", rs!tipo_solicitud)
            oPoliza9(i).num_solicitud = IIf(IsNull(rs!num_solicitud), "", rs!num_solicitud)
            oPoliza9(i).Fec_solicitud = IIf(IsNull(rs!fecha_solicitud), "", rs!fecha_solicitud)
            oPoliza9(i).Cod_afiliado = IIf(IsNull(rs!Cod_afiliado), "", rs!Cod_afiliado)
            oPoliza9(i).doc_identidad = IIf(IsNull(rs!doc_identidad), "", rs!doc_identidad)
            oPoliza9(i).Fec_Ocurrencia = IIf(IsNull(rs!fecha_ocurrencia), "", rs!fecha_ocurrencia)
            oPoliza9(i).Cod_causiniestro = IIf(IsNull(rs!causal), "", rs!causal)
            oPoliza9(i).Fec_aceptacion = IIf(IsNull(rs!fecha_aceptacion), "", rs!fecha_aceptacion)
            oPoliza9(i).situacion_expediente = IIf(IsNull(rs!situacion_expediente), "", rs!situacion_expediente)
               
            rs.MoveNext
            
         Wend
   
       
      End If
   
        conn.Close
        Set objCmd = Nothing
        Set rs = Nothing
        Set conn = Nothing
        
        Exit Sub
        
  
GestionError:
    
    MsgBox "Se ha producido un error. Tipo de error = " & Err.Number & ". Descripción: " & Err.Description, vbCritical
    

End Sub


Private Sub Lst_ListaPolizasIA10(ByVal pfecha_desde As String, ByVal pfecha_hasta As String, ByVal ptipoCarga As Integer)

       On Error GoTo GestionError
  
       Dim conn    As ADODB.Connection
       Dim objCmd As ADODB.Command
       Dim rs As ADODB.Recordset
       Dim itotalReg As Double
       Dim i As Double
       Dim Mensaje As String
       
       
       Set conn = New ADODB.Connection
       Set rs = New ADODB.Recordset ' CreateObject("ADODB.Recordset")
       
       Dim param1 As ADODB.Parameter
       Dim param2 As ADODB.Parameter
       Dim param3 As ADODB.Parameter
       Dim param4 As ADODB.Parameter
       Dim param5 As ADODB.Parameter
    
       Set objCmd = New ADODB.Command ' CreateObject("ADODB.Command")

       conn.Provider = "OraOLEDB.Oracle"
       conn.ConnectionString = "PLSQLRSet=1;Data Source=" & vgNombreBaseDatos & ";" & "User ID=" & vgNombreUsuario & ";Password=" & vgPassWord & ";"
       conn.CursorLocation = adUseClient
       conn.Open
    
       Set objCmd = CreateObject("ADODB.Command")
       Set objCmd.ActiveConnection = conn
       
       objCmd.CommandText = "PKG_INFORMES_IAI.Lst_IAI10"
       objCmd.CommandType = adCmdStoredProc
       
        Set param1 = objCmd.CreateParameter("pfecha_desde", adVarChar, adParamInput, 8, pfecha_desde)
        objCmd.Parameters.Append param1
                
        Set param2 = objCmd.CreateParameter("pfecha_hasta", adVarChar, adParamInput, 8, pfecha_hasta)
        objCmd.Parameters.Append param2
        
        Set param3 = objCmd.CreateParameter("ptipocarga", adInteger, adParamInput)
        param3.Value = ptipoCarga
        objCmd.Parameters.Append param3
    
        Set param4 = objCmd.CreateParameter("p_outNumError", adDouble, adParamOutput)
        objCmd.Parameters.Append param4
        
        Set param5 = objCmd.CreateParameter("p_outMsgError", adVarChar, adParamOutput, 200)
        objCmd.Parameters.Append param5
           
        Set rs = objCmd.Execute
        
        If Not IsNull(objCmd.Parameters.Item("p_outMsgError").Value) Then
            Mensaje = objCmd.Parameters.Item("p_outMsgError").Value
            Err.Raise -2222, "Ins_Causa", Mensaje
        Else
          Mensaje = ""
     
        If rs.RecordCount > 0 Then
          ReDim oPoliza10(rs.RecordCount - 1)
        End If
     
        
          itotalReg = rs.RecordCount
          i = -1
          
    
          While Not rs.EOF
            i = i + 1
            
            oPoliza10(i).cod_fila = IIf(IsNull(rs!cod_fila), "", rs!cod_fila)
            oPoliza10(i).cod_protecta = IIf(IsNull(rs!entidad), "", rs!entidad)
            oPoliza10(i).Cod_Cuspp = IIf(IsNull(rs!Cod_afiliado), "", rs!Cod_afiliado)
            oPoliza10(i).doc_identidad = IIf(IsNull(rs!doc_identidad), "", rs!doc_identidad)
            oPoliza10(i).tipo_solicitud = IIf(IsNull(rs!tipo_solicitud), "", rs!tipo_solicitud)
            oPoliza10(i).num_solicitud = IIf(IsNull(rs!num_solicitud), "", rs!num_solicitud)
            oPoliza10(i).fecha_solicitud = IIf(IsNull(rs!fecha_solicitud), "", rs!fecha_solicitud)
            oPoliza10(i).situacion_cobertura = IIf(IsNull(rs!situacion_cobertura), "", rs!situacion_cobertura)
            oPoliza10(i).causal_cobertura = IIf(IsNull(rs!causal_cobertura), "", rs!causal_cobertura)
            oPoliza10(i).modalidad_pension = IIf(IsNull(rs!mod_pension), "", rs!mod_pension)
            oPoliza10(i).fecha_endoso = IIf(IsNull(rs!Fec_aceptacion), "", rs!Fec_aceptacion)
            oPoliza10(i).fecha_pago = IIf(IsNull(rs!fecha_pago), "", rs!fecha_pago)
            oPoliza10(i).situacion_expediente = IIf(IsNull(rs!situacion_expediente), "", rs!situacion_expediente)
               
            rs.MoveNext
            
         Wend
         

   
       
      End If
   
        conn.Close
        Set objCmd = Nothing
        Set rs = Nothing
        Set conn = Nothing
        
        Exit Sub
        
  
GestionError:
    
    MsgBox "Se ha producido un error. Tipo de error = " & Err.Number & ". Descripción: " & Err.Description, vbCritical
    

End Sub
Private Sub IA09(ByVal sDesde As String, ByVal sHasta As String, ByVal ptipoCarga As Integer)

    Lst_ListaPolizasIA09 sDesde, sHasta, ptipoCarga
  
        Dim Obj_Excel As Object
        Dim Obj_Libro As Object
        Dim Obj_Hoja As Object
        Dim vFila As Integer
        Dim vColumna As Integer
        Dim i As Integer
        Dim vTotaReg As Integer
        
       If ((Not oPoliza9) = -1) Then
            
            MsgBox "No hay registros para mostrar", vbExclamation, "Trama IA09"
            Exit Sub
  
        End If
        
  
        
      Set Obj_Excel = CreateObject("Excel.Application")

      Set Obj_Libro = Obj_Excel.Workbooks.Add
      Set Obj_Hoja = Obj_Libro.Worksheets.Add
  
        vFila = 0
        vColumna = 0
        
      lblprogreso.Caption = "Escribiendo filas, espere un momento..."
 
    Dim Titulos(11) As String
    
  
    Titulos(0) = "FILA"
    Titulos(1) = "ENTIDAD"
    Titulos(2) = "TIPO SOLICITUD"
    Titulos(3) = "NUM SOLICITUD"
    Titulos(4) = "FECHA SOLICITUD"
    Titulos(5) = "COD AFILIADO"
    Titulos(6) = "DOC IDENTIDAD"
    Titulos(7) = "FECHA OCURRENCIA"
    Titulos(8) = "CAUSAL"
    Titulos(9) = "FECHA ACEPTACION"
    Titulos(10) = "SITUACION EXPEDIENTE"
    Titulos(11) = "CONCATENACION"
    
    

    vFila = 1
    Obj_Hoja.Cells(vFila, 1) = "Informe SBS - IAI09"
    Obj_Hoja.rows(vFila).Font.Bold = True
    vFila = 2
    Obj_Hoja.Cells(vFila, 1) = "Periodo: " & Me.dtFechaDesde.Value & " - " & Me.dtFechaHasta.Value
    Obj_Hoja.rows(vFila).Font.Bold = True
   
    vFila = 3
  
    For i = 0 To UBound(Titulos)
        Obj_Hoja.Cells(vFila, i + 1) = Titulos(i)
    Next
 
    vTotaReg = UBound(oPoliza9)
  
    BarraProgreso.Value = 0
    
    If vTotaReg > 0 Then
    
                BarraProgreso.Max = vTotaReg
            
                Obj_Hoja.Columns("A:L").NumberFormat = "@"
                 
                For i = 0 To UBound(oPoliza9)
                     vFila = vFila + 1
                     Obj_Hoja.Cells(vFila, 1) = Trim(oPoliza9(i).cod_fila)
                     Obj_Hoja.Cells(vFila, 2) = oPoliza9(i).Cod_entidad
                     Obj_Hoja.Cells(vFila, 3) = oPoliza9(i).Tipo_siniestro
                     Obj_Hoja.Cells(vFila, 4) = oPoliza9(i).num_solicitud
                     Obj_Hoja.Cells(vFila, 5) = oPoliza9(i).Fec_solicitud
                     Obj_Hoja.Cells(vFila, 6) = oPoliza9(i).Cod_afiliado
                     Obj_Hoja.Cells(vFila, 7) = oPoliza9(i).doc_identidad
                     Obj_Hoja.Cells(vFila, 8) = oPoliza9(i).Fec_Ocurrencia
                     Obj_Hoja.Cells(vFila, 9) = oPoliza9(i).Cod_causiniestro
                     Obj_Hoja.Cells(vFila, 10) = oPoliza9(i).Fec_aceptacion
                     Obj_Hoja.Cells(vFila, 11) = oPoliza9(i).situacion_expediente
                     
            
                     
                        Obj_Hoja.Cells(vFila, 12) = Right(Space(6) & Left(oPoliza9(i).cod_fila, 6), 6) & _
                        Right(Space(5) & Left(oPoliza9(i).Cod_entidad, 5), 5) & _
                        Right(Space(1) & Left(oPoliza9(i).Tipo_siniestro, 1), 1) & _
                        Right(Space(10) & Left(oPoliza9(i).num_solicitud, 10), 10) & _
                        Right(Space(8) & Left(oPoliza9(i).Fec_solicitud, 8), 8) & _
                        Right(Space(12) & Left(oPoliza9(i).Cod_afiliado, 12), 12) & _
                        Right(Space(15) & Left(oPoliza9(i).doc_identidad, 15), 15) & _
                        Right(Space(8) & Left(oPoliza9(i).Fec_Ocurrencia, 8), 8) & _
                        Right(Space(4) & Left(oPoliza9(i).Cod_causiniestro, 4), 4) & _
                        Right(Space(8) & Left(oPoliza9(i).Fec_aceptacion, 8), 8) & _
                        Right(Space(2) & Left(oPoliza9(i).situacion_expediente, 2), 2)
                     
                     
                     
                     
                    BarraProgreso.Value = vFila - 4
                    lblprogreso.Caption = "Procesando " & BarraProgreso.Value & " de " & vTotaReg & " Registros."
                    Me.Refresh
            
            
                Next
                
            
                
                 Obj_Hoja.rows(3).Font.Bold = True
                 Obj_Hoja.Range("A3:L3").Borders.LineStyle = xlContinuous
                 Obj_Hoja.Range("A3:L3").Interior.Color = RGB(0, 32, 96)
                 Obj_Hoja.Range("A3:L3").Font.Color = RGB(237, 125, 49)
                 Obj_Hoja.Range("A3:L3").RowHeight = 34.5
                 Obj_Hoja.Range("A3:L3").VerticalAlignment = xlVAlignCenter
                 
                
               Obj_Hoja.Columns("A:L").AutoFit
               Obj_Hoja.Columns("A").ColumnWidth = 5
            
              
               Obj_Hoja.Range("A1:A1").Select
             
               Obj_Excel.Visible = True
                  
            '       Me.lblfinexcel.Caption = Time
            '       Me.lblfinexcel.Caption = Format((CDate(lblinicio.Caption) - CDate(lblfinexcel.Caption)), "hh:mm:ss")
                     
                
                Set Obj_Hoja = Nothing
                Set Obj_Libro = Nothing
                Set Obj_Excel = Nothing
    Else
        MsgBox "No hay datos para mostrar", vbApplicationModal, "Informe IA9 IA10"
        
                
    End If
    
    
                
    MousePointer = vbDefault
    
    BarraProgreso.Value = 0
    lblprogreso.Caption = ""
    
Gestionaerror:
'If Err.Number <> 0 Then
'    GestiónError
'    Resume Next
'End If

End Sub
Private Sub IA10(ByVal sDesde As String, ByVal sHasta As String, ByVal ptipoCarga As Integer)

    Lst_ListaPolizasIA10 sDesde, sHasta, ptipoCarga


        Dim Obj_Excel As Object
        Dim Obj_Libro As Object
        Dim Obj_Hoja As Object
        Dim vFila As Integer
        Dim vColumna As Integer
        Dim i As Integer
        Dim vTotaReg As Integer
        
        
        If ((Not oPoliza10) = -1) Then
             
             MsgBox "No hay registros para mostrar", vbExclamation, "Trama IA10"
            Exit Sub
  
        End If
  
        
      Set Obj_Excel = CreateObject("Excel.Application")

      Set Obj_Libro = Obj_Excel.Workbooks.Add
      Set Obj_Hoja = Obj_Libro.Worksheets.Add
  
        vFila = 0
        vColumna = 0
        
      lblprogreso.Caption = "Escribiendo filas, espere un momento..."
 
    Dim Titulos(30) As String
  
    Titulos(0) = "FILA"
    Titulos(1) = "ENTIDAD"
    Titulos(2) = "COD AFILIADO"
    Titulos(3) = "DOC IDENTIDAD"
    Titulos(4) = "TIPO SOLICITUD"
    Titulos(5) = "NUM SOLICITUD"
    Titulos(6) = "FECHA SOLICITUD"
    Titulos(7) = "SITUACION COBERTURA"
    Titulos(8) = "CAUSAL"
    Titulos(9) = "MODALIDAD DE PENSION"
    Titulos(10) = "FECHA ACEPTACION"
    Titulos(11) = "FECHA PAGO"
    Titulos(12) = "SITUACION EXPEDIENTE"
    Titulos(13) = "CONCATENACION"
 
    vFila = 1
    Obj_Hoja.Cells(vFila, 1) = "Informe SBS - IAI10"
    Obj_Hoja.rows(vFila).Font.Bold = True
    vFila = 2
    Obj_Hoja.Cells(vFila, 1) = "Periodo: " & Me.dtFechaDesde.Value & " - " & Me.dtFechaHasta.Value
    Obj_Hoja.rows(vFila).Font.Bold = True
   
    vFila = 3
    
    For i = 0 To UBound(Titulos)
       Obj_Hoja.Cells(vFila, i + 1) = Titulos(i)
    Next
 
    vTotaReg = UBound(oPoliza10) + 1
    
    If vTotaReg > 0 Then
       
       
        
    BarraProgreso.Value = 0
    BarraProgreso.Max = vTotaReg
    
    
    Obj_Hoja.Columns("A:N").NumberFormat = "@"
     
    For i = 0 To UBound(oPoliza10)
         vFila = vFila + 1
         Obj_Hoja.Cells(vFila, 1) = Trim(oPoliza10(i).cod_fila)
         Obj_Hoja.Cells(vFila, 2) = oPoliza10(i).cod_protecta
         Obj_Hoja.Cells(vFila, 3) = oPoliza10(i).Cod_Cuspp
         Obj_Hoja.Cells(vFila, 4) = oPoliza10(i).doc_identidad
         Obj_Hoja.Cells(vFila, 5) = oPoliza10(i).tipo_solicitud
         Obj_Hoja.Cells(vFila, 6) = oPoliza10(i).num_solicitud
         Obj_Hoja.Cells(vFila, 7) = oPoliza10(i).fecha_solicitud
         Obj_Hoja.Cells(vFila, 8) = oPoliza10(i).situacion_cobertura
         Obj_Hoja.Cells(vFila, 9) = oPoliza10(i).causal_cobertura
         Obj_Hoja.Cells(vFila, 10) = oPoliza10(i).modalidad_pension
         Obj_Hoja.Cells(vFila, 11) = oPoliza10(i).fecha_endoso
         Obj_Hoja.Cells(vFila, 12) = oPoliza10(i).fecha_pago
         Obj_Hoja.Cells(vFila, 13) = oPoliza10(i).situacion_expediente
         
      
 Obj_Hoja.Cells(vFila, 14) = Right(Space(6) & Left(oPoliza10(i).cod_fila, 6), 6) & _
            Right(Space(5) & Left(oPoliza10(i).cod_protecta, 5), 5) & _
            Right(Space(12) & Left(oPoliza10(i).Cod_Cuspp, 12), 12) & _
            Right(Space(15) & Left(oPoliza10(i).doc_identidad, 15), 15) & _
            Right(Space(1) & Left(oPoliza10(i).tipo_solicitud, 1), 1) & _
            Right(Space(10) & Left(oPoliza10(i).num_solicitud, 10), 10) & _
            Right(Space(8) & Left(oPoliza10(i).fecha_solicitud, 8), 8) & _
            Right(Space(1) & Left(oPoliza10(i).situacion_cobertura, 1), 1) & _
            Right(Space(2) & Left(oPoliza10(i).causal_cobertura, 2), 2) & _
            Right(Space(1) & Left(oPoliza10(i).modalidad_pension, 1), 1) & _
            Right(Space(8) & Left(oPoliza10(i).fecha_endoso, 8), 8) & _
            Right(Space(8) & Left(oPoliza10(i).fecha_pago, 8), 8) & _
            Right(Space(2) & Left(oPoliza10(i).situacion_expediente, 2), 2)

        BarraProgreso.Value = vFila - 4
        lblprogreso.Caption = "Procesando " & BarraProgreso.Value & " de " & vTotaReg & " Registros."
        Me.Refresh


    Next
   
    
     Obj_Hoja.rows(3).Font.Bold = True
     Obj_Hoja.Range("A3:N3").Borders.LineStyle = xlContinuous
     Obj_Hoja.Range("A3:N3").Interior.Color = RGB(0, 32, 96)
     Obj_Hoja.Range("A3:N3").Font.Color = RGB(237, 125, 49)
     Obj_Hoja.Range("A3:N3").RowHeight = 34.5
     Obj_Hoja.Range("A3:N3").VerticalAlignment = xlVAlignCenter
     
    
   Obj_Hoja.Columns("A:N").AutoFit
   Obj_Hoja.Columns("A").ColumnWidth = 5

  
   Obj_Hoja.Range("A1:A1").Select
 
   Obj_Excel.Visible = True
      
'       Me.lblfinexcel.Caption = Time
'       Me.lblfinexcel.Caption = Format((CDate(lblinicio.Caption) - CDate(lblfinexcel.Caption)), "hh:mm:ss")
         
    
    Set Obj_Hoja = Nothing
    Set Obj_Libro = Nothing
    Set Obj_Excel = Nothing
    
   Else
        MsgBox "No hay datos para mostrar", vbApplicationModal, "Informe IA9 IA10"
        
                
    End If
    
    
    MousePointer = vbDefault
    
    BarraProgreso.Value = 0
    lblprogreso.Caption = ""
    
Gestionaerror:
'If Err.Number <> 0 Then
'    GestiónError
'    Resume Next
'End If

End Sub

Private Sub cmd_salir_Click()
 Unload Me
 
End Sub

Private Sub cmdExcel_Click()
Dim fec_desde As String
Dim fec_hasta As String

    fec_desde = Format(Me.dtFechaDesde.Value, "yyyymmdd")
    fec_hasta = Format(Me.dtFechaHasta.Value, "yyyymmdd")
    
   Call InactivarControles
    
    Dim v_tipocarga As Integer
    
   v_tipocarga = Me.cmbTipoCarga.ListIndex + 1
    
   If OptTrama(0).Value = True Then
      IA09 fec_desde, fec_hasta, v_tipocarga
   Else
      IA10 fec_desde, fec_hasta, v_tipocarga
   End If
   
   Call ActivarControles
   
    
End Sub
Private Sub Trama09()
    
    Dim fec_desde As String
    Dim fec_hasta As String
    Dim variable As String
    Dim NombreArchivo As String
    Dim PrimeraLinea As String
    Dim FilaPrint As String
    Dim vFechaActual As String
    Dim vFecAcCabecera As String
    
    
    Dim i As Integer
  
    Erase oPoliza9
    
    vFechaActual = Format(Now, "yymmdd")
    NombreArchivo = "01" & vFechaActual & ".095"
    
    On Error GoTo MyErrorHandler
       
    VentSelecFichero.CancelError = True
    VentSelecFichero.DialogTitle = "Ubicación del archivo"
    VentSelecFichero.InitDir = "C:\"
    VentSelecFichero.Flags = cdlOFNHideReadOnly
    VentSelecFichero.FileName = "C:\" & NombreArchivo
    VentSelecFichero.ShowSave
    variable = VentSelecFichero.FileName
    
 
    
    Dim v_tipocarga As Integer
    
    v_tipocarga = Me.cmbTipoCarga.ListIndex + 1
    
    fec_desde = Format(Me.dtFechaDesde.Value, "yyyymmdd")
    fec_hasta = Format(Me.dtFechaHasta.Value, "yyyymmdd")
 
    Lst_ListaPolizasIA09 fec_desde, fec_hasta, v_tipocarga
   
    If ((Not oPoliza9) = -1) Then
        
            MsgBox "No hay registros para mostrar", vbExclamation, "Trama IA09"
            Exit Sub
  
      End If
      
      
    vFecAcCabecera = Format(Now, "yyyymmdd")
    PrimeraLinea = "00950100209" & vFecAcCabecera & "000"
    
    
    Static obj As Object
    If obj Is Nothing Then Set obj = VBA.CreateObject("ADODB.Stream")
    
    
    If obj.State = 1 Then
        obj.Close
    End If
    
    
    obj.Open
    obj.Charset = "windows-1252"
    obj.Type = 2
      
    obj.WriteText PrimeraLinea & vbCrLf
    
      
    BarraProgreso.Value = 0
    BarraProgreso.Max = UBound(oPoliza9) + 1
      
      For i = 0 To UBound(oPoliza9)
      
            FilaPrint = Right(Space(6) & Left(oPoliza9(i).cod_fila, 6), 6) & _
            Right(Space(5) & Left(oPoliza9(i).Cod_entidad, 5), 5) & _
            Right(Space(1) & Left(oPoliza9(i).Tipo_siniestro, 1), 1) & _
            Right(Space(10) & Left(oPoliza9(i).num_solicitud, 10), 10) & _
            Right(Space(8) & Left(oPoliza9(i).Fec_solicitud, 8), 8) & _
            Right(Space(12) & Left(oPoliza9(i).Cod_afiliado, 12), 12) & _
            Right(Space(15) & Left(oPoliza9(i).doc_identidad, 15), 15) & _
            Right(Space(8) & Left(oPoliza9(i).Fec_Ocurrencia, 8), 8) & _
            Right(Space(4) & Left(oPoliza9(i).Cod_causiniestro, 4), 4) & _
            Right(Space(8) & Left(oPoliza9(i).Fec_aceptacion, 8), 8) & _
            Right(Space(2) & Left(oPoliza9(i).situacion_expediente, 2), 2)

        
            obj.WriteText FilaPrint & vbCrLf
       
            BarraProgreso.Value = i + 1
            lblprogreso.Caption = "Procesando " & BarraProgreso.Value & " de " & BarraProgreso.Max & " Registros."
        
             Me.Refresh
    
      Next
      
        
   If Dir(variable, vbArchive) <> "" Then
    Kill (variable)
   End If

 
   obj.SaveToFile variable
   obj.Close
   

   
   MsgBox "Proceso terminado", vbInformation, "Trama 09"
   
MyErrorHandler:
   
   
End Sub

Private Sub Trama10()
    
    Dim fec_desde As String
    Dim fec_hasta As String
    Dim variable As String
    Dim NombreArchivo As String
    Dim PrimeraLinea As String
    Dim FilaPrint As String
    Dim vFecAcCabecera As String
    Dim vFechaActual As String
    
    Dim i As Integer
    
  
    Erase oPoliza10
    
    vFechaActual = Format(Now, "yymmdd")
    NombreArchivo = "02" & vFechaActual & ".095"
    
    fec_desde = Format(Me.dtFechaDesde.Value, "yyyymmdd")
    fec_hasta = Format(Me.dtFechaHasta.Value, "yyyymmdd")
    
    On Error GoTo MyErrorHandler

    VentSelecFichero.CancelError = True
    
    VentSelecFichero.DialogTitle = "Ubicación del archivo"
    VentSelecFichero.InitDir = "C:\"
    VentSelecFichero.Flags = cdlOFNHideReadOnly
    VentSelecFichero.FileName = "C:\" & NombreArchivo
    VentSelecFichero.ShowSave
    variable = VentSelecFichero.FileName
    

       
    
      Dim v_tipocarga As Integer
    
    v_tipocarga = Me.cmbTipoCarga.ListIndex + 1
 
    Lst_ListaPolizasIA10 fec_desde, fec_hasta, v_tipocarga
    
    If ((Not oPoliza10) = -1) Then
    
            MsgBox "No hay registros para mostrar", vbExclamation, "Trama IA10"
            Exit Sub
  
      End If
    
    vFecAcCabecera = Format(Now, "yyyymmdd")
    PrimeraLinea = "00950200209" & vFecAcCabecera & "000"
    
    Static obj As Object
    Dim sSituacion As String
    
    
   
    
    If obj Is Nothing Then Set obj = VBA.CreateObject("ADODB.Stream")
    
    If obj.State = 1 Then
        obj.Close
    End If
    

    obj.Open
    obj.Charset = "windows-1252"
    obj.Type = 2
    
    obj.WriteText PrimeraLinea & vbCrLf
    
    BarraProgreso.Value = 0
    BarraProgreso.Max = UBound(oPoliza10) + 1
      
      For i = 0 To UBound(oPoliza10)
      
      
            FilaPrint = Right(Space(6) & Left(oPoliza10(i).cod_fila, 6), 6) & _
            Right(Space(5) & Left(oPoliza10(i).cod_protecta, 5), 5) & _
            Right(Space(12) & Left(oPoliza10(i).Cod_Cuspp, 12), 12) & _
            Right(Space(15) & Left(oPoliza10(i).doc_identidad, 15), 15) & _
            Right(Space(1) & Left(oPoliza10(i).tipo_solicitud, 1), 1) & _
            Right(Space(10) & Left(oPoliza10(i).num_solicitud, 10), 10) & _
            Right(Space(8) & Left(oPoliza10(i).fecha_solicitud, 8), 8) & _
            Right(Space(1) & Left(oPoliza10(i).situacion_cobertura, 1), 1) & _
            Right(Space(2) & Left(oPoliza10(i).causal_cobertura, 2), 2) & _
            Right(Space(1) & Left(oPoliza10(i).modalidad_pension, 1), 1) & _
            Right(Space(8) & Left(oPoliza10(i).fecha_endoso, 8), 8) & _
            Right(Space(8) & Left(oPoliza10(i).fecha_pago, 8), 8) & _
            Right(Space(2) & Left(oPoliza10(i).situacion_expediente, 2), 2)
            
           obj.WriteText FilaPrint & vbCrLf
 
          
            
            BarraProgreso.Value = i + 1
            lblprogreso.Caption = "Procesando " & BarraProgreso.Value & " de " & BarraProgreso.Max & " Registros."
            Me.Refresh
            
      
      Next
              
        
   If Dir(variable, vbArchive) <> "" Then
     Kill (variable)
   End If
   
   obj.SaveToFile variable
   obj.Close

   MsgBox "Proceso terminado", vbInformation, "Trama 10"
   
MyErrorHandler:
   
End Sub

Private Sub CmdTrama_Click()
      
      
 Call InactivarControles
 
   If OptTrama(0).Value = True Then
    Call Trama09
   Else
    Call Trama10
   End If
   
   Call ActivarControles
   
 
End Sub

Private Sub InactivarControles()

    dtFechaDesde.Enabled = False
    dtFechaHasta.Enabled = False
    OptTrama(0).Enabled = False
    OptTrama(1).Enabled = False
    
    cmdExcel.Enabled = False
    CmdTrama.Enabled = False

End Sub
Private Sub ActivarControles()

    dtFechaDesde.Enabled = True
    dtFechaHasta.Enabled = True
    OptTrama(0).Enabled = True
    OptTrama(1).Enabled = True
    
    cmdExcel.Enabled = True
    CmdTrama.Enabled = True
    
    BarraProgreso.Value = 0
    Me.lblprogreso.Caption = ""
    

End Sub

Private Sub Form_Load()
    dtFechaDesde = Format(Now, "mm/dd/yyyy")
    dtFechaHasta = Format(Now, "mm/dd/yyyy")
    
    Me.cmbTipoCarga.AddItem ("Primera Carga")
    Me.cmbTipoCarga.AddItem ("Segunda Carga")
    
    Me.cmbTipoCarga.ListIndex = 0
    
    
End Sub


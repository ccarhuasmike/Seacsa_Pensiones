VERSION 5.00
Begin VB.Form Frm_AperturaManual 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Apertura de Periodo Manual"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   4185
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4095
      Begin VB.Label Label1 
         Caption         =   "Periodo a Abrir"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Lbl_Periodo 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1920
         TabIndex        =   4
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Fra_Operaciones 
      Height          =   1095
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   4095
      Begin VB.CommandButton Cmd_Abrir 
         Caption         =   "&Abrir"
         Height          =   675
         Left            =   1080
         Picture         =   "Frm_AperturaManual.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Abrir Periodo"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   2400
         Picture         =   "Frm_AperturaManual.frx":0822
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
   End
End
Attribute VB_Name = "Frm_AperturaManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Formulario que permite ReAbrir un periodo de pago, ya se de primeros
'pagos o de pagos en regimen, para volver a ralizar el calculo definitivo
'del periodo.
'vgApeManualSel= Variable general que entrega el tipo de pago que se
'desea reabrir, según la opción seleccionada del menu.
'Valores posibles:
'ApeManPP= Apertura Manual de Primeros Pagos
'ApeManPR= Apertura Manual de Pagos en Regimen

Dim vlAnno As String
Dim vlMes As String
Dim vlPeriodoSig As String 'Contiene el periodo siguiente al que se desea reabrir
Dim vlPeriodo As String 'Contiene el periodo que se desea reabrir

Const clCodEstadoC As String * 1 = "C" 'Código Estado de Periodo Cerrado
Const clCodEstadoA As String * 1 = "A" 'Código Estado de Periodo Abierto
Const clCodIndApeA As String * 1 = "A" 'Código Indicador de Apertura Automática
Const clCodIndApeM As String * 1 = "M" 'Código Indicador de Apertura Manual
'CORPTEC
Dim sTipoPro As String
Dim num_LogProc As Long
'CORPTEC
Function flLog_Proc() As Boolean
    Dim com As ADODB.Command
    Dim sistema, modulo, opcion, origen, tipo As String
    sistema = "SEACSA"
    modulo = "PENSIONES"
    opcion = "PAGOS RECURRENTES- REPROCESO"
    origen = "A"
    tipo = "A"

    Set com = New ADODB.Command
    
    vgConexionBD.BeginTrans
    com.ActiveConnection = vgConexionBD
    com.CommandText = "SP_LOG_CARGAPROCESO"
    com.CommandType = adCmdStoredProc
    
    com.Parameters.Append com.CreateParameter("ESTADO", adChar, adParamInput, 1, sTipoPro)
    com.Parameters.Append com.CreateParameter("USUARIO", adVarChar, adParamInput, 10, vgLogin)
    com.Parameters.Append com.CreateParameter("SISTEMA", adVarChar, adParamInput, 50, sistema)
    com.Parameters.Append com.CreateParameter("MODULO", adVarChar, adParamInput, 50, modulo)
    com.Parameters.Append com.CreateParameter("OPCION", adVarChar, adParamInput, 50, opcion)
    com.Parameters.Append com.CreateParameter("ORIGEN", adChar, adParamInput, 1, origen)
    com.Parameters.Append com.CreateParameter("TIPO", adChar, adParamInput, 1, tipo)
    com.Parameters.Append com.CreateParameter("IDLOG", adDouble, adParamInput, 2, 0)
    com.Parameters.Append com.CreateParameter("Retorno", adDouble, adParamReturnValue)
    com.Execute
    vgConexionBD.CommitTrans
    num_LogProc = com("Retorno")

End Function

Private Sub Cmd_Abrir_Click()
On Error GoTo Err_Cmd_Abrir_Click

    If vlPeriodo = "" Then
        MsgBox "No exite Periodo a ReAbrir", vbCritical, "Sin Datos"
        Screen.MousePointer = 0
        Exit Sub
    End If
    If flEstadoPeriodoSiguiente(vlPeriodo) = clCodIndApeM Then
        MsgBox "No puede Abrir el Periodo: " & Lbl_Periodo.Caption & " Ya que el Siguiente ya fue ReAbierto", vbExclamation, "Información"
        Exit Sub
    Else
        vgRes = MsgBox("¿ Está seguro que desea ReAbrir el Periodo: " & Lbl_Periodo.Caption & " ?", 4 + 32 + 256, "Operación de Apertura")
        If vgRes <> 6 Then
           Screen.MousePointer = 0
           Exit Sub
        End If
        
        'CORPTEC
        sTipoPro = "I"
        Call flLog_Proc
        
        flGrabarApertura (Trim(vlPeriodo))
        'CORPTEC
        sTipoPro = "F"
        Call flLog_Proc
        
        MsgBox "El Periodo: " & Lbl_Periodo.Caption & " Fue ReAbierto con Exito", vbExclamation, "Información"
    End If

Exit Sub
Err_Cmd_Abrir_Click:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub

Private Sub cmd_salir_Click()
On Error GoTo Err_Salir

    Screen.MousePointer = 11
    Unload Me
    Screen.MousePointer = 0

Exit Sub
Err_Salir:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub

Private Sub Form_Load()
On Error GoTo Err_Carga
    
    Frm_AperturaManual.Left = 0
    Frm_AperturaManual.Top = 0
    vlPeriodo = flUltimoPeriodoCerrado
    Lbl_Periodo = Mid(vlPeriodo, 5, 2) & "-" & Mid(vlPeriodo, 1, 4)
    
Exit Sub
Err_Carga:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub

Function flUltimoPeriodoCerrado() As String
On Error GoTo Err_flUltimoPeriodoCerrado
'Permite obtener el último periodo que se encuentra cerrado, es decir,
'que tiene calculo definitivo

    flUltimoPeriodoCerrado = ""

    vgSql = ""
    vgSql = "SELECT p.num_perpago "
    vgSql = vgSql & "FROM pp_tmae_propagopen p "
    vgSql = vgSql & "WHERE "
    If vgApeManualSel = "ApeManPP" Then
        vgSql = vgSql & "p.cod_estadopri = '" & clCodEstadoC & "' "
    Else
        vgSql = vgSql & "p.cod_estadoreg = '" & clCodEstadoC & "' "
    End If
    vgSql = vgSql & "ORDER BY num_perpago DESC"
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
        flUltimoPeriodoCerrado = Trim(vgRegistro!Num_PerPago)
    End If

Exit Function
Err_flUltimoPeriodoCerrado:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function flEstadoPeriodoSiguiente(iPeriodo As String) As String
On Error GoTo Err_flEstadoPeriodoSiguiente
'Permite obtener el estado del indicador del periodo
'inmediatamente siguiente al que se intenta reabrir
'Esto permite determinar si el siguiente periodo se encuetra abierto
'via manual
    
    flEstadoPeriodoSiguiente = ""
    
    vlAnno = Mid(iPeriodo, 1, 4)
    vlMes = Mid(iPeriodo, 5, 2)
    vlPeriodoSig = DateSerial(vlAnno, vlMes + 1, "01")
    vlPeriodoSig = Format(vlPeriodoSig, "yyyymmdd")
    vlPeriodoSig = Mid(vlPeriodoSig, 1, 6)
    
    vgSql = ""
    If vgApeManualSel = "ApeManPP" Then
        vgSql = "SELECT p.cod_indapepri "
    Else
        vgSql = "SELECT p.cod_indapereg "
    End If
    vgSql = vgSql & "FROM pp_tmae_propagopen p "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "p.num_perpago = '" & vlPeriodoSig & "' "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
        If vgApeManualSel = "ApeManPP" Then
            flEstadoPeriodoSiguiente = (vgRegistro!cod_indapepri)
        Else
            flEstadoPeriodoSiguiente = (vgRegistro!cod_indapereg)
        End If
    End If

Exit Function
Err_flEstadoPeriodoSiguiente:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function flGrabarApertura(iPeriodo As String)
On Error GoTo Err_flGrabarApertura
'Permite modificar el estado del periodo a reabrir
'En definitiva, abre el periodo

    vgSql = ""
    vgSql = "UPDATE pp_tmae_propagopen SET "
    If vgApeManualSel = "ApeManPP" Then
        vgSql = vgSql & "cod_estadopri = '" & clCodEstadoA & "', "
        vgSql = vgSql & "cod_indapepri = '" & clCodIndApeM & "' "
    Else
        vgSql = vgSql & "cod_estadoreg = '" & clCodEstadoA & "', "
        vgSql = vgSql & "cod_indapereg = '" & clCodIndApeM & "' "
    End If
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "num_perpago = '" & iPeriodo & "' "
    vgConexionBD.Execute (vgSql)

Exit Function
Err_flGrabarApertura:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function



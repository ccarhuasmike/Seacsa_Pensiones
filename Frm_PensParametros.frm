VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "Crystl32.OCX"
Begin VB.Form Frm_PensParametros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros del Cálculo de Pagos Recurrentes"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   6885
   Begin VB.Frame Frame1 
      Caption         =   "  Especificación de Períodos de Pago  "
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
      Height          =   3135
      Left            =   80
      TabIndex        =   16
      Top             =   80
      Width           =   6765
      Begin VB.Frame Fra_PagoReg 
         Height          =   1815
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   6525
         Begin VB.TextBox Txt_PagoReg 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3150
            MaxLength       =   10
            TabIndex        =   4
            Top             =   480
            Width           =   1155
         End
         Begin VB.TextBox Txt_CalPagoReg 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3150
            MaxLength       =   10
            TabIndex        =   5
            Top             =   720
            Width           =   1155
         End
         Begin VB.TextBox Txt_ProxPago 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3150
            MaxLength       =   10
            TabIndex        =   6
            Top             =   960
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.OptionButton Opt_AbiertoReg 
            Caption         =   "Abierto"
            Height          =   375
            Left            =   2280
            TabIndex        =   7
            Top             =   1305
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton Opt_ProvisorioReg 
            Caption         =   "Provisorio"
            Height          =   375
            Left            =   3360
            TabIndex        =   8
            Top             =   1305
            Width           =   1215
         End
         Begin VB.OptionButton Opt_CerradoReg 
            Caption         =   "Cerrado"
            Height          =   375
            Left            =   4560
            TabIndex        =   9
            Top             =   1305
            Width           =   975
         End
         Begin VB.Label lbl_nombre 
            Caption         =   "Fecha Pago"
            Height          =   255
            Index           =   1
            Left            =   630
            TabIndex        =   27
            Top             =   480
            Width           =   1875
         End
         Begin VB.Label lbl_nombre 
            Caption         =   "Fecha Cálculo"
            Height          =   255
            Index           =   2
            Left            =   630
            TabIndex        =   26
            Top             =   720
            Width           =   2355
         End
         Begin VB.Label lbl_nombre 
            Caption         =   "Fecha Próximo Pago"
            Height          =   255
            Index           =   4
            Left            =   630
            TabIndex        =   25
            Top             =   960
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.Label lbl_nombre 
            Caption         =   "Definición Pagos Recurrentes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   5
            Left            =   135
            TabIndex        =   24
            Top             =   240
            Width           =   3915
         End
         Begin VB.Label lbl_nombre 
            Caption         =   "Estado del Período"
            Height          =   255
            Index           =   12
            Left            =   630
            TabIndex        =   23
            Top             =   1305
            Width           =   1695
         End
      End
      Begin VB.Frame Fra_Busqueda 
         Height          =   900
         Left            =   120
         TabIndex        =   17
         Top             =   255
         Width           =   6525
         Begin VB.TextBox Txt_Anno 
            Height          =   285
            Left            =   4320
            MaxLength       =   4
            TabIndex        =   2
            Top             =   360
            Width           =   795
         End
         Begin VB.TextBox Txt_Mes 
            Height          =   285
            Left            =   3000
            MaxLength       =   2
            TabIndex        =   1
            Top             =   360
            Width           =   435
         End
         Begin VB.CommandButton Cmd_Buscar 
            Height          =   375
            Left            =   5520
            Picture         =   "Frm_PensParametros.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Efectuar Busqueda de IPC"
            Top             =   255
            Width           =   855
         End
         Begin VB.Label lbl_nombre 
            Caption         =   " Definición Período de Pago"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   11
            Left            =   135
            TabIndex        =   21
            Top             =   15
            Width           =   2475
         End
         Begin VB.Label lbl_nombre 
            Caption         =   "Período de Pago"
            Height          =   255
            Index           =   8
            Left            =   600
            TabIndex        =   20
            Top             =   390
            Width           =   1515
         End
         Begin VB.Label lbl_nombre 
            Caption         =   "Año"
            Height          =   210
            Index           =   7
            Left            =   3855
            TabIndex        =   19
            Top             =   405
            Width           =   570
         End
         Begin VB.Label lbl_nombre 
            Caption         =   "Mes"
            Height          =   270
            Index           =   10
            Left            =   2550
            TabIndex        =   18
            Top             =   405
            Width           =   570
         End
      End
   End
   Begin VB.Frame Fra_Operacion 
      Height          =   1095
      Left            =   80
      TabIndex        =   0
      Top             =   5640
      Width           =   6735
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   3000
         Picture         =   "Frm_PensParametros.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Eliminar 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   1680
         Picture         =   "Frm_PensParametros.frx":07BC
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   4320
         Picture         =   "Frm_PensParametros.frx":0AFE
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5640
         Picture         =   "Frm_PensParametros.frx":11B8
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Grabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   400
         Picture         =   "Frm_PensParametros.frx":12B2
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Grabar Datos"
         Top             =   240
         Width           =   720
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Msf_Grilla 
      Height          =   2295
      Left            =   75
      TabIndex        =   15
      Top             =   3360
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   4048
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   14745599
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Crystal.CrystalReport Rpt_General 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "Frm_PensParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Sql As String

Dim vlPeriodoMin As String
Dim vlPeriodoMax As String
Dim vlPeriodo As String
Dim vlAnno As String
Dim vlMes As String
Dim vlArchivo As String
Dim vlPos As Integer
Dim vlNumero As Integer
Dim vlEstadoPri As String * 1
Dim vlEstadoReg As String * 1


Dim vlGlsUsuarioCrea As Variant
Dim vlFecCrea As Variant
Dim vlHorCrea As Variant
Dim vlGlsUsuarioModi As Variant
Dim vlFecModi As Variant
Dim vlHorModi As Variant
Function flInicializaGrilla()
'Permite limpiar e inicializar la grilla.
'----------------------------------------------------------------------

On Error GoTo Err_flInicializaGrilla

    Msf_Grilla.Clear
    Msf_Grilla.Cols = 4
    Msf_Grilla.Rows = 1
    Msf_Grilla.RowHeight(0) = 250
    Msf_Grilla.Row = 0
        
    Msf_Grilla.Col = 0
    Msf_Grilla.Text = "Periodo Pago"
    Msf_Grilla.ColWidth(0) = 1050
    Msf_Grilla.ColAlignment(0) = 1  'centrado
        
    Msf_Grilla.Col = 1
    Msf_Grilla.Text = "Fecha Pago Recurrente"
    Msf_Grilla.ColWidth(1) = 1800
    
    Msf_Grilla.Col = 2
    Msf_Grilla.Text = "Fecha Cálculo Pago Recurrente"
    Msf_Grilla.ColWidth(2) = 2400
    
'    Msf_Grilla.Col = 3
'    Msf_Grilla.Text = "Proximo Pago"
'    Msf_Grilla.ColWidth(3) = 1500
    
    Msf_Grilla.Col = 3
    Msf_Grilla.Text = "Estado Periodo"
    Msf_Grilla.ColWidth(3) = 1200
    
Exit Function
Err_flInicializaGrilla:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flCargaGrilla()
'Permite cargar la grilla con los datos registrados en la Base de Datos.
'----------------------------------------------------------------------

On Error GoTo Err_Carga
    
    Call flInicializaGrilla
    
    'vgPalabra = Mid(num_perpago, 1, 4)
    
    vgSql = ""
    vgSql = "SELECT num_perpago,fec_calpagoreg,fec_pagoreg, "
    vgSql = vgSql & "fec_pagoproxreg,cod_estadoreg "
    vgSql = vgSql & "FROM pp_tmae_propagopen "
    vgSql = vgSql & "WHERE "
    If vgTipoBase = "ORACLE" Then
       vgSql = vgSql & "substr(num_perpago,1,4) = '" & Trim(Txt_Anno.Text) & "' "
    Else
        vgSql = vgSql & "substring(num_perpago,1,4) = '" & Trim(Txt_Anno.Text) & "' "
    End If
    'vgSql = vgSql & "WHERE num_perpago like '" & Trim(Txt_Anno.Text) & "%' "
    vgSql = vgSql & "ORDER by num_perpago "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
       While Not vgRs.EOF
          vlAnno = (Mid(vgRs!Num_PerPago, 1, 4))
          vlMes = (Mid(vgRs!Num_PerPago, 5, 2))
          Msf_Grilla.AddItem ((Trim(vlMes) & " - " & Trim(vlAnno))) & vbTab _
          & DateSerial(Mid((vgRs!Fec_PagoReg), 1, 4), Mid((vgRs!Fec_PagoReg), 5, 2), Mid((vgRs!Fec_PagoReg), 7, 2)) & vbTab _
          & DateSerial(Mid((vgRs!fec_calpagoreg), 1, 4), Mid((vgRs!fec_calpagoreg), 5, 2), Mid((vgRs!fec_calpagoreg), 7, 2)) & vbTab _
          & (Trim(vgRs!cod_estadoreg))
          '& DateSerial(Mid((vgRs!Fec_PagoProxReg), 1, 4), Mid((vgRs!Fec_PagoProxReg), 5, 2), Mid((vgRs!Fec_PagoProxReg), 7, 2)) & vbTab
          vgRs.MoveNext
       Wend
    End If
    vgRs.Close

Exit Function
Err_Carga:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function flLimpiar()

On Error GoTo Err_flLimpiar
    
    Call flInicializaGrilla
    Txt_Mes.Text = ""
    Txt_Anno.Text = ""
    Txt_PagoReg.Text = ""
    Txt_CalPagoReg.Text = ""
    Txt_ProxPago.Text = ""
'    Opt_Abierto.Value = True

Exit Function
Err_flLimpiar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function flHabilitaIngreso()

On Error GoTo flHabilitaIngreso

    Fra_Busqueda.Enabled = False

'    Txt_Mes.Enabled = False
'    Txt_Anno.Enabled = False
'    Cmd_Buscar.Enabled = False
        
    If Opt_AbiertoReg.Value = True Then
    
       Fra_PagoReg.Enabled = True
    
'       Opt_AbiertoReg.Enabled = True
'       Opt_ProvisorioReg.Enabled = True
'       Opt_CerradoReg.Enabled = True
'       Txt_PagoReg.Enabled = True
'       Txt_CalPagoReg.Enabled = True
'       Txt_ProxPago.Enabled = True
    Else
    
        Fra_PagoReg.Enabled = False
    
'        Opt_AbiertoReg.Enabled = False
'        Opt_ProvisorioReg.Enabled = False
'        Opt_CerradoReg.Enabled = False
'        Txt_PagoReg.Enabled = False
'        Txt_CalPagoReg.Enabled = False
'        Txt_ProxPago.Enabled = False
    End If
    
    If Opt_AbiertoReg = True Then
       If Txt_PagoReg.Enabled = True Then
          Txt_PagoReg.SetFocus
       End If
    Else
        Cmd_Grabar.SetFocus
    End If
    
    
    
'    Txt_PrimerPago.Enabled = True
'    Txt_CalPrimerPago.Enabled = True
'    Txt_PagoReg.Enabled = True
'    Txt_CalPagoReg.Enabled = True
'    Txt_ProxPago.Enabled = True
    
'    Txt_PrimerPago.SetFocus

Exit Function
flHabilitaIngreso:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select


End Function

Function flDeshabilitaIngreso()

On Error GoTo Err_flDeshabilitaIngreso

    Fra_Busqueda.Enabled = True
    
    Fra_PagoReg.Enabled = False
    Opt_AbiertoReg.Value = True
    
    Txt_Mes.SetFocus
    
Exit Function
Err_flDeshabilitaIngreso:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
    

End Function


Private Sub Cmd_Buscar_Click()

On Error GoTo Err_CmdBuscarClick

'Valida Año
    If Txt_Anno.Text = "" Then
       MsgBox "Debe Ingresar Año del Periodo de Pago.", vbCritical, "Error de Datos"
       Txt_Anno.SetFocus
       Exit Sub
    End If
    If CDbl(Txt_Anno.Text) < 1900 Then
       MsgBox "Debe Ingresar un Año Mayor a 1900.", vbCritical, "Error de Datos"
       Txt_Anno.SetFocus
       Exit Sub
    End If
'Valida Mes
    If Txt_Mes.Text = "" Then
       MsgBox "Debe Ingresar Mes del Periodo de Pago.", vbCritical, "Error de Datos"
       Txt_Mes.SetFocus
       Exit Sub
    End If
    If CDbl(Txt_Mes.Text) <= 0 Or CDbl(Txt_Mes.Text) > 12 Then
       MsgBox "El Mes Ingresado No es un Valor Válido.", vbCritical, "Error de Datos"
       Txt_Mes.SetFocus
       Exit Sub
    End If
        
'    vlPeriodo = Trim(Str(Txt_Anno.Text)) + Trim(Str(Txt_Mes.Text))
    vlPeriodo = Format(Txt_Anno.Text, "0000") & Format(Txt_Mes.Text, "00")
                  
    vgSql = "SELECT * "
    vgSql = vgSql & "FROM PP_TMAE_PROPAGOPEN "
    vgSql = vgSql & "WHERE (num_perpago = '" & Trim(vlPeriodo) & "') "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
       
       If (vgRs!cod_estadoreg) = "A" Then
          Opt_AbiertoReg.Value = True
       Else
           If (vgRs!cod_estadoreg) = "P" Then
              Opt_ProvisorioReg.Value = True
           Else
               Opt_CerradoReg.Value = True
           End If
       End If
              
       Txt_PagoReg = DateSerial(Mid((vgRs!Fec_PagoReg), 1, 4), Mid((vgRs!Fec_PagoReg), 5, 2), Mid((vgRs!Fec_PagoReg), 7, 2))
       If Not IsNull(vgRs!fec_calpagoreg) Then
        Txt_CalPagoReg = DateSerial(Mid((vgRs!fec_calpagoreg), 1, 4), Mid((vgRs!fec_calpagoreg), 5, 2), Mid((vgRs!fec_calpagoreg), 7, 2))
       Else
        Txt_CalPagoReg = ""
       End If
       If Not IsNull(vgRs!Fec_PagoProxReg) Then
        Txt_ProxPago = DateSerial(Mid((vgRs!Fec_PagoProxReg), 1, 4), Mid((vgRs!Fec_PagoProxReg), 5, 2), Mid((vgRs!Fec_PagoProxReg), 7, 2))
       Else
        Txt_ProxPago = ""
       End If
    Else
        vgSql = "SELECT * "
        vgSql = vgSql & "FROM PP_TMAE_CALENDARIOPAGO "
        vgSql = vgSql & "WHERE (num_perpago = '" & Trim(vlPeriodo) & "') "
        Set vgRs = vgConexionBD.Execute(vgSql)
        If Not vgRs.EOF Then
           Txt_PagoReg = DateSerial(Mid((vgRs!Fec_PagoReg), 1, 4), Mid((vgRs!Fec_PagoReg), 5, 2), Mid((vgRs!Fec_PagoReg), 7, 2))
           If Not IsNull(vgRs!fec_calpagoreg) Then
            Txt_CalPagoReg = DateSerial(Mid((vgRs!fec_calpagoreg), 1, 4), Mid((vgRs!fec_calpagoreg), 5, 2), Mid((vgRs!fec_calpagoreg), 7, 2))
           Else
            Txt_CalPagoReg = ""
           End If
           If Not IsNull(vgRs!Fec_PagoProxReg) Then
            Txt_ProxPago = DateSerial(Mid((vgRs!Fec_PagoProxReg), 1, 4), Mid((vgRs!Fec_PagoProxReg), 5, 2), Mid((vgRs!Fec_PagoProxReg), 7, 2))
           Else
            Txt_ProxPago = ""
           End If
           Opt_AbiertoReg.Value = True
        End If
       
    End If
                      
    Call flCargaGrilla
    Call flHabilitaIngreso
    
Exit Sub
Err_CmdBuscarClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
    

End Sub

Private Sub Cmd_Eliminar_Click()

On Error GoTo Err_CmdEliminarClick

'Valida Año
    If Txt_Anno.Text = "" Then
       MsgBox "Debe Ingresar Año del Periodo de Pago a Eliminar.", vbCritical, "Error de Datos"
       Txt_Anno.SetFocus
       Exit Sub
    End If
    If CDbl(Txt_Anno.Text) < 1900 Then
       MsgBox "Debe Ingresar un Año Mayor a 1900.", vbCritical, "Error de Datos"
       Txt_Anno.SetFocus
       Exit Sub
    End If
'Valida Mes
    If Txt_Mes.Text = "" Then
       MsgBox "Debe Ingresar Mes del Periodo de Pago a Eliminar.", vbCritical, "Error de Datos"
       Txt_Mes.SetFocus
       Exit Sub
    End If
    If CDbl(Txt_Mes.Text) <= 0 Or CDbl(Txt_Mes.Text) > 12 Then
       MsgBox "El Mes Ingresado No es un Valor Válido.", vbCritical, "Error de Datos"
       Txt_Mes.SetFocus
       Exit Sub
    End If
    
'    vlPeriodo = Trim(Str(Txt_Anno.Text)) + Trim(Str(Txt_Mes.Text))
    vlPeriodo = Format(Txt_Anno.Text, "0000") & Format(Txt_Mes.Text, "00")
     
    vgSql = "SELECT num_perpago "
    vgSql = vgSql & "FROM PP_TMAE_PROPAGOPEN "
    vgSql = vgSql & "WHERE (num_perpago = '" & Trim(vlPeriodo) & "') "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
       Screen.MousePointer = 11
       vlNumero = MsgBox("¿Realmente Desea Eliminar el Registro Seleccionado ?", vbQuestion + vbYesNo + 256, "Confirmación")
       Screen.MousePointer = 0
       If vlNumero <> 6 Then
          Call Cmd_Limpiar_Click
       Else
           Sql = " DELETE PP_TMAE_PROPAGOPEN "
           Sql = Sql & "WHERE (num_perpago = '" & Trim(vlPeriodo) & "') "
           vgConexionBD.Execute Sql
           MsgBox "La Eliminación de Datos fue realizada Correctamente", vbInformation, "Información"
       End If
    Else
        MsgBox "El registro que Desea Eliminar No Existe", vbInformation, "Información"
    End If
           
'    vlPeriodoMin = Trim(Str(Txt_Anno.Text)) + Trim(Str(1))
'    vlPeriodoMax = Trim(Str(Txt_Anno.Text)) + Trim(Str(12))
            
    Call Cmd_Limpiar_Click
    Txt_Anno.Text = (Mid(vlPeriodo, 1, 4))
    Call flCargaGrilla
    Call flDeshabilitaIngreso
    
Exit Sub
Err_CmdEliminarClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
    

End Sub

Private Sub cmd_grabar_Click()

On Error GoTo Err_CmdGrabarClick

'Valida Año
    If Txt_Anno.Text = "" Then
       MsgBox "Debe Ingresar Año del Periodo de Pago.", vbCritical, "Error de Datos"
       Txt_Anno.SetFocus
       Exit Sub
    End If
    If CDbl(Txt_Anno.Text) < 1900 Then
       MsgBox "Debe Ingresar un Año Mayor a 1900.", vbCritical, "Error de Datos"
       Txt_Anno.SetFocus
       Exit Sub
    End If
'Valida Mes
    If Txt_Mes.Text = "" Then
       MsgBox "Debe Ingresar Mes del Periodo de Pago.", vbCritical, "Error de Datos"
       Txt_Mes.SetFocus
       Exit Sub
    End If
    If CDbl(Txt_Mes.Text) <= 0 Or CDbl(Txt_Mes.Text) > 12 Then
       MsgBox "El Mes Ingresado No es un Valor Válido.", vbCritical, "Error de Datos"
       Txt_Mes.SetFocus
       Exit Sub
    End If
'Valida Fecha de Pago en Regimen
    If (Trim(Txt_PagoReg) = "") Then
       MsgBox "Debe Ingresar Fecha de Pago en Régimen", vbCritical, "Error de Datos"
       If Txt_PagoReg.Enabled Then
        Txt_PagoReg.SetFocus
       End If
       Exit Sub
    End If
    If Not IsDate(Txt_PagoReg.Text) Then
       MsgBox "La Fecha Ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
       If Txt_PagoReg.Enabled Then
        Txt_PagoReg.SetFocus
       End If
       Exit Sub
    End If
    If (Year(Txt_PagoReg) < 1900) Then
       MsgBox "La Fecha Ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
       If Txt_PagoReg.Enabled Then
        Txt_PagoReg.SetFocus
       End If
       Exit Sub
    End If
'Valida Fecha de Calculo para Pago en Regimen
    If (Trim(Txt_CalPagoReg) = "") Then
       MsgBox "Debe Ingresar Fecha para Cálculo de Pago en Régimen", vbCritical, "Error de Datos"
       Txt_CalPagoReg.SetFocus
       Exit Sub
    End If
    If Not IsDate(Txt_CalPagoReg.Text) Then
       MsgBox "La Fecha Ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
       Txt_CalPagoReg.SetFocus
       Exit Sub
    End If
    If (Year(Txt_CalPagoReg) < 1900) Then
       MsgBox "La Fecha Ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
       Txt_CalPagoReg.SetFocus
       Exit Sub
    End If
'Valida Fecha de Proximo Pago
'    If (Trim(Txt_ProxPago) = "") Then
'       MsgBox "Debe Ingresar Fecha de Proximo Pago en Régimen", vbCritical, "Error de Datos"
'       Txt_ProxPago.SetFocus
'       Exit Sub
'    End If
'    If Not IsDate(Txt_ProxPago.Text) Then
'       MsgBox "La Fecha Ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
'       Txt_ProxPago.SetFocus
'       Exit Sub
'    End If
'    If (Year(Txt_ProxPago) < 1900) Then
'       MsgBox "La Fecha Ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
'       Txt_ProxPago.SetFocus
'       Exit Sub
'    End If
         
    vlPeriodo = Format(Txt_Anno.Text, "0000") & Format(Txt_Mes.Text, "00")
         
    If Opt_AbiertoReg.Value = True Then
       vlEstadoReg = "A"
    Else
        If Opt_ProvisorioReg.Value = True Then
           vlEstadoReg = "P"
        Else
            vlEstadoReg = "C"
        End If
    End If
     
    vgSql = "SELECT num_perpago "
    vgSql = vgSql & "FROM PP_TMAE_PROPAGOPEN "
    vgSql = vgSql & "WHERE (num_perpago = '" & Trim(vlPeriodo) & "') "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
       Screen.MousePointer = 11
       vlNumero = MsgBox("¿Los Datos Ya Existen en la Base de Datos, Desea Modificarlos ?", vbQuestion + vbYesNo + 256, "Confirmación")
       Screen.MousePointer = 0
       If vlNumero <> 6 Then
          Call Cmd_Limpiar_Click
       Else
           vlGlsUsuarioModi = vgUsuario
           vlFecModi = Format(Date, "yyyymmdd")
           vlHorModi = Format(Time, "hhmmss")
           
           Sql = ""
           Sql = " UPDATE PP_TMAE_PROPAGOPEN SET "
           Sql = Sql & " fec_pagoreg = '" & Format(CDate(Trim(Txt_PagoReg)), "yyyymmdd") & "', "
           If Trim(Txt_CalPagoReg) <> "" Then
            Sql = Sql & " fec_calpagoreg = '" & Format(CDate(Trim(Txt_CalPagoReg)), "yyyymmdd") & "', "
           Else
            Sql = Sql & " fec_calpagoreg = NULL, "
           End If
           If Trim(Txt_ProxPago) <> "" Then
            Sql = Sql & " fec_pagoproxreg = '" & Format(CDate(Trim(Txt_ProxPago)), "yyyymmdd") & "', "
           Else
            Sql = Sql & " fec_pagoproxreg = NULL, "
           End If
           Sql = Sql & " cod_estadoreg = '" & Trim(vlEstadoReg) & "', "
           Sql = Sql & " cod_usuariomodi = '" & vlGlsUsuarioModi & "', "
           Sql = Sql & " fec_modi = '" & vlFecModi & "', "
           Sql = Sql & " hor_modi = '" & vlHorModi & "' "
           Sql = Sql & " WHERE num_perpago =  '" & Trim(vlPeriodo) & "' "
           vgConexionBD.Execute Sql
           MsgBox "La Actualización de Datos fue realizado Correctamente", vbInformation, "Información"
           Call Cmd_Limpiar_Click
       End If
    Else
        vlGlsUsuarioCrea = vgUsuario
        vlFecCrea = Format(Date, "yyyymmdd")
        vlHorCrea = Format(Time, "hhmmss")
        Sql = ""
        Sql = "INSERT INTO PP_TMAE_PROPAGOPEN "
        Sql = Sql & "(num_perpago,fec_pagoreg,"
        Sql = Sql & " fec_calpagoreg,fec_pagoproxreg,cod_estadoreg, "
        Sql = Sql & " cod_usuariocrea,fec_crea,hor_crea "
        Sql = Sql & " ) VALUES ( "
        Sql = Sql & " '" & Trim(Str(vlPeriodo)) & "', "
        Sql = Sql & " '" & Format(CDate(Trim(Txt_PagoReg)), "yyyymmdd") & "', "
        If Trim(Txt_CalPagoReg) <> "" Then
            Sql = Sql & " '" & Format(CDate(Trim(Txt_CalPagoReg)), "yyyymmdd") & "', "
        Else
            Sql = Sql & " NULL, "
        End If
        If Trim(Txt_ProxPago) <> "" Then
            Sql = Sql & " '" & Format(CDate(Trim(Txt_ProxPago)), "yyyymmdd") & "', "
        Else
            Sql = Sql & " NULL, "
        End If
        Sql = Sql & " '" & Trim(vlEstadoReg) & "', "
        Sql = Sql & " '" & vlGlsUsuarioCrea & "', "
        Sql = Sql & " '" & vlFecCrea & "', "
        Sql = Sql & " '" & vlHorCrea & "') "
        vgConexionBD.Execute Sql
        MsgBox "El registro de Datos fue realizado Correctamente", vbInformation, "Información"
        Call Cmd_Limpiar_Click
    End If
    
    Txt_Anno.Text = (Mid(vlPeriodo, 1, 4))
'    Txt_Mes.Text = (Mid(vlPeriodo, 5, 2))
    
'    vlPeriodoMin = Trim(Str(Txt_Anno.Text)) + Trim(Str(1))
'    vlPeriodoMax = Trim(Str(Txt_Anno.Text)) + Trim(Str(12))
            
    Call flDeshabilitaIngreso
    Call flCargaGrilla
    
Exit Sub
Err_CmdGrabarClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
    
       
End Sub

Private Sub Cmd_Imprimir_Click()

On Error GoTo Err_CmdImprimir

   vlArchivo = strRpt & "PP_Rpt_PensParametros.rpt"   '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
       MsgBox "Archivo de Reporte de Parámetros del Cálculo de Pensiones no se encuentra en el Directorio de la Aplicación.", 16, "Archivo no encontrado"
       Screen.MousePointer = 0
       Exit Sub
   End If
   
   Rpt_General.Reset
   Rpt_General.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
   Rpt_General.Connect = vgRutaDataBase       ' o App.Path & "\Nestle.mdb"
  
   Rpt_General.Formulas(0) = ""
   Rpt_General.Formulas(1) = ""
   Rpt_General.Formulas(2) = ""
   
   Rpt_General.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
   Rpt_General.Formulas(1) = "NombreSistema= '" & vgNombreSistema & "'"
   Rpt_General.Formulas(2) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
      
   Rpt_General.WindowState = crptMaximized
   Rpt_General.Destination = crptToWindow
   Rpt_General.WindowTitle = "Informe de Parámetros de Cálculo de Pensiones"
   Rpt_General.Action = 1
   Screen.MousePointer = 0
   
Exit Sub
Err_CmdImprimir:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub

Private Sub Cmd_Limpiar_Click()

On Error GoTo Err_CmdLimpiarClick

    Call flLimpiar
    Call flDeshabilitaIngreso
    
Exit Sub
Err_CmdLimpiarClick:
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

On Error GoTo Err_Cargar

    Frm_PensParametros.Top = 0
    Frm_PensParametros.Left = 0
    
    Call flLimpiar
    Call flInicializaGrilla
    
    Fra_Busqueda.Enabled = True
    
'    Txt_Mes.Enabled = True
'    Txt_Anno.Enabled = True
'    Cmd_Buscar.Enabled = True
    
    Opt_AbiertoReg.Value = True
    
'    Opt_AbiertoPri.Enabled = False
'    Opt_ProvisorioPri.Enabled = False
'    Opt_CerradoPri.Enabled = False
'    Txt_PrimerPago.Enabled = False
'    Txt_CalPrimerPago.Enabled = False
    
    Fra_PagoReg.Enabled = False
    
'    Opt_AbiertoReg.Enabled = False
'    Opt_ProvisorioReg.Enabled = False
'    Opt_CerradoReg.Enabled = False
'    Txt_PagoReg.Enabled = False
'    Txt_CalPagoReg.Enabled = False
'    Txt_ProxPago.Enabled = False
            
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
    
End Sub


Private Sub Msf_Grilla_DblClick()

On Error GoTo Err_GrillaDblClick
    
    Msf_Grilla.Col = 0
    If (Msf_Grilla.Text = "") Or (Msf_Grilla.Row = 0) Then
        MsgBox "No existen Detalles", vbExclamation, "Información"
        Exit Sub
    Else
        Msf_Grilla.Col = 0
        vgPalabra = Trim(Msf_Grilla.Text)
        vlPos = InStr(vgPalabra, "-")
        vlMes = (Trim(Mid(vgPalabra, 1, vlPos - 1)))
        vlAnno = (Trim(Mid(vgPalabra, vlPos + 1, 5)))
        Txt_Mes.Text = vlMes
        Txt_Anno.Text = vlAnno
        Msf_Grilla.Col = 1
        Txt_PagoReg = Msf_Grilla.Text
        Msf_Grilla.Col = 2
        Txt_CalPagoReg = Msf_Grilla.Text
        'Msf_Grilla.Col = 3
        'Txt_ProxPago = Msf_Grilla.Text
        
        Msf_Grilla.Col = 3
        If Trim(Msf_Grilla.Text) = "A" Then
           Opt_AbiertoReg.Value = True
        Else
            If Trim(Msf_Grilla.Text) = "P" Then
               Opt_ProvisorioReg.Value = True
            Else
                Opt_CerradoReg.Value = True
            End If
        End If
                
        Call flHabilitaIngreso
'        Txt_PrimerPago.SetFocus
        
    End If

Exit Sub
Err_GrillaDblClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub

Private Sub Opt_AbiertoPri_KeyPress(KeyAscii As Integer)

On Error GoTo Err_OptAbiertoPriKeyPress

    If KeyAscii = 13 Then
       Txt_PagoReg.SetFocus
    End If
    
Exit Sub
Err_OptAbiertoPriKeyPress:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub

Private Sub Opt_AbiertoReg_KeyPress(KeyAscii As Integer)

On Error GoTo Err_OptAbiertoRegKeyPress

    If KeyAscii = 13 Then
       Cmd_Grabar.SetFocus
    End If
    
Exit Sub
Err_OptAbiertoRegKeyPress:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub

Private Sub Txt_Anno_Change()

On Error GoTo Err_AnnoChange

    If Not IsNumeric(Txt_Anno) Then
       Txt_Anno = ""
    End If
    
Exit Sub
Err_AnnoChange:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub

Private Sub Txt_Anno_KeyPress(KeyAscii As Integer)

On Error GoTo Err_TxtAnnoKeyPress

    If KeyAscii = 13 Then
        If Txt_Anno.Text = "" Then
           MsgBox "Debe Ingresar Año del Periodo de Pago.", vbCritical, "Error de Datos"
           Txt_Anno.SetFocus
           Exit Sub
        End If
    
        If CDbl(Txt_Anno.Text) < 1900 Then
           MsgBox "Debe Ingresar un Año Mayor a 1900.", vbCritical, "Error de Datos"
           Txt_Anno.SetFocus
           Exit Sub
        End If
        Txt_Anno.Text = Format(Txt_Anno.Text, "0000")
        Cmd_Buscar.SetFocus
     End If
     
Exit Sub
Err_TxtAnnoKeyPress:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
     

End Sub

Private Sub Txt_Anno_LostFocus()

    If Txt_Anno.Text = "" Then
       Exit Sub
    End If
    If CDbl(Txt_Anno.Text) < 1900 Then
       Exit Sub
    End If
    Txt_Anno.Text = Format(Txt_Anno.Text, "0000")
        

End Sub

Private Sub Txt_CalPagoReg_KeyPress(KeyAscii As Integer)

On Error GoTo Err_TxtCalPagoRegKP

     If KeyAscii = 13 Then
        If (Trim(Txt_CalPagoReg) = "") Then
           MsgBox "Debe Ingresar Fecha para Cálculo de Pago en Régimen", vbCritical, "Error de Datos"
           Txt_CalPagoReg.SetFocus
           Exit Sub
        End If
        If Not IsDate(Txt_CalPagoReg.Text) Then
           MsgBox "La Fecha Ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
           Txt_CalPagoReg.SetFocus
           Exit Sub
        End If
        If (Year(Txt_CalPagoReg) < 1900) Then
           MsgBox "La Fecha Ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
           Txt_CalPagoReg.SetFocus
           Exit Sub
        End If
        Txt_CalPagoReg = Format(CDate(Trim(Txt_CalPagoReg)), "yyyymmdd")
        Txt_CalPagoReg.Text = DateSerial(Mid((Txt_CalPagoReg.Text), 1, 4), Mid((Txt_CalPagoReg.Text), 5, 2), Mid((Txt_CalPagoReg.Text), 7, 2))
        Txt_ProxPago.SetFocus
     End If
     
Exit Sub
Err_TxtCalPagoRegKP:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub

Private Sub Txt_CalPagoReg_LostFocus()

    If (Trim(Txt_CalPagoReg) = "") Then
       Exit Sub
    End If
    If Not IsDate(Txt_CalPagoReg.Text) Then
       Exit Sub
    End If
    If (Year(Txt_CalPagoReg) < 1900) Then
       Exit Sub
    End If
    Txt_CalPagoReg = Format(CDate(Trim(Txt_CalPagoReg)), "yyyymmdd")
    Txt_CalPagoReg.Text = DateSerial(Mid((Txt_CalPagoReg.Text), 1, 4), Mid((Txt_CalPagoReg.Text), 5, 2), Mid((Txt_CalPagoReg.Text), 7, 2))

End Sub

Private Sub Txt_Mes_Change()

On Error GoTo Err_MesChange

    If Not IsNumeric(Txt_Mes) Then
       Txt_Mes = ""
    End If
    
Exit Sub
Err_MesChange:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Txt_Mes_KeyPress(KeyAscii As Integer)

On Error GoTo Err_TxtMesKeyPress

    If KeyAscii = 13 Then
       If Txt_Mes.Text = "" Then
          MsgBox "Debe Ingresar Mes del Periodo de Pago.", vbCritical, "Error de Datos"
          Txt_Mes.SetFocus
          Exit Sub
       End If
         
       If CDbl(Txt_Mes.Text) <= 0 Or CDbl(Txt_Mes.Text) > 12 Then
          MsgBox "El Mes Ingresado No es un Valor Válido.", vbCritical, "Error de Datos"
          Txt_Mes.SetFocus
          Exit Sub
       End If
        Txt_Mes.Text = Format(Txt_Mes.Text, "00")
        Txt_Anno.SetFocus
    End If
    
Exit Sub
Err_TxtMesKeyPress:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
    

End Sub

Private Sub Txt_Mes_LostFocus()

    If Txt_Mes.Text = "" Then
       Exit Sub
    End If
    If CDbl(Txt_Mes.Text) <= 0 Or CDbl(Txt_Mes.Text) > 12 Then
       Exit Sub
    End If
    Txt_Mes.Text = Format(Txt_Mes.Text, "00")

End Sub

Private Sub Txt_PagoReg_KeyPress(KeyAscii As Integer)

On Error GoTo Err_TxtPagoRegKP
     If KeyAscii = 13 Then
        If (Trim(Txt_PagoReg) = "") Then
           MsgBox "Debe Ingresar Fecha de Pago en Régimen", vbCritical, "Error de Datos"
           Txt_PagoReg.SetFocus
           Exit Sub
        End If
        If Not IsDate(Txt_PagoReg.Text) Then
           MsgBox "La Fecha Ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
           Txt_PagoReg.SetFocus
           Exit Sub
        End If
        If (Year(Txt_PagoReg) < 1900) Then
           MsgBox "La Fecha Ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
           Txt_PagoReg.SetFocus
           Exit Sub
        End If
        Txt_PagoReg = Format(CDate(Trim(Txt_PagoReg)), "yyyymmdd")
        Txt_PagoReg.Text = DateSerial(Mid((Txt_PagoReg.Text), 1, 4), Mid((Txt_PagoReg.Text), 5, 2), Mid((Txt_PagoReg.Text), 7, 2))
        Txt_CalPagoReg.SetFocus
     End If
     
Exit Sub
Err_TxtPagoRegKP:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select


End Sub

Private Sub Txt_PagoReg_LostFocus()

    If (Trim(Txt_PagoReg) = "") Then
       Exit Sub
    End If
    If Not IsDate(Txt_PagoReg.Text) Then
       Exit Sub
    End If
    If (Year(Txt_PagoReg) < 1900) Then
       Exit Sub
    End If
    Txt_PagoReg = Format(CDate(Trim(Txt_PagoReg)), "yyyymmdd")
    Txt_PagoReg.Text = DateSerial(Mid((Txt_PagoReg.Text), 1, 4), Mid((Txt_PagoReg.Text), 5, 2), Mid((Txt_PagoReg.Text), 7, 2))
 
End Sub

Private Sub Txt_ProxPago_KeyPress(KeyAscii As Integer)

On Error GoTo Err_TxtProxPagoKP
     If KeyAscii = 13 Then
        If (Trim(Txt_ProxPago) = "") Then
           MsgBox "Debe Ingresar Fecha de Proximo Pago en Régimen", vbCritical, "Error de Datos"
           Txt_ProxPago.SetFocus
           Exit Sub
        End If
        If Not IsDate(Txt_ProxPago.Text) Then
           MsgBox "La Fecha Ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
           Txt_ProxPago.SetFocus
           Exit Sub
        End If
        If (Year(Txt_ProxPago) < 1900) Then
           MsgBox "La Fecha Ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
           Txt_ProxPago.SetFocus
           Exit Sub
        End If
        If Opt_AbiertoReg.Enabled = True Then
           Txt_ProxPago = Format(CDate(Trim(Txt_ProxPago)), "yyyymmdd")
           Txt_ProxPago.Text = DateSerial(Mid((Txt_ProxPago.Text), 1, 4), Mid((Txt_ProxPago.Text), 5, 2), Mid((Txt_ProxPago.Text), 7, 2))
           Opt_AbiertoReg.SetFocus
        Else
            Txt_ProxPago = Format(CDate(Trim(Txt_ProxPago)), "yyyymmdd")
            Txt_ProxPago.Text = DateSerial(Mid((Txt_ProxPago.Text), 1, 4), Mid((Txt_ProxPago.Text), 5, 2), Mid((Txt_ProxPago.Text), 7, 2))
            Cmd_Grabar.SetFocus
        End If
     End If
     
Exit Sub
Err_TxtProxPagoKP:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub


Private Sub Txt_ProxPago_LostFocus()

    If (Trim(Txt_ProxPago) = "") Then
       Exit Sub
    End If
    If Not IsDate(Txt_ProxPago.Text) Then
       Exit Sub
    End If
    If (Year(Txt_ProxPago) < 1900) Then
       Exit Sub
    End If
    Txt_ProxPago = Format(CDate(Trim(Txt_ProxPago)), "yyyymmdd")
    Txt_ProxPago.Text = DateSerial(Mid((Txt_ProxPago.Text), 1, 4), Mid((Txt_ProxPago.Text), 5, 2), Mid((Txt_ProxPago.Text), 7, 2))
    
End Sub


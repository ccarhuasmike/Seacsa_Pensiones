VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Frm_InfoControlRetJudicial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe Control Retención Judicial"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   4995
   Begin VB.Frame Fra_Cmd 
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   4815
      Begin VB.CommandButton cmd_calcular 
         Caption         =   "&Calcular"
         Height          =   675
         Left            =   1440
         Picture         =   "Frm_InfoControlRetJudicial.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Generar Tabla de Mortalidad"
         Top             =   360
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   3720
         Picture         =   "Frm_InfoControlRetJudicial.frx":04A2
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   2520
         Picture         =   "Frm_InfoControlRetJudicial.frx":059C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   720
      End
      Begin Crystal.CrystalReport Rpt_General 
         Left            =   720
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.Frame Fra_Periodo 
      Caption         =   "  Periodo  "
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
      Height          =   1095
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.TextBox Txt_Anno 
         Height          =   285
         Left            =   2400
         MaxLength       =   4
         TabIndex        =   2
         Top             =   480
         Width           =   795
      End
      Begin VB.TextBox Txt_Mes 
         Height          =   285
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   1
         Top             =   480
         Width           =   435
      End
      Begin VB.Label lbl_nombre 
         Caption         =   "Período de Pago"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   1515
      End
      Begin VB.Label lbl_nombre 
         Caption         =   "(Mes - Año)"
         Height          =   270
         Index           =   10
         Left            =   3360
         TabIndex        =   7
         Top             =   480
         Width           =   930
      End
      Begin VB.Label lbl_nombre 
         Alignment       =   2  'Center
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   2160
         TabIndex        =   6
         Top             =   480
         Width           =   195
      End
   End
End
Attribute VB_Name = "Frm_InfoControlRetJudicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vlFechaInicio As String
Dim vlFechaTermino As String

Dim vlArchivo As String

Private Sub Cmd_Imprimir_Click()

On Error GoTo Err_CmdImprimir
 
   'Valida Mes del Periodo
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
     
   'Valida Año del Periodo
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
   
                 
   Screen.MousePointer = 11
   
   vlArchivo = App.Path & "\Reportes\PP_Rpt_ControlRetJud1.rpt"   '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
      MsgBox "Archivo de Reporte de Control de Retención Judicial no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
      Screen.MousePointer = 0
      Exit Sub
   End If
   
'''   vlFechaInicio = Format(CDate(Trim(Txt_FechaIni.Text)), "yyyymmdd")
'''   vlFechaTermino = Format(CDate(Trim(Txt_FechaTer.Text)), "yyyymmdd")
'''
'''   vgQuery = ""
'''   vgQuery = vgQuery & "{PP_TMAE_RETJUDICIAL.fec_iniret}>= '" & Trim(vlFechaInicio) & "' AND "
'''   vgQuery = vgQuery & "{PP_TMAE_RETJUDICIAL.fec_iniret}<= '" & Trim(vlFechaTermino) & "' "
   
 
   Rpt_General.Reset
   Rpt_General.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
   Rpt_General.Connect = vgRutaDataBase       ' o App.Path & "\Nestle.mdb"
'''   Rpt_General.SelectionFormula = vgQuery
      
'''   vgPalabra = Txt_FechaIni.Text & " - " & Txt_FechaTer.Text
   
'''   Rpt_General.Formulas(0) = "Periodo = '" & vgPalabra & "'"
   Rpt_General.Formulas(1) = ""
   Rpt_General.Formulas(2) = ""
      
   Rpt_General.Formulas(3) = "NombreCompania = '" & vgNombreCompania & "'"
   Rpt_General.Formulas(4) = "NombreSistema= '" & vgNombreSistema & "'"
   Rpt_General.Formulas(5) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
      
   Rpt_General.WindowState = crptMaximized
   Rpt_General.Destination = crptToWindow
   Rpt_General.WindowTitle = "Informe de Control de Retención Judicial"
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

    Frm_InfoControlRetJudicial.Top = 0
    Frm_InfoControlRetJudicial.Left = 0
            
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Txt_Mes_Change()
    If Not IsNumeric(Txt_Mes) Then
       Txt_Mes = ""
    End If
End Sub

Private Sub Txt_Mes_KeyPress(KeyAscii As Integer)
On Error GoTo Err_TxtMesKeyPress

    If KeyAscii = 13 Then
       If Txt_Mes.Text = "" Then
          MsgBox "Debe ingresar Mes del Periodo de Pago.", vbCritical, "Error de Datos"
          Txt_Mes.SetFocus
          Exit Sub
       End If
         
       If CDbl(Txt_Mes.Text) <= 0 Or CDbl(Txt_Mes.Text) > 12 Then
          MsgBox "El Mes ingresado No es un Valor Válido.", vbCritical, "Error de Datos"
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

Private Sub Txt_Anno_Change()
    If Not IsNumeric(Txt_Anno) Then
       Txt_Anno = ""
    End If
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
        cmd_calcular.SetFocus
               
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

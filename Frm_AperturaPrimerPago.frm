VERSION 5.00
Begin VB.Form Frm_AperturaPrimerPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Apertura de Periodo Manual"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   5040
   Begin VB.Frame Frame1 
      Caption         =   "Periodo a Abrir"
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
      Height          =   1215
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4935
      Begin VB.Label Label2 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3720
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Lbl_Periodo 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2040
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Hasta"
         Height          =   255
         Left            =   3120
         TabIndex        =   7
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Polizas recibidas Desde"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Fra_Operaciones 
      Height          =   1095
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   4935
      Begin VB.CommandButton Cmd_Abrir 
         Caption         =   "&Abrir"
         Height          =   675
         Left            =   1680
         Picture         =   "Frm_AperturaPrimerPago.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Abrir Periodo"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   2880
         Picture         =   "Frm_AperturaPrimerPago.frx":0822
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
   End
End
Attribute VB_Name = "Frm_AperturaPrimerPago"
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
    
    Frm_AperturaPrimerPago.Left = 0
    Frm_AperturaPrimerPago.Top = 0
    
Exit Sub
Err_Carga:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub

VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Frm_BuscaCauInv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda de Causal de Invalidez"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   5520
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   5295
      Begin VB.CommandButton Btn_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   3000
         Picture         =   "Frm_BuscaCauInv.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Btn_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   1560
         Picture         =   "Frm_BuscaCauInv.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.TextBox Txt_CodCauInv 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Txt_DescCodCauInv 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label Lbl_Buscador 
         Caption         =   "Código"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Lbl_Buscador 
         Caption         =   "Descripción"
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   5
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Msf_GrillaBuscaCauInv 
      Height          =   1515
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1440
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2672
      _Version        =   393216
      BackColor       =   14745599
   End
   Begin VB.Label Lbl_Buscador 
      Caption         =   "Resultado Búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   1815
   End
End
Attribute VB_Name = "Frm_BuscaCauInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vlSql As String

Dim vlCodigo As String
Dim vlDescripcion As String
Dim vlCauInv As String
Dim vlFila As Integer

Function flLimpiar()
On Error GoTo Err_Limpia

    Txt_CodCauInv.Text = ""
    Txt_DescCodCauInv.Text = ""
    Txt_CodCauInv.SetFocus
    
Exit Function
Err_Limpia:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flCargaGrilla()
Dim vlCuenta As Integer
Dim vlColumna As Integer
Dim vlCodigo As String
Dim vlDescripcion As String

'On Error GoTo Err_Carga
    Msf_GrillaBuscaCauInv.Rows = 1
    
    Msf_GrillaBuscaCauInv.Enabled = True
    Msf_GrillaBuscaCauInv.Cols = 3
    Msf_GrillaBuscaCauInv.Rows = 1
    Msf_GrillaBuscaCauInv.Row = 0
    
    Msf_GrillaBuscaCauInv.Col = 0
    Msf_GrillaBuscaCauInv.ColWidth(0) = 0
    
    Msf_GrillaBuscaCauInv.Col = 1
    Msf_GrillaBuscaCauInv.CellAlignment = 4
    Msf_GrillaBuscaCauInv.ColWidth(1) = 800
    Msf_GrillaBuscaCauInv.Text = "Código"
    Msf_GrillaBuscaCauInv.CellFontBold = True
        
    Msf_GrillaBuscaCauInv.Col = 2
    Msf_GrillaBuscaCauInv.ColWidth(2) = 4000
    Msf_GrillaBuscaCauInv.CellAlignment = 4
    Msf_GrillaBuscaCauInv.Text = "Descripción"
    Msf_GrillaBuscaCauInv.CellFontBold = True
        
Exit Function
Err_Carga:
      Screen.MousePointer = 0
      Select Case Err
        Case Else
          MsgBox "Error grave [" & Err & Space(4) & Err.Description & "]", vbCritical
      End Select
End Function

Private Sub Btn_Salir_Click()
On Error GoTo Err_Volver
'    If vgFormulario = "P" Then
'        Frm_CalPoliza.Enabled = True
'    End If
    Unload Me
Exit Sub
Err_Volver:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Private Sub Btn_Limpiar_Click()
    flLimpiar
End Sub
Private Sub Form_Load()
On Error GoTo Err_Cargar

    Me.Top = 0
    Me.Left = 0
    
    Call flCargaGrilla
    
    Txt_CodCauInv.Text = ""
    Txt_DescCodCauInv.Text = ""
    
Exit Sub
Err_Cargar:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Unload
'    If vgFormulario = "P" Then
    Frm_EndosoPol.Enabled = True
'    End If
Exit Sub
Err_Unload:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

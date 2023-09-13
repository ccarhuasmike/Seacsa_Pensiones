VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Frm_AFInformeEstFin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe Estadistico y Financiero"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   6090
   Begin VB.Frame Frame1 
      Height          =   1035
      Left            =   90
      TabIndex        =   12
      Top             =   2295
      Width           =   5850
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   1200
         Picture         =   "Frm_AFInformeEstFin.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   3630
         Picture         =   "Frm_AFInformeEstFin.frx":06BA
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   2415
         Picture         =   "Frm_AFInformeEstFin.frx":07B4
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   720
      End
      Begin Crystal.CrystalReport Rpt_Reporte 
         Left            =   5175
         Top             =   210
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.Frame Fra_Busqueda 
      Height          =   2175
      Left            =   90
      TabIndex        =   8
      Top             =   60
      Width           =   5880
      Begin VB.ComboBox Cmb_Tipo 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox Txt_Reintegro 
         Height          =   270
         Left            =   2895
         MaxLength       =   12
         TabIndex        =   4
         Top             =   1710
         Width           =   1455
      End
      Begin VB.TextBox Txt_AporteFiscal 
         Height          =   270
         Left            =   2895
         MaxLength       =   12
         TabIndex        =   3
         Top             =   1260
         Width           =   1455
      End
      Begin VB.TextBox Txt_Anno 
         Height          =   255
         Left            =   3585
         MaxLength       =   4
         TabIndex        =   2
         Top             =   780
         Width           =   795
      End
      Begin VB.TextBox Txt_Mes 
         Height          =   255
         Left            =   2910
         MaxLength       =   2
         TabIndex        =   1
         Top             =   795
         Width           =   360
      End
      Begin VB.Label Label7 
         Caption         =   " :"
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
         Height          =   240
         Left            =   2400
         TabIndex        =   20
         Top             =   1725
         Width           =   165
      End
      Begin VB.Label Label6 
         Caption         =   " :"
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
         Height          =   240
         Left            =   2400
         TabIndex        =   19
         Top             =   1260
         Width           =   165
      End
      Begin VB.Label Label5 
         Caption         =   " :"
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
         Height          =   240
         Left            =   2400
         TabIndex        =   18
         Top             =   825
         Width           =   165
      End
      Begin VB.Label Lbl_Contrato 
         Caption         =   "Tipo de Proceso             :"
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
         Index           =   2
         Left            =   285
         TabIndex        =   17
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "$"
         Height          =   375
         Left            =   2730
         TabIndex        =   16
         Top             =   1740
         Width           =   195
      End
      Begin VB.Label Label3 
         Caption         =   "Reintegros de Asig. Fam."
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
         Left            =   285
         TabIndex        =   15
         Top             =   1740
         Width           =   2145
      End
      Begin VB.Label Label2 
         Caption         =   "$"
         Height          =   375
         Left            =   2715
         TabIndex        =   14
         Top             =   1260
         Width           =   195
      End
      Begin VB.Label Label1 
         Caption         =   "Aporte Fiscal del Mes "
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
         Left            =   285
         TabIndex        =   13
         Top             =   1305
         Width           =   2055
      End
      Begin VB.Label lbl_nombre 
         Caption         =   "Período "
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
         Index           =   8
         Left            =   285
         TabIndex        =   11
         Top             =   780
         Width           =   705
      End
      Begin VB.Label lbl_nombre 
         Caption         =   "(Mes - Año)"
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
         Height          =   270
         Index           =   10
         Left            =   1140
         TabIndex        =   10
         Top             =   795
         Width           =   1125
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
         Left            =   3315
         TabIndex        =   9
         Top             =   795
         Width           =   195
      End
   End
End
Attribute VB_Name = "Frm_AFInformeEstFin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


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
    
    Frm_AFInformeEstFin.Left = 0
    Frm_AFInformeEstFin.Top = 0
    

        
Exit Sub
Err_Carga:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub


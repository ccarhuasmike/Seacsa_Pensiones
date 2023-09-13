VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Frm_PensCertificado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Certificado de Estudios y Soltería"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   9195
   Begin VB.Frame Fra_Botones 
      Height          =   975
      Left            =   120
      TabIndex        =   26
      Top             =   6000
      Width           =   8895
      Begin VB.CommandButton cmd_graba 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   1320
         Picture         =   "Frm_PensCertificado.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton cmd_limpia 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   3720
         Picture         =   "Frm_PensCertificado.frx":06BA
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton cmd_salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   6120
         Picture         =   "Frm_PensCertificado.frx":0D74
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   4920
         Picture         =   "Frm_PensCertificado.frx":0E6E
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Eliminar 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   2520
         Picture         =   "Frm_PensCertificado.frx":1528
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Eliminar Año"
         Top             =   200
         Width           =   720
      End
      Begin Crystal.CrystalReport Rpt_Calculo 
         Left            =   7320
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Póliza / Pensionado"
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
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   8895
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1080
         TabIndex        =   15
         Top             =   360
         Width           =   1185
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   5880
         TabIndex        =   14
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00E0FFFF&
         Height          =   285
         Left            =   1080
         TabIndex        =   13
         Top             =   720
         Width           =   6825
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   3960
         TabIndex        =   12
         Top             =   360
         Width           =   1755
      End
      Begin VB.CommandButton Cmd_Buscar 
         Height          =   375
         Left            =   6360
         Picture         =   "Frm_PensCertificado.frx":186A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Buscar Póliza"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "N° Póliza"
         Height          =   375
         Left            =   360
         TabIndex        =   19
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre"
         Height          =   285
         Left            =   360
         TabIndex        =   18
         Top             =   720
         Width           =   990
      End
      Begin VB.Label Label17 
         Caption         =   "Rut Pensionado"
         Height          =   285
         Left            =   2640
         TabIndex        =   17
         Top             =   360
         Width           =   1365
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "-"
         Height          =   255
         Left            =   5640
         TabIndex        =   16
         Top             =   360
         Width           =   255
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2325
      Left            =   120
      TabIndex        =   9
      Top             =   3600
      Width           =   8970
      _ExtentX        =   15822
      _ExtentY        =   4101
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   14745599
      FormatString    =   $"Frm_PensCertificado.frx":196C
   End
   Begin VB.Frame Frame1 
      Caption         =   "Antecedentes de Certificado de Estudios"
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
      Height          =   2025
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   8925
      Begin VB.TextBox Text6 
         BackColor       =   &H00E0FFFF&
         Height          =   285
         Left            =   6840
         TabIndex        =   25
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0FFFF&
         Height          =   285
         Left            =   6840
         TabIndex        =   23
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1680
         TabIndex        =   22
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Top             =   330
         Width           =   1125
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   3600
         TabIndex        =   5
         Top             =   330
         Width           =   1035
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   1440
         Width           =   2820
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   1080
         Width           =   6705
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha de Ingreso"
         Height          =   255
         Left            =   5280
         TabIndex        =   24
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha de Efecto"
         Height          =   375
         Left            =   5280
         TabIndex        =   21
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Recepción  "
         Height          =   255
         Left            =   165
         TabIndex        =   20
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label13 
         Caption         =   "Vigencia Desde "
         Height          =   210
         Left            =   180
         TabIndex        =   8
         Top             =   345
         Width           =   1395
      End
      Begin VB.Label Label11 
         Caption         =   "Hasta"
         Height          =   225
         Left            =   2880
         TabIndex        =   7
         Top             =   360
         Width           =   690
      End
      Begin VB.Label Label10 
         Caption         =   "Régimen de Estudio"
         Height          =   225
         Left            =   165
         TabIndex        =   4
         Top             =   1530
         Width           =   1530
      End
      Begin VB.Label Label9 
         Caption         =   "Institución"
         Height          =   255
         Left            =   165
         TabIndex        =   2
         Top             =   1080
         Width           =   885
      End
   End
End
Attribute VB_Name = "Frm_PensCertificado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Buscar_Click()
    Frm_Busqueda.Show 1
End Sub

Private Sub Command5_Click()
    Unload Me
    
End Sub


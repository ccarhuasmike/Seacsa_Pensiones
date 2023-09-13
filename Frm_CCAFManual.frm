VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Frm_CCAFManual 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Descuentos por Cajas de Compensación                                              "
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   9360
   Begin VB.Frame Fra_Poliza 
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
      Left            =   120
      TabIndex        =   22
      Top             =   0
      Width           =   9165
      Begin VB.TextBox Txt_PenPoliza 
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   0
         Top             =   360
         Width           =   1185
      End
      Begin VB.TextBox Txt_PenDigito 
         Height          =   285
         Left            =   5520
         MaxLength       =   1
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Txt_PenRut 
         Height          =   285
         Left            =   3840
         MaxLength       =   11
         TabIndex        =   1
         Top             =   360
         Width           =   1395
      End
      Begin VB.CommandButton Cmd_BuscarPol 
         Height          =   375
         Left            =   8160
         Picture         =   "Frm_CCAFManual.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Buscar Póliza"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Cmd_Buscar 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8160
         Picture         =   "Frm_CCAFManual.frx":0102
         TabIndex        =   4
         ToolTipText     =   "Buscar"
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Póliza"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   29
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
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
         Index           =   20
         Left            =   5280
         TabIndex        =   28
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Lbl_PenNombre 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   255
         Left            =   1080
         TabIndex        =   27
         Top             =   720
         Width           =   6975
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Rut Pensionado"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   26
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Num. Endoso"
         Height          =   195
         Index           =   42
         Left            =   6360
         TabIndex        =   25
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label Lbl_End 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   7440
         TabIndex        =   24
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   43
         Left            =   120
         TabIndex        =   23
         Top             =   0
         Width           =   1725
      End
   End
   Begin VB.Frame Fra_Botones 
      Height          =   975
      Left            =   120
      TabIndex        =   21
      Top             =   5640
      Width           =   9135
      Begin VB.CommandButton Cmd_Eliminar 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   2400
         Picture         =   "Frm_CCAFManual.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Eliminar Año"
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   3480
         Picture         =   "Frm_CCAFManual.frx":0546
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   6720
         Picture         =   "Frm_CCAFManual.frx":0C00
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   4560
         Picture         =   "Frm_CCAFManual.frx":0CFA
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Grabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   1320
         Picture         =   "Frm_CCAFManual.frx":13B4
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   5640
         Picture         =   "Frm_CCAFManual.frx":1A6E
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   200
         Width           =   730
      End
      Begin Crystal.CrystalReport Rpt_Calculo 
         Left            =   7920
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4425
      Left            =   120
      TabIndex        =   19
      Top             =   1200
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   7805
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Ingreso de Descuentos"
      TabPicture(0)   =   "Frm_CCAFManual.frx":2048
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Lbl_Nombre(3)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Lbl_Nombre(4)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Lbl_FecTermino"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Lbl_Nombre(7)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Lbl_Nombre(8)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Lbl_Nombre(12)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Lbl_Nombre(5)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Msf_GriDoc"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Cmb_ModPago"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Txt_FecInicio"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Txt_NroCuotas"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Txt_MontoCuota"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Txt_MontoTotal"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Fra_Entidad"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Fra_Sus"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Histórico"
      TabPicture(1)   =   "Frm_CCAFManual.frx":2064
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Msf_GriHis"
      Tab(1).ControlCount=   1
      Begin VB.Frame Fra_Sus 
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
         Height          =   735
         Left            =   240
         TabIndex        =   42
         Top             =   3600
         Width           =   8655
         Begin VB.TextBox Txt_FecSus 
            Height          =   285
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   47
            Top             =   255
            Width           =   1185
         End
         Begin VB.ComboBox Cmb_Sus 
            BackColor       =   &H80000018&
            Height          =   315
            Left            =   4440
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   240
            Width           =   4095
         End
         Begin VB.Label Label6 
            Caption         =   "Fecha Suspensión"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   270
            Width           =   1935
         End
         Begin VB.Label Label5 
            Caption         =   "Motivo Suspensión"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3000
            TabIndex        =   44
            Top             =   270
            Width           =   1695
         End
         Begin VB.Label Label4 
            Caption         =   "  Suspensión"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.Frame Fra_Entidad 
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
         Left            =   240
         TabIndex        =   38
         Top             =   360
         Width           =   8775
         Begin VB.ComboBox Cmb_CCAF 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   2880
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   240
            Width           =   5235
         End
         Begin VB.ComboBox Cmb_TipoConcepto 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   2880
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   600
            Width           =   5235
         End
         Begin VB.Label Label3 
            Caption         =   "  Selección de Entidad  "
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
            Left            =   120
            TabIndex        =   41
            Top             =   0
            Width           =   2175
         End
         Begin VB.Label Label2 
            Caption         =   "Caja de Compensación"
            Height          =   255
            Left            =   960
            TabIndex        =   40
            Top             =   255
            Width           =   1875
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo de Concepto"
            Height          =   255
            Left            =   960
            TabIndex        =   39
            Top             =   600
            Width           =   1875
         End
      End
      Begin VB.TextBox Txt_MontoTotal 
         Height          =   285
         Left            =   2280
         MaxLength       =   15
         TabIndex        =   11
         Top             =   3240
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.TextBox Txt_MontoCuota 
         Height          =   285
         Left            =   240
         MaxLength       =   11
         TabIndex        =   10
         Top             =   3240
         Width           =   1305
      End
      Begin VB.TextBox Txt_NroCuotas 
         Height          =   285
         Left            =   1800
         MaxLength       =   5
         TabIndex        =   9
         Top             =   3240
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox Txt_FecInicio 
         Height          =   285
         Left            =   240
         MaxLength       =   10
         TabIndex        =   7
         Top             =   1920
         Width           =   1185
      End
      Begin VB.ComboBox Cmb_ModPago 
         BackColor       =   &H00E0FFFF&
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2520
         Width           =   3465
      End
      Begin MSFlexGridLib.MSFlexGrid Msf_GriHis 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   20
         Top             =   600
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   6376
         _Version        =   393216
         Cols            =   5
         BackColor       =   14745599
         FormatString    =   $"Frm_CCAFManual.frx":2080
      End
      Begin MSFlexGridLib.MSFlexGrid Msf_GriDoc 
         Height          =   2055
         Left            =   4200
         TabIndex        =   12
         Top             =   1560
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   3625
         _Version        =   393216
         Cols            =   6
         BackColor       =   14745599
         FormatString    =   "Desde      |        Hasta |    Modalidad    | Nº Cuotas |Mto. Cuota |Mto. Total"
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Modalidad de Pago"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   37
         Top             =   2280
         Width           =   1785
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Monto Total a Pagar"
         Height          =   255
         Index           =   12
         Left            =   2280
         TabIndex        =   36
         Top             =   3000
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Monto a Descontar"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   35
         Top             =   3000
         Width           =   1785
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Número de Cuotas"
         Height          =   255
         Index           =   7
         Left            =   1680
         TabIndex        =   34
         Top             =   3000
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label Lbl_FecTermino 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1800
         TabIndex        =   33
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Lbl_Nombre 
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
         Index           =   4
         Left            =   1485
         TabIndex        =   32
         Top             =   1920
         Width           =   225
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Fecha del Descuento"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   31
         Top             =   1680
         Width           =   1785
      End
   End
End
Attribute VB_Name = "Frm_CCAFManual"
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
On Error GoTo Err_Cargar

    Frm_CCAFManual.Top = 0
    Frm_CCAFManual.Left = 0
    
    
    SSTab1.Tab = 0
    Fra_Entidad.Enabled = False
    SSTab1.Enabled = True
    
    
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

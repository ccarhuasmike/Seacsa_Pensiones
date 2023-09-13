VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Frm_AFIngresos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Declaración de Ingresos."
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   9120
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   120
      TabIndex        =   31
      Top             =   1200
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Declaración de Ingresos"
      TabPicture(0)   =   "Frm_AFIngresos.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Fra_Efecto"
      Tab(0).Control(1)=   "Fra_Per"
      Tab(0).Control(2)=   "Fra_Detalle"
      Tab(0).Control(3)=   "Fra_Calculo"
      Tab(0).Control(4)=   "Fra_Declaracion"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Historial de Declaraciones de Ingresos"
      TabPicture(1)   =   "Frm_AFIngresos.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "MSF_GrillaHistorico"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Fra_Efecto 
         Caption         =   "  Periodo de Efecto  "
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
         Height          =   855
         Left            =   -74880
         TabIndex        =   52
         Top             =   2400
         Width           =   3615
         Begin VB.Label Label1 
            Caption         =   "  -"
            Height          =   255
            Left            =   1560
            TabIndex        =   55
            Top             =   360
            Width           =   255
         End
         Begin VB.Label lbl_Suspension 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1920
            TabIndex        =   54
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lbl_Efecto 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   360
            TabIndex        =   53
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Fra_Per 
         Caption         =   "  Periodo Vigencia  "
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
         Height          =   855
         Left            =   -74880
         TabIndex        =   47
         Top             =   1440
         Width           =   3615
         Begin VB.Label Lbl_raya 
            Caption         =   "  -"
            Height          =   255
            Left            =   1560
            TabIndex        =   50
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Lbl_FecTer 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1920
            TabIndex        =   49
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Lbl_Fecini 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   360
            TabIndex        =   48
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Fra_Detalle 
         Caption         =   "  Detalle de Declaraciones  de Ingresos  "
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
         Height          =   4815
         Left            =   -71160
         TabIndex        =   36
         Top             =   480
         Width           =   4935
         Begin VB.CommandButton Cmd_CalcularPension 
            Height          =   450
            Left            =   4320
            Picture         =   "Frm_AFIngresos.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Calcular Porcentajes"
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox Txt_MtoPension 
            Height          =   285
            Left            =   1680
            MaxLength       =   12
            TabIndex        =   10
            Top             =   1080
            Width           =   1185
         End
         Begin VB.CommandButton Cmd_LimpiarRango 
            Height          =   450
            Left            =   4320
            Picture         =   "Frm_AFIngresos.frx":05FA
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Limpiar"
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox Txt_MesDec 
            Height          =   285
            Left            =   1680
            MaxLength       =   2
            TabIndex        =   8
            Top             =   360
            Width           =   1185
         End
         Begin VB.TextBox Txt_MtoDec 
            Height          =   285
            Left            =   1680
            MaxLength       =   9
            TabIndex        =   9
            Top             =   720
            Width           =   1185
         End
         Begin VB.CommandButton Cmd_Sumar 
            Height          =   450
            Left            =   3720
            Picture         =   "Frm_AFIngresos.frx":0C6C
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Agregar Beneficiario"
            Top             =   360
            Width           =   495
         End
         Begin VB.CommandButton Cmd_Restar 
            Height          =   450
            Left            =   3720
            Picture         =   "Frm_AFIngresos.frx":0DF6
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Quitar Beneficiario"
            Top             =   840
            Width           =   495
         End
         Begin MSFlexGridLib.MSFlexGrid MSF_GrillaDetDec 
            Height          =   3105
            Left            =   120
            TabIndex        =   37
            Top             =   1560
            Width           =   4635
            _ExtentX        =   8176
            _ExtentY        =   5477
            _Version        =   393216
            Rows            =   1
            Cols            =   3
            BackColor       =   14745599
            FormatString    =   "Mes      |               Pensión R.V.     |           Otros Ingresos"
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Pensión R.V."
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   56
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Detalle de Declaraciones de Ingresos"
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
            Left            =   240
            TabIndex        =   45
            Top             =   0
            Width           =   3375
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Mes Declaración"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   39
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Monto Declarado"
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   38
            Top             =   720
            Width           =   1455
         End
      End
      Begin VB.Frame Fra_Calculo 
         Caption         =   "  Calculo Final  "
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
         Height          =   1935
         Left            =   -74880
         TabIndex        =   33
         Top             =   3360
         Width           =   3615
         Begin VB.CommandButton Cmd_Reliquidar 
            Caption         =   "&Reliquidar ..."
            Height          =   375
            Left            =   2280
            TabIndex        =   51
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox Txt_RtaPromedio 
            BackColor       =   &H00E0FFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   22
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox Txt_ValorCarga 
            BackColor       =   &H00E0FFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   23
            Top             =   720
            Width           =   1455
         End
         Begin VB.CommandButton cmd_calcular 
            Caption         =   "&Calcular"
            Height          =   675
            Left            =   1320
            Picture         =   "Frm_AFIngresos.frx":0F80
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Calcula Valor Carga"
            Top             =   1200
            Width           =   720
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Calculo Final"
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
            Index           =   6
            Left            =   240
            TabIndex        =   43
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label Label16 
            Caption         =   "Valor de Carga"
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Top             =   720
            Width           =   1275
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Renta Promedio"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   34
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Fra_Declaracion 
         Caption         =   "  Año Declaración  "
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
         Height          =   855
         Left            =   -74880
         TabIndex        =   32
         Top             =   480
         Width           =   3615
         Begin VB.TextBox Txt_AnnoDec 
            Height          =   285
            Left            =   960
            MaxLength       =   4
            TabIndex        =   6
            Top             =   360
            Width           =   705
         End
         Begin VB.CommandButton Cmd_BuscarAnno 
            Height          =   495
            Left            =   2280
            Picture         =   "Frm_AFIngresos.frx":1422
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Buscar Póliza"
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Año Declaración"
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
            Index           =   7
            Left            =   240
            TabIndex        =   44
            Top             =   0
            Width           =   1455
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSF_GrillaHistorico 
         Height          =   3960
         Left            =   960
         TabIndex        =   40
         Top             =   600
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   6985
         _Version        =   393216
         Cols            =   6
         BackColor       =   14745599
         FormatString    =   "Año Dec.          |               Renta Promedio          |          Valor de Carga          "
      End
   End
   Begin VB.Frame Fra_Botones 
      Height          =   1095
      Left            =   120
      TabIndex        =   30
      Top             =   6600
      Width           =   8925
      Begin VB.CommandButton Cmd_ImpHistorico 
         Caption         =   "&Historico"
         Height          =   675
         Left            =   4320
         Picture         =   "Frm_AFIngresos.frx":1524
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   6240
         Picture         =   "Frm_AFIngresos.frx":1BDE
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   730
      End
      Begin VB.CommandButton Cmd_Eliminar 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   2400
         Picture         =   "Frm_AFIngresos.frx":21B8
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Eliminar Año"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Detalle"
         Height          =   675
         Left            =   3360
         Picture         =   "Frm_AFIngresos.frx":24FA
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   7200
         Picture         =   "Frm_AFIngresos.frx":2BB4
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   5280
         Picture         =   "Frm_AFIngresos.frx":2CAE
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Grabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   1440
         Picture         =   "Frm_AFIngresos.frx":3368
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   720
      End
      Begin Crystal.CrystalReport Rpt_Calculo 
         Left            =   8160
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.Frame Fra_Poliza 
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
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8895
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
         Picture         =   "Frm_AFIngresos.frx":3A22
         TabIndex        =   5
         ToolTipText     =   "Buscar"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Txt_PenPoliza 
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   1
         Top             =   360
         Width           =   1185
      End
      Begin VB.TextBox Txt_PenDigito 
         Height          =   285
         Left            =   6000
         MaxLength       =   1
         TabIndex        =   3
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox Txt_PenRut 
         Height          =   285
         Left            =   3960
         MaxLength       =   10
         TabIndex        =   2
         Top             =   360
         Width           =   1755
      End
      Begin VB.CommandButton Cmd_BuscarPol 
         Height          =   375
         Left            =   8160
         Picture         =   "Frm_AFIngresos.frx":3B24
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Buscar Póliza"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Lbl_Nombre 
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   46
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nº Endoso"
         Height          =   255
         Index           =   12
         Left            =   6480
         TabIndex        =   42
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Lbl_Endoso 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7440
         TabIndex        =   41
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Lbl_PenNombre 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   29
         Top             =   720
         Width           =   6855
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Póliza"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   27
         Top             =   720
         Width           =   855
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
         Left            =   5760
         TabIndex        =   25
         Top             =   360
         Width           =   255
      End
   End
End
Attribute VB_Name = "Frm_AFIngresos"
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

    Frm_AFIngresos.Top = 0
    Frm_AFIngresos.Left = 0
    SSTab1.Enabled = True
    SSTab1.Tab = 0
    Fra_Declaracion.Enabled = False
    
    Fra_Calculo.Enabled = False
    Fra_Detalle.Enabled = False
    
    
    vgCodIndCalDecIng = flObtenerCodIndCalDecIng
    
    If vgCodIndCalDecIng = "S" Then
        Txt_MtoPension.Enabled = True
        Cmd_CalcularPension.Enabled = True
    Else
        Txt_MtoPension.Enabled = False
        Cmd_CalcularPension.Enabled = False
    End If
    
  
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub


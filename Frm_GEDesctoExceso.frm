VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Frm_GEDesctoExceso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Haberes y Descuentos de Garantía Estatal"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   9105
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
      TabIndex        =   37
      Top             =   0
      Width           =   8895
      Begin VB.CommandButton Cmd_BuscarPol 
         Height          =   375
         Left            =   7680
         Picture         =   "Frm_GEDesctoExceso.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Buscar Póliza"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Txt_PenRut 
         Height          =   285
         Left            =   4080
         TabIndex        =   2
         Top             =   360
         Width           =   1395
      End
      Begin VB.TextBox Txt_PenDigito 
         Height          =   285
         Left            =   5760
         TabIndex        =   3
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox Txt_PenPoliza 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   360
         Width           =   1185
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
         Left            =   7680
         Picture         =   "Frm_GEDesctoExceso.frx":0102
         TabIndex        =   5
         ToolTipText     =   "Buscar"
         Top             =   600
         Width           =   615
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
         Left            =   5520
         TabIndex        =   45
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Rut Pensionado"
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   44
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   43
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Póliza"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   42
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Lbl_End 
         Caption         =   "Endoso"
         Height          =   255
         Left            =   6360
         TabIndex        =   41
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Lbl_Endoso 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   7080
         TabIndex        =   40
         Top             =   360
         Width           =   375
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
         Index           =   11
         Left            =   120
         TabIndex        =   39
         Top             =   0
         Width           =   1935
      End
      Begin VB.Label Lbl_PenNombre 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   38
         Top             =   720
         Width           =   6135
      End
   End
   Begin VB.Frame Fra_Botones 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   5880
      Width           =   8925
      Begin VB.CommandButton Cmd_Grabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   1320
         Picture         =   "Frm_GEDesctoExceso.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   4560
         Picture         =   "Frm_GEDesctoExceso.frx":08BE
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   6720
         Picture         =   "Frm_GEDesctoExceso.frx":0F78
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   3480
         Picture         =   "Frm_GEDesctoExceso.frx":1072
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Eliminar 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   2400
         Picture         =   "Frm_GEDesctoExceso.frx":172C
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Eliminar Año"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   5640
         Picture         =   "Frm_GEDesctoExceso.frx":1A6E
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   730
      End
      Begin Crystal.CrystalReport Rpt_HabDes 
         Left            =   7320
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   120
      TabIndex        =   21
      Top             =   1200
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   8281
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Haberes y Desctos."
      TabPicture(0)   =   "Frm_GEDesctoExceso.frx":2048
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Fra_HaberDescto"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Historia"
      TabPicture(1)   =   "Frm_GEDesctoExceso.frx":2064
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Msf_Grilla"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Fra_HaberDescto 
         Height          =   4215
         Left            =   -74880
         TabIndex        =   22
         Top             =   360
         Width           =   8655
         Begin VB.TextBox Txt_MontoTotal 
            Height          =   285
            Left            =   2160
            MaxLength       =   15
            TabIndex        =   11
            Top             =   2040
            Width           =   1305
         End
         Begin VB.TextBox Txt_MontoCuota 
            Height          =   285
            Left            =   2160
            MaxLength       =   13
            TabIndex        =   10
            Top             =   1680
            Width           =   1305
         End
         Begin VB.TextBox Txt_NroCuotas 
            Height          =   285
            Left            =   2160
            MaxLength       =   3
            TabIndex        =   9
            Top             =   1320
            Width           =   585
         End
         Begin VB.TextBox Txt_FecInicio 
            Height          =   285
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   7
            Top             =   600
            Width           =   1185
         End
         Begin VB.Frame Fra_suspension 
            Caption         =   "  Suspensión Haberes / Descuentos"
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
            Height          =   1695
            Left            =   240
            TabIndex        =   23
            Top             =   2400
            Width           =   8175
            Begin VB.TextBox Txt_Observacion 
               Height          =   615
               Left            =   1920
               MaxLength       =   255
               MultiLine       =   -1  'True
               TabIndex        =   14
               Top             =   960
               Width           =   5895
            End
            Begin VB.TextBox Txt_FecSuspension 
               Height          =   285
               Left            =   1920
               MaxLength       =   10
               TabIndex        =   13
               Top             =   600
               Width           =   1140
            End
            Begin VB.ComboBox Cmb_MotSuspension 
               BackColor       =   &H00E0FFFF&
               Height          =   315
               Left            =   1920
               Style           =   2  'Dropdown List
               TabIndex        =   12
               Top             =   240
               Width           =   5895
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "Observación  "
               Height          =   255
               Index           =   14
               Left            =   240
               TabIndex        =   27
               Top             =   1080
               Width           =   1335
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "Fecha de Suspensión"
               Height          =   255
               Index           =   10
               Left            =   240
               TabIndex        =   26
               Top             =   720
               Width           =   1710
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "Motivo de Suspensión"
               Height          =   255
               Index           =   9
               Left            =   240
               TabIndex        =   25
               Top             =   360
               Width           =   1800
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "Suspensión Haberes / Descuentos"
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
               Index           =   13
               Left            =   240
               TabIndex        =   24
               Top             =   0
               Width           =   3105
            End
         End
         Begin VB.ComboBox Cmb_Moneda 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   960
            Width           =   6105
         End
         Begin VB.ComboBox Cmb_HabDes 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   240
            Width           =   6105
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Monto Total "
            Height          =   255
            Index           =   12
            Left            =   360
            TabIndex        =   35
            Top             =   2040
            Width           =   1785
         End
         Begin VB.Label Lbl_FecTermino 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3600
            TabIndex        =   34
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Monto de la Cuota"
            Height          =   255
            Index           =   8
            Left            =   360
            TabIndex        =   33
            Top             =   1680
            Width           =   1785
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Número de Cuotas"
            Height          =   255
            Index           =   7
            Left            =   360
            TabIndex        =   32
            Top             =   1320
            Width           =   1785
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Código Moneda"
            Height          =   255
            Index           =   5
            Left            =   360
            TabIndex        =   31
            Top             =   960
            Width           =   1785
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
            Left            =   3360
            TabIndex        =   30
            Top             =   600
            Width           =   225
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Fecha de Inicio"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   29
            Top             =   600
            Width           =   1785
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Código Haber/Descto."
            Height          =   255
            Index           =   6
            Left            =   360
            TabIndex        =   28
            Top             =   240
            Width           =   1785
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Msf_Grilla 
         Height          =   3615
         Left            =   240
         TabIndex        =   36
         Top             =   600
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   6376
         _Version        =   393216
         Cols            =   12
         BackColor       =   14745599
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "Frm_GEDesctoExceso"
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

    Frm_GEDesctoExceso.Top = 0
    Frm_GEDesctoExceso.Left = 0
    vlModFecTerNumCuotas = True
    SSTab1.Tab = 0
    
    SSTab1.Enabled = True
    
        
   
    
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub


VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Frm_CtaCorriente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuenta Corriente Pensionado"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   9090
   Begin TabDlg.SSTab SSTab_CtaCte 
      Height          =   5280
      Left            =   120
      TabIndex        =   8
      Top             =   1695
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9313
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Totales Finales"
      TabPicture(0)   =   "Frm_CtaCorriente.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Fra_TotalesEst"
      Tab(0).Control(1)=   "Fra_TotalesCia"
      Tab(0).Control(2)=   "Fra_Resumen"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Detalle de Cta. Corriente"
      TabPicture(1)   =   "Frm_CtaCorriente.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Msf_GrillaCtaCte"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Fra_TotalesEst 
         Caption         =   " Totales por Estado   "
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
         Left            =   -70320
         TabIndex        =   50
         Top             =   1920
         Width           =   3975
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Recupero Estado"
            Height          =   255
            Index           =   36
            Left            =   240
            TabIndex        =   60
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Bono Invierno"
            Height          =   255
            Index           =   37
            Left            =   240
            TabIndex        =   59
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Reintegro p/ Estado"
            Height          =   255
            Index           =   38
            Left            =   240
            TabIndex        =   58
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Reintegro a/ Estado"
            Height          =   255
            Index           =   39
            Left            =   240
            TabIndex        =   57
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Total Mensual"
            Height          =   255
            Index           =   40
            Left            =   240
            TabIndex        =   56
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Label Lbl_MtoRecEst 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2280
            TabIndex        =   55
            Top             =   360
            Width           =   1305
         End
         Begin VB.Label Lbl_BonInvEst 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2280
            TabIndex        =   54
            Top             =   720
            Width           =   1305
         End
         Begin VB.Label Lbl_ReiPEst 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2280
            TabIndex        =   53
            Top             =   1080
            Width           =   1305
         End
         Begin VB.Label Lbl_ReiAEst 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2280
            TabIndex        =   52
            Top             =   1440
            Width           =   1305
         End
         Begin VB.Label Lbl_TotalMensual 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2280
            TabIndex        =   51
            Top             =   1800
            Width           =   1305
         End
      End
      Begin VB.Frame Fra_TotalesCia 
         Caption         =   " Totales por Compañía  "
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
         Left            =   -74760
         TabIndex        =   35
         Top             =   1920
         Width           =   3975
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Total a la Fecha"
            Height          =   255
            Index           =   15
            Left            =   240
            TabIndex        =   49
            Top             =   2520
            Width           =   1575
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Otros Descuentos GE"
            Height          =   255
            Index           =   14
            Left            =   240
            TabIndex        =   48
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Otros Haberes GE"
            Height          =   255
            Index           =   13
            Left            =   240
            TabIndex        =   47
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Bono Invierno"
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   46
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "GE RetroActiva"
            Height          =   255
            Index           =   11
            Left            =   240
            TabIndex        =   45
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Garantía Estatal"
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   44
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Monto Pensión"
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   43
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Lbl_MtoPension 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2160
            TabIndex        =   42
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Lbl_MtoGarEst 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2160
            TabIndex        =   41
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Lbl_MtoGERetro 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2160
            TabIndex        =   40
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Lbl_BonInvCia 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2160
            TabIndex        =   39
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Lbl_MtoOtrosHab 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2160
            TabIndex        =   38
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label Lbl_MtoOtrosDesctos 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2160
            TabIndex        =   37
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label Lbl_TotalFecha 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2160
            TabIndex        =   36
            Top             =   2520
            Width           =   1335
         End
      End
      Begin VB.Frame Fra_Resumen 
         Caption         =   " Resumen "
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
         Height          =   1320
         Left            =   -74760
         TabIndex        =   25
         Top             =   480
         Width           =   8415
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Total Garantía Estatal Pagada por la Cía."
            Height          =   495
            Index           =   3
            Left            =   360
            TabIndex        =   34
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Total Garantía Estatal Recuperada"
            Height          =   495
            Index           =   4
            Left            =   2640
            TabIndex        =   33
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Por Recuperar"
            Height          =   255
            Index           =   5
            Left            =   4800
            TabIndex        =   32
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Porcentaje de Recupero"
            Height          =   495
            Index           =   6
            Left            =   6960
            TabIndex        =   31
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "%"
            Height          =   255
            Index           =   7
            Left            =   7680
            TabIndex        =   30
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Lbl_MtoGEPagCia 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   360
            TabIndex        =   29
            Top             =   360
            Width           =   1665
         End
         Begin VB.Label Lbl_MtoGERec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2640
            TabIndex        =   28
            Top             =   360
            Width           =   1665
         End
         Begin VB.Label Lbl_MtoRec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4800
            TabIndex        =   27
            Top             =   360
            Width           =   1665
         End
         Begin VB.Label Lbl_PrcRec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6960
            TabIndex        =   26
            Top             =   360
            Width           =   615
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Msf_GrillaCtaCte 
         Height          =   4695
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   8281
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   14745599
         AllowBigSelection=   0   'False
         FocusRect       =   2
         SelectionMode   =   1
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
   End
   Begin VB.Frame Fra_TipoCtaCte 
      Height          =   615
      Left            =   120
      TabIndex        =   22
      Top             =   0
      Width           =   8895
      Begin VB.OptionButton Opt_Pensionado 
         Caption         =   "Por Pensionado"
         Height          =   255
         Left            =   2880
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton Opt_General 
         Caption         =   "General"
         Height          =   255
         Left            =   4920
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Lbl_Nombre 
         Alignment       =   2  'Center
         Caption         =   "Tipo de Cuenta Corriente"
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
         Index           =   17
         Left            =   3120
         TabIndex        =   23
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.Frame Fra_Botones 
      Height          =   975
      Left            =   120
      TabIndex        =   21
      Top             =   6960
      Width           =   8895
      Begin VB.CommandButton Cmd_Calcular 
         Caption         =   "&Calcular"
         Height          =   675
         Left            =   2280
         Picture         =   "Frm_CtaCorriente.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Realizar Cálculo de Totales"
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5880
         Picture         =   "Frm_CtaCorriente.frx":04DA
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   3480
         Picture         =   "Frm_CtaCorriente.frx":05D4
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   4680
         Picture         =   "Frm_CtaCorriente.frx":0C8E
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   200
         Width           =   720
      End
      Begin Crystal.CrystalReport Rpt_General 
         Left            =   8280
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
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
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   650
      Width           =   8895
      Begin VB.CommandButton Cmd_Buscar 
         BackColor       =   &H00000000&
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
         Picture         =   "Frm_CtaCorriente.frx":1268
         TabIndex        =   7
         ToolTipText     =   "Buscar"
         Top             =   520
         Width           =   615
      End
      Begin VB.TextBox Txt_PenPoliza 
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   3
         Top             =   240
         Width           =   1185
      End
      Begin VB.TextBox Txt_PenDigito 
         Height          =   285
         Left            =   5880
         MaxLength       =   1
         TabIndex        =   5
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Txt_PenRut 
         Height          =   285
         Left            =   3720
         MaxLength       =   10
         TabIndex        =   4
         Top             =   240
         Width           =   1755
      End
      Begin VB.CommandButton Cmd_BuscarPol 
         Height          =   375
         Left            =   8160
         Picture         =   "Frm_CtaCorriente.frx":136A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Buscar Póliza"
         Top             =   150
         Width           =   615
      End
      Begin VB.Label Lbl_PenNombre 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   20
         Top             =   580
         Width           =   6855
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Póliza"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   18
         Top             =   580
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Rut Pensionado"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   17
         Top             =   240
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
         Left            =   5520
         TabIndex        =   16
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Lbl_PenEndoso 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7425
         TabIndex        =   15
         Top             =   240
         Width           =   525
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nº Endoso"
         Height          =   240
         Index           =   33
         Left            =   6480
         TabIndex        =   14
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "  Póliza / Pensionado"
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
         Index           =   9
         Left            =   120
         TabIndex        =   13
         Top             =   0
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Frm_CtaCorriente"
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

    Frm_CtaCorriente.Top = 0
    Frm_CtaCorriente.Left = 0
    
    Fra_TipoCtaCte.Enabled = True
    
    Fra_Poliza.Enabled = True
    
    SSTab_CtaCte.Tab = 0
    SSTab_CtaCte.Enabled = True
    
    
    vlSwCalcular = False
    
    
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub







VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Frm_GESeguimiento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seguimiento del Beneficio por Pensionado."
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   9855
   Begin TabDlg.SSTab SSTab1 
      Height          =   5760
      Left            =   120
      TabIndex        =   39
      Top             =   1200
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   10160
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Documentación"
      TabPicture(0)   =   "Frm_GESeguimiento.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Fra_Estado"
      Tab(0).Control(1)=   "Fra_Resolucion"
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(3)=   "Fra_Documentacion"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Estados"
      TabPicture(1)   =   "Frm_GESeguimiento.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MSF_GrillaEstado"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Resoluciones"
      TabPicture(2)   =   "Frm_GESeguimiento.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "MSF_GrillaResolucion"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Fra_Estado 
         Caption         =   "Estado del beneficio "
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
         Left            =   -74880
         TabIndex        =   57
         Top             =   3480
         Width           =   9330
         Begin VB.CommandButton Cmd_Reliquidar 
            Caption         =   "&Reliquidar ..."
            Height          =   375
            Left            =   8040
            TabIndex        =   76
            Top             =   600
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txt_IniValidez 
            Height          =   285
            Left            =   6960
            MaxLength       =   10
            TabIndex        =   17
            Top             =   600
            Width           =   975
         End
         Begin VB.CommandButton Cmd_LmpEst 
            Height          =   330
            Left            =   8760
            Picture         =   "Frm_GESeguimiento.frx":0054
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   600
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton Cmd_EliEstado 
            Height          =   330
            Left            =   8760
            Picture         =   "Frm_GESeguimiento.frx":06C6
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Eliminar Año"
            Top             =   240
            Width           =   375
         End
         Begin VB.ComboBox Cmb_Estado 
            BackColor       =   &H80000018&
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   240
            Width           =   3810
         End
         Begin VB.ComboBox Cmb_Suspension 
            BackColor       =   &H80000018&
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   600
            Width           =   3810
         End
         Begin VB.Label Lbl_UltEstado 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   135
            Left            =   2520
            TabIndex        =   74
            Top             =   120
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Label Lbl_FecIngEstado 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   8400
            TabIndex        =   73
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Estado del Beneficio"
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
            Height          =   195
            Left            =   120
            TabIndex        =   65
            Top             =   0
            Width           =   1890
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Estado"
            Height          =   285
            Index           =   5
            Left            =   120
            TabIndex        =   63
            Top             =   240
            Width           =   570
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Estado de Suspensión"
            Height          =   225
            Index           =   6
            Left            =   120
            TabIndex        =   61
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Válido Desde"
            Height          =   255
            Index           =   7
            Left            =   5760
            TabIndex        =   60
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Fecha Efecto"
            Height          =   255
            Index           =   12
            Left            =   5760
            TabIndex        =   59
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Lbl_Efecto 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6960
            TabIndex        =   58
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Fra_Resolucion 
         Caption         =   "Resoluciones aprobadas por el Estado"
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
         Left            =   -74880
         TabIndex        =   66
         Top             =   4560
         Width           =   9345
         Begin VB.CommandButton Cmd_EliResol 
            Height          =   330
            Left            =   8760
            Picture         =   "Frm_GESeguimiento.frx":0A08
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton Cmd_LmpResol 
            Height          =   330
            Left            =   8760
            Picture         =   "Frm_GESeguimiento.frx":0D4A
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Limpia Datos y Trae Porc.Deducción "
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox Txt_NroRes1 
            Height          =   285
            Left            =   2085
            MaxLength       =   5
            TabIndex        =   20
            Top             =   240
            Width           =   630
         End
         Begin VB.TextBox Txt_NroRes2 
            Height          =   285
            Left            =   2760
            MaxLength       =   6
            TabIndex        =   21
            Top             =   240
            Width           =   690
         End
         Begin VB.TextBox Txt_NroRes3 
            Height          =   285
            Left            =   3480
            MaxLength       =   4
            TabIndex        =   22
            Top             =   240
            Width           =   480
         End
         Begin VB.TextBox Txt_FecResolucion 
            Height          =   285
            Left            =   6760
            MaxLength       =   10
            TabIndex        =   23
            Top             =   240
            Width           =   1140
         End
         Begin VB.TextBox Txt_FecVigencia 
            Height          =   285
            Left            =   2085
            MaxLength       =   10
            TabIndex        =   24
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Lbl_PorcDedCal 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6760
            TabIndex        =   75
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Resoluciones Aprobadas por el Estado"
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
            Height          =   195
            Left            =   120
            TabIndex        =   72
            Top             =   0
            Width           =   3405
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "% de Deducción Calculado"
            Height          =   270
            Index           =   11
            Left            =   4560
            TabIndex        =   71
            Top             =   600
            Width           =   1995
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "N° de Resolución"
            Height          =   225
            Index           =   8
            Left            =   120
            TabIndex        =   70
            Top             =   240
            Width           =   1515
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Fecha de Resolución"
            Height          =   270
            Index           =   10
            Left            =   4560
            TabIndex        =   69
            Top             =   240
            Width           =   1590
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Fecha Inicio Gtía Estatal"
            Height          =   240
            Index           =   9
            Left            =   120
            TabIndex        =   68
            Top             =   600
            Width           =   2175
         End
      End
      Begin VB.Frame Frame1 
         Height          =   615
         Left            =   -74880
         TabIndex        =   49
         Top             =   360
         Width           =   9375
         Begin VB.Label Lbl_TipoPension 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   4095
         End
         Begin VB.Label Lbl_CodParentesco 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4560
            TabIndex        =   52
            Top             =   240
            Width           =   4455
         End
         Begin VB.Label Lbl_Parentesco 
            Caption         =   " Parentesco"
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
            Left            =   4560
            TabIndex        =   55
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label Lbl_Pension 
            Caption         =   " Tipo de Pensión"
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
            Left            =   120
            TabIndex        =   53
            Top             =   0
            Width           =   1575
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSF_GrillaEstado 
         Height          =   4815
         Left            =   -74760
         TabIndex        =   43
         Top             =   720
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   8493
         _Version        =   393216
         Cols            =   7
         BackColor       =   14745599
         FormatString    =   $"Frm_GESeguimiento.frx":13BC
      End
      Begin VB.Frame Fra_Documentacion 
         Caption         =   "Documentación"
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
         Height          =   2445
         Left            =   -74880
         TabIndex        =   40
         Top             =   960
         Width           =   9360
         Begin VB.TextBox Txt_FecBono 
            Height          =   285
            Left            =   7275
            MaxLength       =   10
            TabIndex        =   13
            Top             =   1440
            Width           =   1200
         End
         Begin VB.TextBox Txt_FecFotocopia 
            Height          =   285
            Left            =   7275
            MaxLength       =   10
            TabIndex        =   12
            Top             =   1200
            Width           =   1200
         End
         Begin VB.TextBox Txt_FecCertificado 
            Height          =   285
            Left            =   7275
            MaxLength       =   10
            TabIndex        =   11
            Top             =   960
            Width           =   1200
         End
         Begin VB.TextBox Txt_FecDocumento 
            Height          =   285
            Left            =   7275
            MaxLength       =   10
            TabIndex        =   10
            Top             =   720
            Width           =   1200
         End
         Begin VB.TextBox Txt_FecSaldo 
            Height          =   285
            Left            =   7275
            MaxLength       =   10
            TabIndex        =   8
            Top             =   480
            Width           =   1200
         End
         Begin VB.TextBox Txt_FecSolicitud 
            Height          =   285
            Left            =   7275
            MaxLength       =   10
            TabIndex        =   6
            Top             =   240
            Width           =   1200
         End
         Begin VB.CheckBox Chk_Bono 
            Caption         =   "Solicitud de Bono de Invierno"
            Height          =   255
            Left            =   525
            TabIndex        =   67
            Top             =   1440
            Width           =   2730
         End
         Begin VB.TextBox Txt_AnnoCotiz 
            Height          =   285
            Left            =   3960
            MaxLength       =   3
            TabIndex        =   9
            Top             =   720
            Width           =   1230
         End
         Begin VB.TextBox Txt_SaldoCartola 
            Height          =   285
            Left            =   3960
            MaxLength       =   11
            TabIndex        =   7
            Top             =   480
            Width           =   1230
         End
         Begin VB.CheckBox Chk_Saldo 
            Caption         =   "Saldo de Cartola"
            Height          =   210
            Left            =   525
            TabIndex        =   54
            Top             =   480
            Width           =   1935
         End
         Begin VB.CheckBox Chk_Certificados 
            Caption         =   "Certificados Civiles "
            Height          =   210
            Left            =   525
            TabIndex        =   62
            Top             =   960
            Width           =   2760
         End
         Begin VB.CheckBox Chk_Fotocopia 
            Caption         =   "Fotocopia de Carnet de Identidad"
            Height          =   255
            Left            =   525
            TabIndex        =   64
            Top             =   1200
            Width           =   2730
         End
         Begin VB.CheckBox Chk_Solicitud 
            Caption         =   "Solicitud de Garantía Estatal"
            Height          =   195
            Left            =   525
            TabIndex        =   51
            Top             =   240
            Width           =   6750
         End
         Begin VB.TextBox Txt_Observacion 
            Height          =   555
            Left            =   1080
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   14
            Top             =   1800
            Width           =   7995
         End
         Begin VB.CheckBox Chk_Documento 
            Caption         =   "Documento demuestra años de cotización "
            Height          =   195
            Left            =   525
            TabIndex        =   56
            Top             =   720
            Width           =   3360
         End
         Begin VB.Label Label3 
            Caption         =   "Documentación"
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
            Height          =   195
            Left            =   120
            TabIndex        =   47
            Top             =   0
            Width           =   1410
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Observación"
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   42
            Top             =   1800
            Width           =   945
         End
         Begin VB.Label Label18 
            Caption         =   "Fecha de Recepción"
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
            Height          =   240
            Left            =   6960
            TabIndex        =   41
            Top             =   0
            Width           =   1785
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSF_GrillaResolucion 
         Height          =   4815
         Left            =   240
         TabIndex        =   44
         Top             =   720
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   8493
         _Version        =   393216
         Cols            =   7
         BackColor       =   14745599
         FormatString    =   "Fecha Inicio   |Fecha Termino   |Tipo Res.   |Nº             |Año           | Fecha Resol.        |% Deducción           "
      End
   End
   Begin VB.Frame Fra_Botones 
      Height          =   975
      Left            =   120
      TabIndex        =   38
      Top             =   6960
      Width           =   9645
      Begin VB.CommandButton Cmd_Grabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   2400
         Picture         =   "Frm_GESeguimiento.frx":144E
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   4560
         Picture         =   "Frm_GESeguimiento.frx":1B08
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   6720
         Picture         =   "Frm_GESeguimiento.frx":21C2
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   3480
         Picture         =   "Frm_GESeguimiento.frx":22BC
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Eliminar 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   240
         Picture         =   "Frm_GESeguimiento.frx":2976
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Eliminar Año"
         Top             =   200
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   5640
         Picture         =   "Frm_GESeguimiento.frx":2CB8
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   200
         Width           =   730
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
      Width           =   9615
      Begin VB.CommandButton Cmd_BuscarPol 
         Height          =   375
         Left            =   8640
         Picture         =   "Frm_GESeguimiento.frx":3292
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Buscar Póliza"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Txt_PenRut 
         Height          =   285
         Left            =   4200
         TabIndex        =   2
         Top             =   360
         Width           =   1755
      End
      Begin VB.TextBox Txt_PenDigito 
         Height          =   285
         Left            =   6240
         TabIndex        =   3
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox Txt_PenPoliza 
         Height          =   285
         Left            =   1080
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
         Left            =   8640
         Picture         =   "Frm_GESeguimiento.frx":3394
         TabIndex        =   5
         ToolTipText     =   "Buscar"
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Lbl_PenNombre 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   33
         Top             =   720
         Width           =   7335
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   " Póliza / Pensionado"
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
         Index           =   3
         Left            =   120
         TabIndex        =   48
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nº Endoso"
         Height          =   255
         Index           =   19
         Left            =   7200
         TabIndex        =   46
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Lbl_Endoso 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   8160
         TabIndex        =   45
         Top             =   360
         Width           =   255
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
         Left            =   6000
         TabIndex        =   37
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Rut Pensionado"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   36
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   35
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Póliza"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   34
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "Frm_GESeguimiento"
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

    Frm_GESeguimiento.Top = 0
    Frm_GESeguimiento.Left = 0
        
   SSTab1.Tab = 0
   
   
    vlPasa = True
    
    vlSwFecha1 = True
    vlSwFecha2 = True
    vlSwFecha3 = True
    vlSwFecha4 = True
    vlSwFecha5 = True
    vlSwFecha6 = True
    vlSwFechaV = True
    vlSwFechaR = True
    vlSwFechaValidez = True 'hqr 05/03/2005
    
    SSTab1.Enabled = True
    
       
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub


VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Frm_GECalculoPorcentaje 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cálculo de Porcentaje de Deducción"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   9060
   Begin VB.Frame Fra_Botones 
      Height          =   975
      Left            =   120
      TabIndex        =   51
      Top             =   7440
      Width           =   8895
      Begin VB.CommandButton Cmd_Grabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   2400
         Picture         =   "Frm_GECalculoPorcentaje.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   4560
         Picture         =   "Frm_GECalculoPorcentaje.frx":06BA
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   6720
         Picture         =   "Frm_GECalculoPorcentaje.frx":0D74
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   3480
         Picture         =   "Frm_GECalculoPorcentaje.frx":0E6E
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Eliminar 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   480
         Picture         =   "Frm_GECalculoPorcentaje.frx":1528
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Eliminar Año"
         Top             =   200
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   5640
         Picture         =   "Frm_GECalculoPorcentaje.frx":186A
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   200
         Width           =   730
      End
      Begin Crystal.CrystalReport Rpt_Calculo 
         Left            =   8280
         Top             =   240
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
      TabIndex        =   42
      Top             =   0
      Width           =   8895
      Begin VB.CommandButton Cmd_BuscarPol 
         Height          =   375
         Left            =   8160
         Picture         =   "Frm_GECalculoPorcentaje.frx":1E44
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Buscar Póliza"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Txt_PenRut 
         Height          =   285
         Left            =   3720
         MaxLength       =   10
         TabIndex        =   1
         Top             =   360
         Width           =   1755
      End
      Begin VB.TextBox Txt_PenDigito 
         Height          =   285
         Left            =   5865
         MaxLength       =   1
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Txt_PenPoliza 
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   0
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
         Left            =   8160
         Picture         =   "Frm_GECalculoPorcentaje.frx":1F46
         TabIndex        =   4
         ToolTipText     =   "Buscar"
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Lbl_PenEndoso 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   255
         Left            =   7425
         TabIndex        =   53
         Top             =   360
         Width           =   525
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nº Endoso"
         Height          =   240
         Index           =   33
         Left            =   6480
         TabIndex        =   52
         Top             =   360
         Width           =   840
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
         TabIndex        =   47
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Rut Pensionado"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   46
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   45
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Póliza"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   44
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Lbl_PenNombre 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   255
         Left            =   1080
         TabIndex        =   43
         Top             =   720
         Width           =   6855
      End
   End
   Begin TabDlg.SSTab SSTab_Deduccion 
      Height          =   5595
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9869
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Deduccion Excedente / Herencia"
      TabPicture(0)   =   "Frm_GECalculoPorcentaje.frx":2048
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Fra_Vigencia"
      Tab(0).Control(1)=   "Fra_TipoPen"
      Tab(0).Control(2)=   "Fra_Retiros"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Deducción Vejez Anticipada"
      TabPicture(1)   =   "Frm_GECalculoPorcentaje.frx":2064
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Txt_FecPensionAnt"
      Tab(1).Control(1)=   "Fra_Modalidad"
      Tab(1).Control(2)=   "Fra_Capital"
      Tab(1).Control(3)=   "MSFGrilla_Beneficiarios"
      Tab(1).Control(4)=   "Lbl_EdadActuarial"
      Tab(1).Control(5)=   "Lbl_FecEdadLegal"
      Tab(1).Control(6)=   "Label14"
      Tab(1).Control(7)=   "Label15"
      Tab(1).Control(8)=   "Label16"
      Tab(1).Control(9)=   "Label18"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Historia"
      TabPicture(2)   =   "Frm_GECalculoPorcentaje.frx":2080
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Msf_Grilla"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Fra_Vigencia 
         Caption         =   " Vigencia del Porcentaje "
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
         Height          =   700
         Left            =   -74880
         TabIndex        =   73
         Top             =   1000
         Width           =   8655
         Begin VB.TextBox Txt_FechaInicio 
            Height          =   285
            Left            =   1920
            TabIndex        =   6
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Lbl_FechaTermino 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3600
            TabIndex        =   76
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label3 
            Caption         =   " - "
            Height          =   255
            Left            =   3240
            TabIndex        =   75
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label2 
            Caption         =   "Periodo de Vigencia"
            Height          =   255
            Left            =   240
            TabIndex        =   74
            Top             =   240
            Width           =   1695
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Msf_Grilla 
         Height          =   4695
         Left            =   240
         TabIndex        =   23
         Top             =   600
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   8281
         _Version        =   393216
      End
      Begin VB.Frame Fra_TipoPen 
         Caption         =   "  Tipo de Pensión  "
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
         Height          =   615
         Left            =   -74880
         TabIndex        =   70
         Top             =   360
         Width           =   8655
         Begin VB.Label Lbl_TipoPension 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            TabIndex        =   72
            Top             =   195
            Width           =   6600
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo de Pensión"
            Height          =   255
            Left            =   240
            TabIndex        =   71
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Fra_Retiros 
         Caption         =   "  Antecedentes  "
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
         Height          =   3735
         Left            =   -74880
         TabIndex        =   58
         Top             =   1740
         Width           =   8655
         Begin VB.TextBox Txt_FecPension 
            Height          =   285
            Left            =   6600
            MaxLength       =   10
            TabIndex        =   7
            Top             =   330
            Width           =   1680
         End
         Begin VB.Frame Fra_Detalle 
            Caption         =   "  Detalle de Retiros "
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
            Height          =   3375
            Left            =   120
            TabIndex        =   65
            Top             =   230
            Width           =   4335
            Begin VB.CommandButton Cmd_Cancela 
               Height          =   450
               Left            =   3660
               Picture         =   "Frm_GECalculoPorcentaje.frx":209C
               Style           =   1  'Graphical
               TabIndex        =   12
               ToolTipText     =   "Limpiar"
               Top             =   390
               Width           =   495
            End
            Begin VB.TextBox Txt_NumCuotas 
               Height          =   285
               Left            =   1440
               MaxLength       =   10
               TabIndex        =   9
               Top             =   480
               Width           =   1110
            End
            Begin VB.TextBox Txt_FecRetiro 
               Height          =   285
               Left            =   120
               MaxLength       =   10
               TabIndex        =   8
               Top             =   480
               Width           =   1140
            End
            Begin VB.CommandButton Cmd_Sumar 
               Height          =   450
               Left            =   2700
               Picture         =   "Frm_GECalculoPorcentaje.frx":270E
               Style           =   1  'Graphical
               TabIndex        =   10
               ToolTipText     =   "Agregar Beneficiario"
               Top             =   390
               Width           =   495
            End
            Begin VB.CommandButton Cmd_Restar 
               Height          =   450
               Left            =   3180
               Picture         =   "Frm_GECalculoPorcentaje.frx":2898
               Style           =   1  'Graphical
               TabIndex        =   11
               ToolTipText     =   "Quitar Beneficiario"
               Top             =   390
               Width           =   495
            End
            Begin MSFlexGridLib.MSFlexGrid MSFGrilla_Retiros 
               Height          =   2250
               Left            =   135
               TabIndex        =   66
               Top             =   945
               Width           =   4080
               _ExtentX        =   7197
               _ExtentY        =   3969
               _Version        =   393216
               BackColor       =   -2147483624
            End
            Begin VB.Label Label10 
               Caption         =   "N°de Cuotas"
               Height          =   285
               Left            =   1440
               TabIndex        =   68
               Top             =   240
               Width           =   1005
            End
            Begin VB.Label Label9 
               Caption         =   "Fecha Retiro"
               Height          =   210
               Left            =   120
               TabIndex        =   67
               Top             =   240
               Width           =   1260
            End
         End
         Begin VB.Frame Fra_Deduccion 
            Caption         =   "  Cálculo de Deducción  "
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
            Height          =   2925
            Left            =   4560
            TabIndex        =   59
            Top             =   690
            Width           =   3930
            Begin VB.TextBox Txt_SaldoCuenta 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   2280
               MaxLength       =   12
               TabIndex        =   13
               Top             =   795
               Width           =   1470
            End
            Begin VB.CommandButton Cmd_Calcular1 
               Caption         =   "&Calcular"
               Height          =   675
               Left            =   1680
               Picture         =   "Frm_GECalculoPorcentaje.frx":2A22
               Style           =   1  'Graphical
               TabIndex        =   14
               ToolTipText     =   "Generar Tabla de Mortalidad"
               Top             =   1440
               Width           =   720
            End
            Begin VB.Label Label13 
               Caption         =   "% de Deducción"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   480
               TabIndex        =   64
               Top             =   2280
               Width           =   1650
            End
            Begin VB.Label Label12 
               Caption         =   "Saldo Total Inicial de la Cuenta (en Cuotas)"
               Height          =   420
               Left            =   240
               TabIndex        =   63
               Top             =   810
               Width           =   1905
            End
            Begin VB.Label Label11 
               Caption         =   "Suma Total de Cuotas Retiradas"
               Height          =   375
               Left            =   240
               TabIndex        =   62
               Top             =   270
               Width           =   1995
            End
            Begin VB.Label Lbl_SumaCuotas 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Enabled         =   0   'False
               Height          =   285
               Left            =   2280
               TabIndex        =   61
               Top             =   315
               Width           =   1470
            End
            Begin VB.Label Lbl_PrcDed1 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Enabled         =   0   'False
               Height          =   285
               Left            =   2160
               TabIndex        =   60
               Top             =   2235
               Width           =   1605
            End
         End
         Begin VB.Label Label8 
            Caption         =   "Fecha de Pensión"
            Height          =   255
            Left            =   4920
            TabIndex        =   69
            Top             =   330
            Width           =   1620
         End
      End
      Begin VB.TextBox Txt_FecPensionAnt 
         Height          =   285
         Left            =   -72720
         MaxLength       =   10
         TabIndex        =   15
         Top             =   600
         Width           =   1320
      End
      Begin VB.Frame Fra_Modalidad 
         Caption         =   "Marque la Modalidad seleccionada:"
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
         Height          =   1170
         Left            =   -74850
         TabIndex        =   37
         Top             =   1020
         Width           =   8550
         Begin VB.OptionButton Opt_RP 
            Caption         =   "Retiro Programado (Fecha de solicitud de pensión)"
            Height          =   225
            Left            =   165
            TabIndex        =   16
            Top             =   315
            Width           =   4710
         End
         Begin VB.OptionButton Opt_RPcRVI 
            Caption         =   "Retiro Programado con Renta Viitalicia Inmediata (Fecha de Traspaso de Prima)"
            Height          =   225
            Left            =   165
            TabIndex        =   18
            Top             =   855
            Width           =   6705
         End
         Begin VB.OptionButton Opt_RVI 
            Caption         =   "Renta Vitalicia Inmediata (Fecha de Traspaso de Prima)"
            Height          =   225
            Left            =   150
            TabIndex        =   17
            Top             =   585
            Width           =   4425
         End
      End
      Begin VB.Frame Fra_Capital 
         Caption         =   "Cálculo de Capital Necesario Unitario"
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
         Height          =   1425
         Left            =   -74880
         TabIndex        =   31
         Top             =   3900
         Width           =   8640
         Begin VB.TextBox Txt_CNpv 
            Height          =   285
            Left            =   3840
            MaxLength       =   16
            TabIndex        =   21
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox Txt_CNpa 
            Height          =   285
            Left            =   3840
            MaxLength       =   16
            TabIndex        =   20
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox Txt_Tasa 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Left            =   3840
            TabIndex        =   57
            Top             =   405
            Width           =   1455
         End
         Begin VB.CommandButton Cmd_Calcular2 
            Caption         =   "&Calcular"
            Height          =   675
            Left            =   6360
            Picture         =   "Frm_GECalculoPorcentaje.frx":2EC4
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Generar Tabla de Mortalidad"
            Top             =   240
            Width           =   720
         End
         Begin VB.ComboBox Cmb_AFP 
            BackColor       =   &H80000018&
            Height          =   315
            Left            =   825
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   360
            Width           =   2115
         End
         Begin VB.Label Lbl_PrcDed2 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   7200
            TabIndex        =   56
            Top             =   1005
            Width           =   1305
         End
         Begin VB.Label Label19 
            Caption         =   "Capital necesario para pensión anticipada (CNpa)"
            Height          =   255
            Left            =   105
            TabIndex        =   36
            Top             =   840
            Width           =   3705
         End
         Begin VB.Label Label20 
            Caption         =   "A.F.P."
            Height          =   315
            Left            =   120
            TabIndex        =   35
            Top             =   375
            Width           =   690
         End
         Begin VB.Label Label21 
            Caption         =   "Tasa"
            Height          =   225
            Left            =   3240
            TabIndex        =   34
            Top             =   450
            Width           =   555
         End
         Begin VB.Label Label22 
            Caption         =   "Capital necesario para pensión de vejez    (CNpv)"
            Height          =   255
            Left            =   90
            TabIndex        =   33
            Top             =   1080
            Width           =   3705
         End
         Begin VB.Label Label23 
            Caption         =   "% de Deducción"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5640
            TabIndex        =   32
            Top             =   1080
            Width           =   1515
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFGrilla_Beneficiarios 
         Height          =   1395
         Left            =   -74820
         TabIndex        =   29
         Top             =   2430
         Width           =   8550
         _ExtentX        =   15081
         _ExtentY        =   2461
         _Version        =   393216
         BackColor       =   -2147483624
      End
      Begin VB.Label Lbl_EdadActuarial 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   -67080
         TabIndex        =   55
         Top             =   600
         Width           =   525
      End
      Begin VB.Label Lbl_FecEdadLegal 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   -70080
         TabIndex        =   54
         Top             =   600
         Width           =   1305
      End
      Begin VB.Label Label14 
         Caption         =   "Fecha Pensión Antic. (P.A.)"
         Height          =   255
         Left            =   -74760
         TabIndex        =   41
         Top             =   600
         Width           =   2130
      End
      Begin VB.Label Label15 
         Caption         =   "Fecha Edad Legal"
         Height          =   495
         Left            =   -71160
         TabIndex        =   40
         Top             =   600
         Width           =   1140
      End
      Begin VB.Label Label16 
         Caption         =   "Edad actuarial  fecha P.A."
         Height          =   495
         Left            =   -68280
         TabIndex        =   39
         Top             =   600
         Width           =   1140
      End
      Begin VB.Label Label18 
         Caption         =   "Beneficiarios Legales"
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
         Height          =   300
         Left            =   -74805
         TabIndex        =   38
         Top             =   2220
         Width           =   2520
      End
   End
   Begin VB.Frame Fra_Total 
      Height          =   735
      Left            =   120
      TabIndex        =   48
      Top             =   6720
      Width           =   8895
      Begin VB.TextBox Txt_TotalPorDed 
         BackColor       =   &H00E0FFFF&
         Enabled         =   0   'False
         Height          =   330
         Left            =   3720
         TabIndex        =   49
         Top             =   240
         Width           =   1635
      End
      Begin VB.Label Label24 
         Caption         =   "Total Porcentaje de Deducción "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   600
         TabIndex        =   50
         Top             =   240
         Width           =   2745
      End
   End
End
Attribute VB_Name = "Frm_GECalculoPorcentaje"
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

    Frm_GECalculoPorcentaje.Top = 0
    Frm_GECalculoPorcentaje.Left = 0
      
            
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub


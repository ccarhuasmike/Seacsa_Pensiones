VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Frm_PensMantenciónPensionado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantención antecedentes pensionado"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   9750
   Begin VB.Frame Frame9 
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
      TabIndex        =   35
      Top             =   0
      Width           =   9255
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   1080
         TabIndex        =   40
         Top             =   360
         Width           =   1185
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   5880
         TabIndex        =   39
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00E0FFFF&
         Height          =   285
         Left            =   1080
         TabIndex        =   38
         Top             =   720
         Width           =   6825
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   3960
         TabIndex        =   37
         Top             =   360
         Width           =   1755
      End
      Begin VB.CommandButton Cmd_Buscar 
         Height          =   375
         Left            =   6360
         Picture         =   "Frm_PensMantencion.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Buscar Póliza"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label18 
         Caption         =   "N° Póliza"
         Height          =   375
         Left            =   360
         TabIndex        =   44
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label7 
         Caption         =   "Nombre"
         Height          =   285
         Left            =   360
         TabIndex        =   43
         Top             =   720
         Width           =   990
      End
      Begin VB.Label Label17 
         Caption         =   "Rut Causante"
         Height          =   285
         Left            =   2640
         TabIndex        =   42
         Top             =   360
         Width           =   1365
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "-"
         Height          =   255
         Left            =   5640
         TabIndex        =   41
         Top             =   360
         Width           =   255
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6540
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   11536
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Póliza"
      TabPicture(0)   =   "Frm_PensMantencion.frx":0102
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label21"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label22"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label23"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label24"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label25"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label27"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label28"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label29"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label30"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label31"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label26"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label32"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label33"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label34"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label35"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text14"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Combo8"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Combo9"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Combo10"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Combo11"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text5"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text6"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text16"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Combo12"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Text17"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Text18"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Text19"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Text20"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Text21"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text22"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Text23"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text24"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).ControlCount=   34
      TabCaption(1)   =   "Beneficiarios"
      TabPicture(1)   =   "Frm_PensMantencion.frx":011E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MSFlexGrid1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame5"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame3"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.TextBox Text24 
         Height          =   285
         Left            =   2280
         TabIndex        =   78
         Top             =   5160
         Width           =   1215
      End
      Begin VB.TextBox Text23 
         Height          =   285
         Left            =   2280
         TabIndex        =   77
         Top             =   4800
         Width           =   1215
      End
      Begin VB.TextBox Text22 
         Height          =   285
         Left            =   2280
         TabIndex        =   76
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox Text21 
         Height          =   285
         Left            =   2280
         TabIndex        =   72
         Top             =   4080
         Width           =   1215
      End
      Begin VB.TextBox Text20 
         Height          =   285
         Left            =   7320
         TabIndex        =   69
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox Text19 
         Height          =   285
         Left            =   7320
         TabIndex        =   68
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox Text18 
         Height          =   285
         Left            =   2280
         TabIndex        =   67
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox Text17 
         Height          =   285
         Left            =   2280
         TabIndex        =   66
         Top             =   3360
         Width           =   1215
      End
      Begin VB.ComboBox Combo12 
         Height          =   315
         Left            =   2280
         TabIndex        =   65
         Top             =   2280
         Width           =   2775
      End
      Begin VB.TextBox Text16 
         Height          =   285
         Left            =   3720
         TabIndex        =   59
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   2280
         TabIndex        =   58
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2280
         TabIndex        =   56
         Top             =   2640
         Width           =   1215
      End
      Begin VB.ComboBox Combo11 
         Height          =   315
         Left            =   2280
         TabIndex        =   54
         Top             =   1920
         Width           =   2775
      End
      Begin VB.ComboBox Combo10 
         Height          =   315
         Left            =   2280
         TabIndex        =   53
         Top             =   1560
         Width           =   2775
      End
      Begin VB.ComboBox Combo9 
         Height          =   315
         Left            =   2280
         TabIndex        =   52
         Top             =   1200
         Width           =   2775
      End
      Begin VB.ComboBox Combo8 
         Height          =   315
         Left            =   2280
         TabIndex        =   48
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   2280
         TabIndex        =   47
         Top             =   480
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Caption         =   " Plan de Salud     "
         Height          =   1545
         Left            =   -70440
         TabIndex        =   28
         Top             =   4200
         Width           =   4230
         Begin VB.TextBox Text13 
            Height          =   285
            Left            =   990
            TabIndex        =   31
            Top             =   1005
            Width           =   1500
         End
         Begin VB.ComboBox Combo6 
            Height          =   315
            Left            =   975
            TabIndex        =   30
            Top             =   645
            Width           =   1905
         End
         Begin VB.ComboBox Combo7 
            Height          =   315
            Left            =   975
            TabIndex        =   29
            Top             =   285
            Width           =   2955
         End
         Begin VB.Label Label16 
            Caption         =   "Monto"
            Height          =   270
            Left            =   165
            TabIndex        =   34
            Top             =   1050
            Width           =   795
         End
         Begin VB.Label Label19 
            Caption         =   "Moneda"
            Height          =   270
            Left            =   165
            TabIndex        =   33
            Top             =   750
            Width           =   825
         End
         Begin VB.Label Label20 
            Caption         =   "Institución"
            Height          =   285
            Left            =   180
            TabIndex        =   32
            Top             =   405
            Width           =   810
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Tutor"
         Height          =   1545
         Left            =   -70200
         TabIndex        =   20
         Top             =   5400
         Width           =   4230
         Begin VB.TextBox Text4 
            BackColor       =   &H00E0FFFF&
            Height          =   285
            Left            =   1500
            TabIndex        =   24
            Top             =   405
            Width           =   1950
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H00E0FFFF&
            Height          =   285
            Left            =   1500
            TabIndex        =   23
            Top             =   720
            Width           =   2655
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00E0FFFF&
            Height          =   300
            Left            =   1500
            TabIndex        =   22
            Top             =   1020
            Width           =   1935
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00E0FFFF&
            Height          =   285
            Left            =   3525
            TabIndex        =   21
            Top             =   390
            Width           =   435
         End
         Begin VB.Label Label3 
            Caption         =   "Rut "
            Height          =   285
            Left            =   180
            TabIndex        =   27
            Top             =   405
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Nombre"
            Height          =   270
            Left            =   150
            TabIndex        =   26
            Top             =   705
            Width           =   1275
         End
         Begin VB.Label Label1 
            Caption         =   "Fec.Vencimiento"
            Height          =   255
            Left            =   180
            TabIndex        =   25
            Top             =   1080
            Width           =   1260
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Forma de Pago de Pensión"
         Height          =   1980
         Left            =   -74760
         TabIndex        =   11
         Top             =   4200
         Width           =   4230
         Begin VB.TextBox Text12 
            Height          =   285
            Left            =   975
            TabIndex        =   15
            Top             =   1455
            Width           =   2910
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            Left            =   990
            TabIndex        =   14
            Top             =   690
            Width           =   2940
         End
         Begin VB.ComboBox Combo4 
            Height          =   315
            Left            =   990
            TabIndex        =   13
            Top             =   315
            Width           =   2955
         End
         Begin VB.ComboBox Combo5 
            Height          =   315
            Left            =   990
            TabIndex        =   12
            Top             =   1050
            Width           =   2940
         End
         Begin VB.Label Label12 
            Caption         =   "N°Cuenta"
            Height          =   270
            Left            =   180
            TabIndex        =   19
            Top             =   1515
            Width           =   795
         End
         Begin VB.Label Label14 
            Caption         =   "Banco"
            Height          =   270
            Left            =   165
            TabIndex        =   18
            Top             =   780
            Width           =   825
         End
         Begin VB.Label Label15 
            Caption         =   "Vía Pago"
            Height          =   285
            Left            =   165
            TabIndex        =   17
            Top             =   405
            Width           =   810
         End
         Begin VB.Label Label13 
            Caption         =   "Tipo Cta."
            Height          =   270
            Left            =   165
            TabIndex        =   16
            Top             =   1140
            Width           =   825
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Antecedenes Personales"
         Height          =   1980
         Left            =   -74730
         TabIndex        =   4
         Top             =   2040
         Width           =   8430
         Begin VB.TextBox Text27 
            Height          =   285
            Left            =   960
            TabIndex        =   84
            Top             =   720
            Width           =   6975
         End
         Begin VB.TextBox Text26 
            Height          =   285
            Left            =   2520
            TabIndex        =   81
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox Text25 
            Height          =   285
            Left            =   960
            TabIndex        =   80
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox Text10 
            Height          =   300
            Left            =   960
            TabIndex        =   7
            Top             =   1155
            Width           =   7035
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   975
            TabIndex        =   6
            Top             =   1530
            Width           =   2955
         End
         Begin VB.TextBox Text11 
            Height          =   285
            Left            =   5625
            TabIndex        =   5
            Top             =   1560
            Width           =   1500
         End
         Begin VB.Label Label37 
            Caption         =   "Nombre"
            Height          =   255
            Left            =   240
            TabIndex        =   83
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label36 
            Alignment       =   2  'Center
            Caption         =   "-"
            Height          =   255
            Left            =   2400
            TabIndex        =   82
            Top             =   360
            Width           =   135
         End
         Begin VB.Label Label10 
            Caption         =   "Rut"
            Height          =   255
            Left            =   240
            TabIndex        =   79
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label8 
            Caption         =   "Domicilio"
            Height          =   285
            Left            =   180
            TabIndex        =   10
            Top             =   1245
            Width           =   810
         End
         Begin VB.Label Label9 
            Caption         =   "Comuna"
            Height          =   270
            Left            =   165
            TabIndex        =   9
            Top             =   1650
            Width           =   825
         End
         Begin VB.Label Label11 
            Caption         =   "Teléfono"
            Height          =   270
            Left            =   4800
            TabIndex        =   8
            Top             =   1560
            Width           =   795
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   1650
         Left            =   -74670
         TabIndex        =   3
         Top             =   480
         Width           =   8460
         _ExtentX        =   14923
         _ExtentY        =   2910
         _Version        =   393216
         Cols            =   6
         FormatString    =   $"Frm_PensMantencion.frx":013A
      End
      Begin VB.Label Label35 
         Caption         =   "Tasa de Interés Periodo Garantizado"
         Height          =   495
         Left            =   360
         TabIndex        =   75
         Top             =   5160
         Width           =   1575
      End
      Begin VB.Label Label34 
         Caption         =   "Tasa Cto. Reaseguro"
         Height          =   255
         Left            =   360
         TabIndex        =   74
         Top             =   4800
         Width           =   1815
      End
      Begin VB.Label Label33 
         Caption         =   "Tasa de Venta"
         Height          =   255
         Left            =   360
         TabIndex        =   73
         Top             =   4440
         Width           =   1935
      End
      Begin VB.Label Label32 
         Caption         =   "Tasa Cto. Equivalente"
         Height          =   255
         Left            =   360
         TabIndex        =   71
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         Caption         =   "-"
         Height          =   255
         Left            =   3480
         TabIndex        =   70
         Top             =   3000
         Width           =   255
      End
      Begin VB.Label Label31 
         Caption         =   "Tipo de Renta"
         Height          =   255
         Left            =   360
         TabIndex        =   64
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label30 
         Caption         =   "Meses Garantizados"
         Height          =   255
         Left            =   5400
         TabIndex        =   63
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label29 
         Caption         =   "Meses Diferidos"
         Height          =   255
         Left            =   5400
         TabIndex        =   62
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label28 
         Caption         =   "Monto Pensión"
         Height          =   255
         Left            =   360
         TabIndex        =   61
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label27 
         Caption         =   "Monto Prima"
         Height          =   375
         Left            =   360
         TabIndex        =   60
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label25 
         Caption         =   "Periodo de Vigencia"
         Height          =   255
         Left            =   360
         TabIndex        =   57
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label24 
         Caption         =   "Nº Cargas"
         Height          =   255
         Left            =   360
         TabIndex        =   55
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label23 
         Caption         =   "Estado"
         Height          =   255
         Left            =   360
         TabIndex        =   51
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo de Pensión "
         Height          =   300
         Left            =   360
         TabIndex        =   50
         Top             =   1200
         Width           =   1635
      End
      Begin VB.Label Label5 
         Caption         =   "Modalidad"
         Height          =   285
         Left            =   360
         TabIndex        =   49
         Top             =   1920
         Width           =   1260
      End
      Begin VB.Label Label22 
         Caption         =   "AFP"
         Height          =   255
         Left            =   360
         TabIndex        =   46
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "Endoso"
         Height          =   255
         Left            =   360
         TabIndex        =   45
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Grabar"
      Height          =   315
      Left            =   300
      TabIndex        =   1
      Top             =   7920
      Width           =   1185
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cerrar"
      Height          =   330
      Left            =   8400
      TabIndex        =   0
      Top             =   7920
      Width           =   1095
   End
End
Attribute VB_Name = "Frm_PensMantenciónPensionado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command5_Click()

    Unload Me
End Sub


VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Frm_ConsultaEndoso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Endoso"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   9000
   Begin VB.Frame Fra_Botones 
      Height          =   975
      Left            =   120
      TabIndex        =   86
      Top             =   7920
      Width           =   8775
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5280
         Picture         =   "Frm_ConsultaEndoso.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   2160
         Picture         =   "Frm_ConsultaEndoso.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   3720
         Picture         =   "Frm_ConsultaEndoso.frx":07B4
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   200
         Width           =   720
      End
      Begin Crystal.CrystalReport Rpt_Reporte 
         Left            =   8160
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowState     =   2
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
      Left            =   0
      TabIndex        =   84
      Top             =   0
      Width           =   8895
      Begin VB.TextBox Txt_PenEndoso 
         Height          =   285
         Left            =   7560
         MaxLength       =   3
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Txt_PenPoliza 
         Height          =   285
         Left            =   840
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   1185
      End
      Begin VB.TextBox Txt_PenNumIdent 
         Height          =   285
         Left            =   4920
         MaxLength       =   16
         TabIndex        =   3
         Top             =   240
         Width           =   1875
      End
      Begin VB.CommandButton Cmd_BuscarPol 
         Height          =   375
         Left            =   8160
         Picture         =   "Frm_ConsultaEndoso.frx":0D8E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Buscar Póliza"
         Top             =   120
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
         Picture         =   "Frm_ConsultaEndoso.frx":0E90
         TabIndex        =   6
         ToolTipText     =   "Buscar"
         Top             =   480
         Width           =   615
      End
      Begin VB.ComboBox Cmb_PenNumIdent 
         BackColor       =   &H00E0FFFF&
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1995
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Póliza"
         Height          =   255
         Index           =   19
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Lbl_PenNombre 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   600
         Width           =   7215
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Ident."
         Height          =   255
         Index           =   0
         Left            =   2160
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "N° End"
         Height          =   195
         Index           =   42
         Left            =   6960
         TabIndex        =   8
         Top             =   240
         Width           =   510
      End
      Begin VB.Label Label8 
         Caption         =   " Póliza"
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
         TabIndex        =   85
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   87
         Top             =   600
         Width           =   855
      End
   End
   Begin TabDlg.SSTab SSTab_PolizaOriginal 
      Height          =   6735
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   8421376
      TabCaption(0)   =   "Póliza"
      TabPicture(0)   =   "Frm_ConsultaEndoso.frx":0F92
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Lbl_PolMesDif"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Lbl_PolMesGar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Lbl_Nombre(4)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Lbl_Nombre(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Lbl_Nombre(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Lbl_Nombre(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Lbl_PolEstVig"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Lbl_Nombre(28)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Lbl_Nombre(24)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Lbl_Nombre(25)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Lbl_Nombre(30)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Lbl_Nombre(31)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Lbl_Nombre(33)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Lbl_Nombre(34)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Lbl_Nombre(27)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Lbl_Nombre(29)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Lbl_Nombre(26)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Lbl_Nombre(32)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Lbl_Nombre(35)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Lbl_Nombre(36)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Lbl_Nombre(38)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Lbl_PolNumCar"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Lbl_PolIniVig"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Lbl_PolTerVig"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Lbl_PolMtoPri"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Lbl_PolMtoPen"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Lbl_PolTasaCto"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Lbl_PolTasaVta"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Lbl_PolTasaPerGar"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Lbl_PolTipPen"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Lbl_PolTipRta"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Lbl_PolMod"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Lbl_FecDevengue"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Lbl_FecEmision"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Lbl_Nombre(52)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Lbl_Nombre(20)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Lbl_MonPrima"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Lbl_MonPension(0)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Lbl_POCtaInd"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Lbl_POIngBase"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Lbl_POPrcCubierto"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Lbl_PODerGratificacion"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Lbl_Nombre(98)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Lbl_PODerCrecer"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Lbl_Nombre(97)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Lbl_POCoberCon"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Lbl_Nombre(96)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "Lbl_POIndCobertura"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "Lbl_Nombre(95)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "Lbl_PolCuspp"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "Lbl_Nombre(9)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).ControlCount=   51
      TabCaption(1)   =   "Beneficiarios"
      TabPicture(1)   =   "Frm_ConsultaEndoso.frx":0FAE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Msf_BenGrilla"
      Tab(1).Control(1)=   "Fra_Personales"
      Tab(1).ControlCount=   2
      Begin VB.Frame Fra_Personales 
         Caption         =   "Antecedentes Personales"
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
         Height          =   4425
         Left            =   -74760
         TabIndex        =   10
         Top             =   2040
         Width           =   8295
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mto. Pensión Gar."
            Height          =   195
            Index           =   21
            Left            =   3000
            TabIndex        =   110
            Top             =   3840
            Width           =   1275
         End
         Begin VB.Label Label3 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   4320
            TabIndex        =   109
            Top             =   3840
            Width           =   1335
         End
         Begin VB.Label Lbl_MonPension 
            Caption         =   "(TM)"
            Enabled         =   0   'False
            Height          =   300
            Index           =   2
            Left            =   5640
            TabIndex        =   108
            Top             =   3840
            Width           =   495
         End
         Begin VB.Label Lbl_BenNumIden 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   6360
            TabIndex        =   97
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. Ident."
            Height          =   255
            Index           =   10
            Left            =   3360
            TabIndex        =   96
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Lbl_BenTipoIden 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   4260
            TabIndex        =   95
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Lbl_MonPension 
            Caption         =   "(TM)"
            Enabled         =   0   'False
            Height          =   300
            Index           =   1
            Left            =   5640
            TabIndex        =   94
            Top             =   3480
            Width           =   495
         End
         Begin VB.Label Lbl_BenMtoPension 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   4320
            TabIndex        =   47
            Top             =   3480
            Width           =   1335
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mto. Pensión"
            Height          =   195
            Index           =   41
            Left            =   3000
            TabIndex        =   46
            Top             =   3480
            Width           =   1050
         End
         Begin VB.Label Lbl_BenSitInv 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   1080
            TabIndex        =   45
            Top             =   2040
            Width           =   615
         End
         Begin VB.Label Lbl_BenFecNac 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   1080
            TabIndex        =   44
            Top             =   2760
            Width           =   1335
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sit. de Inv."
            Height          =   195
            Index           =   40
            Left            =   120
            TabIndex        =   43
            Top             =   2040
            Width           =   1005
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Nac."
            Height          =   195
            Index           =   39
            Left            =   120
            TabIndex        =   42
            Top             =   2760
            Width           =   840
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Grupo Fam."
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   41
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Parentesco"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   40
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Lbl_BenNumOrd 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   1080
            TabIndex        =   39
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Orden"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   37
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Lbl_BenNomBen 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   1080
            TabIndex        =   36
            Top             =   600
            Width           =   7095
         End
         Begin VB.Label Lbl_BenPar 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   1080
            TabIndex        =   35
            Top             =   960
            Width           =   7095
         End
         Begin VB.Label Lbl_BenGrupFam 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   1080
            TabIndex        =   34
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sexo"
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   33
            Top             =   1680
            Width           =   960
         End
         Begin VB.Label Lbl_BenSexo 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   1080
            TabIndex        =   32
            Top             =   1680
            Width           =   615
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Dº Pensión"
            Height          =   255
            Index           =   12
            Left            =   3300
            TabIndex        =   31
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Lbl_BenDerPen 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   4320
            TabIndex        =   30
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Dº Acrecer"
            Height          =   255
            Index           =   13
            Left            =   3300
            TabIndex        =   29
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Lbl_BenDerAcrecer 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   4320
            TabIndex        =   28
            Top             =   1680
            Width           =   615
         End
         Begin VB.Label Lbl_BenFecFall 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   1080
            TabIndex        =   27
            Top             =   3120
            Width           =   1335
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Mat."
            Height          =   195
            Index           =   14
            Left            =   5880
            TabIndex        =   26
            Top             =   2880
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.Label Lbl_BenFecNHM 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   1080
            TabIndex        =   25
            Top             =   3480
            Width           =   1335
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha NHM"
            Height          =   195
            Index           =   15
            Left            =   120
            TabIndex        =   24
            Top             =   3480
            Width           =   975
         End
         Begin VB.Label Lbl_BenPrcLegal 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   1080
            TabIndex        =   23
            Top             =   3840
            Width           =   855
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Porc. Legal"
            Height          =   195
            Index           =   16
            Left            =   120
            TabIndex        =   22
            Top             =   3840
            Width           =   1050
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "%"
            Height          =   255
            Index           =   17
            Left            =   2040
            TabIndex        =   21
            Top             =   3840
            Width           =   255
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Causal Inv."
            Height          =   195
            Index           =   77
            Left            =   120
            TabIndex        =   20
            Top             =   2400
            Width           =   795
         End
         Begin VB.Label Lbl_BenCauInv 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   1080
            TabIndex        =   19
            Top             =   2400
            Width           =   7095
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Inv."
            Height          =   195
            Index           =   78
            Left            =   3300
            TabIndex        =   18
            Top             =   2040
            Width           =   765
         End
         Begin VB.Label Lbl_BenFecInv 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   4320
            TabIndex        =   17
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Fall."
            Height          =   195
            Index           =   89
            Left            =   120
            TabIndex        =   16
            Top             =   3120
            Width           =   1020
         End
         Begin VB.Label Lbl_BenFecMat 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   6840
            TabIndex        =   15
            Top             =   2880
            Visible         =   0   'False
            Width           =   1335
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Msf_BenGrilla 
         Height          =   1530
         Left            =   -74760
         TabIndex        =   48
         Top             =   480
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   2699
         _Version        =   393216
         Rows            =   1
         Cols            =   6
         FixedCols       =   0
         BackColor       =   14745599
         FormatString    =   $"Frm_ConsultaEndoso.frx":0FCA
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
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "CUSPP"
         Height          =   195
         Index           =   9
         Left            =   240
         TabIndex        =   107
         Top             =   4620
         Width           =   540
      End
      Begin VB.Label Lbl_PolCuspp 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   106
         Top             =   4560
         Width           =   1935
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Ind. Cobertura"
         Height          =   195
         Index           =   95
         Left            =   240
         TabIndex        =   105
         Top             =   5640
         Width           =   1005
      End
      Begin VB.Label Lbl_POIndCobertura 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   104
         Top             =   5580
         Width           =   1095
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Cobertura a la Cónyuge"
         Height          =   195
         Index           =   96
         Left            =   240
         TabIndex        =   103
         Top             =   5955
         Width           =   1665
      End
      Begin VB.Label Lbl_POCoberCon 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   102
         Top             =   5920
         Width           =   1095
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Der. Crecer"
         Height          =   195
         Index           =   97
         Left            =   5640
         TabIndex        =   101
         Top             =   5640
         Width           =   810
      End
      Begin VB.Label Lbl_PODerCrecer 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   7080
         TabIndex        =   100
         Top             =   5580
         Width           =   975
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Gratificación"
         Height          =   195
         Index           =   98
         Left            =   5640
         TabIndex        =   99
         Top             =   5955
         Width           =   885
      End
      Begin VB.Label Lbl_PODerGratificacion 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   7080
         TabIndex        =   98
         Top             =   5920
         Width           =   975
      End
      Begin VB.Label Lbl_POPrcCubierto 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   7080
         TabIndex        =   52
         Top             =   4800
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Lbl_POIngBase 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   7080
         TabIndex        =   53
         Top             =   4440
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Lbl_POCtaInd 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   7080
         TabIndex        =   54
         Top             =   4080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Lbl_MonPension 
         Caption         =   "(TM)"
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   4200
         TabIndex        =   93
         Top             =   3540
         Width           =   495
      End
      Begin VB.Label Lbl_MonPrima 
         Caption         =   "S/."
         Enabled         =   0   'False
         Height          =   285
         Left            =   4200
         TabIndex        =   92
         Top             =   3195
         Width           =   495
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Emisión"
         Height          =   195
         Index           =   20
         Left            =   240
         TabIndex        =   91
         Top             =   4950
         Width           =   1260
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Devengue"
         Height          =   195
         Index           =   52
         Left            =   240
         TabIndex        =   90
         Top             =   5295
         Width           =   1470
      End
      Begin VB.Label Lbl_FecEmision 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   89
         Top             =   4900
         Width           =   1095
      End
      Begin VB.Label Lbl_FecDevengue 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   88
         Top             =   5240
         Width           =   1095
      End
      Begin VB.Label Lbl_PolMod 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   83
         Top             =   2520
         Width           =   4095
      End
      Begin VB.Label Lbl_PolTipRta 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   82
         Top             =   1840
         Width           =   4095
      End
      Begin VB.Label Lbl_PolTipPen 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   81
         Top             =   480
         Width           =   5295
      End
      Begin VB.Label Lbl_PolTasaPerGar 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   7080
         TabIndex        =   80
         Top             =   2880
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Lbl_PolTasaVta 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   79
         Top             =   4220
         Width           =   1095
      End
      Begin VB.Label Lbl_PolTasaCto 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   78
         Top             =   3880
         Width           =   1095
      End
      Begin VB.Label Lbl_PolMtoPen 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   77
         Top             =   3540
         Width           =   1455
      End
      Begin VB.Label Lbl_PolMtoPri 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   76
         Top             =   3200
         Width           =   1455
      End
      Begin VB.Label Lbl_PolTerVig 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   4320
         TabIndex        =   75
         Top             =   1160
         Width           =   1215
      End
      Begin VB.Label Lbl_PolIniVig 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   74
         Top             =   1160
         Width           =   1215
      End
      Begin VB.Label Lbl_PolNumCar 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   73
         Top             =   820
         Width           =   735
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Tasa de Int. Per. Garantizado"
         Height          =   255
         Index           =   38
         Left            =   4800
         TabIndex        =   72
         Top             =   2925
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Tasa de Venta"
         Height          =   195
         Index           =   36
         Left            =   240
         TabIndex        =   71
         Top             =   4275
         Width           =   1050
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Tasa Cto. Equivalente"
         Height          =   195
         Index           =   35
         Left            =   240
         TabIndex        =   70
         Top             =   3930
         Width           =   1575
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
         Index           =   32
         Left            =   3960
         TabIndex        =   69
         Top             =   1160
         Width           =   255
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Renta"
         Height          =   195
         Index           =   26
         Left            =   240
         TabIndex        =   68
         Top             =   1890
         Width           =   1020
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Meses Gar."
         Height          =   195
         Index           =   29
         Left            =   240
         TabIndex        =   67
         Top             =   2910
         Width           =   810
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Meses Dif."
         Height          =   195
         Index           =   27
         Left            =   240
         TabIndex        =   66
         Top             =   2235
         Width           =   750
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Monto Pensión"
         Height          =   195
         Index           =   34
         Left            =   240
         TabIndex        =   65
         Top             =   3585
         Width           =   1065
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Monto Prima"
         Height          =   195
         Index           =   33
         Left            =   240
         TabIndex        =   64
         Top             =   3255
         Width           =   885
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Periodo de Vigencia"
         Height          =   195
         Index           =   31
         Left            =   240
         TabIndex        =   63
         Top             =   1215
         Width           =   1425
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Nº Beneficiarios"
         Height          =   195
         Index           =   30
         Left            =   240
         TabIndex        =   62
         Top             =   870
         Width           =   1125
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Estado de Vigencia"
         Height          =   195
         Index           =   25
         Left            =   240
         TabIndex        =   61
         Top             =   1545
         Width           =   1380
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Pensión "
         Height          =   195
         Index           =   24
         Left            =   240
         TabIndex        =   60
         Top             =   525
         Width           =   1200
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Modalidad"
         Height          =   195
         Index           =   28
         Left            =   240
         TabIndex        =   59
         Top             =   2565
         Width           =   735
      End
      Begin VB.Label Lbl_PolEstVig 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   58
         Top             =   1500
         Width           =   4095
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Cta. Individual"
         Height          =   195
         Index           =   1
         Left            =   5640
         TabIndex        =   57
         Top             =   4080
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Ingreso Base"
         Height          =   195
         Index           =   2
         Left            =   5640
         TabIndex        =   56
         Top             =   4440
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Porc. Cubierto"
         Height          =   195
         Index           =   3
         Left            =   5640
         TabIndex        =   55
         Top             =   4800
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "%"
         Height          =   255
         Index           =   4
         Left            =   8040
         TabIndex        =   51
         Top             =   4800
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Lbl_PolMesGar 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   50
         Top             =   2860
         Width           =   855
      End
      Begin VB.Label Lbl_PolMesDif 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   49
         Top             =   2180
         Width           =   855
      End
   End
End
Attribute VB_Name = "Frm_ConsultaEndoso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vlPoliza As String
Dim vlEndoso As Integer
Dim vlOrden As Integer
Dim vlRutAux As String
Dim vlArchivo As String

Dim vlRptCodInsSalud As String 'Código Institución de Salud
Dim vlRptCodAfp As String 'Código de AFP
Dim vlRptCodTipRen As String 'Código de Tipo de Renta (Cobertura)
Dim vlRptCodTipPension As String 'Código de Tipo de Pensión
Dim vlRptNomInter As String 'Nombre Intermediario
Dim vlRptRutInter As String 'Rut Intermediario
Dim vlRptComInter As String 'Comisión Intermediario
'Dim vlRptNumEndosoEnd As Integer
'Dim vlRptNumEndosoPol As Integer
Dim vlRptIdenBen As String
Dim vlRptNomBen As String
Dim vlRptGlsDirBen As String
Dim vlRptCodDireccion As String
Dim vlRptFonoBen As String
Dim vlRptComuna As String
Dim vlRptRegion As String
Dim vlRptProvincia As String
'Dim vlRptGlsCauEndoso As String
'Dim vlRptFechaVigEndoso As String
'Dim vlRptMtoPension As Double
'Dim vlRptMtoPensionOri As Double
'Dim vlRptMtoRtaMod As Double
'Dim vlRptGlsFactorEndoso As String 'Texto Aumenta o Disminuye
'Constantes Reportes
Const clRptAumenta As String = "Aumenta en "
Const clRptDisminuye As String = "Disminuye en "
Const clRptMantiene As String = "Se Mantiene en "

'Variables para datos de Póliza
Dim vlNumPoliza As String
Dim vlNumEndoso As Integer
Dim vlCodTipPension As String
Dim vlCodEstado As String
Dim vlCodTipRen As String
Dim vlCodModalidad As String
Dim vlNumCargas As Integer
Dim vlFecVigencia As String
Dim vlFecTerVigencia As String
Dim vlMtoPrima As Double
Dim vlMtoPension As Double
Dim vlNumMesDif As Integer
Dim vlNumMesGar As Integer
Dim vlPrcTasaCe As Double
Dim vlPrcTasaVta As Double
Dim vlPrcTasaIntPerGar As Double
Dim vlCuspp As String 'MC - 28-08-2007
Dim vlFecEmision As String
Dim vlFecDevengue As String
Dim vlIndCober As String
Dim vlCobCon As String
Dim vlDerCre As String
Dim vlDerGra As String

'Variables para datos de Beneficiarios
'Dim vlNumPoliza As String
'Dim vlNumEndoso As Integer
Dim vlNumOrden As Integer
Dim vlRutBen As Double
Dim vlDgvBen As String

Dim vlGlsNomBen As String
Dim vlGlsNomSegBen As String
Dim vlGlsPatBen As String
Dim vlGlsMatBen As String
Dim vlCodGruFam As String
Dim vlCodPar As String
Dim vlCodSexo As String
Dim vlCodSitInv As String
Dim vlCodDerCre As String
Dim vlCodEstPension As String
Dim vlCodCauInv As String
Dim vlFecNacBen As String
Dim vlFecNacHM As String
Dim vlFecInvBen As String
'Dim vlMtoPension As Double
Dim vlPrcPension As Double
Dim vlFecFallBen As String
Dim vlCodDerpen As String
Dim vlCodMotReqPen As String
Dim vlMtoPensionGar As Double
Dim vlCodCauSusBen As String
Dim vlFecSusBen As String
Dim vlFecIniPagoPen As String
Dim vlFecTerPagoPenGar As String
Dim vlFecMatrimonio As String

Const clCodSinDerPen As String * 2 = "10"
Const clCodParCau As String * 2 = "99"

Const clRptOrigen As String = "Privada ó Previsional"

Dim vlCodTipoIdenBenCau As String
Dim vlNumIdenBenCau As String

Dim vlCodTipoIdenBen As String
Dim vlNumIdenBen As String

Dim vlLargoTipoIden    As Integer 'sirve para llenar la grilla
Dim vlPosicionTipoIden As Integer 'sirve para llenar la grilla

Private Sub Cmb_PenNumIdent_Click()
If (Cmb_PenNumIdent <> "") Then
    vlPosicionTipoIden = Cmb_PenNumIdent.ListIndex
    vlLargoTipoIden = Cmb_PenNumIdent.ItemData(vlPosicionTipoIden)
    If (vlLargoTipoIden = 0) Then
        Txt_PenNumIdent.Text = "0"
        Txt_PenNumIdent.Enabled = False
    Else
        Txt_PenNumIdent = ""
        Txt_PenNumIdent.Enabled = True
        Txt_PenNumIdent.MaxLength = vlLargoTipoIden
        If (Txt_PenNumIdent <> "") Then Txt_PenNumIdent.Text = Mid(Txt_PenNumIdent, 1, vlLargoTipoIden)
    End If
End If
End Sub

Private Sub Cmb_PenNumIdent_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        If (Txt_PenNumIdent.Enabled = True) Then
            Txt_PenNumIdent.SetFocus
        Else
            Cmd_BuscarPol.SetFocus
        End If
    End If
End Sub

Private Sub Cmd_Buscar_Click()
On Error GoTo Err_CmdBuscarClick

    Frm_EndBusqueda.flInicio ("Frm_ConsultaEndoso")
    
Exit Sub
Err_CmdBuscarClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_BuscarPol_Click()
On Error GoTo Err_CmdBuscarPolClick
        
    If Txt_PenPoliza.Text = "" Then
       If ((Trim(Cmb_PenNumIdent.Text)) = "") Or (Txt_PenNumIdent.Text = "") Then
       ''*Or _(Not ValiRut(Txt_PenRut.Text, Txt_PenDigito.Text))
           MsgBox "Debe Ingresar el Número de Póliza o la Identificación del Pensionado.", vbCritical, "Error de Datos"
           Txt_PenNumIdent.SetFocus
           Exit Sub
       Else
           ''Txt_PenRut = Format(Txt_PenRut, "##,###,##0")
           Txt_PenNumIdent = UCase(Trim(Txt_PenNumIdent))
           Txt_PenNumIdent.SetFocus
           ''vlRutAux = Format(Txt_PenRut, "#0")
       End If
    Else
        Txt_PenPoliza.Text = Trim(Txt_PenPoliza.Text)
    End If
    
    vlCodTipoIdenBenCau = fgObtenerCodigo_TextoCompuesto(Cmb_PenNumIdent)
    vlNumIdenBenCau = Txt_PenNumIdent
    
    vgPalabra = ""
    'Seleccionar beneficiario, según número de póliza y Identificación de beneficiario.
    If (Txt_PenPoliza.Text <> "") And (Cmb_PenNumIdent.Text <> "") And (Txt_PenNumIdent.Text <> "") Then
        ''vlRutAux = Format(Txt_PenRut, "#0")
        vgPalabra = "num_poliza = '" & Txt_PenPoliza.Text & "' AND "
        vgPalabra = vgPalabra & "cod_tipoidenBEN = " & (vlCodTipoIdenBenCau) & " and "
        vgPalabra = vgPalabra & "num_idenben = '" & (vlNumIdenBenCau) & "' "
    Else
        'Seleccionar, según número de póliza, el primer beneficiario con derecho a pensión.
        'En caso de no existir, seleccionar sólo el primer beneficiario sin derecho.
        If Txt_PenPoliza.Text <> "" Then
           vgSql = ""
           vgSql = "SELECT COUNT(num_orden) as NumeroBen "
           vgSql = vgSql & "FROM pp_tmae_ben WHERE "
           vgSql = vgSql & "num_poliza = '" & Txt_PenPoliza.Text & "' "
           'If vlTipoBuscar = "N" Then
           'vgSql = vgSql & "AND cod_estpension <> '" & clCodSinDerPen & "' "
           'End If
           vgSql = vgSql & "ORDER BY num_endoso DESC, num_orden ASC "
           Set vgRegistro = vgConexionBD.Execute(vgSql)
           If (vgRegistro!numeroben) <> 0 Then
              vgPalabra = "num_poliza = '" & Txt_PenPoliza.Text & "' "
              'If vlTipoBuscar = "N" Then
              '   vgPalabraAux = vgPalabra & "AND cod_estpension <> '" & clCodSinDerPen & "' "
              'End If
           Else
               vgPalabra = "num_poliza = '" & Txt_PenPoliza.Text & "' "
           End If
        Else
            'Seleccionar beneficiario, según identificación beneficiario. (Datos de primera póliza encontrada.)
            If Txt_PenNumIdent.Text <> "" Then
               vgPalabra = "cod_tipoidenBEN = " & (vlCodTipoIdenBenCau) & " "
               vgPalabra = vgPalabra & "AND num_idenben = '" & (vlNumIdenBenCau) & "' "
            End If
        End If
    End If
    
    'Validar el ingreso obligatorio del número de endoso
    'CMV 20050926 I
    If Txt_PenEndoso.Text = "" Then
       MsgBox "Debe Ingresar Número de Endoso.", vbCritical, "Error de Datos"
       Txt_PenEndoso.SetFocus
       Exit Sub
    Else
        vgPalabra = vgPalabra & " AND num_endoso = " & Txt_PenEndoso & " "
    End If
    'CMV 20050926 F
    
    'Ejecutar selección según los parámetros correspondientes, contenidos en
    'variable vgpalabra
    vgSql = ""
    vgSql = "SELECT num_endoso,num_orden,gls_nomben,gls_nomsegben, gls_patben,gls_matben, "
    vgSql = vgSql & "cod_estpension,cod_tipoidenben,num_idenben,num_poliza "
    vgSql = vgSql & "FROM pp_tmae_ben WHERE "
    vgSql = vgSql & vgPalabra
    vgSql = vgSql & " ORDER BY num_orden ASC "
    Set vgRs2 = vgConexionBD.Execute(vgSql)
    If Not vgRs2.EOF Then
    
        vlCodTipoIdenBenCau = vgRs2!Cod_TipoIdenBen
        vlNumIdenBenCau = Trim(vgRs2!Num_IdenBen)
    
       If Txt_PenPoliza.Text <> "" Then
            Call fgBuscarPosicionCodigoCombo(vlCodTipoIdenBenCau, Cmb_PenNumIdent)
            Txt_PenNumIdent.Text = vlNumIdenBenCau
       Else
           Txt_PenPoliza.Text = Trim(vgRs2!num_poliza)
       End If
              
       Lbl_PenNombre.Caption = Trim(vgRs2!Gls_NomBen) + " " + IIf(IsNull(vgRs2!Gls_NomSegBen), "", Trim(vgRs2!Gls_NomSegBen) + " " + Trim(vgRs2!Gls_PatBen)) + " " + IIf(IsNull(vgRs2!Gls_MatBen), "", Trim(vgRs2!Gls_MatBen))
       'Lbl_PenEndoso.Caption = (vgRs2!num_endoso)
       vlPoliza = (vgRs2!num_poliza)
       vlEndoso = (vgRs2!num_endoso)
       vlOrden = (vgRs2!Num_Orden)
      
       Call flCargaDatosPoliza
       
       Call flInicializaGrillaBenef
       Call flCargaBeneficiarios
       Msf_BenGrilla.row = 1
       Call Msf_BenGrilla_Click
        
       Fra_Poliza.Enabled = False
       SSTab_PolizaOriginal.Enabled = True
       SSTab_PolizaOriginal.Tab = 0
       SSTab_PolizaOriginal.SetFocus
  
    Else
        MsgBox "El Beneficiario o la Póliza Ingresados, No Existen en la Base de Datos", vbInformation, "Información"
        Txt_PenPoliza.SetFocus
        Exit Sub
    End If
       
Exit Sub
Err_CmdBuscarPolClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Cancelar_Click()
On Error GoTo Err_Cmd_Cancelar_Click

    Call flLimpiarFraPenPoliza
    Call flLimpiarPoliza
    Call flLimpiarBeneficiarios
    
    Call flLimpiarVariables

    Fra_Poliza.Enabled = True
    SSTab_PolizaOriginal.Tab = 0
    SSTab_PolizaOriginal.Enabled = False
        
    Txt_PenPoliza.SetFocus
        
Exit Sub
Err_Cmd_Cancelar_Click:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Imprimir_Click() 'MateriaGris-JRios 12/12/2018 Se modificò para replicar el reporte de generaciòn de endosos
Dim vlNombre As String, vlNombreSeg As String
Dim vlPaterno As String, vlMaterno As String
Dim vlNombreCompleto As String, vlNombreTipoIden As String
Dim vlImpCodMoneda As String, vlImpNomMoneda As String

On Error GoTo Err_flImprimirPoliza

    Dim rsLiq As ADODB.Recordset
    Dim objRep As New ClsReporte
    Dim LNGa As Long
'Imprimir Reporte de Poliza Original
   Screen.MousePointer = 11

   'vlArchivo = strRpt & "PP_Rpt_EndDefPoliza.rpt"   '\Reportes
   vlArchivo = strRpt & "PP_Rpt_EndDefEndoso_2.rpt"   '\Reportes Jaime Rios - Materia Gris
   If Not fgExiste(vlArchivo) Then     ', vbNormal
      MsgBox "Archivo de Reporte de Endoso Póliza de Renta Vitalicia no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   vlRptCodTipPension = ""
   'vlRptNumEndoso = 0
   vlRptCodAfp = ""
   vlRptCodTipRen = ""
   
   vgSql = ""
   vgSql = "SELECT num_endoso,cod_afp,cod_tipren,cod_tippension "
   vgSql = vgSql & ",cod_moneda,mto_valmoneda "
   vgSql = vgSql & "FROM PP_TMAE_POLIZA "
   vgSql = vgSql & "WHERE "
   vgSql = vgSql & "num_poliza = '" & vlPoliza & "' AND "
   vgSql = vgSql & "num_endoso = " & vlEndoso & " "
   'vgSql = vgSql & "ORDER BY num_endoso DESC "
   Set vgRegistro = vgConexionBD.Execute(vgSql)
   If Not vgRegistro.EOF Then
      'vlRptNumEndoso = (vgRegistro!num_endoso)
      vlRptCodAfp = Trim(vgRegistro!cod_afp) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_AFP, Trim(vgRegistro!cod_afp)))
      vlRptCodTipRen = Trim(vgRegistro!Cod_TipRen) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_TipRen, Trim(vgRegistro!Cod_TipRen)))
      vlRptCodTipPension = Trim(vgRegistro!Cod_TipPension) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_TipPen, Trim(vgRegistro!Cod_TipPension)))
      vlImpCodMoneda = vgRegistro!Cod_Moneda
      vlImpNomMoneda = fgBuscarGlosaElemento(vgCodTabla_TipMon, vlImpCodMoneda)
   End If
   vgRegistro.Close
   
   vlRptCodInsSalud = ""
   vlRptNomBen = ""
   vlRptIdenBen = ""
   vlRptGlsDirBen = ""
   vlRptFonoBen = ""
   vlRptCodDireccion = "0"
   
   vgSql = ""
   vgSql = "SELECT cod_inssalud,gls_nomben,gls_nomsegben,gls_patben,gls_matben,cod_tipoidenben,"
   vgSql = vgSql & "num_idenben,cod_direccion,gls_dirben,gls_fonoben "
   vgSql = vgSql & "FROM PP_TMAE_BEN "
   vgSql = vgSql & "WHERE "
   vgSql = vgSql & "num_poliza = '" & vlPoliza & "' AND "
   vgSql = vgSql & "num_endoso = " & (vlEndoso) & " AND "
   vgSql = vgSql & "cod_par = '" & Trim(clCodParCau) & "' "
   Set vgRegistro = vgConexionBD.Execute(vgSql)
   If Not vgRegistro.EOF Then
        If Not IsNull(vgRegistro!Cod_InsSalud) Then
            vlRptCodInsSalud = " " & Trim(vgRegistro!Cod_InsSalud) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_InsSal, Trim(vgRegistro!Cod_InsSalud)))
        End If
        If Not IsNull(vgRegistro!Gls_NomBen) And Not IsNull(vgRegistro!Gls_PatBen) Then
            vlNombre = vgRegistro!Gls_NomBen
            vlNombreSeg = IIf(IsNull(vgRegistro!Gls_NomSegBen), "", vgRegistro!Gls_NomSegBen)
            vlPaterno = vgRegistro!Gls_PatBen
            vlMaterno = IIf(IsNull(vgRegistro!Gls_PatBen), "", vgRegistro!Gls_PatBen)
            vlNombreCompleto = fgFormarNombreCompleto(vlNombre, vlNombreSeg, vlPaterno, vlMaterno)
            
            vlRptNomBen = vlNombreCompleto
        End If
        If Not IsNull(vgRegistro!Cod_TipoIdenBen) And Not IsNull(vgRegistro!Num_IdenBen) Then
            vlNombreTipoIden = fgBuscarNombreTipoIden(vgRegistro!Cod_TipoIdenBen)
            vlRptIdenBen = vlNombreTipoIden & " - " & (Trim(vgRegistro!Num_IdenBen))
        End If
        If Not IsNull(vgRegistro!Gls_DirBen) Then
            vlRptGlsDirBen = Trim(vgRegistro!Gls_DirBen)
        End If
        If Not IsNull(vgRegistro!Gls_FonoBen) Then
            vlRptFonoBen = Trim(vgRegistro!Gls_FonoBen)
        Else
            vlRptFonoBen = " "
        End If
        If Not IsNull(vgRegistro!Cod_Direccion) Then
            vlRptCodDireccion = Trim(vgRegistro!Cod_Direccion)
        End If
   End If
   vgRegistro.Close


'RRR DATOS DEL ULTIMO ENDOSO
   Dim vlNumOrden As String
   vlNumOrden = ""
   vgSql = ""
   vgSql = " SELECT NUM_ORDEN"
   vgSql = vgSql & " FROM (SELECT NUM_ORDEN,COD_TIPOIDENBEN,NUM_IDENBEN,GLS_NOMBEN,GLS_NOMSEGBEN,GLS_PATBEN,GLS_MATBEN,"
   vgSql = vgSql & "              Gls_DirBen , Cod_Direccion, Gls_FonoBen, Gls_CorreoBen, Count(1)"
   vgSql = vgSql & "         From PP_TMAE_BEN"
   vgSql = vgSql & "        Where num_poliza = " & vlPoliza & " "
   vgSql = vgSql & "          AND NUM_ENDOSO IN (" & vlEndoso & ", " & vlEndoso & " - 1)"
   vgSql = vgSql & "        GROUP BY NUM_ORDEN,COD_TIPOIDENBEN,NUM_IDENBEN,GLS_NOMBEN,GLS_NOMSEGBEN,GLS_PATBEN,GLS_MATBEN,"
   vgSql = vgSql & "              Gls_DirBen , Cod_Direccion, Gls_FonoBen, Gls_CorreoBen"
   vgSql = vgSql & "        ORDER BY 12)"
   vgSql = vgSql & " Where ROWNUM = 1"
   Set vgRegistro = vgConexionBD.Execute(vgSql)
   If Not vgRegistro.EOF Then
      vlNumOrden = Trim(vgRegistro!Num_Orden)
   End If
  
   'vlRptNomBenNue = ""
   Dim vlRptGlsDirBenNue As String
   vlRptGlsDirBenNue = ""
   Dim vlRptFonoBenNue As String
   vlRptFonoBenNue = ""
   Dim vlRptCodDireccionNue As String
   vlRptCodDireccionNue = "0"
   Dim vlRptNomBeNue As String
   vlRptNomBeNue = ""
   Dim vlRptCorreo As String
   vlRptCorreo = ""
   Dim vlRptDocNum As String
   vlRptDocNum = ""
   
   vgSql = ""
   vgSql = "SELECT cod_inssalud,gls_nomben,gls_nomsegben,gls_patben,gls_matben,Cod_TipoIdenben, "
   vgSql = vgSql & "Num_Idenben,cod_direccion,gls_dirben,gls_fonoben, gls_correoben "
   vgSql = vgSql & "FROM PP_TMAE_BEN "
   vgSql = vgSql & "WHERE "
   vgSql = vgSql & "num_poliza = '" & vlPoliza & "' AND "
   vgSql = vgSql & "num_endoso = " & (vlEndoso) & " AND "
   'vgSql = vgSql & "cod_derpen = '" & Trim(clCodParCau) & "' AND " 'MateriaGris-JRios 11/01/2018
   vgSql = vgSql & "num_orden = " & (vlNumOrden)
   Set vgRegistro = vgConexionBD.Execute(vgSql)
   If Not vgRegistro.EOF Then
       
        'While Not vgRegistro.EOF
        
            If Not IsNull(vgRegistro!Gls_NomBen) And Not IsNull(vgRegistro!Gls_PatBen) Then
                vlNombre = vgRegistro!Gls_NomBen
                vlNombreSeg = IIf(IsNull(vgRegistro!Gls_NomSegBen), "", vgRegistro!Gls_NomSegBen)
                vlPaterno = vgRegistro!Gls_PatBen
                'I - MC 24/01/2008
                ''vlMaterno = IIf(IsNull(vgRegistro!Gls_PatBen), "", vgRegistro!Gls_PatBen)
                vlMaterno = IIf(IsNull(vgRegistro!Gls_MatBen), "", vgRegistro!Gls_MatBen)
                'F - MC 24/01/2008
                vlNombreCompleto = fgFormarNombreCompleto(vlNombre, vlNombreSeg, vlPaterno, vlMaterno)
                
                vlRptNomBeNue = vlNombreCompleto 'vlRptNomBeNue & vbCrLf & vlNombreCompleto
            End If
            'vgRegistro.MoveNext
        'Wend
        'vgRegistro.MoveFirst
          
        If Not IsNull(vgRegistro!Gls_DirBen) Then
            vlRptGlsDirBenNue = Trim(vgRegistro!Gls_DirBen)
        End If
        If Not IsNull(vgRegistro!Gls_FonoBen) Then
            vlRptFonoBenNue = Trim(vgRegistro!Gls_FonoBen)
        Else
            vlRptFonoBenNue = " "
        End If
        If Not IsNull(vgRegistro!Cod_Direccion) Then
            vlRptCodDireccionNue = Trim(vgRegistro!Cod_Direccion)
        End If
        If Not IsNull(vgRegistro!Gls_CorreoBen) Then
            vlRptCorreo = Trim(vgRegistro!Gls_CorreoBen)
        End If
        If Not IsNull(vgRegistro!Num_IdenBen) Then
            vlRptDocNum = Trim(vgRegistro!Num_IdenBen)
        End If
 
   End If
   vgRegistro.Close

    vgSql = ""
    vgSql = "SELECT c.gls_comuna,p.gls_provincia,r.gls_region "
    vgSql = vgSql & "FROM MA_TPAR_COMUNA c,MA_TPAR_PROVINCIA p,MA_TPAR_REGION r "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "c.cod_direccion = '" & vlRptCodDireccionNue & "' AND "
    vgSql = vgSql & "c.cod_provincia = p.cod_provincia AND "
    vgSql = vgSql & "p.cod_region = r.cod_region AND "
    vgSql = vgSql & "c.cod_region = r.cod_region "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
         If Not IsNull(vgRegistro!gls_comuna) Then vlRptComuna = Trim(vgRegistro!gls_comuna)
         If Not IsNull(vgRegistro!gls_provincia) Then vlRptProvincia = Trim(vgRegistro!gls_provincia)
         If Not IsNull(vgRegistro!gls_region) Then vlRptRegion = Trim(vgRegistro!gls_region)
         vlRptGlsDirBenNue = vlRptGlsDirBenNue & " - " & vlRptComuna & "/" & vlRptProvincia & "/" & vlRptRegion
    End If
    vgRegistro.Close

'RRR

   vlRptComuna = ""
   vlRptProvincia = ""
   vlRptRegion = ""
   
    vgSql = ""
    vgSql = "SELECT c.gls_comuna,p.gls_provincia,r.gls_region "
    vgSql = vgSql & "FROM MA_TPAR_COMUNA c,MA_TPAR_PROVINCIA p,MA_TPAR_REGION r "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "c.cod_direccion = '" & vlRptCodDireccion & "' AND "
    vgSql = vgSql & "c.cod_provincia = p.cod_provincia AND "
    vgSql = vgSql & "p.cod_region = r.cod_region AND "
    vgSql = vgSql & "c.cod_region = r.cod_region "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
        If Not IsNull(vgRegistro!gls_comuna) Then vlRptComuna = Trim(vgRegistro!gls_comuna)
        If Not IsNull(vgRegistro!gls_provincia) Then vlRptProvincia = Trim(vgRegistro!gls_provincia)
        If Not IsNull(vgRegistro!gls_region) Then vlRptRegion = Trim(vgRegistro!gls_region)
        'vlRptGlsDirBenNue = vlRptGlsDirBenNue & "-" & vlRptComuna & "/" & vlRptProvincia & "/" & vlRptRegion
    End If
    vgRegistro.Close
    
    vlRptNomInter = "-"
    vlRptRutInter = "-"
    vlRptComInter = "-"
   
    
   Dim vlRptGlsCauEndoso As String
   vlRptGlsCauEndoso = ""
   Dim vlRptGlsFactorEndoso As String
   vlRptGlsFactorEndoso = ""
   Dim vlRptMtoRtaMod As Double
   vlRptMtoRtaMod = 0
   Dim vlRptMtoPension As Double
   vlRptMtoPension = 0
   Dim vlRptFechaVigEndoso As String
   vlRptFechaVigEndoso = ""
   Dim vlRptMtoPensionOri As Double
   vlRptMtoPensionOri = 0
   Dim vl_FlgSRPT As Integer
   vl_FlgSRPT = 0
    
   vgSql = ""
   vgSql = "SELECT mto_pensionori,mto_pensioncal,fec_efecto, "
   vgSql = vgSql & "cod_cauendoso,cod_tipendoso "
   vgSql = vgSql & "FROM PP_TMAE_ENDOSO "
   vgSql = vgSql & "WHERE "
   vgSql = vgSql & "num_poliza = '" & Trim(Txt_PenPoliza) & "' AND "
   'I--- ABV
   'vgSql = vgSql & "num_endoso = " & vlRptNumEndosoEnd & " "
   vgSql = vgSql & "num_endoso = " & vlEndoso - 1 & " "
   'F--- ABV
   
   
   Set vgRegistro = vgConexionBD.Execute(vgSql)
   If Not vgRegistro.EOF Then
      'vlRptNumEndoso = (vgRegistro!num_endoso)
      '****mvg agrego la ultima else
      If vgRegistro!cod_cauendoso = "14" Then
            vlRptGlsCauEndoso = Trim(fgBuscarGlosaElemento(vgCodTabla_CauEnd, Trim(vgRegistro!cod_cauendoso))) & vbCrLf & "Nombres Actualizados: " & vlRptNomBeNue
      ElseIf vgRegistro!cod_cauendoso = "15" Then
            vlRptGlsCauEndoso = Trim(fgBuscarGlosaElemento(vgCodTabla_CauEnd, Trim(vgRegistro!cod_cauendoso))) & vbCrLf & "Nombre    : " & vlRptNomBeNue & vbCrLf & "La nueva direccion es: " & vlRptGlsDirBenNue 'MateriaGris-JRios 12/12/2018
      ElseIf vgRegistro!cod_cauendoso = "16" Then
            vlRptGlsCauEndoso = Trim(fgBuscarGlosaElemento(vgCodTabla_CauEnd, Trim(vgRegistro!cod_cauendoso))) & vbCrLf & "Nombre    : " & vlRptNomBeNue & vbCrLf & "El nuevo Nùmero Telefonico es: " & vlRptFonoBenNue 'MateriaGris-JRios 12/12/2018
      ElseIf vgRegistro!cod_cauendoso = "17" Then
            vlRptGlsCauEndoso = Trim(fgBuscarGlosaElemento(vgCodTabla_CauEnd, Trim(vgRegistro!cod_cauendoso))) & vbCrLf & "Nombre    : " & vlRptNomBeNue & vbCrLf & "El nuevo DNI es: " & vlRptDocNum 'MateriaGris-JRios 12/12/2018
      ElseIf vgRegistro!cod_cauendoso = "27" Then
            vlRptGlsCauEndoso = Trim(fgBuscarGlosaElemento(vgCodTabla_CauEnd, Trim(vgRegistro!cod_cauendoso))) & vbCrLf & "Nombre    : " & vlRptNomBeNue
            vlRptGlsCauEndoso = vlRptGlsCauEndoso & vbCrLf & "Nro. documento : " & vlRptDocNum & vbCrLf & "E-mail       : " & vlRptCorreo
            vlRptGlsCauEndoso = vlRptGlsCauEndoso & vbCrLf & "Direccion : " & vlRptGlsDirBenNue & vbCrLf & "Telèfono   : " & vlRptFonoBenNue
      Else
            vlRptGlsCauEndoso = Trim(fgBuscarGlosaElemento(vgCodTabla_CauEnd, Trim(vgRegistro!cod_cauendoso)))
      End If
      
      
      vlRptFechaVigEndoso = DateSerial(Mid((vgRegistro!FEC_EFECTO), 1, 4), Mid((vgRegistro!FEC_EFECTO), 5, 2), Mid((vgRegistro!FEC_EFECTO), 7, 2))
      vlRptMtoPension = Format((vgRegistro!mto_pensioncal), "###,###,##0.00")
      vlRptMtoPensionOri = Format((vgRegistro!mto_pensionori), "###,###,##0.00")
      If vlRptMtoPensionOri <> vlRptMtoPension Then
         If vlRptMtoPensionOri > vlRptMtoPension Then
            vlRptMtoRtaMod = Format((vlRptMtoPensionOri - vlRptMtoPension), "#0.00")
            vlRptGlsFactorEndoso = Trim(clRptDisminuye)
         End If
         If vlRptMtoPensionOri < vlRptMtoPension Then
            vlRptMtoRtaMod = Format((vlRptMtoPension - vlRptMtoPensionOri), "#0.00")
            vlRptGlsFactorEndoso = Trim(clRptAumenta)
         End If
      Else
          'CMV 20050928 I
          'Modificado para que en el informe muestre: "La Renta se mantiene en"
          'junto al valor de la renta original y no : "La Renta se mantiene en"
          'junto al valor de la modificación que sería 0.
          'vlRptMtoRtaMod = 0
          vlRptMtoRtaMod = vlRptMtoPensionOri
          'CMV 20050928 F
          vlRptGlsFactorEndoso = Trim(clRptMantiene)
      End If
      
      
   End If
     ' vlRptGlsDirBen = Left(vlRptGlsDirBen, 50)
      
      
    vgSql = ""
    vgSql = " select a.num_poliza, a.num_endoso, b.num_idenben, a.fec_vigencia, a.mto_prima, a.mto_pension, a.num_mesdif, a.num_mesgar,"
    vgSql = vgSql & " Gls_NomBen , Gls_NomSegBen, b.Gls_PatBen, b.Gls_MatBen, Cod_Par, Cod_Sexo, Cod_SitInv, b.Mto_Pension as Mto_ben, c.gls_elemento, d.gls_tipoiden, b.fec_nacben,e.cod_tipendoso, b.gls_dirben"
    ',f.gls_elemento AS tipoendoso
    vgSql = vgSql & " from pp_tmae_poliza a"
    vgSql = vgSql & " join pp_tmae_ben b on a.num_poliza=b.num_poliza and a.num_endoso=b.num_endoso"
    vgSql = vgSql & " join ma_tpar_tabcod c on b.cod_par=c.cod_elemento and c.cod_tabla='PA'"
    vgSql = vgSql & " join ma_tpar_tipoiden d on b.cod_tipoidenben=d.cod_tipoiden"
    vgSql = vgSql & " join pp_tmae_endoso e on a.num_poliza=e.num_poliza and a.num_endoso=(e.num_endoso + 1) "
    'vgSql = vgSql & " JOIN ma_tpar_tabcod F on e.cod_cauendoso=f.cod_elemento and f.cod_tabla='CE'"
    vgSql = vgSql & " Where a.num_poliza = '" & vlPoliza & "' And a.num_endoso = " & (vlEndoso) & ""
    vgSql = vgSql & " order by b.num_orden"
      
    Set rsLiq = New ADODB.Recordset
    rsLiq.CursorLocation = adUseClient
    rsLiq.Open vgSql, vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
    'vlRutCliente = flRutCliente 'vgRutCliente + " - " + vgDgvCliente
    
    LNGa = CreateFieldDefFile(rsLiq, Replace(UCase(strRpt & "Estructura\PP_Rpt_EndDefEndoso.rpt"), ".RPT", ".TTX"), 1)

    If objRep.CargaReporte(strRpt & "", "PP_Rpt_EndDefEndoso_2.rpt", "Informe de Liquidación de Rentas Vitalicias", rsLiq, True, _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("NombreSistema", vgNombreSistema), _
                            ArrFormulas("NombreSubSistema", vgNombreSubSistema), _
                            ArrFormulas("TipPension", vlRptCodTipPension), _
                            ArrFormulas("Afp", vlRptCodAfp), _
                            ArrFormulas("TipRta", vlRptCodTipRen), _
                            ArrFormulas("NombreCausante", vlRptNomBen), _
                            ArrFormulas("RutCausante", vlRptIdenBen), _
                            ArrFormulas("Direccion", vlRptGlsDirBen), _
                            ArrFormulas("Fono", vlRptFonoBen), _
                            ArrFormulas("Comuna", vlRptComuna), _
                            ArrFormulas("Provincia", vlRptProvincia), _
                            ArrFormulas("Region", vlRptRegion), _
                            ArrFormulas("Origen", clRptOrigen), _
                            ArrFormulas("CodMoneda", cgCodTipMonedaUF), _
                            ArrFormulas("MotivoEndoso", vlRptGlsCauEndoso), _
                            ArrFormulas("GlsFactorEndoso", vlRptGlsFactorEndoso), _
                            ArrFormulas("MtoRtaMod", str(vlRptMtoRtaMod)), _
                            ArrFormulas("MtoPension", str(vlRptMtoPension)), _
                            ArrFormulas("FechaVigEndoso", vlRptFechaVigEndoso), _
                            ArrFormulas("MtoRtaOri", str(vlRptMtoPensionOri)), _
                            ArrFormulas("CodMonedaCor", vlImpCodMoneda), _
                            ArrFormulas("FlgSRPT", vl_FlgSRPT)) = False Then

        MsgBox "No se pudo abrir el reporte", vbInformation
        'Exit Function
    End If
   Screen.MousePointer = 0

Exit Sub
Err_flImprimirPoliza:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub cmd_salir_Click()
    Screen.MousePointer = 11
    Unload Me
    Screen.MousePointer = 0
End Sub



Private Sub Form_Load()
    Frm_ConsultaEndoso.Top = 0
    Frm_ConsultaEndoso.Left = 0
    SSTab_PolizaOriginal.Tab = 0
    SSTab_PolizaOriginal.Enabled = False
    
    'Carga combo tipo de identificación
    fgComboTipoIdentificacion Cmb_PenNumIdent
        
End Sub

Private Sub Msf_BenGrilla_Click()
On Error GoTo Err_Msf_BenGrilla_Click
Dim vlPos As Integer
Dim vlGlosaPar As String
Dim vlGlosaCauInv As String
        
    Msf_BenGrilla.Col = 0
    Lbl_BenNumOrd = Msf_BenGrilla.Text
    
    Msf_BenGrilla.Col = 1
    ''*vlPos = InStr(Msf_BenGrilla.Text, "-")
    Lbl_BenTipoIden = Trim(Msf_BenGrilla.Text)
    Msf_BenGrilla.Col = 2
    Lbl_BenNumIden = Trim(Msf_BenGrilla.Text)
    ''*Lbl_BenDgvBen = Trim(Mid(Msf_BenGrilla.Text, vlPos + 1, 2))
    
    Msf_BenGrilla.Col = 3
    Lbl_BenNomBen = Trim(Msf_BenGrilla.Text)
    Msf_BenGrilla.Col = 4
    Lbl_BenNomBen = Lbl_BenNomBen & " " & Trim(Msf_BenGrilla.Text)
    Msf_BenGrilla.Col = 5
    Lbl_BenNomBen = Lbl_BenNomBen & " " & Trim(Msf_BenGrilla.Text)
    Msf_BenGrilla.Col = 6
    Lbl_BenNomBen = Lbl_BenNomBen & " " & Trim(Msf_BenGrilla.Text)
        
    Msf_BenGrilla.Col = 7
    vlGlosaPar = " " & Trim(Msf_BenGrilla.Text) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_Par, Trim(Msf_BenGrilla.Text)))
    Lbl_BenPar = vlGlosaPar
    
    Msf_BenGrilla.Col = 8
    Lbl_BenGrupFam = Trim(Msf_BenGrilla.Text)
    
    Msf_BenGrilla.Col = 9
    Lbl_BenSexo = Trim(Msf_BenGrilla.Text)
    
    Msf_BenGrilla.Col = 10
    Lbl_BenSitInv = Trim(Msf_BenGrilla.Text)
    
    Msf_BenGrilla.Col = 11
    Lbl_BenDerPen = Trim(Msf_BenGrilla.Text)
    
    Msf_BenGrilla.Col = 11
    Lbl_BenDerAcrecer = Trim(Msf_BenGrilla.Text)
    
    Msf_BenGrilla.Col = 14
    Lbl_BenCauInv = Trim(Msf_BenGrilla.Text) & " - " & Trim(flBuscarGlosaCauInv(Trim(Msf_BenGrilla.Text)))
        
    Msf_BenGrilla.Col = 15
    Lbl_BenFecNac = Trim(Msf_BenGrilla.Text)
    'Txt_BMFecNac = DateSerial(Mid((Msf_BMGrilla.Text), 1, 4), Mid((Msf_BMGrilla.Text), 5, 2), Mid((Msf_BMGrilla.Text), 7, 2))
        
    Msf_BenGrilla.Col = 16
    If (Msf_BenGrilla.Text) = "" Then
       Lbl_BenFecNHM = ""
    Else
        Lbl_BenFecNHM = Trim(Msf_BenGrilla.Text)
        'Lbl_BMFecNHM = DateSerial(Mid((Msf_BMGrilla.Text), 1, 4), Mid((Msf_BMGrilla.Text), 5, 2), Mid((Msf_BMGrilla.Text), 7, 2))
    End If
        
    Msf_BenGrilla.Col = 17
    If (Msf_BenGrilla.Text) = "" Then
       Lbl_BenFecInv = ""
    Else
        Lbl_BenFecInv = Trim(Msf_BenGrilla.Text)
        'Txt_BMFecInv = DateSerial(Mid((Msf_BMGrilla.Text), 1, 4), Mid((Msf_BMGrilla.Text), 5, 2), Mid((Msf_BMGrilla.Text), 7, 2))
    End If
    
    Msf_BenGrilla.Col = 18
    Lbl_BenMtoPension = Format(Msf_BenGrilla.Text, "###,###,##0.00")
    
    Msf_BenGrilla.Col = 19
    Lbl_BenPrcLegal = Format(Msf_BenGrilla.Text, "##0.00")
    
    Msf_BenGrilla.Col = 20
    If (Msf_BenGrilla.Text) = "" Then
       Lbl_BenFecFall = ""
    Else
        Lbl_BenFecFall = Trim(Msf_BenGrilla.Text)
        'Txt_BMFecFall = DateSerial(Mid((Msf_BMGrilla.Text), 1, 4), Mid((Msf_BMGrilla.Text), 5, 2), Mid((Msf_BMGrilla.Text), 7, 2))
    End If
    
    Msf_BenGrilla.Col = 28
    If (Msf_BenGrilla.Text) = "" Then
       Lbl_BenFecMat = ""
    Else
        Lbl_BenFecMat = Trim(Msf_BenGrilla.Text)
        'Txt_BMFecMat = DateSerial(Mid((Msf_BMGrilla.Text), 1, 4), Mid((Msf_BMGrilla.Text), 5, 2), Mid((Msf_BMGrilla.Text), 7, 2))
    End If
    
''********************
'Dim vlGlosaPar As String
'Dim vlGlosaCauInv As String
'
'    Msf_BenGrilla.Col = 0
'    If (Msf_BenGrilla.Text = "") Or (Msf_BenGrilla.Row = 0) Then
'        MsgBox "No existen Detalles", vbExclamation, "Información"
'        Exit Sub
'    Else
'        'Buscar Glosas
'        vlGlosaPar = " " & Trim(vlCodPar) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_Par, Trim(vlCodPar)))
'        vlGlosaCauInv = " " & Trim(vlCodCauInv) & " - " & Trim(flBuscarGlosaCauInv(Trim(vlCodCauInv)))
'        'Mostrar Datos
'        Lbl_BenNumOrd = vlNumOrden
'        Lbl_BenRutBen = Format((Trim(vlRutBen)), "##,###,##0")
'        Lbl_BenDgvBen = vlDgvBen
'        Lbl_BenNomBen = Trim(vlGlsNomBen) + " " + Trim(vlGlsPatBen) + " " + Trim(vlGlsMatBen)
'        Lbl_BenPar = vlGlosaPar
'        Lbl_BenGrupFam = Trim(vlCodGruFam)
'        Lbl_BenDerPen = Trim(vlCodEstPension)
'        Lbl_BenSexo = Trim(vlCodSexo)
'        Lbl_BenDerAcrecer = Trim(vlCodDerCre)
'        Lbl_BenSitInv = Trim(vlCodSitInv)
'        If (vlFecInvBen) = "" Then
'           Lbl_BenFecInv = ""
'        Else
'            Lbl_BenFecInv = DateSerial(Mid((vlFecInvBen), 1, 4), Mid((vlFecInvBen), 5, 2), Mid((vlFecInvBen), 7, 2))
'        End If
'        Lbl_BenCauInv = vlGlosaCauInv
'        'Lbl_BenFecNac = DateSerial(Mid((vlFecNacBen), 1, 4), Mid((vlFecNacBen), 5, 2), Mid((vlFecNacBen), 7, 2))
'        Lbl_BenFecNac = (vlFecNacBen)
'        If (vlFecFallBen) = "" Then
'           Lbl_BenFecFall = ""
'        Else
'            Lbl_BenFecFall = DateSerial(Mid((vlFecFallBen), 1, 4), Mid((vlFecFallBen), 5, 2), Mid((vlFecFallBen), 7, 2))
'        End If
'        If (vlFecMatrimonio) = "" Then
'           Lbl_BenFecMat = ""
'        Else
'            Lbl_BenFecMat = DateSerial(Mid((vlFecMatrimonio), 1, 4), Mid((vlFecMatrimonio), 5, 2), Mid((vlFecMatrimonio), 7, 2))
'        End If
'        Lbl_BenPrcLegal = Format(vlPrcPension, "##0.00")
'        If (vlFecNacHM) = "" Then
'           Lbl_BenFecNHM = ""
'        Else
'            Lbl_BenFecNHM = DateSerial(Mid((vlFecNacHM), 1, 4), Mid((vlFecNacHM), 5, 2), Mid((vlFecNacHM), 7, 2))
'        End If
'        Lbl_BenMtoPension = Format(vlMtoPension, "###,###,##0.00")
'    End If


Exit Sub
Err_Msf_BenGrilla_Click:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Txt_PenEndoso_Change()
    If Not IsNumeric(Txt_PenEndoso) Then
       Txt_PenEndoso = ""
    End If
End Sub

Private Sub Txt_PenEndoso_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Txt_PenEndoso_KeyPress

    If KeyAscii = 13 Then
       If Txt_PenEndoso.Text = "" Then
          MsgBox "Debe Ingresar Número de Endoso.", vbCritical, "Error de Datos"
          Txt_PenEndoso.SetFocus
          Exit Sub
       End If
       Cmd_BuscarPol.SetFocus
    End If
    
Exit Sub
Err_Txt_PenEndoso_KeyPress:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Txt_PenEndoso_LostFocus()
On Error GoTo Err_Txt_PenEndoso_LostFocus
      
    Txt_PenEndoso.Text = Trim(Txt_PenEndoso.Text)
    If Trim(Txt_PenEndoso = "") Then
       Exit Sub
    End If
        
Exit Sub
Err_Txt_PenEndoso_LostFocus:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Txt_PenNumIdent_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If (Trim(Txt_PenNumIdent) <> "") Then
            Txt_PenNumIdent = UCase(Trim(Txt_PenNumIdent))
        End If
        Txt_PenEndoso.SetFocus
    End If
End Sub

Private Sub txt_pennumident_lostfocus()
    Txt_PenNumIdent = Trim(UCase(Txt_PenNumIdent))
End Sub

Private Sub Txt_PenPoliza_KeyPress(KeyAscii As Integer)
On Error GoTo Err_TxtPenPolizaKeyPress

    If KeyAscii = 13 Then
       Txt_PenPoliza = UCase(Trim(Txt_PenPoliza))
       Txt_PenPoliza = Format(Txt_PenPoliza, "0000000000")
       Cmb_PenNumIdent.SetFocus
    End If
    
Exit Sub
Err_TxtPenPolizaKeyPress:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Txt_PenPoliza_LostFocus()
    Txt_PenPoliza = UCase(Trim(Txt_PenPoliza))
    Txt_PenPoliza = Format(Txt_PenPoliza, "0000000000")
End Sub

Function flRecibe(vlNumPol, vlCodTipoIden, vlNumIden, vlNumEnd)
    Txt_PenPoliza = vlNumPol
    Call fgBuscarPosicionCodigoCombo(vlCodTipoIden, Cmb_PenNumIdent)
    Txt_PenNumIdent = vlNumIden
    Txt_PenEndoso = vlNumEnd
    Call Cmd_BuscarPol_Click
End Function

Function flInicializaGrillaBenef()
On Error GoTo Err_flInicializaGrillaBenef
    
    Msf_BenGrilla.Clear
    Msf_BenGrilla.Cols = 30
    Msf_BenGrilla.rows = 1
    Msf_BenGrilla.RowHeight(0) = 250
    Msf_BenGrilla.row = 0

    Msf_BenGrilla.Col = 0
    Msf_BenGrilla.Text = "Nº Orden"
    Msf_BenGrilla.ColWidth(0) = 700

    Msf_BenGrilla.Col = 1
    Msf_BenGrilla.Text = "Tipo Ident."
    Msf_BenGrilla.ColWidth(1) = 1000

    Msf_BenGrilla.Col = 2
    Msf_BenGrilla.Text = "Nº Ident."
    Msf_BenGrilla.ColWidth(2) = 1000
    
    Msf_BenGrilla.Col = 3
    Msf_BenGrilla.Text = "Nombre"
    Msf_BenGrilla.ColWidth(3) = 1500

    Msf_BenGrilla.Col = 4
    Msf_BenGrilla.Text = "Seg. Nombre"
    Msf_BenGrilla.ColWidth(4) = 1000


    Msf_BenGrilla.Col = 5
    Msf_BenGrilla.Text = "Ap. Paterno"
    Msf_BenGrilla.ColWidth(5) = 1000

    Msf_BenGrilla.Col = 6
    Msf_BenGrilla.Text = "Ap. Materno"
    Msf_BenGrilla.ColWidth(6) = 1000

    Msf_BenGrilla.Col = 7
    Msf_BenGrilla.Text = "Par."
    Msf_BenGrilla.ColWidth(7) = 500

    Msf_BenGrilla.Col = 8
    Msf_BenGrilla.Text = "Gru.Fam."
    Msf_BenGrilla.ColWidth(8) = 700

    Msf_BenGrilla.Col = 9
    Msf_BenGrilla.Text = "Sexo"
    Msf_BenGrilla.ColWidth(9) = 500

    Msf_BenGrilla.Col = 10
    Msf_BenGrilla.Text = "Sit. Inv."
    Msf_BenGrilla.ColWidth(10) = 600

    Msf_BenGrilla.Col = 11
    Msf_BenGrilla.Text = "Dº Pen." 'cod_estpension
    Msf_BenGrilla.ColWidth(11) = 600

    Msf_BenGrilla.Col = 12
    Msf_BenGrilla.Text = "Dº Acrecer"
    Msf_BenGrilla.ColWidth(12) = 800

    Msf_BenGrilla.Col = 13
    Msf_BenGrilla.Text = "num_poliza"
    Msf_BenGrilla.ColWidth(13) = 0

    Msf_BenGrilla.Col = 14
    Msf_BenGrilla.Text = "num_endoso"
    Msf_BenGrilla.ColWidth(14) = 0

    Msf_BenGrilla.Col = 15
    Msf_BenGrilla.Text = "cod_cauinv"
    Msf_BenGrilla.ColWidth(15) = 0

    Msf_BenGrilla.Col = 16
    Msf_BenGrilla.Text = "Fec.Nac."
    Msf_BenGrilla.ColWidth(16) = 900

    Msf_BenGrilla.Col = 17
    Msf_BenGrilla.Text = "Fec.Nac.HM"
    Msf_BenGrilla.ColWidth(17) = 900

    Msf_BenGrilla.Col = 18
    Msf_BenGrilla.Text = "fec_invben"
    Msf_BenGrilla.ColWidth(18) = 0

    Msf_BenGrilla.Col = 19
    Msf_BenGrilla.Text = "Mto.Pensión"
    Msf_BenGrilla.ColWidth(19) = 900

    Msf_BenGrilla.Col = 20
    Msf_BenGrilla.Text = "Prc. Pensión"
    Msf_BenGrilla.ColWidth(20) = 900

    Msf_BenGrilla.Col = 21
    Msf_BenGrilla.Text = "Fec.Fallec."
    Msf_BenGrilla.ColWidth(21) = 900

    Msf_BenGrilla.Col = 22
    Msf_BenGrilla.Text = "cod_derpen"
    Msf_BenGrilla.ColWidth(22) = 0

    Msf_BenGrilla.Col = 23
    Msf_BenGrilla.Text = "cod_motreqpen"
    Msf_BenGrilla.ColWidth(23) = 0

    Msf_BenGrilla.Col = 24
    Msf_BenGrilla.Text = "Mto.Pen.Gar."
    Msf_BenGrilla.ColWidth(24) = 900

    Msf_BenGrilla.Col = 25
    Msf_BenGrilla.Text = "cod_caususben"
    Msf_BenGrilla.ColWidth(25) = 0

    Msf_BenGrilla.Col = 26
    Msf_BenGrilla.Text = "fec_susben"
    Msf_BenGrilla.ColWidth(26) = 0

    Msf_BenGrilla.Col = 27
    Msf_BenGrilla.Text = "fec_inipagopen"
    Msf_BenGrilla.ColWidth(27) = 0

    Msf_BenGrilla.Col = 28
    Msf_BenGrilla.Text = "fec_terpagopengar"
    Msf_BenGrilla.ColWidth(28) = 0

    Msf_BenGrilla.Col = 29
    Msf_BenGrilla.Text = "fec_matrimonio"
    Msf_BenGrilla.ColWidth(29) = 0

Exit Function
Err_flInicializaGrillaBenef:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flCargaBeneficiarios()
On Error GoTo Err_flCargaBeneficiarios

    vgSql = ""
    vgSql = "SELECT num_poliza,num_endoso,num_orden,cod_tipoidenben,num_idenben, "
    vgSql = vgSql & "gls_nomben,gls_nomsegben,gls_patben,gls_matben,cod_grufam, "
    vgSql = vgSql & "cod_par,cod_sexo,cod_sitinv,cod_dercre,cod_derpen, "
    vgSql = vgSql & "cod_cauinv,fec_nacben,fec_nachm,fec_invben, "
    vgSql = vgSql & "mto_pension,prc_pension,fec_fallben,cod_estpension, "
    vgSql = vgSql & "cod_motreqpen,mto_pensiongar,cod_caususben,fec_susben, "
    vgSql = vgSql & "fec_inipagopen,fec_terpagopengar,fec_matrimonio "
    vgSql = vgSql & "FROM pp_tmae_ben WHERE "
    vgSql = vgSql & "num_poliza = '" & Trim(vlPoliza) & "' AND "
    vgSql = vgSql & "num_endoso = " & vlEndoso & " "
    vgSql = vgSql & " ORDER BY num_orden ASC"
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
       While Not vgRegistro.EOF
            If IsNull(vgRegistro!num_poliza) Then
               vlNumPoliza = ""
            Else
                vlNumPoliza = (vgRegistro!num_poliza)
            End If
            If IsNull(vgRegistro!num_endoso) Then
               vlNumEndoso = 0
            Else
                vlNumEndoso = (vgRegistro!num_endoso)
            End If
            If IsNull(vgRegistro!Num_Orden) Then
               vlNumOrden = 0
            Else
                vlNumOrden = (vgRegistro!Num_Orden)
            End If
            If IsNull(vgRegistro!Cod_TipoIdenBen) Then
               vlCodTipoIdenBen = "0"
            Else
                vlCodTipoIdenBen = Trim(vgRegistro!Cod_TipoIdenBen) & " - " & fgBuscarNombreTipoIden(Trim(vgRegistro!Cod_TipoIdenBen), False)
            End If
            If IsNull(vgRegistro!Num_IdenBen) Then
               vlNumIdenBen = ""
            Else
                vlNumIdenBen = (vgRegistro!Num_IdenBen)
            End If
            If IsNull(vgRegistro!Gls_NomBen) Then
               vlGlsNomBen = ""
            Else
                vlGlsNomBen = (vgRegistro!Gls_NomBen)
            End If
            If IsNull(vgRegistro!Gls_NomSegBen) Then
               vlGlsNomSegBen = ""
            Else
                vlGlsNomSegBen = (vgRegistro!Gls_NomSegBen)
            End If
            If IsNull(vgRegistro!Gls_PatBen) Then
               vlGlsPatBen = ""
            Else
                vlGlsPatBen = (vgRegistro!Gls_PatBen)
            End If
            If IsNull(vgRegistro!Gls_MatBen) Then
               vlGlsMatBen = ""
            Else
                vlGlsMatBen = (vgRegistro!Gls_MatBen)
            End If
            If IsNull(vgRegistro!Cod_GruFam) Then
               vlCodGruFam = ""
            Else
                vlCodGruFam = (vgRegistro!Cod_GruFam)
            End If
            If IsNull(vgRegistro!Cod_Par) Then
               vlCodPar = ""
            Else
                vlCodPar = (vgRegistro!Cod_Par)
            End If
            If IsNull(vgRegistro!Cod_Sexo) Then
               vlCodSexo = ""
            Else
                vlCodSexo = (vgRegistro!Cod_Sexo)
            End If
            If IsNull(vgRegistro!Cod_SitInv) Then
               vlCodSitInv = ""
            Else
                vlCodSitInv = (vgRegistro!Cod_SitInv)
            End If
            If IsNull(vgRegistro!Cod_DerCre) Then
               vlCodDerCre = ""
            Else
                vlCodDerCre = (vgRegistro!Cod_DerCre)
            End If
            If IsNull(vgRegistro!Cod_EstPension) Then
               vlCodEstPension = ""
            Else
                vlCodEstPension = (vgRegistro!Cod_EstPension)
            End If
            If IsNull(vgRegistro!Cod_CauInv) Then
               vlCodCauInv = ""
            Else
                vlCodCauInv = (vgRegistro!Cod_CauInv)
            End If
            If IsNull(vgRegistro!Fec_NacBen) Then
               vlFecNacBen = ""
            Else
                vlFecNacBen = (vgRegistro!Fec_NacBen)
            End If
            If IsNull(vgRegistro!Fec_NacHM) Then
               vlFecNacHM = ""
            Else
                vlFecNacHM = (vgRegistro!Fec_NacHM)
            End If
            If IsNull(vgRegistro!Fec_InvBen) Then
               vlFecInvBen = ""
            Else
                vlFecInvBen = (vgRegistro!Fec_InvBen)
            End If
            If IsNull(vgRegistro!Mto_Pension) Then
               vlMtoPension = 0
            Else
                vlMtoPension = (vgRegistro!Mto_Pension)
            End If
            If IsNull(vgRegistro!Prc_Pension) Then
               vlPrcPension = 0
            Else
                vlPrcPension = (vgRegistro!Prc_Pension)
            End If
            If IsNull(vgRegistro!Fec_FallBen) Then
               vlFecFallBen = ""
            Else
                vlFecFallBen = (vgRegistro!Fec_FallBen)
            End If
            If IsNull(vgRegistro!Cod_DerPen) Then
               vlCodDerpen = ""
            Else
                vlCodDerpen = (vgRegistro!Cod_DerPen)
            End If
            If IsNull(vgRegistro!Cod_MotReqPen) Then
               vlCodMotReqPen = ""
            Else
                vlCodMotReqPen = (vgRegistro!Cod_MotReqPen)
            End If
            If IsNull(vgRegistro!Mto_PensionGar) Then
               vlMtoPensionGar = ""
            Else
                vlMtoPensionGar = (vgRegistro!Mto_PensionGar)
            End If
            If IsNull(vgRegistro!Cod_CauSusBen) Then
               vlCodCauSusBen = 0
            Else
                vlCodCauSusBen = (vgRegistro!Cod_CauSusBen)
            End If
            If IsNull(vgRegistro!Fec_SusBen) Then
               vlFecSusBen = ""
            Else
                vlFecSusBen = (vgRegistro!Fec_SusBen)
            End If
            If IsNull(vgRegistro!Fec_IniPagoPen) Then
               vlFecIniPagoPen = ""
            Else
                vlFecIniPagoPen = (vgRegistro!Fec_IniPagoPen)
            End If
            If IsNull(vgRegistro!Fec_TerPagoPenGar) Then
               vlFecTerPagoPenGar = ""
            Else
                vlFecTerPagoPenGar = (vgRegistro!Fec_TerPagoPenGar)
            End If
            If IsNull(vgRegistro!Fec_Matrimonio) Then
               vlFecMatrimonio = ""
            Else
                vlFecMatrimonio = (vgRegistro!Fec_Matrimonio)
            End If
                  
            'Formatear la Fecha de Nacimiento
            If (Trim(vlFecNacBen) <> "") Then
                vlFecNacBen = DateSerial(CInt(Mid(vlFecNacBen, 1, 4)), CInt(Mid(vlFecNacBen, 5, 2)), CInt(Mid(vlFecNacBen, 7, 2)))
            Else
                vlFecNacBen = ""
            End If
            'Formatear la Fecha de Nacimiento del Hijo Menor
            If (Trim(vlFecNacHM) <> "") Then
                vlFecNacHM = DateSerial(CInt(Mid(vlFecNacHM, 1, 4)), CInt(Mid(vlFecNacHM, 5, 2)), CInt(Mid(vlFecNacHM, 7, 2)))
            Else
                vlFecNacHM = ""
            End If
            'Formatear la Fecha de Invalidez
            If (Trim(vlFecInvBen) <> "") Then
                vlFecInvBen = DateSerial(CInt(Mid(vlFecInvBen, 1, 4)), CInt(Mid(vlFecInvBen, 5, 2)), CInt(Mid(vlFecInvBen, 7, 2)))
            Else
                vlFecInvBen = ""
            End If
            'Formatear la Fecha de Fallecimiento
            If (Trim(vlFecFallBen) <> "") Then
                vlFecFallBen = DateSerial(CInt(Mid(vlFecFallBen, 1, 4)), CInt(Mid(vlFecFallBen, 5, 2)), CInt(Mid(vlFecFallBen, 7, 2)))
            Else
                vlFecFallBen = ""
            End If
            'Formatear la Fecha de Suspención del Beneficiario
            If (Trim(vlFecSusBen) <> "") Then
                vlFecSusBen = DateSerial(CInt(Mid(vlFecSusBen, 1, 4)), CInt(Mid(vlFecSusBen, 5, 2)), CInt(Mid(vlFecSusBen, 7, 2)))
            Else
                vlFecSusBen = ""
            End If
            'Formatear la Fecha de Inicio de Pago de Pensiones
            If (Trim(vlFecIniPagoPen) <> "") Then
                vlFecIniPagoPen = DateSerial(CInt(Mid(vlFecIniPagoPen, 1, 4)), CInt(Mid(vlFecIniPagoPen, 5, 2)), CInt(Mid(vlFecIniPagoPen, 7, 2)))
            Else
                vlFecIniPagoPen = ""
            End If
            'Formatear la Fecha de Termino de Pago del Periodo Garantizado
            If (Trim(vlFecTerPagoPenGar) <> "") Then
                vlFecTerPagoPenGar = DateSerial(CInt(Mid(vlFecTerPagoPenGar, 1, 4)), CInt(Mid(vlFecTerPagoPenGar, 5, 2)), CInt(Mid(vlFecTerPagoPenGar, 7, 2)))
            Else
                vlFecTerPagoPenGar = ""
            End If
            'Formatear la Fecha de Termino de Pago del Periodo Garantizado
            If (Trim(vlFecMatrimonio) <> "") Then
                vlFecMatrimonio = DateSerial(CInt(Mid(vlFecMatrimonio, 1, 4)), CInt(Mid(vlFecMatrimonio, 5, 2)), CInt(Mid(vlFecMatrimonio, 7, 2)))
            Else
                vlFecMatrimonio = ""
            End If

            Msf_BenGrilla.AddItem (vlNumOrden) & vbTab _
            & (Trim(vlCodTipoIdenBen)) & vbTab & (Trim(vlNumIdenBen)) & vbTab _
            & (Trim(vlGlsNomBen)) & vbTab & (Trim(vlGlsNomSegBen)) & vbTab & (Trim(vlGlsPatBen)) & vbTab & (Trim(vlGlsMatBen)) & vbTab _
            & (Trim(vlCodPar)) & vbTab _
            & (Trim(vlCodGruFam)) & vbTab _
            & (Trim(vlCodSexo)) & vbTab _
            & (Trim(vlCodSitInv)) & vbTab _
            & (Trim(vlCodEstPension)) & vbTab _
            & (Trim(vlCodDerCre)) & vbTab _
            & (Trim(vlNumPoliza)) & vbTab & (Trim(vlNumEndoso)) & vbTab _
            & (Trim(vlCodCauInv)) & vbTab _
            & (vlFecNacBen) & vbTab _
            & (vlFecNacHM) & vbTab _
            & (vlFecInvBen) & vbTab _
            & ((vlMtoPension)) & vbTab _
            & ((vlPrcPension)) & vbTab _
            & (vlFecFallBen) & vbTab _
            & (Trim(vlCodDerpen)) & vbTab _
            & (Trim(vlCodMotReqPen)) & vbTab _
            & ((vlMtoPensionGar)) & vbTab _
            & (Trim(vlCodCauSusBen)) & vbTab _
            & (vlFecSusBen) & vbTab _
            & (vlFecIniPagoPen) & vbTab _
            & (vlFecTerPagoPenGar) & vbTab _
            & (vlFecMatrimonio)
                  
             vgRegistro.MoveNext
       Wend
    End If

Exit Function
Err_flCargaBeneficiarios:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flCargaDatosPoliza()
On Error GoTo Err_flCargaDatosPoliza

    vgSql = ""
    vgSql = "SELECT num_poliza,num_endoso,cod_tippension,cod_estado, "
    vgSql = vgSql & "cod_tipren,cod_modalidad,num_cargas,fec_vigencia, "
    vgSql = vgSql & "fec_tervigencia,mto_prima,mto_pension,num_mesdif, "
    vgSql = vgSql & "num_mesgar,prc_tasace,prc_tasavta,prc_tasaintpergar, "
    vgSql = vgSql & "cod_cuspp,fec_emision,fec_dev,ind_cob,prc_facpenella,cod_dercre,cod_dergra "
    vgSql = vgSql & "FROM pp_tmae_poliza WHERE "
    vgSql = vgSql & "num_poliza = '" & Trim(vlPoliza) & "' AND "
    vgSql = vgSql & "num_endoso = " & vlEndoso & " "
    vgSql = vgSql & " ORDER BY num_endoso DESC"
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
        If IsNull(vgRegistro!num_poliza) Then
           vlNumPoliza = ""
        Else
            vlNumPoliza = (vgRegistro!num_poliza)
        End If
        If IsNull(vgRegistro!num_endoso) Then
           vlNumEndoso = 0
        Else
            vlNumEndoso = (vgRegistro!num_endoso)
        End If
        If IsNull(vgRegistro!Cod_TipPension) Then
           vlCodTipPension = ""
        Else
            vlCodTipPension = (vgRegistro!Cod_TipPension)
        End If
        If IsNull(vgRegistro!Cod_Estado) Then
           vlCodEstado = ""
        Else
            vlCodEstado = (vgRegistro!Cod_Estado)
        End If
        If IsNull(vgRegistro!Cod_TipRen) Then
           vlCodTipRen = ""
        Else
            vlCodTipRen = (vgRegistro!Cod_TipRen)
        End If
        If IsNull(vgRegistro!Cod_Modalidad) Then
           vlCodModalidad = ""
        Else
            vlCodModalidad = (vgRegistro!Cod_Modalidad)
        End If
        If IsNull(vgRegistro!Num_Cargas) Then
           vlNumCargas = 0
        Else
            vlNumCargas = (vgRegistro!Num_Cargas)
        End If
        If IsNull(vgRegistro!Fec_Vigencia) Then
           vlFecVigencia = ""
        Else
            vlFecVigencia = (vgRegistro!Fec_Vigencia)
        End If
        vlFecTerVigencia = (vgRegistro!Fec_TerVigencia)
        If IsNull(vgRegistro!Fec_TerVigencia) Then
           vlMtoPrima = 0
        Else
            vlMtoPrima = (vgRegistro!Mto_Prima)
        End If
        If IsNull(vgRegistro!Mto_Pension) Then
           vlMtoPension = 0
        Else
            vlMtoPension = (vgRegistro!Mto_Pension)
        End If
        If IsNull(vgRegistro!Num_MesDif) Then
           vlNumMesDif = 0
        Else
            vlNumMesDif = (vgRegistro!Num_MesDif)
        End If
        If IsNull(vgRegistro!Num_MesGar) Then
           vlNumMesGar = 0
        Else
            vlNumMesGar = (vgRegistro!Num_MesGar)
        End If
        If IsNull(vgRegistro!Prc_TasaCe) Then
           vlPrcTasaCe = 0
        Else
            vlPrcTasaCe = (vgRegistro!Prc_TasaCe)
        End If
        If IsNull(vgRegistro!Prc_TasaVta) Then
           vlPrcTasaVta = 0
        Else
            vlPrcTasaVta = (vgRegistro!Prc_TasaVta)
        End If
        If IsNull(vgRegistro!Prc_TasaIntPerGar) Then
           vlPrcTasaIntPerGar = 0
        Else
            vlPrcTasaIntPerGar = (vgRegistro!Prc_TasaIntPerGar)
        End If
        If IsNull(vgRegistro!Cod_Cuspp) Then
           vlCuspp = ""
        Else
            vlCuspp = (vgRegistro!Cod_Cuspp)
        End If
        If IsNull(vgRegistro!Fec_Emision) Then
           vlFecEmision = ""
        Else
            vlFecEmision = (vgRegistro!Fec_Emision)
        End If
        If IsNull(vgRegistro!fec_dev) Then
           vlFecDevengue = ""
        Else
            vlFecDevengue = (vgRegistro!fec_dev)
        End If
        If IsNull(vgRegistro!Ind_Cob) Then
           vlIndCober = ""
        Else
            vlIndCober = (vgRegistro!Ind_Cob)
        End If
        If IsNull(vgRegistro!Prc_FacPenElla) Then
           vlCobCon = ""
        Else
            vlCobCon = (vgRegistro!Prc_FacPenElla)
        End If
        If IsNull(vgRegistro!Cod_DerCre) Then
           vlDerCre = ""
        Else
            vlDerCre = (vgRegistro!Cod_DerCre)
        End If
        If IsNull(vgRegistro!Cod_DerGra) Then
           vlDerGra = ""
        Else
            vlDerGra = (vgRegistro!Cod_DerGra)
        End If
        
        'Buscar Glosas
        vlCodTipPension = " " & Trim(vlCodTipPension) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_TipPen, Trim(vlCodTipPension)))
        vlCodEstado = " " & Trim(vlCodEstado) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_TipVigPol, Trim(vlCodEstado)))
        vlCodTipRen = " " & Trim(vlCodTipRen) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_TipRen, Trim(vlCodTipRen)))
        vlCodModalidad = " " & Trim(vlCodModalidad) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_AltPen, Trim(vlCodModalidad)))
        'Mostrar Datos
        Lbl_PolTipPen = Trim(vlCodTipPension)
        Lbl_PolNumCar = vlNumCargas
        Lbl_PolIniVig = DateSerial(Mid((vlFecVigencia), 1, 4), Mid((vlFecVigencia), 5, 2), Mid((vlFecVigencia), 7, 2))
        Lbl_PolTerVig = DateSerial(Mid((vlFecTerVigencia), 1, 4), Mid((vlFecTerVigencia), 5, 2), Mid((vlFecTerVigencia), 7, 2))
        Lbl_PolEstVig = Trim(vlCodEstado)
        Lbl_PolTipRta = Trim(vlCodTipRen)
        Lbl_PolMesDif = vlNumMesDif
        Lbl_PolMod = Trim(vlCodModalidad)
        Lbl_PolMesGar = vlNumMesGar
        Lbl_PolMtoPri = Format(vlMtoPrima, "###,###,##0.00")
        Lbl_PolMtoPen = Format(vlMtoPension, "###,###,##0.00")
        Lbl_PolTasaCto = Format(vlPrcTasaCe, "##0.00")
        Lbl_PolTasaVta = Format(vlPrcTasaVta, "##0.00")
        Lbl_PolTasaPerGar = Format(vlPrcTasaIntPerGar, "##0.00")
        Lbl_PolCuspp = vlCuspp
        Lbl_FecEmision = vlFecEmision
        Lbl_FecDevengue = vlFecDevengue
        Lbl_POIndCobertura = vlIndCober
        Lbl_POCoberCon = vlCobCon
        Lbl_PODerCrecer = vlDerCre
        Lbl_PODerGratificacion = vlDerGra
        
    End If

Exit Function
Err_flCargaDatosPoliza:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flLimpiarFraPenPoliza()
On Error GoTo Err_flLimpiarFraPenPoliza

    Txt_PenPoliza = ""
    If (Cmb_PenNumIdent.ListCount <> 0) Then
        Cmb_PenNumIdent.ListIndex = 0
    End If
    Txt_PenNumIdent = ""
    Txt_PenEndoso = ""
    Lbl_PenNombre = ""

Exit Function
Err_flLimpiarFraPenPoliza:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flLimpiarPoliza()
On Error GoTo Err_flLimpiarPoliza

    Lbl_PolTipPen = ""
    Lbl_PolNumCar = ""
    Lbl_PolIniVig = ""
    Lbl_PolTerVig = ""
    Lbl_PolEstVig = ""
    Lbl_PolTipRta = ""
    Lbl_PolMesDif = ""
    Lbl_PolMod = ""
    Lbl_PolMesGar = ""
    Lbl_PolMtoPri = ""
    Lbl_PolMtoPen = ""
    Lbl_PolTasaCto = ""
    Lbl_PolTasaVta = ""
    Lbl_PolTasaPerGar = ""
    Lbl_PolCuspp = ""
    Lbl_FecEmision = ""
    Lbl_FecDevengue = ""
    Lbl_POIndCobertura = ""
    Lbl_POCoberCon = ""
    Lbl_PODerCrecer = ""
    Lbl_PODerGratificacion = ""

Exit Function
Err_flLimpiarPoliza:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flLimpiarBeneficiarios()
On Error GoTo Err_flLimpiarBeneficiarios

    Lbl_BenNumOrd = ""
    Lbl_BenNomBen = ""
    Lbl_BenPar = ""
    Lbl_BenGrupFam = ""
    Lbl_BenSexo = ""
    Lbl_BenSitInv = ""
    Lbl_BenCauInv = ""
    Lbl_BenFecNac = ""
    Lbl_BenFecFall = ""
    Lbl_BenFecMat = ""
    Lbl_BenFecNHM = ""
'''    Lbl_BenRutBen = ""
'''    Lbl_BenDgvBen = ""
    Lbl_BenDerPen = ""
    Lbl_BenDerAcrecer = ""
    Lbl_BenFecInv = ""
    Lbl_BenPrcLegal = ""
    Lbl_BenMtoPension = ""

Exit Function
Err_flLimpiarBeneficiarios:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flLimpiarVariables()
On Error GoTo Err_flLimpiarVariables

vlPoliza = ""
vlEndoso = 0
vlOrden = 0

'Variables para datos de Póliza
vlNumPoliza = ""
vlNumEndoso = 0
vlCodTipPension = ""
vlCodEstado = ""
vlCodTipRen = ""
vlCodModalidad = ""
vlNumCargas = 0
vlFecVigencia = ""
vlFecTerVigencia = ""
vlMtoPrima = 0
vlMtoPension = 0
vlNumMesDif = 0
vlNumMesGar = 0
vlPrcTasaCe = 0
vlPrcTasaVta = 0
vlPrcTasaIntPerGar = 0

'Variables para datos de Beneficiarios
'Dim vlNumPoliza As String
'Dim vlNumEndoso As Integer
vlNumOrden = 0
vlRutBen = 0
vlDgvBen = ""
vlGlsNomBen = ""
vlGlsPatBen = ""
vlGlsMatBen = ""
vlCodGruFam = ""
vlCodPar = ""
vlCodSexo = ""
vlCodSitInv = ""
vlCodDerCre = ""
vlCodEstPension = ""
vlCodCauInv = ""
vlFecNacBen = ""
vlFecNacHM = ""
vlFecInvBen = ""
'Dim vlMtoPension As Double
vlPrcPension = 0
vlFecFallBen = ""
vlCodDerpen = ""
vlCodMotReqPen = ""
vlMtoPensionGar = 0
vlCodCauSusBen = ""
vlFecSusBen = ""
vlFecIniPagoPen = ""
vlFecTerPagoPenGar = ""
vlFecMatrimonio = ""

Exit Function
Err_flLimpiarVariables:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flBuscarGlosaCauInv(iElemento) As String
On Error GoTo Err_flBuscarGlosaCauInv

    flBuscarGlosaCauInv = ""
    
    vgSql = ""
    vgSql = "SELECT gls_patologia "
    vgSql = vgSql & "FROM MA_TPAR_PATOLOGIA WHERE "
    vgSql = vgSql & "cod_patologia = '" & iElemento & "' "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
        flBuscarGlosaCauInv = (vgRegistro!gls_patologia)
    End If
    vgRegistro.Close
    
Exit Function
Err_flBuscarGlosaCauInv:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function



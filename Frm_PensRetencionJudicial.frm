VERSION 5.00
Begin VB.Form Frm_PensRetencionJudicial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retención Judicial"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   9750
   Begin VB.CommandButton Command4 
      Caption         =   "Grabar"
      Height          =   450
      Left            =   285
      TabIndex        =   60
      Top             =   8445
      Width           =   1200
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Eliminar"
      Height          =   465
      Left            =   1785
      TabIndex        =   59
      Top             =   8445
      Width           =   1050
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cerrar"
      Height          =   420
      Left            =   8430
      TabIndex        =   58
      Top             =   8475
      Width           =   945
   End
   Begin VB.TextBox Text15 
      Height          =   315
      Left            =   1170
      TabIndex        =   23
      Top             =   120
      Width           =   1755
   End
   Begin VB.TextBox Text14 
      BackColor       =   &H00E0FFFF&
      Height          =   285
      Left            =   8700
      TabIndex        =   22
      Top             =   585
      Width           =   705
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00E0FFFF&
      Height          =   345
      Left            =   6375
      TabIndex        =   21
      Top             =   990
      Width           =   2985
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00E0FFFF&
      Height          =   330
      Left            =   3810
      TabIndex        =   20
      Top             =   1005
      Width           =   1530
   End
   Begin VB.TextBox Text5 
      Height          =   345
      Left            =   1230
      TabIndex        =   19
      Top             =   1005
      Width           =   1185
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0FFFF&
      Height          =   345
      Left            =   4260
      TabIndex        =   18
      Top             =   105
      Width           =   5145
   End
   Begin VB.TextBox Text3 
      Height          =   345
      Left            =   3060
      TabIndex        =   17
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0FFFF&
      Height          =   330
      Left            =   2295
      TabIndex        =   16
      Top             =   570
      Width           =   1650
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0FFFF&
      Height          =   315
      Left            =   5190
      TabIndex        =   15
      Top             =   585
      Width           =   1755
   End
   Begin VB.Frame Frame1 
      Caption         =   "Retención Judicial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6465
      Left            =   270
      TabIndex        =   0
      Top             =   1620
      Width           =   9090
      Begin VB.Frame Frame2 
         Caption         =   "Forma de Pago de Pensión"
         Height          =   1305
         Left            =   270
         TabIndex        =   43
         Top             =   4935
         Width           =   8655
         Begin VB.ComboBox Combo5 
            Height          =   315
            Left            =   990
            TabIndex        =   47
            Top             =   735
            Width           =   2940
         End
         Begin VB.ComboBox Combo4 
            Height          =   315
            Left            =   990
            TabIndex        =   46
            Top             =   315
            Width           =   2955
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            Left            =   5250
            TabIndex        =   45
            Top             =   285
            Width           =   2940
         End
         Begin VB.TextBox Text12 
            Height          =   315
            Left            =   5250
            TabIndex        =   44
            Top             =   675
            Width           =   2910
         End
         Begin VB.Label Label13 
            Caption         =   "Tipo Cta."
            Height          =   270
            Left            =   165
            TabIndex        =   51
            Top             =   795
            Width           =   825
         End
         Begin VB.Label Label10 
            Caption         =   "Vía Pago"
            Height          =   285
            Left            =   165
            TabIndex        =   50
            Top             =   405
            Width           =   810
         End
         Begin VB.Label Label14 
            Caption         =   "Banco"
            Height          =   270
            Left            =   4440
            TabIndex        =   49
            Top             =   315
            Width           =   825
         End
         Begin VB.Label Label12 
            Caption         =   "N°Cuenta"
            Height          =   270
            Left            =   4395
            TabIndex        =   48
            Top             =   720
            Width           =   795
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Antecedentes Retención"
         Height          =   2085
         Left            =   270
         TabIndex        =   32
         Top             =   285
         Width           =   8655
         Begin VB.CommandButton Command1 
            Caption         =   "Seleccionar carga"
            Height          =   360
            Left            =   4380
            TabIndex        =   57
            Top             =   1590
            Width           =   1665
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Si"
            Height          =   360
            Left            =   3735
            TabIndex        =   56
            Top             =   1590
            Width           =   600
         End
         Begin VB.OptionButton Option1 
            Caption         =   "No"
            Height          =   225
            Left            =   2775
            TabIndex        =   55
            Top             =   1665
            Width           =   750
         End
         Begin VB.ComboBox Combo8 
            Height          =   315
            Left            =   1800
            TabIndex        =   53
            Top             =   285
            Width           =   2505
         End
         Begin VB.TextBox Text9 
            BackColor       =   &H00E0FFFF&
            Height          =   345
            Left            =   7410
            TabIndex        =   42
            Top             =   195
            Width           =   1080
         End
         Begin VB.TextBox Text8 
            Height          =   330
            Left            =   5775
            TabIndex        =   39
            Top             =   195
            Width           =   1020
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   1800
            TabIndex        =   38
            Top             =   1140
            Width           =   4440
         End
         Begin VB.TextBox Text13 
            Height          =   345
            Left            =   4530
            TabIndex        =   34
            Top             =   705
            Width           =   1620
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   1800
            TabIndex        =   33
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label Label11 
            Caption         =   "Rentención con carga familiar:"
            Height          =   285
            Left            =   135
            TabIndex        =   54
            Top             =   1635
            Width           =   2460
         End
         Begin VB.Label Label9 
            Caption         =   "Tipo de Retención"
            Height          =   270
            Left            =   150
            TabIndex        =   52
            Top             =   315
            Width           =   1620
         End
         Begin VB.Label Label8 
            Caption         =   "Hasta"
            Height          =   285
            Left            =   6840
            TabIndex        =   41
            Top             =   225
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "Vigencia:  Desde"
            Height          =   345
            Left            =   4410
            TabIndex        =   40
            Top             =   255
            Width           =   1320
         End
         Begin VB.Label Label1 
            Caption         =   "Monto"
            Height          =   270
            Left            =   3975
            TabIndex        =   37
            Top             =   780
            Width           =   795
         End
         Begin VB.Label Label19 
            Caption         =   "Moneda"
            Height          =   270
            Left            =   165
            TabIndex        =   36
            Top             =   765
            Width           =   1425
         End
         Begin VB.Label Label20 
            Caption         =   "Juzgado de Menores"
            Height          =   285
            Left            =   150
            TabIndex        =   35
            Top             =   1155
            Width           =   1500
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Antecedenes Personales"
         Height          =   2340
         Left            =   285
         TabIndex        =   1
         Top             =   2490
         Width           =   8625
         Begin VB.TextBox Text16 
            Height          =   300
            Left            =   960
            TabIndex        =   8
            Top             =   300
            Width           =   1620
         End
         Begin VB.ComboBox Combo6 
            Height          =   315
            Left            =   975
            TabIndex        =   7
            Top             =   1500
            Width           =   2955
         End
         Begin VB.ComboBox Combo7 
            Height          =   315
            Left            =   4905
            TabIndex        =   6
            Top             =   1485
            Width           =   2835
         End
         Begin VB.TextBox Text17 
            Height          =   285
            Left            =   975
            TabIndex        =   5
            Top             =   1875
            Width           =   1500
         End
         Begin VB.TextBox Text18 
            Height          =   345
            Left            =   2700
            TabIndex        =   4
            Top             =   285
            Width           =   405
         End
         Begin VB.TextBox Text19 
            Height          =   285
            Left            =   960
            TabIndex        =   3
            Top             =   690
            Width           =   6750
         End
         Begin VB.TextBox Text20 
            Height          =   330
            Left            =   975
            TabIndex        =   2
            Top             =   1080
            Width           =   6750
         End
         Begin VB.Label Label21 
            Caption         =   "Rut"
            Height          =   285
            Left            =   180
            TabIndex        =   14
            Top             =   405
            Width           =   810
         End
         Begin VB.Label Label22 
            Caption         =   "Comuna"
            Height          =   270
            Left            =   180
            TabIndex        =   13
            Top             =   1530
            Width           =   825
         End
         Begin VB.Label Label23 
            Caption         =   "Ciudad"
            Height          =   225
            Left            =   4290
            TabIndex        =   12
            Top             =   1530
            Width           =   540
         End
         Begin VB.Label Label24 
            Caption         =   "Teléfono"
            Height          =   270
            Left            =   180
            TabIndex        =   11
            Top             =   1920
            Width           =   795
         End
         Begin VB.Label Label25 
            Caption         =   "Nombre"
            Height          =   255
            Left            =   165
            TabIndex        =   10
            Top             =   735
            Width           =   795
         End
         Begin VB.Label Label26 
            Caption         =   "Dirección"
            Height          =   240
            Left            =   150
            TabIndex        =   9
            Top             =   1140
            Width           =   735
         End
      End
   End
   Begin VB.Label Label17 
      Caption         =   "Rut"
      Height          =   285
      Left            =   420
      TabIndex        =   31
      Top             =   135
      Width           =   645
   End
   Begin VB.Label Label16 
      Caption         =   "Inválido"
      Height          =   255
      Left            =   8025
      TabIndex        =   30
      Top             =   600
      Width           =   585
   End
   Begin VB.Label Label15 
      Caption         =   "Parentesco"
      Height          =   315
      Left            =   5445
      TabIndex        =   29
      Top             =   1050
      Width           =   1185
   End
   Begin VB.Label Label6 
      Caption         =   "Tipo de Pensión"
      Height          =   345
      Left            =   2520
      TabIndex        =   28
      Top             =   1035
      Width           =   1290
   End
   Begin VB.Label Label5 
      Caption         =   "N° Póliza"
      Height          =   375
      Left            =   435
      TabIndex        =   27
      Top             =   1020
      Width           =   825
   End
   Begin VB.Label Label4 
      Caption         =   "Nombre"
      Height          =   285
      Left            =   3570
      TabIndex        =   26
      Top             =   180
      Width           =   990
   End
   Begin VB.Line Line2 
      X1              =   330
      X2              =   9375
      Y1              =   1485
      Y2              =   1470
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha de Nacimiento"
      Height          =   225
      Left            =   405
      TabIndex        =   25
      Top             =   600
      Width           =   1800
   End
   Begin VB.Label Label2 
      Caption         =   "Sexo"
      Height          =   195
      Left            =   4635
      TabIndex        =   24
      Top             =   615
      Width           =   540
   End
End
Attribute VB_Name = "Frm_PensRetencionJudicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command3_Click()
    Unload Me
    
End Sub

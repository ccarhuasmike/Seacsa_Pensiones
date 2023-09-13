VERSION 5.00
Begin VB.Form Frm_GEConcPensionado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conciliación por pensionado"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8955
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   8955
   Begin VB.CommandButton Command6 
      Caption         =   "Imprimir histórico"
      Height          =   495
      Left            =   105
      TabIndex        =   69
      Top             =   4710
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   7425
      TabIndex        =   68
      Top             =   4725
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   870
      TabIndex        =   61
      Top             =   675
      Width           =   840
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3000
      TabIndex        =   60
      Top             =   660
      Width           =   1545
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   825
      TabIndex        =   59
      Top             =   255
      Width           =   1485
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2535
      TabIndex        =   58
      Top             =   270
      Width           =   405
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3780
      TabIndex        =   57
      Top             =   300
      Width           =   5040
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   5430
      TabIndex        =   56
      Top             =   690
      Width           =   2085
   End
   Begin VB.TextBox Text17 
      Height          =   285
      Left            =   8250
      TabIndex        =   55
      Top             =   690
      Width           =   555
   End
   Begin VB.Frame Frame1 
      Caption         =   "Conciliacion histórica"
      Height          =   3045
      Left            =   120
      TabIndex        =   0
      Top             =   1335
      Width           =   8580
      Begin VB.TextBox Text86 
         Height          =   255
         Left            =   6030
         TabIndex        =   53
         Top             =   2115
         Width           =   675
      End
      Begin VB.TextBox Text85 
         Height          =   255
         Left            =   6840
         TabIndex        =   52
         Top             =   2130
         Width           =   675
      End
      Begin VB.TextBox Text84 
         Height          =   255
         Left            =   7665
         TabIndex        =   51
         Top             =   2145
         Width           =   720
      End
      Begin VB.TextBox Text83 
         Height          =   240
         Left            =   180
         TabIndex        =   40
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text82 
         Height          =   240
         Left            =   1275
         TabIndex        =   39
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text81 
         Height          =   240
         Left            =   2325
         TabIndex        =   38
         Top             =   840
         Width           =   405
      End
      Begin VB.TextBox Text80 
         Height          =   240
         Left            =   2880
         TabIndex        =   37
         Top             =   840
         Width           =   675
      End
      Begin VB.TextBox Text79 
         Height          =   255
         Left            =   3675
         TabIndex        =   36
         Top             =   840
         Width           =   675
      End
      Begin VB.TextBox Text78 
         Height          =   255
         Left            =   4485
         TabIndex        =   35
         Top             =   840
         Width           =   675
      End
      Begin VB.TextBox Text77 
         Height          =   240
         Left            =   5295
         TabIndex        =   34
         Top             =   840
         Width           =   645
      End
      Begin VB.TextBox Text76 
         Height          =   255
         Left            =   6015
         TabIndex        =   33
         Top             =   840
         Width           =   675
      End
      Begin VB.TextBox Text75 
         Height          =   255
         Left            =   6825
         TabIndex        =   32
         Top             =   855
         Width           =   675
      End
      Begin VB.TextBox Text74 
         Height          =   255
         Left            =   7650
         TabIndex        =   31
         Top             =   870
         Width           =   720
      End
      Begin VB.TextBox Text73 
         Height          =   255
         Left            =   3690
         TabIndex        =   30
         Top             =   1125
         Width           =   675
      End
      Begin VB.TextBox Text72 
         Height          =   240
         Left            =   2895
         TabIndex        =   29
         Top             =   1125
         Width           =   675
      End
      Begin VB.TextBox Text71 
         Height          =   240
         Left            =   2340
         TabIndex        =   28
         Top             =   1125
         Width           =   405
      End
      Begin VB.TextBox Text70 
         Height          =   240
         Left            =   1290
         TabIndex        =   27
         Top             =   1125
         Width           =   975
      End
      Begin VB.TextBox Text69 
         Height          =   240
         Left            =   195
         TabIndex        =   26
         Top             =   1125
         Width           =   975
      End
      Begin VB.TextBox Text68 
         Height          =   255
         Left            =   4500
         TabIndex        =   25
         Top             =   1125
         Width           =   675
      End
      Begin VB.TextBox Text67 
         Height          =   240
         Left            =   5310
         TabIndex        =   24
         Top             =   1125
         Width           =   645
      End
      Begin VB.TextBox Text66 
         Height          =   255
         Left            =   6030
         TabIndex        =   23
         Top             =   1125
         Width           =   675
      End
      Begin VB.TextBox Text65 
         Height          =   255
         Left            =   6840
         TabIndex        =   22
         Top             =   1140
         Width           =   675
      End
      Begin VB.TextBox Text64 
         Height          =   255
         Left            =   7665
         TabIndex        =   21
         Top             =   1155
         Width           =   720
      End
      Begin VB.TextBox Text63 
         Height          =   255
         Left            =   3690
         TabIndex        =   20
         Top             =   1425
         Width           =   675
      End
      Begin VB.TextBox Text62 
         Height          =   240
         Left            =   2895
         TabIndex        =   19
         Top             =   1425
         Width           =   675
      End
      Begin VB.TextBox Text61 
         Height          =   240
         Left            =   2340
         TabIndex        =   18
         Top             =   1425
         Width           =   405
      End
      Begin VB.TextBox Text60 
         Height          =   240
         Left            =   1290
         TabIndex        =   17
         Top             =   1425
         Width           =   975
      End
      Begin VB.TextBox Text59 
         Height          =   240
         Left            =   195
         TabIndex        =   16
         Top             =   1425
         Width           =   975
      End
      Begin VB.TextBox Text58 
         Height          =   255
         Left            =   4500
         TabIndex        =   15
         Top             =   1425
         Width           =   675
      End
      Begin VB.TextBox Text57 
         Height          =   240
         Left            =   5310
         TabIndex        =   14
         Top             =   1425
         Width           =   645
      End
      Begin VB.TextBox Text56 
         Height          =   255
         Left            =   6030
         TabIndex        =   13
         Top             =   1425
         Width           =   675
      End
      Begin VB.TextBox Text55 
         Height          =   255
         Left            =   6840
         TabIndex        =   12
         Top             =   1440
         Width           =   675
      End
      Begin VB.TextBox Text54 
         Height          =   255
         Left            =   7665
         TabIndex        =   11
         Top             =   1455
         Width           =   720
      End
      Begin VB.TextBox Text53 
         Height          =   255
         Left            =   3690
         TabIndex        =   10
         Top             =   1710
         Width           =   675
      End
      Begin VB.TextBox Text52 
         Height          =   240
         Left            =   2895
         TabIndex        =   9
         Top             =   1710
         Width           =   675
      End
      Begin VB.TextBox Text51 
         Height          =   240
         Left            =   2340
         TabIndex        =   8
         Top             =   1710
         Width           =   405
      End
      Begin VB.TextBox Text50 
         Height          =   240
         Left            =   1290
         TabIndex        =   7
         Top             =   1710
         Width           =   975
      End
      Begin VB.TextBox Text49 
         Height          =   240
         Left            =   195
         TabIndex        =   6
         Top             =   1710
         Width           =   975
      End
      Begin VB.TextBox Text48 
         Height          =   255
         Left            =   4500
         TabIndex        =   5
         Top             =   1710
         Width           =   675
      End
      Begin VB.TextBox Text47 
         Height          =   240
         Left            =   5310
         TabIndex        =   4
         Top             =   1710
         Width           =   645
      End
      Begin VB.TextBox Text46 
         Height          =   255
         Left            =   6030
         TabIndex        =   3
         Top             =   1710
         Width           =   675
      End
      Begin VB.TextBox Text45 
         Height          =   255
         Left            =   6840
         TabIndex        =   2
         Top             =   1725
         Width           =   675
      End
      Begin VB.TextBox Text44 
         Height          =   255
         Left            =   7665
         TabIndex        =   1
         Top             =   1740
         Width           =   720
      End
      Begin VB.Label Label11 
         Caption         =   "Totales"
         Height          =   255
         Left            =   5235
         TabIndex        =   54
         Top             =   2160
         Width           =   555
      End
      Begin VB.Label Label10 
         Caption         =   "Tesorería"
         Height          =   180
         Left            =   6000
         TabIndex        =   50
         Top             =   615
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Compañía"
         Height          =   180
         Left            =   6795
         TabIndex        =   49
         Top             =   615
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "% Deduc."
         Height          =   180
         Left            =   5250
         TabIndex        =   48
         Top             =   615
         Width           =   690
      End
      Begin VB.Label Label7 
         Caption         =   "P.Mínima"
         Height          =   255
         Left            =   4470
         TabIndex        =   47
         Top             =   600
         Width           =   750
      End
      Begin VB.Label Label6 
         Caption         =   "Pens. $"
         Height          =   180
         Left            =   3780
         TabIndex        =   46
         Top             =   600
         Width           =   675
      End
      Begin VB.Label Label5 
         Caption         =   "Pens. UF"
         Height          =   195
         Left            =   2895
         TabIndex        =   45
         Top             =   600
         Width           =   765
      End
      Begin VB.Label Label4 
         Caption         =   "Días"
         Height          =   165
         Left            =   2340
         TabIndex        =   44
         Top             =   600
         Width           =   480
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Período"
         Height          =   195
         Left            =   1365
         TabIndex        =   43
         Top             =   615
         Width           =   705
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Resolución"
         Height          =   195
         Left            =   150
         TabIndex        =   42
         Top             =   585
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Diferencia"
         Height          =   180
         Left            =   7695
         TabIndex        =   41
         Top             =   630
         Width           =   735
      End
   End
   Begin VB.Line Line2 
      X1              =   90
      X2              =   8715
      Y1              =   4575
      Y2              =   4575
   End
   Begin VB.Line Line1 
      X1              =   165
      X2              =   8805
      Y1              =   1095
      Y2              =   1095
   End
   Begin VB.Label Label16 
      Caption         =   "N° Póliza"
      Height          =   255
      Left            =   105
      TabIndex        =   67
      Top             =   660
      Width           =   735
   End
   Begin VB.Label Label15 
      Caption         =   "Tipo de Pensión"
      Height          =   210
      Left            =   1770
      TabIndex        =   66
      Top             =   690
      Width           =   1185
   End
   Begin VB.Label Label14 
      Caption         =   "Rut"
      Height          =   210
      Left            =   150
      TabIndex        =   65
      Top             =   270
      Width           =   285
   End
   Begin VB.Label Label13 
      Caption         =   "Nombre"
      Height          =   270
      Left            =   3060
      TabIndex        =   64
      Top             =   300
      Width           =   660
   End
   Begin VB.Label Label12 
      Caption         =   "Parentesco"
      Height          =   225
      Left            =   4605
      TabIndex        =   63
      Top             =   705
      Width           =   870
   End
   Begin VB.Label Label17 
      Caption         =   "Inválido"
      Height          =   270
      Left            =   7620
      TabIndex        =   62
      Top             =   720
      Width           =   600
   End
End
Attribute VB_Name = "Frm_GEConcPensionado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

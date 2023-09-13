VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "Crystl32.OCX"
Begin VB.Form Frm_AFReliquidacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reliquidaciones"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   11370
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   120
      TabIndex        =   9
      Top             =   2625
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   6588
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      ForeColor       =   4210752
      TabCaption(0)   =   "Beneficiarios Afectados"
      TabPicture(0)   =   "Frm_AFReliquidacion.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Rpt_Reliquidacion"
      Tab(0).Control(1)=   "Frm_Beneficiarios"
      Tab(0).Control(2)=   "Msf_Beneficiarios"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Pagos Reliquidados por Beneficiario"
      TabPicture(1)   =   "Frm_AFReliquidacion.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frm_Pagos"
      Tab(1).Control(1)=   "Msf_Pagos"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Modalidad de Pago"
      TabPicture(2)   =   "Frm_AFReliquidacion.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frm_Modalidad"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Msf_Modalidad"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin MSFlexGridLib.MSFlexGrid Msf_Beneficiarios 
         Height          =   2655
         Left            =   -74880
         TabIndex        =   65
         Top             =   960
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   4683
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         BackColor       =   14745599
         WordWrap        =   -1  'True
         AllowUserResizing=   1
         FormatString    =   $"Frm_AFReliquidacion.frx":0054
      End
      Begin MSFlexGridLib.MSFlexGrid Msf_Pagos 
         Height          =   2655
         Left            =   -74880
         TabIndex        =   64
         Top             =   960
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   4683
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         BackColor       =   14745599
         AllowUserResizing=   1
         FormatString    =   $"Frm_AFReliquidacion.frx":00F0
      End
      Begin MSFlexGridLib.MSFlexGrid Msf_Modalidad 
         Height          =   2655
         Left            =   120
         TabIndex        =   63
         Top             =   960
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   4683
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         BackColor       =   14745599
         AllowUserResizing=   1
         FormatString    =   $"Frm_AFReliquidacion.frx":018C
      End
      Begin VB.Frame Frm_Modalidad 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1695
         Left            =   120
         TabIndex        =   51
         Top             =   360
         Width           =   10935
         Begin VB.CommandButton Cmd_ModRestar 
            Height          =   405
            Left            =   10440
            Picture         =   "Frm_AFReliquidacion.frx":021D
            Style           =   1  'Graphical
            TabIndex        =   62
            ToolTipText     =   "Quitar Beneficiario"
            Top             =   1200
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CommandButton Cmd_ModSumar 
            Height          =   405
            Left            =   10440
            Picture         =   "Frm_AFReliquidacion.frx":03A7
            Style           =   1  'Graphical
            TabIndex        =   61
            ToolTipText     =   "Agregar Beneficiario"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txt_ModUltCuota 
            BackColor       =   &H00E0FFFF&
            Height          =   285
            Left            =   9000
            TabIndex        =   60
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txt_ModCuota 
            BackColor       =   &H00E0FFFF&
            Height          =   285
            Left            =   8040
            TabIndex        =   59
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txt_ModTotal 
            BackColor       =   &H00E0FFFF&
            Height          =   285
            Left            =   7080
            TabIndex        =   58
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txt_ModMoneda 
            BackColor       =   &H00E0FFFF&
            Height          =   285
            Left            =   6120
            TabIndex        =   57
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txt_ModFecFin 
            BackColor       =   &H00E0FFFF&
            Height          =   285
            Left            =   5160
            TabIndex        =   56
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txt_ModNumCuotas 
            Height          =   285
            Left            =   4440
            MaxLength       =   3
            TabIndex        =   55
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txt_ModFecInicio 
            Height          =   285
            Left            =   3360
            MaxLength       =   10
            TabIndex        =   54
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txt_ModConcepto 
            BackColor       =   &H00E0FFFF&
            Height          =   285
            Left            =   840
            TabIndex        =   53
            Top             =   240
            Width           =   2535
         End
         Begin VB.TextBox txt_ModNumOrden 
            BackColor       =   &H00E0FFFF&
            Height          =   285
            Left            =   0
            TabIndex        =   52
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frm_Pagos 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1695
         Left            =   -74880
         TabIndex        =   39
         Top             =   360
         Width           =   10935
         Begin VB.TextBox txt_PagMontoAct 
            Height          =   285
            Left            =   7440
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txt_PagMontoAnt 
            BackColor       =   &H00E0FFFF&
            Height          =   285
            Left            =   6360
            Locked          =   -1  'True
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txt_PagConcepto 
            BackColor       =   &H00E0FFFF&
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   240
            Width           =   2895
         End
         Begin VB.CommandButton Cmd_PagRestar 
            Height          =   405
            Left            =   10440
            Picture         =   "Frm_AFReliquidacion.frx":0531
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Quitar Beneficiario"
            Top             =   1200
            Width           =   495
         End
         Begin VB.CommandButton Cmd_PagSumar 
            Height          =   405
            Left            =   10440
            Picture         =   "Frm_AFReliquidacion.frx":06BB
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Agregar Beneficiario"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txt_PagPeriodo 
            BackColor       =   &H00E0FFFF&
            Height          =   285
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txt_PagFecInicio 
            BackColor       =   &H00E0FFFF&
            Height          =   285
            Left            =   4440
            Locked          =   -1  'True
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txt_PagFecFin 
            BackColor       =   &H00E0FFFF&
            Height          =   285
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txt_PagMonto 
            BackColor       =   &H00E0FFFF&
            Height          =   285
            Left            =   8400
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txt_PagMoneda 
            BackColor       =   &H00E0FFFF&
            Height          =   285
            Left            =   9240
            Locked          =   -1  'True
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txt_PagNumOrden 
            BackColor       =   &H00E0FFFF&
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frm_Beneficiarios 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1695
         Left            =   -74880
         TabIndex        =   28
         Top             =   360
         Width           =   10935
         Begin VB.TextBox txt_BenPerHasta 
            BackColor       =   &H00E0FFFF&
            Height          =   285
            Left            =   9000
            TabIndex        =   34
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txt_BenPerDesde 
            BackColor       =   &H00E0FFFF&
            Height          =   285
            Left            =   7920
            TabIndex        =   33
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txt_BenPar 
            BackColor       =   &H00E0FFFF&
            Height          =   285
            Left            =   6960
            Locked          =   -1  'True
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txt_BenNom 
            BackColor       =   &H00E0FFFF&
            Height          =   285
            Left            =   4080
            Locked          =   -1  'True
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   240
            Width           =   2895
         End
         Begin VB.TextBox txt_BenNroIden 
            BackColor       =   &H00E0FFFF&
            Height          =   285
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txt_BenTipoIden 
            BackColor       =   &H00E0FFFF&
            Height          =   285
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Cmd_BenRestar 
            Height          =   405
            Left            =   10440
            Picture         =   "Frm_AFReliquidacion.frx":0845
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Quitar Beneficiario"
            Top             =   1200
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CommandButton cmd_BenSumar 
            Height          =   405
            Left            =   10440
            Picture         =   "Frm_AFReliquidacion.frx":09CF
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Agregar Beneficiario"
            Top             =   720
            Width           =   495
         End
         Begin VB.CheckBox chk_Pension 
            Enabled         =   0   'False
            Height          =   195
            Left            =   720
            TabIndex        =   36
            Top             =   240
            Width           =   375
         End
         Begin VB.CheckBox chk_Reliq 
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txt_BenNumOrden 
            BackColor       =   &H00E0FFFF&
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   240
            Width           =   735
         End
      End
      Begin Crystal.CrystalReport Rpt_Reliquidacion 
         Left            =   -74640
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.Frame Fra_Poliza 
      Caption         =   "Póliza / Pensionado que recibe el Beneficio"
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
      TabIndex        =   19
      Top             =   0
      Width           =   11175
      Begin VB.TextBox Txt_PenNumIdent 
         Height          =   285
         Left            =   5400
         MaxLength       =   16
         TabIndex        =   67
         Top             =   360
         Width           =   1875
      End
      Begin VB.ComboBox Cmb_PenNumIdent 
         BackColor       =   &H00E0FFFF&
         Height          =   315
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   360
         Width           =   2235
      End
      Begin VB.CommandButton Cmd_BuscarPol 
         Height          =   375
         Left            =   10080
         Picture         =   "Frm_AFReliquidacion.frx":0B59
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Buscar Póliza"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Txt_PenPoliza 
         Height          =   285
         Left            =   1080
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
         Left            =   10080
         Picture         =   "Frm_AFReliquidacion.frx":0C5B
         TabIndex        =   3
         ToolTipText     =   "Buscar"
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Ident."
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   68
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Lbl_Endoso 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   8400
         TabIndex        =   27
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nº Endoso"
         Height          =   255
         Index           =   12
         Left            =   7320
         TabIndex        =   26
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   21
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Póliza"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Lbl_PenNombre 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   1
         Top             =   720
         Width           =   6735
      End
   End
   Begin VB.Frame Fra_Operacion 
      Height          =   1095
      Index           =   0
      Left            =   120
      TabIndex        =   18
      Top             =   6360
      Width           =   11175
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   7680
         Picture         =   "Frm_AFReliquidacion.frx":0D5D
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   5640
         Picture         =   "Frm_AFReliquidacion.frx":0E57
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Imprimir Reporte Montos Cargas Familiares"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   10320
         Picture         =   "Frm_AFReliquidacion.frx":1511
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Eliminar 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   3720
         Picture         =   "Frm_AFReliquidacion.frx":1BCB
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Eliminar Vigencia"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Grabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   1560
         Picture         =   "Frm_AFReliquidacion.frx":1F0D
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Grabar datos de Vigencia"
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Fra_Antecedentes 
      Caption         =   "Antecedentes Reliquidación"
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
      Height          =   1455
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   11175
      Begin VB.CommandButton cmdCalcular 
         Height          =   450
         Left            =   10080
         Picture         =   "Frm_AFReliquidacion.frx":25C7
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Calcular Porcentajes"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txt_Comentarios 
         Height          =   615
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   8535
      End
      Begin VB.TextBox txt_FecReliq 
         BackColor       =   &H00E0FFFF&
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txt_NumReliq 
         BackColor       =   &H00E0FFFF&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txt_PerHasta 
         BackColor       =   &H00E0FFFF&
         Height          =   285
         Left            =   8880
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txt_PerDesde 
         BackColor       =   &H00E0FFFF&
         Height          =   285
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Comentarios    :"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Número Reliq. :"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Reliq. :"
         Height          =   255
         Left            =   2640
         TabIndex        =   22
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "hasta"
         Height          =   255
         Left            =   8280
         TabIndex        =   17
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Periodo a Reliquidar (mm/aaaa)"
         Height          =   375
         Left            =   5040
         TabIndex        =   16
         Top             =   360
         Width           =   2415
      End
   End
End
Attribute VB_Name = "Frm_AFReliquidacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variables con las que se llama a la función flObtieneBeneficiarios
Public vpIndAF As Integer
Public vpIndGE As Integer
Public vpIndPension As Integer
Public vpReliquidacion As Integer
Public vpNumOrden As Integer
Public vpNumOrdenRec As Integer
Public vpSQLUpdate As String
Public vpSQLWhere As String
Public vpEfecto As Date 'Fecha de Efecto
Public vpAccion As String 'M: Modificar, C: Consultar, N: Nueva
Public vpEstado As String 'A: Abierto, C: Cerrado (Estado del Formulario)
Dim vlNumOrden As Long 'Orden del Beneficiario que recibe la Asignación Familiar
Dim vlArrayPagos() As stPagos
Dim vlConceptoAFMenor As String, vlConceptoAFMayor As String 'Conceptos de Reliquidación de Asignación Familiar
Dim vlConceptoModif As String
Dim vlConceptoPEMenor As String, vlConceptoPEMayor As String 'Conceptos de Reliquidación de Pensión
Dim vlConceptoGEMenor As String, vlConceptoGEMayor As String 'Conceptos de Reliquidación de Garantía Estatal

Function flActualizaReliquidación(iNumReliq)
    'Actualiza Antecedentes de la Reliquidación
End Function

Function flAcumulaPagos()
    Dim vlNumOrden As Long
    Dim vlConcepto As String
    Dim vlMoneda As String
    Dim vlMonto As Double
    Dim vlFila As Long
    Dim i, j As Integer
    
    'Modalidad de Pago
    vlFila = 1
    For i = 1 To Msf_Pagos.Rows - 2
        Msf_Pagos.Row = i
        Msf_Pagos.Col = 0
        vlNumOrden = Msf_Pagos
        Msf_Pagos.Col = 2
        vlConcepto = Msf_Pagos
        Msf_Pagos.Col = 8
        vlMoneda = Msf_Pagos
        vlMonto = 0
        If vlArrayPagos(i).Ind_Acumulado = 0 Then
            vlMonto = vlArrayPagos(i).monto
            For j = i + 1 To Msf_Pagos.Rows - 2
                If vlArrayPagos(j).Ind_Acumulado = 0 Then
                    If vlArrayPagos(j).Num_Orden = vlNumOrden Then
                        If vlArrayPagos(j).Concepto = vlConcepto And vlArrayPagos(j).Moneda = vlMoneda Then
                            vlMonto = vlMonto + vlArrayPagos(j).monto
                            vlArrayPagos(j).Ind_Acumulado = 1
                        End If
                    Else
                        Exit For
                    End If
                End If
            Next j
            
            If vlMonto > 0 Then
                Msf_Modalidad.AddItem vlNumOrden & vbTab & vlConcepto & vbTab & _
                vpEfecto & vbTab & "" & vbTab & "" & vbTab & vlMoneda & vbTab & vlMonto & _
                vbTab & "" & vbTab & "", vlFila
                vlFila = vlFila + 1
            End If
            vlMonto = 0
        End If
    Next i
End Function

Function flEliminaReliquidacion(iNumReliq As Double) As Boolean
On Error GoTo Errores
Dim vlSQLBase As String, vlSql As String
Dim vlSQLHab As String, vlSQLBaseHab As String 'Para los Haberes y Descuentos
Dim vlFecInicio As Date, vlFecTermino As Date 'Para el último periodo
Dim vlUltCuota As Boolean
Dim vlNumCuotas As Long
Dim vlMontoTotas As Double
Dim i As Integer
flEliminaReliquidacion = False

'Actualiza Tabla por la cual se hizo la Reliquidación
'con la query traspasada desde el Formulario que Llama a la Reliquidación
vlSql = "UPDATE pp_tmae_certificado"
vlSql = vlSql & " SET num_reliq = 0"
vlSql = vlSql & " WHERE num_reliq = " & iNumReliq
vgConexionTransac.Execute (vlSql)

'Elimina Haber y Descuento en Tabla General de Haberes y Descuentos
vlSql = "DELETE FROM pp_tmae_habdes"
vlSql = vlSql & " WHERE num_reliq = " & iNumReliq
vgConexionTransac.Execute (vlSql)

'Elimina Modalidad de Pago
vlSql = "DELETE FROM pp_tmae_detpagoreliq"
vlSql = vlSql & " WHERE num_reliq = " & iNumReliq
vgConexionTransac.Execute (vlSql)

'Elimina Beneficiarios reliquidados
vlSql = "DELETE FROM pp_tmae_detcalcreliq"
vlSql = vlSql & " WHERE num_reliq = " & iNumReliq
vgConexionTransac.Execute (vlSql)

'Elimina Beneficiarios de la Reliquidacion
vlSql = "DELETE FROM pp_tmae_benreliq"
vlSql = vlSql & " WHERE num_reliq = " & iNumReliq
vgConexionTransac.Execute (vlSql)

'Elimina Reliquidacion
vlSql = "DELETE FROM pp_tmae_reliq"
vlSql = vlSql & " WHERE num_reliq = " & iNumReliq
vgConexionTransac.Execute (vlSql)

flEliminaReliquidacion = True

Errores:
    If Err.Number <> 0 Then
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End If
End Function

Function flGrabaReliquidacion(iNumReliq) As Boolean
On Error GoTo Errores
Dim vlSQLBase As String, vlSql As String
Dim vlSQLHab As String, vlSQLBaseHab As String 'Para los Haberes y Descuentos
Dim vlFecInicio As Date, vlFecTermino As Date 'Para el último periodo
Dim vlUltCuota As Boolean
Dim vlNumCuotas As Long
Dim vlMontoTotas As Double
Dim i As Integer
flGrabaReliquidacion = False

'Graba la Reliquidación
vlSql = "INSERT INTO pp_tmae_reliq"
vlSql = vlSql & " (num_reliq, num_poliza, num_endoso, fec_reliq,"
vlSql = vlSql & " num_perdesde, num_perhasta, gls_observacion)"
vlSql = vlSql & " VALUES ("
vlSql = vlSql & iNumReliq & ",'" & Txt_PenPoliza & "',"
If vgTipoBase = "ORACLE" Then
    vlSql = vlSql & Lbl_Endoso & ",TO_CHAR(SYSDATE,'YYYYMMDD')" & ",'"
Else
    vlSql = vlSql & Lbl_Endoso & ",CONVERT(CHAR(8),GETDATE(),112)" & ",'"
End If
vlSql = vlSql & Mid(txt_PerDesde, 4, 4) & Mid(txt_PerDesde, 1, 2) & "','" & Mid(txt_PerHasta, 4, 4) & Mid(txt_PerHasta, 1, 2) & "','" & txt_Comentarios & "')"
vgConexionTransac.Execute (vlSql)

'Graba Beneficiarios de la Reliquidacion
vlSQLBase = "INSERT INTO pp_tmae_benreliq"
vlSQLBase = vlSQLBase & " (num_reliq, num_poliza, num_endoso, num_orden,"
vlSQLBase = vlSQLBase & " cod_indreliq, num_perdesde, num_perhasta, "
vlSQLBase = vlSQLBase & " cod_indpension)"
vlSQLBase = vlSQLBase & " VALUES ("
vlSQLBase = vlSQLBase & iNumReliq & ",'" & Txt_PenPoliza & "',"
vlSQLBase = vlSQLBase & Lbl_Endoso & ","
For i = 1 To Msf_Beneficiarios.Rows - 2
    Msf_Beneficiarios.Row = i
    Msf_Beneficiarios.Col = 2 'Num Orden
    vlSql = Msf_Beneficiarios.Text & ","
    Msf_Beneficiarios.Col = 0 'Ind Reliq
    If Msf_Beneficiarios.Text = "X" Then
        vlSql = vlSql & "'1',"
    Else
        vlSql = vlSql & "'0',"
    End If
    Msf_Beneficiarios.Col = 7 'PerDesde
    vlSql = vlSql & "'" & Mid(Msf_Beneficiarios.Text, 4, 4) & Mid(Msf_Beneficiarios.Text, 1, 2) & "',"
    Msf_Beneficiarios.Col = 8 'PerHasta
    vlSql = vlSql & "'" & Mid(Msf_Beneficiarios.Text, 4, 4) & Mid(Msf_Beneficiarios.Text, 1, 2) & "',"
    Msf_Beneficiarios.Col = 1 'Ind Pension
    If Msf_Beneficiarios.Text = "X" Then
        vlSql = vlSql & "'1')"
    Else
        vlSql = vlSql & "'0')"
    End If
    vlSql = vlSQLBase & vlSql
    vgConexionTransac.Execute (vlSql)
Next i
        
'Graba Beneficiarios reliquidados
vlSQLBase = "INSERT INTO pp_tmae_detcalcreliq"
vlSQLBase = vlSQLBase & " (num_reliq, num_orden, num_perpago,"
vlSQLBase = vlSQLBase & " cod_conhabdes, fec_inipago, fec_terpago,"
vlSQLBase = vlSQLBase & " mto_conhabdes, cod_moneda, mto_conhabdesant,"
vlSQLBase = vlSQLBase & " mto_diferencia)"
vlSQLBase = vlSQLBase & " VALUES ("
vlSQLBase = vlSQLBase & iNumReliq & ","
For i = 1 To Msf_Pagos.Rows - 2
    Msf_Pagos.Row = i
    Msf_Pagos.Col = 0 'Num Orden
    vlSql = Msf_Pagos.Text & " ,'"
    Msf_Pagos.Col = 1 'Num Perpago
    vlSql = vlSql & Mid(Msf_Pagos.Text, 4, 4) & Mid(Msf_Pagos.Text, 1, 2) & "' ,"
    Msf_Pagos.Col = 2 'Código Concepto
    vlSql = vlSql & "'" & Trim(Mid(Msf_Pagos.Text, 1, InStr(1, Msf_Pagos.Text, "-") - 1)) & "',"
    Msf_Pagos.Col = 3 'Fecha Inicio
    vlSql = vlSql & "'" & Format(Msf_Pagos.Text, "YYYYMMDD") & "',"
    Msf_Pagos.Col = 4 'Fecha Termino
    vlSql = vlSql & "'" & Format(Msf_Pagos.Text, "YYYYMMDD") & "',"
    Msf_Pagos.Col = 6 'Monto Concepto
    vlSql = vlSql & Str(Msf_Pagos.Text) & ",'"
    Msf_Pagos.Col = 8 'Moneda
    'vlSQL = vlSQL & Trim(Mid(Msf_Pagos.Text, 1, InStr(1, Msf_Pagos, "-") - 1)) & "',"
    vlSql = vlSql & Trim(Msf_Pagos.Text) & "',"
    Msf_Pagos.Col = 5 'Monto Anterior
    vlSql = vlSql & Str(Msf_Pagos.Text) & ","
    Msf_Pagos.Col = 7 'Diferencia
    vlSql = vlSql & Str(Msf_Pagos.Text) & ")"
    vlSql = vlSQLBase & vlSql
    vgConexionTransac.Execute (vlSql)
Next i
    
'Graba Modalidad de Pago
vlSQLBase = "INSERT INTO pp_tmae_detpagoreliq"
vlSQLBase = vlSQLBase & " (num_reliq, num_orden, cod_conhabdes,"
vlSQLBase = vlSQLBase & " fec_inihabdes, fec_terhabdes, num_cuotas,"
vlSQLBase = vlSQLBase & " mto_cuota, mto_total, cod_moneda, mto_ultcuota)"
vlSQLBase = vlSQLBase & " VALUES ("
vlSQLBase = vlSQLBase & iNumReliq & ","
For i = 1 To Msf_Modalidad.Rows - 2
    Msf_Modalidad.Row = i
    Msf_Modalidad.Col = 0 'Num Orden
    vlSql = Msf_Modalidad.Text & " ,"
    Msf_Modalidad.Col = 1 'Código Concepto
    vlSql = vlSql & "'" & Trim(Mid(Msf_Modalidad.Text, 1, InStr(1, Msf_Modalidad.Text, "-") - 1)) & "',"
    Msf_Modalidad.Col = 2 'Fecha Inicio
    vlSql = vlSql & "'" & Format(Msf_Modalidad.Text, "YYYYMMDD") & "',"
    Msf_Modalidad.Col = 4 'Fecha Termino
    vlSql = vlSql & "'" & Format(Msf_Modalidad.Text, "YYYYMMDD") & "',"
    Msf_Modalidad.Col = 3 'Num Cuotas
    vlSql = vlSql & Msf_Modalidad.Text & ","
    Msf_Modalidad.Col = 7 'Monto Cuota
    vlSql = vlSql & Str(Msf_Modalidad.Text) & ","
    Msf_Modalidad.Col = 6 'Monto Cuotas
    vlSql = vlSql & Str(Msf_Modalidad.Text) & ",'"
    Msf_Modalidad.Col = 5 'Moneda
    vlSql = vlSql & Trim(Msf_Modalidad.Text) & "',"
    Msf_Modalidad.Col = 8 'Monto Ultima Cuota
    vlSql = vlSql & Str(Msf_Modalidad.Text) & ")"
    vlSql = vlSQLBase & vlSql
    vgConexionTransac.Execute (vlSql)
Next i

'Graba Haber y Descuento en Tabla General de Haberes y Descuentos
vlSQLBaseHab = "INSERT INTO pp_tmae_habdes"
vlSQLBaseHab = vlSQLBaseHab & " (num_poliza, num_endoso, num_orden, "
vlSQLBaseHab = vlSQLBaseHab & " cod_conhabdes, fec_inihabdes,"
vlSQLBaseHab = vlSQLBaseHab & " fec_terhabdes, num_cuotas, mto_cuota, "
vlSQLBaseHab = vlSQLBaseHab & " mto_total, cod_moneda, cod_motsushabdes, fec_sushabdes,"
vlSQLBaseHab = vlSQLBaseHab & " gls_obshabdes, cod_usuariocrea, fec_crea,"
vlSQLBaseHab = vlSQLBaseHab & " hor_crea, num_reliq) VALUES ('"
vlSQLBaseHab = vlSQLBaseHab & Txt_PenPoliza & "'," & Lbl_Endoso & ","
For i = 1 To Msf_Modalidad.Rows - 2
    Msf_Modalidad.Row = i
    Msf_Modalidad.Col = 8
    If CDbl(Msf_Modalidad.Text) > 0 Then
        vlUltCuota = True
    Else
        vlUltCuota = False
    End If
    Msf_Modalidad.Col = 0 'Num Orden
    vlSql = Msf_Modalidad.Text & ","
    Msf_Modalidad.Col = 1 'Código Concepto
    vlSql = vlSql & "'" & Trim(Mid(Msf_Modalidad.Text, 1, InStr(1, Msf_Modalidad.Text, "-") - 1)) & "',"
    Msf_Modalidad.Col = 2 'Fecha Inicio
    vlSql = vlSql & "'" & Format(Msf_Modalidad.Text, "YYYYMMDD") & "',"
    
    Msf_Modalidad.Col = 4 'Fecha Termino
    If vlUltCuota Then
        vlFecTermino = DateAdd("d", -1, DateAdd("m", -1, DateAdd("d", 1, Msf_Modalidad)))
    Else
        vlFecTermino = Msf_Modalidad.Text
    End If
    
    vlSql = vlSql & "'" & Format(vlFecTermino, "YYYYMMDD") & "',"
    Msf_Modalidad.Col = 3 'Num Cuotas
    If vlUltCuota Then
        vlNumCuotas = CLng(Msf_Modalidad.Text - 1)
    Else
        vlNumCuotas = CLng(Msf_Modalidad.Text)
    End If
    vlSql = vlSql & vlNumCuotas & ","
    Msf_Modalidad.Col = 7 'Monto Cuota
    vlSql = vlSql & Str(Msf_Modalidad.Text) & ","
    vlSql = vlSql & Str(CDbl(Msf_Modalidad.Text) * vlNumCuotas) & ",'"
    Msf_Modalidad.Col = 5 'Moneda
    vlSql = vlSql & Trim(Msf_Modalidad.Text) & "',"
    vlSql = vlSql & "'00', NULL, NULL, " 'Motivo Suspensión, Fecha Suspensión
    vlSql = vlSql & "'" & vgUsuario & "',"
    vlSql = vlSql & "'" & Format(Date, "yyyymmdd") & "',"
    vlSql = vlSql & "'" & Format(Time, "hhmmss") & "',"
    vlSql = vlSql & iNumReliq & ")"
    vlSql = vlSQLBaseHab & vlSql
    vgConexionTransac.Execute (vlSql)
    
    'Si la última cuota tiene valor, se debe insertar un nuevo haber/descuento
    If vlUltCuota Then
        Msf_Modalidad.Row = i
        Msf_Modalidad.Col = 0 'Num Orden
        vlSql = Msf_Modalidad.Text & ","
        Msf_Modalidad.Col = 1 'Código Concepto
        vlSql = vlSql & "'" & Trim(Mid(Msf_Modalidad.Text, 1, InStr(1, Msf_Modalidad.Text, "-") - 1)) & "',"
        Msf_Modalidad.Col = 2 'Fecha Inicio
        vlFecInicio = DateAdd("d", 1, vlFecTermino)
        vlSql = vlSql & "'" & Format(vlFecInicio, "YYYYMMDD") & "',"
        Msf_Modalidad.Col = 4 'Fecha Termino
        vlFecTermino = Msf_Modalidad 'DateAdd("d", -1, DateAdd("m", 1, vlFecInicio))
        vlSql = vlSql & "'" & Format(vlFecTermino, "YYYYMMDD") & "',"
        Msf_Modalidad.Col = 3 'Num Cuotas
        vlSql = vlSql & "1,"
        Msf_Modalidad.Col = 8 'Monto Cuota
        vlSql = vlSql & Str(Msf_Modalidad.Text) & ","
        Msf_Modalidad.Col = 8 'Monto Cuotas
        vlSql = vlSql & Str(Msf_Modalidad.Text) & ",'"
        Msf_Modalidad.Col = 5 'Moneda
        vlSql = vlSql & Trim(Msf_Modalidad.Text) & "',"
        vlSql = vlSql & "'00', NULL, NULL, " 'Motivo Suspensión, Fecha Suspensión
        vlSql = vlSql & "'" & vgUsuario & "',"
        vlSql = vlSql & "'" & Format(Date, "yyyymmdd") & "',"
        vlSql = vlSql & "'" & Format(Time, "hhmmss") & "',"
        vlSql = vlSql & iNumReliq & ")"
        vlSql = vlSQLBaseHab & vlSql
        vgConexionTransac.Execute (vlSql)
    End If
    
Next i

'Actualiza Tabla por la cual se hizo la Reliquidación
'con la query traspasada desde el Formulario que Llama a la Reliquidación

vlSql = vpSQLUpdate
vlSql = vlSql & " SET num_reliq = " & iNumReliq
vlSql = vlSql & vpSQLWhere
vgConexionTransac.Execute (vlSql)

flGrabaReliquidacion = True
MsgBox "Reliquidación Grabada Exitosamente", vbInformation

Errores:
    If Err.Number <> 0 Then
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End If
End Function

Function flObtieneDatosReliquidacion(iNumReliq As Double, oMinFecha As Date)
    Dim vlSql As String
    Dim vlTB As ADODB.Recordset
    Dim vlLinea As String
    Dim vlI As Integer
    
    'Obtiene los Datos de una Reliquidacion
    vlSql = "SELECT rel.* FROM pp_tmae_reliq rel"
    vlSql = vlSql & " WHERE rel.num_reliq = " & iNumReliq
    Set vlTB = vgConexionBD.Execute(vlSql)
    If Not vlTB.EOF Then
        txt_FecReliq = DateSerial(Mid(vlTB!FEC_RELIQ, 1, 4), Mid(vlTB!FEC_RELIQ, 5, 2), Mid(vlTB!FEC_RELIQ, 7, 2))
        txt_PerDesde = Mid(vlTB!NUM_PERDESDE, 5, 2) & "/" & Mid(vlTB!NUM_PERDESDE, 1, 4)
        txt_PerHasta = Mid(vlTB!NUM_PERHASTA, 5, 2) & "/" & Mid(vlTB!NUM_PERHASTA, 1, 4)
        txt_Comentarios = IIf(IsNull(vlTB!GLS_OBSERVACION), "", vlTB!GLS_OBSERVACION)
    End If
    
    'Obtiene Datos de los Beneficiarios Reliquidados
    vlSql = "SELECT rel.*, ben.num_idenben, ben.gls_nomben,ben.gls_nomsegben,"
    vlSql = vlSql & " ben.gls_patben, ben.gls_matben, ben.cod_par as par, iden.gls_tipoidencor"
    vlSql = vlSql & " FROM pp_tmae_ben ben, pp_tmae_benreliq rel, ma_tpar_tipoiden iden"
    vlSql = vlSql & " WHERE rel.num_poliza = ben.num_poliza"
    vlSql = vlSql & " AND rel.num_endoso = ben.num_endoso"
    vlSql = vlSql & " AND rel.num_orden = ben.num_orden"
    vlSql = vlSql & " AND iden.cod_tipoiden = ben.cod_tipoidenben"
    vlSql = vlSql & " AND rel.num_reliq = " & iNumReliq
    Set vlTB = vgConexionBD.Execute(vlSql)
    If Not vlTB.EOF Then
        Msf_Beneficiarios.Rows = 1
        Msf_Beneficiarios.Rows = 2
        Msf_Beneficiarios.Row = 1
        vlI = 0
        Do While Not vlTB.EOF
            vlLinea = ""
            vlLinea = vlLinea & IIf(vlTB!COD_INDRELIQ = "1", "X", "") & vbTab
            vlLinea = vlLinea & IIf(vlTB!COD_INDPENSION = "1", "X", "") & vbTab
            vlLinea = vlLinea & vlTB!Num_Orden & vbTab
            vlLinea = vlLinea & vlTB!gls_tipoidencor & vbTab
            vlLinea = vlLinea & vlTB!Num_IdenBen & vbTab
            vlLinea = vlLinea & Trim(vlTB!Gls_NomBen) & " " & Trim(IIf(IsNull(vlTB!Gls_NomSegBen), "", vlTB!Gls_NomSegBen)) & " " & Trim(vlTB!Gls_PatBen) & " " & Trim(IIf(IsNull(vlTB!Gls_MatBen), "", vlTB!Gls_MatBen)) & vbTab
            vlLinea = vlLinea & vlTB!PAR & vbTab
            vlLinea = vlLinea & Mid(vlTB!NUM_PERDESDE, 5, 2) & "/" & Mid(vlTB!NUM_PERDESDE, 1, 4) & vbTab
            vlLinea = vlLinea & Mid(vlTB!NUM_PERHASTA, 5, 2) & "/" & Mid(vlTB!NUM_PERHASTA, 1, 4) & vbTab
            vlI = vlI + 1
            Msf_Beneficiarios.AddItem vlLinea, vlI
            vlTB.MoveNext
        Loop
    End If
    
    
    vlSql = "SELECT rel.*, con.gls_conhabdes FROM"
    vlSql = vlSql & " pp_tmae_detcalcreliq rel, pp_tmae_detpagoreliq pag, "
    vlSql = vlSql & " ma_tpar_conhabdes con"
    vlSql = vlSql & " WHERE con.cod_conhabdes = pag.cod_conhabdes"
    vlSql = vlSql & " AND pag.num_reliq = rel.num_reliq"
    vlSql = vlSql & " AND pag.cod_conhabdes = rel.cod_conhabdes"
    vlSql = vlSql & " AND rel.num_reliq = " & iNumReliq
    Set vlTB = vgConexionBD.Execute(vlSql)
    If Not vlTB.EOF Then
        Msf_Pagos.Rows = 1
        Msf_Pagos.Rows = 2
        Msf_Pagos.Row = 1
        vlI = 0
        Do While Not vlTB.EOF
            vlLinea = ""
            vlLinea = vlLinea & vlTB!Num_Orden & vbTab
            vlLinea = vlLinea & Mid(vlTB!Num_PerPago, 5, 2) & "/" & Mid(vlTB!Num_PerPago, 1, 4) & vbTab
            vlLinea = vlLinea & " " & vlTB!Cod_ConHabDes & " - " & vlTB!gls_ConHabDes & vbTab
            vlLinea = vlLinea & DateSerial(Mid(vlTB!Fec_IniPago & vbTab, 1, 4), Mid(vlTB!Fec_IniPago & vbTab, 5, 2), Mid(vlTB!Fec_IniPago & vbTab, 7, 2)) & vbTab
            vlLinea = vlLinea & DateSerial(Mid(vlTB!Fec_TerPago & vbTab, 1, 4), Mid(vlTB!Fec_TerPago & vbTab, 5, 2), Mid(vlTB!Fec_TerPago & vbTab, 7, 2)) & vbTab
            vlLinea = vlLinea & Format(vlTB!MTO_CONHABDESANT, "##0.00") & vbTab
            vlLinea = vlLinea & Format(vlTB!Mto_ConHabDes, "##0.00") & vbTab
            vlLinea = vlLinea & Format(vlTB!MTO_DIFERENCIA, "##0.00") & vbTab
            vlLinea = vlLinea & vlTB!Cod_Moneda & vbTab
            'vlLinea = vlLinea & vlTB!NUM_ORDENREC & vbTab

            vlI = vlI + 1
            Msf_Pagos.AddItem vlLinea, vlI
            vlTB.MoveNext
        Loop
    End If
    
    vlSql = "SELECT rel.*, con.gls_conhabdes"
    vlSql = vlSql & " FROM pp_tmae_detpagoreliq rel, ma_tpar_conhabdes con"
    vlSql = vlSql & " WHERE con.cod_conhabdes = rel.cod_conhabdes"
    vlSql = vlSql & " AND rel.num_reliq = " & iNumReliq
    Set vlTB = vgConexionBD.Execute(vlSql)
    If Not vlTB.EOF Then
        Msf_Modalidad.Rows = 1
        Msf_Modalidad.Rows = 2
        Msf_Modalidad.Row = 1
        vlI = 0
        oMinFecha = DateSerial(Mid(vlTB!Fec_IniHabDes, 1, 4), Mid(vlTB!Fec_IniHabDes, 5, 2), Mid(vlTB!Fec_IniHabDes, 7, 2))
        Do While Not vlTB.EOF
            vlLinea = ""
            vlLinea = vlLinea & vlTB!Num_Orden & vbTab
            vlLinea = vlLinea & " " & vlTB!Cod_ConHabDes & " - " & vlTB!gls_ConHabDes & vbTab
            vlLinea = vlLinea & DateSerial(Mid(vlTB!Fec_IniHabDes & vbTab, 1, 4), Mid(vlTB!Fec_IniHabDes & vbTab, 5, 2), Mid(vlTB!Fec_IniHabDes & vbTab, 7, 2)) & vbTab
            vlLinea = vlLinea & vlTB!Num_Cuotas & vbTab
            vlLinea = vlLinea & DateSerial(Mid(vlTB!FEC_TERHabDes & vbTab, 1, 4), Mid(vlTB!FEC_TERHabDes & vbTab, 5, 2), Mid(vlTB!FEC_TERHabDes & vbTab, 7, 2)) & vbTab
            vlLinea = vlLinea & vlTB!Cod_Moneda & vbTab
            vlLinea = vlLinea & Format(vlTB!mto_total, "##0.00") & vbTab
            vlLinea = vlLinea & Format(vlTB!MTO_CUOTA, "##0.00") & vbTab
            vlLinea = vlLinea & Format(vlTB!MTO_ULTCUOTA, "##0.00") & vbTab
            vlI = vlI + 1
            Msf_Modalidad.AddItem vlLinea, vlI
            If oMinFecha > vlTB!Fec_IniHabDes Then
                oMinFecha = vlTB!Fec_IniHabDes
            End If
            vlTB.MoveNext
        Loop
    End If
    
End Function

Function flObtieneNumReliq() As Double
Dim vlSql As String
Dim vlTB As ADODB.Recordset
vlSql = "SELECT MAX(num_reliq) as reliq FROM PP_TMAE_RELIQ"
Set vlTB = vgConexionBD.Execute(vlSql)
If Not vlTB.EOF Then
    flObtieneNumReliq = IIf(IsNull(vlTB!reliq), 0, vlTB!reliq) + 1
Else
    flObtieneNumReliq = 1
End If
End Function


Private Sub chk_Reliq_Click()
    If chk_Reliq.Value = "1" Then
        chk_Pension.Enabled = True
    Else
        chk_Pension.Value = "0"
        chk_Pension.Enabled = False
    End If
End Sub

Private Sub cmd_BenSumar_Click()
On Error GoTo Err_Sumar
Dim vlEncontrado As Boolean
Dim i As Integer
    'Validar Campos
    If chk_Reliq.Value = 1 Then
        'Validar que seleccione por lo menos una de las opciones
        If chk_Pension.Value = 0 Then
            MsgBox "Debe seleccionar opcion de Reliquidacion de Pensión", vbCritical, Me.Caption
            chk_Pension.SetFocus
            Exit Sub
        End If
    End If
    
    For i = 1 To Msf_Beneficiarios.Rows - 2
        Msf_Beneficiarios.Row = i
        Msf_Beneficiarios.Col = 2
        If Msf_Beneficiarios = txt_BenNumOrden Then
            Msf_Beneficiarios.Col = 0
            If chk_Reliq = "1" Then
                Msf_Beneficiarios = "X"
            Else
                Msf_Beneficiarios = ""
            End If
            
            Msf_Beneficiarios.Col = 1
            If chk_Pension = "1" Then
                Msf_Beneficiarios = "X"
            Else
                Msf_Beneficiarios = ""
            End If
            
            Msf_Beneficiarios.Col = 7
            Msf_Beneficiarios = txt_BenPerDesde
            
            Msf_Beneficiarios.Col = 8
            Msf_Beneficiarios = txt_BenPerHasta
            
            'Si se actualizan los datos de los beneficiarios se debe volver a calcular
            cmdCalcular.Enabled = True
            flLimpiaPagos
            flLimpiaModalidad
            Exit For
        End If
    Next i

Exit Sub
Err_Sumar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Buscar_Click()
On Error GoTo Err_Buscar

    Frm_Busqueda.flInicio ("Frm_AFReliquidacion")

Exit Sub
Err_Buscar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub

Public Sub Cmd_BuscarPol_Click()
Dim vlMinFecha As Date  'Menor fecha de Inicio de Pago
On Error GoTo Err_Buscar

   If Trim(Txt_PenPoliza) <> "" Or Trim(Txt_PenNumIdent) <> "" Then
       If Trim(Txt_PenNumIdent) <> "" Then
          If Trim(Cmb_PenNumIdent) = "" Then
             MsgBox "Debe seleccionar el Tipo de Identificación.", vbCritical, "Error de Datos"
             Cmb_PenNumIdent.SetFocus
             Exit Sub
          End If
       End If
       'Permite Buscar los Datos del Beneficiario
       flValidarBen
       If vpAccion = "N" Then
            'Obtiene Beneficiarios a Reliquidar
            If Not flObtieneBeneficiarios(vpIndPension, vpNumOrden) Then
                MsgBox "No se encontraron Beneficiarios a Reliquidar", vbCritical, Me.Caption
                Unload Me
                Exit Sub
            End If
       Else
            'Obtiene Datos de la Reliquidación
            Call flObtieneDatosReliquidacion(txt_NumReliq, vlMinFecha)
            'construir Verifica si ya se están realizando pagos de la reliquidación para deshabilitar lo que no se puede modificar
            If vpAccion = "C" Then 'Consultar
                'Deja todo deshabilitado
                Fra_Antecedentes.Enabled = False
                Frm_Beneficiarios.Enabled = False
                Frm_Pagos.Enabled = False
                Frm_Modalidad.Enabled = False
                Cmd_Grabar.Enabled = False
                Cmd_Eliminar.Enabled = False
            Else
                'Modificar
                If vlMinFecha > vpEfecto Then
                    Cmd_Eliminar.Enabled = False
                End If
            End If
       End If
   Else
     MsgBox "Debe ingresar el Nº de Póliza o Número de Identificación del Pensionado", vbCritical, "Error de Datos"
     Txt_PenPoliza.SetFocus
   End If

Exit Sub
Err_Buscar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Eliminar_Click()
    Dim vlNumReliq As Double
        
    'Validar que Se puede eliminar si no se ha empezado a pagar alguno de los beneficios
    If txt_NumReliq = "" Then
        MsgBox "No existe Reliquidación a Eliminar", vbCritical
        Exit Sub
    End If
            
    If MsgBox("¿ Está Seguro que desea Eliminar la Reliquidación ?", 32 + 4, Me.Caption) <> 6 Then
        Exit Sub
    End If
    
    vlNumReliq = txt_NumReliq
    'Abre conexión a la BD
    If Not fgConexionBaseDatos(vgConexionTransac) Then
        MsgBox "Error en Conexion a la Base de Datos", vbCritical, Me.Caption
        Exit Sub
    End If
    vgConexionTransac.BeginTrans
    'Elimina Reliquidación Anterior
    If Not flEliminaReliquidacion(vlNumReliq) Then
        GoTo ErrEliminar
    End If
    vgConexionTransac.CommitTrans
    MsgBox "Reliquidación Eliminada Exitosamente", vbInformation
    
    Unload Me
    Exit Sub

ErrEliminar:
    vgConexionTransac.RollbackTrans
    MsgBox "Se han producido Errores al Eliminar la Reliquidacion." & Chr(13) & Err.Description, vbCritical
End Sub

Private Sub cmd_grabar_Click()
    Dim vlNumReliq As Double
    Dim vlFecha As Date
    
    'Obtiene Máximo Número de Reliquidación para generar el próximo
    If Not flValidaDatos Then
        Exit Sub
    End If
        
    If vpAccion = "N" Then 'Se trata de una nueva Reliquidación
        vlNumReliq = flObtieneNumReliq
        'Abre conexión a la BD
        If Not fgConexionBaseDatos(vgConexionTransac) Then
            MsgBox "Error en Conexion a la Base de Datos", vbCritical, Me.Caption
            Exit Sub
        End If
        vgConexionTransac.BeginTrans
   Else
        vlNumReliq = txt_NumReliq
        'Abre conexión a la BD
        If Not fgConexionBaseDatos(vgConexionTransac) Then
            MsgBox "Error en Conexion a la Base de Datos", vbCritical, Me.Caption
            Exit Sub
        End If
        vgConexionTransac.BeginTrans
        'Elimina Reliquidación Anterior
        If Not flEliminaReliquidacion(vlNumReliq) Then
            GoTo ErrGrabar
        End If
   End If
   
   'Graba Reliquidación
   If Not flGrabaReliquidacion(vlNumReliq) Then
        GoTo ErrGrabar
   End If
   vgConexionTransac.CommitTrans
   
   vlFecha = fgBuscaFecServ()
   txt_NumReliq = vlNumReliq
   txt_FecReliq = vlFecha
   'Unload Me
   vpAccion = "M"
   Exit Sub
   
'Error
ErrGrabar:
    vgConexionTransac.RollbackTrans
    MsgBox "Se han producido Errores al Grabar la Reliquidacion." & Chr(13) & Err.Description, vbCritical
End Sub

Private Function flValidaDatos() As Boolean
    Dim vlReliq As String, vlAF As String
    Dim vlGE As String, vlPension As String
    Dim vlI As Long
    flValidaDatos = False
    
    'Valida Datos a Grabar
    If txt_PerDesde = "" Then
        MsgBox "Debe ingresar Periodo Desde", vbCritical
        txt_PerDesde.SetFocus
        Exit Function
    End If
    
    If txt_PerHasta = "" Then
        MsgBox "Debe ingresar Periodo Hasta", vbCritical
        txt_PerHasta.SetFocus
        Exit Function
    End If
    
    'Valida Grilla de Beneficiarios
    If Msf_Beneficiarios.Rows < 3 Then
        MsgBox "No puede ingresar Reliquidación sin Beneficiarios", vbCritical
        Exit Function
    End If
    For vlI = 1 To Msf_Beneficiarios.Rows - 2
        Msf_Beneficiarios.Row = vlI
        Msf_Beneficiarios.Col = 0
        vlReliq = IIf(Msf_Beneficiarios = "X", 1, 0)
        Msf_Beneficiarios.Col = 1
        vlPension = IIf(Msf_Beneficiarios = "X", 1, 0)
        If vlReliq = 1 Then
            'Validar que seleccione por lo menos una de las opciones
            If vlPension = 0 Then
                MsgBox "En la Fila [" & vlI & "] de Beneficiarios debe seleccionar por lo menos una de las opciones de Reliquidacion", vbCritical, Me.Caption
                SSTab1.Tab = 0
                Exit Function
            End If
        End If
    Next vlI
    
    'Validar Grilla de Pagos Reliquidados por Beneficiario
    If Msf_Pagos.Rows < 3 Then
        MsgBox "Debe efectuar Cálculo de Reliquidación", vbCritical
        If cmdCalcular.Enabled Then
            cmdCalcular.SetFocus
        Else
            SSTab1.Tab = 1
        End If
        Exit Function
    End If
    For vlI = 1 To Msf_Pagos.Rows - 2
        Msf_Pagos.Row = vlI
        Msf_Pagos.Col = 6
        If Not IsNumeric(Msf_Pagos) Then
            SSTab1.Tab = 1
            MsgBox "Monto Actual incorrecto en la Fila [" & vlI & "] de los Pagos Reliquidados por Beneficiario", vbCritical
            Exit Function
        End If
    Next vlI
    
    'Valida Grilla de Modalidad de Pago
    If Msf_Modalidad.Rows < 3 Then
        MsgBox "Debe efectuar Calculo de Reliquidación", vbCritical
        If cmdCalcular.Enabled Then
            cmdCalcular.SetFocus
        Else
            SSTab1.Tab = 2
        End If
        Exit Function
    End If
    For vlI = 1 To Msf_Modalidad.Rows - 2
        Msf_Modalidad.Row = vlI
        Msf_Modalidad.Col = 2
        If Not IsDate(Msf_Modalidad) Then
            SSTab1.Tab = 2
            MsgBox "Debe ingresar una Fecha de Inicio del Pago Válida en la Fila [" & vlI & " ] de la Modalidad de Pago", vbCritical
            Exit Function
        End If
        Msf_Modalidad.Col = 3
        If Not IsNumeric(Msf_Modalidad) Then
            SSTab1.Tab = 2
            MsgBox "Debe ingresar un Número de Cuotas Válido en la Fila [" & vlI & " ] de la Modalidad de Pago", vbCritical
            Exit Function
        End If
        
        If Msf_Modalidad <= 0 Then
            SSTab1.Tab = 2
            MsgBox "Número de Cuotas no válido en la Fila [" & vlI & " ] de la Modalidad de Pago", vbCritical
            Exit Function
        End If
    Next vlI
    
    flValidaDatos = True
    
End Function

Private Sub Cmd_Imprimir_Click()
On Error GoTo Err_Imprimir

'    If Fra_Poliza.Enabled = True Then
'       Exit Sub
'    End If
    
    'Imprime el Reporte de Variables
    flImpresion

Exit Sub
Err_Imprimir:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Sub flImpresion()
Dim vlArchivo As String
Dim vlFecNac As String

Err.Clear
On Error GoTo Errores1
   
   Screen.MousePointer = 11
   
   vlArchivo = strRpt & "PP_Rpt_Reliquidacion.rpt"   '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
       MsgBox "Archivo de Reporte de Listados no se encuentra en el directorio de la Aplicación.", 16, "Archivo no encontrado"
       Screen.MousePointer = 0
       Exit Sub
   End If
   
   If Txt_PenPoliza = "" Then
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   'I GGO 22/03/2005
   If txt_NumReliq = "" Then
      MsgBox "Debe grabar la Reliquidación antes de Imprimir", vbCritical, Me.Caption
      Screen.MousePointer = 0
      Exit Sub
   End If
   'F GGO
   
   vgQuery = ""
   vgQuery = "{PP_TMAE_RELIQ.NUM_RELIQ} = " & (txt_NumReliq)
   Rpt_Reliquidacion.Reset
   Rpt_Reliquidacion.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
   Rpt_Reliquidacion.Connect = vgRutaDataBase
   Rpt_Reliquidacion.SelectionFormula = vgQuery
   Rpt_Reliquidacion.Formulas(0) = ""
   Rpt_Reliquidacion.Formulas(1) = ""
   Rpt_Reliquidacion.Formulas(2) = ""
   
   Rpt_Reliquidacion.Formulas(3) = ""
   Rpt_Reliquidacion.Formulas(4) = ""
   Rpt_Reliquidacion.Formulas(5) = ""
   Rpt_Reliquidacion.Formulas(6) = ""
   Rpt_Reliquidacion.Formulas(7) = ""
   
   Rpt_Reliquidacion.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
   Rpt_Reliquidacion.Formulas(1) = "NombreSistema= '" & vgNombreSistema & "'"
   Rpt_Reliquidacion.Formulas(2) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
   
   Rpt_Reliquidacion.SubreportToChange = "Beneficiarios"
   Rpt_Reliquidacion.Connect = vgRutaDataBase
   Rpt_Reliquidacion.SubreportToChange = "PagosRel"
   Rpt_Reliquidacion.Connect = vgRutaDataBase
   Rpt_Reliquidacion.SubreportToChange = "Modalidad"
   Rpt_Reliquidacion.Connect = vgRutaDataBase
'   vlRut = Trim(Txt_PenRut) + "-" + Trim(Txt_PenDigito)
'
'   Rpt_Reliquidacion.Formulas(3) = "Poliza = '" & Trim(Txt_PenPoliza) & "'"
'   Rpt_Reliquidacion.Formulas(4) = "Endoso = '" & Trim(Lbl_Endoso) & "'"
'   Rpt_Reliquidacion.Formulas(5) = "Rut = '" & Trim(vlRut) & "'"
'   Rpt_Reliquidacion.Formulas(6) = "Nombre_bene = '" & Trim(Lbl_PenNombre) & "'"
'   Rpt_Reliquidacion.Formulas(7) = "Fec_Nac = '" & (vlFechaNac) & "'"
      
   Rpt_Reliquidacion.Destination = crptToWindow
   Rpt_Reliquidacion.WindowState = crptMaximized
   Rpt_Reliquidacion.WindowTitle = "Informe de Reliquidación"
   Rpt_Reliquidacion.Action = 1
   
   Screen.MousePointer = 0
   
Exit Sub
Errores1:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " " & "Numero de Error : " & Err.Number & Err.HelpFile & Err.HelpContext, vbExclamation, "Mensaje de error ..."
        Screen.MousePointer = 0
    End If
End Sub


Private Sub Cmd_Limpiar_Click()
    Fra_Poliza.Enabled = True
    Fra_Antecedentes.Enabled = False
    SSTab1.Enabled = False
    
    Txt_PenPoliza = ""
    Txt_PenNumIdent = ""
    Cmb_PenNumIdent.ListIndex = -1
    Lbl_PenNombre = ""
    txt_NumReliq = ""
    txt_FecReliq = ""
    txt_PerDesde = ""
    txt_PerHasta = ""
    txt_Comentarios = ""
    
    'Carpeta Beneficiarios Afectados
    chk_Reliq = 0
    chk_Pension = 0
    txt_BenNumOrden = ""
    txt_BenNroIden = ""
    txt_BenTipoIden = ""
    txt_BenNom = ""
    txt_BenPar = ""
    txt_BenPerDesde = ""
    txt_BenPerHasta = ""
    Msf_Beneficiarios.Rows = 1
    Msf_Beneficiarios.Rows = 2
    
    flLimpiaPagos
    
    flLimpiaModalidad
    
End Sub

Private Sub Cmd_ModSumar_Click()
Dim vlNumOrden As Long
Dim vlConcepto As String
Dim i As Integer

'Valida que los datos ingresados sean válidos
If Not IsDate(txt_ModFecInicio) Then
    MsgBox "Debe ingresar una Fecha de Inicio del Pago Válida", vbCritical
    txt_ModFecInicio.SetFocus
    Exit Sub
End If

If Not IsNumeric(txt_ModNumCuotas) Then
    MsgBox "Debe ingresar un Número de Cuotas Válido", vbCritical
    txt_ModNumCuotas.SetFocus
    Exit Sub
End If

txt_ModNumCuotas = CLng(txt_ModNumCuotas)
If txt_ModNumCuotas <= 0 Then
    MsgBox "Número de Cuotas no válido", vbCritical
    Exit Sub
End If

'If txt_ModNumCuotas > 120 Then
'    MsgBox "Número de Cuotas no válido", vbCritical 'No puede ser mayor a 10 años
'    Exit Sub
'End If

'Obtiene Número de Cuotas y Monto de las Cuotas
For i = 1 To Msf_Modalidad.Rows - 2
    Msf_Modalidad.Row = i
    
    Msf_Modalidad.Col = 0
    vlNumOrden = Msf_Modalidad
    
    Msf_Modalidad.Col = 1
    vlConcepto = Msf_Modalidad
    
    If vlNumOrden = txt_ModNumOrden And vlConcepto = txt_ModConcepto Then
        Msf_Modalidad.Col = 2
        Msf_Modalidad = txt_ModFecInicio
        
        Msf_Modalidad.Col = 3
        Msf_Modalidad = txt_ModNumCuotas
        
        Msf_Modalidad.Col = 4
        Msf_Modalidad = txt_ModFecFin
        
        Msf_Modalidad.Col = 7
        Msf_Modalidad = txt_ModCuota
        
        Msf_Modalidad.Col = 8
        Msf_Modalidad = txt_ModUltCuota
        
    End If
Next i
End Sub

Private Sub Cmd_PagRestar_Click()
    Dim vlMontoAnterior As Double
    Dim vlNumOrden As Long
    Dim vlConcepto As String
    Dim i As Integer, j As Integer, k As Integer
    
    If Msf_Pagos.Row <> 0 And Msf_Pagos.Row <> Msf_Pagos.Rows - 1 Then
    
        i = Msf_Pagos.Row
        
        Msf_Pagos.Col = 0
        vlNumOrden = Msf_Pagos
        
        Msf_Pagos.Col = 2
        vlConcepto = Msf_Pagos
        
        Msf_Pagos.Col = 5
        vlMontoAnterior = Msf_Pagos
            
        'Actualiza Monto
        For j = 1 To Msf_Modalidad.Rows - 2
            Msf_Modalidad.Row = j
            Msf_Modalidad.Col = 0
            vlNumOrden = Msf_Modalidad
            
            Msf_Modalidad.Col = 1
            vlConcepto = Msf_Modalidad
            If vlArrayPagos(i).Num_Orden = vlNumOrden And vlArrayPagos(i).Concepto = vlConcepto Then
                Msf_Modalidad.Col = 6
                vlArrayPagos(i).monto = 0
                Msf_Modalidad = Msf_Modalidad - vlMontoAnterior
                
                If CDbl(Msf_Modalidad) = 0 Then
                    Msf_Modalidad.RemoveItem i
                    For k = i To Msf_Pagos.Rows - 3
                        vlArrayPagos(k).Concepto = vlArrayPagos(k + 1).Concepto
                        vlArrayPagos(k).Ind_Acumulado = vlArrayPagos(i + 1).Ind_Acumulado
                        vlArrayPagos(k).Moneda = vlArrayPagos(k + 1).Moneda
                        vlArrayPagos(k).monto = vlArrayPagos(k + 1).monto
                        vlArrayPagos(k).Num_Orden = vlArrayPagos(k + 1).Num_Orden
                    Next k
                    ReDim Preserve vlArrayPagos(k)
                Else
                    'Limpia Forma de Pago ingresada anteriormente
                    Msf_Modalidad.Col = 2
                    Msf_Modalidad = ""
                    Msf_Modalidad.Col = 3
                    Msf_Modalidad = ""
                    Msf_Modalidad.Col = 7
                    Msf_Modalidad = ""
                    Msf_Modalidad.Col = 8
                    Msf_Modalidad = ""
                End If
                Exit For
            End If
        Next j
        Msf_Pagos.RemoveItem i
        'iRes = ADEL()(vlArrayPagos, Msf_Pagos.Row)
    End If
End Sub

Private Sub Cmd_PagSumar_Click()
On Error GoTo Err_Sumar
Dim vlNumOrden As Long, vlPeriodo As String, vlConcepto As String
Dim vlMontoAnt As Double
Dim i As Integer, j As Integer
    'Validar Campos

    For i = 1 To Msf_Pagos.Rows - 2
        Msf_Pagos.Row = i
        Msf_Pagos.Col = 0
        vlNumOrden = Msf_Pagos
        
        Msf_Pagos.Col = 1
        vlPeriodo = Msf_Pagos
        
        Msf_Pagos.Col = 2
        vlConcepto = Msf_Pagos
        If vlNumOrden = txt_PagNumOrden And vlPeriodo = txt_PagPeriodo And vlConceptoModif = txt_PagConcepto Then
            
            Msf_Pagos.Col = 7
            vlMontoAnt = Msf_Pagos
            Msf_Pagos = txt_PagMonto
            vlArrayPagos(i).monto = txt_PagMonto
            Msf_Pagos.Col = 6
            Msf_Pagos = txt_PagMontoAct
            'Actualiza Monto
            For j = 1 To Msf_Modalidad.Rows - 2
                Msf_Modalidad.Row = j
                Msf_Modalidad.Col = 0
                vlNumOrden = Msf_Modalidad
                
                Msf_Modalidad.Col = 1
                vlConcepto = Msf_Modalidad
                If vlArrayPagos(i).Num_Orden = vlNumOrden And vlArrayPagos(i).Concepto = vlConcepto Then
                    Msf_Modalidad.Col = 6
                    Msf_Modalidad = Msf_Modalidad + vlArrayPagos(i).monto - vlMontoAnt
                    
                    'Limpia Forma de Pago ingresada anteriormente
'                    Msf_Modalidad.Col = 2
'                    Msf_Modalidad = ""
                    Msf_Modalidad.Col = 3
                    Msf_Modalidad = ""
                    Msf_Modalidad.Col = 7
                    Msf_Modalidad = ""
                    Msf_Modalidad.Col = 8
                    Msf_Modalidad = ""
                End If
            Next j
            Exit For
        End If
    Next i

Exit Sub
Err_Sumar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub cmd_salir_Click()
    Unload Me
End Sub

Private Sub cmdCalcular_Click()
   'Función que realiza el Cálculo del Monto por Reliquidaciones
   Dim vlPerDesde As Date, vlPerHasta As Date
   Dim vlOrden As Long, vlTotAsigFam As Double 'Monto Total de Asignación Familiar
   Dim vlFechaHasta As Date, vlConcepto As String, vlMoneda As String
   Dim vlFila As Long
   Dim vlMontoAnterior As Double, vlDiferencia As Double
   Dim i As Integer
   Dim vlVarCarga As Double, vlNumCargas As Integer
   Dim vlAsigFamiliar As Double
   Dim vlCalcPension As Boolean
   Dim vlTipoPension As String
   Dim vlIndPension As Boolean, vlIndGarantia As Boolean  'Indica si se debe Pagar Pensión o Garantía Estatal
   Dim vlFecTerPagoPenGar As String, vlPensionGar As Double
   Dim vlPension As Double
   'Validación de Datos
   
   'Limpia Grillas
   Msf_Pagos.Rows = 1
   Msf_Pagos.Rows = 2
   Msf_Modalidad.Rows = 1
   Msf_Modalidad.Rows = 2
   
   'Conceptos de Pensión
   vlConceptoPEMenor = fgObtieneDescripcionConcepto(stDatGenerales.Cod_ConceptoPensionCobro)
   vlConceptoPEMayor = fgObtieneDescripcionConcepto(stDatGenerales.Cod_ConceptoPensionPago)
   
   'Cálculo de Opciones Seleccionadas
   vlTotAsigFam = 0
   vlFila = 1
    For i = 1 To Msf_Beneficiarios.Rows - 2
        Msf_Beneficiarios.Row = i
        Msf_Beneficiarios.Col = 0
        vlCalcPension = False
        If Msf_Beneficiarios = "X" Then 'Está para reliquidación
            Msf_Beneficiarios.Col = 7
            vlPerDesde = DateSerial(Mid(Msf_Beneficiarios, 4, 4), Mid(Msf_Beneficiarios, 1, 2), 1)
            
            Msf_Beneficiarios.Col = 8
            vlPerHasta = DateSerial(Mid(Msf_Beneficiarios, 4, 4), Mid(Msf_Beneficiarios, 1, 2), 1)
            Msf_Beneficiarios.Col = 2
            vlOrden = Msf_Beneficiarios
            Msf_Beneficiarios.Col = 1
               If Msf_Beneficiarios = "X" Then
                    vlIndPension = True
               Else
                    vlIndPension = False
               End If
               vlFecTerPagoPenGar = ""
               vlPensionGar = 0
               vlPension = 0
               vlCalcPension = True 'Se realizó cálculo de la Pensión
               If vlIndPension Then
                    Dim vlPensionNormal As Double
                    Dim vlSql As String, vlTB As ADODB.Recordset
                    Dim vlCodPar As String, vlFecHasta As Date
                    Dim vlEdad As Integer, vlEdadAños As Integer, bResp As Integer
                    Dim vlFecIniPago As String, vlTB2 As ADODB.Recordset
                    
                    vlSql = "SELECT pol.cod_tippension, pol.mto_pension, "
                    vlSql = vlSql & " ben.fec_terpagopengar, ben.cod_par, ben.fec_nacben, "
                    vlSql = vlSql & " ben.fec_inipagopen, ben.cod_sitinv, ben.cod_sexo,"
                    vlSql = vlSql & " ben.prc_pension, ben.prc_pensiongar, pol.cod_moneda"
                    vlSql = vlSql & ", pol.cod_tipreajuste" 'hqr 03/12/2010
                    vlSql = vlSql & " FROM pp_tmae_ben ben, pp_tmae_poliza pol"
                    vlSql = vlSql & " WHERE pol.num_poliza = ben.num_poliza"
                    vlSql = vlSql & " AND pol.num_endoso = ben.num_endoso"
                    vlSql = vlSql & " AND pol.num_poliza = '" & Txt_PenPoliza & "'"
                    vlSql = vlSql & " AND pol.num_endoso = " & Lbl_Endoso
                    'vlSql = vlSql & " AND ben.num_orden = " & vlNumOrden
                    vlSql = vlSql & " AND ben.num_orden = " & vpNumOrden
                    
                    Set vlTB = vgConexionBD.Execute(vlSql)
                    If Not vlTB.EOF Then
                        vlTipoPension = vlTB!Cod_TipPension
                        vlPensionNormal = vlTB!Mto_Pension
                                                
                        vlFecIniPago = DateSerial(Mid(vlTB!Fec_IniPagoPen, 1, 4), Mid(vlTB!Fec_IniPagoPen, 5, 2), Mid(vlTB!Fec_IniPagoPen, 7, 2))
                        vlCodPar = vlTB!Cod_Par
                        If Not IsNull(vlTB!Fec_TerPagoPenGar) Then
                            vlFecTerPagoPenGar = vlTB!Fec_TerPagoPenGar
                            vlPensionGar = vlTB!Mto_Pension
                        End If
                    
                        Do While vlPerDesde <= vlPerHasta
                        'If vlPerDesde = "01/10/2009" Then MsgBox "yata"
                            'Obtiene Monto de la Pensión
                            If CDate(vlFecIniPago) > vlPerDesde Then
                                GoTo Siguiente 'Aun no le corresponde pension
                            End If
                            'If vlTB!Cod_Moneda = "NS" Then
                            If vlTB!Cod_TipReajuste <> cgSINAJUSTE Then 'hqr 03/12/2010
                                'Obtiene Monto de la Pensión Actualizada
                                vlSql = "SELECT a.mto_pension "
                                vlSql = vlSql & "FROM pp_tmae_pensionact a "
                                vlSql = vlSql & "WHERE a.num_poliza = '" & Txt_PenPoliza & "' "
                                vlSql = vlSql & "AND a.num_endoso = " & Lbl_Endoso & " "
                                vlSql = vlSql & "AND a.fec_desde = ("
                                    vlSql = vlSql & "SELECT max(fec_desde) FROM pp_tmae_pensionact b "
                                    vlSql = vlSql & "WHERE b.num_poliza = a.num_poliza "
                                    vlSql = vlSql & "AND b.num_endoso = a.num_endoso "
                                    vlSql = vlSql & "AND b.fec_desde <= '" & Format(vlPerDesde, "yyyymmdd") & "'"
                                    vlSql = vlSql & ")"
                                Set vlTB2 = vgConexionBD.Execute(vlSql)
                                If Not vlTB2.EOF Then
                                    vlPensionNormal = vlTB2!Mto_Pension
                                End If
                            End If
                            
                            vlPension = Format(vlPensionNormal * vlTB!Prc_Pension / 100, "#0.00")
                            
                            If vlFecTerPagoPenGar <> "" Then 'Hay fecha de Pago Garantizado
                                If vlFecTerPagoPenGar >= Format(vlPerDesde, "yyyymmdd") Then
                                    vlPension = Format(vlPensionNormal * vlTB!Prc_PensionGar / 100, "#0.00")
                                End If
                            End If
                            
                            vlFecHasta = DateAdd("d", -1, DateAdd("m", 1, vlPerDesde))
                            bResp = fgCalculaEdad(vlTB!Fec_NacBen, vlFecHasta)
                            If bResp = "-1" Then 'Error
                                'GoTo Deshacer
                                Exit Sub
                            End If
                            vlEdad = bResp
                            vlEdadAños = fgConvierteEdadAños(vlEdad)
                            
                            'Si son Hijos se Calcula la Edad y se Verifica Certificado de Estudios
                            If vlCodPar >= 30 And vlCodPar <= 35 Then 'Hijos
                                'bResp = fgCalculaEdad(vlTB2!FEC_NACBEN, vgFecIniPag)

                                If vlEdad >= stDatGenerales.MesesEdad18 And vlTB!Cod_SitInv = "N" Then 'Hijos Sanos
                                    'OBS: Se asume que el mes de los 18 años se paga completo
                                        GoTo Siguiente
                                    End If
                            End If

                            If vlIndPension Then
                                'Obtener Pensión Anterior
                                vlMontoAnterior = fgObtieneMontoConcepto(Txt_PenPoliza, vlNumOrden, vlOrden, stDatGenerales.Cod_ConceptoPension, Format(vlPerDesde, "yyyymm"), "PE") 'Estaría en Pesos
                                'vlMontoAnterior = Format(vlMontoAnterior / vlUF, "##0.00") 'Lo Transforma a UF
                                vlDiferencia = Format(vlMontoAnterior - vlPension, "#0.00")
                                If vlDiferencia <> 0 Then
                                    'variables en duro -- se deben cambiar
                                    If vlDiferencia > 0 Then
                                        vlConcepto = vlConceptoPEMenor
                                    Else
                                        vlConcepto = vlConceptoPEMayor
                                    End If
                                    vlDiferencia = Abs(vlDiferencia)
                                    vlMoneda = vlTB!Cod_Moneda
                                    'Hasta Acá
                                    ReDim Preserve vlArrayPagos(vlFila)
                                    Msf_Pagos.AddItem vlOrden & vbTab & Format(vlPerDesde, "mm/yyyy") & vbTab & _
                                    vlConcepto & vbTab & vlPerDesde & vbTab & vlFecHasta & vbTab & vlMontoAnterior & vbTab & vlPension & vbTab & vlDiferencia & vbTab & vlMoneda & vbTab & vlNumOrden, vlFila
                                    vlArrayPagos(vlFila).Ind_Acumulado = 0
                                    vlArrayPagos(vlFila).monto = vlDiferencia 'vlAsigFamiliar
                                    vlArrayPagos(vlFila).Concepto = vlConcepto
                                    vlArrayPagos(vlFila).Moneda = vlMoneda
                                    vlArrayPagos(vlFila).Num_Orden = vlNumOrden 'vlOrden
                                    vlFila = vlFila + 1
                                End If
                            End If
Siguiente:
                            vlPerDesde = DateAdd("m", 1, vlPerDesde) 'Incrementa el Periodo
                        Loop
                    End If
                    vlTB.Close
                End If
            End If
'''        End If
    Next i
    
    flAcumulaPagos
    cmdCalcular.Enabled = False
'    Fra_Antecedentes.Enabled = False
    
End Sub



Private Sub Form_Load()
On Error GoTo Err_Cargar

''    vpIndAF = 1
''    vpIndGE = 0
''    vpIndPension = 0
''    vpNumOrden = -1
''    vpNumOrdenRec = 1

    Frm_AFReliquidacion.Top = 0
    Frm_AFReliquidacion.Left = 0
    SSTab1.Enabled = False
    SSTab1.Tab = 0
    Fra_Antecedentes.Enabled = False
    'Carga Combo de Tipo de Identificación Causante
    fgComboTipoIdentificacion Cmb_PenNumIdent
''    'Valores para pruebas de Cálculo
''    txt_PerDesde = "10/2004"
''    txt_PerHasta = "01/2005"
    vpEstado = "A"
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    vpEstado = "C" 'Cerrado
End Sub

Private Sub Msf_Beneficiarios_DblClick()
    If Msf_Beneficiarios.Row <> 0 And Msf_Beneficiarios.Row <> Msf_Beneficiarios.Rows - 1 Then
        Msf_Beneficiarios.Col = 0
        If Msf_Beneficiarios = "X" Then
            chk_Reliq.Value = "1"
        Else
            chk_Reliq.Value = "0"
        End If
                
        Msf_Beneficiarios.Col = 1
        If Msf_Beneficiarios = "X" Then
            chk_Pension.Value = "1"
        Else
            chk_Pension.Value = "0"
        End If
        
        Msf_Beneficiarios.Col = 2
        txt_BenNumOrden = Msf_Beneficiarios
        
        Msf_Beneficiarios.Col = 3
        txt_BenTipoIden = Msf_Beneficiarios
        
        Msf_Beneficiarios.Col = 4
        txt_BenNroIden = Msf_Beneficiarios

        Msf_Beneficiarios.Col = 5
        txt_BenNom = Msf_Beneficiarios
        
        Msf_Beneficiarios.Col = 6
        txt_BenPar = Msf_Beneficiarios
        
        Msf_Beneficiarios.Col = 7
        txt_BenPerDesde = Msf_Beneficiarios
        
        Msf_Beneficiarios.Col = 8
        txt_BenPerHasta = Msf_Beneficiarios
    End If
End Sub





Private Sub Msf_Modalidad_Click()
    If Msf_Modalidad.Row <> 0 And Msf_Modalidad.Row <> Msf_Modalidad.Rows - 1 Then
        Msf_Modalidad.Col = 0
        txt_ModNumOrden = Msf_Modalidad.Text
        
        Msf_Modalidad.Col = 1
        txt_ModConcepto = Msf_Modalidad
        
        Msf_Modalidad.Col = 2
        txt_ModFecInicio = Msf_Modalidad
        
        Msf_Modalidad.Col = 3
        txt_ModNumCuotas = Msf_Modalidad
        
        Msf_Modalidad.Col = 4
        txt_ModFecFin = Msf_Modalidad
        
        Msf_Modalidad.Col = 5
        txt_ModMoneda = Msf_Modalidad
        
        Msf_Modalidad.Col = 6
        txt_ModTotal = Msf_Modalidad
        
        Msf_Modalidad.Col = 7
        txt_ModCuota = Msf_Modalidad
        
        Msf_Modalidad.Col = 8
        txt_ModUltCuota = Msf_Modalidad
    End If
End Sub

Private Sub Msf_Pagos_DblClick()
    If Msf_Pagos.Row <> 0 And Msf_Pagos.Row <> Msf_Pagos.Rows - 1 Then
        Msf_Pagos.Col = 0
        txt_PagNumOrden = Msf_Pagos.Text
        
        Msf_Pagos.Col = 1
        txt_PagPeriodo = Msf_Pagos
        
        Msf_Pagos.Col = 2
        txt_PagConcepto = Msf_Pagos
        vlConceptoModif = txt_PagConcepto
        
        Msf_Pagos.Col = 3
        txt_PagFecInicio = Msf_Pagos
        
        Msf_Pagos.Col = 4
        txt_PagFecFin = Msf_Pagos
        
        Msf_Pagos.Col = 5
        txt_PagMontoAnt = Msf_Pagos
        
        Msf_Pagos.Col = 6
        txt_PagMontoAct = Msf_Pagos
        
        Msf_Pagos.Col = 7
        txt_PagMonto = Msf_Pagos
        
        Msf_Pagos.Col = 8
        txt_PagMoneda = Msf_Pagos
    End If
End Sub


Private Sub txt_ModFecInicio_GotFocus()
    txt_ModFecInicio.SelStart = 0
    txt_ModFecInicio.SelLength = Len(txt_ModFecInicio)
End Sub

Private Sub txt_ModFecInicio_KeyPress(KeyAscii As Integer)
    'Solo acepta números y los separadores de fecha
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 47 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub


Private Sub txt_ModFecInicio_LostFocus()
    Dim vlFechaEfecto As Date 'Fecha de Efecto Real
    If Not IsDate(txt_ModFecInicio) Then
        txt_ModFecInicio = ""
        Exit Sub
    Else
        txt_ModFecInicio = CDate(txt_ModFecInicio)
    End If
    vlFechaEfecto = fgValidaFechaEfecto(txt_ModFecInicio, Txt_PenPoliza, vpNumOrden)
    'If txt_ModFecInicio < vpEfecto Then
    If txt_ModFecInicio < vlFechaEfecto Then
        MsgBox "Fecha de Inicio no puede ser anterior al próximo periodo de Pago: (" & vlFechaEfecto & ")", vbCritical
        txt_ModFecInicio = ""
        Exit Sub
    End If
    If Day(txt_ModFecInicio) <> 1 Then
        MsgBox "Fecha de Inicio debe comenzar el dia 1º", vbCritical
        txt_ModFecInicio = ""
        Exit Sub
    End If
    flCalculaCuotas
End Sub


Private Sub txt_ModNumCuotas_GotFocus()
    txt_ModNumCuotas.SelStart = 0
    txt_ModNumCuotas.SelLength = Len(txt_ModNumCuotas)
End Sub

Private Sub txt_ModNumCuotas_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub


Private Sub txt_ModNumCuotas_LostFocus()
    flCalculaCuotas
End Sub


Function flCalculaCuotas()
    Dim vlMonto As Double
    Dim vlCuota As Double, vlUltCuota As Double
    Dim vlNumCuotas As Long
    
    'Calcula la distribución de las Cuotas
    If Not IsDate(txt_ModFecInicio) Then
        Exit Function
    End If
    If Not IsNumeric(txt_ModNumCuotas) Then
        Exit Function
    End If
    If txt_ModNumCuotas = 0 Then
        Exit Function
    End If
    
    'Se asume Monto en Pesos revisar
    vlMonto = txt_ModTotal
    vlNumCuotas = txt_ModNumCuotas
    If txt_ModMoneda = "PESOS" Then
        vlCuota = Format(vlMonto / txt_ModNumCuotas, "##0")
    Else
        vlCuota = Format(vlMonto / txt_ModNumCuotas, "##0.00")
    End If
    vlUltCuota = 0
    txt_ModFecFin = DateAdd("d", -1, DateAdd("m", vlNumCuotas, txt_ModFecInicio))
    If Format(vlCuota * vlNumCuotas, "##0.00") <> Format(vlMonto, "##0.00") Then
        vlNumCuotas = vlNumCuotas - 1
        vlUltCuota = vlMonto - (vlCuota * vlNumCuotas)
    End If
    txt_ModCuota = vlCuota
    txt_ModUltCuota = vlUltCuota

End Function



Private Sub txt_PagMontoAct_Change()
    Dim vlConcepto As String
    If Not IsNumeric(txt_PagMontoAct) Then
        txt_PagMontoAct = "0"
        txt_PagMonto = txt_PagMontoAnt
    Else
        Dim vlDiferencia As Double
        vlDiferencia = Format(txt_PagMontoAnt - txt_PagMontoAct, "#0.00")
        txt_PagMonto = Abs(vlDiferencia)
        'variables en duro -- se deben cambiar
        If vlDiferencia > 0 Then
            vlConcepto = vlConceptoPEMenor
        Else
            vlConcepto = vlConceptoPEMayor
        End If
        txt_PagConcepto = vlConcepto
    End If
End Sub

Private Sub txt_PagMontoAct_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 44 And KeyAscii <> 46 Then
        KeyAscii = 0
    End If
End Sub


Private Sub Txt_PenPoliza_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Poliza

   If KeyAscii = 13 Then
      If Trim(Txt_PenPoliza) <> "" Then
        Txt_PenPoliza = Trim(UCase(Txt_PenPoliza))
        Txt_PenPoliza = Format(Txt_PenPoliza, "0000000000")
        Cmb_PenNumIdent.SetFocus
      Else
        Cmb_PenNumIdent.SetFocus
      End If
   End If

Exit Sub
Err_Poliza:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Txt_PenPoliza_LostFocus()
    Txt_PenPoliza = Trim(UCase(Txt_PenPoliza))
    Txt_PenPoliza = Format(Txt_PenPoliza, "0000000000")
End Sub

Function flValidarBen()
Dim vlFechaActual As String
Dim vlRegistro As ADODB.Recordset
Dim vlSwGrilla As Boolean
On Error GoTo Err_Validar

   Screen.MousePointer = 11
    

  
  'Verificar Número de Póliza, y saca el último Endoso
   vgPalabra = ""
   vgSql = ""
   If Txt_PenPoliza <> "" And Txt_PenNumIdent <> "" And Cmb_PenNumIdent.Text <> "" Then
      vgPalabra = "NUM_POLIZA = '" & Txt_PenPoliza & "' AND "
      vgPalabra = vgPalabra & "NUM_IDENBEN = '" & Txt_PenNumIdent & "' AND "
      vgPalabra = vgPalabra & "COD_TIPOIDENBEN = " & fgObtenerCodigo_TextoCompuesto(Cmb_PenNumIdent)
   Else
     If Txt_PenPoliza <> "" Then
        vgSql = "SELECT COUNT(NUM_POLIZA) AS REG_POLIZA"
        vgSql = vgSql & " FROM PP_TMAE_BEN WHERE"
        vgSql = vgSql & " NUM_POLIZA = '" & Txt_PenPoliza & "' and "
        vgSql = vgSql & " COD_ESTPENSION <> '10'"
'        vgSql = vgSql & " ORDER BY NUM_ENDOSO DESC, NUM_ORDEN ASC"
        Set vlRegistro = vgConexionBD.Execute(vgSql)
        If Not vlRegistro.EOF Then
           If (vlRegistro!reg_poliza) > 0 Then
               vgPalabra = "NUM_POLIZA = '" & Txt_PenPoliza & "' AND"
               vgPalabra = vgPalabra & " COD_ESTPENSION <> '10'"
           Else
               vgPalabra = "NUM_POLIZA = '" & Txt_PenPoliza & "'"
           End If
        Else
           vgPalabra = "NUM_POLIZA = '" & Txt_PenPoliza & "'"
        End If
     Else
        If Txt_PenNumIdent <> "" Then
           vgPalabra = "NUM_IDENBEN = '" & Txt_PenNumIdent & "' AND"
           vgPalabra = "COD_TIPOIDENBEN = " & fgObtenerCodigo_TextoCompuesto(Cmb_PenNumIdent)
        End If
     End If
   End If
                
    vlSwGrilla = True
    vgSql = "SELECT NUM_POLIZA,NUM_ENDOSO,NUM_ORDEN,NUM_IDENBEN,B.COD_TIPOIDENBEN, A.GLS_TIPOIDENCOR,"
    vgSql = vgSql & " COD_ESTPENSION,GLS_NOMBEN,GLS_NOMSEGBEN,GLS_PATBEN,GLS_MATBEN,FEC_MATRIMONIO "
    vgSql = vgSql & " FROM PP_TMAE_BEN B, MA_TPAR_TIPOIDEN A"
    vgSql = vgSql & " WHERE A.COD_TIPOIDEN = B.COD_TIPOIDENBEN AND "
    vgSql = vgSql & vgPalabra
    vgSql = vgSql & " ORDER BY num_orden asc, NUM_ENDOSO DESC "
    Set vlRegistro = vgConexionBD.Execute(vgSql)
    If Not vlRegistro.EOF Then
       If (vlRegistro!Cod_EstPension) = "10" Then
           MsgBox " El Beneficiario Seleccionado No Tiene Derecho a Pensión " & Chr(13) & _
                  "          Sólo podrá Consultar los Datos del Registro", vbInformation, "Información"
           vlSwGrilla = False
       End If
       Txt_PenNumIdent = vlRegistro!Num_IdenBen
       Cmb_PenNumIdent = " " & vlRegistro!Cod_TipoIdenBen & " - " & vlRegistro!gls_tipoidencor
       Txt_PenPoliza = (vlRegistro!num_poliza)
       Lbl_Endoso = (vlRegistro!num_endoso)
       vlNumOrden = (vlRegistro!Num_Orden)
       Lbl_PenNombre = (vlRegistro!Gls_NomBen) + " " + IIf(IsNull(vlRegistro!Gls_NomSegBen), "", vlRegistro!Gls_NomSegBen) + " " + (vlRegistro!Gls_PatBen) + " " + IIf(IsNull(vlRegistro!Gls_MatBen), "", vlRegistro!Gls_MatBen)
   Else
      Lbl_Endoso = ""
      MsgBox "El Beneficiario/Pensionado no corresponde o No se encuentra registrado.", vbCritical, "Error de Datos"
      Txt_PenPoliza.SetFocus
   End If
    vlRegistro.Close
             
    Fra_Poliza.Enabled = False
    Fra_Antecedentes.Enabled = True
    SSTab1.Enabled = True
    Screen.MousePointer = 0
   
Exit Function
Err_Validar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function
'---------------------FUNCIONES----------------------------------------

Function flRecibe(vlNumPoliza, vlNumIdent, vlTipoIdent, vlNumEndoso)
    Txt_PenPoliza = vlNumPoliza
    Txt_PenNumIdent = vlNumIdent
    Cmb_PenNumIdent = vlTipoIdent
'    Lbl_PenEndoso = vlNumEndoso
    Cmd_BuscarPol_Click
End Function


Public Function flObtieneBeneficiarios(iOrigenPE, iNumOrden) As Boolean
'Obtiene Beneficiarios Posiblemente Afectados
'iOrigenPE: 1:Indica reliquidar Pensión, 0:Indica no Reliquidar Pensión
'iNumOrden: Número de Orden del Beneficiario al que se le Reliquidará

Dim vlLinea As String, vlI As Long
Dim vlLineaBase As String, vlSql As String
Dim vlTB As ADODB.Recordset

flObtieneBeneficiarios = False
vlSql = "SELECT b.num_orden, b.num_idenben, b.cod_tipoidenben, b.gls_nomben, b.gls_nomsegben,"
vlSql = vlSql & " b.gls_patben, b.gls_matben, b.cod_par as par, c.gls_tipoidencor"
vlSql = vlSql & " FROM pp_tmae_ben b, ma_tpar_tipoiden c"
vlSql = vlSql & " WHERE b.cod_tipoidenben= c.cod_tipoiden"
If iOrigenPE = 1 Then
    vlSql = vlSql & " AND b.num_poliza = '" & Txt_PenPoliza & "'"
    vlSql = vlSql & " AND b.num_endoso = " & Lbl_Endoso
    If iNumOrden <> -1 Then
        vlSql = vlSql & " AND b.num_orden = " & iNumOrden
    End If
    If vgTipoBase = "ORACLE" Then
        vlSql = vlSql & " AND (b.fec_inipagopen IS NOT NULL OR b.fec_inipagopen <= TO_CHAR(SYSDATE,'YYYYMMDD'))"
    Else
        vlSql = vlSql & " AND (b.fec_inipagopen IS NOT NULL OR b.fec_inipagopen <= CONVERT(CHAR(8),GETDATE(),112))"
    End If
    vlLineaBase = "X" & vbTab & IIf(iOrigenPE = 1, "X", "") & vbTab
End If

vlI = 1
Set vlTB = vgConexionBD.Execute(vlSql)
If Not vlTB.EOF Then
    flObtieneBeneficiarios = True 'Encontró Datos
    Do While Not vlTB.EOF
        Msf_Beneficiarios.Row = vlI
        vlLinea = vlTB!Num_Orden & vbTab & vlTB!gls_tipoidencor & vbTab & vlTB!Num_IdenBen & vbTab
        vlLinea = vlLinea & Trim(vlTB!Gls_NomBen) & " " & Trim(vlTB!Gls_PatBen) & " " & Trim(vlTB!Gls_MatBen) & vbTab
        vlLinea = vlLinea & vlTB!PAR & vbTab & txt_PerDesde & vbTab & txt_PerHasta
        Msf_Beneficiarios.AddItem vlLineaBase & vlLinea, vlI
        vlTB.MoveNext
        vlI = vlI + 1
    Loop
End If

End Function

Function flLimpiaPagos()
    'Carpeta Pagos Reliquidados
    txt_PagNumOrden = ""
    txt_PagPeriodo = ""
    txt_PagConcepto = ""
    txt_PagFecInicio = ""
    txt_PagFecFin = ""
    txt_PagMonto = ""
    txt_PagMoneda = ""
    Msf_Pagos.Rows = 1
    Msf_Pagos.Rows = 2
End Function
Function flLimpiaModalidad()
    'Carpeta Modalidad de Pago
    txt_ModNumOrden = ""
    txt_ModConcepto = ""
    txt_ModFecInicio = ""
    txt_ModNumCuotas = ""
    txt_ModFecFin = ""
    txt_ModMoneda = ""
    txt_ModTotal = ""
    txt_ModCuota = ""
    txt_ModUltCuota = ""
    Msf_Modalidad.Rows = 1
    Msf_Modalidad.Rows = 2
End Function



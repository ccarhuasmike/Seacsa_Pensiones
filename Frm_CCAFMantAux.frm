VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Frm_CCAFMantAux 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenedor Carga de Conceptos CCAF"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   8280
   Begin VB.Frame Fra_CCAF 
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
      TabIndex        =   20
      Top             =   0
      Width           =   8055
      Begin VB.CommandButton Cmd_BuscarPeriodo 
         Height          =   375
         Left            =   6890
         Picture         =   "Frm_CCAFMantAux.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Buscar Póliza"
         Top             =   160
         Width           =   615
      End
      Begin VB.TextBox Txt_Anno 
         Height          =   285
         Left            =   3000
         MaxLength       =   4
         TabIndex        =   3
         Top             =   240
         Width           =   795
      End
      Begin VB.TextBox Txt_Mes 
         Height          =   285
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   2
         Top             =   240
         Width           =   360
      End
      Begin VB.ComboBox Cmb_CCAF 
         BackColor       =   &H00E0FFFF&
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   5235
      End
      Begin VB.ComboBox Cmb_TipoCon 
         BackColor       =   &H00E0FFFF&
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   960
         Width           =   5235
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Período "
         ForeColor       =   &H80000017&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   26
         Top             =   315
         Width           =   705
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "(Mes - Año)"
         ForeColor       =   &H80000017&
         Height          =   270
         Index           =   10
         Left            =   1095
         TabIndex        =   25
         Top             =   315
         Width           =   1125
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
         Index           =   11
         Left            =   2760
         TabIndex        =   24
         Top             =   255
         Width           =   195
      End
      Begin VB.Label Label9 
         Caption         =   "  Selección de Entidad  "
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
         TabIndex        =   23
         Top             =   0
         Width           =   2175
      End
      Begin VB.Label Label8 
         Caption         =   "Caja de Compensación"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   650
         Width           =   1875
      End
      Begin VB.Label Label7 
         Caption         =   "Tipo de Concepto"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1020
         Width           =   1875
      End
   End
   Begin VB.Frame Fra_Botones 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Width           =   8055
      Begin VB.CommandButton Cmd_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   5160
         Picture         =   "Frm_CCAFMantAux.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   200
         Width           =   730
      End
      Begin VB.CommandButton Cmd_Grabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   840
         Picture         =   "Frm_CCAFMantAux.frx":06DC
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   4080
         Picture         =   "Frm_CCAFMantAux.frx":0D96
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   6240
         Picture         =   "Frm_CCAFMantAux.frx":1450
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   3000
         Picture         =   "Frm_CCAFMantAux.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Eliminar 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   1920
         Picture         =   "Frm_CCAFMantAux.frx":1C04
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Eliminar Año"
         Top             =   200
         Width           =   720
      End
      Begin Crystal.CrystalReport Rpt_General 
         Left            =   7080
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin TabDlg.SSTab SStab_CCAF 
      Height          =   3225
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   5689
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Detalle por Póliza"
      TabPicture(0)   =   "Frm_CCAFMantAux.frx":1F46
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Fra_Detalle"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Fra_Rut"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Pólizas a Descontar"
      TabPicture(1)   =   "Frm_CCAFMantAux.frx":1F62
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Msf_GrillaCCAF"
      Tab(1).ControlCount=   1
      Begin VB.Frame Fra_Rut 
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
         Height          =   735
         Left            =   120
         TabIndex        =   36
         Top             =   360
         Width           =   7695
         Begin VB.TextBox Txt_Dgv 
            Height          =   285
            Left            =   3120
            MaxLength       =   1
            TabIndex        =   9
            Top             =   240
            Width           =   315
         End
         Begin VB.TextBox Txt_Rut 
            Height          =   285
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   8
            Top             =   240
            Width           =   1515
         End
         Begin VB.CommandButton Cmd_BuscarRut 
            Height          =   375
            Left            =   5520
            Picture         =   "Frm_CCAFMantAux.frx":1F7E
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Buscar Póliza"
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Lbl_Nombre 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
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
            Index           =   22
            Left            =   2880
            TabIndex        =   38
            Top             =   240
            Width           =   135
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Rut"
            Height          =   255
            Index           =   21
            Left            =   240
            TabIndex        =   37
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Fra_Detalle 
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
         Height          =   1815
         Left            =   120
         TabIndex        =   27
         Top             =   1200
         Width           =   7695
         Begin VB.TextBox Txt_MtoDescto 
            Height          =   285
            Left            =   1920
            MaxLength       =   10
            TabIndex        =   11
            Top             =   1320
            Width           =   1305
         End
         Begin VB.ComboBox Cmb_ModPago 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   960
            Width           =   3465
         End
         Begin VB.Label Lbl_FechaIni 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1920
            TabIndex        =   40
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Lbl_NumPoliza 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1920
            TabIndex        =   1
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "N° Póliza"
            Height          =   255
            Index           =   20
            Left            =   240
            TabIndex        =   35
            Top             =   280
            Width           =   855
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Modalidad de Pago"
            Height          =   255
            Index           =   19
            Left            =   240
            TabIndex        =   34
            Top             =   1000
            Width           =   1785
         End
         Begin VB.Label Lbl_FechaTer 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3600
            TabIndex        =   33
            Top             =   600
            Width           =   1215
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
            Index           =   16
            Left            =   3240
            TabIndex        =   32
            Top             =   600
            Width           =   225
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Fecha del Descuento"
            Height          =   255
            Index           =   15
            Left            =   240
            TabIndex        =   31
            Top             =   620
            Width           =   1785
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Monto a Descontar"
            Height          =   255
            Index           =   14
            Left            =   240
            TabIndex        =   30
            Top             =   1350
            Width           =   1785
         End
         Begin VB.Label Lbl_NumOrden 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   6120
            TabIndex        =   29
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Lbl_Nombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Orden"
            Height          =   255
            Index           =   13
            Left            =   5040
            TabIndex        =   28
            Top             =   240
            Width           =   735
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   39
         Top             =   600
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   6376
         _Version        =   393216
         Cols            =   5
         BackColor       =   14745599
         FormatString    =   $"Frm_CCAFMantAux.frx":2080
      End
      Begin MSFlexGridLib.MSFlexGrid Msf_GrillaCCAF 
         Height          =   2415
         Left            =   -74880
         TabIndex        =   18
         Top             =   600
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   4260
         _Version        =   393216
         Cols            =   5
         BackColor       =   14745599
         FormatString    =   $"Frm_CCAFMantAux.frx":2111
      End
   End
End
Attribute VB_Name = "Frm_CCAFMantAux"
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

    Frm_CCAFMantAux.Top = 0
    Frm_CCAFMantAux.Left = 0
    
    
        
    SStab_CCAF.Tab = 0
    SStab_CCAF.Enabled = True
    
    Fra_Rut.Enabled = True
    
    Fra_Detalle.Enabled = True
    


    
    
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Msf_GrillaCCAF_DblClick()
On Error GoTo Err_Msf_GrillaCCAF_DblClick
    
    Msf_GrillaCCAF.Col = 0
    If (Msf_GrillaCCAF.Text = "") Or (Msf_GrillaCCAF.Row = 0) Then
        MsgBox "No existen Detalles", vbExclamation, "Información"
        Exit Sub
    Else
        Msf_GrillaCCAF.Col = 0
      
        'Call flMostrarDatosPoliza
        
        SStab_CCAF.Tab = 0
        
    End If

Exit Sub
Err_Msf_GrillaCCAF_DblClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub



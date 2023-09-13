VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "Crystl32.OCX"
Begin VB.Form Frm_PensHabDescto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenedor de Haberes y Descuentos por Pensionado."
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   9120
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
      Height          =   1095
      Left            =   120
      TabIndex        =   37
      Top             =   0
      Width           =   8805
      Begin VB.TextBox Txt_PenPoliza 
         Height          =   285
         Left            =   840
         MaxLength       =   10
         TabIndex        =   0
         Top             =   360
         Width           =   1185
      End
      Begin VB.TextBox Txt_PenNumIdent 
         Height          =   285
         Left            =   5160
         MaxLength       =   16
         TabIndex        =   2
         Top             =   360
         Width           =   1635
      End
      Begin VB.CommandButton Cmd_BuscarPol 
         Height          =   375
         Left            =   8040
         Picture         =   "Frm_PensHabDescto.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Buscar Póliza"
         Top             =   240
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
         Left            =   8040
         Picture         =   "Frm_PensHabDescto.frx":0102
         TabIndex        =   4
         ToolTipText     =   "Buscar"
         Top             =   600
         Width           =   615
      End
      Begin VB.ComboBox Cmb_PenNumIdent 
         BackColor       =   &H00E0FFFF&
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   2235
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Póliza"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   44
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Lbl_PenNombre 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   255
         Left            =   840
         TabIndex        =   42
         Top             =   720
         Width           =   7095
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Ident."
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   41
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "N° End"
         Height          =   195
         Index           =   42
         Left            =   6840
         TabIndex        =   40
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Lbl_End 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   7440
         TabIndex        =   39
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   43
         Left            =   120
         TabIndex        =   38
         Top             =   0
         Width           =   1725
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   43
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Frame Fra_Botones 
      Height          =   1095
      Left            =   120
      TabIndex        =   23
      Top             =   5880
      Width           =   8925
      Begin VB.CommandButton Cmd_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   5640
         Picture         =   "Frm_PensHabDescto.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   730
      End
      Begin VB.CommandButton Cmd_Eliminar 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   2400
         Picture         =   "Frm_PensHabDescto.frx":07DE
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Eliminar Año"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   3480
         Picture         =   "Frm_PensHabDescto.frx":0B20
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   6720
         Picture         =   "Frm_PensHabDescto.frx":11DA
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   4560
         Picture         =   "Frm_PensHabDescto.frx":12D4
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Grabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   1320
         Picture         =   "Frm_PensHabDescto.frx":198E
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   720
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
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Haberes y Desctos."
      TabPicture(0)   =   "Frm_PensHabDescto.frx":2048
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Fra_HaberDescto"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Historia"
      TabPicture(1)   =   "Frm_PensHabDescto.frx":2064
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Msf_Grilla"
      Tab(1).ControlCount=   1
      Begin VB.Frame Fra_HaberDescto 
         Height          =   4215
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   8655
         Begin VB.ComboBox Cmb_HabDes 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   240
            Width           =   6105
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
            TabIndex        =   25
            Top             =   2400
            Width           =   8175
            Begin VB.ComboBox Cmb_MotSuspension 
               BackColor       =   &H00E0FFFF&
               Height          =   315
               Left            =   1920
               Style           =   2  'Dropdown List
               TabIndex        =   12
               Top             =   240
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
            Begin VB.TextBox Txt_Observacion 
               Height          =   615
               Left            =   1920
               MaxLength       =   255
               MultiLine       =   -1  'True
               TabIndex        =   14
               Top             =   960
               Width           =   5895
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
               TabIndex        =   36
               Top             =   0
               Width           =   3105
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "Motivo de Suspensión"
               Height          =   255
               Index           =   9
               Left            =   240
               TabIndex        =   28
               Top             =   360
               Width           =   1800
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "Fecha de Suspensión"
               Height          =   255
               Index           =   10
               Left            =   240
               TabIndex        =   27
               Top             =   720
               Width           =   1710
            End
            Begin VB.Label Lbl_Observacion 
               Caption         =   "Observación  "
               Height          =   255
               Left            =   240
               TabIndex        =   26
               Top             =   1080
               Width           =   1335
            End
         End
         Begin VB.TextBox Txt_FecInicio 
            Height          =   285
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   6
            Top             =   600
            Width           =   1185
         End
         Begin VB.TextBox Txt_NroCuotas 
            Height          =   285
            Left            =   2160
            MaxLength       =   3
            TabIndex        =   9
            Top             =   1320
            Width           =   585
         End
         Begin VB.TextBox Txt_MontoCuota 
            Height          =   285
            Left            =   2160
            MaxLength       =   13
            TabIndex        =   10
            Top             =   1680
            Width           =   1305
         End
         Begin VB.TextBox Txt_MontoTotal 
            Height          =   285
            Left            =   2160
            MaxLength       =   15
            TabIndex        =   11
            Top             =   2040
            Width           =   1305
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Código Haber/Descto."
            Height          =   255
            Index           =   6
            Left            =   360
            TabIndex        =   35
            Top             =   240
            Width           =   1785
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Fecha de Inicio"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   34
            Top             =   600
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
            TabIndex        =   33
            Top             =   600
            Width           =   225
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Código Moneda"
            Height          =   255
            Index           =   5
            Left            =   360
            TabIndex        =   32
            Top             =   960
            Width           =   1785
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Número de Cuotas"
            Height          =   255
            Index           =   7
            Left            =   360
            TabIndex        =   31
            Top             =   1320
            Width           =   1785
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Monto de la Cuota"
            Height          =   255
            Index           =   8
            Left            =   360
            TabIndex        =   30
            Top             =   1680
            Width           =   1785
         End
         Begin VB.Label Lbl_FecTermino 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3600
            TabIndex        =   7
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Monto Total "
            Height          =   255
            Index           =   12
            Left            =   360
            TabIndex        =   29
            Top             =   2040
            Width           =   1785
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Msf_Grilla 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   22
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
Attribute VB_Name = "Frm_PensHabDescto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vlRegistro   As ADODB.Recordset
Dim vlRegistro1  As ADODB.Recordset

Dim vlNumEndoso  As String
Dim vlNumOrden   As Integer
Dim vlNombre     As String
Dim vlCodDerpen  As String
Dim vlSwNo       As Boolean
Dim vlCodTab     As String
Dim vlCodFre     As String
Dim Habilita     As Boolean
Dim vlFecNac     As String
Dim vlPasa       As Boolean
Dim vlPasaBen    As Boolean
Dim vlSw         As Boolean
Dim vlSwVigPol   As Boolean
Dim vlSwGrilla   As Boolean
Dim vlSwSuspension As Boolean
Dim inivig       As String
Dim vlOperacion  As String
Dim vlCodMSus    As String
Dim vlFechaSus   As String
Dim vlOp         As String
Dim vlFecIni     As String
Dim vlFecTer     As String
Dim vlCodMon     As String
Dim vlCodHabDes  As String
Dim vlPos        As String
Dim vlCodSus     As String
Dim vlBuscaHaDes As Boolean

Enum enMoneda1
    numero = 7
    monto = 8
End Enum

Const clTopeMinimo = 0
Const clTopeMaximo = 9999999
Const clTopeMinimoPorc = 0
Const clTopeMaximoPorc = 100

Dim vlModFecTerNumCuotas As Boolean 'Indica si se debe modificar el valor del número de cuotas y la fecha de termino

Dim vlCodTipoIdenBenCau As String
Dim vlNumIdenBenCau As String

Function flLmpHabDes()

  Txt_FecInicio = ""
  Txt_NroCuotas = ""
  Txt_MontoCuota = ""
  Txt_MontoTotal = ""
  Lbl_FecTermino = ""
  
  If Cmb_Moneda.Text <> "" Then
     Cmb_Moneda.ListIndex = 0
  End If
  If Cmb_HabDes.Text <> "" Then
     Cmb_HabDes.ListIndex = 0
  End If
  If Cmb_MotSuspension.Text <> "" Then
     Cmb_MotSuspension.ListIndex = 0
  End If
  Txt_FecSuspension = ""
  Txt_Observacion = ""
  vlModFecTerNumCuotas = True
End Function

Function flValidarBen()
Dim vlFecActual As String
On Error GoTo Err_Validar

   Screen.MousePointer = 11
      
   vlSwVigPol = True
   
'   vlFecActual = fgBuscaFecServ

    vlCodTipoIdenBenCau = fgObtenerCodigo_TextoCompuesto(Cmb_PenNumIdent)
    vlNumIdenBenCau = Txt_PenNumIdent
    
   'Verificar Número de Póliza, y saca el último Endoso
   vgPalabra = ""
   vgSql = ""
   If Txt_PenPoliza <> "" And cmd_PenNumIdent And Txt_PenNumIdent <> "" Then
        vgPalabra = "NUM_POLIZA = '" & Txt_PenPoliza & "' AND "
        vgPalabra = vgPalabra & "cod_tipoidenBEN = " & (vlCodTipoIdenBenCau) & " and "
        vgPalabra = vgPalabra & "num_idenben = '" & (vlNumIdenBenCau) & "' "
   Else
     If Txt_PenPoliza <> "" Then
        vgSql = "SELECT COUNT(NUM_POLIZA) AS REG_POLIZA"
        vgSql = vgSql & " FROM PP_TMAE_BEN WHERE"
        vgSql = vgSql & " NUM_POLIZA = '" & Txt_PenPoliza & "' and "
        vgSql = vgSql & " COD_ESTPENSION <> '10'"
        'vgSql = vgSql & " ORDER BY NUM_ENDOSO DESC, NUM_ORDEN ASC"
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
        If Txt_PenNumIdent.Text <> "" Then
           ''*vgPalabra = "RUT_BEN = " & Format(Txt_PenRut, "#0") & ""
            vgPalabra = "cod_tipoidenBEN = " & (vlCodTipoIdenBenCau) & " "
            vgPalabra = vgPalabra & "AND num_idenben = '" & (vlNumIdenBenCau) & "' "
        End If
     End If
   End If
                
   vlSwGrilla = True
   
   vgSql = ""
   vgSql = "SELECT NUM_POLIZA,NUM_ENDOSO,FEC_NACBEN,NUM_ORDEN,COD_TIPOIDENBEN,NUM_IDENBEN,MTO_PENSION,"
   vgSql = vgSql & " COD_ESTPENSION,GLS_NOMBEN,GLS_NOMSEGBEN,GLS_PATBEN,GLS_MATBEN,FEC_MATRIMONIO "
   vgSql = vgSql & " FROM PP_TMAE_BEN"
   vgSql = vgSql & " Where "
   vgSql = vgSql & vgPalabra
   vgSql = vgSql & " ORDER BY num_orden asc, NUM_ENDOSO DESC "
   Set vlRegistro = vgConexionBD.Execute(vgSql)
   If Not vlRegistro.EOF Then
      If (vlRegistro!Cod_EstPension) = "10" Then
          MsgBox " El Beneficiario Seleccionado No Tiene Derecho a Pensión " & Chr(13) & _
                 "          Sólo podrá Consultar los Datos del Registro", vbInformation, "Información"
          vlSwGrilla = False
      End If
      
    vlCodTipoIdenBenCau = vlRegistro!Cod_TipoIdenBen
    vlNumIdenBenCau = Trim(vlRegistro!Num_IdenBen)
    
     Txt_PenPoliza = (vlRegistro!num_poliza)
     Call fgBuscarPosicionCodigoCombo(vlCodTipoIdenBenCau, Cmb_PenNumIdent)
     Txt_PenNumIdent.Text = vlNumIdenBenCau
          
     vlPasaBen = False
     vlCodDerpen = (vlRegistro!Cod_EstPension)
     vlNumEndoso = (vlRegistro!num_endoso)
     Lbl_End = vlNumEndoso
     vlNumOrden = (vlRegistro!Num_Orden)
     vlNombre = (vlRegistro!Gls_NomBen) + " " + IIf(IsNull(vlRegistro!Gls_NomSegBen), "", (vlRegistro!Gls_NomSegBen)) + " " + (vlRegistro!Gls_PatBen) + " " + IIf(IsNull(vlRegistro!Gls_MatBen), "", (vlRegistro!Gls_MatBen))
     Lbl_PenNombre = vlNombre
     Fra_Poliza.Enabled = False
'    Cmb_HabDes.Enabled = True
'    Txt_FecInicio.Enabled = True
     vlFecNac = (vlRegistro!Fec_NacBen)
     SSTab1.Enabled = True
     Cmb_HabDes.Enabled = True
     Cmb_HabDes.SetFocus
   Else
      'I---- ABV 23/08/2004 ---
      Lbl_End = ""
      MsgBox "El Beneficiario/Pensionado No tiene Derecho a Pensión o No se encuentra registrado.", vbCritical, "Error de Datos"
      'Txt_PenRut = ""
      'Txt_PenDigito = ""
      'Txt_PenPoliza = ""
      Txt_PenPoliza.SetFocus
      'F---- ABV 23/08/2004 ---
   End If
   vlRegistro.Close
     
   If (Lbl_End <> "") And (Lbl_PenNombre <> "") And (Txt_PenPoliza <> "") And (Txt_PenNumIdent <> "") And (Cmb_PenNumIdent <> "") Then
       flLmpGrilla
       flBuscarHistorico
       
       Cmd_Limpiar.Enabled = True
       Cmb_HabDes.SetFocus
   End If
     
   Screen.MousePointer = 0
     
Exit Function
Err_Validar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flBuscarHistorico()
On Error GoTo Err_Carga
     
     vgSql = ""
     vlCodMon = "TM"
     vlCodSus = "SHD"
     vgSql = "SELECT m.COD_MODORIGEN,p.COD_CONHABDES,m.GLS_CONHABDES,p.FEC_INIHABDES,p.FEC_TERHABDES,"
     vgSql = vgSql & "p.NUM_CUOTAS,p.MTO_CUOTA,p.MTO_TOTAL,p.COD_MONEDA,"
     'vgSql = vgSql & "t.GLS_ELEMENTO,"
     vgSql = vgSql & "p.COD_MOTSUSHABDES,m.COD_INTERNO,"
     vgSql = vgSql & "c.GLS_ELEMENTO as elemento , p.FEC_SUSHABDES,p.GLS_OBSHABDES, m.COD_TIPMOV "
     'vgSql = vgSql & "FROM PP_TMAE_HABDES P, MA_TPAR_CONHABDES M, MA_TPAR_TABCOD T,MA_TPAR_TABCOD C "
     vgSql = vgSql & "FROM PP_TMAE_HABDES P, MA_TPAR_CONHABDES M, MA_TPAR_TABCOD C "
     vgSql = vgSql & "Where p.NUM_POLIZA =  '" & (Txt_PenPoliza) & "' AND "
     'vgSql = vgSql & "p.NUM_ENDOSO = '" & (vlNumEndoso) & "' AND "
     vgSql = vgSql & "p.NUM_ORDEN = " & (vlNumOrden) & " AND "
     vgSql = vgSql & "p.COD_CONHABDES = m.COD_CONHABDES AND "
     'vgSql = vgSql & "t.COD_TABLA = '" & (vlCodMon) & "' AND "
     'vgSql = vgSql & "p.COD_MONEDA = t.COD_ELEMENTO AND "
     vgSql = vgSql & "C.COD_TABLA = '" & (vlCodSus) & "'   AND "
     'vgSql = vgSql & "p.COD_MOTSUSHABDES = - C.COD_ELEMENTO  "
     vgSql = vgSql & "p.COD_MOTSUSHABDES =  C.COD_ELEMENTO  "
     vgSql = vgSql & "ORDER BY p.FEC_INIHABDES desc,p.COD_CONHABDES"
     Set vlRegistro = vgConexionBD.Execute(vgSql)
     If Not vlRegistro.EOF Then
        While Not vlRegistro.EOF
                
              vlInicio = (vlRegistro!Fec_IniHabDes)
              vlAnno = Mid(vlInicio, 1, 4)
              vlMes = Mid(vlInicio, 5, 2)
              vlDia = Mid(vlInicio, 7, 2)
              vlInicio = DateSerial((vlAnno), (vlMes), (vlDia))

              vlTermino = (vlRegistro!FEC_TERHabDes)
              vlAnno = Mid(vlTermino, 1, 4)
              vlMes = Mid(vlTermino, 5, 2)
              vlDia = Mid(vlTermino, 7, 2)
              vlTermino = DateSerial((vlAnno), (vlMes), (vlDia))

              If IsNull(vlRegistro!Fec_SusHabDes) Then
                 vlFecSus = ""
              Else
                 vlFecSus = (vlRegistro!Fec_SusHabDes)
                 vlAnno = Mid(vlFecSus, 1, 4)
                 vlMes = Mid(vlFecSus, 5, 2)
                 vlDia = Mid(vlFecSus, 7, 2)
                 vlFecSus = DateSerial((vlAnno), (vlMes), (vlDia))
              End If
              'Msf_Grilla.AddItem (Trim(vlRegistro!Cod_ConHabDes) & " - " & _
                                  Trim(vlRegistro!gls_ConHabDes)) & vbTab & _
                                 (vlInicio) & vbTab & (vlTermino) & vbTab & _
                                  (vlRegistro!Num_Cuotas) & vbTab & _
                                 (Format(vlRegistro!MTO_CUOTA, "#,#0.00")) & vbTab & _
                                 (Trim(vlRegistro!COD_MONEDA) & " - " & _
                                  Trim(vlRegistro!gls_elemento)) & vbTab & _
                                  (Trim(vlRegistro!Cod_MotSusHabDes) & " - " & _
                                  Trim(vlRegistro!elemento)) & vbTab & _
                                 (vlFecSus) & vbTab & Trim(vlRegistro!GLS_OBSHABDES)
              
              Msf_Grilla.AddItem (Trim(vlRegistro!Cod_ConHabDes) & " - " & _
                                 Trim(vlRegistro!gls_ConHabDes)) & _
                                 "   (" & vlRegistro!cod_tipmov & ")" & vbTab & _
                                 (vlInicio) & vbTab & (vlTermino) & vbTab & _
                                 Format(vlRegistro!Num_Cuotas, "#0") & vbTab & _
                                 (Format(vlRegistro!MTO_CUOTA, "#,#0.00")) & vbTab & _
                                 (Format(vlRegistro!mto_total, "#,#0.00")) & vbTab & _
                                 Trim(vlRegistro!Cod_Moneda) & vbTab & _
                                 (Trim(vlRegistro!Cod_MotSusHabDes) & " - " & _
                                 Trim(vlRegistro!elemento)) & vbTab & _
                                 (vlFecSus) & vbTab & Trim(vlRegistro!GLS_OBSHABDES) & vbTab & _
                                 (vlRegistro!COD_INTERNO) & vbTab & _
                                 (vlRegistro!COD_MODORIGEN)
              vlRegistro.MoveNext
        Wend
     End If
     Screen.MousePointer = 0
     vlRegistro.Close

Exit Function
Err_Carga:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flLmpGrilla()

    Msf_Grilla.Clear
    Msf_Grilla.Rows = 1
    Msf_Grilla.RowHeight(0) = 250
    Msf_Grilla.Row = 0
    
    Msf_Grilla.Col = 0
    Msf_Grilla.Text = "Haber/Descto"
    Msf_Grilla.ColWidth(0) = 3200
    Msf_Grilla.ColAlignment(0) = 1
    
    Msf_Grilla.Col = 1
    Msf_Grilla.Text = "F.Inicio"
    Msf_Grilla.ColWidth(1) = 950
    Msf_Grilla.ColAlignment(1) = 1
    
    Msf_Grilla.Col = 2
    Msf_Grilla.Text = "F.Término"
    Msf_Grilla.ColWidth(2) = 950
    Msf_Grilla.ColAlignment(2) = 1
    
    Msf_Grilla.Col = 3
    Msf_Grilla.Text = "NºCuota"
    Msf_Grilla.ColWidth(3) = 600
    Msf_Grilla.ColAlignment(3) = 1
    
    Msf_Grilla.Col = 4
    Msf_Grilla.Text = "Mto.Cuotas"
    Msf_Grilla.ColWidth(4) = 1100
    
    Msf_Grilla.Col = 5
    Msf_Grilla.Text = "Mto. Total"
    Msf_Grilla.ColWidth(5) = 1300
    
    Msf_Grilla.Col = 6
    Msf_Grilla.Text = "Moneda"
    Msf_Grilla.ColWidth(6) = 800

    Msf_Grilla.Col = 7
    Msf_Grilla.Text = "Motivo de Suspensión"
    Msf_Grilla.ColWidth(7) = 3000
    Msf_Grilla.ColAlignment(7) = 1
    
    Msf_Grilla.Col = 8
    Msf_Grilla.Text = "F.Suspensión"
    Msf_Grilla.ColWidth(8) = 1200
    Msf_Grilla.ColAlignment(8) = 1
    
    Msf_Grilla.Col = 9
    Msf_Grilla.Text = "Observación"
    Msf_Grilla.ColWidth(9) = 0

    Msf_Grilla.Col = 10
    Msf_Grilla.Text = "Cod_Interno"
    Msf_Grilla.ColWidth(10) = 0

    Msf_Grilla.Col = 11
    Msf_Grilla.Text = "Mod.Origen"
    Msf_Grilla.ColWidth(11) = 800
    
End Function

Function flHaberDescuento()
On Error GoTo Err_HD

     vlFechaIni = Trim(Txt_FecInicio)
     If (Txt_FecInicio) <> "" Then
         If (flValidaFecha(vlFechaIni) = True) Then
            'transforma la fecha al formato yyyymmdd
             Screen.MousePointer = 11
             vlSwVigPol = True
             If fgValidaVigenciaPoliza(Txt_PenPoliza, Trim(Txt_FecInicio)) = False Then
                MsgBox "La Fecha Ingresada no se Encuentra dentro del Rango de Vigencia de la Póliza" & Chr(13) & _
                       "O esta No Vigente. No se Ingresara Ni Modificara Información. ", vbCritical, "Operación Cancelada"
                Screen.MousePointer = 0
                vlSwVigPol = False
                Exit Function
             End If
                          
'             If fgValidaPagoPension(Trim(Txt_FecInicio), Txt_PenPoliza, vlNumOrden) = False Then
'                MsgBox "Ya se ha relizado el Proceso de Cálculo de Pensión para ésta Fecha", vbCritical, "Operación Cancelada"
'                Screen.MousePointer = 0
'                Exit Function
'             End If
'
             vlFechaIni = Format(CDate(Trim(vlFechaIni)), "yyyymmdd")
            'se valida que exista información para esa fecha en la BD
             Call flBuscaHaDes(vlFechaIni)
             If vlBuscaHaDes = True Then
                Txt_FecInicio.Enabled = False
                Cmb_HabDes.Enabled = False
             Else
               If vlSwNo = True Then
                  Screen.MousePointer = 0
                  Txt_FecInicio.SetFocus
                  Exit Function
               End If
             End If
             Cmb_Moneda.SetFocus
         End If
       
     Else
        MsgBox "Debe Ingresar la Fecha de Inicio de la Vigencia ", vbCritical, "Falta Información"
        Txt_FecInicio.SetFocus
        Exit Function
     End If
     Screen.MousePointer = 0

Exit Function
Err_HD:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(2) & Err.Description & " ]"
    End Select
End Function

Function flBuscaHaDes(iFecha)
On Error GoTo Err_buscavig
    vlBuscaHaDes = False
    
    vlSwSuspension = True
    
    vlCodHabDes = Mid(Cmb_HabDes.Text, 1, (InStr(1, Cmb_HabDes, "-") - 1))
    vgSql = ""
    vgSql = "select * from PP_TMAE_HABDES where "
    vgSql = vgSql & " NUM_POLIZA = '" & Txt_PenPoliza & "' and "
    vgSql = vgSql & " NUM_ENDOSO = " & vlNumEndoso & " and "
    vgSql = vgSql & " NUM_ORDEN = " & vlNumOrden & " and "
    vgSql = vgSql & " COD_CONHABDES = '" & Trim(vlCodHabDes) & "' and "
    vgSql = vgSql & " FEC_INIHABDES = '" & iFecha & "'"
    Set vlRegistro1 = vgConexionBD.Execute(vgSql)
    If Not (vlRegistro1.EOF) Then
         
            If fgValidaPagoPension(Trim(Txt_FecInicio), Txt_PenPoliza, vlNumOrden) = False Then
               Fra_Suspension.Enabled = True
               vlSwSuspension = True
            Else
               Fra_Suspension.Enabled = False
               vlSwSuspension = False
            End If
   
            vlFecFin = (vlRegistro1!FEC_TERHabDes)
            vlAnno = Mid(vlFecFin, 1, 4)
            vlMes = Mid(vlFecFin, 5, 2)
            vlDia = Mid(vlFecFin, 7, 2)
            Lbl_FecTermino = DateSerial((vlAnno), (vlMes), (vlDia))
            vlCodMon = (vlRegistro1!Cod_Moneda)
            If Cmb_Moneda.ListCount <> 0 Then
               For vlI = 0 To Cmb_Moneda.ListCount - 1
                   Cmb_Moneda.ListIndex = vlI
                   vlCodigo = Trim(Mid(Cmb_Moneda, 1, (InStr(1, Cmb_Moneda, "-") - 1)))
                   If vlCodigo = vlCodMon Then
                      Exit For
                   End If
               Next vlI
            End If
            
            Txt_NroCuotas = (vlRegistro1!Num_Cuotas)
            Txt_MontoCuota = Format((vlRegistro1!MTO_CUOTA), "#,#0.00")
            Txt_MontoTotal = Format((vlRegistro1!mto_total), "#,#0.00")
            
            If (vlRegistro1!Fec_SusHabDes) <> "" Then
                vlFecSuspension = (vlRegistro1!Fec_SusHabDes)
                vlAnno = Mid(vlFecSuspension, 1, 4)
                vlMes = Mid(vlFecSuspension, 5, 2)
                vlDia = Mid(vlFecSuspension, 7, 2)
                Txt_FecSuspension = DateSerial((vlAnno), (vlMes), (vlDia))
            Else
                Txt_FecSuspension = ""
            End If
            vlCodMSus = (vlRegistro1!Cod_MotSusHabDes)
            If Cmb_MotSuspension.ListCount <> 0 Then
               For vlI = 0 To Cmb_MotSuspension.ListCount - 1
                   Cmb_MotSuspension.ListIndex = vlI
                   vlCodigo = Trim(Mid(Cmb_MotSuspension, 1, (InStr(1, Cmb_MotSuspension, "-") - 1)))
                   If vlCodigo = vlCodMSus Then
                      Exit For
                   End If
               Next vlI
            End If
            If IsNull(vlRegistro1!GLS_OBSHABDES) Then
                Txt_Observacion = ""
            Else
                Txt_Observacion = (vlRegistro1!GLS_OBSHABDES)
            End If
            vlBuscaHaDes = True
    Else
      If fgValidaPagoPension(Trim(Txt_FecInicio), Txt_PenPoliza, vlNumOrden) = False Then
         MsgBox "Ya se ha relizado el Proceso de Cálculo de Pensión para ésta Fecha", vbCritical, "Operación Cancelada"
         Screen.MousePointer = 0
         vlBuscaHaDes = False
         Fra_Suspension.Enabled = True
         vlSwNo = True
         Exit Function
      Else
         Fra_Suspension.Enabled = False
         vlSwSuspension = False
         vlSwNo = False
      End If
         
         
         
'      End If
    End If
    vlRegistro1.Close
Exit Function

Err_buscavig:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flValidaFecha(iFecha)
On Error GoTo Err_valfecha
      flValidaFecha = False
     
     'valida que la fecha este correcta
      If Trim(iFecha <> "") Then
         If Not IsDate(iFecha) Then
                MsgBox "La Fecha ingresada no es una fecha válida.", vbCritical, "Dato Incorrecto"
                Exit Function
         End If

         If (Year(iFecha) < 1900) Then
             MsgBox "La Fecha ingresada es menor a la mínima que se puede ingresar (1900).", vbCritical, "Dato Incorrecto"
             Exit Function
         End If
         flValidaFecha = True
     End If

Exit Function
Err_valfecha:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(2) & Err.Description & " ]"
    End Select
End Function

Function flRecibe(NPoliza, Rut, Digito, NEndoso)

    Txt_PenPoliza = NPoliza
    Txt_PenRut = Rut
    Txt_PenDigito = Digito
    Lbl_End = NEndoso
    Cmd_BuscarPol_Click
    
End Function
Function flCalculaFecha()

 If (Txt_FecInicio > "") Then
    vlFechaIni = Format(Txt_FecInicio, "yyyymmdd")
    vlAño = Mid(vlFechaIni, 1, 4)
    vlMes = Mid(vlFechaIni, 5, 2)
    vlMes = Int(vlMes) + Int(Txt_NroCuotas)
    vlDia = Mid(vlFechaIni, 7, 2)
    Lbl_FecTermino = DateAdd("d", -1, DateSerial(vlAño, vlMes, vlDia))
 End If
 
End Function

Sub flImpresion()
Dim vlTipoI As String
Dim vlArchivo As String

Err.Clear
On Error GoTo Errores1
   
   Screen.MousePointer = 11
   
   vlArchivo = strRpt & "PP_Rpt_PensHabDescto.rpt"   '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
       MsgBox "Archivo de Reporte de Listados no se encuentra en el directorio de la Aplicación.", 16, "Archivo no encontrado"
       Screen.MousePointer = 0
       Exit Sub
   End If
   
   vgQuery = ""
   vgQuery = "{PP_TMAE_HABDES.NUM_POLIZA} = '" & Trim(Txt_PenPoliza) & "' and "
   vgQuery = vgQuery & " {PP_TMAE_HABDES.NUM_ORDEN} = " & vlNumOrden & ""
   'vgQuery = vgQuery & " {MA_TPAR_CONHABDES.COD_MODORIGEN} <> '" & vgCodTabla_GarEst & "' AND "
   'vgQuery = vgQuery & " {MA_TPAR_CONHABDES.COD_MODORIGEN} <> 'CCAF' "
   'vgQuery = vgQuery & " AND {PP_TMAE_BEN.NUM_POLIZA} = '" & Trim(Txt_PenPoliza) & "' "
   'vgQuery = vgQuery & " AND {PP_TMAE_BEN.NUM_ENDOSO} = " & vlNumEndoso & " "
   
   
   Rpt_HabDes.Reset
   Rpt_HabDes.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
   'Rpt_General.DataFiles(0) = vgRutaBasedeDatos       ' o App.Path & "\Nestle.mdb"
   'Rpt_General.Connect = "ODBC;DATABASE= " & vgNombreBaseDatos & ";DSN=" & vgDsn
   Rpt_HabDes.Connect = vgRutaDataBase
'   Rpt_General.SelectionFormula = ""
   Rpt_HabDes.SelectionFormula = vgQuery
   Rpt_HabDes.Formulas(0) = ""
   Rpt_HabDes.Formulas(1) = ""
   Rpt_HabDes.Formulas(2) = ""
   
   vlTipoI = Trim(vlCodTipoIdenBenCau) & " - " & fgBuscarNombreTipoIden(vlCodTipoIdenBenCau, True)
   Rpt_HabDes.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
   Rpt_HabDes.Formulas(1) = "NombreSistema= '" & vgNombreSistema & "'"
   Rpt_HabDes.Formulas(2) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
   Rpt_HabDes.Formulas(3) = "Poliza = '" & Trim(Txt_PenPoliza) & "'"
   Rpt_HabDes.Formulas(4) = "Endoso = '" & vlNumEndoso & "'"
   Rpt_HabDes.Formulas(5) = "TipoIden = '" & vlTipoI & "'"
   Rpt_HabDes.Formulas(6) = "NumIden = '" & Trim(Txt_PenNumIdent) & "'"
   Rpt_HabDes.Formulas(7) = "Nombre_Bene = '" & Trim(Lbl_PenNombre) & "'"
   Rpt_HabDes.Formulas(8) = "Fec_Nac = '" & DateSerial(Mid(vlFecNac, 1, 4), Mid(vlFecNac, 5, 2), Mid(vlFecNac, 7, 2)) & "'"
   
   Rpt_HabDes.Destination = crptToWindow
   Rpt_HabDes.WindowState = crptMaximized
   Rpt_HabDes.WindowTitle = "Informe Haberes y Descuentos por Pensionado"
   Rpt_HabDes.Action = 1
   
   Screen.MousePointer = 0
   
Exit Sub
Errores1:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " " & "Numero de Error : " & Err.Number & Err.HelpFile & Err.HelpContext, vbExclamation, "Mensaje de error ..."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub Cmb_Moneda_Click()
If (Cmb_Moneda <> "") Then
    vgPalabra = Trim(Mid(Cmb_Moneda, 1, (InStr(1, Cmb_Moneda, "-") - 1)))
    If (vgPalabra = "PRCIM") Or (vgPalabra = "PRCTR") Or (vgPalabra = "PRIMF") Or (vgPalabra = "PRTMF") Then
        Lbl_Nombre(enMoneda1.monto) = "Porcentaje de la Renta"
        Txt_MontoTotal.Enabled = True
    Else
        Lbl_Nombre(enMoneda1.monto) = "Monto de la Cuota"
        Txt_MontoTotal.Enabled = False
        'Validar que cuadren Nro Cuotas y Monto.
        If vlModFecTerNumCuotas Then
            If IsNumeric(Txt_NroCuotas) Then
               flCalculaFecha
            End If
            If (Txt_MontoTotal.Enabled = False) Then
                If (Txt_MontoCuota <> "") And (Txt_NroCuotas <> "") Then
                    Txt_MontoTotal = CDbl(Txt_MontoCuota) * CDbl(Txt_NroCuotas)
                    Txt_MontoTotal = Format(Txt_MontoTotal, "#,#0.00")
                End If
            End If
        End If
    End If
    'Si se trata de Porcentaje Fijo, no se ingresa Número de Cuotas
    If (vgPalabra = "PRIMF") Or (vgPalabra = "PRTMF") Then
        Txt_NroCuotas.Enabled = False
        If vlModFecTerNumCuotas Then
            Txt_NroCuotas.Text = "999"
            Lbl_FecTermino = DateSerial(Mid(vgTopeFecFin, 1, 4), Mid(vgTopeFecFin, 5, 2), Mid(vgTopeFecFin, 7, 2))
        End If
    Else
        Txt_NroCuotas.Enabled = True
    End If
    
End If
End Sub

Private Sub Cmb_Moneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Txt_NroCuotas.Enabled = True Then
        Txt_NroCuotas.SetFocus
    Else
        Txt_MontoCuota.SetFocus
    End If
End If
End Sub

Private Sub Cmb_HabDes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       'If Cmb_HabDes.Text <> "" Then
          Txt_FecInicio.Enabled = True
          Txt_FecInicio.SetFocus
       'End If
    End If
End Sub

Private Sub Cmb_MotSuspension_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       'If Cmb_MotSuspension.Text <> "" Then
          Txt_FecSuspension.SetFocus
       'End If
    End If
End Sub

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
On Error GoTo Err_Buscar

    Frm_Busqueda.flInicio ("Frm_PensHabDescto")

Exit Sub
Err_Buscar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_BuscarPol_Click()
On Error GoTo Err_Buscar

  If Trim(Txt_PenPoliza) <> "" Or Trim(Txt_PenNumIdent) <> "" Then
''       If Trim(Cmb_PenNumIdent) <> "" Then
''          If Trim(Txt_PenNumIdent) = "" Then
''             MsgBox "Debe ingresar el Número de Identificación.", vbCritical, "Error de Datos"
''             Txt_PenNumIdent.SetFocus
''             Exit Sub
''          End If
          Txt_PenNumIdent = Trim(UCase(Txt_PenNumIdent))
''       End If
       'Permite Buscar los Datos del Beneficiario
        flValidarBen
   Else
     MsgBox "Debe ingresar el NºPóliza o la Identificación del Pensionado", vbCritical, "Error de Datos"
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

Private Sub Cmd_Cancelar_Click()
On Error GoTo Err_Cancelar

    flLmpHabDes
'    flDesHabDes
'    Cmd_Limpiar.Enabled = False
    SSTab1.Tab = 0
    SSTab1.Enabled = False
'    Cmb_HabDes.Enabled = False
'    Txt_FecInicio.Enabled = False
    Lbl_End = ""
    Fra_Poliza.Enabled = True
'    Cmd_BuscarPol.Enabled = True
'    Cmd_Buscar.Enabled = True
'    Txt_PenPoliza.Enabled = True
'    Txt_PenRut.Enabled = True
'    Txt_PenDigito.Enabled = True
'    Lbl_End.Enabled = True
'    Lbl_PenNombre.Enabled = True
    Txt_PenPoliza = ""
    If (Cmb_PenNumIdent.ListCount <> 0) Then
        Cmb_PenNumIdent.ListIndex = 0
    End If
    Txt_PenNumIdent = ""
    Lbl_End = ""
    Lbl_PenNombre = ""
    flLmpGrilla
    Txt_PenPoliza.SetFocus
    
Exit Sub
Err_Cancelar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Eliminar_Click()
On Error GoTo Err_Eliminar

''    If Fra_Poliza.Enabled = True Then
''       Exit Sub
''    End If
    
    If vlSwGrilla = False Then
       MsgBox " El Beneficiario Seleccionado No Tiene Derecho a Pensión ", vbCritical, "Operacion Cancelada"
       Exit Sub
    End If
    
'    If vlSwVigPol = False Then
'       MsgBox "Póliza Ingresada No se Encuentra vigente para el Sistema", vbCritical, "Operación Cancelada"
'       Exit Sub
'    End If
   
    'Validar el ingreso de las Claves para la Eliminación del Concepto
    If Txt_PenPoliza = "" Then
        MsgBox "Debe ingresar el Nº de la Póliza.", vbInformation, "Error de Datos"
        Txt_PenPoliza.SetFocus
        Exit Sub
    End If
    
    If Cmb_PenNumIdent = "" Then
        MsgBox "Debe ingresar la Identificación del Pensionado/Beneficiario.", vbInformation, "Error de Datos"
        Cmb_PenNumIdent.SetFocus
        Exit Sub
    End If
    
    If Txt_PenNumIdent = "" Then
        MsgBox "Debe ingresar la Identificación del Pensionado/Beneficiario.", vbInformation, "Error de Datos"
        Txt_PenNumIdent.SetFocus
        Exit Sub
    End If
    
    If (Trim(Cmb_HabDes) = "") Then
        MsgBox "Debe indicar el Tipo de Concepto de Haber/Descuento a ingresar.", vbInformation, "Error de Datos"
        Cmb_Moneda.SetFocus
        Exit Sub
    End If
    
    If (Txt_FecInicio = "") Or (Not IsDate(Txt_FecInicio)) Then
        MsgBox "Debe Ingresar la Fecha de Inicio de Vigencia del Haber/Descuento.", vbInformation, "Error de Datos"
        Txt_FecInicio.SetFocus
        Exit Sub
    End If
    
    vlSwVigPol = True
    If fgValidaVigenciaPoliza(Txt_PenPoliza, Trim(Txt_FecInicio)) = False Then
          MsgBox "La Fecha Ingresada no se Encuentra dentro del Rango de Vigencia de la Póliza" & Chr(13) & _
                 "O esta No Vigente. No se Ingresara Ni Modificara Información. ", vbCritical, "Operación Cancelada"
       Screen.MousePointer = 0
       vlSwVigPol = False
       Exit Sub
    End If
 
    If fgValidaPagoPension(Trim(Txt_FecInicio), Txt_PenPoliza, vlNumOrden) = False Then
       MsgBox "Ya se ha relizado el Proceso de Cálculo de Pensión para ésta fecha" & Chr(13) & _
              "              El Registro No se puede Eliminar", vbCritical, "Operación Cancelada"
       Screen.MousePointer = 0
       Cmd_Salir.SetFocus
       Exit Sub
    End If
    
    vlOperacion = ""
    Screen.MousePointer = 11
    vlCodHabDes = Trim(Mid(Cmb_HabDes, 1, (InStr(1, Cmb_HabDes, "-") - 1)))
    inivig = Txt_FecInicio
    inivig = Format(CDate(Trim(inivig)), "yyyymmdd")
    
    vgQuery = "SELECT NUM_POLIZA,NUM_ORDEN,COD_CONHABDES,FEC_INIHABDES "
    vgQuery = vgQuery & "FROM PP_TMAE_HABDES WHERE "
    vgQuery = vgQuery & "NUM_POLIZA = '" & Txt_PenPoliza & "' And "
    vgQuery = vgQuery & "NUM_ORDEN = " & vlNumOrden & " AND "
    vgQuery = vgQuery & "COD_CONHABDES = '" & (vlCodHabDes) & "' and "
    vgQuery = vgQuery & "FEC_INIHABDES = '" & inivig & "'"
    Set vgRs = vgConexionBD.Execute(vgQuery)
    If Not (vgRs.EOF) Then
        vlOperacion = "E"
    End If
    vgRs.Close

    If (vlOperacion = "E") Then
        vgRes = MsgBox(" ¿ Esta seguro que desea Eliminar los Datos ? ", vbQuestion + vbYesNo + 256, "Operación de Eliminación")
        If vgRes <> 6 Then
            Screen.MousePointer = 0
            Exit Sub
        End If
        
        vgQuery = "DELETE FROM PP_TMAE_HABDES WHERE "
        vgQuery = vgQuery & "NUM_POLIZA = '" & Txt_PenPoliza & "' And "
        vgQuery = vgQuery & "NUM_ORDEN = " & vlNumOrden & " AND "
        vgQuery = vgQuery & "COD_CONHABDES = '" & (vlCodHabDes) & "' and "
        vgQuery = vgQuery & "FEC_INIHABDES = '" & (inivig) & "'"
        vgConexionBD.Execute (vgQuery)
                 
        flLmpGrilla
        flBuscarHistorico
        Cmd_Limpiar_Click
        Txt_FecInicio.Enabled = True
        Txt_FecInicio.SetFocus
    End If
    
    Screen.MousePointer = 0
    
Exit Sub
Err_Eliminar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub cmd_grabar_Click()
On Error GoTo Err_Grabar

'    If vlSwVigPol = False Then
'       MsgBox "Póliza Ingresada No se Encuentra vigente para el Sistema", vbCritical, "Operación Cancelada"
'       Exit Sub
'    End If
        
    If Txt_PenPoliza = "" Then
        MsgBox "Debe ingresar el Nº de la Póliza.", vbInformation, "Error de Datos"
        Txt_PenPoliza.SetFocus
        Exit Sub
    End If
    
    If Cmb_PenNumIdent = "" Then
        MsgBox "Debe ingresar la Identificación del Pensionado/Beneficiario.", vbInformation, "Error de Datos"
        Cmb_PenNumIdent.SetFocus
        Exit Sub
    End If
    
    If Txt_PenNumIdent = "" Then
        MsgBox "Debe ingresar la Identificación del Pensionado/Beneficiario.", vbInformation, "Error de Datos"
        Txt_PenNumIdent.SetFocus
        Exit Sub
    End If
          
    If Fra_Poliza.Enabled = True Then
       Exit Sub
    End If
    
    If vlSwGrilla = False Then
       MsgBox " El Beneficiario Seleccionado No Tiene Derecho a Pensión ", vbCritical, "Operacion Cancelada"
       Exit Sub
    End If
      
    If (Trim(Cmb_HabDes) = "") Then
        MsgBox "Debe indicar el Tipo de Concepto de Haber/Descuento a ingresar.", vbCritical, "Error de Datos"
        Cmb_Moneda.SetFocus
        Exit Sub
    End If
    
    If (Trim(Cmb_Moneda) = "") Then
        MsgBox "Debe indicar la Moneda utilizada para el Haber/Descuento.", vbCritical, "Error de Datos"
        Cmb_Moneda.SetFocus
        Exit Sub
    End If
    
    If (Txt_FecInicio = "") Or (Not IsDate(Txt_FecInicio)) Then
        MsgBox "Debe Ingresar la Fecha de Inicio de Vigencia del Haber/Descuento.", vbInformation, "Error de Datos"
        Txt_FecInicio.SetFocus
        Exit Sub
    End If
        
     vlCodMSus = Trim(Mid(Cmb_MotSuspension, 1, (InStr(1, Cmb_MotSuspension, "-") - 1)))
     If Txt_FecSuspension <> "" Then
        If vlCodMSus <> "00" Then
           vlFechaSus = Trim(Txt_FecSuspension)
           If (flValidaFecha(vlFechaSus) = True) Then
               If CDate(Txt_FecSuspension) < CDate(Txt_FecInicio) Or _
                  CDate(Txt_FecSuspension) > CDate(Lbl_FecTermino) Then
                  MsgBox "Fecha de Suspensión Debe estar entre la  Fecha de Inicio y la de Término.", vbInformation, "Error de Datos"
                  Txt_FecSuspension.SetFocus
                  Exit Sub
               End If
           End If
        Else
           MsgBox "No puede Ingresar la Fecha de Suspensión, Código No Corresponde", vbInformation, "Error de Datos"
           Cmb_MotSuspension.SetFocus
           Exit Sub
        End If
     Else
      If vlCodMSus <> "00" Then
         MsgBox "Código No Corresponde, Falta Fecha de Suspensión", vbInformation, "Operación Cancelada"
         Cmd_Salir.SetFocus
         Exit Sub
      End If
     End If
    
    If Trim(Txt_FecSuspension) <> "" Then
             
       If fgValidaPagoPension(Trim(Txt_FecSuspension), Txt_PenPoliza, vlNumOrden) = False Then
          MsgBox "Ya se ha realizado el Proceso de Cálculo de Pensión para ésta fecha" & Chr(13) & _
                 "              Ingrese una Nueva Fecha           ", vbCritical, "Operación Cancelada"
          Screen.MousePointer = 0
          Txt_FecSuspension.SetFocus
          Exit Sub
       End If
    End If
  
    If Trim(Txt_FecSuspension) = "" Then
       vlSwVigPol = True
       If fgValidaVigenciaPoliza(Txt_PenPoliza, Trim(Txt_FecInicio)) = False Then
          MsgBox "La Fecha Ingresada no se Encuentra dentro del Rango de Vigencia de la Póliza" & Chr(13) & _
                 "O esta No Vigente. No se Ingresara Ni Modificara Información. ", vbCritical, "Operación Cancelada"
          Screen.MousePointer = 0
          vlSwVigPol = False
          Txt_FecInicio.SetFocus
          Exit Sub
       End If
 
       If fgValidaPagoPension(Trim(Txt_FecInicio), Txt_PenPoliza, vlNumOrden) = False Then
          MsgBox "Ya se ha relizado el Proceso de Cálculo de Pensión para ésta fecha" & Chr(13) & _
                 "                Ingrese una Nueva Fecha", vbCritical, "Operación Cancelada"
          Screen.MousePointer = 0
          If Txt_FecInicio.Enabled Then
            Txt_FecInicio.SetFocus
          End If
          Exit Sub
       End If
    End If
    
    If Txt_NroCuotas = "" Then
        MsgBox "Debe Ingresar el Nº de Cuotas del Haber/Descuento.", vbInformation, "Error de Datos"
        Txt_NroCuotas.SetFocus
        Exit Sub
    End If
    
    If Txt_MontoCuota = "" Then
        MsgBox "Debe Ingresar el Valor de las Cuotas a considerar.", vbInformation, "Error de Datos"
        Txt_MontoCuota.SetFocus
        Exit Sub
    Else
        If (Txt_MontoTotal.Enabled = False) Then
            If (CDbl(Txt_MontoCuota) > clTopeMaximo) Or (CDbl(Txt_MontoCuota) < clTopeMinimo) Then
                MsgBox "El Monto en Cuotas a registrar debe estar entre " & CStr(clTopeMinimo) & " y " & Format(clTopeMaximo, "#,#0.00"), vbCritical, "Error de Datos"
                Txt_MontoCuota.SetFocus
                Exit Sub
            End If
        Else
            If (CDbl(Txt_MontoCuota) > clTopeMaximoPorc) Or (CDbl(Txt_MontoCuota) < clTopeMinimoPorc) Then
                MsgBox "El Porcentaje a registrar debe estar entre " & CStr(clTopeMinimoPorc) & " y " & CStr(clTopeMaximoPorc), vbCritical, "Error de Datos"
                Txt_MontoCuota.SetFocus
                Exit Sub
            End If
        End If
    End If
    
    'Validar el Monto Total por el Concepto de Haber/Descuento
    If Txt_MontoTotal = "" Then
        MsgBox "Debe Ingresar el Monto Total del Haber/Descuento.", vbInformation, "Error de Datos"
        Txt_MontoCuota.SetFocus
        Exit Sub
    Else
        If (Txt_MontoTotal.Enabled = False) Then
            'If (CDbl(Txt_MontoTotal) > clTopeMaximo) Or (CDbl(Txt_MontoTotal) < clTopeMinimo) Then
            '    MsgBox "El Monto Total debe estar entre " & CStr(clTopeMinimo) & " y " & Format(clTopeMaximo, "#,#0.00"), vbCritical, "Error de Datos"
            '    Txt_MontoCuota.SetFocus
            '    Exit Sub
            'End If
        Else
            If (CDbl(Txt_MontoTotal) > clTopeMaximo) Or (CDbl(Txt_MontoTotal) < clTopeMinimo) Then
                MsgBox "El Monto Total a registrar debe estar entre " & CStr(clTopeMinimo) & " y " & Format(clTopeMaximo, "#,#0.00"), vbCritical, "Error de Datos"
                Txt_MontoTotal.SetFocus
                Exit Sub
            End If
        End If
    End If
    
    If (Trim(Cmb_MotSuspension) = "") Then
        MsgBox "Debe indicar si existe o no Motivo de Suspensión del Concepto.", vbCritical, "Error de Datos"
        Cmb_MotSuspension.SetFocus
        Exit Sub
    End If
 

       
'    vlCodMSus = Trim(Mid(Cmb_MotSuspension, 1, (InStr(1, Cmb_MotSuspension, "-") - 1)))
'    If Txt_FecSuspension <> "" Then
'       If vlCodMSus <> "00" Then
'          vlFechaSus = Trim(Txt_FecSuspension)
'          If (flValidaFecha(vlFechaSus) = True) Then
'              If CDate(Txt_FecSuspension) < CDate(Txt_FecInicio) Or _
'                 CDate(Txt_FecSuspension) > CDate(Lbl_FecTermino) Then
'
'                 MsgBox "Fecha de Suspensión tiene que estar Dentro del Rango.", vbInformation, "Error de Datos"
'                 Txt_FecSuspension.SetFocus
'                 Exit Sub
'              End If
'          End If
'       Else
'            MsgBox "No puede Ingresar la Fecha de Suspensión, Código No Corresponde", vbInformation, "Error de Datos"
'            Cmb_MotSuspension.SetFocus
'            Exit Sub
'       End If
'    End If
    
    
    If vlCodDerpen = "10" Then
      'No tiene Derecho
       MsgBox "El Pensionado No Tiene Derecho a Pensión, por ende no puede Ingresar/Actualizar los Conceptos de Haberes Y Descuentos.", vbInformation, "Información"
       Screen.MousePointer = 0
       Exit Sub
    End If
    
    Screen.MousePointer = 11
    
    vlOp = ""
    vlCodMon = Trim(Mid(Cmb_Moneda, 1, (InStr(1, Cmb_Moneda, "-") - 1)))
    vlCodHabDes = Trim(Mid(Cmb_HabDes, 1, (InStr(1, Cmb_HabDes, "-") - 1)))
    
    vlFecIni = Txt_FecInicio
    vlFecIni = Format(CDate(Trim(vlFecIni)), "yyyymmdd")
    vlFecTer = Lbl_FecTermino
    vlFecTer = Format(CDate(Trim(vlFecTer)), "yyyymmdd")
    
   'Verifica la existencia del Haber/Descuento
    vgSql = ""
    vgSql = "select num_poliza from PP_TMAE_HABDES where "
    vgSql = vgSql & " NUM_POLIZA = '" & Trim(Txt_PenPoliza) & "' and "
    vgSql = vgSql & " NUM_ORDEN = " & Trim(vlNumOrden) & " and "
    vgSql = vgSql & " COD_CONHABDES = '" & (vlCodHabDes) & "' and "
    vgSql = vgSql & " FEC_INIHABDES = '" & vlFecIni & "'"
    Set vlRegistro = vgConexionBD.Execute(vgSql)
    If vlRegistro.EOF Then
        vlOp = "I"
    Else
        vlOp = "A"
    End If
    vlRegistro.Close
    
    If (vlOp = "A") Then
        vgRes = MsgBox("¿ Está seguro que desea Modificar los Datos ?", 4 + 32 + 256, "Operación de Actualización")
        If vgRes <> 6 Then
            Cmd_Salir.SetFocus
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If
    If (vlOp = "I") Then
        vlResp = MsgBox(" ¿ Está seguro que desea Ingresar los Datos ?", 4 + 32 + 256, "Proceso de Ingreso de Datos")
        If vlResp <> 6 Then
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If
   'Si la accion es modificar
    If (vlOp = "A") Then
        Sql = "update PP_TMAE_HABDES set"
        Sql = Sql & " FEC_TERHABDES = '" & (vlFecTer) & "',"
        Sql = Sql & " NUM_ENDOSO = " & Trim(vlNumEndoso) & ","
        Sql = Sql & " NUM_CUOTAS = " & (Txt_NroCuotas) & ","
        Sql = Sql & " MTO_CUOTA = " & Str(Txt_MontoCuota) & ","
        Sql = Sql & " MTO_TOTAL = " & Str(Txt_MontoTotal) & ","
        Sql = Sql & " COD_MOTSUSHABDES = '" & (vlCodMSus) & "',"
        Sql = Sql & " COD_MONEDA = '" & (vlCodMon) & "',"
        If (Txt_FecSuspension) <> "" Then
            vlFecSus = Txt_FecSuspension
            vlFecSus = Format(CDate(Trim(vlFecSus)), "yyyymmdd")
            Sql = Sql & " FEC_SUSHABDES = '" & (vlFecSus) & "',"
        Else
            Sql = Sql & " FEC_SUSHABDES = NULL, "
        End If
        
        If Txt_Observacion <> "" Then
           Sql = Sql & " GLS_OBSHABDES = '" & (Txt_Observacion) & "',"
        Else
           Sql = Sql & " GLS_OBSHABDES = NULL, "
        End If
        Sql = Sql & " COD_USUARIOMODI = '" & (vgUsuario) & "',"
        Sql = Sql & " FEC_MODI = '" & Format(Date, "yyyymmdd") & "',"
        Sql = Sql & " HOR_MODI = '" & Format(Time, "hhmmss") & "'"
        Sql = Sql & " Where "
        Sql = Sql & " NUM_POLIZA = '" & Trim(Txt_PenPoliza) & "' and "
        Sql = Sql & " NUM_ORDEN = " & Trim(vlNumOrden) & " and "
        Sql = Sql & " COD_CONHABDES = '" & (vlCodHabDes) & "' and "
        Sql = Sql & " FEC_INIHABDES = '" & (vlFecIni) & "'"
        vgConexionBD.Execute (Sql)
            
    Else
       'Inserta los Datos en la Tabla PP_TMAE_HABDES
        Sql = ""
        Sql = "insert into PP_TMAE_HABDES ("
        Sql = Sql & "NUM_POLIZA,NUM_ORDEN,COD_CONHABDES,FEC_INIHABDES,"
        Sql = Sql & "FEC_TERHABDES,NUM_ENDOSO,NUM_CUOTAS,MTO_CUOTA,"
        Sql = Sql & "MTO_TOTAL,"
        Sql = Sql & "COD_MONEDA,COD_MOTSUSHABDES,"
        Sql = Sql & "FEC_SUSHABDES,GLS_OBSHABDES,COD_USUARIOCREA,FEC_CREA,HOR_CREA"
        Sql = Sql & " "
        Sql = Sql & ") values ("
        Sql = Sql & "'" & Trim(Txt_PenPoliza) & "',"
        Sql = Sql & " " & Trim(vlNumOrden) & ","
        Sql = Sql & "'" & (vlCodHabDes) & "',"
        Sql = Sql & "'" & (vlFecIni) & "',"
        Sql = Sql & "'" & (vlFecTer) & "',"
        Sql = Sql & " " & Trim(vlNumEndoso) & ","
        Sql = Sql & " " & Str(Txt_NroCuotas) & ","
        Sql = Sql & " " & Str(Txt_MontoCuota) & ","
        Sql = Sql & " " & Str(Txt_MontoTotal) & ","
        Sql = Sql & "'" & (vlCodMon) & "',"
        Sql = Sql & "'" & (vlCodMSus) & "',"
                
        If (Txt_FecSuspension) <> "" Then
            vlFecSus = Txt_FecSuspension
            vlFecSus = Format(CDate(Trim(vlFecSus)), "yyyymmdd")
            Sql = Sql & "'" & (vlFecSus) & "',"
        Else
            Sql = Sql & "NULL ,"
        End If
                
        If Trim(Txt_Observacion) <> "" Then
            Sql = Sql & "'" & (Txt_Observacion) & "',"
        Else
            Sql = Sql & "NULL ,"
        End If
        
        Sql = Sql & "'" & (vgUsuario) & "',"
        Sql = Sql & "'" & Format(Date, "yyyymmdd") & "',"
        Sql = Sql & "'" & Format(Time, "hhmmss") & "'"
        Sql = Sql & ")"
        vgConexionBD.Execute (Sql)
    End If
     
    If (vlOp <> "") Then
        'Limpia los Datos de la Pantalla
       ' Cmd_Limpiar_Click
        flLmpGrilla
        flBuscarHistorico
    End If
    
    Screen.MousePointer = 0
        
Exit Sub
Err_Grabar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(2) & Err.Description & " ]"
    End Select
End Sub

Private Sub Cmd_Imprimir_Click()
On Error GoTo Err_Imprimir
    'Valida que se haya seleccionado al Pensionado
    If Txt_PenPoliza = "" Then
        MsgBox "Debe ingresar el Nº de la Póliza.", vbInformation, "Error de Datos"
        Exit Sub
    End If
    If Cmb_PenNumIdent = "" Then
        MsgBox "Debe ingresar la Identificación del Pensionado/Beneficiario.", vbInformation, "Error de Datos"
        Cmb_PenNumIdent.SetFocus
        Exit Sub
    End If
    If Txt_PenNumIdent = "" Then
        MsgBox "Debe ingresar la Identificación del Pensionado/Beneficiario.", vbInformation, "Error de Datos"
        Txt_PenNumIdent.SetFocus
        Exit Sub
    End If
    
    If Fra_Poliza.Enabled = True Then
       Exit Sub
    End If
 
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

Private Sub Cmd_Limpiar_Click()
On Error GoTo Err_Limpiar

  If SSTab1.Tab = 0 Then
     If Fra_Poliza.Enabled = False Then
        flLmpHabDes
'       flDesHabDes
        Cmb_HabDes.Enabled = True
        Txt_FecInicio.Enabled = True
        Txt_FecInicio.SetFocus
     End If
  End If
  
Exit Sub
Err_Limpiar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

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

    Frm_PensHabDescto.Top = 0
    Frm_PensHabDescto.Left = 0
    vlModFecTerNumCuotas = True
    SSTab1.Tab = 0
    
    SSTab1.Enabled = False
        
    'Carga Tipo de Identificación
    fgComboTipoIdentificacion Cmb_PenNumIdent
    
   'carga Código de Moneda
    fgComboGeneral vgCodTabla_TipMon, Cmb_Moneda
    
   'carga Código de Suspensión
    fgComboGeneral vgCodTabla_SitHabDes, Cmb_MotSuspension
    
   'carga Código de Haberes/Descuentos
    flCargaCodHabDesc Cmb_HabDes
    
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Function flCargaCodHabDesc(iCombo As ComboBox)
On Error GoTo Err_ComboHabDes

    vgSql = ""
    vgSql = "SELECT cod_conhabdes,gls_conhabdes,cod_tipmov "
    vgSql = vgSql & "FROM ma_tpar_conhabdes WHERE "
    vgSql = vgSql & "cod_modorigen <> 'GE' AND "
    vgSql = vgSql & "cod_modorigen <> 'CCAF' AND "
    vgSql = vgSql & "cod_interno = 'N' ORDER BY cod_conhabdes"
    Set vlRegistro = vgConexionBD.Execute(vgSql)
    Do While Not vlRegistro.EOF
        iCombo.AddItem ((Trim(vlRegistro!Cod_ConHabDes) & " - " & Trim(vlRegistro!gls_ConHabDes)) & "   (" & Trim(vlRegistro!cod_tipmov) & ")")
        vlRegistro.MoveNext
    Loop
    vlRegistro.Close
    
    If iCombo.ListCount <> 0 Then
        iCombo.ListIndex = 0
    End If

Exit Function
Err_ComboHabDes:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Private Sub Msf_Grilla_Click()
On Error GoTo Err_Grilla

    Msf_Grilla.Col = 0
    vlPos = Msf_Grilla.RowSel
    Msf_Grilla.Row = vlPos
    If (Msf_Grilla.Text = "") Or (Msf_Grilla.Row = 0) Then
        Exit Sub
    End If
    Screen.MousePointer = 11
    
'    flHabHabDes
    
    Msf_Grilla.Col = 10
    If Msf_Grilla.Text = "S" Then
       Screen.MousePointer = 0
       Exit Sub
    End If
    
    Msf_Grilla.Col = 11
    If Msf_Grilla.Text = "CCAF" Or Msf_Grilla.Text = "GE" Then
       Screen.MousePointer = 0
       Exit Sub
    End If
    
    Msf_Grilla.Col = 0
    
    vlCodHabDes = Trim(Msf_Grilla.Text)
    If Cmb_HabDes.Text <> "" Then
       For vlI = 0 To Cmb_HabDes.ListCount - 1
           Cmb_HabDes.ListIndex = vlI
           If Cmb_HabDes.Text = vlCodHabDes Then
              Exit For
           End If
       Next vlI
    End If
            
    Msf_Grilla.Col = 1
    Txt_FecInicio = (Msf_Grilla.Text)
    
    If fgValidaPagoPension(Trim(Txt_FecInicio), Txt_PenPoliza, vlNumOrden) = False Then
       Fra_Suspension.Enabled = True
       vlSwSuspension = True
    Else
       Fra_Suspension.Enabled = False
       vlSwSuspension = False
    End If
       
    Msf_Grilla.Col = 2
    Lbl_FecTermino = (Msf_Grilla.Text)
    
    Msf_Grilla.Col = 3
    Txt_NroCuotas = (Msf_Grilla.Text)
    
    Msf_Grilla.Col = 4
    Txt_MontoCuota = (Msf_Grilla.Text)
    
    'Calculo del Monto Total del Haber o Descuento
    Msf_Grilla.Col = 5
    Txt_MontoTotal = (Msf_Grilla.Text)
    
    'Determinar la Moneda utilizada
    Msf_Grilla.Col = 6
    vlCodMon = Trim(Msf_Grilla.Text)
    vlModFecTerNumCuotas = False
    If Cmb_Moneda.Text <> "" Then
       For vlI = 0 To Cmb_Moneda.ListCount - 1
           Cmb_Moneda.ListIndex = vlI
           vgPalabra = Trim(Mid(Cmb_Moneda, 1, (InStr(1, Cmb_Moneda, "-") - 1)))
           'If Cmb_Moneda.Text = vlCodMon Then
           If vgPalabra = vlCodMon Then
              Exit For
           End If
       Next vlI
    End If
    vlModFecTerNumCuotas = True
    Msf_Grilla.Col = 7
    vlCodMSus = (Msf_Grilla.Text)
    If Cmb_MotSuspension.Text <> "" Then
       For vlI = 0 To Cmb_MotSuspension.ListCount - 1
           Cmb_MotSuspension.ListIndex = vlI
           If Cmb_MotSuspension.Text = vlCodMSus Then
              Exit For
           End If
       Next vlI
    End If
    
'    vlCodMSus = Trim(Msf_Grilla.Text)
'    If vlCodMSus <> "" Then
'       If Cmb_MotSuspension.Text <> "" Then
'          For vlI = 0 To Cmb_MotSuspension.ListCount - 1
'              Cmb_MotSuspension.ListIndex = vlI
'              vlCodSus = Trim(Mid(Cmb_MotSuspension, 1, (InStr(1, Cmb_MotSuspension, "-") - 1)))
'              If vlCodSus = vlCodMSus Then
'                 Exit For
'              End If
'          Next vlI
'       End If
'    Else
'     Cmb_MotSuspension.ListIndex = 0
'    End If
    
    
    
    
    
    
    
    Msf_Grilla.Col = 8
    Txt_FecSuspension = (Msf_Grilla.Text)
    
    Msf_Grilla.Col = 9
    Txt_Observacion = (Msf_Grilla.Text)
    
    
    ''Calculo del Monto Total del Haber o Descuento
    'Txt_MontoTotal = CDbl(Txt_NroCuotas) * CDbl(Txt_MontoCuota)
    'Txt_MontoTotal = Format(Txt_MontoTotal, "#,#0.00")
    
    Cmb_HabDes.Enabled = False
    Txt_FecInicio.Enabled = False
    Cmb_Moneda.SetFocus
    
    Screen.MousePointer = 0
    SSTab1.Tab = 0
    
Exit Sub
Err_Grilla:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Txt_FecInicio_GotFocus()
    vlPasa = False
    Txt_FecInicio.SelStart = 0
    Txt_FecInicio.SelLength = Len(Txt_FecInicio)
End Sub

Private Sub Txt_FecInicio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If (Trim(Txt_FecInicio <> "")) Then
       vlPasa = True
'       flHabHabDes
       Call flHaberDescuento
      
    End If
End If
End Sub

Private Sub Txt_FecInicio_LostFocus()
    If Txt_FecInicio <> "" Then
       If vlPasa = False Then
'          flHabHabDes
          Call flHaberDescuento
       End If
       Txt_FecInicio = Format(CDate(Trim(Txt_FecInicio)), "yyyymmdd")
       Txt_FecInicio = DateSerial(Mid((Txt_FecInicio), 1, 4), Mid((Txt_FecInicio), 5, 2), Mid((Txt_FecInicio), 7, 2))
       
    End If
End Sub

Private Sub Txt_FecSuspension_GotFocus()
    Txt_FecSuspension.SelStart = 0
    Txt_FecSuspension.SelLength = Len(Txt_FecSuspension)
End Sub

Private Sub Txt_FecSuspension_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Txt_Observacion.SetFocus
    End If
End Sub

Private Sub Txt_FecSuspension_LostFocus()
    If Txt_FecSuspension = "" Then
       Exit Sub
    End If
    If Not IsDate(Txt_FecSuspension) Then
       Txt_FecSuspension = ""
       Exit Sub
    End If
    If Txt_FecSuspension <> "" Then
       Txt_FecSuspension = Format(CDate(Trim(Txt_FecSuspension)), "yyyymmdd")
       Txt_FecSuspension = DateSerial(Mid((Txt_FecSuspension), 1, 4), Mid((Txt_FecSuspension), 5, 2), Mid((Txt_FecSuspension), 7, 2))
    End If
End Sub

Private Sub Txt_MontoCuota_Change()
If Not IsNumeric(Txt_MontoCuota) Then
    Txt_MontoCuota = ""
End If
End Sub

Private Sub Txt_MontoCuota_GotFocus()
    Txt_MontoCuota.SelStart = 0
    Txt_MontoCuota.SelLength = Len(Txt_MontoCuota)
End Sub

Private Sub Txt_MontoCuota_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If IsNumeric(Txt_MontoCuota) Then
        Txt_MontoCuota = Format(Txt_MontoCuota, "#,#0.00")
        If (Txt_MontoTotal.Enabled = False) Then
            If (CDbl(Txt_MontoCuota) > clTopeMaximo) Or (CDbl(Txt_MontoCuota) < clTopeMinimo) Then
                MsgBox "El Monto en Cuotas a registrar debe estar entre " & CStr(clTopeMinimo) & " y " & Format(clTopeMaximo, "#,#0.00"), vbCritical, "Error de Datos"
                Exit Sub
            End If
            If vlSwSuspension = True Then
               Fra_Suspension.Enabled = True
               Cmb_MotSuspension.SetFocus
            Else
               Fra_Suspension.Enabled = False
               Cmd_Grabar.SetFocus
            End If
        Else
            If CDbl(Txt_MontoCuota) > clTopeMaximoPorc Or CDbl(Txt_MontoCuota) < clTopeMinimoPorc Then
                MsgBox "El Porcentaje a registrar debe estar entre " & CStr(clTopeMinimoPorc) & " y " & CStr(clTopeMaximoPorc), vbCritical, "Error de Datos"
                Exit Sub
            End If
            Txt_MontoTotal.SetFocus
        End If
    End If
End If
End Sub

Private Sub Txt_MontoCuota_LostFocus()
If IsNumeric(Txt_MontoCuota) Then
    Txt_MontoCuota = Format(Txt_MontoCuota, "#,#0.00")
    If (Txt_MontoTotal.Enabled = False) Then
        If (CDbl(Txt_MontoCuota) > clTopeMaximo) Or (CDbl(Txt_MontoCuota) < clTopeMinimo) Then
            MsgBox "El Monto en Cuotas a registrar debe estar entre " & CStr(clTopeMinimo) & " y " & Format(clTopeMaximo, "#,#0.00"), vbCritical, "Error de Datos"
            Exit Sub
        End If
        If (Txt_MontoCuota <> "") And (Txt_NroCuotas <> "") Then
            Txt_MontoTotal = CDbl(Txt_MontoCuota) * CDbl(Txt_NroCuotas)
            Txt_MontoTotal = Format(Txt_MontoTotal, "#,#0.00")
        End If
    Else
        If CDbl(Txt_MontoCuota) > clTopeMaximoPorc Or CDbl(Txt_MontoCuota) < clTopeMinimoPorc Then
            MsgBox "El Porcentaje de Renta debe ser un valor comprendido entre 0% y 100%.", vbExclamation, "Error de Datos"
        End If
    End If
End If
End Sub

Private Sub Txt_MontoTotal_Change()
If Not IsNumeric(Txt_MontoTotal) Then
    Txt_MontoTotal = ""
End If
End Sub

Private Sub Txt_MontoTotal_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    If IsNumeric(Txt_MontoTotal) Then
        Txt_MontoTotal = Format(Txt_MontoTotal, "#,#0.00")
        If (CDbl(Txt_MontoTotal) > clTopeMaximo) Or (CDbl(Txt_MontoTotal) < clTopeMinimo) Then
            MsgBox "El Monto Total a registrar debe estar entre " & CStr(clTopeMinimo) & " y " & Format(clTopeMaximo, "#,#0.00"), vbCritical, "Error de Datos"
            Exit Sub
        End If
        If vlSwSuspension = True Then
           Fra_Suspension.Enabled = True
           Cmb_MotSuspension.SetFocus
        Else
          Fra_Suspension.Enabled = False
          Cmd_Grabar.SetFocus
        End If
    End If
End If
End Sub

Private Sub Txt_MontoTotal_LostFocus()
If IsNumeric(Txt_MontoTotal) Then
    Txt_MontoTotal = Format(Txt_MontoTotal, "#,#0.00")
    If (CDbl(Txt_MontoTotal) > clTopeMaximo) Or (CDbl(Txt_MontoTotal) < clTopeMinimo) Then
        MsgBox "El Monto Total a registrar debe estar entre " & CStr(clTopeMinimo) & " y " & Format(clTopeMaximo, "#,#0.00"), vbCritical, "Error de Datos"
    End If
End If
End Sub

Private Sub Txt_NroCuotas_Change()
If Not IsNumeric(Txt_NroCuotas) Then
    Txt_NroCuotas = ""
End If
End Sub

Private Sub Txt_NroCuotas_GotFocus()
    Txt_NroCuotas.SelStart = 0
    Txt_NroCuotas.SelLength = Len(Txt_NroCuotas)
End Sub

Private Sub Txt_NroCuotas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If IsNumeric(Txt_NroCuotas) Then
       Txt_MontoCuota.SetFocus
    End If
End If
End Sub

Private Sub Txt_NroCuotas_LostFocus()
    If IsNumeric(Txt_NroCuotas) Then
       flCalculaFecha
    End If
    If (Txt_MontoTotal.Enabled = False) Then
        If (Txt_MontoCuota <> "") And (Txt_NroCuotas <> "") Then
            Txt_MontoTotal = CDbl(Txt_MontoCuota) * CDbl(Txt_NroCuotas)
            Txt_MontoTotal = Format(Txt_MontoTotal, "#,#0.00")
        End If
    End If
End Sub

Private Sub Txt_Observacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Txt_Observacion = Trim(UCase(Txt_Observacion))
       KeyAscii = 0
       Cmd_Grabar.SetFocus
    End If
End Sub

Private Sub Txt_Observacion_LostFocus()
    Txt_Observacion = Trim(UCase(Txt_Observacion))
End Sub

Private Sub Txt_PenNumIdent_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If (Trim(Txt_PenNumIdent) <> "") Then
            Txt_PenNumIdent = UCase(Trim(Txt_PenNumIdent))
        End If
        Cmd_BuscarPol.SetFocus
    End If
End Sub

Private Sub txt_pennumident_lostfocus()
    Txt_PenNumIdent = Trim(UCase(Txt_PenNumIdent))
End Sub

Private Sub Txt_PenPoliza_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(Txt_PenPoliza) <> "" Then
        Txt_PenPoliza = UCase(Txt_PenPoliza)
        Txt_PenPoliza = Format(Txt_PenPoliza, "0000000000")
        Cmb_PenNumIdent.SetFocus
      Else
         Cmb_PenNumIdent.SetFocus
      End If
   End If
End Sub

Private Sub Txt_PenPoliza_LostFocus()
    Txt_PenPoliza = Trim(UCase(Txt_PenPoliza))
    Txt_PenPoliza = Format(Txt_PenPoliza, "0000000000")
End Sub

VERSION 5.00
Begin VB.Form Frm_SisNivel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenedor de Niveles de Acceso."
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   Icon            =   "Frm_SisNivel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   10200
   Begin VB.CheckBox Chk_Validaciones 
      Caption         =   "Validaciones"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6600
      TabIndex        =   37
      Top             =   2520
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CheckBox Chk_InfoLiqPago 
      Caption         =   "Liquidación de Pago"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6600
      TabIndex        =   36
      Top             =   2760
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox Chk_InfoSalud 
      Caption         =   "Plan de Salud"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6600
      TabIndex        =   35
      Top             =   3000
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox Chk_InfoContables 
      Caption         =   "Informes Contables"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6600
      TabIndex        =   34
      Top             =   3240
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox Chk_InfoSVS 
      Caption         =   "Informes SBS"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6600
      TabIndex        =   33
      Top             =   3480
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox Chk_InfoCierre 
      Caption         =   "Cierre de Pensiones"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6600
      TabIndex        =   32
      Top             =   3720
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox Chk_CalInformes 
      Caption         =   "Emisión de Informes"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6360
      TabIndex        =   31
      Top             =   2280
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.CheckBox Chk_PriProvisorio 
      Caption         =   "Provisorio"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6600
      TabIndex        =   30
      Top             =   1800
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox Chk_PriDefinitivo 
      Caption         =   "Definitivo"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6600
      TabIndex        =   29
      Top             =   2040
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox Chk_CalParametros 
      Caption         =   "Parámetros de Cálculo de Pagos Recurrentes"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6360
      TabIndex        =   28
      Top             =   1320
      Value           =   1  'Checked
      Width           =   3615
   End
   Begin VB.CheckBox Chk_CalPrimeras 
      Caption         =   "Pensiones Recurrentes"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6360
      TabIndex        =   27
      Top             =   1560
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.CheckBox Chk_ManEndosos 
      Caption         =   "Endosos"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   480
      TabIndex        =   25
      Top             =   6000
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.Frame Fra_Nivel 
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
      TabIndex        =   11
      Top             =   0
      Width           =   9975
      Begin VB.TextBox Txt_Nivel 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   5520
         MaxLength       =   50
         TabIndex        =   1
         Top             =   240
         Width           =   4095
      End
      Begin VB.ComboBox Cmb_Nivel 
         Height          =   315
         ItemData        =   "Frm_SisNivel.frx":0442
         Left            =   2520
         List            =   "Frm_SisNivel.frx":0444
         TabIndex        =   0
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Descripción Nivel :"
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
         Left            =   3840
         TabIndex        =   48
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Nivel de Acceso :"
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
         Left            =   840
         TabIndex        =   12
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Fra_Menu 
      Caption         =   "  Menú de Acceso  "
      ForeColor       =   &H00800000&
      Height          =   6315
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   9975
      Begin VB.CheckBox Chk_PagRecurrentes 
         Caption         =   "Pagos Recurrentes"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6480
         TabIndex        =   56
         Top             =   4320
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox Chk_GastosSepelio 
         Caption         =   "Gastos de Sepelio"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6480
         TabIndex        =   55
         Top             =   4560
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox Chk_PeriodoGar 
         Caption         =   "Periodo Garantizado"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6480
         TabIndex        =   54
         Top             =   4800
         Value           =   1  'Checked
         Width           =   3255
      End
      Begin VB.CheckBox Chk_ArchContables 
         Caption         =   "Archivos Contables"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   6240
         TabIndex        =   53
         Top             =   4080
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CheckBox Chk_ConPagosTer 
         Caption         =   "Consulta de Pagos a Terceros"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6480
         TabIndex        =   52
         Top             =   3840
         Value           =   1  'Checked
         Width           =   3255
      End
      Begin VB.CheckBox Chk_PerGarantizado 
         Caption         =   "Periodo Garantizado"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6480
         TabIndex        =   51
         Top             =   3600
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox chk_GastoSepelio 
         Caption         =   "Gastos de Sepelio"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6480
         TabIndex        =   50
         Top             =   3360
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox Chk_CertSuperv 
         Caption         =   "Certificado de Supervivencia"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   49
         Top             =   2280
         Value           =   1  'Checked
         Width           =   3135
      End
      Begin VB.CheckBox Chk_HabDesAuto 
         Caption         =   "Carga Automática"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   47
         Top             =   3960
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox Chk_HabDesMan 
         Caption         =   "Ingreso Manual"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   46
         Top             =   3720
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox Chk_CalApeManual 
         Caption         =   "Habilitar Reproceso de pagos Recurrentes"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   6240
         TabIndex        =   45
         Top             =   5280
         Value           =   1  'Checked
         Width           =   3615
      End
      Begin VB.CheckBox Chk_AntArchBen 
         Caption         =   "Generación Archivo Datos Beneficiarios"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   44
         Top             =   2520
         Value           =   1  'Checked
         Width           =   3135
      End
      Begin VB.CheckBox Chk_ConConsulta 
         Caption         =   "Consulta General por Pensionado"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   6240
         TabIndex        =   43
         Top             =   5760
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox Chk_ConsultaEndosos 
         Caption         =   "Consulta de Endosos"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   42
         Top             =   5640
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox Chk_GenerarEndosos 
         Caption         =   "Generar Endosos"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   41
         Top             =   5400
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox Chk_CalCalendario 
         Caption         =   "Calendario de Pagos Recurrentes"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   6240
         TabIndex        =   40
         Top             =   5040
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox Chk_CalPagosTer 
         Caption         =   "Registro de Pagos a Terceros"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   6240
         TabIndex        =   39
         Top             =   3120
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CheckBox Chk_Consultas 
         Caption         =   "Consultas"
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
         Left            =   6000
         TabIndex        =   38
         Top             =   5520
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox Chk_CalPension 
         Caption         =   "Generación Cálculo de Pensión"
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
         Left            =   6000
         TabIndex        =   26
         Top             =   240
         Value           =   1  'Checked
         Width           =   3135
      End
      Begin VB.CheckBox Chk_MenAutomaticos 
         Caption         =   "Automáticos"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   24
         Top             =   4680
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox Chk_MenCarga 
         Caption         =   "Carga desde Archivo"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   23
         Top             =   4920
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox Chk_RetInfoControl 
         Caption         =   "Informe de Control"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   22
         Top             =   3240
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox Chk_RetOrdenJud 
         Caption         =   "Ingreso de Orden Judicial"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   21
         Top             =   3000
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox Chk_AntGrales 
         Caption         =   "Mantención Antecedentes Generales"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   20
         Top             =   1800
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.CheckBox Chk_AntTutores 
         Caption         =   "Asignación de Tutor/Apoderado"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   19
         Top             =   2040
         Value           =   1  'Checked
         Width           =   3975
      End
      Begin VB.CheckBox Chk_MenIndividuales 
         Caption         =   "Individuales (por pensionado)"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   18
         Top             =   4440
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CheckBox Chk_ManMensajes 
         Caption         =   "Generación de Mensajes"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   4200
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox Chk_ManHabDesc 
         Caption         =   "Ingreso de Haberes y Descuentos"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   3480
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.CheckBox Chk_ManRetJud 
         Caption         =   "Retención Judicial"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   2760
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox Chk_ManAntecedentes 
         Caption         =   "Antecedentes Pensionado"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   1560
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox Chk_ManInformacion 
         Caption         =   "Mantención de Información"
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
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox Chk_Niveles 
         Caption         =   "Nivel de Acceso"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   960
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox Chk_Contrasena 
         Caption         =   "Contraseña"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   720
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox Chk_Usuarios 
         Caption         =   "Usuarios "
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox Chk_Sistema 
         Caption         =   "Adm. del Sistema"
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
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   1  'Checked
         Width           =   2175
      End
   End
   Begin VB.Frame Fra_Operacion 
      Height          =   950
      Left            =   120
      TabIndex        =   9
      Top             =   7200
      Width           =   9975
      Begin VB.CommandButton cmd_salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5760
         Picture         =   "Frm_SisNivel.frx":0446
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   170
         Width           =   720
      End
      Begin VB.CommandButton cmd_limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   4620
         Picture         =   "Frm_SisNivel.frx":0540
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   170
         Width           =   720
      End
      Begin VB.CommandButton cmd_grabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   3420
         Picture         =   "Frm_SisNivel.frx":0BFA
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   170
         Width           =   720
      End
   End
End
Attribute VB_Name = "Frm_SisNivel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vlVerificar       As String
Dim vlOperacion       As String
Dim vlNivel           As Integer
Dim vlSw              As Boolean

Dim vlGlsUsuarioCrea As Variant
Dim vlFecCrea As Variant
Dim vlHorCrea As Variant
Dim vlGlsUsuarioModi As Variant
Dim vlFecModi As Variant
Dim vlHorModi As Variant


Function flRegistrarNivel()
On Error GoTo Err_Registrar
    Dim i As Integer
    Dim vlNivelAgregado As Long
    
    Screen.MousePointer = 11
    vlNivel = CLng(Cmb_Nivel)
    vlNivelAgregado = CLng(Cmb_Nivel)
    'Verificar existencia de Código Nivel para el Ingreso/Actualización
    vgQuery = "SELECT cod_nivel FROM MA_TPAR_NIVEL WHERE "
    vgQuery = vgQuery & "cod_sistema = '" & vgTipoSistema & "' AND "
    vgQuery = vgQuery & "cod_nivel = " & vlNivel & ""
    Set vgRs = vgConexionBD.Execute(vgQuery)
    If (vgRs.EOF) Then
        vlOperacion = "I"
    Else
        vlOperacion = "A"
    End If
    vgRs.Close
    
    If (vlOperacion = "I") Then
        'Ingresar
        vlGlsUsuarioCrea = vgUsuario
        vlFecCrea = Format(Date, "yyyymmdd")
        vlHorCrea = Format(Time, "hhmmss")
        vgQuery = ""
        
        vgQuery = "INSERT INTO MA_TPAR_NIVEL ("
        vgQuery = vgQuery & "cod_sistema,"
        vgQuery = vgQuery & "cod_nivel,"
        vgQuery = vgQuery & "gls_nivel,"
        vgQuery = vgQuery & "num_menu_1,"
        vgQuery = vgQuery & "num_menu_1_1,"
        vgQuery = vgQuery & "num_menu_1_2,"
        vgQuery = vgQuery & "num_menu_1_3,"
    
        vgQuery = vgQuery & "num_menu_2,"
        vgQuery = vgQuery & "num_menu_2_1,"
        vgQuery = vgQuery & "num_menu_2_1_1,"
        vgQuery = vgQuery & "num_menu_2_1_2,"
        vgQuery = vgQuery & "num_menu_2_1_3,"
        vgQuery = vgQuery & "num_menu_2_1_4,"
        
        vgQuery = vgQuery & "num_menu_2_2,"
        vgQuery = vgQuery & "num_menu_2_2_1,"
        vgQuery = vgQuery & "num_menu_2_2_2,"
        
        vgQuery = vgQuery & "num_menu_2_3,"
        vgQuery = vgQuery & "num_menu_2_3_1,"
        vgQuery = vgQuery & "num_menu_2_3_2,"
        
        vgQuery = vgQuery & "num_menu_2_4,"
        vgQuery = vgQuery & "num_menu_2_4_1,"
        vgQuery = vgQuery & "num_menu_2_4_2,"
        vgQuery = vgQuery & "num_menu_2_4_3,"
       
        vgQuery = vgQuery & "num_menu_2_5,"
        vgQuery = vgQuery & "num_menu_2_5_1,"
        vgQuery = vgQuery & "num_menu_2_5_2,"
    
        vgQuery = vgQuery & "num_menu_3,"
        vgQuery = vgQuery & "num_menu_3_1,"
        vgQuery = vgQuery & "num_menu_3_2,"
        vgQuery = vgQuery & "num_menu_3_2_1,"
        vgQuery = vgQuery & "num_menu_3_2_2,"
        vgQuery = vgQuery & "num_menu_3_3,"
        vgQuery = vgQuery & "num_menu_3_3_1,"
        vgQuery = vgQuery & "num_menu_3_3_2,"
        vgQuery = vgQuery & "num_menu_3_3_3,"
        vgQuery = vgQuery & "num_menu_3_3_4,"
        vgQuery = vgQuery & "num_menu_3_3_5,"
        vgQuery = vgQuery & "num_menu_3_3_6,"
        
        vgQuery = vgQuery & "num_menu_3_4,"
        vgQuery = vgQuery & "num_menu_3_4_1,"
        vgQuery = vgQuery & "num_menu_3_4_2,"
        vgQuery = vgQuery & "num_menu_3_4_3,"
       
        vgQuery = vgQuery & "num_menu_3_5,"
        vgQuery = vgQuery & "num_menu_3_5_1,"
        vgQuery = vgQuery & "num_menu_3_5_2,"
        vgQuery = vgQuery & "num_menu_3_5_3,"
        
        vgQuery = vgQuery & "num_menu_3_6,"
        vgQuery = vgQuery & "num_menu_3_7,"
        
        vgQuery = vgQuery & "num_menu_4,"
        vgQuery = vgQuery & "num_menu_4_1,"
        vgQuery = vgQuery & "cod_usuariocrea,"
        vgQuery = vgQuery & "fec_crea,"
        vgQuery = vgQuery & "hor_crea "
        
        vgQuery = vgQuery & ") VALUES ("
        vgQuery = vgQuery & "'" & vgTipoSistema & "',"
        vgQuery = vgQuery & " " & vlNivel & ", "
        vgQuery = vgQuery & "'" & Trim(Txt_Nivel) & "', "
        vgQuery = vgQuery & " " & Chk_Sistema.Value & ", "
        vgQuery = vgQuery & " " & Chk_Usuarios.Value & ", "
        vgQuery = vgQuery & " " & Chk_Contrasena.Value & ", "
        vgQuery = vgQuery & " " & Chk_Niveles.Value & ", "
        
        vgQuery = vgQuery & " " & Chk_ManInformacion.Value & ", "
        vgQuery = vgQuery & " " & Chk_ManAntecedentes.Value & ", "
        vgQuery = vgQuery & " " & Chk_AntGrales.Value & ", "
        vgQuery = vgQuery & " " & Chk_AntTutores.Value & ", "
         vgQuery = vgQuery & " " & Chk_CertSuperv.Value & ", "
        vgQuery = vgQuery & " " & Chk_AntArchBen.Value & ", "
        
        vgQuery = vgQuery & " " & Chk_ManRetJud.Value & ", "
        vgQuery = vgQuery & " " & Chk_RetOrdenJud.Value & ", "
        vgQuery = vgQuery & " " & Chk_RetInfoControl.Value & ", "
    
        vgQuery = vgQuery & " " & Chk_ManHabDesc.Value & ", "
        vgQuery = vgQuery & " " & Chk_HabDesMan.Value & ", "
        vgQuery = vgQuery & " " & Chk_HabDesAuto.Value & ", "
        
        vgQuery = vgQuery & " " & Chk_ManMensajes.Value & ", "
        vgQuery = vgQuery & " " & Chk_MenIndividuales.Value & ", "
        vgQuery = vgQuery & " " & Chk_MenAutomaticos.Value & ", "
        vgQuery = vgQuery & " " & Chk_MenCarga.Value & ", "
        
        vgQuery = vgQuery & " " & Chk_ManEndosos.Value & ", "
        vgQuery = vgQuery & " " & Chk_GenerarEndosos.Value & ", "
        vgQuery = vgQuery & " " & Chk_ConsultaEndosos.Value & ", "
        'nivel 3
        vgQuery = vgQuery & " " & Chk_CalPension.Value & ", "
        vgQuery = vgQuery & " " & Chk_CalParametros.Value & ", "
        vgQuery = vgQuery & " " & Chk_CalPrimeras.Value & ", "
        vgQuery = vgQuery & " " & Chk_PriProvisorio.Value & ", "
        vgQuery = vgQuery & " " & Chk_PriDefinitivo.Value & ", "
        
        vgQuery = vgQuery & " " & Chk_CalInformes.Value & ", "
        vgQuery = vgQuery & " " & Chk_Validaciones.Value & ", "
        vgQuery = vgQuery & " " & Chk_InfoLiqPago.Value & ", "
        vgQuery = vgQuery & " " & Chk_InfoSalud.Value & ", "
        vgQuery = vgQuery & " " & Chk_InfoContables.Value & ", "
        vgQuery = vgQuery & " " & Chk_InfoSVS.Value & ", "
        vgQuery = vgQuery & " " & Chk_InfoCierre.Value & ", "
        
        vgQuery = vgQuery & " " & Chk_CalPagosTer.Value & ", "
        vgQuery = vgQuery & " " & chk_GastoSepelio.Value & ", "
        vgQuery = vgQuery & " " & Chk_PerGarantizado.Value & ", "
        vgQuery = vgQuery & " " & Chk_ConPagosTer.Value & ", "
        
        vgQuery = vgQuery & " " & Chk_ArchContables.Value & ", "
        vgQuery = vgQuery & " " & Chk_PagRecurrentes.Value & ", "
        vgQuery = vgQuery & " " & Chk_GastosSepelio.Value & ", "
        vgQuery = vgQuery & " " & Chk_PeriodoGar.Value & ", "
              
        vgQuery = vgQuery & " " & Chk_CalCalendario.Value & ", "
        vgQuery = vgQuery & " " & Chk_CalApeManual.Value & ", "
        
        vgQuery = vgQuery & " " & Chk_Consultas.Value & ", "
        vgQuery = vgQuery & " " & Chk_ConConsulta.Value & ", "

        vgQuery = vgQuery & " '" & vlGlsUsuarioCrea & "', "
        vgQuery = vgQuery & " '" & vlFecCrea & "', "
        vgQuery = vgQuery & " '" & vlHorCrea & "' "
        vgQuery = vgQuery & " " & ") "
        vgConexionBD.Execute (vgQuery)
        
        MsgBox "El registro de Datos fue realizado Correctamente", vbInformation, "Información"
        'Call Cmd_Limpiar_Click '05/11/2005 Comentado por sugerencia de AVB
    Else
        If (vlOperacion = "A") Then
            'Actualizar
            vgRes = MsgBox("¿ Está seguro que desea Modificar los Niveles de Acceso ?", 4 + 32 + 256, "Operación de Actualización")
            If vgRes <> 6 Then
                Screen.MousePointer = 0
                Exit Function
            End If
            
            vlGlsUsuarioModi = vgUsuario
            vlFecModi = Format(Date, "yyyymmdd")
            vlHorModi = Format(Time, "hhmmss")
            
            vgQuery = "UPDATE MA_TPAR_NIVEL SET "
            vgQuery = vgQuery & "gls_nivel ='" & Trim(Txt_Nivel) & "', "
            vgQuery = vgQuery & "num_menu_1 =" & Chk_Sistema.Value & ", "
            vgQuery = vgQuery & "num_menu_1_1 =" & Chk_Usuarios.Value & ", "
            vgQuery = vgQuery & "num_menu_1_2 =" & Chk_Contrasena.Value & ", "
            vgQuery = vgQuery & "num_menu_1_3 =" & Chk_Niveles.Value & ", "
            
            vgQuery = vgQuery & "num_menu_2 =" & Chk_ManInformacion.Value & ", "
            vgQuery = vgQuery & "num_menu_2_1 =" & Chk_ManAntecedentes.Value & ", "
            vgQuery = vgQuery & "num_menu_2_1_1 =" & Chk_AntGrales.Value & ", "
            vgQuery = vgQuery & "num_menu_2_1_2 =" & Chk_AntTutores.Value & ", "
            vgQuery = vgQuery & "num_menu_2_1_3 =" & Chk_CertSuperv.Value & ", "
            vgQuery = vgQuery & "num_menu_2_1_4 =" & Chk_AntArchBen.Value & ", "
            
            vgQuery = vgQuery & "num_menu_2_2 =" & Chk_ManRetJud.Value & ", "
            vgQuery = vgQuery & "num_menu_2_2_1 =" & Chk_RetOrdenJud.Value & ", "
            vgQuery = vgQuery & "num_menu_2_2_2 =" & Chk_RetInfoControl.Value & ", "
            
            vgQuery = vgQuery & "num_menu_2_3 =" & Chk_ManHabDesc.Value & ", "
            vgQuery = vgQuery & "num_menu_2_3_1 =" & Chk_HabDesMan.Value & ", "
            vgQuery = vgQuery & "num_menu_2_3_2 =" & Chk_HabDesAuto.Value & ", "
            
            vgQuery = vgQuery & "num_menu_2_4 =" & Chk_ManMensajes.Value & ", "
            vgQuery = vgQuery & "num_menu_2_4_1 =" & Chk_MenIndividuales.Value & ", "
            vgQuery = vgQuery & "num_menu_2_4_2 =" & Chk_MenAutomaticos.Value & ", "
            vgQuery = vgQuery & "num_menu_2_4_3 =" & Chk_MenCarga.Value & ", "
                        
            vgQuery = vgQuery & "num_menu_2_5 =" & Chk_ManEndosos.Value & ", "
            vgQuery = vgQuery & "num_menu_2_5_1 =" & Chk_GenerarEndosos.Value & ", "
            vgQuery = vgQuery & "num_menu_2_5_2 =" & Chk_ConsultaEndosos.Value & ", "
          
                           
            vgQuery = vgQuery & "num_menu_3 =" & Chk_CalPension.Value & ", "
            vgQuery = vgQuery & "num_menu_3_1 =" & Chk_CalParametros.Value & ", "
            vgQuery = vgQuery & "num_menu_3_2 =" & Chk_CalPrimeras.Value & ", "
            vgQuery = vgQuery & "num_menu_3_2_1 =" & Chk_PriProvisorio.Value & ", "
            vgQuery = vgQuery & "num_menu_3_2_2 =" & Chk_PriDefinitivo.Value & ", "
            '
            vgQuery = vgQuery & "num_menu_3_3 =" & Chk_CalInformes.Value & ", "
            vgQuery = vgQuery & "num_menu_3_3_1 =" & Chk_Validaciones.Value & ", "
            vgQuery = vgQuery & "num_menu_3_3_2 =" & Chk_InfoLiqPago.Value & ", "
            vgQuery = vgQuery & "num_menu_3_3_3 =" & Chk_InfoSalud.Value & ", "
            vgQuery = vgQuery & "num_menu_3_3_4 =" & Chk_InfoContables.Value & ", "
            vgQuery = vgQuery & "num_menu_3_3_5 =" & Chk_InfoSVS.Value & ", "
            vgQuery = vgQuery & "num_menu_3_3_6 =" & Chk_InfoCierre.Value & ", "
            
            vgQuery = vgQuery & "num_menu_3_4 =" & Chk_CalPagosTer.Value & ", "
            vgQuery = vgQuery & "num_menu_3_4_1 =" & chk_GastoSepelio.Value & ", "
            vgQuery = vgQuery & "num_menu_3_4_2 =" & Chk_PerGarantizado.Value & ", "
            vgQuery = vgQuery & "num_menu_3_4_3 =" & Chk_ConPagosTer.Value & ", "
            
            vgQuery = vgQuery & "num_menu_3_5 =" & Chk_ArchContables.Value & ", "
            vgQuery = vgQuery & "num_menu_3_5_1 =" & Chk_PagRecurrentes.Value & ", "
            vgQuery = vgQuery & "num_menu_3_5_2 =" & Chk_GastosSepelio.Value & ", "
            vgQuery = vgQuery & "num_menu_3_5_3 =" & Chk_PeriodoGar.Value & ", "
            
            vgQuery = vgQuery & "num_menu_3_6 =" & Chk_CalCalendario.Value & ", "
            vgQuery = vgQuery & "num_menu_3_7 =" & Chk_CalApeManual.Value & ", "
            
            vgQuery = vgQuery & "num_menu_4 =" & Chk_Consultas.Value & ", "
            vgQuery = vgQuery & "num_menu_4_1 =" & Chk_ConConsulta.Value & ", "
                         
            vgQuery = vgQuery & "cod_usuariomodi = '" & vlGlsUsuarioModi & "', "
            vgQuery = vgQuery & "fec_modi = '" & vlFecModi & "', "
            vgQuery = vgQuery & "hor_modi = '" & vlHorModi & "' "
            
            vgQuery = vgQuery & "WHERE "
            vgQuery = vgQuery & "cod_sistema = '" & vgTipoSistema & "' AND "
            vgQuery = vgQuery & "cod_nivel = " & vlNivel & ""
            vgConexionBD.Execute (vgQuery)
            
            MsgBox "La Actualización de Datos fue realizado Correctamente", vbInformation, "Información"
            'Call Cmd_Limpiar_Click '05/11/2005 Comentado por sugerencia de AVB
            
        End If
    End If
    
    If (vlOperacion = "I") Then
        fgComboNivel Cmb_Nivel
        
        '05/11/2005 Busca el nivel Agregado y deja el combo posicionado en esa fila (sugerencia AVB)
        For i = 0 To Cmb_Nivel.ListCount
            If Cmb_Nivel.List(i) = vlNivelAgregado Then
                Cmb_Nivel.ListIndex = i
                Exit For
            End If
        Next i
    End If

    Screen.MousePointer = 0
    
Exit Function
Err_Registrar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'--------------------------------------------------------
'Permite mostrar datos niveles
'--------------------------------------------------------
Function flMostrarDatos(iNivel As Integer)
On Error GoTo Err_mostrar
    
    'Consulta por nivel
    vgQuery = ""
    
    vgQuery = "SELECT gls_nivel, "
    vgQuery = vgQuery & "num_menu_1,"
    vgQuery = vgQuery & "num_menu_1_1,"
    vgQuery = vgQuery & "num_menu_1_2,"
    vgQuery = vgQuery & "num_menu_1_3,"
    
    vgQuery = vgQuery & "num_menu_2,"
    vgQuery = vgQuery & "num_menu_2_1,"
    vgQuery = vgQuery & "num_menu_2_1_1,"
    vgQuery = vgQuery & "num_menu_2_1_2,"
    vgQuery = vgQuery & "num_menu_2_1_3,"
    vgQuery = vgQuery & "num_menu_2_1_4,"
    
    vgQuery = vgQuery & "num_menu_2_2,"
    vgQuery = vgQuery & "num_menu_2_2_1,"
    vgQuery = vgQuery & "num_menu_2_2_2,"
    
    vgQuery = vgQuery & "num_menu_2_3,"
    vgQuery = vgQuery & "num_menu_2_3_1,"
    vgQuery = vgQuery & "num_menu_2_3_2,"
    
    vgQuery = vgQuery & "num_menu_2_4,"
    vgQuery = vgQuery & "num_menu_2_4_1,"
    vgQuery = vgQuery & "num_menu_2_4_2,"
    vgQuery = vgQuery & "num_menu_2_4_3,"
    
    vgQuery = vgQuery & "num_menu_2_5,"
    vgQuery = vgQuery & "num_menu_2_5_1,"
    vgQuery = vgQuery & "num_menu_2_5_2,"
    
    vgQuery = vgQuery & "num_menu_3,"
    vgQuery = vgQuery & "num_menu_3_1,"
    
    vgQuery = vgQuery & "num_menu_3_2,"
    vgQuery = vgQuery & "num_menu_3_2_1,"
    vgQuery = vgQuery & "num_menu_3_2_2,"
    
    vgQuery = vgQuery & "num_menu_3_3,"
    vgQuery = vgQuery & "num_menu_3_3_1,"
    vgQuery = vgQuery & "num_menu_3_3_2,"
    vgQuery = vgQuery & "num_menu_3_3_3,"
    vgQuery = vgQuery & "num_menu_3_3_4,"
    vgQuery = vgQuery & "num_menu_3_3_5,"
    vgQuery = vgQuery & "num_menu_3_3_6,"
    
    vgQuery = vgQuery & "num_menu_3_4, "
    vgQuery = vgQuery & "num_menu_3_4_1,"
    vgQuery = vgQuery & "num_menu_3_4_2,"
    vgQuery = vgQuery & "num_menu_3_4_3,"
    
    vgQuery = vgQuery & "num_menu_3_5,"
    vgQuery = vgQuery & "num_menu_3_5_1,"
    vgQuery = vgQuery & "num_menu_3_5_2,"
    vgQuery = vgQuery & "num_menu_3_5_3,"
    
    vgQuery = vgQuery & "num_menu_3_6,"
    vgQuery = vgQuery & "num_menu_3_7,"
    
    vgQuery = vgQuery & "num_menu_4, "
    vgQuery = vgQuery & "num_menu_4_1 "
    
    vgQuery = vgQuery & "FROM MA_TPAR_NIVEL WHERE "
    vgQuery = vgQuery & "cod_sistema = '" & vgTipoSistema & "' AND "
    vgQuery = vgQuery & "cod_nivel = " & iNivel
    Set vgRs = vgConexionBD.Execute(vgQuery)
    If Not (vgRs.EOF) Then
        Txt_Nivel = IIf(IsNull(vgRs!gls_nivel), "", vgRs!gls_nivel)
        Chk_Sistema.Value = vgRs!num_menu_1
        Chk_Usuarios.Value = vgRs!num_menu_1_1
        Chk_Contrasena.Value = vgRs!num_menu_1_2
        Chk_Niveles.Value = vgRs!num_menu_1_3
        
        Chk_ManInformacion.Value = vgRs!num_menu_2
        Chk_ManAntecedentes.Value = vgRs!num_menu_2_1
        Chk_AntGrales.Value = vgRs!num_menu_2_1_1
        Chk_AntTutores.Value = vgRs!num_menu_2_1_2
        Chk_CertSuperv.Value = vgRs!num_menu_2_1_3
        Chk_AntArchBen.Value = vgRs!num_menu_2_1_4
        
        Chk_ManRetJud.Value = vgRs!num_menu_2_2
        Chk_RetOrdenJud.Value = vgRs!num_menu_2_2_1
        Chk_RetInfoControl.Value = vgRs!num_menu_2_2_2

        Chk_ManHabDesc.Value = vgRs!num_menu_2_3
        Chk_HabDesMan.Value = vgRs!num_menu_2_3_1
        Chk_HabDesAuto.Value = vgRs!num_menu_2_3_2
     
        Chk_ManMensajes.Value = vgRs!num_menu_2_4
        Chk_MenIndividuales.Value = vgRs!num_menu_2_4_1
        Chk_MenAutomaticos.Value = vgRs!num_menu_2_4_2
        Chk_MenCarga.Value = vgRs!num_menu_2_4_3
        
        Chk_ManEndosos.Value = vgRs!num_menu_2_5
        Chk_GenerarEndosos.Value = vgRs!num_menu_2_5_1
        Chk_ConsultaEndosos.Value = vgRs!num_menu_2_5_2
    

        Chk_CalPension.Value = vgRs!num_menu_3
        Chk_CalParametros.Value = vgRs!num_menu_3_1
        
        Chk_CalPrimeras.Value = vgRs!num_menu_3_2
        Chk_PriProvisorio.Value = vgRs!num_menu_3_2_1
        Chk_PriDefinitivo.Value = vgRs!num_menu_3_2_2
        
        Chk_CalInformes.Value = vgRs!num_menu_3_3
        Chk_Validaciones.Value = vgRs!num_menu_3_3_1
        Chk_InfoLiqPago.Value = vgRs!num_menu_3_3_2
        Chk_InfoSalud.Value = vgRs!num_menu_3_3_3
        Chk_InfoContables.Value = vgRs!num_menu_3_3_4
        Chk_InfoSVS.Value = vgRs!num_menu_3_3_5
        Chk_InfoCierre.Value = vgRs!num_menu_3_3_6
        
        Chk_CalPagosTer.Value = vgRs!num_menu_3_4
        chk_GastoSepelio.Value = vgRs!num_menu_3_4_1
        Chk_PerGarantizado.Value = vgRs!num_menu_3_4_2
        Chk_ConPagosTer.Value = vgRs!num_menu_3_4_3
                
        Chk_ArchContables.Value = vgRs!num_menu_3_5
        Chk_PagRecurrentes.Value = vgRs!num_menu_3_5_1
        Chk_GastosSepelio.Value = vgRs!num_menu_3_5_2
        Chk_PeriodoGar.Value = vgRs!num_menu_3_5_3
        
        Chk_CalCalendario.Value = vgRs!num_menu_3_6
        Chk_CalApeManual.Value = vgRs!num_menu_3_7
        
        Chk_Consultas.Value = vgRs!num_menu_4
        Chk_ConConsulta.Value = vgRs!num_menu_4_1
                
    Else
        Chk_Sistema.Value = 0
        Chk_ManInformacion.Value = 0
        Chk_CalPension.Value = 0
        Chk_Consultas.Value = 0
        Txt_Nivel = ""
    End If
    vgRs.Close
    Screen.MousePointer = 0

Exit Function
Err_mostrar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Private Sub Chk_ArchContables_Click()

    If Chk_ArchContables.Value = 1 Then
    
        Chk_PagRecurrentes.Value = 1
        Chk_GastosSepelio.Value = 1
        Chk_PeriodoGar.Value = 1
        
        Chk_PagRecurrentes.Enabled = True
        Chk_GastosSepelio.Enabled = True
        Chk_PeriodoGar.Enabled = True
        
    Else
    
        Chk_PagRecurrentes.Value = 0
        Chk_GastosSepelio.Value = 0
        Chk_PeriodoGar.Value = 0
        
        Chk_PagRecurrentes.Enabled = False
        Chk_GastosSepelio.Enabled = False
        Chk_PeriodoGar.Enabled = False
    
    End If
    

End Sub

Private Sub Chk_CalInformes_Click()

    If Chk_CalInformes.Value = 1 Then
    
       Chk_Validaciones.Value = 1
       Chk_InfoLiqPago.Value = 1
       Chk_InfoSalud.Value = 1
       Chk_InfoContables.Value = 1
       Chk_InfoSVS.Value = 1
       Chk_InfoCierre.Value = 1
       
       Chk_Validaciones.Enabled = True
       Chk_InfoLiqPago.Enabled = True
       Chk_InfoSalud.Enabled = True
       Chk_InfoContables.Enabled = True
       Chk_InfoSVS.Enabled = True
       Chk_InfoCierre.Enabled = True
       
    Else
    
        Chk_Validaciones.Value = 0
        Chk_InfoLiqPago.Value = 0
        Chk_InfoSalud.Value = 0
        Chk_InfoContables.Value = 0
        Chk_InfoSVS.Value = 0
        Chk_InfoCierre.Value = 0
        
        Chk_Validaciones.Enabled = False
        Chk_InfoLiqPago.Enabled = False
        Chk_InfoSalud.Enabled = False
        Chk_InfoContables.Enabled = False
        Chk_InfoSVS.Enabled = False
        Chk_InfoCierre.Enabled = False
    
    End If

End Sub

Private Sub Chk_CalPagosTer_Click()

    If Chk_CalPagosTer.Value = 1 Then

       chk_GastoSepelio.Value = 1
       Chk_PerGarantizado.Value = 1
       Chk_ConPagosTer.Value = 1
   
       chk_GastoSepelio.Enabled = True
       Chk_PerGarantizado.Enabled = True
       Chk_ConPagosTer.Enabled = True
           
    Else
    
       chk_GastoSepelio.Value = 0
       Chk_PerGarantizado.Value = 0
       Chk_ConPagosTer.Value = 0
   
       chk_GastoSepelio.Enabled = False
       Chk_PerGarantizado.Enabled = False
       Chk_ConPagosTer.Enabled = False
    
   
    End If

End Sub



Private Sub Chk_CalPension_Click()

    If Chk_CalPension.Value = 1 Then
    
       Chk_CalParametros.Value = 1
       Chk_CalPrimeras.Value = 1
       Chk_CalInformes.Value = 1
       Chk_CalPagosTer.Value = 1
       Chk_ArchContables.Value = 1
       Chk_CalCalendario.Value = 1
       Chk_CalApeManual.Value = 1
    
       Chk_CalParametros.Enabled = True
       Chk_CalPrimeras.Enabled = True
       Chk_CalInformes.Enabled = True
       Chk_CalPagosTer.Enabled = True
       Chk_ArchContables.Enabled = True
       Chk_CalCalendario.Enabled = True
       Chk_CalApeManual.Enabled = True
    
    Else
    
       Chk_CalParametros.Value = 0
       Chk_CalPrimeras.Value = 0
       Chk_CalInformes.Value = 0
       Chk_CalPagosTer.Value = 0
       Chk_ArchContables.Value = 0
       Chk_CalCalendario.Value = 0
       Chk_CalApeManual.Value = 0
    
       Chk_CalParametros.Enabled = False
       Chk_CalPrimeras.Enabled = False
       Chk_CalInformes.Enabled = False
       Chk_CalPagosTer.Enabled = False
       Chk_ArchContables.Enabled = False
       Chk_CalCalendario.Enabled = False
       Chk_CalApeManual.Enabled = False
    
    End If
    

End Sub

Private Sub Chk_CalPrimeras_Click()

    If Chk_CalPrimeras.Value = 1 Then
    
       Chk_PriProvisorio.Value = 1
       Chk_PriDefinitivo.Value = 1
       
       Chk_PriProvisorio.Enabled = True
       Chk_PriDefinitivo.Enabled = True
    
    Else
    
        Chk_PriProvisorio.Value = 0
        Chk_PriDefinitivo.Value = 0
       
        Chk_PriProvisorio.Enabled = False
        Chk_PriDefinitivo.Enabled = False
    
    End If

End Sub

Private Sub Chk_Consultas_Click()

    If Chk_Consultas.Value = 1 Then
    
       Chk_ConConsulta.Value = 1
   
       Chk_ConConsulta.Enabled = True
           
    Else
    
        Chk_ConConsulta.Value = 0

       
        Chk_ConConsulta.Enabled = False
   
    End If

End Sub

Private Sub Chk_ManAntecedentes_Click()

    If Chk_ManAntecedentes.Value = 1 Then
    
       Chk_AntGrales.Value = 1
       Chk_AntTutores.Value = 1
       Chk_CertSuperv.Value = 1
       Chk_AntArchBen.Value = 1
       
       Chk_AntGrales.Enabled = True
       Chk_AntTutores.Enabled = True
       Chk_CertSuperv.Enabled = True
       Chk_AntArchBen.Enabled = True
    
    Else
    
        Chk_AntGrales.Value = 0
        Chk_AntTutores.Value = 0
        Chk_CertSuperv.Value = 0
        Chk_AntArchBen.Value = 0
       
        Chk_AntGrales.Enabled = False
        Chk_AntTutores.Enabled = False
        Chk_CertSuperv.Enabled = False
        Chk_AntArchBen.Enabled = False
    
    End If

End Sub


Private Sub Chk_ManEndosos_Click()
    If Chk_ManEndosos.Value = 1 Then
    
       Chk_GenerarEndosos.Value = 1
       Chk_ConsultaEndosos.Value = 1
       
       Chk_GenerarEndosos.Enabled = True
       Chk_ConsultaEndosos.Enabled = True
    
    Else
    
        Chk_GenerarEndosos.Value = 0
        Chk_ConsultaEndosos.Value = 0
       
        Chk_GenerarEndosos.Enabled = False
        Chk_ConsultaEndosos.Enabled = False
    
    End If

End Sub



Private Sub Chk_ManHabDesc_Click()

    If Chk_ManHabDesc.Value = 1 Then
        
       Chk_HabDesMan.Value = 1
       Chk_HabDesAuto.Value = 1
       
       Chk_HabDesMan.Enabled = True
       Chk_HabDesAuto.Enabled = True
       
    Else
    
        Chk_HabDesMan.Value = 0
        Chk_HabDesAuto.Value = 0
       
        Chk_HabDesMan.Enabled = False
        Chk_HabDesAuto.Enabled = False
    
    End If

End Sub

Private Sub Chk_ManInformacion_Click()

    If Chk_ManInformacion.Value = 1 Then
       
       Chk_ManAntecedentes.Value = 1
       Chk_ManRetJud.Value = 1
       Chk_ManHabDesc.Value = 1
       Chk_ManMensajes.Value = 1
       Chk_ManEndosos.Value = 1
       
       Chk_ManAntecedentes.Enabled = True
       Chk_ManRetJud.Enabled = True
       Chk_ManHabDesc.Enabled = True
       Chk_ManMensajes.Enabled = True
       Chk_ManEndosos.Enabled = True
       
    Else
    
        Chk_ManAntecedentes.Value = 0
        Chk_ManRetJud.Value = 0
        Chk_ManHabDesc.Value = 0
        Chk_ManMensajes.Value = 0
        Chk_ManEndosos.Value = 0
        
        Chk_ManAntecedentes.Enabled = False
        Chk_ManRetJud.Enabled = False
        Chk_ManHabDesc.Enabled = False
        Chk_ManMensajes.Enabled = False
        Chk_ManEndosos.Enabled = False
    
    End If

End Sub

Private Sub Chk_ManMensajes_Click()

    If Chk_ManMensajes.Value = 1 Then
    
       Chk_MenIndividuales.Value = 1
       Chk_MenAutomaticos.Value = 1
       Chk_MenCarga.Value = 1
       
       Chk_MenIndividuales.Enabled = True
       Chk_MenAutomaticos.Enabled = True
       Chk_MenCarga.Enabled = True
    
    Else
    
        Chk_MenIndividuales.Value = 0
        Chk_MenAutomaticos.Value = 0
        Chk_MenCarga.Value = 0
       
        Chk_MenIndividuales.Enabled = False
        Chk_MenAutomaticos.Enabled = False
        Chk_MenCarga.Enabled = False
    
    End If

End Sub

Private Sub Chk_ManRetJud_Click()

    If Chk_ManRetJud.Value = 1 Then
    
       Chk_RetOrdenJud.Value = 1
       Chk_RetInfoControl.Value = 1
       
       Chk_RetOrdenJud.Enabled = True
       Chk_RetInfoControl.Enabled = True
    
    Else
        
        Chk_RetOrdenJud.Value = 0
        Chk_RetInfoControl.Value = 0
       
        Chk_RetOrdenJud.Enabled = False
        Chk_RetInfoControl.Enabled = False
        
    End If
    

End Sub

Private Sub Cmb_Nivel_Change()
If Not IsNumeric(Cmb_Nivel) Then
    Cmb_Nivel = ""
End If
If (Cmb_Nivel) = "" Then
    Chk_Sistema.Value = 0
    Chk_Usuarios.Value = 0
    Chk_Contrasena.Value = 0
End If
End Sub

Private Sub Cmb_Nivel_Click()
On Error GoTo Err_Nivel

    If IsNumeric(Cmb_Nivel) Then
        vlNivel = CLng(Cmb_Nivel)
        flMostrarDatos vlNivel
    Else
        Chk_Sistema.Value = 0
        Chk_Usuarios.Value = 0
        Chk_Contrasena.Value = 0
    End If

Exit Sub
Err_Nivel:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmb_Nivel_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    If (Trim(Cmb_Nivel) = "") Then
        MsgBox "Debe Ingresar un Valor para el Nivel a Registrar.", vbInformation, "Error de Datos"
    Else
        Txt_Nivel.SetFocus
    End If
Else
    If Cmb_Nivel <> "" Then
        'Validar que no sobrepase los 20 caracteres
        If Len(Cmb_Nivel) > 2 And KeyAscii <> 8 Then
            KeyAscii = 0
        End If
    End If
End If
End Sub

Private Sub Cmb_Nivel_LostFocus()

    Cmb_Nivel_Click

End Sub


Private Sub Cmd_Grabar_Click()
On Error GoTo Err_Grabar
    Cmb_Nivel = Format(Cmb_Nivel, "#0")
    
    'Validar ingreso de Nivel
    If (Not IsNumeric(Cmb_Nivel)) Then
        MsgBox "Debe Ingresar el Nivel a Registrar.", vbCritical, "Error de Datos"
        Cmb_Nivel.SetFocus
        Exit Sub
    End If
    
    'Validar rangos de Nivel
    If CLng(Cmb_Nivel) < 0 Then
        MsgBox "Debe Ingresar un Valor Mayor a 0 (Cero) para el Nivel a Registrar.", vbCritical, "Error de Datos"
        Cmb_Nivel.SetFocus
        Exit Sub
    End If
    If CLng(Cmb_Nivel) > 100 Then
        MsgBox "Debe Ingresar un Valor Menor a 100 (Cien) para el Nivel a Registrar.", vbCritical, "Error de Datos"
        Cmb_Nivel.SetFocus
        Exit Sub
    End If
    
    flRegistrarNivel

Exit Sub
Err_Grabar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Limpiar_Click()
On Error GoTo Err_Limpiar

    Chk_Sistema.Value = 0
    Chk_ManInformacion.Value = 0
    Chk_CalPension.Value = 0
    ''Chk_CalPagosTer.Value = 0
    Chk_Consultas.Value = 0
    ''Chk_ArchContables.Value = 0

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

    Frm_SisNivel.Left = 0
    Frm_SisNivel.Top = 0

    fgComboNivel Cmb_Nivel
    
    If IsNumeric(Cmb_Nivel) Then
        vlNivel = CLng(Cmb_Nivel)
        flMostrarDatos vlNivel
    End If

Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Chk_Sistema_Click()

    If Chk_Sistema.Value = 1 Then
        Chk_Usuarios.Value = 1
        Chk_Contrasena.Value = 1
        Chk_Niveles.Value = 1
        
        Chk_Usuarios.Enabled = True
        Chk_Contrasena.Enabled = True
        Chk_Niveles.Enabled = True
       
    Else
        Chk_Usuarios.Value = 0
        Chk_Contrasena.Value = 0
        Chk_Niveles.Value = 0
        
        Chk_Usuarios.Enabled = False
        Chk_Contrasena.Enabled = False
        Chk_Niveles.Enabled = False
       
    End If

End Sub

Private Sub Txt_Nivel_LostFocus()
    Txt_Nivel = UCase(Txt_Nivel)
End Sub


Private Sub Txt_Nivel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Chk_Sistema.SetFocus
End If
End Sub



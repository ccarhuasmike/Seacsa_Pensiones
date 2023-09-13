VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "Crystl32.OCX"
Begin VB.Form Frm_SisUsuario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenedor de Usuarios del Sistema."
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   Icon            =   "Frm_SisUsuario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   8040
   Begin VB.Frame Fra_Datos 
      Height          =   4575
      Index           =   1
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   5775
      Begin VB.Frame Frame1 
         ForeColor       =   &H00800000&
         Height          =   1455
         Index           =   2
         Left            =   720
         TabIndex        =   18
         Top             =   3000
         Width           =   4455
         Begin VB.TextBox Txt_Contraseña 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   2520
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   7
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox Txt_Repetir 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   2520
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   8
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Lbl_Nombre 
            Alignment       =   2  'Center
            Caption         =   "( Contraseña entre 6 - 10 caracteres )"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   9
            Left            =   480
            TabIndex        =   21
            Top             =   1080
            Width           =   3495
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Contraseña                  :"
            Height          =   255
            Index           =   6
            Left            =   480
            TabIndex        =   20
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Repitir Contraseña       :"
            Height          =   255
            Index           =   7
            Left            =   480
            TabIndex        =   19
            Top             =   720
            Width           =   1935
         End
      End
      Begin VB.TextBox Txt_Nom 
         Height          =   285
         Left            =   2160
         MaxLength       =   25
         TabIndex        =   3
         Top             =   1440
         Width           =   3375
      End
      Begin VB.TextBox Txt_Mat 
         Height          =   285
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   5
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox Txt_Pat 
         Height          =   285
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   4
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox Txt_NumIdent 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         MaxLength       =   20
         TabIndex        =   2
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox Txt_Usuario 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   0
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox Cmb_Nivel 
         BackColor       =   &H00E0FFFF&
         Height          =   315
         ItemData        =   "Frm_SisUsuario.frx":0442
         Left            =   2160
         List            =   "Frm_SisUsuario.frx":0444
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2520
         Width           =   3375
      End
      Begin VB.ComboBox Cmb_TipoIdent 
         BackColor       =   &H00E0FFFF&
         Height          =   315
         ItemData        =   "Frm_SisUsuario.frx":0446
         Left            =   2160
         List            =   "Frm_SisUsuario.frx":0448
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   3375
      End
      Begin Crystal.CrystalReport Rpt_Usuarios 
         Left            =   120
         Top             =   3360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowState     =   2
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nivel de Acceso           :"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   28
         Top             =   2565
         Width           =   1815
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nombre                        :"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   27
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Ap. Materno                 :"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   26
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Ap. Paterno                  :"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   25
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Tipo Identificación        :"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   24
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Usuario                          :"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Número Identificación   :"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   22
         Top             =   1080
         Width           =   1935
      End
   End
   Begin VB.Frame Fra_Botones 
      Height          =   1095
      Index           =   4
      Left            =   120
      TabIndex        =   16
      Top             =   4560
      Width           =   5775
      Begin VB.CommandButton cmd_salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   4700
         Picture         =   "Frm_SisUsuario.frx":044A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton cmd_imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   2500
         Picture         =   "Frm_SisUsuario.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton cmd_eliminar 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   1400
         Picture         =   "Frm_SisUsuario.frx":0BFE
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton cmd_limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   3600
         Picture         =   "Frm_SisUsuario.frx":0F40
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton cmd_grabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   300
         Picture         =   "Frm_SisUsuario.frx":15FA
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Fra_Usuarios 
      Caption         =   "  Usuarios  "
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
      Height          =   5655
      Index           =   3
      Left            =   6000
      TabIndex        =   15
      Top             =   0
      Width           =   1935
      Begin VB.ListBox Lst_Usuarios 
         BackColor       =   &H00E0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5100
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Frm_SisUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vlRecord    As ADODB.Recordset

Dim vlVerificar       As String
Dim vlOperacion       As String
Dim vlRut             As String
Dim vlDigito          As String
Dim vlSw              As Boolean
Dim vlPassword        As String
Dim vlOculto          As String
Dim vlUsuario         As String
Dim vlPosicion        As Long
Dim vlNivel           As Long

Dim vlLargoTipoIden    As Integer 'sirve para llenar la grilla
Dim vlPosicionTipoIden As Integer 'sirve para llenar la grilla
Dim vlNumIdent         As Integer
Dim vlCodTipoIden As Integer 'sirve para guardar el código

'--------------------------------------------------------
'                   F U N C I O N E S
'--------------------------------------------------------
Function flLimpiar()

    Txt_Usuario = ""
    
    If (Cmb_TipoIdent.ListCount > 0) Then
        Cmb_TipoIdent.ListIndex = 0
    End If
    
    Txt_NumIdent = ""
    Txt_Nom = ""
    Txt_Pat = ""
    Txt_Mat = ""
    Txt_Contraseña = ""
    Txt_Repetir = ""
    Txt_Usuario.Enabled = True
  
End Function

'--------------------------------------------------------
'Permite actualizar la Lista de los códigos existentes
'--------------------------------------------------------
Function flActLista()
On Error GoTo Err_Actualizar

    Lst_Usuarios.Clear
    vgSql = "SELECT cod_usuario FROM MA_tmae_usuario WHERE "
    vgSql = vgSql & "cod_sistema = '" & vgTipoSistema & "'"
    vgSql = vgSql & "ORDER BY cod_usuario "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
        While Not vgRs.EOF
            'Lst_Usuarios.a
            'vgPalabra = CStr(vgRs!rut_usu) & " - " & Trim(CStr(vgRs!dgv_usu))
            'vgPalabra = Space(14 - Len(vgPalabra)) & vgPalabra
            vgPalabra = UCase(Trim(vgRs!Cod_Usuario))
            Lst_Usuarios.AddItem vgPalabra
            vgRs.MoveNext
        Wend
    End If
    vgRs.Close

Exit Function
Err_Actualizar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'--------------------------------------------------------
'Permite llenar los casilleros de información con los
'datos del Usuario seleccionado
'--------------------------------------------------------
Function flBuscarUsuario(iUsuario)
Dim vlNivelAux As String
On Error GoTo Err_Buscar
    
    vgSql = "SELECT * FROM MA_tmae_usuario WHERE "
    vgSql = vgSql & "cod_sistema = '" & vgTipoSistema & "' AND "
    vgSql = vgSql & "cod_usuario = '" & iUsuario & "' "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
    
        'Busca el Nivel asignado al Usuario
        vgI = 0
        Cmb_Nivel.ListIndex = 0
        Do While vgI < Cmb_Nivel.ListCount
            If (Trim(Cmb_Nivel) <> "") Then
'                If (CStr(vgRs!cod_nivel) = Trim(Cmb_Nivel.Text)) Then
                vlNivelAux = Trim(Mid(Cmb_Nivel, 1, InStr(1, Cmb_Nivel, "-") - 1))
                If (CStr(vgRs!cod_nivel) = Trim(vlNivelAux)) Then
                    Exit Do
                End If
            End If
            vgI = vgI + 1
            If (vgI = Cmb_Nivel.ListCount) Then
                MsgBox "El Nivel del Usuario no se encuentra identificada o no existe.", vbExclamation, "Dato Inexistente"
                Cmb_Nivel.ListIndex = -1
                Exit Do
            End If
            Cmb_Nivel.ListIndex = vgI
        Loop
        
        'Buscar el Còdigo del Tipo de Identificación
        'JT
        vlPosicionTipoIden = fgBuscarPosicionCodigoCombo(vgRs!Cod_TipoIdenusu, Cmb_TipoIdent)
        If (vlPosicionTipoIden >= 0) Then
            Cmb_TipoIdent.ListIndex = vlPosicionTipoIden
        End If
        
        Txt_NumIdent = Trim(vgRs!num_idenusu)
        Txt_Nom = Trim(vgRs!gls_nombre)
        Txt_Pat = Trim(vgRs!gls_paterno)
        Txt_Mat = Trim(vgRs!gls_materno)
        Txt_Usuario = Trim(vgRs!Cod_Usuario)
        Txt_Contraseña = fgDesPassword(Trim(vgRs!gls_password))
        Txt_Repetir = Txt_Contraseña
        Txt_Usuario.Enabled = False
        
    End If
    vgRs.Close

'    Txt_Nom.SetFocus

Exit Function
Err_Buscar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'------------------------------------------------------------
'Permite Ingresar o Modificar los datos de los Usuarios
'con autorización al uso del Sistema
'------------------------------------------------------------
Function flRegistrar()
Dim vlSuc As String
On Error GoTo Err_Registrar
Dim vlCodTipoIden As String
    
    vlSuc = "0000"
    'Valida Usuario
    If Trim(Txt_Usuario) = "" Then
        MsgBox "Debe ingresar el Usuario.", vbCritical, "Error de Datos"
        Txt_Usuario.SetFocus
        Exit Function
    End If
    
    'Valida el Tipo de Identificación
    If Cmb_TipoIdent.Text = "" Then
        MsgBox "Debe Seleccionar un Tipo de Identificación.", vbCritical, "Error de Datos"
        Cmb_TipoIdent.SetFocus
        Exit Function
    End If

    'Validar Número de Identificación
    If Txt_NumIdent = "" Then
        MsgBox "Debe Ingresar el Número de Identificación.", vbCritical, "Error de Datos"
        Txt_NumIdent.SetFocus
        Exit Function
    End If

   'Valida el Nombre del Usuario
    If Trim(Txt_Nom) = "" Then
        MsgBox "Debe ingresar el Nombre de Usuario.", vbCritical, "Error de Datos"
        Txt_Nom.SetFocus
        Exit Function
    End If
    
    'Validar Apellido Paterno Usuario
    If Trim(Txt_Pat) = "" Then
        MsgBox "Debe ingresar el Apellido Paterno de Usuario.", vbCritical, "Error de Datos"
        Txt_Pat.SetFocus
        Exit Function
    End If
    
    'Validar Apellido Materno
    If Trim(Txt_Mat) = "" Then
        MsgBox "Debe ingresar el Apellido Materno de Usuario.", vbCritical, "Error de Datos"
        Txt_Mat.SetFocus
        Exit Function
    End If
    
    'Validar la Contraseña y su Repetición
    If (Trim(Txt_Contraseña) <> "") And (Trim(Txt_Repetir) <> "") Then
        Txt_Contraseña = Trim(Txt_Contraseña)
        Txt_Repetir = Trim(Txt_Repetir)
        If (Len(Txt_Contraseña) < 6) Then
            MsgBox "La contraseña debe contar como mínimo de 6 caracteres, y como máximo 10.", vbInformation, "Error de Dato"
            Txt_Contraseña.SetFocus
            Exit Function
        End If
        If (Txt_Contraseña <> Txt_Repetir) Then
            MsgBox "Las Contraseñas registradas son distintas, vuelva a registrarlas.", vbExclamation, "Error de Contraseña"
            Txt_Contraseña = ""
            Txt_Repetir = ""
            Txt_Contraseña.SetFocus
            Exit Function
        End If
    Else
        MsgBox "Debe ingresar la Contraseña y su correspondiente repetición para ser registrado.", vbCritical, "Error de Contraseña"
        Txt_Contraseña.SetFocus
        Exit Function
    End If
    'Validar selección de Nivel
    If Trim(Cmb_Nivel) = "" Then
        MsgBox "Debe seleccionar el nivel de acceso de usuario.", vbCritical, "Error de Contraseña"
        Cmb_Nivel.SetFocus
        Exit Function
    End If
    
    Txt_Usuario = Trim(UCase(Txt_Usuario))
    Txt_NumIdent = Trim(UCase(Txt_NumIdent))
    Txt_Nom = UCase(Trim(Txt_Nom))
    Txt_Pat = UCase(Trim(Txt_Pat))
    Txt_Mat = UCase(Trim(Txt_Mat))
    Txt_Contraseña = Trim(UCase(Txt_Contraseña))
    Txt_Repetir = Trim(UCase(Txt_Repetir))
    'vlNivel = CLng(Cmb_Nivel)
    vlNivel = CLng(Mid(Cmb_Nivel, 1, InStr(1, Cmb_Nivel, "-") - 1))
    vlCodTipoIden = Trim(Mid(Cmb_TipoIdent.Text, 1, (InStr(1, Cmb_TipoIdent, "-") - 1)))
    
    vlOperacion = ""
    vlSw = False
    
    'Encriptar Contraseña
    vlPassword = fgEncPassword(Txt_Contraseña)
    If (vlPassword = "") Then
        MsgBox "Error en la transformación de la Password o Contraseña por el Sistema.", vbCritical, "Error de Transformación"
        Exit Function
    End If
    
    Screen.MousePointer = 11
                
    'Verificar existencia de Rut de Usuario para el Ingreso/Actualización
    vgSql = "SELECT cod_usuario FROM MA_tmae_usuario WHERE "
    vgSql = vgSql & "cod_sistema = '" & vgTipoSistema & "' AND "
    vgSql = vgSql & "cod_usuario = '" & Txt_Usuario & "' "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If (vgRs.EOF) Then
        vlOperacion = "I"
    Else
        vlOperacion = "A"
    End If
    vgRs.Close
    
    If (vlOperacion = "I") Then
        'Ingresar Usuario
        vgQuery = "INSERT INTO MA_tmae_usuario ("
        vgQuery = vgQuery & "cod_usuario,cod_sistema,cod_nivel,gls_password, "
        vgQuery = vgQuery & "cod_tipoidenusu,num_idenusu,"
        vgQuery = vgQuery & "gls_nombre,gls_paterno,gls_materno,"
        'vgQuery = vgQuery & "cod_sucursal, "
        vgQuery = vgQuery & "cod_usuariocrea, "
        vgQuery = vgQuery & "fec_crea, hor_crea "
        vgQuery = vgQuery & ") VALUES ("
        vgQuery = vgQuery & "'" & Txt_Usuario & "', "
        vgQuery = vgQuery & "'" & vgTipoSistema & "', "
        vgQuery = vgQuery & " " & vlNivel & ", "
        vgQuery = vgQuery & "'" & vlPassword & "', "
        vgQuery = vgQuery & " " & vlCodTipoIden & ", "
        vgQuery = vgQuery & "'" & Txt_NumIdent & "', "
        vgQuery = vgQuery & "'" & Txt_Nom & "', "
        vgQuery = vgQuery & "'" & Txt_Pat & "', "
        vgQuery = vgQuery & "'" & Txt_Mat & "', "
        'vgQuery = vgQuery & "'" & vlSuc & "', "
        vgQuery = vgQuery & "'" & vgUsuario & "', "
        vgQuery = vgQuery & "'" & Format(Date, "yyyymmdd") & "', "
        vgQuery = vgQuery & "'" & Format(Time, "hhmmss") & "'"
        vgQuery = vgQuery & ")"
        vgConexionBD.Execute (vgQuery)
        vlSw = True
    Else
        If (vlOperacion = "A") Then
            'Actualiza Datos del Usuario
            vgRes = MsgBox("¿ Está seguro que desea Modificar los Datos del Usuario ?", 4 + 32 + 256, "Operación de Actualización")
            If vgRes <> 6 Then
                Cmd_Salir.SetFocus
                Screen.MousePointer = 0
                Exit Function
            End If
            
            vgQuery = ""
            vgQuery = "UPDATE MA_tmae_usuario SET "
            vgQuery = vgQuery & "cod_nivel = " & vlNivel & ", "
            vgQuery = vgQuery & "gls_password = '" & vlPassword & "', "
            vgQuery = vgQuery & "cod_tipoidenusu = " & vlCodTipoIden & ", "
            vgQuery = vgQuery & "num_idenusu = '" & Txt_NumIdent & "', "
            vgQuery = vgQuery & "gls_nombre = '" & Txt_Nom & "', "
            vgQuery = vgQuery & "gls_paterno = '" & Txt_Pat & "', "
            vgQuery = vgQuery & "gls_materno = '" & Txt_Mat & "', "
'            vgquery = vgquery & "cod_sucursal = '" & vlsuc & ", "
            vgQuery = vgQuery & "fec_modi = '" & Format(Date, "yyyymmdd") & "', "
            vgQuery = vgQuery & "hor_modi = '" & Format(Time, "hhmmss") & "', "
            vgQuery = vgQuery & "cod_usuariomodi = '" & vgUsuario & "' "
            vgQuery = vgQuery & "WHERE "
            vgQuery = vgQuery & "cod_sistema = '" & vgTipoSistema & "' AND "
            vgQuery = vgQuery & "cod_usuario = '" & Txt_Usuario & "' "
            vgConexionBD.Execute (vgQuery)
            vlSw = False
        End If
    End If
    
    If (vgUsuario = Txt_Usuario) Then
        vgContraseña = vlPassword
    End If
    
    'Permite limpiar los casilleros de información
    fgComboNivelGlosa Cmb_Nivel
    
    'Limpia los Casilleros de Información de Usuario
    flLimpiar
    If (vlOperacion = "I") Then
        'Actualizar Lista de Usuarios con Acceso
        Call flActLista
    End If
    Txt_Usuario.SetFocus
    
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
'Permite Eliminar al Usuario de la Base de Datos
'--------------------------------------------------------
Function flEliminar()
On Error GoTo Err_Eliminar

     
    'Validar la selección del Usuario a Eliminar
    If (Trim(Txt_Usuario) = "") Then
        MsgBox "Debe ingresar el Login del Usuario a eliminar.", vbCritical, "Error de Datos"
        Txt_Usuario.SetFocus
        Exit Function
    End If
    
    vlOperacion = ""
    vlSw = False
    
    Screen.MousePointer = 11
    
    'Verificar existencia de Rut del Usuario para la Eliminación
    vgQuery = "SELECT cod_usuario FROM MA_tmae_usuario WHERE "
    vgQuery = vgQuery & "cod_sistema = '" & vgTipoSistema & "' AND "
    vgQuery = vgQuery & "cod_usuario = '" & Txt_Usuario & "' "
    Set vgRs = vgConexionBD.Execute(vgQuery)
    If Not (vgRs.EOF) Then
        vlOperacion = "E"
    End If
    vgRs.Close
   
    If (vlOperacion = "E") Then
        vgRes = MsgBox("  ¿ Está seguro que desea Eliminar este Usuario ?  ", 4 + 32 + 256, "Operación de Eliminación")
        If vgRes <> 6 Then
            Screen.MousePointer = 0
            Cmd_Salir.SetFocus
            Exit Function
        End If
        
        vgQuery = "DELETE FROM MA_tmae_usuario WHERE "
        vgQuery = vgQuery & "cod_sistema = '" & vgTipoSistema & "' AND "
        vgQuery = vgQuery & "cod_usuario = '" & Txt_Usuario & "' "
        vgConexionBD.Execute (vgQuery)
        
        vlSw = True
    End If

    If (vlSw = True) Then
        'Actualizar Lista de Usuarios
        Call flActLista
    End If
    'Permite limpiar los casilleros de información
    fgComboNivelGlosa Cmb_Nivel
    Call flLimpiar
    Txt_Usuario.SetFocus
    
    Screen.MousePointer = 0
    
Exit Function
Err_Eliminar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Sub flImpresion()
Dim vlArchivo As String

Err.Clear
On Error GoTo Errores1
   
   Screen.MousePointer = 11
   
   vlArchivo = strRpt & "MA_Rpt_Usuarios.rpt"   '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
       MsgBox "Archivo de Reporte de Listados no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
       Screen.MousePointer = 0
       Exit Sub
   End If
   
    vgQuery = "SELECT * FROM MA_TPAR_SISTEMA WHERE "
    vgQuery = vgQuery & "cod_sistema = '" & vgTipoSistema & "'"
    Set vgRs = vgConexionBD.Execute(vgQuery)
    If Not vgRs.EOF Then
        vgPalabra = IIf(Not IsNull(vgRs!gls_sistema), Trim(vgRs!gls_sistema), "")
    Else
        vgPalabra = " "
    End If
    vgRs.Close
   
   vgQuery = "{MA_TMAE_USUARIO.COD_SISTEMA} = '" & vgTipoSistema & "'"
   
   Rpt_Usuarios.Reset
   Rpt_Usuarios.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
   'MDIPrincipal.Rpt.DataFiles(0) = vgRutaBasedeDatos       ' o App.Path & "\Nestle.mdb"
   Rpt_Usuarios.Connect = vgRutaDataBase
   Rpt_Usuarios.SelectionFormula = ""
   Rpt_Usuarios.SelectionFormula = vgQuery
   Rpt_Usuarios.Formulas(0) = ""
   Rpt_Usuarios.Formulas(1) = ""
   Rpt_Usuarios.Formulas(2) = ""
   Rpt_Usuarios.Formulas(3) = ""
   
   Rpt_Usuarios.Formulas(0) = "TipoSistema = '" & vgPalabra & "'"
   Rpt_Usuarios.Formulas(1) = "NombreCompania = '" & vgNombreCompania & "'"
   Rpt_Usuarios.Formulas(2) = "NombreSistema = '" & vgNombreSistema & "'"
   Rpt_Usuarios.Formulas(3) = "NombreSubSistema = '" & vgNombreSubSistema & "'"
   
   Rpt_Usuarios.WindowState = crptMaximized
   Rpt_Usuarios.Destination = crptToWindow
   Rpt_Usuarios.WindowTitle = "Informe de Usuarios del Sistema"
   Rpt_Usuarios.Action = 1
   
   Screen.MousePointer = 0
   
Exit Sub
Errores1:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " " & "Numero de Error : " & Err.Number & Err.HelpFile & Err.HelpContext, vbExclamation, "Mensaje de error ..."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub Cmb_Nivel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Txt_Contraseña.SetFocus
End If
End Sub

Private Sub Cmb_TipoIdent_Click()
If (Cmb_TipoIdent <> "") Then
    vlPosicionTipoIden = Cmb_TipoIdent.ListIndex
    vlLargoTipoIden = Cmb_TipoIdent.ItemData(vlPosicionTipoIden)
    Txt_NumIdent.MaxLength = vlLargoTipoIden
    If (Txt_NumIdent <> "") Then Txt_NumIdent.Text = Mid(Txt_NumIdent, 1, vlLargoTipoIden)
End If
End Sub

Private Sub Cmb_TipoIdent_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    If (Cmb_TipoIdent <> "") Then
        Txt_NumIdent.SetFocus
    End If
End If
End Sub

'----------------------------------------------------------
'  PROCEDIMIENTOS DE LOS OBJETOS
'----------------------------------------------------------

Private Sub Cmd_Eliminar_Click()
On Error GoTo Err_Eliminar
    
    'Permite eliminar al Usuario de la Base de Datos
    Call flEliminar
    
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
    
    'Permite registrar los datos de los Usuarios
    Call flRegistrar
    
Exit Sub
Err_Grabar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Imprimir_Click()
On Error GoTo Err_Imprimir
    
    'Validar si se ha seleccionado el Cliente
    'If (Trim(vgCliente) = "") Or Not IsNumeric(vgCliente) Then
    '    MsgBox "No se ha seleccionado el Cliente sobre el cual se realizará la Operación.", vbCritical, "Operación Cancelada"
    '    Screen.MousePointer = 0
    '    Exit Sub
    'End If
    
    'Permite limpiar los casilleros de información
    fgComboNivelGlosa Cmb_Nivel
    
    'Limpia los Datos del Formulario
    Call flLimpiar
    'Imprime el Reporte de Errores
    flImpresion
            
Exit Sub
Err_Imprimir:
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

    Frm_SisUsuario.Left = 0
    Frm_SisUsuario.Top = 0
    
    fgComboNivelGlosa Cmb_Nivel
    fgComboTipoIdentificacion Cmb_TipoIdent
  
    'Actualizar Lista de Usuarios con Acceso al Sistema
    Call flActLista
    Call flLimpiar
    
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Limpiar_Click()
On Error GoTo Err_Limpiar
    
    'Permite limpiar los casilleros de información
    fgComboNivelGlosa Cmb_Nivel
    
    Call flLimpiar
    Txt_Usuario.SetFocus
    
Exit Sub
Err_Limpiar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub



Private Sub Lst_Usuarios_Click()
On Error GoTo Err_Error

    If (Lst_Usuarios.ListCount > 0) Then
        If (Lst_Usuarios.Text <> "") Then
            'vlPosicion = InStr(1, Lst_Usuarios.Text, "-")
            'vlRut = Trim(Mid(Lst_Usuarios.Text, 1, vlPosicion - 1))
            vlRut = UCase(Trim(Lst_Usuarios.Text))
            'VLNumero = Trim(Mid(Lst_Usuarios.Text, vlPosicion + 1, Len(Lst_Usuarios.Text)))
            Call flLimpiar
            Call flBuscarUsuario(vlRut)
        End If
    End If

Exit Sub
Err_Error:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub


Private Sub Txt_Contraseña_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
'    If (Trim(Txt_Contraseña) <> "") And (Txt_Repetir <> "") Then
'        Txt_Contraseña = Trim(Txt_Contraseña)
'        Txt_Repetir = Trim(Txt_Repetir)
'        If (Len(Txt_Contraseña) < 6) Then
'            MsgBox "La contraseña debe contar como mínimo de 6 caracteres, y como máximo 10.", vbInformation, "Error de Dato"
'            Txt_Contraseña.SetFocus
'        End If
'        If (Txt_Contraseña <> Txt_Repetir) Then
'            MsgBox "Las Contraseñas registradas son distintas, vuelva a registrarlas.", vbExclamation, "Error de Contraseña"
'            Txt_Contraseña = ""
'            Txt_Repetir = ""
'        Else
'            Cmd_Grabar.SetFocus
'        End If
'    Else
        If (Trim(Txt_Contraseña) <> "") Then
            If (Len(Trim(Txt_Contraseña)) < 6) Then
                MsgBox "La contraseña debe contar como mínimo de 6 caracteres, y como máximo 10.", vbInformation, "Error de Dato"
                Txt_Contraseña.SetFocus
            Else
                Txt_Contraseña = UCase(Trim(Txt_Contraseña))
                Txt_Repetir.SetFocus
            End If
        End If
'    End If
End If
End Sub


Private Sub Txt_Mat_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    If (Trim(Txt_Mat) <> "") Then
        Txt_Mat = UCase(Trim(Txt_Mat))
        Cmb_Nivel.SetFocus
    End If
End If
End Sub

Private Sub Txt_Nom_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    If (Trim(Txt_Nom) <> "") Then
        Txt_Nom = UCase(Trim(Txt_Nom))
        Txt_Pat.SetFocus
    End If
End If
End Sub

Private Sub Txt_NumIdent_GotFocus()
    Txt_NumIdent.SelStart = 0
    Txt_NumIdent.SelLength = Len(Txt_NumIdent)
End Sub

Private Sub Txt_NumIdent_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    If (Trim(Txt_NumIdent) <> "") Then
        Txt_NumIdent = UCase(Trim(Txt_NumIdent))
        Txt_Nom.SetFocus
    End If
End If
End Sub

Private Sub Txt_NumIdent_LostFocus()
If (Trim(Txt_NumIdent) <> "") Then
    Txt_NumIdent = UCase(Trim(Txt_NumIdent))
End If
End Sub

Private Sub Txt_Pat_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    If (Trim(Txt_Pat) <> "") Then
        Txt_Pat = UCase(Trim(Txt_Pat))
        Txt_Mat.SetFocus
    End If
End If
End Sub

Private Sub Txt_Repetir_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    If (Trim(Txt_Contraseña) <> "") And (Txt_Repetir <> "") Then
        Txt_Contraseña = UCase(Trim(Txt_Contraseña))
        Txt_Repetir = UCase(Trim(Txt_Repetir))
        If (Len(Txt_Contraseña) < 6) Then
            MsgBox "La contraseña debe contar como mínimo de 6 caracteres, y como máximo 10.", vbInformation, "Error de Dato"
            Txt_Contraseña.SetFocus
        End If
        If (Txt_Contraseña <> Txt_Repetir) Then
            MsgBox "Las Contraseñas registradas son distintas, vuelva a registrarlas.", vbExclamation, "Error de Contraseña"
            Txt_Contraseña = ""
            Txt_Repetir = ""
            Txt_Contraseña.SetFocus
        Else
            Cmd_Grabar.SetFocus
        End If
    Else
        MsgBox "Debe ingresar primeramente la Contraseña, y posteriormente repetirla.", vbExclamation, "Error de Contraseña"
        Txt_Contraseña = ""
        Txt_Repetir = ""
        Txt_Contraseña.SetFocus
    End If
End If
End Sub

Private Sub Txt_Usuario_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    If (Trim(Txt_Usuario) <> "") Then
        Txt_Usuario = Trim(UCase(Txt_Usuario))
        Cmb_TipoIdent.SetFocus
    End If
End If
End Sub

Private Sub Txt_Usuario_LostFocus()
On Error GoTo Err_Error

    If (Txt_Usuario.Text <> "") Then
        'vlPosicion = InStr(1, Lst_Usuarios.Text, "-")
        'vlRut = Trim(Mid(Lst_Usuarios.Text, 1, vlPosicion - 1))
        vlRut = UCase(Trim(Txt_Usuario.Text))
        'VLNumero = Trim(Mid(Lst_Usuarios.Text, vlPosicion + 1, Len(Lst_Usuarios.Text)))
        'Call flLimpiar
        Txt_Nom = ""
        Txt_Pat = ""
        Txt_Mat = ""
        Txt_Contraseña = ""
        'Txt_Usuario = ""
        Txt_Repetir = ""
        Txt_NumIdent = ""
        Txt_Usuario.Enabled = True
        'Opt_1.Value = True
        Call flBuscarUsuario(vlRut)
    End If

Exit Sub
Err_Error:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub



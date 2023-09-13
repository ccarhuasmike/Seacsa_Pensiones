VERSION 5.00
Begin VB.Form Frm_InicioSesion 
   Caption         =   "Inicio de Sesión"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6255
   ClipControls    =   0   'False
   Icon            =   "Frm_InicioSesion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Creación de DSN "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1035
      Left            =   120
      TabIndex        =   16
      Top             =   3480
      Width           =   6000
      Begin VB.TextBox Txt_DSN 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         MaxLength       =   30
         TabIndex        =   5
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre de DSN                  : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   930
         TabIndex        =   17
         Top             =   480
         Width           =   2610
      End
      Begin VB.Image Image 
         Height          =   480
         Index           =   2
         Left            =   270
         Picture         =   "Frm_InicioSesion.frx":030A
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Inicio de Sesión "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1400
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   6000
      Begin VB.TextBox Txt_Inicio 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         MaxLength       =   30
         TabIndex        =   3
         Top             =   300
         Width           =   1935
      End
      Begin VB.TextBox Txt_Contraseña 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   3600
         MaxLength       =   30
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   870
         Width           =   1935
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre de Inicio de Sesión  : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   930
         TabIndex        =   14
         Top             =   360
         Width           =   2625
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   "Contraseña                         : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   960
         TabIndex        =   13
         Top             =   840
         Width           =   2610
      End
      Begin VB.Image Image 
         Height          =   480
         Index           =   1
         Left            =   270
         Picture         =   "Frm_InicioSesion.frx":074C
         Top             =   450
         Width           =   480
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Servidor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1995
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   6000
      Begin VB.TextBox txt_Proveedor 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         MaxLength       =   30
         TabIndex        =   0
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txt_server 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         MaxLength       =   30
         TabIndex        =   1
         Top             =   930
         Width           =   1935
      End
      Begin VB.TextBox txt_bd 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   3600
         MaxLength       =   30
         TabIndex        =   2
         Top             =   1470
         Width           =   1935
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre Proveedor              : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   960
         TabIndex        =   15
         Top             =   480
         Width           =   2550
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre o IP del Servidor     : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   960
         TabIndex        =   10
         Top             =   960
         Width           =   2565
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre Base de Datos       : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   930
         TabIndex        =   9
         Top             =   1500
         Width           =   2565
      End
      Begin VB.Image Image 
         Height          =   480
         Index           =   0
         Left            =   250
         Picture         =   "Frm_InicioSesion.frx":0A56
         Top             =   450
         Width           =   480
      End
   End
   Begin VB.Frame Fra_Botones 
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   4560
      Width           =   6015
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1005
         TabIndex        =   6
         Top             =   240
         Width           =   1500
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3570
         TabIndex        =   7
         Top             =   225
         Width           =   1500
      End
   End
End
Attribute VB_Name = "Frm_InicioSesion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
On Error GoTo Err_Salir

    Unload Me
    'End

Exit Sub
Err_Salir:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub cmdOK_Click()
On Error GoTo error_inicio
             
    txt_Proveedor.Text = UCase(Trim(txt_Proveedor.Text))
    txt_server.Text = UCase(Trim(txt_server.Text))
    txt_bd.Text = UCase(Trim(txt_bd.Text))
    Txt_Inicio.Text = UCase(Trim(Txt_Inicio.Text))
    Txt_Contraseña.Text = UCase(Trim(Txt_Contraseña.Text))
    Txt_DSN.Text = UCase(Trim(Txt_DSN.Text))
  
    resp = EscribeArchivoIni("Conexion", "Proveedor", txt_Proveedor.Text, App.Path & "\AdmPrevBD.ini")
    resp = EscribeArchivoIni("Conexion", "Servidor", txt_server.Text, App.Path & "\AdmPrevBD.ini")
    resp = EscribeArchivoIni("Conexion", "BaseDatos", txt_bd.Text, App.Path & "\AdmPrevBD.ini")
    resp = EscribeArchivoIni("Conexion", "Usuario", fgEncPassword(Txt_Inicio.Text), App.Path & "\AdmPrevBD.ini")
    resp = EscribeArchivoIni("Conexion", "Password", fgEncPassword(Txt_Contraseña.Text), App.Path & "\AdmPrevBD.ini")
    resp = EscribeArchivoIni("Conexion", "DSN", Txt_DSN.Text, App.Path & "\AdmPrevBD.ini")
     
     ProviderName = txt_Proveedor.Text
     ServerName = txt_server.Text
     DatabaseName = txt_bd.Text
     UserName = Txt_Inicio.Text
     PasswordName = Txt_Contraseña.Text
             
    MsgBox "El Archivo de Configuración ha sido generado exitosamente", vbInformation, "Creación Archivo Configuración"
    Unload Me
    Exit Sub
             
Exit Sub
error_inicio:
    MsgBox "Error " & Err & " : " & error, vbExclamation, "Atención"
    ''inises = "sincn"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Cargar

    txt_Proveedor.Text = LeeArchivoIni("Conexion", "Proveedor", "", App.Path & "\AdmPrevBD.ini")
    txt_server.Text = LeeArchivoIni("Conexion", "Servidor", "", App.Path & "\AdmPrevBD.ini")
    txt_bd.Text = LeeArchivoIni("Conexion", "BaseDatos", "", App.Path & "\AdmPrevBD.ini")
    Txt_Inicio.Text = LeeArchivoIni("Conexion", "Usuario", "", App.Path & "\AdmPrevBD.ini")
    Txt_Contraseña.Text = LeeArchivoIni("Conexion", "Password", "", App.Path & "\AdmPrevBD.ini")
    Txt_DSN.Text = LeeArchivoIni("Conexion", "DSN", "", App.Path & "\AdmPrevBD.ini")

    
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub txt_bd_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    'txt_DSN.SetFocus
    'cmdOK.SetFocus
    Txt_Inicio.SetFocus
End If
End Sub

Private Sub Txt_Contraseña_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Txt_DSN.SetFocus
End If
End Sub

Private Sub txt_DSN_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    cmdOK.SetFocus
End If
End Sub

Private Sub Txt_Inicio_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Txt_Contraseña.SetFocus
End If
End Sub

Private Sub txt_Proveedor_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    txt_server.SetFocus
End If
End Sub

Private Sub txt_server_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    txt_bd.SetFocus
End If
End Sub

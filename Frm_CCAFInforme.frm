VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Frm_CCAFInforme 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Control"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4200
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   3975
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   2040
         Picture         =   "Frm_CCAFInforme.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   960
         Picture         =   "Frm_CCAFInforme.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Imprimir Estadísticas"
         Top             =   240
         Width           =   790
      End
      Begin Crystal.CrystalReport Rpt_Calculo 
         Left            =   120
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.Frame Frame1 
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.TextBox Txt_Fecha 
         Height          =   285
         Left            =   2040
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de Pago"
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
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Frm_CCAFInforme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd_Imprimir_Click()
On Error GoTo Err_Imprimir

   vlArchivo = App.Path & "\Reportes\PP_Rpt_CCAFcontrol.rpt"   '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
       MsgBox "Archivo de Reporte de Póliza no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
       Screen.MousePointer = 0
       Exit Sub
   End If
      
    vlFecha = Format(Trim(Txt_Fecha), "yyyymmdd")
    
            
    vgQuery = "{PP_TMAE_HABDESCCAF.FEC_INIHABDES}<= '" & vlFecha & "' AND "
    vgQuery = vgQuery & "{PP_TMAE_HABDESCCAF.FEC_TERHABDES}>= '" & vlFecha & "'"

    Rpt_Calculo.Reset
    Rpt_Calculo.WindowState = crptMaximized
    Rpt_Calculo.ReportFileName = vlArchivo
    Rpt_Calculo.Connect = vgRutaDataBase
    Rpt_Calculo.SelectionFormula = ""
    Rpt_Calculo.SelectionFormula = vgQuery
    
    Rpt_Calculo.Formulas(0) = "NombreCompania='" & vgNombreCompania & "'"
    Rpt_Calculo.Formulas(1) = "NombreSistema= '" & vgNombreSistema & "'"
    Rpt_Calculo.Formulas(2) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
    Rpt_Calculo.Formulas(3) = "Fecha= '" & vlFecha & "'"
    Rpt_Calculo.SubreportToChange = ""
    Rpt_Calculo.Destination = crptToWindow
    Rpt_Calculo.WindowTitle = "Informe de Control Cajas de Compensación"
    Rpt_Calculo.Action = 1
    Screen.MousePointer = 0
    


Exit Sub
Err_Imprimir:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub

Private Sub cmd_salir_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Frm_CCAFInforme.Top = 0
    Frm_CCAFInforme.Left = 0
End Sub

Private Sub Txt_Fecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If IsDate(Trim(Txt_Fecha)) Then
        Txt_Fecha.Text = Format(CDate(Trim(Txt_Fecha)), "yyyymmdd")
        Txt_Fecha.Text = DateSerial(Mid((Txt_Fecha.Text), 1, 4), Mid((Txt_Fecha.Text), 5, 2), Mid((Txt_Fecha.Text), 7, 2))
        
        If CDate(Trim(Txt_Fecha)) > CDate(Date) Then
            MsgBox "La Fecha de Proceso es superior a la Fecha Actual.", vbCritical, "Operación Cancelada"
            Exit Sub
        Else
            Cmd_Imprimir.SetFocus
        End If
    Else
        MsgBox "La Fecha de Proceso Ingresada no es válida.", vbCritical, "Operación Cancelada"
        Txt_Fecha = ""
    End If
End If

End Sub

Private Sub Txt_Fecha_LostFocus()
Txt_Fecha = Trim(Txt_Fecha)
If IsDate(Txt_Fecha) Then
    Txt_Fecha.Text = Format(CDate(Trim(Txt_Fecha)), "yyyymmdd")
    Txt_Fecha.Text = DateSerial(Mid((Txt_Fecha.Text), 1, 4), Mid((Txt_Fecha.Text), 5, 2), Mid((Txt_Fecha.Text), 7, 2))
Else
    Txt_Fecha = ""
End If
End Sub

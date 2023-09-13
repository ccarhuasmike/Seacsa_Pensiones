VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Frm_Archivos_PDT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Información para PDT por Periodo"
   ClientHeight    =   6570
   ClientLeft      =   1995
   ClientTop       =   1740
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   10020
   Begin VB.Frame Frame2 
      Height          =   4215
      Left            =   90
      TabIndex        =   15
      Top             =   1230
      Width           =   9795
      Begin MSComctlLib.ListView lvw_lista 
         Height          =   3855
         Left            =   120
         TabIndex        =   5
         Top             =   210
         Width           =   9525
         _ExtentX        =   16801
         _ExtentY        =   6800
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin MSComDlg.CommonDialog ComDialogo 
      Left            =   60
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Fra_Busqueda 
      Height          =   1215
      Left            =   90
      TabIndex        =   10
      Top             =   0
      Width           =   9795
      Begin VB.OptionButton OptBajas 
         Caption         =   "Bajas"
         Height          =   240
         Left            =   6570
         TabIndex        =   19
         Top             =   405
         Width           =   795
      End
      Begin VB.OptionButton optAltas 
         Caption         =   "Altas"
         Height          =   240
         Left            =   5775
         TabIndex        =   18
         Top             =   405
         Width           =   735
      End
      Begin VB.CommandButton Cmd_BuscarPol 
         Height          =   375
         Left            =   3990
         Picture         =   "Frm_Archivos_PDT.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Buscar Póliza"
         Top             =   750
         Width           =   615
      End
      Begin VB.TextBox Txt_Mes 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2310
         MaxLength       =   2
         TabIndex        =   2
         Top             =   780
         Width           =   360
      End
      Begin VB.TextBox Txt_Anno 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3030
         MaxLength       =   4
         TabIndex        =   3
         Top             =   780
         Width           =   795
      End
      Begin VB.ComboBox Cmb_Tipo 
         BackColor       =   &H80000018&
         Height          =   315
         ItemData        =   "Frm_Archivos_PDT.frx":0102
         Left            =   2340
         List            =   "Frm_Archivos_PDT.frx":0104
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   3315
      End
      Begin VB.Label lbl_tipcam 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   8700
         TabIndex        =   17
         Top             =   450
         Width           =   795
      End
      Begin VB.Label lbl_nombre 
         Caption         =   "Tipo de Cambio:"
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
         Height          =   405
         Index           =   1
         Left            =   7800
         TabIndex        =   16
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lbl_nombre 
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
         Left            =   2745
         TabIndex        =   14
         Top             =   795
         Width           =   195
      End
      Begin VB.Label lbl_nombre 
         Caption         =   "(Mes - Año)"
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
         Height          =   270
         Index           =   10
         Left            =   1140
         TabIndex        =   13
         Top             =   795
         Width           =   1005
      End
      Begin VB.Label lbl_nombre 
         Caption         =   "Período "
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
         Index           =   8
         Left            =   285
         TabIndex        =   12
         Top             =   780
         Width           =   705
      End
      Begin VB.Label lbl_nombre 
         Caption         =   "Tipo de Archivo     :"
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
         Index           =   2
         Left            =   285
         TabIndex        =   0
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lbl_nombre 
         Caption         =   " :"
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
         Height          =   240
         Index           =   0
         Left            =   2100
         TabIndex        =   11
         Top             =   825
         Width           =   165
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1035
      Left            =   90
      TabIndex        =   9
      Top             =   5460
      Width           =   9795
      Begin VB.CommandButton Cmd_Cargar 
         Caption         =   "&Archivo"
         Height          =   675
         Left            =   4575
         Picture         =   "Frm_Archivos_PDT.frx":0106
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Exportar Datos a Archivo"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5700
         Picture         =   "Frm_Archivos_PDT.frx":0928
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   3480
         Picture         =   "Frm_Archivos_PDT.frx":0A22
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   720
      End
      Begin Crystal.CrystalReport Rpt_Reporte 
         Left            =   120
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
End
Attribute VB_Name = "Frm_Archivos_PDT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmb_Tipo_Click()
    lvw_lista.ListItems.Clear
    Call p_formato_lista(lvw_lista, Mid(Cmb_Tipo.Text, 1, 3))
    If Mid(Cmb_Tipo.Text, 1, 3) = "P00" Then
        optAltas.Visible = True
        OptBajas.Visible = True
        optAltas.Value = True
    Else
        optAltas.Visible = False
        OptBajas.Visible = False
    End If
    'Call Cmd_BuscarPol_Click
End Sub

Private Sub Cmd_BuscarPol_Click()

    If Val(Txt_Mes.Text) = 1 Then
        Txt_Anno.Tag = Val(Txt_Anno.Text) - 1
        Txt_Mes.Tag = 12
    Else
        Txt_Anno.Tag = Txt_Anno.Text
        Txt_Mes.Tag = Val(Txt_Mes.Text) - 1
    End If
    lbl_tipcam.Caption = fTipoCambioSBS("US", Format$(Txt_Anno.Tag, "0000") & Format$(Txt_Mes.Tag, "00"))
    If Mid(Cmb_Tipo.Text, 1, 3) = "DER" Then
        Call p_llena_datos_der
    Else
        Call p_llena_datos
    End If

End Sub

Private Sub Cmd_Cargar_Click()

Dim vlOpen As Boolean
Dim sNombreArchivo As String
Dim iFila As Long
Dim iColumn As Byte
Dim vlLinea As String

On Error GoTo Err_Cargar

'Valida Año
    If Txt_Anno.Text = "" Then
       MsgBox "Debe Ingresar Año del Periodo de Pago.", vbCritical, "Error de Datos"
       Txt_Anno.SetFocus
       Exit Sub
    End If
    If CDbl(Txt_Anno.Text) < 1900 Then
       MsgBox "Debe Ingresar un Año Mayor a 1900.", vbCritical, "Error de Datos"
       Txt_Anno.SetFocus
       Exit Sub
    End If
'Valida Mes
    If Txt_Mes.Text = "" Then
       MsgBox "Debe Ingresar Mes del Periodo de Pago.", vbCritical, "Error de Datos"
       Txt_Mes.SetFocus
       Exit Sub
    End If
    If CDbl(Txt_Mes.Text) <= 0 Or CDbl(Txt_Mes.Text) > 12 Then
       MsgBox "El Mes Ingresado No es un Valor Válido.", vbCritical, "Error de Datos"
       Txt_Mes.SetFocus
       Exit Sub
    End If

    
    If Mid(Cmb_Tipo.Text, 1, 3) = "PEN" Then
        sNombreArchivo = "0601" & Format(Txt_Anno.Text, "0000") & Format(Txt_Mes.Text, "00")
        sNombreArchivo = sNombreArchivo & "20517207331.pen"
    ElseIf Mid(Cmb_Tipo.Text, 1, 3) = "T00" Then
        sNombreArchivo = "RP_"
        sNombreArchivo = sNombreArchivo & "20517207331.ide"
    ElseIf Mid(Cmb_Tipo.Text, 1, 3) = "DER" Then
        sNombreArchivo = "RP_"
        sNombreArchivo = sNombreArchivo & "20517207331.ide"
    ElseIf Mid(Cmb_Tipo.Text, 1, 3) = "T02" Then
        sNombreArchivo = "RP_"
        sNombreArchivo = sNombreArchivo & "20517207331.pen"
    ElseIf Mid(Cmb_Tipo.Text, 1, 3) = "PER" Then
        sNombreArchivo = "0601" & Format(Txt_Anno.Text, "0000") & Format(Txt_Mes.Text, "00")
        sNombreArchivo = sNombreArchivo & "20517207331.poc"
    ElseIf Mid(Cmb_Tipo.Text, 1, 3) = "P00" Then
        sNombreArchivo = "RP_"
        sNombreArchivo = sNombreArchivo & "20517207331.per"
    End If
    
    Screen.MousePointer = 11

    'Selección del Archivo de Resumen de Reservas
    ComDialogo.CancelError = True
    ComDialogo.FileName = sNombreArchivo
    ComDialogo.DialogTitle = "Generar archivo PDT"
    ComDialogo.Filter = "*." & Mid(Cmb_Tipo.Text, 1, 3)
    ComDialogo.FilterIndex = 1
    ComDialogo.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    ComDialogo.ShowSave
    sNombreArchivo = ComDialogo.FileName
    If sNombreArchivo = "" Then
        Screen.MousePointer = 0
        Exit Sub
    End If

    Open sNombreArchivo For Output As #1
    vlOpen = True
    
Dim vDNI As String
Dim vdDev As Double
Dim vdPen As Double
vDNI = ""
vlLinea = "": vdDev = 0: vdPen = 0
iFila = 1

If Mid(Cmb_Tipo.Text, 1, 3) = "PEN" Then

    Do
            If vDNI = "05845085" Then
                vDNI = vDNI
            End If
            If Len(vlLinea) > 1 Then
                If vDNI <> Trim(lvw_lista.ListItems(iFila).SubItems(10)) Then
                    Print #1, vlLinea
                End If
            End If
            vlLinea = ""
            If vDNI <> Trim(lvw_lista.ListItems(iFila).SubItems(10)) Then
                For iColumn = 9 To lvw_lista.ColumnHeaders.Count - 1
                        vlLinea = vlLinea & Trim(lvw_lista.ListItems(iFila).SubItems(iColumn)) & "|"
                Next
            Else
                For iColumn = 9 To lvw_lista.ColumnHeaders.Count - 1
                    If iColumn = 12 Then
                        vdDev = Trim(lvw_lista.ListItems(iFila - 1).SubItems(12))
                        vdDev = vdDev + CDbl(Trim(lvw_lista.ListItems(iFila).SubItems(12)))
                        vlLinea = vlLinea & vdDev & "|"
                    ElseIf iColumn = 13 Then
                        vdPen = Trim(lvw_lista.ListItems(iFila - 1).SubItems(13))
                        vdPen = vdPen + CDbl(Trim(lvw_lista.ListItems(iFila).SubItems(13)))
                        vlLinea = vlLinea & vdPen & "|"
                    Else
                        vlLinea = vlLinea & Trim(lvw_lista.ListItems(iFila).SubItems(iColumn)) & "|"
                    End If
                Next
            End If
            
            vDNI = Trim(lvw_lista.ListItems(iFila).SubItems(10))
            iFila = iFila + 1
            
    Loop Until iFila = lvw_lista.ListItems.Count + 1
    Print #1, vlLinea

Else
    vDNI = ""
    For iFila = 1 To lvw_lista.ListItems.Count
        vlLinea = ""
        'Materia Gris - Jaime Rios 05/03/2018 inicio
        'For iColumn = 5 To lvw_lista.ColumnHeaders.Count - 1
        '    If Mid(Cmb_Tipo.Text, 1, 3) = "PEN" Then
        '        'If Not (iColumn = 8 Or iColumn = 9) Then 'Se excluyen moneda y monto de origen que son auxiliares para el reporte
        '            vlLinea = vlLinea & Trim(lvw_lista.ListItems(iFila).SubItems(iColumn)) & "|"
        '        'End If
        '    Else
        '        vlLinea = vlLinea & Trim(lvw_lista.ListItems(iFila).SubItems(iColumn)) & "|"
        '    End If
        'Next

'            If Mid(Cmb_Tipo.Text, 1, 3) = "PEN" Then
'                For iColumn = 9 To lvw_lista.ColumnHeaders.Count - 1
'                        vlLinea = vlLinea & Trim(lvw_lista.ListItems(iFila).SubItems(iColumn)) & "|"
'                Next
'            Else
                'INI Giovanni Cruz 20211118
                If vDNI <> Trim(lvw_lista.ListItems(iFila).SubItems(6)) Then
                
                    vDNI = Trim(lvw_lista.ListItems(iFila).SubItems(6))
                    For iColumn = 5 To lvw_lista.ColumnHeaders.Count - 1
                        vlLinea = vlLinea & Trim(lvw_lista.ListItems(iFila).SubItems(iColumn)) & "|"
                    Next
                    
                    Print #1, vlLinea
                End If
                'FIN Giovanni Cruz 20211118
'        End If
        'Materia Gris - Jaime Rios 05/03/2018 fin
        

    Next
    
End If
    

    Close #1

    vlOpen = False
    MsgBox "La Exportación de datos al Archivo ha sido finalizada exitosamente.", vbInformation, "Estado de Generación Archivo"

    Screen.MousePointer = 0


Exit Sub

Err_Cargar:
    Screen.MousePointer = 0
    If vlOpen Then
        Close #1
    End If
    If Err.Number = 32755 Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
        End If
    End If

End Sub

Private Sub Cmd_Imprimir_Click()

Dim xlapp As Excel.Application
Dim vlArchivo As String
Dim iFila As Long
Dim iCol As Integer

    If lvw_lista.ListItems.Count = 0 Then Exit Sub
    
    vlArchivo = strRpt & "PP_Archivo_PDT.xls"
    Set xlapp = CreateObject("excel.application")

    xlapp.Visible = True 'para ver vista previa
    xlapp.WindowState = 2 ' minimiza excel
    xlapp.Workbooks.Open (vlArchivo)
    iFila = 1
    
    For iCol = 2 To lvw_lista.ColumnHeaders.Count
        If iCol <= 27 Then
            xlapp.Range(Chr(iCol + 63) & iFila) = lvw_lista.ColumnHeaders(iCol).Text
        Else
            xlapp.Range("A" & Chr(iCol + 63 - 26) & iFila) = lvw_lista.ColumnHeaders(iCol).Text
        End If
    Next
    
    xlapp.Sheets("PDT").Name = Mid(Cmb_Tipo.Text, 1, 3)
    
    For iFila = 1 To lvw_lista.ListItems.Count
        For iCol = 1 To lvw_lista.ColumnHeaders.Count - 1
            If iCol <= 26 Then
                xlapp.Range(Chr(iCol + 64) & iFila + 1) = lvw_lista.ListItems(iFila).SubItems(iCol)
            Else
                xlapp.Range("A" & Chr(iCol + 64 - 26) & iFila + 1) = lvw_lista.ListItems(iFila).SubItems(iCol)
            End If
        Next
    Next
    GoSub s_marco
    
    xlapp.WindowState = 1 ' maxima excel

    Exit Sub
    
s_marco:
'*******
    If iCol <= 26 Then
        xlapp.Range("A1:" & Chr(iCol + 63) & iFila).Select
    Else
        xlapp.Range("A1:A" & Chr(iCol + 63 - 26) & iFila).Select
    End If
    xlapp.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    xlapp.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With xlapp.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .ColorIndex = xlAutomatic
    End With
    With xlapp.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .ColorIndex = xlAutomatic
    End With
    With xlapp.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .ColorIndex = xlAutomatic
    End With
    With xlapp.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .ColorIndex = xlAutomatic
    End With
    With xlapp.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        '.Weight = 1 'xlThick  ' xlMedium
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With xlapp.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With

    Return
    
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

    Frm_Archivos_PDT.Left = 0
    Frm_Archivos_PDT.Top = 0
    
    Txt_Mes.Text = Month(Date)
    Txt_Anno.Text = Year(Date)
    
    Call fgComboPDT(Cmb_Tipo)
    optAltas.Value = True
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub p_formato_lista(lista As ListView, ByVal sTipo As String)

Dim objCol As ColumnHeader
   
    lista.FullRowSelect = True
    lista.CheckBoxes = False
    lista.View = lvwReport
    lista.Gridlines = True
    lista.HotTracking = False
    lista.HoverSelection = False
    lista.LabelEdit = lvwManual
       
    lista.SortOrder = lvwDescending
    lista.SortKey = 0
    lista.Sorted = True
       
    lista.ColumnHeaders.Clear
    Set objCol = lista.ColumnHeaders.Add(, , , 0)
    Select Case sTipo
        Case "T00"
            Set objCol = lista.ColumnHeaders.Add(, , "Nº", 500, lvwColumnCenter)                  '1
            Set objCol = lista.ColumnHeaders.Add(, , "Póliza", 1200, lvwColumnCenter)             '2
            Set objCol = lista.ColumnHeaders.Add(, , "Tipo", 800, lvwColumnCenter)                '3
            Set objCol = lista.ColumnHeaders.Add(, , "TipPen", 800, lvwColumnLeft)                '4
            
            
            Set objCol = lista.ColumnHeaders.Add(, , "Tipo Doc.", 1000, lvwColumnCenter)          '5            '1
            Set objCol = lista.ColumnHeaders.Add(, , "Nº Documento", 1000, lvwColumnCenter)       '6            '2
            Set objCol = lista.ColumnHeaders.Add(, , "Pais.emisor del Doc.", 1000, lvwColumnCenter)       '6    '3
            Set objCol = lista.ColumnHeaders.Add(, , "Fec. Nac.", 1200, lvwColumnCenter)          '10           '4
            Set objCol = lista.ColumnHeaders.Add(, , "Ap. Paterno", 2000, lvwColumnLeft)          '7            '5
            Set objCol = lista.ColumnHeaders.Add(, , "Ap. Materno", 2000, lvwColumnLeft)          '8            '6
            Set objCol = lista.ColumnHeaders.Add(, , "Nombres", 2000, lvwColumnLeft)              '9            '7
            Set objCol = lista.ColumnHeaders.Add(, , "Sexo", 800, lvwColumnCenter)                '11           '8
            Set objCol = lista.ColumnHeaders.Add(, , "Nacion.", 1200, lvwColumnCenter)             '12          '9
            Set objCol = lista.ColumnHeaders.Add(, , "Telef.Cod.Larga", 300, lvwColumnCenter)              '13  '10
            Set objCol = lista.ColumnHeaders.Add(, , "Telef.", 1200, lvwColumnCenter)              '13          '11
            Set objCol = lista.ColumnHeaders.Add(, , "Correo Electrónico", 3000, lvwColumnLeft)   '14           '12

            'Set objCol = lista.ColumnHeaders.Add(, , "Ind. domi.", 800, lvwColumnCenter)          '16
            Set objCol = lista.ColumnHeaders.Add(, , "Tipo de Via", 800, lvwColumnCenter)         '17           '13
            Set objCol = lista.ColumnHeaders.Add(, , "Nombre de Via", 3000, lvwColumnLeft)        '18           '14
            Set objCol = lista.ColumnHeaders.Add(, , "Numero de Via", 800, lvwColumnCenter)       '19           '15
            
            Set objCol = lista.ColumnHeaders.Add(, , "Departamento", 500, lvwColumnLeft)            '20         '16
            Set objCol = lista.ColumnHeaders.Add(, , "Interior", 500, lvwColumnLeft)            '20             '17
            Set objCol = lista.ColumnHeaders.Add(, , "Manzana", 500, lvwColumnLeft)            '20              '18
            Set objCol = lista.ColumnHeaders.Add(, , "Lote", 500, lvwColumnLeft)            '20                 '19
            Set objCol = lista.ColumnHeaders.Add(, , "Kilometro", 500, lvwColumnLeft)            '20            '20
            Set objCol = lista.ColumnHeaders.Add(, , "Block", 500, lvwColumnLeft)            '20                '21
            Set objCol = lista.ColumnHeaders.Add(, , "Etapa", 500, lvwColumnLeft)            '20                '22
            
            Set objCol = lista.ColumnHeaders.Add(, , "Tipo de Zona", 800, lvwColumnCenter)        '21           '23
            Set objCol = lista.ColumnHeaders.Add(, , "Nombre de Zona", 2000, lvwColumnLeft)       '22           '24
            Set objCol = lista.ColumnHeaders.Add(, , "Referencia", 2000, lvwColumnCenter)         '23           '25
            Set objCol = lista.ColumnHeaders.Add(, , "Ubigeo", 800, lvwColumnCenter)              '24           '26
            
            'Direccion dos (solo siregistro dos direcciones)
            
            'Set objCol = lista.ColumnHeaders.Add(, , "Ind. domi.", 800, lvwColumnCenter)          '16
            Set objCol = lista.ColumnHeaders.Add(, , "Tipo de Via 2", 800, lvwColumnCenter)         '25         '27
            Set objCol = lista.ColumnHeaders.Add(, , "Nombre de Via 2", 3000, lvwColumnLeft)        '26         '28
            Set objCol = lista.ColumnHeaders.Add(, , "Numero de Via 2", 800, lvwColumnCenter)       '27         '29
            
            Set objCol = lista.ColumnHeaders.Add(, , "Departamento 2", 500, lvwColumnLeft)            '28       '30
            Set objCol = lista.ColumnHeaders.Add(, , "Interior 2", 500, lvwColumnLeft)            '29           '31
            Set objCol = lista.ColumnHeaders.Add(, , "Manzana 2", 500, lvwColumnLeft)            '30            '32
            Set objCol = lista.ColumnHeaders.Add(, , "Lote 2", 500, lvwColumnLeft)            '31               '33
            Set objCol = lista.ColumnHeaders.Add(, , "Kilometro 2", 500, lvwColumnLeft)            '32          '34
            Set objCol = lista.ColumnHeaders.Add(, , "Block 2", 500, lvwColumnLeft)            '33              '35
            Set objCol = lista.ColumnHeaders.Add(, , "Etapa 2", 500, lvwColumnLeft)            '34              '36
                
            Set objCol = lista.ColumnHeaders.Add(, , "Tipo de Zona 2", 800, lvwColumnCenter)        '35         '37
            Set objCol = lista.ColumnHeaders.Add(, , "Nombre de Zona 2", 2000, lvwColumnLeft)       '36         '38
            Set objCol = lista.ColumnHeaders.Add(, , "Referencia 2", 2000, lvwColumnCenter)         '37         '39
            Set objCol = lista.ColumnHeaders.Add(, , "Ubigeo 2", 800, lvwColumnCenter)              '38         '40
            
            'Asesoria Asistencial Essalud
            Set objCol = lista.ColumnHeaders.Add(, , "IndCA EsSaLud 2", 800, lvwColumnCenter)       '39         '41
            
'            Set objCol = lista.ColumnHeaders.Add(, , "Indicador", 800, lvwColumnCenter)              '24
'            Set objCol = lista.ColumnHeaders.Add(, , "ESSALUD", 800, lvwColumnCenter)             '15
        Case "T02"
            Set objCol = lista.ColumnHeaders.Add(, , "Nº", 500, lvwColumnCenter)                  '1
            Set objCol = lista.ColumnHeaders.Add(, , "Póliza", 1200, lvwColumnCenter)             '2
            Set objCol = lista.ColumnHeaders.Add(, , "Tipo", 800, lvwColumnCenter)                '3
            Set objCol = lista.ColumnHeaders.Add(, , "TipPen", 800, lvwColumnLeft)                '4
            Set objCol = lista.ColumnHeaders.Add(, , "Tipo Doc.", 1000, lvwColumnCenter)          '5
            Set objCol = lista.ColumnHeaders.Add(, , "Nº Documento", 1000, lvwColumnCenter)       '6
            Set objCol = lista.ColumnHeaders.Add(, , "Pais.emisor del Doc.", 1000, lvwColumnCenter)       '6
            Set objCol = lista.ColumnHeaders.Add(, , "Tipo Pens.", 1000, lvwColumnLeft)           '7
            Set objCol = lista.ColumnHeaders.Add(, , "Regimen Pens.", 1200, lvwColumnCenter)      '8
            'Set objCol = lista.ColumnHeaders.Add(, , "Fec. Insc.", 1200, lvwColumnCenter)         '9
            Set objCol = lista.ColumnHeaders.Add(, , "CUSPP", 1800, lvwColumnCenter)              '10
            'Set objCol = lista.ColumnHeaders.Add(, , "Situación", 800, lvwColumnCenter)           '11
            Set objCol = lista.ColumnHeaders.Add(, , "Tipo Pago.", 800, lvwColumnCenter)          '12
        Case "P00"
            Set objCol = lista.ColumnHeaders.Add(, , "Nº", 500, lvwColumnCenter)                  '1
            Set objCol = lista.ColumnHeaders.Add(, , "Póliza", 1200, lvwColumnCenter)             '2
            Set objCol = lista.ColumnHeaders.Add(, , "Tipo", 800, lvwColumnCenter)                '3
            Set objCol = lista.ColumnHeaders.Add(, , "TipPen", 800, lvwColumnLeft)                '4
            Set objCol = lista.ColumnHeaders.Add(, , "Tipo Doc.", 1000, lvwColumnCenter)          '5
            Set objCol = lista.ColumnHeaders.Add(, , "Nº Documento", 1000, lvwColumnCenter)       '6
            Set objCol = lista.ColumnHeaders.Add(, , "Pais Emisor", 1000, lvwColumnCenter)       '6
            Set objCol = lista.ColumnHeaders.Add(, , "Categoria", 1000, lvwColumnLeft)            '7
            Set objCol = lista.ColumnHeaders.Add(, , "Tipo Registro", 1000, lvwColumnCenter)       '6
            Set objCol = lista.ColumnHeaders.Add(, , "Fec. Inicio", 1200, lvwColumnCenter)        '8
            Set objCol = lista.ColumnHeaders.Add(, , "Fec. Final", 1200, lvwColumnCenter)         '9
            Set objCol = lista.ColumnHeaders.Add(, , "Mot. Fin", 2000, lvwColumnCenter)           '10
            Set objCol = lista.ColumnHeaders.Add(, , "Tipo Mod.", 800, lvwColumnCenter)           '11
        Case "PEN"
            Set objCol = lista.ColumnHeaders.Add(, , "Nº", 500, lvwColumnCenter)                  '1
            Set objCol = lista.ColumnHeaders.Add(, , "Póliza", 1200, lvwColumnCenter)             '2
            Set objCol = lista.ColumnHeaders.Add(, , "Tipo", 800, lvwColumnCenter)                '3
            Set objCol = lista.ColumnHeaders.Add(, , "TipPen", 800, lvwColumnLeft)                '4
            'Materia Gris - Jaime Rios 05/03/2018 inicio
            Set objCol = lista.ColumnHeaders.Add(, , "CodAFP", 800, lvwColumnCenter)
            Set objCol = lista.ColumnHeaders.Add(, , "AFP", 1800, lvwColumnCenter)
            Set objCol = lista.ColumnHeaders.Add(, , "CUSPP", 1600, lvwColumnCenter)
            Set objCol = lista.ColumnHeaders.Add(, , "Fec. Nac.", 1200, lvwColumnCenter)
            'Materia Gris - Jaime Rios 05/03/2018 fin
            Set objCol = lista.ColumnHeaders.Add(, , "Tipo Doc.", 1000, lvwColumnCenter)          '
            Set objCol = lista.ColumnHeaders.Add(, , "Nº Documento", 1000, lvwColumnCenter)       '
            Set objCol = lista.ColumnHeaders.Add(, , "Concepto", 1000, lvwColumnLeft)             '
'            Set objCol = lista.ColumnHeaders.Add(, , "Moneda", 1000, lvwColumnCenter)             '
'            Set objCol = lista.ColumnHeaders.Add(, , "Monto Origen", 1200, lvwColumnCenter)       '
            Set objCol = lista.ColumnHeaders.Add(, , "Monto Devengado", 1200, lvwColumnCenter)    '
            Set objCol = lista.ColumnHeaders.Add(, , "Monto Pagado", 1200, lvwColumnCenter)       '
        Case "DER"
            Set objCol = lista.ColumnHeaders.Add(, , "Nº", 500, lvwColumnCenter)                  '1
            Set objCol = lista.ColumnHeaders.Add(, , "Póliza", 1200, lvwColumnCenter)             '2
            Set objCol = lista.ColumnHeaders.Add(, , "Tipo", 800, lvwColumnCenter)                '3
            Set objCol = lista.ColumnHeaders.Add(, , "TipPen", 800, lvwColumnLeft)                '4
            Set objCol = lista.ColumnHeaders.Add(, , "TD Trab", 1000, lvwColumnCenter)            '
            Set objCol = lista.ColumnHeaders.Add(, , "Nº Doc. Trab", 1000, lvwColumnCenter)       '
            Set objCol = lista.ColumnHeaders.Add(, , "TD Dh", 1000, lvwColumnCenter)              '
            Set objCol = lista.ColumnHeaders.Add(, , "Nº Doc. Dh", 1000, lvwColumnCenter)         '
            Set objCol = lista.ColumnHeaders.Add(, , "Ap. Paterno Dh", 2000, lvwColumnLeft)       '
            Set objCol = lista.ColumnHeaders.Add(, , "Ap. Materno Dh", 2000, lvwColumnLeft)       '
            Set objCol = lista.ColumnHeaders.Add(, , "Nombres Dh", 2000, lvwColumnLeft)           '
            Set objCol = lista.ColumnHeaders.Add(, , "Fec. Nac.", 1200, lvwColumnCenter)          '
            Set objCol = lista.ColumnHeaders.Add(, , "Sexo", 800, lvwColumnCenter)                '
            Set objCol = lista.ColumnHeaders.Add(, , "Vinculo Fam.", 800, lvwColumnCenter)        '
            Set objCol = lista.ColumnHeaders.Add(, , "TD Acred.", 800, lvwColumnCenter)           '
            Set objCol = lista.ColumnHeaders.Add(, , "Nº Doc Acred.", 3000, lvwColumnLeft)        '
            Set objCol = lista.ColumnHeaders.Add(, , "Situación Dh", 1000, lvwColumnCenter)        '
            Set objCol = lista.ColumnHeaders.Add(, , "Fec. Alta", 1200, lvwColumnCenter)           '
            Set objCol = lista.ColumnHeaders.Add(, , "Tipo Baja", 800, lvwColumnCenter)           '
            Set objCol = lista.ColumnHeaders.Add(, , "Fec. Baja", 800, lvwColumnCenter)           '
            Set objCol = lista.ColumnHeaders.Add(, , "Nº Resol.", 800, lvwColumnCenter)           '
            
            Set objCol = lista.ColumnHeaders.Add(, , "Ind. domi.", 800, lvwColumnCenter)          '
            Set objCol = lista.ColumnHeaders.Add(, , "Tipo de Via", 800, lvwColumnCenter)         '
            Set objCol = lista.ColumnHeaders.Add(, , "Nombre de Via", 3000, lvwColumnLeft)      '
            Set objCol = lista.ColumnHeaders.Add(, , "Numero de Via", 800, lvwColumnCenter)       '
            Set objCol = lista.ColumnHeaders.Add(, , "Interior", 800, lvwColumnCenter)            '
            Set objCol = lista.ColumnHeaders.Add(, , "Tipo de Zona", 800, lvwColumnCenter)        '
            Set objCol = lista.ColumnHeaders.Add(, , "Nombre de Zona", 2000, lvwColumnLeft)       '
            Set objCol = lista.ColumnHeaders.Add(, , "Referencia", 2000, lvwColumnCenter)         '
            Set objCol = lista.ColumnHeaders.Add(, , "Ubigeo", 800, lvwColumnCenter)              '
        Case "PER"
            Set objCol = lista.ColumnHeaders.Add(, , "Nº", 500, lvwColumnCenter)                  '1
            Set objCol = lista.ColumnHeaders.Add(, , "Póliza", 1200, lvwColumnCenter)             '2
            Set objCol = lista.ColumnHeaders.Add(, , "Tipo", 800, lvwColumnCenter)                '3
            Set objCol = lista.ColumnHeaders.Add(, , "TipPen", 800, lvwColumnLeft)                '4
            Set objCol = lista.ColumnHeaders.Add(, , "Tipo Doc.", 1000, lvwColumnCenter)          '
            Set objCol = lista.ColumnHeaders.Add(, , "Nº Documento", 1200, lvwColumnCenter)       '
'            Set objCol = lista.ColumnHeaders.Add(, , "Concepto", 1000, lvwColumnLeft)             '
'            Set objCol = lista.ColumnHeaders.Add(, , "Moneda", 1000, lvwColumnCenter)             '
'            Set objCol = lista.ColumnHeaders.Add(, , "Monto Origen", 1200, lvwColumnCenter)       '
            Set objCol = lista.ColumnHeaders.Add(, , "Aporta", 750, lvwColumnCenter)    '
    End Select
    
    Set objCol = Nothing

End Sub

Private Sub p_llena_datos()

Dim rs_Temp As ADODB.Recordset
Dim sSql As String
Dim dFecIni As Long
Dim dFecFin As Long
Dim iCantidad As Long
Dim dFecEval As Long
Dim proc As ADODB.Command

    dFecIni = Txt_Anno.Text & Format(Txt_Mes.Text, "00") & "01"
    dFecFin = Txt_Anno.Text & Format(Txt_Mes.Text, "00") & f_dia_ultimo(Val(Txt_Anno.Text), Val(Txt_Mes.Text))
    

    If Mid(Cmb_Tipo.Text, 1, 3) = "P00" And OptBajas.Value = True Then
            Dim FECHA_ANT2 As String
            Dim FECHA_ANT As Date
            Dim FECHAX As Date
            Dim fecha_ant_mes As String
            Dim anio, mes As String
            
            FECHA_ANT2 = "01/" & Format(Txt_Mes.Text, "00") & "/" & Txt_Anno.Text
            FECHAX = FECHA_ANT2
            FECHA_ANT = DateAdd("d", -1, FECHAX)
            anio = Year(FECHA_ANT)
            mes = Month(FECHA_ANT)
            fecha_ant_mes = anio & Format(mes, "00") & "01"
            
            
            
'            Set rs_Temp = New ADODB.Recordset
'            rs_Temp.CursorLocation = adUseClient
'            rs_Temp.Open "PP_LISTA_PDT.LISTA_PDT_BAJAS('" & Txt_Anno.Text & "','" & Format(Txt_Mes.Text, "00") & "','" & anio & "','" & Format(mes, "00") & "','" & Mid(Cmb_Tipo.Text, 1, 3) & "','" & IIf(OptBajas.Value = True, 1, 0) & "','" & dFecIni & "','" & dFecFin & "')", vgConexionBD, adOpenStatic, adLockReadOnly
'RRR 08/08/2019
            Set proc = New ADODB.Command
            Set proc.ActiveConnection = vgConexionBD
            proc.CommandType = adCmdStoredProc
            proc.CommandText = "PP_LISTA_PDT.LISTA_PDT_BAJAS"
            proc.Prepared = False
            'proc.Parameters.Delete
            proc.Parameters.Append proc.CreateParameter("ANIO", adVarChar, adParamInput, 4, Txt_Anno.Text)
            proc.Parameters.Append proc.CreateParameter("MES", adVarChar, adParamInput, 2, Format(Txt_Mes.Text, "00"))
            proc.Parameters.Append proc.CreateParameter("ANIO_ANT", adVarChar, adParamInput, 4, anio)
            proc.Parameters.Append proc.CreateParameter("MES_ANT", adVarChar, adParamInput, 2, Format(mes, "00"))
            proc.Parameters.Append proc.CreateParameter("TIPO", adVarChar, adParamInput, 3, Mid(Cmb_Tipo.Text, 1, 3))
            proc.Parameters.Append proc.CreateParameter("B_A", adVarChar, adParamInput, 1, IIf(OptBajas.Value = True, 1, 0))
            proc.Parameters.Append proc.CreateParameter("FECHAINI", adVarChar, adParamInput, 8, dFecIni)
            proc.Parameters.Append proc.CreateParameter("FRECHAFIN", adVarChar, adParamInput, 8, dFecFin)
            'proc.Parameters.Append proc.CreateParameter("lista1", adIUnknown, adParamOutput, 0)

            Set rs_Temp = New ADODB.Recordset
            Set rs_Temp = proc.Execute

    Else
            
            'Primeros pagos solo beneficiarios de sobrevivencia
            'Set rs_Temp = New ADODB.Recordset
            'rs_Temp.CursorLocation = adUseClient
            'rs_Temp.Open "PP_LISTA_PDT.LISTA_PDT('" & Txt_Anno.Text & "','" & Format(Txt_Mes.Text, "00") & "','" & Mid(Cmb_Tipo.Text, 1, 3) & "','" & IIf(OptBajas.Value = True, 1, 0) & "'," & lbl_tipcam.Caption & ")", vgConexionBD, adOpenStatic, adLockReadOnly

'RRR 08/08/2019
            Set proc = New ADODB.Command
            Set proc.ActiveConnection = vgConexionBD
            proc.CommandType = adCmdStoredProc
            proc.CommandText = "PP_LISTA_PDT.LISTA_PDT"
            proc.Prepared = False
            'proc.Parameters.Delete
            proc.Parameters.Append proc.CreateParameter("p_ANIO", adVarChar, adParamInput, 4, Txt_Anno.Text)
            proc.Parameters.Append proc.CreateParameter("p_MES", adVarChar, adParamInput, 2, Format(Txt_Mes.Text, "00"))
            proc.Parameters.Append proc.CreateParameter("p_TIPO", adVarChar, adParamInput, 3, Mid(Cmb_Tipo.Text, 1, 3))
            proc.Parameters.Append proc.CreateParameter("p_B_A", adVarChar, adParamInput, 1, IIf(OptBajas.Value = True, 1, 0))
            proc.Parameters.Append proc.CreateParameter("p_TC", adVarChar, adParamInput, 10, lbl_tipcam.Caption)
            'proc.Parameters.Append proc.CreateParameter("lista1", adIUnknown, adParamOutput, 0)

            Set rs_Temp = New ADODB.Recordset
            Set rs_Temp = proc.Execute
    End If
    
    
    'Dim a As Integer
    
    iCantidad = 0
    lvw_lista.ListItems.Clear
    If Not rs_Temp.EOF Then
        Do Until rs_Temp.EOF
            iCantidad = iCantidad + 1
            'If rs_Temp!numpol = "0000000792" Then
            '     a = 1
            'End If
            
            
            GoSub s_cabecera
            Select Case Mid(Cmb_Tipo.Text, 1, 3)
                Case "T00"
                    GoSub s_llena_lista_T00
                Case "T02"
                    GoSub s_llena_lista_T02
                Case "P00"
                    GoSub s_llena_lista_P00
                Case "PEN"
                    GoSub s_llena_lista_PEN
                Case "PER"
                    GoSub s_llena_lista_PER
            End Select
            rs_Temp.MoveNext
        Loop
    End If
    
    Exit Sub

s_cabecera:
'**********
    Set objItem = lvw_lista.ListItems.Add
    Set objSubItem = objItem.ListSubItems.Add(Text:=iCantidad)
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!numPol), "", Trim(rs_Temp!numPol)))
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!tipo), "", Trim(rs_Temp!tipo)))
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!tippen), "", Trim(rs_Temp!tippen)))

    
    Return
    
   
s_llena_lista_T00:
'*****************

    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!tipdoc), "", fTipoDoc(Trim(rs_Temp!tipdoc))))
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!NumDoc), "", Trim(rs_Temp!NumDoc)))
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!paisEmisor), "", Mid(Trim(rs_Temp!paisEmisor), 1, 3)))
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!FecNac), "", f_amd_dma(rs_Temp!FecNac)))
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!apepat), "", Mid(Trim(rs_Temp!apepat), 1, 40)))
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!apemat), "", Mid(Trim(rs_Temp!apemat), 1, 40)))
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!nomben1), "", Mid(Trim(rs_Temp!nomben1) & " " & IIf(IsNull(rs_Temp!nomben2), "", Trim(rs_Temp!nomben2)), 1, 40)))
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!Sexo), "", fSexo(Trim(rs_Temp!Sexo))))
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!nacionalidad), "", Mid(fNacionalidad(Trim(rs_Temp!nacionalidad)), 1, 4)))
    Set objSubItem = objItem.ListSubItems.Add(Text:="")
    Set objSubItem = objItem.ListSubItems.Add(Text:="")
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!correo), "", Mid(Trim(rs_Temp!correo), 1, 50)))
    
    If Trim(rs_Temp!tipdoc) = "1" Then

        Set objSubItem = objItem.ListSubItems.Add(Text:="99")
        'Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!direccion), " ", Trim(rs_Temp!direccion)))
        Set objSubItem = objItem.ListSubItems.Add(Text:=".")
        Set objSubItem = objItem.ListSubItems.Add(Text:="0")
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:="0")
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:="") 'Nombre de zona
        Set objSubItem = objItem.ListSubItems.Add(Text:="") 'Referencia
        Set objSubItem = objItem.ListSubItems.Add(Text:="") 'Ubigeo
    Else

        Set objSubItem = objItem.ListSubItems.Add(Text:="99")
        'Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!direccion), " ", Trim(rs_Temp!direccion)))
        Set objSubItem = objItem.ListSubItems.Add(Text:=".")
        Set objSubItem = objItem.ListSubItems.Add(Text:="0")
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:="0")
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:="") 'Nombre de zona
        Set objSubItem = objItem.ListSubItems.Add(Text:="") 'Referencia
        Set objSubItem = objItem.ListSubItems.Add(Text:="") 'Ubigeo
    End If
    
     'otra direccion registrada
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        'Indicador Centro Asistencial EsSalud
        Set objSubItem = objItem.ListSubItems.Add(Text:="1")
    
    
    Return

s_llena_lista_T02:
'*****************
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!tipdoc), "", fTipoDoc(Trim(rs_Temp!tipdoc))))
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!NumDoc), "", Mid(Trim(rs_Temp!NumDoc), 1, 15)))
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!paisEmisor), "", Mid(Trim(rs_Temp!paisEmisor), 1, 3)))
    Set objSubItem = objItem.ListSubItems.Add(Text:="24")
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!AFP), "", Mid(Trim(rs_Temp!AFP), 1, 2)))
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!Cuspp), "", Trim(rs_Temp!Cuspp)))
    Set objSubItem = objItem.ListSubItems.Add(Text:="2")
    
    Return
    
s_llena_lista_P00:
'*****************

    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!tipdoc), "", fTipoDoc(Trim(rs_Temp!tipdoc))))
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!NumDoc), "", Trim(rs_Temp!NumDoc)))
    Set objSubItem = objItem.ListSubItems.Add(Text:="") ' pais emisor
    Set objSubItem = objItem.ListSubItems.Add(Text:="2") 'Fijo siempre
    Set objSubItem = objItem.ListSubItems.Add(Text:="1") 'tipo de registro
    If optAltas.Value = True Then
'    If Trim(rs_Temp!numPol) = "0000004275" Then
'        MsgBox "LLEGO"
'    End If
        If Not IsNull(rs_Temp!FEC_EFECTO) Then
            Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!FEC_EFECTO), "", f_amd_dma(rs_Temp!FEC_EFECTO)))
        Else
            If (rs_Temp!Diferido > 0 And rs_Temp!Cod_TipRen <> "1") And (rs_Temp!Diferido > 0 And rs_Temp!Cod_TipRen <> "6") Then 'inmediatas
                dFecEval = Format(DateAdd("yyyy", rs_Temp!Diferido, f_amd_dma(Trim(rs_Temp!fecdev))), "YYYYMMDD")
                If dFecEval > rs_Temp!Emision Then
                    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!fecdev), "", f_amd_dma(dFecEval)))
                Else
                    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!Emision), "", "01" & Mid(f_amd_dma(rs_Temp!Emision), 3)))
                End If
            Else
                dFecEval = Trim(rs_Temp!fecdev)
                If dFecEval > rs_Temp!Emision Then
                    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!fecdev), "", f_amd_dma(dFecEval)))
                Else
                    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!Emision), "", "01" & Mid(f_amd_dma(rs_Temp!Emision), 3)))
                End If
            End If
        End If
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:="")  'Fijo siempre
    Else
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!FechaBaja), "", f_amd_dma(rs_Temp!FechaBaja)))
        Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!TipoSuspension), "", Trim(rs_Temp!TipoSuspension)))
        Set objSubItem = objItem.ListSubItems.Add(Text:="")  'Fijo siempre
    End If
    

    Return
    
s_llena_lista_PEN:
'*****************
    'Materia Gris - Jaime Rios 05/03/2018 inicio
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!CodAFP), "", Trim(rs_Temp!CodAFP)))
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!nom_AFP), "", Trim(rs_Temp!nom_AFP)))
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!Cuspp), "", Trim(rs_Temp!Cuspp)))
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!FecNac), "", f_amd_dma(rs_Temp!FecNac)))
    'Materia Gris - Jaime Rios 05/03/2018 fin
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!tipdoc), "", fTipoDoc(Trim(rs_Temp!tipdoc))))
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!NumDoc), "", Trim(rs_Temp!NumDoc)))
    Set objSubItem = objItem.ListSubItems.Add(Text:="0912") 'Fijo siempre
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!MontoDevengado), "0.00", Trim(rs_Temp!MontoDevengado)))
    'Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!Moneda), "", Trim(rs_Temp!Moneda)))
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!Pension), "0.00", Trim(rs_Temp!Pension)))
    'MVG 23/02/2016
'    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!Moneda), "0.00", IIf(rs_Temp!Moneda = "US", Format(rs_Temp!Pension * CDbl(lbl_tipcam.Caption), "####0"), Format(rs_Temp!Pension, "###0"))))
'    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!Moneda), "0.00", IIf(rs_Temp!Moneda = "US", Format(rs_Temp!Pension * CDbl(lbl_tipcam.Caption), "####0"), Format(rs_Temp!Pension, "###0"))))
    
    Return
    
s_llena_lista_PER:
'*****************
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!tipdoc), "", fTipoDoc(Trim(rs_Temp!tipdoc))))
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!NumDoc), "", Trim(rs_Temp!NumDoc)))
    Set objSubItem = objItem.ListSubItems.Add(Text:="0") 'Fijo siempre
    
    Return
End Sub

Private Sub p_llena_datos_der()

Dim rs_Temp As ADODB.Recordset
Dim rx_Temp As ADODB.Recordset
Dim sSql As String
Dim dFecIni As Long
Dim dFecFin As Long
Dim iCantidad As Long
Dim iEdad As Integer
Dim sPolizaAux As String
Dim iEndosoAux As Integer

    dFecIni = Txt_Anno.Text & Format(Txt_Mes.Text, "00") & "01"
    dFecFin = Txt_Anno.Text & Format(Txt_Mes.Text, "00") & f_dia_ultimo(Val(Txt_Anno.Text), Val(Txt_Mes.Text))

    'Información de Primeros Pagos: SOBREVIVENCIA
    '  No se muestra informacion de derecho habientes para sobrevivencia debido a que ellos ya estan cobrando de forma directa
    
    'Información de Primeros Pagos: DIFERENTES A SOBREVIVENCIA
    '  Se mostrará la informaciónde los datos de los beneficiarios que no son el titular
    'RVF 20090918      se adicionan campos adicionales a los querys
    
    'RVF 20100112  se puso en comentario la primera linea del query, ya que estaba tomando endosos, y esto hace que se
    '              duplique la información
    'sSql = " select distinct p.num_poliza as numpol, be.num_endoso as endoso, 'PP' as Tipo, p.cod_tippension as tippen, be.cod_tipoidenben as tipdoc, be.num_idenben as numdoc, be.gls_patben as apepat, be.gls_matben as apemat, be.gls_nomben as nomben1, be.gls_nomsegben as nomben2,"
    sSql = " select distinct p.num_poliza as numpol, 'PP' as Tipo, p.cod_tippension as tippen, be.cod_tipoidenben as tipdoc, be.num_idenben as numdoc, be.gls_patben as apepat, be.gls_matben as apemat, be.gls_nomben as nomben1, be.gls_nomsegben as nomben2,"
    sSql = sSql & " be.Fec_NacBen as FecNac, be.Cod_Sexo as Sexo, p.Cod_Cuspp as Cuspp, pr.Fec_Vigencia as Vigencia, pr.Mto_Pension as Pension, p.fec_dev as FecDev, p.Cod_Moneda as Moneda, be.Cod_Par as TipPar,"
    sSql = sSql & " p.gls_direccion as direccion, p.gls_fono as fono, p.gls_correo as correo, p.gls_nacionalidad as nacionalidad, p.cod_afp as AFP, "
    sSql = sSql & " p.COD_TIPOIDENAFI as tipdoctit, p.NUM_IDENAFI as numdoctit, "
    sSql = sSql & " p.fec_ingvigencia as Inicio, p.num_mesdif/12 as Diferido, p.fec_emision as Emision, "
    sSql = sSql & " p.cod_tipvia as TipoVia, p.gls_nomvia as NombreVia, p.gls_numdmc as NumVia, p.gls_intdmc as Interior,"
    sSql = sSql & " p.cod_tipzon as TipoZona, p.gls_nomzon as NomZona, p.gls_referencia as Referencia, p.cod_direccion as Ubigeo,"
    sSql = sSql & " p.Num_MesGar , p.fec_finpergar, be.cod_sitinv"
    '200912 DCM
    'sSql = sSql & " FROM pd_tmae_poliza p, pd_tmae_polprirec pr, ma_tpar_tabcod t, ma_tpar_tabcod r, ma_tpar_tabcod m, pd_tmae_polben be,"
    sSql = sSql & " FROM pd_tmae_poliza p, pd_tmae_polprirec pr, ma_tpar_tabcod t, ma_tpar_tabcod r, ma_tpar_tabcod m, pp_tmae_ben be,"
    sSql = sSql & " ma_tpar_tipoiden a, ma_tpar_cobercon b "
    sSql = sSql & " WHERE p.fec_pripago <= " & dFecFin & " AND p.num_poliza = pr.num_poliza"
    sSql = sSql & " AND p.num_poliza = be.num_poliza "
    'sSql = sSql & " AND p.num_endoso = be.num_endoso"  '200912 DCM   'RVF 20100112  se puso en comentario
    sSql = sSql & " AND t.cod_tabla = 'TP'"
    sSql = sSql & " AND be.Cod_Par <> '99'"
    sSql = sSql & " AND t.cod_elemento = p.cod_tippension AND r.cod_tabla = 'TR' AND r.cod_elemento = p.cod_tipren AND m.cod_tabla = 'AL'"
    sSql = sSql & " AND m.cod_elemento = p.cod_modalidad AND p.cod_tipoidenafi = a.cod_tipoiden And p.Cod_CoberCon = b.Cod_CoberCon"
    sSql = sSql & " AND p.Cod_TipPension <> '08'"
    'sSql = sSql & " AND p.num_poliza like '00%'" 'temporal
    sSql = sSql & " order by p.num_poliza " ', be.num_endoso desc"
    
    Set rs_Temp = New ADODB.Recordset
    Set rs_Temp = vgConexionBD.Execute(sSql)
    
    iCantidad = 0
    lvw_lista.ListItems.Clear
    If Not rs_Temp.EOF Then
        sPolizaAux = ""
        iEndosoAux = 0
        Do Until rs_Temp.EOF   '20100112 RVF   se puso en comentario las evaluaciones de pólizas y endosos, ya que
                               '               no son necesarias porque ya no se trabaja con el endoso
            'If sPolizaAux <> rs_Temp!numpol Then 'Si se cambia de poliza se evalua nuevamente
                GoSub s_evalua_pago
            'Else
            '    If iEndosoAux = rs_Temp!endoso Then
            '        GoSub s_evalua_pago
            '    'Else
            '    '    MsgBox "XX"
            '    End If
            'End If
            sPolizaAux = rs_Temp!numPol
            'iEndosoAux = rs_Temp!endoso
            rs_Temp.MoveNext
        Loop
    End If
    
    Exit Sub

s_evalua_pago:
'*************
    If rs_Temp!TipPar <> "30" Then
        GoSub s_busca_pago
    Else
        iEdad = Int(DateDiff("m", f_amd_dma(rs_Temp!FecNac), f_amd_dma(dFecIni)) / 12)
        If iEdad < 18 Then
            GoSub s_busca_pago
        Else
            If rs_Temp!Cod_SitInv <> "N" Then
                GoSub s_busca_pago
            'Else
            '    MsgBox "no pasa " & rs_Temp!apepat & " " & rs_Temp!apemat & " " & rs_Temp!Nomben1
            End If
        End If
    End If

    Return
    
s_busca_pago:
'************
    sSql = "select nvl(count(*),0) as cuenta from PP_TMAE_LIQPAGOPENDEF where num_poliza='" & rs_Temp!numPol & "'"
    'sSql = sSql & " and num_endoso=" & rs_Temp!endoso    '20100112 RVF   se puso en comentario
    sSql = sSql & " and num_idenreceptor='" & rs_Temp!NumDoc & "'"
    ssl = sSql & " and fec_pago >= '" & dFecIni & "' and fec_pago <= '" & dFecFin & "' and cod_tipopago = 'R'"
    Set rx_Temp = New ADODB.Recordset
    Set rx_Temp = vgConexionBD.Execute(sSql)
    If Not rx_Temp.EOF Then
        If rx_Temp!CUENTA = 0 Then
            iCantidad = iCantidad + 1
            GoSub s_cabecera
            GoSub s_llena_lista_DER
        End If
    End If
    
    Return

s_cabecera:
'**********
    Set objItem = lvw_lista.ListItems.Add
    Set objSubItem = objItem.ListSubItems.Add(Text:=iCantidad)
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!numPol), "", Trim(rs_Temp!numPol)))
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!tipo), "", Trim(rs_Temp!tipo)))
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!tippen), "", Trim(rs_Temp!tippen)))

    Return
    
    
s_llena_lista_DER:
'*****************
    Set objSubItem = objItem.ListSubItems.Add(Text:=fTipoDoc(rs_Temp!tipdoctit))
    Set objSubItem = objItem.ListSubItems.Add(Text:=rs_Temp!numdoctit)
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!tipdoc), "", fTipoDoc(Trim(rs_Temp!tipdoc))))
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!NumDoc), "", Trim(rs_Temp!NumDoc)))
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!apepat), "", Trim(rs_Temp!apepat)))
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!apemat), "", Trim(rs_Temp!apemat)))
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!nomben1), "", Trim(rs_Temp!nomben1)) & " " & IIf(IsNull(rs_Temp!nomben2), "", Trim(rs_Temp!nomben2)))
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!FecNac), "", f_amd_dma(rs_Temp!FecNac)))
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!Sexo), "", fSexo(Trim(rs_Temp!Sexo))))
    Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!TipPar), "", fVinculo(Trim(rs_Temp!TipPar))))
    Set objSubItem = objItem.ListSubItems.Add(Text:="")
    Set objSubItem = objItem.ListSubItems.Add(Text:="")
    Set objSubItem = objItem.ListSubItems.Add(Text:="10")
    If rs_Temp!Diferido > 0 Then
        dFecEval = Format(DateAdd("yyyy", rs_Temp!Diferido, f_amd_dma(Trim(rs_Temp!fecdev))), "YYYYMMDD")
        'Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!fecdev), "", f_amd_dma(Trim(rs_Temp!Inicio))))
        If dFecEval > rs_Temp!Emision Then
            Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!fecdev), "", f_amd_dma(dFecEval)))
        Else
            Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!Emision), "", "01" & Mid(f_amd_dma(rs_Temp!Emision), 3)))
        End If
    Else
        'Set objSubItem = objItem.ListSubItems.Add(Text:="01" & Mid(f_amd_dma(rs_Temp!Emision), 3))
        dFecEval = Trim(rs_Temp!fecdev)
        If dFecEval > rs_Temp!Emision Then
            Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!fecdev), "", f_amd_dma(dFecEval)))
        Else
            Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!Emision), "", "01" & Mid(f_amd_dma(rs_Temp!Emision), 3)))
        End If
    End If
    Set objSubItem = objItem.ListSubItems.Add(Text:="")
    Set objSubItem = objItem.ListSubItems.Add(Text:="")
    Set objSubItem = objItem.ListSubItems.Add(Text:="")

    Set objSubItem = objItem.ListSubItems.Add(Text:="0")    'Indicador de domicialiado por defecto
    'Estos valores solo se llenan si es menor de edad o tipo de documento con carnet de extranjeria, pero en POL_BENEF no existen estos datos los saca del titular
    
    'RVF 20090918
    If Trim(rs_Temp!tipdoc) = "2" Or Trim(rs_Temp!tipdoc) = "6" Then
        Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!TipoVia), "", Trim(rs_Temp!TipoVia)))
        Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!NombreVia), "", Left(Trim(rs_Temp!NombreVia), 20)))
        Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!NumVia), "", Trim(rs_Temp!NumVia)))
        Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!Interior), "", Trim(rs_Temp!Interior)))
        Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!TipoZona), "", Trim(rs_Temp!TipoZona)))
        Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!NomZona), "", Left(Trim(rs_Temp!NomZona), 20)))
        Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!Referencia), "", Left(Trim(rs_Temp!Referencia), 40)))
        Set objSubItem = objItem.ListSubItems.Add(Text:=IIf(IsNull(rs_Temp!Ubigeo), "", Format$(fBuscaUbigeo(rs_Temp!Ubigeo), "000000")))
    Else
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
        Set objSubItem = objItem.ListSubItems.Add(Text:="")
    End If
    '*************
    
    Return
    
End Sub

Private Sub Txt_Anno_GotFocus()
    Call p_sombrea_texto(Txt_Anno)
End Sub

Private Sub Txt_Anno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = f_valida_numeros(KeyAscii)
    End If
End Sub

Private Sub Txt_Mes_GotFocus()
    Call p_sombrea_texto(Txt_Mes)
End Sub

Private Function fTipoDoc(ByVal sTipDoc) As String

    fTipoDoc = ""
    Select Case sTipDoc
        Case "1"  'DNI
            fTipoDoc = "01"
        Case "2"  'Carnet extranjeria
            fTipoDoc = "04"
        Case "5"  'Pasaporte
            fTipoDoc = "07"
        Case "6"  'Partida Nacimiento
            fTipoDoc = "11"
    End Select

End Function

Private Function fSexo(ByVal sSexo) As String
    
    fSexo = ""
    Select Case sSexo
        Case "M"
            fSexo = "1"
        Case "F"
            fSexo = "2"
    End Select

End Function

Private Function fNacionalidad(ByVal sNacionalidad) As String
    
    fNacionalidad = ""
    Select Case sNacionalidad
        Case "PERUANA"
            fNacionalidad = "9589"
    End Select

End Function

Private Function fBuscaUbigeo(pCodigo)
'RVF  20090918
Dim vlCodComuna

On Error GoTo Err_Ubigeo

    vlCodComuna = 0
    vgSql = ""
    vgSql = "SELECT * FROM MA_TPAR_COMUNA "
    vgSql = vgSql & " Where cod_direccion=" & pCodigo
    Set vgRs = vgConexionBD.Execute(vgSql)
        
    If Not (vgRs.EOF) Then
        vlCodComuna = IIf(IsNull(vgRs!cod_comuna), "", vgRs!cod_comuna)
    End If
    vgRs.Close
    
    If vlCodComuna <> "" Then
        vgSql = ""
        vgSql = "SELECT * FROM MA_TVAL_EQUIVUBIGEO"
        vgSql = vgSql & " Where num_codsistema=" & vlCodComuna
        Set vgRs = vgConexionBD.Execute(vgSql)
            
        If Not (vgRs.EOF) Then
            fBuscaUbigeo = IIf(IsNull(vgRs!num_codsbs), "", vgRs!num_codsbs)
        End If
    End If
    
Exit Function

Err_Ubigeo:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select

End Function

Private Function fAFP(ByVal sAFP) As String
    
    fAFP = ""
    Select Case sAFP
        Case "242"  'Integra
            fAFP = "21"
        Case "241"  'Horizonte
            fAFP = "22"
        Case "243"  'Profuturo
            fAFP = "23"
        Case "245"  'Prima
            fAFP = "24"
        Case "244"  'Union Vida
            fAFP = "24"
    End Select

End Function

Private Function fVinculo(ByVal sVinculo) As String

    fVinculo = ""
    Select Case sVinculo
        Case "30"               'Hijo
            fVinculo = "1"
        Case "10", "11"         'Conyugue
            fVinculo = "2"
        Case "20", "21"         'Concubina
            fVinculo = ""
        Case ""                 'Gestante
            fVinculo = ""
        Case "41", "42"         'Padres
            fVinculo = "2"
    End Select

End Function

Private Sub Txt_Mes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'SendKeys "{TAB}"
    Else
        KeyAscii = f_valida_numeros(KeyAscii)
    End If
End Sub

Private Function fTipoCambioSBS(ByVal sCodMon As String, ByVal dFecha As String) As Double

Dim rs_Temp As ADODB.Recordset
Dim Sql As String
    'Para tipo de cambio del mes
    Sql = "SELECT mto_moneda FROM MA_TVAL_MONEDA_SBS WHERE "
    Sql = Sql & "cod_moneda = '" & sCodMon & "' AND "
    Sql = Sql & "fec_moneda = '" & dFecha & "'"
    Set rs_Temp = New ADODB.Recordset
    Set rs_Temp = vgConexionBD.Execute(Sql)
    If rs_Temp.EOF Then
        fTipoCambioSBS = 0
    Else
        fTipoCambioSBS = rs_Temp!Mto_Moneda
    End If

End Function


VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frm_PlanillaSVSVta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Archivo de Pólizas vendidas en el Periodo."
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   6195
   Begin MSComDlg.CommonDialog ComDialogo 
      Left            =   4920
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   1035
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   5895
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   3600
         Picture         =   "Frm_PlanillaSVSVta.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   2640
         Picture         =   "Frm_PlanillaSVSVta.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Cargar 
         Caption         =   "&Archivo"
         Height          =   675
         Left            =   1680
         Picture         =   "Frm_PlanillaSVSVta.frx":07B4
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Exportar Datos a Archivo"
         Top             =   240
         Width           =   720
      End
      Begin Crystal.CrystalReport Rpt_Reporte 
         Left            =   5280
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.Frame Fra_Busqueda 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5925
      Begin VB.TextBox Txt_Anno 
         Height          =   285
         Left            =   3600
         MaxLength       =   4
         TabIndex        =   2
         Top             =   420
         Width           =   795
      End
      Begin VB.TextBox Txt_Mes 
         Height          =   285
         Left            =   2880
         MaxLength       =   2
         TabIndex        =   1
         Top             =   420
         Width           =   360
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
         Left            =   2400
         TabIndex        =   9
         Top             =   465
         Width           =   165
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
         TabIndex        =   8
         Top             =   420
         Width           =   705
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
         TabIndex        =   7
         Top             =   435
         Width           =   1125
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
         Left            =   3315
         TabIndex        =   6
         Top             =   435
         Width           =   195
      End
   End
End
Attribute VB_Name = "Frm_PlanillaSVSVta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'CMV/20050331

Dim vlCodEstado As String

Dim vlRegistroPagos As ADODB.Recordset

Dim vlOpcion As String
Dim vlGlosaOpcion As String
Dim vlArchivo As String
Dim vlNumPerPago As String

Dim vlMtoPensionUF As Double
Dim vlMtoPromedioUF As Double
Dim vlNumPensiones As Integer

'Variables para Formulas Informe
'Constantes de Tipos de Pensión
Const clCodTipPension04 As String = "('04')" 'Vejez Normal
Const clCodTipPension05 As String = "('05')" 'Vejez Anticipada
Const clCodTipPension06 As String = "('06')" 'Invalidez Total
Const clCodTipPension07 As String = "('07')" 'Invalidez Parcial
Const clCodTipPension08 As String = "('08','09','10','11','12')" 'Sobrevivencia
'Constantes de Tipo de Parentesco
Const clCodParPensionado As String = "('99')" 'Causante
Const clCodParConyuge As String = "('10','11')"  'Conyuge
Const clCodParHijos As String = "('30','35')"  'Hijos
Const clCodParOtros As String = "('20','21','41','42','77')"  'Otros

Const clCodPagoPen As String * 2 = "01"
Const clCodTipRecR As String * 1 = "R"
Const clCodTipoPago As String * 1 = "P"

'I--- ABV 04/12/2006 ---
Dim vlA As String, vlB As String
Dim vlRegistro       As ADODB.Recordset
Dim vlLinea          As String

Dim vlNumPerReserva As String
Dim vlFechaReserva  As String
Dim vlNumPerReservaMes As Integer
Dim vlNumPerReservaAno As Long

Const clCodCliente As Integer = 1
Const clCodEstadoReservaCerrado As String * 1 = "C"
Const clCodEstadoPolizaScomp As String * 1 = "S"

Dim vlNumPoliza       As String
Dim vlRutAfiliado     As String, vlDgVAfiliado As String
Dim vlCodTipPension   As String
Dim vlCodTipModalidad As String
Dim vlCodTipRenta     As String
Dim vlMtoRenta        As String, vlMtoRentaTexto As String
Dim vlMtoPrima        As String, vlMtoPrimaTexto As String
Dim vlTasaVenta       As String, vlTasaVentaTexto As String
Dim vlCodTipCorredor  As String
Dim vlRutCorredor     As String, vlDgVCorredor As String
Dim vlPrcCorredor     As String, vlPrcCorredorTexto As String
Dim vlCodTramitadaScomp As String

Dim vlCodTipPensionSVS   As String
'F--- ABV 04/12/2006 ---

Private Sub Cmd_Cargar_Click()
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

    vlNumPerPago = Format(Txt_Anno.Text, "0000") & Format(Txt_Mes.Text, "00")
    vlFechaReserva = Format(DateSerial(CLng(Txt_Anno), CInt(Txt_Mes), 1 - 1), "yyyymmdd")
    vlNumPerReserva = Mid(vlFechaReserva, 1, 6)
    vlNumPerReservaMes = Mid(vlFechaReserva, 5, 2)
    vlNumPerReservaAno = Mid(vlFechaReserva, 1, 4)
    
    'Valida si el periodo de Reservas a sido generado
    If flValidaEstadoProceso(Trim(vlFechaReserva), clCodEstadoReservaCerrado) = False Then
        MsgBox "El Tipo de Proceso Seleccionado no se encuentra Cerrado o no ha sido generado para el periodo : " & vlNumPerReservaMes & "-" & vlNumPerReservaAno & ".", vbCritical, "Error de Datos"
        Exit Sub
    End If
    'CMV-20060222 F

    Screen.MousePointer = 11

    'Selección del Archivo de Resumen de Reservas
    ComDialogo.CancelError = True
    ComDialogo.FileName = "R" & vlNumPerPago & ".txt"
    ComDialogo.DialogTitle = "Guardar Pólizas de Rtas. Vit. del Mes como"
    ComDialogo.Filter = "*.txt"
    ComDialogo.FilterIndex = 1
    ComDialogo.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    ComDialogo.ShowSave
    vlArchivo = ComDialogo.FileName
    If vlArchivo = "" Then
        Screen.MousePointer = 0
        Exit Sub
    End If

    'Permite generar el Archivo con la Opción Indicada a través del Menú
    Call flArchivoAnexoSVS

    Screen.MousePointer = 0

Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    If Err.Number = 32755 Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
        End If
    End If
End Sub

Private Sub Cmd_Limpiar_Click()
On Error GoTo Err_Limpiar

    Call flLimpiarVariables

    Txt_Mes = ""
    Txt_Anno = ""
    Txt_Mes.SetFocus

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

    Me.Left = 0
    Me.Top = 0

Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Function flLimpiarVariables()
On Error GoTo Err_flLimpiarVariables

   vlMtoPensionUF = 0
   vlMtoPromedioUF = 0
   vlNumPensiones = 0

Exit Function
Err_flLimpiarVariables:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Private Sub Txt_Anno_Change()
If Not IsNumeric(Txt_Anno) Then
   Txt_Anno = ""
End If
End Sub

Private Sub Txt_Anno_KeyPress(KeyAscii As Integer)

On Error GoTo Err_TxtAnnoKeyPress

    If KeyAscii = 13 Then
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
        Txt_Anno.Text = Format(Txt_Anno.Text, "0000")
        Cmd_Cargar.SetFocus
     End If

Exit Sub
Err_TxtAnnoKeyPress:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select


End Sub

Private Sub Txt_Anno_LostFocus()

    If Txt_Anno.Text = "" Then
       Exit Sub
    End If
    If CDbl(Txt_Anno.Text) < 1900 Then
       Exit Sub
    End If
    Txt_Anno.Text = Format(Txt_Anno.Text, "0000")


End Sub

Private Sub Txt_Mes_Change()
If Not IsNumeric(Txt_Mes) Then
   Txt_Mes = ""
End If
End Sub

Private Sub Txt_Mes_KeyPress(KeyAscii As Integer)

On Error GoTo Err_TxtMesKeyPress

    If KeyAscii = 13 Then
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
        Txt_Mes.Text = Format(Txt_Mes.Text, "00")
        Txt_Anno.SetFocus
    End If

Exit Sub
Err_TxtMesKeyPress:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select


End Sub

Private Sub Txt_Mes_LostFocus()

    If Txt_Mes.Text = "" Then
       Exit Sub
    End If
    If CDbl(Txt_Mes.Text) <= 0 Or CDbl(Txt_Mes.Text) > 12 Then
       Exit Sub
    End If
    Txt_Mes.Text = Format(Txt_Mes.Text, "00")

End Sub

Function flValidaEstadoProceso(iPeriodo As String, iCodEstado As String) As Boolean
'Valida el Estado del Periodo de Reservas para obtener la información
On Error GoTo Err_flValidaTipoProceso

    flValidaEstadoProceso = False

    vgSql = "SELECT cod_cliente,fec_calculo,num_periodo,cod_estperiodo "
    vgSql = vgSql & "FROM pr_tmae_procaldef "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "cod_cliente = 1 AND "
    vgSql = vgSql & "fec_calculo = '" & Trim(iPeriodo) & "' "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
        If Trim(vgRegistro!cod_estperiodo) = clCodEstadoReservaCerrado Then
            flValidaEstadoProceso = True
        End If
    End If
    vgRegistro.Close

Exit Function
Err_flValidaTipoProceso:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flArchivoAnexoSVS()
Dim vlNumRegistros As Integer
Dim vlAumento As Double
Dim vlFiller1 As String, vlFiller2 As String
Dim vlOpen As Boolean
On Error GoTo Err_flInformeAnexoSVS

    vlFiller1 = Space(37)
    vlFiller2 = Space(88)
    vgI = 0

    vlNumRegistros = 0

    Open vlArchivo For Output As #1
    vlOpen = True

    'Registro Tipo 1
    'vlLinea = "1" & vlNumPerPago & Format(vgRutCompania, "000000000") & _
              vgDgVCompania & vgNombreCompania & Space(80 - Len(vgNombreCompania))
    vlLinea = "1" & vlNumPerPago & vgNumIdenCompania & _
              vgTipoIdenCompania & vgNombreCompania & Space(80 - Len(vgNombreCompania))
    Print #1, vlLinea
    vgI = vgI + 1


    'Registro Tipo 2
    'Números de Pólizas que se vendieron en el Mes anterior al Mes indicado
    vgSql = "SELECT "
    vgSql = vgSql & "num_poliza,cod_plan,cod_tipren,cod_modalidad,  "
    vgSql = vgSql & "fec_vigencia,mto_prima,mto_pension,num_mesdif,num_mesgar,"
    vgSql = vgSql & "prc_tasavta,rut_afi,dgv_afi "
    vgSql = vgSql & "FROM pr_this_poliza1 "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "fec_calculo = '" & vlFechaReserva & "' "
    vgSql = vgSql & "AND cod_cliente = " & clCodCliente & " "
    vgSql = vgSql & "AND fec_vigencia <= '" & vlFechaReserva & "' "
    vgSql = vgSql & "AND fec_vigencia >= '" & Mid(vlFechaReserva, 1, 6) & "01" & "' "
    vgSql = vgSql & "order by num_poliza "
    Set vlRegistro = vgConexionBD.Execute(vgSql)
    While Not vlRegistro.EOF
        
        vlNumPoliza = vlRegistro!Num_Poliza
        vlRutAfiliado = Format(vlRegistro!rut_afi, "000000000")
        vlDgVAfiliado = vlRegistro!dgv_afi
        vlCodTipPension = vlRegistro!cod_plan
        vlCodTipPensionSVS = "00"
        
        'Definir Tipo de Pensión SVS
        If (vlRegistro!Cod_TipRen = "1") Then
            'Si es Inmediata
            Select Case vlRegistro!cod_plan
                Case "04": vlCodTipPensionSVS = "30" 'Vejez Normal
                Case "05": vlCodTipPensionSVS = "35" 'Vejez Anticipada
                Case "06": vlCodTipPensionSVS = "80" 'Inv. Parcial
                Case "07": vlCodTipPensionSVS = "40" 'Inv. Total
                Case "08": vlCodTipPensionSVS = "70" 'Sobrevivencia
            End Select
        Else
            'Si es Diferida
            Select Case vlRegistro!cod_plan
                Case "04": vlCodTipPensionSVS = "38" 'Vejez Normal
                Case "05": vlCodTipPensionSVS = "39" 'Vejez Anticipada
                Case "06": vlCodTipPensionSVS = "88" 'Inv. Parcial
                Case "07": vlCodTipPensionSVS = "48" 'Inv. Total
                Case "08": vlCodTipPensionSVS = "78" 'Sobrevivencia
            End Select
        End If
        
        vlCodTipModalidad = vlRegistro!Cod_Modalidad & Format(vlRegistro!Num_MesGar, "000")
        vlCodTipRenta = vlRegistro!Cod_TipRen & Format(vlRegistro!Num_MesDif, "000")
        
        vlMtoRenta = vlRegistro!Mto_Pension
        vlMtoRentaTexto = Replace(Format(vlMtoRenta, "000.00"), ",", ".")
        vlA = Mid(vlMtoRentaTexto, 1, 3)
        vlB = Mid(vlMtoRentaTexto, 5, 2)
        vlMtoRentaTexto = vlA & vlB
        
        vlMtoPrima = vlRegistro!Mto_Prima
        vlMtoPrimaTexto = Replace(Format(vlMtoPrima, "00000.00"), ",", ".")
        vlA = Mid(vlMtoPrimaTexto, 1, 5)
        vlB = Mid(vlMtoPrimaTexto, 7, 2)
        vlMtoPrimaTexto = vlA & vlB
        
        vlTasaVenta = vlRegistro!Prc_TasaVta
        vlTasaVentaTexto = Replace(Format(vlTasaVenta, "000.00"), ",", ".")
        vlA = Mid(vlTasaVentaTexto, 1, 3)
        vlB = Mid(vlTasaVentaTexto, 5, 2)
        vlTasaVentaTexto = vlA & vlB
        
        vlCodTramitadaScomp = clCodEstadoPolizaScomp
        
        If (flBuscarDatosIntermediario(vlNumPoliza, vlRutCorredor, vlDgVCorredor, _
        vlPrcCorredor, vlCodTipCorredor) = False) Then
            vlCodTipCorredor = Space(1)
            vlRutCorredor = Format(0, "000000000")
            vlDgVCorredor = Space(1)
            vlPrcCorredor = Format(0, "00.00")
            vlPrcCorredorTexto = "0000"
        Else
            vlRutCorredor = Format(vlRutCorredor, "000000000")
            vlPrcCorredor = Format(vlPrcCorredor, "00.00")
            vlPrcCorredorTexto = Replace(Format(vlPrcCorredor, "00.00"), ",", ".")
            vlA = Mid(vlPrcCorredorTexto, 1, 2)
            vlB = Mid(vlPrcCorredorTexto, 4, 2)
            vlPrcCorredorTexto = vlA & vlB
        End If
        
        vlLinea = "2" & vlRutAfiliado & vlDgVAfiliado & vlNumPoliza & _
            vlCodTipPensionSVS & vlCodTipModalidad & _
            vlMtoRentaTexto & vlMtoPrimaTexto & vlTasaVentaTexto & _
            vlCodTipCorredor & vlRutCorredor & vlDgVCorredor & _
            vlPrcCorredorTexto & vlCodTramitadaScomp & _
            vlFiller1
        Print #1, vlLinea
        vgI = vgI + 1
        
        vlRegistro.MoveNext
    Wend
    vlRegistro.Close
    
    'Registro Tipo 3
    vgI = vgI + 1
    vlLinea = "3" & Format(vgI, "00000000") & vlFiller2
    Print #1, vlLinea

    Close #1

    vlOpen = False
    MsgBox "La Exportación de datos al Archivo ha sido finalizada exitosamente.", vbInformation, "Estado de Generación Archivo"

    Screen.MousePointer = 0

Exit Function
Err_flInformeAnexoSVS:
    Screen.MousePointer = 0
    If vlOpen Then
        Close #1
    End If
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flBuscarDatosIntermediario(iNumPoliza As String, _
oRutCorredor As String, oDgVCorredor As String, oPrcCorredor As String, oCodTipCorredor As String) As Boolean

On Error GoTo Err_BuscarIntermediario

    flBuscarDatosIntermediario = False
    oRutCorredor = ""
    oDgVCorredor = ""
    oPrcCorredor = ""
    oCodTipCorredor = ""

    vgSql = "SELECT p.num_poliza,p.rut_corredor,p.dgv_corredor,p.prc_corcom, "
    vgSql = vgSql & "p.mto_corcom,c.cod_corredor "
    vgSql = vgSql & "FROM pd_tmae_poliza p, pt_tmae_corredor c "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "p.Num_Poliza = '" & iNumPoliza & "' "
    vgSql = vgSql & " AND p.rut_corredor = c.rut_cor "
    vgSql = vgSql & "order by p.num_endoso "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
        If Not IsNull(vgRs!rut_corredor) Then oRutCorredor = vgRs!rut_corredor
        If Not IsNull(vgRs!dgv_corredor) Then oDgVCorredor = Trim(vgRs!dgv_corredor)
        If Not IsNull(vgRs!prc_corcom) Then oPrcCorredor = vgRs!prc_corcom
        If Not IsNull(vgRs!cod_corredor) Then oCodTipCorredor = Trim(vgRs!cod_corredor)
        
        flBuscarDatosIntermediario = True
    End If
    vgRs.Close

Exit Function
Err_BuscarIntermediario:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        'MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

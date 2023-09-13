VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "Crystl32.OCX"
Begin VB.Form Frm_MenMasivo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de Mensajes Automáticos - Masivos."
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   8235
   Begin VB.Frame Frame2 
      Caption         =   "  Periodo de Pago  "
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
      Height          =   855
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   7995
      Begin VB.TextBox Txt_Anno 
         Height          =   285
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Txt_Mes 
         Height          =   285
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   1
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "(Mes-Año)"
         Height          =   255
         Left            =   3840
         TabIndex        =   19
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "-"
         Height          =   255
         Left            =   2520
         TabIndex        =   18
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label2 
         Caption         =   "Periodo de Pago"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Fra_Botones 
      Height          =   975
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   7965
      Begin VB.CommandButton Cmd_Grabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   1560
         Picture         =   "Frm_MenMasivo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   3720
         Picture         =   "Frm_MenMasivo.frx":06BA
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   4800
         Picture         =   "Frm_MenMasivo.frx":0D74
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   3720
         Picture         =   "Frm_MenMasivo.frx":0E6E
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   200
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Eliminar 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   2640
         Picture         =   "Frm_MenMasivo.frx":1528
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Eliminar Año"
         Top             =   200
         Width           =   720
      End
      Begin Crystal.CrystalReport Rpt_Calculo 
         Left            =   7200
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "  Asignación de Mensajes Automáticos  "
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
      Height          =   1605
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   7995
      Begin VB.TextBox Txt_CerEst 
         Height          =   285
         Left            =   6000
         MaxLength       =   10
         TabIndex        =   8
         Top             =   1155
         Width           =   1305
      End
      Begin VB.TextBox Txt_BonInv 
         Height          =   285
         Left            =   6000
         MaxLength       =   10
         TabIndex        =   6
         Top             =   810
         Width           =   1305
      End
      Begin VB.TextBox Txt_GarEst 
         Height          =   285
         Left            =   6000
         MaxLength       =   10
         TabIndex        =   4
         Top             =   480
         Width           =   1305
      End
      Begin VB.CheckBox Chk_CerEst 
         Caption         =   "Navidad"
         Height          =   270
         Left            =   450
         TabIndex        =   7
         Top             =   1185
         Width           =   3945
      End
      Begin VB.CheckBox Chk_BonInv 
         Caption         =   "Fiestas Patrias"
         Height          =   315
         Left            =   450
         TabIndex        =   5
         Top             =   810
         Width           =   3975
      End
      Begin VB.CheckBox Chk_GarEst 
         Caption         =   "Cumpleaños"
         Height          =   270
         Left            =   450
         TabIndex        =   3
         Top             =   480
         Width           =   3930
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Mes de Pago"
         Height          =   210
         Left            =   5865
         TabIndex        =   14
         Top             =   210
         Width           =   1530
      End
   End
End
Attribute VB_Name = "Frm_MenMasivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vlFechaCalculo As String
Dim vlFecIniTramo As String
Dim vlFecTerTramo As String
Dim vlDia As String
Dim vlMes As String
Dim vlAnno As String
Dim vlAnnoIniTramo As String
Dim vlAnnoTerTramo As String
Dim vlNumPerPago As String
Dim vlNumPolizaAnt As String

Dim vlPeriodo As String

Dim vlTiposPensionSob As String

Dim vlCodMenGarEst As Integer
Dim vlCodMenCerEst As Integer
Dim vlCodMenAsiFam As Integer
Dim vlCodMEnBonInv As Integer

Dim vlSwGarEst As String * 1
Dim vlSwBonInv As String * 1
Dim vlSwCerEst As String * 1
Dim vlSwAsiFam As String * 1

Dim vlGlsUsuarioCrea As Variant
Dim vlFecCrea As Variant
Dim vlHorCrea As Variant
Dim vlGlsUsuarioModi As Variant
Dim vlFecModi As Variant
Dim vlHorModi As Variant

Dim vlMtoPensionUF As Double
Dim vlMtoPensionPesos As Double
Dim vlMtoPensionMin As Double
Dim vlEdadBen As Integer
Dim vlMtoMoneda As Double

Const clCodNoInv As String * 1 = "N"
Const clCodActiva As String * 1 = "A"
Const clCodTipoIng As String * 1 = "M"
Const clCodTipoResGE As String = "('09','10','11','20','21')"

Const clCodSexoM As String * 1 = "M"
Const clCodSexoF As String * 1 = "F"
Const clEdadTope65 As Integer = 65 'Edad Tope para Hombres
Const clEdadTope60 As Integer = 60 'Edad Tope para Mujeres


Function flAsignaMensajesGarEst()
On Error GoTo Err_flAsignaMensajesGarEst

    vlSwGarEst = "N"

    vlFechaCalculo = Format(CDate(Trim(Txt_GarEst.Text)), "yyyymmdd")
        
'Se seleccionan las polizas con resolucion de G.E. vigente a la fecha
        
    vgSql = ""
    vgSql = "SELECT g.num_poliza,g.num_endoso,g.num_orden "
    vgSql = vgSql & "FROM pp_tmae_garestres g "
    vgSql = vgSql & "WHERE cod_tipres IN " & clCodTipoResGE & " AND "
    vgSql = vgSql & "g.fec_inires <= '" & vlFechaCalculo & "' AND "
    vgSql = vgSql & "g.fec_terres >= '" & vlFechaCalculo & "' "
    vgSql = vgSql & "ORDER BY g.fec_inires "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
    
       vlSwGarEst = "S"
    
       vlNumPerPago = Format(Txt_Anno.Text, "0000") & Format(Txt_Mes.Text, "00")
       vlCodMenGarEst = 0
       vgSql = ""
       vgSql = "SELECT cod_mensajegarest "
       vgSql = vgSql & "FROM MA_TCOD_GENERAL "
       Set vgRs2 = vgConexionBD.Execute(vgSql)
       If Not vgRs2.EOF Then
          vlCodMenGarEst = (vgRs2!cod_mensajegarest)
       End If
       vlNumPolizaAnt = ""
       
'Se envian mensajes generados por concepto de garantia estatal
'a el/los beneficiarios (con Certificado) Solo a el pensionado.
       
       While Not vgRs.EOF
             If Trim(vlNumPolizaAnt) <> (vgRs!num_poliza) Then
                Call flInsertarMensajes((vgRs!num_poliza), (vgRs!num_endoso), _
                                        (vgRs!Num_Orden), vlCodMenGarEst)
             End If
             vlNumPolizaAnt = (vgRs!num_poliza)
             vgRs.MoveNext
       Wend
       MsgBox "Los Datos para Mensajes por Garantía Estatal, han sido Ingresados Satisfactoriamente.", vbInformation, "Información"
    End If
    
'    MsgBox "El Proceso de Asignación de Mensajes por Garantía Estatal No se encuentra Disponible.", vbInformation, "Información"
    
Exit Function
Err_flAsignaMensajesGarEst:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
    
End Function

Function flAsignaMensajesBonInv()

    vlSwBonInv = "N"

    vlFechaCalculo = Format(CDate(Trim(Txt_BonInv.Text)), "yyyymmdd")
    
    'buscar monto de moneda
    vlMtoMoneda = 0
    vgSql = ""
    vgSql = "SELECT mto_moneda "
    vgSql = vgSql & "FROM ma_tval_moneda "
    vgSql = vgSql & "WHERE cod_moneda = '" & Trim(cgCodTipMonedaUF) & "' AND "
    vgSql = vgSql & "fec_moneda = '" & Trim(vlFechaCalculo) & "' "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
        vlMtoMoneda = (vgRegistro!Mto_Moneda)
    Else
        MsgBox "Debe Ingresar el Valor de UF para la Fecha Indicada.", vbInformation, "Información"
        Exit Function
    End If

        
'Se seleccionan las polizas con resolucion de G.E. vigente a la fecha
        
    vgSql = ""
    vgSql = "SELECT g.num_poliza,g.num_endoso,g.num_orden, "
    vgSql = vgSql & "b.cod_par,b.cod_sexo,b.fec_nacben,b.cod_sitinv, "
    vgSql = vgSql & "b.mto_pension,b.mto_pensiongar,b.fec_terpagopengar "
    vgSql = vgSql & "FROM pp_tmae_garestres g, pp_tmae_ben b "
    vgSql = vgSql & "WHERE cod_tipres IN " & clCodTipoResGE & " AND "
    vgSql = vgSql & "g.fec_inires <= '" & vlFechaCalculo & "' AND "
    vgSql = vgSql & "g.fec_terres >= '" & vlFechaCalculo & "' AND "
    vgSql = vgSql & "g.num_poliza = b.num_poliza AND "
    vgSql = vgSql & "g.num_endoso = b.num_endoso AND "
    vgSql = vgSql & "g.num_orden = b.num_orden "
    vgSql = vgSql & "ORDER BY g.fec_inires "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
    
       vlSwBonInv = "S"
    
       vlNumPerPago = Format(Txt_Anno.Text, "0000") & Format(Txt_Mes.Text, "00")
       vlCodMEnBonInv = 0
       vgSql = ""
       vgSql = "SELECT cod_mensajeboninv "
       vgSql = vgSql & "FROM MA_TCOD_GENERAL "
       Set vgRs2 = vgConexionBD.Execute(vgSql)
       If Not vgRs2.EOF Then
          vlCodMEnBonInv = (vgRs2!cod_mensajeboninv)
       End If
       vlNumPolizaAnt = ""
       
       While Not vgRs.EOF
       
            vlMtoPensionUF = 0
            vlMtoPensionPesos = 0
            vlMtoPensionMin = 0
            vlEdadBen = 0
            'Seleccion Monto pension normal o monto pension garantizada (en UF)
            If IsNull(vgRs!Fec_TerPagoPenGar) Then
               vlMtoPensionUF = (vgRs!Mto_Pension)
            Else
                If vlFechaCalculo > (vgRs!Fec_TerPagoPenGar) Then
                   vlMtoPensionUF = (vgRs!Mto_Pension)
                Else
                    If vlFechaCalculo < (vgRs!Fec_TerPagoPenGar) Then
                       vlMtoPensionUF = (vgRs!Mto_PensionGar)
                    End If
                End If
            End If
            'Calcular edad del beneficiario
            vlEdadBen = fgCalculaEdad((vgRs!Fec_NacBen), fgBuscaFecServ)
            vlEdadBen = fgConvierteEdadAños(vlEdadBen)
            
            vlMtoPensionPesos = vlMtoPensionUF * vlMtoMoneda
            
            'buscar monto de pension minima para el beneficiario
            vgSql = ""
            vgSql = "SELECT mto_penminfin "
            vgSql = vgSql & "FROM pp_tval_penminima "
            vgSql = vgSql & "WHERE fec_inipenmin <= '" & vlFechaCalculo & " ' AND "
            vgSql = vgSql & "cod_par = '" & (vgRs!Cod_Par) & "' AND "
            vgSql = vgSql & "cod_sitinv = '" & (vgRs!Cod_SitInv) & "' AND "
            vgSql = vgSql & "cod_sexo = '" & (vgRs!Cod_Sexo) & "' AND "
            vgSql = vgSql & "num_edadini <= " & vlEdadBen & " AND "
            vgSql = vgSql & "num_edadfin >= " & vlEdadBen & " AND "
            vgSql = vgSql & "fec_terpenmin >= '" & vlFechaCalculo & "' AND "
            vgSql = vgSql & "mto_penmin > " & Str(vlMtoPensionPesos) & " "
            vgSql = vgSql & "ORDER BY mto_penminfin "
            Set vgRegistro = vgConexionBD.Execute(vgSql)
            If Not vgRegistro.EOF Then
                vlMtoPensionMin = (vgRegistro!mto_penminfin)
            End If
            
            'Se envian mensajes generados por concepto de bono de invierno
            'Solo a el pensionado.
            If vlMtoPensionMin > 0 Then
                If Trim(vgRs!Cod_Sexo) = clCodSexoM Then
                    If vlEdadBen >= clEdadTope65 Then
                        If Trim(vlNumPolizaAnt) <> (vgRs!num_poliza) Then
                           Call flInsertarMensajes((vgRs!num_poliza), (vgRs!num_endoso), _
                                                   (vgRs!Num_Orden), vlCodMEnBonInv)
                        End If
                    End If
                Else
                    If Trim(vgRs!Cod_Sexo) = clCodSexoF Then
                        If vlEdadBen >= clEdadTope60 Then
                            If Trim(vlNumPolizaAnt) <> (vgRs!num_poliza) Then
                               Call flInsertarMensajes((vgRs!num_poliza), (vgRs!num_endoso), _
                                                       (vgRs!Num_Orden), vlCodMEnBonInv)
                            End If
                        End If
                    End If
                End If
            End If
            
            vlNumPolizaAnt = (vgRs!num_poliza)
            vgRs.MoveNext
       Wend
       MsgBox "Los Datos para Mensajes por Bono de Invierno, han sido Ingresados Satisfactoriamente.", vbInformation, "Información"
    End If


'    MsgBox "El Proceso de Asignación de Mensajes por Bonos de Invierno No se encuentra Disponible.", vbInformation, "Información"

End Function

Function flAsignaMensajesCerEst()

On Error GoTo Err_flAsignaMensajesCerEst

    vlSwCerEst = "N"

    vlFechaCalculo = Format(CDate(Trim(Txt_CerEst.Text)), "yyyymmdd")
        
    vlAnno = Mid(Trim(vlFechaCalculo), 1, 4)
    vlMes = Mid(Trim(vlFechaCalculo), 5, 2)
    vlDia = Mid(Trim(vlFechaCalculo), 7, 2)
    
    'Ej.: 2004-24 = 1980
    vlFecIniTramo = DateSerial(Trim(vlAnno) - 24, Trim(vlMes), Trim(vlDia))
    'Ej.: 2004-18 = 1986
    vlFecTerTramo = DateSerial(Trim(vlAnno) - 18, Trim(vlMes), Trim(vlDia))
    
    vlAnnoIniTramo = Format(CDate(Trim(vlFecIniTramo)), "yyyymmdd")
    vlAnnoIniTramo = Mid(Trim(vlAnnoIniTramo), 1, 4)
    vlAnnoTerTramo = Format(CDate(Trim(vlFecTerTramo)), "yyyymmdd")
    vlAnnoTerTramo = Mid(Trim(vlAnnoTerTramo), 1, 4)
    
    vlTiposPensionSob = flAsignaCodPensionSob
    
'Seleccionar Certificados de Estudios de Pensiones de Sobrevivencia

    vgSql = ""
    vgSql = "SELECT c.num_poliza,c.num_endoso,c.num_orden, "
    vgSql = vgSql & "b.fec_nacben,b.cod_sitinv, "
    vgSql = vgSql & "p.Cod_TipPension, "
    vgSql = vgSql & "a.cod_estvigencia "
    vgSql = vgSql & "FROM PP_TMAE_CERESTUDIO c, PP_TMAE_BEN b, "
    vgSql = vgSql & "PP_TMAE_POLIZA p, PP_TMAE_ASIGFAM a "
    vgSql = vgSql & "WHERE c.fec_tercerest = "
    vgSql = vgSql & "(SELECT MAX(fec_tercerest) from pp_tmae_cerestudio "
    vgSql = vgSql & "WHERE num_poliza = c.num_poliza and num_endoso = c.num_endoso and "
    vgSql = vgSql & " num_orden = c.num_orden) AND "
    vgSql = vgSql & "c.num_poliza = b.num_poliza AND "
    vgSql = vgSql & "c.num_endoso = b.num_endoso AND "
    vgSql = vgSql & "c.num_orden = b.num_orden AND "
    vgSql = vgSql & "c.num_poliza = p.num_poliza and "
    vgSql = vgSql & "c.num_endoso = p.num_endoso AND "
    vgSql = vgSql & "c.num_poliza = a.num_poliza AND "
    vgSql = vgSql & "c.num_orden = a.num_orden AND "
    vgSql = vgSql & "c.fec_tercerest >= '" & vlFechaCalculo & "' AND "
    
    If vgTipoBase = "ORACLE" Then
       vgSql = vgSql & "SUBSTR(b.fec_nacben,1,4) >= '" & vlAnnoIniTramo & "' AND "
       vgSql = vgSql & "SUBSTR(b.fec_nacben,1,4) <= '" & vlAnnoTerTramo & "' AND "
    Else
       vgSql = vgSql & "SUBSTRING(b.fec_nacben,1,4) >= '" & vlAnnoIniTramo & "' AND "
       vgSql = vgSql & "SUBSTRING(b.fec_nacben,1,4) <= '" & vlAnnoTerTramo & "') AND "
    End If
    
'    vgSql = vgSql & "(substr(b.fec_nacben,1,4) >= '" & vlAnnoIniTramo & "' AND "
'    vgSql = vgSql & "substr(b.fec_nacben,1,4) <= '" & vlAnnoTerTramo & "') AND "
    vgSql = vgSql & "b.cod_sitinv = '" & clCodNoInv & "' AND "
    vgSql = vgSql & "p.cod_tippension IN " & vlTiposPensionSob & " and "
    vgSql = vgSql & "a.cod_estvigencia = '" & clCodActiva & "' "
    vgSql = vgSql & "ORDER BY c.num_poliza,c.num_orden "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
    
       vlSwCerEst = "S"
    
       vlNumPerPago = Format(Txt_Anno.Text, "0000") & Format(Txt_Mes.Text, "00")
       vlCodMenCerEst = 0
       vgSql = ""
       vgSql = "SELECT cod_mensajecerest "
       vgSql = vgSql & "FROM MA_TCOD_GENERAL "
       Set vgRs2 = vgConexionBD.Execute(vgSql)
       If Not vgRs2.EOF Then
          vlCodMenCerEst = (vgRs2!cod_mensajecerest)
       End If
       
       vlNumPolizaAnt = ""
       
'Se envian mensajes, para solicitar Certificados de Estudios
'a cada uno de los beneficiarios (con Certificado) de la poliza.
       
       While Not vgRs.EOF
             Call flInsertarMensajes((vgRs!num_poliza), (vgRs!num_endoso), _
                                     (vgRs!Num_Orden), vlCodMenCerEst)
             vlNumPolizaAnt = (vgRs!num_poliza)
             vgRs.MoveNext
       Wend
'       MsgBox "Los Datos para Mensajes por Certificados de Estudio, han sido Ingresados Satisfactoriamente.", vbInformation, "Información"
    End If
    
'Seleccionar Certificados de Estudios de Pensiones de Vejez, Invalidez y Vejez Anticipada

    vgSql = ""
    vgSql = "SELECT c.num_poliza,c.num_endoso,c.num_orden, "
    vgSql = vgSql & "b.fec_nacben,b.cod_sitinv, "
    vgSql = vgSql & "p.Cod_TipPension, "
    vgSql = vgSql & "a.cod_estvigencia "
    vgSql = vgSql & "FROM PP_TMAE_CERESTUDIO c, PP_TMAE_BEN b, "
    vgSql = vgSql & "PP_TMAE_POLIZA p, PP_TMAE_ASIGFAM a "
    vgSql = vgSql & "WHERE c.fec_tercerest = "
    vgSql = vgSql & "(select max(fec_tercerest) from pp_tmae_cerestudio "
    vgSql = vgSql & "where num_poliza = c.num_poliza and num_endoso = c.num_endoso AND "
    vgSql = vgSql & " num_orden = c.num_orden) AND "
    vgSql = vgSql & "c.num_poliza = b.num_poliza AND "
    vgSql = vgSql & "c.num_endoso = b.num_endoso AND "
    vgSql = vgSql & "c.num_orden = b.num_orden AND "
    vgSql = vgSql & "c.num_poliza = p.num_poliza and "
    vgSql = vgSql & "c.num_endoso = p.num_endoso AND "
    vgSql = vgSql & "c.num_poliza = a.num_poliza AND "
    vgSql = vgSql & "c.num_orden = a.num_orden AND "
    vgSql = vgSql & "c.fec_tercerest >= '" & vlFechaCalculo & "' AND "
    vgSql = vgSql & "(substr(b.fec_nacben,1,4) >= '" & vlAnnoIniTramo & "' AND "
    vgSql = vgSql & "substr(b.fec_nacben,1,4) <= '" & vlAnnoTerTramo & "') AND "
    vgSql = vgSql & "b.cod_sitinv = '" & clCodNoInv & "' AND "
    vgSql = vgSql & "p.cod_tippension NOT IN " & vlTiposPensionSob & " and "
    vgSql = vgSql & "a.cod_estvigencia = '" & clCodActiva & "' "
    vgSql = vgSql & "ORDER BY c.num_poliza,c.num_orden "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
    
       vlSwCerEst = "S"
    
       vlNumPerPago = Format(Txt_Anno.Text, "0000") & Format(Txt_Mes.Text, "00")
       vlCodMenCerEst = 0
       vgSql = ""
       vgSql = "SELECT cod_mensajecerest "
       vgSql = vgSql & "FROM MA_TCOD_GENERAL "
       Set vgRs2 = vgConexionBD.Execute(vgSql)
       If Not vgRs2.EOF Then
          vlCodMenCerEst = (vgRs2!cod_mensajecerest)
       End If
       
       vlNumPolizaAnt = ""
       
'Se envian mensajes para solicitar Certificados de Estudios de
'el/los beneficiarios (con Certificado) Solo a el pensionado.
       
       While Not vgRs.EOF
             If Trim(vlNumPolizaAnt) <> (vgRs!num_poliza) Then
                Call flInsertarMensajes((vgRs!num_poliza), (vgRs!num_endoso), _
                                        (vgRs!Num_Orden), vlCodMenCerEst)
             End If
             vlNumPolizaAnt = (vgRs!num_poliza)
             vgRs.MoveNext
       Wend
'       MsgBox "Los Datos para Mensajes por Certificados de Estudio, han sido Ingresados Satisfactoriamente.", vbInformation, "Información"
    End If
    
    If vlSwCerEst = "N" Then
       MsgBox "No Existen Certificados de Estudio para enviar Mensajes.", vbInformation, "Información"
       Exit Function
    Else
        MsgBox "Los Datos para Mensajes por Certificados de Estudio, han sido Ingresados Satisfactoriamente.", vbInformation, "Información"
    End If
    
Exit Function
Err_flAsignaMensajesCerEst:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
    
End Function
''
''Function flAsignaMensajesAsiFam()
''
''On Error GoTo Err_flAsignaMensajesAsiFam
''
''    vlSwAsiFam = "N"
''
''    vlFechaCalculo = Format(CDate(Trim(Txt_AsiFam.Text)), "yyyymmdd")
''    vlTiposPensionSob = flAsignaCodPensionSob
''
'''Seleccionar a todas las cargas familiares de polizas con tipo de
'''pension de vejez, invalidez y vejez anticipada.
''
''
'''''    vgSql = "SELECT a.num_poliza,a.num_endoso,a.num_orden, "
'''''    vgSql = vgSql & " p.cod_tippension "
'''''    vgSql = vgSql & "FROM PP_TMAE_ASIGFAM a, PP_TMAE_POLIZA P "
'''''    vgSql = vgSql & "WHERE cod_estvigencia = '" & clCodActiva & "' AND "
'''''    vgSql = vgSql & "fec_iniactiva <= '" & vlFechaCalculo & "' AND "
'''''    vgSql = vgSql & "fec_teractiva >= '" & vlFechaCalculo & "' AND "
'''''    vgSql = vgSql & "p.num_poliza = a.num_poliza AND "
'''''    vgSql = vgSql & "p.num_endoso = a.num_endoso AND "
'''''    vgSql = vgSql & "p.cod_tippension NOT IN '" & vlTiposPensionSob & "' "
''
''    vgSql = ""
''    vgSql = "SELECT p.num_poliza,p.num_endoso,p.num_orden, "
''    vgSql = vgSql & "l.Fec_Pago "
''    vgSql = vgSql & "FROM PP_TMAE_PAGOPENDEF p, PP_TMAE_LIQPAGOPENDEF l "
''    vgSql = vgSql & "WHERE substr(l.fec_pago,1,6) = '" & Trim(Mid(vlFechaCalculo, 1, 6)) & "' AND "
''    vgSql = vgSql & "p.num_perpago = l.num_perpago AND "
''    vgSql = vgSql & "p.num_poliza = l.num_poliza AND "
''    vgSql = vgSql & "p.num_orden = l.num_orden AND "
''    vgSql = vgSql & "p.cod_conhabdes = '08' "
''    vgSql = vgSql & "ORDER BY p.num_poliza,p.num_orden ASC "
''
''    Set vgRs = vgConexionBD.Execute(vgSql)
''    If Not vgRs.EOF Then
''
''       vlSwAsiFam = "S"
''
''       vlNumPerPago = Format(Txt_Anno.Text, "0000") & Format(Txt_Mes.Text, "00")
''       vlCodMenAsiFam = 0
''       vgSql = ""
''       vgSql = "SELECT cod_mensajeasifam "
''       vgSql = vgSql & "FROM MA_TCOD_GENERAL "
''       Set vgRs2 = vgConexionBD.Execute(vgSql)
''       If Not vgRs2.EOF Then
''          vlCodMenAsiFam = (vgRs2!cod_mensajeasifam)
''       End If
''       vlNumPolizaAnt = ""
''       While Not vgRs.EOF
''             If Trim(vlNumPolizaAnt) <> (vgRs!Num_Poliza) Then
''                Call flInsertarMensajes((vgRs!Num_Poliza), (vgRs!num_endoso), _
''                                        (vgRs!Num_Orden), vlCodMenAsiFam)
''             End If
''             vlNumPolizaAnt = (vgRs!Num_Poliza)
''             vgRs.MoveNext
''       Wend
''          MsgBox "Los Datos para Mensajes por Asignación Familiar, han sido Ingresados Satisfactoriamente.", vbInformation, "Información"
''    End If
''
''    If vlSwAsiFam = "N" Then
''       MsgBox "No Existen Cargas Familiares Activas para enviar Mensajes.", vbInformation, "Información"
''    End If
''
''Exit Function
''Err_flAsignaMensajesAsiFam:
''    Screen.MousePointer = 0
''    Select Case Err
''        Case Else
''        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
''    End Select
''
''End Function

Function flAsignaCodPensionSob() As String

On Error GoTo Err_flAsignaCodPensionSob

    flAsignaCodPensionSob = "('03','08','09','10','11','12','13','15')"
    
Exit Function
Err_flAsignaCodPensionSob:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
        
End Function

Function flInsertarMensajes(iNumPoliza As String, inumendoso As Integer, _
                            iNumOrden As Integer, iCodMensaje As Integer)

On Error GoTo Err_flInsertarMensajes

    vgSql = ""
    vgSql = "SELECT num_poliza "
    vgSql = vgSql & "FROM pp_tmae_menpoliza m "
    vgSql = vgSql & "WHERE m.num_poliza = '" & Trim(iNumPoliza) & "' AND "
    vgSql = vgSql & "m.num_orden = " & Trim(Str(iNumOrden)) & " AND "
    vgSql = vgSql & "m.cod_mensaje = " & Trim(Str(iCodMensaje)) & " AND "
    vgSql = vgSql & "m.num_perpago = '" & Trim(vlNumPerPago) & "' "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If vgRegistro.EOF Then
    
        vlGlsUsuarioCrea = vgUsuario
        vlFecCrea = Format(Date, "yyyymmdd")
        vlHorCrea = Format(Time, "hhmmss")
    
        vgSql = ""
        vgSql = "INSERT INTO PP_TMAE_MENPOLIZA "
        vgSql = vgSql & "(num_poliza,num_endoso,num_orden, "
        vgSql = vgSql & " cod_mensaje,num_perpago,cod_tipoing, "
        vgSql = vgSql & " cod_usuariocrea,fec_crea,hor_crea "
        vgSql = vgSql & " ) VALUES ( "
        vgSql = vgSql & " '" & Trim(iNumPoliza) & "' , "
        vgSql = vgSql & " " & Trim(Str(inumendoso)) & ", "
        vgSql = vgSql & " " & Trim(Str(iNumOrden)) & ", "
        vgSql = vgSql & " " & Trim(Str(iCodMensaje)) & ", "
        vgSql = vgSql & " '" & Trim(vlNumPerPago) & "', "
        vgSql = vgSql & " '" & Trim(clCodTipoIng) & "', "
        vgSql = vgSql & "'" & Trim(vlGlsUsuarioCrea) & "', "
        vgSql = vgSql & "'" & Trim(vlFecCrea) & "', "
        vgSql = vgSql & "'" & Trim(vlHorCrea) & "') "
        vgConexionBD.Execute vgSql
    End If
    
Exit Function
Err_flInsertarMensajes:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

''Private Sub Chk_AsiFam_KeyPress(KeyAscii As Integer)
''    If KeyAscii = 13 Then
''       Txt_AsiFam.SetFocus
''    End If
''End Sub

Private Sub Chk_BonInv_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Txt_BonInv.SetFocus
    End If
End Sub

Private Sub Chk_CerEst_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Txt_CerEst.SetFocus
    End If
End Sub

Private Sub Chk_GarEst_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Txt_GarEst.SetFocus
    End If
End Sub

Private Sub Cmd_Eliminar_Click()

    MsgBox "El Proceso de Eliminación No se encuentra Disponible.", vbInformation, "Información"

End Sub

'*****************************************************************

Private Sub cmd_grabar_Click()

On Error GoTo Err_CmdGrabarClick

'Valida Mes Ingresado en Periodo de Pago

    If Txt_Mes.Text = "" Then
       MsgBox "Debe ingresar Mes del Periodo de Pago.", vbCritical, "Error de Datos"
       Txt_Mes.SetFocus
       Exit Sub
    End If
      
    If CDbl(Txt_Mes.Text) <= 0 Or CDbl(Txt_Mes.Text) > 12 Then
       MsgBox "El Mes ingresado No es un Valor Válido.", vbCritical, "Error de Datos"
       Txt_Mes.SetFocus
       Exit Sub
    End If
    
    Txt_Mes.Text = Format(Txt_Mes.Text, "00")
    
'Valida Año Ingresado en Periodo de Pago

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
    
    vlPeriodo = Format(Txt_Anno, "0000") & Format(Txt_Mes, "00")
    
    'Validar si Periodo se encuentra abierto
    'vlSwEstPeriodo = C    Periodo Cerrado
    'vlSwEstPeriodo = A    Periodo Abierto
    'vlSwEstPeriodo = "C"
    
    vgSql = ""
    vgSql = "SELECT num_perpago "
    vgSql = vgSql & "FROM PP_TMAE_PROPAGOPEN "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "num_perpago = '" & Trim(vlPeriodo) & "' AND "
    'vgSql = vgSql & "(cod_estadoreg = 'A' OR cod_estadopri = 'A') "
    vgSql = vgSql & "(cod_estadoreg <> 'C' OR cod_estadopri <> 'C') "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If vgRs.EOF Then
       'vlSwEstPeriodo = "C"
       'El Periodo se encuentra CERRADO
       MsgBox "El Período Ingresado se Encuentra Cerrado, Debe Ingresar un Nuevo Periodo.", vbCritical, "Error de Datos"
       Exit Sub
    'Else
        'vlSwEstPeriodo = "A"
    End If

'Valida Fecha Ingresada en Garantía Estatal

    If (Trim(Txt_GarEst.Text) = "") Then
       Chk_GarEst.Value = 0
    Else
        If Not IsDate(Txt_GarEst.Text) Then
           MsgBox "La Fecha Ingresada para Garantía Estatal No es una Fecha Válida.", vbCritical, "Error de Datos"
           Txt_GarEst.SetFocus
           Exit Sub
        End If
        If (Year(Txt_GarEst.Text) < 1900) Then
           MsgBox "La Fecha Ingresada para Garantía Estatal es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
           Txt_GarEst.SetFocus
           Exit Sub
        End If
        
        Txt_GarEst.Text = Format(CDate(Trim(Txt_GarEst.Text)), "yyyymmdd")
        Txt_GarEst.Text = DateSerial(Mid((Txt_GarEst.Text), 1, 4), Mid((Txt_GarEst.Text), 5, 2), Mid((Txt_GarEst.Text), 7, 2))
        
        
    End If
    
'Valida Fecha Ingresada en Bono de Invierno

    If (Trim(Txt_BonInv.Text) = "") Then
       Chk_BonInv.Value = 0
    Else
        If Not IsDate(Txt_BonInv.Text) Then
           MsgBox "La Fecha Ingresada para Bono de Invierno No es una Fecha Válida.", vbCritical, "Error de Datos"
           Txt_BonInv.SetFocus
           Exit Sub
        End If
        If (Year(Txt_BonInv.Text) < 1900) Then
           MsgBox "La Fecha Ingresada para Bono de Invierno es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
           Txt_BonInv.SetFocus
           Exit Sub
        End If
        
        Txt_BonInv.Text = Format(CDate(Trim(Txt_BonInv.Text)), "yyyymmdd")
        Txt_BonInv.Text = DateSerial(Mid((Txt_BonInv.Text), 1, 4), Mid((Txt_BonInv.Text), 5, 2), Mid((Txt_BonInv.Text), 7, 2))
    End If
       
'Valida Fecha Ingresada en Certificado de Estudios
       
    If (Trim(Txt_CerEst.Text) = "") Then
       Chk_CerEst.Value = 0
    Else
        If Not IsDate(Txt_CerEst.Text) Then
           MsgBox "La Fecha Ingresada para Certificado de Estudio No es una Fecha Válida.", vbCritical, "Error de Datos"
           Txt_CerEst.SetFocus
           Exit Sub
        End If
        If (Year(Txt_CerEst.Text) < 1900) Then
           MsgBox "La Fecha Ingresada para Certificado de Estudio es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
           Txt_CerEst.SetFocus
           Exit Sub
        End If
        
        Txt_CerEst.Text = Format(CDate(Trim(Txt_CerEst.Text)), "yyyymmdd")
        Txt_CerEst.Text = DateSerial(Mid((Txt_CerEst.Text), 1, 4), Mid((Txt_CerEst.Text), 5, 2), Mid((Txt_CerEst.Text), 7, 2))
    End If
                     
'Valida la Fecha Ingresada en Asignación Familiar

'''    If (Trim(Txt_AsiFam.Text) = "") Then
'''       Chk_AsiFam.Value = 0
'''    Else
'''        If Not IsDate(Txt_AsiFam.Text) Then
'''           MsgBox "La Fecha Ingresada para Asignación Familiar No es una Fecha Válida.", vbCritical, "Error de Datos"
'''           Txt_AsiFam.SetFocus
'''           Exit Sub
'''        End If
'''        If (Year(Txt_AsiFam.Text) < 1900) Then
'''           MsgBox "La Fecha Ingresada para Asignación Familiar es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
'''           Txt_AsiFam.SetFocus
'''           Exit Sub
'''        End If
'''
'''        Txt_AsiFam.Text = Format(CDate(Trim(Txt_AsiFam.Text)), "yyyymmdd")
'''        Txt_AsiFam.Text = DateSerial(Mid((Txt_AsiFam.Text), 1, 4), Mid((Txt_AsiFam.Text), 5, 2), Mid((Txt_AsiFam.Text), 7, 2))
'''    End If
               
    Screen.MousePointer = 11
           
    If Chk_GarEst.Value = 1 Then
       Call flAsignaMensajesGarEst
    End If
     
    If Chk_BonInv.Value = 1 Then
       Call flAsignaMensajesBonInv
    End If
     
    If Chk_CerEst.Value = 1 Then
       Call flAsignaMensajesCerEst
    End If
       
'''    If Chk_AsiFam.Value = 1 Then
'''       Call flAsignaMensajesAsiFam
'''    End If
   
    Screen.MousePointer = 0
   
Exit Sub
Err_CmdGrabarClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Imprimir_Click()
Dim vlArchivo As String '20-10-2004
On Error GoTo Err_CmdImprimir
  
   Screen.MousePointer = 11
   
   vlArchivo = strRpt & "PP_Rpt_MenMasivo.rpt"   '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
      MsgBox "Archivo de Reporte de Mensajes Masivos no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
      Screen.MousePointer = 0
      Exit Sub
   End If
   
'   vgQuery = "{PP_TMAE_RETJUDICIAL.NUM_POLIZA} = '" & Trim(Txt_PenPoliza.Text) & "' "
      
   Rpt_Calculo.Reset
   Rpt_Calculo.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
   Rpt_Calculo.Connect = vgRutaDataBase       ' o App.Path & "\Nestle.mdb"
'   Rpt_Calculo.SelectionFormula = vgQuery
   'Rpt_Calculo.Formulas(0) = "RutPensionado = '" & (Trim(Txt_PenRut.Text)) & " - " & (Trim(Txt_PenDigito.Text)) & "' "
   'Rpt_Calculo.Formulas(1) = "NombrePensionado = '" & Trim(Lbl_PenNombre.Caption) & "' "
   Rpt_Calculo.Formulas(2) = ""
   
   Rpt_Calculo.Formulas(3) = "NombreCompania = '" & vgNombreCompania & "'"
   Rpt_Calculo.Formulas(4) = "NombreSistema= '" & vgNombreSistema & "'"
   Rpt_Calculo.Formulas(5) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
      
   Rpt_Calculo.WindowState = crptMaximized
   Rpt_Calculo.Destination = crptToWindow
   Rpt_Calculo.WindowTitle = "Informe de Retención Judicial por Pensionado"
   Rpt_Calculo.Action = 1
   Screen.MousePointer = 0
   
Exit Sub
Err_CmdImprimir:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub

Private Sub Cmd_Limpiar_Click()

    Txt_Mes.Text = ""
    Txt_Anno.Text = ""

    Chk_GarEst.Value = 0
    Chk_BonInv.Value = 0
    Chk_CerEst.Value = 0
'''    Chk_AsiFam.Value = 0
    
    Txt_GarEst.Text = ""
    Txt_BonInv.Text = ""
    Txt_CerEst.Text = ""
'''    Txt_AsiFam.Text = ""
    
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

    Frm_MenMasivo.Top = 0
    Frm_MenMasivo.Left = 0
            
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
    
End Sub

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
        Chk_GarEst.SetFocus
         
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
'''
'''Private Sub Txt_AsiFam_KeyPress(KeyAscii As Integer)
'''
'''    If KeyAscii = 13 Then
'''       If (Trim(Txt_AsiFam.Text) = "") Then
'''          Exit Sub
'''       End If
'''       If Not IsDate(Txt_AsiFam.Text) Then
'''          MsgBox "La Fecha Ingresada para Asignación Familiar No es una Fecha Válida.", vbCritical, "Error de Datos"
'''          Txt_AsiFam.SetFocus
'''          Exit Sub
'''       End If
'''       If (Year(Txt_AsiFam.Text) < 1900) Then
'''          MsgBox "La Fecha Ingresada para Asignación Familiar es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
'''          Txt_AsiFam.SetFocus
'''          Exit Sub
'''       End If
'''
'''       Txt_AsiFam.Text = Format(CDate(Trim(Txt_AsiFam.Text)), "yyyymmdd")
'''       Txt_AsiFam.Text = DateSerial(Mid((Txt_AsiFam.Text), 1, 4), Mid((Txt_AsiFam.Text), 5, 2), Mid((Txt_AsiFam.Text), 7, 2))
'''
'''       Cmd_Grabar.SetFocus
'''
'''    End If
'''
'''End Sub

Private Sub Txt_AsiFam_LostFocus()
    Cmd_Grabar.SetFocus
End Sub

Private Sub Txt_BonInv_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       If (Trim(Txt_BonInv.Text) = "") Then
          Exit Sub
       End If
       If Not IsDate(Txt_BonInv.Text) Then
          MsgBox "La Fecha Ingresada para Bono de Invierno No es una Fecha Válida.", vbCritical, "Error de Datos"
          Txt_BonInv.SetFocus
          Exit Sub
       End If
       If (Year(Txt_BonInv.Text) < 1900) Then
          MsgBox "La Fecha Ingresada para Bono de Invierno es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
          Txt_BonInv.SetFocus
          Exit Sub
       End If
       
       Txt_BonInv.Text = Format(CDate(Trim(Txt_BonInv.Text)), "yyyymmdd")
       Txt_BonInv.Text = DateSerial(Mid((Txt_BonInv.Text), 1, 4), Mid((Txt_BonInv.Text), 5, 2), Mid((Txt_BonInv.Text), 7, 2))
       
      
      
    End If

End Sub

Private Sub Txt_BonInv_LostFocus()

    If (Trim(Txt_BonInv.Text) = "") Then
       Exit Sub
    End If
    If Not IsDate(Txt_BonInv.Text) Then
       Exit Sub
    End If
    If (Year(Txt_BonInv.Text) < 1900) Then
       Exit Sub
    End If
    
    Txt_BonInv.Text = Format(CDate(Trim(Txt_BonInv.Text)), "yyyymmdd")
    Txt_BonInv.Text = DateSerial(Mid((Txt_BonInv.Text), 1, 4), Mid((Txt_BonInv.Text), 5, 2), Mid((Txt_BonInv.Text), 7, 2))
    
    

End Sub

Private Sub Txt_CerEst_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       If (Trim(Txt_CerEst.Text) = "") Then
          Exit Sub
       End If
       If Not IsDate(Txt_CerEst.Text) Then
          MsgBox "La Fecha Ingresada para Certificado de Estudio No es una Fecha Válida.", vbCritical, "Error de Datos"
          Txt_CerEst.SetFocus
          Exit Sub
       End If
       If (Year(Txt_CerEst.Text) < 1900) Then
          MsgBox "La Fecha Ingresada para Certificado de Estudio es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
          Txt_CerEst.SetFocus
          Exit Sub
       End If
       
       Txt_CerEst.Text = Format(CDate(Trim(Txt_CerEst.Text)), "yyyymmdd")
       Txt_CerEst.Text = DateSerial(Mid((Txt_CerEst.Text), 1, 4), Mid((Txt_CerEst.Text), 5, 2), Mid((Txt_CerEst.Text), 7, 2))
       
'''       Chk_AsiFam.SetFocus
      
    End If

End Sub

Private Sub Txt_CerEst_LostFocus()

    If (Trim(Txt_CerEst.Text) = "") Then
       Exit Sub
    End If
    If Not IsDate(Txt_CerEst.Text) Then
       Exit Sub
    End If
    If (Year(Txt_CerEst.Text) < 1900) Then
       Exit Sub
    End If
    
    Txt_CerEst.Text = Format(CDate(Trim(Txt_CerEst.Text)), "yyyymmdd")
    Txt_CerEst.Text = DateSerial(Mid((Txt_CerEst.Text), 1, 4), Mid((Txt_CerEst.Text), 5, 2), Mid((Txt_CerEst.Text), 7, 2))
    
   
  
End Sub

Private Sub Txt_GarEst_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       If (Trim(Txt_GarEst.Text) = "") Then
          Exit Sub
       End If
       If Not IsDate(Txt_GarEst.Text) Then
          MsgBox "La Fecha Ingresada para Garantía Estatal No es una Fecha Válida.", vbCritical, "Error de Datos"
          Txt_GarEst.SetFocus
          Exit Sub
       End If
       If (Year(Txt_GarEst.Text) < 1900) Then
          MsgBox "La Fecha Ingresada para Garantía Estatal es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
          Txt_GarEst.SetFocus
          Exit Sub
       End If
       
       Txt_GarEst.Text = Format(CDate(Trim(Txt_GarEst.Text)), "yyyymmdd")
       Txt_GarEst.Text = DateSerial(Mid((Txt_GarEst.Text), 1, 4), Mid((Txt_GarEst.Text), 5, 2), Mid((Txt_GarEst.Text), 7, 2))
       
       Chk_BonInv.SetFocus
      
    End If
    
End Sub

Private Sub Txt_GarEst_LostFocus()

    If (Trim(Txt_GarEst.Text) = "") Then
       Exit Sub
    End If
    If Not IsDate(Txt_GarEst.Text) Then
       Exit Sub
    End If
    If (Year(Txt_GarEst.Text) < 1900) Then
       Exit Sub
    End If
    
    Txt_GarEst.Text = Format(CDate(Trim(Txt_GarEst.Text)), "yyyymmdd")
    Txt_GarEst.Text = DateSerial(Mid((Txt_GarEst.Text), 1, 4), Mid((Txt_GarEst.Text), 5, 2), Mid((Txt_GarEst.Text), 7, 2))
    
   

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
          MsgBox "Debe ingresar Mes del Periodo de Pago.", vbCritical, "Error de Datos"
          Txt_Mes.SetFocus
          Exit Sub
       End If
         
       If CDbl(Txt_Mes.Text) <= 0 Or CDbl(Txt_Mes.Text) > 12 Then
          MsgBox "El Mes ingresado No es un Valor Válido.", vbCritical, "Error de Datos"
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



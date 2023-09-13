VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Frm_CargaArchBen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generación Archivo Datos "
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   3705
   Begin VB.Frame Fra_Operaciones 
      Height          =   1095
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   3495
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   1920
         Picture         =   "Frm_CargaArchBen.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Cargar 
         Caption         =   "&Archivo"
         Height          =   675
         Left            =   720
         Picture         =   "Frm_CargaArchBen.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Exportar Datos a Archivo"
         Top             =   240
         Width           =   720
      End
      Begin MSComDlg.CommonDialog ComDialogo 
         Left            =   120
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.Label Lbl_FechaActual 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1400
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Actual"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Frm_CargaArchBen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***** Variables Generación de Archivo de Datos de Beneficiarios ***
'Datos de Poliza
Dim vlNumPoliza As String
Dim vlNumEndoso As String
Dim vlCodAFP As String
Dim vlCodTipPension As String
Dim vlCodEstado As String
Dim vlCodTipRen As String
Dim vlCodModalidad As String
Dim vlNumCargas As String
Dim vlFecVigencia As String
Dim vlFecTerVigencia As String
Dim vlMtoPrima As String
Dim vlMtoPension As String
Dim vlNumMesDif As String
Dim vlNumMesGar As String
Dim vlPrcTasaCe As String
Dim vlPrcTasaVta As String
Dim vlPrcTasaCtoRea As String
Dim vlPrcTasaIntPerGar As String
Dim vlFecIniPagoPen As String
Dim vlCodTipOrigen As String
Dim vlNumIndQuiebra As String
'Datos de Beneficiario
Dim vlNumOrden As String
Dim vlFecIngreso As String
Dim vlCodTipoIdenBen As String
Dim vlNumIden As String
Dim vlGlsNomBen As String
Dim vlGlsSegNomBen As String
Dim vlGlsPatBen As String
Dim vlGlsMatBen As String
Dim vlGlsDirBen As String
Dim vlCodDireccion As String

Dim vlGlsComuna As String
Dim vlGlsProvincia As String
Dim vlGlsRegion As String

Dim vlGlsFonoBen As String
Dim vlGlsCorreoBen As String
Dim vlCodGruFam As String
Dim vlCodPar As String
Dim vlCodSexo As String
Dim vlCodSitInv As String
Dim vlCodDerCre As String
Dim vlCodDerpen As String
Dim vlCodCauInv As String
Dim vlFecNacBen As String
Dim vlFecNacHM As String
Dim vlFecInvBen As String
Dim vlCodMotReqPen As String
Dim vlMtoPensionB As String
Dim vlMtoPensionGar As String
Dim vlPrcPension As String
Dim vlCodInsSalud As String
Dim vlCodModSalud As String
Dim vlMtoPlanSalud As String
Dim vlCodEstPension As String
Dim vlCodCajaCompen As String
Dim vlCodViaPago As String
Dim vlCodBanco As String
Dim vlCodTipCuenta As String
Dim vlNumCuenta As String
Dim vlCodSucursal As String
Dim vlFecFallBen As String
Dim vlFecMatrimonio As String
Dim vlCodCauSusBen As String
Dim vlFecSusBen As String
Dim vlFecIniPagoPenB As String
Dim vlFecTerPagoPenGarB As String
Dim vlMtoPlanSalud2 As String
Dim vlCodModSalud2 As String
Dim vlNumFun As String

'Carga de Códigos
Dim stAfp()     As String
Dim stTipPen()  As String
Dim stEstPol()  As String
Dim stTipRen()  As String
Dim stMod()     As String
Dim stCodPar()  As String
Dim stDerpen()  As String
Dim stInsSal()  As String
Dim stEstPen()  As String
Dim stViaPago() As String
Dim stBanco()   As String
Dim stTipCta()  As String
Dim stTipIden() As String
Dim stSuc()     As String

Private Sub Cmd_Cargar_Click()
On Error GoTo Err_ExportarDatos
    
    'Permite imprimir la Opción Indicada a través del Menú
    Select Case vgNomInfSeleccionado
        Case "InfGeneraArchDatosBen"    'Genera Archivo de Datos de Beneficiarios
            Call flExportarDatosBen
            
    End Select
    
Exit Sub
Err_ExportarDatos:
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

'************** Generación de Archivo de Datos de Beneficiarios *********
Function flExportarDatosBen()
On Error GoTo Err_flExportarDatosBen
Dim vlLinea As String
Dim vlArchivo As String, vlOpen As Boolean
Dim vlContador As Long
Dim vlAumento As Integer
    
    vgSql = ""
    vgSql = "SELECT  p.num_poliza,p.num_endoso,p.cod_afp,p.cod_tippension, "
    vgSql = vgSql & "p.cod_estado,p.cod_tipren,p.cod_modalidad,p.num_cargas,"
    vgSql = vgSql & "p.fec_vigencia,p.fec_tervigencia,p.mto_prima,p.mto_pension,"
    vgSql = vgSql & "p.num_mesdif,p.num_mesgar,p.prc_tasace,p.prc_tasavta,"
    vgSql = vgSql & "p.prc_tasactorea,p.prc_tasaintpergar,p.fec_inipagopen,"
    vgSql = vgSql & "p.cod_tiporigen,p.num_indquiebra,"
    vgSql = vgSql & "b.num_orden,b.fec_ingreso,b.cod_tipoidenben,b.num_idenben,b.gls_nomben,"
    vgSql = vgSql & "b.gls_nomsegben,b.gls_patben,b.gls_matben,b.gls_dirben,b.cod_direccion,"
    vgSql = vgSql & "b.gls_fonoben,b.gls_correoben,b.cod_grufam,b.cod_par,"
    vgSql = vgSql & "b.cod_sexo,b.cod_sitinv,b.cod_dercre,b.cod_derpen,"
    vgSql = vgSql & "b.cod_cauinv,b.fec_nacben,b.fec_nachm,b.fec_invben,"
    vgSql = vgSql & "b.cod_motreqpen,b.mto_pension,b.mto_pensiongar,"
    vgSql = vgSql & "b.prc_pension,b.cod_inssalud,b.cod_modsalud,"
    vgSql = vgSql & "b.mto_plansalud,b.cod_estpension, "
    vgSql = vgSql & "b.cod_viapago,b.cod_banco,b.cod_tipcuenta,b.num_cuenta,"
    vgSql = vgSql & "b.cod_sucursal,b.fec_fallben, "
    vgSql = vgSql & "b.cod_caususben,b.fec_susben,b.fec_inipagopen,"
    vgSql = vgSql & "b.fec_terpagopengar "
    vgSql = vgSql & "FROM pp_tmae_poliza p, pp_tmae_ben b "
    vgSql = vgSql & "WHERE b.num_poliza = p.num_poliza AND "
    vgSql = vgSql & "b.num_endoso = "
    vgSql = vgSql & "(SELECT MAX(p.num_endoso) FROM pp_tmae_poliza p WHERE "
    vgSql = vgSql & "p.num_poliza = b.num_poliza) "
    vgSql = vgSql & "ORDER BY p.num_poliza, b.num_orden "
    
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
    
        'Selección del Archivo en el que se generarán los pagos
        ComDialogo.CancelError = True
        ComDialogo.FileName = "ArchivoDatosBen"
        ComDialogo.DialogTitle = "Guardar Datos de Beneficiarios como"
        ComDialogo.Filter = "*.txt"
        ComDialogo.FilterIndex = 1
        ComDialogo.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
        ComDialogo.ShowSave
        vlArchivo = ComDialogo.FileName & ".txt"
        If vlArchivo = "" Then
            Exit Function
        End If

        Screen.MousePointer = 11

        Open vlArchivo For Output As #1
        vlOpen = True

        Frm_BarraProg.Show
        Frm_BarraProg.ProgressBar1.Value = 0
        Frm_BarraProg.Refresh
        Frm_BarraProg.Lbl_Texto = "Generando Archivo " & vlArchivo 'vlNombre_Archivo
        Frm_BarraProg.Refresh
        Frm_BarraProg.ProgressBar1.Visible = True
        Frm_BarraProg.Refresh
        vlAumento = 100 / (1 + 2 + 3)
        
        'Carga Matrices de Códigos
        Call flCargaCodigos
        
        'Generar Línea de Registro de Descripciones de Datos de Beneficiarios
        vlLinea = "Número de Póliza" & ";" & "Número de Endoso" & ";" & "Código de AFP" & ";" & _
                  "Tipo de Pensión" & ";" & "Código Estado" & ";" & "Código de Tipo de Renta" & ";" & _
                  "Código de Modalidad" & ";" & "Número de Beneficiarios" & ";" & "Fecha de Vigencia" & ";" & _
                  "Fecha de Termino de Vigencia" & ";" & "Monto Prima" & ";" & "Monto Pensión" & ";" & _
                  "Número de Meses Diferidos" & ";" & "Número de Meses Garantizados" & ";" & "Porcentaje Tasa de Costo Equivalente" & ";" & _
                  "Porcentaje Tasa de Venta" & ";" & "Porcentaje Tasa Costo Reaseguro" & ";" & "Porcentaje Tasa Interés Periodo Garantizado" & ";" & _
                  "Fecha de Inicio de Pago Pensión" & ";" & _
                  "Número de Orden" & ";" & "Fecha de Ingreso" & ";" & "Tipo Documento" & ";" & "Número de Documento" & ";" & _
                  "Nombre Beneficiario" & ";" & "Segundo Nombre Beneficiario" & ";" & "Apellido Paterno Beneficiario" & ";" & "Apellido Materno Beneficiario" & ";" & "Código Dirección" & ";" & _
                  "Descripción Dirección" & ";" & "Descripción Distrito" & ";" & _
                  "Descripción Provincia" & ";" & "Descripción Departamento" & ";" & "Teléfono" & ";" & _
                  "Correo" & ";" & "Código Grupo Familiar" & ";" & "Código Parentesco" & ";" & _
                  "Código Sexo" & ";" & "Código Situación de Invalidez" & ";" & "Código Derecho a Crecer" & ";" & _
                  "Código Derecho a Pensión" & ";" & "Código Causal de Invalidez." & ";" & "Fecha de Nacimiento" & ";" & _
                  "Fecha Nacimiento Hijo Menor" & ";" & "Fecha Invalidez" & ";" & "Código Motivo Requisito Pensión" & ";" & _
                  "Monto Pensión" & ";" & "Monto Pensión Garantizada" & ";" & "Porcentaje Pensión" & ";" & _
                  "Código Institución de Salud" & ";" & "Código Modalidad de Salud" & ";" & "Monto Plan Salud" & ";" & _
                  "Código Estado Pensión" & ";" & "Código Vía de Pago" & ";" & _
                  "Código de Banco" & ";" & "Código Tipo de Cuenta" & ";" & "Número de Cuenta" & ";" & _
                  "Código Sucursal AFP" & ";" & "Fecha de Fallecimiento" & ";" & _
                  "Código de Causa de Suspensión" & ";" & "Fecha de Suspensión" & ";" & "Fecha de Inicio Pago Pensión" & ";" & _
                  "Fecha de Término de Pago Pensión Garantizada"
                  
        Print #1, vlLinea
        
        While Not vgRegistro.EOF

            Call flLimpiarVariables
                
            'Datos de Poliza
            vlNumPoliza = (vgRegistro!Num_Poliza)
            vlNumEndoso = CStr(vgRegistro!num_endoso)
            vlCodAFP = flObtDescripcion(stAfp, vgRegistro!cod_afp) '(vgRegistro!COD_AFP)
            vlCodTipPension = flObtDescripcion(stTipPen, vgRegistro!Cod_TipPension) '(vgRegistro!Cod_TipPension)
            vlCodEstado = flObtDescripcion(stEstPol, vgRegistro!Cod_Estado) '(vgRegistro!Cod_Estado)
            vlCodTipRen = flObtDescripcion(stTipRen, vgRegistro!Cod_TipRen) '(vgRegistro!Cod_TipRen)
            vlCodModalidad = flObtDescripcion(stMod, vgRegistro!Cod_Modalidad) '(vgRegistro!Cod_Modalidad)
            vlNumCargas = CStr(vgRegistro!Num_Cargas)
            vlFecVigencia = (vgRegistro!Fec_Vigencia)
            vlFecTerVigencia = (vgRegistro!Fec_TerVigencia)
            vlMtoPrima = CStr(vgRegistro!Mto_Prima)
            vlMtoPension = CStr(vgRegistro!Mto_Pension)
            vlNumMesDif = CStr(vgRegistro!Num_MesDif)
            vlNumMesGar = CStr(vgRegistro!Num_MesGar)
            vlPrcTasaCe = CStr(vgRegistro!Prc_TasaCe)
            vlPrcTasaVta = CStr(vgRegistro!Prc_TasaVta)
            vlPrcTasaCtoRea = CStr(vgRegistro!prc_tasactorea)
            vlPrcTasaIntPerGar = CStr(vgRegistro!Prc_TasaIntPerGar)
            vlFecIniPagoPen = (vgRegistro!Fec_IniPagoPen)
            'vlCodTipOrigen = (vgRegistro!cod_tiporigen)
            vlNumIndQuiebra = (vgRegistro!num_indquiebra)
            'Datos de Beneficiario
            vlNumOrden = CStr(vgRegistro!Num_Orden)
            vlFecIngreso = (vgRegistro!Fec_Ingreso)
            vlCodTipoIdenBen = flObtDescripcion(stTipIden, vgRegistro!Cod_TipoIdenBen) 'CStr(vgRegistro!Cod_TipoIdenBen)
            vlNumIden = (vgRegistro!Num_IdenBen)
            vlGlsNomBen = (vgRegistro!Gls_NomBen)
            If Not IsNull(vgRegistro!Gls_NomSegBen) Then
                vlGlsSegNomBen = (vgRegistro!Gls_NomSegBen)
            Else
                vlGlsSegNomBen = ""
            End If
            If Not IsNull(vgRegistro!Gls_PatBen) Then
                vlGlsPatBen = (vgRegistro!Gls_PatBen)
            Else
                vlGlsPatBen = ""
            End If
            If Not IsNull(vgRegistro!Gls_MatBen) Then
                vlGlsMatBen = (vgRegistro!Gls_MatBen)
            Else
                vlGlsMatBen = ""
            End If
            vlGlsDirBen = (vgRegistro!Gls_DirBen)
            vlCodDireccion = CStr(vgRegistro!Cod_Direccion)
            'Buscar descripciones de direccion
            vlGlsComuna = flBuscaNombreComuna(vgRegistro!Cod_Direccion)
            Call fgBuscarNombreProvinciaRegion(vgRegistro!Cod_Direccion)
            vlGlsProvincia = vgNombreProvincia
            vlGlsRegion = vgNombreRegion
            
            If Not IsNull(vgRegistro!Gls_FonoBen) Then
                vlGlsFonoBen = (vgRegistro!Gls_FonoBen)
            Else
                vlGlsFonoBen = ""
            End If
            If Not IsNull(vgRegistro!Gls_CorreoBen) Then
                vlGlsCorreoBen = (vgRegistro!Gls_CorreoBen)
            Else
                vlGlsCorreoBen = ""
            End If
            vlCodGruFam = (vgRegistro!Cod_GruFam)
            vlCodPar = flObtDescripcion(stCodPar, vgRegistro!Cod_Par) '(vgRegistro!Cod_Par)
            vlCodSexo = (vgRegistro!Cod_Sexo)
            vlCodSitInv = (vgRegistro!Cod_SitInv)
            vlCodDerCre = (vgRegistro!Cod_DerCre)
            vlCodDerpen = flObtDescripcion(stDerpen, vgRegistro!Cod_DerPen) '(vgRegistro!Cod_DerPen)
            vlCodCauInv = (vgRegistro!Cod_CauInv)
            vlFecNacBen = (vgRegistro!Fec_NacBen)
            If Not IsNull(vgRegistro!Fec_NacHM) Then
                vlFecNacHM = (vgRegistro!Fec_NacHM)
            Else
                vlFecNacHM = ""
            End If
            If Not IsNull(vgRegistro!Fec_InvBen) Then
                vlFecInvBen = (vgRegistro!Fec_InvBen)
            Else
                vlFecInvBen = ""
            End If
            vlCodMotReqPen = (vgRegistro!Cod_MotReqPen)
            vlMtoPensionB = CStr(vgRegistro!Mto_Pension)
            vlMtoPensionGar = CStr(vgRegistro!Mto_PensionGar)
            vlPrcPension = CStr(vgRegistro!Prc_Pension)
            vlCodInsSalud = flObtDescripcion(stInsSal, vgRegistro!Cod_InsSalud) '(vgRegistro!Cod_InsSalud)
            vlCodModSalud = (vgRegistro!Cod_ModSalud)
            vlMtoPlanSalud = CStr(vgRegistro!Mto_PlanSalud)
            vlCodEstPension = flObtDescripcion(stEstPen, vgRegistro!Cod_EstPension) '(vgRegistro!Cod_EstPension)
            vlCodViaPago = flObtDescripcion(stViaPago, vgRegistro!Cod_ViaPago) '(vgRegistro!Cod_ViaPago)
            vlCodBanco = flObtDescripcion(stBanco, vgRegistro!Cod_Banco) '(vgRegistro!Cod_Banco)
            vlCodTipCuenta = flObtDescripcion(stTipCta, vgRegistro!Cod_TipCuenta) '(vgRegistro!Cod_TipCuenta)
            If Not IsNull(vgRegistro!Num_Cuenta) Then
                vlNumCuenta = (vgRegistro!Num_Cuenta)
            Else
                vlNumCuenta = ""
            End If
            vlCodSucursal = flObtDescripcion(stSuc, vgRegistro!Cod_Sucursal) '(vgRegistro!Cod_Sucursal)
            If Not IsNull(vgRegistro!Fec_FallBen) Then
                vlFecFallBen = (vgRegistro!Fec_FallBen)
            Else
                vlFecFallBen = ""
            End If
            If Not IsNull(vgRegistro!Cod_CauSusBen) Then
                vlCodCauSusBen = (vgRegistro!Cod_CauSusBen)
            Else
                vlCodCauSusBen = ""
            End If
            If Not IsNull(vgRegistro!Fec_SusBen) Then
                vlFecSusBen = (vgRegistro!Fec_SusBen)
            Else
                vlFecSusBen = ""
            End If
            vlFecIniPagoPenB = (vgRegistro!Fec_IniPagoPen)
            If Not IsNull(vgRegistro!Fec_TerPagoPenGar) Then
                vlFecTerPagoPenGarB = (vgRegistro!Fec_TerPagoPenGar)
            Else
                vlFecTerPagoPenGarB = ""
            End If
                
            'Generar Línea de Registro de Datos de Beneficiarios
                      
            vlLinea = vlNumPoliza & ";" & vlNumEndoso & ";" & vlCodAFP & ";" & _
                      vlCodTipPension & ";" & vlCodEstado & ";" & vlCodTipRen & ";" & _
                      vlCodModalidad & ";" & vlNumCargas & ";" & vlFecVigencia & ";" & _
                      vlFecTerVigencia & ";" & vlMtoPrima & ";" & vlMtoPension & ";" & _
                      vlNumMesDif & ";" & vlNumMesGar & ";" & vlPrcTasaCe & ";" & _
                      vlPrcTasaVta & ";" & vlPrcTasaCtoRea & ";" & vlPrcTasaIntPerGar & ";" & _
                      vlFecIniPagoPen & ";" & _
                      vlNumOrden & ";" & vlFecIngreso & ";" & vlCodTipoIdenBen & ";" & vlNumIden & ";" & _
                      vlGlsNomBen & ";" & vlGlsSegNomBen & ";" & vlGlsPatBen & ";" & vlGlsMatBen & ";" & _
                      vlCodDireccion & ";" & vlGlsDirBen & ";" & vlGlsComuna & ";" & _
                      vlGlsProvincia & ";" & vlGlsRegion & ";" & vlGlsFonoBen & ";" & _
                      vlGlsCorreoBen & ";" & vlCodGruFam & ";" & vlCodPar & ";" & _
                      vlCodSexo & ";" & vlCodSitInv & ";" & vlCodDerCre & ";" & _
                      vlCodDerpen & ";" & vlCodCauInv & ";" & vlFecNacBen & ";" & _
                      vlFecNacHM & ";" & vlFecInvBen & ";" & vlCodMotReqPen & ";" & _
                      vlMtoPensionB & ";" & vlMtoPensionGar & ";" & vlPrcPension & ";" & _
                      vlCodInsSalud & ";" & vlCodModSalud & ";" & vlMtoPlanSalud & ";" & _
                      vlCodEstPension & ";" & vlCodViaPago & ";" & _
                      vlCodBanco & ";" & vlCodTipCuenta & ";" & vlNumCuenta & ";" & _
                      vlCodSucursal & ";" & vlFecFallBen & ";" & _
                      vlCodCauSusBen & ";" & vlFecSusBen & ";" & vlFecIniPagoPenB & ";" & _
                      vlFecTerPagoPenGarB
                      

            Print #1, vlLinea
            
            vgRegistro.MoveNext
            
            If Frm_BarraProg.ProgressBar1.Value + vlAumento < 100 Then
                Frm_BarraProg.ProgressBar1.Value = Frm_BarraProg.ProgressBar1.Value + vlAumento
            End If

        Wend

        Close #1

        Unload Frm_BarraProg
        Screen.MousePointer = 0
        vlOpen = False
        MsgBox "La Exportación de Datos al Archivo ha sido Finalizada Exitosamente.", vbInformation, "Estado de Exportación"
        Screen.MousePointer = vbDefault

    Else
        MsgBox "No existe Información para este Rango de Fechas", vbInformation, "Operacion Cancelada"
        Exit Function
    End If
    
Exit Function
Err_flExportarDatosBen:
Screen.MousePointer = vbDefault
'Error por hacer click en boton cancelar de pantalla guardar como
If Err.Number = 32755 Then
    Exit Function
Else
    If vlOpen Then
        Close #1
    End If
    MsgBox "Se ha producido el siguiente error : " & Err.Description, vbCritical, "Error"
End If

End Function

Function flLimpiarVariables()
On Error GoTo Err_flLimpiarVariables

    'Datos de Poliza
    vlNumPoliza = ""
    vlNumEndoso = ""
    vlCodAFP = ""
    vlCodTipPension = ""
    vlCodEstado = ""
    vlCodTipRen = ""
    vlCodModalidad = ""
    vlNumCargas = ""
    vlFecVigencia = ""
    vlFecTerVigencia = ""
    vlMtoPrima = ""
    vlMtoPension = ""
    vlNumMesDif = ""
    vlNumMesGar = ""
    vlPrcTasaCe = ""
    vlPrcTasaVta = ""
    vlPrcTasaCtoRea = ""
    vlPrcTasaIntPerGar = ""
    vlFecIniPagoPen = ""
    vlCodTipOrigen = ""
    vlNumIndQuiebra = ""
    'Datos de Beneficiario
    vlNumOrden = ""
    vlFecIngreso = ""
    vlCodTipoIdenBen = ""
    vlNumIden = ""
    vlGlsNomBen = ""
    vlGlsSegNomBen = ""
    vlGlsPatBen = ""
    vlGlsMatBen = ""
    vlGlsDirBen = ""
    vlCodDireccion = ""
    
    vlGlsComuna = ""
    vlGlsProvincia = ""
    vlGlsRegion = ""
    
    vlGlsFonoBen = ""
    vlGlsCorreoBen = ""
    vlCodGruFam = ""
    vlCodPar = ""
    vlCodSexo = ""
    vlCodSitInv = ""
    vlCodDerCre = ""
    vlCodDerpen = ""
    vlCodCauInv = ""
    vlFecNacBen = ""
    vlFecNacHM = ""
    vlFecInvBen = ""
    vlCodMotReqPen = ""
    vlMtoPensionB = ""
    vlMtoPensionGar = ""
    vlPrcPension = ""
    vlCodInsSalud = ""
    vlCodModSalud = ""
    vlMtoPlanSalud = ""
    vlCodEstPension = ""
    vlCodViaPago = ""
    vlCodBanco = ""
    vlCodTipCuenta = ""
    vlNumCuenta = ""
    vlCodSucursal = ""
    vlFecFallBen = ""
    vlCodCauSusBen = ""
    vlFecSusBen = ""
    vlFecIniPagoPenB = ""
    vlFecTerPagoPenGarB = ""
    
Exit Function
Err_flLimpiarVariables:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function flBuscaNombreComuna(Codigo As Integer) As String

    flBuscaNombreComuna = ""

    vgSql = ""
    vgSql = vgSql & "SELECT gls_comuna "
    vgSql = vgSql & "FROM MA_TPAR_COMUNA "
    vgSql = vgSql & "WHERE cod_direccion = " & Trim(Codigo) & " "
    Set vgRs4 = vgConexionBD.Execute(vgSql)
    If Not vgRs4.EOF Then
       flBuscaNombreComuna = (vgRs4!gls_comuna)
    Else
        flBuscaNombreComuna = ""
    End If

End Function

Private Sub Form_Load()
On Error GoTo Err_Carga
    
    Frm_CargaArchBen.Left = 0
    Frm_CargaArchBen.Top = 0
    Lbl_FechaActual = ""
    Lbl_FechaActual = fgBuscaFecServ
    
Exit Sub
Err_Carga:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Function flCargaCodigos()
    
    Call flCargaMatriz(stAfp, vgCodTabla_AFP)           'Afp
    Call flCargaMatriz(stTipPen, vgCodTabla_TipPen)     'Tipo de Pensión
    Call flCargaMatriz(stEstPol, vgCodTabla_TipVigPol)  'Estado de la Poliza
    Call flCargaMatriz(stTipRen, vgCodTabla_TipRen)     'Tipo de Renta
    Call flCargaMatriz(stMod, vgCodTabla_AltPen)        'Modalidad
    Call flCargaMatriz(stCodPar, vgCodTabla_Par)        'Parentesco
    Call flCargaMatriz(stDerpen, vgCodTabla_DerPen)     'Derecho a Pensión
    Call flCargaMatriz(stInsSal, vgCodTabla_InsSal)     'Salud
    Call flCargaMatriz(stEstPen, vgCodTabla_TipVig)     'Estado de la Pensión
    Call flCargaMatriz(stViaPago, vgCodTabla_ViaPago)   'Via Pago
    Call flCargaMatriz(stBanco, vgCodTabla_Bco)         'Banco
    Call flCargaMatriz(stTipCta, vgCodTabla_TipCta)     'Tipo de Cuenta
    Call flCargaTipoIden(stTipIden)                     'Tipo de Identidad
    Call flCargaSucursal(stSuc)                         'Sucursal

End Function

Private Function flCargaMatriz(iMatriz, iCodTabla As String)
'Carga en la matriz los codigos existentes en la BD
On Error GoTo Err_CargaMatriz
Dim vlcont, i As Integer

    vgSql = ""
    vgSql = "select count(1) as contador from ma_tpar_tabcod "
    vgSql = vgSql & "where cod_tabla='" & iCodTabla & "' "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
        vlcont = vgRs!contador
    End If
    vgRs.Close
    
    ReDim iMatriz(vlcont, 2) As String
    
    i = 1
    vgSql = ""
    vgSql = "select cod_elemento,gls_elemento from ma_tpar_tabcod "
    vgSql = vgSql & "where cod_tabla='" & iCodTabla & "'"
    vgSql = vgSql & "order by cod_elemento,gls_elemento "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
        While Not vgRs.EOF
            iMatriz(i, 1) = vgRs!cod_elemento
            iMatriz(i, 2) = vgRs!gls_elemento
            i = i + 1
            vgRs.MoveNext
        Wend
    End If
    vgRs.Close
    
Exit Function
Err_CargaMatriz:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Private Function flCargaSucursal(iMatriz)
'Carga en la matriz los codigos de sucursal existentes en la BD
On Error GoTo Err_CargaSuc
Dim vlcont, i As Integer

    vgSql = ""
    vgSql = "select count(1) as contador from ma_tpar_sucursal "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
        vlcont = vgRs!contador
    End If
    vgRs.Close
    
    ReDim iMatriz(vlcont, 2) As String
    
    i = 1
    vgSql = ""
    vgSql = "select cod_sucursal,gls_sucursal from ma_tpar_sucursal "
    vgSql = vgSql & "order by cod_sucursal,gls_sucursal "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
        While Not vgRs.EOF
            iMatriz(i, 1) = vgRs!Cod_Sucursal
            iMatriz(i, 2) = vgRs!gls_sucursal
            i = i + 1
            vgRs.MoveNext
        Wend
    End If
    vgRs.Close
    
Exit Function
Err_CargaSuc:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Private Function flCargaTipoIden(iMatriz)
'Carga en la matriz los codigos de tipo de identidad existentes en la BD
On Error GoTo Err_CargaTipIden
Dim vlcont, i As Integer

    vgSql = ""
    vgSql = "select count(1) as contador from ma_tpar_tipoiden "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
        vlcont = vgRs!contador
    End If
    vgRs.Close
    
    ReDim iMatriz(vlcont, 2) As String
    
    i = 1
    vgSql = ""
    vgSql = "select cod_tipoiden,gls_tipoidencor from ma_tpar_tipoiden "
    vgSql = vgSql & "order by cod_tipoiden,gls_tipoidencor "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
        While Not vgRs.EOF
            iMatriz(i, 1) = vgRs!cod_tipoiden
            iMatriz(i, 2) = vgRs!gls_tipoidencor
            i = i + 1
            vgRs.MoveNext
        Wend
    End If
    vgRs.Close
    
Exit Function
Err_CargaTipIden:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Private Function flObtDescripcion(iMatriz, iCodigo As String) As String
'Permite obtener la descripción del codigo de entrada
'Parametros de Entrada:
'-iMatriz: Matriz de códigos
'-iCodigo: código a buscar en la matriz
On Error GoTo Err_ObtDes
Dim i As Integer

    For i = 1 To UBound(iMatriz)
        If (iCodigo = iMatriz(i, 1)) Then
            flObtDescripcion = iMatriz(i, 2)
            Exit Function
        End If
    Next
    flObtDescripcion = ""
    
Exit Function
Err_ObtDes:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.MDIForm Frm_Menu 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Sistema de Pago de Pensiones de Renta Vitalicia"
   ClientHeight    =   5610
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   12555
   Icon            =   "MDIForm.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   3  'Windows Default
   Tag             =   "0"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   5340
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   6509
            Text            =   "                                                                                "
            TextSave        =   "                                                                                "
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "07/08/2023"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "11:52 a.m."
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3678
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3678
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "MAYÚS"
         EndProperty
      EndProperty
   End
   Begin VB.Menu Mnu_AdmSistema 
      Caption         =   "Administración del &Sistema"
      Begin VB.Menu Mnu_SisUsuarios 
         Caption         =   "&Usuarios"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_SisContrasena 
         Caption         =   "&Contraseña"
      End
      Begin VB.Menu Mnu_SisNivel 
         Caption         =   "&Nivel de Acceso"
      End
      Begin VB.Menu mnuSepara0 
         Caption         =   "-"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMantenciondeInformacion 
      Caption         =   "&Mantención de Información"
      Begin VB.Menu mnuAntecedentePensionado 
         Caption         =   "&Antecedentes Pensionado"
         Begin VB.Menu mnuGenerales 
            Caption         =   "&Mantención Antecedentes Generales"
         End
         Begin VB.Menu mnuTutores 
            Caption         =   "&Asignación de Tutor/Apoderado"
         End
         Begin VB.Menu mnuCertificado 
            Caption         =   "&Certificados de Estudios"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuCertificadoSup 
            Caption         =   "&Certificado de Supervivencia"
            Begin VB.Menu mnuMantenedorSuperv 
               Caption         =   "Mantenedor"
            End
            Begin VB.Menu mnuCargarExcel 
               Caption         =   "Cargar Automatica"
            End
            Begin VB.Menu mnuReliquidacion 
               Caption         =   "Reliquidación"
            End
         End
         Begin VB.Menu mnuseparaGenArchBen 
            Caption         =   "-"
         End
         Begin VB.Menu mnuGenArchBen 
            Caption         =   "Generación Archivo Datos Beneficiarios"
         End
      End
      Begin VB.Menu mnuRetencion 
         Caption         =   "&Retención Judicial"
         Begin VB.Menu mnuRetIngreso 
            Caption         =   "&Ingreso de Orden Judicial"
         End
         Begin VB.Menu mnuSepara10 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRetInforme 
            Caption         =   "&Informe de Control"
         End
      End
      Begin VB.Menu mnuAsigFam 
         Caption         =   "Asignación &Familiar"
         Visible         =   0   'False
         Begin VB.Menu mnuValoresAF 
            Caption         =   "Mantenedor &Valores de Cargas Familiares"
         End
         Begin VB.Menu mnuDeclaracionIngresos 
            Caption         =   "&Ingreso de Declaración de Ingresos"
         End
         Begin VB.Menu mnuAFNoBenef 
            Caption         =   "&Mantenedor No Beneficiarios"
         End
         Begin VB.Menu mnuActivacionAF 
            Caption         =   "&Activación y Desactivación de Carga Familiar"
         End
         Begin VB.Menu mnuSepara 
            Caption         =   "-"
         End
         Begin VB.Menu mnuReporteControlAF 
            Caption         =   "&Informe de Control"
         End
      End
      Begin VB.Menu mnuGarantiaEstatal 
         Caption         =   "&Garantía Estatal"
         Visible         =   0   'False
         Begin VB.Menu mnuGEPenMinimas 
            Caption         =   "&Mantención Pensiones Mínimas"
         End
         Begin VB.Menu mnuGEPorDeduccion 
            Caption         =   "Cálculo &Porcentaje de Deducción"
         End
         Begin VB.Menu mnuGESegBenPensionado 
            Caption         =   "&Seguimiento del Beneficio por Pensionado"
         End
         Begin VB.Menu mnuGEDescuento 
            Caption         =   "Haberes y Descuentos de Garantía Estatal"
         End
         Begin VB.Menu mnusepara1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuGECargaArchivo 
            Caption         =   "&Carga Archivos"
            Begin VB.Menu mnuGECargaRecuperos 
               Caption         =   "Recuperos"
            End
            Begin VB.Menu mnuGECargaOtrosRec 
               Caption         =   "Otros Recuperos"
            End
         End
         Begin VB.Menu mnuGEInformesConciliacion 
            Caption         =   "&Informes de Conciliación"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuSepara2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuGEInformesControl 
            Caption         =   "In&formes de Control"
         End
      End
      Begin VB.Menu mnuCCAF 
         Caption         =   "&Cajas de Compensación"
         Visible         =   0   'False
         Begin VB.Menu mnuCarga 
            Caption         =   "Carga &Automática"
         End
         Begin VB.Menu mnuManual 
            Caption         =   "Ingreso &Manual"
         End
         Begin VB.Menu mnusepara12 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCCInforme 
            Caption         =   "&Informe de Control"
         End
         Begin VB.Menu mnuSeparaCCAF2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuModCargaCCAF 
            Caption         =   "Modificar Carga CCAF"
         End
      End
      Begin VB.Menu mnuSepara7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHabDescuento 
         Caption         =   "Ingreso de &Haberes y Descuentos"
         Begin VB.Menu mnuHabDescManual 
            Caption         =   "Ingreso Manual"
         End
         Begin VB.Menu mnuHabDescAuto 
            Caption         =   "Carga Automática"
         End
      End
      Begin VB.Menu mnuIngPresMed 
         Caption         =   "Ingreso de Préstamos &Médicos"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSepara6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMensajes 
         Caption         =   "Generación de &Mensajes"
         Begin VB.Menu mnuMsgIndividual 
            Caption         =   "&Individuales (por pensionado)"
         End
         Begin VB.Menu mnuMsgAutomaticos 
            Caption         =   "&Automáticos"
         End
         Begin VB.Menu mnuSepara5 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMsgCarga 
            Caption         =   "&Carga desde Archivo"
         End
      End
      Begin VB.Menu mnuSepara23 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPorcPension 
         Caption         =   "Porcentaje Disminución Pensión"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSepara4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEndosos 
         Caption         =   "&Endosos"
         Begin VB.Menu mnuEndososGenera 
            Caption         =   "Generar Endosos"
         End
         Begin VB.Menu mnuEndososConsulta 
            Caption         =   "Consulta de Endosos"
         End
         Begin VB.Menu mnu_separa45 
            Caption         =   "-"
         End
         Begin VB.Menu mnu_pago_heren 
            Caption         =   "Pago de Herencia"
         End
      End
   End
   Begin VB.Menu mnuCalculoPension 
      Caption         =   "&Generación Cálculo de Pensión"
      Begin VB.Menu mnuParametrosCalculo 
         Caption         =   "&Parámetros de Cálculo de Pensiones"
         Visible         =   0   'False
         Begin VB.Menu mnuParametroPriPagos 
            Caption         =   "&Primeros Pagos"
         End
      End
      Begin VB.Menu mnuParametro 
         Caption         =   "&Parámetros de Cálculo de Pagos Recurrentes"
      End
      Begin VB.Menu mnuPrimerasPensiones 
         Caption         =   "P&rimeros Pagos"
         Visible         =   0   'False
         Begin VB.Menu mnuPrimerProvisorio 
            Caption         =   "&Provisorio"
         End
         Begin VB.Menu mnuPrimerDefinitivo 
            Caption         =   "&Definitivo"
         End
      End
      Begin VB.Menu mnuPensionesEnRegimen 
         Caption         =   "Pensiones &Recurrentes"
         Begin VB.Menu mnuRegimenProvisorio 
            Caption         =   "&Provisorio"
         End
         Begin VB.Menu mnuRegimenDefinitivo 
            Caption         =   "&Definitivo"
         End
      End
      Begin VB.Menu mnuSeparar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEmiInformes 
         Caption         =   "&Emisión de Informes"
         Begin VB.Menu mnuInfConsultaHistorica 
            Caption         =   "Consulta Histórica de Haberes y Desctos."
            Visible         =   0   'False
         End
         Begin VB.Menu Mnuconsultapensionado 
            Caption         =   "Consulta por Pensionado"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuSepara8 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuValidacion 
            Caption         =   "Validaciones"
            Begin VB.Menu mnuCertEstudios 
               Caption         =   "Certificados de &Supervivencia por Caducar"
            End
            Begin VB.Menu mnuCertVenc 
               Caption         =   "Certificados de Supervivencia &Vencidos"
            End
            Begin VB.Menu mnuInfo18 
               Caption         =   "Beneficiarios por Cumplir 18 Años"
            End
            Begin VB.Menu mnuInfo24 
               Caption         =   "Beneficiarios por Cumplir 24 Años"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuCieInfTutores 
               Caption         =   "&Vencimiento de Tutores"
            End
            Begin VB.Menu mnuFallAct 
               Caption         =   "Endosos Fallecimiento y Activación"
            End
            Begin VB.Menu mnuBenGar18 
               Caption         =   "&Beneficiarios Garantizados mayores a 18 años"
            End
            Begin VB.Menu mnuHijosParNac 
               Caption         =   "&Beneficiario con Partida de Nacimiento"
            End
            Begin VB.Menu mnRentTempIni 
               Caption         =   "Renta Temporal a Iniciar"
            End
            Begin VB.Menu mnuHijosMaySInCert 
               Caption         =   "Hijos Mayores Sin Certificado de Estudios"
            End
         End
         Begin VB.Menu mnuLiqPago 
            Caption         =   "&Liquidación de Pago"
            Begin VB.Menu mnuLiqPensiones 
               Caption         =   "&Liquidaciones de Pensiones"
            End
            Begin VB.Menu mnuLiqPenCarta 
               Caption         =   "Liquidaciones de Pensiones Tipo Carta"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuLiqViasPagoRec 
               Caption         =   "Distribución de &Vías de Pago por Recep."
            End
            Begin VB.Menu mnuLiqViasPago 
               Caption         =   "Distribución de &Vías de Pago por Pens."
            End
            Begin VB.Menu mnuLiqconError 
               Caption         =   "&Pensiones con Error"
            End
            Begin VB.Menu mnuSeparaLiq 
               Caption         =   "-"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuLiqDepBancario 
               Caption         =   "&Generación Archivo Depósito Bancario"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuLiqChequeCont 
               Caption         =   "Generación Archivo Cheque Contable"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuLiqPrevired 
               Caption         =   "Generación Archivo Previred"
               Visible         =   0   'False
            End
         End
         Begin VB.Menu mnuGEInformes 
            Caption         =   "&Garantía Estatal"
            Visible         =   0   'False
            Begin VB.Menu mnuGEInfPenMinDef 
               Caption         =   "&Libro de Pensiones Minimas"
            End
            Begin VB.Menu mnuGEInformesPenMin 
               Caption         =   "Informes de Pensiones Minimas"
            End
            Begin VB.Menu mnuGEInfPenMinSinDer 
               Caption         =   "&Pensiones Mínimas sin Derecho"
            End
            Begin VB.Menu mnuGEInfHabDesDef 
               Caption         =   "&Informe de Haberes y Desctos de GE"
            End
            Begin VB.Menu mnuGEEstadoBen 
               Caption         =   "Informe por Estado del &Beneficiario"
            End
            Begin VB.Menu mnuGEConciliacion 
               Caption         =   "&Conciliación Garantía Estatal"
            End
            Begin VB.Menu mnuGEInfExcesos 
               Caption         =   "&Excesos Garantía Estatal"
            End
            Begin VB.Menu mnuGEInfDeficit 
               Caption         =   "&Déficit Garantía Estatal"
            End
            Begin VB.Menu mnuGEdePago 
               Caption         =   "Informe de &Pago"
               Visible         =   0   'False
            End
         End
         Begin VB.Menu mnuInfAF 
            Caption         =   "&Asignación Familiar"
            Visible         =   0   'False
            Begin VB.Menu mnuAFInfPagos 
               Caption         =   "&Nómina de Pago Mensual"
            End
         End
         Begin VB.Menu mnuCC 
            Caption         =   "&Cajas de Compensación"
            Visible         =   0   'False
            Begin VB.Menu mnuCCAportes 
               Caption         =   "&Pago de Aportes"
            End
            Begin VB.Menu mnuCCCreditos 
               Caption         =   "Pago de &Créditos"
            End
            Begin VB.Menu mnuCCOtrasPrestaciones 
               Caption         =   "Pago de &Otras Prestaciones"
            End
            Begin VB.Menu mnuSepara9 
               Caption         =   "-"
            End
            Begin VB.Menu mnuCCArchivo 
               Caption         =   "&Archivo para Caja de Compensación"
            End
         End
         Begin VB.Menu mnuInfPlanSalud 
            Caption         =   "&Plan de Salud"
            Begin VB.Menu mnuPSInfCotizaciones 
               Caption         =   "Planilla de Pago de &Cotizaciones"
            End
            Begin VB.Menu mnuPSInfPagoSalud 
               Caption         =   "Resumen de Planilla de Pago de &Cotizaciones"
            End
            Begin VB.Menu mnuPSInfPrestamos 
               Caption         =   "Planilla de Pago de &Préstamos Médicos Fonasa"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuSeparaSalud 
               Caption         =   "-"
            End
            Begin VB.Menu mnuPSInfArchivoCotFonasa 
               Caption         =   "Generación Archivo Cotización &ESSALUD"
            End
            Begin VB.Menu mnuPSInfArchivoPresMed 
               Caption         =   "Generación Archivo Prestamos Médicos"
               Visible         =   0   'False
            End
         End
         Begin VB.Menu mnuInfContables 
            Caption         =   "&Informes Contables"
            Begin VB.Menu mnuConInfCentralizacion 
               Caption         =   "&Centralización Contable"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuConInfFecu 
               Caption         =   "&Datos para FECU"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuSeparaCenCont 
               Caption         =   "-"
            End
            Begin VB.Menu mnuConGenArchCenCont 
               Caption         =   "Generación Archivo Centralización Contable"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuConGenArchContable 
               Caption         =   "Generación Archivo Contable"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuConGenArchRentas 
               Caption         =   "Generación Archivo de Rentas"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuConInfReaseguros 
               Caption         =   "Contabilización &Reaseguros"
               Visible         =   0   'False
            End
         End
         Begin VB.Menu Mnu_InfAdministrativos 
            Caption         =   "Informes Administrativos"
            Begin VB.Menu Mnu_UnivPolizas 
               Caption         =   "Universo de Pólizas"
            End
            Begin VB.Menu Mnu_EndososPol 
               Caption         =   "Endosos"
            End
            Begin VB.Menu Mnu_PolEmitidas 
               Caption         =   "Pólizas Emitidas"
            End
            Begin VB.Menu Mnu_ConsolidPagos 
               Caption         =   "Consolidado de Pagos"
            End
         End
         Begin VB.Menu mnuInfSVS 
            Caption         =   "Informes &SBS"
            Begin VB.Menu Mnu_Anexo18 
               Caption         =   "SBS Anexo 18"
            End
            Begin VB.Menu mnu_oficioMultiple 
               Caption         =   "Oficio Multiple 24729"
            End
            Begin VB.Menu mnu_formato0730 
               Caption         =   "Formato 0730"
            End
            Begin VB.Menu mnu_pagosmensuales 
               Caption         =   "Trama AFP"
            End
            Begin VB.Menu mnu_tramaiai9 
               Caption         =   "Trama IAI9 y IAI10"
            End
            Begin VB.Menu mnuAFInfSSS 
               Caption         =   "&Informe Estadístico y Financiero Asig. Fam."
               Visible         =   0   'False
            End
            Begin VB.Menu mnuSVSInfPromedio 
               Caption         =   "Informe &Promedio de Pensiones Pagadas"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuInfCircular 
               Caption         =   "&Circular 1410"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuSVSInfTasas 
               Caption         =   "&Tasas Implícitas y Resumen de Pensiones"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuSepara3 
               Caption         =   "-"
            End
            Begin VB.Menu mnuSVSInfArchivoCir1815 
               Caption         =   "Archivo de Pólizas Vendidas"
               Visible         =   0   'False
            End
         End
         Begin VB.Menu mnuCieInforme 
            Caption         =   "C&ierre de Pensiones"
            Begin VB.Menu mnuLibroPensiones 
               Caption         =   "&Libro de Pensiones"
            End
            Begin VB.Menu mnuLibroPensionesInt 
               Caption         =   "Libro de Pensiones Interno"
            End
            Begin VB.Menu mnuLibroPensionesCont 
               Caption         =   "Libro de Pensiones Contable"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuCieInfHyD 
               Caption         =   "&Haberes y Descuentos"
            End
            Begin VB.Menu mnuInfImpuesto 
               Caption         =   "&Impuesto Único"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuCieInfRJ 
               Caption         =   "&Retenciones Judiciales"
            End
            Begin VB.Menu mnuAnexoDos 
               Caption         =   "&Anexo 2"
            End
         End
         Begin VB.Menu mnuInfQuiebra 
            Caption         =   "&Quiebra"
            Visible         =   0   'False
            Begin VB.Menu mnuInfQuiComparativo 
               Caption         =   "&Comparativo Pensión Normal y por Quiebra"
            End
         End
         Begin VB.Menu mnu_archivos_pdt 
            Caption         =   "Archivos PDT"
         End
         Begin VB.Menu mnuRegula 
            Caption         =   "Regularizaciones"
         End
      End
      Begin VB.Menu mnuSeparar13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRegistroPagosATerceros 
         Caption         =   "&Registro de Pagos a Terceros"
         Begin VB.Menu mnuRegistroPagosATerCuoMor 
            Caption         =   "Gastos de Sepelio"
         End
         Begin VB.Menu mnuRegistroPagosATerPerGar 
            Caption         =   "Periodo Garantizado"
         End
         Begin VB.Menu mnuSeparar5 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRegistroPagosATerConsulta 
            Caption         =   "Consulta de Pagos a Terceros"
         End
      End
      Begin VB.Menu mnuSeparar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArchContable 
         Caption         =   "&Archivos Contables"
         Begin VB.Menu mnuArchContablePagReg 
            Caption         =   "Pagos Recurrentes"
         End
         Begin VB.Menu mnuArchContableGtoSep 
            Caption         =   "Gastos de Sepelio"
         End
         Begin VB.Menu mnuArchContablePerGar 
            Caption         =   "Periodo Garantizado"
         End
         Begin VB.Menu mnuHabDes 
            Caption         =   "Reliquidación"
         End
      End
      Begin VB.Menu mnuSeparar11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCalendario 
         Caption         =   "&Calendario de Pagos "
         Visible         =   0   'False
         Begin VB.Menu mnuCalPrimerosPagos 
            Caption         =   "&Primeros Pagos"
         End
      End
      Begin VB.Menu mnuCalPagosCartera 
         Caption         =   "&Calendario de Pagos  Recurrentes"
      End
      Begin VB.Menu mnuseparaApeMan 
         Caption         =   "-"
      End
      Begin VB.Menu mnuApertura 
         Caption         =   "&Habilitar Reproceso"
         Visible         =   0   'False
         Begin VB.Menu mnuApeManPP 
            Caption         =   "Primeros &Pagos"
         End
      End
      Begin VB.Menu mnuApeManPR 
         Caption         =   "&Habilitar Reproceso de Pagos Recurrentes"
      End
      Begin VB.Menu mnu_separadorprueba 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_EndAutoPrueba 
         Caption         =   "Endosos Automáticos (Prueba)"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuConsultas 
      Caption         =   "Consultas"
      Begin VB.Menu mnuConsulta 
         Caption         =   "&Consulta General por Pensionado"
      End
      Begin VB.Menu mnuConsultaCtaCte 
         Caption         =   "Consulta Cta. Corr&iente"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnurepor 
      Caption         =   "&Reportes"
      Begin VB.Menu mnureporlog 
         Caption         =   "&Reporte de Log"
      End
   End
   Begin VB.Menu mnuAcerca 
      Caption         =   "&Acerca de..."
   End
   Begin VB.Menu mnuSalir 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "Frm_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
 
Private Sub MDIForm_Load()
On Error GoTo Err_Carga
'02-12-2002***********************************
'Frm_Menu.mnuAdmin.Enabled = False
'Frm_Menu.mnuParametros.Enabled = False
'Frm_Menu.mnuValeInf.Enabled = False
'Frm_Menu.mnuAcerca.Enabled = False
'*********************************************
On Error GoTo Siguiente

    If (vgNombreCortoCompania <> "") Then
        Frm_Menu.Caption = vgNombreCortoCompania & " - " & Frm_Menu.Caption
    End If
    
    Me.Picture = LoadPicture(App.Path & "\ModuloPensiones.bmp")
    StatusBar1.Panels(1) = "Sistema : " & vgNombreSistema
    StatusBar1.Panels(4) = "BD : " & vgNombreBaseDatos

Siguiente:
    If Err.Number <> 0 Then
        MsgBox "No se encontró el archivo: " & App.Path & "\Logo.bmp" & Chr(13) & "Se continuará con la carga del Sistema", vbExclamation
    End If
    
Exit Sub
Err_Carga:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub
Private Sub mnRentTempIni_Click()
    frmTemporalSinCert.Show
End Sub

Private Sub Mnu_Anexo18_Click()
    Screen.MousePointer = vbHourglass
    'Frm_InfoEstBeneficios.Show
    Frm_Circular18.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnu_archivos_pdt_Click()
    Frm_Archivos_PDT.Show
End Sub

Private Sub Mnu_ConsolidPagos_Click()
    Screen.MousePointer = vbHourglass
    Frm_RptConsolPagos.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnu_EndAutoPrueba_Click()
    On Error GoTo Errores
    vgFecIniPag = DateSerial(2005, 7, 1)
    vgFecPago = "20050720"
    vgTipoPago = "R" 'Pagos en Regimen
   
    Screen.MousePointer = 11

Errores:
    If Err.Number <> 0 Then
       ' vgConexionTransac.RollbackTrans
    Else
       ' vgConexionTransac.CommitTrans
    End If
End Sub



Private Sub mnu_oficio_multiple_Click()

End Sub

Private Sub Mnu_EndososPol_Click()
    Screen.MousePointer = vbHourglass
    Frm_RptListaEndosos.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnu_formato0730_Click()
Screen.MousePointer = vbHourglass

  Frm_InfoSBSPagosMensuales.lblTipo_rpt = "3"
  Frm_InfoSBSPagosMensuales.Show
    
  Screen.MousePointer = vbDefault
End Sub

Private Sub mnu_oficioMultiple_Click()
  Screen.MousePointer = vbHourglass

  Frm_InfoSBSPagosMensuales.lblTipo_rpt = "2"
  Frm_InfoSBSPagosMensuales.Show
    
  Screen.MousePointer = vbDefault
End Sub

Private Sub mnu_pago_heren_Click()
    Screen.MousePointer = vbHourglass
    Frm_EndosoHerencia.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnu_pagosmensuales_Click()
  Screen.MousePointer = vbHourglass

Frm_InfoSBSPagosMensuales.lblTipo_rpt = "1"
  Frm_InfoSBSPagosMensuales.Show
    
  Screen.MousePointer = vbDefault
End Sub

Private Sub Mnu_PolEmitidas_Click()
    Screen.MousePointer = vbHourglass
    Frm_RptPoliasEmitidas.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub Mnu_SisContrasena_Click()
    Screen.MousePointer = vbHourglass
    vgValorAr = 0
    Frm_SisContrasena.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub Mnu_SisNivel_Click()
    Screen.MousePointer = vbHourglass
    Frm_SisNivel.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub Mnu_SisUsuarios_Click()
    Screen.MousePointer = vbHourglass
    Frm_SisUsuario.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnu_tramaiai9_Click()

  Screen.MousePointer = vbHourglass
  Frm_TramaIAI910.Show
  Screen.MousePointer = vbDefault

End Sub

Private Sub Mnu_UnivPolizas_Click()
    Screen.MousePointer = vbHourglass
    Frm_RptUniversoPolizas.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuAcerca_Click()
    Screen.MousePointer = vbHourglass
    Frm_SisAbout.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuActivacionAF_Click()
    Screen.MousePointer = vbHourglass
    Frm_AFActivaDesactiva.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuAFInfPagos_Click()
    Screen.MousePointer = vbHourglass
    'Informe de Nómina Mensual de Asignación Familiar
    vgNombreInformeSeleccionado = "InfPagoMensual"
    Frm_PlanillaPago.Show
    Frm_PlanillaPago.Caption = "Nómina Mensual de Asignación Familiar."
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuAFInfSSS_Click()
    Screen.MousePointer = vbHourglass
    vgNombreInformeSeleccionado = "InfEstFin"
    Frm_AFInformeEstFin.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuAFNoBenef_Click()
    Screen.MousePointer = vbHourglass
    Frm_AFMantNoBenef.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuAnexoDos_Click()
    Screen.MousePointer = vbHourglass
    Frm_ExcelAnexo2.Show
    Frm_ExcelAnexo2.Caption = "Anexo 2"
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuApeManPP_Click()
    Screen.MousePointer = vbHourglass
    'Apertura Manual de Primeras pensiones
    'vgApeManualSel = "ApeManPP"
    Frm_AperturaPrimerPago.Show
    Frm_AperturaPrimerPago.Caption = "Primeros Pagos"
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuApeManPR_Click()
    Screen.MousePointer = vbHourglass
    'Apertura Manual de Pensiones en Regimen
    vgApeManualSel = "ApeManPR"
    Frm_AperturaManual.Show
    Frm_AperturaManual.Caption = "Pagos Recurrentes"
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuArchContableGtoSep_Click()
    Screen.MousePointer = 11
    ''Frm_ArchContable.Show
    Frm_ContableArch.flInicio ("GtoSep")
    Screen.MousePointer = 0
End Sub

Private Sub mnuArchContablePagReg_Click()
    Screen.MousePointer = 11
    ''Frm_ArchContable.Show
    Frm_ContableArch.flInicio ("PagRec")
    Screen.MousePointer = 0
End Sub

Private Sub mnuArchContablePerGar_Click()
    Screen.MousePointer = 11
    ''Frm_ArchContable.Show
    Frm_ContableArch.flInicio ("PerGar")
    Screen.MousePointer = 0
End Sub

Private Sub mnuBenGar18_Click()
    Screen.MousePointer = vbHourglass
    frmReporteBeneficiariosGarantizados.Show
    Screen.MousePointer = vbDefault
End Sub





Private Sub mnuCalPagosCartera_Click()
    Screen.MousePointer = vbHourglass
    Frm_PensCalendario.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuCalPrimerosPagos_Click()
    Screen.MousePointer = vbHourglass
    Frm_PensCalendarioPrimerPago.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub Mnucarga_Click()
    Screen.MousePointer = vbHourglass
    Frm_CCAFImportar.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuCargarExcel_Click()
frm_CargarExcel_Superv.Show
End Sub

Private Sub mnuCCArchivo_Click()
    Screen.MousePointer = vbHourglass
    'Generación de Archivo de Descuentos de CCAF
    vgNomInfSeleccionado = "InfGeneraArchCCAF"
    Frm_CCAFGeneraArchivoCCAF.Show
    Frm_CCAFGeneraArchivoCCAF.Caption = "Generación de Archivo de Descuentos para CCAF."
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuCCInforme_Click()
    Screen.MousePointer = 11
        
    vgTipoInforme = "CCAF"
    vgTituloInfControl = "Informe Control Cajas de Compensación"
    
    'Descarga y Carga el formulario para refrescar los datos
    Unload Frm_InformeControl
    Frm_InformeControl.Show
     
    Screen.MousePointer = 0
End Sub

Private Sub mnuCertEstudios_Click()
    Screen.MousePointer = vbHourglass
    'Informe Definitivo de Certificados de Estudio por Caducar
    vgNombreInformePeriodoSeleccionado = "InfCerEstCad"
    Frm_InformePeriodo.Command1.Visible = True
    Frm_InformePeriodo.Show
    Frm_InformePeriodo.Caption = "Certificados de Supervivencia por Caducar."
    Screen.MousePointer = vbDefault
End Sub

Private Sub Mnucertificado_Click()
    Screen.MousePointer = vbHourglass
    Frm_AntCertificado.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuCertificadoSup_Click()
'    Screen.MousePointer = vbHourglass
'    Frm_AntCertificadoSup.Show
'    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuCertVenc_Click()
Screen.MousePointer = vbHourglass
frmCertificadosVencidos.Show
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuCieInfHyD_Click()
    Screen.MousePointer = vbHourglass
    'Informe de Libro de Haberes y Descuentos
    vgNombreInformeSeleccionadoInd = "InfLibHabDes"
    Frm_PlanillaPensionado.Show
    Frm_PlanillaPensionado.Caption = "Informe de Libro de Haberes y Descuentos."
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuCieInfRJ_Click()
    Screen.MousePointer = vbHourglass
    'Informe Definitivo de Retenciones Judiciales
    vgNombreInformeSeleccionado = "InfDefRetJud"
    Frm_PlanillaPago.Show
    Frm_PlanillaPago.Caption = "Informe Definitivo de Retenciones Judiciales."
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuCieInfTutores_Click()
    Screen.MousePointer = vbHourglass
    'Informe de los Vencimientos de Tutores
    vgNombreInformeSeleccionado = "InfVenTut"
    Frm_PlanillaPago.Show
    Frm_PlanillaPago.Caption = "Informe de Vencimiento de Tutores."
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuConGenArchCenCont_Click()
    Screen.MousePointer = vbHourglass
    'Generación de Archivo de Deposito Bancario
    vgNomInfSeleccionado = "InfGeneraArchCenCont"
    Frm_CargaArchivo.Show
    Frm_CargaArchivo.Caption = "Generación de Archivo de Centralización Contable."
    Frm_CargaArchivo.Lbl_FecProceso.Visible = True
    Frm_CargaArchivo.Txt_FecProceso.Visible = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuConGenArchContable_Click()
    Screen.MousePointer = vbHourglass
    'Generación de Archivo Contable
    vgNomInfSeleccionado = "InfGeneraArchContable"
    Frm_CargaArchivo.Show
    Frm_CargaArchivo.Caption = "Generación de Archivo Contable."
    Frm_CargaArchivo.Lbl_FecProceso.Visible = False
    Frm_CargaArchivo.Txt_FecProceso.Visible = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuConGenArchRentas_Click()
    Screen.MousePointer = vbHourglass
    'Generación de Archivo de Rentas
    vgNomInfCargaArchivoAnno = "InfGeneraArchRenta"
    Frm_CargaArchivoAnno.Show
    Frm_CargaArchivoAnno.Caption = "Generación de Archivo de Rentas."
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuConInfCentralizacion_Click()
    Screen.MousePointer = vbHourglass
    'Informe de Vías de Pago de Pensiones
    vgNombreInformeSeleccionado = "InfCenContable"
    Frm_PlanillaPago.Show
    Frm_PlanillaPago.Caption = "Informe de Centralización Contable."
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuConInfFecu_Click()
    Screen.MousePointer = vbHourglass
    'Informe de Datos para FECU (Contable)
    vgNombreInformePlanillaCont = "InfFECUContable"
    Frm_PlanillaCont.Show
    Frm_PlanillaCont.Caption = "Informe de Datos para FECU."
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuConsulta_Click()
    Screen.MousePointer = 11
    Frm_Consulta.Show
    Screen.MousePointer = 0
End Sub

Private Sub mnuConsultaCtaCte_Click()
    Screen.MousePointer = 11
    Frm_CtaCorriente.Show
    Screen.MousePointer = 0
End Sub

Private Sub Mnuconsultapensionado_Click()
    Screen.MousePointer = 11
    Frm_ConPensionado.Show
    Screen.MousePointer = 0
End Sub

Private Sub Mnudeclaracioningresos_Click()
    Screen.MousePointer = vbHourglass
    Frm_AFIngresos.Show
    Screen.MousePointer = vbDefault
End Sub

'Private Sub mnuEndosos_Click()
'    Screen.MousePointer = vbHourglass
'    Frm_EndosoPol.Show
'    Screen.MousePointer = vbDefault
'End Sub

Private Sub mnuEndososConsulta_Click()
    Screen.MousePointer = vbHourglass
    Frm_ConsultaEndoso.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuEndososGenera_Click()
    Screen.MousePointer = vbHourglass
    Frm_EndosoPol.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuFallAct_Click()
    'frm_rep_Endoso.Show
    
    Screen.MousePointer = vbHourglass
    If Me.Tag = "0" Then
            frm_rep_Endoso.Tag = "A"
            frm_rep_Endoso.Show
            Frm_Menu.Tag = "2"
    Else
        If Me.Tag = "1" Then
            MsgBox "Debe Cerrar ventana de Primeros Pagos para poder Ejecutar esta Opción", vbCritical, Me.Caption
        Else
            MsgBox "Debe Cerrar ventana de Pagos en Régiman para poder Ejecutar esta Opción", vbCritical, Me.Caption
        End If
    End If
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub mnuHijosMaySInCert_Click()
frmMayoresSinCertEst.Show

End Sub

Private Sub mnuHijosParNac_Click()

    Screen.MousePointer = vbHourglass
    If Me.Tag = "0" Then
            frm_rep_Endoso.Tag = "B"
            frm_rep_Endoso.Show
            Frm_Menu.Tag = "1"
    Else
        If Me.Tag = "2" Then
            MsgBox "Debe Cerrar ventana de Primeros Pagos para poder Ejecutar esta Opción", vbCritical, Me.Caption
        Else
            MsgBox "Debe Cerrar ventana de Pagos de Cartera para poder Ejecutar esta Opción", vbCritical, Me.Caption
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuGECargaOtrosRec_Click()
    Screen.MousePointer = vbHourglass
    Frm_GECargaRecupOtro.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuGECargaRecuperos_Click()
    Screen.MousePointer = vbHourglass
    Frm_GECargaRecup.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuGEConciliacion_Click()
    Screen.MousePointer = vbHourglass
    'Informe Definitivo de Conciliacion de Garantia Estatal
    vgNombreInfSeleccionadoProceso = "InfConciliacion"
    vgPalabra = ""
    Frm_PlanillaProceso.Show
    Frm_PlanillaProceso.Caption = "Informe Definitivo de Conciliación de Garantía Estatal."
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuGEDescuento_Click()
    Screen.MousePointer = vbHourglass
    Frm_GEDesctoExceso.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuGEInfHabDes_Click()
    Screen.MousePointer = 11

    vgTipoInforme = "HD"
    vgTituloInfControl = "Informe de Haberes y Descuentos de Garantía Estatal"

    'Descarga y Carga el formulario para refrescar los datos
    Unload Frm_InformeControl
    Frm_InformeControl.Show

    Screen.MousePointer = 0
End Sub



Private Sub mnuGEEstadoBen_Click()
    Screen.MousePointer = vbHourglass
    'Informe Definitivo por Estado Beneficiairo de G.E.
    'vgNombreInformeSeleccionadoInd = "InfConciliacion"
    Frm_PlanillaEstadoBenGE.Show
    Frm_PlanillaEstadoBenGE.Caption = "Informe Definitivo por Estado Beneficiario de G.E."
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuGEInfDeficit_Click()
    Screen.MousePointer = vbHourglass
    'Informe Definitivo de Pagos efectuados de menos de Garantia Estatal
    vgNombreInfSeleccionadoProceso = "InfPagoMenos"
    vgPalabra = ""
    Frm_PlanillaProceso.Show
    Frm_PlanillaProceso.Caption = "Informe Definitivo de Pagos Efectuados de Menos de Garantía Estatal."
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuGEInfExcesos_Click()
    Screen.MousePointer = vbHourglass
    'Informe Definitivo de Pagos efectuados en exceso de Garantia Estatal
    vgNombreInfSeleccionadoProceso = "InfPagoExceso"
    vgPalabra = ""
    Frm_PlanillaProceso.Show
    Frm_PlanillaProceso.Caption = "Informe Definitivo de Pagos Efectuados En Exceso de Garantía Estatal."
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuGEInfHabDesDef_Click()
    Screen.MousePointer = vbHourglass
    'Informe Definitivo de Haberes y Descuentos de Garantía Estatal
    vgNombreInformeSeleccionado = "InfHabDesGE"
    Frm_PlanillaPago.Show
    Frm_PlanillaPago.Caption = "Informe Definitivo de Haberes y Descuentos de Garantía Estatal."
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuGEInfPenMin_Click()
    Screen.MousePointer = 11

    vgTipoInforme = "PM"
    vgTituloInfControl = "Informe Libro de Pensiones Mínimas"

    'Descarga y Carga el formulario para refrescar los datos
    Unload Frm_InformeControl
    Frm_InformeControl.Show

    Screen.MousePointer = 0
End Sub

Private Sub mnuGEInformesControl_Click()
    Screen.MousePointer = 11

    vgTipoInforme = "GE"
    vgTituloInfControl = "Informes de Control de Garantía Estatal"

    'Descarga y Carga el formulario para refrescar los datos
    Unload Frm_InformeControl
    Frm_InformeControl.Show

    Screen.MousePointer = 0
End Sub

Private Sub mnuGEInformesPenMin_Click()
    Screen.MousePointer = vbHourglass
    'Informes Varios de Pensiones Minimas
    Frm_PlanillaPenMin.Show
    Frm_PlanillaPenMin.Caption = "Informes de Pensiones Minimas."
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuGEInfPenMinDef_Click()
    Screen.MousePointer = vbHourglass
    'Informe Definitivo de Libro de Pensiones Minimas
    vgNombreInformeSeleccionado = "InfLibroPenMin"
    Frm_PlanillaPago.Show
    Frm_PlanillaPago.Caption = "Informe Definitivo de Libro de Pensiones Minimas."
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuGEInfPenMinSinDer_Click()
    Screen.MousePointer = vbHourglass
    'Informe Definitivo de Pensiones Mínimas sin Derecho
    vgNombreInformeSeleccionado = "InfPenMinSinDer"
    Frm_PlanillaPago.Show
    Frm_PlanillaPago.Caption = "Informe Definitivo de Pensiones Mínimas sin Derecho."
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuGenArchBen_Click()
    Screen.MousePointer = vbHourglass
    'Generación de Archivo de Datos de Beneficiarios
    vgNomInfSeleccionado = "InfGeneraArchDatosBen"
    Frm_CargaArchBen.Show
    Frm_CargaArchBen.Caption = "Generación de Archivo de Datos de Beneficiarios."
    Screen.MousePointer = vbDefault
End Sub

Private Sub Mnugenerales_Click()
    Screen.MousePointer = vbHourglass
    Frm_AntPensionado.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuGEPenMinimas_Click()
    Screen.MousePointer = vbHourglass
    Frm_GEPensionesMinimas.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuGEPorDeducción_Click()
    Screen.MousePointer = vbHourglass
    Frm_GECalculoPorcentaje.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuGEPorDeduccion_Click()
    Screen.MousePointer = vbHourglass
    Frm_GECalculoPorcentaje.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuGESegBenPensionado_Click()
    Screen.MousePointer = vbHourglass
    Frm_GESeguimiento.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuHabDes_Click()
    Screen.MousePointer = 11
    ''Frm_ArchContable.Show
    Frm_ContableArch.flInicio ("HabDes")
    Screen.MousePointer = 0
End Sub

Private Sub mnuHabDescAuto_Click()
    Screen.MousePointer = vbHourglass
    Frm_HabDesImportar.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuHabDescManual_Click()
    Screen.MousePointer = vbHourglass
    Frm_PensHabDescto.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuHabDescuento_Click()
'    Screen.MousePointer = vbHourglass
'    Frm_PensHabDescto.Show
'    Screen.MousePointer = vbDefault
End Sub


Private Sub mnuInfConsultaHistorica_Click()
    Screen.MousePointer = vbHourglass
    Frm_ConHabDescto.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuInfImpuesto_Click()
    Screen.MousePointer = vbHourglass
    'Informe de Libro de Impuesto Único
    vgNombreInformeSeleccionado = "InfImpUni"
    Frm_PlanillaPago.Show
    Frm_PlanillaPago.Caption = "Informe de Impuesto Único"
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuInfo18_Click()
    Screen.MousePointer = vbHourglass
    'Informe Definitivo de Beneficiarios por ciumplir 18 años de edad
    vgNombreInformePeriodoSeleccionado = "InfBen18"
    'Frm_InformePeriodo.Command1.Visible = False
    Frm_InformePeriodo.Show
    Frm_InformePeriodo.Caption = "Beneficiarios por Cumplir 18 Años."
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuInfo24_Click()
    Screen.MousePointer = vbHourglass
    'Informe Definitivo de Beneficiarios por ciumplir 24 años de edad
    vgNombreInformePeriodoSeleccionado = "InfBen24"
    Frm_InformePeriodo.Show
    Frm_InformePeriodo.Caption = "Beneficiarios por Cumplir 24 Años."
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuIngPresMed_Click()
    Screen.MousePointer = vbHourglass
    Frm_PensPresMed.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuLibroPensiones_Click()
    Screen.MousePointer = vbHourglass
    'Informe de Libro de Pensiones
    vgNombreInformeSeleccionado = "InfLibPen"
    Frm_PlanillaPago.Show
    Frm_PlanillaPago.Caption = "Informe de Libro de Pensiones."
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuLibroPensionesCont_Click()
    Screen.MousePointer = vbHourglass
    'Informe de Libro de Pensiones Contable (Ordenado por Tipo de Pensión)
    vgNombreInformeSeleccionado = "InfLibPenCont"
    Frm_PlanillaPago.Show
    Frm_PlanillaPago.Caption = "Informe de Libro de Pensiones Contable."
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuLibroPensionesInt_Click()
    Screen.MousePointer = vbHourglass
    'Informe de Libro de Pensiones Interno
    vgNombreInformeSeleccionado = "InfLibPenInt"
    Frm_PlanillaPago.Show
    Frm_PlanillaPago.Caption = "Informe de Libro de Pensiones Interno de la Compañia."
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuLiqChequeCont_Click()
    Screen.MousePointer = vbHourglass
    'Generación de Archivo Cheque Contable
    vgNomInfSeleccionado = "InfGeneraArchChequeCon"
    Frm_CargaArchivo.Show
    Frm_CargaArchivo.Caption = "Generación de Archivo de Cheque Contable."
    Frm_CargaArchivo.Lbl_FecProceso.Visible = True
    Frm_CargaArchivo.Txt_FecProceso.Visible = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuLiqconError_Click()
    Screen.MousePointer = vbHourglass
    'Informe de Pensiones con Error
    vgNombreInformeSeleccionado = "InfLiqconError"
    Frm_PlanillaPago.Show
    Frm_PlanillaPago.Caption = "Informe de Pensiondes con Error."
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuLiqDepBancario_Click()
    Screen.MousePointer = vbHourglass
    'Generación de Archivo de Deposito Bancario
    vgNomInfSeleccionado = "InfGeneraArchPagoBco"
    Frm_CargaArchivo.Show
    Frm_CargaArchivo.Caption = "Generación de Archivo de Deposito Bancario."
    Frm_CargaArchivo.Lbl_FecProceso.Visible = True
    Frm_CargaArchivo.Txt_FecProceso.Visible = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuLiqPenCarta_Click()
    Screen.MousePointer = vbHourglass
    'Informe de Liquidaciones de Pago
    vgNombreInformeSeleccionadoInd = "InfLiqPagoCarta"
    vgPalabra = ""
    Frm_PlanillaPensionado.Show
    Frm_PlanillaPensionado.Caption = "Informe de Liquidación de Pensiones Tipo Carta."
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuLiqPensiones_Click()
    Screen.MousePointer = vbHourglass
    'Informe de Liquidaciones de Pago
    vgNombreInformeSeleccionadoInd = "InfLiqPago"
    vgPalabra = ""
    Frm_PlanillaPensionado.Show
    Frm_PlanillaPensionado.Caption = "Informe de Liquidación de Pensiones."
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuLiqPrevired_Click()
    Screen.MousePointer = vbHourglass
    'Generación de Archivo de Pagos Previsionales Previred
    vgNomInfSeleccionado = "InfGeneraArchPrevired"
    Frm_CargaArchivo.Show
    Frm_CargaArchivo.Caption = "Generación de Archivo de Pagos Previsionales."
    Frm_CargaArchivo.Lbl_FecProceso.Visible = False
    Frm_CargaArchivo.Txt_FecProceso.Visible = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuLiqViasPago_Click()
    Screen.MousePointer = vbHourglass
    'Informe de Vías de Pago de Pensiones por Pensionado
    vgNombreInformeSeleccionado = "InfViaPago"
    Frm_PlanillaPago.Show
    Frm_PlanillaPago.Caption = "Informe de Vías de Pago de Pensiones por Pensionado."
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuLiqViasPagoRec_Click()
    Screen.MousePointer = vbHourglass
    'Informe de Vías de Pago de Pensiones por Receptor
    vgNombreInformeSeleccionado = "InfViaPagoRec"
    Frm_PlanillaPago.Show
    Frm_PlanillaPago.Caption = "Informe de Vías de Pago de Pensiones por Receptor."
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuMantenedorSuperv_Click()
    Screen.MousePointer = vbHourglass
    Frm_AntCertificadoSup.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub Mnumanual_Click()
    Screen.MousePointer = vbHourglass
    Frm_CCAFManual.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuModCargaCCAF_Click()
    Screen.MousePointer = vbHourglass
    Frm_CCAFMantAux.Show
    Screen.MousePointer = vbDefault
End Sub

'Private Sub mnuMsgAutomáticos_Click()
'    Screen.MousePointer = vbHourglass
'    Frm_PensMensajesMasivos.Show
'    Screen.MousePointer = vbDefault
'End Sub

Private Sub mnuMsgAutomaticos_Click()
    Screen.MousePointer = vbHourglass
    Frm_MenMasivo.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuMsgCarga_Click()
    Screen.MousePointer = vbHourglass
    Frm_MenImportar.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub MnuMsgIndividual_Click()
    Screen.MousePointer = vbHourglass
    Frm_MenIndividual.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuParametro_Click()
    Screen.MousePointer = vbHourglass
    Frm_PensParametros.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuParametroPriPagos_Click()
    Screen.MousePointer = vbHourglass
    Frm_PensParametrosPrimerPago.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuPorcPension_Click()
    Screen.MousePointer = vbHourglass
    Frm_PorcCastigo.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuPrimerDefinitivo_Click()
    Screen.MousePointer = vbHourglass
    If Me.Tag = "0" Then
            vgTipoPago = "P" 'Primero Pago
            Frm_PensPrimerosPagos.txt_TipoCalculo = "DEFINITIVO"
            Frm_PensPrimerosPagos.Tag = "D"
            Frm_PensPrimerosPagos.Show
            Frm_Menu.Tag = "2"
    Else
        If Me.Tag = "1" Then
            MsgBox "Debe Cerrar ventana de Primeros Pagos para poder Ejecutar esta Opción", vbCritical, Me.Caption
        Else
            MsgBox "Debe Cerrar ventana de Pagos en Régiman para poder Ejecutar esta Opción", vbCritical, Me.Caption
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuPrimerProvisorio_Click()
    Screen.MousePointer = vbHourglass
    If Me.Tag = "0" Then
            vgTipoPago = "P" 'Primero Pago
            Frm_PensPrimerosPagos.txt_TipoCalculo = "PROVISORIO"
            Frm_PensPrimerosPagos.Tag = "P"
            Frm_PensPrimerosPagos.Show
            Frm_Menu.Tag = "1"
    Else
        If Me.Tag = "2" Then
            MsgBox "Debe Cerrar ventana de Primeros Pagos para poder Ejecutar esta Opción", vbCritical, Me.Caption
        Else
            MsgBox "Debe Cerrar ventana de Pagos de Cartera para poder Ejecutar esta Opción", vbCritical, Me.Caption
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuPSInfArchivoCotFonasa_Click()
    Screen.MousePointer = vbHourglass
    'Generación de Archivo de Cotizaciones Fonasa
    vgNomInfSeleccionado = "InfGeneraArchCotFonasa"
    Frm_CargaArchivo.Show
    Frm_CargaArchivo.Caption = "Generación de Archivo de Cotizaciones EsSalud."
    Frm_CargaArchivo.Lbl_FecProceso.Visible = False
    Frm_CargaArchivo.Txt_FecProceso.Visible = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuPSInfArchivoPresMed_Click()
    Screen.MousePointer = vbHourglass
    'Generación de Archivo de Prestamos Médicos
    vgNomInfSeleccionado = "InfGeneraArchPresMed"
    Frm_CargaArchivo.Show
    Frm_CargaArchivo.Caption = "Generación de Archivo de Prestamos Médicos."
    Frm_CargaArchivo.Lbl_FecProceso.Visible = False
    Frm_CargaArchivo.Txt_FecProceso.Visible = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuPSInfCotizaciones_Click()
    Screen.MousePointer = vbHourglass
    'Planilla de Cotizaciones de Salud
    vgNombreInformeSeleccionado = "InfPlaCotSalud"
    Frm_PlanillaPago.Show
    Frm_PlanillaPago.Caption = "Planilla de Cotizaciones de Salud."
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuPSInfPagoSalud_Click()
    Screen.MousePointer = vbHourglass
    'Resumen de Planilla de Pagos de Cotizaciones
    vgNombreInformeSeleccionado = "InfPagSalPenRV"
    Frm_PlanillaPago.Show
    Frm_PlanillaPago.Caption = "Resumen de Planilla de Pago de Cotizaciones."
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuPSInfPrestamos_Click()
    Screen.MousePointer = vbHourglass
    'Planilla de Pago de Prestamos Medicos Fonasa
    vgNombreInformeSeleccionado = "InfPreMedFon"
    Frm_PlanillaPago.Show
    Frm_PlanillaPago.Caption = "Planilla de Pago de Préstamos Médicos Fonasa."
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRegimenDefinitivo_Click()
    Screen.MousePointer = vbHourglass
    If Me.Tag = "0" Then
        If flObtieneDatosRegimen Then
            vgTipoPago = "R" 'Pago en Regimen
            Frm_PensPagosRegimen.Tag = "D"
            Frm_PensPagosRegimen.txt_TipoCalculo = "DEFINITIVO"
            Frm_PensPagosRegimen.Show
            Frm_Menu.Tag = "4"
        End If
    Else
        If Me.Tag = "1" Or Me.Tag = "2" Then
            MsgBox "Debe Cerrar ventana de Primeras Pensiones para poder Ejecutar esta Opción", vbCritical, Me.Caption
        Else
            MsgBox "Debe Cerrar ventana de Pagos Recurrentes para poder Ejecutar esta Opción", vbCritical, Me.Caption
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub
Private Sub mnuRegimenProvisorio_Click()
    Screen.MousePointer = vbHourglass
    If Me.Tag = "0" Then
        If flObtieneDatosRegimen Then
            vgTipoPago = "R" 'Pago en Regimen
            Frm_PensPagosRegimen.Show
            Frm_PensPagosRegimen.txt_TipoCalculo = "PROVISORIO"
            Frm_PensPagosRegimen.Tag = "P"
            Frm_Menu.Tag = "3"
        End If
    Else
        If Me.Tag = "1" Or Me.Tag = "2" Then
            MsgBox "Debe Cerrar ventana de Primeros Pagos para poder Ejecutar esta Opción", vbCritical, Me.Caption
        Else
            MsgBox "Debe Cerrar ventana de Pagos Recurrentes para poder Ejecutar esta Opción", vbCritical, Me.Caption
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRegula_Click()
 frmMemos.Show
End Sub

Private Sub mnuReliquidacion_Click()
'    Screen.MousePointer = vbHourglass
'    Frm_AFReliquidacion.Show
'    Screen.MousePointer = vbDefault
        frmReliqMasiva.Show
End Sub

Private Sub mnuRegistroPagosATerConsulta_Click()
    Screen.MousePointer = vbHourglass
    Frm_PensRegistroPagosCon.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRegistroPagosATerCuoMor_Click()
    Screen.MousePointer = vbHourglass
    Frm_PensRegistroPagos.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRegistroPagosATerPerGar_Click()
    Screen.MousePointer = vbHourglass
    Frm_PensRegistroPagosGar.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnureporlog_Click()
  frm_reportelog.Show
End Sub

Private Sub mnuReporteControlAF_Click()
    Screen.MousePointer = 11
        
    vgTipoInforme = "AF"
    vgTituloInfControl = "Informe Control Asignación Familiar"
    
    'Descarga y Carga el formulario para refrescar los datos
    Unload Frm_InformeControl
    Frm_InformeControl.Show
     
    Screen.MousePointer = 0
End Sub

Private Sub mnuRetInforme_Click()
    Screen.MousePointer = 11
        
    vgTipoInforme = "RJ"
    vgTituloInfControl = "Informe Control Retención Judicial"
    
    'Descarga y Carga el formulario para refrescar los datos
    Unload Frm_InformeControl
    Frm_InformeControl.Show
     
    Screen.MousePointer = 0
End Sub

Private Sub mnuRetIngreso_Click()
    Screen.MousePointer = vbHourglass
    Frm_RetJudicial.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuSalir_Click()
Dim x%
    
    x% = MsgBox("¿ Está Seguro que desea Salir del Sistema ?", 32 + 4, "Salir")
    If x% = 6 Then
    ''CORPTEC
     Call fgLogOut_Pen
        End
    Else
        Cancel = 1
    End If
End Sub

 ''CORPTEC
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Or UnloadMode = 1 Then
        Dim x%
        x% = MsgBox("¿ Está Seguro que desea Salir del Sistema?", 32 + 4, "Salir")
        If x% = 6 Then
            
            Call fgLogOut_Pen
            End
        Else
            Cancel = True
        End If
    End If
End Sub

Private Sub mnuSVSInfArchivoCir1815_Click()
    Screen.MousePointer = vbHourglass
    Frm_PlanillaSVSVta.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuSVSInfPromedio_Click()
    Screen.MousePointer = vbHourglass
    'Informe Definitivo de Anexo SVS
    vgPlanillaPeriodoSeleccionada = "InfAnexoSVS"
    Frm_PlanillaSVS.Show
    Frm_PlanillaSVS.Caption = "Anexo SBS."
    Screen.MousePointer = vbDefault
End Sub

Private Sub Mnututores_Click()
    Screen.MousePointer = vbHourglass
    Frm_AntTutores.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub Mnuvaloresaf_Click()
    Screen.MousePointer = vbHourglass
    Frm_AFParametros.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuCCAportes_Click()
   Screen.MousePointer = 11
        
    vgPagoCCaf = "25"
    vgNombreInformeSeleccionado = "infPagoCCaf"
    'Descarga y Carga el formulario para refrescar los datos
    Unload Frm_PlanillaPago
    Frm_PlanillaPago.Caption = "Informe Detalle Descto. CCAF Aporte"
    Frm_PlanillaPago.Show
     
    Screen.MousePointer = 0
End Sub

Private Sub mnuCCCreditos_Click()
   Screen.MousePointer = 11
        
    vgPagoCCaf = "26"
    vgNombreInformeSeleccionado = "infPagoCCaf"
    'Descarga y Carga el formulario para refrescar los datos
    Unload Frm_PlanillaPago
    Frm_PlanillaPago.Caption = "Informe Detalle Descto. CCAF Créditos"
    Frm_PlanillaPago.Show
     
    Screen.MousePointer = 0
End Sub
Private Sub mnuCCOtrasPrestaciones_Click()
   Screen.MousePointer = 11
        
    vgPagoCCaf = "27"
    vgNombreInformeSeleccionado = "infPagoCCaf"
    'Descarga y Carga el formulario para refrescar los datos
    Unload Frm_PlanillaPago
    Frm_PlanillaPago.Caption = "Informe Detalle Descto. CCAF Otras Prestaciones"
    Frm_PlanillaPago.Show
     
    Screen.MousePointer = 0
End Sub


Function flObtieneDatosRegimen() As Boolean
'Obtiene Datos desplegados en la Pantalla
Dim vlSql As String
Dim vlUF As Double
Dim vlFecPago As String
Dim vlUltDiaMesFecPago As Date

flObtieneDatosRegimen = False
vlUF = 0
vlSql = "SELECT * FROM PP_TMAE_PROPAGOPEN A"
vlSql = vlSql & " WHERE A.COD_ESTADOREG IN ('A','P')"
vlSql = vlSql & " AND A.NUM_PERPAGO ="
vlSql = vlSql & " (SELECT MIN(NUM_PERPAGO)"
vlSql = vlSql & " FROM PP_TMAE_PROPAGOPEN WHERE COD_ESTADOREG IN ('A','P'))"
Set vgRs = vgConexionBD.Execute(vlSql)
If Not vgRs.EOF Then
    vlFecPago = DateSerial(Mid(vgRs!fec_calpagoreg, 1, 4), Mid(vgRs!fec_calpagoreg, 5, 2), Mid(vgRs!fec_calpagoreg, 7, 2))
    Frm_PensPagosRegimen.Txt_Periodo = Mid(vgRs!Num_PerPago, 5, 2) & "/" & Mid(vgRs!Num_PerPago, 1, 4)
    Frm_PensPagosRegimen.Txt_FecPago = DateSerial(Mid(vgRs!Fec_PagoReg, 1, 4), Mid(vgRs!Fec_PagoReg, 5, 2), Mid(vgRs!Fec_PagoReg, 7, 2))
    Frm_PensPagosRegimen.Txt_FecCalculo = vlFecPago
    'If Not fgObtieneConversion(vgRs!fec_calpagoreg, "UF", vlUF) Then
    '    MsgBox "Debe ingresar Valor UF a la Fecha de Calculo : " & vlFecPago, vbCritical, Me.Caption
    '    Exit Function
    'End If
    'Frm_PensPagosRegimen.txt_FecProxPago = DateSerial(Mid(vgRs!Fec_PagoProxReg, 1, 4), Mid(vgRs!Fec_PagoProxReg, 5, 2), Mid(vgRs!Fec_PagoProxReg, 7, 2))
    'Frm_PensPagosRegimen.txt_UF = Format(vlUF, "###,##0.00")
    
    vgFecIniPag = DateSerial(Mid(vgRs!Num_PerPago, 1, 4), Mid(vgRs!Num_PerPago, 5, 2), 1)
    vgFecTerPag = DateAdd("d", -1, DateAdd("m", 1, vgFecIniPag))
    vgPerPago = vgRs!Num_PerPago
    vgFecPago = vgRs!fec_calpagoreg

    'Valor UF al último día del mes
    'vlUltDiaMesFecPago = DateSerial(Mid(vgFecPago, 1, 4), Mid(vgFecPago, 5, 2) + 1, 0)
    'If Not fgObtieneConversion(Format(vlUltDiaMesFecPago, "yyyymmdd"), "UF", vlUF) Then
    '    MsgBox "Debe ingresar Valor UF al Último Día del Mes : " & vlUltDiaMesFecPago, vbCritical, Me.Caption
    '    Exit Function
    'End If
    'stDatGenerales.Val_UFUltDiaMes = vlUF
Else
    MsgBox "Debe Definir el Periodo que se va a Calcular", vbCritical, Me.Caption
    Exit Function
End If
flObtieneDatosRegimen = True
End Function





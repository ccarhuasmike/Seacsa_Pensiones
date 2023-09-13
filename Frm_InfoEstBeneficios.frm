VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_InfoEstBeneficios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Información Estadistica de Beneficios"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleMode       =   0  'User
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   8640
      TabIndex        =   14
      Top             =   3000
      Width           =   1335
   End
   Begin VB.OptionButton optSobrevivencia 
      Caption         =   "Pensiones de Sobrevivencia"
      Height          =   375
      Left            =   5640
      TabIndex        =   13
      Top             =   2400
      Width           =   2895
   End
   Begin VB.OptionButton optJubilacion 
      Caption         =   "Pensión de Jubilación"
      Height          =   375
      Left            =   3600
      TabIndex        =   12
      Top             =   2400
      Width           =   2535
   End
   Begin VB.OptionButton OptModalidades 
      Caption         =   "Infomación de Pensiones por Modalidades"
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Value           =   -1  'True
      Width           =   3375
   End
   Begin VB.CommandButton cmd_exportarPM 
      Caption         =   "Exportar txt"
      Height          =   495
      Left            =   8640
      TabIndex        =   8
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Selecciona AFP"
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   9975
      Begin VB.CheckBox chkAnexo 
         Caption         =   "Habitat"
         Height          =   255
         Index           =   3
         Left            =   6000
         TabIndex        =   7
         Tag             =   "19"
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox chkAnexo 
         Caption         =   "Prima"
         Height          =   255
         Index           =   2
         Left            =   4320
         TabIndex        =   6
         Tag             =   "16"
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox chkAnexo 
         Caption         =   "Profuturo"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   5
         Tag             =   "13"
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox chkAnexo 
         Caption         =   "Integra"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   4
         Tag             =   "10"
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtro"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10095
      Begin VB.TextBox txt_TC 
         Height          =   285
         Left            =   8880
         TabIndex        =   16
         Top             =   360
         Width           =   855
      End
      Begin VB.ComboBox CmbMesExtrae 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "TC"
         Height          =   375
         Left            =   8520
         TabIndex        =   15
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Mes Extraer"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
   End
   Begin MSComctlLib.ProgressBar Barra 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblMensaje 
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3480
      Width           =   6855
   End
End
Attribute VB_Name = "Frm_InfoEstBeneficios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const IdFormato As String = "0712"
Const IdSBSEnpresaVigilada = "00209"
Const IdExpresionMontos = "012"
Const IdDatoControl = "0"
Const DelimitadorFinRegistro = "|"
Dim FechaReporte As String

Dim vlSqlXMODCount As String
Dim vlSqlXMODQuery As String
Dim vlSqlXMOD As String

Dim vlSqlXPENCount As String
Dim vlSqlXPENQuery As String
Dim vlSqlXPEN As String

Dim vlSqlXSobCount As String
Dim vlSqlXSobQuery As String
Dim vlSqlXSov As String


Dim IdAnexo As String
Dim Cabecera As String
Dim CountRegistros As String


Private Sub InicializaCadenasSobrevivencia()


vlSqlXSobCount = ""
vlSqlXSobQuery = ""
vlSqlXSov = ""

vlSqlXSobQuery = vlSqlXSobQuery & "SELECT "
vlSqlXSobQuery = vlSqlXSobQuery & "LPAD(NVL(A.NUM_FILA,0),4,'0')||" & Chr(13)
vlSqlXSobQuery = vlSqlXSobQuery & "LPAD(NVL(B.EDAD_18,0),15,'0')||" & Chr(13)
vlSqlXSobQuery = vlSqlXSobQuery & "LPAD(NVL(B.EDAD_18_25,0),15,'0')||" & Chr(13)
vlSqlXSobQuery = vlSqlXSobQuery & "LPAD(NVL(B.EDAD_26_35,0),15,'0')||" & Chr(13)
vlSqlXSobQuery = vlSqlXSobQuery & "LPAD(NVL(B.EDAD_36_40,0),15,'0')||" & Chr(13)
vlSqlXSobQuery = vlSqlXSobQuery & "LPAD(NVL(B.EDAD_41_45,0),15,'0')||" & Chr(13)
vlSqlXSobQuery = vlSqlXSobQuery & "LPAD(NVL(B.EDAD_46_50,0),15,'0')||" & Chr(13)
vlSqlXSobQuery = vlSqlXSobQuery & "LPAD(NVL(B.EDAD_51_55,0),15,'0')||" & Chr(13)
vlSqlXSobQuery = vlSqlXSobQuery & "LPAD(NVL(B.EDAD_56_60,0),15,'0')||" & Chr(13)
vlSqlXSobQuery = vlSqlXSobQuery & "LPAD(NVL(B.EDAD_60,0),15,'0')||" & Chr(13)
vlSqlXSobQuery = vlSqlXSobQuery & "LPAD('0',15,'0')||" & Chr(13)
vlSqlXSobQuery = vlSqlXSobQuery & "LPAD(NVL(C.NUM_SEXO,0),15,'0')||" & Chr(13)
vlSqlXSobQuery = vlSqlXSobQuery & "LPAD(NVL(D.NUM_PEN,0),15,'0')||" & Chr(13)
vlSqlXSobQuery = vlSqlXSobQuery & "LPAD((abs(NVL(D.MTO_HABER,0))-floor(NVL(D.MTO_HABER,0)))*100,2,'0')||" & Chr(13)
vlSqlXSobQuery = vlSqlXSobQuery & "LPAD(NVL(E.NUM_PEN,0),15,'0')||" & Chr(13)
vlSqlXSobQuery = vlSqlXSobQuery & "LPAD((abs(NVL(E.MTO_HABER,0))-floor(NVL(E.MTO_HABER,0)))*100,2,'0') as fila" & Chr(13)


vlSqlXSobCount = vlSqlXSobCount & "SELECT COUNT(1) AS cantidad " & Chr(13)


vlSqlXSov = vlSqlXSov & "FROM PP_TTMP_TAB_ANEX_12151821 A" & Chr(13)
vlSqlXSov = vlSqlXSov & "Left Join" & Chr(13)
vlSqlXSov = vlSqlXSov & "(" & Chr(13)
vlSqlXSov = vlSqlXSov & "SELECT CASE WHEN B.COD_PAR IN (10,11,20,21) THEN 'C' ELSE" & Chr(13)
vlSqlXSov = vlSqlXSov & "         CASE WHEN B.COD_PAR IN (30) THEN 'H' ELSE 'P' END END COD_PAR," & Chr(13)
vlSqlXSov = vlSqlXSov & "                 SUM(CASE WHEN TRUNC(months_between(SYSDATE, TO_DATE(B.FEC_NACBEN, 'YYYYMMDD')) /12,0) < 18 THEN 1 ELSE 0 END) EDAD_18 ," & Chr(13)
vlSqlXSov = vlSqlXSov & "         SUM(CASE WHEN TRUNC(months_between(SYSDATE, TO_DATE(B.FEC_NACBEN, 'YYYYMMDD')) /12,0) BETWEEN 18 AND 25 THEN 1 ELSE 0 END) EDAD_18_25 ," & Chr(13)
vlSqlXSov = vlSqlXSov & "         SUM(CASE WHEN TRUNC(months_between(SYSDATE, TO_DATE(B.FEC_NACBEN, 'YYYYMMDD')) /12,0) BETWEEN 26 AND 35 THEN 1 ELSE 0 END) EDAD_26_35 ," & Chr(13)
vlSqlXSov = vlSqlXSov & "         SUM(CASE WHEN TRUNC(months_between(SYSDATE, TO_DATE(B.FEC_NACBEN, 'YYYYMMDD')) /12,0) BETWEEN 36 AND 40 THEN 1 ELSE 0 END) EDAD_36_40," & Chr(13)
vlSqlXSov = vlSqlXSov & "         SUM(CASE WHEN TRUNC(months_between(SYSDATE, TO_DATE(B.FEC_NACBEN, 'YYYYMMDD')) /12,0) BETWEEN 41 AND 45 THEN 1 ELSE 0 END) EDAD_41_45 ," & Chr(13)
vlSqlXSov = vlSqlXSov & "         SUM(CASE WHEN TRUNC(months_between(SYSDATE, TO_DATE(B.FEC_NACBEN, 'YYYYMMDD')) /12,0) BETWEEN 46 AND 50 THEN 1 ELSE 0 END) EDAD_46_50," & Chr(13)
vlSqlXSov = vlSqlXSov & "         SUM(CASE WHEN TRUNC(months_between(SYSDATE, TO_DATE(B.FEC_NACBEN, 'YYYYMMDD')) /12,0) BETWEEN 51 AND 55 THEN 1 ELSE 0 END) EDAD_51_55," & Chr(13)
vlSqlXSov = vlSqlXSov & "         SUM(CASE WHEN TRUNC(months_between(SYSDATE, TO_DATE(B.FEC_NACBEN, 'YYYYMMDD')) /12,0) BETWEEN 56 AND 60 THEN 1 ELSE 0 END) EDAD_56_60," & Chr(13)
vlSqlXSov = vlSqlXSov & "         SUM(CASE WHEN TRUNC(months_between(SYSDATE, TO_DATE(B.FEC_NACBEN, 'YYYYMMDD')) /12,0) BETWEEN 61 AND 120 THEN 1 ELSE 0 END) EDAD_60," & Chr(13)
vlSqlXSov = vlSqlXSov & "         COUNT(B.COD_SEXO) NUM_SEXO" & Chr(13)
vlSqlXSov = vlSqlXSov & "         FROM PD_TMAE_POLBEN B" & Chr(13)
vlSqlXSov = vlSqlXSov & "         JOIN PD_TMAE_POLIZA P ON B.NUM_POLIZA=P.NUM_POLIZA AND B.NUM_ENDOSO=P.NUM_ENDOSO" & Chr(13)
vlSqlXSov = vlSqlXSov & "         JOIN PD_TMAE_POLPRIREC E ON P.NUM_POLIZA=E.NUM_POLIZA" & Chr(13)
vlSqlXSov = vlSqlXSov & "         Where B.Num_Endoso = 1" & Chr(13)
vlSqlXSov = vlSqlXSov & "         AND NUM_ORDEN<>1" & Chr(13)
vlSqlXSov = vlSqlXSov & "         AND COD_TIPPENSION IN ('08','09','10','11','12')" & Chr(13)
vlSqlXSov = vlSqlXSov & "         AND P.COD_AFP IN ([CodAFP])" & Chr(13)
vlSqlXSov = vlSqlXSov & "         AND E.FEC_TRASPASO BETWEEN [InicioPer] AND [InicioPer] " & Chr(13)
vlSqlXSov = vlSqlXSov & "         group by CASE WHEN B.COD_PAR IN (10,11,20,21) THEN 'C' ELSE" & Chr(13)
vlSqlXSov = vlSqlXSov & "                CASE WHEN B.COD_PAR IN (30) THEN 'H' ELSE 'P' END END" & Chr(13)
vlSqlXSov = vlSqlXSov & "         ) B ON A.COD_PAR=B.COD_PAR" & Chr(13)
vlSqlXSov = vlSqlXSov & "         Left Join" & Chr(13)
vlSqlXSov = vlSqlXSov & "         (" & Chr(13)
vlSqlXSov = vlSqlXSov & "                 SELECT CASE WHEN B.COD_PAR IN (10,11,20,21) THEN 'C' ELSE CASE WHEN B.COD_PAR IN (30) THEN 'H' ELSE 'P' END END COD_PAR," & Chr(13)
vlSqlXSov = vlSqlXSov & "                COUNT(B.COD_SEXO) NUM_SEXO" & Chr(13)
vlSqlXSov = vlSqlXSov & "                 FROM PD_TMAE_POLIZA A" & Chr(13)
vlSqlXSov = vlSqlXSov & "                 JOIN PD_TMAE_POLBEN B ON A.NUM_POLIZA=B.NUM_POLIZA AND A.NUM_ENDOSO=B.NUM_ENDOSO" & Chr(13)
vlSqlXSov = vlSqlXSov & "                 JOIN MA_TPAR_TABCOD T ON A.COD_VEJEZ=T.COD_ELEMENTO AND T.COD_TABLA='TV'" & Chr(13)
vlSqlXSov = vlSqlXSov & "                 JOIN PD_TMAE_POLPRIREC E ON A.NUM_POLIZA=E.NUM_POLIZA" & Chr(13)
vlSqlXSov = vlSqlXSov & "                 Where a.Num_Endoso = 1" & Chr(13)
vlSqlXSov = vlSqlXSov & "                 AND B.NUM_ORDEN<>1" & Chr(13)
vlSqlXSov = vlSqlXSov & "                 AND A.COD_TIPPENSION IN ('08','09','10','11','12')" & Chr(13)
vlSqlXSov = vlSqlXSov & "                 AND A.COD_AFP IN ([CodAFP]) " & Chr(13)
vlSqlXSov = vlSqlXSov & "                 AND E.FEC_TRASPASO BETWEEN [InicioPer] AND [FinPer] " & Chr(13)
vlSqlXSov = vlSqlXSov & "                 group by CASE WHEN B.COD_PAR IN (10,11,20,21) THEN 'C' ELSE CASE WHEN B.COD_PAR IN (30) THEN 'H' ELSE 'P' END END" & Chr(13)
vlSqlXSov = vlSqlXSov & "         ) C ON C.COD_PAR=A.COD_PAR" & Chr(13)
vlSqlXSov = vlSqlXSov & "         Left Join" & Chr(13)
vlSqlXSov = vlSqlXSov & "         (" & Chr(13)
vlSqlXSov = vlSqlXSov & "                SELECT CASE WHEN B.COD_PAR IN (10,11,20,21) THEN 'C' ELSE CASE WHEN B.COD_PAR IN (30) THEN 'H' ELSE 'P' END END COD_PAR, COUNT(*) NUM_PEN, SUM(MTO_HABER) MTO_HABER" & Chr(13)
vlSqlXSov = vlSqlXSov & "                 FROM PP_TMAE_LIQPAGOPENDEF L" & Chr(13)
vlSqlXSov = vlSqlXSov & "                 JOIN PD_TMAE_POLBEN B ON L.NUM_POLIZA=B.NUM_POLIZA AND L.NUM_ORDEN=B.NUM_ORDEN" & Chr(13)
vlSqlXSov = vlSqlXSov & "                JOIN PD_TMAE_POLIZA P ON B.NUM_POLIZA=P.NUM_POLIZA AND B.NUM_ENDOSO=P.NUM_ENDOSO" & Chr(13)
vlSqlXSov = vlSqlXSov & "                 WHERE P.COD_TIPPENSION IN ('08','09','10','11','12')" & Chr(13)
vlSqlXSov = vlSqlXSov & "                 AND L.COD_TIPOPAGO='P'" & Chr(13)
vlSqlXSov = vlSqlXSov & "                AND B.NUM_ORDEN<>1" & Chr(13)
vlSqlXSov = vlSqlXSov & "                AND B.NUM_ENDOSO=1" & Chr(13)
vlSqlXSov = vlSqlXSov & "                AND P.COD_AFP IN ([CodAFP]) " & Chr(13)
vlSqlXSov = vlSqlXSov & "                AND L.NUM_PERPAGO=[FinPer]" & Chr(13)
vlSqlXSov = vlSqlXSov & "                group by CASE WHEN B.COD_PAR IN (10,11,20,21) THEN 'C' ELSE CASE WHEN B.COD_PAR IN (30) THEN 'H' ELSE 'P' END END" & Chr(13)
vlSqlXSov = vlSqlXSov & "                ) D ON D.COD_PAR=A.COD_PAR" & Chr(13)
vlSqlXSov = vlSqlXSov & "                Left Join" & Chr(13)
vlSqlXSov = vlSqlXSov & "                (" & Chr(13)
vlSqlXSov = vlSqlXSov & "                SELECT CASE WHEN B.COD_PAR IN (10,11,20,21) THEN 'C' ELSE CASE WHEN B.COD_PAR IN (30) THEN 'H' ELSE 'P' END END COD_PAR, COUNT(*) NUM_PEN, SUM(MTO_HABER) MTO_HABER" & Chr(13)
vlSqlXSov = vlSqlXSov & "                FROM PP_TMAE_LIQPAGOPENDEF L" & Chr(13)
vlSqlXSov = vlSqlXSov & "                JOIN PD_TMAE_POLBEN B ON L.NUM_POLIZA=B.NUM_POLIZA AND L.NUM_ORDEN=B.NUM_ORDEN" & Chr(13)
vlSqlXSov = vlSqlXSov & "                JOIN PD_TMAE_POLIZA P ON B.NUM_POLIZA=P.NUM_POLIZA AND B.NUM_ENDOSO=P.NUM_ENDOSO" & Chr(13)
vlSqlXSov = vlSqlXSov & "                WHERE P.COD_TIPPENSION IN ('08','09','10','11','12')" & Chr(13)
vlSqlXSov = vlSqlXSov & "                AND L.COD_TIPOPAGO='R'" & Chr(13)
vlSqlXSov = vlSqlXSov & "                AND B.NUM_ORDEN<>1" & Chr(13)
vlSqlXSov = vlSqlXSov & "                AND B.NUM_ENDOSO=1" & Chr(13)
vlSqlXSov = vlSqlXSov & "                AND P.COD_AFP IN ([CodAFP]) " & Chr(13)
vlSqlXSov = vlSqlXSov & "                AND L.NUM_PERPAGO=[PERIODO]" & Chr(13)
vlSqlXSov = vlSqlXSov & "                group by CASE WHEN B.COD_PAR IN (10,11,20,21) THEN 'C' ELSE CASE WHEN B.COD_PAR IN (30) THEN 'H' ELSE 'P' END END" & Chr(13)
vlSqlXSov = vlSqlXSov & "                ) E ON E.COD_PAR=A.COD_PAR" & Chr(13)
vlSqlXSov = vlSqlXSov & " ORDER BY 1"


End Sub

Private Sub InicializaCadenasJubilacion()
'*******Pensiones por jubilado

vlSqlXPEN = ""
vlSqlXPENQuery = ""
vlSqlXPENCount = ""


vlSqlXPENQuery = vlSqlXPENQuery & "Select lpad(N.NUM_FILA,4,'0')||" & Chr(13)
vlSqlXPENQuery = vlSqlXPENQuery & "lpad(NVL(EDAD_55,0),15,'0')||" & Chr(13)
vlSqlXPENQuery = vlSqlXPENQuery & "lpad(NVL(EDAD_55_60,0),15,'0')||" & Chr(13)
vlSqlXPENQuery = vlSqlXPENQuery & "lpad(NVL(EDAD_61_65,0),15,'0')||" & Chr(13)
vlSqlXPENQuery = vlSqlXPENQuery & "lpad(NVL(EDAD_66_70,0),15,'0')||" & Chr(13)
vlSqlXPENQuery = vlSqlXPENQuery & "lpad(NVL(EDAD_71_75,0),15,'0')||" & Chr(13)
vlSqlXPENQuery = vlSqlXPENQuery & "lpad(NVL(EDAD_75,0),15,'0')||" & Chr(13)
vlSqlXPENQuery = vlSqlXPENQuery & "lpad(NVL(A.NUM_SEXO,0),15,'0')||" & Chr(13)
vlSqlXPENQuery = vlSqlXPENQuery & "lpad(NVL(B.NUM_PEN,0),15,'0')||" & Chr(13)
vlSqlXPENQuery = vlSqlXPENQuery & "lpad(TRUNC(NVL(B.MTO_PEN,0)),13,'0')||lpad((abs(NVL(B.MTO_PEN,0))-floor(NVL(B.MTO_PEN,0)))*100,2,'0')||" & Chr(13)
vlSqlXPENQuery = vlSqlXPENQuery & "lpad(NVL(C.NUM_PEN,0),15,'0')||lpad(trunc(NVL(C.MTO_PEN,0)),13,'0')||lpad((abs(NVL(C.MTO_PEN,0))-floor(NVL(C.MTO_PEN,0)))*100,2,'0')||" & Chr(13)
vlSqlXPENQuery = vlSqlXPENQuery & "lpad('0',15,'0') as fila " & Chr(13)

vlSqlXPENCount = vlSqlXPENCount & "SELECT COUNT(1) AS cantidad "


vlSqlXPEN = vlSqlXPEN & " FROM PP_TTMP_TAB_ANEX_11141720 N  " & Chr(13)
vlSqlXPEN = vlSqlXPEN & " Left Join " & Chr(13)
vlSqlXPEN = vlSqlXPEN & " (SELECT A.COD_TIPPENSION, A.COD_VEJEZ, T.GLS_ELEMENTO TIPO, B.COD_SEXO , COUNT(B.COD_SEXO) NUM_SEXO  " & Chr(13)
vlSqlXPEN = vlSqlXPEN & " FROM PD_TMAE_POLIZA A " & Chr(13)
vlSqlXPEN = vlSqlXPEN & " JOIN PD_TMAE_POLBEN B ON A.NUM_POLIZA=B.NUM_POLIZA AND A.NUM_ENDOSO=B.NUM_ENDOSO " & Chr(13)
vlSqlXPEN = vlSqlXPEN & " JOIN MA_TPAR_TABCOD T ON A.COD_VEJEZ=T.COD_ELEMENTO AND T.COD_TABLA='TV' " & Chr(13)
vlSqlXPEN = vlSqlXPEN & " JOIN PD_TMAE_POLPRIREC E ON A.NUM_POLIZA=E.NUM_POLIZA " & Chr(13)
vlSqlXPEN = vlSqlXPEN & " WHERE A.NUM_ENDOSO=(SELECT MAX(NUM_ENDOSO) FROM PD_TMAE_POLIZA WHERE NUM_POLIZA=A.NUM_POLIZA) " & Chr(13)
vlSqlXPEN = vlSqlXPEN & " AND A.COD_TIPPENSION IN ('04','05') " & Chr(13)
vlSqlXPEN = vlSqlXPEN & " AND B.NUM_ORDEN=1 " & Chr(13)
vlSqlXPEN = vlSqlXPEN & " AND E.FEC_TRASPASO BETWEEN [InicioPer] AND [FinPer] " & Chr(13)
vlSqlXPEN = vlSqlXPEN & " AND A.COD_AFP IN ([CodAFP]) " & Chr(13)
vlSqlXPEN = vlSqlXPEN & " group by A.COD_TIPPENSION, A.COD_VEJEZ, T.GLS_ELEMENTO, B.COD_SEXO " & Chr(13)
vlSqlXPEN = vlSqlXPEN & " ) A ON N.COD_TIPPENSION=A.COD_TIPPENSION AND N.COD_VEJEZ=A.COD_VEJEZ AND N.COD_SEXO=A.COD_SEXO " & Chr(13)
vlSqlXPEN = vlSqlXPEN & " LEFT JOIN ( " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "            SELECT A.COD_TIPPENSION, A.COD_VEJEZ, A.COD_SEXO, COUNT(*) NUM_PEN, SUM(MTO_HABER) MTO_PEN FROM ( " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "             SELECT DISTINCT P.COD_TIPPENSION, P.COD_VEJEZ, B.COD_SEXO, L.NUM_POLIZA, SUM(MTO_HABER) MTO_HABER " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "             FROM PP_TMAE_LIQPAGOPENDEF L " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "             JOIN PD_TMAE_POLBEN B ON L.NUM_POLIZA=B.NUM_POLIZA AND L.NUM_ORDEN=B.NUM_ORDEN " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "             JOIN PD_TMAE_POLIZA P ON B.NUM_POLIZA=P.NUM_POLIZA AND B.NUM_ENDOSO=P.NUM_ENDOSO " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "             WHERE P.COD_TIPPENSION IN ('04','05') " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "             AND P.COD_AFP IN ([CodAFP]) " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "             AND L.COD_TIPOPAGO='P' " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "             AND B.NUM_ORDEN=1 " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "             AND B.NUM_ENDOSO=1 " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "             AND L.NUM_PERPAGO=[PERIODO] " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "             group by P.COD_TIPPENSION, P.COD_VEJEZ, B.COD_SEXO, L.NUM_POLIZA " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "           ) A group by A.COD_TIPPENSION, A.COD_VEJEZ, A.COD_SEXO " & Chr(13)
vlSqlXPEN = vlSqlXPEN & " ) B ON N.COD_TIPPENSION=B.COD_TIPPENSION AND N.COD_VEJEZ=B.COD_VEJEZ AND N.COD_SEXO=B.COD_SEXO " & Chr(13)
vlSqlXPEN = vlSqlXPEN & " LEFT JOIN ( " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "            SELECT A.COD_TIPPENSION, A.COD_VEJEZ, A.COD_SEXO, COUNT(*) NUM_PEN, SUM(MTO_HABER) MTO_PEN FROM ( " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "             SELECT DISTINCT P.COD_TIPPENSION, P.COD_VEJEZ, B.COD_SEXO, L.NUM_POLIZA, SUM(MTO_HABER) MTO_HABER " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "             FROM PP_TMAE_LIQPAGOPENDEF L " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "             JOIN PD_TMAE_POLBEN B ON L.NUM_POLIZA=B.NUM_POLIZA AND L.NUM_ORDEN=B.NUM_ORDEN " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "             JOIN PD_TMAE_POLIZA P ON B.NUM_POLIZA=P.NUM_POLIZA AND B.NUM_ENDOSO=P.NUM_ENDOSO " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "             WHERE P.COD_TIPPENSION IN ('04','05') " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "            AND P.COD_AFP IN ([CodAFP]) " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "             AND L.COD_TIPOPAGO='R' " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "             AND B.NUM_ORDEN=1 " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "             AND B.NUM_ENDOSO=1 " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "             AND L.NUM_PERPAGO=[PERIODO] " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "             group by P.COD_TIPPENSION, P.COD_VEJEZ, B.COD_SEXO, L.NUM_POLIZA " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "           ) A group by A.COD_TIPPENSION, A.COD_VEJEZ, A.COD_SEXO " & Chr(13)
vlSqlXPEN = vlSqlXPEN & " ) C ON N.COD_TIPPENSION=C.COD_TIPPENSION AND N.COD_VEJEZ=C.COD_VEJEZ AND N.COD_SEXO=C.COD_SEXO " & Chr(13)
vlSqlXPEN = vlSqlXPEN & " Left Join " & Chr(13)
vlSqlXPEN = vlSqlXPEN & " ("
vlSqlXPEN = vlSqlXPEN & "       SELECT P.COD_TIPPENSION, P.COD_VEJEZ, B.COD_SEXO, " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "       SUM(CASE WHEN months_between(SYSDATE, TO_DATE(B.FEC_NACBEN, 'YYYYMMDD')) /12 < 55 THEN 1 ELSE 0 END) EDAD_55 , " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "       SUM(CASE WHEN months_between(SYSDATE, TO_DATE(B.FEC_NACBEN, 'YYYYMMDD')) /12 BETWEEN 55 AND 60 THEN 1 ELSE 0 END) EDAD_55_60 , " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "       SUM(CASE WHEN months_between(SYSDATE, TO_DATE(B.FEC_NACBEN, 'YYYYMMDD')) /12 BETWEEN 61 AND 65 THEN 1 ELSE 0 END) EDAD_61_65 , " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "       SUM(CASE WHEN months_between(SYSDATE, TO_DATE(B.FEC_NACBEN, 'YYYYMMDD')) /12 BETWEEN 66 AND 70 THEN 1 ELSE 0 END) EDAD_66_70, " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "       SUM(CASE WHEN months_between(SYSDATE, TO_DATE(B.FEC_NACBEN, 'YYYYMMDD')) /12 BETWEEN 71 AND 75 THEN 1 ELSE 0 END) EDAD_71_75 , " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "       SUM(CASE WHEN months_between(SYSDATE, TO_DATE(B.FEC_NACBEN, 'YYYYMMDD')) /12 BETWEEN 76 AND 150 THEN 1 ELSE 0 END) EDAD_75 " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "       FROM PD_TMAE_POLBEN B " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "       JOIN PD_TMAE_POLIZA P ON B.NUM_POLIZA=P.NUM_POLIZA AND B.NUM_ENDOSO=P.NUM_ENDOSO " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "       JOIN PD_TMAE_POLPRIREC E ON P.NUM_POLIZA=E.NUM_POLIZA " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "       Where B.Num_Endoso = 1 " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "       AND NUM_ORDEN=1 " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "       AND P.COD_TIPPENSION IN ('04','05') " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "       AND P.COD_AFP IN ([CodAFP]) " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "       AND E.FEC_TRASPASO BETWEEN [InicioPer] AND [FinPer] " & Chr(13)
vlSqlXPEN = vlSqlXPEN & "       group by P.COD_TIPPENSION, P.COD_VEJEZ, B.COD_SEXO " & Chr(13)
vlSqlXPEN = vlSqlXPEN & " ) D ON N.COD_TIPPENSION=D.COD_TIPPENSION AND N.COD_VEJEZ=D.COD_VEJEZ AND N.COD_SEXO=D.COD_SEXO " & Chr(13)
vlSqlXPEN = vlSqlXPEN & " ORDER BY 1 " & Chr(13)



End Sub

Private Sub InicializaCadenasModalidades()

'Query 1 Por Modalidades
'****************************************************
    vlSqlXMODQuery = ""
    vlSqlXMODQuery = " SELECT lpad(L.NUM_FILA,4,'0')||"
    vlSqlXMODQuery = vlSqlXMODQuery & "lpad(MAX(NUM_POL_40),15,'0')||"
    vlSqlXMODQuery = vlSqlXMODQuery & "lpad(MAX(NUM_PEN_50),15,'0')||"
    vlSqlXMODQuery = vlSqlXMODQuery & "lpad(trunc(MAX(MTO_PEN_60)),13,'0')|| lpad((abs(MAX(MTO_PEN_60))-floor(MAX(MTO_PEN_60)))*100,2,'0')||"
    vlSqlXMODQuery = vlSqlXMODQuery & "lpad(MAX(NUM_PEN_70),15,'0')||"
    vlSqlXMODQuery = vlSqlXMODQuery & "lpad(trunc(MAX(MTO_PEN_80)),13,'0')|| lpad((abs(MAX(MTO_PEN_80))-floor(MAX(MTO_PEN_80)))*100,2,'0') as fila "
    
    vlSqlXMODCount = ""
    vlSqlXMODCount = "select count(SUM(NUM_POL_40)) as cantidad "
    
    vlSqlXMOD = ""
    vlSqlXMOD = vlSqlXMOD & " FROM PP_TTMP_TABLAS_ANEXOS L " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & " JOIN ( " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "           SELECT B.NUM_FILA, sum(NVL(NUM_NUEPOL,0)) NUM_POL_40, sum(NVL(L.NUM_PEN,0)) NUM_PEN_50, sum(NVL(L.MTO_PEN,0)) MTO_PEN_60, sum(NVL(LR.NUM_PEN,0)) NUM_PEN_70, sum(NVL(LR.MTO_PEN,0)) MTO_PEN_80 " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "           FROM PP_TTMP_TABLAS_ANEXOS B " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "           LEFT JOIN (" & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                 SELECT COD_TIPPENSION, COD_TIPREN, COD_MONEDA, COD_TIPREAJUSTE, COUNT(*) NUM_NUEPOL FROM ( " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                       SELECT DISTINCT B.COD_TIPPENSION, B.COD_TIPREN, B.COD_MONEDA, B.COD_TIPREAJUSTE, A.NUM_POLIZA " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                       FROM PP_TMAE_LIQPAGOPENDEF A " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                       JOIN PP_TMAE_POLIZA B ON A.NUM_POLIZA=B.NUM_POLIZA and b.num_endoso=1 " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                       WHERE b.COD_TIPPENSION IN ('04','05') " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                       AND COD_TIPOPAGO='P' " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                       AND COD_TIPRECEPTOR<>'R' " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                       AND A.fec_pago between [InicioPer] AND [FinPer] " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                       AND B.COD_AFP IN ([CodAFP]) " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                       AND B.COD_TIPREN<>2 " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                 ) A GROUP BY COD_TIPPENSION, COD_TIPREN, COD_MONEDA, COD_TIPREAJUSTE " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "           ) A ON A.COD_TIPPENSION=B.COD_TIPPENSION AND A.COD_TIPREN=B.COD_TIPREN AND A.COD_MONEDA=B.COD_MONEDA AND A.COD_TIPREAJUSTE=B.COD_TIPREAJUSTE " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "           LEFT JOIN ( " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                 SELECT B.COD_TIPPENSION, B.COD_TIPREN, B.COD_MONEDA, B.COD_TIPREAJUSTE, COUNT(*) NUM_PEN, SUM(MTO_HABER) MTO_PEN " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                 FROM PP_TMAE_LIQPAGOPENDEF A " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                 JOIN Pp_TMAE_POLIZA B ON A.NUM_POLIZA=B.NUM_POLIZA and b.num_endoso=1 " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                 WHERE b.COD_TIPPENSION IN ('04','05') " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                 AND COD_TIPOPAGO='P' " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                 AND COD_TIPRECEPTOR<>'R' " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                 AND A.fec_pago between [InicioPer] AND [FinPer] " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                 AND B.COD_AFP IN ([CodAFP]) " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                 GROUP BY B.COD_TIPPENSION, B.COD_TIPREN, B.COD_MONEDA, B.COD_TIPREAJUSTE " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "           ) L ON B.COD_TIPPENSION=L.COD_TIPPENSION AND B.COD_TIPREN=L.COD_TIPREN AND B.COD_MONEDA=L.COD_MONEDA AND B.COD_TIPREAJUSTE=L.COD_TIPREAJUSTE " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "           LEFT JOIN ( " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                 SELECT B.COD_TIPPENSION, B.COD_TIPREN, B.COD_MONEDA, B.COD_TIPREAJUSTE, COUNT(*) NUM_PEN, SUM(MTO_HABER) MTO_PEN " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                 FROM PP_TMAE_LIQPAGOPENDEF A " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                 JOIN Pp_TMAE_POLIZA B ON A.NUM_POLIZA=B.NUM_POLIZA " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                 WHERE b.COD_TIPPENSION IN ('04','05') " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                 AND COD_TIPOPAGO='R' " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                 AND COD_TIPRECEPTOR<>'R' " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                 AND A.fec_pago between [InicioPer] AND [FinPer] " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                 AND B.COD_AFP IN ([CodAFP]) " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                 AND B.NUM_ENDOSO=(SELECT MAX(NUM_ENDOSO) FROM PP_TMAE_POLIZA WHERE NUM_POLIZA=B.NUM_POLIZA) " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                 GROUP BY B.COD_TIPPENSION, B.COD_TIPREN, B.COD_MONEDA, B.COD_TIPREAJUSTE " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "          ) LR ON B.COD_TIPPENSION=LR.COD_TIPPENSION AND B.COD_TIPREN=LR.COD_TIPREN AND B.COD_MONEDA=LR.COD_MONEDA AND B.COD_TIPREAJUSTE=LR.COD_TIPREAJUSTE " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "          group by B.NUM_FILA " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "          Union All " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "          SELECT B.NUM_FILA, sum(NVL(NUM_NUEPOL,0)) NUM_POL_40, sum(NVL(L.NUM_PEN,0)) NUM_PEN_50, sum(NVL(L.MTO_PEN,0)) MTO_PEN_60, sum(NVL(LR.NUM_PEN,0)) NUM_PEN_70, sum(NVL(LR.MTO_PEN,0)) MTO_PEN_80 " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "          FROM PP_TTMP_TABLAS_ANEXOS B " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "          LEFT JOIN ( " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                SELECT COD_TIPPENSION, COD_TIPREN, COD_MONEDA, COD_TIPREAJUSTE, COUNT(*) NUM_NUEPOL FROM ( " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                      SELECT DISTINCT B.COD_TIPPENSION, B.COD_TIPREN, B.COD_MONEDA, B.COD_TIPREAJUSTE, A.NUM_POLIZA " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                      FROM PP_TMAE_LIQPAGOPENDEF A " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                      JOIN PP_TMAE_POLIZA B ON A.NUM_POLIZA=B.NUM_POLIZA and b.num_endoso=1 " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                      WHERE b.COD_TIPPENSION IN ('09','10') " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                      AND COD_TIPOPAGO='P' " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                      AND COD_TIPRECEPTOR<>'R' " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                      AND A.fec_pago between [InicioPer] AND [FinPer] " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                      AND B.COD_AFP IN ([CodAFP]) " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                      AND B.COD_TIPREN<>2 " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                ) A GROUP BY COD_TIPPENSION, COD_TIPREN, COD_MONEDA, COD_TIPREAJUSTE " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "          ) A ON A.COD_TIPPENSION=B.COD_TIPPENSION AND A.COD_TIPREN=B.COD_TIPREN AND A.COD_MONEDA=B.COD_MONEDA AND A.COD_TIPREAJUSTE=B.COD_TIPREAJUSTE " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "          LEFT JOIN ( " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                SELECT B.COD_TIPPENSION, B.COD_TIPREN, B.COD_MONEDA, B.COD_TIPREAJUSTE, COUNT(*) NUM_PEN, SUM(MTO_HABER) MTO_PEN " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                FROM PP_TMAE_LIQPAGOPENDEF A " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                JOIN Pp_TMAE_POLIZA B ON A.NUM_POLIZA=B.NUM_POLIZA and b.num_endoso=1 " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                WHERE b.COD_TIPPENSION IN ('09','10') " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                AND COD_TIPOPAGO='P' " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                AND COD_TIPRECEPTOR<>'R' " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                AND A.fec_pago between [InicioPer] AND [FinPer] " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                AND B.COD_AFP IN ([CodAFP]) " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                GROUP BY B.COD_TIPPENSION, B.COD_TIPREN, B.COD_MONEDA, B.COD_TIPREAJUSTE " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "          ) L ON B.COD_TIPPENSION=L.COD_TIPPENSION AND B.COD_TIPREN=L.COD_TIPREN AND B.COD_MONEDA=L.COD_MONEDA AND B.COD_TIPREAJUSTE=L.COD_TIPREAJUSTE " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "          LEFT JOIN ( " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                SELECT B.COD_TIPPENSION, B.COD_TIPREN, B.COD_MONEDA, B.COD_TIPREAJUSTE, COUNT(*) NUM_PEN, SUM(MTO_HABER) MTO_PEN " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                FROM PP_TMAE_LIQPAGOPENDEF A " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                JOIN Pp_TMAE_POLIZA B ON A.NUM_POLIZA=B.NUM_POLIZA " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                WHERE b.COD_TIPPENSION IN ('09','10') " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                AND COD_TIPOPAGO='R' " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                AND COD_TIPRECEPTOR<>'R' " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                AND A.fec_pago between [InicioPer] AND [FinPer] " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                AND B.NUM_ENDOSO=(SELECT MAX(NUM_ENDOSO) FROM PP_TMAE_POLIZA WHERE NUM_POLIZA=B.NUM_POLIZA) " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                AND B.COD_AFP IN ([CodAFP]) " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "                GROUP BY B.COD_TIPPENSION, B.COD_TIPREN, B.COD_MONEDA, B.COD_TIPREAJUSTE " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "          ) LR ON B.COD_TIPPENSION=LR.COD_TIPPENSION AND B.COD_TIPREN=LR.COD_TIPREN AND B.COD_MONEDA=LR.COD_MONEDA AND B.COD_TIPREAJUSTE=LR.COD_TIPREAJUSTE " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & "          group by B.NUM_FILA " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & " ) A ON A.NUM_FILA=L.NUM_FILA " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & " GROUP BY L.NUM_FILA " & Chr(13)
    vlSqlXMOD = vlSqlXMOD & " ORDER BY 1"
'
End Sub
Private Sub GeneraArchivo(ByVal Cabecera As String, _
                           ByVal CodAFP As String, _
                           ByVal NombreFile As String, _
                           ByVal TipoArchivo As String, _
                           ByVal CadenaQuery As String, _
                           ByVal CadenaCount As String)

    Dim intFile As Integer
    Dim strFile As String
    
    Dim strCountFinal As String
    Dim strSelectFinal As String
 
    Dim RSPAL As ADODB.Recordset
    Dim RS As ADODB.Recordset
    
    Set RS = New ADODB.Recordset
    RS.CursorLocation = adUseServer
    RS.Open CadenaCount, vgConexionBD, adOpenStatic, adLockOptimistic
    
    strFile = "c:\" & NombreFile 'the file you want to save to
    intFile = FreeFile
 
    'Set RS = vgConexionBD.Execute(CadenaCount)
    CountRegistros = RS!cantidad
    
    Barra.Max = CountRegistros
    Barra.Min = 0
    Barra.Value = 0
      
  
  Open strFile For Output As #intFile
      Print #intFile, Cabecera
         
       Set RSPAL = vgConexionBD.Execute(CadenaQuery)
      
    While Not RSPAL.EOF
         Print #intFile, RSPAL!Fila & "|"
         Barra.Value = Barra.Value + 1
         RSPAL.MoveNext
         lblMensaje.Caption = "Procesando " & Barra.Value & " de " & CountRegistros & " Registros."
    Wend
    'Call InicializaCadenasModalidades
    RSPAL.Close
    lblMensaje.Caption = "Terminado"
    Close #intFile

End Sub

Private Sub cmd_exportarPM_Click()


Dim Anexo As String
Dim i As Integer
Dim Arr() As String
Dim vcount As Integer
Dim NombreArchivo As String
Dim swEscogioTipoFile As Boolean
Dim TipoArchivo As String
Dim CadenaQuery As String
Dim CadenaCount As String
Dim CodAFP As String
Dim Periodo As String
Dim InicioPer As String
Dim FinPer As String
Dim DIASF As Integer

vcount = -1
FechaReporte = Format(Now, "YYYYMMDD")

  Periodo = CmbMesExtrae.Text

  DIASF = Day(DateSerial(Left(Periodo, 4), Right(Periodo, 2) + 1, 1 - 1))
  
  InicioPer = Periodo & "01"
  FinPer = Periodo & DIASF


'•   “10” si la información corresponde a AFP Integra.
'•   “13” si la información corresponde a AFP Profuturo.
'•   “16” si la información corresponde a AFP Prima.
'•   “19” si la información corresponde a AFP Habitat.

'
'/*
'241 HORIZONTE
'242 INTEGRA
'243 PROFUTURO
'244 HABITAT
'245 PRIMA
'246 INTEGRA -2
'247 PROFUTURO -2
'*/

CodAFP = ""
NombreArchivo = ""
swEscogioTipoFile = False

For i = 0 To chkAnexo.UBound

         If chkAnexo(i).Value = 1 Then
            vcount = vcount + 1
            IdAnexo = chkAnexo(i).Tag
            
        Select Case IdAnexo
        Case 10
            CodAFP = "242"
            NombreArchivo = "INTEGRA"

        Case 13
            CodAFP = "243"
            NombreArchivo = "PROFUTURO"
        Case 16
            CodAFP = "245"
            NombreArchivo = "PRIMA"
         
        Case 19
            CodAFP = "244"
            NombreArchivo = "HABITAT"
          
        End Select
        
        
         If OptModalidades.Value Then
            NombreArchivo = NombreArchivo & "_Modalidades_" & FechaReporte & ".txt"
            TipoArchivo = "M"
            Call InicializaCadenasModalidades
            vlSqlXMOD = Replace(vlSqlXMOD, "[InicioPer]", InicioPer)
            vlSqlXMOD = Replace(vlSqlXMOD, "[FinPer]", FinPer)
            vlSqlXMOD = Replace(vlSqlXMOD, "[CodAFP]", CodAFP)
            vlSqlXMOD = Replace(vlSqlXMOD, "[PERIODO]", CmbMesExtrae.Text)
            'CadenaQuery = vlSqlXMODQuery & " " & vlSqlXMOD
            'CadenaCount = vlSqlXMODCount & " " & vlSqlXMOD
            
            CadenaQuery = "SELECT lpad(L.FILA,4,'0')|| lpad(COL_40,15,'0')|| lpad(COL_50,15,'0')|| lpad(trunc(COL_60),13,'0')|| lpad((abs(COL_60)-floor(COL_60))*100,2,'0')|| lpad(COL_70,15,'0')||lpad(trunc(COL_80),13,'0')|| lpad((abs(COL_80)-floor(COL_80))*100,2,'0') as fila FROM PP_TTMP_ANEXO10_INT L "
            CadenaCount = "select count(*) cantidad from PP_TTMP_ANEXO10_INT"
            swEscogioTipoFile = True
            
        ElseIf optJubilacion.Value Then
        
            NombreArchivo = NombreArchivo & "_Jubilacion_" & FechaReporte & ".txt"
            swEscogioTipoFile = True
            TipoArchivo = "J"
            'Call InicializaCadenasJubilacion
            vlSqlXPEN = Replace(vlSqlXPEN, "[InicioPer]", InicioPer)
            vlSqlXPEN = Replace(vlSqlXPEN, "[FinPer]", FinPer)
            vlSqlXPEN = Replace(vlSqlXPEN, "[CodAFP]", CodAFP)
            vlSqlXPEN = Replace(vlSqlXPEN, "[PERIODO]", CmbMesExtrae.Text)
            'CadenaQuery = vlSqlXPENQuery & " " & vlSqlXPEN
            'CadenaCount = vlSqlXPENCount & " " & Replace(vlSqlXPEN, " ORDER BY 1,2", "")
            
            CadenaQuery = "SELECT lpad(FILA,4,'0')||"
            CadenaQuery = CadenaQuery & " lpad(NVL(COL_20,0),15,'0')||"
            CadenaQuery = CadenaQuery & " lpad(NVL(COL_30,0),15,'0')||"
            CadenaQuery = CadenaQuery & " lpad(NVL(COL_40,0),15,'0')||"
            CadenaQuery = CadenaQuery & " lpad(NVL(COL_50,0),15,'0')||"
            CadenaQuery = CadenaQuery & " lpad(NVL(COL_60,0),15,'0')||"
            CadenaQuery = CadenaQuery & " lpad(NVL(COL_70,0),15,'0')||"
            CadenaQuery = CadenaQuery & " lpad(NVL(COL_80,0),15,'0')||"
            CadenaQuery = CadenaQuery & " lpad(NVL(COL_90,0),15,'0')||"
            CadenaQuery = CadenaQuery & " lpad(TRUNC(NVL(COL_100,0)),13,'0')||lpad((abs(NVL(COL_100,0))-floor(NVL(COL_100,0)))*100,2,'0')||"
            CadenaQuery = CadenaQuery & " lpad(NVL(COL_110,0),15,'0')||"
            CadenaQuery = CadenaQuery & " lpad(trunc(NVL(COL_120,0)),13,'0')||lpad((abs(NVL(COL_120,0))-floor(NVL(COL_120,0)))*100,2,'0')||"
            CadenaQuery = CadenaQuery & " lpad(trunc(NVL(COL_130,0)),13,'0')||lpad((abs(NVL(COL_130,0))-floor(NVL(COL_130,0)))*100,2,'0') as fila"
            CadenaQuery = CadenaQuery & " FROM PP_TTMP_ANEXO11_INT ORDER BY FILA"

            CadenaCount = "SELECT COUNT(*) cantidad FROM PP_TTMP_ANEXO11_INT"
            IdAnexo = Val(IdAnexo) + 1
            
        ElseIf optSobrevivencia.Value Then
           NombreArchivo = NombreArchivo & "_Sobrevivencia_" & FechaReporte & ".txt"
             swEscogioTipoFile = True
             TipoArchivo = "S"
             IdAnexo = Val(IdAnexo) + 2
             'Call InicializaCadenasSobrevivencia
             
             vlSqlXSov = Replace(vlSqlXSov, "[InicioPer]", InicioPer)
             vlSqlXSov = Replace(vlSqlXSov, "[FinPer]", FinPer)
             vlSqlXSov = Replace(vlSqlXSov, "[CodAFP]", CodAFP)
             vlSqlXSov = Replace(vlSqlXSov, "[PERIODO]", CmbMesExtrae.Text)
             
            'CadenaQuery = vlSqlXSobQuery & " " & vlSqlXSov
            'CadenaCount = vlSqlXSobCount & " " & Replace(vlSqlXSov, " ORDER BY 1", "")
             
            CadenaQuery = CadenaQuery & " select"
            CadenaQuery = CadenaQuery & " LPAD(NVL(FILA,0),4,'0')||"
            CadenaQuery = CadenaQuery & " LPAD(NVL(COL_20,0),15,'0')||"
            CadenaQuery = CadenaQuery & " LPAD(NVL(COL_30,0),15,'0')||"
            CadenaQuery = CadenaQuery & " LPAD(NVL(COL_40,0),15,'0')||"
            CadenaQuery = CadenaQuery & " LPAD(NVL(COL_50,0),15,'0')||"
            CadenaQuery = CadenaQuery & " LPAD(NVL(COL_60,0),15,'0')||"
            CadenaQuery = CadenaQuery & " LPAD(NVL(COL_70,0),15,'0')||"
            CadenaQuery = CadenaQuery & " LPAD(NVL(COL_80,0),15,'0')||"
            CadenaQuery = CadenaQuery & " LPAD(NVL(COL_90,0),15,'0')||"
            CadenaQuery = CadenaQuery & " LPAD(NVL(COL_100,0),15,'0')||"
            CadenaQuery = CadenaQuery & " LPAD('0',15,'0')||"
            CadenaQuery = CadenaQuery & " LPAD(NVL(COL_120,0),15,'0')||"
            CadenaQuery = CadenaQuery & " LPAD(NVL(COL_130,0),15,'0')||"
            CadenaQuery = CadenaQuery & " lpad(trunc(NVL(COL_140,0)),13,'0')||lpad((abs(NVL(COL_140,0))-floor(NVL(COL_140,0)))*100,2,'0')||"
            CadenaQuery = CadenaQuery & " LPAD(NVL(COL_150,0),15,'0')||"
            CadenaQuery = CadenaQuery & " lpad(trunc(NVL(COL_160,0)),13,'0')||lpad((abs(NVL(COL_160,0))-floor(NVL(COL_160,0)))*100,2,'0') as fila"
            CadenaQuery = CadenaQuery & " from PP_TTMP_ANEXO12_INT order by 1"

            CadenaCount = "SELECT COUNT(*) cantidad FROM PP_TTMP_ANEXO12_INT"
            
             
        End If
        
        If CodAFP = "" Or Not swEscogioTipoFile Then
            MsgBox "Debe indicar la AFP y el tipo de información a generar ", vbCritical, "Información Estadistica de Beneficios"
         
            Exit Sub
        End If
        
           Cabecera = IdFormato & IdAnexo & IdSBSEnpresaVigilada & FechaReporte & IdExpresionMontos & "0|"
           Call GeneraArchivo(Cabecera, CodAFP, NombreArchivo, TipoArchivo, CadenaQuery, CadenaCount)
           
            MsgBox ("Se generó el archivo: C\:" & NombreArchivo)
        
         End If


Next


End Sub

Private Sub cmdsalir_Click()
 Unload Me
 
End Sub

Private Sub Form_Load()

Dim AnnioActual As Integer
Dim MesActual As Integer
Dim ItemAnioMes As String
Dim ItemInicio As String

ItemInicio = "201801"

AnnioActual = Format(Now, "YYYY")
MesActual = Format(Now, "MM")

ItemAnioMes = AnnioActual & Right(String(2, "0") & MesActual, 2)



Dim i As Long

For i = ItemInicio To ItemAnioMes

CmbMesExtrae.AddItem (i)

If Right(i, 2) = 12 Then
    i = i + 88
End If



Next


CmbMesExtrae.Text = i - 1


End Sub



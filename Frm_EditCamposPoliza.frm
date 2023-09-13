VERSION 5.00
Begin VB.Form Frm_EditCamposPoliza 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editar Campos Pre Poliza"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   6870
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox txtdirec 
      Height          =   405
      Left            =   360
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   45
      Top             =   6120
      Width           =   6255
   End
   Begin VB.TextBox txt_fono2 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   4080
      TabIndex        =   44
      Top             =   1320
      Width           =   2295
   End
   Begin VB.ComboBox cboCiudad2 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   43
      Top             =   840
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.ComboBox cboTipoTelefono2 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "Frm_EditCamposPoliza.frx":0000
      Left            =   4080
      List            =   "Frm_EditCamposPoliza.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   42
      Top             =   480
      Width           =   2295
   End
   Begin VB.ComboBox cboconj 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2520
      TabIndex        =   22
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Frame Direccion 
      Caption         =   "Direccion"
      Height          =   3855
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Width           =   6375
      Begin VB.CommandButton Cmd_BuscarDir 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Efectuar Busqueda de Dirección"
         Top             =   2610
         Width           =   300
      End
      Begin VB.TextBox Txt_Referencia 
         Height          =   285
         Left            =   1560
         MaxLength       =   40
         TabIndex        =   33
         Top             =   3390
         Width           =   4335
      End
      Begin VB.TextBox txtnombreconj 
         Height          =   285
         Left            =   4800
         TabIndex        =   31
         Top             =   1750
         Width           =   1215
      End
      Begin VB.TextBox txtnumero 
         Height          =   285
         Left            =   960
         TabIndex        =   28
         Top             =   810
         Width           =   735
      End
      Begin VB.TextBox txtdireccion 
         Height          =   285
         Left            =   3960
         TabIndex        =   25
         Top             =   360
         Width           =   2175
      End
      Begin VB.ComboBox cbotipovia 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox Txt_Interior 
         Height          =   285
         Left            =   5040
         MaxLength       =   4
         TabIndex        =   14
         Top             =   840
         Width           =   495
      End
      Begin VB.ComboBox cbodepart 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   800
         Width           =   1455
      End
      Begin VB.TextBox txtmanzana 
         Height          =   285
         Left            =   1155
         TabIndex        =   12
         Top             =   1275
         Width           =   855
      End
      Begin VB.TextBox txtlote 
         Height          =   285
         Left            =   2880
         TabIndex        =   11
         Top             =   1275
         Width           =   735
      End
      Begin VB.TextBox txtetapa 
         Height          =   285
         Left            =   4800
         TabIndex        =   10
         Top             =   1275
         Width           =   735
      End
      Begin VB.ComboBox cbobloque 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   1515
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtbloque 
         Height          =   285
         Left            =   4155
         TabIndex        =   8
         Top             =   2235
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Provincia"
         Height          =   255
         Left            =   2640
         TabIndex        =   40
         Top             =   2685
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Referencia"
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   39
         Top             =   3435
         Width           =   1005
      End
      Begin VB.Label Lbl_DistritoEdit 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   38
         Top             =   3015
         Width           =   3255
      End
      Begin VB.Label Lbl_ProvinciaEdit 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   3480
         TabIndex        =   37
         Top             =   2640
         Width           =   1725
      End
      Begin VB.Label Lbl_DepartamentoEdit 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         TabIndex        =   36
         Top             =   2640
         Width           =   1650
      End
      Begin VB.Label Lbl_Afiliado 
         Caption         =   "Distrito"
         Height          =   255
         Index           =   12
         Left            =   720
         TabIndex        =   35
         Top             =   3090
         Width           =   615
      End
      Begin VB.Label Lbl_Afiliado 
         Caption         =   "Dpto."
         Height          =   255
         Index           =   15
         Left            =   240
         TabIndex        =   34
         Top             =   2685
         Width           =   405
      End
      Begin VB.Label Label6 
         Caption         =   "Nombre Conj.Habit"
         Height          =   255
         Left            =   3240
         TabIndex        =   32
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Conj.Habit"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Numero :"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Direccion :"
         Height          =   375
         Left            =   3120
         TabIndex        =   26
         Top             =   400
         Width           =   855
      End
      Begin VB.Label lblTipoVia 
         Caption         =   "Tipo de Via :"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   400
         Width           =   1215
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Interior"
         Height          =   255
         Index           =   13
         Left            =   4395
         TabIndex        =   21
         Top             =   840
         Width           =   585
      End
      Begin VB.Label lblmanzana 
         Caption         =   "Manzana"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lbllote 
         Caption         =   "Lote"
         Height          =   255
         Left            =   2280
         TabIndex        =   19
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lbletapa 
         Caption         =   "Etapa"
         Height          =   255
         Left            =   3960
         TabIndex        =   18
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Num/Letra"
         Height          =   255
         Left            =   3120
         TabIndex        =   17
         Top             =   2250
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Pref.Depart"
         Height          =   255
         Left            =   1800
         TabIndex        =   16
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblbloque 
         Caption         =   "Bloque/Chalet"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   2235
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Telefono"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.ComboBox cboTipoTelefono 
         BackColor       =   &H80000018&
         Height          =   315
         ItemData        =   "Frm_EditCamposPoliza.frx":0004
         Left            =   1560
         List            =   "Frm_EditCamposPoliza.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   2295
      End
      Begin VB.ComboBox cboCiudad 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   780
         Width           =   2295
      End
      Begin VB.TextBox Txt_Fono 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label lbltipotelefono 
         Caption         =   "Tipo Telefono : "
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label lblciudad 
         Caption         =   "Ciudad : "
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   780
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblnumerotelf 
         Caption         =   "Numero :"
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   1260
         Width           =   855
      End
   End
End
Attribute VB_Name = "Frm_EditCamposPoliza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vNumPoliza As String
Public vNumendoso As String
Public vNumOrden As String
Public vlCodDireccion As String
Public vTipoTelefono As String
Public vNumTelefono As String
Public vCodigoTelefono As String
Public vTipoTelefono2 As String
Public vNumTelefono2 As String
Public vCodigoTelefono2 As String
Public vTipoVia As String
Public vDireccion As String
Public vNumero As String
Public vTipoPref As String
Public vInterior As String
Public vManzana As String
Public vLote As String
Public vEtapa As String
Public vTipoConj As String
Public vConjHabit As String
Public vTipoBlock As String
Public vNumBlock As String
Public vReferencia As String
Public vcodeDepar As String
Public vcodeProv As String
Public vCodeDistr As String
Public vConcatDirec As String
Public vCodLoad As Integer




Private Sub cboTipoTelefono_Click()
Dim valor As String
valor = fgObtenerCodigo_TextoCompuesto(cboTipoTelefono)
If valor <> "2" Then
 cboCiudad.Visible = True
 lblciudad.Visible = True
 Else
 cboCiudad.Visible = False
 lblciudad.Visible = False
End If
End Sub
Private Sub cboTipoTelefono2_Click()
Dim valor As String
valor = fgObtenerCodigo_TextoCompuesto(cboTipoTelefono2)
If valor <> "2" Then
 cboCiudad.Visible = True
 lblciudad.Visible = True
 Else
 cboCiudad.Visible = False
 lblciudad.Visible = False
End If
End Sub

Private Sub cbotipovia_Click()
If fgObtenerCodigo_TextoCompuesto(cbotipovia.Text) = "99" Then
    txtdireccion.Locked = True
    txtnumero.Locked = True
    If vCodLoad = 0 Then
    txtnumero.Text = ""
    txtdireccion.Text = ""
    End If
    Else
    txtdireccion.Locked = False
    txtnumero.Locked = False
End If
End Sub

Private Sub Cmd_BuscarDir_Click()
On Error GoTo Err_Buscar

  Frm_BusDireccion.flInicio ("Frm_EditCamposPoliza")

  
    
    
Exit Sub
Err_Buscar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
 
End Sub

Private Function ValidarNull(valor As String) As String
Dim retornar As String
If Trim(valor) = "" Then
    retornar = ""
Else
    retornar = valor
End If
ValidarNull = retornar
End Function
Private Sub setValirables()
vCodLoad = 1
Txt_Fono.Text = vNumTelefono
txt_fono2.Text = vNumTelefono2
txtdireccion.Text = vDireccion
txtnumero.Text = vNumero
Txt_Interior.Text = vInterior
txtmanzana.Text = vManzana
txtlote.Text = vLote
txtetapa.Text = vEtapa
txtnombreconj.Text = vConjHabit
txtbloque.Text = vNumBlock
Txt_Referencia.Text = vReferencia
If vlCodDireccion <> "" Then
Call fgBuscarNombreComunaProvinciaRegion(vlCodDireccion)
End If

If vTipoTelefono = "" Then
vTipoTelefono = 4
End If
If vTipoTelefono2 = "" Then
vTipoTelefono2 = 2
End If

If vCodigoTelefono = "" Then
vCodigoTelefono = "1"
End If


Call llenarComboValue(cboTipoTelefono, vTipoTelefono)
Call llenarComboValue(cboCiudad, vCodigoTelefono)
Call llenarComboValue(cboTipoTelefono2, vTipoTelefono2)
Call llenarComboValue(cboCiudad2, vCodigoTelefono2)
Call llenarComboValue(cbotipovia, vTipoVia)
Call llenarComboValue(cbodepart, vTipoPref)
Call llenarComboValue(cboconj, vTipoConj)
Call llenarComboValue(cbobloque, vTipoBlock)



Call cboTipoTelefono_Click

vCodLoad = 0

End Sub

Private Sub llenarComboValue(combo As ComboBox, Value As String)
If Value <> "" Then
Call fgBuscaPos(combo, Value)
Else
combo.ListIndex = -1
End If
End Sub

Private Sub GetVariables()
vTipoTelefono = UCase(ValidarNull(CStr(fgObtenerCodigo_TextoCompuesto(cboTipoTelefono.Text))))
vNumTelefono = UCase(ValidarNull(Txt_Fono.Text))
vCodigoTelefono = UCase(ValidarNull(fgObtenerCodigo_TextoCompuesto(cboCiudad.Text)))
vTipoTelefono2 = UCase(ValidarNull(CStr(fgObtenerCodigo_TextoCompuesto(cboTipoTelefono2.Text))))
vNumTelefono2 = UCase(ValidarNull(txt_fono2.Text))
vCodigoTelefono2 = ValidarNull(fgObtenerCodigo_TextoCompuesto(cboCiudad2.Text))
vTipoVia = UCase(ValidarNull(fgObtenerCodigo_TextoCompuesto(cbotipovia.Text)))
vDireccion = UCase(ValidarNull(txtdireccion.Text))
vNumero = UCase(ValidarNull(txtnumero.Text))
vTipoPref = UCase(ValidarNull(fgObtenerCodigo_TextoCompuesto(cbodepart.Text)))
vInterior = UCase(ValidarNull(Txt_Interior.Text))
vManzana = UCase(ValidarNull(txtmanzana.Text))
vLote = UCase(ValidarNull(txtlote.Text))
vEtapa = UCase(ValidarNull(txtetapa.Text))
vTipoConj = ValidarNull(fgObtenerCodigo_TextoCompuesto(cboconj.Text))
vConjHabit = UCase(ValidarNull(txtnombreconj.Text))
vTipoBlock = UCase(ValidarNull(fgObtenerCodigo_TextoCompuesto(cbobloque.Text)))
vNumBlock = UCase(ValidarNull(txtbloque.Text))
vReferencia = UCase(ValidarNull(Txt_Referencia.Text))
End Sub

Private Sub Command1_Click()
Call GetVariables
If vTipoTelefono = "2" Then
 vCodigoTelefono = ""
End If
Call ConcatenarDireccion
Call Frm_EndosoPol.flRecibeParametrosEditDirc(vNumOrden, vlCodDireccion, vTipoTelefono, vNumTelefono, vCodigoTelefono, vTipoTelefono2, vNumTelefono2, vCodigoTelefono2, vTipoVia, vDireccion, vNumero, vTipoPref, vInterior, vManzana, vLote, vEtapa, vTipoConj, vConjHabit, vTipoBlock, vNumBlock, vReferencia, vConcatDirec)
Unload Me
End Sub
Private Sub ConcatenarDireccion()
Dim vlRegistroDir As ADODB.Recordset
On Error GoTo Err_Buscar
 Dim tipo_via As String
 Dim tipo_bloque As String
 Dim tipo_interior As String
 Dim tipo_cjht As String
Call fgBuscarNombreComunaProvinciaRegion(vlCodDireccion)

     vgSql = "SELECT"
     vgSql = vgSql + "(SELECT TRIM(GLS_DESCRIPCION)"
     vgSql = vgSql + " FROM MA_TPAR_TIPO_VIA T "
     vgSql = vgSql + " WHERE T.COD_DIRE_VIA = '" + vTipoVia + "') as TIPO_VIA,"
     vgSql = vgSql + " (SELECT TRIM(GLS_DESCRIPCION)"
     vgSql = vgSql + " FROM MA_TPAR_TIPO_BLOQUE T"
     vgSql = vgSql + " WHERE T.COD_BLOCKCHALET =  '" + vTipoBlock + "') AS TIPO_BLOQUE,"
     vgSql = vgSql + " (SELECT TRIM(GLS_DESCRIPCION)"
     vgSql = vgSql + " FROM MA_TPAR_TIPO_INTERIOR T"
     vgSql = vgSql + " WHERE T.COD_INTERIOR = '" + vTipoPref + "') AS TIPO_INTERIOR,"
     vgSql = vgSql + "(SELECT TRIM(GLS_DESCRIPCION)"
     vgSql = vgSql + " FROM MA_TPAR_TIPO_CJHT T"
     vgSql = vgSql + " WHERE T.COD_CJHT = '" + vTipoConj + "' ) AS TIPO_CJHT"
     vgSql = vgSql + " FROM DUAL"
     Set vlRegistroDir = vgConexionBD.Execute(vgSql)
     If Not vlRegistroDir.EOF Then
        tipo_via = IIf(IsNull(vlRegistroDir!tipo_via), "", vlRegistroDir!tipo_via)
        tipo_bloque = IIf(IsNull(vlRegistroDir!tipo_bloque), "", vlRegistroDir!tipo_bloque)
        tipo_interior = IIf(IsNull(vlRegistroDir!tipo_interior), "", vlRegistroDir!tipo_interior)
        tipo_cjht = IIf(IsNull(vlRegistroDir!tipo_cjht), "", vlRegistroDir!tipo_cjht)
     End If
     vlRegistroDir.Close
     Dim Strmanzana As String
     If Trim(vManzana) <> "" Then
       Strmanzana = " Manzana " & vManzana
     Else
       Strmanzana = ""
     End If
     
     Dim StrLote As String
     If Trim(vLote) <> "" Then
       StrLote = " Lote " & vLote
     Else
       StrLote = ""
     End If
     Dim StrEtapa As String
     If Trim(vEtapa) <> "" Then
       StrEtapa = " Etapa " & vEtapa
     Else
       StrEtapa = ""
     End If
     Dim strBloque As String
     If Trim(tipo_bloque) <> "" And Trim(vNumBlock) <> "" Then
     strBloque = " " & tipo_bloque & " " & Trim(vNumBlock)
     Else
     strBloque = ""
     End If
     Dim StrInterior As String
     If Trim(tipo_interior) <> "" And Trim(vInterior) <> "" Then
     StrInterior = " " & tipo_interior & " " & Trim(vInterior)
     Else
     StrInterior = ""
     End If
     Dim StrCjht As String
     If vTipoConj = "99" Then
     StrCjht = " " & vConjHabit
     Else
        If Trim(tipo_cjht) <> "" And Trim(vConjHabit) <> "" Then
        StrCjht = " " & tipo_cjht & " " & vConjHabit
        Else
        StrCjht = ""
        End If
     End If
     
     Dim StrDireccion As String
     If vTipoVia = "88" Or vTipoVia = "99" Then
     StrDireccion = " " & vDireccion
    Else
        If Trim(tipo_via) <> "" And Trim(vDireccion) <> "" Then
        StrDireccion = " " & tipo_via & " " & vDireccion
        Else
        StrDireccion = ""
        End If
     End If
    
     vConcatDirec = Trim(UCase(IIf(StrDireccion = "", "", StrDireccion) & IIf(Trim(vNumero) = "", "", " " & Trim(vNumero)) _
        & "" & strBloque & "" & StrInterior _
        & "" & StrCjht & "" _
        & StrEtapa & Strmanzana & StrLote _
        & IIf(IsNull(vgNombreRegion), "", " " & Trim(vgNombreRegion)) & "" _
        & IIf(IsNull(vgNombreProvincia), "", " " & Trim(vgNombreProvincia)) & "" _
        & IIf(IsNull(vgNombreComuna), "", " " & Trim(vgNombreComuna))))
     
     
    
       
 Exit Sub
Err_Buscar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub
Public Function CambiarNullxVacio(valor As String) As String
If valor = "Null" Then
CambiarNullxVacio = ""
Else
CambiarNullxVacio = valor
End If
End Function
Public Sub bloquearOpciones(SubTipoEndoso As String)
If SubTipoEndoso = "15" Then
Txt_Fono.Enabled = False
txt_fono2.Enabled = False
cboTipoTelefono.Enabled = False
cboTipoTelefono.Enabled = False
ElseIf SubTipoEndoso = "16" Then
cbotipovia.Enabled = False
txtdireccion.Enabled = False
txtnumero.Enabled = False
cbodepart.Enabled = False
Txt_Interior.Enabled = False
txtmanzana.Enabled = False
txtlote.Enabled = False
txtetapa.Enabled = False
cboconj.Enabled = False
txtnombreconj.Enabled = False
cbobloque.Enabled = False
txtbloque.Enabled = False
Txt_Referencia.Enabled = False
txtdirec.Enabled = False
ElseIf SubTipoEndoso = "27" Then
cboTipoTelefono.Enabled = True
cboTipoTelefono.Enabled = True
cbotipovia.Enabled = True
txtdireccion.Enabled = True
txtnumero.Enabled = True
cbodepart.Enabled = True
Txt_Interior.Enabled = True
txtmanzana.Enabled = True
txtlote.Enabled = True
txtetapa.Enabled = True
cboconj.Enabled = True
txtnombreconj.Enabled = True
cbobloque.Enabled = True
txtbloque.Enabled = True
Txt_Referencia.Enabled = True
txtdirec.Enabled = True
Txt_Fono.Enabled = True
txt_fono2.Enabled = True
End If
End Sub
Function flIniciarValores(vTipoEndoso As String, vSubTipoEndoso, vNumOrden1 As String, pgCodDireccion As String, pTipoTelefono As String, pNumTelefono As String, pCodigoTelefono As String, pTipoTelefono2 As String, pNumTelefono2 As String, pCodigoTelefono2 As String, pTipoVia As String, pDireccion As String, pNumero As String, pTipoPref As String, pInterior As String, pManzana As String, pLote As String, pEtapa As String, pTipoConj As String, pConjHabit As String, pTipoBlock As String, pNumBlock As String, pReferencia As String, pTextoDirec As String)
If vTipoEndoso = "O" Then
 'Cmd_BuscarDir.Enabled = False
Else
Cmd_BuscarDir.Enabled = True
End If
bloquearOpciones (vSubTipoEndoso)
vNumOrden = CambiarNullxVacio(vNumOrden1)
 vlCodDireccion = CambiarNullxVacio(pgCodDireccion)
 vTipoTelefono = CambiarNullxVacio(pTipoTelefono)
 vNumTelefono = CambiarNullxVacio(pNumTelefono)
 vCodigoTelefono = CambiarNullxVacio(pCodigoTelefono)
 vTipoTelefono2 = CambiarNullxVacio(pTipoTelefono2)
 vNumTelefono2 = CambiarNullxVacio(pNumTelefono2)
 vCodigoTelefono2 = CambiarNullxVacio(pCodigoTelefono2)
 vTipoVia = CambiarNullxVacio(pTipoVia)
 vDireccion = CambiarNullxVacio(pDireccion)
 vNumero = CambiarNullxVacio(pNumero)
 vTipoPref = CambiarNullxVacio(pTipoPref)
 vInterior = CambiarNullxVacio(pInterior)
 vManzana = CambiarNullxVacio(pManzana)
 vLote = CambiarNullxVacio(pLote)
 vEtapa = CambiarNullxVacio(pEtapa)
 vTipoConj = CambiarNullxVacio(pTipoConj)
 vConjHabit = CambiarNullxVacio(pConjHabit)
 vTipoBlock = CambiarNullxVacio(pTipoBlock)
 vNumBlock = CambiarNullxVacio(pNumBlock)
 vReferencia = CambiarNullxVacio(pReferencia)
 txtdirec.Text = pTextoDirec
 Call setValirables
End Function
Private Sub Limpiar()
Txt_Fono.Text = ""
txt_fono2.Text = ""
txtdireccion.Text = ""
txtnumero.Text = ""
Txt_Interior.Text = ""
txtmanzana.Text = ""
txtlote.Text = ""
txtetapa.Text = ""
txtnombreconj.Text = ""
txtbloque.Text = ""
Txt_Referencia.Text = ""
End Sub
Private Sub Form_Load()
vCodLoad = 1
Call Limpiar
Call llenarTipoInterior(cbodepart)
Call llenarTipoBloque(cbobloque)
Call llenarTipoConj(cboconj)
Call llenarTipoVia(cbotipovia)
Call llenarcomboTipoTelefono(cboTipoTelefono)
Call llenarCodigoTelefono(cboCiudad)
Call llenarcomboTipoTelefono(cboTipoTelefono2)
Call llenarCodigoTelefono(cboCiudad2)
Call setValirables
If vTipoTelefono = "" Then
Call llenarComboValue(cboTipoTelefono, "2")
End If
cboTipoTelefono.Enabled = False
cboTipoTelefono2.Enabled = False
Cmd_BuscarDir.Visible = False
vCodLoad = 0
End Sub
Function flRecibeDireccionEdit(iNomDepartamento As String, iNomProvincia As String, iNomDistrito As String, iCodDir As String)
'FUNCION QUE RECIBE LOS DATOS DEL FORMULARIO DE BUSQUEDA de Dirección
 Lbl_DepartamentoEdit = Trim(iNomDepartamento)
    Lbl_ProvinciaEdit = Trim(iNomProvincia)
    Lbl_DistritoEdit = Trim(iNomDistrito)
    vlCodDireccion = iCodDir
    Frm_EditCamposPoliza.Enabled = True
End Function
Private Sub Form_Unload(Cancel As Integer)
    Frm_EndosoPol.Enabled = True
End Sub
Public Sub llenarcomboTipoTelefono(iCombo As ComboBox)
Dim vlRsCombo As ADODB.Recordset
Dim vlDescripcion As String * 50

On Error GoTo Err_ComboGeneral

    iCombo.Clear
    
    iCombo.AddItem ""
    vgSql = "select cod_tipo_telefono,gls_descripcion  from ma_tpar_tipo_telefono where "
    vgSql = vgSql & "cod_estado = 1"
    Set vlRsCombo = vgConexionBD.Execute(vgSql)
    Do While Not vlRsCombo.EOF
        vlDescripcion = (vlRsCombo!COD_TIPO_TELEFONO) & " - " & Trim((vlRsCombo!gls_descripcion))
        'iCombo.AddItem vlDescripcion & Trim(vlRsCombo!cod_scomp)
        iCombo.AddItem vlDescripcion
        vlRsCombo.MoveNext
    Loop
    vlRsCombo.Close
    
    If iCombo.ListCount <> 0 Then
        iCombo.ListIndex = 0
    End If

Exit Sub
Err_ComboGeneral:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub

Public Sub llenarCodigoTelefono(iCombo As ComboBox)
Dim vlRsCombo As ADODB.Recordset
Dim vlDescripcion As String * 50

On Error GoTo Err_ComboGeneral

    iCombo.Clear
    
    iCombo.AddItem ""
  
    vgSql = "SELECT COD_AREA,GLS_REGION FROM MA_TPAR_TIPO_AREA MTTA , MA_TPAR_REGION MTR "
    vgSql = vgSql & "Where MTTA.COD_REGION = MTR.COD_REGION ORDER BY COD_AREA ASC"
    Set vlRsCombo = vgConexionBD.Execute(vgSql)
    Do While Not vlRsCombo.EOF
        vlDescripcion = (vlRsCombo!COD_AREA) & " - " & Trim((vlRsCombo!gls_region))
        'iCombo.AddItem vlDescripcion & Trim(vlRsCombo!cod_scomp)
        iCombo.AddItem vlDescripcion
        vlRsCombo.MoveNext
    Loop
    vlRsCombo.Close
    
    If iCombo.ListCount <> 0 Then
        iCombo.ListIndex = 0
    End If

Exit Sub
Err_ComboGeneral:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub

Public Sub llenarTipoInterior(iCombo As ComboBox)
Dim vlRsCombo As ADODB.Recordset
Dim vlDescripcion As String * 50

On Error GoTo Err_ComboGeneral

    iCombo.Clear
    
    iCombo.AddItem ""
    vgSql = "select cod_interior,gls_descripcion as  from ma_tpar_tipo_interior where "
    vgSql = vgSql & "cod_estado = 1"
    Set vlRsCombo = vgConexionBD.Execute(vgSql)
    Do While Not vlRsCombo.EOF
        vlDescripcion = (vlRsCombo!COD_INTERIOR) & " - " & Trim((vlRsCombo!gls_descripcion))
        'iCombo.AddItem vlDescripcion & Trim(vlRsCombo!cod_scomp)
        iCombo.AddItem vlDescripcion
        vlRsCombo.MoveNext
    Loop
    vlRsCombo.Close
    
    If iCombo.ListCount <> 0 Then
        iCombo.ListIndex = 0
    End If

Exit Sub
Err_ComboGeneral:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub


Public Sub llenarTipoBloque(iCombo As ComboBox)
Dim vlRsCombo As ADODB.Recordset
Dim vlDescripcion As String * 50

On Error GoTo Err_ComboGeneral

    iCombo.Clear
    
    iCombo.AddItem ""
    vgSql = "select cod_blockchalet ,gls_descripcion  from MA_TPAR_TIPO_BLOQUE where "
    vgSql = vgSql & "cod_Estado = 1"
    Set vlRsCombo = vgConexionBD.Execute(vgSql)
    Do While Not vlRsCombo.EOF
        vlDescripcion = (vlRsCombo!cod_blockchalet) & " - " & Trim((vlRsCombo!gls_descripcion))
        'iCombo.AddItem vlDescripcion & Trim(vlRsCombo!cod_scomp)
        iCombo.AddItem vlDescripcion
        vlRsCombo.MoveNext
    Loop
    vlRsCombo.Close
    
    If iCombo.ListCount <> 0 Then
        iCombo.ListIndex = 0
    End If

Exit Sub
Err_ComboGeneral:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Public Sub llenarTipoVia(iCombo As ComboBox)
Dim vlRsCombo As ADODB.Recordset
Dim vlDescripcion As String * 50

On Error GoTo Err_ComboGeneral

    iCombo.Clear
    
    iCombo.AddItem ""
    vgSql = "select COD_DIRE_VIA ,GLS_DESCRIPCION  from MA_TPAR_TIPO_VIA where "
    vgSql = vgSql & "COD_ESTADO = 1"
    Set vlRsCombo = vgConexionBD.Execute(vgSql)
    Do While Not vlRsCombo.EOF
        vlDescripcion = (vlRsCombo!cod_dire_via) & " - " & Trim((vlRsCombo!gls_descripcion))
        'iCombo.AddItem vlDescripcion & Trim(vlRsCombo!cod_scomp)
        iCombo.AddItem vlDescripcion
        vlRsCombo.MoveNext
    Loop
    vlRsCombo.Close
    
    If iCombo.ListCount <> 0 Then
        iCombo.ListIndex = 0
    End If

Exit Sub
Err_ComboGeneral:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Public Sub llenarTipoConj(iCombo As ComboBox)
Dim vlRsCombo As ADODB.Recordset
Dim vlDescripcion As String * 50

On Error GoTo Err_ComboGeneral

    iCombo.Clear
    iCombo.AddItem ""
    vgSql = "select COD_CJHT ,GLS_DESCRIPCION  from MA_TPAR_TIPO_CJHT where "
    vgSql = vgSql & "COD_ESTADO = 1 AND FLG_USO_RV = 1"
    Set vlRsCombo = vgConexionBD.Execute(vgSql)
    Do While Not vlRsCombo.EOF
        vlDescripcion = (vlRsCombo!cod_cjht) & " - " & Trim((vlRsCombo!gls_descripcion))
        'iCombo.AddItem vlDescripcion & Trim(vlRsCombo!cod_scomp)
        iCombo.AddItem vlDescripcion
        vlRsCombo.MoveNext
    Loop
    vlRsCombo.Close
    
    If iCombo.ListCount <> 0 Then
        iCombo.ListIndex = 0
    End If

Exit Sub
Err_ComboGeneral:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Public Function fgObtenerCodigo_TextoCompuesto(iTexto As String) As String
'Función: Permite obtener el Código de un Texto que tiene el Código y la
'Descripción separados por un Guión
'Parámetros de Entrada :
'- iTexto     => Texto que contiene el Código y Descripción
'Parámetros de Salida :
'- Devuelve el código del Texto
    
    If (InStr(1, iTexto, "-") <> 0) Then
        fgObtenerCodigo_TextoCompuesto = Trim(Mid(iTexto, 1, InStr(1, iTexto, "-") - 1))
    Else
        fgObtenerCodigo_TextoCompuesto = UCase(Trim(iTexto))
    End If

End Function

Function fgBuscarNombreComunaProvinciaRegion(vlCodDir As String)
Dim vlRegistroDir As ADODB.Recordset
On Error GoTo Err_Buscar

     vgSql = "SELECT r.Cod_Region,r.Gls_Region,p.Cod_Provincia,p.Gls_Provincia,c.Cod_Comuna,c.Gls_Comuna"
     vgSql = vgSql & " FROM MA_TPAR_COMUNA c, MA_TPAR_PROVINCIA p, MA_TPAR_REGION r"
     vgSql = vgSql & " Where c.Cod_Direccion = '" & vlCodDir & "' and  "
     vgSql = vgSql & " c.cod_region = p.cod_region and"
     vgSql = vgSql & " c.cod_provincia = p.cod_provincia and"
     vgSql = vgSql & " p.cod_region = r.cod_region"
     Set vlRegistroDir = vgConexionBD.Execute(vgSql)
     If Not vlRegistroDir.EOF Then
        Lbl_DepartamentoEdit = (vlRegistroDir!gls_region)
        Lbl_ProvinciaEdit = (vlRegistroDir!gls_provincia)
        Lbl_DistritoEdit = (vlRegistroDir!gls_comuna)
     End If
     vlRegistroDir.Close

Exit Function
Err_Buscar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function
Function fgBuscaPos(iCombo As ComboBox, iCodigo)
        vgI = 0
        iCombo.ListIndex = 0
        Do While vgI < iCombo.ListCount
            If (Trim(iCombo) <> "") Then
                If (Trim(iCodigo) = Trim(Mid(iCombo.Text, 1, (InStr(1, iCombo, "-") - 1)))) Then
                    Exit Do
                End If
            End If
            vgI = vgI + 1
            If (vgI = iCombo.ListCount) Then
                iCombo.ListIndex = 0
                Exit Do
            End If
                iCombo.ListIndex = vgI
        Loop
End Function



Attribute VB_Name = "Mod_BasDat"
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Function EscribeArchivoIni(ByVal Seccion$, ByVal Item$, ByVal Default$, ByVal NombreArchivo$) As Integer
    EscribeArchivoIni = WritePrivateProfileString(Seccion$, Item$, Default$, NombreArchivo$)
End Function

Function fgConexionBaseDatos(oConBD As ADODB.Connection)
    Dim StringConexion As String
    'Por defecto supone que falla la Conexión
    fgConexionBaseDatos = False

On Local Error GoTo Err_ConsultaBD
    'String de Conexión
    If vgTipoBase = "ORACLE" Then
        StringConexion = "Provider=" & ProviderName & ";"
        StringConexion = StringConexion & "Server= " & vgNombreServidor & " ;"
        StringConexion = StringConexion & "User ID= " & vgNombreUsuario & " ;"
        StringConexion = StringConexion & "Password= " & vgPassWord & ";"
        StringConexion = StringConexion & "Data Source=" & vgNombreBaseDatos & " "
    Else
        StringConexion = "driver={Sql Server}; server=" & vgNombreServidor & ";UID=" & vgNombreUsuario & ";PWD=" & vgPassWord & ";database=" & vgNombreBaseDatos
    End If

    Set oConBD = New ADODB.Connection
    oConBD.ConnectionString = StringConexion
   
    'oConBD.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Data Source=" & vgRutaBasedeDatos & ";Persist Security Info=False"
    'oConBD.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & vgRutaBasedeDatos & ";Persist Security Info=False"
    oConBD.ConnectionTimeout = 1800
    oConBD.CommandTimeout = 1800
    oConBD.Open
    'oConBD.BeginTrans
    'La Conexión fue realizada
    fgConexionBaseDatos = True
   ' oConBD.CommitTrans
    Exit Function

Err_ConsultaBD:
    'MsgBox "La consulta a la Base de Datos no fue llevada a cabo. Error : " & vbCrLf & Err.Description, vbCritical, "Error de Consulta a Base de Datos"
    'Err.Clear
    'Text3.Text = "Errores en la Impresión de Solicitudes"
    MsgBox "La consulta a la Base de Datos no fue llevada a cabo. Error : [ " & Err & Space(4) & Err.Description & " ]"
End Function

Function fgGetPrivateIni(section, key$, FnameIni)
Dim retVal As String
Dim AppName As String
Dim worked As Integer

retVal = String$(255, 0)
worked = GetPrivateProfileString(section, key, "", retVal, Len(retVal), FnameIni)
If (worked = 0) Then
    fgGetPrivateIni = "DESCONOCIDO"
Else
    fgGetPrivateIni = Left(retVal, InStr(retVal, Chr(0)) - 1)
End If
End Function

Function LeeArchivoIni(ByVal Seccion$, ByVal Item$, ByVal Default$, ByVal NombreArchivo$) As String
Dim temp As String
Dim x   As Integer

    temp = String$(2048, 32)
    x = GetPrivateProfileString(Seccion$, Item$, Default$, temp, Len(temp), NombreArchivo$)
    LeeArchivoIni = Mid$(temp, 1, x)

End Function

    

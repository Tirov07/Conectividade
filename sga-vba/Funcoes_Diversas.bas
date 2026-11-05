Attribute VB_Name = "Funcoes_Diversas"
Option Compare Database
Option Explicit
Private oemail As CanadaUtils.Email
Private oFTP As CanadaUtils.FTP

#If VBA7 Then
    'Public Declare PtrSafe PtrSafePtrSafePtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As LongPtr _
    '  , ByVal szURL As String _
    '  , ByVal szFileName As String _
    '  , ByVal dwReserved As LongPtr _
    ' ByVal lpfnCB As LongPtr) As Long
    'Public Declare PtrSafe PtrSafePtrSafePtrSafe Sub 'sleep Lib "kernel32" ( _
    '    ByVal dwMilliseconds As LongLong)
#Else
   Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias _
    "URLDownloadToFileA" (ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long) As Long
    Public Declare PtrSafe Sub Sleep Lib "kernel32" Alias "sleep" (ByVal dwMilliseconds As Long)
    Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Public Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
    Public Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
    Public Declare PtrSafe Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Public Const GW_HWNDNEXT = 2
    Public Const SW_MAXIMIZE = 1
#End If


Type Dados
     Seq As String * 4
     Item As String * 3
     Status As String * 1
     Transacao As String * 8
     Escape As String * 2
     Reducao As String * 1
     Liq As String * 12
     Bruto As String * 12
     Erro As String * 1
     Data As String * 6
     Hora As String * 4
     Verao As String * 1
     Menserro As String * 35
     Ecf As String * 3
     Papel As String * 1
     Abrev As String * 1
     Falta As String * 12
     REC As String * 12
     Fazx As String * 1
     Abrir As String * 1
     Prog As String * 1
     Vinc As String * 1
End Type

Public Declare PtrSafe Function InternetGetConnectedState Lib "wininet" _
(ByRef dwFlags As Long, _
ByVal dwReserved As Long) As Long

Dim irow
Dim icol

Public Sub testearq()

Call UpLoad_Extrato_FTP(5, "C:\sga\EXT_85087190944_23092020_155618.txt")

End Sub

Function RetornaURLExtrato(ByVal wCodCliente As Long) As String
RetornaURLExtrato = "https://www.canainfo.com.br/" & wCodCliente & Nroserie & wCodCliente & "/ExtPlacas.txt"
End Function

Function UpLoad_Extrato_FTP(wCodCliente As Long, wFileName As String) As Boolean
On Error GoTo TratamentodeErro:
Set oFTP = New CanadaUtils.FTP
Dim WUsuario, wsenha As String

Dim pDIrectoryFTPClient As String
pDIrectoryFTPClient = wCodCliente & Nroserie & wCodCliente

Dim pAddressFtp As String
pAddressFtp = "ftp://ftp.canainfo.com.br/public_html/"

WUsuario = Captura_Conteudo_Campos_Empresa(Empresa_Ativa(), "Usuario_Ftp")
wsenha = Captura_Conteudo_Campos_Empresa(Empresa_Ativa(), "Senha_Ftp")

oFTP.ContaFTP = WUsuario
oFTP.SenhaFTP = wsenha
oFTP.CaminhoLocal = wFileName
oFTP.CaminhoUpServidor = pAddressFtp & pDIrectoryFTPClient & "/ExtPlacas.txt"

If oFTP.DirectoryExists(pAddressFtp & pDIrectoryFTPClient & "/") = False Then
   oFTP.Create_Directory_FTP pAddressFtp, pDIrectoryFTPClient
End If

oFTP.UpLoad_File
If oFTP.StatusProcesso = True Then
   UpLoad_Extrato_FTP = True
Else
  ' MsgBox oFTP.MensagemProcesso
   
   UpLoad_Extrato_FTP = False
End If
Set oFTP = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Ret_Mod_Equip_Cons_InMetro(WBusca As String, WChave As Long) As String
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Dim WValor As String
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Consulta_Selagem_Completa", DB_OPEN_TABLE)

If WChave = 1 Then
   tb.index = "Placa"
Else
   tb.index = "Renavam"
End If

tb.Seek "=", WBusca
If tb.NoMatch Then
   GoSub Finaliza:
Else
   If Not IsNull(tb!ModeloEquip) Then
      WValor = tb!ModeloEquip
   Else
      WValor = "INDEF"
   End If
End If

Finaliza:
  Ret_Mod_Equip_Cons_InMetro = WValor
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Ret_Marc_Equip_Cons_InMetro(WBusca As String, WChave As Long) As String
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Dim WValor As String
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Consulta_Selagem_Completa", DB_OPEN_TABLE)

If WChave = 1 Then
   tb.index = "Placa"
Else
   tb.index = "Renavam"
End If

tb.Seek "=", WBusca
If tb.NoMatch Then
   GoSub Finaliza:
Else
   If Not IsNull(tb!MarcaEquip) Then
      WValor = tb!MarcaEquip
   Else
      WValor = "INDEF"
   End If
End If

Finaliza:
  Ret_Marc_Equip_Cons_InMetro = WValor
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function
Function Ret_Ren_Veic_Cons_InMetro(WBusca As String, WChave As Long) As String
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Dim WValor As String
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Consulta_Selagem_Completa", DB_OPEN_TABLE)

If WChave = 1 Then
   tb.index = "Placa"
Else
   tb.index = "Renavam"
End If

tb.Seek "=", WBusca
If tb.NoMatch Then
   GoSub Finaliza:
Else
   If Not IsNull(tb!Renavam) Then
      WValor = tb!Renavam
   Else
      WValor = "INDEF"
   End If
End If

Finaliza:
  Ret_Ren_Veic_Cons_InMetro = WValor
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Ret_Plac_Veic_Cons_InMetro(WBusca As String, WChave As Long) As String
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Dim WValor As String
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Consulta_Selagem_Completa", DB_OPEN_TABLE)

If WChave = 1 Then
   tb.index = "Placa"
Else
   tb.index = "Renavam"
End If

tb.Seek "=", WBusca
If tb.NoMatch Then
   GoSub Finaliza:
Else
   If Not IsNull(tb!Placa) Then
      WValor = tb!Placa
   Else
      WValor = "INDEF"
   End If
End If

Finaliza:
  Ret_Plac_Veic_Cons_InMetro = WValor
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Ret_Pneu_Veic_Cons_InMetro(WBusca As String, WChave As Long) As String
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Dim WValor As String
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Consulta_Selagem_Completa", DB_OPEN_TABLE)

If WChave = 1 Then
   tb.index = "Placa"
Else
   tb.index = "Renavam"
End If

tb.Seek "=", WBusca
If tb.NoMatch Then
   GoSub Finaliza:
Else
   If Not IsNull(tb!Pneu) Then
      WValor = tb!Pneu
   Else
      WValor = "INDEF"
   End If
End If

Finaliza:
  Ret_Pneu_Veic_Cons_InMetro = WValor
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Ret_Aro_Veic_Cons_InMetro(WBusca As String, WChave As Long) As String
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Dim WValor As String
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Consulta_Selagem_Completa", DB_OPEN_TABLE)

If WChave = 1 Then
   tb.index = "Placa"
Else
   tb.index = "Renavam"
End If

tb.Seek "=", WBusca
If tb.NoMatch Then
   GoSub Finaliza:
Else
   If Not IsNull(tb!Aro) Then
      WValor = tb!Aro
   Else
      WValor = "INDEF"
   End If
End If

Finaliza:
  Ret_Aro_Veic_Cons_InMetro = WValor
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Ret_Ano_Veic_Cons_InMetro(WBusca As String, WChave As Long) As String
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Dim WValor As String
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Consulta_Selagem_Completa", DB_OPEN_TABLE)

If WChave = 1 Then
   tb.index = "Placa"
Else
   tb.index = "Renavam"
End If

tb.Seek "=", WBusca
If tb.NoMatch Then
   GoSub Finaliza:
Else
   If Not IsNull(tb!AnoVeiculo) Then
      WValor = tb!AnoVeiculo
   Else
      WValor = "INDEF"
   End If
End If

Finaliza:
  Ret_Ano_Veic_Cons_InMetro = WValor
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Ret_Mod_Veic_Cons_InMetro(WBusca As String, WChave As Long) As String
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Dim WValor As String
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Consulta_Selagem_Completa", DB_OPEN_TABLE)

If WChave = 1 Then
   tb.index = "Placa"
Else
   tb.index = "Renavam"
End If

tb.Seek "=", WBusca
If tb.NoMatch Then
   GoSub Finaliza:
Else
   If Not IsNull(tb!ModeloVeiculo) Then
      WValor = tb!ModeloVeiculo
   Else
      WValor = "INDEF"
   End If
End If

Finaliza:
  Ret_Mod_Veic_Cons_InMetro = WValor
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Ret_Marc_Veic_Cons_InMetro(WBusca As String, WChave As Long) As String
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Dim WValor As String
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Consulta_Selagem_Completa", DB_OPEN_TABLE)

If WChave = 1 Then
   tb.index = "Placa"
Else
   tb.index = "Renavam"
End If

tb.Seek "=", WBusca
If tb.NoMatch Then
   GoSub Finaliza:
Else
   If Not IsNull(tb!MarcaVeiculo) Then
      WValor = tb!MarcaVeiculo
   Else
      WValor = "Indef"
   End If
End If

Finaliza:
  Ret_Marc_Veic_Cons_InMetro = WValor
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Ret_Data_Selagem_Cons_InMetro(WBusca As String, WChave As Long) As String
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Dim WValor As String
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Consulta_Selagem_Completa", DB_OPEN_TABLE)

If WChave = 1 Then
   tb.index = "Placa"
Else
   tb.index = "Renavam"
End If

tb.Seek "=", WBusca
If tb.NoMatch Then
   GoSub Finaliza:
Else
   If Not IsNull(tb!DataSelagem) Then
      WValor = tb!DataSelagem
   Else
      WValor = " "
   End If
End If

Finaliza:
  Ret_Data_Selagem_Cons_InMetro = WValor
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Ret_Tipo_Certif(WBusca As String, WChave As Long) As String
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Dim WValor As String
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Consulta_Selagem_Completa", DB_OPEN_TABLE)

If WChave = 1 Then
   tb.index = "Placa"
Else
   tb.index = "Renavam"
End If

tb.Seek "=", WBusca
If tb.NoMatch Then
   GoSub Finaliza:
Else
   If Not IsNull(tb!TipoCertificadoResultado) Then
      WValor = tb!TipoCertificadoResultado
   Else
      WValor = " "
   End If
End If

Finaliza:
  Ret_Tipo_Certif = WValor
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function
Function Ret_Data_ValidadeCert(WBusca As String, WChave As Long) As String
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Dim WValor As String
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Consulta_Selagem_Completa", DB_OPEN_TABLE)

If WChave = 1 Then
   tb.index = "Placa"
Else
   tb.index = "Renavam"
End If

tb.Seek "=", WBusca
If tb.NoMatch Then
   GoSub Finaliza:
Else
   If Not IsNull(tb!DataValidadeResultado) Then
      WValor = tb!DataValidadeResultado
   Else
      WValor = " "
   End If
End If

Finaliza:
  Ret_Data_ValidadeCert = WValor
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Ret_Data_Ensaio(WBusca As String, WChave As Long) As String
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Dim WValor As String
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Consulta_Selagem_Completa", DB_OPEN_TABLE)

If WChave = 1 Then
   tb.index = "Placa"
Else
   tb.index = "Renavam"
End If

tb.Seek "=", WBusca
If tb.NoMatch Then
   GoSub Finaliza:
Else
   If Not IsNull(tb!DataEnsaio) Then
      WValor = tb!DataEnsaio
   Else
      WValor = " "
   End If
End If

Finaliza:
  Ret_Data_Ensaio = WValor
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Ret_Posto_Selagem(WBusca As String, WChave As Long) As String

On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Dim WValor As String
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Consulta_Selagem_Completa", DB_OPEN_TABLE)

If WChave = 1 Then
   tb.index = "Placa"
Else
   tb.index = "Renavam"
End If

tb.Seek "=", WBusca
If tb.NoMatch Then
   GoSub Finaliza:
Else
   If Not IsNull(tb!PostoSelagem) Then
      WValor = tb!PostoSelagem
   Else
      WValor = " "
   End If
End If

Finaliza:
  Ret_Posto_Selagem = WValor
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function
Function Ret_Num_Protoc_Selagem(WBusca As String, WChave As Long) As String
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Dim WValor As String
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Consulta_Selagem_Completa", DB_OPEN_TABLE)

If WChave = 1 Then
   tb.index = "Placa"
Else
   tb.index = "Renavam"
End If

tb.Seek "=", WBusca
If tb.NoMatch Then
   GoSub Finaliza:
Else
   If Not IsNull(tb!NumeroProtocoloSelagem) Then
      WValor = tb!NumeroProtocoloSelagem
   Else
      WValor = " "
   End If
End If

Finaliza:
  Ret_Num_Protoc_Selagem = WValor
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Ret_Chassi_Veic_Cons_InMetro(WBusca As String, WChave As Long) As String
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Dim WValor As String
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Consulta_Selagem_Completa", DB_OPEN_TABLE)

If WChave = 1 Then
   tb.index = "Placa"
Else
   tb.index = "Renavam"
End If

tb.Seek "=", WBusca
If tb.NoMatch Then
   GoSub Finaliza:
Else
   If Not IsNull(tb!CHASSI) Then
      WValor = tb!CHASSI
   Else
      WValor = "INDEF"
   End If
End If

Finaliza:
  Ret_Chassi_Veic_Cons_InMetro = WValor
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Ret_Gru_Veic_Cons_InMetro(WBusca As String, WChave As Long) As String
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Dim WValor As String
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Consulta_Selagem_Completa", DB_OPEN_TABLE)

If WChave = 1 Then
   tb.index = "Placa"
Else
   tb.index = "Renavam"
End If

tb.Seek "=", WBusca
If tb.NoMatch Then
   GoSub Finaliza:
Else
   If Not IsNull(tb!GRU) Then
      WValor = tb!GRU
   Else
      WValor = "INDEF"
   End If
End If

Finaliza:
  Ret_Gru_Veic_Cons_InMetro = WValor
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Ret_ValorGru_Veic_Cons_InMetro(WBusca As String, WChave As Long) As String
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Dim WValor As String
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Consulta_Selagem_Completa", DB_OPEN_TABLE)

If WChave = 1 Then
   tb.index = "Placa"
Else
   tb.index = "Renavam"
End If

tb.Seek "=", WBusca
If tb.NoMatch Then
   GoSub Finaliza:
Else
   If Not IsNull(tb!ValorGru) Then
      WValor = tb!ValorGru
   Else
      WValor = "0.00"
   End If
End If

Finaliza:
  Ret_ValorGru_Veic_Cons_InMetro = WValor
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Ret_ValorPagoGru_Veic_Cons_InMetro(WBusca As String, WChave As Long) As String
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Dim WValor As String
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Consulta_Selagem_Completa", DB_OPEN_TABLE)

If WChave = 1 Then
   tb.index = "Placa"
Else
   tb.index = "Renavam"
End If

tb.Seek "=", WBusca
If tb.NoMatch Then
   GoSub Finaliza:
Else
   If Not IsNull(tb!ValorPago) Then
      WValor = tb!ValorPago
   Else
      WValor = "0.00"
   End If
End If

Finaliza:
  Ret_ValorPagoGru_Veic_Cons_InMetro = WValor
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function
Function Ret_DataPagoGru_Veic_Cons_InMetro(WBusca As String, WChave As Long) As String
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Dim WValor As String
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Consulta_Selagem_Completa", DB_OPEN_TABLE)

If WChave = 1 Then
   tb.index = "Placa"
Else
   tb.index = "Renavam"
End If

tb.Seek "=", WBusca
If tb.NoMatch Then
   GoSub Finaliza:
Else
   If Not IsNull(tb!DataPagamento) Then
      WValor = tb!DataPagamento
   Else
      WValor = "0.00"
   End If
End If

Finaliza:
  Ret_DataPagoGru_Veic_Cons_InMetro = WValor
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function


Function Inicializa_Placa_Veic_Cons_Tabela_Completa(wPlaca As String) As Long
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Dim WValor As String
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Consulta_Selagem_Completa", DB_OPEN_TABLE)

tb.index = "Placa"
BeginTrans
tb.Seek "=", wPlaca
If Not tb.NoMatch Then
   While tb!Placa = wPlaca
         tb.Delete
         tb.MoveNext
         If tb.EOF Then
            GoSub Finaliza:
         End If
   Wend
End If

Finaliza:
  CommitTrans
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Ret_Nro_Equip_Cons_InMetro(WBusca As String, WChave As Long) As String
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Dim WValor As String
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Consulta_Selagem_Completa", DB_OPEN_TABLE)

If WChave = 1 Then
   tb.index = "Placa"
Else
   tb.index = "Renavam"
End If

tb.Seek "=", WBusca
If tb.NoMatch Then
   GoSub Finaliza:
Else
   If Not IsNull(tb!NroSerieEquip) Then
      WValor = tb!NroSerieEquip
   Else
      WValor = "INDEF"
   End If
End If

Finaliza:
  Ret_Nro_Equip_Cons_InMetro = WValor
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Ret_Redutor_Equip_Cons_InMetro(WBusca As String, WChave As Long) As String
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Dim WValor As String
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Consulta_Selagem_Completa", DB_OPEN_TABLE)

If WChave = 1 Then
   tb.index = "Placa"
Else
   tb.index = "Renavam"
End If

tb.Seek "=", WBusca
If tb.NoMatch Then
   GoSub Finaliza:
Else
   If Not IsNull(tb!Redutor) Then
      WValor = tb!Redutor
   Else
      WValor = "INDEF"
   End If
End If

Finaliza:
  Ret_Redutor_Equip_Cons_InMetro = WValor
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function
Function Ret_ConstK_Equip_Cons_InMetro(WBusca As String, WChave As Long) As String
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Dim WValor As String
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Consulta_Selagem_Completa", DB_OPEN_TABLE)

If WChave = 1 Then
   tb.index = "Placa"
Else
   tb.index = "Renavam"
End If

tb.Seek "=", WBusca
If tb.NoMatch Then
   GoSub Finaliza:
Else
   If Not IsNull(tb!ConstK) Then
      WValor = tb!ConstK
   Else
      WValor = "INDEF"
   End If
End If

Finaliza:
  Ret_ConstK_Equip_Cons_InMetro = WValor
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function GeraOcorrenciaSelos(WNUMOC As Long, WDataOC As Variant, WTipoOC As String, WTipoSel As String, WObserv As String, WNumos As Long, WNumero As String, wdv As String) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tbs, tbp As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())

Set tbp = bd.OpenRecordset("SELOSINUTILIZADOSPRINC", DB_OPEN_TABLE)
Set tbs = bd.OpenRecordset("SELOSINUTILIZADOS", DB_OPEN_TABLE)
tbs.index = "PrimaryKey"
tbp.index = "PrimaryKey"
BeginTrans

GoSub GeraOcorP:
GoSub GeraOcorS:

Finaliza:
  CommitTrans
  tbp.Close
  tbs.Close
  bd.Close
  Set tbp = Nothing
  Set tbs = Nothing
  Set bd = Nothing
  Exit Function
Return

GeraOcorS:
  tbs.Seek "=", WNUMOC, WDataOC, WTipoOC, WTipoSel, WNumero
  If tbs.NoMatch Then
     tbs.AddNew
  Else
     tbs.Edit
  End If
  tbs!NumOC = WNUMOC
  tbs!DataOC = WDataOC
  tbs!TipoOC = WTipoOC
  tbs!TipoSel = WTipoSel
  tbs!Numero = WNumero
  tbs!DV = wdv
  tbs.Update
Return

GeraOcorP:
  tbp.Seek "=", WNUMOC, WDataOC, WTipoOC, WTipoSel
  If tbp.NoMatch Then
     tbp.AddNew
  Else
     tbp.Edit
  End If
  tbp!NumOC = WNUMOC
  tbp!DataOC = WDataOC
  tbp!TipoOC = WTipoOC
  tbp!TipoSel = WTipoSel
  tbp!Observacao = WObserv
  tbp!Numos = WNumos
  tbp.Update
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function


Function Gera_Tabela_Relatorio_Mensal_InMetro_MG(WClassif As Long) As Long

On Error GoTo TratamentodeErro:
Dim bd, bdt As DataBase
Dim tb, tba As Recordset
Dim Qtb As QueryDef, numreg As Long, Lote, X, Y As Long, Ret As Variant
Set bd = CurrentDb()
Set bdt = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
If WClassif = 1 Then
   Set Qtb = bd.QueryDefs("CadastraOSAssistTecnicaEnsaiosRel_MOD06")
Else
   Set Qtb = bd.QueryDefs("CadastraOSAssistTecnicaEnsaiosRel_MOD06_Data")
End If
Qtb.Parameters(0) = Forms!DialogoSelecaoMesRelInMetro!DataInicial
Qtb.Parameters(1) = Forms!DialogoSelecaoMesRelInMetro!DataFinal
Set tb = Qtb.OpenRecordset()
Set tba = bdt.OpenRecordset("Tabela_Temp_Rel_Mensal_InMetro_MG", DB_OPEN_TABLE)
tba.index = "PrimaryKey"

X = 1
Lote = 1
Y = 1

If tba.RecordCount <> 0 Then
   BeginTrans
   While Not tba.EOF
         tba.Delete
         tba.MoveNext
   Wend
   CommitTrans
End If

If tb.RecordCount <> 0 Then
   BeginTrans
   While Not tb.EOF
         tba.Seek "=", Lote, X
         If tba.NoMatch Then
            tba.AddNew
         Else
            tba.Edit
         End If
         tba!Sequencia = X
         tba!Lote = Lote
         tba!Placa = tb!PlacaVeiculo
         tba!Renavam = tb!RENAVAN
         tba!DTEnsaio = tb!DTEnsaio
         tba!GRU = tb!GRU_Ensaio
         tba.Update
         X = X + 1
         
         If X > 30 Then
            X = 1
            Lote = Lote + 1
         End If
         
         Y = Y + 1
         If Y Mod 200 = 0 Then
            CommitTrans
            BeginTrans
         End If
         tb.MoveNext
   
   Wend
   CommitTrans
Else
   MsgBox "Nenhum registro selecionado...", vbCritical, "SGa"
End If

Finaliza:
'   CommitTrans
   tb.Close
   tba.Close
   bd.Close
   bdt.Close
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function Inicializa_Tabela_Temp_Rel_Mensal_InMetro_MG()

On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Dim numreg As Long, Lote, X, Y As Long, Ret As Variant
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Tabela_Temp_Rel_Mensal_InMetro_MG", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

If tb.RecordCount <> 0 Then
   BeginTrans
   While Not tb.EOF
         tb.Delete
         tb.MoveNext
   Wend
   CommitTrans
End If

Finaliza:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
    End Function

Function Elimina_Placas_Invalidas() As Long
On Error GoTo TratamentodeErro:
    
Dim bd As DataBase, tb As Recordset
Dim X As Long
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("PLACXCLI", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

BeginTrans

If tb.RecordCount <> 0 Then
   While Not tb.EOF
         If Len(tb!Placa) <> 7 Or IsNull(tb!Placa) Or IsNull(EliminaEspacos(tb!Placa)) Then
            tb.Delete
         End If
         tb.MoveNext
         X = X + 1
         If X Mod 200 = 0 Then
            CommitTrans
            BeginTrans
        End If
   Wend
End If
    
Finaliza:
    CommitTrans
    tb.Close
    bd.Close
    Set tb = Nothing
    Set bd = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    ElseIf Err.Number = 13 Then
           GoSub Finaliza:
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
      Resume Next
    End If

End Function

Function Retorna_Dados_Placas_Consultadas() As String
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb, tbc As Recordset
Dim qtbc As QueryDef
Dim X As Long
Dim WString As String

Set bd = CurrentDb()
Set qtbc = bd.QueryDefs("Consulta_Plac_Cli_Conec_Retorna_Dados_Placas_Selec")
qtbc.Parameters(0) = Forms!DialogoSelecaoPlacXCliXPrinc!Cliente
qtbc.Parameters(1) = Forms!DialogoSelecaoPlacXCliXPrinc!Cidade
qtbc.Parameters(2) = Forms!DialogoSelecaoPlacXCliXPrinc!UF
qtbc.Parameters(3) = Forms!DialogoSelecaoPlacXCliXPrinc!Placa
Set tbc = qtbc.OpenRecordset()

If tbc.RecordCount <> 0 Then
   While Not tbc.EOF
         WString = WString & "(" & UCase(tbc!Placa) & " - " & tbc!Data_Validade_Resultado & ") "
         tbc.MoveNext
    Wend
End If

Finaliza:
    Retorna_Dados_Placas_Consultadas = WString
    tbc.Close
    bd.Close
    Set tbc = Nothing
    Set bd = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
       Resume
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume
    End If

End Function

Function Valida_Qtde_Cli_Consulta_Placas_InMetro(WFlag As Long) As String
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb, tbc As Recordset
Dim qtbc As QueryDef
Dim X As Long

Set bd = CurrentDb()
Set qtbc = bd.QueryDefs("Consulta_Plac_Cli_Conec_Valida_Qtde_Cli")
qtbc.Parameters(0) = Forms!DialogoSelecaoPlacXCliXPrinc!Cliente
qtbc.Parameters(1) = Forms!DialogoSelecaoPlacXCliXPrinc!Cidade
qtbc.Parameters(2) = Forms!DialogoSelecaoPlacXCliXPrinc!UF
qtbc.Parameters(3) = Forms!DialogoSelecaoPlacXCliXPrinc!Placa
qtbc.Parameters(4) = Forms!DialogoSelecaoPlacXCliXPrinc!VenctoI
qtbc.Parameters(5) = Forms!DialogoSelecaoPlacXCliXPrinc!VenctoF
qtbc.Parameters(6) = Forms!DialogoSelecaoPlacXCliXPrinc!ConsultaI
qtbc.Parameters(7) = Forms!DialogoSelecaoPlacXCliXPrinc!ConsultaF
Set tbc = qtbc.OpenRecordset()

If tbc.RecordCount = 0 Then
   MsgBox "Nenhum registro selecionado...", vbCritical, "SGa"
   Valida_Qtde_Cli_Consulta_Placas_InMetro = 0
ElseIf tbc.RecordCount > 1 Then
       MsgBox "Numero de clientes selecionados excede quantidade suportada para envio de email...", vbCritical, "SGa"
       Valida_Qtde_Cli_Consulta_Placas_InMetro = 9
ElseIf tbc.RecordCount = 1 Then
       If WFlag = 0 Then
          Valida_Qtde_Cli_Consulta_Placas_InMetro = 1
       ElseIf WFlag = 1 Then
              While Not tbc.EOF
                   If IsNull(tbc!EndElet) Then
                      Valida_Qtde_Cli_Consulta_Placas_InMetro = "*"
                   Else
                       Valida_Qtde_Cli_Consulta_Placas_InMetro = tbc!EndElet
                   End If
                   tbc.MoveNext
              Wend
       ElseIf WFlag = 2 Then
              While Not tbc.EOF
                   Valida_Qtde_Cli_Consulta_Placas_InMetro = tbc!EndElet
                   tbc.MoveNext
              Wend
       
       
       End If
End If

Finaliza:
    tbc.Close
    bd.Close
    Set tbc = Nothing
    Set bd = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
       Resume
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume
    End If

End Function

Function Cria_Base_Consulta_Lote() As Boolean

On Error GoTo TratamentodeErro:
Dim bd, bdc As DataBase
Dim tb, tbc As Recordset
Dim qtbc As QueryDef
Dim X As Long
Set bdc = CurrentDb()
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Base_Dados_Lote", DB_OPEN_TABLE)
Set qtbc = bdc.QueryDefs("Consulta_Plac_Cli_Conec")
qtbc.Parameters(0) = Forms!DialogoSelecaoPlacXCliXPrinc!Cliente
qtbc.Parameters(1) = Forms!DialogoSelecaoPlacXCliXPrinc!Cidade
qtbc.Parameters(2) = Forms!DialogoSelecaoPlacXCliXPrinc!UF
qtbc.Parameters(3) = Forms!DialogoSelecaoPlacXCliXPrinc!Placa
qtbc.Parameters(4) = Forms!DialogoSelecaoPlacXCliXPrinc!VenctoI
qtbc.Parameters(5) = Forms!DialogoSelecaoPlacXCliXPrinc!VenctoF
qtbc.Parameters(6) = Forms!DialogoSelecaoPlacXCliXPrinc!ConsultaI
qtbc.Parameters(7) = Forms!DialogoSelecaoPlacXCliXPrinc!ConsultaF
qtbc.Parameters(8) = Forms!DialogoSelecaoPlacXCliXPrinc!TipoCertif
Set tbc = qtbc.OpenRecordset()
tb.index = "PrimaryKey"

Cria_Base_Consulta_Lote = True

GoSub Inicializa_Base_Lote:
GoSub Atualiza_Base_Consulta:

Finaliza:
    tb.Close
    tbc.Close
    bd.Close
    Set tb = Nothing
    Set tbc = Nothing
    Set bd = Nothing
    Exit Function
Return

Atualiza_Base_Consulta:
    BeginTrans
    If tbc.RecordCount <> 0 Then
       While Not tbc.EOF
             If tbc!Selecionar = True Then
                X = 1
                tb.Seek "=", tbc!Placa, tbc!CodigoCliente
                If tb.NoMatch Then
                    tb.AddNew
                Else
                    tb.Edit
                End If
                tb!Placa = tbc!Placa
                tb!CodigoCliente = tbc!CodigoCliente
                tb!CodEmpresa = Form_Menu.WEmpresa
                tb.Enviar = False
                tb.Status = Null
                tb.Update
             End If
             tbc.MoveNext
             X = X + 1
             If X Mod 200 = 0 Then
                CommitTrans
                BeginTrans
            End If
            If tbc.EOF Then
                CommitTrans
                GoSub Finaliza:
            End If
       Wend
    End If
    CommitTrans
Return

Inicializa_Base_Lote:
    BeginTrans
    If tb.RecordCount <> 0 Then
       While Not tb.EOF
                tb.Delete

             X = X + 1
             If X Mod 200 = 0 Then
                CommitTrans
                BeginTrans
             End If
             tb.MoveNext
       Wend
    End If
    CommitTrans
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
       Resume
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume
    End If



End Function

Function Atualiza_Data_Consulta_Validade_Certificado() As Long
On Error GoTo TratamentodeErro:
    
Dim bd As DataBase
Dim tb As Recordset
Dim X As Long
    
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("PLACXCLI", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

BeginTrans
If tb.RecordCount <> 0 Then
   X = 0
   While Not tb.EOF
         tb.Edit
         If IsNull(tb!DataConsulta) Or tb!DataConsulta = "" Or IsEmpty(tb!DataConsulta) Then
            tb!Data_Ultima_Consulta = #1/1/2001#
         Else
            tb!Data_Ultima_Consulta = IIf(IsDate(tb!DataConsulta), tb!DataConsulta, #1/1/2001#)
         End If
                  
         If IsNull(tb!DataValidadeResultado) Or tb!DataValidadeResultado = "" Or IsEmpty(tb!DataValidadeResultado) Then
            tb!Data_Validade_Resultado = #1/1/2001#
         Else
            If Mid$(tb!DataValidadeResultado, 1, 1) <> "-" Then
                tb!Data_Validade_Resultado = IIf(IsDate(Mid$(tb!DataValidadeResultado, 1, 10)), Mid$(tb!DataValidadeResultado, 1, 10), #1/1/2001#)
            End If
         End If
         tb.Update
         X = X + 1
         If X Mod 200 = 0 Then
            CommitTrans
            BeginTrans
         End If
         tb.MoveNext
   Wend
End If
    
Finaliza:
    CommitTrans
    tb.Close
    bd.Close
    Set tb = Nothing
    Set bd = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    ElseIf Err.Number = 13 Then
           GoSub Finaliza:
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
      Resume Next
    End If



End Function

Function Atualiza_GRU_Ensaio_OSV(WNumos As Long, WGRU_Ensaio As String) As Long

On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

BeginTrans
tb.Seek "=", WNumos
If Not tb.NoMatch Then
   tb.Edit
   tb!GRU_Ensaio = WGRU_Ensaio
   tb.Update
End If

Finaliza:
    CommitTrans
    tb.Close
    bd.Close
    Set tb = Nothing
    Set bd = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function Efetua_Ajuste_GRU_Selagem(WGRUSelagem As String, WQuantidade As Long, WOperacao As String) As Boolean
On Error GoTo TratamentodeErro
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("SELAGENS_SUB", DB_OPEN_TABLE)
tb.index = "GRU_Selagem"

BeginTrans
tb.Seek "=", WGRUSelagem
If Not tb.NoMatch Then
   
   tb.Edit
   If WOperacao = "A" Then
      tb!Quantidade_Selagens = tb!Quantidade_Selagens + WQuantidade
       Efetua_Ajuste_GRU_Selagem = True
        tb.Update
       GoSub Finaliza:
   ElseIf WOperacao = "D" Then
          If WQuantidade <= tb!Quantidade_Selagens Then
             tb!Quantidade_Selagens = tb!Quantidade_Selagens - WQuantidade
             Efetua_Ajuste_GRU_Selagem = True
              tb.Update
             GoSub Finaliza:
          Else
             MsgBox "Quantidade informada superior a quantidade encontrada GRU, impossivel prosseguir...", vbCritical, "SGa"
             Efetua_Ajuste_GRU_Selagem = False
              tb.Update
             GoSub Finaliza:
          End If
   End If
   tb.Update
End If

Finaliza:
    CommitTrans
    Call LogSGa("GRUSELAGEM", "ADMSELAGEMPRINC", WOperacao, Captura_Conteudo_Campo_RESP(Usuario_Ativo(), "nome"), Format$(Date, "dd/mm/yyyy"), Time(), "gru " & WGRUSelagem, Environ("ComputerName"))
  '  Time(), "Tit " & Val(wnronf) & " " & Format$(WDataEmissao, "dd/mm/yyyy"), Environ("ComputerName"))
    tb.Close
    bd.Close
    Set tb = Nothing
    Set bd = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function Retorna_Links_Uteis(WCodigo As Long) As String
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("Links_Uteis", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

tb.Seek "=", WCodigo
If Not tb.NoMatch Then
   If Not IsNull(tb!Descricao_Link) Then
      Retorna_Links_Uteis = tb!Descricao_Link
   Else
      MsgBox "Descricao do link invalida...", vbCritical, "SGa"
      Retorna_Links_Uteis = "#"
   End If

End If

Finaliza:
    tb.Close
    bd.Close
    Set tb = Nothing
    Set bd = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function



Function Valida_Controlador_Repasse(WCodigoServico As Long) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb, tba As Recordset
Dim WValida As Long
Dim X As Long
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tba = bd.OpenRecordset("EMPXTPSERVXREPASSE", DB_OPEN_TABLE)
tba.index = "PrimaryKey"

tba.Seek "=", 1, WCodigoServico
If Not tba.NoMatch Then
   If tba!controlador = 0 Or IsNull(tba!controlador) Then
      MsgBox "Ausencia de Controlador na Tabela de Repasse, Impossivel Prosseguir", vbCritical, "SGA"
      WValida = 1  ' CONTROLADODOR INVALIDO
   ElseIf tba!Repasse = 0 Or IsNull(tba!Repasse) Then
      MsgBox "Ausencia de Aliquota de Repasse, Impossivel Prosseguir", vbCritical, "SGA"
      WValida = 1  ' ALIQUOTA DE REPASSE INVALIDO
   ElseIf tba!CentroCusto = 0 Or IsNull(tba!CentroCusto) Then
      MsgBox "Ausencia de Centro de Custo na Tabela de Repasse, Impossivel Prosseguir", vbCritical, "SGA"
      WValida = 1  ' CENTRO DE CUSTO INVALIDO
   ElseIf tba!NaturezaGasto = 0 Or IsNull(tba!NaturezaGasto) Then
      MsgBox "Ausencia de Natureza de Gasto na Tabela de Repasse, Impossivel Prosseguir", vbCritical, "SGA"
      WValida = 1  ' NATUREZA DE GASTO
   ElseIf tba!CodigoPagto = 0 Or IsNull(tba!CodigoPagto) Then
      MsgBox "Ausencia de Operacao na Tabela de Repasse, Impossivel Prosseguir", vbCritical, "SGA"
      WValida = 1  ' CODIGO PAGTO
   Else
      WValida = 0
   End If
End If

Finaliza:
    Valida_Controlador_Repasse = WValida
    tba.Close
    bd.Close
    Set tba = Nothing
    Set bd = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function


Function Saldo_Disp_Selagens_GRU(WGRU As String) As Long
On Error GoTo TratamentodeErro:
Dim bd, bdc As DataBase
Dim tb, tba As Recordset
Dim WQuant_S As Long
Dim WQuant_E As Long
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("Selagens_Sub", DB_OPEN_TABLE)
Set tba = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "GRU_Selagem"
tba.index = "GRU_Selagem"

GoSub Apura_Quant_Selagens_Inseridas:
GoSub Apura_Quant_Selagens_Efetuadas:

Saldo_Disp_Selagens_GRU = (WQuant_E - WQuant_S)

Finaliza:
    tb.Close
    tba.Close
    bd.Close
    Set tb = Nothing
    Set tba = Nothing
    Set bd = Nothing
    Exit Function
Return

Apura_Quant_Selagens_Efetuadas:
    tba.Seek "=", WGRU
    If Not tba.NoMatch Then
       While Not tba.EOF
             If tba!GRU_Selagem <> WGRU Then
                Return
             End If
             If tba!TipoLaudo <> "E" Then
                WQuant_S = WQuant_S + 1
             End If
             tba.MoveNext
             If tba.EOF Then
                Return
             End If
       Wend
    End If
Return

Apura_Quant_Selagens_Inseridas:
    tb.Seek "=", WGRU
    If Not tb.NoMatch Then
       WQuant_E = tb!Quantidade_Selagens
    End If
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function Saldo_Disp_Selagens(WData As Variant) As Long
On Error GoTo TratamentodeErro:
Dim bd, bdc As DataBase
Dim tb, tba, tbc As Recordset
Dim qtbc As QueryDef, numreg As Long, X As Long, Ret As Variant
Dim WQuant_S As Long
Dim WQuant_E As Long
Set bdc = CurrentDb()
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set qtbc = bdc.QueryDefs("Query_OSV_Selagens_Efetuadas")
Set tbc = qtbc.OpenRecordset()
Set tb = bd.OpenRecordset("Selagens_Sub", DB_OPEN_TABLE)
Set tba = bd.OpenRecordset("Selagens", DB_OPEN_TABLE)
tb.index = "Data_Aquisicao"
tba.index = "GRU_Selagem"

GoSub Apura_Quant_Selagens_Efetuadas:
GoSub Apura_Quant_Selagens_Inseridas:

Saldo_Disp_Selagens = (WQuant_E - WQuant_S)

Finaliza:
    tb.Close
    tbc.Close
    tba.Close
    bdc.Close
    bd.Close
    Set tb = Nothing
    Set tbc = Nothing
    Set tba = Nothing
    Set bdc = Nothing
    Set bd = Nothing
    Exit Function
Return

Apura_Quant_Selagens_Inseridas:
    If tb.RecordCount <> 0 Then
       While Not tb.EOF
             If tb!Data_Aquisicao > WData Then
                Return
             End If
             WQuant_E = WQuant_E + tb!Quantidade_Selagens
             tb.MoveNext
             If tb.EOF Then
                Return
             End If
       Wend
    End If
Return

Apura_Quant_Selagens_Efetuadas:
    If tbc.RecordCount <> 0 Then
       While Not tbc.EOF
             If tbc!datafechamento > WData Then
                Return
             End If
             tba.Seek "=", tbc!GRU_Selagem
             If Not tba.NoMatch Then
                WQuant_S = WQuant_S + 1
             End If
             tbc.MoveNext
             If tbc.EOF Then
                Return
             End If
       Wend
    End If
Return



TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Valida_Plac_Cli_Conec(WCliente As String, WCidade As String, WUF As String, wPlaca As String, wVenctoI As Variant, wVenctoF As Variant, WConsultaI As Variant, WConsultaF As Variant, WtipoCertificado As String, WSegmento As String) As Boolean

On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Dim Qtb As QueryDef, numreg As Long, X As Long, Ret As Variant
Set bd = CurrentDb()
Set Qtb = bd.QueryDefs("Plac_Cli_Conec")
Qtb.Parameters(0) = WCliente
Qtb.Parameters(1) = WCidade
Qtb.Parameters(2) = WUF
Qtb.Parameters(3) = wPlaca
Qtb.Parameters(4) = WtipoCertificado
Qtb.Parameters(5) = wVenctoI
Qtb.Parameters(6) = wVenctoF
Qtb.Parameters(7) = WConsultaI
Qtb.Parameters(8) = WConsultaF
Qtb.Parameters(9) = WSegmento


Set tb = Qtb.OpenRecordset()

If tb.RecordCount <> 0 Then
   Valida_Plac_Cli_Conec = True
Else
   Valida_Plac_Cli_Conec = False
End If

tb.Close
bd.Close
Set tb = Nothing
Set bd = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function


Function Valida_Consulta_Lote_Placas() As Boolean
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Dim Qtb As QueryDef, numreg As Long, X As Long, Ret As Variant
Set bd = CurrentDb()
Set Qtb = bd.QueryDefs("ConsultaLotePlacasXInMetro")
Set tb = Qtb.OpenRecordset()

If tb.RecordCount <> 0 Then
   Valida_Consulta_Lote_Placas = True
Else
   Valida_Consulta_Lote_Placas = False
End If

tb.Close
bd.Close
Set tb = Nothing
Set bd = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Retorna_Data_Validade_PlacxCli(WData) As Variant

If WData = "" Or IsNull(WData) Or WData = "*" Or WData = "#" Or WData = Space(1) Then
   Retorna_Data_Validade_PlacxCli = #1/1/2001#
Else
   Retorna_Data_Validade_PlacxCli = WData
End If

End Function

Function Relatorio_Mensal(WClassif As String, WDataInicial As Variant, WDataFinal As Variant) As Long
DoCmd.Hourglass True
On Error GoTo TratamentodeErro:

Dim Qtb As QueryDef
Dim tb, tbc As Recordset
Dim bd, bdt As DataBase
Dim WIndex As Long
Dim WStr As String

Set bd = CurrentDb()
Set bdt = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
If WClassif = "1" Then
   Set Qtb = bd.QueryDefs("RelatorioMensalInMetro_Placa")
Else
   Set Qtb = bd.QueryDefs("RelatorioMensalInMetro_Data")
End If
Qtb.Parameters(0) = WDataInicial
Qtb.Parameters(1) = WDataFinal
Set tb = Qtb.OpenRecordset()
Set tbc = bdt.OpenRecordset("Ens_Sel_Mes", DB_OPEN_TABLE)
tbc.index = "Primarykey"
BeginTrans

WIndex = 0

GoSub InicializaTBC

If tb.RecordCount <> 0 Then
   While Not tb.EOF
         
         WIndex = WIndex + 1
         tbc.AddNew
         tbc!datafechamento = tb!datafechamento
         tbc!CodigoCliente = tb!CodigoCliente
         tbc!Numos = tb!Numos
         tbc!PlacaVeiculo = tb!PlacaVeiculo
         tbc!Marca = tb!Marca
         tbc!Modelo = tb!Modelo
         tbc!RazaoSocial = tb!RazaoSocial
         tbc!InMetro = InscInMetro()
         tbc!Email = Email()
         tbc!NumeroTacografo = tb!NumeroTacografo
         If tb!SelosPlastAplicados <> "" Then
            tbc!SelosPlastAplicados = tb!SelosPlastAplicados
         End If
         If tb!SelosAdesivostAplicados <> "" Then
            tbc!SelosAdesivosAplicados = tb!SelosAdesivostAplicados
         End If
         tbc!Oficina = Empresa()
         tbc!Endereco = Endereco()
         tbc!Estado = Estado
         tbc!Municipio = Cidade()
         tbc!Fone = Telefone()
         tbc!Site = Site()
         tbc!Mes = Format$(tb!datafechamento, "mmm")
         tbc!bAIr = Bairro()
         tbc!Cep = Cep()
         tbc!UF = Estado()
         tbc!Laudo = tb!Laudo
         tbc!CoeficienteW = tb!CoeficienteW
         tbc!ConstanteK = tb!ConstanteK
         tbc!ModeloVeiculo = tb!ModeloVeiculo
         tbc!MarcaVeiculo = tb!MarcaVeiculo
         If TipoPosto() = "E" Then
            tbc!GRUR = tb!GRU_Ensaio
         Else
            tbc!GRUR = tb!GRU_Selagem
         End If
         tbc!AnoVeiculo = tb!AnoVeiculo
         tbc!RENAVAN = tb!RENAVAN
         tbc!CHASSI = tb!CHASSI
         tbc!Pneu = tb!Pneu
         tbc!DTEnsaio = tb!DTEnsaio
         tbc!Nro = NroEnd()
         tbc!GRU_Ensaio = tb!GRU_Ensaio
         tbc!GRU_Selagem = tb!GRU_Selagem
         tbc!Protocolo_Selagem = tb!Protocolo_Selagem
         tbc!Sequencia = WIndex
         tbc.Update
         tb.MoveNext
   Wend
End If

Finaliza:
    DoCmd.Hourglass False
    CommitTrans
    tb.Close
    tbc.Close
    bd.Close
    bdt.Close
    Set tb = Nothing
    Set tbc = Nothing
    Set bd = Nothing
    Set bdt = Nothing
    Exit Function
Return

InicializaTBC:
   If tbc.RecordCount <> 0 Then
      While Not tbc.EOF
            tbc.Delete
            tbc.MoveNext
      Wend
   End If
Return


TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Public Function ValidaSelosAplicTempxSelosAplic(ByVal WNumos As Long) As Boolean
On Error GoTo TratamentodeErro:
Dim bdt As DataBase
Dim tbt As Recordset

Dim SeloInformado As Boolean

Set bdt = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tbt = bdt.OpenRecordset("SelosAplicadosTemp", DB_OPEN_TABLE)

SeloInformado = True

BeginTrans
tbt.index = "QuintaChave"
tbt.Seek "=", WNumos
If Not tbt.NoMatch Then
    While Not tbt.EOF
        If tbt.Fields("NumOs") = WNumos Then
            If Not IsNull(tbt.Fields("NumeroSubst")) And tbt.Fields("NumeroSubst") <> " " Then
                SeloInformado = True
                If ValidaSelosDetEntradasxSubXEns(tbt.Fields("NumeroSubst"), tbt.Fields("TipoSelo")) = 0 Then
                    MsgBox "Selo " & UCase(tbt.Fields("NumeroSubst")) & " informado inexistente na base de dados, impossivel prosseguir...", vbCritical, "Aviso"
                    ValidaSelosAplicTempxSelosAplic = False
                    GoTo Finaliza
                ElseIf ValidaSeloAplicadoXSubstXEns(tbt.Fields("NumeroSubst"), tbt.Fields("TipoSelo")) <> 0 Then
                    MsgBox "Selo  " & UCase(tbt.Fields("NumeroSubst")) & " indisponivel para aplica��o, utiliza��o efetuada na OSV " & Format$(ValidaSeloAplicadoXSubstXEns(tbt.Fields("NumeroSubst"), tbt.Fields("TipoSelo")), "000,000") & " impossivel prosseguir...", vbCritical, "Aviso"
                    ValidaSelosAplicTempxSelosAplic = False
                    GoTo Finaliza
                ElseIf ValidaSeloInutilizadoXSubstXEns(tbt.Fields("NumeroSubst"), tbt.Fields("TipoSelo")) <> 0 Then
                    MsgBox "Selo inutilizado, impossivel prosseguir...", vbCritical, "Aviso"
                    ValidaSelosAplicTempxSelosAplic = False
                    GoTo Finaliza
                End If
            End If
            tbt.MoveNext
        Else
            ValidaSelosAplicTempxSelosAplic = SeloInformado
            GoTo Finaliza
        End If
    Wend
    ValidaSelosAplicTempxSelosAplic = SeloInformado
    GoTo Finaliza
    
Else
    MsgBox "Nenhum Selo Foi Informado Para Esta Operacao, Impossivel Prosseguir.", vbCritical
    ValidaSelosAplicTempxSelosAplic = False
End If

Finaliza:
   CommitTrans
   tbt.Close
   bdt.Close
   Set tbt = Nothing
   Set bdt = Nothing
   DoCmd.Close acForm, "CadastroSubstEnsaioPrinc", acSaveYes
   Exit Function
   
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
   
End Function

Function Valida_Modelo_Selo_Temp(WNumos As Long) As Boolean
On Error GoTo TratamentodeErro:
Dim tbt As Recordset
Dim bdt As DataBase
Set bdt = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tbt = bdt.OpenRecordset("SelosAplicadosTemp", DB_OPEN_TABLE)
tbt.index = "QuintaChave"
 
tbt.Seek "=", WNumos
If tbt.NoMatch Then
   MsgBox "OSV nao existente...", vbCritical, "SGa"
   Valida_Modelo_Selo_Temp = False
Else
   Valida_Modelo_Selo_Temp = True
   While tbt!Numos = WNumos
         If tbt!TipoSelo = "P" And Not IsNull(tbt!NumeroSubst) Then
            If IsNull(tbt!ModeloSelo) Then
               MsgBox "Ausencia informa��o modelo selo " & tbt!Numero & " ,impossivel prosseguir..", vbCritical, "SGa"
               Valida_Modelo_Selo_Temp = False
               GoSub Finaliza:
            End If
         End If
         tbt.MoveNext
         If tbt.EOF Then
            GoSub Finaliza:
         End If
         
   Wend
End If

Finaliza:
   tbt.Close
   bdt.Close
   Set tbt = Nothing
   Set bdt = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function IncluiSelosAplicTempXSelosAplic(WNumos As Long) As Long

On Error GoTo TratamentodeErro:
Dim tb, tbt, tbr As Recordset
Dim bd, bdt As DataBase
Set bdt = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Dim WNumeroStr As String
Dim wDVStr As String
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tbt = bdt.OpenRecordset("SelosAplicadosTemp", DB_OPEN_TABLE)
Set tb = bd.OpenRecordset("SelosAplicados", DB_OPEN_TABLE)
Set tbr = bd.OpenRecordset("SelosRemovidos", DB_OPEN_TABLE)
tbt.index = "PrimaryKey"
tb.index = "PrimaryKey"
tbr.index = "PrimaryKey"

 
BeginTrans

If tbt.RecordCount <> 0 Then
    While Not tbt.EOF
          GoSub EfetuaInclusao:
          If tbt!NumeroSubst <> " " Then
             GoSub EfetuaInclusaoRemocao:
          End If
          tbt.MoveNext
     Wend
End If

Finaliza:
   CommitTrans
   tb.Close
   tbt.Close
   tbr.Close
   bd.Close
   bdt.Close
   Set tb = Nothing
   Set tbt = Nothing
   Set tbr = Nothing
   Set bd = Nothing
   Set bdt = Nothing
   MsgBox "Fim normal de execu��o", vbInformation, "SGa"
   DoCmd.Close acForm, "CadastroSubstEnsaioPrinc", acSaveYes
   Exit Function
Return

EfetuaInclusaoRemocao:
     tbr.Seek "=", Trim$(tbt!NumeroSubst & tbt!DVSUBST), WNumos, tbt!TipoSelo
     If tbr.NoMatch Then
        tbr.AddNew
     Else
        tbr.Edit
     End If
     tbr!Numero = Trim$(tbt!NumeroSubst & tbt!DVSUBST)
     tbr!Numos = WNumos
     tbr!TipoSelo = tbt!TipoSelo
     'TBR!DV = TBT!DVSubst
     tbr.Update
Return

EfetuaInclusao:
     tb.Seek "=", tbt!Numero, WNumos, tbt!TipoSelo
     If tb.NoMatch Then
        tb.AddNew
     Else
        tb.Edit
     End If
     tb!Numero = tbt!Numero
     tb!Numos = WNumos
     tb!TipoSelo = tbt!TipoSelo
     tb!DV = tbt!DV
     tb!NumeroSubst = tbt!NumeroSubst
     tb!DVSUBST = tbt!DVSUBST
     tb.Update
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function IncluiSelosAplicXSubstXEnsaios(WNumos As Long, WNumeroSelo As String, wdv As String, WNumeroSeloSt As String, WDVSt As String, WTipoSelo As String) As Long
On Error GoTo TratamentodeErro:
Dim tb As Recordset
Dim bd As DataBase
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("SelosAplicadosTemp", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
 
BeginTrans
tb.Seek "=", WNumeroSelo, WNumos, WTipoSelo
If tb.NoMatch Then
   tb.AddNew
Else
   tb.Edit
End If
tb!Numero = WNumeroSelo
tb!Numos = WNumos
tb!TipoSelo = WTipoSelo
tb!DV = wdv
tb!NumeroSubst = WNumeroSeloSt
tb!DVSUBST = WDVSt
tb.Update

Finaliza:
   CommitTrans
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ValidaSeloAplicadoXSubstXEnsTemp(WNumeroSelo As String, WTipoSelo As String) As Long
On Error GoTo TratamentodeErro:
Dim tb As Recordset
Dim bd As DataBase
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("SelosAplicadosTemp", DB_OPEN_TABLE)
tb.index = "QuartaChave"
 
tb.Seek "=", WNumeroSelo, WTipoSelo
If tb.NoMatch Then
   ValidaSeloAplicadoXSubstXEnsTemp = 0
Else
   ValidaSeloAplicadoXSubstXEnsTemp = tb!Numos
End If

Finaliza:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
   
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function InicializaTabSelosTemp() As Long
On Error GoTo TratamentodeErro:
Dim tb, tbr As Recordset
Dim bd As DataBase
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tbr = bd.OpenRecordset("SelosRemovidosTemp", DB_OPEN_TABLE)
Set tb = bd.OpenRecordset("SelosAplicadosTemp", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tbr.index = "PrimaryKey"
 
BeginTrans
If tb.RecordCount <> 0 Then
   While Not tb.EOF
         tb.Delete
         tb.MoveNext
   Wend
End If

If tbr.RecordCount <> 0 Then
   While Not tbr.EOF
         tbr.Delete
         tbr.MoveNext
   Wend
End If

Finaliza:
   CommitTrans
   tb.Close
   tbr.Close
   bd.Close
   Set tb = Nothing
   Set tbr = Nothing
   Set bd = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ApuraNumOSVxNumeroSeloRemovido(WNumero As String, WTipoSelo As String) As Long
On Error GoTo TratamentodeErro:
Dim SA, OS As Recordset
Dim bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set SA = bd1.OpenRecordset("SelosRemovidos", DB_OPEN_TABLE)
Set OS = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
SA.index = "SecondKey"
OS.index = "PrimaryKey"
 
'BeginTrans
SA.Seek "=", WNumero, WTipoSelo
If SA.NoMatch Then
 '  MsgBox "Selo inexistente, impossivel prosseguir...", vbCritical, "SGa"
   ApuraNumOSVxNumeroSeloRemovido = -999
   GoSub Finaliza:
Else
 
   GoSub ApuraOSV:
End If

Finaliza:
'   CommitTrans
   SA.Close
   OS.Close
   bd1.Close
   Set SA = Nothing
   Set OS = Nothing
   Set bd1 = Nothing
   Exit Function
Return


ApuraOSV:
   OS.Seek "=", SA!Numos
   If Not OS.NoMatch Then
      ApuraNumOSVxNumeroSeloRemovido = SA!Numos
   Else
      ApuraNumOSVxNumeroSeloRemovido = 0
   End If
Return


TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CriaLacrePlastTracado(WNumLAcre As String, wdv As String, WNumos As Long) As Long
On Error GoTo TratamentodeErro:
Dim SA As Recordset
Dim bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set SA = bd1.OpenRecordset("SelosAplicados", DB_OPEN_TABLE)
SA.index = "PrimaryKey"
 
BeginTrans
SA.Seek "=", WNumLAcre, WNumos, "P"
If SA.NoMatch Then
   SA.AddNew
Else
   SA.Edit
End If
SA!Numero = WNumLAcre
SA!DV = wdv
SA!Numos = WNumos
SA!TipoSelo = "P"
SA.Update

Finaliza:
   CriaLacrePlastTracado = 0
   CommitTrans
   SA.Close
   bd1.Close
   Set SA = Nothing
   Set bd1 = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function
 
Function ApuraNumOSVxNumeroSelo(WNumero As String, WTipoSelo As String) As Long
On Error GoTo TratamentodeErro:
Dim SA, OS As Recordset
Dim bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set SA = bd1.OpenRecordset("SelosAplicados", DB_OPEN_TABLE)
Set OS = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
SA.index = "SecondKey"
OS.index = "PrimaryKey"
  

SA.Seek "=", WNumero, WTipoSelo
If SA.NoMatch Then
   ApuraNumOSVxNumeroSelo = -999
   GoSub Finaliza:
Else
   GoSub ApuraOSV:
End If

Finaliza:
   SA.Close
   OS.Close
   bd1.Close
   Set SA = Nothing
   Set OS = Nothing
   Set bd1 = Nothing
   
   Exit Function
Return

ApuraOSV:
   OS.Seek "=", SA!Numos
   If Not OS.NoMatch Then
      ApuraNumOSVxNumeroSelo = SA!Numos
   Else
      ApuraNumOSVxNumeroSelo = 0
   End If
Return


TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function RemoveSelosRemovidos(WNumero As String, WTipoSelo As String, WNumos As Long) As Long
On Error GoTo TratamentodeErro:
Dim SA As Recordset
Dim bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set SA = bd1.OpenRecordset("SelosRemovidos", DB_OPEN_TABLE)
SA.index = "PrimaryKey"
 
BeginTrans
SA.Seek "=", WNumero, WNumos, WTipoSelo
If SA.NoMatch Then
   MsgBox "Selo inexistente, impossivel prosseguir...", vbCritical, "SGa"
   GoSub Finaliza:
Else
   SA.Delete
   MsgBox "Opera��o efetuada com sucesso...", vbCritical, "SGa"
   GoSub Finaliza:
End If

Finaliza:
   CommitTrans
   SA.Close
   bd1.Close
   Set SA = Nothing
   Set bd1 = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function RemoveSelosAplicados(WNumero As String, WTipoSelo As String, WNumos As Long) As Long
On Error GoTo TratamentodeErro:
Dim SA As Recordset
Dim bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set SA = bd1.OpenRecordset("SelosAplicados", DB_OPEN_TABLE)
SA.index = "PrimaryKey"
 
BeginTrans
SA.Seek "=", WNumero, WNumos, WTipoSelo
If SA.NoMatch Then
   MsgBox "Selo inexistente, impossivel prosseguir...", vbCritical, "SGa"
   GoSub Finaliza:
Else
   SA.Delete
   MsgBox "Opera��o efetuada com sucesso...", vbCritical, "SGa"
   GoSub Finaliza:
End If

Finaliza:
   CommitTrans
   SA.Close
   bd1.Close
   Set SA = Nothing
   Set bd1 = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function RetornaRedutorXTacografos(WNumos As Long) As String
On Error GoTo TratamentodeErro:
Dim TAC As Recordset, bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set TAC = bd1.OpenRecordset("TACOGRAFOSOS", DB_OPEN_TABLE)
TAC.index = "SecondKey"
 
RetornaRedutorXTacografos = 0
 
TAC.Seek "=", WNumos
If Not TAC.NoMatch Then
   If Not IsNull(TAC!MedRedutor) Then
      RetornaRedutorXTacografos = TAC!MedRedutor
   Else
      RetornaRedutorXTacografos = 0
   End If
End If

Finaliza:
   TAC.Close
   bd1.Close
   Set TAC = Nothing
   Set bd1 = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function
Function RetornaObservTacografos(WNumos As Long) As String
On Error GoTo TratamentodeErro:
Dim TAC As Recordset, bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set TAC = bd1.OpenRecordset("TACOGRAFOSOS", DB_OPEN_TABLE)
TAC.index = "SecondKey"
 
RetornaObservTacografos = " "
 
TAC.Seek "=", WNumos
If Not TAC.NoMatch Then
   If Not IsNull(TAC!ConstanteK) Then
      If Not IsNull(TAC!ObservacoesEnsaio) Then
         RetornaObservTacografos = TAC!ObservacoesEnsaio
      End If
   Else
      RetornaObservTacografos = " "
   End If
End If

Finaliza:
   TAC.Close
   bd1.Close
   Set TAC = Nothing
   Set bd1 = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Retorna_Nro_Protocolo_Selagem(WNumos As Long) As String
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
 
'BeginTrans
Retorna_Nro_Protocolo_Selagem = " "
 
tb.Seek "=", WNumos
If Not tb.NoMatch Then
   If Not IsNull(tb!Protocolo_Selagem) Then
      Retorna_Nro_Protocolo_Selagem = tb!Protocolo_Selagem
   Else
      Retorna_Nro_Protocolo_Selagem = ""
   End If
End If

Finaliza:
'   CommitTrans
   tb.Close
   bd.Close
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Retorna_GRU_Selagem(WNumos As Long) As String
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
 
'BeginTrans
Retorna_GRU_Selagem = ""
 
tb.Seek "=", WNumos
If Not tb.NoMatch Then
   If Not IsNull(tb!GRU_Selagem) Then
      Retorna_GRU_Selagem = tb!GRU_Selagem
   Else
      Retorna_GRU_Selagem = ""
   End If
End If

Finaliza:
'   CommitTrans
   tb.Close
   bd.Close
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Retorna_Razao_Social_Simul_Ensaio(WCNPJ As String) As String
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("EMPXFABRIC_PISTA", DB_OPEN_TABLE)
tb.index = "CNPJ_Fabric"
 
tb.Seek "=", WCNPJ
If Not tb.NoMatch Then
   If Not IsNull(tb!Razao_Social) Then
      Retorna_Razao_Social_Simul_Ensaio = tb!Razao_Social
   Else
      Retorna_Razao_Social_Simul_Ensaio = " "
   End If
End If

Finaliza:
   tb.Close
   bd.Close
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function Retorna_GRU_Ensaio(WNumos As Long) As String
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
 
'BeginTrans
Retorna_GRU_Ensaio = " "
 
tb.Seek "=", WNumos
If Not tb.NoMatch Then
   If Not IsNull(tb!GRU_Ensaio) Then
      Retorna_GRU_Ensaio = tb!GRU_Ensaio
   Else
      Retorna_GRU_Ensaio = " "
   End If
End If

Finaliza:
'   CommitTrans
   tb.Close
   bd.Close
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function RetornaVelIndicXTacografos(WNumos As Long) As String
On Error GoTo TratamentodeErro:
Dim TAC As Recordset, bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set TAC = bd1.OpenRecordset("TACOGRAFOSOS", DB_OPEN_TABLE)
TAC.index = "SecondKey"
 
BeginTrans
RetornaVelIndicXTacografos = " "
 
TAC.Seek "=", WNumos
If Not TAC.NoMatch Then
   If Not IsNull(TAC!VelIndic) Then
      RetornaVelIndicXTacografos = TAC!VelIndic
   Else
      RetornaVelIndicXTacografos = " "
   End If
End If

Finaliza:
   CommitTrans
   TAC.Close
   bd1.Close
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function RetornaConstXTacografos(WNumos As Long) As String
On Error GoTo TratamentodeErro:
Dim TAC As Recordset, bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set TAC = bd1.OpenRecordset("TACOGRAFOSOS", DB_OPEN_TABLE)
TAC.index = "SecondKey"
 
'BeginTrans
RetornaConstXTacografos = " "
 
TAC.Seek "=", WNumos
If Not TAC.NoMatch Then
   If Not IsNull(TAC!ConstanteK) Then
      RetornaConstXTacografos = TAC!ConstanteK
   Else
      RetornaConstXTacografos = " "
   End If
End If

Finaliza:
 '  CommitTrans
   TAC.Close
   bd1.Close
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function RetornaComplEndXTAcografosXOS(WNumos As Long) As String
On Error GoTo TratamentodeErro:
Dim TAC As Recordset, bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set TAC = bd1.OpenRecordset("TACOGRAFOSOS", DB_OPEN_TABLE)
TAC.index = "SecondKey"
'BeginTrans
RetornaComplEndXTAcografosXOS = " "
TAC.Seek "=", WNumos
If Not TAC.NoMatch Then
   If Not IsNull(TAC!Compl) Then
     RetornaComplEndXTAcografosXOS = TAC!Compl
   Else
     RetornaComplEndXTAcografosXOS = Space(1)
   End If
End If

Finaliza:
'   CommitTrans
   TAC.Close
   bd1.Close
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function RetornaNroEndXTAcografosXOS(WNumos As Long) As String
On Error GoTo TratamentodeErro:
Dim TAC As Recordset, bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set TAC = bd1.OpenRecordset("TACOGRAFOSOS", DB_OPEN_TABLE)
TAC.index = "SecondKey"
'BeginTrans
RetornaNroEndXTAcografosXOS = " "
TAC.Seek "=", WNumos
If Not TAC.NoMatch Then
   If Not IsNull(TAC!Nro) Then
     RetornaNroEndXTAcografosXOS = TAC!Nro
   Else
     RetornaNroEndXTAcografosXOS = "0"
   End If
End If

Finaliza:
'   CommitTrans
   TAC.Close
   bd1.Close
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function RetornaEmailXTacografosos(WNumos As Long) As String
On Error GoTo TratamentodeErro:
Dim TAC As Recordset, bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set TAC = bd1.OpenRecordset("TACOGRAFOSOS", DB_OPEN_TABLE)
TAC.index = "SecondKey"
'BeginTrans
RetornaEmailXTacografosos = " "
TAC.Seek "=", WNumos
If Not TAC.NoMatch Then
   If IsNull(TAC!Email) Or TAC!Email = "*" Or TAC!Email = "" Then
      RetornaEmailXTacografosos = " "
   Else
     RetornaEmailXTacografosos = TAC!Email
   End If
End If

Finaliza:
   'CommitTrans
   TAC.Close
   bd1.Close
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function RetornaUFXTAcografosXOS(WNumos As Long) As String
On Error GoTo TratamentodeErro:
Dim TAC As Recordset, bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set TAC = bd1.OpenRecordset("TACOGRAFOSOS", DB_OPEN_TABLE)
TAC.index = "SecondKey"
'BeginTrans
RetornaUFXTAcografosXOS = " "
TAC.Seek "=", WNumos
If Not TAC.NoMatch Then
   If Not IsNull(TAC!UF) And TAC!UF <> Space(1) Then
     RetornaUFXTAcografosXOS = TAC!UF
   Else
     RetornaUFXTAcografosXOS = " "
   End If
End If

Finaliza:
   'CommitTrans
   TAC.Close
   bd1.Close
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function RetornaBairroXTacografosXOS(WNumos As Long) As String
On Error GoTo TratamentodeErro:
Dim TAC As Recordset, bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set TAC = bd1.OpenRecordset("TACOGRAFOSOS", DB_OPEN_TABLE)
TAC.index = "SecondKey"
'BeginTrans
RetornaBairroXTacografosXOS = " "
TAC.Seek "=", WNumos
If Not TAC.NoMatch Then
   If Not IsNull(TAC!Bairro) And TAC!Bairro <> Space(1) Then
     RetornaBairroXTacografosXOS = TAC!Bairro
   Else
     RetornaBairroXTacografosXOS = " "
   End If
End If

Finaliza:
 '  CommitTrans
   TAC.Close
   bd1.Close
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function RetornaCNPJXTacografosXOS(WNumos As Long) As String
On Error GoTo TratamentodeErro:
Dim TAC As Recordset, bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set TAC = bd1.OpenRecordset("TACOGRAFOSOS", DB_OPEN_TABLE)
TAC.index = "SecondKey"
'BeginTrans
RetornaCNPJXTacografosXOS = " "
TAC.Seek "=", WNumos
If Not TAC.NoMatch Then
   If Not IsNull(TAC!CNPJ) And TAC!CNPJ <> Space(1) Then
     RetornaCNPJXTacografosXOS = TAC!CNPJ
   Else
     RetornaCNPJXTacografosXOS = " "
   End If
End If

Finaliza:
 '   CommitTrans
   TAC.Close
   bd1.Close
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function RetornaTelefoneXTAcografosXOS(WNumos As Long) As String
On Error GoTo TratamentodeErro:
Dim TAC As Recordset, bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set TAC = bd1.OpenRecordset("TACOGRAFOSOS", DB_OPEN_TABLE)
TAC.index = "SecondKey"
 
RetornaTelefoneXTAcografosXOS = " "
TAC.Seek "=", WNumos
If Not TAC.NoMatch Then
   If Not IsNull(TAC!Telefone) And TAC!Telefone <> Space(1) Then
     RetornaTelefoneXTAcografosXOS = TAC!Telefone
   Else
     RetornaTelefoneXTAcografosXOS = " "
   End If
End If

Finaliza:
   TAC.Close
   bd1.Close
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function RetornaCidadeXTacografosXOS(WNumos As Long) As String
On Error GoTo TratamentodeErro:
Dim TAC As Recordset, bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set TAC = bd1.OpenRecordset("TACOGRAFOSOS", DB_OPEN_TABLE)
TAC.index = "SecondKey"
'BeginTrans
RetornaCidadeXTacografosXOS = " "
TAC.Seek "=", WNumos
If Not TAC.NoMatch Then
   If Not IsNull(TAC!Cidade) And TAC!Cidade <> Space(1) Then
     RetornaCidadeXTacografosXOS = TAC!Cidade
   Else
     RetornaCidadeXTacografosXOS = " "
   End If
End If

Finaliza:
 '  CommitTrans
   TAC.Close
   bd1.Close
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function RetornaCepXTacografosXOS(WNumos As Long) As String
On Error GoTo TratamentodeErro:
Dim TAC As Recordset, bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set TAC = bd1.OpenRecordset("TACOGRAFOSOS", DB_OPEN_TABLE)
TAC.index = "SecondKey"
'BeginTrans
RetornaCepXTacografosXOS = " "
TAC.Seek "=", WNumos
If Not TAC.NoMatch Then
   If Not IsNull(TAC!Cep) And TAC!Cep <> Space(1) Then
     RetornaCepXTacografosXOS = TAC!Cep
   Else
     RetornaCepXTacografosXOS = " "
   End If
End If

Finaliza:
 '  CommitTrans
   TAC.Close
   bd1.Close
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function RetornaEnderecoXTacografosXOS(WNumos As Long) As String
On Error GoTo TratamentodeErro:
Dim TAC As Recordset, bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set TAC = bd1.OpenRecordset("TACOGRAFOSOS", DB_OPEN_TABLE)
TAC.index = "SecondKey"
'BeginTrans
RetornaEnderecoXTacografosXOS = " "
TAC.Seek "=", WNumos
If Not TAC.NoMatch Then
   If Not IsNull(TAC!Endereco) And TAC!Endereco <> Space(1) Then
     RetornaEnderecoXTacografosXOS = TAC!Endereco
   Else
     RetornaEnderecoXTacografosXOS = " "
   End If
End If

Finaliza:
   'CommitTrans
   TAC.Close
   bd1.Close
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function
Function Retorna_Tipo_Registro_Crono(WNumos As Long) As String

On Error GoTo TratamentodeErro:
Dim TAC As Recordset, bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set TAC = bd1.OpenRecordset("TACOGRAFOSOS", DB_OPEN_TABLE)
TAC.index = "SecondKey"
'BeginTrans
Retorna_Tipo_Registro_Crono = " "
TAC.Seek "=", WNumos
If Not TAC.NoMatch Then
   If Not IsNull(TAC!TipoRegistro) And TAC!TipoRegistro <> Space(1) Then
     Retorna_Tipo_Registro_Crono = TAC!TipoRegistro
   Else
     Retorna_Tipo_Registro_Crono = 1
   End If
End If

Finaliza:
 '   CommitTrans
   TAC.Close
   bd1.Close
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function Retorna_PAM_Crono(WNumos As Long) As String
On Error GoTo TratamentodeErro:
Dim TAC As Recordset, bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set TAC = bd1.OpenRecordset("TACOGRAFOSOS", DB_OPEN_TABLE)
TAC.index = "SecondKey"
'BeginTrans
Retorna_PAM_Crono = " "
TAC.Seek "=", WNumos
If Not TAC.NoMatch Then
   If Not IsNull(TAC!PAM) And TAC!PAM <> Space(1) Then
     Retorna_PAM_Crono = TAC!PAM
   Else
     Retorna_PAM_Crono = " "
   End If
End If

Finaliza:
 '   CommitTrans
   TAC.Close
   bd1.Close
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Retorna_Ano_Fabric_Crono(WNumos As Long) As String
On Error GoTo TratamentodeErro:
Dim TAC As Recordset, bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set TAC = bd1.OpenRecordset("TACOGRAFOSOS", DB_OPEN_TABLE)
TAC.index = "SecondKey"
'BeginTrans
Retorna_Ano_Fabric_Crono = " "
TAC.Seek "=", WNumos
If Not TAC.NoMatch Then
   If Not IsNull(TAC!ANOFABRICCRONO) And TAC!ANOFABRICCRONO <> Space(1) Then
     Retorna_Ano_Fabric_Crono = TAC!ANOFABRICCRONO
   Else
     Retorna_Ano_Fabric_Crono = " "
   End If
End If

Finaliza:
 '   CommitTrans
   TAC.Close
   bd1.Close
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Retorna_Escala_Veloc_Crono(WNumos As Long) As String
On Error GoTo TratamentodeErro:
Dim TAC As Recordset, bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set TAC = bd1.OpenRecordset("TACOGRAFOSOS", DB_OPEN_TABLE)
TAC.index = "SecondKey"
'BeginTrans
Retorna_Escala_Veloc_Crono = " "
TAC.Seek "=", WNumos
If Not TAC.NoMatch Then
   If Not IsNull(TAC!ESCALAVELOC) And TAC!ESCALAVELOC <> Space(1) Then
     Retorna_Escala_Veloc_Crono = TAC!ESCALAVELOC
   Else
     Retorna_Escala_Veloc_Crono = " "
   End If
End If

Finaliza:
 '   CommitTrans
   TAC.Close
   bd1.Close
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function RetornaRazaoSocialXTacografosXOS(WNumos As Long) As String
On Error GoTo TratamentodeErro:
Dim TAC As Recordset, bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set TAC = bd1.OpenRecordset("TACOGRAFOSOS", DB_OPEN_TABLE)
TAC.index = "SecondKey"
'BeginTrans
RetornaRazaoSocialXTacografosXOS = " "
TAC.Seek "=", WNumos
If Not TAC.NoMatch Then
   If Not IsNull(TAC!RazaoSocial) And TAC!RazaoSocial <> Space(1) Then
     RetornaRazaoSocialXTacografosXOS = TAC!RazaoSocial
   Else
     RetornaRazaoSocialXTacografosXOS = " "
   End If
End If

Finaliza:
 '   CommitTrans
   TAC.Close
   bd1.Close
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function AjustaTacografosXOS() As Long
On Error GoTo TratamentodeErro:
Dim OSV, TAC, Cli As Recordset, bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set TAC = bd1.OpenRecordset("TACOGRAFOSOS", DB_OPEN_TABLE)
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
Set OSV = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
TAC.index = "SecondKey"
Cli.index = "PrimaryKey"
OSV.index = "PrimaryKey"
'BeginTrans
If TAC.RecordCount <> 0 Then
   While Not TAC.EOF
         OSV.Seek "=", TAC!Numos
         If Not OSV.NoMatch Then
            Cli.Seek "=", OSV!CodigoCliente
            If Not Cli.NoMatch Then
               TAC.Edit
               If IsNull(TAC!RazaoSocial) Or TAC!RazaoSocial = Space(1) Then
                  If Not IsNull(Cli!RazaoSocial) Then
                     TAC!RazaoSocial = UCase(Mid$(Cli!RazaoSocial, 1, 50))
                  End If
               End If
               
               If IsNull(TAC!Endereco) Or TAC!Endereco = Space(1) Then
                  If Not IsNull(Cli!Endereco) Then
                     TAC!Endereco = UCase(Cli!Endereco)
                  End If
               End If
               
               If IsNull(TAC!Cep) Or TAC!Cep = Space(1) Then
                  If Not IsNull(Cli!EndPostal) Then
                     TAC!Cep = UCase(Cli!EndPostal)
                  End If
               End If
               
               If IsNull(TAC!Cidade) Or TAC!Cidade = Space(1) Then
                  If Not IsNull(Cli!Cidade) Then
                     TAC!Cidade = UCase(Cli!Cidade)
                  End If
               End If
               
               If IsNull(TAC!CNPJ) Or TAC!CNPJ = Space(1) Then
                  If Not IsNull(Cli!CgcCpf) Then
                     TAC!CNPJ = UCase(Cli!CgcCpf)
                  End If
               End If
               
               If IsNull(TAC!Bairro) Or TAC!Bairro = Space(1) Then
                  If Not IsNull(Cli!Bairro) Then
                     TAC!Bairro = UCase(Cli!Bairro)
                  End If
               End If

               If IsNull(TAC!UF) Or TAC!UF = Space(1) Then
                  If Not IsNull(Cli!Estado) Then
                     TAC!UF = UCase(Cli!Estado)
                  End If
               End If

               
            TAC.Update
            End If
            
         End If
         'Tac.Update
         TAC.MoveNext
   Wend
End If


Finaliza:
   CommitTrans
   TAC.Close
   Cli.Close
   OSV.Close
   bd1.Close
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function AtualizaRegistroTacografosXOS(WNumos As Long, WNroSerie As String, wTelefone As String, WRazaoSocial As String, WEndereco As String, wcep As String, WCidade As String, WCNPJ As String, WBairro As String, WUF As String, WMarca As String, WModelo As String, WConstanteK As String, WMedRedutor As String, wnro As String, wcompl As String, WVelocIndic As String, WEmail As String, wTipoRegistro As String, wEscalaVelocidade As String, wAnoFabricacao As String, wPAM As String, wConfirma As Boolean) As Long
On Error GoTo TratamentodeErro:
Dim TAC As Recordset, bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set TAC = bd1.OpenRecordset("TACOGRAFOSOS", DB_OPEN_TABLE)
TAC.index = "PrimaryKey"
BeginTrans
TAC.Seek "=", WNumos, WNroSerie
If TAC.NoMatch Then
   TAC.AddNew
Else
   TAC.Edit
End If
 TAC!Numos = WNumos
 TAC!NumeroTacografo = WNroSerie
 TAC!Marca = WMarca
 TAC!Modelo = WModelo
 TAC!ConstanteK = WConstanteK
' Call AtualizaCOEFWXOSV(wNumOS, CLng(WConstanteK))
 TAC!RazaoSocial = WRazaoSocial
 TAC!Endereco = WEndereco
 TAC!Cep = wcep
 TAC!Cidade = WCidade
 TAC!CNPJ = WCNPJ
 TAC!Bairro = WBairro
 TAC!UF = WUF
 TAC!Telefone = wTelefone
 If IsNull(WMedRedutor) Or WMedRedutor = "" Then
    TAC!MedRedutor = "0"
 Else
    TAC!MedRedutor = WMedRedutor
 End If
 TAC!Nro = wnro
 'TAC!Compl = WCompl
 TAC!VelIndic = WVelocIndic
 TAC!Email = WEmail
 TAC!TipoRegistro = wTipoRegistro
 TAC!ESCALAVELOC = wEscalaVelocidade
 TAC!ANOFABRICCRONO = wAnoFabricacao
 TAC!PAM = wPAM
 TAC!Confirma = wConfirma
 TAC.Update
 
Finaliza:
   CommitTrans
   DoCmd.Close acForm, "DialogoSelecaoTipoFrotaXOSV", acSaveYes
   TAC.Close
   bd1.Close
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function RetornaNroSeriexTacografos(WNumos As Long) As String
On Error GoTo TratamentodeErro:
Dim TAC As Recordset, bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set TAC = bd1.OpenRecordset("TACOGRAFOSOS", DB_OPEN_TABLE)
TAC.index = "SecondKey"
RetornaNroSeriexTacografos = "INDEF"
TAC.Seek "=", WNumos
If Not TAC.NoMatch Then
   If Not IsNull(TAC!NumeroTacografo) Then
      If TAC!NumeroTacografo = "0" Then
         RetornaNroSeriexTacografos = " "
      ElseIf Len(TAC!NumeroTacografo) <= 2 Then
            RetornaNroSeriexTacografos = " "
      Else
         RetornaNroSeriexTacografos = TAC!NumeroTacografo
      End If
   Else
      RetornaNroSeriexTacografos = " "
   End If
End If

Finaliza:
   TAC.Close
   bd1.Close
   
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function Retorna_Tipo_Tacografos_Str(WNumos As Long) As String
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

Retorna_Tipo_Tacografos_Str = " "
tb.Seek "=", WNumos
If Not tb.NoMatch Then
   If Not IsNull(tb!Tipo_Tco) Then
      If tb!Tipo_Tco = "1" Then
         Retorna_Tipo_Tacografos_Str = "DI�RIO"
      Else
          Retorna_Tipo_Tacografos_Str = "SEMANAL"
      End If
   Else
      Retorna_Tipo_Tacografos_Str = " "
   End If
End If

Finaliza:

   tb.Close
   bd.Close
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Retorna_Volts_Tacografos(WNumos As Long) As String
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

Retorna_Volts_Tacografos = " "
tb.Seek "=", WNumos
If Not tb.NoMatch Then
   If Not IsNull(tb!Volts) Then
      Retorna_Volts_Tacografos = tb!Volts & " VOLTS"
   Else
      Retorna_Volts_Tacografos = " "
   End If
End If

Finaliza:

   tb.Close
   bd.Close
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function RetornaModeloxTacografos(WNumos As Long) As String
On Error GoTo TratamentodeErro:
Dim TAC As Recordset, bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set TAC = bd1.OpenRecordset("TACOGRAFOSOS", DB_OPEN_TABLE)
TAC.index = "SecondKey"

RetornaModeloxTacografos = " "
TAC.Seek "=", WNumos
If Not TAC.NoMatch Then
   If Not IsNull(TAC!Modelo) Then
      RetornaModeloxTacografos = TAC!Modelo
   Else
      RetornaModeloxTacografos = " "
   End If
End If

Finaliza:

   TAC.Close
   bd1.Close
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function


Function Retorna_Protoc_OSV(WNumos As Long) As String
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", WNumos
If tb.NoMatch Then
   Retorna_Protoc_OSV = Space(1)
Else
   If IsNull(tb!Protocolo_Selagem) Then
      Retorna_Protoc_OSV = Space(1)
   Else
      Retorna_Protoc_OSV = tb!Protocolo_Selagem
   End If
End If

Finaliza:
   tb.Close
   bd.Close
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function FormatSelos(wSelo As String) As String
Dim SeloFormatado As String
Dim contador  As Integer
SeloFormatado = ""
For contador = 0 To Len(wSelo) - 1
   FormatSelos = FormatSelos & "0"
Next
End Function

Function GeracaoAutomaticaSelos(WPref As String, WNroInicial As String, WNroFinal As String, WFornec As Long, WDOcto As Long, WDtEntrada As Variant, WDtDocumento As Variant, WTipoSelo As String) As Long
On Error GoTo TratamentodeErro:
Dim Det As Recordset, bd1 As DataBase
Dim WNroSeq As String
Dim WChave As String
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Det = bd1.OpenRecordset("DETENTRADASSELOS", DB_OPEN_TABLE)
Det.index = "PrimaryKey"
WNroSeq = WNroInicial
  
While WNroSeq <= Val(WNroFinal)
   If WPref = Space(1) Then
      WChave = WNroSeq
   Else
      'WChave = UCase(WPref) & Format$(WNroSeq, "00000000")
      WChave = UCase(WPref) & Format$(WNroSeq, FormatSelos(WNroInicial))
   End If
   Det.Seek "=", WFornec, WDOcto, WDtEntrada, WDtDocumento, WTipoSelo, WChave
    If Det.NoMatch Then
       Det.AddNew
    Else
       Det.Edit
    End If
    Det!CodigoFornecedor = WFornec
    Det!NroDocumento = WDOcto
    Det!DataEntrada = WDtEntrada
    Det!DataDocumento = WDtDocumento
    'Det!Numero = WNroSeq
    Det!TipoSelo = WTipoSelo
    Det!Numero = WChave
    Det!CodEmpresa = Form_Menu.WEmpresa
    Det.Update
    WNroSeq = WNroSeq + 1
Wend

Finaliza:
   Det.Close
   bd1.Close
   MsgBox "Fim Normal de gera��o...", vbInformation, "SGa"
   DoCmd.Close acForm, "DialogoRecepcionarSelos", acSaveYes
   
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function MascaraSelos(WTipoSel As String, WNumero As String) As String

Dim bd1 As DataBase
Dim SA As Recordset
Dim wdv As String * 1

Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set SA = bd1.OpenRecordset("SelosAplicados", DB_OPEN_TABLE)
SA.index = "Secondkey"

SA.Seek "=", WNumero, WTipoSel
If SA.NoMatch Then
   wdv = Space(1)
Else
   If Not IsNull(SA!DV) Then
      wdv = SA!DV
   Else
      wdv = Space(1)
   End If
End If

If WTipoSel = "A" Then
   'MascaraSelos = UCase(Format$(WNumero & wdv, "!@ @.@@@.@@@-@"))
   MascaraSelos = UCase(Format$(WNumero, "!@@@@@@@@-@"))
Else
   'MascaraSelos = UCase(Format$(WNumero & wdv, "!@@@@@@@@-@"))
   MascaraSelos = UCase(Format$(WNumero, "!@@@@@@@@-@"))
End If

'bd1.Close
SA.Close
bd1.Close
Exit Function


End Function

Function Return_Selo_Caixa(WNumos As Long, WTipoSelo As String) As String

On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset

Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("SelosAplicados", DB_OPEN_TABLE)
tb.index = "TerceiraChave"

tb.Seek "=", WNumos, WTipoSelo

While WNumos = tb!Numos
    If tb!Entre_Sensor_Caixa = True Then
       Return_Selo_Caixa = tb!Numero & tb!DV
       GoTo FecharTabelas:
    End If
    tb.MoveNext
Wend

FecharTabelas:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function CapturaSelosInutilizados(WTipoSelo As String, WNumero As String) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim NREQ As Recordset
Dim WAux  As String

Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set NREQ = bd1.OpenRecordset("SelosInutilizados", DB_OPEN_TABLE)
NREQ.index = "SecondKey"

'If WNumero = "a0089077" Then
'   Stop
'End If

BeginTrans
NREQ.Seek "=", WNumero, WTipoSelo
If NREQ.NoMatch Then
   CapturaSelosInutilizados = "OFF"
   GoSub FecharTabelas:
Else
   CapturaSelosInutilizados = "ON"
End If

FecharTabelas:
   CommitTrans
   NREQ.Close
   bd1.Close
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function CapturaStatusSelos(WTipoSel As String, WNumero As String) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim SA, SD As Recordset
Dim WST  As String

Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set SA = bd1.OpenRecordset("SelosAplicados", DB_OPEN_TABLE)
Set SD = bd1.OpenRecordset("SelosInutilizados", DB_OPEN_TABLE)
SA.index = "SecondKey"
SD.index = "SecondKey"

'BeginTrans

WST = "L"

GoSub ApuraAplicacao:
GoSub ApuraDevolucao:
  

Finaliza:
   CapturaStatusSelos = WST
 '  CommitTrans
   SA.Close
   SD.Close
   bd1.Close
   Exit Function
Return

ApuraDevolucao:
   SD.Seek "=", WNumero, WTipoSel
   If SD.NoMatch Then
      WST = "L"
      GoSub Finaliza:
   Else
      WST = Mid$(SD!TipoOC, 1, 1)
      GoSub Finaliza:
   End If
Return

ApuraAplicacao:
   SA.Seek "=", WNumero, WTipoSel
   If SA.NoMatch Then
      WST = "L"
   Else
      WST = "U"
      GoSub Finaliza:
   End If
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function


Function CapturaSelosXOSV(WNumos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim NREQ As Recordset
Dim WAux  As String

Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set NREQ = bd1.OpenRecordset("SelosAplicados", DB_OPEN_TABLE)
NREQ.index = "QuintaChave"

NREQ.Seek "=", WNumos
If NREQ.NoMatch Then
   CapturaSelosXOSV = Space(1)
   GoSub FecharTabelas:
Else
   While Not NREQ.EOF
      If WNumos <> NREQ!Numos Then
         GoSub FecharTabelas:
      End If
      If NREQ!Numos = WNumos Then
         WAux = WAux & Mid$(NREQ!Numero, 1, 8) & NREQ!DV
      End If
      NREQ.MoveNext
   Wend
End If

FecharTabelas:
   CapturaSelosXOSV = WAux
   NREQ.Close
   bd1.Close
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaSelosAplicadosXCodigoSelo(WNumos As Long, WTipoSelo As String) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim NREQ As Recordset
Dim WAux  As String


Dim wtam_sel As String

Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set NREQ = bd1.OpenRecordset("SelosAplicados", DB_OPEN_TABLE)
NREQ.index = "TerceiraChave"

BeginTrans
NREQ.Seek "=", WNumos, WTipoSelo
If NREQ.NoMatch Then
   CapturaSelosAplicadosXCodigoSelo = Space(1)
   GoSub FecharTabelas:
Else
   While Not NREQ.EOF
      If WNumos <> NREQ!Numos Then
         GoSub FecharTabelas:
      End If
      'If WTipoSelo = "P" Then Stop
      If NREQ!Numos = WNumos And NREQ!TipoSelo = WTipoSelo Then
         
         If Len(NREQ!Numero) = 8 Then
            wtam_sel = Mid$(NREQ!Numero, 1, 8)
         ElseIf Len(NREQ!Numero) = 9 Then
                 wtam_sel = Mid$(NREQ!Numero, 1, 9)
         Else
            wtam_sel = Mid$(NREQ!Numero, 1, Len(NREQ!Numero))
         End If
         
         If IsNumeric(NREQ!ModeloSelo) Then
            WAux = WAux & Retorna_Value_Modelo_Selos(NREQ!ModeloSelo) & wtam_sel & NREQ!DV & "  "
         Else
            WAux = WAux & wtam_sel & NREQ!DV & "  "
         End If
         'WAux = WAux & Retorna_Value_Modelo_Selos(NREQ!ModeloSelo) & Mid$(NREQ!Numero, 1, 8) & NREQ!DV & "  "
      End If
      NREQ.MoveNext
   Wend
End If

FecharTabelas:
   CapturaSelosAplicadosXCodigoSelo = WAux
   CommitTrans
   NREQ.Close
   bd1.Close
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Public Sub Clear_Table_Status_Operacao()
Dim Table_Resultado As Recordset
Dim db As DataBase
Set db = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set Table_Resultado = db.OpenRecordset("Status_Operacao", DB_OPEN_TABLE)
BeginTrans
If Table_Resultado.RecordCount <> 0 Then
    While Not Table_Resultado.EOF
        Table_Resultado.Delete
        Table_Resultado.MoveNext
    Wend
End If
CommitTrans
End Sub
Public Sub Add_Message_Status_Operacao(ByVal wMensagem As String)
Dim Table_Resultado As Recordset
Dim db As DataBase
Set db = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set Table_Resultado = db.OpenRecordset("Status_Operacao", DB_OPEN_TABLE)
Clear_Table_Status_Operacao
BeginTrans
Table_Resultado.AddNew
Table_Resultado.Fields("StatusOperacao") = UCase(wMensagem)
Table_Resultado.Update
CommitTrans
End Sub
Public Sub Clear_Table_Consulta_Selagem_Completa()
Dim Table_Resultado As Recordset
Dim db As DataBase
Set db = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set Table_Resultado = db.OpenRecordset("Consulta_Selagem_Completa", DB_OPEN_TABLE)
BeginTrans
If Table_Resultado.RecordCount <> 0 Then
    While Not Table_Resultado.EOF
        Table_Resultado.Delete
        Table_Resultado.MoveNext
    Wend
End If
CommitTrans
End Sub
Public Function RetornaURL(WTipo As Long) As String
On Error GoTo TratamentodeErro:
Dim bd As DataBase, tb  As Recordset
Set bd = CurrentDb()
Set tb = bd.OpenRecordset("URLINMETRO", DB_OPEN_TABLE)
tb.index = "PrimaryKEy"
tb.Seek "=", WTipo
 If tb.NoMatch Then
     MsgBox "URL inexistente, impossivel prosseguir...", vbCritical, "Aviso"
     RetornaURL = "999999"
 Else
     RetornaURL = tb!DescricaoURL
 End If
Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Exit Function
    End If
End Function
Public Function VerificaSelagemxSubstituicaoInformada(ByVal WNumos As Long) As Boolean
On Error GoTo TratamentodeErro:
Dim bd As DataBase, tb  As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", WNumos
If tb.NoMatch Then
    MsgBox "Ordem de Servicos Inexistente, Impossivel Prosseguir.", vbCritical, "Aviso"
    VerificaSelagemxSubstituicaoInformada = False
Else
    If IsNull(tb.Fields("DataRegSelagem").Value) Then
       MsgBox "Selagem ou Substituicao de Selos Nao Informados, Impossivel Prosseguir.", vbCritical
       VerificaSelagemxSubstituicaoInformada = False
    Else
       VerificaSelagemxSubstituicaoInformada = True
    End If
End If
Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Exit Function
    End If
End Function
Public Sub RegistraSelagemxOSV(ByVal WNumos As Long, ByVal WData As Date)
On Error GoTo TratamentodeErro:
Dim bd As DataBase, tb  As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
BeginTrans
tb.index = "PrimaryKey"
tb.Seek "=", WNumos
If tb.NoMatch Then
    MsgBox "Ordem de Servicos Inexistente, Impossivel Prosseguir.", vbCritical, "Aviso"
    'VerificaSelagemxSubstituicaoInformada = False
    Exit Sub
Else
   tb.Edit
   tb.Fields("DataRegSelagem").Value = WData
   tb.Update
   CommitTrans
End If
Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Sub
Return
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Exit Sub
    End If

End Sub
Public Sub RegistraEnsaioxOSV(ByVal WNumos As Long, ByVal WData As Date)
On Error GoTo TratamentodeErro:
Dim bd As DataBase, tb  As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
BeginTrans
tb.index = "PrimaryKey"
tb.Seek "=", WNumos
If tb.NoMatch Then
    MsgBox "Ordem de Servicos Inexistente, Impossivel Prosseguir.", vbCritical, "Aviso"
    'VerificaSelagemxSubstituicaoInformada = False
    Exit Sub
Else
   tb.Edit
   tb.Fields("DataRegEnsaio").Value = WData
   tb.Update
   CommitTrans
End If
Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Sub
Return
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Exit Sub
    End If
End Sub

Function Cria_Tab_PlaxcliTemp_Certif_vencer() As Long

On Error GoTo TratamentodeErro:
Dim bd, bdt As DataBase
Dim tb, tbc As Recordset
Dim qtbc As QueryDef
Dim X As Long
Dim WDias As Long
Dim WVencto As Variant

Set bd = CurrentDb()
Set bdt = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bdt.OpenRecordset("PLACXCLITEMP_Cert", DB_OPEN_TABLE)
Set qtbc = bd.QueryDefs("CERTIFICADOS_A_VENCER")
Set tbc = qtbc.OpenRecordset()
tb.index = "PrimaryKey"

GoSub Inicializa_Tabela:

BeginTrans

If tbc.RecordCount <> 0 Then
   While Not tbc.EOF
   
         If IsDate(tbc!DataValidadeResultado) Then
         
         If DateValue(Format$(tbc!DataValidadeResultado, "dd/mm/yyyy")) Then
            WDias = Val(DateDiff("d", Date, DateValue(tbc!DataValidadeResultado)))
            If WDias > 0 And WDias <= 30 Then
               tb.Seek "=", tbc!Placa, tbc!CodigoCliente
               If tb.NoMatch Then
                  tb.AddNew
               Else
                  tb.Edit
               End If
               tb!Placa = tbc!Placa
               tb!CodigoCliente = tbc!CodigoCliente
               tb!MarcaVeiculo = tbc!MarcaVeiculo
               tb!ModeloVeiculo = tbc!ModeloVeiculo
               tb!RENAVAN = tbc!RENAVAN
               tb!DataValidadeResultado = tbc!DataValidadeResultado
               tb!DiasVencer = WDias
               tb.Update
            End If
         End If
         End If
         X = X + 1
         If X Mod 200 = 0 Then
             CommitTrans
             BeginTrans
         End If
         tbc.MoveNext
   Wend

End If

Finaliza:
   CommitTrans
   tbc.Close
   bd.Close
   bdt.Close
   Exit Function
Return

Inicializa_Tabela:
BeginTrans
If tbc.RecordCount <> 0 Then
   While Not tb.EOF
         tb.Delete
         tb.MoveNext
   Wend
End If
CommitTrans
Return


TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
       Resume
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume
    End If


End Function

Function Certificados_A_Vencer() As Boolean
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb, tbc As Recordset
Dim qtbc As QueryDef

Set bd = CurrentDb()
Set qtbc = bd.QueryDefs("PLAC_CLI_CONEC_VALID")
qtbc.Parameters(0) = Forms!Menu!DataInicial
qtbc.Parameters(1) = Forms!Menu!DataFinal
Set tbc = qtbc.OpenRecordset()

If tbc.RecordCount <> 0 Then
   Certificados_A_Vencer = True
Else
   Certificados_A_Vencer = False
End If

Finaliza:
   tbc.Close
   bd.Close
   Set tbc = Nothing
   Set bd = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
       Resume
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume
    End If

End Function

Function Demonstrativo_Lote_Placas(ByVal wCodCliente As Long, ByVal WSeq_Arq As Long) As String
On Error GoTo TratamentodeErro:

Dim WTraco As String * 55
Dim wHeaderRelatorio As String * 55
Dim wHeaderProprietario As String * 55
Dim wHeaderDados As String * 55
Dim wHeaderAssinaturaCanada As String * 55
Dim wHeaderRodape As String * 55
Dim WArquivo As String
Dim WEmpresa As String * 55
Dim WEndereco As String * 55
Dim wCidadeUF As String * 42
Dim WEmail As String * 40
Dim wfone As String * 15
Dim wPlaca As String * 8
Dim WFrota As String * 8
Dim wRenavam As String * 11
Dim wModeloVeiculo As String * 14
Dim wDataVencimento As String * 10
Dim wCaminhoArquivo As String
Dim WBairro As String
Dim ContVeiculos As Integer
Dim bd As DataBase, BaseLote As Recordset
Dim ofuncoes As CanadaConnectivity.FuncoesDiversas

ContVeiculos = 0
'WCont_Extrato = 0

Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set BaseLote = bd.OpenRecordset("Base_Dados_Lote", dbOpenTable)
Set ofuncoes = New CanadaConnectivity.FuncoesDiversas

wHeaderAssinaturaCanada = "                                        SGA - canainfo"
wHeaderRodape = "                                 Qtd. Veiclos "
wHeaderRelatorio = "             POSTO DE ENSAIO CRONOTACOGRAFOS          "
WTraco = "+----------------------------------------------------+"
'wHeaderDados = "PLACA       RENAVAM         MODELO         VENCIMENTO "
 wHeaderDados = "PLACA    FROTA    RENAVAM       MODELO      VENCIMENTO "
wHeaderProprietario = "              PROPRIETARIO DO VEICULO                  "
If ofuncoes.ExisteDiretorio(CapturaCaminhoBD() & "\Extratos") = False Then
   Call ofuncoes.CriaPastas(CapturaCaminhoBD(), "Extratos")
End If
Set ofuncoes = Nothing
BaseLote.index = "CodigoCliente"
BaseLote.Seek "=", wCodCliente
GoSub Le_Dados_Tabela:
GoSub Finaliza:

Le_Dados_Tabela:
   If BaseLote.RecordCount <> 0 Then
      wCaminhoArquivo = CapturaCaminhoBD() & "\Extratos" & "\EXT_" & Captura_Conteudo_Campos_CLI(wCodCliente, "CGCCPF") & "_" & Format(WSeq_Arq, "000000") & "_" & Format(Date, "ddMMyyyy") & "_" & Format(Now(), "hhmmss") & ".txt"
      Open wCaminhoArquivo For Output As #1
      'Open CapturaCaminhoBD() & "\Extratos" & "\EXT_" & CapturaCNPJCli(wCodCliente) & "_" & Format(Date, "ddMMyyyy") & "_" & Format(Now(), "hhmmss") & ".txt" For Output As #1
      GoSub Monta_Cabecalho_Principal:
      GoSub Monta_Dados_Relatorio:
      GoSub Monta_Rodape_Relatorio:
   End If
Return

Monta_Dados_Relatorio:
       Do While BaseLote.Fields("Codigocliente") = wCodCliente
          ContVeiculos = ContVeiculos + 1
          wPlaca = ""
          wRenavam = ""
          wModeloVeiculo = ""
          wDataVencimento = ""
          wPlaca = LTrim(Format$(BaseLote.Fields("Placa").Value, "!@@@@@@@@@@"))
          WFrota = LTrim(Format$(Captura_Campos_PlacXCli(wCodCliente, BaseLote.Fields("Placa").Value, "FROTA"), "!@@@@@@@@"))
          wRenavam = Format$(ApuraRENAVeicxOSVxCliente(wCodCliente, BaseLote.Fields("Placa").Value), "!@@@@@@@@@@@")
          wModeloVeiculo = Format$(ApuraModVeicxOSVxCliente(wCodCliente, BaseLote.Fields("Placa").Value), "!@@@@@@@@@@@@@@@")
          wDataVencimento = CapturaDataValidadeCertificadoVeiculoXPlaca(wCodCliente, BaseLote.Fields("Placa").Value)
          Print #1, wPlaca & Space(1) & WFrota & wRenavam & Space(1) & wModeloVeiculo & Space(1) & wDataVencimento
          BaseLote.MoveNext
          If BaseLote.EOF Then
             Exit Do
          End If
       Loop
       wHeaderRodape = "                                 Qtd. Veiculos: " & Format(CInt(ContVeiculos), "000000")

Return

Finaliza:
   BaseLote.Close
   bd.Close
   Set BaseLote = Nothing
   Set bd = Nothing
   Demonstrativo_Lote_Placas = wCaminhoArquivo
   Exit Function
Return

Monta_Cabecalho_Principal:
   If Nroserie() = 142356 Then
      WBairro = Space(1)(1)
      WEmpresa = Format$(Mid$(Fantasia(), 1, 50), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
   Else
      WBairro = Bairro()
      WEmpresa = Format$(Mid$(Captura_Conteudo_Campos_Empresa(Form_Menu.WEmpresa, "RazaoSocial"), 1, 50), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
   End If
   wfone = Fax()
   WEndereco = Format$(Endereco() & " - " & WBairro, "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
   wCidadeUF = Format$(Cidade & " - " & Estado(), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
   WEmail = Format(LCase(Email()), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
   Print #1, WTraco
   Print #1, wHeaderRelatorio
   Print #1, WTraco
   'Print #1, WEmpresa & WFone
   Print #1, WEmpresa
   Print #1, WEndereco
   Print #1, wCidadeUF
   Print #1, WEmail & wfone
   Print #1, WTraco
   Print #1, wHeaderProprietario
   Print #1, WTraco
   Print #1, "CNPJ/CPF: " & MascaraCgcCpf(Captura_Conteudo_Campos_CLI(wCodCliente, "CGCCPF"), Captura_Conteudo_Campos_CLI(wCodCliente, "Tipo"))
   Print #1, "RAZAO SOCIAL: " & UCase(Mid(Captura_Conteudo_Campos_CLI(wCodCliente, "RazaoSocial"), 1, 40))
   Print #1, WTraco
   Print #1, wHeaderDados
Return

Monta_Rodape_Relatorio:
  Print #1, WTraco
  Print #1, wHeaderRodape
  Print #1, WTraco
  Print #1, wHeaderAssinaturaCanada
  Close #1
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Public Function Enviar_Extrato_Lote_Placas(ByVal wVenctoI As Date, ByVal wVenctoF As Date, ByVal WAssunto As String, ByVal wMensagem As String, ByVal wRodape As String, ByVal wMensagemPadrao As Boolean) As Boolean
Dim bd As DataBase, tb As Recordset, qd As QueryDef
Dim wCodCliente As Long
Dim WCont_Extrato As Long
Dim X As Long
Set bd = CurrentDb()
Set qd = bd.QueryDefs("Consulta_Plac_Cli_Conec_Envio_SMS_S")
Set tb = qd.OpenRecordset()
Dim ofuncoes As CanadaConnectivity.FuncoesDiversas
Set ofuncoes = New CanadaConnectivity.FuncoesDiversas
WCont_Extrato = 1
If tb.RecordCount <> 0 Then
    Do While Not tb.EOF
      If tb.Fields("Enviar") = True Then
         If Not IsNull(tb.Fields("EndElet").Value) Then
            Dim FileNameExt As String
            FileNameExt = Demonstrativo_Lote_Placas(tb.Fields("CodigoCliente").Value, WCont_Extrato)
            WCont_Extrato = WCont_Extrato + 1
            Dim oCanadaUtils As CanadaUtils.Email
            Set oCanadaUtils = New CanadaUtils.Email
            Dim pAssunto As String
            Dim pMensagem As String
            If wMensagemPadrao = True Then
                pAssunto = AssuntoEmailCertificadoCronotacografo(tb.Fields("CodigoCliente").Value)
                pMensagem = MensagemEmailLotePlacas(wVenctoI, wVenctoF, tb.Fields("CodigoCliente").Value) & Chr(10) & Chr(10) & AssinaturaEmpresaAtivaEmail
            Else
                pAssunto = WAssunto
                pMensagem = wMensagem & Chr(10) & Chr(10) & wRodape
            End If
            oCanadaUtils.EmailDestino = tb.Fields("EndElet").Value
            oCanadaUtils.Assunto = pAssunto
            oCanadaUtils.Mensagem = pMensagem
            
            If oCanadaUtils.EmailAddressCheck(oCanadaUtils.EmailDestino) = False Then
                Call Atualiza_Status_EMAIL(tb.Fields("CodigoCliente").Value, "EMAIL DO DESTINATARIO INVALIDO")
            Else
                Call oCanadaUtils.EnviarEmailCronotacografos(Captura_Conteudo_Campos_Empresa(Empresa_Ativa(), "RazaoSocial"), oCanadaUtils.EmailDestino, pAssunto, pMensagem, FileNameExt)
                If Form_Envio_Email_Consulta_Lotes_Placas.CC = True Then
                    oCanadaUtils.EmailDestino = Email
                    If oCanadaUtils.EmailAddressCheck(oCanadaUtils.EmailDestino) = True Then Call oCanadaUtils.EnviarEmailCronotacografos(Captura_Conteudo_Campos_Empresa(Empresa_Ativa(), "RazaoSocial"), oCanadaUtils.EmailDestino, pAssunto, pMensagem, FileNameExt)
                    Call Atualiza_Status_EMAIL(tb.Fields("CodigoCliente").Value, oCanadaUtils.MensagemEnvio)
                End If
             End If
         Else
             Call Atualiza_Status_EMAIL(tb.Fields("CodigoCliente").Value, "EMAIL DO DESTINATARIO INVALIDO")
         End If
      End If
      tb.MoveNext
   Loop
End If
Finaliza:
tb.Close
bd.Close
Set tb = Nothing
Set bd = Nothing
Enviar_Extrato_Lote_Placas = True
Call Atualiza_Status_Envio_Email_Diario(2, PrimDiaMesCor, PrimDiaMesCor, Empresa_Ativa())
End Function



Public Function CapturaDataValidadeCertificadoVeiculoXPlaca(wCodCliente As Long, wPlaca As String) As Date
Dim bd As DataBase, tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("PlacxCli", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", wPlaca, wCodCliente
If Not tb.NoMatch Then
   CapturaDataValidadeCertificadoVeiculoXPlaca = FormatDateValue(tb!DataValidadeResultado)
Else
   CapturaDataValidadeCertificadoVeiculoXPlaca = "01/01/2001"
End If
Finaliza:
tb.Close
bd.Close
Set tb = Nothing
Set bd = Nothing
End Function

Function FormatDateValue(ByVal WValue) As Variant
If Not IsDate(WValue) Then
   WValue = #1/1/2001#
Else
   WValue = DateValue(WValue)
End If
FormatDateValue = WValue
End Function

Function Efetua_Importacao_ECF_Salc() As Long

Call Inicializa_ECF_Importar
Call Ler_Diretorio_Salc
Call Ler_Arquivos_ECF_Salc
Call Ler_Arquivos_Importados_ECF_Salc

Exit Function

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Deleta_Arq_Salc(wNomeArquivo As String) As Boolean

On Error GoTo TratamentodeErro:

Kill (CapturaCaminhoIntegracaoECF() & "\Resp\" & wNomeArquivo)

Exit Function

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       If Err.Number <> 53 Then
          MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       End If
          Resume Next
    End If

End Function

Function Importa_ECF_Salc(wNomeArquivo As String) As Boolean
On Error GoTo TratamentodeErro:
Dim db As DataBase
Dim tbe, tbi As Recordset
Dim X As Long
Dim RegTXT As String
Dim WSeq As String * 1
Dim Ret As Variant

Dim WCodFunc As String
Dim WCodigoCliente As String
Dim WFormaPagto As String
Dim WDataVenda As String
Dim WCOO As String
Dim WProtocolo_Aut_NFCe As String
Dim WChaveAcesso_NFCe As String
Dim WDataEmissao_NFC_e As String
Dim WNumos As String

Dim WCOO_I As String
Dim Wnro_Seq As String
Dim WCodItem As String
Dim WQuantidade As String
Dim WPrecoUnit As String
Dim WA_D As String
Dim WValor_A_D As String
Dim WCodigo_Trib As String

Set db = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tbe = db.OpenRecordset("ECF", DB_OPEN_TABLE)
Set tbi = db.OpenRecordset("ECF_Itens", DB_OPEN_TABLE)
tbe.index = "PrimaryKey"
tbi.index = "PrimaryKey"

Open wNomeArquivo For Input As #1

BeginTrans
While Not EOF(1)

   Line Input #1, RegTXT ' L� a linha para a vari�vel.
    
   WSeq = Mid$(RegTXT, 1, 1)
 
   X = X + 1
   Ret = SysCmd(SYSCMD_SETSTATUS, "Lendo registro " & Str(X))
 
   If WSeq = "0" Then
      GoSub Monta_Reg_Cabec:
      GoSub Inclui_Reg_Cabec:
   ElseIf WSeq = "1" Then
      GoSub Monta_Reg_Itens:
      GoSub Inclui_Reg_Itens:
   Else
      GoSub Finaliza:
   End If
   
   X = X + 1
   If X Mod 200 = 0 Then
      CommitTrans
      BeginTrans
   End If
Wend

Finaliza:
  CommitTrans
  tbe.Close
  tbi.Close
  db.Close
  Close #1
  Set tbe = Nothing
  Set tbi = Nothing
  Set db = Nothing
  Exit Function
Return

Inclui_Reg_Cabec:
  tbe.Seek "=", WCOO
  If tbe.NoMatch Then
     tbe.AddNew
  Else
     tbe.Edit
  End If
  tbe!CodigoResp = Val(WCodFunc)
  tbe!CodigoCliente = Val(WCodigoCliente)
  tbe!CodigoPagto = Val(WFormaPagto)
  tbe!datavenda = DateValue(WDataVenda)
  tbe!NRO_COO = WCOO
 ' If WProtocolo_Aut_NFCe <> "" Then
     tbe!protocolo_Aut_NFCe = WProtocolo_Aut_NFCe
 ' End If
 ' If WChaveAcesso_NFCe <> "" Then
     tbe!chaveacesso_nfce = WChaveAcesso_NFCe
 ' End If
'  If Wnumos <> Space(10) Then
     tbe!Numos = WNumos
'  Else
'     tbe!Numos = "9999999999"
'  End If
  
  
  tbe.Update
Return


Monta_Reg_Cabec:
  WCodFunc = Mid$(RegTXT, 2, 3)
  WCodigoCliente = Mid$(RegTXT, 5, 9)
  WFormaPagto = Mid$(RegTXT, 14, 2)
  WDataVenda = Mid$(RegTXT, 36, 10)
  WCOO = Mid$(RegTXT, 50, 9)
  WProtocolo_Aut_NFCe = Mid$(RegTXT, 103, 20)
  WChaveAcesso_NFCe = Mid$(RegTXT, 123, 44)
  WDataEmissao_NFC_e = Mid$(RegTXT, 197, 30)
  WNumos = Mid$(RegTXT, 26, 10)
Return

Inclui_Reg_Itens:
  tbi.Seek "=", WCOO, Val(Wnro_Seq)
  If tbi.NoMatch Then
     tbi.AddNew
  Else
     tbi.Edit
  End If
  tbi!NRO_COO = WCOO
  tbi!Nro_Sequencial = Val(Wnro_Seq)
  tbi!CodItem = Val(WCodItem)
  tbi!Quantidade = Val(WQuantidade) / 1000
  tbi!PrecoUnitario = Val(WPrecoUnit) / 100
  tbi!valor_acresc_Desc = Val(WValor_A_D) / 100
  tbi!codigo_tributo = WCodigo_Trib
  tbi.Update

Return

Monta_Reg_Itens:
  WCOO_I = WCOO
  Wnro_Seq = Mid$(RegTXT, 2, 3)
  WCodItem = Mid$(RegTXT, 5, 7)
  WQuantidade = Mid$(RegTXT, 12, 7)
  WPrecoUnit = Mid$(RegTXT, 26, 7)
  WA_D = Mid$(RegTXT, 58, 1)
  WValor_A_D = Mid$(RegTXT, 59, 8)
  WCodigo_Trib = Mid$(RegTXT, 67, 2)
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CriaOSV(wCodCliente As Long, WDataAbertura As Variant, WDepto As Long, WResponsavel As Long, WOperacao As Long) As Long
On Error GoTo TratamentodeErro:
Dim OSV As Recordset, bd1 As DataBase
Dim WNumos As Long
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set OSV = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
OSV.index = "PrimaryKey"
BeginTrans
WNumos = ProximaOSV(1)
OSV.Seek "=", WNumos
If OSV.NoMatch Then
   OSV.AddNew
Else
   OSV.Edit
End If
OSV!Numos = WNumos
OSV!TipoOs = "1"
OSV!CodigoCliente = wCodCliente
OSV!CodigoPagto = WOperacao
OSV!Emitirnf = "1"
OSV!Responsavel = WResponsavel
OSV!DataAbertura = WDataAbertura
OSV!HoraAbertura = Time()
OSV!datafechamento = WDataAbertura
OSV!HoraFechamento = Time()
OSV!DataOrcamento = WDataAbertura
OSV!CodigoDepto = WDepto
OSV!CodEmpresa = Empresa_Ativa()
OSV!EmitirLaudo = "N"
OSV.Update


Finaliza:
   CommitTrans
   OSV.Close
   bd1.Close
   CriaOSV = WNumos
   Set OSV = Nothing
   Set bd1 = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Inicialica_Modelo_Email() As Long

On Error GoTo TratamentodeErro:
Dim bd As DataBase, tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("MODELO_EMAIL", DB_OPEN_TABLE)

tb.index = "PrimaryKey"
BeginTrans
tb.Seek "=", 0
If tb.NoMatch Then
   tb.AddNew
Else
   tb.Edit
End If
tb!Codigo = 0
tb!NomeMensagem = "Envio de WhatsApp"
tb!AssuntoMensagem = "Validade de Certificados Cronotacografos"
tb.Update

Finaliza:
  CommitTrans
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return


TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function RetornaCodigoCategoriaxCodItem(ByVal WCodItem As Integer) As Integer
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("EST", DB_OPEN_TABLE)

tb.index = "PrimaryKey"
tb.Seek "=", WCodItem
If Not tb.NoMatch Then
    RetornaCodigoCategoriaxCodItem = UCase(tb!CodigoCategoria)
Else
    RetornaCodigoCategoriaxCodItem = 0
End If

Finaliza:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
       GoSub Finaliza:
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       GoSub Finaliza:
    End If
End Function

Function RetornaDescricaoResponsavel(ByVal WCodResp As Integer) As String
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("RESP", DB_OPEN_TABLE)

tb.index = "PrimaryKey"
tb.Seek "=", WCodResp
If Not tb.NoMatch Then
    RetornaDescricaoResponsavel = IIf(IsNull(tb!Codinome), "INDEFINIDO", UCase(tb!Codinome))
Else
    RetornaDescricaoResponsavel = ""
End If

Finaliza:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
       GoSub Finaliza:
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       GoSub Finaliza:
    End If
End Function

Function RetornaDescricaoCATSERV(ByVal WCodServ As Integer) As String
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("CATSERV", DB_OPEN_TABLE)

tb.index = "PrimaryKey"
tb.Seek "=", WCodServ
If Not tb.NoMatch Then
    If Not IsNull(tb!Descricao) Then
       RetornaDescricaoCATSERV = UCase(tb!Descricao)
    Else
       RetornaDescricaoCATSERV = ""
    End If
Else
    RetornaDescricaoCATSERV = ""
End If

Finaliza:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
       GoSub Finaliza:
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       GoSub Finaliza:
    End If
End Function

Function Ler_Arquivos_Importados_ECF_Salc() As Boolean

On Error GoTo TratamentodeErro:
Dim db, dba As DataBase
Dim tb, tbi, tbc, tba, tbo, tbf As Recordset
Dim WST As String
Dim WNumos As Long
Dim WNumrequisicao As Long
Dim wsaldo As Double
Dim WMOD_BC_ICMS_ST As Long
Dim WMod_BC_ICMS As Long
Dim WAliq_Red_BC_ST As Double
Dim WCodCSTPis As Long
Dim WCodCSTCofins As Long
Dim WAliqPis As Double
Dim WAliqCofins As Double
Dim WAliqINSS As Double
Dim WAliqIRPJ As Double
Dim WAliqCSLL As Double
Dim WCFOP As Long
Dim WCST As Double
Dim WGera_Financ As Boolean

Set db = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set dba = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = db.OpenRecordset("ECF", DB_OPEN_TABLE)
Set tbi = db.OpenRecordset("ECF_Itens", DB_OPEN_TABLE)
Set tbc = dba.OpenRecordset("CLI", DB_OPEN_TABLE)
Set tba = dba.OpenRecordset("RESP", DB_OPEN_TABLE)
Set tbf = dba.OpenRecordset("CONDPAGTO", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tbi.index = "Nro_COO"
tbc.index = "PrimaryKey"
tba.index = "PrimaryKey"
tbf.index = "FormaPagto"

WST = "N"

If tb.RecordCount <> 0 Then
   While Not tb.EOF
         
         If Valida_Nro_COOCCF_OSV(tb!NRO_COO) = False Then
            GoSub Valida_Cliente:
            If WST = "P" Then
               If IsNumeric(tb!Numos) Then            ' PEDIDO GERADO PELO SGA
                  'GERAR SOMENTE NOTA FISCAL
               Else
                  
                  WNumos = CriaOSV(tb!CodigoCliente, tb!datavenda, 1, tb!CodigoResp, tbf!CodigoPagto)
                  Call LogSGa("OS", "Import ECFs", "Criando OSV", Captura_Conteudo_Campo_RESP(Usuario_Ativo(), "nome"), Date, Time(), "OSV: " & WNumos & " -> " & CapturaRazaoSocialCli(tb!CodigoCliente) & " Opera��o : " & CapturaDescricaoPrazo(tbf!CodigoPagto), Environ("ComputerName"))
                  Call Atualiza_Nro_COOCCF_OSV(WNumos, tb!NRO_COO)
                  Call LogSGa("OS", "Import ECFs", "Atualizando COOCCF em  OSV", Captura_Conteudo_Campo_RESP(Usuario_Ativo(), "nome"), Date, Time(), "OSV: " & WNumos & " -> " & CapturaRazaoSocialCli(tb!CodigoCliente) & " Opera��o : " & CapturaDescricaoPrazo(tbf!CodigoPagto), Environ("ComputerName"))
                  GoSub Le_Itens_ECF:
                  Call RegravaOsFechamento(WNumos, tb!datavenda)
                  Call LogSGa("OS", "Import ECFs", "Encerrando OSV", Captura_Conteudo_Campo_RESP(Usuario_Ativo(), "nome"), Date, Time(), "OSV: " & WNumos & " -> " & CapturaRazaoSocialCli(tb!CodigoCliente) & " Opera��o : " & CapturaDescricaoPrazo(tbf!CodigoPagto), Environ("ComputerName"))
                  Call CriaDemonsReceitas(WNumos, (TotalServicosOrdemServico(WNumos) + CalculoMercadorias(WNumos)), 0, 0, "DIN")
                  Call GeraContasReceber(WNumos, tbf!CodigoPagto, tb!CodigoCliente, (TotalServicosOrdemServico(WNumos) + CalculoMercadorias(WNumos)), tb!datavenda, 0, 1, "NRO COO:" & tb!NRO_COO, RetornaPrevRecXPrazo(tbf!CodigoPagto), 1, CapturaCarteiraRecebtosXCondicaoPagto(tbf!CodigoPagto), 0, tbf!GeraFinanceiro, tbf!Composicao, 0, tbf!Descricao, Captura_Conteudo_Campos_CONDPAGTO(tbf!CodigoPagto, "CarteirasVinculada"))
                  Call LogSGa("OS", "Import ECFs", "Gerando Financ OSV", Captura_Conteudo_Campo_RESP(Usuario_Ativo(), "nome"), Date, Time(), "OSV: " & WNumos & " -> " & CapturaRazaoSocialCli(tb!CodigoCliente) & " Opera��o : " & CapturaDescricaoPrazo(tbf!CodigoPagto) & " Total OSV -> " & Format$((TotalServicosOrdemServico(WNumos) + CalculoMercadorias(WNumos)), "0.00"), Environ("ComputerName"))
               End If
               
             End If
         End If
         tb.MoveNext
   Wend
End If

Finaliza:
  tb.Close
  tbi.Close
  tba.Close
  tbc.Close
  tbf.Close
  db.Close
  dba.Close
  Set tb = Nothing
  Set tbi = Nothing
  Set tba = Nothing
  Set tbc = Nothing
  Set tbf = Nothing
  Set db = Nothing
  Set dba = Nothing
  Exit Function
Return

Le_Itens_ECF:
  tbi.Seek "=", tb!NRO_COO
  If Not tbi.NoMatch Then
     While tbi!NRO_COO = tb!NRO_COO
           WMOD_BC_ICMS_ST = Captura_Conteudo_Campo_Tabelas_Basicas(tbi!CodItem, "EST", "ModBCICMSST", "PrimaryKey")
           WMod_BC_ICMS = Captura_Conteudo_Campo_Tabelas_Basicas(tbi!CodItem, "EST", "ModalidadeBCICMS", "PrimaryKey")
           WAliq_Red_BC_ST = Retorna_Aliq_RED_BC_ST(Captura_CodigoNCM(tbi!CodItem), CapturaUFCadastroCliente(tbc!CodigoCliente))
        
           'WNumRequisicao = CriaRequisicao_Atendimento(WNumos)
           
           WCFOP = CapturaCFOPXEST(tbi!CodItem, tb!CodigoCliente)
           WCST = Captura_CST_EST(tbi!CodItem)
           If Captura_Conteudo_Campo_Tabelas_Basicas(tbi!CodItem, "EST", "CodigoCST_PIS", "PrimaryKey") = 0 Then
              WCodCSTPis = 4
           Else
              WCodCSTPis = Captura_Conteudo_Campo_Tabelas_Basicas(tbi!CodItem, "EST", "CodigoCST_PIS", "PrimaryKey")
           End If
           If Captura_Conteudo_Campo_Tabelas_Basicas(tbi!CodItem, "EST", "CodigoCST_COFINS", "PrimaryKey") = 0 Then
              WCodCSTCofins = 4
           Else
              WCodCSTCofins = Captura_Conteudo_Campo_Tabelas_Basicas(tbi!CodItem, "EST", "CodigoCST_COFINS", "PrimaryKey")
           End If
           If WCodCSTPis = 4 Then
              WAliqPis = 0
              WAliqCofins = 0
           Else
              WAliqPis = PIS()
              WAliqCofins = COFIS()
           End If
           WAliqINSS = INSS()
           WAliqIRPJ = IRPJ()
           WAliqCSLL = CSSLL()
           
           WNumrequisicao = Captura_Conteudo_Campos_Empresa(1, "Proxima_Req")
           Call Atualiza_Campos_EMPRESA(1, "Proxima_Req", WNumrequisicao + 1)
           Call CriaRequisicao(WNumos, WNumrequisicao)
           
           Call IncluiItemRequisitado(WNumrequisicao, Empresa_Ativa(), tbi!CodItem, tbi!Quantidade, tbi!PrecoUnitario, tbi!valor_acresc_Desc, WCFOP, Apura_Perc_Comis_Pecas_Mod_02_Func(Val(tb!CodigoResp)), tb!CodigoResp, Saldo_Estoque(tbi!CodItem, tb!datavenda, Empresa_Ativa()), WNumos, tbi!PrecoUnitario * (Captura_Aliq_IPI_EST(tbi!CodItem) / 100), 0, 0, 0, 0, WCST, 0, 0, 0, 0, WMOD_BC_ICMS_ST, WMod_BC_ICMS, WAliq_Red_BC_ST, AliqICMSCupom, WCodCSTPis, WCodCSTCofins, WAliqPis, WAliqCofins, WAliqINSS, WAliqIRPJ, WAliqCSLL)
                                      'WNumRequisicao, WCodEmpresa    , WCodItem   , WQuantidade   , WValorUnit       , WDesconto            , WCFOP, WAliq_Comis_Atend                                     , WAtend_Resp  , WSaldo                                                   , WNumos , wipi                                                        , WVlr_Fret, WVlr_Out, WVlr_ICMSST, WBase_ICMSST , WCodigo_CST , WP_ICMSST , WAliq_MVA, WVlr_Import   , WVlr_Seguro, WMOD_BC_ICMS_ST, WMod_BC_ICMS, WAliq_Red_BC_ST, WAliq_ICMS    , WCodCSTPis, WCodCSTCofins, WAliqPis, WAliqCofins, WAliqINSS, WAliqIRPJ, WAliqCSLL)
           tbi.MoveNext
           If tbi.EOF Then
              Return
           End If
        
     Wend
  End If

Return

Valida_Cliente:
   tbc.Seek "=", tb!CodigoCliente
   If tbc.NoMatch Then
      WST = "N"
      MsgBox "Cliente n�o cadastrado, impossivel prosseguir...", vbCritical, "SGa"
   Else
      WST = "P"
      GoSub Valida_Func:
   End If
Return


Valida_Func:
   tba.Seek "=", tb!CodigoResp
   If tba.NoMatch Then
      WST = "N"
      MsgBox "Atendente n�o cadastrado, impossivel prosseguir...", vbCritical, "SGa"
   Else
      WST = "P"
      GoSub Valida_Forma_Pagto:
   End If
Return

Valida_Forma_Pagto:
  If tb!CodigoPagto = 1 Then
     tbf.Seek "=", "DIN"
  Else
     tbf.Seek "=", "CAR"
  End If
  If tbf.NoMatch Then
     WST = "N"
  Else
     WST = "P"
  End If
Return


TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If



End Function

Function Inicializa_ECF_Importar() As Boolean
On Error GoTo TratamentodeErro:
Dim db As DataBase
Dim tb, tba, tbb As Recordset

Set db = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = db.OpenRecordset("ECF_A_Importar", DB_OPEN_TABLE)
Set tba = db.OpenRecordset("ECF", DB_OPEN_TABLE)
Set tbb = db.OpenRecordset("ECF_Itens", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tba.index = "PrimaryKey"
tbb.index = "PrimaryKey"

GoSub Inicializa_ECF_A_Importar:
GoSub Inicializa_ECF:
GoSub Inicializa_ECF_Itens:

Finaliza:
  tb.Close
  tba.Close
  tbb.Close
  db.Close
  Set tb = Nothing
  Set db = Nothing
  Exit Function
Return

Inicializa_ECF:
  BeginTrans
  If tba.RecordCount <> 0 Then
     While Not tba.EOF
           tba.Delete
           tba.MoveNext
     Wend
  End If
  CommitTrans
Return

Inicializa_ECF_Itens:
  BeginTrans
  If tbb.RecordCount <> 0 Then
     While Not tbb.EOF
           tbb.Delete
           tbb.MoveNext
     Wend
  End If
  CommitTrans
Return

Inicializa_ECF_A_Importar:
  BeginTrans
  If tb.RecordCount <> 0 Then
     While Not tb.EOF
           Call Deleta_Arq_Salc(tb!NomeArquivo)
           tb.Delete
           tb.MoveNext
     Wend
  End If
  CommitTrans
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Ler_Arquivos_ECF_Salc() As Long

On Error GoTo TratamentodeErro:
Dim db As DataBase
Dim tb As Recordset

Set db = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = db.OpenRecordset("ECF_A_Importar", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

If tb.RecordCount <> 0 Then
   While Not tb.EOF
        ' If tb!NomeArquivo = "01000000484.ecf" Then
         '   Stop
         'End If
         Call Importa_ECF_Salc(CapturaCaminhoIntegracaoECF() & "\Resp\" & tb!NomeArquivo)
         tb.MoveNext
   Wend
End If

Finaliza:
   tb.Close
   db.Close
   Set tb = Nothing
   Set db = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
    
End Function

Function ECF_A_Importar(WNomeTabela As String)

On Error GoTo TratamentodeErro:
Dim db As DataBase
Dim tb As Recordset

Set db = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = db.OpenRecordset("ECF_A_Importar", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

BeginTrans
tb.Seek "=", WNomeTabela
If tb.NoMatch Then
   tb.AddNew
Else
   tb.Edit
End If
tb!NomeArquivo = WNomeTabela
tb.Update
CommitTrans

Finaliza:
  tb.Close
  db.Close
  Set tb = Nothing
  Set db = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function Ler_Diretorio_Salc() As Long
On Error GoTo TratamentodeErro:
Dim DirectoryName
Dim path As String
Dim contador As Long
Dim irow
Dim icol
        
Dim WECF As String
contador = 0
        
irow = 11
     
path = CapturaCaminhoIntegracaoECF() & "\Resp\"
'Path = CapturaCaminhoIntegracaoECF() & "\"
     
DirectoryName = Dir(path, vbDirectory)
                
Do Until DirectoryName = ""
   If DirectoryName <> "." And DirectoryName <> ".." Then
      contador = contador + 1
      Debug.Print DirectoryName
      If (GetAttr(path & DirectoryName) And vbDirectory) = vbDirectory Then
      End If
   End If
   icol = 2
   ' Cells(irow, icol).Value = DirectoryName
   icol = icol + 1
   DirectoryName = Dir
  ' MsgBox "ext" & Mid$(DirectoryName, Len(DirectoryName) - 3, 4)
   If Len(DirectoryName) > 5 Then
      If Mid$(DirectoryName, Len(DirectoryName) - 3, 4) = ".ECF" Then
         WECF = DirectoryName
         Call ECF_A_Importar(WECF)
      End If
   End If
            
Loop
Debug.Print contador
Exit Function

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Corpo_Mensagem_Consulta_Certif_Whats_Avulsa(WCodigoCliente As Long, wPlaca As String, WVencto As Variant, WTipo_Certif As String) As String
On Error GoTo TratamentodeErro:
Dim wMensagem As String
Dim WRazaoSocial As String
'WRazaoSocial = CapturaRazaoSocialCli(WCodigoCliente)

wMensagem = "DADOS DO VEICULO " & Chr(10) & Chr(10) & "PLACA: " & wPlaca & Chr(10) & "MARCA: " & Retorna_Dados_Consulta_Selagem_Completa(wPlaca, "MARCAVEICULO") & Chr(10) & "MODELO: " & Retorna_Dados_Consulta_Selagem_Completa(wPlaca, "MODELOVEICULO") & Chr(10) & "ANO: " & Retorna_Dados_Consulta_Selagem_Completa(wPlaca, "ANOVEICULO") & Chr(10) & "RENAVAN: " & Retorna_Dados_Consulta_Selagem_Completa(wPlaca, "RENAVAM") & Chr(10) & "VENCIMENTO:" & WVencto & Chr(10) & "TIPO: " & WTipo_Certif
    
 
Corpo_Mensagem_Consulta_Certif_Whats_Avulsa = wMensagem


Exit Function


TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Ret_Tipo_Notif(WTipo As String) As String
On Error GoTo TratamentodeErro:
If WTipo = "1" Then
   Ret_Tipo_Notif = "Envio po Email"
ElseIf WTipo = "2" Then
       Ret_Tipo_Notif = "Envio por SMS"
Else
    Ret_Tipo_Notif = "Envio por What�s APP"
End If
Exit Function

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Inicializa_Tabela_Importacao_CLI() As Long

DoCmd.Hourglass True
On Error GoTo TratamentodeErro:
Dim bdc As DataBase
Dim tbc As Recordset

Dim X As Long

Set bdc = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tbc = bdc.OpenRecordset("Tabela_CLI_CSV", DB_OPEN_TABLE)

tbc.index = "PrimaryKey"

BeginTrans
If tbc.RecordCount <> 0 Then
   While Not tbc.EOF
         tbc.Delete
         
         X = X + 1
         If X Mod 200 = 0 Then
            CommitTrans
            BeginTrans
         End If
         
         tbc.MoveNext
   Wend
End If

Finaliza:
   CommitTrans
   tbc.Close
   bdc.Close
   Set tbc = Nothing
   Set bdc = Nothing
   DoCmd.Hourglass False
   Exit Function
Return


TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume
    End If

End Function

Function Inicializa_Tabela_Importacao_VDO() As Long

DoCmd.Hourglass True
On Error GoTo TratamentodeErro:
Dim bdc As DataBase
Dim tbc As Recordset

Dim X As Long

Set bdc = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tbc = bdc.OpenRecordset("Tabela_VDO_CSV", DB_OPEN_TABLE)

tbc.index = "PrimaryKey"

BeginTrans
If tbc.RecordCount <> 0 Then
   While Not tbc.EOF
         tbc.Delete
         
         X = X + 1
         If X Mod 200 = 0 Then
            CommitTrans
            BeginTrans
         End If
         
         tbc.MoveNext
   Wend
End If

Finaliza:
   CommitTrans
   tbc.Close
   bdc.Close
   Set tbc = Nothing
   Set bdc = Nothing
   
   DoCmd.Hourglass False
   Exit Function
Return


TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume
    End If




End Function

Function Importa_Planiha_VDO(WCodigoCAtegoria As Long, WCodigoMarca As Long) As Long

DoCmd.Hourglass True
On Error GoTo TratamentodeErro:
Dim bd, bdc As DataBase
Dim tb, tba, tbc As Recordset

Dim WCodItem As Long
Dim Item_Existente As Boolean

Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set bdc = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tbc = bdc.OpenRecordset("Tabela_VDO_CSV", DB_OPEN_TABLE)
Set tb = bd.OpenRecordset("EST", DB_OPEN_TABLE)
Set tba = bd.OpenRecordset("EMPXEST", DB_OPEN_TABLE)

tbc.index = "PrimaryKey"
tba.index = "PrimaryKey"

If tbc.RecordCount <> 0 Then
   While Not tbc.EOF
         If tbc!Selecionar Then
            GoSub Valida_Codigo_Fabricante:
            If Item_Existente = False Then
               GoSub Grava_EST:
               GoSub Grava_EMPXEST:
            End If
         End If
         tbc.MoveNext
   Wend
End If

Finaliza:
   tbc.Close
   tba.Close
   tb.Close
   bd.Close
   bdc.Close
   Set tbc = Nothing
   Set tba = Nothing
   Set tb = Nothing
   Set bd = Nothing
   Set bdc = Nothing
   DoCmd.Hourglass False
   Exit Function
Return

Valida_Codigo_Fabricante:
   tb.index = "CodigoFabricante"
   tb.Seek "=", tbc!CodigoFabricante
   If tb.NoMatch Then
      Item_Existente = False
    Else
      Item_Existente = True
    End If
Return

Grava_EMPXEST:
   BeginTrans
   tba.Seek "=", Val(Form_Menu.WEmpresa), WCodItem
   If tba.NoMatch Then
      tba.AddNew
   Else
      tba.Edit
   End If
   
   tba!CodigoEmpresa = Val(Form_Menu.WEmpresa)
   tba!CodigoItem = WCodItem
   tba!Localizacao = "ST"
   If tbc!PrecoVenda = 0 Then
      tba!PrecoVenda = 0.001
   Else
      tba!PrecoVenda = tbc!PrecoVenda
   End If
   tba!Ativo = -1
   tba.Update
   CommitTrans
Return

Grava_EST:
   tb.index = "PrimaryKey"
   WCodItem = ProxCodItem()
   BeginTrans
   tb.Seek "=", WCodItem
   If tb.NoMatch Then
      tb.AddNew
    Else
      tb.Edit
    End If
    tb!CodItem = WCodItem
    tb!Codinome = Mid$(tbc!Descricao, 1, 255)
    tb!CodigoFabricante = Mid$(tbc!CodigoFabricante, 1, 25)
    tb!CodigoCategoria = WCodigoCAtegoria
    tb!Descricao = Mid$(tbc!Descricao, 1, 255)
    tb!Ativo = -1
    tb!CodigoMarca = WCodigoMarca
    tb!Codigoncm = tbc!Codigoncm
    tb!ClassificacaoFiscal = tbc!Codigoncm
    tb!CodigoInterno = WCodItem
    tb!CodigoCST = 60
    tb!AliqIPI = 0
    tb!ModalidadeBCICMS = 0
    tb!ReducaoBC = 0
    tb!ModBCICMSST = 0
    tb!ReducaoBCST = 0
    tb!MargICMSST = 0
    tb!REFORIG = tbc!CodigoFabricante
    tb!GeneroItem = 90
    tb!CFOP = 112
    tb!PesoBruto = tbc!Peso
    tb!CodBarra = tbc!CodigoBarras
    tb!CodigoCST_PIS = 1
    tb!CodigoCST_IPI = 1
    tb!CodigoCST_COFINS = 1
    tb.Update
    CommitTrans
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume
    End If



End Function
Function Atualiza_Precos_Planilha_VDO() As Long

DoCmd.Hourglass True
On Error GoTo TratamentodeErro:
Dim bd, bdc As DataBase
Dim tb, tba, tbc As Recordset

Dim WCodItem As Long
Dim Item_Existente As Boolean


Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set bdc = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tbc = bdc.OpenRecordset("Tabela_VDO_CSV", DB_OPEN_TABLE)
Set tb = bd.OpenRecordset("EST", DB_OPEN_TABLE)
Set tba = bd.OpenRecordset("EMPXEST", DB_OPEN_TABLE)

tbc.index = "PrimaryKey"
tba.index = "PrimaryKey"

If tbc.RecordCount <> 0 Then
   While Not tbc.EOF
         If tbc!Selecionar Then
            GoSub Valida_Codigo_Fabricante:
            'If Item_Existente = True Then
            '   GoSub Atualiza_Preco_EMPXEST:
            'End If
         End If
         tbc.MoveNext
   Wend
End If

Finaliza:
   tbc.Close
   tba.Close
   tb.Close
   bd.Close
   bdc.Close
   Set tbc = Nothing
   Set tba = Nothing
   Set tb = Nothing
   Set bd = Nothing
   Set bdc = Nothing
   DoCmd.Hourglass False
   Exit Function
Return

Valida_Codigo_Fabricante:
   tb.index = "CodigoFabricante"
   tb.Seek "=", tbc!CodigoFabricante
   If tb.NoMatch Then
      Item_Existente = False
      Return
    Else
      Item_Existente = True
      While tbc!CodigoFabricante = tb!CodigoFabricante
            If Not IsNull(tbc!CodigoBarras) Then
               If IsNull(tb!CodBarra) Then
                  BeginTrans
                  tb.Edit
                  tb!CodBarra = Str(tbc!CodigoBarras)
                  tb.Update
                  CommitTrans
               End If
            End If
            GoSub Atualiza_Preco_EMPXEST:
            tb.MoveNext
            If tb.EOF Then
               Return
            End If
      Wend
    End If
Return

Atualiza_Preco_EMPXEST:
   BeginTrans
   tba.Seek "=", Val(Form_Menu.WEmpresa), tb!CodItem
   If Not tba.NoMatch Then
      tba.Edit
      If tba!Ativo Then
         tba!PrecoVenda = tbc!PrecoVenda
      End If
      tba.Update
   End If
   CommitTrans
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume
    End If


End Function

Function Incluir_Clientes_Planilha_CSV() As Long

DoCmd.Hourglass True
On Error GoTo TratamentodeErro:
Dim bd, bdc As DataBase
Dim tb, tbb, tba, tbc As Recordset

Dim WCodigoCliente As Long
Dim Cli_Existente As Boolean
Dim X As Long
Dim Ret As Variant
Dim numreg As Long

'CAMPOS TABELA

Dim WRazaoSocial As String
Dim WCodinome As String
Dim WTipo As String
Dim WInscricao As String
Dim wCGCCPF As String
Dim WSegmento As Long
Dim WTabela As Long
Dim WEndereco As String
Dim WBairro As String
Dim WEndPostal As String
Dim WCidade As String
Dim WEstado As String
Dim wTelefone As String
Dim WFax As String
Dim WEmpresa As Long
Dim WCodigoTransportadora As String
Dim WEndElet As String
Dim WEnderecoCobranca As String
Dim WBairroCobranca As String
Dim WEndPostalCobranca As String
Dim WCidadeCobranca As String
Dim WEstadoCobranca As String
Dim WFoneCobranca As String
Dim WCodigoVendedor As String
Dim WTipoAbordagem As String
Dim WDDD As String
Dim WCelular As String
Dim wnro As Long
Dim WComplemento As String
Dim WInscMunic As String
Dim WTabelaServ  As Long
Dim WCodigoPagto As String
Dim wobservacoes As String
Dim WCodigo_Cli_Antigo As String

Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set bdc = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tbc = bdc.OpenRecordset("Tabela_CLI_CSV", DB_OPEN_TABLE)
Set tb = bd.OpenRecordset("CLI", DB_OPEN_TABLE)
Set tba = bd.OpenRecordset("BAI", DB_OPEN_TABLE)
Set tbb = bd.OpenRecordset("CID", DB_OPEN_TABLE)
tbc.index = "PrimaryKey"
tba.index = "PrimaryKey"
tbc.index = "PrimaryKey"
tb.index = "PrimaryKey"

X = 0

BeginTrans
If tbc.RecordCount <> 0 Then
   numreg = tbc.RecordCount
   Ret = SysCmd(acSysCmdInitMeter, "Aguarde, importando registros...", numreg)
   While Not tbc.EOF
         'If Tbc!Selecionar Then
            GoSub Valida_CNPJCPF:
            If Cli_Existente = False Then
               GoSub Monta_Registro:
               GoSub Cria_Bairro:
               GoSub Cria_Cliente:
            End If
         'End If
         X = X + 1
          Ret = SysCmd(SYSCMD_UPDATEMETER, X)
          If X Mod 200 = 0 Then
            CommitTrans
            BeginTrans
         End If
         tbc.MoveNext
   Wend
End If

Finaliza:
   CommitTrans
   Ret = SysCmd(acSysCmdRemoveMeter)
   tbc.Close
   tba.Close
   tb.Close
   bd.Close
   bdc.Close
   Set tbc = Nothing
   Set tba = Nothing
   Set tb = Nothing
   Set bd = Nothing
   Set bdc = Nothing
   DoCmd.Hourglass False
   Exit Function
Return

Cria_Cliente:
   WCodigoCliente = ProximoCodigoCliente(1)
   tb.Seek "=", WCodigoCliente
   If tb.NoMatch Then
      tb.AddNew
   Else
      tb.Edit
   End If
   tb!CodigoCliente = WCodigoCliente
   tb!RazaoSocial = tbc!RazaoSocial
   tb!Codinome = tbc!Codinome
   tb!Tipo = WTipo
   tb!Inscricao = tbc!Inscricao
   tb!CgcCpf = Mid$(tbc!CgcCpf, 1, 14)
   tb!Segmento = 8
   tb!tABELA = 1
   tb!Endereco = tbc!Endereco
   tb!Bairro = tbc!Bairro
   If Len(tbc!EndPostal) > 8 Then
      tb!EndPostal = Cep()
   Else
      tb!EndPostal = Mid$(tbc!EndPostal, 1, 8)
   End If
   tb!Cidade = tbc!Cidade
   If Len(tbc!Estado) > 2 Then
      tb!Estado = Estado()
   Else
      tb!Estado = Mid$(tbc!Estado, 1, 2)
   End If
   tb!Telefone = Mid$(tbc!Telefone, 1, 14)
   tb!Empresa = 1
   tb!CodigoTransportadora = 0
   tb!EndElet = tbc!EndElet
   tb!EnderecoCobranca = tbc!Endereco
   tb!BairroCobranca = tbc!Bairro
   tb!EndPostalCobranca = tbc!EndPostal
   tb!CidadeCobranca = tbc!Cidade
   If Len(tbc!Estado) > 2 Then
      tb!EstadoCobranca = Estado()
   Else
      tb!EstadoCobranca = tbc!Estado
   End If
   tb!FoneCobranca = 0
   tb!CodigoVendedor = 0
   tb!TipoAbordadem = "IND"
   tb!DDD = tbc!DDD
   tb!Celular = tbc!Celular
   If IsNumeric(tbc!Nro) Then
      tb!Nro = tbc!Nro
   End If
   tb!Complemento = tbc!Complemento
   tb!InscMunic = tbc!InscMunic
   tb!TabelaServ = 1
   tb!CodigoPagto = 0
   tb!Observacoes = tbc!Observacoes
   tb!CodigoPais = 1058
   tb!CodigoCli_Antigo = tbc!CodigoCli_Antigo
   tb.Update
Return

Cria_Cidade:
  tbb.Seek "=", WCidade
  If tbb.NoMatch Then
     tbb.AddNew
  Else
     tbb.Edit
  End If
  tbb!DescricaoCidade = WCidade
  tbb.Update
Return


Cria_Bairro:
  tba.Seek "=", WBairro
  If tba.NoMatch Then
     tba.AddNew
  Else
     tba.Edit
  End If
  tba!DescricaoBAirro = WBairro
  tba.Update
Return

Valida_CNPJCPF:
   tb.index = "CGCCPF"
   tb.Seek "=", tbc!CgcCpf
   If tb.NoMatch Then
      Cli_Existente = False
    Else
      Cli_Existente = True
    End If
Return

Monta_Registro:

   WRazaoSocial = tbc!RazaoSocial
   WCodinome = tbc!Codinome
   If Len(tbc!CgcCpf) = 11 Then
      WTipo = 1
   Else
      WTipo = 1
   End If
   WInscricao = tbc!Inscricao
   wCGCCPF = tbc!CgcCpf
   WSegmento = 8
   WTabela = 1
   WEndereco = tbc!Endereco
   WBairro = tbc!Bairro
   WEndPostal = tbc!EndPostal
   WCidade = tbc!Cidade
   WEstado = tbc!Estado
   wTelefone = tbc!Telefone
   WFax = tbc!Telefone
   WEmpresa = 1
   WCodigoTransportadora = 0
   WEndElet = tbc!EndElet
   WEnderecoCobranca = tbc!Endereco
   WBairroCobranca = tbc!Bairro
   WEndPostalCobranca = tbc!EndPostal
   WCidadeCobranca = tbc!Cidade
   WEstadoCobranca = tbc!Estado
   WFoneCobranca = 0
   WCodigoVendedor = 0
   WTipoAbordagem = "IND"
   WDDD = tbc!DDD
   If IsNull(tbc!Celular) Then
      WCelular = 0
   Else
      WCelular = tbc!Celular
   End If
   If IsNumeric(tbc!Nro) Then
      wnro = Val(tbc!Nro)
   End If
   WComplemento = tbc!Complemento
   WInscMunic = tbc!InscMunic
   WTabelaServ = 1
   WCodigoPagto = 0
   wobservacoes = tbc!Observacoes
   WCodigo_Cli_Antigo = tbc!CodigoCli_Antigo
Return


TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Importa_Planilha_VDO_CSV(wNomeArquivo As String) As Boolean
DoCmd.Hourglass True
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset

Dim Ret As Variant
Dim X As Long
Dim I As Long
Dim WLog As Long
Dim WSeq As String * 1
Dim wult_campo As Long
Dim WVlr_Campo As String
Dim WColuna As Long
Dim numreg As Long

Dim RegTXT As String * 500

' CAMPOS DA TABELA

Dim WCodigoFabricante As String
Dim WDescricao As String
Dim wLote As String
Dim WCodigoNCM As String
Dim wCodigoBarras As String
Dim WPeso As Double
Dim WUnidade As String
Dim WPrecovenda As Double

Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Tabela_VDO_CSV", DB_OPEN_TABLE)

tb.index = "PrimaryKey"

Open wNomeArquivo For Input As #1
BeginTrans
GoSub Inicializa_Tabela_Aux:
CommitTrans

BeginTrans
X = 0
wult_campo = 0
numreg = 11300
Ret = SysCmd(SYSCMD_INITMETER, "Aguarde, Atualizando Tabela " & UCase(wNomeArquivo) & " ...", numreg)
While Not EOF(1)

      Line Input #1, RegTXT ' L� a linha para a vari�vel.
   
      X = X + 1
      If X > 1 Then
        GoSub Efetua_Leitura:
      End If
      
      Ret = SysCmd(SYSCMD_UPDATEMETER, X)
      If X Mod 500 = 0 Then
         CommitTrans
         BeginTrans
      End If
Wend

Finaliza:
   CommitTrans
   Close #1
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
  
   DoCmd.Hourglass False
   Ret = SysCmd(SYSCMD_REMOVEMETER)
 '  Call Rodape_SGa
   Exit Function
Return

Inicializa_Tabela_Aux:
  If tb.RecordCount <> 0 Then
     While Not tb.EOF
           tb.Delete
           tb.MoveNext
     Wend
  End If
Return

Efetua_Leitura:
  For I = 1 To Len(RegTXT)
      If wult_campo = 0 Then
         wult_campo = 1
      End If
          
      If Mid$(RegTXT, I, 1) = ";" Then
         WColuna = WColuna + 1
         WVlr_Campo = Mid$(RegTXT, wult_campo, (I - wult_campo))
         If WVlr_Campo <> ";" Then
            GoSub MontaRegistro:
         End If
         wult_campo = I + 1
       End If
   Next I
   WColuna = WColuna + 1
   
   If Mid$(RegTXT, 1, 1) = "" Or IsNull(Mid$(RegTXT, 1, 1)) Then
      GoSub Finaliza:
   End If
   WVlr_Campo = Trim(Mid$(RegTXT, wult_campo, (Len(RegTXT) - wult_campo)))
   GoSub MontaRegistro:
   GoSub GravaRegistro:
   WColuna = 0
   wult_campo = 0
Return

MontaRegistro:
  
  If WVlr_Campo <> "" Then
     If WColuna = 1 Then
        WCodigoFabricante = WVlr_Campo
     ElseIf WColuna = 2 Then
            WDescricao = WVlr_Campo
     ElseIf WColuna = 3 Then
            wLote = WVlr_Campo
     ElseIf WColuna = 4 Then
            WCodigoNCM = WVlr_Campo
     ElseIf WColuna = 5 Then
            wCodigoBarras = WVlr_Campo
     ElseIf WColuna = 6 Then
            WPeso = WVlr_Campo
     ElseIf WColuna = 7 Then
            WUnidade = WVlr_Campo
     ElseIf WColuna = 8 Then
            If IsNumeric(WVlr_Campo) Then
                If WVlr_Campo <> 0 Then
                   WPrecovenda = WVlr_Campo
                Else
                   WPrecovenda = 0.001
                End If
            End If
     End If
  End If
Return

GravaRegistro:
      tb.Seek "=", WCodigoFabricante
      If tb.NoMatch Then
         tb.AddNew
      Else
         tb.Edit
      End If
     tb!CodigoFabricante = WCodigoFabricante
     tb!Descricao = WDescricao
     tb!Lote = wLote
     tb!Codigoncm = WCodigoNCM
     tb!CodigoBarras = wCodigoBarras
     tb!Peso = WPeso
     tb!Unidade = WUnidade
     tb!PrecoVenda = WPrecovenda
     tb.Update
Return '

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume
    End If

End Function

Function Existe_Clientes_Importar() As Boolean

'DoCmd.hourglass True
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset

Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Tabela_CLI_CSV", DB_OPEN_TABLE)

tb.index = "PrimaryKey"


If tb.RecordCount <> 0 Then
   Existe_Clientes_Importar = True
Else
   Existe_Clientes_Importar = False
End If

Finaliza:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
  
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume
    End If


End Function

Function Importa_Planilha_CLI_CSV(wNomeArquivo As String, Wcidade_A As String, WEstado_A As String) As Boolean

DoCmd.Hourglass True
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset

Dim Ret As Variant
Dim X As Long
Dim I As Long
Dim WLog As Long
Dim WSeq As String * 1
Dim wult_campo As Long
Dim WVlr_Campo As String
Dim WColuna As Long
Dim numreg As Long

Dim RegTXT As String * 500

' CAMPOS DA TABELA

Dim WRazaoSocial As String
Dim WCodinome As String
Dim WInscricao As String
Dim wCGCCPF As String
Dim WInscMunic As String
Dim WEndereco As String
Dim wnro As String
Dim WComplemento As String
Dim WBairro As String
Dim WEndPostal As String
Dim WCidade As String
Dim WEstado As String
Dim wTelefone As String
Dim WEndElet As String
Dim WDDD As String
Dim WCelular As String
Dim WContato As String
Dim wobservacoes As String
Dim WCodigoCli_Antigo As String

Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Tabela_CLI_CSV", DB_OPEN_TABLE)

tb.index = "PrimaryKey"

Open wNomeArquivo For Input As #1
BeginTrans
GoSub Inicializa_Tabela_Aux:
CommitTrans

BeginTrans
X = 0
wult_campo = 0
numreg = 11300
Ret = SysCmd(SYSCMD_INITMETER, "Aguarde, Atualizando Tabela " & UCase(wNomeArquivo) & " ...", numreg)
While Not EOF(1)

      Line Input #1, RegTXT ' L� a linha para a vari�vel.
    
      X = X + 1
      
      If X > 1 Then
        GoSub Inicializa_Campos
        GoSub Efetua_Leitura:
      End If
      
      Ret = SysCmd(SYSCMD_UPDATEMETER, X)
      If X Mod 500 = 0 Then
         CommitTrans
         BeginTrans
      End If
Wend

Finaliza:
   CommitTrans
   Close #1
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
  
   DoCmd.Hourglass False
   Ret = SysCmd(SYSCMD_REMOVEMETER)
 '  Call Rodape_SGa
   Exit Function
Return

Inicializa_Campos:
 WInscricao = "NAO FORNECIDO"
 WInscMunic = "NAO FORNECIDO"
 wnro = 0
 WComplemento = Space(1)
 WBairro = "INDEFINIDO"
 WEndPostal = 0
 WCidade = Wcidade_A
 WEstado = WEstado_A
 wTelefone = 0
 WEndElet = Space(1)
 WDDD = "0"
 WCelular = 0
 WContato = Space(1)
 wobservacoes = Space(1)
 WCodigoCli_Antigo = "SPACE(1)"
Return


Inicializa_Tabela_Aux:
  If tb.RecordCount <> 0 Then
     While Not tb.EOF
           tb.Delete
           tb.MoveNext
     Wend
  End If
Return

Efetua_Leitura:
  For I = 1 To Len(RegTXT)
      If wult_campo = 0 Then
         wult_campo = 1
      End If
          
      If Mid$(RegTXT, I, 1) = ";" Then
         WColuna = WColuna + 1
         WVlr_Campo = Mid$(RegTXT, wult_campo, (I - wult_campo))
         If WVlr_Campo <> ";" Then
            GoSub MontaRegistro:
         End If
         wult_campo = I + 1
       End If
   Next I
   WColuna = WColuna + 1
   
   If Mid$(RegTXT, 1, 1) = "" Or IsNull(Mid$(RegTXT, 1, 1)) Then
      GoSub Finaliza:
   End If
   WVlr_Campo = Trim(Mid$(RegTXT, wult_campo, (Len(RegTXT) - wult_campo)))
   GoSub MontaRegistro:
   GoSub GravaRegistro:
   WColuna = 0
   wult_campo = 0
Return

MontaRegistro:
  
  If WVlr_Campo <> "" Then
    'If WVlr_Campo = "727" Then
    '   Stop
    'End If
     If WColuna = 1 Then
        WRazaoSocial = UCase(WVlr_Campo)
     ElseIf WColuna = 2 Then
            WCodinome = UCase(WVlr_Campo)
     ElseIf WColuna = 3 Then
            If Len(WVlr_Campo) > 0 Then
               WInscricao = UCase(WVlr_Campo)
            End If
     ElseIf WColuna = 4 Then
            wCGCCPF = UCase(WVlr_Campo)
     ElseIf WColuna = 5 Then
            If Len(WVlr_Campo) > 0 Then
               WInscMunic = WVlr_Campo
            End If
     ElseIf WColuna = 6 Then
            If Len(WVlr_Campo) > 0 Then
                WEndereco = UCase(WVlr_Campo)
            End If
     ElseIf WColuna = 7 Then
            If IsNumeric(WVlr_Campo) Then
               wnro = WVlr_Campo
            End If
     ElseIf WColuna = 8 Then
            If Len(WVlr_Campo) > 0 Then
               WComplemento = WVlr_Campo
            End If
     ElseIf WColuna = 9 Then
            WBairro = UCase(WVlr_Campo)
     ElseIf WColuna = 10 Then
            WEndPostal = WVlr_Campo
     ElseIf WColuna = 11 Then
            WCidade = WVlr_Campo
     ElseIf WColuna = 12 Then
            WEstado = WVlr_Campo
     ElseIf WColuna = 13 Then
            If Len(WVlr_Campo) > 0 Then
               wTelefone = WVlr_Campo
            End If
     ElseIf WColuna = 14 Then
            If Len(WVlr_Campo) > 0 Then
               WEndElet = LCase(WVlr_Campo)
            End If
     ElseIf WColuna = 15 Then
            WDDD = WVlr_Campo
     ElseIf WColuna = 16 Then
            If IsNumeric(WColuna) Then
               WCelular = WVlr_Campo
            End If
     ElseIf WColuna = 17 Then
            If Len(WVlr_Campo) > 0 Then
               WContato = UCase(WVlr_Campo)
            End If
     ElseIf WColuna = 18 Then
            If Len(WVlr_Campo) > 0 Then
               wobservacoes = WVlr_Campo
            End If
     ElseIf WColuna = 19 Then
            If Len(WVlr_Campo) > 0 Then
               WCodigoCli_Antigo = WVlr_Campo
            End If
     End If
  End If
Return

GravaRegistro:
     
      tb.Seek "=", wCGCCPF
      If tb.NoMatch Then
         tb.AddNew
      Else
         tb.Edit
      End If
     tb!CgcCpf = wCGCCPF
     tb!RazaoSocial = Mid$(WRazaoSocial, 1, 255)
     tb!Codinome = Trim$(Mid$(WCodinome, 1, 255))
     tb!Inscricao = Mid$(WInscricao, 1, 20)
     tb!CgcCpf = wCGCCPF
     tb!InscMunic = Mid$(WInscMunic, 1, 50)
     tb!Endereco = Mid$(WEndereco, 1, 255)
     tb!Nro = wnro
     If WComplemento <> "" Then
        tb!Complemento = WComplemento
     Else
        tb!Complemento = "*"
     End If
     tb!Bairro = WBairro
     tb!EndPostal = WEndPostal
     tb!Cidade = WCidade
     tb!Estado = WEstado
     tb!Telefone = wTelefone
     tb!EndElet = WEndElet
     tb!DDD = WDDD
     If IsNumeric(WCelular) Then
        tb!Celular = WCelular
     End If
     tb!Contato = WContato
     tb!Observacoes = wobservacoes
     tb!CodigoCli_Antigo = WCodigoCli_Antigo
     tb.Update
Return '

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume
    End If
End Function


Function Proximo_Compart_Tanque_Descontam_Disp(WNumos As Long, WNum_Tanque As String) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Dim WCompart
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("DESCONTAM_TANQUE_OSV_COMPART", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

WCompart = 1

tb.Seek "=", WNumos, WNum_Tanque, WCompart

If tb.NoMatch Then
   GoSub Finaliza:
Else
   While WCompart < 100
         WCompart = WCompart + 1
         tb.Seek "=", WNumos, WNum_Tanque, WCompart
         If tb.NoMatch Then
            GoSub Finaliza
         End If
   Wend
End If


Finaliza:
    Proximo_Compart_Tanque_Descontam_Disp = WCompart
    tb.Close
    bd.Close
    Set tb = Nothing
    Set bd = Nothing
  
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaPrazoOSV(Numos As Long) As String
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim OSV, PRZ As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set OSV = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
Set PRZ = bd1.OpenRecordset("CONDPAGTO", DB_OPEN_TABLE)
OSV.index = "PrimaryKey"
PRZ.index = "PrimaryKey"
'BeginTrans
OSV.Seek "=", Numos
If Not OSV.NoMatch Then
   PRZ.Seek "=", OSV!CodigoPagto
   If Not PRZ.NoMatch Then
      CapturaPrazoOSV = Mid$(PRZ!Descricao, 1, 16)
   End If
End If
'CommitTrans
PRZ.Close
OSV.Close
bd1.Close
Set PRZ = Nothing
Set OSV = Nothing
Set bd1 = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Proxima_Sequencia_Imagem_Vinc(WNumos As Long, WSequencia As Long) As Long

On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OSV_IMAGENS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

Do Until tb.NoMatch
     WSequencia = WSequencia + 1
     tb.Seek "=", WNumos, WSequencia
Loop

Finaliza:
  Proxima_Sequencia_Imagem_Vinc = WSequencia
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function Check_ImagemXNumOs(WNumos As Long) As Boolean

On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OSV_IMAGENS", DB_OPEN_TABLE)
tb.index = "NumOS"
tb.Seek "=", WNumos
Check_ImagemXNumOs = True
If Not tb.NoMatch Then
   While WNumos = tb!Numos
      Dim ofuncoes As CanadaConnectivity.FuncoesDiversas
      Set ofuncoes = New CanadaConnectivity.FuncoesDiversas
      If Not IsNull(tb!Caminho_Imagem) Then
         If ofuncoes.CheckFile(tb!Caminho_Imagem) = False Then
            MsgBox "Verifique as Imagens Anexadas e Tente Novamente.", vbCritical
            Check_ImagemXNumOs = False
            GoTo Finaliza
         End If
      Else
         MsgBox "Verifique as Imagens Anexadas e Tente Novamente.", vbCritical
         Check_ImagemXNumOs = False
         GoTo Finaliza
      End If
      tb.MoveNext
      If tb.EOF Then
         GoTo Finaliza
      End If
   Wend
Else
   MsgBox "Verifique as Imagens Anexadas e Tente Novamente.", vbCritical
   Check_ImagemXNumOs = False
   GoTo Finaliza
End If
Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Retorna_ImagemXNumOsxSeq(WNumos As Long, WSeq As Integer) As String

On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OSV_IMAGENS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", WNumos, WSeq
If Not tb.NoMatch Then
   If Not IsNull(tb!Caminho_Imagem) Then
      Retorna_ImagemXNumOsxSeq = tb!Caminho_Imagem
   Else
      Retorna_ImagemXNumOsxSeq = 0
   End If
Else
   Retorna_ImagemXNumOsxSeq = 0
End If
Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If



End Function

Function RetornaValueAroPneuVeiculo(WDescricao As String) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("AROPNEU", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", WDescricao
If Not tb.NoMatch Then
   If Not IsNull(tb!Value) Then
      RetornaValueAroPneuVeiculo = tb!Value
   Else
      RetornaValueAroPneuVeiculo = 0
   End If
End If
Finaliza:
  tb.Close
  bd.Close
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Captura_Campos_PlacXCli(WCodigoCliente As Long, WPlacaVeiculo As String, WCampo) As Variant
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("PLACXCLI", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

tb.Seek "=", WPlacaVeiculo, WCodigoCliente

If Not tb.NoMatch Then
   If tb.Fields(WCampo).Value <> "" Or Not IsNull(tb.Fields(WCampo).Value) Then
       Captura_Campos_PlacXCli = tb.Fields(WCampo).Value
   Else
      Captura_Campos_PlacXCli = 0
   End If
   
Else
   Captura_Campos_PlacXCli = 0
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Captura_Campos_TacografosOS(WNumos As Long, WCampo As String) As String
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("TACOGRAFOsOS", DB_OPEN_TABLE)
tb.index = "SecondKey"

tb.Seek "=", WNumos

If Not tb.NoMatch Then
   If tb.Fields(WCampo).Value <> "" Or Not IsNull(tb.Fields(WCampo).Value) Then
       Captura_Campos_TacografosOS = tb.Fields(WCampo).Value
   Else
      Captura_Campos_TacografosOS = 0
   End If
   
Else
   Captura_Campos_TacografosOS = 0
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Prev_Rec_Cont_Vlr_OSV_Abert(WFlag As Long) As Double
On Error GoTo TratamentodeErro
Dim db As DataBase, tb As Recordset, Qtb As QueryDef
Dim WTotal As Double
Set db = CurrentDb()
Set Qtb = db.QueryDefs("Prev_Rec_Cont_OSV_Abertas")
Set tb = Qtb.OpenRecordset()

WTotal = 0

If tb.RecordCount <> 0 Then
   While Not tb.EOF
         
          If WFlag = 1 Then
             WTotal = WTotal + tb!Total_Produtos
          ElseIf WFlag = 2 Then
                 WTotal = WTotal + tb!Total_Servicos
          Else
             WTotal = WTotal + tb!Total_Servicos + tb!Total_Produtos
          End If
          tb.MoveNext
          
   Wend
End If

Finaliza:
  Prev_Rec_Cont_Vlr_OSV_Abert = Round(WTotal)
  tb.Close
  db.Close
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Total_HRS_TRAB_EMPRESA_MES(Wmes As Long) As Long
On Error GoTo TratamentodeErro
Dim bd As DataBase, tbf As Recordset
Dim WTotal As Long
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tbf = bd.OpenRecordset("HRS_TRAB_TECNICOS", DB_OPEN_TABLE)
tbf.index = "Codigo"

If tbf.RecordCount <> 0 Then
   While Not tbf.EOF
         If Wmes = 1 Then
            WTotal = WTotal + tbf!Jan
         ElseIf Wmes = 2 Then
                WTotal = WTotal + tbf!Fev
         ElseIf Wmes = 3 Then
                WTotal = WTotal + tbf!Mar
         ElseIf Wmes = 4 Then
                WTotal = WTotal + tbf!Abr
         ElseIf Wmes = 5 Then
                WTotal = WTotal + tbf!Mai
         ElseIf Wmes = 6 Then
                WTotal = WTotal + tbf!Jun
         ElseIf Wmes = 7 Then
                WTotal = WTotal + tbf!Jul
         ElseIf Wmes = 8 Then
                WTotal = WTotal + tbf!Ago
         ElseIf Wmes = 9 Then
                WTotal = WTotal + tbf!Set
         ElseIf Wmes = 10 Then
                WTotal = WTotal + tbf!OUT
         ElseIf Wmes = 11 Then
                WTotal = WTotal + tbf!Nov
         ElseIf Wmes = 12 Then
                WTotal = WTotal + tbf!Dez
  
         End If
         
         tbf.MoveNext
         If tbf.EOF Then
            GoSub Finaliza:
         End If

   Wend
End If

Finaliza:
   Total_HRS_TRAB_EMPRESA_MES = WTotal
  tbf.Close
  bd.Close
  Exit Function
Return


TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Total_HRS_TRAB() As Long
On Error GoTo TratamentodeErro
Dim bd As DataBase, tbf As Recordset
Dim WTotal As Long
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tbf = bd.OpenRecordset("HRS_TRAB_TECNICOS", DB_OPEN_TABLE)
tbf.index = "Codigo"

If tbf.RecordCount <> 0 Then
   While Not tbf.EOF
         WTotal = WTotal + (tbf!Jan + tbf!Fev + tbf!Mar + tbf!Abr + tbf!Mai + tbf!Jun + tbf!Jul + tbf!Ago + tbf!Set + tbf!OUT + tbf!Nov + tbf!Dez)
         tbf.MoveNext
         If tbf.EOF Then
            GoSub Finaliza:
         End If

   Wend
End If

Finaliza:
   Total_HRS_TRAB = WTotal
  tbf.Close
  bd.Close
  Exit Function
Return


TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Total_HRS_TRAB_Tecnico(WTec_Exec As Long) As Long
On Error GoTo TratamentodeErro
Dim bd As DataBase, tbf As Recordset
Dim WTotal As Long
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tbf = bd.OpenRecordset("HRS_TRAB_TECNICOS", DB_OPEN_TABLE)
tbf.index = "Codigo"

tbf.Seek "=", WTec_Exec
If Not tbf.NoMatch Then
   While tbf!Codigo = WTec_Exec
         WTotal = WTotal + (tbf!Jan + tbf!Fev + tbf!Mar + tbf!Abr + tbf!Mai + tbf!Jun + tbf!Jul + tbf!Ago + tbf!Set + tbf!OUT + tbf!Nov + tbf!Dez)
         tbf.MoveNext
         If tbf.EOF Then
            GoSub Finaliza:
         End If

   Wend
End If

Finaliza:
   Total_HRS_TRAB_Tecnico = WTotal
  tbf.Close
  bd.Close
  Exit Function
Return


TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Total_HRS_TRAB_SERV_Anual_Tecnico(WTec_Exec As Long, WTipoServico As Long) As Long
On Error GoTo TratamentodeErro
Dim bd As DataBase, tbf As Recordset
Dim WTotal As Long
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tbf = bd.OpenRecordset("HRS_TRAB_TECNICOS", DB_OPEN_TABLE)
tbf.index = "PrimaryKey"

tbf.Seek "=", WTec_Exec, WTipoServico
If Not tbf.NoMatch Then
   While tbf!Codigo = WTec_Exec And tbf!TipoServico = WTipoServico
         WTotal = WTotal + (tbf!Jan + tbf!Fev + tbf!Mar + tbf!Abr + tbf!Mai + tbf!Jun + tbf!Jul + tbf!Ago + tbf!Set + tbf!OUT + tbf!Nov + tbf!Dez)
         tbf.MoveNext
         If tbf.EOF Then
            GoSub Finaliza:
         End If

   Wend
End If

Finaliza:
   Total_HRS_TRAB_SERV_Anual_Tecnico = WTotal
  tbf.Close
  bd.Close
  Set tbf = Nothing
  Set bd = Nothing
  
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function
Function Demonstrativo_Servicos_HR_TRAB_Tecnicos() As Long
On Error GoTo TratamentodeErro
Dim bd, bdt As DataBase, tbf, tb As Recordset, Qtb As QueryDef
Dim X As Long
Set bd = CurrentDb()
Set bdt = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tbf = bdt.OpenRecordset("HRS_TRAB_TECNICOS", DB_OPEN_TABLE)
Set Qtb = bd.QueryDefs("Demonstrativo_Servicos_HR_TRAB_Tecnicos")
Qtb.Parameters(0) = [Forms]![Demonstrativo_Servicos_Executados]![Placa]
Qtb.Parameters(1) = [Forms]![Demonstrativo_Servicos_Executados]![DataInicial]
Qtb.Parameters(2) = [Forms]![Demonstrativo_Servicos_Executados]![DataFinal]
Set tb = Qtb.OpenRecordset()
tbf.index = "PrimaryKey"
DoCmd.Hourglass True
GoSub Inicializa_Tbf:

BeginTrans
If tb.RecordCount <> 0 Then
   While Not tb.EOF
         tbf.Seek "=", tb!Tec_Exec, tb!TipoServico
         If tbf.NoMatch Then
            tbf.AddNew
         Else
            tbf.Edit
         End If
         tbf!Codigo = tb!Tec_Exec
         tbf!TipoServico = tb!TipoServico
         If Month(tb!datafechamento) = 1 Then
            tbf!Jan = (tbf!Jan + tb!TempoRealizadoSeg)
         ElseIf Month(tb!datafechamento) = 2 Then
            tbf!Fev = (tbf!Fev + tb!TempoRealizadoSeg)
         ElseIf Month(tb!datafechamento) = 3 Then
            tbf!Mar = (tbf!Mar + tb!TempoRealizadoSeg)
         ElseIf Month(tb!datafechamento) = 4 Then
            tbf!Abr = (tbf!Abr + tb!TempoRealizadoSeg)
         ElseIf Month(tb!datafechamento) = 5 Then
            tbf!Mai = (tbf!Mai + tb!TempoRealizadoSeg)
         ElseIf Month(tb!datafechamento) = 6 Then
            tbf!Jun = (tbf!Jun + tb!TempoRealizadoSeg)
         ElseIf Month(tb!datafechamento) = 7 Then
            tbf!Jul = (tbf!Jul + tb!TempoRealizadoSeg)
         ElseIf Month(tb!datafechamento) = 8 Then
            tbf!Ago = (tbf!Ago + tb!TempoRealizadoSeg)
         ElseIf Month(tb!datafechamento) = 9 Then
            tbf!Set = (tbf!Set + tb!TempoRealizadoSeg)
         ElseIf Month(tb!datafechamento) = 10 Then
            tbf!OUT = (tbf!OUT + tb!TempoRealizadoSeg)
         ElseIf Month(tb!datafechamento) = 11 Then
            tbf!Nov = (tbf!Nov + tb!TempoRealizadoSeg)
         ElseIf Month(tb!datafechamento) = 12 Then
            tbf!Dez = (tbf!Dez + tb!TempoRealizadoSeg)
         End If
         tbf.Update
         X = X + 1
         If X Mod 500 = 0 Then
            CommitTrans
            BeginTrans
         End If
         tb.MoveNext
          
   Wend
End If

Finaliza:
  DoCmd.Hourglass False
  CommitTrans
  tb.Close
  tbf.Close
  bd.Close
  bdt.Close
  Set tb = Nothing
  Set tbf = Nothing
  Set bd = Nothing
  Set bdt = Nothing
  Exit Function
Return

Inicializa_Tbf:
   If tbf.RecordCount <> 0 Then
      BeginTrans
      While Not tbf.EOF
            tbf.Delete
            tbf.MoveNext
      Wend
      CommitTrans
   End If
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Atualiza_NfeEmitida_CCusto_OSV(WNumos As Long) As Long

On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"


BeginTrans
tb.Seek "=", WNumos

If Not tb.NoMatch Then
   tb.Edit
   tb!Emitirnf = "1"
   tb!NFEmitida = 1
   tb.Update
End If
CommitTrans
Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

'Function Atualiza_Campos_PlacXCli(WCodigoCliente As Long, wPlaca As String, WCampo As String, WValor_Campo As Variant) As Long

'On Error GoTo TratamentodeErro:

'Dim bd As DataBase
'Dim tb As Recordset
'Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
'Set tb = bd.OpenRecordset("PLACXCLI", DB_OPEN_TABLE)
'tb.Index = "PrimaryKey"


'BeginTrans
'tb.Seek "=", wPlaca, WCodigoCliente

'If Not tb.NoMatch Then
'   If WValor_Campo <> "" And Not IsNull(WValor_Campo) Then
'      tb.Edit
'      tb.Fields(WCampo).Value = WValor_Campo
'      tb.Update
'   End If
'End If
'CommitTrans
'Finaliza:
'  tb.Close
'  bd.Close
'  Set tb = Nothing
'  Set bd = Nothing
'  Exit Function
'Return
  
'TratamentodeErro:
'    If Err.Number = 3001 Then
'      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
'    Else
'      MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description, Vbcritical,"SGa" '
 '      Resume Next
 '   End If


'End Function



Function Peso_Aplic_Faixa_Serv_Exec(WValorServico As Double, WFaixa1_I As Double, WFaixa1_F As Double, WFaixa2_I As Double, WFaixa2_F As Double, WFaixa3_I As Double, WFaixa3_F As Double, WPeso_F1 As Double, WPeso_F2 As Double, WPeso_F3 As Double) As Double

If WValorServico = 0 Then
   Peso_Aplic_Faixa_Serv_Exec = 0
ElseIf WValorServico >= WFaixa1_I And WValorServico <= WFaixa1_F Then
   Peso_Aplic_Faixa_Serv_Exec = WPeso_F1
ElseIf WValorServico >= WFaixa2_I And WValorServico <= WFaixa2_F Then
       Peso_Aplic_Faixa_Serv_Exec = WPeso_F2
Else
   Peso_Aplic_Faixa_Serv_Exec = WPeso_F3
End If
End Function

Function Tempo_Execucao(WData_Ini As Variant, WHora_Ini As Variant, WData_fin As Variant, WHora_fin As Variant) As Double
Dim wsegundoshrexec As Double
If Not IsNull(WData_Ini) Then
'  wsegundoshrexec = DateDiff("n", Format$(WData_Ini, "dd/mm/yyyy") & " " & Format$(WHora_Ini, "hh:mm:ss"), WData_fin & " " & WHora_fin) * 60
  
Tempo_Execucao = Val(wsegundoshrexec)
End If
End Function


Function Captura_Depto_Resp_OSV(Numos As Long) As Long
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Numos

If Not tb.NoMatch Then
   Captura_Depto_Resp_OSV = tb!CodigoDepto
Else
   Captura_Depto_Resp_OSV = 0
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Captura_CodigoResponsavelOSV(Numos As Long) As Long
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Numos

If Not tb.NoMatch Then
   Captura_CodigoResponsavelOSV = tb!Responsavel
Else
   Captura_CodigoResponsavelOSV = 0
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function


Function Captura_NumOSFab(Numos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Numos

If Not tb.NoMatch Then
   If IsNull(tb!NUMOSFAB) Then
      Captura_NumOSFab = " "
   Else
      Captura_NumOSFab = tb!NUMOSFAB
   End If
Else
   Captura_NumOSFab = 0
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function


Function Captura_EmitirLaudo(Numos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Numos

If Not tb.NoMatch Then
   If IsNull(tb!EmitirLaudo) Then
       Captura_EmitirLaudo = "N"
   Else
      Captura_EmitirLaudo = tb!EmitirLaudo
   End If
Else
   
   Captura_EmitirLaudo = "N"
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Captura_DestacaTR(Numos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Numos

If Not tb.NoMatch Then
   If IsNull(tb!destacarTR) Then
       Captura_DestacaTR = "N"
   Else
      Captura_DestacaTR = tb!destacarTR
   End If
Else
   
   Captura_DestacaTR = "N"
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Captura_FrotaPropia(Numos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Numos

If Not tb.NoMatch Then
   If IsNull(tb!FrotaProp) Then
       Captura_FrotaPropia = "N"
   Else
      Captura_FrotaPropia = tb!FrotaProp
   End If
Else
   
   Captura_FrotaPropia = "N"
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Captura_Status_Selagem(Numos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Numos

If Not tb.NoMatch Then
   If IsNull(tb!Status_Selagem) Then
       Captura_Status_Selagem = "#"
   Else
      Captura_Status_Selagem = tb!Status_Selagem
   End If
Else
   
   Captura_Status_Selagem = "#"
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Captura_EmitirNF(Numos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Numos

If Not tb.NoMatch Then
   If IsNull(tb!Emitirnf) Then
       Captura_EmitirNF = "N"
   Else
      Captura_EmitirNF = tb!Emitirnf
   End If
Else
   
   Captura_EmitirNF = "N"
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing

  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Captura_NumRegPref(Numos As Long) As Long
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Numos

If Not tb.NoMatch Then
   If Not IsNull(tb!NumRegPref) Then
      Captura_NumRegPref = tb!NumRegPref
   End If
Else
   
   Captura_NumRegPref = 0
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing

  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Captura_OrigemOSV(Numos As Long) As Long
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Numos

If Not tb.NoMatch Then
   If IsNull(tb!Emitirnf) Then
       Captura_OrigemOSV = 0
   Else
      Captura_OrigemOSV = tb!OrigemOSV
   End If
Else
   
   Captura_OrigemOSV = 0
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing

  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Captura_CNPJ_Fabric_Simul_OSV(Numos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Numos

If Not tb.NoMatch Then
   If IsNull(tb!CNPJ_Fabricante) Then
       Captura_CNPJ_Fabric_Simul_OSV = 0
   Else
      Captura_CNPJ_Fabric_Simul_OSV = tb!CNPJ_Fabricante
   End If
Else
   
   Captura_CNPJ_Fabric_Simul_OSV = 0
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing

  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Captura_Data_Hora_Ensaio_OSV(Numos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Numos

If Not tb.NoMatch Then
   If IsNull(tb!Data_Hora_Ensaio) Then
       Captura_Data_Hora_Ensaio_OSV = 0
   Else
      Captura_Data_Hora_Ensaio_OSV = tb!Data_Hora_Ensaio
   End If
Else
   
   Captura_Data_Hora_Ensaio_OSV = 0
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing

  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function
Function Captura_Data_Hora_Envio_Simul_OSV(Numos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Numos

If Not tb.NoMatch Then
   If IsNull(tb!Data_Hora_Envio_Simulador) Then
       Captura_Data_Hora_Envio_Simul_OSV = 0
   Else
      Captura_Data_Hora_Envio_Simul_OSV = tb!Data_Hora_Envio_Simulador
   End If
Else
   
   Captura_Data_Hora_Envio_Simul_OSV = 0
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing

  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Captura_PNEU_Ensaio_Simul_OSV(Numos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Numos

If Not tb.NoMatch Then
   If IsNull(tb!Pneu_Ensaio) Then
       Captura_PNEU_Ensaio_Simul_OSV = 0
   Else
      Captura_PNEU_Ensaio_Simul_OSV = tb!Pneu_Ensaio
   End If
Else
   
   Captura_PNEU_Ensaio_Simul_OSV = 0
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing

  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Captura_AROPNEU_Ensaio_Simul_OSV(Numos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Numos

If Not tb.NoMatch Then
   If IsNull(tb!AroPneu_Ensaio) Then
       Captura_AROPNEU_Ensaio_Simul_OSV = 0
   Else
      Captura_AROPNEU_Ensaio_Simul_OSV = tb!AroPneu_Ensaio
   End If
Else
   
   Captura_AROPNEU_Ensaio_Simul_OSV = 0
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing

  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Captura_NroSerie_Ensaio_Simul_OSV(Numos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Numos

If Not tb.NoMatch Then
   If IsNull(tb!NroSerieSimulador) Then
       Captura_NroSerie_Ensaio_Simul_OSV = 0
   Else
      Captura_NroSerie_Ensaio_Simul_OSV = tb!NroSerieSimulador
   End If
Else
   
   Captura_NroSerie_Ensaio_Simul_OSV = 0
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing

  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Captura_Marca_Ensaio_Simul_OSV(Numos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Numos

If Not tb.NoMatch Then
   If IsNull(tb!MarcaSimulador) Then
       Captura_Marca_Ensaio_Simul_OSV = 0
   Else
      Captura_Marca_Ensaio_Simul_OSV = tb!MarcaSimulador
   End If
Else
   
   Captura_Marca_Ensaio_Simul_OSV = 0
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing

  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Captura_Modelo_Ensaio_Simul_OSV(Numos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Numos

If Not tb.NoMatch Then
   If IsNull(tb!MarcaSimulador) Then
       Captura_Modelo_Ensaio_Simul_OSV = 0
   Else
      Captura_Modelo_Ensaio_Simul_OSV = tb!ModeloSimulador
   End If
Else
   
   Captura_Modelo_Ensaio_Simul_OSV = 0
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing

  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function



Function Captura_KMEntradaOSV(Numos As Long) As Long
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Numos

If Not tb.NoMatch Then
   If IsNull(tb!KMEntrada) Then
       Captura_KMEntradaOSV = 0
   Else
      Captura_KMEntradaOSV = tb!KMEntrada
   End If
Else
   
   Captura_KMEntradaOSV = 0
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing

  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Captura_KMSaidaOSV(Numos As Long) As Long
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Numos

If Not tb.NoMatch Then
   If IsNull(tb!KMSaida) Then
       Captura_KMSaidaOSV = 0
   Else
      Captura_KMSaidaOSV = tb!KMSaida
   End If
Else
   
   Captura_KMSaidaOSV = 0
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing

  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Captura_FrotaOSV(Numos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Numos

If Not tb.NoMatch Then
   If Not IsNull(tb!FROTA) Then
   '    Captura_FrotaOSV = 0
   'Else
      Captura_FrotaOSV = tb!FROTA
   End If
Else
   
   Captura_FrotaOSV = 0
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing

  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Valida_Nro_COOCCF_OSV(WNro_COOCCF As String) As Boolean
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "COOCCF"


tb.Seek "=", WNro_COOCCF
If Not tb.NoMatch Then
   Valida_Nro_COOCCF_OSV = True
Else
   Valida_Nro_COOCCF_OSV = False
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  
  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Atualiza_Nro_COOCCF_OSV(WNumos As Long, WNro_COOCCF As String) As Long

On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

BeginTrans

tb.Seek "=", WNumos
If Not tb.NoMatch Then
   tb.Edit
   tb!COOCCF = WNro_COOCCF
   tb.Update
End If

Finaliza:
  CommitTrans
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  
  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If



End Function

Function Captura_CodigoFipe(Numos As Long) As Long
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Numos

If Not tb.NoMatch Then
   If IsNull(tb!CodigoFipe) Then
       Captura_CodigoFipe = 0
   Else
      Captura_CodigoFipe = tb!CodigoFipe
   End If
Else
   
   Captura_CodigoFipe = 0
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing

  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Captura_TipoCliente(wCodCliente As Long) As Long
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("CLI", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", wCodCliente

If Not tb.NoMatch Then
   If IsNull(tb!Tipo) Then
       Captura_TipoCliente = 2
   Else
      Captura_TipoCliente = tb!Tipo
   End If
Else
   
   Captura_TipoCliente = 2
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function



Function Captura_ValorEntradaOSV(Numos As Long) As Long
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Numos

If Not tb.NoMatch Then
   If IsNull(tb!ValorEntrada) Then
       Captura_ValorEntradaOSV = 0
   Else
      Captura_ValorEntradaOSV = tb!ValorEntrada
   End If
Else
   
   Captura_ValorEntradaOSV = 0
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing

  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function ProximaRequisicao(WNumReq) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("RQM", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
If tb.RecordCount = 0 Then
   WNumReq = 1
Else
   WNumReq = DMax("[NumRequisicao]", "RQM") + 1
End If

tb.Seek "=", WNumReq

If Not tb.NoMatch Then
   While Not tb.NoMatch
      WNumReq = WNumReq + 1
      tb.Seek "=", WNumReq
   Wend
   ProximaRequisicao = WNumReq
Else
   ProximaRequisicao = WNumReq
End If

Finaliza:
   tb.Close
   bd.Close
  Set tb = Nothing
   Set bd = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If ''

End Function

Function Captura_Proxima_Requisicao() As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Dim WReq As Long
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("EMPRESA", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

tb.Seek "=", 1
If tb.NoMatch Then
   MsgBox "Empresa n�o encontrada para capturar proxima requisicao disponivel...", vbCritical, "SGa"
Else
   WReq = tb!Proxima_Req
   BeginTrans
   tb.Edit
   tb!Proxima_Req = WReq + 1
   tb.Update
   CommitTrans
End If

Finaliza:
   Captura_Proxima_Requisicao = WReq
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If ''

End Function


Function Captura_Numos_RQM(WNumrequisicao As Long) As Long

On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("RQM", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", WNumrequisicao

If tb.NoMatch Then
   Captura_Numos_RQM = 0
Else
   Captura_Numos_RQM = tb!Numos
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function Captura_DtAfericao(Numos As Long) As Variant
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Numos

If Not tb.NoMatch Then
   If IsNull(tb!DTAfericao) Then
       Captura_DtAfericao = "" 'Alterado por Guilherme 26/07/2019
   Else
      Captura_DtAfericao = tb!DTAfericao
   End If
Else
   Captura_DtAfericao = 0
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing

  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Captura_ProtocoloSelagem(Numos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Numos

If Not tb.NoMatch Then
   If IsNull(tb!Protocolo_Selagem) Then
       Captura_ProtocoloSelagem = "" 'Alterado por Guilherme 26/07/2019
   Else
      Captura_ProtocoloSelagem = tb!Protocolo_Selagem
   End If
Else
   Captura_ProtocoloSelagem = 0
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing

  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Captura_Pulso(Numos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Numos

If Not tb.NoMatch Then
   If IsNull(tb!PULSO) Then
       Captura_Pulso = 0
   Else
      Captura_Pulso = tb!PULSO
   End If
Else
   Captura_Pulso = 0
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing

  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Captura_Volts(Numos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Numos

If Not tb.NoMatch Then
   If IsNull(tb!Volts) Then
      Captura_Volts = 0
   Else
      Captura_Volts = tb!Volts
   End If
Else
   Captura_Volts = 0
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing

  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function




Function Captura_TermoSelagemImpresso(Numos As Long) As Boolean
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Numos

If Not tb.NoMatch Then
   If IsNull(tb!TermoSelagemImpresso) Then
       Captura_TermoSelagemImpresso = False
   Else
      Captura_TermoSelagemImpresso = tb!TermoSelagemImpresso
   End If
Else
   Captura_TermoSelagemImpresso = False
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing

  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Captura_DescricaoDefeito(Numos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Numos

If Not tb.NoMatch Then
   If IsNull(tb!DESCRICAODEFEITO) Then
       Captura_DescricaoDefeito = Space(1)
   Else
      Captura_DescricaoDefeito = tb!DESCRICAODEFEITO
   End If
Else
   Captura_DescricaoDefeito = Space(1)
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing

  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function


Function Captura_ImagemVinc(Numos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Numos

If Not tb.NoMatch Then
   If IsNull(tb!imagemvinc) Then
       Captura_ImagemVinc = Space(1)
   Else
      Captura_ImagemVinc = tb!imagemvinc
   End If
Else
   Captura_ImagemVinc = Space(1)
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing

  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Captura_ImagemVinc1(Numos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Numos

If Not tb.NoMatch Then
   If IsNull(tb!ImagemVinc1) Then
       Captura_ImagemVinc1 = Space(1)
   Else
      Captura_ImagemVinc1 = tb!ImagemVinc1
   End If
Else
   Captura_ImagemVinc1 = Space(1)
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing

  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Retorna_Mascara_NumregPref(WNumos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", WNumos

If Not tb.NoMatch Then
   If Not IsNull(tb!NumRegPref) Then
      If tb!NumRegPref <> 0 Then
         Retorna_Mascara_NumregPref = "CADASTRO SERVI�O " & tb!NumRegPref
      Else
         Retorna_Mascara_NumregPref = ""
      End If
   
   End If
Else
   Retorna_Mascara_NumregPref = ""
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing

  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Atualiza_Dados_OSV(WNumos As Long, WIndice As Long, WCampo As String) As Long
On Error GoTo TratamentodeErro:

DoCmd.Hourglass True
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", WNumos
BeginTrans
If Not tb.NoMatch Then
   tb.Edit
   If WIndice = 1 Then
      tb!FLE = Val(WCampo)
   ElseIf WIndice = 2 Then
          tb!LFE = Val(WCampo)
   ElseIf WIndice = 3 Then
          tb!LCE = Val(WCampo)
   ElseIf WIndice = 4 Then
          tb!LTE = Val(WCampo)
   ElseIf WIndice = 5 Then
          tb!TLE = Val(WCampo)
   ElseIf WIndice = 6 Then
          tb!TC = Val(WCampo)
   ElseIf WIndice = 7 Then
          tb!FC = Val(WCampo)
   ElseIf WIndice = 8 Then
          tb!FLD = Val(WCampo)
   ElseIf WIndice = 9 Then
          tb!LFD = Val(WCampo)
   ElseIf WIndice = 10 Then
          tb!LCD = Val(WCampo)
   ElseIf WIndice = 11 Then
          tb!LTD = Val(WCampo)
   ElseIf WIndice = 12 Then
          tb!TLD = Val(WCampo)
   ElseIf WIndice = 13 Then
          tb!Responsavel = Val(WCampo)
   ElseIf WIndice = 14 Then
          tb!CodigoPagto = Val(WCampo)
   ElseIf WIndice = 15 Then
          tb!CodigoDepto = WCampo
   ElseIf WIndice = 16 Then
          If WCampo = "#" Then
             tb!NUMOSFAB = Null
          Else
             tb!NUMOSFAB = WCampo
          End If
   ElseIf WIndice = 17 Then
          If WCampo = "#" Then
             tb!NroRequisicao = Null
          Else
             tb!NroRequisicao = WCampo
          End If
   ElseIf WIndice = 18 Then
          If WCampo = "N" Then
             tb!TipoLaudo = Null
          End If
          tb!EmitirLaudo = WCampo
   ElseIf WIndice = 19 Then
          tb!TipoLaudo = WCampo
   ElseIf WIndice = 20 Then
          tb!destacarTR = WCampo
   ElseIf WIndice = 21 Then
          tb!FrotaProp = WCampo
   ElseIf WIndice = 22 Then
          tb!Emitirnf = WCampo
   ElseIf WIndice = 23 Then
          tb!NumRegPref = WCampo
   ElseIf WIndice = 25 Then
          tb!NumRegPref = Val(WCampo)
   ElseIf WIndice = 24 Then
          tb!OrigemOSV = WCampo
   ElseIf WIndice = 26 Then
          tb!Observacoes = UCase(WCampo)
   ElseIf WIndice = 27 Then
          tb!KMEntrada = WCampo
   ElseIf WIndice = 28 Then
          tb!KMSaida = WCampo
   ElseIf WIndice = 29 Then
          tb!PlacaVeiculo = WCampo
   ElseIf WIndice = 30 Then
          tb!MarcaVeiculo = WCampo
   ElseIf WIndice = 31 Then
          tb!AnoVeiculo = Val(WCampo)
   ElseIf WIndice = 32 Then
          tb!ModeloVeiculo = WCampo
   ElseIf WIndice = 33 Then
          If WCampo = "#" Then
             tb!RENAVAN = Null
          Else
             tb!RENAVAN = WCampo
          End If
   ElseIf WIndice = 34 Then
          If WCampo = "#" Then
             tb!CHASSI = Null
          ElseIf WCampo = "" Then
                 tb!CHASSI = Null
          Else
                tb!CHASSI = WCampo
          End If
   ElseIf WIndice = 35 Then
          tb!Pneu = WCampo
   ElseIf WIndice = 36 Then
          tb!AroPneu = WCampo
   ElseIf WIndice = 37 Then
          If WCampo = "#" Then
             tb!FROTA = Null
          Else
             tb!FROTA = WCampo
          End If
   ElseIf WIndice = 38 Then
          tb!CodigoFipe = WCampo
   ElseIf WIndice = 39 Then
          tb!ValorEntrada = WCampo
   ElseIf WIndice = 40 Then
          If IsNull(WCampo) Or WCampo = "" Or WCampo = Empty Then
          'If Not IsNumeric(WCampo) Then
             tb!GRU_Selagem = Null
          ElseIf WCampo = "#" Then
             tb!GRU_Selagem = Null
          Else
             tb!GRU_Selagem = WCampo
          End If
   ElseIf WIndice = 41 Then
          If WCampo = "#" Then
             tb!DTAfericao = Null
          Else
             tb!DTAfericao = DateValue(Format$(WCampo, "dd/mm/yyyy"))
          End If
   ElseIf WIndice = 42 Then
          If WCampo = "#" Then
             tb!Protocolo_Selagem = Null
          Else
             tb!Protocolo_Selagem = WCampo
          End If
   ElseIf WIndice = 43 Then
          If WCampo = "#" Then
             tb!PULSO = Null
          Else
             tb!PULSO = WCampo
          End If
   ElseIf WIndice = 44 Then
          If WCampo = "#" Then
             tb!Volts = Null
          Else
             tb!Volts = WCampo
          End If
   ElseIf WIndice = 45 Then
          tb!TermoSelagemImpresso = WCampo
   ElseIf WIndice = 46 Then
          If WCampo = "#" Then
             tb!TipoPistaEnsaio = Null
          Else
             tb!TipoPistaEnsaio = WCampo
          End If
   ElseIf WIndice = 47 Then
          If WCampo = "#" Then
             tb!VelMaxEns = Null
          Else
             tb!VelMaxEns = WCampo
          End If
   ElseIf WIndice = 48 Then
          If WCampo = "#" Then
             tb!CoeficienteW = Null
          Else
             tb!CoeficienteW = WCampo
          End If
   ElseIf WIndice = 49 Then
   
          If IsNull(WCampo) Or WCampo = "" Or WCampo = Empty Then
             tb!GRU_Ensaio = Null
          ElseIf WCampo = "#" Then
             tb!GRU_Ensaio = Null
          Else
             tb!GRU_Ensaio = WCampo
          End If
   ElseIf WIndice = 50 Then
          If WCampo = "#" Then
             tb!DTEnsaio = Null
          Else
             tb!DTEnsaio = DateValue(Format$(WCampo, "dd/mm/yyyy"))
          End If
   ElseIf WIndice = 51 Then
          If WCampo = "#" Then
             tb!DESCRICAODEFEITO = Null
          Else
             tb!DESCRICAODEFEITO = UCase(WCampo)
          End If
          
   ElseIf WIndice = 52 Then
          If WCampo = "#" Then
             tb!ImagVinc = Null
          Else
             tb!ImagVinc = WCampo
          End If
   ElseIf WIndice = 53 Then
          If WCampo = "#" Then
             tb!ImagVinc1 = Null
          Else
             tb!ImagVinc1 = WCampo
          End If
   ElseIf WIndice = 64 Then
          If WCampo = "#" Then
             tb!COR_Veiculo = Null
          Else
             tb!COR_Veiculo = WCampo
          End If
          
    ElseIf WIndice = 65 Then
           tb!DataAbertura = WCampo
   End If
   tb.Update
End If
DoCmd.Hourglass False
CommitTrans
Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing

  Exit Function
Return
  
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Captura_CodigoPadrao_Servico(WCodServico As Long) As String
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("TPSERV", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", WCodServico

If Not tb.NoMatch Then
   If IsNull(tb!Aplicacao) Then
      Captura_CodigoPadrao_Servico = "#"
    Else
      Captura_CodigoPadrao_Servico = tb!CodigoPadrao
    End If
Else
   Captura_CodigoPadrao_Servico = "#"
End If

tb.Close
bd.Close
Set tb = Nothing
  Set bd = Nothing

Exit Function

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function Captura_Aplicacao_Servico(WCodServico As Long) As String
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("TPSERV", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", WCodServico

If Not tb.NoMatch Then
   If IsNull(tb!Aplicacao) Then
      Captura_Aplicacao_Servico = " "
    Else
      Captura_Aplicacao_Servico = tb!Aplicacao
    End If
Else
   Captura_Aplicacao_Servico = " "
End If

tb.Close
bd.Close
Set tb = Nothing
Set bd = Nothing

Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function


Function CapturaTempoProgramado(WTipoServ As Long) As Double
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim TPS  As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set TPS = bd1.OpenRecordset("tpserv", DB_OPEN_TABLE)
TPS.index = "PrimaryKey"

TPS.Seek "=", WTipoServ

If Not TPS.NoMatch Then
   CapturaTempoProgramado = TPS!TempoMedio
End If

Finaliza:
 TPS.Close
 bd1.Close
 Set TPS = Nothing
 Set bd1 = Nothing

 Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function
Function CapturaCodigoPadraoServico(WCodigoServico As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim TPS As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set TPS = bd1.OpenRecordset("TPSERV", DB_OPEN_TABLE)
TPS.index = "PrimaryKey"
'BeginTrans
TPS.Seek "=", WCodigoServico
If TPS.NoMatch Then
   CapturaCodigoPadraoServico = "0000000000000000000"
Else
   If IsNull(TPS!TipoServico) Then
      CapturaCodigoPadraoServico = "0000000000000000000"
   Else
      CapturaCodigoPadraoServico = Format$(TPS!TipoServico, "0000000000000000000")
   End If
End If

'CommitTrans
TPS.Close
bd1.Close
Set TPS = Nothing
Set bd1 = Nothing

Exit Function

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Proximo_Codigo_Vago_Servico(WCodigoServico As Long) As Long

On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("TPSERV", DB_OPEN_TABLE)

tb.index = "PrimaryKey"
tb.Seek "=", WCodigoServico

If Not tb.NoMatch Then
   While Not tb.NoMatch
         WCodigoServico = WCodigoServico + 1
         tb.Seek "=", WCodigoServico
   Wend
   Proximo_Codigo_Vago_Servico = WCodigoServico
Else
   Proximo_Codigo_Vago_Servico = WCodigoServico
End If

Finaliza:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing

   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function


Function Captura_Genero_Item_TPServ(WCodigoServico As Long) As Long

On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim TPS As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set TPS = bd1.OpenRecordset("TPSERV", DB_OPEN_TABLE)
TPS.index = "PrimaryKey"
'BeginTrans
TPS.Seek "=", WCodigoServico
If TPS.NoMatch Then
   Captura_Genero_Item_TPServ = 0
Else
   If IsNull(TPS!GeneroItem) Then
      Captura_Genero_Item_TPServ = 0
   Else
      Captura_Genero_Item_TPServ = TPS!GeneroItem
   End If
End If

'CommitTrans
TPS.Close
bd1.Close
Set TPS = Nothing
Set bd1 = Nothing

Exit Function

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function Captura_CodigoCST_PIS(WCodigoServico As Long) As Long
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim TPS As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set TPS = bd1.OpenRecordset("TPSERV", DB_OPEN_TABLE)
TPS.index = "PrimaryKey"
'BeginTrans
TPS.Seek "=", WCodigoServico
If TPS.NoMatch Then
   Captura_CodigoCST_PIS = 0
Else
   If IsNull(TPS!CodigoCST_PIS) Then
      Captura_CodigoCST_PIS = 0
   Else
      Captura_CodigoCST_PIS = TPS!CodigoCST_PIS
   End If
End If

'CommitTrans
TPS.Close
bd1.Close
Set TPS = Nothing
Set bd1 = Nothing

Exit Function

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Captura_Status_TPSERV(WCodigoServico As Long) As Double

On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim TPS As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set TPS = bd1.OpenRecordset("TPSERVXEMP", DB_OPEN_TABLE)
TPS.index = "PrimaryKey"
'BeginTrans
TPS.Seek "=", WCodigoServico, Empresa_Ativa()
If TPS.NoMatch Then
   Captura_Status_TPSERV = 0
Else
   If IsNull(TPS!Ativo) Then
      Captura_Status_TPSERV = 0
   Else
      Captura_Status_TPSERV = TPS!Ativo
   End If
End If

'CommitTrans
TPS.Close
bd1.Close
Set TPS = Nothing
Set bd1 = Nothing

Exit Function

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function Captura_CodigoCST_Cofis(WCodigoServico As Long) As Long
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim TPS As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set TPS = bd1.OpenRecordset("TPSERV", DB_OPEN_TABLE)
TPS.index = "PrimaryKey"
'BeginTrans
TPS.Seek "=", WCodigoServico
If TPS.NoMatch Then
   Captura_CodigoCST_Cofis = 0
Else
   If IsNull(TPS!CodigoCST_COFINS) Then
      Captura_CodigoCST_Cofis = 0
   Else
      Captura_CodigoCST_Cofis = TPS!CodigoCST_COFINS
   End If
End If

'CommitTrans
TPS.Close
bd1.Close
Set TPS = Nothing
Set bd1 = Nothing

Exit Function

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Captura_Descricao_Categoria_Servicos(WCodCategoria As Long) As String
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("CATSERV", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", WCodCategoria

If Not tb.NoMatch Then
   If IsNull(tb!Descricao) Then
      Captura_Descricao_Categoria_Servicos = "DESCRICAO NULA..."
    Else
      Captura_Descricao_Categoria_Servicos = tb!Descricao
    End If
Else
   Captura_Descricao_Categoria_Servicos = "INEXISTENTE"
End If

Finaliza:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function
Function Ajusta_CST_PIS_COFINS_TPSERV() As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("TPSERV", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

If tb.RecordCount <> 0 Then
   While Not tb.EOF
         BeginTrans
         If IsNull(tb!CodigoCST_PIS) Then
            tb.Edit
            tb!CodigoCST_PIS = 1
            tb.Update
         End If
         If IsNull(tb!CodigoCST_COFINS) Then
            tb.Edit
            tb!CodigoCST_COFINS = 1
            tb.Update
         End If
         tb.MoveNext
         CommitTrans
   Wend
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing

  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function RetornaCodigoCategoriaxCodServ(ByVal WTipoServico As Long) As Long
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("TPSERV", DB_OPEN_TABLE)

tb.index = "PrimaryKey"
tb.Seek "=", WTipoServico
If Not tb.NoMatch Then
    If Not IsNull(tb!Categoria) Then
       RetornaCodigoCategoriaxCodServ = IIf(IsNull(tb!Categoria), 0, UCase(tb!Categoria))
    Else
      RetornaCodigoCategoriaxCodServ = 0
    End If
Else
    RetornaCodigoCategoriaxCodServ = 0
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing

  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaDescricaoServico(WCodigoServico As Long) As String
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim TPS  As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set TPS = bd1.OpenRecordset("TPSERV", DB_OPEN_TABLE)
TPS.index = "PrimaryKey"

TPS.Seek "=", WCodigoServico

If Not TPS.NoMatch Then
   If IsNull(TPS!Descricao) Then
      CapturaDescricaoServico = "Descricao nula..."
   Else
      CapturaDescricaoServico = Mid$(TPS!Descricao, 1, 50)
   End If
Else
   CapturaDescricaoServico = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
End If

Finaliza:
 TPS.Close
 bd1.Close
 Set TPS = Nothing
 Set bd1 = Nothing

 Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function RetValorSemCasasDec(WValor As Double) As Double

On Error GoTo TratamentodeErro:
Dim WC As String * 1
Dim I As Long
Dim WValorConv As String
WValorConv = ""

For I = 1 To Len(WValor)
    If Mid$(WValor, I, 1) <> "" And Not IsNull(Mid$(WValor, I, 1)) Then
       WC = Mid$(WValor, I, 1)
       If WC <> "," Then
          WValorConv = WValorConv + WC
        Else
          GoSub Finaliza:
       End If
       
    End If
Next I

Finaliza:
  RetValorSemCasasDec = Val(WValorConv)
  Exit Function
Return
    
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Tipo_Pessoa(WTipo As Long) As String
On Error GoTo TratamentodeErro:


If WTipo = 1 Then
   Tipo_Pessoa = "Pessoa Fisica"
ElseIf WTipo = 2 Then
       Tipo_Pessoa = "Pessoa Juridica"
ElseIf WTipo = 3 Then
       Tipo_Pessoa = "Produtor Rural"
Else
  Tipo_Pessoa = " Estrangeiro"
End If

Finaliza:
    Exit Function
Return
    
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function


Function Calculo_DVNNSicred(ByVal strcampo As String) As String

On Error GoTo TratamentodeErro:

'############# ESTA ROTINA FOI FORNECIDA PELO BANCO

Dim intSequencia As Integer
Dim intPeso As Integer
Dim lngResultado As Long
Dim lngDigito As Long

'Atribuir os pesos (2-9) correspondentes � cada d�gito, da direita para a esquerda, efetuadno cada multiplica��o
' somando o resultado de cada multiplica��o
intPeso = 2
For intSequencia = Len(strcampo) To 1 Step -1
    lngResultado = lngResultado + (Val(Mid(strcampo, intSequencia, 1)) * intPeso)
    intPeso = IIf(intPeso >= 9, 2, intPeso + 1)
Next intSequencia

'Dividir o resultado por 11, achar o resto da divis�o, subtrai de 11 e se for maior que 9, ser� igual a 0
lngDigito = 11 - (lngResultado - ((lngResultado \ 11) * 11))

'Normalmente, se a subtra��o resultar em 10 ou 11, o d�gito ser� zero
If lngDigito > 9 Then
   lngDigito = 0
End If
   
   Calculo_DVNNSicred = lngDigito
   Exit Function
   
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function


Function Calculo_DVBradesco11(ByVal STRNUMERO As String) As String

'declara as vari�veis
Dim intContador As Integer
Dim intNumero As Integer

Dim intTotalNumero As Integer

Dim intMultiplicador As Integer

Dim intResto As Integer
Dim WResult As Long
Dim wCalcA As Long
Dim WCalcB As Long


Dim WTotalA As Long
Dim WTotalB As Long

' se nao for um valor numerico sai da fun��o
If Not IsNumeric(STRNUMERO) Then
  Calculo_DVBradesco11 = ""
  Exit Function
End If

'inicia o multiplicador
intMultiplicador = 2

'pega cada caracter do numero a partir da direita
For intContador = Len(STRNUMERO) To 1 Step -1

'extrai o caracter e multiplica prlo multiplicador
intNumero = Val(Mid(STRNUMERO, intContador, 1)) * intMultiplicador

'If Val(Mid(strnumero, intContador, 1)) = 9 Then
'   Stop
'End If


WResult = intNumero
GoSub ApuraResult
'intNumero = WCalcB

'soma o resultado para totaliza��o
intTotalNumero = intTotalNumero + intNumero

'se o multiplicador for menor  que 7 incrementa-o caso contrario atribuir valor padrao original
intMultiplicador = IIf(intMultiplicador < 7, intMultiplicador + 1, 2)

Next

'calcula o resto da divisao do total por 11
'intResto = Val((wCalcA & intTotalNumero)) Mod 11
intResto = Val((WTotalA & WTotalB)) Mod 11

'verifica as exce��es ( 0 -> DV=0    10 -> DV=X (para o BB) e retorna o DV
Select Case intResto
  Case 0
    Calculo_DVBradesco11 = "0"
  Case 1
    Calculo_DVBradesco11 = "P"
  Case 10
    Calculo_DVBradesco11 = "1"
  Case 11
    Calculo_DVBradesco11 = "1"
  Case Else
    Calculo_DVBradesco11 = 11 - Str(intResto)
End Select

Exit Function

ApuraResult:
  If WResult > 9 Then
     WTotalA = WTotalA + Val(Mid$(WResult, 1, 1))
     WTotalB = WTotalB + Val(Mid$(WResult, 2, 1))
     
     'wCalcA = wCalcA + Val(Mid$(WResult, 1, 1))
     'WCalcB = Val(Mid$(WResult, 2, 1))
  Else
     WTotalA = WTotalA + 0
     WTotalB = WTotalB + Val(Mid$(Str(WResult), 2, 1))
     'WCalcB = Mid$(str(WResult), 2, 1)
  End If
Return

End Function

Function Calculo_DV11(ByVal STRNUMERO As String) As String

'declara as vari�veis
Dim intContador As Integer
Dim intNumero As Integer

Dim intTotalNumero As Integer

Dim intMultiplicador As Integer

Dim intResto As Integer

' se nao for um valor numerico sai da fun��o
If Not IsNumeric(STRNUMERO) Then
  Calculo_DV11 = ""
  Exit Function
End If

'inicia o multiplicador
intMultiplicador = 9

'pega cada caracter do numero a partir da direita
For intContador = Len(STRNUMERO) To 1 Step -1

'extrai o caracter e multiplica prlo multiplicador
intNumero = Val(Mid(STRNUMERO, intContador, 1)) * intMultiplicador

'soma o resultado para totaliza��o
intTotalNumero = intTotalNumero + intNumero

'se o multiplicador for maior que 2 decrementa-o caso contrario atribuir valor padrao original
intMultiplicador = IIf(intMultiplicador > 2, intMultiplicador - 1, 9)

Next

'calcula o resto da divisao do total por 11
intResto = intTotalNumero Mod 11

'verifica as exce��es ( 0 -> DV=0    10 -> DV=X (para o BB) e retorna o DV
Select Case intResto
  Case 0
    Calculo_DV11 = "1"
  Case 10
    Calculo_DV11 = "1"
  Case 11
    Calculo_DV11 = "1"
  Case Else
    Calculo_DV11 = Str(intResto)
End Select


End Function

Function calculo_dv10(Snumber As String) As String
Dim tmpTotal As Integer
Dim I As Integer, f As Byte, tmpStr As String
Dim mdNums() As Integer
For I = 1 To Len(Snumber)
  f = f + 1
  If f = 2 Then
    tmpStr = CInt(Mid$(Snumber, I, 1)) * 2
    If Len(tmpStr) > 1 Then
      tmpTotal = tmpTotal + CInt(Mid$(tmpStr, 1, 1)) + CInt(Mid$(tmpStr, 2, 1))
    Else
      tmpTotal = tmpTotal + CInt(tmpStr)
    End If
    tmpStr = ""
    f = 0
  Else
    tmpTotal = tmpTotal + CInt(Mid$(Snumber, I, 1))
  End If
Next I
If Right$(CStr(tmpTotal), 1) = "0" Then
  tmpTotal = 0
Else
  tmpTotal = ((tmpTotal + 10) - CInt(Right$(CStr(tmpTotal), _
1))) - tmpTotal
End If
calculo_dv10 = tmpTotal


End Function


Function Monta_CodBarras(Banco As String, Moeda As String, Valor As Single, Vencimento As Date, Livre As String)

Dim codigo_sequencia As String
Dim DataBase As Date
Dim fator As Integer
Dim intDac As Integer

'database para calculo do fator
DataBase = CDate("7/10/1997")
fator = DateDiff("d", DataBase, Format(Vencimento, "dd/mm/yyyy"))
Valor = Int(Valor * 100)
Livre = Format(Livre, "0000000000000000000000000")

' sequencia sem o DV
codigo_sequencia = Banco & Moeda & fator & Format(Valor, "0000000000") & Livre

' calculo do DV
intDac = calcula_DV_CodBarras(codigo_sequencia)

' monta a sequencia para o codigo de barras com o DV
Monta_CodBarras = Left(codigo_sequencia, 4) & intDac & Right(codigo_sequencia, 39)

End Function

Function calcula_DV_CodBarras(Sequencia As String) As Integer

Dim intContador, intNumero, intTotalNumero As Integer
Dim intMultiplicador, intResto, intresultado As Integer
Dim caracter As String

intMultiplicador = 2

For intContador = 1 To 43
   caracter = Mid(Right(Sequencia, intContador), 1, 1)
   If intMultiplicador > 9 Then
          intMultiplicador = 2
         intNumero = 0
   End If
   If IsNumeric(caracter) Then
      intNumero = caracter * intMultiplicador
      intTotalNumero = intTotalNumero + intNumero
      intMultiplicador = intMultiplicador + 1
   End If
Next

intResto = intTotalNumero Mod 11

intresultado = 11 - intResto

If intresultado = 10 Or intresultado = 11 Or intresultado = 0 Then
    calcula_DV_CodBarras = 1
Else
    calcula_DV_CodBarras = intresultado
End If

End Function

Function Enviar_Email(WAssunto As String, WCorpo As String, WRemetente As String, wRodape As String, WEmail_Dest As String, WAnexo As String) As Long

DoCmd.Hourglass True

On Error GoTo TratamentodeErro:
   
Set oemail = New CanadaUtils.Email
        
oemail.Assunto = WAssunto
oemail.ContaEMail = Conta_Email()
oemail.EmailDestino = WEmail_Dest
oemail.Mensagem = WCorpo
oemail.SenhaEmail = Senha_Conta_Email()
oemail.SMTP = Servidor_Conta_Email()
If WAnexo = "#" Then
   oemail.Anexo = LogomarcaEmpresa()
Else
   oemail.Anexo = WAnexo
End If
If oemail.ValidaEnvioEmail = True Then
   oemail.EnviarEmail
   If oemail.StatusEnvio = True Then
      MsgBox oemail.MensagemEnvio, vbCritical, "SGa"
   Else
      DoCmd.Hourglass False
      MsgBox oemail.MensagemEnvio, vbCritical, "SGa"
      GoSub Finaliza:
   End If
End If
       

Finaliza:
   DoCmd.Hourglass False
   Set oemail = Nothing
   Exit Function
Return

TratamentodeErro:
    DoCmd.Hourglass False
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGv"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function
Private Sub testeemail()
Call Enviar_Email_RMD("ENVIO RMD EMAIL NOVO", "TESTANDO ENVIAR RMD", "GUILHEREM ANASTACIO", "guigba1@hotmail.com", "C:\sga\ImagemIndisponivel.bmp")
End Sub



Function Excluir_Selos_BD(WTipo_Selo As String) As Long
On Error GoTo TratamentodeErro
Dim db, dbc As DataBase, tb As Recordset, tbc As Recordset, qtbc As QueryDef, numreg As Long, Ret As Variant, X As Long
Dim WTotal As Double
Dim wSelo As String
Set dbc = CurrentDb()
Set db = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = db.OpenRecordset("SELOSAPLICADOS", DB_OPEN_TABLE)
Set qtbc = dbc.QueryDefs("Excluir_Selos_BD")
qtbc.Parameters(0) = [Forms]![Dialogo_Excluir_Selos_BD]![SeloInicial]
qtbc.Parameters(1) = [Forms]![Dialogo_Excluir_Selos_BD]![SeloFinal]
Set tbc = qtbc.OpenRecordset()

tb.index = "SecondKey"

BeginTrans
If tbc.RecordCount <> 0 Then
   numreg = tbc.RecordCount
   Ret = SysCmd(SYSCMD_INITMETER, "Processando, aguarde...", numreg)
   X = 1
   While Not tbc.EOF
          
         tb.Seek "=", tbc!Numero, WTipo_Selo
         If tb.NoMatch Then
            wSelo = tbc!Numero
            tbc.Delete
            Call LogSGa("DetEntradasSelos", "Dialogo_Excluir_Selos_BD", "Excluir Selos BD", Captura_Conteudo_Campo_RESP(Usuario_Ativo(), "nome"), Date, Time(), wSelo & "-" & WTipo_Selo, Environ("ComputerName"))
         Else
            MsgBox "Selo/Lacre " & Format$(tbc!Numero, "0000000") & " vinculado a OSV " & Format$(tb!Numos, "00000") & ", processamento ser� interrompido...", vbCritical, "SGa"
            GoSub Finaliza:
         End If
          
         X = X + 1
         If X Mod 500 = 0 Then
            CommitTrans
            BeginTrans
         End If
         If Not tbc.EOF Then
            tbc.MoveNext
         End If
         Ret = SysCmd(SYSCMD_UPDATEMETER, X)
   Wend
End If

Finaliza_Com_Sucesso:
    MsgBox "Fim Normal de Execu��o...", vbInformation, "SGa"
    
    GoSub Finaliza:

Return

Finaliza:
  CommitTrans
  tbc.Close
  tb.Close
  dbc.Close
  db.Close
  Set tbc = Nothing
  Set tb = Nothing
  Set dbc = Nothing
  Set db = Nothing
  Ret = SysCmd(SYSCMD_REMOVEMETER)
'  Call Rodape_SGa
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function VerificaInternet() As Long
     'By JPaulo � Maximo Access
     Dim strResultado As Long
     VerificaInternet = InternetGetConnectedState(strResultado, 0)
End Function

Function UltDiaMes(WDataRef As Variant) As Variant
UltDiaMes = DateSerial(Year(WDataRef), Month(WDataRef) + 1, 1) - 1
End Function

Function UltDiaMesUlt()
UltDiaMesUlt = DateSerial(Year(Now), Month(Now), 1) - 1
End Function

Function PesoBrutoEFD(WNroSeq As Long) As Double
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim INF, Est As Recordset, WValor, WPesoBruto As Double
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set INF = bd1.OpenRecordset("ITENSESTNOTAFISCAL", DB_OPEN_TABLE)
Set Est = bd1.OpenRecordset("EST", DB_OPEN_TABLE)
    
INF.index = "TerceiraChave"
Est.index = "PrimaryKey"
WValor = 0
BeginTrans

GoSub ApuraVolumes:

Finaliza:
  CommitTrans
   PesoBrutoEFD = WValor
   INF.Close
   Est.Close
   bd1.Close
   Set INF = Nothing
   Set Est = Nothing
   Set bd1 = Nothing
   Exit Function
Return

ApuraVolumes:
   INF.Seek "=", WNroSeq
   If Not INF.NoMatch Then
      While Not INF.EOF
         If INF!NroSeq <> WNroSeq Then
            GoSub Finaliza:
         Else
            Est.Seek "=", INF!CodigoItem
            If Not Est.NoMatch Then
               If IsNull(Est!PesoBruto) Or Est!PesoBruto = 0 Then
                  WPesoBruto = 0.0001
               Else
                  WPesoBruto = Est!PesoBruto
               End If
               WValor = WValor + (INF!Quantidade * WPesoBruto)
            End If
         End If
         INF.MoveNext
      Wend
   End If
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function QtdeVolumesEFD(WNroSeq As Long) As Double
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim INF As Recordset, WValor, WValorServ As Double
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set INF = bd1.OpenRecordset("ITENSESTNOTAFISCAL", DB_OPEN_TABLE)
    
INF.index = "TerceiraChave"
WValor = 0
BeginTrans

GoSub ApuraVolumes:

Finaliza:
  CommitTrans
   QtdeVolumesEFD = WValor
   INF.Close
   bd1.Close
   Set INF = Nothing
   Set bd1 = Nothing
   Exit Function
Return

ApuraVolumes:
   INF.Seek "=", WNroSeq
   If Not INF.NoMatch Then
      While Not INF.EOF
         If INF!NroSeq <> WNroSeq Then
            GoSub Finaliza:
         Else
            WValor = WValor + INF!Quantidade
         End If
         INF.MoveNext
      Wend
   End If
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function CodigoNumerico(wnro As String) As String
On Error GoTo TratamentodeErro:
Dim WResult As Long
Dim wresto As Long
Dim wdig As Long
Dim WCont As Long
Dim WNumero As String * 9

WNumero = Mid$(wnro, 1, 9)

For WCont = 1 To Len(WNumero) Step 1
    If WCont = 1 Then
       WResult = WResult + (Val(Mid$(WNumero, WCont, 1)) * 5)
    ElseIf WCont = 2 Then
       WResult = WResult + (Val(Mid$(WNumero, WCont, 1)) * 4)
    ElseIf WCont = 3 Then
       WResult = WResult + (Val(Mid$(WNumero, WCont, 1)) * 3)
    ElseIf WCont = 4 Then
       WResult = WResult + (Val(Mid$(WNumero, WCont, 1)) * 2)
    ElseIf WCont = 5 Then
       WResult = WResult + (Val(Mid$(WNumero, WCont, 1)) * 7)
    ElseIf WCont = 6 Then
       WResult = WResult + (Val(Mid$(WNumero, WCont, 1)) * 6)
    ElseIf WCont = 7 Then
       WResult = WResult + (Val(Mid$(WNumero, WCont, 1)) * 5)
    ElseIf WCont = 8 Then
       WResult = WResult + (Val(Mid$(WNumero, WCont, 1)) * 4)
    ElseIf WCont = 9 Then
       WResult = WResult + (Val(Mid$(WNumero, WCont, 1)) * 3)
    'Else
    '   WResult = WResult + (Val(Mid$(WNumero, WCont, 1)) * 2)
    End If
    
  '   WResto = WResult Mod 11
  '   WDig = 11 - WResto
     
'     Stop
Next WCont
wresto = WResult Mod 9
wdig = 9 - wresto
CodigoNumerico = Format$(WNumero, "000000000") + Format$(wdig, "0")
'MsgBox "N/N" & GeraNossoNumero
Exit Function

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ValorContabilEFD(WCodFornec As Long, WNotaFiscal As Long, WDataEntrada As Variant) As Double
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim X As Long, tb As Recordset
Dim WTotal As Double
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd1.OpenRecordset("DETENTRADAS", DB_OPEN_TABLE)
tb.index = "TerceiraChave"

'BeginTrans
tb.Seek "=", WDataEntrada, WNotaFiscal, WCodFornec
If tb.NoMatch Then
   GoSub Finaliza:
Else
   While Not tb.EOF
         If tb!CodigoFornecedor = WCodFornec And tb!NotaFiscal = WNotaFiscal And tb!DataEntrada = WDataEntrada Then
             WTotal = WTotal + ((tb!ValorUnit - tb!Desconto + tb!OUTRASDESPESAS + tb!VlrFrete + tb!ICMSRETIDO + Round(tb!ValorUnit * tb!IPI / 100)) * tb!Quantidade)
         Else
             GoSub Finaliza:
         End If
         tb.MoveNext
      Wend
End If

Finaliza:
 ' CommitTrans
  ValorContabilEFD = WTotal
  tb.Close
  bd1.Close
  Set tb = Nothing
  Set bd1 = Nothing
   
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ValorFreteEFD(WCodFornec As Long, WNotaFiscal As Long, WDataEntrada As Variant) As Double
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim X As Long, tb As Recordset
Dim WTotal As Double
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd1.OpenRecordset("DETENTRADAS", DB_OPEN_TABLE)
tb.index = "TerceiraChave"

'BeginTrans
tb.Seek "=", WDataEntrada, WNotaFiscal, WCodFornec
If tb.NoMatch Then
   GoSub Finaliza:
Else
   While Not tb.EOF
         If tb!CodigoFornecedor = WCodFornec And tb!NotaFiscal = WNotaFiscal And tb!DataEntrada = WDataEntrada Then
             WTotal = WTotal + (tb!VlrFrete * tb!Quantidade)
         Else
             GoSub Finaliza:
         End If
         tb.MoveNext
      Wend
End If

Finaliza:
 ' CommitTrans
  ValorFreteEFD = WTotal
  tb.Close
  bd1.Close
  Set tb = Nothing
Set bd1 = Nothing

  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ValorDescontoEFD(WCodFornec As Long, WNotaFiscal As Long, WDataEntrada As Variant) As Double
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim X As Long, tb As Recordset
Dim WTotal As Double
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd1.OpenRecordset("DETENTRADAS", DB_OPEN_TABLE)
tb.index = "TerceiraChave"

'BeginTrans
tb.Seek "=", WDataEntrada, WNotaFiscal, WCodFornec
If tb.NoMatch Then
   GoSub Finaliza:
Else
   While Not tb.EOF
         If tb!CodigoFornecedor = WCodFornec And tb!NotaFiscal = WNotaFiscal And tb!DataEntrada = WDataEntrada Then
             WTotal = WTotal + (tb!Desconto * tb!Quantidade)
         Else
             GoSub Finaliza:
         End If
         tb.MoveNext
      Wend
End If

Finaliza:
 ' CommitTrans
  ValorDescontoEFD = WTotal
  tb.Close
  bd1.Close
  Set tb = Nothing
Set bd1 = Nothing

  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ValorOUTRASDESPESASEFD(WCodFornec As Long, WNotaFiscal As Long, WDataEntrada As Long) As Double

On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim QRs As QueryDef
Dim rs  As Recordset
Dim WTotal As Double
Set bd = CurrentDb()
Set QRs = bd.QueryDefs("Q_DETENTRADAS_TERCEIRA_CHAVE")
QRs.Parameters(0) = WDataEntrada
QRs.Parameters(1) = WNotaFiscal
QRs.Parameters(2) = WCodFornec
Set rs = QRs.OpenRecordset()
WTotal = 0

If rs.RecordCount <> 0 Then
   While Not rs.EOF
         WTotal = WTotal + (rs!OUTRASDESPESAS * rs!Quantidade)
         rs.MoveNext
      Wend
End If

Finaliza:
  ValorOUTRASDESPESASEFD = WTotal
  rs.Close
  bd.Close
  Set rs = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


'On Error GoTo TratamentodeErro:
''Dim bd1 As DataBase
'Dim X As Long, tb As Recordset
'Dim wtotal As Double
'Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
'Set tb = bd1.OpenRecordset("DETENTRADAS", DB_OPEN_TABLE)
'tb.index = "TerceiraChave"

'BeginTrans
'tb.Seek "=", WDataEntrada, WNotaFiscal, WCodFornec
'If tb.NoMatch Then
'   GoSub Finaliza:
'Else
'   While Not tb.EOF
'         If tb!CodigoFornecedor = WCodFornec And tb!NotaFiscal = WNotaFiscal And tb!DataEntrada = WDataEntrada Then
'             wtotal = wtotal + (tb!OUTRASDESPESAS * tb!Quantidade)
'         Else
'             GoSub Finaliza:
'         End If
'         tb.MoveNext
'      Wend
'End If

'Finaliza:
' ' CommitTrans
'  ValorOUTRASDESPESASEFD = wtotal
'  tb.Close
'  bd1.Close
'  Set tb = Nothing
'  Set bd1 = Nothing

'  Exit Function
'Return

'TratamentodeErro:
 '   If Err.Number = 3001 Then
'       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
'    Else
'       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description, Vbcritical,"SGa"
'       Resume Next
'    End If

End Function

Function ValorICMSDESONEFD(WCodFornec As Long, WNotaFiscal As Long, WDataEntrada As Long) As Double
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim X As Long, tb As Recordset
Dim WTotal As Double
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd1.OpenRecordset("DETENTRADAS", DB_OPEN_TABLE)
tb.index = "TerceiraChave"

'BeginTrans
tb.Seek "=", WDataEntrada, WNotaFiscal, WCodFornec
If tb.NoMatch Then
   GoSub Finaliza:
Else
   While Not tb.EOF
         If tb!CodigoFornecedor = WCodFornec And tb!NotaFiscal = WNotaFiscal And tb!DataEntrada = WDataEntrada Then
             WTotal = WTotal + (tb!ICMs_Deson * tb!Quantidade)
         Else
             GoSub Finaliza:
         End If
         tb.MoveNext
      Wend
End If

Finaliza:
 ' CommitTrans
  ValorICMSDESONEFD = WTotal
  tb.Close
  bd1.Close
  Set tb = Nothing
  Set bd1 = Nothing

  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function


Function ValorICMSSTEFD(WCodFornec As Long, WNotaFiscal As Long, WDataEntrada As Long) As Double
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim X As Long, tb As Recordset
Dim WTotal As Double
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd1.OpenRecordset("DETENTRADAS", DB_OPEN_TABLE)
tb.index = "TerceiraChave"

'BeginTrans
tb.Seek "=", WDataEntrada, WNotaFiscal, WCodFornec
If tb.NoMatch Then
   GoSub Finaliza:
Else
   While Not tb.EOF
         If tb!CodigoFornecedor = WCodFornec And tb!NotaFiscal = WNotaFiscal And tb!DataEntrada = WDataEntrada Then
             WTotal = WTotal + (tb!ICMSRETIDO * tb!Quantidade)
         Else
             GoSub Finaliza:
         End If
         tb.MoveNext
      Wend
End If

Finaliza:
 ' CommitTrans
  ValorICMSSTEFD = WTotal
  tb.Close
  bd1.Close
  Set tb = Nothing
  Set bd1 = Nothing

  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function



Function ValorProdutosEFD(WCodFornec As Long, WNotaFiscal As Long, WDataEntrada As Variant) As Double
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim X As Long, tb, tbe As Recordset
Dim WTotal As Double
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tbe = bd1.OpenRecordset("ENTRADAS", DB_OPEN_TABLE)
Set tb = bd1.OpenRecordset("DETENTRADAS", DB_OPEN_TABLE)
tb.index = "TErceiraChave"
tbe.index = "ChavePrimaria"

'BeginTrans
tb.Seek "=", WDataEntrada, WNotaFiscal, WCodFornec
If tb.NoMatch Then
   tbe.Seek "=", WDataEntrada, WNotaFiscal, WCodFornec
   If tbe.NoMatch Then
      WTotal = 0
      GoSub Finaliza:
   Else
      WTotal = tbe!ValorTotal
      GoSub Finaliza
   End If
Else
   While Not tb.EOF
         If tb!CodigoFornecedor = WCodFornec And tb!NotaFiscal = WNotaFiscal And tb!DataEntrada = WDataEntrada Then
            WTotal = WTotal + Round(tb!ValorUnit * tb!Quantidade)
         Else
             GoSub Finaliza:
         End If
         tb.MoveNext
      Wend
End If

Finaliza:
 ' CommitTrans
  ValorProdutosEFD = WTotal
  tb.Close
  tbe.Close
  bd1.Close
  Set tb = Nothing
  Set tbe = Nothing
  Set bd1 = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ApuraUltimaNFDia(WDataPedido As Variant) As Long

On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim tb As Recordset
Dim WNroNotaFiscal As Long
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd1.OpenRecordset("NotaFiscal", DB_OPEN_TABLE)
tb.index = "DataEmissaoXNroNotaFiscalXEmpresa"

'BeginTrans
WNroNotaFiscal = 1

GoSub ApuraPrimeira:

While Not tb.EOF
      If tb!DataPedido = WDataPedido Then
          If tb!CodEmpresa = Form_Menu.WEmpresa Then
             WNroNotaFiscal = tb!NroNotaFiscal
          End If
      Else
         GoSub Finaliza:
      End If
      tb.MoveNext
Wend

Finaliza:
   ApuraUltimaNFDia = WNroNotaFiscal
 '  CommitTrans
   tb.Close
   bd1.Close
   Set tb = Nothing
   Set bd1 = Nothing

   Exit Function
Return

ApuraPrimeira:
tb.Seek "=", WDataPedido, WNroNotaFiscal, Form_Menu.WEmpresa
If tb.NoMatch Then
   While tb.NoMatch
         WNroNotaFiscal = WNroNotaFiscal + 1
         If WNroNotaFiscal > ApuraUltimaNFEmitida() Then
             WNroNotaFiscal = 0
             GoSub Finaliza:
         End If
         tb.Seek "=", WDataPedido, WNroNotaFiscal
   Wend
End If

Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ApuraPrimeiraNFDia(WDataPedido As Variant) As Long

On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim tb As Recordset
Dim WNroNotaFiscal As Long
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd1.OpenRecordset("NotaFiscal", DB_OPEN_TABLE)
tb.index = "DataEmissaoXNroNotaFiscalXemnpresa"

'BeginTrans
WNroNotaFiscal = 1
tb.Seek "=", WDataPedido, WNroNotaFiscal, Form_Menu.WEmpresa
If tb.NoMatch Then
   While tb.NoMatch
         WNroNotaFiscal = WNroNotaFiscal + 1
         If WNroNotaFiscal > ApuraUltimaNFEmitida() Then
             WNroNotaFiscal = 0
             GoSub Finaliza:
         End If
         tb.Seek "=", WDataPedido, WNroNotaFiscal
   Wend
Else
   ApuraPrimeiraNFDia = WNroNotaFiscal
End If

Finaliza:
   ApuraPrimeiraNFDia = WNroNotaFiscal
'   CommitTrans
   tb.Close
   bd1.Close
   Set tb = Nothing
   Set bd1 = Nothing

   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ApuraTotalNotaFiscalServ(WNroSeq As Long) As Double
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim NOTA  As Recordset, X As Long, Est As Recordset, serv As Recordset
Dim WVALORITEMEST As Double
Dim WTotalLiqServ As Double
Dim WTotalBrutoServ As Double
Dim WTotalDescServ As Double
Dim WTotalliqEst As Double
Dim WTotalBrutoEst As Double
Dim WTotalDescEst As Double
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set NOTA = bd1.OpenRecordset("NOTAFISCAL", DB_OPEN_TABLE)
Set Est = bd1.OpenRecordset("ITENSESTNOTAFISCAL", DB_OPEN_TABLE)
Set serv = bd1.OpenRecordset("ITENSSERVNOTAFISCAL", DB_OPEN_TABLE)
Est.index = "TerceiraChave"
serv.index = "TerceiraChave"
NOTA.index = "PrimaryKey"


NOTA.Seek "=", WNroSeq
If NOTA.NoMatch Then
   ApuraTotalNotaFiscalServ = 0
   GoSub Finaliza:
Else
   GoSub ApuraServicos:
'   GoSub ApuraItensEstoque:
   ApuraTotalNotaFiscalServ = WTotalliqEst + WTotalLiqServ
End If

Finaliza:
  Est.Close
  NOTA.Close
  serv.Close
  bd1.Close
  Set Est = Nothing
  Set NOTA = Nothing
  Set serv = Nothing
  Set bd1 = Nothing
  Exit Function
Return

ApuraItensEstoque:
   Est.Seek "=", WNroSeq
   If Est.NoMatch Then
      Return
   Else
      Do Until Est.EOF
         If Est!NroSeq <> WNroSeq Then
            Exit Do
         Else
            If NOTA!DECRETO43349 = "1" Then
               WVALORITEMEST = (Est!ValorUnitario - Est!Desconto) - Round((Est!ValorUnitario - Est!Desconto) * NOTA!AliquotaICMS)
               WTotalliqEst = WTotalliqEst + (WVALORITEMEST * Est!Quantidade)
            Else
               WTotalliqEst = WTotalliqEst + ((Est!ValorUnitario - Est!Desconto) * Est!Quantidade) + Est!ICMSST
            End If
            Est.MoveNext
         End If
      Loop
   End If
Return

ApuraServicos:
   serv.Seek "=", WNroSeq
   If serv.NoMatch Then
      Return
   Else
      Do Until serv.EOF
         If serv!NroSeq <> WNroSeq Then
            Exit Do
         Else
            WTotalLiqServ = WTotalLiqServ + ((serv!ValorMaoObra - serv!Desconto) * serv!Quantidade)
            WTotalBrutoServ = WTotalBrutoServ + (serv!ValorMaoObra * serv!Quantidade)
            WTotalDescServ = WTotalDescServ + (serv!Desconto * serv!Quantidade)
            serv.MoveNext
         End If
      Loop
   End If
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function ApuraTotalNotaFiscalPecas(WNroSeq As Long) As Double
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim NOTA  As Recordset, X As Long, Est As Recordset, serv As Recordset
Dim WVALORITEMEST As Double
Dim WTotalLiqServ As Double
Dim WTotalBrutoServ As Double
Dim WTotalDescServ As Double
Dim WTotalliqEst As Double
Dim WTotalBrutoEst As Double
Dim WTotalDescEst As Double
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set NOTA = bd1.OpenRecordset("NOTAFISCAL", DB_OPEN_TABLE)
Set Est = bd1.OpenRecordset("ITENSESTNOTAFISCAL", DB_OPEN_TABLE)
Set serv = bd1.OpenRecordset("ITENSSERVNOTAFISCAL", DB_OPEN_TABLE)
Est.index = "TerceiraChave"
serv.index = "TerceiraChave"
NOTA.index = "PrimaryKey"


NOTA.Seek "=", WNroSeq
If NOTA.NoMatch Then
   ApuraTotalNotaFiscalPecas = 0
   GoSub Finaliza:
Else
   'GoSub ApuraServicos:
   GoSub ApuraItensEstoque:
   ApuraTotalNotaFiscalPecas = WTotalliqEst + WTotalLiqServ
End If

Finaliza:
  Est.Close
  NOTA.Close
  serv.Close
  bd1.Close
  Set Est = Nothing
  Set NOTA = Nothing
  Set serv = Nothing
  Set bd1 = Nothing
  Exit Function
Return

ApuraItensEstoque:
   Est.Seek "=", WNroSeq
   If Est.NoMatch Then
      Return
   Else
      Do Until Est.EOF
         If Est!NroSeq <> WNroSeq Then
            Exit Do
         Else
            If NOTA!DECRETO43349 = "1" Then
               WVALORITEMEST = (Est!ValorUnitario - Est!Desconto) - Round((Est!ValorUnitario - Est!Desconto) * NOTA!AliquotaICMS)
               WTotalliqEst = WTotalliqEst + (WVALORITEMEST * Est!Quantidade)
            Else
               WTotalliqEst = WTotalliqEst + ((Est!ValorUnitario - Est!Desconto) * Est!Quantidade)
            End If
            Est.MoveNext
         End If
      Loop
   End If
Return

ApuraServicos:
   serv.Seek "=", WNroSeq
   If serv.NoMatch Then
      Return
   Else
      Do Until serv.EOF
         If serv!NroSeq <> WNroSeq Then
            Exit Do
         Else
            WTotalLiqServ = WTotalLiqServ + ((serv!ValorMaoObra - serv!Desconto) * serv!Quantidade)
            WTotalBrutoServ = WTotalBrutoServ + (serv!ValorMaoObra * serv!Quantidade)
            WTotalDescServ = WTotalDescServ + (serv!Desconto * serv!Quantidade)
            serv.MoveNext
         End If
      Loop
   End If
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function
'Fun��o para checar refer�ncia se n�o existe cria
'Call RefArquivo("DAO350.dll") office 97
'Call RefArquivo("Dao360.dll") office 2K
Function RefArquivo(strNomeBiblioteca As String)
On Error GoTo Error_RefArquivo
Dim Ref As Reference, strCaminho As String
'Cortesia: Carlos Moura - crpmoura@ig.com.br

'Verifica se j� existe refer�ncia
For Each Ref In References
If Ref.Name = strNomeBiblioteca Then
Exit Function
End If
Next

strCaminho = CapturaCaminhoBD & "\"
'Cria refer�ncia a biblioteca
Set Ref = References.AddFromFile(strCaminho & strNomeBiblioteca)

Exit_RefArquivo:
Exit Function
Error_RefArquivo:
If Err = 32813 Then
   Exit Function
End If

MsgBox Err & ": " & Err.Description
Resume Exit_RefArquivo
End Function

Function RemoverRefPerdida()
'Remove referencias perdidas
'Autor: Carlos Moura - www.geocities.com/espaco_access
Dim Ref As Reference
On Error Resume Next
For Each Ref In References
If Ref.IsBroken = True Then
References.Remove Ref
End If
Next
End Function
Function MascaraRedutor(WRedutor As String) As String
  MascaraRedutor = Format$(WRedutor, "@.@@@@:@")
End Function

Function Retorna_Proprietario_Descontam_OSV(WNumos As Long, WNum_Tanque As String) As String
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("DESCONTAM_TANQUE_OSV", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", WNumos, WNum_Tanque
If tb.NoMatch Then
   Retorna_Proprietario_Descontam_OSV = "INDEFINIDO"
Else
   If IsNull(tb!Marca_Tanque) Then
      Retorna_Proprietario_Descontam_OSV = "NULL"
   Else
      Retorna_Proprietario_Descontam_OSV = tb!Proprietario
   End If
End If

Finaliza:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing

   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function Retorna_ModeloTanque_Descontam_OSV(WNumos As Long, WNum_Tanque As String) As String
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("DESCONTAM_TANQUE_OSV", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", WNumos, WNum_Tanque
If tb.NoMatch Then
   Retorna_ModeloTanque_Descontam_OSV = "INDEFINIDO"
Else
   If IsNull(tb!Marca_Tanque) Then
      Retorna_ModeloTanque_Descontam_OSV = "NULL"
   Else
      Retorna_ModeloTanque_Descontam_OSV = tb!Modelo_Tanque
   End If
End If

Finaliza:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing

   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Retorna_MarcaTanque_Descontam_OSV(WNumos As Long, WNum_Tanque As String) As String
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("DESCONTAM_TANQUE_OSV", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", WNumos, WNum_Tanque
If tb.NoMatch Then
   Retorna_MarcaTanque_Descontam_OSV = "INDEFINIDO"
Else
   If IsNull(tb!Marca_Tanque) Then
      Retorna_MarcaTanque_Descontam_OSV = "NULL"
   Else
      Retorna_MarcaTanque_Descontam_OSV = tb!Marca_Tanque
   End If
End If

Finaliza:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing

   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Retorna_NumTanque_Descontam_OSV(WNumos As Long) As String
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("DESCONTAM_TANQUE_OSV", DB_OPEN_TABLE)
tb.index = "NumOS"
tb.Seek "=", WNumos
If tb.NoMatch Then
   Retorna_NumTanque_Descontam_OSV = "INDEFINIDO"
Else
   While tb!Numos = WNumos
         If IsNull(tb!Num_Tanque) Then
            Retorna_NumTanque_Descontam_OSV = "NULL"
         Else
             Retorna_NumTanque_Descontam_OSV = tb!Num_Tanque
         End If
         
         tb.MoveNext
         If tb.EOF Then
            GoSub Finaliza:
         End If
   
   Wend
End If

Finaliza:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing

   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function




Function Ret_Ini_Status_OSV(WNumos As Long) As Variant
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim Qtb As QueryDef
Dim tb  As Recordset
Dim Wini As Variant
Set bd = CurrentDb()
Set Qtb = bd.QueryDefs("Ret_Status_OSV")
Qtb.Parameters(0) = WNumos
Set tb = Qtb.OpenRecordset()


If tb.RecordCount <> 0 Then
  While Not tb.EOF
        Wini = tb!Inicio
        tb.MoveNext
   Wend
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
   Set bd = Nothing

  Ret_Ini_Status_OSV = Wini
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function


Function Status_Atual_OSV(WNumos As Long) As String
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim Qtb As QueryDef
Dim tb  As Recordset
Dim WST As String
Set bd = CurrentDb()
Set Qtb = bd.QueryDefs("Ret_Status_OSV")
Qtb.Parameters(0) = WNumos
Set tb = Qtb.OpenRecordset()

WST = "AN"

If tb.RecordCount <> 0 Then
  While Not tb.EOF
        WST = tb!codigo_status
        tb.MoveNext
   Wend
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
   Set bd = Nothing

  Status_Atual_OSV = Ret_Descr_Status_OSV(WST)
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Exit Function
    End If

End Function

Function Ret_Ult_Status_OSV(WNumos As Long) As String
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim Qtb As QueryDef
Dim tb  As Recordset
Dim WST As String
Set bd = CurrentDb()
Set Qtb = bd.QueryDefs("Ret_Ult_Status_OSV")
Qtb.Parameters(0) = WNumos
Set tb = Qtb.OpenRecordset()

WST = "AN"

If tb.RecordCount <> 0 Then
  While Not tb.EOF
        WST = tb!Descricao
        GoSub Finaliza:
        tb.MoveNext
   Wend
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
   Set bd = Nothing

  Ret_Ult_Status_OSV = WST
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Exit Function
    End If

End Function




Function CapturaCodItemXCodigoFabricante(WCodfabricante As String) As Long
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim Est  As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Est = bd1.OpenRecordset("EST", DB_OPEN_TABLE)
Est.index = "CodigoFabricante"

Est.Seek "=", WCodfabricante

If Not Est.NoMatch Then
   If IsNull(Est!CodItem) Then
      CapturaCodItemXCodigoFabricante = 1
   Else
      CapturaCodItemXCodigoFabricante = Est!CodItem
   End If
Else
   CapturaCodItemXCodigoFabricante = 0
End If

Finaliza:
 Est.Close
 bd1.Close
 Set Est = Nothing
 Set bd1 = Nothing
 Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function IncluiModeloVeic(WModeloVeic As String) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase, tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("MODELOVEIC", DB_OPEN_TABLE)

tb.index = "Primarykey"

BeginTrans
tb.Seek "=", WModeloVeic
If tb.NoMatch Then
   tb.AddNew
Else
   tb.Edit
End If
tb!DescricaoModelo = WModeloVeic
tb.Update

Finaliza:
  CommitTrans
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function


Function ImportacaoPlacasXClientes(WCodigoCliente As Long, Wcaminho As String) As Long
DoCmd.Hourglass True
On Error GoTo TratamentodeErro:

Dim LFCR As String * 2, Msg As String

Dim numreg As Long
Dim WPar As Long
Dim Ret As Variant
Dim X As Long
Dim WNome As String
Dim WSeq As String * 2

Dim MeuRegistro As String * 18

'Descricao do Registro
Dim WRegPlaca As String * 7
Dim WRegRenavam As String * 11

Dim db As DataBase
Dim tb As Recordset

Set db = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = db.OpenRecordset("PlacXCli", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
LFCR = Chr(10) & Chr(13)
X = 0

WNome = Wcaminho

Open Wcaminho For Input As #1
BeginTrans
'Ret = SysCmd(SYSCMD_INITMETER, "Aguarde, gerando registro 00...", NumReg)
While Not EOF(1)

      Line Input #1, MeuRegistro ' L� a linha para a vari�vel.
      WRegPlaca = Mid$(MeuRegistro, 1, 7)
      WRegRenavam = Mid$(MeuRegistro, 9, 11)

      Ret = SysCmd(SYSCMD_SETSTATUS, "Importando placa " & WRegPlaca)
     
      GoSub IncluiRegistro:
   
      X = X + 1
      If X Mod 200 = 0 Then
         CommitTrans
         BeginTrans
      End If
Wend

Finaliza:
  GoSub ResumoProcessamento:
  CommitTrans
  Close #1
  tb.Close
  db.Close
  Set tb = Nothing
   Set db = Nothing

  DoCmd.Hourglass False
  Exit Function
 ' Call Rodape_SGa
Return

ResumoProcessamento:
  Ret = SysCmd(acSysCmdRemoveMeter)
  Msg = "Importa��o de dados !" & LFCR & LFCR
  Msg = Msg & LFCR & "Empresa .: " & UCase(CapturaRazaoSocialCli(WCodigoCliente))
  Msg = Msg & LFCR & Format$(X, "0,000") & " registros importados com sucesso...  "
  Msg = Msg & LFCR & "Status arquivo : " & " OK"
  MsgBox Msg
 ' Call Rodape_SGa
Return

IncluiRegistro:
  tb.Seek "=", WRegPlaca, WCodigoCliente
  If tb.NoMatch Then
     tb.AddNew
  Else
     tb.Edit
  End If
  tb!Placa = WRegPlaca
  tb!CodigoCliente = WCodigoCliente
  tb!Data_Validade_Resultado = #1/1/2001#
  tb!Data_Ultima_Consulta = #1/1/2001#
  tb!TipoCertificadoResultado = "CERTIFICADO"
  tb!DataCertificado2 = #1/1/2001#
  tb!ValidadeCertificado2 = #1/1/2001#
  tb!DataResultado = #1/1/2001#
  tb!DataValidadeResultado = #1/1/2001#
  tb!Tipo_Veic = "C"
  tb!Status = 1
  If IsNumeric(WRegRenavam) Then
     tb!Renavam = WRegRenavam
  End If
  tb.Update
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "AVISO"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       CommitTrans
       Close #1
       tb.Close
       db.Close
       Set tb = Nothing
       Set db = Nothing
       DoCmd.Hourglass False
       Exit Function
'       Resume
    End If

End Function


Function UltAquisicaoItemFornec(WCodItem As Long) As Variant
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim Qtb As QueryDef
Dim tb  As Recordset
Dim WRazaoSocial As String
Set bd = CurrentDb()
Set Qtb = bd.QueryDefs("UltAquisicaoItemEst")
Qtb.Parameters(0) = Forms!cadastroitensestoque!CodItem
'Qtb.Parameters(0) = WCodItem
Set tb = Qtb.OpenRecordset()

If tb.RecordCount <> 0 Then
  While Not tb.EOF
    WRazaoSocial = IIf(IsNull(tb!RazaoSocial), "###", tb![RazaoSocial])
    GoSub Finaliza:
   Wend
End If


Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
   Set bd = Nothing

  UltAquisicaoItemFornec = WRazaoSocial
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function


Function UltAquisicaoItemPrazo(WCodItem As Long) As Variant
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim Qtb As QueryDef
Dim tb  As Recordset
Dim WPrazo As String
Set bd = CurrentDb()
Set Qtb = bd.QueryDefs("UltAquisicaoItemEst")
Qtb.Parameters(0) = Forms!cadastroitensestoque!CodItem
'Qtb.Parameters(0) = WCodItem
Set tb = Qtb.OpenRecordset()

If tb.RecordCount <> 0 Then
  While Not tb.EOF
    WPrazo = IIf(IsNull(tb![Descricao]), "###", tb![Descricao])
    GoSub Finaliza:
   Wend
End If


Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  UltAquisicaoItemPrazo = WPrazo
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function UltAquisicaoItemDataEntrada(WCodItem As Long) As Variant
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim Qtb As QueryDef
Dim tb  As Recordset
Dim WDataEntrada As Variant
Set bd = CurrentDb()
Set Qtb = bd.QueryDefs("UltAquisicaoItemEst")
Qtb.Parameters(0) = Forms!cadastroitensestoque!CodItem
Set tb = Qtb.OpenRecordset()

If tb.RecordCount <> 0 Then
  While Not tb.EOF
    WDataEntrada = IIf(IsNull(tb![DataEntrada]), #1/1/2001#, tb![DataEntrada])
    GoSub Finaliza:
   Wend
End If


Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  UltAquisicaoItemDataEntrada = WDataEntrada
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function


Function UltAquisicaoItem(WCodItem As Long) As Double
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim Qtb As QueryDef
Dim tb  As Recordset
Dim WPrecoCompra As Double
Dim WValor As Double
Dim WFrete As Double
Dim WIcmsretido As Double
Dim WIPI As Double
Dim WOutrasdespesas As Double
Dim WDesconto As Double
Set bd = CurrentDb()
Set Qtb = bd.QueryDefs("UltAquisicaoItemEst")
Qtb.Parameters(0) = Forms!cadastroitensestoque!CodItem
Set tb = Qtb.OpenRecordset()

WValor = 0

If tb.RecordCount <> 0 Then
  While Not tb.EOF
    GoSub ApuraPreco:
    WPrecoCompra = (WValor + WFrete + WIcmsretido + WIPI + WOutrasdespesas - WDesconto)
    GoSub Finaliza:
   Wend
End If


Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  UltAquisicaoItem = WPrecoCompra
  Exit Function
Return

ApuraPreco:
   WValor = IIf(IsNull(tb!PrecoCompra), 0, tb!PrecoCompra)
   'WFrete = IIf(IsNull(tb!�ltimodeVlrFrete), 0, tb!�ltimodeVlrFrete)
   'WIcmsretido = IIf(IsNull(tb!�ltimodeICMSRETIDO), 0, tb!�ltimodeICMSRETIDO)
   'WIPI = IIf(IsNull(tb!�ltimodeIPI), 0, (tb!�ltimoDeValorUnit * tb!�ltimodeIPI) / 100)
   'WOutrasdespesas = IIf(IsNull(tb!�ltimodeOUTRASDESPESAS), 0, tb!�ltimodeOUTRASDESPESAS)
   'WDesconto = IIf(IsNull(tb!�ltimodeDesconto), 0, tb!�ltimodeDesconto)

Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function
Function Ultima_Compra_Est(WCodItem As Long) As Double
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim Qtb As QueryDef
Dim tb  As Recordset
Dim WPrecoCompra As Double
Dim WValor As Double
Dim WFrete As Double
Dim WIcmsretido As Double
Dim WIPI As Double
Dim WOutrasdespesas As Double
Dim WDesconto As Double
Set bd = CurrentDb()
Set Qtb = bd.QueryDefs("UltAquisicaoItemEst")
Qtb.Parameters(0) = WCodItem
Set tb = Qtb.OpenRecordset()

WValor = 0

If tb.RecordCount <> 0 Then
  While Not tb.EOF
    GoSub ApuraPreco:
    WPrecoCompra = (WValor + WFrete + WIcmsretido + WIPI + WOutrasdespesas - WDesconto)
    GoSub Finaliza:
   Wend
End If


Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Ultima_Compra_Est = WPrecoCompra
  Exit Function
Return

ApuraPreco:
   WValor = IIf(IsNull(tb!PrecoCompra), 0, tb!PrecoCompra)
   'WFrete = IIf(IsNull(tb!�ltimodeVlrFrete), 0, tb!�ltimodeVlrFrete)
   'WIcmsretido = IIf(IsNull(tb!�ltimodeICMSRETIDO), 0, tb!�ltimodeICMSRETIDO)
   'WIPI = IIf(IsNull(tb!�ltimodeIPI), 0, (tb!�ltimoDeValorUnit * tb!�ltimodeIPI) / 100)
   'WOutrasdespesas = IIf(IsNull(tb!�ltimodeOUTRASDESPESAS), 0, tb!�ltimodeOUTRASDESPESAS)
   'WDesconto = IIf(IsNull(tb!�ltimodeDesconto), 0, tb!�ltimodeDesconto)

Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function CapturaNomeContatoCli(WCodigoCliente As Long) As String
On Error GoTo TratamentodeErro:
Dim CNT As Recordset, bd1 As DataBase
Dim WCont As String
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set CNT = bd1.OpenRecordset("CNT", DB_OPEN_TABLE)
CNT.index = "SecondKey"
WCont = Space(1)
CNT.Seek "=", WCodigoCliente
If CNT.NoMatch Then
   CapturaNomeContatoCli = Space(10)
Else
  While Not CNT.EOF
      If WCodigoCliente <> CNT!CodigoCliente Then
         GoSub Finaliza:
      End If
      If IsNull(CNT!Codinome) Then
         WCont = WCont & Space(1)
      Else
          WCont = WCont & CNT!Codinome & " / "
      End If
      
      CNT.MoveNext
   Wend
End If

Finaliza:
  CapturaNomeContatoCli = WCont
  CNT.Close
  bd1.Close
  Set CNT = Nothing
  Set bd1 = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "AVISO"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Function_Exclui_Finalidade_Descontam_OSV(WNumos As Long, WNum_Tanque As String, WFinalidade As Long) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("DESCONTAM_TANQUE_OSV_FINALID", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

BeginTrans

tb.Seek "=", WNumos, WNum_Tanque, WFinalidade
If Not tb.NoMatch Then
   tb.Delete
   MsgBox "Exclus�o efetuada com sucesso...", vbCritical, "SGa"
End If

Finaliza:
    CommitTrans
    tb.Close
    bd.Close
    Set tb = Nothing
    Set bd = Nothing
    Exit Function
Return



TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Excluir_Compart_Descontam_OSV(WNumos As Long, WNum_Tanque As String, WCompart As Long) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("DESCONTAM_TANQUE_OSV_COMPART", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

BeginTrans

tb.Seek "=", WNumos, WNum_Tanque, WCompart
If Not tb.NoMatch Then
   tb.Delete
   MsgBox "Exclus�o efetuada com sucesso...", vbCritical, "SGa"
End If

Finaliza:
    CommitTrans
    tb.Close
    bd.Close
    Set tb = Nothing
    Set bd = Nothing
    Exit Function
Return



TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function



Function Capacidade_Lanc_Tanque_Descontam_OSV(WNumos As Long) As Double

On Error GoTo TratamentodeErro:
Dim WCapac As Double
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("DESCONTAM_TANQUE_OSV_COMPART", DB_OPEN_TABLE)
tb.index = "NumOS"

WCapac = 0

tb.Seek "=", WNumos
If tb.NoMatch Then
   WCapac = 0
Else
   While tb!Numos = WNumos
         If IsNull(tb!Volume_C) Then
            WCapac = 0
         Else
            WCapac = WCapac + Val(tb!Volume_C)
         End If
         tb.MoveNext
         If tb.EOF Then
            GoSub Finaliza:
         End If
   Wend
End If


Finaliza:
    Capacidade_Lanc_Tanque_Descontam_OSV = WCapac
    tb.Close
    bd.Close
    Set tb = Nothing
  Set bd = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function


Function Valida_Exclusao_Nro_Seq_Tanque_Descontam_OSV(WNumos As Long, WNumTanque As String, WCompart As Long) As Boolean
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("DESCONTAM_TANQUE_OSV_COMPART", DB_OPEN_TABLE)
tb.index = "Numos"

Valida_Exclusao_Nro_Seq_Tanque_Descontam_OSV = True

tb.Seek "=", WNumos
If tb.NoMatch Then
   MsgBox "Compartimento inexistente...", vbCritical, "SGa"
   Valida_Exclusao_Nro_Seq_Tanque_Descontam_OSV = False
   tb.Close
   bd.Close
   Exit Function
Else
   While tb!Numos = WNumos
         If tb!num_seq > WCompart Then
            Valida_Exclusao_Nro_Seq_Tanque_Descontam_OSV = False
            MsgBox "Compartimento nao disponivel para exclusao...", vbCritical, "SGa"
            GoSub Finaliza:
         End If
         tb.MoveNext
         If tb.EOF Then
            GoSub Finaliza:
         End If
   Wend
End If


Finaliza:
    tb.Close
    bd.Close
    Set tb = Nothing
  Set bd = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Capac_Tanque_Descontam_OSV(WNumos As Long) As Double
On Error GoTo TratamentodeErro:
Dim WCapac As Double
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("DESCONTAM_TANQUE_OSV", DB_OPEN_TABLE)
tb.index = "NumOS"

WCapac = 0

tb.Seek "=", WNumos
If tb.NoMatch Then
   WCapac = 0
Else
   While tb!Numos = WNumos
         WCapac = Val(tb!Capac_Tanque)
         tb.MoveNext
         If tb.EOF Then
            GoSub Finaliza:
         End If
   Wend
End If


Finaliza:
    Capac_Tanque_Descontam_OSV = WCapac
    tb.Close
    bd.Close
    Set tb = Nothing
  Set bd = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Altera_NumTanque_Descontam(WNumos As Long, WNumTanque As String) As Long

On Error GoTo TratamentodeErro:
Dim WCapac As Double
Dim bd As DataBase
Dim tbc, tbf, tbp As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tbc = bd.OpenRecordset("DESCONTAM_TANQUE_OSV_COMPART", DB_OPEN_TABLE)
Set tbf = bd.OpenRecordset("DESCONTAM_TANQUE_OSV_FINALID", DB_OPEN_TABLE)
Set tbp = bd.OpenRecordset("DESCONTAM_TANQUE_OSV_PROC", DB_OPEN_TABLE)

tbc.index = "NumOS"
tbf.index = "NumOS"
tbp.index = "NumOS"

BeginTrans
GoSub Altera_Compart:
GoSub Altera_Finalidade:
GoSub Altera_Processos:

Finaliza:
    CommitTrans
    tbc.Close
    tbf.Close
    tbp.Close
    bd.Close
    Set tbc = Nothing
    Set tbf = Nothing
    Set tbp = Nothing
    Set bd = Nothing
    Exit Function
Return

Altera_Processos:
    tbp.Seek "=", WNumos
    If Not tbp.NoMatch Then
       While tbp!Numos = WNumos
             tbp.Edit
             tbp!Num_Tanque = WNumTanque
             tbp.Update
             tbp.MoveNext
             If tbp.EOF Then
                Return
             End If
       Wend
    End If
Return

Altera_Finalidade:
    tbf.Seek "=", WNumos
    If Not tbf.NoMatch Then
       While tbf!Numos = WNumos
             tbf.Edit
             tbf!Num_Tanque = WNumTanque
             tbf.Update
             tbf.MoveNext
             If tbf.EOF Then
                Return
             End If
       Wend
    End If
Return



Altera_Compart:
    tbc.Seek "=", WNumos
    If Not tbc.NoMatch Then
       While tbc!Numos = WNumos
             tbc.Edit
             tbc!Num_Tanque = WNumTanque
             tbc.Update
             tbc.MoveNext
             If tbc.EOF Then
                Return
             End If
       Wend
    End If
Return


TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Incluir_Finalidade_Descontam_OSV(WNumos As Long, WNum_Tanque As String, WFinalidade As Long) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("DESCONTAM_TANQUE_OSV_FINALID", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

BeginTrans
GoSub IncluiReg:

Finaliza:
    CommitTrans
    tb.Close
    bd.Close
    Set tb = Nothing
  Set bd = Nothing
    Exit Function
Return

IncluiReg:
    tb.Seek "=", WNumos, WNum_Tanque, WFinalidade
    If tb.NoMatch Then
       tb.AddNew
    Else
       tb.Edit
    End If
    tb!Numos = WNumos
    tb!Num_Tanque = WNum_Tanque
    tb!CodigoFinalidade = WFinalidade
    tb.Update
Return



TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function IncluirProcDescontTanque(WNumos As Long, WNum_Tanque As String, WCodigoProcesso As Long) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("DESCONTAM_TANQUE_OSV_PROC", DB_OPEN_TABLE)

BeginTrans
GoSub ExcluiRegAnt:
GoSub IncluiRegAtu:

Finaliza:
    CommitTrans
    tb.Close
    bd.Close
    Set tb = Nothing
  Set bd = Nothing
    Exit Function
Return

IncluiRegAtu:
    tb.index = "PrimaryKey"
    tb.Seek "=", WNumos, WNum_Tanque, WCodigoProcesso
    If tb.NoMatch Then
       tb.AddNew
    Else
       tb.Edit
    End If
    tb!Numos = WNumos
    tb!Num_Tanque = WNum_Tanque
    tb!CodigoProcesso = WCodigoProcesso
    tb.Update
Return


ExcluiRegAnt:
    tb.index = "Numos"
    tb.Seek "=", WNumos
    If Not tb.NoMatch Then
       While tb!Numos = WNumos
             tb.Delete
             tb.MoveNext
             If tb.EOF Then
                Return
             End If
       Wend
    End If
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function IncluirServicoxOSV(WCodigoServico As Long, WQuant As Double, WValorMaoObra As Double, WDesconto As Double, WNumos As Long) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("ITENSSERVOSV", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

BeginTrans
GoSub CriaServico:

Finaliza:
    CommitTrans
    tb.Close
    bd.Close
    Set tb = Nothing
    Set bd = Nothing
    Exit Function
Return

CriaServico:
   tb.Seek "=", WNumos, WCodigoServico
   If tb.NoMatch Then
      tb.AddNew
       tb!ValorMaoObra = WValorMaoObra
   Else
      tb.Edit
   End If
   tb!TipoServico = WCodigoServico
   tb!Numos = WNumos
   tb!Quantidade = WQuant
   tb!Desconto = WDesconto
   tb!Vlr_RetencaoISSQN = 0
   tb!Aliq_RetencaoISSQN = 0
   tb!MovimentaCusto = -1
   tb!Comissionado = -1
   tb!Tec_Exec = 0
   tb!Data_Ini = Date
   tb!Hora_Ini = Time()
   tb!Ocorr = 0
   tb!Aliq_Comis_Exec = 0
   
   tb.Update
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function


Function ProxCodItem() As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Dim WCodigo As Long
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("EST", DB_OPEN_TABLE)

tb.index = "PrimaryKey"

WCodigo = 1

tb.Seek "=", WCodigo

If Not tb.NoMatch Then
   While Not tb.NoMatch
      WCodigo = WCodigo + 1
      tb.Seek "=", WCodigo
   Wend
   ProxCodItem = WCodigo
Else
   ProxCodItem = WCodigo
End If

Finaliza:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CaminhoExiste(sCaminho As String) As Boolean
    On Error GoTo TratamentodeErro:
    'Testa se um caminho ou arquivo existe (ou � v�lido)
    
    'Caso deseja testar se um caminho existe, certifique-se que o
    'par�metro de entrada sCaminho TERMINE com uma \, como, por
    'exemplo, Debug.Print CaminhoExiste("c:\temp\felipe\"). Caso
    'contr�rio, a fun��o buscar� um arquivo chamado felipe, sem
    'extens�o, na pasta c:\temp.
    
    If Dir(sCaminho) = vbNullString Then
        CaminhoExiste = False
    Else
        CaminhoExiste = True
    End If
    
    'A forma abreviada da fun��o pode ser escrita como:
    'CaminhoExiste = Dir(sCaminho) <> vbNullString
    
    Exit Function
    
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function AbreArquivoLogSGA() As Long
'Shell "notepad " & CapturaCaminhoBD & "\LOGSGA.TXT", vbMaximizedFocus
DoCmd.OpenForm "Log_SGa_Princ", acNormal
End Function

Function LogSGa(WTabela As String, WFORM As String, WOcorr As String, WUsuario As String, WData As Variant, WHora As Variant, WNumReg As String, WEstacaotrabalho As String) As Long
'On Error GoTo TratamentodeErro:

On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim QRs As QueryDef
Dim rs  As Recordset
Set bd = CurrentDb()
Set QRs = bd.QueryDefs("Q_LOG_SGa")
QRs.Parameters(0) = WData
QRs.Parameters(1) = WHora
QRs.Parameters(2) = WUsuario
QRs.Parameters(3) = WTabela
QRs.Parameters(4) = WFORM
Set rs = QRs.OpenRecordset()

BeginTrans
If rs.RecordCount = 0 Then
   rs.AddNew
Else
   rs.Edit
End If
rs!data_log = WData
rs!hora_log = WHora
rs!usuario_log = WUsuario
rs!tabela_log = WTabela
rs!formulario_log = WFORM
rs!ocorrencia_log = WOcorr
rs!registro_Alt_log = WNumReg
rs!maquina_Log = WEstacaotrabalho
rs.Update
CommitTrans
Finaliza:
   rs.Close
   bd.Close
   Set rs = Nothing
   Set bd = Nothing
   Exit Function
Return


'Dim WLog As String
'Dim WTraco As String * 195
    
'WLog = CapturaCaminhoBD() & "\LOGSGa.txt"

'If CaminhoExiste(WLog) Then
'    Open CapturaCaminhoBD() & "\LOGSGa.txt" For Append As #10
'Else
'    WTraco = "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
'    Open CapturaCaminhoBD() & "\LOGSGa.txt" For Output As #10
'    Print #10, WTraco
'    Print #10, Format$("Data      ", "!@@@@@@@@") & Format$("Hora", "!@@@@@@@@@") & Format$("Tabela", "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") & " " & Format$("Registro", "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") & Format$("Formulario", "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") & " " & Format$("Ocorrencia", "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") & " " & Format$("Usuario", "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") & " " & Format$(" Estacao de trabalho", "!@@@@@@@@")
'    Print #10, WTraco
'End If

'Print #10, Format$(wDATA, "!@@@@@@@@") & ";" & Format$(wHora, "!@@@@@@@@") & ";" & Format$(WTabela, "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") & ";" & Format$(WNumReg, "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") & ";" & Format$(WFORM, "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") & ";" & Format$(WOcorr, "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") & ";" & Format$(WUsuario, "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") & ";" & Format$(WEstacaotrabalho, "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")

'Close #10


TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Operacao_Reemb_ContrMetr() As Long
On Error GoTo TratamentodeErro:
Dim Dbtab As DataBase
Dim emp As Recordset

Set Dbtab = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set emp = Dbtab.OpenRecordset("Empresa", DB_OPEN_TABLE)
emp.index = "PrimaryKey"

If emp.RecordCount <> 0 Then
   Do Until emp.EOF
      If Form_Menu.WEmpresa = emp!CodigoEmpresa Then
         If IsNull(emp!ControladorMetrologico) Then
            Operacao_Reemb_ContrMetr = 0
         Else
            If IsNull(emp!Operacao_Reemb) Then
               Operacao_Reemb_ContrMetr = 0
            Else
               Operacao_Reemb_ContrMetr = emp!Operacao_Reemb
            End If
         End If
      End If
   
      emp.MoveNext
   Loop
End If
        
Finaliza:
    emp.Close
    Dbtab.Close
    Set emp = Nothing
    Set Dbtab = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function


Function RetornaServicoXReembXOSV(WNumos As Long) As Double

On Error GoTo TratamentodeErro:
Dim db As DataBase
Dim tb As Recordset
Dim WTotalReemb As Double

Set db = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = db.OpenRecordset("ITENSSERVOSV", DB_OPEN_TABLE)
tb.index = "SecondaryKey"
WTotalReemb = 0
tb.Seek "=", WNumos

If tb.NoMatch Then
   RetornaServicoXReembXOSV = 0
Else
   While Not tb.EOF
         If tb!Numos <> WNumos Then
            GoSub Finaliza:
         End If
         WTotalReemb = WTotalReemb + VlrReembServXEmp(tb!TipoServico)
         tb.MoveNext
         If tb.EOF Then
            GoSub Finaliza:
         End If
   Wend
End If
   
Finaliza:
    RetornaServicoXReembXOSV = WTotalReemb
    tb.Close
    db.Close
    Set tb = Nothing
    Set db = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Total_Vlr_Repasse_OSV(WNumos As Long) As Double
On Error GoTo TratamentodeErro:
Dim db As DataBase
Dim tb, tbo As Recordset
Dim WValorRepasse As Double

Set db = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = db.OpenRecordset("ITENSSERVOSV", DB_OPEN_TABLE)
Set tbo = db.OpenRecordset("OCORR_SERV", DB_OPEN_TABLE)
tb.index = "SecondaryKey"
tbo.index = "PrimaryKey"
tb.Seek "=", WNumos

If tb.NoMatch Then
   Total_Vlr_Repasse_OSV = 0
Else
  While tb!Numos = WNumos
        tbo.Seek "=", tb!Ocorr
        If Not tbo.NoMatch Then
           If tbo!Gera_Financeiro Then
              WValorRepasse = WValorRepasse + VlrRepasseXEmp(tb!TipoServico, (tb!ValorMaoObra - tb!Desconto))
           End If
        End If
        tb.MoveNext
        If tb.EOF Then
           GoSub Finaliza:
        End If
  
  Wend
End If
   
Finaliza:
    tb.Close
    tbo.Close
    db.Close
    Set tb = Nothing
    Set tbo = Nothing
    Set db = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function VlrRepasseXEmp(WCodigoServico As Long, WValorServico As Double) As Double
On Error GoTo TratamentodeErro:
Dim db As DataBase
Dim tb As Recordset

Set db = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = db.OpenRecordset("EMPXTPSERVXREPASSE", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

tb.Seek "=", Form_Menu.WEmpresa, WCodigoServico

If tb.NoMatch Then
   VlrRepasseXEmp = 0
Else
   If IsNull(tb!Repasse) Then
      VlrRepasseXEmp = 0
   Else
      VlrRepasseXEmp = Round((WValorServico * tb!Repasse) / 100)
   End If
End If
   
Finaliza:
    tb.Close
    db.Close
    Set tb = Nothing
    Set db = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function VlrReembServXEmp(WCodigoServico As Long) As Double

On Error GoTo TratamentodeErro:
Dim Dbtab As DataBase
Dim emp As Recordset

Set Dbtab = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set emp = Dbtab.OpenRecordset("EMPXTPSERVXINMETRO", DB_OPEN_TABLE)
emp.index = "PrimaryKey"

emp.Seek "=", Form_Menu.WEmpresa, WCodigoServico

If emp.NoMatch Then
   VlrReembServXEmp = 0
Else
   If IsNull(emp!Reemb) Then
      VlrReembServXEmp = 0
   Else
      VlrReembServXEmp = emp!Reemb
   End If
End If
   
Finaliza:
    emp.Close
    Dbtab.Close
    Set emp = Nothing
    Set Dbtab = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function AtualizaLoginInMetro(wLoginInMetro As String, WSenhaInMetro As String) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
   
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("EMPRESA", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

BeginTrans
tb.Seek "=", Form_Menu.WEmpresa
If Not tb.NoMatch Then
   tb.Edit
Else
   tb.AddNew
End If
tb!LoginInMetro = wLoginInMetro
tb!SenhaInMetro = WSenhaInMetro
tb.Update
CommitTrans

Finaliza:
    tb.Close
    bd.Close
    Set tb = Nothing
    Set bd = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ExcluiReembContrMetr(WNumos As Long, WDataFechamento As Variant, WCodigoCliente As Long) As Long

On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
   
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("CTREC", DB_OPEN_TABLE)
tb.index = "SetimaChave"

BeginTrans
tb.Seek "=", WNumos, WDataFechamento, WCodigoCliente
If Not tb.NoMatch Then
   If Form_Menu.WEmpresa = tb!CodEmpresa Then
      tb.Delete
   End If
End If
CommitTrans

Finaliza:
    tb.Close
    bd.Close
    Set tb = Nothing
    Set bd = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function RetornaCodigoFornecXCNPJ(WCNPJ As String) As Long
    On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim tb As Recordset
    Dim WTotalEstoque As Double
    
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd1.OpenRecordset("FOR", DB_OPEN_TABLE)
    
    tb.index = "CGCCPF"
    tb.Seek "=", Trim(WCNPJ)
        
    If tb.NoMatch Then
      RetornaCodigoFornecXCNPJ = 0
    Else
      RetornaCodigoFornecXCNPJ = tb!CodigoFornecedor
    End If
        
Finaliza:
    tb.Close
    bd1.Close
    Set tb = Nothing
    Set bd1 = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ApuraChassixOSVxCliente(WCodigoCliente As Long, wPlaca As String) As String
 On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim tb As Recordset
    Dim WTotalEstoque As Double
    
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd1.OpenRecordset("PLACXCLI", DB_OPEN_TABLE)
    
    tb.index = "PrimaryKey"
    tb.Seek "=", wPlaca, WCodigoCliente
        
    If tb.NoMatch Then
      ApuraChassixOSVxCliente = "1"
      GoSub Finaliza:
    Else
      ApuraChassixOSVxCliente = "1"
       If Not IsNull(tb!CHASSI) Then
         If tb!CHASSI <> " " And tb!CHASSI <> Space(1) Then
            ApuraChassixOSVxCliente = UCase(tb!CHASSI)
         Else
            ApuraChassixOSVxCliente = ""
         End If
      End If
    End If
        
Finaliza:
    tb.Close
    bd1.Close
    Set tb = Nothing
    Set bd1 = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function VinculaESTXAPLIC(waplic As String) As Long
 On Error GoTo TratamentodeErro:
    Dim bd As DataBase
    Dim tb, tbe As Recordset
    Dim WTotalEstoque As Double
    
    Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd.OpenRecordset("ESTXAPLIC", DB_OPEN_TABLE)
    Set tbe = bd.OpenRecordset("EST", DB_OPEN_TABLE)
    tb.index = "PrimaryKey"
    tbe.index = "PrimaryKey"
    BeginTrans
    If tbe.RecordCount <> 0 Then
        While Not tbe.EOF
              If tbe!Selecionar Then
                 tb.Seek "=", tbe!CodItem, Trim$(waplic)
                 If tb.NoMatch Then
                    tb.AddNew
                 Else
                    tb.Edit
                 End If
                 tb!CodItem = tbe!CodItem
                 tb!Aplic = waplic
                 tb!Observ = Space(1)
                 tb.Update
              End If
              tbe.MoveNext
        Wend
    End If
    
Finaliza:
    CommitTrans
    MsgBox "Fim Normal de Execu��o...", vbInformation, "Aviso"
    DoCmd.Close acForm, "CadastroAplicVincEstPrinc"
    tbe.Close
    tb.Close
    bd.Close
    Set tb = Nothing
    Set tbe = Nothing
    Set bd = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function


Function ApuraPneuxOSVxCliente(WCodigoCliente As Long, wPlaca As String) As String
 On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim tb As Recordset
    Dim WTotalEstoque As Double
    
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd1.OpenRecordset("PLACXCLI", DB_OPEN_TABLE)
    
    tb.index = "PrimaryKey"
    tb.Seek "=", wPlaca, WCodigoCliente
    
    
    If tb.NoMatch Then
      ApuraPneuxOSVxCliente = "1"
      GoSub Finaliza:
    Else
     ApuraPneuxOSVxCliente = "1"
      If Not IsNull(tb!Pneu) Then
         ApuraPneuxOSVxCliente = UCase(tb!Pneu)
      End If
    End If
        
Finaliza:
    tb.Close
    bd1.Close
    Set tb = Nothing
    Set bd1 = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ApuraPlacAuxxOSVxCliente(WCodigoCliente As Long, wPlaca As String) As String
 On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim tb As Recordset
    Dim WTotalEstoque As Double
    
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd1.OpenRecordset("PLACXCLI", DB_OPEN_TABLE)
    
    tb.index = "PrimaryKey"
    tb.Seek "=", wPlaca, WCodigoCliente
    
    
    If tb.NoMatch Then
      ApuraPlacAuxxOSVxCliente = "1"
      GoSub Finaliza:
    Else
      ApuraPlacAuxxOSVxCliente = "1"
      If Not IsNull(tb!PlacaVeiculoAux) Then
         ApuraPlacAuxxOSVxCliente = UCase(tb!PlacaVeiculoAux)
      End If
    End If
        
Finaliza:
    tb.Close
    bd1.Close
    Set tb = Nothing
    Set bd1 = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ApuraMarcVeicxOSVxCliente(WCodigoCliente As Long, wPlaca As String) As String
 On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim tb As Recordset
    Dim WTotalEstoque As Double
    
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd1.OpenRecordset("PLACXCLI", DB_OPEN_TABLE)
    
    tb.index = "PrimaryKey"
    tb.Seek "=", wPlaca, WCodigoCliente
    
    
    If tb.NoMatch Then
      ApuraMarcVeicxOSVxCliente = "1"
      GoSub Finaliza:
    Else
     ApuraMarcVeicxOSVxCliente = "1"
      If Not IsNull(tb!MarcaVeiculo) Then
         ApuraMarcVeicxOSVxCliente = UCase(tb!MarcaVeiculo)
      End If
    End If
        
Finaliza:
    tb.Close
    bd1.Close
    Set tb = Nothing
    Set bd1 = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function



Function ApuraAnoVeicxOSVxCliente(WCodigoCliente As Long, wPlaca As String) As String
 On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim tb As Recordset
    Dim WTotalEstoque As Double
    
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd1.OpenRecordset("PLACXCLI", DB_OPEN_TABLE)
    
    tb.index = "PrimaryKey"
    tb.Seek "=", wPlaca, WCodigoCliente
    
    
    If tb.NoMatch Then
      ApuraAnoVeicxOSVxCliente = "1"
      GoSub Finaliza:
    Else
        ApuraAnoVeicxOSVxCliente = "1"
       If Not IsNull(tb!AnoVeiculo) Then
          ApuraAnoVeicxOSVxCliente = UCase(tb!AnoVeiculo)
       End If
    End If
        
Finaliza:
    tb.Close
    bd1.Close
    Set tb = Nothing
    Set bd1 = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ApuraModeloVeiculoXOSV(WNumos As Long) As String
On Error GoTo TratamentodeErro:
    Dim bd As DataBase
    Dim tb As Recordset
    Dim WTotalEstoque As Double
    
    Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
    
    tb.index = "PrimaryKey"
    tb.Seek "=", WNumos
    
    
    If tb.NoMatch Then
      ApuraModeloVeiculoXOSV = Empty
    Else
     If IsNull(tb!ModeloVeiculo) Then
        ApuraModeloVeiculoXOSV = Empty
     Else
      ApuraModeloVeiculoXOSV = UCase(tb!ModeloVeiculo)
     End If
    End If
        
Finaliza:
    tb.Close
    bd.Close
    Set tb = Nothing
    Set bd = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ApuraMarcaVeiculoXOSV(WNumos As Long) As String
On Error GoTo TratamentodeErro:
    Dim bd As DataBase
    Dim tb As Recordset
    Dim WTotalEstoque As Double
    
    Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
    
    tb.index = "PrimaryKey"
    tb.Seek "=", WNumos
    
    
    If tb.NoMatch Then
      ApuraMarcaVeiculoXOSV = Empty
    Else
     If IsNull(tb!MarcaVeiculo) Then
        ApuraMarcaVeiculoXOSV = Empty
     Else
      ApuraMarcaVeiculoXOSV = UCase(tb!MarcaVeiculo)
     End If
    End If
        
Finaliza:
    tb.Close
    bd.Close
    Set tb = Nothing
    Set bd = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ApuraCodigoPistalEnsaioXOSV(WNumos As Long) As String
On Error GoTo TratamentodeErro:
    Dim bd As DataBase
    Dim tb As Recordset
    Dim WTotalEstoque As Double
    
    Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
    
    tb.index = "PrimaryKey"
    tb.Seek "=", WNumos
    
    
    If tb.NoMatch Then
      ApuraCodigoPistalEnsaioXOSV = Space(1)
    Else
     If IsNull(tb!TipoPistaEnsaio) Then
        ApuraCodigoPistalEnsaioXOSV = Space(1)
     Else
      ApuraCodigoPistalEnsaioXOSV = tb!TipoPistaEnsaio
     End If
    End If
        
Finaliza:
    tb.Close
    bd.Close
    Set tb = Nothing
    Set bd = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ApuraVelMaxEnsXOSV(WNumos As Long) As String
On Error GoTo TratamentodeErro:
    Dim bd As DataBase
    Dim tb As Recordset
    Dim WTotalEstoque As Double
    
    Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
    
    tb.index = "PrimaryKey"
    tb.Seek "=", WNumos
      
    If tb.NoMatch Then
      ApuraVelMaxEnsXOSV = ""
    Else
     If IsNull(tb!VelMaxEns) Then
        ApuraVelMaxEnsXOSV = ""
     Else
      ApuraVelMaxEnsXOSV = tb!VelMaxEns
     End If
    End If
        
Finaliza:
    tb.Close
    bd.Close
    Set tb = Nothing
    Set bd = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function GravaGRUXOSV(WNumos As Long, WGRU As String) As Long
On Error GoTo TratamentodeErro:
    Dim bd As DataBase
    Dim tb As Recordset
    Dim WTotalEstoque As Double
    
    Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
    
    tb.index = "PrimaryKey"
    
    BeginTrans
    tb.Seek "=", WNumos
    If Not tb.NoMatch Then
       tb.Edit
       tb!GRU = WGRU
       tb.Update
    End If
        
Finaliza:
    CommitTrans
    tb.Close
    bd.Close
    Set tb = Nothing
    Set bd = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ApuraAroPneuXOSV(WNumos As Long) As String
On Error GoTo TratamentodeErro:
    Dim bd As DataBase
    Dim tb As Recordset
    Dim WTotalEstoque As Double
    
    Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
    
    tb.index = "PrimaryKey"
    tb.Seek "=", WNumos
    
    
    If tb.NoMatch Then
      ApuraAroPneuXOSV = Empty
    Else
     If IsNull(tb!AroPneu) Then
        ApuraAroPneuXOSV = Empty
     Else
      ApuraAroPneuXOSV = UCase(tb!AroPneu)
     End If
    End If
        
Finaliza:
    tb.Close
    bd.Close
    Set tb = Nothing
    Set bd = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ApuraTipoLaudoXOSV(WNumos As Long) As String
On Error GoTo TratamentodeErro:
    Dim bd As DataBase
    Dim tb As Recordset
    Dim WTotalEstoque As Double
    
    Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
    
    tb.index = "PrimaryKey"
    tb.Seek "=", WNumos
    
    
    If tb.NoMatch Then
      ApuraTipoLaudoXOSV = "Z"
    Else
     If IsNull(tb!TipoLaudo) Then
        ApuraTipoLaudoXOSV = "Z"
     Else
        ApuraTipoLaudoXOSV = UCase(tb!TipoLaudo)
     End If
    End If
        
Finaliza:
    tb.Close
    bd.Close
    Set tb = Nothing
    Set bd = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ApuraPneuXOSV(WNumos As Long) As String
On Error GoTo TratamentodeErro:
    Dim bd As DataBase
    Dim tb As Recordset
    Dim WTotalEstoque As Double
    
    Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
    
    tb.index = "PrimaryKey"
    tb.Seek "=", WNumos
   
    If tb.NoMatch Then
      ApuraPneuXOSV = Space(1)
    Else
     If IsNull(tb!Pneu) Then
        ApuraPneuXOSV = Empty
     Else
      ApuraPneuXOSV = UCase(tb!Pneu)
     End If
    End If
        
Finaliza:
    tb.Close
    bd.Close
    Set tb = Nothing
    Set bd = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Retorna_Marca_Taximetro(WCodmarca As Long) As String
On Error GoTo TratamentodeErro:
    Dim bd As DataBase
    Dim tb As Recordset
    Dim WTotalEstoque As Double
    
    Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd.OpenRecordset("MARCA_TAXIM", DB_OPEN_TABLE)
    
    tb.index = "PrimaryKey"
    tb.Seek "=", Val(WCodmarca)
   
    If tb.NoMatch Then
      Retorna_Marca_Taximetro = "INDEFINIDO"
    Else
     If IsNull(tb!Descricao) Then
        Retorna_Marca_Taximetro = "INDEFINIDO"
     Else
      Retorna_Marca_Taximetro = UCase(tb!Descricao)
     End If
    End If
        
Finaliza:
    tb.Close
    bd.Close
    Set tb = Nothing
    Set bd = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function


Function ApuraAnoVeiculoXOSV(WNumos As Long) As String
On Error GoTo TratamentodeErro:
    Dim bd As DataBase
    Dim tb As Recordset
    Dim WTotalEstoque As Double
    
    Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
    
    tb.index = "PrimaryKey"
    tb.Seek "=", WNumos
    
    
    If tb.NoMatch Then
      ApuraAnoVeiculoXOSV = Empty
    Else
     If IsNull(tb!AnoVeiculo) Then
        ApuraAnoVeiculoXOSV = Empty
     Else
      ApuraAnoVeiculoXOSV = UCase(tb!AnoVeiculo)
     End If
    End If
        
Finaliza:
    tb.Close
    bd.Close
    Set tb = Nothing
    Set bd = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ApuraDtEnsaioXOSV(WNumos As Long) As Variant
On Error GoTo TratamentodeErro:
    Dim bd As DataBase
    Dim tb As Recordset
    Dim WTotalEstoque As Double
    
    Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
    
    tb.index = "PrimaryKey"
    tb.Seek "=", WNumos
   
    If tb.NoMatch Then
      ApuraDtEnsaioXOSV = ""
    Else
     If IsNull(tb!DTEnsaio) Then
        ApuraDtEnsaioXOSV = ""
     Else
      ApuraDtEnsaioXOSV = UCase(tb!DTEnsaio)
     End If
    End If
        
Finaliza:
    tb.Close
    bd.Close
    Set tb = Nothing
    Set bd = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function ApuraChassiXOSV(WNumos As Long) As String
On Error GoTo TratamentodeErro:
    Dim bd As DataBase
    Dim tb As Recordset
    Dim WTotalEstoque As Double
    
    Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
    
    tb.index = "PrimaryKey"
    tb.Seek "=", WNumos
    
    
    If tb.NoMatch Then
      ApuraChassiXOSV = Space(1)
    Else
     If IsNull(tb!CHASSI) Then
        ApuraChassiXOSV = Space(1)
     Else
      ApuraChassiXOSV = UCase(tb!CHASSI)
     End If
    End If
        
Finaliza:
    tb.Close
    bd.Close
    Set tb = Nothing
    Set bd = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ApuraModVeicxOSVxCliente(WCodigoCliente As Long, wPlaca As String) As String
 On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim tb As Recordset
    Dim WTotalEstoque As Double
    
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd1.OpenRecordset("PLACXCLI", DB_OPEN_TABLE)
    
    tb.index = "PrimaryKey"
    tb.Seek "=", wPlaca, WCodigoCliente
    
    
    If tb.NoMatch Then
      ApuraModVeicxOSVxCliente = "1"
      GoSub Finaliza:
    Else
    ApuraModVeicxOSVxCliente = "1"
      If Not IsNull(tb!ModeloVeiculo) Then
         ApuraModVeicxOSVxCliente = UCase(tb!ModeloVeiculo)
      End If
    End If
        
Finaliza:
    tb.Close
    bd1.Close
    Set tb = Nothing
    Set bd1 = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ApuraFrotaVeicxOSVxCliente(WCodigoCliente As Long, wPlaca As String) As String
 On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim tb As Recordset
    Dim WTotalEstoque As Double
    
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd1.OpenRecordset("PLACXCLI", DB_OPEN_TABLE)
    
    tb.index = "PrimaryKey"
    tb.Seek "=", wPlaca, WCodigoCliente
    
    
    If tb.NoMatch Then
      ApuraFrotaVeicxOSVxCliente = "1"
      GoSub Finaliza:
    Else
      ApuraFrotaVeicxOSVxCliente = "1"
       If Not IsNull(tb!FROTA) Then
          ApuraFrotaVeicxOSVxCliente = UCase(tb!FROTA)
       End If
    End If
        
Finaliza:
    tb.Close
    bd1.Close
    Set tb = Nothing
    Set bd1 = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ApuraRenavanXOSV(WNumos As Long) As String
On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim tb As Recordset
    Dim WTotalEstoque As Double
    
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
    
    tb.index = "PrimaryKey"
    tb.Seek "=", WNumos
    
    If Not tb.NoMatch Then
       If Not IsNull(tb!RENAVAN) Then
          ApuraRenavanXOSV = tb!RENAVAN
       Else
          ApuraRenavanXOSV = Empty
       End If
    End If
        
Finaliza:
    tb.Close
    bd1.Close
    Set tb = Nothing
    Set bd1 = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ApuraCor_VeiculoXOSV(WNumos As Long) As String
On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim tb As Recordset
    Dim WTotalEstoque As Double
    
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
    
    tb.index = "PrimaryKey"
    tb.Seek "=", WNumos
    
    If Not tb.NoMatch Then
       If Not IsNull(tb!COR_Veiculo) Then
          ApuraCor_VeiculoXOSV = tb!COR_Veiculo
       Else
          ApuraCor_VeiculoXOSV = Empty
       End If
    End If
        
Finaliza:
    tb.Close
    bd1.Close
    Set tb = Nothing
    Set bd1 = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ApuraRENAVeicxOSVxCliente(WCodigoCliente As Long, wPlaca As String) As String
 On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim tb As Recordset
    Dim WTotalEstoque As Double
    
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd1.OpenRecordset("PLACXCLI", DB_OPEN_TABLE)
    
    tb.index = "PrimaryKey"
    tb.Seek "=", wPlaca, WCodigoCliente
    
    
    If tb.NoMatch Then
      ApuraRENAVeicxOSVxCliente = "1"
      GoSub Finaliza:
    Else
     ApuraRENAVeicxOSVxCliente = "1"
      If Not IsNull(tb!RENAVAN) Then
         ApuraRENAVeicxOSVxCliente = UCase(tb!RENAVAN)
      End If
    End If
        
Finaliza:
    tb.Close
    bd1.Close
    Set tb = Nothing
    Set bd1 = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ApuraCodClienteXPlacxCli(wPlaca As String) As Long
 On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim tb As Recordset
    Dim WTotalEstoque As Double
    
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd1.OpenRecordset("PLACXCLI", DB_OPEN_TABLE)
    
    tb.index = "SecondKey"
    tb.Seek "=", wPlaca
        
    If tb.NoMatch Then
       ApuraCodClienteXPlacxCli = 0
       GoSub Finaliza:
    Else
       ApuraCodClienteXPlacxCli = 0
       If Not IsNull(tb!CodigoCliente) Then
          ApuraCodClienteXPlacxCli = tb!CodigoCliente
       End If
    End If
        
Finaliza:
    tb.Close
    bd1.Close
    Set tb = Nothing
    Set bd1 = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ApuraAroVeicxOSVxCliente(WCodigoCliente As Long, wPlaca As String) As String
 On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim tb As Recordset
    Dim WTotalEstoque As Double
    
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd1.OpenRecordset("PLACXCLI", DB_OPEN_TABLE)
    
    tb.index = "PrimaryKey"
    tb.Seek "=", wPlaca, WCodigoCliente
    
    
    If tb.NoMatch Then
      ApuraAroVeicxOSVxCliente = "1"
      GoSub Finaliza:
    Else
     ApuraAroVeicxOSVxCliente = "1"
      If Not IsNull(tb!AroPneu) Then
         ApuraAroVeicxOSVxCliente = UCase(tb!AroPneu)
      End If
    End If
        
Finaliza:
    tb.Close
    bd1.Close
    Set tb = Nothing
    Set bd1 = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ApuraCorVeicxOSVxCliente(WCodigoCliente As Long, wPlaca As String) As String
 On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim tb As Recordset
    Dim WTotalEstoque As Double
    
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd1.OpenRecordset("PLACXCLI", DB_OPEN_TABLE)
    
    tb.index = "PrimaryKey"
    tb.Seek "=", wPlaca, WCodigoCliente
    
    
    If tb.NoMatch Then
      ApuraCorVeicxOSVxCliente = "1"
      GoSub Finaliza:
    Else
     ApuraCorVeicxOSVxCliente = "1"
      If Not IsNull(tb!COR_Veiculo) Then
         ApuraCorVeicxOSVxCliente = UCase(tb!COR_Veiculo)
      End If
    End If
        
Finaliza:
    tb.Close
    bd1.Close
    Set tb = Nothing
    Set bd1 = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function


Function ApuraTotalEstXOSVXNFE(wCodCliente As Long) As Double

    On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim tb As Recordset
    Dim WTotalEstoque As Double
    
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
    
    tb.index = "SecondKey"
    WTotalEstoque = 0
    tb.Seek "=", wCodCliente
    
    If tb.NoMatch Then
       GoSub Finaliza:
    Else
    
       While Not tb.EOF
             If tb!CodigoCliente <> wCodCliente Then
                GoSub Finaliza:
             End If
             
             If tb!SelecionaOSV Then
                WTotalEstoque = WTotalEstoque + CalculoMercadorias(tb!Numos)
             End If
             
             tb.MoveNext
       Wend
    End If
        
Finaliza:
    tb.Close
    ApuraTotalEstXOSVXNFE = WTotalEstoque
    bd1.Close
    Set tb = Nothing
    Set bd1 = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function


Function Inicializa_TecxCatxEst(WCodigoFunc As Long) As Long

    On Error GoTo TratamentodeErro:
    Dim bd As DataBase
    Dim tb As Recordset
    
    Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd.OpenRecordset("TECXCATXEST", DB_OPEN_TABLE)
    tb.index = "CodigoTec"
    BeginTrans
    tb.Seek "=", WCodigoFunc
    
    If Not tb.NoMatch Then
       While tb!CodigoTec = WCodigoFunc
             tb.Delete
             tb.MoveNext
       Wend
    End If
        
Finaliza:
    CommitTrans
    tb.Close
    bd.Close
    Set tb = Nothing
    Set bd = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Inicializa_TecxServ(WCodigoFunc As Long) As Long

    On Error GoTo TratamentodeErro:
    Dim bd As DataBase
    Dim tb As Recordset
    
    Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd.OpenRecordset("TECXSERV", DB_OPEN_TABLE)
    tb.index = "CodigoTec"
    BeginTrans
    tb.Seek "=", WCodigoFunc
    
    If Not tb.NoMatch Then
       While tb!CodigoTec = WCodigoFunc
             tb.Delete
             tb.MoveNext
             If tb.EOF Then
                GoSub Finaliza:
             End If
       Wend
    End If
        
Finaliza:
    CommitTrans
    tb.Close
    bd.Close
    Set tb = Nothing
    Set bd = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CriaTecXServ(WCodigoFunc As Long) As Long
    On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim tb, tb1 As Recordset
    Dim WTotalEstoque As Double
    
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd1.OpenRecordset("TECXSERV", DB_OPEN_TABLE)
    Set tb1 = bd1.OpenRecordset("TPSERV", DB_OPEN_TABLE)
    
    tb.index = "PrimaryKey"
    tb1.index = "PrimaryKey"
    
    If tb1.RecordCount <> 0 Then
       While Not tb1.EOF
            If Captura_Conteudo_Campo_TPSERVXEMP(tb1!TipoServico, "Ativo") Then
               tb.Seek "=", WCodigoFunc, tb1!TipoServico
               If tb.NoMatch Then
                  tb.AddNew
                  tb!Comissao = 0
               Else
                  tb.Edit
               End If
               tb!CodigoTec = WCodigoFunc
               tb!CodigoServico = tb1!TipoServico
               tb.Update
            End If
               tb1.MoveNext
       Wend
    End If
        
Finaliza:
    tb.Close
    tb1.Close
    bd1.Close
    Set tb1 = Nothing
    Set bd1 = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function


Function CapturaPerComisServ(WCodigoFunc As Long, WCodigoServ As Long, WNumos As Long) As Double
    On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim tb As Recordset
    
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd1.OpenRecordset("TECXSERVXOSV", DB_OPEN_TABLE)
    
    tb.index = "PrimaryKey"
    tb.Seek "=", WCodigoFunc, WCodigoServ, WNumos
    
    If tb.NoMatch Then
       CapturaPerComisServ = 0
       GoSub Finaliza:
    Else
       If IsNull(tb!Comissao) Then
           CapturaPerComisServ = 0
       Else
           CapturaPerComisServ = tb!Comissao
       End If
    End If
        
Finaliza:
    tb.Close
    bd1.Close
    Set tb = Nothing
    Set bd1 = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function


'Function CapturaPerComisCatEst(WCodigoFunc As Long, WCodigoCat As Long, WNumos As Long) As Double
'    On Error GoTo TratamentodeErro:
'    Dim bd1 As DataBase
'    Dim tb As Recordset
'
'    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
'    Set tb = bd1.OpenRecordset("TECXCATXOSV", DB_OPEN_TABLE)
    
'    tb.index = "PrimaryKey"
'    tb.Seek "=", WCodigoFunc, WCodigoCat, WNumos
    
    
'    If tb.NoMatch Then
'       CapturaPerComisCatEst = 0
'       GoSub Finaliza:
'    Else
'       If IsNull(tb!Comissao) Then
'           CapturaPerComisCatEst = 0
'       Else
'           CapturaPerComisCatEst = tb!Comissao
'       End If
'    End If
        
'Finaliza:
'    tb.Close
'    bd1.Close
'    Set tb = Nothing
'    Set bd1 = Nothing
'    Exit Function
'Return

'TratamentodeErro:
'    If Err.Number = 3001 Then
'       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
'    Else
'       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description, Vbcritical,"SGa"
'       Resume Next
'    End If'

'End Function

Function ValidaComissoesCatESt(WCodigoFunc As Long, WCodigoCat As Long) As Double

    On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim tb As Recordset
    
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd1.OpenRecordset("TECXCATXEST", DB_OPEN_TABLE)
    
    tb.index = "PrimaryKey"
    tb.Seek "=", WCodigoFunc, WCodigoCat
   
    If tb.NoMatch Then
        ValidaComissoesCatESt = 0
       GoSub Finaliza:
    Else
       ValidaComissoesCatESt = 1
    End If
        
Finaliza:
    tb.Close
    bd1.Close
    Set tb = Nothing
    Set bd1 = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function



Function ApuraTotalServXOSVXNFE(wCodCliente As Long) As Double

    On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim tb As Recordset
    Dim WTotalServicos As Double
    
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
    
    tb.index = "SecondKey"
    WTotalServicos = 0
    tb.Seek "=", wCodCliente
    
    
    If tb.NoMatch Then
       GoSub Finaliza:
    Else
    
       While Not tb.EOF
             If tb!CodigoCliente <> wCodCliente Then
                GoSub Finaliza:
             End If
             
             If tb!SelecionaOSV Then
                WTotalServicos = WTotalServicos + TotalServicosOrdemServico(tb!Numos)
             End If
             
             tb.MoveNext
       Wend
    End If
        
Finaliza:
    tb.Close
    ApuraTotalServXOSVXNFE = WTotalServicos
    bd1.Close
    Set tb = Nothing
    Set bd1 = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function



Function CodigoVagoCliXRef() As Long
Dim db  As DataBase, Cli As Recordset, WCod As Long
On Error GoTo TratamentodeErro:
Set db = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Cli = db.OpenRecordset("CliXRef", DB_OPEN_TABLE)
Cli.index = "SecondKey"
WCod = 1
Cli.Seek "=", WCod
If Not Cli.NoMatch Then
   Do Until Cli.NoMatch
      WCod = WCod + 1
      Cli.Seek "=", WCod
   Loop
End If

Finaliza:
  CodigoVagoCliXRef = WCod
  Cli.Close
  db.Close
  Set Cli = Nothing
  Set db = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function



Function DesmarcaEST() As Long
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim X As Long, Est As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Est = bd1.OpenRecordset("EST", DB_OPEN_TABLE)
Est.index = "PrimaryKey"

BeginTrans
If Est.RecordCount <> 0 Then
      While Not Est.EOF
               Est.Edit
               Est!Selecionar = 0
               Est.Update
               X = X + 1
            Est.MoveNext
      Wend
End If
CommitTrans

Finaliza:
  Est.Close
  bd1.Close
  Set Est = Nothing
  Set bd1 = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function DesmarcaCli() As Long
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim X As Long, Cli As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
Cli.index = "PrimaryKey"

BeginTrans
If Cli.RecordCount <> 0 Then
      While Not Cli.EOF
            Cli.Edit
            Cli!Selecionar = 0
            Cli.Update
            X = X + 1
            If X Mod 500 = 0 Then
               CommitTrans
               BeginTrans
            End If

            Cli.MoveNext
      Wend
End If
CommitTrans

Finaliza:
  Cli.Close
  bd1.Close
  Set Cli = Nothing
  Set bd1 = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function MarcaCli() As Long
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim X As Long, Cli As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
Cli.index = "PrimaryKey"

BeginTrans
If Cli.RecordCount <> 0 Then
      While Not Cli.EOF
            Cli.Edit
            Cli!Selecionar = 1
            Cli.Update
            X = X + 1
            If X Mod 500 = 0 Then
               CommitTrans
               BeginTrans
            End If
            Cli.MoveNext
      Wend
End If
CommitTrans

Finaliza:
  Cli.Close
  bd1.Close
  Set Cli = Nothing
  Set bd1 = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function MarcaOSV() As Long
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim X As Long, OSV As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set OSV = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
OSV.index = "PrimaryKey"

BeginTrans
If OSV.RecordCount <> 0 Then
      While Not OSV.EOF
            OSV.Edit
            OSV!SelecionaOSV = 1
            OSV.Update
            X = X + 1
            If X Mod 500 = 0 Then
               CommitTrans
               BeginTrans
            End If
            OSV.MoveNext
      Wend
End If
CommitTrans

Finaliza:
  OSV.Close
  bd1.Close
  Set OSV = Nothing
  Set bd1 = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function DesmarcaCategorias() As Long
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim X As Long, Est As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Est = bd1.OpenRecordset("CATPROD", DB_OPEN_TABLE)
Est.index = "PrimaryKey"

BeginTrans
If Est.RecordCount <> 0 Then
      While Not Est.EOF
            
               Est.Edit
               Est!Selecionar = 0
               Est.Update
               X = X + 1
            Est.MoveNext
      Wend
End If
CommitTrans

Finaliza:
  Est.Close
  bd1.Close
  Set Est = Nothing
  Set bd1 = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function


Function RetornaCodigoCategoria(WCodItem As Long) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("EST", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

'BeginTrans
tb.Seek "=", WCodItem
If tb.NoMatch Then
   RetornaCodigoCategoria = 0
   Else
  RetornaCodigoCategoria = tb!CodigoCategoria
End If
'CommitTrans

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function RetornaDescricaoMarcaEstoque(WCodItem As Long) As String
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb, tba As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("EST", DB_OPEN_TABLE)
Set tba = bd.OpenRecordset("Marcas", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tba.index = "PrimaryKey"


'BeginTrans
tb.Seek "=", WCodItem
If Not tb.NoMatch Then
   tba.Seek "=", tb!CodigoMarca
   If Not tba.NoMatch Then
      If IsNull(tba!Descricao) Then
         RetornaDescricaoMarcaEstoque = "."
      Else
         RetornaDescricaoMarcaEstoque = UCase(tba!Descricao)
      End If
   End If
End If
'CommitTrans

Finaliza:
  tb.Close
  tba.Close
  bd.Close
  Set tb = Nothing
  Set tba = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function
Function RetornaEspeciftecEstoque(WCodItem As Long) As String
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("EST", DB_OPEN_TABLE)
tb.index = "PrimaryKey"


'BeginTrans
tb.Seek "=", WCodItem
If Not tb.NoMatch Then
      If IsNull(tb!Aplicacao) Then
         RetornaEspeciftecEstoque = " "
      Else
         RetornaEspeciftecEstoque = UCase(tb!Aplicacao)
      End If
End If
'CommitTrans

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function RetornaDescricaoCAtegoria(WCodigo As Long) As String
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("CATPROD", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

'BeginTrans
tb.Seek "=", WCodigo
If tb.NoMatch Then
   RetornaDescricaoCAtegoria = "."
Else
   If IsNull(tb!DescricaoCategoria) Then
      RetornaDescricaoCAtegoria = "."
   Else
      RetornaDescricaoCAtegoria = UCase(tb!DescricaoCategoria)
   End If
End If
'CommitTrans

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function RetornaLocalizacaoEstoque(WCodItem As Long) As String
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("EMPXEST", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

'BeginTrans
tb.Seek "=", Form_Menu.WEmpresa, WCodItem
If tb.NoMatch Then
   RetornaLocalizacaoEstoque = " "
Else
   If IsNull(tb!Localizacao) Then
      RetornaLocalizacaoEstoque = " "
   Else
      RetornaLocalizacaoEstoque = UCase(tb!Localizacao)
   End If
End If
'CommitTrans

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function SelecionaCategorias() As Long
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim X As Long, Est As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Est = bd1.OpenRecordset("CATPROD", DB_OPEN_TABLE)
Est.index = "PrimaryKey"

BeginTrans
If Est.RecordCount <> 0 Then
      While Not Est.EOF
            
               Est.Edit
               Est!Selecionar = 1
               Est.Update
               X = X + 1
            Est.MoveNext
      Wend
End If
CommitTrans

Finaliza:
  Est.Close
  bd1.Close
  Set Est = Nothing
  Set bd1 = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function RateioEntradas(WFornec As Long, WNotaFiscal As Long, WCFOP As Long, WRateio As Double, WTipo As String) As Long
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim Ent, DEnt As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Ent = bd1.OpenRecordset("ENTRADAS", DB_OPEN_TABLE)
Set DEnt = bd1.OpenRecordset("DETENTRADAS", DB_OPEN_TABLE)
Ent.index = "ChavePrimaria"
DEnt.index = "QuintaChave"
BeginTrans
Ent.Seek "=", WFornec, WNotaFiscal, WCFOP
If Ent.NoMatch Then
   GoSub Finaliza
Else
   While Not Ent.EOF
      If Ent!NotaFiscal = WNotaFiscal And Ent!CodigoFornecedor = WFornec And WCFOP = Ent!cfo Then
         GoSub AplicaRateio:
      Else
         GoSub Finaliza:
      End If
      Ent.MoveNext
   Wend
End If

Finaliza:
  CommitTrans
  Ent.Close
  DEnt.Close
  bd1.Close
  Set Ent = Nothing
  Set DEnt = Nothing
  Set bd1 = Nothing
  Exit Function
Return

AplicaRateio:
   DEnt.Seek "=", WFornec, WNotaFiscal, WCFOP
   If Not DEnt.NoMatch Then
      While Not DEnt.EOF
            If DEnt!NotaFiscal = WNotaFiscal And DEnt!CodigoFornecedor = WFornec And WCFOP = DEnt!CFOP Then
               DEnt.Edit
               If WTipo = "2" Then
                  If DEnt!CodigoCST = 10 Or DEnt!CodigoCST = 60 Or DEnt!CodigoCST = 90 Or DEnt!CodigoCST = 70 Or DEnt!CodigoCST = 30 Or DEnt!CodigoCST = 110 Or DEnt!CodigoCST = 160 Or DEnt!CodigoCST = 170 Or DEnt!CodigoCST = 130 Or DEnt!CodigoCST = 210 Then
                     DEnt!ICMSRETIDO = (DEnt!ValorUnit * WRateio) / 100
                  End If
               ElseIf WTipo = "1" Then
                  DEnt!VlrFrete = (DEnt!ValorUnit * WRateio) / 100
               'ElseIf WTipo = "3" Then
               '   DEnt!Desconto = (DEnt!ValorUnit * WRateio) / 100
               ElseIf WTipo = "4" Then
                  DEnt!OUTRASDESPESAS = (DEnt!ValorUnit * WRateio) / 100
               ElseIf WTipo = "5" Then
                  If DEnt!CodigoCST = 10 Or DEnt!CodigoCST = 60 Or DEnt!CodigoCST = 70 Or DEnt!CodigoCST = 30 Or DEnt!CodigoCST = 110 Or DEnt!CodigoCST = 160 Or DEnt!CodigoCST = 170 Or DEnt!CodigoCST = 130 Or DEnt!CodigoCST = 210 Then
                     DEnt!BaseCalculoICMSST = (DEnt!ValorUnit * WRateio) / 100
                  End If
               End If
               DEnt.Update
            Else
               Return
            End If
            DEnt.MoveNext
      Wend
   
   End If
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function RetornaTipoLaudo(WTipo As String) As String
If WTipo = "A" Then
   RetornaTipoLaudo = "SELAGEM"
ElseIf WTipo = "S" Then
   RetornaTipoLaudo = "SUBSTITUI��O"
ElseIf WTipo = "T" Then
   RetornaTipoLaudo = "SELAGEM BANCADA"
ElseIf WTipo = "E" Then
   RetornaTipoLaudo = "ENSAIO"
ElseIf WTipo = "F" Then
   RetornaTipoLaudo = "SELAGEM E ENSAIO"
ElseIf WTipo = "B" Then
   RetornaTipoLaudo = "SUBSTITUICAO E ENSAIO"
End If
End Function

Function DesmarcaOSV() As Long
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim OSV As Recordset
Dim X As Long
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set OSV = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
OSV.index = "PrimaryKey"

If OSV.RecordCount <> 0 Then
   While Not OSV.EOF
      If OSV!SelecionaOSV Then
         BeginTrans
         OSV.Edit
         OSV!SelecionaOSV = 0
         OSV.Update
         X = X + 1
         If X Mod 500 = 0 Then
               CommitTrans
               BeginTrans
            End If
         CommitTrans
      End If
      OSV.MoveNext
   Wend
End If

Finaliza:

  OSV.Close
  bd1.Close
  Set OSV = Nothing
  Set bd1 = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ValidaSenhaResp(WResp As Long) As String
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("RESP", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

tb.Seek "=", WResp
If tb.NoMatch Then
   MsgBox "Responsavel invalido, impossivel prosseguir...", vbCritical, "SGa"
   ValidaSenhaResp = "USUARIO INEXISTENTE"
Else
   If IsNull(tb!Sen) Then
      ValidaSenhaResp = "Livre"
   Else
      ValidaSenhaResp = tb!Sen
   End If
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function



Function ValidaEst(WCodItem As Long) As Long
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim tb As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd1.OpenRecordset("EST", DB_OPEN_TABLE)
tb.index = "primarykey"

tb.Seek "=", WCodItem
If tb.NoMatch Then
   ValidaEst = 0
Else
   ValidaEst = 1
End If

Finaliza:
  tb.Close
  bd1.Close
  Set tb = Nothing
  Set bd1 = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function



Function TiraPontos(WMAscara As String) As String
Dim X, Comp As Integer
Dim WRet As String
If IsNull(WMAscara) Then
   Exit Function
End If
For X = 1 To Len(WMAscara)
    If Mid$(WMAscara, X, 1) <> " " And Mid$(WMAscara, X, 1) <> "-" And Mid$(WMAscara, X, 1) <> "." And Mid$(WMAscara, X, 1) <> "/" Then
       WRet = WRet & Mid$(WMAscara, X, 1)
    End If
Next X
TiraPontos = WRet

End Function

Function TiraEspacos(WMAscara As String) As String
Dim X, Comp As Integer
Dim WRet As String
If IsNull(WMAscara) Then
   Exit Function
End If
For X = 1 To Len(WMAscara)
    If Mid$(WMAscara, X, 1) <> " " Or Not IsEmpty(Mid$(WMAscara, X, 1)) Or Mid$(WMAscara, X, 1) <> Space(1) Then
       WRet = WRet & Mid$(WMAscara, X, 1)
    End If
Next X
TiraEspacos = WRet

End Function

Function ApuraImpressaoSaldos(WImprimirSaldos As String) As Long
    On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim Reg As Recordset
    Dim WTotal As Double
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
    Set Reg = bd1.OpenRecordset("REGINV", DB_OPEN_TABLE)
     BeginTrans
    If Reg.RecordCount <> 0 Then
       While Not Reg.EOF
          If Reg.EOF Then
             GoTo Finaliza:
          End If
          
          If WImprimirSaldos = "1" Then
             If Reg!Saldo > 0 Then
                Reg.Delete
             End If
          ElseIf WImprimirSaldos = "2" Then
             If Reg!Saldo <= 0 Then
                Reg.Delete
             End If
          End If
          Reg.MoveNext
       Wend
    End If

Finaliza:
    CommitTrans
    Reg.Close
    bd1.Close
    Set Reg = Nothing
    Set bd1 = Nothing
    DoCmd.Hourglass False
        Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ApuraTotalrEGInvScat(WCat As String) As Double
    On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim Reg As Recordset
    Dim WTotal As Double
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
    Set Reg = bd1.OpenRecordset("REGINV", DB_OPEN_TABLE)
    Reg.MoveLast
    Reg.MoveFirst
    While Not Reg.EOF
       If Reg.EOF Then
          GoTo Finaliza:
       End If
       If WCat = Reg!DescricaoCategoria Then
          If Reg!Saldo > 0 Then
             WTotal = WTotal + (Reg!VlrUnitario * Reg!Saldo)
           'Else
           ' ' WTotal = WTotal + (REG!vlrunitario * (REG!Saldos * -1))
           '  WTotal = WTotal + (REG!VlrUnitario * REG!Saldo)
          End If
       End If
       Reg.MoveNext
   Wend

Finaliza:
    Reg.Close
    DoCmd.Hourglass False
    ApuraTotalrEGInvScat = WTotal
    bd1.Close
    Set Reg = Nothing
    Set bd1 = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ApuraTotalrEGInvS() As Double
    On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim Reg As Recordset
    Dim WTotal As Double
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
    Set Reg = bd1.OpenRecordset("REGINV", DB_OPEN_TABLE)
    If Reg.RecordCount <> 0 Then
       Reg.MoveLast
       Reg.MoveFirst
    While Not Reg.EOF
       If Reg.EOF Then
          GoTo Finaliza:
       End If
       
       If Reg!Saldos >= 0 Then
          WTotal = WTotal + (Reg!VlrUnitario * Reg!Saldo)
        Else
         ' WTotal = WTotal + (REG!vlrunitario * (REG!Saldos * -1))
          WTotal = WTotal + (Reg!VlrUnitario * Reg!Saldo)
       End If
       Reg.MoveNext
   Wend
End If
Finaliza:
    Reg.Close
    DoCmd.Hourglass False
    ApuraTotalrEGInvS = WTotal
    bd1.Close
    Set Reg = Nothing
    Set bd1 = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function


Function ApuraTotalrEGInvAcat(WCat As String) As Double
    On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim Reg As Recordset
    Dim WTotal As Double
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
    Set Reg = bd1.OpenRecordset("REGINV", DB_OPEN_TABLE)
    Reg.MoveLast
    Reg.MoveFirst
    While Not Reg.EOF
       If Reg.EOF Then
          GoTo Finaliza:
       End If
       If WCat = Reg!DescricaoCategoria Then
          If Reg!SaldoA >= 0 Then
             WTotal = WTotal + (Reg!VlrUnitario * Reg!Saldo)
           Else
            'WTotal = WTotal + (REG!vlrunitario * (REG!SaldoA * -1))
             WTotal = WTotal + (Reg!VlrUnitario * Reg!Saldo)
          End If
       End If
       Reg.MoveNext
   Wend

Finaliza:
    Reg.Close
    DoCmd.Hourglass False
    ApuraTotalrEGInvAcat = WTotal
    bd1.Close
    Set Reg = Nothing
    Set bd1 = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function


Function ApuraTotalrEGInvA() As Double
    On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim Reg As Recordset
    Dim WTotal As Double
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
    Set Reg = bd1.OpenRecordset("REGINV", DB_OPEN_TABLE)
    'REG.MoveLast
    'REG.MoveFirst
    If Reg.RecordCount <> 0 Then
    While Not Reg.EOF
       If Reg.EOF Then
          GoTo Finaliza:
       End If
       
       If Reg!SaldoA >= 0 Then
          WTotal = WTotal + (Reg!VlrUnitario * Reg!Saldo)
        Else
          'WTotal = WTotal + (REG!vlrunitario * (REG!SaldoA * -1))
          WTotal = WTotal + (Reg!VlrUnitario * Reg!Saldo)
       End If
       Reg.MoveNext
   Wend
   End If

Finaliza:
    Reg.Close
    DoCmd.Hourglass False
    ApuraTotalrEGInvA = WTotal
    bd1.Close
    Set Reg = Nothing
    Set bd1 = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ApuraTotalrEGInvCAT(WCat As String) As Double
    On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim Reg As Recordset
    Dim WTotal As Double
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
    Set Reg = bd1.OpenRecordset("REGINV", DB_OPEN_TABLE)
    Reg.MoveLast
    Reg.MoveFirst
    While Not Reg.EOF
       If Reg.EOF Then
          GoTo Finaliza:
       End If
       If WCat = Reg!DescricaoCategoria Then
          If Reg!Saldo > 0 Then
             WTotal = WTotal + (Reg!VlrUnitario * Reg!Saldo)
          Else
             WTotal = WTotal + (Reg!VlrUnitario * (Reg!Saldo * -1))
          End If
       End If
       Reg.MoveNext
   Wend

Finaliza:
    Reg.Close
    DoCmd.Hourglass False
    ApuraTotalrEGInvCAT = WTotal
    bd1.Close
    Set Reg = Nothing
    Set bd1 = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function TotalRegInvPrecoVenda() As Double
    On Error GoTo TratamentodeErro:
    Dim bd As DataBase
    Dim tb As Recordset
    Dim WTotal As Double
    Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
    Set tb = bd.OpenRecordset("REGINV", DB_OPEN_TABLE)
    
    If tb.RecordCount <> 0 Then
       While Not tb.EOF
             If tb!Saldo > 0 Then
                WTotal = WTotal + (tb!VlrUnitario * tb!Saldo)
 '            ElseIf tb!Saldo < 0 Then
  '                   WTotal = WTotal + (CapturaPrecoVenda(tb!CodItem, Form_Menu.WEmpresa) * (tb!Saldo * -1))
             End If
             tb.MoveNext
             If tb.EOF Then
                GoTo Finaliza:
             End If
        Wend
    End If

Finaliza:
    tb.Close
    DoCmd.Hourglass False
    TotalRegInvPrecoVenda = WTotal
    bd.Close
    Set tb = Nothing
    Set bd = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function TotalRegInvUltCompra() As Double
On Error GoTo TratamentodeErro:
    Dim bd As DataBase
    Dim tb As Recordset
    Dim WTotal As Double
    Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
    Set tb = bd.OpenRecordset("REGINV", DB_OPEN_TABLE)
    
    If tb.RecordCount <> 0 Then
       While Not tb.EOF
             If tb!Saldo > 0 Then
                WTotal = WTotal + (Round(ApuraValorUltimaAquisicao(tb!CodItem)) * tb!Saldo)
             ElseIf tb!Saldo < 0 Then
                     WTotal = WTotal + (Round(ApuraValorUltimaAquisicao(tb!CodItem)) * (tb!Saldo * -1))
             End If
             tb.MoveNext
             If tb.EOF Then
                GoTo Finaliza:
             End If
        Wend
    End If

Finaliza:
    tb.Close
    DoCmd.Hourglass False
    TotalRegInvUltCompra = WTotal
    bd.Close
    Set tb = Nothing
    Set bd = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function TotalRegInvCustoMedio() As Double
    On Error GoTo TratamentodeErro:
    Dim bd As DataBase
    Dim tb As Recordset
    Dim WTotal As Double
    Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
    Set tb = bd.OpenRecordset("REGINV", DB_OPEN_TABLE)
    
    If tb.RecordCount <> 0 Then
       While Not tb.EOF
             If tb!Saldo > 0 Then
                WTotal = WTotal + (Round(tb!CustoMedio) * tb!Saldo)
             'ElseIf tb!Saldo < 0 Then
             '        WTotal = WTotal + (Round(CapturaCustoMedio(tb!CodItem)) * (tb!Saldo * -1))
             End If
             tb.MoveNext
             If tb.EOF Then
                GoTo Finaliza:
             End If
        Wend
    End If

Finaliza:
    tb.Close
    DoCmd.Hourglass False
    TotalRegInvCustoMedio = WTotal
    bd.Close
    Set tb = Nothing
    Set bd = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ApuraTotalrEGInv() As Double
    On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim bd As DataBase
    Dim tbe As Recordset
    Dim tbc As Recordset
    Dim Reg As Recordset
    Dim WTotal As Double
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
    Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tbe = bd.OpenRecordset("EST", DB_OPEN_TABLE)
    Set tbc = bd.OpenRecordset("CATPROD", DB_OPEN_TABLE)
    Set Reg = bd1.OpenRecordset("REGINV", DB_OPEN_TABLE)
    tbc.index = "PrimaryKey"
    tbe.index = "PrimaryKey"
    While Not Reg.EOF
       
       tbc.Seek "=", Reg!CodigoCategoria
       If Not tbc.NoMatch Then
          If tbc!Selecionar Then
             tbe.Seek "=", Reg!CodItem
             If tbe!Selecionar Then
                If Not tbe.NoMatch Then
                   If Reg!Saldo > 0 Then
                      WTotal = WTotal + (Reg!VlrUnitario * Reg!Saldo)
                   ElseIf Reg!Saldo < 0 Then
                          WTotal = WTotal + (Reg!VlrUnitario * (Reg!Saldo * -1))
                   End If
                End If
             End If
          End If
       End If
       Reg.MoveNext
       If Reg.EOF Then
          GoTo Finaliza:
       End If
   Wend

Finaliza:
    Reg.Close
    tbe.Close
    tbc.Close
    DoCmd.Hourglass False
    ApuraTotalrEGInv = WTotal
    bd.Close
    bd1.Close
    Set Reg = Nothing
    Set tbe = Nothing
    Set tbc = Nothing
    Set bd = Nothing
    Set bd1 = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function PagamentosNoDiaGeral(WDataInicial As Variant, WDataFinal As Variant, WBanco As Long) As Double
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim Receb As Recordset
Dim TotalReceb As Double
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Receb = bd1.OpenRecordset("LIQCTPAG", DB_OPEN_TABLE)
Receb.index = "Carteira"
Receb.Seek "=", WBanco
TotalReceb = 0
If Not Receb.NoMatch Then
   Do Until Receb.EOF
      If Receb!Carteira = WBanco Then
            If Receb!DataPagto >= WDataInicial And Receb!DataPagto <= WDataFinal Then
               If Form_Menu.WEmpresa = Receb!CodEmpresa Then
                  TotalReceb = TotalReceb + (Receb!ValorPago + Receb!Juros)
               End If
            End If
      Else
        Exit Do
      End If
      Receb.MoveNext
   Loop
   Receb.Close
End If

Finaliza:
   PagamentosNoDiaGeral = TotalReceb
   Receb.Close
   bd1.Close
   Set Receb = Nothing
   Set bd1 = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function


Function RecebimentosNoDiaGeral(WDataInicial As Variant, WDataFinal As Variant, WBanco As Long) As Double
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim Receb As Recordset
Dim TotalReceb As Double
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Receb = bd1.OpenRecordset("PGCTREC", DB_OPEN_TABLE)
Receb.index = "Carteira"
Receb.Seek "=", WBanco
TotalReceb = 0
If Not Receb.NoMatch Then
   Do Until Receb.EOF
      If Receb!Carteira = WBanco Then
         If Receb!DataPagto >= WDataInicial And Receb!DataPagto <= WDataFinal Then
            If Form_Menu.WEmpresa = Receb!CodEmpresa Then
               TotalReceb = TotalReceb + (Receb!ValorPago + Receb!Juros)
            End If
         End If
      Else
         Exit Do
      End If
      Receb.MoveNext
   Loop
   Receb.Close
End If

Finaliza:
  RecebimentosNoDiaGeral = TotalReceb
  Receb.Close
  bd1.Close
  Set Receb = Nothing
  Set bd1 = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Corrige_Campo_Nulo(WNomeTabela As String, WCampo As String, WValor As Variant) As Long

On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim X As Long, tb, tbc As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset(WNomeTabela, DB_OPEN_TABLE)

BeginTrans
If tb.RecordCount <> 0 Then
      X = 0
      While Not tb.EOF
            If tb.Fields(WCampo).Value = "" Or IsNull(tb.Fields(WCampo).Value) Then
               tb.Edit
               tb.Fields(WCampo).Value = WValor
               tb.Update
             End If
               
             X = X + 1
             If X Mod 200 = 0 Then
                CommitTrans
                BeginTrans
             End If
             tb.MoveNext
      Wend
End If

Finaliza:
  CommitTrans
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return


TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       GoSub Finaliza:
    End If


End Function


Function Captura_Identif_Equip(WNumos As Long) As String
On Error GoTo TratamentodeErro:
Dim WNumSerie As String
Dim bd As DataBase
Dim tb, tba, tbb, tbc As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("TACOGRAFOSOS", DB_OPEN_TABLE)
Set tba = bd.OpenRecordset("DESCONTAM_TANQUE_OSV", DB_OPEN_TABLE)
Set tbb = bd.OpenRecordset("TAXIMETRO_OSV", DB_OPEN_TABLE)
Set tbc = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "SecondKey"
tba.index = "NumOS"
tbb.index = "PrimaryKey"
tbc.index = "PrimaryKey"

tbc.Seek "=", WNumos
If Not tbc.NoMatch Then
   If Not IsNull(tbc!EmitirLaudo) And tbc!EmitirLaudo <> "N" Then
      If tbc!EmitirLaudo = "S" Then
         If Not IsNull(tbc!TipoLaudo) Then
            If tbc!TipoLaudo = "AFT" Then
               GoSub Apura_Taximetro:
            ElseIf tbc!TipoLaudo = "DST" Then
                   GoSub Apura_Tanque:
            Else
                GoSub Apura_Equipamento:
            End If
         End If
      End If
   Else
      GoSub Apura_Equipamento:
   End If
Else
   GoSub Apura_Equipamento:
End If

Finaliza:
   Captura_Identif_Equip = WNumSerie
   tb.Close
   tba.Close
   tbb.Close
   tbc.Close
   bd.Close
   Set tb = Nothing
   Set tba = Nothing
   Set tbb = Nothing
   Set tbc = Nothing
   Set bd = Nothing
   Exit Function
Return
   
Apura_Equipamento:
    tb.Seek "=", WNumos
    If tb.NoMatch Then
       WNumSerie = "#"
    Else
       If IsNull(tb!NumeroTacografo) Then
          WNumSerie = "nulo"
       Else
          WNumSerie = tb!NumeroTacografo
       End If
    End If
Return
   
Apura_Tanque:
    tba.Seek "=", WNumos
    If tba.NoMatch Then
       WNumSerie = "#"
    Else
       If IsNull(tba!Num_Tanque) Then
          WNumSerie = "nulo"
       Else
          WNumSerie = tba!Num_Tanque
       End If
    End If
Return
   
Apura_Taximetro:
    tbb.Seek "=", WNumos
    If tbb.NoMatch Then
       WNumSerie = "#"
    Else
       If IsNull(tbb!Nro_Taximetro) Then
          WNumSerie = "nulo"
       Else
          WNumSerie = tbb!Nro_Taximetro
       End If
    End If
Return
   
   
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function


Function CriaHRPRXHREX() As Long
On Error GoTo TratamentodeErro:
Dim db, dbt As DataBase, Ret As Variant, X As Long, numreg As Long
Dim QGR As QueryDef, GR As Recordset, GRT As Recordset

Set db = CurrentDb()
Set dbt = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set QGR = db.QueryDefs("GraficoHRProgxHrExec")
QGR.Parameters(0) = Forms![GraficoHRProgxHrExec]!Mes
QGR.Parameters(1) = Forms![GraficoHRProgxHrExec]!Ano
QGR.Parameters(2) = Forms![GraficoHRProgxHrExec]!Categoria
Set GR = QGR.OpenRecordset()
Set GRT = dbt.OpenRecordset("HRPRXHREX", DB_OPEN_TABLE)
GRT.index = "PrimaryKey"
'Ent.Seek "=", CodItem, DataEnt

If GRT.RecordCount <> 0 Then
   BeginTrans
   Do Until GRT.EOF
      GRT.Delete
      GRT.MoveNext
  Loop
  CommitTrans
End If

If GR.RecordCount <> 0 Then
   BeginTrans
   Do Until GR.EOF
      GRT.Seek "=", GR!Mes, GR!Ano, GR!TipoServico
      If GRT.NoMatch Then
         GRT.AddNew
       Else
         GRT.Edit
      End If
      GRT!Mes = GR!Mes
      GRT!Ano = GR!Ano
      GRT!TipoServico = GR!TipoServico
      GRT!TP = GRT!TP + GR!TempoProgramado
      GRT!TE = GRT!TE + GR!TempoExecucao
      GRT.Update
      GR.MoveNext
   Loop
   CommitTrans
End If

Finaliza:
 GRT.Close
 GR.Close
 db.Close
 dbt.Close
 Set GRT = Nothing
 Set GR = Nothing
 Set db = Nothing
 Set dbt = Nothing
 Exit Function
Return

TratamentodeErro:
   If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function



Function ConverteHorasSegundos(WHora As Variant, WMin As Variant) As Variant
On Error GoTo TratamentodeErro:
Dim Hor As Double
Dim Min As Double



If Not IsNull(WHora) And IsNumeric(WHora) Then
   
   Hor = (WHora * 3600)
Else
   Hor = 0
End If

If Not IsNull(WMin) And IsNumeric(WMin) Then
   Min = (WMin * 60)
Else
   Min = 0
End If

ConverteHorasSegundos = Hor + Min

Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
Return

End Function

Function ConverteHorasSegundos_Selagens(WHora As Variant, WMin As Variant) As Variant
On Error GoTo TratamentodeErro:
Dim Hor As Double
Dim Min As Double
Dim Seg

Seg = Format$(Time(), "ss")


If Not IsNull(WHora) Then
   Hor = (WHora * 3600)
Else
   Hor = 0
End If

If Not IsNull(WMin) Then
   Min = (WMin * 60)
Else
   Min = 0
End If

ConverteHorasSegundos_Selagens = Hor + Min + Seg

Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
Return

End Function

Function ConverteSegundosHoras(WSeg As Long) As String
On Error GoTo TratamentodeErro:
Dim WH As Long, WM As Long, ws As Long
Dim h As Long
Dim m As Long
Dim s As Long

ws = 0

h = (WSeg / 3600)
m = (WSeg - (3600 * h)) / 60
s = 0

WH = Int(WSeg / 3600)
WM = Int((WSeg - (3600 * WH)) / 60)

ConverteSegundosHoras = Format$(WH, "00") & ":" & Format$(WM, "00")

Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
Return

End Function

Function ConverteSegundosMinutos(WSeg As Double) As String
On Error GoTo TratamentodeErro:
Dim WH As Double, WM As Double, ws As Double
Dim h As Double
Dim m As Double
Dim s As Double

ws = 0

h = (WSeg / 3600)
m = (WSeg - (3600 * h))
s = 0

WH = Int(WSeg / 3600)
WM = Int((WSeg - (3600 * WH)) / 60)

ConverteSegundosMinutos = h + m

Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
Return

End Function

Function ConverteSegundosFracao(WSeg As Double) As String
On Error GoTo TratamentodeErro:
Dim WH As Double, WM As Double, ws As Double
Dim db As DataBase, fr As Recordset
Set db = CurrentDb()
Set fr = db.OpenRecordset("FracaoHoras", DB_OPEN_TABLE)
fr.index = "SecondKey"

ws = 0
'WH = Int(WSeg / 3600)
'WM = Int((WSeg - (3600 * WH)) / 60)
WH = Int(WSeg / 3600)
WM = Round((WSeg / 3600) - WH)

fr.Seek "=", WM
If Not fr.NoMatch Then
   ConverteSegundosFracao = Format$(WH, "00") & "," & Format$(fr!Min, "00")
Else
  ConverteSegundosFracao = Format$(WH, "00") & "," & "00"
End If
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
Return
End Function
Function GeraTabelaProdutividadeHoras(WQF As Integer, WQD As Integer, WHP As Double, WDataRef As Variant)
On Error GoTo TratamentodeErro:
Dim WHD As Double, WTotH As Double, WTotA As Double, WTotb As Double, WTotC As Double, WH As String
Dim db, dbt As DataBase, QHE As QueryDef, HE As Recordset, HP As Recordset
Set db = CurrentDb()
Set dbt = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set QHE = db.QueryDefs("HorasExecutadasPeriodo")
'QHE.Parameters(0) = Forms!GraficoProdutividadeServicos!DataInicial
'QHE.Parameters(1) = Forms!GraficoProdutividadeServicos!DataFinal
Set HE = QHE.OpenRecordset()
Set HP = dbt.OpenRecordset("HorasProdutividade", DB_OPEN_TABLE)
HP.index = "PrimaryKey"

If HP.RecordCount <> 0 Then
   BeginTrans
   Do Until HP.EOF
      HP.Delete
      HP.MoveNext
   Loop
   CommitTrans
End If

WTotA = WHP * WQD
WH = ConverteSegundosHoras(CLng(WTotA))

If HE.RecordCount <> 0 Then
   BeginTrans
   Do Until HE.EOF
      If Not IsNull(HE!datafechamento) Then
         If Format$(HE!datafechamento, "mm/yyyy") = Format$(WDataRef, "mm/yyyy") Then
            GoSub GeraRealizado:
         End If
         If Format$(HE!DataAbertura, "mm/yyyy") <> Format$(WDataRef, "mm/yyyy") Then
            If Format$(HE!datafechamento, "mm/yyyy") = Format$(WDataRef, "mm/yyyy") Then
               GoSub GeraExecutadasMesesAnt:
            End If
         Else
            If Format$(HE!datafechamento, "mm/yyyy") = Format$(WDataRef, "mm/yyyy") Then
                GoSub GeraExecutadasMesAtual:
            End If
         End If
      End If
      
      If Format$(HE!DataAbertura, "mm/yyyy") = Format$(WDataRef, "mm/yyyy") Then
         If IsNull(HE!datafechamento) Or Format$(HE!datafechamento, "mm/yyyy") <> Format$(WDataRef, "mm/yyyy") Then
            GoSub GeraPendentes:
         End If
      End If
      
      If Format$(HE!DataAbertura, "mm/yyyy") = Format$(WDataRef, "mm/yyyy") Then
            GoSub GeraAbertasMes:
      End If
      
      If Format$(HE!PrevisaoFechamento, "mm/yyyy") = Format$(WDataRef, "mm/yyyy") Then
         GoSub GeraProgramadas:
      End If
      
      HE.MoveNext
   Loop
   
   GoSub GeraHP:
   GoSub ApuraHD:
   GoSub GeraHE:
   CommitTrans
End If

Finaliza:
 HP.Close
 HE.Close
 db.Close
 dbt.Close
 Set HP = Nothing
 Set HE = Nothing
 Set db = Nothing
 Set dbt = Nothing
 Exit Function
Return

GeraHP:
 HP.Seek "=", 1, "HORAS DISPONIVEIS"
 If HP.NoMatch Then
    HP.AddNew
 Else
    HP.Edit
 End If
 HP!Sequencia = 1
 HP!Servico = "HORAS DISPONIVEIS"
 HP!QtdeSegundos = WHP
 HP!qtdefracao = ConverteSegundosFracao(WHP)
 HP!qtdehoras = ConverteSegundosHoras(CLng(WHP))
 HP.Update
Return

ApuraHD:
  HP.Seek "=", 2, "EXECUCOES NO MES"
  WHD = HP!QtdeSegundos
  Return

GeraRealizado:
   HP.Seek "=", 2, "EXECUCOES NO MES"
   If HP.NoMatch Then
      HP.AddNew
   Else
      HP.Edit
   End If
   HP!Sequencia = 2
   HP!Servico = "EXECUCOES NO MES"
   HP!QtdeSegundos = HP!QtdeSegundos + HE!TempoExecucao
   HP!qtdefracao = ConverteSegundosFracao(HP!QtdeSegundos)
   HP!qtdehoras = ConverteSegundosHoras(HP!QtdeSegundos)
   HP.Update
Return

GeraHE:
  If WHD - WHP > 0 Then
     HP.Seek "=", 3, "HORAS EXCEDENTES"
     If HP.NoMatch Then
        HP.AddNew
      Else
        HP.Edit
     End If
     HP!Sequencia = 3
     HP!Servico = "HORAS EXCEDENTES"
     HP!QtdeSegundos = WHD - WHP
     HP!qtdefracao = ConverteSegundosFracao(WHD - WHP)
     HP!qtdehoras = ConverteSegundosHoras(HP!QtdeSegundos)
     HP.Update
  End If
Return


GeraPendentes:
 HP.Seek "=", 5, "HORAS PENDENTES"
   If HP.NoMatch Then
      HP.AddNew
   Else
      HP.Edit
   End If
   HP!Sequencia = 5
   HP!Servico = "HORAS PENDENTES"
   HP!QtdeSegundos = HP!QtdeSegundos + HE!TempoExecucao
   HP!qtdefracao = ConverteSegundosFracao(HP!QtdeSegundos)
   HP!qtdehoras = ConverteSegundosHoras(HP!QtdeSegundos)
   HP.Update
Return

GeraProgramadas:
   HP.Seek "=", 4, "HORAS PROGRAMADAS"
   If HP.NoMatch Then
      HP.AddNew
   Else
      HP.Edit
   End If
   HP!Sequencia = 4
   HP!Servico = "HORAS PROGRAMADAS"
   HP!QtdeSegundos = HP!QtdeSegundos + HE!TempoProgramado
   HP!qtdefracao = ConverteSegundosFracao(HP!QtdeSegundos)
   HP!qtdehoras = ConverteSegundosHoras(HP!QtdeSegundos)
   HP.Update
Return

GeraExecutadasMesesAnt:
   HP.Seek "=", 5, "EXECUCOES MESES ANT"
   If HP.NoMatch Then
      HP.AddNew
   Else
      HP.Edit
   End If
   HP!Sequencia = 5
   HP!Servico = "EXECUCOES MESES ANT"
   HP!QtdeSegundos = HP!QtdeSegundos + HE!TempoExecucao
   HP!qtdefracao = ConverteSegundosFracao(HP!QtdeSegundos)
   HP!qtdehoras = ConverteSegundosHoras(HP!QtdeSegundos)
   HP.Update

Return

GeraExecutadasMesAtual:
   HP.Seek "=", 6, "EXECUCOES MES ATUAL"
   If HP.NoMatch Then
      HP.AddNew
   Else
      HP.Edit
   End If
   HP!Sequencia = 6
   HP!Servico = "EXECUCOES MES ATUAL"
   HP!QtdeSegundos = HP!QtdeSegundos + HE!TempoExecucao
   HP!qtdefracao = ConverteSegundosFracao(HP!QtdeSegundos)
   HP!qtdehoras = ConverteSegundosHoras(HP!QtdeSegundos)
   HP.Update
Return

GeraAbertasMes:

   HP.Seek "=", 7, "ABERTAS NO MES"
   If HP.NoMatch Then
      HP.AddNew
   Else
      HP.Edit
   End If
   HP!Sequencia = 7
   HP!Servico = "ABERTAS NO MES"
   HP!QtdeSegundos = HP!QtdeSegundos + HE!TempoExecucao
   HP!qtdefracao = ConverteSegundosFracao(HP!QtdeSegundos)
   HP!qtdehoras = ConverteSegundosHoras(HP!QtdeSegundos)
   HP.Update

Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
Return
End Function

Function InicializaDEMSERV()
On Error GoTo TratamentodeErro:
Dim db As DataBase, DS As Recordset
Set db = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set DS = db.OpenRecordset("DEMSERV", DB_OPEN_TABLE)
DS.index = "PrimaryKey"

If DS.RecordCount <> 0 Then
   BeginTrans
   Do Until DS.EOF
      DS.Delete
      DS.MoveNext
   Loop
   CommitTrans
End If

Finaliza:
 DS.Close
 db.Close
 Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
Return
End Function

Function GeraTabelaDemonstServicos()
On Error GoTo TratamentodeErro:
Dim db, dbt As DataBase, QDA As QueryDef, DS As Recordset, Da As Recordset
Set db = CurrentDb()
Set dbt = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set QDA = db.QueryDefs("GraficoDemonstrativoServicos")
QDA.Parameters(0) = Forms!GraficoDemonstServicos!DataInicial
QDA.Parameters(1) = Forms!GraficoDemonstServicos!DataFinal
Set Da = QDA.OpenRecordset()
Set DS = dbt.OpenRecordset("DEMSERV", DB_OPEN_TABLE)
DS.index = "PrimaryKey"

If DS.RecordCount <> 0 Then
   BeginTrans
   Do Until DS.EOF
      DS.Delete
      DS.MoveNext
   Loop
   CommitTrans
End If

If Da.RecordCount <> 0 Then
   BeginTrans
   Do Until Da.EOF
      DS.Seek "=", Da!CodigoServico, Da!datafechamento
      If DS.NoMatch Then
         DS.AddNew
      Else
         DS.Edit
      End If
      DS!DataServico = Da!datafechamento
      DS!CodigoServico = Da!CodigoServico
      If IsNull(Da!Descricao) Then
         DS!Descricao = "DESCRICAO NULA"
      Else
         DS!Descricao = Da!Descricao
      End If
      DS!TempoExecucao = DS!TempoExecucao + Da!Execucao
      DS.Update
      Da.MoveNext
   Loop
   CommitTrans
End If

Finaliza:
 DS.Close
 Da.Close
 db.Close
 dbt.Close
 Set DS = Nothing
 Set Da = Nothing
 Set db = Nothing
 Set dbt = Nothing
 Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
Return
End Function

Function DateCheck(ctlDate As TextBox)
    Dim datepassed As Variant
    datepassed = ctlDate
    ctlDate = GetDate(datepassed)
End Function


Function ServicosInvalidos(WNumos As Long) As Long
On Error GoTo TratamentodeErro:
DoCmd.Hourglass True
Dim db As DataBase, TPServ, serv As Recordset
Set db = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set serv = db.OpenRecordset("ITENSSERVOSV", DB_OPEN_TABLE)
Set TPServ = db.OpenRecordset("TPSERV", DB_OPEN_TABLE)
serv.index = "SecondaryKey"
TPServ.index = "PrimaryKey"

BeginTrans
serv.Seek "=", WNumos
If serv.NoMatch Then
   GoSub Finaliza:
Else
   While Not serv.EOF
         If serv!Numos <> WNumos Then
            GoSub Finaliza:
         End If
         If IsNull(serv!TipoServico) Then
            serv.Delete
         Else
            TPServ.Seek "=", serv!TipoServico
            If TPServ.NoMatch Then
               serv.Delete
            End If
         End If
       
         serv.MoveNext
         If serv.EOF Then
            GoSub Finaliza:
         End If
   Wend
End If

Finaliza:
  DoCmd.Hourglass False
  CommitTrans
  serv.Close
  TPServ.Close
  db.Close
  Set serv = Nothing
  Set TPServ = Nothing
  Set db = Nothing
  Exit Function
Return
 
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function



Function AtualizaPrecoVenda(WCodItem As Long, WCustoAquisicao As Double)
On Error GoTo TratamentodeErro:
DoCmd.Hourglass True
Dim db As DataBase, Prec As Recordset, Cat As Recordset, Est As Recordset, Impostos As Double, Ref As Double, Result As Double, PV As Double
Set db = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Prec = db.OpenRecordset("EMPXEST", DB_OPEN_TABLE)
Set Cat = db.OpenRecordset("CATPROD", DB_OPEN_TABLE)
Set Est = db.OpenRecordset("EST", DB_OPEN_TABLE)
Est.index = "PrimaryKey"
Prec.index = "SecondaryKey"
Cat.index = "PrimaryKey"

Ref = 1#

BeginTrans
Est.Seek "=", WCodItem
If Est.NoMatch Then
   GoSub Finaliza:
Else
   GoSub VerificaCategoria:
   GoSub VerificaEMPXEST:
   Result = Ref - Impostos
   PV = Round(WCustoAquisicao / Result)
   Prec.Edit
   Prec!PrecoVenda = PV
   Prec.Update
End If

Finaliza:
  DoCmd.Hourglass False
  CommitTrans
  Prec.Close
  Est.Close
  Cat.Close
  db.Close
  Set Prec = Nothing
  Set Est = Nothing
  Set Cat = Nothing
  Set db = Nothing
  Exit Function
Return
 
VerificaCategoria:
  Cat.Seek "=", Est!CodigoCategoria
  If Cat.NoMatch Then
     GoSub Finaliza:
  End If
  Impostos = Cat!Comissao + Cat!Margem + Cat!OutDesp + Cat!Frete + Cat!Quebras + Cat!IPI + Cat!ISS + Cat!IR + Cat!PIS + Cat!COFIS + Cat!ICMS + Cat!SIMPLES + Cat!BDI + Cat!CFINANC + Cat!CSSL
Return

VerificaEMPXEST:
 Prec.Seek "=", WCodItem
 If Prec.NoMatch Then
    GoSub Finaliza:
 End If
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function MarcarTodos()
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase, bd As DataBase, QAju As QueryDef, Aju As Recordset, Prec As Recordset, WPrecovenda As Double
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set bd = CurrentDb()
Set Prec = bd1.OpenRecordset("EMPXEST", DB_OPEN_TABLE)
Set QAju = bd.QueryDefs("MarcarTodosAtualizarPrecoVenda")
QAju.Parameters(0) = Forms!AtualizaPrecoVendaXMinimoPrinc!CodigoCategoria
Set Aju = QAju.OpenRecordset()
Prec.index = "Primarykey"

If Aju.RecordCount <> 0 Then
   BeginTrans
   Do Until Aju.EOF
      Aju.Edit
      Aju!AtualizarPreco = 1
      Aju.Update
      Aju.MoveNext
   Loop
   CommitTrans
End If
Prec.Close
bd.Close
Set Prec = Nothing
Set bd = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function TABXEST()
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim Cat As Recordset, tb As Recordset, TBXEST As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Cat = bd1.OpenRecordset("CATPROD", DB_OPEN_TABLE)
Set TBXEST = bd1.OpenRecordset("TABXEST", DB_OPEN_TABLE)
Set tb = bd1.OpenRecordset("TAB", DB_OPEN_TABLE)
Cat.index = "Primarykey"
TBXEST.index = "PrimaryKey"
tb.index = "PrimaryKey"

If Cat.RecordCount <> 0 Then
   BeginTrans
   Do Until Cat.EOF
         If Not IsNull(Cat!DescricaoCategoria) Then
            TBXEST.Seek "=", 1, Cat!CodigoCategoria
            If TBXEST.NoMatch Then
               TBXEST.AddNew
            Else
               TBXEST.Edit
            End If
            TBXEST!tABELA = 1
            TBXEST!CodCategoria = Cat!CodigoCategoria
            TBXEST!Obs = "."
            TBXEST.Update
         End If
         
      Cat.MoveNext
   Loop
   CommitTrans
End If
TBXEST.Close
Cat.Close
tb.Close
bd1.Close
Set Cat = Nothing
Set tb = Nothing
Set bd1 = Nothing

Exit Function

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function RetornaVlrUnitNFS(WNroSeq As Long) As Double
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("ITENSESTNOTAFISCAL", DB_OPEN_TABLE)
tb.index = "TerceiraChave"

tb.Seek "=", WNroSeq

If tb.NoMatch Then
   RetornaVlrUnitNFS = 0
   GoSub Finaliza:
Else
   GoSub ValidaCamposNulos:
   RetornaVlrUnitNFS = tb!ValorUnitario + tb!IPI - tb!Desconto + tb!ICMSST + tb!OUTRASDESP
End If


Finaliza:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing

   Exit Function
Return

ValidaCamposNulos:
   BeginTrans
   tb.Edit
   If IsNull(tb!ValorUnitario) Then
      tb!ValorUnitario = 0
   End If
   If IsNull(tb!Desconto) Then
      tb!Desconto = 0
   End If
   If IsNull(tb!IPI) Then
      tb!IPI = 0
   End If
   If IsNull(tb!ICMSST) Then
      tb!ICMSST = 0
   End If
   If IsNull(tb!BASECALCULOST) Then
      tb!BASECALCULOST = 0
   End If
   If IsNull(tb!OUTRASDESP) Then
      tb!OUTRASDESP = 0
   End If
   If IsNull(tb!AliqAprov) Then
      tb!AliqAprov = 0
   End If
   tb.Update
   CommitTrans
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Desmarcar()
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim X As Long

Dim Est As Recordset, tb As Recordset, TBXEST As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Est = bd1.OpenRecordset("EST", DB_OPEN_TABLE)
Est.index = "Primarykey"

If Est.RecordCount <> 0 Then
   BeginTrans
   Do Until Est.EOF
      Est.Edit
      Est!AtualizarPreco = 0
      Est.Update
      X = X + 1
      If X Mod 200 = 0 Then
            CommitTrans
            BeginTrans
       End If
      
      Est.MoveNext
   Loop
   CommitTrans
End If
Est.Close
bd1.Close
Set Est = Nothing
Set bd1 = Nothing
Exit Function

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ImprimeOSVMod92003(WPorta As String, WPedido As Long) As Long
On Error GoTo TratamentodeErro

DoCmd.Hourglass True
Dim db As DataBase, db1 As DataBase, WValorFaturar As Double, GerarReceber As Double
Dim PED As Recordset, GAR As Recordset, Cli As Recordset, Prod As Recordset, PRZ As Recordset, Ite As Recordset, Resp As Recordset, iped As Recordset, RQM As Recordset, ISERV As Recordset, TPServ As Recordset
Dim wMensagem As String * 70
' CAMPOS EMPRESA
Dim WCont As Integer
Dim WBrancos As String * 77
Dim WTrac1  As String * 77
Dim WTrac2 As String * 77
Dim WTrac As String * 74
Dim WCol As String * 1
Dim WEmpresa As String * 48
Dim WEndEmpresa As String * 41
Dim wTelefone As String * 15
Dim WEmail As String * 20
Dim wcIDADEeMP As String * 17
Dim WGarantia As String * 10
Dim WKMEnt As Double
Dim WKMSai As Double

' CAMPOS CLIENTES
Dim wobs, WNomeCli  As String * 50
Dim WCidade   As String * 20
Dim WEndCli As String * 45
Dim WFANTASIA As String * 30
Dim WUF As String * 2
Dim WCNPJ  As String * 14
Dim wInsc As String * 13
Dim WTelcli As String * 12
Dim wdESCpRZ As String * 50
Dim WTecnicos As String * 50
Dim WVend As String * 16
Dim WDescProd As String * 34
Dim wDescServ As String * 34
Dim WMarca As String * 11
Dim wPlaca As String * 8
Dim WModelo As String * 11
Dim WFrota As String * 11
Dim wcep As String * 8
Dim WBairro As String * 20
Dim wnro As String * 5

'CAMPOS NUMERICOS
Dim WValorUnit As Double
Dim WVALORUNITSTR As String * 8
Dim WValorCarreto As Double
Dim WValorCarretosStr As String * 8
Dim WValorDesconto As Double
Dim WValorDescontoStr As String * 8
Dim WValorJuros As Double
Dim WValorJurosStr As String * 8
Dim WValorIcmRepasse As Double
Dim WValorIcmsRepasseStr As String * 8

Dim wvALORUnitTotalStr As String * 10
Dim WValorUnitTotal As Double
Dim WTotalMerc As Double
Dim WTotalCarreto As Double
Dim WTotalDesconto As Double
Dim WTotalGeral As Double
Dim WTotalMercStr As String * 10
Dim WTotalCarretoStr As String * 10
Dim WTOTALDESCONTOSTR As String * 10
Dim WTotalGeralStr As String * 10
Dim WQuant As Double
Dim WQuantstr As String * 7
Dim WCodFAB As String * 9
Dim WTitulo As String * 10

Dim WQuantS As Double
Dim WValorUnitS As Double
Dim WValorDescontoS As Double
Dim WValorTotalunitS As Double
Dim WTotalS As Double
Dim WQuantStrS As String * 7
Dim WValorUnitStrS As String * 8
Dim WValorDescontostrs As String * 8
Dim WvalorTotalUnitStrS As String * 10
Dim WTotalStrS As String * 10
Dim WtotalDescontoS As Double
Dim WTotalDescontoStrS As String * 10

Set db = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set db1 = CurrentDb()
Set Cli = db.OpenRecordset("CLI", DB_OPEN_TABLE)
Set Prod = db.OpenRecordset("EST", DB_OPEN_TABLE)
Set PRZ = db.OpenRecordset("CONDPAGTO", DB_OPEN_TABLE)
Set Resp = db.OpenRecordset("RESP", DB_OPEN_TABLE)
Set PED = db.OpenRecordset("OS", DB_OPEN_TABLE)
Set RQM = db.OpenRecordset("RQM", DB_OPEN_TABLE)
Set iped = db.OpenRecordset("RMXEMPXEST", DB_OPEN_TABLE)
Set ISERV = db.OpenRecordset("ITENSSERVOSV", DB_OPEN_TABLE)
Set TPServ = db.OpenRecordset("TPSERV", DB_OPEN_TABLE)
Set GAR = db.OpenRecordset("GARANTIA", DB_OPEN_TABLE)

GAR.index = "PrimaryKey"
Cli.index = "primaryKey"
Prod.index = "PrimaryKey"
PRZ.index = "PrimaryKey"
Resp.index = "PrimaryKey"
iped.index = "RQM"
ISERV.index = "sECONDARYKey"
RQM.index = "sECONDARYKey"
PED.index = "PrimaryKey"
TPServ.index = "PrimaryKey"

Open WPorta For Output As #1
'Open "c:\sga\pedido_mult.txt" For Output As #1

WTrac1 = "---------------------------------------------------------------------------------"
WTrac2 = "================================================================================="
WTrac = "=============================================================================="
If Month(Now()) = 12 Then
   wMensagem = "            <<< BOAS FESTAS E FELIZ ANO NOVO >>> "
Else
   If Mensagem() = "1" Then
      wMensagem = "            OBRIGADO PELA PREFERENCIA, VOLTE SEMPRE"
    Else
       wMensagem = UCase(Mensagem())
   End If
End If

             '123456789 123456789 123456789 123456789 123456789 123456789 123456789
WEndEmpresa = Endereco() & Space(1) & NroEnd()
WEmpresa = Mid$(NomeFantasia(), 1, 41)
wTelefone = Telefone()
WEmail = Mid$(Email(), 1, 20)
wcIDADEeMP = Mid$(Cidade(), 1, 17)
PED.Seek "=", WPedido
If PED.NoMatch Then
   MsgBox "Sr." & Captura_Conteudo_Campo_RESP(Usuario_Ativo(), "nome") & "  n�o existe OSV para os parametros informados...", 64, "SGa"
   GoSub Finaliza:
End If
'BeginTrans
WCont = 0
While Not PED.EOF
   If PED!Numos = WPedido Then
      GoSub ImprimeCabecalhoEmpresa:
      GoSub VerificaCliente:
      GoSub VerificaGarantia:
      GoSub VerificaPRZ:
      GoSub VerificaVENd:
      GoSub ImprimeCabecalhoCliente:
      GoSub ImprimeServicos:
      GoSub ApuraItensPedidos:
      WCont = WCont + 1
   End If
   PED.MoveNext
Wend
GoSub ImprimeFechamento:
GoSub InicializaCampos:
'CommitTrans

Finaliza:
    Cli.Close
    Prod.Close
    PRZ.Close
    PED.Close
    iped.Close
    ISERV.Close
    Resp.Close
    TPServ.Close
    RQM.Close
    GAR.Close
    Close #1
    db.Close
    db1.Close
    Set Cli = Nothing
    Set Prod = Nothing
    Set PRZ = Nothing
    Set PED = Nothing
    Set iped = Nothing
    Set ISERV = Nothing
    Set Resp = Nothing
    Set TPServ = Nothing
    Set RQM = Nothing
    Set GAR = Nothing
    Close #1
    Set db = Nothing
    Set db1 = Nothing
    DoCmd.Hourglass False
    'MsgBox "Fim normal de execu��o", 64, "SGa"
    Exit Function
Return

ImprimeServicos:
   ISERV.Seek "=", PED!Numos
   If Not ISERV.NoMatch Then
      Do Until ISERV.EOF
         If ISERV!Numos <> WPedido Then
            Exit Do
         End If
         TPServ.Seek "=", ISERV!TipoServico
         If TPServ.NoMatch Then
            MsgBox "Servico inexistente, OSV n�o pode ser impressa...", 64, "SGa"
            GoSub Finaliza:
        End If
        WValorUnitS = ISERV!ValorMaoObra
        WValorDescontoS = ISERV!Desconto
        WValorTotalunitS = ISERV!Quantidade * (ISERV!ValorMaoObra - ISERV!Desconto)
        WValorDescontostrs = PreencheEspacos(Format(WValorDescontoS, "00000.00"), 8)
        WValorUnitStrS = PreencheEspacos(Format(WValorUnitS, "00000.00"), 8)
        WvalorTotalUnitStrS = PreencheEspacos(Format(WValorTotalunitS, "0000000.00"), 10)
        WQuantStrS = PreencheEspacos(Format(ISERV!Quantidade, "0000.00"), 7)
        WTotalS = WTotalS + WValorTotalunitS
        
        If IsNull(TPServ!Descricao) Then
           wDescServ = "..."
        Else
           wDescServ = Mid$(TPServ!Descricao, 1, 34)
        End If
        WCont = WCont + 1
        If WCont = 25 Then
             GoSub ImprimeFechamentoParcial:
             GoSub ImprimeCabecalhoEmpresa:
             GoSub ImprimeCabecalhoCliente:
             WCont = 15
        End If
        Print #1, WQuantStrS & " " & "UN" & " " & Format$(UCase(Mid$(wDescServ, 1, 34)), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") & Space(1) & WValorUnitStrS & Space(3) & WValorDescontostrs & Space(1) & WvalorTotalUnitStrS
        ISERV.MoveNext
      Loop
   End If
Return

ApuraItensPedidos:
 RQM.Seek "=", WPedido
 If RQM.NoMatch Then
'    MsgBox "n�o existe mercadorias a serem impressas...", 64, "SGa"
    GoSub ImprimeFechamento:
    GoSub Finaliza:
 End If
 Do Until RQM.EOF
    If RQM!Numos <> WPedido Then
       Exit Do
    End If
    iped.Seek "=", RQM!NumRequisicao
    If Not iped.NoMatch Then
       Do Until iped.EOF
          If (iped!Quantidade - iped!Devolucao) <> 0 Then
             If iped!NumRequisicao <> RQM!NumRequisicao Then
                Exit Do
             End If
             WCont = WCont + 1
          If WCont = 25 Then
             GoSub ImprimeFechamentoParcial:
             GoSub ImprimeCabecalhoEmpresa:
             GoSub ImprimeCabecalhoCliente:
             WCont = 15
          End If
          GoSub ImprimeItem:
          End If
          iped.MoveNext
       Loop
    End If
    RQM.MoveNext
 Loop
Return

ImprimeItem:
   Prod.Seek "=", iped!CodItem
   If Not Prod.NoMatch Then
      WValorUnit = iped!ValorUnit + iped!Valor_IPI
      WValorDesconto = iped!Desconto
      WValorDescontoStr = PreencheEspacos(Format(WValorDesconto, "00000.00"), 8)
      WVALORUNITSTR = PreencheEspacos(Format(WValorUnit, "00000.00"), 8)
      WValorUnitTotal = (iped!Quantidade - iped!Devolucao) * (iped!ValorUnit + iped!Valor_IPI - iped!Desconto)
      wvALORUnitTotalStr = PreencheEspacos(Format(WValorUnitTotal, "0000000.00"), 10)
      WQuantstr = PreencheEspacos(Format((iped!Quantidade - iped!Devolucao), "0000.00"), 7)
      WTotalMerc = WTotalMerc + WValorUnitTotal
      If IsNull(Prod!CodigoFabricante) Then
        WCodFAB = Space(9)
      Else
        If Nroserie() = 52007 Then
           If IsNull(Prod!CodigoFabricante) Then
              WCodFAB = Space(9)
           Else
              WCodFAB = Mid$(Prod!CodigoInterno, 1, 9)
           End If
        ElseIf Nroserie() = 100830 Then
              If IsNull(Prod!CodigoInterno) Then
                 WCodFAB = Mid$(Prod!CodigoFabricante, 1, 9)
              Else
                 WCodFAB = Mid$(Prod!CodigoInterno, 1, 9)
              End If
        
        Else
           If IsNull(Prod!CodigoFabricante) Then
              WCodFAB = Space(9)
           Else
              WCodFAB = Mid$(Prod!CodigoFabricante, 1, 9)
           End If
        End If
      End If
      If Nroserie() = 82003 Then
         WDescProd = Mid$(UCase(Prod!Descricao), 1, 35)
      Else
         WDescProd = LTrim$(Mid$(UCase(WCodFAB), 1, 9) & "-" & Mid$(UCase(Prod!Descricao), 1, 24))
      End If
   End If
   Print #1, WQuantstr & " " & Format$(Prod!Unidade) & " " & Format$(UCase(Mid$(WDescProd, 1, 34)), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") & Space(1) & WVALORUNITSTR & Space(3) & WValorDescontoStr & Space(1) & wvALORUnitTotalStr
Return

ImprimeFechamento:
  WTotalMercStr = PreencheEspacos(Format(WTotalMerc, "0000000.00"), 10)
  WTotalGeral = WTotalMerc - WTotalDesconto + WTotalS
  WValorCarretosStr = PreencheEspacos(Format(WTotalCarreto, "00000.00"), 8)
  WTotalGeralStr = PreencheEspacos(Format(WTotalGeral, "0000000.00"), 10)
  
  Do While WCont <= 25
     Print #1, WBrancos
     WCont = WCont + 1
  Loop
  
  Print #1, WTrac1
  
  Print #1, "Recebi(emos),as mercadorias e/ou servicos OSV:" & Format$(WPedido, "000,000") & " TOTAL GERAL:" & WTotalGeralStr
  If CapturaChequesPre(WPedido) = "." Then
    Print #1, WBrancos
  Else
    Print #1, "CHP:" & Mid$(CapturaChequesPre(WPedido), 1, 71)
  End If
  
  Print #1, wcIDADEeMP & "_____/_____/_____" & "Ass.Cliente: ____________________________"
             '123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789
  Print #1, WTrac1
  Print #1, "--->" & Space(1) & Format$(Mid$(wMensagem, 1, 66), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") & " <---"
 ' Print #1, wtRAC2
  Print #1, WBrancos
  Print #1, "                                                                                "
  Print #1, WBrancos
  WCont = 0
Return

ImprimeFechamentoParcial:
  Do While WCont <= 25
     Print #1, WBrancos
     WCont = WCont + 1
  Loop
  
  Print #1, WTrac1
  Print #1, Space(6) & "          " & "           " & Space(15) & "                        " & "Continua..."
  Print #1, WBrancos
  
  Print #1, wcIDADEeMP & "_____/_____/_____" & "Ass.Cliente: ____________________________"
             '123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789
  Print #1, WTrac1
  Print #1, "--->" & Space(1) & Format$(Mid$(wMensagem, 1, 66), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") & " <---"
  'Print #1, wtRAC2
  Print #1, WBrancos
  Print #1, "                                                                                "
  Print #1, WBrancos
  WCont = 0
Return

InicializaCampos:
 WTotalGeral = 0
 WTotalMerc = 0
 WTotalCarreto = 0
 WTotalDesconto = 0
Return

ImprimeCabecalhoEmpresa:
   Print #1, Chr(27) + Chr(18) + Chr(71) + WTrac ' comentado para teste em bo despacho
   Print #1, "### " & Format$(UCase(WEmpresa), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") & Format$(WEmail, "!@@@@@@@@@@@@@@@@@@@") & " ###"
   Print #1, "### " & UCase(WEndEmpresa) & Space(1) & "       " & Format(PED!datafechamento, "DD/mm/YYYY") & "-" & Format$(PED!HoraFechamento, "hh:mm:ss") & " ###"
   Print #1, "### " & wTelefone & Space(32) & Space(3) & Format$(IIf(IsNull(PED!datafechamento), "ORCAMENTO ", WTitulo), "!@@@@@@@@@@") & Format$(PED!Numos, "000,000") & "  ###"
   Print #1, WTrac2
   WCont = 6
Return

VerificaCliente:
  Cli.Seek "=", PED!CodigoCliente
  If Not Cli.NoMatch Then
  
     If IsNull(PED!FROTA) Then
        WFrota = " "
     Else
        WFrota = UCase(Mid$(PED!FROTA, 1, 11))
     End If
  
     If IsNull(PED!Observacoes) Then
        wobs = " "
     Else
        wobs = PED!Observacoes
     End If
     
     If IsNull(PED!PlacaVeiculo) Then
        wPlaca = "."
     Else
        wPlaca = Mid$(PED!PlacaVeiculo, 1, 3) & "-" & Mid$(PED!PlacaVeiculo, 4, 4)
     End If
     
     If IsNull(PED!MarcaVeiculo) Then
        WMarca = "."
     Else
        WMarca = Mid$(PED!MarcaVeiculo, 1, 11)
     End If
     
     If IsNull(PED!ModeloVeiculo) Then
        WModelo = "."
     Else
        WModelo = Mid$(PED!ModeloVeiculo, 1, 11)
    End If
     
     If IsNull(Cli!RazaoSocial) Then
        WNomeCli = "XXXXXXXXXXX"
     Else
        WNomeCli = Cli!RazaoSocial
     End If
     
     If IsNull(Cli!Nro) Then
        wnro = Space(5)
     Else
        wnro = Format$(Cli!Nro, "00000")
     End If
     
     If IsNull(Cli!Endereco) Then
        WEndCli = "."
     Else
        WEndCli = Trim$(Mid$(Cli!Endereco, 1, 45) & " " & wnro)
     End If
               
     If IsNull(Cli!Bairro) Then
        WBairro = "."
     Else
        WBairro = Mid$(Cli!Bairro, 1, 20)
        
     End If
     
     If IsNull(Cli!Cidade) Then
        WCidade = "."
     Else
        WCidade = Mid$(Cli!Cidade, 1, 20)
     End If
     
     If IsNull(Cli!EndPostal) Then
        wcep = Cep()
     Else
        wcep = Cli!EndPostal
     End If
     
     
     If IsNull(Cli!Estado) Then
        WUF = "."
     Else
        WUF = Cli!Estado
     End If
     If IsNull(Cli!CgcCpf) Then
        WCNPJ = "NAO FORNECIDO"
     Else
        WCNPJ = Cli!CgcCpf
     End If
     If IsNull(Cli!Inscricao) Then
        wInsc = "NAO FORNECIDO"
     Else
        wInsc = Cli!Inscricao
     End If
     If IsNull(Cli!Telefone) Then
        WTelcli = "."
     Else
        WTelcli = CapturaTelefoneCli(Cli!CodigoCliente)
     End If
     
     WTecnicos = CapturaEquipeTecnica(PED!Numos)
     
  End If
Return

ImprimeCabecalhoCliente:

   If IsNull(PED!KMEntrada) Then
      WKMEnt = 0
   Else
      WKMEnt = PED!KMEntrada
   End If
   
   If IsNull(PED!KMSaida) Then
      WKMSai = 0
   Else
      WKMSai = PED!KMSaida
   End If

   Print #1, "CLI..: " & Format$(PED!CodigoCliente, "000,000") & "-" & Format$(UCase(Mid$(WNomeCli, 1, 40)), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") & Space(1) & "FONE.: " & WTelcli
   Print #1, "END..: " & Format$(UCase(WEndCli), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") & Space(1) & "ATEND:" & Format$(Mid$(WVend, 1, 12), "!@@@@@@@@@@@@")
   Print #1, "CID..: " & Format$(wcep, "!@@@@@@@@") & "-" & Format$(UCase(WBairro), "!@@@@@@@@@@@@@@@@@@@@") & "-" & Format$(UCase(WCidade), "!@@@@@@@@@@@@@@@@@@@@") & Space(1) & "GARANTIA:" & WGarantia
   Print #1, "CNPJ.: " & Format$(WCNPJ, "!@@@@@@@@@@@@@@") & Space(8) & Space(29) & "KM..: " & Format$(WKMEnt, "000000") & " " & Format$(WKMSai, "000000")
   Print #1, "PRAZO: " & Format$(wdESCpRZ, "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") & " MODELO: " & WModelo
   Print #1, "TEC..: " & Format$(WTecnicos, "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") & " FROTA.: " & WFrota
   Print #1, "OBS..: " & Format$(UCase(Mid$(wobs, 1, 51)), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") & "PLACA.: " & wPlaca
   Print #1, WTrac1
   Print #1, "QUANT   UN DISCRIMINACOES                      VR.UNIT    VR.DESC   VLR.TOTAL"
   Print #1, WTrac1
   WCont = WCont + 10
Return

VerificaGarantia:
 GAR.Seek "=", WPedido
 If Not GAR.NoMatch Then
    If IsNull(GAR!DataFinal) Then
       WGarantia = Space(8)
    Else
       WGarantia = (Format$(GAR!DataFinal, "dd/mm/yyyy"))
    End If
 End If
Return


VerificaVENd:
 Resp.Seek "=", PED!Responsavel
 If Not Resp.NoMatch Then
    If IsNull(Resp!Nome) Then
       WVend = "."
    Else
       WVend = Mid$(Resp!Nome, 1, 16)
    End If
 End If
Return


VerificaPRZ:
  PRZ.Seek "=", PED!CodigoPagto
  If Not PRZ.NoMatch Then
     If IsNull(PRZ!Descricao) Then
        wdESCpRZ = "PRAZO NAO CADASTRADO"
     Else
        wdESCpRZ = PRZ!Descricao
     End If
  End If
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Public Function GetDate(Optional varDate As Variant) As Variant
    
    ' If varTempDate is missing then use today's date.
    ' Otherwise, set the date to the date passed.
    
    Dim varTempDate As Variant

    ' Set calendar date
    varTempDate = IIf(IsMissing(varDate), Date, varDate)
    
    ' Validate date
    If Not IsDate(varTempDate) Then varTempDate = Date
    DoCmd.OpenForm FormName:="frmCalendarControl", WindowMode:=acDialog, OpenArgs:=varTempDate

    ' If frmCalendarControl is still loaded, then the user clicked OK so get the
    ' date from the form.  If the form isn't open return a value of Null.
    If IsLoaded("frmCalendarControl") Then
        GetDate = Forms("frmCalendarControl").Calendar1.Value
        DoCmd.Close acForm, "frmCalendarControl"
    Else
        GetDate = Null
    End If
    
End Function

Function Conv_Data_NFE_Data(WData As String) As Variant
Conv_Data_NFE_Data = DateValue(Format$(Mid$(WData, 1, 10), "dd/mm/yyyy"))

End Function



Function Gera_Plac_Cli_Temp() As Long

On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Dim Qtb As QueryDef, numreg As Long, X As Long, Ret As Variant
Set bd = CurrentDb()
Set Qtb = bd.QueryDefs("OSVXAfeXEns")
Qtb.Parameters(0) = Forms!DialogoSelecaoMesRelInMetro!DataInicial
Qtb.Parameters(1) = Forms!DialogoSelecaoMesRelInMetro!DataFinal
Set tb = Qtb.OpenRecordset()

BeginTrans

If tb.RecordCount <> 0 Then
   While Not tb.EOF
         
         tb.Edit
         tb!Selecionar = 0
         tb.Update
         X = X + 1
         If X Mod 200 = 0 Then
            CommitTrans
            BeginTrans
         End If
         tb.MoveNext
   
   Wend
Else
   MsgBox "Nenhum registro selecionado...", vbCritical, "Aviso"
End If

Finaliza:
   CommitTrans
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Desmarca_OSV_AFE_ENS() As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Dim Qtb As QueryDef, numreg As Long, X As Long, Ret As Variant
Set bd = CurrentDb()
Set Qtb = bd.QueryDefs("OSVXAfeXEns")
Qtb.Parameters(0) = Forms!DialogoSelecaoMesRelInMetro!DataInicial
Qtb.Parameters(1) = Forms!DialogoSelecaoMesRelInMetro!DataFinal
Set tb = Qtb.OpenRecordset()

BeginTrans

If tb.RecordCount <> 0 Then
   While Not tb.EOF
         
         tb.Edit
         tb!Selecionar = 0
         tb.Update
         X = X + 1
         If X Mod 200 = 0 Then
            CommitTrans
            BeginTrans
         End If
         tb.MoveNext
   
   Wend
Else
   MsgBox "Nenhum registro selecionado...", vbCritical, "Aviso"
End If

Finaliza:
   CommitTrans
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function Seleciona_OSV_AFE_ENS() As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Dim Qtb As QueryDef, numreg As Long, X As Long, Ret As Variant
Set bd = CurrentDb()
Set Qtb = bd.QueryDefs("OSVXAfeXEns")
Qtb.Parameters(0) = Forms!DialogoSelecaoMesRelInMetro!DataInicial
Qtb.Parameters(1) = Forms!DialogoSelecaoMesRelInMetro!DataFinal
Set tb = Qtb.OpenRecordset()

BeginTrans

If tb.RecordCount <> 0 Then
   While Not tb.EOF
         
         tb.Edit
         tb!Selecionar = 1
         tb.Update
         X = X + 1
         If X Mod 200 = 0 Then
            CommitTrans
            BeginTrans
         End If
         tb.MoveNext
   
   Wend
Else
   MsgBox "Nenhum registro selecionado...", vbCritical, "Aviso"
End If

Finaliza:
   CommitTrans
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function


Function Valida_Emissao_Laudo_Ensaio_Afericao(WNumos As Long) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb, tba As Recordset
Dim WValida As Long
Dim X As Long
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("ITENSSERVOSV", DB_OPEN_TABLE)
Set tba = bd.OpenRecordset("EMPXTPSERVXREPASSE", DB_OPEN_TABLE)
tb.index = "SecondaryKey"
tba.index = "PrimaryKey"

WValida = 0

' 0 -> nao emite laudo
' 1 -> emite laudo

tb.Seek "=", WNumos
If Not tb.NoMatch Then
   While tb!Numos = WNumos
         tba.Seek "=", 1, tb!TipoServico
         If Not tba.NoMatch Then
            WValida = 1
            'GoSub Finaliza
        'Else
        '    WValida = 0
         End If
         tb.MoveNext
         If tb.EOF Then
            GoSub Finaliza:
         End If
   Wend
End If

Finaliza:
    Valida_Emissao_Laudo_Ensaio_Afericao = Round(WValida)
    tb.Close
    tba.Close
    bd.Close
    Set tb = Nothing
    Set tba = Nothing
    Set bd = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Ret_Vlr_Repasse_Serv_Inmetro_OSV(WNumos As Long) As Double
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb, tba As Recordset
Dim WTotal_Repasse  As Double
Dim X As Long
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("ITENSSERVOSV", DB_OPEN_TABLE)
Set tba = bd.OpenRecordset("EMPXTPSERVXREPASSE", DB_OPEN_TABLE)
tb.index = "SecondaryKey"
tba.index = "PrimaryKey"

WTotal_Repasse = 0

tb.Seek "=", WNumos
If Not tb.NoMatch Then
   While tb!Numos = WNumos
         tba.Seek "=", Form_Menu.WEmpresa, tb!TipoServico
         If Not tba.NoMatch Then
            'If Tba!retencao Then
               WTotal_Repasse = (((tb!ValorMaoObra - tb!Desconto)) * tba!Repasse) / 100
            'End If
         End If
         tb.MoveNext
         If tb.EOF Then
            GoSub Finaliza:
         End If
   Wend
End If

Finaliza:
    Ret_Vlr_Repasse_Serv_Inmetro_OSV = Round(WTotal_Repasse)
    tb.Close
    tba.Close
    bd.Close
    Set tb = Nothing
    Set tba = Nothing
    Set bd = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function EnviarMail(WRemetente As String, WDestinatario As String, WAssunto As String, WAnexo As String, WCorpo As String)

On Error GoTo TratamentodeErro:

DoCmd.Hourglass True

On Error GoTo TratamentodeErro:
   
Set oemail = New CanadaUtils.Email
oemail.Assunto = WAssunto
oemail.ContaEMail = "sistema.canada@gmail.com" ' Conta de email da empresa
oemail.EmailDestino = WDestinatario
oemail.Mensagem = WCorpo & vbCrLf & vbCrLf & vbCrLf
oemail.Mensagem = oemail.Mensagem & Captura_Conteudo_Campos_Empresa(Form_Menu.WEmpresa, "RazaoSocial") & vbCrLf
oemail.Mensagem = oemail.Mensagem & "Tel: " & Telefone() & vbCrLf
oemail.Mensagem = oemail.Mensagem & Email() & vbCrLf
oemail.Mensagem = oemail.Mensagem & Cidade() & "," & Format$(Date, "dd/mm/yyyy") & " " & Format$(Time(), "hh:mm:ss")
oemail.SenhaEmail = "N@mbuco1986" ' Senha email empresa
oemail.SMTP = "smtp.gmail.com" ' Servidor email da empresa

If WAnexo = "#" Then
   oemail.Anexo = LogomarcaEmpresa()
Else
   oemail.Anexo = WAnexo
End If
If oemail.ValidaEnvioEmail = True Then
   oemail.EnviarEmail
   If Not oemail.StatusEnvio = True Then
      DoCmd.Hourglass False
      MsgBox oemail.MensagemEnvio, vbCritical, "SGa"
      GoSub Finaliza:
   End If
End If
Finaliza:
   DoCmd.Hourglass False
   Set oemail = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function



Function RetornaItensAtivosXNivelEstMin() As Double
'=SeImed([Sequencia]<=[NumClientes]*0,25;"A";SeImed([Sequencia]>[NumClientes]*0,25 E [Sequencia]<=[NumClientes]*0,6;"B";"C"))

On Error GoTo TratamentodeErro:
DoCmd.Hourglass True
Dim db As DataBase, tb, tba As Recordset
Dim WAtivos As Long
Set db = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = db.OpenRecordset("EST", DB_OPEN_TABLE)
Set tba = db.OpenRecordset("EMPXEST", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tba.index = "SecondaryKey"

GoSub ApuraQtdeItensAtivos:

Finaliza:
  DoCmd.Hourglass False
  RetornaItensAtivosXNivelEstMin = WAtivos
  tb.Close
  tba.Close
  db.Close
  Set tb = Nothing
  Set tba = Nothing
  Set db = Nothing
  Exit Function
Return

ApuraQtdeItensAtivos:
  If tb.RecordCount <> 0 Then
     While Not tb.EOF
           If tb!Ativo Then
              tba.Seek "=", tb!CodItem
              If Not tba.NoMatch Then
                 If Form_Menu.WEmpresa = tba!CodigoEmpresa Then
                    If Not IsNull(tba!Minimo) Then
                       If tba!Minimo > 0 Then
                          WAtivos = WAtivos + 1
                       End If
                    End If
                 End If
              End If
           End If
           tb.MoveNext
     Wend
  End If
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Retorna_ChequesPredatados_Dem_Rec(WNumos As Long) As Double
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("DEMONSTRATIVORECEITAS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"


tb.Seek "=", WNumos
If Not tb.NoMatch Then
   If IsNull(tb!ChequesPreDatados) Then
      Retorna_ChequesPredatados_Dem_Rec = 0
   Else
      Retorna_ChequesPredatados_Dem_Rec = tb!ChequesPreDatados
   End If
Else
   Retorna_ChequesPredatados_Dem_Rec = 0
End If


Finaliza:
    tb.Close
    bd.Close
    Set tb = Nothing
    Set bd = Nothing

    Exit Function
Return


TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Retorna_Vale_Dem_Rec(WNumos As Long) As Double
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("DEMONSTRATIVORECEITAS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"


tb.Seek "=", WNumos
If Not tb.NoMatch Then
   If IsNull(tb!Vale) Then
      Retorna_Vale_Dem_Rec = 0
   Else
      Retorna_Vale_Dem_Rec = tb!Vale
   End If
Else
   Retorna_Vale_Dem_Rec = 0
End If

Finaliza:
    tb.Close
    bd.Close
    Set tb = Nothing
    Set bd = Nothing

    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Retorna_Garantia_Dem_Rec(WNumos As Long) As Double
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("DEMONSTRATIVORECEITAS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"


tb.Seek "=", WNumos
If Not tb.NoMatch Then
   If IsNull(tb!Garantia) Then
      Retorna_Garantia_Dem_Rec = 0
   Else
      Retorna_Garantia_Dem_Rec = tb!Garantia
   End If
Else
   Retorna_Garantia_Dem_Rec = 0
End If


Finaliza:
    tb.Close
    bd.Close
    Set tb = Nothing
    Set bd = Nothing

    Exit Function
Return


TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Retorna_Duplicatas_Dem_Rec(WNumos As Long) As Double
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("DEMONSTRATIVORECEITAS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"


tb.Seek "=", WNumos
If Not tb.NoMatch Then
   If IsNull(tb!Duplicatas) Then
      Retorna_Duplicatas_Dem_Rec = 0
   Else
      Retorna_Duplicatas_Dem_Rec = tb!Duplicatas
   End If
Else
   Retorna_Duplicatas_Dem_Rec = 0
End If


Finaliza:
    tb.Close
    bd.Close
    Set tb = Nothing
    Set bd = Nothing

    Exit Function
Return


TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Retorna_ChequesVista_Dem_Rec(WNumos As Long) As Double
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("DEMONSTRATIVORECEITAS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"


tb.Seek "=", WNumos
If Not tb.NoMatch Then
   If IsNull(tb!ChequesVista) Then
      Retorna_ChequesVista_Dem_Rec = 0
   Else
      Retorna_ChequesVista_Dem_Rec = tb!ChequesVista
   End If
Else
   Retorna_ChequesVista_Dem_Rec = 0
End If

Finaliza:
    tb.Close
    bd.Close
    Set tb = Nothing
    Set bd = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Retorna_Moedas_Dem_Rec(WNumos As Long) As Double
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("DEMONSTRATIVORECEITAS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"


tb.Seek "=", WNumos
If Not tb.NoMatch Then
   If IsNull(tb!Moedas) Then
      Retorna_Moedas_Dem_Rec = 0
   Else
      Retorna_Moedas_Dem_Rec = tb!Moedas
   End If
Else
   Retorna_Moedas_Dem_Rec = 0
End If


Finaliza:
    tb.Close
    bd.Close
    Set tb = Nothing
    Set bd = Nothing
    Exit Function
Return


TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Retorna_Dinheiro_Dem_Rec(WNumos As Long) As Double
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("DEMONSTRATIVORECEITAS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"


tb.Seek "=", WNumos
If Not tb.NoMatch Then
   If IsNull(tb!Dinheiro) Then
      Retorna_Dinheiro_Dem_Rec = 0
   Else
      Retorna_Dinheiro_Dem_Rec = tb!Dinheiro
   End If
Else
   Retorna_Dinheiro_Dem_Rec = 0
End If


Finaliza:
    tb.Close
    bd.Close
    Set tb = Nothing
    Set bd = Nothing
    Exit Function
Return


TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function


Function CriaModeloEquipTemp(WMarca As String) As Long

On Error GoTo TratamentodeErro:
Dim bd, bdc As DataBase
Dim tb, tbc As Recordset
Set bdc = CurrentDb()
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("MODELOEQUIP", DB_OPEN_TABLE)
Set tbc = bdc.OpenRecordset("MODELOEQUIPTEMP", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
'tbc.Index = "PrimaryKey"

GoSub InicializaTemp:
GoSub AtualizaTemp:


Finaliza:
    tb.Close
    tbc.Close
    bd.Close
    bdc.Close
    Set tb = Nothing
    Set tbc = Nothing
    Set bd = Nothing
    Set bdc = Nothing
    Exit Function
Return

AtualizaTemp:
    If tb.RecordCount <> 0 Then
       While Not tb.EOF
             BeginTrans
             If tb!Marca = WMarca Then
                tbc.AddNew
                tbc!Modelo = tb!Modelo
                tbc.Update
             End If
             CommitTrans
             tb.MoveNext
       Wend
    End If
Return

InicializaTemp:
  If tbc.RecordCount <> 0 Then
     BeginTrans
     While Not tbc.EOF
           tbc.Delete
           tbc.MoveNext
     Wend
     CommitTrans
  End If

Return


TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function SubstituiItemNFXItemOSV(WNumos As Long, WCodItemNovo As Long, WCodItemVelho As Long) As Long
On Error GoTo TratamentodeErro:
DoCmd.Hourglass True
Dim db As DataBase
Dim tb As Recordset
Dim tba As Recordset
Set db = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = db.OpenRecordset("RQM", DB_OPEN_TABLE)
Set tba = db.OpenRecordset("RMXEMPXEST", DB_OPEN_TABLE)
tb.index = "SecondaryKey"
tba.index = "RQM"

BeginTrans

tb.Seek "=", WNumos
If Not tb.NoMatch Then
   While Not tb.EOF
         If tb!Numos <> WNumos Then
            GoSub Finaliza:
         End If
         GoSub ProcuraItem:
         tb.MoveNext
   Wend
End If


Finaliza:
  DoCmd.Hourglass False
  CommitTrans
  tb.Close
  tba.Close
  db.Close
  Set tb = Nothing
  Set tba = Nothing
  Set db = Nothing
  Exit Function
Return

ProcuraItem:
  tba.Seek "=", tb!NumRequisicao
  If Not tba.NoMatch Then
     While Not tba.EOF
          If tba!NumRequisicao <> tb!NumRequisicao Then
             Return
          End If
          If tba!CodItem = WCodItemVelho Then
             tba.Edit
             tba!CodItem = WCodItemNovo
             tba.Update
             Return
          End If
          tba.MoveNext
     Wend
  End If
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function AtualizaABCEstoque(WQTdeAtivos As Long) As Long

'=SeImed([Sequencia]<=[NumClientes]*0,25;"A";SeImed([Sequencia]>[NumClientes]*0,25 E [Sequencia]<=[NumClientes]*0,6;"B";"C"))

On Error GoTo TratamentodeErro:
DoCmd.Hourglass True
Dim db As DataBase
Dim Qtb As QueryDef
Dim tb As Recordset
Dim WSeq As Long
Dim WNivel As String * 1
Set db = CurrentDb()
Set Qtb = db.QueryDefs("ClassifABCEstoque")
Set tb = Qtb.OpenRecordset()

BeginTrans

If tb.RecordCount <> 0 Then
   While Not tb.EOF
         WSeq = WSeq + 1
         If WSeq <= (WQTdeAtivos * 0.2) Then
            WNivel = "A"
         ElseIf WSeq > (WQTdeAtivos * 0.2) And WSeq <= (WQTdeAtivos * 0.3) Then
                WNivel = "B"
         Else
             WNivel = "C"
         End If
    
         tb.Edit
         tb!NIVELESTMIN = WNivel
         tb.Update
         tb.MoveNext
   Wend
End If


Finaliza:
  DoCmd.Hourglass False
  CommitTrans
  tb.Close
  db.Close
  Set tb = Nothing
  Set db = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function


Function Retorna_Numosv_GRU(WGRU As String) As String
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "GRU"

tb.Seek "=", WGRU
If tb.NoMatch Then
   Retorna_Numosv_GRU = 0
Else
   Retorna_Numosv_GRU = tb!Numos
End If

Finaliza:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function RetornaCodigoPagtoCliente(WCodigoCliente As Long) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("CLI", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

tb.Seek "=", WCodigoCliente
If tb.NoMatch Then
   RetornaCodigoPagtoCliente = 0
Else
   If IsNull(tb!CodigoPagto) Then
      RetornaCodigoPagtoCliente = 0
   Else
      RetornaCodigoPagtoCliente = tb!CodigoPagto
   End If
End If

Finaliza:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function RetornaTipoLaudoXOSV(WNumos As Long) As String
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

tb.Seek "=", WNumos
If tb.NoMatch Then
   RetornaTipoLaudoXOSV = "Z"
Else
   If IsNull(tb!TipoLaudo) Then
      RetornaTipoLaudoXOSV = "Z"
   Else
      RetornaTipoLaudoXOSV = tb!TipoLaudo
   End If
End If

Finaliza:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

'Function ReCalculaComissao(WNumos As Long, WTipo As Long, WApuracao As Long) As Long
'On Error GoTo TratamentodeErro:
'Dim TBOSV As Recordset, bd As DataBase
'Dim QTBOSV As QueryDef
'Dim X As Long
'Set bd = CurrentDb()
'Set QTBOSV = bd.QueryDefs("RecalculaComissoes")
'QTBOSV.Parameters(0) = WNumos
'Set TBOSV = QTBOSV.OpenRecordset()

'If TBOSV.RecordCount <> 0 Then
'   While Not TBOSV.EOF
'
'         If WApuracao = 1 Then
'            If WTipo = 1 Then
'               Call GeraAtexServXOSV(WNumos) ' OK REVISAO FEITA PARA QUERY
'            ElseIf WTipo = 2 Then
'                   Call GeraTecxServXOSV(WNumos) ' OK REVISAO FEITA PARA QUERY
'            Else
'               Call GeraVenxServXOSV(WNumos) ' OK REVISAO FEITA PARA QUERY
'            End If
'         ElseIf WApuracao = 2 Then
'                If WTipo = 1 Then
'                   Call GeraAteXCatXOSV(WNumos) ' OK , REVISAO FEITA PARA QUERY
'                ElseIf WTipo = 2 Then
'                       Call GeraTecxCatXOSV(WNumos) ' OK REVISAO FEITA PARA QUERY
'                Else
'                    Call GeraVenXCatXOSV(WNumos) ' OK REVISAO FEITA PARA QUERY
'                End If
'         End If
'         TBOSV.MoveNext
'   Wend
'End If

'Finaliza:
'   TBOSV.Close
'   bd.Close
'   Set TBOSV = Nothing
'   Set bd = Nothing
'   Exit Function
'Return


'TratamentodeErro:
'    If Err.Number = 3001 Then
'      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
'    Else
'       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description, Vbcritical,"SGa"
'       Resume Next
'    End If
'End Function
Function CriaEstxEtiq() As Long
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim QTbQ As QueryDef
Dim tb, Tbq As Recordset
Dim WGer As Long
Dim WQuant As Long
Set bd1 = CurrentDb()
Set tb = bd1.OpenRecordset("ESTTEMP", DB_OPEN_TABLE)
Set QTbQ = bd1.QueryDefs("EstSelecionados")
Set Tbq = QTbQ.OpenRecordset()

BeginTrans

If tb.RecordCount <> 0 Then
   While Not tb.EOF
         tb.Delete
         tb.MoveNext
   Wend
End If

If Tbq.RecordCount <> 0 Then
   While Not Tbq.EOF
        If Tbq!Selecionar Then
           WQuant = Tbq!Quantidade_Eti
           GoSub GeraEstTemp
        End If
        Tbq.MoveNext
   Wend
End If

Finaliza:
  CommitTrans
  tb.Close
  Tbq.Close
  bd1.Close
  Set tb = Nothing
  Set Tbq = Nothing
  Set bd1 = Nothing
  Exit Function
Return


GeraEstTemp:
  WGer = 0
  While WGer < WQuant
        tb.AddNew
        tb!CodItem = Tbq!CodItem
        tb.Update
        WGer = WGer + 1
   Wend
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function
Function CriaClixETI() As Long
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim QTbQ As QueryDef
Dim tb, Tbq As Recordset
Dim WGer As Long
Dim WQuant As Long
Set bd1 = CurrentDb()
Set tb = bd1.OpenRecordset("CLITEMP", DB_OPEN_TABLE)
Set QTbQ = bd1.QueryDefs("CliSelecionados")
Set Tbq = QTbQ.OpenRecordset()

BeginTrans

If tb.RecordCount <> 0 Then
   While Not tb.EOF
         tb.Delete
         tb.MoveNext
   Wend
End If

If Tbq.RecordCount <> 0 Then
   While Not Tbq.EOF
        If Tbq!Selecionar Then
           WQuant = Tbq!Quantidade
           GoSub GeraEstTemp
        End If
        Tbq.MoveNext
   Wend
End If

Finaliza:
  CommitTrans
  tb.Close
  Tbq.Close
  bd1.Close
  Set tb = Nothing
  Set Tbq = Nothing
  Set bd1 = Nothing
  Exit Function
Return

GeraEstTemp:
  WGer = 1
  While WGer <= WQuant
        tb.AddNew
        tb!CodigoCliente = Tbq!CodigoCliente
        tb.Update
        WGer = WGer + 1
   Wend
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CriaEtiCli() As Long
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim QTbQ As QueryDef
Dim tb, Tbq As Recordset
Dim WGer As Long
Set bd1 = CurrentDb()
Set tb = bd1.OpenRecordset("CLIEST", DB_OPEN_TABLE)
Set QTbQ = bd1.QueryDefs("CliQ")
Set Tbq = QTbQ.OpenRecordset()

BeginTrans

If tb.RecordCount <> 0 Then
   While Not tb.EOF
         tb.Delete
         tb.MoveNext
   Wend
End If

If Tbq.RecordCount <> 0 Then
   While Not Tbq.EOF
        tb.AddNew
        tb!CodigoCliente = Tbq!CodigoCliente
        tb!Quantidade = 1
        tb.Update
        Tbq.MoveNext
   Wend
End If

Finaliza:
  CommitTrans
  tb.Close
  Tbq.Close
  bd1.Close
  Set tb = Nothing
  Set Tbq = Nothing
  Set bd1 = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function CriaEtiEst() As Long
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim QTbQ As QueryDef
Dim tb, Tbq As Recordset
Dim WGer As Long
Set bd1 = CurrentDb()
Set tb = bd1.OpenRecordset("ETIEST", DB_OPEN_TABLE)
Set QTbQ = bd1.QueryDefs("EstQ")
Set Tbq = QTbQ.OpenRecordset()

BeginTrans

If tb.RecordCount <> 0 Then
   While Not tb.EOF
         tb.Delete
         tb.MoveNext
   Wend
End If

If Tbq.RecordCount <> 0 Then
   While Not Tbq.EOF
        tb.AddNew
        tb!CodItem = Tbq!CodItem
        tb!Quantidade = 1
        tb.Update
        Tbq.MoveNext
   Wend
End If

Finaliza:
  CommitTrans
  tb.Close
  Tbq.Close
  bd1.Close
  Set tb = Nothing
  Set Tbq = Nothing
  Set bd1 = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function InicializaCliXTemp() As Long
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim tb As Recordset
Set bd1 = CurrentDb()
Set tb = bd1.OpenRecordset("CLITEMP", DB_OPEN_TABLE)

BeginTrans

If tb.RecordCount <> 0 Then
   While Not tb.EOF
         tb.Delete
         tb.MoveNext
   Wend
End If

Finaliza:
  CommitTrans
  tb.Close
  bd1.Close
  Set tb = Nothing
  Set bd1 = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function
Function Inicializa_PlacxCli_Temp() As Long
On Error GoTo TratamentodeErro:
Dim bdc, bdt As DataBase
Dim tba, tbc As Recordset
Dim X As Long
Set bdc = CurrentDb()
Set bdt = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Dim qtbc As QueryDef
Set qtbc = bdc.QueryDefs("ConsultaLotePlacasxInmetro")
qtbc.Parameters(0) = Forms!DialogoSelecaoPlacXCliXPrinc!Cliente
qtbc.Parameters(1) = Forms!DialogoSelecaoPlacXCliXPrinc!Cidade
qtbc.Parameters(2) = Forms!DialogoSelecaoPlacXCliXPrinc!UF
qtbc.Parameters(3) = Forms!DialogoSelecaoPlacXCliXPrinc!Placa

Set tbc = qtbc.OpenRecordset()
Set tba = bdt.OpenRecordset("PLACXCLITEMP", DB_OPEN_TABLE)
tba.index = "PrimaryKey"

GoSub LimpaPlacXCliXTemp:

If tbc.RecordCount <> 0 Then
   X = 0
   BeginTrans
   
   While Not tbc.EOF
   
         tba.Seek "=", tbc!Placa, tbc!CodigoCliente
         If tba.NoMatch Then
            tba.AddNew
         Else
            tba.Edit
         End If
         'tba!RazaoSocial = tbc!RazaoSocial
         'tba!Cidade = tbc!Cidade
         'tba!Estado = tbc!Estado
         tba!CodigoCliente = tbc!CodigoCliente
         tba!Placa = tbc!Placa
         tba!RENAVAN = tbc!RENAVAN
'         tba!StatusPlaca = tbc!StatusPlaca
         tba!VenctoCert = tbc!VenctoCert
         tba!CHASSI = tbc!CHASSI
         tba!Pneu = tbc!Pneu
         tba!MarcaVeiculo = tbc!MarcaVeiculo
         tba!ModeloVeiculo = tbc!ModeloVeiculo
         tba!FROTA = tbc!FROTA
         tba!AnoVeiculo = tbc!AnoVeiculo
         tba!AroPneu = tbc!AroPneu
         tba!DataSelagem = tbc!DataSelagem
         tba!PostoSelagem = tbc!PostoSelagem
         tba!DataEnsaio = tbc!DataEnsaio
         tba!DataResultado = tbc!DataResultado
         tba!TipoCertificadoResultado = tbc!TipoCertificadoResultado
         tba!NumeroCertificadoDefinitivo = tbc!NumeroCertificadoDefinitivo
         tba!MarcaEquip = tbc!MarcaEquip
         tba!ModeloEquip = tbc!ModeloEquip
         tba!NroSerieEquip = tbc!NroSerieEquip
         tba!DataConsulta = tbc!DataConsulta
         tba!StatusPlaca = tbc!StatusPlaca
         tba!VenctoCert = tbc!VenctoCert
         If IsNull(tba!Status) Then
             tba!Status = 1
         End If
         tba.Update
         X = X + 1
         If X Mod 200 = 0 Then
            CommitTrans
            BeginTrans
         End If
         tbc.MoveNext
   Wend
   CommitTrans
End If

Finaliza:
  tba.Close
  tbc.Close
  bdc.Close
  bdt.Close
  Set tba = Nothing
  Set tbc = Nothing
  Set bdc = Nothing
  Set bdt = Nothing
  Exit Function
Return

LimpaPlacXCliXTemp:
   BeginTrans
   X = 0
   If tba.RecordCount <> 0 Then
      While Not tba.EOF
            tba.Delete
            tba.MoveNext
            X = X + 1
            If X Mod 200 = 0 Then
              CommitTrans
              BeginTrans
            End If
      Wend
   End If
   CommitTrans
Return


TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Limpa_Plac_Cli_Temp() As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Dim X As Long
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("PLACXCLITEMP", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

If tb.RecordCount <> 0 Then
   X = 0
   BeginTrans
   While Not tb.EOF
         tb.Delete
         X = X + 1
         If X Mod 200 = 0 Then
            CommitTrans
            BeginTrans
         End If
         tb.MoveNext
   Wend
   CommitTrans
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function InicializaEstXTemp() As Long
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim tb As Recordset
Set bd1 = CurrentDb()
Set tb = bd1.OpenRecordset("ESTTEMP", DB_OPEN_TABLE)

BeginTrans

If tb.RecordCount <> 0 Then
   While Not tb.EOF
         tb.Delete
         tb.MoveNext
   Wend
End If

Finaliza:
  CommitTrans
  tb.Close
  bd1.Close
  Set tb = Nothing
  Set bd1 = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function


Function ValidaOSV(WNumos As Long) As Boolean
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb  As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

tb.Seek "=", WNumos
If tb.NoMatch Then
   ValidaOSV = False
Else
   ValidaOSV = True
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function




Function TipoLaudoStr(WTipo As String) As String
On Error GoTo TratamentodeErro:

      If WTipo = "A" Then
         TipoLaudoStr = "SELAGEM"
      ElseIf WTipo = "S" Then
           TipoLaudoStr = "SUBSTITUICAO"
      ElseIf WTipo = "T" Then
              TipoLaudoStr = "SELAGEM BANCADA"
      ElseIf WTipo = "E" Then
              TipoLaudoStr = "ENSAIO"
      ElseIf WTipo = "F" Then
              TipoLaudoStr = "SELAGEM E ENSAIO"
      ElseIf WTipo = "B" Then
              TipoLaudoStr = "SUBSTITUICAO E ENSAIO"
      Else
              TipoLaudoStr = ""
      End If
   Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function CapturaTabelaServXClientes(WCodigoCliente As Long) As Long
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim tb As Recordset
Dim WTabela As Double
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())

Set tb = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
BeginTrans

tb.Seek "=", WCodigoCliente
If tb.NoMatch Then
   WTabela = 1
Else
   If Not IsNull(tb!TabelaServ) Then
      WTabela = tb!TabelaServ
   Else
      tb.Edit
      tb!TabelaServ = 1
      tb.Update
      WTabela = 1
   End If
End If
   
Finaliza:
  CommitTrans
  tb.Close
  CapturaTabelaServXClientes = WTabela
  bd1.Close
  Set tb = Nothing
  Set bd1 = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function CapturaDescricaoTabPecas(WTabela As Long) As String
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim tb As Recordset
Dim WDelta As String
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd1.OpenRecordset("TAB", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

tb.Seek "=", WTabela
If Not tb.NoMatch Then
   If Not IsNull(tb!Nome) Then
      CapturaDescricaoTabPecas = tb!Nome
   Else
      CapturaDescricaoTabPecas = "X"
   End If
End If
   
Finaliza:
  tb.Close
  bd1.Close
  Set tb = Nothing
  Set bd1 = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function



Function CapturaDescricaoTabServ(WTabela As Long) As String
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim tb As Recordset
Dim WDelta As String
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd1.OpenRecordset("TABSERV", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

tb.Seek "=", WTabela
If Not tb.NoMatch Then
   If Not IsNull(tb!Nome) Then
      CapturaDescricaoTabServ = tb!Nome
   Else
      CapturaDescricaoTabServ = "X"
   End If
End If
   
Finaliza:
  tb.Close
  bd1.Close
  Set tb = Nothing
  Set bd1 = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function Inclui_Altera_Fornecedor(WCodigoFornec As Long, WRazaoSocial As String, WInscricao As String, WCNPJ As String, WEndereco As String, WBairro As String, WEndPostal As String, WCidade As String, WEstado As String, wnro As String) As Long
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("FOR", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
BeginTrans
tb.Seek "=", WCodigoFornec
If tb.NoMatch Then
   tb.AddNew
   tb!CodigoFornecedor = WCodigoFornec
Else
   tb.Edit
End If
tb!RazaoSocial = Mid$(UCase(WRazaoSocial), 1, 50)
tb!Codinome = Mid$(UCase(WRazaoSocial), 1, 30)
tb!Tipo = 2
tb!Inscricao = WInscricao
tb!CgcCpf = Mid$(WCNPJ, 1, 14)
tb!Endereco = Mid$(UCase(WEndereco), 1, 50)
tb!Bairro = Mid$(UCase(WBairro), 1, 20)
tb!EndPostal = Mid$(WEndPostal, 1, 8)
tb!Cidade = Mid$(UCase(WCidade), 1, 30)
tb!Estado = Mid$(UCase(WEstado), 1, 2)
tb!Nro = IIf(IsNumeric(wnro), wnro, 0)
tb.Update
CommitTrans
Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Ret_Cod_Fornec_CNPJ(WCNPJ As String) As Long
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("FOR", DB_OPEN_TABLE)
tb.index = "cgccpf"

tb.Seek "=", WCNPJ
If Not tb.NoMatch Then
  Ret_Cod_Fornec_CNPJ = tb!CodigoFornecedor
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function FornecInvalido(WCNPJ As String) As Boolean
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("FOR", DB_OPEN_TABLE)
tb.index = "cgccpf"

tb.Seek "=", WCNPJ
If tb.NoMatch Then
   FornecInvalido = True
Else
   FornecInvalido = False
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Retorna_Tipo_Recurso_Trans(WCodigo As Long) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("TRANSXCART", DB_OPEN_TABLE)

tb.index = "PrimaryKey"
tb.Seek "=", WCodigo

If Not tb.NoMatch Then
   Retorna_Tipo_Recurso_Trans = tb!TipoRecurso
End If

Finaliza:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Retorna_Aprop_Trans(WCodigo As Long) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("TRANSXCART", DB_OPEN_TABLE)

tb.index = "PrimaryKey"
tb.Seek "=", WCodigo

If Not tb.NoMatch Then
   Retorna_Aprop_Trans = tb!Apropriacao
End If

Finaliza:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Retorna_Hist_Destino_Trans(WCodigo As Long)
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("TRANSXCART", DB_OPEN_TABLE)

tb.index = "PrimaryKey"
tb.Seek "=", WCodigo

If Not tb.NoMatch Then
   Retorna_Hist_Destino_Trans = tb!HistoricoDestino
End If

Finaliza:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Retorna_Tipo_Destino_Trans(WCodigo As Long) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("TRANSXCART", DB_OPEN_TABLE)

tb.index = "PrimaryKey"
tb.Seek "=", WCodigo

If Not tb.NoMatch Then
   Retorna_Tipo_Destino_Trans = tb!TipoDestino
End If

Finaliza:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Atualiza_Status_Selagem_OSV(WNumos As Long, WStatus As String) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

tb.Seek "=", WNumos
BeginTrans
If Not tb.NoMatch Then
   tb.Edit
   tb!Status_Selagem = WStatus
   tb.Update
End If
CommitTrans

Finaliza:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Retorna_Status_Selagem_OSv(WNumos As Long) As String

On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)

tb.index = "PrimaryKey"
tb.Seek "=", WNumos

If Not tb.NoMatch Then
   If IsNull(tb!Status_Selagem) Then
      Retorna_Status_Selagem_OSv = Space(1)
   Else
      Retorna_Status_Selagem_OSv = tb!Status_Selagem
   End If
Else
   Retorna_Status_Selagem_OSv = Space(1)
   
End If

Finaliza:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ProxCodVagoTabela(WTabela As String) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Dim WCodigo As Long
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset(WTabela, DB_OPEN_TABLE)

tb.index = "PrimaryKey"
WCodigo = 1
tb.Seek "=", WCodigo

If Not tb.NoMatch Then
   While Not tb.NoMatch
      WCodigo = WCodigo + 1
      tb.Seek "=", WCodigo
   Wend
   ProxCodVagoTabela = WCodigo
Else
   ProxCodVagoTabela = WCodigo
End If

Finaliza:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function ProximoCodigoFornec(WCodigo As Long) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("FOR", DB_OPEN_TABLE)

tb.index = "PrimaryKey"
tb.Seek "=", WCodigo

If Not tb.NoMatch Then
   While Not tb.NoMatch
      WCodigo = WCodigo + 1
      tb.Seek "=", WCodigo
   Wend
   ProximoCodigoFornec = WCodigo
Else
   ProximoCodigoFornec = WCodigo
End If

Finaliza:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function


Function ProximoCodigoCliente(WCodigo As Long) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("CLI", DB_OPEN_TABLE)

tb.index = "PrimaryKey"
tb.Seek "=", WCodigo

If Not tb.NoMatch Then
   While Not tb.NoMatch
      WCodigo = WCodigo + 1
      tb.Seek "=", WCodigo
   Wend
   ProximoCodigoCliente = WCodigo
Else
   ProximoCodigoCliente = WCodigo
End If

Finaliza:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function ProximoTRansp(WCodigo As Long) As Long
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim tb As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd1.OpenRecordset("Transp", DB_OPEN_TABLE)

tb.index = "PrimaryKey"
    
'BeginTrans
tb.Seek "=", WCodigo

If Not tb.NoMatch Then
   While Not tb.NoMatch
      WCodigo = WCodigo + 1
      tb.Seek "=", WCodigo
   Wend
   ProximoTRansp = WCodigo
Else
   ProximoTRansp = WCodigo
End If
'CommitTrans
tb.Close
bd1.Close
Set tb = Nothing
Set bd1 = Nothing
Exit Function

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function ProximoCodigoRef(WCodigo As Long) As Long
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim tb As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd1.OpenRecordset("ClixRef", DB_OPEN_TABLE)

tb.index = "SecondKey"

'BeginTrans
tb.Seek "=", WCodigo

If Not tb.NoMatch Then
   While Not tb.NoMatch
      WCodigo = WCodigo + 1
      tb.Seek "=", WCodigo
   Wend
   ProximoCodigoRef = WCodigo
Else
  ProximoCodigoRef = WCodigo
End If
'CommitTrans
tb.Close
bd1.Close
Set tb = Nothing
Set bd1 = Nothing
Exit Function

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaUltimaAquisicaoRegInv(WCodItem As Long) As Double
On Error GoTo TratamentodeErro:
Dim Ult As Recordset
Dim bd As DataBase
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set Ult = bd.OpenRecordset("UltAquiRegInv", DB_OPEN_TABLE)
Ult.index = "PrimaryKey"

Ult.Seek "=", WCodItem
If Ult.NoMatch Then
   CapturaUltimaAquisicaoRegInv = 0
Else
   CapturaUltimaAquisicaoRegInv = Ult!ValorUnitario
End If


Finaliza:
  Ult.Close
  bd.Close
  Set Ult = Nothing
   Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function



Function CriaUltAquiRegInv() As Long
On Error GoTo TratamentodeErro:
Dim Ult, UltAQRI As Recordset
Dim WVlrUnitario As Double
Dim X As Long
Dim WOutras, WFrete, WIcmsretido As Double
Dim bd, bdt As DataBase
Dim QUlt As QueryDef
Set bd = CurrentDb()
Set bdt = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set QUlt = bd.QueryDefs("QUltAquis")
Set UltAQRI = bdt.OpenRecordset("UltAquiRegInv", DB_OPEN_TABLE)
Set Ult = QUlt.OpenRecordset()
UltAQRI.index = "PrimaryKey"

If UltAQRI.RecordCount <> 0 Then
   BeginTrans
   X = 0
   While Not UltAQRI.EOF
         UltAQRI.Delete
         X = X + 1
         If X Mod 500 = 0 Then
            CommitTrans
            BeginTrans
         End If
         UltAQRI.MoveNext
   Wend
   CommitTrans
End If
BeginTrans
If Ult.RecordCount <> 0 Then
   X = 0
   While Not Ult.EOF
      Ult.Edit
       
      If IsNull(Ult!OUTRASDESPESAS) Then
         Ult!OUTRASDESPESAS = 0
      End If
          
      If IsNull(Ult!VlrFrete) Then
         Ult!VlrFrete = 0
      End If
      
     If IsNull(Ult!ICMSRETIDO) Then
        Ult!ICMSRETIDO = 0
     End If
          
     WOutras = Ult!OUTRASDESPESAS
     WFrete = Ult!VlrFrete
     WIcmsretido = Ult!ICMSRETIDO
          
      WVlrUnitario = Ult!ValorUnit + WFrete + WIcmsretido + WOutras + (Round(Ult!ValorUnit * Ult!IPI / 100))
      If WVlrUnitario > 0 Then
         UltAQRI.Seek "=", Ult!CodItem
         If UltAQRI.NoMatch Then
            UltAQRI.AddNew
         Else
            UltAQRI.Edit
         End If
      
         UltAQRI!CodItem = Ult!CodItem
         UltAQRI!ValorUnitario = WVlrUnitario
         UltAQRI.Update
      End If
      X = X + 1
      If X Mod 500 = 0 Then
         CommitTrans
         BeginTrans
         End If
      Ult.MoveNext
   Wend
   
End If


Finaliza:
  CommitTrans
  Ult.Close
  UltAQRI.Close
  bd.Close
  bdt.Close
  Set Ult = Nothing
  Set UltAQRI = Nothing
  Set bd = Nothing
  Set bdt = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function SelecionaFormBaixaCTRECXBalan() As Long

 If Nroserie() = 111999 Or Nroserie() = 12007 Then
      DoCmd.OpenForm "EfetuarRecebimentosPrinc", acNormal
   Else
      DoCmd.OpenForm "RecebimentoAutomaticoPrinc", acNormal
   End If


End Function


Function SelecionaFormCTRECXBalan() As Long
   Call Liquida_Tit_Vencto_Aut_CTREC(Date)
   If Nroserie() = 111999 Or Nroserie() = 12007 Then
      DoCmd.OpenForm "CadastraCTREC", acNormal
   Else
      DoCmd.OpenForm "ManutencaoCTRECPrinc", acNormal
   End If

End Function

Function SelecionaFormBaixaCTPAGBalan() As Long
   
  ' If Nroserie() = 111999 Or Nroserie() = 12007 Then
  '    DoCmd.OpenForm "EfetuarPagamentosPrinc", acNormal
  ' Else
      DoCmd.OpenForm "PagamentosAutomaticoPrinc", acNormal
  ' End If

End Function

Function SelecionaFormCTPAGXBalan() As Long
   Call Liquida_Tit_Vencto_Aut_CTPAG(Date)
  ' If Nroserie() = 111999 Or Nroserie() = 12007 Then
  '    DoCmd.OpenForm "ManutContasPagar", acNormal
  ' Else
      DoCmd.OpenForm "ManutencaoCTPAGPrinc", acNormal
  ' End If

End Function

Function MascaraCelular(Fone As Variant)
Dim Masctel As String
If Not IsNull(Fone) Then
   Masctel = "(" & Mid$(Fone, 1, 2) & ")" & Mid$(Fone, 3, 5) & "-" & Mid$(Fone, 8, 4)
'11991834985
Else
   Masctel = Space(1)
End If
MascaraCelular = Masctel
End Function

Function Mascarafone(Fone As Variant)
Dim Masctel As String
'31998914961
If Not IsNull(Fone) Then
   If Len(Fone) = 11 Then
      Masctel = MascaraCelular(Fone)
   Else
      Masctel = "(" & Mid$(Fone, 1, 2) & ")" & Mid$(Fone, 3, 4) & "-" & Mid$(Fone, 7, 4)
   End If
Else
   Masctel = Space(1)
End If
Mascarafone = Masctel
End Function

Function DesmarcaOSVSelecionadas() As Long
On Error GoTo TratamentodeErro:
Dim db As DataBase, Ret As Variant, X As Long, numreg As Long
Dim QOSV As QueryDef, OSV As Recordset
Set db = CurrentDb()
Set QOSV = db.QueryDefs("OSVSELECIONADAS")
Set OSV = QOSV.OpenRecordset()
BeginTrans
If OSV.RecordCount <> 0 Then
   While Not OSV.EOF
        If OSV.EOF Then
           GoSub Finaliza:
        End If
        OSV.Edit
        OSV!SelecionaOSV = 0
        OSV.Update
        OSV.MoveNext
   Wend
End If

Finaliza:
 CommitTrans
 OSV.Close
 db.Close
 Set OSV = Nothing
   Set db = Nothing
 Exit Function
Return

TratamentodeErro:
   If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaTotalNotaFiscal(WNum As Long) As Double
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim NF, Est, serv As Recordset
Dim WVALORITEMEST As Double
Dim WTotalLiqServ As Double
Dim WTotalBrutoServ As Double
Dim WTotalDescServ As Double
Dim WTotalliqEst As Double
Dim WTotalBrutoEst As Double
Dim WTotalDescEst As Double

Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set NF = bd1.OpenRecordset("NOTAFISCAL", DB_OPEN_TABLE)
Set Est = bd1.OpenRecordset("ITENSESTNOTAFISCAL", DB_OPEN_TABLE)
Set serv = bd1.OpenRecordset("ITENSSERVNOTAFISCAL", DB_OPEN_TABLE)
Est.index = "TerceiraChave"
serv.index = "TerceiraChave"

NF.index = "TerceiraChave"

NF.Seek "=", WNum

If Not NF.NoMatch Then
   While NF!NroNotaFiscal = WNum
   
         If NF!CodEmpresa = Form_Menu.WEmpresa Then
            GoSub ApuraServicos:
            GoSub ApuraItensEstoque:
            CapturaTotalNotaFiscal = WTotalliqEst + WTotalLiqServ
         End If
         
         NF.MoveNext
         If NF.EOF Then
            GoSub Finaliza:
         End If
          
   Wend
Else
   CapturaTotalNotaFiscal = 0
End If

Finaliza:
 NF.Close
 Est.Close
 serv.Close
 bd1.Close
 Set NF = Nothing
 Set Est = Nothing
 Set serv = Nothing
 Set bd1 = Nothing
 Exit Function
Return


ApuraItensEstoque:
   Est.Seek "=", NF!NroSeq
   If Est.NoMatch Then
      Return
   Else
      Do Until Est.EOF
         If Est!NroSeq <> NF!NroSeq Then
            Exit Do
         Else
            If NF!DECRETO43349 = "1" Then
               WVALORITEMEST = (Est!ValorUnitario - Est!Desconto) - Round((Est!ValorUnitario - Est!Desconto) * NF!AliquotaICMS)
               WTotalliqEst = WTotalliqEst + (WVALORITEMEST * Est!Quantidade)
            Else
               WTotalliqEst = WTotalliqEst + ((Est!ValorUnitario - Est!Desconto) * Est!Quantidade)
            End If
            Est.MoveNext
         End If
      Loop
   End If
Return

ApuraServicos:
   serv.Seek "=", NF!NroSeq
   If serv.NoMatch Then
      Return
   Else
      Do Until serv.EOF
         If serv!NroSeq <> NF!NroSeq Then
            Exit Do
         Else
            WTotalLiqServ = WTotalLiqServ + ((serv!ValorMaoObra - serv!Desconto) * serv!Quantidade)
            WTotalBrutoServ = WTotalBrutoServ + (serv!ValorMaoObra * serv!Quantidade)
            WTotalDescServ = WTotalDescServ + (serv!Desconto * serv!Quantidade)
            serv.MoveNext
         End If
      Loop
   End If
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function ValidaItemEstoqueXNFXPeriodo(WCodItem As Long, WDataInicial As Variant, WDataFinal As Variant) As Long
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim NF, Est, serv As Recordset
Dim WQuant As Double

Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set NF = bd1.OpenRecordset("NOTAFISCAL", DB_OPEN_TABLE)
Set Est = bd1.OpenRecordset("ITENSESTNOTAFISCAL", DB_OPEN_TABLE)
Est.index = "CodigoItem"
NF.index = "PrimaryKey"
WQuant = 0

Est.Seek "=", WCodItem

If Not Est.NoMatch Then
   While Not Est.EOF
   
         If Est!CodigoItem <> WCodItem Then
            GoSub Finaliza:
         End If
         If Est!CodEmpresa = Form_Menu.WEmpresa Then
         
            NF.Seek "=", Est!NroSeq
            If Not NF.NoMatch Then
         
               If NF!DataEmissao >= WDataInicial Then
                  If NF!DataEmissao <= WDataFinal Then
                     WQuant = WQuant + Est!Quantidade
                  End If
               End If
            End If
         End If
         Est.MoveNext
   Wend
Else
   WQuant = 0
End If

Finaliza:
 ValidaItemEstoqueXNFXPeriodo = WQuant
 NF.Close
 Est.Close
 bd1.Close
 Set NF = Nothing
 Set Est = Nothing
 Set bd1 = Nothing
 Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CriaAquisicaoSelos(WDataEntrada As Variant, WDataDocumento As Variant, WFornec As Long, WTipo As String, wNroDocumento As Long) As Long

On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim tb  As Recordset
Dim WUlt As Long
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd1.OpenRecordset("ENTRADASSELOS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

BeginTrans
tb.Seek "=", WFornec, wNroDocumento, WDataEntrada, WDataDocumento, WTipo
If tb.NoMatch Then
   tb.AddNew
Else
   tb.Edit
End If

tb!CodigoFornecedor = WFornec
tb!NroDocumento = Format$(wNroDocumento, "0000000")
tb!DataEntrada = WDataEntrada
tb!DataDocumento = WDataDocumento
tb!TipoSelos = WTipo
tb!CodEmpresa = Form_Menu.WEmpresa
tb.Update

Finaliza:
  CommitTrans
  tb.Close
  bd1.Close
  Set tb = Nothing
  Set bd1 = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function PrimDiaMes(WDataRef) As Variant
PrimDiaMes = DateValue(Format$((DateSerial(Year(WDataRef), Month(WDataRef), 1)), "dd/mm/yyyy"))
'Stop
End Function
Function PrimDiaMesCor()
PrimDiaMesCor = DateSerial(Year(Now), Month(Now), 1)
End Function
Function UltDiaMesCor()
UltDiaMesCor = DateSerial(Year(Now), Month(Now) + 1, 1) - 1
End Function
Function UltDiaMesMMAA(Wmes As Long, WAno As Double) As Variant
'UltDiaMesMMAA = DateSerial(Year(WAno), Month(WMes) + 1, 1) - 1
UltDiaMesMMAA = Format$(DateSerial(Year(WAno), Month(Wmes) + 1, 1) - 1, "dd/mm/yyyy")
'Stop
End Function

Function N_Round(Pvar As Variant) As Currency

Dim cont As Long
If IsNull(Pvar) Then
   N_Round = 0
Else
   For cont = 1 To Len(Pvar) Step 1
       If Mid$(Pvar, cont, 1) = "," Then
          Exit Function
       End If
       N_Round = N_Round & Mid$(Pvar, cont, 1)
   Next cont
   
   'N_Round = Int((CCur(Pvar) + 0.005) * 100) / 100
   'N_Round = (CCur(Pvar) / 100)
End If

End Function
 


Function Captura_Descricao_Depto_Resp(WCodigoDepto As Long) As String
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("DEPTORESP", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
'BeginTrans
tb.Seek "=", WCodigoDepto
If Not tb.NoMatch Then
   If IsNull(tb!Descricao) Then
       Captura_Descricao_Depto_Resp = "ERRO DESCRICAO DEPTO"
    Else
     Captura_Descricao_Depto_Resp = UCase(tb!Descricao)
   End If
End If
'CommitTrans
tb.Close
bd.Close
Set tb = Nothing
Set bd = Nothing
Exit Function

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function CapturaDescricaoPrazo(WPrazo As Long) As String
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim PRZ As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set PRZ = bd1.OpenRecordset("CONDPAGTO", DB_OPEN_TABLE)
PRZ.index = "PrimaryKey"
'BeginTrans
PRZ.Seek "=", WPrazo
If Not PRZ.NoMatch Then
   If IsNull(PRZ!Descricao) Then
       CapturaDescricaoPrazo = "ERRO PRAZO"
    Else
     CapturaDescricaoPrazo = Mid$(PRZ!Descricao, 1, 16)
   End If
End If
'CommitTrans
PRZ.Close
bd1.Close
Set PRZ = Nothing
Set bd1 = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If



End Function

Function CapturaPrazoXMovtaEstoque(WCodigoPagto As Long) As Boolean

On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("CONDPAGTO", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", WCodigoPagto
If Not tb.NoMatch Then
   If tb!GeraEstoque Then
      CapturaPrazoXMovtaEstoque = True
   Else
      CapturaPrazoXMovtaEstoque = False
  
   End If
       
End If
tb.Close
bd.Close
Set tb = Nothing
Set bd = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function CapturaCGCCPF(Numos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim OSV, Cli As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set OSV = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
OSV.index = "PrimaryKey"
Cli.index = "PrimaryKey"
'BeginTrans
OSV.Seek "=", Numos
If Not OSV.NoMatch Then
   Cli.Seek "=", OSV!CodigoCliente
   If Not Cli.NoMatch Then
      CapturaCGCCPF = MascaraCgcCpf(Cli!CgcCpf, Cli!Tipo)
   End If
End If
'CommitTrans
Cli.Close
OSV.Close
bd1.Close
Set Cli = Nothing
Set OSV = Nothing
Set bd1 = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function CapturaEnderecoCliente(Numos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim OSV, Cli As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set OSV = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
OSV.index = "PrimaryKey"
Cli.index = "PrimaryKey"
'BeginTrans
OSV.Seek "=", Numos
If Not OSV.NoMatch Then
   Cli.Seek "=", OSV!CodigoCliente
   If Not Cli.NoMatch Then
      If IsNull(Cli!Endereco) Then
          CapturaEnderecoCliente = "NAO FORNECIDO"
      Else
          CapturaEnderecoCliente = Cli!Endereco
      End If
   End If
End If
'CommitTrans
Cli.Close
OSV.Close
bd1.Close
Set Cli = Nothing
Set OSV = Nothing
Set bd1 = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function CapturaLocalizacaoItemEstoque(WCodItem As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim Est As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Est = bd1.OpenRecordset("EMPXEST", DB_OPEN_TABLE)
Est.index = "PrimaryKey"
'BeginTrans
Est.Seek "=", Form_Menu.WEmpresa, WCodItem
If Not Est.NoMatch Then
   If Not IsNull(Est!Localizacao) Then
      CapturaLocalizacaoItemEstoque = Est!Localizacao
   End If
Else
   CapturaLocalizacaoItemEstoque = "ST"
End If
'CommitTrans
Est.Close
bd1.Close
Set Est = Nothing
Set bd1 = Nothing

Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaNroCliente(Numos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim OSV, Cli As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set OSV = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
OSV.index = "PrimaryKey"
Cli.index = "PrimaryKey"
'BeginTrans
OSV.Seek "=", Numos
If Not OSV.NoMatch Then
   Cli.Seek "=", OSV!CodigoCliente
   If Not Cli.NoMatch Then
      If IsNull(Cli!Nro) Then
          CapturaNroCliente = "0"
      Else
          CapturaNroCliente = Cli!Nro
      End If
   End If
End If
'CommitTrans
Cli.Close
OSV.Close
bd1.Close
Set Cli = Nothing
Set OSV = Nothing
Set bd1 = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaComplCliente(Numos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim OSV, Cli As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set OSV = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
OSV.index = "PrimaryKey"
Cli.index = "PrimaryKey"
'BeginTrans
OSV.Seek "=", Numos
If Not OSV.NoMatch Then
   Cli.Seek "=", OSV!CodigoCliente
   If Not Cli.NoMatch Then
      If IsNull(Cli!Complemento) Then
          CapturaComplCliente = Space(1)
      Else
          CapturaComplCliente = Cli!Complemento
      End If
   End If
End If
'CommitTrans
Cli.Close
OSV.Close
bd1.Close
Set Cli = Nothing
Set OSV = Nothing
Set bd1 = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaBairroCliente(Numos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim OSV, Cli As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set OSV = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
OSV.index = "PrimaryKey"
Cli.index = "PrimaryKey"
'BeginTrans
OSV.Seek "=", Numos
If Not OSV.NoMatch Then
   Cli.Seek "=", OSV!CodigoCliente
   If Not Cli.NoMatch Then
      If IsNull(Cli!Endereco) Then
          CapturaBairroCliente = "NAO FORNECIDO"
      Else
          If IsNull(Cli!Bairro) Then
             CapturaBairroCliente = Space(1)
          Else
             CapturaBairroCliente = Cli!Bairro
          End If
      End If
   End If
End If
'CommitTrans
Cli.Close
OSV.Close
bd1.Close
Set Cli = Nothing
Set OSV = Nothing
Set bd1 = Nothing

Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaVendedorCliente(WCodigoCliente As Long) As Long
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("CLI", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", WCodigoCliente
If Not tb.NoMatch Then
   If IsNull(tb!CodigoVendedor) Then
      CapturaVendedorCliente = 0
   Else
      CapturaVendedorCliente = tb!CodigoVendedor
   End If
Else
   CapturaVendedorCliente = 0
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaNroEquipamentos(WNumos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim NREQ As Recordset
Dim WAux  As String

Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set NREQ = bd1.OpenRecordset("TACOGRAFOSOS", DB_OPEN_TABLE)
NREQ.index = "SecondKey"

'BeginTrans
NREQ.Seek "=", WNumos
If NREQ.NoMatch Then
   CapturaNroEquipamentos = Space(1)
   GoSub FecharTabelas:
Else
   Do Until NREQ.EOF
      If WNumos <> NREQ!Numos Then
         Exit Do
      End If
      WAux = WAux & NREQ!NumeroTacografo & "-"
      NREQ.MoveNext
   Loop
   CapturaNroEquipamentos = WAux
End If

FecharTabelas:
 '  CommitTrans
   NREQ.Close
   bd1.Close
   Set NREQ = Nothing
   Set bd1 = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaPlacaVeicXOSV(WNumos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset

Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

tb.Seek "=", WNumos
If tb.NoMatch Then
   CapturaPlacaVeicXOSV = Space(1)
Else
   If IsNull(tb!PlacaVeiculo) Then
      CapturaPlacaVeicXOSV = Space(1)
   Else
      CapturaPlacaVeicXOSV = tb!PlacaVeiculo
   End If
End If

FecharTabelas:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaChequesXOSV(WNumos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim NREQ As Recordset
Dim WAux  As String

Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set NREQ = bd1.OpenRecordset("CHEQUESPREDATADOS", DB_OPEN_TABLE)
NREQ.index = "SecondKey"

'BeginTrans
NREQ.Seek "=", WNumos
If NREQ.NoMatch Then
   CapturaChequesXOSV = Space(1)
   GoSub FecharTabelas:
Else
   Do Until NREQ.EOF
      If WNumos <> NREQ!Numos Then
         Exit Do
      End If
      WAux = WAux & " (" & NREQ!NumeroCheque & "-" & Format$(NREQ!DataVencto, "dd/mm/yyyy") & "-" & Format$(NREQ!Valor, "###0.00") & ")"
      NREQ.MoveNext
   Loop
   CapturaChequesXOSV = WAux
End If

FecharTabelas:
 '  CommitTrans
   NREQ.Close
   bd1.Close
   Set NREQ = Nothing
   Set bd1 = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaEmitxChequesXOSV(WNumos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim NREQ As Recordset
Dim WAux  As String

Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set NREQ = bd1.OpenRecordset("CHEQUESPREDATADOS", DB_OPEN_TABLE)
NREQ.index = "SecondKey"

'BeginTrans
NREQ.Seek "=", WNumos
If NREQ.NoMatch Then
   CapturaEmitxChequesXOSV = Space(1)
   GoSub FecharTabelas:
Else
   Do Until NREQ.EOF
      If WNumos <> NREQ!Numos Then
         Exit Do
      End If
      If Not IsNull(NREQ!Emitente) Then
         WAux = WAux & "(" & UCase(NREQ!Emitente) & ")"
      Else
         WAux = Space(1)
      End If
      Exit Do
   Loop
   CapturaEmitxChequesXOSV = WAux
End If

FecharTabelas:
 '  CommitTrans
   NREQ.Close
   bd1.Close
   Set NREQ = Nothing
   Set bd1 = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Total_Orcamentos_Agrupado(WCodigoCliente As Long) As Double
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset

Dim WTotal As Double
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "Secondkey"

tb.Seek "=", WCodigoCliente
If tb.NoMatch Then
   WTotal = 0
Else
   While tb!CodigoCliente = WCodigoCliente
         If tb!SelecionaOSV Then
            WTotal = WTotal + (CalculoMercadorias(tb!Numos) + TotalServicosOrdemServico(tb!Numos))
         End If
         tb.MoveNext
         If tb.EOF Then
            GoSub Finaliza:
        End If
   Wend
End If
Finaliza:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
   Total_Orcamentos_Agrupado = WTotal
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaEstadoCliente(WCodigoCliente As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim OSV, Cli As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set OSV = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
OSV.index = "PrimaryKey"
Cli.index = "PrimaryKey"
'BeginTrans
Cli.Seek "=", WCodigoCliente
If Not Cli.NoMatch Then
   If IsNull(Cli!Estado) Then
      CapturaEstadoCliente = Estado()
   Else
      CapturaEstadoCliente = Cli!Estado
   End If
End If
'CommitTrans
Cli.Close
OSV.Close
bd1.Close
Set Cli = Nothing
Set OSV = Nothing
Set bd1 = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function Valida_CliXBaseLote(wCodCliente As Long) As Boolean
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Base_Dados_Lote", DB_OPEN_TABLE)
tb.index = "CodigoCliente"

Valida_CliXBaseLote = True

tb.Seek "=", wCodCliente
If tb.NoMatch Then
   Valida_CliXBaseLote = False
End If

Finaliza:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
   Exit Function
Return


TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Captura_Enviar_Base_Dados_Lote(wCodCliente As Long) As Boolean
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Base_Dados_Lote", DB_OPEN_TABLE)
tb.index = "CodigoCliente"
tb.Seek "=", wCodCliente
Captura_Enviar_Base_Dados_Lote = False
If tb.NoMatch Then GoSub Finaliza
Do While tb!CodigoCliente = wCodCliente
    If tb!Enviar = True Then
        Captura_Enviar_Base_Dados_Lote = True
        GoTo Finaliza
    End If
    tb.MoveNext
    If tb.EOF = True Then
       Exit Do
    End If
Loop
Finaliza:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
   Exit Function
Return
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function CapturaUFCadastroCliente(wCodCliente As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim Cli As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
Cli.index = "PrimaryKey"
'BeginTrans
Cli.Seek "=", wCodCliente
If Cli.NoMatch Then
   MsgBox "Cliente inexistente...", vbCritical, "SGa"
Else
   CapturaUFCadastroCliente = Cli!Estado
End If
'CommitTrans
Cli.Close
bd1.Close
Set Cli = Nothing
Set bd1 = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

'Function CriaBairro(WBairro As String) As Long
'On Error GoTo TratamentodeErro:

'Dim bd As DataBase
'Dim tb As Recordset
'Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
'Set tb = bd.OpenRecordset("BAI", DB_OPEN_TABLE)
'tb.Index = "PrimaryKey"
'BeginTrans

'tb.Seek "=", WBairro
'If tb.NoMatch Then
'   tb.AddNew
'Else
'   tb.Edit
'End If
'tb!DescricaoBAirro = WBairro
'tb!CodigoCEP = "00000000"
'tb.Update

'Finaliza:
'  C ommitTrans
'  tb.Close
'  bd.Close
'  Set tb = Nothing
'  Set bd = Nothing
'  Exit Function
'Return
'TratamentodeErro:
'    If Err.Number = 3001 Then
'      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
'    Else
'       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description, Vbcritical,"SGa"
'       Resume Next
'    End If


'End Function

Function DesmarcaCFOP() As Long
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim cfo As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set cfo = bd1.OpenRecordset("IDENTXCFOP", DB_OPEN_TABLE)
'CFO.Index = "PrimaryKey"
BeginTrans
If cfo.RecordCount <> 0 Then
   While Not cfo.EOF
         cfo.Edit
         cfo!Selecionar = 0
         cfo.Update
         cfo.MoveNext
   Wend
   
End If

Finaliza:
  CommitTrans
  cfo.Close
  bd1.Close
  Set cfo = Nothing
  Set bd1 = Nothing
  Exit Function
Return
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function SelecionaPlacXCliXConec() As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim Qtb As QueryDef
Dim tb As Recordset
Set bd = CurrentDb()
Set Qtb = bd.QueryDefs("Plac_Cli_Conec")
DoCmd.Hourglass True

Set tb = Qtb.OpenRecordset
If tb.RecordCount <> 0 Then
   While Not tb.EOF
         tb.Edit
         tb!Selecionar = 1
         tb.Update
         tb.MoveNext
   Wend
End If

Finaliza:
      DoCmd.Hourglass False
      tb.Close
      bd.Close
      Set tb = Nothing
      Set bd = Nothing
      Exit Function
Return
    
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function SelecionarCFOP() As Long
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim cfo As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set cfo = bd1.OpenRecordset("IDENTXCFOP", DB_OPEN_TABLE)
'CFO.Index = "PrimaryKey"
'BeginTrans
If cfo.RecordCount <> 0 Then
   While Not cfo.EOF
         cfo.Edit
         cfo!Selecionar = 1
         cfo.Update
         cfo.MoveNext
   Wend
   
End If

Finaliza:
  'CommitTrans
  cfo.Close
  bd1.Close
  Set cfo = Nothing
  Set bd1 = Nothing
  Exit Function
Return
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaCEPCliente(Numos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim OSV, Cli As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set OSV = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
OSV.index = "PrimaryKey"
Cli.index = "PrimaryKey"
'BeginTrans
OSV.Seek "=", Numos
If Not OSV.NoMatch Then
   Cli.Seek "=", OSV!CodigoCliente
   If Not Cli.NoMatch Then
      If IsNull(Cli!Endereco) Then
          CapturaCEPCliente = "NAO FORNECIDO"
      Else
          CapturaCEPCliente = Cli!EndPostal
      End If
   End If
End If
'CommitTrans
Cli.Close
OSV.Close
bd1.Close
Set Cli = Nothing
Set OSV = Nothing
Set bd1 = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function


Function CapturaNumeroEquipamentoOSV(WNumos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim OSV As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set OSV = bd1.OpenRecordset("TACOGRAFOSOS", DB_OPEN_TABLE)
OSV.index = "SecondKey"
'BeginTrans
OSV.Seek "=", WNumos
If OSV.NoMatch Then
   CapturaNumeroEquipamentoOSV = Space(1)
Else
   CapturaNumeroEquipamentoOSV = Mid$(OSV!NumeroTacografo, 1, 22)
End If
'CommitTrans
OSV.Close
bd1.Close
Set OSV = Nothing
Set bd1 = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    
    End If

End Function

Function CapturaObservacoesOSV(Numos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim OSV As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set OSV = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
OSV.index = "PrimaryKey"
'BeginTrans
OSV.Seek "=", Numos
If Not OSV.NoMatch Then
   If Not IsNull(OSV!Observacoes) Then
      CapturaObservacoesOSV = OSV!Observacoes
   Else
      CapturaObservacoesOSV = Space(1)
   End If
Else
  CapturaObservacoesOSV = Space(1)
End If
'CommitTrans
OSV.Close
bd1.Close
Set OSV = Nothing
Set bd1 = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaDataAberturaOSV(Numos As Long) As Variant
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim OSV As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set OSV = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
OSV.index = "PrimaryKey"
'BeginTrans
OSV.Seek "=", Numos
If Not OSV.NoMatch Then
   CapturaDataAberturaOSV = OSV!DataAbertura
Else
   CapturaDataAberturaOSV = "Abertura invalida..."
End If
'CommitTrans
OSV.Close
bd1.Close
Set OSV = Nothing
Set bd1 = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function CapturaNroRequisicaoOSV(Numos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim OSV As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set OSV = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
OSV.index = "PrimaryKey"
'BeginTrans
OSV.Seek "=", Numos
If Not OSV.NoMatch Then
   If Not IsNull(OSV!NroRequisicao) Then
      CapturaNroRequisicaoOSV = OSV!NroRequisicao
   End If
Else
   CapturaNroRequisicaoOSV = "XXX"
End If
'CommitTrans
OSV.Close
bd1.Close
Set OSV = Nothing
Set bd1 = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If



End Function

Function CapturaHoraAberturaOSv(Numos As Long) As Variant
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim OSV As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set OSV = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
OSV.index = "PrimaryKey"
'BeginTrans
OSV.Seek "=", Numos
If Not OSV.NoMatch Then
   CapturaHoraAberturaOSv = OSV!HoraAbertura
Else
   CapturaHoraAberturaOSv = "Hora invalida..."
End If
'CommitTrans
OSV.Close
bd1.Close
Set OSV = Nothing
Set bd1 = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaHoraFechamentoOSv(Numos As Long) As Variant
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim OSV As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set OSV = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
OSV.index = "PrimaryKey"
'BeginTrans
OSV.Seek "=", Numos
If Not OSV.NoMatch Then
   CapturaHoraFechamentoOSv = OSV!HoraFechamento
Else
   CapturaHoraFechamentoOSv = "Hora invalida..."
End If
'CommitTrans
OSV.Close
bd1.Close
Set OSV = Nothing
Set bd1 = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaDataFechamentoOSV(Numos As Long) As Variant
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim OSV As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set OSV = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
OSV.index = "PrimaryKey"
'BeginTrans
OSV.Seek "=", Numos
If Not OSV.NoMatch Then
   CapturaDataFechamentoOSV = OSV!datafechamento
Else
   CapturaDataFechamentoOSV = "Fechamento invalida..."
End If
'CommitTrans
OSV.Close
bd1.Close
Set OSV = Nothing
Set bd1 = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaResponsavelOSV(Numos As Long) As String
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim OSV, Resp As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set OSV = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
Set Resp = bd1.OpenRecordset("RESP", DB_OPEN_TABLE)
OSV.index = "PrimaryKey"
Resp.index = "PrimaryKey"
'BeginTrans
OSV.Seek "=", Numos
If Not OSV.NoMatch Then
   Resp.Seek "=", OSV!Responsavel
   If Not Resp.NoMatch Then
      If IsNull(Resp!Nome) Then
          CapturaResponsavelOSV = "NAO FORNECIDO"
      Else
          CapturaResponsavelOSV = Resp!Nome
      End If
   End If
End If
'CommitTrans
Resp.Close
OSV.Close
bd1.Close
Set Resp = Nothing
Set OSV = Nothing
Set bd1 = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function
Function CapturaCodIBGEXCliente(WCodigo As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim CID, Cli As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set CID = bd1.OpenRecordset("CID", DB_OPEN_TABLE)
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
CID.index = "PrimaryKey"
Cli.index = "PrimaryKey"
'BeginTrans
Cli.Seek "=", WCodigo
If Not Cli.NoMatch Then
   CID.Seek "=", Cli!Cidade
   If CID.NoMatch Then
      CapturaCodIBGEXCliente = 0
   Else
      If Cli!CodigoPais = 1058 Then
         If IsNull(CID!CodigoIBGE) Then
            CapturaCodIBGEXCliente = 0
         Else
            CapturaCodIBGEXCliente = CID!CodigoIBGE
         End If
      Else
        CapturaCodIBGEXCliente = 9999999
      End If
   End If
End If
'CommitTrans
Cli.Close
CID.Close
bd1.Close
Set Cli = Nothing
Set CID = Nothing
Set bd1 = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaCodMunicNFSENetXCliente(WCodigo As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim CID, Cli As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set CID = bd1.OpenRecordset("CID", DB_OPEN_TABLE)
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
CID.index = "PrimaryKey"
Cli.index = "PrimaryKey"
'BeginTrans
Cli.Seek "=", WCodigo
If Not Cli.NoMatch Then
   CID.Seek "=", Cli!Cidade
   If CID.NoMatch Then
      CapturaCodMunicNFSENetXCliente = 0
   Else
      If Cli!CodigoPais = 1058 Then
         If IsNull(CID!CodMunicNFSe) Then
            CapturaCodMunicNFSENetXCliente = 0
         Else
            CapturaCodMunicNFSENetXCliente = Format$(Val(CID!CodMunicNFSe), "000000000000000")
         End If
      Else
        CapturaCodMunicNFSENetXCliente = 9999999
      End If
   End If
End If
'CommitTrans
Cli.Close
CID.Close
bd1.Close
Set Cli = Nothing
Set CID = Nothing
Set bd1 = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaCidadeCliente(Numos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim OSV, Cli As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set OSV = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
OSV.index = "PrimaryKey"
Cli.index = "PrimaryKey"
'BeginTrans
OSV.Seek "=", Numos
If Not OSV.NoMatch Then
   Cli.Seek "=", OSV!CodigoCliente
   If Not Cli.NoMatch Then
      If IsNull(Cli!Endereco) Then
          CapturaCidadeCliente = "NAO FORNECIDO"
      Else
          If IsNull(Cli!Cidade) Then
              CapturaCidadeCliente = Space(1)
          Else
             CapturaCidadeCliente = Cli!Cidade
          End If
      End If
   End If
End If
'CommitTrans
Cli.Close
OSV.Close
bd1.Close
Set Cli = Nothing
Set OSV = Nothing
Set bd1 = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function CapturaTipoCliente(wCodCliente As Long) As Long
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim Cli As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
Cli.index = "PrimaryKey"
'BeginTrans
Cli.Seek "=", wCodCliente
If Not Cli.NoMatch Then
   If IsNull(Cli!Tipo) Then
       CapturaTipoCliente = 1
   Else
       CapturaTipoCliente = Cli!Tipo
   End If
End If

Finaliza:
 '  CommitTrans
    Cli.Close
    bd1.Close
    Set Cli = Nothing
    Set bd1 = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function CapturaMarcaEnsaio() As String

On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("EMPRESA", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
'BeginTrans

If tb.RecordCount <> 0 Then
   While Not tb.EOF
         If tb!CodigoEmpresa = Form_Menu.WEmpresa Then
            If Not IsNull(tb!MarcaEns) Then
               CapturaMarcaEnsaio = tb!MarcaEns
            Else
               CapturaMarcaEnsaio = "XXX"
            End If
        End If
        tb.MoveNext
   Wend
 
End If

Finaliza:
 '  CommitTrans
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaNotaFiscalEmitida(WNumos As Long) As Long
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim OSV As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set OSV = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
OSV.index = "PrimaryKey"
'BeginTrans
OSV.Seek "=", WNumos
If Not OSV.NoMatch Then
   
   If Not OSV!NFEmitida And Not OSV!NFSEmitida Then
      CapturaNotaFiscalEmitida = 0
   ElseIf OSV!NFEmitida Then
          CapturaNotaFiscalEmitida = 1 ' NF PECA EMITIDA
   ElseIf OSV!NFSEmitida Then
          CapturaNotaFiscalEmitida = 2 'NF  SERVICO EMITIDA
   ElseIf OSV!NFSEmitida And OSV!NFEmitida Then
          CapturaNotaFiscalEmitida = 3 'NFS  PECA E SERVICO EMITIDA
   End If
End If
'CommitTrans
OSV.Close
bd1.Close
Set OSV = Nothing
Set bd1 = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaBairroCli(WCodigoCliente As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim Cli As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
Cli.index = "PrimaryKey"
Cli.Seek "=", WCodigoCliente
If Not Cli.NoMatch Then
   If IsNull(Cli!Bairro) Then
      CapturaBairroCli = "NAO FORNECIDO"
   Else
      CapturaBairroCli = Cli!Bairro
   End If
End If

Cli.Close
bd1.Close
Set Cli = Nothing
Set bd1 = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function
Function CapturaCaminhoImagVinc1(WNumos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", WNumos
If Not tb.NoMatch Then
   If IsNull(tb!ImagVinc1) Then
      CapturaCaminhoImagVinc1 = "X"
   Else
      CapturaCaminhoImagVinc1 = tb!ImagVinc1
   End If
End If
tb.Close
bd.Close
Set tb = Nothing
Set bd = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function
Function CapturaCaminhoImagVinc(WNumos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", WNumos
If Not tb.NoMatch Then
   If IsNull(tb!ImagVinc) Then
      CapturaCaminhoImagVinc = "X"
   Else
      CapturaCaminhoImagVinc = tb!ImagVinc
   End If
End If

tb.Close
bd.Close
Set tb = Nothing
Set bd = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Captura_Imag_Item_Estoque(WCodItem As Long) As String
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("EST", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", WCodItem
If Not tb.NoMatch Then
   If IsNull(tb!ImagVinc) Then
      Captura_Imag_Item_Estoque = "X"
   Else
      Captura_Imag_Item_Estoque = tb!ImagVinc
   End If
End If

tb.Close
bd.Close
Set tb = Nothing
Set bd = Nothing

Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Cli_Ativo(WCodigoCliente As Long) As Boolean
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim Cli As Recordset
Dim WDDD As String
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
Cli.index = "PrimaryKey"
Cli.Seek "=", WCodigoCliente
If Not Cli.NoMatch Then
   If Cli!Ativo Then
      Cli_Ativo = True
   Else
      Cli_Ativo = False
   End If
End If

Cli.Close
bd1.Close
Set Cli = Nothing
Set bd1 = Nothing

Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function CapturaTelefoneCelCli(WCodigoCliente As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim Cli As Recordset
Dim WDDD As String
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
Cli.index = "PrimaryKey"
Cli.Seek "=", WCodigoCliente
If Not Cli.NoMatch Then
   If IsNull(Cli!Celular) Then
      CapturaTelefoneCelCli = "0"
   Else
      CapturaTelefoneCelCli = Cli!Celular
   End If
End If

Cli.Close
bd1.Close
Set Cli = Nothing
Set bd1 = Nothing

Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaDDDTelefoneCli(WCodigoCliente As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim Cli As Recordset
Dim WDDD As String
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
Cli.index = "PrimaryKey"
Cli.Seek "=", WCodigoCliente
If Not Cli.NoMatch Then
   If IsNull(Cli!DDD) Then
      CapturaDDDTelefoneCli = "0"
   Else
      CapturaDDDTelefoneCli = Cli!DDD
   End If
End If

Cli.Close
bd1.Close
Set Cli = Nothing
Set bd1 = Nothing

Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaTelefoneCli(WCodigoCliente As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim Cli As Recordset
Dim WDDD As String
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
Cli.index = "PrimaryKey"
Cli.Seek "=", WCodigoCliente
If Not Cli.NoMatch Then
   If IsNull(Cli!Telefone) Then
      CapturaTelefoneCli = "0"
   Else
      If IsNull(Cli!DDD) Or Cli!DDD = "" Then
         WDDD = 0
      Else
         WDDD = Cli!DDD
      End If
      CapturaTelefoneCli = IIf(IsNull(WDDD), Cli!Telefone, WDDD & "-" & Cli!Telefone)
   End If
End If

Cli.Close
bd1.Close
Set Cli = Nothing
Set bd1 = Nothing

Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaFoneCli(WCodigoCliente As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim Cli As Recordset
Dim WDDD As String
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
Cli.index = "PrimaryKey"
Cli.Seek "=", WCodigoCliente
If Not Cli.NoMatch Then
   If IsNull(Cli!Telefone) Then
      CapturaFoneCli = "0"
   Else
      CapturaFoneCli = Cli!Telefone
   End If
End If

Cli.Close
bd1.Close
Set Cli = Nothing
Set bd1 = Nothing

Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function


Function CapturaEmailCli(WCodigoCliente As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim Cli As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
Cli.index = "PrimaryKey"
Cli.Seek "=", WCodigoCliente
If Not Cli.NoMatch Then
   If IsNull(Cli!EndElet) Then
      CapturaEmailCli = Space(1)
   Else
      CapturaEmailCli = LCase(Cli!EndElet)
   End If
End If

Cli.Close
bd1.Close
Set Cli = Nothing
Set bd1 = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaObservacoesCli(WCodigoCliente As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim Cli As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
Cli.index = "PrimaryKey"
Cli.Seek "=", WCodigoCliente
If Not Cli.NoMatch Then
   If IsNull(Cli!Observacoes) Then
      CapturaObservacoesCli = Space(1)
   Else
      CapturaObservacoesCli = Cli!Observacoes
   End If
Else
   CapturaObservacoesCli = Space(1)
End If

Cli.Close
bd1.Close
Set Cli = Nothing
Set bd1 = Nothing

Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaInscEstCli(WCodigoCliente As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim Cli As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
Cli.index = "PrimaryKey"
Cli.Seek "=", WCodigoCliente
If Not Cli.NoMatch Then
   If IsNull(Cli!Inscricao) Then
      CapturaInscEstCli = "NAO FORNECIDO"
   Else
      CapturaInscEstCli = Cli!Inscricao
   End If
End If

Cli.Close
bd1.Close
Set Cli = Nothing
Set bd1 = Nothing

Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaCNPJCli(WCodigoCliente As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim Cli As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
Cli.index = "PrimaryKey"
Cli.Seek "=", WCodigoCliente
If Not Cli.NoMatch Then
   If IsNull(Cli!CgcCpf) Then
      CapturaCNPJCli = "NAO FORNECIDO"
   Else
      CapturaCNPJCli = Cli!CgcCpf
   End If
End If

Cli.Close
bd1.Close
Set Cli = Nothing
Set bd1 = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaCodEmpresa_OSV(WNumos As Long) As Long
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", WNumos
If Not tb.NoMatch Then
   CapturaCodEmpresa_OSV = tb!CodEmpresa
Else
   CapturaCodEmpresa_OSV = Form_Menu.WEmpresa
End If

Finaliza:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing

   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaRazaoSocialCli(WCodigoCliente As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim Cli As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
Cli.index = "PrimaryKey"
Cli.Seek "=", WCodigoCliente
If Not Cli.NoMatch Then
   If IsNull(Cli!RazaoSocial) Then
      CapturaRazaoSocialCli = "NAO FORNECIDO"
   Else
      CapturaRazaoSocialCli = Cli!RazaoSocial
   End If
Else
   CapturaRazaoSocialCli = "NAO IDENTIFICADO"
End If

Cli.Close
bd1.Close
Set Cli = Nothing
Set bd1 = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaCodinomeCli(WCodigoCliente As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim Cli As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
Cli.index = "PrimaryKey"
Cli.Seek "=", WCodigoCliente
If Not Cli.NoMatch Then
   If IsNull(Cli!Codinome) Then
      CapturaCodinomeCli = "NAO FORNECIDO"
   Else
      CapturaCodinomeCli = Cli!Codinome
   End If
Else
   CapturaCodinomeCli = "INEXISTENTE"
End If

Cli.Close
bd1.Close
Set Cli = Nothing
Set bd1 = Nothing

Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function


Function CapturaRazaoSocialCli_CHP(WCodigoCliente As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim Cli As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
Cli.index = "PrimaryKey"
Cli.Seek "=", WCodigoCliente
If Not Cli.NoMatch Then
   If IsNull(Cli!RazaoSocial) Then
      CapturaRazaoSocialCli_CHP = "NAO FORNECIDO"
   Else
      CapturaRazaoSocialCli_CHP = Cli!RazaoSocial
   End If
Else
   CapturaRazaoSocialCli_CHP = Captura_Conteudo_Campos_Empresa(Form_Menu.WEmpresa, "RazaoSocial")
End If

Cli.Close
bd1.Close
Set Cli = Nothing
Set bd1 = Nothing

Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Captura_Ult_Calibr_Oximetro(WDescricao As String) As Variant
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OXIMETRO_TANQUE", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Form_Menu.WEmpresa, WDescricao
If Not tb.NoMatch Then
   If IsNull(tb!Data_Ult_Calib) Then
      Captura_Ult_Calibr_Oximetro = Date
    Else
      Captura_Ult_Calibr_Oximetro = DateValue(tb!Data_Ult_Calib)
    End If
Else
   Captura_Ult_Calibr_Oximetro = Date
End If

tb.Close
bd.Close
Set tb = Nothing
Set bd = Nothing

Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Captura_Reg_Tec_Aplic(WDescricao As String) As String
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("EXPLOSIMETRO_TANQUE", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Form_Menu.WEmpresa, WDescricao
If Not tb.NoMatch Then
   If IsNull(tb!Reg_Tec_Aplic) Then
      Captura_Reg_Tec_Aplic = "INDEFINIDO"
    Else
      Captura_Reg_Tec_Aplic = tb!Reg_Tec_Aplic
    End If
Else
   Captura_Reg_Tec_Aplic = "INEXISTENTE"
End If

tb.Close
bd.Close
Set tb = Nothing
Set bd = Nothing

Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Captura_Ult_Calib_Explosimetro(WDescricao As String) As Variant
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("EXPLOSIMETRO_TANQUE", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Form_Menu.WEmpresa, WDescricao
If Not tb.NoMatch Then
   If IsNull(tb!Data_Ult_Calib) Then
      Captura_Ult_Calib_Explosimetro = Date
    Else
      Captura_Ult_Calib_Explosimetro = tb!Data_Ult_Calib
    End If
Else
   Captura_Ult_Calib_Explosimetro = Date
End If

tb.Close
bd.Close
Set tb = Nothing
Set bd = Nothing

Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function Captura_Proc_Aplic_Oximetro(WDescricao As String) As String
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OXIMETRO_TANQUE", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Form_Menu.WEmpresa, WDescricao
If Not tb.NoMatch Then
   If IsNull(tb!Proc_Aplic) Then
      Captura_Proc_Aplic_Oximetro = "INDEFINIDO"
    Else
      Captura_Proc_Aplic_Oximetro = tb!Proc_Aplic
    End If
Else
   Captura_Proc_Aplic_Oximetro = "INEXISTENTE"
End If

tb.Close
bd.Close
Set tb = Nothing
Set bd = Nothing

Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaCodigoClienteXPlaca(wPlaca As String) As Long
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("PLACXCLI", DB_OPEN_TABLE)
tb.index = "SecondKey"
tb.Seek "=", wPlaca
If Not tb.NoMatch Then
   CapturaCodigoClienteXPlaca = tb!CodigoCliente
Else
   CapturaCodigoClienteXPlaca = 0
End If

tb.Close
bd.Close
Set tb = Nothing
Set bd = Nothing

Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaNroEndCli(WCodigoCliente As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim Cli As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
Cli.index = "PrimaryKey"
Cli.Seek "=", WCodigoCliente
If Not Cli.NoMatch Then
   If IsNull(Cli!Nro) Then
      CapturaNroEndCli = Space(1)
   Else
      CapturaNroEndCli = Cli!Nro
   End If
End If

Cli.Close
bd1.Close
Set Cli = Nothing
Set bd1 = Nothing

Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaEstadoCli(WCodigoCliente As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim Cli As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
Cli.index = "PrimaryKey"
Cli.Seek "=", WCodigoCliente
If Not Cli.NoMatch Then
   If IsNull(Cli!Estado) Then
      CapturaEstadoCli = Space(1)
   Else
      CapturaEstadoCli = Cli!Estado
   End If
End If

Cli.Close
bd1.Close
Set Cli = Nothing
Set bd1 = Nothing

Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Captura_Aliq_Repasse_InMetro(WEmpresa As Long, WTipoServico As Long) As Double
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("EMPXTPSERVXREPASSE", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

tb.Seek "=", WEmpresa, WTipoServico
If Not tb.NoMatch Then
   If IsNull(tb!Repasse) Then
      Captura_Aliq_Repasse_InMetro = 0
   Else
      Captura_Aliq_Repasse_InMetro = tb!Repasse
   End If
Else
   Captura_Aliq_Repasse_InMetro = 0
End If

tb.Close
bd.Close
Set tb = Nothing
Set bd = Nothing

Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function CapturaCidadeCli(WCodigoCliente As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim Cli As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
Cli.index = "PrimaryKey"
Cli.Seek "=", WCodigoCliente
If Not Cli.NoMatch Then
   If IsNull(Cli!Cidade) Then
      CapturaCidadeCli = Space(1)
   Else
      CapturaCidadeCli = Cli!Cidade
   End If
End If

Cli.Close
bd1.Close
Set Cli = Nothing
Set bd1 = Nothing

Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaEndPostalCli(WCodigoCliente As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim Cli As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
Cli.index = "PrimaryKey"
Cli.Seek "=", WCodigoCliente
If Not Cli.NoMatch Then
   If IsNull(Cli!EndPostal) Then
      CapturaEndPostalCli = "00000000"
   Else
      CapturaEndPostalCli = Cli!EndPostal
   End If
End If

Cli.Close
bd1.Close
Set Cli = Nothing
Set bd1 = Nothing

Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaEnderecoCli(WCodigoCliente As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim Cli As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
Cli.index = "PrimaryKey"
Cli.Seek "=", WCodigoCliente
If Not Cli.NoMatch Then
   If IsNull(Cli!Endereco) Then
      CapturaEnderecoCli = "NAO FORNECIDO"
   Else
      CapturaEnderecoCli = Cli!Endereco
   End If
End If

Cli.Close
bd1.Close
Set Cli = Nothing
Set bd1 = Nothing

Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaRazaoSocialCliente(Numos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim OSV, Cli As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set OSV = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
OSV.index = "PrimaryKey"
Cli.index = "PrimaryKey"
'BeginTrans
OSV.Seek "=", Numos
If Not OSV.NoMatch Then
   Cli.Seek "=", OSV!CodigoCliente
   If Not Cli.NoMatch Then
      If IsNull(Cli!RazaoSocial) Then
         CapturaRazaoSocialCliente = "NAO FORNECIDO"
      Else
         CapturaRazaoSocialCliente = Cli!RazaoSocial
      End If
   End If
End If
'CommitTrans
Cli.Close
OSV.Close
bd1.Close
Set Cli = Nothing
Set OSV = Nothing
Set bd1 = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaCOEFWXOSV(Numos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Numos
If Not tb.NoMatch Then
   If Not IsNull(tb!CoeficienteW) Then
      CapturaCOEFWXOSV = tb!CoeficienteW
   Else
      CapturaCOEFWXOSV = " "
   End If
End If

Finaliza:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing

   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function
Function AtualizaCOEFWXOSV(Numos As Long, WCoeficienteW As Long) As String
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Numos
BeginTrans
If Not tb.NoMatch Then
   tb.Edit
   tb!CoeficienteW = WCoeficienteW
   tb.Update
End If
Finaliza:
   CommitTrans
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing

   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function


Function CapturaOSVXGRU(WGRU As String) As Long
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "GRU"
tb.Seek "=", WGRU
If Not tb.NoMatch Then
   CapturaOSVXGRU = tb!Numos
Else
   CapturaOSVXGRU = 0
End If

Finaliza:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing

   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function CapturaGRUXOSV(Numos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", Numos
If Not tb.NoMatch Then
   If Not IsNull(tb!GRU) Then
      
      CapturaGRUXOSV = tb!GRU
   Else
      CapturaGRUXOSV = ""
   End If
End If

Finaliza:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing

   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaFoneCliente(Numos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim OSV, Cli As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set OSV = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
OSV.index = "PrimaryKey"
Cli.index = "PrimaryKey"
'BeginTrans
OSV.Seek "=", Numos
If Not OSV.NoMatch Then
   Cli.Seek "=", OSV!CodigoCliente
   If Not Cli.NoMatch Then
      If IsNull(Cli!Telefone) Then
         CapturaFoneCliente = 0
      Else
         CapturaFoneCliente = Cli!Telefone
      End If
   End If
End If
'CommitTrans'
Cli.Close
OSV.Close
bd1.Close
Set Cli = Nothing
Set OSV = Nothing
Set bd1 = Nothing

Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function CapturaEmailCliente(Numos As Long) As String
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim OSV, Cli As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set OSV = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
OSV.index = "PrimaryKey"
Cli.index = "PrimaryKey"
'BeginTrans
OSV.Seek "=", Numos
If Not OSV.NoMatch Then
   Cli.Seek "=", OSV!CodigoCliente
   If Not Cli.NoMatch Then
      If IsNull(Cli!EndElet) Then
         CapturaEmailCliente = Space(1)
      Else
         CapturaEmailCliente = Cli!EndElet
      End If
   End If
End If
'CommitTrans
Cli.Close
OSV.Close
bd1.Close
Set Cli = Nothing
Set OSV = Nothing
Set bd1 = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function TipoECF() As Long
On Error GoTo TratamentodeErro:
Dim emp As Recordset, bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set emp = bd1.OpenRecordset("Empresa", DB_OPEN_TABLE)

Do Until emp.EOF
   If Form_Menu.WEmpresa = emp!CodigoEmpresa Then
      TipoECF = emp!TipoECF
   End If
   emp.MoveNext
Loop
emp.Close
bd1.Close
Set emp = Nothing
Set bd1 = Nothing
   Exit Function

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function



Function CriaContatoCli(WCodigoCliente As Long, WCodigoContato As Long, WRazao As String) As Long

On Error GoTo TratamentodeErro:
Dim CNT As Recordset, bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set CNT = bd1.OpenRecordset("CNT", DB_OPEN_TABLE)
CNT.index = "PrimaryKey"
BeginTrans
CNT.Seek "=", WCodigoCliente
If CNT.NoMatch Then
   CNT.AddNew
   CNT!Nome = UCase(Mid$(WRazao, 1, 50))
   CNT!Codinome = UCase(Mid$(WRazao, 1, 20))
Else
   CNT.Edit
End If
CNT!CodigoContato = WCodigoContato
CNT!CodigoCliente = WCodigoCliente
CNT.Update

Finaliza:
  CommitTrans
  CNT.Close
  bd1.Close
  Set CNT = Nothing
  Set bd1 = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function CapturaDataNascCli(WCodigoCliente As Long) As String
On Error GoTo TratamentodeErro:
Dim CNT As Recordset, bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set CNT = bd1.OpenRecordset("CNT", DB_OPEN_TABLE)
CNT.index = "SecondKey"

CNT.Seek "=", WCodigoCliente
If CNT.NoMatch Then
   CapturaDataNascCli = "0000"
Else
   CapturaDataNascCli = CNT!Nascimento
End If

Finaliza:
  CNT.Close
  bd1.Close
  Set CNT = Nothing
  Set bd1 = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function capturaSexoCli(WCodigoCliente As Long) As String

On Error GoTo TratamentodeErro:
Dim CNT As Recordset, bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set CNT = bd1.OpenRecordset("CNT", DB_OPEN_TABLE)
CNT.index = "SecondKey"

CNT.Seek "=", WCodigoCliente
If CNT.NoMatch Then
   capturaSexoCli = "M"
Else
   capturaSexoCli = CNT!Sexo
End If

Finaliza:
  CNT.Close
  bd1.Close
  Set CNT = Nothing
  Set bd1 = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaTipoReg() As String
On Error GoTo TratamentodeErro:
Dim emp As Recordset, bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set emp = bd1.OpenRecordset("Empresa", DB_OPEN_TABLE)

Do Until emp.EOF
   If emp!CodigoEmpresa = Form_Menu.WEmpresa Then
      CapturaTipoReg = emp!TIPOREG
   End If
   emp.MoveNext
Loop
emp.Close
bd1.Close
Set emp = Nothing
Set bd1 = Nothing
   Exit Function

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function GeraDiasUteis()
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim DU As Recordset
Dim X As Long
Dim wQtde As Long
Dim WDataA As Variant, WDataB As Variant
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set DU = bd1.OpenRecordset("DIASUTEIS", DB_OPEN_TABLE)
DU.index = "PrimaryKey"
BeginTrans
WDataA = DateValue(Format$("01/01/2011", "dd/mm/yyyy"))
'Du.Seek "=", WDataA
    
'If Du.NoMatch Then
   While WDataA < #12/31/2050#
      DU.Seek "=", WDataA
      If DU.NoMatch Then
         DU.AddNew
      Else
         DU.Edit
      End If
      DU!DataDiasUteis = WDataA
      If WeekDay(WDataA) = 1 Then
         DU!Diautil = 0
      Else
         DU!Diautil = 1
      End If
      X = X + 1
      If X Mod 500 = 0 Then
         CommitTrans
         BeginTrans
      End If
      DU.Update
      WDataA = WDataA + 1
   Wend
'End If
CommitTrans
DU.Close
Set DU = Nothing
Set bd1 = Nothing

Exit Function

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function extenso(Valor As Variant)
Dim X, UNID, DZNA, CTNA As Integer
    Dim UN, DU, DZ, CT, Txt As String
    If Valor = 0 Or Valor > 999 Then    'valor � v�lido ?
        Exit Function
    End If
    If Valor = 100 Then
        extenso = "CEM"
        Exit Function
    End If
    UN = "UM    DOIS  TR�S  QUATROCINCO SEIS  SETE  OITO  NOVE"
    DZ = "DEZ      VINTE    TRINTA   QUARENTA CINQUENTASESSENTA SETENTA  OITENTA  NOVENTA"
    DU = "ONZE     DOZE     TREZE    QUATORZE QUINZE   DEZESSEISDEZESSETEDEZOITO  DEZENOVE"
    CT = "CENTO       DUZENTOS    TREZENTOS   QUATROCENTOSQUINHENTOS  SEISCENTOS  SETECENTOS  OITOCENTOS  NOVECENTOS"
    '-- desmonta o valor em partes
    CTNA = Int(Valor / 100)
    DZNA = Int((Valor Mod 100) / 10)
    UNID = Valor Mod 10
    '-- monta o valor por extenso
    If CTNA > 0 Then
        Txt = Trim$(Mid$(CT, 12 * CTNA - 11, 12)) & IIf(DZNA + UNID > 0, " E ", "")
    End If
    If DZNA = 1 And UNID > 0 Then
        Txt = Txt & Trim$(Mid$(DU, 9 * UNID - 8, 9))
    ElseIf DZNA > 0 Then
        Txt = Txt & Trim$(Mid$(DZ, 9 * DZNA - 8, 9)) & IIf(UNID > 0, " E ", "")
    End If
    If DZNA <> 1 And UNID > 0 Then
        Txt = Txt & Trim$(Mid$(UN, 6 * UNID - 5, 6))
    End If
    '-- devolve o extenso
    extenso = Txt

End Function
Function ValExtenso(ValExt) As String
    '-- extensifica��o de valores monet�rios
    Dim Vtxt As String
    Dim Ast1, Ast2, Ast3  As String * 50
    Dim Ast4 As String * 2
    Dim Part1 As Long, Part2 As Long, Part3 As Long
    '-- verifica se � v�lido
    If ValExt <= 0 Or ValExt >= 1000000 Then
        Exit Function
    End If
    
    '-- desmembra o valor a extensificar
    Part1 = Int(ValExt / 1000)          'milhares
    Part2 = Int(ValExt) - Part1 * 1000  'unidades
    Part3 = ValExt * 100 - Part2 * 100 - Part1 * 100000
    Vtxt = IIf(Part1 = 0, "", extenso(CInt(Part1)) & " MIL")
    Vtxt = Vtxt & IIf(Part1 > 0, IIf(Part2 > 0, IIf(Part3 > 0, ", ", " E "), ""), "")
    Vtxt = Vtxt & IIf(Part2 = 0, "", extenso(CInt(Part2)))
    Vtxt = Vtxt & IIf(Part1 + Part2 > 0, " REA" & IIf(Part1 = 0 And Part2 = 1, "L", "IS"), "")
    Vtxt = Vtxt & IIf(Part3 > 0, IIf(Part1 + Part2 > 0, " E ", "") & extenso(CInt(Part3)) & " CENTAVO" & IIf(Part3 > 1, "S", ""), "")
    '-- retorna o valor por extenso
    ValExtenso = Vtxt

End Function

Function GraficoFisicoItem()
On Error GoTo TratamentodeErro:
Dim bd As DataBase, QGraf As QueryDef, GrafTmp As Recordset
Dim Graf As Recordset
Set bd = CurrentDb()
Set GrafTmp = bd.OpenRecordset("GraficoFisico", DB_OPEN_TABLE)
Set QGraf = bd.QueryDefs("GraficoFisicoItemQuery")
QGraf.Parameters(0) = Forms!GraficoVendasFisicas!DataInicial
QGraf.Parameters(1) = Forms!GraficoVendasFisicas!DataFinal
Set Graf = QGraf.OpenRecordset()
GrafTmp.index = "PrimaryKey"
If GrafTmp.RecordCount <> 0 Then
   BeginTrans
   While Not GrafTmp.EOF
      GrafTmp.Delete
      GrafTmp.MoveNext
   Wend
   CommitTrans
End If

If Graf.RecordCount <> 0 Then
   BeginTrans
   While Not Graf.EOF
      GrafTmp.Seek "=", Graf!CodItem, Graf!MESANO
      If GrafTmp.NoMatch Then
         GrafTmp.AddNew
       Else
         GrafTmp.Edit
      End If
      GrafTmp!CodItem = Graf!CodItem
      GrafTmp!MESANO = Graf!MESANO
      GrafTmp!Venda = GrafTmp!Venda + Graf!Venda
      GrafTmp.Update
      Graf.MoveNext
   Wend
   CommitTrans
End If
Graf.Close
GrafTmp.Close
bd.Close
Set Graf = Nothing
Set GrafTmp = Nothing
Set bd = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function SenhaDigitada() As String
On Error GoTo TratamentodeErro:
Dim bd As DataBase, Sen As Recordset
Dim Graf As Recordset
Set bd = CurrentDb()
Set Sen = bd.OpenRecordset("SenResp", DB_OPEN_TABLE)
'BeginTrans
If Sen.RecordCount <> 0 Then
   While Not Sen.EOF
      If IsNull(Sen!SenDigitada) Then
         SenhaDigitada = "LIVRE"
      Else
         SenhaDigitada = Sen!SenDigitada
      End If
      Sen.MoveNext
   Wend
Else
   SenhaDigitada = "LIVRE"
End If
'CommitTrans
Sen.Close
bd.Close
Set Sen = Nothing
Set bd = Nothing

Exit Function

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function
Function LimpaSenhaDigitada()
On Error GoTo TratamentodeErro:
Dim bd As DataBase, Sen As Recordset
Dim Graf As Recordset
Set bd = CurrentDb()
Set Sen = bd.OpenRecordset("SenResp", DB_OPEN_TABLE)
If Sen.RecordCount <> 0 Then
   BeginTrans
   While Not Sen.EOF
      Sen.Edit
      Sen!SenDigitada = "LIVRE"
      Sen.Update
      Sen.MoveNext
   Wend
   CommitTrans
End If
Sen.Close
bd.Close
Set Sen = Nothing
Set bd = Nothing
Exit Function

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function TituloOrdemServico(WNroSerie As Long) As String

On Error GoTo TratamentodeErro:
Dim emp As Recordset, bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set emp = bd1.OpenRecordset("Empresa", DB_OPEN_TABLE)

Do Until emp.EOF
   If Not IsNull(emp!TituloOSV) Then
      If emp!CodigoEmpresa = Form_Menu.WEmpresa Then
         TituloOrdemServico = UCase(emp!TituloOSV)
      End If
   Else
      TituloOrdemServico = "ORDEM DE SERVICO E VENDA"
   End If
   emp.MoveNext
Loop
emp.Close
bd1.Close
Set emp = Nothing
Set bd1 = Nothing
   Exit Function

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function


Function GraficoFisicoCategoria()
Dim bd As DataBase, QGraf As QueryDef, GrafTmp As Recordset
Dim Graf As Recordset
Set bd = CurrentDb()
Set GrafTmp = bd.OpenRecordset("GraficoFisico", DB_OPEN_TABLE)
Set QGraf = bd.QueryDefs("GraficoFisicoCategoriaQuery")
QGraf.Parameters(0) = Forms!GraficoVendasFisicasCategoria!AnoDesejado
QGraf.Parameters(1) = Forms!GraficoVendasFisicasCategoria!Item1
QGraf.Parameters(2) = Forms!GraficoVendasFisicasCategoria!Item2
QGraf.Parameters(3) = Forms!GraficoVendasFisicasCategoria!Item3
Set Graf = QGraf.OpenRecordset()
GrafTmp.index = "PrimaryKey"
If GrafTmp.RecordCount <> 0 Then
   BeginTrans
   Do Until GrafTmp.EOF
      GrafTmp.Delete
      GrafTmp.MoveNext
   Loop
   CommitTrans
End If

If Graf.RecordCount <> 0 Then
   BeginTrans
   Do Until Graf.EOF
      GrafTmp.Seek "=", Graf!CodigoCategoria, Graf!MESANO
      If GrafTmp.NoMatch Then
         GrafTmp.AddNew
       Else
         GrafTmp.Edit
      End If
      GrafTmp!CodItem = Graf!CodigoCategoria
      GrafTmp!MESANO = Graf!MESANO
      GrafTmp!Venda = GrafTmp!Venda + Graf!Venda
      GrafTmp.Update
      Graf.MoveNext
   Loop
   CommitTrans
End If
Graf.Close
GrafTmp.Close
bd.Close
End Function

Function Gera_CTRECAUX_Boletos(WNumos As Variant) As Long
On Error GoTo TratamentodeErro:
Dim bd, bdt As DataBase, QCTR As QueryDef, CTR As Recordset
Dim CTAUX As Recordset
Set bd = CurrentDb()
Set bdt = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set CTAUX = bdt.OpenRecordset("CTRECAUX", DB_OPEN_TABLE)
Set QCTR = bd.QueryDefs("Gera_CTRECAUX_Boletos")
QCTR.Parameters(0) = WNumos
Set CTR = QCTR.OpenRecordset()
CTAUX.index = "PrimaryKey"

If CTAUX.RecordCount <> 0 Then
   BeginTrans
   While Not CTAUX.EOF
         If CTAUX!CodEmpresa = Form_Menu.WEmpresa Then
            If CurrentUser() = CTAUX!Current_User Then
               CTAUX.Delete
            End If
         End If
        CTAUX.MoveNext
   Wend
   CommitTrans
End If

If CTR.RecordCount <> 0 Then
   BeginTrans
   While Not CTR.EOF
      If CTR!CodEmpresa = Form_Menu.WEmpresa Then
      If CTR!EmAberto Then
         If CTR!Carteira = 0 Then
            CTAUX.Seek "=", CTR!NroConta
            If CTAUX.NoMatch Then
               CTAUX.AddNew
             Else
               CTAUX.Edit
             End If
            ' CTAUX!NroConta = CTR!NroConta
             CTAUX!NotaOSCupon = CTR!NotaOSCupon
             CTAUX!DataEmissao = CTR!DataEmissao
             CTAUX!CodigoCliente = CTR!CodigoCliente
             CTAUX!DataVencimento = CTR!DataVencimento
             CTAUX!Sequencia = CTR!Sequencia
             If CTR!ReterISSQN <> 0 Then
                CTAUX!Valor = (CTR!Valor - CTR!TotalPago)
                CTAUX!Observacoes = "Ret.ISSQN Aliq % " & Format$(CTR!Aliq_RetencaoISSQN, "0.00") & " R$ " & Format$(CTR!Vlr_RetencaoISSQN, "0.00") & " " & CTR!Observacoes
             Else
                CTAUX!Observacoes = CTR!Observacoes
                CTAUX!Valor = (CTR!Valor - CTR!TotalPago)
             End If
             CTAUX!Vlr_RetencaoISSQN = CTR!Vlr_RetencaoISSQN
             CTAUX!Aliq_RetencaoISSQN = CTR!Aliq_RetencaoISSQN
             CTAUX!CodigoDepto = CTR!CodigoDepto
             CTAUX!CodEmpresa = Form_Menu.WEmpresa
             CTAUX!NroConta_CTREC = CTR!NroConta
             CTAUX!Current_User = CurrentUser()
             CTAUX!Numos = CTR!Numos
             CTAUX!Cart_Digital = CTR!Cart_Digital
             CTAUX.Update
         End If
      End If
      End If
      CTR.MoveNext
   Wend
   CommitTrans
End If

CTAUX.Close
CTR.Close
bd.Close
bdt.Close
Set CTAUX = Nothing
Set CTR = Nothing
Set bd = Nothing
Set bdt = Nothing
Exit Function

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function RetornaPrevRecXPrazo(WCodigoPagto As Long) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Dim WTotal As Double
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("CONDPAGTO", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
'BeginTrans
tb.Seek "=", WCodigoPagto
If tb.NoMatch Then
   RetornaPrevRecXPrazo = 0
Else
   If IsNull(tb!CodigoRecurso) Then
       RetornaPrevRecXPrazo = 0
   Else
      RetornaPrevRecXPrazo = tb!CodigoRecurso
   End If
End If

Finaliza:
 ' CommitTrans
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
 
End Function

Function GeraDupAUX(WCodigoCliente As Long)
On Error GoTo TratamentodeErro:
Dim bd, bdt As DataBase, QCTR As QueryDef, CTR As Recordset
Dim CTAUX As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set bdt = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set CTAUX = bdt.OpenRecordset("CTRECAUX", DB_OPEN_TABLE)
Set CTR = bd.OpenRecordset("CTREC", DB_OPEN_TABLE)
CTAUX.index = "PrimaryKey"
CTR.index = "QuintaChave"


GoSub Inicializa_CTRECAUX:

BeginTrans

CTR.Seek "=", WCodigoCliente
If Not CTR.NoMatch Then
   While CTR!CodigoCliente = WCodigoCliente
         CTAUX.Seek "=", CTR!NroConta
         If CTAUX.NoMatch Then
            CTAUX.AddNew
          Else
            CTAUX.Edit
         End If
         CTAUX!NroConta = CTR!NroConta
         CTAUX!NotaOSCupon = CTR!NotaOSCupon
         CTAUX!DataEmissao = CTR!DataEmissao
         CTAUX!CodigoCliente = CTR!CodigoCliente
         CTAUX!DataVencimento = CTR!DataVencimento
         If CTR!DataEmissao <> CTR!DataVencimento Then
            CTAUX!Valor = CTR!Valor - ApuraPagtosCTREC(CTR!NotaOSCupon, CTR!CodigoCliente, CTR!DataVencimento)
         Else
             CTAUX!Valor = CTR!Valor
         End If
         CTAUX!Sequencia = CTR!Sequencia
         CTAUX!NossoNumero = CTR!NossoNumero
         CTAUX!Carteira = CTR!Carteira
         CTAUX!CodigoDepto = CTR!CodigoDepto
         CTAUX!CodEmpresa = Form_Menu.WEmpresa
         CTAUX!NroNotaFiscal = CTR!NotaOSCupon
         CTAUX.Update
         CTR.MoveNext
         If CTR.EOF Then
            GoSub Finaliza
         End If
   Wend
End If

Finaliza:
   CommitTrans
   CTAUX.Close
   CTR.Close
   bd.Close
   bdt.Close
   Set CTAUX = Nothing
   Set CTR = Nothing
  Set bd = Nothing
  Set bdt = Nothing
  Exit Function
Return

Inicializa_CTRECAUX:
   If CTAUX.RecordCount <> 0 Then
      BeginTrans
      While Not CTAUX.EOF
         CTAUX.Delete
         CTAUX.MoveNext
      Wend
      CommitTrans
   End If
Return


TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function AtualizaPrecoMedioCompra()
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim bd As DataBase
Dim ATUPREC, Est, EMPXEST As Recordset
Dim Ent As Recordset
Dim WTotalCompras As Double, WValorTotalCompras As Double, WPrecoMedioCompras As Double
Dim WCodItem As Double
Dim QATUPREC As QueryDef
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set EMPXEST = bd1.OpenRecordset("EMPXEST", DB_OPEN_TABLE)
Set Ent = bd1.OpenRecordset("DetEntradas", DB_OPEN_TABLE)
Set Est = bd1.OpenRecordset("EST", DB_OPEN_TABLE)
Est.index = "PrimaryKey"
EMPXEST.index = "PrimaryKey"
Ent.index = "CodItem"
If Ent.RecordCount <> 0 Then
   BeginTrans
   Do Until Ent.EOF
       Est.Seek "=", Ent!CodItem
       If Not Est.NoMatch Then
          If WCodItem = 0 Then
             WCodItem = Ent!CodItem
          End If
          If WCodItem <> Ent!CodItem Then
             GoSub CalculaPrecoMedioCompras
             WCodItem = Ent!CodItem
             WTotalCompras = 0
             WValorTotalCompras = 0
         End If
         WTotalCompras = WTotalCompras + Ent!Quantidade
         WValorTotalCompras = WValorTotalCompras + (Ent!Quantidade * Ent!ValorUnit)
      End If
      Ent.MoveNext
   Loop
   GoSub CalculaPrecoMedioCompras
   CommitTrans
End If
Ent.Close
EMPXEST.Close
Est.Close
bd1.Close
Set Ent = Nothing
Set EMPXEST = Nothing
Set Est = Nothing
Set bd1 = Nothing
Exit Function

CalculaPrecoMedioCompras:
    WPrecoMedioCompras = WValorTotalCompras / WTotalCompras
    EMPXEST.Seek "=", 1, WCodItem
    If EMPXEST.NoMatch Then
       EMPXEST.AddNew
    Else
       EMPXEST.Edit
   End If
   EMPXEST!CodigoEmpresa = 1
   EMPXEST!CodigoItem = WCodItem
   EMPXEST!CustoMedio = WPrecoMedioCompras
   EMPXEST.Update
 Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function


Function GeraTabelaXPrecosXServ() As Long
On Error GoTo TratamentodeErro:
DoCmd.Hourglass True
Dim bd1 As DataBase
Dim bd As DataBase
Dim TabPrec, TABC, serv As Recordset
Dim WDelta As Double
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set TABC = bd1.OpenRecordset("TABSERV", DB_OPEN_TABLE)
Set TabPrec = bd.OpenRecordset("TabelaxPrecos", DB_OPEN_TABLE)
Set serv = bd1.OpenRecordset("TPSERV", DB_OPEN_TABLE)
TabPrec.index = "PrimaryKey"
TABC.index = "PrimaryKey"
serv.index = "PrimaryKey"
If TabPrec.RecordCount <> 0 Then
   BeginTrans
   While Not TabPrec.EOF
      TabPrec.Delete
      TabPrec.MoveNext
   Wend
   CommitTrans
End If

BeginTrans
If TABC.RecordCount <> 0 Then
   While Not TABC.EOF
         GoSub CriaTabela:
         TABC.MoveNext
   Wend
End If

Finaliza:
  CommitTrans
  TABC.Close
  TabPrec.Close
  serv.Close
  DoCmd.Hourglass False
  bd.Close
  bd1.Close
  Set TABC = Nothing
  Set TabPrec = Nothing
  Set serv = Nothing
  Set bd = Nothing
  Set bd1 = Nothing
  Exit Function
Return

CriaTabela:
   If serv.RecordCount <> 0 Then
      serv.MoveFirst
      While Not serv.EOF
            TabPrec.Seek "=", serv!TipoServico, TABC!CODIGOtABELA
            If TabPrec.NoMatch Then
               TabPrec.AddNew
            Else
               TabPrec.Edit
            End If
            TabPrec!CodItem = serv!TipoServico
            TabPrec!tABELA = TABC!CODIGOtABELA
            If IsNull(TABC!Tipo) Then
               TabPrec!PrecoVenda = (Int(((serv!ValorMaoObra - (Int((((serv!ValorMaoObra * TABC!Delta) / 100) + 0.005) * 100)) / 100) + 0.005) * 100)) / 100
            ElseIf TABC!Tipo = "D" Then
                   TabPrec!PrecoVenda = (Int(((serv!ValorMaoObra - (Int((((serv!ValorMaoObra * TABC!Delta) / 100) + 0.005) * 100)) / 100) + 0.005) * 100)) / 100
            Else
                  TabPrec!PrecoVenda = (Int(((serv!ValorMaoObra + (Int((((serv!ValorMaoObra * TABC!Delta) / 100) + 0.005) * 100)) / 100) + 0.005) * 100)) / 100
            End If
            TabPrec.Update
            serv.MoveNext
      Wend
   End If

Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Captura_Vlr_Frete(CodItem As Long, CodEmpresa As Long) As Double
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim EMPXEST  As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set EMPXEST = bd1.OpenRecordset("EmpxEst", DB_OPEN_TABLE)
EMPXEST.index = "PrimaryKey"

EMPXEST.Seek "=", Form_Menu.WEmpresa, CodItem

If EMPXEST.NoMatch Then
   Captura_Vlr_Frete = 0
Else
   If IsNull(EMPXEST!Vlr_Fret) Then
      Captura_Vlr_Frete = 0
   ElseIf EMPXEST!Vlr_Fret < 0 Then
          BeginTrans
          EMPXEST.Edit
          EMPXEST!Vlr_Fret = 0
          EMPXEST.Update
          CommitTrans
          Captura_Vlr_Frete = 0
   Else
      Captura_Vlr_Frete = EMPXEST!Vlr_Fret
   End If
End If

Finaliza:
   EMPXEST.Close
   bd1.Close
   Set EMPXEST = Nothing
   Set bd1 = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function


Function Captura_Vlr_Out_Desp(CodItem As Long, CodEmpresa As Long) As Double
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim EMPXEST  As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set EMPXEST = bd1.OpenRecordset("EmpxEst", DB_OPEN_TABLE)
EMPXEST.index = "PrimaryKey"

EMPXEST.Seek "=", Form_Menu.WEmpresa, CodItem

If EMPXEST.NoMatch Then
   Captura_Vlr_Out_Desp = 0
Else
   If IsNull(EMPXEST!Vlr_Out_Desp) Then
      Captura_Vlr_Out_Desp = 0
   ElseIf EMPXEST!Vlr_Out_Desp < 0 Then
          BeginTrans
          EMPXEST.Edit
          EMPXEST!Vlr_Out_Desp = 0
          EMPXEST.Update
          CommitTrans
          Captura_Vlr_Out_Desp = 0
   Else
      Captura_Vlr_Out_Desp = EMPXEST!Vlr_Out_Desp
   End If
End If


Finaliza:
   EMPXEST.Close
   bd1.Close
   Set EMPXEST = Nothing
   Set bd1 = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function


Function Captura_Vlr_IPI(CodItem As Long, CodEmpresa As Long) As Double
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim EMPXEST  As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set EMPXEST = bd1.OpenRecordset("EmpxEst", DB_OPEN_TABLE)
EMPXEST.index = "PrimaryKey"

EMPXEST.Seek "=", Form_Menu.WEmpresa, CodItem

If EMPXEST.NoMatch Then
   Captura_Vlr_IPI = 0
Else
   If IsNull(EMPXEST!IPI) Then
      Captura_Vlr_IPI = 0
   ElseIf EMPXEST!IPI < 0 Then
          BeginTrans
          EMPXEST.Edit
          EMPXEST!IPI = 0
          EMPXEST.Update
          CommitTrans
          Captura_Vlr_IPI = 0
   Else
      Captura_Vlr_IPI = (EMPXEST!PrecoVenda * EMPXEST!IPI) / 100
   End If
End If


Finaliza:
   EMPXEST.Close
   bd1.Close
   Set EMPXEST = Nothing
   Set bd1 = Nothing
   Exit Function
Return

Valida_Campos_Nulo:
   BeginTrans
   EMPXEST.Edit
   If IsNull(EMPXEST!IPI) Then
      EMPXEST!IPI = 0
   End If
   If IsNull(EMPXEST!PrecoVenda) Then
      EMPXEST!PrecoVenda = 0
   End If
   EMPXEST.Update
  CommitTrans
Return


TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function CapturaPrecoVenda(CodItem As Long, CodEmpresa As Long) As Double
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim EMPXEST  As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set EMPXEST = bd1.OpenRecordset("EmpxEst", DB_OPEN_TABLE)
EMPXEST.index = "PrimaryKey"

EMPXEST.Seek "=", Form_Menu.WEmpresa, CodItem

If EMPXEST.NoMatch Then
   CapturaPrecoVenda = 0
Else
   If IsNull(EMPXEST!PrecoVenda) Then
      CapturaPrecoVenda = 0
   Else
      CapturaPrecoVenda = EMPXEST!PrecoVenda
   End If
End If

Finaliza:
   EMPXEST.Close
   bd1.Close
   Set EMPXEST = Nothing
   Set bd1 = Nothing
   Exit Function
Return

Valida_Campos_Nulo:
   BeginTrans
   EMPXEST.Edit
   If IsNull(EMPXEST!IPI) Then
      EMPXEST!IPI = 0
   End If
   If IsNull(EMPXEST!PrecoVenda) Then
      EMPXEST!PrecoVenda = 0
   End If
   EMPXEST.Update
  CommitTrans
Return


TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function


Function CapturaEstoqueMinimo(CodItem As Long, CodEmpresa As Long) As Double ' convertido
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim EMPXEST  As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set EMPXEST = bd1.OpenRecordset("EmpxEst", DB_OPEN_TABLE)
EMPXEST.index = "PrimaryKey"

EMPXEST.Seek "=", Form_Menu.WEmpresa, CodItem

If EMPXEST.NoMatch Then
  CapturaEstoqueMinimo = 0
Else
  CapturaEstoqueMinimo = EMPXEST!Minimo
End If

EMPXEST.Close
bd1.Close
Set EMPXEST = Nothing
Set bd1 = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function CapturaDescricaoProcessoDescont_Tanque(WCodigo As Long) As String
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb  As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("PROC_DESCONT", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

tb.Seek "=", WCodigo

If Not tb.NoMatch Then
   If IsNull(tb!Descricao) Then
      CapturaDescricaoProcessoDescont_Tanque = "nula"
   Else
      CapturaDescricaoProcessoDescont_Tanque = tb!Descricao
   End If
Else
   CapturaDescricaoProcessoDescont_Tanque = "Processo inexistente"
End If

Finaliza:
 tb.Close
 bd.Close
 Set tb = Nothing
 Set bd = Nothing
 Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function
Function CapturaDescricaoBanco(WBanco As Long) As String
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim Bco  As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Bco = bd1.OpenRecordset("BCO", DB_OPEN_TABLE)
Bco.index = "PrimaryKey"

Bco.Seek "=", WBanco

If Not Bco.NoMatch Then
   CapturaDescricaoBanco = Bco!RazaoSocial
Else
   CapturaDescricaoBanco = "Banco nao cadastrado..."
End If

Finaliza:
 Bco.Close
 bd1.Close
 Set Bco = Nothing
 Set bd1 = Nothing
 Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaSituacaoBloqueio(wCodCliente As Long) As String

On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim Cli  As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
Cli.index = "PrimaryKey"

Cli.Seek "=", wCodCliente

If Not Cli.NoMatch Then
   If IsNull(Cli!Bloqueado) Or Cli!Bloqueado <> "S" Then
      CapturaSituacaoBloqueio = "N"
   ElseIf Cli!Bloqueado = "S" Then
      CapturaSituacaoBloqueio = "S"
   End If
Else
   CapturaSituacaoBloqueio = "N"
End If

Finaliza:
 Cli.Close
 bd1.Close
 Set Cli = Nothing
 Set bd1 = Nothing
 Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaRefInt(WCodItem As Long) As String
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim Est  As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Est = bd1.OpenRecordset("EST", DB_OPEN_TABLE)
Est.index = "PrimaryKey"

Est.Seek "=", WCodItem

If Not Est.NoMatch Then
   If Not IsNull(Est!CodigoInterno) Then
      CapturaRefInt = Est!CodigoInterno
   Else
     CapturaRefInt = "x"
   End If
Else
   CapturaRefInt = " "
End If

Finaliza:
 Est.Close
 bd1.Close
 Set Est = Nothing
 Set bd1 = Nothing
 Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function CapturaRefOrig(WCodItem As Long) As String
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim Est  As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Est = bd1.OpenRecordset("EST", DB_OPEN_TABLE)
Est.index = "PrimaryKey"

Est.Seek "=", WCodItem

If Not Est.NoMatch Then
   If Not IsNull(Est!REFORIG) Then
      CapturaRefOrig = Est!REFORIG
   Else
     CapturaRefOrig = "x"
   End If
Else
   CapturaRefOrig = "X"
End If

Finaliza:
 Est.Close
 bd1.Close
 Set Est = Nothing
 Set bd1 = Nothing
 Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaDescricaoItemEstoque(WCodItem As Long) As String
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim Est  As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Est = bd1.OpenRecordset("EST", DB_OPEN_TABLE)
Est.index = "PrimaryKey"

Est.Seek "=", WCodItem

If Not Est.NoMatch Then
   If Not IsNull(Est!Descricao) And Est!Descricao <> "" Then
      CapturaDescricaoItemEstoque = Mid$(Est!Descricao, 1, 50)
   Else
      CapturaDescricaoItemEstoque = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
   End If
Else
   CapturaDescricaoItemEstoque = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
End If

Finaliza:
 Est.Close
 bd1.Close
 Set Est = Nothing
 Set bd1 = Nothing
 Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaCodItemXDescricao(WDescricao As String) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb  As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("EST", DB_OPEN_TABLE)
tb.index = "Descricao"

tb.Seek "=", WDescricao

If Not tb.NoMatch Then
   CapturaCodItemXDescricao = tb!CodItem
Else
   CapturaCodItemXDescricao = 0
End If

Finaliza:
 tb.Close
 bd.Close
 Set tb = Nothing
 Set bd = Nothing
 Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaCodigoFabricante(WCodItem As Long) As String
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim Est  As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Est = bd1.OpenRecordset("EST", DB_OPEN_TABLE)
    Est.index = "PrimaryKey"

Est.Seek "=", WCodItem

If Not Est.NoMatch Then
   If IsNull(Est!CodigoFabricante) Then
      CapturaCodigoFabricante = "0000000000000000000"
   Else
      CapturaCodigoFabricante = Mid$(Est!CodigoFabricante, 1, 19)
   End If
Else
   CapturaCodigoFabricante = "0000000000000000000"
End If

Finaliza:
 Est.Close
 bd1.Close
 Set Est = Nothing
 Set bd1 = Nothing
 Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaConsumoCategoria(WCategoria As Long) As Boolean
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim Cat  As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Cat = bd1.OpenRecordset("CATPROD", DB_OPEN_TABLE)
Cat.index = "PrimaryKey"

Cat.Seek "=", WCategoria

If Not Cat.NoMatch Then
   CapturaConsumoCategoria = Cat!ConsumoInterno
End If

Finaliza:
 Cat.Close
 bd1.Close
 Set Cat = Nothing
 Set bd1 = Nothing
 Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaCategoriaEST(WCodItem As Long) As Long
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim Est  As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Est = bd1.OpenRecordset("EST", DB_OPEN_TABLE)
Est.index = "PrimaryKey"

Est.Seek "=", WCodItem

If Not Est.NoMatch Then
   CapturaCategoriaEST = Est!CodigoCategoria
End If

Finaliza:
 Est.Close
 bd1.Close
 Set Est = Nothing
 Set bd1 = Nothing
 Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaAbreviaturaMarcaItemEstoque(WCodItem As Long) As String
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim Est, Mar  As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Est = bd1.OpenRecordset("EST", DB_OPEN_TABLE)
Set Mar = bd1.OpenRecordset("MARCAS", DB_OPEN_TABLE)
Est.index = "PrimaryKey"
Mar.index = "PrimaryKey"
Est.Seek "=", WCodItem

If Not Est.NoMatch Then
   Mar.Seek "=", Est!CodigoMarca
   If Mar.NoMatch Then
      CapturaAbreviaturaMarcaItemEstoque = "BOSCH"
      GoSub Finaliza:
   End If
   If IsNull(Mar!ABREVIATURA) Then
      CapturaAbreviaturaMarcaItemEstoque = "BOSCH"
      GoSub Finaliza:
   End If
   CapturaAbreviaturaMarcaItemEstoque = Mar!ABREVIATURA
Else
   CapturaAbreviaturaMarcaItemEstoque = "BOSCH"
End If

Finaliza:
 Est.Close
 Mar.Close
 bd1.Close
 Set Est = Nothing
 Set Mar = Nothing
 Set bd1 = Nothing
 Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CriaDiasVisitas(CodCliente As Long)
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim Cli  As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Cli = bd1.OpenRecordset("CLIXVISITA", DB_OPEN_TABLE)
Cli.index = "primarykey"

BeginTrans
Cli.Seek "=", CodCliente, 0


If Cli.NoMatch Then
   Cli.AddNew
   Cli!CodigoCliente = CodCliente
   Cli!CodigoVisita = 0
   Cli.Update
End If
CommitTrans
Cli.Close
bd1.Close
Set Cli = Nothing
Set bd1 = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function GeraTabelaxPrecosGeral()
On Error GoTo TratamentodeErro:
DoCmd.Hourglass True
Dim bd1 As DataBase
Dim bd As DataBase
Dim X As Long
Dim TabPrec, TABC, Est, EMPXEST As Recordset
Dim WDelta As Double, WTabela As Double
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set TABC = bd1.OpenRecordset("TAB", DB_OPEN_TABLE)
Set TabPrec = bd.OpenRecordset("TabelaxPrecos", DB_OPEN_TABLE)
Set EMPXEST = bd1.OpenRecordset("EMPXEST", DB_OPEN_TABLE)
Set Est = bd1.OpenRecordset("EST", DB_OPEN_TABLE)
TabPrec.index = "PrimaryKey"
EMPXEST.index = "PrimaryKey"
TABC.index = "PrimaryKey"
Est.index = "PrimaryKey"
WTabela = 0
If TabPrec.RecordCount <> 0 Then
   BeginTrans
   While Not TabPrec.EOF
      TabPrec.Delete
      TabPrec.MoveNext
   Wend
   CommitTrans
End If

If TABC.RecordCount <> 0 Then
   While Not TABC.EOF
      If IsNull(TABC!Delta) Then
         WDelta = 0
        Else
         WDelta = TABC!Delta
      End If
      BeginTrans
      While Not Est.EOF
            If Est!Ativo Then
               TabPrec.Seek "=", Est!CodItem, TABC!CODIGOtABELA
               If TabPrec.NoMatch Then
                  TabPrec.AddNew
               Else
                  TabPrec.Edit
               End If
               TabPrec!CodItem = Est!CodItem
               TabPrec!tABELA = TABC!CODIGOtABELA
            
               EMPXEST.Seek "=", Form_Menu.WEmpresa, Est!CodItem
               If EMPXEST.NoMatch Then
                  TabPrec!PrecoVenda = 0
               Else
                  TabPrec!PrecoVenda = (Int(((EMPXEST!PrecoVenda - (Int((((EMPXEST!PrecoVenda * TABC!Delta) / 100) + 0.005) * 100)) / 100) + 0.005) * 100)) / 100
               End If
               If Not IsNull(Est!Aplicacao) Then
                  TabPrec!Aplicacao = Mid$(Est!Aplicacao, 1, 50)
               End If
               'TabPrec!SALDOESTOQUE = SaldoFinalItemEstoque(Est!CodItem, Date)
               TabPrec.Update
            End If
         
            X = X + 1
            If X Mod 500 = 0 Then
               CommitTrans
               BeginTrans
            End If
            Est.MoveNext
      Wend
      CommitTrans
      Est.MoveFirst
      TABC.MoveNext
   Wend
End If
TABC.Close
TabPrec.Close
EMPXEST.Close
Est.Close
bd1.Close
bd.Close
Set TABC = Nothing
Set TabPrec = Nothing
Set EMPXEST = Nothing
Set Est = Nothing
Set bd = Nothing
Set bd1 = Nothing

DoCmd.Hourglass False

Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function PrecoMedioCompraCodItem(CodItem As Long)
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim bd As DataBase
Dim ATUPREC, Est, EMPXEST As Recordset
Dim Ent As Recordset
Dim WTotalCompras As Double, WValorTotalCompras As Double, WPrecoMedioCompras As Double
Dim WCodItem As Double
Dim QATUPREC As QueryDef
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set EMPXEST = bd1.OpenRecordset("EMPXEST", DB_OPEN_TABLE)
Set Ent = bd1.OpenRecordset("DetEntradas", DB_OPEN_TABLE)
Set Est = bd1.OpenRecordset("EST", DB_OPEN_TABLE)
Est.index = "PrimaryKey"
EMPXEST.index = "PrimaryKey"
Ent.index = "CodItem"

WTotalCompras = 0
WValorTotalCompras = 0
WPrecoMedioCompras = 0
If CodItem <> 0 Then
    GoSub VerificaItem:
    GoSub ApuraEntradas:
    GoSub CalculaPrecoMedioCompras:
    'Ent.MoveNext
End If
Ent.Close
EMPXEST.Close
Est.Close
bd1.Close
Set Ent = Nothing
Set EMPXEST = Nothing
Set Est = Nothing
Set bd1 = Nothing
Exit Function

VerificaItem:
  Est.Seek "=", CodItem
  If Est.NoMatch Then
     MsgBox "Produto Inexistente, contacte o suporte...", 64, "SGa"
     Exit Function
   End If
Return

ApuraEntradas:
  Ent.Seek "=", CodItem
  If Ent.NoMatch Then
     WPrecoMedioCompras = 0
   Else
     Do Until Ent.EOF
          If Ent!CodItem <> CodItem Then
            Exit Do
         Else
            WTotalCompras = WTotalCompras + Ent!Quantidade
            WValorTotalCompras = WValorTotalCompras + (Ent!Quantidade * Ent!ValorUnit)
         End If
         Ent.MoveNext
     Loop
  End If
Return

CalculaPrecoMedioCompras:
If WTotalCompras > 0 Then
   BeginTrans
    WPrecoMedioCompras = WValorTotalCompras / WTotalCompras
    EMPXEST.Seek "=", 1, CodItem
    If EMPXEST.NoMatch Then
       EMPXEST.AddNew
    Else
       EMPXEST.Edit
   End If
   EMPXEST!CodigoEmpresa = 1
   EMPXEST!CodigoItem = CodItem
   EMPXEST!CustoMedio = WPrecoMedioCompras
   EMPXEST.Update
   CommitTrans
End If
 Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function ForXEstItem(WCodItem As Long, WFornec As Long) As Long
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim FOREST, Est, EMPXEST As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Est = bd1.OpenRecordset("EST", DB_OPEN_TABLE)
Set EMPXEST = bd1.OpenRecordset("EMPXEST", DB_OPEN_TABLE)
Set FOREST = bd1.OpenRecordset("FORXEST", DB_OPEN_TABLE)
FOREST.index = "Primarykey"
EMPXEST.index = "PrimaryKey"
Est.index = "PrimaryKey"

GoSub AnalisaEMPXEST:
GoSub AnaliseFORXEST:
GoSub Finaliza:

AnalisaEMPXEST:
   BeginTrans
   EMPXEST.Seek "=", Form_Menu.WEmpresa, WCodItem
   If EMPXEST.NoMatch Then
      EMPXEST.AddNew
   Else
      EMPXEST.Edit
   End If
   EMPXEST!CodigoEmpresa = Form_Menu.WEmpresa
   EMPXEST!CodigoItem = WCodItem
   EMPXEST.Update
   CommitTrans
Return
    
AnaliseFORXEST:
   BeginTrans
   FOREST.Seek "=", Form_Menu.WEmpresa, WCodItem, WFornec
   If FOREST.NoMatch Then
      FOREST.AddNew
   Else
      FOREST.Edit
   End If
   FOREST!CodigoEmpresa = Form_Menu.WEmpresa
   FOREST!CodigoItem = WCodItem
   FOREST!CodigoFornecedor = WFornec
   FOREST.Update
   CommitTrans
Return

Finaliza:
   FOREST.Close
   EMPXEST.Close
   Est.Close
   bd1.Close
   Set FOREST = Nothing
   Set EMPXEST = Nothing
   Set Est = Nothing
   Set bd1 = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function FORXEST() As Long
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase

Dim FOREST, DETENT As Recordset, Est As Recordset, EMPXEST As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Est = bd1.OpenRecordset("EST", DB_OPEN_TABLE)
Set EMPXEST = bd1.OpenRecordset("EMPXEST", DB_OPEN_TABLE)
Set FOREST = bd1.OpenRecordset("FORXEST", DB_OPEN_TABLE)
Set DETENT = bd1.OpenRecordset("DetEntradas", DB_OPEN_TABLE)
FOREST.index = "Primarykey"
EMPXEST.index = "PrimaryKey"
Est.index = "PrimaryKey"



If Est.RecordCount <> 0 Then
   BeginTrans
   Do While Not Est.EOF
      EMPXEST.Seek "=", Form_Menu.WEmpresa, Est!CodItem
      If EMPXEST.NoMatch Then
         EMPXEST.AddNew
      Else
         EMPXEST.Edit
      End If
      EMPXEST!CodigoEmpresa = Form_Menu.WEmpresa
      EMPXEST!CodigoItem = Est!CodItem
      EMPXEST.Update
      Est.MoveNext
      If Est.EOF Then
         Exit Do
      End If
   Loop
   CommitTrans
End If

If FOREST.RecordCount <> 0 Then
   BeginTrans
   Do Until FOREST.EOF
      FOREST.Delete
      FOREST.MoveNext
   Loop
   CommitTrans
End If

    
If DETENT.RecordCount <> 0 Then
   BeginTrans
   Do Until DETENT.EOF
         
   
      GoSub VerificaCamposNulos:
      If DETENT!CodItem <> 0 Then
         EMPXEST.Seek "=", DETENT!CodEmpresa, DETENT!CodItem
         If EMPXEST.NoMatch Then
            EMPXEST.AddNew
            EMPXEST!CodigoEmpresa = DETENT!CodEmpresa
            EMPXEST!CodigoItem = DETENT!CodItem
            EMPXEST.Update
         End If
   
         FOREST.Seek "=", DETENT!CodEmpresa, DETENT!CodItem, DETENT!CodigoFornecedor
         If FOREST.NoMatch Then
            FOREST.AddNew
           Else
            FOREST.Edit
         End If
         FOREST!CodigoEmpresa = DETENT!CodEmpresa
         FOREST!CodigoItem = DETENT!CodItem
         FOREST!CodigoFornecedor = DETENT!CodigoFornecedor
         FOREST.Update
      End If
      DETENT.MoveNext
   Loop
   CommitTrans
End If
DETENT.Close
FOREST.Close
EMPXEST.Close
Est.Close
bd1.Close
Set DETENT = Nothing
Set FOREST = Nothing
Set EMPXEST = Nothing
Set Est = Nothing
Set bd1 = Nothing
Exit Function


VerificaCamposNulos:
  If DETENT!CodEmpresa <> 1 Then
     DETENT.Edit
     DETENT!CodEmpresa = Form_Menu.WEmpresa
     DETENT.Update
  End If
      
  If IsNull(DETENT!ICMS) Then
     DETENT.Edit
     DETENT!ICMS = 0
     DETENT.Update
  End If
  If IsNull(DETENT!IPI) Then
     DETENT.Edit
     DETENT!IPI = 0
     DETENT.Update
  End If
      
  If IsNull(DETENT!Quantidade) Then
     DETENT.Edit
     DETENT!Quantidade = 0
     DETENT.Update
  End If
  If IsNull(DETENT!ValorUnit) Then
     DETENT.Edit
     DETENT!ValorUnit = 0
     DETENT.Update
  End If
  If IsNull(DETENT!VlrFrete) Then
     DETENT.Edit
     DETENT!VlrFrete = 0
     DETENT.Update
  End If
  If IsNull(DETENT!NroCotacao) Then
     DETENT.Edit
     DETENT!NroCotacao = 0
     DETENT.Update
  End If

Return


TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function
Function EMPXEST()
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim EMPEST, Est As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set EMPEST = bd1.OpenRecordset("EMPXEST", DB_OPEN_TABLE)
Set Est = bd1.OpenRecordset("EST", DB_OPEN_TABLE)
EMPEST.index = "Primarykey"
If Est.RecordCount <> 0 Then
   BeginTrans
   While Not Est.EOF
      EMPEST.Seek "=", Form_Menu.WEmpresa, Est!CodItem
      If EMPEST.NoMatch Then
         EMPEST.AddNew
        Else
         EMPEST.Edit
      End If
      EMPEST!CodigoEmpresa = Form_Menu.WEmpresa
      EMPEST!CodigoItem = Est!CodItem
      EMPEST.Update
      Est.MoveNext
   Wend
   CommitTrans
End If
Est.Close
EMPEST.Close
bd1.Close
Set Est = Nothing
Set EMPEST = Nothing
Set bd1 = Nothing
Exit Function

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function SaidasNoMes(CodItem As Long) As Long
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim RQM As Recordset
Dim WSaidas As Long, WData As Variant, WDataA As Variant, WDataB As Variant
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set RQM = bd1.OpenRecordset("RMXEMPXEST", DB_OPEN_TABLE)
RQM.index = "SecondKey"
WDataA = Format$(Date, "mm/yyyy")
WDataB = Format$("01/" & WDataA, "dd/mm/yyyy")
WData = WDataB
RQM.Seek "=", CodItem, WData
If RQM.NoMatch Then
   WSaidas = 0
   RQM.Close
   Exit Function
End If

WSaidas = 0
'BeginTrans
Do Until RQM.EOF
   If Not IsNull(RQM!DataAtendimento) Then
      If Format$(RQM!DataAtendimento, "mm") <> Format$(Date, "mm") Then
         Exit Do
      Else
         If RQM!CodItem <> CodItem Then
            Exit Do
         Else
            If RQM!CodEmpresa = Form_Menu.WEmpresa Then
               WSaidas = WSaidas + (RQM!Quantidade - RQM!Devolucao)
            End If
         End If
     End If
   End If
   RQM.MoveNext
Loop
'CommitTrans
RQM.Close
bd1.Close
Set RQM = Nothing
Set bd1 = Nothing
SaidasNoMes = WSaidas
Exit Function

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function ConsisteLucro(CodItem As Long, CodEmpresa As Long, PrecoDesejado As Double)
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim Prec As Recordset, Ent As Recordset, Cat As Recordset, Est As Recordset
Dim Ref As Long
Dim Result As Long
Dim WPerc As Double
Dim WPrecoCusto As Double
Dim WPrecoMinimo As Variant
Dim WPrecoMinimoStr As String * 10
Dim WPrecoDesejadoStr As String * 10
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Prec = bd1.OpenRecordset("EMPXEST", DB_OPEN_TABLE)
Set Est = bd1.OpenRecordset("EST", DB_OPEN_TABLE)

Prec.index = "PrimaryKey"
Est.index = "PrimaryKey"

WPerc = DescontoMaximo()

ConsisteLucro = 1

Est.Seek "=", CodItem

If Not Est.NoMatch Then
   Prec.Seek "=", Form_Menu.WEmpresa, CodItem
   If Not Prec.NoMatch Then
      
      WPrecoMinimo = Prec!PrecoVenda - ((Prec!PrecoVenda * WPerc) / 100)
      If WPrecoMinimo > PrecoDesejado Then
         ConsisteLucro = 1
         GoSub Finaliza:
      Else
         ConsisteLucro = 0
         GoSub Finaliza:
      End If
        
   End If
       
End If

Finaliza:
  Est.Close
  Prec.Close
  bd1.Close
  Set Est = Nothing
  Set Prec = Nothing
  Set bd1 = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function SenhaResp(WCodigo_Func As Long) As String
On Error GoTo TratamentodeErro:
If Not IsNull(Forms!CadastraOSPrinc!Combina��o193) Then
   Dim bd1 As DataBase
   Dim Sen As Recordset
   Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
   Set Sen = bd1.OpenRecordset("RESP", DB_OPEN_TABLE)
   Sen.index = "PrimaryKey"
   Sen.Seek "=", WCodigo_Func
   'BeginTrans
   If Not Sen.NoMatch Then
      If IsNull(Sen!Sen) Then
         SenhaResp = "LIVRE"
       Else
         SenhaResp = Sen!Sen
      End If
   Else
      SenhaResp = "LIVRE"
   End If
   'CommitTrans
    Sen.Close
    bd1.Close
    Set Sen = Nothing
    Set bd1 = Nothing
End If

Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function


Function AlteraSenhaResp(CodResponsavel As Variant, SenhaAnterior As String, NovaSenha As String, Confirmacao As String)
Dim bd1 As DataBase
Dim Resp As Recordset, Estoque As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Resp = bd1.OpenRecordset("Resp", DB_OPEN_TABLE)
Resp.index = "PrimaryKey"
If Not IsNull(CodResponsavel) Then
   BeginTrans
   Resp.Seek "=", CodResponsavel
   Resp.Edit
   If Len(Resp!Sen & "") > 0 Then
      If SenhaAnterior = Resp!Sen Then
         If NovaSenha = Confirmacao Then
            Resp!Sen = Confirmacao
            GoSub Atualiza:
         Else
            MsgBox "Nova Senha diferente CONFIRMACAO...", 64, "SGa"
            GoSub Atualiza:
            Exit Function
         End If
      Else
         MsgBox "Senha anterior invalida...", 64, "SGa"
         GoSub Atualiza:
         Exit Function
      End If
   Else
      Resp!Sen = Confirmacao
      GoSub Atualiza:
   End If
Else
  Resp.Close
End If

Exit Function

Atualiza:
  Resp.Update
  CommitTrans
  Resp.Close
  bd1.Close
  Set Resp = Nothing
  Set bd1 = Nothing
Return

End Function
Function CapturaSenhaFuncionario(WCodFunc As Long) As String
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim Resp As Recordset, Estoque As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Resp = bd1.OpenRecordset("Resp", DB_OPEN_TABLE)
Resp.index = "PrimaryKey"
Resp.Seek "=", WCodFunc
If Resp.NoMatch Then
   CapturaSenhaFuncionario = "livre"
Else
   If IsNull(Resp!Sen) Then
      CapturaSenhaFuncionario = "livre"
   Else
      CapturaSenhaFuncionario = Resp!Sen
   End If
End If
 
Finaliza:
  Resp.Close
  bd1.Close
  Set Resp = Nothing
  Set bd1 = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaCelularFuncionario(WCodFunc As Long) As String
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim Resp As Recordset, Estoque As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Resp = bd1.OpenRecordset("Resp", DB_OPEN_TABLE)
Resp.index = "PrimaryKey"
Resp.Seek "=", WCodFunc
If Resp.NoMatch Then
   CapturaCelularFuncionario = "Func Inexistente"
Else
   If IsNull(Resp!Nome) Then
      CapturaCelularFuncionario = 0
   Else
      CapturaCelularFuncionario = Resp!Celular
   End If
End If

 
Finaliza:
  Resp.Close
  bd1.Close
  Set Resp = Nothing
  Set bd1 = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function CapturaNomeFuncionario(WCodFunc As Long) As String
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim Resp As Recordset, Estoque As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Resp = bd1.OpenRecordset("Resp", DB_OPEN_TABLE)
Resp.index = "PrimaryKey"
Resp.Seek "=", WCodFunc
If Resp.NoMatch Then
   CapturaNomeFuncionario = "Func Inexistente"
Else
   If IsNull(Resp!Nome) Then
      CapturaNomeFuncionario = "Null"
   Else
      CapturaNomeFuncionario = UCase(Resp!Nome)
   End If
End If

 
Finaliza:
  Resp.Close
  bd1.Close
  Set Resp = Nothing
  Set bd1 = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CriaSaldoInicial(CodigoItem As Long) As Long
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim ItemEstoque, SaldoEstoque As Recordset, Estoque As Recordset, emp As Recordset, Custo As Recordset
Dim DataSaldo As Variant
Dim WEmpresa As Long
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set ItemEstoque = bd1.OpenRecordset("EMPXEST", DB_OPEN_TABLE)
Set SaldoEstoque = bd1.OpenRecordset("SDOEST", DB_OPEN_TABLE)
Set Custo = bd1.OpenRecordset("ESTXCUSTO", DB_OPEN_TABLE)
Set emp = bd1.OpenRecordset("EMPRESA", DB_OPEN_TABLE)
ItemEstoque.index = "PrimaryKey"
SaldoEstoque.index = "PrimaryKey"
Custo.index = "Primarykey"
emp.index = "Primarykey"
WEmpresa = Form_Menu.WEmpresa

If emp.RecordCount <> 0 Then
   While Not emp.EOF
         If emp!CodigoEmpresa = WEmpresa Then
            DataSaldo = #1/2/1999#
            GoSub Cria_ItemEstoque:
            GoSub Cria_Custo:
          ' GoSub Cria_SaldoEstoque:
         End If
         emp.MoveNext
   Wend
End If
        
Finaliza:
    SaldoEstoque.Close
    ItemEstoque.Close
    Custo.Close
    emp.Close
    bd1.Close
    Set SaldoEstoque = Nothing
    Set ItemEstoque = Nothing
    Set Custo = Nothing
    Set emp = Nothing
    Set bd1 = Nothing
    Exit Function
Return
    
Cria_ItemEstoque:
    BeginTrans
    ItemEstoque.Seek "=", emp!CodigoEmpresa, CodigoItem
    If ItemEstoque.NoMatch Then
       ItemEstoque.AddNew
       ItemEstoque!CodigoEmpresa = emp!CodigoEmpresa
       ItemEstoque!CodigoItem = CodigoItem
       ItemEstoque!Localizacao = "ST"
       ItemEstoque!CustoMedio = 0
       ItemEstoque!PrecoVenda = 0.001
       ItemEstoque!IPI = 0
       ItemEstoque!Minimo = 0
       ItemEstoque!Vlr_Out_Desp = 0
       ItemEstoque!Vlr_Fret = 0
       
       ItemEstoque.Update
    End If
    CommitTrans
Return

Cria_Custo:
    BeginTrans
    Custo.Seek "=", CodigoItem, emp!CodigoEmpresa
    If Custo.NoMatch Then
       Custo.AddNew
       Custo!CodItem = CodigoItem
       Custo!CodEmpresa = emp!CodigoEmpresa
       Custo!PrecoCusto = 0
       Custo.Update
    End If
    CommitTrans
Return

'Cria_SaldoEstoque:
'    BeginTrans
'    SaldoEstoque.Seek "=", emp!CodigoEmpresa, CodigoItem, DataSaldo
'    If SaldoEstoque.NoMatch Then
'       SaldoEstoque.AddNew
'       SaldoEstoque!CodigoEmpresa = emp!CodigoEmpresa
'       SaldoEstoque!CodItem = CodigoItem
'       SaldoEstoque!Data = DataSaldo
'       SaldoEstoque!SaldoInicial = 0
'       SaldoEstoque!CustoMedio = 0
'       SaldoEstoque.Update
'    End If
'    CommitTrans
'Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function
Function NomeFantasiaEmpresa() As String
On Error GoTo TratamentodeErro:
Dim Empr As Recordset, bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Empr = bd1.OpenRecordset("Empresa", DB_OPEN_TABLE)

Do Until Empr.EOF
   If Empr!CodigoEmpresa = Form_Menu.WEmpresa Then
       NomeFantasiaEmpresa = Empr!NomeFantasia
   End If
   Empr.MoveNext
Loop
Empr.Close
bd1.Close
Set Empr = Nothing
Set bd1 = Nothing
   Exit Function

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function DiasDecorridos(WData As Variant)
On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim DU As Recordset
    Dim wQtde As Long
    Dim WDataA As Variant, WDataB As Variant
    
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())

    Set DU = bd1.OpenRecordset("DIASUTEIS", DB_OPEN_TABLE)
    DU.index = "PrimaryKey"
    BeginTrans
    wQtde = 0
    WDataA = Format$(WData, "mm/yyyy")
    WDataB = Format$("01/" & WDataA, "dd/mm/yyyy")
    DU.Seek "=", WDataA
    
    If DU.NoMatch Then
       MsgBox "Tabela de DIASUTEIS n�o gerada para este m�s, contacte o suporte", 16, "SGa"
       Exit Function
    Else
       If Format$(WDataA, "mm/yyyy") = Format$(Date, "mm/yyyy") Then
          Do Until DU!DataDiasUteis = Date
             If DU!Diautil <> 0 Then
                wQtde = wQtde + 1
             End If
             DU.MoveNext
          Loop
        Else
          Do Until Format$(DU!DataDiasUteis, "mm/yyyy") <> Format$(WData, "mm/yyyy")
             If DU!Diautil <> 0 Then
                wQtde = wQtde + 1
             End If
             DU.MoveNext
          Loop
        End If
    End If
    CommitTrans
    DiasDecorridos = wQtde
    DU.Close
    bd1.Close
    Set DU = Nothing
    Set bd1 = Nothing
    Exit Function

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function DiasUteis(WData As Variant)
    On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim DU As Recordset
    Dim wQtde As Long, WDataA As Variant, WDataB As Variant
    
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())

    Set DU = bd1.OpenRecordset("DIASUTEIS", DB_OPEN_TABLE)
    DU.index = "PrimaryKey"
'    BeginTrans
    wQtde = 0
    WDataA = Format$(WData, "mm/yyyy")
    WDataB = Format$("01/" & WDataA, "dd/mm/yyyy")
    DU.Seek "=", WDataA
    If DU.NoMatch Then
       MsgBox "Tabela de DIASUTEIS n�o gerada para este m�s, contacte o suporte", 16, "SGa"
       Exit Function
    Else
       Do Until DU.EOF
          If Format$(DU!DataDiasUteis, "mm/yyyy") = Format$(WData, "mm/yyyy") Then
             If DU!Diautil <> 0 Then
                wQtde = wQtde + 1
             End If
          Else
             Exit Do
          End If
          DU.MoveNext
       Loop
    End If
 '   CommitTrans
    DiasUteis = wQtde
    DU.Close
    bd1.Close
    Set DU = Nothing
    Set bd1 = Nothing
    
    Exit Function
    
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function Rodape_Termo_Selagem() As String

On Error GoTo TratamentodeErro:
Dim bd As DataBase, tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("EMPRESA", DB_OPEN_TABLE)
tb.index = "Primarykey"


tb.Seek "=", Form_Menu.WEmpresa

If Not tb.NoMatch Then
   If Not IsNull(tb!RodapeTermoSelagem) Then
      Rodape_Termo_Selagem = tb!RodapeTermoSelagem
   Else
      Rodape_Termo_Selagem = Space(1)
   End If
Else
   Rodape_Termo_Selagem = Space(1)
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function Rodape_Termo_Ensaios() As String
On Error GoTo TratamentodeErro:
Dim bd As DataBase, tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("EMPRESA", DB_OPEN_TABLE)
tb.index = "Primarykey"


tb.Seek "=", Form_Menu.WEmpresa

If Not tb.NoMatch Then
   If Not IsNull(tb!RodapeTermoEnsaios) Then
      Rodape_Termo_Ensaios = tb!RodapeTermoEnsaios
   Else
      Rodape_Termo_Ensaios = Space(1)
   End If
Else
   Rodape_Termo_Ensaios = Space(1)
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function TermoSelagem() As String
On Error GoTo TratamentodeErro:
Dim bd As DataBase, tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("EMPRESA", DB_OPEN_TABLE)
tb.index = "Primarykey"


tb.Seek "=", Form_Menu.WEmpresa

If Not tb.NoMatch Then
   If Not IsNull(tb!TermoSelagem) Then
      TermoSelagem = tb!TermoSelagem
   Else
      TermoSelagem = Space(1)
   End If
Else
   TermoSelagem = Space(1)
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Termoensaios() As String

On Error GoTo TratamentodeErro:
Dim bd As DataBase, tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("EMPRESA", DB_OPEN_TABLE)
tb.index = "Primarykey"

tb.Seek "=", Form_Menu.WEmpresa

If Not tb.NoMatch Then
   If Not IsNull(tb!Termoensaios) Then
      Termoensaios = tb!Termoensaios
   Else
      Termoensaios = Space(1)
   End If
Else
   Termoensaios = Space(1)
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function PreencheZeros(ByVal Campo As String, comprimento As Integer) As String
    Dim X, Tam, Comp As Integer               '**********************************'
    Dim Zeros1 As String * 14                 '* Fun��o que insere zeros        *'
    For X = 1 To Len(Campo)                   '* na posi��o menos significativa *'
        Mid$(Zeros1, X, X) = "0"              '* de uma string num�rica para    *'
    Next X                                    '* ser processada no ECF          *'
    For X = 1 To Len(Campo)                   '**********************************'
        If Mid$(Campo, X, 1) <> " " Then
           Tam = Tam + 1
           Else
           X = comprimento
        End If
    Next X
    Comp = comprimento - Tam
    Campo = Mid$(Zeros1, 1, Comp) + Campo
    PreencheZeros = Campo
    'MsgBox Campo
End Function

Function PreencheZeros_MFE(WCampo As String, Wcomprimento As Integer) As String
    Dim X, Tam, Comp As Integer               '**********************************'
    Dim WNovo_Campo As String
    Dim WZeros As String
    Tam = Len(Trim(WCampo))
    
    For X = 1 To Wcomprimento - Tam
        WZeros = WZeros & "0"
     
    Next X
    'MsgBox "" & WZeros
    'MsgBox "" & Len(WZeros)
    
    PreencheZeros_MFE = Trim(WZeros & WCampo)
    'MsgBox Campo
End Function
Function MascaraCgcCpf(ByVal Campo As String, TipoCli As String) As String
                                               '11.111.111/1111-11
                                               '12.345.678/9012-34
    Dim Ident As String
    
    If TipoCli = 1 Then
       Ident = Mid$(Campo, 1, 3) + "." + Mid$(Campo, 4, 3) + "."
       Ident = Ident + Mid$(Campo, 7, 3) + "-" + Mid$(Campo, 10, 2)
    Else
       Ident = Mid$(Campo, 1, 2) + "." + Mid$(Campo, 3, 3) + "."
       Ident = Ident + Mid$(Campo, 6, 3) + "/" + Mid$(Campo, 9, 4)
       Ident = Ident + "-" + Mid$(Campo, 13, 2)
    End If
    MascaraCgcCpf = Ident
End Function
Function DataPagtoFormAtivo() As Variant
    If Est�Carregado("BaixaAutContasPagar") Then
       DataPagtoFormAtivo = "31/12/1998"
      Else
       DataPagtoFormAtivo = Date + 3650
       End If
End Function
Function DataVencFormAtivo() As Variant
    If Est�Carregado("ManutContasPagar") Then
       DataVencFormAtivo = Date + 3650
       Else
       DataVencFormAtivo = Date + 1
    End If
    If Est�Carregado("CadastroFornecedores") Then
       DataVencFormAtivo = Date - 1
    End If
End Function
Function NumOsFormAtivo() As Long
    If Est�Carregado("PesquisarOs") Then
       NumOsFormAtivo = Forms!PesquisarOS!Numos
    End If
    If Est�Carregado("CadastraOsPrinc") Then
       NumOsFormAtivo = Forms!CadastraOSPrinc!Numos
    End If
End Function

Function Est�Carregado(strNomeDoFrm As String) As Boolean
   
    Const conEstruturaDoForm = 0
    Dim intX As Integer
    
    Est�Carregado = False
    For intX = 0 To Forms.Count - 1
        If Forms(intX).FormName = strNomeDoFrm Then
            If Forms(intX).CurrentView <> conEstruturaDoForm Then
                Est�Carregado = True
                Exit Function  ' Sai da fun��o, uma vez localizado o formul�rio.
            End If
        End If
    Next

End Function

Function CapturaCodigoClienteXOSV(WNumos As Long) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", WNumos
If Not tb.NoMatch Then
   If Not IsNull(tb!CodigoCliente) Then
      CapturaCodigoClienteXOSV = tb!CodigoCliente
   Else
      CapturaCodigoClienteXOSV = 0
   End If
Else
   CapturaCodigoClienteXOSV = 0
End If
tb.Close
bd.Close
Set tb = Nothing
Set bd = Nothing
Exit Function

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CapturaCodigoPagtoXOSV(WNumos As Long) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", WNumos
If Not tb.NoMatch Then
   If Not IsNull(tb!CodigoPagto) Then
      CapturaCodigoPagtoXOSV = tb!CodigoPagto
   Else
      CapturaCodigoPagtoXOSV = 0
   End If
Else
   CapturaCodigoPagtoXOSV = 0
End If
tb.Close
bd.Close
Set tb = Nothing
Set bd = Nothing
Exit Function

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function RegravaOsFechamento(Numos As Long, WData As Variant) As Long
    On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim OrdemServico As Recordset
    
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set OrdemServico = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
    OrdemServico.index = "PrimaryKey"
    BeginTrans
    OrdemServico.Seek "=", Numos
    If Not OrdemServico.NoMatch Then
       OrdemServico.Edit
       OrdemServico!datafechamento = WData
       OrdemServico!HoraFechamento = Time
       OrdemServico.Update
    End If

Finaliza:
    CommitTrans
    OrdemServico.Close
    bd1.Close
    Set OrdemServico = Nothing
    Set bd1 = Nothing
    Exit Function
Return
    
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function RegravaOsvTipoPistaEnsaio(Numos As Long, WTipoPista As String)
    On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim OrdemServico As Recordset
    
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set OrdemServico = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
    OrdemServico.index = "PrimaryKey"
    BeginTrans
    OrdemServico.Seek "=", Numos
    If Not OrdemServico.NoMatch Then
       OrdemServico.Edit
       OrdemServico!TipoPistaEnsaio = WTipoPista
       OrdemServico.Update
    End If

Finaliza:
    CommitTrans
    OrdemServico.Close
    bd1.Close
    Set OrdemServico = Nothing
    Set bd1 = Nothing
    Exit Function
Return
    
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function RegravaOSVDataEnsaio(Numos As Long)
    On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim OrdemServico As Recordset
    
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set OrdemServico = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
    OrdemServico.index = "PrimaryKey"
    BeginTrans
    OrdemServico.Seek "=", Numos
    If Not OrdemServico.NoMatch Then
       OrdemServico.Edit
       OrdemServico!DTEnsaio = Date
       OrdemServico.Update
    End If

Finaliza:
    CommitTrans
    OrdemServico.Close
    bd1.Close
    Set OrdemServico = Nothing
    Set bd1 = Nothing
    Exit Function
Return
    
    
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function RegravaOSCupon(Numos As Long)
On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim OS As Recordset
    
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set OS = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
    OS.index = "PrimaryKey"
    BeginTrans
    OS.Seek "=", Numos
    If Not OS.NoMatch Then
       OS.Edit
       OS!DataEmisCupon = Date
       OS.Update
    End If
    CommitTrans
    OS.Close
    bd1.Close
    Set OS = Nothing
    Set bd1 = Nothing
    Exit Function
    
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

    
End Function

Function Retorna_Status_OSV(WStatus As String) As String

If IsNull(WStatus) Then
   Retorna_Status_OSV = Space(1)
ElseIf WStatus = "AN" Then
       Retorna_Status_OSV = "Analise"
ElseIf WStatus = "AP" Then
       Retorna_Status_OSV = "Aprovado"
ElseIf WStatus = "RE" Then
    Retorna_Status_OSV = "Reprovado"
Else
    Retorna_Status_OSV = Space(1)
End If

End Function

Function GeraComparativoAnualVendas(WAno_1 As Long, WAno_2 As Long)
DoCmd.Hourglass True
Dim bd1 As DataBase, bd As DataBase
Dim QOSV As QueryDef, OSV As Recordset, CompAno1 As Recordset, CompAno2 As Recordset
Dim wmesano As Variant, Wmes As Variant, Ret As Variant, X As Long, numreg As Long
Dim WAno As Variant, WTotProd As Double, WTotDescProd As Double, WTotServ As Double, WTotDescServ As Double
Set bd = CurrentDb()
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set QOSV = bd.QueryDefs("ComparativoVendasAnual")
'QOSV.Parameters(0) = Forms!ComparativoAnualVendas![Ano-1]
'QOSV.Parameters(1) = Forms!ComparativoAnualVendas![Ano-2]
QOSV.Parameters(0) = WAno_1
QOSV.Parameters(1) = WAno_2
'QOSV.Parameters(2) = Forms!ComparativoAnualVendas![MovimentaFinanceiro]
Set OSV = QOSV.OpenRecordset()
Set CompAno1 = bd1.OpenRecordset("ComparativoVendasAno-1", DB_OPEN_TABLE)
Set CompAno2 = bd1.OpenRecordset("ComparativoVendasAno-2", DB_OPEN_TABLE)
CompAno1.index = "PrimaryKey"
CompAno2.index = "PrimaryKey"

If CompAno1.RecordCount <> 0 Then
   BeginTrans
   While Not CompAno1.EOF
      CompAno1.Delete
      CompAno1.MoveNext
   Wend
   CommitTrans
End If

If CompAno2.RecordCount <> 0 Then
   BeginTrans
   While Not CompAno2.EOF
      CompAno2.Delete
      CompAno2.MoveNext
   Wend
   CommitTrans
End If

If OSV.RecordCount <> 0 Then
   OSV.MoveLast
   OSV.MoveFirst
   numreg = OSV.RecordCount
   Ret = SysCmd(acSysCmdInitMeter, "Aguarde, gerando informacoes...", numreg)
   X = 1
   BeginTrans
   While Not OSV.EOF
    
         If Not IsNull(OSV!datafechamento) Then
            WAno = Format$(OSV!datafechamento, "yyyy")
            wmesano = Format$(OSV!datafechamento, "mm/yyyy")
            Wmes = Format$(OSV!datafechamento, "mm")
            If WAno = Forms!ComparativoAnualVendas![Ano-1] Then
               GoSub ApuraValorOSV:
               GoSub GravaCompAno1:
            ElseIf WAno = Forms!ComparativoAnualVendas![Ano-2] Then
                   GoSub ApuraValorOSV:
                   GoSub GravaCompAno2:
            End If
         End If
      
         Ret = SysCmd(acSysCmdUpdateMeter, X)
         X = X + 1
         If X Mod 500 = 0 Then
            CommitTrans
            BeginTrans
         End If
         OSV.MoveNext
   Wend
   CommitTrans
   
End If

Finaliza:
  OSV.Close
  CompAno1.Close
  CompAno2.Close
  bd.Close
  bd1.Close
  Set OSV = Nothing
  Set CompAno1 = Nothing
  Set CompAno2 = Nothing
  Ret = SysCmd(acSysCmdRemoveMeter)
  DoCmd.Hourglass False
 ' Call Rodape_SGa
  Exit Function
Return

GravaCompAno1:
 CompAno1.Seek "=", wmesano
 If CompAno1.NoMatch Then
    CompAno1.AddNew
 Else
    CompAno1.Edit
 End If
 CompAno1![MesAno-1] = wmesano
 CompAno1![Mes-1] = Wmes
 CompAno1![TotalProdutos-1] = CompAno1![TotalProdutos-1] + WTotProd
 CompAno1![TotalDescontoProdutos-1] = CompAno1![TotalDescontoProdutos-1] + WTotDescProd
 CompAno1![TotalServicos-1] = CompAno1![TotalServicos-1] + WTotServ
 CompAno1![TotalDescontoServicos-1] = CompAno1![TotalDescontoServicos-1] + WTotDescServ
 CompAno1.Update
Return

GravaCompAno2:
 CompAno2.Seek "=", wmesano
 If CompAno2.NoMatch Then
    CompAno2.AddNew
 Else
    CompAno2.Edit
 End If
 CompAno2![MesAno-2] = wmesano
 CompAno2![Mes-2] = Wmes
 CompAno2![TotalProdutos-2] = CompAno2![TotalProdutos-2] + WTotProd
 CompAno2![TotalDescontoProdutos-2] = CompAno2![TotalDescontoProdutos-2] + WTotDescProd
 CompAno2![TotalServicos-2] = CompAno2![TotalServicos-2] + WTotServ
 CompAno2![TotalDescontoServicos-2] = CompAno2![TotalDescontoServicos-2] + WTotDescServ
 CompAno2.Update
Return

ApuraValorOSV:
 WTotProd = 0
 WTotDescProd = 0
 WTotServ = 0
 WTotDescServ = 0
 If Not IsNull(OSV!tot_itens) Then
    WTotProd = OSV!tot_itens
 End If
 WTotDescProd = 0
 If Not IsNull(OSV!Tot_serv) Then
    WTotServ = OSV!Tot_serv
 End If
 WTotDescServ = 0
Return
    
    
End Function

Function GeraGraficoComparativoVendasAnual()
DoCmd.Hourglass True
Dim bd1 As DataBase, bd As DataBase
Dim QOSV As QueryDef, OSV As Recordset, CompAno1 As Recordset, CompAno2 As Recordset
Dim wmesano As Variant, Wmes As Variant, Ret As Variant, X As Long, numreg As Long
Dim WAno As Variant, WTotProd As Double, WTotDescProd As Double, WTotServ As Double, WTotDescServ As Double
Set bd = CurrentDb()
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set QOSV = bd.QueryDefs("GeraComparativoVendasAnual")
QOSV.Parameters(0) = Forms!GraficoComparativoVendasAnual![Ano-1]
QOSV.Parameters(1) = Forms!GraficoComparativoVendasAnual![Ano-2]
Set OSV = QOSV.OpenRecordset()
Set CompAno1 = bd1.OpenRecordset("ComparativoVendasAno-1", DB_OPEN_TABLE)
Set CompAno2 = bd1.OpenRecordset("ComparativoVendasAno-2", DB_OPEN_TABLE)
CompAno1.index = "PrimaryKey"
CompAno2.index = "PrimaryKey"

If CompAno1.RecordCount <> 0 Then
   BeginTrans
   Do Until CompAno1.EOF
      CompAno1.Delete
      CompAno1.MoveNext
   Loop
   CommitTrans
End If

If CompAno2.RecordCount <> 0 Then
   BeginTrans
   Do Until CompAno2.EOF
      CompAno2.Delete
      CompAno2.MoveNext
   Loop
   CommitTrans
End If

If OSV.RecordCount <> 0 Then
   OSV.MoveLast
   OSV.MoveFirst
   numreg = OSV.RecordCount
   Ret = SysCmd(acSysCmdInitMeter, "Aguarde, gerando informacoes...", numreg)
   X = 1
   BeginTrans
   Do Until OSV.EOF
   
      If Not IsNull(OSV!datafechamento) Then
         WAno = Format$(OSV!datafechamento, "yyyy")
         wmesano = Format$(OSV!datafechamento, "mm/yyyy")
         Wmes = Format$(OSV!datafechamento, "mm")
         If WAno = Forms!GraficoComparativoVendasAnual![Ano-1] Then
            GoSub ApuraValorOSV:
            GoSub GravaCompAno1:
         ElseIf WAno = Forms!GraficoComparativoVendasAnual![Ano-2] Then
            GoSub ApuraValorOSV:
            GoSub GravaCompAno2:
         End If
      End If
      
      Ret = SysCmd(acSysCmdUpdateMeter, X)
      X = X + 1
      If X Mod 500 = 0 Then
         CommitTrans
         BeginTrans
      End If
      OSV.MoveNext
   Loop
   CommitTrans
   
End If
OSV.Close
CompAno1.Close
CompAno2.Close
bd.Close
bd1.Close
Set OSV = Nothing
Set CompAno1 = Nothing
Set CompAno2 = Nothing
Ret = SysCmd(acSysCmdRemoveMeter)
DoCmd.Hourglass False
Exit Function

GravaCompAno1:
 CompAno1.Seek "=", wmesano
 If CompAno1.NoMatch Then
    CompAno1.AddNew
 Else
    CompAno1.Edit
 End If
 CompAno1![MesAno-1] = wmesano
 CompAno1![Mes-1] = Wmes
 CompAno1![TotalProdutos-1] = CompAno1![TotalProdutos-1] + WTotProd
 CompAno1![TotalDescontoProdutos-1] = CompAno1![TotalDescontoProdutos-1] + WTotDescProd
 CompAno1![TotalServicos-1] = CompAno1![TotalServicos-1] + WTotServ
 CompAno1![TotalDescontoServicos-1] = CompAno1![TotalDescontoServicos-1] + WTotDescServ
 CompAno1.Update
Return

GravaCompAno2:
 CompAno2.Seek "=", wmesano
 If CompAno2.NoMatch Then
    CompAno2.AddNew
 Else
    CompAno2.Edit
 End If
 CompAno2![MesAno-2] = wmesano
 CompAno2![Mes-2] = Wmes
 CompAno2![TotalProdutos-2] = CompAno2![TotalProdutos-2] + WTotProd
 CompAno2![TotalDescontoProdutos-2] = CompAno2![TotalDescontoProdutos-2] + WTotDescProd
 CompAno2![TotalServicos-2] = CompAno2![TotalServicos-2] + WTotServ
 CompAno2![TotalDescontoServicos-2] = CompAno2![TotalDescontoServicos-2] + WTotDescServ
 CompAno2.Update
Return

ApuraValorOSV:
 WTotProd = 0
 WTotDescProd = 0
 WTotServ = 0
 WTotDescServ = 0
 WTotProd = ResumoFinanceiroMercadorias(OSV!Numos)
 WTotDescProd = CalculoDescontoMercadorias(OSV!Numos)
 WTotServ = ResumoFinanceiroServico(OSV!Numos)
 WTotDescServ = TotalDescontosServicosOs(OSV!Numos)
Return
    
    
End Function



Function EliminaEspacos(ByVal WString As String)
On Error GoTo TratamentodeErro:
Dim I As Long
Dim WC As String * 1
Dim CaracConv As String
For I = 1 To Len(WString)
    If Mid$(WString, I, 1) <> " " Then
       WC = Mid$(WString, I, 1)
      CaracConv = CaracConv + WC
    End If
Next I

Finaliza:
  EliminaEspacos = CaracConv
  Exit Function
Return
    
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function
Function PreencheEspacos(ByVal Campo As String, comprimento As Integer) As String
    Dim X, Tam, Comp As Integer
    For X = 1 To Len(Campo) - 4
        If Mid$(Campo, X, 1) = "0" Or Mid$(Campo, X, 1) = "." Then
           Mid$(Campo, X, 1) = " "
           Else
           comprimento = X
           X = Len(Campo)
           End If
    Next X
    PreencheEspacos = Campo
           
End Function
Function CriaTabelaEstoqueMinimo()
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase, bd As DataBase, QCons As QueryDef, Cons As Recordset, ESTM As Recordset
Dim wsaldo As Long, WSaidas As Long, WMedia As Long, WTendencia As Long
Dim WDU As Integer, WDD As Integer, WDR As Integer
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set bd = CurrentDb()
Set ESTM = bd1.OpenRecordset("EstM", DB_OPEN_TABLE)
Set QCons = bd.QueryDefs("EstoquePontoPedido")
QCons.Parameters(0) = Forms![EstoquePontoPedido]!categoriadesejada
QCons.Parameters(1) = Forms![EstoquePontoPedido]!ProdutoDesejado
Set Cons = QCons.OpenRecordset()
ESTM.index = "PrimaryKey"

If ESTM.RecordCount <> 0 Then
   BeginTrans
   Do Until ESTM.EOF
      ESTM.Delete
      ESTM.MoveNext
   Loop
   CommitTrans
End If

WDU = DiasUteis(Date)
WDD = DiasDecorridos(Date)
WDR = (WDU - WDD)

If Cons.RecordCount <> 0 Then
   BeginTrans
   Do Until Cons.EOF
      wsaldo = Saldo_Estoque(Cons!CodigoItem, Date, Empresa_Ativa())
      WSaidas = SaidasNoMes(Cons!CodigoItem)
      WMedia = Int(WSaidas / WDD)
      WTendencia = (WMedia * WDU)
      If wsaldo <= Cons!Minimo Then
          ESTM.Seek "=", Cons!CodigoItem
          If ESTM.NoMatch Then
             ESTM.AddNew
          Else
             ESTM.Edit
          End If
          ESTM!CodigoItem = Cons!CodigoItem
          ESTM!Minimo = Cons!Minimo
          ESTM!EstoqueAtual = wsaldo
          ESTM!SaidasNoMes = WSaidas
          ESTM!MediaDia = WMedia
          ESTM!DU = WDU
          ESTM!DD = WDD
          ESTM!dr = WDR
          
          If WMedia <= 0 Then
             ESTM!TendenciaMes = 1
          Else
             ESTM!TendenciaMes = WMedia * WDU
          End If
              
          ESTM!CustoMedio = Cons!CustoMedio
          ESTM.Update
        End If
      Cons.MoveNext
   Loop
   CommitTrans
End If
Cons.Close
ESTM.Close
bd.Close
bd1.Close
Set Cons = Nothing
Set ESTM = Nothing
Set bd = Nothing
Set bd1 = Nothing
Exit Function

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function TotalServicosComissionadoOrdemServico(Numos As Long) As Double
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim ItensServicos, TPS, tbo As Recordset
    
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set ItensServicos = bd1.OpenRecordset("ITENSSERVOSV", DB_OPEN_TABLE)
Set TPS = bd1.OpenRecordset("TPSERV", DB_OPEN_TABLE)
Set tbo = bd1.OpenRecordset("OCORR_SERV", DB_OPEN_TABLE)
    
Dim NUMOSATUAL, highvalue  As Long
Dim SomaTotal, ValorDesconto As Double
Dim Quantidade As Double
    
ItensServicos.index = "SecondaryKey"
TPS.index = "PrimaryKey"
tbo.index = "PrimaryKey"
SomaTotal = 0
'BeginTrans
ItensServicos.Seek "=", Numos
If ItensServicos.NoMatch Then
   SomaTotal = 0
 '  CommitTrans
   ItensServicos.Close
   bd1.Close
   Exit Function
Else
   While ItensServicos!Numos = Numos
      
         GoSub VerificaCamposNulos:
         tbo.Seek "=", ItensServicos!Ocorr
         If Not tbo.NoMatch Then
            If tbo!Gera_Financeiro Then
               If ItensServicos!Quantidade > 0 And ItensServicos!ValorMaoObra > 0 Then
                  GoSub SomaServicos
               End If
            End If
         End If
         ItensServicos.MoveNext
         If ItensServicos.EOF Then
            GoSub Finaliza:
         End If
   Wend
End If

Finaliza:
   TotalServicosComissionadoOrdemServico = SomaTotal
   ItensServicos.Close
   tbo.Close
   bd1.Close
   Set ItensServicos = Nothing
   Set tbo = Nothing
   Set bd1 = Nothing
   Exit Function
Return
    
SomaServicos:
    TPS.Seek "=", ItensServicos!TipoServico
    If Not TPS.NoMatch Then
       If ItensServicos!Comissionado Then
          Quantidade = ItensServicos!Quantidade
          ValorDesconto = ItensServicos!Desconto
          SomaTotal = SomaTotal + ((Quantidade * ItensServicos!ValorMaoObra) - (Quantidade * ValorDesconto))
       End If
    End If
Return

VerificaCamposNulos:
  BeginTrans
  If IsNull(ItensServicos!Quantidade) Then
     ItensServicos.Edit
     ItensServicos!Quantidade = 0
     ItensServicos.Update
  End If
  If IsNull(ItensServicos!ValorMaoObra) Then
     ItensServicos.Edit
     ItensServicos!ValorMaoObra = 0
     ItensServicos.Update
  End If
  If IsNull(ItensServicos!Desconto) Then
     ItensServicos.Edit
     ItensServicos!Desconto = 0
     ItensServicos.Update
  End If
  CommitTrans
Return


TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function



Function ResumoFinanceiroServico(Numos As Long) As Double
Dim bd1 As DataBase
    Dim ItensServicos, tbo As Recordset
    
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set ItensServicos = bd1.OpenRecordset("ITENSSERVOSV", DB_OPEN_TABLE)
    Set tbo = bd1.OpenRecordset("OCORR_SERV", DB_OPEN_TABLE)
    
    Dim NUMOSATUAL, highvalue  As Long
    Dim SomaTotal, ValorDesconto As Double
    Dim Quantidade As Long
    
    ItensServicos.index = "SecondaryKey"
    tbo.index = "PrimaryKey"
    highvalue = 999999999
    SomaTotal = 0
    NUMOSATUAL = highvalue
    
    ItensServicos.Seek "=", Numos
    If ItensServicos.NoMatch Then
       NUMOSATUAL = highvalue
     Else
       NUMOSATUAL = ItensServicos!Numos
    End If
    
    While NUMOSATUAL = Numos
          If ItensServicos!Quantidade > 0 And ItensServicos!ValorMaoObra > 0 Then
             tbo.Seek "=", ItensServicos!Ocorr
             If Not tbo.NoMatch Then
                If tbo!Gera_Financeiro Then
                   GoSub SomaServicos
                End If
             End If
          End If
          ItensServicos.MoveNext
          If ItensServicos.EOF Then
             NUMOSATUAL = highvalue
           Else
             NUMOSATUAL = ItensServicos!Numos
          End If
    Wend
    
    ResumoFinanceiroServico = SomaTotal
    ItensServicos.Close
    tbo.Close
    bd1.Close
    Set ItensServicos = Nothing
    Set tbo = Nothing
    Set bd1 = Nothing
    Exit Function
    
SomaServicos:
    Quantidade = ItensServicos!Quantidade
    ValorDesconto = ItensServicos!Desconto
    SomaTotal = SomaTotal + (Quantidade * ItensServicos!ValorMaoObra)
    Return
End Function


Function Juros(Numos As Double, CodigoPagto As Double)

' esta funcao foi descontinuada pois o conceito do campo delta  � desconto e nao juros
GoSub Finaliza:

On Error GoTo TratamentodeErro:
If CodigoPagto <> 0 Or Not IsNull(CodigoPagto) Then
   Dim bd1 As DataBase
   Dim RQM As Recordset
   Dim RM As Recordset
   Dim PRZ As Recordset
   Dim Est As Recordset
   Dim Cli As Recordset
   Dim TABC As Recordset
   Dim WDelta As Double
   
   Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
   Set RQM = bd1.OpenRecordset("RQM", DB_OPEN_TABLE)
   Set RM = bd1.OpenRecordset("RMXEMPXEST", DB_OPEN_TABLE)
   Set PRZ = bd1.OpenRecordset("CONDPAGTO", DB_OPEN_TABLE)
   Set Est = bd1.OpenRecordset("EMPXEST", DB_OPEN_TABLE)
   Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
   Set TABC = bd1.OpenRecordset("TAB", DB_OPEN_TABLE)
   Cli.index = "Primarykey"
   TABC.index = "PrimaryKey"
   Est.index = "PrimaryKey"
   RQM.index = "SecondaryKey"
   RM.index = "RQM"
   PRZ.index = "PrimaryKey"
   
   BeginTrans
   PRZ.Seek "=", CodigoPagto
   If PRZ.NoMatch Then
      BeginTrans
      GoSub Finaliza:
   End If
   RQM.Seek "=", Numos
   If Not RQM.NoMatch Then
      Do Until RQM.EOF
         If RQM!Numos <> Numos Then
            Exit Do
         Else
            RM.Seek "=", RQM!NumRequisicao
            If Not RM.NoMatch Then
               Do Until RM.EOF
                  If RM!NumRequisicao <> RQM!NumRequisicao Then
                     Exit Do
                  Else
                     GoSub VerificaDelta:
                     Est.Seek "=", RM!CodEmpresa, RM!CodItem
                     If Not Est.NoMatch Then
                        RM.Edit
                        RM!ValorUnit = RM!ValorUnit + Round((RM!ValorUnit * WDelta) / 100)
                        RM.Update
                     End If
                  End If
                  RM.MoveNext
               Loop
            End If
         End If
         RQM.MoveNext
      Loop
   End If
   CommitTrans
End If

Finaliza:
  PRZ.Close
  RQM.Close
  RM.Close
  Est.Close
  Cli.Close
  TABC.Close
  bd1.Close
  Set PRZ = Nothing
  Set RQM = Nothing
  Set RM = Nothing
  Set Est = Nothing
  Set Est = Nothing
  Set Cli = Nothing
  Set TABC = Nothing
  Set bd1 = Nothing
  
  Exit Function
Return

VerificaDelta:
   Cli.Seek "=", Forms!CadastraOSPrinc!CodigoCliente
   If Cli.NoMatch Then
      MsgBox "Cliente inexistente...", 64, "SGa"
      Cli.Close
      TABC.Close
      Exit Function
    Else
      TABC.Seek "=", Cli!tABELA
      If TABC.NoMatch Then
         MsgBox "Tabela inexistente", 64, "SGa"
         Cli.Close
         TABC.Close
         Exit Function
       Else
         WDelta = TABC!Delta
      End If
    End If
Return:


TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function


Function CalculoMercadoriasXCat(WNumos As Long, WCategoria As Long) As Double

   On Error GoTo TratamentodeErro:

    Dim bd1 As DataBase
    Dim Requisicao, saidas, Est As Recordset
    
    Dim SomaTotal As Double
    Dim NumOSV As Double
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set Requisicao = bd1.OpenRecordset("RQM", DB_OPEN_TABLE)
    Set saidas = bd1.OpenRecordset("RMXEMPXEST", DB_OPEN_TABLE)
    Set Est = bd1.OpenRecordset("EST", DB_OPEN_TABLE)
    
    Requisicao.index = "SecondaryKey"
    Est.index = "PrimaryKey"
    saidas.index = "RQM"
    SomaTotal = 0
    NumOSV = WNumos
    
    Requisicao.Seek "=", WNumos
    If Not Requisicao.NoMatch Then
       While Requisicao!Numos = WNumos
             GoSub ApuraSaidas:
             Requisicao.MoveNext
             If Requisicao.EOF Then
                GoSub Finaliza:
             End If
       Wend
    End If
    
Finaliza:
    CalculoMercadoriasXCat = SomaTotal
    saidas.Close
    Requisicao.Close
    bd1.Close
    Set saidas = Nothing
    Set Requisicao = Nothing
    Set bd1 = Nothing
    Exit Function
Return

ApuraSaidas:
  saidas.Seek "=", Requisicao!NumRequisicao
  If Not saidas.NoMatch Then
     While saidas!NumRequisicao = Requisicao!NumRequisicao
           GoSub VerificaCamposNulos:
           Est.Seek "=", saidas!CodItem
           If Not Est.NoMatch Then
              If Est!CodigoCategoria = WCategoria Then
                 SomaTotal = SomaTotal + ((saidas!ValorUnit + saidas!Vlr_Icms_ST + saidas!Vlr_Seguro + saidas!Vlr_Out_Desp + saidas!Vlr_Fret + saidas!ValorIPI - saidas!Desconto) * (saidas!Quantidade - saidas!Devolucao))
              End If
           End If
           saidas.MoveNext
           If saidas.EOF Then
              Return
           End If
      Wend
   End If
Return


VerificaCamposNulos:
  BeginTrans
  If IsNull(saidas!ValorUnit) Then
     saidas.Edit
     saidas!ValorUnit = 0
     saidas.Update
  End If
  If IsNull(saidas!Desconto) Then
     saidas.Edit
     saidas!Desconto = 0
     saidas.Update
  End If

  If IsNull(saidas!Quantidade) Then
     saidas.Edit
     saidas!Quantidade = 0
     saidas.Update
  End If
  If IsNull(saidas!Devolucao) Then
     saidas.Edit
     saidas!Devolucao = 0
     saidas.Update
  End If
  CommitTrans
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
    
End Function

Function Total_OSV_Serv_Sel(Numos As Long) As Double
 On Error GoTo TratamentodeErro:

    Dim bd As DataBase
    Dim tbs, tbt, tbo As Recordset
    
    Dim SomaTotal_Est As Double
    Dim SomaTotal_Serv As Double
    Dim NumOSV As Double
    Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tbs = bd.OpenRecordset("ITENSSERVOSV", DB_OPEN_TABLE)
    Set tbt = bd.OpenRecordset("TPSERV", DB_OPEN_TABLE)
    Set tbo = bd.OpenRecordset("OCORR_SERV", DB_OPEN_TABLE)
    
    tbt.index = "PrimaryKey"
    tbo.index = "PrimaryKey"
    tbs.index = "SecondaryKey"
    
    SomaTotal_Serv = 0
    NumOSV = Numos
    
    tbs.Seek "=", Numos
    If Not tbs.NoMatch Then
       While tbs!Numos = Numos
             tbt.Seek "=", tbs!TipoServico
             If Not tbt.NoMatch Then
                If tbt!Selecionar Then
                   tbo.Seek "=", tbs!Ocorr
                   If Not tbo.NoMatch Then
                      If tbo!Gera_Financeiro Then
                         SomaTotal_Serv = SomaTotal_Serv + ((tbs!Quantidade * tbs!ValorMaoObra) - (tbs!Quantidade * tbs!Desconto))
                      End If
                   End If
                End If
             End If
             tbs.MoveNext
             If tbs.EOF Then
                GoSub Finaliza:
             End If
       Wend
    End If
    
Finaliza:
    Total_OSV_Serv_Sel = SomaTotal_Serv
    tbs.Close
    tbt.Close
    tbo.Close
    bd.Close
    Set tbs = Nothing
    Set tbt = Nothing
    Set tbo = Nothing
    Set bd = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Total_OSV_Cat_Est_Serv_Sel(Numos As Long) As Double
    On Error GoTo TratamentodeErro:

    Dim bd As DataBase
    Dim tbr, tbs, tbt, tbi, tbc, tbe, tbo As Recordset
    
    Dim SomaTotal_Est As Double
    Dim SomaTotal_Serv As Double
    Dim NumOSV As Double
    Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tbr = bd.OpenRecordset("RQM", DB_OPEN_TABLE)
    Set tbi = bd.OpenRecordset("RMXEMPXEST", DB_OPEN_TABLE)
    Set tbc = bd.OpenRecordset("CATPROD", DB_OPEN_TABLE)
    Set tbe = bd.OpenRecordset("EST", DB_OPEN_TABLE)
    Set tbs = bd.OpenRecordset("ITENSSERVOSV", DB_OPEN_TABLE)
    Set tbt = bd.OpenRecordset("TPSERV", DB_OPEN_TABLE)
    Set tbo = bd.OpenRecordset("OCORR_SERV", DB_OPEN_TABLE)
    tbr.index = "SecondaryKey"
    tbi.index = "RQM"
    tbc.index = "PrimaryKey"
    tbe.index = "PrimaryKey"
    tbt.index = "PrimaryKey"
    tbo.index = "PrimaryKey"
    tbs.index = "SecondaryKey"
    
    SomaTotal_Est = 0
    SomaTotal_Serv = 0
    NumOSV = Numos
    
    GoSub Apura_Total_Serv:
    
    tbr.Seek "=", Numos
    If Not tbr.NoMatch Then
       While tbr!Numos = Numos
             GoSub Apura_Total_Est:
             tbr.MoveNext
             If tbr.EOF Then
                GoSub Finaliza:
             End If
       Wend
    End If
    
Finaliza:
    Total_OSV_Cat_Est_Serv_Sel = SomaTotal_Serv + SomaTotal_Est
    tbi.Close
    tbr.Close
    tbc.Close
    tbe.Close
    tbs.Close
    tbt.Close
    tbo.Close
    bd.Close
    Set tbi = Nothing
    Set tbr = Nothing
    Set tbc = Nothing
    Set tbe = Nothing
    Set tbs = Nothing
    Set tbt = Nothing
    Set tbo = Nothing
    Set bd = Nothing
    Exit Function
Return

Apura_Total_Serv:
    tbs.Seek "=", Numos
    If Not tbs.NoMatch Then
       While tbs!Numos = Numos
       
             tbt.Seek "=", tbs!TipoServico
             If Not tbt.NoMatch Then
                If tbt!Selecionar Then
                   tbo.Seek "=", tbs!Ocorr
                   If Not tbo.NoMatch Then
                      If tbo!Gera_Financeiro Then
                         SomaTotal_Serv = SomaTotal_Serv + ((tbs!Quantidade * tbs!ValorMaoObra) - (tbs!Quantidade * tbs!Desconto))
                      End If
                   End If
                End If
             End If
             tbs.MoveNext
             If tbs.EOF Then
                Return
             End If
             
       Wend
    End If
Return

Apura_Total_Est:
  tbi.Seek "=", tbr!NumRequisicao
  If Not tbi.NoMatch Then
     While tbi!NumRequisicao = tbr!NumRequisicao
           tbe.Seek "=", tbi!CodItem
           If Not tbe.NoMatch Then
              If tbe!Selecionar Then
                 tbc.Seek "=", tbe!CodigoCategoria
                 If Not tbc.NoMatch Then
                    If tbc!Selecionar Then
                        SomaTotal_Est = SomaTotal_Est + ((tbi!ValorUnit + tbi!Valor_IPI - tbi!Desconto) * (tbi!Quantidade - tbi!Devolucao))
                    End If
                 End If
              End If
           End If
           tbi.MoveNext
           If tbi.EOF Then
              Return
           End If
      Wend
   End If
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
   
End Function

Function Total_OSV_Cat_Est_Sel(Numos As Long) As Double
On Error GoTo TratamentodeErro:

    Dim bd As DataBase
    Dim tbr, tbi, tbc, tbe As Recordset
    
    Dim SomaTotal_Est As Double
    Dim NumOSV As Double
    Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tbr = bd.OpenRecordset("RQM", DB_OPEN_TABLE)
    Set tbi = bd.OpenRecordset("RMXEMPXEST", DB_OPEN_TABLE)
    Set tbc = bd.OpenRecordset("CATPROD", DB_OPEN_TABLE)
    Set tbe = bd.OpenRecordset("EST", DB_OPEN_TABLE)
    
    tbr.index = "SecondaryKey"
    tbi.index = "RQM"
    tbc.index = "PrimaryKey"
    tbe.index = "PrimaryKey"
    
    SomaTotal_Est = 0
    NumOSV = Numos
    
    tbr.Seek "=", Numos
    If Not tbr.NoMatch Then
       While tbr!Numos = Numos
             GoSub Apura_Total_Est:
             tbr.MoveNext
             If tbr.EOF Then
                GoSub Finaliza:
             End If
       Wend
    End If
    
Finaliza:
    Total_OSV_Cat_Est_Sel = SomaTotal_Est
    tbi.Close
    tbr.Close
    tbc.Close
    tbe.Close
    bd.Close
    Set tbi = Nothing
    Set tbr = Nothing
    Set tbc = Nothing
    Set tbe = Nothing
    Set bd = Nothing
    Exit Function
Return

Apura_Total_Est:
  tbi.Seek "=", tbr!NumRequisicao
  If Not tbi.NoMatch Then
     While tbi!NumRequisicao = tbr!NumRequisicao
           tbe.Seek "=", tbi!CodItem
           If Not tbe.NoMatch Then
              If tbe!Selecionar Then
                 tbc.Seek "=", tbe!CodigoCategoria
                 If Not tbc.NoMatch Then
                    If tbc!Selecionar Then
                        SomaTotal_Est = SomaTotal_Est + ((tbi!ValorUnit + tbi!Vlr_Seguro + tbi!Vlr_Fret + tbi!vlr_import + tbi!Valor_IPI - tbi!Desconto) * (tbi!Quantidade - tbi!Devolucao))
                    End If
                 End If
              End If
           End If
           tbi.MoveNext
           If tbi.EOF Then
              Return
           End If
      Wend
   End If
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function ResumoFinanceiroMercadorias(Numos As Long) As Double
    On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim Requisicao, saidas As Recordset
    
    Dim SomaTotal As Double
    Dim NumOSV As Double
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set Requisicao = bd1.OpenRecordset("RQM", DB_OPEN_TABLE)
    Set saidas = bd1.OpenRecordset("RMXEMPXEST", DB_OPEN_TABLE)
    
    Requisicao.index = "SecondaryKey"
    saidas.index = "RQM"
    SomaTotal = 0
    NumOSV = Numos
    
    'GoSub LerRequisicao
    
    
    Requisicao.Seek "=", Numos
    If Not Requisicao.NoMatch Then
       Do Until Requisicao.EOF
          If Requisicao!Numos = Numos Then
             saidas.Seek "=", Requisicao!NumRequisicao
             If Not saidas.NoMatch Then
                Do Until saidas.EOF
                   If saidas!NumRequisicao = Requisicao!NumRequisicao Then
                      SomaTotal = SomaTotal + ((saidas!ValorUnit + saidas!Valor_IPI) * (saidas!Quantidade - saidas!Devolucao))
                   Else
                      Exit Do
                   End If
                   saidas.MoveNext
                Loop
             End If
          Else
             Exit Do
          End If
          Requisicao.MoveNext
       Loop
       ResumoFinanceiroMercadorias = SomaTotal
       saidas.Close
       Requisicao.Close
       bd1.Close
       Set saidas = Nothing
       Set Requisicao = Nothing
       Set bd1 = Nothing
    End If
    Exit Function
    
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
    


End Function

Function ResumoFinanceiroMercadoriasCM(Numos As Long) As Double
    Dim bd1 As DataBase
    Dim Requisicao, saidas As Recordset, Est As Recordset
    
    Dim SomaTotal As Double
    Dim NumOSV As Double
    Dim WCustoMedio As Double
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set Requisicao = bd1.OpenRecordset("RQM", DB_OPEN_TABLE)
    Set saidas = bd1.OpenRecordset("RMXEMPXEST", DB_OPEN_TABLE)
    Set Est = bd1.OpenRecordset("EMPXEST", DB_OPEN_TABLE)
    Est.index = "SecondaryKey"
    Requisicao.index = "SecondaryKey"
    saidas.index = "RQM"
    SomaTotal = 0
    NumOSV = Numos
    WCustoMedio = 0
    
    Requisicao.Seek "=", Numos
    If Not Requisicao.NoMatch Then
       Do Until Requisicao.EOF
          If Requisicao!Numos = Numos Then
             saidas.Seek "=", Requisicao!NumRequisicao
             If Not saidas.NoMatch Then
                Do Until saidas.EOF
                   If saidas!NumRequisicao = Requisicao!NumRequisicao Then
                      Est.Seek "=", saidas!CodItem
                      If Est.NoMatch Then
                         WCustoMedio = 0
                      Else
                         WCustoMedio = Est!CustoMedio
                      End If
                      SomaTotal = SomaTotal + (WCustoMedio * (saidas!Quantidade - saidas!Devolucao))
                   Else
                      Exit Do
                   End If
                   saidas.MoveNext
                Loop
             End If
          Else
             Exit Do
          End If
          Requisicao.MoveNext
       Loop
       ResumoFinanceiroMercadoriasCM = SomaTotal
       saidas.Close
       Est.Close
       Requisicao.Close
       bd1.Close
       Set saidas = Nothing
       Set Est = Nothing
       Set Requisicao = Nothing
       Set bd1 = Nothing
    End If
    
    Exit Function
    
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function


Function VerificaFechamentoOSV(WNumos As Long) As Double
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase
Dim OSV As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set OSV = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
OSV.index = "Primarykey"
OSV.Seek "=", WNumos
If OSV.NoMatch Then
   MsgBox "Ordem de Servi�o e Venda Inexistente...", 64, "SGa"
   OSV.Close
   Exit Function
Else
   If IsNull(OSV!datafechamento) Then
      VerificaFechamentoOSV = 1
   Else
      VerificaFechamentoOSV = 0
   End If
End If
OSV.Close
bd1.Close
Set OSV = Nothing
Set bd1 = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CalculoDescontoMercadorias(Numos As Long)
 On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim Requisicao, saidas As Recordset
    
    Dim SomaTotal As Double
    Dim NumOSV As Double
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set Requisicao = bd1.OpenRecordset("RQM", DB_OPEN_TABLE)
    Set saidas = bd1.OpenRecordset("RMXEMPXEST", DB_OPEN_TABLE)
    
    Requisicao.index = "SecondaryKey"
    saidas.index = "RQM"
    SomaTotal = 0
    NumOSV = Numos
    
    Requisicao.Seek "=", Numos
    If Not Requisicao.NoMatch Then
       Do Until Requisicao.EOF
          If Requisicao!Numos = Numos Then
             saidas.Seek "=", Requisicao!NumRequisicao
             If Not saidas.NoMatch Then
                Do Until saidas.EOF
                   If saidas!NumRequisicao = Requisicao!NumRequisicao Then
                      SomaTotal = SomaTotal + (saidas!Desconto * (saidas!Quantidade - saidas!Devolucao))
                   Else
                      Exit Do
                   End If
                   saidas.MoveNext
                Loop
             End If
          Else
             Exit Do
          End If
          Requisicao.MoveNext
       Loop
       CalculoDescontoMercadorias = SomaTotal
       saidas.Close
       Requisicao.Close
       bd1.Close
       Set saidas = Nothing
       Set Requisicao = Nothing
       Set bd1 = Nothing
    Else
       CalculoDescontoMercadorias = 0
    End If
    Exit Function
    
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function



Function TotalMercadoriasOrdemServico(Numos As Long) As Double
    Dim bd1 As DataBase
    Dim Requisicao, saidas As Recordset
    
    Dim NumRequisicao, NUMOSATUAL, highvalue  As Long
    Dim SomaTotal, ValorDesconto As Double
    Dim Quantidade As Long
    
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set Requisicao = bd1.OpenRecordset("RQM", DB_OPEN_TABLE)
    Set saidas = bd1.OpenRecordset("RMXEMPXEST", DB_OPEN_TABLE)
    
    Requisicao.index = "SecondaryKey"
    saidas.index = "PrimaryKey"
    
    highvalue = 999999999
    SomaTotal = 0
    NUMOSATUAL = highvalue
    
    Requisicao.Seek "=", Numos
    If Requisicao.NoMatch Then
       NUMOSATUAL = highvalue
     Else
       NUMOSATUAL = Numos
    End If
    
    While NUMOSATUAL = Numos
          Requisicao.Edit
          NumRequisicao = Requisicao!NumRequisicao
          GoSub LerSaidas
          Requisicao.MoveNext
          If Requisicao.EOF Then
             NUMOSATUAL = highvalue
           Else
             NUMOSATUAL = Requisicao!Numos
          End If
    Wend

Finaliza:
    TotalMercadoriasOrdemServico = SomaTotal
    Requisicao.Close
    saidas.Close
    bd1.Close
    Set Requisicao = Nothing
    Set saidas = Nothing
    Set bd1 = Nothing
    Exit Function
    
LerSaidas:
    saidas.Seek ">=", NumRequisicao, 1, 1
    If saidas.NoMatch Then
       Return
       End If
    While saidas!NumRequisicao = NumRequisicao
          saidas.Edit
          If Not IsNull(saidas!DataAtendimento) And ((saidas!Quantidade - saidas!Devolucao) > 0) Then
             GoSub SomaSaidas
             End If
          saidas.MoveNext
          If saidas.EOF Then
             Return
             End If
    Wend
    Return

SomaSaidas:
    Quantidade = saidas!Quantidade - saidas!Devolucao
    ValorDesconto = saidas!Desconto
    SomaTotal = SomaTotal + ((Quantidade * Format(saidas!ValorUnit + saidas!Valor_IPI + saidas!Vlr_Seguro + saidas!Vlr_Icms_ST + saidas!vlr_import + saidas!Vlr_Out_Desp + saidas!Vlr_Fret, "##,###.##")) - (Quantidade * ValorDesconto))
    Return


End Function

Function Total_Acrescimos_OSV(Numos As Long) As Double

    Dim bd1 As DataBase
    Dim Requisicao, saidas As Recordset
    
    Dim NumRequisicao, NUMOSATUAL, highvalue  As Long
    Dim SomaTotal, ValorDesconto As Double
    Dim Quantidade As Double
    
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set Requisicao = bd1.OpenRecordset("RQM", DB_OPEN_TABLE)
    Set saidas = bd1.OpenRecordset("RMXEMPXEST", DB_OPEN_TABLE)
    
    Requisicao.index = "SecondaryKey"
    saidas.index = "PrimaryKey"
    
    highvalue = 999999999
    SomaTotal = 0
    NUMOSATUAL = highvalue
    
    Requisicao.Seek "=", Numos
    If Requisicao.NoMatch Then
       NUMOSATUAL = highvalue
     Else
       NUMOSATUAL = Numos
    End If
    
    While NUMOSATUAL = Numos
          Requisicao.Edit
          NumRequisicao = Requisicao!NumRequisicao
          GoSub LerSaidas
          Requisicao.MoveNext
          If Requisicao.EOF Then
             NUMOSATUAL = highvalue
           Else
             NUMOSATUAL = Requisicao!Numos
          End If
    Wend

Finaliza:
    Total_Acrescimos_OSV = SomaTotal
    Requisicao.Close
    saidas.Close
    bd1.Close
    Set Requisicao = Nothing
    Set saidas = Nothing
    Set bd1 = Nothing
    Exit Function
    
LerSaidas:
    saidas.Seek ">=", NumRequisicao, 1, 1
    If saidas.NoMatch Then
       Return
       End If
    While saidas!NumRequisicao = NumRequisicao
          saidas.Edit
          If Not IsNull(saidas!DataAtendimento) And ((saidas!Quantidade - saidas!Devolucao) > 0) Then
             GoSub SomaSaidas
             End If
          saidas.MoveNext
          If saidas.EOF Then
             Return
             End If
    Wend
    Return

SomaSaidas:
    Quantidade = saidas!Quantidade - saidas!Devolucao
    ValorDesconto = saidas!Desconto
    SomaTotal = SomaTotal + ((Quantidade * (saidas!Valor_IPI + saidas!Vlr_Seguro + saidas!Vlr_Icms_ST + saidas!vlr_import + saidas!Vlr_Out_Desp + saidas!Vlr_Fret)))
    Return

End Function

Function ConsisteCampo(Campo As String) As Boolean
    If IsNull(Campo) Then
       ConsisteCampo = False
      Else
       ConsisteCampo = True
    End If
End Function
Function MediaAtrazoCliente(CodigoCliente As Long)
    Dim bd1 As DataBase
    Dim ContasReceber As Recordset
    Dim SomaAtrazo, cont As Long
    Dim Data As Variant
    DoCmd.Hourglass True
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set ContasReceber = bd1.OpenRecordset("CTREC", DB_OPEN_TABLE)
    Data = Date - 180
    cont = 0
    ContasReceber.index = "TercKey"
    ContasReceber.Seek ">=", CodigoCliente, Data
    If ContasReceber.NoMatch Then
       CodigoCliente = highvalue
       cont = 1
      Else
       If ContasReceber!CodigoCliente <> CodigoCliente Then
          CodigoCliente = highvalue
          cont = 1
       End If
    End If
    SomaAtrazo = 0
    While CodigoCliente <> highvalue
          ContasReceber.Edit
          GoSub CalculaAtrazo
          ContasReceber.MoveNext
          If ContasReceber.EOF Then
             CodigoCliente = highvalue
            Else
             If ContasReceber!CodigoCliente > CodigoCliente Then
                CodigoCliente = highvalue
               Else
                CodigoCliente = ContasReceber!CodigoCliente
             End If
          End If
    Wend
    If cont = 0 Then
       cont = 1
    End If
    MediaAtrazoCliente = Int(SomaAtrazo / cont)
    
    DoCmd.Hourglass False
    ContasReceber.Close
    bd1.Close
    Set ContasReceber = Nothing
    Set bd1 = Nothing
    Exit Function

CalculaAtrazo:
    If IsNull(ContasReceber!DataPagto) And ContasReceber!DataVencimento > Date Then
       ContasReceber!DataPagto = ContasReceber!DataVencimento
       cont = cont - 1
       End If
    If IsNull(ContasReceber!DataPagto) And ContasReceber!DataVencimento <= Date Then
       ContasReceber!DataPagto = Date
       End If
    
    SomaAtrazo = SomaAtrazo + (ContasReceber!DataPagto - ContasReceber!DataVencimento)
    cont = cont + 1
    Return
    
End Function
Function CalculoEstoqueMinimo()
Dim bd1 As DataBase, bd As DataBase, QCalc As QueryDef, Calc As Recordset, WQtdeForn, WCarForn, WCodForn As Long
Dim Fornec As Recordset, Est As Recordset, WPar As Integer, WCodItem As Long, WMediaDiaria As Long
Dim WEstAt As Double, WCusto As Double, WTend As Double, WSaidasNoMes As Double
Dim RM, FORXEST, EMPXEST As Recordset
Dim WDU, WDD As Double
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set bd = CurrentDb()
Set Fornec = bd1.OpenRecordset("FOR", DB_OPEN_TABLE)
Set FORXEST = bd1.OpenRecordset("FORXEST", DB_OPEN_TABLE)
Set RM = bd1.OpenRecordset("RMXEMPXEST", DB_OPEN_TABLE)
Set Est = bd1.OpenRecordset("Estm", DB_OPEN_TABLE)
Set QCalc = bd.QueryDefs("CalculaEstoqueMinimo")
Set Calc = QCalc.OpenRecordset()
Est.index = "Primarykey"
FORXEST.index = "primarykey"
RM.index = "CODITEM"
WDU = DiasUteis(Date)
WDD = DiasDecorridos(Date)
WPar = 30
If Est.RecordCount <> 0 Then
   BeginTrans
   Do Until Est.EOF
      Est.Edit
      Est!Minimo = 0
      Est.Update
      Est.MoveNext
   Loop
   CommitTrans
End If

If Calc.RecordCount <> 0 Then
 
   WCodItem = 0
   Fornec.index = "Primarykey"
   BeginTrans
   Do Until Calc.EOF
      If WCodItem = 0 Then
         WCodItem = Calc!CodItem
         WCodForn = Calc!CodigoFornecedor
      End If
         
      If Calc!CodItem = WCodItem Then
         GoSub LerFornecedor:
       Else
         WCarForn = Int(WCarForn / WQtdeForn)
         GoSub AtualizaEstoqueMinimo:
         WCarForn = 0
         WQtdeForn = 0
         GoSub LerFornecedor:
         WCodItem = Calc!CodItem
         WCodForn = Calc!CodigoFornecedor
      End If
      
      WMediaDiaria = Int(Calc!saidas / WPar)
      WSaidasNoMes = Calc!saidas
      WEstAt = Saldo_Estoque(WCodItem, Date, Empresa_Ativa())
      WCusto = Calc!CustoMedio
      WTend = Int(WMediaDiaria * Calc!DU)
      Calc.MoveNext
   Loop
   GoSub AtualizaEstoqueMinimo:
   CommitTrans
End If
Calc.Close
Fornec.Close
Est.Close
FORXEST.Close
RM.Close
bd1.Close
bd.Close
Set Calc = Nothing
Set Fornec = Nothing
Set Est = Nothing
Set FORXEST = Nothing
Set RM = Nothing
Set bd1 = Nothing
Set bd = Nothing
Exit Function

AtualizaEstoqueMinimo:
 Est.Seek "=", WCodItem
 
 If WMediaDiaria < 1 Then
    WMediaDiaria = 1
 End If
 
 If Not Est.NoMatch Then
    Est.Edit
 Else
    Est.AddNew
 End If
 Est!CodigoItem = WCodItem
 Est!Minimo = (WMediaDiaria * WCarForn) + (((WMediaDiaria * WCarForn) * 10) / 100)
 Est!EstoqueAtual = WEstAt
 Est!SaidasNoMes = WSaidasNoMes
 Est!MediaDia = WMediaDiaria
 Est!TendenciaMes = WTend
 Est!CustoMedio = WCusto
 Est.Update
 
Return

LerFornecedor:
 Fornec.Seek "=", WCodForn
 WCarForn = WCarForn + Fornec!PrazoEntrega
 WQtdeForn = WQtdeForn + 1
Return


End Function

Function PrazoEntregaFornec()
On Error GoTo TratamentodeErro:
Dim bd1 As DataBase, bd As DataBase, QPrz As QueryDef, PRZ As Recordset, Fornec As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set bd = CurrentDb()
Set Fornec = bd1.OpenRecordset("FOR", DB_OPEN_TABLE)
Set QPrz = bd.QueryDefs("MediaPrazoEntregaFornecedores")
Set PRZ = QPrz.OpenRecordset()
If PRZ.RecordCount <> 0 Then
   Fornec.index = "PrimaryKey"
   BeginTrans
   Do Until PRZ.EOF
      Fornec.Seek "=", PRZ!CodigoFornecedor
      If Not Fornec.NoMatch Then
         Fornec.Edit
         Fornec!PrazoEntrega = PRZ!nrodias
         Fornec.Update
      End If
      PRZ.MoveNext
   Loop
   CommitTrans
End If

PRZ.Close
Fornec.Close
bd1.Close
Set PRZ = Nothing
Set Fornec = Nothing
Set bd1 = Nothing
Exit Function

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function CriaGraficoOSABXOSEXXOSABD(WDataInicial As Variant) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase, QGrafA, QGrafd, QGraff As QueryDef, GrafA, Grafd, Graff As Recordset, TGraf As Recordset
Set bd = CurrentDb()
Set TGraf = bd.OpenRecordset("GraficoOSABXOSEXXOSABD", DB_OPEN_TABLE)
Set QGrafA = bd.QueryDefs("OSABXOSEXXOSABA")
'QGrafA.Parameters (0) <= WDataInicial
'QGrafA.Parameters(1) = WDataFinal
Set GrafA = QGrafA.OpenRecordset()
Set QGrafd = bd.QueryDefs("OSABXOSEXXOSABD")
QGrafd.Parameters(0) = WDataInicial
'QGrafd.Parameters(1) = WDataFinal
Set Grafd = QGrafd.OpenRecordset()
Set QGraff = bd.QueryDefs("OSABXOSEXXOSABF")
QGraff.Parameters(0) = WDataInicial
'QGraff.Parameters(1) = WDataFinal
Set Graff = QGraff.OpenRecordset()

TGraf.index = "PrimaryKey"

BeginTrans
  TGraf.Seek "=", 1
  If TGraf.NoMatch Then
     TGraf.AddNew
  Else
     TGraf.Edit
  End If
  TGraf!Ident = 1
  TGraf!OSVABERTAS = 0
  TGraf!OSVfechadasnodia = 0
  TGraf!OSVABERTASnodia = 0
  TGraf.Update
CommitTrans

If GrafA.RecordCount <> 0 Then
   BeginTrans
   Do Until GrafA.EOF
      TGraf.Seek "=", 1
      If TGraf.NoMatch Then
         TGraf.AddNew
       Else
         TGraf.Edit
      End If
      TGraf!Ident = 1
      TGraf!OSVABERTAS = TGraf!OSVABERTAS + 1
      TGraf.Update
      GrafA.MoveNext
   Loop
   CommitTrans
End If

If Grafd.RecordCount <> 0 Then
   BeginTrans
   Do Until Grafd.EOF
      TGraf.Seek "=", 1
      If TGraf.NoMatch Then
         TGraf.AddNew
       Else
         TGraf.Edit
      End If
      TGraf!Ident = 1
      TGraf!OSVfechadasnodia = TGraf!OSVfechadasnodia + 1
      TGraf.Update
      Grafd.MoveNext
   Loop
   CommitTrans
End If


If Graff.RecordCount <> 0 Then
   BeginTrans
   Do Until Graff.EOF
      TGraf.Seek "=", 1
      If TGraf.NoMatch Then
         TGraf.AddNew
       Else
         TGraf.Edit
      End If
      TGraf!Ident = 1
      TGraf!OSVABERTASnodia = TGraf!OSVABERTASnodia + 1
      TGraf.Update
      Graff.MoveNext
   Loop
   CommitTrans
End If

Finaliza:
  GrafA.Close
  Graff.Close
  Grafd.Close
  TGraf.Close
  bd.Close
  Set GrafA = Nothing
  Set Graff = Nothing
  Set Grafd = Nothing
  Set TGraf = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Ativa_Conectividade_Autoexec() As Long
'[CONECTIVIDADE_NOVA] Abertura do app antigo (C:\SGA\conectividade.exe) desativada.
'A nova extensao Chrome substitui o app antigo e usa as mesmas pastas
'C:\Sga\Req e C:\Sga\Resp, entao nao e mais necessario abrir o conectividade.exe
'no login. A funcao foi esvaziada (as chamadas a ela continuam existindo e
'nao quebram, mas nao abrem mais nada).
'
'If ConectividadeAtiva(Empresa_Ativa()) = True Then
'   Dim WAPPName As String
'   Dim ofuncoes As CanadaConnectivity.FuncoesDiversas
'   Set ofuncoes = New CanadaConnectivity.FuncoesDiversas
'   WAPPName = "C:\SGA\conectividade.exe"
'   If ofuncoes.CheckFile(WAPPName) = True Then
'      Call Shell(WAPPName, 1)
'   Else
'      MsgBox "Ocorreu um erro ao ativar a Conectividade, entre em contato com o Suporte SGA.", vbCritical
'   End If
' End If
End Function

Function ExisteModeloVeic(WModeloVeic As String, WMarcaVeic As String) As Boolean
On Error GoTo TratamentodeErro:
Dim bd As DataBase, tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("MODELOVEIC", DB_OPEN_TABLE)

tb.index = "ChavePrimaria"
tb.Seek "=", WMarcaVeic, WModeloVeic

If tb.NoMatch Then
   ExisteModeloVeic = False
Else
   ExisteModeloVeic = True
End If


Finaliza:
  
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Public Function GeraArquivoLoginInmetro() As Boolean
If ConectividadeAtiva(Empresa_Ativa()) = True Then
    Dim oReq As CanadaConnectivity.ReqLogin
    Set oReq = New CanadaConnectivity.ReqLogin
    Set oReq.ologin = New CanadaConnectivity.Login
    oReq.ologin.User = RetornaLoginInMetro
    oReq.ologin.password = RetornaSenhaInMetro
    oReq.CreateLogin ("www.canainfo.com.br")
    Set oReq = Nothing
End If
GeraArquivoLoginInmetro = True
End Function
Function Inicializa_STATUS_PLACXCLI_SMS() As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb  As Recordset
Dim X As Long
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Base_Dados_Lote", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
BeginTrans
If tb.RecordCount <> 0 Then
   While Not tb.EOF
         tb.Edit
         tb!Status = Null
         tb!Enviar = True
         tb.Update
         X = X + 1
         If X Mod 200 = 0 Then
            CommitTrans
            BeginTrans
         End If
         tb.MoveNext
   Wend
End If


Finaliza:
  CommitTrans
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function Inicializa_Status_SMS() As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb  As Recordset
Dim X As Long
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("CLI", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
BeginTrans
If tb.RecordCount <> 0 Then
   While Not tb.EOF
         If tb!Ativo Then
            tb.Edit
            tb!Status = Null
            tb!Enviar = True
            tb.Update
         End If
         X = X + 1
         If X Mod 200 = 0 Then
            CommitTrans
            BeginTrans
         End If
         tb.MoveNext
   Wend
End If


Finaliza:
  CommitTrans
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function Atualiza_Status_SMS_Cli(WCodigoCliente As Integer, wMensagem As String) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb  As Recordset
Dim X As Long
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Base_Dados_Lote", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", WCodigoCliente
If tb.NoMatch Then GoSub Finaliza:
BeginTrans
If tb.RecordCount <> 0 Then
   tb.Edit
   tb!Status_SMS = wMensagem
   tb.Update
End If
CommitTrans
Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function
Function Atualiza_Status_SMS(WCodigoCliente As Integer, wMensagem As String) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb  As Recordset
Dim X As Long
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Base_Dados_Lote", DB_OPEN_TABLE)
tb.index = "CodigoCliente"
tb.Seek "=", WCodigoCliente
BeginTrans
If tb.NoMatch Then GoSub Finaliza:
If tb.RecordCount <> 0 Then
   While tb!CodigoCliente = WCodigoCliente
      tb.Edit
      tb!Status = wMensagem
      tb.Update
      tb.MoveNext
      If tb.EOF Then
         GoSub Finaliza
      End If
   Wend
End If
Finaliza:
  CommitTrans
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function
Function Atualiza_Status_EMAIL(WCodigoCliente As Long, wMensagem As String) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb  As Recordset
Dim X As Long
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Base_Dados_Lote", DB_OPEN_TABLE)
tb.index = "CodigoCliente"
tb.Seek "=", WCodigoCliente
BeginTrans
If tb.NoMatch Then GoSub Finaliza:
If tb.RecordCount <> 0 Then
   While tb!CodigoCliente = WCodigoCliente
      tb.Edit
      tb!Status = wMensagem
      tb.Update
      tb.MoveNext
      If tb.EOF Then
         GoSub Finaliza
      End If
   Wend
End If
Finaliza:
  CommitTrans
  tb.Close
  bd.Close
  
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Atualiza_Status_Enviar(WCodigoCliente As Long, WEnviar As Boolean) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb  As Recordset
Dim X As Long
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Base_Dados_Lote", DB_OPEN_TABLE)
BeginTrans
    tb.index = "CodigoCliente"
    tb.Seek "=", WCodigoCliente
    If Not tb.NoMatch Then
       While tb!CodigoCliente = WCodigoCliente
            tb.Edit
            tb!Enviar = WEnviar
            tb.Update
            tb.MoveNext
            If tb.EOF Then
                GoSub Finaliza
            End If
       Wend
    End If
Finaliza:
  CommitTrans
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Desmarca_Todos_PLACXCLI_SMS() As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb  As Recordset
Dim X As Long
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("PLACXCLI", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
BeginTrans
If tb.RecordCount <> 0 Then
   While Not tb.EOF
         'If Tb!Ativo Then
            tb.Edit
            tb!Enviar_SMS = 0
            tb.Update
         'End If
         X = X + 1
         If X Mod 200 = 0 Then
            CommitTrans
            BeginTrans
         End If
         tb.MoveNext
   Wend
End If


Finaliza:
  CommitTrans
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Desmarca_Todos_Cli_SMS() As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb  As Recordset
Dim X As Long
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("CLI", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
BeginTrans
If tb.RecordCount <> 0 Then
   While Not tb.EOF
         If tb!Ativo Then
            tb.Edit
            tb!Enviar_SMS = 0
            tb.Update
         End If
         X = X + 1
         If X Mod 200 = 0 Then
            CommitTrans
            BeginTrans
         End If
         tb.MoveNext
   Wend
End If


Finaliza:
  CommitTrans
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Seleciona_Todas_PLACXCLI_SMS() As Long

On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb  As Recordset
Dim X As Long
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("PLACXCLI", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
BeginTrans
If tb.RecordCount <> 0 Then
   While Not tb.EOF
            tb.Edit
            tb!Enviar_SMS = 1
            tb.Update
         X = X + 1
         If X Mod 200 = 0 Then
            CommitTrans
            BeginTrans
         End If
         tb.MoveNext
   Wend
End If


Finaliza:
  CommitTrans
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       
       Resume Next
    End If


End Function
Function Seleciona_Todas_PLACXBaseLote_SMS(ByVal WStatus As Boolean) As Long

On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb  As Recordset
Dim X As Long
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("Base_Dados_Lote", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
BeginTrans
If tb.RecordCount <> 0 Then
   While Not tb.EOF
         tb.Edit
         tb!Enviar = WStatus
         tb.Update
         X = X + 1
         If X Mod 200 = 0 Then
            CommitTrans
            BeginTrans
         End If
         tb.MoveNext
   Wend
End If


Finaliza:
  CommitTrans
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       
       Resume Next
    End If


End Function

Function Seleciona_Todos_Cli_SMS() As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb  As Recordset
Dim X As Long
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("CLI", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
BeginTrans
If tb.RecordCount <> 0 Then
   While Not tb.EOF
         If tb!Ativo Then
            tb.Edit
            tb!Enviar_SMS = 1
            tb.Update
         End If
         X = X + 1
         If X Mod 200 = 0 Then
            CommitTrans
            BeginTrans
         End If
         tb.MoveNext
   Wend
End If


Finaliza:
  CommitTrans
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function Relacao_Placas_Consultadas_LotexCodCliente(ByVal wCodCliente As Long) As String
On Error GoTo TratamentodeErro:
Dim bdc As DataBase
Dim tb As Recordset
Dim pStringPlaca As String
Set bdc = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bdc.OpenRecordset("Base_Dados_Lote", DB_OPEN_TABLE)
tb.index = "CodigoCliente"
tb.Seek "=", wCodCliente
If tb.NoMatch Then GoTo Finaliza
While tb!CodigoCliente = wCodCliente
    pStringPlaca = pStringPlaca & Space(1) & tb!Placa
    tb.MoveNext
    If tb.EOF Then
        GoTo Finaliza
    End If
Wend
Finaliza:
  Relacao_Placas_Consultadas_LotexCodCliente = pStringPlaca
  tb.Close
  bdc.Close
  Set tb = Nothing
  Set bdc = Nothing
  Exit Function
Return
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function


Function Envio_SMS_Clientes_Analit(wMensagem As String, wVenctoF As Variant) As Long
On Error GoTo TratamentodeErro:
Dim bdc As DataBase
Dim tbc As Recordset
Dim qtbc As QueryDef
Dim WTelefone_Cel As String
Dim WMensagem_SMS As String
Dim WStatus_venc As String
Dim WCert_In As String
Dim WObj As CanadaUtils.SMS

Dim X As Long
Set bdc = CurrentDb()
Set qtbc = bdc.QueryDefs("Consulta_Plac_Cli_Conec_Envio_SMS_A")
Set tbc = qtbc.OpenRecordset()
BeginTrans
If tbc.RecordCount = 0 Then
   Envio_SMS_Clientes_Analit = 0
Else
   X = 0
   While Not tbc.EOF
         If tbc!Enviar Then
           ' tbc.Edit
            Set WObj = New CanadaUtils.SMS
         
            WTelefone_Cel = (EliminaEspacos(EliminaCaracEspec(tbc!Celular)))
            WStatus_venc = IIf(tbc!Validade = "01/01/2001" Or DateValue(Format$(tbc!Validade, "dd/mm/yyyy")) < wVenctoF, " Expirado ", tbc!Validade)
            WCert_In = IIf(tbc!NumeroCertificadoDefinitivo = "" Or IsEmpty(tbc!NumeroCertificadoDefinitivo) Or IsNull(tbc!NumeroCertificadoDefinitivo), " inexistente ", tbc!NumeroCertificadoDefinitivo)
            WMensagem_SMS = StrConv(Identif_SMS(), vbProperCase) & ":Veiculo Placa " & UCase(Trim(tbc!Placa)) & " Cert.Ensaio Cronotacografo " & Trim$(WCert_In) & " Com Vencto " & Trim(WStatus_venc) & ", Entre Em Contato Conosco e Saiba Mais " & Telefone()
            
            WObj.NumPorta = Porta_Comunic_SMS()
            WObj.Mensagem = WMensagem_SMS
            WObj.Telefone = WTelefone_Cel
            
            If WObj.ValidaDadosMensagem = True Then
               WObj.EnviarMensagem
               Call Atualiza_Status_SMS(tbc!CodigoCliente, WObj.StatusMensagem)
            Else
               Call Atualiza_Status_SMS(tbc!CodigoCliente, WObj.StatusMensagem)
            End If

            Call LogSGa("CLI  ", "ENVIO_SMS_CLIE", " ENVIO SMS CLIENTES  ", Captura_Conteudo_Campo_RESP(Usuario_Ativo(), "nome"), Date, Time(), "CEL " & WTelefone_Cel, Environ("ComputerName"))
           ' tbc.Update
            Set WObj = Nothing
            
         End If
         X = X + 1
         If X Mod Retorna_Quant_SMS_Remessa() = 0 Then
            CommitTrans
            'sleep (Retorna_Interv_Remessa())
            BeginTrans
         End If
         tbc.MoveNext
   Wend
End If

Finaliza:
  CommitTrans
  MsgBox "Fim normal de execu��o", vbInformation, "Sga"
  tbc.Close
  bdc.Close
  Set tbc = Nothing
  Set bdc = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Envio_Email_Clientes_Analit(WRemetente As String, WAssunto As String, wMensagem As String, wRodape As String) As Long
On Error GoTo TratamentodeErro:
Dim bdc As DataBase
Dim Tbcs As Recordset
Dim Qtbcs As QueryDef
Dim WMensagem_Email As String
Dim WStatus_venc As String
Dim WCert_In As String
Dim WQuant_Plac_S As Long
Dim WQuant_Reg As Double
Dim X As Long
Dim MensagemRetorno As String
Dim WRazaoSocial As String
Dim WCNPJ As String
Dim WEndElet As String
Dim WCodigoCliente As Long

Set bdc = CurrentDb()
Set Qtbcs = bdc.QueryDefs("Consulta_Plac_Cli_Conec_Envio_SMS_A")
Set Tbcs = Qtbcs.OpenRecordset()

BeginTrans
If Tbcs.RecordCount = 0 Then
   Envio_Email_Clientes_Analit = 0
Else
   X = 0
   Tbcs.MoveLast
   Tbcs.MoveFirst
   WQuant_Reg = Tbcs.RecordCount
   While Not Tbcs.EOF
         If Tbcs!Enviar Then
            If WCodigoCliente = 0 Then
               WCodigoCliente = Tbcs!codigoCli
               WRazaoSocial = Tbcs!RazaoSocial
               WCNPJ = "CNPJ" & Tbcs!CgcCpf
               WEndElet = Tbcs!EndElet
               Open "C:\SGa\Placas_" & WCNPJ & ".csv" For Output As #1
               GoSub Grava_Cabecalho:
             ElseIf WCodigoCliente <> Tbcs!codigoCli Then
                    Close #1
                    GoSub Envia_Planilha:
                    WCodigoCliente = Tbcs!codigoCli
                    WRazaoSocial = Tbcs!RazaoSocial
                    WEndElet = Tbcs!EndElet
                    WCNPJ = "CNPJ" & Tbcs!CgcCpf
                    Open "C:\SGa\Placas_" & WCNPJ & ".csv" For Output As #1
                    GoSub Grava_Cabecalho:
              End If
              Print #1, Tbcs!Placa & ";" & Tbcs!RENAVAN & ";" & Tbcs!NroSerieEquip & ";" & Tbcs!DataValidadeResultado
         End If
         X = X + 1
         If X = 0 Then
            CommitTrans
            BeginTrans
         End If
         Tbcs.MoveNext
   Wend
End If

Finaliza:
  CommitTrans
  Close #1
  GoSub Envia_Planilha:
  MsgBox "Fim normal de execu��o", vbInformation, "Sga"
  Tbcs.Close
  bdc.Close
  Set Tbcs = Nothing
  Set bdc = Nothing
  Exit Function
Return

Envia_Planilha:
 
 Call EnviarMail(Email(), WEndElet, WAssunto, "C:\SGa\Placas_" & WCNPJ & ".csv", wMensagem)
 Call LogSGa("CLI  ", "ENVIO_MAIL_EMAIL", " CONSULTA LOTE PLACAS  " & WRazaoSocial, Captura_Conteudo_Campo_RESP(Usuario_Ativo(), "nome"), Date, Time(), "MAQUINA ", Environ("ComputerName"))
 'Form_Envio_Email_Consulta_Lotes_Placas.Dialogo_Envio_EMAIL_Conect_Sub_S.Requery

Return


Grava_Cabecalho:

   Print #1, "PLACA VEICULO;RENAVAM;NROSERIE;VENCIMENTO"

Return

Grava_Rodape:

   Print #1, WRazaoSocial
   Print #1, "Data" & Date & ";Hora" & Time()

Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Envio_Email_Clientes_Sint(WRemetente As String, WAssunto As String, wMensagem As String, wRodape As String) As Long
On Error GoTo TratamentodeErro:
Dim bdc As DataBase
Dim Tbcs As Recordset
Dim Qtbcs As QueryDef
Dim WMensagem_Email As String
Dim WStatus_venc As String
Dim WCert_In As String
Dim WQuant_Plac_S As Long
Dim WQuant_Reg As Double
Dim X As Long
Dim MensagemRetorno As String

Set bdc = CurrentDb()
Set Qtbcs = bdc.QueryDefs("Consulta_Plac_Cli_Conec_Envio_SMS_S")
Set Tbcs = Qtbcs.OpenRecordset()
BeginTrans
If Tbcs.RecordCount = 0 Then
   Envio_Email_Clientes_Sint = 0
Else
   X = 0
   Tbcs.MoveLast
   Tbcs.MoveFirst
   WQuant_Reg = Tbcs.RecordCount
   While Not Tbcs.EOF
         If Tbcs!Enviar Then
            MensagemRetorno = ""
            MensagemRetorno = Enviar_Email_Lote_Placas(UCase(WAssunto), wMensagem & ". " & vbCrLf & vbCrLf & " Relacao de Placas.: " & Tbcs!RelacaoPlacas & vbCrLf & vbCrLf & vbCrLf & wRodape, WRemetente, Tbcs!EndElet, "#")
            Call Atualiza_Status_SMS_Cli(Tbcs!CodigoCliente, MensagemRetorno)
            Call LogSGa("CLI  ", "ENVIO_MAIL_EMAIL", " ENVIO MAIL CLIENTES  ", Captura_Conteudo_Campo_RESP(Usuario_Ativo(), "nome"), Date, Time(), "MAQUINA ", Environ("ComputerName"))
         End If
         X = X + 1
         If X Mod Retorna_Quant_SMS_Remessa() = 0 Then
            CommitTrans
            'sleep (Retorna_Interv_Remessa())
            BeginTrans
         End If
         Tbcs.MoveNext
   Wend
End If

Finaliza:
  CommitTrans
  MsgBox "Fim normal de execu��o", vbInformation, "Sga"
  Tbcs.Close
  bdc.Close
  Set Tbcs = Nothing
  Set bdc = Nothing
  Exit Function
Return
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function
Function Envio_SMS_Clientes_Sint(wMensagem As String, wVenctoF As Variant) As Long
On Error GoTo TratamentodeErro:
Dim bdc As DataBase
Dim Tbcs As Recordset
Dim Qtbcs As QueryDef
Dim WTelefone_Cel As String
Dim WMensagem_SMS As String
Dim WStatus_venc As String
Dim WCert_In As String
Dim WQuant_Plac_S As Long
Dim WQuant_Reg As Double
Dim WObj As CanadaUtils.SMS
Dim X As Long

Set bdc = CurrentDb()
Set Qtbcs = bdc.QueryDefs("Consulta_Plac_Cli_Conec_Envio_SMS_S")

Set Tbcs = Qtbcs.OpenRecordset()
BeginTrans
If Tbcs.RecordCount = 0 Then
   Envio_SMS_Clientes_Sint = 0
Else
   X = 0
   Tbcs.MoveLast
   Tbcs.MoveFirst
   WQuant_Reg = Tbcs.RecordCount
   While Not Tbcs.EOF
         If Tbcs!Enviar Then
            Set WObj = New CanadaUtils.SMS
         '   Tbcs.Edit
            WQuant_Plac_S = Retorna_Quant_Plac_Cli_Venc_Conect(Tbcs!CodigoCliente, wVenctoF)
            WTelefone_Cel = (EliminaEspacos(EliminaCaracEspec(Tbcs!Celular)))
            WMensagem_SMS = StrConv(Identif_SMS(), vbProperCase) & ": Existem " & WQuant_Plac_S & " Veiculos Com Certificado Ensaio Cronotacografos Com Vencimento Ate " & Format$(wVenctoF, "dd/mm/yyyy") & " Entre Em Contato Conosco e Saiba Mais " & Telefone()
            
            WObj.NumPorta = Porta_Comunic_SMS()
            WObj.Mensagem = WMensagem_SMS
            WObj.Telefone = WTelefone_Cel
            
            If WObj.ValidaDadosMensagem = True Then
               WObj.EnviarMensagem
               Call Atualiza_Status_SMS(Tbcs!CodigoCliente, WObj.StatusMensagem)
            Else
               Call Atualiza_Status_SMS(Tbcs!CodigoCliente, WObj.StatusMensagem)
            End If
            
            Call LogSGa("CLI  ", "ENVIO_SMS_CLIE", " ENVIO SMS CLIENTES  ", Captura_Conteudo_Campo_RESP(Usuario_Ativo(), "nome"), Date, Time(), "CEL " & WTelefone_Cel, Environ("ComputerName"))
          '  Tbcs.Update
            
            Set WObj = Nothing
            
         End If
         X = X + 1
         If X Mod Retorna_Quant_SMS_Remessa() = 0 Then
            CommitTrans
            'sleep (Retorna_Interv_Remessa())
            BeginTrans
         End If
         Tbcs.MoveNext
   Wend
End If

Finaliza:
  CommitTrans
  MsgBox "Fim normal de execu��o", vbInformation, "Sga"
  Tbcs.Close
  bdc.Close
  Set Tbcs = Nothing
  Set bdc = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Envio_SMS_Clientes(wMensagem As String) As Long
On Error GoTo TratamentodeErro:
Dim bdc As DataBase
Dim tbc As Recordset
Dim qtbc As QueryDef
Dim WTelefone_Cel As String
Dim WStatus_Env As String
Dim X As Long
Dim WObj As CanadaUtils.SMS
Set bdc = CurrentDb()
Set qtbc = bdc.QueryDefs("Enviar_SMS_Clientes")
Set tbc = qtbc.OpenRecordset()
BeginTrans
If tbc.RecordCount = 0 Then
   Envio_SMS_Clientes = 0
Else
   X = 0
   While Not tbc.EOF
         If tbc!Enviar_SMS Then
            Set WObj = New CanadaUtils.SMS
            tbc.Edit
            WTelefone_Cel = (EliminaEspacos(EliminaCaracEspec(tbc!Celular)))
                       
            WObj.NumPorta = Porta_Comunic_SMS()
            WObj.Mensagem = wMensagem
            WObj.Telefone = WTelefone_Cel
            
            If WObj.ValidaDadosMensagem = True Then
               WObj.EnviarMensagem
               tbc!Status_SMS = WObj.StatusMensagem
            Else
               tbc!Status_SMS = WObj.StatusMensagem
            End If
                     
            tbc.Update
                     
            Set WObj = Nothing
            
            Call LogSGa("CLI  ", "ENVIO_SMS_CLIE", " ENVIO SMS CLIENTES  ", Captura_Conteudo_Campo_RESP(Usuario_Ativo(), "nome"), Date, Time(), "CEL " & WTelefone_Cel, Environ("ComputerName"))
         End If
         X = X + 1
         If X Mod Retorna_Quant_SMS_Remessa() = 0 Then
            CommitTrans
            'sleep (Retorna_Interv_Remessa())
            BeginTrans
         End If
         
         tbc.MoveNext
   Wend
End If

Finaliza:
  CommitTrans
  'MsgBox "Kbo", vbCritical, "Sga"
  tbc.Close
  bdc.Close
  Set tbc = Nothing
  Set bdc = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Enviar_Email_Lote_Placas(WAssunto As String, WCorpo As String, WRemetente As String, WEmail_Dest As String, WAnexo As String) As String
DoCmd.Hourglass True
On Error GoTo TratamentodeErro:
Set oemail = New CanadaUtils.Email

oemail.NomeRemetente = Captura_Conteudo_Campos_Empresa(Form_Menu.WEmpresa, "RazaoSocial")
oemail.Assunto = WAssunto
oemail.ContaEMail = Conta_Email()
oemail.EmailDestino = WEmail_Dest
oemail.Mensagem = WCorpo
oemail.SenhaEmail = Senha_Conta_Email()
oemail.SMTP = Servidor_Conta_Email()

Dim outils As CanadaConnectivity.FuncoesDiversas
Set outils = New CanadaConnectivity.FuncoesDiversas

If outils.EmailAddressCheck(EmailAlternativo) = True Then
   oemail.EmailAlternativo = EmailAlternativo
End If

If WAnexo = "#" Then
   oemail.Anexo = LogomarcaEmpresa()
Else
   oemail.Anexo = WAnexo
End If
If oemail.ValidaEnvioEmail = True Then
   oemail.EnviarEmail
End If
Enviar_Email_Lote_Placas = oemail.MensagemEnvio
GoSub Finaliza
Finaliza:
   
DoCmd.Hourglass False
Set oemail = Nothing
Set outils = Nothing
Exit Function
Return
TratamentodeErro:
    DoCmd.Hourglass False
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGv"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Retorna_Quant_Plac_Cli_Venc_Conect(WCodigoCliente As Long, WDataF As Variant) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb  As Recordset
Dim WDataR As Variant
Dim WdataC As Variant
Dim X As Long
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("PLACXCLI", DB_OPEN_TABLE)
tb.index = "CodigoCliente"

tb.Seek "=", WCodigoCliente

If Not tb.NoMatch Then
   While tb!CodigoCliente = WCodigoCliente
         If Not IsNull(tb!DataValidadeResultado) And tb!DataValidadeResultado <> "" Then
            WDataR = DateValue(Format$(tb!DataValidadeResultado, "dd/mm/yyyy"))
         Else
            WDataR = DateValue(Format$("01/01/2001", "dd/mm/yyyy"))
          
         End If
         
         If Not IsNull(tb!DataConsulta) And tb!DataConsulta <> "" Then
            WdataC = DateValue(Format$(tb!DataConsulta, "dd/mm/yyyy"))
         Else
            WdataC = DateValue(Format$("01/01/2001", "dd/mm/yyyy"))
          
         End If
   
         If WdataC <= WDataF Then
            If WDataR <= WDataF Then
               X = X + 1
            End If
         End If
          
         tb.MoveNext
         
         If tb.EOF Then
            GoSub Finaliza:
         End If
   
   Wend
End If


Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Retorna_Quant_Plac_Cli_Venc_Conect = X
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Retorna_Modelo_Padrao_SMS(WCodigo As Long) As String
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb  As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("MODELOS_SMS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", WCodigo
If tb.NoMatch Then
   Retorna_Modelo_Padrao_SMS = "#"
Else
   Retorna_Modelo_Padrao_SMS = Identif_SMS() & " : " & tb!Descricao
End If
Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Public Function AtualizaVencimentoPlacas() As Integer
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tbOS  As Recordset
Dim tbPL  As Recordset
Dim X As Long
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tbOS = bd.OpenRecordset("OS", DB_OPEN_TABLE)
Set tbPL = bd.OpenRecordset("PLACXCLI", DB_OPEN_TABLE)

tbOS.index = "PrimaryKey"
tbPL.index = "PrimaryKey"
X = 0

BeginTrans
If tbOS.RecordCount = 0 Then
   GoSub Finaliza
Else
   While Not tbOS.EOF
         If Not IsNull(tbOS!CodigoCliente) And Not IsNull(tbOS!PlacaVeiculo) Then
            GoSub AtualizaPlaca
          End If
          X = X + 1
          If X Mod 1000 = 0 Then
             CommitTrans
             BeginTrans
          End If
          tbOS.MoveNext
    Wend
End If

Finaliza:
    'MsgBox "Atualizacao Realizada Com Sucesso.", vbInformation
    CommitTrans
    tbOS.Close
    tbPL.Close
    bd.Close
    Set tbOS = Nothing
    Set tbPL = Nothing
    Set bd = Nothing
    Exit Function
Return

AtualizaPlaca:
   tbPL.Seek "=", tbOS!PlacaVeiculo, tbOS!CodigoCliente
   If Not tbPL.NoMatch Then
      If Not IsDate(tbPL!Data_Ultima_Consulta) Or tbPL!Data_Validade_Resultado = "01/01/2001" Then
     ' If Not IsDate(tbPL!Data_Ultima_Consulta) Then
          tbPL.Edit
          tbPL!ValidadeCertificado1 = IIf(Not IsDate(tbOS!datafechamento), DateAdd("yyyy", 2, tbOS!DataAbertura), DateAdd("yyyy", 2, tbOS!datafechamento))
          tbPL!ValidadeCertificado2 = IIf(Not IsDate(tbOS!datafechamento), DateAdd("yyyy", 2, tbOS!DataAbertura), DateAdd("yyyy", 2, tbOS!datafechamento))
          tbPL!DataValidadeResultado = IIf(Not IsDate(tbOS!datafechamento), DateAdd("yyyy", 2, tbOS!DataAbertura), DateAdd("yyyy", 2, tbOS!datafechamento))
          tbPL!Data_Validade_Resultado = IIf(Not IsDate(tbOS!datafechamento), DateAdd("yyyy", 2, tbOS!DataAbertura), DateAdd("yyyy", 2, tbOS!datafechamento))
          tbPL.Update
     End If
  End If
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function AtualizaCadastroClientexConsultaLotePlacas(ByVal WCodigoCliente As Integer, ByVal WEmail As String, ByVal WCelular As String) As Boolean
On Error GoTo TratamentodeErro:

Dim bd1 As DataBase
Dim Cli As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Cli = bd1.OpenRecordset("CLI", DB_OPEN_TABLE)
Cli.index = "PrimaryKey"
Cli.Seek "=", WCodigoCliente
BeginTrans
If Not Cli.NoMatch Then
   Cli.Edit
   Cli!EndElet = WEmail
   Cli!Celular = WCelular
   Cli.Update
End If

AtualizaCadastroClientexConsultaLotePlacas = True
CommitTrans
Cli.Close
bd1.Close
Set Cli = Nothing
Set bd1 = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function
Public Function RetornaDescricaoFormaPagtoXCondPagto(ByVal wCodPagto As Long) As String

Dim bd1 As DataBase
Dim CondPagto As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set CondPagto = bd1.OpenRecordset("CONDPAGTO", DB_OPEN_TABLE)
CondPagto.index = "PrimaryKey"
CondPagto.Seek "=", wCodPagto

If CondPagto.NoMatch Then
   RetornaDescricaoFormaPagtoXCondPagto = "INDEFINIDO"
Else
   If IsNull(CondPagto!FormaPagto) Then
       RetornaDescricaoFormaPagtoXCondPagto = "INDEFINIDO"
   ElseIf CondPagto!FormaPagto = "DIN" Then
       RetornaDescricaoFormaPagtoXCondPagto = "DINHEIRO"
   ElseIf CondPagto!FormaPagto = "CAR" Then
       RetornaDescricaoFormaPagtoXCondPagto = "CARTAO DE CREDITO"
   ElseIf CondPagto!FormaPagto = "DEB" Then
       RetornaDescricaoFormaPagtoXCondPagto = "CARTAO DE DEBITO"
   ElseIf CondPagto!FormaPagto = "FIN" Then
       RetornaDescricaoFormaPagtoXCondPagto = "FINANCIAMENTO"
   ElseIf CondPagto!FormaPagto = "CHQ" Then
       RetornaDescricaoFormaPagtoXCondPagto = "CHEQUE"
   ElseIf CondPagto!FormaPagto = "CHP" Then
       RetornaDescricaoFormaPagtoXCondPagto = "CHEQUE PRE"
   ElseIf CondPagto!FormaPagto = "OUT" Then
       RetornaDescricaoFormaPagtoXCondPagto = "OUTROS"
   ElseIf CondPagto!FormaPagto = "PIX" Then
       RetornaDescricaoFormaPagtoXCondPagto = "PIX"
   ElseIf CondPagto!FormaPagto = "TRA" Then
       RetornaDescricaoFormaPagtoXCondPagto = "TRANSFERENCIA BANCARIA"
   ElseIf CondPagto!FormaPagto = "DEP" Then
       RetornaDescricaoFormaPagtoXCondPagto = "DEPOSITO EM CONTA"
       
   Else
      RetornaDescricaoFormaPagtoXCondPagto = "OUTROS"
   End If
End If

Finaliza:
   CondPagto.Close
   bd1.Close
   Set CondPagto = Nothing
   Set bd1 = Nothing
   
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function
'Fun��o que identifica a exist�ncia do arquivo
Public Function VerificaArquivo(ByVal lStr As String) As Boolean
Dim outils As CanadaConnectivity.FuncoesDiversas
Set outils = New CanadaConnectivity.FuncoesDiversas
VerificaArquivo = True
If outils.CheckFile(lStr) = False Then
    VerificaArquivo = False
Else
    VerificaArquivo = True
End If
Set outils = Nothing
    'VerificaArquivo = True
    'Identifica se o arquivo existe
    'If Dir(lStr) = vbNullString Then
    '    VerificaArquivo = False
   ' Else
    '    VerificaArquivo = True
    'End If
End Function
Public Function ImagemIndisponivel() As String
ImagemIndisponivel = "C:\SGA\ImagemIndisponivel.bmp"
End Function
Public Function CaminhoArquivoCRLV() As String
CaminhoArquivoCRLV = CapturaCaminhoBD() & "\CRLV\" & RetornaArquivoCRLV(Form_RegistrosInMetro!Numos) & ".pdf"
End Function
Public Function RetornaArquivoCRLV(ByVal WNumos As Long) As String

Dim bd1 As DataBase
Dim oOS As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set oOS = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
oOS.index = "PrimaryKey"
oOS.Seek "=", WNumos

If oOS.NoMatch Then
   RetornaArquivoCRLV = Empty
Else
   RetornaArquivoCRLV = IIf(IsNull(oOS!ArquivoCRLV), Empty, oOS!ArquivoCRLV)
End If

oOS.Close
bd1.Close
Set oOS = Nothing
Set bd1 = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Public Function AtualizaArquivoCRLV(ByVal WNumos As Long, ByVal wNomeArquivo As String) As Integer

Dim bd1 As DataBase
Dim oOS As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set oOS = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
oOS.index = "PrimaryKey"
oOS.Seek "=", WNumos
If oOS.NoMatch Then
   AtualizaArquivoCRLV = 0
Else
   BeginTrans
   oOS.Edit
   oOS!ArquivoCRLV = wNomeArquivo
   oOS.Update
   CommitTrans
   AtualizaArquivoCRLV = 1
End If
oOS.Close
bd1.Close
Set oOS = Nothing
Set bd1 = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Public Function AtualizaaVelocidadeEnsaio(ByVal WNumos As Long, ByVal wVelocidade As String) As Integer

Dim bd1 As DataBase
Dim oOS As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set oOS = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
oOS.index = "PrimaryKey"
oOS.Seek "=", WNumos
If oOS.NoMatch Then
   AtualizaaVelocidadeEnsaio = 0
Else
   BeginTrans
   oOS.Edit
   oOS!VelMaxEns = wVelocidade
   oOS.Update
   CommitTrans
   AtualizaaVelocidadeEnsaio = 1
End If
oOS.Close
bd1.Close
Set oOS = Nothing
Set bd1 = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Public Function AtualizaNCMProdutos() As Boolean
Dim bd1 As DataBase
Dim oest As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set oest = bd1.OpenRecordset("EST", DB_OPEN_TABLE)
Dim cont As Integer
cont = 0
If oest.RecordCount > 0 Then
    BeginTrans
    While Not oest.EOF
        oest.Edit
        'oest!CodigoNCM = "90299010"
        oest.Update
        If cont Mod 1000 = 0 Then
            CommitTrans
            BeginTrans
        End If
        oest.MoveNext
    Wend
End If
CommitTrans
oest.Close
bd1.Close
Set oest = Nothing
Set bd1 = Nothing
MsgBox "Atualizacao Realizada Com Sucesso", vbInformation
AtualizaNCMProdutos = True
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Public Function AtualizaCESTxNCMS() As Boolean
Dim bd1 As DataBase
Dim oNCM As Recordset
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set oNCM = bd1.OpenRecordset("NCM", DB_OPEN_TABLE)
Dim cont As Integer
cont = 0
BeginTrans
If oNCM.RecordCount > 0 Then
    While Not oNCM.EOF
        oNCM.Edit
        oNCM!CEST = "0108000"
        oNCM.Update
        If cont Mod 1000 = 0 Then
            CommitTrans
            BeginTrans
        End If
        oNCM.MoveNext
    Wend
End If
CommitTrans
oNCM.Close
bd1.Close
Set oNCM = Nothing
Set bd1 = Nothing
AtualizaCESTxNCMS = True
MsgBox "Atualizacao Realizada Com Sucesso", vbInformation
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function ApuraExisteChassixOSVxCliente(WCodigoCliente As Long, wPlaca As String) As Boolean
 On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim tb As Recordset
    
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd1.OpenRecordset("PLACXCLI", DB_OPEN_TABLE)
    
    tb.index = "PrimaryKey"
    tb.Seek "=", wPlaca, WCodigoCliente
        
    If Not tb.NoMatch Then
      ApuraExisteChassixOSVxCliente = True
    Else
     ApuraExisteChassixOSVxCliente = False
    End If
        
Finaliza:
    tb.Close
    bd1.Close
    Set tb = Nothing
    Set bd1 = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function
Public Function AtualizaChassiPlacaxCli(wPlaca As String, WCodigoCliente As Long, wChassi As String) As Integer
On Error GoTo TratamentodeErro:
    Dim bd1 As DataBase
    Dim tb As Recordset
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set tb = bd1.OpenRecordset("PLACXCLI", DB_OPEN_TABLE)
    tb.index = "PrimaryKey"
    tb.Seek "=", wPlaca, WCodigoCliente
    BeginTrans
    If Not tb.NoMatch Then
         tb.Edit
         tb!CHASSI = wChassi
         tb.Update
    End If
Finaliza:
    CommitTrans
    tb.Close
    bd1.Close
    Set tb = Nothing
    Set bd1 = Nothing
    Exit Function
Return
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function ApuraConfirmaTacografosOs(WNumos As Long) As Boolean
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("TACOGRAFOSOS", DB_OPEN_TABLE)
tb.index = "SecondKey"
tb.Seek "=", WNumos
If tb.NoMatch Then
  ApuraConfirmaTacografosOs = False
Else
  ApuraConfirmaTacografosOs = tb!Confirma
End If
tb.Close
bd.Close
Set tb = Nothing
Set bd = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    
    End If
End Function
Function AtualizaConfirmaTacografosOs(WNumos As Long, WStatus As Boolean) As Boolean
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("TACOGRAFOSOS", DB_OPEN_TABLE)
tb.index = "SecondKey"
tb.Seek "=", WNumos
BeginTrans
If Not tb.NoMatch Then
  tb.Edit
  tb!Confirma = WStatus
  tb.Update
End If
CommitTrans
tb.Close
bd.Close
Set tb = Nothing
Set bd = Nothing
Exit Function
TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    
    End If
End Function



Function Retorna_Dados_Equipamentos(WEquip As Long, WModelo As String, WMarca As String, WNroSerie As String, WIndex As Long) As Variant

On Error GoTo TratamentodeErro:

Dim db As DataBase
Dim Qtb As QueryDef
Dim tb As Recordset

Set db = CurrentDb()
Set Qtb = db.QueryDefs("WS_Retorna_Dados_Equipamento")
Qtb.Parameters(0) = WEquip
Set tb = Qtb.OpenRecordset()
If tb.RecordCount <> 0 Then
   While Not tb.EOF
   
         If tb!ModeloEns = WModelo And tb!MarcaEns = WMarca Then
            If WIndex = 1 Then
               Retorna_Dados_Equipamentos = tb!CertCalibEns
            ElseIf WIndex = 2 Then
                   Retorna_Dados_Equipamentos = tb!DtUltCalibr
            ElseIf WIndex = 3 Then
                   Retorna_Dados_Equipamentos = tb!Data_Vencimento
            ElseIf WIndex = 4 Then
                   If tb!Data_Vencimento < Date Then
                      Retorna_Dados_Equipamentos = "Certificado Invalido"
                   ElseIf tb!Data_Vencimento > Date Then
                          Retorna_Dados_Equipamentos = "Certificado Valido"
                   Else
                       Retorna_Dados_Equipamentos = "Certificado Invalido"
                   End If
            Else
               Retorna_Dados_Equipamentos = "01/01/2001"
            End If
         End If
          
         tb.MoveNext
      
   Wend
End If

Finaliza:
 tb.Close
 db.Close
 Set tb = Nothing
 Set db = Nothing
 Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGv"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume
    End If





End Function

Function View_Programadores(WStatus As String) As Long

On Error GoTo TratamentodeErro:

Dim db As DataBase
Dim Qtb As QueryDef
Dim tb As Recordset

Set db = CurrentDb()
Set Qtb = db.QueryDefs("RMD_Validade_Equipamentos")
Set tb = Qtb.OpenRecordset()
If tb.RecordCount <> 0 Then
   While Not tb.EOF
   
   '      Call Atualiza_View_Programadores(tb!Modelo, tb!Marca, tb!NroSerieEns, tb!CertCalibEns, tb!�ltimoDeDtUltCalibr, WStatus)
         
         tb.MoveNext
      
   Wend
End If

Finaliza:
 tb.Close
 db.Close
 Set tb = Nothing
 Set db = Nothing
 Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGv"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume
    End If


End Function

Function Atualiza_View_Programadores(WModelo As String, WMarca As String, WNroSerie As String, WCertCalib As String, WDtUltCalibr As Variant, WStatus As String) As Long

On Error GoTo TratamentodeErro:

Dim db As DataBase
Dim tb As Recordset

Set db = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = db.OpenRecordset("MODXENS", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

tb.Seek "=", WModelo, WMarca, WNroSerie, WCertCalib, WDtUltCalibr
If Not tb.NoMatch Then
   BeginTrans
   tb.Edit
   tb!Exibir_Vencto = WStatus
   tb.Update
   CommitTrans
End If

Finaliza:
 tb.Close
 db.Close
 Set tb = Nothing
 Set db = Nothing
 Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGv"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume
    End If



End Function

Function Retorna_Saldo_Dias_Programadores() As Long

On Error GoTo TratamentodeErro:

Call CriaCampoTabela("MODXENS", "Exibir_Vencto", 2, 1)

Dim db As DataBase
Dim Qtb As QueryDef
Dim tb As Recordset
Dim X As Long

Set db = CurrentDb()
Set Qtb = db.QueryDefs("RMD_Validade_Equipamentos")
Set tb = Qtb.OpenRecordset()
X = 0
If tb.RecordCount <> 0 Then
   While Not tb.EOF
   
        If tb!WDias < 30 Then
           If tb!Exibir_Vencto = "S" Or tb!Exibir_Vencto = "" Or IsNull(tb!Exibir_Vencto) Then
              X = tb!WDias
              GoSub Finaliza
           End If
        End If
        tb.MoveNext
      
   Wend
End If

Finaliza:
 tb.Close
 db.Close
 Set tb = Nothing
 Set db = Nothing
 If X > 0 Then
    DoCmd.OpenForm "View_Programadores", , , , , acDialog
 End If
 Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGv"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function

Function RetornaValueMarcaVeiculo(WDescricao As String) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("MARC", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tb.Seek "=", WDescricao
If Not tb.NoMatch Then
   If Not IsNull(tb!Value) Then
      RetornaValueMarcaVeiculo = tb!Value
   Else
      RetornaValueMarcaVeiculo = 0
   End If
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Inclui_OSV_CHK_TQ(WNumos As Long) As Long

On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OSV_CHK_TQ", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
BeginTrans

tb.Seek "=", WNumos, "CHK_TQ"
If tb.NoMatch Then
   tb.AddNew
Else
   tb.Edit
End If
tb!Numos = WNumos
tb!Nro_Tq_Chk = "CHK_TQ"
tb.Update

Finaliza:
  CommitTrans
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Aplica_Desconto_Acima_Limite_Item_Requisitado(WCodItem As Long, WNumrequisicao As Long, WEmpresa As Long, WDesconto_Conc As Double) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("RMXEMPXEST", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
BeginTrans

tb.Seek "=", WNumrequisicao, WEmpresa, WCodItem
If Not tb.NoMatch Then
   tb.Edit
   tb!Desconto = WDesconto_Conc
   tb!DESC_AUT = "S"
   tb.Update
End If

Finaliza:
  CommitTrans
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function



Function Inclui_Depto_Resp(WCodigo As Long, WDescricao As String) As Long

On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("DEPTORESP", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
BeginTrans

tb.Seek "=", WCodigo
If tb.NoMatch Then
   tb.AddNew
Else
   tb.Edit
End If
tb!CodigoDepto = WCodigo
tb!Descricao = WDescricao
tb!Selecionar = -1
tb.Update

Finaliza:
  CommitTrans
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If



End Function

Function Inclui_OCORR_SERV(WCodigo As Long, WDescricao As String, WGera_Fin As Double) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OCORR_SERV", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
BeginTrans

tb.Seek "=", WCodigo
If tb.NoMatch Then
   tb.AddNew
Else
   tb.Edit
End If
tb!Codigo = WCodigo
tb!Descricao = WDescricao
tb!Gera_Financeiro = WGera_Fin
tb.Update

Finaliza:
  CommitTrans
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Padroniza_Modelo_Veic(WMarca As String, WModelo As String, WValue As String) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase, tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("MODELOVEIC", DB_OPEN_TABLE)
tb.index = "ChavePrimaria"

BeginTrans
tb.Seek "=", WMarca, WModelo
If tb.NoMatch Then
   tb.AddNew
Else
   tb.Edit
End If
tb!DescricaoMarca = WMarca
tb!DescricaoModelo = WModelo
tb!Value = WValue
tb.Update
          
Finaliza:
  CommitTrans
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Padroniza_Modelo_Equip(ByVal WMarca As String, ByVal WModelo As String, ByVal wMinAdesivos As Integer, ByVal wMinPlasticos As Integer, ByVal wMaxAdesivos As Integer, ByVal wMaxPlasticos As Integer, ByVal WTipo As String) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase, tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("MODELOEQUIP", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
BeginTrans
tb.Seek "=", WModelo, WMarca
If tb.NoMatch Then
   tb.AddNew
Else
   tb.Edit
End If

tb!Modelo = WModelo
tb!Marca = WMarca
tb!QtdeSelosPlasticos = wMinPlasticos
tb!QtdeSelosAdesivos = wMinAdesivos
tb!Plasticos_Maximo = wMaxPlasticos
tb!Adesivos_Maximo = wMaxAdesivos
tb!Tipo = WTipo
tb.Update
          
Finaliza:
  CommitTrans
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Inicializa_Modelo_Equip() As Long

On Error GoTo TratamentodeErro:
Dim bd As DataBase, tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("MODELOEQUIP", DB_OPEN_TABLE)

tb.index = "PrimaryKey"
GoSub Inicializa_Tabela:

Finaliza:
  CommitTrans
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

Inicializa_Tabela:
  BeginTrans
  If Not tb.NoMatch Then
    While Not tb.EOF
        tb.Delete
        tb.MoveNext
        If tb.EOF Then
            GoTo Finaliza
        End If
    Wend
  End If
  GoTo Finaliza
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
End Function

Function Inclui_Modelo_Selos(WCodigo As String, WDescricao As String, WValue As String) As Long

On Error GoTo TratamentodeErro:
Dim bd As DataBase, tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("MODELO_SELOS", DB_OPEN_TABLE)

tb.index = "PrimaryKey"
GoSub Inclui_Registro:

Finaliza:
  CommitTrans
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

Inclui_Registro:
  BeginTrans
  tb.Seek "=", WCodigo
  If tb.NoMatch Then
     tb.AddNew
  Else
     tb.Edit
  End If
  tb!Codigo = WCodigo
  tb!Descricao = WDescricao
  tb!WValue = WValue
  tb.Update
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function
Function Inicializa_Aro_Pneu_Veic() As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase, tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("AROPNEU", DB_OPEN_TABLE)

GoSub Inicializa_Tabela:

Finaliza:
 
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

Inicializa_Tabela:
  BeginTrans
  If tb.RecordCount <> 0 Then
     While Not tb.EOF
           tb.Delete
           tb.MoveNext
     Wend
  End If
  CommitTrans
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function
Function Load_ResumoVendasPeriodo(ByVal WDataInicial As Date, ByVal WDataFinal As Date) As Long

On Error GoTo TratamentodeErro:

Dim tbResumoVendas, tbOS, tbRQM, tbRMXEMPXEST, tbItensServOSV As Recordset, bd, bd1 As DataBase
Dim pEmpresa, pDepto, pNumOS, pDataAbertura, pDataFechamento, pAtendente, pCNPJCPF, pRAzaoSocial, pPlaca, pPrazo, pNFe, pNFSe, pTipoItem, pCategoriaItem, pDescricaoItem, pVendedor As String
Dim pValorUnit, pValorDesc, pQuant As Double
Dim X As Long

Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tbResumoVendas = bd.OpenRecordset("ResumoVendasPeriodo", DB_OPEN_TABLE)

Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tbOS = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
Set tbRQM = bd1.OpenRecordset("RQM", DB_OPEN_TABLE)
Set tbRMXEMPXEST = bd1.OpenRecordset("RMXEMPXEST", DB_OPEN_TABLE)
Set tbItensServOSV = bd1.OpenRecordset("ITENSSERVOSV", DB_OPEN_TABLE)

tbItensServOSV.index = "SecondaryKey"
tbOS.index = "DataFechamento"

'BeginTrans

If tbResumoVendas.RecordCount <> 0 Then Call Inicializa_ResumoVendasPeriodo
   tbOS.Seek ">=", WDataInicial

If Not tbOS.NoMatch Then
   
    While tbOS!datafechamento >= WDataInicial And tbOS!datafechamento <= WDataFinal
        
        If tbOS!CodEmpresa = Form_Menu.WEmpresa Then
    
            pEmpresa = UCase(Retorna_RazaoSocial_Empresa(Empresa_Ativa))
            pDepto = UCase(Captura_Descricao_Depto_Resp(tbOS!CodigoDepto))
            pNumOS = CLng(tbOS!Numos)
            pDataAbertura = CDate(tbOS!DataAbertura & " " & tbOS!HoraAbertura)
            pDataFechamento = CDate(tbOS!datafechamento & " " & tbOS!HoraFechamento)
            pAtendente = UCase(RetornaDescricaoResponsavel(tbOS!Responsavel))
            pCNPJCPF = CapturaCNPJCli(tbOS!CodigoCliente)
            pRAzaoSocial = UCase(CapturaRazaoSocialCli(tbOS!CodigoCliente))
            pPlaca = tbOS!PlacaVeiculo
            pPrazo = UCase(CapturaDescricaoPrazo(tbOS!CodigoPagto))
            pNFe = tbOS!NFEmitida
            pNFSe = tbOS!NFSEmitida
            pVendedor = RetornaDescricaoResponsavel(IIf(IsNumeric(CapturaVendedorCliente(tbOS!CodigoCliente)), CapturaVendedorCliente(tbOS!CodigoCliente), 0))
            
            tbItensServOSV.Seek "=", tbOS!Numos
            If Not tbItensServOSV.NoMatch Then
               BeginTrans
               X = 0
                Do While tbOS!Numos = tbItensServOSV!Numos
                
                    tbResumoVendas.AddNew
                    tbResumoVendas("EMPRESA") = pEmpresa
                    tbResumoVendas("DEPTO") = pDepto
                    tbResumoVendas("NUMOSV") = pNumOS
                    tbResumoVendas("DATAABERTURA") = pDataAbertura
                    tbResumoVendas("DATAFECHAMENTO") = pDataFechamento
                    tbResumoVendas("ATENDENTE") = pAtendente
                    tbResumoVendas("CNPJCPF") = pCNPJCPF
                    tbResumoVendas("RAZAOSOCIAL") = pRAzaoSocial
                    tbResumoVendas("PLACA") = pPlaca
                    tbResumoVendas("PRAZO") = pPrazo
                    tbResumoVendas("NFE") = pNFe
                    tbResumoVendas("NFSE") = pNFSe
                    tbResumoVendas("TIPOITEM") = "SERVICO"
                    tbResumoVendas("CATEGORIAITEM") = UCase(RetornaDescricaoCATSERV(RetornaCodigoCategoriaxCodServ(tbItensServOSV!TipoServico)))
                    tbResumoVendas("CATEGORIAITEM") = IIf(tbResumoVendas("CATEGORIAITEM") = "" Or IsNull(tbResumoVendas("CATEGORIAITEM")), "INDEFINIDO", tbResumoVendas("CATEGORIAITEM"))
                    tbResumoVendas("DESCRICAOITEM") = UCase(CapturaDescricaoServico(tbItensServOSV!TipoServico))
                    tbResumoVendas("VALORUNIT") = tbItensServOSV!ValorMaoObra
                    tbResumoVendas("QTD") = tbItensServOSV!Quantidade
                    tbResumoVendas("DESCONTO") = tbItensServOSV!Desconto
                    tbResumoVendas("VENDEDORCLIENTE") = pVendedor
                    tbResumoVendas.Update
                    X = X + 1
                    If X Mod 500 = 0 Then
                       BeginTrans
                       CommitTrans
                    End If
                    
                    tbItensServOSV.MoveNext
                    If tbItensServOSV.EOF Then
                        Exit Do
                    End If
                Loop
                CommitTrans
            End If
            
            tbRQM.index = "SecondaryKey"
            tbRQM.Seek "=", tbOS!Numos
            
            If Not tbRQM.NoMatch Then
               BeginTrans
               X = 0
                Do While tbOS!Numos = tbRQM!Numos
                    
                    tbRMXEMPXEST.index = "RQM"
                    tbRMXEMPXEST.Seek "=", tbRQM!NumRequisicao
                    
                    If Not tbRMXEMPXEST.NoMatch Then
                        Do While tbRQM!NumRequisicao = tbRMXEMPXEST!NumRequisicao
                            If (tbRMXEMPXEST!Quantidade - tbRMXEMPXEST!Devolucao) > 0 Then
                                tbResumoVendas.AddNew
                                tbResumoVendas("EMPRESA") = pEmpresa
                                tbResumoVendas("DEPTO") = pDepto
                                tbResumoVendas("NUMOSV") = pNumOS
                                tbResumoVendas("DATAABERTURA") = pDataAbertura
                                tbResumoVendas("DATAFECHAMENTO") = pDataFechamento
                                tbResumoVendas("ATENDENTE") = pAtendente
                                tbResumoVendas("CNPJCPF") = pCNPJCPF
                                tbResumoVendas("RAZAOSOCIAL") = pRAzaoSocial
                                tbResumoVendas("PLACA") = pPlaca
                                tbResumoVendas("PRAZO") = pPrazo
                                tbResumoVendas("NFE") = pNFe
                                tbResumoVendas("NFSE") = pNFSe
                                tbResumoVendas("TIPOITEM") = "PECA"
                                tbResumoVendas("CATEGORIAITEM") = UCase(RetornaDescricaoCATPROD(RetornaCodigoCategoria(tbRMXEMPXEST!CodItem)))
                                tbResumoVendas("DESCRICAOITEM") = UCase(CapturaDescricaoItemEstoque(tbRMXEMPXEST!CodItem))
                                tbResumoVendas("VALORUNIT") = tbRMXEMPXEST!ValorUnit
                                tbResumoVendas("QTD") = tbRMXEMPXEST!Quantidade - tbRMXEMPXEST!Devolucao
                                tbResumoVendas("DESCONTO") = tbRMXEMPXEST!Desconto
                                tbResumoVendas("VENDEDORCLIENTE") = pVendedor
                                tbResumoVendas.Update
                                X = X + 1
                                If X Mod 500 = 0 Then
                                    BeginTrans
                                    CommitTrans
                                End If
                            End If
                            tbRMXEMPXEST.MoveNext
                            If tbRMXEMPXEST.EOF Then
                                Exit Do
                            End If
                        Loop
                    End If
                                    
                    tbRQM.MoveNext
                    If tbRQM.EOF Then
                        Exit Do
                    End If
                                                   
                Loop
                CommitTrans
            End If
                        
        End If
          
        tbOS.MoveNext
        If tbOS.EOF Then
            GoSub Finaliza
        End If
    Wend
'Else
'    GoTo Finaliza
    
End If


Finaliza:
  'CommitTrans
    tbOS.Close
    tbRQM.Close
    tbRMXEMPXEST.Close
    tbItensServOSV.Close
    bd1.Close
    bd.Close
    Set tbOS = Nothing
    Set tbRQM = Nothing
    Set tbRMXEMPXEST = Nothing
    Set tbItensServOSV = Nothing
    Set bd = Nothing
    Set bd1 = Nothing
    Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
       Resume Next:
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next:
    End If
End Function


Function Inicializa_ResumoVendasPeriodo() As Long
On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set tb = bd.OpenRecordset("ResumoVendasPeriodo", DB_OPEN_TABLE)

BeginTrans

If tb.RecordCount <> 0 Then
   While Not tb.EOF
         tb.Delete
         tb.MoveNext
         If tb.EOF Then
            GoSub Finaliza:
         End If
   Wend
End If

Finaliza:
  CommitTrans
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
       Resume Next
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function
Function Gera_Estoque_Ent_Salc(WCodigoFornec As Long, WNotaFiscal As Long, WCFOP As Long) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb, tba As Recordset
Dim wNomeArquivo As String * 256
Dim wsaldo As Double

Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("ENTRADAS", DB_OPEN_TABLE)
Set tba = bd.OpenRecordset("DETENTRADAS", DB_OPEN_TABLE)
tb.index = "ChavePrimaria"
tba.index = "QuintaChave"

wNomeArquivo = CapturaCaminhoIntegracaoECF() & "\Req\ENV" & RetornaNroSerieECF() & ".EST"

Open wNomeArquivo For Output As #1

tb.Seek "=", WCodigoFornec, WNotaFiscal, WCFOP
If Not tb.NoMatch Then
   While tb!CodigoFornecedor = WCodigoFornec And tb!NotaFiscal = WNotaFiscal And tb!cfo = WCFOP
   
         If tb!Empresa = Form_Menu.WEmpresa Then
            GoSub ApuraItens:
         End If
            
         tb.MoveNext
         If tb.EOF Then
            GoSub Finaliza:
         End If
   Wend
End If

Finaliza:
    Print #1, "9"
    tb.Close
    tba.Close
    Close #1
    bd.Close
    Set tb = Nothing
    Set tba = Nothing
    Set bd = Nothing
    
    Exit Function
Return

ApuraItens:
   tba.Seek "=", tb!CodigoFornecedor, tb!NotaFiscal, tb!cfo
   If Not tba.NoMatch Then
      While tba!CodigoFornecedor = tb!CodigoFornecedor And tba!NotaFiscal = tb!NotaFiscal And tba!CFOP = tb!cfo
      
            GoSub GravaReg11:
  
            tba.MoveNext
            
            If tba.EOF Then
               Return
            End If
      Wend
  End If

Return


GravaReg11:
   wsaldo = (SaldoContabilEstoque(tba!CodItem, Date, Form_Menu.WEmpresa)) * 1000
   If wsaldo < 0 Then
      wsaldo = 0
   End If
   
   Print #1, "1" & Format$(tba!CodItem, "0000000") & Format$(wsaldo, "00000000000000") & Space(177) & "X"
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Gera_Estoque_OSV_Salc(WNumos As Long) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb, tba As Recordset
Dim wNomeArquivo As String * 256
Dim wsaldo As Double

Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("RQM", DB_OPEN_TABLE)
Set tba = bd.OpenRecordset("RMXEMPXEST", DB_OPEN_TABLE)
tb.index = "SecondaryKey"
tba.index = "RQM"

wNomeArquivo = CapturaCaminhoIntegracaoECF() & "\Req\ENV" & RetornaNroSerieECF() & ".EST"

Open wNomeArquivo For Output As #1

tb.Seek "=", WNumos
If Not tb.NoMatch Then
   While tb!Numos = WNumos
   
         GoSub ApuraItens:
            
         tb.MoveNext
         If tb.EOF Then
            GoSub Finaliza:
         End If
   Wend
End If

Finaliza:
    Print #1, "9"
    tb.Close
    tba.Close
    Close #1
    bd.Close
    Set tb = Nothing
    Set tba = Nothing
    Set bd = Nothing
    Exit Function
Return

ApuraItens:
   tba.Seek "=", tb!NumRequisicao
   If Not tba.NoMatch Then
      While tba!NumRequisicao = tb!NumRequisicao
      
            GoSub GravaReg11:
  
            tba.MoveNext
            
            If tba.EOF Then
               Return
            End If
      Wend
  End If

Return


GravaReg11:
   wsaldo = (SaldoContabilEstoque(tba!CodItem, Date, Form_Menu.WEmpresa)) * 1000
   If wsaldo < 0 Then
      wsaldo = 0
   End If
   
   Print #1, "1" & Format$(tba!CodItem, "0000000") & Format$(wsaldo, "00000000000000") & Space(177) & "X"
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If


End Function


Function Total_ICMS_ST_OSV(Numos As Long) As Double

   On Error GoTo TratamentodeErro:

    Dim bd1 As DataBase
    Dim Requisicao, saidas As Recordset
    
    Dim SomaTotal As Double
    Dim NumOSV As Double
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set Requisicao = bd1.OpenRecordset("RQM", DB_OPEN_TABLE)
    Set saidas = bd1.OpenRecordset("RMXEMPXEST", DB_OPEN_TABLE)
    
    Requisicao.index = "SecondaryKey"
    saidas.index = "RQM"
    SomaTotal = 0
    NumOSV = Numos
    
    Requisicao.Seek "=", Numos
    If Not Requisicao.NoMatch Then
       While Requisicao!Numos = Numos
             GoSub ApuraSaidas:
             Requisicao.MoveNext
             If Requisicao.EOF Then
                GoSub Finaliza:
             End If
       Wend
    End If
    
Finaliza:
    Total_ICMS_ST_OSV = SomaTotal
    saidas.Close
    Requisicao.Close
    bd1.Close
    Set saidas = Nothing
    Set Requisicao = Nothing
    Set bd1 = Nothing
    Exit Function
Return

ApuraSaidas:
  saidas.Seek "=", Requisicao!NumRequisicao
  If Not saidas.NoMatch Then
     While saidas!NumRequisicao = Requisicao!NumRequisicao
           SomaTotal = SomaTotal + ((saidas!Vlr_Icms_ST) * (saidas!Quantidade - saidas!Devolucao))
           saidas.MoveNext
           If saidas.EOF Then
              Return
           End If
      Wend
   End If
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Total_BASE_CALC_ST_OSV(Numos As Long) As Double

   On Error GoTo TratamentodeErro:

    Dim bd1 As DataBase
    Dim Requisicao, saidas As Recordset
    
    Dim SomaTotal As Double
    Dim NumOSV As Double
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set Requisicao = bd1.OpenRecordset("RQM", DB_OPEN_TABLE)
    Set saidas = bd1.OpenRecordset("RMXEMPXEST", DB_OPEN_TABLE)
    
    Requisicao.index = "SecondaryKey"
    saidas.index = "RQM"
    SomaTotal = 0
    NumOSV = Numos
    
    Requisicao.Seek "=", Numos
    If Not Requisicao.NoMatch Then
       While Requisicao!Numos = Numos
             GoSub ApuraSaidas:
             Requisicao.MoveNext
             If Requisicao.EOF Then
                GoSub Finaliza:
             End If
       Wend
    End If
    
Finaliza:
    Total_BASE_CALC_ST_OSV = SomaTotal
    saidas.Close
    Requisicao.Close
    bd1.Close
    Set saidas = Nothing
    Set Requisicao = Nothing
    Set bd1 = Nothing
    Exit Function
Return

ApuraSaidas:
  saidas.Seek "=", Requisicao!NumRequisicao
  If Not saidas.NoMatch Then
     While saidas!NumRequisicao = Requisicao!NumRequisicao
           SomaTotal = SomaTotal + ((saidas!Base_Calc_ST) * (saidas!Quantidade - saidas!Devolucao))
           saidas.MoveNext
           If saidas.EOF Then
              Return
           End If
      Wend
   End If
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Total_PRODUTOS_OSV(Numos As Long) As Double

   On Error GoTo TratamentodeErro:

    Dim bd1 As DataBase
    Dim Requisicao, saidas As Recordset
    
    Dim SomaTotal As Double
    Dim NumOSV As Double
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set Requisicao = bd1.OpenRecordset("RQM", DB_OPEN_TABLE)
    Set saidas = bd1.OpenRecordset("RMXEMPXEST", DB_OPEN_TABLE)
    
    Requisicao.index = "SecondaryKey"
    saidas.index = "RQM"
    SomaTotal = 0
    NumOSV = Numos
    
    Requisicao.Seek "=", Numos
    If Not Requisicao.NoMatch Then
       While Requisicao!Numos = Numos
             GoSub ApuraSaidas:
             Requisicao.MoveNext
             If Requisicao.EOF Then
                GoSub Finaliza:
             End If
       Wend
    End If
    
Finaliza:
   Total_PRODUTOS_OSV = SomaTotal
    saidas.Close
    Requisicao.Close
    bd1.Close
    Set saidas = Nothing
    Set Requisicao = Nothing
    Set bd1 = Nothing
    Exit Function
Return

ApuraSaidas:
  saidas.Seek "=", Requisicao!NumRequisicao
  If Not saidas.NoMatch Then
     While saidas!NumRequisicao = Requisicao!NumRequisicao
           SomaTotal = SomaTotal + ((saidas!ValorUnit) * (saidas!Quantidade - saidas!Devolucao))
           saidas.MoveNext
           If saidas.EOF Then
              Return
           End If
      Wend
   End If
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function


Function Total_FRETE_OSV(Numos As Long) As Double

   On Error GoTo TratamentodeErro:

    Dim bd1 As DataBase
    Dim Requisicao, saidas As Recordset
    
    Dim SomaTotal As Double
    Dim NumOSV As Double
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set Requisicao = bd1.OpenRecordset("RQM", DB_OPEN_TABLE)
    Set saidas = bd1.OpenRecordset("RMXEMPXEST", DB_OPEN_TABLE)
    
    Requisicao.index = "SecondaryKey"
    saidas.index = "RQM"
    SomaTotal = 0
    NumOSV = Numos
    
    Requisicao.Seek "=", Numos
    If Not Requisicao.NoMatch Then
       While Requisicao!Numos = Numos
             GoSub ApuraSaidas:
             Requisicao.MoveNext
             If Requisicao.EOF Then
                GoSub Finaliza:
             End If
       Wend
    End If
    
Finaliza:
   Total_FRETE_OSV = SomaTotal
    saidas.Close
    Requisicao.Close
    bd1.Close
    Set saidas = Nothing
    Set Requisicao = Nothing
    Set bd1 = Nothing
    Exit Function
Return

ApuraSaidas:
  saidas.Seek "=", Requisicao!NumRequisicao
  If Not saidas.NoMatch Then
     While saidas!NumRequisicao = Requisicao!NumRequisicao
           SomaTotal = SomaTotal + ((saidas!Vlr_Fret) * (saidas!Quantidade - saidas!Devolucao))
           saidas.MoveNext
           If saidas.EOF Then
              Return
           End If
      Wend
   End If
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Total_SEGURO_OSV(Numos As Long) As Double

   On Error GoTo TratamentodeErro:

    Dim bd1 As DataBase
    Dim Requisicao, saidas As Recordset
    
    Dim SomaTotal As Double
    Dim NumOSV As Double
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set Requisicao = bd1.OpenRecordset("RQM", DB_OPEN_TABLE)
    Set saidas = bd1.OpenRecordset("RMXEMPXEST", DB_OPEN_TABLE)
    
    Requisicao.index = "SecondaryKey"
    saidas.index = "RQM"
    SomaTotal = 0
    NumOSV = Numos
    
    Requisicao.Seek "=", Numos
    If Not Requisicao.NoMatch Then
       While Requisicao!Numos = Numos
             GoSub ApuraSaidas:
             Requisicao.MoveNext
             If Requisicao.EOF Then
                GoSub Finaliza:
             End If
       Wend
    End If
    
Finaliza:
   Total_SEGURO_OSV = SomaTotal
    saidas.Close
    Requisicao.Close
    bd1.Close
    Set saidas = Nothing
    Set Requisicao = Nothing
    Set bd1 = Nothing
    Exit Function
Return

ApuraSaidas:
  saidas.Seek "=", Requisicao!NumRequisicao
  If Not saidas.NoMatch Then
     While saidas!NumRequisicao = Requisicao!NumRequisicao
           SomaTotal = SomaTotal + ((saidas!Vlr_Seguro) * (saidas!Quantidade - saidas!Devolucao))
           saidas.MoveNext
           If saidas.EOF Then
              Return
           End If
      Wend
   End If
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Total_DESCONTO_OSV(Numos As Long) As Double

   On Error GoTo TratamentodeErro:

    Dim bd1 As DataBase
    Dim Requisicao, saidas As Recordset
    
    Dim SomaTotal As Double
    Dim NumOSV As Double
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set Requisicao = bd1.OpenRecordset("RQM", DB_OPEN_TABLE)
    Set saidas = bd1.OpenRecordset("RMXEMPXEST", DB_OPEN_TABLE)
    
    Requisicao.index = "SecondaryKey"
    saidas.index = "RQM"
    SomaTotal = 0
    NumOSV = Numos
    
    Requisicao.Seek "=", Numos
    If Not Requisicao.NoMatch Then
       While Requisicao!Numos = Numos
             GoSub ApuraSaidas:
             Requisicao.MoveNext
             If Requisicao.EOF Then
                GoSub Finaliza:
             End If
       Wend
    End If
    
Finaliza:
   Total_DESCONTO_OSV = SomaTotal
    saidas.Close
    Requisicao.Close
    bd1.Close
    Set saidas = Nothing
    Set Requisicao = Nothing
    Set bd1 = Nothing
    Exit Function
Return

ApuraSaidas:
  saidas.Seek "=", Requisicao!NumRequisicao
  If Not saidas.NoMatch Then
     While saidas!NumRequisicao = Requisicao!NumRequisicao
           SomaTotal = SomaTotal + ((saidas!Desconto) * (saidas!Quantidade - saidas!Devolucao))
           saidas.MoveNext
           If saidas.EOF Then
              Return
           End If
      Wend
   End If
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Total_OUTRAS_DESP_OSV(Numos As Long) As Double

   On Error GoTo TratamentodeErro:

    Dim bd1 As DataBase
    Dim Requisicao, saidas As Recordset
    
    Dim SomaTotal As Double
    Dim NumOSV As Double
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set Requisicao = bd1.OpenRecordset("RQM", DB_OPEN_TABLE)
    Set saidas = bd1.OpenRecordset("RMXEMPXEST", DB_OPEN_TABLE)
    
    Requisicao.index = "SecondaryKey"
    saidas.index = "RQM"
    SomaTotal = 0
    NumOSV = Numos
    
    Requisicao.Seek "=", Numos
    If Not Requisicao.NoMatch Then
       While Requisicao!Numos = Numos
             GoSub ApuraSaidas:
             Requisicao.MoveNext
             If Requisicao.EOF Then
                GoSub Finaliza:
             End If
       Wend
    End If
    
Finaliza:
   Total_OUTRAS_DESP_OSV = SomaTotal
    saidas.Close
    Requisicao.Close
    bd1.Close
    Set saidas = Nothing
    Set Requisicao = Nothing
    Set bd1 = Nothing
    Exit Function
Return

ApuraSaidas:
  saidas.Seek "=", Requisicao!NumRequisicao
  If Not saidas.NoMatch Then
     While saidas!NumRequisicao = Requisicao!NumRequisicao
           SomaTotal = SomaTotal + ((saidas!Vlr_Out_Desp) * (saidas!Quantidade - saidas!Devolucao))
           saidas.MoveNext
           If saidas.EOF Then
              Return
           End If
      Wend
   End If
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function


Function Total_IPI_OSV(Numos As Long) As Double

   On Error GoTo TratamentodeErro:

    Dim bd1 As DataBase
    Dim Requisicao, saidas As Recordset
    
    Dim SomaTotal As Double
    Dim NumOSV As Double
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set Requisicao = bd1.OpenRecordset("RQM", DB_OPEN_TABLE)
    Set saidas = bd1.OpenRecordset("RMXEMPXEST", DB_OPEN_TABLE)
    
    Requisicao.index = "SecondaryKey"
    saidas.index = "RQM"
    SomaTotal = 0
    NumOSV = Numos
    
    Requisicao.Seek "=", Numos
    If Not Requisicao.NoMatch Then
       While Requisicao!Numos = Numos
             GoSub ApuraSaidas:
             Requisicao.MoveNext
             If Requisicao.EOF Then
                GoSub Finaliza:
             End If
       Wend
    End If
    
Finaliza:
   Total_IPI_OSV = SomaTotal
    saidas.Close
    Requisicao.Close
    bd1.Close
    Set saidas = Nothing
    Set Requisicao = Nothing
    Set bd1 = Nothing
    Exit Function
Return

ApuraSaidas:
  saidas.Seek "=", Requisicao!NumRequisicao
  If Not saidas.NoMatch Then
     While saidas!NumRequisicao = Requisicao!NumRequisicao
           SomaTotal = SomaTotal + ((saidas!Valor_IPI) * (saidas!Quantidade - saidas!Devolucao))
           saidas.MoveNext
           If saidas.EOF Then
              Return
           End If
      Wend
   End If
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function


Function Total_BC_ICMS_OSV(Numos As Long) As Double

   On Error GoTo TratamentodeErro:

    Dim bd1 As DataBase
    Dim Requisicao, saidas As Recordset
    
    Dim SomaTotal As Double
    Dim NumOSV As Double
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set Requisicao = bd1.OpenRecordset("RQM", DB_OPEN_TABLE)
    Set saidas = bd1.OpenRecordset("RMXEMPXEST", DB_OPEN_TABLE)
    
    Requisicao.index = "SecondaryKey"
    saidas.index = "RQM"
    SomaTotal = 0
    NumOSV = Numos
    
    Requisicao.Seek "=", Numos
    If Not Requisicao.NoMatch Then
       While Requisicao!Numos = Numos
             GoSub ApuraSaidas:
             Requisicao.MoveNext
             If Requisicao.EOF Then
                GoSub Finaliza:
             End If
       Wend
    End If
    
Finaliza:
   Total_BC_ICMS_OSV = SomaTotal
    saidas.Close
    Requisicao.Close
    bd1.Close
    Set saidas = Nothing
    Set Requisicao = Nothing
    Set bd1 = Nothing
    Exit Function
Return

ApuraSaidas:
  saidas.Seek "=", Requisicao!NumRequisicao
  If Not saidas.NoMatch Then
     While saidas!NumRequisicao = Requisicao!NumRequisicao
           If saidas!CodigoSituacaoTributaria = 0 Or saidas!CodigoSituacaoTributaria = 10 Or saidas!CodigoSituacaoTributaria = 110 Or saidas!CodigoSituacaoTributaria = 20 Or saidas!CodigoSituacaoTributaria = 90 Then
              SomaTotal = SomaTotal + ((saidas!ValorUnit + saidas!Vlr_Fret + saidas!Vlr_Out_Desp + saidas!Valor_IPI) * (saidas!Quantidade - saidas!Devolucao))
           End If
           saidas.MoveNext
           If saidas.EOF Then
              Return
           End If
      Wend
   End If
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Total_VR_ICMS_OSV(Numos As Long) As Double

   On Error GoTo TratamentodeErro:

    Dim bd1 As DataBase
    Dim Requisicao, saidas As Recordset
    
    Dim SomaTotal As Double
    Dim NumOSV As Double
    Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
    Set Requisicao = bd1.OpenRecordset("RQM", DB_OPEN_TABLE)
    Set saidas = bd1.OpenRecordset("RMXEMPXEST", DB_OPEN_TABLE)
    
    Requisicao.index = "SecondaryKey"
    saidas.index = "RQM"
    SomaTotal = 0
    NumOSV = Numos
    
    Requisicao.Seek "=", Numos
    If Not Requisicao.NoMatch Then
       While Requisicao!Numos = Numos
             GoSub ApuraSaidas:
             Requisicao.MoveNext
             If Requisicao.EOF Then
                GoSub Finaliza:
             End If
       Wend
    End If
    
Finaliza:
   Total_VR_ICMS_OSV = SomaTotal
    saidas.Close
    Requisicao.Close
    bd1.Close
    Set saidas = Nothing
    Set Requisicao = Nothing
    Set bd1 = Nothing
    Exit Function
Return

ApuraSaidas:
  saidas.Seek "=", Requisicao!NumRequisicao
  If Not saidas.NoMatch Then
     While saidas!NumRequisicao = Requisicao!NumRequisicao
           If saidas!CodigoSituacaoTributaria = 0 Or saidas!CodigoSituacaoTributaria = 10 Or saidas!CodigoSituacaoTributaria = 110 Or saidas!CodigoSituacaoTributaria = 20 Or saidas!CodigoSituacaoTributaria = 90 Then
              SomaTotal = SomaTotal + ((((saidas!ValorUnit + saidas!Vlr_Fret + saidas!Vlr_Out_Desp + saidas!Valor_IPI) * (saidas!Quantidade - saidas!Devolucao)) * saidas!Aliq_ICMS) / 100)
           End If
           saidas.MoveNext
           If saidas.EOF Then
              Return
           End If
      Wend
   End If
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Cria_OSV_Salc(WNumos As Long, wCodCliente As Long, WOperacao As Long, WResponsavel As Long, WDataAbertura As Variant, WDepto As Long, wobservacoes As String, WCOOCCF As String) As Long
On Error GoTo TratamentodeErro:
Dim OSV As Recordset, bd1 As DataBase
Set bd1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set OSV = bd1.OpenRecordset("OS", DB_OPEN_TABLE)
OSV.index = "PrimaryKey"
BeginTrans
OSV.Seek "=", WNumos
If OSV.NoMatch Then
   OSV.AddNew
Else
   OSV.Edit
End If

wCodCliente = 1

OSV!Numos = WNumos
OSV!TipoOs = "1"
OSV!CodigoCliente = wCodCliente
OSV!CodigoPagto = WOperacao
OSV!Emitirnf = "1"
OSV!Responsavel = WResponsavel
OSV!DataAbertura = WDataAbertura
OSV!HoraAbertura = Time()
OSV!datafechamento = WDataAbertura
OSV!HoraFechamento = Time()
OSV!DataOrcamento = WDataAbertura
OSV!CodigoDepto = WDepto
OSV!Observacoes = wobservacoes
OSV!COOCCF = WCOOCCF
OSV!CodEmpresa = Form_Menu.WEmpresa
OSV.Update

Finaliza:
  CommitTrans
  OSV.Close
  bd1.Close
  Set OSV = Nothing
  Set bd1 = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Cancela_Cupons_Salc(WCOOCCF As String) As Long

On Error GoTo TratamentodeErro:
Dim tb As Recordset, bd As DataBase
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
tb.index = "COOCCF"

tb.Seek "=", WCOOCCF

If Not tb.NoMatch Then
   Call Cancela_OSV_Salc(tb!Numos, tb!CodigoCliente, tb!datafechamento)
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function Cancela_OSV_Salc(WNumos As Long, WCodigoCliente As Long, WDataEmissao As Variant) As Long

On Error GoTo TratamentodeErro:
DoCmd.Hourglass True
Dim db As DataBase
Dim CTREC, PGCTREC, DEM, RQM, IRQM As Recordset
Set db = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set CTREC = db.OpenRecordset("CTREC", DB_OPEN_TABLE)
Set PGCTREC = db.OpenRecordset("PGCTREC", DB_OPEN_TABLE)
Set DEM = db.OpenRecordset("DEMONSTRATIVORECEITAS", DB_OPEN_TABLE)
Set RQM = db.OpenRecordset("RQM", DB_OPEN_TABLE)
Set IRQM = db.OpenRecordset("RMXEMPXEST", DB_OPEN_TABLE)

CTREC.index = "SetimaChave"
PGCTREC.index = "QuintaChave"
DEM.index = "Primarykey"
RQM.index = "Secondarykey"
IRQM.index = "RQM"

BeginTrans

'CTREC.Seek "=", WNumos, WDataEmissao, WCodigocliente
'If Not CTREC.NoMatch Then
   GoSub DeletaConta:
   GoSub DeletaDEM:
   GoSub EfetuaLeituraRequisicao:
'End If

Finaliza:
  CommitTrans
  CTREC.Close
  PGCTREC.Close
  DEM.Close
  RQM.Close
  IRQM.Close
  db.Close
  Set CTREC = Nothing
  Set PGCTREC = Nothing
  Set DEM = Nothing
  Set RQM = Nothing
  Set IRQM = Nothing
  Set db = Nothing
  DoCmd.Hourglass False
  Exit Function
Return

DeletaDEM:
  
  DEM.Seek "=", WNumos
  If Not DEM.NoMatch Then
     DEM.Edit
     DEM!Dinheiro = 0
     DEM!Moedas = 0
     DEM!ChequesVista = 0
     DEM!ChequesPreDatados = 0
     DEM!Duplicatas = 0
     DEM!Garantia = 0
     DEM!Vale = 0
     DEM.Update
  End If
Return

DeletaPagto:
  PGCTREC.Seek "=", WNumos, WCodigoCliente, CTREC!DataVencimento
  If Not PGCTREC.NoMatch Then
     While Not PGCTREC.EOF
        If PGCTREC!NotaOSCupon = WNumos And PGCTREC!CodigoCliente = WCodigoCliente And PGCTREC!DataVencimento = CTREC!DataVencimento Then
           PGCTREC.Delete
        Else
           Return
        End If
        PGCTREC.MoveNext
        If PGCTREC.EOF Then
           Return
        End If
     Wend
   End If
Return
    
DeletaConta:
    CTREC.Seek "=", WNumos, WDataEmissao, WCodigoCliente
    If Not CTREC.NoMatch Then
       While Not CTREC.EOF
             If CTREC!NotaOSCupon = WNumos And CTREC!CodigoCliente = WCodigoCliente And CTREC!DataEmissao = WDataEmissao Then
                GoSub DeletaPagto
                CTREC.Delete
             Else
                Return
             End If
             CTREC.MoveNext
             If CTREC.EOF Then
                Return
             End If
        Wend
    End If
Return

EfetuaLeituraRequisicao:

   RQM.Seek "=", WNumos
   If Not RQM.NoMatch Then
      While RQM!Numos = WNumos
             
            GoSub DevolveItens:
       
            RQM.MoveNext
            If RQM.EOF Then
               Return
            End If
      Wend
   
   End If

Return

DevolveItens:
   IRQM.Seek "=", RQM!NumRequisicao
   If Not IRQM.NoMatch Then
      While IRQM!NumRequisicao = RQM!NumRequisicao
            IRQM.Edit
            IRQM!Devolucao = IRQM!Quantidade
            IRQM.Update
            IRQM.MoveNext
            If IRQM.EOF Then
               Return
            End If
      Wend
   End If
Return



TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If



End Function

Function Importa_Cupons_Cancelados_Salc(ByVal caminho As String) As Long

'DESVIO TEMPORARIO ATE CORRIGIR PROCESSAMENTE DE RETORNO ...ECF SALC

Exit Function


 On Error GoTo TratamentodeErro:
    'Aten��o: Fa�a refer�ncia � biblioteca Micrsoft Scripting Runtime
    Dim FSO As New FileSystemObject
    Dim X As Long
    Dim WSeq_Reg As Long
    Dim Result() As String
    Dim Pasta As Folder
    Dim arquivo As File
    Dim Indice As Long
    Dim Wcaminho As String
    Dim WReg As String * 100
    Dim WSeq As String * 1
    
    Wcaminho = caminho
 
    ReDim Result(0) As String
    If FSO.FolderExists(Wcaminho) Then
        Set Pasta = FSO.GetFolder(Wcaminho)
        For Each arquivo In Pasta.Files
            Indice = IIf(Result(0) = "", 0, Indice + 1)
            ReDim Preserve Result(Indice) As String
            Result(Indice) = arquivo.Name
            If Len(Result(Indice)) <> 0 Then
               If Mid$(Result(Indice), 13, 3) = "DEL" Then
                  GoSub EfetuaLeitura:
                  Close #2
                  Kill Wcaminho & Result(Indice)
               End If
            End If
        Next
    End If
 
    
Finaliza:
  Set FSO = Nothing
  Set Pasta = Nothing
  Set arquivo = Nothing
  Exit Function
Return
 
 
EfetuaLeitura:
    Open Wcaminho & Result(Indice) For Input As #2
    
    While Not EOF(2)
 
          Line Input #2, WReg ' L� a linha para a vari�vel.
    
          WSeq = Mid$(WReg, 1, 1)
          If WSeq = "8" Then
              Call Cancela_Cupons_Salc(Mid$(WReg, 6, 9))
          End If
   
    Wend
Return
 

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGv"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
    
End Function

Function Importa_Arquivos_Salc(ByVal caminho As String) As Long

'DESVIO TEMPORARIO ATE CORRIGIR PROCESSAMENTE DE RETORNO ...ECF SALC

Exit Function

 On Error GoTo TratamentodeErro:
    'Aten��o: Fa�a refer�ncia � biblioteca Micrsoft Scripting Runtime
    Dim FSO As New FileSystemObject
    Dim X As Long
    Dim WSeq_Reg As Long
    Dim Result() As String
    Dim Pasta As Folder
    Dim arquivo As File
    Dim Indice As Long
    Dim Wcaminho As String
    Dim WReg As String * 100
    Dim WSeq As String * 1
    Dim db As DataBase
    Dim tb, tbi As Recordset
    
    ' Dados OSV
    Dim WProximaOSV As Long
    Dim WCodigoCliente As Long
    Dim WCodigoResp As Long
    Dim WCodigoPagto As Long
    Dim WDataVenda As Variant
    Dim WObserv As String
    Dim WTotalVenda As Double
    Dim WCOOCCF As String
    
    ' Dados Itens
    Dim WCodItem As Long
    Dim WPrecovenda As Double
    Dim WQuantidade As Double
    Dim WSequencia As Long
    Dim WDesconto As Double
    Dim WNumrequisicao As Long
    Dim WAliq_IPI As Double
    Dim WVlr_IPI As Double
    Dim WMOD_BC_ICMS_ST As Long
    Dim WMod_BC_ICMS As Long
    Dim WAliq_Red_BC_ST As Double
    Dim WCodCSTPis As Long
    Dim WCodCSTCofins As Long
    Dim WAliqPis As Double
    Dim WAliqCofins As Double
    Dim WAliqINSS As Double
    Dim WAliqIRPJ As Double
    Dim WAliqCSLL As Double
    Dim WCFOP As Long
    Dim WCST As Long

    Set db = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
    Set tb = db.OpenRecordset("Contr_ECF", DB_OPEN_TABLE)
    Set tbi = db.OpenRecordset("Itens_Contr_ECF", DB_OPEN_TABLE)
    
    tb.index = "PrimaryKey"
    
    WSeq_Reg = 0
    
    GoSub Inicializa_TB:
    GoSub Inicializa_TBI:
    
    Wcaminho = caminho
 
    ReDim Result(0) As String
    If FSO.FolderExists(Wcaminho) Then
        Set Pasta = FSO.GetFolder(Wcaminho)
        For Each arquivo In Pasta.Files
            Indice = IIf(Result(0) = "", 0, Indice + 1)
            ReDim Preserve Result(Indice) As String
            Result(Indice) = arquivo.Name
            If Len(Result(Indice)) <> 0 Then
               If Mid$(Result(Indice), 13, 3) = "ECF" Then
                  GoSub EfetuaLeitura:
               End If
            End If
        Next
    End If
 
    GoSub ImportaVendaSga:
    
Finaliza:
  tb.Close
  tbi.Close
  db.Close
  Set FSO = Nothing
  Set Pasta = Nothing
  Set arquivo = Nothing
  Exit Function
 
Return
  
ImportaVendaSga:
  If tb.RecordCount <> 0 Then
     tb.MoveFirst
     While Not tb.EOF
           WProximaOSV = ProximaOSV(0)
           WSequencia = 0
           WCodigoCliente = 1
           WCodigoResp = 1
           WCodigoPagto = 1
           WDataVenda = tb!datavenda
           WObserv = "ECF: " & Format$(tb!nroecf, "0000")
           WCOOCCF = Format$(tb!nrocoo, "000000000")
 
           Call Cria_OSV_Salc(WProximaOSV, WCodigoCliente, WCodigoPagto, WCodigoResp, WDataVenda, 1, WObserv, WCOOCCF)
                    
           GoSub ImportaItensVendaSga:
           
           Call GeraContasReceber(WProximaOSV, WCodigoPagto, WCodigoCliente, WTotalVenda, WDataVenda, 0, 1, WObserv, 0, 1, 0, 0, 0, "NAO MOVIMENTA", 0, "VENDA VIA COPUN FISCAL", Captura_Conteudo_Campos_CONDPAGTO(WCodigoPagto, "Cart_Vinculada"))
           
           tb.MoveNext
     Wend
  End If
Return

ImportaItensVendaSga:
  tbi.index = "Secondkey"
  tbi.Seek "=", tb!Controle
  If Not tbi.NoMatch Then
     While Not tbi.EOF
           If tbi!Controle <> tb!Controle Then
              Return
           End If
           WCodItem = tbi!CodItem
           WPrecovenda = tbi!ValorUnit
           WQuantidade = tbi!Quantidade
           WSequencia = WSequencia + 1
           WDesconto = tbi!Desconto
           'WNumRequisicao = CriaRequisicao_Atendimento(WProximaOSV)
           
           WAliq_IPI = Captura_Aliq_IPI_EST(tbi!CodItem)
           If WAliq_IPI <> 0 Then
              WPrecovenda = tbi!ValorUnit - ((tbi!ValorUnit * WAliq_IPI) / 100)
              WVlr_IPI = ((tbi!ValorUnit * WAliq_IPI) / 100)
           Else
             WPrecovenda = tbi!ValorUnit
             WVlr_IPI = 0
           End If
           
           WMOD_BC_ICMS_ST = Captura_Conteudo_Campo_Tabelas_Basicas(WCodItem, "EST", "MOD_BC_ICMS_ST", "PrimaryKey")
           WMod_BC_ICMS = Captura_Conteudo_Campo_Tabelas_Basicas(WCodItem, "EST", "MOD_BC_ICMS", "PrimaryKey")
           WAliq_Red_BC_ST = Retorna_Aliq_RED_BC_ST(Captura_CodigoNCM(WCodItem), CapturaUFCadastroCliente(WCodigoCliente))
           WCFOP = CapturaCFOPXEST(WCodItem, WCodigoCliente)
           WCST = Captura_CST_EST(WCodItem)
           If Captura_Conteudo_Campo_Tabelas_Basicas(WCodItem, "EST", "CodigoCST_PIS", "PrimaryKey") = 0 Then
              WCodCSTPis = 4
           Else
              WCodCSTPis = Captura_Conteudo_Campo_Tabelas_Basicas(WCodItem, "EST", "CodigoCST_PIS", "PrimaryKey")
           End If
           If Captura_Conteudo_Campo_Tabelas_Basicas(WCodItem, "EST", "CodigoCST_COFINS", "PrimaryKey") = 0 Then
              WCodCSTCofins = 4
           Else
              WCodCSTCofins = Captura_Conteudo_Campo_Tabelas_Basicas(WCodItem, "EST", "CodigoCST_COFINS", "PrimaryKey")
           End If
          If WCodCSTPis = 4 Then
             WAliqPis = 0
             WAliqCofins = 0
          Else
            WAliqPis = PIS()
            WAliqCofins = COFIS()
          End If
          WAliqINSS = INSS()
          WAliqIRPJ = IRPJ()
          WAliqCSLL = CSSLL()
                      
          WNumrequisicao = Captura_Conteudo_Campos_Empresa(1, "Proxima_Req")
          Call Atualiza_Campos_EMPRESA(1, "Proxima_Req", WNumrequisicao + 1)
          Call CriaRequisicao(WProximaOSV, WNumrequisicao)
          Call IncluiItemRequisitado_Salc(WNumrequisicao, Form_Menu.WEmpresa, WCodItem, WQuantidade, WPrecovenda, WDesconto, WCFOP, tb!datavenda, WVlr_IPI, 0, 0, 0, 0, 0, 0, 0, 0, WMOD_BC_ICMS_ST, WMod_BC_ICMS, WAliq_Red_BC_ST, AliqICMSCupom(), WCodCSTPis, WCodCSTCofins, WAliqPis, WAliqCofins, WAliqINSS, WAliqIRPJ, WAliqCSLL, WCST)
          WTotalVenda = WTotalVenda + (WQuantidade * (WPrecovenda - WDesconto))
               
           tbi.MoveNext
      Wend
  End If
Return
 
EfetuaLeitura:
    Open Wcaminho & Result(Indice) For Input As #1
    
    While Not EOF(1)
 
          Line Input #1, WReg ' L� a linha para a vari�vel.
    
          WSeq = Mid$(WReg, 1, 1)
          If WSeq = "0" Then
             BeginTrans
             'Cabecalho da venda
             tb.Seek "=", Val(Mid$(Result(Indice), 3, 9))
             If tb.NoMatch Then
                tb.AddNew
             Else
                tb.Edit
             End If
             tb!Controle = Val(Mid$(Result(Indice), 3, 9))
             tb!codigofunc = Val(Mid$(WReg, 2, 3))
             tb!CodigoCliente = Val(Mid$(WReg, 5, 9))
             tb!codigoprazo = Val(Mid$(WReg, 14, 2))
             tb!datavenda = DateValue(Mid$(WReg, 36, 10))
             tb!nroecf = Val(Mid$(WReg, 46, 4))
             tb!nrocoo = Val(Mid$(WReg, 50, 9))
             tb!NroCCF = Val(Mid$(Result(Indice), 3, 9))
             tb.Update
             CommitTrans
          ElseIf WSeq = "1" Then
                 'Itens vendidos
                 BeginTrans
                 tbi.index = "Primarykey"
                 WSeq_Reg = WSeq_Reg + 1
                 tbi.Seek "=", Val(Mid$(Result(Indice), 3, 9)), Val(Mid$(WReg, 5, 7)), WSeq_Reg
                 If tbi.NoMatch Then
                    tbi.AddNew
                 Else
                    tbi.Edit
                 End If
                  tbi!Controle = Val(Mid$(Result(Indice), 3, 9))
                  tbi!CodItem = Val(Mid$(WReg, 5, 7))
                  tbi!SeqReg = WSeq_Reg
                  tbi!Quantidade = Val(Mid$(WReg, 12, 7)) / 1000
                  tbi!ValorUnit = Val(Mid$(WReg, 26, 7)) / 100
                  tbi!Desconto = Val(Mid$(WReg, 59, 8)) / 100
                  tbi.Update
                  CommitTrans
          ElseIf WSeq = "9" Then
                 'Fim de arquivo
                 
'                 GoSub ImportaVendaSga:
                 Close #1
                 Kill Wcaminho & Result(Indice)
                 
 '                GoSub Inicializa_TB:
  '               GoSub Inicializa_TBI:
                 
                 'MsgBox "Fim de arquivo"
                 
                 Return
          End If
       
   
    Wend
 
Return
 
Inicializa_TBI:
    If tbi.RecordCount <> 0 Then
       BeginTrans
       While Not tbi.EOF
            tbi.Delete
            tbi.MoveNext
       Wend
       CommitTrans
    End If
Return
  
Inicializa_TB:
    If tb.RecordCount <> 0 Then
       BeginTrans
       While Not tb.EOF
            tb.Delete
            tb.MoveNext
       Wend
       CommitTrans
    End If
Return
   

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGv"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If
   
End Function

Function GeraCuponXSalc(WNumos As Long, wMensagem As String) As Long

On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb, tba, tbb, tbc, tbd, tbe, tbp As Recordset
Dim WSeq As Long
Dim WCodItem_Serv As String * 7
Dim WTotal_Desconto_Itens As Double
Dim WTotal_Desconto_Serv As Double
Dim WTotal_Desconto As Double
Dim WMensagem_Cupon As String * 48
Dim WFormaPagto As String * 2

Dim WSaidas As Double
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
Set tba = bd.OpenRecordset("RQM", DB_OPEN_TABLE)
Set tbb = bd.OpenRecordset("RMXEMPXEST", DB_OPEN_TABLE)
Set tbc = bd.OpenRecordset("EST", DB_OPEN_TABLE)
Set tbd = bd.OpenRecordset("ITENSSERVOSV", DB_OPEN_TABLE)
Set tbe = bd.OpenRecordset("TPSERV", DB_OPEN_TABLE)
Set tbp = bd.OpenRecordset("CONDPAGTO", DB_OPEN_TABLE)

tb.index = "PrimaryKey"
tba.index = "secondaryKey"
tbb.index = "RQM"
tbc.index = "PrimaryKey"
tbd.index = "secondaryKey"
tbe.index = "PrimaryKey"
tbp.index = "PrimaryKey"

WSeq = 0
Open CapturaCaminhoIntegracaoECF() & "\Req\P" & Format$(Val(RetornaNroSerieECF()), "0000") & "0001" & Format$(WNumos, "0000000000") & ".PRE" For Output As #2

BeginTrans
tb.Seek "=", WNumos
If Not tb.NoMatch Then

   WTotal_Desconto_Itens = TotalDescontoItensOs(WNumos)
   WTotal_Desconto_Serv = TotalDescontosServicosOs(WNumos)
   
   GoSub Apura_Forma_Pagto:
   
   GoSub GravaReg0:
   GoSub GravaReg1:
   If ECF_CONJUGADO() Then
      If TotalServicosOrdemServico(WNumos) <> 0 Then
         GoSub GravaReg1_Serv:
      End If
   End If
   GoSub GravaReg2:
End If

Finaliza:
   Print #2, "9"
   tb.Close
   tba.Close
   tbb.Close
   tbc.Close
   tbd.Close
   tbe.Close
   Close #2
   bd.Close
   CommitTrans
   Set tb = Nothing
   Set tba = Nothing
   Set tbb = Nothing
   Set tbc = Nothing
   Set tbd = Nothing
   Set tbe = Nothing
   Set bd = Nothing
   Exit Function
Return

Apura_Forma_Pagto:
   '"DIN";"DINHEIRO";"CHQ";"CHEQUE";"CAR";"CARTAO DE CREDITO";"FIN";"FINANCIAMENTO";"CHP";"CHEQUE PRE";"DEB";"CARTAO DE DEBITO";"OUT";"OUTROS"
   '
   '  ---> SITUACOES DO SALC 01=DINHEIRO 02=CARTAO DEB/CRE 03=         04=
   '
   '01=Dinheiro
   '02=Cheque
   '03=Cart�o de Cr�dito
   '04=Cart�o de D�bito
   '05=Cr�dito Loja
   '10=Vale Alimenta��o
   '11=Vale Refei��o
   '12=Vale Presente
   ' 13=Vale Combust�vel
   '99=Outros
   
    tbp.Seek "=", tb!CodigoPagto
    If Not tbp.NoMatch Then
       If IsNull(tbp!FormaPagto) Then
           WFormaPagto = "01"
       ElseIf tbp!FormaPagto = "FIN" Then
              WFormaPagto = "05"
       ElseIf tbp!FormaPagto = "CAR" Then
              WFormaPagto = "03"
       ElseIf tbp!FormaPagto = "DEB" Then
              WFormaPagto = "02"
       ElseIf tbp!FormaPagto = "OUT" Then
              WFormaPagto = "99"
       ElseIf tbp!FormaPagto = "PIX" Then
              WFormaPagto = "99"
       ElseIf tbp!FormaPagto = "TRA" Then
              WFormaPagto = "99"
       ElseIf tbp!FormaPagto = "DIN" Then
              WFormaPagto = "01"
       ElseIf tbp!FormaPagto = "CHQ" Or tbp!FormaPagto = "CHP" Then
            WFormaPagto = "01"
            
       Else
          WFormaPagto = "01"
       End If
     End If
Return

GravaReg0:

   If ECF_CONJUGADO() Then
      WTotal_Desconto = WTotal_Desconto_Itens + WTotal_Desconto_Serv
   Else
      WTotal_Desconto = WTotal_Desconto_Itens
   End If
   If wMensagem = "#" Then
      WMensagem_Cupon = Space(48)
   Else
      WMensagem_Cupon = UCase(Mid$(wMensagem, 1, 48))
   End If
   Print #2, "0" & Format(tb!CodigoCliente, "000000000") & Format$(tb!Numos, "0000000000") & "DAV" & Format$(tb!Numos, "0000000000") & Space(35) & "000000000" & "000000000" & Format$(tb!datafechamento, "ddmmyyyy") & "PEDIDO" & Space(24) & Format$(WMensagem_Cupon, "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") & "L" & Space(26) & "X"
Return

GravaReg1:
  tba.Seek "=", WNumos
  If Not tba.NoMatch Then
     While Not tba.EOF
           If WNumos <> tba!Numos Then
              Return
           End If
           GoSub GravaReg11:
           tba.MoveNext
     Wend
  End If
Return

GravaReg1_Serv:
  tbd.Seek "=", WNumos
  If Not tbd.NoMatch Then
     While Not tbd.EOF
           If tbd!Numos <> WNumos Then
              Return
           End If
           
           WSaidas = tbd!Quantidade
           
           If WSaidas <> 0 Then
              tbe.Seek "=", tbd!TipoServico
              WCodItem_Serv = "1" & Format$(tbd!TipoServico, "000000")
              If Not tbc.NoMatch Then
                 WSeq = WSeq + 1
                 Print #2, "1" & WCodItem_Serv & Format$(WSaidas * 1000, "000000000") & "0001000" & Format$(tbd!ValorMaoObra * 100, "0000000") & "C" & Format$(tb!datafechamento, "ddmmyyyy") & Format$((tbd!Desconto * WSaidas) * 100, "00000000") & "00000000" & Format$(WSeq, "000") & Space(10) & "X"
              End If
           End If
           tbd.MoveNext
     Wend
  End If
Return

GravaReg11:
  tbb.Seek "=", tba!NumRequisicao
  If Not tbb.NoMatch Then
     While Not tbb.EOF
           If tbb!NumRequisicao <> tba!NumRequisicao Then
              Return
           End If
           
           WSaidas = (tbb!Quantidade - tbb!Devolucao)
           
           If WSaidas <> 0 Then
              tbc.Seek "=", tbb!CodItem
              If Not tbc.NoMatch Then
                 WSeq = WSeq + 1
                 Print #2, "1" & Format$(tbb!CodItem, "0000000") & Format$(WSaidas * 1000, "000000000") & "0001000" & Format$(tbb!ValorUnit * 100, "0000000") & "C" & Format$(tb!datafechamento, "ddmmyyyy") & Format$((tbb!Desconto * WSaidas) * 100, "00000000") & "00000000" & Format$(WSeq, "000") & Space(10) & "X"
              End If
           End If
           tbb.MoveNext
     Wend
  End If
Return

GravaReg2:
  
  If ECF_CONJUGADO() Then
     Print #2, "2" & WFormaPagto & Format$((CalculoMercadorias(WNumos) + TotalServicosOrdemServico(WNumos) + WTotal_Desconto) * 100, "0000000000000") & Space(53) & "X"
  Else
     Print #2, "2" & WFormaPagto & Format$(CalculoMercadorias(WNumos) * 100, "0000000000000") & Space(53) & "X"
  End If
   
Return:

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
       Rollback
        
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
       Rollback
        
    End If

End Function
 
Function Gera_Todos_Produtos_Salc() As Long
 
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb, tba, tbb, tbc As Recordset
Dim WSaidas As Long
Dim wNomeArquivo As String * 256
' Campos auxiliares
Dim X As Long
Dim WSeq As Long
Dim Ret As Variant
Dim WDescricao As String * 40
Dim WUnidade As String * 2
Dim WReferencia As String * 15
Dim WPreco As String * 12
Dim WTrib As String * 3
Dim WCodBar As String * 13
Dim WCFOP As String * 4
Dim WCodigoNCM As String * 8
Dim wCodigoCEST As String * 7
Dim WTipoItem As String * 2
Dim WCST_ECF As String * 2

Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tbc = bd.OpenRecordset("EST", DB_OPEN_TABLE)
Set tbb = bd.OpenRecordset("EMPXEST", DB_OPEN_TABLE)
tbc.index = "PrimaryKey"
tbb.index = "PrimaryKey"

WSeq = 0
wNomeArquivo = CapturaCaminhoIntegracaoECF() & "\Req\ENV" & RetornaNroSerieECF() & ".PRO"
Open wNomeArquivo For Output As #1
BeginTrans

If tbc.RecordCount <> 0 Then
   While Not tbc.EOF
         If tbc!Ativo Then
            tbb.Seek "=", Form_Menu.WEmpresa, tbc!CodItem
            If Not tbb.NoMatch Then
               If Not IsNull(tbc!Descricao) Then
                  WDescricao = Mid$(UCase(ConverteCaracEsp_Mip(Trim(tbc!Descricao))), 1, 40)
               Else
                  WDescricao = Space(40)
               End If
               WUnidade = "UN"
              WReferencia = IIf(IsNull(tbc!CodigoInterno), Space(15), tbc!CodigoInterno)
               WPreco = Format$((tbb!PrecoVenda * 100), "000000000000")
               WTrib = Format$(Converte_CST_Simples_Nacional(tbc!CodigoCST), "000")
               WCodBar = "             "
               WCFOP = "5405"
               WCodigoNCM = IIf(Not IsNumeric(tbc!Codigoncm), "90299010", Format$(tbc!Codigoncm, "!@@@@@@@@"))
               wCodigoCEST = IIf(Not IsNumeric(CapturaCodigoCEST(WCodigoNCM)), "0108000", IIf(Not IsNumeric(CapturaCodigoCEST(WCodigoNCM)), "0108000", CapturaCodigoCEST(WCodigoNCM)))
               WTipoItem = "00"
               WCST_ECF = Retorna_CST_ECF_EST(tbc!CodigoCST)
               Print #1, "1" & Format$(tbc!CodItem, "0000000") & WDescricao & WUnidade & Space(3) & WReferencia & WPreco & WCST_ECF & WCodBar & "S" & "S" & "N" & "N" & "S" & "N" & "A" & "T" & WTipoItem & WCodigoNCM & Space(6) & "0000" & WTrib & Space(25) & WCFOP & wCodigoCEST & Space(37) & "X"
            End If
         End If
         tbc.MoveNext
         X = X + 1
         If X Mod 500 = 0 Then
           CommitTrans
           BeginTrans
         End If
         Ret = SysCmd(SYSCMD_SETSTATUS, "Lendo registro " & Str(X))
   Wend
End If

Finaliza:
   CommitTrans
    Ret = SysCmd(acSysCmdRemoveMeter)
   tbc.Close
   tbb.Close
   Close #1
   bd.Close
   Set tbc = Nothing
   Set tbb = Nothing
   Set bd = Nothing
   'MsgBox "Fim normal de execu��o", vbCritical, "SGa"
   Exit Function
Return

GravaReg2:
  Print #1, "2" & "0000000" & Space(191) & "X"
Return:

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
       Rollback
        
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       
       
       Resume Next
       Rollback
        
    End If
End Function

Function Gera_Produtos_OSV_Salc(WNumos As Long) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb, tba, tbb, tbc As Recordset
Dim WSaidas As Long
Dim wNomeArquivo As String * 256
' Campos auxiliares
Dim X As Long
Dim WSeq As Long
Dim Ret As Variant
Dim WDescricao As String * 40
Dim WUnidade As String * 2
Dim WReferencia As String * 15
Dim WPreco As String * 12
Dim WTrib As String * 3
Dim WCodBar As String * 13
Dim WCFOP As String * 4
Dim WCodigoNCM As String * 8
Dim wCodigoCEST As String * 7
Dim WTipoItem As String * 2
Dim WCST_ECF As String * 2

Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tbc = bd.OpenRecordset("EST", DB_OPEN_TABLE)
Set tbb = bd.OpenRecordset("EMPXEST", DB_OPEN_TABLE)
Set tb = bd.OpenRecordset("RQM", DB_OPEN_TABLE)
Set tba = bd.OpenRecordset("RMXEMPXEST", DB_OPEN_TABLE)
tbc.index = "PrimaryKey"
tbb.index = "PrimaryKey"
tb.index = "SecondaryKey"
tba.index = "RQM"

WSeq = 0

wNomeArquivo = CapturaCaminhoIntegracaoECF() & "\Req\ENV" & RetornaNroSerieECF() & ".PRO"
Open wNomeArquivo For Output As #1

BeginTrans
tb.Seek "=", WNumos
If Not tb.NoMatch Then
   GoSub Apura_Itens:
End If

Finaliza:
   CommitTrans
   tbc.Close
   tb.Close
   tba.Close
   tbb.Close
   Close #1
   bd.Close
   Set tbc = Nothing
   Set tb = Nothing
   Set tba = Nothing
   Set tbb = Nothing
   Set bd = Nothing
   Exit Function
Return

Apura_Itens:
  tba.Seek "=", tb!NumRequisicao
  If Not tba.NoMatch Then
     While tba!NumRequisicao = tb!NumRequisicao
           GoSub Grava_Produto:
            X = X + 1
            If X Mod 100 = 0 Then
               CommitTrans
               BeginTrans
            End If
            tba.MoveNext
            If tba.EOF Then
               Return
            End If
     Wend
  End If
Return

Grava_Produto:
   tbc.Seek "=", tba!CodItem
   If Not tbc.NoMatch Then
      If tbc!Ativo Then
            tbb.Seek "=", Form_Menu.WEmpresa, tbc!CodItem
            If Not tbb.NoMatch Then
               If Not IsNull(tbc!Descricao) Then
                  WDescricao = Mid$(UCase(ConverteCaracEsp_Mip(tbc!Descricao)), 1, 40)
               Else
                  WDescricao = Space(40)
               End If
               WUnidade = "UN"
               WReferencia = IIf(IsNull(tbc!CodigoInterno), Space(15), tbc!CodigoInterno)
               WPreco = Format$((tbb!PrecoVenda * 100), "000000000000")
               WTrib = Format$(Converte_CST_Simples_Nacional(tbc!CodigoCST), "000")
               WCodBar = "             "
               WCFOP = "5405"
               WCodigoNCM = IIf(Not IsNumeric(tbc!Codigoncm), "90299010", Format$(tbc!Codigoncm, "!@@@@@@@@"))
               wCodigoCEST = IIf(Not IsNumeric(CapturaCodigoCEST(WCodigoNCM)), "0108000", IIf(Not IsNumeric(CapturaCodigoCEST(WCodigoNCM)), "0108000", CapturaCodigoCEST(WCodigoNCM)))
               WTipoItem = "00"
               WCST_ECF = Retorna_CST_ECF_EST(tbc!CodigoCST)
               Print #1, "1" & Format$(tbc!CodItem, "0000000") & WDescricao & WUnidade & Space(3) & WReferencia & WPreco & WCST_ECF & WCodBar & "S" & "S" & "N" & "N" & "S" & "N" & "A" & "T" & WTipoItem & WCodigoNCM & Space(6) & "0000" & WTrib & Space(25) & WCFOP & wCodigoCEST & Space(37) & "X"
            End If
         End If
   End If
Return



GravaReg2:
  Print #1, "2" & "0000000" & Space(191) & "X"
Return:

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
       Rollback
        
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       
       
       Resume Next
       Rollback
        
    End If

End Function


Function Gera_Itens_Produtos_Salc(WNumos As Long) As Long
 
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb, tba, tbb, tbc As Recordset
Dim WSaidas As Long
Dim WSeq As Long
Dim wNomeArquivo As String * 256
' Campos auxiliares

Dim WDescricao As String * 40
Dim WUnidade As String * 2
Dim WReferencia As String * 15
Dim WPreco As String * 12
Dim WTrib As String * 3
Dim WCodBar As String * 13
Dim WCFOP As String * 4
Dim WCodigoNCM As String * 8
Dim wCodigoCEST As String * 7
Dim WTipoItem As String * 2
Dim WCST_ECF As String * 2

Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
Set tba = bd.OpenRecordset("RQM", DB_OPEN_TABLE)
Set tbb = bd.OpenRecordset("RMXEMPXEST", DB_OPEN_TABLE)
Set tbc = bd.OpenRecordset("EST", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tba.index = "secondaryKey"
tbb.index = "RQM"
tbc.index = "PrimaryKey"

WSeq = 0

wNomeArquivo = CapturaCaminhoIntegracaoECF() & "\Req\ENV" & RetornaNroSerieECF() & ".PRO"

Open wNomeArquivo For Output As #1

BeginTrans

tb.Seek "=", WNumos
If Not tb.NoMatch Then
   GoSub GravaReg1:
   'GoSub GravaReg2:
End If

Finaliza:
   CommitTrans
   If ECF_CONJUGADO() Then
      If TotalServicosOrdemServico(WNumos) = 0 Then
         Print #1, "9"
      End If
   End If
   tb.Close
   tba.Close
   tbb.Close
   tbc.Close
   Close #1
   bd.Close
   Set tb = Nothing
   Set tba = Nothing
   Set tbb = Nothing
   Set tbc = Nothing
   Set bd = Nothing
   Exit Function
Return

GravaReg1:
  tba.Seek "=", WNumos
  If Not tba.NoMatch Then
     While Not tba.EOF
           If WNumos <> tba!Numos Then
              Return
           End If
           GoSub GravaReg11:
           tba.MoveNext
     Wend
  End If
Return

GravaReg11:
  tbb.Seek "=", tba!NumRequisicao
  If Not tbb.NoMatch Then
     While Not tbb.EOF
           If tbb!NumRequisicao <> tba!NumRequisicao Then
              Return
           End If
           
           WSaidas = (tbb!Quantidade - tbb!Devolucao)
           
           If WSaidas <> 0 Then
              tbc.Seek "=", tbb!CodItem
              If Not tbc.NoMatch Then
                 WDescricao = IIf(IsNull(tbc!Descricao), Space(40), Mid$(UCase(ConverteCaracEsp_Mip(tbc!Descricao)), 1, 40))
                 WUnidade = "UN"
                 WReferencia = IIf(IsNull(tbc!CodigoInterno), Space(15), tbc!CodigoInterno)
                 WPreco = Format$((tbb!ValorUnit * 100), "000000000000")
                 WTrib = Format$(Converte_CST_Simples_Nacional(tbc!CodigoCST), "000")
                 WCodBar = "             "
                 WCFOP = "5405"
                 WCodigoNCM = IIf(Not IsNumeric(tbc!Codigoncm), "90299010", Format$(tbc!Codigoncm, "!@@@@@@@@"))
                 wCodigoCEST = IIf(Not IsNumeric(tbc!Codigoncm), "0108000", IIf(Not IsNumeric(CapturaCodigoCEST(WCodigoNCM)), "0108000", CapturaCodigoCEST(WCodigoNCM)))
                 WTipoItem = "00"
                 WCST_ECF = Retorna_CST_ECF_EST(tbc!CodigoCST)
                 Print #1, "1" & Format$(tbb!CodItem, "0000000") & WDescricao & WUnidade & Space(3) & WReferencia & WPreco & WCST_ECF & WCodBar & "S" & "S" & "N" & "N" & "S" & "N" & "A" & "T" & WTipoItem & WCodigoNCM & Space(6) & "0000" & WTrib & Space(25) & WCFOP & wCodigoCEST & Space(37) & "X"
              End If
           End If
           tbb.MoveNext
     Wend
  End If
Return

GravaReg2:
  Print #1, "2" & "0000000" & Space(191) & "X"
Return:

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
       Rollback
        
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
       Rollback
        
    End If


 End Function
 
Function GeraEstoqueXSalc() As Long

On Error GoTo TratamentodeErro:
DoCmd.Hourglass True
Dim bd As DataBase
Dim tb As Recordset
Dim WSeq As Long
Dim wNomeArquivo As String * 256
Dim X As Long
Dim Ret As Variant

' Campos auxiliares
Dim wsaldo As Double

Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("EST", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

WSeq = 0

wNomeArquivo = CapturaCaminhoBD() & "\Req\ENV" & RetornaNroSerieECF() & ".EST"

Open wNomeArquivo For Output As #1

BeginTrans

If tb.RecordCount <> 0 Then
   While Not tb.EOF
         If tb!Ativo Then
            GoSub GravaReg11:
         End If
         
         X = X + 1
         If X Mod 500 = 0 Then
           CommitTrans
           BeginTrans
         End If
         Ret = SysCmd(SYSCMD_SETSTATUS, "Lendo " & (UCase(CapturaDescricaoItemEstoque(tb!CodItem))))
         
         tb.MoveNext
            
   Wend
End If

Finaliza:
   DoCmd.Hourglass False
   CommitTrans
    Ret = SysCmd(acSysCmdRemoveMeter)
   Print #1, "9"
   tb.Close
   Close #1
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
   Exit Function
Return

GravaReg11:
   wsaldo = (SaldoContabilEstoque(tb!CodItem, Date, Empresa_Ativa())) * 1000
   If wsaldo < 0 Then
      wsaldo = 0
   End If
   
   Print #1, "1" & Format$(tb!CodItem, "0000000") & Format$(wsaldo, "00000000000000") & Space(177) & "X"
Return

GravaReg2:
  Print #1, "2" & "0000000" & Space(191) & "X"
Return:

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGv"
       Rollback
        
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
       Rollback
        
    End If

End Function

Function PadronizaAroPneu(WDescricao As String, WValue As String) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("AROPNEU", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
BeginTrans

tb.Seek "=", WDescricao
If tb.NoMatch Then
   tb.AddNew
Else
   tb.Edit
End If
tb!DescricaoAroPneu = WDescricao
tb!Value = WValue
tb.Update

Finaliza:
  CommitTrans
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function PadronizaMarcaVeiculo(WDescricao As String, WValue As String) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("MARC", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
BeginTrans

tb.Seek "=", WDescricao
If tb.NoMatch Then
   tb.AddNew
Else
   tb.Edit
End If
tb!DescricaoMarca = WDescricao
tb!Value = WValue
tb.Update

Finaliza:
  CommitTrans
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function PadronizaMarcaEquip(WDescricao As String) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("MarcaEquip", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
BeginTrans
tb.Seek "=", WDescricao
If tb.NoMatch Then
   tb.AddNew
Else
   tb.Edit
End If
tb!Descricao = WDescricao
tb!ABREVIATURA = WDescricao
tb.Update

Finaliza:
  CommitTrans
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Function LimpaRegistrosPneus() As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("PNEU", DB_OPEN_TABLE)
BeginTrans
Do While Not tb.EOF
    If Len(tb("DescricaoPneu").Value) > 6 Then
        tb.Delete
    End If
    tb.MoveNext
Loop
CommitTrans

Finaliza:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function


Function Gera_Itens_Servicos_Salc(WNumos As Long) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb, tba, tbb As Recordset
Dim WSaidas As Long
Dim wNomeArquivo As String * 256
' Campos auxiliares

Dim WCodItem As String * 7
Dim WDescricao As String * 40
Dim WUnidade As String * 2
Dim WReferencia As String * 15
Dim WPreco As String * 12
Dim WTrib As String * 3
Dim WCodBar As String * 13
Dim WCFOP As String * 4
Dim WCodigoNCM As String * 9
Dim WTipoItem As String * 2
Dim WSeq As Long
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
Set tba = bd.OpenRecordset("ITENSSERVOSV", DB_OPEN_TABLE)
Set tbb = bd.OpenRecordset("TPSERV", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tba.index = "secondaryKey"
tbb.index = "PrimaryKey"

WSeq = 0

wNomeArquivo = CapturaCaminhoIntegracaoECF() & "\Req\ENV" & RetornaNroSerieECF() & ".PRO"

Open wNomeArquivo For Append As #1

BeginTrans

tb.Seek "=", WNumos
If Not tb.NoMatch Then
   GoSub GravaReg1:
   'GoSub GravaReg2:
End If

Finaliza:
   CommitTrans
   Print #1, "9"
   tb.Close
   tba.Close
   tbb.Close
   Close #1
   bd.Close
   Set tb = Nothing
   Set tba = Nothing
   Set tbb = Nothing
   Set bd = Nothing
   Exit Function
Return

GravaReg1:
  tba.Seek "=", WNumos
  If Not tba.NoMatch Then
     While Not tba.EOF
           If tba!Numos <> WNumos Then
              Return
           End If
           
           WSaidas = tba!Quantidade
           
           If WSaidas <> 0 Then
              tbb.Seek "=", tba!TipoServico
              If Not tbb.NoMatch Then
               
                 WCodItem = "1" & Format$(tba!TipoServico, "000000")
                 WDescricao = IIf(IsNull(tbb!Descricao), Space(40), Mid$(UCase(ConverteCaracEsp_Mip(tbb!Descricao)), 1, 40))
                 WUnidade = "UN"
                 WReferencia = Space(15)
                 WPreco = Format$((tba!ValorMaoObra * 100), "000000000000")
                 WTrib = "900"
                 WCodBar = "             "
                 WCFOP = "5932"
                 WCodigoNCM = "120013100"
                 WTipoItem = "09"
                 Print #1, "1" & WCodItem & WDescricao & WUnidade & Space(3) & WReferencia & WPreco & "IS" & WCodBar & "S" & "S" & "N" & "N" & "S" & "N" & "A" & "T" & WTipoItem & WCodigoNCM & Space(3) & "00" & "1402" & WTrib & Space(25) & WCFOP & Space(44) & "X"
                 
              End If
           End If
           tba.MoveNext
     Wend
  End If
Return

GravaReg2:
  Print #1, "2" & "0000000" & Space(191) & "X"
Return:

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
       Rollback
        
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
       Rollback
        
    End If


End Function
 
Function GeraResponsaveisXSalc(WNumos As Long) As Long

On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb, tba, tbb As Recordset
Dim WSeq As Long
Dim WSaidas As Long

' Campos Auxiliares

Dim WNomeFunc As String * 40
Dim WEndereco As String * 40
Dim WBairro As String * 20
Dim wcep As String * 10
Dim WCidade As String * 30
Dim WUF As String * 2
Dim wTelefone As String * 14
Dim WFax As String * 14
Dim WCNPJ As String * 14
Dim wInsc As String * 15
Dim WFormaPagto As String * 2
Dim wcodigoibge As String * 7

Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
Set tba = bd.OpenRecordset("RESP", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tba.index = "PrimaryKey"

Open CapturaCaminhoIntegracaoECF() & "\Req\ENV" & RetornaNroSerieECF() & ".FUN" For Output As #1


BeginTrans

tb.Seek "=", WNumos
If Not tb.NoMatch Then
   GoSub GravaReg1:
   'GoSub GravaReg2:
End If

Finaliza:
   CommitTrans
   Print #1, "9"
   tb.Close
   tba.Close
   Close #1
   bd.Close
   Set tb = Nothing
   Set tba = Nothing
   Set bd = Nothing
   Exit Function
Return

GravaReg1:
  tba.Seek "=", tb!Responsavel
  If Not tba.NoMatch Then
     WNomeFunc = IIf(IsNull(tba!Nome), Space(40), Format$(UCase(ConverteCaracEsp_Mip(tba!Nome)), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"))
     WEndereco = IIf(IsNull(Endereco()), Space(40), Format$(UCase(ConverteCaracEsp_Mip(Endereco())), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"))
     WBairro = IIf(IsNull(Bairro()), Space(20), Format$(UCase(ConverteCaracEsp_Mip(Bairro())), "!@@@@@@@@@@@@@@@@@@@@"))
     wcep = IIf(IsNull(Cep()), Space(10), Format$(UCase(Cep()), "!@@@@@@@@@@"))
     WCidade = IIf(IsNull(Cidade()), Space(30), Format$(UCase(Mid$(ConverteCaracEsp_Mip(Cidade()), 1, 28)) & Estado(), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"))
     wTelefone = Space(14)
     
     Print #1, "1" & Format$(tb!Responsavel, "000") & WNomeFunc & "VENDEDOR            " & WEndereco & WBairro & wcep & WCidade & wTelefone & Space(8) & "S" & "3" & Space(5) & Space(20) & "S" & "X"
     
  End If
Return

GravaReg2:
  Print #1, "2" & "000" & Space(210) & "X"
Return:

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
       Rollback
        
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
       Rollback
        
    End If

End Function
 
Function GeraTributXSalc(WSP As String) As Long

On Error GoTo TratamentodeErro:


' Campos Auxiliares


Dim WCodTrib As String * 2
Dim WAliqTrib As String * 5
Dim WTPTrib As String * 2
Dim WPercIPI As String * 5
Dim WCST As String * 3
Dim WBrancos As String * 2
Dim WFimReg As String * 1

Open CapturaCaminhoIntegracaoECF() & "\Req\ENV" & RetornaNroSerieECF() & ".TRI" For Output As #1


If WSP = "P" Then
   GoSub GravaReg1P:
ElseIf WSP = "S" Then
      GoSub GravaReg1S:
Else
    GoSub GravaReg1PS:
End If

Finaliza:
   Print #1, "9"
   Close #1
   Exit Function
Return

GravaReg1P:
   WCodTrib = "FF"
   WAliqTrib = "00000"
   WTPTrib = "N"
   WPercIPI = "00000"
   WCST = "500"
   WBrancos = Space(2)
   WFimReg = "X"
   Print #1, "1" & WCodTrib & WAliqTrib & WTPTrib & WPercIPI & WCST & WBrancos & WFimReg
Return

GravaReg1S:
   WCodTrib = "IS"
   WAliqTrib = "00000"
   WTPTrib = "S"
   WPercIPI = "00000"
   WCST = "900"
   WBrancos = Space(2)
   WFimReg = "X"
   Print #1, "1" & WCodTrib & WAliqTrib & WTPTrib & WPercIPI & WCST & WBrancos & WFimReg
Return

GravaReg1PS:

   ' Grava Reg trib pe�as
   WCodTrib = "FF"
   WAliqTrib = "00000"
   WTPTrib = "N"
   WPercIPI = "00000"
   WCST = "060"
   WBrancos = Space(2)
   WFimReg = "X"
   Print #1, "1" & WCodTrib & WAliqTrib & WTPTrib & WPercIPI & WCST & WBrancos & WFimReg
   
   ' Grava Reg trib serv
   WCodTrib = "IS"
   WAliqTrib = "00000"
   WTPTrib = "S"
   WPercIPI = "00000"
   WCST = "090"
   WBrancos = Space(2)
   WFimReg = "X"
   Print #1, "1" & WCodTrib & WAliqTrib & WTPTrib & WPercIPI & WCST & WBrancos & WFimReg
Return

GravaReg2:
  Print #1, "2" & "00" & Space(16) & "X"
Return:

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
       Rollback
        
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
       Rollback
        
    End If

End Function
 
Function GeraFormasPagtoXSalc(WNumos As Long) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb, tba, tbb As Recordset
Dim WSeq As Long
Dim WSaidas As Long

' Campos Auxiliares

Dim WDescricao As String * 30
Dim WCodinome As String * 20
Dim WFormaPagto As String * 2
Dim WParcelas As String * 2
Dim WCoefic As String * 5
Dim WNroDias As String * 1
Dim WProxVenc As String * 2

Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
Set tba = bd.OpenRecordset("CONDPAGTO", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tba.index = "PrimaryKey"

Open CapturaCaminhoIntegracaoECF() & "\Req\ENV" & RetornaNroSerieECF() & ".FOR" For Output As #1

BeginTrans
tb.Seek "=", WNumos
If Not tb.NoMatch Then
   WParcelas = "00"
   WCoefic = "10000"
   GoSub GravaReg1:
   'GoSub GravaReg2:
End If

Finaliza:

   Print #1, "9"
   tb.Close
   tba.Close
   Close #1
   bd.Close
   Set tb = Nothing
   Set tba = Nothing
   Set bd = Nothing
   CommitTrans
   Exit Function
Return

GravaReg1:
 tba.Seek "=", tb!CodigoPagto
 If Not tba.NoMatch Then
     
     If IsNull(tba!FormaPagto) Then
         WFormaPagto = "01"
         WDescricao = Format$("AVISTA", "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
         WCodinome = Format$("AVISTA", "!@@@@@@@@@@@@@@@@@@@@")
         WNroDias = "0"
         WProxVenc = "00"
     ElseIf tba!FormaPagto = "CAR" Then
         WFormaPagto = "02"
         WDescricao = Format$("CARTAO", "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
         WCodinome = Format$("CARTAO", "!@@@@@@@@@@@@@@@@@@@@")
         WParcelas = "01"
         WNroDias = "0"
         WProxVenc = "00"
     ElseIf tba!FormaPagto = "DEB" Then
         WFormaPagto = "02"
         WDescricao = Format$("CARTAO", "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
         WCodinome = Format$("CARTAO", "!@@@@@@@@@@@@@@@@@@@@")
         WParcelas = "01"
         WNroDias = "0"
         WProxVenc = "00"
     ElseIf tba!FormaPagto = "OUT" Then
         WFormaPagto = "02"
         WDescricao = Format$(tba!Descricao, "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
         WCodinome = Format$(tba!Descricao, "!@@@@@@@@@@@@@@@@@@@@")
         WParcelas = "01"
         WNroDias = "0"
         WProxVenc = "00"
     Else
         WFormaPagto = "01"
         WDescricao = Format$("AVISTA", "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
         WCodinome = Format$("AVISTA", "!@@@@@@@@@@@@@@@@@@@@")
         WNroDias = "0"
         WProxVenc = "00"
     End If
     
     Print #1, "1" & WFormaPagto & WDescricao & WCodinome & "10000" & WParcelas & "0000000000" & "00" & WNroDias & WProxVenc & "0000" & "10000" & Space(9) & "N" & "N" & "N" & "N" & "N" & "N" & Space(12) & "0" & "N" & Space(9) & "N" & Space(76) & "X"
     

 End If
Return

GravaReg2:
  Print #1, "2" & "00" & Space(196) & "X"
Return:

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
       Rollback
        
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
       Rollback
        
    End If
End Function
 
Function Gera_Todas_Formas_Pagto_Salc() As Long

If MsgBox("Dados est�o corretos ? ", 4 + 32, "ATEN��O") <> 6 Then
   Exit Function
End If
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb, tba, tbb As Recordset
Dim WSeq As Long
Dim WSaidas As Long

' Campos Auxiliares

Dim WDescricao As String * 30
Dim WCodinome As String * 20
Dim WFormaPagto As String * 2
Dim WParcelas As String * 2
Dim WCoefic As String * 5
Dim WNroDias As String * 1
Dim WProxVenc As String * 2

Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
Set tba = bd.OpenRecordset("CONDPAGTO", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tba.index = "PrimaryKey"

Open CapturaCaminhoIntegracaoECF() & "\Req\ENV" & RetornaNroSerieECF() & ".FOR" For Output As #1

BeginTrans
If tba.RecordCount <> 0 Then
   While Not tba.EOF
      If tba!FormaPagto <> "CAR" And tba!FormaPagto <> "DEB" Then
         WParcelas = "00"
         WCoefic = "10000"
         GoSub GravaReg1:
      End If
         tba.MoveNext
   Wend
End If


Finaliza:

   Print #1, "9"
   tb.Close
   tba.Close
   Close #1
   bd.Close
   CommitTrans
   Set tb = Nothing
   Set tba = Nothing
   Set bd = Nothing
   MsgBox "Fim de execu��o", vbCritical, "SGa"
   Exit Function
Return

GravaReg1:
     If IsNull(tba!FormaPagto) Then
         WFormaPagto = "01"
         WDescricao = Format$("AVISTA", "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
         WCodinome = Format$("AVISTA", "!@@@@@@@@@@@@@@@@@@@@")
         WNroDias = "0"
         WProxVenc = "00"
     ElseIf tba!FormaPagto = "CAR" Then
         WFormaPagto = "02"
         WDescricao = Format$("CARTAO", "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
         WCodinome = Format$("CARTAO", "!@@@@@@@@@@@@@@@@@@@@")
         WParcelas = "01"
         WNroDias = "0"
         WProxVenc = "00"
     ElseIf tba!FormaPagto = "DEB" Then
         WFormaPagto = "02"
         WDescricao = Format$("CARTAO", "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
         WCodinome = Format$("CARTAO", "!@@@@@@@@@@@@@@@@@@@@")
         WParcelas = "01"
         WNroDias = "0"
         WProxVenc = "00"
     ElseIf tba!FormaPagto = "OUT" Then
         WFormaPagto = "02"
         WDescricao = Format$(tba!Descricao, "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
         WCodinome = Format$(tba!Descricao, "!@@@@@@@@@@@@@@@@@@@@")
         WParcelas = "01"
         WNroDias = "0"
         WProxVenc = "00"
     Else
         WFormaPagto = "01"
         WDescricao = Format$("AVISTA", "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
         WCodinome = Format$("AVISTA", "!@@@@@@@@@@@@@@@@@@@@")
         WNroDias = "0"
         WProxVenc = "00"
     End If
     
     Print #1, "1" & WFormaPagto & WDescricao & WCodinome & "10000" & WParcelas & "0000000000" & "00" & WNroDias & WProxVenc & "0000" & "10000" & Space(9) & "N" & "N" & "N" & "N" & "N" & "N" & Space(12) & "0" & "N" & Space(9) & "N" & Space(76) & "X"
     
Return

GravaReg2:
  Print #1, "2" & "00" & Space(196) & "X"
Return:

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
       Rollback
        
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
       Rollback
        
    End If


End Function
 
Function GeraClientesXSalc(WNumos As Long, WCNPJCPF As String, WRazaoSocial_ECF As String) As Long

On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb, tba, tbb, tbc As Recordset
Dim WSeq As Long
Dim WSaidas As Long

Dim wNomeArquivo As String * 256

' Campos Auxiliares

Dim WNomeFantasia As String * 40
Dim WRazaoSocial As String * 40
Dim WEndereco As String * 40
Dim WBairro As String * 20
Dim wcep As String * 10
Dim WCidade As String * 30
Dim WUF As String * 2
Dim wTelefone As String * 14
Dim WFax As String * 14
Dim WCNPJ As String * 14
Dim wInsc As String * 15
Dim WFormaPagto As String * 2
Dim wcodigoibge As String * 7
Dim WNroEnd As String * 10

Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OS", DB_OPEN_TABLE)
Set tba = bd.OpenRecordset("CLI", DB_OPEN_TABLE)
Set tbb = bd.OpenRecordset("CID", DB_OPEN_TABLE)
Set tbc = bd.OpenRecordset("CONDPAGTO", DB_OPEN_TABLE)
tb.index = "PrimaryKey"
tba.index = "PrimaryKey"
tbb.index = "PrimaryKey"
tbc.index = "PrimaryKey"

wNomeArquivo = CapturaCaminhoIntegracaoECF() & "\Req\ENV" & RetornaNroSerieECF() & ".CLI"

Open wNomeArquivo For Output As #1
BeginTrans
tb.Seek "=", WNumos
If Not tb.NoMatch Then
   GoSub GravaReg1:
   'GoSub GravaReg2:
End If

Finaliza:

   Print #1, "9"
   tb.Close
   tba.Close
   Close #1
   bd.Close
   CommitTrans
   Set tb = Nothing
   Set tba = Nothing
   Set bd = Nothing
   Exit Function
Return

GravaReg1:
  tb.Seek "=", WNumos
  If Not tb.NoMatch Then
     tba.Seek "=", tb!CodigoCliente
     If Not tba.NoMatch Then
        'WNomeFantasia = IIf(IsNull(tba!Codinome), Space(40), Format$(UCase(ConverteCaracEsp(tba!Codinome)), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"))
        'WRazaoSocial = IIf(IsNull(tba!RazaoSocial), Space(40), Format$(UCase(ConverteCaracEsp(tba!RazaoSocial)), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"))
        WNomeFantasia = Format$(UCase(ConverteCaracEsp_Mip(Mid$(WRazaoSocial_ECF, 1, 40))), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
        WRazaoSocial = Format$(UCase(ConverteCaracEsp_Mip(Mid$(WRazaoSocial_ECF, 1, 40))), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
        WEndereco = IIf(IsNull(tba!Endereco) Or tba!Endereco = "", Format$(Endereco, "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"), IIf(IsNull(tba!Endereco), Space(40), Format$(UCase(ConverteCaracEsp_Mip(tba!Endereco)), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")))
        WBairro = IIf(IsNull(tba!Bairro) Or tba!Bairro = "", Format$(UCase(ConverteCaracEsp_Mip(Bairro)), "!@@@@@@@@@@@@@@@@@@@@"), IIf(IsNull(tba!Bairro), Space(20), Format$(UCase(ConverteCaracEsp_Mip(tba!Bairro)), "!@@@@@@@@@@@@@@@@@@@@")))
        wcep = IIf(Not IsNumeric(tba!EndPostal), Format$(UCase(Cep), "!@@@@@@@@@@"), IIf(IsNull(tba!EndPostal), Space(10), Format$(UCase(tba!EndPostal), "!@@@@@@@@@@")))
        WCidade = IIf(IsNull(tba!Cidade), Format$(UCase(ConverteCaracEsp_Mip(Cidade)), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"), Format$(UCase(ConverteCaracEsp_Mip(tba!Cidade)), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"))
        WUF = IIf(IsNull(tba!Estado), Format$(UCase(Estado), "!@@"), Format$(UCase(tba!Estado), "!@@"))
        'Wtelefone = IIf(Len(Tba!Telefone) < 8, Format$(UCase(Telefone), "!@@@@@@@@@@@@@@"), IIf(IsNull(Tba!Telefone), Space(14), Format$(UCase(Tba!Telefone), "!@@@@@@@@@@@@@@")))
         wTelefone = "3132012308"
         WFax = "3132012308"
        'WFax = IIf(IsNull(tba!Fax), Space(14), Format$(UCase(tba!Fax), "!@@@@@@@@@@@@@@"))
        WNroEnd = IIf(Not IsNumeric(tba!Nro), "01", tba!Nro)
        'WCNPJ = IIf(IsNull(tba!CgcCpf), Space(14), Format$(UCase(tba!CgcCpf), "!@@@@@@@@@@@@@@"))
        WCNPJ = Format$(WCNPJCPF, "!@@@@@@@@@@@@@@")
        wInsc = IIf(IsNull(tba!Inscricao), Space(15), Format$(UCase(tba!Inscricao), "!@@@@@@@@@@@@@@@"))
        
        tbc.Seek "=", tb!CodigoPagto
        If tbc.NoMatch Then
           WFormaPagto = "01"
        Else
           If IsNull(tbc!FormaPagto) Then
              WFormaPagto = "01"
           ElseIf tbc!FormaPagto = "CAR" Or tbc!FormaPagto = "DEB" Then
                  WFormaPagto = "02"
           Else
               WFormaPagto = "01"
           End If
        End If
        
        tbb.Seek "=", tba!Cidade
        wcodigoibge = IIf(IsNull(tbb!CodigoIBGE), CodigoMunicipio(), Format$(UCase(tbb!CodigoIBGE), "0000000"))
        Print #1, "1" & Format$(tba!CodigoCliente, "000000000") & WNomeFantasia & WRazaoSocial & WEndereco & WBairro & wcep & WCidade & WUF & wTelefone & WFax & WCNPJ & wInsc & WFormaPagto & "N" & wcodigoibge & "000000000" & "01058" & Format$(Date, "ddmmyyyy") & Format$(WNroEnd, "!@@@@@@@@@@") & Space(60) & Space(14) & "N" & Space(33) & "X"
        'Print #1, "1" & Format$(tba!CodigoCliente, "000000000") & WNomeFantasia & wRazaoSocial & WEndereco & wBairro & WCep & WCidade & WUF & wTelefone & WFax & WCNPJ & WInsc & "  N" & wcodigoibge & "000000000" & "01058" & Format$(Date, "ddmmyyyy") & Format$(WNroEnd, "!@@@@@@@@@@") & Space(60) & Space(14) & "N" & Space(33) & "X"
     End If
  End If
Return


GravaReg2:
  
  Print #1, "2" & "000000000" & Space(389) & "X"
   
Return:

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
       Rollback
        
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
       Rollback
        
    End If
 
End Function


Function Captura_Gera_Financeiro_OCORR_SERV(WCodigo As Long) As Boolean
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OCORR_SERV", DB_OPEN_TABLE)
tb.index = "PrimaryKey"

tb.Seek "=", WCodigo
If tb.NoMatch Then
   Captura_Gera_Financeiro_OCORR_SERV = False
Else
   Captura_Gera_Financeiro_OCORR_SERV = tb!Gera_Financeiro
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function


Function RetornaValueModeloVeiculo(WMarcaVeiculo As String, wModeloVeiculo As String) As Long
On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("MODELOVEIC", DB_OPEN_TABLE)
tb.index = "ChavePrimaria"
tb.Seek "=", WMarcaVeiculo, wModeloVeiculo
If Not tb.NoMatch Then
   If Not IsNull(tb!Value) Then
      RetornaValueModeloVeiculo = tb!Value
   Else
      RetornaValueModeloVeiculo = 0
   End If
End If

Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informa��es eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de n�mero " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function


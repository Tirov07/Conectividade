Attribute VB_Name = "Integracao_SGA_Inmetro_CRLV"
Option Compare Database
Option Explicit
Private Status As Boolean
Private ofuncoes As CanadaConnectivity.FuncoesDiversas
Private oconnreq As CanadaConnectivity.ReqGerarArquivoCRLV
Private oConnResp As CanadaConnectivity.LoadRespGerarArquivoCRLV
Public Function Integracao_SGA_Inmetro_Gerar_CRLV(ByVal WNumos As Long) As Boolean

Set oconnreq = New CanadaConnectivity.ReqGerarArquivoCRLV
Set oconnreq.oCRLV = New CanadaConnectivity.CRLV
Set ofuncoes = New CanadaConnectivity.FuncoesDiversas

If VerificaExisteArquivo("C:\SGA\CanainfoPrint.exe") = False Then
    MsgBox "O APLICATIVO NECESSARIO PARA REALIZAR ESSA OPERAÇÃO NÃO SE ENCONTRA CONFIGURADO. ENTRE EM CONTATO COM SUPORTE DA CANADÁ", vbCritical
    Integracao_SGA_Inmetro_Gerar_CRLV = False
    Set ofuncoes = Nothing
    Exit Function

End If

ofuncoes.ExcluiArquivos "C:\SGA\REQ\" & ofuncoes.ReturnUserWindows & ".xml"
ofuncoes.ExcluiArquivos "C:\SGA\RESP\" & ofuncoes.ReturnUserWindows & ".xml"

With oconnreq
    .oCRLV.Linha1 = UCase(Empresa)
    .oCRLV.Linha2 = UCase(Endereco & " - " & NroEnd & " - " & " - " & Bairro & " - " & Cidade & " - " & Estado)
    .oCRLV.Linha3 = "CNPJ: " & MascaraCgcCpf(CNPJ(), 2) & Ret_Pipe() & "INSC.EST.: " & InscricaoEstadual() & Ret_Pipe() & " INSC.MUN.: " & InscMunic()
    .oCRLV.Linha4 = Site() & Ret_Pipe() & Email() & Ret_Pipe() & Telefone()
    .oCRLV.Logomarca = LogomarcaEmpresa
    .oCRLV.NomeArquivo = ofuncoes.GeraNomeArquivoCRLV(WNumos, Now)
    .oCRLV.NomeArquivoCompleto = ofuncoes.GeraNomeArquivoCRLVCompleto(.oCRLV.NomeArquivo, CapturaCaminhoBD)
    LoadImagens WNumos
    If .oCRLV.CountImagens > 0 Then
       .oCRLV.ValidateImagem
       .CreateGerarArquivoCRLV 14, "www.canainfo.com.br"
    Else
    
       MsgBox "Nenhuma Imagem Vinculada, Impossivel Gerar o Arquivo.", vbCritical
       ofuncoes.ExcluiArquivos "C:\SGA\REQ\" & ofuncoes.ReturnUserWindows & ".xml"
       ofuncoes.ExcluiArquivos "C:\SGA\RESP\" & ofuncoes.ReturnUserWindows & ".xml"
       Integracao_SGA_Inmetro_Gerar_CRLV = False
       
       Set ofuncoes = Nothing
       Set oconnreq = Nothing
       Set oConnResp = Nothing
       
    End If
End With
'sleep (2000)

Shell "C:\SGA\CanainfoPrint.exe"

'AppActivate ("Conectividade - " & RetornaLoginInMetro)
Do While ofuncoes.CheckFile("C:\SGA\RESP\" & ofuncoes.ReturnUserWindows & ".xml") = False
Loop
'sleep (2000)
Set oConnResp = New CanadaConnectivity.LoadRespGerarArquivoCRLV
oConnResp.Load_Objetos ofuncoes.ReturnUserWindows
AppActivate "SGa"

If CBool(oConnResp.Status) = True Then
    Call AtualizaArquivoCRLV(WNumos, oconnreq.oCRLV.NomeArquivo)
    MsgBox oConnResp.Message, vbInformation, "SGA"
Else
    MsgBox oConnResp.Message, vbCritical, "SGA"
End If
ofuncoes.ExcluiArquivos "C:\SGA\REQ\" & ofuncoes.ReturnUserWindows & ".xml"
ofuncoes.ExcluiArquivos "C:\SGA\RESP\" & ofuncoes.ReturnUserWindows & ".xml"
Integracao_SGA_Inmetro_Gerar_CRLV = oConnResp.Status
Set ofuncoes = Nothing
Set oconnreq = Nothing
Set oConnResp = Nothing
End Function
Public Function Integracao_SGA_Inmetro_Informar_CRLV(ByVal WNumos As Long) As Boolean

Dim oconnreq As CanadaConnectivity.ReqInformarArquivoCRLV
Dim oConnResp As CanadaConnectivity.LoadRespGerarArquivoCRLV

Set ofuncoes = New CanadaConnectivity.FuncoesDiversas

'[CONECTIVIDADE_NOVA] validacao da Conectividade antiga desativada:
'If ofuncoes.ConectividadeIsAlive("Conectividade") = False Then
    'MsgBox "CONECTIVIDADE NAO ESTA ATIVA, IMPOSSIVEL PROSSEGUIR.", vbCritical
    'Integracao_SGA_Inmetro_Informar_CRLV = False
    'Set ofuncoes = Nothing
    'Exit Function
'End If

Set oconnreq = New CanadaConnectivity.ReqInformarArquivoCRLV

ofuncoes.ExcluiArquivos "C:\SGA\REQ\" & ofuncoes.ReturnUserWindows & ".xml"
ofuncoes.ExcluiArquivos "C:\SGA\RESP\" & ofuncoes.ReturnUserWindows & ".xml"

Set oconnreq.oCRLV = New CanadaConnectivity.CRLV
Set oconnreq.oVeiculo = New CanadaConnectivity.Veiculo

oconnreq.oCRLV.NomeArquivoCompleto = ofuncoes.GeraNomeArquivoCRLVCompleto(RetornaArquivoCRLV(WNumos), CapturaCaminhoBD())
oconnreq.oVeiculo.Placa = CapturaPlacaVeicXOSV(WNumos)
oconnreq.CreateInformarArquivoCRVL False, Retorna_Links_Uteis(7), 15, "www.canainfo.com.br"

'sleep (2000)
'[CONECTIVIDADE_NOVA] ativacao da janela do app antigo desativada:
'AppActivate ("Conectividade - " & RetornaLoginInMetro)
Do While ofuncoes.CheckFile("C:\SGA\RESP\" & ofuncoes.ReturnUserWindows & ".xml") = False
Loop

'sleep (2000)
Set oConnResp = New CanadaConnectivity.LoadRespGerarArquivoCRLV
oConnResp.Load_Objetos ofuncoes.ReturnUserWindows
ofuncoes.ExcluiArquivos "C:\SGA\REQ\" & ofuncoes.ReturnUserWindows & ".xml"
ofuncoes.ExcluiArquivos "C:\SGA\RESP\" & ofuncoes.ReturnUserWindows & ".xml"
AppActivate "SGa"

If CBool(oConnResp.Status) = True Then

    oconnreq.CreateInformarArquivoCRVL True, Retorna_Links_Uteis(7), 15, "www.canainfo.com.br"
    'sleep (2000)
'[CONECTIVIDADE_NOVA] ativacao da janela do app antigo desativada:
    'AppActivate ("Conectividade - " & RetornaLoginInMetro)
    
    Do While ofuncoes.CheckFile("C:\SGA\RESP\" & ofuncoes.ReturnUserWindows & ".xml") = False
    Loop

    'sleep (2000)
    Set oConnResp = New CanadaConnectivity.LoadRespGerarArquivoCRLV
    oConnResp.Load_Objetos ofuncoes.ReturnUserWindows
    
    If CBool(oConnResp.Status) = True Then
    
       'AppActivate "SGa"
       'MsgBox oConnResp.Message, vbInformation, "SGA"
    
    Else
    
       AppActivate "SGa"
       MsgBox oConnResp.Message, vbCritical, "SGA"
    
    End If
    
Else
    AppActivate "SGa"
    MsgBox oConnResp.Message, vbCritical, "SGA"
    
End If

ofuncoes.ExcluiArquivos "C:\SGA\REQ\" & ofuncoes.ReturnUserWindows & ".xml"
ofuncoes.ExcluiArquivos "C:\SGA\RESP\" & ofuncoes.ReturnUserWindows & ".xml"
Integracao_SGA_Inmetro_Informar_CRLV = oConnResp.Status

MsgBox "Processo Finalizado.", vbInformation

Set ofuncoes = Nothing
Set oconnreq = Nothing
Set oConnResp = Nothing

End Function

Private Sub LoadImagens(WNumos As Long)

On Error GoTo TratamentodeErro:
Dim bd As DataBase
Dim tb As Recordset
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("OSV_IMAGENS", DB_OPEN_TABLE)
Dim cont As Integer
cont = 1
tb.index = "NumOs"
tb.Seek "=", WNumos
If Not tb.NoMatch Then
   oconnreq.oCRLV.InicializaTable
   While tb!Numos = WNumos
      oconnreq.oCRLV.Add cont, tb!Caminho_Imagem
      cont = cont + 1
      tb.MoveNext
      If tb.EOF Then
          GoTo Finaliza
      End If
   Wend
End If
Finaliza:
  tb.Close
  bd.Close
  Set tb = Nothing
  Set bd = Nothing
  Exit Sub
TratamentodeErro:
    If Err.Number = 3001 Then
       MsgBox "Ocorreu um erro quando suas informações eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de número " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Sub


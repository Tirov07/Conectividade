Attribute VB_Name = "Integracao_SGA_Inmetro_Emitir_GRU"
Option Compare Database
Option Explicit
Private Status As Boolean
Private oconnreq As CanadaConnectivity.ReqEmitirGRUEnsaio
Private oConnResp As CanadaConnectivity.LoadRespEmitirGRUEnsaio
Private ofuncoes As CanadaConnectivity.FuncoesDiversas

Public Function Integracao_SGA_Inmetro_Emitir_GRU(ByVal wRenavam As String, ByVal wProtocolo As String, ByVal wPlaca As String, ByVal WNumos As Long) As Boolean

If UCase(TipoPosto) = "S" Then
    Add_Message_Status_Operacao "OPERACAO INDISPONIVEL PARA EMPRESAS DO TIPO POSTO DE SELAGEM, IMPOSSIVEL PROSSEGUIR."
    Integracao_SGA_Inmetro_Emitir_GRU = False
    Exit Function
End If

Set ofuncoes = New CanadaConnectivity.FuncoesDiversas

'[CONECTIVIDADE_NOVA] validacao da Conectividade antiga desativada:
'If ofuncoes.ConectividadeIsAlive("Conectividade") = False Then
    'Add_Message_Status_Operacao "CONECTIVIDADE NAO ESTA ATIVA, IMPOSSIVEL PROSSEGUIR."
    'Integracao_SGA_Inmetro_Emitir_GRU = False
    'Exit Function
'End If

ofuncoes.ExcluiArquivos "C:\SGA\REQ\" & ofuncoes.ReturnUserWindows & ".xml"
ofuncoes.ExcluiArquivos "C:\SGA\RESP\" & ofuncoes.ReturnUserWindows & ".xml"

Set oconnreq = New CanadaConnectivity.ReqEmitirGRUEnsaio
Set oconnreq.oSelagem = New CanadaConnectivity.Selagem
oconnreq.oSelagem.Protocolo = wProtocolo
oconnreq.CreateEmitirGRUEnsaio wPlaca, wRenavam, Retorna_Links_Uteis(4), Retorna_Links_Uteis(10), 4, "www.canainfo.com.br"
'sleep (2000)
'[CONECTIVIDADE_NOVA] ativacao da janela do app antigo desativada:
'AppActivate ("Conectividade - " & RetornaLoginInMetro)

Do While ofuncoes.CheckFile("C:\SGA\RESP\" & ofuncoes.ReturnUserWindows & ".xml") = False
Loop

'sleep (2000)

Set oConnResp = New LoadRespEmitirGRUEnsaio
oConnResp.Load_Objetos ofuncoes.ReturnUserWindows

If CBool(oConnResp.Status) = True Then
    Call Atualiza_GRU_Ensaio_OSV(WNumos, oConnResp.NumGru)
End If
Add_Message_Status_Operacao oConnResp.Message

ofuncoes.ExcluiArquivos "C:\SGA\REQ\" & ofuncoes.ReturnUserWindows & ".xml"
ofuncoes.ExcluiArquivos "C:\SGA\RESP\" & ofuncoes.ReturnUserWindows & ".xml"
Integracao_SGA_Inmetro_Emitir_GRU = oConnResp.Status
Set ofuncoes = Nothing
Set oconnreq = Nothing
Set oConnResp = Nothing
End Function


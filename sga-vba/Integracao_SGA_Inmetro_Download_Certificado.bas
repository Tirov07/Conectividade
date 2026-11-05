Attribute VB_Name = "Integracao_SGA_Inmetro_Download_Certificado"
Option Compare Database
Option Explicit

Private oconnreq As CanadaConnectivity.ReqDownLoadCertificado
Private oConnResp As CanadaConnectivity.LoadRespDownloadCertificado
Private ofuncoes As CanadaConnectivity.FuncoesDiversas

Public Function Integracao_SGA_Inmetro_Download_Certificado(ByVal WUsuario As String, ByVal wsenha As String, ByVal wRenavam As String, ByVal wPlaca As String) As Boolean

Set ofuncoes = New CanadaConnectivity.FuncoesDiversas

'[CONECTIVIDADE_NOVA] validacao da Conectividade antiga desativada:
'If ofuncoes.ConectividadeIsAlive("Conectividade") = False Then
    'Add_Message_Status_Operacao "CONECTIVIDADE NAO ESTA ATIVA, IMPOSSIVEL PROSSEGUIR."
    'Integracao_SGA_Inmetro_Download_Certificado = False
    'Exit Function
'End If

ofuncoes.ExcluiArquivos "C:\SGA\REQ\" & ofuncoes.ReturnUserWindows & ".xml"
ofuncoes.ExcluiArquivos "C:\SGA\RESP\" & ofuncoes.ReturnUserWindows & ".xml"

Set oconnreq = New CanadaConnectivity.ReqDownLoadCertificado
Set oconnreq.oVeiculo = New CanadaConnectivity.Veiculo
oconnreq.oVeiculo.Placa = wPlaca
oconnreq.oVeiculo.Renavam = wRenavam
oconnreq.CreateDownloadCertificado Retorna_Links_Uteis(7), 7, False, "www.canainfo.com.br"
'[CONECTIVIDADE_NOVA] ativacao da janela do app antigo desativada:
'AppActivate ("Conectividade - " & RetornaLoginInMetro)
Do While ofuncoes.CheckFile("C:\SGA\RESP\" & ofuncoes.ReturnUserWindows & ".xml") = False
Loop
'sleep (4000)
AppActivate "SGa"
Set oConnResp = New CanadaConnectivity.LoadRespDownloadCertificado
oConnResp.Load_Objetos ofuncoes.ReturnUserWindows
ofuncoes.ExcluiArquivos "C:\SGA\REQ\" & ofuncoes.ReturnUserWindows & ".xml"
ofuncoes.ExcluiArquivos "C:\SGA\RESP\" & ofuncoes.ReturnUserWindows & ".xml"
If CBool(oConnResp.Status) = False Then
    Add_Message_Status_Operacao UCase(oConnResp.Message)
Else
   oconnreq.CreateDownloadCertificado Retorna_Links_Uteis(7), 7, True, "www.canainfo.com.br"
'[CONECTIVIDADE_NOVA] ativacao da janela do app antigo desativada:
   'AppActivate ("Conectividade - " & RetornaLoginInMetro)
   Do While ofuncoes.CheckFile("C:\SGA\RESP\" & ofuncoes.ReturnUserWindows & ".xml") = False
   Loop
   'sleep (2000)
   Set oConnResp = New CanadaConnectivity.LoadRespDownloadCertificado
   oConnResp.Load_Objetos ofuncoes.ReturnUserWindows
   ofuncoes.ExcluiArquivos "C:\SGA\REQ\" & ofuncoes.ReturnUserWindows & ".xml"
   ofuncoes.ExcluiArquivos "C:\SGA\RESP\" & ofuncoes.ReturnUserWindows & ".xml"
End If
Add_Message_Status_Operacao UCase(oConnResp.Message)
'AppActivate "SGa"
Integracao_SGA_Inmetro_Download_Certificado = CBool(oConnResp.Status)
Set ofuncoes = Nothing
Set oconnreq = Nothing
Set oConnResp = Nothing
End Function




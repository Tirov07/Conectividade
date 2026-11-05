Attribute VB_Name = "Integracao_SGA_Inmetro_Cadastrar_Veiculo"
Option Compare Database
Option Explicit

Private oconn As CanadaConnectivity.ReqCadastrarVeiculo
Private oConnResp As CanadaConnectivity.LoadRespCadastrarVeiculo
Private ofuncoes As CanadaConnectivity.FuncoesDiversas

Public Function Integracao_SGA_Inmetro_Cadastrar_Veiculo(ByVal wRenavam As String, ByVal wPlaca As String, ByVal wChassi As String, ByVal wTipoEmplacamento As String, ByVal WMarcaVeiculo As String, ByVal WModelo As String, ByVal WAno As String, ByVal wAro As String, ByVal WPneu As String, ByVal wPaisOrigem As String, ByVal WRazaoSocial As String, ByVal wNroDocumento As String, ByVal wcep As String, ByVal WUF As String, ByVal WCidade As String, ByVal WBairro As String, ByVal WEndereco As String, ByVal wnro As String, ByVal WComplemento As String, ByVal WEmail As String, ByVal wTelefone As String) As Boolean

Set ofuncoes = New CanadaConnectivity.FuncoesDiversas
Set oconn = New CanadaConnectivity.ReqCadastrarVeiculo
Set oconn.oVeiculo = New CanadaConnectivity.Veiculo
Set oconn.oProprietario = New CanadaConnectivity.Proprietario
Set oConnResp = New CanadaConnectivity.LoadRespCadastrarVeiculo

'[CONECTIVIDADE_NOVA] validacao da Conectividade antiga desativada:
'If ofuncoes.ConectividadeIsAlive("Conectividade") = False Then
    'Add_Message_Status_Operacao "CONECTIVIDADE NAO ESTA ATIVA, IMPOSSIVEL PROSSEGUIR."
    'Integracao_SGA_Inmetro_Cadastrar_Veiculo = False
    'Exit Function
'End If

ofuncoes.ExcluiArquivos "C:\SGA\REQ\" & ofuncoes.ReturnUserWindows & ".xml"
ofuncoes.ExcluiArquivos "C:\SGA\RESP\" & ofuncoes.ReturnUserWindows & ".xml"

With oconn.oVeiculo
     .Renavam = wRenavam
     .Placa = wPlaca
     .Ano = WAno
     .Pneu = WPneu
     .Aro = wAro
     .AroValue = RetornaValueAroPneuVeiculo(wAro)
     .CHASSI = wChassi
     .Marca = WMarcaVeiculo
     .MarcaValue = RetornaValueMarcaVeiculo(WMarcaVeiculo)
     .ModeloValue = RetornaValueModeloVeiculo(WMarcaVeiculo, WModelo)
     .Modelo = WModelo
     .Pais = "Brasil"
     .TipoEmplacamento = "Veículo emplacado no país"
End With

With oconn.oProprietario
    .RazaoSocial = WRazaoSocial
    .Documento = wNroDocumento
    .Cidade = WCidade
    .Bairro = WBairro
    .Endereco = WEndereco
    .Numero = wnro
    .Complemento = WComplemento
    .Email = WEmail
    .Telefone = wTelefone
    .UF = WUF
    .Cep = wcep
    .Pais = "BRASIL"
End With

oconn.CreateCadastraVeiculo Retorna_Links_Uteis(2), 2, "www.canainfo.com.br"
'Sleep (2000) '
'[CONECTIVIDADE_NOVA] ativacao da janela do app antigo desativada:
'AppActivate ("Conectividade - " & RetornaLoginInMetro)
Do While ofuncoes.CheckFile("C:\SGA\RESP\" & ofuncoes.ReturnUserWindows & ".xml") = False
Loop

'Sleep (2000)

oConnResp.Load_Objetos ofuncoes.ReturnUserWindows
Add_Message_Status_Operacao oConnResp.Message
AppActivate "SGa"
Integracao_SGA_Inmetro_Cadastrar_Veiculo = oConnResp.Status
ofuncoes.ExcluiArquivos "C:\SGA\REQ\" & ofuncoes.ReturnUserWindows & ".xml"
ofuncoes.ExcluiArquivos "C:\SGA\RESP\" & ofuncoes.ReturnUserWindows & ".xml"

Set ofuncoes = Nothing
Set oconn = Nothing
Set oConnResp = Nothing

End Function


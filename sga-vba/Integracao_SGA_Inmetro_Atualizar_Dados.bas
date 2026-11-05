Attribute VB_Name = "Integracao_SGA_Inmetro_Atualizar_Dados"
Option Compare Database
Option Explicit
Private ofuncoes As CanadaConnectivity.FuncoesDiversas
Private cont As Integer
Public CountVeiculosCNPJ As Long

Public Function Integracao_SGA_Inmetro_Atualizar_Aro_Pneu_Veiculo() As Boolean

Dim oconnreq As CanadaConnectivity.ReqAtualizarAroPneuVeiculo
Dim oConnResp As CanadaConnectivity.LoadRespAtualizarAroPneuVeiculo

Set ofuncoes = New CanadaConnectivity.FuncoesDiversas

'[CONECTIVIDADE_NOVA] validacao da Conectividade antiga desativada:
'If ofuncoes.ConectividadeIsAlive("Conectividade") = False Then
    'MsgBox "CONECTIVIDADE NAO ESTA ATIVA, IMPOSSIVEL PROSSEGUIR.", vbCritical
    'Integracao_SGA_Inmetro_Atualizar_Aro_Pneu_Veiculo = False
    'Exit Function
'End If

ofuncoes.ExcluiArquivos "C:\SGA\REQ\" & ofuncoes.ReturnUserWindows & ".xml"
ofuncoes.ExcluiArquivos "C:\SGA\RESP\" & ofuncoes.ReturnUserWindows & ".xml"

Set oconnreq = New CanadaConnectivity.ReqAtualizarAroPneuVeiculo
Set oconnreq.oVeiculo = New CanadaConnectivity.Veiculo
oconnreq.oVeiculo.Placa = "xxx1111"

Loading_Page
oconnreq.CreateAtualizarMarcasVeiculo Retorna_Links_Uteis(2), 12, "www.canainfo.com.br"
'[CONECTIVIDADE_NOVA] ativacao da janela do app antigo desativada:
'AppActivate ("Conectividade - " & RetornaLoginInMetro)

Do While ofuncoes.CheckFile("C:\SGA\RESP\" & ofuncoes.ReturnUserWindows & ".xml") = False
Loop

'sleep (2000)
Set oConnResp = New CanadaConnectivity.LoadRespAtualizarAroPneuVeiculo
oConnResp.Load_Objetos ofuncoes.ReturnUserWindows

If oConnResp.Status = True Then
    For cont = 0 To oConnResp.oVeiculo.AroPneuCount - 1
        Status_Update Trim$(UCase(oConnResp.oVeiculo.String_AroPneu(cont)))
        Call PadronizaAroPneu(Trim$(UCase(oConnResp.oVeiculo.String_AroPneu(cont))), oConnResp.oVeiculo.String_AroPneu_Value(cont))
    Next
    MsgBox "Atualizacao Realizada Com Sucesso!", vbInformation
Else
    MsgBox oConnResp.Message, vbCritical
End If
Status_Finalize

AppActivate "SGa"
Integracao_SGA_Inmetro_Atualizar_Aro_Pneu_Veiculo = oConnResp.Status
ofuncoes.ExcluiArquivos "C:\SGA\REQ\" & ofuncoes.ReturnUserWindows & ".xml"
ofuncoes.ExcluiArquivos "C:\SGA\RESP\" & ofuncoes.ReturnUserWindows & ".xml"
Set ofuncoes = Nothing
Set oconnreq = Nothing
Set oConnResp = Nothing
End Function

Public Function Integracao_SGA_Inmetro_Atualizar_Marca_Veiculo() As Boolean

Dim oconnreq As CanadaConnectivity.ReqAtualizarMarcasVeiculo
Dim oConnResp As CanadaConnectivity.LoadRespAtualizarMarcasVeiculo

Set ofuncoes = New CanadaConnectivity.FuncoesDiversas

'[CONECTIVIDADE_NOVA] validacao da Conectividade antiga desativada:
'If ofuncoes.ConectividadeIsAlive("Conectividade") = False Then
    'MsgBox "CONECTIVIDADE NAO ESTA ATIVA, IMPOSSIVEL PROSSEGUIR.", vbCritical
    'Integracao_SGA_Inmetro_Atualizar_Marca_Veiculo = False
    'Exit Function
'End If

ofuncoes.ExcluiArquivos "C:\SGA\REQ\" & ofuncoes.ReturnUserWindows & ".xml"
ofuncoes.ExcluiArquivos "C:\SGA\RESP\" & ofuncoes.ReturnUserWindows & ".xml"

Set oconnreq = New CanadaConnectivity.ReqAtualizarMarcasVeiculo
Set oconnreq.oVeiculo = New CanadaConnectivity.Veiculo
oconnreq.oVeiculo.Placa = "xxx1111"

Loading_Page
oconnreq.CreateAtualizarMarcasVeiculo Retorna_Links_Uteis(2), 11, "www.canainfo.com.br"
'[CONECTIVIDADE_NOVA] ativacao da janela do app antigo desativada:
'AppActivate ("Conectividade - " & RetornaLoginInMetro)

Do While ofuncoes.CheckFile("C:\SGA\RESP\" & ofuncoes.ReturnUserWindows & ".xml") = False
Loop

'sleep (2000)
Set oConnResp = New CanadaConnectivity.LoadRespAtualizarMarcasVeiculo
oConnResp.Load_Objetos ofuncoes.ReturnUserWindows

If oConnResp.Status = True Then
    For cont = 0 To oConnResp.oVeiculo.MarcasCount - 1
        Status_Update Trim$(UCase(oConnResp.oVeiculo.String_Marca(cont)))
        Call PadronizaMarcaVeiculo(Trim$(UCase(oConnResp.oVeiculo.String_Marca(cont))), oConnResp.oVeiculo.String_Marca_Value(cont))
    Next
    MsgBox "Atualizacao Realizada Com Sucesso!", vbInformation
Else
    MsgBox oConnResp.Message, vbCritical
End If
Status_Finalize

AppActivate "SGa"
Integracao_SGA_Inmetro_Atualizar_Marca_Veiculo = oConnResp.Status
ofuncoes.ExcluiArquivos "C:\SGA\REQ\" & ofuncoes.ReturnUserWindows & ".xml"
ofuncoes.ExcluiArquivos "C:\SGA\RESP\" & ofuncoes.ReturnUserWindows & ".xml"
Set ofuncoes = Nothing
Set oconnreq = Nothing
Set oConnResp = Nothing
End Function

Public Function Integracao_SGA_Inmetro_Atualizar_Modelo(ByVal WMarcaVeiculo As String) As Boolean

Dim oconnreq As CanadaConnectivity.ReqAtualizarModelosVeiculo
Dim oConnResp As CanadaConnectivity.LoadRespAtualizarModelosVeiculo

Set ofuncoes = New CanadaConnectivity.FuncoesDiversas

'[CONECTIVIDADE_NOVA] validacao da Conectividade antiga desativada:
'If ofuncoes.ConectividadeIsAlive("Conectividade") = False Then
    'MsgBox "CONECTIVIDADE NAO ESTA ATIVA, IMPOSSIVEL PROSSEGUIR.", vbCritical
    'Integracao_SGA_Inmetro_Atualizar_Modelo = False
    'Exit Function
'End If

Set oconnreq = New CanadaConnectivity.ReqAtualizarModelosVeiculo
Set oconnreq.oVeiculo = New CanadaConnectivity.Veiculo

ofuncoes.ExcluiArquivos "C:\SGA\REQ\" & ofuncoes.ReturnUserWindows & ".xml"
ofuncoes.ExcluiArquivos "C:\SGA\RESP\" & ofuncoes.ReturnUserWindows & ".xml"

oconnreq.oVeiculo.Marca = WMarcaVeiculo
oconnreq.oVeiculo.MarcaValue = RetornaValueMarcaVeiculo(WMarcaVeiculo)
Loading_Page_Marcas_Veiculos WMarcaVeiculo
oconnreq.CreateAtualizarModelosVeiculo Retorna_Links_Uteis(2), 10, "www.canainfo.com.br"
'[CONECTIVIDADE_NOVA] ativacao da janela do app antigo desativada:
'AppActivate ("Conectividade - " & RetornaLoginInMetro)

Do While ofuncoes.CheckFile("C:\SGA\RESP\" & ofuncoes.ReturnUserWindows & ".xml") = False
Loop

'sleep (2000)
Set oConnResp = New CanadaConnectivity.LoadRespAtualizarModelosVeiculo
oConnResp.Load_Objetos ofuncoes.ReturnUserWindows

If CBool(oConnResp.Status) = True Then
    For cont = 0 To oConnResp.oVeiculo.ModelosCount - 1
        Status_Modelo UCase(Trim$(oConnResp.oVeiculo.String_Marca_Modelo(cont))), UCase(Trim$(oConnResp.oVeiculo.String_Modelo(cont)))
        Call Padroniza_Modelo_Veic(UCase(Trim$(oConnResp.oVeiculo.String_Marca_Modelo(cont))), UCase(Trim$(oConnResp.oVeiculo.String_Modelo(cont))), oConnResp.oVeiculo.String_Modelo_Value(cont))
    Next
End If
Status_Finalize
AppActivate "SGa"
Integracao_SGA_Inmetro_Atualizar_Modelo = oConnResp.Status
ofuncoes.ExcluiArquivos "C:\SGA\REQ\" & ofuncoes.ReturnUserWindows & ".xml"
ofuncoes.ExcluiArquivos "C:\SGA\RESP\" & ofuncoes.ReturnUserWindows & ".xml"
Set ofuncoes = Nothing
Set oconnreq = Nothing
Set oConnResp = Nothing
End Function
Public Function Integracao_SGA_Inmetro_Atualizar_Marca_Modelo_Equipamento() As Boolean

Dim oconnreq As CanadaConnectivity.ReqAtualizarMarcasModelosEquipamentos
Dim oConnResp As CanadaConnectivity.LoadRespAtualizarModelosEquipamento
Dim cont As Integer

Set ofuncoes = New CanadaConnectivity.FuncoesDiversas

'[CONECTIVIDADE_NOVA] validacao da Conectividade antiga desativada:
'If ofuncoes.ConectividadeIsAlive("Conectividade") = False Then
    'MsgBox "CONECTIVIDADE NAO ESTA ATIVA, IMPOSSIVEL PROSSEGUIR.", vbCritical
    'Integracao_SGA_Inmetro_Atualizar_Marca_Modelo_Equipamento = False
    'Exit Function
'End If

Set oconnreq = New CanadaConnectivity.ReqAtualizarMarcasModelosEquipamentos

ofuncoes.ExcluiArquivos "C:\SGA\REQ\" & ofuncoes.ReturnUserWindows & ".xml"
ofuncoes.ExcluiArquivos "C:\SGA\RESP\" & ofuncoes.ReturnUserWindows & ".xml"

oconnreq.CreateAtualizarModelosEquipamentos Retorna_Links_Uteis(9), 9, "www.canainfo.com.br"
'[CONECTIVIDADE_NOVA] ativacao da janela do app antigo desativada:
'AppActivate ("Conectividade - " & RetornaLoginInMetro)

Do While ofuncoes.CheckFile("C:\SGA\RESP\" & ofuncoes.ReturnUserWindows & ".xml") = False
Loop

'sleep (2000)
Set oConnResp = New CanadaConnectivity.LoadRespAtualizarModelosEquipamento
oConnResp.Load_Objetos ofuncoes.ReturnUserWindows

AppActivate "SGa"

If CBool(oConnResp.Status) = True Then
    For cont = 0 To oConnResp.oEquipamento.ModelosCount - 1
        Status_Page_Modelos_Equipamento oConnResp.oEquipamento.StringMarca(cont), oConnResp.oEquipamento.StringModelo(cont)
        PadronizaMarcaEquip (Trim(UCase(oConnResp.oEquipamento.StringMarca(cont))))
        Padroniza_Modelo_Equip oConnResp.oEquipamento.StringMarca(cont), oConnResp.oEquipamento.StringModelo(cont), oConnResp.oEquipamento.AdesivoMinimo(cont), oConnResp.oEquipamento.PlasticoMinimo(cont), oConnResp.oEquipamento.AdesivoMaximo(cont), oConnResp.oEquipamento.PlasticoMaximo(cont), oConnResp.oEquipamento.Tipo(cont)
    Next cont
    MsgBox "Atualização Realizada Com Sucesso!", vbInformation
    Status_Finalize
End If
Integracao_SGA_Inmetro_Atualizar_Marca_Modelo_Equipamento = oConnResp.Status
ofuncoes.ExcluiArquivos "C:\SGA\REQ\" & ofuncoes.ReturnUserWindows & ".xml"
ofuncoes.ExcluiArquivos "C:\SGA\RESP\" & ofuncoes.ReturnUserWindows & ".xml"
Set ofuncoes = Nothing
Set oconnreq = Nothing
Set oConnResp = Nothing
End Function

Private Sub Status_Page_Modelos_Equipamento(ByVal WMarca As String, ByVal WModelo As String)
Dim Ret As Variant
Ret = SysCmd(SYSCMD_SETSTATUS, "Atualizando Dados Dos Equipamentos Marca: " & WMarca & " >>> " & "Modelo: " & WModelo & ".")
End Sub

Private Sub Loading_Page_Marcas_Veiculos(ByVal WString As String)
Dim Ret As Variant
Ret = SysCmd(SYSCMD_SETSTATUS, "Atualizando Modelos da Marca..." & WString & ">>> Aguarde Por Favor.")
End Sub
Private Sub Status_Modelo(ByVal WMarca As String, ByVal WModelo As String)
Dim Ret As Variant
Ret = SysCmd(SYSCMD_SETSTATUS, "Atualizando Marca >>> " & WMarca & ". Modelo >>> " & WModelo & ". Aguarde Por Favor.")
End Sub
Private Sub Status_Finalize()
Dim Ret As Variant
Ret = SysCmd(SYSCMD_SETSTATUS, "Concluido.")
End Sub
Private Sub Loading_Page()
Dim Ret As Variant
   Ret = SysCmd(SYSCMD_SETSTATUS, "Processando... Aguarde Por Favor.")
End Sub
Private Sub Status_Update(ByVal WStatus As String)
Dim Ret As Variant
Ret = SysCmd(SYSCMD_SETSTATUS, "Atualizando Marca >>> " & WStatus & ".")
End Sub

Public Function Integracao_SGA_Inmetro_Atualizar_Veiculos_CNPJ(ByVal WCodigoCliente As Long) As Boolean

Dim oconnreq As CanadaConnectivity.ReqConsultaVeiculoxCNPJ
Dim oConnResp As CanadaConnectivity.LoadRespConsultarVeiculoCNPJ

Set ofuncoes = New CanadaConnectivity.FuncoesDiversas

InicializaVariaveisConectividade

If IsCloud = False Then
'[CONECTIVIDADE_NOVA] validacao da Conectividade antiga desativada:
   'If ofuncoes.ConectividadeIsAlive("Conectividade") = False Then
      'MsgBox "CONECTIVIDADE NAO ESTA ATIVA, IMPOSSIVEL PROSSEGUIR.", vbCritical
      'Integracao_SGA_Inmetro_Atualizar_Veiculos_CNPJ = False
      'Exit Function
   'End If
Else
   If ofuncoes.ExisteDiretorio(CaminhoSGaRemote) = False Then
      MsgBox "CONEXAO OFF LINE, IMPOSSIVEL PROSSEGUIR.", vbCritical
      Integracao_SGA_Inmetro_Atualizar_Veiculos_CNPJ = False
      Exit Function
   End If
End If
ofuncoes.ExcluiArquivos pFileNameReq
ofuncoes.ExcluiArquivos pFileNameResp

Set oconnreq = New CanadaConnectivity.ReqConsultaVeiculoxCNPJ
Set oconnreq.oProprietario = New CanadaConnectivity.Proprietario
oconnreq.oProprietario.Documento = Captura_Conteudo_Campos_CLI(WCodigoCliente, "CgcCpf")

Loading_Page
oconnreq.CreateConsultaVeiculoDocumento Retorna_Links_Uteis(17), 17, "www.canainfo.com.br"

If IsCloud = False Then
'[CONECTIVIDADE_NOVA] ativacao da janela do app antigo desativada:
    'AppActivate ("Conectividade - " & RetornaLoginInMetro)
End If

Do While ofuncoes.CheckFile(pFileNameResp) = False
Loop

'sleep (2000)
Set oConnResp = New CanadaConnectivity.LoadRespConsultarVeiculoCNPJ
oConnResp.Load_Objetos pNameUserCloud, IsCloud

If oConnResp.Status = True Then
    Dim contadorVeiculos As Integer
    contadorVeiculos = 0
    CountVeiculosCNPJ = oConnResp.oVeiculo.VeiculosCount
    For contadorVeiculos = 0 To oConnResp.oVeiculo.VeiculosCount - 1
       If oConnResp.oVeiculo.String_Placa(contadorVeiculos) = "-" Then
          '##### Veiculos Nao Emplacados #####
          Dim sPlacaAux As String
          sPlacaAux = Mid(oConnResp.oVeiculo.String_Chassi(contadorVeiculos), 1, 7)
          Call Atualiza_Tabela_PLACXCLI_Placa_Renavam_Chassi(sPlacaAux, WCodigoCliente, oConnResp.oVeiculo.String_Renavam(contadorVeiculos), oConnResp.oVeiculo.String_Chassi(contadorVeiculos), 1, "C")
       Else
          '##### Veiculos Emplacados #####
          If VerificaExistePlacaxCli_Chassi(oConnResp.oVeiculo.String_Chassi(contadorVeiculos), WCodigoCliente) = True Then
             Call Atualiza_Tabela_PLACXCLI_Inmetro_Chassi(oConnResp.oVeiculo.String_Chassi(contadorVeiculos), WCodigoCliente, oConnResp.oVeiculo.String_Renavam(contadorVeiculos), oConnResp.oVeiculo.String_Placa(contadorVeiculos), 1, "C")
          Else
             '########## Criei uma funcao Atualiza_Tabela_PlacxCli_placa_Renavam_Chassi onde atualiza-se os 3 campos de uma vez.. Placa Renavam e Chassi.. Atualizando separadamente fica lento.. pode-se testar.#########
             Call Atualiza_Tabela_PLACXCLI_Placa_Renavam_Chassi(oConnResp.oVeiculo.String_Placa(contadorVeiculos), WCodigoCliente, oConnResp.oVeiculo.String_Renavam(contadorVeiculos), oConnResp.oVeiculo.String_Chassi(contadorVeiculos), 1, "C")
          End If
       End If
    Next
Else
    MsgBox oConnResp.Message, vbCritical
End If
Status_Finalize
AppActivate "SGa"
Integracao_SGA_Inmetro_Atualizar_Veiculos_CNPJ = oConnResp.Status
ofuncoes.ExcluiArquivos "C:\SGA\REQ\" & ofuncoes.ReturnUserWindows & ".xml"
ofuncoes.ExcluiArquivos "C:\SGA\RESP\" & ofuncoes.ReturnUserWindows & ".xml"
ofuncoes.ExcluiArquivos pFileNameReq
ofuncoes.ExcluiArquivos pFileNameResp
Integracao_SGA_Inmetro_Atualizar_Veiculos_CNPJ = oConnResp.Status
Set ofuncoes = Nothing
Set oconnreq = Nothing
Set oConnResp = Nothing
End Function
Public Sub ConsultaVeiculosxCNPJ_Inmetro_Analitico(ByVal WCodigoCliente As Long)
 If ConectividadeAtiva(Empresa_Ativa) = True Then
    If Integracao_SGA_Inmetro_Atualizar_Veiculos_CNPJ(WCodigoCliente) = True Then
       Call Integracao_SGA_Inmetro_Consulta_Selagem_Detalhada_Veiculo_CNPJ(WCodigoCliente)
       AppActivate "SGa"
       MsgBox "O SGa Importou Todos Os " & Format$(CountVeiculosCNPJ, "00000") & " Veiculos Vinculados a Esse Cliente Na Base de Dados do Inmetro.", vbInformation
    End If
End If
AppActivate "SGa"
End Sub
Public Sub ConsultaVeiculosxCNPJ_Inmetro_Sintetico(ByVal WCodigoCliente As Long)
CountVeiculosCNPJ = 0
'If Captura_Conteudo_Campos_Empresa(Empresa_Ativa, "TipoPosto") = "S" Or Captura_Conteudo_Campos_Empresa(Empresa_Ativa, "TipoPosto") = "E" Then
If ConectividadeAtiva(Empresa_Ativa) = True Then
   If MsgBox("Deseja Sincronizar Os Veiculos da Base de Dados Inmetro ? ", 4 + 32, "ATENÇÃO") = 6 Then
      If Integracao_SGA_Inmetro_Atualizar_Veiculos_CNPJ(WCodigoCliente) = True Then
         AppActivate "SGa"
         'If CountVeiculosCNPJ <= 5 Then
            Call Integracao_SGA_Inmetro_Consulta_Selagem_Detalhada_Veiculo_CNPJ(WCodigoCliente)
         'End If
         AppActivate "SGa"
         MsgBox "O SGa Importou Todos Os " & Format$(CountVeiculosCNPJ, "00000") & " Veiculos Vinculados a Esse Cliente Na Base de Dados do Inmetro.", vbInformation
     End If
   End If
End If
AppActivate "SGa"
End Sub
Private Sub testeconsultacnpj()
'Call ConsultaVeiculoxCNPJ_Inmetro(57489)
End Sub


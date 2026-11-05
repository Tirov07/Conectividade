Attribute VB_Name = "Integracao_SGA_Inmetro_Informar_Selagem"
Option Compare Database
Option Explicit
#If VBA7 And Win64 Then
     Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
 ' Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
Private oconnreq As CanadaConnectivity.ReqInformarSelagem
Private oConnResp As CanadaConnectivity.LoadRespInformarSelagem
Private ofuncoes As CanadaConnectivity.FuncoesDiversas
Public Function Integracao_SGA_Inmetro_Informar_Selagem(ByVal wNumGRU As String, ByVal WNumSerieTacografo As String, ByVal WConstK As String, ByVal WSelos_Adesivos As String, ByVal WSelos_Plasticos As String, ByVal WMarcaTacografo As String, ByVal WModeloTacografo As String, ByVal WVelInf As String, ByVal WRedutor As String, ByVal wRenavam As String, ByVal wPlaca As String, ByVal wAro As String, ByVal WPneu As String, ByVal wTipoRegistro As String, ByVal wEscalaVelocidade As String, ByVal wAnoTacografo As String, ByVal wPAM As String, ByVal WNumos As String, ByVal wChassi As String) As Boolean
Set ofuncoes = New CanadaConnectivity.FuncoesDiversas

InicializaVariaveisConectividade

If IsCloud = False Then
'[CONECTIVIDADE_NOVA] validacao da Conectividade antiga desativada:
   'If ofuncoes.ConectividadeIsAlive("Conectividade") = False Then
      'MsgBox "CONECTIVIDADE NAO ESTA ATIVA, IMPOSSIVEL PROSSEGUIR.", vbCritical
      'Integracao_SGA_Inmetro_Informar_Selagem = False
      'Exit Function
   'End If
Else
   If ofuncoes.ExisteDiretorio(CaminhoSGaRemote) = False Then
      MsgBox "CONEXAO OFF LINE, IMPOSSIVEL PROSSEGUIR.", vbCritical
      Integracao_SGA_Inmetro_Informar_Selagem = False
      Exit Function
   End If
End If

ofuncoes.ExcluiArquivos pFileNameReq
ofuncoes.ExcluiArquivos pFileNameResp

Set oconnreq = New CanadaConnectivity.ReqInformarSelagem
Set oConnResp = New CanadaConnectivity.LoadRespInformarSelagem

Set oconnreq.oEquipamento = New CanadaConnectivity.Equipamento
oconnreq.oEquipamento.Nroserie = WNumSerieTacografo
oconnreq.oEquipamento.Marca = WMarcaTacografo
oconnreq.oEquipamento.Modelo = WModeloTacografo
oconnreq.oEquipamento.Redutor = WRedutor
oconnreq.oEquipamento.ConstK = WConstK
oconnreq.oEquipamento.TipoRegistro = wTipoRegistro
oconnreq.oEquipamento.EscalaVelocidade = wEscalaVelocidade
oconnreq.oEquipamento.AnoInstrumento = wAnoTacografo
oconnreq.oEquipamento.PAM = wPAM

Set oconnreq.oVeiculo = New CanadaConnectivity.Veiculo
oconnreq.oVeiculo.Placa = wPlaca
oconnreq.oVeiculo.Renavam = wRenavam
oconnreq.oVeiculo.Pneu = WPneu
oconnreq.oVeiculo.Aro = wAro
oconnreq.oVeiculo.CHASSI = wChassi

Set oconnreq.oSelagem = New CanadaConnectivity.Selagem
oconnreq.oSelagem.GRUSELAGEM = wNumGRU
oconnreq.oSelagem.Velocidade = WVelInf
oconnreq.oSelagem.SelosAdesivos = CapturaSelosAplicadosXCodigoSelo(CLng(WNumos), "A")
oconnreq.oSelagem.SelosPlasticos = CapturaSelosAplicadosXCodigoSelo(CLng(WNumos), "P")
oconnreq.oSelagem.SeloPlasticoCaixaSensor = Return_Selo_Caixa(CLng(WNumos), "P")

Loading_Page

oconnreq.CreateInformarSelagem TipoPosto, Retorna_Links_Uteis(3), Retorna_Links_Uteis(7), Retorna_Links_Uteis(2), 3, "www.canainfo.com.br"

If IsCloud = False Then
'[CONECTIVIDADE_NOVA] ativacao da janela do app antigo desativada:
    'AppActivate ("Conectividade - " & RetornaLoginInMetro)
End If

Do While ofuncoes.CheckFile(pFileNameResp) = False
   Sleep (500)
   DoEvents
Loop

'sleep (2000)

oConnResp.Load_Objetos pNameUserCloud, IsCloud
If CBool(oConnResp.Status) = True Then
    Atualiza_Nro_Protocolo_OSV Val(WNumos), oConnResp.NumeroProtocolo
    Call Atualiza_Status_Selagem_OSV(Val(WNumos), oConnResp.Situacao)
End If

Add_Message_Status_Operacao UCase(oConnResp.Message)
Finalize_Page
Integracao_SGA_Inmetro_Informar_Selagem = oConnResp.Status
If UCase(TipoPosto()) = "E" Then AppActivate "SGa"

ofuncoes.ExcluiArquivos CaminhoReqLocal & "\" & ofuncoes.ReturnUserWindows & ".xml"
ofuncoes.ExcluiArquivos CaminhoRespLocal & "\" & ofuncoes.ReturnUserWindows & ".xml"
   
ofuncoes.ExcluiArquivos pFileNameReq
ofuncoes.ExcluiArquivos pFileNameResp

Set ofuncoes = Nothing
Set oconnreq = Nothing
Set oConnResp = Nothing

End Function
Private Sub Loading_Page()
Dim Ret As Variant
Ret = SysCmd(SYSCMD_SETSTATUS, "Aguarde, processando...")
End Sub
Private Sub Finalize_Page()
Dim Ret As Variant
Ret = SysCmd(SYSCMD_SETSTATUS, "Processo Finalizado...")
End Sub


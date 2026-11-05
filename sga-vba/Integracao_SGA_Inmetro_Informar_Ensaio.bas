Attribute VB_Name = "Integracao_SGA_Inmetro_Informar_Ensaio"
Option Compare Database
Option Explicit
#If VBA7 And Win64 Then
     Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
 ' Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private oconnreq As CanadaConnectivity.ReqInformarEnsaio
Private oConnResp As CanadaConnectivity.LoadRespInformarEnsaio
Private ofuncoes As CanadaConnectivity.FuncoesDiversas
Public Function Integracao_SGA_Inmetro_Informar_Ensaio(ByVal wNumGRU As String, ByVal wData_Emissao_Laudo As String, ByVal WVel_Max As String, ByVal WPadrao_Utilizado As String, ByVal wRenavam As String, ByVal wPlaca As String, ByVal wNumeroProtocolo As String) As Boolean

If UCase(TipoPosto) = "S" Then
    Add_Message_Status_Operacao "OPERACAO INDISPONIVEL PARA EMPRESAS DO TIPO POSTO DE SELAGEM, IMPOSSIVEL PROSSEGUIR."
    Integracao_SGA_Inmetro_Informar_Ensaio = False
    Exit Function
End If

Set ofuncoes = New CanadaConnectivity.FuncoesDiversas

InicializaVariaveisConectividade

If IsCloud = False Then
'[CONECTIVIDADE_NOVA] validacao da Conectividade antiga desativada:
   'If ofuncoes.ConectividadeIsAlive("Conectividade") = False Then
      'MsgBox "CONECTIVIDADE NAO ESTA ATIVA, IMPOSSIVEL PROSSEGUIR.", vbCritical
      'Integracao_SGA_Inmetro_Informar_Ensaio = False
      'Exit Function
   'End If
Else
   If ofuncoes.ExisteDiretorio(CaminhoSGaRemote) = False Then
      MsgBox "CONEXAO OFF LINE, IMPOSSIVEL PROSSEGUIR.", vbCritical
      Integracao_SGA_Inmetro_Informar_Ensaio = False
      Exit Function
   End If
End If

ofuncoes.ExcluiArquivos pFileNameReq
ofuncoes.ExcluiArquivos pFileNameResp

Set oconnreq = New CanadaConnectivity.ReqInformarEnsaio

Set oconnreq.oEnsaio = New CanadaConnectivity.Ensaio
oconnreq.oEnsaio.DataEnsaio = wData_Emissao_Laudo
oconnreq.oEnsaio.NumeroGRU = wNumGRU
oconnreq.oEnsaio.PadraoUtilizado = WPadrao_Utilizado
oconnreq.oEnsaio.VelMaxEnsaio = WVel_Max

Set oconnreq.oSelagem = New CanadaConnectivity.Selagem
oconnreq.oSelagem.Protocolo = wNumeroProtocolo

Set oconnreq.oVeiculo = New CanadaConnectivity.Veiculo
oconnreq.oVeiculo.Placa = wPlaca
oconnreq.oVeiculo.Renavam = wRenavam

Loading_Page
oconnreq.CreateInformarEnsaio Retorna_Links_Uteis(5), Retorna_Links_Uteis(7), 5, "www.canainfo.com.br"

If IsCloud = False Then
'[CONECTIVIDADE_NOVA] ativacao da janela do app antigo desativada:
    'AppActivate ("Conectividade - " & RetornaLoginInMetro)
End If

Do While ofuncoes.CheckFile(pFileNameResp) = False
   Sleep (500)
   DoEvents
Loop

Sleep (2000)

Set oConnResp = New LoadRespInformarEnsaio
oConnResp.Load_Objetos pNameUserCloud, IsCloud

Add_Message_Status_Operacao oConnResp.Message
Integracao_SGA_Inmetro_Informar_Ensaio = oConnResp.Status

ofuncoes.ExcluiArquivos CaminhoReqLocal & "\" & ofuncoes.ReturnUserWindows & ".xml"
ofuncoes.ExcluiArquivos CaminhoRespLocal & "\" & ofuncoes.ReturnUserWindows & ".xml"
   
ofuncoes.ExcluiArquivos pFileNameReq
ofuncoes.ExcluiArquivos pFileNameResp

AppActivate "SGa"
Set ofuncoes = Nothing
Set oconnreq = Nothing
Set oConnResp = Nothing
End Function

Public Function Integracao_SGA_Inmetro_Informar_PreEnsaio(ByVal WNumos As Long, ByVal wNumGRU As String, ByVal wRenavam As String, ByVal wPlaca As String, ByVal wChassi As String, ByVal WMarcaVeiculo As String, ByVal wModeloVeiculo As String, ByVal wAnoVeiculo As String, ByVal wAroVeiculo As String, ByVal wPneuVeiculo As String, ByVal wDataEnsaio As Date) As Boolean

On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim tb As Recordset
Dim pCNPJ As String
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set tb = bd.OpenRecordset("EMPXFABRIC_PISTA", DB_OPEN_TABLE)

tb.index = "CNPJ_Fabric"

If tb.RecordCount = 0 Then
   MsgBox "Não Existem Fabricantes de Simuladores De Ensaio Cadastrados", vbCritical
   Integracao_SGA_Inmetro_Informar_PreEnsaio = False
   GoTo Finaliza
End If

pCNPJ = "00000000000191"
While Not tb.EOF
   If pCNPJ = "00000000000191" Then
      If tb!CodEmpresa = Empresa_Ativa() Then
         If Not IsNull(tb!CNPJ_Fabric) Then
            pCNPJ = tb!CNPJ_Fabric
            Call Integracao_SGA_Inmetro_Informar_PreEnsaio_Sub(WNumos, UCase(wNumGRU), wRenavam, wPlaca, wChassi, WMarcaVeiculo, wModeloVeiculo, wAnoVeiculo, wAroVeiculo, wPneuVeiculo, wDataEnsaio, pCNPJ, IIf(IsNull(tb!Razao_Social), "INDEFINIDO", tb!Razao_Social))
         End If
      End If
   Else
      If tb!CodEmpresa = Empresa_Ativa() Then
         If Not IsNull(tb!CNPJ_Fabric) Then
            If pCNPJ <> tb!CNPJ_Fabric Then
               pCNPJ = tb!CNPJ_Fabric
               Call Integracao_SGA_Inmetro_Informar_PreEnsaio_Sub(WNumos, UCase(wNumGRU), wRenavam, wPlaca, wChassi, WMarcaVeiculo, wModeloVeiculo, wAnoVeiculo, wAroVeiculo, wPneuVeiculo, wDataEnsaio, pCNPJ, IIf(IsNull(tb!Razao_Social), "INDEFINIDO", tb!Razao_Social))
            End If
         End If
      End If
   End If
   tb.MoveNext
Wend

Finaliza:
   tb.Close
   bd.Close
   Set tb = Nothing
   Set bd = Nothing
   Exit Function
Return

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informações eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de número " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Public Function Integracao_SGA_Inmetro_Informar_PreEnsaio_Sub(ByVal WNumos As Long, ByVal wNumGRU As String, ByVal wRenavam As String, ByVal wPlaca As String, ByVal wChassi As String, ByVal WMarcaVeiculo As String, ByVal wModeloVeiculo As String, ByVal wAnoVeiculo As String, ByVal wAroVeiculo As String, ByVal wPneuVeiculo As String, ByVal wDataEnsaio As Date, ByVal wCNPJFabricante As String, ByVal wMarcaSimulador As String) As Boolean

Dim oconnreq As CanadaWebService.RequestEnviarPreEnsaio
Dim oConnResp As CanadaWebService.LoadResponseEnviarPreEnsaio
Dim oService As CanadaWebService.EnviarPreEnsaio

If UCase(TipoPosto) = "S" Then
    Add_Message_Status_Operacao "OPERACAO INDISPONIVEL PARA EMPRESAS DO TIPO POSTO DE SELAGEM, IMPOSSIVEL PROSSEGUIR."
    Integracao_SGA_Inmetro_Informar_PreEnsaio_Sub = False
    Exit Function
End If

Set ofuncoes = New CanadaConnectivity.FuncoesDiversas
Set oconnreq = New CanadaWebService.RequestEnviarPreEnsaio

Set oconnreq.oFabricante = New CanadaWebService.Fabricante
'oconnreq.oFabricante.cNPJ = "00000000000191"
'oconnreq.oFabricante.cNPJ = "25069711000143" ' MACHINE
'oconnreq.oFabricante.cNPJ = "10206534000122" ' ETM
oconnreq.oFabricante.CNPJ = wCNPJFabricante

Set oconnreq.oPAC = New CanadaWebService.PAC
'oconnreq.oPAC.CGC = CNPJ()
oconnreq.oPAC.CGC = RetornaLoginInMetro()
'oconnreq.oPAC.CGC = "22202378000100"
'oconnreq.oPAC.CodigoIdent = "71503"
oconnreq.oPAC.CodigoIdent = NumeroSerie(Empresa_Ativa)

Set oconnreq.oGRU = New CanadaWebService.GRU
oconnreq.oGRU.NumeroGRU = UCase(wNumGRU)

Set oconnreq.oEnsaio = New CanadaWebService.Ensaio
oconnreq.oEnsaio.DataEnsaio = Format(wDataEnsaio, "yyyy-MM-dd")
oconnreq.oDataHoraEnvio = Now()

Set oconnreq.oVeiculo = New CanadaWebService.Veiculo
oconnreq.oVeiculo.PlacaVeiculo = wPlaca
oconnreq.oVeiculo.RenavamVeiculo = wRenavam
oconnreq.oVeiculo.ChassiVeiculo = wChassi
oconnreq.oVeiculo.MarcaVeiculo = WMarcaVeiculo
oconnreq.oVeiculo.ModeloVeiculo = wModeloVeiculo
oconnreq.oVeiculo.AnoVeiculo = wAnoVeiculo
oconnreq.oVeiculo.PneuVeiculo = wPneuVeiculo
oconnreq.oVeiculo.AroVeiculo = wAroVeiculo
oconnreq.CreateXml WNumos, Environ("Computername"), CurrentUser()

Set oService = New CanadaWebService.EnviarPreEnsaio
oService.Send (oconnreq.StringXml)

Set oConnResp = New CanadaWebService.LoadResponseEnviarPreEnsaio
oConnResp.Load_Objetos (oService.MessageReturn)

If oConnResp.Status = True Then
    MsgBox "Pre Ensaio Registrado No Simulador " & wMarcaSimulador & ".", vbInformation
    Call Add_PreEnsaioXOSV(WNumos, wDataEnsaio, oconnreq.oDataHoraEnvio, wPlaca, wRenavam, wChassi, WMarcaVeiculo, wModeloVeiculo, wAnoVeiculo, wPneuVeiculo, wAroVeiculo, Environ("Computername"), Captura_Conteudo_Campo_RESP(Usuario_Ativo(), "nome"), UCase(wNumGRU))
   'MsgBox oConnResp.Message, vbInformation
Else
   MsgBox oConnResp.Message, vbCritical
End If

Integracao_SGA_Inmetro_Informar_PreEnsaio_Sub = oConnResp.Status

Set oConnResp = Nothing
Set oConnResp = Nothing
Set ofuncoes = Nothing

Exit Function

TratamentodeErro:
    If Err.Number = 3001 Then
      MsgBox "Ocorreu um erro quando suas informações eram processadas, para mante-las atualizadas, reinicie o SISTEMA DE GERENCIAMENTO ", 16, "SGa"
    Else
       MsgBox "Erro de número " & Err.Number & ": " & Err.Description
       Resume Next
    End If

End Function

Private Sub Loading_Page()
Dim Ret As Variant
Ret = SysCmd(SYSCMD_SETSTATUS, "Aguarde, processando...")
End Sub
Private Sub Finalize_Page()
Dim Ret As Variant
Ret = SysCmd(SYSCMD_SETSTATUS, "Processo Finalizado...")
End Sub




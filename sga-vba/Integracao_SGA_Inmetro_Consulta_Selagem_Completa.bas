Attribute VB_Name = "Integracao_SGA_Inmetro_Consulta_Selagem_Completa"
Option Compare Database
Option Explicit
#If VBA7 And Win64 Then
     Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
 ' Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If


Private Table_Resultado As Recordset
Private Table_StatusOperaca As Recordset
Private db As DataBase

Private oconnreq As CanadaConnectivity.ReqConsultaSelagem
Private oConnResp As CanadaConnectivity.LoadRespConsultaSelagem
Private oConnReqGru As CanadaConnectivity.ReqConsultarGRU
Private oConnRespGru As CanadaConnectivity.LoadRespConsultarGRU
Private ofuncoes As CanadaConnectivity.FuncoesDiversas

Public Sub InicializaVariaveisConectividade()
Set ofuncoes = New CanadaConnectivity.FuncoesDiversas
If IsCloud = False Then
   pCaminhoSGa = CaminhoSGaLocal
   pCaminhoReq = CaminhoReqLocal
   pCaminhoResp = CaminhoRespLocal
   pNameUserCloud = ofuncoes.ReturnUserWindows
   pFileNameReq = pCaminhoReq & "\" & ofuncoes.ReturnUserWindows & ".xml"
   pFileNameResp = pCaminhoResp & "\" & ofuncoes.ReturnUserWindows & ".xml"
Else
   pCaminhoSGa = CaminhoSGaRemote
   pCaminhoReq = CaminhoReqRemote
   pCaminhoResp = CaminhoRespRemote
   pNameUserCloud = "SGaCloud"
   pFileNameReq = pCaminhoReq & "\" & pNameUserCloud & ".xml"   '<<<<-------  AQUI   <<<<<<<<<<<
   pFileNameResp = pCaminhoResp & "\" & pNameUserCloud & ".xml"
End If

End Sub
Public Function Integracao_SGA_Inmetro_InicializaConectividadeCloud(ByVal wRenavam As String, ByVal wPlaca As String, Optional wChassi As String, Optional ByVal WNumos As Long = 0) As Boolean

Set db = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set Table_Resultado = db.OpenRecordset("Consulta_Selagem_Completa", DB_OPEN_TABLE)
Set Table_StatusOperaca = db.OpenRecordset("Status_Operacao", DB_OPEN_TABLE)
Set oconnreq = New CanadaConnectivity.ReqConsultaSelagem
Set ofuncoes = New CanadaConnectivity.FuncoesDiversas

InicializaVariaveisConectividade

If ofuncoes.ExisteDiretorio(CaminhoSGaRemote) = False Then
   MsgBox "CONEXAO OFF LINE, IMPOSSIVEL PROSSEGUIR.", vbCritical, "SGa"
   Integracao_SGA_Inmetro_InicializaConectividadeCloud = False
   Exit Function
End If
ofuncoes.ExcluiArquivos pFileNameReq
ofuncoes.ExcluiArquivos pFileNameResp

oconnreq.CreateConsultaSelagem wPlaca, wRenavam, wChassi, False, Retorna_Links_Uteis(6), Retorna_Links_Uteis(10), 6, "www.canainfo.com.br"

Table_Resultado.Close
Table_StatusOperaca.Close
db.Close
Set Table_StatusOperaca = Nothing
Set Table_Resultado = Nothing
Set db = Nothing

Integracao_SGA_Inmetro_InicializaConectividadeCloud = True

Set ofuncoes = Nothing
Set oconnreq = Nothing
Set oConnResp = Nothing

End Function

Public Function Integracao_SGA_Inmetro_Consulta_Selagem_Detalhada(ByVal wRenavam As String, ByVal wPlaca As String, Optional wChassi As String, Optional ByVal WNumos As Long = 0) As Boolean

Set db = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set Table_Resultado = db.OpenRecordset("Consulta_Selagem_Completa", DB_OPEN_TABLE)
Set Table_StatusOperaca = db.OpenRecordset("Status_Operacao", DB_OPEN_TABLE)
Set oconnreq = New CanadaConnectivity.ReqConsultaSelagem
Set ofuncoes = New CanadaConnectivity.FuncoesDiversas

InicializaVariaveisConectividade

If IsCloud = False Then
'[CONECTIVIDADE_NOVA] validacao da Conectividade antiga desativada:
   'If ofuncoes.ConectividadeIsAlive("Conectividade") = False Then
      'MsgBox "CONECTIVIDADE NAO ESTA ATIVA, IMPOSSIVEL PROSSEGUIR.", vbCritical
      'Integracao_SGA_Inmetro_Consulta_Selagem_Detalhada = False
      'Exit Function
   'End If
Else
   If ofuncoes.ExisteDiretorio(CaminhoSGaRemote) = False Then
      MsgBox "CONEXAO OFF LINE, IMPOSSIVEL PROSSEGUIR.", vbCritical
      Integracao_SGA_Inmetro_Consulta_Selagem_Detalhada = False
      Exit Function
   End If
End If
ofuncoes.ExcluiArquivos pFileNameReq
ofuncoes.ExcluiArquivos pFileNameResp
Call Inicializa_Placa_Veic_Cons_Tabela_Completa(wPlaca)
oconnreq.CreateConsultaSelagem wPlaca, wRenavam, wChassi, False, Retorna_Links_Uteis(6), Retorna_Links_Uteis(10), 6, "www.canainfo.com.br"

If IsCloud = False Then
'[CONECTIVIDADE_NOVA] ativacao da janela do app antigo desativada:
    'AppActivate ("Conectividade - " & RetornaLoginInMetro)
End If

Do While ofuncoes.CheckFile(pFileNameResp) = False
   Sleep (500)
   DoEvents
Loop

'Sleep (2000)
Set oConnResp = New CanadaConnectivity.LoadRespConsultaSelagem
oConnResp.Load_Objetos pNameUserCloud, IsCloud
AppActivate "SGa"

If CBool(oConnResp.oSTATUS) = True Then

   ofuncoes.ExcluiArquivos CaminhoReqLocal & "\" & ofuncoes.ReturnUserWindows & ".xml"
   ofuncoes.ExcluiArquivos CaminhoRespLocal & "\" & ofuncoes.ReturnUserWindows & ".xml"
   
   ofuncoes.ExcluiArquivos pFileNameReq
   ofuncoes.ExcluiArquivos pFileNameResp
   
   Table_Resultado.AddNew
   Load_Data_Veiculos
   Load_Data_Cronotacografo
   Load_Data_Selagem
   Load_Data_Ensaio
   Load_Data_Resultado
   
   Table_Resultado.Fields("Gru") = oConnResp.oGRU.NumGru
   Table_Resultado.Fields("NumOS") = WNumos
   Table_Resultado.Fields("DataHoraConsulta") = Now
   Table_Resultado.Update
   
Else
    Add_Message_Status_Operacao oConnResp.oMessageError
End If

AppActivate "SGa"
Table_Resultado.Close
Table_StatusOperaca.Close
db.Close
Set Table_StatusOperaca = Nothing
Set Table_Resultado = Nothing
Set db = Nothing

ofuncoes.ExcluiArquivos CaminhoReqLocal & "\" & ofuncoes.ReturnUserWindows & ".xml"
ofuncoes.ExcluiArquivos CaminhoRespLocal & "\" & ofuncoes.ReturnUserWindows & ".xml"
 
ofuncoes.ExcluiArquivos pFileNameReq
ofuncoes.ExcluiArquivos pFileNameResp

Integracao_SGA_Inmetro_Consulta_Selagem_Detalhada = CBool(oConnResp.oSTATUS)

Set ofuncoes = Nothing
Set oconnreq = Nothing
Set oConnResp = Nothing

End Function
Private Sub Load_Data_Gru_Div()
Table_Resultado.Fields("Gru").Value = IIf(oConnRespGru Is Nothing, " ", oConnRespGru.oGRU.NumGru)
Table_Resultado.Fields("ValorGru").Value = IIf(oConnRespGru Is Nothing, " ", oConnRespGru.oGRU.ValorGru)
Table_Resultado.Fields("ValorPago").Value = IIf(oConnRespGru Is Nothing, " ", oConnRespGru.oGRU.ValorPag)
Table_Resultado.Fields("DataPagamento").Value = IIf(oConnRespGru Is Nothing, " ", oConnRespGru.oGRU.DataPag)
End Sub
Private Sub Load_Data_Gru_Table()
Table_Resultado.Fields("Gru").Value = IIf(oConnRespGru Is Nothing, " ", oConnRespGru.oGRU.NumGru)
Table_Resultado.Fields("ValorGru").Value = IIf(oConnRespGru Is Nothing, " ", oConnRespGru.oGRU.ValorGru)
Table_Resultado.Fields("ValorPago").Value = IIf(oConnRespGru Is Nothing, " ", oConnRespGru.oGRU.ValorPag)
Table_Resultado.Fields("DataPagamento").Value = IIf(oConnRespGru Is Nothing, " ", oConnRespGru.oGRU.DataPag)
End Sub
Private Sub Load_Data_Veiculos()
Table_Resultado.Fields("CHASSI").Value = oConnResp.oVeiculo.CHASSI
Table_Resultado.Fields("Placa").Value = oConnResp.oVeiculo.Placa
Table_Resultado.Fields("MarcaVeiculo").Value = oConnResp.oVeiculo.Marca
Table_Resultado.Fields("ModeloVeiculo").Value = oConnResp.oVeiculo.Modelo
Table_Resultado.Fields("AnoVeiculo").Value = oConnResp.oVeiculo.Ano
Table_Resultado.Fields("ARO").Value = oConnResp.oVeiculo.Aro
Table_Resultado.Fields("Pneu").Value = oConnResp.oVeiculo.Pneu
Table_Resultado.Fields("Renavam").Value = oConnResp.oVeiculo.Renavam
End Sub
Private Sub Load_Data_Cronotacografo()
Table_Resultado.Fields("MarcaEquip").Value = oConnResp.oEquipamento.Marca
Table_Resultado.Fields("ModeloEquip").Value = oConnResp.oEquipamento.Modelo
Table_Resultado.Fields("NroSerieEquip").Value = oConnResp.oEquipamento.Nroserie
Table_Resultado.Fields("ConstK").Value = oConnResp.oEquipamento.ConstK
Table_Resultado.Fields("Redutor").Value = oConnResp.oEquipamento.Redutor
End Sub
Private Sub Load_Data_Selagem()
Table_Resultado.Fields("DataSelagem").Value = oConnResp.oSelagem.DataSelagem
Table_Resultado.Fields("SelosPlasticos").Value = oConnResp.oSelagem.SelosPlasticos
Table_Resultado.Fields("SelosAdesivos").Value = oConnResp.oSelagem.SelosAdesivos
Table_Resultado.Fields("PostoSelagem").Value = oConnResp.oSelagem.PostoSelagem
Table_Resultado.Fields("NumeroProtocoloSelagem").Value = oConnResp.oSelagem.Protocolo
Table_Resultado.Fields("SituacaoSelagem").Value = oConnResp.oSelagem.Situacao
End Sub
Private Sub Load_Data_Ensaio()
Table_Resultado.Fields("UFEnsaio").Value = oConnResp.oEnsaio.UFEnsaio
Table_Resultado.Fields("PostoEnsaio").Value = oConnResp.oEnsaio.PostoEnsaio
Table_Resultado.Fields("DataEnsaio").Value = oConnResp.oEnsaio.DataEnsaio
Table_Resultado.Fields("VelMaxEnsaio").Value = oConnResp.oEnsaio.VelMaxEnsaio
Table_Resultado.Fields("PadraoUtilizadoEnsaio").Value = oConnResp.oEnsaio.PadraoUtilizado
End Sub

Private Sub Load_Data_Certificado()

'Table_Resultado.Fields("NumeroCertificado2").Value = IIf(oconnResp.oCertificado= "", ".", oNroDocto2)
'Table_Resultado.Fields("DataCertificado2").Value = IIf(oDataDocto2 = "", ".", oDataDocto2)
'Table_Resultado.Fields("ValidadeCertificado2").Value = IIf(oValidadeDocto2 = "", ".", oValidadeDocto2)
'Table_Resultado.Fields("TipoCertificado2").Value = IIf(oTipo2 = "", ".", oTipo2)

End Sub
Private Sub Load_Data_Leitura()
'Table_Resultado.Fields("EstadoResultado").Value = IIf(IsNull(oUFLeitura) Or oUFLeitura = "", ".", oUFLeitura)
'Table_Resultado.Fields("ResponsavelResultado").Value = IIf(IsNull(oRespTecnico) Or oRespTecnico = "", ".", oRespTecnico)
'Table_Resultado.Fields("DataResultado").Value = IIf(IsNull(oDataLeitura) Or oDataLeitura = "", "01/01/1964", oDataLeitura)
'Table_Resultado.Fields("NaoConformidade").Value = IIf(oConformidades = "", ".", oConformidades)
End Sub
Private Sub Load_Data_Resultado()

   If IsDate(oConnResp.oResultado.DataValidade) = True Then
       'Table_Consulta.Fields("EstadoResultado").Value = oConnResp.oLotePlacas.EstadoResultado(cont)
       'Table_Consulta.Fields("ResponsavelResultado").Value = oConnResp.oLotePlacas.ResponsavelResultado(cont)
       'Table_Consulta.Fields("DataResultado").Value = Format(IIf(Not IsDate(oConnResp.oLotePlacas.DataResultado(cont)), "01/01/2001", oConnResp.oLotePlacas.DataResultado(cont)), "dd/MM/yyyy")
       Table_Resultado.Fields("DataValidadeResultado").Value = Format(oConnResp.oResultado.DataValidade, "dd/MM/yyyy")
       Table_Resultado.Fields("TipoCertificadoResultado").Value = oConnResp.oResultado.TipoCertificado
     '  Table_Consulta.Fields("NumeroCertificadoDefinitivo").Value = IIf(IsNull(oConnResp.oLotePlacas.NumeroCertificadoDefinitivo(cont)) Or oConnResp.oLotePlacas.NumeroCertificadoDefinitivo(cont) = "", "000000000", oConnResp.oLotePlacas.NumeroCertificadoDefinitivo(cont))
  ' ElseIf oConnResp.oLotePlacas.TipoCertificado2(cont) <> "" And Not IsNull(oConnResp.oLotePlacas.TipoCertificado2(cont)) Then
  '     Table_Consulta.Fields("tipocertificadoresultado").Value = oConnResp.oLotePlacas.TipoCertificado2(cont)
   End If




'Table_Resultado.Fields("EstadoResultado").Value = oUFResultado
'Table_Resultado.Fields("ResponsavelResultado").Value = oRespTecnico
'Table_Resultado.Fields("DataResultado").Value = oDataEmissao
'Table_Resultado.Fields("DataValidadeResultado").Value = oDataValidade
'Table_Resultado.Fields("NumeroCertificadoDefinitivo").Value = oNumCertificado
End Sub


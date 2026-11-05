Attribute VB_Name = "Integracao_SGA_Inmetro_Informar_Substituicao"
Option Compare Database
Option Explicit


Private Table_SELOS_APLICADOS As Recordset
Private db1 As DataBase
Private Status As Boolean
Private oconnreq As CanadaConnectivity.ReqInformarSubstituicao
Private ofuncoes As CanadaConnectivity.FuncoesDiversas
Private oConnResp As CanadaConnectivity.LoadRespInformarSubstituicao

Public Const CaminhoSGaRemote As String = "\\tsclient\c\sga"
Public Const CaminhoReqRemote As String = "\\tsclient\c\sga\Req"
Public Const CaminhoRespRemote As String = "\\tsclient\c\sga\Resp"
Public Const CaminhoWhatsAppRemote As String = "\\tsclient\c\sga\WhatsApp"

Public Const CaminhoSGaLocal As String = "C:\Sga"
Public Const CaminhoReqLocal As String = "C:\Sga\Req"
Public Const CaminhoRespLocal As String = "C:\Sga\Resp"

Public pNameUserCloud As String

Public pCaminhoSGa As String
Public pCaminhoReq As String
Public pCaminhoResp As String

Public pFileNameReq As String
Public pFileNameResp As String

Public Function IsCloud() As Boolean
If Captura_Conteudo_Campos_Empresa(1, "Tipo_Base_Dados") = 1 Then
   IsCloud = True
Else
   IsCloud = False
End If
End Function

Public Function Integracao_SGA_Inmetro_Informar_Substituicao(ByVal wNumOSV As Long, ByVal wNumGRU As String, ByVal wRenavam As String, ByVal wPlaca As String, wNumeroProtocolo As String) As Boolean

Set ofuncoes = New CanadaConnectivity.FuncoesDiversas

InicializaVariaveisConectividade

If IsCloud = False Then
'[CONECTIVIDADE_NOVA] validacao da Conectividade antiga desativada:
   'If ofuncoes.ConectividadeIsAlive("Conectividade") = False Then
      'MsgBox "CONECTIVIDADE NAO ESTA ATIVA, IMPOSSIVEL PROSSEGUIR.", vbCritical
      'Integracao_SGA_Inmetro_Informar_Substituicao = False
      'Exit Function
   'End If
Else
   If ofuncoes.ExisteDiretorio(CaminhoSGaRemote) = False Then
      MsgBox "CONEXAO OFF LINE, IMPOSSIVEL PROSSEGUIR.", vbCritical
      Integracao_SGA_Inmetro_Informar_Substituicao = False
      Exit Function
   End If
End If
ofuncoes.ExcluiArquivos pFileNameReq
ofuncoes.ExcluiArquivos pFileNameResp

Set db1 = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Table_SELOS_APLICADOS = db1.OpenRecordset("SELOSAPLICADOS", DB_OPEN_TABLE)

Clear_Table_Status_Operacao

Set oconnreq.oSubstituicao = New CanadaConnectivity.Substituicao
Set oconnreq.oSubstituicao.oSelagem = New CanadaConnectivity.Selagem
Set oconnreq.oSubstituicao.oVeiculo = New CanadaConnectivity.Veiculo

oconnreq.oSubstituicao.oVeiculo.Placa = wPlaca
oconnreq.oSubstituicao.oVeiculo.Renavam = wRenavam

oconnreq.oSubstituicao.oSelagem.Protocolo = wNumeroProtocolo
oconnreq.oSubstituicao.oSelagem.GRUSELAGEM = wNumGRU

InputSelosSubstituidos wNumOSV, "A"
InputSelosSubstituidos wNumOSV, "P"

oconnreq.CreateInformarSubstituicao Retorna_Links_Uteis(8), Retorna_Links_Uteis(7), 8, "www.canainfo.com.br"

'sleep (2000)
If IsCloud = False Then
'[CONECTIVIDADE_NOVA] ativacao da janela do app antigo desativada:
    'AppActivate ("Conectividade - " & RetornaLoginInMetro)
End If

Do While ofuncoes.CheckFile(pFileNameResp) = False
Loop

'sleep (2000)

oConnResp.Load_Objetos pNameUserCloud, IsCloud

Add_Message_Status_Operacao oConnResp.Message

Integracao_SGA_Inmetro_Informar_Substituicao = oConnResp.Status

ofuncoes.ExcluiArquivos CaminhoReqLocal & "\" & ofuncoes.ReturnUserWindows & ".xml"
ofuncoes.ExcluiArquivos CaminhoRespLocal & "\" & ofuncoes.ReturnUserWindows & ".xml"
   
ofuncoes.ExcluiArquivos pFileNameReq
ofuncoes.ExcluiArquivos pFileNameResp

Table_SELOS_APLICADOS.Close
db1.Close
Set ofuncoes = Nothing
Set oconnreq = Nothing
Set oConnResp = Nothing
End Function

Private Sub InputSelosSubstituidos(ByVal wNumOSV As Long, ByVal wTpSelo As String)
Table_SELOS_APLICADOS.index = "QuintaChave"
Table_SELOS_APLICADOS.Seek "=", wNumOSV
Dim WAux As String
WAux = ""
If Not Table_SELOS_APLICADOS.NoMatch Then
    While Not Table_SELOS_APLICADOS.EOF
        If Table_SELOS_APLICADOS.Fields("NumOS") = wNumOSV Then
            If UCase(Table_SELOS_APLICADOS.Fields("TipoSelo")) = wTpSelo Then
                 If IsNumeric(Table_SELOS_APLICADOS.Fields("ModeloSelo")) Then
                    WAux = WAux & Retorna_Value_Modelo_Selos(Table_SELOS_APLICADOS.Fields("ModeloSelo")) & Mid$(Table_SELOS_APLICADOS.Fields("Numero"), 1, 8) & Table_SELOS_APLICADOS.Fields("DV")
                 Else
                    WAux = WAux & Mid$(Table_SELOS_APLICADOS.Fields("Numero"), 1, 8) & Table_SELOS_APLICADOS.Fields("DV")
                 End If
                      
                 If wTpSelo = "A" Then oconnreq.oSubstituicao.Add_Selo_Adesivo WAux, Table_SELOS_APLICADOS.Fields("NumeroSubst").Value & Table_SELOS_APLICADOS.Fields("DVSubst").Value
                 If wTpSelo = "P" Then oconnreq.oSubstituicao.Add_Selo_Plastico WAux, Table_SELOS_APLICADOS.Fields("NumeroSubst").Value & Table_SELOS_APLICADOS.Fields("DVSubst").Value
                 WAux = ""
                 Table_SELOS_APLICADOS.MoveNext
            Else
                 Table_SELOS_APLICADOS.MoveNext
            End If
        Else
            Exit Sub
        End If
    Wend
Else
    Status = False
    Add_Message_Status_Operacao (UCase("A Ordem de Servicos Nao Possui Selos Aplicados. Impossivel Prosseguir."))
End If
End Sub


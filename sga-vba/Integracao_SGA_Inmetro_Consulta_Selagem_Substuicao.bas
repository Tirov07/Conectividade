Attribute VB_Name = "Integracao_SGA_Inmetro_Consulta_Selagem_Substuicao"
Option Compare Database
Option Explicit

Private pNumOS As Long
Private Table_Consulta As Recordset
Private db As DataBase
Private Status As Boolean
Private oconnreq As CanadaConnectivity.ReqConsultaSelagem
Private oConnResp As CanadaConnectivity.LoadRespConsultaSelagem
Private ofuncoes As CanadaConnectivity.FuncoesDiversas
Dim oSelos_Adesivos As Variant
Public Function Integracao_SGA_Inmetro_Consulta_Selagem_Substituicao(ByVal wRenavam As String, ByVal wPlaca As String, ByVal WNumos As Long) As Boolean
Set db = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set Table_Consulta = db.OpenRecordset("SELOSAPLICADOSTEMP", DB_OPEN_TABLE)
Set oconnreq = New CanadaConnectivity.ReqConsultaSelagem
Set ofuncoes = New CanadaConnectivity.FuncoesDiversas

InicializaVariaveisConectividade

If IsCloud = False Then
'[CONECTIVIDADE_NOVA] validacao da Conectividade antiga desativada:
   'If ofuncoes.ConectividadeIsAlive("Conectividade") = False Then
      'MsgBox "CONECTIVIDADE NAO ESTA ATIVA, IMPOSSIVEL PROSSEGUIR.", vbCritical
      'Integracao_SGA_Inmetro_Consulta_Selagem_Substituicao = False
      'Exit Function
   'End If
Else
   If ofuncoes.ExisteDiretorio(CaminhoSGaRemote) = False Then
      MsgBox "CONEXAO OFF LINE, IMPOSSIVEL PROSSEGUIR.", vbCritical
      Integracao_SGA_Inmetro_Consulta_Selagem_Substituicao = False
      Exit Function
   End If
End If
ofuncoes.ExcluiArquivos pFileNameReq
ofuncoes.ExcluiArquivos pFileNameResp

oconnreq.CreateConsultaSelagem wPlaca, wRenavam, "", False, Retorna_Links_Uteis(6), Retorna_Links_Uteis(7), 6, "www.canainfo.com.br"
If IsCloud = False Then
'[CONECTIVIDADE_NOVA] ativacao da janela do app antigo desativada:
    'AppActivate ("Conectividade - " & RetornaLoginInMetro)
End If

Do While ofuncoes.CheckFile(pFileNameResp) = False
Loop

'sleep (2000)

Set oConnResp = New CanadaConnectivity.LoadRespConsultaSelagem
oConnResp.Load_Objetos pNameUserCloud, IsCloud
AppActivate "SGa"

If CBool(oConnResp.oSTATUS) = True Then
    pNumOS = WNumos
    Add_table_Substituicao_Adesivos
    Add_table_Substituicao_Plasticos
        
Else
    Add_Message_Status_Operacao oConnResp.oMessageError
End If

Table_Consulta.Close
db.Close
Set Table_Consulta = Nothing
Set db = Nothing

'ofuncoes.ExcluiArquivos "C:\SGA\REQ\" & ofuncoes.ReturnUserWindows & ".xml"
'ofuncoes.ExcluiArquivos "C:\SGA\RESP\" & ofuncoes.ReturnUserWindows & ".xml"

ofuncoes.ExcluiArquivos CaminhoReqLocal & "\" & ofuncoes.ReturnUserWindows & ".xml"
ofuncoes.ExcluiArquivos CaminhoRespLocal & "\" & ofuncoes.ReturnUserWindows & ".xml"
   
ofuncoes.ExcluiArquivos pFileNameReq
ofuncoes.ExcluiArquivos pFileNameResp

'[CONECTIVIDADE_NOVA] ativacao da janela do app antigo desativada:
'AppActivate ("Conectividade - " & RetornaLoginInMetro)
'AppActivate "SGa"
'MsgBox "Consulta Finalizada!", vbInformation
Integracao_SGA_Inmetro_Consulta_Selagem_Substituicao = oConnResp.oSTATUS
Set ofuncoes = Nothing
Set oconnreq = Nothing
Set oConnResp = Nothing
End Function

Private Sub Add_table_Substituicao_Adesivos()
Dim ocont As Integer
If oSelos_Adesivos <> "-" Then
    For ocont = 0 To oConnResp.oSelagem.CountSelosAdesivos - 1
        Table_Consulta.AddNew
        Table_Consulta.Fields("Numero").Value = oConnResp.oSelagem.SeloAdesivo_Numero(ocont)
        Table_Consulta.Fields("NumOS").Value = pNumOS
        Table_Consulta.Fields("TipoSelo").Value = "A"
        Table_Consulta.Fields("DV").Value = oConnResp.oSelagem.SeloAdesivo_DV(ocont)
        Table_Consulta.Update
    Next ocont
End If
End Sub
Private Sub Add_table_Substituicao_Plasticos()
Dim ocont As Integer
For ocont = 0 To oConnResp.oSelagem.CountSelosPlasticos - 1
    Table_Consulta.AddNew
    Table_Consulta.Fields("Numero").Value = oConnResp.oSelagem.SeloPlastico_Numero(ocont)
    Table_Consulta.Fields("NumOS").Value = pNumOS
    Table_Consulta.Fields("TipoSelo").Value = "P"
    Table_Consulta.Fields("DV").Value = oConnResp.oSelagem.SeloPlastico_DV(ocont)
    Table_Consulta.Update
Next ocont
End Sub


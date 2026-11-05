Attribute VB_Name = "Integracao_SGA_Inmetro_Consulta_Selagem_Completa_Lote_Placas_NEw"
Option Compare Database
Option Explicit
#If VBA7 And Win64 Then
     Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
 ' Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private ologin, osenha As String
Private Table_Consulta, Table_Aux As Recordset
Private bd, bdc As DataBase
Private Status As Boolean
Private oconnreq As CanadaConnectivity.ReqConsultaLotePlacas
Private oConnResp As CanadaConnectivity.LoadRespConsultaLotePlacas
Private ofuncoes As CanadaConnectivity.FuncoesDiversas
Private cont As Integer
Dim Ret As Variant

Public Function Integracao_SGA_Inmetro_Consulta_Selagem_Detalhada_Lote_New(ByVal wLogin As String, ByVal wsenha As String) As Boolean
Set bd = DBEngine.Workspaces(0).OpenDatabase(Banco1())
Set Table_Consulta = bd.OpenRecordset("PLACXCLI", DB_OPEN_TABLE)
Set bdc = DBEngine.Workspaces(0).OpenDatabase(Banco_Temp())
Set Table_Aux = bdc.OpenRecordset("Base_Dados_Lote", DB_OPEN_TABLE)
ologin = wLogin
osenha = wsenha
Status = True
Make_Consulta

Table_Aux.Close
bdc.Close
Table_Consulta.Close
bd.Close
Set Table_Aux = Nothing
Set Table_Consulta = Nothing
Set bdc = Nothing
Set bd = Nothing
Set oconnreq = Nothing
Set oConnResp = Nothing
Set ofuncoes = Nothing
AppActivate "SGa"

MsgBox "Consulta Finalizada.", vbInformation
Integracao_SGA_Inmetro_Consulta_Selagem_Detalhada_Lote_New = Status
End Function
Private Sub Make_Consulta()
Dim contPlacas As Integer
Dim oLote As CanadaConnectivity.LotePlacas
contPlacas = 0
Set ofuncoes = New CanadaConnectivity.FuncoesDiversas

InicializaVariaveisConectividade

If IsCloud = False Then
'[CONECTIVIDADE_NOVA] validacao da Conectividade antiga desativada:
   'If ofuncoes.ConectividadeIsAlive("Conectividade") = False Then
      'MsgBox "CONECTIVIDADE NAO ESTA ATIVA, IMPOSSIVEL PROSSEGUIR.", vbCritical
      'Status = False
      'Exit Sub
   'End If
Else
   If ofuncoes.ExisteDiretorio(CaminhoSGaRemote) = False Then
      MsgBox "CONEXAO OFF LINE, IMPOSSIVEL PROSSEGUIR.", vbCritical
      Status = False
      Exit Sub
   End If
End If
ofuncoes.ExcluiArquivos pFileNameReq
ofuncoes.ExcluiArquivos pFileNameResp
    If Table_Aux.RecordCount <> 0 Then
            Set oLote = New CanadaConnectivity.LotePlacas
            Set oconnreq = New CanadaConnectivity.ReqConsultaLotePlacas
            oLote.InicializaTable
            While Not Table_Aux.EOF And contPlacas <= 300
                Table_Consulta.index = "PrimaryKey"
                Table_Consulta.Seek "=", Table_Aux!Placa, Table_Aux!CodigoCliente
                oLote.AddVeiculo
                contPlacas = contPlacas + 1
                oLote.CodCliente = Table_Aux!CodigoCliente
                oLote.Placa = Table_Aux!Placa
                oLote.Update
                Table_Aux.MoveNext
            Wend
            Ret = SysCmd(SYSCMD_SETSTATUS, "Realizando a Consulta de Lotes de Placas >>>> Por Favor Aguarde.")
            oconnreq.CreateConsultaLotePlacasEnvio oLote, Retorna_Links_Uteis(6), Retorna_Links_Uteis(7), 13, "www.canainfo.com.br"
            
            If IsCloud = False Then
'[CONECTIVIDADE_NOVA] ativacao da janela do app antigo desativada:
               'AppActivate ("Conectividade - " & RetornaLoginInMetro)
            End If
                        
            Do While ofuncoes.CheckFile(pFileNameResp) = False
               Sleep (500)
               DoEvents
            Loop
         '   Sleep (2000)
            
            
            Set oConnResp = New CanadaConnectivity.LoadRespConsultaLotePlacas
            oConnResp.Load_Objetos pNameUserCloud, IsCloud
            AppActivate "SGa"
            
            If oConnResp.Status = False Then
                Exit Sub
            End If
            Table_Consulta.index = "PrimaryKey"
            BeginTrans
            For cont = 0 To oConnResp.oLotePlacas.RowsCount - 1
                Table_Consulta.Seek "=", oConnResp.oLotePlacas.Placa(cont), oConnResp.oLotePlacas.CodCliente(cont)
                If Not Table_Consulta.NoMatch Then
                    Table_Consulta.Edit
                    Load_Table_Veiculo
                    Load_Data_Selagem
                    Load_Data_Ensaio
                    Load_Data_Cronotacografo
                    Load_Data_Certificado
                    Load_Data_Resultado
                    Table_Consulta!DataConsulta = Date
                    Table_Consulta.Update
                End If
            Next
            CommitTrans
    End If
Ret = SysCmd(SYSCMD_SETSTATUS, "Consulta Finalizada.")

ofuncoes.ExcluiArquivos "C:\SGA\REQ\" & ofuncoes.ReturnUserWindows & ".xml"
ofuncoes.ExcluiArquivos "C:\SGA\RESP\" & ofuncoes.ReturnUserWindows & ".xml"
ofuncoes.ExcluiArquivos pFileNameReq
ofuncoes.ExcluiArquivos pFileNameResp

Status = True
End Sub

Private Sub Load_Table_Veiculo_New(ByVal WCodigoCliente As Long)
Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "Placa", oConnResp.oLotePlacas.Placa(cont))
Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "Renavan", oConnResp.oLotePlacas.Renavam(cont))
Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "MarcaVeiculo", oConnResp.oLotePlacas.MarcaVeiculo(cont))
Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "ModeloVeiculo", oConnResp.oLotePlacas.ModeloVeiculo(cont))
Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "AnoVeiculo", oConnResp.oLotePlacas.AnoVeiculo(cont))
Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "AROPNEU", oConnResp.oLotePlacas.AroPneu(cont))
Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "Pneu", oConnResp.oLotePlacas.Pneu(cont))
End Sub
Private Sub Load_Table_Veiculo()
Table_Consulta.Fields("Renavan").Value = IIf(IsNull(oConnResp.oLotePlacas.Renavam(cont)), " ", oConnResp.oLotePlacas.Renavam(cont))
Table_Consulta.Fields("Placa").Value = IIf(IsNull(oConnResp.oLotePlacas.Placa(cont)), " ", oConnResp.oLotePlacas.Placa(cont))
Table_Consulta.Fields("MarcaVeiculo").Value = IIf(IsNull(oConnResp.oLotePlacas.MarcaVeiculo(cont)), ".", oConnResp.oLotePlacas.MarcaVeiculo(cont))
Table_Consulta.Fields("ModeloVeiculo").Value = IIf(IsNull(oConnResp.oLotePlacas.ModeloVeiculo(cont)), ".", oConnResp.oLotePlacas.ModeloVeiculo(cont))
Table_Consulta.Fields("AnoVeiculo").Value = IIf(IsNull(oConnResp.oLotePlacas.AnoVeiculo(cont)), ".", oConnResp.oLotePlacas.AnoVeiculo(cont))
Table_Consulta.Fields("AROPNEU").Value = IIf(IsNull(oConnResp.oLotePlacas.AroPneu(cont)), ".", oConnResp.oLotePlacas.AroPneu(cont))
Table_Consulta.Fields("Pneu").Value = IIf(IsNull(oConnResp.oLotePlacas.Pneu(cont)), ".", oConnResp.oLotePlacas.Pneu(cont))
End Sub
Private Sub Load_Data_Selagem_New(ByVal WCodigoCliente As Long)
Dim odataselagem As String
odataselagem = Format(IIf(Not IsDate(Mid(oConnResp.oLotePlacas.DataSelagem(cont), 1, 10)), "01/01/1964", Mid(oConnResp.oLotePlacas.DataSelagem(cont), 1, 10)), "dd/MM/yyyy")
Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "DataSelagem", CDate(odataselagem))
Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "PostoSelagem", oConnResp.oLotePlacas.PostoSelagem(cont))
Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "DataValidadeResultado", Format(CDate(odataselagem) + 30, "dd/MM/yyyy"))
Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "TipoCertificadoResultado", "CERT. PRELIMINAR")
End Sub
Private Sub Load_Data_Selagem()
Dim odataselagem As String
odataselagem = Format(IIf(Not IsDate(Mid(oConnResp.oLotePlacas.DataSelagem(cont), 1, 10)), "01/01/1964", Mid(oConnResp.oLotePlacas.DataSelagem(cont), 1, 10)), "dd/MM/yyyy")
Table_Consulta.Fields("DataSelagem").Value = CDate(odataselagem)
Table_Consulta.Fields("PostoSelagem").Value = oConnResp.oLotePlacas.PostoSelagem(cont)
Table_Consulta.Fields("DataValidadeResultado").Value = Format(CDate(odataselagem) + 30, "dd/MM/yyyy")
Table_Consulta.Fields("TipoCertificadoResultado").Value = "CERT. PRELIMINAR"
End Sub
Private Sub Load_Data_Ensaio_New(ByVal WCodigoCliente As Long)
Dim oDataEnsaio As String
oDataEnsaio = Format(IIf(Not IsDate(oConnResp.oLotePlacas.DataEnsaio(cont)), "01/01/2001", oConnResp.oLotePlacas.DataEnsaio(cont)), "dd/MM/yyyy")
Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "UFEnsaio", oConnResp.oLotePlacas.UFEnsaio(cont))
Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "PostoEnsaio", IIf(oConnResp.oLotePlacas.PostoEnsaio(cont) = "", ".", oConnResp.oLotePlacas.PostoEnsaio(cont)))
Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "DataEnsaio", CDate(oDataEnsaio))
Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "VelMaxEnsaio", IIf(oConnResp.oLotePlacas.VelMaxEnsaio(cont) = "", "0", oConnResp.oLotePlacas.VelMaxEnsaio(cont)))
Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "PadraoUtilizadoEnsaio", IIf(oConnResp.oLotePlacas.PadraoUtilizadoEnsaio(cont) = "", ".", oConnResp.oLotePlacas.PadraoUtilizadoEnsaio(cont)))
Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "DataValidadeResultado", Format(CDate(oDataEnsaio) + 30, "dd/MM/yyyy"))
Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "TipoCertificadoResultado", "CERT. PRELIMINAR")
End Sub
Private Sub Load_Data_Ensaio()
Dim oDataEnsaio As String
oDataEnsaio = Format(IIf(Not IsDate(oConnResp.oLotePlacas.DataEnsaio(cont)), "01/01/2001", oConnResp.oLotePlacas.DataEnsaio(cont)), "dd/MM/yyyy")
Table_Consulta.Fields("UFEnsaio").Value = IIf(oConnResp.oLotePlacas.UFEnsaio(cont) = "", ".", oConnResp.oLotePlacas.UFEnsaio(cont))
Table_Consulta.Fields("PostoEnsaio").Value = IIf(oConnResp.oLotePlacas.PostoEnsaio(cont) = "", ".", oConnResp.oLotePlacas.PostoEnsaio(cont))
Table_Consulta.Fields("DataEnsaio").Value = CDate(oDataEnsaio)
Table_Consulta.Fields("VelMaxEnsaio").Value = IIf(oConnResp.oLotePlacas.VelMaxEnsaio(cont) = "", "0", oConnResp.oLotePlacas.VelMaxEnsaio(cont))
Table_Consulta.Fields("PadraoUtilizadoEnsaio").Value = IIf(oConnResp.oLotePlacas.PadraoUtilizadoEnsaio(cont) = "", ".", oConnResp.oLotePlacas.PadraoUtilizadoEnsaio(cont))
Table_Consulta.Fields("DataValidadeResultado").Value = Format(CDate(oDataEnsaio) + 30, "dd/MM/yyyy")
Table_Consulta.Fields("TipoCertificadoResultado").Value = "CERT. PRELIMINAR"
End Sub
Private Sub Load_Data_Certificado_New(ByVal WCodigoCliente As Long)
If Not oConnResp.oLotePlacas.TipoCertificado1(cont) = "." Then
   Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "NumeroCertificado2", IIf(oConnResp.oLotePlacas.NumeroCertificado2(cont) = "", ".", oConnResp.oLotePlacas.NumeroCertificado2(cont)))
   Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "DataCertificado2", IIf(oConnResp.oLotePlacas.DataCertificado2(cont) = "", ".", oConnResp.oLotePlacas.DataCertificado2(cont)))
   Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "ValidadeCertificado2", Format(IIf(oConnResp.oLotePlacas.ValidadeCertificado2(cont) = "", ".", oConnResp.oLotePlacas.ValidadeCertificado2(cont)), "dd/MM/yyyy"))
   Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "TipoCertificado2", IIf(oConnResp.oLotePlacas.TipoCertificado2(cont) = "", ".", oConnResp.oLotePlacas.TipoCertificado2(cont)))
    '######### este é o campo que armazena validade do certificado mais recente
    If oConnResp.oLotePlacas.TipoCertificado1(cont) <> "" Then
        Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "TipoCertificadoResultado", oConnResp.oLotePlacas.TipoCertificado1(cont))
    End If
    If oConnResp.oLotePlacas.DataCertificado1(cont) <> "" Then
        Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "DataResultado", oConnResp.oLotePlacas.DataCertificado1(cont))
    End If
    If oConnResp.oLotePlacas.ValidadeCertificado1(cont) <> "" Then
        Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "DataValidadeResultado", Format(oConnResp.oLotePlacas.ValidadeCertificado1(cont), "dd/MM/yyyy"))
    End If
    If oConnResp.oLotePlacas.NumeroCertificado1(cont) <> "" Then
       Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "NumeroCertificadoDefinitivo", oConnResp.oLotePlacas.NumeroCertificado1(cont))
    End If
Else
    
    Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "NumeroCertificado2", ".")
    Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "DataCertificado2", "01/01/2001")
    Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "ValidadeCertificado2", "01/01/2001")
    Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "TipoCertificado2", ".")
    
End If

End Sub



Private Sub Load_Data_Certificado()
If Not oConnResp.oLotePlacas.TipoCertificado1(cont) = "." Then
    Table_Consulta.Fields("NumeroCertificado2").Value = IIf(oConnResp.oLotePlacas.NumeroCertificado2(cont) = "", ".", oConnResp.oLotePlacas.NumeroCertificado2(cont))
    Table_Consulta.Fields("DataCertificado2").Value = IIf(oConnResp.oLotePlacas.DataCertificado2(cont) = "", ".", oConnResp.oLotePlacas.DataCertificado2(cont))
    Table_Consulta.Fields("ValidadeCertificado2").Value = Format(IIf(oConnResp.oLotePlacas.ValidadeCertificado2(cont) = "", ".", oConnResp.oLotePlacas.ValidadeCertificado2(cont)), "dd/MM/yyyy")
    Table_Consulta.Fields("TipoCertificado2").Value = IIf(oConnResp.oLotePlacas.TipoCertificado2(cont) = "", ".", oConnResp.oLotePlacas.TipoCertificado2(cont))
    '######### este é o campo que armazena validade do certificado mais recente
    If oConnResp.oLotePlacas.TipoCertificado1(cont) <> "" Then
        Table_Consulta.Fields("TipoCertificadoResultado").Value = oConnResp.oLotePlacas.TipoCertificado1(cont)
    End If
    If oConnResp.oLotePlacas.DataCertificado1(cont) <> "" Then
        Table_Consulta.Fields("DataResultado").Value = oConnResp.oLotePlacas.DataCertificado1(cont)
    End If
    If oConnResp.oLotePlacas.ValidadeCertificado1(cont) <> "" Then
        Table_Consulta.Fields("DataValidadeResultado").Value = Format(oConnResp.oLotePlacas.ValidadeCertificado1(cont), "dd/MM/yyyy")
    End If
    If oConnResp.oLotePlacas.NumeroCertificado1(cont) <> "" Then
        Table_Consulta.Fields("NumeroCertificadoDefinitivo").Value = oConnResp.oLotePlacas.NumeroCertificado1(cont)
    End If
Else
    Table_Consulta.Fields("NumeroCertificado2").Value = "."
    Table_Consulta.Fields("DataCertificado2").Value = "01/01/2001"
    Table_Consulta.Fields("ValidadeCertificado2").Value = "01/01/2001"
    Table_Consulta.Fields("TipoCertificado2").Value = "."
End If

End Sub
Private Sub Load_Data_Leitura_New(ByVal WCodigoCliente As Long)
If Not oConnResp.oLeitura Is Nothing Then

   Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "EstadoResultado", IIf(IsNull(oConnResp.oLeitura.UFLeitura(cont)) Or oConnResp.oLeitura.UFLeitura(cont) = "", ".", oConnResp.oLeitura.UFLeitura(cont)))
   Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "ResponsavelResultado", IIf(IsNull(oConnResp.oLeitura.RespTecnico(cont)) Or oConnResp.oLeitura.RespTecnico(cont) = "", ".", oConnResp.oLeitura.RespTecnico(cont)))
   Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "DataResultado", IIf(IsNull(oConnResp.oLeitura.DataLeitura(cont)) Or oConnResp.oLeitura.DataLeitura(cont) = "", "01/01/1964", oConnResp.oLeitura.DataLeitura(cont)))
   Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "NaoConformidade", IIf(oConnResp.oLeitura.Conformidades(cont) = "", ".", oConnResp.oLeitura.Conformidades(cont)))

End If
End Sub

Private Sub Load_Data_Leitura()
If Not oConnResp.oLeitura Is Nothing Then
   Table_Consulta.Fields("EstadoResultado").Value = IIf(IsNull(oConnResp.oLeitura.UFLeitura(cont)) Or oConnResp.oLeitura.UFLeitura(cont) = "", ".", oConnResp.oLeitura.UFLeitura(cont))
   Table_Consulta.Fields("ResponsavelResultado").Value = IIf(IsNull(oConnResp.oLeitura.RespTecnico(cont)) Or oConnResp.oLeitura.RespTecnico(cont) = "", ".", oConnResp.oLeitura.RespTecnico(cont))
   Table_Consulta.Fields("DataResultado").Value = IIf(IsNull(oConnResp.oLeitura.DataLeitura(cont)) Or oConnResp.oLeitura.DataLeitura(cont) = "", "01/01/1964", oConnResp.oLeitura.DataLeitura(cont))
   Table_Consulta.Fields("NaoConformidade").Value = IIf(oConnResp.oLeitura.Conformidades(cont) = "", ".", oConnResp.oLeitura.Conformidades(cont))
End If
End Sub

Private Sub Load_Data_Resultado_New(ByVal WCodigoCliente As Long)
   If IsDate(oConnResp.oLotePlacas.DataValidadeResultado(cont)) = True Then
       Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "EstadoResultado", oConnResp.oLotePlacas.EstadoResultado(cont))
       Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "ResponsavelResultado", oConnResp.oLotePlacas.ResponsavelResultado(cont))
       Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "DataResultado", Format(IIf(Not IsDate(oConnResp.oLotePlacas.DataResultado(cont)), "01/01/2001", oConnResp.oLotePlacas.DataResultado(cont)), "dd/MM/yyyy"))
       Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "DataValidadeResultado", Format(oConnResp.oLotePlacas.DataValidadeResultado(cont), "dd/MM/yyyy"))
       Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "Data_Validade_Resultado", Format(oConnResp.oLotePlacas.DataValidadeResultado(cont), "dd/MM/yyyy"))
       Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "tipocertificadoresultado", oConnResp.oLotePlacas.TipoCertificado2(cont))
       Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "NumeroCertificadoDefinitivo", IIf(IsNull(oConnResp.oLotePlacas.NumeroCertificadoDefinitivo(cont)) Or oConnResp.oLotePlacas.NumeroCertificadoDefinitivo(cont) = "", "000000000", oConnResp.oLotePlacas.NumeroCertificadoDefinitivo(cont)))
   ElseIf oConnResp.oLotePlacas.TipoCertificado2(cont) <> "" And Not IsNull(oConnResp.oLotePlacas.TipoCertificado2(cont)) Then
      Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "tipocertificadoresultado", oConnResp.oLotePlacas.TipoCertificado2(cont))
      Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "tipocertificadoresultado", oConnResp.oLotePlacas.TipoCertificado2(cont))
   End If
End Sub

Private Sub Load_Data_Resultado()
   If IsDate(oConnResp.oLotePlacas.DataValidadeResultado(cont)) = True Then
       Table_Consulta.Fields("EstadoResultado").Value = oConnResp.oLotePlacas.EstadoResultado(cont)
       Table_Consulta.Fields("ResponsavelResultado").Value = oConnResp.oLotePlacas.ResponsavelResultado(cont)
       Table_Consulta.Fields("DataResultado").Value = Format(IIf(Not IsDate(oConnResp.oLotePlacas.DataResultado(cont)), "01/01/2001", oConnResp.oLotePlacas.DataResultado(cont)), "dd/MM/yyyy")
       Table_Consulta.Fields("DataValidadeResultado").Value = Format(oConnResp.oLotePlacas.DataValidadeResultado(cont), "dd/MM/yyyy")
       Table_Consulta.Fields("Data_Validade_Resultado").Value = Format(oConnResp.oLotePlacas.DataValidadeResultado(cont), "dd/MM/yyyy")
       Table_Consulta.Fields("tipocertificadoresultado").Value = oConnResp.oLotePlacas.TipoCertificado2(cont)
       Table_Consulta.Fields("NumeroCertificadoDefinitivo").Value = IIf(IsNull(oConnResp.oLotePlacas.NumeroCertificadoDefinitivo(cont)) Or oConnResp.oLotePlacas.NumeroCertificadoDefinitivo(cont) = "", "000000000", oConnResp.oLotePlacas.NumeroCertificadoDefinitivo(cont))
   ElseIf oConnResp.oLotePlacas.TipoCertificado2(cont) <> "" And Not IsNull(oConnResp.oLotePlacas.TipoCertificado2(cont)) Then
       Table_Consulta.Fields("tipocertificadoresultado").Value = oConnResp.oLotePlacas.TipoCertificado2(cont)
   End If
End Sub
Private Sub Load_Data_Cronotacografo_NEw(ByVal WCodigoCliente As Long)

Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "MarcaEquip", IIf(IsNull(oConnResp.oLotePlacas.MarcaEquip(cont)), " ", oConnResp.oLotePlacas.MarcaEquip(cont)))
Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "ModeloEquip", IIf(IsNull(oConnResp.oLotePlacas.ModeloEquip(cont)), " ", oConnResp.oLotePlacas.ModeloEquip(cont)))
Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "NroSerieEquip", IIf(IsNull(oConnResp.oLotePlacas.NroSerieEquip(cont)), " ", oConnResp.oLotePlacas.NroSerieEquip(cont)))
      
End Sub
Private Sub Load_Data_Cronotacografo()
Table_Consulta.Fields("MarcaEquip").Value = IIf(IsNull(oConnResp.oLotePlacas.MarcaEquip(cont)), " ", oConnResp.oLotePlacas.MarcaEquip(cont))
Table_Consulta.Fields("ModeloEquip").Value = IIf(IsNull(oConnResp.oLotePlacas.ModeloEquip(cont)), " ", oConnResp.oLotePlacas.ModeloEquip(cont))
Table_Consulta.Fields("NroSerieEquip").Value = IIf(IsNull(oConnResp.oLotePlacas.NroSerieEquip(cont)), " ", oConnResp.oLotePlacas.NroSerieEquip(cont))
End Sub
Public Function Integracao_SGA_Inmetro_Consulta_Selagem_Detalhada_Veiculo_CNPJ(ByVal WCodigoCliente As Long) As Boolean
Dim contPlacas As Integer
Dim oLote As CanadaConnectivity.LotePlacas

Dim bd As DataBase
Dim QRs As QueryDef
Dim rs  As Recordset
Set bd = CurrentDb()
Set QRs = bd.QueryDefs("Q_PLAXCLI_CODCLIENTE")
QRs.Parameters(0) = WCodigoCliente
Set rs = QRs.OpenRecordset()

contPlacas = 0
Set ofuncoes = New CanadaConnectivity.FuncoesDiversas
InicializaVariaveisConectividade
If IsCloud = False Then
'[CONECTIVIDADE_NOVA] validacao da Conectividade antiga desativada:
   'If ofuncoes.ConectividadeIsAlive("Conectividade") = False Then
      'MsgBox "CONECTIVIDADE NAO ESTA ATIVA, IMPOSSIVEL PROSSEGUIR.", vbCritical
      'Integracao_SGA_Inmetro_Consulta_Selagem_Detalhada_Veiculo_CNPJ = False
      'Exit Function
   'End If
Else
   If ofuncoes.ExisteDiretorio(CaminhoSGaRemote) = False Then
      MsgBox "CONEXAO OFF LINE, IMPOSSIVEL PROSSEGUIR.", vbCritical
      Integracao_SGA_Inmetro_Consulta_Selagem_Detalhada_Veiculo_CNPJ = False
      Exit Function
   End If
End If

ofuncoes.ExcluiArquivos pFileNameReq
ofuncoes.ExcluiArquivos pFileNameResp

If rs.RecordCount <> 0 Then
   Set oLote = New CanadaConnectivity.LotePlacas
   Set oconnreq = New CanadaConnectivity.ReqConsultaLotePlacas
   oLote.InicializaTable
   rs.MoveFirst
   While Not rs.EOF
       oLote.AddVeiculo
       contPlacas = contPlacas + 1
       oLote.CodCliente = WCodigoCliente
       oLote.Placa = rs!Placa
       oLote.Update
       
       rs.MoveNext
   Wend
   Ret = SysCmd(SYSCMD_SETSTATUS, "Realizando a Consulta de Lotes de Placas >>>> Por Favor Aguarde.")
   oconnreq.CreateConsultaLotePlacasEnvio oLote, Retorna_Links_Uteis(6), Retorna_Links_Uteis(7), 13, "www.canainfo.com.br"
  
   If IsCloud = False Then
'[CONECTIVIDADE_NOVA] ativacao da janela do app antigo desativada:
      'AppActivate ("Conectividade - " & RetornaLoginInMetro)
   End If

   Do While ofuncoes.CheckFile(pFileNameResp) = False
      Sleep (500)
      DoEvents
   Loop

   Sleep (2000)
   Set oConnResp = New CanadaConnectivity.LoadRespConsultaLotePlacas
   oConnResp.Load_Objetos pNameUserCloud, IsCloud
   AppActivate "SGa"
   If oConnResp.Status = False Then
      Exit Function
   End If
   For cont = 0 To oConnResp.oLotePlacas.RowsCount - 1
       Call Atualiza_Tabela_PLACXCLI_ConsultaLotePlacas(oLote.Placa(cont), WCodigoCliente)
       'Load_Table_Veiculo_New (WCodigoCliente)
       'Load_Data_Selagem_New (WCodigoCliente)
       'Load_Data_Ensaio_New (WCodigoCliente)
       'Load_Data_Cronotacografo_NEw (WCodigoCliente)
       'Load_Data_Certificado_New (WCodigoCliente)
       'Load_Data_Resultado_New (WCodigoCliente)
       'Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "DataConsulta", Date)
       'Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "Tipo_Veic", "C")
       'Call Atualiza_Tabela_PLACXCLI(oConnResp.oLotePlacas.Placa(cont), WCodigoCliente, "Status", 1)
   Next
   
End If

Ret = SysCmd(SYSCMD_SETSTATUS, "Consulta Finalizada.")
Status = True

Finaliza:
   rs.Close
   QRs.Close
   bd.Close
   Set bd = Nothing
   Set QRs = Nothing
   Set rs = Nothing

End Function

Function Atualiza_Tabela_PLACXCLI_ConsultaLotePlacas(wPlaca As String, WCodigoCliente As Long) As Long
On Error GoTo TratamentodeErro:

Dim bd As DataBase
Dim QRs As QueryDef
Dim rs  As Recordset
Set bd = CurrentDb()
Set QRs = bd.QueryDefs("Q_PlacXCLI")
QRs.Parameters(0) = wPlaca
QRs.Parameters(1) = WCodigoCliente
Set rs = QRs.OpenRecordset()

BeginTrans
If rs.RecordCount <> 0 Then
   rs.Edit
Else
   rs.AddNew
   rs.Fields("Placa").Value = wPlaca
   rs.Fields("CodigoCliente").Value = WCodigoCliente
End If

rs.Fields("Placa").Value = wPlaca
rs.Fields("CodigoCliente").Value = WCodigoCliente
rs.Fields("Renavan").Value = oConnResp.oLotePlacas.Renavam(cont)
rs.Fields("MarcaVeiculo").Value = oConnResp.oLotePlacas.MarcaVeiculo(cont)
rs.Fields("ModeloVeiculo").Value = oConnResp.oLotePlacas.ModeloVeiculo(cont)
rs.Fields("AnoVeiculo").Value = oConnResp.oLotePlacas.AnoVeiculo(cont)
rs.Fields("AROPNEU").Value = oConnResp.oLotePlacas.AroPneu(cont)
rs.Fields("Pneu").Value = oConnResp.oLotePlacas.Pneu(cont)

Dim odataselagem As String
odataselagem = Format(IIf(Not IsDate(Mid(oConnResp.oLotePlacas.DataSelagem(cont), 1, 10)), "01/01/1964", Mid(oConnResp.oLotePlacas.DataSelagem(cont), 1, 10)), "dd/MM/yyyy")
rs.Fields("DataSelagem").Value = CDate(odataselagem)
rs.Fields("PostoSelagem").Value = oConnResp.oLotePlacas.PostoSelagem(cont)
rs.Fields("DataValidadeResultado").Value = Format(CDate(odataselagem) + 30, "dd/MM/yyyy")
rs.Fields("TipoCertificadoResultado").Value = "CERT.PRELIMINAR"

Dim oDataEnsaio As String
oDataEnsaio = Format(IIf(Not IsDate(oConnResp.oLotePlacas.DataEnsaio(cont)), "01/01/2001", oConnResp.oLotePlacas.DataEnsaio(cont)), "dd/MM/yyyy")
rs.Fields("UFEnsaio").Value = oConnResp.oLotePlacas.UFEnsaio(cont)
rs.Fields("PostoEnsaio").Value = IIf(oConnResp.oLotePlacas.PostoEnsaio(cont) = "", ".", oConnResp.oLotePlacas.PostoEnsaio(cont))
rs.Fields("DataEnsaio").Value = CDate(oDataEnsaio)
rs.Fields("VelMaxEnsaio").Value = IIf(oConnResp.oLotePlacas.VelMaxEnsaio(cont) = "", "0", oConnResp.oLotePlacas.VelMaxEnsaio(cont))
rs.Fields("PadraoUtilizadoEnsaio").Value = IIf(oConnResp.oLotePlacas.PadraoUtilizadoEnsaio(cont) = "", ".", oConnResp.oLotePlacas.PadraoUtilizadoEnsaio(cont))
rs.Fields("DataValidadeResultado").Value = Format(CDate(oDataEnsaio) + 30, "dd/MM/yyyy")
rs.Fields("Data_Validade_Resultado").Value = Format(CDate(oDataEnsaio) + 30, "dd/MM/yyyy")
rs.Fields("TipoCertificadoResultado").Value = "CERT. PRELIMINAR"

rs.Fields("MarcaEquip").Value = IIf(IsNull(oConnResp.oLotePlacas.MarcaEquip(cont)), " ", oConnResp.oLotePlacas.MarcaEquip(cont))
rs.Fields("ModeloEquip").Value = IIf(IsNull(oConnResp.oLotePlacas.ModeloEquip(cont)), " ", oConnResp.oLotePlacas.ModeloEquip(cont))
rs.Fields("NroSerieEquip").Value = IIf(IsNull(oConnResp.oLotePlacas.NroSerieEquip(cont)), " ", oConnResp.oLotePlacas.NroSerieEquip(cont))

If Not oConnResp.oLotePlacas.TipoCertificado1(cont) = "." Then
   rs.Fields("NumeroCertificado2").Value = IIf(oConnResp.oLotePlacas.NumeroCertificado2(cont) = "", ".", oConnResp.oLotePlacas.NumeroCertificado2(cont))
   rs.Fields("DataCertificado2").Value = IIf(oConnResp.oLotePlacas.DataCertificado2(cont) = "", ".", oConnResp.oLotePlacas.DataCertificado2(cont))
   rs.Fields("ValidadeCertificado2").Value = Format(IIf(oConnResp.oLotePlacas.ValidadeCertificado2(cont) = "", ".", oConnResp.oLotePlacas.ValidadeCertificado2(cont)), "dd/MM/yyyy")
   rs.Fields("TipoCertificado2").Value = IIf(oConnResp.oLotePlacas.TipoCertificado2(cont) = "", ".", oConnResp.oLotePlacas.TipoCertificado2(cont))
   '######### este é o campo que armazena validade do certificado mais recente
   If oConnResp.oLotePlacas.TipoCertificado1(cont) <> "" Then
     rs.Fields("TipoCertificado1").Value = oConnResp.oLotePlacas.TipoCertificado1(cont)
   End If
   If oConnResp.oLotePlacas.DataCertificado1(cont) <> "" Then
      rs.Fields("DataResultado").Value = oConnResp.oLotePlacas.DataCertificado1(cont)
   End If
   If oConnResp.oLotePlacas.ValidadeCertificado1(cont) <> "" Then
      rs.Fields("DataValidadeResultado").Value = Format(oConnResp.oLotePlacas.ValidadeCertificado1(cont), "dd/MM/yyyy")
      rs.Fields("Data_Validade_Resultado").Value = Format(oConnResp.oLotePlacas.ValidadeCertificado1(cont), "dd/MM/yyyy")
   End If
   If oConnResp.oLotePlacas.NumeroCertificado1(cont) <> "" Then
      rs.Fields("NumeroCertificadoDefinitivo").Value = oConnResp.oLotePlacas.NumeroCertificado1(cont)
   End If
Else
    rs.Fields("NumeroCertificado2").Value = "."
    rs.Fields("DataCertificado2").Value = "01/01/2001"
    rs.Fields("ValidadeCertificado2").Value = "01/01/2001"
    rs.Fields("TipoCertificado2").Value = "."
End If

If IsDate(oConnResp.oLotePlacas.DataValidadeResultado(cont)) = True Then
       rs.Fields("EstadoResultado") = oConnResp.oLotePlacas.EstadoResultado(cont)
       rs.Fields("ResponsavelResultado") = oConnResp.oLotePlacas.ResponsavelResultado(cont)
       rs.Fields("DataValidadeResultado") = oConnResp.oLotePlacas.DataValidadeResultado(cont)
       rs.Fields("Data_Validade_Resultado") = oConnResp.oLotePlacas.DataValidadeResultado(cont)
       rs.Fields("tipocertificadoresultado") = oConnResp.oLotePlacas.TipoCertificado2(cont)
       rs.Fields("NumeroCertificadoDefinitivo") = "0000000"
   ElseIf oConnResp.oLotePlacas.TipoCertificado2(cont) <> "" And Not IsNull(oConnResp.oLotePlacas.TipoCertificado2(cont)) Then
      rs.Fields("tipocertificadoresultado") = "CERT. PRELIMINAR"
   End If

rs.Fields("Data_Ultima_Consulta").Value = Date
rs.Fields("DataConsulta").Value = Date
rs.Fields("Tipo_Veic").Value = "C"
rs.Fields("Status").Value = 1

rs.Update
Finaliza:
  CommitTrans
  rs.Close
  bd.Close
  Set rs = Nothing
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

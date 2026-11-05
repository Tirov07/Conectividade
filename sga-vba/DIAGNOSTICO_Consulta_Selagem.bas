Attribute VB_Name = "DIAGNOSTICO_Consulta_Selagem"
Option Compare Database
Option Explicit
#If VBA7 And Win64 Then
     Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

'===========================================================================
'  MODULO DE DIAGNOSTICO (TEMPORARIO)
'---------------------------------------------------------------------------
'  Objetivo: descobrir POR QUE o Load_Objetos estoura "objeto nao definido"
'  (80004003) na consulta de selagem, SEM depender de recompilar o .exe.
'
'  Como usar:
'   1. Importe este modulo no SGA (Alt+F11 > Arquivo > Importar).
'   2. Faça uma consulta normal pela tela (a que gera o erro). O modulo real
'      ainda vai estourar o erro -- tudo bem.
'   3. LOGO EM SEGUIDA, antes de fazer qualquer outra coisa, va no VBA:
'      janela Verificacao Imediata (Ctrl+G) e rode:  Diagnostico_Resp
'      (ou execute Diagnostico_Resp pela lista de macros).
'   4. Copie TODO o texto que aparecer na caixa de mensagem e me mande.
'
'  Ele inspeciona o arquivo C:\Sga\Resp\<usuario>.xml que a extensao gravou:
'  tamanho, primeiros bytes (BOM?), se o MSXML consegue carregar, e tenta
'  rodar o Load_Objetos capturando o erro exato.
'===========================================================================

Public Sub Diagnostico_Resp()
    Dim sUser As String, sPath As String, sMsg As String
    Dim ofuncoes As CanadaConnectivity.FuncoesDiversas

    On Error Resume Next
    Set ofuncoes = New CanadaConnectivity.FuncoesDiversas
    sUser = ofuncoes.ReturnUserWindows
    On Error GoTo 0

    sPath = "C:\Sga\Resp\" & sUser & ".xml"

    sMsg = "=== DIAGNOSTICO CONECTIVIDADE ===" & vbCrLf
    sMsg = sMsg & "Usuario Windows: [" & sUser & "]" & vbCrLf
    sMsg = sMsg & "Caminho resp:    " & sPath & vbCrLf

    '--- 1) O arquivo existe? ---
    Dim existe As Boolean
    existe = (Dir(sPath) <> "")
    sMsg = sMsg & "Arquivo existe:  " & existe & vbCrLf

    If Not existe Then
        ' tenta listar o que tem na pasta
        Dim sLista As String, sF As String
        sF = Dir("C:\Sga\Resp\*.*")
        Do While sF <> ""
            sLista = sLista & sF & "  "
            sF = Dir
        Loop
        sMsg = sMsg & "Conteudo da pasta Resp: " & sLista & vbCrLf
        MsgBox sMsg, vbInformation, "Diagnostico"
        Exit Sub
    End If

    '--- 2) Tamanho e primeiros bytes (detecta BOM / encoding) ---
    Dim iFile As Integer, lTam As Long
    Dim b1 As Byte, b2 As Byte, b3 As Byte, b4 As Byte
    iFile = FreeFile
    Open sPath For Binary Access Read As #iFile
    lTam = LOF(iFile)
    If lTam >= 1 Then Get #iFile, 1, b1
    If lTam >= 2 Then Get #iFile, 2, b2
    If lTam >= 3 Then Get #iFile, 3, b3
    If lTam >= 4 Then Get #iFile, 4, b4
    Close #iFile

    sMsg = sMsg & "Tamanho (bytes): " & lTam & vbCrLf
    sMsg = sMsg & "4 primeiros bytes (hex): " _
        & Right("00" & Hex(b1), 2) & " " _
        & Right("00" & Hex(b2), 2) & " " _
        & Right("00" & Hex(b3), 2) & " " _
        & Right("00" & Hex(b4), 2) & vbCrLf
    If b1 = &HEF And b2 = &HBB And b3 = &HBF Then
        sMsg = sMsg & ">> Tem BOM UTF-8 (EF BB BF)" & vbCrLf
    ElseIf b1 = &H3C Then
        sMsg = sMsg & ">> SEM BOM (comeca direto com '<')" & vbCrLf
    Else
        sMsg = sMsg & ">> Inicio inesperado" & vbCrLf
    End If

    '--- 3) O MSXML consegue carregar o arquivo? ---
    Dim oXml As Object
    On Error Resume Next
    Set oXml = CreateObject("MSXML2.DOMDocument.6.0")
    If oXml Is Nothing Then Set oXml = CreateObject("MSXML2.DOMDocument")
    oXml.async = False
    oXml.validateOnParse = False
    Dim okLoad As Boolean
    okLoad = oXml.Load(sPath)
    sMsg = sMsg & "MSXML Load OK:   " & okLoad & vbCrLf
    If Not okLoad Then
        sMsg = sMsg & "MSXML parseError: " & oXml.parseError.reason _
            & " (linha " & oXml.parseError.Line & ", col " & oXml.parseError.linepos & ")" & vbCrLf
    Else
        sMsg = sMsg & "Raiz do XML:     " & oXml.documentElement.nodeName & vbCrLf
        sMsg = sMsg & "Namespace raiz:  [" & oXml.documentElement.namespaceURI & "]" & vbCrLf
        sMsg = sMsg & "Qtd filhos raiz: " & oXml.documentElement.childNodes.Length & vbCrLf
    End If
    On Error GoTo 0

    '--- 4) Tenta o Load_Objetos de verdade, capturando o erro ---
    Dim oResp As CanadaConnectivity.LoadRespConsultaSelagem
    Dim bIsCloud As Boolean
    On Error Resume Next
    bIsCloud = IsCloud()
    Err.Clear
    Set oResp = New CanadaConnectivity.LoadRespConsultaSelagem
    sMsg = sMsg & "---- Teste Load_Objetos ----" & vbCrLf
    sMsg = sMsg & "oResp instanciado (Is Nothing)? " & (oResp Is Nothing) & vbCrLf
    sMsg = sMsg & "IsCloud: " & bIsCloud & "  | pNameUserCloud p/ chamada: [" & sUser & "]" & vbCrLf
    oResp.Load_Objetos sUser, bIsCloud
    If Err.Number <> 0 Then
        sMsg = sMsg & "ERRO no Load_Objetos: " & Err.Number & " - " & Err.Description & vbCrLf
    Else
        sMsg = sMsg & "Load_Objetos OK! STATUS=" & oResp.oSTATUS & vbCrLf
        Err.Clear
        sMsg = sMsg & "oGRU Is Nothing? " & (oResp.oGRU Is Nothing) & vbCrLf
        If Err.Number <> 0 Then sMsg = sMsg & "  (erro ao acessar oGRU: " & Err.Description & ")" & vbCrLf
    End If
    On Error GoTo 0

    MsgBox sMsg, vbInformation, "Diagnostico Conectividade"
End Sub

'===========================================================================
'  Diagnostico_Consulta_Completa
'---------------------------------------------------------------------------
'  Faz a consulta do ZERO (gera o req, espera a extensao gravar o resp) e
'  JA roda o diagnostico em cima do resp recem-gerado -- sem depender de o
'  arquivo "sobrar" de uma consulta anterior.
'
'  Troque a placa de teste abaixo se quiser. Rode assim na janela imediata:
'       Diagnostico_Consulta_Completa "CVN8588"
'===========================================================================
Public Sub Diagnostico_Consulta_Completa(Optional ByVal wPlaca As String = "CVN8588")
    Dim ofuncoes As CanadaConnectivity.FuncoesDiversas
    Dim oconnreq As CanadaConnectivity.ReqConsultaSelagem
    Dim sUser As String, sReq As String, sResp As String
    Dim bIsCloud As Boolean

    Set ofuncoes = New CanadaConnectivity.FuncoesDiversas
    sUser = ofuncoes.ReturnUserWindows
    bIsCloud = IsCloud()

    If bIsCloud Then
        sUser = "SGaCloud"
        sReq = "\\tsclient\c\sga\Req\" & sUser & ".xml"
        sResp = "\\tsclient\c\sga\Resp\" & sUser & ".xml"
    Else
        sReq = "C:\Sga\Req\" & sUser & ".xml"
        sResp = "C:\Sga\Resp\" & sUser & ".xml"
    End If

    ' limpa req/resp antes
    ofuncoes.ExcluiArquivos sReq
    ofuncoes.ExcluiArquivos sResp

    ' gera a requisicao (op 6)
    Set oconnreq = New CanadaConnectivity.ReqConsultaSelagem
    oconnreq.CreateConsultaSelagem wPlaca, "", "", False, _
        Retorna_Links_Uteis(6), Retorna_Links_Uteis(10), 6, "www.canainfo.com.br"

    ' espera o resp aparecer (ate ~60s)
    Dim i As Integer
    For i = 1 To 120
        If ofuncoes.CheckFile(sResp) = True Then Exit For
        Sleep 500
        DoEvents
    Next i

    If ofuncoes.CheckFile(sResp) = False Then
        MsgBox "O resp NAO apareceu em 60s. Verifique se a extensao/host estao rodando.", vbCritical
        Exit Sub
    End If

    ' pequena folga e roda o diagnostico sobre o arquivo gerado
    Sleep 300
    Diagnostico_Resp
End Sub

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

'===========================================================================
'  Diagnostico_Testa_Namespace
'---------------------------------------------------------------------------
'  Testa a hipotese de que a DLL Load_Objetos so acha os nos quando o XML
'  esta SEM namespace (ou com outro namespace). Pega o resp atual, gera 3
'  variantes em C:\Sga\Resp e tenta Load_Objetos em cada uma:
'     A) como esta  (xmlns="www.canainfo.com.br")
'     B) SEM nenhum xmlns
'     C) com xmlns="http://www.canainfo.com.br"  (com http://)
'
'  Rode DEPOIS de uma consulta (para existir o resp). Ex.:
'       Diagnostico_Testa_Namespace
'  Me mande o texto da caixa.
'===========================================================================
Public Sub Diagnostico_Testa_Namespace()
    Dim ofuncoes As CanadaConnectivity.FuncoesDiversas
    Dim sUser As String, sResp As String, sMsg As String
    Dim sOrig As String

    Set ofuncoes = New CanadaConnectivity.FuncoesDiversas
    sUser = ofuncoes.ReturnUserWindows
    sResp = "C:\Sga\Resp\" & sUser & ".xml"

    If Dir(sResp) = "" Then
        MsgBox "Nao existe " & sResp & ". Faca uma consulta primeiro e rode de novo.", vbCritical
        Exit Sub
    End If

    ' le o conteudo atual
    sOrig = LeTexto(sResp)
    sMsg = "=== TESTE DE NAMESPACE ===" & vbCrLf

    ' --- Variante A: como esta ---
    sMsg = sMsg & "A) original  -> " & TentaLoad(sResp, sOrig, sUser) & vbCrLf

    ' --- Variante B: SEM xmlns ---
    Dim sSemNs As String
    sSemNs = Replace(sOrig, " xmlns=""www.canainfo.com.br""", "")
    sSemNs = Replace(sSemNs, " xmlns=""http://www.canainfo.com.br""", "")
    sMsg = sMsg & "B) sem xmlns -> " & TentaLoad(sResp, sSemNs, sUser) & vbCrLf

    ' --- Variante C: com http:// ---
    Dim sHttp As String
    sHttp = Replace(sOrig, "xmlns=""www.canainfo.com.br""", "xmlns=""http://www.canainfo.com.br""")
    sMsg = sMsg & "C) http://   -> " & TentaLoad(sResp, sHttp, sUser) & vbCrLf

    ' restaura o original
    EscreveTexto sResp, sOrig

    MsgBox sMsg, vbInformation, "Teste Namespace"
End Sub

Private Function TentaLoad(ByVal sResp As String, ByVal sConteudo As String, ByVal sUser As String) As String
    On Error Resume Next
    EscreveTexto sResp, sConteudo
    Dim oResp As CanadaConnectivity.LoadRespConsultaSelagem
    Set oResp = New CanadaConnectivity.LoadRespConsultaSelagem
    Err.Clear
    oResp.Load_Objetos sUser, False
    If Err.Number <> 0 Then
        TentaLoad = "ERRO " & Err.Number & " (" & Err.Description & ")"
    Else
        TentaLoad = "OK  STATUS=" & oResp.oSTATUS
    End If
    Set oResp = Nothing
    On Error GoTo 0
End Function

Private Function LeTexto(ByVal sPath As String) As String
    Dim iFile As Integer, s As String
    iFile = FreeFile
    Open sPath For Binary Access Read As #iFile
    s = Space$(LOF(iFile))
    Get #iFile, 1, s
    Close #iFile
    LeTexto = s
End Function

Private Sub EscreveTexto(ByVal sPath As String, ByVal sConteudo As String)
    Dim iFile As Integer
    If Dir(sPath) <> "" Then Kill sPath
    iFile = FreeFile
    Open sPath For Binary Access Write As #iFile
    Put #iFile, 1, sConteudo
    Close #iFile
End Sub

'===========================================================================
'  Diagnostico_Isola_Classe
'---------------------------------------------------------------------------
'  Descobre se o problema e (1) a DLL inteira, (2) so a classe
'  LoadRespConsultaSelagem, ou (3) o metodo Load_Objetos.
'
'  Testa instanciar varias classes da DLL e chamar Load_Objetos passo a passo,
'  mostrando em qual exato ponto o 80004003 aparece.
'
'  Rode (NAO precisa de resp existir para a 1a parte):
'       Diagnostico_Isola_Classe
'===========================================================================
Public Sub Diagnostico_Isola_Classe()
    Dim sMsg As String
    sMsg = "=== ISOLAR CLASSE / DLL ===" & vbCrLf

    '--- 1) Instanciar classes da DLL (sem chamar metodo) ---
    sMsg = sMsg & InstanciaOK("FuncoesDiversas") & vbCrLf
    sMsg = sMsg & InstanciaOK("ReqConsultaSelagem") & vbCrLf
    sMsg = sMsg & InstanciaOK("LoadRespConsultaSelagem") & vbCrLf
    sMsg = sMsg & InstanciaOK("LoadRespConsultaLotePlacas") & vbCrLf

    '--- 2) Na classe LoadRespConsultaSelagem, acessar propriedades ANTES de Load ---
    sMsg = sMsg & "---- antes do Load_Objetos ----" & vbCrLf
    Dim oResp2 As CanadaConnectivity.LoadRespConsultaSelagem
    On Error Resume Next
    Set oResp2 = New CanadaConnectivity.LoadRespConsultaSelagem
    Err.Clear
    Dim dummy As Variant

    dummy = oResp2.oSTATUS
    sMsg = sMsg & "ler oSTATUS antes do Load: " & IIf(Err.Number = 0, "OK (" & dummy & ")", "ERRO " & Err.Number) & vbCrLf
    Err.Clear

    Dim oGruNada As Boolean
    oGruNada = (oResp2.oGRU Is Nothing)
    sMsg = sMsg & "oGRU Is Nothing antes do Load: " & IIf(Err.Number = 0, CStr(oGruNada), "ERRO " & Err.Number & " (" & Err.Description & ")") & vbCrLf
    Err.Clear

    Dim oVeicNada As Boolean
    oVeicNada = (oResp2.oVeiculo Is Nothing)
    sMsg = sMsg & "oVeiculo Is Nothing antes do Load: " & IIf(Err.Number = 0, CStr(oVeicNada), "ERRO " & Err.Number & " (" & Err.Description & ")") & vbCrLf
    On Error GoTo 0

    MsgBox sMsg, vbInformation, "Isolar Classe"
End Sub

Private Function InstanciaOK(ByVal sClasse As String) As String
    On Error Resume Next
    Dim o As Object
    Err.Clear
    Set o = CreateObject("CanadaConnectivity." & sClasse)
    If Err.Number <> 0 Then
        ' tenta via New nos tipos conhecidos (CreateObject pode nao achar pelo ProgID)
        InstanciaOK = sClasse & ": CreateObject ERRO " & Err.Number & " (" & Err.Description & ")"
    Else
        InstanciaOK = sClasse & ": CreateObject OK"
    End If
    Set o = Nothing
    On Error GoTo 0
End Function

'===========================================================================
'  Diagnostico_Varre_Formato
'---------------------------------------------------------------------------
'  Como namespace nao resolveu, este teste varre HIPOTESES de formato que a
'  DLL Load_Objetos pode exigir, gerando variantes do resp e tentando o Load
'  em cada uma. A primeira que der "OK" revela o que a DLL precisa.
'
'  Variantes testadas:
'   0) original (como a extensao gravou)
'   1) sem a 1a linha <?xml ...?>  (declaracao removida)
'   2) declaracao trocada p/ ISO-8859-1
'   3) declaracao trocada p/ Windows-1252
'   4) STATUS minusculo <STATUS>true</STATUS>
'   5) raiz sem xmlns E declaracao ISO-8859-1 juntas
'   6) tudo em UMA linha (sem quebras/identacao)
'
'  Rode DEPOIS de uma consulta (precisa existir o resp). Ex.:
'       Diagnostico_Varre_Formato
'===========================================================================
Public Sub Diagnostico_Varre_Formato()
    Dim ofuncoes As CanadaConnectivity.FuncoesDiversas
    Dim sUser As String, sResp As String, sMsg As String, sOrig As String

    Set ofuncoes = New CanadaConnectivity.FuncoesDiversas
    sUser = ofuncoes.ReturnUserWindows
    sResp = "C:\Sga\Resp\" & sUser & ".xml"

    If Dir(sResp) = "" Then
        MsgBox "Nao existe " & sResp & ". Faca uma consulta primeiro.", vbCritical
        Exit Sub
    End If
    sOrig = LeTexto(sResp)

    sMsg = "=== VARRE FORMATO ===" & vbCrLf

    ' 0) original
    sMsg = sMsg & "0) original                 -> " & TentaLoad(sResp, sOrig, sUser) & vbCrLf

    ' 1) sem a declaracao <?xml?>
    Dim s1 As String
    s1 = sOrig
    If InStr(s1, "?>") > 0 Then s1 = Mid(s1, InStr(s1, "?>") + 2)
    s1 = LTrim(Replace(Replace(s1, vbCr, ""), vbLf, "")) ' tira quebra inicial
    s1 = sOrig ' recomputa de forma segura abaixo
    Dim p As Long
    p = InStr(sOrig, "?>")
    If p > 0 Then s1 = Mid(sOrig, p + 2) Else s1 = sOrig
    sMsg = sMsg & "1) sem <?xml?>              -> " & TentaLoad(sResp, s1, sUser) & vbCrLf

    ' 2) declaracao ISO-8859-1
    Dim s2 As String
    s2 = Replace(sOrig, "encoding=""utf-8""", "encoding=""ISO-8859-1""")
    s2 = Replace(s2, "encoding=""UTF-8""", "encoding=""ISO-8859-1""")
    sMsg = sMsg & "2) encoding ISO-8859-1      -> " & TentaLoad(sResp, s2, sUser) & vbCrLf

    ' 3) declaracao Windows-1252
    Dim s3 As String
    s3 = Replace(sOrig, "encoding=""utf-8""", "encoding=""Windows-1252""")
    s3 = Replace(s3, "encoding=""UTF-8""", "encoding=""Windows-1252""")
    sMsg = sMsg & "3) encoding Windows-1252    -> " & TentaLoad(sResp, s3, sUser) & vbCrLf

    ' 4) STATUS minusculo
    Dim s4 As String
    s4 = Replace(sOrig, "<STATUS>True</STATUS>", "<STATUS>true</STATUS>")
    s4 = Replace(s4, "<STATUS>False</STATUS>", "<STATUS>false</STATUS>")
    sMsg = sMsg & "4) STATUS minusculo         -> " & TentaLoad(sResp, s4, sUser) & vbCrLf

    ' 5) sem xmlns + ISO-8859-1
    Dim s5 As String
    s5 = Replace(s2, " xmlns=""www.canainfo.com.br""", "")
    s5 = Replace(s5, " xmlns=""http://www.canainfo.com.br""", "")
    sMsg = sMsg & "5) sem xmlns + ISO-8859-1   -> " & TentaLoad(sResp, s5, sUser) & vbCrLf

    ' 6) tudo numa linha so
    Dim s6 As String
    s6 = Replace(sOrig, vbCrLf, "")
    s6 = Replace(s6, vbCr, "")
    s6 = Replace(s6, vbLf, "")
    sMsg = sMsg & "6) linha unica              -> " & TentaLoad(sResp, s6, sUser) & vbCrLf

    ' restaura
    EscreveTexto sResp, sOrig
    MsgBox sMsg, vbInformation, "Varre Formato"
End Sub

'===========================================================================
'  Diagnostico_Captura_Resp
'---------------------------------------------------------------------------
'  Faz a consulta do ZERO e, no instante em que o resp aparece, COPIA o
'  conteudo para C:\Sga\Resp\CAPTURA.txt (que NAO some), para voce conseguir
'  abrir/copiar com calma. Tambem mostra o conteudo numa caixa.
'
'  Rode:  Diagnostico_Captura_Resp "CVN8588"
'===========================================================================
Public Sub Diagnostico_Captura_Resp(Optional ByVal wPlaca As String = "CVN8588")
    Dim ofuncoes As CanadaConnectivity.FuncoesDiversas
    Dim oconnreq As CanadaConnectivity.ReqConsultaSelagem
    Dim sUser As String, sReq As String, sResp As String, sCaptura As String
    Dim bIsCloud As Boolean, i As Integer

    Set ofuncoes = New CanadaConnectivity.FuncoesDiversas
    sUser = ofuncoes.ReturnUserWindows
    bIsCloud = IsCloud()

    If bIsCloud Then
        sUser = "SGaCloud"
        sReq = "\\tsclient\c\sga\Req\" & sUser & ".xml"
        sResp = "\\tsclient\c\sga\Resp\" & sUser & ".xml"
        sCaptura = "\\tsclient\c\sga\Resp\CAPTURA.txt"
    Else
        sReq = "C:\Sga\Req\" & sUser & ".xml"
        sResp = "C:\Sga\Resp\" & sUser & ".xml"
        sCaptura = "C:\Sga\Resp\CAPTURA.txt"
    End If

    ofuncoes.ExcluiArquivos sReq
    ofuncoes.ExcluiArquivos sResp

    Set oconnreq = New CanadaConnectivity.ReqConsultaSelagem
    oconnreq.CreateConsultaSelagem wPlaca, "", "", False, _
        Retorna_Links_Uteis(6), Retorna_Links_Uteis(10), 6, "www.canainfo.com.br"

    ' assim que o resp aparecer, copia IMEDIATAMENTE (antes de qualquer Load)
    Dim capturou As Boolean
    capturou = False
    For i = 1 To 240  ' ~120s
        If Dir(sResp) <> "" Then
            On Error Resume Next
            Dim conteudo As String
            conteudo = LeTexto(sResp)
            If Len(conteudo) > 0 Then
                EscreveTexto sCaptura, conteudo
                capturou = True
                Exit For
            End If
            On Error GoTo 0
        End If
        Sleep 250
        DoEvents
    Next i

    If Not capturou Then
        MsgBox "O resp nao apareceu. Host/extensao rodando?", vbCritical
        Exit Sub
    End If

    MsgBox "Capturei o resp em:" & vbCrLf & sCaptura & vbCrLf & vbCrLf & _
           "Tamanho: " & Len(conteudo) & " chars" & vbCrLf & vbCrLf & _
           "---- CONTEUDO ----" & vbCrLf & conteudo, vbInformation, "Captura Resp"
End Sub

'===========================================================================
'  Diagnostico_Final
'---------------------------------------------------------------------------
'  Teste definitivo: faz a consulta, captura o resp assim que aparece, e
'  JA na sequencia (com o arquivo garantidamente presente) tenta o
'  Load_Objetos nas variantes de formato - tudo de uma vez.
'
'  Assim elimina a duvida do "arquivo sumiu antes de testar".
'
'  Rode:  Diagnostico_Final "CVN8588"
'===========================================================================
Public Sub Diagnostico_Final(Optional ByVal wPlaca As String = "CVN8588")
    Dim ofuncoes As CanadaConnectivity.FuncoesDiversas
    Dim oconnreq As CanadaConnectivity.ReqConsultaSelagem
    Dim sUser As String, sReq As String, sResp As String, sMsg As String
    Dim bIsCloud As Boolean, i As Integer, sOrig As String

    Set ofuncoes = New CanadaConnectivity.FuncoesDiversas
    sUser = ofuncoes.ReturnUserWindows
    bIsCloud = IsCloud()
    sReq = "C:\Sga\Req\" & sUser & ".xml"
    sResp = "C:\Sga\Resp\" & sUser & ".xml"

    ofuncoes.ExcluiArquivos sReq
    ofuncoes.ExcluiArquivos sResp

    Set oconnreq = New CanadaConnectivity.ReqConsultaSelagem
    oconnreq.CreateConsultaSelagem wPlaca, "", "", False, _
        Retorna_Links_Uteis(6), Retorna_Links_Uteis(10), 6, "www.canainfo.com.br"

    ' captura o resp assim que aparece
    sOrig = ""
    For i = 1 To 240
        If Dir(sResp) <> "" Then
            sOrig = LeTexto(sResp)
            If Len(sOrig) > 0 Then Exit For
        End If
        Sleep 250
        DoEvents
    Next i

    If Len(sOrig) = 0 Then
        MsgBox "resp nao apareceu.", vbCritical
        Exit Sub
    End If

    ' salva uma copia permanente para analise
    EscreveTexto "C:\Sga\Resp\CAPTURA.txt", sOrig

    sMsg = "=== DIAGNOSTICO FINAL ===" & vbCrLf
    sMsg = sMsg & "resp capturado: " & Len(sOrig) & " chars" & vbCrLf
    sMsg = sMsg & "IsCloud=" & bIsCloud & "  user=[" & sUser & "]" & vbCrLf & vbCrLf

    ' agora testa variantes COM o arquivo presente
    sMsg = sMsg & "0) original               -> " & TentaLoad(sResp, sOrig, sUser) & vbCrLf

    Dim p As Long, s1 As String
    p = InStr(sOrig, "?>")
    If p > 0 Then s1 = Mid(sOrig, p + 2) Else s1 = sOrig
    sMsg = sMsg & "1) sem <?xml?>            -> " & TentaLoad(sResp, s1, sUser) & vbCrLf

    Dim s2 As String
    s2 = Replace(sOrig, "encoding=""utf-8""", "encoding=""ISO-8859-1""")
    s2 = Replace(s2, "encoding=""UTF-8""", "encoding=""ISO-8859-1""")
    sMsg = sMsg & "2) encoding ISO-8859-1    -> " & TentaLoad(sResp, s2, sUser) & vbCrLf

    Dim s3 As String
    s3 = Replace(s2, " xmlns=""www.canainfo.com.br""", "")
    sMsg = sMsg & "3) ISO + sem xmlns        -> " & TentaLoad(sResp, s3, sUser) & vbCrLf

    Dim s4 As String
    s4 = Replace(sOrig, " xmlns=""www.canainfo.com.br""", "")
    sMsg = sMsg & "4) sem xmlns              -> " & TentaLoad(sResp, s4, sUser) & vbCrLf

    ' restaura original
    EscreveTexto sResp, sOrig

    MsgBox sMsg, vbInformation, "Diagnostico Final"
End Sub

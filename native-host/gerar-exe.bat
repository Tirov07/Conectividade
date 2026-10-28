@echo off
REM ===========================================================================
REM  gerar-exe.bat  -  Gera o ConectividadeHost.exe (native host) e, opcional-
REM  mente, copia para C:\sga\conectividade\.
REM
REM  Como usar: de um duplo-clique neste arquivo (ou rode no terminal, dentro
REM  da pasta native-host).  Requer o .NET SDK instalado.
REM ===========================================================================

setlocal
cd /d "%~dp0"

REM Pasta de destino onde o host roda no cliente (ajuste se necessario).
set "DESTINO=C:\sga\conectividade"

echo.
echo ============================================================
echo  Gerando o ConectividadeHost.exe ...
echo ============================================================
echo.

REM Verifica se o dotnet esta disponivel.
where dotnet >nul 2>&1
if errorlevel 1 (
    echo [ERRO] O comando "dotnet" nao foi encontrado.
    echo        Instale o .NET SDK e tente novamente.
    echo.
    pause
    exit /b 1
)

REM Compila como .exe unico e autossuficiente (nao exige .NET no cliente).
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
if errorlevel 1 (
    echo.
    echo [ERRO] Falha ao gerar o .exe. Veja as mensagens acima.
    echo.
    pause
    exit /b 1
)

set "PUBLISH=%~dp0bin\Release\net8.0-windows\win-x64\publish"
set "EXE=%PUBLISH%\ConectividadeHost.exe"

echo.
echo ============================================================
echo  .EXE gerado com sucesso em:
echo    %EXE%
echo ============================================================
echo.

if not exist "%EXE%" (
    echo [AVISO] Nao encontrei o .exe no caminho esperado.
    echo         Verifique a pasta: %PUBLISH%
    echo.
    pause
    exit /b 1
)

REM Pergunta se quer copiar para a pasta de destino.
set /p "COPIAR=Copiar o .exe para %DESTINO% agora? (S/N): "
if /i "%COPIAR%"=="S" (
    if not exist "%DESTINO%" mkdir "%DESTINO%"
    copy /y "%EXE%" "%DESTINO%\ConectividadeHost.exe" >nul
    if errorlevel 1 (
        echo.
        echo [ERRO] Nao foi possivel copiar. O host pode estar em uso.
        echo        Feche o Chrome por completo e rode de novo, ou copie o
        echo        arquivo manualmente.
    ) else (
        echo.
        echo  Copiado para %DESTINO%\ConectividadeHost.exe
        echo.
        echo  PROXIMO PASSO: recarregue a extensao em chrome://extensions
        echo  (isso reinicia o host com a versao nova).
    )
) else (
    echo.
    echo  Nada copiado. Copie manualmente o arquivo acima para %DESTINO%
    echo  quando quiser.
)

echo.
pause
endlocal

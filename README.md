# Conectividade Canadá — Cronotacógrafo (Inmetro)

Substitui o **navegador à parte** da Conectividade atual por uma **extensão do
Chrome** que faz a consulta de selagem dentro do próprio navegador do cliente
(na sessão que ele já logou manualmente), integrada ao SGA por um pequeno
**native host** em C#.

> **Status:** estrutura inicial. Implementada a **Operação 6** (consulta de
> selagem por placa) ponta a ponta. Demais operações virão por etapas.

## Por que essa arquitetura

A Conectividade de hoje abre um **segundo navegador** controlado por Selenium.
Isso causa os problemas conhecidos:

| Problema atual | Causa |
|---|---|
| Máquinas fracas travando | 2º navegador consumindo recursos |
| O *service* às vezes não abre a Conectividade | processo separado e pesado no boot |
| Captcha difícil de passar | navegador automatizado é mais "suspeito" |
| Inmetro derruba por inatividade | sessão isolada que ninguém usa |

A proposta usa **o Chrome que o cliente já tem aberto e logado**:

- ✅ **Nada de 2º navegador** → alívio em máquinas fracas.
- ✅ **Captcha** some ou fica raro → é uma sessão humana real.
- ✅ **Inatividade** → é a aba que a pessoa usa.
- ✅ Visualmente **não abre nada a mais**: a extensão é invisível e o native
  host roda em background sem janela.

## Como funciona

```
   SGA (desktop)
      │  grava ENVIO.xml
      ▼
  C:\sga\req ───────────────┐
      │                      │  grava RESPOSTA.xml
      ▼                      │
 ┌──────────────────────────────────────┐
 │  Native Host (C#)  ConectividadeHost  │  ← lê req / escreve resp
 │  - vigia C:\sga\req (FileSystemWatcher)│
 │  - fala com a extensão (Native Msg)    │
 └──────────────────────────────────────┘
      ▲  Native Messaging (stdin/stdout)
      │
 ┌──────────────────────────────────────┐
 │  Extensão Chrome (Manifest V3)        │
 │  - usa a aba já logada no Inmetro      │
 │  - percorre as 3 etapas da consulta    │
 │  - devolve os dados                    │
 └──────────────────────────────────────┘
      │
      ▼
 cronotacografo.rbmlq.gov.br  (login manual do operador)
```

### As 3 etapas da consulta (site CakePHP, não é SPA)

1. **Busca** (`/selagens/consultar`): preenche a placa em `#Gruveiculo_placa` e
   envia o form `#SelagemConsultarForm`.
2. **Resultados**: lê a `table.table-list-resultados`, pega a **1ª linha** (mais
   recente) e abre o link **"ver"**.
3. **Detalhes** (`/selagens/ver...`): extrai veículo, cronotacógrafo, selagem,
   ensaio, resultado e GRU — a maior parte do XML de retorno sai daqui.

> **Por que extensão + host, e não só extensão?** Uma extensão do Chrome **não
> pode ler/escrever arquivos arbitrários** do disco (como `C:\sga\req`). O host
> faz essa parte. Ele é minúsculo, sem janela, e roda só quando a extensão o
> chama via `connectNative()`.

## Estrutura

```
Conectividade/
├── extension/                       Extensão Chrome (Manifest V3)
│   ├── manifest.json
│   ├── background.js                orquestra: host <-> abas
│   ├── content.js                   "dirige" a página (as 3 etapas)
│   ├── popup.html / popup.js        indicadores de status
│   └── icons/
├── native-host/                     Programa C# (.NET 8)
│   ├── ConectividadeHost.csproj
│   ├── Program.cs                   watcher de C:\sga\req + fluxo
│   ├── NativeMessaging.cs           protocolo stdin/stdout do Chrome
│   ├── XmlModels.cs                 classes dos XMLs (envio/retorno)
│   ├── Log.cs                       log em arquivo
│   ├── install/
│   │   ├── com.canada.conectividade.json   manifesto do host
│   │   └── registrar-host.reg              registro no Chrome
│   └── exemplos/ENVIO-exemplo.xml
└── README.md
```

## Instalação no cliente

### 1. Compilar o native host (.exe único)

```powershell
cd native-host
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
```

Copie o `ConectividadeHost.exe` gerado para `C:\sga\conectividade\`.

### 2. Instalar a extensão

Enquanto não está publicada na Chrome Web Store, carregue como extensão
descompactada:

1. Chrome → `chrome://extensions`
2. Ative o **Modo do desenvolvedor**
3. **Carregar sem compactação** → selecione a pasta `extension/`
4. **Copie o ID da extensão** que aparece (ex.: `abcdefgh...`)

### 3. Registrar o native host

1. Edite `native-host/install/com.canada.conectividade.json`:
   - `path`: caminho do `.exe` (ex.: `C:\\sga\\conectividade\\ConectividadeHost.exe`)
   - `allowed_origins`: troque `SUBSTITUA_PELO_ID_DA_EXTENSAO` pelo ID copiado no passo 2
2. Copie esse JSON para `C:\sga\conectividade\`
3. Edite e execute `registrar-host.reg` (ajuste o caminho do JSON) — ou crie a
   chave de registro equivalente via instalador.

### 4. Usar

1. Abra o Chrome e **faça login manualmente** no Cronotacógrafo.
2. O SGA grava os pedidos em `C:\sga\req` normalmente — o host faz o resto.
3. Clique no ícone da extensão para ver o status (host conectado + sessão).

## Testar sem o SGA

É possível testar o host apontando as pastas para um diretório de teste:

```powershell
$env:CONECT_REQ_DIR = "C:\temp\req"
$env:CONECT_RESP_DIR = "C:\temp\resp"
```

Copie `exemplos/ENVIO-exemplo.xml` para a pasta `req` e observe o `resp`.
Os logs ficam em `%LOCALAPPDATA%\CanadaConectividade\host.log`.

## Roadmap

- [x] Operação 6 — consulta de selagem por placa (placa/renavam/chassi), testada em máquina real
- [x] Operação 6 — "sincronizar última selagem" (é a própria Op. 6 com os 3 campos preenchidos)
- [x] Operação 13 — consulta de lote de placas (reusa o fluxo da Op. 6 por placa)
- [ ] Validar a Op. 13 numa máquina real com o SGA
- [ ] Keep-alive leve para evitar queda por inatividade
- [ ] Instalador único (host + registro + extensão)

### Operações suportadas

| Operação | XML raiz | Descrição |
|---|---|---|
| **6** | `ConsultaSelagemEnvio` | Consulta/sincroniza a última selagem de uma placa (aceita placa, renavam e/ou chassi). |
| **13** | `ConsultaLotePlacasEnvio` | Consulta um lote de placas (até ~500). Faz uma consulta por placa, sequencialmente, na tela de selagem, e devolve um `<Consulta Id="n">` por placa com `StatusConsulta` individual. A tela de certificados (`UrlConsulta`) é ignorada de propósito — passou a exigir captcha por placa; a `DataValidade` vem do campo "Informações do resultado da verificação" da tela de selagem. |
```

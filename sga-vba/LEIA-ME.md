# Ajustes no SGA (VBA) para usar a extensão Conectividade

> Esta pasta faz parte do projeto **Conectividade** (extensão Chrome + native
> host). São os ajustes do lado do **SGA (Access/VBA)** necessários para o SGA
> parar de depender do app antigo e passar a conversar com a **nova extensão**,
> que usa as mesmas pastas `C:\Sga\Req` e `C:\Sga\Resp`.

Estes 11 módulos `.bas` (+ `Funcoes_Diversas.bas`) são as versões **ajustadas**
dos módulos de integração
SGA ↔ Inmetro para funcionarem com a **nova extensão Chrome** (que substitui a
Conectividade antiga). Eles usam as **mesmas pastas** `C:\Sga\Req` e
`C:\Sga\Resp`, então a gravação do pedido e a leitura da resposta **não mudaram**.

## O que foi alterado (e só isso)

Em cada módulo, apenas **dois trechos** foram desativados (comentados com `'`),
todos marcados com o comentário `'[CONECTIVIDADE_NOVA]` para você localizar:

1. **Validação do app antigo** — o bloco que impedia prosseguir quando a
   Conectividade antiga não estava aberta:
   ```vba
   If ofuncoes.ConectividadeIsAlive("Conectividade") = False Then
      MsgBox "CONECTIVIDADE NAO ESTA ATIVA, IMPOSSIVEL PROSSEGUIR.", vbCritical
      ... = False
      Exit Function
   End If
   ```
   Como a nova extensão roda dentro do Chrome que você já deixa logado, essa
   checagem não faz mais sentido — por isso foi desativada.

2. **Ativação da janela do app antigo** — a linha que trazia a janela
   "Conectividade - login" para frente:
   ```vba
   AppActivate ("Conectividade - " & RetornaLoginInMetro)
   ```
   Essa janela não existe mais; a linha daria erro, por isso foi comentada.

## O que foi MANTIDO de propósito

- `AppActivate "SGa"` — continua trazendo o SGA de volta (válido).
- `Shell "C:\SGA\CanainfoPrint.exe"` (no módulo CRLV) — é outro aplicativo
  (impressão de CRLV), **não** a Conectividade. Mantido.
- Gravação em `C:\Sga\Req` e leitura de `C:\Sga\Resp` — a extensão usa as
  mesmas pastas. Mantido.
- A checagem de nuvem `ExisteDiretorio(CaminhoSGaRemote)` — continua valendo
  para o modo nuvem (`IsCloud = True`). Mantida.

## Como aplicar no SGA

Para **cada** arquivo desta pasta:

1. Abra o SGA → Editor do VBA (`Alt + F11`).
2. No painel de módulos, localize o módulo de mesmo nome (ex.:
   `Integracao_SGA_Inmetro_Informar_Selagem`).
3. **Remova** o módulo antigo: clique com o botão direito → *Remover* →
   quando perguntar se quer exportar antes, **exporte** (backup) e então remova.
4. **Importe** o novo: menu *Arquivo → Importar Arquivo…* → selecione o `.bas`
   desta pasta.
5. Repita para os 11 módulos (ou só os que você usa — ver abaixo).
6. Salve/compile (`Debug → Compilar`) para confirmar que não há erro.

> Dica: você pode começar só pelos 3 que já usa no dia a dia:
> `Informar_Selagem` (Op. 3), `Consulta_Selagem_Completa` (Op. 6) e
> `Consulta_Selagem_Completa_Lote_Placas_NEw` (Op. 13). Os demais são para
> quando migrar as outras operações.

## Lista dos módulos ajustados

| Módulo | Operação | Alterações |
|---|---|---|
| Integracao_SGA_Inmetro_Informar_Selagem | Informar selagem (3) | 1 validação + 1 AppActivate |
| Integracao_SGA_Inmetro_Consulta_Selagem_Completa | Consulta selagem (6) | 1 + 1 |
| Integracao_SGA_Inmetro_Consulta_Selagem_Completa_Lote_Placas_NEw | Lote de placas (13) | 2 + 2 |
| Integracao_SGA_Inmetro_Emitir_GRU | Emitir GRU | 1 + 1 |
| Integracao_SGA_Inmetro_Informar_Ensaio | Informar ensaio | 1 + 1 |
| Integracao_SGA_Inmetro_Cadastrar_Veiculo | Cadastrar veículo | 1 + 1 |
| Integracao_SGA_Inmetro_Download_Certificado | Download certificado | 1 + 2 |
| Integracao_SGA_Inmetro_CRLV | CRLV | 1 + 2 |
| Integracao_SGA_Inmetro_Atualizar_Dados | Atualizar marcas/modelos/CNPJ | 5 + 5 |
| Integracao_SGA_Inmetro_Informar_Substituicao | Informar substituição | 1 + 1 |
| Integracao_SGA_Inmetro_Consulta_Selagem_Substuicao | Consulta substituição | 1 + 2 |

**Total:** 16 validações e 19 ativações de janela desativadas.

## Abertura automática do app antigo no login (IMPORTANTE)

Além dos 11 módulos acima, o SGA abria o aplicativo antigo
`C:\SGA\conectividade.exe` **automaticamente no login** (e na troca de empresa).
Isso ficava na função `Ativa_Conectividade_Autoexec`, dentro de
`Funcoes_Diversas.bas`, chamada pelos formulários `Form_Logon` e
`Form_DialogoSelecaoAcessoEmpresa`.

Esse arquivo (`Funcoes_Diversas.bas`) também foi incluído aqui, com **apenas a
função `Ativa_Conectividade_Autoexec` esvaziada** (comentada). Resultado:

- O `conectividade.exe` antigo **não abre mais** quando o cliente entra no SGA.
- As 6 chamadas à função espalhadas nos formulários **continuam existindo** e
  não quebram — simplesmente não abrem mais nada.

> ⚠️ `Funcoes_Diversas.bas` é um módulo **enorme** (32 mil linhas) e cheio de
> outras funções usadas pelo sistema todo. Se preferir **não** reimportar o
> arquivo inteiro, você pode apenas abrir a função `Ativa_Conectividade_Autoexec`
> no seu SGA e comentar o corpo dela manualmente (é rápido, são ~10 linhas). O
> efeito é idêntico.

### O que NÃO foi mexido no login (de propósito)
- `GeraArquivoLoginInmetro` — gera um XML de login na pasta `Req`. **Não é
  bloqueante** (não fica esperando resposta), então não trava o SGA mesmo que a
  extensão ignore esse XML. Deixado como está.
- `ConectividadeAtiva(empresa)` — é apenas o **flag de cadastro** da empresa
  ("usa Inmetro? S/N"). **Mantenha ligado (S)**: é ele que habilita os menus e
  botões do Inmetro no SGA. Não tem relação com o app antigo.
- `Shell "...Canada.Pagamentos.app.exe"` — é o app de **Pagamentos**, outro
  sistema. Mantido.

## Como foram gerados

A partir dos módulos originais exportados do SGA, rodando o script
`aplicar_conectividade_nova.py` (incluído nesta pasta):

```
python aplicar_conectividade_nova.py <pasta_com_os_bas_originais> <pasta_destino>
```

Se o VBA do seu SGA estiver um pouco diferente, basta reexportar os `.bas`
originais, rodar o script de novo e reimportar — ele encontra e comenta os dois
padrões automaticamente.

> Observação: a função `Ativa_Conectividade_Autoexec` (que abria o
> `C:\SGA\conectividade.exe` no login) está dentro de `Funcoes_Diversas.bas` e
> **não** é tratada por este script — ela foi ajustada manualmente. Veja a seção
> "Abertura automática do app antigo no login" acima.

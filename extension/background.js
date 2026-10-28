/**
 * background.js — Service worker da extensão Conectividade Canadá.
 *
 * Papel: ponte entre o NATIVE HOST (programa C# que lê C:\sga\req e escreve
 * C:\sga\resp) e a PÁGINA do Cronotacógrafo (onde a consulta acontece de fato,
 * dentro da sessão já logada manualmente pelo usuário).
 *
 * Fluxo de uma consulta (Operação 6 — consulta de selagem por placa):
 *   1. Native host detecta ENVIO.xml em C:\sga\req, extrai a placa e manda
 *      aqui uma mensagem { tipo: "consultar_selagem", placa, chassi, renavam, id }.
 *   2. Aqui a gente encontra (ou abre) a aba do cronotacografo e delega ao
 *      content.js, que "dirige" a página pelas 3 etapas e devolve os dados.
 *   3. A gente repassa o resultado de volta pro native host, que monta o
 *      RETORNO.xml e grava em C:\sga\resp.
 *
 * IMPORTANTE: a extensão NÃO lê/escreve arquivos do disco (o Chrome não
 * permite). Quem faz isso é o native host. A extensão só sabe "consultar placa
 * e devolver dados".
 */

// Deve bater com o "name" no manifesto do native host (com.canada.conectividade.json).
const NATIVE_HOST_NAME = "com.canada.conectividade";
const BASE_URL = "https://cronotacografo.rbmlq.gov.br";
const URL_CONSULTA_SELAGEM = `${BASE_URL}/selagens/consultar`;

let porta = null; // conexão persistente com o native host

// Consultas em andamento, aguardando o content.js terminar as 3 etapas.
// chave = id do pedido (vindo do native host) -> { resolve }
const pendentes = new Map();

/**
 * Conecta no native host e mantém a conexão viva. Se cair, tenta reconectar.
 */
function conectarNativeHost() {
  if (porta) return porta;

  try {
    porta = chrome.runtime.connectNative(NATIVE_HOST_NAME);
    log("Conectado ao native host:", NATIVE_HOST_NAME);

    porta.onMessage.addListener(onMensagemDoHost);

    porta.onDisconnect.addListener(() => {
      const erro = chrome.runtime.lastError;
      log("Native host desconectou.", erro ? erro.message : "(sem erro)");
      porta = null;
      // Reconecta depois de um tempinho — o SGA pode abrir/fechar o host.
      setTimeout(conectarNativeHost, 3000);
    });
  } catch (e) {
    log("Falha ao conectar no native host:", e.message);
    porta = null;
    setTimeout(conectarNativeHost, 5000);
  }

  return porta;
}

/**
 * Recebe os pedidos vindos do native host.
 */
async function onMensagemDoHost(msg) {
  log("Mensagem do host:", msg);

  if (!msg || !msg.tipo) return;

  if (msg.tipo === "consultar_selagem") {
    const resultado = await consultarSelagem(msg);
    // Devolve pro host casar com o id do pedido e montar o RETORNO.xml.
    // (resultado traz status, dados?, message e naoEncontrado)
    responderHost({ tipo: "resultado_selagem", id: msg.id, ...resultado });
    return;
  }

  if (msg.tipo === "informar_selagem") {
    const resultado = await informarSelagem(msg);
    responderHost({ tipo: "resultado_informar", id: msg.id, ...resultado });
    return;
  }

  if (msg.tipo === "progresso_lote") {
    atualizarProgressoLote(msg);
    return;
  }

  if (msg.tipo === "ping") {
    responderHost({ tipo: "pong" });
    return;
  }
}

/**
 * Orquestra a operação 3 (informar selagem). Igual à consulta, atravessa
 * recarregamentos de página; a resposta volta via runtime.sendMessage do
 * informar.js com o id da operação.
 */
async function informarSelagem(pedido) {
  try {
    const aba = await garantirAbaCronotacografo();

    const resultado = await new Promise((resolve, reject) => {
      const idConsulta = `i-${Date.now()}-${Math.random().toString(36).slice(2)}`;

      // Informar selagem pode demorar (3 AJAX encadeados + preenchimento).
      const timeout = setTimeout(() => {
        pendentes.delete(idConsulta);
        reject(new Error("Tempo esgotado ao informar a selagem."));
      }, 115000);

      pendentes.set(idConsulta, { resolve, timeout });

      enviarParaContentScript(aba.id, {
        acao: "informar_selagem",
        idConsulta,
        veiculo: pedido.veiculo,
        equipamento: pedido.equipamento,
        selagem: pedido.selagem
      }).catch((e) => {
        clearTimeout(timeout);
        pendentes.delete(idConsulta);
        reject(e);
      });
    });

    return resultado;
  } catch (e) {
    log("Erro ao informar selagem:", e.message);
    return { status: false, message: `Erro ao informar selagem: ${e.message}` };
  }
}

/**
 * Atualiza o indicador de progresso do lote: um badge no ícone da extensão e
 * um overlay flutuante nas abas do Cronotacógrafo.
 */
// Guarda o último progresso conhecido do lote, para reexibir o overlay quando
// a página recarrega durante o lote (cada consulta recarrega a aba).
let ultimoProgressoLote = null;

async function atualizarProgressoLote(msg) {
  const { atual, total, placa, concluido } = msg;

  ultimoProgressoLote = concluido ? null : { atual, total, placa };

  // 1) Badge no ícone da extensão (visível mesmo fora da aba do site).
  try {
    if (concluido) {
      await chrome.action.setBadgeText({ text: "" });
    } else {
      await chrome.action.setBadgeBackgroundColor({ color: "#2d7" });
      await chrome.action.setBadgeText({ text: `${atual}` });
    }
  } catch (e) {
    log("Falha ao atualizar badge:", e.message);
  }

  // O progresso detalhado é exibido no POPUP da extensão (janelinha do ícone),
  // que consulta `ultimoProgressoLote` via __pedir_progresso_lote. Não usamos
  // mais overlay flutuante na página.
}

/**
 * Orquestra a consulta de uma placa: garante uma aba no cronotacografo e pede
 * ao content.js pra executar as 3 etapas. Em caso de erro (ex.: sessão não
 * logada, captcha), devolve STATUS=false com a mensagem.
 */
async function consultarSelagem(pedido) {
  try {
    const aba = await garantirAbaCronotacografo();

    // A consulta atravessa recarregamentos de página (3 etapas). Por isso a
    // resposta NÃO volta pelo retorno do sendMessage, e sim por um
    // runtime.sendMessage posterior do content.js (com o id da consulta).
    // Montamos uma Promise que o ouvinte de mensagens resolve quando chegar.
    const resultado = await new Promise((resolve, reject) => {
      const idConsulta = `c-${Date.now()}-${Math.random().toString(36).slice(2)}`;

      // 70s: um pouco abaixo dos 75s que o host espera por placa, para que a
      // extensão responda (com erro) ANTES de o host desistir — assim, num
      // lote, a placa vira StatusConsulta=False e a consulta segue para a
      // próxima em vez de o host estourar o tempo sozinho.
      const timeout = setTimeout(() => {
        pendentes.delete(idConsulta);
        reject(new Error("Tempo esgotado aguardando a resposta da consulta."));
      }, 70000);

      pendentes.set(idConsulta, { resolve, timeout });

      enviarParaContentScript(aba.id, {
        acao: "consultar_selagem",
        idConsulta,
        placa: (pedido.placa || "").trim(),
        chassi: (pedido.chassi || "").trim(),
        renavam: (pedido.renavam || "").trim()
      }).catch((e) => {
        clearTimeout(timeout);
        pendentes.delete(idConsulta);
        reject(e);
      });
    });

    return resultado; // { status, dados?, message }
  } catch (e) {
    log("Erro na consulta:", e.message);
    return {
      status: false,
      message: `Erro ao consultar no navegador: ${e.message}`
    };
  }
}

/**
 * Encontra uma aba já aberta no cronotacografo (pra reusar a sessão logada) ou
 * abre uma nova na página de consulta de selagem.
 */
async function garantirAbaCronotacografo() {
  const abas = await chrome.tabs.query({ url: `${BASE_URL}/*` });

  if (abas.length > 0) {
    // Reusa a primeira aba logada do cliente.
    return abas[0];
  }

  // Nenhuma aba aberta — abre a página de consulta.
  // (Se a sessão não estiver logada, o site redireciona pro login; o
  // content.js detecta isso e devolve STATUS=false pedindo login manual.)
  const nova = await chrome.tabs.create({
    url: URL_CONSULTA_SELAGEM,
    active: false
  });

  // Espera a aba terminar de carregar.
  await esperarAbaCarregar(nova.id);
  return nova;
}

function esperarAbaCarregar(tabId) {
  return new Promise((resolve) => {
    function ouvinte(id, info) {
      if (id === tabId && info.status === "complete") {
        chrome.tabs.onUpdated.removeListener(ouvinte);
        resolve();
      }
    }
    chrome.tabs.onUpdated.addListener(ouvinte);
  });
}

/**
 * Manda uma mensagem pro content.js daquela aba e espera a resposta.
 * Se o content.js ainda não estiver injetado (ex.: aba recém-aberta), injeta.
 */
async function enviarParaContentScript(tabId, mensagem) {
  try {
    return await chrome.tabs.sendMessage(tabId, mensagem);
  } catch (e) {
    // Content script ainda não presente — injeta e tenta de novo.
    log("Injetando content.js na aba", tabId);
    await chrome.scripting.executeScript({
      target: { tabId },
      files: ["content.js"]
    });
    return await chrome.tabs.sendMessage(tabId, mensagem);
  }
}

// Responde ao popup com o status da conexão com o native host.
chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  if (msg && msg.__popup === "status_host") {
    sendResponse({ conectado: !!porta });
    return true;
  }
  // content.js pergunta, ao carregar, se há um lote em andamento (para
  // reexibir o overlay após o recarregamento de página).
  if (msg && msg.__pedir_progresso_lote) {
    sendResponse({ progresso: ultimoProgressoLote });
    return true;
  }
});

// Recebe o resultado final que o content.js envia após percorrer as 3 etapas.
chrome.runtime.onMessage.addListener((msg) => {
  if (msg && msg.__conectividade_resultado && msg.idConsulta) {
    const pend = pendentes.get(msg.idConsulta);
    if (pend) {
      clearTimeout(pend.timeout);
      pendentes.delete(msg.idConsulta);
      pend.resolve({
        status: !!msg.status,
        dados: msg.dados,
        message: msg.message || "",
        naoEncontrado: !!msg.naoEncontrado
      });
    }
  }
});

function responderHost(obj) {
  if (!porta) {
    log("Sem conexão com o host; não foi possível responder:", obj);
    return;
  }
  porta.postMessage(obj);
}

function log(...args) {
  console.log("[Conectividade]", ...args);
}

// Conecta assim que o service worker sobe.
conectarNativeHost();

// Reconecta quando o Chrome reinicia a extensão.
chrome.runtime.onStartup.addListener(conectarNativeHost);
chrome.runtime.onInstalled.addListener(conectarNativeHost);

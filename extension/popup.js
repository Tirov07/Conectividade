/**
 * popup.js — mostra ao cliente dois indicadores de status:
 *   1. se o native host (comunicação com o SGA) está conectado;
 *   2. se há uma aba logada no Cronotacógrafo.
 */

const BASE_URL = "https://cronotacografo.rbmlq.gov.br";

async function atualizar() {
  // Status da sessão: existe alguma aba aberta no cronotacografo?
  const abas = await chrome.tabs.query({ url: `${BASE_URL}/*` });
  setDot("dot-login", abas.length > 0);

  // Status do host: pergunta ao background.
  try {
    const resp = await chrome.runtime.sendMessage({ __popup: "status_host" });
    setDot("dot-host", !!(resp && resp.conectado));
  } catch {
    setDot("dot-host", false);
  }
}

function setDot(id, ok) {
  const el = document.getElementById(id);
  el.classList.remove("ok", "off");
  el.classList.add(ok ? "ok" : "off");
}

atualizar();

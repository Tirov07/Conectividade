/**
 * popup.js — mostra ao cliente, na janelinha do ícone:
 *   1. se o native host (comunicação com o SGA) está conectado;
 *   2. se há uma aba logada no Cronotacógrafo;
 *   3. o progresso de uma consulta em lote (ex.: 30/120), quando houver.
 *
 * O popup abre ao clicar no ícone e fecha ao clicar fora (comportamento padrão
 * do Chrome). Enquanto está aberto, atualiza o progresso do lote de tempos em
 * tempos para acompanhar em tempo real.
 */

const BASE_URL = "https://cronotacografo.rbmlq.gov.br";

async function atualizarStatus() {
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

async function atualizarLote() {
  try {
    const resp = await chrome.runtime.sendMessage({ __pedir_progresso_lote: true });
    const p = resp && resp.progresso;
    const box = document.getElementById("lote");

    if (!p) {
      box.classList.remove("ativo"); // sem lote: esconde a seção
      return;
    }

    const pct = p.total > 0 ? Math.round((p.atual / p.total) * 100) : 0;
    document.getElementById("lote-contagem").textContent = `${p.atual}/${p.total} (${pct}%)`;
    document.getElementById("lote-placa").textContent = p.placa ? `— ${p.placa}` : "";
    document.getElementById("lote-barra").style.width = `${pct}%`;
    box.classList.add("ativo");
  } catch {
    /* background ocupado; tenta de novo no próximo tick */
  }
}

function setDot(id, ok) {
  const el = document.getElementById(id);
  el.classList.remove("ok", "off");
  el.classList.add(ok ? "ok" : "off");
}

// Primeira carga.
atualizarStatus();
atualizarLote();

// Enquanto o popup está aberto, atualiza o progresso do lote a cada 1s.
setInterval(atualizarLote, 1000);

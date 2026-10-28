/**
 * informar.js — Operação 3: INFORMAR SELAGEM (escrita no Inmetro).
 *
 * ⚠️ OPERAÇÃO DE ESCRITA FISCAL: grava uma selagem oficial e consome uma GRU
 * real. Por isso, por padrão, roda em MODO SEGURO: preenche TODA a tela e PARA
 * antes de clicar em "Enviar", deixando o operador conferir e enviar à mão.
 * Quando o preenchimento estiver validado, basta mudar ENVIAR_AUTOMATICO.
 *
 * Fluxo (site CakePHP com AJAX parcial no #form-ajax):
 *   Etapa 1 (busca):   preenche placa/renavam/chassi/GRU e envia -> recarrega.
 *   Etapa 2 (preencher):
 *     a) Marca  (select, dispara AJAX)  -> espera recarregar
 *     b) Modelo (select, dispara AJAX)  -> espera recarregar (faz surgir selos)
 *     c) Nº série (texto, dispara AJAX) -> espera recarregar
 *     d) ConstK, TipoRegistro, Escala, Ano(se válido), PAM
 *     e) Selos plásticos e adesivos (tipo + número) nos campos que existirem
 *     f) Radio do selo plástico da caixa/sensor (SeloPlasticoCaixaSensor)
 *     g) Vel. indicada + checkbox "Atesto"
 *     h) MODO SEGURO: para aqui. (ou Enviar, se ENVIAR_AUTOMATICO)
 *
 * O estado atravessa o recarregamento da busca via sessionStorage.
 */

(() => {
  "use strict";

  // 🔒 INTERRUPTOR DE SEGURANÇA.
  // false = preenche a tela e PARA (operador confere e clica Enviar).
  // true  = clica "Enviar" automaticamente (só ligar após validar MUITO).
  const ENVIAR_AUTOMATICO = false;

  const CC = window.CC;
  const CHAVE = "__conectividade_informar__";
  const URL_INFORMAR = "https://cronotacografo.rbmlq.gov.br/selagens/informar";

  // ---- Estado persistido entre recarregamentos ----------------------------

  function salvar(estado) {
    sessionStorage.setItem(CHAVE, JSON.stringify(estado));
  }
  function ler() {
    try {
      return JSON.parse(sessionStorage.getItem(CHAVE));
    } catch {
      return null;
    }
  }
  function limpar() {
    sessionStorage.removeItem(CHAVE);
  }

  // ---- Detecção de etapa --------------------------------------------------

  function ehPaginaBusca() {
    // Etapa 1: só o formulário pequeno de busca, sem os dados do veículo.
    return (
      !!document.querySelector("form#SelagemInformarForm #GruVeiculoPlaca") &&
      !document.querySelector("#SelagemInstrumentoMarcaId")
    );
  }

  function ehPaginaPreenchimento() {
    // Etapa 2: já carregou a tela grande com o select de marca do instrumento.
    return !!document.querySelector("#SelagemInstrumentoMarcaId");
  }

  // ---- Etapa 1: busca -----------------------------------------------------

  function buscar(estado) {
    const form = document.querySelector("form#SelagemInformarForm");
    const placa = document.querySelector("#GruVeiculoPlaca");
    const renavam = document.querySelector("#GruVeiculoRenavam");
    const chassi = document.querySelector("#CrVeiculoDsChassi");
    const gru = document.querySelector("#GruNumDocumento");

    if (!form) throw new Error("Formulário de busca (informar) não encontrado.");

    const v = estado.veiculo || {};
    const s = estado.selagem || {};
    if (s.gru && gru) CC.setInput(gru, s.gru);
    if (v.placa && placa) CC.setInput(placa, (v.placa || "").toUpperCase());
    if (v.renavam && renavam) CC.setInput(renavam, v.renavam);
    if (v.chassi && chassi) CC.setInput(chassi, (v.chassi || "").toUpperCase());

    estado.etapa = "preencher";
    salvar(estado);
    form.submit(); // recarrega a página inteira -> Etapa 2
  }

  // ---- Etapa 2: preenchimento (os AJAX encadeados) ------------------------

  async function preencher(estado) {
    const eq = estado.equipamento || {};
    const sel = estado.selagem || {};

    // (a) Marca -> dispara AJAX, recarrega o #form-ajax.
    const selMarca = document.querySelector("#SelagemInstrumentoMarcaId");
    if (!CC.selectByText(selMarca, eq.marca)) {
      return falhar(estado, `Marca de instrumento não encontrada na lista: ${eq.marca}`);
    }
    await CC.esperarAjax();

    // (b) Modelo -> dispara AJAX (faz surgir os campos de selo).
    const selModelo = await CC.esperarElemento("#SelagemInstrumentoModeloId");
    if (!CC.selectByText(selModelo, eq.modelo)) {
      return falhar(estado, `Modelo de instrumento não encontrado na lista: ${eq.modelo}`);
    }
    await CC.esperarAjax();

    // (c) Nº série -> dispara AJAX.
    const inpSerie = await CC.esperarElemento("#SelagemNumSerie");
    CC.setInput(inpSerie, (eq.nroSerie || "").toUpperCase());
    await CC.esperarAjax();

    // Após os 3 AJAX, o form está "estável". Reobtém os elementos (o DOM foi
    // substituído) e preenche o restante.
    const g = (id) => document.getElementById(id);

    // (d) Campos simples do instrumento.
    if (eq.constK) CC.setInput(g("SelagemConstanteK"), eq.constK);
    if (eq.tipoRegistro) CC.selectByValue(g("SelagemTpRegistro"), eq.tipoRegistro);
    if (eq.escalaVelocidade) CC.selectByValue(g("SelagemEscalaVelocidade"), eq.escalaVelocidade);
    if (eq.anoInstrumento) CC.selectByValue(g("SelagemAno"), eq.anoInstrumento); // vazio = não preenche
    if (eq.pam) CC.selectByValue(g("informarSelagemSimilar"), eq.pam);

    // Aro e Pneu do veículo (já vêm preenchidos; só reforça se vier no XML).
    if (estado.veiculo && estado.veiculo.aro) {
      CC.selectByText(g("CrVeiculoVeaId"), estado.veiculo.aro);
    }
    if (estado.veiculo && estado.veiculo.pneu) {
      CC.setInput(g("CrVeiculoDsPneu"), estado.veiculo.pneu);
    }

    // (e) Selos plásticos: cada um = { tipo (value do select), numero }.
    preencherSelos(sel.selosPlasticos || [], "SelagemSeloPlasticoDocId", "SelagemSeloPlastico");
    // Selos adesivos.
    preencherSelos(sel.selosAdesivos || [], "SelagemSeloAdesivoDocId", "SelagemSeloAdesivo");

    // (f) Radio do selo plástico da caixa/sensor: marca o radio da LINHA cujo
    // selo plástico == SeloPlasticoCaixaSensor.
    marcarRadioCaixaSensor(sel.selosPlasticos || [], sel.seloPlasticoCaixaSensor || "");

    // (g) Velocidade indicada + atesto.
    if (sel.velocidade) CC.setInput(g("SelagemVelIndicada"), sel.velocidade);
    CC.marcar(g("SelagemCondicao"), true);

    // (h) MODO SEGURO: para aqui, com a tela preenchida.
    if (!ENVIAR_AUTOMATICO) {
      finalizar(estado, {
        status: true,
        message:
          "SELAGEM PREENCHIDA (modo seguro). Confira os dados e clique em Enviar manualmente."
      });
      return;
    }

    // Envio automático (só quando habilitado). O site tem um confirm() no
    // submit; como estamos na sessão do usuário, deixamos o fluxo nativo.
    const form = document.querySelector("form#SelagemInformarForm");
    const submit = form && form.querySelector('input[type="submit"]');
    if (submit) {
      estado.etapa = "aguardando_resultado";
      salvar(estado);
      submit.click();
    } else {
      finalizar(estado, { status: false, message: "Botão Enviar não encontrado." });
    }
  }

  /**
   * Preenche uma sequência de selos nos campos numerados (1, 2, 3...).
   * idTipoPrefix: ex. "SelagemSeloPlasticoDocId" -> ...DocId1, ...DocId2
   * idSeloPrefix: ex. "SelagemSeloPlastico"       -> ...1, ...2
   */
  function preencherSelos(selos, idTipoPrefix, idSeloPrefix) {
    selos.forEach((s, i) => {
      const n = i + 1;
      const selTipo = document.getElementById(`${idTipoPrefix}${n}`);
      const inpSelo = document.getElementById(`${idSeloPrefix}${n}`);
      if (selTipo) CC.selectByValue(selTipo, s.tipo);
      if (inpSelo) CC.setInput(inpSelo, (s.numero || "").toUpperCase());
    });
  }

  /**
   * Marca o radio "Entre o sensor ou cabo e a caixa do veículo" da linha do
   * selo plástico cujo número corresponde ao SeloPlasticoCaixaSensor.
   */
  function marcarRadioCaixaSensor(plasticos, numeroCaixa) {
    if (!numeroCaixa) return;
    const alvo = numeroCaixa.toUpperCase();
    let idx = plasticos.findIndex((s) => (s.numero || "").toUpperCase() === alvo);
    if (idx < 0) idx = 0; // fallback: primeira linha
    const n = idx + 1;
    const radio = document.getElementById(`SelagemMarcaSubstit${n}1`);
    if (radio) CC.marcar(radio, true);
  }

  // ---- Finalização --------------------------------------------------------

  function finalizar(estado, resultado) {
    limpar();
    chrome.runtime.sendMessage({
      __conectividade_resultado: true,
      idConsulta: estado.idConsulta,
      ...resultado
    });
  }

  function falhar(estado, mensagem) {
    finalizar(estado, { status: false, message: mensagem });
  }

  // ---- Máquina de estados a cada carga de página --------------------------

  async function processar() {
    const estado = ler();
    if (!estado || estado.__tipo !== "informar") return;

    if (!CC.estaLogado()) {
      finalizar(estado, {
        status: false,
        message:
          "Sessão não autenticada no Cronotacógrafo. Faça login manualmente e tente novamente."
      });
      return;
    }

    try {
      if (estado.etapa === "preencher" && ehPaginaPreenchimento()) {
        await preencher(estado);
        return;
      }
      if ((!estado.etapa || estado.etapa === "busca") && ehPaginaBusca()) {
        buscar(estado);
        return;
      }
      // (etapa "aguardando_resultado": leitura do retorno será feita quando o
      //  XML de retorno da Op.3 for definido pelo cliente.)
    } catch (e) {
      finalizar(estado, { status: false, message: `Erro ao informar selagem: ${e.message}` });
    }
  }

  // ---- Entrada: mensagem do background ------------------------------------

  chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
    if (msg && msg.acao === "informar_selagem") {
      const estado = {
        __tipo: "informar",
        idConsulta: msg.idConsulta,
        veiculo: msg.veiculo || {},
        equipamento: msg.equipamento || {},
        selagem: msg.selagem || {},
        etapa: "busca"
      };
      salvar(estado);

      if (!ehPaginaBusca() && !ehPaginaPreenchimento()) {
        // Não estamos na tela de informar: vai pra ela primeiro.
        location.href = URL_INFORMAR;
        sendResponse({ status: "em_andamento" });
        return true;
      }
      processar();
      sendResponse({ status: "em_andamento" });
      return true;
    }
  });

  // A cada carga de página, continua uma operação em andamento.
  processar();
})();

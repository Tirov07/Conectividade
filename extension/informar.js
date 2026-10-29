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

  // ---- Leitura do resultado após o envio ----------------------------------

  /**
   * Após enviar a selagem, o site responde de três formas:
   *  - ERRO:    div.alert.alert-error com um <h4> e um <ul><li> detalhando.
   *  - SUCESSO: (a confirmar com o XML de retorno de sucesso do cliente)
   *             normalmente uma mensagem de confirmação / protocolo gerado.
   *  - Se ainda estiver na tela de preenchimento, o envio não concluiu.
   *
   * Retorna { pronto, status, message } — pronto=false se a página ainda não
   * trouxe um desfecho (deve aguardar).
   */
  function lerResultadoEnvio() {
    // Caso de ERRO.
    const erro = document.querySelector(".alert.alert-error, .alert-error");
    if (erro) {
      // A mensagem específica fica no <li> (dentro de ul.nav); o <h4> é o
      // título genérico ("Não foi possível informar a selagem.").
      const li = erro.querySelector("ul li");
      const h4 = erro.querySelector(".alert-heading, h4");
      const detalhe = li ? CC.limpar(li.textContent) : "";
      const titulo = h4 ? CC.limpar(h4.textContent) : "";
      const msg = detalhe || titulo || "Não foi possível informar a selagem.";
      return { pronto: true, status: false, message: msg };
    }

    // Caso de SUCESSO — box VERDE no topo (mesmo padrão do erro, mas verde),
    // com "Selagem informada com sucesso"; a selagem passa a existir com um
    // PROTOCOLO (~18 dígitos) e SITUAÇÃO ("Ativa"). O site usa Bootstrap antigo,
    // então o verde pode vir como .alert-success OU .alert-info — e, por
    // segurança, também aceitamos qualquer .alert cujo texto fale em "sucesso".
    const ok = acharAlertaSucesso();
    if (ok) {
      // Pega o texto do <h4>/<li> se houver, senão o texto do box; remove o "×".
      const h4 = ok.querySelector(".alert-heading, h4");
      const li = ok.querySelector("ul li");
      let texto = CC.limpar((h4 && h4.textContent) || (li && li.textContent) || ok.textContent || "");
      texto = texto.replace(/^×\s*/, "").trim();
      return {
        pronto: true,
        status: true,
        message: texto || "Selagem informada com sucesso.",
        protocolo: extrairProtocolo(),
        situacao: extrairSituacao()
      };
    }

    // Ainda sem desfecho visível.
    return { pronto: false };
  }

  /**
   * Acha o box de alerta de SUCESSO. O site usa Bootstrap antigo, então o
   * alerta verde pode ter classe .alert-success ou .alert-info. Como garantia
   * extra, também aceita qualquer .alert (que NÃO seja de erro) cujo texto
   * mencione "sucesso". Assim cobrimos variações sem confundir com o erro.
   */
  function acharAlertaSucesso() {
    // 1) classes verdes conhecidas
    const porClasse = document.querySelector(".alert-success, .alert.alert-info");
    if (porClasse) return porClasse;
    // 2) qualquer alerta (não-erro) que fale em "sucesso"
    for (const el of document.querySelectorAll(".alert")) {
      if (el.classList.contains("alert-error") || el.classList.contains("alert-danger")) continue;
      if (/sucesso/i.test(el.textContent || "")) return el;
    }
    return null;
  }

  /**
   * Extrai o número de protocolo da selagem recém-informada. O protocolo tem
   * ~18 dígitos (ex.: 518020160558362884). Procura por um rótulo "Protocolo"
   * seguido do número; se não achar, pega o maior número longo da página.
   * (Seletores finos serão calibrados quando tivermos o HTML da tela de
   *  sucesso; esta heurística já cobre o caso comum.)
   */
  function extrairProtocolo() {
    // 1) procura por "Protocolo" em um .fakeinput/label próximo
    const prot = CC.valorPorLabel ? CC.valorPorLabel(document, "Protocolo") : "";
    if (prot && /\d{10,}/.test(prot)) return prot.match(/\d{10,}/)[0];
    // 2) fallback: maior sequência de 15+ dígitos no texto da página
    const m = (document.body.textContent || "").match(/\d{15,}/g);
    return m ? m.sort((a, b) => b.length - a.length)[0] : "";
  }

  /**
   * Extrai a situação da selagem (ex.: "Ativa"). Procura por um rótulo
   * "Situação"; se não achar, devolve "Ativa" (situação típica de uma selagem
   * recém-informada com sucesso).
   */
  function extrairSituacao() {
    const sit = CC.valorPorLabel ? CC.valorPorLabel(document, "Situação") : "";
    return sit || "Ativa";
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
      // ERRO pode aparecer DEPOIS de uma ação nossa (busca ou envio):
      //  - logo após a busca (ex.: "A GRU informada já foi utilizada") — nesse
      //    caso a tela de preenchimento NUNCA carrega;
      //  - após o envio final.
      // Só consideramos o alerta se já passamos da busca (etapa preencher ou
      // aguardando_resultado), para não confundir com um erro remanescente.
      if (estado.etapa === "preencher" || estado.etapa === "aguardando_resultado") {
        const r = lerResultadoEnvio();
        if (r.pronto && r.status === false) {
          finalizar(estado, { status: false, message: r.message });
          return;
        }
        // Após o envio (automático): se deu sucesso, finaliza.
        if (estado.etapa === "aguardando_resultado") {
          if (r.pronto) finalizar(estado, { status: r.status, message: r.message });
          return; // sem desfecho ainda: aguarda próxima carga
        }
      }

      // Etapa 2: a tela de preenchimento carregou (busca deu certo).
      if (estado.etapa === "preencher" && ehPaginaPreenchimento()) {
        await preencher(estado);
        return;
      }

      // Etapa 2 pedida, mas a tela de preenchimento ainda NÃO apareceu.
      // Três desfechos possíveis após a busca:
      //   (1) carrega a tela de preenchimento -> segue;
      //   (2) aparece um alert de erro -> retorna o erro;
      //   (3) FICA NA TELA DE BUSCA sem erro -> GRU inválida (o site trava sem
      //       avisar nada quando o número da GRU está incorreto).
      // Fazemos uma espera ativa de ~6s cobrindo os três casos.
      if (estado.etapa === "preencher") {
        const prazoMs = 6000;
        const inicio = Date.now();
        while (Date.now() - inicio < prazoMs) {
          if (ehPaginaPreenchimento()) {
            await preencher(estado);
            return;
          }
          const r = lerResultadoEnvio();
          if (r.pronto && r.status === false) {
            finalizar(estado, { status: false, message: r.message });
            return;
          }
          await CC.sleep(300);
        }
        // Esgotou o tempo e continuamos na tela de busca, sem erro nem avanço:
        // o número da GRU é inválido.
        if (ehPaginaBusca()) {
          finalizar(estado, { status: false, message: "NÚMERO DE GRU INCORRETO." });
          return;
        }
        // Situação inesperada (nem busca, nem preenchimento, nem erro).
        finalizar(estado, {
          status: false,
          message: "NÃO FOI POSSÍVEL PROSSEGUIR COM A SELAGEM. VERIFIQUE OS DADOS."
        });
        return;
      }

      // Etapa 1: busca.
      if ((!estado.etapa || estado.etapa === "busca") && ehPaginaBusca()) {
        buscar(estado);
        return;
      }
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

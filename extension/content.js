/**
 * content.js — roda DENTRO das páginas do cronotacografo.rbmlq.gov.br.
 *
 * É o equivalente ao que o Selenium fazia, só que "por dentro" da página, na
 * sessão já logada pelo usuário. Executa as 3 etapas da consulta de selagem:
 *
 *   Etapa 1 (/selagens/consultar) — preenche a placa e clica em "Buscar".
 *   Etapa 2 (lista de resultados)  — lê a tabela, pega a 1ª linha (mais
 *                                    recente) e abre o link "ver".
 *   Etapa 3 (ver_selagem_antecipada / ver) — extrai TODOS os dados (veículo,
 *                                    cronotacógrafo, selagem, ensaio,
 *                                    resultado, GRU).
 *
 * Como o site é CakePHP clássico (form POST → página recarrega), cada etapa é
 * uma página nova. Para atravessar as etapas sem perder o controle, o content
 * script guarda o "estado da consulta" em sessionStorage e retoma a etapa certa
 * a cada carga de página.
 *
 * Seletores confirmados a partir do HTML real das 3 telas (nov/2026).
 */

(() => {
  "use strict";

  const CHAVE_ESTADO = "__conectividade_consulta__";

  // ---- Utilidades de leitura do DOM --------------------------------------

  /** Retorna o texto de um .fakeinput pelo texto do <label> que o precede. */
  function valorPorLabel(raiz, textoLabel) {
    const labels = (raiz || document).querySelectorAll("label");
    for (const label of labels) {
      if (norm(label.textContent) === norm(textoLabel)) {
        // O valor fica no .fakeinput irmão logo após o label.
        const campo = label.nextElementSibling;
        if (campo && campo.classList.contains("fakeinput")) {
          return limpar(campo.textContent);
        }
        // Às vezes está dentro do mesmo container.
        const pai = label.parentElement;
        const fake = pai && pai.querySelector(".fakeinput");
        if (fake) return limpar(fake.textContent);
      }
    }
    return "";
  }

  function norm(s) {
    return (s || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim()
      .toLowerCase();
  }

  function limpar(s) {
    return (s || "").replace(/\s+/g, " ").trim();
  }

  // ---- Estado da consulta (persistido entre recarregamentos) --------------

  function salvarEstado(estado) {
    sessionStorage.setItem(CHAVE_ESTADO, JSON.stringify(estado));
  }

  function lerEstado() {
    try {
      return JSON.parse(sessionStorage.getItem(CHAVE_ESTADO));
    } catch {
      return null;
    }
  }

  function limparEstado() {
    sessionStorage.removeItem(CHAVE_ESTADO);
  }

  // ---- Detecção de login --------------------------------------------------

  /** Se não houver sessão logada, o site mostra a tela de login. */
  function estaLogado() {
    // A navbar do usuário logado traz o bloco .user-info com o PAC.
    return !!document.querySelector(".user-info, #header-nav");
  }

  // ---- Etapa 1: busca -----------------------------------------------------

  function ehPaginaBusca() {
    return !!document.querySelector("#Gruveiculo_placa form, form#SelagemConsultarForm");
  }

  function preencherEbuscar(estado) {
    const campoPlaca = document.querySelector("#Gruveiculo_placa");
    const campoChassi = document.querySelector("#Gruveiculo_chassi");
    const campoRenavam = document.querySelector("#Gruveiculo_renavam");
    const form = document.querySelector("form#SelagemConsultarForm");

    if (!campoPlaca || !form) {
      throw new Error("Formulário de consulta não encontrado na página.");
    }

    if (estado.placa) setInput(campoPlaca, estado.placa.toUpperCase());
    if (estado.chassi && campoChassi) setInput(campoChassi, estado.chassi.toUpperCase());
    if (estado.renavam && campoRenavam) setInput(campoRenavam, estado.renavam);

    // Marca que a próxima página será a lista de resultados.
    estado.etapa = "resultados";
    salvarEstado(estado);

    form.submit();
  }

  function setInput(el, valor) {
    el.value = valor;
    el.dispatchEvent(new Event("input", { bubbles: true }));
    el.dispatchEvent(new Event("change", { bubbles: true }));
  }

  // ---- Etapa 2: lista de resultados --------------------------------------

  function ehPaginaResultados() {
    return !!document.querySelector("table.table-list-resultados");
  }

  /**
   * Lê a tabela de resultados. Retorna a lista de linhas e o link "ver" da
   * primeira (mais recente). Também captura dados que só existem na lista
   * (protocolo, selos) para completar o retorno.
   */
  function lerResultados() {
    const tabela = document.querySelector("table.table-list-resultados");
    const linhas = [...tabela.querySelectorAll("tr")].filter(
      (tr) => tr.querySelectorAll("td").length > 0
    );

    if (linhas.length === 0) {
      return { encontrou: false };
    }

    // 1ª linha de dados = selagem mais recente.
    const tr = linhas[0];
    const tds = tr.querySelectorAll("td");

    // Colunas: PROTOCOLO, RENAVAM, PLACA, GRU Nº, MARCA/MODELO, Nº SÉRIE,
    // SELOS PLÁSTICOS, SELOS ADESIVOS, DATA, POSTO, SITUAÇÃO, AÇÕES
    const lista = {
      protocolo: limpar(tds[0]?.textContent),
      renavam: limpar(tds[1]?.textContent),
      placa: limpar(tds[2]?.textContent),
      numGru: limpar(tds[3]?.textContent),
      marcaModelo: limpar(tds[4]?.textContent),
      numSerie: limpar(tds[5]?.textContent),
      selosPlasticos: juntarSelos(tds[6]),
      selosAdesivos: juntarSelos(tds[7]),
      data: limpar(tds[8]?.textContent),
      posto: limpar(tds[9]?.textContent),
      situacao: limpar(tds[10]?.textContent)
    };

    // Link "ver" (abre a Etapa 3 com todos os detalhes).
    const linkVer = tr.querySelector('a[href*="/selagens/ver"]');

    return {
      encontrou: true,
      lista,
      urlVer: linkVer ? linkVer.href : null
    };
  }

  /** Concatena os selos (vêm separados por <br>) sem separador, como no XML. */
  function juntarSelos(td) {
    if (!td) return "";
    return [...td.childNodes]
      .map((n) => (n.textContent || "").trim())
      .filter(Boolean)
      .join("");
  }

  // ---- Etapa 3: detalhes --------------------------------------------------

  function ehPaginaDetalhes() {
    return location.pathname.includes("/selagens/ver");
  }

  /** Extrai todos os blocos de dados da página de detalhes. */
  function lerDetalhes() {
    const d = {};

    // Veículo
    d.renavam = valorPorLabel(document, "RENAVAM");
    d.placa = valorPorLabel(document, "Placa");
    d.chassi = valorPorLabel(document, "Chassi");
    d.marca = valorPorLabel(document, "Marca"); // 1º "Marca" = veículo
    d.modelo = valorPorLabel(document, "Modelo");
    d.ano = valorPorLabel(document, "Ano");
    d.aro = valorPorLabel(document, "Aro");
    d.pneu = valorPorLabel(document, "Pneu");

    // Cronotacógrafo / equipamento
    d.constanteK = valorPorLabel(document, "Constante K");
    d.redutor = valorPorLabel(document, "Redutor");
    d.numSerie = valorPorLabel(document, "Nº de série");

    // Ensaio
    d.velMax = valorPorLabel(document, "Vel. máx. total (km/h)");

    // Resultado
    d.dataEmissao = valorPorLabel(document, "Data emissão");
    d.dataValidade = valorPorLabel(document, "Data validade");
    d.tipoCertificado = valorPorLabel(document, "Documento");
    d.numCertificado = valorPorLabel(document, "Doc. nº");

    // GRU
    d.numGru = valorPorLabel(document, "N° da GRU") || valorPorLabel(document, "Nº da GRU");
    d.valorGru = valorPorLabel(document, "Valor");
    d.dataDocGru = valorPorLabel(document, "Dt. documento");
    d.dataPagGru = valorPorLabel(document, "Dt. pagamento");

    // Marca/Modelo do cronotacógrafo e posto/UF da selagem são mais seguros
    // de pegar na própria tabela de selagem desta página.
    extrairEquipamentoEPostoDetalhe(d);

    return d;
  }

  /**
   * Na página de detalhes há dois blocos "Marca/Modelo": o do veículo e o do
   * cronotacógrafo. O valorPorLabel pega o primeiro; aqui refinamos o do
   * equipamento e o posto/UF da selagem usando os cabeçalhos de seção.
   */
  function extrairEquipamentoEPostoDetalhe(d) {
    const legends = [...document.querySelectorAll(".legend")];
    for (const legend of legends) {
      const titulo = norm(legend.textContent);
      const fieldset = legend.nextElementSibling;
      if (!fieldset) continue;

      if (titulo.includes("cronotacografo")) {
        d.equipMarca = valorPorLabel(fieldset, "Marca");
        d.equipModelo = valorPorLabel(fieldset, "Modelo");
        d.numSerie = valorPorLabel(fieldset, "Nº de série") || d.numSerie;
        d.constanteK = valorPorLabel(fieldset, "Constante K") || d.constanteK;
        d.redutor = valorPorLabel(fieldset, "Redutor") || d.redutor;
      }

      if (titulo.includes("ensaio")) {
        d.ufEnsaio = valorPorLabel(fieldset, "Estado");
        d.postoEnsaio = valorPorLabel(fieldset, "Posto de ensaio / verificação");
        d.dataEnsaio = valorPorLabel(fieldset, "Dt. Realizado");
      }

      if (titulo.includes("resultado da verificacao")) {
        d.ufResultado = valorPorLabel(fieldset, "Estado");
        d.respTecnico = valorPorLabel(fieldset, "Responsável técnico");
      }
    }
  }

  // ---- Máquina de estados: decide o que fazer nesta carga de página -------

  function processar() {
    const estado = lerEstado();
    if (!estado) return; // Nenhuma consulta em andamento.

    if (!estaLogado()) {
      finalizar(estado, {
        status: false,
        message:
          "Sessão não autenticada no Cronotacógrafo. Faça login manualmente e tente novamente."
      });
      return;
    }

    try {
      // Etapa 3 — detalhes
      if (ehPaginaDetalhes() && estado.etapa === "detalhes") {
        const detalhes = lerDetalhes();
        const dados = montarDados(estado, detalhes);
        finalizar(estado, { status: true, dados, message: `CONSULTA DA PLACA ${estado.placa} FINALIZADA.` });
        return;
      }

      // Etapa 2 — resultados
      if (ehPaginaResultados() && estado.etapa === "resultados") {
        const r = lerResultados();
        if (!r.encontrou) {
          finalizar(estado, {
            status: false,
            message: `Nenhuma selagem encontrada para a placa ${estado.placa}.`
          });
          return;
        }
        estado.lista = r.lista; // guarda dados da lista p/ compor o retorno
        if (r.urlVer) {
          estado.etapa = "detalhes";
          salvarEstado(estado);
          location.href = r.urlVer; // vai pra Etapa 3
        } else {
          // Sem link "ver" — retorna só o que a lista tem.
          finalizar(estado, {
            status: true,
            dados: montarDados(estado, {}),
            message: `CONSULTA DA PLACA ${estado.placa} FINALIZADA (sem detalhes).`
          });
        }
        return;
      }

      // Etapa 1 — busca
      if (ehPaginaBusca() && (!estado.etapa || estado.etapa === "busca")) {
        preencherEbuscar(estado);
        return;
      }
    } catch (e) {
      finalizar(estado, { status: false, message: `Erro ao ler a página: ${e.message}` });
    }
  }

  /** Mescla os dados da lista (Etapa 2) com os detalhes (Etapa 3). */
  function montarDados(estado, detalhes) {
    const lista = estado.lista || {};
    return {
      veiculo: {
        renavam: detalhes.renavam || lista.renavam || "",
        placa: detalhes.placa || lista.placa || estado.placa || "",
        marca: detalhes.marca || "",
        modelo: detalhes.modelo || "",
        ano: detalhes.ano || "",
        aro: detalhes.aro || "",
        pneu: detalhes.pneu || "",
        chassi: detalhes.chassi || estado.chassi || ""
      },
      equipamento: {
        marca: detalhes.equipMarca || "",
        modelo: detalhes.equipModelo || "",
        nroSerie: detalhes.numSerie || lista.numSerie || "",
        constK: detalhes.constanteK || "",
        redutor: detalhes.redutor || "-"
      },
      selagem: {
        dataSelagem: (lista.data || "").split(" - ")[0] || detalhes.dataEnsaio || "",
        postoSelagem: lista.posto || "",
        protocolo: lista.protocolo || "",
        selosAdesivos: lista.selosAdesivos || "",
        selosPlasticos: lista.selosPlasticos || "",
        situacao: lista.situacao || ""
      },
      ensaio: {
        ufEnsaio: detalhes.ufEnsaio || "",
        postoEnsaio: detalhes.postoEnsaio || "",
        dataEnsaio: detalhes.dataEnsaio || "",
        velMax: detalhes.velMax || "",
        padraoUtilizado: ""
      },
      resultado: {
        ufResultado: detalhes.ufResultado || "",
        respTecnico: detalhes.respTecnico || "",
        dataEmissao: detalhes.dataEmissao || "",
        dataValidade: detalhes.dataValidade || "",
        numCertificado: detalhes.numCertificado || "",
        tipoCertificado: detalhes.tipoCertificado || ""
      },
      gru: {
        numGru: detalhes.numGru || lista.numGru || "",
        dataPag: detalhes.dataPagGru || "",
        valorPag: "",
        valorGru: detalhes.valorGru || ""
      }
    };
  }

  /** Envia o resultado final pro background e limpa o estado. */
  function finalizar(estado, resultado) {
    limparEstado();
    // O background está esperando a resposta do sendMessage que iniciou a
    // consulta; mas como houve navegação, respondemos por runtime.sendMessage
    // com o id da consulta pra ele casar.
    chrome.runtime.sendMessage({
      __conectividade_resultado: true,
      idConsulta: estado.idConsulta,
      ...resultado
    });
  }

  // ---- Entrada: mensagens do background + auto-processamento --------------

  chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
    if (msg && msg.acao === "consultar_selagem") {
      // Inicia uma consulta nova nesta aba.
      const estado = {
        idConsulta: msg.idConsulta,
        placa: msg.placa || "",
        chassi: msg.chassi || "",
        renavam: msg.renavam || "",
        etapa: "busca"
      };

      // Se já não estamos na página de busca, vamos pra ela primeiro.
      if (!ehPaginaBusca()) {
        salvarEstado(estado);
        location.href = "https://cronotacografo.rbmlq.gov.br/selagens/consultar";
        // A resposta virá depois via runtime.sendMessage (finalizar()).
        sendResponse({ status: "em_andamento" });
        return true;
      }

      salvarEstado(estado);
      processar();
      sendResponse({ status: "em_andamento" });
      return true;
    }
  });

  // A cada carga de página, verifica se há consulta em andamento e continua.
  processar();
})();

/**
 * util.js — helpers compartilhados pelos content scripts (consulta e informar).
 *
 * Centraliza a "interação real" com a página (opção A): preencher inputs e
 * selects disparando os eventos que o site (jQuery) escuta, e esperar os
 * recarregamentos AJAX parciais do #form-ajax.
 *
 * Exposto em window.CC (namespace da extensão) para os demais scripts usarem.
 */

(() => {
  "use strict";

  if (window.CC) return; // evita redefinir se injetado duas vezes

  const CC = {};

  /** Normaliza texto: sem acento, minúsculo, trim. */
  CC.norm = function (s) {
    return (s || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim()
      .toLowerCase();
  };

  CC.limpar = function (s) {
    return (s || "").replace(/\s+/g, " ").trim();
  };

  CC.sleep = function (ms) {
    return new Promise((r) => setTimeout(r, ms));
  };

  /**
   * Preenche um <input> de texto disparando os eventos que o site escuta
   * (input, keyup, change, blur). Importante para campos com máscara jQuery.
   */
  CC.setInput = function (el, valor) {
    if (!el) return false;
    el.focus();
    el.value = valor;
    for (const tipo of ["input", "keyup", "change", "blur"]) {
      el.dispatchEvent(new Event(tipo, { bubbles: true }));
    }
    return true;
  };

  /**
   * Seleciona uma <option> de um <select> pelo TEXTO visível (case/acentos
   * tolerantes) e dispara 'change'. Retorna true se encontrou a opção.
   * Também atualiza o bootstrap-select (selectpicker) se presente.
   */
  CC.selectByText = function (selectEl, texto) {
    if (!selectEl) return false;
    const alvo = CC.norm(texto);
    let achou = false;
    for (const opt of selectEl.options) {
      if (CC.norm(opt.textContent) === alvo) {
        selectEl.value = opt.value;
        achou = true;
        break;
      }
    }
    if (!achou) {
      // tentativa mais frouxa: contém o texto
      for (const opt of selectEl.options) {
        if (CC.norm(opt.textContent).includes(alvo) && alvo.length > 0) {
          selectEl.value = opt.value;
          achou = true;
          break;
        }
      }
    }
    if (achou) CC.dispararChangeSelect(selectEl);
    return achou;
  };

  /** Seleciona uma <option> pelo VALUE e dispara 'change'. */
  CC.selectByValue = function (selectEl, value) {
    if (!selectEl) return false;
    const existe = [...selectEl.options].some((o) => o.value === String(value));
    if (!existe) return false;
    selectEl.value = String(value);
    CC.dispararChangeSelect(selectEl);
    return true;
  };

  /** Dispara 'change' no select e sincroniza o bootstrap-select (se houver). */
  CC.dispararChangeSelect = function (selectEl) {
    selectEl.dispatchEvent(new Event("change", { bubbles: true }));
    // Se o site usa bootstrap-select (selectpicker via jQuery), atualiza o
    // botão visual para refletir o valor escolhido.
    try {
      if (window.jQuery && window.jQuery(selectEl).selectpicker) {
        window.jQuery(selectEl).selectpicker("refresh");
      }
    } catch {
      /* sem jQuery/selectpicker; o change nativo basta */
    }
  };

  /** Marca um checkbox/radio e dispara os eventos. */
  CC.marcar = function (el, marcado = true) {
    if (!el) return false;
    el.checked = !!marcado;
    el.dispatchEvent(new Event("click", { bubbles: true }));
    el.dispatchEvent(new Event("change", { bubbles: true }));
    return true;
  };

  /**
   * Espera o overlay de carregamento (#loading) do site aparecer e sumir —
   * é o sinal de que um AJAX parcial começou e terminou. Com fallback por
   * tempo caso o site não use o #loading naquele passo.
   */
  CC.esperarAjax = async function (timeoutMs = 20000) {
    const loading = () => document.querySelector("#loading");
    const visivel = () => {
      const el = loading();
      if (!el) return false;
      const st = getComputedStyle(el);
      return st.display !== "none" && st.visibility !== "hidden";
    };

    const inicio = Date.now();
    // 1) espera aparecer (até 1,5s) — alguns AJAX são rápidos e nem mostram.
    while (!visivel() && Date.now() - inicio < 1500) await CC.sleep(50);
    // 2) espera sumir (até timeout).
    while (visivel() && Date.now() - inicio < timeoutMs) await CC.sleep(80);
    // 3) margem para o DOM do #form-ajax ser substituído.
    await CC.sleep(300);
  };

  /**
   * Espera um elemento aparecer no DOM (após um AJAX que troca o #form-ajax).
   * Retorna o elemento ou null no timeout.
   */
  CC.esperarElemento = async function (selector, timeoutMs = 20000) {
    const inicio = Date.now();
    while (Date.now() - inicio < timeoutMs) {
      const el = document.querySelector(selector);
      if (el) return el;
      await CC.sleep(100);
    }
    return null;
  };

  /** Detecta se há sessão logada (navbar do usuário). */
  CC.estaLogado = function () {
    return !!document.querySelector(".user-info, #header-nav");
  };

  window.CC = CC;
})();

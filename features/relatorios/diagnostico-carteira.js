(function setupDiagnosticoCarteira(global, document) {
  "use strict";

  const DEFAULT_TITLE = "Diagnostico Executivo da Carteira";
  const CRITICAL_WINDOW_DAYS = 7;
  const summaryFields = {
    total: "diagResumoTotal",
    planejadas: "diagResumoPlanejadas",
    aPlanejar: "diagResumoAPlanejar",
    replanejadas: "diagResumoReplanejadas",
    realizadas: "diagResumoRealizadas",
    vencidas: "diagResumoVencidas",
    criticas: "diagResumoCriticas",
    comKm: "diagResumoComKm",
  };
  const sectionFields = {
    capa: "diagSecaoCapa",
    resumo: "diagSecaoResumoExecutivo",
    kpis: "diagSecaoKpis",
    prioridades: "diagSecaoPrioridades",
    planejamento: "diagSecaoPlanejamento",
    centros: "diagSecaoCentros",
    planejadores: "diagSecaoPlanejadores",
    km: "diagSecaoKm",
    patios: "diagSecaoPatios",
    tolerancias: "diagSecaoTolerancias",
    vencimentos: "diagSecaoVencimentos",
    clima: "diagSecaoClima",
    listaCritica: "diagSecaoListaCritica",
    recomendacoes: "diagSecaoRecomendacoes",
  };
  const optionFields = {
    incluirGraficos: "diagIncluirGraficos",
    incluirTabelaOms: "diagIncluirTabelaOms",
    priorizarDemandasCriticas: "diagIncluirSomenteCriticas",
    destacarDemandasSemKm: "diagIncluirDemandasSemKm",
    incluirDetalhamentoCompleto: "diagIncluirDetalhamentoCompleto",
  };

  let diagnosticoContext = {};
  let handlersBound = false;
  let diagnosticoGenerating = false;

  function $(selector) {
    return document.getElementById(selector);
  }

  function safeDemandas() {
    const demandas = diagnosticoContext.getDemandasFiltradas?.();
    return Array.isArray(demandas) ? demandas : [];
  }

  function formatNumber(value) {
    return new Intl.NumberFormat("pt-BR").format(Number(value) || 0);
  }

  function setText(id, value) {
    const node = $(id);
    if (node) node.textContent = formatNumber(value);
  }

  function trimValue(value) {
    return String(value || "").trim();
  }

  function hasKm(demand) {
    return Boolean(
      trimValue(demand?.kmInicio) ||
        trimValue(demand?.kmFim) ||
        trimValue(demand?.km) ||
        trimValue(demand?.KmInicio) ||
        trimValue(demand?.KmFim),
    );
  }

  function isClosedStatus(status) {
    return status === "Realizado" || status === "Cancelado";
  }

  function addDays(date, amount) {
    const result = new Date(date.getTime());
    result.setDate(result.getDate() + amount);
    return result;
  }

  function startOfDay(date) {
    const normalized = new Date(date.getTime());
    normalized.setHours(0, 0, 0, 0);
    return normalized;
  }

  function isOverdue(demand, today) {
    const status = diagnosticoContext.primaryStatusOf?.(demand) || "";
    if (isClosedStatus(status)) return false;
    const dueDateRaw = diagnosticoContext.toDate?.(demand?.vencimento);
    const dueDate = dueDateRaw ? startOfDay(dueDateRaw) : null;
    return Boolean(dueDate && dueDate < today);
  }

  function isCritical(demand, today, criticalLimit) {
    const status = diagnosticoContext.primaryStatusOf?.(demand) || "";
    if (isClosedStatus(status)) return false;
    const dueDateRaw = diagnosticoContext.toDate?.(demand?.vencimento);
    const dueDate = dueDateRaw ? startOfDay(dueDateRaw) : null;
    return Boolean(dueDate && dueDate <= criticalLimit);
  }

  function buildSummary(demandas) {
    const today = startOfDay(new Date());
    const criticalLimit = addDays(today, CRITICAL_WINDOW_DAYS);
    const summary = {
      total: demandas.length,
      planejadas: 0,
      aPlanejar: 0,
      replanejadas: 0,
      realizadas: 0,
      vencidas: 0,
      criticas: 0,
      comKm: 0,
    };

    demandas.forEach((demand) => {
      const status = diagnosticoContext.primaryStatusOf?.(demand) || "";

      if (status === "Planejado") summary.planejadas += 1;
      if (status === "A Planejar") summary.aPlanejar += 1;
      if (status === "Replanejado") summary.replanejadas += 1;
      if (status === "Realizado") summary.realizadas += 1;
      if (isOverdue(demand, today)) summary.vencidas += 1;
      if (isCritical(demand, today, criticalLimit)) summary.criticas += 1;
      if (hasKm(demand)) summary.comKm += 1;
    });

    return summary;
  }

  function renderSummary(summary) {
    Object.entries(summaryFields).forEach(([key, id]) => {
      setText(id, summary[key] || 0);
    });
  }

  function currentUserLabel() {
    const state = diagnosticoContext.getState?.() || {};
    const user = state.currentUser || {};
    return trimValue(user.nome) || trimValue(user.email) || "";
  }

  function ensureDefaultFields() {
    const titleField = $("diagTituloRelatorio");
    const responsibleField = $("diagResponsavel");
    const climateHint = $("diagClimateHint");
    const tipoRelatorio = $("diagTipoRelatorio");

    if (titleField && !trimValue(titleField.value)) {
      titleField.value = DEFAULT_TITLE;
    }

    if (tipoRelatorio && !trimValue(tipoRelatorio.value)) {
      tipoRelatorio.value = "executivo";
    }

    if (responsibleField) {
      responsibleField.value = currentUserLabel();
    }

    if (climateHint) {
      const hasClimateRuntime = Boolean(
        global.CCEClimate || global.CCEClimaFeature || global.CCEClimateConfig,
      );
      climateHint.classList.toggle("hidden", hasClimateRuntime);
    }
  }

  function updateReportTypeDescription(preset) {
    const el = $("diagTipoRelatorioDescricao");
    if (!el) return;

    el.textContent = preset?.descricao || "";
    el.classList.toggle("hidden", !preset?.descricao);
  }

  function applyReportTypePreset(tipoRelatorio, options = {}) {
    const presets = global.CCEDiagnosticoConfig?.REPORT_SECTION_PRESETS || {};
    const preset = presets[tipoRelatorio];

    if (!preset) {
      updateReportTypeDescription(null);
      return;
    }

    Object.entries(preset.secoes || {}).forEach(([fieldId, checked]) => {
      const input = $(fieldId);
      if (input) input.checked = Boolean(checked);
    });

    Object.entries(preset.opcoes || {}).forEach(([fieldId, checked]) => {
      const input = $(fieldId);
      if (input) input.checked = Boolean(checked);
    });

    const tituloInput = $("diagTituloRelatorio");
    if (tituloInput && (!trimValue(tituloInput.value) || options.forceTitle)) {
      tituloInput.value = preset.tituloPadrao || tituloInput.value;
    }

    updateReportTypeDescription(preset);
  }

  function openDialog(dialog) {
    if (!dialog) return;
    if (typeof dialog.showModal === "function") {
      if (!dialog.open) dialog.showModal();
      return;
    }
    dialog.setAttribute("open", "open");
  }

  function closeDialog(dialog) {
    if (!dialog) return;
    if (typeof dialog.close === "function") {
      if (dialog.open) dialog.close();
      return;
    }
    dialog.removeAttribute("open");
  }

  function openDiagnosticoModal() {
    const dialog = $("diagnosticoCarteiraDialog");
    if (!dialog) return;

    const demandas = safeDemandas();
    renderSummary(buildSummary(demandas));
    ensureDefaultFields();
    const tipoRelatorio = $("diagTipoRelatorio");
    if (tipoRelatorio) {
      applyReportTypePreset(tipoRelatorio.value || "executivo", {
        forceTitle: false,
      });
    }
    openDialog(dialog);
  }

  function closeDiagnosticoModal() {
    closeDialog($("diagnosticoCarteiraDialog"));
  }

  function readCheckboxMap(fieldMap) {
    return Object.fromEntries(
      Object.entries(fieldMap).map(([key, id]) => [key, Boolean($(id)?.checked)]),
    );
  }

  function buildDiagnosticoConfig() {
    const demandas = safeDemandas();

    return {
      tipoRelatorio: $("diagTipoRelatorio")?.value || "executivo",
      titulo: trimValue($("diagTituloRelatorio")?.value) || DEFAULT_TITLE,
      responsavel: trimValue($("diagResponsavel")?.value) || currentUserLabel(),
      observacaoTecnica: trimValue($("diagObservacaoTecnica")?.value),
      escopoDiagnostico: trimValue($("diagEscopoDiagnostico")?.value),
      secoes: readCheckboxMap(sectionFields),
      opcoes: readCheckboxMap(optionFields),
      totalDemandasFiltradas: demandas.length,
      criadoEm: new Date().toISOString(),
    };
  }

  function notifyDiagnostico(message, type = "info") {
    const host = document.getElementById("toastHost");

    if (host) {
      const toast = document.createElement("div");
      toast.className = `toast ${type}`.trim();
      toast.setAttribute("role", "status");
      toast.setAttribute("aria-live", "polite");
      toast.textContent = message;

      host.appendChild(toast);

      requestAnimationFrame(() => {
        toast.classList.add("is-visible");
      });

      setTimeout(() => {
        toast.classList.remove("is-visible");
        setTimeout(() => toast.remove(), 300);
      }, 4200);

      return;
    }

    if (type === "error" && typeof global.alert === "function") {
      global.alert(message);
      return;
    }

    console.log(message);
  }

  function showDiagnosticoLoadingToast(message) {
    const host = document.getElementById("toastHost");
    if (!host) return null;

    const toast = document.createElement("div");
    toast.className = "toast loading is-visible";
    toast.setAttribute("role", "status");
    toast.setAttribute("aria-live", "polite");
    toast.innerHTML = `<span class="toast-icon toast-spinner">⟳</span><span>${message}</span>`;
    host.appendChild(toast);

    return {
      dismiss(successMessage) {
        if (successMessage) notifyDiagnostico(successMessage, "success");
        toast.classList.remove("is-visible");
        setTimeout(() => toast.remove(), 300);
      },
      error(errorMessage) {
        if (errorMessage) notifyDiagnostico(errorMessage, "error");
        toast.classList.remove("is-visible");
        setTimeout(() => toast.remove(), 300);
      },
    };
  }

  function setDiagnosticoGeneratingState(isGenerating) {
    const dialog = $("diagnosticoCarteiraDialog");
    const primaryButton = $("generateDiagnosticoCarteira");
    const cancelButton = $("cancelDiagnosticoCarteira");
    const closeButton = $("closeDiagnosticoCarteira");

    if (dialog) {
      dialog.setAttribute("aria-busy", isGenerating ? "true" : "false");
    }

    if (primaryButton) {
      if (!primaryButton.dataset.defaultLabel) {
        primaryButton.dataset.defaultLabel =
          trimValue(primaryButton.textContent) || "Gerar diagnóstico";
      }

      primaryButton.disabled = isGenerating;
      primaryButton.classList.toggle("is-loading", isGenerating);
      primaryButton.setAttribute("aria-busy", isGenerating ? "true" : "false");
      primaryButton.textContent = isGenerating
        ? "Gerando diagnóstico..."
        : primaryButton.dataset.defaultLabel;
    }

    [cancelButton, closeButton].forEach((button) => {
      if (!button) return;
      button.disabled = isGenerating;
      button.setAttribute("aria-disabled", isGenerating ? "true" : "false");
    });
  }

  function handleGenerateDiagnostico() {
    if (diagnosticoGenerating) return;

    try {
      const demandasFiltradas = safeDemandas();

      if (!demandasFiltradas.length) {
        notifyDiagnostico(
          "Não há demandas no recorte filtrado para gerar o diagnóstico.",
          "error",
        );
        return;
      }

      const config = buildDiagnosticoConfig();
      diagnosticoGenerating = true;
      setDiagnosticoGeneratingState(true);

      const loadingToast = showDiagnosticoLoadingToast(
        "Gerando diagnóstico da carteira. Aguarde alguns instantes...",
      );

      global.requestAnimationFrame(() => {
        global.setTimeout(() => {
          try {
            const diagnostico = global.CCEDiagnosticoEngine?.gerarDiagnostico?.(
              demandasFiltradas,
              {
                ...config,
                usuarioAtual: diagnosticoContext.getState?.()?.currentUser || null,
                filtrosAtuais: diagnosticoContext.getState?.()?.filters || {},
                helpers: {
                  primaryStatusOf: diagnosticoContext.primaryStatusOf,
                  substatusListOf: diagnosticoContext.substatusListOf,
                  dueClassOf: diagnosticoContext.dueClassOf,
                  dateText: diagnosticoContext.dateText,
                  toDate: diagnosticoContext.toDate,
                  formatDate: diagnosticoContext.formatDate,
                  formatDateTime: diagnosticoContext.formatDateTime,
                },
              },
            );

            if (!diagnostico) {
              throw new Error("A engine não retornou o diagnóstico técnico.");
            }

            console.log("Diagnóstico técnico gerado", diagnostico);

            if (!global.CCEDiagnosticoPDF?.gerarPDF) {
              throw new Error("Módulo de geração de PDF não carregado.");
            }

            global.CCEDiagnosticoPDF.gerarPDF(diagnostico, config);

            if (loadingToast) {
              loadingToast.dismiss("Diagnóstico da Carteira gerado com sucesso.");
            } else {
              notifyDiagnostico(
                "Diagnóstico da Carteira gerado com sucesso.",
                "success",
              );
            }
          } catch (error) {
            console.error("Erro ao gerar Diagnóstico da Carteira:", error);

            const message = `Erro ao gerar diagnóstico: ${
              error?.message || "erro inesperado"
            }`;

            if (loadingToast) {
              loadingToast.error(message);
            } else {
              notifyDiagnostico(message, "error");
            }
          } finally {
            diagnosticoGenerating = false;
            setDiagnosticoGeneratingState(false);
          }
        }, 40);
      });
    } catch (error) {
      console.error("Erro ao gerar Diagnóstico da Carteira:", error);

      notifyDiagnostico(
        `Erro ao gerar diagnóstico: ${error?.message || "erro inesperado"}`,
        "error",
      );
    }
  }

  function bindEvents() {
    if (handlersBound) return;
    handlersBound = true;

    $("openDiagnosticoCarteira")?.addEventListener(
      "click",
      openDiagnosticoModal,
    );
    $("closeDiagnosticoCarteira")?.addEventListener(
      "click",
      closeDiagnosticoModal,
    );
    $("cancelDiagnosticoCarteira")?.addEventListener(
      "click",
      closeDiagnosticoModal,
    );
    $("generateDiagnosticoCarteira")?.addEventListener(
      "click",
      handleGenerateDiagnostico,
    );
    $("diagTipoRelatorio")?.addEventListener("change", (event) => {
      applyReportTypePreset(event.currentTarget?.value || "executivo", {
        forceTitle: true,
      });
    });
  }

  function initDiagnosticoCarteira(options = {}) {
    diagnosticoContext = { ...options };
    bindEvents();

    global.CCEDiagnosticoEngine?.init?.(diagnosticoContext);
    global.CCEDiagnosticoCharts?.init?.({
      config: global.CCEDiagnosticoConfig || null,
    });
    global.CCEDiagnosticoPDF?.init?.({
      config: global.CCEDiagnosticoConfig || null,
    });

    console.log("Diagnostico da Carteira inicializado", {
      hasDocument: Boolean(document),
      reportName: global.CCEDiagnosticoConfig?.reportName || null,
    });
  }

  function getContext() {
    return diagnosticoContext;
  }

  global.CCEDiagnosticoCarteira = {
    init: initDiagnosticoCarteira,
    getContext,
    open: openDiagnosticoModal,
    close: closeDiagnosticoModal,
  };
})(window, document);

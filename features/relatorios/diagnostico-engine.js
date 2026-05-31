(function setupDiagnosticoEngine(global) {
  "use strict";

  const OPEN_STATUS = new Set(["A Planejar", "Planejado", "Replanejado"]);
  const CLOSED_STATUS = new Set(["Realizado", "Cancelado"]);
  const TOP_CRITICAL_LIMIT = 100;
  const DEMANDAS_SEM_KM_LIMIT = 50;
  const CLIMA_CRITICO_LIMIT = 50;
  const CLIMA_PDF_LIMIT = 30;

  const HIGH_SENSITIVITY_RULES = [
    {
      categoria: "Vegetacao",
      termos: ["roco", "roço", "poda", "vegetacao", "vegetação", "capina"],
      justificativa:
        "Atividade exposta a chuva e umidade, com impacto direto na mobilizacao e seguranca de campo.",
    },
    {
      categoria: "Drenagem",
      termos: ["bueiro", "drenagem", "canaleta", "valeta", "galeria"],
      justificativa:
        "Atividade sensivel a acumulacao de agua e alteracao rapida das condicoes de solo.",
    },
    {
      categoria: "Talude/Aterro",
      termos: ["talude", "aterro", "corte", "erosao", "erosão"],
      justificativa:
        "Atividade em zona de instabilidade geotecnica com sensibilidade elevada a chuva.",
    },
    {
      categoria: "Energia Externa",
      termos: [
        "spda",
        "aterramento",
        "iluminacao externa",
        "iluminação externa",
        "rede aerea",
        "rede aérea",
        "subestacao externa",
        "subestação externa",
      ],
      justificativa:
        "Atividade externa com maior exposicao a chuva, descarga atmosferica e condicao insegura de trabalho.",
    },
    {
      categoria: "Via Permanente",
      termos: [
        "passagem em nivel",
        "passagem em nível",
        " pn ",
        "soldagem de trilho",
        "tls",
        "esmerilhamento",
      ],
      justificativa:
        "Intervencao em via permanente com restricoes de execucao sob chuva e pista molhada.",
    },
  ];

  const MEDIUM_SENSITIVITY_RULES = [
    {
      categoria: "Sinalizacao",
      termos: [
        "amv",
        "jac",
        "mch",
        "maquina de chave",
        "máquina de chave",
        "junta isolante",
        "jic",
        "circuito de via",
      ],
      justificativa:
        "Atividade com interferencia de campo e sensibilidade moderada a umidade e acesso.",
    },
    {
      categoria: "Sinalizacao",
      termos: ["telecom", "fibra", "cctv"],
      justificativa:
        "Atividade externa com dependencia de acesso, seguranca e visibilidade.",
    },
    {
      categoria: "Geral",
      termos: ["inspecao", "inspeção", "preventiva externa"],
      justificativa:
        "Atividade de campo com sensibilidade moderada a chuva e restricao operacional.",
    },
  ];

  let engineContext = {};

  function safeText(value) {
    if (value === null || value === undefined) return "";
    return String(value).trim();
  }

  function normalizeText(value) {
    return safeText(value)
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase();
  }

  function toNumber(value) {
    if (typeof value === "number") return Number.isFinite(value) ? value : 0;
    const normalized = safeText(value).replace(",", ".");
    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : 0;
  }

  function dateText(value, helpers = {}) {
    if (typeof helpers.dateText === "function") {
      return safeText(helpers.dateText(value));
    }

    const date = toDate(value, helpers);
    if (!date) return "";
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(
      date.getDate(),
    ).padStart(2, "0")}`;
  }

  function toDate(value, helpers = {}) {
    if (typeof helpers.toDate === "function") {
      const helperDate = helpers.toDate(value);
      if (helperDate instanceof Date && !Number.isNaN(helperDate.getTime())) {
        return helperDate;
      }
    }

    if (!value) return null;
    if (value instanceof Date) {
      return Number.isNaN(value.getTime()) ? null : value;
    }

    const raw = safeText(value);
    if (!raw) return null;

    const brMatch = raw.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (brMatch) {
      const [, dd, mm, yyyy] = brMatch;
      return new Date(Number(yyyy), Number(mm) - 1, Number(dd));
    }

    const parsed = new Date(raw);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }

  function formatPercent(value, total) {
    if (!total) return 0;
    return Number(((value / total) * 100).toFixed(1));
  }

  function groupBy(rows, selector) {
    return rows.reduce((acc, row) => {
      const key = safeText(selector(row)) || "Nao informado";
      if (!acc[key]) acc[key] = [];
      acc[key].push(row);
      return acc;
    }, {});
  }

  function countBy(rows, selector) {
    return rows.reduce((acc, row) => {
      const key = safeText(selector(row)) || "Nao informado";
      acc[key] = (acc[key] || 0) + 1;
      return acc;
    }, {});
  }

  function sortEntriesDesc(mapOrObject) {
    const entries = mapOrObject instanceof Map
      ? Array.from(mapOrObject.entries())
      : Object.entries(mapOrObject || {});

    return entries.sort((a, b) => {
      if (b[1] !== a[1]) return b[1] - a[1];
      return safeText(a[0]).localeCompare(safeText(b[0]), "pt-BR");
    });
  }

  function uniqueValues(rows, selector) {
    return Array.from(
      new Set(
        rows
          .map((row) => safeText(selector(row)))
          .filter(Boolean),
      ),
    );
  }

  function isFilled(value) {
    return safeText(value) !== "";
  }

  function startOfDay(date) {
    const normalized = new Date(date.getTime());
    normalized.setHours(0, 0, 0, 0);
    return normalized;
  }

  function addDays(date, amount) {
    const next = new Date(date.getTime());
    next.setDate(next.getDate() + amount);
    return next;
  }

  function isOpenStatus(status) {
    return OPEN_STATUS.has(status);
  }

  function getStatus(demanda, helpers = {}) {
    if (typeof helpers.primaryStatusOf === "function") {
      return safeText(helpers.primaryStatusOf(demanda));
    }

    const explicit = safeText(demanda?.statusOperacional);
    if (explicit) return explicit;
    if (isFilled(demanda?.dataRealizada)) return "Realizado";
    if (isFilled(demanda?.dataReplanejadaAtual)) return "Replanejado";
    if (isFilled(demanda?.dataPlanejada)) return "Planejado";
    return "A Planejar";
  }

  function getSubstatusList(demanda, helpers = {}) {
    if (typeof helpers.substatusListOf === "function") {
      const result = helpers.substatusListOf(demanda);
      if (Array.isArray(result)) {
        return result.map((item) => safeText(item)).filter(Boolean);
      }
    }

    const explicit = safeText(demanda?.substatusOperacional);
    return explicit
      ? explicit.split("|").map((item) => safeText(item)).filter(Boolean)
      : [];
  }

  function isDemandVencida(demanda, helpers = {}) {
    const status = getStatus(demanda, helpers);
    if (CLOSED_STATUS.has(status)) return false;

    const dueDate = toDate(demanda?.vencimento, helpers);
    if (!dueDate) return false;
    return startOfDay(dueDate) < startOfDay(new Date());
  }

  function hasKm(demanda) {
    return Boolean(
      safeText(demanda?.kmInicio) ||
        safeText(demanda?.kmFim) ||
        safeText(demanda?.KmInicio) ||
        safeText(demanda?.KmFim),
    );
  }

  function isCriticalDemand(demanda) {
    return normalizeText(demanda?.critico) === "sim";
  }

  function isCentroCadastrado(demanda) {
    return (
      normalizeText(demanda?.centroTrabalhoStatus) === "cadastrado" ||
      demanda?.centroTrabalhoCadastrado === true
    );
  }

  function effectivePlanningDate(demanda, helpers = {}) {
    return (
      toDate(demanda?.dataReplanejadaAtual, helpers) ||
      toDate(demanda?.dataPlanejada, helpers)
    );
  }

  function weekKey(date) {
    if (!date) return "";
    const current = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    const dayNum = current.getUTCDay() || 7;
    current.setUTCDate(current.getUTCDate() + 4 - dayNum);
    const yearStart = new Date(Date.UTC(current.getUTCFullYear(), 0, 1));
    const weekNum = Math.ceil((((current - yearStart) / 86400000) + 1) / 7);
    return `${current.getUTCFullYear()}-W${String(weekNum).padStart(2, "0")}`;
  }

  function buildFiltersResumo(contexto = {}) {
    return contexto.filtrosAtuais && typeof contexto.filtrosAtuais === "object"
      ? { ...contexto.filtrosAtuais }
      : {};
  }

  function getClimaRuntime() {
    return {
      climate: global.CCEClimate || null,
      climaFeature: global.CCEClimaFeature || null,
      config: global.CCEClimateConfig || null,
    };
  }

  function climaAvailabilityText(runtime) {
    if (runtime.climate) {
      return "Sensibilidade climatica por atividade com configuracao climatica carregada.";
    }
    if (runtime.config) {
      return "Sensibilidade climatica por atividade com apoio de configuracao regional do modulo clima.";
    }
    return "Sensibilidade climatica por atividade sem dados climaticos carregados.";
  }

  function buildMeta(demandas, contexto = {}) {
    return {
      titulo: safeText(contexto.titulo) || "Diagnostico da Carteira",
      tipoRelatorio: safeText(contexto.tipoRelatorio) || "executivo",
      responsavel: safeText(contexto.responsavel),
      observacaoTecnica: safeText(contexto.observacaoTecnica),
      criadoEm: contexto.criadoEm || new Date().toISOString(),
      totalDemandas: demandas.length,
      usuarioAtual: contexto.usuarioAtual || null,
      filtrosResumo: buildFiltersResumo(contexto),
    };
  }

  function buildResumo(demandas, helpers = {}) {
    const today = startOfDay(new Date());
    const limit7 = addDays(today, 7);
    const limit20 = addDays(today, 20);

    const resumo = {
      total: demandas.length,
      aPlanejar: 0,
      planejadas: 0,
      replanejadas: 0,
      realizadas: 0,
      canceladas: 0,
      vencidas: 0,
      proximasVencimento7d: 0,
      proximasVencimento20d: 0,
      criticas: 0,
      comKm: 0,
      semKm: 0,
      comCentroCadastrado: 0,
      semCentroCadastrado: 0,
      percentualPlanejado: 0,
      percentualRealizado: 0,
      percentualComKm: 0,
      percentualCritico: 0,
      percentualVencido: 0,
    };

    demandas.forEach((demanda) => {
      const status = getStatus(demanda, helpers);
      const dueDate = toDate(demanda?.vencimento, helpers);
      const hasDue = dueDate ? startOfDay(dueDate) : null;

      if (status === "A Planejar") resumo.aPlanejar += 1;
      if (status === "Planejado") resumo.planejadas += 1;
      if (status === "Replanejado") resumo.replanejadas += 1;
      if (status === "Realizado") resumo.realizadas += 1;
      if (status === "Cancelado") resumo.canceladas += 1;
      if (isDemandVencida(demanda, helpers)) resumo.vencidas += 1;
      if (isCriticalDemand(demanda)) resumo.criticas += 1;
      if (hasKm(demanda)) resumo.comKm += 1;
      else resumo.semKm += 1;
      if (isCentroCadastrado(demanda)) resumo.comCentroCadastrado += 1;
      else resumo.semCentroCadastrado += 1;

      if (hasDue && isOpenStatus(status) && hasDue >= today && hasDue <= limit7) {
        resumo.proximasVencimento7d += 1;
      }

      if (hasDue && isOpenStatus(status) && hasDue >= today && hasDue <= limit20) {
        resumo.proximasVencimento20d += 1;
      }
    });

    resumo.percentualPlanejado = formatPercent(
      resumo.planejadas + resumo.replanejadas,
      resumo.total,
    );
    resumo.percentualRealizado = formatPercent(resumo.realizadas, resumo.total);
    resumo.percentualComKm = formatPercent(resumo.comKm, resumo.total);
    resumo.percentualCritico = formatPercent(resumo.criticas, resumo.total);
    resumo.percentualVencido = formatPercent(resumo.vencidas, resumo.total);

    return resumo;
  }

  function rankedDistribution(rows, labelField, helpers = {}) {
    const total = rows.length || 1;
    return sortEntriesDesc(groupBy(rows, labelField)).map(([label, group]) => ({
      nome: label,
      quantidade: group.length,
      percentual: formatPercent(group.length, total),
      planejadas: group.filter((item) => getStatus(item, helpers) === "Planejado").length,
      vencidas: group.filter((item) => isDemandVencida(item, helpers)).length,
      criticas: group.filter(isCriticalDemand).length,
    }));
  }

  function buildStatusAnalysis(demandas, helpers = {}) {
    const porStatusMap = countBy(demandas, (item) => getStatus(item, helpers));
    const porSubstatusMap = {};

    demandas.forEach((item) => {
      getSubstatusList(item, helpers).forEach((substatus) => {
        porSubstatusMap[substatus] = (porSubstatusMap[substatus] || 0) + 1;
      });
    });

    const porStatus = sortEntriesDesc(porStatusMap).map(([status, quantidade]) => ({
      status,
      quantidade,
      percentual: formatPercent(quantidade, demandas.length),
    }));

    const porSubstatus = sortEntriesDesc(porSubstatusMap).map(([substatus, quantidade]) => ({
      substatus,
      quantidade,
      percentual: formatPercent(quantidade, demandas.length),
    }));

    const leitura = [];
    if (porStatus[0]) {
      leitura.push(`A carteira apresenta maior concentracao em ${porStatus[0].status}.`);
    }
    if (demandas.some((item) => isDemandVencida(item, helpers))) {
      leitura.push("Ha volume relevante de demandas vencidas no recorte.");
    }
    if ((porStatusMap.Replanejado || 0) > 0) {
      leitura.push(
        "A presenca de replanejamentos indica necessidade de acompanhamento da estabilidade da programacao.",
      );
    }

    return {
      porStatus,
      porSubstatus,
      leitura,
    };
  }

  function compactDemand(demanda, helpers = {}) {
    return {
      id: safeText(demanda?.id),
      ordem: safeText(demanda?.ordem),
      descricao: safeText(demanda?.descricao),
      centroTrabalho: safeText(demanda?.centroTrabalho),
      localInstalacao: safeText(demanda?.localInstalacao),
      kmInicio: safeText(demanda?.kmInicio),
      kmFim: safeText(demanda?.kmFim),
      vencimento: dateText(demanda?.vencimento, helpers),
      dataPlanejada: dateText(demanda?.dataPlanejada, helpers),
      status: getStatus(demanda, helpers),
      critico: safeText(demanda?.critico),
    };
  }

  function buildVencimentosAnalysis(demandas, helpers = {}) {
    const today = startOfDay(new Date());
    const limit7 = addDays(today, 7);
    const limit20 = addDays(today, 20);
    const perDate = {};
    const perWeek = {};

    const vencidas = [];
    const proximas7d = [];
    const proximas20d = [];

    demandas.forEach((demanda) => {
      const dueDate = toDate(demanda?.vencimento, helpers);
      const status = getStatus(demanda, helpers);
      if (!dueDate) return;

      const dayKey = dateText(dueDate, helpers);
      perDate[dayKey] = (perDate[dayKey] || 0) + 1;
      perWeek[weekKey(dueDate)] = (perWeek[weekKey(dueDate)] || 0) + 1;

      const normalizedDue = startOfDay(dueDate);
      const base = compactDemand(demanda, helpers);

      if (isDemandVencida(demanda, helpers)) vencidas.push(base);
      if (isOpenStatus(status) && normalizedDue >= today && normalizedDue <= limit7) {
        proximas7d.push(base);
      }
      if (isOpenStatus(status) && normalizedDue >= today && normalizedDue <= limit20) {
        proximas20d.push(base);
      }
    });

    vencidas.sort((a, b) => safeText(a.vencimento).localeCompare(safeText(b.vencimento)));
    proximas7d.sort((a, b) => safeText(a.vencimento).localeCompare(safeText(b.vencimento)));
    proximas20d.sort((a, b) => safeText(a.vencimento).localeCompare(safeText(b.vencimento)));

    const leitura = [];
    if (vencidas.length) {
      leitura.push(`Foram identificadas ${vencidas.length} demandas vencidas com necessidade de tratamento prioritario.`);
    }
    if (proximas7d.length) {
      leitura.push(`Ha ${proximas7d.length} demandas em janela de vencimento nos proximos 7 dias.`);
    }

    return {
      vencidas,
      proximas7d,
      proximas20d,
      porData: sortEntriesDesc(perDate).map(([data, quantidade]) => ({ data, quantidade })),
      porSemana: sortEntriesDesc(perWeek).map(([semana, quantidade]) => ({ semana, quantidade })),
      leitura,
    };
  }

  function buildPlanejamentoAnalysis(demandas, helpers = {}) {
    const planejadas = [];
    const porDataPlanejada = {};
    const porSemanaPlanejada = {};
    let semDataPlanejada = 0;

    demandas.forEach((demanda) => {
      const status = getStatus(demanda, helpers);
      const effectiveDate = effectivePlanningDate(demanda, helpers);

      if (status === "Planejado" || status === "Replanejado") {
        planejadas.push({
          id: safeText(demanda?.id),
          ordem: safeText(demanda?.ordem),
          descricao: safeText(demanda?.descricao),
          dataPlanejada: dateText(effectiveDate, helpers),
          status,
        });
      }

      if (!effectiveDate) {
        semDataPlanejada += 1;
        return;
      }

      const key = dateText(effectiveDate, helpers);
      porDataPlanejada[key] = (porDataPlanejada[key] || 0) + 1;
      porSemanaPlanejada[weekKey(effectiveDate)] =
        (porSemanaPlanejada[weekKey(effectiveDate)] || 0) + 1;
    });

    const diasComMaiorCarga = sortEntriesDesc(porDataPlanejada)
      .slice(0, 10)
      .map(([data, quantidade]) => ({ data, quantidade }));

    const leitura = [];
    if (diasComMaiorCarga[0]) {
      leitura.push(
        `A maior concentracao de carga planejada ocorre em ${diasComMaiorCarga[0].data}, com ${diasComMaiorCarga[0].quantidade} demandas.`,
      );
    }
    if (semDataPlanejada > 0) {
      leitura.push(
        `Ha ${semDataPlanejada} demandas sem data efetiva de planejamento definida.`,
      );
    }

    return {
      planejadas,
      porDataPlanejada: sortEntriesDesc(porDataPlanejada).map(([data, quantidade]) => ({
        data,
        quantidade,
      })),
      porSemanaPlanejada: sortEntriesDesc(porSemanaPlanejada).map(
        ([semana, quantidade]) => ({ semana, quantidade }),
      ),
      semDataPlanejada,
      diasComMaiorCarga,
      leitura,
    };
  }

  function buildCentrosAnalysis(demandas, helpers = {}) {
    const groups = groupBy(demandas, (item) => item.centroTrabalho || "Sem centro");
    const total = demandas.length || 1;
    const ranking = sortEntriesDesc(groups).map(([centroTrabalho, rows]) => ({
      centroTrabalho,
      quantidade: rows.length,
      percentual: formatPercent(rows.length, total),
      planejadas: rows.filter((item) => getStatus(item, helpers) === "Planejado").length,
      aPlanejar: rows.filter((item) => getStatus(item, helpers) === "A Planejar").length,
      replanejadas: rows.filter((item) => getStatus(item, helpers) === "Replanejado").length,
      realizadas: rows.filter((item) => getStatus(item, helpers) === "Realizado").length,
      vencidas: rows.filter((item) => isDemandVencida(item, helpers)).length,
      criticas: rows.filter(isCriticalDemand).length,
      comKm: rows.filter(hasKm).length,
      semKm: rows.filter((item) => !hasKm(item)).length,
    }));

    const centrosSemCadastro = ranking
      .filter(({ centroTrabalho }) =>
        groups[centroTrabalho].some((item) => !isCentroCadastrado(item)),
      )
      .map((item) => item.centroTrabalho);

    const leitura = [];
    if (ranking[0]) {
      leitura.push(
        `O centro ${ranking[0].centroTrabalho} concentra o maior volume do recorte filtrado.`,
      );
    }
    if (centrosSemCadastro.length) {
      leitura.push(
        `Foram identificados ${centrosSemCadastro.length} centros com necessidade de saneamento cadastral.`,
      );
    }

    return {
      totalCentros: ranking.length,
      ranking,
      centrosSemCadastro,
      leitura,
    };
  }

  function buildPlanejadoresAnalysis(demandas, helpers = {}) {
    const groups = groupBy(
      demandas,
      (item) => item.planejadorCurto || "Sem planejador",
    );
    const total = demandas.length || 1;
    const ranking = sortEntriesDesc(groups).map(([planejadorCurto, rows]) => ({
      planejadorCurto,
      quantidade: rows.length,
      percentual: formatPercent(rows.length, total),
      planejadas: rows.filter((item) => getStatus(item, helpers) === "Planejado").length,
      aPlanejar: rows.filter((item) => getStatus(item, helpers) === "A Planejar").length,
      replanejadas: rows.filter((item) => getStatus(item, helpers) === "Replanejado").length,
      realizadas: rows.filter((item) => getStatus(item, helpers) === "Realizado").length,
      vencidas: rows.filter((item) => isDemandVencida(item, helpers)).length,
      criticas: rows.filter(isCriticalDemand).length,
    }));

    const semPlanejador = groups["Sem planejador"]?.length || 0;
    const leitura = [];
    if (ranking[0]) {
      leitura.push(
        `O maior volume de carteira esta atribuido ao planejador ${ranking[0].planejadorCurto}.`,
      );
    }
    if (semPlanejador) {
      leitura.push(`Ha ${semPlanejador} demandas sem planejador de curto definido.`);
    }

    return {
      totalPlanejadores: ranking.length,
      ranking,
      semPlanejador,
      leitura,
    };
  }

  function buildSupervisoesAnalysis(demandas, helpers = {}) {
    const ranking = rankedDistribution(
      demandas,
      (item) => item.supervisao || "Sem supervisao",
      helpers,
    );
    return {
      total: ranking.length,
      ranking,
      leitura: ranking[0]
        ? [`A supervisao ${ranking[0].nome} lidera o volume do recorte analisado.`]
        : [],
    };
  }

  function buildGerenciasAnalysis(demandas, helpers = {}) {
    const ranking = rankedDistribution(
      demandas,
      (item) => item.gerencia || "Sem gerencia",
      helpers,
    );
    return {
      total: ranking.length,
      ranking,
      leitura: ranking[0]
        ? [`A gerencia ${ranking[0].nome} concentra o maior volume da carteira filtrada.`]
        : [],
    };
  }

  function buildKmAnalysis(demandas, helpers = {}) {
    const comKmRows = demandas.filter(hasKm);
    const semKmRows = demandas.filter((item) => !hasKm(item));
    const trechos = {};

    comKmRows.forEach((item) => {
      const trecho = safeText(item.kmInicio) && safeText(item.kmFim)
        ? `${safeText(item.kmInicio)} - ${safeText(item.kmFim)}`
        : safeText(item.kmInicio) || safeText(item.kmFim) || "Sem trecho";

      if (!trechos[trecho]) {
        trechos[trecho] = {
          trecho,
          kmInicio: safeText(item.kmInicio),
          kmFim: safeText(item.kmFim),
          quantidade: 0,
          centros: new Set(),
          locais: new Set(),
        };
      }

      trechos[trecho].quantidade += 1;
      if (isFilled(item.centroTrabalho)) trechos[trecho].centros.add(item.centroTrabalho);
      if (isFilled(item.localInstalacao)) trechos[trecho].locais.add(item.localInstalacao);
    });

    const rankingTrechos = Object.values(trechos)
      .sort((a, b) => b.quantidade - a.quantidade)
      .map((item) => ({
        trecho: item.trecho,
        kmInicio: item.kmInicio,
        kmFim: item.kmFim,
        quantidade: item.quantidade,
        centros: Array.from(item.centros),
        locais: Array.from(item.locais),
      }));

    const percentualComKm = formatPercent(comKmRows.length, demandas.length);
    const leitura = [];
    if (percentualComKm >= 80) {
      leitura.push(
        `O recorte possui ${percentualComKm}% das demandas com KM informado, permitindo boa rastreabilidade territorial.`,
      );
    } else {
      leitura.push(
        `O recorte possui ${percentualComKm}% das demandas com KM informado, indicando necessidade de melhoria na rastreabilidade territorial.`,
      );
    }
    if (semKmRows.length) {
      leitura.push(
        "Ha demandas sem KM que devem ser saneadas para melhorar analise de interferencia e planejamento de campo.",
      );
    }

    return {
      comKm: comKmRows.length,
      semKm: semKmRows.length,
      percentualComKm,
      rankingTrechos,
      demandasSemKm: semKmRows
        .slice(0, DEMANDAS_SEM_KM_LIMIT)
        .map((item) => compactDemand(item, helpers)),
      leitura,
    };
  }

  function buildToleranciasAnalysis(demandas, helpers = {}) {
    let comTolerancia = 0;
    let semTolerancia = 0;
    let foraJanela = 0;
    let dentroJanela = 0;
    const vencimentoForaTolerancia = [];

    demandas.forEach((item) => {
      const minDate = toDate(item.toleranciaMin, helpers);
      const maxDate = toDate(item.toleranciaMax, helpers);
      const effectiveDate = effectivePlanningDate(item, helpers);

      if (!minDate && !maxDate) {
        semTolerancia += 1;
        return;
      }

      comTolerancia += 1;
      if (!effectiveDate) return;

      const beforeMin = minDate && effectiveDate < minDate;
      const afterMax = maxDate && effectiveDate > maxDate;
      if (beforeMin || afterMax) {
        foraJanela += 1;
        vencimentoForaTolerancia.push({
          ...compactDemand(item, helpers),
          toleranciaMin: dateText(item.toleranciaMin, helpers),
          toleranciaMax: dateText(item.toleranciaMax, helpers),
          dataPlanejadaEfetiva: dateText(effectiveDate, helpers),
        });
      } else {
        dentroJanela += 1;
      }
    });

    const leitura = [];
    if (foraJanela > 0) {
      leitura.push(
        `Foram identificadas ${foraJanela} demandas planejadas fora da janela de tolerancia.`,
      );
    }
    if (semTolerancia > 0) {
      leitura.push(`Ha ${semTolerancia} demandas sem tolerancias definidas na base.`);
    }

    return {
      comTolerancia,
      semTolerancia,
      foraJanela,
      dentroJanela,
      vencimentoForaTolerancia,
      leitura,
    };
  }

  function riscoReasonsFor(demanda, helpers = {}) {
    const reasons = [];
    const status = getStatus(demanda, helpers);
    const dueDate = toDate(demanda?.vencimento, helpers);
    const plannedDate = effectivePlanningDate(demanda, helpers);
    const today = startOfDay(new Date());
    const soonLimit = addDays(today, 7);

    if (isDemandVencida(demanda, helpers)) reasons.push("Vencida");
    if (isCriticalDemand(demanda) && isOpenStatus(status)) reasons.push("Critica aberta");
    if (dueDate && plannedDate && plannedDate > dueDate) {
      reasons.push("Planejada apos vencimento");
    }
    if (!hasKm(demanda)) reasons.push("Sem KM");
    if (!isCentroCadastrado(demanda)) reasons.push("Sem centro cadastrado");
    if (!isFilled(demanda?.planejadorCurto)) reasons.push("Sem planejador");
    const normalizedDue = dueDate ? startOfDay(dueDate) : null;
    if (
      normalizedDue &&
      normalizedDue >= today &&
      normalizedDue <= soonLimit &&
      isOpenStatus(status)
    ) {
      reasons.push("Proxima do vencimento");
    }

    return reasons;
  }

  function buildRiscosAnalysis(demandas, helpers = {}) {
    const summary = {
      vencidas: demandas.filter((item) => isDemandVencida(item, helpers)).length,
      criticasAbertas: demandas.filter(
        (item) => isCriticalDemand(item) && isOpenStatus(getStatus(item, helpers)),
      ).length,
      semKm: demandas.filter((item) => !hasKm(item)).length,
      semCentro: demandas.filter((item) => !isCentroCadastrado(item)).length,
      semPlanejador: demandas.filter((item) => !isFilled(item.planejadorCurto)).length,
      proximasVencimento: demandas.filter((item) =>
        riscoReasonsFor(item, helpers).includes("Proxima do vencimento"),
      ).length,
      planejamentoAtrasado: demandas.filter((item) =>
        riscoReasonsFor(item, helpers).includes("Planejada apos vencimento"),
      ).length,
      reprogramacoes: demandas.filter((item) => isFilled(item.dataReplanejadaAtual)).length,
    };

    const factors = [
      {
        fator: "Demandas vencidas",
        quantidade: summary.vencidas,
        severidade: "Alta",
        descricao: "Demandas abertas com vencimento expirado.",
      },
      {
        fator: "Demandas criticas abertas",
        quantidade: summary.criticasAbertas,
        severidade: "Alta",
        descricao: "Demandas criticas ainda sem encerramento operacional.",
      },
      {
        fator: "Demandas planejadas apos vencimento",
        quantidade: summary.planejamentoAtrasado,
        severidade: "Alta",
        descricao: "Programacao posicionada depois da data de vencimento.",
      },
      {
        fator: "Demandas sem KM",
        quantidade: summary.semKm,
        severidade: "Moderada",
        descricao: "Falta de rastreabilidade territorial da carteira.",
      },
      {
        fator: "Demandas sem centro cadastrado",
        quantidade: summary.semCentro,
        severidade: "Alta",
        descricao: "Centro de trabalho sem cadastro consistente para gestao.",
      },
      {
        fator: "Demandas sem planejador de curto",
        quantidade: summary.semPlanejador,
        severidade: "Moderada",
        descricao: "Atividades sem atribuicao clara de planejamento.",
      },
      {
        fator: "Demandas reprogramadas",
        quantidade: summary.reprogramacoes,
        severidade: "Moderada",
        descricao: "Sinal de instabilidade na programacao operacional.",
      },
      {
        fator: "Demandas proximas do vencimento",
        quantidade: summary.proximasVencimento,
        severidade: "Moderada",
        descricao: "Carteira exigindo resposta de curtissimo prazo.",
      },
    ].filter((item) => item.quantidade > 0);

    let rawScore = 0;
    rawScore += Math.min(summary.vencidas * 3, 36);
    rawScore += Math.min(summary.criticasAbertas * 2, 24);
    rawScore += Math.min(summary.semKm, 12);
    rawScore += Math.min(summary.semCentro * 2, 16);
    rawScore += Math.min(summary.semPlanejador, 8);
    rawScore += Math.min(summary.proximasVencimento, 8);
    rawScore += Math.min(summary.planejamentoAtrasado * 2, 14);
    rawScore += Math.min(summary.reprogramacoes, 10);

    const scoreRiscoGeral = Math.min(100, rawScore);
    const nivelRisco =
      scoreRiscoGeral <= 25
        ? "Baixo"
        : scoreRiscoGeral <= 50
          ? "Moderado"
          : scoreRiscoGeral <= 75
            ? "Alto"
            : "Critico";

    const demandasRiscoAlto = demandas
      .map((item) => ({
        ...compactDemand(item, helpers),
        motivos: riscoReasonsFor(item, helpers),
      }))
      .filter((item) => item.motivos.length >= 2)
      .sort((a, b) => b.motivos.length - a.motivos.length)
      .slice(0, TOP_CRITICAL_LIMIT);

    const leitura = [];
    leitura.push(`O recorte apresenta nivel de risco ${nivelRisco.toLowerCase()} com score ${scoreRiscoGeral}.`);
    if (summary.vencidas > 0) {
      leitura.push("Demandas vencidas permanecem como principal fator de exposicao operacional.");
    }
    if (summary.semKm > 0 || summary.semCentro > 0) {
      leitura.push("Ha fragilidades de cadastro que reduzem a qualidade do planejamento e da rastreabilidade.");
    }

    return {
      scoreRiscoGeral,
      nivelRisco,
      fatores: factors,
      demandasRiscoAlto,
      leitura,
    };
  }

  function getDataOperacional(demanda, helpers = {}) {
    const dataReplanejadaAtual = dateText(demanda?.dataReplanejadaAtual, helpers);
    if (dataReplanejadaAtual) {
      return { data: dataReplanejadaAtual, origem: "Replanejada" };
    }

    const dataPlanejada = dateText(demanda?.dataPlanejada, helpers);
    if (dataPlanejada) {
      return { data: dataPlanejada, origem: "Planejada" };
    }

    const vencimento = dateText(demanda?.vencimento, helpers);
    if (vencimento) {
      return { data: vencimento, origem: "Vencimento" };
    }

    return { data: "", origem: "Sem data" };
  }

  function sensitivitySourceText(demanda) {
    return [
      demanda?.descricao,
      demanda?.tipoOM,
      demanda?.centroTrabalho,
      demanda?.localInstalacao,
      demanda?.observacao,
      demanda?.comentario,
    ].join(" ");
  }

  function classificarSensibilidadeClimatica(demanda) {
    const haystack = normalizeText(sensitivitySourceText(demanda));

    for (const rule of HIGH_SENSITIVITY_RULES) {
      if (rule.termos.some((term) => haystack.includes(normalizeText(term)))) {
        return {
          nivel: "Alto",
          categoria: rule.categoria,
          justificativa: rule.justificativa,
        };
      }
    }

    for (const rule of MEDIUM_SENSITIVITY_RULES) {
      if (rule.termos.some((term) => haystack.includes(normalizeText(term)))) {
        return {
          nivel: "Medio",
          categoria: rule.categoria,
          justificativa: rule.justificativa,
        };
      }
    }

    return {
      nivel: "Baixo",
      categoria: "Geral",
      justificativa:
        "Nao foram identificados elementos de alta exposicao climatica na descricao da atividade.",
    };
  }

  function locateClimateDistrict(demanda, runtime) {
    const config = runtime?.config;
    if (!config?.distritos?.length) return null;

    const gerencia = normalizeText(demanda?.gerencia);
    const centro = normalizeText(demanda?.centroTrabalho);
    const kmInicio = toNumber(safeText(demanda?.kmInicio).replace("KM", ""));
    const kmFim = toNumber(safeText(demanda?.kmFim).replace("KM", ""));

    return (
      config.distritos.find((distrito) => {
        const ga = normalizeText(distrito.ga);
        if (gerencia && ga && ga === gerencia) return true;

        const centros = Array.isArray(distrito.centrosTrabalho)
          ? distrito.centrosTrabalho.map(normalizeText)
          : [];
        if (centro && centros.includes(centro)) return true;

        const start = Number(distrito.kmInicial);
        const end = Number(distrito.kmFinal);
        if (Number.isFinite(kmInicio) && Number.isFinite(start) && Number.isFinite(end)) {
          return kmInicio >= Math.min(start, end) && kmInicio <= Math.max(start, end);
        }
        if (Number.isFinite(kmFim) && Number.isFinite(start) && Number.isFinite(end)) {
          return kmFim >= Math.min(start, end) && kmFim <= Math.max(start, end);
        }
        return false;
      }) || null
    );
  }

  function getClimateRealSignal(_demanda, runtime) {
    if (!runtime?.climate) return null;
    return null;
  }

  function calcularRiscoClimaticoDemanda(demanda, contexto = {}, helpers = {}) {
    const runtime = contexto.runtime || getClimaRuntime();
    const sensibilidade = classificarSensibilidadeClimatica(demanda);
    const dataOperacional = getDataOperacional(demanda, helpers);
    const dueDate = toDate(demanda?.vencimento, helpers);
    const today = startOfDay(new Date());
    const soonLimit = addDays(today, 7);
    const fatores = [];
    let score = 0;

    if (sensibilidade.nivel === "Alto") {
      score += 40;
      fatores.push("Atividade com alta sensibilidade a chuva");
    } else if (sensibilidade.nivel === "Medio") {
      score += 25;
      fatores.push("Atividade com sensibilidade moderada a condicoes climaticas");
    } else {
      score += 10;
      fatores.push("Atividade com baixa sensibilidade climatica");
    }

    if (isCriticalDemand(demanda)) {
      score += 15;
      fatores.push("Demanda critica");
    }

    if (isDemandVencida(demanda, helpers)) {
      score += 15;
      fatores.push("Demanda vencida");
    } else if (
      dueDate &&
      isOpenStatus(getStatus(demanda, helpers)) &&
      startOfDay(dueDate) >= today &&
      startOfDay(dueDate) <= soonLimit
    ) {
      score += 15;
      fatores.push("Demanda proxima do vencimento");
    }

    if (
      !hasKm(demanda) ||
      !isFilled(demanda?.localInstalacao) ||
      !isFilled(demanda?.centroTrabalho)
    ) {
      score += 10;
      fatores.push("Localizacao operacional incompleta");
    }

    if (!dataOperacional.data) {
      score += 10;
      fatores.push("Sem data operacional definida");
    }

    const realSignal = getClimateRealSignal(demanda, runtime);
    if (realSignal?.nivel === "Alto") {
      score += 20;
      fatores.push("Sinal climatico real elevado");
    } else if (realSignal?.nivel === "Medio") {
      score += 10;
      fatores.push("Sinal climatico real moderado");
    }

    const normalizedScore = Math.min(100, score);
    const nivel =
      normalizedScore <= 30
        ? "Baixo"
        : normalizedScore <= 60
          ? "Medio"
          : "Alto";

    return {
      score: normalizedScore,
      nivel,
      sensibilidade,
      dataOperacional,
      fatores,
      distrito: locateClimateDistrict(demanda, runtime),
      riscoReal: realSignal,
    };
  }

  function sortClimateGroups(rows, keyField) {
    return rows.sort((a, b) => {
      if (b.alto !== a.alto) return b.alto - a.alto;
      if (b.quantidade !== a.quantidade) return b.quantidade - a.quantidade;
      return safeText(a[keyField]).localeCompare(safeText(b[keyField]), "pt-BR");
    });
  }

  function climateRecommendationByLevel(level) {
    if (level === "Alto") {
      return "Reavaliar janela de execucao e confirmar condicao climatica antes da mobilizacao.";
    }
    if (level === "Medio") {
      return "Manter acompanhamento climatico e validar condicao de campo.";
    }
    return "Sem restricao climatica relevante identificada no diagnostico.";
  }

  function buildClimaAnalysis(demandas, helpers = {}, contexto = {}) {
    const runtime = getClimaRuntime();
    const rows = [];
    const porDataMap = {};
    const porCentroMap = {};
    const porTipoAtividadeMap = {};
    const resumo = {
      totalAnalisado: demandas.length,
      planejadasComData: 0,
      sensiveisAoClima: 0,
      altoRiscoClimatico: 0,
      medioRiscoClimatico: 0,
      baixoRiscoClimatico: 0,
      semDataPlanejada: 0,
      semLocalOuCentro: 0,
    };

    demandas.forEach((demanda) => {
      const clima = calcularRiscoClimaticoDemanda(demanda, { ...contexto, runtime }, helpers);
      const status = getStatus(demanda, helpers);
      const criticalItem = {
        id: safeText(demanda?.id),
        ordem: safeText(demanda?.ordem),
        descricao: safeText(demanda?.descricao),
        status,
        centroTrabalho: safeText(demanda?.centroTrabalho),
        localInstalacao: safeText(demanda?.localInstalacao),
        kmInicio: safeText(demanda?.kmInicio),
        kmFim: safeText(demanda?.kmFim),
        dataOperacional: safeText(clima.dataOperacional.data),
        origemData: safeText(clima.dataOperacional.origem),
        vencimento: dateText(demanda?.vencimento, helpers),
        critico: safeText(demanda?.critico),
        sensibilidade: clima.sensibilidade.nivel,
        categoria: clima.sensibilidade.categoria,
        riscoClimatico: clima.nivel,
        score: clima.score,
        fatores: clima.fatores,
        recomendacao: climateRecommendationByLevel(clima.nivel),
      };

      rows.push(criticalItem);

      if (clima.sensibilidade.nivel !== "Baixo") resumo.sensiveisAoClima += 1;
      if (clima.nivel === "Alto") resumo.altoRiscoClimatico += 1;
      else if (clima.nivel === "Medio") resumo.medioRiscoClimatico += 1;
      else resumo.baixoRiscoClimatico += 1;
      if (clima.dataOperacional.data) resumo.planejadasComData += 1;
      else resumo.semDataPlanejada += 1;
      if (!hasKm(demanda) || !isFilled(demanda?.localInstalacao) || !isFilled(demanda?.centroTrabalho)) {
        resumo.semLocalOuCentro += 1;
      }

      const dateKey = clima.dataOperacional.data || "Sem data";
      if (!porDataMap[dateKey]) {
        porDataMap[dateKey] = { data: dateKey, quantidade: 0, alto: 0, medio: 0, baixo: 0 };
      }
      porDataMap[dateKey].quantidade += 1;
      if (clima.nivel === "Alto") porDataMap[dateKey].alto += 1;
      else if (clima.nivel === "Medio") porDataMap[dateKey].medio += 1;
      else porDataMap[dateKey].baixo += 1;

      const centerKey = safeText(demanda?.centroTrabalho) || "Sem centro";
      if (!porCentroMap[centerKey]) {
        porCentroMap[centerKey] = {
          centroTrabalho: centerKey,
          quantidade: 0,
          alto: 0,
          medio: 0,
          baixo: 0,
        };
      }
      porCentroMap[centerKey].quantidade += 1;
      if (clima.nivel === "Alto") porCentroMap[centerKey].alto += 1;
      else if (clima.nivel === "Medio") porCentroMap[centerKey].medio += 1;
      else porCentroMap[centerKey].baixo += 1;

      const categoryKey = clima.sensibilidade.categoria || "Geral";
      if (!porTipoAtividadeMap[categoryKey]) {
        porTipoAtividadeMap[categoryKey] = {
          categoria: categoryKey,
          quantidade: 0,
          alto: 0,
          medio: 0,
          baixo: 0,
        };
      }
      porTipoAtividadeMap[categoryKey].quantidade += 1;
      if (clima.nivel === "Alto") porTipoAtividadeMap[categoryKey].alto += 1;
      else if (clima.nivel === "Medio") porTipoAtividadeMap[categoryKey].medio += 1;
      else porTipoAtividadeMap[categoryKey].baixo += 1;
    });

    const porData = sortClimateGroups(Object.values(porDataMap), "data");
    const porCentro = sortClimateGroups(Object.values(porCentroMap), "centroTrabalho");
    const porTipoAtividade = sortClimateGroups(
      Object.values(porTipoAtividadeMap),
      "categoria",
    );
    const demandasCriticasClima = rows
      .sort((a, b) => {
        if (b.score !== a.score) return b.score - a.score;
        return safeText(a.dataOperacional).localeCompare(safeText(b.dataOperacional));
      })
      .slice(0, CLIMA_CRITICO_LIMIT);

    const leitura = [];
    leitura.push(
      `O recorte possui ${resumo.sensiveisAoClima} demandas com sensibilidade climatica elevada ou moderada, concentradas em atividades de campo e interferencia externa.`,
    );
    if (resumo.altoRiscoClimatico > 0) {
      leitura.push(
        `Foram identificadas ${resumo.altoRiscoClimatico} demandas com alto risco climatico operacional para a janela analisada.`,
      );
    }
    if (resumo.semLocalOuCentro > 0) {
      leitura.push(
        "A ausencia de KM ou local de instalacao em parte das demandas reduz a precisao da avaliacao climatica por trecho.",
      );
    }
    if (!runtime.climate) {
      leitura.push("Nao havia dados climaticos carregados no momento da geracao do relatorio.");
    }

    const recomendacoes = [];
    if (resumo.altoRiscoClimatico > 0) {
      recomendacoes.push({
        prioridade: "Alta",
        titulo: "Validar janela de execucao para demandas sensiveis",
        descricao:
          "Demandas com alto risco climatico exigem confirmacao de condicao de campo antes da mobilizacao.",
        acaoSugerida:
          "Revisar sequenciamento, janela operacional e contingencia de chuva para as atividades de maior score.",
      });
    }
    if (porData[0]?.quantidade > 8) {
      recomendacoes.push({
        prioridade: "Media",
        titulo: "Redistribuir concentracao por data operacional",
        descricao:
          "Foi observada concentracao de demandas sensiveis em datas especificas do recorte.",
        acaoSugerida:
          "Avaliar redistribuicao da execucao para reduzir exposicao operacional a variacao climatica.",
      });
    }
    if (
      porTipoAtividade.some((item) =>
        ["Drenagem", "Talude/Aterro"].includes(item.categoria) && item.alto > 0,
      )
    ) {
      recomendacoes.push({
        prioridade: "Alta",
        titulo: "Antecipar verificacoes de drenagem e talude",
        descricao:
          "Atividades de drenagem e talude aparecem com sensibilidade elevada no recorte atual.",
        acaoSugerida:
          "Executar inspeção preventiva e confirmar condicao de estabilidade antes de janela de chuva.",
      });
    }
    if (
      porTipoAtividade.some((item) => item.categoria === "Energia Externa" && item.quantidade > 0)
    ) {
      recomendacoes.push({
        prioridade: "Media",
        titulo: "Reforcar avaliacao de seguranca em energia externa",
        descricao:
          "Atividades de SPDA, aterramento ou iluminacao externa demandam cuidado adicional sob chuva.",
        acaoSugerida:
          "Validar seguranca eletrica e condicao de campo antes da liberacao da equipe.",
      });
    }

    return {
      disponivel: true,
      fonte: climaAvailabilityText(runtime),
      resumo,
      porData,
      porCentro,
      porTipoAtividade,
      demandasCriticasClima,
      leitura,
      recomendacoes,
    };
  }

  function buildListaCritica(demandas, helpers = {}) {
    const rows = demandas
      .map((item) => {
        const motivos = riscoReasonsFor(item, helpers);
        if (!motivos.length) return null;

        return {
          id: safeText(item.id),
          ordem: safeText(item.ordem),
          descricao: safeText(item.descricao),
          status: getStatus(item, helpers),
          substatus: getSubstatusList(item, helpers).join(" | "),
          gerencia: safeText(item.gerencia),
          supervisao: safeText(item.supervisao),
          centroTrabalho: safeText(item.centroTrabalho),
          localInstalacao: safeText(item.localInstalacao),
          kmInicio: safeText(item.kmInicio),
          kmFim: safeText(item.kmFim),
          vencimento: dateText(item.vencimento, helpers),
          dataPlanejada: dateText(item.dataPlanejada, helpers),
          dataReplanejadaAtual: dateText(item.dataReplanejadaAtual, helpers),
          toleranciaMin: dateText(item.toleranciaMin, helpers),
          toleranciaMax: dateText(item.toleranciaMax, helpers),
          critico: safeText(item.critico),
          motivo: motivos.join(" | "),
          prioridadeTecnica: motivos.length,
        };
      })
      .filter(Boolean)
      .sort((a, b) => {
        if (b.prioridadeTecnica !== a.prioridadeTecnica) {
          return b.prioridadeTecnica - a.prioridadeTecnica;
        }
        return safeText(a.vencimento).localeCompare(safeText(b.vencimento));
      })
      .slice(0, TOP_CRITICAL_LIMIT);

    return rows.map(({ prioridadeTecnica, ...item }) => item);
  }

  function mergeRecommendations(base, extra) {
    const seen = new Set();
    return [...(base || []), ...(extra || [])].filter((item) => {
      const key = `${safeText(item?.titulo)}|${safeText(item?.descricao)}`;
      if (!key || seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }

  function buildRecomendacoes(diagnosticoParcial) {
    const recomendacoes = [];
    const {
      resumo,
      planejamento,
      centros,
      planejadores,
      tolerancias,
      kms,
      clima,
    } = diagnosticoParcial;

    if (resumo.vencidas > 0) {
      recomendacoes.push({
        prioridade: "Alta",
        titulo: "Reavaliar demandas vencidas",
        descricao: `Existem ${resumo.vencidas} demandas vencidas no recorte atual.`,
        acaoSugerida: "Priorizar regularizacao, replanejamento ou encerramento tecnico das demandas vencidas.",
      });
    }

    if (resumo.criticas > 0) {
      recomendacoes.push({
        prioridade: "Alta",
        titulo: "Priorizar demandas criticas abertas",
        descricao: `Foram identificadas ${resumo.criticas} demandas classificadas como criticas.`,
        acaoSugerida: "Estabelecer sequenciamento prioritario para atividades criticas ainda abertas.",
      });
    }

    if (planejamento.diasComMaiorCarga[0]?.quantidade > 10) {
      recomendacoes.push({
        prioridade: "Media",
        titulo: "Redistribuir concentracao de carga",
        descricao: "Ha concentracao relevante de atividades em dias especificos da programacao.",
        acaoSugerida: "Avaliar redistribuicao da carga para reduzir risco de saturacao operacional.",
      });
    }

    if (centros.centrosSemCadastro.length > 0) {
      recomendacoes.push({
        prioridade: "Alta",
        titulo: "Completar cadastro de centros de trabalho",
        descricao: "Existem centros de trabalho sem cadastro consistente na carteira filtrada.",
        acaoSugerida: "Regularizar cadastro dos centros antes do fechamento do planejamento executivo.",
      });
    }

    if (tolerancias.foraJanela > 0) {
      recomendacoes.push({
        prioridade: "Media",
        titulo: "Revisar planejamento fora da tolerancia",
        descricao: `Foram encontradas ${tolerancias.foraJanela} demandas planejadas fora da janela permitida.`,
        acaoSugerida: "Revisar datas planejadas e ajustar a programacao conforme a tolerancia operacional.",
      });
    }

    if (planejadores.semPlanejador > 0) {
      recomendacoes.push({
        prioridade: "Alta",
        titulo: "Atribuir planejadores de curto",
        descricao: `Ha ${planejadores.semPlanejador} demandas sem planejador de curto definido.`,
        acaoSugerida: "Designar responsaveis para evitar perda de rastreabilidade de planejamento.",
      });
    }

    return mergeRecommendations(recomendacoes, clima?.recomendacoes || []);
  }

  function buildNarrativaExecutiva(diagnostico) {
    const paragraphs = [];
    const { resumo, centros, riscos, vencimentos, planejamento, clima } = diagnostico;

    paragraphs.push(
      `A carteira filtrada possui ${resumo.total} demandas, com ${resumo.percentualPlanejado}% do volume em programacao e ${resumo.percentualRealizado}% ja realizado.`,
    );

    if (centros.ranking[0]) {
      paragraphs.push(
        `O maior peso operacional esta concentrado no centro ${centros.ranking[0].centroTrabalho}, exigindo acompanhamento de capacidade e aderencia da execucao.`,
      );
    }

    if (vencimentos.vencidas.length || resumo.proximasVencimento20d) {
      paragraphs.push(
        `Foram identificadas ${vencimentos.vencidas.length} demandas vencidas e ${resumo.proximasVencimento20d} proximas do vencimento em 20 dias, indicando necessidade de priorizacao no curto prazo.`,
      );
    }

    paragraphs.push(
      `O nivel de risco consolidado do recorte foi classificado como ${riscos.nivelRisco.toLowerCase()}, com score ${riscos.scoreRiscoGeral}.`,
    );

    if (planejamento.diasComMaiorCarga[0]) {
      paragraphs.push(
        `A maior concentracao de carga planejada ocorre em ${planejamento.diasComMaiorCarga[0].data}, ponto que merece avaliacao de balanceamento operacional.`,
      );
    }

    if (clima?.resumo?.altoRiscoClimatico > 0) {
      paragraphs.push(
        `A leitura climatica identificou ${clima.resumo.altoRiscoClimatico} demandas com alta exposicao meteorologica para a janela operacional considerada.`,
      );
    }

    return paragraphs;
  }

  function gerarDiagnostico(demandas, contexto = {}) {
    const rows = Array.isArray(demandas) ? demandas : [];
    const helpers = contexto.helpers || {};

    const diagnostico = {
      meta: buildMeta(rows, contexto),
      resumo: buildResumo(rows, helpers),
      status: buildStatusAnalysis(rows, helpers),
      vencimentos: buildVencimentosAnalysis(rows, helpers),
      planejamento: buildPlanejamentoAnalysis(rows, helpers),
      centros: buildCentrosAnalysis(rows, helpers),
      planejadores: buildPlanejadoresAnalysis(rows, helpers),
      supervisoes: buildSupervisoesAnalysis(rows, helpers),
      gerencias: buildGerenciasAnalysis(rows, helpers),
      kms: buildKmAnalysis(rows, helpers),
      tolerancias: buildToleranciasAnalysis(rows, helpers),
      riscos: null,
      clima: null,
      listaCritica: [],
      recomendacoes: [],
      narrativaExecutiva: [],
    };

    diagnostico.riscos = buildRiscosAnalysis(rows, helpers);
    diagnostico.clima = buildClimaAnalysis(rows, helpers, contexto);
    diagnostico.listaCritica = buildListaCritica(rows, helpers);
    diagnostico.recomendacoes = buildRecomendacoes(diagnostico);
    diagnostico.narrativaExecutiva = buildNarrativaExecutiva(diagnostico);

    return diagnostico;
  }

  function initDiagnosticoEngine(options = {}) {
    engineContext = { ...options };
    console.log("Diagnostico Engine inicializado", engineContext);
    return engineContext;
  }

  function getContext() {
    return engineContext;
  }

  global.CCEDiagnosticoEngine = {
    init: initDiagnosticoEngine,
    getContext,
    gerarDiagnostico,
  };
})(window);

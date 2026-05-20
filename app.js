(function bootstrapCentral(global, document) {
  const $ = (selector, root = document) => root.querySelector(selector);
  const $$ = (selector, root = document) =>
    Array.from(root.querySelectorAll(selector));

  const STATUS_OPTIONS = [
    "A Planejar",
    "Planejado",
    "Replanejado",
    "Realizado",
    "Cancelado",
  ];

  const LARGE_BATCH_REFRESH_LIMIT = 500;

  const PROFILE_RULES = {
    Administrador: {
      planejar: true,
      replanejar: true,
      realizar: true,
      configurar: true,
      exportar: true,
      cargaLote: true,
      indicadores: true,
    },
    Editor: {
      planejar: true,
      replanejar: true,
      realizar: true,
      configurar: false,
      exportar: true,
      cargaLote: true,
      indicadores: true,
    },
    Planejador: {
      planejar: true,
      replanejar: true,
      realizar: false,
      configurar: false,
      exportar: true,
      cargaLote: true,
      indicadores: true,
    },
    Gestor: {
      planejar: false,
      replanejar: false,
      realizar: false,
      configurar: false,
      exportar: true,
      cargaLote: false,
      indicadores: true,
    },
    Visualizador: {
      planejar: false,
      replanejar: false,
      realizar: false,
      configurar: false,
      exportar: true,
      cargaLote: false,
      indicadores: false,
    },
  };

  const FILTER_DEFINITIONS = [
    { key: "gerencia", label: "Gerencia", field: "gerencia" },
    { key: "supervisao", label: "Supervisao", field: "supervisao" },
    {
      key: "centroTrabalho",
      label: "Centro Trabalho",
      field: "centroTrabalho",
    },
    { key: "tipoDemanda", label: "Tipo Demanda", field: "tipoDemanda" },
    { key: "origem", label: "Origem", field: "origem" },
    { key: "tipoOM", label: "Tipo OM", field: "tipoOM" },
    { key: "competencia", label: "Competencia", field: "competencia" },
    {
      key: "statusOperacional",
      label: "Status Operacional",
      special: "status",
    },
    { key: "prioridade", label: "Prioridade", field: "prioridade" },
    { key: "critico", label: "Crítico", field: "critico" },
    { key: "planejadorOM", label: "Planejador OM", field: "planejadorOM" },
    { key: "substatus", label: "Substatus", special: "substatus" },
    {
      key: "planejadorCurto",
      label: "Planejador Curto",
      field: "planejadorCurto",
    },
    { key: "programador", label: "Programador", field: "programador" },
    { key: "centroStatus", label: "Cadastro Centro", special: "centroStatus" },
    {
      key: "localInstalacao",
      label: "Local Instalacao",
      field: "localInstalacao",
    },
    {
      key: "statusSistema",
      label: "Status Sistema",
      field: "statusSistema",
      formatter: "sapStatus",
    },
    {
      key: "statusUsuario",
      label: "Status Usuario",
      field: "statusUsuario",
      formatter: "sapStatus",
    },
    { key: "anoVencimento", label: "Ano Vencimento", special: "anoVencimento" },
    { key: "mesVencimento", label: "Mes Vencimento", special: "mesVencimento" },
  ];

  const QUALITY_TYPE_DEFINITIONS = [
    {
      key: "om-duplicada",
      label: "OM duplicada",
      cardLabel: "OMs duplicadas",
    },
    {
      key: "id-duplicado-sem-om",
      label: "ID duplicado sem OM",
      cardLabel: "IDs duplicados sem OM",
    },
    {
      key: "sem-chave",
      label: "Demanda sem chave",
      cardLabel: "Demandas sem chave",
    },
    {
      key: "centro-sem-cadastro",
      label: "Centro sem cadastro",
      cardLabel: "Centros sem cadastro",
    },
  ];

  const NAV_GROUPS = {
    carteira: ["carteira", "lote", "notificacoes", "historico-carteira"],
    qualidade: [
      "futuras",
      "qualidade",
      "qualidade-local",
      "qualidade-centros",
      "qualidade-divergencias",
    ],
    administracao: ["administracao", "logs", "saude-integracao"],
  };

  const state = {
    repo: null,
    db: null,
    currentUser: null,
    currentView: "carteira",
    navGroups: {
      carteira: true,
      qualidade: false,
      administracao: false,
    },
    adminTab: "usuarios",
    selectedDemandId: "",
    page: 1,
    pageSize: 12,

    futurePage: 1,
    futurePageSize: 50,
    futureSearch: "",

    advancedFilters: false,
    identity: null,
    realizedAutoSynced: false,
    filters: {},
    indicatorFilters: {},
    indicatorFiltersReady: false,
    indicatorFiltersVisible: false,
    filterOptionsCache: null,
    filterSearchTimer: null,
    indicatorSearchTimer: null,
    lastDataUpdateAt: "",
    loginReady: false,
    batch: {
      rows: [],
      valid: [],
      warnings: [],
      errors: [],
      fileName: "",
    },

    notifications: {
      typeFilter: "",
      search: "",
      cache: [],
      cacheKey: "",
      byId: new Map(),
    },

    portfolioHistory: {
      search: "",
      action: "",
      user: "",
      startDate: "",
      endDate: "",
    },

    quality: {
      typeFilter: "",
      search: "",
      duplicateOmSearch: "",
      localSearch: "",
      localFilters: {},
      localFiltersVisible: false,
      localFiltersReady: false,
      localGroupsCache: null,
      selectedLocal: "",
      centersSearch: "",
      divergenceSearch: "",
      selectedIssueId: "",
      selectedPrimarySequence: "",
      issuesCache: [],
      filteredCache: [],
      page: 1,
      pageSize: 80,
    },
    actionContext: null,
  };

  function iconSvg(name) {
    const icons = {
      play: '<path d="M8 5v14l11-7z"></path>',
      sync: '<path d="M21 12a9 9 0 0 0-15.3-6.4L3 8"></path><path d="M3 3v5h5"></path><path d="M3 12a9 9 0 0 0 15.3 6.4L21 16"></path><path d="M16 16h5v5"></path>',
      check: '<path d="M20 6 9 17l-5-5"></path>',
      history:
        '<path d="M3 12a9 9 0 1 0 3-6.7"></path><path d="M3 3v6h6"></path><path d="M12 7v5l3 2"></path>',
      bell: '<path d="M18 8a6 6 0 0 0-12 0c0 7-3 7-3 9h18c0-2-3-2-3-9"></path><path d="M13.7 21a2 2 0 0 1-3.4 0"></path>',
      help: '<circle cx="12" cy="12" r="10"></circle><path d="M9.1 9a3 3 0 1 1 5.8 1c-.7 1.2-2.1 1.4-2.6 2.6"></path><path d="M12 17h.01"></path>',
      grid: '<rect x="4" y="4" width="6" height="6"></rect><rect x="14" y="4" width="6" height="6"></rect><rect x="4" y="14" width="6" height="6"></rect><rect x="14" y="14" width="6" height="6"></rect>',
      target:
        '<circle cx="12" cy="12" r="8"></circle><circle cx="12" cy="12" r="3"></circle>',
      calendar:
        '<path d="M8 2v4"></path><path d="M16 2v4"></path><rect x="3" y="5" width="18" height="16" rx="2"></rect><path d="M3 10h18"></path>',
      shield:
        '<path d="M12 2 5 5v6c0 5 3.3 9.4 7 11 3.7-1.6 7-6 7-11V5Z"></path><path d="m9 12 2 2 4-5"></path>',
      chart:
        '<path d="M4 19V5"></path><path d="M4 19h17"></path><path d="m8 16 3-5 4 3 4-8"></path>',
      settings:
        '<path d="M12 15a3 3 0 1 0 0-6 3 3 0 0 0 0 6Z"></path><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.9l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.9-.3 1.7 1.7 0 0 0-1 1.6V22a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1-1.6 1.7 1.7 0 0 0-1.9.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.9 1.7 1.7 0 0 0-1.6-1H2a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.6-1 1.7 1.7 0 0 0-.3-1.9l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.9.3 1.7 1.7 0 0 0 1-1.6V2a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.6 1.7 1.7 0 0 0 1.9-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.9 1.7 1.7 0 0 0 1.6 1H22a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.6 1Z"></path>',
      logs: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8Z"></path><path d="M14 2v6h6"></path><path d="M8 13h8"></path><path d="M8 17h5"></path>',
      "chevron-left": '<path d="m15 18-6-6 6-6"></path>',
      "chevron-down": '<path d="m6 9 6 6 6-6"></path>',
      filter:
        '<path d="M3 5h18"></path><path d="M6 12h12"></path><path d="M10 19h4"></path>',
    };
    return `<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">${icons[name] || icons.help}</svg>`;
  }

  function renderStaticIcons() {
    $$("[data-icon]").forEach((element) => {
      element.innerHTML = iconSvg(element.dataset.icon);
    });
  }

  function navGroupForView(view) {
    return Object.entries(NAV_GROUPS).find(([, views]) =>
      views.includes(view),
    )?.[0];
  }

  function syncNavigation(view) {
    $$("[data-nav-group]").forEach((group) => {
      const groupKey = group.dataset.navGroup;
      const isOpen = Boolean(state.navGroups[groupKey]);
      group.classList.toggle("is-open", Boolean(isOpen));
      const toggle = group.querySelector("[data-nav-group-toggle]");
      toggle?.setAttribute("aria-expanded", isOpen ? "true" : "false");
    });

    $$(".nav-item").forEach((item) => {
      const groupKey = item.dataset.navGroupToggle;
      const isActive = groupKey
        ? NAV_GROUPS[groupKey]?.includes(view)
        : item.dataset.view === view;
      item.classList.toggle("is-active", Boolean(isActive));
    });

    $$(".nav-subitem").forEach((item) => {
      item.classList.toggle("is-active", item.dataset.view === view);
    });
  }

  function todayText() {
    return new Date().toISOString().slice(0, 10);
  }

  function escapeHtml(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function compact(value, fallback = "-") {
    return value || fallback;
  }

  function toDate(value) {
    if (!value) return null;
    if (value instanceof Date)
      return Number.isNaN(value.getTime()) ? null : value;
    if (typeof value === "number" && Number.isFinite(value)) {
      const date = excelSerialToDate(value);
      return date && !Number.isNaN(date.getTime()) ? date : null;
    }
    const text = String(value).trim();
    if (!text) return null;

    if (/^\d{5}$/.test(text)) {
      const date = excelSerialToDate(Number(text));
      return date && !Number.isNaN(date.getTime()) ? date : null;
    }

    let normalized = text;
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(text)) {
      const [day, month, year] = text.split("/");
      normalized = `${year}-${month}-${day}`;
    } else if (/^\d{2}\/\d{2}\/\d{2}$/.test(text)) {
      const [day, month, year] = text.split("/");
      normalized = `20${year}-${month}-${day}`;
    } else if (/^\d{4}-\d{2}$/.test(text)) {
      normalized = `${text}-01`;
    } else if (/^\d{6}$/.test(text) && text.startsWith("20")) {
      normalized = `${text.slice(0, 4)}-${text.slice(4, 6)}-01`;
    }

    const date = new Date(`${normalized.slice(0, 10)}T12:00:00`);
    return Number.isNaN(date.getTime()) ? null : date;
  }

  function formatDate(value) {
    const date = toDate(value);
    if (!date) return "-";
    return new Intl.DateTimeFormat("pt-BR", { timeZone: "UTC" }).format(date);
  }

  function formatDateTime(value) {
    const date = value ? new Date(value) : null;
    if (!date || Number.isNaN(date.getTime())) return "-";
    return new Intl.DateTimeFormat("pt-BR", {
      dateStyle: "short",
      timeStyle: "short",
    }).format(date);
  }

  function dateText(value) {
    const date = toDate(value);
    return date ? date.toISOString().slice(0, 10) : "";
  }

  function monthName(month) {
    const names = [
      "Jan",
      "Fev",
      "Mar",
      "Abr",
      "Mai",
      "Jun",
      "Jul",
      "Ago",
      "Set",
      "Out",
      "Nov",
      "Dez",
    ];
    return names[Number(month) - 1] || month;
  }

  function excelSerialToDate(value) {
    const serial = Number(value);
    if (!Number.isFinite(serial) || serial < 20000 || serial > 80000)
      return null;
    const epoch = Date.UTC(1899, 11, 30);
    return new Date(epoch + serial * 86400000);
  }

  function normalizeCompetencia(value) {
    if (value === null || value === undefined || value === "") return "";
    if (typeof value === "number" && Number.isFinite(value)) {
      const date = excelSerialToDate(value);
      return date ? date.toISOString().slice(0, 7) : "";
    }

    const text = String(value).trim();
    if (!text) return "";

    if (/^\d{4}-\d{2}$/.test(text)) return text;
    if (/^\d{6}$/.test(text) && text.startsWith("20")) {
      return `${text.slice(0, 4)}-${text.slice(4, 6)}`;
    }
    if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text.slice(0, 7);
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(text)) {
      const [, month, year] = text.split("/");
      return `${year}-${month}`;
    }
    if (/^\d{5}$/.test(text)) {
      const date = excelSerialToDate(Number(text));
      return date ? date.toISOString().slice(0, 7) : "";
    }

    const monthAliases = {
      JAN: "01",
      JANEIRO: "01",
      FEB: "02",
      FEV: "02",
      FEVEREIRO: "02",
      MAR: "03",
      MARCO: "03",
      MARÇO: "03",
      APR: "04",
      ABR: "04",
      ABRIL: "04",
      MAY: "05",
      MAI: "05",
      MAIO: "05",
      JUN: "06",
      JUNHO: "06",
      JUL: "07",
      JULHO: "07",
      AUG: "08",
      AGO: "08",
      AGOSTO: "08",
      SEP: "09",
      SET: "09",
      SETEMBRO: "09",
      OCT: "10",
      OUT: "10",
      OUTUBRO: "10",
      NOV: "11",
      NOVEMBRO: "11",
      DEC: "12",
      DEZ: "12",
      DEZEMBRO: "12",
    };
    const normalized = normalizeText(text).replace(/[^A-Z0-9]+/g, " ");
    const parts = normalized.split(" ").filter(Boolean);
    const year = parts.find((part) => /^20\d{2}$/.test(part));
    const month = parts.map((part) => monthAliases[part]).find(Boolean);
    return year && month ? `${year}-${month}` : text;
  }

  function normalizePrioridade(value) {
    const text = String(value ?? "").trim();
    if (!text) return "Nao informado";
    const normalized = normalizeText(text);
    if (/^1($|[^0-9])/.test(normalized) || normalized.includes("ALTO"))
      return "Alto";
    if (
      /^2($|[^0-9])/.test(normalized) ||
      normalized.includes("MEDIO") ||
      normalized.includes("MEDIA")
    )
      return "Medio";
    if (/^3($|[^0-9])/.test(normalized) || normalized.includes("BAIXO"))
      return "Baixo";
    return text;
  }

  function normalizeCritico(value) {
    const text = String(value ?? "").trim();

    if (!text) return "Não informado";

    const normalized = normalizeText(text);

    if (
      normalized === "SIM" ||
      normalized === "S" ||
      normalized === "TRUE" ||
      normalized === "1" ||
      normalized === "X" ||
      normalized.includes("CRITICO") ||
      normalized.includes("CRITICA")
    ) {
      return "Sim";
    }

    if (
      normalized === "NAO" ||
      normalized === "NÃO" ||
      normalized === "N" ||
      normalized === "FALSE" ||
      normalized === "0"
    ) {
      return "Não";
    }

    return text;
  }

  function profileDefaults(profile) {
    return PROFILE_RULES[profile] || PROFILE_RULES.Visualizador;
  }

  function getRules() {
    const defaults = profileDefaults(state.currentUser?.perfil);
    const user = state.currentUser || {};
    return {
      ...defaults,
      planejar: user.permissaoPlanejar ?? defaults.planejar,
      replanejar: user.permissaoReplanejar ?? defaults.replanejar,
      realizar: user.permissaoRealizar ?? defaults.realizar,
      configurar: user.permissaoConfigurar ?? defaults.configurar,
      exportar: user.permissaoExportar ?? defaults.exportar,
      cargaLote: user.permissaoCargaLote ?? defaults.cargaLote,
    };
  }

  function configItems(group) {
    const raw = state.db?.configuracoes?.[group] || [];
    return raw
      .map((item) =>
        typeof item === "string"
          ? { id: global.CCEData.slugify(item), nome: item, ativo: true }
          : item,
      )
      .filter((item) => item.ativo !== false);
  }

  function configNames(group) {
    return configItems(group).map((item) => item.nome);
  }

  function configKeyByName(group, name) {
    const item = configItems(group).find(
      (config) => config.nome === name || config.id === name,
    );
    return item?.id || global.CCEData.slugify(name || "").toUpperCase();
  }

  function childConfigNames(group, parentKey, parentIdOrName) {
    const parent = configItems(parentKey).find(
      (item) => item.id === parentIdOrName || item.nome === parentIdOrName,
    );
    const parentId = parent?.id || parentIdOrName;
    const field = group === "justificativas" ? "motivoId" : "perfilId";
    const items = configItems(group).filter(
      (item) =>
        !parentId || item[field] === parentId || item[field] === parent?.nome,
    );
    return items.length ? items.map((item) => item.nome) : configNames(group);
  }

  function canEdit() {
    const rules = getRules();
    return rules.planejar || rules.replanejar || rules.realizar;
  }

  function canAdmin() {
    return getRules().configurar;
  }

  function canExport() {
    return getRules().exportar;
  }

  function canBatch() {
    return getRules().cargaLote;
  }

  function canPlan() {
    return getRules().planejar;
  }

  function canReplan() {
    return getRules().replanejar;
  }

  function canRealizar() {
    return getRules().realizar;
  }

  function replanHistoryIssues(demand) {
    const history = state.db?.historicoReplanejamento || [];
    return history
      .filter((item) => item.demandaId === demand.id)
      .flatMap((item) => {
        const issues = [];
        if (!item.motivo) issues.push("Sem motivo replanejamento");
        if (!item.justificativa)
          issues.push("Sem justificativa replanejamento");
        return issues;
      });
  }

  function dueClassOf(demand) {
    if (!demand.dataRealizada) return "";
    const realized = toDate(demand.dataRealizada);
    const min = toDate(demand.toleranciaMin);
    const max = toDate(demand.toleranciaMax || demand.vencimento);
    if ((min && realized < min) || (max && realized > max))
      return "Fora do Prazo";
    return "No Prazo";
  }

  function hasSapStatusText(demand, token) {
    const statusSistema = normalizeText(demand.statusSistema);
    const statusUsuario = normalizeText(demand.statusUsuario);
    const alvo = normalizeText(token);

    return statusSistema.includes(alvo) || statusUsuario.includes(alvo);
  }

  function isCanceledBySap(demand) {
    return hasSapStatusText(demand, "CANC");
  }

  function hasSapSystemClosureStatus(demand) {
    const statusSistema = normalizeText(demand.statusSistema);
    return statusSistema.includes("ENTE") || statusSistema.includes("ENCE");
  }

  function isRealizedBySapStatus(demand) {
    const statusSistema = normalizeText(demand.statusSistema);
    const statusUsuario = normalizeText(demand.statusUsuario);

    return (
      statusSistema.includes("ENTE") ||
      statusSistema.includes("ENCE") ||
      statusUsuario.includes("ENCR") ||
      statusUsuario.includes("ENTE") ||
      statusUsuario.includes("ENCE")
    );
  }

  function hasRealizedDate(demand) {
    return Boolean(dateText(demand.dataRealizada));
  }

  function isWaitingClosure(demand) {
    if (isCanceledBySap(demand)) return false;
    if (!hasRealizedDate(demand)) return false;

    const statusSistema = normalizeText(demand.statusSistema);
    const hasOpenSystemStatus =
      statusSistema.includes("LIB") || statusSistema.includes("CONF");

    return hasOpenSystemStatus && !hasSapSystemClosureStatus(demand);
  }

  function isWaitingTechnicalClosure(demand) {
    if (isCanceledBySap(demand)) return false;
    if (!hasRealizedDate(demand)) return false;
    if (hasSapSystemClosureStatus(demand)) return false;

    const statusSistema = normalizeText(demand.statusSistema);

    return (
      statusSistema.includes("LIB") ||
      statusSistema.includes("CONF")
    );
  }

  function hasLibConfStatus(demand) {
    const statusSistema = normalizeText(demand.statusSistema);
    const statusUsuario = normalizeText(demand.statusUsuario);
    const statusText = `${statusSistema} ${statusUsuario}`;
    return statusText.includes("LIB") && statusText.includes("CONF");
  }

  function primaryStatusOf(demand) {
    if (isCanceledBySap(demand)) {
      return "Cancelado";
    }

    if (isRealizedBySapStatus(demand)) {
      return "Realizado";
    }

    if (demand.dataRealizada && hasLibConfStatus(demand)) {
      return "Realizado";
    }

    if (demand.dataReplanejadaAtual) {
      return "Replanejado";
    }

    if (demand.dataPlanejada) {
      return "Planejado";
    }

    return "A Planejar";
  }

  function pendingIssuesOf(demand) {
    const issues = [];
    if (demand.perda && !demand.motivoPerda) issues.push("Sem motivo de perda");
    if (demand.perda && !demand.justificativaPerda)
      issues.push("Sem justificativa de perda");
    issues.push(...replanHistoryIssues(demand));
    return Array.from(new Set(issues));
  }

  function substatusListOf(demand) {
    const status = primaryStatusOf(demand);
    const substatuses = [];

    if (status === "Cancelado") {
      substatuses.push("Cancelado");
      return Array.from(new Set(substatuses));
    }

    if (status === "Realizado" && !hasRealizedDate(demand)) {
      substatuses.push("Encerrado no SAP BO sem data realizada");
    }

    if (status === "Realizado" && isWaitingClosure(demand)) {
      substatuses.push("Ag Encerramento");
    }

    if (hasRealizedDate(demand) && !hasLibConfStatus(demand)) {
      substatuses.push("Avaliar Status no SAP");
    }

    if (demand.perda) {
      substatuses.push("Perda");
    }

    if (status === "Realizado" && !isWaitingClosure(demand)) {
      const dueClass = dueClassOf(demand);
      if (dueClass) substatuses.push(dueClass);
    }

    if (status !== "Realizado") {
      const dueClass = dueClassOf(demand);
      if (dueClass) substatuses.push(dueClass);
    }

    if (pendingIssuesOf(demand).length) {
      substatuses.push("Pendente");
    }

    return Array.from(new Set(substatuses));
  }

  function statusListOf(demand) {
    return [primaryStatusOf(demand), ...substatusListOf(demand)];
  }

  function statusOf(demand) {
    return primaryStatusOf(demand);
  }

  function prepareDemandForSave(demand, extra = {}) {
    const prepared = normalizeDemandRecord({
      ...demand,
      ...extra,
      dataUltimaAtualizacao: new Date().toISOString(),
    });
    prepared.statusOperacional = primaryStatusOf(prepared);
    prepared.substatusOperacional = substatusListOf(prepared).join(" | ");
    prepared.usuarioResponsavel =
      prepared.usuarioResponsavel || state.currentUser?.email || "";
    return prepared;
  }

  function statusClass(status) {
    if (status === "A Planejar") return "status-planejar";
    if (status === "Planejado") return "status-planejado";
    if (status === "Replanejado") return "status-replanejado";
    if (status === "Realizado" || status === "No Prazo")
      return "status-realizado";
    if (status === "Cancelado" || status === "Cancelado") return "status-perda";
    if (status === "Cadastrado") return "status-realizado";
    if (status === "Nao cadastrado" || status === "Sem centro")
      return "status-perda";
    if (status === "Fora do Prazo") return "status-fora-prazo";
    if (status === "Avaliar Status no SAP") return "status-fora-prazo";
    if (status === "Perda" || status === "Pendente") return "status-perda";
    return "status-planejado";
  }

  function statusChip(status) {
    return `<span class="status-chip ${statusClass(status)}">${escapeHtml(status)}</span>`;
  }

  function statusChipGroup(statuses) {
    return statuses.length
      ? `<span class="status-stack">${statuses.map(statusChip).join("")}</span>`
      : '<span class="muted">-</span>';
  }

  function allowedActionsFor(demand) {
    const status = primaryStatusOf(demand);

    return {
      planejar: status === "A Planejar",
      replanejar: status === "Planejado" || status === "Replanejado",
      realizado: status !== "Cancelado",
      historico: true,
    };
  }

  function actionButton(action, id, disabled = false, withText = false) {
    const meta = {
      planejar: ["Planejar", "play", "primary"],
      replanejar: ["Replanejar", "sync", "warning"],
      realizado: ["Realizado/Perda", "check", "success"],
      historico: ["Histórico", "history", "neutral"],
    }[action];
    const label = meta[0];
    return `
      <button class="${withText ? "detail-action-button" : "icon-action"} ${meta[2]}" data-action="${action}" data-id="${escapeHtml(id)}" aria-label="${label}" title="${label}" ${disabled ? "disabled" : ""}>
        ${iconSvg(meta[1])}
        ${withText ? `<span>${label}</span>` : ""}
      </button>
    `;
  }

  function normalizeText(value) {
    return global.CCEData.normalizeText(value);
  }

  function formatSapStatusFilter(value) {
    const text = String(value || "")
      .replace(/_/g, " ")
      .replace(/\s+/g, " ")
      .trim();

    if (!text) return "";

    return text.slice(0, 9).trim();
  }

  function uniqueOptions(values) {
    return Array.from(new Set(values.filter(Boolean))).sort((a, b) =>
      String(a).localeCompare(String(b), "pt-BR"),
    );
  }

  function populateSelect(element, values, allLabel = "Todos") {
    if (!element) return;
    const previous = element.value;
    element.innerHTML = [`<option value="">${allLabel}</option>`]
      .concat(
        values.map(
          (value) =>
            `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`,
        ),
      )
      .join("");
    if (values.includes(previous)) {
      element.value = previous;
    }
  }

  function showToast(message, type = "") {
    const toast = document.createElement("div");
    toast.className = `toast ${type}`;
    toast.textContent = message;
    $("#toastHost").appendChild(toast);
    setTimeout(() => toast.remove(), 3600);
  }

  function showBatchStatus(type, title, detail = "") {
    const panel = $(".validation-panel");
    if (!panel) return;

    let bar = $("#batchStatusBar");

    if (!bar) {
      bar = document.createElement("div");
      bar.id = "batchStatusBar";
      panel.prepend(bar);
    }

    bar.className = `batch-status-bar ${type || ""}`;
    bar.innerHTML = `
    <div class="batch-status-icon">
      ${
        type === "success"
          ? iconSvg("check")
          : type === "error"
            ? iconSvg("help")
            : iconSvg("sync")
      }
    </div>

    <div class="batch-status-content">
      <strong>${escapeHtml(title)}</strong>
      ${detail ? `<span>${escapeHtml(detail)}</span>` : ""}
    </div>
  `;

    bar.classList.remove("hidden");
  }

  function downloadFile(filename, text, type = "text/csv;charset=utf-8") {
    const blob = new Blob([text], { type });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    link.click();
    URL.revokeObjectURL(url);
  }

  function toCsv(rows) {
    const escapeCell = (value) =>
      `"${String(value ?? "").replace(/"/g, '""')}"`;
    return rows.map((row) => row.map(escapeCell).join(";")).join("\n");
  }

  async function fetchJsonArray(baseUrl, label) {
    const response = await fetch(`${baseUrl}?v=${Date.now()}`, {
      method: "GET",
      cache: "no-store",
      headers: {
        Accept: "application/json",
      },
    });

    if (!response.ok) {
      throw new Error(`Erro ao carregar ${label}: ${response.status}`);
    }

    const data = await response.json();

    if (!Array.isArray(data)) {
      throw new Error(`${label} precisa ser um array JSON.`);
    }

    return data;
  }

  function mapBaseItemToDemand(item, origemPadrao) {
    const ordem = String(
      item.OrdemSAP || item.Ordem || item.ordem_sap || item.ordem || "",
    ).trim();

    const idDemandaInformado = String(
      item.ID_Demanda_Controle ||
        item.id_demanda_controle ||
        item.IdDemandaControle ||
        item.id ||
        "",
    ).trim();

    const vencimentoBase =
      item.Vencimento ||
      item.vencimento ||
      item.DataVencimento ||
      item.data_vencimento ||
      "";

    const competenciaInformada =
      item.Competencia ||
      item.competencia ||
      item.COMPETENCIA ||
      item.Competência ||
      item.competência ||
      "";

    const origemNormalizada = normalizeText(origemPadrao);

    const competenciaBase =
      origemNormalizada.includes("ORDENS") ||
      origemNormalizada.includes("REALIZADOS")
        ? vencimentoBase
        : competenciaInformada;

    const competenciaNormalizada = normalizeCompetencia(competenciaBase);
    const vencimentoNormalizado =
      dateText(vencimentoBase) || vencimentoBase || "";

    const idDemanda =
      idDemandaInformado ||
      (ordem
        ? `DEM-SAP-${ordem}`
        : global.CCEData.stableDemandId({
            ordem: "",
            descricao: item.Descricao || item.descricao || "",
            centroTrabalho: item.CentroTrabalho || item.centro_trabalho || "",
            localInstalacao:
              item.LocalInstalacao || item.local_instalacao || "",
            competencia: competenciaNormalizada,
            vencimento: vencimentoNormalizado,
            origem: origemPadrao,
          }));

    return {
      id: String(idDemanda).trim(),
      idDemandaInformado,

      ordem,
      tipoDemanda:
        item.TipoDemanda ||
        item.tipo_demanda ||
        (origemPadrao.includes("Futuras") ? "Futura" : ""),

      tipoOM: item.TipoOM || item.tipo_om || "",
      descricao: item.Descricao || item.descricao || "",
      gerencia: item.Gerencia || item.gerencia || "",
      supervisao: item.Supervisao || item.supervisao || "",
      centroTrabalho: item.CentroTrabalho || item.centro_trabalho || "",
      localInstalacao: item.LocalInstalacao || item.local_instalacao || "",

      statusSistema: item.StatusSistema || item.status_sistema || "",
      statusUsuario: item.StatusUsuario || item.status_usuario || "",

      competencia: competenciaNormalizada,
      dataRealizada: item.DataRealizada || item.data_realizada || "",
      vencimento: vencimentoNormalizado,

      prioridade: normalizePrioridade(item.Prioridade || item.prioridade),

      critico: normalizeCritico(
        item.Critico ||
          item.Crítico ||
          item.critico ||
          item.crítico ||
          item.CRITICO ||
          item.CRÍTICO ||
          "",
      ),

      toleranciaMin: item.ToleranciaMin || item.tolerancia_min || "",
      toleranciaMax: item.ToleranciaMax || item.tolerancia_max || "",

      dataPlanejada: item.DataPlanejada || item.data_planejada || "",
      dataReplanejadaAtual:
        item.DataReplanejada ||
        item.DataReplanejadaAtual ||
        item.data_replanejada ||
        "",

      perda: item.Perda === true || item.Perda === "Sim" || item.perda === true,
      motivoPerda: item.MotivoPerda || item.motivo_perda || "",
      justificativaPerda:
        item.JustificativaPerda || item.justificativa_perda || "",

      comentario: item.Comentario || item.comentario || "",
      usuarioResponsavel:
        item.UsuarioResponsavel || item.usuario_responsavel || "",

      dataUltimaAtualizacao:
        item.DataUltimaAtualizacao || item.data_ultima_atualizacao || "",

      origem: item.Origem || item.origem || origemPadrao,

      quantidadeReplanejamentos:
        Number(
          item.QuantidadeReplanejamentos ||
            item.quantidade_replanejamentos ||
            0,
        ) || 0,

      frequencia: item.Frequencia || item.frequencia || "",
      observacao: item.Observacao || item.observacao || "",
      vinculadaEm: item.VinculadaEm || item.vinculada_em || "",
    };
  }

  function mergeBaseOrdensEFuturas(baseOrdens, baseFuturas) {
    const mapa = new Map();
    const futurasPorOrdem = new Map();

    const ordemKey = (item) => String(item?.ordem || "").trim();

    const mergeOrdemComFutura = (ordemItem, futuraItem) => {
      const merged = {
        ...futuraItem,

        // REGRA PRINCIPAL:
        // se a OM existe na futura e também na base ordens,
        // o ID oficial fica sendo o ID da futura.
        id: futuraItem.id || ordemItem.id,

        idDemandaInformado:
          futuraItem.idDemandaInformado ||
          futuraItem.id ||
          ordemItem.idDemandaInformado ||
          "",

        ordem: ordemItem.ordem || futuraItem.ordem || "",

        // A futura continua sendo a origem conceitual da demanda,
        // mas recebe os dados atualizados da ordem SAP.
        tipoDemanda:
          futuraItem.tipoDemanda || ordemItem.tipoDemanda || "Futura",

        // Campos próprios da demanda futura não devem ser perdidos.
        frequencia: futuraItem.frequencia || ordemItem.frequencia || "",
        observacao: futuraItem.observacao || ordemItem.observacao || "",
        vinculadaEm: futuraItem.vinculadaEm || ordemItem.vinculadaEm || "",

        // Campos operacionais já existentes na futura também devem ser preservados
        // quando a base ordens não trouxer valor.
        dataPlanejada:
          futuraItem.dataPlanejada || ordemItem.dataPlanejada || "",
        dataReplanejadaAtual:
          futuraItem.dataReplanejadaAtual ||
          ordemItem.dataReplanejadaAtual ||
          "",
        dataRealizada:
          ordemItem.dataRealizada || futuraItem.dataRealizada || "",

        statusSistema:
          ordemItem.statusSistema || futuraItem.statusSistema || "",

        statusUsuario:
          ordemItem.statusUsuario || futuraItem.statusUsuario || "",
        perda: futuraItem.perda ?? ordemItem.perda ?? false,
        motivoPerda: futuraItem.motivoPerda || ordemItem.motivoPerda || "",
        justificativaPerda:
          futuraItem.justificativaPerda || ordemItem.justificativaPerda || "",
        comentario: futuraItem.comentario || ordemItem.comentario || "",
        usuarioResponsavel:
          futuraItem.usuarioResponsavel || ordemItem.usuarioResponsavel || "",

        quantidadeReplanejamentos:
          futuraItem.quantidadeReplanejamentos ??
          ordemItem.quantidadeReplanejamentos ??
          0,

        dataUltimaAtualizacao:
          ordemItem.dataUltimaAtualizacao ||
          futuraItem.dataUltimaAtualizacao ||
          "",

        origem: "SAP BO - Ordens | SAP BO - Demandas Futuras",

        fontesConsolidadas: Array.from(
          new Set(
            [
              futuraItem.fontesConsolidadas,
              ordemItem.fontesConsolidadas,
              futuraItem.origem,
              ordemItem.origem,
              "SAP BO - Demandas Futuras",
              "SAP BO - Ordens",
            ]
              .filter(Boolean)
              .flatMap((item) =>
                String(item)
                  .split("|")
                  .map((value) => value.trim()),
              )
              .filter(Boolean),
          ),
        ).join(" | "),
      };

      return normalizeDemandRecord(merged);
    };

    // Primeiro entram as futuras, porque elas têm o ID de controle correto.
    baseFuturas.forEach((item) => {
      if (!item.id) return;

      mapa.set(item.id, item);

      const ordem = ordemKey(item);
      if (ordem && !futurasPorOrdem.has(ordem)) {
        futurasPorOrdem.set(ordem, item);
      }
    });

    // Depois entram as ordens.
    // Se a ordem tiver a mesma OM de uma futura, mantém o ID da futura.
    baseOrdens.forEach((item) => {
      if (!item.id) return;

      const ordem = ordemKey(item);
      const futuraPorMesmaOrdem = ordem ? futurasPorOrdem.get(ordem) : null;
      const futuraPorMesmoId = mapa.get(item.id);

      if (futuraPorMesmaOrdem) {
        const merged = mergeOrdemComFutura(item, futuraPorMesmaOrdem);

        // Garante que não fique um DEM-SAP duplicado para a mesma OM.
        mapa.delete(item.id);

        mapa.set(merged.id, merged);
        return;
      }

      if (futuraPorMesmoId) {
        const merged = mergeOrdemComFutura(item, futuraPorMesmoId);
        mapa.set(merged.id, merged);
        return;
      }

      mapa.set(item.id, item);
    });

    return Array.from(mapa.values());
  }
  function demandIsRealizada(demanda) {
    return primaryStatusOf(demanda) === "Realizado";
  }

  function demandHasFutureSource(demanda) {
    return normalizeText(
      [
        demanda?.origem,
        demanda?.fontesConsolidadas,
        demanda?.tipoDemanda,
      ].join(" "),
    ).includes("FUTUR");
  }

  function mergeCarteiraDuplicateByOrder(principal, secundario) {
    const merged = { ...secundario, ...principal };

    const camposComplementares = [
      "descricao",
      "gerencia",
      "supervisao",
      "centroTrabalho",
      "localInstalacao",
      "competencia",
      "tipoOM",
      "prioridade",
      "critico",
      "vencimento",
      "toleranciaMin",
      "toleranciaMax",
      "frequencia",
      "observacao",
      "vinculadaEm",

      // campos SAP/realização que podem vir da base ordens ou realizados
      "statusSistema",
      "statusUsuario",
      "dataRealizada",
      "origemRealizacao",
    ];

    camposComplementares.forEach((campo) => {
      if (
        (merged[campo] === null ||
          merged[campo] === undefined ||
          String(merged[campo]).trim() === "") &&
        secundario[campo]
      ) {
        merged[campo] = secundario[campo];
      }
    });

    const camposControleSupabase = [
      "dataPlanejada",
      "dataReplanejadaAtual",
      "dataRealizada",
      "perda",
      "motivoPerda",
      "justificativaPerda",
      "comentario",
      "usuarioResponsavel",
      "quantidadeReplanejamentos",
    ];

    camposControleSupabase.forEach((campo) => {
      if (
        secundario[campo] !== null &&
        secundario[campo] !== undefined &&
        String(secundario[campo]).trim() !== "" &&
        String(secundario[campo]).trim() !== "false"
      ) {
        if (
          merged[campo] === null ||
          merged[campo] === undefined ||
          String(merged[campo]).trim() === "" ||
          merged[campo] === false
        ) {
          merged[campo] = secundario[campo];
        }
      }
    });

    if (isCanceledBySap(secundario) && !isCanceledBySap(merged)) {
      merged.statusSistema = secundario.statusSistema || merged.statusSistema;
      merged.statusUsuario = secundario.statusUsuario || merged.statusUsuario;
    }

    merged.fontesConsolidadas = Array.from(
      new Set(
        [
          principal.fontesConsolidadas,
          secundario.fontesConsolidadas,
          principal.origem,
          secundario.origem,
        ]
          .filter(Boolean)
          .flatMap((item) =>
            String(item)
              .split("|")
              .map((v) => v.trim()),
          )
          .filter(Boolean),
      ),
    ).join(" | ");

    merged.statusOperacional = primaryStatusOf(merged);
    merged.substatusOperacional = substatusListOf(merged).join(" | ");

    return normalizeDemandRecord(merged);
  }

  function consolidateCarteiraByRealizedOrder(carteira) {
    const semOrdem = [];
    const porOrdem = new Map();

    carteira.forEach((demanda) => {
      const ordem = String(demanda.ordem || "").trim();

      if (!ordem) {
        semOrdem.push(demanda);
        return;
      }

      const atual = porOrdem.get(ordem);

      if (!atual) {
        porOrdem.set(ordem, demanda);
        return;
      }

      const atualRealizada = demandIsRealizada(atual);
      const novaRealizada = demandIsRealizada(demanda);
      const atualCancelada = isCanceledBySap(atual);
      const novaCancelada = isCanceledBySap(demanda);
      const atualVemDaFutura = demandHasFutureSource(atual);
      const novaVemDaFutura = demandHasFutureSource(demanda);

      let principal = atual;
      let secundario = demanda;

      if (novaVemDaFutura && !atualVemDaFutura) {
        principal = demanda;
        secundario = atual;
      } else if (!novaVemDaFutura && atualVemDaFutura) {
        principal = atual;
        secundario = demanda;
      } else if (novaCancelada && !atualCancelada) {
        principal = demanda;
        secundario = atual;
      } else if (!novaCancelada && atualCancelada) {
        principal = atual;
        secundario = demanda;
      } else if (novaRealizada && !atualRealizada) {
        principal = demanda;
        secundario = atual;
      } else if (novaRealizada === atualRealizada) {
        const atualTemSupabase =
          String(atual.origem || "").includes("Supabase") ||
          atual.dataPlanejada ||
          atual.dataReplanejadaAtual ||
          atual.comentario;

        const novaTemSupabase =
          String(demanda.origem || "").includes("Supabase") ||
          demanda.dataPlanejada ||
          demanda.dataReplanejadaAtual ||
          demanda.comentario;

        if (novaTemSupabase && !atualTemSupabase) {
          principal = demanda;
          secundario = atual;
        }
      }

      porOrdem.set(ordem, mergeCarteiraDuplicateByOrder(principal, secundario));
    });

    return [...semOrdem, ...Array.from(porOrdem.values())];
  }

  function applyBaseRealizadosToCarteira(carteira, baseRealizados) {
    if (!Array.isArray(baseRealizados) || !baseRealizados.length) {
      return carteira;
    }

    const realizadosPorOrdem = new Map();

    baseRealizados.forEach((realizado) => {
      const ordem = String(realizado.ordem || "").trim();
      if (!ordem) return;

      const atual = realizadosPorOrdem.get(ordem);

      if (!atual) {
        realizadosPorOrdem.set(ordem, realizado);
        return;
      }

      const dataAtual = toDate(atual.dataRealizada);
      const dataNova = toDate(realizado.dataRealizada);

      if (dataNova && (!dataAtual || dataNova > dataAtual)) {
        realizadosPorOrdem.set(ordem, realizado);
      }
    });

    const ordensExistentesNaCarteira = new Set(
      carteira
        .map((demanda) => String(demanda.ordem || "").trim())
        .filter(Boolean),
    );

    const carteiraAtualizada = carteira.map((demanda) => {
      const ordem = String(demanda.ordem || "").trim();
      if (!ordem) return demanda;

      const realizado = realizadosPorOrdem.get(ordem);
      if (!realizado) return demanda;

      const dataRealizada =
        realizado.dataRealizada || demanda.dataRealizada || "";

      const atualizado = {
        ...demanda,

        statusSistema: realizado.statusSistema || demanda.statusSistema || "",

        statusUsuario: realizado.statusUsuario || demanda.statusUsuario || "",

        dataRealizada,

        origemRealizacao: realizado.dataRealizada
          ? "SAP BO - Realizados"
          : demanda.dataRealizada
            ? demanda.origemRealizacao || "SAP BO - Ordens/Futuras"
            : "",
        dataUltimaAtualizacao:
          realizado.dataUltimaAtualizacao ||
          demanda.dataUltimaAtualizacao ||
          new Date().toISOString(),

        fontesConsolidadas: [
          demanda.origem || "Carteira",
          "SAP BO - Realizados",
        ]
          .filter(Boolean)
          .join(" | "),
      };

      atualizado.statusOperacional = primaryStatusOf(atualizado);
      atualizado.substatusOperacional = substatusListOf(atualizado).join(" | ");

      return normalizeDemandRecord(atualizado);
    });

    const realizadosSomenteNaBase = [];

    realizadosPorOrdem.forEach((realizado, ordem) => {
      if (ordensExistentesNaCarteira.has(ordem)) return;

      const dataRealizada = realizado.dataRealizada || "";

      const novoRealizado = {
        ...realizado,

        id: realizado.id || `DEM-SAP-${ordem}`,
        idDemandaInformado:
          realizado.idDemandaInformado || realizado.id || `REAL-SAP-${ordem}`,

        ordem,
        tipoDemanda: realizado.tipoDemanda || "Realizada",
        origem: "SAP BO - Realizados",

        statusSistema: realizado.statusSistema || "",
        statusUsuario: realizado.statusUsuario || "",
        dataRealizada,

        origemRealizacao: "SAP BO - Realizados",
        dataUltimaAtualizacao:
          realizado.dataUltimaAtualizacao || new Date().toISOString(),

        fontesConsolidadas: "SAP BO - Realizados",

        dataPlanejada: realizado.dataPlanejada || "",
        dataReplanejadaAtual: realizado.dataReplanejadaAtual || "",

        perda: realizado.perda ?? false,
        motivoPerda: realizado.motivoPerda || "",
        justificativaPerda: realizado.justificativaPerda || "",
        comentario: realizado.comentario || "",
        usuarioResponsavel: realizado.usuarioResponsavel || "",
        quantidadeReplanejamentos:
          Number(realizado.quantidadeReplanejamentos || 0) || 0,
      };

      novoRealizado.statusOperacional = primaryStatusOf(novoRealizado);
      novoRealizado.substatusOperacional =
        substatusListOf(novoRealizado).join(" | ");

      realizadosSomenteNaBase.push(normalizeDemandRecord(novoRealizado));
    });

    return [...carteiraAtualizada, ...realizadosSomenteNaBase];
  }

  async function loadBaseSourcesFromJson() {
    const [baseOrdensRaw, baseFuturasRaw, baseRealizadosRaw] =
      await Promise.all([
        fetchJsonArray("./base/base_ordens.json", "base_ordens.json"),

        fetchJsonArray("./base/base_futuras.json", "base_futuras.json").catch(
          (error) => {
            console.warn("base_futuras.json não carregada:", error);
            return [];
          },
        ),

        fetchJsonArray(
          "./base/base_realizados.json",
          "base_realizados.json",
        ).catch((error) => {
          console.warn("base_realizados.json não carregada:", error);
          return [];
        }),
      ]);

    const baseOrdens = baseOrdensRaw.map((item) =>
      mapBaseItemToDemand(item, "SAP BO - Ordens"),
    );

    const baseFuturas = baseFuturasRaw.map((item) =>
      mapBaseItemToDemand(item, "SAP BO - Demandas Futuras"),
    );

    const baseRealizados = baseRealizadosRaw.map((item) =>
      mapBaseItemToDemand(item, "SAP BO - Realizados"),
    );

    return { baseOrdens, baseFuturas, baseRealizados };
  }

  async function loadBaseFromJson() {
    const { baseOrdens, baseFuturas, baseRealizados } =
      await loadBaseSourcesFromJson();

    const carteiraBase = mergeBaseOrdensEFuturas(baseOrdens, baseFuturas);

    return applyBaseRealizadosToCarteira(carteiraBase, baseRealizados);
  }
  function mergeDemandWithSupabase(baseDemand, delta) {
    if (!delta) return baseDemand;

    return {
      ...baseDemand,
      ...delta,

      id: baseDemand.id || delta.id,
      ordem: baseDemand.ordem || delta.ordem,

      descricao: baseDemand.descricao || delta.descricao,
      gerencia: baseDemand.gerencia || delta.gerencia,
      supervisao: baseDemand.supervisao || delta.supervisao,
      centroTrabalho: baseDemand.centroTrabalho || delta.centroTrabalho,
      localInstalacao: baseDemand.localInstalacao || delta.localInstalacao,
      competencia: normalizeCompetencia(
        baseDemand.competencia || delta.competencia,
      ),
      tipoOM: baseDemand.tipoOM || delta.tipoOM,
      prioridade: normalizePrioridade(
        baseDemand.prioridade || delta.prioridade,
      ),
      critico: delta.critico || baseDemand.critico || "Não informado",
      vencimento: baseDemand.vencimento || delta.vencimento,
      statusSistema: baseDemand.statusSistema || delta.statusSistema || "",
      statusUsuario: baseDemand.statusUsuario || delta.statusUsuario || "",

      dataPlanejada: delta.dataPlanejada || baseDemand.dataPlanejada || "",
      dataReplanejadaAtual:
        delta.dataReplanejadaAtual || baseDemand.dataReplanejadaAtual || "",

      // Data realizada pode vir do Supabase manualmente,
      // mas se a base_realizados existir, ela será aplicada depois
      // em applyBaseRealizadosToCarteira().
      dataRealizada: baseDemand.dataRealizada || delta.dataRealizada || "",

      perda: delta.perda ?? baseDemand.perda ?? false,
      motivoPerda: delta.motivoPerda || baseDemand.motivoPerda || "",
      justificativaPerda:
        delta.justificativaPerda || baseDemand.justificativaPerda || "",
      comentario: delta.comentario || baseDemand.comentario || "",
      usuarioResponsavel:
        delta.usuarioResponsavel || baseDemand.usuarioResponsavel || "",

      quantidadeReplanejamentos:
        delta.quantidadeReplanejamentos ??
        baseDemand.quantidadeReplanejamentos ??
        0,

      origem: delta.origem || baseDemand.origem || "SAP BO",
      dataUltimaAtualizacao:
        delta.dataUltimaAtualizacao || baseDemand.dataUltimaAtualizacao || "",
    };
  }
  function normalizeCentroTrabalho(value) {
    return String(value || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toUpperCase()
      .trim();
  }

  function normalizeScopeValue(value) {
    return normalizeCentroTrabalho(value).replace(/\s+/g, " ");
  }

  function centroResponsabilidadeNivel(cadastro) {
    const nivel = normalizeText(cadastro?.nivelResponsabilidade || "");
    if (nivel.includes("GERENCIA")) return "gerencia";
    if (nivel.includes("SUPERVISAO")) return "supervisao";
    if (nivel.includes("CENTRO")) return "centro";
    if (cadastro?.centroTrabalho || cadastro?.centroTrabalhoChave)
      return "centro";
    if (cadastro?.supervisao) return "supervisao";
    return "gerencia";
  }

  function centroResponsabilidadeChave(cadastro) {
    const chaveInformada = normalizeScopeValue(cadastro?.centroTrabalhoChave);
    if (/^(CENTRO|SUPERVISAO|GERENCIA)::/.test(chaveInformada)) {
      return chaveInformada;
    }
    const nivel = centroResponsabilidadeNivel(cadastro);
    const gerencia = normalizeScopeValue(cadastro?.gerencia);
    const supervisao = normalizeScopeValue(cadastro?.supervisao);
    const centro = normalizeScopeValue(
      chaveInformada || cadastro?.centroTrabalho,
    );

    if (nivel === "centro") return `CENTRO::${centro}`;
    if (nivel === "supervisao") return `SUPERVISAO::${gerencia}::${supervisao}`;
    return `GERENCIA::${gerencia}`;
  }

  function findResponsabilidadeForDemand(demanda, cadastros) {
    const gerencia = normalizeScopeValue(demanda.gerencia);
    const supervisao = normalizeScopeValue(demanda.supervisao);
    const centro = normalizeScopeValue(demanda.centroTrabalho);
    const candidates = [
      `CENTRO::${centro}`,
      `SUPERVISAO::${gerencia}::${supervisao}`,
      `GERENCIA::${gerencia}`,
    ];
    return candidates.map((key) => cadastros.get(key)).find(Boolean);
  }

  function enrichDemandWithCentroTrabalho(demanda, mapaCentros) {
    const chaveCentro = normalizeCentroTrabalho(demanda.centroTrabalho);
    const cadastro = findResponsabilidadeForDemand(demanda, mapaCentros);

    if (!cadastro) {
      return {
        ...demanda,
        centroTrabalhoChave: chaveCentro,
        centroTrabalhoCadastrado: !chaveCentro ? null : false,
        centroTrabalhoStatus: !chaveCentro ? "Sem centro" : "Nao cadastrado",
      };
    }

    return {
      ...demanda,

      // Se a base SAP já vier com gerência/supervisão, você pode escolher manter ou sobrescrever.
      // Aqui estou sobrescrevendo pelo Supabase, porque ele vira a fonte oficial.
      gerencia: cadastro.gerencia || demanda.gerencia || "",
      supervisao: cadastro.supervisao || demanda.supervisao || "",

      planejadorCurto: cadastro.planejadorCurto || "",
      planejadorCurtoEmail: cadastro.planejadorCurtoEmail || "",
      planejadorCurtoMatricula: cadastro.planejadorCurtoMatricula || "",

      planejadorOM: cadastro.planejadorOM || "",
      planejadorOMEmail: cadastro.planejadorOMEmail || "",
      planejadorOMMatricula: cadastro.planejadorOMMatricula || "",

      programador: cadastro.programador || "",
      programadorEmail: cadastro.programadorEmail || "",
      programadorMatricula: cadastro.programadorMatricula || "",

      centroTrabalhoChave: cadastro.centroTrabalhoChave || chaveCentro,
      centroTrabalhoCadastrado: true,
      centroTrabalhoStatus: "Cadastrado",
      centroResponsabilidadeNivel: centroResponsabilidadeNivel(cadastro),
    };
  }

  async function loadDatabase() {
    const previousSelection = state.selectedDemandId;
    const previousUserEmail =
      state.currentUser?.email || getStoredSessionEmail();

    const baseSources = await loadBaseSourcesFromJson();

    const base = mergeBaseOrdensEFuturas(
      baseSources.baseOrdens,
      baseSources.baseFuturas,
    );

    const supabaseData = await state.repo.getAll();

    const mapaCentrosTrabalho = new Map(
      (supabaseData.centrosTrabalho || [])
        .filter((item) => item.ativo !== false)
        .map((item) => [centroResponsabilidadeChave(item), item]),
    );

    const baseEnriquecida = base.map((demanda) =>
      enrichDemandWithCentroTrabalho(demanda, mapaCentrosTrabalho),
    );

    const deltasById = new Map(
      (supabaseData.demandas || []).map((item) => [item.id, item]),
    );

    const baseIds = new Set(baseEnriquecida.map((item) => item.id));

    const qualitySourceRecords = buildQualitySourceRecords({
      baseOrdens: baseSources.baseOrdens,
      baseFuturas: baseSources.baseFuturas,
      supabaseDemandas: supabaseData.demandas || [],
      baseIds,
      deltasById,
      mapaCentrosTrabalho,
    });

    const mergedBase = baseEnriquecida.map((item) =>
      enrichDemandWithCentroTrabalho(
        mergeDemandWithSupabase(item, deltasById.get(item.id)),
        mapaCentrosTrabalho,
      ),
    );

    const demandasSomenteSupabase = (supabaseData.demandas || [])
      .filter((item) => !baseIds.has(item.id))
      .map((item) => enrichDemandWithCentroTrabalho(item, mapaCentrosTrabalho));

    const demandasAntesRealizados = [
      ...demandasSomenteSupabase,
      ...mergedBase,
    ].map(normalizeDemandRecord);

    const demandasComRealizados = applyBaseRealizadosToCarteira(
      demandasAntesRealizados,
      baseSources.baseRealizados,
    ).map(normalizeDemandRecord);

    const demandas = consolidateCarteiraByRealizedOrder(demandasComRealizados)
      .map((demanda) =>
        enrichDemandWithCentroTrabalho(demanda, mapaCentrosTrabalho),
      )
      .map(normalizeDemandRecord);

    state.db = {
      demandas,
      qualitySourceRecords,
      qualityBaseRealizados: baseSources.baseRealizados || [],
      usuarios: supabaseData.usuarios || [],
      centrosTrabalho: supabaseData.centrosTrabalho || [],
      feriasSubstituicoes: supabaseData.feriasSubstituicoes || [],
      configuracoes: supabaseData.configuracoes || {},
      parametros: supabaseData.parametros || {},
      parametrosDisponiveis: supabaseData.parametrosDisponiveis === true,
      historicoPlanejamento: supabaseData.historicoPlanejamento || [],
      historicoReplanejamento: supabaseData.historicoReplanejamento || [],
      historicoRealizadoPerdas: supabaseData.historicoRealizadoPerdas || [],
      logs: supabaseData.logs || [],
    };

    state.quality.issuesCache = [];
    state.quality.filteredCache = [];
    state.quality.selectedIssueId = "";
    state.quality.selectedPrimarySequence = "";
    state.quality.page = 1;
    state.quality.localGroupsCache = null;
    state.quality.localFiltersReady = false;
    state.notifications.cache = [];
    state.notifications.cacheKey = "";
    state.notifications.byId = new Map();
    state.filterOptionsCache = null;
    state.indicatorFiltersReady = false;

    clearBatchLookup();

    setCurrentUserFromEmail(previousUserEmail);

    state.lastDataUpdateAt = latestDataUpdateAt();

    state.selectedDemandId = demandas.some(
      (item) => item.id === previousSelection,
    )
      ? previousSelection
      : demandas[0]?.id || "";
  }

  function normalizeDemandRecord(demanda) {
    return {
      ...demanda,
      competencia: normalizeCompetencia(demanda.competencia),
      prioridade: normalizePrioridade(demanda.prioridade),
      critico: normalizeCritico(demanda.critico),
      planejadorCurto: demanda.planejadorCurto || "",
      planejadorCurtoEmail: demanda.planejadorCurtoEmail || "",
      planejadorCurtoMatricula: demanda.planejadorCurtoMatricula || "",
      centroTrabalhoChave:
        demanda.centroTrabalhoChave ||
        normalizeCentroTrabalho(demanda.centroTrabalho),
    };
  }

  function decorateQualitySource(record, fonteQualidade, sequencia) {
    return {
      ...record,
      fonteQualidade: fonteQualidade || record.origem || "Sistema",
      qualidadeSequencia: sequencia,
      qualidadeIdInformado:
        record.qualidadeIdInformado || record.idDemandaInformado || "",
    };
  }

  function buildQualitySourceRecords({
    baseOrdens,
    baseFuturas,
    supabaseDemandas,
    baseIds,
    deltasById,
    mapaCentrosTrabalho,
  }) {
    const baseRecords = [
      ...baseOrdens.map((item, index) =>
        decorateQualitySource(item, "SAP BO - Ordens", `ordens-${index + 1}`),
      ),
      ...baseFuturas.map((item, index) =>
        decorateQualitySource(
          item,
          "SAP BO - Demandas Futuras",
          `futuras-${index + 1}`,
        ),
      ),
    ];

    const supabaseOnlyRecords = (supabaseDemandas || [])
      .filter((item) => !baseIds.has(item.id))
      .map((item, index) =>
        decorateQualitySource(
          {
            ...item,
            qualidadeIdInformado: item.id || "",
          },
          item.origem || "Supabase",
          `supabase-${index + 1}`,
        ),
      );

    return [...baseRecords, ...supabaseOnlyRecords].map((item) => {
      const delta =
        normalizeText(item.fonteQualidade).includes("SUPABASE") ||
        !deltasById.has(item.id)
          ? null
          : deltasById.get(item.id);
      const merged = mergeDemandWithSupabase(item, delta);
      return normalizeDemandRecord(
        enrichDemandWithCentroTrabalho(
          {
            ...merged,
            fonteQualidade: item.fonteQualidade,
            qualidadeSequencia: item.qualidadeSequencia,
            qualidadeIdInformado:
              item.qualidadeIdInformado || item.idDemandaInformado || "",
          },
          mapaCentrosTrabalho,
        ),
      );
    });
  }

  function clearSensitiveUrlParams() {
    const url = new URL(global.location.href);

    const sensitiveParams = ["email", "matricula", "user"];

    let changed = false;

    sensitiveParams.forEach((param) => {
      if (url.searchParams.has(param)) {
        url.searchParams.delete(param);
        changed = true;
      }
    });

    if (!changed) return;

    const cleanUrl =
      url.pathname +
      (url.searchParams.toString() ? `?${url.searchParams.toString()}` : "") +
      url.hash;

    global.history.replaceState({}, document.title, cleanUrl);
  }

  function getStoredSessionEmail() {
    try {
      const session = JSON.parse(
        global.localStorage.getItem("cce.session") || "null",
      );
      return session?.email || "";
    } catch (error) {
      return "";
    }
  }

  function setCurrentUserFromEmail(email) {
    if (!email) {
      state.currentUser = null;
      return;
    }
    state.currentUser =
      state.db.usuarios.find(
        (user) =>
          user.ativo && normalizeText(user.email) === normalizeText(email),
      ) || null;
  }

  function latestDataUpdateAt() {
    const dates = [
      ...state.db.demandas.map((item) => item.dataUltimaAtualizacao),
      ...state.db.logs.map((item) => item.dataHora),
    ]
      .map((value) => (value ? new Date(value) : null))
      .filter((date) => date && !Number.isNaN(date.getTime()))
      .sort((a, b) => b - a);
    return dates[0]?.toISOString() || "";
  }
  function demandById(id) {
    return state.db.demandas.find((item) => item.id === id);
  }

  function historiesFor(id) {
    const planejamento = state.db.historicoPlanejamento
      .filter((item) => item.demandaId === id)
      .map((item) => ({
        type: "Planejamento",
        date: item.dataHora,
        title: `Planejado para ${formatDate(item.novaData)}`,
        detail: item.comentario || item.usuario,
      }));
    const replanejamento = state.db.historicoReplanejamento
      .filter((item) => item.demandaId === id)
      .map((item) => ({
        type: "Replanejamento",
        date: item.dataHora,
        title: `${formatDate(item.dataAnterior)} para ${formatDate(item.novaData)}`,
        detail: `${item.motivo || "-"} | ${item.justificativa || "-"}`,
      }));
    const realizados = state.db.historicoRealizadoPerdas
      .filter((item) => item.demandaId === id)
      .map((item) => ({
        type: item.perda ? "Perda" : "Realizado",
        date: item.dataHora,
        title: item.perda
          ? item.motivoPerda || "Perda registrada"
          : `Realizado em ${formatDate(item.dataRealizada)}`,
        detail: item.comentario || item.justificativaPerda || item.usuario,
      }));
    return [...planejamento, ...replanejamento, ...realizados].sort(
      (a, b) => new Date(b.date) - new Date(a.date),
    );
  }

  function buildPortfolioHistoryRows() {
    const demandMap = new Map(
      (state.db?.demandas || []).map((demanda) => [demanda.id, demanda]),
    );

    const rowFor = (item, action, title, detail) => {
      const demand = demandMap.get(item.demandaId) || {};
      return {
        id: item.id || `${action}-${item.demandaId}-${item.dataHora}`,
        demandId: item.demandaId || "",
        ordem: demand.ordem || item.ordem || "",
        descricao: demand.descricao || item.descricao || "",
        usuario: item.usuario || demand.usuarioResponsavel || "",
        action,
        date: item.dataHora || item.data || "",
        title,
        detail,
      };
    };

    const planejamento = (state.db?.historicoPlanejamento || []).map((item) =>
      rowFor(
        item,
        "Planejamento",
        `Planejado para ${formatDate(item.novaData)}`,
        item.comentario || "Planejamento registrado",
      ),
    );

    const replanejamento = (state.db?.historicoReplanejamento || []).map(
      (item) =>
        rowFor(
          item,
          "Replanejamento",
          `${formatDate(item.dataAnterior)} para ${formatDate(item.novaData)}`,
          [item.motivo, item.justificativa].filter(Boolean).join(" | ") ||
            "Replanejamento registrado",
        ),
    );

    const realizados = (state.db?.historicoRealizadoPerdas || []).map((item) =>
      rowFor(
        item,
        item.perda ? "Perda" : "Realizado",
        item.perda
          ? item.motivoPerda || "Perda registrada"
          : `Realizado em ${formatDate(item.dataRealizada)}`,
        item.comentario || item.justificativaPerda || "Registro operacional",
      ),
    );

    return [...planejamento, ...replanejamento, ...realizados].sort(
      (a, b) => new Date(b.date || 0) - new Date(a.date || 0),
    );
  }

  function filteredPortfolioHistoryRows() {
    const filters = state.portfolioHistory;
    const search = normalizeText(filters.search);
    const user = normalizeText(filters.user);

    return buildPortfolioHistoryRows().filter((row) => {
      const rowDate = toDate(row.date)?.toISOString().slice(0, 10) || "";

      if (filters.action && row.action !== filters.action) return false;
      if (filters.startDate && rowDate && rowDate < filters.startDate)
        return false;
      if (filters.endDate && rowDate && rowDate > filters.endDate)
        return false;
      if (user && !normalizeText(row.usuario).includes(user)) return false;

      if (search) {
        const haystack = normalizeText(
          [
            row.demandId,
            row.ordem,
            row.descricao,
            row.usuario,
            row.action,
            row.title,
            row.detail,
          ].join(" "),
        );
        if (!haystack.includes(search)) return false;
      }

      return true;
    });
  }

  function renderPortfolioHistory() {
    const filters = state.portfolioHistory;
    const rows = filteredPortfolioHistoryRows();
    const visibleRows = rows.slice(0, 500);

    $("#portfolioHistorySearch").value = filters.search;
    $("#portfolioHistoryAction").value = filters.action;
    $("#portfolioHistoryUser").value = filters.user;
    $("#portfolioHistoryStartDate").value = filters.startDate;
    $("#portfolioHistoryEndDate").value = filters.endDate;

    $("#portfolioHistoryCount").textContent =
      rows.length > visibleRows.length
        ? `${rows.length} registros encontrados - exibindo os primeiros ${visibleRows.length}`
        : `${rows.length} registros encontrados`;

    $("#portfolioHistoryTableBody").innerHTML = visibleRows.length
      ? visibleRows
          .map(
            (row) => `
              <tr>
                <td>${formatDateTime(row.date)}</td>
                <td>${escapeHtml(row.action)}</td>
                <td>
                  <strong>${escapeHtml(row.ordem || row.demandId || "-")}</strong>
                  <div class="muted">${escapeHtml(row.demandId || "-")}</div>
                </td>
                <td class="description-cell">
                  ${escapeHtml(row.descricao || "-")}
                  <div class="muted">${escapeHtml(row.title || "-")}</div>
                </td>
                <td>${escapeHtml(row.usuario || "-")}</td>
                <td>${escapeHtml(row.detail || "-")}</td>
                <td>
                  <button
                    class="button secondary"
                    type="button"
                    data-history-demand="${escapeHtml(row.demandId)}"
                    ${row.demandId ? "" : "disabled"}
                  >
                    Abrir
                  </button>
                </td>
              </tr>
            `,
          )
          .join("")
      : `<tr>
          <td colspan="7">
            <div class="empty-detail">
              <strong>Nenhum histórico encontrado</strong>
              <span>Ajuste os filtros para consultar alterações da carteira.</span>
            </div>
          </td>
        </tr>`;
  }

  function hydrateStaticUi() {
    $("#storageMode").textContent = state.repo.mode;
    $("#pageSize").value = String(state.pageSize);
    const userSelect = $("#userSelect");
    userSelect.innerHTML = state.db.usuarios
      .filter((user) => user.ativo)
      .map(
        (user) =>
          `<option value="${escapeHtml(user.email)}">${escapeHtml(user.nome)}</option>`,
      )
      .join("");
    if (state.currentUser) userSelect.value = state.currentUser.email;
    $(".dev-user-select").classList.toggle(
      "hidden",
      !(
        state.repo.mode === "Local" &&
        new URLSearchParams(global.location.search).get("debugUsers") === "1"
      ),
    );
    $("#lastUpdateSide").textContent = state.lastDataUpdateAt
      ? formatDateTime(state.lastDataUpdateAt)
      : "-";
    renderRole();
    collectFilters();
    buildFilterOptions();
    state.indicatorFiltersReady = false;
    renderAlerts();
  }

  function renderRole() {
    $("#userName").textContent =
      state.currentUser?.nome || state.identity?.nome || "Aguardando login";
    $("#roleChip").textContent = state.currentUser?.perfil || "Visualizador";
    $("#logoutButton")?.classList.toggle("hidden", !state.currentUser);
    applyPermissions();
  }
  function setLoginUiState({
    loading = false,
    ready = false,
    buttonText = "",
    statusText = "",
    errorText = "",
  } = {}) {
    const button = $("#loginSubmit");
    const status = $("#loginStatus");
    const error = $("#loginError");

    if (button) {
      button.disabled = loading || !ready;
      button.textContent =
        buttonText || (ready ? "Entrar na Central" : "Carregando acesso...");
      button.classList.toggle("is-loading", loading);
    }

    if (status) {
      status.textContent =
        statusText ||
        (ready
          ? "Acesso pronto. Informe e-mail e matrícula."
          : "Preparando tela de login...");
    }

    if (error) {
      error.textContent = errorText || "";
    }
  }

  async function loadLoginData() {
    const loginData = state.repo.getLoginData
      ? await state.repo.getLoginData()
      : await state.repo.getAll();

    state.db = loginData;

    state.loginReady = true;

    setLoginUiState({
      ready: true,
      loading: false,
      buttonText: "Entrar na Central",
      statusText: "Acesso pronto. Informe e-mail e matrícula.",
    });
  }

  function renderLoginState() {
    const loginScreen = $("#loginScreen");
    if (!loginScreen) return;

    const isLogged = Boolean(state.currentUser);

    loginScreen.classList.toggle("hidden", isLogged);
    document.body.classList.toggle("is-login-required", !isLogged);

    if (isLogged) return;

    if (!state.loginReady) {
      setLoginUiState({
        ready: false,
        loading: false,
        buttonText: "Carregando acesso...",
        statusText: "Preparando tela de login...",
      });
      return;
    }

    setLoginUiState({
      ready: true,
      loading: false,
      buttonText: "Entrar na Central",
      statusText: "Acesso pronto. Informe e-mail e matrícula.",
    });
  }

  async function handleLogin(event) {
    event.preventDefault();

    if (state.loginSubmitting) return;

    const formElement = event.currentTarget;
    const errorElement = $("#loginError");

    if (!state.loginReady || !state.db?.usuarios?.length) {
      setLoginUiState({
        ready: false,
        loading: false,
        buttonText: "Carregando acesso...",
        statusText: "Aguarde o carregamento da base de usuários.",
        errorText:
          "O sistema ainda está preparando o acesso. Tente novamente em alguns segundos.",
      });
      return;
    }

    state.loginSubmitting = true;

    setLoginUiState({
      ready: true,
      loading: true,
      buttonText: "Validando...",
      statusText: "Conferindo usuário e matrícula...",
    });

    const form = new FormData(formElement);
    const email = String(form.get("email") || "")
      .trim()
      .toLowerCase();
    const matricula = String(form.get("matricula") || "").trim();

    const user = state.db.usuarios.find(
      (item) => normalizeText(item.email) === normalizeText(email),
    );

    try {
      if (!user || user.ativo === false) {
        setLoginUiState({
          ready: true,
          loading: false,
          buttonText: "Entrar na Central",
          statusText: "Acesso pronto. Informe e-mail e matrícula.",
          errorText: "Usuário não encontrado ou inativo.",
        });

        state.loginSubmitting = false;
        return;
      }

      if (String(user.matricula || "").trim() !== matricula) {
        setLoginUiState({
          ready: true,
          loading: false,
          buttonText: "Entrar na Central",
          statusText: "Acesso pronto. Informe e-mail e matrícula.",
          errorText: "Matrícula inválida para o e-mail informado.",
        });

        state.loginSubmitting = false;
        return;
      }

      state.currentUser = user;

      global.localStorage.setItem(
        "cce.session",
        JSON.stringify({
          email: user.email,
          createdAt: new Date().toISOString(),
        }),
      );

      clearSensitiveUrlParams();

      setLoginUiState({
        ready: true,
        loading: true,
        buttonText: "Carregando carteira...",
        statusText: "Acesso validado. Carregando carteira operacional...",
      });

      await loadDatabase();
      await autoSyncRealizadosFromSharePoint();

      state.pageSize = Number(state.db.parametros?.pageSizeDefault || 12);

      formElement.reset();

      hydrateStaticUi();
      renderLoginState();
      renderRole();
      renderCurrentView();

      await state.repo.addLog?.({
        usuario: user.email,
        acao: "Login",
        lista: "usuarios_central_eletrovia",
        referencia: user.email,
        detalhe: "Login realizado com e-mail e matrícula.",
        modulo: "LOGIN",
        status: "SUCESSO",
      });
    } catch (error) {
      console.error(error);

      state.currentUser = null;

      setLoginUiState({
        ready: true,
        loading: false,
        buttonText: "Entrar na Central",
        statusText: "Acesso pronto. Informe e-mail e matrícula.",
        errorText: `Erro ao carregar a carteira: ${error.message}`,
      });
    } finally {
      state.loginSubmitting = false;
    }
  }

  async function logout() {
    const email = state.currentUser?.email || "";
    global.localStorage.removeItem("cce.session");
    state.currentUser = null;
    renderRole();
    renderLoginState();
    if (email) {
      await state.repo.addLog?.({
        usuario: email,
        acao: "Logout",
        lista: "usuarios_central_eletrovia",
        referencia: email,
        detalhe: "Sessao local encerrada.",
        modulo: "LOGIN",
        status: "SUCESSO",
      });
    }
  }

  function applyPermissions() {
    const rules = getRules();
    $$(".admin-only").forEach((element) => {
      element.classList.toggle("hidden", !rules.configurar);
      element.disabled = !rules.configurar;
    });
    $$(".editor-only").forEach((element) => {
      element.disabled = !canEdit();
    });
    $$(".batch-only").forEach((element) => {
      element.disabled = !canBatch();
    });
    $$(".planner-only").forEach((element) => {
      element.disabled = !canPlan();
    });
    $$(".realizer-only").forEach((element) => {
      element.disabled = !canRealizar();
    });
    $$(
      '[data-view="administracao"], [data-view="logs"], [data-view="saude-integracao"]',
    ).forEach((element) => {
      element.disabled = !rules.configurar;
      element.title = rules.configurar ? "" : "Disponivel para Administrador";
    });
    $$('[data-view="lote"]').forEach((element) => {
      element.disabled = !rules.cargaLote;
      element.title = rules.cargaLote ? "" : "Sem permissao para carga em lote";
    });
    $("#exportCsv").disabled = !canExport();
  }

  function filterOptionsFor(definition) {
    if (!state.filterOptionsCache) state.filterOptionsCache = {};
    if (state.filterOptionsCache[definition.key]) {
      return state.filterOptionsCache[definition.key];
    }

    const rows = state.db?.demandas || [];
    const values =
      definition.special === "substatus"
        ? rows.flatMap((item) =>
            String(filterValueFor(item, definition))
              .split(" | ")
              .map((option) => option.trim())
              .filter(Boolean),
          )
        : rows.map((item) => filterValueFor(item, definition));

    state.filterOptionsCache[definition.key] = uniqueOptions(values);
    return state.filterOptionsCache[definition.key];
  }

  function buildFilterOptions({ includeHidden = false } = {}) {
    const filters = state.filters || {};
    FILTER_DEFINITIONS.forEach((definition) => {
      const host = $(`[data-multi-filter="${definition.key}"]`);
      if (!host) return;
      if (
        !includeHidden &&
        host.classList.contains("advanced-filter") &&
        host.classList.contains("hidden")
      ) {
        return;
      }
      const selected = filters[definition.key] || [];
      const options = uniqueOptions([
        ...filterOptionsFor(definition),
        ...selected,
      ]);
      renderMultiFilter(host, definition, options, selected);
    });
  }

  function buildIndicatorFilterOptions() {
    const filters = state.indicatorFilters || {};
    FILTER_DEFINITIONS.forEach((definition) => {
      const host = $(`[data-indicator-multi-filter="${definition.key}"]`);
      if (!host) return;
      const scopedRows = (state.db?.demandas || []).filter((item) =>
        demandMatchesFilters(item, filters, definition.key),
      );
      const rowOptions =
        definition.special === "substatus"
          ? scopedRows.flatMap((item) =>
              String(filterValueFor(item, definition))
                .split(" | ")
                .map((option) => option.trim())
                .filter(Boolean),
            )
          : scopedRows.map((item) => filterValueFor(item, definition));
      const availableOptions = uniqueOptions(rowOptions);
      const selected = (filters[definition.key] || []).filter((option) =>
        availableOptions.includes(option),
      );
      filters[definition.key] = selected;
      const options = uniqueOptions([
        ...availableOptions,
        ...selected,
      ]);
      renderMultiFilter(host, definition, options, selected);
    });
    state.indicatorFilters = filters;
    state.indicatorFiltersReady = true;
  }

  function buildQualityLocalFilterOptions() {
    const filters = state.quality.localFilters || {};
    FILTER_DEFINITIONS.forEach((definition) => {
      const host = $(`[data-quality-local-multi-filter="${definition.key}"]`);
      if (!host) return;

      const scopedRows = (state.db?.demandas || []).filter((item) =>
        demandMatchesFilters(item, filters, definition.key),
      );
      const rowOptions =
        definition.special === "substatus"
          ? scopedRows.flatMap((item) =>
              String(filterValueFor(item, definition))
                .split(" | ")
                .map((option) => option.trim())
                .filter(Boolean),
            )
          : scopedRows.map((item) => filterValueFor(item, definition));
      const availableOptions = uniqueOptions(rowOptions);
      const selected = (filters[definition.key] || []).filter((option) =>
        availableOptions.includes(option),
      );

      filters[definition.key] = selected;
      renderMultiFilter(
        host,
        definition,
        uniqueOptions([...availableOptions, ...selected]),
        selected,
      );
    });

    state.quality.localFilters = filters;
    state.quality.localFiltersReady = true;
  }

  function collectFilters() {
    const filters = {};
    $$("[data-multi-filter]").forEach((field) => {
      filters[field.dataset.multiFilter] = $$(
        "[data-multi-option]:checked",
        field,
      ).map((input) => input.value);
    });
    $$("[data-filter]").forEach((field) => {
      filters[field.dataset.filter] = field.value;
    });
    filters.quickSearch = $("#quickSearch").value.trim();
    state.filters = filters;
    return filters;
  }

  function collectIndicatorFilters() {
    const filters = {};
    $$("[data-indicator-multi-filter]").forEach((field) => {
      filters[field.dataset.indicatorMultiFilter] = $$(
        "[data-multi-option]:checked",
        field,
      ).map((input) => input.value);
    });
    filters.quickSearch = $("#indicatorQuickSearch")?.value.trim() || "";
    state.indicatorFilters = filters;
    return filters;
  }

  function collectQualityLocalFilters() {
    const filters = {};
    $$("[data-quality-local-multi-filter]").forEach((field) => {
      filters[field.dataset.qualityLocalMultiFilter] = $$(
        "[data-multi-option]:checked",
        field,
      ).map((input) => input.value);
    });
    filters.quickSearch = $("#qualityLocalQuickSearch")?.value.trim() || "";
    state.quality.localFilters = filters;
    state.quality.localGroupsCache = null;
    return filters;
  }

  function updateMultiFilterSummary(host) {
    if (!host) return;
    const checked = $$("[data-multi-option]:checked", host).map(
      (input) => input.value,
    );
    const summary =
      checked.length === 0
        ? "Todos"
        : checked.length === 1
          ? checked[0]
          : `${checked.length} selecionados`;
    const summaryText = $(".multi-summary-text", host);
    const summaryElement = $("summary", host);
    if (summaryText) summaryText.textContent = summary;
    if (summaryElement) summaryElement.title = summary;
  }

  function updateAllMultiFilterSummaries(root) {
    $$("[data-multi-filter], [data-indicator-multi-filter]", root || document)
      .forEach(updateMultiFilterSummary);
  }

  function filterValueFor(item, definition) {
    if (!item.__filterCache) {
      Object.defineProperty(item, "__filterCache", {
        value: {},
        enumerable: false,
        configurable: true,
      });
    }
    if (Object.prototype.hasOwnProperty.call(item.__filterCache, definition.key))
      return item.__filterCache[definition.key];

    let result = "";

    if (definition.special === "status") {
      result = primaryStatusOf(item);
      item.__filterCache[definition.key] = result;
      return result;
    }

    if (definition.special === "centroStatus") {
      result = item.centroTrabalhoStatus || "Nao cadastrado";
      item.__filterCache[definition.key] = result;
      return result;
    }

    if (definition.special === "substatus") {
      result = substatusListOf(item).join(" | ") || "Sem substatus";
      item.__filterCache[definition.key] = result;
      return result;
    }

    if (definition.special === "anoVencimento") {
      result = dateText(item.vencimento).slice(0, 4) || "";
      item.__filterCache[definition.key] = result;
      return result;
    }

    if (definition.special === "mesVencimento") {
      const month = dateText(item.vencimento).slice(5, 7) || "";
      result = month ? `${month} - ${monthName(month)}` : "";
      item.__filterCache[definition.key] = result;
      return result;
    }

    const value = item[definition.field] || "";

    if (definition.formatter === "sapStatus") {
      result = formatSapStatusFilter(value);
      item.__filterCache[definition.key] = result;
      return result;
    }

    item.__filterCache[definition.key] = value;
    return value;
  }

  function renderMultiFilter(host, definition, options, selected) {
    const selectedSet = new Set(selected || []);
    const checkedCount = selectedSet.size;

    const summary =
      checkedCount === 0
        ? "Todos"
        : checkedCount === 1
          ? selected[0]
          : `${checkedCount} selecionados`;

    const normalizedOptions = uniqueOptions(options);

    host.innerHTML = `
    <label class="multi-label">${escapeHtml(definition.label)}</label>

    <details class="multi-select filter-select-pro" data-filter-details>
      <summary title="${escapeHtml(summary)}">
        <span class="multi-summary-text">${escapeHtml(summary)}</span>
      </summary>

      <div class="multi-menu filter-menu-pro">
        <div class="multi-search-wrap">
          ${iconSvg("filter")}
          <input
            data-multi-search
            type="search"
            placeholder="Pesquisar em ${escapeHtml(definition.label)}..."
            aria-label="Pesquisar ${escapeHtml(definition.label)}"
          />
        </div>

        <div class="multi-menu-actions">
          <button
            class="mini-filter-button"
            type="button"
            data-multi-select-visible
          >
            Selecionar visíveis
          </button>

          <button
            class="mini-filter-button secondary"
            type="button"
            data-multi-clear
          >
            Limpar seleção
          </button>
        </div>

        <div class="multi-options">
          ${
            normalizedOptions.length
              ? normalizedOptions
                  .map(
                    (option) => `
              <label class="multi-option">
                <input
                  data-multi-option
                  type="checkbox"
                  value="${escapeHtml(option)}"
                  ${selectedSet.has(option) ? "checked" : ""}
                />
                <span>${escapeHtml(option)}</span>
              </label>
            `,
                  )
                  .join("")
              : '<span class="muted multi-empty">Sem opções no recorte</span>'
          }
        </div>
      </div>
    </details>
  `;
  }
  function demandMatchesFilters(item, filters, ignoredKey = "") {
    const search = normalizeText(filters.quickSearch);

    for (const definition of FILTER_DEFINITIONS) {
      if (definition.key === ignoredKey) continue;
      const selected = filters[definition.key] || [];
      if (selected.length) {
        const value = String(filterValueFor(item, definition));
        if (definition.special === "substatus") {
          const substatuses = value
            .split(" | ")
            .map((option) => option.trim())
            .filter(Boolean);
          if (!selected.some((option) => substatuses.includes(option)))
            return false;
        } else if (!selected.includes(value)) {
          return false;
        }
      }
    }

    if (filters.perda === "sim" && !item.perda) return false;
    if (filters.perda === "nao" && item.perda) return false;
    if (
      filters.planejado === "sim" &&
      !item.dataPlanejada &&
      !item.dataReplanejadaAtual
    )
      return false;
    if (
      filters.planejado === "nao" &&
      (item.dataPlanejada || item.dataReplanejadaAtual)
    )
      return false;
    if (filters.realizado === "sim" && !item.dataRealizada) return false;
    if (filters.realizado === "nao" && item.dataRealizada) return false;

    if (search) {
      if (!item.__filterSearchText) {
        Object.defineProperty(item, "__filterSearchText", {
          value: normalizeText(
            [
              item.id,
              item.ordem,
              item.descricao,
              item.centroTrabalho,
              item.localInstalacao,
              item.gerencia,
              item.supervisao,
              item.usuarioResponsavel,
              item.planejadorCurto,
              item.planejadorOM,
              item.programador,
            ].join(" "),
          ),
          enumerable: false,
          configurable: true,
        });
      }
      const haystack = item.__filterSearchText;
      if (!haystack.includes(search)) return false;
    }

    return true;
  }

  function filteredDemandas() {
    const filters = collectFilters();
    return state.db.demandas.filter((item) =>
      demandMatchesFilters(item, filters),
    );
  }

  function dashboardStats(demands) {
    const total = demands.length;
    const counts = STATUS_OPTIONS.reduce((acc, status) => {
      acc[status] = 0;
      return acc;
    }, {});
    demands.forEach((item) => {
      const status = primaryStatusOf(item);
      counts[status] = (counts[status] || 0) + 1;
    });
    const realizedOnTime = demands.filter(
      (item) =>
        primaryStatusOf(item) === "Realizado" &&
        dueClassOf(item) === "No Prazo",
    ).length;
    const planned = demands.filter(
      (item) => item.dataPlanejada || item.dataReplanejadaAtual,
    ).length;
    const adherence = planned
      ? Math.round((realizedOnTime / planned) * 100)
      : 0;
    return {
      total,
      aPlanejar: counts["A Planejar"] || 0,
      planejadas: counts.Planejado || 0,
      replanejadas: counts.Replanejado || 0,
      realizadas: counts.Realizado || 0,
      perdas: demands.filter((item) => item.perda).length,
      aderencia: adherence,
      pendentesPerda: demands.filter((item) => pendingIssuesOf(item).length)
        .length,
    };
  }

  function qualityTypeLabel(typeKey) {
    return (
      QUALITY_TYPE_DEFINITIONS.find((item) => item.key === typeKey)?.label ||
      typeKey
    );
  }

  function qualitySourceLabel(record) {
    return record.fonteQualidade || record.origem || "Sistema";
  }

  function qualityExplicitId(record) {
    const informed = String(
      record.qualidadeIdInformado || record.idDemandaInformado || "",
    ).trim();
    if (informed) return informed;
    if (normalizeText(qualitySourceLabel(record)).includes("SUPABASE")) {
      return String(record.id || "").trim();
    }
    return "";
  }

  function addQualityGroup(map, key, record) {
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(record);
  }

  function joinedDistinct(records, selector, fallback = "-") {
    const values = uniqueOptions(
      records.map(selector).map((value) => String(value || "").trim()),
    );
    return values.length ? values.join(" | ") : fallback;
  }

  function firstFilled(records, selector) {
    return records.map(selector).find((value) => String(value || "").trim());
  }

  function makeQualityIssue(typeKey, key, records) {
    const issue = {
      typeKey,
      typeLabel: qualityTypeLabel(typeKey),
      key,
      quantidade: records.length,
      records,
      origem: joinedDistinct(records, qualitySourceLabel),
      descricao:
        firstFilled(records, (item) => item.descricao) || "Nao informado",
      centroTrabalho:
        firstFilled(records, (item) => item.centroTrabalho) || "-",
      gerencia: firstFilled(records, (item) => item.gerencia) || "-",
      supervisao: firstFilled(records, (item) => item.supervisao) || "-",
      competencia: firstFilled(records, (item) => item.competencia) || "-",
    };

    issue.searchText = normalizeText(
      [
        issue.typeLabel,
        issue.key,
        issue.origem,
        issue.descricao,
        issue.centroTrabalho,
        issue.gerencia,
        issue.supervisao,
        issue.competencia,
        ...records.flatMap((record) => [
          record.id,
          record.ordem,
          qualityExplicitId(record),
          record.descricao,
          record.localInstalacao,
        ]),
      ].join(" "),
    );

    return issue;
  }

  function buildQualityIssues() {
    const records = state.db?.qualitySourceRecords || state.db?.demandas || [];
    const omGroups = new Map();
    const idGroups = new Map();
    const missingCenterGroups = new Map();
    const noKeyRecords = [];

    records.forEach((record) => {
      const ordem = String(record.ordem || "").trim();
      const id = qualityExplicitId(record);

      if (ordem) {
        addQualityGroup(omGroups, ordem, record);
      } else if (id) {
        addQualityGroup(idGroups, id, record);
      } else {
        noKeyRecords.push(record);
      }

      if (
        record.centroTrabalhoChave &&
        record.centroTrabalhoCadastrado === false
      ) {
        addQualityGroup(
          missingCenterGroups,
          record.centroTrabalhoChave || record.centroTrabalho,
          record,
        );
      }
    });

    const issues = [];

    omGroups.forEach((group, key) => {
      if (group.length > 1)
        issues.push(makeQualityIssue("om-duplicada", key, group));
    });

    idGroups.forEach((group, key) => {
      if (group.length > 1)
        issues.push(makeQualityIssue("id-duplicado-sem-om", key, group));
    });

    noKeyRecords.forEach((record, index) => {
      issues.push(
        makeQualityIssue("sem-chave", `Sem chave ${index + 1}`, [record]),
      );
    });

    missingCenterGroups.forEach((group, key) => {
      issues.push(makeQualityIssue("centro-sem-cadastro", key, group));
    });

    const order = QUALITY_TYPE_DEFINITIONS.reduce((acc, item, index) => {
      acc[item.key] = index;
      return acc;
    }, {});

    return issues
      .sort((a, b) => {
        const typeOrder = order[a.typeKey] - order[b.typeKey];
        if (typeOrder) return typeOrder;
        if (b.quantidade !== a.quantidade) return b.quantidade - a.quantidade;
        return String(a.key).localeCompare(String(b.key), "pt-BR");
      })
      .map((issue, index) => ({ ...issue, id: `${issue.typeKey}-${index}` }));
  }

  function filteredQualityIssues(issues) {
    const typeFilter = state.quality.typeFilter;
    const search = normalizeText(state.quality.search);

    return issues.filter((issue) => {
      if (typeFilter && issue.typeKey !== typeFilter) return false;
      if (!search) return true;
      return (issue.searchText || "").includes(search);
    });
  }

  function qualityIssueStats(issues, typeKey) {
    const scoped = issues.filter((issue) => issue.typeKey === typeKey);
    return {
      groups: scoped.length,
      records: scoped.reduce((sum, issue) => sum + issue.quantidade, 0),
    };
  }

  function renderQualityCards(issues) {
    $("#qualityCards").innerHTML = QUALITY_TYPE_DEFINITIONS.map(
      (definition) => {
        const stats = qualityIssueStats(issues, definition.key);
        const note =
          definition.key === "sem-chave"
            ? "sem OM e sem ID"
            : definition.key === "centro-sem-cadastro"
              ? `${stats.records} demandas`
              : `${stats.records} registros envolvidos`;
        const active = state.quality.typeFilter === definition.key;
        return `
        <button class="quality-card ${active ? "is-active" : ""}" data-quality-card="${definition.key}" type="button">
          <span>${escapeHtml(definition.cardLabel)}</span>
          <strong>${stats.groups}</strong>
          <small>${escapeHtml(note)}</small>
        </button>
      `;
      },
    ).join("");
  }

  function recordCompletenessScore(record) {
    return [
      record.ordem,
      qualityExplicitId(record),
      record.descricao,
      record.centroTrabalho,
      record.localInstalacao,
      record.gerencia,
      record.supervisao,
      record.competencia,
      record.vencimento,
      record.tipoOM,
      record.prioridade,
      record.dataPlanejada,
      record.dataReplanejadaAtual,
      record.dataRealizada,
      record.comentario,
      record.observacao,
      record.usuarioResponsavel,
    ].filter((value) => String(value || "").trim()).length;
  }

  function recommendedQualityPrimary(records) {
    const fromOrders = records.find((record) =>
      normalizeText(qualitySourceLabel(record)).includes("ORDENS"),
    );
    if (fromOrders) {
      return {
        record: fromOrders,
        reason: "base_ordens.json com OM existente",
      };
    }

    const [mostComplete] = [...records].sort(
      (a, b) => recordCompletenessScore(b) - recordCompletenessScore(a),
    );
    return mostComplete
      ? { record: mostComplete, reason: "registro com mais campos preenchidos" }
      : null;
  }

  function selectedQualityIssue() {
    return (state.quality.filteredCache || getQualityIssuesCached()).find(
      (issue) => issue.id === state.quality.selectedIssueId,
    );
  }

  function updateQualityRowSelection() {
    $$("#qualityIssueTableBody [data-quality-issue-id]").forEach((row) => {
      row.classList.toggle(
        "is-selected",
        row.dataset.qualityIssueId === state.quality.selectedIssueId,
      );
    });
  }

  function focusQualityDetailPanel() {
    const panel = $("#qualityDetailPanel");
    if (!panel) return;
    panel.classList.add("is-focused");
    panel.scrollIntoView({ behavior: "smooth", block: "nearest" });
    global.setTimeout(() => panel.classList.remove("is-focused"), 900);
  }

  function selectQualityIssue(issueId, scrollToDetail = false) {
    if (!issueId || issueId === state.quality.selectedIssueId) {
      if (scrollToDetail) focusQualityDetailPanel();
      return;
    }

    state.quality.selectedIssueId = issueId;
    state.quality.selectedPrimarySequence = "";
    updateQualityRowSelection();
    renderQualityDetail(selectedQualityIssue());

    if (scrollToDetail) focusQualityDetailPanel();
  }

  function mergeQualityRecords(issue, primarySequence) {
    if (!issue?.records?.length) return null;

    const primary =
      issue.records.find(
        (record) => record.qualidadeSequencia === primarySequence,
      ) ||
      recommendedQualityPrimary(issue.records)?.record ||
      issue.records[0];

    const merged = { ...primary };

    const fieldsToMerge = [
      "id",
      "ordem",
      "tipoDemanda",
      "tipoOM",
      "descricao",
      "gerencia",
      "supervisao",
      "centroTrabalho",
      "localInstalacao",
      "statusSistema",
      "statusUsuario",
      "competencia",
      "vencimento",
      "prioridade",
      "toleranciaMin",
      "toleranciaMax",
      "dataPlanejada",
      "dataReplanejadaAtual",
      "dataRealizada",
      "perda",
      "motivoPerda",
      "justificativaPerda",
      "comentario",
      "usuarioResponsavel",
      "quantidadeReplanejamentos",
      "frequencia",
      "observacao",
      "vinculadaEm",
    ];

    issue.records.forEach((record) => {
      fieldsToMerge.forEach((field) => {
        const current = merged[field];
        const candidate = record[field];

        const currentEmpty =
          current === null ||
          current === undefined ||
          String(current).trim() === "" ||
          current === "-";

        const candidateFilled =
          candidate !== null &&
          candidate !== undefined &&
          String(candidate).trim() !== "" &&
          String(candidate).trim() !== "-";

        if (currentEmpty && candidateFilled) {
          merged[field] = candidate;
        }
      });
    });

    merged.id = primary.id || qualityExplicitId(primary) || issue.key;
    merged.ordem = primary.ordem || merged.ordem || "";
    merged.origem = "Qualidade da Base - Mesclado";
    merged.comentario = [
      merged.comentario,
      `Ajuste de qualidade: ${issue.typeLabel} | Chave ${issue.key} | ${issue.quantidade} registros analisados.`,
    ]
      .filter(Boolean)
      .join(" | ");

    return prepareDemandForSave(merged);
  }

  function showQualityMergePreview() {
    const issue = selectedQualityIssue();

    if (!issue) {
      showToast("Selecione um problema para mesclar.", "error");
      return;
    }

    const merged = mergeQualityRecords(
      issue,
      state.quality.selectedPrimarySequence,
    );

    if (!merged) {
      showToast("Não foi possível montar a prévia da mesclagem.", "error");
      return;
    }

    const preview = $("#qualityMergePreview");

    if (!preview) return;

    preview.classList.remove("hidden");

    preview.innerHTML = `
    <strong>Prévia da mesclagem</strong>
    <span>O sistema manterá o registro principal e preencherá campos vazios com dados dos registros duplicados.</span>

    <div class="detail-grid">
      <div class="detail-item">
        <span>ID final</span>
        <strong>${escapeHtml(merged.id || "-")}</strong>
      </div>
      <div class="detail-item">
        <span>OM final</span>
        <strong>${escapeHtml(merged.ordem || "-")}</strong>
      </div>
      <div class="detail-item">
        <span>Descrição</span>
        <strong>${escapeHtml(merged.descricao || "-")}</strong>
      </div>
      <div class="detail-item">
        <span>Centro</span>
        <strong>${escapeHtml(merged.centroTrabalho || "-")}</strong>
      </div>
      <div class="detail-item">
        <span>Gerência</span>
        <strong>${escapeHtml(merged.gerencia || "-")}</strong>
      </div>
      <div class="detail-item">
        <span>Supervisão</span>
        <strong>${escapeHtml(merged.supervisao || "-")}</strong>
      </div>
      <div class="detail-item">
        <span>Competência</span>
        <strong>${escapeHtml(merged.competencia || "-")}</strong>
      </div>
      <div class="detail-item">
        <span>Vencimento</span>
        <strong>${formatDate(merged.vencimento)}</strong>
      </div>
    </div>
  `;
  }

  async function saveQualityMerge() {
    const issue = selectedQualityIssue();

    if (!issue) {
      showToast("Selecione um problema para salvar.", "error");
      return;
    }

    if (!canEdit() && !canAdmin()) {
      showToast(
        "Perfil sem permissão para salvar ajuste de qualidade.",
        "error",
      );
      return;
    }

    const merged = mergeQualityRecords(
      issue,
      state.quality.selectedPrimarySequence,
    );

    if (!merged) {
      showToast("Não foi possível montar o registro mesclado.", "error");
      return;
    }

    try {
      const saved = await state.repo.upsertDemanda(merged);
      const savedRecord = normalizeDemandRecord(saved || merged);
      const existingIndex = state.db.demandas.findIndex(
        (item) => item.id === savedRecord.id,
      );

      if (existingIndex >= 0) {
        state.db.demandas[existingIndex] = {
          ...state.db.demandas[existingIndex],
          ...savedRecord,
        };
      } else {
        state.db.demandas.unshift(savedRecord);
      }

      await state.repo.addLog?.({
        usuario: state.currentUser?.email || "",
        acao: "Qualidade da Base - Mesclagem",
        lista: "controle_demandas_eletrovia",
        referencia: merged.id,
        detalhe: `${issue.typeLabel} | Chave ${issue.key} | ${issue.quantidade} registros analisados.`,
        modulo: "QUALIDADE_BASE",
        status: "SUCESSO",
      });

      state.lastDataUpdateAt = new Date().toISOString();
      $("#lastUpdateSide").textContent = formatDateTime(state.lastDataUpdateAt);

      state.quality.issuesCache = [];
      state.quality.filteredCache = [];

      showToast("Ajuste salvo no controle de demandas.", "success");
      renderQuality();
    } catch (error) {
      console.error(error);
      showToast(`Erro ao salvar ajuste: ${error.message}`, "error");
    }
  }

  function renderQualityDetail(issue) {
    const panel = $("#qualityDetailPanel");

    if (!issue) {
      panel.innerHTML = `
      <div class="empty-detail">
        <strong>Selecione um problema</strong>
        <span>Os registros envolvidos aparecem aqui.</span>
      </div>
    `;
      return;
    }

    const recommendation = recommendedQualityPrimary(issue.records);
    const recommendedSequence =
      recommendation?.record?.qualidadeSequencia || "";

    if (!state.quality.selectedPrimarySequence) {
      state.quality.selectedPrimarySequence = recommendedSequence;
    }

    const selectedPrimarySequence =
      state.quality.selectedPrimarySequence || recommendedSequence;

    const primaryRecord =
      issue.records.find(
        (record) => record.qualidadeSequencia === selectedPrimarySequence,
      ) ||
      recommendation?.record ||
      issue.records[0];

    const hiddenRecords = Math.max(0, issue.records.length - 40);
    const canSaveQuality = canEdit() || canAdmin();

    panel.innerHTML = `
    <div class="detail-title">
      <span>${escapeHtml(issue.typeLabel)}</span>
      <h3>${escapeHtml(issue.key)}</h3>
    </div>

    <div class="detail-grid">
      <div class="detail-item">
        <span>Quantidade</span>
        <strong>${issue.quantidade}</strong>
      </div>
      <div class="detail-item">
        <span>Origem</span>
        <strong>${escapeHtml(issue.origem)}</strong>
      </div>
      <div class="detail-item">
        <span>Centro</span>
        <strong>${escapeHtml(issue.centroTrabalho)}</strong>
      </div>
      <div class="detail-item">
        <span>Competência</span>
        <strong>${escapeHtml(issue.competencia)}</strong>
      </div>
    </div>

    ${
      recommendation
        ? `
          <div class="quality-recommendation">
            <strong>Principal sugerido</strong>
            <span>
              ${escapeHtml(recommendation.record.id || qualityExplicitId(recommendation.record) || recommendation.record.ordem || "-")}
              | ${escapeHtml(recommendation.reason)}
            </span>
          </div>
        `
        : ""
    }

    <div class="quality-selected-primary">
      <span>Registro principal selecionado</span>
      <strong>
        ${escapeHtml(primaryRecord?.id || qualityExplicitId(primaryRecord) || primaryRecord?.ordem || "-")}
      </strong>
      <small>
        ${escapeHtml(qualitySourceLabel(primaryRecord || {}))}
      </small>
    </div>

    <div class="quality-actions">
      <button
        class="button secondary"
        type="button"
        data-quality-action="set-primary"
      >
        Definir principal
      </button>

      <button
        class="button secondary"
        type="button"
        data-quality-action="merge-preview"
      >
        Mesclar
      </button>

      <button
        class="button"
        type="button"
        data-quality-action="save-merge"
        ${canSaveQuality ? "" : "disabled"}
        title="${canSaveQuality ? "Salvar ajuste no Supabase" : "Perfil sem permissÃ£o para salvar ajuste"}"
      >
        Salvar ajuste
      </button>
    </div>

    <div id="qualityMergePreview" class="quality-merge-preview hidden"></div>

    <div class="quality-record-list">
      ${issue.records
        .slice(0, 40)
        .map((record, index) => {
          const isRecommended =
            recommendedSequence &&
            recommendedSequence === record.qualidadeSequencia;

          const isSelected =
            selectedPrimarySequence &&
            selectedPrimarySequence === record.qualidadeSequencia;

          return `
            <article class="quality-record ${isRecommended ? "is-recommended" : ""} ${isSelected ? "is-primary-selected" : ""}" data-quality-primary="${escapeHtml(record.qualidadeSequencia)}">
              <header>
                <label class="quality-primary-option">
                  <input
                    type="radio"
                    name="qualityPrimaryRecord"
                    value="${escapeHtml(record.qualidadeSequencia)}"
                    ${isSelected ? "checked" : ""}
                  />
                  <strong>Registro ${index + 1}</strong>
                </label>

                <div class="quality-record-badges">
                  ${isRecommended ? statusChip("Principal sugerido") : ""}
                  ${isSelected ? statusChip("Principal selecionado") : ""}
                </div>
              </header>

              <div class="detail-grid">
                <div class="detail-item">
                  <span>ID</span>
                  <strong>${escapeHtml(record.id || "-")}</strong>
                </div>
                <div class="detail-item">
                  <span>ID informado</span>
                  <strong>${escapeHtml(qualityExplicitId(record) || "-")}</strong>
                </div>
                <div class="detail-item">
                  <span>OM</span>
                  <strong>${escapeHtml(record.ordem || "-")}</strong>
                </div>
                <div class="detail-item">
                  <span>Origem</span>
                  <strong>${escapeHtml(qualitySourceLabel(record))}</strong>
                </div>
                <div class="detail-item">
                  <span>Descrição</span>
                  <strong>${escapeHtml(record.descricao || "-")}</strong>
                </div>
                <div class="detail-item">
                  <span>Centro</span>
                  <strong>${escapeHtml(record.centroTrabalho || "-")}</strong>
                </div>
                <div class="detail-item">
                  <span>Gerência</span>
                  <strong>${escapeHtml(record.gerencia || "-")}</strong>
                </div>
                <div class="detail-item">
                  <span>Supervisão</span>
                  <strong>${escapeHtml(record.supervisao || "-")}</strong>
                </div>
                <div class="detail-item">
                  <span>Competência</span>
                  <strong>${escapeHtml(record.competencia || "-")}</strong>
                </div>
                <div class="detail-item">
                  <span>Vencimento</span>
                  <strong>${formatDate(record.vencimento)}</strong>
                </div>
                <div class="detail-item">
                  <span>Status</span>
                  <strong>${statusChip(primaryStatusOf(record))}</strong>
                </div>
                <div class="detail-item">
                  <span>Responsável</span>
                  <strong>${escapeHtml(record.usuarioResponsavel || "-")}</strong>
                </div>
              </div>

              ${
                record.comentario || record.observacao
                  ? `
                    <div class="quality-note">
                      ${escapeHtml([record.comentario, record.observacao].filter(Boolean).join(" | "))}
                    </div>
                  `
                  : ""
              }
            </article>
          `;
        })
        .join("")}

      ${
        hiddenRecords
          ? `<span class="muted">Mais ${hiddenRecords} registros neste grupo.</span>`
          : ""
      }
    </div>
  `;
  }

  function getQualityIssuesCached() {
    if (!state.quality.issuesCache.length) {
      state.quality.issuesCache = buildQualityIssues();
    }

    return state.quality.issuesCache;
  }

  function renderQuality() {
    const allIssues = getQualityIssuesCached();
    const filtered = filteredQualityIssues(allIssues);

    state.quality.filteredCache = filtered;

    const totalPages = Math.max(
      1,
      Math.ceil(filtered.length / state.quality.pageSize),
    );

    if (state.quality.page > totalPages) {
      state.quality.page = totalPages;
    }

    const selectedInFilter = filtered.some(
      (issue) => issue.id === state.quality.selectedIssueId,
    );

    if (!selectedInFilter) {
      state.quality.selectedIssueId = filtered[0]?.id || "";
      state.quality.selectedPrimarySequence = "";
    }

    const start = (state.quality.page - 1) * state.quality.pageSize;
    const pageRows = filtered.slice(start, start + state.quality.pageSize);

    $("#qualityTypeFilter").value = state.quality.typeFilter;
    $("#qualitySearch").value = state.quality.search;

    $("#qualityCount").textContent =
      `${filtered.length} problemas encontrados • Página ${state.quality.page} de ${totalPages}`;

    renderQualityCards(allIssues);

    const tbody = $("#qualityIssueTableBody");

    tbody.innerHTML = pageRows.length
      ? pageRows
          .map((issue) => {
            const selected =
              issue.id === state.quality.selectedIssueId ? "is-selected" : "";

            return `
            <tr class="${selected}" data-quality-issue-id="${escapeHtml(issue.id)}">
              <td>${statusChip(issue.typeLabel)}</td>
              <td><strong>${escapeHtml(issue.key)}</strong></td>
              <td>${issue.quantidade}</td>
              <td>${escapeHtml(issue.origem)}</td>
              <td class="description-cell">${escapeHtml(issue.descricao)}</td>
              <td>${escapeHtml(issue.centroTrabalho)}</td>
              <td>${escapeHtml(issue.gerencia)}</td>
              <td>${escapeHtml(issue.supervisao)}</td>
              <td>${escapeHtml(issue.competencia)}</td>
              <td>
                <button
                  class="button secondary compact-button"
                  data-quality-detail="${escapeHtml(issue.id)}"
                  type="button"
                >
                  Ver detalhes
                </button>
              </td>
            </tr>
          `;
          })
          .join("")
      : `
      <tr>
        <td colspan="10">
          <div class="empty-detail">
            <strong>Nenhum problema no recorte</strong>
            <span>Ajuste os filtros para consultar outros grupos.</span>
          </div>
        </td>
      </tr>
    `;

    const selectedIssue = filtered.find(
      (issue) => issue.id === state.quality.selectedIssueId,
    );

    renderQualityDetail(selectedIssue);

    renderQualityPager(totalPages);
  }

  function renderQualityPager(totalPages) {
    const panel = $("#qualityCount")?.closest(".panel-toolbar");
    if (!panel) return;

    let pager = $("#qualityPager");

    if (!pager) {
      pager = document.createElement("div");
      pager.id = "qualityPager";
      pager.className = "quality-pager";
      panel.appendChild(pager);
    }

    pager.innerHTML = `
    <button
      class="button secondary compact-button"
      id="qualityPrevPage"
      type="button"
      ${state.quality.page <= 1 ? "disabled" : ""}
    >
      Anterior
    </button>

    <span>Página <strong>${state.quality.page}</strong> de ${totalPages}</span>

    <button
      class="button secondary compact-button"
      id="qualityNextPage"
      type="button"
      ${state.quality.page >= totalPages ? "disabled" : ""}
    >
      Próxima
    </button>
  `;

    $("#qualityPrevPage")?.addEventListener("click", () => {
      state.quality.page = Math.max(1, state.quality.page - 1);
      renderQuality();
    });

    $("#qualityNextPage")?.addEventListener("click", () => {
      state.quality.page += 1;
      renderQuality();
    });
  }

  function distinctValues(records, selector) {
    return uniqueOptions(
      records
        .map(selector)
        .map((value) => String(value || "").trim())
        .filter(Boolean),
    );
  }

  function mostFrequentValue(records, selector) {
    const counts = new Map();
    records.forEach((record) => {
      const value = String(selector(record) || "").trim();
      if (!value) return;
      counts.set(value, (counts.get(value) || 0) + 1);
    });

    return [...counts.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] || "-";
  }

  function qualityDivergentFields(records) {
    const fields = [
      ["ID", (record) => qualityExplicitId(record) || record.id],
      ["Origem", qualitySourceLabel],
      ["Descricao", (record) => record.descricao],
      ["Centro", (record) => record.centroTrabalho],
      ["Local", (record) => record.localInstalacao],
      ["Competencia", (record) => record.competencia],
      ["Vencimento", (record) => record.vencimento],
      ["Status SAP", (record) => record.statusSistema],
      ["Status usuario", (record) => record.statusUsuario],
      ["Realizado", (record) => record.dataRealizada],
    ];

    return fields
      .filter(([, selector]) => distinctValues(records, selector).length > 1)
      .map(([label]) => label);
  }

  function qualityPrimaryRecommendation(records) {
    const realizedSource = records.find((record) =>
      normalizeText(qualitySourceLabel(record)).includes("REALIZADOS"),
    );
    if (realizedSource) {
      return `${qualityExplicitId(realizedSource) || realizedSource.id || realizedSource.ordem} | realizados`;
    }

    const futureSource = records.find((record) =>
      normalizeText(qualitySourceLabel(record)).includes("FUTURAS"),
    );
    const orderSource = records.find((record) =>
      normalizeText(qualitySourceLabel(record)).includes("ORDENS"),
    );

    if (futureSource && orderSource) {
      return `${qualityExplicitId(futureSource) || futureSource.id || futureSource.ordem} | manter ID futura e status da ordem`;
    }

    const richer = [...records].sort((a, b) => {
      const score = (record) =>
        [
          record.id,
          qualityExplicitId(record),
          record.ordem,
          record.descricao,
          record.centroTrabalho,
          record.localInstalacao,
          record.statusSistema,
          record.statusUsuario,
          record.dataRealizada,
          record.vencimento,
          record.competencia,
        ].filter((value) => String(value || "").trim()).length;
      return score(b) - score(a);
    })[0];

    return `${qualityExplicitId(richer) || richer?.id || richer?.ordem || "-"} | mais completo`;
  }

  function filteredDuplicateOmIssues() {
    const search = normalizeText(state.quality.duplicateOmSearch);
    return getQualityIssuesCached().filter((issue) => {
      if (issue.typeKey !== "om-duplicada") return false;
      if (!search) return true;
      return issue.searchText.includes(search);
    });
  }

  function openQualityIssueFromFocusedView(issueId) {
    state.quality.typeFilter = "";
    state.quality.search = "";
    state.quality.selectedIssueId = issueId;
    state.quality.selectedPrimarySequence = "";
    switchView("qualidade");
  }

  function renderDuplicateOms() {
    const rows = filteredDuplicateOmIssues();
    const involved = rows.reduce((total, issue) => total + issue.quantidade, 0);
    const withDivergence = rows.filter(
      (issue) => qualityDivergentFields(issue.records).length,
    ).length;

    $("#duplicateOmSearch").value = state.quality.duplicateOmSearch;
    $("#duplicateOmSummary").innerHTML = [
      ["OMs duplicadas", rows.length, "grupos encontrados"],
      ["Registros envolvidos", involved, "linhas nas fontes"],
      ["Com divergência", withDivergence, "campos conflitantes"],
    ]
      .map(
        ([label, value, hint]) => `
          <div class="quality-insight-card">
            <span>${escapeHtml(label)}</span>
            <strong>${escapeHtml(value)}</strong>
            <small>${escapeHtml(hint)}</small>
          </div>
        `,
      )
      .join("");

    $("#duplicateOmCount").textContent = `${rows.length} OMs duplicadas`;

    $("#duplicateOmTableBody").innerHTML = rows.length
      ? rows
          .map((issue) => {
            const divergent = qualityDivergentFields(issue.records);
            return `
              <tr data-duplicate-om="${escapeHtml(issue.id)}">
                <td><strong>${escapeHtml(issue.key)}</strong></td>
                <td>${issue.quantidade}</td>
                <td>${escapeHtml(issue.origem)}</td>
                <td>${escapeHtml(divergent.join(", ") || "Sem divergencia visivel")}</td>
                <td>${escapeHtml(qualityPrimaryRecommendation(issue.records))}</td>
                <td class="description-cell">${escapeHtml(issue.descricao)}</td>
                <td>${escapeHtml(issue.centroTrabalho)}</td>
                <td>${escapeHtml(firstFilled(issue.records, (item) => item.localInstalacao) || "-")}</td>
                <td>
                  <button
                    class="button secondary compact-button"
                    data-open-quality-issue="${escapeHtml(issue.id)}"
                    type="button"
                  >
                    Ver saneamento
                  </button>
                </td>
              </tr>
            `;
          })
          .join("")
      : `<tr>
          <td colspan="9">
            <div class="empty-detail">
              <strong>Nenhuma OM duplicada no recorte</strong>
              <span>Ajuste a busca para consultar outros grupos.</span>
            </div>
          </td>
        </tr>`;
  }

  function buildQualityLocalGroups() {
    const filters = state.quality.localFilters || {};
    const cacheKey = JSON.stringify({
      filters,
      total: state.db?.demandas?.length || 0,
      lastUpdate: state.lastDataUpdateAt || "",
    });

    if (state.quality.localGroupsCache?.key === cacheKey) {
      return state.quality.localGroupsCache.groups;
    }

    const groups = new Map();
    const rows = (state.db?.demandas || []).filter((demand) =>
      demandMatchesFilters(demand, filters),
    );

    rows.forEach((demand) => {
      const key = String(demand.localInstalacao || "").trim();
      if (!key) return;
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(demand);
    });

    const result = [...groups.entries()]
      .map(([local, demandas]) => {
        const ordens = distinctValues(demandas, (item) => item.ordem);
        return {
          local,
          demandas,
          quantidade: demandas.length,
          ordens: ordens.length,
          semOm: demandas.filter((item) => !String(item.ordem || "").trim())
            .length,
          centroPrincipal: mostFrequentValue(
            demandas,
            (item) => item.centroTrabalho,
          ),
          searchText: normalizeText(
            [
              local,
              ...demandas.flatMap((item) => [
                item.id,
                item.ordem,
                item.descricao,
                item.centroTrabalho,
                item.gerencia,
                item.supervisao,
              ]),
            ].join(" "),
          ),
        };
      })
      .sort(
        (a, b) =>
          b.quantidade - a.quantidade || a.local.localeCompare(b.local),
      );

    state.quality.localGroupsCache = {
      key: cacheKey,
      groups: result,
    };

    return result;
  }

  function filteredQualityLocalGroups() {
    const search = normalizeText(state.quality.localSearch);
    return buildQualityLocalGroups().filter(
      (group) => !search || group.searchText.includes(search),
    );
  }

  function renderQualityByLocal() {
    renderQualityLocalFilterVisibility();

    if (state.quality.localFiltersVisible && !state.quality.localFiltersReady) {
      buildQualityLocalFilterOptions();
    }

    const groups = filteredQualityLocalGroups();
    const selectedExists = groups.some(
      (group) => group.local === state.quality.selectedLocal,
    );

    if (!selectedExists) {
      state.quality.selectedLocal = groups[0]?.local || "";
    }

    const selected = groups.find(
      (group) => group.local === state.quality.selectedLocal,
    );

    $("#qualityLocalSearch").value = state.quality.localSearch;
    $("#qualityLocalQuickSearch").value =
      state.quality.localFilters?.quickSearch || "";
    $("#qualityLocalCount").textContent = `${groups.length} locais encontrados`;
    $("#qualityLocalDemandCount").textContent = selected
      ? `${selected.quantidade} demandas em ${selected.local}`
      : "Selecione um local";
    $("#qualityLocalSelection").innerHTML = selected
      ? `
        <div>
          <span>Local selecionado</span>
          <strong>${escapeHtml(selected.local)}</strong>
        </div>
        <div>
          <span>Demandas</span>
          <strong>${selected.quantidade}</strong>
        </div>
        <div>
          <span>OMs</span>
          <strong>${selected.ordens}</strong>
        </div>
        <div>
          <span>Sem OM</span>
          <strong>${selected.semOm}</strong>
        </div>
        <div>
          <span>Centro</span>
          <strong>${escapeHtml(selected.centroPrincipal)}</strong>
        </div>
      `
      : "";

    $("#qualityLocalGroupTableBody").innerHTML = groups.length
      ? groups
          .slice(0, 500)
          .map(
            (group) => `
              <tr
                class="${group.local === state.quality.selectedLocal ? "is-selected" : ""}"
                data-quality-local="${escapeHtml(group.local)}"
              >
                <td><strong>${escapeHtml(group.local)}</strong></td>
                <td class="metric-cell">${group.quantidade}</td>
                <td class="metric-cell">${group.ordens}</td>
                <td class="metric-cell">${group.semOm}</td>
                <td class="metric-cell">${escapeHtml(group.centroPrincipal)}</td>
                <td>
                  <button
                    class="button secondary compact-button"
                    data-select-quality-local="${escapeHtml(group.local)}"
                    type="button"
                  >
                    Ver demandas
                  </button>
                </td>
              </tr>
            `,
          )
          .join("")
      : `<tr>
          <td colspan="6">
            <div class="empty-detail">
              <strong>Nenhum local encontrado</strong>
              <span>Ajuste a busca ou os filtros ocultos para consultar outros locais.</span>
            </div>
          </td>
        </tr>`;

    $("#qualityLocalDemandTableBody").innerHTML = selected?.demandas?.length
      ? selected.demandas
          .slice()
          .sort((a, b) =>
            String(a.vencimento || "").localeCompare(String(b.vencimento || "")),
          )
          .map(
            (demand) => `
              <tr data-local-demand="${escapeHtml(demand.id)}">
                <td><strong>${escapeHtml(demand.id)}</strong></td>
                <td>${escapeHtml(demand.ordem || "-")}</td>
                <td class="description-cell">${escapeHtml(demand.descricao || "-")}</td>
                <td>${statusChip(primaryStatusOf(demand))}</td>
                <td>${escapeHtml(demand.centroTrabalho || "-")}</td>
                <td>${escapeHtml(demand.competencia || "-")}</td>
                <td>${formatDate(demand.vencimento)}</td>
                <td>${formatDate(demand.dataPlanejada)}</td>
                <td>${formatDate(demand.dataReplanejadaAtual)}</td>
              </tr>
            `,
          )
          .join("")
      : `<tr>
          <td colspan="9">
            <div class="empty-detail">
              <strong>Nenhuma demanda selecionada</strong>
              <span>Escolha um local para aplicar o filtro interno.</span>
            </div>
          </td>
        </tr>`;
  }

  function filteredMissingCenterIssues() {
    const search = normalizeText(state.quality.centersSearch);
    return getQualityIssuesCached().filter((issue) => {
      if (issue.typeKey !== "centro-sem-cadastro") return false;
      if (!search) return true;
      return issue.searchText.includes(search);
    });
  }

  function renderQualityCenters() {
    const rows = filteredMissingCenterIssues();
    $("#qualityCentersSearch").value = state.quality.centersSearch;
    $("#qualityCentersCount").textContent = `${rows.length} centros encontrados`;

    $("#qualityCentersTableBody").innerHTML = rows.length
      ? rows
          .map(
            (issue) => `
              <tr>
                <td><strong>${escapeHtml(issue.key)}</strong></td>
                <td>${issue.quantidade}</td>
                <td>${escapeHtml(issue.gerencia)}</td>
                <td>${escapeHtml(issue.supervisao)}</td>
                <td class="description-cell">${escapeHtml(issue.descricao)}</td>
                <td>
                  <button
                    class="button secondary compact-button admin-only"
                    data-open-center-admin="${escapeHtml(issue.key)}"
                    type="button"
                  >
                    Cadastrar
                  </button>
                </td>
              </tr>
            `,
          )
          .join("")
      : `<tr>
          <td colspan="6">
            <div class="empty-detail">
              <strong>Nenhum centro pendente no recorte</strong>
              <span>Os centros cadastrados pela Administração saem daqui.</span>
            </div>
          </td>
        </tr>`;
  }

  function buildQualityDivergenceRows() {
    const records = [
      ...(state.db?.qualitySourceRecords || []),
      ...(state.db?.qualityBaseRealizados || []).map((record, index) => ({
        ...record,
        fonteQualidade: "SAP BO - Realizados",
        qualidadeSequencia: `realizados-${index + 1}`,
      })),
    ];

    const groups = new Map();
    records.forEach((record) => {
      const key =
        String(record.ordem || "").trim() ||
        qualityExplicitId(record) ||
        String(record.id || "").trim();
      if (!key) return;
      addQualityGroup(groups, key, record);
    });

    return [...groups.entries()]
      .map(([key, group]) => {
        const fontes = distinctValues(group, qualitySourceLabel);
        if (fontes.length < 2) return null;

        const statusSistema = distinctValues(group, (item) =>
          formatSapStatusFilter(item.statusSistema),
        );
        const statusUsuario = distinctValues(group, (item) =>
          formatSapStatusFilter(item.statusUsuario),
        );
        const dataRealizada = distinctValues(group, (item) =>
          formatDate(item.dataRealizada),
        );
        const statusOperacional = distinctValues(group, primaryStatusOf);

        const divergencias = [];
        if (statusSistema.length > 1) divergencias.push("Status sistema");
        if (statusUsuario.length > 1) divergencias.push("Status usuario");
        if (dataRealizada.length > 1) divergencias.push("Data realizada");
        if (statusOperacional.length > 1) divergencias.push("Status operacional");

        if (!divergencias.length) return null;

        return {
          key,
          fontes,
          divergencias,
          statusPorFonte: group
            .map(
              (item) =>
                `${qualitySourceLabel(item)}: ${formatSapStatusFilter(item.statusSistema) || "-"} / ${formatSapStatusFilter(item.statusUsuario) || "-"}`,
            )
            .join(" | "),
          dataRealizada: dataRealizada.join(" | ") || "-",
          recommendation: group.some((item) =>
            normalizeText(qualitySourceLabel(item)).includes("REALIZADOS"),
          )
            ? "Usar base_realizados como fonte oficial de baixa."
            : "Comparar fonte operacional antes de salvar saneamento.",
          searchText: normalizeText(
            [
              key,
              fontes.join(" "),
              divergencias.join(" "),
              ...group.flatMap((item) => [
                item.id,
                qualityExplicitId(item),
                item.descricao,
                item.centroTrabalho,
                item.localInstalacao,
                item.statusSistema,
                item.statusUsuario,
              ]),
            ].join(" "),
          ),
        };
      })
      .filter(Boolean)
      .sort((a, b) => a.key.localeCompare(b.key, "pt-BR"));
  }

  function renderQualityDivergences() {
    const search = normalizeText(state.quality.divergenceSearch);
    const rows = buildQualityDivergenceRows().filter(
      (row) => !search || row.searchText.includes(search),
    );

    $("#qualityDivergenceSearch").value = state.quality.divergenceSearch;
    $("#qualityDivergenceCount").textContent =
      `${rows.length} divergências encontradas`;

    $("#qualityDivergenceTableBody").innerHTML = rows.length
      ? rows
          .slice(0, 500)
          .map(
            (row) => `
              <tr>
                <td><strong>${escapeHtml(row.key)}</strong></td>
                <td>${escapeHtml(row.fontes.join(" | "))}</td>
                <td>${escapeHtml(row.divergencias.join(", "))}</td>
                <td class="description-cell">${escapeHtml(row.statusPorFonte)}</td>
                <td>${escapeHtml(row.dataRealizada)}</td>
                <td>${escapeHtml(row.recommendation)}</td>
              </tr>
            `,
          )
          .join("")
      : `<tr>
          <td colspan="6">
            <div class="empty-detail">
              <strong>Nenhuma divergência no recorte</strong>
              <span>Quando as bases discordarem, o conflito aparecerá aqui.</span>
            </div>
          </td>
        </tr>`;
  }

  function activeVacationSubstitutionsForCurrentUser() {
    const email = normalizeText(state.currentUser?.email || "");
    const today = toDate(todayText());

    return (state.db.feriasSubstituicoes || []).filter((item) => {
      if (!item.ativo) return false;

      const inicio = toDate(item.dataInicio);
      const fim = toDate(item.dataFim);

      if (!inicio || !fim) return false;

      const isInPeriod = today >= inicio && today <= fim;
      const isSubstitute = normalizeText(item.emailSubstituto) === email;

      return isInPeriod && isSubstitute;
    });
  }

  function responsibleEmailsForDemand(demand) {
    return [
      demand.usuarioResponsavel,
      demand.planejadorCurtoEmail,
      demand.planejadorOMEmail,
      demand.programadorEmail,
    ]
      .filter(Boolean)
      .map(normalizeText);
  }

  function demandMatchesVacationScope(demand, substitution) {
    const gerencia = normalizeText(substitution.escopoGerencia || "");
    const centro = normalizeText(substitution.escopoCentroTrabalho || "");

    const demandGerencia = normalizeText(demand.gerencia || "");
    const demandCentro = normalizeText(demand.centroTrabalho || "");

    const gerenciaOk = !gerencia || demandGerencia === gerencia;
    const centroOk = !centro || demandCentro === centro;

    return gerenciaOk && centroOk;
  }

  function demandIsVisibleForCurrentUserAlerts(demand) {
    const profile = state.currentUser?.perfil || "";

    if (profile === "Administrador" || profile === "Gestor") {
      return true;
    }

    const currentEmail = normalizeText(state.currentUser?.email || "");
    const demandEmails = responsibleEmailsForDemand(demand);

    if (demandEmails.includes(currentEmail)) {
      return true;
    }

    const substitutions = activeVacationSubstitutionsForCurrentUser();

    return substitutions.some((substitution) => {
      const absentEmail = normalizeText(substitution.emailAusente || "");

      return (
        demandEmails.includes(absentEmail) &&
        demandMatchesVacationScope(demand, substitution)
      );
    });
  }

  function notificationTypeMeta(type) {
    const meta = {
      vencida: {
        label: "Vencida",
        group: "vencida",
        criticality: "Crítica",
        className: "critical",
      },
      vencendo: {
        label: "Perto do vencimento",
        group: "vencendo",
        criticality: "Atenção",
        className: "warning",
      },
      "replanejamento-incompleto": {
        label: "Replanejamento incompleto",
        group: "replanejamento-incompleto",
        criticality: "Atenção",
        className: "warning",
      },
      "realizada-fora-prazo": {
        label: "Realizada fora do prazo",
        group: "realizada-fora-prazo",
        criticality: "Crítica",
        className: "critical",
      },
      "perda-incompleta": {
        label: "Perda incompleta",
        group: "perda-incompleta",
        criticality: "Crítica",
        className: "critical",
      },
    };

    return (
      meta[type] || {
        label: type,
        group: type,
        criticality: "Informativo",
        className: "info",
      }
    );
  }

  function daysFromToday(value) {
    const date = toDate(value);
    const today = toDate(todayText());

    if (!date || !today) return null;

    return Math.ceil((date - today) / 86400000);
  }

  function isDemandOpenForDueNotification(demand) {
    const status = primaryStatusOf(demand);

    return status !== "Realizado" && status !== "Cancelado";
  }

  function latestRealizadoPerdaHistory(demandId) {
    return [...(state.db.historicoRealizadoPerdas || [])]
      .filter((item) => item.demandaId === demandId)
      .sort((a, b) => new Date(b.dataHora || 0) - new Date(a.dataHora || 0))[0];
  }

  function makeNotification({
    type,
    demand,
    message,
    days = null,
    action = "abrir",
    historyId = "",
    source = "",
  }) {
    const meta = notificationTypeMeta(type);

    return {
      id: [type, demand.id, historyId, source, message]
        .filter(Boolean)
        .join("|"),
      type,
      group: meta.group,
      label: meta.label,
      criticality: meta.criticality,
      className: meta.className,
      demandId: demand.id,
      ordem: demand.ordem || "",
      descricao: demand.descricao || "",
      gerencia: demand.gerencia || "",
      supervisao: demand.supervisao || "",
      centroTrabalho: demand.centroTrabalho || "",
      vencimento: demand.vencimento || "",
      status: primaryStatusOf(demand),
      message,
      days,
      action,
      historyId,
    };
  }

  function notificationsCacheKey() {
    return [
      state.db?.demandas?.length || 0,
      state.db?.historicoReplanejamento?.length || 0,
      state.db?.historicoRealizadoPerdas?.length || 0,
      state.lastDataUpdateAt || "",
      state.currentUser?.email || "",
      todayText(),
    ].join("|");
  }

  function clearNotificationsCache() {
    state.notifications.cache = [];
    state.notifications.cacheKey = "";
    state.notifications.byId = new Map();
  }

  function groupByDemandId(rows) {
    const map = new Map();

    (rows || []).forEach((row) => {
      if (!row.demandaId) return;

      if (!map.has(row.demandaId)) {
        map.set(row.demandaId, []);
      }

      map.get(row.demandaId).push(row);
    });

    return map;
  }

  function getNotificationsCached() {
    const key = notificationsCacheKey();

    if (
      state.notifications.cacheKey === key &&
      Array.isArray(state.notifications.cache) &&
      state.notifications.cache.length
    ) {
      return state.notifications.cache;
    }

    const notifications = buildNotificationsOptimized();

    state.notifications.cache = notifications;
    state.notifications.cacheKey = key;
    state.notifications.byId = new Map(
      notifications.map((item) => [item.id, item]),
    );

    return notifications;
  }

  function buildNotificationsOptimized() {
    const notifications = [];
    const today = toDate(todayText());

    const replanByDemandId = groupByDemandId(
      state.db.historicoReplanejamento || [],
    );

    const visibleDemands = (state.db.demandas || []).filter(
      demandIsVisibleForCurrentUserAlerts,
    );

    visibleDemands.forEach((demand) => {
      const status = primaryStatusOf(demand);
      const due = toDate(demand.vencimento);
      const days = due && today ? Math.ceil((due - today) / 86400000) : null;

      if (isDemandOpenForDueNotification(demand) && days !== null && days < 0) {
        notifications.push(
          makeNotification({
            type: "vencida",
            demand,
            days,
            action:
              demand.dataPlanejada || demand.dataReplanejadaAtual
                ? "replanejar"
                : "planejar",
            message: `Vencida há ${Math.abs(days)} dia(s).`,
          }),
        );
      }

      if (
        isDemandOpenForDueNotification(demand) &&
        days !== null &&
        days >= 0 &&
        days <= 7
      ) {
        notifications.push(
          makeNotification({
            type: "vencendo",
            demand,
            days,
            action:
              demand.dataPlanejada || demand.dataReplanejadaAtual
                ? "abrir"
                : "planejar",
            message: `Vence em ${days} dia(s).`,
          }),
        );
      }

      const replanHistory = replanByDemandId.get(demand.id) || [];

      replanHistory.forEach((history) => {
        if (!history.motivo || !history.justificativa) {
          const missing = [
            !history.motivo ? "motivo" : "",
            !history.justificativa ? "justificativa" : "",
          ]
            .filter(Boolean)
            .join(" e ");

          notifications.push(
            makeNotification({
              type: "replanejamento-incompleto",
              demand,
              historyId: history.id,
              action: "regularizar-replanejamento",
              source: "historico-replanejamento",
              message: `Replanejamento sem ${missing}.`,
            }),
          );
        }
      });

      const foraPrazo =
        status === "Realizado" && dueClassOf(demand) === "Fora do Prazo";

      if (foraPrazo && !demand.perda) {
        notifications.push(
          makeNotification({
            type: "realizada-fora-prazo",
            demand,
            action: "regularizar-perda",
            source: "fora-prazo-sem-perda",
            message:
              "Realizada fora da tolerância e ainda sem registro de perda.",
          }),
        );
      }

      if (
        foraPrazo &&
        demand.perda &&
        (!demand.motivoPerda || !demand.justificativaPerda)
      ) {
        const missing = [
          !demand.motivoPerda ? "perfil de perda" : "",
          !demand.justificativaPerda ? "justificativa de perda" : "",
        ]
          .filter(Boolean)
          .join(" e ");

        notifications.push(
          makeNotification({
            type: "perda-incompleta",
            demand,
            action: "regularizar-perda",
            source: "perda-incompleta-fora-prazo",
            message: `Perda fora do prazo sem ${missing}.`,
          }),
        );
      }

      if (demand.perda && (!demand.motivoPerda || !demand.justificativaPerda)) {
        const missing = [
          !demand.motivoPerda ? "perfil de perda" : "",
          !demand.justificativaPerda ? "justificativa de perda" : "",
        ]
          .filter(Boolean)
          .join(" e ");

        notifications.push(
          makeNotification({
            type: "perda-incompleta",
            demand,
            action: "regularizar-perda",
            source: "perda-incompleta",
            message: `Perda sem ${missing}.`,
          }),
        );
      }
    });

    const unique = new Map();

    notifications.forEach((notification) => {
      if (!unique.has(notification.id)) {
        unique.set(notification.id, notification);
      }
    });

    return Array.from(unique.values()).sort((a, b) => {
      const priority = {
        Crítica: 1,
        Atenção: 2,
        Informativo: 3,
      };

      const priorityDiff =
        (priority[a.criticality] || 9) - (priority[b.criticality] || 9);

      if (priorityDiff) return priorityDiff;

      return String(a.vencimento || "").localeCompare(
        String(b.vencimento || ""),
      );
    });
  }

  function buildNotifications() {
    return getNotificationsCached();
  }
  function notificationStats(notifications) {
    const stats = {
      total: notifications.length,
      vencida: 0,
      vencendo: 0,
      "replanejamento-incompleto": 0,
      "realizada-fora-prazo": 0,
      "perda-incompleta": 0,
    };

    notifications.forEach((item) => {
      stats[item.group] = (stats[item.group] || 0) + 1;
    });

    return stats;
  }

  function filteredNotifications() {
    const typeFilter = state.notifications.typeFilter;
    const search = normalizeText(state.notifications.search);

    return getNotificationsCached().filter((notification) => {
      if (typeFilter && notification.group !== typeFilter) return false;

      if (!search) return true;

      const haystack = normalizeText(
        [
          notification.label,
          notification.criticality,
          notification.ordem,
          notification.demandId,
          notification.descricao,
          notification.gerencia,
          notification.supervisao,
          notification.centroTrabalho,
          notification.message,
          notification.status,
        ].join(" "),
      );

      return haystack.includes(search);
    });
  }

  function notificationActionLabel(action) {
    const labels = {
      planejar: "Planejar",
      replanejar: "Replanejar",
      abrir: "Abrir demanda",
      "regularizar-replanejamento": "Regularizar replanejamento",
      "regularizar-perda": "Regularizar perda",
    };

    return labels[action] || "Abrir";
  }

  function findNotification(notificationId) {
    getNotificationsCached();

    return (
      state.notifications.byId.get(notificationId) ||
      state.notifications.cache.find((item) => item.id === notificationId)
    );
  }

  function renderAlerts() {
    const notifications = getNotificationsCached();
    const stats = notificationStats(notifications);

    $("#alertCount").textContent = String(stats.total);
    $("#alertButton").classList.toggle("has-alerts", stats.total > 0);
    $("#alertButton").classList.toggle(
      "has-critical-alerts",
      notifications.some((item) => item.criticality === "Crítica"),
    );

    $("#alertMenu").innerHTML =
      `<strong>Alertas operacionais</strong>` +
      (notifications.length
        ? `
        <div class="alert-menu-summary">
          <span>${stats.vencida || 0} vencidas</span>
          <span>${stats.vencendo || 0} vencendo</span>
          <span>${stats["replanejamento-incompleto"] || 0} replanejamentos</span>
          <span>${stats["perda-incompleta"] || 0} perdas</span>
        </div>

        ${notifications
          .slice(0, 8)
          .map(
            (notification) => `
              <button
                type="button"
                class="alert-menu-item ${notification.className}"
                data-alert-notification="${escapeHtml(notification.id)}"
              >
                <span>${escapeHtml(notification.label)} | ${escapeHtml(notification.criticality)}</span>
                <strong>${escapeHtml(notification.ordem || notification.demandId)}</strong>
                <small>${escapeHtml(notification.message)}</small>
              </button>
            `,
          )
          .join("")}

        <button
          type="button"
          class="alert-menu-open-panel"
          data-alert-open-panel
        >
          Abrir Tratamento Operacional
        </button>
      `
        : "<p>Sem alertas críticos no momento.</p>");
  }

  function renderKpis(demands) {
    const stats = dashboardStats(demands);
    const kpis = [
      ["Total", stats.total, "registros filtrados"],
      ["A Planejar", stats.aPlanejar, "sem data planejada"],
      ["Planejadas", stats.planejadas, "com data ativa"],
      ["Replanejadas", stats.replanejadas, "alteradas"],
      ["Realizadas", stats.realizadas, "com baixa"],
      ["Perdas", stats.perdas, `${stats.pendentesPerda} pendentes`],
    ];
    $("#kpiStrip").innerHTML = kpis
      .map(
        ([label, value, note]) =>
          `<article class="kpi-card"><span>${label}</span><strong>${value}</strong><small>${note}</small></article>`,
      )
      .join("");
  }

  function renderCarteira() {
    const filtered = filteredDemandas();
    renderKpis(filtered);

    const totalPages = Math.max(1, Math.ceil(filtered.length / state.pageSize));
    if (state.page > totalPages) state.page = totalPages;
    const start = (state.page - 1) * state.pageSize;
    const pageRows = filtered.slice(start, start + state.pageSize);
    const tbody = $("#demandTableBody");

    tbody.innerHTML = pageRows
      .map((item) => {
        const status = primaryStatusOf(item);
        const substatuses = substatusListOf(item);
        const allowed = allowedActionsFor(item);
        const selected =
          item.id === state.selectedDemandId ? "is-selected" : "";
        const planDisabled = !canPlan();
        const replanDisabled = !canReplan();
        const realizedDisabled = !canRealizar();
        return `
          <tr class="${selected}" data-demand-id="${escapeHtml(item.id)}">
            <td><strong>${escapeHtml(item.id)}</strong></td>
            <td>${escapeHtml(compact(item.ordem))}</td>

            <td class="description-cell">
              ${escapeHtml(item.descricao)}
              <div class="muted">${escapeHtml(item.usuarioResponsavel || "")}</div>
            </td>

            <td class="sap-status-cell">${escapeHtml(formatSapStatusFilter(item.statusSistema) || "-")}</td>
            <td class="sap-status-cell">${escapeHtml(formatSapStatusFilter(item.statusUsuario) || "-")}</td>

            <td>${statusChip(status)}</td>
            <td>${statusChipGroup(substatuses)}</td>
            <td>${escapeHtml(item.origem || "-")}</td>
            <td>${escapeHtml(item.gerencia || "-")}</td>
            <td>${escapeHtml(item.supervisao || "-")}</td>
            <td>${formatDate(item.vencimento)}</td>
            <td>${escapeHtml(item.competencia || "-")}</td>
            <td>${escapeHtml(item.tipoOM || "-")}</td>
            <td>${escapeHtml(item.centroTrabalho || "-")}</td>
            <td>${statusChip(item.centroTrabalhoStatus || "Nao cadastrado")}</td>
            <td>${escapeHtml(item.planejadorCurto || "-")}</td>
            <td>${escapeHtml(item.planejadorOM || "-")}</td>
            <td>${escapeHtml(item.programador || "-")}</td>
            <td>${escapeHtml(item.localInstalacao || "-")}</td>
            <td>${escapeHtml(item.prioridade || "-")}</td>
            <td>${escapeHtml(item.critico || "Não informado")}</td>
            <td>${formatDate(item.dataPlanejada)}</td>
            <td>${formatDate(item.dataReplanejadaAtual)}</td>
            <td>${formatDate(item.dataRealizada)}</td>
            <td>${item.perda ? "Sim" : "Não"}</td>
            <td>${escapeHtml(compact(item.motivoPerda))}</td>
            <td>${formatDateTime(item.dataUltimaAtualizacao)}</td>
            <td>
              <div class="row-actions">
                ${actionButton("planejar", item.id, planDisabled || !allowed.planejar)}
                ${actionButton("replanejar", item.id, replanDisabled || !allowed.replanejar)}
                ${actionButton("realizado", item.id, realizedDisabled || !allowed.realizado)}
                ${actionButton("historico", item.id)}
              </div>
            </td>
          </tr>
        `;
      })
      .join("");

    $("#resultCount").textContent = `${filtered.length} registros filtrados`;
    $("#pageInfo").textContent = `Página ${state.page} de ${totalPages}`;
    $("#prevPage").disabled = state.page <= 1;
    $("#nextPage").disabled = state.page >= totalPages;
    renderDetail();
  }

  function renderDetail() {
    const panel = $("#detailPanel");
    const demand = demandById(state.selectedDemandId);
    if (!demand) {
      panel.innerHTML = `
        <div class="empty-detail">
          <strong>Selecione uma demanda</strong>
          <span>O resumo operacional, histórico e ações permitidas aparecem aqui.</span>
        </div>
      `;
      return;
    }
    const status = primaryStatusOf(demand);
    const substatuses = substatusListOf(demand);
    const pending = pendingIssuesOf(demand);
    const allowed = allowedActionsFor(demand);
    const timeline = historiesFor(demand.id).slice(0, 5);
    panel.innerHTML = `
      <div class="detail-title">
        <span>${escapeHtml(demand.id)} ${demand.ordem ? `| Ordem ${escapeHtml(demand.ordem)}` : "| Sem ordem SAP"}</span>
        <h3>${escapeHtml(demand.descricao)}</h3>
      </div>
      <div class="detail-grid">
        <div class="detail-item"><span>Status</span><strong>${statusChip(status)}</strong></div>
        <div class="detail-item"><span>Última atualização</span><strong>${formatDateTime(demand.dataUltimaAtualizacao)}</strong></div>

        <div class="detail-item"><span>Gerência</span><strong>${escapeHtml(demand.gerencia || "-")}</strong></div>
        <div class="detail-item"><span>Supervisão</span><strong>${escapeHtml(demand.supervisao || "-")}</strong></div>

        <div class="detail-item"><span>Vencimento</span><strong>${formatDate(demand.vencimento)}</strong></div>
        <div class="detail-item"><span>Centro</span><strong>${escapeHtml(demand.centroTrabalho || "-")}</strong></div>
        <div class="detail-item"><span>Crítico</span><strong>${statusChip(demand.critico || "Não informado")}</strong></div>

        <div class="detail-item"><span>Planejador Curto</span><strong>${escapeHtml(demand.planejadorCurto || "-")}</strong></div>
        <div class="detail-item"><span>Planejador OM</span><strong>${escapeHtml(demand.planejadorOM || "-")}</strong></div>
        <div class="detail-item"><span>Programador</span><strong>${escapeHtml(demand.programador || "-")}</strong></div>

        <div class="detail-item"><span>Local</span><strong>${escapeHtml(demand.localInstalacao || "-")}</strong></div>

        <div class="detail-item"><span>Tolerância Mín.</span><strong>${formatDate(demand.toleranciaMin)}</strong></div>
        <div class="detail-item"><span>Tolerância Máx.</span><strong>${formatDate(demand.toleranciaMax)}</strong></div>

        <div class="detail-item"><span>Planejada</span><strong>${formatDate(demand.dataPlanejada)}</strong></div>
        <div class="detail-item"><span>Replanejada</span><strong>${formatDate(demand.dataReplanejadaAtual)}</strong></div>
      </div>
      ${
        pending.length
          ? `<div class="pending-box"><strong>Pendências</strong>${pending.map((item) => `<span>${escapeHtml(item)}</span>`).join("")}</div>`
          : ""
      }
      <div class="detail-actions">
        ${actionButton("planejar", demand.id, !canPlan() || !allowed.planejar, true)}
        ${actionButton("replanejar", demand.id, !canReplan() || !allowed.replanejar, true)}
        ${actionButton("realizado", demand.id, !canRealizar() || !allowed.realizado, true)}
        ${actionButton("historico", demand.id, false, true)}
      </div>
      <div class="timeline">
        <h3>Histórico recente</h3>
        ${
          timeline.length
            ? timeline
                .map(
                  (item) => `
                    <div class="timeline-item">
                      <strong>${escapeHtml(item.type)} | ${escapeHtml(item.title)}</strong>
                      <span>${formatDateTime(item.date)} | ${escapeHtml(item.detail || "-")}</span>
                    </div>
                  `,
                )
                .join("")
            : '<span class="muted">Sem histórico operacional registrado.</span>'
        }
      </div>
    `;
  }

  function optionsMarkup(values, selected = "") {
    return values
      .map(
        (value) =>
          `<option value="${escapeHtml(value)}" ${value === selected ? "selected" : ""}>${escapeHtml(value)}</option>`,
      )
      .join("");
  }

  function bindDependentSelects() {
    const motivo = $('[name="motivo"]');
    const justificativa = $('[name="justificativa"]');
    if (motivo && justificativa) {
      const update = () => {
        const selected =
          justificativa.value || justificativa.dataset.selected || "";
        justificativa.innerHTML = optionsMarkup(
          childConfigNames("justificativas", "motivos", motivo.value),
          selected,
        );
      };
      motivo.addEventListener("change", update);
      update();
    }
    const perfil = $('[name="motivoPerda"]');
    const justificativaPerda = $('[name="justificativaPerda"]');
    if (perfil && justificativaPerda) {
      const update = () => {
        const selected =
          justificativaPerda.value || justificativaPerda.dataset.selected || "";
        justificativaPerda.innerHTML = `<option value=""></option>${optionsMarkup(childConfigNames("justificativasPerda", "perfisPerda", perfil.value), selected)}`;
      };
      perfil.addEventListener("change", update);
      update();
    }
  }

  function openAction(action, demandId) {
    const demand = demandById(demandId);
    if (!demand) return;
    const permissionByAction = {
      planejar: canPlan(),
      replanejar: canReplan(),
      realizado: canRealizar(),
      historico: true,
    };
    if (action !== "historico" && !permissionByAction[action]) {
      showToast("Perfil sem permissão para alterar registros.", "error");
      return;
    }
    const allowed = allowedActionsFor(demand);
    if (action !== "historico" && !allowed[action]) {
      showToast("Ação bloqueada para o status atual da demanda.", "error");
      return;
    }
    state.actionContext = { action, demandId };
    $("#modalDemandId").textContent =
      `${demand.id} | ${demand.ordem || "Sem ordem SAP"}`;

    if (action === "planejar") {
      $("#modalTitle").textContent = "Planejar Demanda";
      $("#modalBody").innerHTML = `
        <div class="modal-grid">
          <label>Data planejada<input name="dataPlanejada" type="date" value="${dateText(demand.dataPlanejada) || todayText()}" required /></label>
          <label>Responsável<input name="responsavel" value="${escapeHtml(demand.usuarioResponsavel || state.currentUser.email)}" /></label>
          <label class="span-2">Comentário<textarea name="comentario" rows="4">${escapeHtml(demand.comentario || "")}</textarea></label>
        </div>
      `;
      $("#modalSave").classList.remove("hidden");
    }

    if (action === "replanejar") {
      $("#modalTitle").textContent = "Replanejar Demanda";
      $("#modalBody").innerHTML = `
        <div class="modal-grid">
          <label>Data anterior<input value="${formatDate(demand.dataReplanejadaAtual || demand.dataPlanejada || demand.vencimento)}" disabled /></label>
          <label>Nova data<input name="novaData" type="date" value="${dateText(demand.dataReplanejadaAtual) || todayText()}" required /></label>
          <label>Motivo<select name="motivo" required>${optionsMarkup(configNames("motivos"))}</select></label>
          <label>Justificativa<select name="justificativa" required></select></label>
          <label class="span-2">Comentário<textarea name="comentario" rows="4">${escapeHtml(demand.comentario || "")}</textarea></label>
        </div>
      `;
      $("#modalSave").classList.remove("hidden");
    }

    if (action === "realizado") {
      $("#modalTitle").textContent = "Registrar Realizado/Perda";

      $("#modalBody").innerHTML = `
    <div class="modal-grid">
      <label>
        Data realizada
        <input
          name="dataRealizada"
          type="date"
          value="${dateText(demand.dataRealizada) || todayText()}"
        />
      </label>

      <label>
        Perda
        <select name="perda">
          <option value="nao" ${!demand.perda ? "selected" : ""}>Não</option>
          <option value="sim" ${demand.perda ? "selected" : ""}>Sim</option>
        </select>
      </label>

      <label>
        Perfil da perda
        <select name="motivoPerda">
          <option value=""></option>
          ${optionsMarkup(configNames("perfisPerda"), demand.motivoPerda)}
        </select>
      </label>

      <label>
        Justificativa perda
        <select
          name="justificativaPerda"
          data-selected="${escapeHtml(demand.justificativaPerda || "")}"
        ></select>
      </label>

      <label class="span-2">
        Comentário
        <textarea name="comentario" rows="4">${escapeHtml(demand.comentario || "")}</textarea>
      </label>

      <label class="span-2">
        Evidência
        <input
          name="evidencia"
          placeholder="URL ou referência do anexo no SharePoint"
        />
      </label>
    </div>
  `;

      $("#modalSave").classList.remove("hidden");
    }
    if (action === "historico") {
      $("#modalTitle").textContent = "Histórico da Demanda";
      const timeline = historiesFor(demand.id);
      $("#modalBody").innerHTML = `
        <div class="timeline">
          ${
            timeline.length
              ? timeline
                  .map(
                    (item) => `
                    <div class="timeline-item">
                      <strong>${escapeHtml(item.type)} | ${escapeHtml(item.title)}</strong>
                      <span>${formatDateTime(item.date)} | ${escapeHtml(item.detail || "-")}</span>
                    </div>
                  `,
                  )
                  .join("")
              : '<span class="muted">Sem histórico operacional registrado.</span>'
          }
        </div>
      `;
      $("#modalSave").classList.add("hidden");
    }

    const dialog = $("#actionDialog");
    bindDependentSelects();
    if (dialog.showModal) dialog.showModal();
    else dialog.setAttribute("open", "open");
  }

  function openReplanNotificationDialog(notification) {
    const demand = demandById(notification.demandId);

    if (!demand) {
      showToast("Demanda não encontrada.", "error");
      return;
    }

    const history = (state.db.historicoReplanejamento || []).find(
      (item) => item.id === notification.historyId,
    );

    if (!history) {
      showToast("Histórico de replanejamento não encontrado.", "error");
      return;
    }

    state.actionContext = {
      action: "regularizarReplanejamento",
      demandId: demand.id,
      historyId: history.id,
    };

    $("#modalDemandId").textContent =
      `${demand.id} | ${demand.ordem || "Sem ordem SAP"}`;
    $("#modalTitle").textContent = "Regularizar Replanejamento";

    $("#modalBody").innerHTML = `
    <div class="pending-box">
      <strong>Pendência identificada</strong>
      <span>${escapeHtml(notification.message)}</span>
    </div>

    <div class="modal-grid">
      <label>
        Data anterior
        <input value="${formatDate(history.dataAnterior)}" disabled />
      </label>

      <label>
        Nova data
        <input value="${formatDate(history.novaData)}" disabled />
      </label>

      <label>
        Motivo
        <select name="motivo" required>
          <option value=""></option>
          ${optionsMarkup(configNames("motivos"), history.motivo)}
        </select>
      </label>

      <label>
        Justificativa
        <select
          name="justificativa"
          data-selected="${escapeHtml(history.justificativa || "")}"
          required
        ></select>
      </label>

      <label class="span-2">
        Comentário
        <textarea name="comentario" rows="4">${escapeHtml(history.comentario || demand.comentario || "")}</textarea>
      </label>
    </div>
  `;

    $("#modalSave").classList.remove("hidden");

    const dialog = $("#actionDialog");
    bindDependentSelects();

    if (dialog.showModal) dialog.showModal();
    else dialog.setAttribute("open", "open");
  }

  function openLossNotificationDialog(notification) {
    const demand = demandById(notification.demandId);

    if (!demand) {
      showToast("Demanda não encontrada.", "error");
      return;
    }

    const history = latestRealizadoPerdaHistory(demand.id);

    state.actionContext = {
      action: "regularizarPerda",
      demandId: demand.id,
      historyId: history?.id || "",
    };

    $("#modalDemandId").textContent =
      `${demand.id} | ${demand.ordem || "Sem ordem SAP"}`;
    $("#modalTitle").textContent = "Regularizar Perda";

    $("#modalBody").innerHTML = `
    <div class="pending-box">
      <strong>Pendência identificada</strong>
      <span>${escapeHtml(notification.message)}</span>
    </div>

    <div class="modal-grid">
      <label>
        Data realizada
        <input
          name="dataRealizada"
          type="date"
          value="${dateText(demand.dataRealizada) || ""}"
        />
      </label>

      <label>
        Classe de prazo
        <input value="${escapeHtml(dueClassOf(demand) || "-")}" disabled />
      </label>

      <label>
        Perda
        <select name="perda" required>
          <option value="sim" selected>Sim</option>
        </select>
      </label>

      <label>
        Perfil da perda
        <select name="motivoPerda" required>
          <option value=""></option>
          ${optionsMarkup(configNames("perfisPerda"), demand.motivoPerda)}
        </select>
      </label>

      <label class="span-2">
        Justificativa perda
        <select
          name="justificativaPerda"
          data-selected="${escapeHtml(demand.justificativaPerda || "")}"
          required
        ></select>
      </label>

      <label class="span-2">
        Comentário
        <textarea name="comentario" rows="4">${escapeHtml(demand.comentario || history?.comentario || "")}</textarea>
      </label>

      <label class="span-2">
        Evidência
        <input name="evidencia" placeholder="URL ou referência do anexo no SharePoint" />
      </label>
    </div>
  `;

    $("#modalSave").classList.remove("hidden");

    const dialog = $("#actionDialog");
    bindDependentSelects();

    if (dialog.showModal) dialog.showModal();
    else dialog.setAttribute("open", "open");
  }

  async function saveAction() {
    const context = state.actionContext;
    if (!context) return;
    const demand = demandById(context.demandId);
    if (!demand) return;
    const form = new FormData($("#actionForm"));
    const userEmail = state.currentUser.email;

    if (context.action === "regularizarReplanejamento") {
      const motivo = form.get("motivo") || "";
      const justificativa = form.get("justificativa") || "";

      if (!motivo || !justificativa) {
        showToast("Informe motivo e justificativa do replanejamento.", "error");
        return;
      }

      await state.repo.updateReplanHistory(context.historyId, {
        demandaId: demand.id,
        ordem: demand.ordem,
        motivo,
        motivoChave: configKeyByName("motivos", motivo),
        justificativa,
        justificativaChave: configKeyByName("justificativas", justificativa),
        comentario: form.get("comentario") || "",
        usuario: userEmail,
        origemAlteracao: "NOTIFICACAO_OPERACIONAL",
      });

      demand.comentario = form.get("comentario") || demand.comentario || "";
      Object.assign(demand, prepareDemandForSave(demand));

      await state.repo.upsertDemanda(demand);

      await state.repo.addLog?.({
        usuario: userEmail,
        acao: "Regularização Replanejamento",
        lista: "historico_replanejamento",
        referencia: demand.id,
        detalhe: `${motivo} | ${justificativa}`,
        modulo: "NOTIFICACOES",
        status: "SUCESSO",
      });

      $("#actionDialog").close();
      clearNotificationsCache();
      await refreshAll();
      showToast("Replanejamento regularizado com sucesso.", "success");
      return;
    }

    if (context.action === "regularizarPerda") {
      const motivoPerda = form.get("motivoPerda") || "";
      const justificativaPerda = form.get("justificativaPerda") || "";

      if (!motivoPerda || !justificativaPerda) {
        showToast("Informe perfil e justificativa de perda.", "error");
        return;
      }

      demand.perda = true;
      demand.dataRealizada =
        form.get("dataRealizada") || demand.dataRealizada || "";
      demand.motivoPerda = motivoPerda;
      demand.justificativaPerda = justificativaPerda;
      demand.comentario = form.get("comentario") || demand.comentario || "";
      demand.usuarioResponsavel = userEmail;

      Object.assign(demand, prepareDemandForSave(demand));

      await state.repo.upsertDemanda(demand);

      await state.repo.updateRealizadoPerdaHistory(context.historyId, {
        demandaId: demand.id,
        ordem: demand.ordem,
        dataRealizada: demand.dataRealizada,
        perda: true,
        motivoPerda,
        motivoPerdaChave: configKeyByName("perfisPerda", motivoPerda),
        justificativaPerda,
        justificativaPerdaChave: configKeyByName(
          "justificativasPerda",
          justificativaPerda,
        ),
        comentario: demand.comentario,
        evidencia: form.get("evidencia") || "",
        usuario: userEmail,
        origemAlteracao: "NOTIFICACAO_OPERACIONAL",
      });

      await state.repo.addLog?.({
        usuario: userEmail,
        acao: "Regularização Perda",
        lista: "historico_realizado_perdas",
        referencia: demand.id,
        detalhe: `${motivoPerda} | ${justificativaPerda}`,
        modulo: "NOTIFICACOES",
        status: "SUCESSO",
      });

      $("#actionDialog").close();
      clearNotificationsCache();
      await refreshAll();
      showToast("Perda regularizada com sucesso.", "success");
      return;
    }

    if (context.action === "planejar") {
      const previous = demand.dataPlanejada || "";
      const nextDate = form.get("dataPlanejada");
      if (!nextDate) {
        showToast("Informe a data planejada.", "error");
        return;
      }
      demand.dataPlanejada = nextDate;
      demand.usuarioResponsavel = form.get("responsavel") || userEmail;
      demand.comentario = form.get("comentario") || "";
      Object.assign(demand, prepareDemandForSave(demand));
      await state.repo.upsertDemanda(demand);
      await state.repo.addHistory("planejamento", {
        demandaId: demand.id,
        ordem: demand.ordem,
        dataAnterior: previous,
        novaData: nextDate,
        usuario: userEmail,
        comentario: demand.comentario,
      });
      await state.repo.addLog({
        usuario: userEmail,
        acao: "Planejamento",
        lista: "Controle_Demandas_Eletrovia",
        referencia: demand.id,
        detalhe: `Planejado para ${nextDate}`,
      });
    }

    if (context.action === "replanejar") {
      const previous =
        demand.dataReplanejadaAtual ||
        demand.dataPlanejada ||
        demand.vencimento;
      const nextDate = form.get("novaData");
      if (!nextDate || !form.get("motivo") || !form.get("justificativa")) {
        showToast(
          "Replanejamento exige nova data, motivo e justificativa.",
          "error",
        );
        return;
      }
      demand.dataReplanejadaAtual = nextDate;
      demand.quantidadeReplanejamentos =
        Number(demand.quantidadeReplanejamentos || 0) + 1;
      demand.comentario = form.get("comentario") || "";
      demand.usuarioResponsavel = userEmail;
      Object.assign(demand, prepareDemandForSave(demand));
      await state.repo.upsertDemanda(demand);
      await state.repo.addHistory("replanejamento", {
        demandaId: demand.id,
        ordem: demand.ordem,
        motivo: form.get("motivo"),
        motivoChave: configKeyByName("motivos", form.get("motivo")),
        justificativa: form.get("justificativa"),
        justificativaChave: configKeyByName(
          "justificativas",
          form.get("justificativa"),
        ),
        dataAnterior: previous,
        novaData: nextDate,
        usuario: userEmail,
        quantidadeReplanejamentos: demand.quantidadeReplanejamentos,
        comentario: demand.comentario,
      });
      await state.repo.addLog({
        usuario: userEmail,
        acao: "Replanejamento",
        lista: "Controle_Demandas_Eletrovia",
        referencia: demand.id,
        detalhe: `${previous} para ${nextDate}`,
      });
    }

    if (context.action === "realizado") {
      const dataRealizadaInformada = form.get("dataRealizada") || "";
      const lossSelected = form.get("perda") === "sim";
      const motivoPerda = form.get("motivoPerda") || "";
      const justificativaPerda = form.get("justificativaPerda") || "";

      if (!dataRealizadaInformada && !lossSelected) {
        showToast(
          "Informe a data realizada ou marque perda para registrar a ocorrência.",
          "error",
        );
        return;
      }

      if (lossSelected && (!motivoPerda || !justificativaPerda)) {
        showToast("Perda exige perfil e justificativa.", "error");
        return;
      }

      demand.dataRealizada = dataRealizadaInformada;
      demand.perda = lossSelected;
      demand.motivoPerda = lossSelected ? motivoPerda : "";
      demand.justificativaPerda = lossSelected ? justificativaPerda : "";
      demand.comentario = form.get("comentario") || "";
      demand.usuarioResponsavel = userEmail;

      Object.assign(demand, prepareDemandForSave(demand));

      await state.repo.upsertDemanda(demand);

      await state.repo.addHistory("realizadoPerda", {
        demandaId: demand.id,
        ordem: demand.ordem,
        dataRealizada: demand.dataRealizada,
        perda: demand.perda,
        motivoPerda: demand.motivoPerda,
        motivoPerdaChave: configKeyByName("perfisPerda", demand.motivoPerda),
        justificativaPerda: demand.justificativaPerda,
        justificativaPerdaChave: configKeyByName(
          "justificativasPerda",
          demand.justificativaPerda,
        ),
        comentario: demand.comentario,
        evidencia: form.get("evidencia") || "",
        usuario: userEmail,
      });

      await state.repo.addLog({
        usuario: userEmail,
        acao: demand.perda ? "Perda" : "Realizado",
        lista: "Historico_Realizado_Perdas",
        referencia: demand.id,
        detalhe: demand.perda
          ? `${demand.motivoPerda} | ${demand.justificativaPerda}`
          : `Realizado em ${demand.dataRealizada}`,
      });
    }

    $("#actionDialog").close();
    await refreshAfterSave("Registro salvo com sucesso.");
  }

  function renderCurrentView() {
    syncNavigation(state.currentView);
    if (state.currentView === "carteira") renderCarteira();
    if (state.currentView === "lote") renderBatch();
    if (state.currentView === "futuras") renderFutureDemandas();
    if (state.currentView === "qualidade") renderQuality();
    if (state.currentView === "qualidade-local") renderQualityByLocal();
    if (state.currentView === "qualidade-centros") renderQualityCenters();
    if (state.currentView === "qualidade-divergencias")
      renderQualityDivergences();
    if (state.currentView === "notificacoes") renderNotifications();
    if (state.currentView === "historico-carteira") renderPortfolioHistory();
    if (state.currentView === "indicadores") {
      renderIndicatorFilterVisibility();
      if (state.indicatorFiltersVisible && !state.indicatorFiltersReady) {
        collectIndicatorFilters();
        buildIndicatorFilterOptions();
      }
      renderIndicators();
    }
    if (state.currentView === "administracao") renderAdmin();
    if (state.currentView === "logs") renderLogs();
    if (state.currentView === "saude-integracao") renderIntegrationHealth();
    applyPermissions();
  }

  async function refreshAfterSave(message = "Registro salvo com sucesso.") {
    await refreshAll();
    showToast(message, "success");
  }

  async function refreshAll() {
    await loadDatabase();
    await autoSyncRealizadosFromSharePoint();
    hydrateStaticUi();
    renderCurrentView();
    const time = new Date().toLocaleTimeString("pt-BR", {
      hour: "2-digit",
      minute: "2-digit",
    });
    $("#lastSync").textContent = `Sincronizado ${time}`;
    $("#lastUpdateSide").textContent = state.lastDataUpdateAt
      ? formatDateTime(state.lastDataUpdateAt)
      : "-";
  }

  function switchView(view) {
    if (navGroupForView(view) === "administracao" && !canAdmin()) {
      showToast(
        "Administração disponível somente para Administrador.",
        "error",
      );
      return;
    }
    if (view === "lote" && !canBatch()) {
      showToast("Perfil sem permissao para carga em lote.", "error");
      return;
    }
    state.currentView = view;
    const groupKey = navGroupForView(view);
    Object.keys(state.navGroups).forEach((key) => {
      state.navGroups[key] = key === groupKey;
    });
    syncNavigation(view);
    $$("[data-view-panel]").forEach((panel) =>
      panel.classList.toggle("is-active", panel.dataset.viewPanel === view),
    );
    renderCurrentView();
  }

  function renderBatch() {
    const summary = $("#validationSummary");
    const groups = $("#validationGroups");
    const saveValidButton = $("#saveValidBatch");
    const saveConfirmedButton = $("#saveConfirmedBatch");
    const confirmWarningsLabel = $("#confirmWarnings")?.closest(
      ".confirm-alerts",
    );

    const updateSaveActions = () => {
      const hasRows = state.batch.rows.length > 0;
      const hasValid = state.batch.valid.length > 0;
      const hasWarnings = state.batch.warnings.length > 0;

      saveValidButton?.classList.toggle(
        "hidden",
        !hasRows || !hasValid || hasWarnings,
      );
      saveConfirmedButton?.classList.toggle(
        "hidden",
        !hasRows || !hasValid || !hasWarnings,
      );
      confirmWarningsLabel?.classList.toggle(
        "hidden",
        !hasRows || !hasValid || !hasWarnings,
      );

      if (!hasWarnings && $("#confirmWarnings")) {
        $("#confirmWarnings").checked = false;
      }
    };
    if (!state.batch.rows.length) {
      summary.innerHTML = "";
      groups.innerHTML =
        '<div class="empty-detail"><strong>Nenhum arquivo validado</strong><span>Selecione um arquivo para iniciar a validação.</span></div>';
      updateSaveActions();
      return;
    }

    summary.innerHTML = [
      ["Válidos", state.batch.valid.length, "status-realizado"],
      ["Alertas", state.batch.warnings.length, "status-planejar"],
      ["Erros", state.batch.errors.length, "status-perda"],
    ]
      .map(
        ([label, count, klass]) => `
        <div class="validation-count">
          <span class="status-chip ${klass}">${label}</span>
          <strong>${count}</strong>
        </div>
      `,
      )
      .join("");

    const renderGroup = (title, rows, type) => `
      <section class="validation-group">
        <h3>${title}</h3>
        <div class="validation-list">
          ${
            rows.length
              ? rows
                  .slice(0, 30)
                  .map(
                    (item) => `
                <div class="validation-row ${type}">
                  <strong>Linha ${item.line} | ${escapeHtml(item.record.ordem || item.record.id || item.record.descricao || "-")}</strong>
                  <div>${escapeHtml(item.message)}</div>
                </div>
              `,
                  )
                  .join("")
              : '<span class="muted">Nenhum registro.</span>'
          }
        </div>
      </section>
    `;
    groups.innerHTML =
      renderGroup("Registros Válidos", state.batch.valid, "valid") +
      renderGroup("Registros com Alerta", state.batch.warnings, "warning") +
      renderGroup("Registros com Erro", state.batch.errors, "error");
    updateSaveActions();
  }

  function renderIndicatorFilterVisibility() {
    const panel = $("#indicatorFilterPanel");
    const button = $("#toggleIndicatorFilters");
    if (!panel || !button) return;

    panel.classList.toggle("hidden", !state.indicatorFiltersVisible);
    button.textContent = state.indicatorFiltersVisible
      ? "Ocultar filtros"
      : "Mostrar filtros";
  }

  function renderQualityLocalFilterVisibility() {
    const panel = $("#qualityLocalFilterPanel");
    const button = $("#toggleQualityLocalFilters");
    if (!panel || !button) return;

    panel.classList.toggle("hidden", !state.quality.localFiltersVisible);
    button.textContent = state.quality.localFiltersVisible
      ? "Ocultar filtros"
      : "Mostrar filtros";
  }

  function parseCsv(text) {
    const rows = [];
    let current = "";
    let row = [];
    let quoted = false;
    for (let index = 0; index < text.length; index += 1) {
      const char = text[index];
      const next = text[index + 1];
      if (char === '"' && quoted && next === '"') {
        current += '"';
        index += 1;
      } else if (char === '"') {
        quoted = !quoted;
      } else if ((char === ";" || char === ",") && !quoted) {
        row.push(current);
        current = "";
      } else if ((char === "\n" || char === "\r") && !quoted) {
        if (current || row.length) {
          row.push(current);
          rows.push(row);
        }
        current = "";
        row = [];
        if (char === "\r" && next === "\n") index += 1;
      } else {
        current += char;
      }
    }
    if (current || row.length) {
      row.push(current);
      rows.push(row);
    }
    const headers = rows.shift() || [];
    return rows
      .filter((cells) => cells.some((cell) => String(cell).trim()))
      .map((cells) =>
        headers.reduce((acc, header, index) => {
          acc[header] = cells[index] || "";
          return acc;
        }, {}),
      );
  }

  function normalizeHeader(header) {
    const text = normalizeText(header).replace(/[^A-Z0-9]/g, "");
    const map = {
      ORDEM: "ordem",
      ORDEMSAP: "ordem",
      ID: "id",
      IDDEMANDA: "id",
      IDDEMANDACONTROLE: "id",
      DESCRICAO: "descricao",
      DESCRIO: "descricao",
      CENTROTRABALHO: "centroTrabalho",
      CT: "centroTrabalho",
      LOCALINSTALACAO: "localInstalacao",
      LOCALINSTALAO: "localInstalacao",
      COMPETENCIA: "competencia",
      VENCIMENTO: "vencimento",
      DATAPLANEJADA: "dataPlanejada",
      DATAREPLANEJADA: "dataReplanejadaAtual",
      DATAREALIZADA: "dataRealizada",
      DATAFIMREAL: "dataRealizada",
      DATACONCLUSAO: "dataRealizada",
      DATACONCLUSÃO: "dataRealizada",
      REALIZADOEM: "dataRealizada",
      STATUSREALIZADO: "statusRealizado",
      MOTIVO: "motivo",
      JUSTIFICATIVA: "justificativa",
      PERDA: "perda",
      MOTIVOPERDA: "motivoPerda",
      JUSTIFICATIVAPERDA: "justificativaPerda",
      COMENTARIO: "comentario",
      TIPOOM: "tipoOM",
      PRIORIDADE: "prioridade",
      RESPONSAVEL: "usuarioResponsavel",
    };
    return map[text] || header;
  }

  function normalizeDateInput(value) {
    if (typeof value === "number" && global.XLSX) {
      const parsed = global.XLSX.SSF.parse_date_code(value);
      if (parsed)
        return `${parsed.y}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
    }
    const date = toDate(value);
    return dateText(date);
  }

  function normalizeBatchRecord(row) {
    const normalized = {};
    Object.entries(row).forEach(([key, value]) => {
      normalized[normalizeHeader(key)] =
        typeof value === "string" ? value.trim() : value;
    });
    [
      "vencimento",
      "dataPlanejada",
      "dataReplanejadaAtual",
      "dataRealizada",
    ].forEach((key) => {
      if (normalized[key])
        normalized[key] = normalizeDateInput(normalized[key]);
    });
    if (normalized.competencia)
      normalized.competencia = normalizeCompetencia(normalized.competencia);
    if (normalized.prioridade)
      normalized.prioridade = normalizePrioridade(normalized.prioridade);
    if (normalized.perda) {
      normalized.perda = ["SIM", "S", "TRUE", "1"].includes(
        normalizeText(normalized.perda),
      );
    }
    normalized.origem = "Carga em Lote";
    if (normalized.ordem) normalized.ordem = String(normalized.ordem).trim();
    if (normalized.id) normalized.id = String(normalized.id).trim();
    return normalized;
  }

  function isBlockedForBatchUpdate(demand) {
    if (!demand) return false;

    const status = primaryStatusOf(demand);

    return status === "Realizado" || status === "Cancelado";
  }

  function blockedBatchMessage(demand, record) {
    const ordem = record?.ordem || demand?.ordem || "";
    const id = record?.id || demand?.id || "";
    const status = primaryStatusOf(demand);

    if (status === "Cancelado") {
      return `A ordem ${ordem || id} está cancelada/encerrada e não pode ser alterada por carga em lote.`;
    }

    if (status === "Realizado") {
      return `A ordem ${ordem || id} está realizada/encerrada e não pode ser alterada por carga em lote.`;
    }

    return `A ordem ${ordem || id} está bloqueada para alteração por carga em lote.`;
  }

  function validateBatchRows(rows) {
    const valid = [];
    const warnings = [];
    const errors = [];

    rows.forEach((row, index) => {
      const record = normalizeBatchRecord(row);
      const messages = [];
      const alerts = [];
      const existing = findDemandForBatch(record);

      if (existing && isBlockedForBatchUpdate(existing)) {
        messages.push(blockedBatchMessage(existing, record));
      }

      if (!record.ordem && !record.id) {
        messages.push("Informe Ordem SAP ou ID_Demanda_Controle.");
      }
      if ((record.ordem || record.id) && !existing) {
        alerts.push(
          "Demanda nao encontrada na carteira atual. Linha nao sera gravada.",
        );
      }
      if (record.dataPlanejada && !toDate(record.dataPlanejada))
        messages.push("Data planejada inválida.");
      if (record.dataReplanejadaAtual && !toDate(record.dataReplanejadaAtual))
        messages.push("Data replanejada inválida.");
      if (record.dataRealizada && !toDate(record.dataRealizada))
        messages.push("Data realizada inválida.");
      if (record.perda && (!record.motivoPerda || !record.justificativaPerda)) {
        messages.push("Perda exige motivo e justificativa.");
      }
      if (!record.ordem && record.id)
        alerts.push(
          "Demanda sem ordem SAP sera atualizada pelo ID_Demanda_Controle.",
        );

      const item = {
        line: index + 2,
        record,
        message:
          messages.concat(alerts).join(" ") || "Registro pronto para gravação.",
      };
      if (messages.length) errors.push(item);
      else if (alerts.length) warnings.push(item);
      else valid.push(item);
    });

    state.batch = { ...state.batch, rows, valid, warnings, errors };
  }

  function buildBatchPartialUpdate(record) {
    const ignoredKeys = new Set(["id", "ordem", "origem"]);

    const partial = {};

    Object.entries(record || {}).forEach(([key, value]) => {
      if (ignoredKeys.has(key)) return;

      if (value === null || value === undefined) return;

      if (typeof value === "string" && value.trim() === "") return;

      partial[key] = value;
    });

    return partial;
  }

  function emptyBatchHistoryEntries() {
    return {
      planejamento: [],
      replanejamento: [],
      realizadoPerda: [],
    };
  }

  function countBatchHistoryEntries(historyEntries) {
    return (
      (historyEntries?.planejamento?.length || 0) +
      (historyEntries?.replanejamento?.length || 0) +
      (historyEntries?.realizadoPerda?.length || 0)
    );
  }

  function mergeBatchHistoryEntries(target, source) {
    target.planejamento.push(...(source.planejamento || []));
    target.replanejamento.push(...(source.replanejamento || []));
    target.realizadoPerda.push(...(source.realizadoPerda || []));
    return target;
  }

  function buildBatchHistoryEntries(record, previousDemand = {}, savedDemand) {
    const entries = emptyBatchHistoryEntries();
    const userEmail = state.currentUser?.email || "";
    const now = new Date().toISOString();
    const comentario =
      record.comentario ||
      savedDemand.comentario ||
      previousDemand.comentario ||
      "";

    if (record.dataPlanejada) {
      entries.planejamento.push({
        demandaId: savedDemand.id,
        ordem: savedDemand.ordem,
        dataAnterior: previousDemand.dataPlanejada || "",
        novaData: record.dataPlanejada,
        usuario: userEmail,
        comentario,
        dataHora: now,
        tipoAlteracao: "CARGA_LOTE_PLANEJAMENTO",
        origemAlteracao: "CARGA_LOTE",
      });
    }

    if (record.dataReplanejadaAtual) {
      const motivo = record.motivo || "";
      const justificativa = record.justificativa || "";

      entries.replanejamento.push({
        demandaId: savedDemand.id,
        ordem: savedDemand.ordem,
        motivo,
        motivoChave: configKeyByName("motivos", motivo),
        justificativa,
        justificativaChave: configKeyByName("justificativas", justificativa),
        dataAnterior:
          previousDemand.dataReplanejadaAtual ||
          previousDemand.dataPlanejada ||
          previousDemand.vencimento ||
          "",
        novaData: record.dataReplanejadaAtual,
        usuario: userEmail,
        quantidadeReplanejamentos: savedDemand.quantidadeReplanejamentos || 0,
        comentario,
        dataHora: now,
        origemAlteracao: "CARGA_LOTE",
      });
    }

    if (record.dataRealizada || record.perda === true) {
      entries.realizadoPerda.push({
        demandaId: savedDemand.id,
        ordem: savedDemand.ordem,
        dataRealizada: record.dataRealizada || savedDemand.dataRealizada || "",
        perda: record.perda === true || savedDemand.perda === true,
        motivoPerda: record.motivoPerda || savedDemand.motivoPerda || "",
        motivoPerdaChave: configKeyByName(
          "perfisPerda",
          record.motivoPerda || savedDemand.motivoPerda || "",
        ),
        justificativaPerda:
          record.justificativaPerda || savedDemand.justificativaPerda || "",
        justificativaPerdaChave: configKeyByName(
          "justificativasPerda",
          record.justificativaPerda || savedDemand.justificativaPerda || "",
        ),
        comentario,
        evidencia: record.evidencia || "",
        usuario: userEmail,
        dataHora: now,
        origemAlteracao: "CARGA_LOTE",
      });
    }

    return entries;
  }

  async function saveBatchHistoryEntries(historyEntries) {
    const total = countBatchHistoryEntries(historyEntries);
    if (!total) return historyEntries || emptyBatchHistoryEntries();

    if (state.repo.bulkAddHistories) {
      return state.repo.bulkAddHistories(historyEntries);
    }

    for (const entry of historyEntries.planejamento || []) {
      await state.repo.addHistory("planejamento", entry);
    }

    for (const entry of historyEntries.replanejamento || []) {
      await state.repo.addHistory("replanejamento", entry);
    }

    for (const entry of historyEntries.realizadoPerda || []) {
      await state.repo.addHistory("realizadoPerda", entry);
    }

    return historyEntries;
  }

  function appendBatchHistoriesLocally(historyEntries) {
    if (!state.db || !historyEntries) return;

    const now = new Date().toISOString();
    const localId = (type, index) =>
      `LOCAL-CARGA-${type}-${Date.now()}-${index}`;

    state.db.historicoPlanejamento = [
      ...(state.db.historicoPlanejamento || []),
      ...(historyEntries.planejamento || []).map((entry, index) => ({
        id: localId("PLAN", index),
        ...entry,
        dataHora: entry.dataHora || now,
      })),
    ];

    state.db.historicoReplanejamento = [
      ...(state.db.historicoReplanejamento || []),
      ...(historyEntries.replanejamento || []).map((entry, index) => ({
        id: localId("REPLAN", index),
        ...entry,
        dataHora: entry.dataHora || now,
      })),
    ];

    state.db.historicoRealizadoPerdas = [
      ...(state.db.historicoRealizadoPerdas || []),
      ...(historyEntries.realizadoPerda || []).map((entry, index) => ({
        id: localId("REAL", index),
        ...entry,
        dataHora: entry.dataHora || now,
      })),
    ];
  }

  function applySavedBatchLocally(records) {
    if (!Array.isArray(records) || !records.length || !state.db?.demandas) {
      return;
    }

    const byId = new Map(
      state.db.demandas.map((item) => [String(item.id || ""), item]),
    );

    const byOrdem = new Map(
      state.db.demandas
        .filter((item) => String(item.ordem || "").trim())
        .map((item) => [String(item.ordem || "").trim(), item]),
    );

    records.forEach((record) => {
      const id = String(record.id || "").trim();
      const ordem = String(record.ordem || "").trim();

      const existing = (id && byId.get(id)) || (ordem && byOrdem.get(ordem));

      const normalized = normalizeDemandRecord(
        prepareDemandForSave({
          ...(existing || {}),
          ...record,
        }),
      );

      if (existing) {
        Object.assign(existing, normalized);
        return;
      }

      state.db.demandas.unshift(normalized);
    });

    state.lastDataUpdateAt = new Date().toISOString();
    state.quality.issuesCache = [];
    state.quality.filteredCache = [];
    clearNotificationsCache?.();
    clearBatchLookup();
  }

  function finishBatchWithoutFullRefresh(records, resumoCarga, historyEntries) {
    applySavedBatchLocally(records);
    appendBatchHistoriesLocally(historyEntries);

    state.batch = {
      rows: [],
      valid: [],
      warnings: [],
      errors: [],
      fileName: "",
    };

    const batchFile = $("#batchFile");
    if (batchFile) {
      batchFile.value = "";
    }

    renderBatch();

    if (state.currentView === "carteira") {
      renderCarteira();
    }

    if (state.currentView === "notificacoes") {
      clearNotificationsCache();
      renderNotifications();
    }

    const time = new Date().toLocaleTimeString("pt-BR", {
      hour: "2-digit",
      minute: "2-digit",
    });

    const lastSync = $("#lastSync");
    if (lastSync) {
      lastSync.textContent = `Salvo ${time}`;
    }

    const lastUpdateSide = $("#lastUpdateSide");
    if (lastUpdateSide) {
      lastUpdateSide.textContent = formatDateTime(state.lastDataUpdateAt);
    }

    showBatchStatus(
      "success",
      "Carga em lote enviada com sucesso.",
      `${resumoCarga.processados} registro(s) salvos e ${resumoCarga.historicos || 0} histórico(s) gravado(s). Como foi uma carga grande, o sistema atualizou a memória local e não recarregou todo o banco para evitar travamento. Use Atualizar somente se precisar recarregar tudo.`,
    );

    showToast(
      `${resumoCarga.processados} registros salvos sem recarregar toda a base.`,
      "success",
    );
  }
  function getBatchLookupMaps() {
    if (state.batchLookup?.sourceLength === state.db?.demandas?.length) {
      return state.batchLookup;
    }

    const byId = new Map();
    const byOrdem = new Map();

    (state.db?.demandas || []).forEach((item) => {
      const id = String(item.id || "").trim();
      const ordem = String(item.ordem || "").trim();

      if (id && !byId.has(id)) {
        byId.set(id, item);
      }

      if (ordem && !byOrdem.has(ordem)) {
        byOrdem.set(ordem, item);
      }
    });

    state.batchLookup = {
      byId,
      byOrdem,
      sourceLength: state.db?.demandas?.length || 0,
    };

    return state.batchLookup;
  }

  function clearBatchLookup() {
    state.batchLookup = null;
  }

  function findDemandForBatch(record) {
    const lookup = getBatchLookupMaps();

    const id = String(record?.id || "").trim();
    const ordem = String(record?.ordem || "").trim();

    if (id && lookup.byId.has(id)) {
      return lookup.byId.get(id);
    }

    if (ordem && lookup.byOrdem.has(ordem)) {
      return lookup.byOrdem.get(ordem);
    }

    return null;
  }

  async function readBatchFile(file) {
    const extension = file.name.split(".").pop().toLowerCase();
    if (["xlsx", "xls"].includes(extension)) {
      if (!global.XLSX) {
        throw new Error(
          "Biblioteca XLSX indisponível. Use CSV ou habilite o CDN no ambiente SharePoint.",
        );
      }
      const buffer = await file.arrayBuffer();
      const workbook = global.XLSX.read(buffer, {
        type: "array",
        cellDates: true,
      });
      const sheetName = workbook.SheetNames[0];
      return global.XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
        defval: "",
      });
    }
    const text = await file.text();
    return parseCsv(text);
  }

  async function validateBatchFile() {
    const file = $("#batchFile").files[0];
    if (!file) {
      showToast("Selecione um arquivo para validar.", "error");
      return;
    }
    try {
      showBatchStatus(
        "loading",
        "Validando arquivo...",
        "Lendo o arquivo e conferindo as linhas antes da gravação.",
      );

      const rows = await readBatchFile(file);
      state.batch.fileName = file.name;
      validateBatchRows(rows);
      renderBatch();

      showBatchStatus(
        "success",
        "Arquivo validado com sucesso.",
        `${state.batch.valid.length} válidos, ${state.batch.warnings.length} alertas e ${state.batch.errors.length} erros encontrados.`,
      );

      showToast("Arquivo validado.", "success");
    } catch (error) {
      showBatchStatus("error", "Erro ao validar arquivo.", error.message);
      showToast(error.message, "error");
    }
  }

  async function saveBatch(includeWarnings) {
    if (!canBatch()) {
      showToast("Perfil sem permissão para salvar carga em lote.", "error");
      return;
    }

    const candidates = includeWarnings
      ? state.batch.valid.concat(state.batch.warnings)
      : state.batch.valid;

    if (
      includeWarnings &&
      state.batch.warnings.length &&
      !$("#confirmWarnings").checked
    ) {
      showToast("Confirme os registros com alerta antes de salvar.", "error");
      return;
    }

    if (!candidates.length) {
      showToast("Nenhum registro disponível para salvar.", "error");
      return;
    }

    showBatchStatus(
      "loading",
      "Enviando carga em lote...",
      `${candidates.length} registro(s) em processamento. Aguarde a gravação no Supabase.`,
    );

    const blockedInSave = [];
    const notFoundInSave = [];
    const preparedBatchRows = [];
    const historyEntries = emptyBatchHistoryEntries();

    candidates.forEach((item) => {
      const existing = findDemandForBatch(item.record);

      if (!existing) {
        notFoundInSave.push({
          ...item,
          message:
            "Demanda nao encontrada na carteira atual. Linha nao sera gravada.",
        });
        return;
      }

      if (isBlockedForBatchUpdate(existing)) {
        blockedInSave.push({
          ...item,
          message: blockedBatchMessage(existing, item.record),
        });
        return;
      }

      const partial = buildBatchPartialUpdate(item.record);

      if (
        partial.dataReplanejadaAtual &&
        partial.dataReplanejadaAtual !== existing.dataReplanejadaAtual
      ) {
        partial.quantidadeReplanejamentos =
          Number(existing.quantidadeReplanejamentos || 0) + 1;
      }

      const demand = prepareDemandForSave({
        ...existing,
        ...partial,

        id: existing.id,
        ordem: existing.ordem || item.record.ordem || "",

        origem: existing.origem || item.record.origem || "Carga em Lote",

        usuarioResponsavel:
          partial.usuarioResponsavel ||
          existing.usuarioResponsavel ||
          state.currentUser.email,
      });

      mergeBatchHistoryEntries(
        historyEntries,
        buildBatchHistoryEntries(item.record, existing, demand),
      );

      preparedBatchRows.push({ demand });
    });

    const records = preparedBatchRows.map((item) => item.demand);

    if (blockedInSave.length) {
      state.batch.errors = [
        ...state.batch.errors,
        ...blockedInSave.map((item) => ({
          ...item,
          status: "ERRO",
          acao: "BLOQUEADO_STATUS_FINAL",
        })),
      ];

      renderBatch();

      showToast(
        `${blockedInSave.length} registro(s) bloqueado(s) por status Realizado/Cancelado.`,
        "error",
      );
    }

    if (notFoundInSave.length) {
      const warningLines = new Set(
        state.batch.warnings.map((item) => item.line),
      );
      const newNotFoundWarnings = notFoundInSave.filter(
        (item) => !warningLines.has(item.line),
      );

      state.batch.warnings = [
        ...state.batch.warnings,
        ...newNotFoundWarnings.map((item) => ({
          ...item,
          status: "ALERTA",
          acao: "NAO_ENCONTRADO_CARTEIRA",
        })),
      ];

      renderBatch();

      showToast(
        `${notFoundInSave.length} registro(s) não encontrado(s) na carteira atual e não gravado(s).`,
        "warning",
      );
    }

    if (!records.length) {
      showBatchStatus(
        "error",
        "Nenhum registro foi preparado para gravação.",
        blockedInSave.length
          ? "Todas as linhas selecionadas estavam bloqueadas por status Realizado/Cancelado."
          : notFoundInSave.length
            ? "Todas as linhas selecionadas não foram encontradas na carteira atual."
            : "As linhas válidas não tinham Ordem SAP nem ID_Demanda_Controle.",
      );

      showToast("Nenhum registro foi preparado para gravação.", "error");
      return;
    }

    try {
      await state.repo.bulkUpsertDemandas(records);
      const savedHistoryEntries = await saveBatchHistoryEntries(historyEntries);
      const historyCount = countBatchHistoryEntries(historyEntries);

      const batchRun = await state.repo.createBatchRun?.({
        nomeArquivo: state.batch.fileName || "arquivo_sem_nome",
        tipoCarga: "ATUALIZACAO_DEMANDAS",
        usuario: state.currentUser?.nome || state.currentUser?.email || "",
        usuarioEmail: state.currentUser?.email || "",
        totalLinhas: state.batch.rows.length,
        linhasValidas: state.batch.valid.length,
        linhasAlerta: state.batch.warnings.length,
        linhasComErro: state.batch.errors.length,
        linhasProcessadas: records.length,
        status: state.batch.errors.length
          ? "PROCESSADO_COM_ERRO"
          : "PROCESSADO",
        detalheErro: state.batch.errors.length
          ? `${state.batch.errors.length} linhas com erro de validação. ${historyCount} histórico(s) gravado(s).`
          : "",
      });

      const loteId = batchRun?.lote_id || batchRun?.id;

      if (loteId) {
        const totalAuditItems =
          state.batch.valid.length +
          state.batch.warnings.length +
          state.batch.errors.length;

        if (totalAuditItems <= 500) {
          const auditItems = [
            ...state.batch.valid.map((item) => ({
              ...item,
              status: "VALIDO",
              acao: "UPSERT_DEMANDA",
            })),
            ...state.batch.warnings.map((item) => ({
              ...item,
              status: "ALERTA",
              acao: "UPSERT_DEMANDA_COM_ALERTA",
            })),
            ...state.batch.errors.map((item) => ({
              ...item,
              status: "ERRO",
              acao: "VALIDACAO_ERRO",
            })),
          ];

          await state.repo.addBatchItems?.(loteId, auditItems);
        } else {
          await state.repo.addLog?.({
            usuario: state.currentUser?.email || "",
            acao: "Carga em Lote - Auditoria resumida",
            lista: "cargas_lote_itens",
            referencia: loteId,
            detalhe: `Auditoria detalhada não gravada item a item para evitar travamento. Total de linhas: ${totalAuditItems}.`,
            modulo: "CARGA_LOTE",
            status: "SUCESSO",
          });
        }
      }

      await state.repo.addLog({
        usuario: state.currentUser.email,
        acao: "Carga em Lote",
        lista: "cargas_lote",
        referencia: `${records.length} registros`,
        detalhe: includeWarnings
          ? `Válidos e alertas confirmados | ${historyCount} histórico(s) gravado(s)`
          : `Somente válidos | ${historyCount} histórico(s) gravado(s)`,
        modulo: "CARGA_LOTE",
        status: "SUCESSO",
      });

      const resumoCarga = {
        processados: records.length,
        validos: state.batch.valid.length,
        alertas: includeWarnings ? state.batch.warnings.length : 0,
        erros: state.batch.errors.length,
        loteId: loteId || "",
        historicos: historyCount,
      };

      if (records.length > LARGE_BATCH_REFRESH_LIMIT) {
        finishBatchWithoutFullRefresh(
          records,
          resumoCarga,
          savedHistoryEntries,
        );
        return;
      }

      await refreshAll();

      state.batch = {
        rows: [],
        valid: [],
        warnings: [],
        errors: [],
        fileName: "",
      };

      const batchFile = $("#batchFile");
      if (batchFile) {
        batchFile.value = "";
      }

      renderBatch();

      showBatchStatus(
        "success",
        "Carga em lote enviada com sucesso.",
        `${resumoCarga.processados} registro(s) processado(s) e ${resumoCarga.historicos} histórico(s) gravado(s). ${
          resumoCarga.loteId ? `Lote: ${resumoCarga.loteId}. ` : ""
        }${resumoCarga.alertas ? `${resumoCarga.alertas} alerta(s) confirmado(s). ` : ""}${
          resumoCarga.erros
            ? `${resumoCarga.erros} linha(s) ficaram com erro de validação.`
            : ""
        }`,
      );

      showToast(`${records.length} registros salvos.`, "success");
    } catch (error) {
      console.error(error);

      showBatchStatus(
        "error",
        "Erro ao enviar carga em lote.",
        error.message || "Falha inesperada ao gravar registros no Supabase.",
      );

      showToast("Erro ao salvar carga em lote.", "error");
    }
  }

  async function syncRealizedRows(rows, sourceName, showResult = false) {
    const normalized = rows.map(normalizeBatchRecord);
    const updates = [];
    const unmatched = [];
    for (const record of normalized) {
      const ordem = String(record.ordem || "").trim();
      if (!ordem || !record.dataRealizada) {
        unmatched.push(record);
        continue;
      }
      const demand = state.db.demandas.find(
        (item) => String(item.ordem) === ordem,
      );
      if (!demand) {
        unmatched.push(record);
        continue;
      }
      const before = demand.dataRealizada || "";
      demand.dataRealizada = record.dataRealizada;
      demand.perda = Boolean(record.perda || demand.perda);
      demand.motivoPerda = record.motivoPerda || demand.motivoPerda || "";
      demand.justificativaPerda =
        record.justificativaPerda || demand.justificativaPerda || "";
      demand.comentario =
        record.comentario ||
        demand.comentario ||
        "Realizado sincronizado pela base SAP BO.";
      demand.usuarioResponsavel = state.currentUser.email;
      Object.assign(demand, prepareDemandForSave(demand));
      updates.push({ demand, before });
    }
    if (!updates.length) return { updates: 0, unmatched: unmatched.length };
    await state.repo.bulkUpsertDemandas(updates.map((item) => item.demand));
    for (const { demand, before } of updates) {
      if (before !== demand.dataRealizada) {
        await state.repo.addHistory("realizadoPerda", {
          demandaId: demand.id,
          dataRealizada: demand.dataRealizada,
          perda: demand.perda,
          motivoPerda: demand.motivoPerda,
          justificativaPerda: demand.justificativaPerda,
          comentario: "Sincronizado automaticamente pela base de realizado.",
          evidencia: sourceName,
          usuario: state.currentUser.email,
        });
      }
    }
    await state.repo.addLog({
      usuario: state.currentUser.email,
      acao: "Sincronização Realizados",
      lista: "Historico_Realizado_Perdas",
      referencia: sourceName,
      detalhe: `${updates.length} realizadas atualizadas. ${unmatched.length} sem correspondência.`,
    });
    if (showResult)
      showToast(
        `${updates.length} ordens realizadas sincronizadas.`,
        "success",
      );
    return { updates: updates.length, unmatched: unmatched.length };
  }

  function workbookRowsFromBuffer(buffer) {
    if (!global.XLSX)
      throw new Error(
        "Biblioteca XLSX indisponível para sincronizar realizados.",
      );
    const workbook = global.XLSX.read(buffer, {
      type: "array",
      cellDates: true,
    });
    const sheetName = workbook.SheetNames[0];
    return global.XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      defval: "",
    });
  }

  async function autoSyncRealizadosFromSharePoint() {
    if (
      state.realizedAutoSynced ||
      state.repo.mode !== "SharePoint REST" ||
      !state.repo.getRealizadosFileBuffer
    )
      return;
    state.realizedAutoSynced = true;
    try {
      const buffer = await state.repo.getRealizadosFileBuffer(
        state.db.parametros || {},
      );
      const rows = workbookRowsFromBuffer(buffer);
      const result = await syncRealizedRows(
        rows,
        state.db.parametros?.realizedExcelFileName ||
          "base_realizados_sap.xlsx",
      );
      if (result.updates) await loadDatabase();
    } catch (error) {
      await state.repo.addLog?.({
        usuario: state.currentUser?.email || "",
        acao: "Falha Sincronização Realizados",
        lista: "Base_Realizados_SAP",
        referencia:
          state.db.parametros?.realizedExcelFileName ||
          "base_realizados_sap.xlsx",
        detalhe: error.message,
      });
    }
  }

  function isFutureDemand(item) {
    const origem = normalizeText(item.origem);
    const tipo = normalizeText(item.tipoDemanda);

    return (
      origem.includes("DEMANDAS FUTURAS") ||
      origem.includes("DEMANDA ANTECIPADA") ||
      tipo.includes("FUTURA") ||
      tipo.includes("SISTEMATICA") ||
      tipo.includes("SISTEMÁTICA") ||
      !item.ordem
    );
  }

  function filteredFutureDemandas() {
    const search = normalizeText(state.futureSearch);

    return state.db.demandas.filter((item) => {
      if (!isFutureDemand(item)) return false;

      if (!search) return true;

      const haystack = normalizeText(
        [
          item.id,
          item.ordem,
          item.descricao,
          item.centroTrabalho,
          item.localInstalacao,
          item.gerencia,
          item.supervisao,
          item.competencia,
          item.origem,
          item.frequencia,
        ].join(" "),
      );

      return haystack.includes(search);
    });
  }

  function renderFutureDemandas() {
    const futures = filteredFutureDemandas();

    const totalPages = Math.max(
      1,
      Math.ceil(futures.length / state.futurePageSize),
    );

    if (state.futurePage > totalPages) {
      state.futurePage = totalPages;
    }

    const start = (state.futurePage - 1) * state.futurePageSize;
    const pageRows = futures.slice(start, start + state.futurePageSize);

    $("#futureCount").textContent =
      `${futures.length} demandas encontradas • Página ${state.futurePage} de ${totalPages}`;

    const toolbarHtml = `
      <div class="future-toolbar-pro">
        <div class="future-toolbar-main">
          <label class="future-search-box">
            <span>Buscar demanda futura</span>
            <div class="future-search-input-wrap">
              ${iconSvg("filter")}
              <input
                id="futureSearch"
                type="search"
                placeholder="Digite ID, OM, descrição, centro, local, gerência ou supervisão"
                value="${escapeHtml(state.futureSearch)}"
              />
            </div>
          </label>

          <label class="future-page-size">
            <span>Linhas por página</span>
            <select id="futurePageSize">
              <option value="25" ${state.futurePageSize === 25 ? "selected" : ""}>25</option>
              <option value="50" ${state.futurePageSize === 50 ? "selected" : ""}>50</option>
              <option value="100" ${state.futurePageSize === 100 ? "selected" : ""}>100</option>
              <option value="200" ${state.futurePageSize === 200 ? "selected" : ""}>200</option>
            </select>
          </label>
        </div>

        <div class="future-pagination-pro">
          <button
            class="future-page-button"
            id="futurePrevPage"
            type="button"
            ${state.futurePage <= 1 ? "disabled" : ""}
          >
            ${iconSvg("chevron-left")}
            <span>Anterior</span>
          </button>

          <div class="future-page-indicator">
            <strong>${state.futurePage}</strong>
            <span>de ${totalPages}</span>
          </div>

          <button
            class="future-page-button"
            id="futureNextPage"
            type="button"
            ${state.futurePage >= totalPages ? "disabled" : ""}
          >
            <span>Próxima</span>
            ${iconSvg("chevron-down")}
          </button>
        </div>
      </div>
    `;

    const listHtml =
      pageRows
        .map((item) => {
          const possuiOM = Boolean(String(item.ordem || "").trim());

          return `
          <article class="future-card" data-future-card="${escapeHtml(item.id)}">
            <header>
              <div>
                <h3>${escapeHtml(item.descricao || "-")}</h3>
                <span class="muted">
                  ${escapeHtml(item.id)}
                  |
                  ${possuiOM ? `OM ${escapeHtml(item.ordem)}` : "Sem OM SAP"}
                  |
                  ${escapeHtml(item.centroTrabalho || "-")}
                  |
                  ${escapeHtml(item.localInstalacao || "-")}
                  |
                  ${escapeHtml(item.competencia || "-")}
                </span>
              </div>
              ${statusChipGroup(statusListOf(item))}
            </header>

            <div class="detail-grid" style="margin-top: 10px;">
              <div class="detail-item">
                <span>Vencimento</span>
                <strong>${formatDate(item.vencimento)}</strong>
              </div>
              <div class="detail-item">
                <span>Frequência</span>
                <strong>${escapeHtml(item.frequencia || "-")}</strong>
              </div>
              <div class="detail-item">
                <span>Gerência</span>
                <strong>${escapeHtml(item.gerencia || "-")}</strong>
              </div>
              <div class="detail-item">
                <span>Supervisão</span>
                <strong>${escapeHtml(item.supervisao || "-")}</strong>
              </div>
            </div>

            <div class="suggestion-list" id="suggestions-${escapeHtml(item.id)}">
              ${
                possuiOM
                  ? `<span class="muted">Esta demanda já possui OM SAP vinculada. Sugestão não necessária.</span>`
                  : `<button class="button secondary planner-only" type="button" data-load-suggestions="${escapeHtml(item.id)}">
                      Ver sugestões de vínculo
                    </button>`
              }
            </div>
          </article>
        `;
        })
        .join("") ||
      '<div class="empty-detail"><strong>Nenhuma demanda futura encontrada</strong><span>Verifique a busca, a base_futuras.json ou os registros criados no sistema.</span></div>';

    $("#futureDemandList").innerHTML = toolbarHtml + listHtml;

    const searchInput = $("#futureSearch");
    if (searchInput) {
      searchInput.addEventListener("input", (event) => {
        state.futureSearch = event.target.value;
        state.futurePage = 1;
        renderFutureDemandas();
      });
    }

    const pageSize = $("#futurePageSize");
    if (pageSize) {
      pageSize.addEventListener("change", (event) => {
        state.futurePageSize = Number(event.target.value);
        state.futurePage = 1;
        renderFutureDemandas();
      });
    }

    const prev = $("#futurePrevPage");
    if (prev) {
      prev.addEventListener("click", () => {
        state.futurePage = Math.max(1, state.futurePage - 1);
        renderFutureDemandas();
      });
    }

    const next = $("#futureNextPage");
    if (next) {
      next.addEventListener("click", () => {
        state.futurePage = Math.min(totalPages, state.futurePage + 1);
        renderFutureDemandas();
      });
    }

    applyPermissions();
  }

  function renderFutureSuggestions(futureId) {
    const future = demandById(futureId);
    const container = document.getElementById(`suggestions-${futureId}`);

    if (!future || !container) return;

    if (future.ordem) {
      container.innerHTML =
        '<span class="muted">Esta demanda já possui OM SAP vinculada. Sugestão não necessária.</span>';
      return;
    }

    container.innerHTML = '<span class="muted">Calculando sugestões...</span>';

    setTimeout(() => {
      const suggestions = linkSuggestions(future).slice(0, 5);

      container.innerHTML = suggestions.length
        ? suggestions
            .map(
              (suggestion) => `
                <div class="suggestion-item">
                  <div>
                    <strong>
                      ${escapeHtml(suggestion.target.ordem)}
                      |
                      ${escapeHtml(suggestion.target.descricao || "-")}
                    </strong>
                    <div class="muted">
                      ${suggestion.score}% de similaridade
                      |
                      ${escapeHtml(suggestion.target.centroTrabalho || "-")}
                      |
                      ${escapeHtml(suggestion.target.localInstalacao || "-")}
                      |
                      ${escapeHtml(suggestion.target.competencia || "-")}
                    </div>
                  </div>

                  <button
                    class="button editor-only"
                    data-link-future="${escapeHtml(future.id)}"
                    data-link-target="${escapeHtml(suggestion.target.id)}"
                    type="button"
                  >
                    Vincular
                  </button>
                </div>
              `,
            )
            .join("")
        : '<span class="muted">Nenhuma sugestão encontrada para esta demanda.</span>';

      applyPermissions();
    }, 50);
  }

  function tokenOverlap(a, b) {
    const left = new Set(
      normalizeText(a)
        .split(/\s+/)
        .filter((token) => token.length > 3),
    );
    const right = new Set(
      normalizeText(b)
        .split(/\s+/)
        .filter((token) => token.length > 3),
    );
    if (!left.size || !right.size) return 0;
    let hits = 0;
    left.forEach((token) => {
      if (right.has(token)) hits += 1;
    });
    return hits / Math.max(left.size, right.size);
  }

  function linkSuggestions(future) {
    if (future.ordem) return [];
    return state.db.demandas
      .filter((item) => item.ordem && item.id !== future.id)
      .map((target) => {
        let score = 0;
        if (target.centroTrabalho === future.centroTrabalho) score += 30;
        if (target.localInstalacao === future.localInstalacao) score += 30;
        if (target.competencia === future.competencia) score += 20;
        score += Math.round(
          tokenOverlap(future.descricao, target.descricao) * 20,
        );
        return { target, score };
      })
      .filter((item) => item.score >= 45)
      .sort((a, b) => b.score - a.score);
  }

  async function linkFutureDemand(futureId, targetId) {
    if (!canPlan()) return;
    const future = demandById(futureId);
    const target = demandById(targetId);
    if (!future || !target) return;
    future.ordem = target.ordem;
    future.origem = "Demanda Antecipada Vinculada ao SAP";
    future.vinculadaEm = new Date().toISOString();
    future.statusSistema = target.statusSistema;
    future.prioridade = target.prioridade;
    future.toleranciaMin = target.toleranciaMin;
    future.toleranciaMax = target.toleranciaMax;
    await state.repo.upsertDemanda(prepareDemandForSave(future));
    await state.repo.addLog({
      usuario: state.currentUser.email,
      acao: "Vínculo Demanda/Ordem",
      lista: "Controle_Demandas_Eletrovia",
      referencia: future.id,
      detalhe: `Vinculada à ordem ${target.ordem}`,
    });
    await refreshAll();
    showToast("Demanda futura vinculada à ordem SAP.", "success");
  }

  async function createFutureDemand(event) {
    event.preventDefault();
    if (!canPlan()) {
      showToast("Perfil sem permissão para criar demanda futura.", "error");
      return;
    }
    const form = new FormData(event.currentTarget);
    const record = Object.fromEntries(form.entries());

    record.ordem = "";

    const idManual = String(record.id || "").trim();

    record.id = idManual || global.CCEData.stableDemandId(record);

    record.tipoOM = record.tipoDemanda;
    const mapaCentros = new Map(
      (state.db.centrosTrabalho || [])
        .filter((item) => item.ativo !== false)
        .map((item) => [centroResponsabilidadeChave(item), item]),
    );
    const centro = findResponsabilidadeForDemand(record, mapaCentros);
    record.gerencia = centro?.gerencia || "";
    record.supervisao = centro?.supervisao || "";
    record.planejadorCurto = centro?.planejadorCurto || "";
    record.planejadorCurtoEmail = centro?.planejadorCurtoEmail || "";
    record.planejadorOM = centro?.planejadorOM || "";
    record.planejadorOMEmail = centro?.planejadorOMEmail || "";
    record.programador = centro?.programador || "";
    record.programadorEmail = centro?.programadorEmail || "";
    record.prioridade = "Nao informado";
    record.statusSistema = "PREV";
    record.toleranciaMin = record.vencimento;
    record.toleranciaMax = record.vencimento;
    record.dataPlanejada = "";
    record.dataReplanejadaAtual = "";
    record.dataRealizada = "";
    record.perda = false;
    record.competencia = normalizeCompetencia(record.competencia);
    record.origem = "Sistema - Demanda Futura";
    record.usuarioResponsavel = state.currentUser.email;
    await state.repo.upsertDemanda(prepareDemandForSave(record));
    await state.repo.addLog({
      usuario: state.currentUser.email,
      acao: "Criação Demanda Futura",
      lista: "Controle_Demandas_Eletrovia",
      referencia: record.id,
      detalhe: record.descricao,
    });
    event.currentTarget.reset();
    await refreshAll();
    showToast("Demanda futura criada.", "success");
  }

  function countBy(demands, selector) {
    return demands.reduce((acc, item) => {
      const key = selector(item) || "Não informado";
      acc[key] = (acc[key] || 0) + 1;
      return acc;
    }, {});
  }

  function renderBars(element, counts) {
    const entries = Object.entries(counts)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 6);
    const max = Math.max(1, ...entries.map(([, count]) => count));
    element.innerHTML =
      entries
        .map(
          ([label, count]) => `
        <div class="bar-row">
          <span>${escapeHtml(label)}</span>
          <div class="bar-track"><div class="bar-value" style="width: ${(count / max) * 100}%"></div></div>
          <strong>${count}</strong>
        </div>
      `,
        )
        .join("") || '<span class="muted">Sem dados no recorte.</span>';
  }

  function renderNotificationCards(notifications) {
    const stats = notificationStats(notifications);

    const cards = [
      ["Todas", "total", stats.total, "pendências abertas"],
      ["Vencidas", "vencida", stats.vencida || 0, "prazo estourado"],
      ["Vencendo", "vencendo", stats.vencendo || 0, "até 7 dias"],
      [
        "Replanejamento",
        "replanejamento-incompleto",
        stats["replanejamento-incompleto"] || 0,
        "motivo/justificativa",
      ],
      [
        "Fora do prazo",
        "realizada-fora-prazo",
        stats["realizada-fora-prazo"] || 0,
        "regularizar perda",
      ],
      [
        "Perda incompleta",
        "perda-incompleta",
        stats["perda-incompleta"] || 0,
        "perfil/justificativa",
      ],
    ];

    $("#notificationCards").innerHTML = cards
      .map(
        ([label, type, value, note]) => `
        <button
          class="notification-card ${
            state.notifications.typeFilter === type ||
            (!state.notifications.typeFilter && type === "total")
              ? "is-active"
              : ""
          }"
          type="button"
          data-notification-card="${type}"
        >
          <span>${escapeHtml(label)}</span>
          <strong>${value}</strong>
          <small>${escapeHtml(note)}</small>
        </button>
      `,
      )
      .join("");
  }

  function renderNotifications() {
    const all = getNotificationsCached();
    const rows = filteredNotifications();
    const visibleRows = rows.slice(0, 300);

    renderNotificationCards(all);

    $("#notificationTypeFilter").value = state.notifications.typeFilter;
    $("#notificationSearch").value = state.notifications.search;

    $("#notificationCount").textContent =
      rows.length > visibleRows.length
        ? `${rows.length} pendências encontradas • exibindo as primeiras ${visibleRows.length}`
        : `${rows.length} pendências encontradas`;

    const tbody = $("#notificationTableBody");

    tbody.innerHTML = visibleRows.length
      ? visibleRows
          .map(
            (notification) => `
            <tr>
              <td>
                <span class="notification-badge ${notification.className}">
                  ${escapeHtml(notification.criticality)}
                </span>
              </td>
              <td>${escapeHtml(notification.label)}</td>
              <td>
                <strong>${escapeHtml(notification.ordem || notification.demandId)}</strong>
                <div class="muted">${escapeHtml(notification.demandId)}</div>
              </td>
              <td class="description-cell">${escapeHtml(notification.descricao)}</td>
              <td>${escapeHtml(notification.gerencia || "-")}</td>
              <td>${escapeHtml(notification.supervisao || "-")}</td>
              <td>${escapeHtml(notification.centroTrabalho || "-")}</td>
              <td>${formatDate(notification.vencimento)}</td>
              <td>${statusChip(notification.status)}</td>
              <td>${escapeHtml(notification.message)}</td>
              <td>
                <div class="row-actions notification-actions">
                  <button
                    class="button compact-button"
                    type="button"
                    data-notification-action="${escapeHtml(notification.id)}"
                  >
                    ${escapeHtml(notificationActionLabel(notification.action))}
                  </button>

                  <button
                    class="button secondary compact-button"
                    type="button"
                    data-notification-open="${escapeHtml(notification.demandId)}"
                  >
                    Carteira
                  </button>
                </div>
              </td>
            </tr>
          `,
          )
          .join("")
      : `
      <tr>
        <td colspan="11">
          <div class="empty-detail">
            <strong>Nenhuma pendência no recorte</strong>
            <span>Altere os filtros ou verifique se as pendências já foram tratadas.</span>
          </div>
        </td>
      </tr>
    `;
  }

  function openDemandFromNotification(demandId) {
    const demand = demandById(demandId);

    if (!demand) {
      showToast("Demanda não encontrada na carteira atual.", "error");
      return;
    }

    state.selectedDemandId = demandId;
    $("#quickSearch").value = demand.ordem || demand.id;
    state.page = 1;

    collectFilters();
    buildFilterOptions();
    switchView("carteira");
    renderCarteira();
  }

  function openNotificationAction(notificationId) {
    const notification = findNotification(notificationId);

    if (!notification) {
      showToast("Notificação não encontrada.", "error");
      return;
    }

    if (notification.action === "planejar") {
      openAction("planejar", notification.demandId);
      return;
    }

    if (notification.action === "replanejar") {
      openAction("replanejar", notification.demandId);
      return;
    }

    if (notification.action === "regularizar-replanejamento") {
      openReplanNotificationDialog(notification);
      return;
    }

    if (notification.action === "regularizar-perda") {
      openLossNotificationDialog(notification);
      return;
    }

    openDemandFromNotification(notification.demandId);
  }

  function renderIndicators() {
    const indicatorFilters = collectIndicatorFilters();
    const demands = state.db.demandas.filter((item) =>
      demandMatchesFilters(item, indicatorFilters),
    );
    const stats = dashboardStats(demands);
    const dueSoon = demands
      .filter((item) => {
        const due = toDate(item.vencimento);
        const today = toDate(todayText());
        if (!due || item.dataRealizada) return false;
        const days = (due - today) / 86400000;
        return days >= 0 && days <= 20;
      })
      .sort((a, b) => toDate(a.vencimento) - toDate(b.vencimento));
    const overdue = demands.filter(
      (item) =>
        toDate(item.vencimento) < toDate(todayText()) && !item.dataRealizada,
    ).length;
    const cards = [
      ["Total de Demandas", stats.total, "recorte filtrado"],
      ["Total a Planejar", stats.aPlanejar, "sem data"],
      ["Planejadas", stats.planejadas, "ativas"],
      ["Replanejadas", stats.replanejadas, "com histórico"],
      ["Realizadas", stats.realizadas, "baixadas"],
      ["Ordens Vencidas", overdue, "sem realização"],
      ["Próximas do Vencimento", dueSoon.length, "20 dias"],
    ];
    $("#indicatorGrid").innerHTML = cards
      .map(
        ([label, value, note]) =>
          `<article class="indicator-card"><span>${label}</span><strong>${value}</strong><small>${note}</small></article>`,
      )
      .join("");
    renderBars(
      $("#lossByCenter"),
      countBy(demands, (item) => item.gerencia),
    );
    renderBars(
      $("#lossReasons"),
      countBy(demands, (item) => item.supervisao),
    );
    renderBars(
      $("#replanRanking"),
      countBy(demands, (item) => item.centroTrabalho),
    );
    $("#dueSoonList").innerHTML =
      dueSoon
        .slice(0, 6)
        .map(
          (item) => `
        <div class="due-item">
          <strong>${formatDate(item.vencimento)}</strong>
          <span>${escapeHtml(item.ordem || item.id)} | ${escapeHtml(item.descricao)}</span>
          ${statusChipGroup(statusListOf(item))}
        </div>
      `,
        )
        .join("") ||
      '<span class="muted">Sem vencimentos nos próximos 20 dias.</span>';
    renderBars(
      $("#statusChart"),
      countBy(demands, (item) => primaryStatusOf(item)),
    );
    renderBars(
      $("#competenceChart"),
      countBy(demands, (item) => item.competencia),
    );
  }

  function renderAdmin() {
    if (!canAdmin()) {
      $("#adminContent").innerHTML =
        '<div class="empty-detail"><strong>Acesso restrito</strong><span>Somente Administrador pode alterar cadastros.</span></div>';
      return;
    }
    $$("#adminTabs button").forEach((button) =>
      button.classList.toggle(
        "is-active",
        button.dataset.adminTab === state.adminTab,
      ),
    );
    if (state.adminTab === "usuarios") renderUserAdmin();
    else if (state.adminTab === "centrosTrabalho") renderCentrosTrabalhoAdmin();
    else if (state.adminTab === "ferias") renderFeriasAdmin();
    else if (state.adminTab === "parametros") renderParameterAdmin();
    else renderConfigAdmin(state.adminTab);
  }

  function renderUserAdmin() {
    $("#adminContent").innerHTML = `
      <div class="admin-grid admin-grid-wide">
        <form class="admin-form" id="addUserForm">
          <label>Nome<input name="nome" required /></label>
          <label>E-mail<input name="email" type="email" required /></label>
          <label>Matrícula<input name="matricula" required /></label>
          <label>Área<input name="area" /></label>
          <label>Perfil<select name="perfil">${optionsMarkup(Object.keys(PROFILE_RULES), "Visualizador")}</select></label>
          <label>Status<select name="ativo"><option value="true">Ativo</option><option value="false">Inativo</option></select></label>
          <fieldset class="permission-grid">
            <legend>Permissões</legend>
            <label><input name="permissaoPlanejar" type="checkbox" /> Planejar</label>
            <label><input name="permissaoReplanejar" type="checkbox" /> Replanejar</label>
            <label><input name="permissaoRealizar" type="checkbox" /> Realizar/perda</label>
            <label><input name="permissaoConfigurar" type="checkbox" /> Configurar</label>
            <label><input name="permissaoExportar" type="checkbox" checked /> Exportar</label>
            <label><input name="permissaoCargaLote" type="checkbox" /> Carga em lote</label>
          </fieldset>
          <button class="button" type="submit">Salvar Usuario</button>
        </form>
        <div class="admin-list table-scroll admin-list-scroll">
          <table class="data-table admin-table">
            <thead>
              <tr><th>Nome</th><th>E-mail</th><th>Matrícula</th><th>Área</th><th>Perfil</th><th>Status</th><th>Permissões</th><th>Ações</th></tr>
            </thead>
            <tbody>
              ${state.db.usuarios
                .map(
                  (user) => `
                <tr>
                  <td>${escapeHtml(user.nome)}</td>
                  <td>${escapeHtml(user.email)}</td>
                  <td>${escapeHtml(user.matricula || "-")}</td>
                  <td>${escapeHtml(user.area || "-")}</td>
                  <td>${escapeHtml(user.perfil)}</td>
                  <td>${user.ativo ? "Ativo" : "Inativo"}</td>
                  <td>${[
                    user.permissaoPlanejar && "planejar",
                    user.permissaoReplanejar && "replanejar",
                    user.permissaoRealizar && "realizar",
                    user.permissaoConfigurar && "configurar",
                    user.permissaoExportar && "exportar",
                    user.permissaoCargaLote && "lote",
                  ]
                    .filter(Boolean)
                    .join(", ")}</td>
                  <td><button class="button secondary" type="button" data-edit-user="${escapeHtml(user.email)}">Editar</button></td>
                </tr>
              `,
                )
                .join("")}
            </tbody>
          </table>
        </div>
      </div>
    `;
    $("#addUserForm").perfil.addEventListener("change", (event) => {
      applyProfilePermissionDefaults(event.currentTarget.form);
    });
    applyProfilePermissionDefaults($("#addUserForm"));
    $("#addUserForm").addEventListener("submit", async (event) => {
      event.preventDefault();

      const form = event.currentTarget;
      const payload = Object.fromEntries(new FormData(form).entries());

      payload.email = String(payload.email || "")
        .trim()
        .toLowerCase();
      payload.nome = String(payload.nome || "").trim();
      payload.matricula = String(payload.matricula || "").trim();
      payload.area = String(payload.area || "").trim();

      payload.permissaoPlanejar = form.permissaoPlanejar.checked;
      payload.permissaoReplanejar = form.permissaoReplanejar.checked;
      payload.permissaoRealizar = form.permissaoRealizar.checked;
      payload.permissaoConfigurar = form.permissaoConfigurar.checked;
      payload.permissaoExportar = form.permissaoExportar.checked;
      payload.permissaoCargaLote = form.permissaoCargaLote.checked;

      const usuarioLogado =
        state.currentUser?.email ||
        getStoredSessionEmail() ||
        "usuario_nao_identificado";

      try {
        await state.repo.addUser(payload);

        await state.repo.addLog?.({
          usuario: usuarioLogado,
          acao: "Cadastro Usuário",
          lista: "usuarios_central_eletrovia",
          referencia: payload.email,
          detalhe: `Usuário salvo pela administração: ${payload.nome || payload.email}`,
          modulo: "ADMINISTRACAO",
          status: "SUCESSO",
        });

        form.reset();
        applyProfilePermissionDefaults(form);

        await refreshAfterSave("Usuário salvo com sucesso.");
      } catch (error) {
        console.error(error);

        showToast(
          `Erro ao salvar usuário: ${error.message || "falha inesperada."}`,
          "error",
        );
      }
    });
    $("#adminContent").addEventListener("click", (event) => {
      const button = event.target.closest("[data-edit-user]");
      if (!button) return;
      const user = state.db.usuarios.find(
        (item) => item.email === button.dataset.editUser,
      );
      if (!user) return;
      const form = $("#addUserForm");
      form.nome.value = user.nome || "";
      form.email.value = user.email || "";
      form.matricula.value = user.matricula || "";
      form.area.value = user.area || "";
      form.perfil.value = user.perfil || "Visualizador";
      form.ativo.value = user.ativo ? "true" : "false";
      form.permissaoPlanejar.checked = user.permissaoPlanejar;
      form.permissaoReplanejar.checked = user.permissaoReplanejar;
      form.permissaoRealizar.checked = user.permissaoRealizar;
      form.permissaoConfigurar.checked = user.permissaoConfigurar;
      form.permissaoExportar.checked = user.permissaoExportar;
      form.permissaoCargaLote.checked = user.permissaoCargaLote;
    });
  }

  function applyProfilePermissionDefaults(form) {
    const defaults = profileDefaults(form.perfil.value);
    form.permissaoPlanejar.checked = defaults.planejar;
    form.permissaoReplanejar.checked = defaults.replanejar;
    form.permissaoRealizar.checked = defaults.realizar;
    form.permissaoConfigurar.checked = defaults.configurar;
    form.permissaoExportar.checked = defaults.exportar;
    form.permissaoCargaLote.checked = defaults.cargaLote;
  }

  function renderConfigAdmin(group) {
    const labels = {
      motivos: {
        title: "Motivos e Justificativas",
        parent: "Motivo",
        child: "Justificativa",
        parentGroup: "motivos",
        childGroup: "justificativas",
        childField: "motivoId",
      },
      perfisPerda: {
        title: "Perfis e Justificativas de Perda",
        parent: "Perfil de perda",
        child: "Justificativa de perda",
        parentGroup: "perfisPerda",
        childGroup: "justificativasPerda",
        childField: "perfilId",
      },
    };
    const config = labels[group] || labels.motivos;
    const parents = configItems(config.parentGroup);
    const children = configItems(config.childGroup);
    $("#adminContent").innerHTML = `
      <div class="admin-grid admin-grid-wide">
        <div class="admin-form-stack">
          <form class="admin-form" id="addConfigForm">
            <label>Novo ${config.parent}<input name="value" required /></label>
            <button class="button" type="submit">Adicionar ${config.parent}</button>
          </form>
          <form class="admin-form" id="addChildConfigForm">
            <label>${config.parent}<select name="parentId" required>${parents.map((item) => `<option value="${escapeHtml(item.id)}">${escapeHtml(item.nome)}</option>`).join("")}</select></label>
            <label>Nova ${config.child}<input name="value" required /></label>
            <button class="button secondary" type="submit">Adicionar ${config.child}</button>
          </form>
        </div>
        <div class="admin-list admin-list-scroll">
          ${parents
            .map((parent) => {
              const childList = children.filter(
                (child) => child[config.childField] === parent.id,
              );
              return `
                <div class="admin-list-item grouped-config">
                  <div>
                    <strong>${escapeHtml(parent.nome)}</strong>
                    <div class="muted">${childList.length ? childList.map((child) => escapeHtml(child.nome)).join(" | ") : `Sem ${config.child.toLowerCase()} cadastrada`}</div>
                  </div>
                </div>
              `;
            })
            .join("")}
        </div>
      </div>
    `;
    $("#addConfigForm").addEventListener("submit", async (event) => {
      event.preventDefault();
      await state.repo.addConfigItem(
        config.parentGroup,
        new FormData(event.currentTarget).get("value"),
      );
      await state.repo.addLog({
        usuario: state.currentUser.email,
        acao: "Cadastro Configuração",
        lista: "Configuracoes",
        referencia: config.title,
        detalhe: new FormData(event.currentTarget).get("value"),
      });
      await refreshAfterSave("Configuração salva com sucesso.");
    });
    $("#addChildConfigForm").addEventListener("submit", async (event) => {
      event.preventDefault();
      const form = new FormData(event.currentTarget);
      await state.repo.addConfigItem(
        config.childGroup,
        form.get("value"),
        form.get("parentId"),
      );
      await state.repo.addLog({
        usuario: state.currentUser.email,
        acao: "Cadastro Subgrupo",
        lista: "Configuracoes",
        referencia: config.title,
        detalhe: form.get("value"),
      });
      await refreshAfterSave("Subgrupo salvo com sucesso.");
    });
  }

  function renderCentrosTrabalhoAdmin() {
    const centros = state.db.centrosTrabalho || [];
    const centrosNaoCadastrados = missingCentrosTrabalho();

    $("#adminContent").innerHTML = `
    <div class="admin-grid admin-grid-wide">
      <form class="admin-form" id="centroTrabalhoForm">
        <label>
          Nivel de Responsabilidade
          <select name="nivelResponsabilidade">
            <option value="centro" selected>Centro de Trabalho</option>
            <option value="supervisao">Supervisao</option>
            <option value="gerencia">Gerencia</option>
          </select>
        </label>

        <label>
          Centro de Trabalho
          <input name="centroTrabalho" placeholder="Ex.: EVT-PCM-01" />
        </label>

        <label>
          Gerência
          <input name="gerencia" required placeholder="Ex.: Gerência Eletrovia" />
        </label>

        <label>
          Supervisão
          <input name="supervisao" required placeholder="Ex.: Supervisão PCM Eletrovia" />
        </label>

        <label>
          Planejador de Curto
          <input name="planejadorCurto" placeholder="Nome do planejador de curto" />
        </label>

        <label>
          E-mail Planejador Curto
          <input name="planejadorCurtoEmail" type="email" placeholder="planejador.curto@vale.com" />
        </label>

        <label>
          Matricula Planejador Curto
          <input name="planejadorCurtoMatricula" placeholder="000000" />
        </label>

        <label>
          Planejador de OM
          <input name="planejadorOM" placeholder="Nome do planejador de OM" />
        </label>

        <label>
          E-mail Planejador OM
          <input name="planejadorOMEmail" type="email" placeholder="planejador@vale.com" />
        </label>

        <label>
          Matrícula Planejador OM
          <input name="planejadorOMMatricula" placeholder="000000" />
        </label>

        <label>
          Programador
          <input name="programador" placeholder="Nome do programador" />
        </label>

        <label>
          E-mail Programador
          <input name="programadorEmail" type="email" placeholder="programador@vale.com" />
        </label>

        <label>
          Matrícula Programador
          <input name="programadorMatricula" placeholder="000000" />
        </label>

        <label>
          Área
          <input name="area" placeholder="Ex.: PCM Eletrovia" />
        </label>

        <label class="span-2">
          Observação
          <textarea name="observacao" rows="3"></textarea>
        </label>

        <label>
          Ativo
          <select name="ativo">
            <option value="true" selected>Sim</option>
            <option value="false">Não</option>
          </select>
        </label>

        <div class="admin-form-actions span-2">
          <button class="button" type="submit">
            Salvar Centro de Trabalho
          </button>
          <button
            class="button secondary"
            id="clearCentroTrabalhoForm"
            type="button"
          >
            Limpar
          </button>
        </div>
      </form>

      <div class="admin-list admin-list-scroll">
        ${
          centrosNaoCadastrados.length
            ? `
              <section class="missing-centers">
                <strong>Centros encontrados no JSON sem cadastro mestre</strong>
                <div class="missing-center-list">
                  ${centrosNaoCadastrados
                    .slice(0, 60)
                    .map(
                      (item) => `
                        <button class="button secondary" type="button" data-new-centro="${escapeHtml(item.centro)}">
                          ${escapeHtml(item.centro)} (${item.total})
                        </button>
                      `,
                    )
                    .join("")}
                </div>
              </section>
            `
            : ""
        }
        ${
          centros.length
            ? centros
                .map(
                  (item) => `
                    <div class="admin-list-item">
                      <div>
                        <strong>${escapeHtml(item.centroTrabalho || "-")}</strong>
                        <div class="muted">
                          ${escapeHtml(item.gerencia || "-")}
                          |
                          ${escapeHtml(item.supervisao || "-")}
                        </div>
                        <div class="muted">
                          Nivel: ${escapeHtml(centroResponsabilidadeNivel(item))}
                          |
                          Planejador Curto: ${escapeHtml(item.planejadorCurto || "-")}
                          |
                          Planejador OM: ${escapeHtml(item.planejadorOM || "-")}
                          |
                          Programador: ${escapeHtml(item.programador || "-")}
                        </div>
                        <div class="muted">
                          ${item.ativo !== false ? "Ativo" : "Inativo"}
                        </div>
                      </div>
                      <button
                        class="button secondary"
                        type="button"
                        data-edit-centro="${escapeHtml(centroResponsabilidadeChave(item))}"
                      >
                        Editar
                      </button>
                    </div>
                  `,
                )
                .join("")
            : '<div class="empty-detail"><strong>Nenhum centro de trabalho cadastrado</strong><span>Cadastre o primeiro centro para enriquecer a carteira.</span></div>'
        }
      </div>
    </div>
  `;

    $("#centroTrabalhoForm").addEventListener("submit", async (event) => {
      event.preventDefault();

      const form = new FormData(event.currentTarget);
      const record = Object.fromEntries(form.entries());

      if (
        record.nivelResponsabilidade === "centro" &&
        !String(record.centroTrabalho || "").trim()
      ) {
        showToast("Informe o Centro de Trabalho para cadastro por centro.", "error");
        return;
      }

      record.ativo = record.ativo === "true";
      record.usuario = state.currentUser.email;

      await state.repo.upsertCentroTrabalho(record);

      await state.repo.addLog({
        usuario: state.currentUser.email,
        acao: "Cadastro Centro de Trabalho",
        lista: "cadastro_centros_trabalho",
        referencia: record.centroTrabalho || record.supervisao || record.gerencia,
        detalhe: `${record.nivelResponsabilidade || "centro"} | ${record.gerencia || "-"} | ${record.supervisao || "-"}`,
        modulo: "CONFIGURACOES",
      });

      await refreshAfterSave("Centro de trabalho salvo com sucesso.");
    });

    $("#clearCentroTrabalhoForm").addEventListener("click", () => {
      const form = $("#centroTrabalhoForm");
      form.reset();
      form.nivelResponsabilidade.value = "centro";
      form.ativo.value = "true";
      form.centroTrabalho.focus();
    });

    $("#adminContent").addEventListener("click", (event) => {
      const button = event.target.closest("[data-edit-centro]");
      const newButton = event.target.closest("[data-new-centro]");
      if (!button && !newButton) return;

      if (newButton) {
        const form = $("#centroTrabalhoForm");
        form.reset();
        form.nivelResponsabilidade.value = "centro";
        form.centroTrabalho.value = newButton.dataset.newCentro || "";
        form.ativo.value = "true";
        form.gerencia.focus();
        return;
      }

      const centro = centros.find(
        (item) => centroResponsabilidadeChave(item) === button.dataset.editCentro,
      );

      if (!centro) return;

      const form = $("#centroTrabalhoForm");

      form.nivelResponsabilidade.value = centroResponsabilidadeNivel(centro);
      form.centroTrabalho.value = centro.centroTrabalho || "";
      form.gerencia.value = centro.gerencia || "";
      form.supervisao.value = centro.supervisao || "";
      form.planejadorCurto.value = centro.planejadorCurto || "";
      form.planejadorCurtoEmail.value = centro.planejadorCurtoEmail || "";
      form.planejadorCurtoMatricula.value = centro.planejadorCurtoMatricula || "";
      form.planejadorOM.value = centro.planejadorOM || "";
      form.planejadorOMEmail.value = centro.planejadorOMEmail || "";
      form.planejadorOMMatricula.value = centro.planejadorOMMatricula || "";
      form.programador.value = centro.programador || "";
      form.programadorEmail.value = centro.programadorEmail || "";
      form.programadorMatricula.value = centro.programadorMatricula || "";
      form.area.value = centro.area || "";
      form.observacao.value = centro.observacao || "";
      form.ativo.value = centro.ativo !== false ? "true" : "false";
    });
  }

  function missingCentrosTrabalho() {
    const counts = new Map();
    (state.db.demandas || []).forEach((item) => {
      if (item.centroTrabalhoCadastrado !== false) return;
      const centro = item.centroTrabalho || "";
      if (!centro) return;
      counts.set(centro, (counts.get(centro) || 0) + 1);
    });
    return Array.from(counts, ([centro, total]) => ({ centro, total })).sort(
      (a, b) => b.total - a.total || a.centro.localeCompare(b.centro),
    );
  }

  function renderFeriasAdmin() {
    const ferias = state.db.feriasSubstituicoes || [];

    $("#adminContent").innerHTML = `
    <div class="admin-grid admin-grid-wide">
      <form class="admin-form" id="feriasForm">
        <label>
          E-mail de quem vai sair de férias
          <input name="emailAusente" type="email" required placeholder="planejador@vale.com" />
        </label>

        <label>
          Matrícula de quem vai sair
          <input name="matriculaAusente" placeholder="000000" />
        </label>

        <label>
          E-mail do substituto
          <input name="emailSubstituto" type="email" required placeholder="substituto@vale.com" />
        </label>

        <label>
          Matrícula do substituto
          <input name="matriculaSubstituto" placeholder="000000" />
        </label>

        <label>
          Data início
          <input name="dataInicio" type="date" required />
        </label>

        <label>
          Data fim
          <input name="dataFim" type="date" required />
        </label>

        <label>
          Escopo Gerência
          <input name="escopoGerencia" placeholder="Ex.: GAL I - deixar vazio para todos" />
        </label>

        <label>
          Escopo Centro de Trabalho
          <input name="escopoCentroTrabalho" placeholder="Ex.: CESSL2 - deixar vazio para todos" />
        </label>

        <label>
          Ativo
          <select name="ativo">
            <option value="true" selected>Sim</option>
            <option value="false">Não</option>
          </select>
        </label>

        <label class="span-2">
          Observação
          <textarea name="observacao" rows="3" placeholder="Ex.: Substituição durante férias do planejador responsável."></textarea>
        </label>

        <button class="button" type="submit">
          Salvar Substituição
        </button>
      </form>

      <div class="admin-list admin-list-scroll">
        ${
          ferias.length
            ? ferias
                .map(
                  (item) => `
                    <div class="admin-list-item">
                      <div>
                        <strong>
                          ${escapeHtml(item.emailAusente)}
                          →
                          ${escapeHtml(item.emailSubstituto)}
                        </strong>

                        <div class="muted">
                          ${formatDate(item.dataInicio)}
                          até
                          ${formatDate(item.dataFim)}
                          |
                          ${item.ativo ? "Ativo" : "Inativo"}
                        </div>

                        <div class="muted">
                          Gerência:
                          ${escapeHtml(item.escopoGerencia || "Todas")}
                          |
                          Centro:
                          ${escapeHtml(item.escopoCentroTrabalho || "Todos")}
                        </div>

                        <div class="muted">
                          ${escapeHtml(item.observacao || "")}
                        </div>
                      </div>

                      <button
                        class="button secondary"
                        type="button"
                        data-edit-ferias="${escapeHtml(item.id)}"
                      >
                        Editar
                      </button>
                    </div>
                  `,
                )
                .join("")
            : '<div class="empty-detail"><strong>Nenhuma substituição cadastrada</strong><span>Cadastre férias para redirecionar notificações e responsabilidades temporárias.</span></div>'
        }
      </div>
    </div>
  `;

    $("#feriasForm").addEventListener("submit", async (event) => {
      event.preventDefault();

      const formElement = event.currentTarget;
      const form = new FormData(formElement);
      const record = Object.fromEntries(form.entries());

      if (toDate(record.dataFim) < toDate(record.dataInicio)) {
        showToast("A data fim não pode ser menor que a data início.", "error");
        return;
      }

      record.ativo = record.ativo === "true";

      await state.repo.upsertFeriasSubstituicao(record);

      await state.repo.addLog({
        usuario: state.currentUser.email,
        acao: "Cadastro Férias/Substituição",
        lista: "ferias_substituicoes",
        referencia: `${record.emailAusente} -> ${record.emailSubstituto}`,
        detalhe: `${record.dataInicio} até ${record.dataFim}`,
        modulo: "ADMINISTRACAO",
        status: "SUCESSO",
      });

      await refreshAfterSave("Substituição de férias salva com sucesso.");
    });

    $("#adminContent").addEventListener("click", (event) => {
      const button = event.target.closest("[data-edit-ferias]");
      if (!button) return;

      const item = ferias.find(
        (registro) => registro.id === button.dataset.editFerias,
      );
      if (!item) return;

      const form = $("#feriasForm");

      form.emailAusente.value = item.emailAusente || "";
      form.matriculaAusente.value = item.matriculaAusente || "";
      form.emailSubstituto.value = item.emailSubstituto || "";
      form.matriculaSubstituto.value = item.matriculaSubstituto || "";
      form.dataInicio.value = item.dataInicio || "";
      form.dataFim.value = item.dataFim || "";
      form.escopoGerencia.value = item.escopoGerencia || "";
      form.escopoCentroTrabalho.value = item.escopoCentroTrabalho || "";
      form.ativo.value = item.ativo !== false ? "true" : "false";
      form.observacao.value = item.observacao || "";

      let hiddenId = form.querySelector('[name="id"]');

      if (!hiddenId) {
        hiddenId = document.createElement("input");
        hiddenId.type = "hidden";
        hiddenId.name = "id";
        form.appendChild(hiddenId);
      }

      hiddenId.value = item.id;
    });
  }

  function renderParameterAdmin() {
    const params = state.db.parametros || {};
    if (!state.db.parametrosDisponiveis) {
      $("#adminContent").innerHTML = `
        <div class="empty-detail">
          <strong>Parametros ainda nao migrados</strong>
          <span>A tabela parametros_sistema nao existe no Supabase atual. A aba fica somente informativa para evitar gravacao falsa.</span>
        </div>
      `;
      return;
    }
    $("#adminContent").innerHTML = `
      <form class="admin-form admin-parameter-form" id="parameterForm">
        <label>Competência atual<input name="currentCompetencia" value="${escapeHtml(params.currentCompetencia || "")}" /></label>
        <label>Tolerância padrão antes (dias)<input name="defaultToleranceBeforeDays" type="number" value="${escapeHtml(params.defaultToleranceBeforeDays || 3)}" /></label>
        <label>Tolerância padrão depois (dias)<input name="defaultToleranceAfterDays" type="number" value="${escapeHtml(params.defaultToleranceAfterDays || 5)}" /></label>
        <label>Biblioteca SAP BO<input name="sharePointLibrary" value="${escapeHtml(params.sharePointLibrary || "")}" /></label>
        <label>Arquivo SAP BO<input name="sapExcelFileName" value="${escapeHtml(params.sapExcelFileName || "")}" /></label>
        <label>Arquivo de Realizados SAP BO<input name="realizedExcelFileName" value="${escapeHtml(params.realizedExcelFileName || "base_realizados_sap.xlsx")}" /></label>
        <button class="button" type="submit">Salvar Parâmetros</button>
      </form>
    `;
    $("#parameterForm").addEventListener("submit", async (event) => {
      event.preventDefault();
      await state.repo.updateParameters(
        Object.fromEntries(new FormData(event.currentTarget).entries()),
      );
      await state.repo.addLog({
        usuario: state.currentUser.email,
        acao: "Parâmetros",
        lista: "Parametros_Sistema",
        referencia: "Geral",
        detalhe: "Parâmetros atualizados.",
      });
      await refreshAfterSave("Parâmetros salvos com sucesso.");
    });
  }

  function renderLogs() {
    $("#logsTableBody").innerHTML = state.db.logs
      .slice(0, 200)
      .map(
        (log) => `
        <tr>
          <td>${formatDateTime(log.dataHora)}</td>
          <td>${escapeHtml(log.usuario || "-")}</td>
          <td>${escapeHtml(log.acao || "-")}</td>
          <td>${escapeHtml(log.lista || "-")}</td>
          <td>${escapeHtml(log.referencia || "-")}</td>
          <td>${escapeHtml(log.detalhe || "-")}</td>
        </tr>
      `,
      )
      .join("");
  }

  function integrationStatusChip(status) {
    const normalized = normalizeText(status);
    const className = normalized.includes("FALHA")
      ? "status-perda"
      : normalized.includes("ALERTA")
        ? "status-planejar"
        : "status-realizado";
    return `<span class="status-chip ${className}">${escapeHtml(status)}</span>`;
  }

  function integrationFailureLogs() {
    const failTerms = [
      "ERRO",
      "FALHA",
      "API",
      "SUPABASE",
      "JSON",
      "BASE",
      "PARAMETRO",
      "SINCRONIZACAO",
      "WORKER",
      "CLOUDFLARE",
    ];

    return (state.db?.logs || [])
      .filter((log) => {
        const text = normalizeText(
          [log.acao, log.lista, log.referencia, log.detalhe, log.status].join(
            " ",
          ),
        );
        return failTerms.some((term) => text.includes(term));
      })
      .sort((a, b) => new Date(b.dataHora || 0) - new Date(a.dataHora || 0));
  }

  function integrationSourceRows() {
    const qualityRecords = state.db?.qualitySourceRecords || [];
    const ordens = qualityRecords.filter((record) =>
      normalizeText(qualitySourceLabel(record)).includes("ORDENS"),
    ).length;
    const futuras = qualityRecords.filter((record) =>
      normalizeText(qualitySourceLabel(record)).includes("FUTURAS"),
    ).length;
    const realizados = state.db?.qualityBaseRealizados?.length || 0;
    const supabaseDemandas = (state.db?.demandas || []).filter((record) =>
      normalizeText(record.origem).includes("SUPABASE"),
    ).length;
    const centros = state.db?.centrosTrabalho?.length || 0;
    const cloudflareUrl = global.CCESupabase?.url || "";
    const lastUpdate = state.lastDataUpdateAt || latestDataUpdateAt();
    const failures = integrationFailureLogs();

    return [
      {
        component: "base_ordens.json",
        status: ordens ? "OK" : "Alerta",
        updatedAt: lastUpdate,
        detail: `${ordens} registros carregados da base de ordens.`,
      },
      {
        component: "base_futuras.json",
        status: futuras ? "OK" : "Alerta",
        updatedAt: lastUpdate,
        detail: `${futuras} registros carregados da base futura.`,
      },
      {
        component: "base_realizados.json",
        status: realizados ? "OK" : "Alerta",
        updatedAt: lastUpdate,
        detail: `${realizados} registros carregados da base de realizados.`,
      },
      {
        component: "Supabase",
        status: state.repo?.mode ? "OK" : "Falha",
        updatedAt: lastUpdate,
        detail: `${state.repo?.mode || "Repositorio indisponivel"} | ${supabaseDemandas} demandas somente Supabase.`,
      },
      {
        component: "Cloudflare Worker",
        status: cloudflareUrl ? "OK" : "Alerta",
        updatedAt: lastUpdate,
        detail: cloudflareUrl || "Worker nao informado no adaptador.",
      },
      {
        component: "cadastro_centros_trabalho",
        status: centros ? "OK" : "Alerta",
        updatedAt: lastUpdate,
        detail: `${centros} centros/responsabilidades cadastrados.`,
      },
      {
        component: "Fluxo de carga",
        status: failures.length ? "Alerta" : "OK",
        updatedAt: lastUpdate,
        detail: failures.length
          ? `${failures.length} falhas ou alertas recentes nos logs.`
          : "Sem falhas recentes registradas em logs.",
      },
      {
        component: "Parametros do sistema",
        status: state.db?.parametrosDisponiveis ? "OK" : "Alerta",
        updatedAt: lastUpdate,
        detail: state.db?.parametrosDisponiveis
          ? "Tabela de parametros disponivel."
          : "Parametros ainda nao migrados ou indisponiveis.",
      },
    ];
  }

  function renderIntegrationHealth() {
    const rows = integrationSourceRows();
    const failures = integrationFailureLogs();
    const okCount = rows.filter((row) => row.status === "OK").length;
    const alertCount = rows.filter((row) => row.status === "Alerta").length;
    const failCount = rows.filter((row) => row.status === "Falha").length;
    const lastUpdate = state.lastDataUpdateAt || latestDataUpdateAt();

    $("#integrationHealthSummary").textContent =
      `${okCount} OK | ${alertCount} alertas | ${failCount} falhas`;

    $("#integrationHealthCards").innerHTML = [
      ["Status geral", failCount ? "Falha" : alertCount ? "Alerta" : "OK", "Bases e integrações"],
      ["Última atualização", formatDateTime(lastUpdate), "Dados carregados"],
      ["Supabase", state.repo?.mode || "-", "Repositório ativo"],
      ["Falhas recentes", failures.length, "logs técnicos"],
    ]
      .map(
        ([label, value, hint]) => `
          <div class="integration-health-card">
            <span>${escapeHtml(label)}</span>
            <strong>${escapeHtml(value)}</strong>
            <small>${escapeHtml(hint)}</small>
          </div>
        `,
      )
      .join("");

    $("#integrationHealthTableBody").innerHTML = rows
      .map(
        (row) => `
          <tr>
            <td><strong>${escapeHtml(row.component)}</strong></td>
            <td>${integrationStatusChip(row.status)}</td>
            <td>${formatDateTime(row.updatedAt)}</td>
            <td class="description-cell">${escapeHtml(row.detail)}</td>
          </tr>
        `,
      )
      .join("");

    $("#integrationFailureCount").textContent =
      `${failures.length} falhas encontradas`;

    $("#integrationFailureTableBody").innerHTML = failures.length
      ? failures
          .slice(0, 80)
          .map(
            (log) => `
              <tr>
                <td>${formatDateTime(log.dataHora)}</td>
                <td>${escapeHtml(log.acao || "-")}</td>
                <td>${escapeHtml(log.referencia || log.lista || "-")}</td>
                <td class="description-cell">${escapeHtml(log.detalhe || "-")}</td>
              </tr>
            `,
          )
          .join("")
      : `<tr>
          <td colspan="4">
            <div class="empty-detail">
              <strong>Sem falhas recentes</strong>
              <span>Os logs não indicam erro técnico ou falha de integração.</span>
            </div>
          </td>
        </tr>`;
  }

  function exportCurrentCarteira() {
    const rows = filteredDemandas();
    const header = [
      "ID_Demanda_Controle",
      "Ordem SAP",
      "Descrição",
      "Origem",
      "Gerencia",
      "Supervisao",
      "Vencimento",
      "Competência",
      "Tipo OM",
      "Centro Trabalho",
      "Cadastro Centro",
      "Planejador Curto",
      "Planejador OM",
      "Programador",
      "Local Instalação",
      "Prioridade",
      "Status Operacional",
      "Substatus",
      "Data Planejada",
      "Data Replanejada",
      "Data Realizada",
      "Perda",
      "Motivo Perda",
      "Ultima Atualizacao",
    ];
    const body = rows.map((item) => [
      item.id,
      item.ordem,
      item.descricao,
      item.origem,
      item.gerencia,
      item.supervisao,
      item.vencimento,
      item.competencia,
      item.tipoOM,
      item.centroTrabalho,
      item.centroTrabalhoStatus,
      item.planejadorCurto,
      item.planejadorOM,
      item.programador,
      item.localInstalacao,
      item.prioridade,
      primaryStatusOf(item),
      substatusListOf(item).join(" | "),
      item.dataPlanejada,
      item.dataReplanejadaAtual,
      item.dataRealizada,
      item.perda ? "Sim" : "Não",
      item.motivoPerda,
      item.dataUltimaAtualizacao,
    ]);
    downloadFile(
      `carteira-eletrovia-${todayText()}.csv`,
      toCsv([header, ...body]),
    );
  }

  function downloadTemplate() {
    const header = [
      "ID_Demanda_Controle",
      "ordem",
      "descricao",
      "centro trabalho",
      "local instalacao",
      "competencia",
      "vencimento",
      "data planejada",
      "data replanejada",
      "motivo",
      "justificativa",
      "perda",
      "motivo perda",
      "justificativa perda",
      "comentario",
    ];
    const example = [
      "ID-188042-2379636-20260427-409",
      "910123456",
      "Inspeção preventiva em subestação",
      "EVT-ENE-04",
      "ELV-SE-0207",
      "2026-06",
      "2026-06-18",
      "2026-06-16",
      "2026-06-20",
      "Conflito com janela operacional",
      "Atendimento agrupado por rota",
      "Não",
      "",
      "",
      "Carga modelo",
    ];
    downloadFile("modelo-carga-eletrovia.csv", toCsv([header, example]));
  }

  function downloadRealizedTemplate() {
    const header = [
      "ordem",
      "data realizada",
      "perda",
      "motivo perda",
      "justificativa perda",
      "comentario",
    ];
    const example = [
      "910123456",
      "2026-06-18",
      "Não",
      "",
      "",
      "Realizado sincronizado pelo SAP BO",
    ];
    downloadFile("modelo-realizados-eletrovia.csv", toCsv([header, example]));
  }

  function bindEvents() {
    $("#loginForm").addEventListener("submit", handleLogin);
    $("#logoutButton").addEventListener("click", logout);

    $("#collapseSidebar").addEventListener("click", () => {
      document.body.classList.toggle("sidebar-collapsed");
    });

    $("#alertButton").addEventListener("click", () => {
      $("#alertMenu").classList.toggle("hidden");
    });
    $("#alertMenu").addEventListener("click", (event) => {
      const openPanel = event.target.closest("[data-alert-open-panel]");
      if (openPanel) {
        $("#alertMenu").classList.add("hidden");
        switchView("notificacoes");
        return;
      }

      const notificationButton = event.target.closest(
        "[data-alert-notification]",
      );
      if (notificationButton) {
        $("#alertMenu").classList.add("hidden");
        openNotificationAction(notificationButton.dataset.alertNotification);
        return;
      }

      const demandButton = event.target.closest("[data-alert-demand]");
      if (demandButton) {
        state.selectedDemandId = demandButton.dataset.alertDemand;
        $("#alertMenu").classList.add("hidden");
        switchView("carteira");
        renderCarteira();
      }
    });

    $("#mainNav").addEventListener("click", (event) => {
      const groupToggle = event.target.closest("[data-nav-group-toggle]");
      if (groupToggle?.dataset.navGroupToggle) {
        const groupKey = groupToggle.dataset.navGroupToggle;
        const nextOpen = !state.navGroups[groupKey];

        Object.keys(state.navGroups).forEach((key) => {
          state.navGroups[key] = key === groupKey ? nextOpen : false;
        });

        if (nextOpen && navGroupForView(state.currentView) !== groupKey) {
          switchView(groupToggle.dataset.view || NAV_GROUPS[groupKey]?.[0]);
          return;
        }

        syncNavigation(state.currentView);
        return;
      }

      const button = event.target.closest("[data-view]");
      if (button) switchView(button.dataset.view);
    });

    $("#userSelect").addEventListener("change", async (event) => {
      state.currentUser = state.db.usuarios.find(
        (user) => user.email === event.target.value,
      );
      global.localStorage.setItem("cce.currentUser", state.currentUser.email);
      renderRole();
      renderCurrentView();
      await state.repo.addLog({
        usuario: state.currentUser.email,
        acao: "Troca de Perfil",
        lista: "Usuarios_Central_Eletrovia",
        referencia: state.currentUser.perfil,
        detalhe: "Perfil selecionado no protótipo operacional.",
      });
    });

    $(".filter-panel").addEventListener("change", (event) => {
      if (
        !event.target.matches("[data-filter]") &&
        !event.target.matches("[data-multi-option]")
      ) {
        return;
      }

      state.page = 1;
      collectFilters();
      updateMultiFilterSummary(event.target.closest("[data-multi-filter]"));
      renderCarteira();
    });

    $(".filter-panel").addEventListener("input", (event) => {
      if (event.target.matches("[data-multi-search]")) {
        const query = normalizeText(event.target.value);
        const menu = event.target.closest(".multi-menu");

        $$(".multi-option", menu).forEach((option) => {
          option.classList.toggle(
            "hidden",
            Boolean(query) &&
              !normalizeText(option.textContent).includes(query),
          );
        });

        return;
      }

      if (event.target.id === "quickSearch") {
        state.page = 1;
        global.clearTimeout(state.filterSearchTimer);
        state.filterSearchTimer = global.setTimeout(renderCarteira, 160);
      }
    });

    $(".filter-panel").addEventListener("click", (event) => {
      const selectVisibleButton = event.target.closest(
        "[data-multi-select-visible]",
      );
      const clearButton = event.target.closest("[data-multi-clear]");

      if (!selectVisibleButton && !clearButton) return;

      const menu = event.target.closest(".multi-menu");
      if (!menu) return;

      event.preventDefault();
      event.stopPropagation();

      const visibleOptions = $$("[data-multi-option]", menu).filter(
        (option) =>
          !option.closest(".multi-option")?.classList.contains("hidden"),
      );

      if (selectVisibleButton) {
        visibleOptions.forEach((option) => {
          option.checked = true;
        });
      }

      if (clearButton) {
        $$("[data-multi-option]", menu).forEach((option) => {
          option.checked = false;
        });
      }

      state.page = 1;
      collectFilters();
      updateMultiFilterSummary(event.target.closest("[data-multi-filter]"));
      renderCarteira();
    });
    $("#clearFilters").addEventListener("click", () => {
      $$("[data-filter]").forEach((field) => {
        field.value = "";
      });
      $$("[data-multi-option]").forEach((field) => {
        field.checked = false;
      });
      $("#quickSearch").value = "";
      state.page = 1;
      collectFilters();
      buildFilterOptions();
      updateAllMultiFilterSummaries($(".filter-panel"));
      renderCarteira();
    });
    $("#toggleAdvancedFilters").addEventListener("click", () => {
      state.advancedFilters = !state.advancedFilters;
      $$(".advanced-filter").forEach((element) =>
        element.classList.toggle("hidden", !state.advancedFilters),
      );
      if (state.advancedFilters) {
        collectFilters();
        buildFilterOptions({ includeHidden: true });
      }
      $("#toggleAdvancedFilters").innerHTML =
        `${iconSvg("filter")} ${state.advancedFilters ? "Ocultar avançados" : "Filtros avançados"}`;
    });
    $("#pageSize").addEventListener("change", (event) => {
      state.pageSize = Number(event.target.value);
      state.page = 1;
      renderCarteira();
    });
    $("#prevPage").addEventListener("click", () => {
      state.page = Math.max(1, state.page - 1);
      renderCarteira();
    });
    $("#nextPage").addEventListener("click", () => {
      state.page += 1;
      renderCarteira();
    });
    $("#demandTableBody").addEventListener("click", (event) => {
      const actionButton = event.target.closest("[data-action]");
      if (actionButton) {
        openAction(actionButton.dataset.action, actionButton.dataset.id);
        return;
      }
      const row = event.target.closest("[data-demand-id]");
      if (row) {
        state.selectedDemandId = row.dataset.demandId;
        renderCarteira();
      }
    });
    $("#detailPanel").addEventListener("click", (event) => {
      const actionButton = event.target.closest("[data-action]");
      if (actionButton)
        openAction(actionButton.dataset.action, actionButton.dataset.id);
    });
    $("#modalSave").addEventListener("click", saveAction);
    $("#exportCsv").addEventListener("click", exportCurrentCarteira);
    $("#refreshData").addEventListener("click", async () => {
      await state.repo.reset();
      await refreshAll();
      showToast("Dados atualizados do Supabase e JSON.", "success");
    });
    $("#validateBatch").addEventListener("click", validateBatchFile);
    $("#clearBatch").addEventListener("click", () => {
      state.batch = {
        rows: [],
        valid: [],
        warnings: [],
        errors: [],
        fileName: "",
      };
      $("#batchFile").value = "";
      renderBatch();
    });
    $("#saveValidBatch").addEventListener("click", () => saveBatch(false));
    $("#saveConfirmedBatch").addEventListener("click", () => saveBatch(true));
    $("#downloadTemplate").addEventListener("click", downloadTemplate);
    $("#futureDemandForm").addEventListener("submit", createFutureDemand);
    $("#futureDemandList").addEventListener("click", (event) => {
      const suggestionButton = event.target.closest("[data-load-suggestions]");

      if (suggestionButton) {
        renderFutureSuggestions(suggestionButton.dataset.loadSuggestions);
        return;
      }

      const linkButton = event.target.closest("[data-link-future]");

      if (linkButton) {
        linkFutureDemand(
          linkButton.dataset.linkFuture,
          linkButton.dataset.linkTarget,
        );
      }
    });
    $("#qualityTypeFilter").addEventListener("change", (event) => {
      state.quality.typeFilter = event.target.value;
      state.quality.selectedIssueId = "";
      state.quality.selectedPrimarySequence = "";
      state.quality.page = 1;
      renderQuality();
    });

    $("#qualitySearch").addEventListener("input", (event) => {
      state.quality.search = event.target.value;
      state.quality.selectedIssueId = "";
      state.quality.selectedPrimarySequence = "";
      state.quality.page = 1;
      global.clearTimeout(state.quality.searchTimer);
      state.quality.searchTimer = global.setTimeout(renderQuality, 180);
    });

    $("#qualityClearFilters").addEventListener("click", () => {
      state.quality.typeFilter = "";
      state.quality.search = "";
      state.quality.selectedIssueId = "";
      state.quality.selectedPrimarySequence = "";
      state.quality.page = 1;
      renderQuality();
    });
    $("#qualityCards").addEventListener("click", (event) => {
      const card = event.target.closest("[data-quality-card]");
      if (!card) return;
      state.quality.typeFilter =
        state.quality.typeFilter === card.dataset.qualityCard
          ? ""
          : card.dataset.qualityCard;
      state.quality.selectedIssueId = "";
      state.quality.selectedPrimarySequence = "";
      state.quality.page = 1;
      renderQuality();
    });
    $("#qualityIssueTableBody").addEventListener("click", (event) => {
      const target =
        event.target.closest("[data-quality-detail]") ||
        event.target.closest("[data-quality-issue-id]");

      if (!target) return;

      selectQualityIssue(
        target.dataset.qualityDetail || target.dataset.qualityIssueId,
        Boolean(target.dataset.qualityDetail),
      );
    });
    $("#qualityDetailPanel").addEventListener("change", (event) => {
      if (!event.target.matches('input[name="qualityPrimaryRecord"]')) return;

      state.quality.selectedPrimarySequence = event.target.value;
      renderQualityDetail(selectedQualityIssue());
    });

    $("#qualityDetailPanel").addEventListener("click", async (event) => {
      const primaryCard = event.target.closest("[data-quality-primary]");
      if (primaryCard && !event.target.closest("[data-quality-action]")) {
        state.quality.selectedPrimarySequence =
          primaryCard.dataset.qualityPrimary || "";
        renderQualityDetail(selectedQualityIssue());
        return;
      }

      const action = event.target.closest("[data-quality-action]")?.dataset
        .qualityAction;

      if (!action) return;

      if (action === "set-primary") {
        const checked = $(
          'input[name="qualityPrimaryRecord"]:checked',
          $("#qualityDetailPanel"),
        );

        if (!checked) {
          showToast("Selecione um registro principal.", "error");
          return;
        }

        state.quality.selectedPrimarySequence = checked.value;
        renderQualityDetail(selectedQualityIssue());
        showToast("Registro principal definido.", "success");
      }

      if (action === "merge-preview") {
        showQualityMergePreview();
      }

      if (action === "save-merge") {
        await saveQualityMerge();
      }
    });
    $("#duplicateOmSearch")?.addEventListener("input", (event) => {
      state.quality.duplicateOmSearch = event.target.value;
      global.clearTimeout(state.quality.duplicateOmSearchTimer);
      state.quality.duplicateOmSearchTimer = global.setTimeout(
        renderDuplicateOms,
        160,
      );
    });

    $("#duplicateOmClearSearch")?.addEventListener("click", () => {
      state.quality.duplicateOmSearch = "";
      renderDuplicateOms();
    });

    $("#duplicateOmTableBody")?.addEventListener("click", (event) => {
      const button = event.target.closest("[data-open-quality-issue]");
      if (!button) return;
      openQualityIssueFromFocusedView(button.dataset.openQualityIssue);
    });

    $("#qualityLocalSearch")?.addEventListener("input", (event) => {
      state.quality.localSearch = event.target.value;
      global.clearTimeout(state.quality.localSearchTimer);
      state.quality.localSearchTimer = global.setTimeout(
        renderQualityByLocal,
        160,
      );
    });

    $("#qualityLocalClearSearch")?.addEventListener("click", () => {
      state.quality.localSearch = "";
      state.quality.selectedLocal = "";
      renderQualityByLocal();
    });

    $("#toggleQualityLocalFilters")?.addEventListener("click", () => {
      state.quality.localFiltersVisible = !state.quality.localFiltersVisible;
      if (state.quality.localFiltersVisible && !state.quality.localFiltersReady) {
        collectQualityLocalFilters();
        buildQualityLocalFilterOptions();
      }
      renderQualityLocalFilterVisibility();
    });

    $("#qualityLocalFilterPanel")?.addEventListener("change", (event) => {
      if (!event.target.matches("[data-multi-option]")) return;
      collectQualityLocalFilters();
      buildQualityLocalFilterOptions();
      state.quality.selectedLocal = "";
      renderQualityByLocal();
    });

    $("#qualityLocalFilterPanel")?.addEventListener("input", (event) => {
      if (event.target.matches("[data-multi-search]")) {
        const query = normalizeText(event.target.value);
        const menu = event.target.closest(".multi-menu");

        $$(".multi-option", menu).forEach((option) => {
          option.classList.toggle(
            "hidden",
            Boolean(query) &&
              !normalizeText(option.textContent).includes(query),
          );
        });
        return;
      }

      if (event.target.id === "qualityLocalQuickSearch") {
        collectQualityLocalFilters();
        state.quality.selectedLocal = "";
        global.clearTimeout(state.quality.localFilterSearchTimer);
        state.quality.localFilterSearchTimer = global.setTimeout(
          renderQualityByLocal,
          160,
        );
      }
    });

    $("#qualityLocalFilterPanel")?.addEventListener("click", (event) => {
      const selectVisibleButton = event.target.closest(
        "[data-multi-select-visible]",
      );
      const clearButton = event.target.closest("[data-multi-clear]");

      if (!selectVisibleButton && !clearButton) return;

      const menu = event.target.closest(".multi-menu");
      if (!menu) return;

      event.preventDefault();
      event.stopPropagation();

      const visibleOptions = $$("[data-multi-option]", menu).filter(
        (option) =>
          !option.closest(".multi-option")?.classList.contains("hidden"),
      );

      if (selectVisibleButton) {
        visibleOptions.forEach((option) => {
          option.checked = true;
        });
      }

      if (clearButton) {
        $$("[data-multi-option]", menu).forEach((option) => {
          option.checked = false;
        });
      }

      collectQualityLocalFilters();
      buildQualityLocalFilterOptions();
      state.quality.selectedLocal = "";
      renderQualityByLocal();
    });

    $("#qualityLocalClearFilters")?.addEventListener("click", () => {
      $$("[data-multi-option]", $("#qualityLocalFilterPanel")).forEach(
        (field) => {
          field.checked = false;
        },
      );
      $("#qualityLocalQuickSearch").value = "";
      state.quality.localFilters = {};
      state.quality.localGroupsCache = null;
      state.quality.localFiltersReady = false;
      state.quality.selectedLocal = "";
      buildQualityLocalFilterOptions();
      renderQualityByLocal();
    });

    $("#qualityLocalGroupTableBody")?.addEventListener("click", (event) => {
      const target =
        event.target.closest("[data-select-quality-local]") ||
        event.target.closest("[data-quality-local]");
      const local =
        target?.dataset.selectQualityLocal || target?.dataset.qualityLocal;
      if (!local) return;
      state.quality.selectedLocal = local;
      renderQualityByLocal();
      $("#qualityLocalDemandPanel")?.scrollIntoView({
        behavior: "smooth",
        block: "nearest",
      });
    });

    $("#qualityLocalDemandTableBody")?.addEventListener("click", (event) => {
      const row = event.target.closest("[data-local-demand]");
      if (!row) return;
      state.selectedDemandId = row.dataset.localDemand;
      switchView("carteira");
      renderCarteira();
    });

    $("#qualityCentersSearch")?.addEventListener("input", (event) => {
      state.quality.centersSearch = event.target.value;
      global.clearTimeout(state.quality.centersSearchTimer);
      state.quality.centersSearchTimer = global.setTimeout(
        renderQualityCenters,
        160,
      );
    });

    $("#qualityCentersClearSearch")?.addEventListener("click", () => {
      state.quality.centersSearch = "";
      renderQualityCenters();
    });

    $("#qualityCentersTableBody")?.addEventListener("click", (event) => {
      const button = event.target.closest("[data-open-center-admin]");
      if (!button) return;
      state.adminTab = "centrosTrabalho";
      switchView("administracao");
      renderAdmin();
    });

    $("#qualityDivergenceSearch")?.addEventListener("input", (event) => {
      state.quality.divergenceSearch = event.target.value;
      global.clearTimeout(state.quality.divergenceSearchTimer);
      state.quality.divergenceSearchTimer = global.setTimeout(
        renderQualityDivergences,
        160,
      );
    });

    $("#qualityDivergenceClearSearch")?.addEventListener("click", () => {
      state.quality.divergenceSearch = "";
      renderQualityDivergences();
    });
    $("#notificationTypeFilter")?.addEventListener("change", (event) => {
      state.notifications.typeFilter = event.target.value;
      renderNotifications();
    });

    $("#notificationSearch")?.addEventListener("input", (event) => {
      state.notifications.search = event.target.value;
      renderNotifications();
    });

    $("#notificationClearFilters")?.addEventListener("click", () => {
      state.notifications.typeFilter = "";
      state.notifications.search = "";
      renderNotifications();
    });

    $("#notificationCards")?.addEventListener("click", (event) => {
      const card = event.target.closest("[data-notification-card]");
      if (!card) return;

      const type = card.dataset.notificationCard;

      state.notifications.typeFilter = type === "total" ? "" : type;
      renderNotifications();
    });

    $("#notificationTableBody")?.addEventListener("click", (event) => {
      const actionButton = event.target.closest("[data-notification-action]");
      if (actionButton) {
        openNotificationAction(actionButton.dataset.notificationAction);
        return;
      }

      const openButton = event.target.closest("[data-notification-open]");
      if (openButton) {
        openDemandFromNotification(openButton.dataset.notificationOpen);
      }
    });
    $("#portfolioHistorySearch")?.addEventListener("input", (event) => {
      state.portfolioHistory.search = event.target.value;
      global.clearTimeout(state.portfolioHistory.searchTimer);
      state.portfolioHistory.searchTimer = global.setTimeout(
        renderPortfolioHistory,
        160,
      );
    });

    $("#portfolioHistoryAction")?.addEventListener("change", (event) => {
      state.portfolioHistory.action = event.target.value;
      renderPortfolioHistory();
    });

    $("#portfolioHistoryUser")?.addEventListener("input", (event) => {
      state.portfolioHistory.user = event.target.value;
      global.clearTimeout(state.portfolioHistory.userTimer);
      state.portfolioHistory.userTimer = global.setTimeout(
        renderPortfolioHistory,
        160,
      );
    });

    $("#portfolioHistoryStartDate")?.addEventListener("change", (event) => {
      state.portfolioHistory.startDate = event.target.value;
      renderPortfolioHistory();
    });

    $("#portfolioHistoryEndDate")?.addEventListener("change", (event) => {
      state.portfolioHistory.endDate = event.target.value;
      renderPortfolioHistory();
    });

    $("#portfolioHistoryClearFilters")?.addEventListener("click", () => {
      state.portfolioHistory.search = "";
      state.portfolioHistory.action = "";
      state.portfolioHistory.user = "";
      state.portfolioHistory.startDate = "";
      state.portfolioHistory.endDate = "";
      renderPortfolioHistory();
    });

    $("#portfolioHistoryTableBody")?.addEventListener("click", (event) => {
      const button = event.target.closest("[data-history-demand]");
      if (!button?.dataset.historyDemand) return;

      state.selectedDemandId = button.dataset.historyDemand;
      switchView("carteira");
      renderCarteira();
    });
    $("#toggleIndicatorFilters")?.addEventListener("click", () => {
      state.indicatorFiltersVisible = !state.indicatorFiltersVisible;
      if (state.indicatorFiltersVisible && !state.indicatorFiltersReady) {
        collectIndicatorFilters();
        buildIndicatorFilterOptions();
      }
      renderIndicatorFilterVisibility();
    });
    $("#indicatorFilterPanel")?.addEventListener("change", (event) => {
      if (!event.target.matches("[data-multi-option]")) return;
      collectIndicatorFilters();
      buildIndicatorFilterOptions();
      renderIndicators();
    });
    $("#indicatorFilterPanel")?.addEventListener("input", (event) => {
      if (event.target.matches("[data-multi-search]")) {
        const query = normalizeText(event.target.value);
        const menu = event.target.closest(".multi-menu");

        $$(".multi-option", menu).forEach((option) => {
          option.classList.toggle(
            "hidden",
            Boolean(query) &&
              !normalizeText(option.textContent).includes(query),
          );
        });
        return;
      }

      if (event.target.id === "indicatorQuickSearch") {
        collectIndicatorFilters();
        global.clearTimeout(state.indicatorSearchTimer);
        state.indicatorSearchTimer = global.setTimeout(renderIndicators, 160);
      }
    });
    $("#indicatorFilterPanel")?.addEventListener("click", (event) => {
      const selectVisibleButton = event.target.closest(
        "[data-multi-select-visible]",
      );
      const clearButton = event.target.closest("[data-multi-clear]");

      if (!selectVisibleButton && !clearButton) return;

      const menu = event.target.closest(".multi-menu");
      if (!menu) return;

      event.preventDefault();
      event.stopPropagation();

      const visibleOptions = $$("[data-multi-option]", menu).filter(
        (option) =>
          !option.closest(".multi-option")?.classList.contains("hidden"),
      );

      if (selectVisibleButton) {
        visibleOptions.forEach((option) => {
          option.checked = true;
        });
      }

      if (clearButton) {
        $$("[data-multi-option]", menu).forEach((option) => {
          option.checked = false;
        });
      }

      collectIndicatorFilters();
      buildIndicatorFilterOptions();
      renderIndicators();
    });
    $("#indicatorClearFilters")?.addEventListener("click", () => {
      $$("[data-multi-option]", $("#indicatorFilterPanel")).forEach((field) => {
        field.checked = false;
      });
      $("#indicatorQuickSearch").value = "";
      state.indicatorFilters = {};
      collectIndicatorFilters();
      buildIndicatorFilterOptions();
      renderIndicators();
    });
    $("#adminTabs").addEventListener("click", (event) => {
      const button = event.target.closest("[data-admin-tab]");
      if (!button) return;
      state.adminTab = button.dataset.adminTab;
      renderAdmin();
    });
    $("#exportLogs").addEventListener("click", () => {
      const rows = state.db.logs.map((log) => [
        log.dataHora,
        log.usuario,
        log.acao,
        log.lista,
        log.referencia,
        log.detalhe,
      ]);
      downloadFile(
        `logs-eletrovia-${todayText()}.csv`,
        toCsv([
          ["Data/Hora", "Usuário", "Ação", "Lista", "Referência", "Detalhe"],
          ...rows,
        ]),
      );
    });
  }

  async function init() {
    renderStaticIcons();

    $("#toggleAdvancedFilters").innerHTML =
      `${iconSvg("filter")} Filtros avançados`;

    state.repo = global.CCEData.createRepository();

    bindEvents();

    renderLoginState();

    try {
      await loadLoginData();
    } catch (error) {
      console.error(error);

      setLoginUiState({
        ready: false,
        loading: false,
        buttonText: "Falha no acesso",
        statusText: "Não foi possível carregar a base de usuários.",
        errorText: error.message,
      });
    }
  }

  document.addEventListener("DOMContentLoaded", init);
})(window, document);

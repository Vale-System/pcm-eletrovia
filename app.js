(function bootstrapCentral(global, document) {
  const $ = (selector, root = document) => root.querySelector(selector);
  const $$ = (selector, root = document) =>
    Array.from(root.querySelectorAll(selector));

  const STATUS_OPTIONS = [
    "A Planejar",
    "Planejado",
    "Replanejado",
    "Realizado",
  ];

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
    { key: "planejadorOM", label: "Planejador OM", field: "planejadorOM" },
    { key: "programador", label: "Programador", field: "programador" },
    { key: "centroStatus", label: "Cadastro Centro", special: "centroStatus" },
    {
      key: "localInstalacao",
      label: "Local Instalacao",
      field: "localInstalacao",
    },
    { key: "statusSistema", label: "Status Sistema", field: "statusSistema" },
    { key: "anoVencimento", label: "Ano Vencimento", special: "anoVencimento" },
    { key: "mesVencimento", label: "Mes Vencimento", special: "mesVencimento" },
  ];

  const state = {
    repo: null,
    db: null,
    currentUser: null,
    currentView: "carteira",
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
    lastDataUpdateAt: "",
    loginReady: false,
    batch: {
      rows: [],
      valid: [],
      warnings: [],
      errors: [],
      fileName: "",
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

  function primaryStatusOf(demand) {
    if (demand.dataRealizada) {
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
    const substatuses = [];
    if (demand.perda) substatuses.push("Perda");
    const dueClass = dueClassOf(demand);
    if (dueClass) substatuses.push(dueClass);
    if (pendingIssuesOf(demand).length) substatuses.push("Pendente");
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
    if (status === "Cadastrado") return "status-realizado";
    if (status === "Nao cadastrado" || status === "Sem centro")
      return "status-perda";
    if (status === "Fora do Prazo") return "status-fora-prazo";
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
      realizado: true,
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

    const idDemanda =
      item.ID_Demanda_Controle ||
      item.id_demanda_controle ||
      item.IdDemandaControle ||
      item.id ||
      (ordem
        ? `DEM-SAP-${ordem}`
        : global.CCEData.stableDemandId({
            ordem: "",
            descricao: item.Descricao || item.descricao || "",
            centroTrabalho: item.CentroTrabalho || item.centro_trabalho || "",
            localInstalacao:
              item.LocalInstalacao || item.local_instalacao || "",
            competencia: item.Competencia || item.competencia || "",
            vencimento: item.Vencimento || item.vencimento || "",
            origem: origemPadrao,
          }));

    return {
      id: String(idDemanda).trim(),

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

      competencia: normalizeCompetencia(item.Competencia || item.competencia),
      dataRealizada: item.DataRealizada || item.data_realizada || "",
      vencimento: item.Vencimento || item.vencimento || "",

      prioridade: normalizePrioridade(item.Prioridade || item.prioridade),
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

    // Primeiro entram as futuras.
    // Depois as ordens entram por cima, caso tenham o mesmo ID.
    baseFuturas.forEach((item) => {
      if (!item.id) return;
      mapa.set(item.id, item);
    });

    baseOrdens.forEach((item) => {
      if (!item.id) return;

      const futura = mapa.get(item.id);

      if (!futura) {
        mapa.set(item.id, item);
        return;
      }

      // Se existir nas duas, a ordem SAP/base ordens vence,
      // mas preserva campos úteis da futura quando a ordem estiver vazia.
      mapa.set(item.id, {
        ...futura,
        ...item,

        id: item.id,
        ordem: item.ordem || futura.ordem || "",

        tipoDemanda: item.tipoDemanda || futura.tipoDemanda || "",
        frequencia: item.frequencia || futura.frequencia || "",
        observacao: item.observacao || futura.observacao || "",
        vinculadaEm: item.vinculadaEm || futura.vinculadaEm || "",

        origem: item.origem || "SAP BO - Ordens",
      });
    });

    return Array.from(mapa.values());
  }

  async function loadBaseFromJson() {
    const [baseOrdensRaw, baseFuturasRaw] = await Promise.all([
      fetchJsonArray("./base/base_ordens.json", "base_ordens.json"),
      fetchJsonArray("./base/base_futuras.json", "base_futuras.json").catch(
        (error) => {
          console.warn("base_futuras.json não carregada:", error);
          return [];
        },
      ),
    ]);

    const baseOrdens = baseOrdensRaw.map((item) =>
      mapBaseItemToDemand(item, "SAP BO - Ordens"),
    );

    const baseFuturas = baseFuturasRaw.map((item) =>
      mapBaseItemToDemand(item, "SAP BO - Demandas Futuras"),
    );

    return mergeBaseOrdensEFuturas(baseOrdens, baseFuturas);
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
      vencimento: baseDemand.vencimento || delta.vencimento,

      dataPlanejada: delta.dataPlanejada || baseDemand.dataPlanejada || "",
      dataReplanejadaAtual:
        delta.dataReplanejadaAtual || baseDemand.dataReplanejadaAtual || "",
      dataRealizada: delta.dataRealizada || baseDemand.dataRealizada || "",

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

  function enrichDemandWithCentroTrabalho(demanda, mapaCentros) {
    const chaveCentro = normalizeCentroTrabalho(demanda.centroTrabalho);
    const cadastro = mapaCentros.get(chaveCentro);

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

      planejadorOM: cadastro.planejadorOM || "",
      planejadorOMEmail: cadastro.planejadorOMEmail || "",
      planejadorOMMatricula: cadastro.planejadorOMMatricula || "",

      programador: cadastro.programador || "",
      programadorEmail: cadastro.programadorEmail || "",
      programadorMatricula: cadastro.programadorMatricula || "",

      centroTrabalhoChave: cadastro.centroTrabalhoChave || chaveCentro,
      centroTrabalhoCadastrado: true,
      centroTrabalhoStatus: "Cadastrado",
    };
  }

  async function loadDatabase() {
    const previousSelection = state.selectedDemandId;
    const previousUserEmail =
      state.currentUser?.email || getStoredSessionEmail();
    const base = await loadBaseFromJson();
    const supabaseData = await state.repo.getAll();
    const mapaCentrosTrabalho = new Map(
      (supabaseData.centrosTrabalho || [])
        .filter((item) => item.ativo !== false)
        .map((item) => [
          item.centroTrabalhoChave ||
            normalizeCentroTrabalho(item.centroTrabalho),
          item,
        ]),
    );

    const baseEnriquecida = base.map((demanda) =>
      enrichDemandWithCentroTrabalho(demanda, mapaCentrosTrabalho),
    );

    const deltasById = new Map(
      (supabaseData.demandas || []).map((item) => [item.id, item]),
    );

    const baseIds = new Set(baseEnriquecida.map((item) => item.id));

    const mergedBase = baseEnriquecida.map((item) =>
      enrichDemandWithCentroTrabalho(
        mergeDemandWithSupabase(item, deltasById.get(item.id)),
        mapaCentrosTrabalho,
      ),
    );

    const demandasSomenteSupabase = (supabaseData.demandas || [])
      .filter((item) => !baseIds.has(item.id))
      .map((item) => enrichDemandWithCentroTrabalho(item, mapaCentrosTrabalho));

    const demandas = [...demandasSomenteSupabase, ...mergedBase].map(
      normalizeDemandRecord,
    );

    state.db = {
      demandas,
      usuarios: supabaseData.usuarios || [],
      centrosTrabalho: supabaseData.centrosTrabalho || [],
      configuracoes: supabaseData.configuracoes || {},
      parametros: supabaseData.parametros || {},
      parametrosDisponiveis: supabaseData.parametrosDisponiveis === true,
      historicoPlanejamento: supabaseData.historicoPlanejamento || [],
      historicoReplanejamento: supabaseData.historicoReplanejamento || [],
      historicoRealizadoPerdas: supabaseData.historicoRealizadoPerdas || [],
      logs: supabaseData.logs || [],
    };

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
      centroTrabalhoChave:
        demanda.centroTrabalhoChave ||
        normalizeCentroTrabalho(demanda.centroTrabalho),
    };
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
    renderAlerts();
  }

  function renderRole() {
    $("#userName").textContent =
      state.currentUser?.nome || state.identity?.nome || "Aguardando login";
    $("#roleChip").textContent = state.currentUser?.perfil || "Visualizador";
    $("#logoutButton")?.classList.toggle("hidden", !state.currentUser);
    applyPermissions();
  }

  function renderLoginState() {
    const loginScreen = $("#loginScreen");
    if (!loginScreen) return;
    loginScreen.classList.toggle("hidden", Boolean(state.currentUser));
    document.body.classList.toggle("is-login-required", !state.currentUser);
    $("#loginError").textContent = "";
  }

  async function handleLogin(event) {
    event.preventDefault();
    const form = new FormData(event.currentTarget);
    const email = String(form.get("email") || "")
      .trim()
      .toLowerCase();
    const matricula = String(form.get("matricula") || "").trim();
    const user = state.db.usuarios.find(
      (item) => normalizeText(item.email) === normalizeText(email),
    );
    const errorElement = $("#loginError");

    if (!user || user.ativo === false) {
      errorElement.textContent = "Usuario nao encontrado ou inativo.";
      await state.repo.addLog?.({
        usuario: email,
        acao: "Login",
        lista: "usuarios_central_eletrovia",
        referencia: email,
        detalhe: "Tentativa de login com usuario inexistente ou inativo.",
        nivel: "AVISO",
        modulo: "LOGIN",
        status: "ERRO",
      });
      return;
    }

    if (String(user.matricula || "").trim() !== matricula) {
      errorElement.textContent = "Matricula invalida para o e-mail informado.";
      await state.repo.addLog?.({
        usuario: email,
        acao: "Login",
        lista: "usuarios_central_eletrovia",
        referencia: email,
        detalhe: "Tentativa de login com matricula invalida.",
        nivel: "AVISO",
        modulo: "LOGIN",
        status: "ERRO",
      });
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
    event.currentTarget.reset();
    renderLoginState();
    renderRole();
    renderCurrentView();
    await state.repo.addLog?.({
      usuario: user.email,
      acao: "Login",
      lista: "usuarios_central_eletrovia",
      referencia: user.email,
      detalhe: "Login realizado com e-mail e matricula.",
      modulo: "LOGIN",
      status: "SUCESSO",
    });
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
    $$('[data-view="administracao"]').forEach((element) => {
      element.disabled = !rules.configurar;
      element.title = rules.configurar ? "" : "Disponivel para Administrador";
    });
    $$('[data-view="lote"]').forEach((element) => {
      element.disabled = !rules.cargaLote;
      element.title = rules.cargaLote ? "" : "Sem permissao para carga em lote";
    });
    $("#exportCsv").disabled = !canExport();
  }

  function buildFilterOptions() {
    const filters = state.filters || {};
    FILTER_DEFINITIONS.forEach((definition) => {
      const host = $(`[data-multi-filter="${definition.key}"]`);
      if (!host) return;
      const scopedRows = state.db.demandas.filter((item) =>
        demandMatchesFilters(item, filters, definition.key),
      );
      const selected = filters[definition.key] || [];
      const options = uniqueOptions([
        ...scopedRows.map((item) => filterValueFor(item, definition)),
        ...selected,
      ]);
      renderMultiFilter(host, definition, options, selected);
    });
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

  function filterValueFor(item, definition) {
    if (definition.special === "status") return primaryStatusOf(item);
    if (definition.special === "centroStatus")
      return item.centroTrabalhoStatus || "Nao cadastrado";
    if (definition.special === "anoVencimento")
      return item.vencimento?.slice(0, 4) || "";
    if (definition.special === "mesVencimento") {
      const month = item.vencimento?.slice(5, 7) || "";
      return month ? `${month} - ${monthName(month)}` : "";
    }
    return item[definition.field] || "";
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
    host.innerHTML = `
      <label class="multi-label">${escapeHtml(definition.label)}</label>
      <details class="multi-select" data-filter-details>
        <summary title="${escapeHtml(summary)}">${escapeHtml(summary)}</summary>
        <div class="multi-menu">
          <input data-multi-search type="search" placeholder="Pesquisar..." aria-label="Pesquisar ${escapeHtml(definition.label)}" />
          <div class="multi-options">
            ${
              options.length
                ? options
                    .map(
                      (option) => `
                <label class="multi-option">
                  <input data-multi-option type="checkbox" value="${escapeHtml(option)}" ${selectedSet.has(option) ? "checked" : ""} />
                  <span>${escapeHtml(option)}</span>
                </label>
              `,
                    )
                    .join("")
                : '<span class="muted multi-empty">Sem opcoes no recorte</span>'
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
      if (
        selected.length &&
        !selected.includes(String(filterValueFor(item, definition)))
      )
        return false;
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
      const haystack = normalizeText(
        [
          item.id,
          item.ordem,
          item.descricao,
          item.centroTrabalho,
          item.localInstalacao,
        ].join(" "),
      );
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

  function alertItems() {
    const today = toDate(todayText());
    return state.db.demandas
      .map((item) => {
        const due = toDate(item.vencimento);
        const days = due ? Math.ceil((due - today) / 86400000) : 9999;
        const pending = pendingIssuesOf(item);
        if (pending.length) {
          return { item, type: "Pendente", message: pending.join(" | ") };
        }
        if (!item.dataRealizada && days >= 0 && days <= 7) {
          return {
            item,
            type: "Vencimento",
            message: `Vence em ${days} dia${days === 1 ? "" : "s"}.`,
          };
        }
        return null;
      })
      .filter(Boolean)
      .slice(0, 30);
  }

  function renderAlerts() {
    const alerts = alertItems();
    $("#alertCount").textContent = String(alerts.length);
    $("#alertButton").classList.toggle("has-alerts", alerts.length > 0);
    $("#alertMenu").innerHTML =
      `<strong>Alertas operacionais</strong>` +
      (alerts.length
        ? alerts
            .slice(0, 8)
            .map(
              ({ item, type, message }) => `
              <button type="button" data-alert-demand="${escapeHtml(item.id)}">
                <span>${escapeHtml(type)}</span>
                <strong>${escapeHtml(item.ordem || item.id)}</strong>
                <small>${escapeHtml(message)}</small>
              </button>
            `,
            )
            .join("")
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
      ["Aderência", `${stats.aderencia}%`, "realizado no prazo"],
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
            <td class="description-cell">${escapeHtml(item.descricao)}<div class="muted">${escapeHtml(item.usuarioResponsavel || "")}</div></td>
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
            <td>${escapeHtml(item.planejadorOM || "-")}</td>
            <td>${escapeHtml(item.programador || "-")}</td>
            <td>${escapeHtml(item.localInstalacao || "-")}</td>
            <td>${escapeHtml(item.prioridade || "-")}</td>
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
        <div class="detail-item"><span>Substatus</span><strong>${statusChipGroup(substatuses)}</strong></div>
        <div class="detail-item"><span>Origem</span><strong>${escapeHtml(demand.origem || "-")}</strong></div>
        <div class="detail-item"><span>Ultima atualizacao</span><strong>${formatDateTime(demand.dataUltimaAtualizacao)}</strong></div>
        <div class="detail-item"><span>Gerencia</span><strong>${escapeHtml(demand.gerencia || "-")}</strong></div>
        <div class="detail-item"><span>Supervisao</span><strong>${escapeHtml(demand.supervisao || "-")}</strong></div>
        <div class="detail-item"><span>Vencimento</span><strong>${formatDate(demand.vencimento)}</strong></div>
        <div class="detail-item"><span>Centro</span><strong>${escapeHtml(demand.centroTrabalho || "-")}</strong></div>
        <div class="detail-item"><span>Cadastro centro</span><strong>${statusChip(demand.centroTrabalhoStatus || "Nao cadastrado")}</strong></div>
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
          <label>Data planejada<input name="dataPlanejada" type="date" value="${dateText(demand.dataPlanejada || demand.vencimento)}" required /></label>
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
          <label>Nova data<input name="novaData" type="date" value="${dateText(demand.dataReplanejadaAtual || demand.dataPlanejada || demand.vencimento)}" required /></label>
          <label>Motivo<select name="motivo" required>${optionsMarkup(configNames("motivos"))}</select></label>
          <label>Justificativa<select name="justificativa" required></select></label>
          <label class="span-2">Comentário<textarea name="comentario" rows="4">${escapeHtml(demand.comentario || "")}</textarea></label>
        </div>
      `;
      $("#modalSave").classList.remove("hidden");
    }

    if (action === "realizado") {
      const realizedOnlyLoss = primaryStatusOf(demand) === "Realizado";
      $("#modalTitle").textContent = realizedOnlyLoss
        ? "Complementar Perda"
        : "Registrar Realizado/Perda";
      $("#modalBody").innerHTML = `
        <div class="modal-grid">
          <label class="${realizedOnlyLoss ? "hidden" : ""}">Data realizada<input name="dataRealizada" type="date" value="${dateText(demand.dataRealizada || todayText())}" ${realizedOnlyLoss ? "disabled" : ""} /></label>
          <label>Perda<select name="perda"><option value="nao" ${!demand.perda ? "selected" : ""} ${realizedOnlyLoss ? "disabled" : ""}>Não</option><option value="sim" ${demand.perda || realizedOnlyLoss ? "selected" : ""}>Sim</option></select></label>
          <label>Perfil da perda<select name="motivoPerda"><option value=""></option>${optionsMarkup(configNames("perfisPerda"), demand.motivoPerda)}</select></label>
          <label>Justificativa perda<select name="justificativaPerda" data-selected="${escapeHtml(demand.justificativaPerda || "")}"></select></label>
          <label class="span-2">Comentário<textarea name="comentario" rows="4">${escapeHtml(demand.comentario || "")}</textarea></label>
          <label class="span-2">Evidência<input name="evidencia" placeholder="URL ou referência do anexo no SharePoint" /></label>
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

  async function saveAction() {
    const context = state.actionContext;
    if (!context) return;
    const demand = demandById(context.demandId);
    if (!demand) return;
    const form = new FormData($("#actionForm"));
    const userEmail = state.currentUser.email;

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
      const alreadyRealized = primaryStatusOf(demand) === "Realizado";
      const lossSelected = alreadyRealized ? true : form.get("perda") === "sim";
      const motivoPerda = form.get("motivoPerda") || "";
      const justificativaPerda = form.get("justificativaPerda") || "";
      if (lossSelected && (!motivoPerda || !justificativaPerda)) {
        showToast("Perda exige perfil e justificativa.", "error");
        return;
      }
      demand.dataRealizada = alreadyRealized
        ? demand.dataRealizada
        : form.get("dataRealizada") || "";
      demand.perda = lossSelected;
      demand.motivoPerda = form.get("motivoPerda") || "";
      demand.justificativaPerda = form.get("justificativaPerda") || "";
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
          ? demand.motivoPerda
          : `Realizado em ${demand.dataRealizada}`,
      });
    }

    $("#actionDialog").close();
    await refreshAll();
    showToast("Registro salvo com sucesso.", "success");
  }

  function renderCurrentView() {
    if (state.currentView === "carteira") renderCarteira();
    if (state.currentView === "lote") renderBatch();
    if (state.currentView === "futuras") renderFutureDemandas();
    if (state.currentView === "indicadores") renderIndicators();
    if (state.currentView === "administracao") renderAdmin();
    if (state.currentView === "logs") renderLogs();
    applyPermissions();
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
    if (view === "administracao" && !canAdmin()) {
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
    $$(".nav-item").forEach((item) =>
      item.classList.toggle("is-active", item.dataset.view === view),
    );
    $$("[data-view-panel]").forEach((panel) =>
      panel.classList.toggle("is-active", panel.dataset.viewPanel === view),
    );
    renderCurrentView();
  }

  function renderBatch() {
    const summary = $("#validationSummary");
    const groups = $("#validationGroups");
    if (!state.batch.rows.length) {
      summary.innerHTML = "";
      groups.innerHTML =
        '<div class="empty-detail"><strong>Nenhum arquivo validado</strong><span>Selecione um arquivo para iniciar a validação.</span></div>';
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

  function validateBatchRows(rows) {
    const valid = [];
    const warnings = [];
    const errors = [];

    rows.forEach((row, index) => {
      const record = normalizeBatchRecord(row);
      const messages = [];
      const alerts = [];
      const existing = findDemandForBatch(record);
      if (!record.ordem && !record.id) {
        messages.push("Informe Ordem SAP ou ID_Demanda_Controle.");
      }
      if ((record.ordem || record.id) && !existing) {
        messages.push(
          "Demanda nao encontrada na carteira JSON/Supabase atual. Linha nao sera gravada.",
        );
      }
      if (record.dataPlanejada && !toDate(record.dataPlanejada))
        messages.push("Data planejada inválida.");
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

  function findDemandForBatch(record) {
    if (record.id) {
      const byId = state.db.demandas.find((item) => item.id === record.id);
      if (byId) return byId;
    }
    if (record.ordem) {
      return state.db.demandas.find(
        (item) => String(item.ordem) === String(record.ordem),
      );
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
      const rows = await readBatchFile(file);
      state.batch.fileName = file.name;
      validateBatchRows(rows);
      renderBatch();
      showToast("Arquivo validado.", "success");
    } catch (error) {
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

    const records = candidates
      .map((item) => {
        const existing = findDemandForBatch(item.record);

        if (!existing) return null;

        const partial = buildBatchPartialUpdate(item.record);

        return prepareDemandForSave({
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
      })
      .filter(Boolean);

    await state.repo.bulkUpsertDemandas(records);
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
      status: state.batch.errors.length ? "PROCESSADO_COM_ERRO" : "PROCESSADO",
      detalheErro: state.batch.errors.length
        ? `${state.batch.errors.length} linhas com erro de validação.`
        : "",
    });
    const loteId = batchRun?.lote_id || batchRun?.id;

    if (loteId) {
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
    }
    await state.repo.addLog({
      usuario: state.currentUser.email,
      acao: "Carga em Lote",
      lista: "cargas_lote",
      referencia: `${records.length} registros`,
      detalhe: includeWarnings
        ? "Válidos e alertas confirmados"
        : "Somente válidos",
      modulo: "CARGA_LOTE",
    });
    await refreshAll();
    state.batch = {
      rows: [],
      valid: [],
      warnings: [],
      errors: [],
      fileName: "",
    };
    renderBatch();
    showToast(`${records.length} registros salvos.`, "success");
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
      `${futures.length} demandas | Página ${state.futurePage} de ${totalPages}`;

    const toolbarHtml = `
    <div class="future-toolbar" style="display:flex; gap:12px; align-items:end; flex-wrap:wrap; margin-bottom:14px;">
      <label style="min-width:280px; flex:1;">
        Buscar demanda futura
        <input
          id="futureSearch"
          type="search"
          placeholder="Digite ID, OM, descrição, centro, local, gerência ou supervisão"
          value="${escapeHtml(state.futureSearch)}"
        />
      </label>

      <label>
        Linhas por página
        <select id="futurePageSize">
          <option value="25" ${state.futurePageSize === 25 ? "selected" : ""}>25</option>
          <option value="50" ${state.futurePageSize === 50 ? "selected" : ""}>50</option>
          <option value="100" ${state.futurePageSize === 100 ? "selected" : ""}>100</option>
          <option value="200" ${state.futurePageSize === 200 ? "selected" : ""}>200</option>
        </select>
      </label>

      <button
        class="button secondary"
        id="futurePrevPage"
        type="button"
        ${state.futurePage <= 1 ? "disabled" : ""}
      >
        Anterior
      </button>

      <button
        class="button secondary"
        id="futureNextPage"
        type="button"
        ${state.futurePage >= totalPages ? "disabled" : ""}
      >
        Próxima
      </button>
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
    record.id = global.CCEData.stableDemandId(record);
    record.tipoOM = record.tipoDemanda;
    const centro = (state.db.centrosTrabalho || []).find(
      (item) =>
        normalizeCentroTrabalho(item.centroTrabalho) ===
        normalizeCentroTrabalho(record.centroTrabalho),
    );
    record.gerencia = centro?.gerencia || "";
    record.supervisao = centro?.supervisao || "";
    record.planejadorOM = centro?.planejadorOM || "";
    record.programador = centro?.programador || "";
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
      .slice(0, 8);
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

  function renderIndicators() {
    const demands = filteredDemandas();
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
    const futureNoOrder = demands.filter((item) => !item.ordem).length;
    const sapDemands = demands.filter((item) => item.ordem).length;
    const futureDemands = demands.filter(
      (item) =>
        !item.ordem ||
        normalizeText(item.origem).includes("DEMANDAS FUTURAS") ||
        normalizeText(item.tipoDemanda).includes("FUTURA"),
    ).length;
    const missingCenters = demands.filter(
      (item) => item.centroTrabalhoCadastrado === false,
    ).length;
    const cards = [
      ["Total de Demandas", stats.total, "recorte filtrado"],
      ["Demandas SAP", sapDemands, "com ordem"],
      ["Demandas Futuras", futureDemands, "JSON ou sistema"],
      ["Total a Planejar", stats.aPlanejar, "sem data"],
      ["Planejadas", stats.planejadas, "ativas"],
      ["Replanejadas", stats.replanejadas, "com histórico"],
      ["Realizadas", stats.realizadas, "baixadas"],
      ["Perdas", stats.perdas, "registradas"],
      ["Aderência", `${stats.aderencia}%`, "no prazo"],
      ["Ordens Vencidas", overdue, "sem realização"],
      ["Próximas do Vencimento", dueSoon.length, "20 dias"],
      ["Futuras sem Ordem", futureNoOrder, "aguardando vínculo"],
      ["Centros sem Cadastro", missingCenters, "alerta mestre"],
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
        .slice(0, 8)
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
      payload.permissaoPlanejar = form.permissaoPlanejar.checked;
      payload.permissaoReplanejar = form.permissaoReplanejar.checked;
      payload.permissaoRealizar = form.permissaoRealizar.checked;
      payload.permissaoConfigurar = form.permissaoConfigurar.checked;
      payload.permissaoExportar = form.permissaoExportar.checked;
      payload.permissaoCargaLote = form.permissaoCargaLote.checked;
      await state.repo.addUser(payload);
      await state.repo.addLog({
        usuario: state.currentUser.email,
        acao: "Cadastro Usuário",
        lista: "Usuarios_Central_Eletrovia",
        referencia: event.currentTarget.email.value,
        detalhe: "Usuário incluído pela administração.",
      });
      await refreshAll();
      showToast("Usuário cadastrado.", "success");
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
      await refreshAll();
      showToast("Configuração cadastrada.", "success");
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
      await refreshAll();
      showToast("Subgrupo cadastrado.", "success");
    });
  }

  function renderCentrosTrabalhoAdmin() {
    const centros = state.db.centrosTrabalho || [];
    const centrosNaoCadastrados = missingCentrosTrabalho();

    $("#adminContent").innerHTML = `
    <div class="admin-grid admin-grid-wide">
      <form class="admin-form" id="centroTrabalhoForm">
        <label>
          Centro de Trabalho
          <input name="centroTrabalho" required placeholder="Ex.: EVT-PCM-01" />
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

        <button class="button" type="submit">
          Salvar Centro de Trabalho
        </button>
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
                        data-edit-centro="${escapeHtml(item.centroTrabalhoChave || normalizeCentroTrabalho(item.centroTrabalho))}"
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

      record.ativo = record.ativo === "true";
      record.usuario = state.currentUser.email;

      await state.repo.upsertCentroTrabalho(record);

      await state.repo.addLog({
        usuario: state.currentUser.email,
        acao: "Cadastro Centro de Trabalho",
        lista: "cadastro_centros_trabalho",
        referencia: record.centroTrabalho,
        detalhe: `${record.gerencia || "-"} | ${record.supervisao || "-"}`,
        modulo: "CONFIGURACOES",
      });

      await refreshAll();
      showToast("Centro de trabalho salvo com sucesso.", "success");
    });

    $("#adminContent").addEventListener("click", (event) => {
      const button = event.target.closest("[data-edit-centro]");
      const newButton = event.target.closest("[data-new-centro]");
      if (!button && !newButton) return;

      if (newButton) {
        const form = $("#centroTrabalhoForm");
        form.reset();
        form.centroTrabalho.value = newButton.dataset.newCentro || "";
        form.ativo.value = "true";
        form.gerencia.focus();
        return;
      }

      const centro = centros.find(
        (item) =>
          (item.centroTrabalhoChave ||
            normalizeCentroTrabalho(item.centroTrabalho)) ===
          button.dataset.editCentro,
      );

      if (!centro) return;

      const form = $("#centroTrabalhoForm");

      form.centroTrabalho.value = centro.centroTrabalho || "";
      form.gerencia.value = centro.gerencia || "";
      form.supervisao.value = centro.supervisao || "";
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
      await refreshAll();
      showToast("Parâmetros salvos.", "success");
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
      "ordem",
      "descricao",
      "centro trabalho",
      "local instalacao",
      "competencia",
      "vencimento",
      "data planejada",
      "motivo",
      "justificativa",
      "perda",
      "motivo perda",
      "justificativa perda",
      "comentario",
    ];
    const example = [
      "910123456",
      "Inspeção preventiva em subestação",
      "EVT-ENE-04",
      "ELV-SE-0207",
      "2026-06",
      "2026-06-18",
      "2026-06-16",
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
      const button = event.target.closest("[data-alert-demand]");
      if (!button) return;
      state.selectedDemandId = button.dataset.alertDemand;
      $("#alertMenu").classList.add("hidden");
      switchView("carteira");
      renderCarteira();
    });

    $("#mainNav").addEventListener("click", (event) => {
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
      )
        return;
      state.page = 1;
      collectFilters();
      buildFilterOptions();
      renderCarteira();
    });
    $(".filter-panel").addEventListener("input", (event) => {
      if (event.target.matches("[data-multi-search]")) {
        const query = normalizeText(event.target.value);
        $$(".multi-option", event.target.closest(".multi-menu")).forEach(
          (option) => {
            option.classList.toggle(
              "hidden",
              query && !normalizeText(option.textContent).includes(query),
            );
          },
        );
        return;
      }
      if (event.target.id === "quickSearch") {
        state.page = 1;
        renderCarteira();
      }
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
      renderCarteira();
    });
    $("#toggleAdvancedFilters").addEventListener("click", () => {
      state.advancedFilters = !state.advancedFilters;
      $$(".advanced-filter").forEach((element) =>
        element.classList.toggle("hidden", !state.advancedFilters),
      );
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
    await loadDatabase();
    await autoSyncRealizadosFromSharePoint();
    state.pageSize = Number(state.db.parametros?.pageSizeDefault || 12);
    hydrateStaticUi();
    bindEvents();
    renderLoginState();
    renderCurrentView();
  }

  document.addEventListener("DOMContentLoaded", init);
})(window, document);

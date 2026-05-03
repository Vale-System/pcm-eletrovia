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
    Visualizador: { edit: false, admin: false, approve: false, export: true },
    Editor: { edit: true, admin: false, approve: false, export: true },
    Administrador: { edit: true, admin: true, approve: true, export: true },
    Gestor: { edit: false, admin: false, approve: true, export: true },
  };

  const state = {
    repo: null,
    db: null,
    currentUser: null,
    currentView: "carteira",
    adminTab: "usuarios",
    selectedDemandId: "",
    page: 1,
    pageSize: 12,
    advancedFilters: false,
    identity: null,
    realizedAutoSynced: false,
    filters: {},
    batch: {
      rows: [],
      valid: [],
      warnings: [],
      errors: [],
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
    const text = String(value).trim();
    if (!text) return null;
    const normalized = text.includes("/")
      ? text.split("/").reverse().join("-")
      : text;
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

  function getRules() {
    return (
      PROFILE_RULES[state.currentUser?.perfil] || PROFILE_RULES.Visualizador
    );
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
    return getRules().edit;
  }

  function canAdmin() {
    return getRules().admin;
  }

  function canExport() {
    return getRules().export;
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

  function statusClass(status) {
    if (status === "A Planejar") return "status-planejar";
    if (status === "Planejado") return "status-planejado";
    if (status === "Replanejado") return "status-replanejado";
    if (status === "Realizado" || status === "No Prazo")
      return "status-realizado";
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

  async function loadBaseFromJson() {
  const baseUrl = "./base/base_ordens.json";

  const response = await fetch(`${baseUrl}?v=${Date.now()}`, {
    method: "GET",
    cache: "no-store",
    headers: {
      Accept: "application/json"
    }
  });

  if (!response.ok) {
    throw new Error(`Erro ao carregar base_ordens.json do GitHub: ${response.status}`);
  }

  const data = await response.json();

  if (!Array.isArray(data)) {
    throw new Error("base_ordens.json precisa ser um array JSON.");
  }

  return data.map((item) => {
    const ordem = String(item.OrdemSAP || "").trim();

    return {
      id:
        item.ID_Demanda_Controle ||
        (ordem
          ? `DEM-SAP-${ordem}`
          : global.CCEData.stableDemandId({
              ordem: "",
              descricao: item.Descricao || "",
              centroTrabalho: item.CentroTrabalho || "",
              localInstalacao: item.LocalInstalacao || "",
              competencia: item.Competencia || "",
              vencimento: item.Vencimento || "",
              origem: "SAP BO"
            })),

      ordem,
      tipoOM: item.TipoOM || "",
      descricao: item.Descricao || "",
      gerencia: item.Gerencia || "",
      supervisao: item.Supervisao || "",
      centroTrabalho: item.CentroTrabalho || "",
      localInstalacao: item.LocalInstalacao || "",
      statusSistema: item.StatusSistema || "",
      statusUsuario: item.StatusUsuario || "",
      competencia: item.Competencia || "",
      dataRealizada: item.DataRealizada || "",
      vencimento: item.Vencimento || "",
      prioridade: item.Prioridade || "",
      toleranciaMin: item.ToleranciaMin || "",
      toleranciaMax: item.ToleranciaMax || "",

      dataPlanejada: "",
      dataReplanejadaAtual: "",
      perda: false,
      motivoPerda: "",
      justificativaPerda: "",
      comentario: "",
      usuarioResponsavel: "",
      dataUltimaAtualizacao: "",
      origem: item.Origem || "SAP BO",
      quantidadeReplanejamentos: 0,
      frequencia: "",
      observacao: "",
      vinculadaEm: ""
    };
  });
}
  async function loadDatabase() {
    const base = await loadBaseFromJson();

    state.db = {
      demandas: base,
      usuarios: [],
      configuracoes: {},
      parametros: {},
      historicoPlanejamento: [],
      historicoReplanejamento: [],
      historicoRealizadoPerdas: [],
      logs: [],
    };

    state.currentUser = {
      nome: "Weslley",
      email: "weslley.santos@vale.com",
      perfil: "Administrador",
    };

    state.selectedDemandId = base[0]?.id || "";
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
    $("#lastUpdateSide").textContent =
      todayText().split("-").reverse().join("/") + " agora";
    renderRole();
    buildFilterOptions();
    renderAlerts();
  }

  function renderRole() {
    $("#userName").textContent =
      state.currentUser?.nome || state.identity?.nome || "Usuário";
    $("#roleChip").textContent = state.currentUser?.perfil || "Visualizador";
    applyPermissions();
  }

  function applyPermissions() {
    const rules = getRules();
    $$(".admin-only").forEach((element) => {
      element.classList.toggle("hidden", !rules.admin);
      element.disabled = !rules.admin;
    });
    $$(".editor-only").forEach((element) => {
      element.disabled = !rules.edit;
    });
    $$('[data-view="administracao"]').forEach((element) => {
      element.disabled = !rules.admin;
      element.title = rules.admin ? "" : "Disponível para Administrador";
    });
    $("#exportCsv").disabled = !canExport();
  }

  function buildFilterOptions() {
    const demands = state.db.demandas;
    populateSelect(
      $("#filterGerencia"),
      uniqueOptions(demands.map((item) => item.gerencia)),
    );
    populateSelect(
      $("#filterCentro"),
      uniqueOptions(demands.map((item) => item.centroTrabalho)),
    );
    populateSelect(
      $("#filterTipo"),
      uniqueOptions(demands.map((item) => item.tipoOM)),
    );
    populateSelect(
      $("#filterCompetencia"),
      uniqueOptions(demands.map((item) => item.competencia)),
    );
    populateSelect(
      $("#filterLocal"),
      uniqueOptions(demands.map((item) => item.localInstalacao)),
    );
    populateSelect(
      $("#filterPrioridade"),
      uniqueOptions(demands.map((item) => item.prioridade)),
    );
    populateSelect(
      $("#filterStatusSistema"),
      uniqueOptions(demands.map((item) => item.statusSistema)),
    );
    populateSelect($("#filterStatusOperacional"), STATUS_OPTIONS);
    populateSelect(
      $("#filterAno"),
      uniqueOptions(demands.map((item) => item.vencimento?.slice(0, 4))),
    );
    populateSelect(
      $("#filterMes"),
      uniqueOptions(demands.map((item) => item.vencimento?.slice(5, 7))).map(
        (month) => `${month} - ${monthName(month)}`,
      ),
    );
  }

  function collectFilters() {
    const filters = {};
    $$("[data-filter]").forEach((field) => {
      filters[field.dataset.filter] = field.value;
    });
    filters.quickSearch = $("#quickSearch").value.trim();
    state.filters = filters;
  }

  function filteredDemandas() {
    collectFilters();
    const filters = state.filters;
    const search = normalizeText(filters.quickSearch);

    return state.db.demandas.filter((item) => {
      const fields = {
        gerencia: item.gerencia,
        centroTrabalho: item.centroTrabalho,
        tipoOM: item.tipoOM,
        competencia: item.competencia,
        localInstalacao: item.localInstalacao,
        prioridade: item.prioridade,
        statusSistema: item.statusSistema,
        anoVencimento: item.vencimento?.slice(0, 4),
        mesVencimento: item.vencimento
          ? `${item.vencimento.slice(5, 7)} - ${monthName(item.vencimento.slice(5, 7))}`
          : "",
      };

      for (const [key, value] of Object.entries(fields)) {
        if (filters[key] && String(value) !== filters[key]) return false;
      }
      if (
        filters.statusOperacional &&
        primaryStatusOf(item) !== filters.statusOperacional
      )
        return false;

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
    });
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
        const editDisabled = !canEdit();
        return `
          <tr class="${selected}" data-demand-id="${escapeHtml(item.id)}">
            <td><strong>${escapeHtml(item.id)}</strong></td>
            <td>${escapeHtml(compact(item.ordem))}</td>
            <td class="description-cell">${escapeHtml(item.descricao)}<div class="muted">${escapeHtml(item.origem || "")}</div></td>
            <td>${statusChip(status)}</td>
            <td>${statusChipGroup(substatuses)}</td>
            <td>${formatDate(item.vencimento)}</td>
            <td>${escapeHtml(item.competencia || "-")}</td>
            <td>${escapeHtml(item.tipoOM || "-")}</td>
            <td>${escapeHtml(item.centroTrabalho || "-")}</td>
            <td>${escapeHtml(item.localInstalacao || "-")}</td>
            <td>${escapeHtml(item.prioridade || "-")}</td>
            <td>${formatDate(item.dataPlanejada)}</td>
            <td>${formatDate(item.dataReplanejadaAtual)}</td>
            <td>${formatDate(item.dataRealizada)}</td>
            <td>${item.perda ? "Sim" : "Não"}</td>
            <td>${escapeHtml(compact(item.motivoPerda))}</td>
            <td>
              <div class="row-actions">
                ${actionButton("planejar", item.id, editDisabled || !allowed.planejar)}
                ${actionButton("replanejar", item.id, editDisabled || !allowed.replanejar)}
                ${actionButton("realizado", item.id, editDisabled || !allowed.realizado)}
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
    const editDisabled = !canEdit();
    panel.innerHTML = `
      <div class="detail-title">
        <span>${escapeHtml(demand.id)} ${demand.ordem ? `| Ordem ${escapeHtml(demand.ordem)}` : "| Sem ordem SAP"}</span>
        <h3>${escapeHtml(demand.descricao)}</h3>
      </div>
      <div class="detail-grid">
        <div class="detail-item"><span>Status</span><strong>${statusChip(status)}</strong></div>
        <div class="detail-item"><span>Substatus</span><strong>${statusChipGroup(substatuses)}</strong></div>
        <div class="detail-item"><span>Vencimento</span><strong>${formatDate(demand.vencimento)}</strong></div>
        <div class="detail-item"><span>Centro</span><strong>${escapeHtml(demand.centroTrabalho || "-")}</strong></div>
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
        ${actionButton("planejar", demand.id, editDisabled || !allowed.planejar, true)}
        ${actionButton("replanejar", demand.id, editDisabled || !allowed.replanejar, true)}
        ${actionButton("realizado", demand.id, editDisabled || !allowed.realizado, true)}
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
    if (action !== "historico" && !canEdit()) {
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
      await state.repo.upsertDemanda(demand);
      await state.repo.addHistory("planejamento", {
        demandaId: demand.id,
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
      await state.repo.upsertDemanda(demand);
      await state.repo.addHistory("replanejamento", {
        demandaId: demand.id,
        motivo: form.get("motivo"),
        justificativa: form.get("justificativa"),
        dataAnterior: previous,
        novaData: nextDate,
        usuario: userEmail,
        quantidadeReplanejamentos: demand.quantidadeReplanejamentos,
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
      await state.repo.upsertDemanda(demand);
      await state.repo.addHistory("realizadoPerda", {
        demandaId: demand.id,
        dataRealizada: demand.dataRealizada,
        perda: demand.perda,
        motivoPerda: demand.motivoPerda,
        justificativaPerda: demand.justificativaPerda,
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
    $("#lastUpdateSide").textContent =
      todayText().split("-").reverse().join("/") + ` ${time}`;
  }

  function switchView(view) {
    if (view === "administracao" && !canAdmin()) {
      showToast(
        "Administração disponível somente para Administrador.",
        "error",
      );
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
    if (normalized.perda) {
      normalized.perda = ["SIM", "S", "TRUE", "1"].includes(
        normalizeText(normalized.perda),
      );
    }
    normalized.origem = normalized.ordem
      ? "Carga em Lote"
      : "Demanda Antecipada";
    normalized.id = normalized.id || global.CCEData.stableDemandId(normalized);
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
      if (
        !record.ordem &&
        !(
          record.descricao &&
          record.centroTrabalho &&
          record.localInstalacao &&
          record.competencia &&
          record.vencimento
        )
      ) {
        messages.push(
          "Sem ordem SAP e sem campos mínimos para criar ID interno de demanda futura.",
        );
      }
      if (record.dataPlanejada && !toDate(record.dataPlanejada))
        messages.push("Data planejada inválida.");
      if (record.dataRealizada && !toDate(record.dataRealizada))
        messages.push("Data realizada inválida.");
      if (record.perda && (!record.motivoPerda || !record.justificativaPerda)) {
        messages.push("Perda exige motivo e justificativa.");
      }
      if (
        record.ordem &&
        !state.db.demandas.some((item) => item.ordem === String(record.ordem))
      ) {
        alerts.push(
          "Ordem não encontrada na carteira atual; será criada na camada de controle.",
        );
      }
      if (!record.ordem)
        alerts.push(
          "Demanda sem ordem SAP será salva com ID_Demanda_Controle.",
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

    state.batch = { rows, valid, warnings, errors };
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
      validateBatchRows(rows);
      renderBatch();
      showToast("Arquivo validado.", "success");
    } catch (error) {
      showToast(error.message, "error");
    }
  }

  async function saveBatch(includeWarnings) {
    if (!canEdit()) {
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

    const records = candidates.map((item) => {
      const existing = item.record.ordem
        ? state.db.demandas.find(
            (demand) => demand.ordem === String(item.record.ordem),
          )
        : null;
      return {
        ...(existing || {}),
        ...item.record,
        ordem: item.record.ordem ? String(item.record.ordem) : "",
        id: existing?.id || item.record.id,
        usuarioResponsavel: state.currentUser.email,
        dataUltimaAtualizacao: new Date().toISOString(),
      };
    });

    await state.repo.bulkUpsertDemandas(records);
    await state.repo.addLog({
      usuario: state.currentUser.email,
      acao: "Carga em Lote",
      lista: "Controle_Demandas_Eletrovia",
      referencia: `${records.length} registros`,
      detalhe: includeWarnings
        ? "Válidos e alertas confirmados"
        : "Somente válidos",
    });
    await refreshAll();
    state.batch = { rows: [], valid: [], warnings: [], errors: [] };
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

  function renderFutureDemandas() {
    const futures = state.db.demandas.filter(
      (item) =>
        !item.ordem &&
        ["Demanda Antecipada", "Futura"].includes(
          item.origem || item.tipoDemanda,
        ),
    );
    $("#futureCount").textContent = `${futures.length} demandas`;
    $("#futureDemandList").innerHTML =
      futures
        .map((item) => {
          const suggestions = linkSuggestions(item).slice(0, 3);
          return `
            <article class="future-card">
              <header>
                <div>
                  <h3>${escapeHtml(item.descricao)}</h3>
                  <span class="muted">${escapeHtml(item.id)} | ${escapeHtml(item.centroTrabalho)} | ${escapeHtml(item.localInstalacao)} | ${escapeHtml(item.competencia)}</span>
                </div>
                ${statusChipGroup(statusListOf(item))}
              </header>
              <div class="detail-grid" style="margin-top: 10px;">
                <div class="detail-item"><span>Vencimento</span><strong>${formatDate(item.vencimento)}</strong></div>
                <div class="detail-item"><span>Frequência</span><strong>${escapeHtml(item.frequencia || "-")}</strong></div>
              </div>
              <div class="suggestion-list">
                ${
                  suggestions.length
                    ? suggestions
                        .map(
                          (suggestion) => `
                        <div class="suggestion-item">
                          <div>
                            <strong>${escapeHtml(suggestion.target.ordem)} | ${escapeHtml(suggestion.target.descricao)}</strong>
                            <div class="muted">${suggestion.score}% de similaridade</div>
                          </div>
                          <button class="button editor-only" data-link-future="${escapeHtml(item.id)}" data-link-target="${escapeHtml(suggestion.target.id)}" type="button">Vincular</button>
                        </div>
                      `,
                        )
                        .join("")
                    : '<span class="muted">Sem sugestão automática no recorte atual.</span>'
                }
              </div>
            </article>
          `;
        })
        .join("") ||
      '<div class="empty-detail"><strong>Nenhuma demanda futura pendente</strong><span>Todas as demandas futuras estão vinculadas ou realizadas.</span></div>';
    applyPermissions();
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
    if (!canEdit()) return;
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
    await state.repo.upsertDemanda(future);
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
    if (!canEdit()) {
      showToast("Perfil sem permissão para criar demanda futura.", "error");
      return;
    }
    const form = new FormData(event.currentTarget);
    const record = Object.fromEntries(form.entries());
    record.ordem = "";
    record.id = global.CCEData.stableDemandId(record);
    record.tipoOM = record.tipoDemanda;
    record.gerencia = "GME Corredor";
    record.prioridade = "Média";
    record.statusSistema = "PREV";
    record.toleranciaMin = record.vencimento;
    record.toleranciaMax = record.vencimento;
    record.dataPlanejada = "";
    record.dataReplanejadaAtual = "";
    record.dataRealizada = "";
    record.perda = false;
    record.origem = "Demanda Antecipada";
    record.usuarioResponsavel = state.currentUser.email;
    await state.repo.upsertDemanda(record);
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
    const demands = state.db.demandas;
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
    const cards = [
      ["Total de Ordens", stats.total, "carteira tratada"],
      ["Total a Planejar", stats.aPlanejar, "sem data"],
      ["Planejadas", stats.planejadas, "ativas"],
      ["Replanejadas", stats.replanejadas, "com histórico"],
      ["Realizadas", stats.realizadas, "baixadas"],
      ["Perdas", stats.perdas, "registradas"],
      ["Aderência", `${stats.aderencia}%`, "no prazo"],
      ["Ordens Vencidas", overdue, "sem realização"],
      ["Próximas do Vencimento", dueSoon.length, "20 dias"],
      ["Futuras sem Ordem", futureNoOrder, "aguardando vínculo"],
    ];
    $("#indicatorGrid").innerHTML = cards
      .map(
        ([label, value, note]) =>
          `<article class="indicator-card"><span>${label}</span><strong>${value}</strong><small>${note}</small></article>`,
      )
      .join("");
    renderBars(
      $("#lossByCenter"),
      countBy(
        demands.filter((item) => item.perda),
        (item) => item.centroTrabalho,
      ),
    );
    renderBars(
      $("#lossReasons"),
      countBy(
        demands.filter((item) => item.perda),
        (item) => item.motivoPerda,
      ),
    );
    renderBars(
      $("#replanRanking"),
      countBy(
        demands.filter(
          (item) => Number(item.quantidadeReplanejamentos || 0) > 0,
        ),
        (item) => item.centroTrabalho,
      ),
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
    else if (state.adminTab === "parametros") renderParameterAdmin();
    else renderConfigAdmin(state.adminTab);
  }

  function renderUserAdmin() {
    $("#adminContent").innerHTML = `
      <div class="admin-grid">
        <form class="admin-form" id="addUserForm">
          <label>Nome<input name="nome" required /></label>
          <label>E-mail<input name="email" type="email" required /></label>
          <label>Matrícula<input name="matricula" /></label>
          <label>Área<input name="area" /></label>
          <label>Perfil<select name="perfil">${optionsMarkup(Object.keys(PROFILE_RULES), "Visualizador")}</select></label>
          <button class="button" type="submit">Adicionar Usuário</button>
        </form>
        <div class="admin-list">
          ${state.db.usuarios
            .map(
              (user) => `
              <div class="admin-list-item">
                <div><strong>${escapeHtml(user.nome)}</strong><div class="muted">${escapeHtml(user.email)} | ${escapeHtml(user.perfil)} | ${user.ativo ? "Ativo" : "Inativo"}</div></div>
              </div>
            `,
            )
            .join("")}
        </div>
      </div>
    `;
    $("#addUserForm").addEventListener("submit", async (event) => {
      event.preventDefault();
      await state.repo.addUser(
        Object.fromEntries(new FormData(event.currentTarget).entries()),
      );
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
      <div class="admin-grid">
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
        <div class="admin-list">
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

  function renderParameterAdmin() {
    const params = state.db.parametros || {};
    $("#adminContent").innerHTML = `
      <form class="admin-form" id="parameterForm" style="max-width: 680px;">
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
      "Vencimento",
      "Competência",
      "Tipo OM",
      "Centro Trabalho",
      "Local Instalação",
      "Prioridade",
      "Status Operacional",
      "Substatus",
      "Data Planejada",
      "Data Replanejada",
      "Data Realizada",
      "Perda",
      "Motivo Perda",
    ];
    const body = rows.map((item) => [
      item.id,
      item.ordem,
      item.descricao,
      item.vencimento,
      item.competencia,
      item.tipoOM,
      item.centroTrabalho,
      item.localInstalacao,
      item.prioridade,
      primaryStatusOf(item),
      substatusListOf(item).join(" | "),
      item.dataPlanejada,
      item.dataReplanejadaAtual,
      item.dataRealizada,
      item.perda ? "Sim" : "Não",
      item.motivoPerda,
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

    $$("[data-filter]").forEach((field) => {
      field.addEventListener("change", () => {
        state.page = 1;
        renderCarteira();
      });
    });
    $("#quickSearch").addEventListener("input", () => {
      state.page = 1;
      renderCarteira();
    });
    $("#clearFilters").addEventListener("click", () => {
      $$("[data-filter]").forEach((field) => {
        field.value = "";
      });
      $("#quickSearch").value = "";
      state.page = 1;
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
      showToast("Dados de demonstração recarregados.", "success");
    });
    $("#validateBatch").addEventListener("click", validateBatchFile);
    $("#clearBatch").addEventListener("click", () => {
      state.batch = { rows: [], valid: [], warnings: [], errors: [] };
      $("#batchFile").value = "";
      renderBatch();
    });
    $("#saveValidBatch").addEventListener("click", () => saveBatch(false));
    $("#saveConfirmedBatch").addEventListener("click", () => saveBatch(true));
    $("#downloadTemplate").addEventListener("click", downloadTemplate);
    $("#futureDemandForm").addEventListener("submit", createFutureDemand);
    $("#futureDemandList").addEventListener("click", (event) => {
      const button = event.target.closest("[data-link-future]");
      if (button)
        linkFutureDemand(button.dataset.linkFuture, button.dataset.linkTarget);
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
    renderCurrentView();
  }

  document.addEventListener("DOMContentLoaded", init);
})(window, document);

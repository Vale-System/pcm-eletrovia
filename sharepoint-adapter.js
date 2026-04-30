(function defineRepositories(global) {
  const STORAGE_KEY = "cce.v2.database";

  const LIST_NAMES = {
    baseOrdensSap: "Base_Ordens_SAP",
    controleDemandas: "Controle_Demandas_Eletrovia",
    historicoPlanejamento: "Historico_Planejamento",
    historicoReplanejamento: "Historico_Replanejamento",
    historicoRealizadoPerdas: "Historico_Realizado_Perdas",
    usuarios: "Usuarios_Central_Eletrovia",
    motivos: "Configuracoes_Motivos",
    justificativas: "Configuracoes_Justificativas",
    perfisPerda: "Configuracoes_Perfil_Perdas",
    justificativasPerda: "Configuracoes_Justificativas_Perdas",
    logs: "Logs_Sistema",
  };

  function clone(value) {
    return JSON.parse(JSON.stringify(value));
  }

  function uid(prefix) {
    if (global.crypto && global.crypto.randomUUID) {
      return `${prefix}-${global.crypto.randomUUID().slice(0, 8).toUpperCase()}`;
    }
    return `${prefix}-${Date.now().toString(36).toUpperCase()}${Math.random().toString(36).slice(2, 6).toUpperCase()}`;
  }

  function normalizeText(value) {
    return String(value || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toUpperCase()
      .trim();
  }

  function stableDemandId(record) {
    const raw = [
      record.ordem,
      record.centroTrabalho,
      record.localInstalacao,
      record.descricao,
      record.tipoOM || record.tipoDemanda,
      record.competencia,
      record.vencimento,
      record.origem,
    ]
      .filter(Boolean)
      .map(normalizeText)
      .join("|");
    let hash = 0;
    for (let index = 0; index < raw.length; index += 1) {
      hash = (hash << 5) - hash + raw.charCodeAt(index);
      hash |= 0;
    }
    const suffix = Math.abs(hash)
      .toString(36)
      .toUpperCase()
      .padStart(6, "0")
      .slice(0, 6);
    return record.ordem
      ? `DEM-SAP-${record.ordem}`
      : `DF-${record.competencia || "SEMCOMP"}-${suffix}`;
  }

  function slugify(value) {
    return normalizeText(value)
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-|-$/g, "");
  }

  function mapSharePointUser(item) {
    return {
      id: item.ID || item.Id || item.id,
      nome: item.Nome || item.Title || item.nome || "",
      email: item.Email || item.EMail || item.email || "",
      matricula: item.Matricula || item.Matrícula || item.matricula || "",
      area: item.Area || item.Área || item.area || "",
      perfil: item.PerfilAcesso || item.Perfil || item.perfil || "Visualizador",
      ativo: item.Ativo !== false && item.Ativo !== "Não",
    };
  }

  function createSeedDatabase() {
    const sample = global.CCE_SAMPLE || {};
    return {
      demandas: clone(sample.demands || []),
      usuarios: clone(sample.users || []),
      configuracoes: clone(sample.configs || {}),
      parametros: clone(sample.parameters || {}),
      historicoPlanejamento: clone(sample.historicoPlanejamento || []),
      historicoReplanejamento: clone(sample.historicoReplanejamento || []),
      historicoRealizadoPerdas: clone(sample.historicoRealizadoPerdas || []),
      logs: clone(sample.logs || []),
    };
  }

  class LocalRepository {
    constructor() {
      this.mode = "Local";
      this.database = this.load();
    }

    load() {
      const current = global.localStorage.getItem(STORAGE_KEY);
      if (!current) {
        const seed = createSeedDatabase();
        global.localStorage.setItem(STORAGE_KEY, JSON.stringify(seed));
        return seed;
      }
      try {
        return JSON.parse(current);
      } catch (error) {
        const seed = createSeedDatabase();
        global.localStorage.setItem(STORAGE_KEY, JSON.stringify(seed));
        return seed;
      }
    }

    persist() {
      global.localStorage.setItem(STORAGE_KEY, JSON.stringify(this.database));
    }

    async reset() {
      this.database = createSeedDatabase();
      this.persist();
      return clone(this.database);
    }

    async getAll() {
      return clone(this.database);
    }

    async getDemandas() {
      return clone(this.database.demandas);
    }

    async upsertDemanda(record) {
      const next = clone(record);
      next.id = next.id || stableDemandId(next);
      next.dataUltimaAtualizacao = new Date().toISOString();
      const index = this.database.demandas.findIndex(
        (item) =>
          item.id === next.id || (next.ordem && item.ordem === next.ordem),
      );
      if (index >= 0) {
        this.database.demandas[index] = {
          ...this.database.demandas[index],
          ...next,
        };
      } else {
        this.database.demandas.unshift(next);
      }
      this.persist();
      return clone(next);
    }

    async bulkUpsertDemandas(records) {
      const saved = [];
      for (const record of records) {
        saved.push(await this.upsertDemanda(record));
      }
      return saved;
    }

    async addHistory(type, entry) {
      const map = {
        planejamento: "historicoPlanejamento",
        replanejamento: "historicoReplanejamento",
        realizadoPerda: "historicoRealizadoPerdas",
      };
      const key = map[type];
      if (!key) return null;
      const next = {
        id: uid(
          type === "realizadoPerda"
            ? "HPR"
            : type === "replanejamento"
              ? "HR"
              : "HP",
        ),
        dataHora: new Date().toISOString(),
        ...entry,
      };
      this.database[key].unshift(next);
      this.persist();
      return clone(next);
    }

    async addLog(entry) {
      const next = {
        id: uid("LOG"),
        dataHora: new Date().toISOString(),
        ...entry,
      };
      this.database.logs.unshift(next);
      this.persist();
      return clone(next);
    }

    async addConfigItem(group, value, parentId = "") {
      const clean = String(value || "").trim();
      if (!clean) return null;
      this.database.configuracoes[group] =
        this.database.configuracoes[group] || [];
      const items = this.database.configuracoes[group];
      const exists = items.some((item) => {
        const name = typeof item === "string" ? item : item.nome;
        const parent = item.motivoId || item.perfilId || "";
        return (
          normalizeText(name) === normalizeText(clean) &&
          (!parentId || parent === parentId)
        );
      });
      if (!exists) {
        const item = { id: slugify(clean), nome: clean, ativo: true };
        if (group === "justificativas") item.motivoId = parentId;
        if (group === "justificativasPerda") item.perfilId = parentId;
        this.database.configuracoes[group].push(item);
        this.database.configuracoes[group].sort((a, b) =>
          String(a.nome || a).localeCompare(String(b.nome || b), "pt-BR"),
        );
        this.persist();
      }
      return clean;
    }

    async addUser(user) {
      const next = {
        id: uid("USR"),
        ativo: true,
        ...user,
      };
      this.database.usuarios.push(next);
      this.persist();
      return clone(next);
    }

    async updateParameters(parameters) {
      this.database.parametros = { ...this.database.parametros, ...parameters };
      this.persist();
      return clone(this.database.parametros);
    }

    async getCurrentUserIdentity() {
      const params = new URLSearchParams(global.location.search);
      const email =
        params.get("user") ||
        (params.get("debugUsers") === "1"
          ? global.localStorage.getItem("cce.currentUser")
          : "") ||
        "";
      const user =
        this.database.usuarios.find(
          (item) => normalizeText(item.email) === normalizeText(email),
        ) || this.database.usuarios.find((item) => item.ativo);
      return {
        email: user?.email || email,
        nome: user?.nome || email || "Usuário local",
      };
    }
  }

  class SharePointRestRepository {
    constructor(options = {}) {
      this.mode = "SharePoint REST";
      this.webUrl =
        options.webUrl || global._spPageContextInfo?.webAbsoluteUrl || "";
      this.digest = "";
    }

    async ensureDigest() {
      if (this.digest) return this.digest;
      const response = await fetch(`${this.webUrl}/_api/contextinfo`, {
        method: "POST",
        headers: {
          Accept: "application/json;odata=nometadata",
        },
      });
      const payload = await response.json();
      this.digest = payload.FormDigestValue;
      return this.digest;
    }

    async spFetch(path, options = {}) {
      const method = options.method || "GET";
      const headers = {
        Accept: "application/json;odata=nometadata",
        "Content-Type": "application/json;odata=nometadata",
        ...(options.headers || {}),
      };
      if (method !== "GET") {
        headers["X-RequestDigest"] = await this.ensureDigest();
      }
      const response = await fetch(`${this.webUrl}/_api/${path}`, {
        ...options,
        method,
        headers,
      });
      if (!response.ok) {
        throw new Error(
          `SharePoint ${response.status}: ${await response.text()}`,
        );
      }
      if (response.status === 204) return null;
      return response.json();
    }

    async listItems(listName, query = "") {
      const payload = await this.spFetch(
        `web/lists/getbytitle('${listName}')/items${query}`,
      );
      return payload.value || [];
    }

    async createItem(listName, item) {
      return this.spFetch(`web/lists/getbytitle('${listName}')/items`, {
        method: "POST",
        body: JSON.stringify(item),
      });
    }

    async updateItem(listName, itemId, item) {
      return this.spFetch(
        `web/lists/getbytitle('${listName}')/items(${itemId})`,
        {
          method: "POST",
          headers: {
            "IF-MATCH": "*",
            "X-HTTP-Method": "MERGE",
          },
          body: JSON.stringify(item),
        },
      );
    }

    async getAll() {
      const [
        demandas,
        usuarios,
        logs,
        motivos,
        justificativas,
        perfisPerda,
        justificativasPerda,
      ] = await Promise.all([
        this.listItems(LIST_NAMES.controleDemandas, "?$top=5000"),
        this.listItems(LIST_NAMES.usuarios, "?$top=5000"),
        this.listItems(LIST_NAMES.logs, "?$top=5000&$orderby=Created desc"),
        this.listItems(LIST_NAMES.motivos, "?$top=5000"),
        this.listItems(LIST_NAMES.justificativas, "?$top=5000"),
        this.listItems(LIST_NAMES.perfisPerda, "?$top=5000"),
        this.listItems(LIST_NAMES.justificativasPerda, "?$top=5000"),
      ]);
      return {
        demandas,
        usuarios: usuarios.map(mapSharePointUser),
        logs,
        configuracoes: {
          motivos: motivos.map((item) => ({
            id: item.Chave || slugify(item.Title),
            nome: item.Title,
            ativo: item.Ativo !== false,
          })),
          justificativas: justificativas.map((item) => ({
            id: item.Chave || slugify(item.Title),
            motivoId:
              item.MotivoChave || item.MotivoId || slugify(item.Motivo || ""),
            nome: item.Title,
            ativo: item.Ativo !== false,
          })),
          perfisPerda: perfisPerda.map((item) => ({
            id: item.Chave || slugify(item.Title),
            nome: item.Title,
            ativo: item.Ativo !== false,
          })),
          justificativasPerda: justificativasPerda.map((item) => ({
            id: item.Chave || slugify(item.Title),
            perfilId:
              item.PerfilChave ||
              item.PerfilPerdaId ||
              slugify(item.PerfilPerda || ""),
            nome: item.Title,
            ativo: item.Ativo !== false,
          })),
        },
        parametros: {
          sharePointLibrary: "Documentos Compartilhados/SAP_BO",
          sapExcelFileName: "base_ordens_sap.xlsx",
          realizedExcelFileName: "base_realizados_sap.xlsx",
        },
        historicoPlanejamento: [],
        historicoReplanejamento: [],
        historicoRealizadoPerdas: [],
      };
    }

    async upsertDemanda(record) {
      const next = {
        ...record,
        ID_Demanda_Controle: record.id || stableDemandId(record),
      };
      await this.createItem(LIST_NAMES.controleDemandas, next);
      return next;
    }

    async bulkUpsertDemandas(records) {
      const saved = [];
      for (const record of records) {
        saved.push(await this.upsertDemanda(record));
      }
      return saved;
    }

    async addHistory(type, entry) {
      const list = {
        planejamento: LIST_NAMES.historicoPlanejamento,
        replanejamento: LIST_NAMES.historicoReplanejamento,
        realizadoPerda: LIST_NAMES.historicoRealizadoPerdas,
      }[type];
      return list ? this.createItem(list, entry) : null;
    }

    async addLog(entry) {
      return this.createItem(LIST_NAMES.logs, entry);
    }

    async getRealizadosFileBuffer(parameters = {}) {
      const libraryPath =
        parameters.sharePointLibrary || "Documentos Compartilhados/SAP_BO";
      const fileName =
        parameters.realizedExcelFileName || "base_realizados_sap.xlsx";
      const serverRelativeBase =
        global._spPageContextInfo?.webServerRelativeUrl || "";
      const cleanBase = serverRelativeBase.replace(/\/$/, "");
      const cleanLibrary = String(libraryPath).replace(/^\/|\/$/g, "");
      const serverRelativeUrl =
        `${cleanBase}/${cleanLibrary}/${fileName}`.replace(/\/{2,}/g, "/");
      const response = await fetch(
        `${this.webUrl}/_api/web/GetFileByServerRelativeUrl('${serverRelativeUrl.replace(/'/g, "''")}')/$value`,
        {
          headers: { Accept: "application/octet-stream" },
        },
      );
      if (!response.ok) {
        throw new Error(
          `SharePoint ${response.status}: não foi possível ler ${fileName}`,
        );
      }
      return response.arrayBuffer();
    }

    async addConfigItem(group, value, parentId = "") {
      const title = String(value || "").trim();
      if (!title) return null;
      const listName = {
        motivos: LIST_NAMES.motivos,
        justificativas: LIST_NAMES.justificativas,
        perfisPerda: LIST_NAMES.perfisPerda,
        justificativasPerda: LIST_NAMES.justificativasPerda,
      }[group];
      const item = {
        Title: title,
        Chave: slugify(title),
        Ativo: true,
      };
      if (group === "justificativas") item.MotivoChave = parentId;
      if (group === "justificativasPerda") item.PerfilChave = parentId;
      return this.createItem(listName, item);
    }

    async addUser(user) {
      return this.createItem(LIST_NAMES.usuarios, {
        Title: user.nome,
        Email: user.email,
        Matricula: user.matricula,
        Area: user.area,
        PerfilAcesso: user.perfil,
        Ativo: true,
      });
    }

    async updateParameters(parameters) {
      await this.addLog({
        Acao: "Parâmetros",
        Detalhe: JSON.stringify(parameters),
      });
      return parameters;
    }

    async getCurrentUserIdentity() {
      if (global._spPageContextInfo?.userEmail) {
        return {
          email: global._spPageContextInfo.userEmail,
          nome:
            global._spPageContextInfo.userDisplayName ||
            global._spPageContextInfo.userLoginName ||
            global._spPageContextInfo.userEmail,
        };
      }
      const payload = await this.spFetch(
        "web/currentuser?$select=Title,Email,LoginName",
      );
      return {
        email: payload.Email || "",
        nome: payload.Title || payload.LoginName || payload.Email || "",
      };
    }
  }

  function createRepository() {
    const params = new URLSearchParams(global.location.search);
    const wantsSharePoint = params.get("data") === "sharepoint";
    if (wantsSharePoint && global._spPageContextInfo?.webAbsoluteUrl) {
      return new SharePointRestRepository();
    }
    return new LocalRepository();
  }
  const SHAREPOINT_SITE_URL =
    "https://globalvale.sharepoint.com/sites/ControlePCMEletrovia";

  const SHAREPOINT_LISTS = {
    demandas: "Controle_Demandas_Eletrovia",
    usuarios: "Usuarios_Central_Eletrovia",
    motivos: "Configuracoes_Motivos",
    justificativas: "Configuracoes_Justificativas",
    perfilPerdas: "Configuracoes_Perfil_Perdas",
    justificativasPerdas: "Configuracoes_Justificativas_Perdas",
    historicoPlanejamento: "Historico_Planejamento",
    historicoReplanejamento: "Historico_Replanejamento",
    historicoRealizadoPerdas: "Historico_Realizado_Perdas",
    logs: "Logs_Sistema",
  };

  async function spRequestDigest() {
    const response = await fetch(`${SHAREPOINT_SITE_URL}/_api/contextinfo`, {
      method: "POST",
      credentials: "include",
      headers: {
        Accept: "application/json;odata=verbose",
      },
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error("Erro ao obter Request Digest: " + errorText);
    }

    const data = await response.json();
    return data.d.GetContextWebInformation.FormDigestValue;
  }

  async function spGetItems(listName, select = "*", filter = "", top = 100) {
    let url = `${SHAREPOINT_SITE_URL}/_api/web/lists/getbytitle('${listName}')/items?$select=${select}&$top=${top}`;

    if (filter) {
      url += `&$filter=${encodeURIComponent(filter)}`;
    }

    const response = await fetch(url, {
      method: "GET",
      credentials: "include",
      headers: {
        Accept: "application/json;odata=verbose",
      },
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error("Erro ao buscar itens SharePoint: " + errorText);
    }

    const data = await response.json();
    return data.d.results;
  }

  async function spCreateItem(listName, itemType, fields) {
    const digest = await spRequestDigest();

    const response = await fetch(
      `${SHAREPOINT_SITE_URL}/_api/web/lists/getbytitle('${listName}')/items`,
      {
        method: "POST",
        credentials: "include",
        headers: {
          Accept: "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
          "X-RequestDigest": digest,
        },
        body: JSON.stringify({
          __metadata: {
            type: itemType,
          },
          ...fields,
        }),
      },
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error("Erro ao criar item SharePoint: " + errorText);
    }

    const data = await response.json();
    return data.d;
  }

  async function spUpdateItem(listName, itemType, itemId, fields) {
    const digest = await spRequestDigest();

    const response = await fetch(
      `${SHAREPOINT_SITE_URL}/_api/web/lists/getbytitle('${listName}')/items(${itemId})`,
      {
        method: "POST",
        credentials: "include",
        headers: {
          Accept: "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
          "X-RequestDigest": digest,
          "IF-MATCH": "*",
          "X-HTTP-Method": "MERGE",
        },
        body: JSON.stringify({
          __metadata: {
            type: itemType,
          },
          ...fields,
        }),
      },
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error("Erro ao atualizar item SharePoint: " + errorText);
    }

    return true;
  }

  async function criarDemandaFuturaSharePoint(demanda) {
    const digest = await getRequestDigest();

    const response = await fetch(
      `${SHAREPOINT_SITE_URL}/_api/web/lists/getbytitle('Controle_Demandas_Eletrovia')/items`,
      {
        method: "POST",
        credentials: "include",
        headers: {
          Accept: "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
          "X-RequestDigest": digest,
        },
        body: JSON.stringify({
          __metadata: {
            type: "SP.Data.Controle_x005f_Demandas_x005f_EletroviaListItem",
          },

          Title: demanda.ID_Demanda_Controle,
          OrdemSAP: demanda.OrdemSAP || "",
          TipoDemanda: "Antecipada",
          Descricao: demanda.Descricao,
          CentroTrabalho: demanda.CentroTrabalho,
          LocalInstalacao: demanda.LocalInstalacao,
          Competencia: demanda.Competencia,
          TipoOM: demanda.TipoOM || "Sistematica",
          Prioridade: demanda.Prioridade || "Media",
          StatusOperacional: "A Planejar",
          SubstatusOperacional: "Pendente",
          Perda: false,
          OrigemInformacao: "Demanda Antecipada",
          QuantidadeReplanejamentos: 0,
          DataUltimaAtualizacao: new Date().toISOString(),
        }),
      },
    );

    if (!response.ok) {
      const error = await response.text();
      console.error("Erro completo:", error);
      throw new Error(error);
    }

    return await response.json();
  }

  async function getRequestDigest() {
    const response = await fetch(`${SHAREPOINT_SITE_URL}/_api/contextinfo`, {
      method: "POST",
      credentials: "include",
      headers: {
        Accept: "application/json;odata=verbose",
      },
    });

    const data = await response.json();
    return data.d.GetContextWebInformation.FormDigestValue;
  }

  global.CCEData = {
    LIST_NAMES,
    LocalRepository,
    SharePointRestRepository,
    createRepository,
    stableDemandId,
    normalizeText,
    slugify,
  };
})(window);

(function defineSupabaseRepository(global) {
  const SUPABASE_URL = "https://vgtqolxlnljdrgtrsubz.supabase.co";

  // Cole aqui a sua anon public key do Supabase.
  // Nunca use service_role aqui.
  const SUPABASE_ANON_KEY =
    "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InZndHFvbHhsbmxqZHJndHJzdWJ6Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzc1ODYwODUsImV4cCI6MjA5MzE2MjA4NX0.BR6xVYREcQxunHp2E7sCQ3F9llcKTkMlvqJcCCXFBcQ";

  const PAGE_SIZE = 1000;

  const TABLES = {
    controle: "controle_demandas_eletrovia",
    usuarios: "usuarios_central_eletrovia",
    motivos: "configuracoes_motivos",
    justificativas: "configuracoes_justificativas",
    perfisPerda: "configuracoes_perfil_perdas",
    justificativasPerda: "configuracoes_justificativas_perdas",
    historicoPlanejamento: "historico_planejamento",
    historicoReplanejamento: "historico_replanejamento",
    historicoRealizadoPerdas: "historico_realizado_perdas",
    logs: "logs_sistema",
    cargasLote: "cargas_lote",
    cargasLoteItens: "cargas_lote_itens",
    evidencias: "evidencias",
    centrosTrabalho: "cadastro_centros_trabalho",
    // parametros: "parametros_sistema",
  };

  function normalizeText(value) {
    return String(value || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toUpperCase()
      .trim();
  }

  function slugify(value) {
    return normalizeText(value)
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-|-$/g, "");
  }

  function cleanDate(value) {
    if (!value) return null;
    return String(value).slice(0, 10);
  }

  function cleanDateTime(value) {
    if (!value) return null;
    return new Date(value).toISOString();
  }

  function excelSerialToDate(value) {
    const serial = Number(value);
    if (!Number.isFinite(serial) || serial < 20000 || serial > 80000)
      return null;
    return new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
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
    if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text.slice(0, 7);
    if (/^\d{6}$/.test(text) && text.startsWith("20")) {
      return `${text.slice(0, 4)}-${text.slice(4, 6)}`;
    }
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(text)) {
      const [, month, year] = text.split("/");
      return `${year}-${month}`;
    }
    if (/^\d{5}$/.test(text)) {
      const date = excelSerialToDate(Number(text));
      return date ? date.toISOString().slice(0, 7) : "";
    }
    return text;
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

  function stableDemandId(record) {
    const ordem = String(record.ordem || record.OrdemSAP || "").trim();

    if (ordem) {
      return `DEM-SAP-${ordem}`;
    }

    const raw = [
      record.centroTrabalho || record.centro_trabalho,
      record.localInstalacao || record.local_instalacao,
      record.descricao,
      record.tipoOM ||
        record.tipo_om ||
        record.tipoDemanda ||
        record.tipo_demanda,
      record.competencia,
      record.vencimento,
      record.origem || record.origem_informacao,
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

    return `DF-${record.competencia || "SEMCOMP"}-${suffix}`;
  }

  async function request(path, options = {}) {
    const response = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, {
      ...options,
      headers: {
        apikey: SUPABASE_ANON_KEY,
        Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
        Accept: "application/json",
        ...(options.headers || {}),
      },
    });

    if (!response.ok) {
      const text = await response.text();
      throw new Error(`Supabase ${response.status}: ${text}`);
    }

    if (response.status === 204) return null;

    const text = await response.text();
    return text ? JSON.parse(text) : null;
  }

  async function selectAll(table, query = "select=*") {
    const rows = [];
    let from = 0;

    while (true) {
      const to = from + PAGE_SIZE - 1;

      const batch = await request(`${table}?${query}`, {
        method: "GET",
        headers: {
          Range: `${from}-${to}`,
          Prefer: "count=exact",
        },
      });

      if (!Array.isArray(batch) || !batch.length) break;

      rows.push(...batch);

      if (batch.length < PAGE_SIZE) break;

      from += PAGE_SIZE;
    }

    return rows;
  }

  async function selectAllOptional(table, query = "select=*") {
    try {
      return await selectAll(table, query);
    } catch (error) {
      if (String(error.message || "").includes("PGRST205")) return [];
      throw error;
    }
  }

  async function insert(table, payload) {
    return request(table, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Prefer: "return=representation",
      },
      body: JSON.stringify(payload),
    });
  }

  async function insertMany(table, payload) {
    if (!payload.length) return [];
    return request(table, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Prefer: "return=representation",
      },
      body: JSON.stringify(payload),
    });
  }

  async function upsert(table, payload, conflictKey) {
    return request(`${table}?on_conflict=${conflictKey}`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Prefer: "resolution=merge-duplicates,return=representation",
      },
      body: JSON.stringify(payload),
    });
  }

  function mapControleToDemand(row) {
    return {
      id: row.id_demanda_controle,
      ordem: row.ordem_sap || "",
      tipoDemanda: row.tipo_demanda || "",
      descricao: row.descricao || "",
      gerencia: row.gerencia || "",
      centroTrabalho: row.centro_trabalho || "",
      localInstalacao: row.local_instalacao || "",
      vencimento: cleanDate(row.vencimento) || "",
      competencia: normalizeCompetencia(row.competencia),
      tipoOM: row.tipo_om || "",
      prioridade: normalizePrioridade(row.prioridade),
      statusSistema: row.status_sistema || "",
      statusUsuario: row.status_usuario || "",
      toleranciaMin: cleanDate(row.tolerancia_min) || "",
      toleranciaMax: cleanDate(row.tolerancia_max) || "",
      dataPlanejada: cleanDate(row.data_planejada) || "",
      dataReplanejadaAtual: cleanDate(row.data_replanejada) || "",
      dataRealizada: cleanDate(row.data_realizada) || "",
      statusOperacional: row.status_operacional || "",
      substatusOperacional: row.substatus_operacional || "",
      perda: row.perda === true,
      motivoPerda: row.motivo_perda || "",
      justificativaPerda: row.justificativa_perda || "",
      comentario: row.comentario || "",
      usuarioResponsavel: row.usuario_responsavel || "",
      dataUltimaAtualizacao:
        row.data_ultima_atualizacao || row.updated_at || "",
      origem: row.origem_informacao || row.fonte_registro || "Supabase",
      quantidadeReplanejamentos: Number(row.quantidade_replanejamentos || 0),
      frequencia: row.frequencia || "",
      observacao: row.observacao || "",
      vinculadaEm: row.vinculada_em || "",
    };
  }

  function mapDemandToControle(record) {
    const idDemanda =
      record.id || record.id_demanda_controle || stableDemandId(record);

    return {
      id_demanda_controle: idDemanda,
      ordem_sap: record.ordem ? String(record.ordem) : "",
      tipo_demanda: record.tipoDemanda || "",
      descricao: record.descricao || "",
      gerencia: record.gerencia || "",
      centro_trabalho: record.centroTrabalho || "",
      local_instalacao: record.localInstalacao || "",
      vencimento: cleanDate(record.vencimento),
      competencia: normalizeCompetencia(record.competencia),
      tipo_om: record.tipoOM || "",
      prioridade: normalizePrioridade(record.prioridade),
      status_sistema: record.statusSistema || "",
      status_usuario: record.statusUsuario || "",
      tolerancia_min: cleanDate(record.toleranciaMin),
      tolerancia_max: cleanDate(record.toleranciaMax),
      data_planejada: cleanDate(record.dataPlanejada),
      data_replanejada: cleanDate(record.dataReplanejadaAtual),
      data_realizada: cleanDate(record.dataRealizada),
      status_operacional: record.statusOperacional || "",
      substatus_operacional: record.substatusOperacional || "",
      perda: Boolean(record.perda),
      motivo_perda: record.motivoPerda || "",
      justificativa_perda: record.justificativaPerda || "",
      comentario: record.comentario || "",
      usuario_responsavel: record.usuarioResponsavel || "",
      data_ultima_atualizacao: new Date().toISOString(),
      origem_informacao: record.origem || "SISTEMA",
      quantidade_replanejamentos: Number(record.quantidadeReplanejamentos || 0),
      frequencia: record.frequencia || "",
      observacao: record.observacao || "",
      vinculada_em: record.vinculadaEm || null,
      updated_by: record.usuarioResponsavel || "",
      fonte_registro: record.origem || "SISTEMA",
    };
  }

  function mapUser(row) {
    return {
      id: row.id,
      nome: row.nome || "",
      email: row.email || "",
      matricula: row.matricula || "",
      area: row.area || "",
      perfil: row.perfil_acesso || "Visualizador",
      ativo: row.ativo !== false,
      permissaoPlanejar: row.permissao_planejar === true,
      permissaoReplanejar: row.permissao_replanejar === true,
      permissaoRealizar: row.permissao_realizar === true,
      permissaoConfigurar: row.permissao_configurar === true,
      permissaoExportar: row.permissao_exportar === true,
      permissaoCargaLote: row.permissao_carga_lote === true,
    };
  }

  function mapHistoricoPlanejamento(row) {
    return {
      id: row.id,
      demandaId: row.id_demanda_controle,
      ordem: row.ordem_sap || "",
      dataAnterior: cleanDate(row.data_anterior) || "",
      novaData: cleanDate(row.nova_data) || "",
      usuario: row.usuario || row.usuario_email || "",
      dataHora: row.data_hora_alteracao || row.created_at,
      comentario: row.comentario || "",
    };
  }

  function mapHistoricoReplanejamento(row) {
    return {
      id: row.id,
      demandaId: row.id_demanda_controle,
      ordem: row.ordem_sap || "",
      motivo: row.motivo || "",
      motivoChave: row.motivo_chave || "",
      justificativa: row.justificativa || "",
      justificativaChave: row.justificativa_chave || "",
      dataAnterior: cleanDate(row.data_anterior) || "",
      novaData: cleanDate(row.nova_data) || "",
      usuario: row.usuario || row.usuario_email || "",
      dataHora: row.data_hora_alteracao || row.created_at,
      quantidadeReplanejamentos: Number(row.quantidade_replanejamentos || 0),
      comentario: row.comentario || "",
    };
  }

  function mapHistoricoRealizadoPerda(row) {
    return {
      id: row.id,
      demandaId: row.id_demanda_controle,
      ordem: row.ordem_sap || "",
      dataRealizada: cleanDate(row.data_realizada) || "",
      perda: row.perda === true,
      motivoPerda: row.motivo_perda || row.perfil_perda || "",
      justificativaPerda: row.justificativa_perda || "",
      comentario: row.comentario || "",
      evidencia: row.evidencia || row.url_evidencia || "",
      usuario: row.usuario || row.usuario_email || "",
      dataHora: row.data_hora_registro || row.created_at,
    };
  }

  function normalizeCentroTrabalho(value) {
    return String(value || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toUpperCase()
      .trim();
  }

  function mapCentroTrabalho(row) {
    return {
      id: row.id,
      centroTrabalho: row.centro_trabalho || "",
      centroTrabalhoChave: row.centro_trabalho_chave || "",
      gerencia: row.gerencia || "",
      supervisao: row.supervisao || "",
      planejadorOM: row.planejador_om || "",
      planejadorOMEmail: row.planejador_om_email || "",
      planejadorOMMatricula: row.planejador_om_matricula || "",
      programador: row.programador || "",
      programadorEmail: row.programador_email || "",
      programadorMatricula: row.programador_matricula || "",
      area: row.area || "",
      observacao: row.observacao || "",
      ativo: row.ativo !== false,
    };
  }

  function mapLog(row) {
    return {
      id: row.id,
      acao: row.acao || "",
      dataHora: row.data_hora || row.created_at,
      usuario: row.usuario || "",
      lista: row.lista || "",
      referencia: row.referencia || "",
      detalhe: row.detalhe || row.erro || "",
      nivel: row.nivel || "INFO",
      modulo: row.modulo || "",
      status: row.status || "",
    };
  }

  class SupabaseRepository {
    constructor() {
      this.mode = "Supabase REST";
    }

    async reset() {
      return this.getAll();
    }

    async getAll() {
      const [
        demandas,
        usuarios,
        motivos,
        justificativas,
        perfisPerda,
        justificativasPerda,
        historicoPlanejamento,
        historicoReplanejamento,
        historicoRealizadoPerdas,
        logs,
        centrosTrabalho,
      ] = await Promise.all([
        selectAll(TABLES.controle, "select=*&ativo=eq.true"),
        selectAll(TABLES.usuarios, "select=*&ativo=eq.true"),
        selectAll(
          TABLES.motivos,
          "select=*&ativo=eq.true&order=ordem_exibicao.asc",
        ),
        selectAll(
          TABLES.justificativas,
          "select=*&ativo=eq.true&order=ordem_exibicao.asc",
        ),
        selectAll(
          TABLES.perfisPerda,
          "select=*&ativo=eq.true&order=ordem_exibicao.asc",
        ),
        selectAll(
          TABLES.justificativasPerda,
          "select=*&ativo=eq.true&order=ordem_exibicao.asc",
        ),
        selectAll(
          TABLES.historicoPlanejamento,
          "select=*&order=data_hora_alteracao.desc",
        ),
        selectAll(
          TABLES.historicoReplanejamento,
          "select=*&order=data_hora_alteracao.desc",
        ),
        selectAll(
          TABLES.historicoRealizadoPerdas,
          "select=*&order=data_hora_registro.desc",
        ),
        selectAll(TABLES.logs, "select=*&order=data_hora.desc"),
        selectAll(TABLES.centrosTrabalho, "select=*&order=centro_trabalho.asc"),
      ]);

      const parametrosMap = {};

      return {
        demandas: demandas.map(mapControleToDemand),
        usuarios: usuarios.map(mapUser),
        centrosTrabalho: centrosTrabalho.map(mapCentroTrabalho),
        configuracoes: {
          motivos: motivos.map((item) => ({
            id: item.chave,
            nome: item.motivo,
            ativo: item.ativo !== false,
          })),
          justificativas: justificativas.map((item) => ({
            id: item.chave,
            motivoId: item.motivo_chave,
            nome: item.justificativa,
            ativo: item.ativo !== false,
          })),
          perfisPerda: perfisPerda.map((item) => ({
            id: item.chave,
            nome: item.perfil_perda,
            ativo: item.ativo !== false,
          })),
          justificativasPerda: justificativasPerda.map((item) => ({
            id: item.chave,
            perfilId: item.perfil_chave,
            nome: item.justificativa_perda,
            ativo: item.ativo !== false,
          })),
        },
        parametros: parametrosMap,
        parametrosDisponiveis: false,
        historicoPlanejamento: historicoPlanejamento.map(
          mapHistoricoPlanejamento,
        ),
        historicoReplanejamento: historicoReplanejamento.map(
          mapHistoricoReplanejamento,
        ),
        historicoRealizadoPerdas: historicoRealizadoPerdas.map(
          mapHistoricoRealizadoPerda,
        ),
        logs: logs.map(mapLog),
      };
    }

    async upsertDemanda(record) {
      const payload = mapDemandToControle(record);
      const saved = await upsert(
        TABLES.controle,
        payload,
        "id_demanda_controle",
      );
      return saved?.[0] ? mapControleToDemand(saved[0]) : record;
    }

    async bulkUpsertDemandas(records) {
      if (!records.length) return [];

      const payload = records.map(mapDemandToControle);
      const saved = await upsert(
        TABLES.controle,
        payload,
        "id_demanda_controle",
      );

      return Array.isArray(saved) ? saved.map(mapControleToDemand) : records;
    }

    async addHistory(type, entry) {
      if (type === "planejamento") {
        const payload = {
          id_demanda_controle: entry.demandaId,
          ordem_sap: entry.ordem || "",
          data_anterior: cleanDate(entry.dataAnterior),
          nova_data: cleanDate(entry.novaData),
          usuario: entry.usuario || "",
          usuario_email: entry.usuario || "",
          data_hora_alteracao:
            cleanDateTime(entry.dataHora) || new Date().toISOString(),
          comentario: entry.comentario || "",
          tipo_alteracao: entry.tipoAlteracao || "PLANEJAMENTO_INICIAL",
          origem_alteracao: entry.origemAlteracao || "SISTEMA_WEB",
          valor_anterior_texto: entry.dataAnterior || "",
          valor_novo_texto: entry.novaData || "",
        };

        const saved = await insert(TABLES.historicoPlanejamento, payload);
        return saved?.[0] ? mapHistoricoPlanejamento(saved[0]) : payload;
      }

      if (type === "replanejamento") {
        const payload = {
          id_demanda_controle: entry.demandaId,
          ordem_sap: entry.ordem || "",
          motivo: entry.motivo || "",
          motivo_chave: entry.motivoChave || slugify(entry.motivo || ""),
          justificativa: entry.justificativa || "",
          justificativa_chave:
            entry.justificativaChave || slugify(entry.justificativa || ""),
          data_anterior: cleanDate(entry.dataAnterior),
          nova_data: cleanDate(entry.novaData),
          usuario: entry.usuario || "",
          usuario_email: entry.usuario || "",
          data_hora_alteracao:
            cleanDateTime(entry.dataHora) || new Date().toISOString(),
          quantidade_replanejamentos: Number(
            entry.quantidadeReplanejamentos || 0,
          ),
          comentario: entry.comentario || "",
          origem_alteracao: entry.origemAlteracao || "SISTEMA_WEB",
        };

        const saved = await insert(TABLES.historicoReplanejamento, payload);
        return saved?.[0] ? mapHistoricoReplanejamento(saved[0]) : payload;
      }

      if (type === "realizadoPerda") {
        const payload = {
          id_demanda_controle: entry.demandaId,
          ordem_sap: entry.ordem || "",
          data_realizada: cleanDate(entry.dataRealizada),
          perda: Boolean(entry.perda),
          perfil_perda: entry.motivoPerda || "",
          perfil_perda_chave:
            entry.motivoPerdaChave || slugify(entry.motivoPerda || ""),
          motivo_perda: entry.motivoPerda || "",
          motivo_perda_chave:
            entry.motivoPerdaChave || slugify(entry.motivoPerda || ""),
          justificativa_perda: entry.justificativaPerda || "",
          justificativa_perda_chave:
            entry.justificativaPerdaChave ||
            slugify(entry.justificativaPerda || ""),
          comentario: entry.comentario || "",
          evidencia: entry.evidencia || "",
          url_evidencia: entry.evidencia || "",
          usuario: entry.usuario || "",
          usuario_email: entry.usuario || "",
          data_hora_registro:
            cleanDateTime(entry.dataHora) || new Date().toISOString(),
          origem_alteracao: entry.origemAlteracao || "SISTEMA_WEB",
        };

        const saved = await insert(TABLES.historicoRealizadoPerdas, payload);
        return saved?.[0] ? mapHistoricoRealizadoPerda(saved[0]) : payload;
      }

      return null;
    }

    async addLog(entry) {
      const payload = {
        acao: entry.acao || "",
        usuario: entry.usuario || "",
        lista: entry.lista || "",
        referencia: entry.referencia || "",
        detalhe: entry.detalhe || "",
        nivel: entry.nivel || "INFO",
        modulo: entry.modulo || "",
        origem: entry.origem || "SISTEMA_WEB",
        status: entry.status || "SUCESSO",
        data_hora: new Date().toISOString(),
      };

      const saved = await insert(TABLES.logs, payload);
      return saved?.[0] ? mapLog(saved[0]) : payload;
    }

    async addConfigItem(group, value, parentId = "") {
      const nome = String(value || "").trim();
      if (!nome) return null;

      const chave = slugify(nome).toUpperCase().replace(/-/g, "_");

      if (group === "motivos") {
        await upsert(
          TABLES.motivos,
          {
            motivo: nome,
            chave,
            ativo: true,
            aplicavel_replanejamento: true,
          },
          "chave",
        );
      }

      if (group === "justificativas") {
        await upsert(
          TABLES.justificativas,
          {
            justificativa: nome,
            chave,
            motivo_chave: parentId,
            ativo: true,
          },
          "chave,motivo_chave",
        );
      }

      if (group === "perfisPerda") {
        await upsert(
          TABLES.perfisPerda,
          {
            perfil_perda: nome,
            chave,
            ativo: true,
          },
          "chave",
        );
      }

      if (group === "justificativasPerda") {
        await upsert(
          TABLES.justificativasPerda,
          {
            justificativa_perda: nome,
            chave,
            perfil_chave: parentId,
            ativo: true,
          },
          "chave,perfil_chave",
        );
      }

      return nome;
    }

    async addUser(user) {
      const perfil = user.perfil || user.perfil_acesso || "Visualizador";
      const defaults =
        {
          Administrador: {
            planejar: true,
            replanejar: true,
            realizar: true,
            configurar: true,
            exportar: true,
            cargaLote: true,
          },
          Editor: {
            planejar: true,
            replanejar: true,
            realizar: true,
            configurar: false,
            exportar: true,
            cargaLote: true,
          },
          Planejador: {
            planejar: true,
            replanejar: true,
            realizar: false,
            configurar: false,
            exportar: true,
            cargaLote: true,
          },
          Gestor: {
            planejar: false,
            replanejar: false,
            realizar: false,
            configurar: false,
            exportar: true,
            cargaLote: false,
          },
          Visualizador: {
            planejar: false,
            replanejar: false,
            realizar: false,
            configurar: false,
            exportar: true,
            cargaLote: false,
          },
        }[perfil] || {};
      const boolValue = (value, fallback) =>
        value === true || value === "on" || value === "true"
          ? true
          : value === false || value === "false"
            ? false
            : fallback;

      const payload = {
        nome: user.nome || "",
        email: user.email || "",
        matricula: user.matricula || "",
        area: user.area || "",
        perfil_acesso: perfil,
        ativo: user.ativo === false || user.ativo === "false" ? false : true,
        permissao_planejar: boolValue(
          user.permissaoPlanejar,
          defaults.planejar,
        ),
        permissao_replanejar: boolValue(
          user.permissaoReplanejar,
          defaults.replanejar,
        ),
        permissao_realizar: boolValue(
          user.permissaoRealizar,
          defaults.realizar,
        ),
        permissao_configurar: boolValue(
          user.permissaoConfigurar,
          defaults.configurar,
        ),
        permissao_exportar: boolValue(
          user.permissaoExportar,
          defaults.exportar,
        ),
        permissao_carga_lote: boolValue(
          user.permissaoCargaLote,
          defaults.cargaLote,
        ),
      };

      const saved = await upsert(TABLES.usuarios, payload, "email");
      return saved?.[0] ? mapUser(saved[0]) : payload;
    }

    async updateParameters(parameters) {
      await this.addLog({
        acao: "PARAMETROS_NAO_MIGRADOS",
        usuario: "",
        lista: "PARAMETROS",
        referencia: "updateParameters",
        detalhe: JSON.stringify(parameters),
        modulo: "CONFIGURACOES",
        nivel: "AVISO",
      });

      return parameters;
    }

    async createBatchRun(summary) {
      const loteId =
        summary.loteId ||
        summary.lote_id ||
        `LOTE-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;

      const payload = {
        lote_id: loteId,
        nome_arquivo: summary.nomeArquivo || summary.nome_arquivo || "",
        usuario: summary.usuario || "",
        usuario_email: summary.usuario || "",
        total_linhas: Number(summary.totalLinhas || summary.total_linhas || 0),
        linhas_validas: Number(
          summary.linhasValidas || summary.linhas_validas || 0,
        ),
        linhas_alerta: Number(
          summary.linhasAlerta || summary.linhas_alerta || 0,
        ),
        linhas_com_erro: Number(
          summary.linhasComErro || summary.linhas_com_erro || 0,
        ),
        status: summary.status || "PROCESSADO",
        data_hora: new Date().toISOString(),
        updated_at: new Date().toISOString(),
      };

      const saved = await insert(TABLES.cargasLote, payload);

      return Array.isArray(saved) ? saved[0] : payload;
    }
    async addBatchItems(loteId, items) {
      if (!loteId || !items?.length) return [];

      const payload = items.map((item) => ({
        lote_id: loteId,
        linha: Number(item.line || item.linha || 0),
        status: item.status || "",
        mensagem_validacao: item.message || item.mensagem || "",
        id_demanda_controle: item.record?.id || item.idDemandaControle || "",
        ordem_sap: item.record?.ordem ? String(item.record.ordem) : "",
        payload_original: item.record || {},
        created_at: new Date().toISOString(),
      }));

      return insertMany(TABLES.cargasLoteItens, payload);
    }
    async getCurrentUserIdentity() {
      const params = new URLSearchParams(global.location.search);
      const email = params.get("user") || "";

      return {
        email,
        nome: email || "Usuário",
      };
    }
    async upsertCentroTrabalho(record) {
      const centroTrabalho = String(record.centroTrabalho || "").trim();

      if (!centroTrabalho) {
        throw new Error("Centro de trabalho é obrigatório.");
      }

      const payload = {
        centro_trabalho: centroTrabalho,
        centro_trabalho_chave: normalizeCentroTrabalho(centroTrabalho),
        gerencia: record.gerencia || "",
        supervisao: record.supervisao || "",

        planejador_om: record.planejadorOM || "",
        planejador_om_email: record.planejadorOMEmail || "",
        planejador_om_matricula: record.planejadorOMMatricula || "",

        programador: record.programador || "",
        programador_email: record.programadorEmail || "",
        programador_matricula: record.programadorMatricula || "",

        area: record.area || "",
        observacao: record.observacao || "",
        ativo: record.ativo !== false,

        created_by: record.usuario || "",
        updated_by: record.usuario || "",
      };

      const saved = await upsert(
        TABLES.centrosTrabalho,
        payload,
        "centro_trabalho_chave",
      );

      return saved?.[0] ? mapCentroTrabalho(saved[0]) : record;
    }
  }

  const previous = global.CCEData || {};

  global.CCEData = {
    ...previous,
    slugify: previous.slugify || slugify,
    stableDemandId: previous.stableDemandId || stableDemandId,
    normalizeText: previous.normalizeText || normalizeText,
    normalizeCompetencia: previous.normalizeCompetencia || normalizeCompetencia,
    normalizePrioridade: previous.normalizePrioridade || normalizePrioridade,
    createRepository() {
      return new SupabaseRepository();
    },
  };

  global.CCESupabase = {
    url: SUPABASE_URL,
    tables: TABLES,
    selectAll,
    request,
  };
})(window);

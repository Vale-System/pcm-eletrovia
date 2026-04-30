(function seedSampleData(global) {
  const gerencias = ["GME Norte", "GME Sul", "GME Corredor", "GME Oficina"];
  const centros = ["EVT-PCM-01", "EVT-PCM-02", "EVT-LIN-03", "EVT-ENE-04", "EVT-SUB-05"];
  const locais = ["ELV-TR-0001", "ELV-TR-0002", "ELV-LC-0104", "ELV-SE-0207", "ELV-PN-0302", "ELV-OF-0401"];
  const tipos = ["Preventiva", "Corretiva", "Inspeção", "Calibração", "Sistemática"];
  const prioridades = ["Alta", "Média", "Baixa"];
  const statusSistema = ["ABER", "LIB", "PROG", "AGPR", "ENCE"];
  const descricoes = [
    "Inspeção de via energizada e aterramento",
    "Preventiva em chave seccionadora",
    "Revisão de banco de baterias",
    "Correção de falha em circuito de comando",
    "Medição de isolamento em alimentador",
    "Limpeza técnica em painel de proteção",
    "Aferição de relé de sobrecorrente",
    "Substituição programada de conector",
    "Inspeção termográfica em subestação",
    "Rotina sistemática de lubrificação",
    "Revisão de sinalização operacional",
    "Teste funcional de intertravamento"
  ];

  const users = [
    {
      id: "USR-001",
      nome: "Weslley Santos",
      email: "weslley.santos@vale.com",
      matricula: "V102938",
      area: "Central Eletrovia",
      perfil: "Administrador",
      ativo: true
    },
    {
      id: "USR-002",
      nome: "Ana Planejamento",
      email: "ana.planejamento@vale.com",
      matricula: "V204560",
      area: "PCM",
      perfil: "Editor",
      ativo: true
    },
    {
      id: "USR-003",
      nome: "Carlos Gestor",
      email: "carlos.gestor@vale.com",
      matricula: "V305671",
      area: "Gerência Eletrovia",
      perfil: "Gestor",
      ativo: true
    },
    {
      id: "USR-004",
      nome: "Marina Consulta",
      email: "marina.consulta@vale.com",
      matricula: "V409812",
      area: "Operação",
      perfil: "Visualizador",
      ativo: true
    }
  ];

  const configs = {
    motivos: [
      { id: "janela-operacional", nome: "Conflito com janela operacional", ativo: true },
      { id: "material-indisponivel", nome: "Material indisponível", ativo: true },
      { id: "equipe-indisponivel", nome: "Equipe indisponível", ativo: true },
      { id: "priorizacao-emergencial", nome: "Priorização emergencial", ativo: true },
      { id: "bloqueio-acesso", nome: "Bloqueio de acesso", ativo: true },
      { id: "condicao-climatica", nome: "Condição climática", ativo: true }
    ],
    justificativas: [
      { id: "dentro-tolerancia", motivoId: "janela-operacional", nome: "Replanejado dentro da tolerância", ativo: true },
      { id: "seguranca-operacional", motivoId: "janela-operacional", nome: "Necessário preservar segurança operacional", ativo: true },
      { id: "liberacao-operacao", motivoId: "bloqueio-acesso", nome: "Dependência de liberação da operação", ativo: true },
      { id: "regularizacao-material", motivoId: "material-indisponivel", nome: "Aguardando regularização de material", ativo: true },
      { id: "agrupado-rota", motivoId: "equipe-indisponivel", nome: "Atendimento agrupado por rota", ativo: true }
    ],
    perfisPerda: [
      { id: "perda-operacional", nome: "Perda Operacional", ativo: true },
      { id: "perda-pcm", nome: "Perda PCM", ativo: true },
      { id: "perda-manutencao", nome: "Perda Manutenção", ativo: true },
      { id: "perda-externa", nome: "Perda Externa", ativo: true }
    ],
    justificativasPerda: [
      { id: "janela-cancelada", perfilId: "perda-operacional", nome: "Janela operacional cancelada", ativo: true },
      { id: "nao-executada-prazo", perfilId: "perda-pcm", nome: "Ordem não executada no prazo", ativo: true },
      { id: "equipe-emergencia", perfilId: "perda-manutencao", nome: "Equipe deslocada para emergência", ativo: true },
      { id: "material-critico", perfilId: "perda-pcm", nome: "Material crítico não entregue", ativo: true },
      { id: "demanda-cancelada", perfilId: "perda-externa", nome: "Demanda cancelada pela área solicitante", ativo: true }
    ]
  };

  const parameters = {
    pageSizeDefault: 12,
    currentCompetencia: "2026-04",
    defaultToleranceBeforeDays: 3,
    defaultToleranceAfterDays: 5,
    mandatoryInitialScope: "Abertas, competência atual e vencimentos próximos",
    sharePointLibrary: "Documentos Compartilhados/SAP_BO",
    sapExcelFileName: "base_ordens_sap.xlsx",
    realizedExcelFileName: "base_realizados_sap.xlsx"
  };

  function pad(value) {
    return String(value).padStart(2, "0");
  }

  function addDays(base, days) {
    const date = new Date(`${base}T12:00:00`);
    date.setDate(date.getDate() + days);
    return date.toISOString().slice(0, 10);
  }

  function competenciaFromDate(dateText) {
    return dateText.slice(0, 7);
  }

  const baseDate = "2026-04-10";
  const demands = Array.from({ length: 64 }, (_, index) => {
    const vencimento = addDays(baseDate, index - 18);
    const hasPlanned = index % 5 !== 0;
    const hasReplanned = index % 7 === 0;
    const hasRealized = index % 8 === 0 || index % 11 === 0;
    const isLoss = index % 13 === 0;
    const planned = hasPlanned ? addDays(vencimento, -((index % 4) + 1)) : "";
    const replanned = hasReplanned ? addDays(vencimento, (index % 5) - 1) : "";
    const realized = hasRealized ? addDays(vencimento, index % 11 === 0 ? 8 : (index % 4) - 1) : "";
    const centro = centros[index % centros.length];
    const local = locais[index % locais.length];
    const tipoOM = tipos[index % tipos.length];
    const ordem = index < 56 ? String(910000000 + index * 17) : "";

    return {
      id: `DEM-2026-${pad(index + 1)}`,
      ordem,
      tipoDemanda: ordem ? "SAP BO" : "Futura",
      descricao: descricoes[index % descricoes.length],
      centroTrabalho: centro,
      localInstalacao: local,
      gerencia: gerencias[index % gerencias.length],
      vencimento,
      prioridade: prioridades[index % prioridades.length],
      statusSistema: statusSistema[index % statusSistema.length],
      competencia: competenciaFromDate(vencimento),
      tipoOM,
      toleranciaMin: addDays(vencimento, -3),
      toleranciaMax: addDays(vencimento, 5),
      dataPlanejada: planned,
      dataReplanejadaAtual: replanned,
      dataRealizada: realized,
      perda: isLoss,
      motivoPerda: isLoss ? configs.perfisPerda[index % configs.perfisPerda.length].nome : "",
      justificativaPerda: isLoss ? configs.justificativasPerda[index % configs.justificativasPerda.length].nome : "",
      comentario: index % 6 === 0 ? "Acompanhar disponibilidade da janela operacional." : "",
      usuarioResponsavel: users[index % users.length].email,
      dataUltimaAtualizacao: addDays("2026-04-01", index % 22),
      origem: ordem ? "SAP BO" : "Demanda Antecipada",
      quantidadeReplanejamentos: hasReplanned ? (index % 3) + 1 : 0,
      frequencia: ordem ? "" : ["Mensal", "Trimestral", "Semestral"][index % 3],
      observacao: ordem ? "" : "Demanda cadastrada antes da geração da ordem SAP.",
      vinculadaEm: ""
    };
  });

  demands.push(
    {
      id: "DF-2026-ELV-010",
      ordem: "",
      tipoDemanda: "Sistemática",
      descricao: "Rotina sistemática futura de inspeção do alimentador principal",
      centroTrabalho: "EVT-ENE-04",
      localInstalacao: "ELV-SE-0207",
      gerencia: "GME Corredor",
      vencimento: "2026-06-18",
      prioridade: "Média",
      statusSistema: "PREV",
      competencia: "2026-06",
      tipoOM: "Sistemática",
      toleranciaMin: "2026-06-15",
      toleranciaMax: "2026-06-23",
      dataPlanejada: "2026-06-16",
      dataReplanejadaAtual: "",
      dataRealizada: "",
      perda: false,
      motivoPerda: "",
      justificativaPerda: "",
      comentario: "Criada para preservar planejamento antes da ordem SAP.",
      usuarioResponsavel: "ana.planejamento@vale.com",
      dataUltimaAtualizacao: "2026-04-23",
      origem: "Demanda Antecipada",
      quantidadeReplanejamentos: 0,
      frequencia: "Mensal",
      observacao: "Vincular quando a ordem oficial aparecer.",
      vinculadaEm: ""
    },
    {
      id: "DEM-2026-SUG-001",
      ordem: "910000986",
      tipoDemanda: "SAP BO",
      descricao: "Inspeção do alimentador principal com revisão de proteção",
      centroTrabalho: "EVT-ENE-04",
      localInstalacao: "ELV-SE-0207",
      gerencia: "GME Corredor",
      vencimento: "2026-06-19",
      prioridade: "Alta",
      statusSistema: "ABER",
      competencia: "2026-06",
      tipoOM: "Inspeção",
      toleranciaMin: "2026-06-16",
      toleranciaMax: "2026-06-24",
      dataPlanejada: "",
      dataReplanejadaAtual: "",
      dataRealizada: "",
      perda: false,
      motivoPerda: "",
      justificativaPerda: "",
      comentario: "",
      usuarioResponsavel: "",
      dataUltimaAtualizacao: "2026-04-24",
      origem: "SAP BO",
      quantidadeReplanejamentos: 0,
      frequencia: "",
      observacao: "",
      vinculadaEm: ""
    }
  );

  const historicoPlanejamento = [
    {
      id: "HP-001",
      demandaId: "DEM-2026-02",
      dataAnterior: "",
      novaData: "2026-03-25",
      usuario: "ana.planejamento@vale.com",
      dataHora: "2026-03-20T09:14:00",
      comentario: "Planejamento inicial pela competência."
    }
  ];

  const historicoReplanejamento = [
    {
      id: "HR-001",
      demandaId: "DEM-2026-08",
      motivo: "Conflito com janela operacional",
      justificativa: "Necessário preservar segurança operacional",
      dataAnterior: "2026-03-30",
      novaData: "2026-04-02",
      usuario: "ana.planejamento@vale.com",
      dataHora: "2026-03-29T15:44:00",
      quantidadeReplanejamentos: 1
    }
  ];

  const historicoRealizadoPerdas = [
    {
      id: "HPR-001",
      demandaId: "DEM-2026-01",
      dataRealizada: "2026-03-24",
      perda: false,
      motivoPerda: "",
      justificativaPerda: "",
      comentario: "Executada dentro da janela.",
      evidencia: "",
      usuario: "ana.planejamento@vale.com",
      dataHora: "2026-03-24T17:05:00"
    }
  ];

  const logs = [
    {
      id: "LOG-001",
      dataHora: "2026-04-24T08:00:00",
      usuario: "weslley.santos@vale.com",
      acao: "Acesso",
      lista: "Logs_Sistema",
      referencia: "Sessão",
      detalhe: "Acesso inicial ao painel operacional."
    },
    {
      id: "LOG-002",
      dataHora: "2026-04-24T08:05:00",
      usuario: "weslley.santos@vale.com",
      acao: "Importação",
      lista: "Base_Ordens_SAP",
      referencia: "base_ordens_sap.xlsx",
      detalhe: "Carga SAP BO validada para demonstração."
    }
  ];

  global.CCE_SAMPLE = {
    demands,
    users,
    configs,
    parameters,
    historicoPlanejamento,
    historicoReplanejamento,
    historicoRealizadoPerdas,
    logs
  };
})(window);

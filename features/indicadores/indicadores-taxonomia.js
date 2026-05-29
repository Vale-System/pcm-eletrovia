(function setupTaxonomia(global) {
  "use strict";

  const DISCIPLINAS = {
    "Via Permanente": {
      icone: "ti-train",
      descricao:
        "Atividades ligadas à via, AMV, trilhos, dormentes, geometria, juntas, lastro e soldagem.",
      subtipos: [
        "Manutenção de Mecanismos (MCH/JAC)",
        "Lubrificação de AMV/Curva",
        "Inspeção de Risco (AMV/Trilho)",
        "Manutenção Preventiva AMV",
        "Juntas Isolantes (JIC)",
        "Substituição de Dormentes",
        "Esmerilhamento de Trilhos",
        "Substituição de Trilhos",
        "Geometria e Socaria",
        "Fixação, Lastro e Pequenas Peças",
        "Soldagem de Trilhos (TLS)",
        "Geral",
      ],
    },

    Eletroeletrônica: {
      icone: "ti-cpu",
      descricao:
        "Atividades ligadas à sinalização, circuitos de via, máquinas de chave, ATC, telemetria, telecomunicações e eventos eletroeletrônicos.",
      subtipos: [
        "Inspeção e Preventiva",
        "Máquinas de Chave (MCH/AMV)",
        "Circuitos de Via e Detecção",
        "Telemetria e Sistemas Embarcados",
        "Telecomunicações e CCTV",
        "Sinalização e ATC",
        "Corretivas e Eventos",
        "Geral",
      ],
    },

    "Força & Energia": {
      icone: "ti-bolt",
      descricao:
        "Atividades ligadas a energia, iluminação, subestações, aterramento, UPS, geradores, tração e cancelas elétricas.",
      subtipos: [
        "Iluminação",
        "Proteção Catódica e Aterramento",
        "Subestações e Transformadores",
        "Geradores e UPS/Nobreak",
        "Equipamentos de Tração",
        "Cancelas elétricas",
        "Geral",
      ],
    },

    Infraestrutura: {
      icone: "ti-mountain",
      descricao:
        "Atividades civis ligadas a drenagem, bueiros, taludes, aterros, roço, poda, acessos, edificações e rede de distribuição.",
      subtipos: [
        "Bueiros e Drenagem",
        "Aterros, Taludes e Cortes",
        "Passagens em Nível",
        "Roço e Poda de Vegetação",
        "Manutenção de Rede de Distribuição",
        "Edificações e Acessos",
        "Geral",
      ],
    },
    "Estaleiro de Solda": {
      icone: "ti-tool",
      descricao:
        "Atividades ligadas ao estaleiro de solda, soldagem de trilhos, rolo motor e apoio operacional da solda.",
      subtipos: ["Rolo Motor", "Soldagem", "Geral"],
    },
  };

  const PREFIXOS_DISCIPLINA = [
    {
      disciplina: "Força & Energia",
      prefixos: ["ECFE"],
    },
    {
      disciplina: "Eletroeletrônica",
      prefixos: ["ECEE", "CSAT", "CCU", "INT"],
    },
    {
      disciplina: "Via Permanente",
      prefixos: [
        "CIN",
        "CCG",
        "CDL",
        "CDA",
        "CDM",
        "CIV",
        "CMA",
        "CRG",
        "CRN",
        "CSC",
        "CTG",
        "CTT",
        "CVI",
      ],
    },
    {
      disciplina: "Infraestrutura",
      prefixos: ["CIF", "CMI", "CSV"],
    },
    {
      disciplina: "Estaleiro de Solda",
      prefixos: ["CES"],
    },
  ];

  const REGRAS_SUBTIPO = {
    "Via Permanente": [
      {
        subtipo: "Manutenção de Mecanismos (MCH/JAC)",
        termos: [
          "MCH",
          "MECANISMO",
          "MECANISMOS",
          "MAQUINA DE CHAVE",
          "MAQUINAS DE CHAVE",
          "MÁQUINA DE CHAVE",
          "MÁQUINAS DE CHAVE",
          "JAC",
          "JACARE",
          "JACARÉ",
        ],
      },
      {
        subtipo: "Lubrificação de AMV/Curva",
        termos: [
          "LUBRIFICACAO",
          "LUBRIFICAÇÃO",
          "LUBRIFICAR",
          "GRAXA",
          "AMV",
          "CURVA",
        ],
      },
      {
        subtipo: "Inspeção de Risco (AMV/Trilho)",
        termos: [
          "INSPECAO DE RISCO",
          "INSPEÇÃO DE RISCO",
          "RISCO AMV",
          "RISCO TRILHO",
          "INSPECAO AMV",
          "INSPEÇÃO AMV",
          "INSPECAO TRILHO",
          "INSPEÇÃO TRILHO",
        ],
      },
      {
        subtipo: "Manutenção Preventiva AMV",
        termos: [
          "PREVENTIVA AMV",
          "MANUTENCAO PREVENTIVA AMV",
          "MANUTENÇÃO PREVENTIVA AMV",
          "MP AMV",
          "AMV",
        ],
      },
      {
        subtipo: "Juntas Isolantes (JIC)",
        termos: [
          "JIC",
          "JUNTA ISOLANTE",
          "JUNTAS ISOLANTES",
          "JUNTA ISOLADA",
          "JUNTA",
        ],
      },
      {
        subtipo: "Substituição de Dormentes",
        termos: [
          "DORMENTE",
          "DORMENTES",
          "SUBSTITUICAO DE DORMENTE",
          "SUBSTITUIÇÃO DE DORMENTE",
          "TROCA DE DORMENTE",
        ],
      },
      {
        subtipo: "Esmerilhamento de Trilhos",
        termos: [
          "ESMERILHAMENTO",
          "ESMERILHAR",
          "ESMERIL",
          "TRILHO ESMERILHADO",
        ],
      },
      {
        subtipo: "Substituição de Trilhos",
        termos: [
          "SUBSTITUICAO DE TRILHO",
          "SUBSTITUIÇÃO DE TRILHO",
          "TROCA DE TRILHO",
          "TROCA TRILHO",
          "BARRA LONGA",
          "TRILHO",
        ],
      },
      {
        subtipo: "Geometria e Socaria",
        termos: [
          "GEOMETRIA",
          "SOCARIA",
          "ALINHAMENTO",
          "NIVELAMENTO",
          "CONFORMACAO",
          "CONFORMAÇÃO",
          "BITOLA",
          "EMPENO",
        ],
      },
      {
        subtipo: "Fixação, Lastro e Pequenas Peças",
        termos: [
          "FIXACAO",
          "FIXAÇÃO",
          "GRAMPO",
          "GRAMPO ELASTICO",
          "GRAMPO ELÁSTICO",
          "TIREFAO",
          "TIREFÃO",
          "PLACA",
          "LASTRO",
          "BRITA",
          "TALA",
          "PARAFUSO",
          "PEQUENAS PECAS",
          "PEQUENAS PEÇAS",
        ],
      },
      {
        subtipo: "Soldagem de Trilhos (TLS)",
        termos: [
          "TLS",
          "SOLDA",
          "SOLDAGEM",
          "SOLDAGEM DE TRILHO",
          "SOLDA ALUMINOTERMICA",
          "SOLDA ALUMINOTÉRMICA",
        ],
      },
    ],

    Eletroeletrônica: [
      {
        subtipo: "Inspeção e Preventiva",
        termos: [
          "INSPECAO",
          "INSPEÇÃO",
          "PREVENTIVA",
          "MANUTENCAO PREVENTIVA",
          "MANUTENÇÃO PREVENTIVA",
          "ROTINA",
          "CHECKLIST",
        ],
      },
      {
        subtipo: "Máquinas de Chave (MCH/AMV)",
        termos: [
          "MCH",
          "MAQUINA DE CHAVE",
          "MÁQUINA DE CHAVE",
          "MAQUINAS DE CHAVE",
          "MÁQUINAS DE CHAVE",
          "AMV",
          "ACIONAMENTO",
          "CHAVE ELETRICA",
          "CHAVE ELÉTRICA",
        ],
      },
      {
        subtipo: "Circuitos de Via e Detecção",
        termos: [
          "CIRCUITO DE VIA",
          "CIRCUITOS DE VIA",
          "DETECCAO",
          "DETECÇÃO",
          "OCUPACAO",
          "OCUPAÇÃO",
          "TRACK CIRCUIT",
          "EIXO",
          "CONTADOR DE EIXO",
        ],
      },
      {
        subtipo: "Telemetria e Sistemas Embarcados",
        termos: [
          "TELEMETRIA",
          "EMBARCADO",
          "EMBARCADOS",
          "RADIO MODEM",
          "RÁDIO MODEM",
          "MODEM",
          "IHM",
          "CLP",
          "PLC",
        ],
      },
      {
        subtipo: "Telecomunicações e CCTV",
        termos: [
          "TELECOM",
          "TELECOMUNICACAO",
          "TELECOMUNICAÇÃO",
          "CCTV",
          "CAMERA",
          "CÂMERA",
          "FIBRA",
          "RADIO",
          "RÁDIO",
          "ANTENA",
          "REDE",
          "SWITCH",
        ],
      },
      {
        subtipo: "Sinalização e ATC",
        termos: [
          "SINALIZACAO",
          "SINALIZAÇÃO",
          "ATC",
          "SINAL",
          "SEMÁFORO",
          "SEMAFORO",
          "LICENCIAMENTO",
          "INTERTRAVAMENTO",
        ],
      },
      {
        subtipo: "Corretivas e Eventos",
        termos: [
          "CORRETIVA",
          "FALHA",
          "EVENTO",
          "EMERGENCIAL",
          "ANOMALIA",
          "OCORRENCIA",
          "OCORRÊNCIA",
        ],
      },
    ],

    "Força & Energia": [
      {
        subtipo: "Iluminação",
        termos: [
          "ILUMINACAO",
          "ILUMINAÇÃO",
          "LUMINARIA",
          "LUMINÁRIA",
          "REFLETOR",
          "POSTE",
          "LAMPADA",
          "LÂMPADA",
        ],
      },
      {
        subtipo: "Proteção Catódica e Aterramento",
        termos: [
          "PROTECAO CATODICA",
          "PROTEÇÃO CATÓDICA",
          "CATODICA",
          "CATÓDICA",
          "ATERRAMENTO",
          "MALHA DE TERRA",
          "SPDA",
          "PARA-RAIO",
          "PARA RAIO",
          "DESCARGA ATMOSFERICA",
          "DESCARGA ATMOSFÉRICA",
        ],
      },
      {
        subtipo: "Subestações e Transformadores",
        termos: [
          "SUBESTACAO",
          "SUBESTAÇÃO",
          "SE ",
          "TRAFO",
          "TRANSFORMADOR",
          "TRANSFORMADORES",
          "DISJUNTOR",
          "RELIGADOR",
          "CUBICULO",
          "CUBÍCULO",
        ],
      },
      {
        subtipo: "Geradores e UPS/Nobreak",
        termos: [
          "GERADOR",
          "GERADORES",
          "GRUPO GERADOR",
          "UPS",
          "NOBREAK",
          "NO BREAK",
          "BANCO DE BATERIA",
          "BATERIA",
          "RETIFICADOR",
        ],
      },
      {
        subtipo: "Equipamentos de Tração",
        termos: [
          "TRACAO",
          "TRAÇÃO",
          "REDE AEREA",
          "REDE AÉREA",
          "CATENARIA",
          "CATENÁRIA",
          "PANTOGRAFO",
          "PANTÓGRAFO",
          "ALIMENTADOR",
          "SECCIONADORA",
        ],
      },
      {
        subtipo: "Cancelas elétricas",
        termos: [
          "CANCELA",
          "CANCELAS",
          "PN ELETRICA",
          "PN ELÉTRICA",
          "PASSAGEM EM NIVEL ELETRICA",
          "PASSAGEM EM NÍVEL ELÉTRICA",
          "SINALIZADOR DE PN",
        ],
      },
    ],

    Infraestrutura: [
      {
        subtipo: "Bueiros e Drenagem",
        termos: [
          "BUEIRO",
          "BUEIROS",
          "DRENAGEM",
          "DRENO",
          "VALETA",
          "CANALETA",
          "GALERIA",
          "OBRA DE ARTE CORRENTE",
          "OAC",
        ],
      },
      {
        subtipo: "Aterros, Taludes e Cortes",
        termos: [
          "ATERRO",
          "ATERROS",
          "TALUDE",
          "TALUDES",
          "CORTE",
          "CORTES",
          "EROSAO",
          "EROSÃO",
          "ESCORREGAMENTO",
          "CONTENCAO",
          "CONTENÇÃO",
        ],
      },
      {
        subtipo: "Passagens em Nível",
        termos: [
          "PASSAGEM EM NIVEL",
          "PASSAGEM EM NÍVEL",
          "PN",
          "PAVIMENTO",
          "PAVIMENTACAO",
          "PAVIMENTAÇÃO",
          "PLACA DE CONCRETO",
          "ACESSO DE PN",
        ],
      },
      {
        subtipo: "Roço e Poda de Vegetação",
        termos: [
          "ROCO",
          "ROÇO",
          "PODA",
          "VEGETACAO",
          "VEGETAÇÃO",
          "CAPINA",
          "MATO",
          "ARVORE",
          "ÁRVORE",
          "ARBUSTO",
          "SUPRESSAO VEGETAL",
          "SUPRESSÃO VEGETAL",
        ],
      },
      {
        subtipo: "Manutenção de Rede de Distribuição",
        termos: [
          "REDE DE DISTRIBUICAO",
          "REDE DE DISTRIBUIÇÃO",
          "REDE DISTRIBUICAO",
          "REDE DISTRIBUIÇÃO",
          "ADUTORA",
          "DISTRIBUICAO",
          "DISTRIBUIÇÃO",
          "TUBULACAO",
          "TUBULAÇÃO",
        ],
      },
      {
        subtipo: "Edificações e Acessos",
        termos: [
          "EDIFICACAO",
          "EDIFICAÇÃO",
          "EDIFICACOES",
          "EDIFICAÇÕES",
          "PREDIO",
          "PRÉDIO",
          "ACESSO",
          "ACESSOS",
          "ESCADA",
          "PASSARELA",
          "CERCA",
          "PORTAO",
          "PORTÃO",
        ],
      },
    ],
    "Estaleiro de Solda": [
      {
        subtipo: "Rolo Motor",
        termos: ["ROLO MOTOR", "ROLO MOTO"],
      },
      {
        subtipo: "Soldagem",
        termos: ["SOLDAGEM", "TRILHOS", "SOLDA"],
      },
    ],
  };

  function normalizar(value) {
    return String(value ?? "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toUpperCase()
      .replace(/\s+/g, " ")
      .trim();
  }

  function normalizarCentro(value) {
    return normalizar(value).replace(/[^A-Z0-9]/g, "");
  }

  function pegarPrimeiroCampo(item, campos) {
    for (const campo of campos) {
      const valor = item?.[campo];
      if (valor !== undefined && valor !== null && String(valor).trim()) {
        return valor;
      }
    }

    return "";
  }

  function textoDaDemanda(item) {
    return [
      pegarPrimeiroCampo(item, [
        "descricao",
        "descrição",
        "descricaoOrdem",
        "textoBreve",
        "Texto breve",
        "texto",
        "observacao",
        "observação",
        "comentario",
        "comentário",
      ]),
      pegarPrimeiroCampo(item, [
        "tipoOM",
        "tipoOrdem",
        "tipoDemanda",
        "atividade",
        "categoria",
      ]),
    ].join(" ");
  }

  function centroDaDemanda(item) {
    return pegarPrimeiroCampo(item, [
      "centroTrabalho",
      "centro_trabalho",
      "Centro de trabalho",
      "Centro Trabalho",
      "centro",
      "workCenter",
      "work_center",
    ]);
  }

  function identificarDisciplinaPorCentro(centroTrabalho) {
    const centro = normalizarCentro(centroTrabalho);

    if (!centro) return "";

    const regra = PREFIXOS_DISCIPLINA.find((item) =>
      item.prefixos.some((prefixo) => centro.startsWith(prefixo)),
    );

    return regra?.disciplina || "";
  }

  function contemTermo(textoNormalizado, termo) {
    const termoNormalizado = normalizar(termo);

    if (!termoNormalizado) return false;

    return textoNormalizado.includes(termoNormalizado);
  }

  function identificarSubtipoPorDisciplina(disciplina, textoNormalizado) {
    const regras = REGRAS_SUBTIPO[disciplina] || [];

    for (const regra of regras) {
      if (regra.termos.some((termo) => contemTermo(textoNormalizado, termo))) {
        return regra.subtipo;
      }
    }

    return "Geral";
  }

  function identificarDisciplinaPorTexto(textoNormalizado) {
    for (const disciplina of Object.keys(REGRAS_SUBTIPO)) {
      const subtipo = identificarSubtipoPorDisciplina(
        disciplina,
        textoNormalizado,
      );

      if (subtipo !== "Geral") {
        return disciplina;
      }
    }

    return "Via Permanente";
  }

  function classificarItem(item) {
    const centroTrabalho = centroDaDemanda(item);
    const textoNormalizado = normalizar(textoDaDemanda(item));

    const disciplinaPorCentro = identificarDisciplinaPorCentro(centroTrabalho);

    const disciplina =
      disciplinaPorCentro || identificarDisciplinaPorTexto(textoNormalizado);

    const subtipo = identificarSubtipoPorDisciplina(
      disciplina,
      textoNormalizado,
    );

    return {
      ...item,
      disciplina,
      subtipo,
      taxonomiaOrigem: disciplinaPorCentro ? "Centro de Trabalho" : "Descrição",
      taxonomiaCentroTrabalho: centroTrabalho,
    };
  }

  function classificar(demandas) {
    if (!Array.isArray(demandas)) return [];

    return demandas.map(classificarItem);
  }

  global.CCETaxonomia = {
    classificar,
    DISCIPLINAS,
  };
})(window);

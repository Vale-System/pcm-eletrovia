(function setupDiagnosticoPDF(global) {
  "use strict";

  const PAGE = {
    width: 297,
    height: 210,
    marginX: 12,
    contentTop: 24,
    contentBottom: 194,
  };

  const SECTION_ALIASES = Object.freeze({
    capa: ["capa", "capaExecutiva", "diagSecaoCapa"],
    resumo: ["resumo", "resumoExecutivo", "diagSecaoResumoExecutivo"],
    kpis: ["kpis", "kpisCarteira", "diagSecaoKpis"],
    planejamento: [
      "planejamento",
      "planejamentoPorData",
      "diagSecaoPlanejamento",
    ],
    centros: ["centros", "analiseCentros", "diagSecaoCentros"],
    planejadores: [
      "planejadores",
      "analisePlanejadores",
      "diagSecaoPlanejadores",
    ],
    km: ["km", "analiseKmTrecho", "diagSecaoKm"],
    patios: ["patios", "analisePatios", "diagSecaoPatios"],
    tolerancias: [
      "tolerancias",
      "toleranciasJanelas",
      "diagSecaoTolerancias",
    ],
    vencimentos: [
      "vencimentos",
      "vencimentosAtrasos",
      "diagSecaoVencimentos",
    ],
    clima: ["clima", "riscoClimatico", "diagSecaoClima"],
    listaCritica: ["listaCritica", "diagSecaoListaCritica"],
    recomendacoes: [
      "recomendacoes",
      "recomendacoesTecnicas",
      "diagSecaoRecomendacoes",
    ],
  });

  const THEME = {
    green: [13, 92, 68],
    greenDark: [9, 66, 49],
    greenSoft: [237, 245, 241],
    dark: [23, 52, 42],
    text: [28, 45, 36],
    muted: [93, 112, 104],
    gray: [100, 116, 139],
    lightGray: [245, 247, 250],
    lightBlue: [239, 246, 255],
    border: [219, 228, 223],
    white: [255, 255, 255],
    red: [180, 35, 24],
    redSoft: [252, 232, 230],
    orange: [201, 122, 0],
    orangeSoft: [255, 247, 237],
    blue: [29, 78, 216],
    blueSoft: [239, 246, 255],
    darkSoft: [241, 245, 249],
  };

  let pdfContext = {};

  function safeText(value, fallback = "-") {
    const text = String(value ?? "").trim();
    return text || fallback;
  }

  function formatNumber(value) {
    return new Intl.NumberFormat("pt-BR").format(Number(value) || 0);
  }

  function formatPercent(value) {
    const numeric = Number(value) || 0;
    return `${numeric.toFixed(1).replace(".", ",")}%`;
  }

  function formatDateBR(value) {
    if (!value) return "-";
    if (String(value).includes("T")) {
      const date = new Date(value);
      if (!Number.isNaN(date.getTime())) {
        return new Intl.DateTimeFormat("pt-BR", {
          day: "2-digit",
          month: "2-digit",
          year: "numeric",
          hour: "2-digit",
          minute: "2-digit",
        }).format(date);
      }
    }

    const text = String(value).trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
      const [yyyy, mm, dd] = text.split("-");
      return `${dd}/${mm}/${yyyy}`;
    }

    return text;
  }

  function splitText(doc, text, maxWidth) {
    return doc.splitTextToSize(safeText(text, ""), maxWidth);
  }

  function truncateTextValue(value, max = 80) {
    const text = safeText(value, "");
    if (!text) return "-";
    return text.length > max ? `${text.slice(0, max - 1)}...` : text;
  }

  function normalizePatioPdfText(value) {
    return String(value ?? "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toUpperCase()
      .replace(/\s+/g, " ")
      .trim();
  }

  function isPatioText(value) {
    const text = normalizePatioPdfText(value);
    const compact = text.replace(/[^A-Z0-9]/g, "");

    if (!text) return false;

    if (/PM[12]/.test(compact) || /CESSL[0-9]+/.test(compact)) return true;

    const patioTokens = [
      "PATIO",
      "QPMPA",
      "QPEPA",
      "QPMPTL",
      "QPEPC",
      "QPM",
      "QPE",
      "ESTAISL",
      "ESTAI",
    ];

    return patioTokens.some((token) => text.includes(token));
  }

  function slugFileName(value) {
    return safeText(value, "relatorio")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-zA-Z0-9]+/g, "_")
      .replace(/^_+|_+$/g, "");
  }

  function reportTypeLabel(rawType) {
    const type = safeText(rawType, "executivo").toLowerCase();
    const configTypes = global.CCEDiagnosticoConfig?.reportTypes || [];
    const found = configTypes.find((item) => item.id === type);
    return safeText(found?.label, type);
  }

  function downloadFileName(diagnostico, config = {}) {
    const reportType = reportTypeLabel(
      config.tipoRelatorio || diagnostico?.meta?.tipoRelatorio || "executivo",
    );
    const slugType = slugFileName(reportType) || "Executivo";
    const date = safeText(config.criadoEm || diagnostico?.meta?.criadoEm, "")
      .slice(0, 10)
      .replace(/[^0-9-]/g, "") || new Date().toISOString().slice(0, 10);

    return `Diagnostico_Carteira_${slugType}_${date}.pdf`;
  }

  function ensureLibraries() {
    if (!global.jspdf?.jsPDF) {
      throw new Error("Biblioteca jsPDF nao carregada.");
    }

    const jsPDFCtor = global.jspdf.jsPDF;
    const testDoc = new jsPDFCtor({
      orientation: "landscape",
      unit: "mm",
      format: "a4",
    });

    if (typeof testDoc.autoTable !== "function") {
      throw new Error("Plugin jsPDF-AutoTable nao carregado.");
    }

    return jsPDFCtor;
  }

  function applyText(doc, fontSize, color, style) {
    doc.setFont("helvetica", style || "normal");
    doc.setFontSize(fontSize);
    doc.setTextColor(...color);
  }

  function drawRoundedBox(doc, x, y, w, h, fillColor, drawColor) {
    doc.setFillColor(...fillColor);
    doc.setDrawColor(...drawColor);
    doc.roundedRect(x, y, w, h, 3, 3, "FD");
  }

  function drawLine(doc, x1, y1, x2, y2, color) {
    doc.setDrawColor(...color);
    doc.setLineWidth(0.2);
    doc.line(x1, y1, x2, y2);
  }

  function addHeader(doc, title, subtitle, options = {}) {
    const generatedAt = pdfContext.generatedAt || new Date().toISOString();
    doc.setFillColor(...THEME.green);
    doc.rect(0, 0, PAGE.width, 6, "F");

    applyText(doc, 7.5, THEME.muted, "normal");
    doc.text(
      safeText(options.eyebrow || "Central de Controle PCM Eletrovia"),
      PAGE.marginX,
      11,
    );

    applyText(doc, 12, THEME.dark, "bold");
    doc.text(safeText(title), PAGE.marginX, 17);

    applyText(doc, 7.5, THEME.gray, "normal");
    doc.text(safeText(subtitle, ""), PAGE.marginX, 21);
    doc.text(
      `Gerado em ${formatDateBR(generatedAt)}`,
      PAGE.width - PAGE.marginX,
      11,
      { align: "right" },
    );

    drawLine(doc, PAGE.marginX, 23, PAGE.width - PAGE.marginX, 23, THEME.border);
  }

  function addFooter(doc, diagnostico, config, pageNumber, totalPages) {
    const generatedAt =
      config?.criadoEm || diagnostico?.meta?.criadoEm || pdfContext.generatedAt;

    drawLine(doc, PAGE.marginX, 199, PAGE.width - PAGE.marginX, 199, THEME.border);
    applyText(doc, 7.5, THEME.gray, "normal");
    doc.text("Central de Controle PCM Eletrovia", PAGE.marginX, 204);
    doc.text("Diagnostico da Carteira", 102, 204);
    doc.text(`Gerado em ${formatDateBR(generatedAt)}`, 196, 204);
    doc.text(`Pagina ${pageNumber} / ${totalPages}`, PAGE.width - PAGE.marginX, 204, {
      align: "right",
    });
  }

  function addPage(doc, title, subtitle, options = {}) {
    doc.addPage();
    addHeader(doc, title, subtitle, options);
    return PAGE.contentTop;
  }

  function checkPageBreak(doc, y, requiredHeight) {
    return y + requiredHeight > PAGE.contentBottom;
  }

  function ensureSpace(doc, currentY, requiredHeight, title, subtitle, options = {}) {
    const pageHeight = doc.internal.pageSize.getHeight();
    const footerSafeY = pageHeight - 18;
    if (currentY + requiredHeight > footerSafeY) {
      return addPage(doc, title, subtitle, options);
    }
    return currentY;
  }

  function drawSectionTitle(doc, title, x, y, note) {
    applyText(doc, 10.5, THEME.dark, "bold");
    doc.text(safeText(title), x, y);
    if (note) {
      applyText(doc, 7.5, THEME.gray, "normal");
      doc.text(safeText(note), PAGE.width - PAGE.marginX, y, { align: "right" });
    }
  }

  function drawExecutiveCard(doc, x, y, w, h, payload = {}) {
    const paletteByVariant = {
      default: { fill: THEME.white, accent: THEME.green, text: THEME.dark },
      success: { fill: THEME.greenSoft, accent: THEME.green, text: THEME.dark },
      warning: { fill: THEME.orangeSoft, accent: THEME.orange, text: THEME.dark },
      danger: { fill: THEME.redSoft, accent: THEME.red, text: THEME.dark },
      info: { fill: THEME.blueSoft, accent: THEME.blue, text: THEME.dark },
      dark: { fill: THEME.darkSoft, accent: THEME.dark, text: THEME.dark },
    };

    const variant = paletteByVariant[payload.variant] || paletteByVariant.default;
    drawRoundedBox(doc, x, y, w, h, variant.fill, THEME.border);

    doc.setFillColor(...variant.accent);
    doc.roundedRect(x, y, 3.2, h, 2, 2, "F");

    applyText(doc, 7.5, THEME.muted, "bold");
    doc.text(safeText(payload.label), x + 6, y + 6);

    applyText(doc, 16, variant.text, "bold");
    doc.text(safeText(payload.value), x + 6, y + 15);

    if (payload.note) {
      applyText(doc, 7.2, THEME.gray, "normal");
      const noteLines = splitText(doc, payload.note, w - 10).slice(0, 3);
      doc.text(noteLines, x + 6, y + 20);
    }
  }

  function drawCard(doc, x, y, w, h, label, value, note, accent) {
    const variant =
      accent === THEME.red
        ? "danger"
        : accent === THEME.orange
          ? "warning"
          : accent === THEME.blue
            ? "info"
            : accent === THEME.dark
              ? "dark"
              : accent === THEME.green
                ? "success"
                : "default";
    drawExecutiveCard(doc, x, y, w, h, { label, value, note, variant });
  }

  function addTechnicalReadingBox(doc, x, y, w, title, paragraphs, options = {}) {
    const items = Array.isArray(paragraphs)
      ? paragraphs.filter(Boolean)
      : [safeText(paragraphs, "")].filter(Boolean);
    const headerHeight = 7;
    const lineHeight = options.lineHeight || 4.3;
    const gap = 1.6;
    const bodyWidth = w - 8;

    let textHeight = 0;
    items.forEach((paragraph) => {
      const lines = splitText(doc, paragraph, bodyWidth);
      textHeight += lines.length * lineHeight + gap;
    });

    const boxHeight = Math.max(options.minHeight || 18, headerHeight + textHeight + 7);
    let boxY = y;

    if (options.pageTitle && options.pageSubtitle) {
      boxY = ensureSpace(
        doc,
        y,
        boxHeight + (options.bottomSpacing || 0),
        options.pageTitle,
        options.pageSubtitle,
        options.pageOptions || {},
      );
    }

    drawRoundedBox(doc, x, boxY, w, boxHeight, options.fillColor || THEME.lightGray, THEME.border);

    applyText(doc, 8.2, THEME.greenDark, "bold");
    doc.text(safeText(title), x + 4, boxY + 5.5);

    let cursor = boxY + 10;
    applyText(doc, 8.2, THEME.text, "normal");
    items.forEach((paragraph) => {
      const lines = splitText(doc, paragraph, bodyWidth);
      doc.text(lines, x + 4, cursor);
      cursor += lines.length * lineHeight + gap;
    });

    return boxY + boxHeight;
  }

  function tableCommonStyles(themeType = "green") {
    const headColor =
      themeType === "danger"
        ? THEME.red
        : themeType === "warning"
          ? THEME.orange
          : themeType === "info"
            ? THEME.blue
            : THEME.green;

    return {
      theme: "grid",
      styles: {
        fontSize: 7.4,
        cellPadding: 2,
        overflow: "linebreak",
        lineColor: THEME.border,
        lineWidth: 0.1,
        textColor: THEME.text,
        valign: "middle",
      },
      headStyles: {
        fillColor: headColor,
        textColor: THEME.white,
        fontStyle: "bold",
      },
      alternateRowStyles: {
        fillColor: THEME.lightGray,
      },
      margin: { left: PAGE.marginX, right: PAGE.marginX },
    };
  }

  function addAutoTable(doc, options = {}) {
    const rows = Array.isArray(options.body) ? options.body.slice() : [];
    const body = rows.length
      ? (options.maxRows ? rows.slice(0, options.maxRows) : rows)
      : [[options.emptyMessage || "Sem dados para exibir"]];
    const head = rows.length ? options.head || [] : [["Informacao"]];

    doc.autoTable({
      ...tableCommonStyles(options.themeType),
      startY: options.startY,
      head,
      body,
      columnStyles: options.columnStyles || {},
    });

    return doc.lastAutoTable.finalY;
  }

  function mapRowsForChart(rows, labelField, valueField, maxItems = 10) {
    return (rows || [])
      .slice(0, maxItems)
      .map((item) => ({
        label:
          item?.[labelField] ||
          item?.nome ||
          item?.status ||
          item?.substatus ||
          item?.centroTrabalho ||
          item?.categoria ||
          "-",
        value: Number(item?.[valueField] || item?.quantidade || 0),
      }))
      .filter((item) => item.value > 0);
  }

  function getCharts() {
    if (!global.CCEDiagnosticoCharts) return null;
    return global.CCEDiagnosticoCharts;
  }

  function shouldIncludeCharts() {
    return pdfContext?.config?.opcoes?.incluirGraficos !== false;
  }

  function isSectionEnabled(config, key) {
    const sections = config?.secoes || {};
    const aliases = SECTION_ALIASES[key] || [key];

    for (const alias of aliases) {
      if (Object.prototype.hasOwnProperty.call(sections, alias)) {
        return sections[alias] === true;
      }
    }

    return true;
  }

  function currentPageInfo(doc) {
    return doc.getCurrentPageInfo().pageNumber;
  }

  function periodLabel(diagnostico) {
    const dates = [];
    (diagnostico?.planejamento?.porDataPlanejada || []).forEach((item) => {
      if (item?.data && item.data !== "-") dates.push(item.data);
    });
    (diagnostico?.vencimentos?.porData || []).forEach((item) => {
      if (item?.data && item.data !== "-") dates.push(item.data);
    });

    const normalized = dates
      .map((item) => safeText(item))
      .filter((item) => /^\d{4}-\d{2}-\d{2}$/.test(item))
      .sort();

    if (!normalized.length) return "Periodo nao consolidado";
    return `${formatDateBR(normalized[0])} a ${formatDateBR(
      normalized[normalized.length - 1],
    )}`;
  }

  function formatFiltersResumo(filters) {
    if (!filters || typeof filters !== "object") return "Recorte filtrado atual da carteira";

    const entries = Object.entries(filters)
      .map(([key, value]) => {
        if (Array.isArray(value)) {
          const visible = value.filter(Boolean);
          if (!visible.length) return null;
          return `${key}: ${visible.slice(0, 3).join(", ")}${visible.length > 3 ? "..." : ""}`;
        }
        if (value && typeof value === "object") return null;
        const text = safeText(value, "");
        if (!text) return null;
        return `${key}: ${text}`;
      })
      .filter(Boolean);

    return entries.length
      ? entries.slice(0, 4).join(" | ")
      : "Recorte filtrado atual da carteira";
  }

  function buildExecutiveAlerts(diagnostico) {
    const alerts = [];
    const resumo = diagnostico?.resumo || {};
    const riscos = diagnostico?.riscos || {};
    const planejadores = diagnostico?.planejadores || {};

    if (Number(resumo.vencidas) > 0) {
      alerts.push({
        nivel: "Alta",
        texto: `Existem ${formatNumber(resumo.vencidas)} demandas vencidas no recorte analisado.`,
      });
    }
    if (Number(resumo.criticas) > 0) {
      alerts.push({
        nivel: "Alta",
        texto: `Ha ${formatNumber(resumo.criticas)} demandas criticas exigindo acompanhamento prioritario.`,
      });
    }
    if (Number(riscos.scoreRiscoGeral) >= 76 || safeText(riscos.nivelRisco).toLowerCase() === "critico") {
      alerts.push({
        nivel: "Alta",
        texto: `O score de risco consolidado esta em ${formatNumber(riscos.scoreRiscoGeral)}, com nivel ${safeText(riscos.nivelRisco)}.`,
      });
    }
    if (Number(planejadores.semPlanejador) > 0) {
      alerts.push({
        nivel: "Media",
        texto: `Ha ${formatNumber(planejadores.semPlanejador)} demandas sem planejador de curto definido.`,
      });
    }

    if (!alerts.length) {
      alerts.push({
        nivel: "Baixa",
        texto: "Nao foram identificados alertas executivos acima do patamar de atencao para o recorte atual.",
      });
    }

    return alerts.slice(0, 5);
  }

  function riskVariant(level) {
    const normalized = safeText(level).toLowerCase();
    if (normalized === "alta" || normalized === "alto" || normalized === "critico") {
      return "danger";
    }
    if (normalized === "media" || normalized === "medio") {
      return "warning";
    }
    if (normalized === "baixa" || normalized === "baixo") {
      return "info";
    }
    return "default";
  }

  function addRecommendationBlocks(doc, recomendacoes, startY, pageTitle, pageSubtitle) {
    let cursor = startY;

    (recomendacoes || []).forEach((item, index) => {
      const description = safeText(item?.descricao, "");
      const action = safeText(item?.acaoSugerida, "");
      const textLines = [
        ...splitText(doc, description, 255),
        ...splitText(doc, `Acao sugerida: ${action}`, 255),
      ];
      const blockHeight = Math.max(20, 13 + textLines.length * 3.7);
      cursor = ensureSpace(doc, cursor, blockHeight + 4, pageTitle, pageSubtitle);

      const variant = riskVariant(item?.prioridade);
      const fill =
        variant === "danger"
          ? THEME.redSoft
          : variant === "warning"
            ? THEME.orangeSoft
            : variant === "info"
              ? THEME.blueSoft
              : THEME.greenSoft;
      const accent =
        variant === "danger"
          ? THEME.red
          : variant === "warning"
            ? THEME.orange
            : variant === "info"
              ? THEME.blue
              : THEME.green;

      drawRoundedBox(doc, PAGE.marginX, cursor, 273, blockHeight, fill, THEME.border);
      doc.setFillColor(...accent);
      doc.roundedRect(PAGE.marginX, cursor, 4, blockHeight, 2, 2, "F");

      applyText(doc, 8.1, THEME.muted, "bold");
      doc.text(
        `${index + 1}. Prioridade ${safeText(item?.prioridade, "Media")}`,
        PAGE.marginX + 7,
        cursor + 5,
      );

      applyText(doc, 10, THEME.dark, "bold");
      doc.text(safeText(item?.titulo), PAGE.marginX + 7, cursor + 11);

      applyText(doc, 8, THEME.text, "normal");
      doc.text(textLines, PAGE.marginX + 7, cursor + 16);
      cursor += blockHeight + 4;
    });

    return cursor;
  }

  function drawBulletList(doc, items, x, startY, maxWidth) {
    let cursor = startY;
    applyText(doc, 8.2, THEME.text, "normal");
    (items || []).forEach((item) => {
      const lines = splitText(doc, safeText(item), maxWidth - 6);
      doc.text("•", x, cursor);
      doc.text(lines, x + 4, cursor);
      cursor += lines.length * 4.2 + 1.2;
    });
    return cursor;
  }

  function addCoverPage(doc, diagnostico, config) {
    doc.setFillColor(...THEME.green);
    doc.rect(0, 0, PAGE.width, 30, "F");
    doc.setFillColor(...THEME.greenDark);
    doc.rect(0, PAGE.height - 28, PAGE.width, 28, "F");

    applyText(doc, 9, THEME.white, "bold");
    doc.text("CENTRAL DE CONTROLE PCM ELETROVIA", PAGE.marginX, 12);

    applyText(doc, 23, THEME.white, "bold");
    doc.text(safeText(diagnostico?.meta?.titulo, "Diagnostico da Carteira"), PAGE.marginX, 23);

    applyText(doc, 9.2, THEME.white, "normal");
    doc.text(
      "Relatorio tecnico executivo para planejamento ferroviario, acompanhamento operacional e governanca da carteira filtrada.",
      PAGE.marginX,
      28,
    );

    drawRoundedBox(doc, PAGE.marginX, 42, 176, 72, THEME.white, THEME.border);
    applyText(doc, 8, THEME.muted, "bold");
    doc.text("INFORMACOES DO RELATORIO", PAGE.marginX + 5, 49);

    const metaRows = [
      ["Tipo do relatorio", reportTypeLabel(config?.tipoRelatorio || diagnostico?.meta?.tipoRelatorio)],
      ["Recorte filtrado atual", formatFiltersResumo(diagnostico?.meta?.filtrosResumo)],
      ["Responsavel", safeText(config?.responsavel || diagnostico?.meta?.responsavel)],
      ["Data/hora de geracao", formatDateBR(config?.criadoEm || diagnostico?.meta?.criadoEm)],
      ["Total de demandas", formatNumber(diagnostico?.resumo?.total)],
      ["Periodo analisado", periodLabel(diagnostico)],
    ];

    let metaY = 58;
    metaRows.forEach(([label, value]) => {
      applyText(doc, 7.6, THEME.muted, "bold");
      doc.text(`${label}:`, PAGE.marginX + 5, metaY);
      applyText(doc, 8.4, THEME.text, "normal");
      const lines = splitText(doc, value, 115);
      doc.text(lines, PAGE.marginX + 46, metaY);
      metaY += Math.max(6, lines.length * 4.2);
    });

    if (diagnostico?.meta?.observacaoTecnica) {
      addTechnicalReadingBox(
        doc,
        194,
        42,
        91,
        "Observacao tecnica",
        [diagnostico.meta.observacaoTecnica],
        { fillColor: THEME.greenSoft, minHeight: 72 },
      );
    } else {
      addTechnicalReadingBox(
        doc,
        194,
        42,
        91,
        "Escopo executivo",
        [
          "Leitura consolidada da carteira filtrada com foco em vencimentos, planejamento de curto prazo, rastreabilidade por KM e exposicao operacional.",
          "Documento preparado para reunioes de M+1, supervisao, gerencia ou alinhamento tatico de campo.",
        ],
        { fillColor: THEME.greenSoft, minHeight: 72 },
      );
    }

    const cards = [
      { label: "Total", value: formatNumber(diagnostico?.resumo?.total), note: "Demandas no recorte", variant: "dark" },
      { label: "Planejadas", value: formatNumber(diagnostico?.resumo?.planejadas), note: "Programacao ativa", variant: "info" },
      { label: "Realizadas", value: formatNumber(diagnostico?.resumo?.realizadas), note: "Execucao concluida", variant: "success" },
      { label: "Vencidas", value: formatNumber(diagnostico?.resumo?.vencidas), note: "Tratamento prioritario", variant: "danger" },
      { label: "Criticas", value: formatNumber(diagnostico?.resumo?.criticas), note: "Prioridade operacional", variant: "warning" },
      { label: "Com KM", value: formatNumber(diagnostico?.resumo?.comKm), note: formatPercent(diagnostico?.resumo?.percentualComKm), variant: "default" },
    ];

    const cardY = 126;
    const cardW = 42.5;
    cards.forEach((card, index) => {
      drawExecutiveCard(
        doc,
        PAGE.marginX + index * (cardW + 3),
        cardY,
        cardW,
        28,
        card,
      );
    });

    applyText(doc, 8, THEME.white, "normal");
    doc.text(
      "Diagnostico gerado automaticamente a partir do recorte filtrado atual da carteira.",
      PAGE.marginX,
      193,
    );
  }

  function addExecutiveSummaryIndex(doc, diagnostico, config, reuseCurrentPage = false) {
    let y;
    if (reuseCurrentPage) {
      addHeader(doc, "Sumario Executivo", "Estrutura do relatorio e escopo do diagnostico");
      y = PAGE.contentTop;
    } else {
      y = addPage(
        doc,
        "Sumario Executivo",
        "Estrutura do relatorio e escopo do diagnostico",
      );
    }

    const description = [
      "Este relatorio consolida a leitura tecnica da carteira filtrada no sistema, com foco em planejamento de curto prazo, vencimentos, tolerancias, rastreabilidade por KM, concentracao operacional e riscos associados a execucao ferroviaria.",
      "O conteudo respeita integralmente o recorte filtrado da carteira no momento da geracao.",
    ];

    y = addTechnicalReadingBox(
      doc,
      PAGE.marginX,
      y,
      273,
      "Objetivo do relatorio",
      description,
      { fillColor: THEME.greenSoft, minHeight: 30 },
    ) + 8;

    drawSectionTitle(doc, "Secoes do documento", PAGE.marginX, y);
    y += 8;

    const sections = [
      { key: "resumo", label: "Resumo executivo" },
      { key: "vencimentos", label: "Status e vencimentos" },
      { key: "planejamento", label: "Planejamento e centros de trabalho" },
      { key: "planejadores", label: "Planejadores, supervisoes e gerencias" },
      { key: "km", label: "KM, tolerancias e riscos" },
      { key: "clima", label: "Diagnostico climatico" },
      { key: "listaCritica", label: "Lista critica" },
      { key: "recomendacoes", label: "Recomendacoes tecnicas" },
    ].filter((item) => isSectionEnabled(config, item.key));

    y = drawBulletList(
      doc,
      sections.map((item, index) => `${index + 1}. ${item.label}`),
      PAGE.marginX,
      y,
      130,
    );

    addTechnicalReadingBox(
      doc,
      164,
      84,
      121,
      "Escopo do diagnostico",
      [
        `Tipo do relatorio: ${reportTypeLabel(config?.tipoRelatorio || diagnostico?.meta?.tipoRelatorio)}`,
        `Periodo analisado: ${periodLabel(diagnostico)}`,
        `Responsavel: ${safeText(config?.responsavel || diagnostico?.meta?.responsavel)}`,
        `Recorte atual: ${formatFiltersResumo(diagnostico?.meta?.filtrosResumo)}`,
      ],
      { fillColor: THEME.lightBlue, minHeight: 42 },
    );
  }

  function addResumoPage(doc, diagnostico, Charts, config) {
    let y = addPage(doc, "Resumo Executivo", "Indicadores executivos da carteira filtrada");
    const pageHeight = doc.internal.pageSize.getHeight();
    const footerSafeY = pageHeight - 18;

    if (isSectionEnabled(config, "kpis")) {
      const cardData = [
        { label: "Total", value: formatNumber(diagnostico.resumo.total), note: "Volume filtrado", variant: "dark" },
        { label: "A Planejar", value: formatNumber(diagnostico.resumo.aPlanejar), note: "Sem programacao", variant: "default" },
        { label: "Planejadas", value: formatNumber(diagnostico.resumo.planejadas), note: "Programacao ativa", variant: "info" },
        { label: "Replanejadas", value: formatNumber(diagnostico.resumo.replanejadas), note: "Revisao de agenda", variant: "warning" },
        { label: "Realizadas", value: formatNumber(diagnostico.resumo.realizadas), note: formatPercent(diagnostico.resumo.percentualRealizado), variant: "success" },
        { label: "Vencidas", value: formatNumber(diagnostico.resumo.vencidas), note: "Atraso acumulado", variant: "danger" },
        { label: "Criticas", value: formatNumber(diagnostico.resumo.criticas), note: formatPercent(diagnostico.resumo.percentualCritico), variant: "warning" },
        { label: "Com KM", value: formatNumber(diagnostico.resumo.comKm), note: formatPercent(diagnostico.resumo.percentualComKm), variant: "default" },
      ];

      const columns = 4;
      const cardW = 64.5;
      const cardH = 24;
      cardData.forEach((card, index) => {
        const col = index % columns;
        const row = Math.floor(index / columns);
        drawExecutiveCard(
          doc,
          PAGE.marginX + col * (cardW + 3),
          y + row * (cardH + 4),
          cardW,
          cardH,
          card,
        );
      });
      y += 56;
    }

    const alerts = buildExecutiveAlerts(diagnostico);
    const narrativeX = PAGE.marginX;
    const narrativeW = 156;
    const alertX = 175;
    const alertW = 110;
    const narrativeTop = y;
    const riskCardBottom = (() => {
      drawExecutiveCard(doc, alertX, narrativeTop, alertW, 22, {
        label: "Score de risco",
        value: formatNumber(diagnostico.riscos.scoreRiscoGeral),
        note: `Nivel ${safeText(diagnostico.riscos.nivelRisco)}`,
        variant: riskVariant(diagnostico.riscos.nivelRisco),
      });
      return narrativeTop + 22;
    })();

    const narrativeBottom = addTechnicalReadingBox(
      doc,
      narrativeX,
      narrativeTop,
      narrativeW,
      "Narrativa executiva",
      diagnostico.narrativaExecutiva,
      { fillColor: THEME.white, minHeight: 36 },
    );

    const alertBottom = addTechnicalReadingBox(
      doc,
      alertX,
      riskCardBottom + 4,
      alertW,
      "Principais alertas",
      alerts.map((item) => `${item.nivel}: ${item.texto}`),
      { fillColor: THEME.orangeSoft, minHeight: 28 },
    );

    let chartY = Math.max(narrativeBottom, alertBottom) + 8;
    let chartH = Math.min(48, footerSafeY - chartY - 4);

    if (chartH < 34) {
      chartY = addPage(
        doc,
        "Resumo Executivo",
        "Indicadores executivos da carteira filtrada",
      );
      chartH = Math.min(48, footerSafeY - chartY - 4);
    }

    if (shouldIncludeCharts() && Charts) {
      Charts.drawDonutChart(doc, {
        x: PAGE.marginX,
        y: chartY,
        size: 28,
        title: "Distribuicao por status",
        data: mapRowsForChart(diagnostico.status.porStatus, "status", "quantidade", 6).map(
          (item, index) => ({
            ...item,
            color: [THEME.green, THEME.blue, THEME.orange, THEME.red, [109, 40, 217], THEME.gray][index],
          }),
        ),
      });
      Charts.drawRiskMeter(doc, {
        x: 102,
        y: chartY,
        w: 177,
        h: Math.max(22, chartH - 18),
        title: "Leitura consolidada de risco operacional",
        score: diagnostico.riscos.scoreRiscoGeral,
        nivel: diagnostico.riscos.nivelRisco,
      });
    }
  }

  function addStatusPage(doc, diagnostico, Charts) {
    let y = addPage(doc, "Status e Vencimentos", "Distribuicao operacional e agenda de vencimentos");

    if (shouldIncludeCharts() && Charts) {
      Charts.drawMiniTimeline(doc, {
        x: PAGE.marginX,
        y,
        w: 273,
        h: 38,
        title: "Vencimentos por data",
        data: mapRowsForChart(diagnostico.vencimentos.porData, "data", "quantidade", 12),
      });
      y += 46;
    }

    y = addTechnicalReadingBox(
      doc,
      PAGE.marginX,
      y,
      273,
      "Leitura tecnica",
      [
        ...(diagnostico.status.leitura || []),
        ...(diagnostico.vencimentos.leitura || []),
      ],
      { fillColor: THEME.white, minHeight: 20 },
    ) + 5;

    drawSectionTitle(doc, "Distribuicao por status", PAGE.marginX, y);
    y += 3;
    y = addAutoTable(doc, {
      startY: y + 2,
      head: [["Status", "Quantidade", "%"]],
      body: diagnostico.status.porStatus.map((item) => [
        safeText(item.status),
        formatNumber(item.quantidade),
        formatPercent(item.percentual),
      ]),
      themeType: "green",
    }) + 5;

    y = ensureSpace(doc, y, 34, "Status e Vencimentos", "Distribuicao operacional e agenda de vencimentos");
    drawSectionTitle(doc, "Demandas vencidas prioritarias", PAGE.marginX, y);
    addAutoTable(doc, {
      startY: y + 2,
      head: [["ID", "OM", "Descricao", "Vencimento", "Status", "Critico"]],
      body: diagnostico.vencimentos.vencidas.slice(0, 12).map((item) => [
        safeText(item.id),
        safeText(item.ordem),
        truncateTextValue(item.descricao, 66),
        formatDateBR(item.vencimento),
        safeText(item.status),
        safeText(item.critico),
      ]),
      themeType: "danger",
      emptyMessage: "Sem demandas vencidas no recorte filtrado.",
    });
  }

  function addPlanejamentoCentrosPage(doc, diagnostico, Charts, config) {
    let y = addPage(doc, "Planejamento e Centros", "Carga planejada e concentracao operacional");

    if (shouldIncludeCharts() && Charts) {
      if (isSectionEnabled(config, "planejamento")) {
        Charts.drawVerticalBarChart(doc, {
          x: PAGE.marginX,
          y,
          w: 128,
          h: 42,
          title: "Planejamento por data",
          data: mapRowsForChart(diagnostico.planejamento.porDataPlanejada, "data", "quantidade", 12),
        });
      }

      if (isSectionEnabled(config, "centros")) {
        Charts.drawHorizontalBarChart(doc, {
          x: 149,
          y,
          w: 136,
          h: 42,
          title: "Ranking de centros de trabalho",
          data: mapRowsForChart(
            diagnostico.centros.ranking.map((item) => ({
              label: item.centroTrabalho,
              value: item.quantidade,
            })),
            "label",
            "value",
            8,
          ),
          showPercent: true,
          total: diagnostico.resumo.total,
        });
      }
      y += 48;
    }

    y = addTechnicalReadingBox(
      doc,
      PAGE.marginX,
      y,
      273,
      "Leitura tecnica",
      [
        ...(isSectionEnabled(config, "planejamento") ? diagnostico.planejamento.leitura || [] : []),
        ...(isSectionEnabled(config, "centros") ? diagnostico.centros.leitura || [] : []),
      ],
      { fillColor: THEME.white, minHeight: 20 },
    ) + 5;

    if (isSectionEnabled(config, "planejamento")) {
      drawSectionTitle(doc, "Planejamento por data", PAGE.marginX, y);
      y = addAutoTable(doc, {
        startY: y + 2,
        head: [["Data planejada", "Quantidade"]],
        body: diagnostico.planejamento.porDataPlanejada.slice(0, 15).map((item) => [
          formatDateBR(item.data),
          formatNumber(item.quantidade),
        ]),
        themeType: "info",
        emptyMessage: "Sem datas planejadas no recorte analisado.",
      }) + 5;
    }

    if (isSectionEnabled(config, "centros")) {
      y = ensureSpace(doc, y, 42, "Planejamento e Centros", "Carga planejada e concentracao operacional");
      drawSectionTitle(doc, "Centros de trabalho", PAGE.marginX, y);
      addAutoTable(doc, {
        startY: y + 2,
        head: [[
          "Centro",
          "Qtd",
          "%",
          "Planejadas",
          "A Planejar",
          "Replanejadas",
          "Realizadas",
          "Vencidas",
          "Criticas",
        ]],
        body: diagnostico.centros.ranking.slice(0, 15).map((item) => [
          safeText(item.centroTrabalho),
          formatNumber(item.quantidade),
          formatPercent(item.percentual),
          formatNumber(item.planejadas),
          formatNumber(item.aPlanejar),
          formatNumber(item.replanejadas),
          formatNumber(item.realizadas),
          formatNumber(item.vencidas),
          formatNumber(item.criticas),
        ]),
        themeType: "green",
      });
    }
  }

  function addResponsabilidadesPage(doc, diagnostico, Charts) {
    let y = addPage(
      doc,
      "Planejadores, Supervisoes e Gerencias",
      "Distribuicao de responsabilidade operacional",
    );

    if (shouldIncludeCharts() && Charts) {
      Charts.drawHorizontalBarChart(doc, {
        x: PAGE.marginX,
        y,
        w: 132,
        h: 42,
        title: "Demandas por planejador de curto",
        data: mapRowsForChart(
          diagnostico.planejadores.ranking.map((item) => ({
            label: item.planejadorCurto,
            value: item.quantidade,
          })),
          "label",
          "value",
          8,
        ),
      });
      Charts.drawHorizontalBarChart(doc, {
        x: 149,
        y,
        w: 136,
        h: 42,
        title: "Demandas por supervisao",
        data: mapRowsForChart(
          diagnostico.supervisoes.ranking.map((item) => ({
            label: item.nome,
            value: item.quantidade,
          })),
          "label",
          "value",
          8,
        ),
        color: THEME.blue,
      });
      y += 48;
    }

    y = addTechnicalReadingBox(
      doc,
      PAGE.marginX,
      y,
      273,
      "Leitura tecnica",
      [
        ...(diagnostico.planejadores.leitura || []),
        ...(diagnostico.supervisoes.leitura || []),
        ...(diagnostico.gerencias.leitura || []),
      ],
      { fillColor: THEME.white, minHeight: 20 },
    ) + 5;

    drawSectionTitle(doc, "Planejadores de curto", PAGE.marginX, y);
    y = addAutoTable(doc, {
      startY: y + 2,
      head: [["Planejador", "Qtd", "%", "Planejadas", "Vencidas", "Criticas"]],
      body: diagnostico.planejadores.ranking.slice(0, 15).map((item) => [
        safeText(item.planejadorCurto),
        formatNumber(item.quantidade),
        formatPercent(item.percentual),
        formatNumber(item.planejadas),
        formatNumber(item.vencidas),
        formatNumber(item.criticas),
      ]),
      themeType: "info",
    }) + 5;

    y = ensureSpace(
      doc,
      y,
      56,
      "Planejadores, Supervisoes e Gerencias",
      "Distribuicao de responsabilidade operacional",
    );
    drawSectionTitle(doc, "Supervisoes e gerencias", PAGE.marginX, y);
    y = addAutoTable(doc, {
      startY: y + 2,
      head: [["Supervisao", "Qtd", "%", "Planejadas", "Vencidas", "Criticas"]],
      body: diagnostico.supervisoes.ranking.slice(0, 12).map((item) => [
        safeText(item.nome),
        formatNumber(item.quantidade),
        formatPercent(item.percentual),
        formatNumber(item.planejadas),
        formatNumber(item.vencidas),
        formatNumber(item.criticas),
      ]),
      themeType: "green",
    }) + 5;

    addAutoTable(doc, {
      startY: y,
      head: [["Gerencia", "Qtd", "%", "Planejadas", "Vencidas", "Criticas"]],
      body: diagnostico.gerencias.ranking.slice(0, 12).map((item) => [
        safeText(item.nome),
        formatNumber(item.quantidade),
        formatPercent(item.percentual),
        formatNumber(item.planejadas),
        formatNumber(item.vencidas),
        formatNumber(item.criticas),
      ]),
      themeType: "green",
    });
  }

  function addKmRiscosPage(doc, diagnostico, Charts, config) {
    let y = addPage(doc, "KM, Tolerancias e Riscos", "Rastreabilidade ferroviaria e exposicao operacional");
    const kms = diagnostico.kms || {};
    const tolerancias = diagnostico.toleranciasLineares || diagnostico.tolerancias || {};
    const riscos = diagnostico.riscosLineares || diagnostico.riscos || {};
    const kmTrechos = (kms.rankingTrechos || [])
      .map((item) => ({
        ...item,
        centros: (item.centros || []).filter((centro) => !isPatioText(centro)),
        locais: (item.locais || []).filter((local) => !isPatioText(local)),
      }))
      .filter(
        (item) =>
          !isPatioText(item.trecho) &&
          !isPatioText(item.kmInicio) &&
          !isPatioText(item.kmFim) &&
          !(item.centros || []).some(isPatioText) &&
          !(item.locais || []).some(isPatioText),
      );

    const cards = [];
    if (isSectionEnabled(config, "km")) {
      cards.push(
        { label: "Com KM", value: formatNumber(kms.comKm), note: formatPercent(kms.percentualComKm), variant: "default" },
        { label: "Sem KM", value: formatNumber(kms.semKm), note: "Sem KM informado", variant: "warning" },
      );
    }
    if (isSectionEnabled(config, "tolerancias")) {
      cards.push(
        { label: "Com tolerancia", value: formatNumber(tolerancias.comTolerancia), note: "Base parametrizada", variant: "success" },
        { label: "Sem tolerancia", value: formatNumber(tolerancias.semTolerancia), note: "Analise de janela", variant: "warning" },
        { label: "Fora da janela", value: formatNumber(tolerancias.foraJanela), note: "Planejamento fora do limite", variant: "danger" },
      );
    }
    cards.push({
      label: "Risco geral",
      value: formatNumber(riscos.scoreRiscoGeral),
      note: `Nivel ${safeText(riscos.nivelRisco)}`,
      variant: riskVariant(riscos.nivelRisco),
    });

    const cardW = 42.5;
    cards.slice(0, 6).forEach((card, index) => {
      drawExecutiveCard(doc, PAGE.marginX + index * (cardW + 3), y, cardW, 24, card);
    });
    y += 30;

    if (shouldIncludeCharts() && Charts) {
      if (isSectionEnabled(config, "km")) {
        Charts.drawHorizontalBarChart(doc, {
          x: PAGE.marginX,
          y,
          w: 132,
          h: 42,
          title: "Rastreabilidade por trecho",
          data: mapRowsForChart(
            kmTrechos.map((item) => ({
              label: item.trecho,
              value: item.quantidade,
            })),
            "label",
            "value",
            8,
          ),
        });
      }
      Charts.drawHorizontalBarChart(doc, {
        x: 149,
        y,
        w: 136,
        h: 42,
        title: "Fatores de risco",
        data: mapRowsForChart(
          (riscos.fatores || []).map((item) => ({
            label: item.fator,
            value: item.quantidade,
          })),
          "label",
          "value",
          8,
        ),
        color: THEME.red,
      });
      y += 48;
    }

    y = addTechnicalReadingBox(
      doc,
      PAGE.marginX,
      y,
      273,
      "Leitura tecnica",
      [
        ...(isSectionEnabled(config, "km") ? kms.leitura || [] : []),
        ...(isSectionEnabled(config, "tolerancias") ? tolerancias.leitura || [] : []),
        ...(riscos.leitura || []),
      ],
      { fillColor: THEME.white, minHeight: 20 },
    ) + 5;

    if (isSectionEnabled(config, "km")) {
      drawSectionTitle(doc, "Trechos com maior concentracao", PAGE.marginX, y);
      y = addAutoTable(doc, {
        startY: y + 2,
        head: [["Trecho", "KM Inicio", "KM Fim", "Qtd", "Centros", "Locais"]],
        body: kmTrechos.slice(0, 12).map((item) => [
          truncateTextValue(item.trecho, 30),
          safeText(item.kmInicio),
          safeText(item.kmFim),
          formatNumber(item.quantidade),
          truncateTextValue((item.centros || []).slice(0, 3).join(", "), 28),
          truncateTextValue((item.locais || []).slice(0, 2).join(", "), 32),
        ]),
        themeType: "info",
        emptyMessage: "Sem trechos com KM informado no recorte.",
      }) + 5;
    }

    if (isSectionEnabled(config, "tolerancias")) {
      y = ensureSpace(doc, y, 42, "KM, Tolerancias e Riscos", "Rastreabilidade ferroviaria e exposicao operacional");
      drawSectionTitle(doc, "Tolerancias e fatores de risco", PAGE.marginX, y);
      addAutoTable(doc, {
        startY: y + 2,
        head: [["Fator de risco", "Quantidade", "Severidade", "Descricao"]],
        body: (riscos.fatores || []).map((item) => [
          safeText(item.fator),
          formatNumber(item.quantidade),
          safeText(item.severidade),
          truncateTextValue(item.descricao, 96),
        ]),
        themeType: "danger",
      });
    }
  }

  function addPatiosRiscosPage(doc, diagnostico, Charts) {
    const patios = diagnostico.patios || {};
    const resumo = patios.resumo || {};
    let y = addPage(
      doc,
      "Patios, Tolerancias e Riscos",
      "Leitura operacional dedicada aos patios ferroviarios",
    );

    const cards = [
      { label: "Demandas em patio", value: formatNumber(resumo.totalPatio), note: "Recorte de patio", variant: "default" },
      { label: "Patios distintos", value: formatNumber(resumo.patiosDistintos), note: "Locais agrupados", variant: "info" },
      { label: "Sem tolerancia", value: formatNumber(resumo.semTolerancia), note: "Sem janela definida", variant: "warning" },
      { label: "Vencidas", value: formatNumber(resumo.vencidas), note: "Exposicao operacional", variant: "danger" },
      { label: "Criticas", value: formatNumber(resumo.criticas), note: "Sensibilidade elevada", variant: "danger" },
    ];

    const cardW = 51;
    cards.forEach((card, index) => {
      drawExecutiveCard(doc, PAGE.marginX + index * (cardW + 4), y, cardW, 24, card);
    });
    y += 30;

    if (shouldIncludeCharts() && Charts) {
      Charts.drawHorizontalBarChart(doc, {
        x: PAGE.marginX,
        y,
        w: 132,
        h: 42,
        title: "Patios com maior concentracao",
        data: mapRowsForChart(
          (patios.rankingPatios || []).map((item) => ({
            label: item.patio,
            value: item.quantidade,
          })),
          "label",
          "value",
          8,
        ),
      });
      Charts.drawHorizontalBarChart(doc, {
        x: 149,
        y,
        w: 136,
        h: 42,
        title: "Fatores de risco em patio",
        data: mapRowsForChart(
          (patios.fatores || []).map((item) => ({
            label: item.fator,
            value: item.quantidade,
          })),
          "label",
          "value",
          8,
        ),
        color: THEME.red,
      });
      y += 48;
    }

    y = addTechnicalReadingBox(
      doc,
      PAGE.marginX,
      y,
      273,
      "Leitura tecnica",
      patios.leitura || ["Nao foram identificadas demandas de patio no recorte filtrado."],
      { fillColor: THEME.white, minHeight: 20 },
    ) + 5;

    drawSectionTitle(doc, "Patios com maior concentracao", PAGE.marginX, y);
    addAutoTable(doc, {
      startY: y + 2,
      head: [["Patio", "Qtd", "Centros", "Gerencias", "Vencidas", "Criticas", "Sem tolerancia", "Fora da janela"]],
      body: (patios.rankingPatios || []).slice(0, 12).map((item) => [
        truncateTextValue(item.patio, 42),
        formatNumber(item.quantidade),
        truncateTextValue((item.centros || []).slice(0, 3).join(", "), 28),
        truncateTextValue((item.gerencias || []).slice(0, 2).join(", "), 24),
        formatNumber(item.vencidas),
        formatNumber(item.criticas),
        formatNumber(item.semTolerancia),
        formatNumber(item.foraJanela),
      ]),
      themeType: "info",
      emptyMessage: "Sem demandas de patio no recorte atual.",
    });
  }

  function addClimaPage(doc, diagnostico, Charts) {
    let y = addPage(doc, "Diagnostico Climatico", "Sensibilidade climatica e risco operacional de campo");

    const clima = diagnostico.clima || {};
    const resumo = clima.resumo || {};
    const segmentos = clima.segmentos || {};
    const malhaLinear = segmentos.malhaLinear || {};
    const areasPatio = segmentos.areasPatio || {};

    const cards = [
      { label: "Demandas analisadas", value: formatNumber(resumo.totalAnalisado), note: "Recorte avaliado", variant: "default" },
      { label: "Sensiveis ao clima", value: formatNumber(resumo.sensiveisAoClima), note: "Exposicao direta", variant: "warning" },
      { label: "Alto risco", value: formatNumber(resumo.altoRiscoClimatico), note: "Validacao de janela", variant: "danger" },
      { label: "Medio risco", value: formatNumber(resumo.medioRiscoClimatico), note: "Monitoramento", variant: "warning" },
      { label: "Sem data", value: formatNumber(resumo.semDataPlanejada), note: "Sem referencia operacional", variant: "default" },
      { label: "Sem local/centro/KM", value: formatNumber(resumo.semLocalOuCentro), note: "Sem referencia", variant: "info" },
    ];

    const cardW = 42.5;
    cards.forEach((card, index) => {
      drawExecutiveCard(doc, PAGE.marginX + index * (cardW + 3), y, cardW, 24, card);
    });
    y += 30;

    const segmentY = y;
    const segmentBottomLeft = addTechnicalReadingBox(
      doc,
      PAGE.marginX,
      segmentY,
      132,
      "Clima na malha linear",
      [
        `${formatNumber(malhaLinear.total)} demandas avaliadas na malha linear.`,
        `${formatNumber(malhaLinear.alto)} em alto risco e ${formatNumber(malhaLinear.medio)} em risco medio.`,
        malhaLinear.rankingCentros?.[0]
          ? `Maior concentracao em ${safeText(malhaLinear.rankingCentros[0].nome)}.`
          : "Sem concentracao linear relevante no recorte.",
        ],
      { fillColor: THEME.white, minHeight: 28 },
    );
    const segmentBottomRight = addTechnicalReadingBox(
      doc,
      149,
      segmentY,
      136,
      "Clima em areas de patio",
      [
        `${formatNumber(areasPatio.total)} demandas avaliadas em areas de patio.`,
        `${formatNumber(areasPatio.alto)} em alto risco e ${formatNumber(areasPatio.medio)} em risco medio.`,
        areasPatio.rankingLocais?.[0]
          ? `Maior concentracao em ${safeText(areasPatio.rankingLocais[0].nome)}.`
          : "Sem concentracao de patio relevante no recorte.",
        ],
      { fillColor: THEME.white, minHeight: 28 },
    );
    y = Math.max(segmentBottomLeft, segmentBottomRight) + 5;

    if (shouldIncludeCharts() && Charts) {
      Charts.drawHorizontalBarChart(doc, {
        x: PAGE.marginX,
        y,
        w: 132,
        h: 42,
        title: "Risco climatico por centro",
        data: mapRowsForChart(clima.porCentro, "centroTrabalho", "alto", 8),
        color: THEME.red,
      });
      Charts.drawHorizontalBarChart(doc, {
        x: 149,
        y,
        w: 136,
        h: 42,
        title: "Risco por tipo de atividade",
        data: mapRowsForChart(clima.porTipoAtividade, "categoria", "alto", 8),
        color: THEME.orange,
      });
      y += 48;

      Charts.drawMiniTimeline(doc, {
        x: PAGE.marginX,
        y,
        w: 273,
        h: 30,
        title: "Demandas sensiveis por data",
        data: mapRowsForChart(clima.porData, "data", "quantidade", 10),
      });
      y += 36;
    }

    y = addTechnicalReadingBox(
      doc,
      PAGE.marginX,
      y,
      273,
      "Leitura tecnica climatica",
      clima.leitura || [
        "Nao havia dados climaticos carregados no momento da geracao do relatorio.",
      ],
      {
        fillColor: THEME.white,
        minHeight: 20,
        bottomSpacing: 18,
        pageTitle: "Diagnostico Climatico",
        pageSubtitle: "Sensibilidade climatica e risco operacional de campo",
      },
    ) + 5;

    y = ensureSpace(
      doc,
      y,
      30,
      "Diagnostico Climatico",
      "Sensibilidade climatica e risco operacional de campo",
    );
    drawSectionTitle(doc, "Recorte climatico por ambiente", PAGE.marginX, y);
    y = addAutoTable(doc, {
      startY: y + 2,
      head: [["Ambiente", "Demandas", "Sensiveis", "Alto risco", "Medio risco", "Baixo risco", "Maior concentracao"]],
      body: [
        [
          "Malha linear",
          formatNumber(malhaLinear.total),
          formatNumber(malhaLinear.sensiveis),
          formatNumber(malhaLinear.alto),
          formatNumber(malhaLinear.medio),
          formatNumber(malhaLinear.baixo),
          safeText(malhaLinear.rankingCentros?.[0]?.nome, "-"),
        ],
        [
          "Areas de patio",
          formatNumber(areasPatio.total),
          formatNumber(areasPatio.sensiveis),
          formatNumber(areasPatio.alto),
          formatNumber(areasPatio.medio),
          formatNumber(areasPatio.baixo),
          safeText(areasPatio.rankingLocais?.[0]?.nome, "-"),
        ],
      ],
      themeType: "info",
    }) + 5;

    drawSectionTitle(doc, "Demandas criticas de clima", PAGE.marginX, y);
    addAutoTable(doc, {
      startY: y + 2,
      head: [[
        "Risco",
        "Score",
        "Data",
        "Origem Data",
        "OM",
        "ID",
        "Centro",
        "Local",
        "KM Inicio",
        "KM Fim",
        "Categoria",
        "Recomendacao",
      ]],
      body: (clima.demandasCriticasClima || []).slice(0, 30).map((item) => [
        safeText(item.riscoClimatico),
        formatNumber(item.score),
        formatDateBR(item.dataOperacional),
        safeText(item.origemData),
        safeText(item.ordem),
        safeText(item.id),
        truncateTextValue(item.centroTrabalho, 18),
        truncateTextValue(item.localInstalacao, 22),
        safeText(item.kmInicio),
        safeText(item.kmFim),
        truncateTextValue(item.categoria, 18),
        truncateTextValue(item.recomendacao, 42),
      ]),
      themeType: "warning",
      emptyMessage: "Sem demandas climaticas criticas para o recorte atual.",
    });
  }

  function formatKmRange(item) {
    const start = safeText(item?.kmInicio, "");
    const end = safeText(item?.kmFim, "");
    if (start && end) return `${start} - ${end}`;
    return start || end || "-";
  }

  function addListaCriticaPage(doc, diagnostico) {
    let y = addPage(doc, "Lista Critica", "Demandas com necessidade de acompanhamento prioritario");

    y = addTechnicalReadingBox(
      doc,
      PAGE.marginX,
      y,
      273,
      "Diretriz de leitura",
      [
        "A lista abaixo prioriza demandas com combinacao de vencimento, criticidade, ausencia de referencia territorial e risco tecnico consolidado.",
      ],
      { fillColor: THEME.white, minHeight: 18 },
    ) + 5;

    addAutoTable(doc, {
      startY: y,
      head: [[
        "Motivo",
        "Status",
        "OM",
        "Centro",
        "KM",
        "Vencimento",
        "Planejada",
        "Critico",
        "Descricao",
      ]],
      body: (diagnostico.listaCritica || []).slice(0, 50).map((item) => [
        truncateTextValue(item.motivo, 20),
        safeText(item.status),
        safeText(item.ordem),
        truncateTextValue(item.centroTrabalho, 18),
        formatKmRange(item),
        formatDateBR(item.vencimento),
        formatDateBR(item.dataPlanejada || item.dataReplanejadaAtual),
        safeText(item.critico),
        truncateTextValue(item.descricao, 42),
      ]),
      themeType: "danger",
      emptyMessage: "Sem itens criticos no recorte atual.",
    });
  }

  function addRecomendacoesPage(doc, diagnostico) {
    let y = addPage(doc, "Recomendacoes Tecnicas", "Direcionadores de acao para a carteira filtrada");

    y = addRecommendationBlocks(
      doc,
      diagnostico.recomendacoes || [],
      y + 2,
      "Recomendacoes Tecnicas",
      "Direcionadores de acao para a carteira filtrada",
    );

    y = ensureSpace(
      doc,
      y + 2,
      24,
      "Recomendacoes Tecnicas",
      "Direcionadores de acao para a carteira filtrada",
    );

    addTechnicalReadingBox(
      doc,
      PAGE.marginX,
      y,
      273,
      "Nota tecnica",
      [
        "Este diagnostico foi gerado automaticamente a partir do recorte filtrado da carteira no momento da emissao. As recomendacoes devem ser avaliadas pelo planejamento responsavel, considerando restricoes operacionais, disponibilidade de equipe, materiais, janela ferroviaria, condicoes climaticas e criticidade do ativo.",
      ],
      { fillColor: THEME.lightGray, minHeight: 20 },
    );
  }

  function finalizeFooters(doc, diagnostico, config) {
    const totalPages = doc.internal.getNumberOfPages();
    for (let page = 1; page <= totalPages; page += 1) {
      doc.setPage(page);
      addFooter(doc, diagnostico, config, page, totalPages);
    }
  }

  function gerarPDF(diagnostico, config = {}) {
    const jsPDF = ensureLibraries();

    if (!diagnostico) {
      throw new Error("Diagnostico nao informado.");
    }
    if (!diagnostico.resumo) {
      throw new Error("Diagnostico sem resumo disponivel.");
    }
    if (!diagnostico.resumo.total) {
      throw new Error("Nao ha demandas no recorte filtrado para gerar o diagnostico.");
    }

    pdfContext = {
      config,
      generatedAt: config.criadoEm || diagnostico?.meta?.criadoEm || new Date().toISOString(),
    };

    const Charts = getCharts();
    const doc = new jsPDF({
      orientation: "landscape",
      unit: "mm",
      format: "a4",
    });

    const hasCover = isSectionEnabled(config, "capa");
    if (hasCover) {
      addCoverPage(doc, diagnostico, config);
    }

    addExecutiveSummaryIndex(doc, diagnostico, config, !hasCover);

    if (isSectionEnabled(config, "resumo") || isSectionEnabled(config, "kpis")) {
      addResumoPage(doc, diagnostico, Charts, config);
    }
    if (isSectionEnabled(config, "vencimentos")) {
      addStatusPage(doc, diagnostico, Charts);
    }
    if (isSectionEnabled(config, "planejamento") || isSectionEnabled(config, "centros")) {
      addPlanejamentoCentrosPage(doc, diagnostico, Charts, config);
    }
    if (isSectionEnabled(config, "planejadores")) {
      addResponsabilidadesPage(doc, diagnostico, Charts);
    }
    if (isSectionEnabled(config, "km") || isSectionEnabled(config, "tolerancias")) {
      addKmRiscosPage(doc, diagnostico, Charts, config);
    }
    if ((diagnostico.patios?.resumo?.totalPatio || 0) > 0) {
      addPatiosRiscosPage(doc, diagnostico, Charts);
    }
    if (isSectionEnabled(config, "clima")) {
      addClimaPage(doc, diagnostico, Charts);
    }
    if (isSectionEnabled(config, "listaCritica")) {
      addListaCriticaPage(doc, diagnostico);
    }
    if (isSectionEnabled(config, "recomendacoes")) {
      addRecomendacoesPage(doc, diagnostico);
    }

    finalizeFooters(doc, diagnostico, config);
    doc.save(downloadFileName(diagnostico, config));
    return doc;
  }

  function initDiagnosticoPDF(options = {}) {
    pdfContext = { ...options };
    console.log("Diagnostico PDF inicializado", options);
    return options;
  }

  global.CCEDiagnosticoPDF = {
    init: initDiagnosticoPDF,
    gerarPDF,
  };
})(window);

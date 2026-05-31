(function setupDiagnosticoCharts(global) {
  "use strict";

  const CHART_THEME = {
    green: [0, 102, 51],
    greenLight: [220, 252, 231],
    blue: [37, 99, 235],
    orange: [217, 119, 6],
    red: [185, 28, 28],
    purple: [109, 40, 217],
    gray: [100, 116, 139],
    grayDark: [30, 41, 59],
    grayLight: [241, 245, 249],
    border: [226, 232, 240],
    white: [255, 255, 255],
  };

  let chartsContext = {};

  function safeText(value, fallback = "-") {
    const text = String(value ?? "").trim();
    return text || fallback;
  }

  function formatNumber(value) {
    return new Intl.NumberFormat("pt-BR").format(Number(value) || 0);
  }

  function formatPercent(value, total) {
    if (!total) return "0,0%";
    return `${(((Number(value) || 0) / total) * 100).toFixed(1).replace(".", ",")}%`;
  }

  function normalizeRows(data, labelKey, valueKey, maxItems) {
    return (data || [])
      .slice(0, maxItems || data.length || 0)
      .map((item) => ({
        label: safeText(item?.[labelKey] ?? item?.label),
        value: Number(item?.[valueKey] ?? item?.value ?? 0),
        color: item?.color || null,
      }))
      .filter((item) => item.value > 0);
  }

  function truncateText(doc, text, maxWidth) {
    const source = safeText(text, "");
    if (!source) return "";
    if (doc.getTextWidth(source) <= maxWidth) return source;

    let current = source;
    while (current.length > 1 && doc.getTextWidth(`${current}...`) > maxWidth) {
      current = current.slice(0, -1);
    }
    return `${current}...`;
  }

  function formatShortDateLabel(value) {
    const text = String(value || "").trim();

    if (/^\d{4}-\d{2}-\d{2}/.test(text)) {
      const [, month, day] = text.slice(0, 10).split("-");
      return `${day}/${month}`;
    }

    if (/^\d{2}\/\d{2}\/\d{4}/.test(text)) {
      return text.slice(0, 5);
    }

    return text.length > 8 ? text.slice(0, 8) : text;
  }

  function drawEmptyState(doc, x, y, w, h, title) {
    doc.setDrawColor(...CHART_THEME.border);
    doc.setFillColor(...CHART_THEME.white);
    doc.roundedRect(x, y, w, h, 3, 3, "FD");
    doc.setFont("helvetica", "bold");
    doc.setFontSize(9);
    doc.setTextColor(...CHART_THEME.grayDark);
    doc.text(safeText(title), x + 3, y + 6);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(8);
    doc.setTextColor(...CHART_THEME.gray);
    doc.text("Sem dados para exibir", x + w / 2, y + h / 2, {
      align: "center",
    });
  }

  function drawChartFrame(doc, x, y, w, h, title) {
    doc.setDrawColor(...CHART_THEME.border);
    doc.setFillColor(...CHART_THEME.white);
    doc.roundedRect(x, y, w, h, 3, 3, "FD");
    doc.setFont("helvetica", "bold");
    doc.setFontSize(9);
    doc.setTextColor(...CHART_THEME.grayDark);
    doc.text(safeText(title), x + 3, y + 6);
  }

  function drawLegend(doc, options = {}) {
    const {
      x = 0,
      y = 0,
      items = [],
      size = 3.2,
      gap = 6,
      fontSize = 7.5,
    } = options;
    let cursorY = y;

    items.forEach((item) => {
      const color = item.color || CHART_THEME.green;
      doc.setFillColor(...color);
      doc.roundedRect(x, cursorY - 2.4, size, size, 0.8, 0.8, "F");
      doc.setFont("helvetica", "normal");
      doc.setFontSize(fontSize);
      doc.setTextColor(...CHART_THEME.grayDark);
      doc.text(safeText(item.label), x + size + 2, cursorY);
      cursorY += gap;
    });

    return cursorY;
  }

  function drawHorizontalBarChart(doc, options = {}) {
    const {
      x = 0,
      y = 0,
      w = 100,
      h = 60,
      title = "Grafico",
      data = [],
      labelKey = "label",
      valueKey = "value",
      maxItems = 10,
      color = CHART_THEME.green,
      showPercent = false,
      total = null,
    } = options;

    const rows = normalizeRows(data, labelKey, valueKey, maxItems);
    if (!rows.length) {
      drawEmptyState(doc, x, y, w, h, title);
      return;
    }

    drawChartFrame(doc, x, y, w, h, title);
    const contentTop = y + 11;
    const rowHeight = Math.min(6.4, (h - 16) / rows.length);
    const labelWidth = Math.max(24, w * 0.3);
    const valueWidth = Math.max(18, w * 0.18);
    const barX = x + labelWidth + 5;
    const barW = w - labelWidth - valueWidth - 12;
    const maxValue = Math.max(...rows.map((item) => item.value), 1);

    rows.forEach((item, index) => {
      const rowY = contentTop + index * rowHeight;
      const width = Math.max(1, (item.value / maxValue) * barW);
      const display = showPercent
        ? `${formatNumber(item.value)} | ${formatPercent(item.value, total || rows.reduce((sum, row) => sum + row.value, 0))}`
        : formatNumber(item.value);

      doc.setFont("helvetica", "normal");
      doc.setFontSize(7.3);
      doc.setTextColor(...CHART_THEME.grayDark);
      doc.text(truncateText(doc, item.label, labelWidth - 2), x + 3, rowY + 2.8);

      doc.setFillColor(...CHART_THEME.grayLight);
      doc.roundedRect(barX, rowY, barW, 3.6, 1, 1, "F");
      doc.setFillColor(...(item.color || color));
      doc.roundedRect(barX, rowY, width, 3.6, 1, 1, "F");

      doc.setTextColor(...CHART_THEME.grayDark);
      doc.text(display, x + w - 3, rowY + 2.8, { align: "right" });
    });
  }

  function drawVerticalBarChart(doc, options = {}) {
    const {
      x = 0,
      y = 0,
      w = 100,
      h = 60,
      title = "Grafico",
      data = [],
      labelKey = "label",
      valueKey = "value",
      maxItems = 12,
      color = CHART_THEME.blue,
    } = options;

    const rows = normalizeRows(data, labelKey, valueKey, maxItems);
    if (!rows.length) {
      drawEmptyState(doc, x, y, w, h, title);
      return;
    }

    drawChartFrame(doc, x, y, w, h, title);
    const chartLeft = x + 8;
    const chartBottom = y + h - 9;
    const chartTop = y + 14;
    const chartHeight = chartBottom - chartTop;
    const chartWidth = w - 14;
    const maxValue = Math.max(...rows.map((item) => item.value), 1);
    const gap = 2;
    const barWidth = Math.max(4, (chartWidth - gap * (rows.length - 1)) / rows.length);

    doc.setDrawColor(...CHART_THEME.border);
    doc.line(chartLeft, chartBottom, chartLeft + chartWidth, chartBottom);

    rows.forEach((item, index) => {
      const barHeight = Math.max(1, (item.value / maxValue) * chartHeight);
      const barX = chartLeft + index * (barWidth + gap);
      const barY = chartBottom - barHeight;
      doc.setFillColor(...(item.color || color));
      doc.roundedRect(barX, barY, barWidth, barHeight, 1, 1, "F");

      doc.setFont("helvetica", "bold");
      doc.setFontSize(6.8);
      doc.setTextColor(...CHART_THEME.grayDark);
      doc.text(formatNumber(item.value), barX + barWidth / 2, barY - 1.2, { align: "center" });

      doc.setFont("helvetica", "normal");
      doc.setFontSize(6.5);
      const shouldShowLabel =
        barWidth >= 12 || (barWidth >= 8 ? index % 2 === 0 : index % 3 === 0);
      if (shouldShowLabel) {
        doc.text(
          formatShortDateLabel(item.label),
          barX + barWidth / 2,
          chartBottom + 4.8,
          { align: "center" },
        );
      }
    });
  }

  function drawDonutChart(doc, options = {}) {
    const {
      x = 0,
      y = 0,
      size = 40,
      title = "Distribuicao",
      data = [],
      labelKey = "label",
      valueKey = "value",
      colors = [],
    } = options;

    const rows = normalizeRows(data, labelKey, valueKey, 8);
    const boxWidth = size + 60;
    if (!rows.length) {
      drawEmptyState(doc, x, y, boxWidth, size + 16, title);
      return;
    }

    drawChartFrame(doc, x, y, boxWidth, size + 16, title);
    const total = rows.reduce((sum, row) => sum + row.value, 0);
    const centerX = x + size / 2 + 4;
    const centerY = y + size / 2 + 10;
    const radius = size / 2 - 2;

    doc.setFillColor(...CHART_THEME.grayLight);
    doc.circle(centerX, centerY, radius, "F");
    doc.setFillColor(...CHART_THEME.white);
    doc.circle(centerX, centerY, radius * 0.58, "F");

    let startY = y + 17;
    rows.forEach((item, index) => {
      const lineY = startY + index * 5.8;
      doc.setFillColor(...(item.color || colors[index] || CHART_THEME.green));
      doc.roundedRect(x + size + 10, lineY - 2.2, 3, 3, 0.6, 0.6, "F");
      doc.setFont("helvetica", "normal");
      doc.setFontSize(7.3);
      doc.setTextColor(...CHART_THEME.grayDark);
      doc.text(
        `${truncateText(doc, item.label, 32)} ${formatNumber(item.value)} (${formatPercent(item.value, total)})`,
        x + size + 15,
        lineY,
      );
    });

    doc.setFont("helvetica", "bold");
    doc.setFontSize(12);
    doc.setTextColor(...CHART_THEME.grayDark);
    doc.text(formatNumber(total), centerX, centerY - 1, { align: "center" });
    doc.setFont("helvetica", "normal");
    doc.setFontSize(7);
    doc.text("Total", centerX, centerY + 3.8, { align: "center" });
  }

  function drawRiskMeter(doc, options = {}) {
    const {
      x = 0,
      y = 0,
      w = 100,
      h = 22,
      title = "Risco",
      score = 0,
      nivel = "Baixo",
    } = options;

    drawChartFrame(doc, x, y, w, h, title);
    const scoreValue = Math.max(0, Math.min(100, Number(score) || 0));
    const meterX = x + 4;
    const meterY = y + 10;
    const meterW = w - 8;
    const segmentW = meterW / 4;
    const segments = [
      CHART_THEME.green,
      CHART_THEME.blue,
      CHART_THEME.orange,
      CHART_THEME.red,
    ];

    segments.forEach((segment, index) => {
      doc.setFillColor(...segment);
      doc.roundedRect(meterX + index * segmentW, meterY, segmentW, 4, 0.8, 0.8, "F");
    });

    const markerX = meterX + (scoreValue / 100) * meterW;
    doc.setDrawColor(...CHART_THEME.grayDark);
    doc.setLineWidth(0.7);
    doc.line(markerX, meterY - 1.5, markerX, meterY + 5.2);
    doc.setFillColor(...CHART_THEME.grayDark);
    doc.circle(markerX, meterY - 1.9, 0.9, "F");

    doc.setFont("helvetica", "bold");
    doc.setFontSize(7.5);
    doc.setTextColor(...CHART_THEME.grayDark);
    doc.text(`Score: ${formatNumber(scoreValue)}/100`, x + 4, y + h - 3);
    doc.text(`Nivel: ${safeText(nivel)}`, x + w - 4, y + h - 3, { align: "right" });
  }

  function drawMiniTimeline(doc, options = {}) {
    return drawVerticalBarChart(doc, {
      ...options,
      maxItems: options.maxItems || 10,
      color: options.color || CHART_THEME.orange,
    });
  }

  function initDiagnosticoCharts(options = {}) {
    chartsContext = { ...options };
    console.log("Diagnostico Charts inicializado", chartsContext);
    return chartsContext;
  }

  global.CCEDiagnosticoCharts = {
    init: initDiagnosticoCharts,
    drawHorizontalBarChart,
    drawVerticalBarChart,
    drawDonutChart,
    drawRiskMeter,
    drawMiniTimeline,
    drawLegend,
  };
})(window);

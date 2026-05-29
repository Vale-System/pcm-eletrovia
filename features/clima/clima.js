(function setupClimateModule(global) {
  const DEFAULT_DATA_URLS = {
    distritos: "./data/distritos_ferrovia.json",
    centros: "./data/centros_trabalho_distritos.json",
    coordenadas: "./data/coordenadas_km.json",
  };

  let externalDataCache = null;
  let externalDataPromise = null;

  function normalizeText(value) {
    return String(value || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toUpperCase()
      .replace(/\s+/g, " ")
      .trim();
  }

  function normalizeKey(value) {
    return normalizeText(value)
      .replace(/[^A-Z0-9]+/g, "_")
      .replace(/^_+|_+$/g, "");
  }

  function escapeHtml(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/\"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function toDateText(value) {
    if (!value) return "";

    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      return value.toISOString().slice(0, 10);
    }

    const text = String(value).trim();
    if (!text) return "";

    if (/^\d{4}-\d{2}-\d{2}/.test(text)) return text.slice(0, 10);

    if (/^\d{2}\/\d{2}\/\d{4}$/.test(text)) {
      const [day, month, year] = text.split("/");
      return `${year}-${month}-${day}`;
    }

    return text;
  }

  function formatDatePt(dateText) {
    if (!dateText || !/^\d{4}-\d{2}-\d{2}$/.test(dateText)) return "-";
    const [year, month, day] = dateText.split("-");
    return `${day}/${month}/${year}`;
  }

  function safeNumber(value) {
    if (value === null || value === undefined || value === "") return null;
    if (typeof value === "number" && Number.isFinite(value)) return value;

    const text = String(value)
      .trim()
      .replace(/KM/gi, "")
      .replace(/TU/gi, "")
      .replace(/\s+/g, "")
      .replace(",", ".");

    if (!text) return null;

    if (/^\d{1,4}\+\d{1,3}$/.test(text)) {
      const [km, metros] = text.split("+");
      return Number(
        (Number(km) + Number(String(metros).padEnd(3, "0")) / 1000).toFixed(3),
      );
    }

    const number = Number(text);
    return Number.isFinite(number) ? number : null;
  }

  function getKmFromBaseJson(demanda) {
    const candidates = [
      demanda.km,
      demanda.Km,
      demanda.KM,
      demanda.kmReferencia,
      demanda.KmReferencia,
      demanda.quilometro,
      demanda.Quilometro,
      demanda.quilometragem,
      demanda.Quilometragem,
    ];

    for (const value of candidates) {
      const km = safeNumber(value);
      if (Number.isFinite(km)) return km;
    }

    return null;
  }

  function extractKmFromText(value) {
    const text = normalizeText(value);
    if (!text) return null;

    const patterns = [
      /\bKM\s*(\d{1,4})[,+.](\d{1,3})\b/,
      /\bKM\s*(\d{1,4})\s*\+\s*(\d{1,3})\b/,
      /\bKM\s*(\d{1,4})\b/,

      /\bTU\s*(\d{1,4})[,+.](\d{1,3})\b/,
      /\bTU\s*(\d{1,4})\s*\+\s*(\d{1,3})\b/,
      /\bTU\s*(\d{1,4})\b/,

      /\b(\d{1,4})\s*\+\s*(\d{1,3})\b/,
    ];

    for (const pattern of patterns) {
      const match = text.match(pattern);
      if (!match) continue;

      const kmInteiro = Number(match[1]);
      if (!Number.isFinite(kmInteiro)) continue;

      if (match[2]) {
        const decimal = Number(String(match[2]).padEnd(3, "0")) / 1000;
        return Number((kmInteiro + decimal).toFixed(3));
      }

      return kmInteiro;
    }

    return null;
  }

  function textHasAny(value, tokens) {
    const text = normalizeText(value);
    return tokens.some((token) => text.includes(normalizeText(token)));
  }

  function demandSearchText(demanda) {
    return [
      demanda.id,
      demanda.ordem,
      demanda.descricao,
      demanda.localInstalacao,
      demanda.centroTrabalho,
      demanda.tipoOM,
      demanda.gerencia,
      demanda.supervisao,
    ].join(" ");
  }

  function getProgramDateInfo(demanda) {
    const replanejada = toDateText(
      demanda.dataReplanejadaAtual || demanda.dataReplanejada,
    );

    if (replanejada) {
      return {
        data: replanejada,
        origem: "Data replanejada",
      };
    }

    const planejada = toDateText(demanda.dataPlanejada);

    if (planejada) {
      return {
        data: planejada,
        origem: "Data planejada",
      };
    }

    const vencimento = toDateText(demanda.vencimento);

    if (vencimento) {
      return {
        data: vencimento,
        origem: "Vencimento",
      };
    }

    return {
      data: "",
      origem: "Sem data",
    };
  }

  function isAllowedStatusForClimate(demanda) {
    const text = normalizeText(
      [
        demanda.statusOperacional,
        demanda.status,
        demanda.substatusOperacional,
        demanda.statusSistema,
        demanda.statusUsuario,
      ].join(" "),
    );

    if (toDateText(demanda.dataRealizada)) return false;
    if (text.includes("REALIZADO")) return false;
    if (text.includes("CANCELADO") || text.includes("CANC")) return false;
    if (text.includes("ENTE") || text.includes("ENCE")) return false;

    return true;
  }

  async function fetchJsonArray(url, label) {
    try {
      const response = await fetch(`${url}?v=${Date.now()}`, {
        cache: "no-store",
        headers: {
          Accept: "application/json",
        },
      });

      if (!response.ok) {
        console.warn(`Clima: ${label} não carregou (${response.status}).`);
        return [];
      }

      const data = await response.json();
      return Array.isArray(data) ? data : [];
    } catch (error) {
      console.warn(`Clima: erro ao carregar ${label}.`, error);
      return [];
    }
  }

  async function loadExternalClimateData(config) {
    if (externalDataCache) return externalDataCache;
    if (externalDataPromise) return externalDataPromise;

    const urls = {
      ...DEFAULT_DATA_URLS,
      ...(config.dataUrls || {}),
    };

    externalDataPromise = Promise.all([
      fetchJsonArray(urls.distritos, "distritos_ferrovia.json"),
      fetchJsonArray(urls.centros, "centros_trabalho_distritos.json"),
      fetchJsonArray(urls.coordenadas, "coordenadas_km.json"),
    ]).then(([distritos, centros, coordenadas]) => {
      externalDataCache = {
        distritos,
        centros,
        coordenadas,
      };

      return externalDataCache;
    });

    return externalDataPromise;
  }

  function normalizeDistrictFromJson(item) {
    const sede = item.sede || item.codigo || item.nome || item.distrito || "";
    const nome = item.distrito || item.nome || sede;
    const codigo = normalizeKey(sede || nome);

    return {
      codigo,
      nome,
      ga: item.gerencia || item.ga || "",
      sede,
      kmInicial: safeNumber(item.km_inicio ?? item.kmInicial),
      kmFinal: safeNumber(item.km_fim ?? item.kmFinal),
      cidadeReferencia: item.cidadeReferencia || nome,
      uf: item.uf || "",
      latitude: item.latitude ?? null,
      longitude: item.longitude ?? null,
      centrosTrabalho: item.centrosTrabalho || [],
      regraEspecial: item.regraEspecial || normalizeText(sede).includes("RFSP"),
      observacao: item.observacao || "",
    };
  }

  function buildEffectiveConfig(config, externalData) {
    const externalDistricts = (externalData.distritos || []).map(
      normalizeDistrictFromJson,
    );

    const configDistricts = (config.distritos || []).map((item) => ({
      ...item,
      codigo:
        item.codigo || normalizeKey(item.sede || item.nome || item.distrito),
      sede: item.sede || item.codigo || item.nome,
      kmInicial: safeNumber(item.kmInicial),
      kmFinal: safeNumber(item.kmFinal),
    }));

    return {
      ...config,

      distritos: externalDistricts.length ? externalDistricts : configDistricts,

      centrosTrabalhoDistritos:
        externalData.centros || config.centrosTrabalhoDistritos || [],

      coordenadasKm: (externalData.coordenadas || config.coordenadasKm || [])
        .map((item) => ({
          ...item,
          km: safeNumber(item.km ?? item.km_key),
          kmKey: safeNumber(item.km_key ?? item.km),
          latitude: safeNumber(item.latitude),
          longitude: safeNumber(item.longitude),
        }))
        .filter((item) => Number.isFinite(item.km))
        .sort((a, b) => a.km - b.km),
    };
  }

  function findDistrictByName(name, config) {
    const key = normalizeText(name);
    if (!key) return null;

    return (
      (config.distritos || []).find((distrito) => {
        return [
          distrito.nome,
          distrito.distrito,
          distrito.sede,
          distrito.codigo,
          distrito.cidadeReferencia,
        ]
          .map(normalizeText)
          .some(
            (value) =>
              value &&
              (value === key || value.includes(key) || key.includes(value)),
          );
      }) || null
    );
  }

  function findDistrictByCodeOrName(value, config) {
    const key = normalizeText(value);
    if (!key) return null;

    return (
      (config.distritos || []).find((distrito) => {
        return [
          distrito.codigo,
          distrito.sede,
          distrito.nome,
          distrito.cidadeReferencia,
        ]
          .map(normalizeText)
          .some(
            (item) =>
              item &&
              (item === key || item.includes(key) || key.includes(item)),
          );
      }) || null
    );
  }

  function findDistrictByKm(km, config) {
    if (!Number.isFinite(km)) return null;

    return (
      (config.distritos || []).find((distrito) => {
        if (!Number.isFinite(distrito.kmInicial)) return false;
        if (!Number.isFinite(distrito.kmFinal)) return false;

        return (
          km >= Number(distrito.kmInicial) && km <= Number(distrito.kmFinal)
        );
      }) || null
    );
  }

  function findNearestCoordinateByKm(km, config) {
    if (!Number.isFinite(km)) return null;

    const coordenadas = config.coordenadasKm || [];

    if (!coordenadas.length) return null;

    let nearest = null;
    let bestDistance = Infinity;

    for (const item of coordenadas) {
      const distance = Math.abs(Number(item.km) - km);

      if (distance < bestDistance) {
        bestDistance = distance;
        nearest = item;
      }
    }

    if (!nearest) return null;

    return {
      ...nearest,
      distanciaKm: Number(bestDistance.toFixed(3)),
    };
  }

  function resolveKmLocation(km, config, tipoVinculo) {
    const coordenada = findNearestCoordinateByKm(km, config);

    const distritoPelaCoordenada = coordenada?.distrito
      ? findDistrictByName(coordenada.distrito, config)
      : null;

    const distritoPorFaixa = findDistrictByKm(km, config);

    const distrito = distritoPelaCoordenada || distritoPorFaixa;

    if (!distrito) return null;

    return {
      tipoVinculo,
      km,
      distrito,
      coordenada,
      semInfluenciaClimatica: false,
    };
  }

  function findDistrictByCentroTrabalho(demanda, config) {
    const centro = normalizeText(demanda.centroTrabalho);

    if (!centro) return null;

    const regra = (config.centrosTrabalhoDistritos || []).find((item) => {
      if (item.ativo === false) return false;
      return normalizeText(item.centroTrabalho) === centro;
    });

    if (!regra) return null;

    const escopo = normalizeText(regra.escopo);

    const influenciaClima =
      regra.influenciaClima ??
      regra.considerarRiscoClimatico ??
      regra.riscoClimatico;

    if (escopo.includes("FERROVIA_TODA") || influenciaClima === false) {
      return {
        tipoVinculo: "Centro sem influência climática",
        km: null,
        distrito: null,
        coordenada: null,
        semInfluenciaClimatica: true,
      };
    }

    const distrito =
      findDistrictByCodeOrName(
        regra.sede || regra.codigoDistrito || regra.codigo || "",
        config,
      ) ||
      findDistrictByName(regra.distrito || regra.nomeDistrito || "", config);

    if (!distrito) return null;

    return {
      tipoVinculo: "Distrito por centro de trabalho",
      km: null,
      distrito,
      coordenada: null,
      semInfluenciaClimatica: false,
    };
  }

  function inferDistrict(demanda, config) {
    const searchText = demandSearchText(demanda);

    const kmBaseJson = getKmFromBaseJson(demanda);

    if (Number.isFinite(kmBaseJson)) {
      const location = resolveKmLocation(kmBaseJson, config, "KM da base JSON");
      if (location) return location;
    }

    if (textHasAny(searchText, ["SLS", "TFPM", "TFPM 1", "TFPM 2"])) {
      const distrito =
        findDistrictByCodeOrName("TFPM", config) ||
        findDistrictByName("São Luís", config);

      const coordenada = findNearestCoordinateByKm(0, config);

      if (distrito) {
        return {
          tipoVinculo: "Regra especial SLS/TFPM",
          km: 0,
          distrito,
          coordenada,
          semInfluenciaClimatica: false,
        };
      }
    }

    if (
      textHasAny(searchText, [
        "RFSP",
        "RAMAL FERROVIARIO SUL DO PARA",
        "RAMAL FERROVIÁRIO SUL DO PARÁ",
      ])
    ) {
      const distrito =
        findDistrictByCodeOrName("RFSP", config) ||
        findDistrictByName("Ramal Ferroviário", config);

      if (distrito) {
        return {
          tipoVinculo: "Regra especial RFSP",
          km: null,
          distrito,
          coordenada: null,
          semInfluenciaClimatica: false,
        };
      }
    }

    const kmDescricao = extractKmFromText(demanda.descricao);

    if (Number.isFinite(kmDescricao)) {
      const location = resolveKmLocation(
        kmDescricao,
        config,
        "KM/TU extraído da descrição",
      );

      if (location) return location;
    }

    const kmLocal = extractKmFromText(demanda.localInstalacao);

    if (Number.isFinite(kmLocal)) {
      const location = resolveKmLocation(
        kmLocal,
        config,
        "KM/TU extraído do local",
      );

      if (location) return location;
    }

    const distritoPorCentro = findDistrictByCentroTrabalho(demanda, config);

    if (distritoPorCentro) return distritoPorCentro;

    return {
      tipoVinculo: "Não identificado",
      km: null,
      distrito: null,
      coordenada: null,
      semInfluenciaClimatica: false,
    };
  }

  function activitySensitivity(demanda, config) {
    const text = normalizeText(
      `${demanda.descricao || ""} ${demanda.tipoOM || ""} ${
        demanda.localInstalacao || ""
      }`,
    );

    const found = (config.atividadesSensiveis || []).find((rule) =>
      text.includes(normalizeText(rule.chave)),
    );

    return found?.sensibilidade || "Normal";
  }

  function isCriticalDemand(demanda) {
    const value = normalizeText(demanda.critico);

    return (
      value === "SIM" ||
      value === "S" ||
      value === "TRUE" ||
      value === "CRITICO" ||
      value === "ALTO"
    );
  }

  let realWeatherCache = null;
  let realWeatherPromise = null;
  let realWeatherCacheKey = "";

  function currentDateTextLocal(timezone = "America/Fortaleza") {
    const parts = new Intl.DateTimeFormat("en-CA", {
      timeZone: timezone,
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
    }).formatToParts(new Date());

    const get = (type) => parts.find((part) => part.type === type)?.value || "";
    return `${get("year")}-${get("month")}-${get("day")}`;
  }

  function currentMonthText() {
    return currentDateTextLocal("America/Fortaleza").slice(0, 7);
  }

  function daysOfMonth(monthText) {
    const [year, month] = monthText.split("-").map(Number);
    const lastDay = new Date(year, month, 0).getDate();
    const days = [];

    for (let day = 1; day <= lastDay; day++) {
      days.push(
        `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(
          2,
          "0",
        )}`,
      );
    }

    return days;
  }

  function dateToNoon(dateText) {
    return new Date(`${dateText}T12:00:00`);
  }

  function weekRangeFromDate(dateText) {
    const base = dateToNoon(dateText || currentDateTextLocal());
    const day = base.getDay();
    const diffToMonday = day === 0 ? -6 : 1 - day;

    const start = new Date(base);
    start.setDate(base.getDate() + diffToMonday);

    const days = [];

    for (let index = 0; index < 7; index++) {
      const current = new Date(start);
      current.setDate(start.getDate() + index);
      days.push(current.toISOString().slice(0, 10));
    }

    return days;
  }

  function sameMonth(dateText, monthText) {
    return String(dateText || "").slice(0, 7) === monthText;
  }

  function getClimateViewMode(container, config) {
    return (
      container.dataset.viewMode || config.weatherApi?.visaoPadrao || "mensal"
    );
  }

  function getPeriodFilter(container) {
    return container.dataset.periodo || "todos";
  }

  function periodKeysFromFilter(value) {
    if (value === "manha") return ["manha"];
    if (value === "tarde") return ["tarde"];
    return ["manha", "tarde"];
  }

  function periodLabelFromFilter(value) {
    if (value === "manha") return "Manhã";
    if (value === "tarde") return "Tarde";
    return "Integral";
  }

  const DATE_ORIGIN_OPTIONS = [
    {
      value: "Data planejada",
      label: "Planejamento",
    },
    {
      value: "Data replanejada",
      label: "Replanejamento",
    },
    {
      value: "Vencimento",
      label: "Vencimento",
    },
  ];

  function getSelectedDateOrigins(container) {
    const raw = container.dataset.origensData || "";

    if (!raw || raw === "todos") {
      return DATE_ORIGIN_OPTIONS.map((item) => item.value);
    }

    const values = raw
      .split("|")
      .map((item) => item.trim())
      .filter(Boolean);

    return values.length
      ? values
      : DATE_ORIGIN_OPTIONS.map((item) => item.value);
  }

  function setSelectedDateOrigins(container, origins) {
    const unique = Array.from(new Set(origins || []));

    if (!unique.length || unique.length === DATE_ORIGIN_OPTIONS.length) {
      container.dataset.origensData = "todos";
      return;
    }

    container.dataset.origensData = unique.join("|");
  }

  function selectedDateOriginsLabel(origins) {
    if (!origins?.length || origins.length === DATE_ORIGIN_OPTIONS.length) {
      return "Todos";
    }

    return DATE_ORIGIN_OPTIONS.filter((item) => origins.includes(item.value))
      .map((item) => item.label)
      .join(" + ");
  }

  function originLabel(value) {
    if (value === "Data replanejada") return "Replanejamento";
    if (value === "Data planejada") return "Planejamento";
    if (value === "Vencimento") return "Vencimento";
    return value || "-";
  }

  function rowMatchesDateOrigin(row, selectedOrigins) {
    if (!Array.isArray(selectedOrigins) || !selectedOrigins.length) return true;
    return selectedOrigins.includes(row.origemDataProgramada);
  }

  function dateMatchesCurrentView(
    dateText,
    monthText,
    viewMode,
    weekReference,
  ) {
    if (viewMode === "semanal") {
      return weekRangeFromDate(weekReference).includes(dateText);
    }

    return sameMonth(dateText, monthText);
  }

  function riskClassFromLevel(level, fallback = "sem-dados") {
    if (level >= 4) return "critico";
    if (level === 3) return "alto";
    if (level === 2) return "atencao";
    if (level === 1) return "favoravel";
    return fallback || "sem-dados";
  }

  function representativePointFromDistrict(distrito, config) {
    if (!distrito) return null;

    if (
      Number.isFinite(Number(distrito.latitude)) &&
      Number.isFinite(Number(distrito.longitude))
    ) {
      return {
        id: distrito.codigo || distrito.sede || distrito.nome,
        nome: distrito.nome,
        distrito: distrito.nome,
        latitude: Number(distrito.latitude),
        longitude: Number(distrito.longitude),
        km:
          Number.isFinite(Number(distrito.kmInicial)) &&
          Number.isFinite(Number(distrito.kmFinal))
            ? Number(
                (
                  (Number(distrito.kmInicial) + Number(distrito.kmFinal)) /
                  2
                ).toFixed(3),
              )
            : null,
        tipo: "distrito",
      };
    }

    const coordenadas = config.coordenadasKm || [];
    if (!coordenadas.length) return null;

    let kmReferencia = null;

    if (
      Number.isFinite(Number(distrito.kmInicial)) &&
      Number.isFinite(Number(distrito.kmFinal))
    ) {
      kmReferencia = Number(
        ((Number(distrito.kmInicial) + Number(distrito.kmFinal)) / 2).toFixed(
          3,
        ),
      );
    } else if (normalizeText(distrito.sede).includes("TFPM")) {
      kmReferencia = 0;
    }

    if (kmReferencia === null) return null;

    let nearest = null;
    let bestDistance = Infinity;

    for (const item of coordenadas) {
      const itemKm = Number(item.km);
      if (!Number.isFinite(itemKm)) continue;

      const distance = Math.abs(itemKm - kmReferencia);

      if (distance < bestDistance) {
        bestDistance = distance;
        nearest = item;
      }
    }

    if (!nearest) return null;

    return {
      id: distrito.codigo || distrito.sede || distrito.nome,
      nome: distrito.nome,
      distrito: distrito.nome,
      latitude: Number(nearest.latitude),
      longitude: Number(nearest.longitude),
      km: Number(nearest.km),
      tipo: "distrito",
    };
  }

  function buildWeatherApiPoints(config) {
    const points = [];

    (config.distritos || []).forEach((distrito) => {
      const isRfsp = normalizeText(
        `${distrito.codigo || ""} ${distrito.sede || ""} ${distrito.nome || ""}`,
      ).includes("RFSP");

      if (isRfsp) return;

      const point = representativePointFromDistrict(distrito, config);
      if (!point) return;

      points.push(point);
    });

    const unique = new Map();

    points.forEach((point) => {
      const key = point.id || `${point.latitude},${point.longitude}`;
      if (!unique.has(key)) unique.set(key, point);
    });

    return Array.from(unique.values());
  }

  function forecastKeyForDistrict(distrito) {
    if (!distrito) return "";
    return String(distrito.codigo || distrito.sede || distrito.nome || "");
  }

  function numberFromPeriod(period, keys, fallback = 0) {
    for (const key of keys) {
      const value = period?.[key];

      if (value !== null && value !== undefined && value !== "") {
        const number = Number(value);
        if (Number.isFinite(number)) return number;
      }
    }

    return fallback;
  }

  function normalizePeriodWeather(period) {
    if (!period) {
      return {
        label: "-",
        precipitationMm: 0,
        precipitationProbability: 0,
        windGustKmh: 0,
        precipitationHours: 0,
        riscoLabel: "Sem dados",
        riscoBase: "sem-dados",
        nivel: 0,
        icone: "☁️",
      };
    }

    return {
      label: period.label || "-",

      precipitationMm: numberFromPeriod(period, [
        "precipitationMm",
        "chuvaMm",
        "precipitation",
        "precipitation_sum",
      ]),

      precipitationProbability: numberFromPeriod(period, [
        "precipitationProbability",
        "probabilidade",
        "probabilidadeChuva",
        "precipitation_probability",
        "precipitation_probability_max",
      ]),

      windGustKmh: numberFromPeriod(period, [
        "windGustKmh",
        "ventoRajadaKmh",
        "wind_gusts_10m",
      ]),

      precipitationHours: numberFromPeriod(period, [
        "precipitationHours",
        "horasChuva",
      ]),

      riscoLabel: period.riscoLabel || "Sem dados",
      riscoBase: period.riscoBase || "sem-dados",
      nivel: Number(period.nivel || 0),
      icone: period.icone || "☁️",
    };
  }

  function normalizeRealWeatherResponse(data) {
    const map = new Map();

    if (!data?.ok || !Array.isArray(data.results)) return map;

    data.results.forEach((result) => {
      const point = result.point || {};
      const key = String(point.id || point.nome || point.distrito || "");
      if (!key) return;

      const dailyMap = new Map();

      (result.daily || []).forEach((day) => {
        if (!day.date) return;

        const periodos = {
          manha: normalizePeriodWeather(day.periodos?.manha),
          tarde: normalizePeriodWeather(day.periodos?.tarde),
        };

        dailyMap.set(day.date, {
          riscoBase: day.riscoBase || "sem-dados",
          riscoLabel: day.riscoLabel || "Sem dados",
          nivel: Number(day.nivel || 0),

          chuvaMm: Number(day.precipitationMm || 0),
          probabilidade: Number(day.precipitationProbability || 0),
          ventoRajadaKmh: Number(day.windGustKmh || 0),
          horasChuva: Number(day.precipitationHours || 0),

          weatherCode: day.weatherCode,
          icone: day.icone || "☁️",
          fonte: "api",

          periodoCritico: day.periodoCritico || "",
          periodos,
          periodosConsiderados: day.periodosConsiderados || data.periodos || [],
          timezone: data.timezone || "America/Fortaleza",
        });
      });

      map.set(key, dailyMap);
    });

    return map;
  }

  async function loadRealWeatherForecast(config, selectedPeriodKeys = []) {
    const apiConfig = config.weatherApi || {};

    if (!apiConfig.enabled || !apiConfig.workerUrl) {
      return new Map();
    }

    const periodos =
      Array.isArray(selectedPeriodKeys) && selectedPeriodKeys.length
        ? selectedPeriodKeys
        : apiConfig.periodosPadrao || ["manha", "tarde"];

    const cacheKey = [
      apiConfig.workerUrl,
      apiConfig.forecastDays || 16,
      periodos.join("|"),
      apiConfig.timezone || "America/Fortaleza",
    ].join("::");

    if (realWeatherCache && realWeatherCacheKey === cacheKey) {
      return realWeatherCache;
    }

    if (realWeatherPromise && realWeatherCacheKey === cacheKey) {
      return realWeatherPromise;
    }

    const points = buildWeatherApiPoints(config);

    if (!points.length) {
      realWeatherCache = new Map();
      realWeatherCacheKey = cacheKey;
      return realWeatherCache;
    }

    realWeatherCacheKey = cacheKey;

    realWeatherPromise = fetch(apiConfig.workerUrl, {
      method: "POST",
      cache: "no-store",
      headers: {
        "Content-Type": "application/json",
        Accept: "application/json",
      },
      body: JSON.stringify({
        forecastDays: apiConfig.forecastDays || 16,
        periodos,
        points,
      }),
    })
      .then(async (response) => {
        if (!response.ok) {
          throw new Error(`Worker clima retornou ${response.status}`);
        }

        return response.json();
      })
      .then((data) => {
        realWeatherCache = normalizeRealWeatherResponse(data);
        return realWeatherCache;
      })
      .catch((error) => {
        console.warn(
          "Clima: falha ao carregar API real. Usando simulado.",
          error,
        );
        realWeatherCache = new Map();
        return realWeatherCache;
      })
      .finally(() => {
        realWeatherPromise = null;
      });

    return realWeatherPromise;
  }

  function realWeatherFor(dateText, distrito, realForecastMap) {
    const key = forecastKeyForDistrict(distrito);
    if (!key || !realForecastMap?.has(key)) return null;

    const dailyMap = realForecastMap.get(key);
    return dailyMap.get(dateText) || null;
  }

  function pseudoWeather(dateText, distrito, coordenada) {
    if (!dateText || !distrito) {
      return {
        riscoBase: "sem-dados",
        riscoLabel: "Sem dados",
        nivel: 0,
        chuvaMm: 0,
        probabilidade: 0,
        ventoRajadaKmh: 0,
        icone: "☁️",
        periodos: {
          manha: normalizePeriodWeather(null),
          tarde: normalizePeriodWeather(null),
        },
      };
    }

    const seedText = `${dateText}-${distrito.codigo}-${coordenada?.km ?? "D"}`;
    let seed = 0;

    for (let i = 0; i < seedText.length; i++) {
      seed += seedText.charCodeAt(i) * (i + 1);
    }

    const chance = seed % 100;
    const chuvaMm = Number(((seed % 35) + (chance > 70 ? 8 : 0)).toFixed(1));

    let base = {
      riscoBase: "favoravel",
      riscoLabel: "Favorável",
      nivel: 1,
      chuvaMm,
      probabilidade: Math.max(5, chance),
      ventoRajadaKmh: 0,
      icone: "☀️",
    };

    if (chuvaMm >= 25 || chance >= 88) {
      base = {
        riscoBase: "critico",
        riscoLabel: "Crítico",
        nivel: 4,
        chuvaMm,
        probabilidade: Math.min(95, chance),
        ventoRajadaKmh: 0,
        icone: "⛈️",
      };
    } else if (chuvaMm >= 12 || chance >= 70) {
      base = {
        riscoBase: "alto",
        riscoLabel: "Alto",
        nivel: 3,
        chuvaMm,
        probabilidade: Math.min(90, chance),
        ventoRajadaKmh: 0,
        icone: "🌧️",
      };
    } else if (chuvaMm >= 4 || chance >= 45) {
      base = {
        riscoBase: "atencao",
        riscoLabel: "Atenção",
        nivel: 2,
        chuvaMm,
        probabilidade: Math.min(80, chance),
        ventoRajadaKmh: 0,
        icone: "🌦️",
      };
    }

    return {
      ...base,
      periodos: {
        manha: normalizePeriodWeather({
          label: "Manhã",
          precipitationMm: chuvaMm / 2,
          precipitationProbability: base.probabilidade,
          windGustKmh: 0,
          riscoLabel: base.riscoLabel,
          riscoBase: base.riscoBase,
          nivel: base.nivel,
          icone: base.icone,
        }),
        tarde: normalizePeriodWeather({
          label: "Tarde",
          precipitationMm: chuvaMm / 2,
          precipitationProbability: base.probabilidade,
          windGustKmh: 0,
          riscoLabel: base.riscoLabel,
          riscoBase: base.riscoBase,
          nivel: base.nivel,
          icone: base.icone,
        }),
      },
    };
  }

  function getWeather(dateText, distrito, coordenada, realForecastMap) {
    const real = realWeatherFor(dateText, distrito, realForecastMap);
    if (real) return real;

    return pseudoWeather(dateText, distrito, coordenada);
  }

  function calculateRisk(demanda, clima, sensibilidade) {
    if (!clima || clima.riscoBase === "sem-dados") {
      return {
        classe: "sem-dados",
        label: "Sem dados",
        nivel: 0,
      };
    }

    let nivel =
      {
        favoravel: 1,
        atencao: 2,
        alto: 3,
        critico: 4,
      }[clima.riscoBase] || 0;

    if (sensibilidade === "Alta" && nivel >= 2) nivel += 1;
    if (isCriticalDemand(demanda) && nivel >= 2) nivel += 1;

    nivel = Math.min(4, nivel);

    const map = {
      1: {
        classe: "favoravel",
        label: "Favorável",
      },
      2: {
        classe: "atencao",
        label: "Atenção",
      },
      3: {
        classe: "alto",
        label: "Alto",
      },
      4: {
        classe: "critico",
        label: "Crítico",
      },
    };

    return {
      ...map[nivel],
      nivel,
    };
  }

  function buildClimateRows(demandas, config, realForecastMap) {
    return (demandas || [])
      .filter(isAllowedStatusForClimate)
      .map((demanda) => {
        const dataInfo = getProgramDateInfo(demanda);
        const localizacao = inferDistrict(demanda, config);

        if (localizacao.semInfluenciaClimatica) return null;

        const sensibilidade = activitySensitivity(demanda, config);

        const clima = getWeather(
          dataInfo.data,
          localizacao.distrito,
          localizacao.coordenada,
          realForecastMap,
        );

        const risco = calculateRisk(demanda, clima, sensibilidade);

        return {
          demanda,
          dataProgramada: dataInfo.data,
          origemDataProgramada: dataInfo.origem,
          localizacao,
          sensibilidade,
          clima,
          risco,
        };
      })
      .filter((row) => row && row.dataProgramada);
  }

  function cleanTooltipDescription(value) {
    return String(value || "Sem descrição")
      .replace(/\bID-\S+/gi, "")
      .replace(/\bOM\s*\d+\b/gi, "")
      .replace(/\b\d{8,}\b/g, "")
      .replace(/\s+/g, " ")
      .trim();
  }

  function climateTooltipHtml(date, weather, items) {
    const manha = normalizePeriodWeather(weather?.periodos?.manha);
    const tarde = normalizePeriodWeather(weather?.periodos?.tarde);

    const probAM = Number(manha.precipitationProbability || 0).toFixed(0);
    const probPM = Number(tarde.precipitationProbability || 0).toFixed(0);
    const mmAM = Number(manha.precipitationMm || 0).toFixed(1);
    const mmPM = Number(tarde.precipitationMm || 0).toFixed(1);
    const mmTotal = Number(weather?.chuvaMm || 0).toFixed(1);
    const vento = Number(weather?.ventoRajadaKmh || 0).toFixed(0);
    const risco = weather?.riscoLabel || "Sem dados";
    const riscoBase = weather?.riscoBase || "sem-dados";

    const riscoColors = {
      critico: { bg: "#fee2e2", color: "#b91c1c", border: "#f3aaa5" },
      alto: { bg: "#ffedd5", color: "#b45309", border: "#ffc18d" },
      atencao: { bg: "#fef3c7", color: "#9a6500", border: "#f1d391" },
      favoravel: { bg: "#dcfce7", color: "#08773e", border: "#bde7c9" },
    };
    const rc = riscoColors[riscoBase] || {
      bg: "#f3f5f4",
      color: "#58726b",
      border: "#d6ddda",
    };

    const periodsEqual = probAM === probPM && mmAM === mmPM;

    const periodRows = periodsEqual
      ? `
        <tr>
          <td style="color:#6b7280;font-size:10px;padding:3px 0 3px 0;white-space:nowrap;">☔ Prob. chuva</td>
          <td style="text-align:right;font-weight:700;font-size:11px;padding:3px 0 3px 8px;">${probAM}%</td>
          <td style="text-align:right;font-size:10px;color:#9ca3af;padding:3px 0 3px 4px;">AM = PM</td>
        </tr>
        <tr>
          <td style="color:#6b7280;font-size:10px;padding:3px 0;white-space:nowrap;">🌧️ Chuva</td>
          <td style="text-align:right;font-weight:700;font-size:11px;padding:3px 0 3px 8px;">${mmAM} mm</td>
          <td style="text-align:right;font-size:10px;color:#9ca3af;padding:3px 0 3px 4px;">AM = PM</td>
        </tr>`
      : `
        <tr>
          <td style="color:#6b7280;font-size:10px;padding:3px 0;white-space:nowrap;">☔ Prob. AM</td>
          <td style="text-align:right;font-weight:700;font-size:11px;padding:3px 0 3px 8px;">${probAM}%</td>
          <td style="text-align:right;font-size:10px;color:#9ca3af;padding:3px 0 3px 4px;"></td>
        </tr>
        <tr>
          <td style="color:#6b7280;font-size:10px;padding:3px 0;white-space:nowrap;">☔ Prob. PM</td>
          <td style="text-align:right;font-weight:700;font-size:11px;padding:3px 0 3px 8px;">${probPM}%</td>
          <td style="text-align:right;font-size:10px;color:#9ca3af;padding:3px 0 3px 4px;"></td>
        </tr>
        <tr>
          <td style="color:#6b7280;font-size:10px;padding:3px 0;white-space:nowrap;">🌧️ Chuva AM</td>
          <td style="text-align:right;font-weight:700;font-size:11px;padding:3px 0 3px 8px;">${mmAM} mm</td>
          <td style="text-align:right;font-size:10px;color:#9ca3af;padding:3px 0 3px 4px;"></td>
        </tr>
        <tr>
          <td style="color:#6b7280;font-size:10px;padding:3px 0;white-space:nowrap;">🌧️ Chuva PM</td>
          <td style="text-align:right;font-weight:700;font-size:11px;padding:3px 0 3px 8px;">${mmPM} mm</td>
          <td style="text-align:right;font-size:10px;color:#9ca3af;padding:3px 0 3px 4px;"></td>
        </tr>`;

    return `
      <div style="
        min-width:210px;max-width:270px;
        background:#1e293b;
        color:#f1f5f9;
        border-radius:14px;
        padding:0;
        box-shadow:0 12px 32px rgba(0,0,0,0.35);
        font-family:inherit;
        overflow:hidden;
        pointer-events:none;
      ">
        <div style="
          background:${rc.bg};
          border-bottom:1px solid ${rc.border};
          padding:10px 14px 9px;
          display:flex;
          align-items:center;
          gap:10px;
        ">
          <span style="font-size:22px;line-height:1;">${weather?.icone || "☁️"}</span>
          <div>
            <div style="font-size:10px;color:${rc.color};font-weight:800;text-transform:uppercase;letter-spacing:0.05em;">${escapeHtml(risco)}</div>
            <div style="font-size:13px;font-weight:900;color:#0f172a;">${formatDatePt(date)}</div>
          </div>
          <div style="margin-left:auto;text-align:right;">
            <div style="font-size:10px;color:#64748b;">Atividades</div>
            <div style="font-size:16px;font-weight:900;color:#0f172a;">${items.length}</div>
          </div>
        </div>

        <div style="padding:10px 14px 12px;">
          <table style="width:100%;border-collapse:collapse;">
            <tbody>
              ${periodRows}
              <tr style="border-top:1px solid #334155;">
                <td style="color:#94a3b8;font-size:10px;padding:6px 0 3px;white-space:nowrap;">💧 Total chuva</td>
                <td style="text-align:right;font-weight:700;font-size:11px;padding:6px 0 3px 8px;">${mmTotal} mm</td>
                <td></td>
              </tr>
              <tr>
                <td style="color:#94a3b8;font-size:10px;padding:3px 0;white-space:nowrap;">💨 Rajada vento</td>
                <td style="text-align:right;font-weight:700;font-size:11px;padding:3px 0 3px 8px;">${vento} km/h</td>
                <td></td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    `
      .replace(/\s*\n\s*/g, "")
      .trim();
  }

  function daysForCalendar(monthText, viewMode, weekReference) {
    if (viewMode === "semanal") return weekRangeFromDate(weekReference);
    return daysOfMonth(monthText);
  }

  function renderCalendar({
    monthText,
    viewMode,
    weekReference,
    distrito,
    rows,
    realForecastMap,
  }) {
    const days = daysForCalendar(monthText, viewMode, weekReference);
    const byDate = new Map();

    rows.forEach((row) => {
      if (!byDate.has(row.dataProgramada)) byDate.set(row.dataProgramada, []);
      byDate.get(row.dataProgramada).push(row);
    });

    return `
    <div class="climate-calendar-grid ${viewMode === "semanal" ? "is-weekly" : "is-monthly"}">
      ${days
        .map((date) => {
          const items = byDate.get(date) || [];

          const weather =
            items[0]?.clima ||
            (distrito
              ? getWeather(date, distrito, null, realForecastMap)
              : {
                  riscoBase: "sem-dados",
                  riscoLabel: "Sem dados",
                  nivel: 0,
                  chuvaMm: 0,
                  probabilidade: 0,
                  ventoRajadaKmh: 0,
                  icone: "☁️",
                  periodos: {
                    manha: normalizePeriodWeather(null),
                    tarde: normalizePeriodWeather(null),
                  },
                });

          const highestRisk = items.reduce(
            (max, item) => Math.max(max, item.risco.nivel || 0),
            Number(weather?.nivel || 0),
          );

          const riskClass = riskClassFromLevel(
            highestRisk,
            weather?.riscoBase || "sem-dados",
          );

          return `
            <div
              class="climate-day climate-risk-${riskClass}"
              data-tooltip="${escapeHtml(climateTooltipHtml(date, weather, items))}"
            >
              <div class="climate-day-top">
                <strong>${date.slice(-2)}</strong>
                <span>${items[0]?.clima?.icone || weather.icone || "☁️"}</span>
              </div>

              <small>${items.length ? `${items.length} ativ.` : weather.riscoLabel}</small>

              <div class="climate-day-metrics">
                <b>${Number(weather.probabilidade || 0).toFixed(0)}%</b>
                <em>${Number(weather.chuvaMm || 0).toFixed(1)} mm</em>
              </div>
            </div>
          `;
        })
        .join("")}
    </div>
  `;
  }

  function initClimateTooltips(container) {
    let tooltipEl = document.getElementById("cce-climate-tooltip");

    if (!tooltipEl) {
      tooltipEl = document.createElement("div");
      tooltipEl.id = "cce-climate-tooltip";
      tooltipEl.style.cssText = [
        "position:fixed",
        "z-index:9999",
        "pointer-events:none",
        "opacity:0",
        "transition:opacity 0.15s ease",
        "max-width:280px",
      ].join(";");
      document.body.appendChild(tooltipEl);
    }

    const days = container.querySelectorAll("[data-tooltip]");

    days.forEach((day) => {
      day.addEventListener("mouseenter", (event) => {
        const html = day.dataset.tooltip;
        if (!html) return;

        tooltipEl.innerHTML = html;
        tooltipEl.style.opacity = "1";
        positionTooltip(event, tooltipEl);
      });

      day.addEventListener("mousemove", (event) => {
        positionTooltip(event, tooltipEl);
      });

      day.addEventListener("mouseleave", () => {
        tooltipEl.style.opacity = "0";
      });
    });

    function positionTooltip(event, el) {
      const margin = 14;
      const vw = window.innerWidth;
      const vh = window.innerHeight;
      const rect = el.getBoundingClientRect();
      const w = rect.width || 270;
      const h = rect.height || 180;

      let x = event.clientX + margin;
      let y = event.clientY + margin;

      if (x + w > vw - 8) x = event.clientX - w - margin;
      if (y + h > vh - 8) y = event.clientY - h - margin;

      el.style.left = `${Math.max(8, x)}px`;
      el.style.top = `${Math.max(8, y)}px`;
    }
  }

  function renderExecutiveView({
    filteredRows,
    riskRows,
    criticalRows,
    kmExactRows,
    month,
    effectiveConfig,
    realForecastMap,
    selectedDistrict,
    quinzena,
  }) {
    const quinzenaNum = quinzena === "2" ? 2 : 1;
    function dayOfWeekShort(dateText) {
      const days = ["Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sáb"];
      return days[new Date(`${dateText}T12:00:00`).getDay()];
    }

    function getDateRange(dates) {
      if (!dates.length) return "";
      const sorted = [...dates].sort();
      const start = sorted[0];
      const end = sorted[sorted.length - 1];
      if (start === end) return formatDatePt(start).slice(0, 5);
      const [, sm, sd] = start.split("-");
      const [, em, ed] = end.split("-");
      if (sm === em) return `${sd}–${ed}/${sm}`;
      return `${formatDatePt(start).slice(0, 5)}–${formatDatePt(end).slice(0, 5)}`;
    }

    function buildRanges(dates) {
      if (!dates.length) return [];
      const ranges = [];
      let current = [dates[0]];
      for (let i = 1; i < dates.length; i++) {
        const prev = new Date(`${dates[i - 1]}T12:00:00`);
        const curr = new Date(`${dates[i]}T12:00:00`);
        const diff = Math.round((curr - prev) / 86400000);
        if (diff <= 1) {
          current.push(dates[i]);
        } else {
          ranges.push(current);
          current = [dates[i]];
        }
      }
      ranges.push(current);
      return ranges;
    }

    const ACTIVITY_TYPES = [
      {
        key: "esmeril",
        label: "Esmerilhamento",
        keywords: ["ESMERIL"],
      },
      {
        key: "trilho",
        label: "Trilho / TLS",
        keywords: ["TRILHO", "TLS", "DORMENTE", "SUBS TRILH"],
      },
      {
        key: "eletrica",
        label: "Elétrica",
        keywords: ["ELETR", "CATEN", "SUBESTAC", "CABINE", "TRANSF"],
      },
      {
        key: "inspecao",
        label: "Inspeção / Roço",
        keywords: ["INSPE", "ROCO", "ROÇO", "PODA", "VISTORI"],
      },
      { key: "normal", label: "Normal", keywords: [] },
    ];

    function classifyRow(row) {
      const text = normalizeText(
        `${row.demanda.descricao || ""} ${row.demanda.tipoOM || ""} ${row.demanda.localInstalacao || ""}`,
      );
      for (const type of ACTIVITY_TYPES) {
        if (type.keywords.some((kw) => text.includes(kw))) return type.key;
      }
      return "normal";
    }

    // Favorable days
    const districtsForFavorable = selectedDistrict
      ? [selectedDistrict]
      : (effectiveConfig.distritos || []).filter(
          (d) =>
            !normalizeText(`${d.codigo || ""} ${d.sede || ""}`).includes(
              "RFSP",
            ),
        );

    const allDays = daysOfMonth(month);

    const diasFavoraveis = allDays.filter((date) => {
      if (!districtsForFavorable.length) return false;
      const worstNivel = Math.max(
        ...districtsForFavorable.map((d) => {
          const w = getWeather(date, d, null, realForecastMap);
          return w.nivel || 0;
        }),
      );
      return worstNivel <= 1;
    });

    // Activity count per day
    const byDateCount = {};
    for (const row of filteredRows) {
      const d = row.dataProgramada;
      byDateCount[d] = (byDateCount[d] || 0) + 1;
    }

    // Risk by date (for banners)
    const byDate = {};
    for (const row of riskRows) {
      const d = row.dataProgramada;
      if (!byDate[d])
        byDate[d] = {
          criticos: 0,
          altos: 0,
          total: 0,
          probMax: 0,
          ventoMax: 0,
          rows: [],
        };
      byDate[d].total += 1;
      if (row.risco.nivel >= 4) byDate[d].criticos += 1;
      if (row.risco.nivel >= 3) byDate[d].altos += 1;
      byDate[d].probMax = Math.max(
        byDate[d].probMax,
        row.clima.probabilidade || 0,
      );
      byDate[d].ventoMax = Math.max(
        byDate[d].ventoMax,
        row.clima.ventoRajadaKmh || 0,
      );
      byDate[d].rows.push(row);
    }

    const sortedRiskDates = Object.keys(byDate).sort();
    const dateRanges = buildRanges(sortedRiskDates);

    const scoredRanges = dateRanges
      .map((dates) => ({
        dates,
        criticos: dates.reduce((s, d) => s + (byDate[d]?.criticos || 0), 0),
        altos: dates.reduce((s, d) => s + (byDate[d]?.altos || 0), 0),
        total: dates.reduce((s, d) => s + (byDate[d]?.total || 0), 0),
        probMax: Math.max(...dates.map((d) => byDate[d]?.probMax || 0)),
        ventoMax: Math.max(...dates.map((d) => byDate[d]?.ventoMax || 0)),
      }))
      .sort(
        (a, b) =>
          b.criticos - a.criticos ||
          b.altos - a.altos ||
          b.total - a.total,
      );

    const topRange = scoredRanges[0] || null;
    const secondRange = scoredRanges[1] || null;

    // Intelligence matrix
    const intelMatrix = {};
    for (const type of ACTIVITY_TYPES)
      intelMatrix[type.key] = { atencao: 0, alto: 0, critico: 0 };
    for (const row of riskRows) {
      const typeKey = classifyRow(row);
      const risco = row.risco.classe;
      if (
        intelMatrix[typeKey] &&
        (risco === "atencao" || risco === "alto" || risco === "critico")
      ) {
        intelMatrix[typeKey][risco] += 1;
      }
    }

    // District risk map
    const distRiskMap = {};
    for (const row of riskRows) {
      const nome = row.localizacao.distrito?.nome || "Não identificado";
      if (!distRiskMap[nome]) distRiskMap[nome] = { count: 0, maxNivel: 0 };
      distRiskMap[nome].count += 1;
      distRiskMap[nome].maxNivel = Math.max(
        distRiskMap[nome].maxNivel,
        row.risco.nivel,
      );
    }
    const top6Distritos = Object.entries(distRiskMap)
      .sort((a, b) => b[1].count - a[1].count)
      .slice(0, 6);

    // Capacity calculations
    const totalEmRisco = riskRows.length;
    const diasFavoraveisCount = diasFavoraveis.length;
    const mediaAtivPorDia =
      allDays.length > 0 ? filteredRows.length / allDays.length : 0;
    const capacidadeRealocacao = Math.round(
      diasFavoraveisCount * mediaAtivPorDia,
    );
    const deficit = Math.max(0, totalEmRisco - capacidadeRealocacao);

    const sensiveisAChuva = riskRows.filter(
      (r) => r.sensibilidade === "Alta",
    ).length;

    // Quick indicators
    const mediaChuvaRisco =
      riskRows.length > 0
        ? (
            riskRows.reduce((s, r) => s + (r.clima.chuvaMm || 0), 0) /
            riskRows.length
          ).toFixed(1)
        : "0.0";

    const rajadaMaxRow =
      riskRows.length > 0
        ? riskRows.reduce((best, r) =>
            (r.clima.ventoRajadaKmh || 0) > (best.clima.ventoRajadaKmh || 0)
              ? r
              : best,
          )
        : null;
    const rajadaMax = rajadaMaxRow
      ? Number(rajadaMaxRow.clima.ventoRajadaKmh || 0).toFixed(0)
      : "0";
    const rajadaMaxDayStr = rajadaMaxRow?.dataProgramada
      ? ` (dia ${rajadaMaxRow.dataProgramada.slice(-2)})`
      : "";

    const distCriticosByDate = {};
    for (const row of criticalRows) {
      const d = row.dataProgramada;
      if (!distCriticosByDate[d]) distCriticosByDate[d] = new Set();
      if (row.localizacao.distrito?.nome)
        distCriticosByDate[d].add(row.localizacao.distrito.nome);
    }
    const maxDistCritico = Math.max(
      0,
      ...Object.values(distCriticosByDate).map((s) => s.size),
    );
    const totalDistritos = (effectiveConfig.distritos || []).filter(
      (d) =>
        !normalizeText(`${d.codigo || ""} ${d.sede || ""}`).includes("RFSP"),
    ).length;

    const semKm = filteredRows.filter(
      (r) => !Number.isFinite(r.localizacao.km),
    ).length;
    const pctSemKm =
      filteredRows.length > 0
        ? Math.round((semKm / filteredRows.length) * 100)
        : 0;

    // Forecast district
    const forecastDistrict =
      selectedDistrict ||
      (() => {
        const topNome = top6Distritos[0]?.[0];
        if (topNome)
          return (
            (effectiveConfig.distritos || []).find((d) => d.nome === topNome) ||
            null
          );
        return (effectiveConfig.distritos || [])[0] || null;
      })();

    const firstQDays = allDays.filter((d) => parseInt(d.slice(-2), 10) <= 15);
    const secondQDays = allDays.filter((d) => parseInt(d.slice(-2), 10) > 15);
    const stripDays = quinzenaNum === 2 ? secondQDays : firstQDays;

    const forecastDays = stripDays.map((date) => {
      const w = forecastDistrict
        ? getWeather(date, forecastDistrict, null, realForecastMap)
        : {
            riscoBase: "sem-dados",
            nivel: 0,
            probabilidade: 0,
            icone: "☁️",
            riscoLabel: "Sem dados",
            chuvaMm: 0,
          };
      return { date, weather: w, actCount: byDateCount[date] || 0 };
    });

    // Banner data
    function getBannerData(range, isFirst) {
      if (!range) return null;
      const dateStr = getDateRange(range.dates);
      const topTiposKeys = [
        ...new Set(
          range.dates.flatMap((d) =>
            (byDate[d]?.rows || [])
              .filter((r) => r.sensibilidade === "Alta")
              .map((r) => classifyRow(r)),
          ),
        ),
      ]
        .filter((k) => k !== "normal")
        .slice(0, 3);
      const topTiposLabels = topTiposKeys.map(
        (k) => ACTIVITY_TYPES.find((t) => t.key === k)?.label || k,
      );
      const prob = Math.round(range.probMax);
      const vento = Math.round(range.ventoMax);

      if (isFirst) {
        return {
          title: `Janela crítica: ${dateStr} · Prob. chuva >${prob}%${vento > 0 ? ` · Vento rajada ${vento} km/h previsto` : ""}`,
          body: `${range.criticos || range.altos} atividades de alto/crítico${topTiposLabels.length ? ` com ${topTiposLabels.join(", ")}` : ""} programadas nessa janela. Recomenda-se revisão imediata.`,
          level: "critico",
          badge: "CRÍTICO",
          icon: "⚠️",
        };
      } else {
        return {
          title: `Próxima janela de atenção: ${dateStr} · ${range.total} atividades afetadas`,
          body: `Prob. chuva ${prob}%${forecastDistrict ? ` em ${forecastDistrict.nome}` : ""}. Avaliar remanejamento${topTiposLabels.length ? ` de ${topTiposLabels[0]}` : ""}.`,
          level: "atencao",
          badge: "ATENÇÃO",
          icon: "🌧️",
        };
      }
    }

    const banner1 = getBannerData(topRange, true);
    const banner2 = getBannerData(secondRange, false);

    // Activity type labels for context
    const topTipoLabels = [
      ...new Set(
        riskRows
          .filter((r) => r.sensibilidade === "Alta")
          .map((r) => classifyRow(r))
          .filter((k) => k !== "normal")
          .map((k) => ACTIVITY_TYPES.find((t) => t.key === k)?.label || k),
      ),
    ].slice(0, 3);

    // Sorted table
    const sensOrder = { Alta: 2, Média: 1, Normal: 0 };
    const execTableRows = [...riskRows].sort((a, b) => {
      if (b.risco.nivel !== a.risco.nivel) return b.risco.nivel - a.risco.nivel;
      const sa = sensOrder[a.sensibilidade] ?? 0;
      const sb = sensOrder[b.sensibilidade] ?? 0;
      if (sb !== sa) return sb - sa;
      return String(a.dataProgramada).localeCompare(String(b.dataProgramada));
    });

    // Recommendation chips + text
    const topDistNomes = top6Distritos.slice(0, 3).map(([nome]) => nome);

    const recParts = [];
    if (deficit > 0) {
      const exemDias = diasFavoraveis
        .slice(0, 4)
        .map((d) => `dia ${d.slice(-2)}`)
        .join(", ");
      recParts.push(
        `Realocar ${deficit} atividades críticas sensíveis para janelas favoráveis (${exemDias || "nenhuma identificada"}).`,
      );
    } else if (diasFavoraveisCount > 0) {
      recParts.push(
        `A capacidade de realocação (${capacidadeRealocacao} ativ.) cobre o volume em risco.`,
      );
    }
    if (topTipoLabels.length) {
      recParts.push(
        `Bloquear autorização de ${topTipoLabels.join(", ")} nos dias críticos.`,
      );
    }
    const recText =
      recParts.join(" ") ||
      "Monitorar previsão e revisar programação em dias de risco alto ou crítico.";

    const chips = [
      topRange?.dates.length
        ? `Bloquear dias ${getDateRange(topRange.dates)}`
        : null,
      diasFavoraveis.length
        ? `Priorizar janela ${getDateRange(diasFavoraveis.slice(0, 5))}`
        : null,
      topTipoLabels.length ? topTipoLabels.join(" · ") : null,
      topDistNomes.length
        ? `Alertar ${topDistNomes.slice(0, 2).join(" e ")}`
        : null,
      semKm > 0 ? `Geocodificar ${semKm} OS sem KM` : null,
    ].filter(Boolean);

    // Presentation constants
    const maxDistBar = top6Distritos[0]?.[1]?.count || 1;
    const progBarMaxVal = Math.max(
      diasFavoraveisCount,
      capacidadeRealocacao,
      totalEmRisco,
      1,
    );
    const matrixRiskCols = ["atencao", "alto", "critico"];
    const matrixColLabels = { atencao: "Atenção", alto: "Alto", critico: "Crítico" };
    const matrixColColors = {
      atencao: { bg: "#fef9e7", color: "#9a6500", border: "#f1d391" },
      alto: { bg: "#fff3e4", color: "#b45309", border: "#ffc18d" },
      critico: { bg: "#fef2f2", color: "#b91c1c", border: "#f3aaa5" },
    };

    return `
      <div class="climate-executive-view">

        ${(banner1 || banner2) ? `
        <div class="climate-exec-banners">
          ${banner1 ? `
          <div class="climate-exec-banner climate-exec-banner-${banner1.level}">
            <span class="climate-exec-banner-icon">${banner1.icon}</span>
            <div class="climate-exec-banner-body">
              <strong>${escapeHtml(banner1.title)}</strong>
              <span>${escapeHtml(banner1.body)}</span>
            </div>
            <span class="climate-exec-badge climate-exec-badge-${banner1.level}">${banner1.badge}</span>
          </div>` : ""}
          ${banner2 ? `
          <div class="climate-exec-banner climate-exec-banner-${banner2.level}">
            <span class="climate-exec-banner-icon">${banner2.icon}</span>
            <div class="climate-exec-banner-body">
              <strong>${escapeHtml(banner2.title)}</strong>
              <span>${escapeHtml(banner2.body)}</span>
            </div>
            <span class="climate-exec-badge climate-exec-badge-${banner2.level}">${banner2.badge}</span>
          </div>` : ""}
        </div>
        ` : ""}

        <div class="climate-exec-kpis">
          <article class="climate-exec-kpi-total">
            <div class="climate-exec-kpi-head">
              <span class="climate-exec-kpi-icon">⚡</span>
              <span>Total em risco</span>
            </div>
            <strong>${riskRows.length}</strong>
            <small>atenção, alto ou crítico</small>
          </article>
          <article class="climate-exec-kpi-item">
            <div class="climate-exec-kpi-head">
              <span class="climate-exec-kpi-icon">🔥</span>
              <span>Críticas</span>
            </div>
            <strong class="exec-color-critico">${criticalRows.length}</strong>
            <small>risco crítico confirmado</small>
          </article>
          <article class="climate-exec-kpi-item">
            <div class="climate-exec-kpi-head">
              <span class="climate-exec-kpi-icon">🌧️</span>
              <span>Sensíveis à chuva</span>
            </div>
            <strong class="exec-color-alto">${sensiveisAChuva}</strong>
            <small>${topTipoLabels.length ? topTipoLabels.join(", ") : "alta sensibilidade"}</small>
            ${sensiveisAChuva > 0 ? `<span class="climate-exec-kpi-tag">Prioridade máxima</span>` : ""}
          </article>
          <article class="climate-exec-kpi-item">
            <div class="climate-exec-kpi-head">
              <span class="climate-exec-kpi-icon">📍</span>
              <span>Com KM/coordenada</span>
            </div>
            <strong>${kmExactRows.length}</strong>
            <small>precisão geoespacial</small>
            ${filteredRows.length > 0 ? `<span class="climate-exec-kpi-sub">${Math.round((kmExactRows.length / filteredRows.length) * 100)}% do total</span>` : ""}
          </article>
          <article class="climate-exec-kpi-item">
            <div class="climate-exec-kpi-head">
              <span class="climate-exec-kpi-icon">🗓️</span>
              <span>Janelas disponíveis</span>
            </div>
            <strong class="exec-color-favoravel">${diasFavoraveisCount}</strong>
            <small>dias favoráveis no mês</small>
            ${diasFavoraveis.length > 0 ? `<span class="climate-exec-kpi-sub">Dias ${diasFavoraveis.slice(0, 4).map((d) => d.slice(-2)).join(", ")}${diasFavoraveis.length > 4 ? "..." : ""}</span>` : ""}
          </article>
        </div>

        <div class="climate-exec-grid-2">
          <div class="climate-exec-left-col">

            <div class="climate-card">
              <div class="climate-card-header">
                <div>
                  <span>Previsão detalhada — ${escapeHtml(forecastDistrict?.nome || "Sem distrito")}</span>
                  <strong>${quinzenaNum === 1 ? `Dias 01–15 de ${escapeHtml(month)}` : `Dias 16–${allDays.length} de ${escapeHtml(month)}`}</strong>
                </div>
                <div class="climate-exec-qswitch">
                  <button class="climate-exec-qbtn ${quinzenaNum === 1 ? "is-active" : ""}" data-quinzena="1" type="button">1ª Quinzena</button>
                  <button class="climate-exec-qbtn ${quinzenaNum === 2 ? "is-active" : ""}" data-quinzena="2" type="button">2ª Quinzena</button>
                </div>
              </div>
              <div class="climate-exec-forecast-strip">
                ${forecastDays.map(({ date, weather, actCount }) => `
                  <div class="climate-exec-fday climate-risk-${weather.riscoBase || "sem-dados"}"
                    title="${escapeHtml(formatDatePt(date))} — ${escapeHtml(weather.riscoLabel)} ${Number(weather.probabilidade || 0).toFixed(0)}%">
                    <span class="climate-exec-fday-dow">${dayOfWeekShort(date)}</span>
                    <span class="climate-exec-fday-icon">${weather.icone || "☁️"}</span>
                    <span class="climate-exec-fday-prob">${Number(weather.probabilidade || 0).toFixed(0)}%</span>
                    <span class="climate-exec-fday-mm">${Number(weather.chuvaMm || 0).toFixed(0)}mm</span>
                    <span class="climate-exec-fday-acts">${actCount > 0 ? `${actCount} atv.` : "—"}</span>
                  </div>
                `).join("")}
              </div>
            </div>

            <div class="climate-card">
              <div class="climate-card-header">
                <div>
                  <span>Exposição por distrito</span>
                  <strong>Atividades críticas + alto</strong>
                </div>
              </div>
              <div class="climate-exec-dist-bars">
                ${top6Distritos.length > 0
                  ? top6Distritos
                      .map(([nome, info]) => {
                        const pct = Math.round((info.count / maxDistBar) * 100);
                        const barClass = riskClassFromLevel(info.maxNivel);
                        return `
                          <div class="climate-exec-bar-row">
                            <span class="climate-exec-bar-label" title="${escapeHtml(nome)}">${escapeHtml(nome)}</span>
                            <div class="climate-exec-bar-track">
                              <div class="climate-exec-bar-fill climate-exec-bar-${barClass}" style="width:${pct}%">
                                <span class="climate-exec-bar-inside">${info.count}</span>
                              </div>
                            </div>
                            <span class="climate-exec-bar-val">${info.count}</span>
                          </div>
                        `;
                      })
                      .join("")
                  : `<div class="climate-empty">Nenhum distrito em risco.</div>`}
              </div>
            </div>

          </div>

          <div class="climate-exec-right-col">
            <div class="climate-card climate-exec-table-card">
              <div class="climate-card-header climate-exec-risk-header">
                <div>
                  <span>Matriz de risco</span>
                  <strong>Atividades em janela crítica — ordenadas por impacto</strong>
                </div>
                <span class="climate-exec-total-badge">${execTableRows.length}<br><small>total</small></span>
              </div>
              <div class="climate-exec-table-wrap">
                <table class="climate-exec-risk-table">
                  <thead>
                    <tr>
                      <th>Data / Clima</th>
                      <th>Risco</th>
                      <th>Atividade</th>
                      <th>KM</th>
                      <th>Sens.</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${execTableRows
                        .slice(0, 80)
                        .map((row) => {
                          const coord = row.localizacao.coordenada;
                          const kmText = Number.isFinite(row.localizacao.km)
                            ? Number(row.localizacao.km)
                                .toFixed(3)
                                .replace(".000", "")
                            : "—";
                          const coordText =
                            coord?.latitude && coord?.longitude
                              ? `${Number(coord.latitude).toFixed(2)}°, ${Number(coord.longitude).toFixed(2)}°`
                              : "";
                          const manha = row.clima.periodos?.manha;
                          const tarde = row.clima.periodos?.tarde;
                          const probAM = Number(
                            manha?.precipitationProbability ||
                              row.clima.probabilidade ||
                              0,
                          ).toFixed(0);
                          const probPM = Number(
                            tarde?.precipitationProbability ||
                              row.clima.probabilidade ||
                              0,
                          ).toFixed(0);
                          const probLine =
                            probAM === probPM
                              ? `${probAM}% AM=PM`
                              : `AM ${probAM}% / PM ${probPM}%`;
                          const sensKey =
                            row.sensibilidade === "Alta"
                              ? "alta"
                              : row.sensibilidade === "Média"
                                ? "media"
                                : "normal";
                          return `
                            <tr>
                              <td class="exec-td-date">
                                <strong>${formatDatePt(row.dataProgramada).slice(0, 5)}</strong>
                                <span class="exec-td-icon">${row.clima.icone || "☁️"} ${Number(row.clima.chuvaMm || 0).toFixed(0)}mm</span>
                                <span class="exec-td-prob">${probLine}</span>
                              </td>
                              <td>
                                <span class="climate-risk-pill climate-risk-${row.risco.classe}">${escapeHtml(row.risco.label.toUpperCase())}</span>
                              </td>
                              <td class="exec-td-activ">
                                <strong>${escapeHtml(row.demanda.ordem || row.demanda.id || "-")}</strong>
                                <span>${escapeHtml(row.demanda.descricao || "-")}</span>
                                <small>${escapeHtml(row.localizacao.distrito?.nome || "Não identificado")}${row.localizacao.distrito?.ga ? ` · ${row.localizacao.distrito.ga}` : ""}</small>
                              </td>
                              <td class="exec-td-km">
                                <strong>${escapeHtml(kmText)}</strong>
                                ${coordText ? `<small>${escapeHtml(coordText)}</small>` : ""}
                              </td>
                              <td>
                                <span class="exec-sens exec-sens-${sensKey}">${escapeHtml(row.sensibilidade)}</span>
                              </td>
                            </tr>
                          `;
                        })
                        .join("") ||
                      `<tr><td colspan="5"><div class="climate-empty">Nenhuma atividade em risco.</div></td></tr>`}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        </div>

        <div class="climate-exec-grid-2">
          <div class="climate-card climate-exec-intel-card">
            <div class="climate-card-header">
              <div>
                <span>Inteligência de risco</span>
                <strong>Matriz sensibilidade × nível climático</strong>
              </div>
            </div>
            <p class="climate-exec-matrix-desc">Qtd. atividades por cruzamento de sensibilidade e risco climático atual</p>
            <div class="climate-exec-matrix-wrap">
              <table class="climate-exec-matrix">
                <thead>
                  <tr>
                    <th></th>
                    ${matrixRiskCols
                        .map(
                          (col) =>
                            `<th style="color:${matrixColColors[col].color};">${matrixColLabels[col]}</th>`,
                        )
                        .join("")}
                  </tr>
                </thead>
                <tbody>
                  ${ACTIVITY_TYPES.map((type) => {
                      const rowData = intelMatrix[type.key] || {};
                      return `
                        <tr>
                          <td class="climate-exec-matrix-label">${escapeHtml(type.label)}</td>
                          ${matrixRiskCols
                              .map((col) => {
                                const count = rowData[col] || 0;
                                return `<td class="climate-exec-matrix-cell" style="${count > 0 ? `background:${matrixColColors[col].bg};border:1px solid ${matrixColColors[col].border};color:${matrixColColors[col].color};` : ""}">${count > 0 ? `<strong>${count}</strong><small>atividades</small>` : `<span style="color:#d1d5db;">—</span>`}</td>`;
                              })
                              .join("")}
                        </tr>
                      `;
                    }).join("")}
                </tbody>
              </table>
            </div>
          </div>

          <div class="climate-card climate-exec-windows-card">
            <div class="climate-card-header">
              <div>
                <span>Gestão de janelas</span>
                <strong>Capacidade de realocação no mês</strong>
              </div>
            </div>
            <div class="climate-exec-windows">

              <div class="climate-exec-wb-list">
                <div class="climate-exec-wb-row">
                  <span class="climate-exec-wb-label">Dias favoráveis disponíveis</span>
                  <div class="climate-exec-wb-right">
                    <div class="climate-exec-wb-track">
                      <div class="climate-exec-wb-fill wbf-favoravel" style="width:${Math.min(100, Math.round((diasFavoraveisCount / progBarMaxVal) * 100))}%"></div>
                    </div>
                    <span class="climate-exec-wb-val exec-color-favoravel">${diasFavoraveisCount} dias</span>
                  </div>
                </div>
                <div class="climate-exec-wb-row">
                  <span class="climate-exec-wb-label">Capacidade teórica de realocação</span>
                  <div class="climate-exec-wb-right">
                    <div class="climate-exec-wb-track">
                      <div class="climate-exec-wb-fill wbf-azul" style="width:${Math.min(100, Math.round((capacidadeRealocacao / progBarMaxVal) * 100))}%"></div>
                    </div>
                    <span class="climate-exec-wb-val">~${capacidadeRealocacao} atividades</span>
                  </div>
                </div>
                <div class="climate-exec-wb-row">
                  <span class="climate-exec-wb-label">Déficit estimado (sem janelas)</span>
                  <div class="climate-exec-wb-right">
                    <div class="climate-exec-wb-track">
                      <div class="climate-exec-wb-fill wbf-critico" style="width:${Math.min(100, Math.round((deficit / progBarMaxVal) * 100))}%"></div>
                    </div>
                    <span class="climate-exec-wb-val exec-color-critico">${deficit} atividades</span>
                  </div>
                </div>
              </div>

              ${diasFavoraveis.length > 0 ? `
              <div class="climate-exec-fav-section">
                <span class="climate-exec-fav-label">Janelas favoráveis identificadas</span>
                <div class="climate-exec-fav-pills">
                  ${diasFavoraveis
                      .slice(0, 10)
                      .map((date) => {
                        const w = forecastDistrict
                          ? getWeather(date, forecastDistrict, null, realForecastMap)
                          : { icone: "☀️", probabilidade: 0 };
                        return `<span class="climate-exec-fav-pill" title="${escapeHtml(formatDatePt(date))}">Dia ${date.slice(-2)} ${w.icone || "☀️"} ${Number(w.probabilidade || 0).toFixed(0)}%</span>`;
                      })
                      .join("")}
                </div>
              </div>` : ""}

              <div class="climate-exec-indicators">
                <span class="climate-exec-fav-label">Indicadores adicionais</span>
                <div class="climate-exec-ind-row">
                  <span>Média de chuva no período crítico</span>
                  <span class="exec-ind-val exec-color-alto">${mediaChuvaRisco} mm/dia</span>
                </div>
                <div class="climate-exec-ind-row">
                  <span>Rajada máxima prevista</span>
                  <span class="exec-ind-val exec-color-alto">${rajadaMax} km/h${escapeHtml(rajadaMaxDayStr)}</span>
                </div>
                <div class="climate-exec-ind-row">
                  <span>Distritos com risco crítico simultâneo</span>
                  <span class="exec-ind-val exec-color-critico">${maxDistCritico}${totalDistritos > 0 ? ` de ${totalDistritos}` : ""}</span>
                </div>
                <div class="climate-exec-ind-row">
                  <span>% atividades sem KM/localização</span>
                  <span class="exec-ind-val exec-color-alto">${pctSemKm}%${semKm > 0 ? ` (${semKm} OS)` : ""}</span>
                </div>
              </div>

            </div>
          </div>
        </div>

        <div class="climate-exec-rec">
          <div class="climate-exec-rec-icon">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" aria-hidden="true">
              <path d="M12 2 5 5v6c0 5 3.3 9.4 7 11 3.7-1.6 7-6 7-11V5Z"/>
              <path d="m9 12 2 2 4-5"/>
            </svg>
          </div>
          <div class="climate-exec-rec-body">
            <span>Recomendação do sistema</span>
            <p>${escapeHtml(recText)}</p>
            ${chips.length > 0 ? `
            <div class="climate-exec-rec-chips">
              ${chips.map((c) => `<span class="climate-exec-rec-chip">${escapeHtml(c)}</span>`).join("")}
            </div>` : ""}
          </div>
        </div>

      </div>
    `;
  }

  async function render({ container, demandas, config }) {
    container.innerHTML = `
    <div class="climate-card">
      <div class="climate-empty">
        Carregando regras climáticas, distritos, coordenadas e previsão...
      </div>
    </div>
  `;

    const externalData = await loadExternalClimateData(config);
    const effectiveConfig = buildEffectiveConfig(config || {}, externalData);

    const month = container.dataset.month || currentMonthText();
    const viewMode = getClimateViewMode(container, effectiveConfig);
    const weekReference =
      container.dataset.weekReference ||
      currentDateTextLocal(
        effectiveConfig.weatherApi?.timezone || "America/Fortaleza",
      );

    const periodoFiltro = getPeriodFilter(container);
    const selectedPeriodKeys = periodKeysFromFilter(periodoFiltro);
    const selectedDateOrigins = getSelectedDateOrigins(container);

    const realForecastMap = await loadRealWeatherForecast(
      effectiveConfig,
      selectedPeriodKeys,
    );

    const rows = buildClimateRows(demandas, effectiveConfig, realForecastMap);

    const selectedDistrictCode = container.dataset.district || "__ALL__";

    const selectedDistrict =
      selectedDistrictCode === "__ALL__"
        ? null
        : (effectiveConfig.distritos || []).find(
            (item) => item.codigo === selectedDistrictCode,
          ) || null;

    const scopedDateRows = rows.filter((row) => {
      if (!rowMatchesDateOrigin(row, selectedDateOrigins)) return false;

      return dateMatchesCurrentView(
        row.dataProgramada,
        month,
        viewMode,
        weekReference,
      );
    });

    const filteredRows = scopedDateRows.filter((row) => {
      if (!selectedDistrict) return true;
      return row.localizacao.distrito?.codigo === selectedDistrict.codigo;
    });

    const riskRows = filteredRows
      .filter((row) => row.risco.nivel >= 2)
      .sort((a, b) => {
        if (b.risco.nivel !== a.risco.nivel)
          return b.risco.nivel - a.risco.nivel;
        return String(a.dataProgramada).localeCompare(String(b.dataProgramada));
      });

    const criticalRows = riskRows.filter((row) => row.risco.nivel >= 4);
    const highRows = riskRows.filter((row) => row.risco.nivel >= 3);
    const unidentifiedRows = scopedDateRows.filter(
      (row) => !row.localizacao.distrito,
    );
    const kmExactRows = filteredRows.filter((row) =>
      Number.isFinite(row.localizacao.km),
    );

    const worstDay = riskRows[0];
    const periodLabel = periodLabelFromFilter(periodoFiltro);
    const calendarHeaderDate =
      viewMode === "semanal"
        ? `Semana de ${formatDatePt(weekReference)}`
        : month;

    const activeTab = container.dataset.activeTab || "calendario";
    const quinzena = container.dataset.quinzena || "1";

    container.innerHTML = `
    <div class="climate-page">
      <div class="climate-hero">
        <div class="climate-title-area">
          <span class="climate-eyebrow">Clima & Programação Operacional</span>
          <h2>Risco Climático da Eletrovia</h2>
        </div>

        <div class="climate-filters climate-filters-advanced">
          <label>
            <span>Distrito</span>
            <select id="climateDistrict">
              <option value="__ALL__" ${
                selectedDistrictCode === "__ALL__" ? "selected" : ""
              }>Todos os distritos</option>
              ${(effectiveConfig.distritos || [])
                .map(
                  (distrito) =>
                    `<option value="${escapeHtml(distrito.codigo)}" ${
                      distrito.codigo === selectedDistrict?.codigo
                        ? "selected"
                        : ""
                    }>${escapeHtml(distrito.nome)}</option>`,
                )
                .join("")}
            </select>
          </label>

          <label>
            <span>Visão</span>
            <select id="climateViewMode">
              <option value="mensal" ${viewMode === "mensal" ? "selected" : ""}>Mensal</option>
              <option value="semanal" ${viewMode === "semanal" ? "selected" : ""}>Semanal</option>
            </select>
          </label>

          <label>
            <span>Mês</span>
            <input id="climateMonth" type="month" value="${escapeHtml(month)}" />
          </label>

          <label class="${viewMode === "semanal" ? "" : "is-muted-control"}">
            <span>Semana referência</span>
            <input id="climateWeekReference" type="date" value="${escapeHtml(weekReference)}" />
          </label>

          <label>
            <span>Período</span>
            <select id="climatePeriodFilter">
              <option value="todos" ${periodoFiltro === "todos" ? "selected" : ""}>Todos</option>
              <option value="manha" ${periodoFiltro === "manha" ? "selected" : ""}>Manhã</option>
              <option value="tarde" ${periodoFiltro === "tarde" ? "selected" : ""}>Tarde</option>
            </select>
          </label>

          <div class="climate-multi-field">
            <span>Status da data</span>

            <div class="climate-multi-select" id="climateDateOriginMulti">
              <button class="climate-multi-button" type="button" id="climateDateOriginButton">
                ${escapeHtml(selectedDateOriginsLabel(selectedDateOrigins))}
              </button>

              <div class="climate-multi-panel">
                ${DATE_ORIGIN_OPTIONS.map(
                  (option) => `
                    <label>
                      <input
                        type="checkbox"
                        data-climate-date-origin="${escapeHtml(option.value)}"
                        ${selectedDateOrigins.includes(option.value) ? "checked" : ""}
                      />
                      ${escapeHtml(option.label)}
                    </label>
                  `,
                ).join("")}
              </div>
            </div>
          </div>
        </div>
      </div>

      <div class="climate-tabs" role="tablist">
        <button class="climate-tab ${activeTab === "calendario" ? "is-active" : ""}" data-climate-tab="calendario" role="tab" type="button">Calendário</button>
        <button class="climate-tab ${activeTab === "executivo" ? "is-active" : ""}" data-climate-tab="executivo" role="tab" type="button">Visão Executiva</button>
      </div>

      ${activeTab === "executivo" ? renderExecutiveView({
        filteredRows,
        riskRows,
        criticalRows,
        kmExactRows,
        month,
        effectiveConfig,
        realForecastMap,
        selectedDistrict,
        quinzena,
      }) : `
      <div class="climate-kpis">
        <article>
          <span>Total em risco</span>
          <strong>${riskRows.length}</strong>
          <small>atenção, alto ou crítico</small>
        </article>

        <article>
          <span>Críticas</span>
          <strong>${criticalRows.length}</strong>
          <small>risco climático crítico</small>
        </article>

        <article>
          <span>Com KM/coordenada</span>
          <strong>${kmExactRows.length}</strong>
          <small>base JSON, KM ou TU extraído</small>
        </article>

        <article>
          <span>Sem distrito</span>
          <strong>${unidentifiedRows.length}</strong>
          <small>precisam regra de centro/local</small>
        </article>
      </div>

      <div class="climate-layout">
        <section class="climate-card climate-calendar-card">
          <div class="climate-card-header">
            <div>
              <span>Calendário climático</span>
              <strong>${escapeHtml(selectedDistrict?.nome || "Todos os distritos")}</strong>
            </div>
            <small>${escapeHtml(calendarHeaderDate)} | ${escapeHtml(periodLabel)}</small>
          </div>

          ${renderCalendar({
            monthText: month,
            viewMode,
            weekReference,
            distrito: selectedDistrict,
            rows: filteredRows,
            realForecastMap,
          })}

          <div class="climate-legend">
            <span><i class="legend-dot favoravel"></i>Favorável</span>
            <span><i class="legend-dot atencao"></i>Atenção</span>
            <span><i class="legend-dot alto"></i>Alto</span>
            <span><i class="legend-dot critico"></i>Crítico</span>
          </div>
        </section>

        <section class="climate-card climate-table-card">
          <div class="climate-card-header climate-danger-header">
            <div>
              <span>Atividades programadas em dias de risco</span>
              <strong>${riskRows.length} ocorrência(s)</strong>
            </div>
          </div>

          <div class="climate-risk-table-wrap">
            <table class="climate-risk-table">
              <thead>
                <tr>
                  <th>Data</th>
                  <th>Risco</th>
                  <th>Atividade / Descrição</th>
                  <th>Distrito</th>
                  <th>KM / Coordenada</th>
                  <th>Vínculo</th>
                </tr>
              </thead>

              <tbody>
                ${
                  riskRows
                    .slice(0, 120)
                    .map((row) => {
                      const demanda = row.demanda;
                      const coord = row.localizacao.coordenada;

                      const kmText = Number.isFinite(row.localizacao.km)
                        ? Number(row.localizacao.km)
                            .toFixed(3)
                            .replace(".000", "")
                        : "-";

                      const coordText =
                        coord?.latitude && coord?.longitude
                          ? `${Number(coord.latitude).toFixed(5)}, ${Number(
                              coord.longitude,
                            ).toFixed(5)}`
                          : "Sem coordenada";

                      return `
                        <tr>
                          <td>
                            <strong>${formatDatePt(row.dataProgramada)}</strong>
                            <small>${escapeHtml(row.clima.icone)} ${Number(row.clima.chuvaMm || 0).toFixed(1)} mm</small>
                            <small>Prob. chuva ${Number(row.clima.probabilidade || 0).toFixed(0)}%</small>
                            <small>${escapeHtml(originLabel(row.origemDataProgramada))}</small>
                          </td>

                          <td>
                            <span class="climate-risk-pill climate-risk-${row.risco.classe}">
                              ${escapeHtml(row.risco.label)}
                            </span>
                          </td>

                          <td>
                            <strong>${escapeHtml(demanda.ordem || demanda.id || "-")}</strong>
                            <small>${escapeHtml(demanda.descricao || "-")}</small>
                          </td>

                          <td>
                            <strong>${escapeHtml(
                              row.localizacao.distrito?.nome ||
                                "Não identificado",
                            )}</strong>
                            <small>${escapeHtml(row.localizacao.distrito?.ga || "")}</small>
                          </td>

                          <td>
                            <strong>${escapeHtml(kmText)}</strong>
                            <small>${escapeHtml(coordText)}</small>
                          </td>

                          <td>${escapeHtml(row.localizacao.tipoVinculo)}</td>
                        </tr>
                      `;
                    })
                    .join("") ||
                  `
                    <tr>
                      <td colspan="6">
                        <div class="climate-empty">
                          Nenhuma atividade em risco para o filtro selecionado.
                        </div>
                      </td>
                    </tr>
                  `
                }
              </tbody>
            </table>
          </div>
        </section>
      </div>

      <div class="climate-bottom-grid">
        <section class="climate-card climate-summary-card">
          <div class="climate-card-header">
            <div>
              <span>Resumo executivo</span>
              <strong>${riskRows.length} atividades em atenção climática</strong>
            </div>
          </div>

          <div class="climate-summary-content">
            <div>
              <span class="climate-big-number">${riskRows.length}</span>
              <small>Total de atividades com risco climático</small>
            </div>

            <ul>
              <li><strong>${criticalRows.length}</strong> em nível crítico</li>
              <li><strong>${highRows.length}</strong> em alto/crítico</li>
              <li><strong>${kmExactRows.length}</strong> com KM/coordenada</li>
              <li><strong>${unidentifiedRows.length}</strong> sem distrito identificado</li>
            </ul>
          </div>
        </section>

        <section class="climate-card climate-attention-card">
          <div class="climate-alert-icon">!</div>

          <div>
            <span>Maior ponto de atenção</span>

            <strong>
              ${
                worstDay
                  ? `${formatDatePt(worstDay.dataProgramada)} — ${worstDay.risco.label}`
                  : "Sem risco relevante"
              }
            </strong>

            <p>
              ${
                worstDay
                  ? `Atividade em ${escapeHtml(
                      worstDay.localizacao.distrito?.nome ||
                        "distrito não identificado",
                    )}. Probabilidade ${Number(worstDay.clima.probabilidade || 0).toFixed(0)}%.`
                  : "Nenhuma atividade programada com risco climático alto ou crítico neste filtro."
              }
            </p>
          </div>
        </section>

        <section class="climate-card climate-recommendation-card">
          <div class="climate-shield">✓</div>

          <div>
            <span>Recomendação</span>
            <strong>Reavaliar programação em dias de risco alto/crítico</strong>
            <p>
              Priorizar atividades críticas, externas, de trilho, TLS,
              esmerilhamento e intervenções sensíveis à chuva.
            </p>
          </div>
        </section>
      </div>
      `}
    </div>
  `;

    initClimateTooltips(container);

    container.querySelectorAll("[data-climate-tab]").forEach((btn) => {
      btn.addEventListener("click", () => {
        container.dataset.activeTab = btn.dataset.climateTab;
        render({ container, demandas, config });
      });
    });

    container.querySelectorAll("[data-quinzena]").forEach((btn) => {
      btn.addEventListener("click", () => {
        container.dataset.quinzena = btn.dataset.quinzena;
        render({ container, demandas, config });
      });
    });

    container
      .querySelector("#climateDistrict")
      ?.addEventListener("change", (event) => {
        container.dataset.district = event.target.value;
        render({ container, demandas, config });
      });

    container
      .querySelector("#climateViewMode")
      ?.addEventListener("change", (event) => {
        container.dataset.viewMode = event.target.value;
        render({ container, demandas, config });
      });

    container
      .querySelector("#climateMonth")
      ?.addEventListener("change", (event) => {
        container.dataset.month = event.target.value;
        render({ container, demandas, config });
      });

    container
      .querySelector("#climateWeekReference")
      ?.addEventListener("change", (event) => {
        container.dataset.weekReference = event.target.value;
        render({ container, demandas, config });
      });

    container
      .querySelector("#climatePeriodFilter")
      ?.addEventListener("change", (event) => {
        container.dataset.periodo = event.target.value;

        realWeatherCache = null;
        realWeatherPromise = null;
        realWeatherCacheKey = "";

        render({ container, demandas, config });
      });

    container
      .querySelector("#climateDateOriginButton")
      ?.addEventListener("click", () => {
        const multi = container.querySelector("#climateDateOriginMulti");
        multi?.classList.toggle("is-open");
      });

    container
      .querySelectorAll("[data-climate-date-origin]")
      .forEach((input) => {
        input.addEventListener("change", () => {
          const selected = Array.from(
            container.querySelectorAll("[data-climate-date-origin]:checked"),
          ).map((item) => item.dataset.climateDateOrigin);

          setSelectedDateOrigins(container, selected);
          render({ container, demandas, config });
        });
      });
  }
  global.CCEClimate = {
    render,
  };
})(window);

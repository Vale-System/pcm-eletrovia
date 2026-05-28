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

      if (!unique.has(key)) {
        unique.set(key, point);
      }
    });

    return Array.from(unique.values());
  }

  function forecastKeyForDistrict(distrito) {
    if (!distrito) return "";

    return String(distrito.codigo || distrito.sede || distrito.nome || "");
  }

  function normalizeRealWeatherResponse(data) {
    const map = new Map();

    if (!data?.ok || !Array.isArray(data.results)) {
      return map;
    }

    data.results.forEach((result) => {
      const point = result.point || {};
      const key = String(point.id || point.nome || point.distrito || "");

      if (!key) return;

      const dailyMap = new Map();

      (result.daily || []).forEach((day) => {
        if (!day.date) return;
        dailyMap.set(day.date, {
          riscoBase: day.riscoBase || "sem-dados",
          riscoLabel: day.riscoLabel || "Sem dados",
          chuvaMm: Number(day.precipitationMm || 0),
          probabilidade: Number(day.precipitationProbability || 0),
          ventoRajadaKmh: Number(day.windGustKmh || 0),
          horasChuva: Number(day.precipitationHours || 0),
          weatherCode: day.weatherCode,
          icone: day.icone || "☁️",
          fonte: "api",
        });
      });

      map.set(key, dailyMap);
    });

    return map;
  }

  async function loadRealWeatherForecast(config) {
    const apiConfig = config.weatherApi || {};

    if (!apiConfig.enabled || !apiConfig.workerUrl) {
      return new Map();
    }

    if (realWeatherCache) return realWeatherCache;
    if (realWeatherPromise) return realWeatherPromise;

    const points = buildWeatherApiPoints(config);

    if (!points.length) {
      realWeatherCache = new Map();
      return realWeatherCache;
    }

    realWeatherPromise = fetch(apiConfig.workerUrl, {
      method: "POST",
      cache: "no-store",
      headers: {
        "Content-Type": "application/json",
        Accept: "application/json",
      },
      body: JSON.stringify({
        forecastDays: apiConfig.forecastDays || 16,
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
      });

    return realWeatherPromise;
  }

  function realWeatherFor(dateText, distrito, realForecastMap) {
    const key = forecastKeyForDistrict(distrito);

    if (!key || !realForecastMap?.has(key)) return null;

    const dailyMap = realForecastMap.get(key);

    return dailyMap.get(dateText) || null;
  }

  function getWeather(dateText, distrito, coordenada, realForecastMap) {
    const real = realWeatherFor(dateText, distrito, realForecastMap);

    if (real) return real;

    return pseudoWeather(dateText, distrito, coordenada);
  }

  function pseudoWeather(dateText, distrito, coordenada) {
    if (!dateText || !distrito) {
      return {
        riscoBase: "sem-dados",
        riscoLabel: "Sem dados",
        chuvaMm: 0,
        probabilidade: 0,
        icone: "☁️",
      };
    }

    const seedText = `${dateText}-${distrito.codigo}-${coordenada?.km ?? "D"}`;

    let seed = 0;

    for (let i = 0; i < seedText.length; i++) {
      seed += seedText.charCodeAt(i) * (i + 1);
    }

    const chance = seed % 100;
    const chuvaMm = Number(((seed % 35) + (chance > 70 ? 8 : 0)).toFixed(1));

    if (chuvaMm >= 25 || chance >= 88) {
      return {
        riscoBase: "critico",
        riscoLabel: "Crítico",
        chuvaMm,
        probabilidade: Math.min(95, chance),
        icone: "⛈️",
      };
    }

    if (chuvaMm >= 12 || chance >= 70) {
      return {
        riscoBase: "alto",
        riscoLabel: "Alto",
        chuvaMm,
        probabilidade: Math.min(90, chance),
        icone: "🌧️",
      };
    }

    if (chuvaMm >= 4 || chance >= 45) {
      return {
        riscoBase: "atencao",
        riscoLabel: "Atenção",
        chuvaMm,
        probabilidade: Math.min(80, chance),
        icone: "🌦️",
      };
    }

    return {
      riscoBase: "favoravel",
      riscoLabel: "Favorável",
      chuvaMm,
      probabilidade: Math.max(5, chance),
      icone: "☀️",
    };
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

  function currentMonthText() {
    return new Date().toISOString().slice(0, 7);
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

  function renderCalendar(monthText, distrito, rows, realForecastMap) {
    const days = daysOfMonth(monthText);
    const byDate = new Map();

    rows.forEach((row) => {
      if (!byDate.has(row.dataProgramada)) byDate.set(row.dataProgramada, []);
      byDate.get(row.dataProgramada).push(row);
    });

    return `
    <div class="climate-calendar-grid">
      ${days
        .map((date) => {
          const items = byDate.get(date) || [];

          const weather = distrito
            ? getWeather(date, distrito, null, realForecastMap)
            : {
                riscoBase: "sem-dados",
                riscoLabel: "Sem dados",
                icone: "☁️",
              };

          const highestRisk = items.reduce(
            (max, item) => Math.max(max, item.risco.nivel || 0),
            0,
          );

          const riskClass =
            highestRisk >= 4
              ? "critico"
              : highestRisk === 3
                ? "alto"
                : highestRisk === 2
                  ? "atencao"
                  : weather.riscoBase;

          return `
            <div class="climate-day climate-risk-${riskClass}">
              <div class="climate-day-top">
                <strong>${date.slice(-2)}</strong>
                <span>${items[0]?.clima?.icone || weather.icone}</span>
              </div>
              <small>${items.length ? `${items.length} ativ.` : weather.riscoLabel}</small>
            </div>
          `;
        })
        .join("")}
    </div>
  `;
  }

  function sameMonth(dateText, monthText) {
    return String(dateText || "").slice(0, 7) === monthText;
  }

  async function render({ container, demandas, config }) {
    container.innerHTML = `
      <div class="climate-card">
        <div class="climate-empty">
          Carregando regras climáticas, distritos e coordenadas...
        </div>
      </div>
    `;

    const externalData = await loadExternalClimateData(config);
    const effectiveConfig = buildEffectiveConfig(config || {}, externalData);

    const realForecastMap = await loadRealWeatherForecast(effectiveConfig);

    const rows = buildClimateRows(demandas, effectiveConfig, realForecastMap);

    const month = container.dataset.month || currentMonthText();
    const selectedDistrictCode = container.dataset.district || "__ALL__";

    const selectedDistrict =
      selectedDistrictCode === "__ALL__"
        ? null
        : (effectiveConfig.distritos || []).find(
            (item) => item.codigo === selectedDistrictCode,
          ) || null;

    const monthRows = rows.filter((row) =>
      sameMonth(row.dataProgramada, month),
    );

    const filteredRows = monthRows.filter((row) => {
      if (!selectedDistrict) return true;
      return row.localizacao.distrito?.codigo === selectedDistrict.codigo;
    });

    const riskRows = filteredRows
      .filter((row) => row.risco.nivel >= 2)
      .sort((a, b) => {
        if (b.risco.nivel !== a.risco.nivel) {
          return b.risco.nivel - a.risco.nivel;
        }

        return String(a.dataProgramada).localeCompare(String(b.dataProgramada));
      });

    const criticalRows = riskRows.filter((row) => row.risco.nivel >= 4);
    const highRows = riskRows.filter((row) => row.risco.nivel >= 3);
    const unidentifiedRows = monthRows.filter(
      (row) => !row.localizacao.distrito,
    );
    const kmExactRows = filteredRows.filter((row) =>
      Number.isFinite(row.localizacao.km),
    );

    const worstDay = riskRows[0];

    container.innerHTML = `
      <div class="climate-page">
        <div class="climate-hero">
          <div>
            <span class="climate-eyebrow">Clima & Programação Operacional</span>
            <h2>Risco Climático da Eletrovia</h2>
            
          </div>

          <div class="climate-filters">
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
              <span>Mês</span>
              <input id="climateMonth" type="month" value="${escapeHtml(month)}" />
            </label>
          </div>
        </div>

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
              <small>${escapeHtml(month)}</small>
            </div>

            ${renderCalendar(month, selectedDistrict, filteredRows, realForecastMap)}

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
                      .slice(0, 100)
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
                              <small>${escapeHtml(row.origemDataProgramada)}</small>
                              <small>${escapeHtml(row.clima.icone)} ${
                                row.clima.chuvaMm
                              } mm • ${row.clima.probabilidade}%</small>
                            </td>

                            <td>
                              <span class="climate-risk-pill climate-risk-${
                                row.risco.classe
                              }">
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
                    ? `${formatDatePt(worstDay.dataProgramada)} — ${
                        worstDay.risco.label
                      }`
                    : "Sem risco relevante"
                }
              </strong>

              <p>
                ${
                  worstDay
                    ? `Atividade ${escapeHtml(
                        worstDay.demanda.ordem || worstDay.demanda.id,
                      )} em ${escapeHtml(
                        worstDay.localizacao.distrito?.nome ||
                          "distrito não identificado",
                      )}.`
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
      </div>
    `;

    container
      .querySelector("#climateDistrict")
      ?.addEventListener("change", (event) => {
        container.dataset.district = event.target.value;
        render({ container, demandas, config });
      });

    container
      .querySelector("#climateMonth")
      ?.addEventListener("change", (event) => {
        container.dataset.month = event.target.value;
        render({ container, demandas, config });
      });
  }

  global.CCEClimate = {
    render,
  };
})(window);

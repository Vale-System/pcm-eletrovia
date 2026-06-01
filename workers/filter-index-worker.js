(function setupFilterIndexWorker() {
  "use strict";

  function normalizeText(value) {
    return String(value || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toUpperCase()
      .trim();
  }

  function formatSapStatusFilter(value) {
    const text = String(value || "")
      .replace(/_/g, " ")
      .replace(/\s+/g, " ")
      .trim();

    if (!text) return "";
    return text.slice(0, 9).trim();
  }

  function dateText(value) {
    return String(value || "").trim().slice(0, 10);
  }

  function toDate(value) {
    const text = dateText(value);
    if (!text) return null;
    const parsed = new Date(`${text}T00:00:00`);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }

  function dueClassOf(demand) {
    const dueClasses = dueClassesOf(demand);
    return (
      dueClasses.find((item) =>
        ["Antecipado", "Fora do Prazo", "No Prazo", "Vencido", "Vence em 20d"].includes(item),
      ) || ""
    );
  }

  function hasSapStatusText(demand, token) {
    const statusSistema = normalizeText(demand.statusSistema);
    const statusUsuario = normalizeText(demand.statusUsuario);
    const alvo = normalizeText(token);
    return statusSistema.includes(alvo) || statusUsuario.includes(alvo);
  }

  function isCanceledBySap(demand) {
    return hasSapStatusText(demand, "CANC");
  }

  function hasSapSystemClosureStatus(demand) {
    return normalizeText(demand.statusSistema).includes("ENTE") ||
      normalizeText(demand.statusSistema).includes("ENCE");
  }

  function isRealizedBySapStatus(demand) {
    const statusSistema = normalizeText(demand.statusSistema);
    const statusUsuario = normalizeText(demand.statusUsuario);
    return (
      statusSistema.includes("ENTE") ||
      statusSistema.includes("ENCE") ||
      ((statusUsuario.includes("ENCR") ||
        statusUsuario.includes("ENTE") ||
        statusUsuario.includes("ENCE")) &&
        !statusSistema.includes("LIB"))
    );
  }

  function hasRealizedDate(demand) {
    return Boolean(dateText(demand.dataRealizada));
  }

  function hasUserWaitingClosureStatus(demand) {
    return normalizeText(demand.statusUsuario).includes("ENCR");
  }

  function isWaitingClosure(demand) {
    if (isCanceledBySap(demand)) return false;
    if (!hasRealizedDate(demand)) return false;
    const statusSistema = normalizeText(demand.statusSistema);
    const hasOpenSystemStatus =
      statusSistema.includes("LIB") || statusSistema.includes("CONF");
    return hasOpenSystemStatus && !hasSapSystemClosureStatus(demand);
  }

  function needsTechnicalClosure(demand) {
    if (isCanceledBySap(demand)) return false;
    if (hasSapSystemClosureStatus(demand)) return false;
    return hasRealizedDate(demand) || hasUserWaitingClosureStatus(demand);
  }

  function sanitizeSubstatusLabel(value) {
    const text = String(value || "").trim();
    if (!text || text === "Avaliar Status no SAP") return "";
    if (text === "Encerrado no SAP BO sem data realizada") {
      return "Ag Encerramento";
    }
    return text;
  }

  function normalizeSubstatusLabels(substatuses = []) {
    return Array.from(
      new Set((substatuses || []).map(sanitizeSubstatusLabel).filter(Boolean)),
    );
  }

  function isFutureControlId(value) {
    return /^ID-[^-]+-[^-]+-/i.test(String(value || "").trim());
  }

  function hasSemVinculoSubstatus(demand) {
    return demand?.semVinculoOperacional === true;
  }

  function dueClassesOf(demand) {
    if (isCanceledBySap(demand)) return [];

    const today = toDate(new Date().toISOString());
    const due = toDate(demand.vencimento);
    const min = toDate(demand.toleranciaMin);
    const max = toDate(demand.toleranciaMax);
    const realized = hasRealizedDate(demand) ? toDate(demand.dataRealizada) : null;
    const referenceLimit = max || due;

    if (realized) {
      if (min && realized < min) {
        return ["Antecipado"];
      } else if (referenceLimit && realized > referenceLimit) {
        return ["Fora do Prazo"];
      } else {
        return ["No Prazo"];
      }
    }

    if (!today || (!due && !referenceLimit)) return [];

    if (referenceLimit && today > referenceLimit) {
      return ["Vencido"];
    }

    if (due) {
      const diffDays = Math.round((due - today) / 86400000);
      if (diffDays >= 0 && diffDays <= 20) {
        return ["Vence em 20d"];
      }
    }

    return ["No Prazo"];
  }

  function primaryStatusOf(demand) {
    if (isCanceledBySap(demand)) return "Cancelado";
    if (isRealizedBySapStatus(demand)) return "Realizado";
    if (demand.dataReplanejadaAtual) return "Replanejado";
    if (demand.dataPlanejada) return "Planejado";
    return "A Planejar";
  }

  function substatusListOf(demand) {
    const status = primaryStatusOf(demand);
    if (status === "Cancelado") return ["Cancelado"];
    if (hasSemVinculoSubstatus(demand)) return ["Sem Vinculo"];
    if (
      (status === "Realizado" && !hasRealizedDate(demand)) ||
      needsTechnicalClosure(demand) ||
      (status === "Realizado" && isWaitingClosure(demand))
    ) {
      return ["Ag Encerramento"];
    }
    if (demand.perda) return ["Perda"];

    const dueClasses = normalizeSubstatusLabels(dueClassesOf(demand));
    if (dueClasses.length) return [dueClasses[0]];

    return [];
  }

  function monthName(month) {
    return (
      {
        "01": "Jan",
        "02": "Fev",
        "03": "Mar",
        "04": "Abr",
        "05": "Mai",
        "06": "Jun",
        "07": "Jul",
        "08": "Ago",
        "09": "Set",
        "10": "Out",
        "11": "Nov",
        "12": "Dez",
      }[month] || month
    );
  }

  function filterValueFor(item, definition) {
    if (definition.special === "status") {
      return primaryStatusOf(item);
    }
    if (definition.special === "centroStatus") {
      return item.centroTrabalhoStatus || "Nao cadastrado";
    }
    if (definition.special === "substatus") {
      return substatusListOf(item);
    }
    if (definition.special === "anoVencimento") {
      return dateText(item.vencimento).slice(0, 4) || "";
    }
    if (definition.special === "mesVencimento") {
      const month = dateText(item.vencimento).slice(5, 7) || "";
      return month ? `${month} - ${monthName(month)}` : "";
    }

    const value = item[definition.field] || "";
    if (definition.formatter === "sapStatus") {
      return formatSapStatusFilter(value);
    }
    return value;
  }

  self.onmessage = (event) => {
    const { demandas = [], definitions = [] } = event.data || {};
    const indexes = {};
    const options = {};

    definitions.forEach((definition) => {
      indexes[definition.key] = {};
      options[definition.key] = [];
    });

    demandas.forEach((item) => {
      definitions.forEach((definition) => {
        const raw = filterValueFor(item, definition);
        const values = Array.isArray(raw) ? raw : [raw];

        values
          .map((value) => String(value || "").trim())
          .filter(Boolean)
          .forEach((value) => {
            if (!indexes[definition.key][value]) {
              indexes[definition.key][value] = [];
              options[definition.key].push(value);
            }
            indexes[definition.key][value].push(item.id);
          });
      });
    });

    Object.keys(options).forEach((key) => {
      options[key].sort((a, b) => a.localeCompare(b, "pt-BR"));
    });

    self.postMessage({ indexes, options });
  };
})();

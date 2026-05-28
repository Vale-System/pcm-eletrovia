(function configurarJsonCacheService(global) {
  "use strict";

  const DEFAULT_TTL_MS = 5 * 60 * 1000;
  const cache = new Map();

  function now() {
    return Date.now();
  }

  function normalizarUrl(url) {
    return String(url || "").trim();
  }

  function montarUrlComCacheBust(url) {
    const separador = url.includes("?") ? "&" : "?";
    return `${url}${separador}_=${Date.now()}`;
  }

  async function getJson(url, options = {}) {
    const finalUrl = normalizarUrl(url);

    if (!finalUrl) {
      throw new Error("URL do JSON não informada.");
    }

    const ttlMs = Number(options.ttlMs ?? DEFAULT_TTL_MS);
    const force = options.force === true;
    const cacheBust = options.cacheBust === true;
    const cacheKey = finalUrl;
    const cached = cache.get(cacheKey);

    if (
      !force &&
      cached &&
      cached.expiresAt > now() &&
      cached.status === "resolved"
    ) {
      return cached.data;
    }

    if (!force && cached?.status === "pending") {
      return cached.promise;
    }

    const requestUrl = cacheBust ? montarUrlComCacheBust(finalUrl) : finalUrl;

    const promise = fetch(requestUrl, {
      method: "GET",
      cache: cacheBust ? "no-store" : "default",
      headers: {
        Accept: "application/json",
      },
    }).then(async (response) => {
      if (!response.ok) {
        throw new Error(
          `Falha ao carregar ${finalUrl}. HTTP ${response.status}`,
        );
      }

      const data = await response.json();

      cache.set(cacheKey, {
        status: "resolved",
        data,
        expiresAt: now() + ttlMs,
        loadedAt: new Date().toISOString(),
      });

      return data;
    });

    cache.set(cacheKey, {
      status: "pending",
      promise,
      expiresAt: now() + ttlMs,
      loadedAt: new Date().toISOString(),
    });

    try {
      return await promise;
    } catch (error) {
      cache.delete(cacheKey);
      throw error;
    }
  }

  async function getFirstAvailableJson(urls, options = {}) {
    const lista = Array.isArray(urls) ? urls : [urls];
    const errors = [];

    for (const url of lista) {
      try {
        return await getJson(url, options);
      } catch (error) {
        errors.push(`${url}: ${error.message}`);
      }
    }

    throw new Error(`Nenhum JSON disponível. ${errors.join(" | ")}`);
  }

  function setJson(url, data, options = {}) {
    const finalUrl = normalizarUrl(url);
    const ttlMs = Number(options.ttlMs ?? DEFAULT_TTL_MS);

    cache.set(finalUrl, {
      status: "resolved",
      data,
      expiresAt: now() + ttlMs,
      loadedAt: new Date().toISOString(),
    });

    return data;
  }

  function clear(url) {
    if (url) {
      cache.delete(normalizarUrl(url));
      return;
    }

    cache.clear();
  }

  function snapshot() {
    return Array.from(cache.entries()).map(([url, item]) => ({
      url,
      status: item.status,
      expiresAt: item.expiresAt,
      loadedAt: item.loadedAt,
    }));
  }

  global.CCEJsonCache = {
    getJson,
    getFirstAvailableJson,
    setJson,
    clear,
    snapshot,
  };
})(window);

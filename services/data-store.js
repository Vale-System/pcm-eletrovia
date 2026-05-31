(function configureDataStore(global) {
  "use strict";

  const META_STORAGE_KEY = "cce.data-meta";

  const runtime = {
    bases: new Map(),
    derived: new Map(),
    views: new Map(),
    indexes: new Map(),
  };

  function loadMeta() {
    try {
      return JSON.parse(global.localStorage.getItem(META_STORAGE_KEY) || "{}");
    } catch (error) {
      return {};
    }
  }

  function saveMeta(meta) {
    try {
      global.localStorage.setItem(META_STORAGE_KEY, JSON.stringify(meta));
    } catch (error) {
      console.warn("Nao foi possivel salvar metadata local:", error);
    }
  }

  function getBucket(bucket) {
    if (!runtime[bucket]) runtime[bucket] = new Map();
    return runtime[bucket];
  }

  function get(bucket, key) {
    return getBucket(bucket).get(key);
  }

  function set(bucket, key, value) {
    getBucket(bucket).set(key, value);
    return value;
  }

  function clear(bucket, key) {
    if (!bucket) {
      Object.values(runtime).forEach((item) => item.clear());
      return;
    }
    if (!key) {
      getBucket(bucket).clear();
      return;
    }
    getBucket(bucket).delete(key);
  }

  function has(bucket, key) {
    return getBucket(bucket).has(key);
  }

  global.CCEDataStore = {
    runtime,
    loadMeta,
    saveMeta,
    get,
    set,
    clear,
    has,
  };
})(window);

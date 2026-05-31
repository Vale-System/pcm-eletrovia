(function configureDataStore(global) {
  "use strict";

  const META_STORAGE_KEY = "cce.data-meta";
  const DB_NAME = "cce-data-store";
  const DB_VERSION = 1;
  const STORE_NAME = "entries";

  const runtime = {
    bases: new Map(),
    derived: new Map(),
    views: new Map(),
    indexes: new Map(),
  };

  let dbPromise = null;

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

  function openDb() {
    if (!global.indexedDB) {
      return Promise.resolve(null);
    }

    if (!dbPromise) {
      dbPromise = new Promise((resolve) => {
        const request = global.indexedDB.open(DB_NAME, DB_VERSION);

        request.onupgradeneeded = () => {
          const db = request.result;
          if (!db.objectStoreNames.contains(STORE_NAME)) {
            db.createObjectStore(STORE_NAME, { keyPath: "id" });
          }
        };

        request.onsuccess = () => resolve(request.result);
        request.onerror = () => {
          console.warn("Nao foi possivel abrir IndexedDB local:", request.error);
          resolve(null);
        };
      });
    }

    return dbPromise;
  }

  async function withStore(mode, executor) {
    const db = await openDb();
    if (!db) return null;

    return new Promise((resolve) => {
      const transaction = db.transaction(STORE_NAME, mode);
      const store = transaction.objectStore(STORE_NAME);

      executor(store, resolve);

      transaction.onerror = () => {
        console.warn("Falha ao acessar cache persistente:", transaction.error);
        resolve(null);
      };
    });
  }

  function buildPersistentId(bucket, key) {
    return `${bucket}::${key}`;
  }

  async function getPersistent(bucket, key) {
    const runtimeValue = get(bucket, key);
    if (runtimeValue !== undefined) {
      return runtimeValue;
    }

    const id = buildPersistentId(bucket, key);
    const entry = await withStore("readonly", (store, resolve) => {
      const request = store.get(id);
      request.onsuccess = () => resolve(request.result || null);
      request.onerror = () => resolve(null);
    });

    if (!entry || entry.value === undefined) return null;

    set(bucket, key, entry.value);
    return entry.value;
  }

  async function setPersistent(bucket, key, value) {
    set(bucket, key, value);

    const id = buildPersistentId(bucket, key);
    await withStore("readwrite", (store, resolve) => {
      const request = store.put({
        id,
        bucket,
        key,
        value,
        updatedAt: new Date().toISOString(),
      });
      request.onsuccess = () => resolve(true);
      request.onerror = () => resolve(false);
    });

    return value;
  }

  async function clearPersistent(bucket, key) {
    if (!bucket && !key) {
      await withStore("readwrite", (store, resolve) => {
        const request = store.clear();
        request.onsuccess = () => resolve(true);
        request.onerror = () => resolve(false);
      });
      clear();
      return;
    }

    if (bucket && !key) {
      const db = await openDb();
      if (!db) {
        clear(bucket);
        return;
      }

      await new Promise((resolve) => {
        const transaction = db.transaction(STORE_NAME, "readwrite");
        const store = transaction.objectStore(STORE_NAME);
        const cursorRequest = store.openCursor();

        cursorRequest.onsuccess = (event) => {
          const cursor = event.target.result;
          if (!cursor) {
            resolve(true);
            return;
          }

          if (cursor.value?.bucket === bucket) {
            cursor.delete();
          }
          cursor.continue();
        };

        cursorRequest.onerror = () => resolve(false);
      });

      clear(bucket);
      return;
    }

    clear(bucket, key);
    await withStore("readwrite", (store, resolve) => {
      const request = store.delete(buildPersistentId(bucket, key));
      request.onsuccess = () => resolve(true);
      request.onerror = () => resolve(false);
    });
  }

  global.CCEDataStore = {
    runtime,
    loadMeta,
    saveMeta,
    get,
    set,
    clear,
    has,
    openDb,
    getPersistent,
    setPersistent,
    clearPersistent,
  };
})(window);

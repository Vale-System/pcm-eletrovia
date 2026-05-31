(function configureIndexBuilder(global) {
  "use strict";

  function normalize(value) {
    return String(value || "").trim();
  }

  function buildDemandIndexes(demandas = [], definitions = []) {
    const indexes = {};
    const options = {};

    definitions.forEach((definition) => {
      indexes[definition.key] = new Map();
      options[definition.key] = [];
    });

    demandas.forEach((item) => {
      definitions.forEach((definition) => {
        const raw = definition.getValue(item);
        const values = Array.isArray(raw) ? raw : [raw];

        values
          .map(normalize)
          .filter(Boolean)
          .forEach((value) => {
            if (!indexes[definition.key].has(value)) {
              indexes[definition.key].set(value, []);
              options[definition.key].push(value);
            }
            indexes[definition.key].get(value).push(item.id);
          });
      });
    });

    Object.keys(options).forEach((key) => {
      options[key].sort((a, b) => a.localeCompare(b, "pt-BR"));
    });

    return {
      indexes,
      options,
    };
  }

  global.CCEIndexBuilder = {
    buildDemandIndexes,
  };
})(window);

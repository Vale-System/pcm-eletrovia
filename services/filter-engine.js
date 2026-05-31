(function configureFilterEngine(global) {
  "use strict";

  function uniqueSorted(values = []) {
    return Array.from(new Set(values.filter(Boolean))).sort((a, b) =>
      String(a).localeCompare(String(b), "pt-BR"),
    );
  }

  function collectOptionsFromRows(rows, definitions) {
    const options = {};

    definitions.forEach((definition) => {
      options[definition.key] = [];
    });

    rows.forEach((item) => {
      definitions.forEach((definition) => {
        const raw = definition.getValue(item);
        const values = Array.isArray(raw) ? raw : [raw];
        values
          .map((value) => String(value || "").trim())
          .filter(Boolean)
          .forEach((value) => {
            options[definition.key].push(value);
          });
      });
    });

    Object.keys(options).forEach((key) => {
      options[key] = uniqueSorted(options[key]);
    });

    return options;
  }

  global.CCEFilterEngine = {
    collectOptionsFromRows,
    uniqueSorted,
  };
})(window);

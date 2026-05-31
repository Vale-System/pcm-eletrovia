(function configurePerformanceLogger(global) {
  "use strict";

  const measures = new Map();

  function start(label) {
    measures.set(label, performance.now());
  }

  function end(label, options = {}) {
    if (!measures.has(label)) return 0;
    const duration = performance.now() - measures.get(label);
    measures.delete(label);

    if (options.log !== false) {
      console.info(
        `[perf] ${label}: ${duration.toFixed(1)}ms`,
        options.meta || "",
      );
    }

    return duration;
  }

  function measure(label, fn, options = {}) {
    start(label);
    try {
      return fn();
    } finally {
      end(label, options);
    }
  }

  async function measureAsync(label, fn, options = {}) {
    start(label);
    try {
      return await fn();
    } finally {
      end(label, options);
    }
  }

  global.CCEPerformance = {
    start,
    end,
    measure,
    measureAsync,
  };
})(window);

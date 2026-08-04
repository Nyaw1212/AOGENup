//----------------------------------
// Performance Timer
//----------------------------------

function perfStart_(label) {
  return {
    label: label,
    start: Date.now(),
    marks: {},
    last: Date.now(),
  };
}

function perfMark_(perf, name) {
  const now = Date.now();
  perf.marks[name] = now - perf.last;
  perf.last = now;
}

function perfEnd_(perf, extra) {
  const total = Date.now() - perf.start;

  const summary = Object.assign(
    {
      totalMs: total,
    },
    perf.marks,
    extra || {}
  );

  console.log(
    `[PERF][${perf.label}] END ` +
      Object.entries(summary)
        .map(([k, v]) => `${k}=${v}`)
        .join(" ")
  );

  return summary;
}
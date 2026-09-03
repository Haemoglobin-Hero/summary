// Keep the workbook's Transaction Description separate from Part Description.
// The source workbook contains both headers, so the generic description matcher
// must never see "Transaction Description" as a Part Description candidate.
(function () {
  function install() {
    if (!window.XLSX || !XLSX.utils || !XLSX.utils.sheet_to_json) return;
    if (XLSX.utils.sheet_to_json.__summaryHeaderFix) return;

    const original = XLSX.utils.sheet_to_json;
    const wrapped = function (sheet, opts) {
      const result = original.call(this, sheet, opts);
      if (!opts || opts.header !== 1 || !Array.isArray(result)) return result;

      return result.map(row => {
        if (!Array.isArray(row)) return row;
        return row.map(cell => {
          const value = String(cell ?? '').trim().replace(/\s+/g, ' ');
          return value.toLowerCase() === 'transaction description' ? 'Txn Description' : cell;
        });
      });
    };

    wrapped.__summaryHeaderFix = true;
    XLSX.utils.sheet_to_json = wrapped;
  }

  install();
  window.addEventListener('load', install, { once: true });
})();

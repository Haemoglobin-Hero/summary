// Force the exact source-workbook column mapping before app.js parses rows.
// Source layout (1-based): B=Part No, D=Transaction Code, E=Transaction Description,
// F=Quantity, J=Part Description, K=UoM. The summary must use J for Part Description.
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

        const normalized = row.map(cell => String(cell ?? '').trim().replace(/\s+/g, ' ').toLowerCase());
        const isSourceHeader =
          normalized[1] === 'part no' &&
          normalized[3] === 'transaction code' &&
          normalized[4] === 'transaction description' &&
          normalized[5] === 'quantity' &&
          normalized[9] === 'part description' &&
          normalized[10] === 'uom';

        if (isSourceHeader) {
          // Keep J4 as the ONLY description header visible to the app's matcher.
          const fixed = row.slice();
          fixed[4] = 'Txn Description';
          fixed[9] = 'Part Description';
          return fixed;
        }

        return row;
      });
    };

    wrapped.__summaryHeaderFix = true;
    XLSX.utils.sheet_to_json = wrapped;
  }

  install();
  window.addEventListener('load', install, { once: true });
})();

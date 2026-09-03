// Force the exact source-workbook column mapping before app.js parses rows.
// Source layout (1-based): B=Part No, D=Transaction Code, E=Transaction Description,
// F=Quantity, J=Part Description, K=UoM.
// The source data begins at Excel row 4, so row 4 is preserved as the first data row.
(function () {
  function install() {
    if (!window.XLSX || !XLSX.utils || !XLSX.utils.sheet_to_json) return;
    if (XLSX.utils.sheet_to_json.__summaryHeaderFix) return;

    const original = XLSX.utils.sheet_to_json;
    const wrapped = function (sheet, opts) {
      const result = original.call(this, sheet, opts);
      if (!opts || opts.header !== 1 || !Array.isArray(result)) return result;

      // The summary parser previously guessed columns by matching any header text.
      // That can select E (Transaction Description) instead of J (Part Description).
      // Replace the parsed rows with a deterministic header + the actual data from
      // Excel row 4 onward. This makes B/D/F/J/K the only source of the required fields.
      const canonicalHeader = [];
      canonicalHeader[1] = 'Part No';
      canonicalHeader[3] = 'Transaction Code';
      canonicalHeader[4] = 'Txn Description';
      canonicalHeader[5] = 'Quantity';
      canonicalHeader[9] = 'Part Description';
      canonicalHeader[10] = 'UoM';

      // Excel row 4 is zero-based array index 3.
      const dataRows = result.slice(3).map(row => Array.isArray(row) ? row.slice() : row);
      return [canonicalHeader, ...dataRows];
    };

    wrapped.__summaryHeaderFix = true;
    XLSX.utils.sheet_to_json = wrapped;
  }

  install();
  window.addEventListener('load', install, { once: true });
})();

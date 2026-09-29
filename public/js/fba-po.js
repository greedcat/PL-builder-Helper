  const fbaInput   = document.getElementById('fbaInput');
  const poInput    = document.getElementById('poInput');
  const output     = document.getElementById('output');
  const outputSec  = document.getElementById('outputSection');
  const lineCount  = document.getElementById('lineCount');
  const errorMsg   = document.getElementById('errorMsg');
  const btnCopy    = document.getElementById('btnCopy');

  function parseLines(text) {
    return text.split('\n').map(l => l.trim()).filter(l => l.length > 0);
  }

  document.getElementById('btnGenerate').addEventListener('click', () => {
    errorMsg.style.display = 'none';

    const fbas = parseLines(fbaInput.value);
    const pos  = parseLines(poInput.value);

    if (!fbas.length && !pos.length) {
      errorMsg.textContent = 'Please enter at least one FBA or PO value.';
      errorMsg.style.display = 'block';
      outputSec.classList.remove('visible');
      return;
    }

    const rawCount = Math.max(fbas.length, pos.length);

    // Deduplicate: keep first occurrence of each FBA+PO pair
    const seen = new Set();
    const pairs = [];
    for (let i = 0; i < rawCount; i++) {
      const fba = fbas[i] ?? '';
      const po  = pos[i]  ?? '';
      const key = `${fba}||${po}`;
      if (!seen.has(key)) {
        seen.add(key);
        pairs.push({ fba, po });
      }
    }

    const count = pairs.length;
    output.innerHTML = '';

    for (let i = 0; i < count; i++) {
      const { fba, po } = pairs[i];

      const row = document.createElement('div');
      row.className = 'line-item';
      row.innerHTML =
        `<span class="line-num">${i + 1}.</span>` +
        `<span class="line-fba">${escapeHtml(fba)}</span>` +
        (fba && po ? `<span class="line-sep"> / </span>` : '') +
        `<span class="line-po">${escapeHtml(po)}</span>`;
      output.appendChild(row);
    }

    const skipped = rawCount - count;
    lineCount.textContent = `${count} line${count !== 1 ? 's' : ''}${skipped > 0 ? ` (${skipped} duplicate${skipped !== 1 ? 's' : ''} removed)` : ''}`;
    outputSec.classList.add('visible');

    // Export to output.xlsx
    const headers = [['ARN','PRO', 'BOL List (use , as separator)','Vendor Name','Pallet Count','Carton Count','Unit Count','PO List (use , as separator) *']];
    const dataRows = [];
    for (let i = 0; i < count; i++) {
      dataRows.push(['', pairs[i].fba, '', '', i === 0 ? 1 : '', '', '', pairs[i].po]);
    }
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([...headers, ...dataRows]);
    ws['!cols'] = [{ wch: 14 }, { wch: 24 }, { wch: 30 }, { wch: 20 }, { wch: 13 }, { wch: 13 }, { wch: 11 }, { wch: 32 }];
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, 'output.xlsx');
  });

  document.getElementById('btnClear').addEventListener('click', () => {
    fbaInput.value = '';
    poInput.value  = '';
    output.innerHTML = '';
    outputSec.classList.remove('visible');
    errorMsg.style.display = 'none';
    btnCopy.textContent = 'Copy';
    btnCopy.classList.remove('copied');
  });

  btnCopy.addEventListener('click', () => {
    const fbas = parseLines(fbaInput.value);
    const pos  = parseLines(poInput.value);
    if (!fbas.length && !pos.length) return;

    const count = Math.max(fbas.length, pos.length);
    const lines = [];
    for (let i = 0; i < count; i++) {
      const fba = fbas[i] ?? '';
      const po  = pos[i]  ?? '';
      if (fba && po) lines.push(`${fba} / ${po}`);
      else           lines.push(fba || po);
    }

    navigator.clipboard.writeText(lines.join('\n')).then(() => {
      btnCopy.textContent = 'Copied!';
      btnCopy.classList.add('copied');
      setTimeout(() => {
        btnCopy.textContent = 'Copy';
        btnCopy.classList.remove('copied');
      }, 2000);
    });
  });

  function escapeHtml(str) {
    return str.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  }

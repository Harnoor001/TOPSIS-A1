const fileInput = document.getElementById('fileInput');
const weightsInput = document.getElementById('weightsInput');
const impactsInput = document.getElementById('impactsInput');
const analyzeBtn = document.getElementById('analyzeBtn');
const message = document.getElementById('message');
const resultSection = document.getElementById('resultSection');
const resultTable = document.getElementById('resultTable');
const downloadBtn = document.getElementById('downloadBtn');

let latestResult = [];

analyzeBtn.addEventListener('click', async () => {
  clearMessage();
  try {
    const file = fileInput.files[0];
    if (!file) throw new Error('Please select a CSV/XLSX file.');

    const rows = await parseFile(file);
    const data = validateAndPrepare(rows);
    const weights = parseWeights(weightsInput.value, data.criteriaCount);
    const impacts = parseImpacts(impactsInput.value, data.criteriaCount);

    const result = runTopsis(data.names, data.matrix, weights, impacts);
    latestResult = result;
    renderTable(result);
    resultSection.hidden = false;
    showMessage('Analysis completed successfully.', false);
  } catch (err) {
    showMessage(err.message, true);
  }
});

downloadBtn.addEventListener('click', () => {
  if (!latestResult.length) return;
  const csv = Papa.unparse(latestResult);
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'topsis-result.csv';
  a.click();
  URL.revokeObjectURL(url);
});

function parseFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  if (ext === 'csv') {
    return new Promise((resolve, reject) => {
      Papa.parse(file, {
        complete: (res) => resolve(res.data),
        error: () => reject(new Error('Unable to parse CSV file.')),
      });
    });
  }

  if (['xlsx', 'xls'].includes(ext)) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const workbook = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
          const sheet = workbook.Sheets[workbook.SheetNames[0]];
          resolve(XLSX.utils.sheet_to_json(sheet, { header: 1 }));
        } catch {
          reject(new Error('Unable to parse Excel file.'));
        }
      };
      reader.onerror = () => reject(new Error('Unable to read selected file.'));
      reader.readAsArrayBuffer(file);
    });
  }

  throw new Error('Unsupported file format. Use .csv, .xls, or .xlsx.');
}

function validateAndPrepare(rows) {
  const cleaned = rows.filter((r) => r && r.length && r.some((v) => `${v}`.trim() !== ''));
  if (cleaned.length < 2) throw new Error('File must include header row and at least one data row.');

  const header = cleaned[0];
  if (header.length < 3) throw new Error('Dataset must have at least 3 columns.');

  const names = [];
  const matrix = [];

  for (let i = 1; i < cleaned.length; i++) {
    const row = cleaned[i];
    names.push(String(row[0] ?? `Option ${i}`));
    const vals = row.slice(1, header.length).map((v) => Number(v));

    if (vals.some((v) => Number.isNaN(v))) {
      throw new Error(`Non-numeric value detected in row ${i + 1}.`);
    }
    matrix.push(vals);
  }

  return { names, matrix, criteriaCount: header.length - 1, header };
}

function parseWeights(input, expected) {
  const weights = input.split(',').map((n) => Number(n.trim())).filter((n) => !Number.isNaN(n));
  if (weights.length !== expected) throw new Error(`Please provide exactly ${expected} weights.`);
  return weights;
}

function parseImpacts(input, expected) {
  const impacts = input.split(',').map((s) => s.trim());
  if (impacts.length !== expected || impacts.some((s) => s !== '+' && s !== '-')) {
    throw new Error(`Impacts must include exactly ${expected} values using only + or -.`);
  }
  return impacts;
}

function runTopsis(names, matrix, weights, impacts) {
  const rows = matrix.length;
  const cols = matrix[0].length;

  const denom = Array(cols).fill(0);
  for (let j = 0; j < cols; j++) {
    denom[j] = Math.sqrt(matrix.reduce((sum, row) => sum + row[j] ** 2, 0));
  }

  const weighted = matrix.map((row) =>
    row.map((v, j) => (v / denom[j]) * weights[j])
  );

  const idealBest = [];
  const idealWorst = [];

  for (let j = 0; j < cols; j++) {
    const colValues = weighted.map((row) => row[j]);
    const max = Math.max(...colValues);
    const min = Math.min(...colValues);
    idealBest.push(impacts[j] === '+' ? max : min);
    idealWorst.push(impacts[j] === '+' ? min : max);
  }

  const scoreRows = weighted.map((row, i) => {
    const dPlus = Math.sqrt(row.reduce((sum, v, j) => sum + (v - idealBest[j]) ** 2, 0));
    const dMinus = Math.sqrt(row.reduce((sum, v, j) => sum + (v - idealWorst[j]) ** 2, 0));
    const score = dMinus / (dPlus + dMinus);
    return { Alternative: names[i], 'TOPSIS Score': Number(score.toFixed(6)) };
  });

  const sorted = [...scoreRows].sort((a, b) => b['TOPSIS Score'] - a['TOPSIS Score']);
  const rankMap = new Map(sorted.map((item, idx) => [item.Alternative, idx + 1]));

  return scoreRows.map((item) => ({
    ...item,
    Rank: rankMap.get(item.Alternative),
  }));
}

function renderTable(data) {
  if (!data.length) return;
  const columns = Object.keys(data[0]);
  resultTable.innerHTML = `
    <thead><tr>${columns.map((c) => `<th>${c}</th>`).join('')}</tr></thead>
    <tbody>${data
      .map((row) => `<tr>${columns.map((c) => `<td>${row[c]}</td>`).join('')}</tr>`)
      .join('')}</tbody>
  `;
}

function showMessage(text, isError) {
  message.textContent = text;
  message.className = `message ${isError ? 'error' : 'success'}`;
}

function clearMessage() {
  message.textContent = '';
  message.className = 'message';
}

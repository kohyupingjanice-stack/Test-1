const SPREADSHEET_URL =
  'https://raw.githubusercontent.com/kohyupingjanice-stack/Test-1/main/2026%20TG%20allocation_Sem%201_9%20Jan%20(2).xlsx';

const searchInput = document.getElementById('searchInput');
const statusEl = document.getElementById('status');
const resultsCountEl = document.getElementById('resultsCount');
const resultsTable = document.getElementById('resultsTable');
const tableHead = document.getElementById('tableHead');
const tableBody = document.getElementById('tableBody');

const normalize = (value) => String(value ?? '').trim().toLowerCase();

function isStudentNameKey(key) {
  const text = normalize(key);
  if (!text) return false;
  if (/subject/.test(text)) return false;
  if (/teacher|class|grade|group|year|course|code|id|email|phone|status|gender/.test(text)) {
    return false;
  }
  return /student|name|learner|pupil|candidate/.test(text);
}

function getStudentNameKey(row) {
  if (!row || typeof row !== 'object') return '';
  const keys = Object.keys(row);
  const exact = keys.find((key) => isStudentNameKey(key));
  if (exact) return exact;
  const fallback = keys.find((key) => /name/.test(normalize(key)));
  if (fallback) return fallback;
  return keys[0] || '';
}

function renderRows(rows, studentNameKey) {
  const columns = Object.keys(rows[0] || {}).filter((key) => key !== '_studentName');
  const visibleColumns = columns.filter((key) => key !== studentNameKey).slice(0, 6);

  tableHead.innerHTML = `
    <tr>
      <th>${studentNameKey || 'Student Name'}</th>
      ${visibleColumns.map((key) => `<th>${key}</th>`).join('')}
    </tr>
  `;

  tableBody.innerHTML = rows
    .map((row, index) => {
      const studentName = row[studentNameKey] ?? row._studentName ?? 'N/A';
      const cells = visibleColumns
        .map((key) => `<td>${row[key] ?? '—'}</td>`)
        .join('');
      return `
        <tr>
          <td>${studentName}</td>
          ${cells}
        </tr>
      `;
    })
    .join('');

  resultsTable.hidden = false;
}

async function loadSpreadsheet() {
  try {
    const response = await fetch(SPREADSHEET_URL);
    if (!response.ok) {
      throw new Error('Spreadsheet could not be loaded.');
    }

    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });

    const allRows = [];
    let detectedStudentKey = '';

    workbook.SheetNames.forEach((sheetName) => {
      const sheet = workbook.Sheets[sheetName];
      const jsonRows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });

      jsonRows.forEach((row) => {
        if (!row || typeof row !== 'object') return;
        const key = getStudentNameKey(row);
        if (!detectedStudentKey && key) detectedStudentKey = key;
        const cleanRow = { ...row };
        cleanRow._studentName = cleanRow[key] ?? '';
        allRows.push(cleanRow);
      });
    });

    let rows = allRows;
    const filterRows = () => {
      const query = normalize(searchInput.value);
      if (!query) {
        renderRows(rows, detectedStudentKey || 'Student Name');
        resultsCountEl.hidden = false;
        resultsCountEl.textContent = `${rows.length} student${rows.length === 1 ? '' : 's'} loaded`;
        return;
      }

      const filtered = rows.filter((row) => {
        const value = row[detectedStudentKey || 'Student Name'] ?? row._studentName ?? '';
        return normalize(value).includes(query);
      });

      renderRows(filtered, detectedStudentKey || 'Student Name');
      resultsCountEl.hidden = false;
      resultsCountEl.textContent = `${filtered.length} matching student${filtered.length === 1 ? '' : 's'}`;

      if (filtered.length === 0) {
        statusEl.textContent = 'No student matches your search.';
        statusEl.classList.add('empty');
      } else {
        statusEl.textContent = 'Search complete.';
        statusEl.classList.remove('empty');
      }
    };

    searchInput.addEventListener('input', filterRows);
    filterRows();
    statusEl.textContent = 'Ready to search.';
    statusEl.classList.remove('error');
  } catch (error) {
    console.error(error);
    statusEl.textContent = 'Unable to load the spreadsheet data. Please check the file link and try again.';
    statusEl.classList.add('error');
  }
}

loadSpreadsheet();

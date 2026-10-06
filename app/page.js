'use client';

import { useEffect, useMemo, useState } from 'react';
import * as XLSX from 'xlsx';

const SPREADSHEET_URL =
  'https://raw.githubusercontent.com/kohyupingjanice-stack/Test-1/main/2026%20TG%20allocation_Sem%201_9%20Jan%20(2).xlsx';

const normalize = (value) =>
  String(value ?? '')
    .trim()
    .toLowerCase();

const isStudentNameKey = (key) => {
  const text = normalize(key);
  if (!text) return false;
  if (/subject/.test(text)) return false;
  if (/teacher|class|grade|group|year|course|code|id|email|phone|status|gender/.test(text)) {
    return false;
  }
  return /student|name|learner|pupil|candidate/.test(text);
};

const getStudentNameKey = (row) => {
  if (!row || typeof row !== 'object') return '';
  const keys = Object.keys(row);

  const exactMatch = keys.find((key) => isStudentNameKey(key));
  if (exactMatch) return exactMatch;

  const fallback = keys.find((key) => /name/.test(normalize(key)));
  if (fallback) return fallback;

  return keys[0] || '';
};

export default function Home() {
  const [rows, setRows] = useState([]);
  const [search, setSearch] = useState('');
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [studentNameKey, setStudentNameKey] = useState('');

  useEffect(() => {
    const loadSpreadsheet = async () => {
      try {
        setLoading(true);
        const response = await fetch(SPREADSHEET_URL);

        if (!response.ok) {
          throw new Error('Spreadsheet could not be loaded.');
        }

        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });

        const extractedRows = [];
        let detectedKey = '';

        workbook.SheetNames.forEach((sheetName) => {
          const worksheet = workbook.Sheets[sheetName];
          const jsonRows = XLSX.utils.sheet_to_json(worksheet, {
            defval: '',
            raw: false,
          });

          jsonRows.forEach((row) => {
            if (!row || typeof row !== 'object') return;

            const key = getStudentNameKey(row);
            if (!detectedKey && key) detectedKey = key;

            const cleanRow = { ...row };
            const nameValue = cleanRow[key] ?? '';
            cleanRow._studentName = nameValue;
            extractedRows.push(cleanRow);
          });
        });

        setStudentNameKey(detectedKey || 'Name');
        setRows(extractedRows);
        setError('');
      } catch (loadError) {
        console.error(loadError);
        setError('Unable to read the spreadsheet data. Please check the file link and try again.');
      } finally {
        setLoading(false);
      }
    };

    loadSpreadsheet();
  }, []);

  const visibleRows = useMemo(() => {
    const query = normalize(search);
    if (!query) return rows;

    return rows.filter((row) => {
      const studentName = row[studentNameKey] ?? row._studentName ?? '';
      return normalize(studentName).includes(query);
    });
  }, [rows, search, studentNameKey]);

  const tableColumns = useMemo(() => {
    const firstRow = rows[0] ?? {};
    return Object.keys(firstRow).filter((key) => key !== '_studentName');
  }, [rows]);

  return (
    <main className="page-shell">
      <section className="card">
        <p className="eyebrow">Student Directory</p>
        <h1>Search Students by Name</h1>

        <div className="search-wrap">
          <input
            type="text"
            value={search}
            onChange={(event) => setSearch(event.target.value)}
            placeholder="Type a student name..."
            aria-label="Search by student name"
          />
        </div>

        {loading ? (
          <p className="status">Loading student records...</p>
        ) : error ? (
          <p className="status error">{error}</p>
        ) : (
          <>
            <p className="results-count">
              {visibleRows.length} matching student{visibleRows.length === 1 ? '' : 's'}
            </p>

            {visibleRows.length === 0 ? (
              <p className="status empty">No student matches your search.</p>
            ) : (
              <div className="table-wrap">
                <table>
                  <thead>
                    <tr>
                      <th>{studentNameKey || 'Student Name'}</th>
                      {tableColumns
                        .filter((key) => key !== studentNameKey)
                        .slice(0, 6)
                        .map((key) => (
                          <th key={key}>{key}</th>
                        ))}
                    </tr>
                  </thead>
                  <tbody>
                    {visibleRows.map((row, index) => (
                      <tr key={`${row._studentName || 'student'}-${index}`}>
                        <td>{row[studentNameKey] ?? row._studentName ?? 'N/A'}</td>
                        {tableColumns
                          .filter((key) => key !== studentNameKey)
                          .slice(0, 6)
                          .map((key) => (
                            <td key={`${key}-${index}`}>{row[key] ?? '—'}</td>
                          ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </>
        )}
      </section>
    </main>
  );
}

/**
 * Compares the "DEZEMBRO 25" and "Testes" sheets ROW BY ROW
 * Shows exactly where values are in different columns
 */

import 'dotenv/config';
import path from 'path';
import { fileURLToPath } from 'url';
import ExcelJS from 'exceljs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const VALIDATED_SHEET = 'DEZEMBRO 25';
const TEST_SHEET = process.env.EXCEL_ABA || 'Testes';
const INITIAL_ROW = 8;

// All value columns (procedures + payments) - actual spreadsheet names
const COLUMNS = {
  D: 'AVALIACAO', E: 'TRATAMENTO', F: 'DR_RIGATTI', G: 'LIVRO', H: 'DR_VICTOR',
  I: 'IMPLANTE', J: 'SORO_DR', K: 'TIRZEPATIDA', L: 'APLICACOES', M: 'ONLINE',
  N: 'COMISSAO', O: 'NUTRIS', P: 'EXTRA', Q: 'ESTORNO_PROC',
  R: 'PGTO_OUTROS', S: 'DINHEIRO', T: 'SICOOB', U: 'SAFRA', V: 'PIX_SAFRA',
  W: 'INFINITE', X: 'PIX', Y: 'CHEQUE', Z: 'BOLETO', AB: 'JUROS_CARTAO',
};

function getValue(ws, row, column) {
  const cell = ws.getCell(`${column}${row}`);
  const val = cell.value;
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'object' && val.result !== undefined) return Number(val.result) || 0;
  return Number(val) || 0;
}

function getName(ws, row) {
  const val = ws.getCell(`B${row}`).value;
  if (!val) return null;
  return String(val).trim();
}

function getDate(ws, row) {
  const val = ws.getCell(`A${row}`).value;
  if (!val) return null;
  if (val instanceof Date) {
    const d = val;
    return `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`;
  }
  return String(val).trim();
}

function normalizeName(name) {
  return name.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
}

/**
 * Reads all rows from a sheet as array (preserving multiple entries per patient)
 */
function readRows(ws) {
  const rows = [];
  let row = INITIAL_ROW;
  while (row <= ws.rowCount) {
    const name = getName(ws, row);
    if (!name) { row++; continue; }
    const date = getDate(ws, row);
    const values = {};
    for (const col of Object.keys(COLUMNS)) {
      const v = getValue(ws, row, col);
      if (v !== 0) values[col] = v;
    }
    rows.push({ row, name, normalizedName: normalizeName(name), date, values });
    row++;
  }
  return rows;
}

/**
 * Groups rows by patient+date (sums values from duplicate rows)
 */
function groupRows(rows) {
  const groups = new Map();
  for (const r of rows) {
    const key = `${r.normalizedName}__${r.date}`;
    if (groups.has(key)) {
      const g = groups.get(key);
      for (const [col, val] of Object.entries(r.values)) {
        g.values[col] = (g.values[col] || 0) + val;
      }
      g.originalRows.push(r.row);
    } else {
      groups.set(key, {
        name: r.name,
        normalizedName: r.normalizedName,
        date: r.date,
        values: { ...r.values },
        originalRows: [r.row],
      });
    }
  }
  return groups;
}

async function main() {
  const filePath = path.resolve(__dirname, '../onedrive', process.env.EXCEL_FILE_PATH);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const wsVal = workbook.getWorksheet(VALIDATED_SHEET);
  const wsTest = workbook.getWorksheet(TEST_SHEET);

  if (!wsVal || !wsTest) {
    console.error('Sheet not found!');
    process.exit(1);
  }

  const validatedRows = readRows(wsVal);
  const testRows = readRows(wsTest);

  console.log(`Sheet "${VALIDATED_SHEET}": ${validatedRows.length} rows`);
  console.log(`Sheet "${TEST_SHEET}": ${testRows.length} rows\n`);

  // Group validated rows (may have multiple rows per patient+date)
  const validatedGroups = groupRows(validatedRows);
  const testGroups = groupRows(testRows);

  console.log(`Validated groups: ${validatedGroups.size} | Test groups: ${testGroups.size}\n`);

  let testRowCount = 0;
  let rowsOk = 0;
  let rowsWithDiff = 0;
  const divergences = [];

  // Compare each test group with validated
  for (const [key, test] of testGroups) {
    testRowCount++;

    // Try to find match in validated
    let valMatch = validatedGroups.get(key);
    if (!valMatch) {
      // Try by name only
      for (const [vKey, vGroup] of validatedGroups) {
        if (vGroup.normalizedName === test.normalizedName) {
          valMatch = vGroup;
          break;
        }
      }
    }

    if (!valMatch) {
      divergences.push({
        row: test.originalRows[0],
        patient: test.name,
        date: test.date,
        type: 'ONLY IN TEST',
        details: Object.entries(test.values)
          .map(([c, v]) => `${COLUMNS[c]}(${c})=${v.toFixed(2)}`)
          .join(', '),
      });
      rowsWithDiff++;
      continue;
    }

    // Compare values column by column
    const allColumns = new Set([...Object.keys(test.values), ...Object.keys(valMatch.values)]);
    const diffs = [];

    for (const col of allColumns) {
      const vTest = test.values[col] || 0;
      const vVal = valMatch.values[col] || 0;

      if (Math.abs(vTest - vVal) > 0.01) {
        diffs.push({
          column: col,
          columnName: COLUMNS[col],
          validated: vVal,
          test: vTest,
        });
      }
    }

    if (diffs.length === 0) {
      rowsOk++;
    } else {
      rowsWithDiff++;
      const details = diffs.map(d => {
        if (d.validated === 0) return `${d.columnName}(${d.column}): TEST has ${d.test.toFixed(2)} / VALIDATED does not`;
        if (d.test === 0) return `${d.columnName}(${d.column}): VALIDATED has ${d.validated.toFixed(2)} / TEST does not`;
        return `${d.columnName}(${d.column}): VAL=${d.validated.toFixed(2)} vs TEST=${d.test.toFixed(2)}`;
      });

      divergences.push({
        row: test.originalRows[0],
        patient: test.name,
        date: test.date,
        type: 'DIFFERENCE',
        details: details.join('\n          '),
      });
    }
  }

  // Check patients that are ONLY in validated
  for (const [key, val] of validatedGroups) {
    let testMatch = testGroups.get(key);
    if (!testMatch) {
      for (const [tKey, tGroup] of testGroups) {
        if (tGroup.normalizedName === val.normalizedName) {
          testMatch = tGroup;
          break;
        }
      }
    }
    if (!testMatch) {
      const hasValues = Object.keys(val.values).length > 0;
      if (hasValues) {
        divergences.push({
          row: val.originalRows[0],
          patient: val.name,
          date: val.date,
          type: 'ONLY IN VALIDATED',
          details: Object.entries(val.values)
            .map(([c, v]) => `${COLUMNS[c]}(${c})=${v.toFixed(2)}`)
            .join(', '),
        });
        rowsWithDiff++;
      }
    }
  }

  // Result
  console.log('='.repeat(80));
  console.log(`  SUMMARY: ${rowsOk} rows OK | ${rowsWithDiff} rows with differences`);
  console.log('='.repeat(80));

  if (divergences.length > 0) {
    console.log('\n');
    for (const d of divergences) {
      console.log(`--- Row ${d.row} | ${d.patient} | ${d.date} | [${d.type}]`);
      console.log(`          ${d.details}`);
      console.log('');
    }
  }
}

main().catch(console.error);

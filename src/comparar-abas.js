/**
 * Script to compare the validated sheet (DEZEMBRO 25) with the test sheet
 * Compares column by column, showing differences per patient
 *
 * Usage: node src/comparar-abas.js [column]
 * Example: node src/comparar-abas.js D
 * If no column specified, compares all
 */

import 'dotenv/config';
import path from 'path';
import { fileURLToPath } from 'url';
import ExcelJS from 'exceljs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const VALIDATED_SHEET = 'DEZEMBRO 25';
const TEST_SHEET = process.env.EXCEL_ABA || 'Testes';
const INITIAL_ROW = 8; // Data starts at row 8

// Procedure/value columns to compare
const COLUMNS_TO_COMPARE = {
  D: 'AVALIACAO',
  E: 'TRATAMENTO',
  F: 'DR_RIGATTI',
  G: 'LIVRO',
  H: 'DR_VICTOR',
  I: 'IMPLANTE',
  J: 'SORO_DR',
  K: 'TIRZEPATIDA',
  L: 'APLICACOES',
  M: 'ONLINE',
  N: 'COMISSAO',
  O: 'NUTRIS',
  P: 'EXTRA',
  Q: 'ESTORNO_PROC',
  S: 'DINHEIRO',
  T: 'CIELO',
  U: 'SAFRA',
  V: 'PIX',
  W: 'CHEQUE',
  X: 'BOLETO',
  Y: 'SICOOB',
  Z: 'JUROS_CARTAO',
};

function getValue(worksheet, row, column) {
  const cell = worksheet.getCell(`${column}${row}`);
  const val = cell.value;
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'object' && val.result !== undefined) return Number(val.result) || 0;
  return Number(val) || 0;
}

function getName(worksheet, row) {
  const cell = worksheet.getCell(`B${row}`);
  const val = cell.value;
  if (!val) return null;
  return String(val).trim();
}

function getDate(worksheet, row) {
  const cell = worksheet.getCell(`A${row}`);
  const val = cell.value;
  if (!val) return null;
  // Normalize dates - can be Date object or string DD/MM/YYYY
  if (val instanceof Date) {
    const d = val;
    const day = String(d.getDate()).padStart(2, '0');
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const year = d.getFullYear();
    return `${day}/${month}/${year}`;
  }
  return String(val).trim();
}

/**
 * Reads all data from a sheet, indexed by "name_date"
 * Sums values when there are multiple rows for the same patient+date
 */
function readSheetData(worksheet) {
  const data = new Map();
  let row = INITIAL_ROW;

  while (row <= worksheet.rowCount) {
    const name = getName(worksheet, row);
    if (!name) {
      row++;
      continue;
    }

    const date = getDate(worksheet, row);
    const normalizedName = normalizeName(name);
    const key = `${normalizedName}__${date}`;

    const values = {};
    for (const column of Object.keys(COLUMNS_TO_COMPARE)) {
      values[column] = getValue(worksheet, row, column);
    }

    if (data.has(key)) {
      // Sum values from duplicate rows (same patient + date)
      const existing = data.get(key);
      for (const column of Object.keys(COLUMNS_TO_COMPARE)) {
        existing.values[column] = (existing.values[column] || 0) + (values[column] || 0);
      }
    } else {
      data.set(key, { name, date, row, values });
    }

    row++;
  }

  return data;
}

function normalizeName(name) {
  return name.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
}

/**
 * Tries to find the best match for a patient
 */
function findMatch(key, sourceData, targetData) {
  // Key is already normalized (normalizedName__date), try direct match
  if (targetData.has(key)) return key;

  // Try by name only (without date) - patient may have different date
  const { name } = sourceData.get(key);
  const normalizedName = normalizeName(name);

  for (const [otherKey, otherEntry] of targetData) {
    if (normalizeName(otherEntry.name) === normalizedName) {
      return otherKey;
    }
  }

  return null;
}

async function main() {
  const columnFilter = process.argv[2]?.toUpperCase() || null;

  const filePath = path.resolve(__dirname, '../onedrive', process.env.EXCEL_FILE_PATH);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const wsValidated = workbook.getWorksheet(VALIDATED_SHEET);
  const wsTest = workbook.getWorksheet(TEST_SHEET);

  if (!wsValidated) {
    console.error(`Sheet "${VALIDATED_SHEET}" not found!`);
    process.exit(1);
  }
  if (!wsTest) {
    console.error(`Sheet "${TEST_SHEET}" not found!`);
    process.exit(1);
  }

  console.log(`\nComparing: "${VALIDATED_SHEET}" vs "${TEST_SHEET}"\n`);

  const validatedData = readSheetData(wsValidated);
  const testData = readSheetData(wsTest);

  console.log(`Rows in validated sheet: ${validatedData.size}`);
  console.log(`Rows in test sheet: ${testData.size}\n`);

  // Determine columns to compare
  const columns = columnFilter
    ? { [columnFilter]: COLUMNS_TO_COMPARE[columnFilter] || columnFilter }
    : COLUMNS_TO_COMPARE;

  for (const [column, columnName] of Object.entries(columns)) {
    console.log('='.repeat(70));
    console.log(`  COLUMN ${column} - ${columnName}`);
    console.log('='.repeat(70));

    let validatedTotal = 0;
    let testTotal = 0;
    const differences = [];
    const onlyInValidated = [];
    const onlyInTest = [];
    const processedPatients = new Set();

    // Iterate validated sheet
    for (const [key, entry] of validatedData) {
      const validatedValue = entry.values[column] || 0;
      validatedTotal += validatedValue;

      const matchKey = findMatch(key, validatedData, testData);
      processedPatients.add(key);

      if (matchKey) {
        const testEntry = testData.get(matchKey);
        const testValue = testEntry.values[column] || 0;

        if (Math.abs(validatedValue - testValue) > 0.01) {
          differences.push({
            patient: entry.name,
            date: entry.date,
            validated: validatedValue,
            test: testValue,
            diff: testValue - validatedValue,
          });
        }
      } else if (validatedValue > 0) {
        onlyInValidated.push({
          patient: entry.name,
          date: entry.date,
          value: validatedValue,
        });
      }
    }

    // Iterate test sheet to find entries only in test
    for (const [key, entry] of testData) {
      const testValue = entry.values[column] || 0;
      testTotal += testValue;

      const matchKey = findMatch(key, testData, validatedData);
      if (!matchKey && testValue > 0) {
        onlyInTest.push({
          patient: entry.name,
          date: entry.date,
          value: testValue,
        });
      }
    }

    console.log(`\n  Validated Total: R$ ${validatedTotal.toFixed(2)}`);
    console.log(`  Test Total:      R$ ${testTotal.toFixed(2)}`);
    console.log(`  Difference:      R$ ${(testTotal - validatedTotal).toFixed(2)}\n`);

    if (differences.length > 0) {
      console.log(`  --- DIFFERENT values (${differences.length}) ---`);
      console.table(differences.map(d => ({
        patient: d.patient,
        date: d.date,
        validated: d.validated.toFixed(2),
        test: d.test.toFixed(2),
        difference: (d.diff > 0 ? '+' : '') + d.diff.toFixed(2),
      })));
    }

    if (onlyInValidated.length > 0) {
      console.log(`  --- Only in VALIDATED (${onlyInValidated.length}) ---`);
      console.table(onlyInValidated.map(d => ({
        patient: d.patient,
        date: d.date,
        value: d.value.toFixed(2),
      })));
    }

    if (onlyInTest.length > 0) {
      console.log(`  --- Only in TEST (${onlyInTest.length}) ---`);
      console.table(onlyInTest.map(d => ({
        patient: d.patient,
        date: d.date,
        value: d.value.toFixed(2),
      })));
    }

    if (differences.length === 0 && onlyInValidated.length === 0 && onlyInTest.length === 0) {
      console.log('  All matching! No differences found.');
    }

    console.log('');
  }
}

main().catch(console.error);

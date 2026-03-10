/**
 * Script to retroactively fill CPF (column AI) for existing patients in the spreadsheet.
 *
 * Strategy:
 * 1. For each month sheet, fetch the financial report for that month's date range
 * 2. Build a map of NomePaciente → PacienteID from the report
 * 3. For each row without CPF, look up the patient ID and fetch CPF from API
 *
 * Usage: node -r dotenv/config src/preencher-cpf.js
 */

import 'dotenv/config';
import ExcelJS from 'exceljs';
import path from 'path';
import { fetchTransactions, fetchPatientCPF } from './services/feegow.js';
import { resolveSheetName, MONTH_NAMES_PT } from './services/excel.js';
import { createLogger, logSeparator } from './services/logger.js';

const log = createLogger('CPF');
const ONEDRIVE_PATH = path.join(process.cwd(), 'onedrive');

// Yellow fill for CPF column
const FILL_YELLOW = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFFFE598' },
  bgColor: { argb: 'FFFFFF99' },
};

const CELL_FONT = { name: 'Arial', size: 11 };

/** Remove accents for comparison: ÉDER → EDER, THAÍS → THAIS */
function normalize(str) {
  return str.toUpperCase().trim().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

/** Levenshtein distance between two strings */
function levenshtein(a, b) {
  const m = a.length, n = b.length;
  const dp = Array.from({ length: m + 1 }, () => Array(n + 1).fill(0));
  for (let i = 0; i <= m; i++) dp[i][0] = i;
  for (let j = 0; j <= n; j++) dp[0][j] = j;
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      dp[i][j] = a[i - 1] === b[j - 1]
        ? dp[i - 1][j - 1]
        : 1 + Math.min(dp[i - 1][j - 1], dp[i - 1][j], dp[i][j - 1]);
    }
  }
  return dp[m][n];
}

/**
 * Fuzzy match: find the closest name in the map.
 * Only accepts if edit distance ≤ maxDist (default 3) and the match is
 * unambiguous (no other name within the same distance).
 */
function fuzzyMatch(nameNorm, patientMap, maxDist = 3) {
  let bestName = null;
  let bestDist = Infinity;
  let ambiguous = false;

  for (const [mapName] of patientMap) {
    // Quick length filter — skip names too different in length
    if (Math.abs(mapName.length - nameNorm.length) > maxDist) continue;
    const dist = levenshtein(nameNorm, mapName);
    if (dist < bestDist) {
      bestDist = dist;
      bestName = mapName;
      ambiguous = false;
    } else if (dist === bestDist && mapName !== bestName) {
      ambiguous = true;
    }
  }

  if (bestDist <= maxDist && !ambiguous) {
    return { name: bestName, dist: bestDist };
  }
  return null;
}

const CELL_BORDER = {
  top: { style: 'dotted' },
  left: { style: 'dotted' },
  bottom: { style: 'dotted' },
  right: { style: 'dotted' },
};

/**
 * Fetches the financial report for a full month and builds a name→ID map
 */
async function buildPatientMap(year, month) {
  const mm = String(month).padStart(2, '0');
  const lastDay = new Date(year, month, 0).getDate();
  const startDate = `01/${mm}/${year}`;
  const endDate = `${lastDay}/${mm}/${year}`;

  log.info(`Buscando relatório: ${startDate} a ${endDate}`);
  const transactions = await fetchTransactions(startDate, endDate);

  const map = new Map();
  for (const t of transactions) {
    if (t.NomePaciente && t.PacienteID) {
      const name = normalize(t.NomePaciente);
      if (!map.has(name)) {
        map.set(name, t.PacienteID);
      }
    }
  }

  log.info(`${map.size} pacientes únicos encontrados no relatório`);
  return map;
}

async function main() {
  logSeparator('PREENCHIMENTO RETROATIVO DE CPF');

  const filePath = path.join(ONEDRIVE_PATH, process.env.EXCEL_FILE_PATH);
  log.info(`Planilha: ${filePath}`);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  // Usage: node -r dotenv/config src/preencher-cpf.js [year]
  // Default: 2026. Pass 2025 to process 2025 sheets.
  const year = Number(process.argv[2]) || 2026;
  const months = Array.from({ length: 12 }, (_, i) => i + 1); // All months, skips missing tabs

  let totalFilled = 0;
  let totalNotFound = 0;
  const notFoundNames = new Set();

  for (const month of months) {
    // 2025 uses "JANEIRO 25", 2026 uses "JANEIRO 2026"
    const yearSuffix = year <= 2025 ? String(year).slice(-2) : String(year);
    const sheetName = `${MONTH_NAMES_PT[month - 1]} ${yearSuffix}`;
    const worksheet = workbook.getWorksheet(sheetName);

    if (!worksheet) {
      log.warn(`Aba "${sheetName}" não encontrada, pulando...`);
      continue;
    }

    logSeparator(`${sheetName}`);

    // Build patient map + date index from financial report
    const patientMap = await buildPatientMap(year, month);

    // Iterate rows: fill missing CPFs, verify existing ones, update font on all
    let filled = 0;
    let verified = 0;
    let mismatch = 0;
    let notFound = 0;
    let fontUpdated = 0;

    for (let r = 8; r <= 1000; r++) {
      const nameCell = worksheet.getCell(`B${r}`);
      const name = nameCell.value;
      if (!name || String(name).trim() === '') break;

      const cpfCell = worksheet.getCell(`AI${r}`);
      const existingCpf = cpfCell.value ? String(cpfCell.value).trim() : '';
      const nameNorm = normalize(String(name));
      let patientId = patientMap.get(nameNorm);

      // Level 2: Fuzzy match (dist ≤ 2)
      if (!patientId) {
        const fuzzy = fuzzyMatch(nameNorm, patientMap, 2);
        if (fuzzy) {
          patientId = patientMap.get(fuzzy.name);
          log.info(`FUZZY linha ${r}: "${nameNorm}" → "${fuzzy.name}" (dist=${fuzzy.dist})`);
        }
      }

      if (!patientId) {
        if (!existingCpf) {
          notFound++;
          notFoundNames.add(nameNorm);
        } else {
          // Already has CPF but can't verify — just update font
          cpfCell.font = CELL_FONT;
          cpfCell.fill = FILL_YELLOW;
          cpfCell.border = CELL_BORDER;
          fontUpdated++;
        }
        continue;
      }

      const cpf = await fetchPatientCPF(patientId);

      if (!cpf) {
        if (!existingCpf) {
          notFound++;
          notFoundNames.add(nameNorm);
        }
        continue;
      }

      if (existingCpf) {
        // Verify: existing CPF must match API
        if (existingCpf === cpf) {
          verified++;
        } else {
          mismatch++;
          log.warn(`DIVERGÊNCIA linha ${r}: ${String(name).substring(0, 30)} | Planilha: "${existingCpf}" | API: "${cpf}"`);
          cpfCell.value = cpf;
        }
        cpfCell.font = CELL_FONT;
        cpfCell.fill = FILL_YELLOW;
        cpfCell.border = CELL_BORDER;
        fontUpdated++;
      } else {
        // Fill new CPF
        cpfCell.value = cpf;
        cpfCell.fill = FILL_YELLOW;
        cpfCell.font = CELL_FONT;
        cpfCell.border = CELL_BORDER;
        filled++;
      }
    }

    log.info(`${sheetName}: ${filled} novos, ${verified} verificados OK, ${mismatch} divergências, ${fontUpdated} fontes atualizadas, ${notFound} não encontrados`);
    totalFilled += filled;
    totalNotFound += notFound;
  }

  // Save
  logSeparator('SALVANDO PLANILHA');
  await workbook.xlsx.writeFile(filePath);
  log.info('Planilha salva com sucesso');

  // Summary
  logSeparator('RESUMO');
  log.info(`Total CPFs preenchidos: ${totalFilled}`);
  log.info(`Total não encontrados: ${totalNotFound}`);

  if (notFoundNames.size > 0) {
    log.warn('Pacientes não encontrados no relatório:');
    for (const name of notFoundNames) {
      console.log(`  - ${name}`);
    }
  }

  logSeparator('CONCLUÍDO');
}

main().catch(e => {
  log.error('Erro fatal', { message: e.message, stack: e.stack });
  process.exit(1);
});

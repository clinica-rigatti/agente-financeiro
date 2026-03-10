/**
 * Generates a consolidated patient spreadsheet from multiple year spreadsheets.
 * Columns: Data, Nome, Valor, CPF, Telefone, Email
 *
 * Usage: node -r dotenv/config src/gerar-consolidado.js
 */

import 'dotenv/config';
import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { fetchTransactions, fetchPatient } from './services/feegow.js';
import { MONTH_NAMES_PT } from './services/excel.js';
import { createLogger, logSeparator } from './services/logger.js';

const log = createLogger('Consolidado');
const ONEDRIVE_PATH = path.join(process.cwd(), 'onedrive');

/** Detect payment columns dynamically from header row */
function detectPaymentCols(worksheet, headerRow) {
  const payCols = [];
  const row = worksheet.getRow(headerRow);
  row.eachCell({ includeEmpty: false }, (cell, col) => {
    const val = String(cell.value || '').toUpperCase().trim();
    // Skip tax/fee columns
    if (val.includes('TAXA') || val.startsWith('TX ')) return;
    // Skip NP/future columns
    if (val.includes('BOLETOS FUTUROS') || val.includes('NP ')) return;
    // Payment column detection
    const isPayment =
      val.includes('DINHEIRO') ||
      (val.includes('CIELO') && !val.includes('TX')) ||
      (val.includes('PAGBANK') && !val.includes('TX')) ||
      (val.includes('PAG BANK') && !val.includes('TX')) ||
      (val.includes('SAFRA') && !val.includes('TX')) ||
      val.includes('PIX') || val.includes('TRANSF') || val.includes('TRANF') ||
      val.includes('CHEQUE') ||
      val.includes('LINK') ||
      val.includes('STONE') ||
      val.includes('PAY PAL') || val.includes('PAYPAL') ||
      (val.includes('CART') && !val.includes('TAXA') && !val.includes('TX')) || // CARTÃO, CREDITO/DEBITO
      val.includes('CREDITO') ||
      val.includes('RECORRENTE');
    // BOLETO (exact) — only the first one (payment), not duplicates at end
    const isBoleto = val === 'BOLETO' && !payCols.some(c => c.header === 'BOLETO');
    if (isPayment || isBoleto) {
      const letter = col <= 26 ? String.fromCharCode(64 + col) : 'A' + String.fromCharCode(64 + col - 26);
      payCols.push({ col: letter, header: val });
    }
  });
  return payCols.map(c => c.col);
}

// Spreadsheet configs per year
const SPREADSHEETS = [
  {
    year: 2019,
    file: 'planilhas/extras/Entradas.xlsx',
    paymentCols: 'auto',
    cpfCol: null,
    sheetSuffix: '2019',
    dataStartRow: 'auto',
  },
  {
    year: 2020,
    file: 'planilhas/TESTES LOCAL ENTRADAS ANO 2020 A 2021.xlsx',
    paymentCols: 'auto',
    cpfCol: null,
    sheetSuffix: '2020',
    dataStartRow: 'auto',
  },
  {
    year: 2021,
    file: 'planilhas/TESTES LOCAL ENTRADAS ANO 2020 A 2021.xlsx',
    paymentCols: 'auto',
    cpfCol: null,
    sheetSuffix: '', // 2021 tabs have no year suffix (just "JANEIRO", "FEVEREIRO", etc.)
    dataStartRow: 'auto',
  },
  {
    year: 2022,
    file: 'planilhas/TESTES LOCAL ENTRADAS ANO 2022.xlsx',
    paymentCols: 'auto',
    cpfCol: null,
    sheetSuffix: '2022',
    dataStartRow: 'auto',
  },
  {
    year: 2023,
    file: 'planilhas/TESTES LOCAL ENTRADAS ANO 2023.xlsx',
    paymentCols: 'auto', // Detected dynamically per sheet
    cpfCol: null,
    sheetSuffix: 'auto', // Mixed: "2023" and "23"
    dataStartRow: 'auto', // Either 8 or 9 depending on sheet
  },
  {
    year: 2024,
    file: 'planilhas/TESTES LOCAL ENTRADAS ANO 2024.xlsx',
    paymentCols: ['F', 'G', 'H', 'I'], // 2024: F=DINHEIRO, G=CIELO, H=SAFRA, I=PIX
    cpfCol: null, // No CPF column in 2024
    sheetSuffix: '24', // "JANEIRO 24"
    dataStartRow: 2, // Headers in row 1, data starts row 2
  },
  {
    year: 2025,
    file: 'planilhas/TESTES LOCAL ENTRADAS ANO 2025.xlsx',
    paymentCols: ['S', 'T', 'U', 'V', 'W', 'X'], // 2025: S=DINHEIRO .. X=BOLETO
    cpfCol: 'AI',
    sheetSuffix: '25', // "JANEIRO 25"
  },
  {
    year: 2026,
    file: 'planilhas/TESTES LOCAL ENTRADAS ANO 2026.xlsx',
    paymentCols: ['Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X'], // 2026: Q=DINHEIRO .. X=BOLETO
    cpfCol: 'AI',
    sheetSuffix: '2026', // "JANEIRO 2026"
  },
];

/** Check if name should be excluded (test entries, commissions, refunds, etc.) */
function isBlacklistedName(name) {
  const n = name.toUpperCase().trim();
  // Pure numbers or very short
  if (/^\d+$/.test(n) || n.length <= 2) return true;
  // Test transactions
  if (n.includes('TESTE') || n === 'TEST' || n.startsWith('TEST ')) return true;
  // Tax entries
  if (n.startsWith('TAXA ') || n === 'TAXA') return true;
  // Commission entries (COMISSÃO, COMISSAO, COMISSAÃO)
  if (n.includes('COMISS')) return true;
  // Commission check withdrawals
  if (n.startsWith('CHEQUE MASTER') || n.startsWith('CHEQUE COMPESANDO') || n.startsWith('CHEQUE COMISS')) return true;
  if (n.endsWith('CHEQUE MASTER')) return true;
  // Refund/return entries
  if (n.includes('ESTORNO') || n.includes('DEVOLU')) return true;
  // Machine tests
  if (n.includes('MAQUINA') || n.includes('MÁQUINA')) return true;
  // Internal entries
  if (n.includes('PACIENTE INTERNO')) return true;
  return false;
}

/** Remove accents for comparison */
function normalize(str) {
  return str.toUpperCase().trim().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

/** Levenshtein distance */
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

/** Fuzzy match: find closest name in map (dist ≤ 2, unambiguous) */
function fuzzyMatch(nameNorm, patientMap) {
  let bestName = null;
  let bestDist = Infinity;
  let ambiguous = false;

  for (const [mapName] of patientMap) {
    if (Math.abs(mapName.length - nameNorm.length) > 2) continue;
    const dist = levenshtein(nameNorm, mapName);
    if (dist < bestDist) {
      bestDist = dist;
      bestName = mapName;
      ambiguous = false;
    } else if (dist === bestDist && mapName !== bestName) {
      ambiguous = true;
    }
  }

  if (bestDist <= 2 && !ambiguous) return bestName;
  return null;
}

/** Build name→patientId map from Feegow financial report */
async function buildPatientMap(year, month) {
  const mm = String(month).padStart(2, '0');
  const lastDay = new Date(year, month, 0).getDate();
  const startDate = `01/${mm}/${year}`;
  const endDate = `${lastDay}/${mm}/${year}`;

  log.info(`Buscando relatório Feegow: ${startDate} a ${endDate}`);
  const transactions = await fetchTransactions(startDate, endDate);

  const map = new Map();
  for (const t of transactions) {
    if (t.NomePaciente && t.PacienteID) {
      const name = normalize(t.NomePaciente);
      if (!map.has(name)) map.set(name, t.PacienteID);
    }
  }
  return map;
}

// Cache patient contact data (Feegow API)
const patientCache = new Map();

async function fetchPatientContact(patientId) {
  const key = String(patientId);
  if (patientCache.has(key)) return patientCache.get(key);

  const patient = await fetchPatient(patientId);
  const contact = {
    cpf: patient?.documentos?.cpf || '',
    telefone: patient?.celulares?.[0] || patient?.telefones?.[0] || '',
    email: patient?.email?.[0] || '',
  };
  patientCache.set(key, contact);
  return contact;
}

/** Load MedX backup contacts (name→contact map) */
let medxContactMap = null;

function loadMedxContacts() {
  if (medxContactMap) return medxContactMap;

  const filePath = path.join(process.cwd(), 'medx-backup', 'Contatos.json');
  log.info(`Carregando contatos MedX: ${filePath}`);
  const data = fs.readFileSync(filePath, 'utf8');
  const contatos = JSON.parse(data);

  medxContactMap = new Map();
  for (const c of contatos) {
    if (!c.Nome) continue;
    const name = normalize(c.Nome);
    if (medxContactMap.has(name)) continue;
    const cel = c.Celular || '';
    medxContactMap.set(name, {
      cpf: (c['CPF/CGC'] || '').trim(),
      telefone: cel.includes('ERRO') ? '' : cel.trim(),
      email: (c.Email || '').trim(),
    });
  }
  log.info(`${medxContactMap.size} contatos MedX carregados`);
  return medxContactMap;
}

function lookupMedxContact(nameNorm) {
  const map = loadMedxContacts();
  let contact = map.get(nameNorm);
  if (contact) return contact;

  // Fuzzy match against MedX contacts
  const fuzzyName = fuzzyMatch(nameNorm, map);
  if (fuzzyName) return map.get(fuzzyName);
  return null;
}

/** Process one spreadsheet file and return rows */
async function processSpreadsheet(config) {
  const filePath = path.join(ONEDRIVE_PATH, config.file);
  log.info(`Lendo planilha: ${filePath}`);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const rows = [];
  const months = Array.from({ length: 12 }, (_, i) => i + 1);

  for (const month of months) {
    // Handle mixed suffixes (2023 uses both "2023" and "23")
    let worksheet = null;
    let sheetName = '';
    if (config.sheetSuffix === 'auto') {
      for (const suffix of ['2023', '23']) {
        sheetName = `${MONTH_NAMES_PT[month - 1]} ${suffix}`;
        worksheet = workbook.getWorksheet(sheetName);
        if (worksheet) break;
      }
    } else if (config.sheetSuffix === '') {
      // No suffix — tab name is just the month (e.g., "JANEIRO")
      sheetName = MONTH_NAMES_PT[month - 1];
      worksheet = workbook.getWorksheet(sheetName);
    } else {
      sheetName = `${MONTH_NAMES_PT[month - 1]} ${config.sheetSuffix}`;
      worksheet = workbook.getWorksheet(sheetName);
    }

    if (!worksheet) continue;

    // Detect data start row: scan rows 6-9 for header ("DATA" in col A)
    let startRow;
    let headerRow;
    if (config.dataStartRow === 'auto') {
      let dataRow = null;
      for (let r = 6; r <= 9; r++) {
        const val = String(worksheet.getCell('A' + r).value || '').toUpperCase().trim();
        if (val.includes('DATA')) { dataRow = r; break; }
      }
      if (dataRow) {
        // Check if next row is a sub-header (col B empty) or actual data
        const nextB = String(worksheet.getCell('B' + (dataRow + 1)).value || '').trim();
        if (nextB === '') {
          // Double header (e.g., 2019): sub-header row has payment col names
          headerRow = dataRow + 1;
          startRow = dataRow + 2;
        } else {
          // Single header row
          headerRow = dataRow;
          startRow = dataRow + 1;
        }
      } else {
        headerRow = 7; startRow = 8;
      }
    } else {
      startRow = config.dataStartRow || 8;
      headerRow = startRow - 1;
    }

    // Check if sheet has data
    const firstName = worksheet.getCell(`B${startRow}`).value;
    if (!firstName || String(firstName).trim() === '') continue;

    logSeparator(sheetName);

    // Detect payment columns dynamically if needed
    const paymentCols = config.paymentCols === 'auto'
      ? detectPaymentCols(worksheet, headerRow)
      : config.paymentCols;

    const useFeegow = config.year >= 2025;
    const patientMap = useFeegow ? await buildPatientMap(config.year, month) : new Map();
    let count = 0;
    let noContact = 0;
    let skipped = 0;

    for (let r = startRow; r <= 2000; r++) {
      const nameVal = worksheet.getCell(`B${r}`).value;
      if (!nameVal || String(nameVal).trim() === '') break;

      const name = String(nameVal).trim();

      // Skip blacklisted names (test entries, commissions, refunds, etc.)
      if (isBlacklistedName(name)) {
        skipped++;
        continue;
      }

      // Read date (can be Date, serial number, or DD/MM/YYYY string)
      const dateVal = worksheet.getCell(`A${r}`).value;
      let dateObj = null;
      if (dateVal instanceof Date) {
        dateObj = dateVal;
      } else if (typeof dateVal === 'number' && dateVal > 1000) {
        // Excel serial number: days since 1900-01-01 (with Excel's 1900 leap year bug)
        const excelEpoch = new Date(1899, 11, 30); // Dec 30, 1899
        dateObj = new Date(excelEpoch.getTime() + dateVal * 86400000);
      } else if (dateVal) {
        const parts = String(dateVal).split('/');
        if (parts.length === 3) {
          dateObj = new Date(parts[2], parts[1] - 1, parts[0]);
        }
      }

      // Skip rows with dates outside the expected year
      if (dateObj && dateObj.getFullYear() !== config.year) {
        skipped++;
        continue;
      }

      // Sum payment columns (different per year)
      let valor = 0;
      for (const col of paymentCols) {
        const v = worksheet.getCell(`${col}${r}`).value;
        if (v && typeof v === 'number') valor += v;
      }
      valor = Math.round(valor * 100) / 100;

      // Skip rows with no payment value
      if (valor === 0) {
        skipped++;
        continue;
      }

      // Read CPF (some years don't have a CPF column)
      let cpf = '';
      if (config.cpfCol) {
        const cpfVal = worksheet.getCell(`${config.cpfCol}${r}`).value;
        cpf = cpfVal ? String(cpfVal).trim() : '';
      }

      // Find contact data
      const nameNorm = normalize(name);
      let telefone = '';
      let email = '';

      if (useFeegow) {
        // 2025+: use Feegow API
        let patientId = patientMap.get(nameNorm);
        if (!patientId) {
          const fuzzyName = fuzzyMatch(nameNorm, patientMap);
          if (fuzzyName) patientId = patientMap.get(fuzzyName);
        }
        if (patientId) {
          const contact = await fetchPatientContact(patientId);
          telefone = contact.telefone;
          email = contact.email;
          if (!cpf) cpf = contact.cpf;
        } else {
          noContact++;
        }
      } else {
        // 2024: use MedX backup contacts
        const contact = lookupMedxContact(nameNorm);
        if (contact) {
          telefone = contact.telefone;
          email = contact.email;
          if (!cpf) cpf = contact.cpf;
        } else {
          noContact++;
        }
      }

      rows.push({ dateObj, name, valor, cpf, telefone, email });
      count++;
    }

    log.info(`${count} linhas lidas, ${noContact} sem dados de contato${skipped ? `, ${skipped} ignoradas` : ''}`);
  }

  return rows;
}

async function main() {
  logSeparator('GERAR PLANILHA CONSOLIDADA');

  const allRows = [];

  for (const config of SPREADSHEETS) {
    logSeparator(`ANO ${config.year}`);
    const rows = await processSpreadsheet(config);
    allRows.push(...rows);
    log.info(`${config.year}: ${rows.length} linhas totais`);
  }

  // Load 2024 workbook for Histórico tab (2022/2023 name-only tabs)
  const file2024 = path.join(ONEDRIVE_PATH, 'planilhas/TESTES LOCAL ENTRADAS ANO 2024.xlsx');
  const wb2024 = new ExcelJS.Workbook();
  await wb2024.xlsx.readFile(file2024);

  // Sort by date descending
  allRows.sort((a, b) => {
    if (!a.dateObj && !b.dateObj) return 0;
    if (!a.dateObj) return 1;
    if (!b.dateObj) return -1;
    return b.dateObj.getTime() - a.dateObj.getTime();
  });

  // Create new workbook
  logSeparator('CRIANDO PLANILHA CONSOLIDADA');
  const newWorkbook = new ExcelJS.Workbook();
  const sheet = newWorkbook.addWorksheet('Consolidado');

  const cellAlignment = { horizontal: 'center', vertical: 'middle' };

  // Headers
  const headers = ['Data', 'Nome', 'Valor', 'CPF', 'Telefone', 'Email'];
  const headerRow = sheet.getRow(1);
  headers.forEach((h, i) => {
    const cell = headerRow.getCell(i + 1);
    cell.value = h;
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF4472C4' },
    };
    cell.font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.alignment = cellAlignment;
    cell.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };
  });

  // Column widths
  sheet.getColumn(1).width = 14;  // Data
  sheet.getColumn(2).width = 42;  // Nome
  sheet.getColumn(3).width = 16;  // Valor
  sheet.getColumn(4).width = 16;  // CPF
  sheet.getColumn(5).width = 18;  // Telefone
  sheet.getColumn(6).width = 35;  // Email

  // Data rows
  const cellBorder = {
    top: { style: 'dotted' },
    left: { style: 'dotted' },
    bottom: { style: 'dotted' },
    right: { style: 'dotted' },
  };
  const cellFont = { name: 'Arial', size: 11 };

  for (let i = 0; i < allRows.length; i++) {
    const r = allRows[i];
    const row = sheet.getRow(i + 2);

    const dateCell = row.getCell(1);
    if (r.dateObj) {
      dateCell.value = r.dateObj;
      dateCell.numFmt = 'dd"/"mm"/"yyyy';
    }
    dateCell.font = cellFont;
    dateCell.border = cellBorder;
    dateCell.alignment = cellAlignment;

    const nameCell = row.getCell(2);
    nameCell.value = r.name;
    nameCell.font = cellFont;
    nameCell.border = cellBorder;
    nameCell.alignment = cellAlignment;

    const valorCell = row.getCell(3);
    valorCell.value = r.valor;
    valorCell.numFmt = '_-"R$" * #,##0.00_-;-"R$" * #,##0.00_-;_-"R$" * "-"??_-;_-@';
    valorCell.font = cellFont;
    valorCell.border = cellBorder;
    valorCell.alignment = cellAlignment;

    const cpfCell = row.getCell(4);
    cpfCell.value = r.cpf;
    cpfCell.font = cellFont;
    cpfCell.border = cellBorder;
    cpfCell.alignment = cellAlignment;

    const telCell = row.getCell(5);
    telCell.value = r.telefone;
    telCell.font = cellFont;
    telCell.border = cellBorder;
    telCell.alignment = cellAlignment;

    const emailCell = row.getCell(6);
    emailCell.value = r.email;
    emailCell.font = cellFont;
    emailCell.border = cellBorder;
    emailCell.alignment = cellAlignment;
  }

  // ── Historic tab: patients from 2023 and earlier (MedX data, no values) ──
  logSeparator('ABA HISTÓRICO (PRE-2024)');

  const historicRows = [];

  // 2023 and 2022 name-only tabs (date + name, no values)
  for (const tabName of ['2023', '2022']) {
    const ws = wb2024.getWorksheet(tabName);
    if (!ws) continue;
    // Multiple date/name pairs across columns: A/B, D/E, G/H, J/K, M/N, P/Q, S/T, V/W, Y/Z, AB/AC
    const pairs = [
      [1, 2], [4, 5], [7, 8], [10, 11], [13, 14],
      [16, 17], [19, 20], [22, 23], [25, 26], [28, 29],
    ];
    let count = 0;
    for (const [dc, nc] of pairs) {
      for (let r = 2; r <= 2000; r++) {
        const nameVal = ws.getCell(r, nc).value;
        if (!nameVal || String(nameVal).trim() === '') break;
        const name = String(nameVal).trim();
        const dateVal = ws.getCell(r, dc).value;
        let dateObj = null;
        if (dateVal instanceof Date) {
          dateObj = dateVal;
        } else if (typeof dateVal === 'number' && dateVal > 1000) {
          const excelEpoch = new Date(1899, 11, 30);
          dateObj = new Date(excelEpoch.getTime() + dateVal * 86400000);
        }
        // Filter: tab "2023" → only 2023 dates, tab "2022" → only 2022 dates
        const expectedYear = parseInt(tabName, 10);
        if (dateObj && dateObj.getFullYear() !== expectedYear) continue;

        const nameNorm = normalize(name);
        const contact = lookupMedxContact(nameNorm);
        historicRows.push({
          dateObj,
          name,
          valor: 0,
          cpf: contact?.cpf || '',
          telefone: contact?.telefone || '',
          email: contact?.email || '',
        });
        count++;
      }
    }
    log.info(`Aba ${tabName}: ${count} linhas`);
  }

  // Sort historic rows by date descending
  historicRows.sort((a, b) => {
    if (!a.dateObj && !b.dateObj) return 0;
    if (!a.dateObj) return 1;
    if (!b.dateObj) return -1;
    return b.dateObj.getTime() - a.dateObj.getTime();
  });

  // Write historic sheet
  const histSheet = newWorkbook.addWorksheet('Histórico (2023-)');
  const histHeaders = ['Data', 'Nome', 'Valor', 'CPF', 'Telefone', 'Email'];
  const histHeaderRow = histSheet.getRow(1);
  histHeaders.forEach((h, i) => {
    const cell = histHeaderRow.getCell(i + 1);
    cell.value = h;
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF70AD47' } };
    cell.font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.alignment = cellAlignment;
    cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
  });
  histSheet.getColumn(1).width = 14;
  histSheet.getColumn(2).width = 42;
  histSheet.getColumn(3).width = 16;
  histSheet.getColumn(4).width = 16;
  histSheet.getColumn(5).width = 18;
  histSheet.getColumn(6).width = 35;

  for (let i = 0; i < historicRows.length; i++) {
    const r = historicRows[i];
    const row = histSheet.getRow(i + 2);

    const dateCell = row.getCell(1);
    if (r.dateObj) {
      dateCell.value = r.dateObj;
      dateCell.numFmt = 'dd"/"mm"/"yyyy';
    }
    dateCell.font = cellFont;
    dateCell.border = cellBorder;
    dateCell.alignment = cellAlignment;

    const nameCell = row.getCell(2);
    nameCell.value = r.name;
    nameCell.font = cellFont;
    nameCell.border = cellBorder;
    nameCell.alignment = cellAlignment;

    const valorCell = row.getCell(3);
    if (r.valor > 0) {
      valorCell.value = r.valor;
      valorCell.numFmt = '_-"R$" * #,##0.00_-;-"R$" * #,##0.00_-;_-"R$" * "-"??_-;_-@';
    }
    valorCell.font = cellFont;
    valorCell.border = cellBorder;
    valorCell.alignment = cellAlignment;

    const cpfCell = row.getCell(4);
    cpfCell.value = r.cpf;
    cpfCell.font = cellFont;
    cpfCell.border = cellBorder;
    cpfCell.alignment = cellAlignment;

    const telCell = row.getCell(5);
    telCell.value = r.telefone;
    telCell.font = cellFont;
    telCell.border = cellBorder;
    telCell.alignment = cellAlignment;

    const emailCell = row.getCell(6);
    emailCell.value = r.email;
    emailCell.font = cellFont;
    emailCell.border = cellBorder;
    emailCell.alignment = cellAlignment;
  }

  log.info(`Aba Histórico: ${historicRows.length} linhas totais`);
  const histWithContact = historicRows.filter(r => r.telefone || r.email).length;
  log.info(`  Com contato: ${histWithContact} | Sem contato: ${historicRows.length - histWithContact}`);

  // Save
  const outputPath = path.join(ONEDRIVE_PATH, 'planilhas', 'consolidado-pacientes.xlsx');
  await newWorkbook.xlsx.writeFile(outputPath);
  log.info(`Planilha salva em: ${outputPath}`);

  logSeparator('RESUMO');
  log.info(`Aba Consolidado: ${allRows.length} linhas (2024-2026)`);
  log.info(`Aba Histórico: ${historicRows.length} linhas (2023 e antes)`);
  log.info(`Pacientes com contato (Consolidado): ${allRows.filter(r => r.telefone || r.email).length}`);
  log.info(`Pacientes com contato (Histórico): ${histWithContact}`);
  logSeparator('CONCLUÍDO');
}

main().catch(e => {
  log.error('Erro fatal', { message: e.message, stack: e.stack });
  process.exit(1);
});

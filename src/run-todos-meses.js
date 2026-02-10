import dotenv from 'dotenv';
dotenv.config();

import ExcelJS from 'exceljs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const filePath = path.join(__dirname, '../onedrive/planilhas/TESTES - ENTRADAS  ANO 2025.xlsx');

const MONTHS = [
  { name: 'JANEIRO',   start: '01/01/2025', end: '31/01/2025', sheet: 'TESTES-JANEIRO-25' },
  { name: 'FEVEREIRO',  start: '01/02/2025', end: '28/02/2025', sheet: 'TESTES-FEVEREIRO-25' },
  { name: 'MARÇO',      start: '01/03/2025', end: '31/03/2025', sheet: 'TESTES-MARÇO-25' },
  { name: 'ABRIL',      start: '01/04/2025', end: '30/04/2025', sheet: 'TESTES-ABRIL-25' },
  { name: 'MAIO',       start: '01/05/2025', end: '31/05/2025', sheet: 'TESTES-MAIO-25' },
  { name: 'JUNHO',      start: '01/06/2025', end: '30/06/2025', sheet: 'TESTES-JUNHO-25' },
  { name: 'JULHO',      start: '01/07/2025', end: '31/07/2025', sheet: 'TESTES-JULHO-25' },
  { name: 'AGOSTO',     start: '01/08/2025', end: '31/08/2025', sheet: 'TESTES-AGOSTO-25' },
  { name: 'SETEMBRO',   start: '01/09/2025', end: '30/09/2025', sheet: 'TESTES-SETEMBRO-25' },
  { name: 'OUTUBRO',    start: '01/10/2025', end: '31/10/2025', sheet: 'TESTES-OUTUBRO-25' },
  { name: 'NOVEMBRO',   start: '01/11/2025', end: '30/11/2025', sheet: 'TESTES-NOVEMBRO-25' },
  { name: 'DEZEMBRO',   start: '01/12/2025', end: '31/12/2025', sheet: 'TESTES-DEZEMBRO-25' },
];

async function main() {
  // Dynamic imports to ensure env vars are already loaded
  const { fetchDetailedTransactions } = await import('./services/feegow.js');
  const { updateSpreadsheetCustom } = await import('./services/excel.js');

  console.log(`Abrindo: ${filePath}`);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const startTotal = Date.now();

  for (const month of MONTHS) {
    console.log(`\n${'='.repeat(60)}`);
    console.log(`  ${month.name} 2025 (${month.start} - ${month.end})`);
    console.log(`${'='.repeat(60)}`);

    const ws = workbook.getWorksheet(month.sheet);
    if (!ws) {
      console.log(`  Aba "${month.sheet}" não encontrada, pulando...`);
      continue;
    }

    // Clear data (keep headers rows 1-7)
    for (let r = 8; r <= 500; r++) {
      const row = ws.getRow(r);
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (!cell.formula) cell.value = null;
      });
    }

    // Fetch transactions
    const start = Date.now();
    const transactions = await fetchDetailedTransactions(month.start, month.end);
    console.log(`  ${transactions.length} movimentações (${((Date.now() - start) / 1000).toFixed(1)}s)`);

    // Insert into sheet
    await updateSpreadsheetCustom(workbook, month.sheet, transactions, month.start);
  }

  console.log('\n\nSalvando arquivo...');
  await workbook.xlsx.writeFile(filePath);
  console.log(`Arquivo salvo! Tempo total: ${((Date.now() - startTotal) / 1000).toFixed(1)}s`);
}

main().catch(console.error);

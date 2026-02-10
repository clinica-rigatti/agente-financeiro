/**
 * Diagnostic script to analyze API data vs mapping
 * Shows exactly what comes from the API and how it's being classified
 */

import 'dotenv/config';
import { fetchDetailedTransactions, fetchGroupMapping } from './services/feegow.js';
import { createLogger, logSeparator } from './services/logger.js';

const log = createLogger('Diag');

// Name-based mapping (fallback when ID is not available)
const PROCEDURE_MAPPING = {
  'nutricional': 'NUTRIS',
  'avalia': 'AVALIACAO',
  'rigatti': 'DR_RIGATTI',
  'recompra': 'DR_RIGATTI',
  'protocolo': 'DR_RIGATTI',
  'consulta': 'AVALIACAO',
  'victor': 'EXTRA',
  'implante': 'IMPLANTE',
  'soro': 'SORO_DR',
  'soroterapia': 'SORO_DR',
  'noripurum': 'SORO_DR',
  'tirzepatida': 'TIRZEPATIDA',
  'injetável': 'APLICACOES',
  'injetavel': 'APLICACOES',
  'cipionato': 'APLICACOES',
  'testosterona': 'APLICACOES',
  'blend': 'APLICACOES',
  'gh ': 'APLICACOES',
  'genotropin': 'APLICACOES',
  'omnitrope': 'APLICACOES',
  'nutri': 'NUTRIS',
  'tratamento proposto': 'DR_RIGATTI',
  'juros': 'JUROS_CARTAO',
  'estorno': 'ESTORNO_PROC',
};

let groupMapping = null;

function classifyProcedure(name, procedureId = null) {
  // 1) Try by ID (most precise)
  if (procedureId && groupMapping) {
    const columnById = groupMapping.get(Number(procedureId));
    if (columnById !== undefined) {
      return columnById; // Returns null for discarded
    }
  }

  // 2) Fallback: pattern matching by name
  const normalizedName = name.toLowerCase();
  for (const [pattern, column] of Object.entries(PROCEDURE_MAPPING)) {
    if (normalizedName.includes(pattern.toLowerCase())) {
      return column;
    }
  }
  return 'EXTRA';
}

async function main() {
  const args = process.argv.slice(2);
  const startDate = args[0] || '01/12/2025';
  const endDate = args[1] || '31/12/2025';

  logSeparator(`DIAGNOSTIC - ${startDate} to ${endDate}`);

  // Load group mapping from API
  groupMapping = await fetchGroupMapping();

  const transactions = await fetchDetailedTransactions(startDate, endDate);

  log.info(`Total: ${transactions.length} transactions`);

  // Analyze all detailed items
  const allItems = [];
  const unclassifiedItems = [];
  const summaryByColumn = {};

  for (const tx of transactions) {
    for (const item of (tx.detailedItems || [])) {
      const column = classifyProcedure(item.nome, item.procedimento_id);
      if (column === null) continue; // Discarded (e.g., Conversa/Dúvidas)
      allItems.push({
        patient: tx.NomePaciente,
        apiName: item.nome,
        procId: item.procedimento_id,
        reportProcName: tx.NomeProcedimento,
        itemValue: item.valor,
        transactionValue: tx.Value,
        column,
        source: tx.NomeProcedimento?.includes(',') ? 'INVOICE' : 'REPORT',
      });

      summaryByColumn[column] = (summaryByColumn[column] || 0) + item.valor;

      if (column === 'EXTRA') {
        unclassifiedItems.push({
          name: item.nome,
          value: item.valor,
          patient: tx.NomePaciente,
        });
      }
    }
  }

  // Show summary by column
  logSeparator('SUMMARY BY COLUMN');
  console.table(Object.entries(summaryByColumn)
    .sort((a, b) => b[1] - a[1])
    .map(([col, val]) => ({ column: col, total: val.toFixed(2) }))
  );

  // Show items that went to EXTRA (unclassified)
  logSeparator('UNCLASSIFIED ITEMS (EXTRA)');
  const uniqueNames = {};
  for (const item of unclassifiedItems) {
    if (!uniqueNames[item.name]) {
      uniqueNames[item.name] = { name: item.name, count: 0, totalValue: 0 };
    }
    uniqueNames[item.name].count++;
    uniqueNames[item.name].totalValue += item.value;
  }

  console.table(
    Object.values(uniqueNames)
      .sort((a, b) => b.totalValue - a.totalValue)
      .map(i => ({ name: i.name, occurrences: i.count, total: i.totalValue.toFixed(2) }))
  );

  // Show all unique procedure names (with ID and classification)
  logSeparator('ALL PROCEDURE NAMES (UNIQUE)');
  const uniqueProcs = new Map();
  for (const item of allItems) {
    const key = `${item.procId}_${item.apiName}`;
    if (!uniqueProcs.has(key)) {
      uniqueProcs.set(key, { name: item.apiName, id: item.procId, column: item.column });
    }
  }
  const sortedProcs = Array.from(uniqueProcs.values()).sort((a, b) => a.name.localeCompare(b.name));
  for (const proc of sortedProcs) {
    const idStr = proc.id ? `#${String(proc.id).padStart(3)}` : '  ? ';
    console.log(`  [${proc.column.padEnd(12)}] ${idStr} | ${proc.name}`);
  }

  // Show details of "Procedimento #XX" (items without report name)
  logSeparator('DETAILS OF "Procedimento #XX" (BY PATIENT)');
  const noNameItems = allItems.filter(i => i.apiName.startsWith('Procedimento #'));
  for (const item of noNameItems) {
    console.log(`  Patient: ${item.patient} | ${item.apiName} | Value: R$ ${item.itemValue.toFixed(2)} | Invoice proc: ${item.reportProcName}`);
  }
  if (noNameItems.length === 0) {
    console.log('  No "Procedimento #XX" items found');
  }

  // Compare source INVOICE vs REPORT
  logSeparator('DATA SOURCE');
  const sourceInvoice = allItems.filter(i => i.source === 'INVOICE');
  const sourceReport = allItems.filter(i => i.source === 'REPORT');
  console.log(`  Items via INVOICE (list-invoice): ${sourceInvoice.length}`);
  console.log(`  Items via REPORT (financial-movement): ${sourceReport.length}`);
  console.log(`  Total value INVOICE: ${sourceInvoice.reduce((s, i) => s + i.itemValue, 0).toFixed(2)}`);
  console.log(`  Total value REPORT: ${sourceReport.reduce((s, i) => s + i.itemValue, 0).toFixed(2)}`);
}

main().catch(console.error);

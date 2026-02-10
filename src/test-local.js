/**
 * Local test script using mock data
 * Does not make calls to the Feegow API
 */

import 'dotenv/config';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { updateSpreadsheet } from './services/excel.js';
import { createLogger, logSeparator, logTable } from './services/logger.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const log = createLogger('Test');

async function loadMockData() {
  const mockPath = path.join(__dirname, 'mocks/movimentacoes-api-mock.json');

  log.info(`Loading mock data from: ${mockPath}`);

  const data = JSON.parse(fs.readFileSync(mockPath, 'utf-8'));

  log.info('Mock data loaded', {
    period: data.periodo,
    count: data.quantidade,
  });

  return data.movimentacoes;
}

/**
 * Simulates enrichment with detailed items
 * Since we're using mock, we simulate items based on NomeProcedimento
 */
function simulateDetailedItems(transaction) {
  const { NomeProcedimento, Value } = transaction;

  // If no procedures or single one, return simple
  if (!NomeProcedimento || !NomeProcedimento.includes(',')) {
    return {
      ...transaction,
      detailedItems: [{
        procedimento_id: transaction.ProcedimentoID,
        nome: NomeProcedimento || 'Procedimento',
        valor: Value,
        quantidade: 1,
      }]
    };
  }

  // Split procedures and distribute value
  const procedures = NomeProcedimento.split(',').map(p => p.trim());
  const valuePerItem = Value / procedures.length;

  const detailedItems = procedures.map((name, index) => ({
    procedimento_id: `${transaction.ProcedimentoID}_${index}`,
    nome: name,
    valor: valuePerItem,
    quantidade: 1,
  }));

  return {
    ...transaction,
    detailedItems,
  };
}

async function main() {
  const startTime = Date.now();

  logSeparator('LOCAL TEST - AGENTE FINANCEIRO');
  log.info('Starting test with mock data');

  try {
    // 1. Load mock data
    logSeparator('STEP 1: LOADING MOCK DATA');
    const rawTransactions = await loadMockData();

    // 2. Simulate enrichment with detailed items
    logSeparator('STEP 2: PROCESSING TRANSACTIONS');
    log.info(`Processing ${rawTransactions.length} transactions`);

    const transactions = rawTransactions.map(simulateDetailedItems);

    // Log first transactions
    const summary = transactions.slice(0, 5).map(m => ({
      patient: m.NomePaciente,
      value: m.Value,
      payment: m.FormaPagamento,
      account: m.AccountName,
      items: m.detailedItems.length,
    }));

    logTable(summary, 'Sample of processed transactions');

    // 3. Update spreadsheet
    logSeparator('STEP 3: UPDATING SPREADSHEET');
    const referenceDate = transactions[0]?.Data || '05/02/2026';
    log.info(`Reference date: ${referenceDate}`);

    await updateSpreadsheet(transactions, referenceDate);

    // Finalization
    const duration = Date.now() - startTime;
    logSeparator('TEST FINISHED SUCCESSFULLY');

    log.info('Test summary', {
      processedTransactions: transactions.length,
      totalDuration: `${duration}ms`,
    });

  } catch (error) {
    log.error('Error during test', {
      message: error.message,
      stack: error.stack,
    });

    logSeparator('TEST FINISHED WITH ERROR');
    process.exit(1);
  }
}

main();

/**
 * Script to run the financial agent for a specific period
 * Usage: node src/run-periodo.js <startDate> <endDate>
 * Example: node src/run-periodo.js 01/12/2025 31/12/2025
 */

import 'dotenv/config';
import { fetchDetailedTransactions } from './services/feegow.js';
import { updateSpreadsheet } from './services/excel.js';
import { createLogger, logSeparator, logTable } from './services/logger.js';

const log = createLogger('Periodo');

async function main() {
  const args = process.argv.slice(2);

  if (args.length < 2) {
    console.log('Usage: node src/run-periodo.js <startDate> <endDate>');
    console.log('Example: node src/run-periodo.js 01/12/2025 31/12/2025');
    process.exit(1);
  }

  const [startDate, endDate] = args;
  const startTime = Date.now();

  logSeparator('AGENTE FINANCEIRO - MANUAL PERIOD');
  log.info(`Período: ${startDate} a ${endDate}`);

  try {
    // 1. Fetch transactions from API
    logSeparator('STEP 1: FETCHING DATA FROM FEEGOW API');
    log.info('Buscando movimentações...');

    const transactions = await fetchDetailedTransactions(startDate, endDate);

    if (transactions.length === 0) {
      log.warn('Nenhuma movimentação encontrada para este período');
      logSeparator('EXECUTION FINISHED (NO DATA)');
      return;
    }

    log.info(`Encontradas ${transactions.length} movimentações`);

    // Log summary
    const summary = transactions.slice(0, 10).map(m => ({
      patient: m.NomePaciente,
      value: m.Value,
      payment: m.FormaPagamento,
      items: m.detailedItems?.length || 1,
    }));
    logTable(summary, 'Primeiras 10 movimentações');

    if (transactions.length > 10) {
      log.info(`... e mais ${transactions.length - 10} movimentações`);
    }

    // 2. Update spreadsheet
    logSeparator('STEP 2: UPDATING SPREADSHEET');
    await updateSpreadsheet(transactions, startDate);

    // Finalization
    const duration = Date.now() - startTime;
    logSeparator('EXECUTION FINISHED SUCCESSFULLY');
    log.info('Summary', {
      period: `${startDate} a ${endDate}`,
      transactions: transactions.length,
      duration: `${(duration / 1000).toFixed(1)}s`,
    });

  } catch (error) {
    log.error('Erro durante a execução', {
      message: error.message,
      stack: error.stack,
    });
    logSeparator('EXECUTION FINISHED WITH ERROR');
    process.exit(1);
  }
}

main();

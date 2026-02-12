/**
 * Script to run the financial agent for specific patients only
 * Usage: node src/run-pacientes.js <startDate> <endDate> "Nome1" "Nome2" ...
 */

import 'dotenv/config';
import { fetchDetailedTransactions } from './services/feegow.js';
import { updateSpreadsheet } from './services/excel.js';
import { createLogger, logSeparator, logTable } from './services/logger.js';

const log = createLogger('Pacientes');

async function main() {
  const args = process.argv.slice(2);

  if (args.length < 3) {
    console.log('Usage: node src/run-pacientes.js <startDate> <endDate> "Nome1" "Nome2" ...');
    console.log('Example: node src/run-pacientes.js 06/02/2026 10/02/2026 "Ana Paula" "Anderson"');
    process.exit(1);
  }

  const [startDate, endDate, ...patientFilters] = args;
  const startTime = Date.now();

  logSeparator('AGENTE FINANCEIRO - FILTERED PATIENTS');
  log.info(`Período: ${startDate} a ${endDate}`);
  log.info(`Filtro: ${patientFilters.length} pacientes`);

  try {
    logSeparator('STEP 1: FETCHING DATA');
    const allTransactions = await fetchDetailedTransactions(startDate, endDate);
    log.info(`${allTransactions.length} movimentações totais`);

    // Filter by patient names (partial match, case insensitive)
    const normalizedFilters = patientFilters.map(f => f.toLowerCase());
    const transactions = allTransactions.filter(mov => {
      const name = (mov.NomePaciente || '').toLowerCase();
      return normalizedFilters.some(filter => name.includes(filter));
    });

    log.info(`${transactions.length} movimentações após filtro`);

    if (transactions.length === 0) {
      log.warn('Nenhuma movimentação encontrada para os pacientes filtrados');
      process.exit(0);
    }

    const summary = transactions.map(m => ({
      patient: m.NomePaciente,
      date: m.Data,
      value: m.Value,
      payment: m.FormaPagamento,
    }));
    logTable(summary, 'Movimentações filtradas');

    logSeparator('STEP 2: UPDATING SPREADSHEET');
    await updateSpreadsheet(transactions, startDate);

    const duration = Date.now() - startTime;
    logSeparator('FINISHED');
    log.info('Summary', {
      period: `${startDate} a ${endDate}`,
      filtered: transactions.length,
      total: allTransactions.length,
      duration: `${(duration / 1000).toFixed(1)}s`,
    });

  } catch (error) {
    log.error('Erro', { message: error.message, stack: error.stack });
    process.exit(1);
  }
}

main();

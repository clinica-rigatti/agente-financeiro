import 'dotenv/config';
import dayjs from 'dayjs';
import { fetchDetailedTransactions, isoToFeegow } from './services/feegow.js';
import { updateSpreadsheet } from './services/excel.js';
import { createLogger, logSeparator, logTable } from './services/logger.js';

const log = createLogger('Main');

async function main() {
  const startTime = Date.now();

  logSeparator('AGENTE FINANCEIRO');
  log.info('Iniciando execução');

  // Define fetch date (previous day if FETCH_PREVIOUS_DAY=true)
  const fetchPreviousDay = process.env.FETCH_PREVIOUS_DAY === 'true';
  const referenceDate = fetchPreviousDay
    ? dayjs().subtract(1, 'day')
    : dayjs();

  const formattedDate = referenceDate.format('YYYY-MM-DD');

  log.info('Configurações', {
    referenceDate: formattedDate,
    fetchPreviousDay,
    environment: process.env.NODE_ENV || 'development',
  });

  try {
    // 1. Fetch transactions from Feegow (with detailed items)
    logSeparator('STEP 1: DATA FETCH');
    log.info('Buscando dados da API Feegow...');

    const feegowDate = isoToFeegow(formattedDate);
    log.debug(`Data no formato Feegow: ${feegowDate}`);

    const transactions = await fetchDetailedTransactions(feegowDate, feegowDate);

    if (transactions.length === 0) {
      log.warn('Nenhuma movimentação encontrada para esta data');
      logSeparator('EXECUTION FINISHED (NO DATA)');
      return;
    }

    log.info(`Encontradas ${transactions.length} movimentações`);

    // Log transaction summary
    const transactionSummary = transactions.map(m => ({
      patient: m.NomePaciente,
      value: m.Value,
      paymentMethod: m.FormaPagamento,
      items: m.detailedItems?.length || 1,
    }));

    logTable(transactionSummary.slice(0, 10), 'Primeiras 10 movimentações');

    if (transactions.length > 10) {
      log.debug(`... e mais ${transactions.length - 10} movimentações`);
    }

    // 2. Update Excel spreadsheet
    logSeparator('STEP 2: SPREADSHEET UPDATE');
    await updateSpreadsheet(transactions, formattedDate);

    // Finalization
    const duration = Date.now() - startTime;
    logSeparator('EXECUTION FINISHED SUCCESSFULLY');

    log.info('Resumo da execução', {
      processedTransactions: transactions.length,
      totalDuration: `${duration}ms`,
      processedDate: formattedDate,
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

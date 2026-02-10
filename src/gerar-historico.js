/**
 * Script to generate history JSON without modifying the spreadsheet
 * Usage: node src/gerar-historico.js <startDate> <endDate>
 * Example: node src/gerar-historico.js 01/02/2026 10/02/2026
 */

import 'dotenv/config';
import { fetchDetailedTransactions } from './services/feegow.js';
import { groupTransactions, buildRow, loadGroupMapping } from './services/excel.js';
import { saveHistory } from './services/historico.js';
import { createLogger, logSeparator } from './services/logger.js';

const log = createLogger('GerarHistorico');

/**
 * Parse DD/MM/YYYY to { day, month, year } and generate all dates in range
 */
function parseDateRange(startStr, endStr) {
  const [sd, sm, sy] = startStr.split('/').map(Number);
  const [ed, em, ey] = endStr.split('/').map(Number);
  const start = new Date(sy, sm - 1, sd);
  const end = new Date(ey, em - 1, ed);

  const dates = [];
  const current = new Date(start);
  while (current <= end) {
    const d = String(current.getDate()).padStart(2, '0');
    const m = String(current.getMonth() + 1).padStart(2, '0');
    const y = current.getFullYear();
    dates.push(`${d}/${m}/${y}`);
    current.setDate(current.getDate() + 1);
  }
  return dates;
}

async function main() {
  const args = process.argv.slice(2);

  if (args.length < 2) {
    console.log('Usage: node src/gerar-historico.js <startDate> <endDate>');
    console.log('Example: node src/gerar-historico.js 01/02/2026 10/02/2026');
    process.exit(1);
  }

  const [startDate, endDate] = args;
  const startTime = Date.now();

  logSeparator('GERAR HISTÓRICO (SEM PLANILHA)');
  log.info(`Período: ${startDate} a ${endDate}`);

  try {
    // Load procedure group mapping (needed for classification in buildRow)
    await loadGroupMapping();

    // Generate all dates in range
    const dates = parseDateRange(startDate, endDate);
    log.info(`${dates.length} dias no período`);

    let totalPatients = 0;
    let daysWithData = 0;

    for (const date of dates) {
      log.info(`Buscando ${date}...`);

      const transactions = await fetchDetailedTransactions(date, date);

      if (transactions.length === 0) {
        log.debug(`${date}: sem movimentações`);
        continue;
      }

      // Group transactions by patient
      const grouped = groupTransactions(transactions);

      // Build rows (for classification details) without writing to spreadsheet
      const insertedRows = [];
      let fakeRow = 0;

      for (const group of grouped) {
        fakeRow++;
        const { rowData, classificationDetails } = await buildRow(group, date);

        // Calculate totals from rowData
        const procColumns = ['D', 'E', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P'];
        const payColumns = ['Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X'];

        let totalProc = 0;
        let totalPay = 0;
        for (const [col, val] of Object.entries(rowData)) {
          if (procColumns.includes(col) && typeof val === 'number') totalProc += val;
          if (payColumns.includes(col) && typeof val === 'number') totalPay += val;
        }

        insertedRows.push({
          row: fakeRow,
          patient: group.patientName,
          totalProcedures: totalProc,
          totalPayments: totalPay,
          group,
          classificationDetails,
        });
      }

      // Save history for this day
      await saveHistory(insertedRows, date);

      totalPatients += grouped.length;
      daysWithData++;
      log.info(`${date}: ${grouped.length} pacientes → histórico salvo`);
    }

    const duration = Date.now() - startTime;
    logSeparator('CONCLUÍDO');
    log.info('Resumo', {
      periodo: `${startDate} a ${endDate}`,
      diasComDados: daysWithData,
      totalPacientes: totalPatients,
      duracao: `${(duration / 1000).toFixed(1)}s`,
    });

  } catch (error) {
    log.error('Erro', { message: error.message, stack: error.stack });
    process.exit(1);
  }
}

main();

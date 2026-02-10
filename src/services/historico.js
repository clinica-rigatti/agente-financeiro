import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { createLogger } from './logger.js';

const log = createLogger('Historico');

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ONEDRIVE_PATH = path.resolve(__dirname, '../../onedrive');

// Local copies to avoid circular dependency with excel.js
const MONTH_NAMES_PT = [
  'JANEIRO', 'FEVEREIRO', 'MARÇO', 'ABRIL', 'MAIO', 'JUNHO',
  'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO',
];

const PAYMENT_COLUMNS = {
  DINHEIRO: 'Q', SICOOB: 'R', SAFRA: 'S', PIX_SAFRA: 'T',
  INFINITE: 'U', PIX: 'V', CHEQUE: 'W', BOLETO: 'X',
};

/**
 * Resolves the history filename from a date string
 * @param {string} date - Date in YYYY-MM-DD or DD/MM/YYYY format
 * @returns {string} Filename (e.g., "historico-fevereiro-2026.json")
 */
export function resolveHistoryFilename(date) {
  let month, year;
  if (date.includes('-')) {
    [year, month] = date.split('-');
  } else if (date.includes('/')) {
    const parts = date.split('/');
    month = parts[1];
    year = parts[2];
  } else {
    return 'historico-desconhecido.json';
  }

  const monthIndex = parseInt(month, 10) - 1;
  const monthName = MONTH_NAMES_PT[monthIndex];
  if (!monthName) return 'historico-desconhecido.json';

  return `historico-${monthName.toLowerCase()}-${year}.json`;
}

/**
 * Normalizes a date to YYYY-MM-DD format
 * @param {string} date - Date in YYYY-MM-DD or DD/MM/YYYY
 * @returns {string} Date in YYYY-MM-DD
 */
function normalizeDate(date) {
  if (date.includes('-') && date.indexOf('-') === 4) return date;
  if (date.includes('/')) {
    const [d, m, y] = date.split('/');
    return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
  }
  return date;
}

/**
 * Reads existing history file or returns empty structure
 * @param {string} filePath - Full path to history file
 * @param {string} date - Reference date for metadata
 * @returns {Object} History data
 */
function readOrCreateHistory(filePath, date) {
  let month, year;
  if (date.includes('-')) {
    [year, month] = date.split('-');
  } else if (date.includes('/')) {
    const parts = date.split('/');
    month = parts[1];
    year = parts[2];
  }
  const monthIndex = parseInt(month, 10) - 1;
  const monthName = MONTH_NAMES_PT[monthIndex] || 'DESCONHECIDO';

  const emptyHistory = {
    metadata: {
      mes: `${monthName} ${year}`,
      ultimaAtualizacao: null,
      execucoes: [],
    },
    dias: {},
  };

  if (!fs.existsSync(filePath)) return emptyHistory;

  try {
    const content = fs.readFileSync(filePath, 'utf-8');
    return JSON.parse(content);
  } catch (err) {
    log.warn(`Arquivo de histórico corrompido, criando backup: ${err.message}`);
    try {
      fs.renameSync(filePath, filePath + '.bak');
    } catch (e) { /* ignore */ }
    return emptyHistory;
  }
}

/**
 * Builds a patient entry for the history
 * @param {Object} insertedRow - Row data with group and classification
 * @returns {Object} Patient history entry
 */
function buildPatientEntry(insertedRow) {
  const { row, patient, totalProcedures, totalPayments, group, classificationDetails } = insertedRow;

  // Build transaction entries
  const transacoes = group.transactions.map(mov => ({
    idTransacao: mov.IDTransacao,
    numeroInvoice: mov.NumeroInvoice,
    valor: mov.Value,
    formaPagamento: mov.FormaPagamento,
    accountName: mov.AccountName || null,
    colunaPagamento: `${identifyPaymentColumnName(mov)}`,
    bandeira: mov.Bandeira || null,
    parcelas: mov.Parcelas || null,
    profissional: mov.NomeProfissional || null,
    fonteItens: mov.fonteItens || null,
    paymentOnly: mov.paymentOnly || false,
    descricao: mov.Descricao && mov.Descricao.trim() && mov.Descricao !== ',' ? mov.Descricao.trim() : null,
    itens: (mov.detailedItems || []).map(item => {
      const detail = classificationDetails.find(
        cd => cd.nome === item.nome && cd.valor === item.valor
      );
      return {
        nome: item.nome,
        procedimentoId: item.procedimento_id,
        valor: item.valor,
        quantidade: item.quantidade || 1,
        desconto: item.desconto || 0,
        colunaDestino: detail ? `${detail.colunaDestino} (${detail.colunaDestinoNome})` : null,
        classificadoPor: detail?.classificadoPor || null,
      };
    }),
  }));

  // Build resumoColunas from classification details
  const resumoColunas = {};
  for (const detail of classificationDetails) {
    if (detail.colunaDestino && detail.valor > 0) {
      const col = detail.colunaDestino;
      if (!resumoColunas[col]) {
        resumoColunas[col] = { nome: detail.colunaDestinoNome, valor: 0 };
      }
      resumoColunas[col].valor += detail.valor;
    }
  }

  return {
    pacienteId: group.patientId,
    nome: patient,
    linhaPlanilha: row,
    totalProcedimentos: totalProcedures,
    totalPagamentos: totalPayments,
    diferenca: Number((totalProcedures - totalPayments).toFixed(2)),
    observacao: group.transactions.some(t => t.paymentOnly) ? 'Tratamento em andamento' : null,
    transacoes,
    resumoColunas,
  };
}

/**
 * Gets the payment column name for a transaction
 * @param {Object} mov - Transaction
 * @returns {string} Column letter and name
 */
function identifyPaymentColumnName(mov) {
  const { FormaPagamentoID, AccountName } = mov;
  const account = (AccountName || '').toLowerCase();

  let col, name;
  if (account.includes('caixa') || FormaPagamentoID === 1) { col = PAYMENT_COLUMNS.DINHEIRO; name = 'DINHEIRO'; }
  else if (account.includes('sicoob')) { col = PAYMENT_COLUMNS.SICOOB; name = 'SICOOB'; }
  else if (account.includes('safra pay')) { col = PAYMENT_COLUMNS.SAFRA; name = 'SAFRA'; }
  else if (account.includes('safra') && FormaPagamentoID === 4) { col = PAYMENT_COLUMNS.BOLETO; name = 'BOLETO'; }
  else if (account.includes('safra')) { col = PAYMENT_COLUMNS.PIX_SAFRA; name = 'PIX_SAFRA'; }
  else if (account.includes('infinte') || account.includes('infinite')) { col = PAYMENT_COLUMNS.INFINITE; name = 'INFINITE'; }
  else if (account.includes('cora')) { col = PAYMENT_COLUMNS.PIX; name = 'PIX'; }
  else { col = PAYMENT_COLUMNS.PIX; name = 'PIX'; }

  return `${col} (${name})`;
}

/**
 * Saves history data for a given day
 * @param {Array} insertedRows - Rows with group + classificationDetails
 * @param {string} date - Reference date
 */
export async function saveHistory(insertedRows, date) {
  if (!insertedRows || insertedRows.length === 0) return;

  const filename = resolveHistoryFilename(date);
  const dirPath = path.join(ONEDRIVE_PATH, 'planilhas', 'historicos');
  if (!fs.existsSync(dirPath)) fs.mkdirSync(dirPath, { recursive: true });
  const filePath = path.join(dirPath, filename);
  const dateKey = normalizeDate(date);

  log.info(`Salvando histórico: ${filename} (${dateKey})`);

  const history = readOrCreateHistory(filePath, date);

  // Build day entry
  const pacientes = insertedRows.map(row => buildPatientEntry(row));

  const totalProcedimentos = pacientes.reduce((sum, p) => sum + p.totalProcedimentos, 0);
  const totalPagamentos = pacientes.reduce((sum, p) => sum + p.totalPagamentos, 0);

  history.dias[dateKey] = {
    resumo: {
      totalPacientes: pacientes.length,
      totalProcedimentos: Number(totalProcedimentos.toFixed(2)),
      totalPagamentos: Number(totalPagamentos.toFixed(2)),
    },
    pacientes,
  };

  // Update metadata
  history.metadata.ultimaAtualizacao = new Date().toISOString();
  history.metadata.execucoes.push({
    data: dateKey,
    executadoEm: new Date().toISOString(),
    totalPacientes: pacientes.length,
  });

  // Atomic write
  const tmpPath = filePath + '.tmp';
  fs.writeFileSync(tmpPath, JSON.stringify(history, null, 2), 'utf-8');
  fs.renameSync(tmpPath, filePath);

  log.info(`Histórico salvo: ${pacientes.length} pacientes para ${dateKey}`);
}

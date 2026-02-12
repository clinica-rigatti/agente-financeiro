import axios from 'axios';
import { createLogger } from './logger.js';

const log = createLogger('Feegow');

/**
 * Feegow public API client
 */

const api = axios.create({
  baseURL: process.env.FEEGOW_API_URL || 'https://api.feegow.com/v1',
  headers: {
    'Content-Type': 'application/json',
    'x-access-token': process.env.FEEGOW_API_TOKEN,
  },
});

/**
 * Fetches financial transactions from Feegow
 * @param {string} startDate - Start date in DD/MM/YYYY format
 * @param {string} endDate - End date in DD/MM/YYYY format
 * @param {object} options - Additional filter options
 * @returns {Promise<Array>} List of transactions
 */
export async function fetchTransactions(startDate, endDate, options = {}) {
  const {
    accountTypeIds = [3], // 3 = Patient (default)
    transactionType = null, // 1 = Income, -1 = Expense, null = All
  } = options;

  const body = {
    report: 'financial-movement',
    DATA_INICIO: startDate,
    DATA_FIM: endDate,
    TIPO_CONTA_IDS: accountTypeIds,
  };

  if (transactionType) {
    body.TIPO_MOVIMENTACAO_ID = transactionType;
  }

  try {
    const response = await api.post('/api/reports/generate', body);

    if (!response.data.success && response.data.success !== undefined) {
      throw new Error(response.data.message || 'Error generating report');
    }

    const transactions = response.data.data || [];

    log.info(`${transactions.length} movimentações encontradas`);

    return transactions;

  } catch (error) {
    if (error.response?.status === 401) {
      throw new Error('Token de API inválido ou expirado');
    }
    throw new Error(`Erro ao buscar movimentações: ${error.message}`);
  }
}

/**
 * Fetches patient details by ID
 * @param {string|number} patientId - Patient ID
 * @returns {Promise<Object>} Patient data
 */
export async function fetchPatient(patientId) {
  try {
    const response = await api.get('/api/patient/get', {
      params: { paciente_id: patientId }
    });

    return response.data.content || response.data;

  } catch (error) {
    log.warn(`Não foi possível buscar paciente ${patientId}`, { erro: error.message });
    return null;
  }
}

/**
 * Fetches transaction/invoice details by ID
 * @param {string|number} invoiceId - Invoice/transaction ID
 * @returns {Promise<Object>} Transaction data
 */
export async function fetchTransaction(invoiceId) {
  try {
    const response = await api.get('/api/financial/invoice', {
      params: { invoice_id: invoiceId }
    });

    return response.data.content || response.data;

  } catch (error) {
    log.warn(`Não foi possível buscar transação ${invoiceId}`, { erro: error.message });
    return null;
  }
}

/**
 * Fetches invoice details (items, payments, etc)
 * @param {string|number} invoiceId - Invoice ID
 * @param {string} transactionDate - Transaction date (DD/MM/YYYY)
 * @returns {Promise<Object|null>} Invoice data with detailed items
 */
export async function fetchInvoiceDetails(invoiceId, transactionDate) {
  try {
    // Endpoint requires data_start, data_end and tipo_transacao as mandatory
    // Date format: DD-MM-YYYY (with dash, not slash)
    const formattedDate = transactionDate
      ? transactionDate.replace(/\//g, '-')
      : null;

    const params = {
      invoice_id: invoiceId,
      tipo_transacao: 'C', // C = Accounts receivable (patients)
    };

    if (formattedDate) {
      params.data_start = formattedDate;
      params.data_end = formattedDate;
    }

    const response = await api.get('/api/financial/list-invoice', { params });

    const content = response.data.content || [];
    return content[0] || null;

  } catch (error) {
    log.warn(`Não foi possível buscar invoice ${invoiceId}`, { erro: error.message });
    return null;
  }
}

/**
 * Fetches invoice items (procedures)
 * @param {string|number} invoiceId - Invoice ID
 * @param {string} transactionDate - Transaction date (DD/MM/YYYY)
 * @returns {Promise<Array>} List of items with procedimento_id and value
 */
export async function fetchInvoiceItems(invoiceId, transactionDate) {
  const invoice = await fetchInvoiceDetails(invoiceId, transactionDate);
  return invoice?.itens || [];
}

/**
 * Fetches the proposal/creation date of an invoice
 * Uses a wide date range since the invoice may be much older than the payment
 * @param {string|number} invoiceId - Invoice ID
 * @returns {Promise<string|null>} Proposal date in DD/MM/YYYY format, or null
 */
export async function fetchInvoiceProposalDate(invoiceId) {
  try {
    const response = await api.get('/api/financial/list-invoice', {
      params: {
        invoice_id: invoiceId,
        tipo_transacao: 'C',
        data_start: '01-01-2020',
        data_end: '31-12-2030',
      },
    });

    const invoice = (response.data.content || [])[0];
    const detalhes = invoice?.detalhes;
    if (!detalhes || !detalhes.length) return null;

    // detalhes[0].data comes as DD-MM-YYYY
    const rawDate = detalhes[0].data;
    if (!rawDate) return null;

    return rawDate.replace(/-/g, '/');
  } catch (error) {
    log.warn(`Could not fetch proposal date for invoice ${invoiceId}`, { error: error.message });
    return null;
  }
}

/**
 * Fetches procedure information by ID
 * @param {string|number} procedureId - Procedure ID
 * @returns {Promise<Object|null>} Procedure data
 */
export async function fetchProcedure(procedureId) {
  try {
    const response = await api.get('/api/procedures/list', {
      params: { procedimento_id: procedureId }
    });

    const content = response.data.content || [];
    return content[0] || null;

  } catch (error) {
    log.warn(`Não foi possível buscar procedimento ${procedureId}`, { erro: error.message });
    return null;
  }
}

// Procedure cache to avoid multiple requests
const proceduresCache = new Map();

/**
 * Fetches procedure information by ID (with cache)
 * @param {string|number} procedureId - Procedure ID
 * @returns {Promise<Object|null>} Procedure data
 */
export async function fetchProcedureWithCache(procedureId) {
  const id = String(procedureId);

  if (proceduresCache.has(id)) {
    return proceduresCache.get(id);
  }

  const procedure = await fetchProcedure(procedureId);
  if (procedure) {
    proceduresCache.set(id, procedure);
  }
  return procedure;
}

/**
 * Fetches procedure groups and builds ID → column mapping
 * Uses official Feegow groups as base, with overrides for Group 3 (Consultations)
 * @returns {Promise<Map<number, string>>} Mapping of procedimento_id → column
 */
// Proposals cache to avoid multiple requests per patient+date
const proposalsCache = new Map();

let groupsCache = null;

export async function fetchGroupMapping() {
  if (groupsCache) return groupsCache;

  // Group → default column mapping
  const GROUP_TO_COLUMN = {
    1: 'IMPLANTE',      // Hormonal implants
    2: 'APLICACOES',    // Applications
    // 3: Consultations → needs individual override
    4: 'SORO_DR',       // IV therapy
    5: 'EXTRA',         // Exams
  };

  // Overrides for specific procedures (mainly Group 3)
  const OVERRIDES = {
    // Tirzepatida (in Group 2 but has its own column)
    58: 'TIRZEPATIDA',
    144: 'TIRZEPATIDA',  // Tirzepatida 90mg/3.6ml
    // Group 3 - Consultations (each goes to a specific column)
    16: 'AVALIACAO',     // 1st Consultation (Initial Medical Evaluation)
    17: 'ONLINE',         // Online Consultation - Dr Victor
    18: 'NUTRIS',         // Online Consultation - Nutritionist (nutri > online)
    19: 'ONLINE',         // Online Consultation - Dr Luiz
    20: null,            // Chat/Questions - DISCARD
    22: 'NUTRIS',        // Single Nutritional Consultation
    23: 'AVALIACAO',     // Follow-up - Dr (retorno → avaliação)
    24: 'NUTRIS',        // Follow-up - Nutritionist
    104: 'DR_RIGATTI',   // Medical Repurchase Consultation
    105: 'DR_RIGATTI',   // 1st Medical Protocol Consultation
    106: 'DR_RIGATTI',   // 2nd Medical Protocol Consultation
    107: 'DR_RIGATTI',   // 3rd Medical Protocol Consultation
    108: 'DR_RIGATTI',   // 4th Medical Protocol Consultation
    109: 'NUTRIS',       // 1st Nutritional Protocol Consultation
    110: 'NUTRIS',       // 2nd Nutritional Protocol Consultation
    111: 'NUTRIS',       // 3rd Nutritional Protocol Consultation
    112: 'NUTRIS',       // 4th Nutritional Protocol Consultation
    137: 'ONLINE',       // Online Repurchase
  };

  const mapping = new Map();

  try {
    const response = await api.get('/api/procedures/groups');
    const groups = response.data.content || [];

    for (const group of groups) {
      const groupColumn = GROUP_TO_COLUMN[group.id] || null;

      for (const proc of group.procedimentos) {
        // Override takes priority over group
        if (proc.id in OVERRIDES) {
          const column = OVERRIDES[proc.id];
          if (column) mapping.set(proc.id, column);
          // If column is null, don't map (discarded)
        } else if (groupColumn) {
          mapping.set(proc.id, groupColumn);
        }
      }
    }

    log.info(`Mapeamento de grupos carregado: ${mapping.size} procedimentos`);
  } catch (error) {
    log.warn('Não foi possível carregar grupos de procedimentos', { erro: error.message });
  }

  groupsCache = mapping;
  return mapping;
}

/**
 * Fetches executed proposals for a patient on a specific date
 * @param {string|number} pacienteId - Patient ID
 * @param {string} date - Date in DD/MM/YYYY format
 * @returns {Promise<Array>} List of executed proposals with procedure items
 */
export async function fetchPatientProposals(pacienteId, date) {
  const cacheKey = `${pacienteId}_${date}`;
  if (proposalsCache.has(cacheKey)) {
    return proposalsCache.get(cacheKey);
  }

  try {
    const formattedDate = date.replace(/\//g, '-');
    const response = await api.get('/api/proposal/list', {
      params: {
        paciente_id: pacienteId,
        data_inicio: formattedDate,
        data_fim: formattedDate,
      },
    });

    const proposals = (response.data.content || [])
      .filter(p => p.status === 'Executada');

    proposalsCache.set(cacheKey, proposals);
    log.debug(`${proposals.length} executed proposal(s) for patient ${pacienteId} on ${date}`);
    return proposals;
  } catch (error) {
    log.warn(`Could not fetch proposals for patient ${pacienteId}`, { error: error.message });
    proposalsCache.set(cacheKey, []);
    return [];
  }
}

// Track which proposal IDs have already been used to prevent duplication
// when multiple transactions from the same patient match the same proposal
const usedProposalIds = new Set();

/**
 * Finds the best matching proposal for a transaction
 * Requires a confident match (by procedure ID or invoice items), never guesses
 * @param {Object} transaction - Transaction data
 * @param {Array} proposals - List of proposals (already filtered to unused ones)
 * @returns {Promise<Object|null>} Matched proposal or null
 */
async function findMatchingProposal(transaction, proposals) {
  if (proposals.length === 0) return null;

  // Determine if multi-procedure
  const reportNames = transaction.NomeProcedimento
    ? transaction.NomeProcedimento.split(/,(?=\s*[A-Za-zÀ-ú])|,(?=\s*\d+º)/).map(n => n.trim()).filter(n => n)
    : [];
  const isMultiProcedure = reportNames.length > 1;

  // Single procedure: match by ProcedimentoID
  if (!isMultiProcedure && transaction.ProcedimentoID) {
    const candidates = proposals.filter(p =>
      (p.procedimentos?.data || []).some(proc =>
        proc.procedimento_id === transaction.ProcedimentoID
      )
    );

    if (candidates.length === 1) return candidates[0];
    if (candidates.length > 1) {
      candidates.sort((a, b) =>
        Math.abs(a.value - transaction.Value) - Math.abs(b.value - transaction.Value)
      );
      return candidates[0];
    }
    // No proposal contains this procedure ID → no match
    return null;
  }

  // Multi-procedure: match by invoice item procedure IDs (at least 2 matches)
  if (isMultiProcedure && transaction.NumeroInvoice) {
    const items = await fetchInvoiceItems(transaction.NumeroInvoice, transaction.Data);
    if (items.length) {
      const invoiceProcIds = new Set(items.map(i => i.procedimento_id));
      let bestProposal = null;
      let bestMatchCount = 0;

      for (const proposal of proposals) {
        const propProcIds = (proposal.procedimentos?.data || []).map(p => p.procedimento_id);
        const matchCount = propProcIds.filter(id => invoiceProcIds.has(id)).length;
        if (matchCount > bestMatchCount) {
          bestMatchCount = matchCount;
          bestProposal = proposal;
        }
      }

      if (bestProposal && bestMatchCount >= 2) return bestProposal;
    }
  }

  // No confident match found
  return null;
}

/**
 * Enriches a transaction with detailed invoice items
 * @param {Object} transaction - Transaction from financial-movement report
 * @returns {Promise<Object>} Transaction with detailed items
 */
export async function enrichTransactionWithItems(transaction) {
  const { NumeroInvoice, NomeProcedimento } = transaction;

  // Check if this is a payment for an old proposal (installment)
  // If the proposal date is more than 30 days before the transaction, it's payment-only
  if (NumeroInvoice) {
    const proposalDate = await fetchInvoiceProposalDate(NumeroInvoice);
    if (proposalDate && proposalDate !== transaction.Data) {
      const [pDay, pMonth, pYear] = proposalDate.split('/').map(Number);
      const [tDay, tMonth, tYear] = transaction.Data.split('/').map(Number);
      const proposalMs = new Date(pYear, pMonth - 1, pDay).getTime();
      const transactionMs = new Date(tYear, tMonth - 1, tDay).getTime();
      const daysDiff = Math.abs(transactionMs - proposalMs) / (1000 * 60 * 60 * 24);

      if (daysDiff > 30) {
        log.debug(`Invoice ${NumeroInvoice}: proposal ${proposalDate} is ${Math.round(daysDiff)} days before transaction ${transaction.Data} → payment only`);
        return {
          ...transaction,
          paymentOnly: true,
          fonteItens: 'payment-only',
          detailedItems: [],
        };
      } else {
        log.debug(`Invoice ${NumeroInvoice}: proposal ${proposalDate} is ${Math.round(daysDiff)} days before transaction ${transaction.Data} → recent, processing normally`);
      }
    }
  }

  // Try to use proposal data for accurate values (with quantities)
  if (transaction.PacienteID) {
    const proposals = await fetchPatientProposals(transaction.PacienteID, transaction.Data);

    // Filter out proposals already used by another transaction from this patient
    const availableProposals = proposals.filter(p => !usedProposalIds.has(p.proposal_id));

    if (availableProposals.length) {
      const matchedProposal = await findMatchingProposal(transaction, availableProposals);

      if (matchedProposal?.procedimentos?.data?.length) {
        usedProposalIds.add(matchedProposal.proposal_id);

        const detailedItems = matchedProposal.procedimentos.data.map(proc => ({
          procedimento_id: proc.procedimento_id,
          nome: proc.nome,
          valor: proc.valor * (proc.quantidade || 1), // Proposal values are in reais, multiply by qty
          quantidade: proc.quantidade || 1,
          desconto: Number(proc.desconto) || 0,
          acrescimo: Number(proc.acrescimo) || 0,
        }));

        log.debug(`Invoice ${NumeroInvoice || '?'}: using proposal #${matchedProposal.proposal_id} (${detailedItems.length} items, total R$ ${matchedProposal.value})`);

        return {
          ...transaction,
          fonteItens: 'proposal',
          detailedItems,
        };
      }
    }
  }

  // Fallback: use report/invoice data when no proposal is available

  // Procedure names from report (reliable source for classification)
  // Smart split by comma:
  // - Splits when followed by letter: "Implante,Soroterapia"
  // - Splits when followed by digit+º: "Im,3º Consulta"
  // - Does NOT split decimal commas: "Estradiol 12,5Mg" or "Tirzepatida 40Mg/1,6Ml"
  const reportNames = NomeProcedimento
    ? NomeProcedimento.split(/,(?=\s*[A-Za-zÀ-ú])|,(?=\s*\d+º)/).map(n => n.trim()).filter(n => n)
    : [];

  // Single procedure: use data directly from report
  if (reportNames.length <= 1) {
    return {
      ...transaction,
      fonteItens: 'report-single',
      detailedItems: [{
        procedimento_id: transaction.ProcedimentoID,
        nome: NomeProcedimento || 'Procedimento',
        valor: transaction.Value,
        quantidade: transaction.Quantidade || 1,
      }]
    };
  }

  // Multiple procedures: fetch invoice to get individual values
  if (!NumeroInvoice) {
    // No invoice, distribute value equally among procedures
    const valuePerItem = transaction.Value / reportNames.length;
    return {
      ...transaction,
      fonteItens: 'report-distributed',
      detailedItems: reportNames.map(nome => ({
        procedimento_id: null,
        nome,
        valor: valuePerItem,
        quantidade: 1,
      })),
    };
  }

  // Fetch invoice items to get individual values
  const items = await fetchInvoiceItems(NumeroInvoice, transaction.Data);

  if (!items.length) {
    log.warn(`Invoice ${NumeroInvoice} sem itens detalhados, distribuindo valor igualmente`);
    const valuePerItem = transaction.Value / reportNames.length;
    return {
      ...transaction,
      fonteItens: 'report-distributed',
      detailedItems: reportNames.map(nome => ({
        procedimento_id: null,
        nome,
        valor: valuePerItem,
        quantidade: 1,
      })),
    };
  }

  // Use invoice items with procedimento_id for precise classification
  if (items.length !== reportNames.length) {
    log.debug(`Invoice ${NumeroInvoice}: ${items.length} itens vs ${reportNames.length} nomes no relatório`);
  }

  const detailedItems = [];
  for (const item of items) {
    const proc = await fetchProcedureWithCache(item.procedimento_id);
    const nome = proc?.nome || `Procedimento #${item.procedimento_id}`;

    detailedItems.push({
      procedimento_id: item.procedimento_id,
      nome,
      valor: item.valor / 100, // Invoice returns in cents
      quantidade: item.quantidade || 1,
      desconto: (item.desconto || 0) / 100, // Invoice returns in cents
    });
  }

  return {
    ...transaction,
    fonteItens: 'invoice',
    detailedItems,
  };
}

/**
 * Fetches transactions and enriches with detailed items
 * @param {string} startDate - Start date in DD/MM/YYYY format
 * @param {string} endDate - End date in DD/MM/YYYY format
 * @param {object} options - Additional filter options
 * @returns {Promise<Array>} List of enriched transactions
 */
export async function fetchDetailedTransactions(startDate, endDate, options = {}) {
  const transactions = await fetchTransactions(startDate, endDate, options);

  // Reset used proposals tracking for new batch
  usedProposalIds.clear();

  log.info(`Enriquecendo ${transactions.length} movimentações com itens detalhados...`);

  // Process sequentially to not overload Feegow API
  const enrichedTransactions = [];
  for (let i = 0; i < transactions.length; i++) {
    const mov = transactions[i];
    const enriched = await enrichTransactionWithItems(mov);
    enrichedTransactions.push(enriched);

    // Progress log every 10 transactions
    if ((i + 1) % 10 === 0) {
      log.info(`Progresso: ${i + 1}/${transactions.length} movimentações enriquecidas`);
    }
  }

  log.info('Enriquecimento concluído');

  return enrichedTransactions;
}

/**
 * Cache for patient appointments (key: patientId, value: appointments array)
 * Cleared each run cycle to avoid stale data
 */
const appointmentsCache = new Map();

/**
 * Clears the appointments cache (call at start of each run cycle)
 */
export function clearAppointmentsCache() {
  appointmentsCache.clear();
}

/**
 * Fetches patient appointments in a ±7 day window around the given date
 * Uses GET /api/appoints/search endpoint with cache per patient
 * Retries up to 3 times on transient errors (409, 5xx, network)
 * @param {string|number} patientId - Patient ID
 * @param {string} dateStr - Date in DD/MM/YYYY or YYYY-MM-DD format
 * @returns {Promise<{appointments: Array, apiError: boolean}>}
 */
export async function fetchPatientAppointments(patientId, dateStr) {
  const cacheKey = String(patientId);
  if (appointmentsCache.has(cacheKey)) {
    return appointmentsCache.get(cacheKey);
  }

  // Parse date to create ±7 day window
  let dateObj;
  if (dateStr.includes('/')) {
    const [d, m, y] = dateStr.split('/').map(Number);
    dateObj = new Date(y, m - 1, d);
  } else {
    const [y, m, d] = dateStr.split('-').map(Number);
    dateObj = new Date(y, m - 1, d);
  }

  const startDate = new Date(dateObj);
  startDate.setDate(startDate.getDate() - 7);
  const endDate = new Date(dateObj);
  endDate.setDate(endDate.getDate() + 7);

  const fmt = (dt) => {
    const dd = String(dt.getDate()).padStart(2, '0');
    const mm = String(dt.getMonth() + 1).padStart(2, '0');
    const yyyy = dt.getFullYear();
    return `${dd}-${mm}-${yyyy}`;
  };

  const MAX_RETRIES = 3;
  for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    try {
      const response = await api.get('/api/appoints/search', {
        params: {
          paciente_id: patientId,
          data_start: fmt(startDate),
          data_end: fmt(endDate),
        },
      });

      const result = { appointments: response.data.content || [], apiError: false };
      appointmentsCache.set(cacheKey, result);
      return result;

    } catch (error) {
      const status = error.response?.status;
      const isRetryable = !status || status === 409 || status >= 500;

      if (isRetryable && attempt < MAX_RETRIES) {
        log.warn(`Agendamentos paciente ${patientId}: erro ${status || 'network'}, tentativa ${attempt}/${MAX_RETRIES}`);
        await new Promise(r => setTimeout(r, 500 * attempt));
        continue;
      }

      log.warn(`Não foi possível buscar agendamentos do paciente ${patientId} após ${attempt} tentativa(s)`, {
        erro: error.message,
        status,
      });
      const errorResult = { appointments: [], apiError: true };
      appointmentsCache.set(cacheKey, errorResult);
      return errorResult;
    }
  }
}

/**
 * Checks if patient has an appointment on the exact processing date
 * Also detects online consultations from nearby appointments (notas or telemedicina flag)
 * @param {string|number} patientId - Patient ID
 * @param {string} processingDate - Date being processed (DD/MM/YYYY or YYYY-MM-DD)
 * @returns {Promise<{hasAppointmentToday: boolean, isOnlineConsultation: boolean, apiError: boolean}>}
 */
export async function checkAppointmentDate(patientId, processingDate) {
  const { appointments, apiError } = await fetchPatientAppointments(patientId, processingDate);

  if (apiError) {
    return { hasAppointmentToday: false, isOnlineConsultation: false, apiError: true };
  }

  // Normalize processing date to DD-MM-YYYY for comparison
  let procDateNorm;
  if (processingDate.includes('/')) {
    procDateNorm = processingDate.replace(/\//g, '-');
  } else {
    const [y, m, d] = processingDate.split('-');
    procDateNorm = `${d}-${m}-${y}`;
  }

  const hasToday = appointments.some(apt => apt.data === procDateNorm);

  // Check if any nearby appointment is an online consultation
  // "Agendamento Online" = booking channel (AI agents), NOT an online consultation
  // "Consulta ONLINE" in notes = actual online consultation
  const isOnline = appointments.some(apt =>
    apt.telemedicina === true ||
    (apt.notas && apt.notas.toLowerCase().includes('consulta online'))
  );

  return { hasAppointmentToday: hasToday, isOnlineConsultation: isOnline, apiError: false };
}

/**
 * Formats date to Feegow standard (DD/MM/YYYY)
 * @param {Date} date
 * @returns {string}
 */
export function formatDateFeegow(date) {
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  return `${day}/${month}/${year}`;
}

/**
 * Converts ISO date (YYYY-MM-DD) to Feegow format (DD/MM/YYYY)
 * @param {string} isoDate - Date in YYYY-MM-DD format
 * @returns {string}
 */
export function isoToFeegow(isoDate) {
  const [year, month, day] = isoDate.split('-');
  return `${day}/${month}/${year}`;
}

/**
 * Converts Feegow date (DD/MM/YYYY) to ISO (YYYY-MM-DD)
 * @param {string} feegowDate - Date in DD/MM/YYYY format
 * @returns {string}
 */
export function feegowToIso(feegowDate) {
  const [day, month, year] = feegowDate.split('/');
  return `${year}-${month}-${day}`;
}

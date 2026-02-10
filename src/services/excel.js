import ExcelJS from 'exceljs';
import path from 'path';
import { fileURLToPath } from 'url';
import { createLogger, logTable } from './logger.js';
import { fetchGroupMapping, checkAppointmentDate } from './feegow.js';
import { saveHistory } from './historico.js';

const log = createLogger('Excel');

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ONEDRIVE_PATH = path.resolve(__dirname, '../../onedrive');

// Portuguese month names for dynamic tab resolution
const MONTH_NAMES_PT = [
  'JANEIRO', 'FEVEREIRO', 'MARÇO', 'ABRIL', 'MAIO', 'JUNHO',
  'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO',
];

/**
 * Resolves the Excel sheet name from a date string.
 * If EXCEL_ABA env var is set, uses it as override (useful for testing).
 * Otherwise derives from date: "MÊSNAME ANO" (e.g., "FEVEREIRO 2026").
 * @param {string} date - Date in YYYY-MM-DD or DD/MM/YYYY format
 * @returns {string} Sheet name
 */
export function resolveSheetName(date) {
  if (process.env.EXCEL_ABA) {
    return process.env.EXCEL_ABA;
  }

  let month, year;
  if (date.includes('-')) {
    [year, month] = date.split('-');
  } else if (date.includes('/')) {
    const parts = date.split('/');
    month = parts[1];
    year = parts[2];
  } else {
    log.warn(`Formato de data não reconhecido: ${date}, usando fallback`);
    return 'Testes';
  }

  const monthIndex = parseInt(month, 10) - 1;
  const monthName = MONTH_NAMES_PT[monthIndex];

  if (!monthName) {
    log.warn(`Mês inválido: ${month}, usando fallback`);
    return 'Testes';
  }

  return `${monthName} ${year}`;
}

// Initial row for data insertion (after header)
// If defined, skips last row search and uses this value
const INITIAL_ROW = process.env.EXCEL_LINHA_INICIAL ? parseInt(process.env.EXCEL_LINHA_INICIAL) : null;

// Number of header rows to copy when creating a new sheet from template
const HEADER_ROWS = 7;

// Red color for cells that need validation
const PENDING_VALIDATION_COLOR = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFFF6B6B' },
};

// Default font
const CELL_FONT = { name: 'Arial', size: 11 };

// Default column background colors (for zero-value cells)
const FILL_YELLOW = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE598' } };
const FILL_PINK = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF66FF' } };
const FILL_BLUE = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDEEAF6' } };

function getDefaultFill(column) {
  // A-D: yellow
  if (['A', 'B', 'C', 'D'].includes(column)) return FILL_YELLOW;
  // E: pink
  if (column === 'E') return FILL_PINK;
  // F-P: yellow
  if (['F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P'].includes(column)) return FILL_YELLOW;
  // Q-AF: blue
  const blueColumns = ['Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF'];
  if (blueColumns.includes(column)) return FILL_BLUE;
  return null;
}

// Dotted borders for cells with data
const CELL_BORDER = {
  top: { style: 'dotted' },
  left: { style: 'dotted' },
  bottom: { style: 'dotted' },
  right: { style: 'dotted' },
};

// Spreadsheet column mapping (based on 2026 spreadsheet Row 7 headers)
const COLUMNS = {
  // Identification
  DATA: 'A',
  NOME: 'B',
  VENDA: 'C',

  // Procedures (treatment types)
  AVALIACAO: 'D',
  TRATAMENTO: 'E',
  DESCONTO: 'F',
  DR_RIGATTI: 'G',
  IMPLANTE: 'H',
  SORO_DR: 'I',
  TIRZEPATIDA: 'J',
  APLICACOES: 'K',
  ONLINE: 'L',
  COMISSAO: 'M',
  NUTRIS: 'N',
  EXTRA: 'O',
  ESTORNO_PROC: 'P',

  // Payment
  DINHEIRO: 'Q',
  SICOOB: 'R',
  SAFRA: 'S',          // Safra Pay - card
  PIX_SAFRA: 'T',      // PIX SAFRA (Banco Safra - PIX)
  INFINITE: 'U',       // InfinitePay
  PIX: 'V',            // TRANSF/PIX (Banco Cora - PIX)
  CHEQUE: 'W',
  BOLETO: 'X',

  // Observations
  OBSERVACAO: 'AG',     // Observations (e.g., "Tratamento em andamento")

  // Fees and others
  JUROS_CARTAO: 'Z',   // JUROS CARTÃO
  TX_SICOOB: 'AA',     // TX SICCOB
  TX_INFINITE: 'AB',   // TX INFINITE
  TX_SAFRA: 'AC',      // TX SAFRA
  ESTORNO: 'AD',       // ESTORNO
  NP_ABERTA: 'AE',     // NP ABERTA
  BOLETO_2: 'AF',      // BOLETO
};

// Reverse mapping: column letter → readable name (for history file)
const COLUMN_NAMES = Object.fromEntries(
  Object.entries(COLUMNS).map(([name, col]) => [col, name])
);

/**
 * Procedure name to column mapping
 * Keys are patterns (case insensitive) to search in procedure name
 */
const PROCEDURE_MAPPING = {
  // Nutritionist (BEFORE 'consulta' - "Consulta Nutricional" goes to NUTRIS)
  'nutricional': COLUMNS.NUTRIS,

  // Evaluation (initial evaluation only)
  'avalia': COLUMNS.AVALIACAO,

  // Dr Rigatti (Medical Protocol and Repurchase Consultations)
  'rigatti': COLUMNS.DR_RIGATTI,
  'recompra': COLUMNS.DR_RIGATTI,
  'protocolo': COLUMNS.DR_RIGATTI,

  // Online (before generic consultation fallback)
  'online': COLUMNS.ONLINE,

  // Generic consultation (fallback - AFTER more specific ones)
  'consulta': COLUMNS.AVALIACAO,

  // Dr Victor (no dedicated column in 2026 spreadsheet)
  'victor': COLUMNS.EXTRA,

  // Implants
  'implante': COLUMNS.IMPLANTE,

  // IV Therapy
  'soro': COLUMNS.SORO_DR,
  'soroterapia': COLUMNS.SORO_DR,
  'noripurum': COLUMNS.SORO_DR,

  // Tirzepatida
  'tirzepatida': COLUMNS.TIRZEPATIDA,

  // Applications (injectables)
  'injetável': COLUMNS.APLICACOES,
  'injetavel': COLUMNS.APLICACOES,
  'cipionato': COLUMNS.APLICACOES,
  'testosterona': COLUMNS.APLICACOES,
  'blend': COLUMNS.APLICACOES,
  'gh ': COLUMNS.APLICACOES,
  'genotropin': COLUMNS.APLICACOES,
  'omnitrope': COLUMNS.APLICACOES,

  // Nutritionist
  'nutri': COLUMNS.NUTRIS,

  // Proposed Treatment (Dr Rigatti)
  'tratamento proposto': COLUMNS.DR_RIGATTI,

  // Interest/Extras
  'juros': COLUMNS.JUROS_CARTAO,

  // Reversal
  'estorno': COLUMNS.ESTORNO_PROC,
};

/**
 * Payment method to column mapping (fallback when AccountName doesn't match)
 * Based on Feegow FormaPagamentoID
 */
const PAYMENT_MAPPING = {
  1: COLUMNS.DINHEIRO,      // Cash
  2: COLUMNS.CHEQUE,        // Check
  3: COLUMNS.PIX,           // Transfer
  4: COLUMNS.BOLETO,        // Bank slip
  5: COLUMNS.PIX,           // DOC
  6: COLUMNS.PIX,           // TED
  7: COLUMNS.PIX,           // Bank Transfer
  8: COLUMNS.INFINITE,      // Credit Card (default → InfinitePay)
  9: COLUMNS.INFINITE,      // Debit Card (default → InfinitePay)
  10: COLUMNS.INFINITE,     // Credit Card
  15: COLUMNS.PIX,          // PIX
};

// Group mapping (dynamically loaded from API)
let groupMapping = null;

/**
 * Loads the group mapping from the API (needed before calling buildRow externally)
 */
async function loadGroupMapping() {
  groupMapping = await fetchGroupMapping();
  return groupMapping;
}

/**
 * Identifies the correct column for a procedure
 * Priority: 1) Procedure ID (via API groups), 2) Name (pattern matching)
 * @param {string} procedureName - Procedure name
 * @param {number|null} procedureId - Procedure ID (optional)
 * @returns {string|null} Column letter (null = discard)
 */
function identifyProcedureColumn(procedureName, procedureId = null) {
  return classifyProcedure(procedureName, procedureId).column;
}

/**
 * Classifies a procedure and returns both column and classification details
 * @param {string} procedureName - Procedure name
 * @param {number|null} procedureId - Procedure ID (optional)
 * @returns {{ column: string|null, classificadoPor: string }} Column and method
 */
function classifyProcedure(procedureName, procedureId = null) {
  // 1) Try by ID (more precise)
  if (procedureId && groupMapping) {
    const columnById = groupMapping.get(Number(procedureId));
    if (columnById !== undefined) {
      const col = COLUMNS[columnById] || columnById;
      return { column: col, classificadoPor: 'id' };
    }
  }

  // 2) Fallback: pattern matching by name
  const normalizedName = procedureName.toLowerCase();

  for (const [pattern, column] of Object.entries(PROCEDURE_MAPPING)) {
    if (normalizedName.includes(pattern.toLowerCase())) {
      return { column, classificadoPor: 'nome' };
    }
  }

  // If not found, goes to EXTRA (column for uncategorized procedures)
  return { column: COLUMNS.EXTRA, classificadoPor: 'fallback' };
}

/**
 * Identifies the correct column for payment method
 * Combines AccountName + FormaPagamentoID for precise classification
 * @param {Object} transaction - Feegow transaction
 * @returns {string} Column letter
 */
function identifyPaymentColumn(transaction) {
  const { FormaPagamentoID, AccountName } = transaction;
  const account = (AccountName || '').toLowerCase();

  // Cash
  if (account.includes('caixa') || FormaPagamentoID === 1) {
    return COLUMNS.DINHEIRO;
  }

  // Sicoob (credit or debit on card machine)
  if (account.includes('sicoob')) {
    return COLUMNS.SICOOB;
  }

  // Safra Pay (Safra card machine) → SAFRA column
  if (account.includes('safra pay')) {
    return COLUMNS.SAFRA;
  }

  // Banco Safra (bank account) → depends on FormaPagamentoID
  if (account.includes('safra')) {
    if (FormaPagamentoID === 4) return COLUMNS.BOLETO;       // Bank slip via Banco Safra
    return COLUMNS.PIX_SAFRA;                                 // PIX via Banco Safra
  }

  // InfinitePay (note: API has typo "InfintePay" without the 'i')
  if (account.includes('infinte') || account.includes('infinite')) {
    return COLUMNS.INFINITE;
  }

  // Banco Cora → PIX/Transfer
  if (account.includes('cora')) {
    return COLUMNS.PIX;
  }

  // Fallback by FormaPagamentoID
  return PAYMENT_MAPPING[FormaPagamentoID] || COLUMNS.PGTO_OUTROS;
}

/**
 * Formats value to standard number
 * @param {number} value - Numeric value
 * @returns {number} Formatted value
 */
function formatValue(value) {
  return typeof value === 'number' ? value : 0;
}

/**
 * Creates a new sheet by copying the header area (rows 1-7) from the most recent existing sheet.
 * Copies all formatting, formulas, merges, column widths, and row heights.
 * Updates B3:B5 merged cell with the first day of the new month.
 * @param {ExcelJS.Workbook} workbook - The workbook
 * @param {string} sheetName - Name for the new sheet (e.g., "MARÇO 2026")
 * @param {string} date - Reference date (YYYY-MM-DD or DD/MM/YYYY)
 * @returns {ExcelJS.Worksheet} The newly created worksheet
 */
function createSheetFromTemplate(workbook, sheetName, date) {
  // Find the most recent existing month sheet as template
  const existingSheets = workbook.worksheets.filter(ws => {
    const name = ws.name.toUpperCase();
    // Match only real month sheets (e.g., "FEVEREIRO 2026"), skip test sheets
    return MONTH_NAMES_PT.some(m => name.startsWith(m)) && !name.includes('TESTE');
  });

  if (existingSheets.length === 0) {
    throw new Error('Nenhuma aba de mês encontrada para usar como modelo');
  }

  // Use the last month sheet as template
  const template = existingSheets[existingSheets.length - 1];
  log.info(`Criando aba "${sheetName}" usando "${template.name}" como modelo`);

  const newSheet = workbook.addWorksheet(sheetName);

  // Copy column widths
  const allCols = 'A B C D E F G H I J K L M N O P Q R S T U V W X Y Z AA AB AC AD AE AF AG AH'.split(' ');
  for (const colLetter of allCols) {
    const srcCol = template.getColumn(colLetter);
    if (srcCol.width) {
      newSheet.getColumn(colLetter).width = srcCol.width;
    }
  }

  // Copy header rows (1-7) with all formatting and values
  for (let r = 1; r <= HEADER_ROWS; r++) {
    const srcRow = template.getRow(r);
    const dstRow = newSheet.getRow(r);

    // Copy row height
    if (srcRow.height) dstRow.height = srcRow.height;

    for (const colLetter of allCols) {
      const srcCell = template.getCell(`${colLetter}${r}`);
      const dstCell = newSheet.getCell(`${colLetter}${r}`);

      // Skip merged slave cells (they get their value from the master)
      if (srcCell.isMerged && srcCell.master !== srcCell) continue;

      // Copy formula or value
      if (srcCell.formula) {
        dstCell.value = { formula: srcCell.formula };
      } else if (srcCell.value !== null && srcCell.value !== undefined) {
        dstCell.value = srcCell.value;
      }

      // Copy style (font, fill, alignment, border, numFmt)
      if (srcCell.style) {
        dstCell.style = JSON.parse(JSON.stringify(srcCell.style));
      }
    }
  }

  // Copy merges from header area
  const merges = template.model.merges || [];
  for (const merge of merges) {
    const match = merge.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    if (match && parseInt(match[2]) <= HEADER_ROWS && parseInt(match[4]) <= HEADER_ROWS) {
      try { newSheet.mergeCells(merge); } catch (e) { /* already merged */ }
    }
  }

  // Update B3 with the first day of the new month
  let month, year;
  if (date.includes('-')) {
    [year, month] = date.split('-');
  } else if (date.includes('/')) {
    const parts = date.split('/');
    month = parts[1];
    year = parts[2];
  }
  const firstDayOfMonth = new Date(parseInt(year), parseInt(month) - 1, 1);
  newSheet.getCell('B3').value = firstDayOfMonth;

  log.info(`Aba "${sheetName}" criada com sucesso`);
  return newSheet;
}

/**
 * Updates the Excel spreadsheet with transactions
 * @param {Array} transactions - List of Feegow transactions (with detailedItems)
 * @param {string} date - Reference date
 */
export async function updateSpreadsheet(transactions, date) {
  const startTime = log.start('Atualização da planilha');

  const sheetName = resolveSheetName(date);
  const filePath = path.join(ONEDRIVE_PATH, process.env.EXCEL_FILE_PATH);
  log.debug(`Arquivo: ${filePath}`);
  log.debug(`Aba: ${sheetName}`);

  const workbook = new ExcelJS.Workbook();

  try {
    await workbook.xlsx.readFile(filePath);
    log.debug('Planilha carregada com sucesso');
  } catch (error) {
    log.error(`Não foi possível abrir a planilha: ${filePath}`, { erro: error.message });
    throw new Error(`Não foi possível abrir a planilha: ${filePath}`);
  }

  // Find sheet by resolved name, or create from template if it doesn't exist
  let worksheet = workbook.getWorksheet(sheetName);

  if (!worksheet) {
    log.warn(`Aba "${sheetName}" não encontrada — criando automaticamente`);
    worksheet = createSheetFromTemplate(workbook, sheetName, date);
  }

  log.info(`Usando aba: ${worksheet.name}`);

  // Load procedure group mapping from API
  groupMapping = await fetchGroupMapping();

  // Determine initial row for insertion
  // Always scan column B dynamically; INITIAL_ROW sets the minimum starting point
  const startScan = INITIAL_ROW || 8;
  let lastRow = startScan - 1; // default: insert at startScan
  for (let r = startScan; r <= worksheet.rowCount; r++) {
    const nameCell = worksheet.getCell(`B${r}`);
    const name = nameCell.value;
    if (!name || String(name).trim() === '') {
      break; // First empty row found → insert here
    }
    lastRow = r;
  }
  log.debug(`Última linha com nome de paciente: ${lastRow} → inserindo a partir de ${lastRow + 1}`);

  // Group transactions by patient + date to consolidate into one row
  const groupedTransactions = groupTransactions(transactions);
  log.info(`${groupedTransactions.length} grupos de movimentações para inserir`);

  // Prepare target rows: unmerge any merged cells and clear values/formatting
  const finalRow = lastRow + groupedTransactions.length;
  const usedColumns = Object.values(COLUMNS);
  log.debug(`Preparando linhas ${lastRow + 1} a ${finalRow} (colunas: ${usedColumns.length})`);

  // First: unmerge any merged cells in the target range
  const merges = worksheet.model.merges || [];
  for (const merge of merges) {
    // merge format: "C408:N408" - check if it overlaps our target rows
    const match = merge.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    if (match) {
      const mergeStartRow = parseInt(match[2]);
      const mergeEndRow = parseInt(match[4]);
      if (mergeEndRow >= lastRow + 1 && mergeStartRow <= finalRow) {
        try { worksheet.unMergeCells(merge); } catch (e) { /* ignore */ }
      }
    }
  }

  // Note: do NOT clear all cells upfront - only modify cells we actually write to.
  // This preserves pre-existing formatting (e.g., orange background) in unused cells.

  // Add each grouped transaction as a new row
  const insertedRows = [];

  for (const group of groupedTransactions) {
    lastRow++;
    const excelRow = worksheet.getRow(lastRow);

    // Build row data
    const { rowData, classificationDetails } = await buildRow(group, date);

    // Detailed log for each row
    log.debug(`Inserindo linha ${lastRow}`, {
      patient: group.patientName,
      date: group.date,
      totalItems: group.detailedItems.length,
      totalTransactions: group.transactions.length,
    });

    // Column type classification
    const textColumns = [COLUMNS.NOME, COLUMNS.VENDA, COLUMNS.OBSERVACAO];

    // Fill cells with data
    for (const [column, value] of Object.entries(rowData)) {
      const cell = worksheet.getCell(`${column}${lastRow}`);

      // Reset cell before writing (clear any pre-existing style/value)
      cell.value = null;
      cell.style = {};

      if (column === COLUMNS.DATA) {
        // Date column: write as Date object with dd/mm/yy format, left-aligned
        cell.value = value;
        cell.numFmt = 'dd"/"mm"/"yy';
        cell.alignment = { horizontal: 'left' };
      } else if (textColumns.includes(column)) {
        cell.value = value;
      } else {
        cell.value = Number(Number(value).toFixed(2));
        cell.numFmt = '_-"R$" * #,##0.00_-;-"R$" * #,##0.00_-;_-"R$" * "-"??_-;_-@';
      }

      cell.font = CELL_FONT;
      cell.border = CELL_BORDER;

      if (value !== 0) {
        cell.fill = PENDING_VALIDATION_COLOR;
      } else {
        const defFill = getDefaultFill(column);
        if (defFill) cell.fill = defFill;
      }
    }

    insertedRows.push({
      row: lastRow,
      patient: group.patientName,
      totalProcedures: calculateProceduresTotal(rowData),
      totalPayments: calculatePaymentsTotal(rowData),
      group,
      classificationDetails,
    });
  }

  // Force formula recalculation by clearing cached results
  worksheet.eachRow((row, rowNumber) => {
    row.eachCell((cell) => {
      if (cell.formula) {
        cell.value = { formula: cell.formula };
      }
    });
  });

  workbook.calcProperties = {
    fullCalcOnLoad: true,
  };

  await workbook.xlsx.writeFile(filePath);

  log.info(`Planilha atualizada: ${groupedTransactions.length} linhas adicionadas`);

  // Log summary of inserted rows
  if (insertedRows.length > 0) {
    logTable(insertedRows.map(r => ({ row: r.row, patient: r.patient, totalProcedures: r.totalProcedures, totalPayments: r.totalPayments })), 'Resumo das linhas inseridas');
  }

  // Save history file - group by actual transaction date
  try {
    const rowsByDate = {};
    for (const row of insertedRows) {
      const rowDate = row.group.date || date;
      if (!rowsByDate[rowDate]) rowsByDate[rowDate] = [];
      rowsByDate[rowDate].push(row);
    }
    for (const [rowDate, rows] of Object.entries(rowsByDate)) {
      await saveHistory(rows, rowDate);
    }
  } catch (err) {
    log.warn(`Erro ao salvar histórico (não crítico): ${err.message}`);
  }

  log.end('Atualização da planilha', startTime);
}

/**
 * Updates a specific sheet in an already opened workbook (without reading/saving file)
 * @param {ExcelJS.Workbook} workbook - Already loaded workbook
 * @param {string} sheetName - Sheet name to fill
 * @param {Array} transactions - List of Feegow transactions
 * @param {string} date - Reference date
 * @param {number} initialRow - Initial row for insertion
 */
export async function updateSpreadsheetCustom(workbook, sheetName, transactions, date, initialRow = 8) {
  const worksheet = workbook.getWorksheet(sheetName);
  if (!worksheet) {
    throw new Error(`Aba "${sheetName}" não encontrada`);
  }

  log.info(`Usando aba: ${worksheet.name}`);

  // Load procedure group mapping from API
  groupMapping = await fetchGroupMapping();

  let lastRow = initialRow - 1;

  const groupedTransactions = groupTransactions(transactions);
  log.info(`${groupedTransactions.length} grupos de movimentações para inserir`);

  const insertedRows = [];

  for (const group of groupedTransactions) {
    lastRow++;
    const { rowData, classificationDetails } = await buildRow(group, date);

    const textColumns = [COLUMNS.NOME, COLUMNS.VENDA, COLUMNS.OBSERVACAO];

    for (const [column, value] of Object.entries(rowData)) {
      const cell = worksheet.getCell(`${column}${lastRow}`);
      cell.value = null;
      cell.style = {};
      if (column === COLUMNS.DATA) {
        cell.value = value;
        cell.numFmt = 'dd"/"mm"/"yy';
        cell.alignment = { horizontal: 'left' };
      } else if (textColumns.includes(column)) {
        cell.value = value;
      } else {
        cell.value = Number(Number(value).toFixed(2));
        cell.numFmt = '_-"R$" * #,##0.00_-;-"R$" * #,##0.00_-;_-"R$" * "-"??_-;_-@';
      }
      cell.font = CELL_FONT;
      cell.border = CELL_BORDER;

      if (value !== 0) {
        cell.fill = PENDING_VALIDATION_COLOR;
      } else {
        const defFill = getDefaultFill(column);
        if (defFill) cell.fill = defFill;
      }
    }

    insertedRows.push({
      row: lastRow,
      patient: group.patientName,
      totalProcedures: calculateProceduresTotal(rowData),
      totalPayments: calculatePaymentsTotal(rowData),
      group,
      classificationDetails,
    });
  }

  // Force formula recalculation
  worksheet.eachRow((row) => {
    row.eachCell((cell) => {
      if (cell.formula) {
        cell.value = { formula: cell.formula };
      }
    });
  });

  workbook.calcProperties = { fullCalcOnLoad: true };

  log.info(`Aba "${sheetName}" atualizada: ${groupedTransactions.length} linhas adicionadas`);

  // Save history file - group by actual transaction date
  try {
    const rowsByDate = {};
    for (const row of insertedRows) {
      const rowDate = row.group.date || date;
      if (!rowsByDate[rowDate]) rowsByDate[rowDate] = [];
      rowsByDate[rowDate].push(row);
    }
    for (const [rowDate, rows] of Object.entries(rowsByDate)) {
      await saveHistory(rows, rowDate);
    }
  } catch (err) {
    log.warn(`Erro ao salvar histórico (não crítico): ${err.message}`);
  }

  return insertedRows;
}

/**
 * Calculates total procedures for a row
 * Note: TRATAMENTO is already the sum, so we use only it + AVALIACAO
 */
function calculateProceduresTotal(row) {
  return (row[COLUMNS.AVALIACAO] || 0) + (row[COLUMNS.TRATAMENTO] || 0);
}

/**
 * Calculates total payments for a row
 */
function calculatePaymentsTotal(row) {
  const columns = [
    COLUMNS.DINHEIRO, COLUMNS.SICOOB, COLUMNS.SAFRA, COLUMNS.PIX_SAFRA,
    COLUMNS.INFINITE, COLUMNS.PIX, COLUMNS.CHEQUE, COLUMNS.BOLETO,
  ];

  return columns.reduce((total, col) => total + (row[col] || 0), 0);
}

/**
 * Groups transactions by patient + date
 * @param {Array} transactions - List of transactions
 * @returns {Array} Grouped transactions
 */
function groupTransactions(transactions) {
  const groups = new Map();

  for (const mov of transactions) {
    const key = `${mov.PacienteID}_${mov.Data}`;

    if (!groups.has(key)) {
      groups.set(key, {
        patientId: mov.PacienteID,
        patientName: mov.NomePaciente,
        date: mov.Data,
        transactions: [],
        detailedItems: [],
      });
    }

    const group = groups.get(key);
    group.transactions.push(mov);

    // Add detailed items
    if (mov.detailedItems) {
      group.detailedItems.push(...mov.detailedItems);
    }
  }

  return Array.from(groups.values());
}

/**
 * Builds row data for the spreadsheet
 * Returns ONLY columns that have values (does not include zeros)
 * @param {Object} group - Transaction group
 * @param {string} referenceDate - Reference date
 * @returns {Promise<Object>} Object with columns and values (only significant values)
 */
async function buildRow(group, referenceDate) {
  // Temporary objects to accumulate values
  const procedureValues = {};
  const paymentValues = {};
  const classificationDetails = [];

  // Process detailed items (procedures)
  let discountTotal = 0;
  for (const item of group.detailedItems) {
    // Accumulate discounts (desconto is per unit, multiply by quantity)
    if (item.desconto && item.desconto > 0) {
      discountTotal += item.desconto * (item.quantidade || 1);
    }

    const { column, classificadoPor } = classifyProcedure(item.nome, item.procedimento_id);
    if (column === null) {
      classificationDetails.push({
        tipo: 'procedimento',
        procedimentoId: item.procedimento_id,
        nome: item.nome,
        valor: formatValue(item.valor),
        quantidade: item.quantidade || 1,
        desconto: item.desconto || 0,
        colunaDestino: null,
        colunaDestinoNome: 'DESCARTADO',
        classificadoPor,
      });
      continue;
    }

    // Validate AVALIAÇÃO: check if appointment is for processing date or future
    if (column === COLUMNS.AVALIACAO) {
      const { hasAppointmentToday, apiError } = await checkAppointmentDate(
        group.transactions[0].PacienteID,
        referenceDate,
      );

      if (apiError) {
        // API failed after retries — keep evaluation as normal (don't penalize patient)
        log.warn(`Avaliação do paciente ${group.patientName}: API de agendamentos falhou, mantendo como avaliação`);
      } else if (!hasAppointmentToday) {
        // Signal payment (sinal) — no appointment on this date, skip procedure column
        log.debug(`Avaliação do paciente ${group.patientName}: sinal antecipado (sem agendamento no dia)`);
        classificationDetails.push({
          tipo: 'procedimento',
          procedimentoId: item.procedimento_id,
          nome: item.nome,
          valor: formatValue(item.valor),
          quantidade: item.quantidade || 1,
          desconto: item.desconto || 0,
          colunaDestino: null,
          colunaDestinoNome: 'SINAL_ANTECIPADO',
          classificadoPor: 'agendamento-futuro',
        });
        // Undo discount accumulation for this item (it's not a realized procedure)
        if (item.desconto && item.desconto > 0) {
          discountTotal -= item.desconto * (item.quantidade || 1);
        }
        continue;
      }
    }

    procedureValues[column] = (procedureValues[column] || 0) + formatValue(item.valor);
    classificationDetails.push({
      tipo: 'procedimento',
      procedimentoId: item.procedimento_id,
      nome: item.nome,
      valor: formatValue(item.valor),
      quantidade: item.quantidade || 1,
      desconto: item.desconto || 0,
      colunaDestino: column,
      colunaDestinoNome: COLUMN_NAMES[column] || column,
      classificadoPor,
    });
  }

  // Process payment methods
  const notes = [];
  let hasPaymentOnly = false;

  for (const mov of group.transactions) {
    const paymentColumn = identifyPaymentColumn(mov);
    paymentValues[paymentColumn] = (paymentValues[paymentColumn] || 0) + formatValue(mov.Value);

    if (mov.paymentOnly) {
      hasPaymentOnly = true;
    }

    // Collect card information
    if (mov.Bandeira) {
      const cardInfo = `${mov.Bandeira}${mov.Parcelas ? ` ${mov.Parcelas}x` : ''}`;
      if (!notes.includes(cardInfo)) {
        notes.push(cardInfo);
      }
    }

    // Add description if available
    if (mov.Descricao && mov.Descricao.trim() && mov.Descricao !== ',') {
      notes.push(mov.Descricao.trim());
    }
  }

  // Calculate treatment total (sum of all procedure columns, except AVALIACAO and ESTORNO)
  const columnsForTotal = [
    COLUMNS.DR_RIGATTI, COLUMNS.IMPLANTE, COLUMNS.SORO_DR, COLUMNS.TIRZEPATIDA,
    COLUMNS.APLICACOES, COLUMNS.ONLINE, COLUMNS.COMISSAO, COLUMNS.NUTRIS, COLUMNS.EXTRA
  ];

  let treatmentTotal = 0;
  for (const col of columnsForTotal) {
    treatmentTotal += procedureValues[col] || 0;
  }

  // Build final row object
  const row = {};

  // Parse date string (DD/MM/YYYY or YYYY-MM-DD) into Date object
  const dateStr = group.date || referenceDate;
  let dateObj;
  if (dateStr.includes('/')) {
    const [d, m, y] = dateStr.split('/').map(Number);
    dateObj = new Date(y, m - 1, d);
  } else {
    const [y, m, d] = dateStr.split('-').map(Number);
    dateObj = new Date(y, m - 1, d);
  }
  row[COLUMNS.DATA] = dateObj;
  row[COLUMNS.NOME] = (group.patientName || '').toUpperCase();

  // Notes/observations in VENDA column
  if (notes.length > 0) {
    row[COLUMNS.VENDA] = notes.join(' | ');
  }

  // paymentOnly observation goes to dedicated column AG
  if (hasPaymentOnly) {
    row[COLUMNS.OBSERVACAO] = 'Tratamento em andamento';
  }

  // Add procedures that have value > 0
  for (const [column, value] of Object.entries(procedureValues)) {
    if (value > 0) {
      row[column] = value;
    }
  }

  // Add treatment total (subtract discount) if > 0
  if (treatmentTotal > 0) {
    row[COLUMNS.TRATAMENTO] = treatmentTotal - discountTotal;
  }

  // Add discount if > 0
  if (discountTotal > 0) {
    row[COLUMNS.DESCONTO] = discountTotal;
  }

  // Add payments that have value > 0
  for (const [column, value] of Object.entries(paymentValues)) {
    if (value > 0) {
      row[column] = value;
    }
  }

  // Fill explicit zeros in all numeric columns not yet set
  // Excludes: DATA (A), NOME (B), VENDA (C), Y, OBSERVACAO (AG), AH — manually filled
  const manualColumns = new Set([COLUMNS.DATA, COLUMNS.NOME, COLUMNS.VENDA, 'Y', COLUMNS.OBSERVACAO, 'AH']);
  for (const col of Object.values(COLUMNS)) {
    if (!manualColumns.has(col) && !(col in row)) {
      row[col] = 0;
    }
  }

  return { rowData: row, classificationDetails };
}

// Exports for history service
export { COLUMNS, COLUMN_NAMES, MONTH_NAMES_PT, groupTransactions, buildRow, loadGroupMapping };

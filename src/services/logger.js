import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const LOGS_DIR = path.resolve(__dirname, '../../logs');

// Log levels
const LOG_LEVELS = {
  DEBUG: 0,
  INFO: 1,
  WARN: 2,
  ERROR: 3,
};

// Console colors (ANSI)
const COLORS = {
  DEBUG: '\x1b[36m',   // Cyan
  INFO: '\x1b[32m',    // Green
  WARN: '\x1b[33m',    // Yellow
  ERROR: '\x1b[31m',   // Red
  RESET: '\x1b[0m',
  DIM: '\x1b[2m',
  BRIGHT: '\x1b[1m',
};

// Logger configuration
const config = {
  level: LOG_LEVELS[process.env.LOG_LEVEL?.toUpperCase()] ?? LOG_LEVELS.DEBUG,
  logToFile: process.env.LOG_TO_FILE !== 'false',
  logToConsole: true,
  maxFileSize: 10 * 1024 * 1024, // 10MB
};

// Ensures log directory exists
function ensureLogDir() {
  if (!fs.existsSync(LOGS_DIR)) {
    fs.mkdirSync(LOGS_DIR, { recursive: true });
  }
}

function getTimestamp() {
  return new Date().toISOString();
}

function getDateString() {
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
}

function getLogFilePath() {
  ensureLogDir();
  return path.join(LOGS_DIR, `agente-financeiro-${getDateString()}.log`);
}

function formatMessage(level, context, message, data) {
  const timestamp = getTimestamp();
  const contextStr = context ? `[${context}]` : '';
  const dataStr = data ? `\n${JSON.stringify(data, null, 2)}` : '';

  return {
    formatted: `${timestamp} [${level}] ${contextStr} ${message}${dataStr}`,
    json: {
      timestamp,
      level,
      context,
      message,
      data,
    },
  };
}

function writeToFile(formattedMessage) {
  if (!config.logToFile) return;

  try {
    const logPath = getLogFilePath();
    fs.appendFileSync(logPath, formattedMessage + '\n');
  } catch (error) {
    console.error('Error writing log to file:', error.message);
  }
}

function writeToConsole(level, context, message, data) {
  if (!config.logToConsole) return;

  const color = COLORS[level] || COLORS.RESET;
  const timestamp = getTimestamp();
  const contextStr = context ? `${COLORS.DIM}[${context}]${COLORS.RESET}` : '';

  const prefix = `${COLORS.DIM}${timestamp}${COLORS.RESET} ${color}[${level}]${COLORS.RESET}`;

  console.log(`${prefix} ${contextStr} ${message}`);

  if (data) {
    console.log(COLORS.DIM + JSON.stringify(data, null, 2) + COLORS.RESET);
  }
}

function log(level, context, message, data = null) {
  const levelNum = LOG_LEVELS[level];

  if (levelNum < config.level) return;

  const { formatted } = formatMessage(level, context, message, data);

  writeToConsole(level, context, message, data);
  writeToFile(formatted);
}

// Logger with context
class Logger {
  constructor(context = '') {
    this.context = context;
  }

  debug(message, data) {
    log('DEBUG', this.context, message, data);
  }

  info(message, data) {
    log('INFO', this.context, message, data);
  }

  warn(message, data) {
    log('WARN', this.context, message, data);
  }

  error(message, data) {
    log('ERROR', this.context, message, data);
  }

  // Log operation start
  start(operation) {
    this.info(`Iniciando: ${operation}`);
    return Date.now();
  }

  // Log operation end with duration
  end(operation, startTime) {
    const duration = Date.now() - startTime;
    this.info(`Concluído: ${operation}`, { duration: `${duration}ms` });
  }

  // Log progress
  progress(current, total, description = '') {
    const percent = Math.round((current / total) * 100);
    this.debug(`Progresso: ${current}/${total} (${percent}%) ${description}`);
  }

  // Creates a child logger with additional context
  child(subContext) {
    const newContext = this.context ? `${this.context}:${subContext}` : subContext;
    return new Logger(newContext);
  }
}

// Exports default logger and logger creator
export const logger = new Logger();

export function createLogger(context) {
  return new Logger(context);
}

// Export utility functions
export function setLogLevel(level) {
  config.level = LOG_LEVELS[level.toUpperCase()] ?? LOG_LEVELS.DEBUG;
}

export function enableFileLogging(enabled = true) {
  config.logToFile = enabled;
}

export function enableConsoleLogging(enabled = true) {
  config.logToConsole = enabled;
}

// Log separator for better visualization
export function logSeparator(title = '') {
  const line = '='.repeat(60);
  if (title) {
    const padding = Math.max(0, (60 - title.length - 2) / 2);
    const paddedTitle = ' '.repeat(Math.floor(padding)) + title + ' '.repeat(Math.ceil(padding));
    console.log(`\n${COLORS.BRIGHT}${line}\n${paddedTitle}\n${line}${COLORS.RESET}\n`);
    writeToFile(`\n${line}\n${title}\n${line}\n`);
  } else {
    console.log(`\n${COLORS.DIM}${line}${COLORS.RESET}\n`);
    writeToFile(`\n${line}\n`);
  }
}

// Log structured data as table
export function logTable(data, title = '') {
  if (title) {
    console.log(`\n${COLORS.BRIGHT}${title}${COLORS.RESET}`);
  }
  console.table(data);

  // Simplified version for file
  writeToFile(`\n${title}\n${JSON.stringify(data, null, 2)}\n`);
}

export default logger;

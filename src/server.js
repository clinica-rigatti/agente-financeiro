import 'dotenv/config';
import { createServer } from 'http';
import dayjs from 'dayjs';
import { fetchDetailedTransactions, isoToFeegow } from './services/feegow.js';
import { updateSpreadsheet } from './services/excel.js';
import { createLogger, logSeparator } from './services/logger.js';

const log = createLogger('Server');
const PORT = process.env.API_PORT || 3100;
const TOKEN = process.env.API_TOKEN;

if (!TOKEN) {
  log.error('API_TOKEN não configurado no .env — servidor não iniciará');
  process.exit(1);
}

let running = false;

async function executarBoletos() {
  if (running) {
    return { status: 409, body: { success: false, message: 'Já existe uma execução em andamento. Aguarde.' } };
  }

  running = true;
  const startTime = Date.now();

  try {
    logSeparator('EXECUÇÃO BOLETOS (MANUAL)');

    const referenceDate = dayjs().subtract(1, 'day');
    const formattedDate = referenceDate.format('YYYY-MM-DD');
    const feegowDate = isoToFeegow(formattedDate);

    log.info(`Buscando boletos do dia ${formattedDate}...`);

    // Fetch all transactions for the day
    const allTransactions = await fetchDetailedTransactions(feegowDate, feegowDate);

    // Filter only boletos (FormaPagamentoID === 4)
    const boletos = allTransactions.filter(t => t.FormaPagamentoID === 4);

    if (boletos.length === 0) {
      log.info('Nenhum boleto encontrado para esta data');
      return { status: 200, body: { success: true, message: 'Nenhum boleto encontrado', date: formattedDate, count: 0 } };
    }

    log.info(`${boletos.length} boleto(s) encontrado(s) — inserindo na planilha`);

    await updateSpreadsheet(boletos, formattedDate);

    const duration = Date.now() - startTime;
    log.info(`Concluído em ${duration}ms: ${boletos.length} boleto(s) processado(s)`);

    return {
      status: 200,
      body: {
        success: true,
        message: `${boletos.length} boleto(s) inserido(s) na planilha`,
        date: formattedDate,
        count: boletos.length,
        duration: `${duration}ms`,
      },
    };
  } catch (error) {
    log.error('Erro ao processar boletos', { message: error.message, stack: error.stack });
    return { status: 500, body: { success: false, message: error.message } };
  } finally {
    running = false;
  }
}

const ALLOWED_ORIGINS = [
  'https://maria.clinicarigatti.com.br',
  'https://rigatti.tech',
];

const server = createServer(async (req, res) => {
  // CORS: only allow specific origins
  const origin = req.headers.origin;
  if (origin && ALLOWED_ORIGINS.some(o => origin === o || origin.endsWith('.' + o.replace('https://', '')))) {
    res.setHeader('Access-Control-Allow-Origin', origin);
  }
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Authorization, Content-Type');
  res.setHeader('Vary', 'Origin');

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  // Health check (no auth needed)
  if (req.method === 'GET' && req.url === '/health') {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ status: 'ok', running }));
    return;
  }

  // Auth check
  const auth = req.headers.authorization;
  if (!auth || auth !== `Bearer ${TOKEN}`) {
    res.writeHead(401, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ success: false, message: 'Token inválido' }));
    return;
  }

  if (req.method === 'POST' && req.url === '/executar-boletos') {
    const result = await executarBoletos();
    res.writeHead(result.status, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify(result.body));
    return;
  }

  res.writeHead(404, { 'Content-Type': 'application/json' });
  res.end(JSON.stringify({ success: false, message: 'Rota não encontrada' }));
});

server.listen(PORT, () => {
  log.info(`Servidor rodando na porta ${PORT}`);
  log.info('Endpoints:');
  log.info(`  POST /executar-boletos  — Busca e insere boletos do dia anterior`);
  log.info(`  GET  /health            — Health check`);
});

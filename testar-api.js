import 'dotenv/config';
import axios from 'axios';

const api = axios.create({
  baseURL: process.env.FEEGOW_API_URL,
  headers: {
    'Content-Type': 'application/json',
    'x-access-token': process.env.FEEGOW_API_TOKEN,
  },
});

async function testarEndpoint(method, path, body = null) {
  console.log(`\n=== ${method} ${path} ===`);
  if (body) console.log('Body:', JSON.stringify(body, null, 2));

  try {
    const config = { method, url: path };
    if (body) config.data = body;

    const response = await api(config);
    const result = response.data;

    if (Array.isArray(result)) {
      console.log(`Retornou array com ${result.length} itens`);
      if (result.length > 0) {
        console.log('Campos:', Object.keys(result[0]).join(', '));
        console.log('Primeiro item:', JSON.stringify(result[0], null, 2));
      }
    } else if (result.data && Array.isArray(result.data)) {
      console.log(`Retornou ${result.data.length} registros`);
      if (result.data.length > 0) {
        console.log('Campos:', Object.keys(result.data[0]).join(', '));
        console.log('Primeiro:', JSON.stringify(result.data[0], null, 2));
      }
    } else if (result.content && Array.isArray(result.content)) {
      console.log(`Retornou ${result.content.length} registros em content`);
      if (result.content.length > 0) {
        console.log('Campos:', Object.keys(result.content[0]).join(', '));
        console.log('Primeiro:', JSON.stringify(result.content[0], null, 2));
      }
    } else {
      console.log('Resposta:', JSON.stringify(result, null, 2).slice(0, 1000));
    }

    return result;
  } catch (error) {
    console.error('Erro:', error.response?.status, error.response?.data?.message || error.message);
    return null;
  }
}

async function main() {
  // Tenta endpoints comuns de APIs financeiras
  const endpoints = [
    { method: 'GET', path: '/api/financial' },
    { method: 'GET', path: '/api/financial/list' },
    { method: 'GET', path: '/api/financial/movements' },
    { method: 'GET', path: '/api/financial/transactions' },
    { method: 'GET', path: '/api/payment' },
    { method: 'GET', path: '/api/payment/list' },
    { method: 'GET', path: '/api/payment/methods' },
    { method: 'GET', path: '/api/cashflow' },
    { method: 'GET', path: '/api/cashflow/list' },
    { method: 'GET', path: '/api/appoint/search' },
  ];

  for (const ep of endpoints) {
    await testarEndpoint(ep.method, ep.path, ep.body);
  }

  // Tenta buscar agendamentos com pagamento
  console.log('\n' + '='.repeat(60));
  console.log('Testando busca de agendamentos com filtros...');

  await testarEndpoint('POST', '/api/appoint/search', {
    data_start: '2025-01-01',
    data_end: '2025-01-31'
  });
}

main();

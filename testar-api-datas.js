import axios from 'axios';
import { writeFileSync } from 'fs';

const api = axios.create({
  baseURL: 'https://api.feegow.com/v1',
  headers: {
    'Content-Type': 'application/json',
    'x-access-token': 'eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiJmZWVnb3ciLCJhdWQiOiJwdWJsaWNhcGkiLCJpYXQiOjE3MzI4MjQ1MzUsImxpY2Vuc2VJRCI6NDE1NTl9.iNBVZt6Vztq-sT4hf2HrqIte_oTmZCpORL3BQqdCMyI',
  },
});

async function main() {
  console.log('Buscando movimentações via API pública...\n');

  const body = {
    report: 'financial-movement',
    DATA_INICIO: '01/02/2026',
    DATA_FIM: '05/02/2026',
    TIPO_CONTA_IDS: [3]  // 3 = Paciente
  };

  try {
    const response = await api.post('/api/reports/generate', body);
    const data = response.data;

    console.log(`Total: ${data.data.length} registros\n`);
    console.log('Campos disponíveis:', Object.keys(data.data[0]).join(', '));

    // Salva mock completo
    const mockData = {
      periodo: { dataInicio: '01/02/2026', dataFim: '05/02/2026' },
      buscadoEm: new Date().toISOString(),
      quantidade: data.data.length,
      movimentacoes: data.data
    };

    writeFileSync(
      './src/mocks/movimentacoes-api-mock.json',
      JSON.stringify(mockData, null, 2),
      'utf-8'
    );

    console.log('\nMock salvo em: src/mocks/movimentacoes-api-mock.json');

    // Mostra primeiros 3 registros
    console.log('\n--- Primeiros 3 registros ---\n');
    data.data.slice(0, 3).forEach((mov, i) => {
      console.log(`[${i + 1}]`, JSON.stringify(mov, null, 2));
      console.log('');
    });

  } catch (error) {
    console.error('Erro:', error.response?.data || error.message);
  }
}

main();

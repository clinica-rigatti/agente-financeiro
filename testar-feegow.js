import { fetchMovimentacoes } from './src/services/feegow.js';
import { readFileSync, writeFileSync } from 'fs';

async function main() {
  console.log('=== Buscando movimentações e salvando mock ===\n');

  const config = JSON.parse(readFileSync('./config/feegow-session.json', 'utf-8'));

  const dataInicio = '01/02/2026';
  const dataFim = '04/02/2026';

  try {
    const resultado = await fetchMovimentacoes(dataInicio, dataFim, config);

    // Salva o resultado completo em arquivo mock
    const mockData = {
      periodo: { dataInicio, dataFim },
      buscadoEm: new Date().toISOString(),
      total: resultado.total,
      quantidade: resultado.quantidade,
      movimentacoes: resultado.movimentacoes
    };

    writeFileSync(
      './src/mocks/movimentacoes-mock.json',
      JSON.stringify(mockData, null, 2),
      'utf-8'
    );

    console.log('Arquivo salvo em: src/mocks/movimentacoes-mock.json\n');

    // Exibe JSON formatado
    console.log(JSON.stringify(mockData, null, 2));

  } catch (error) {
    console.error('ERRO:', error.message);
  }
}

main();

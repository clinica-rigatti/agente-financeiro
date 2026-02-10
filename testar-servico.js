import 'dotenv/config';
import { fetchMovimentacoes, formatDateFeegow } from './src/services/feegow.js';

async function main() {
  console.log('=== Testando serviço atualizado ===\n');

  const dataInicio = '01/02/2026';
  const dataFim = '05/02/2026';

  console.log(`Período: ${dataInicio} a ${dataFim}\n`);

  try {
    const movimentacoes = await fetchMovimentacoes(dataInicio, dataFim);

    console.log(`\nTotal: ${movimentacoes.length} movimentações\n`);

    // Mostra exemplo da Thaís (R$ 7.500,00) para entender a composição
    const thais = movimentacoes.find(m => m.NomePaciente === 'Thaís Moraes Lange' && m.Value === 7500);

    if (thais) {
      console.log('=== Exemplo: Thaís Moraes Lange (R$ 7.500,00) ===\n');
      console.log('Dados disponíveis:');
      console.log(`  - PacienteID: ${thais.PacienteID}`);
      console.log(`  - NumeroInvoice: ${thais.NumeroInvoice}`);
      console.log(`  - IDTransacao: ${thais.IDTransacao}`);
      console.log(`  - NomeProcedimento: ${thais.NomeProcedimento}`);
      console.log(`  - Categoria: ${thais.Categoria}`);
      console.log(`  - ProcedimentoID: ${thais.ProcedimentoID}`);
      console.log(`  - FormaPagamento: ${thais.FormaPagamento}`);
      console.log(`  - AccountName: ${thais.AccountName}`);
      console.log(`  - Parcelas: ${thais.Parcelas}`);
      console.log(`  - Bandeira: ${thais.Bandeira}`);
    }

    // Lista todas as movimentações agrupadas por paciente
    console.log('\n=== Movimentações por Paciente ===\n');

    const porPaciente = {};
    movimentacoes.forEach(m => {
      const key = `${m.PacienteID} - ${m.NomePaciente}`;
      if (!porPaciente[key]) {
        porPaciente[key] = [];
      }
      porPaciente[key].push({
        valor: m.Value,
        procedimento: m.NomeProcedimento,
        categoria: m.Categoria,
        forma: m.FormaPagamento,
        invoice: m.NumeroInvoice,
      });
    });

    Object.entries(porPaciente).forEach(([paciente, movs]) => {
      const total = movs.reduce((sum, m) => sum + m.valor, 0);
      console.log(`${paciente}`);
      console.log(`  Total: R$ ${total.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}`);
      movs.forEach(m => {
        console.log(`    - R$ ${m.valor.toLocaleString('pt-BR', { minimumFractionDigits: 2 })} | ${m.procedimento || 'N/A'} | ${m.forma}`);
      });
      console.log('');
    });

  } catch (error) {
    console.error('ERRO:', error.message);
  }
}

main();

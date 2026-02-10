import ExcelJS from 'exceljs';

async function analisarPlanilha() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('./onedrive/planilhas/CÓPIA - ENTRADAS  ANO 2025.xlsx');

  console.log('=== ABAS DA PLANILHA ===\n');
  workbook.worksheets.forEach((ws, i) => {
    console.log(`${i + 1}. "${ws.name}" (${ws.rowCount} linhas)`);
  });

  // Analisa cada aba
  for (const worksheet of workbook.worksheets) {
    console.log(`\n${'='.repeat(60)}`);
    console.log(`ABA: "${worksheet.name}"`);
    console.log('='.repeat(60));

    // Pega o cabeçalho (primeira linha com dados)
    const headerRow = worksheet.getRow(1);
    console.log('\n--- COLUNAS (Cabeçalho) ---');

    headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      const letra = getColumnLetter(colNumber);
      console.log(`${letra} = ${cell.value}`);
    });

    // Mostra algumas linhas de exemplo
    console.log('\n--- PRIMEIRAS 5 LINHAS DE DADOS ---');
    for (let i = 2; i <= Math.min(6, worksheet.rowCount); i++) {
      const row = worksheet.getRow(i);
      let rowData = [];
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        let valor = cell.value;
        if (valor && typeof valor === 'object') {
          if (valor.result !== undefined) valor = valor.result;
          else if (valor.text) valor = valor.text;
          else valor = JSON.stringify(valor);
        }
        rowData.push(`${getColumnLetter(colNumber)}:${valor}`);
      });
      if (rowData.length > 0) {
        console.log(`Linha ${i}: ${rowData.join(' | ')}`);
      }
    }
  }
}

function getColumnLetter(colNumber) {
  let letter = '';
  while (colNumber > 0) {
    const mod = (colNumber - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    colNumber = Math.floor((colNumber - 1) / 26);
  }
  return letter;
}

analisarPlanilha().catch(console.error);

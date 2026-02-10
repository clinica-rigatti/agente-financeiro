import 'dotenv/config';
import path from 'path';
import { fileURLToPath } from 'url';
import ExcelJS from 'exceljs';
import { fetchDetailedTransactions } from './services/feegow.js';
import { updateSpreadsheet } from './services/excel.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

async function main() {
  const filePath = path.resolve(__dirname, '../onedrive', process.env.EXCEL_FILE_PATH);
  const sheetName = process.env.EXCEL_ABA;
  const initialRow = parseInt(process.env.EXCEL_LINHA_INICIAL) || 8;

  // 1. Reset
  console.log('Resetting sheet...');
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const ws = workbook.getWorksheet(sheetName);
  ws.eachRow((row, rowNumber) => {
    if (rowNumber >= initialRow) {
      row.eachCell({ includeEmpty: false }, (cell) => {
        cell.value = null; cell.fill = undefined; cell.border = undefined; cell.style = {};
      });
    }
  });
  await workbook.xlsx.writeFile(filePath);

  // 2. Fetch and insert
  console.log('Fetching and inserting...');
  const transactions = await fetchDetailedTransactions('05/02/2026', '05/02/2026');
  await updateSpreadsheet(transactions, '05/02/2026');

  // 3. Read spreadsheet
  const wb2 = new ExcelJS.Workbook();
  await wb2.xlsx.readFile(filePath);
  const ws2 = wb2.getWorksheet(sheetName);

  const cols = 'A B C D E F G H I J K L M N O P Q R S T U V W X Y Z'.split(' ');
  const extraCols = ['AA', 'AB', 'AC', 'AD', 'AE', 'AF'];
  const allCols = [...cols, ...extraCols];

  for (const search of ['ana l', 'lara', 'melina', 'carlos roberto', 'marciana']) {
    console.log(`\n=== ${search.toUpperCase()} ===\n`);
    for (let row = initialRow; row <= ws2.rowCount; row++) {
      const name = ws2.getCell(`B${row}`).value;
      if (!name || !String(name).toLowerCase().includes(search)) continue;
      console.log(`Row ${row}: ${name}`);
      for (const col of allCols) {
        const cell = ws2.getCell(`${col}${row}`);
        let val = cell.value;
        if (val === null || val === undefined || val === '') continue;
        const header = ws2.getCell(`${col}7`).value || ws2.getCell(`${col}4`).value || '(vazio)';
        console.log(`  ${col.padEnd(3)} | ${String(header).padEnd(16)} | ${val}`);
      }
    }
  }
}

main().catch(console.error);

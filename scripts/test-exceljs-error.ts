import ExcelJS from 'exceljs';

async function test() {
  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.load(Buffer.from('not a zip'));
  } catch (err) {
    console.log('Error caught:', err.message);
  }
}

test();

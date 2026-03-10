import puppeteer from 'puppeteer';
import fs from 'fs';
import JSZip from 'jszip';

const tokenUrl = process.argv[2] || 'https://catalogo-vpfe.dian.gov.co/User/AuthToken?pk=10910094%7C1026592934&rk=901965856&token=5490cd9c-815c-4748-927c-0d4680d327a7';

async function test() {
  console.log('Iniciando Puppeteer...');
  const browser = await puppeteer.launch({ 
    headless: true, 
    args: ['--no-sandbox', '--disable-setuid-sandbox'] 
  });
  const page = await browser.newPage();
  
  console.log('Autenticando en DIAN...');
  await page.goto(tokenUrl, { waitUntil: 'networkidle0', timeout: 60000 });
  
  // Esperar a que cargue la pagina de documentos
  try {
    await page.waitForSelector('#ContentPlaceHolder1_GridReceivedDocuments', { timeout: 15000 });
  } catch {
    console.log('Tabla no encontrada, intentando continuar...');
  }
  
  // Poner fechas
  await page.evaluate(() => {
    const startInput = document.querySelector('#StartDate');
    const endInput = document.querySelector('#EndDate');
    if (startInput) startInput.value = '2025-02-01';
    if (endInput) endInput.value = '2025-02-25';
  });
  
  // Buscar
  const searchBtn = await page.$('#btnSearch');
  if (searchBtn) {
    await searchBtn.click();
    await new Promise(r => setTimeout(r, 5000));
  }
  
  // Extraer info de la tabla
  const docs = await page.evaluate(() => {
    const rows = document.querySelectorAll('#ContentPlaceHolder1_GridReceivedDocuments tbody tr');
    return Array.from(rows).slice(0, 3).map(row => {
      const cells = row.querySelectorAll('td');
      const downloadBtn = row.querySelector('a[onclick*="DownloadZipFiles"]');
      const trackId = downloadBtn ? downloadBtn.getAttribute('onclick').match(/trackId=([^'&]+)/)?.[1] : null;
      return {
        docnum: cells[2]?.textContent?.trim(),
        nit: cells[3]?.textContent?.trim(),
        fecha: cells[1]?.textContent?.trim(),
        trackId
      };
    }).filter(d => d.trackId);
  });
  
  console.log('Documentos encontrados:', docs.length);
  if (docs.length > 0) {
    console.log('Primer documento:', JSON.stringify(docs[0], null, 2));
  }
  
  if (docs[0]?.trackId) {
    // Descargar el ZIP
    const cookies = await page.cookies();
    const cookieStr = cookies.map(c => `${c.name}=${c.value}`).join('; ');
    
    const zipUrl = `https://catalogo-vpfe.dian.gov.co/Document/DownloadZipFiles?trackId=${docs[0].trackId}`;
    console.log('Descargando ZIP:', zipUrl.substring(0, 80) + '...');
    
    const response = await fetch(zipUrl, { headers: { Cookie: cookieStr } });
    const buffer = Buffer.from(await response.arrayBuffer());
    
    console.log('ZIP descargado, tamano:', buffer.length, 'bytes');
    
    // Extraer XML del ZIP
    const zip = await JSZip.loadAsync(buffer);
    const files = Object.keys(zip.files);
    console.log('Archivos en ZIP:', files);
    
    for (const [filename, file] of Object.entries(zip.files)) {
      if (filename.toLowerCase().endsWith('.xml') && !file.dir) {
        const content = await file.async('string');
        fs.writeFileSync('/tmp/factura_sample.xml', content);
        console.log('\n========== XML COMPLETO (/tmp/factura_sample.xml) ==========\n');
        console.log(content);
        break;
      }
    }
  } else {
    console.log('No se encontraron documentos con trackId');
  }
  
  await browser.close();
}

test().catch(e => {
  console.error('Error:', e);
  process.exit(1);
});

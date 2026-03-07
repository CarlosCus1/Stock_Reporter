/**
 * @file download-stock.js
 * @author Carlos Cusi
 * @description Descarga el reporte de stock y lo convierte en data_stock.json.
 * Soporta protocolos HTTP y HTTPS dinámicamente.
 */

import https from 'https';
import http from 'http';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import xlsx from 'xlsx';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const STOCK_URL = process.env.STOCK_URL;

async function downloadAndConvert() {
  if (!STOCK_URL) {
    console.error('❌ ERROR: STOCK_URL no configurada en los secrets de GitHub.');
    process.exit(1);
  }

  console.log('📥 Iniciando descarga de stock...');
  
  const protocol = STOCK_URL.startsWith('https') ? https : http;

  protocol.get(STOCK_URL, (res) => {
    if (res.statusCode !== 200) {
      console.error(`❌ Error de servidor: ${res.statusCode}`);
      process.exit(1);
    }

    const chunks = [];
    res.on('data', (chunk) => chunks.push(chunk));
    res.on('end', () => {
      const buffer = Buffer.concat(chunks);
      const workbook = xlsx.read(buffer, { type: 'buffer' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      
      // Leer como texto formateado para preservar ceros (02210)
      const rawData = xlsx.utils.sheet_to_json(sheet, { header: 1, raw: false });
      const stockMap = {};

      rawData.forEach((row, index) => {
        if (index === 0) return; // Saltar cabecera
        
        // Mapeo basado en la estructura estándar de CIPSA:
        // Columna 1: SKU | Columna 9: Almacen | Columna 18: Disponible
        const sku = String(row[1] || '').trim(); // Sin eliminar ceros - preservar como viene
        const almacen = String(row[9] || '').trim().toUpperCase();
        const disponible = parseInt(row[18], 10) || 0;

        if (sku && (almacen === 'VES' || almacen === '')) {
          stockMap[sku] = (stockMap[sku] || 0) + disponible;
        }
      });

      const outputPath = path.join(__dirname, '..', 'Data', 'data_stock.json');
      fs.writeFileSync(outputPath, JSON.stringify({
        lastUpdate: new Date().toISOString(),
        stock: stockMap
      }, null, 2));

      console.log(`✅ Data de stock guardada: ${Object.keys(stockMap).length} productos procesados.`);
    });
  }).on('error', (err) => {
    console.error('❌ Fallo de conexión:', err.message);
    process.exit(1);
  });
}

downloadAndConvert();

/**
 * @file convert-codigos.js
 * @author Carlos Cusi
 * @description Transforma el maestro codigos_generales.xlsx en un JSON optimizado.
 * Este script es el corazón de la taxonomía del sistema.
 */

import xlsx from 'xlsx';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

function convertirCodigosAJson() {
  try {
    const inputPath = path.join(__dirname, '..', 'Data', 'codigos_generales.xlsx');
    console.log(`📂 Procesando Maestro de Productos: ${inputPath}`);
    
    if (!fs.existsSync(inputPath)) {
      throw new Error('CRÍTICO: No se encontró Data/codigos_generales.xlsx');
    }

    const workbook = xlsx.readFile(inputPath);
    const data = xlsx.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
    
    if (data.length === 0) {
      throw new Error('ERROR: El archivo Excel está vacío.');
    }

    // Identificar líneas únicas para la generación automática de pestañas
    const lineasDetectadas = new Set();

    const productos = data.map((row, index) => {
      // Normalización de campos clave
      const sku = String(row.SKU || row.sku || '').trim().replace(/^0+/, '');
      const linea = (row.LINEA || row.linea || 'SIN LINEA').trim().toUpperCase();
      const categoria = (row.CATEGORIA || row.categoria || 'SIN CATEGORIA').trim().toUpperCase();
      const nombre = (row.NOMBRE || row.nombre || '').trim();

      if (linea !== 'SIN LINEA') lineasDetectadas.add(linea);

      return {
        orden: row.ORDEN || row.orden || index + 1,
        sku,
        nombre,
        ean: String(row.EAN || row.ean || row.EAN_13 || '').trim(),
        linea,
        categoria,
        unBx: parseInt(row['UN/BX'] || row.unBx, 10) || 0
      };
    }).filter(p => p.sku !== ''); // Eliminar filas sin SKU

    // Estructura final con metadatos para el generador de reportes
    const output = {
      metadata: {
        totalMaster: productos.length,
        actualizado: new Date().toISOString(),
        lineas: Array.from(lineasDetectadas).sort()
      },
      productos
    };
    
    const outputPath = path.join(__dirname, '..', 'Data', 'productos.json');
    fs.writeFileSync(outputPath, JSON.stringify(output, null, 2));
    
    console.log(`✅ EXITO: ${productos.length} productos procesados.`);
    console.log(`📋 Líneas identificadas: ${output.metadata.lineas.join(', ')}`);
    
  } catch (error) {
    console.error('❌ ERROR CRÍTICO EN CONVERSIÓN:', error.message);
    process.exit(1);
  }
}

convertirCodigosAJson();

import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import xlsx from 'xlsx';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PRIMARY_COLOR = 'FF13DAEC';
const RED_BG = 'FFFEE2E2';
const YELLOW_BG = 'FFFEF3C7';
const GREEN_BG = 'FFD1FAE5';
const RED_ARROW = '↑';
const YELLOW_ARROW = '→';
const GREEN_ARROW = '↓';

const normalizeSKU = (sku) => String(sku || '').trim().replace(/^0+/, '');

const applyProfessionalStyles = (worksheet) => {
  worksheet.columns = [
    { header: '#', key: 'item', width: 6 },
    { header: 'Código', key: 'sku', width: 12 },
    { header: 'EAN', key: 'ean', width: 18 },
    { header: 'Nombre del Producto', key: 'nombre', width: 45 },
    { header: 'U. x Caja', key: 'unBx', width: 10 },
    { header: 'Stock', key: 'stock', width: 10 },
    { header: 'Estado', key: 'flecha', width: 8 },
    { header: 'Color', key: 'color', width: 8 }
  ];
  const headerRow = worksheet.getRow(1);
  headerRow.height = 30;
  headerRow.eachCell((cell) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: PRIMARY_COLOR } };
    cell.font = { bold: true, size: 11 };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  });
  worksheet.autoFilter = { from: 'A1', to: 'H1' };
  worksheet.views = [{ state: 'frozen', ySplit: 1 }];
};

const addDataToSheet = (worksheet, data) => {
  data.forEach((p, index) => {
    const row = worksheet.addRow([index + 1, p.sku, p.ean, p.nombre, p.unBx, p.stock, p.flecha, p.color]);
    // Columna Estado (flecha)
    row.getCell(7).value = p.flecha;
    row.getCell(7).font = { bold: true, size: 14 };
    row.getCell(7).alignment = { horizontal: 'center' };
    // Columna Color (fondo)
    row.getCell(8).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: p.bgColor } };
    row.getCell(8).alignment = { horizontal: 'center' };
  });
};

async function runSnapshotUpdate() {
  try {
    console.log('🚀 Actualizando Snapshots de StockPulse (4 categorías)...');

    const productosPath = path.join(__dirname, '..', 'Data', 'productos.json');
    const { productos, metadata: masterMeta } = JSON.parse(fs.readFileSync(productosPath, 'utf8'));

    // 2. Cargar Stock desde JSON (generado por download-stock.js)
    const stockJsonPath = path.join(__dirname, '..', 'Data', 'data_stock.json');
    if (!fs.existsSync(stockJsonPath)) {
      throw new Error('No se encontró Data/data_stock.json. Ejecute download-stock.js primero.');
    }
    const { stock: stockMap } = JSON.parse(fs.readFileSync(stockJsonPath, 'utf8'));
    console.log('✅ Stock cargado desde JSON intermedio.');

    let countSinStock = 0;
    let countBajoStock = 0;

    const fullData = productos.map(p => {
      const stock = stockMap[p.sku] || 0;
      let bgColor = GREEN_BG, flecha = GREEN_ARROW, color = 'Verde';
      if (stock === 0) { bgColor = RED_BG; flecha = RED_ARROW; color = 'Rojo'; countSinStock++; }
      else if (stock < 10) { bgColor = YELLOW_BG; flecha = YELLOW_ARROW; color = 'Amarillo'; countBajoStock++; }
      return { ...p, stock, flecha, color, bgColor };
    });

    const outputDirs = [
      path.join(__dirname, '..', 'reports'),
      path.join(__dirname, '..', 'frontend', 'public', 'reports')
    ];
    outputDirs.forEach(dir => { if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true }); });

    // --- GENERAR SNAPSHOTS POR CATEGORÍA ---
    const categoriasADescargar = ['PELOTAS', 'ESCOLAR', 'REPRESENTADAS'];

    for (const cat of categoriasADescargar) {
      const workbook = new ExcelJS.Workbook();
      const dataCat = fullData.filter(p => p.categoria === cat);

      if (dataCat.length > 0) {
        if (cat === 'ESCOLAR') {
          const lineas = [...new Set(dataCat.map(p => p.linea))];
          lineas.forEach(lin => {
            const sheet = workbook.addWorksheet(lin.substring(0, 31));
            applyProfessionalStyles(sheet);
            addDataToSheet(sheet, dataCat.filter(p => p.linea === lin));
          });
        } else {
          const sheet = workbook.addWorksheet(cat);
          applyProfessionalStyles(sheet);
          addDataToSheet(sheet, dataCat);
        }
        const fileName = `StockPulse_${cat}.xlsx`;
        for (const dir of outputDirs) await workbook.xlsx.writeFile(path.join(dir, fileName));
      }
    }

    // --- GENERAR SNAPSHOT MAESTRO (TODOS) ---
    const wbAll = new ExcelJS.Workbook();
    const wsResumen = wbAll.addWorksheet('Resumen');
    wsResumen.addRow(['STOCKPULSE - CONSOLIDADO']).font = { bold: true, size: 16 };
    wsResumen.addRow(['Actualización:', new Date().toLocaleString('es-PE')]);
    masterMeta.lineas.forEach(lin => {
      const total = fullData.filter(p => p.linea === lin).reduce((a, b) => a + b.stock, 0);
      wsResumen.addRow([lin, total]);
    });

    masterMeta.lineas.forEach(lin => {
      const ws = wbAll.addWorksheet(lin.substring(0, 31));
      applyProfessionalStyles(ws);
      addDataToSheet(ws, fullData.filter(p => p.linea === lin));
    });

    const masterFileName = `StockPulse_TODOS.xlsx`;
    for (const dir of outputDirs) await wbAll.xlsx.writeFile(path.join(dir, masterFileName));

    // --- GUARDAR METADATOS PARA EL DASHBOARD ---
    const outputJSON = {
      metadata: {
        lastUpdated: new Date().toISOString(),
        totalProducts: fullData.length,
        almacen: 'VES',
        sinStock: countSinStock,
        bajoStock: countBajoStock,
        status: 'OPERATIVO'
      },
      productos: fullData
    };

    const jsonPaths = [
      path.join(__dirname, '..', 'Data', 'productos_con_stock.json'),
      path.join(__dirname, '..', 'frontend', 'public', 'productos_con_stock.json')
    ];
    for (const p of jsonPaths) fs.writeFileSync(p, JSON.stringify(outputJSON, null, 2));

    console.log(`✅ Los 4 Snapshots han sido actualizados.`);

  } catch (error) {
    console.error('❌ Error:', error.message);
  }
}

runSnapshotUpdate();

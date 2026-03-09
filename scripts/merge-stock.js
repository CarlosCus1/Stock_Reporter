import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import xlsx from 'xlsx';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Colores
const PRIMARY_COLOR = 'FF13DAEC';
const SECONDARY_COLOR = 'FF0E7490';
const DARK_COLOR = 'FF0F172A';
const RED_BG = 'FFFEE2E2';
const YELLOW_BG = 'FFFEF3C7';
const GREEN_BG = 'FFD1FAE5';
const LINE_COLORS = ['FF3B82F6', 'FF10B981', 'FFF59E0B', 'FFEF4444', 'FF8B5CF6', 'FFEC4899', 'FF06B6D4', 'FFF97316'];

const getFechaPeru = () => {
  const ahora = new Date();
  return ahora.toLocaleString('es-PE', { timeZone: 'America/Lima', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false });
};

const applyProfessionalStyles = (worksheet) => {
  worksheet.columns = [
    { header: '#', key: 'item', width: 5 },
    { header: 'Código', key: 'sku', width: 12 },
    { header: 'EAN', key: 'ean', width: 18 },
    { header: 'Nombre del Producto', key: 'nombre', width: 45 },
    { header: 'U. x Caja', key: 'unBx', width: 10 },
    { header: 'Stock', key: 'stock', width: 10 },
    { header: 'Estado', key: 'estado', width: 12 }
  ];
  const headerRow = worksheet.getRow(1);
  headerRow.height = 30;
  headerRow.eachCell((cell) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: PRIMARY_COLOR } };
    cell.font = { bold: true, size: 11, color: { argb: 'FFFFFFFF' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  });
  worksheet.autoFilter = { from: 'A1', to: 'G1' };
  worksheet.views = [{ state: 'frozen', ySplit: 1 }];
};

const addDataToSheet = (worksheet, data) => {
  data.forEach((p, index) => {
    const row = worksheet.addRow([index + 1, p.sku, p.ean, p.nombre, p.unBx, p.stock, p.estado]);
    row.getCell(6).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: p.bgColor } };
    row.getCell(6).font = { bold: true, size: 12 };
    row.getCell(6).alignment = { horizontal: 'center' };
    row.getCell(7).value = p.estado;
    row.getCell(7).font = { bold: true, size: 12, color: { argb: p.fontColor } };
    row.getCell(7).alignment = { horizontal: 'center' };
  });
};

const createResumenSimple = (ws, fullData, masterMeta) => {
  // Configurar columnas
  ws.getColumn('A').width = 20;
  ws.getColumn('B').width = 15;
  ws.getColumn('C').width = 18;
  ws.getColumn('D').width = 15;
  
  // ===== ENCABEZADO =====
  ws.mergeCells('A1:D1');
  ws.getCell('A1').value = 'STOCKPULSE - RESUMEN';
  ws.getCell('A1').font = { bold: true, size: 16, color: { argb: 'FFFFFFFF' } };
  ws.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: SECONDARY_COLOR } };
  ws.getCell('A1').alignment = { vertical: 'middle', horizontal: 'center' };
  ws.getRow(1).height = 30;
  
  ws.mergeCells('A2:D2');
  ws.getCell('A2').value = `Actualización: ${getFechaPeru()}`;
  ws.getCell('A2').font = { size: 10, italic: true };
  ws.getCell('A2').alignment = { horizontal: 'right' };
  ws.getRow(2).height = 18;
  
  // ===== TOTAL GENERAL =====
  const totalUnidades = fullData.reduce((acc, p) => acc + p.stock, 0);
  const totalCodigos = fullData.length;
  const countSinStock = fullData.filter(p => p.stock === 0).length;
  const countBajoStock = fullData.filter(p => p.stock > 0 && p.stock < (p.unBx || 1) * 5).length;
  
  const rowTotal = 4;
  ws.mergeCells(`A${rowTotal}:D${rowTotal}`);
  ws.getCell(`A${rowTotal}`).value = 'TOTAL GENERAL';
  ws.getCell(`A${rowTotal}`).font = { bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
  ws.getCell(`A${rowTotal}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARK_COLOR } };
  ws.getCell(`A${rowTotal}`).alignment = { vertical: 'middle', horizontal: 'center' };
  ws.getRow(rowTotal).height = 25;
  
  // Valores totales
  ws.getCell('A6').value = 'Códigos:';
  ws.getCell('A6').font = { bold: true };
  ws.getCell('B6').value = totalCodigos;
  ws.getCell('B6').font = { bold: true, size: 14, color: { argb: SECONDARY_COLOR } };
  
  ws.getCell('C6').value = 'Unidades:';
  ws.getCell('C6').font = { bold: true };
  ws.getCell('D6').value = totalUnidades.toLocaleString('es-PE');
  ws.getCell('D6').font = { bold: true, size: 14, color: { argb: SECONDARY_COLOR } };
  ws.getRow(6).height = 22;
  
  // Estado rápido
  ws.getCell('A7').value = 'Agotados:';
  ws.getCell('A7').font = { bold: true };
  ws.getCell('B7').value = countSinStock;
  ws.getCell('B7').font = { bold: true, size: 12 };
  ws.getCell('B7').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: countSinStock > 0 ? RED_BG : GREEN_BG } };
  
  ws.getCell('C7').value = 'Bajo Stock:';
  ws.getCell('C7').font = { bold: true };
  ws.getCell('D7').value = countBajoStock;
  ws.getCell('D7').font = { bold: true, size: 12 };
  ws.getCell('D7').fill = { type: 'pattern', pattern: 'solid', fgColor: countBajoStock > 0 ? YELLOW_BG : GREEN_BG };
  ws.getRow(7).height = 22;
  
  // ===== POR LÍNEA =====
  const rowLinea = 10;
  ws.mergeCells(`A${rowLinea}:D${rowLinea}`);
  ws.getCell(`A${rowLinea}`).value = 'POR LÍNEA';
  ws.getCell(`A${rowLinea}`).font = { bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
  ws.getCell(`A${rowLinea}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARK_COLOR } };
  ws.getCell(`A${rowLinea}`).alignment = { vertical: 'middle', horizontal: 'center' };
  ws.getRow(rowLinea).height = 25;
  
  // Encabezados
  const rowHeaders = rowLinea + 1;
  ['Línea', 'Códigos', 'Unidades', '%'].forEach((h, i) => {
    const col = String.fromCharCode(65 + i);
    ws.getCell(`${col}${rowHeaders}`).value = h;
    ws.getCell(`${col}${rowHeaders}`).font = { bold: true, size: 10 };
    ws.getCell(`${col}${rowHeaders}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
    ws.getCell(`${col}${rowHeaders}`).alignment = { horizontal: 'center' };
  });
  ws.getRow(rowHeaders).height = 18;
  
  // Datos
  let currentRow = rowHeaders + 1;
  masterMeta.lineas.forEach((lin, index) => {
    const productosLinea = fullData.filter(p => p.linea === lin);
    const codigosLinea = productosLinea.length;
    const unidadesLinea = productosLinea.reduce((acc, p) => acc + p.stock, 0);
    const porcentaje = ((unidadesLinea / totalUnidades) * 100).toFixed(1);
    const color = LINE_COLORS[index % LINE_COLORS.length];
    
    ws.getCell(`A${currentRow}`).value = lin;
    ws.getCell(`A${currentRow}`).font = { bold: true, size: 10, color: { argb: color } };
    
    ws.getCell(`B${currentRow}`).value = codigosLinea;
    ws.getCell(`B${currentRow}`).alignment = { horizontal: 'center' };
    
    ws.getCell(`C${currentRow}`).value = unidadesLinea.toLocaleString('es-PE');
    ws.getCell(`C${currentRow}`).alignment = { horizontal: 'right' };
    
    ws.getCell(`D${currentRow}`).value = porcentaje + '%';
    ws.getCell(`D${currentRow}`).alignment = { horizontal: 'center' };
    ws.getCell(`D${currentRow}`).font = { bold: true, color: { argb: color } };
    
    currentRow++;
  });
  
  // Total
  const totalRow = currentRow;
  ws.getCell(`A${totalRow}`).value = 'TOTAL';
  ws.getCell(`A${totalRow}`).font = { bold: true, size: 10, color: { argb: 'FFFFFFFF' } };
  ws.getCell(`A${totalRow}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARK_COLOR } };
  
  ws.getCell(`B${totalRow}`).value = totalCodigos;
  ws.getCell(`B${totalRow}`).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  ws.getCell(`B${totalRow}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARK_COLOR } };
  ws.getCell(`B${totalRow}`).alignment = { horizontal: 'center' };
  
  ws.getCell(`C${totalRow}`).value = totalUnidades.toLocaleString('es-PE');
  ws.getCell(`C${totalRow}`).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  ws.getCell(`C${totalRow}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARK_COLOR } };
  ws.getCell(`C${totalRow}`).alignment = { horizontal: 'right' };
  
  ws.getCell(`D${totalRow}`).value = '100%';
  ws.getCell(`D${totalRow}`).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  ws.getCell(`D${totalRow}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARK_COLOR } };
  ws.getCell(`D${totalRow}`).alignment = { horizontal: 'center' };
  ws.getRow(totalRow).height = 20;
  
  // Pie
  ws.mergeCells(`A${totalRow + 2}:D${totalRow + 2}`);
  ws.getCell(`A${totalRow + 2}`).value = 'StockPulse - Inteligencia CIPSA';
  ws.getCell(`A${totalRow + 2}`).font = { italic: true, size: 9, color: { argb: 'FF94A3B8' } };
  ws.getCell(`A${totalRow + 2}`).alignment = { horizontal: 'center' };
};

async function runSnapshotUpdate() {
  try {
    console.log('🚀 Actualizando reportes...');

    const productosPath = path.join(__dirname, '..', 'Data', 'productos.json');
    const { productos, metadata: masterMeta } = JSON.parse(fs.readFileSync(productosPath, 'utf8'));

    const stockJsonPath = path.join(__dirname, '..', 'Data', 'data_stock.json');
    if (!fs.existsSync(stockJsonPath)) {
      throw new Error('No se encontró Data/data_stock.json');
    }
    const { stock: stockMap } = JSON.parse(fs.readFileSync(stockJsonPath, 'utf8'));

    const fullData = productos.map(p => {
      const stock = stockMap[p.sku] || 0;
      const minCajas = 5;
      const stockMinimo = (p.unBx || 1) * minCajas;
      let bgColor = GREEN_BG, fontColor = 'FF065F46', estado = '✓ OK';
      if (stock === 0) { bgColor = RED_BG; fontColor = 'FFDC2626'; estado = '✗ AGOTADO'; }
      else if (stock < stockMinimo) { bgColor = YELLOW_BG; fontColor = 'FFD97706'; estado = '⚠ BAJO'; }
      
      // Aseguramos que 'orden' sea un número para evitar ordenamiento lexicográfico
      const ordenNumerico = p.orden ? parseInt(p.orden, 10) : 0;
      
      return { ...p, orden: ordenNumerico, stock, estado, bgColor, fontColor };
    });

    const outputDirs = [
      path.join(__dirname, '..', 'reports'),
      path.join(__dirname, '..', 'frontend', 'public', 'reports')
    ];
    outputDirs.forEach(dir => { if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true }); });

    // Por categoría
    const categorias = ['PELOTAS', 'ESCOLAR', 'REPRESENTADAS'];
    for (const cat of categorias) {
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
        for (const dir of outputDirs) await workbook.xlsx.writeFile(path.join(dir, `StockPulse_${cat}.xlsx`));
      }
    }

    // Maestro
    const wbAll = new ExcelJS.Workbook();
    const wsResumen = wbAll.addWorksheet('Resumen');
    createResumenSimple(wsResumen, fullData, masterMeta);

    masterMeta.lineas.forEach(lin => {
      const ws = wbAll.addWorksheet(lin.substring(0, 31));
      applyProfessionalStyles(ws);
      addDataToSheet(ws, fullData.filter(p => p.linea === lin));
    });

    for (const dir of outputDirs) await wbAll.xlsx.writeFile(path.join(dir, 'StockPulse_TODOS.xlsx'));

    // JSON
    const countSinStock = fullData.filter(p => p.stock === 0).length;
    const countBajoStock = fullData.filter(p => p.stock > 0 && p.stock < (p.unBx || 1) * 5).length;
    
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

    console.log('✅ Reportes actualizados.');

  } catch (error) {
    console.error('❌ Error:', error.message);
  }
}

runSnapshotUpdate();

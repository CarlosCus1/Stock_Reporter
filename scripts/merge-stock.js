import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import xlsx from 'xlsx';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PRIMARY_COLOR = 'FF13DAEC';
const SECONDARY_COLOR = 'FF0E7490';
const DARK_COLOR = 'FF0F172A';
const RED_BG = 'FFFEE2E2';
const YELLOW_BG = 'FFFEF3C7';
const GREEN_BG = 'FFD1FAE5';
const RED_ARROW = '✗ AGOTADO';
const YELLOW_ARROW = '⚠ BAJO';
const GREEN_ARROW = '✓ OK';

// Colores para gráficos por línea
const LINE_COLORS = [
  'FF3B82F6', // azul
  'FF10B981', // verde
  'FFF59E0B', // amarillo
  'FFEF4444', // rojo
  'FF8B5CF6', // morado
  'FFEC4899', // rosa
  'FF06B6D4', // cyan
  'FFF97316', // naranja
];

const getFechaPeru = () => {
  const ahora = new Date();
  return ahora.toLocaleString('es-PE', { 
    timeZone: 'America/Lima',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: false
  });
};

// Normalizar SKU: mantener como string puro con sus ceros originales
const normalizeSKU = (sku) => String(sku || '').trim();

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
    // Celda Stock con color de fondo
    row.getCell(6).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: p.bgColor } };
    row.getCell(6).font = { bold: true, size: 12 };
    row.getCell(6).alignment = { horizontal: 'center' };
    // Celda Estado con flecha y color
    row.getCell(7).value = p.estado;
    row.getCell(7).font = { bold: true, size: 12, color: { argb: p.fontColor } };
    row.getCell(7).alignment = { horizontal: 'center' };
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
      const minCajas = 5; // Umbral de cajas para alerta de stock bajo
      const stockMinimo = (p.unBx || 1) * minCajas;
      let bgColor = GREEN_BG, fontColor = 'FF065F46', estado = '✓ OK';
      if (stock === 0) { bgColor = RED_BG; fontColor = 'FFDC2626'; estado = '✗ AGOTADO'; countSinStock++; }
      else if (stock < stockMinimo) { bgColor = YELLOW_BG; fontColor = 'FFD97706'; estado = '⚠ BAJO'; countBajoStock++; }
      return { ...p, stock, estado, bgColor, fontColor };
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

    // --- GENERAR SNAPSHOT MAESTRO (TODOS) CON RESUMEN PROFESIONAL ---
    const wbAll = new ExcelJS.Workbook();
    const wsResumen = wbAll.addWorksheet('Resumen');
    
    // Configurar ancho de columnas
    wsResumen.getColumn('A').width = 25;
    wsResumen.getColumn('B').width = 20;
    wsResumen.getColumn('C').width = 20;
    wsResumen.getColumn('D').width = 15;
    
    // ===== ENCABEZADO CON TÍTULO =====
    // Fila 1: Título principal
    wsResumen.mergeCells('A1:D1');
    const titleCell = wsResumen.getCell('A1');
    titleCell.value = 'STOCKPULSE - REPORTE CONSOLIDADO';
    titleCell.font = { bold: true, size: 16, color: { argb: 'FFFFFFFF' } };
    titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: SECONDARY_COLOR } };
    titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
    wsResumen.getRow(1).height = 30;
    
    // Fila 2: Fecha
    wsResumen.mergeCells('A2:D2');
    const subtitleCell = wsResumen.getCell('A2');
    subtitleCell.value = `Generado: ${getFechaPeru()}`;
    subtitleCell.font = { size: 10, italic: true, color: { argb: 'FF64748B' } };
    subtitleCell.alignment = { horizontal: 'right' };
    wsResumen.getRow(2).height = 18;
    
    // ===== SECCIÓN: RESUMEN EJECUTIVO =====
    const rowResumen = 4;
    wsResumen.mergeCells(`A${rowResumen}:D${rowResumen}`);
    const resumenHeader = wsResumen.getCell(`A${rowResumen}`);
    resumenHeader.value = 'RESUMEN EJECUTIVO';
    resumenHeader.font = { bold: true, size: 11, color: { argb: 'FFFFFFFF' } };
    resumenHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARK_COLOR } };
    resumenHeader.alignment = { vertical: 'middle', horizontal: 'center' };
    wsResumen.getRow(rowResumen).height = 22;
    
    // Calcular totales
    const totalUnidades = fullData.reduce((acc, p) => acc + p.stock, 0);
    const totalCodigos = fullData.length;
    const countOK = fullData.length - countSinStock - countBajoStock;
    
    // Fila 5: Total general
    wsResumen.getCell('A5').value = 'Total Códigos:';
    wsResumen.getCell('A5').font = { bold: true };
    wsResumen.getCell('B5').value = totalCodigos;
    wsResumen.getCell('B5').font = { bold: true, size: 13, color: { argb: SECONDARY_COLOR } };
    
    wsResumen.getCell('C5').value = 'Stock Total:';
    wsResumen.getCell('C5').font = { bold: true };
    wsResumen.getCell('D5').value = totalUnidades.toLocaleString('es-PE');
    wsResumen.getCell('D5').font = { bold: true, size: 13, color: { argb: SECONDARY_COLOR } };
    wsResumen.getRow(5).height = 22;
    
    // Fila 6: Estado del inventario
    wsResumen.getCell('A6').value = 'Productos OK:';
    wsResumen.getCell('B6').value = countOK;
    wsResumen.getCell('B6').font = { bold: true, size: 11, color: { argb: 'FF059669' } };
    wsResumen.getCell('B6').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GREEN_BG } };
    wsResumen.getCell('B6').alignment = { horizontal: 'center' };
    
    wsResumen.getCell('C6').value = 'Bajo Stock:';
    wsResumen.getCell('D6').value = countBajoStock;
    wsResumen.getCell('D6').font = { bold: true, size: 11, color: { argb: 'FFD97706' } };
    wsResumen.getCell('D6').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: YELLOW_BG } };
    wsResumen.getCell('D6').alignment = { horizontal: 'center' };
    
    wsResumen.getCell('E6').value = 'Agotados:';
    wsResumen.getCell('F6').value = countSinStock;
    wsResumen.getCell('F6').font = { bold: true, size: 11, color: { argb: 'FFDC2626' } };
    wsResumen.getCell('F6').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: RED_BG } };
    wsResumen.getCell('F6').alignment = { horizontal: 'center' };
    wsResumen.getRow(6).height = 22;
    
    // ===== SECCIÓN: DETALLE POR LÍNEA =====
    const rowLinea = 9;
    wsResumen.mergeCells(`A${rowLinea}:D${rowLinea}`);
    const lineaHeader = wsResumen.getCell(`A${rowLinea}`);
    lineaHeader.value = 'DETALLE POR LÍNEA';
    lineaHeader.font = { bold: true, size: 11, color: { argb: 'FFFFFFFF' } };
    lineaHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARK_COLOR } };
    lineaHeader.alignment = { vertical: 'middle', horizontal: 'center' };
    wsResumen.getRow(rowLinea).height = 22;
    
    // Encabezados de tabla por línea
    const rowHeaders = rowLinea + 1;
    wsResumen.getCell(`A${rowHeaders}`).value = 'Línea';
    wsResumen.getCell(`B${rowHeaders}`).value = 'Códigos';
    wsResumen.getCell(`C${rowHeaders}`).value = 'Unidades';
    wsResumen.getCell(`D${rowHeaders}`).value = '% Part.';
    
    [ 'A', 'B', 'C', 'D' ].forEach(col => {
      const cell = wsResumen.getCell(`${col}${rowHeaders}`);
      cell.font = { bold: true, size: 10 };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
      cell.alignment = { horizontal: 'center' };
    });
    wsResumen.getRow(rowHeaders).height = 18;
    
    // Datos por línea
    let currentRow = rowHeaders + 1;
    masterMeta.lineas.forEach((lin, index) => {
      const productosLinea = fullData.filter(p => p.linea === lin);
      const codigosLinea = productosLinea.length;
      const unidadesLinea = productosLinea.reduce((acc, p) => acc + p.stock, 0);
      const porcentaje = ((unidadesLinea / totalUnidades) * 100).toFixed(1);
      const color = LINE_COLORS[index % LINE_COLORS.length];
      
      wsResumen.getCell(`A${currentRow}`).value = lin;
      wsResumen.getCell(`A${currentRow}`).font = { bold: true, size: 10, color: { argb: color } };
      
      wsResumen.getCell(`B${currentRow}`).value = codigosLinea;
      wsResumen.getCell(`B${currentRow}`).alignment = { horizontal: 'center' };
      wsResumen.getCell(`B${currentRow}`).font = { size: 10 };
      
      wsResumen.getCell(`C${currentRow}`).value = unidadesLinea.toLocaleString('es-PE');
      wsResumen.getCell(`C${currentRow}`).alignment = { horizontal: 'right' };
      wsResumen.getCell(`C${currentRow}`).font = { size: 10 };
      
      wsResumen.getCell(`D${currentRow}`).value = `${porcentaje}%`;
      wsResumen.getCell(`D${currentRow}`).alignment = { horizontal: 'center' };
      wsResumen.getCell(`D${currentRow}`).font = { bold: true, size: 10, color: { argb: color } };
      
      currentRow++;
    });
    
    // Totales finales
    const totalRow = currentRow;
    wsResumen.getCell(`A${totalRow}`).value = 'TOTAL';
    wsResumen.getCell(`A${totalRow}`).font = { bold: true, size: 11, color: { argb: 'FFFFFFFF' } };
    wsResumen.getCell(`A${totalRow}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARK_COLOR } };
    
    wsResumen.getCell(`B${totalRow}`).value = totalCodigos;
    wsResumen.getCell(`B${totalRow}`).font = { bold: true, size: 11, color: { argb: 'FFFFFFFF' } };
    wsResumen.getCell(`B${totalRow}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARK_COLOR } };
    wsResumen.getCell(`B${totalRow}`).alignment = { horizontal: 'center' };
    
    wsResumen.getCell(`C${totalRow}`).value = totalUnidades.toLocaleString('es-PE');
    wsResumen.getCell(`C${totalRow}`).font = { bold: true, size: 11, color: { argb: 'FFFFFFFF' } };
    wsResumen.getCell(`C${totalRow}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARK_COLOR } };
    wsResumen.getCell(`C${totalRow}`).alignment = { horizontal: 'right' };
    
    wsResumen.getCell(`D${totalRow}`).value = '100%';
    wsResumen.getCell(`D${totalRow}`).font = { bold: true, size: 11, color: { argb: 'FFFFFFFF' } };
    wsResumen.getCell(`D${totalRow}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARK_COLOR } };
    wsResumen.getCell(`D${totalRow}`).alignment = { horizontal: 'center' };
    wsResumen.getRow(totalRow).height = 20;
    
    // Pie de página
    const footerRow = totalRow + 2;
    wsResumen.mergeCells(`A${footerRow}:D${footerRow}`);
    wsResumen.getCell(`A${footerRow}`).value = 'Este reporte se actualiza automaticamente. StockPulse - Inteligencia CIPSA';
    wsResumen.getCell(`A${footerRow}`).font = { italic: true, size: 9, color: { argb: 'FF94A3B8' } };
    wsResumen.getCell(`A${footerRow}`).alignment = { horizontal: 'center' };

    // Agregar hojas por línea
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

const express = require('express');
const cors = require('cors');
const xlsx = require('xlsx');
const https = require('https');
const http = require('http');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3001;

// Configuración de rutas para archivos
const DATA_DIR = path.join(__dirname, '..', 'Data');
const FRONTEND_PUBLIC_DIR = path.join(__dirname, '..', 'frontend', 'public');
const PRODUCTOS_JSON_PATH = path.join(DATA_DIR, 'productos.json');
const DATA_STOCK_JSON_PATH = path.join(DATA_DIR, 'data_stock.json');
const OUTPUT_JSON_PATH = path.join(DATA_DIR, 'productos_con_stock.json');
const OUTPUT_JSON_PUBLIC_PATH = path.join(FRONTEND_PUBLIC_DIR, 'productos_con_stock.json');

// ============================================
// FUNCIÓN DE LIMPIEZA SEMANAL DE LOGS
// ============================================
function limpiarLogsSemanales() {
  const logFile = path.join(__dirname, 'logs', 'descargas.json');
  const SEMANA_MS = 7 * 24 * 60 * 60 * 1000; // 7 días en milisegundos
  
  try {
    if (fs.existsSync(logFile)) {
      const data = fs.readFileSync(logFile, 'utf8');
      const log = JSON.parse(data);
      
      // Verificar última limpieza
      const ultimaLimpieza = log.ultima_limpieza ? new Date(log.ultima_limpieza) : null;
      const ahora = new Date();
      
      if (!ultimaLimpieza || (ahora - ultimaLimpieza) > SEMANA_MS) {
        // Limpiar: solo mantener la estructura sin registros
        const logLimpio = {
          descargas: [],
          ultima_limpieza: ahora.toISOString(),
          nota: "Log limpiado automáticamente - historial de descargas"
        };
        fs.writeFileSync(logFile, JSON.stringify(logLimpio, null, 2));
        console.log('🧹 Logs de descargas limpiados automáticamente (semanal)');
      }
    }
  } catch (e) {
    console.error('Error en limpieza de logs:', e.message);
  }
}

// Ejecutar limpieza al iniciar
limpiarLogsSemanales();

// Habilitar CORS para permitir solicitudes desde el frontend
app.use(cors());
app.use(express.json());

// Función para descargar archivo desde una URL
function downloadFile(url) {
  return new Promise((resolve, reject) => {
    const protocol = url.startsWith('https') ? https : http;
    
    const options = {
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
      }
    };
    
    protocol.get(url, options, (response) => {
      // Manejar redirecciones
      if (response.statusCode === 301 || response.statusCode === 302) {
        downloadFile(response.headers.location)
          .then(resolve)
          .catch(reject);
        return;
      }
      
      if (response.statusCode !== 200) {
        reject(new Error(`Error HTTP: ${response.statusCode}`));
        return;
      }

      const chunks = [];
      response.on('data', (chunk) => chunks.push(chunk));
      response.on('end', () => {
        resolve(Buffer.concat(chunks));
      });
      response.on('error', reject);
    }).on('error', reject);
  });
}

// Endpoint: Descargar stock desde appweb
app.get('/api/stock/appweb', async (req, res) => {
  try {
    console.log('Descargando desde appweb...');
    
    const url = 'http://appweb.cipsa.com.pe:8054/AlmacenStock/DownLoadFiles?value=%7B%22parametroX1%22:%220%22,%22parametroX2%22:%220%22%7D';
    
    const buffer = await downloadFile(url);
    console.log('Archivo descargado, parseando...');
    
    // Parsear Excel
    const workbook = xlsx.read(buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rawData = xlsx.utils.sheet_to_json(worksheet);
    
    // Procesar datos
    const stockData = rawData
      .filter(row => row.Column2 && row.Column2 !== 'Total' && row.Column2 !== 'TOTAL')
      .map(row => ({
        sku: String(row.Column2).trim(),
        nombre: row.Column3 || '',
        disponible: parseInt(row.Column19, 10) || 0,
        almacen: row.Column10 || '',
        predespacho: parseInt(row.Column17, 10) || 0
      }));
    
    console.log(`Stock procesado: ${stockData.length} productos`);
    
    res.json({
      success: true,
      count: stockData.length,
      data: stockData
    });
  } catch (error) {
    console.error('Error:', error.message);
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// Endpoint: Descargar stock desde Google Drive
app.get('/api/stock/drive', async (req, res) => {
  try {
    const { url } = req.query;
    
    if (!url) {
      return res.status(400).json({
        success: false,
        error: 'Se requiere la URL de Google Drive'
      });
    }
    
    console.log('Descargando desde Google Drive:', url);
    
    // Extraer ID del archivo
    let fileId = '';
    if (url.includes('/d/')) {
      const parts = url.split('/d/');
      const secondPart = parts[1].split('/');
      fileId = secondPart[0];
    } else if (url.includes('file/d/')) {
      const parts = url.split('file/d/');
      const secondPart = parts[1].split('/');
      fileId = secondPart[0];
    }
    
    if (!fileId) {
      throw new Error('No se pudo extraer el ID del archivo');
    }
    
    // Descargar el archivo desde Google Drive
    const buffer = await downloadFile(url);
    
    // Intentar parsear como Excel o JSON
    let stockData;
    try {
      // Intentar como Excel
      const workbook = xlsx.read(buffer, { type: 'buffer' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const rawData = xlsx.utils.sheet_to_json(worksheet);
      
      stockData = rawData
        .filter(row => row.Column2 && row.Column2 !== 'Total' && row.Column2 !== 'TOTAL')
        .map(row => ({
          sku: String(row.Column2).trim(),
          nombre: row.Column3 || '',
          disponible: parseInt(row.Column19, 10) || 0
        }));
    } catch (e) {
      // Intentar como JSON
      const jsonStr = buffer.toString('utf-8');
      const jsonData = JSON.parse(jsonStr);
      
      stockData = Array.isArray(jsonData) ? jsonData : jsonData.productos || jsonData.stock || [];
      
      stockData = stockData.map(item => ({
        sku: String(item.sku || item.codigo || item.Column2 || '').trim(),
        nombre: item.nombre || item.nombre_producto || '',
        disponible: parseInt(item.disponible || item.stock || item.Column19 || 0, 10)
      })).filter(item => item.sku);
    }
    
    console.log(`Stock procesado: ${stockData.length} productos`);
    
    res.json({
      success: true,
      count: stockData.length,
      data: stockData
    });
  } catch (error) {
    console.error('Error:', error.message);
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// Endpoint: Health check
app.get('/api/health', (req, res) => {
  res.json({ status: 'OK', timestamp: new Date().toISOString() });
});

// Endpoint: Forzar actualización de datos (regenerar JSON)
app.post('/api/stock/actualizar', async (req, res) => {
  try {
    console.log('🔄 Iniciando actualización de datos...');
    
    // 1. Verificar que existen los archivos de códigos
    if (!fs.existsSync(PRODUCTOS_JSON_PATH)) {
      return res.status(500).json({ 
        success: false, 
        error: 'No se encontró el archivo de códigos de productos. Ejecute los scripts de merge localmente.' 
      });
    }
    
    const productosData = JSON.parse(fs.readFileSync(PRODUCTOS_JSON_PATH, 'utf8'));
    const productos = productosData.productos || [];
    console.log(`✅ Códigos de productos cargados: ${productos.length}`);
    
    // 2. Descargar stock desde appweb
    const stockUrl = 'http://appweb.cipsa.com.pe:8054/AlmacenStock/DownLoadFiles?value={%22%20%22:%22%22,%22parametroX1%22:%220%22,%22parametroX2%22:%220%22}';
    const buffer = await downloadFile(stockUrl);
    
    // Parsear Excel
    const workbook = xlsx.read(buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rawData = xlsx.utils.sheet_to_json(worksheet);
    
    // Crear mapa de stock
    const stockMap = {};
    rawData.forEach(row => {
      if (row.Column2 && row.Column2 !== 'Total' && row.Column2 !== 'TOTAL') {
        const sku = String(row.Column2).trim();
        stockMap[sku] = parseInt(row.Column19, 10) || 0;
      }
    });
    console.log(`✅ Stock descargado: ${Object.keys(stockMap).length} productos`);
    
    // 3. Merge: combinar productos con stock
    let countSinStock = 0;
    let countBajoStock = 0;
    
    const fullData = productos.map(p => {
      const stock = stockMap[p.sku] || 0;
      const minCajas = 5;
      const stockMinimo = (p.unBx || 1) * minCajas;
      let estado = 'OK';
      if (stock === 0) { estado = 'AGOTADO'; countSinStock++; }
      else if (stock < stockMinimo) { estado = 'BAJO'; countBajoStock++; }
      return { ...p, stock, estado };
    });
    
    // 4. Guardar JSON
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
    
    // Guardar en Data/
    fs.writeFileSync(OUTPUT_JSON_PATH, JSON.stringify(outputJSON, null, 2));
    console.log(`✅ JSON guardado en: ${OUTPUT_JSON_PATH}`);
    
    // Guardar en frontend/public/
    const frontendDir = path.dirname(OUTPUT_JSON_PUBLIC_PATH);
    if (!fs.existsSync(frontendDir)) {
      fs.mkdirSync(frontendDir, { recursive: true });
    }
    fs.writeFileSync(OUTPUT_JSON_PUBLIC_PATH, JSON.stringify(outputJSON, null, 2));
    console.log(`✅ JSON guardado en: ${OUTPUT_JSON_PUBLIC_PATH}`);
    
    res.json({
      success: true,
      message: 'Datos actualizados correctamente',
      metadata: outputJSON.metadata
    });
    
  } catch (error) {
    console.error('❌ Error en actualización:', error.message);
    res.status(500).json({ 
      success: false, 
      error: error.message 
    });
  }
});

// Endpoint: Registrar descarga de reporte
app.post('/api/descargas/registrar', (req, res) => {
  try {
    const { nombre, email, categoria, timestamp } = req.body;
    
    // Validar datos requeridos
    if (!nombre || !email || !categoria) {
      return res.status(400).json({
        success: false,
        error: 'Faltan datos requeridos: nombre, email, categoria'
      });
    }
    
    // Validar formato de email
    const emailValido = /^[\w.-]+@[\w.-]+\.\w{2,}$/.test(email);
    if (!emailValido) {
      return res.status(400).json({
        success: false,
        error: 'Email inválido'
      });
    }
    
    // Validar que sea email corporativo de CIPSA
    if (!email.endsWith('@cipsa.com.pe')) {
      return res.status(403).json({
        success: false,
        error: 'Solo se permiten emails corporativos de CIPSA'
      });
    }
    
    const registro = {
      id: Date.now().toString(36) + Math.random().toString(36).substr(2, 5),
      nombre: nombre.trim(),
      email: email.toLowerCase().trim(),
      categoria,
      timestamp: timestamp || new Date().toISOString(),
      ip: req.ip || req.connection.remoteAddress || 'desconocido',
      userAgent: req.headers['user-agent'] || 'desconocido'
    };
    
    // Guardar en archivo JSON
    const logFile = path.join(__dirname, 'logs', 'descargas.json');
    let logData = { descargas: [], ultima_limpieza: null };
    
    // Crear directorio si no existe
    const logsDir = path.join(__dirname, 'logs');
    if (!fs.existsSync(logsDir)) {
      fs.mkdirSync(logsDir, { recursive: true });
    }
    
    // Leer archivo existente
    if (fs.existsSync(logFile)) {
      try {
        const data = fs.readFileSync(logFile, 'utf8');
        logData = JSON.parse(data);
        // Asegurar estructura
        if (!logData.descargas) logData.descargas = [];
        if (!logData.ultima_limpieza) logData.ultima_limpieza = null;
      } catch (e) {
        logData = { descargas: [], ultima_limpieza: null };
      }
    }
    
    // Agregar nuevo registro
    logData.descargas.unshift(registro); // Agregar al inicio
    
    // Mantener solo últimos 1000 registros
    if (logData.descargas.length > 1000) {
      logData.descargas = logData.descargas.slice(0, 1000);
    }
    
    // Guardar archivo
    fs.writeFileSync(logFile, JSON.stringify(logData, null, 2));
    
    console.log(`📥 Descarga registrada: ${nombre} (${email}) - ${categoria}`);
    
    res.json({
      success: true,
      message: 'Descarga registrada correctamente',
      registroId: registro.id
    });
  } catch (error) {
    console.error('Error al registrar descarga:', error.message);
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// Endpoint: Ver historial de descargas (para administración)
app.get('/api/descargas/historial', (req, res) => {
  try {
    const logFile = path.join(__dirname, 'logs', 'descargas.json');
    
    if (!fs.existsSync(logFile)) {
      return res.json({ success: true, descargas: [], total: 0, ultima_limpieza: null });
    }
    
    const data = fs.readFileSync(logFile, 'utf8');
    const logData = JSON.parse(data);
    
    res.json({
      success: true,
      total: logData.descargas ? logData.descargas.length : 0,
      ultima_limpieza: logData.ultima_limpieza || null,
      descargas: logData.descargas || []
    });
  } catch (error) {
    console.error('Error al leer historial:', error.message);
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

app.listen(PORT, () => {
  console.log(`Stock API corriendo en http://localhost:${PORT}`);
  console.log('Endpoints disponibles:');
  console.log(`  - GET http://localhost:${PORT}/api/stock/appweb`);
  console.log(`  - GET http://localhost:${PORT}/api/stock/drive?url=...`);
});

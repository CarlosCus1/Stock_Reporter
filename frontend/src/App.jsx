/**
 * @file App.jsx
 * @author Carlos Cusi
 * @description Componente principal de StockPulse con Tooltips informativos.
 */

import { useState, useEffect, useMemo } from 'react'
import { validarCorreo } from './utils/reportUtils'

const CATEGORIAS = [
  { id: 'todos', nombre: 'Todos', hint: 'Reporte completo de todas las líneas' },
  { id: 'pelotas', nombre: 'Pelotas', hint: 'Solo productos de la línea Pelotas y Mascotas' },
  { id: 'escolar', nombre: 'Escolar', hint: 'Solo productos de la línea Escolar (Forros, Archivo, etc.)' },
  { id: 'representadas', nombre: 'Representadas', hint: 'Solo productos de marcas representadas' },
]

function App() {
  const [activeTab, setActiveTab] = useState('pulso')
  const [form, setForm] = useState({ nombre: '', email: '', categoria: 'todos' })
  const [ui, setUi] = useState({ 
    isValido: false, reporteGenerado: false, theme: 'light',
    isSearching: false, searchTerm: '', allProducts: [], 
    metadata: { lastUpdated: '', totalProducts: 0, almacen: 'Cipsa', sinStock: 0, bajoStock: 0, status: '...' }
  })

  useEffect(() => {
    const loadProducts = async () => {
      try {
        const baseUrl = import.meta.env.BASE_URL;
        const res = await fetch(`${baseUrl}productos_con_stock.json`)
        if (res.ok) {
          const data = await res.json()
          setUi(prev => ({ ...prev, allProducts: data.productos || [], metadata: data.metadata || prev.metadata }))
        }
      } catch (err) { console.error('Error cargando data:', err) }
    }
    loadProducts()
  }, [])

  const searchResults = useMemo(() => {
    const term = ui.searchTerm.trim().toLowerCase()
    if (!term || term.length < 2) return []
    return ui.allProducts.filter(p => 
      (p.sku && String(p.sku).toLowerCase().includes(term)) || 
      (p.nombre && p.nombre.toLowerCase().includes(term))
    ).slice(0, 15)
  }, [ui.searchTerm, ui.allProducts])

  useEffect(() => {
    const savedTheme = localStorage.getItem('theme') || 'light'
    setUi(prev => ({ ...prev, theme: savedTheme }))
    document.documentElement.classList.toggle('dark', savedTheme === 'dark')
  }, [])

  const toggleTheme = () => {
    const newTheme = ui.theme === 'light' ? 'dark' : 'light'
    setUi(prev => ({ ...prev, theme: newTheme }))
    document.documentElement.classList.toggle('dark', newTheme === 'dark')
    localStorage.setItem('theme', newTheme)
  }

  useEffect(() => {
    setUi(prev => ({ ...prev, isValido: validarCorreo(form.email) && form.nombre.length > 2 }))
  }, [form.email, form.nombre])
const handleDescargar = async () => {
  const baseUrl = import.meta.env.BASE_URL;
  const catSuffix = (form.categoria || 'todos').toUpperCase();
  const serverFileName = `StockPulse_${catSuffix}.xlsx`; 
  const downloadPath = `${baseUrl}reports/${serverFileName}`;

  try {
    // Validar si el archivo existe antes de descargar
    const response = await fetch(downloadPath);
    const contentType = response.headers.get('content-type');

    if (!response.ok || (contentType && contentType.includes('text/html'))) {
      throw new Error('El reporte aún no ha sido procesado por el servidor para hoy.');
    }

    // Si es válido, disparar la descarga con renombrado
    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);

    const opciones = { day: '2-digit', month: '2-digit', year: '2-digit' };
    const fecha = new Intl.DateTimeFormat('es-PE', opciones).format(new Date()).replace(/\//g, '-');
    const downloadName = `StockPulse_${catSuffix}_${fecha}.xlsx`;

    const link = document.createElement('a');
    link.href = url;
    link.download = downloadName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    window.URL.revokeObjectURL(url);
  } catch (err) {
    alert(err.message);
  }
  }

  return (
  <div className="bg-background-light dark:bg-background-dark font-display text-slate-900 dark:text-slate-100 min-h-screen flex flex-col relative transition-colors duration-500">
      {/* SEARCH OVERLAY */}
      {ui.isSearching && (
        <div className="fixed inset-0 z-50 bg-white/98 dark:bg-background-dark/98 backdrop-blur-xl flex flex-col animate-in fade-in duration-200">
          <div className="p-6 flex items-center gap-4 border-b border-slate-200 dark:border-white/5">
            <span className="material-symbols-outlined text-primary text-3xl">search</span>
            <input 
              autoFocus className="flex-1 bg-transparent border-none text-xl font-bold focus:ring-0 outline-none h-12"
              placeholder="SKU, EAN o nombre..." value={ui.searchTerm}
              title="Ingrese al menos 2 caracteres para buscar"
              onChange={(e) => setUi(prev => ({ ...prev, searchTerm: e.target.value }))}
            />
            <button onClick={() => setUi(prev => ({ ...prev, isSearching: false, searchTerm: '' }))} className="size-12 rounded-full bg-slate-100 dark:bg-white/5 flex items-center justify-center hover:bg-red-50 transition-colors">
              <span className="material-symbols-outlined">close</span>
            </button>
          </div>
          <div className="flex-1 overflow-y-auto p-6 space-y-4">
            {searchResults.map(p => (
              <div key={p.sku} className="p-5 bg-white dark:bg-white/5 rounded-2xl border border-slate-100 dark:border-white/5" title={`Línea: ${p.linea}`}>
                <div className="flex items-start justify-between mb-3">
                  <div>
                    <p className="text-xl font-bold text-primary">{p.sku}</p>
                    {p.ean && <p className="text-xs text-slate-400 font-mono mt-1">EAN: {p.ean}</p>}
                  </div>
                  <div className="text-right">
                    <p className="text-xl font-bold">{p.stock}</p>
                    <p className="text-[9px] uppercase opacity-50">{ui.metadata.almacen}</p>
                  </div>
                </div>
                <p className="font-medium text-sm opacity-80">{p.nombre}</p>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* HEADER */}
      <header className="flex items-center p-4 justify-between border-b sticky top-0 z-40 bg-white/80 dark:bg-background-dark/80 backdrop-blur-xl border-slate-200 dark:border-white/5">
        <div className="flex size-11 shrink-0 items-center overflow-hidden rounded-2xl bg-primary/10 p-1">
          <img src="favicon.svg" alt="Logo" className="h-full w-full object-contain" />
        </div>
        <div className="flex flex-1 px-4 flex-col">
          <h2 className="text-xl leading-none font-bold">StockPulse</h2>
          <p className="text-[9px] font-bold text-slate-400 dark:text-slate-500 uppercase tracking-widest mt-1">Inteligencia • CIPSA</p>
        </div>
        <button 
          onClick={() => setUi(prev => ({ ...prev, isSearching: true }))} 
          className="size-11 flex items-center justify-center rounded-2xl bg-slate-100 dark:bg-white/5 hover:text-primary transition-all"
          title="Consulta rápida de stock por SKU"
        >
          <span className="material-symbols-outlined">search</span>
        </button>
      </header>

      {/* MAIN CONTENT */}
      <main className="flex-1 overflow-y-auto pb-32 max-w-2xl mx-auto w-full px-6">
        
        {activeTab === 'pulso' ? (
          <div className="animate-in fade-in slide-in-from-left-4 duration-500">
            <div className="pt-10 pb-4">
              <h1 className="text-3xl font-bold">Generar Reporte</h1>
              <p className="text-slate-500 mt-2" title="Nombre del almacén de origen">Almacén activo: {ui.metadata.almacen}</p>
            </div>

            <form onSubmit={(e) => {e.preventDefault(); setUi(prev => ({...prev, reporteGenerado: true}))}} className="py-6 space-y-8">
              <div className="space-y-6 bg-white dark:bg-white/2 p-6 rounded-3xl border border-slate-100 dark:border-white/5 shadow-sm">
                <div className="flex flex-col gap-2">
                  <label className="text-[13px] font-bold uppercase opacity-60">Solicitante</label>
                  <input className="input-field" placeholder="Su nombre" value={form.nombre} title="Nombre de quien descarga el reporte" onChange={(e) => setForm({ ...form, nombre: e.target.value })} />
                </div>
                <div className="flex flex-col gap-2">
                  <label className="text-[13px] font-bold uppercase opacity-60">Email Corporativo</label>
                  <input className="input-field" placeholder="usuario@cipsa.com.pe" value={form.email} title="Debe usar su correo de @cipsa.com.pe" onChange={(e) => setForm({ ...form, email: e.target.value })} />
                </div>
                <div className="flex flex-col gap-4">
                  <label className="text-[13px] font-bold uppercase opacity-60">Filtro de Reporte</label>
                  <div className="flex flex-wrap gap-2.5">
                    {CATEGORIAS.map(cat => (
                      <button 
                        key={cat.id} type="button" 
                        className={`chip ${form.categoria === cat.id ? 'chip-active' : 'chip-inactive'}`} 
                        onClick={() => setForm({ ...form, categoria: cat.id })}
                        title={cat.hint}
                      >
                        {cat.nombre}
                      </button>
                    ))}
                  </div>
                </div>
              </div>
              <button type="submit" disabled={!ui.isValido} className="btn-primary" title={!ui.isValido ? "Complete los campos para habilitar" : "Generar snapshot de inventario"}>
                <span className="material-symbols-outlined">analytics</span> Generar Inventario
              </button>
            </form>

            {ui.reporteGenerado && (
              <div className="card-success animate-in fade-in zoom-in-95 mt-4 shadow-lg shadow-emerald-500/10">
                <h4 className="font-bold text-xl mb-4">¡Snapshot Listo!</h4>
                <p className="text-sm opacity-70 mb-6">Se ha capturado el stock actual. El archivo se guardará con la fecha de hoy.</p>
                <button onClick={handleDescargar} className="w-full h-14 bg-slate-900 dark:bg-white text-white dark:text-slate-900 font-bold rounded-2xl flex items-center justify-center gap-2 transition-transform active:scale-95" title="Descargar Excel con renombrado automático">
                  <span className="material-symbols-outlined">download</span> DESCARGAR AHORA
                </button>
              </div>
            )}
          </div>
        ) : (
          <div className="animate-in fade-in slide-in-from-right-4 duration-500 pt-10">
            <h1 className="text-3xl font-bold mb-8 text-primary">Salud del Sistema</h1>
            
            <div className="grid grid-cols-2 gap-4">
              <div className="p-6 bg-white dark:bg-white/5 rounded-3xl border border-slate-100 dark:border-white/5 shadow-sm" title="Hora en que el bot procesó los datos por última vez">
                <p className="text-[10px] font-bold uppercase opacity-50 mb-2 tracking-tighter">Último Pulso</p>
                <p className="text-lg font-bold">{new Date(ui.metadata.lastUpdated).toLocaleTimeString()}</p>
                <p className="text-[10px] opacity-40">{new Date(ui.metadata.lastUpdated).toLocaleDateString()}</p>
              </div>
              <div className="p-6 bg-white dark:bg-white/5 rounded-3xl border border-slate-100 dark:border-white/5 shadow-sm" title="Indica que el motor de sincronización está activo">
                <p className="text-[10px] font-bold uppercase opacity-50 mb-2 tracking-tighter">Estado Bot</p>
                <div className="flex items-center gap-2 text-emerald-500">
                  <span className="relative flex h-3 w-3">
                    <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75"></span>
                    <span className="relative inline-flex rounded-full h-3 w-3 bg-emerald-500"></span>
                  </span>
                  <p className="font-bold">CONECTADO</p>
                </div>
              </div>
              <div className="p-6 bg-white dark:bg-white/5 rounded-3xl border border-slate-100 dark:border-white/5 col-span-2 shadow-sm">
                <p className="text-[10px] font-bold uppercase opacity-50 mb-4 tracking-widest text-center">Métricas del Inventario ({ui.metadata.almacen})</p>
                <div className="flex justify-between items-center px-4">
                  <div className="text-center" title="Total de SKUs en el maestro">
                    <p className="text-2xl font-bold">{ui.metadata.totalProducts}</p>
                    <p className="text-[10px] opacity-50">ITEMS</p>
                  </div>
                  <div className="h-10 w-[1px] bg-slate-200 dark:bg-white/10"></div>
                  <div className="text-center text-red-500" title="Productos con stock físico cero">
                    <p className="text-2xl font-bold">{ui.metadata.sinStock}</p>
                    <p className="text-[10px] opacity-50">AGOTADOS</p>
                  </div>
                  <div className="h-10 w-[1px] bg-slate-200 dark:bg-white/10"></div>
                  <div className="text-center text-amber-500" title="Productos con menos de 10 unidades disponibles">
                    <p className="text-2xl font-bold">{ui.metadata.bajoStock}</p>
                    <p className="text-[10px] opacity-50">CRÍTICOS</p>
                  </div>
                </div>
              </div>
            </div>

            <div className="mt-8 p-6 bg-primary/5 rounded-3xl border border-primary/10 flex gap-4 items-start">
              <span className="material-symbols-outlined text-primary">help</span>
              <p className="text-xs leading-relaxed opacity-80">
                StockPulse procesa automáticamente los datos de los almacenes para ofrecer información estratégica. Los reportes generados son propiedad de CIPSA y contienen información confidencial.
              </p>
            </div>
          </div>
        )}
      </main>

      {/* NAV INFERIOR */}
      <nav className="fixed bottom-0 left-0 right-0 bg-white/90 dark:bg-background-dark/90 backdrop-blur-2xl border-t border-slate-200 dark:border-white/5 px-6 pb-10 pt-4 z-40">
        <div className="max-w-md mx-auto flex justify-between items-center px-4">
          <button onClick={() => setActiveTab('pulso')} className={`flex flex-col items-center gap-1.5 transition-all ${activeTab === 'pulso' ? 'text-primary' : 'opacity-40'}`} title="Panel de Descargas">
            <span className="material-symbols-outlined text-[26px]">analytics</span>
            <span className="text-[9px] font-bold tracking-widest">PULSO</span>
          </button>
          <button onClick={() => setActiveTab('estado')} className={`flex flex-col items-center gap-1.5 transition-all ${activeTab === 'estado' ? 'text-primary' : 'opacity-40'}`} title="Dashboard de Salud de Datos">
            <span className="material-symbols-outlined text-[26px]">monitor_heart</span>
            <span className="text-[9px] font-bold tracking-widest">ESTADO</span>
          </button>
          <button onClick={toggleTheme} className="flex flex-col items-center opacity-40 hover:opacity-100 transition-opacity" title="Cambiar Apariencia">
            <span className="material-symbols-outlined text-[26px]">{ui.theme === 'light' ? 'dark_mode' : 'light_mode'}</span>
            <span className="text-[9px] font-bold uppercase">{ui.theme === 'light' ? 'Noche' : 'Día'}</span>
          </button>
        </div>
      </nav>
    </div>
  )
}

export default App

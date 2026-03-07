import { useState, useEffect, useMemo } from 'react'
import { validarCorreo } from './utils/reportUtils'

const CATEGORIAS = [
  { id: 'todos', nombre: 'Todos' },
  { id: 'pelotas', nombre: 'Pelotas' },
  { id: 'escolar', nombre: 'Escolar' },
  { id: 'representadas', nombre: 'Representadas' },
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
    
    return ui.allProducts.filter(p => {
      const skuStr = String(p.sku || '').toLowerCase()
      const eanStr = String(p.ean || '').toLowerCase()
      const nombreStr = String(p.nombre || '').toLowerCase()
      const lineaStr = String(p.linea || '').toLowerCase()
      
      // Búsqueda inteligente en múltiples campos
      return skuStr.includes(term) || 
             skuStr.replace(/^0+/, '').includes(term) || // Por si buscan "2210" en vez de "02210"
             eanStr.includes(term) || 
             nombreStr.includes(term) || 
             lineaStr.includes(term)
    }).slice(0, 15)
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
      const response = await fetch(downloadPath);
      const contentType = response.headers.get('content-type');
      if (!response.ok || (contentType && contentType.includes('text/html'))) {
        throw new Error('Reporte no disponible actualmente.');
      }
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
    } catch (err) { alert(err.message); }
  }

  return (
    <div className="bg-background-light dark:bg-background-dark font-display text-slate-900 dark:text-slate-100 min-h-screen flex flex-col relative">
      
      {/* SEARCH OVERLAY - TOTALMENTE PLANO */}
      {ui.isSearching && (
        <div className="fixed inset-0 z-50 bg-white dark:bg-[#102022] flex flex-col animate-in fade-in duration-200">
          <div className="p-6 flex items-center gap-4 border-b border-slate-200 dark:border-slate-800">
            <span className="material-symbols-outlined text-primary text-3xl">search</span>
            <input 
              autoFocus className="flex-1 bg-transparent border-none text-2xl font-bold focus:ring-0 outline-none"
              placeholder="Buscar..." value={ui.searchTerm}
              onChange={(e) => setUi(prev => ({ ...prev, searchTerm: e.target.value }))}
            />
            <button onClick={() => setUi(prev => ({ ...prev, isSearching: false, searchTerm: '' }))} className="size-12 rounded-full bg-slate-100 dark:bg-[#1a2a2c] flex items-center justify-center">
              <span className="material-symbols-outlined">close</span>
            </button>
          </div>
          <div className="flex-1 overflow-y-auto p-6 space-y-4">
            {searchResults.map(p => (
              <div key={p.sku} className="p-5 bg-white dark:bg-[#1a2a2c] rounded-2xl border border-slate-100 dark:border-slate-800 flex items-center justify-between">
                <div className="flex-1 min-w-0 mr-4">
                  <div className="flex items-center gap-3 mb-1.5">
                    <span className="text-[13px] font-black text-primary px-2.5 py-1 bg-primary/10 rounded-lg border border-primary/20">{p.sku}</span>
                    {p.ean && <span className="text-[10px] font-bold text-slate-400 dark:text-slate-500 tracking-tight">EAN: {p.ean}</span>}
                  </div>
                  <p className="font-bold truncate text-sm text-slate-800 dark:text-slate-200">{p.nombre}</p>
                  <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest mt-1 opacity-70">{p.linea}</p>
                </div>
                <div className="text-right shrink-0">
                  <div className="flex items-center gap-2 justify-end">
                    <span className="text-2xl font-black text-slate-900 dark:text-white">{p.stock}</span>
                    <span className="text-lg">{p.alerta}</span>
                  </div>
                  <p className="text-[9px] uppercase font-black text-slate-400 tracking-tighter">{ui.metadata.almacen}</p>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* HEADER - TOTALMENTE PLANO */}
      <header className="flex items-center p-4 justify-between border-b sticky top-0 z-40 bg-background-light dark:bg-background-dark border-slate-200 dark:border-slate-800">
        <div className="flex size-11 shrink-0 items-center overflow-hidden rounded-2xl bg-primary/10 p-1">
          <img src="favicon.svg" alt="Logo" className="h-full w-full object-contain" />
        </div>
        <div className="flex flex-1 px-4 flex-col">
          <h2 className="text-xl leading-none font-bold">StockPulse</h2>
          <p className="text-[9px] font-bold text-slate-400 dark:text-slate-500 uppercase tracking-widest mt-1">Inteligencia • CIPSA</p>
        </div>
        <button onClick={() => setUi(prev => ({ ...prev, isSearching: true }))} className="size-11 flex items-center justify-center rounded-2xl bg-slate-100 dark:bg-[#1a2a2c]">
          <span className="material-symbols-outlined">search</span>
        </button>
      </header>

      {/* MAIN */}
      <main className="flex-1 overflow-y-auto pb-32 max-w-2xl mx-auto w-full px-6">
        {activeTab === 'pulso' ? (
          <div className="animate-in fade-in slide-in-from-left-4 duration-500">
            <div className="pt-10 pb-4">
              <h1 className="text-3xl font-bold text-slate-900 dark:text-white">Generar Reporte</h1>
              <p className="text-slate-500 mt-2">Almacén: {ui.metadata.almacen}</p>
            </div>

            <form onSubmit={(e) => {e.preventDefault(); setUi(prev => ({...prev, reporteGenerado: true}))}} className="py-6 space-y-8">
              <div className="space-y-6 bg-white dark:bg-[#1a2a2c] p-6 rounded-3xl border border-slate-100 dark:border-slate-800">
            <div className="flex flex-col gap-2">
              <label className="text-[13px] font-bold uppercase opacity-60">Solicitante</label>
              <div className="relative flex items-center">
                <span className="material-symbols-outlined absolute left-4 text-slate-400 z-10 pointer-events-none">person</span>
                <input className="input-field" placeholder="Nombre" value={form.nombre} onChange={(e) => setForm({ ...form, nombre: e.target.value })} />
              </div>
            </div>
            <div className="flex flex-col gap-2">
              <label className="text-[13px] font-bold uppercase opacity-60">Email Corporativo</label>
              <div className="relative flex items-center">
                <span className="material-symbols-outlined absolute left-4 text-slate-400 z-10 pointer-events-none">mail</span>
                <input className="input-field" placeholder="usuario@cipsa.com.pe" value={form.email} onChange={(e) => setForm({ ...form, email: e.target.value })} />
              </div>
            </div>
                <div className="flex flex-col gap-4">
                  <label className="text-[13px] font-bold uppercase opacity-60">Filtro</label>
                  <div className="flex flex-wrap gap-2.5">
                    {CATEGORIAS.map(cat => (
                      <button key={cat.id} type="button" className={`chip ${form.categoria === cat.id ? 'chip-active' : 'chip-inactive'}`} onClick={() => setForm({ ...form, categoria: cat.id })}>
                        {cat.nombre}
                      </button>
                    ))}
                  </div>
                </div>
              </div>
              <button type="submit" disabled={!ui.isValido} className="btn-primary">
                <span className="material-symbols-outlined">analytics</span> Generar Inventario
              </button>
            </form>

            {ui.reporteGenerado && (
              <div className="card-success animate-in fade-in slide-in-from-bottom-4 duration-500 mt-4">
                <h4 className="font-bold text-xl mb-4">¡Snapshot Listo!</h4>
                <div className="flex items-center p-4 bg-white dark:bg-[#102022] rounded-2xl mb-6 border border-slate-100 dark:border-slate-800">
                  <span className="material-symbols-outlined text-emerald-500 mr-3">description</span>
                  <p className="text-sm font-bold truncate">StockPulse_{form.categoria.toUpperCase()}.xlsx</p>
                </div>
                <button onClick={handleDescargar} className="w-full h-14 bg-slate-900 dark:bg-white text-white dark:text-slate-900 font-bold rounded-2xl">
                  DESCARGAR
                </button>
              </div>
            )}
          </div>
        ) : (
          <div className="animate-in fade-in slide-in-from-right-4 duration-500 pt-10">
            <h1 className="text-3xl font-bold mb-8 text-primary">Estado del Sistema</h1>
            <div className="grid grid-cols-2 gap-4">
              <div className="p-6 bg-white dark:bg-[#1a2a2c] rounded-3xl border border-slate-100 dark:border-slate-800">
                <p className="text-[10px] font-bold uppercase opacity-50 mb-2">Último Pulso</p>
                <p className="text-lg font-bold">{new Date(ui.metadata.lastUpdated).toLocaleTimeString()}</p>
              </div>
              <div className="p-6 bg-white dark:bg-[#1a2a2c] rounded-3xl border border-slate-100 dark:border-slate-800">
                <p className="text-[10px] font-bold uppercase opacity-50 mb-2">Estado Bot</p>
                <div className="flex items-center gap-2 text-emerald-500">
                  <span className="relative flex h-3 w-3">
                    <span className="animate-ping absolute h-full w-full rounded-full bg-emerald-400 opacity-75"></span>
                    <span className="relative h-3 w-3 bg-emerald-500 rounded-full"></span>
                  </span>
                  <p className="font-bold">OK</p>
                </div>
              </div>
              <div className="p-6 bg-white dark:bg-[#1a2a2c] rounded-3xl border border-slate-100 dark:border-slate-800 col-span-2">
                <div className="flex justify-between items-center">
                  <div className="text-center flex-1">
                    <p className="text-2xl font-bold">{ui.metadata.totalProducts}</p>
                    <p className="text-[10px] opacity-50 uppercase">Items</p>
                  </div>
                  <div className="text-center flex-1 text-red-500">
                    <p className="text-2xl font-bold">{ui.metadata.sinStock}</p>
                    <p className="text-[10px] opacity-50 uppercase">Agotados</p>
                  </div>
                  <div className="text-center flex-1 text-amber-500">
                    <p className="text-2xl font-bold">{ui.metadata.bajoStock}</p>
                    <p className="text-[10px] opacity-50 uppercase">Críticos</p>
                  </div>
                </div>
              </div>
            </div>
          </div>
        )}
      </main>

      {/* NAV - TOTALMENTE PLANO */}
      <nav className="fixed bottom-0 left-0 right-0 bg-background-light dark:bg-background-dark border-t border-slate-200 dark:border-slate-800 px-6 pb-10 pt-4 z-40">
        <div className="max-w-md mx-auto flex justify-between px-4">
          <button onClick={() => setActiveTab('pulso')} className={`flex flex-col items-center gap-1.5 transition-all ${activeTab === 'pulso' ? 'text-primary' : 'opacity-40'}`}>
            <span className="material-symbols-outlined text-[26px]">analytics</span>
            <span className="text-[9px] font-bold">PULSO</span>
          </button>
          <button onClick={() => setActiveTab('estado')} className={`flex flex-col items-center gap-1.5 transition-all ${activeTab === 'estado' ? 'text-primary' : 'opacity-40'}`}>
            <span className="material-symbols-outlined text-[26px]">monitor_heart</span>
            <span className="text-[9px] font-bold">ESTADO</span>
          </button>
          <button onClick={toggleTheme} className="flex flex-col items-center opacity-40">
            <span className="material-symbols-outlined text-[26px]">{ui.theme === 'light' ? 'dark_mode' : 'light_mode'}</span>
            <span className="text-[9px] font-bold uppercase">{ui.theme === 'light' ? 'Noche' : 'Día'}</span>
          </button>
        </div>
      </nav>
    </div>
  )
}

export default App

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
  const [isRefreshing, setIsRefreshing] = useState(false)
  const [ui, setUi] = useState({ 
    isValido: false, reporteGenerado: false, theme: 'light',
    isSearching: false, searchTerm: '', allProducts: [], 
    metadata: { lastUpdated: '', totalProducts: 0, almacen: 'Cipsa', sinStock: 0, bajoStock: 0, status: '...' },
    isStale: false, emailError: ''
  })

  const loadProducts = async (forceRefresh = false) => {
    setIsRefreshing(true)
    try {
      const baseUrl = import.meta.env.BASE_URL;
      
      // Si forceRefresh=true, forzar descarga fresca sin caché del navegador
      // El servidor estático (GitHub Pages) es la fuente de la verdad
      const cacheBuster = forceRefresh ? `?t=${Date.now()}&v=${Math.random()}` : `?t=${Date.now()}`
      const res = await fetch(`${baseUrl}productos_con_stock.json${cacheBuster}`, {
        cache: 'no-store'
      })
      if (res.ok) {
        const data = await res.json()
        // Verificar si la data tiene más de 1 hora de antigüedad
        const lastUpdated = data.metadata?.lastUpdated ? new Date(data.metadata.lastUpdated) : null
        const now = new Date()
        const oneHour = 60 * 60 * 1000
        const timeDiff = lastUpdated ? (now.getTime() - lastUpdated.getTime()) : null
        const isStale = timeDiff ? timeDiff > oneHour : false
        
        setUi(prev => ({ 
          ...prev, 
          allProducts: data.productos || [], 
          metadata: data.metadata || prev.metadata,
          isStale: isStale
        }))
        
        // Log para debug
        if (timeDiff !== null) {
          console.log(`Data actualizada hace ${Math.round(timeDiff/60000)} minutos. isStale: ${isStale}`);
        }
      }
    } catch (err) { console.error('Error cargando data:', err) }
    setIsRefreshing(false)
  }

  useEffect(() => {
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

  // Convertir texto a formato nombre propio (primera letra mayúscula)
  const toProperCase = (text) => {
    return text.toLowerCase().replace(/\b\w/g, (char) => char.toUpperCase())
  }

  useEffect(() => {
    const emailValido = validarCorreo(form.email)
    setUi(prev => ({ 
      ...prev, 
      isValido: emailValido.valido && form.nombre.length > 2,
      emailError: form.email.length > 0 && !emailValido.valido ? emailValido.mensaje : ''
    }))
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
      
      // Limpiar formulario después de descargar para evitar descargas excesivas
      setForm({ nombre: '', email: '', categoria: 'todos' })
      setUi(prev => ({ ...prev, reporteGenerado: false }))
    } catch (err) { alert(err.message); }
  }

  useEffect(() => {
    if (ui.reporteGenerado) {
      setTimeout(() => {
        window.scrollTo({ top: document.body.scrollHeight, behavior: 'smooth' });
      }, 100);
    }
  }, [ui.reporteGenerado])

  return (
    <div className="bg-background-light dark:bg-background-dark font-display text-slate-900 dark:text-slate-100 min-h-screen flex flex-col relative">
      
      {/* SEARCH OVERLAY - MODAL EN PC, FULLSCREEN EN MÓVIL */}
      {ui.isSearching && (
        <div className="fixed inset-0 z-50 bg-white/80 dark:bg-black/60 backdrop-blur-sm flex flex-col items-center justify-start md:pt-16 animate-in fade-in duration-200">
          <div className="w-full h-full md:h-[75vh] md:max-w-3xl bg-white dark:bg-[#102022] flex flex-col md:rounded-3xl md:shadow-2xl md:border md:border-slate-200 dark:md:border-slate-800 overflow-hidden">
            <div className="p-6 flex items-center gap-4 border-b border-slate-200 dark:border-slate-800">
              <span className="material-symbols-outlined text-primary text-3xl">search</span>
              <input 
                autoFocus className="flex-1 bg-transparent border-none text-2xl font-bold focus:ring-0 outline-none"
                placeholder="Buscar..." value={ui.searchTerm}
                onChange={(e) => setUi(prev => ({ ...prev, searchTerm: e.target.value }))}
              />
              <button onClick={() => setUi(prev => ({ ...prev, isSearching: false, searchTerm: '' }))} className="size-12 rounded-full bg-slate-100 dark:bg-[#1a2a2c] flex items-center justify-center hover:bg-slate-200 dark:hover:bg-slate-800 transition-colors">
                <span className="material-symbols-outlined">close</span>
              </button>
            </div>
            <div className="flex-1 overflow-y-auto p-6 space-y-4">
              {ui.searchTerm && ui.searchTerm.length >= 2 && (
                <div className="flex items-center justify-between text-xs text-slate-400 dark:text-slate-500 mb-2 px-1">
                  <span>{searchResults.length} resultado{searchResults.length !== 1 ? 's' : ''}</span>
                  <span className={`flex items-center gap-1 ${ui.isStale ? 'text-red-500' : ''}`}>
                    <span className="material-symbols-outlined text-[14px]">schedule</span>
                    Actualizado: {ui.metadata.lastUpdated ? new Date(ui.metadata.lastUpdated).toLocaleString('es-PE', { day: '2-digit', month: '2-digit', year: '2-digit', hour: '2-digit', minute: '2-digit' }) : 'N/A'}
                    {ui.isStale && <span className="text-red-500 font-bold ml-1">⚠</span>}
                  </span>
                </div>
              )}
              {searchResults.map(p => (
                <div key={p.sku} className="p-5 bg-white dark:bg-[#1a2a2c] rounded-2xl border border-slate-100 dark:border-slate-800 flex items-center justify-between hover:border-primary/50 transition-colors group">
                  <div className="flex-1 min-w-0 mr-4">
                    <div className="flex items-center gap-3 mb-1.5">
                      <span className="text-[13px] font-black text-primary px-2.5 py-1 bg-primary/10 rounded-lg border border-primary/20 group-hover:bg-primary group-hover:text-slate-900 transition-colors">{p.sku}</span>
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
              {ui.searchTerm && ui.searchTerm.length >= 2 && searchResults.length === 0 && (
                <div className="py-12 text-center">
                  <span className="material-symbols-outlined text-5xl opacity-20 mb-4">search_off</span>
                  <p className="text-slate-400 font-bold">No se encontraron productos</p>
                </div>
              )}
            </div>
          </div>
          {/* Clic fuera para cerrar en PC */}
          <div className="hidden md:block absolute inset-0 -z-10" onClick={() => setUi(prev => ({ ...prev, isSearching: false, searchTerm: '' }))}></div>
        </div>
      )}

      {/* HEADER - OPTIMIZADO PARA PC */}
      <header className="border-b sticky top-0 z-40 bg-background-light/80 dark:bg-background-dark/80 backdrop-blur-md border-slate-200 dark:border-slate-800">
        <div className="max-w-4xl mx-auto flex items-center p-4 justify-between w-full">
          <div className="flex size-11 shrink-0 items-center overflow-hidden rounded-2xl bg-primary/10 p-1">
            <img src="favicon.svg" alt="Logo" className="h-full w-full object-contain" />
          </div>
          <div className="flex flex-1 px-4 flex-col">
            <h2 className="text-xl leading-none font-bold">StockPulse</h2>
            <p className="text-[9px] font-bold text-slate-400 dark:text-slate-500 uppercase tracking-widest mt-1">Inteligencia • CIPSA</p>
          </div>
          <div className="flex items-center gap-3">
            <button onClick={() => setUi(prev => ({ ...prev, isSearching: true }))} className="size-11 flex items-center justify-center rounded-2xl bg-slate-100 dark:bg-[#1a2a2c] hover:bg-primary hover:text-slate-900 transition-all">
              <span className="material-symbols-outlined">search</span>
            </button>
          </div>
        </div>
      </header>

      {/* MAIN */}
      <main className="flex-1 overflow-y-auto pb-32 max-w-4xl mx-auto w-full px-6">
        {activeTab === 'pulso' ? (
          <div className="animate-in fade-in slide-in-from-left-4 duration-500">
            <div className="pt-10 pb-4">
              <h1 className="text-3xl font-bold text-slate-900 dark:text-white">Generar Reporte</h1>
              <p className="text-slate-500 mt-2">Almacén: {ui.metadata.almacen}</p>
              {ui.isStale && (
                <div className="mt-3 p-3 bg-red-50 dark:bg-red-900/20 border border-red-200 dark:border-red-800 rounded-xl flex items-center gap-2">
                  <span className="material-symbols-outlined text-red-500">warning</span>
                  <p className="text-sm text-red-600 dark:text-red-400 font-bold">Data desactualizada. Presiona actualizar.</p>
                  <button 
                    onClick={() => loadProducts(true)} 
                    disabled={isRefreshing}
                    className="ml-auto flex items-center gap-1 px-3 py-1.5 bg-red-500 text-white text-xs font-bold rounded-lg hover:bg-red-600 transition-colors disabled:opacity-50"
                  >
                    <span className={`material-symbols-outlined text-[16px] ${isRefreshing ? 'animate-spin' : ''}`}>refresh</span>
                    {isRefreshing ? 'Actualizando...' : 'Actualizar'}
                  </button>
                </div>
              )}
            </div>

            <form onSubmit={(e) => {e.preventDefault(); setUi(prev => ({...prev, reporteGenerado: true}))}} className="py-6 space-y-8">
              <div className="space-y-6 bg-white dark:bg-[#1a2a2c] p-6 rounded-3xl border border-slate-100 dark:border-slate-800">
            <div className="flex flex-col gap-2">
              <label className="text-[13px] font-bold uppercase opacity-60">Solicitante</label>
              <div className="relative flex items-center">
                <span className="material-symbols-outlined absolute left-4 text-slate-400 z-10 pointer-events-none">person</span>
                <input 
                  className="input-field" 
                  placeholder="Nombre" 
                  value={form.nombre} 
                  onChange={(e) => setForm({ ...form, nombre: e.target.value })}
                  onBlur={(e) => setForm({ ...form, nombre: toProperCase(e.target.value) })}
                />
              </div>
            </div>
            <div className="flex flex-col gap-2">
              <label className="text-[13px] font-bold uppercase opacity-60">Email Corporativo</label>
              <div className="relative flex items-center">
                <span className="material-symbols-outlined absolute left-4 text-slate-400 z-10 pointer-events-none">mail</span>
                <input 
                  className={`input-field ${ui.emailError ? 'border-red-500 focus:border-red-500' : ''}`} 
                  placeholder="usuario@cipsa.com.pe" 
                  value={form.email} 
                  onChange={(e) => setForm({ ...form, email: e.target.value })} 
                />
              </div>
              {ui.emailError && (
                <div className="flex items-center gap-1.5 text-red-500 text-xs font-bold animate-in fade-in duration-300">
                  <span className="material-symbols-outlined text-[16px]">info</span>
                  {ui.emailError}
                </div>
              )}
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
            <h1 className="text-3xl font-bold mb-2 text-primary">Estado del Sistema</h1>
            <p className="text-sm text-slate-500 mb-6">Monitoreo en tiempo real</p>
            <div className="p-4 bg-blue-50 dark:bg-blue-900/20 rounded-xl mb-6 border border-blue-200 dark:border-blue-800">
              <h3 className="font-bold text-sm text-blue-700 dark:text-blue-300 mb-2">ℹ️ Acerca de StockPulse</h3>
              <p className="text-xs text-blue-600 dark:text-blue-400">
                StockPulse es una aplicación que genera reportes de inventario de CIPSA. 
                Los datos se actualizan automáticamente cada hora mediante un bot de GitHub Actions. 
                También puedes forzar una actualización manual desde esta pantalla.
              </p>
            </div>
            <div className="flex items-center justify-between mb-6">
              <button 
                onClick={() => loadProducts(true)} 
                disabled={isRefreshing}
                className="flex items-center gap-2 px-4 py-2 bg-primary text-white text-xs font-bold rounded-xl hover:bg-primary/90 transition-colors disabled:opacity-50"
              >
                <span className={`material-symbols-outlined text-[18px] ${isRefreshing ? 'animate-spin' : ''}`}>refresh</span>
                {isRefreshing ? 'Actualizando...' : 'Forzar Actualización'}
              </button>
            </div>
            <div className="grid grid-cols-2 md:grid-cols-3 gap-4">
              <div className="p-6 bg-white dark:bg-[#1a2a2c] rounded-3xl border border-slate-100 dark:border-slate-800 relative group hover:border-primary/30 transition-colors">
                <p className="text-[10px] font-bold uppercase opacity-50 mb-2">Último Pulso</p>
                <p className={`text-lg font-bold ${ui.metadata.lastUpdated ? (ui.isStale ? 'text-red-500 animate-pulse' : 'text-emerald-500') : 'text-slate-400'}`}>
                  {ui.metadata.lastUpdated ? new Date(ui.metadata.lastUpdated).toLocaleString('es-PE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit', second: '2-digit' }) : 'Sin datos'}
                </p>
                {ui.metadata.lastUpdated && (
                  <p className="text-[10px] text-slate-400 mt-1">
                    UTC: {new Date(ui.metadata.lastUpdated).toISOString().slice(11, 19)}
                    <span className="mx-1">•</span>
                    {Math.round((new Date() - new Date(ui.metadata.lastUpdated)) / 60000)} min atrás
                  </p>
                )}
                {ui.isStale && (
                  <p className="text-[10px] text-red-500 font-bold mt-1">⚠ Data desactualizada</p>
                )}
                {/* Tooltip */}
                <div className="absolute top-2 right-2 opacity-0 group-hover:opacity-100 transition-opacity bg-slate-800 text-white text-xs p-2 rounded-lg w-48 z-10 pointer-events-none shadow-xl border border-slate-700">
                  Hora de la última actualización del stock. Se considera desactualizado después de 1 hora.
                </div>
              </div>
              <div className="p-6 bg-white dark:bg-[#1a2a2c] rounded-3xl border border-slate-100 dark:border-slate-800 relative group hover:border-emerald-500/30 transition-colors">
                <p className="text-[10px] font-bold uppercase opacity-50 mb-2">Estado Bot</p>
                <div className="flex items-center gap-2 text-emerald-500">
                  <span className="relative flex h-3 w-3">
                    <span className="animate-ping absolute h-full w-full rounded-full bg-emerald-400 opacity-75"></span>
                    <span className="relative h-3 w-3 bg-emerald-500 rounded-full"></span>
                  </span>
                  <p className="font-bold">OK</p>
                </div>
                {/* Tooltip */}
                <div className="absolute top-2 right-2 opacity-0 group-hover:opacity-100 transition-opacity bg-slate-800 text-white text-xs p-2 rounded-lg w-48 z-10 pointer-events-none shadow-xl border border-slate-700">
                  Sistema operativo. El bot de GitHub Actions actualiza los datos automáticamente.
                </div>
              </div>
              
              <div className="p-6 bg-white dark:bg-[#1a2a2c] rounded-3xl border border-slate-100 dark:border-slate-800 col-span-2 md:col-span-1 flex flex-col justify-center relative group hover:border-primary/30 transition-colors">
                <div className="flex md:flex-col justify-between items-center md:gap-4">
                  <div className="text-center flex-1">
                    <p className="text-2xl font-bold">{ui.metadata.totalProducts}</p>
                    <p className="text-[10px] opacity-50 uppercase font-bold">Items</p>
                  </div>
                  <div className="text-center flex-1 text-red-500 border-x md:border-x-0 md:border-y border-slate-100 dark:border-slate-800 py-0 md:py-2 w-full">
                    <p className="text-2xl font-bold">{ui.metadata.sinStock}</p>
                    <p className="text-[10px] opacity-50 uppercase font-bold">Agotados</p>
                  </div>
                  <div className="text-center flex-1 text-amber-500">
                    <p className="text-2xl font-bold">{ui.metadata.bajoStock}</p>
                    <p className="text-[10px] opacity-50 uppercase font-bold">Críticos</p>
                  </div>
                </div>
                {/* Tooltip */}
                <div className="absolute top-2 right-2 opacity-0 group-hover:opacity-100 transition-opacity bg-slate-800 text-white text-xs p-2 rounded-lg w-48 z-10 pointer-events-none shadow-xl border border-slate-700">
                  <b>Items:</b> Total de productos<br/>
                  <b>Agotados:</b> Stock = 0<br/>
                  <b>Críticos:</b> Bajo el mínimo
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

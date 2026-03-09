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
  const [tick, setTick] = useState(0) // Para forzar re-render cada minuto y actualizar el "hace X min"
  const [ui, setUi] = useState({ 
    isValido: false, reporteGenerado: false, theme: 'light',
    isSearching: false, searchTerm: '', allProducts: [], 
    metadata: { lastUpdated: '', totalProducts: 0, almacen: 'Cipsa', sinStock: 0, bajoStock: 0, status: '...' },
    isStale: false, emailError: ''
  })

  // Calcular minutos transcurridos de forma reactiva
  const minutesAgo = useMemo(() => {
    if (!ui.metadata.lastUpdated) return null
    return Math.max(0, Math.round((new Date() - new Date(ui.metadata.lastUpdated)) / 60000))
  }, [ui.metadata.lastUpdated, tick])

  const loadProducts = async (forceRefresh = false) => {
    setIsRefreshing(true)
    try {
      const baseUrl = import.meta.env.BASE_URL;
      const cacheBuster = forceRefresh ? `?t=${Date.now()}&v=${Math.random()}` : `?t=${Date.now()}`
      const res = await fetch(`${baseUrl}productos_con_stock.json${cacheBuster}`, {
        cache: 'no-store'
      })
      if (res.ok) {
        const data = await res.json()
        const lastUpdated = data.metadata?.lastUpdated ? new Date(data.metadata.lastUpdated) : null
        const now = new Date()
        const stalenessThreshold = 75 * 60 * 1000
        const timeDiff = lastUpdated ? (now.getTime() - lastUpdated.getTime()) : null
        const isStale = timeDiff ? timeDiff > stalenessThreshold : false
        
        setUi(prev => ({ 
          ...prev, 
          allProducts: data.productos || [], 
          metadata: data.metadata || prev.metadata,
          isStale: isStale
        }))
        console.log(`[Sync] Data cargada. Antigüedad: ${timeDiff ? Math.round(timeDiff/60000) : '?'} min.`);
      }
    } catch (err) { console.error('Error cargando data:', err) }
    setIsRefreshing(false)
  }

  useEffect(() => {
    loadProducts()

    // 1. Polling silencioso cada 30 minutos
    const syncInterval = setInterval(() => {
      loadProducts(true)
    }, 30 * 60 * 1000)

    // 2. Timer para actualizar el contador de "hace X min" cada minuto
    const tickInterval = setInterval(() => {
      setTick(n => n + 1)
    }, 60000)

    const handleVisibilityChange = () => {
      if (document.visibilityState === 'visible') {
        loadProducts()
      }
    }
    document.addEventListener('visibilitychange', handleVisibilityChange)

    return () => {
      clearInterval(syncInterval)
      clearInterval(tickInterval)
      document.removeEventListener('visibilitychange', handleVisibilityChange)
    }
  }, [])

  const searchResults = useMemo(() => {
    const term = ui.searchTerm.trim().toLowerCase()
    if (!term || term.length < 2) return []
    
    return ui.allProducts.filter(p => {
      const skuStr = String(p.sku || '').toLowerCase()
      const eanStr = String(p.ean || '').toLowerCase()
      const nombreStr = String(p.nombre || '').toLowerCase()
      const lineaStr = String(p.linea || '').toLowerCase()
      
      return skuStr.includes(term) || 
             skuStr.replace(/^0+/, '').includes(term) || 
             eanStr.includes(term) || 
             nombreStr.includes(term) || 
             lineaStr.includes(term)
    })
    .sort((a, b) => Number(a.orden) - Number(b.orden))
    .slice(0, 15)
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
    <div className="bg-background-light dark:bg-background-dark font-display text-slate-900 dark:text-slate-100 min-h-screen flex flex-col relative transition-colors duration-300">
      
      {/* SEARCH OVERLAY */}
      {ui.isSearching && (
        <div className="fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-sm flex flex-col items-center justify-start md:pt-16 animate-in fade-in duration-200">
          <div className="w-full h-full md:h-[75vh] md:max-w-3xl bg-white dark:bg-slate-900 flex flex-col md:rounded-3xl md:shadow-2xl md:border md:border-slate-200 dark:md:border-slate-800 overflow-hidden">
            <div className="p-6 flex items-center gap-4 border-b border-slate-100 dark:border-slate-800">
              <span className="material-symbols-outlined text-brand-600 dark:text-brand-400 text-3xl">search</span>
              <input 
                autoFocus className="flex-1 bg-transparent border-none text-2xl font-bold focus:ring-0 outline-none"
                placeholder="Buscar..." value={ui.searchTerm}
                onChange={(e) => setUi(prev => ({ ...prev, searchTerm: e.target.value }))}
              />
              <button onClick={() => setUi(prev => ({ ...prev, isSearching: false, searchTerm: '' }))} className="size-12 rounded-full bg-slate-100 dark:bg-slate-800 flex items-center justify-center hover:bg-slate-200 dark:hover:bg-slate-700 transition-colors">
                <span className="material-symbols-outlined">close</span>
              </button>
            </div>
            <div className="flex-1 overflow-y-auto p-6 space-y-4">
              {ui.searchTerm && ui.searchTerm.length >= 2 && (
                <div className="flex items-center justify-between text-xs text-slate-500 dark:text-slate-400 mb-2 px-1">
                  <span className="font-bold">{searchResults.length} resultado{searchResults.length !== 1 ? 's' : ''}</span>
                  <span className={`flex items-center gap-1 ${ui.isStale ? 'text-red-500' : ''}`}>
                    <span className="material-symbols-outlined text-[14px]">schedule</span>
                    Actualizado: {ui.metadata.lastUpdated ? new Date(ui.metadata.lastUpdated).toLocaleString('es-PE', { day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit' }) : 'N/A'}
                  </span>
                </div>
              )}
              {searchResults.map(p => (
                <div key={p.sku} className="p-5 bg-white dark:bg-slate-800/50 rounded-2xl border border-slate-100 dark:border-slate-700/50 flex items-center justify-between hover:border-brand-500/50 transition-colors group">
                  <div className="flex-1 min-w-0 mr-4">
                    <div className="flex items-center gap-3 mb-1.5">
                      <span className="text-[13px] font-black text-brand-700 dark:text-brand-400 px-2.5 py-1 bg-brand-50 dark:bg-brand-400/10 rounded-lg border border-brand-100 dark:border-brand-400/20 group-hover:bg-brand-600 group-hover:text-white dark:group-hover:bg-brand-400 dark:group-hover:text-brand-950 transition-colors">{p.sku}</span>
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
                  <span className="material-symbols-outlined text-5xl opacity-20 mb-4 text-slate-400">search_off</span>
                  <p className="text-slate-400 font-bold text-lg">No se encontraron productos</p>
                  <p className="text-slate-400 text-sm">Intenta con otro SKU o nombre</p>
                </div>
              )}
            </div>
          </div>
          <div className="hidden md:block absolute inset-0 -z-10" onClick={() => setUi(prev => ({ ...prev, isSearching: false, searchTerm: '' }))}></div>
        </div>
      )}

      {/* HEADER */}
      <header className="border-b sticky top-0 z-40 bg-background-light/80 dark:bg-background-dark/80 backdrop-blur-md border-slate-200 dark:border-slate-800">
        <div className="max-w-4xl mx-auto flex items-center p-4 justify-between w-full">
          <div className="flex size-11 shrink-0 items-center overflow-hidden rounded-2xl bg-white dark:bg-slate-100 p-1.5 shadow-sm border border-slate-200">
            <img src="favicon.svg" alt="Logo" className="h-full w-full object-contain" />
          </div>
          <div className="flex flex-1 px-4 flex-col">
            <h2 className="text-xl leading-none font-black tracking-tight text-slate-900 dark:text-white">StockPulse</h2>
            <p className="text-[9px] font-bold text-brand-600 dark:text-brand-400 uppercase tracking-widest mt-1">Inteligencia • CIPSA</p>
          </div>
          <div className="flex items-center gap-3">
            <button onClick={() => setUi(prev => ({ ...prev, isSearching: true }))} className="size-11 flex items-center justify-center rounded-2xl bg-slate-100 dark:bg-slate-800 hover:bg-brand-100 dark:hover:bg-brand-900/30 text-slate-600 dark:text-slate-300 hover:text-brand-600 dark:hover:text-brand-400 transition-all">
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
              <h1 className="text-3xl font-black text-slate-900 dark:text-white tracking-tight">Generar Reporte</h1>
              <div className="flex items-center gap-2 mt-2">
                <span className={`size-2 rounded-full ${ui.isStale ? 'bg-amber-500' : 'bg-brand-500'} animate-pulse`}></span>
                <p className="text-slate-500 font-bold text-sm flex items-center gap-2">
                  <span>Almacén: {ui.metadata.almacen}</span>
                  {minutesAgo !== null && (
                    <>
                      <span className="text-slate-300 dark:text-slate-700">•</span>
                      <span className="text-slate-400 font-medium text-xs">Sincronizado hace {minutesAgo} min</span>
                    </>
                  )}
                </p>
              </div>
              
              {ui.isStale && (
                <div className="mt-6 p-4 bg-amber-50 dark:bg-amber-900/20 border border-amber-200 dark:border-amber-800 rounded-2xl flex items-center gap-4">
                  <span className="material-symbols-outlined text-amber-500 text-3xl">warning</span>
                  <div className="flex-1">
                    <p className="text-sm text-amber-800 dark:text-amber-400 font-bold">Data desactualizada</p>
                    <p className="text-xs text-amber-700/70 dark:text-amber-500/70">La información tiene más de 1 hora.</p>
                  </div>
                  <button 
                    onClick={() => loadProducts(true)} 
                    disabled={isRefreshing}
                    className="flex items-center gap-1.5 px-4 py-2 bg-amber-500 text-white text-xs font-black rounded-xl hover:bg-amber-600 transition-colors disabled:opacity-50"
                  >
                    <span className={`material-symbols-outlined text-[16px] ${isRefreshing ? 'animate-spin' : ''}`}>refresh</span>
                    {isRefreshing ? '...' : 'ACTUALIZAR'}
                  </button>
                </div>
              )}
            </div>

            <form onSubmit={(e) => {e.preventDefault(); setUi(prev => ({...prev, reporteGenerado: true}))}} className="py-6 space-y-8">
              <div className="space-y-6 bg-white dark:bg-slate-800/40 p-8 rounded-3xl border border-slate-100 dark:border-slate-800 shadow-soft">
                <div className="flex flex-col gap-2">
                  <label className="text-[11px] font-black uppercase tracking-widest text-slate-400">Solicitante</label>
                  <div className="relative flex items-center">
                    <span className="material-symbols-outlined absolute left-4 text-slate-400 z-10">person</span>
                    <input 
                      className="input-field" 
                      placeholder="Tu nombre completo" 
                      value={form.nombre} 
                      onChange={(e) => setForm({ ...form, nombre: e.target.value })}
                      onBlur={(e) => setForm({ ...form, nombre: toProperCase(e.target.value) })}
                    />
                  </div>
                </div>
                <div className="flex flex-col gap-2">
                  <label className="text-[11px] font-black uppercase tracking-widest text-slate-400">Email Corporativo</label>
                  <div className="relative flex items-center">
                    <span className="material-symbols-outlined absolute left-4 text-slate-400 z-10">mail</span>
                    <input 
                      className={`input-field ${ui.emailError ? 'border-red-500 focus:border-red-500 focus:ring-red-500/10' : ''}`} 
                      placeholder="usuario@cipsa.com.pe" 
                      value={form.email} 
                      onChange={(e) => setForm({ ...form, email: e.target.value })} 
                    />
                  </div>
                  {ui.emailError && (
                    <div className="flex items-center gap-1.5 text-red-500 text-[11px] font-bold mt-1 px-1">
                      <span className="material-symbols-outlined text-[14px]">error</span>
                      {ui.emailError}
                    </div>
                  )}
                </div>
                <div className="flex flex-col gap-4">
                  <label className="text-[11px] font-black uppercase tracking-widest text-slate-400">Categoría de Interés</label>
                  <div className="flex flex-wrap gap-2.5">
                    {CATEGORIAS.map(cat => (
                      <button key={cat.id} type="button" className={`chip ${form.categoria === cat.id ? 'chip-active' : 'chip-inactive'}`} onClick={() => setForm({ ...form, categoria: cat.id })}>
                        {cat.nombre}
                      </button>
                    ))}
                  </div>
                </div>
              </div>
              <button type="submit" disabled={!ui.isValido} className="btn-primary disabled:opacity-40 disabled:cursor-not-allowed">
                <span className="material-symbols-outlined">analytics</span> Generar Inventario Real-Time
              </button>
            </form>

            {/* MODAL DE DESCARGA */}
            {ui.reporteGenerado && (
              <div className="fixed inset-0 z-50 flex items-center justify-center p-6 bg-slate-900/60 backdrop-blur-md animate-in fade-in duration-300">
                <div className="bg-white dark:bg-slate-800 w-full max-w-sm rounded-[2.5rem] p-8 shadow-2xl border border-slate-100 dark:border-slate-700 animate-in zoom-in-95 duration-300 relative overflow-hidden">
                  <div className="absolute top-0 right-0 p-6">
                    <button onClick={() => setUi(prev => ({ ...prev, reporteGenerado: false }))} className="size-10 flex items-center justify-center rounded-full bg-slate-100 dark:bg-slate-700 text-slate-500 hover:bg-slate-200 transition-colors">
                      <span className="material-symbols-outlined text-xl">close</span>
                    </button>
                  </div>

                  <div className="flex flex-col items-center text-center">
                    <div className="size-20 bg-emerald-100 dark:bg-emerald-500/20 rounded-full flex items-center justify-center mb-6">
                      <span className="material-symbols-outlined text-4xl text-emerald-600 dark:text-emerald-400">verified</span>
                    </div>
                    
                    <h4 className="font-black text-2xl mb-2 text-slate-900 dark:text-white">¡Snapshot Listo!</h4>
                    <p className="text-slate-500 dark:text-slate-400 text-sm mb-8">El reporte de stock ha sido generado con éxito.</p>
                    
                    <div className="w-full flex items-center p-4 bg-slate-50 dark:bg-slate-900/50 rounded-2xl mb-8 border border-slate-100 dark:border-slate-700">
                      <span className="material-symbols-outlined text-brand-600 dark:text-brand-400 mr-3 text-2xl">description</span>
                      <p className="text-xs font-black truncate text-slate-700 dark:text-slate-300">StockPulse_{form.categoria.toUpperCase()}.xlsx</p>
                    </div>

                    <button onClick={handleDescargar} className="w-full h-16 bg-brand-600 dark:bg-brand-500 text-white dark:text-slate-950 font-black rounded-2xl shadow-xl shadow-brand-500/20 hover:scale-[1.02] active:scale-[0.98] transition-all">
                      DESCARGAR AHORA
                    </button>
                  </div>
                </div>
              </div>
            )}
          </div>
        ) : (
          <div className="animate-in fade-in slide-in-from-right-4 duration-500 pt-10">
            <h1 className="text-3xl font-black text-slate-900 dark:text-white tracking-tight mb-2">Estado del Sistema</h1>
            <p className="text-sm font-bold text-slate-500 mb-8 flex items-center gap-2">
              <span className="size-2 rounded-full bg-emerald-500 animate-pulse"></span>
              Monitoreo en tiempo real
            </p>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <div className="p-8 bg-white dark:bg-slate-800/40 rounded-3xl border border-slate-100 dark:border-slate-800 shadow-soft relative overflow-hidden group">
                <p className="text-[11px] font-black uppercase tracking-widest text-slate-400 mb-4">Última Sincronización</p>
                <p className={`text-2xl font-black ${ui.metadata.lastUpdated ? (ui.isStale ? 'text-amber-500' : 'text-brand-600 dark:text-brand-400') : 'text-slate-300'}`}>
                  {ui.metadata.lastUpdated ? new Date(ui.metadata.lastUpdated).toLocaleTimeString('es-PE', { hour: '2-digit', minute: '2-digit', second: '2-digit' }) : '--:--:--'}
                </p>
                <p className="text-xs font-bold text-slate-400 mt-1">
                  {ui.metadata.lastUpdated ? new Date(ui.metadata.lastUpdated).toLocaleDateString('es-PE', { day: '2-digit', month: 'long', year: 'numeric' }) : 'Esperando datos...'}
                </p>
                {ui.isStale && (
                  <div className="mt-4 inline-flex items-center gap-1.5 px-3 py-1 bg-amber-100 dark:bg-amber-900/30 text-amber-600 dark:text-amber-400 text-[10px] font-black rounded-full">
                    <span className="material-symbols-outlined text-[14px]">history</span>
                    DATA ANTIGUA
                  </div>
                )}
              </div>

              <div className="p-8 bg-white dark:bg-slate-800/40 rounded-3xl border border-slate-100 dark:border-slate-800 shadow-soft">
                <p className="text-[11px] font-black uppercase tracking-widest text-slate-400 mb-4">Salud del Bot</p>
                <div className="flex items-center gap-3">
                  <div className="relative flex h-5 w-5">
                    <span className="animate-ping absolute h-full w-full rounded-full bg-emerald-400 opacity-75"></span>
                    <span className="relative h-5 w-5 bg-emerald-500 rounded-full border-4 border-emerald-100 dark:border-emerald-900/50"></span>
                  </div>
                  <div>
                    <p className="font-black text-emerald-600 dark:text-emerald-400 text-xl">ACTIVO</p>
                    <p className="text-[10px] font-bold text-slate-400">GitHub Actions Online</p>
                  </div>
                </div>
              </div>
              
              <div className="p-8 bg-slate-900 dark:bg-brand-500 rounded-3xl col-span-1 md:col-span-2 shadow-2xl overflow-hidden relative">
                <div className="absolute top-0 right-0 w-64 h-64 bg-white/5 dark:bg-black/5 rounded-full -mr-32 -mt-32 blur-3xl"></div>
                <div className="relative z-10 grid grid-cols-3 gap-8">
                  <div className="text-center">
                    <p className="text-3xl font-black text-white dark:text-slate-950">{ui.metadata.totalProducts}</p>
                    <p className="text-[10px] text-white/50 dark:text-slate-950/50 uppercase font-black tracking-widest">Total SKU</p>
                  </div>
                  <div className="text-center border-x border-white/10 dark:border-black/10">
                    <p className="text-3xl font-black text-red-400 dark:text-red-900">{ui.metadata.sinStock}</p>
                    <p className="text-[10px] text-white/50 dark:text-slate-950/50 uppercase font-black tracking-widest">Agotados</p>
                  </div>
                  <div className="text-center">
                    <p className="text-3xl font-black text-amber-400 dark:text-amber-900">{ui.metadata.bajoStock}</p>
                    <p className="text-[10px] text-white/50 dark:text-slate-950/50 uppercase font-black tracking-widest">Críticos</p>
                  </div>
                </div>
              </div>
            </div>

            <button 
              onClick={() => loadProducts(true)} 
              disabled={isRefreshing}
              className="mt-8 w-full h-16 bg-white dark:bg-slate-800 border-2 border-slate-100 dark:border-slate-700 rounded-2xl flex items-center justify-center gap-3 hover:border-brand-500 transition-all group disabled:opacity-50"
            >
              <span className={`material-symbols-outlined text-2xl group-hover:rotate-180 transition-transform duration-500 ${isRefreshing ? 'animate-spin text-brand-500' : 'text-slate-400'}`}>refresh</span>
              <span className="font-black text-slate-600 dark:text-slate-300">FORZAR RE-SINCRONIZACIÓN</span>
            </button>
          </div>
        )}
      </main>

      {/* NAV */}
      <nav className="fixed bottom-0 left-0 right-0 bg-white/80 dark:bg-slate-900/80 backdrop-blur-xl border-t border-slate-100 dark:border-slate-800 px-6 pb-10 pt-4 z-40">
        <div className="max-w-md mx-auto flex justify-between px-4">
          <button onClick={() => setActiveTab('pulso')} className={`flex flex-col items-center gap-1.5 transition-all group ${activeTab === 'pulso' ? 'text-brand-600 dark:text-brand-400' : 'text-slate-400'}`}>
            <span className={`material-symbols-outlined text-[28px] ${activeTab === 'pulso' ? 'fill-1' : ''}`}>analytics</span>
            <span className="text-[10px] font-black tracking-tighter">PULSO</span>
            {activeTab === 'pulso' && <span className="w-1 h-1 rounded-full bg-brand-500 animate-pulse"></span>}
          </button>
          
          <button onClick={() => setActiveTab('estado')} className={`flex flex-col items-center gap-1.5 transition-all group ${activeTab === 'estado' ? 'text-brand-600 dark:text-brand-400' : 'text-slate-400'}`}>
            <span className={`material-symbols-outlined text-[28px] ${activeTab === 'estado' ? 'fill-1' : ''}`}>monitor_heart</span>
            <span className="text-[10px] font-black tracking-tighter">ESTADO</span>
            {activeTab === 'estado' && <span className="w-1 h-1 rounded-full bg-brand-500 animate-pulse"></span>}
          </button>

          <button onClick={toggleTheme} className="flex flex-col items-center gap-1.5 text-slate-400 hover:text-brand-500 transition-colors">
            <span className="material-symbols-outlined text-[28px]">{ui.theme === 'light' ? 'dark_mode' : 'light_mode'}</span>
            <span className="text-[10px] font-black tracking-tighter uppercase">{ui.theme === 'light' ? 'Noche' : 'Día'}</span>
          </button>
        </div>
      </nav>
    </div>
  )
}

export default App

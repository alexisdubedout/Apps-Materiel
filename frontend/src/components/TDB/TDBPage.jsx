import { useState, useEffect, useCallback } from 'react';
import { useNavigate } from 'react-router-dom';
import { RefreshCw, AlertCircle, Lock, Unlock, Minimize2, LayoutDashboard } from 'lucide-react';
import GridLayout from 'react-grid-layout';
import 'react-grid-layout/css/styles.css';
import 'react-resizable/css/styles.css';
import TDBWidgetTable from './TDBWidgetTable';
import TDBWidgetText from './TDBWidgetText';
import TDBWidgetNotes from './TDBWidgetNotes';
import { excelService } from '@/services/excelService';
import { TDB_CONFIG } from '@/tdbConfig';
import { useTheme } from '@/contexts/ThemeContext';

function SkeletonWidget() {
  return (
    <div className="h-full flex flex-col bg-card dark:bg-surface-800 rounded-lg border border-surface-200 dark:border-slate-700/40 overflow-hidden">
      <div className="h-1.5 bg-surface-200 dark:bg-surface-700 animate-skeleton" />
      <div className="px-3 py-2 border-b border-surface-100 dark:border-slate-700/30">
        <div className="skeleton-line w-32 h-3.5" />
      </div>
      <div className="flex-1 p-3 space-y-2">
        {[...Array(4)].map((_, i) => (
          <div key={i} className="skeleton-line-sm" style={{ width: `${60 + Math.random() * 35}%`, animationDelay: `${i * 150}ms` }} />
        ))}
      </div>
    </div>
  );
}

export default function TDBPage() {
  const { isDark } = useTheme();
  const navigate = useNavigate();
  const [excelData, setExcelData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [layout, setLayout] = useState([]);
  const [isLocked, setIsLocked] = useState(false);
  const [fullscreenWidget, setFullscreenWidget] = useState(null);
  const [refreshing, setRefreshing] = useState(false);
  const [windowWidth, setWindowWidth] = useState(window.innerWidth);

  useEffect(() => { loadData(); loadLayout(); }, []);

  useEffect(() => {
    const handleResize = () => setWindowWidth(window.innerWidth);
    window.addEventListener('resize', handleResize);
    return () => window.removeEventListener('resize', handleResize);
  }, []);

  const loadData = async () => {
    try { setLoading(true); setError(null); setExcelData((await excelService.fetchAndParse()).data); }
    catch (err) { setError(err.message); }
    finally { setLoading(false); }
  };

  const handleRefresh = async () => {
    setRefreshing(true);
    try { setExcelData((await excelService.fetchAndParse()).data); }
    catch (err) { setError(err.message); }
    finally { setRefreshing(false); }
  };

  const loadLayout = () => {
    const savedLayout = localStorage.getItem('tdb-layout');
    const savedLocked = localStorage.getItem('tdb-locked');
    if (savedLayout) setLayout(JSON.parse(savedLayout));
    if (savedLocked) setIsLocked(JSON.parse(savedLocked));
  };

  const saveLayout = (newLayout) => {
    if (!isLocked) { setLayout(newLayout); localStorage.setItem('tdb-layout', JSON.stringify(newLayout)); }
  };

  const toggleLock = () => {
    const n = !isLocked; setIsLocked(n); localStorage.setItem('tdb-locked', JSON.stringify(n));
  };

  const getLayout = useCallback(() => {
    return TDB_CONFIG.widgets.map((w, idx) => {
      const existing = layout.find(l => l.i === w.id);
      if (existing) return { ...existing, static: isLocked };
      return { i: w.id, x: (idx % 2) * 6, y: Math.floor(idx / 2) * 2, w: 6, h: 2, minW: 3, minH: 1, static: isLocked };
    });
  }, [layout, isLocked]);

  const prepareData = (widget) => {
    if (widget.type !== 'table') return null;
    let data = excelService.filterData(excelData, widget.filter);
    if (data && widget.columns) {
      data = data.map(row => {
        const r = {};
        widget.columns.forEach(col => { r[col] = row[col] || ''; });
        if (row['_redmine_url']) r['_redmine_url'] = row['_redmine_url'];
        return r;
      });
    }
    return data;
  };

  if (loading) {
    return (
      <div className="min-h-screen bg-surface-50 dark:bg-surface-950">
        {/* Header skeleton */}
        <header className="sticky top-0 z-40 bg-card/90 dark:bg-surface-900/90 backdrop-blur-lg border-b border-surface-200/80 dark:border-slate-700/30">
          <div className="flex items-center px-4 h-12 gap-2.5">
            <LayoutDashboard className="w-4 h-4 text-teal-600 dark:text-teal-400" />
            <div className="skeleton-line w-32 h-3.5" />
          </div>
        </header>
        <div className="p-4 grid grid-cols-2 gap-3">
          {[...Array(5)].map((_, i) => (
            <div key={i} className="animate-stagger-in" style={{ animationDelay: `${i * 80}ms` }}>
              <SkeletonWidget />
            </div>
          ))}
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="min-h-screen bg-surface-50 dark:bg-surface-950 flex items-center justify-center p-4">
        <div className="bg-card dark:bg-surface-800 rounded-lg shadow-xl p-8 max-w-md w-full border border-surface-200 dark:border-slate-700/40 animate-scale-in">
          <AlertCircle className="w-10 h-10 text-red-500 mx-auto mb-4" />
          <h2 className="text-lg font-display font-bold text-stone-800 dark:text-slate-100 mb-2 text-center">Erreur de chargement</h2>
          <p className="text-stone-500 mb-6 text-center text-sm">{error}</p>
          <div className="flex gap-3">
            <button onClick={() => navigate('/')} className="flex-1 px-4 py-2.5 rounded-lg bg-surface-100 dark:bg-surface-700 text-stone-600 dark:text-slate-300 hover:bg-surface-200 dark:hover:bg-surface-800 transition-colors duration-150 font-medium text-sm">Retour</button>
            <button onClick={loadData} className="flex-1 px-4 py-2.5 rounded-lg bg-teal-600 text-white hover:bg-teal-700 transition-colors duration-150 font-medium text-sm flex items-center justify-center gap-2"><RefreshCw className="w-4 h-4" /> Reessayer</button>
          </div>
        </div>
      </div>
    );
  }

  if (fullscreenWidget) {
    const widget = TDB_CONFIG.widgets.find(w => w.id === fullscreenWidget);
    const displayData = prepareData(widget);

    return (
      <div className="fixed inset-0 bg-surface-50 dark:bg-surface-950 z-50 flex flex-col">
        <div className="flex justify-between items-center px-5 py-3 bg-card dark:bg-surface-900 border-b border-surface-200 dark:border-slate-700/40">
          <div className="flex items-center gap-3">
            <div className={`w-1 h-5 rounded-full bg-gradient-to-b ${widget.color}`} />
            <h2 className="text-sm font-display font-bold text-stone-800 dark:text-slate-100">{widget.title}</h2>
          </div>
          <button onClick={() => setFullscreenWidget(null)} className="flex items-center gap-2 px-3 py-1.5 rounded-lg text-stone-500 hover:text-stone-700 dark:hover:text-slate-300 hover:bg-surface-100 dark:hover:bg-surface-700 transition-colors duration-150 text-sm font-medium">
            <Minimize2 className="w-4 h-4" /> Reduire
          </button>
        </div>
        <div className="flex-1 overflow-hidden p-4">
          {widget?.type === 'table' ? <TDBWidgetTable widget={widget} data={displayData} isDark={isDark} isFullscreen={true} />
          : widget?.type === 'notes' ? <TDBWidgetNotes widget={widget} isDark={isDark} />
          : <TDBWidgetText widget={widget} isDark={isDark} />}
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-surface-50 dark:bg-surface-950 transition-colors duration-300">
      <header className="sticky top-0 z-40 bg-card/90 dark:bg-surface-900/90 backdrop-blur-lg border-b border-surface-200/80 dark:border-slate-700/30">
        <div className="flex items-center justify-between px-4 h-12">
          <div className="flex items-center gap-2">
            <LayoutDashboard className="w-4 h-4 text-teal-600 dark:text-teal-400" />
            <h1 className="text-sm font-display font-bold text-stone-800 dark:text-slate-100">Tableau de bord MCO</h1>
          </div>

          <div className="flex items-center gap-1">
            <button onClick={handleRefresh} disabled={refreshing}
              className="flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg text-xs font-medium text-teal-700 dark:text-teal-400 hover:bg-teal-50 dark:hover:bg-teal-900/20 transition-colors duration-150 disabled:opacity-50">
              <RefreshCw className={`w-3.5 h-3.5 ${refreshing ? 'animate-spin' : ''}`} />
              <span className="hidden sm:inline">Rafraichir</span>
            </button>

            <button onClick={toggleLock}
              className={`flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg text-xs font-medium transition-colors duration-150 ${
                isLocked ? 'bg-amber-50 dark:bg-amber-900/20 text-amber-700 dark:text-amber-400' : 'text-stone-500 hover:bg-surface-100 dark:hover:bg-surface-700'
              }`}>
              {isLocked ? <Lock className="w-3.5 h-3.5" /> : <Unlock className="w-3.5 h-3.5" />}
              <span className="hidden sm:inline">{isLocked ? 'Verrouille' : 'Libre'}</span>
            </button>
          </div>
        </div>
      </header>

      <div className="p-4">
        <GridLayout
          className="layout"
          layout={getLayout()}
          cols={12}
          rowHeight={30}
          width={windowWidth - 32}
          onLayoutChange={saveLayout}
          draggableHandle=".drag-handle"
          compactType="vertical"
          preventCollision={false}
          margin={[12, 12]}
        >
          {TDB_CONFIG.widgets.map((widget, idx) => (
            <div key={widget.id} className={!isLocked ? "drag-handle cursor-move" : ""}>
              <div className="animate-stagger-in h-full" style={{ animationDelay: `${idx * 60}ms` }}>
                {widget.type === 'table' ? (
                  <TDBWidgetTable widget={widget} data={prepareData(widget)} isDark={isDark} onDoubleClick={() => setFullscreenWidget(widget.id)} isLocked={isLocked} />
                ) : widget.type === 'notes' ? (
                  <TDBWidgetNotes widget={widget} isDark={isDark} onDoubleClick={() => setFullscreenWidget(widget.id)} isLocked={isLocked} />
                ) : (
                  <TDBWidgetText widget={widget} isDark={isDark} onDoubleClick={() => setFullscreenWidget(widget.id)} isLocked={isLocked} />
                )}
              </div>
            </div>
          ))}
        </GridLayout>
      </div>
    </div>
  );
}

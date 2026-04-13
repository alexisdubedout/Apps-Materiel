import { useState } from 'react';
import { Lock, ChevronDown, ChevronUp, Filter, Search, X, Inbox } from 'lucide-react';

export default function TDBWidgetTable({ widget, data, isDark, onDoubleClick, isLocked, isFullscreen = false }) {
  const [searchTerm, setSearchTerm] = useState('');
  const [sortConfig, setSortConfig] = useState({ column: null, direction: null });
  const [filters, setFilters] = useState({});
  const [activeFilterColumn, setActiveFilterColumn] = useState(null);

  const columns = widget.columns || [];

  // Empty state
  if (!data || data.length === 0) {
    return (
      <div className={`h-full flex flex-col bg-card dark:bg-surface-800 rounded-lg border border-surface-200 dark:border-slate-700/40 overflow-hidden ${isFullscreen ? 'border-0 rounded-none' : ''}`}>
        <div className={`h-1.5 bg-gradient-to-r ${widget.color}`} />
        <div className="px-3 py-2 border-b border-surface-100 dark:border-slate-700/30">
          <h3 className="text-sm font-semibold text-stone-800 dark:text-slate-100">{widget.title}</h3>
        </div>
        <div className="flex-1 flex flex-col items-center justify-center py-8">
          <Inbox className="w-8 h-8 text-surface-300 dark:text-slate-600 mb-3" />
          <p className="text-sm font-medium text-stone-500 dark:text-slate-400">Aucune donnee</p>
          <p className="text-xs text-stone-400 dark:text-slate-500 mt-1">Les resultats apparaitront ici</p>
        </div>
      </div>
    );
  }

  const handleSort = (col) => {
    setSortConfig(prev => ({ column: col, direction: prev.column === col && prev.direction === 'asc' ? 'desc' : 'asc' }));
  };

  const getUniqueValues = (col) => [...new Set(data.map(r => r[col]).filter(v => v && v !== '-'))].sort();

  const toggleFilter = (col, val) => {
    setFilters(prev => {
      const cf = prev[col] || [];
      const nf = cf.includes(val) ? cf.filter(v => v !== val) : [...cf, val];
      if (nf.length === 0) { const { [col]: _, ...rest } = prev; return rest; }
      return { ...prev, [col]: nf };
    });
  };

  let filteredData = data.filter(row => {
    const matchSearch = Object.values(row).some(v => v?.toString().toLowerCase().includes(searchTerm.toLowerCase()));
    const matchFilters = Object.entries(filters).every(([c, vs]) => vs.length === 0 || vs.includes(row[c]));
    return matchSearch && matchFilters;
  });

  if (sortConfig.column) {
    filteredData.sort((a, b) => {
      const av = a[sortConfig.column] || '', bv = b[sortConfig.column] || '';
      const an = parseFloat(av.toString().replace(/[^\d.-]/g, '')), bn = parseFloat(bv.toString().replace(/[^\d.-]/g, ''));
      if (!isNaN(an) && !isNaN(bn)) return sortConfig.direction === 'asc' ? an - bn : bn - an;
      return sortConfig.direction === 'asc' ? av.toString().localeCompare(bv.toString()) : bv.toString().localeCompare(av.toString());
    });
  }

  const hasFilters = Object.keys(filters).length > 0;

  return (
    <div className={`h-full flex flex-col bg-card dark:bg-surface-800 rounded-lg overflow-hidden border border-surface-200 dark:border-slate-700/40 ${isFullscreen ? 'border-0 rounded-none' : ''}`}>
      {/* Color bar */}
      <div className={`h-1.5 bg-gradient-to-r ${widget.color} flex-shrink-0`} />

      {/* Header */}
      <div
        className={`flex items-center justify-between px-3 py-2 border-b border-surface-100 dark:border-slate-700/30 ${!isLocked && !isFullscreen ? 'cursor-pointer' : ''}`}
        onDoubleClick={onDoubleClick}
        title={!isLocked && !isFullscreen ? "Double-clic pour agrandir" : ""}
      >
        <div className="flex items-center gap-2 min-w-0">
          <h3 className="text-[13px] font-semibold text-stone-800 dark:text-slate-100 truncate">{widget.title}</h3>
          <span className="text-[11px] font-medium text-stone-500 tabular-nums flex-shrink-0">
            {filteredData.length}{hasFilters ? '*' : ''}
          </span>
        </div>
        <div className="flex items-center gap-1.5 flex-shrink-0">
          {hasFilters && (
            <button onClick={() => setFilters({})} className="text-[11px] text-red-500 hover:text-red-600 font-medium px-1.5 py-0.5 rounded-lg hover:bg-red-50 dark:hover:bg-red-900/20 transition-colors duration-150">
              Effacer
            </button>
          )}
          {isLocked && <Lock className="w-3 h-3 text-stone-300 dark:text-slate-600" />}
        </div>
      </div>

      {/* Search */}
      <div className="px-3 py-1.5 border-b border-surface-100 dark:border-slate-700/30">
        <div className="relative">
          <Search className="absolute left-2 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-stone-400" />
          <input type="text" placeholder="Rechercher..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)}
            className="w-full pl-7 pr-7 py-1 text-xs rounded-lg bg-surface-50 dark:bg-surface-900 border border-surface-200 dark:border-slate-700/50 text-stone-800 dark:text-slate-100 placeholder-stone-400 focus:ring-1 focus:ring-teal-500/30 focus:border-teal-500 outline-none transition-all" />
          {searchTerm && (
            <button onClick={() => setSearchTerm('')} className="absolute right-2 top-1/2 -translate-y-1/2 text-stone-400 hover:text-stone-600 transition-colors duration-150"><X className="w-3 h-3" /></button>
          )}
        </div>
      </div>

      {/* Table */}
      <div className="flex-1 overflow-auto">
        <table className="w-full">
          <thead className="sticky top-0 z-10">
            <tr className="bg-surface-50/90 dark:bg-surface-900/90 backdrop-blur-sm">
              {columns.map((col, idx) => (
                <th key={idx} className="px-3 py-1.5 text-left text-[11px] font-medium text-stone-500 dark:text-slate-400 border-b border-surface-200 dark:border-slate-700/50 relative">
                  <div className="flex items-center gap-1 group">
                    <span className="truncate">{col}</span>
                    <div className="flex items-center gap-0.5 opacity-0 group-hover:opacity-100 transition-opacity duration-150">
                      <button onClick={() => handleSort(col)} className="p-0.5 hover:bg-surface-200 dark:hover:bg-surface-700 rounded-lg transition-colors duration-150" title="Trier">
                        {sortConfig.column === col ? (sortConfig.direction === 'asc' ? <ChevronUp className="w-3 h-3 text-teal-600" /> : <ChevronDown className="w-3 h-3 text-teal-600" />) : <ChevronDown className="w-3 h-3 text-stone-400" />}
                      </button>
                      <button onClick={(e) => { e.stopPropagation(); setActiveFilterColumn(activeFilterColumn === col ? null : col); }}
                        className={`p-0.5 hover:bg-surface-200 dark:hover:bg-surface-700 rounded-lg transition-colors duration-150 ${filters[col]?.length > 0 ? 'text-teal-600 !opacity-100' : 'text-stone-400'}`} title="Filtrer">
                        <Filter className="w-3 h-3" />
                      </button>
                    </div>
                    {(sortConfig.column === col || filters[col]?.length > 0) && <div className="w-1 h-1 rounded-full bg-teal-500 flex-shrink-0" />}
                  </div>

                  {activeFilterColumn === col && (
                    <div className="absolute top-full left-0 mt-1 bg-card dark:bg-surface-800 border border-surface-200 dark:border-slate-700/50 rounded-lg shadow-xl z-50 min-w-[180px] max-h-[260px] overflow-y-auto animate-scale-in"
                      onClick={(e) => e.stopPropagation()}>
                      <div className="p-1.5 space-y-0.5">
                        {getUniqueValues(col).map((v, i) => (
                          <label key={i} className="flex items-center gap-2 px-2 py-1.5 hover:bg-surface-50 dark:hover:bg-surface-700/50 rounded-lg cursor-pointer text-xs transition-colors duration-150">
                            <input type="checkbox" checked={filters[col]?.includes(v) || false} onChange={() => toggleFilter(col, v)}
                              className="rounded border-stone-300 dark:border-slate-600 text-teal-500 focus:ring-teal-500 focus:ring-offset-0 w-3.5 h-3.5" />
                            <span className="text-stone-700 dark:text-slate-300 truncate">{v}</span>
                          </label>
                        ))}
                      </div>
                      {filters[col]?.length > 0 && (
                        <div className="border-t border-surface-100 dark:border-slate-700/30 p-1.5">
                          <button onClick={() => setFilters(prev => { const { [col]: _, ...rest } = prev; return rest; })}
                            className="w-full text-xs text-center py-1 text-red-500 hover:bg-red-50 dark:hover:bg-red-900/20 rounded-lg font-medium transition-colors duration-150">Effacer</button>
                        </div>
                      )}
                    </div>
                  )}
                </th>
              ))}
            </tr>
          </thead>
          <tbody className="divide-y divide-surface-100 dark:divide-slate-700/20">
            {filteredData.length === 0 ? (
              <tr>
                <td colSpan={columns.length} className="px-3 py-8 text-center">
                  <Search className="w-6 h-6 text-surface-300 dark:text-slate-600 mx-auto mb-2" />
                  <p className="text-xs text-stone-500 dark:text-slate-400">Aucun resultat</p>
                </td>
              </tr>
            ) : filteredData.map((row, i) => (
              <tr key={i} className="hover:bg-surface-50/50 dark:hover:bg-surface-700/20 transition-colors duration-150">
                {columns.map((col, j) => (
                  <td key={j} className="px-3 py-1.5 text-sm text-stone-700 dark:text-slate-300">
                    {col === 'ID' && row['_redmine_url'] ? (
                      <a href={row['_redmine_url']} target="_blank" rel="noopener noreferrer" className="text-teal-700 dark:text-teal-400 hover:underline font-medium tabular-nums">{row[col] || '-'}</a>
                    ) : (
                      <span className={col === 'ID' || col === 'N° devis' || col === 'ticketid' || col === 'N° LandesK' ? 'font-medium tabular-nums' : ''}>{row[col] || '-'}</span>
                    )}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

import { useState } from 'react';
import { Search, X, User, Info } from 'lucide-react';
import { useToast } from '@/components/ui/ToastProvider';
import { AlertDialog, AlertDialogContent, AlertDialogHeader, AlertDialogTitle, AlertDialogDescription, AlertDialogFooter, AlertDialogAction, AlertDialogCancel } from '@/components/ui/alert-dialog';
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogFooter } from '@/components/ui/dialog';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { useMappings, useUpdateMapping, useDeleteMapping, useMappingValues, useMappingHistory } from '@/hooks/useQueries';

export default function MappingPage() {
  const [searchTerm, setSearchTerm] = useState('');
  const [editingCell, setEditingCell] = useState(null);
  const [editValue, setEditValue] = useState('');
  const [showSuggestions, setShowSuggestions] = useState(false);
  const [suggestions, setSuggestions] = useState([]);
  const [showHistory, setShowHistory] = useState(false);
  const [showAddPrompt, setShowAddPrompt] = useState(false);
  const [newCode, setNewCode] = useState('');
  const [confirmDelete, setConfirmDelete] = useState(null);
  const toast = useToast();

  const { data: mappings = {}, isLoading } = useMappings();
  const updateMapping = useUpdateMapping();
  const deleteMapping = useDeleteMapping();
  const { data: types = [] } = useMappingValues('type');
  const { data: typeEmplacements = [] } = useMappingValues('typeEmplacement');
  const { data: details = [] } = useMappingValues('detail');
  const { data: history = [], isLoading: loadingHistory, refetch: refetchHistory } = useMappingHistory();

  const uniqueValues = { types, typeEmplacements, details };

  const handleDoubleClick = (code, field, value) => {
    setEditingCell({ code, field });
    setEditValue(value || '');
    const valMap = { type: types, typeEmplacement: typeEmplacements, detail: details };
    if (valMap[field]) { setSuggestions(valMap[field]); setShowSuggestions(true); }
  };

  const handleSave = async () => {
    if (!editingCell) return;
    const { code, field } = editingCell;
    const current = mappings[code] || {};
    const updated = { ...current, [field]: editValue };
    try {
      await updateMapping.mutateAsync({ code, ...updated, user: 'Utilisateur' });
      toast.success('Mapping mis à jour');
    } catch {
      toast.error('Erreur lors de la sauvegarde');
    }
    setEditingCell(null);
    setShowSuggestions(false);
  };

  const handleAddNew = async (code) => {
    setShowAddPrompt(false);
    setNewCode('');
    if (mappings[code]) { toast.error('Cet emplacement existe déjà'); return; }
    try {
      await updateMapping.mutateAsync({ code, description: '', type: '', typeEmplacement: '', detail: '', user: 'Utilisateur' });
      toast.success('Emplacement ajouté');
    } catch {
      toast.error("Erreur lors de l'ajout");
    }
  };

  const handleDelete = async (code) => {
    setConfirmDelete(null);
    try {
      await deleteMapping.mutateAsync({ code, user: 'Utilisateur' });
      toast.success('Emplacement supprimé');
    } catch {
      toast.error('Erreur lors de la suppression');
    }
  };

  const filteredMappings = Object.entries(mappings).filter(([code, data]) => {
    const s = searchTerm.toLowerCase();
    return code.toLowerCase().includes(s) || (data.description || '').toLowerCase().includes(s) || (data.typeEmplacement || '').toLowerCase().includes(s) || (data.detail || '').toLowerCase().includes(s);
  });

  const handleKeyDown = (e) => { if (e.key === 'Enter') handleSave(); else if (e.key === 'Escape') { setEditingCell(null); setShowSuggestions(false); } };
  const selectSuggestion = (value) => { setEditValue(value); setShowSuggestions(false); };

  return (
    <>
      <Dialog open={showAddPrompt} onOpenChange={(open) => { setShowAddPrompt(open); if (!open) setNewCode(''); }}>
        <DialogContent className="sm:max-w-sm">
          <DialogHeader>
            <DialogTitle>Nouvel emplacement</DialogTitle>
          </DialogHeader>
          <form onSubmit={(e) => { e.preventDefault(); if (newCode.trim()) handleAddNew(newCode.trim()); }}>
            <Input value={newCode} onChange={(e) => setNewCode(e.target.value)} placeholder="Code de l'emplacement" autoFocus className="mb-4" />
            <DialogFooter>
              <Button type="button" variant="outline" onClick={() => setShowAddPrompt(false)}>Annuler</Button>
              <Button type="submit" disabled={!newCode.trim()}>Ajouter</Button>
            </DialogFooter>
          </form>
        </DialogContent>
      </Dialog>

      <AlertDialog open={!!confirmDelete} onOpenChange={(open) => { if (!open) setConfirmDelete(null); }}>
        <AlertDialogContent>
          <AlertDialogHeader>
            <AlertDialogTitle>Supprimer l'emplacement</AlertDialogTitle>
            <AlertDialogDescription>Voulez-vous vraiment supprimer l'emplacement {confirmDelete} ?</AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel>Annuler</AlertDialogCancel>
            <AlertDialogAction onClick={() => handleDelete(confirmDelete)}>Supprimer</AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>

      <div className="min-h-screen bg-surface-50 dark:bg-surface-950 transition-colors duration-300">
        <div className="max-w-6xl mx-auto px-6 py-8">
          <div className="mb-6">
            <h1 className="text-xl font-display font-bold text-stone-800 dark:text-slate-100">Gestion du Mapping</h1>
            <p className="text-sm text-stone-500 dark:text-slate-400 mt-1">{Object.keys(mappings).length} emplacements</p>
          </div>

          <div className="bg-card dark:bg-surface-800 rounded-lg border border-surface-200 dark:border-slate-700/50 shadow-sm overflow-hidden">
            <div className="flex border-b border-surface-200 dark:border-slate-700/50 px-6">
              {['Mapping', 'Historique'].map((tab, i) => {
                const active = i === 0 ? !showHistory : showHistory;
                return (
                  <button key={tab} onClick={() => { if (i === 0) { setShowHistory(false); } else { setShowHistory(true); refetchHistory(); } }}
                    className={`px-4 py-2.5 text-sm font-medium relative transition-colors duration-150 ${active ? 'text-teal-700 dark:text-teal-400' : 'text-stone-500 hover:text-stone-700 dark:hover:text-slate-300'}`}>
                    {tab}
                    {active && <div className="absolute bottom-0 left-0 right-0 h-0.5 bg-teal-500 rounded-full" />}
                  </button>
                );
              })}
            </div>

            {!showHistory ? (
              <>
                <div className="flex items-center gap-3 px-6 py-3 border-b border-surface-100 dark:border-slate-700/30">
                  <div className="relative flex-1">
                    <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-stone-400" />
                    <input type="text" placeholder="Rechercher..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)}
                      className="w-full pl-9 pr-4 py-2 rounded-lg bg-surface-50 dark:bg-surface-900 border border-surface-200 dark:border-slate-700/50 text-stone-800 dark:text-slate-100 text-sm placeholder-stone-400 focus:ring-2 focus:ring-teal-500/20 focus:border-teal-500 outline-none transition-all" />
                  </div>
                  <Button onClick={() => setShowAddPrompt(true)} size="sm">+ Nouveau</Button>
                </div>
                <div className="overflow-auto max-h-[calc(100vh-300px)]">
                  {isLoading ? (
                    <div className="p-6 space-y-3">
                      {[...Array(8)].map((_, i) => (
                        <div key={i} className="flex gap-4 animate-skeleton" style={{ animationDelay: `${i * 100}ms` }}>
                          <div className="skeleton-line w-24" /><div className="skeleton-line flex-1" /><div className="skeleton-line w-20" /><div className="skeleton-line w-28" /><div className="skeleton-line w-20" />
                        </div>
                      ))}
                    </div>
                  ) : (
                    <table className="w-full">
                      <thead className="sticky top-0 z-10 bg-surface-50 dark:bg-surface-900">
                        <tr>
                          {['Code', 'Description', 'Type', 'Type emplacement', 'Détail', ''].map((h, i) => (
                            <th key={i} className={`px-4 py-2.5 text-left text-xs font-medium text-stone-500 dark:text-slate-400 border-b border-surface-200 dark:border-slate-700/50 ${i === 5 ? 'w-12' : ''}`}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-surface-100 dark:divide-slate-700/30">
                        {filteredMappings.length === 0 ? (
                          <tr>
                            <td colSpan={6} className="px-4 py-16 text-center">
                              <Search className="w-8 h-8 text-surface-300 dark:text-slate-600 mx-auto mb-3" />
                              <p className="text-sm font-medium text-stone-500 dark:text-slate-400">Aucun résultat</p>
                              <p className="text-xs text-stone-400 dark:text-slate-500 mt-1">Essayez un autre terme de recherche</p>
                            </td>
                          </tr>
                        ) : filteredMappings.map(([code, data]) => (
                          <tr key={code} className="hover:bg-surface-50/50 dark:hover:bg-surface-700/30 transition-colors duration-150">
                            <td className="px-4 py-2.5 text-sm font-mono font-medium text-stone-800 dark:text-slate-100">{code}</td>
                            {['description', 'type', 'typeEmplacement', 'detail'].map(field => (
                              <td key={field} className="px-4 py-2.5 text-sm text-stone-600 dark:text-slate-300 cursor-pointer hover:bg-teal-50/50 dark:hover:bg-teal-900/10 relative transition-colors duration-150" onDoubleClick={() => handleDoubleClick(code, field, data[field])}>
                                {editingCell?.code === code && editingCell?.field === field ? (
                                  <div className="relative">
                                    <input type="text" value={editValue} onChange={(e) => { setEditValue(e.target.value); setSuggestions((uniqueValues[field === 'typeEmplacement' ? 'typeEmplacements' : field + 's'] || []).filter(s => s.toLowerCase().includes(e.target.value.toLowerCase()))); }}
                                      onKeyDown={handleKeyDown} onBlur={handleSave} autoFocus
                                      className="w-full px-2 py-1 rounded-lg border-2 border-teal-500 bg-white dark:bg-surface-900 text-stone-800 dark:text-white text-sm outline-none" />
                                    {showSuggestions && suggestions.length > 0 && (
                                      <div className="absolute top-full left-0 mt-1 w-full max-h-36 overflow-y-auto bg-card dark:bg-surface-800 border border-surface-200 dark:border-slate-700/50 rounded-lg shadow-xl z-20">
                                        {suggestions.map((s, idx) => (<div key={idx} onMouseDown={() => selectSuggestion(s)} className="px-3 py-1.5 hover:bg-teal-50 dark:hover:bg-teal-900/20 cursor-pointer text-sm text-stone-700 dark:text-slate-300 transition-colors duration-150">{s}</div>))}
                                      </div>
                                    )}
                                  </div>
                                ) : (<span>{data[field] || <span className="text-stone-300 dark:text-slate-600 italic text-xs">--</span>}</span>)}
                              </td>
                            ))}
                            <td className="px-4 py-2.5 text-center">
                              <button onClick={() => setConfirmDelete(code)} className="p-1 rounded-lg text-stone-300 hover:text-red-500 hover:bg-red-50 dark:hover:bg-red-900/20 transition-colors duration-150"><X className="w-3.5 h-3.5" /></button>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )}
                </div>
              </>
            ) : (
              <div className="overflow-auto max-h-[calc(100vh-300px)] p-4">
                {loadingHistory ? (
                  <div className="space-y-3">
                    {[...Array(6)].map((_, i) => (
                      <div key={i} className="p-3 rounded-lg bg-surface-50 dark:bg-surface-700/30 space-y-2 animate-skeleton" style={{ animationDelay: `${i * 100}ms` }}>
                        <div className="skeleton-line w-48" /><div className="skeleton-line-sm w-full" />
                      </div>
                    ))}
                  </div>
                ) : (
                  <div className="space-y-2">
                    {history.map(entry => (
                      <div key={entry.id} className="p-3 rounded-lg bg-surface-50 dark:bg-surface-700/30 border border-surface-100 dark:border-slate-700/30">
                        <div className="flex items-center justify-between mb-1.5">
                          <div className="flex items-center gap-2">
                            <span className={`px-2 py-0.5 rounded text-[11px] font-semibold ${entry.action === 'CREATE' ? 'bg-emerald-100 text-emerald-700 dark:bg-emerald-900/30 dark:text-emerald-400' : entry.action === 'UPDATE' ? 'bg-teal-100 text-teal-700 dark:bg-teal-900/30 dark:text-teal-400' : 'bg-red-100 text-red-700 dark:bg-red-900/30 dark:text-red-400'}`}>{entry.action}</span>
                            <span className="font-mono text-sm font-semibold text-stone-800 dark:text-slate-100">{entry.emplacement_code}</span>
                          </div>
                          <span className="text-[11px] text-stone-500 tabular-nums">{new Date(entry.created_at).toLocaleString('fr-FR')}</span>
                        </div>
                        {entry.field_changed && (<div className="text-sm text-stone-600 mb-1"><span className="font-medium">{entry.field_changed} :</span> <span className="line-through text-red-400">{entry.old_value || '(vide)'}</span> → <span className="text-emerald-600 dark:text-emerald-400">{entry.new_value || '(vide)'}</span></div>)}
                        <div className="flex items-center gap-1.5 text-[11px] text-stone-500"><User className="w-3 h-3" /><span>{entry.user_name || entry.user_login}</span></div>
                      </div>
                    ))}
                    {history.length === 0 && (
                      <div className="text-center py-16">
                        <Info className="w-8 h-8 text-surface-300 dark:text-slate-600 mx-auto mb-3" />
                        <p className="text-sm font-medium text-stone-500 dark:text-slate-400">Aucun historique</p>
                      </div>
                    )}
                  </div>
                )}
              </div>
            )}
          </div>
        </div>
      </div>
    </>
  );
}

import { useState } from 'react';
import { Edit2, Save, X, Loader2 } from 'lucide-react';
import { useNotes, useSaveNote } from '@/hooks/useQueries';

export default function TDBWidgetNotes({ widget, onDoubleClick }) {
  const [isEditing, setIsEditing] = useState(false);
  const [editValue, setEditValue] = useState('');

  const { data: note, isLoading } = useNotes();
  const saveNote = useSaveNote();

  const content = note?.content || '';
  const updatedBy = note?.updated_by || '';
  const updatedAt = note?.updated_at || '';

  const handleSave = async () => {
    try {
      await saveNote.mutateAsync(editValue);
      setIsEditing(false);
    } catch (e) {
      console.error('Erreur sauvegarde notes:', e);
    }
  };

  if (isLoading) {
    return (
      <div className="h-full flex flex-col bg-card dark:bg-surface-800 rounded-lg border border-surface-200 dark:border-slate-700/40 overflow-hidden">
        <div className={`h-1.5 bg-gradient-to-r ${widget.color} flex-shrink-0`} />
        <div className="px-3 py-2 border-b border-surface-100 dark:border-slate-700/30">
          <div className="skeleton-line w-16 h-3.5" />
        </div>
        <div className="flex-1 p-3 space-y-2">
          <div className="skeleton-line w-full" />
          <div className="skeleton-line w-4/5" style={{ animationDelay: '100ms' }} />
          <div className="skeleton-line w-3/5" style={{ animationDelay: '200ms' }} />
        </div>
      </div>
    );
  }

  return (
    <div className="h-full flex flex-col bg-card dark:bg-surface-800 rounded-lg overflow-hidden border border-surface-200 dark:border-slate-700/40">
      <div className={`h-1.5 bg-gradient-to-r ${widget.color} flex-shrink-0`} />

      <div className="flex items-center justify-between px-3 py-2 border-b border-surface-100 dark:border-slate-700/30" onDoubleClick={onDoubleClick}>
        <h3 className="text-[13px] font-semibold text-stone-800 dark:text-slate-100">{widget.title}</h3>
        <div className="flex items-center gap-1">
          {!isEditing ? (
            <button onClick={() => { setEditValue(content); setIsEditing(true); }}
              className="p-1.5 rounded-lg text-stone-500 hover:text-stone-700 dark:hover:text-slate-300 hover:bg-surface-100 dark:hover:bg-surface-700 transition-colors duration-150" title="Modifier">
              <Edit2 className="w-3.5 h-3.5" />
            </button>
          ) : (
            <>
              <button onClick={handleSave} disabled={saveNote.isPending}
                className="p-1.5 rounded-lg text-emerald-600 hover:bg-emerald-50 dark:hover:bg-emerald-900/20 transition-colors duration-150 disabled:opacity-50" title="Sauvegarder">
                {saveNote.isPending ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <Save className="w-3.5 h-3.5" />}
              </button>
              <button onClick={() => { setEditValue(content); setIsEditing(false); }}
                className="p-1.5 rounded-lg text-stone-500 hover:text-red-500 hover:bg-red-50 dark:hover:bg-red-900/20 transition-colors duration-150" title="Annuler">
                <X className="w-3.5 h-3.5" />
              </button>
            </>
          )}
        </div>
      </div>

      <div className="flex-1 p-3 overflow-auto">
        {isEditing ? (
          <textarea value={editValue} onChange={(e) => setEditValue(e.target.value)} autoFocus
            className="w-full h-full p-2 rounded-lg bg-amber-50/50 dark:bg-surface-900 border border-amber-200 dark:border-slate-700/50 text-stone-800 dark:text-slate-100 text-sm resize-none focus:ring-2 focus:ring-amber-500/20 focus:border-amber-500 outline-none transition-all"
            placeholder="Ecrivez vos notes ici..." />
        ) : (
          <div className="text-stone-600 dark:text-slate-300 text-sm whitespace-pre-wrap leading-relaxed">
            {content || <span className="text-stone-400 italic">Aucune note</span>}
          </div>
        )}
      </div>

      {updatedBy && !isEditing && (
        <div className="px-3 py-1.5 bg-surface-50 dark:bg-surface-900/50 border-t border-surface-100 dark:border-slate-700/30 text-[11px] text-stone-500">
          Modifié par <span className="font-medium text-stone-600 dark:text-slate-400">{updatedBy}</span>
          {updatedAt && <span> le {new Date(updatedAt).toLocaleString('fr-FR')}</span>}
        </div>
      )}
    </div>
  );
}

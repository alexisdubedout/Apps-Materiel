import { useState, useEffect } from 'react';
import { Edit2, Save, X } from 'lucide-react';

export default function TDBWidgetText({ widget, isDark, onDoubleClick, isLocked }) {
  const [text, setText] = useState('');
  const [isEditing, setIsEditing] = useState(false);

  useEffect(() => {
    const saved = localStorage.getItem(`tdb-note-${widget.id}`);
    if (saved) setText(saved);
  }, [widget.id]);

  const handleSave = () => {
    localStorage.setItem(`tdb-note-${widget.id}`, text);
    setIsEditing(false);
  };

  const handleCancel = () => {
    const saved = localStorage.getItem(`tdb-note-${widget.id}`) || '';
    setText(saved);
    setIsEditing(false);
  };

  return (
    <div className="h-full flex flex-col bg-card dark:bg-surface-800 rounded-lg overflow-hidden border border-surface-200 dark:border-slate-700/40">
      {/* Color bar */}
      <div className={`h-1.5 bg-gradient-to-r ${widget.color} flex-shrink-0`} />

      {/* Header */}
      <div className="flex items-center justify-between px-3 py-2 border-b border-surface-100 dark:border-slate-700/30" onDoubleClick={onDoubleClick}>
        <h3 className="text-[13px] font-semibold text-stone-800 dark:text-slate-100">{widget.title}</h3>
        <div className="flex items-center gap-1">
          {!isEditing ? (
            <button onClick={() => setIsEditing(true)}
              className="p-1.5 rounded-lg text-stone-500 hover:text-stone-700 dark:hover:text-slate-300 hover:bg-surface-100 dark:hover:bg-surface-700 transition-colors duration-150" title="Modifier">
              <Edit2 className="w-3.5 h-3.5" />
            </button>
          ) : (
            <>
              <button onClick={handleSave}
                className="p-1.5 rounded-lg text-emerald-600 hover:bg-emerald-50 dark:hover:bg-emerald-900/20 transition-colors duration-150" title="Sauvegarder">
                <Save className="w-3.5 h-3.5" />
              </button>
              <button onClick={handleCancel}
                className="p-1.5 rounded-lg text-stone-500 hover:text-red-500 hover:bg-red-50 dark:hover:bg-red-900/20 transition-colors duration-150" title="Annuler">
                <X className="w-3.5 h-3.5" />
              </button>
            </>
          )}
        </div>
      </div>

      {/* Content */}
      <div className="flex-1 overflow-auto p-3">
        {isEditing ? (
          <textarea value={text} onChange={(e) => setText(e.target.value)} autoFocus
            className="w-full h-full p-2 rounded-lg bg-amber-50/50 dark:bg-surface-900 border border-amber-200 dark:border-slate-700/50 text-stone-800 dark:text-slate-100 text-sm resize-none focus:ring-2 focus:ring-amber-500/20 focus:border-amber-500 outline-none transition-all"
            placeholder="Ecrivez vos notes ici..." />
        ) : (
          <div className="text-stone-600 dark:text-slate-300 text-sm whitespace-pre-wrap leading-relaxed">
            {text || <span className="text-stone-400 italic">Cliquez sur le crayon pour ajouter des notes</span>}
          </div>
        )}
      </div>

      {/* Footer */}
      <div className="px-3 py-1.5 bg-surface-50 dark:bg-surface-900/50 border-t border-surface-100 dark:border-slate-700/30 text-[11px] text-stone-500 tabular-nums">
        {text.length} caracteres
      </div>
    </div>
  );
}

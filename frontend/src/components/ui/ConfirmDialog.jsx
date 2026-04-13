import React from 'react';

export default function ConfirmDialog({ isOpen, title, message, confirmLabel = 'Confirmer', cancelLabel = 'Annuler', variant = 'danger', onConfirm, onCancel }) {
  if (!isOpen) return null;
  const btnColors = variant === 'danger'
    ? 'bg-red-600 hover:bg-red-700 text-white'
    : 'bg-teal-600 hover:bg-teal-700 text-white';
  return (
    <div className="fixed inset-0 z-[90] flex items-center justify-center bg-black/40 backdrop-blur-sm animate-fade-in">
      <div className="w-full max-w-sm mx-4 bg-card dark:bg-surface-800 rounded-lg border border-surface-200 dark:border-slate-700/50 shadow-2xl animate-scale-in p-6">
        <h3 className="text-base font-display font-bold text-stone-800 dark:text-slate-100 mb-2">{title}</h3>
        <p className="text-sm text-stone-600 dark:text-slate-400 mb-6 leading-relaxed">{message}</p>
        <div className="flex gap-3 justify-end">
          <button onClick={onCancel} className="px-4 py-2 rounded-lg text-sm font-medium text-stone-600 dark:text-slate-400 hover:bg-surface-100 dark:hover:bg-surface-700 transition-colors duration-150">
            {cancelLabel}
          </button>
          <button onClick={onConfirm} className={`px-4 py-2 rounded-lg text-sm font-medium transition-colors duration-150 ${btnColors}`}>
            {confirmLabel}
          </button>
        </div>
      </div>
    </div>
  );
}

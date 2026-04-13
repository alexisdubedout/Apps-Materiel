import React, { useState, useEffect } from 'react';

export default function PromptDialog({ isOpen, title, placeholder, confirmLabel = 'Ajouter', onConfirm, onCancel }) {
  const [value, setValue] = useState('');
  useEffect(() => { if (isOpen) setValue(''); }, [isOpen]);
  if (!isOpen) return null;
  const handleSubmit = (e) => { e.preventDefault(); if (value.trim()) onConfirm(value.trim()); };
  return (
    <div className="fixed inset-0 z-[90] flex items-center justify-center bg-black/40 backdrop-blur-sm animate-fade-in">
      <form onSubmit={handleSubmit} className="w-full max-w-sm mx-4 bg-card dark:bg-surface-800 rounded-lg border border-surface-200 dark:border-slate-700/50 shadow-2xl animate-scale-in p-6">
        <h3 className="text-base font-display font-bold text-stone-800 dark:text-slate-100 mb-4">{title}</h3>
        <input type="text" value={value} onChange={e => setValue(e.target.value)} placeholder={placeholder} autoFocus
          className="w-full px-3.5 py-2.5 rounded-lg bg-surface-50 dark:bg-surface-900 border border-surface-200 dark:border-slate-700/50 text-stone-800 dark:text-slate-100 text-sm placeholder-stone-400 focus:border-teal-500 focus:ring-2 focus:ring-teal-500/20 outline-none transition-all mb-6" />
        <div className="flex gap-3 justify-end">
          <button type="button" onClick={onCancel} className="px-4 py-2 rounded-lg text-sm font-medium text-stone-600 dark:text-slate-400 hover:bg-surface-100 dark:hover:bg-surface-700 transition-colors duration-150">
            Annuler
          </button>
          <button type="submit" disabled={!value.trim()} className="px-4 py-2 rounded-lg text-sm font-medium bg-teal-600 hover:bg-teal-700 text-white transition-colors duration-150 disabled:opacity-50 disabled:cursor-not-allowed">
            {confirmLabel}
          </button>
        </div>
      </form>
    </div>
  );
}

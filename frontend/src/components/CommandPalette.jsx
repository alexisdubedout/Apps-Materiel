import React, { useEffect, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { Command } from 'cmdk';
import { Home, FileSpreadsheet, Database, TrendingUp, LayoutDashboard, Map, Sun, Moon, LogOut } from 'lucide-react';
import { useAuth } from '@/contexts/AuthContext';
import { useTheme } from '@/contexts/ThemeContext';

const COMMANDS = [
  { id: 'home', label: 'Accueil', icon: Home, path: '/', group: 'Navigation' },
  { id: 'stock', label: 'Suivi des Stocks', icon: FileSpreadsheet, path: '/app/stock-tracking', group: 'Navigation' },
  { id: 'tri', label: 'Tri Matériel', icon: Database, path: '/app/tri-materiel', group: 'Navigation' },
  { id: 'indicateurs', label: 'Indicateurs', icon: TrendingUp, path: '/app/data-merge', group: 'Navigation' },
  { id: 'dashboard', label: 'Tableau de bord', icon: LayoutDashboard, path: '/dashboard', group: 'Navigation' },
  { id: 'mapping', label: 'Mapping', icon: Map, path: '/mapping', group: 'Navigation', adminOnly: true },
];

export default function CommandPalette({ open, onOpenChange }) {
  const navigate = useNavigate();
  const { user, logout } = useAuth();
  const { isDark, toggleTheme } = useTheme();

  useEffect(() => {
    const down = (e) => {
      if (e.key === 'k' && (e.metaKey || e.ctrlKey)) {
        e.preventDefault();
        onOpenChange(!open);
      }
    };
    document.addEventListener('keydown', down);
    return () => document.removeEventListener('keydown', down);
  }, [open, onOpenChange]);

  const runCommand = (callback) => {
    onOpenChange(false);
    callback();
  };

  const filteredCommands = COMMANDS.filter(cmd => {
    if (cmd.adminOnly && user?.role === 'Client') return false;
    if (user?.role === 'Client' && cmd.group === 'Navigation' && cmd.path !== '/' && cmd.path !== '/dashboard') return false;
    return true;
  });

  if (!open) return null;

  return (
    <div className="fixed inset-0 z-[100]">
      <div className="fixed inset-0 bg-black/40 backdrop-blur-sm" onClick={() => onOpenChange(false)} />
      <div className="fixed top-[20%] left-1/2 -translate-x-1/2 w-full max-w-lg">
        <Command className="rounded-lg border border-surface-200 dark:border-slate-700/50 bg-card dark:bg-surface-800 shadow-2xl overflow-hidden">
          <Command.Input
            placeholder="Rechercher une commande..."
            className="w-full px-4 py-3 text-sm text-stone-800 dark:text-slate-100 bg-transparent border-b border-surface-200 dark:border-slate-700/50 placeholder-stone-400 outline-none"
            autoFocus
          />
          <Command.List className="max-h-[300px] overflow-y-auto p-2">
            <Command.Empty className="py-6 text-center text-sm text-stone-500 dark:text-slate-400">
              Aucun resultat.
            </Command.Empty>

            <Command.Group heading="Navigation" className="text-xs font-medium text-stone-400 dark:text-slate-500 px-2 py-1.5">
              {filteredCommands.map(cmd => {
                const Icon = cmd.icon;
                return (
                  <Command.Item
                    key={cmd.id}
                    value={cmd.label}
                    onSelect={() => runCommand(() => navigate(cmd.path))}
                    className="flex items-center gap-2.5 px-2 py-2 rounded-md text-sm text-stone-700 dark:text-slate-300 cursor-pointer data-[selected=true]:bg-surface-100 dark:data-[selected=true]:bg-surface-700"
                  >
                    <Icon className="w-4 h-4 text-stone-400 dark:text-slate-500" />
                    {cmd.label}
                  </Command.Item>
                );
              })}
            </Command.Group>

            <Command.Group heading="Actions" className="text-xs font-medium text-stone-400 dark:text-slate-500 px-2 py-1.5 mt-1">
              <Command.Item
                value="Changer le thème"
                onSelect={() => runCommand(toggleTheme)}
                className="flex items-center gap-2.5 px-2 py-2 rounded-md text-sm text-stone-700 dark:text-slate-300 cursor-pointer data-[selected=true]:bg-surface-100 dark:data-[selected=true]:bg-surface-700"
              >
                {isDark ? <Sun className="w-4 h-4 text-amber-400" /> : <Moon className="w-4 h-4 text-stone-400" />}
                {isDark ? 'Mode clair' : 'Mode sombre'}
              </Command.Item>
              <Command.Item
                value="Déconnexion"
                onSelect={() => runCommand(() => { logout(); navigate('/login'); })}
                className="flex items-center gap-2.5 px-2 py-2 rounded-md text-sm text-red-600 dark:text-red-400 cursor-pointer data-[selected=true]:bg-red-50 dark:data-[selected=true]:bg-red-900/20"
              >
                <LogOut className="w-4 h-4" />
                Déconnexion
              </Command.Item>
            </Command.Group>
          </Command.List>

          <div className="border-t border-surface-200 dark:border-slate-700/50 px-4 py-2 text-[11px] text-stone-400 dark:text-slate-500">
            <kbd className="px-1.5 py-0.5 rounded bg-surface-100 dark:bg-surface-700 font-mono text-[10px]">↑↓</kbd> naviguer
            <kbd className="px-1.5 py-0.5 rounded bg-surface-100 dark:bg-surface-700 font-mono text-[10px] ml-2">↵</kbd> ouvrir
            <kbd className="px-1.5 py-0.5 rounded bg-surface-100 dark:bg-surface-700 font-mono text-[10px] ml-2">esc</kbd> fermer
          </div>
        </Command>
      </div>
    </div>
  );
}

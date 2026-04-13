import React, { useState } from 'react';
import { NavLink, useNavigate } from 'react-router-dom';
import { Home, FileSpreadsheet, Database, TrendingUp, LayoutDashboard, Map, Sun, Moon, LogOut, ChevronsLeft, ChevronsRight } from 'lucide-react';
import { useAuth } from '@/contexts/AuthContext';
import { useTheme } from '@/contexts/ThemeContext';
import { avatarColor } from '@/lib/helpers';

const NAV_ITEMS = [
  { to: '/', icon: Home, label: 'Accueil', end: true },
  { to: '/app/stock-tracking', icon: FileSpreadsheet, label: 'Suivi des Stocks' },
  { to: '/app/tri-materiel', icon: Database, label: 'Tri Matériel' },
  { to: '/app/data-merge', icon: TrendingUp, label: 'Indicateurs' },
  { to: '/dashboard', icon: LayoutDashboard, label: 'Tableau de bord' },
  { to: '/mapping', icon: Map, label: 'Mapping', adminOnly: true },
];

export default function Sidebar() {
  const [collapsed, setCollapsed] = useState(false);
  const { user, logout } = useAuth();
  const { isDark, toggleTheme } = useTheme();
  const navigate = useNavigate();

  const handleLogout = () => {
    logout();
    navigate('/login');
  };

  const filteredItems = NAV_ITEMS.filter(item => {
    if (item.adminOnly && user?.role === 'Client') return false;
    if (user?.role === 'Client' && item.to !== '/' && item.to !== '/dashboard') return false;
    return true;
  });

  const avatarGradient = avatarColor((user?.firstname || '') + (user?.lastname || ''));

  return (
    <aside className={`flex flex-col h-screen border-r border-surface-200/80 dark:border-slate-700/30 bg-card/80 dark:bg-surface-900/80 backdrop-blur-lg transition-all duration-300 ${collapsed ? 'w-[52px]' : 'w-[240px]'}`}>
      {/* Logo */}
      <div className={`flex items-center h-14 px-3 border-b border-surface-200/80 dark:border-slate-700/30 ${collapsed ? 'justify-center' : 'gap-2.5'}`}>
        <div className="w-7 h-7 rounded-lg bg-gradient-to-br from-teal-500 to-teal-700 flex items-center justify-center logo-glow flex-shrink-0">
          <span className="text-[10px] font-black text-white tracking-tight">M</span>
        </div>
        {!collapsed && (
          <span className="text-sm font-display font-bold text-stone-800 dark:text-slate-100 tracking-tight truncate">MCO Web Apps</span>
        )}
      </div>

      {/* Navigation */}
      <nav className="flex-1 py-2 px-2 space-y-0.5 overflow-y-auto">
        {filteredItems.map(item => {
          const Icon = item.icon;
          return (
            <NavLink
              key={item.to}
              to={item.to}
              end={item.end}
              title={collapsed ? item.label : undefined}
              className={({ isActive }) =>
                `flex items-center gap-2.5 px-2.5 py-2 rounded-lg text-sm font-medium transition-colors duration-150 ${
                  isActive
                    ? 'bg-teal-50 dark:bg-teal-900/20 text-teal-700 dark:text-teal-400'
                    : 'text-stone-600 dark:text-slate-400 hover:bg-surface-100 dark:hover:bg-surface-700/50 hover:text-stone-800 dark:hover:text-slate-200'
                } ${collapsed ? 'justify-center' : ''}`
              }
            >
              <Icon className="w-4 h-4 flex-shrink-0" />
              {!collapsed && <span className="truncate">{item.label}</span>}
            </NavLink>
          );
        })}
      </nav>

      {/* Bottom section */}
      <div className="border-t border-surface-200/80 dark:border-slate-700/30 p-2 space-y-1">
        <button
          onClick={toggleTheme}
          title={isDark ? 'Mode clair' : 'Mode sombre'}
          className={`flex items-center gap-2.5 w-full px-2.5 py-2 rounded-lg text-sm font-medium text-stone-500 dark:text-slate-400 hover:bg-surface-100 dark:hover:bg-surface-700/50 transition-colors duration-150 ${collapsed ? 'justify-center' : ''}`}
        >
          {isDark ? <Sun className="w-4 h-4 text-amber-400 flex-shrink-0" /> : <Moon className="w-4 h-4 flex-shrink-0" />}
          {!collapsed && <span>{isDark ? 'Mode clair' : 'Mode sombre'}</span>}
        </button>

        <button
          onClick={() => setCollapsed(!collapsed)}
          title={collapsed ? 'Déplier' : 'Replier'}
          className={`flex items-center gap-2.5 w-full px-2.5 py-2 rounded-lg text-sm font-medium text-stone-500 dark:text-slate-400 hover:bg-surface-100 dark:hover:bg-surface-700/50 transition-colors duration-150 ${collapsed ? 'justify-center' : ''}`}
        >
          {collapsed ? <ChevronsRight className="w-4 h-4 flex-shrink-0" /> : <ChevronsLeft className="w-4 h-4 flex-shrink-0" />}
          {!collapsed && <span>Replier</span>}
        </button>

        {/* User info */}
        {user && (
          <div className={`flex items-center gap-2 px-2.5 py-2 ${collapsed ? 'justify-center' : ''}`}>
            <div className={`w-7 h-7 rounded-full bg-gradient-to-br ${avatarGradient} flex items-center justify-center flex-shrink-0`}>
              <span className="text-[10px] font-bold text-white">{user.firstname?.[0]}{user.lastname?.[0]}</span>
            </div>
            {!collapsed && (
              <div className="flex-1 min-w-0">
                <p className="text-xs font-medium text-stone-700 dark:text-slate-300 truncate">{user.firstname} {user.lastname}</p>
                <p className="text-[10px] text-stone-500 dark:text-slate-500 truncate">{user.role || 'Utilisateur'}</p>
              </div>
            )}
            <button
              onClick={handleLogout}
              title="Déconnexion"
              className="p-1.5 rounded-lg text-stone-400 hover:text-red-500 hover:bg-red-50 dark:hover:bg-red-900/20 transition-colors duration-150 flex-shrink-0"
            >
              <LogOut className="w-3.5 h-3.5" />
            </button>
          </div>
        )}
      </div>
    </aside>
  );
}

import { useNavigate, useLocation, useParams } from 'react-router-dom';
import { Sun, Moon, LogOut, ArrowLeft } from 'lucide-react';
import { useAuth } from '@/contexts/AuthContext';
import { useTheme } from '@/contexts/ThemeContext';
import { avatarColor } from '@/lib/helpers';
import { APPS_CONFIG } from '@/lib/constants';

export default function TopNav() {
  const navigate = useNavigate();
  const location = useLocation();
  const { appId } = useParams();
  const { user, logout } = useAuth();
  const { isDark, toggleTheme } = useTheme();

  const isHome = location.pathname === '/';
  const currentApp = appId ? APPS_CONFIG.find(a => a.id === appId) : null;
  const pageTitles = { '/dashboard': 'Tableau de bord', '/mapping': 'Mapping' };
  const pageTitle = pageTitles[location.pathname];

  const avatarGradient = avatarColor((user?.firstname || '') + (user?.lastname || ''));

  return (
    <header className="sticky top-0 z-50 h-14 border-b border-surface-200/60 dark:border-white/[0.05] bg-card/80 dark:bg-surface-950/80 backdrop-blur-xl">
      <div className="h-full px-5 flex items-center justify-between max-w-[1400px] mx-auto">

        {/* Left */}
        <div className="flex items-center gap-2.5">
          {!isHome && (
            <button
              onClick={() => navigate(-1)}
              className="p-1.5 rounded-lg text-stone-400 hover:text-stone-700 dark:hover:text-slate-200 hover:bg-surface-100 dark:hover:bg-white/[0.06] transition-all duration-150"
            >
              <ArrowLeft className="w-4 h-4" />
            </button>
          )}
          <div className="flex items-center gap-2.5 cursor-pointer group" onClick={() => navigate('/')}>
            <div className="w-7 h-7 rounded-lg bg-gradient-to-br from-teal-500 to-teal-700 flex items-center justify-center shadow-md shadow-teal-500/25 group-hover:shadow-teal-500/40 transition-shadow duration-200 logo-glow">
              <span className="text-[10px] font-black text-white tracking-tight">M</span>
            </div>
            <span className="text-sm font-display font-bold text-stone-800 dark:text-slate-100 group-hover:text-teal-700 dark:group-hover:text-teal-400 transition-colors duration-150">
              MCO Web Apps
            </span>
          </div>
          {(currentApp || pageTitle) && (
            <>
              <span className="text-stone-300 dark:text-slate-700 text-sm select-none">/</span>
              {currentApp ? (
                <div className={`flex items-center gap-1.5 px-2.5 py-1 rounded-md text-xs font-semibold ${currentApp.colorLight}`}>
                  <currentApp.icon className="w-3.5 h-3.5" />
                  {currentApp.name}
                </div>
              ) : (
                <span className="text-sm font-medium text-stone-600 dark:text-slate-400">{pageTitle}</span>
              )}
            </>
          )}
        </div>

        {/* Right */}
        <div className="flex items-center gap-1">
          <button
            onClick={toggleTheme}
            className="p-2 rounded-lg text-stone-500 dark:text-slate-400 hover:bg-surface-100 dark:hover:bg-white/[0.06] transition-all duration-150"
            title={isDark ? 'Mode clair' : 'Mode sombre'}
          >
            {isDark ? <Sun className="w-4 h-4 text-amber-400" /> : <Moon className="w-4 h-4" />}
          </button>
          <div className="h-5 w-px bg-surface-200 dark:bg-white/10 mx-1" />
          {user && (
            <div className="flex items-center gap-2 px-1">
              <div className={`w-7 h-7 rounded-full bg-gradient-to-br ${avatarGradient} flex items-center justify-center ring-2 ring-white dark:ring-surface-900 shadow-sm flex-shrink-0`}>
                <span className="text-[10px] font-bold text-white select-none">{user.firstname?.[0]}{user.lastname?.[0]}</span>
              </div>
              <span className="text-sm font-medium text-stone-600 dark:text-slate-300 hidden sm:inline">
                {user.firstname} {user.lastname}
              </span>
            </div>
          )}
          <button
            onClick={() => { logout(); navigate('/login'); }}
            className="p-2 rounded-lg text-stone-400 hover:text-red-500 hover:bg-red-50 dark:hover:bg-red-500/10 transition-all duration-150"
            title="Déconnexion"
          >
            <LogOut className="w-4 h-4" />
          </button>
        </div>
      </div>
    </header>
  );
}

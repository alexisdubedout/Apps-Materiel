import { useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { Loader2, AlertCircle, LogIn, Sun, Moon } from 'lucide-react';
import { useAuth } from '@/contexts/AuthContext';
import { useTheme } from '@/contexts/ThemeContext';

export default function LoginPage() {
  const [login, setLogin] = useState('');
  const [password, setPassword] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const { login: authLogin } = useAuth();
  const { isDark, toggleTheme } = useTheme();
  const navigate = useNavigate();

  const API_URL = import.meta.env.VITE_API_URL || '';

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError('');
    setLoading(true);
    try {
      const response = await fetch(`${API_URL}/api/auth/login`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ login, password }),
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data.detail || 'Erreur de connexion');
      authLogin(data.user, data.token);
      navigate('/');
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-surface-50 dark:bg-surface-950 flex relative overflow-hidden transition-colors duration-300">
      <div className="absolute inset-0 dot-grid opacity-50" />
      <div className="absolute inset-0 overflow-hidden pointer-events-none">
        <div className="ambient-orb ambient-orb--1" style={{ top: '-15%', right: '5%' }} />
        <div className="ambient-orb ambient-orb--2" style={{ bottom: '-12%', left: '-5%' }} />
      </div>

      {/* Theme toggle */}
      <button
        onClick={toggleTheme}
        className="fixed top-5 right-5 z-50 p-2.5 rounded-xl bg-card/80 dark:bg-surface-800/80 backdrop-blur-sm border border-surface-200 dark:border-white/[0.08] shadow-sm hover:shadow-md transition-all duration-200"
      >
        {isDark ? <Sun className="w-4 h-4 text-amber-400" /> : <Moon className="w-4 h-4 text-stone-500" />}
      </button>

      {/* Left branding panel */}
      <div className="hidden lg:flex flex-col justify-between w-[400px] flex-shrink-0 relative z-10 p-10 border-r border-surface-200/60 dark:border-white/[0.05] bg-card/40 dark:bg-surface-900/40 backdrop-blur-sm">
        <div className="flex items-center gap-3">
          <div className="w-9 h-9 rounded-xl bg-gradient-to-br from-teal-500 to-teal-700 flex items-center justify-center shadow-lg shadow-teal-500/30 logo-glow">
            <span className="text-sm font-black text-white tracking-tight">M</span>
          </div>
          <span className="text-base font-display font-bold text-stone-800 dark:text-slate-100">MCO Web Apps</span>
        </div>

        <div>
          <div className="text-4xl font-display font-extrabold text-stone-800 dark:text-slate-100 leading-tight mb-4">
            Vos outils métier,<br />
            <span className="text-gradient-warm">centralisés.</span>
          </div>
          <p className="text-stone-500 dark:text-slate-400 text-sm leading-relaxed">
            Traitement Excel, suivi des stocks, tableau de bord temps réel — tout ce dont votre équipe a besoin au même endroit.
          </p>
        </div>

        <p className="text-xs text-stone-400 dark:text-slate-600">
          MCO Matériel — Authentification sécurisée Redmine
        </p>
      </div>

      {/* Right form panel */}
      <div className="flex-1 flex items-center justify-center p-6 relative z-10">
        <div className="w-full max-w-sm animate-hero-in">
          <div className="flex items-center gap-2.5 mb-8 lg:hidden">
            <div className="w-8 h-8 rounded-lg bg-gradient-to-br from-teal-500 to-teal-700 flex items-center justify-center logo-glow">
              <span className="text-[11px] font-black text-white">M</span>
            </div>
            <span className="text-sm font-display font-bold text-stone-800 dark:text-slate-100">MCO Web Apps</span>
          </div>

          <h2 className="text-2xl font-display font-bold text-stone-800 dark:text-slate-100 mb-1.5">Connexion</h2>
          <p className="text-stone-500 dark:text-slate-400 text-sm mb-8">Utilisez vos identifiants Redmine</p>

          <form onSubmit={handleSubmit} className="space-y-5">
            <div className="animate-slide-up space-y-1.5" style={{ animationDelay: '80ms', animationFillMode: 'both' }}>
              <label className="block text-sm font-medium text-stone-700 dark:text-slate-300">Identifiant</label>
              <input
                type="text" value={login} onChange={e => setLogin(e.target.value)}
                placeholder="Votre identifiant" required autoComplete="username"
                className="w-full px-4 py-3 rounded-xl bg-surface-50/80 dark:bg-white/[0.04] border border-surface-200 dark:border-white/[0.08] text-stone-800 dark:text-white text-sm placeholder-stone-400 focus:border-teal-500 focus:ring-2 focus:ring-teal-500/20 outline-none transition-all"
              />
            </div>

            <div className="animate-slide-up space-y-1.5" style={{ animationDelay: '120ms', animationFillMode: 'both' }}>
              <label className="block text-sm font-medium text-stone-700 dark:text-slate-300">Mot de passe</label>
              <input
                type="password" value={password} onChange={e => setPassword(e.target.value)}
                placeholder="Votre mot de passe" required autoComplete="current-password"
                className="w-full px-4 py-3 rounded-xl bg-surface-50/80 dark:bg-white/[0.04] border border-surface-200 dark:border-white/[0.08] text-stone-800 dark:text-white text-sm placeholder-stone-400 focus:border-teal-500 focus:ring-2 focus:ring-teal-500/20 outline-none transition-all"
              />
            </div>

            {error && (
              <div className="p-3.5 rounded-xl bg-red-50 dark:bg-red-500/10 border border-red-200 dark:border-red-500/20 flex items-start gap-2.5 animate-scale-in">
                <AlertCircle className="w-4 h-4 text-red-500 flex-shrink-0 mt-0.5" />
                <p className="text-red-600 dark:text-red-400 text-sm">{error}</p>
              </div>
            )}

            <div className="animate-slide-up" style={{ animationDelay: '160ms', animationFillMode: 'both' }}>
              <button
                type="submit" disabled={loading}
                className={`
                  w-full py-3 rounded-xl font-semibold text-sm
                  flex items-center justify-center gap-2 transition-all duration-200
                  ${loading
                    ? 'bg-surface-200 dark:bg-white/[0.06] text-stone-400 cursor-not-allowed'
                    : 'bg-gradient-to-r from-teal-600 to-teal-500 hover:from-teal-700 hover:to-teal-600 text-white shadow-lg shadow-teal-500/25 hover:shadow-teal-500/35 hover:-translate-y-0.5 active:scale-[0.99]'
                  }
                `}
              >
                {loading
                  ? <><Loader2 className="w-4 h-4 animate-spin" /> Connexion...</>
                  : <><LogIn className="w-4 h-4" /> Se connecter</>
                }
              </button>
            </div>
          </form>
        </div>
      </div>
    </div>
  );
}

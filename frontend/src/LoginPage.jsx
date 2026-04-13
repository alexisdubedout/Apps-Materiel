import React, { useState } from 'react';
import { Loader2, AlertCircle, LogIn, Shield, Sun, Moon } from 'lucide-react';

function LoginPage({ onLoginSuccess, isDark, setIsDark }) {
  const [login, setLogin] = useState('');
  const [password, setPassword] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');

  const API_URL = import.meta.env.VITE_API_URL || '';

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError('');
    setLoading(true);
    try {
      const response = await fetch(`${API_URL}/api/auth/login`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ login, password })
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data.detail || 'Erreur de connexion');
      localStorage.setItem('token', data.token);
      localStorage.setItem('user', JSON.stringify(data.user));
      onLoginSuccess(data.user);
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-surface-50 dark:bg-surface-950 transition-colors duration-300 flex items-center justify-center p-6 relative overflow-hidden">
      {/* Ambient orbs */}
      <div className="fixed inset-0 overflow-hidden pointer-events-none z-0">
        <div className="ambient-orb ambient-orb--1" style={{ top: '-15%', right: '10%' }} />
        <div className="ambient-orb ambient-orb--2" style={{ bottom: '-10%', left: '5%' }} />
      </div>

      {/* Theme toggle */}
      <button
        onClick={() => setIsDark(!isDark)}
        className="fixed top-5 right-5 z-50 p-2.5 rounded-lg bg-card dark:bg-surface-800 border border-surface-200 dark:border-slate-700/50 shadow-sm hover:shadow-md transition-all duration-150"
      >
        {isDark ? <Sun className="w-4 h-4 text-amber-400" /> : <Moon className="w-4 h-4 text-stone-500" />}
      </button>

      <div className="relative z-10 w-full max-w-sm animate-hero-in">
        <div className="bg-card/90 dark:bg-surface-800/90 backdrop-blur-sm rounded-lg p-8 border border-surface-200 dark:border-slate-700/40 shadow-xl shadow-stone-300/20 dark:shadow-black/20">
          <div className="text-center mb-8">
            <div className="inline-flex p-3.5 rounded-xl bg-gradient-to-br from-teal-500 to-teal-700 mb-4 logo-glow">
              <Shield className="w-7 h-7 text-white" />
            </div>
            <h1 className="text-xl font-display font-bold tracking-tight mb-1">
              <span className="text-gradient-warm">MCO Web Apps</span>
            </h1>
            <p className="text-sm text-stone-500 dark:text-slate-400">
              Connexion avec vos identifiants Redmine
            </p>
          </div>

          <form onSubmit={handleSubmit} className="space-y-4">
            <div className="animate-slide-up" style={{ animationDelay: '100ms' }}>
              <label className="block text-sm font-medium text-stone-700 dark:text-slate-300 mb-1.5">Identifiant</label>
              <input
                type="text"
                value={login}
                onChange={(e) => setLogin(e.target.value)}
                placeholder="Votre identifiant"
                required
                className="w-full px-3.5 py-2.5 rounded-lg bg-surface-50/80 dark:bg-surface-900/80 border border-surface-200 dark:border-slate-700/50 text-stone-800 dark:text-white text-sm placeholder-stone-400 focus:border-teal-500 focus:ring-2 focus:ring-teal-500/20 outline-none transition-all"
              />
            </div>

            <div className="animate-slide-up" style={{ animationDelay: '150ms' }}>
              <label className="block text-sm font-medium text-stone-700 dark:text-slate-300 mb-1.5">Mot de passe</label>
              <input
                type="password"
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                placeholder="Votre mot de passe"
                required
                className="w-full px-3.5 py-2.5 rounded-lg bg-surface-50/80 dark:bg-surface-900/80 border border-surface-200 dark:border-slate-700/50 text-stone-800 dark:text-white text-sm placeholder-stone-400 focus:border-teal-500 focus:ring-2 focus:ring-teal-500/20 outline-none transition-all"
              />
            </div>

            {error && (
              <div className="p-3 rounded-lg bg-red-50 dark:bg-red-900/10 border border-red-200 dark:border-red-800/30 flex items-start gap-2.5 animate-scale-in">
                <AlertCircle className="w-4 h-4 text-red-500 flex-shrink-0 mt-0.5" />
                <p className="text-red-600 dark:text-red-400 text-sm">{error}</p>
              </div>
            )}

            <div className="animate-slide-up" style={{ animationDelay: '200ms' }}>
              <button
                type="submit"
                disabled={loading}
                className={`w-full py-2.5 rounded-lg font-semibold text-sm transition-all duration-200 flex items-center justify-center gap-2 mt-2 ${
                  loading
                    ? 'bg-surface-200 dark:bg-slate-700 text-stone-400 cursor-not-allowed'
                    : 'bg-teal-600 hover:bg-teal-700 text-white hover:shadow-lg hover:shadow-teal-500/20 active:scale-[0.98]'
                }`}
              >
                {loading ? (<><Loader2 className="w-4 h-4 animate-spin" /> Connexion...</>) : (<><LogIn className="w-4 h-4" /> Se connecter</>)}
              </button>
            </div>
          </form>

          <p className="text-center text-[11px] text-stone-500 mt-6 animate-fade-in" style={{ animationDelay: '400ms' }}>
            Authentification via Redmine LGM
          </p>
        </div>
      </div>
    </div>
  );
}

export default LoginPage;

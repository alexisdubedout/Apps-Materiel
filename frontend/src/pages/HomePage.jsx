import { useNavigate } from 'react-router-dom';
import { ArrowUpRight } from 'lucide-react';
import { useAuth } from '@/contexts/AuthContext';
import { APPS_CONFIG } from '@/lib/constants';

function AppCard({ app, onClick, index }) {
  const Icon = app.icon;
  return (
    <button
      onClick={onClick}
      className="group relative text-left w-full animate-slide-up"
      style={{ animationDelay: `${80 + index * 70}ms`, animationFillMode: 'both' }}
    >
      <div className={`
        relative overflow-hidden rounded-2xl p-6
        bg-card dark:bg-surface-800/80
        border border-surface-200/80 dark:border-white/[0.06]
        shadow-sm
        hover:shadow-2xl hover:shadow-stone-300/30 dark:hover:shadow-black/40
        hover:-translate-y-2
        transition-all duration-300 ease-out
        border-l-[3px] ${app.accent}
      `}>
        <div className={`
          absolute inset-0 opacity-0 group-hover:opacity-100 transition-opacity duration-500
          bg-gradient-to-br
          ${app.id === 'stock-tracking' ? 'from-teal-500/[0.05]' :
            app.id === 'tri-materiel' ? 'from-emerald-500/[0.05]' :
            app.id === 'data-merge' ? 'from-amber-500/[0.05]' :
            'from-rose-500/[0.05]'}
          to-transparent rounded-2xl
        `} />

        <div className="relative">
          <div className="flex items-start justify-between mb-5">
            <div className={`
              p-3 rounded-xl ${app.colorLight}
              group-hover:scale-110 group-hover:shadow-lg
              transition-all duration-300
            `}>
              <Icon className="w-5 h-5" />
            </div>
            <div className="opacity-0 group-hover:opacity-100 -translate-x-1 group-hover:translate-x-0 transition-all duration-300">
              <ArrowUpRight className="w-4 h-4 text-stone-400 dark:text-slate-500" />
            </div>
          </div>
          <h3 className="text-[15px] font-display font-bold text-stone-800 dark:text-slate-100 mb-1.5 tracking-tight">
            {app.name}
          </h3>
          <p className="text-[13px] text-stone-500 dark:text-slate-400 leading-relaxed">
            {app.description}
          </p>
        </div>
      </div>
    </button>
  );
}

export default function HomePage() {
  const { user } = useAuth();
  const navigate = useNavigate();

  const handleSelectApp = (app) => {
    if (app.id === 'tdb') navigate('/dashboard');
    else navigate(`/app/${app.id}`);
  };

  const filteredApps = APPS_CONFIG.filter(
    app => user?.role !== 'Client' || app.id === 'tdb'
  );

  return (
    <div className="relative min-h-[calc(100vh-3.5rem)] overflow-hidden">
      <div className="absolute inset-0 dot-grid" />
      <div className="absolute inset-0 overflow-hidden pointer-events-none">
        <div className="ambient-orb ambient-orb--1" style={{ top: '-10%', right: '-3%' }} />
        <div className="ambient-orb ambient-orb--2" style={{ bottom: '-8%', left: '-2%' }} />
        <div className="ambient-orb ambient-orb--3" style={{ top: '40%', left: '55%' }} />
      </div>

      <div className="relative z-10 max-w-4xl mx-auto px-6 pt-16 pb-16">
        <div className="mb-12 animate-hero-in">
          <h1 className="text-[42px] font-display font-extrabold tracking-tight leading-[1.1] mb-4">
            <span className="text-stone-800 dark:text-slate-100">Bonjour, </span>
            <span className="text-gradient-warm">{user?.firstname}</span>
            <span className="ml-2">👋</span>
          </h1>
          <p className="text-stone-500 dark:text-slate-400 text-base max-w-md">
            Choisissez un outil pour commencer votre session de travail.
          </p>
        </div>

        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          {filteredApps.map((app, idx) => (
            <AppCard key={app.id} app={app} index={idx} onClick={() => handleSelectApp(app)} />
          ))}
        </div>
      </div>
    </div>
  );
}

import { useState } from 'react';
import { useParams, useNavigate } from 'react-router-dom';
import {
  Upload, FileSpreadsheet, Calendar, CheckCircle2,
  Loader2, Download, AlertCircle, Database, X, File
} from 'lucide-react';
import { useToast } from '@/components/ui/ToastProvider';
import { APPS_CONFIG, FUNNY_MESSAGES } from '@/lib/constants';

function FileUploadZone({ fileConfig, file, onFileChange, onClear }) {
  const [isDragging, setIsDragging] = useState(false);

  const handleDragOver = (e) => { e.preventDefault(); setIsDragging(true); };
  const handleDragLeave = () => setIsDragging(false);
  const handleDrop = (e) => {
    e.preventDefault(); setIsDragging(false);
    if (e.dataTransfer.files[0]) onFileChange(e.dataTransfer.files[0]);
  };
  const handleFileInput = (e) => { if (e.target.files[0]) onFileChange(e.target.files[0]); };

  return (
    <div className="space-y-2">
      <label className="block text-sm font-medium text-stone-700 dark:text-slate-300">
        {fileConfig.label}
      </label>

      {file ? (
        /* File selected state */
        <div className="flex items-center gap-3 p-4 rounded-xl bg-emerald-50 dark:bg-emerald-500/10 border border-emerald-200 dark:border-emerald-500/20 animate-fade-in">
          <div className="w-9 h-9 rounded-lg bg-emerald-100 dark:bg-emerald-500/20 flex items-center justify-center flex-shrink-0">
            <File className="w-4 h-4 text-emerald-600 dark:text-emerald-400" />
          </div>
          <div className="flex-1 min-w-0">
            <p className="text-sm font-medium text-emerald-700 dark:text-emerald-400 truncate">{file.name}</p>
            <p className="text-xs text-emerald-600/70 dark:text-emerald-500 mt-0.5">{(file.size / 1024).toFixed(1)} Ko</p>
          </div>
          <button
            onClick={onClear}
            className="p-1.5 rounded-lg text-emerald-500 hover:text-emerald-700 hover:bg-emerald-100 dark:hover:bg-emerald-500/20 transition-colors"
          >
            <X className="w-3.5 h-3.5" />
          </button>
        </div>
      ) : (
        /* Drop zone */
        <div
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onDrop={handleDrop}
          className={`
            relative group rounded-xl border-2 border-dashed p-8
            flex flex-col items-center justify-center gap-3
            cursor-pointer transition-all duration-200
            ${isDragging
              ? 'border-teal-400 bg-teal-50/60 dark:bg-teal-500/10 scale-[1.01]'
              : 'border-surface-300 dark:border-white/[0.08] bg-surface-50/50 dark:bg-white/[0.02] hover:border-teal-400/60 dark:hover:border-teal-500/40 hover:bg-teal-50/30 dark:hover:bg-teal-500/[0.05]'
            }
          `}
        >
          <input
            type="file"
            accept={fileConfig.accept}
            onChange={handleFileInput}
            className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
          />
          <div className={`
            w-12 h-12 rounded-xl flex items-center justify-center transition-all duration-200
            ${isDragging
              ? 'bg-teal-100 dark:bg-teal-500/20 scale-110'
              : 'bg-surface-100 dark:bg-white/[0.06] group-hover:bg-teal-50 dark:group-hover:bg-teal-500/10 group-hover:scale-105'
            }
          `}>
            <Upload className={`w-5 h-5 transition-colors duration-200 ${isDragging ? 'text-teal-600' : 'text-stone-400 dark:text-slate-500 group-hover:text-teal-500'}`} />
          </div>
          <div className="text-center">
            <p className="text-sm font-medium text-stone-600 dark:text-slate-300">
              {isDragging ? 'Déposez le fichier ici' : 'Glissez votre fichier ici'}
            </p>
            <p className="text-xs text-stone-400 dark:text-slate-500 mt-1">ou <span className="text-teal-600 dark:text-teal-400 font-medium">cliquez pour parcourir</span></p>
          </div>
        </div>
      )}
    </div>
  );
}

function StepProgress({ progress, message }) {
  return (
    <div className="space-y-3 animate-fade-in">
      <div className="relative h-1.5 bg-surface-200 dark:bg-white/[0.08] rounded-full overflow-hidden">
        <div
          className="absolute inset-y-0 left-0 bg-gradient-to-r from-teal-500 to-teal-400 rounded-full transition-all duration-700 ease-out"
          style={{ width: `${progress}%` }}
        />
        {/* Shimmer */}
        <div
          className="absolute inset-y-0 w-1/3 bg-gradient-to-r from-transparent via-white/40 to-transparent rounded-full animate-shimmer"
          style={{ left: `${Math.max(0, progress - 35)}%` }}
        />
      </div>
      <div className="flex items-center justify-between text-xs text-stone-400 dark:text-slate-500">
        <span className="italic">{message}</span>
        <span className="font-mono tabular-nums">{Math.round(progress)}%</span>
      </div>
    </div>
  );
}

export default function ProcessingPage() {
  const { appId } = useParams();
  const navigate = useNavigate();
  const app = APPS_CONFIG.find(a => a.id === appId);

  const [files, setFiles] = useState({});
  const [params, setParams] = useState({});
  const [isProcessing, setIsProcessing] = useState(false);
  const [progress, setProgress] = useState(0);
  const [funnyMessage, setFunnyMessage] = useState('');
  const [result, setResult] = useState(null);
  const [error, setError] = useState(null);
  const toast = useToast();

  if (!app) {
    return (
      <div className="flex items-center justify-center min-h-[calc(100vh-3.5rem)]">
        <div className="text-center">
          <p className="text-stone-500 dark:text-slate-400 mb-4">Application introuvable</p>
          <button onClick={() => navigate('/')} className="text-teal-600 hover:text-teal-700 text-sm font-medium">
            ← Retour à l'accueil
          </button>
        </div>
      </div>
    );
  }

  const handleFileChange = (fileId, file) => { setFiles(prev => ({ ...prev, [fileId]: file })); setError(null); };
  const handleFileClear = (fileId) => setFiles(prev => { const n = { ...prev }; delete n[fileId]; return n; });
  const handleParamChange = (paramId, value) => setParams(prev => ({ ...prev, [paramId]: value }));
  const canProcess = () => app.files.every(f => files[f.id]) && app.params.every(p => params[p.id]);

  const handleProcess = async () => {
    setIsProcessing(true); setProgress(0); setError(null); setResult(null);
    const messageInterval = setInterval(() => {
      setFunnyMessage(FUNNY_MESSAGES[Math.floor(Math.random() * FUNNY_MESSAGES.length)]);
    }, 2800);
    try {
      const API_URL = import.meta.env.VITE_API_URL || '';
      const token = localStorage.getItem('token');
      const formData = new FormData();
      Object.entries(files).forEach(([key, file]) => formData.append(`file_${key}`, file));
      formData.append('params', JSON.stringify(params));

      const progressInterval = setInterval(() => {
        setProgress(prev => prev >= 88 ? prev : prev + Math.random() * 12);
      }, 400);

      const response = await fetch(`${API_URL}/api/process/${app.id}`, {
        method: 'POST',
        headers: { 'Authorization': `Bearer ${token}` },
        body: formData,
      });
      clearInterval(progressInterval);

      if (!response.ok) {
        let errorData;
        try { errorData = await response.json(); } catch { errorData = { detail: `Erreur HTTP ${response.status}` }; }
        throw new Error(
          Array.isArray(errorData.detail)
            ? errorData.detail.map(e => `${e.loc?.join(' > ') || ''}: ${e.msg}`).join('\n')
            : errorData.detail || 'Erreur lors du traitement'
        );
      }

      const cd = response.headers.get('Content-Disposition');
      let filename = `resultat_${app.id}_${Date.now()}.xlsx`;
      if (cd) { const m = cd.match(/filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/); if (m?.[1]) filename = m[1].replace(/['"]/g, ''); }

      const blob = await response.blob();
      setProgress(100);
      setResult({ url: window.URL.createObjectURL(blob), filename });
      toast.success('Traitement terminé avec succès !');
    } catch (err) {
      setError(err.message || 'Une erreur est survenue');
      toast.error('Le traitement a échoué');
    } finally {
      clearInterval(messageInterval);
      setIsProcessing(false);
      setFunnyMessage('');
    }
  };

  const handleDownload = () => {
    let fn = result.filename.replace(/\.xlsx[_\s]*$/i, '.xlsx');
    if (!fn.toLowerCase().endsWith('.xlsx')) fn += '.xlsx';
    const link = document.createElement('a');
    link.href = result.url; link.download = fn; link.style.display = 'none';
    document.body.appendChild(link); link.click(); document.body.removeChild(link);
    setTimeout(() => window.URL.revokeObjectURL(result.url), 100);
  };

  const Icon = app.icon;

  return (
    <div className="min-h-[calc(100vh-3.5rem)] bg-surface-50 dark:bg-surface-950 relative overflow-hidden">
      {/* Subtle background */}
      <div className="absolute inset-0 dot-grid opacity-40" />

      <div className="relative z-10 max-w-2xl mx-auto px-6 py-12">
        {/* Page header */}
        <div className="mb-8 animate-hero-in">
          <div className={`inline-flex items-center gap-2 px-3 py-1.5 rounded-lg ${app.colorLight} text-xs font-semibold mb-4`}>
            <Icon className="w-3.5 h-3.5" />
            {app.name}
          </div>
          <h1 className="text-2xl font-display font-bold text-stone-800 dark:text-slate-100 tracking-tight mb-2">
            Traitement de fichier
          </h1>
          <p className="text-stone-500 dark:text-slate-400 text-sm">{app.description}</p>

          {app.id === 'tri-materiel' && (
            <button
              onClick={() => navigate('/mapping')}
              className="mt-3 inline-flex items-center gap-1.5 text-xs font-medium text-stone-500 dark:text-slate-400 hover:text-teal-600 dark:hover:text-teal-400 transition-colors"
            >
              <Database className="w-3.5 h-3.5" />
              Gérer le mapping des emplacements
            </button>
          )}
        </div>

        {/* Main card */}
        <div className="bg-card dark:bg-surface-800/80 rounded-2xl border border-surface-200/80 dark:border-white/[0.06] shadow-lg shadow-stone-200/50 dark:shadow-black/20 animate-slide-up" style={{ animationDelay: '80ms', animationFillMode: 'both' }}>

          {/* Card header */}
          <div className={`h-1.5 rounded-t-2xl bg-gradient-to-r ${
            app.id === 'stock-tracking' ? 'from-teal-500 to-teal-400' :
            app.id === 'tri-materiel' ? 'from-emerald-500 to-emerald-400' :
            app.id === 'data-merge' ? 'from-amber-500 to-amber-400' :
            'from-rose-500 to-rose-400'
          }`} />

          <div className="p-7 space-y-6">
            {/* Files */}
            <div className="space-y-4">
              <h2 className="text-xs font-semibold text-stone-400 dark:text-slate-500 uppercase tracking-wider flex items-center gap-2">
                <FileSpreadsheet className="w-3.5 h-3.5" />
                Fichiers requis
              </h2>
              {app.files.map(fc => (
                <FileUploadZone
                  key={fc.id}
                  fileConfig={fc}
                  file={files[fc.id]}
                  onFileChange={(file) => handleFileChange(fc.id, file)}
                  onClear={() => handleFileClear(fc.id)}
                />
              ))}
            </div>

            {/* Params */}
            {app.params.length > 0 && (
              <div className="space-y-4">
                <h2 className="text-xs font-semibold text-stone-400 dark:text-slate-500 uppercase tracking-wider flex items-center gap-2">
                  <Calendar className="w-3.5 h-3.5" />
                  Paramètres
                </h2>
                {app.params.map(param => (
                  <div key={param.id} className="space-y-2">
                    <label className="block text-sm font-medium text-stone-700 dark:text-slate-300">{param.label}</label>
                    <input
                      type={param.type}
                      placeholder={param.placeholder}
                      value={params[param.id] || ''}
                      onChange={(e) => handleParamChange(param.id, e.target.value)}
                      className="w-full px-3.5 py-2.5 bg-surface-50 dark:bg-white/[0.04] border border-surface-200 dark:border-white/[0.08] rounded-xl text-stone-800 dark:text-white text-sm placeholder-stone-400 focus:border-teal-500 focus:ring-2 focus:ring-teal-500/20 outline-none transition-all"
                    />
                  </div>
                ))}
              </div>
            )}

            {/* Submit */}
            <button
              onClick={handleProcess}
              disabled={!canProcess() || isProcessing}
              className={`
                w-full py-3 rounded-xl font-semibold text-sm
                flex items-center justify-center gap-2.5
                transition-all duration-200
                ${canProcess() && !isProcessing
                  ? 'bg-gradient-to-r from-teal-600 to-teal-500 hover:from-teal-700 hover:to-teal-600 text-white shadow-lg shadow-teal-500/25 hover:shadow-teal-500/35 hover:-translate-y-0.5 active:translate-y-0 active:scale-[0.99]'
                  : 'bg-surface-200 dark:bg-white/[0.06] text-stone-400 dark:text-slate-500 cursor-not-allowed'
                }
              `}
            >
              {isProcessing
                ? <><Loader2 className="w-4 h-4 animate-spin" /> Traitement en cours...</>
                : <><Upload className="w-4 h-4" /> Lancer le traitement</>
              }
            </button>

            {/* Progress */}
            {isProcessing && <StepProgress progress={progress} message={funnyMessage} />}

            {/* Error */}
            {error && (
              <div className="p-4 bg-red-50 dark:bg-red-500/10 border border-red-200 dark:border-red-500/20 rounded-xl flex items-start gap-3 animate-fade-in">
                <AlertCircle className="w-5 h-5 text-red-500 flex-shrink-0 mt-0.5" />
                <div>
                  <p className="text-red-700 dark:text-red-400 font-semibold text-sm">Erreur de traitement</p>
                  <p className="text-red-600/80 dark:text-red-300/80 text-sm mt-0.5">{error}</p>
                </div>
              </div>
            )}

            {/* Success */}
            {result && (
              <div className="p-5 bg-emerald-50 dark:bg-emerald-500/10 border border-emerald-200 dark:border-emerald-500/20 rounded-xl animate-scale-in">
                <div className="flex items-center gap-4">
                  <div className="w-11 h-11 rounded-xl bg-emerald-100 dark:bg-emerald-500/20 flex items-center justify-center flex-shrink-0">
                    <CheckCircle2 className="w-5 h-5 text-emerald-600 dark:text-emerald-400" />
                  </div>
                  <div className="flex-1 min-w-0">
                    <p className="text-emerald-700 dark:text-emerald-400 font-bold text-sm">Terminé !</p>
                    <p className="text-emerald-600/70 dark:text-emerald-500 text-xs mt-0.5 truncate">{result.filename}</p>
                  </div>
                  <button
                    onClick={handleDownload}
                    className="flex items-center gap-2 px-4 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg font-semibold text-sm transition-all duration-150 shadow-md shadow-emerald-500/25 hover:shadow-emerald-500/35 flex-shrink-0"
                  >
                    <Download className="w-4 h-4" />
                    Télécharger
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

import { FileSpreadsheet, TrendingUp, Database, LayoutDashboard } from 'lucide-react';

export const APPS_CONFIG = [
  {
    id: 'stock-tracking',
    name: 'Suivi des Stocks',
    description: 'Mise à jour du suivi mensuel et semestriel des stocks',
    icon: FileSpreadsheet,
    color: 'bg-teal-600',
    colorLight: 'bg-teal-50 text-teal-700 dark:bg-teal-900/20 dark:text-teal-400',
    accent: 'border-teal-500',
    glow: 'card-glow--teal',
    files: [
      { id: 'tracking', label: 'Fichier de suivi', accept: '.xlsx,.xls' },
      { id: 'export', label: "Fichier d'export", accept: '.xlsx,.xls' }
    ],
    params: [
      { id: 'export_date', label: "Date d'export", type: 'date', placeholder: '' }
    ]
  },
  {
    id: 'tri-materiel',
    name: 'Tri Materiel',
    description: 'Tri automatique des donnees avec mapping des emplacements',
    icon: Database,
    color: 'bg-emerald-600',
    colorLight: 'bg-emerald-50 text-emerald-700 dark:bg-emerald-900/20 dark:text-emerald-400',
    accent: 'border-emerald-500',
    glow: 'card-glow--emerald',
    files: [
      { id: 'export', label: 'Fichier export brut', accept: '.xlsx,.xls' }
    ],
    params: []
  },
  {
    id: 'data-merge',
    name: 'Indicateurs',
    description: 'Aide a la generation des indicateurs mensuels',
    icon: TrendingUp,
    color: 'bg-amber-600',
    colorLight: 'bg-amber-50 text-amber-700 dark:bg-amber-900/20 dark:text-amber-400',
    accent: 'border-amber-500',
    glow: 'card-glow--amber',
    files: [
      { id: 'file1', label: 'Premier fichier', accept: '.xlsx,.xls' },
      { id: 'file2', label: 'Deuxieme fichier', accept: '.xlsx,.xls' }
    ],
    params: []
  },
  {
    id: 'tdb',
    name: 'Tableau de bord',
    description: 'Suivi en temps reel des devis, DSM, OPALES et fiches Redmine',
    icon: LayoutDashboard,
    color: 'bg-rose-600',
    colorLight: 'bg-rose-50 text-rose-700 dark:bg-rose-900/20 dark:text-rose-400',
    accent: 'border-rose-500',
    glow: 'card-glow--rose',
    files: [],
    params: []
  }
];

export const FUNNY_MESSAGES = [
  "Decompte des claviers noyes dans le cafe...",
  "Recensement des souris qui ont pris la fuite...",
  "Inventaire des ecrans qui ont vu des etoiles...",
  "Comptage des ventilateurs en vacances...",
  "Analyse des composants rescapes...",
];

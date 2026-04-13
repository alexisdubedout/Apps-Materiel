// frontend/src/tdbConfig.js

export const TDB_CONFIG = {
  widgets: [
    {
      id: 'fiches-obso',
      title: 'Fiches obso en cours',
      type: 'table',
      color: 'from-teal-400 to-teal-600',
      filter: {
        source: 'redmine',
        column: null,
        value: null
      },
      columns: ['ID', 'Titre', 'Statut', 'Assigne a', 'Echeance']
    },
    {
      id: 'devis-en-attente',
      title: 'Devis en attente',
      type: 'table',
      color: 'from-emerald-400 to-emerald-600',
      filter: {
        source: 'devis',
        column: 'Statut validation',
        value: 'Soumis'
      },
      columns: ['N° devis', 'Date de devis', 'Montant TTC']
    },
    {
      id: 'dsm-en-cours',
      title: 'DSM en cours',
      type: 'table',
      color: 'from-stone-400 to-stone-600',
      filter: {
        source: 'dsm',
        column: 'status',
        value: 'EN COURS'
      },
      columns: ['ticketid', 'description', 'Date de creation']
    },
    {
      id: 'opales-en-cours',
      title: 'OPALES en cours',
      type: 'table',
      color: 'from-rose-400 to-rose-600',
      filter: {
        source: 'opales',
        column: null,
        value: null
      },
      columns: ['N° LandesK', 'Etat', 'Objet']
    },
    {
      id: 'notes',
      title: 'Notes',
      type: 'notes',
      color: 'from-amber-400 to-amber-600'
    }
  ]
};

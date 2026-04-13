// frontend/src/services/excelService.js

const API_URL = import.meta.env.VITE_API_URL || '';

class ExcelService {
  async fetchAndParse() {
    const token = localStorage.getItem('token');
    
    if (!token) {
      throw new Error('Non authentifié');
    }

    const response = await fetch(`${API_URL}/api/tdb/fetch-and-parse`, {
      headers: {
        'Authorization': `Bearer ${token}`
      }
    });

    if (!response.ok) {
      const errorData = await response.json().catch(() => ({}));
      throw new Error(errorData.detail || errorData.error || 'Erreur lors de la récupération des données');
    }

    const data = await response.json();
    
    return {
      data: {
        devis: data.devis || [],
        dsm: data.dsm || [],
        opales: data.opales || [],
        redmine: data.redmine || []
      },
      timestamp: data.timestamp
    };
  }

  filterData(allData, filter) {
    if (!filter || !filter.source) {
      return [];
    }

    const sourceData = allData[filter.source] || [];

    // Si pas de filtre de colonne, retourner toutes les données
    if (!filter.column || !filter.value) {
      return sourceData;
    }

    // Filtrer selon la colonne et la valeur
    return sourceData.filter(row => {
      const cellValue = row[filter.column];
      if (!cellValue) return false;
      
      return cellValue.toString().trim().toUpperCase() === filter.value.toUpperCase();
    });
  }
}

export const excelService = new ExcelService();
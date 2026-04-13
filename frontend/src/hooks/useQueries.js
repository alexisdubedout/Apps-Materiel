import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import { api } from '@/lib/api';

// ─── TDB (Dashboard) ─────────────────────────────

export function useTDBData() {
  return useQuery({
    queryKey: ['tdb'],
    queryFn: async () => {
      const res = await api.get('/api/tdb/fetch-and-parse');
      if (!res.ok) throw new Error('Erreur lors du chargement des données TDB');
      return res.json();
    },
    staleTime: 5 * 60 * 1000, // 5 minutes
    retry: 1,
  });
}

// ─── Mapping ─────────────────────────────────────

export function useMappings() {
  return useQuery({
    queryKey: ['mappings'],
    queryFn: async () => {
      const res = await api.get('/api/mapping/emplacements');
      if (!res.ok) throw new Error('Erreur lors du chargement du mapping');
      return res.json();
    },
  });
}

export function useUpdateMapping() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: (data) => api.post('/api/mapping/emplacements', data).then(r => r.json()),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['mappings'] });
      queryClient.invalidateQueries({ queryKey: ['mapping-values'] });
    },
  });
}

export function useDeleteMapping() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: ({ code, user }) =>
      api.delete(`/api/mapping/emplacements/${code}`, { user }).then(r => r.json()),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['mappings'] });
      queryClient.invalidateQueries({ queryKey: ['mapping-history'] });
    },
  });
}

export function useMappingValues(field) {
  return useQuery({
    queryKey: ['mapping-values', field],
    queryFn: async () => {
      const res = await api.get(`/api/mapping/values/${field}`);
      if (!res.ok) throw new Error('Erreur chargement valeurs');
      return res.json();
    },
  });
}

export function useMappingHistory(limit = 100) {
  return useQuery({
    queryKey: ['mapping-history'],
    queryFn: async () => {
      const res = await api.get(`/api/mapping/history?limit=${limit}`);
      if (!res.ok) throw new Error('Erreur chargement historique');
      return res.json();
    },
    enabled: false, // load on demand
  });
}

// ─── Notes ───────────────────────────────────────

export function useNotes() {
  return useQuery({
    queryKey: ['notes'],
    queryFn: async () => {
      const res = await api.get('/api/tdb/notes');
      if (!res.ok) throw new Error('Erreur chargement notes');
      return res.json();
    },
    staleTime: 30 * 1000, // 30 seconds
  });
}

export function useSaveNote() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: (content) =>
      api.post('/api/tdb/notes', { content }).then(r => r.json()),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['notes'] });
    },
  });
}

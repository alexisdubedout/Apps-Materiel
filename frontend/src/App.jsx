import React from 'react';
import { Routes, Route, Navigate } from 'react-router-dom';
import { useAuth } from '@/contexts/AuthContext';
import AppLayout from '@/components/layout/AppLayout';
import LoginPage from '@/pages/LoginPage';
import HomePage from '@/pages/HomePage';
import ProcessingPage from '@/pages/ProcessingPage';
import DashboardPage from '@/pages/DashboardPage';
import MappingPage from '@/pages/MappingPage';

function ProtectedRoute({ children }) {
  const { user, isLoading } = useAuth();

  if (isLoading) {
    return (
      <div className="min-h-screen bg-surface-50 dark:bg-surface-950 flex items-center justify-center">
        <div className="w-7 h-7 rounded-lg bg-gradient-to-br from-teal-500 to-teal-700 flex items-center justify-center logo-glow animate-pulse">
          <span className="text-[10px] font-black text-white tracking-tight">M</span>
        </div>
      </div>
    );
  }

  if (!user) return <Navigate to="/login" replace />;
  return children;
}

function ClientGuard({ children }) {
  const { user } = useAuth();
  if (user?.role === 'Client') return <Navigate to="/dashboard" replace />;
  return children;
}

export default function App() {
  return (
    <Routes>
      <Route path="/login" element={<LoginPage />} />
      <Route
        path="/"
        element={
          <ProtectedRoute>
            <AppLayout />
          </ProtectedRoute>
        }
      >
        <Route index element={<HomePage />} />
        <Route path="app/:appId" element={<ClientGuard><ProcessingPage /></ClientGuard>} />
        <Route path="dashboard" element={<DashboardPage />} />
        <Route path="mapping" element={<ClientGuard><MappingPage /></ClientGuard>} />
      </Route>
      <Route path="*" element={<Navigate to="/" replace />} />
    </Routes>
  );
}

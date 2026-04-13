import React from 'react';
import { Toaster, toast } from 'sonner';
import { useTheme } from '@/contexts/ThemeContext';

export function useToast() {
  return toast;
}

export function ToastProvider({ children }) {
  const { isDark } = useTheme();

  return (
    <>
      {children}
      <Toaster
        theme={isDark ? 'dark' : 'light'}
        position="bottom-right"
        toastOptions={{
          style: {
            background: isDark ? '#161a26' : '#fefdfb',
            border: isDark ? '1px solid rgba(51, 65, 85, 0.5)' : '1px solid #e0d8cc',
            color: isDark ? '#e2e8f0' : '#44403c',
          },
        }}
        richColors
      />
    </>
  );
}

import { useState } from 'react';
import { Outlet } from 'react-router-dom';
import TopNav from './TopNav';
import CommandPalette from '@/components/CommandPalette';

export default function AppLayout() {
  const [cmdOpen, setCmdOpen] = useState(false);

  return (
    <div className="min-h-screen bg-surface-50 dark:bg-surface-950 transition-colors duration-300">
      <TopNav onCmdOpen={() => setCmdOpen(true)} />
      <main>
        <Outlet />
      </main>
      <CommandPalette open={cmdOpen} onOpenChange={setCmdOpen} />
    </div>
  );
}

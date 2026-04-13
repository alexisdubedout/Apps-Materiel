export function avatarColor(name) {
  let hash = 0;
  for (let i = 0; i < (name || '').length; i++) hash = name.charCodeAt(i) + ((hash << 5) - hash);
  const colors = [
    'from-teal-500 to-teal-700', 'from-emerald-500 to-emerald-700',
    'from-amber-500 to-amber-700', 'from-rose-500 to-rose-700',
    'from-violet-500 to-violet-700', 'from-sky-500 to-sky-700',
  ];
  return colors[Math.abs(hash) % colors.length];
}

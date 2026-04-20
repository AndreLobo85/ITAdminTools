// Icons — stroke-based line icons, 24x24, currentColor
const Icon = ({ name, size = 20, strokeWidth = 1.6 }) => {
  const paths = {
    lock: <><rect x="4" y="11" width="16" height="10" rx="2"/><path d="M8 11V7a4 4 0 018 0v4"/></>,
    key: <><circle cx="8" cy="15" r="3"/><path d="M11 15h9M17 15v3M20 15v2"/></>,
    users: <><circle cx="9" cy="9" r="3"/><path d="M3 20c0-3.3 2.7-6 6-6s6 2.7 6 6M17 10a3 3 0 100-6M21 20c0-2.5-1.5-4.6-3.7-5.5"/></>,
    shield: <><path d="M12 3l8 3v6c0 5-3.5 8.5-8 9-4.5-.5-8-4-8-9V6l8-3z"/></>,
    globe: <><circle cx="12" cy="12" r="9"/><path d="M3 12h18M12 3a13 13 0 010 18M12 3a13 13 0 000 18"/></>,
    cpu: <><rect x="6" y="6" width="12" height="12" rx="1.5"/><rect x="9" y="9" width="6" height="6" rx="0.5"/><path d="M9 3v3M15 3v3M9 18v3M15 18v3M3 9h3M3 15h3M18 9h3M18 15h3"/></>,
    drive: <><rect x="3" y="5" width="18" height="6" rx="1.5"/><rect x="3" y="13" width="18" height="6" rx="1.5"/><circle cx="7" cy="8" r=".8" fill="currentColor"/><circle cx="7" cy="16" r=".8" fill="currentColor"/></>,
    printer: <><path d="M7 8V4h10v4M5 8h14a2 2 0 012 2v5a2 2 0 01-2 2h-2v3H7v-3H5a2 2 0 01-2-2v-5a2 2 0 012-2z"/><circle cx="17.5" cy="11.5" r=".7" fill="currentColor"/></>,
    badge: <><path d="M9 3h6v4H9zM5 7h14v14H5z"/><circle cx="12" cy="13" r="2.5"/><path d="M8.5 19c.5-1.8 1.9-3 3.5-3s3 1.2 3.5 3"/></>,
    mail: <><rect x="3" y="5" width="18" height="14" rx="2"/><path d="M3 7l9 6 9-6"/></>,
    search: <><circle cx="11" cy="11" r="6"/><path d="M20 20l-4.5-4.5"/></>,
    chat: <><path d="M4 6a2 2 0 012-2h12a2 2 0 012 2v8a2 2 0 01-2 2h-7l-4 4v-4H6a2 2 0 01-2-2V6z"/></>,
    folder: <><path d="M3 7a2 2 0 012-2h4l2 2h8a2 2 0 012 2v8a2 2 0 01-2 2H5a2 2 0 01-2-2V7z"/></>,
    device: <><rect x="6" y="3" width="12" height="18" rx="2"/><path d="M10 18h4"/></>,
    play: <><path d="M7 5l12 7-12 7V5z" fill="currentColor" stroke="none"/></>,
    stop: <><rect x="6" y="6" width="12" height="12" rx="1.5" fill="currentColor" stroke="none"/></>,
    chevron: <><path d="M9 6l6 6-6 6"/></>,
    x: <><path d="M6 6l12 12M18 6L6 18"/></>,
    check: <><path d="M5 12l5 5L20 7"/></>,
    copy: <><rect x="8" y="8" width="12" height="12" rx="1.5"/><path d="M16 8V6a2 2 0 00-2-2H6a2 2 0 00-2 2v8a2 2 0 002 2h2"/></>,
    history: <><path d="M3 12a9 9 0 109-9 9 9 0 00-7 3.3M3 4v4h4M12 7v5l3 2"/></>,
    settings: <><circle cx="12" cy="12" r="3"/><path d="M19 12a7 7 0 00-.1-1.2l2-1.5-2-3.5-2.4.8a7 7 0 00-2-1.2L14 3h-4l-.5 2.4a7 7 0 00-2 1.2L5 5.8l-2 3.5 2 1.5a7 7 0 000 2.4l-2 1.5 2 3.5 2.4-.8a7 7 0 002 1.2L10 21h4l.5-2.4a7 7 0 002-1.2l2.4.8 2-3.5-2-1.5c.1-.4.1-.8.1-1.2z"/></>,
    terminal: <><rect x="3" y="4" width="18" height="16" rx="2"/><path d="M7 9l3 3-3 3M13 15h4"/></>,
    sparkle: <><path d="M12 3l1.8 5.2L19 10l-5.2 1.8L12 17l-1.8-5.2L5 10l5.2-1.8L12 3z"/></>,
    refresh: <><path d="M3 12a9 9 0 0115-6.7L21 8M21 4v4h-4M21 12a9 9 0 01-15 6.7L3 16M3 20v-4h4"/></>,
    power: <><path d="M12 3v9M6.3 7.5a8 8 0 1011.4 0"/></>,
    info: <><circle cx="12" cy="12" r="9"/><path d="M12 11v5M12 8v.5"/></>,
    grid: <><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/></>,
  };
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor"
         strokeWidth={strokeWidth} strokeLinecap="round" strokeLinejoin="round" aria-hidden>
      {paths[name] || paths.info}
    </svg>
  );
};

window.Icon = Icon;

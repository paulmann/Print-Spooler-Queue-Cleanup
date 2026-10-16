# Tray icon and notifications (v5)

## Presence

The tray icon appears when *Show tray icon* is on (default). The tooltip always shows the
live counter: `Print Spooler Queue Cleanup — Print queue jobs: N`.

## Double-click

Opens (restores) the main window and forces a refresh — the fastest way to see what is
stuck.

## Right-click menu (top to bottom)

1. **Open application** — default (bold) item;
2. **Print queue jobs: N** — live, greyed informational row (updates on every refresh);
3. **Open print queue** — queue UI of the default printer (first printer with jobs if no
   default), falls back to the printers folder;
4. **Open printer management** — classic Printers control panel (falls back to
   `ms-settings:printers`);
5. **Clear all jobs** — same guarded purge as the toolbar button (confirmation per
   settings);
6. **Pause spooler / Resume spooler** — label reflects current queue state;
7. **Restart spooler service** — stop + start with timeouts;
8. **Theme ▸** — System / Light / Dark with a check on the active mode;
9. **Start with Windows** — checked item, toggles the Run-key entry;
10. **Check for updates**;
11. **About**;
12. **Exit** — really exits (the X button closes to tray by default).

## Balloon notifications

Respect *Show notifications*: purge results (with removed-job count and backup path),
restore results, service restart result, update-available. Icons: info / warning / error.

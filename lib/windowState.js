// Saved window size, position and maximized state.
const { app, screen } = require('electron');
const path = require('path');
const fs = require('fs');

const WINDOW_STATE_FILE = path.join(app.getPath('userData'), 'window-state.json');

function loadWindowState() {
  try {
    const s = JSON.parse(fs.readFileSync(WINDOW_STATE_FILE, 'utf8'));
    if (!(s.width > 0 && s.height > 0)) return null;
    // Drops the position if under 100px of the window is on any display.
    const onScreen = Number.isFinite(s.x) && Number.isFinite(s.y) && screen.getAllDisplays().some(({ workArea: a }) =>
      Math.min(s.x + s.width, a.x + a.width) - Math.max(s.x, a.x) >= 100 &&
      Math.min(s.y + s.height, a.y + a.height) - Math.max(s.y, a.y) >= 100);
    if (!onScreen) { delete s.x; delete s.y; }
    return s;
  } catch { return null; }
}

function saveWindowState(win) {
  try {
    const { x, y, width, height } = win.getNormalBounds();
    fs.writeFileSync(WINDOW_STATE_FILE, JSON.stringify({ x, y, width, height, maximized: win.isMaximized() }));
  } catch {}
}

module.exports = { loadWindowState, saveWindowState };

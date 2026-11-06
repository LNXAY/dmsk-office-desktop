// DMSK Office als Windows-Programm: gleiche Oberfläche wie im Browser, bleibt angemeldet,
// läuft im Hintergrund (Infobereich) und zeigt Windows-Benachrichtigungen für neue Mails und Meldungen.
const { app, BrowserWindow, Tray, Menu, Notification, shell, session, nativeImage } = require('electron');
const fs = require('fs');
const { autoUpdater } = require('electron-updater');
const path = require('path');

const ORIGIN = 'https://office.dmsk-media.de';
const PARTITION = 'persist:dmsk-office';
const POLL_MS = 60 * 1000;

if (!app.requestSingleInstanceLock()) {
    app.quit();
    process.exit(0);
}
app.setAppUserModelId('de.dmsk-media.office');

const trayIcon = nativeImage.createFromPath(path.join(__dirname, 'icons', 'tray.png'));
const trayUnreadIcon = nativeImage.createFromPath(path.join(__dirname, 'icons', 'tray-unread.png'));
const badgeIcon = nativeImage.createFromPath(path.join(__dirname, 'icons', 'badge.png'));
let win = null;
let updateReady = false;
let tray = null;
let quitting = false;
let state = { notif: null, mail: null };
const stateFile = () => path.join(app.getPath('userData'), 'state.json');
const settingsFile = () => path.join(app.getPath('userData'), 'settings.json');

function loadJson(file, fallback) {
    try { return { ...fallback, ...JSON.parse(fs.readFileSync(file, 'utf8')) }; } catch { return fallback; }
}
function saveJson(file, data) {
    try { fs.writeFileSync(file, JSON.stringify(data)); } catch { /* egal */ }
}

function setupUpdates() {
    if (!app.isPackaged) return;
    autoUpdater.autoDownload = true;
    autoUpdater.autoInstallOnAppQuit = true;
    autoUpdater.disableDifferentialDownload = true;
    autoUpdater.on('update-downloaded', (info) => {
        notify('DMSK Office: Update bereit', `Version ${info.version} wird beim nächsten Start installiert.`, ORIGIN);
        updateReady = true;
        if (tray) tray.setContextMenu(buildMenu());
    });
    const check = () => { autoUpdater.checkForUpdates().catch(() => { /* offline oder noch kein Update */ }); };
    setTimeout(check, 20 * 1000);
    setInterval(check, 4 * 60 * 60 * 1000);
}

function isOffice(url) {
    try { return new URL(url).origin === ORIGIN; } catch { return false; }
}

function createWindow(hidden) {
    win = new BrowserWindow({
        width: 1360,
        height: 880,
        minWidth: 900,
        minHeight: 600,
        show: false,
        backgroundColor: '#0f172a',
        autoHideMenuBar: true,
        title: 'DMSK Office',
        icon: path.join(__dirname, 'icons', 'icon.png'),
        webPreferences: { partition: PARTITION, contextIsolation: true, sandbox: true, nodeIntegration: false },
    });
    win.removeMenu();
    win.webContents.setUserAgent(win.webContents.getUserAgent() + ' DMSKOfficeDesktop/' + app.getVersion());
    win.loadURL(ORIGIN);

    win.once('ready-to-show', () => { if (!hidden) win.show(); });
    win.webContents.setWindowOpenHandler(({ url }) => {
        if (isOffice(url)) return { action: 'allow' };
        if (/^https?:/i.test(url)) shell.openExternal(url);
        return { action: 'deny' };
    });
    win.webContents.on('will-navigate', (event, url) => {
        if (!isOffice(url) && !url.startsWith('file:')) {
            event.preventDefault();
            if (/^https?:/i.test(url)) shell.openExternal(url);
        }
    });
    win.webContents.on('did-fail-load', (_e, code, _desc, url, isMainFrame) => {
        if (isMainFrame && code !== -3) win.loadFile(path.join(__dirname, 'offline.html'), { query: { to: url || ORIGIN } });
    });
    win.on('close', (event) => {
        if (!quitting) {
            event.preventDefault();
            win.hide();
        }
    });
    win.on('page-title-updated', () => { /* Titel bleibt wie die Seite ihn setzt */ });
}

function showWindow(url) {
    if (!win) createWindow(false);
    if (url && isOffice(url)) win.loadURL(url);
    if (win.isMinimized()) win.restore();
    win.show();
    win.focus();
}

function buildMenu() {
    const settings = loadJson(settingsFile(), { autostart: true });
    return Menu.buildFromTemplate([
        { label: 'DMSK Office öffnen', click: () => showWindow() },
        { label: 'E-Mail öffnen', click: () => showWindow(ORIGIN + '/mail') },
        ...(updateReady ? [{ label: 'Update installieren und neu starten', click: () => { quitting = true; autoUpdater.quitAndInstall(); } }] : []),
        { type: 'separator' },
        {
            label: 'Mit Windows starten', type: 'checkbox', checked: app.getLoginItemSettings().openAtLogin,
            click: (item) => {
                app.setLoginItemSettings({ openAtLogin: item.checked, args: ['--hidden'] });
                saveJson(settingsFile(), { ...settings, autostart: item.checked });
            },
        },
        { type: 'separator' },
        { label: `Version ${app.getVersion()}`, enabled: false },
        { label: 'Beenden', click: () => { quitting = true; app.quit(); } },
    ]);
}

function createTray() {
    tray = new Tray(trayIcon);
    tray.setToolTip('DMSK Office');
    tray.setContextMenu(buildMenu());
    tray.on('click', () => showWindow());
    tray.on('double-click', () => showWindow());
}

async function poll() {
    try {
        const ses = session.fromPartition(PARTITION);
        const q = new URLSearchParams({ notif: String(state.notif ?? 0), mail: String(state.mail ?? 0) });
        const res = await ses.fetch(ORIGIN + '/desktop/neuigkeiten?' + q, { headers: { Accept: 'application/json' }, redirect: 'manual' });
        if (!res.ok || !(res.headers.get('content-type') || '').includes('json')) return; // nicht angemeldet
        const data = await res.json();

        if (state.notif === null || state.mail === null) {
            // erster Abruf: nur den Stand merken, nichts anzeigen
            state = { notif: data.latest_notification, mail: data.latest_mail };
            saveJson(stateFile(), state);
        } else {
            for (const m of data.mails) {
                notify('Neue E-Mail: ' + m.from, m.subject, m.url);
                state.mail = Math.max(state.mail, m.id);
            }
            for (const n of data.notifications) {
                notify(n.title, n.body || '', n.url);
                state.notif = Math.max(state.notif, n.id);
            }
            if (data.mails.length || data.notifications.length) saveJson(stateFile(), state);
        }
        const unread = Number(data.unread_mail) || 0;
        tray.setToolTip(unread > 0 ? `DMSK Office – ${unread} ungelesene E-Mail(s)` : 'DMSK Office');
        tray.setImage(unread > 0 ? trayUnreadIcon : trayIcon);
        if (win) win.setOverlayIcon(unread > 0 ? badgeIcon : null, unread > 0 ? `${unread} ungelesene E-Mail(s)` : '');
    } catch { /* offline: beim nächsten Mal wieder */ }
}

function notify(title, body, url) {
    if (!Notification.isSupported()) return;
    const n = new Notification({ title: String(title).slice(0, 120), body: String(body || '').slice(0, 240), icon: path.join(__dirname, 'icons', 'icon.png') });
    n.on('click', () => showWindow(url));
    n.show();
}

app.on('second-instance', () => showWindow());

app.whenReady().then(() => {
    const ses = session.fromPartition(PARTITION);
    ses.setPermissionRequestHandler((_wc, permission, callback) => callback(['notifications', 'clipboard-sanitized-write', 'fullscreen'].includes(permission)));

    const settings = loadJson(settingsFile(), { autostart: true, firstRun: true });
    if (settings.firstRun) {
        app.setLoginItemSettings({ openAtLogin: true, args: ['--hidden'] });
        saveJson(settingsFile(), { ...settings, firstRun: false });
    }
    state = loadJson(stateFile(), { notif: null, mail: null });

    createWindow(process.argv.includes('--hidden'));
    createTray();
    setupUpdates();
    setInterval(poll, POLL_MS);
    setTimeout(poll, 8000);
});

app.on('before-quit', () => { quitting = true; });
app.on('window-all-closed', () => { /* läuft im Infobereich weiter */ });

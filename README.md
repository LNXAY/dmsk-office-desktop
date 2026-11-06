# DMSK Office für Windows

Das Windows-Programm für DMSK Office: ein eigenes Fenster für `office.dmsk-media.de` mit dauerhafter Anmeldung,
Windows-Benachrichtigungen (neue Mails, Meldungen) und automatischen Updates.

Der Installer liegt unter `releases/`. Das Programm enthält keine Zugangsdaten und keine Geschäftsdaten:
es öffnet nur die Webseite und fragt als angemeldeter Benutzer nach Neuigkeiten.

Bauen: `npm install`, danach `npm run dist` (erzeugt `dist/`). Neue Version: `version` in `package.json` erhöhen,
bauen und `dist/DMSK-Office-Setup.exe` und `dist/latest.yml` nach `releases/` kopieren.

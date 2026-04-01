<!-- markdownlint-disable -->
# acadon Outlook Signature Add-in

Dieses Projekt enthält ein natives **Outlook Web Add-in** zur automatisierten Verwaltung und Bereitstellung von E-Mail-Signaturen, basierend auf der Office.js API. 

Es ermöglicht eine eigenständige, von den restriktiven Exchange-Signatur-APIs unabhängige Verwaltung von Signaturen über alle modernen Outlook-Clients hinweg (Web, Desktop, Mac).

## ✨ Kernfunktionen

- **Template-Verwaltung:** Vorlagen werden zentral von einem Webserver geladen.
- **Modulare Textbausteine:** Fixe Blöcke (Rechtliches, Logo), variable Platzhalter (z.B. `@@JobTitle@@`) und optionale Zusatz-Bausteine (Event-Banner etc.), die in der Taskpane an- oder abgewählt werden können.
- **Azure AD Integration:** Standardwerte (Telefon, Position, etc.) werden nahtlos via Microsoft Graph API aus dem Azure Active Directory bezogen.
- **Benutzerdefinierte Overrides:** Nutzer können ihre spezifischen Daten in der Taskpane überschreiben. Diese Anpassungen werden in den `roamingSettings` (Exchange-Postfach) gespeichert und über alle Geräte synchronisiert.
- **Automatisches Einfügen:** Die Signatur wird automatisch beim Verfassen einer neuen E-Mail (`OnNewMessageCompose`) sowie beim Antworten (`OnMessageReplyCompose`) eingefügt.

## 🔮 Ausblick & Roadmap

- **Absenderbasierte Zuweisung:** Automatische Auswahl der Signatur basierend auf der gewählten Absenderadresse (z.B. `.de` -> Deutsch, `.nl` -> Niederländisch).
- **Azure AD / Entra ID Sync:** Vollständige Automatisierung der Benutzerdaten (Position, Telefon) nach Abschluss der App-Registrierung.
- **Zentrale Administration:** Optionales Dashboard zur Verwaltung der Templates für alle User.

## 📂 Projektstruktur & Workflow

- **`outlook-addin/`**: Beinhaltet den gesamten Quellcode für das Office-Add-in (Manifest, Webpack-Config, `src/`, `templates/`, `blocks/` etc.).
- **`docs/`**: Enthält die produktiv nutzbaren Dateien. Dies wird in der Regel auf einem Webserver (z. B. GitHub Pages) gehostet.
- **Dazugehörige Scripts**: Mit der Datei `copy-to-docs.bat` werden die finalen Entwicklungs-Artefakte (src, templates, assets, addons, blocks) per Knopfdruck von `outlook-addin/` in den `docs/`-Ordner kopiert.

## 🚀 Lokale Entwicklung

Die aktive Entwicklung findet ausschließlich im Verzeichnis `outlook-addin` statt. Um das Add-in lokal zu testen und mit Hot-Reloading zu entwickeln:

### 1. Voraussetzungen installieren
Wechsle in den Add-in Ordner und installiere alle npm-Abhängigkeiten:
```bash
cd outlook-addin
npm install
```

### 2. Zertifikate für HTTPS generieren (Einmalig)
Outlook Add-ins **müssen** zwingend über HTTPS bereitgestellt werden. Generiere die lokalen Dev-Zertifikate:
```bash
npm run certs
```
*(Bestätige eventuelle Windows-Dialoge, die das Zertifikat installieren wollen.)*

### 3. Dev-Server starten
Starte den lokalen Webpack-Server. Dieser kompiliert den Code im `src/`-Ordner bei Änderungen mit Hot-Reloading neu.
```bash
npm run start
```
Der Server läuft nun unter `https://localhost:3000`.

### 4. Sideloading in Outlook (Manifest importieren)
Outlook muss wissen, woher das lokale Add-in geladen werden soll:

**Der schnellste Weg für Outlook im Web (OWA):**
1. Öffne im Browser: 👉 **[https://aka.ms/olksideload](https://aka.ms/olksideload)**
2. Es öffnet sich Outlook und das Dialogfenster "Add-Ins für Outlook".
3. Klicke links auf **Meine Add-Ins** (My add-ins).
4. Scrolle ganz nach unten zu **Benutzerdefinierte Addins** (Custom addins).
5. Klicke auf **Benutzerdefiniertes Add-in hinzufügen** -> **Aus Datei hinzufügen...** (Add from file...).
6. Wähle die Datei aus und bestätige.




<!-- markdownlint-disable -->
# Konzept: Outlook Web Add-in für die Signatur-Verwaltung

Dieses Dokument beschreibt eine alternative Architektur zur Browser-Erweiterung. Anstatt auf DOM-Ebene im Browser zu agieren, integriert sich diese Lösung als offizielles **Office Web Add-in** (.xml Manifest + lauffähige Web-App) nativ in Outlook (Web, Desktop, Mac und Mobile).

## 1. Zielsetzung
Eine eigenständige, von den restriktiven Exchange-Signatur-APIs unabhängige Verwaltung von Signaturen. 
- **Eigene Vorlagen**: Zentrale Bereitstellung von HTML-Templates über einen Webserver.
- **Dynamische Daten**: Abruf der Nutzerdaten direkt aus dem Azure Active Directory via Microsoft Graph API.
- **Benutzerdefinierte Anpassungen**: Nutzer können ihre Signatur anpassen (z. B. eine abweichende Handynummer oder einen persönlichen Gruß hinzufügen), ohne das Standard-Template zu zerstören.

## 2. Architektur & Funktionsweise

### 2.1 Event-Based Activation (On-Send oder On-Compose)
Das Add-in nutzt die **Event-Based Activation** von Office.js.
- **OnNewMessageCompose**: Sobald der Nutzer eine neue E-Mail schreibt, wird das Add-in unsichtbar im Hintergrund getriggert.
- *Wichtiger Hinweis zu Antworten (Replies):* Es gibt **kein** Event namens `OnMessageReplyCompose`! Ein solcher Eintrag in der `manifest.xml` führt zum Installationsfehler "300" in Outlook Web. Für Replies muss stattdessen `OnMessageCompose` (ab Mailbox 1.10) verwendet werden, falls automatische Einfügungen bei Antworten gewünscht sind.
- Generierung: Das Add-in führt eine Logik aus, generiert die Signatur und fügt sie über die API `Office.context.mailbox.item.body.setSignatureAsync(html, { coercionType: "Html" })` automatisch am Ende der E-Mail ein.

### 2.2 Task Pane (Benutzeroberfläche)
Das Add-in bietet optional eine Seitenleiste (Task Pane), über die der Nutzer interagieren kann:
- Vorschau der aktuellen Signatur.
- Auswahl verschiedener Vorlagen (z. B. "Intern", "Extern", "Kurz", "Lang").
- Eingabefelder für benutzerdefinierte Daten, die im AAD fehlen oder abweichen (z. B. eine spezifische Projekt-Telefonnummer).

## 3. Datenhaltung & Speicherorte im Add-in

Die Frage, wo Daten in einer Web-basierten Add-in Architektur gespeichert werden, ist zentral. Ein Office Add-in ist technisch eine gesicherte Webseite (iframe), die in Outlook läuft. Es gibt mehrere Speicherorte für unterschiedliche Zwecke:

### 3.1 Roaming Settings (Office.context.roamingSettings)
* **Zweck**: Speichern von Benutzereinstellungen (z. B. "Welche Signatur-Vorlage wurde standardmäßig gewählt?", "Hat der Nutzer einen individuellen Grußtext eingetippt?").
* **Eigenschaften**: Die Daten werden im Exchange-Postfach des Benutzers gespeichert und synchronisieren sich über alle Geräte (Outlook im Browser, Desktop-Client auf dem PC, der Mac-App).
* **Limitierungen**: Maximal 32 KB Speicherplatz pro Postfach. Nur für Konfigurations-Strings/JSONs geeignet, nicht für große Bilder.

### 3.2 Browser Storage (LocalStorage / IndexedDB)
* **Zweck**: Caching, um API-Aufrufe zu minimieren.
* **Eigenschaften**: Da das Add-in ein Browser-Fenster im Hintergrund öffnet, stehen die regulären `window.localStorage` oder `IndexedDB` zur Verfügung.
* **Beispiel**: Das Add-in lädt das HTML-Template vom Webserver und cacht es für 24 Stunden im Profil des lokalen Rechners. Ebenso können die Graph-API-Nutzerdaten für ein paar Stunden gecacht werden.

### 3.3 Backend / Webserver (Eigener Server)
* **Zweck**: Speicherung der zentralen Unternehmens-Templates.
* **Eigenschaften**: Ein eigener kleiner Webspace/Server stellt die reinen HTML-Dateien und Bilder bereit (z. B. unter `https://firma.de/signaturen/template_1.html`). Das Add-in ruft diese beim Start einfach via `fetch()` ab.

## 4. Ablauf-Skizze (Der Lebenszyklus)

1. **E-Mail wird verfasst**: Der Nutzer klickt auf "Neue E-Mail".
2. **Add-in Trigger**: Outlook weckt das Add-in-Skript (OnNewMessageCompose).
3. **Datenbeschaffung**:
    - Add-in nutzt *Single Sign-On (SSO)* (`Office.auth.getAccessToken()`), um stumm einen Token für die Microsoft Graph API zu erhalten.
    - Add-in fragt `/me` nach Vorname, Nachname, Telefon etc. an.
    - Parallel lädt das Add-in die HTML-Signatur als String vom eigenen Backend herunter (oder holt es aus dem LocalStorage).
4. **Benutzereinstellungen laden**: Das Add-in prüft die `roamingSettings`, ob der Nutzer z. B. in der Task Pane seine Telefonnummer manuell überschrieben hat.
5. **Platzhalter ersetzen**: Das Skript nimmt das HTML-Template und ersetzt `${Vorname}` mit den Graph- bzw. Nutzer-Daten.
6. **Einfügen**: `item.body.setSignatureAsync(...)` pflanzt die fertige HTML-Signatur direkt unten in das E-Mail-Fenster. Der Nutzer kann sofort lostippen.

## 5. Vorteile gegenüber der Browser-Extension

| Kriterium | Browser-Extension (DOM Manipulation) | Native Outlook Add-in (Office.js) |
| :--- | :--- | :--- |
| **Plattformen** | Nur in Chromium-basierten Browsern (Edge, Chrome). Gilt nur auf Webmail-Seite! | Alle Outllook Versionen: Web, Windows Desktop, Mac Desktop, teilweise Mobile! |
| **Zuverlässigkeit** | **Niedrig**. Ein Update der Microsoft Webseiten-UI killt das System sofort. | **Hoch**. Offizielle Microsoft API (`setSignatureAsync`), ändert sich quasi nie. |
| **Installation** | Muss per Chrome-Police oder manuell verteilt werden. | Kann vom Admin im M365 Admin Center per Knopfdruck ("Integrierten Apps") unternehmensweit für alle Rechner gepusht werden. |
| **Identity & APIs** | Umständliches Abgreifen von Cookies/Tokens aus dem Browser-Speicher. | Native, freigegebene SSO-Tokens auf Knopfdruck (`getAccessToken`). |
| **Trennung von MS** | Manipuliert unfreiwillig die Exchange-Einstellungen für Signaturen. | Legt niemals eine Signatur in den "Outlook-Optionen" an! Es stempelt die Signatur einfach zur Laufzeit live ins Fenster. |

## 6. Nötige Infrastruktur für diesen Ansatz

1. **Manifest.xml**: Die Konfigurationsdatei, die dem O365-Admin hochgeladen wird.
2. **Ein Webserver (z. B. Azure Static Web Apps, Vercel, oder ein IIS)**: Ein Add-in muss im Internet per **HTTPS** erreichbar sein. Es werden simple statische Dateien gehostet (`taskpane.html`, `commands.js`, `templates/`).
3. **App-Registrierung (Entra ID)**: Eine Azure AD App für das Add-in ist nötig, damit das Add-in die Berechtigung bekommt, Graph API Daten (`User.Read`) stumm abzurufen.

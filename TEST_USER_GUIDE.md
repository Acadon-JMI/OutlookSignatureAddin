# Quick Guide: acadon Signatur Add-in (Testphase)

Herzlich willkommen zum Test des neuen Outlook-Signatur-Add-ins! Dieses Add-in ermöglicht eine zentrale und dennoch flexible Verwaltung deiner E-Mail-Signaturen.

## 1. Installation

1. **Datei vorbereiten:** Du hast eine ZIP-Datei erhalten. Entpacke diese (sie enthält die `manifest.xml`).
2. **Sideload Link nutzen:** Öffne diesen Link in deinem Browser: 👉 **[https://aka.ms/olksideload](https://aka.ms/olksideload)**
   - Dieser Link öffnet direkt das Outlook-Dialogfenster "Add-Ins für Outlook".
3. **Importieren:**
   - Klicke im Dialog links auf **Meine Add-Ins** (My add-ins).
   - Scrolle ganz nach unten zu **Benutzerdefinierte Add-Ins** (Custom add-ins).
   - Klicke auf **Benutzerdefiniertes Add-in hinzufügen** -> **Aus Datei hinzufügen...** (Add from file...).
   - Wähle die entpackte `manifest.xml` aus und bestätige.

## 2. Wo finde ich das Add-in?

Das Add-in erscheint in allen Outlook-Versionen (Classic, New, Web) oben im Menüband unter dem **App-Menü** oder direkt als Button **"acadon Signatur -Beta"**.

## 3. Funktionen entdecken

- **Automatisches Einfügen:** 
  - Im **"New Outlook"** und Outlook Web wird die Signatur beim Erstellen einer Nachricht automatisch eingefügt.
  - Im **Outlook Classic** musst du beim Erstellen oder Antworten auf den Button **"acadon Signatur -Beta"** klicken. Die Seitenleiste öffnet sich dann und fügt die Signatur sofort automatisch ein (sofern der entsprechende Toggle aktiv ist).
- **Mehrfache Signaturen:** Du kannst verschiedene Signaturen (z.B. Deutsch/Englisch oder Lang/Kurz) erstellen und verwalten.
- **Block-Editor:** Deine Signatur besteht aus Bausteinen. Du kannst vordefinierte Blöcke anordnen oder komplett eigene **Custom HTML Blöcke** hinzufügen.
- **Zentrale Vorlagen:** Die Grund-Templates kommen direkt von GitHub. Wir können diese zentral aktualisieren, ohne dass deine individuellen Anpassungen verloren gehen.

## 4. Benutzerdaten

- **Aktueller Stand:** Momentan musst du deine **Position (Job Titel)** und **Telefonnummer** noch einmalig manuell in den "Benutzereinstellungen" im Add-in hinterlegen.
- **Ausblick:** Sobald die Azure App Registrierung finalisiert ist, werden diese Daten automatisch aus dem Firmenverzeichnis (Entra ID) geladen.

## 5. Technische Hintergrundinfos

- **Speicherort:** Deine Einstellungen und Signaturen werden im **Outlook Roaming Speicher** gesichert. Das bedeutet: Einmal eingestellt, sind sie auf all deinen Geräten (Laptop, PC, Web) verfügbar.
- **Update-Sicherheit:** Da die Logik zentral gehostet wird, erhältst du Funktions-Updates automatisch, ohne das Add-in neu installieren zu müssen.

---
Viel Erfolg beim Testen!

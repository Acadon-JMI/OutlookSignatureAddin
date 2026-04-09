# Anleitung: Microsoft Entra ID (Azure AD) für die acadon Signatur konfigurieren

> [!IMPORTANT]
> **Wichtiger Hinweis zur Repository-URL:**
> Die Konfiguration der Azure App-Registrierung und des Add-in-Manifests ist fest an die Hosting-URL gebunden (z. B. GitHub Pages). 
> Wenn das Repository verschoben oder umbenannt wird, ändert sich die Basis-URL. In diesem Fall müssen sowohl die **Redirect-URIs in Azure** als auch alle **URL-Einträge im `manifest.xml`** angepasst werden.

Damit das Outlook Add-in die **Position (Job Title)** und die **Telefonnummern** korrekt aus dem Verzeichnis abrufen kann, muss eine App-Registrierung in Microsoft Entra erstellt und im Add-in hinterlegt werden.

## 1. App-Registrierung erstellen

1. Melden Sie sich im [Microsoft Entra Admin Center](https://entra.microsoft.com/) oder im [Azure Portal](https://portal.azure.com/) an.
2. Navigieren Sie zu **Identität** > **Anwendungen** > **App-Registrierungen**.
3. Klicken Sie auf **Neue Registrierung**.
4. Geben Sie einen Namen ein (z. B. `acadon-outlook-signatur`).
5. Wählen Sie unter **Unterstützte Kontotypen** die Option: `Konten in einem beliebigen Organisationsverzeichnis (beliebiger Microsoft Entra-Mandant – mehrinstanzenfähig)`.
6. Klicken Sie auf **Registrieren**.

## 2. Authentifizierung konfigurieren

1. Wählen Sie im linken Menü **Authentifizierung**.
2. Klicken Sie auf **Plattform hinzufügen** > **Web**.
3. Geben Sie die **Redirect-URI** Ihrer Anwendung ein (die Basis-URL Ihres Add-ins), z. B.:
   `https://acadon-jmi.github.io/OutlookSignatureAddin/src/taskpane/taskpane.html`
4. Aktivieren Sie unter **Implizite Gewährung und Hybrid-Flows**:
   - [x] **Zugriffstoken**
   - [x] **ID-Token**
5. Klicken Sie auf **Speichern**.

## 3. Eine API verfügbar machen (Expose an API)

Dies ist für das Office SSO notwendig:
1. Wählen Sie **Eine API verfügbar machen**.
2. Klicken Sie oben bei **Anwendungs-ID-URI** auf **Festlegen**.
3. Ändern Sie `api://[GUID]` auf `api://acadon-jmi.github.io/OutlookSignatureAddin/[CLIENT_ID]`, wobei `[CLIENT_ID]` Ihre Application (client) ID ist.
4. Klicken Sie auf **Bereich hinzufügen** (Add a scope):
   - Bereichsname: `access_as_user`
   - Wer kann einwilligen: `Administratoren und Benutzer`
   - Anzeigename: `Access acadon Signature`
   - Beschreibung: `Allows the add-in to access the user profile.`
   - Zustand: `Aktiviert`
5. Autorisieren Sie die Office-Client-Anwendungen unter **Autorisierte Clientanwendungen**:
   Klicken Sie auf **Clientanwendung hinzufügen** und fügen Sie folgende IDs hinzu (Office-Standard-IDs):
   - `ea2850d5-d859-447a-9171-4127ee3d9516` (Microsoft Office)
   - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Office im Web)
   Wählen Sie den Bereich `api://.../access_as_user` aus.

## 4. Berechtigungen (API Permissions)

1. Wählen Sie **API-Berechtigungen**.
2. Stellen Sie sicher, dass folgende Berechtigungen vorhanden sind (Typ: Delegiert):
   - `User.Read`
   - `profile`
   - `openid`
3. Klicken Sie auf **Administratorzustimmung für [Ihr Tenant] erteilen**.

## 5. Manifest.xml aktualisieren

Öffnen Sie die Datei `outlook-addin/manifest.xml` und ersetzen Sie an zwei Stellen die Dummy-ID `00000000-0000-0000-0000-000000000000` durch Ihre neue **Application (client) ID**:

```xml
<WebApplicationInfo>
  <Id>[IHRE_CLIENT_ID]</Id>
  <Resource>api://acadon-jmi.github.io/OutlookSignatureAddin/[IHRE_CLIENT_ID]</Resource>
  <Scopes>
    <Scope>User.Read</Scope>
    <Scope>profile</Scope>
    <Scope>openid</Scope>
  </Scopes>
</WebApplicationInfo>
```

## Prüfung

Nachdem Sie das Add-in mit dem neuen Manifest neu geladen haben:
- Klicken Sie im Add-in auf "Laden" oder starten Sie es neu.
- Das Add-in sollte nun "Job Title" und "Phone" automatisch in die Felder "Anpassungen" laden.

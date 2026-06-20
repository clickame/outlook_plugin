# Clickame Outlook Assistant

Solució **dual** per a Outlook que afegeix:

1. **CCO/BCC automàtic** configurable a cada correu (en obrir-lo o just abans d'enviar-lo).
2. **Plantilles de correu** predefinides que s'insereixen al cos del missatge (estil "HubSpot templates", però **sense backend**).

Tot funciona **en local**: sense backend propi, **sense Microsoft Graph** i **sense enviar dades fora del PC**.

---

## Dos projectes, dues tecnologies

| | [`/outlook-classic-vsto`](outlook-classic-vsto) | [`/outlook-modern-officejs`](outlook-modern-officejs) |
|---|---|---|
| **Per a** | Outlook **clàssic** d'escriptori (Windows) | **Nou Outlook**, **Outlook web**, Outlook 365 modern |
| **Tecnologia** | C# · .NET Framework 4.8 · VSTO · WinForms | TypeScript · Office.js · HTML/CSS |
| **Distribució** | ClickOnce (`setup.exe` / `.vsto`) | Manifest + allotjament HTTPS (sideloading) |
| **Config local** | `%APPDATA%\ClickameOutlookAssistant\config.json` | `OfficeRuntime.storage` (fallback `localStorage`) |
| **CCO automàtic** | Interop COM (`Recipients`, `ItemSend`) | API de destinataris + Launch Events |
| **Plantilles** | Inserció a `HTMLBody`/`Body` | `setSelectedDataAsync` / `prependAsync` |

> **Per què dos projectes?** VSTO és la via més fiable per a l'Outlook clàssic però
> **no funciona al nou Outlook ni a Outlook web**. Office.js cobreix els clients moderns
> però **necessita manifest i allotjament HTTPS**. Tots dos comparteixen la mateixa
> estructura de configuració (`config.json`).

---

## Per on començar

- **Tens Outlook clàssic d'escriptori?** → [`outlook-classic-vsto/README.md`](outlook-classic-vsto/README.md)
- **Tens el nou Outlook o Outlook web?** → [`outlook-modern-officejs/README.md`](outlook-modern-officejs/README.md)
- **Exemple de configuració** → [`examples/config.example.json`](examples/config.example.json)

---

## Requisits funcionals comuns

- Configuració local: `email_bcc`, `activar_bcc_auto`, `mode_bcc` (`on_new_mail` | `on_send`), `plantilles[]`.
- Configurador per editar adreça CCO, activar/desactivar el CCO automàtic i fer CRUD de plantilles.
- Botó **Plantilles** en mode redacció.
- En seleccionar una plantilla: inserir el cos al cursor (o al final), i omplir l'assumpte si la plantilla en té i el correu no.
- **Evitar duplicar el BCC** si ja existeix.
- Logs locals bàsics per a depuració.

---

## Notes de compatibilitat i limitacions

- **VSTO no funciona al nou Outlook** (només Outlook clàssic d'escriptori). Per al nou Outlook, usa el projecte Office.js.
- **Office.js necessita manifest i allotjament HTTPS**; en desenvolupament s'usa `https://localhost:3000`. Cal *sideloading* o publicació.
- El **BCC automàtic pot dependre de les API disponibles** segons el client d'Outlook:
  - VSTO: fiable via `Application.ItemSend` (mode `on_send` recomanat).
  - Office.js: usa Launch Events; `OnMessageSend` requereix **Mailbox 1.12** i clients moderns. Si l'API de BCC no està disponible, en mode `on_send` **es bloqueja l'enviament amb un avís**.
- **No hi ha backend, ni Graph, ni sortida de dades del PC.** La configuració i les plantilles es desen localment.
- A Office.js, l'emmagatzematge local és **per màquina/perfil** i no se sincronitza entre dispositius.

---

## Estructura del repositori

```
.
├── outlook-classic-vsto/        # Projecte VSTO (Outlook clàssic)
├── outlook-modern-officejs/     # Projecte Office.js (nou Outlook / web)
├── examples/                    # config.example.json i documentació de camps
└── README.md
```

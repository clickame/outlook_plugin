# Clickame Outlook Assistant — Nou Outlook / Outlook web (Office.js)

Office Add-in fet amb **TypeScript + Office.js + HTML/CSS** per al **nou Outlook (Windows/Mac)**,
**Outlook web** i **Outlook 365** modern. Fa el mateix que el projecte germà VSTO:

1. **CCO/BCC automàtic** configurable (en obrir correu nou o just abans d'enviar).
2. **Plantilles** de correu predefinides que s'insereixen al cos.

**Sense backend**, **sense Microsoft Graph**, **sense enviar dades fora del PC**.
La configuració es desa amb `OfficeRuntime.storage` (amb *fallback* a `localStorage`).

---

## ⚠️ Limitacions importants

- Un Office Add-in **necessita un manifest** i **allotjament HTTPS** dels fitxers
  (en desenvolupament, `https://localhost:3000`). No és un simple `.exe`.
- El **BCC automàtic en composició** usa l'API de destinataris, que **depèn del client i del
  *requirement set* de Mailbox**. Si no està disponible, l'add-in ho detecta:
  - en mode `on_send` **bloqueja l'enviament amb un missatge** (Smart Alerts) perquè l'usuari ho resolgui;
  - en mode `on_new_mail` ho registra al log i continua.
- **`OnMessageSend`** (validació en enviar) requereix **Mailbox 1.12** i clients compatibles
  (nou Outlook, Outlook web, Outlook 365 recent). En clients antics, l'event no s'activa.
- `OfficeRuntime.storage` i `localStorage` són **per màquina/perfil**, no se sincronitzen entre dispositius.

---

## Requisits

- **Node.js 18+** i npm.
- Un compte d'Outlook (Microsoft 365 / Exchange / Outlook.com) per fer *sideloading*.
- **nou Outlook** o **Outlook web** per provar els esdeveniments d'inici.

---

## Posada en marxa (desenvolupament)

```bash
cd outlook-modern-officejs
npm install

# 1) Instal·la els certificats de desenvolupament HTTPS (una sola vegada)
npx office-addin-dev-certs install

# 2) Compila i arrenca el servidor local + sideloading automàtic
npm start
```

`npm start` (via `office-addin-debugging`) arrenca el *dev server* a `https://localhost:3000`
i intenta fer *sideload* del `manifest.xml` al teu Outlook. Si prefereixes fer-ho a mà,
mira la secció **Sideloading manual**.

Per només compilar:

```bash
npm run build        # producció -> dist/
npm run build:dev    # desenvolupament
npm run watch        # recompila en desar
npm run validate     # valida el manifest.xml
```

---

## Sideloading manual

### Nou Outlook / Outlook web
1. Obre Outlook web (`https://outlook.office.com`) o el **nou Outlook**.
2. **Configuració** ⚙️ → *Correu* → **Personalitzar accions** o bé el menú
   **Obtenir complements** / **Get Add-ins**.
3. A la finestra de complements: **Els meus complements** → **Complements personalitzats**
   → **Afegeix un complement personalitzat** → **Afegeix des d'un fitxer…**
4. Selecciona `manifest.xml` d'aquest projecte.
5. Accepta l'avís de complement no verificat (és per a ús intern).

> El *dev server* (`https://localhost:3000`) ha d'estar en marxa (`npm run dev-server` o `npm start`).

### Outlook clàssic Windows (opcional)
El nou model d'esdeveniments funciona millor al nou Outlook/web. Per a Outlook clàssic,
és preferible el projecte [`/outlook-classic-vsto`](../outlook-classic-vsto).

---

## Ús

### Configuració del CCO
1. En un correu nou, obre la pestanya amb el grup **Clickame Assistant** → **Configuració**.
2. Introdueix l'**adreça CCO**, marca **Activar CCO automàtic** i tria el **mode**:
   - **on_send** (recomanat): es valida i s'afegeix abans d'enviar.
   - **on_new_mail**: s'afegeix en obrir el correu.
3. **Desar configuració**.

### Plantilles
1. Pestanya **Plantilles** del panell.
2. **Nova** / edita els camps (nom, assumpte opcional, cos HTML o text) → **Desar plantilla**.
3. Amb un correu obert, selecciona una plantilla i **Inserir al correu**:
   - s'insereix al **punt del cursor** (`setSelectedDataAsync`); si no es pot, al principi (`prependAsync`);
   - si la plantilla té assumpte i el correu no en té, **s'omple l'assumpte**.

### Logs
Pestanya **Logs**: mostra els últims missatges de depuració desats localment (`localStorage`).

---

## Publicació (producció)

1. Allotja el contingut de `dist/` en un servidor **HTTPS** (Azure Static Web Apps, GitHub Pages amb HTTPS, etc.).
2. Edita `manifest.xml`:
   - genera un **GUID nou** a `<Id>`,
   - substitueix totes les URL `https://localhost:3000` per la teva URL pública,
   - afegeix les **icones** reals a `assets/` (mira `assets/README.md`).
3. Distribueix el `manifest.xml`:
   - per a una organització: **Centre d'administració de Microsoft 365 → Paràmetres → Aplicacions integrades / complements**;
   - o publicació a AppSource per a distribució pública.

---

## Estructura

```
outlook-modern-officejs/
  src/
    taskpane.html / taskpane.css / taskpane.ts   # UI: CCO, plantilles, logs
    commands.html / commands.ts                  # function file (accions + import d'events)
    events.ts                                     # handlers OnNewMessageCompose / OnMessageSend
    configService.ts                              # OfficeRuntime.storage + fallback localStorage
    templateService.ts                            # CRUD + inserció al cos
    bccService.ts                                 # afegir BCC sense duplicats
    logger.ts                                     # logs locals
    models.ts                                     # tipus i config per defecte
  assets/                                         # icones (afegeix-les)
  manifest.xml
  package.json
  tsconfig.json
  webpack.config.js
  README.md
```

Exemple de configuració a [`../examples/config.example.json`](../examples/config.example.json).

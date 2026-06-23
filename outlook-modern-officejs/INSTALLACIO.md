# Instal·lació — Clickame Outlook Assistant (Nou Outlook / Outlook web / Microsoft 365)

Aquesta és la versió **Office.js** del complement. Funciona al **nou Outlook per Windows**, a
**Outlook al web** (navegador) i a **Outlook de Microsoft 365** (Windows i Mac).

> ❌ Aquesta versió **NO** funciona a l'Outlook clàssic d'escriptori antic sense connexió a Office 365.
> Per a l'Outlook clàssic, fes servir la versió VSTO: [`../outlook-classic-vsto`](../outlook-classic-vsto).

El codi es publica automàticament a **GitHub Pages** i el manifest apunta a:

```
https://clickame.github.io/outlook_plugin/manifest.xml
```

No cal instal·lar res al PC: el complement carrega els fitxers des d'aquesta URL.

---

## ⚠️ Requisit important per al CCO automàtic

El CCO automàtic en enviar (Smart Alerts, `OnMessageSend`) i en obrir correu nou
(`OnNewMessageCompose`) són **esdeveniments d'inici**. Perquè s'executin de manera
fiable per a tota l'organització, el complement **s'ha de desplegar des del Centre
d'administració de Microsoft 365** (mètode B, més avall). En instal·lació individual
(mètode A) funciona per a proves al nou Outlook, però Microsoft recomana el desplegament
centralitzat per als esdeveniments.

Requereix **Mailbox 1.12** (nou Outlook, Outlook web i Microsoft 365 actuals el compleixen).

---

## Mètode A — Instal·lació individual (un sol usuari, per a proves)

La instal·lació personalitzada es fa des d'**Outlook al web** (val per al nou Outlook,
perquè comparteixen la mateixa bústia).

1. Obre **Outlook al web**: <https://outlook.office.com/mail/>
2. A la cinta, obre **Obtenir complements** (icona de complements / *Get Add-ins*).
   - Si no la veus: menú **…** d'un correu obert → **Obtenir complements**.
3. Ves a **Els meus complements** (*My add-ins*).
4. Baixa fins a **Complements personalitzats** (*Custom Addins*) → **Afegeix un complement personalitzat** → **Afegeix des d'una URL** (*Add from URL*).
5. Enganxa la URL del manifest:
   ```
   https://clickame.github.io/outlook_plugin/manifest.xml
   ```
6. **Instal·la** / **Accepta**.
7. **Reinicia** el nou Outlook (tanca'l del tot i torna'l a obrir).

---

## Mètode B — Desplegament centralitzat (recomanat, per a tota l'organització)

Aquest mètode el fa un **administrador de Microsoft 365** i és el recomanat perquè els
esdeveniments (`OnMessageSend`) funcionin de manera garantida.

1. Entra al **Centre d'administració de Microsoft 365**: <https://admin.microsoft.com/>
2. Ves a **Configuració → Aplicacions integrades** (*Settings → Integrated apps*).
3. **Carrega aplicacions personalitzades** (*Upload custom apps*).
4. Tria **Proporcionar un enllaç al fitxer de manifest** i enganxa:
   ```
   https://clickame.github.io/outlook_plugin/manifest.xml
   ```
5. Assigna els **usuaris** (tu, un grup, o tota l'organització).
6. **Accepta els permisos** (`ReadWriteItem`) i finalitza.
7. El desplegament pot trigar **fins a 24 h** a propagar-se (sovint molt menys).

---

## Primer ús (configurar el CCO)

1. Obre un **correu nou**.
2. A la cinta, grup **Clickame Assistant** → botó **Plantilles/Configuració** (obre el panell).
3. Pestanya **CCO**:
   - **Adreça CCO (BCC)**: l'adreça que vols en còpia oculta.
   - ✅ **Activar CCO automàtic**.
   - **Mode**:
     - **En enviar (on_send)** — el CCO **NO** es veu al camp mentre escrius; s'afegeix
       de manera invisible **en enviar**. Comprova-ho a la carpeta **Enviats**.
     - **Al crear correu nou (on_new_mail)** — el CCO **apareix visible** al camp CCO en
       obrir el correu (bo per verificar que funciona).
   - **Desar configuració**.

> 💡 Per comprovar ràpidament que funciona, posa el mode **"Al crear correu nou"**: en obrir
> un correu nou hauries de veure el CCO aparèixer sol al camp CCO.

---

## Actualitzar a una versió nova

- Si **només canvia el codi** (JS/HTML): no cal fer res a Microsoft. N'hi ha prou amb
  **reiniciar l'Outlook** perquè agafi els fitxers nous de GitHub Pages (pot caldre esperar
  uns minuts o esborrar la memòria cau).
- Si **canvia el `manifest.xml`** (URLs, permisos, esdeveniments o el número de `<Version>`):
  - **Mètode A:** treu el complement i torna'l a afegir des de la URL.
  - **Mètode B:** al Centre d'administració, **actualitza** l'aplicació integrada.

Per forçar refrescos fiables, cada desplegament puja el número de `<Version>` del manifest.

---

## Resolució de problemes

| Símptoma | Causa probable | Solució |
|---|---|---|
| El CCO no apareix al camp en redactar | Mode = `on_send` (és normal) | Mira els **Enviats**, o canvia a `on_new_mail` per veure'l |
| No s'afegeix mai el CCO, ni als Enviats | El manifest instal·lat és antic (sense esdeveniments) o codi antic en cau | Re-instal·la el manifest (v1.0.1.0+) i **reinicia** l'Outlook |
| El panell s'obre però la config no es desa entre dispositius | Cau d'Outlook | Reinicia; la config es desa a `roamingSettings` (bústia) |
| No surt el botó a la cinta | Complement no propagat encara | Espera (mètode B pot trigar) i reinicia |

Per veure activitat interna, pestanya **Logs** del panell.

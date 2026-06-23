# Instal·lació — Clickame Outlook Assistant (Outlook clàssic d'escriptori / VSTO)

Aquesta és la versió **VSTO** del complement, per a l'**Outlook clàssic d'escriptori a Windows**
(Outlook 2016 / 2019 / 2021 / 365 "classic"). Fa CCO automàtic i plantilles, **100% en local**.

> ❌ Aquesta versió **NO** funciona al **nou Outlook** ni a **Outlook web**.
> Per a aquests, fes servir la versió moderna: [`../outlook-modern-officejs`](../outlook-modern-officejs).

Hi ha dues maneres d'instal·lar-la:
- **Mètode A** — Usuari final amb un paquet ja compilat (ClickOnce). *(Recomanat per distribuir.)*
- **Mètode B** — Des de Visual Studio (desenvolupament / prova).

---

## Requisits al PC de destí

- **Windows** amb **Outlook clàssic d'escriptori** instal·lat (no la versió "nou Outlook" de la Store).
- **.NET Framework 4.8** (normalment ja present a Windows 10/11 actuals).
- **Visual Studio 2010 Tools for Office Runtime (VSTO Runtime)** — sol venir amb Office;
  si no, es pot instal·lar gratis des de Microsoft.

---

## Mètode A — Instal·lació per a usuari final (ClickOnce, sense Visual Studio)

### 1) Generar el paquet (ho fa qui compila, un sol cop)

A la màquina de desenvolupament, amb Visual Studio:

1. Obre `ClickameOutlookAssistant.sln`.
2. Posa la configuració en **Release**.
3. Clic dret al projecte **ClickameOutlookAssistant → Publish…**.
4. Tria una carpeta de sortida (una carpeta local, una unitat de xarxa compartida, o web).
5. Es generen:
   - `setup.exe`
   - `ClickameOutlookAssistant.vsto`
   - carpeta `Application Files\`
6. *(Recomanat)* A **Project Properties → Signing**, signa el manifest amb un
   **certificat de confiança** de l'organització per evitar avisos de seguretat.

### 2) Instal·lar al PC de l'usuari

1. **Tanca l'Outlook** completament.
2. Copia al PC (o accedeix a la carpeta de xarxa) els fitxers publicats:
   `setup.exe`, `ClickameOutlookAssistant.vsto` i la carpeta `Application Files\`
   (els tres han d'anar **junts**).
3. Executa **`setup.exe`** (o fes doble clic a `ClickameOutlookAssistant.vsto`).
   - S'instal·la **per usuari**, **sense permisos d'administrador**.
   - Si el `.vsto` està en una ubicació no signada/no confiable, ClickOnce pot demanar
     confirmació; signar amb certificat ho evita.
4. Obre **Outlook**. A la pestanya **Missatge** d'un correu nou hi haurà el grup
   **"Clickame Assistant"**.

### Desinstal·lar (Mètode A)

**Tauler de control → Programes i característiques** → busca *ClickameOutlookAssistant* → **Desinstal·la**.

---

## Mètode B — Des de Visual Studio (desenvolupament i prova)

1. Instal·la **Visual Studio 2019/2022** amb la càrrega de treball
   **"Desenvolupament d'Office/SharePoint"** (inclou VSTO) i el **.NET Framework 4.8 Developer Pack**.
2. Obre `ClickameOutlookAssistant.sln`.
3. Si demana restaurar paquets **NuGet**, accepta (instal·la `Newtonsoft.Json 13.0.3`).
4. Comprova que es resolen les referències a `Microsoft.Office.Interop.Outlook` i
   `Microsoft.Office.Tools.Outlook`.
5. Prem **F5**. Visual Studio compila, **registra l'add-in només per a l'usuari actual**
   (clau a `HKCU`) i obre Outlook amb el complement carregat.
6. En aturar la depuració, VSTO **desregistra** l'add-in automàticament.

---

## Primer ús (configurar el CCO)

1. A Outlook, obre un **correu nou**.
2. Pestanya **Missatge** → grup **Clickame Assistant** → botó **Configuració**.
3. Introdueix l'**adreça CCO (BCC)** i marca **Activar CCO automàtic**.
4. Tria el **mode**:
   - **En enviar (recomanat)** — `on_send`: el BCC es garanteix a `Application.ItemSend`,
     just abans d'enviar. És el més fiable.
   - **En obrir correu nou** — `on_new_mail`: el BCC s'afegeix en obrir el compositor
     (l'usuari el veu i el pot esborrar; el mode `on_send` el torna a validar).
5. **Desar**.

La configuració i els logs es desen a:

```
%APPDATA%\ClickameOutlookAssistant\config.json
%APPDATA%\ClickameOutlookAssistant\log.txt
```

---

## Plantilles

- **Configuració → Gestionar plantilles…** o menú **Plantilles → Gestionar plantilles…**.
- Cada plantilla té **nom**, **assumpte opcional** i **cos** (HTML o text pla).
- En un correu en composició, obre **Plantilles** i tria'n una:
  - s'insereix el cos al **punt del cursor** (o al final si no es pot detectar),
  - si la plantilla té assumpte i el correu encara no en té, **s'omple l'assumpte**.

---

## Resolució de problemes

| Símptoma | Causa probable | Solució |
|---|---|---|
| No surt el grup "Clickame Assistant" a la cinta | Add-in desactivat per Outlook | **Fitxer → Opcions → Complements → Complements COM** → activa'l. Mira també *Elements deshabilitats* |
| L'add-in desapareix en reiniciar | S'ha carregat amb F5 (mode dev) | Per ús permanent, instal·la amb el **Mètode A (ClickOnce)** |
| Error en obrir, demana VSTO Runtime | Falta el runtime VSTO | Instal·la *Visual Studio 2010 Tools for Office Runtime* |
| No s'afegeix el BCC | Polítiques de TI bloquegen interop COM | Revisa restriccions de l'organització; prova el mode `on_send` |
| Avís de seguretat en instal·lar | `.vsto` no signat | Signa el manifest amb un certificat de confiança |

---

## Estructura del projecte

Vegeu [`README.md`](README.md) per al detall de fitxers i arquitectura.

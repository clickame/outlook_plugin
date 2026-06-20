# Clickame Outlook Assistant — Outlook Clàssic (VSTO)

Complement **VSTO** per a **Outlook clàssic d'escriptori a Windows** (Outlook 2016/2019/2021/365 "classic").
Fa dues coses:

1. **CCO/BCC automàtic** configurable a cada correu nou o just abans d'enviar.
2. **Plantilles de correu** predefinides que s'insereixen al cos del missatge des d'un botó a la cinta.

Tot funciona **100% en local**. No hi ha backend, no s'usa Microsoft Graph i **no surt cap dada del PC**.

---

## ⚠️ Limitacions importants

- **VSTO NO funciona al "nou Outlook" ni a Outlook web.** Per a aquests, fes servir el projecte germà [`/outlook-modern-officejs`](../outlook-modern-officejs).
- Cal **Outlook d'escriptori clàssic** instal·lat (no la versió Store "nou Outlook").
- El BCC automàtic depèn de les API d'Outlook (interop COM). En entorns molt restringits per polítiques de TI, l'add-in pot estar bloquejat.
- La inserció al **punt del cursor** funciona quan l'editor és Word (cas normal en correus HTML). Si no es pot detectar la selecció, la plantilla s'afegeix **al final** del cos.

---

## Requisits de desenvolupament

- **Visual Studio 2019 o 2022** amb la càrrega de treball **"Desenvolupament d'Office/SharePoint"** (inclou els *Visual Studio Tools for Office / VSTO*).
- **.NET Framework 4.8** (Developer Pack).
- Outlook d'escriptori instal·lat a la màquina de desenvolupament.
- Connexió a NuGet per restaurar **Newtonsoft.Json 13.0.3**.

---

## Compilar i provar (F5)

1. Obre `ClickameOutlookAssistant.sln` amb Visual Studio.
2. Si demana restaurar paquets NuGet, accepta (instal·la `Newtonsoft.Json`).
3. Comprova que les referències a `Microsoft.Office.Interop.Outlook` i `Microsoft.Office.Tools.Outlook` es resolen (s'instal·len amb la càrrega d'Office).
4. Prem **F5**. Visual Studio:
   - compila l'add-in,
   - el registra **només per a l'usuari actual** (clau de registre a `HKCU`),
   - obre Outlook amb l'add-in carregat.
5. A Outlook, crea un **correu nou**. Veuràs el grup **"Clickame Assistant"** a la pestanya *Missatge* amb:
   - **Configuració**
   - **Plantilles** (menú desplegable)

> En aturar la depuració, VSTO **desregistra** l'add-in automàticament.

---

## Configuració (primer ús)

1. Botó **Configuració**.
2. Introdueix l'**adreça CCO (BCC)**.
3. Marca **Activar CCO automàtic**.
4. Tria el **mode**:
   - **En enviar (recomanat)** — `on_send`: el BCC es garanteix a `Application.ItemSend`, just abans d'enviar. És el més fiable.
   - **En obrir correu nou** — `on_new_mail`: el BCC s'afegeix en obrir el compositor. L'usuari el pot veure (i esborrar). El mode `on_send` torna a validar-lo igualment.
5. **Desar**.

La configuració es desa a:

```
%APPDATA%\ClickameOutlookAssistant\config.json
```

Els logs de depuració:

```
%APPDATA%\ClickameOutlookAssistant\log.txt
```

(Pots obrir aquesta carpeta amb l'enllaç de la finestra de Configuració.)

---

## Plantilles

- **Configuració → Gestionar plantilles…** o el menú **Plantilles → Gestionar plantilles…**.
- Cada plantilla té **nom**, **assumpte opcional** i **cos** (HTML o text pla, segons la casella "El cos és HTML").
- En un correu en composició, obre **Plantilles** i tria'n una:
  - s'insereix el cos al **punt del cursor** (o al final si no es pot detectar),
  - si la plantilla té assumpte i el correu encara no en té, **s'omple l'assumpte**.

---

## Instal·lació en un PC d'usuari final (sense Visual Studio)

VSTO es distribueix amb **ClickOnce**. Després de compilar en **Release**:

1. A Visual Studio: clic dret al projecte → **Publish…**.
2. Tria una carpeta de sortida (p. ex. una xarxa compartida o una carpeta local).
3. Es generen, entre d'altres:
   - `ClickameOutlookAssistant.vsto`
   - `setup.exe`
   - carpeta `Application Files\`
4. Al PC de destí:
   - Cal el **runtime de VSTO** (*Microsoft Visual Studio 2010 Tools for Office Runtime*), normalment ja present amb Office; si no, instal·la'l.
   - Executa `setup.exe` (o obre el `.vsto`). S'instal·la **per usuari**, sense permisos d'administrador.
5. Per a desplegaments signats i sense avisos de seguretat, signa el manifest amb un **certificat de confiança** (a *Project Properties → Signing*).

> Si mous el `.vsto` a una ubicació no confiable, ClickOnce pot demanar confirmació. Signar amb un certificat de l'organització ho evita.

---

## Estructura

```
outlook-classic-vsto/
  ClickameOutlookAssistant.sln
  ClickameOutlookAssistant/
    ThisAddIn.cs                  # Esdeveniments: NewInspector i ItemSend
    ThisAddIn.Designer.cs         # Codi VSTO (Globals, Application tipat)
    Ribbon.cs                     # Cinta: Configuració + menú dinàmic de Plantilles
    Ribbon.xml                    # Definició de la cinta (EmbeddedResource)
    ConfigService.cs              # Llegir/desar config.json
    TemplateService.cs            # CRUD de plantilles + inserció al MailItem
    BccService.cs                 # Afegir BCC evitant duplicats
    Logger.cs                     # Logs locals bàsics
    Models/Config.cs
    Models/EmailTemplate.cs
    Forms/SettingsForm.cs         # WinForms: adreça CCO, auto, mode
    Forms/TemplatesForm.cs        # WinForms: crear/editar/eliminar plantilles
    app.manifest
    packages.config
    Properties/AssemblyInfo.cs
  README.md
```

Exemple de `config.json` a [`../examples/config.example.json`](../examples/config.example.json).

# Exemples de configuració

`config.example.json` és un exemple de la configuració local compartida pels dos projectes.

## On va cada còpia

### Outlook clàssic (VSTO)
Copia el contingut a:

```
%APPDATA%\ClickameOutlookAssistant\config.json
```

(Pots obrir aquesta carpeta des de **Configuració → Obrir carpeta de configuració/logs** dins de l'add-in.)

### Nou Outlook / Outlook web (Office.js)
No es desa en un fitxer accessible directament: es guarda dins de
`OfficeRuntime.storage` (o `localStorage`) amb la clau `clickame_config`.
La manera recomanada de carregar-lo és **a través del task pane** (pestanyes CCO i Plantilles).
L'estructura JSON, però, és idèntica a aquest exemple.

## Camps

| Camp                | Tipus     | Descripció                                                        |
|---------------------|-----------|------------------------------------------------------------------|
| `email_bcc`         | string    | Adreça que s'afegeix com a CCO/BCC.                              |
| `activar_bcc_auto`  | boolean   | Activa/desactiva el CCO automàtic.                              |
| `mode_bcc`          | string    | `"on_new_mail"` o `"on_send"`.                                  |
| `plantilles`        | array     | Llista de plantilles.                                            |
| `plantilles[].id`   | string    | Identificador estable.                                           |
| `plantilles[].nom`  | string    | Nom visible.                                                     |
| `plantilles[].assumpte` | string (opcional) | Si té valor i el correu no té assumpte, s'omple.    |
| `plantilles[].cos`  | string    | Cos HTML o text pla.                                             |
| `plantilles[].es_html` | boolean | Indica si `cos` és HTML.                                       |

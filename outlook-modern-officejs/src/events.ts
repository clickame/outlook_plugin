// Handlers d'esdeveniments d'inici (launch events) declarats al manifest:
//  - onNewMessageCompose  -> afegeix BCC si el mode és on_new_mail
//  - onMessageSendHandler -> garanteix BCC abans d'enviar (on_send), o avisa
//
// Aquests handlers s'executen en un runtime sense interfície. Eviten dependre de
// localStorage per a la lectura de config: usen OfficeRuntime.storage quan cal.

import { loadConfig } from "./configService";
import { ensureBcc, canEditBcc } from "./bccService";
import { Logger } from "./logger";

/* global Office */

/** S'executa en obrir un correu nou en composició. */
async function onNewMessageCompose(event: Office.AddinCommands.Event): Promise<void> {
  try {
    const cfg = await loadConfig();
    if (cfg.activar_bcc_auto && cfg.mode_bcc === "on_new_mail") {
      await ensureBcc(cfg);
    }
  } catch (e) {
    Logger.error("Error a onNewMessageCompose.", e);
  } finally {
    event.completed();
  }
}

/**
 * S'executa just abans d'enviar (smart alerts / on-send).
 * - Si pot afegir el BCC, el garanteix i deixa enviar.
 * - Si el BCC està activat però no es pot afegir, bloqueja amb un missatge.
 */
async function onMessageSendHandler(event: Office.AddinCommands.Event): Promise<void> {
  try {
    const cfg = await loadConfig();

    if (!cfg.activar_bcc_auto || !(cfg.email_bcc || "").trim()) {
      // Res a fer: deixa enviar.
      (event as Office.MailboxEvent).completed({ allowEvent: true });
      return;
    }

    if (!canEditBcc()) {
      // No podem garantir el BCC en aquest client: bloquegem i informem.
      (event as Office.MailboxEvent).completed({
        allowEvent: false,
        errorMessage:
          "No s'ha pogut afegir el CCO automàtic en aquest client d'Outlook. " +
          "Afegeix-lo manualment o desactiva el CCO automàtic a la configuració.",
      });
      return;
    }

    const result = await ensureBcc(cfg);
    if (result.added) {
      (event as Office.MailboxEvent).completed({ allowEvent: true });
    } else {
      (event as Office.MailboxEvent).completed({
        allowEvent: false,
        errorMessage:
          "No s'ha pogut afegir el CCO automàtic (" +
          (result.reason || "desconegut") +
          "). Revisa la configuració.",
      });
    }
  } catch (e) {
    Logger.error("Error a onMessageSendHandler.", e);
    // Davant d'un error inesperat, no bloquegem l'enviament de l'usuari.
    (event as Office.MailboxEvent).completed({ allowEvent: true });
  }
}

// Registre dels handlers perquè el runtime els pugui invocar pel nom del manifest.
Office.actions.associate("onNewMessageCompose", onNewMessageCompose);
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);

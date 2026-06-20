// Lògica de BCC automàtic per a Office.js, evitant duplicats.
// Nota: la disponibilitat de l'API de destinataris BCC en composició depèn del
// client i del "requirement set" de Mailbox. Es comprova abans d'usar-la.

import { Config } from "./models";
import { Logger } from "./logger";

/** Comprova si podem manipular el BCC en aquest client. */
export function canEditBcc(): boolean {
  try {
    const item = Office.context.mailbox.item as Office.MessageCompose;
    return !!item && !!item.bcc && typeof item.bcc.getAsync === "function";
  } catch {
    return false;
  }
}

/**
 * Garanteix que el BCC configurat hi és (si està activat i no duplicat).
 * Retorna un resultat amb informació per a missatges a l'usuari / on-send.
 */
export function ensureBcc(cfg: Config): Promise<{ added: boolean; reason?: string }> {
  return new Promise((resolve) => {
    if (!cfg.activar_bcc_auto) {
      resolve({ added: false, reason: "desactivat" });
      return;
    }
    const target = (cfg.email_bcc || "").trim();
    if (!target) {
      resolve({ added: false, reason: "sense_adreca" });
      return;
    }
    if (!canEditBcc()) {
      Logger.warn("Aquest client no permet editar el BCC en composició.");
      resolve({ added: false, reason: "api_no_disponible" });
      return;
    }

    const item = Office.context.mailbox.item as Office.MessageCompose;
    item.bcc.getAsync((getRes) => {
      if (getRes.status !== Office.AsyncResultStatus.Succeeded) {
        Logger.error("No s'ha pogut llegir el BCC actual.", getRes.error);
        resolve({ added: false, reason: "error_lectura" });
        return;
      }

      const existing = getRes.value || [];
      const already = existing.some(
        (r) => (r.emailAddress || "").toLowerCase() === target.toLowerCase()
      );
      if (already) {
        Logger.info("El BCC ja existeix; no es duplica.");
        resolve({ added: true });
        return;
      }

      // addAsync afegeix sense esborrar els existents (accepta adreces SMTP com a string).
      item.bcc.addAsync([target], (addRes) => {
        if (addRes.status === Office.AsyncResultStatus.Succeeded) {
          Logger.info(`BCC afegit: ${target}`);
          resolve({ added: true });
        } else {
          Logger.error("No s'ha pogut afegir el BCC.", addRes.error);
          resolve({ added: false, reason: "error_addicio" });
        }
      });
    });
  });
}

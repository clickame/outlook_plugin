// Function file de comandes. Allotja accions del ribbon i, per import, els
// handlers d'esdeveniments (events.ts) perquè quedin registrats al runtime.

import "./events";
import { Logger } from "./logger";

/* global Office */

Office.onReady(() => {
  Logger.info("Commands runtime carregat.");
});

/**
 * Acció opcional del ribbon. El botó "Plantilles" del manifest usa ShowTaskpane,
 * però deixem aquesta funció com a exemple d'acció executable.
 */
function showTemplatesNotification(event: Office.AddinCommands.Event): void {
  try {
    Office.context.mailbox.item?.notificationMessages.addAsync("clickameInfo", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: "Obre el panell de Clickame per triar una plantilla.",
      icon: "Icon.16x16",
      persistent: false,
    });
  } catch (e) {
    Logger.error("Error mostrant notificació.", e);
  } finally {
    event.completed();
  }
}

// Registre per si s'assigna a un botó al manifest.
Office.actions.associate("showTemplatesNotification", showTemplatesNotification);

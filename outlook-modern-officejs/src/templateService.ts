// CRUD de plantilles i inserció al cos del correu en composició.

import { Config, EmailTemplate, newId } from "./models";
import { loadConfig, saveConfig } from "./configService";
import { Logger } from "./logger";

export async function getTemplates(): Promise<EmailTemplate[]> {
  const cfg = await loadConfig();
  return cfg.plantilles ?? [];
}

export async function upsertTemplate(t: EmailTemplate): Promise<Config> {
  const cfg = await loadConfig();
  if (!t.id) t.id = newId();
  const idx = cfg.plantilles.findIndex((x) => x.id === t.id);
  if (idx >= 0) cfg.plantilles[idx] = t;
  else cfg.plantilles.push(t);
  await saveConfig(cfg);
  Logger.info(`Plantilla desada: ${t.nom}`);
  return cfg;
}

export async function deleteTemplate(id: string): Promise<Config> {
  const cfg = await loadConfig();
  cfg.plantilles = cfg.plantilles.filter((x) => x.id !== id);
  await saveConfig(cfg);
  Logger.info(`Plantilla eliminada: ${id}`);
  return cfg;
}

/**
 * Insereix una plantilla al correu en composició.
 * - Intenta inserir al punt del cursor amb setSelectedDataAsync.
 * - Si falla (p. ex. sense focus al cos), fa fallback a prependAsync.
 * - Si la plantilla té assumpte i el correu no en té, l'omple.
 */
export function insertTemplateIntoMail(t: EmailTemplate): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    const item = Office.context.mailbox.item;
    if (!item || !item.body) {
      reject(new Error("No hi ha cap correu en composició."));
      return;
    }

    const coercion = t.es_html
      ? Office.CoercionType.Html
      : Office.CoercionType.Text;

    // Assumpte opcional: només si la plantilla en porta i el correu no en té.
    maybeFillSubject(t)
      .catch((e) => Logger.warn("No s'ha pogut omplir l'assumpte: " + e))
      .finally(() => {
        item.body.setSelectedDataAsync(
          t.cos,
          { coercionType: coercion },
          (res) => {
            if (res.status === Office.AsyncResultStatus.Succeeded) {
              Logger.info(`Plantilla inserida al cursor: ${t.nom}`);
              resolve();
              return;
            }
            Logger.warn(
              "setSelectedDataAsync ha fallat; es prova prepend. " +
                res.error?.message
            );
            item.body.prependAsync(
              t.cos,
              { coercionType: coercion },
              (res2) => {
                if (res2.status === Office.AsyncResultStatus.Succeeded) {
                  Logger.info(`Plantilla afegida al principi: ${t.nom}`);
                  resolve();
                } else {
                  Logger.error("prependAsync també ha fallat.", res2.error);
                  reject(new Error(res2.error?.message || "Error inserint plantilla"));
                }
              }
            );
          }
        );
      });
  });
}

function maybeFillSubject(t: EmailTemplate): Promise<void> {
  return new Promise<void>((resolve) => {
    const subject = t.assumpte?.trim();
    const item = Office.context.mailbox.item;
    if (!subject || !item || !item.subject || !("setAsync" in item.subject)) {
      resolve();
      return;
    }
    item.subject.getAsync((getRes) => {
      const current =
        getRes.status === Office.AsyncResultStatus.Succeeded ? getRes.value : "";
      if (current && current.trim().length > 0) {
        resolve(); // Ja té assumpte; no el sobreescrivim.
        return;
      }
      (item.subject as Office.Subject).setAsync(subject, () => resolve());
    });
  });
}

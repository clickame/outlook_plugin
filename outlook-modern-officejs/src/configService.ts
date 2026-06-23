// Persistència de configuració local.
// Estratègia: Office.context.roamingSettings (es desa a la bústia de l'usuari i és
// COMPARTIT entre el task pane i els runtimes d'esdeveniments —i entre dispositius—,
// funciona al nou Outlook, OWA, Outlook clàssic i Mac).
//
// IMPORTANT: NO fem servir OfficeRuntime.storage perquè NO està disponible al nou
// Outlook per Windows ni a OWA, i el fallback a localStorage NO es comparteix entre
// el task pane i el runtime dels esdeveniments (cada runtime té el seu localStorage
// aïllat). Això feia que el handler d'enviament llegís sempre la config per defecte
// i mai afegís el CCO.

import { Config, defaultConfig } from "./models";
import { Logger } from "./logger";

const STORAGE_KEY = "clickame_config";

/** Comprova si roamingSettings està disponible (qualsevol client Outlook modern). */
function hasRoamingSettings(): boolean {
  try {
    return (
      typeof Office !== "undefined" &&
      !!Office.context &&
      !!Office.context.roamingSettings &&
      typeof Office.context.roamingSettings.get === "function"
    );
  } catch {
    return false;
  }
}

async function rawGet(key: string): Promise<string | null> {
  if (hasRoamingSettings()) {
    try {
      const v = Office.context.roamingSettings.get(key);
      if (v === undefined || v === null) return null;
      return typeof v === "string" ? v : JSON.stringify(v);
    } catch (e) {
      Logger.error("Error llegint roamingSettings; es prova localStorage.", e);
    }
  }
  try {
    return window.localStorage.getItem(key);
  } catch {
    return null;
  }
}

async function rawSet(key: string, value: string): Promise<void> {
  if (hasRoamingSettings()) {
    await new Promise<void>((resolve, reject) => {
      try {
        Office.context.roamingSettings.set(key, value);
        Office.context.roamingSettings.saveAsync((res) => {
          if (res.status === Office.AsyncResultStatus.Succeeded) {
            resolve();
          } else {
            Logger.error("Error desant roamingSettings.", res.error);
            reject(res.error);
          }
        });
      } catch (e) {
        reject(e);
      }
    });
    return;
  }
  // Fallback (clients sense roamingSettings): no es comparteix entre runtimes.
  window.localStorage.setItem(key, value);
}

/** Carrega la configuració; si no existeix o és invàlida, retorna (i desa) la per defecte. */
export async function loadConfig(): Promise<Config> {
  try {
    const raw = await rawGet(STORAGE_KEY);
    if (!raw) {
      const cfg = defaultConfig();
      await saveConfig(cfg);
      Logger.info("No hi havia configuració; s'ha creat la per defecte.");
      return cfg;
    }
    const parsed = JSON.parse(raw) as Config;
    return normalize(parsed);
  } catch (e) {
    Logger.error("Error carregant configuració; s'usa la per defecte.", e);
    return defaultConfig();
  }
}

export async function saveConfig(config: Config): Promise<void> {
  const normalized = normalize(config);
  await rawSet(STORAGE_KEY, JSON.stringify(normalized));
  Logger.info("Configuració desada.");
}

/** Sanetja camps per evitar valors invàlids. */
function normalize(cfg: Config): Config {
  const base = defaultConfig();
  const out: Config = {
    email_bcc: (cfg.email_bcc ?? "").trim(),
    activar_bcc_auto: !!cfg.activar_bcc_auto,
    mode_bcc: cfg.mode_bcc === "on_new_mail" ? "on_new_mail" : "on_send",
    plantilles: Array.isArray(cfg.plantilles) ? cfg.plantilles : base.plantilles,
  };
  return out;
}

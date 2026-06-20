// Persistència de configuració local.
// Estratègia: OfficeRuntime.storage si està disponible (compartit entre task pane
// i handlers d'esdeveniments), amb fallback a localStorage.

import { Config, defaultConfig } from "./models";
import { Logger } from "./logger";

const STORAGE_KEY = "clickame_config";

/** Comprova si OfficeRuntime.storage està disponible en aquest client. */
function hasOfficeRuntimeStorage(): boolean {
  return (
    typeof OfficeRuntime !== "undefined" &&
    !!OfficeRuntime &&
    !!OfficeRuntime.storage &&
    typeof OfficeRuntime.storage.getItem === "function"
  );
}

async function rawGet(key: string): Promise<string | null> {
  if (hasOfficeRuntimeStorage()) {
    return await OfficeRuntime.storage.getItem(key);
  }
  return window.localStorage.getItem(key);
}

async function rawSet(key: string, value: string): Promise<void> {
  if (hasOfficeRuntimeStorage()) {
    await OfficeRuntime.storage.setItem(key, value);
    return;
  }
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

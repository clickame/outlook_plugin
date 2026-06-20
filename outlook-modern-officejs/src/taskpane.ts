// Lògica del task pane: configuració del CCO, gestió de plantilles i logs.

import { Config, EmailTemplate, newId } from "./models";
import { loadConfig, saveConfig } from "./configService";
import {
  getTemplates,
  upsertTemplate,
  deleteTemplate,
  insertTemplateIntoMail,
} from "./templateService";
import { Logger } from "./logger";

/* global Office, document */

let config: Config;
let selectedTemplateId: string | null = null;

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) {
    return;
  }
  document.addEventListener("DOMContentLoaded", init);
  // Si el DOM ja està llest (segons l'ordre de càrrega), inicialitza igualment.
  if (document.readyState !== "loading") init();
});

let initialized = false;
async function init(): Promise<void> {
  if (initialized) return;
  initialized = true;

  wireTabs();
  wireBccSection();
  wireTemplatesSection();
  wireLogsSection();

  config = await loadConfig();
  fillBccForm();
  await refreshTemplateList();
}

// ---------- Pestanyes ----------
function wireTabs(): void {
  const tabs = Array.from(document.querySelectorAll<HTMLButtonElement>(".tab"));
  tabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      tabs.forEach((t) => t.classList.remove("active"));
      tab.classList.add("active");
      const name = tab.dataset.tab;
      document.querySelectorAll<HTMLElement>(".tab-panel").forEach((p) => {
        p.classList.toggle("active", p.id === `tab-${name}`);
      });
      if (name === "logs") renderLogs();
    });
  });
}

// ---------- CCO ----------
function wireBccSection(): void {
  byId<HTMLButtonElement>("saveBcc").addEventListener("click", saveBcc);
}

function fillBccForm(): void {
  byId<HTMLInputElement>("bccEmail").value = config.email_bcc || "";
  byId<HTMLInputElement>("bccAuto").checked = !!config.activar_bcc_auto;
  byId<HTMLSelectElement>("bccMode").value = config.mode_bcc || "on_send";
}

async function saveBcc(): Promise<void> {
  const email = byId<HTMLInputElement>("bccEmail").value.trim();
  const auto = byId<HTMLInputElement>("bccAuto").checked;
  const mode = byId<HTMLSelectElement>("bccMode").value as Config["mode_bcc"];

  if (auto && !email) {
    setStatus("bccStatus", "Indica una adreça per activar el CCO automàtic.", false);
    return;
  }

  config.email_bcc = email;
  config.activar_bcc_auto = auto;
  config.mode_bcc = mode;

  try {
    await saveConfig(config);
    setStatus("bccStatus", "Configuració desada.", true);
  } catch (e) {
    Logger.error("Error desant CCO.", e);
    setStatus("bccStatus", "Error desant la configuració.", false);
  }
}

// ---------- Plantilles ----------
function wireTemplatesSection(): void {
  byId<HTMLSelectElement>("templateList").addEventListener("change", onSelectTemplate);
  byId<HTMLButtonElement>("newTemplate").addEventListener("click", onNewTemplate);
  byId<HTMLButtonElement>("saveTemplate").addEventListener("click", onSaveTemplate);
  byId<HTMLButtonElement>("deleteTemplate").addEventListener("click", onDeleteTemplate);
  byId<HTMLButtonElement>("insertSelected").addEventListener("click", onInsertTemplate);
}

async function refreshTemplateList(): Promise<void> {
  const list = byId<HTMLSelectElement>("templateList");
  const templates = await getTemplates();
  list.innerHTML = "";
  templates.forEach((t) => {
    const opt = document.createElement("option");
    opt.value = t.id;
    opt.textContent = t.nom || "(sense nom)";
    list.appendChild(opt);
  });
  if (selectedTemplateId) {
    list.value = selectedTemplateId;
  }
}

function onSelectTemplate(): void {
  const id = byId<HTMLSelectElement>("templateList").value;
  selectedTemplateId = id;
  const t = config.plantilles.find((x) => x.id === id);
  fillTemplateEditor(t);
}

function fillTemplateEditor(t?: EmailTemplate): void {
  byId<HTMLInputElement>("tplNom").value = t?.nom || "";
  byId<HTMLInputElement>("tplAssumpte").value = t?.assumpte || "";
  byId<HTMLTextAreaElement>("tplCos").value = t?.cos || "";
  byId<HTMLInputElement>("tplHtml").checked = t?.es_html ?? true;
}

function onNewTemplate(): void {
  selectedTemplateId = null;
  byId<HTMLSelectElement>("templateList").value = "";
  fillTemplateEditor(undefined);
  byId<HTMLInputElement>("tplNom").focus();
}

async function onSaveTemplate(): Promise<void> {
  const nom = byId<HTMLInputElement>("tplNom").value.trim();
  if (!nom) {
    setStatus("tplStatus", "La plantilla necessita un nom.", false);
    return;
  }
  const template: EmailTemplate = {
    id: selectedTemplateId || newId(),
    nom,
    assumpte: byId<HTMLInputElement>("tplAssumpte").value.trim(),
    cos: byId<HTMLTextAreaElement>("tplCos").value,
    es_html: byId<HTMLInputElement>("tplHtml").checked,
  };
  config = await upsertTemplate(template);
  selectedTemplateId = template.id;
  await refreshTemplateList();
  setStatus("tplStatus", "Plantilla desada.", true);
}

async function onDeleteTemplate(): Promise<void> {
  if (!selectedTemplateId) {
    setStatus("tplStatus", "Selecciona una plantilla per eliminar.", false);
    return;
  }
  config = await deleteTemplate(selectedTemplateId);
  selectedTemplateId = null;
  fillTemplateEditor(undefined);
  await refreshTemplateList();
  setStatus("tplStatus", "Plantilla eliminada.", true);
}

async function onInsertTemplate(): Promise<void> {
  const id = byId<HTMLSelectElement>("templateList").value || selectedTemplateId;
  if (!id) {
    setStatus("tplStatus", "Selecciona una plantilla.", false);
    return;
  }
  const t = config.plantilles.find((x) => x.id === id);
  if (!t) {
    setStatus("tplStatus", "Plantilla no trobada.", false);
    return;
  }
  try {
    await insertTemplateIntoMail(t);
    setStatus("tplStatus", "Plantilla inserida al correu.", true);
  } catch (e) {
    Logger.error("Error inserint plantilla.", e);
    setStatus("tplStatus", "No s'ha pogut inserir. Obre un correu en composició.", false);
  }
}

// ---------- Logs ----------
function wireLogsSection(): void {
  byId<HTMLButtonElement>("refreshLogs").addEventListener("click", renderLogs);
  byId<HTMLButtonElement>("clearLogs").addEventListener("click", () => {
    Logger.clear();
    renderLogs();
  });
}

function renderLogs(): void {
  byId<HTMLPreElement>("logOutput").textContent = Logger.read().join("\n");
}

// ---------- Utilitats ----------
function byId<T extends HTMLElement>(id: string): T {
  return document.getElementById(id) as T;
}

function setStatus(id: string, msg: string, ok: boolean): void {
  const el = byId<HTMLElement>(id);
  el.textContent = msg;
  el.className = "status " + (ok ? "ok" : "err");
}

// Models compartits per a l'add-in Office.js.
// La forma del JSON coincideix amb la del projecte VSTO per facilitar la portabilitat.

export type BccMode = "on_new_mail" | "on_send";

export interface EmailTemplate {
  /** Identificador estable per editar/eliminar. */
  id: string;
  /** Nom visible a la llista. */
  nom: string;
  /** Assumpte opcional; si és buit no es toca l'assumpte del correu. */
  assumpte?: string;
  /** Cos de la plantilla (HTML o text pla). */
  cos: string;
  /** Indica si `cos` conté HTML. */
  es_html: boolean;
}

export interface Config {
  email_bcc: string;
  activar_bcc_auto: boolean;
  mode_bcc: BccMode;
  plantilles: EmailTemplate[];
}

export function defaultConfig(): Config {
  return {
    email_bcc: "",
    activar_bcc_auto: false,
    mode_bcc: "on_send",
    plantilles: [
      {
        id: newId(),
        nom: "Salutació inicial",
        assumpte: "Gràcies pel teu interès",
        es_html: true,
        cos:
          "<p>Hola,</p><p>Gràcies per contactar amb nosaltres. " +
          "En breu et donarem resposta.</p><p>Salutacions,<br/>L'equip de Clickame</p>",
      },
    ],
  };
}

/** Genera un id raonablement únic sense dependències externes. */
export function newId(): string {
  return (
    Date.now().toString(36) +
    "-" +
    Math.random().toString(36).slice(2, 10)
  );
}

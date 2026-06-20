// Logger local molt bàsic. Escriu a la consola i manté un buffer circular
// a localStorage perquè es pugui inspeccionar des del task pane (botó "Veure logs").
// No envia res a cap servidor.

const LOG_KEY = "clickame_logs";
const MAX_LINES = 300;

function append(level: string, message: string): void {
  const line = `${new Date().toISOString()} [${level}] ${message}`;
  // eslint-disable-next-line no-console
  console.log(line);
  try {
    const raw = window.localStorage.getItem(LOG_KEY);
    const lines: string[] = raw ? JSON.parse(raw) : [];
    lines.push(line);
    while (lines.length > MAX_LINES) lines.shift();
    window.localStorage.setItem(LOG_KEY, JSON.stringify(lines));
  } catch {
    // Si localStorage no està disponible (mode esdeveniment), només queda la consola.
  }
}

export const Logger = {
  info: (m: string) => append("INFO", m),
  warn: (m: string) => append("WARN", m),
  error: (m: string, e?: unknown) =>
    append("ERROR", e ? `${m} :: ${stringifyError(e)}` : m),
  read(): string[] {
    try {
      const raw = window.localStorage.getItem(LOG_KEY);
      return raw ? (JSON.parse(raw) as string[]) : [];
    } catch {
      return [];
    }
  },
  clear(): void {
    try {
      window.localStorage.removeItem(LOG_KEY);
    } catch {
      /* ignore */
    }
  },
};

function stringifyError(e: unknown): string {
  if (e instanceof Error) return `${e.name}: ${e.message}`;
  try {
    return JSON.stringify(e);
  } catch {
    return String(e);
  }
}

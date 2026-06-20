// Declaració mínima de l'API OfficeRuntime.storage que fem servir.
// @types/office-js no declara el namespace OfficeRuntime; només necessitem storage.

declare namespace OfficeRuntime {
  interface Storage {
    getItem(key: string): Promise<string | null>;
    setItem(key: string, value: string): Promise<void>;
    removeItem(key: string): Promise<void>;
    getItems(keys: string[]): Promise<{ [key: string]: string | null }>;
    setItems(items: { [key: string]: string }): Promise<void>;
    removeItems(keys: string[]): Promise<void>;
  }
  const storage: Storage;
}

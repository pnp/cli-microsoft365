export interface TokenStorage {
  get: () => Promise<string>;
  set: (connectionInfo: string) => Promise<void>;
  remove: () => Promise<void>;
}
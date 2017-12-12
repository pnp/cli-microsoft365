export interface TokenStorage {
  get: (service: string) => Promise<string>;
  set: (service: string, token: string) => Promise<void>;
  remove: (service: string) => Promise<void>;
}
export interface Environment {
  id: string;
  type: string;
  location: string;
  name: string;
  properties: {
    displayName: string;
    isDefault: boolean;
  }
}
export interface Environment {
  id: string;
  type: string;
  location: string;
  name: string;
  properties: {
    isDefault: boolean;
  }
}
export interface Connector {
  name: string;
  id: string;
  displayName?: string;
  properties: {
    displayName: string;
    apiDefinitions: {
      originalSwaggerUrl: string;
    };
    capabilities: any;
    connectionParameters: any;
    iconBrandColor: string;
    iconUri: string;
    swagger: any;
  };
}
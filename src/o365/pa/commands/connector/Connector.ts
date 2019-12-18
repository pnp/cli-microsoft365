export interface Connector {
  name: string;
  id: string;
  properties: {
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
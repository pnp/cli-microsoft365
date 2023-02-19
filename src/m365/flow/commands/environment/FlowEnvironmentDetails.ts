export interface FlowEnvironmentDetails {
  displayName?: string,
  provisioningState?: string,
  environmentSku?: string,
  azureRegionHint?: string,
  isDefault?: boolean,
  name: string,
  location: string,
  type: string,
  id: string,
  properties: FlowEnvironmentProperties
}

interface FlowEnvironmentProperties {
  displayName: string,
  createdTime: string,
  createdBy: {
    id: string,
    displayName: string,
    type: string
  },
  provisioningState: string,
  creationType: string,
  environmentSku: string,
  environmentType: string,
  isDefault: boolean,
  azureRegionHint: string,
  runtimeEndpoints: RuntimeEndpoints
}

interface RuntimeEndpoints<T = string> {
  [key: string]: T;
}
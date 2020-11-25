export interface ClientSideControl {
  controlType: number;
  displayMode: number;
  id: string;
  position: Position;
  webPartId: string;
  addedFromPersistedData: boolean;
  reservedHeight: number;
  reservedWidth: number;
  webPartData: WebPartData;
}

export interface WebPartData {
  id: string;
  instanceId: string;
  title: string;
  description: string;
  dataVersion: string;
  properties: Properties;
  serverProcessedContent?: ServerProcessedContent;
}

export interface Properties {
  [name: string]: any;
}

export interface ServerProcessedContent {
  [name: string]: any;
}

export interface Position {
  zoneIndex: number;
  sectionIndex: number;
  sectionFactor: number;
  layoutIndex: number;
  controlIndex: number;
}
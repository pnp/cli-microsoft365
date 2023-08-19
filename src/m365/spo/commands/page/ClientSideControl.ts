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

interface WebPartData {
  id: string;
  instanceId: string;
  title: string;
  description: string;
  dataVersion: string;
  properties: Properties;
  serverProcessedContent?: ServerProcessedContent;
}

interface Properties {
  [name: string]: any;
}

interface ServerProcessedContent {
  [name: string]: any;
}

interface Position {
  zoneIndex: number;
  sectionIndex: number;
  sectionFactor: number;
  layoutIndex: number;
  controlIndex: number;
}
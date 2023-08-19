
export interface PageControl {
  id: string;
  position: Position;
  emphasis: any;
  controlType?: number;
  displayMode?: number;
  webPartId?: string;
  addedFromPersistedData?: boolean;
  reservedHeight?: number;
  reservedWidth?: number;
  webPartData?: WebPartData;
}

interface WebPartData {
  id: string;
  instanceId: string;
  title: string;
  description: string;
  audiences: any[];
  serverProcessedContent: null[];
  dataVersion: string;
  properties: any[];
}

interface Position {
  zoneIndex: number;
  sectionIndex: number;
  sectionFactor: number;
  layoutIndex: number;
  controlIndex?: number;
}
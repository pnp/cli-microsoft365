export interface Control {
  controlType?: number;
  displayMode: number;
  emphasis: {};
  id?: string;
  position: ControlPosition;
  reservedHeight?: number;
  reservedWidth?: number;
  webPartData?: any;
  webPartId?: string;
}

export interface ControlPosition {
  controlIndex?: number;
  layoutIndex: number;
  sectionFactor: number;
  sectionIndex: number;
  zoneIndex: number;
}
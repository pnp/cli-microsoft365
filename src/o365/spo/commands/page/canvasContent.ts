export interface Control {
  controlType?: number;
  displayMode: number;
  emphasis: {};
  position: ControlPosition;
}

export interface ControlPosition {
  layoutIndex: number;
  sectionFactor: number;
  sectionIndex: number;
  zoneIndex: number;
}
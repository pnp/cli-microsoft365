import { ClientSideControlPosition, ZoneGroupMetadata } from "./clientsidepages";

export interface Control {
  controlType?: number;
  displayMode?: number;
  emphasis?: { zoneEmphasis?: number };
  id?: string;
  position: ClientSideControlPosition;
  reservedHeight?: number;
  reservedWidth?: number;
  webPartData?: any;
  webPartId?: string;
  zoneGroupMetadata?: ZoneGroupMetadata;
}

export interface BackgroundControl {
  controlType: number;
  position?: any;
  webPartData: {
    properties: {
      zoneBackground: {
        [key: string]: {
          type: string;
          gradient?: string;
          imageData?: {
            source: number;
            fileName: string;
            height: number;
            width: number;
          };
          useLightText: boolean;
          overlay: {
            color: string;
            opacity: number;
          }
        }
      }
    },
    serverProcessedContent: {
      htmlStrings: any,
      searchablePlainTexts: any,
      imageSources?: {
        [key: string]: string
      },
      links: any
    },
    dataVersion: string;
  }
}
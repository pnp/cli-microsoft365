import { AllSiteTypeProperties } from './AllSiteTypeProperties';

export interface SPOSitePropertiesEnumerable {
  _Child_Items_: AllSiteTypeProperties[]; 
  NextStartIndex: number;
  NextStartIndexFromSharePoint: string;
}
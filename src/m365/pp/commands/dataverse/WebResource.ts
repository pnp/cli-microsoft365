export interface Webresource {
  webresourceid: string,
  name: string,
  componentstate: number,
  content: string,
  content_binary: string,
  contentfileref: string,
  contentjson: string,
  contentjsonfileref: string,
  createdon: Date,
  dependencyxml: string,
  description: string,
  displayname: string,
  introducedversion: string,
  isavailableformobileoffline: boolean,
  isenabledformobileclient: boolean,
  ismanaged: boolean,
  languagecode: number,
  modifiedon: Date,
  overwritetime: Date,
  silverlightversion: string,
  solutionid: string,
  versionnumber: number,
  webresourceidunique: string,
  webresourcetype: number,
  canbedeleted: BooleanManagedProperty,
  iscustomizable: BooleanManagedProperty;
  ishidden: BooleanManagedProperty;
}

export interface BooleanManagedProperty {
  Value: boolean;
  CanBeChanged: boolean;
  ManagedPropertyLogicalName: string;
}

export const WEBRESOURCE_TYPE_LABELS: string[] = [
  'Webpage (HTML)',
  'Stylesheet (CSS)',
  'Script (JScript)',
  'Data (XML)',
  'PNG Format',
  'JPG Format',
  'GIF Format',
  'Silverlight (XAP)',
  'Stylesheet (XSL)',
  'ICO Format',
  'Vector Format (SVG)',
  'String (RESX)'
];

export const COMPONENT_STATE_LABELS: string[] = [
  'Published',
  'Unpublished',
  'Deleted',
  'Deleted Unpublished'
];

export const IS_MANAGED_LABELS: string[] = [
  'Unmanaged',
  'Managed'
];
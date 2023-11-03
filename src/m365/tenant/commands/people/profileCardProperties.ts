export const profileCardPropertyNames: string[] = [
  'UserPrincipalName',
  'Fax',
  'StreetAddress',
  'PostalCode',
  'StateOrProvince',
  'Alias',
  'customAttribute1',
  'customAttribute2',
  'customAttribute3',
  'customAttribute4',
  'customAttribute5',
  'customAttribute6',
  'customAttribute7',
  'customAttribute8',
  'customAttribute9',
  'customAttribute10',
  'customAttribute11',
  'customAttribute12',
  'customAttribute13',
  'customAttribute14',
  'customAttribute15'
];

export interface ProfileCardProperty {
  directoryPropertyName: string;
  annotations: Annotation[];
}

interface Annotation {
  displayName: string;
  localizations: { displayName: string; languageTag: string }[];
}

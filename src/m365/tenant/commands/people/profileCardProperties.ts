export interface ProfileCardProperty {
  directoryPropertyName: string;
  annotations: Annotation[];
}

interface Annotation {
  displayName: string;
  localizations: { displayName: string; languageTag: string }[];
}
import { WebInstalledLanguageProperties, WebProperties } from './WebProperties.js';

export interface WebPropertiesCollection {
  value: WebProperties[];
}

export interface WebInstalledLanguagePropertiesCollection {
  Items: WebInstalledLanguageProperties[];
}
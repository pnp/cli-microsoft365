import { parse } from 'csv-parse/sync';
import chalk from 'chalk';
import stripJsonComments from 'strip-json-comments';
import { BasePermissions } from '../m365/spo/base-permissions.js';
import { RoleDefinition } from '../m365/spo/commands/roledefinition/RoleDefinition.js';
import { RoleType } from '../m365/spo/commands/roledefinition/RoleType.js';

/**
 * Has the particular check passed or failed
 */
export enum CheckStatus {
  Success,
  Failure,
  Information,
  Warning
}

export const formatting = {
  escapeXml(s: any | undefined): any | undefined {
    if (!s) {
      return s;
    }

    return s.toString().replace(/[<>&"]/g, (c: string): string => {
      let char: string = c;

      switch (c) {
        case '<':
          char = '&lt;';
          break;
        case '>':
          char = '&gt;';
          break;
        case '&':
          char = '&amp;';
          break;
        case '"':
          char = '&quot;';
          break;
      }

      return char;
    });
  },

  parseJsonWithBom(s: string): any {
    return JSON.parse(s.replace(/^\uFEFF/, ''));
  },

  /**
   * Tries to parse a string as JSON. If it fails, returns the original string.
   * @param value JSON string to parse.
   * @returns JSON object or the original string if parsing fails.
   */
  tryParseJson(value: string): any {
    try {
      if (typeof value !== 'string') {
        return value;
      }

      return JSON.parse(value);
    }
    catch {
      return value;
    }
  },

  filterObject(obj: any, propertiesToInclude: string[]): any {
    const objKeys = Object.keys(obj);
    return propertiesToInclude
      .filter(prop => objKeys.includes(prop))
      .reduce((filtered: any, key: string) => {
        filtered[key] = obj[key];
        return filtered;
      }, {});
  },

  setFriendlyPermissions(response: RoleDefinition[]): RoleDefinition[] {
    response.forEach((r: RoleDefinition) => {
      const permissions: BasePermissions = new BasePermissions();
      permissions.high = r.BasePermissions.High as number;
      permissions.low = r.BasePermissions.Low as number;
      r.BasePermissionsValue = permissions.parse();
      r.RoleTypeKindValue = RoleType[r.RoleTypeKind];
    });

    return response;
  },

  parseCsvToJson(s: string, quoteChar: string = '"', delimiter: string = ','): any {
    return parse(s, {
      quote: quoteChar,
      delimiter: delimiter,
      columns: true,
      skipEmptyLines: true,
      ltrim: true,
      cast: true
    });
  },

  encodeQueryParameter(value: string): string {
    if (!value) {
      return value;
    }

    return encodeURIComponent(value).replace(/'/g, "''");
  },

  removeSingleLineComments(s: string): string {
    return stripJsonComments(s);
  },

  splitAndTrim(s: string): string[] {
    return s.split(',').map(c => c.trim());
  },

  openTypesEncoder(value: string): string {
    return value
      .replace(/\%/g, '%25')
      .replace(/\./g, '%2E')
      .replace(/:/g, '%3A')
      .replace(/@/g, '%40')
      .replace(/#/g, '%23');
  },

  /**
   * Rewrites boolean values according to the definition:
   * Booleans are case-insensitive, and are represented by the following values.
   *   True: 1, yes, true, on
   *   False: 0, no, false, off
   * @value Stringified Boolean value to rewrite
   * @returns A stringified boolean with the value 'true' or 'false'. Returns the original value if it does not comply with the definition. 
   */
  rewriteBooleanValue(value: string): string {
    const argValue = value.toLowerCase();
    switch (argValue) {
      case '1':
      case 'true':
      case 'yes':
      case 'on':
        return 'true';
      case '0':
      case 'false':
      case 'no':
      case 'off':
        return 'false';
      default:
        return value;
    }
  },

  /**
   * Converts an object into an xml:
   * @obj the actual objec
   * @returns A string containing the xml 
   */
  objectToXml(obj: any): string {
    let xml = '';
    for (const prop in obj) {
      xml += "<" + prop + ">";
      if (obj[prop] instanceof Array) {
        for (const array in obj[prop]) {
          xml += this.objectToXml(new Object(obj[prop][array]));
        }
      }
      else {
        xml += obj[prop];
      }
      xml += "</" + prop + ">";
    }
    xml = xml.replace(/<\/?[0-9]{1,}>/g, '');
    return xml;
  },

  getStatus(result: CheckStatus, message: string): string {
    const primarySupported: boolean = process.platform !== 'win32' ||
      process.env.CI === 'true' ||
      process.env.TERM === 'xterm-256color';
    const success: string = primarySupported ? '✔' : '√';
    const failure: string = primarySupported ? '✖' : '×';
    const information: string = 'i';
    const warning: string = '!';
    switch (result) {
      case CheckStatus.Success:
        return `${chalk.green(success)} ${message}`;
      case CheckStatus.Failure:
        return `${chalk.red(failure)} ${message}`;
      case CheckStatus.Information:
        return `${chalk.blue(information)} ${message}`;
      case CheckStatus.Warning:
        return `${chalk.yellow(warning)} ${message}`;
    }
  },

  convertArrayToHashTable(key: string, array: any[]): any {
    const resultAsKeyValuePair: any = {};
    array.forEach((obj) => {
      resultAsKeyValuePair[obj[key]] = obj;
    });
    return resultAsKeyValuePair;
  }
};
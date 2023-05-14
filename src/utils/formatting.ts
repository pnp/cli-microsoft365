import * as stripJsonComments from 'strip-json-comments';
import { BasePermissions } from '../m365/spo/base-permissions';
import { RoleDefinition } from '../m365/spo/commands/roledefinition/RoleDefinition';
import { RoleType } from '../m365/spo/commands/roledefinition/RoleType';

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
    const regex = new RegExp(`\\s*(${quoteChar})?(.*?)\\1\\s*(?:${delimiter}|$)`, 'gs');
    const lines: string[] = s.split('\n');

    const match = (line: string): string[] => [...line.matchAll(regex)]
      .map(m => m[2])  // we only want the second capture group
      .slice(0, -1);   // cut off blank match at the end

    const heads = match(lines[0]);

    return lines.slice(1).filter(text => text !== '').map(line => {
      return match(line).reduce((acc, cur, i) => {
        const val = cur;
        const numValue = parseInt(val);
        const key = heads[i];
        return { ...acc, [key]: isNaN(numValue) || numValue.toString() !== val ? val : numValue };
      }, {});
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
  }
};
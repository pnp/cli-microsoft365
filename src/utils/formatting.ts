import * as stripJsonComments from 'strip-json-comments';

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

  parseCsvToJson(s: string): any {
    const rows: string[] = s.split('\n');
    const jsonObj: any = [];
    const headers: string[] = rows[0].split(',');

    for (let i = 1; i < rows.length; i++) {
      const data: string[] = rows[i].split(',');
      const obj: any = {};
      for (let j = 0; j < data.length; j++) {
        const value = data[j].trim();
        const numValue = parseInt(value);
        obj[headers[j].trim()] = isNaN(numValue) || numValue.toString() !== value ? value : numValue;
      }
      jsonObj.push(obj);
    }

    return jsonObj;
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
   * @value Stringied Boolean value to rewrite
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
  }
};
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
    return Object.keys(obj)
      .filter(key => propertiesToInclude.includes(key))
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
  }
};
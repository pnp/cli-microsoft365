const packageJSON = require('../package.json');
import Table = require('easy-table');
import * as os from 'os';
const vorpal: Vorpal = require('./vorpal-init');
import { CommandError } from './Command';

export default class Utils {
  public static escapeXml(s: string | undefined) {
    if (!s) {
      return s;
    }

    return s.replace(/[<>&'"]/g, (c: string): string => {
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
        case '\'':
          char = '&apos;';
          break;
        case '"':
          char = '&quot;';
          break;
      }

      return char;
    });
  }

  public static restore(method: any | any[]): void {
    if (!method) {
      return;
    }

    if (!Array.isArray(method)) {
      method = [method];
    }

    method.forEach((m: any): void => {
      if (m && m.restore) {
        m.restore();
      }
    });
  }

  public static getRequestHeaders(headers: any): any {
    if (!headers) {
      headers = {};
    }

    headers['User-Agent'] = `NONISV|SharePointPnP|Office365CLI/${packageJSON.version}`;

    return headers;
  }

  public static isValidGuid(guid: string): boolean {
    const guidRegEx: RegExp = new RegExp(/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i);

    return guidRegEx.test(guid);
  }

  public static logOutput(stdout: any): any {
    // what comes in, should be an array
    // if it's not, return as-is
    if (!Array.isArray(stdout)) {
      return stdout;
    }

    let logStatement: any = stdout.pop();
    const logStatementType: string = typeof logStatement;

    if (logStatementType === 'undefined') {
      return logStatement;
    }

    if (vorpal._command &&
      vorpal._command.args &&
      vorpal._command.args.options &&
      vorpal._command.args.options.output === 'json') {
      return JSON.stringify(logStatement);
    }

    if (logStatement instanceof CommandError) {
      return vorpal.chalk.red(`Error: ${logStatement.message}`);
    }

    let arrayType: string = '';
    if (!Array.isArray(logStatement)) {
      logStatement = [logStatement];
      arrayType = logStatementType;
    }
    else {
      for (let i: number = 0; i < logStatement.length; i++) {
        const t: string = typeof logStatement[i];
        if (t !== 'undefined') {
          arrayType = t;
          break;
        }
      }
    }

    if (arrayType !== 'object') {
      return logStatement.join(os.EOL);
    }

    const t: Table = new Table();
    logStatement.forEach((r: any) => {
      if (typeof r !== 'object') {
        return;
      }

      Object.getOwnPropertyNames(r).forEach(p => {
        t.cell(p, r[p]);
      })
      t.newRow();
    });

    if (logStatement.length === 1) {
      return t.printTransposed({
        separator: ': '
      });
    }
    else {
      return t.toString();
    }
  }
}
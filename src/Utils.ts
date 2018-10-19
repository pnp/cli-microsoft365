const packageJSON = require('../package.json');
import Table = require('easy-table');
import * as os from 'os';
const vorpal: Vorpal = require('./vorpal-init');
import { CommandError } from './Command';
import * as url from 'url';

export default class Utils {
  public static escapeXml(s: any | undefined) {
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

  public static isValidBoolean(value: string): boolean {
    return value.toLowerCase() === 'true' || value.toLowerCase() === 'false'
  }

  public static logOutput(stdout: any): any {
    // what comes in, should be an array
    // if it's not, return as-is
    if (!Array.isArray(stdout)) {
      return stdout;
    }

    let logStatement: any = stdout.pop();

    if (logStatement instanceof Date) {
      return logStatement.toString();
    }

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

    if (logStatement.length === 1) {
      const obj: any = logStatement[0];
      const propertyNames: string[] = [];
      Object.getOwnPropertyNames(obj).forEach(p => {
        propertyNames.push(p);
      });

      let longestPropertyLength: number = 0;
      propertyNames.forEach(p => {
        if (p.length > longestPropertyLength) {
          longestPropertyLength = p.length;
        }
      });

      const output: string[] = [];
      propertyNames.sort().forEach(p => {
        output.push(`${p.length < longestPropertyLength ? p + new Array(longestPropertyLength - p.length + 1).join(' ') : p}: ${Array.isArray(obj[p]) || typeof obj[p] === 'object' ? JSON.stringify(obj[p]) : obj[p]}`);
      });

      return output.join('\n') + '\n';
    }
    else {
      const t: Table = new Table();
      logStatement.forEach((r: any) => {
        if (typeof r !== 'object') {
          return;
        }

        Object.getOwnPropertyNames(r).forEach(p => {
          t.cell(p, r[p]);
        });
        t.newRow();
      });

      return t.toString();
    }
  }

  public static getUserNameFromAccessToken(accessToken: string): string {
    let userName: string = '';

    if (!accessToken || accessToken.length === 0) {
      return userName;
    }

    const chunks = accessToken.split('.');
    if (chunks.length !== 3) {
      return userName;
    }

    const tokenString: string = Buffer.from(chunks[1], 'base64').toString();
    try {
      const token: any = JSON.parse(tokenString);
      userName = token.upn;
    }
    catch {
    }

    return userName;
  }

  /**
   * Returns server relative path.
   * @param webUrl web full or web relative url e.g. https://contoso.sharepoint.com/sites/team1
   * @param folderRelativePath folder relative path e.g. /Shared Documents
   * @example
   * // returns "/sites/team1"
   * Utils.getServerRelativePath("https://contoso.sharepoint.com/sites/team1");
   * @example
   * // returns "/sites/team1/Shared Documents"
   * Utils.getServerRelativePath("https://contoso.sharepoint.com/sites/team1", "/Shared Documents");
   * @example
   * // returns "/sites/team1/Shared Documents"
   * Utils.getServerRelativePath("/sites/team1/", "/Shared Documents");
   */
  public static getServerRelativePath(webUrl: string, folderRelativePath: string = ""): string {
    const tenantUrl: string = `${url.parse(webUrl).protocol}//${url.parse(webUrl).hostname}`;
    let webRelativePath: string = webUrl.replace(tenantUrl, '');

    // will be used to remove relative path from the folderRelativePath
    // in case the web relative url is included
    let relativePathToRemove: string = webRelativePath;

    // add '/' at 0
    if (webRelativePath.charAt(0) !== '/') {
      webRelativePath = `/${webRelativePath}`;
    }
    else {
      relativePathToRemove = webRelativePath.substring(1);
    }

    // remove last '/' of webRelativePath
    if (webRelativePath.length > 1 &&
      webRelativePath.lastIndexOf('/') === webRelativePath.length - 1) {
      webRelativePath = webRelativePath.substring(0, webRelativePath.length - 1);
    }

    // remove the web relative path if it is contained in the folder relative path
    folderRelativePath = folderRelativePath.replace(relativePathToRemove, '');

    if (folderRelativePath !== '') {
      // add '/' at 0 for siteRelativePath 
      if (folderRelativePath.charAt(0) !== '/') {
        folderRelativePath = `/${folderRelativePath}`;
      }

      // remove last '/' of siteRelativePath
      if (folderRelativePath.lastIndexOf('/') === folderRelativePath.length - 1) {
        folderRelativePath = folderRelativePath.substring(0, folderRelativePath.length - 1);
      }

      if (webRelativePath === '/' && folderRelativePath !== '') {
        webRelativePath = folderRelativePath;
      }
      else {
        webRelativePath = `${webRelativePath}${folderRelativePath}`;
      }
    }

    return webRelativePath.replace('//', '/');
  }
}
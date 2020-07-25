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

  public static isValidGuid(guid: string): boolean {
    const guidRegEx: RegExp = new RegExp(/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i);

    return guidRegEx.test(guid);
  }

  public static isValidTeamsChannelId(guid: string): boolean {
    const guidRegEx: RegExp = new RegExp(/^19:[0-9a-zA-Z]+@thread\.(skype|tacv2)$/i);

    return guidRegEx.test(guid);
  }

  public static isDateInRange(date: string, monthOffset: number): boolean {
    const d: Date = new Date(date);
    let cutoffDate: Date = new Date();
    cutoffDate.setMonth(cutoffDate.getMonth() - monthOffset);
    return d > cutoffDate;
  }

  public static isValidISODate(date: string): boolean {
    const dateRegEx: RegExp = new RegExp(
      /^(19|20)\d\d[- /.](0[1-9]|1[012])[- /.](0[1-9]|[12][0-9]|3[01])$/i
    );
    return dateRegEx.test(date);
  }

  public static isValidISODateDashOnly(date: string): boolean {
    const dateTimeRegEx: RegExp = new RegExp(
      /^(\d{4}-[01]\d-[0-3]\dT[0-2]\d:[0-5]\d:[0-5]\d\.\d+([+-][0-2]\d:[0-5]\d|Z))|(\d{4}-[01]\d-[0-3]\dT[0-2]\d:[0-5]\d:[0-5]\d([+-][0-2]\d:[0-5]\d|Z))|(\d{4}-[01]\d-[0-3]\dT[0-2]\d:[0-5]\d([+-][0-2]\d:[0-5]\d|Z))$/i
    );
    const dateOnlyRegEx: RegExp = new RegExp(
      /^(19|20)\d\d[-](0[1-9]|1[012])[-](0[1-9]|[12][0-9]|3[01])$/i
    );
    return dateTimeRegEx.test(date) ? true : dateOnlyRegEx.test(date);
  }

  public static isValidISODateTime(dateTime: string): boolean {
    const withMilliSecsPattern: RegExp = new RegExp(
      /^[0-9]{4}-((0[13578]|1[02])-(0[1-9]|[12][0-9]|3[01])|(0[469]|11)-(0[1-9]|[12][0-9]|30)|(02)-(0[1-9]|[12][0-9]))T(0[0-9]|1[0-9]|2[0-3]):(0[0-9]|[1-5][0-9]):(0[0-9]|[1-5][0-9])\.[0-9]{3}Z$/);
    if (withMilliSecsPattern.test(dateTime)) {
      return true;
    }
    const withSecsPattern: RegExp = new RegExp(
      /^[0-9]{4}-((0[13578]|1[02])-(0[1-9]|[12][0-9]|3[01])|(0[469]|11)-(0[1-9]|[12][0-9]|30)|(02)-(0[1-9]|[12][0-9]))T(0[0-9]|1[0-9]|2[0-3]):(0[0-9]|[1-5][0-9]):(0[0-9]|[1-5][0-9])Z$/);
    if (withSecsPattern.test(dateTime)) {
      return true;
    }

    const withMinutesPattern: RegExp = new RegExp(
      /^[0-9]{4}-((0[13578]|1[02])-(0[1-9]|[12][0-9]|3[01])|(0[469]|11)-(0[1-9]|[12][0-9]|30)|(02)-(0[1-9]|[12][0-9]))T(0[0-9]|1[0-9]|2[0-3]):(0[0-9]|[1-5][0-9])Z$/);
    if (withMinutesPattern.test(dateTime)) {
      return true;
    }

    const withHoursPattern: RegExp = new RegExp(
      /^[0-9]{4}-((0[13578]|1[02])-(0[1-9]|[12][0-9]|3[01])|(0[469]|11)-(0[1-9]|[12][0-9]|30)|(02)-(0[1-9]|[12][0-9]))T(0[0-9]|1[0-9]|2[0-3])Z$/);
    if (withHoursPattern.test(dateTime)) {
      return true;
    }

    const withoutTimePattern: RegExp = new RegExp(
      /^[0-9]{4}-((0[13578]|1[02])-(0[1-9]|[12][0-9]|3[01])|(0[469]|11)-(0[1-9]|[12][0-9]|30)|(02)-(0[1-9]|[12][0-9]))$/);
    if (withoutTimePattern.test(dateTime)) {
      return true;
    }

    return false;
  }

  public static isValidBoolean(value: string): boolean {
    return value.toLowerCase() === 'true' || value.toLowerCase() === 'false'
  }

  public static getTenantIdFromAccessToken(accessToken: string): string {
    let tenantId: string = '';

    if (!accessToken || accessToken.length === 0) {
      return tenantId;
    }

    const chunks = accessToken.split('.');
    if (chunks.length !== 3) {
      return tenantId;
    }

    const tokenString: string = Buffer.from(chunks[1], 'base64').toString();
    try {
      const token: any = JSON.parse(tokenString);
      tenantId = token.tid;
    }
    catch {
    }

    return tenantId;
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
      // if authenticated using certificate, there is no upn so use
      // app display name instead
      userName = token.upn || token.app_displayname;
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
   * Utils.getServerRelativePath("https://contoso.sharepoint.com/sites/team1", "");
   * @example
   * // returns "/sites/team1/Shared Documents"
   * Utils.getServerRelativePath("https://contoso.sharepoint.com/sites/team1", "/Shared Documents");
   * @example
   * // returns "/sites/team1/Shared Documents"
   * Utils.getServerRelativePath("/sites/team1/", "/Shared Documents");
   */
  public static getServerRelativePath(webUrl: string, folderRelativePath: string): string {
    const tenantUrl: string = `${url.parse(webUrl).protocol}//${url.parse(webUrl).hostname}`;
    let webRelativePath: string = webUrl.replace(tenantUrl, '');

    // will be used to remove relative path from the folderRelativePath
    // in case the web relative url is included
    let relativePathToRemove: string = webRelativePath;

    // add '/' at 0
    if (webRelativePath[0] !== '/') {
      webRelativePath = `/${webRelativePath}`;
    }
    else {
      relativePathToRemove = webRelativePath.substring(1);
    }

    // remove last '/' of webRelativePath
    const webPathLastCharPos: number = webRelativePath.length - 1;

    if (webRelativePath.length > 1 &&
      webRelativePath[webPathLastCharPos] === '/') {
      webRelativePath = webRelativePath.substring(0, webPathLastCharPos);
    }

    // remove the web relative path if it is contained in the folder relative path
    const relativePathToRemoveIdx: number = folderRelativePath.toLowerCase().indexOf(relativePathToRemove.toLowerCase());

    if (relativePathToRemoveIdx > -1) {
      const pos: number = relativePathToRemoveIdx + relativePathToRemove.length;
      folderRelativePath = folderRelativePath.substring(pos, folderRelativePath.length);
    }

    if (folderRelativePath !== '') {
      // add '/' at 0 for siteRelativePath
      if (folderRelativePath[0] !== '/') {
        folderRelativePath = `/${folderRelativePath}`;
      }

      // remove last '/' of siteRelativePath
      const folderPathLastCharPos: number = folderRelativePath.length - 1;

      if (folderRelativePath[folderPathLastCharPos] === '/') {
        folderRelativePath = folderRelativePath.substring(0, folderPathLastCharPos);
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

  /**
   * Returns server relative site url.
   * @param webUrl web full or web relative url e.g. https://contoso.sharepoint.com/sites/team1
   * @example
   * // returns "/sites/team1"
   * Utils.getServerRelativeSiteUrl("https://contoso.sharepoint.com/sites/team1";
   * @example
   * // returns ""
   * Utils.getServerRelativeSiteUrl("https://contoso.sharepoint.com");
   * @example
   * // returns "/sites/team1/Shared Documents"
   * Utils.getServerRelativePath("/sites/team1/", "/Shared Documents");
   */
  public static getServerRelativeSiteUrl(webUrl: string): string {
    const serverRelativeSiteUrl = Utils.getServerRelativePath(webUrl, '');

    // return an empty string instead of / to prevent // replies
    return serverRelativeSiteUrl === '/' ? "" : serverRelativeSiteUrl;
  }

  /**
   * Returns web relative path from webUrl and folderUrl.
   * @param webUrl web full or web relative url e.g. https://contoso.sharepoint.com/sites/team1/
   * @param folderUrl folder server relative url e.g. /sites/team1/Lists/MyList
   * @example
   * // returns "/Lists/MyList"
   * Utils.getWebRelativePath("https://contoso.sharepoint.com/sites/team1/", "/sites/team1/Lists/MyList");
   * @example
   * // returns "/Shared Documents"
   * Utils.getWebRelativePath("/sites/team1/", "/sites/team1/Shared Documents");
   */
  public static getWebRelativePath(webUrl: string, folderUrl: string): string {

    let folderWebRelativePath: string = '';

    const tenantUrl: string = `${url.parse(webUrl).protocol}//${url.parse(webUrl).hostname}`;
    let webRelativePath: string = webUrl.replace(tenantUrl, '');

    // will be used to remove relative path from the folderRelativePath
    // in case the web relative url is included
    let relativePathToRemove: string = webRelativePath;

    // add '/' at 0
    if (webRelativePath[0] !== '/') {
      webRelativePath = `/${webRelativePath}`;
    }
    else {
      relativePathToRemove = webRelativePath.substring(1);
    }

    // remove last '/' of webRelativePath
    const webPathLastCharPos: number = webRelativePath.length - 1;

    if (webRelativePath.length > 1 &&
      webRelativePath[webPathLastCharPos] === '/') {
      webRelativePath = webRelativePath.substring(0, webPathLastCharPos);
    }

    // remove the web relative path if it is contained in the folder relative path
    const relativePathToRemoveIdx: number = folderUrl.toLowerCase().indexOf(relativePathToRemove.toLowerCase());

    if (relativePathToRemoveIdx > -1) {
      const pos: number = relativePathToRemoveIdx + relativePathToRemove.length;
      folderWebRelativePath = folderUrl.substring(pos, folderUrl.length);
    }
    else {
      folderWebRelativePath = folderUrl;
    }

    // add '/' at 0 for folderWebRelativePath
    if (folderWebRelativePath[0] !== '/') {
      folderWebRelativePath = `/${folderWebRelativePath}`;
    }

    // remove last '/' of folderWebRelativePath
    const folderPathLastCharPos: number = folderWebRelativePath.length - 1;

    if (folderWebRelativePath.length > 1 && folderWebRelativePath[folderPathLastCharPos] === '/') {
      folderWebRelativePath = folderWebRelativePath.substring(0, folderPathLastCharPos);
    }

    return folderWebRelativePath.replace('//', '/');
  }

  /**
   * Returns the absolute URL according to a Web URL and the server relative URL of a folder
   * @param webUrl The full URL of a web
   * @param serverRelativeUrl The server relative URL of a folder
   * @example
   * // returns "https://contoso.sharepoint.com/sites/team1/Lists/MyList"
   * Utils.getAbsoluteUrl("https://contoso.sharepoint.com/sites/team1/", "/sites/team1/Lists/MyList");
   */
  public static getAbsoluteUrl(webUrl: string, serverRelativeUrl: string): string {
    const uri: url.UrlWithStringQuery = url.parse(webUrl);
    const tenantUrl: string = `${uri.protocol}//${uri.hostname}`;
    if (serverRelativeUrl[0] !== '/') {
      serverRelativeUrl = `/${serverRelativeUrl}`;
    }
    return `${tenantUrl}${serverRelativeUrl}`;
  }

  public static isJavascriptReservedWord(input: string): boolean {
    const javascriptReservedWords: string[] = [
      "arguments",
      "await",
      "break",
      "case",
      "catch",
      "class",
      "const",
      "continue",
      "debugger",
      "default",
      "delete",
      "do",
      "else",
      "enum",
      "eval",
      "export",
      "extends",
      "false",
      "finally",
      "for",
      "function",
      "if",
      "implements",
      "import",
      "in",
      "instanceof",
      "interface",
      "let",
      "new",
      "null",
      "package",
      "private",
      "protected",
      "public",
      "return",
      "static",
      "super",
      "switch",
      "this",
      "throw",
      "true",
      "try",
      "typeof",
      "var",
      "void",
      "while",
      "with",
      "yield",
      "Array",
      "Date",
      "eval",
      "function",
      "hasOwnProperty",
      "Infinity",
      "isFinite",
      "isNaN",
      "isPrototypeOf",
      "length",
      "Math",
      "NaN",
      "name",
      "Number",
      "Object",
      "prototype",
      "String",
      "toString",
      "undefined",
      "valueOf",
      "alert",
      "all",
      "anchor",
      "anchors",
      "area",
      "assign",
      "blur",
      "button",
      "checkbox",
      "clearInterval",
      "clearTimeout",
      "clientInformation",
      "close",
      "closed",
      "confirm",
      "constructor",
      "crypto",
      "decodeURI",
      "decodeURIComponent",
      "defaultStatus",
      "document",
      "element",
      "elements",
      "embed",
      "embeds",
      "encodeURI",
      "encodeURIComponent",
      "escape",
      "event",
      "fileUpload",
      "focus",
      "form",
      "forms",
      "frame",
      "innerHeight",
      "innerWidth",
      "layer",
      "layers",
      "link",
      "location",
      "mimeTypes",
      "navigate",
      "navigator",
      "frames",
      "frameRate",
      "hidden",
      "history",
      "image",
      "images",
      "offscreenBuffering",
      "open",
      "opener",
      "option",
      "outerHeight",
      "outerWidth",
      "packages",
      "pageXOffset",
      "pageYOffset",
      "parent",
      "parseFloat",
      "parseInt",
      "password",
      "pkcs11",
      "plugin",
      "prompt",
      "propertyIsEnum",
      "radio",
      "reset",
      "screenX",
      "screenY",
      "scroll",
      "secure",
      "select",
      "self",
      "setInterval",
      "setTimeout",
      "status",
      "submit",
      "taint",
      "text",
      "textarea",
      "top",
      "unescape",
      "untaint",
      "window",
      "onblur",
      "onclick",
      "onerror",
      "onfocus",
      "onkeydown",
      "onkeypress",
      "onkeyup",
      "onmouseover",
      "onload",
      "onmouseup",
      "onmousedown",
      "onsubmit"
    ];
    return !!input && !input.split('.').every(value => !~javascriptReservedWords.indexOf(value));
  }

  public static isValidFileName(input: string): boolean {
    return !!input && !/^((\..*)|COM\d|CLOCK\$|LPT\d|AUX|NUL|CON|PRN|(.*[\u{d800}-\u{dfff}]+.*))$/iu.test(input) && !/^(.*\.\..*)$/i.test(input);
  }

  public static getSafeFileName(input: string): string {
    return input
      .replace(/'/g, "''")
  }

  public static isValidTheme(input: string): boolean {
    const validThemeProperties = [
      "themePrimary",
      "themeLighterAlt",
      "themeLighter",
      "themeLight",
      "themeTertiary",
      "themeSecondary",
      "themeDarkAlt",
      "themeDark",
      "themeDarker",
      "neutralLighterAlt",
      "neutralLighter",
      "neutralLight",
      "neutralQuaternaryAlt",
      "neutralQuaternary",
      "neutralTertiaryAlt",
      "neutralTertiary",
      "neutralSecondary",
      "neutralPrimaryAlt",
      "neutralPrimary",
      "neutralDark",
      "black",
      "white"
    ];
    let theme: any;

    try {
      theme = JSON.parse(input);
    }
    catch {
      return false;
    }

    if (Array.isArray(theme)) {
      return false;
    }

    const hasInvalidProperties = validThemeProperties.map((property) => {
      return theme.hasOwnProperty(property);
    }).includes(false);

    if (hasInvalidProperties) {
      return false;
    }

    const regex = new RegExp(/^#([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$/);
    const hasInvalidValues = validThemeProperties.map((property: string) => {
      return regex.test(theme[property]);
    }).includes(false);

    if (hasInvalidValues) {
      return false;
    }

    return true;
  }

  public static parseJsonWithBom(s: string): any {
    return JSON.parse(s.replace(/^\uFEFF/, ''));
  }
}

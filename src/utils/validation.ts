import { FormDigestInfo } from "./spo";

export const validation = {
  isValidGuidArray(guids: string[]): boolean {
    return guids.every(guid => this.isValidGuid(guid));
  },

  isValidGuid(guid?: string): boolean {
    if (!guid) {
      return false;
    }

    const guidRegEx: RegExp = new RegExp(/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i);

    // verify if the guid is a valid guid. @meid will be replaced in a later stage with the actual user id of the logged in user
    return guidRegEx.test(guid) || guid.toLocaleLowerCase().trim() === "@meid";
  },

  isValidTeamsChannelId(guid: string): boolean {
    const guidRegEx: RegExp = new RegExp(/^19:[0-9a-zA-Z-_]+@thread\.(skype|tacv2)$/i);

    return guidRegEx.test(guid);
  },

  isValidTeamsChatId(guid: string): boolean {
    const guidRegEx: RegExp = new RegExp(/^19:[0-9a-zA-Z-_]+(@thread\.v2|@unq\.gbl\.spaces)$/i);

    return guidRegEx.test(guid);
  },

  isValidUserPrincipalName(upn: string): boolean {
    const upnRegEx = new RegExp(/^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)*$/i);

    return upnRegEx.test(upn);
  },

  isDateInRange(date: string, monthOffset: number): boolean {
    const d: Date = new Date(date);
    const cutoffDate: Date = new Date();
    cutoffDate.setMonth(cutoffDate.getMonth() - monthOffset);
    return d > cutoffDate;
  },

  isValidISODate(date: string): boolean {
    const dateRegEx: RegExp = new RegExp(
      /^(19|20)\d\d[- /.](0[1-9]|1[012])[- /.](0[1-9]|[12][0-9]|3[01])$/i
    );
    return dateRegEx.test(date);
  },

  isValidISODateDashOnly(date: string): boolean {
    const dateTimeRegEx: RegExp = new RegExp(
      /^(\d{4}-[01]\d-[0-3]\dT[0-2]\d:[0-5]\d:[0-5]\d\.\d+([+-][0-2]\d:[0-5]\d|Z))|(\d{4}-[01]\d-[0-3]\dT[0-2]\d:[0-5]\d:[0-5]\d([+-][0-2]\d:[0-5]\d|Z))|(\d{4}-[01]\d-[0-3]\dT[0-2]\d:[0-5]\d([+-][0-2]\d:[0-5]\d|Z))$/i
    );
    const dateOnlyRegEx: RegExp = new RegExp(
      /^(19|20)\d\d[-](0[1-9]|1[012])[-](0[1-9]|[12][0-9]|3[01])$/i
    );
    return dateTimeRegEx.test(date) ? true : dateOnlyRegEx.test(date);
  },

  isValidISODateTime(dateTime: string): boolean {
    // Format: 2000-01-01T00:00:00.0000000Z
    const withMilliSecsLongPattern: RegExp = new RegExp(
      /^[0-9]{4}-((0[13578]|1[02])-(0[1-9]|[12][0-9]|3[01])|(0[469]|11)-(0[1-9]|[12][0-9]|30)|(02)-(0[1-9]|[12][0-9]))T(0[0-9]|1[0-9]|2[0-3]):(0[0-9]|[1-5][0-9]):(0[0-9]|[1-5][0-9])\.[0-9]{7}Z$/);
    if (withMilliSecsLongPattern.test(dateTime)) {
      return true;
    }

    // Format: 2000-01-01T00:00:00.000Z
    const withMilliSecsShortPattern: RegExp = new RegExp(
      /^[0-9]{4}-((0[13578]|1[02])-(0[1-9]|[12][0-9]|3[01])|(0[469]|11)-(0[1-9]|[12][0-9]|30)|(02)-(0[1-9]|[12][0-9]))T(0[0-9]|1[0-9]|2[0-3]):(0[0-9]|[1-5][0-9]):(0[0-9]|[1-5][0-9])\.[0-9]{3}Z$/);
    if (withMilliSecsShortPattern.test(dateTime)) {
      return true;
    }

    // Format: 2000-01-01T00:00:00Z
    const withSecsPattern: RegExp = new RegExp(
      /^[0-9]{4}-((0[13578]|1[02])-(0[1-9]|[12][0-9]|3[01])|(0[469]|11)-(0[1-9]|[12][0-9]|30)|(02)-(0[1-9]|[12][0-9]))T(0[0-9]|1[0-9]|2[0-3]):(0[0-9]|[1-5][0-9]):(0[0-9]|[1-5][0-9])Z$/);
    if (withSecsPattern.test(dateTime)) {
      return true;
    }

    // Format: 2000-01-01T00:00Z
    const withMinutesPattern: RegExp = new RegExp(
      /^[0-9]{4}-((0[13578]|1[02])-(0[1-9]|[12][0-9]|3[01])|(0[469]|11)-(0[1-9]|[12][0-9]|30)|(02)-(0[1-9]|[12][0-9]))T(0[0-9]|1[0-9]|2[0-3]):(0[0-9]|[1-5][0-9])Z$/);
    if (withMinutesPattern.test(dateTime)) {
      return true;
    }

    // Format: 2000-01-01T00Z
    const withHoursPattern: RegExp = new RegExp(
      /^[0-9]{4}-((0[13578]|1[02])-(0[1-9]|[12][0-9]|3[01])|(0[469]|11)-(0[1-9]|[12][0-9]|30)|(02)-(0[1-9]|[12][0-9]))T(0[0-9]|1[0-9]|2[0-3])Z$/);
    if (withHoursPattern.test(dateTime)) {
      return true;
    }

    // Format: 2000-01-01
    const withoutTimePattern: RegExp = new RegExp(
      /^[0-9]{4}-((0[13578]|1[02])-(0[1-9]|[12][0-9]|3[01])|(0[469]|11)-(0[1-9]|[12][0-9]|30)|(02)-(0[1-9]|[12][0-9]))$/);
    if (withoutTimePattern.test(dateTime)) {
      return true;
    }

    return false;
  },

  isValidBoolean(value: string): boolean {
    return value.toLowerCase() === 'true' || value.toLowerCase() === 'false';
  },

  isJavaScriptReservedWord(input: string): boolean {
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
  },

  isValidFileName(input: string): boolean {
    return !!input && !/^((\..*)|COM\d|CLOCK\$|LPT\d|AUX|NUL|CON|PRN|(.*[\u{d800}-\u{dfff}]+.*))$/iu.test(input) && !/^(.*\.\..*)$/i.test(input);
  },

  isValidTheme(input: string): boolean {
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
  },

  isValidSharePointUrl(url: string): boolean | string {
    if (!url) {
      return false;
    }

    if (url.indexOf('https://') !== 0) {
      return `${url} is not a valid SharePoint Online site URL`;
    }
    else {
      return true;
    }
  },

  isValidFormDigest(contextInfo: FormDigestInfo | undefined): boolean {
    if (!contextInfo) {
      return false;
    }

    const now: Date = new Date();
    if (contextInfo.FormDigestValue && now < contextInfo.FormDigestExpiresAt) {
      return true;
    }

    return false;
  },

  /**
  * Tests the two letter country code (ISO standard 3166).
  */
  isValidCountryCode(countryCode: string): boolean {
    const regex = new RegExp(/^(A(D|E|F|G|I|L|M|N|O|R|S|T|Q|U|W|X|Z)|B(A|B|D|E|F|G|H|I|J|L|M|N|O|R|S|T|V|W|Y|Z)|C(A|C|D|F|G|H|I|K|L|M|N|O|R|U|V|X|Y|Z)|D(E|J|K|M|O|Z)|E(C|E|G|H|R|S|T)|F(I|J|K|M|O|R)|G(A|B|D|E|F|G|H|I|L|M|N|P|Q|R|S|T|U|W|Y)|H(K|M|N|R|T|U)|I(D|E|Q|L|M|N|O|R|S|T)|J(E|M|O|P)|K(E|G|H|I|M|N|P|R|W|Y|Z)|L(A|B|C|I|K|R|S|T|U|V|Y)|M(A|C|D|E|F|G|H|K|L|M|N|O|Q|P|R|S|T|U|V|W|X|Y|Z)|N(A|C|E|F|G|I|L|O|P|R|U|Z)|OM|P(A|E|F|G|H|K|L|M|N|R|S|T|W|Y)|QA|R(E|O|S|U|W)|S(A|B|C|D|E|G|H|I|J|K|L|M|N|O|R|T|V|Y|Z)|T(C|D|F|G|H|J|K|L|M|N|O|R|T|V|W|Z)|U(A|G|M|S|Y|Z)|V(A|C|E|G|I|N|U)|W(F|S)|Y(E|T)|Z(A|M|W))$/);
    if (regex.test(countryCode)) {
      return true;
    }

    return false;
  },

  /**
  * Tests the language code (ISO 639-1).
  */
  isValidLanguageCode(languageCode: string): boolean {
    const regex = new RegExp(/^[a-zA-Z]{2,3}(?:-[a-zA-Z0-9]{2,4}){0,2}$/);
    if (regex.test(languageCode)) {
      return true;
    }

    return false;
  }
};
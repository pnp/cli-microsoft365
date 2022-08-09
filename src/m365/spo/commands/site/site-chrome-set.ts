import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

enum HeaderLayout {
  Standard = 1,
  Compact,
  Minimal,
  Extended
}

enum FooterLayout {
  Simple = 1,
  Extended
}

enum Alignment {
  Left = 0,
  Center,
  Right
}

enum Emphasis {
  Lightest = 0,
  Light,
  Dark,
  Darkest
}

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
  headerLayout?: HeaderLayout;
  headerEmphasis?: Emphasis;
  logoAlignment?: Alignment;
  footerLayout?: FooterLayout;
  footerEmphasis?: Emphasis;
  disableMegaMenu?: string;
  hideTitleInHeader?: string;
  disableFooter?: string;
}

class SpoSiteChromeSetCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_CHROME_SET;
  }

  public get description(): string {
    return 'Set the chrome header and footer for the specified site';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        headerLayout: args.options.headerLayout,
        headerEmphasis: args.options.headerEmphasis,
        disableMegaMenu: args.options.disableMegaMenu,
        hideTitleInHeader: args.options.hideTitleInHeader,
        logoAlignment: args.options.logoAlignment,
        disableFooter: args.options.disableFooter,
        footerLayout: args.options.footerLayout,
        footerEmphasis: args.options.footerEmphasis
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --url <url>'
      },
      {
        option: '--headerLayout [headerLayout]',
        autocomplete: ['Standard', 'Compact', 'Minimal', 'Extended']
      },
      {
        option: '--headerEmphasis [headerEmphasis]',
        autocomplete: ['Lightest', 'Light', 'Dark', 'Darkest']
      },
      {
        option: '--logoAlignment [logoAlignment]',
        autocomplete: ['Left', 'Center', 'Right']
      },
      {
        option: '--footerLayout [footerLayout]',
        autocomplete: ['Simple', 'Extended']
      },
      {
        option: '--footerEmphasis [footerEmphasis]',
        autocomplete: ['Lightest', 'Light', 'Dark', 'Darkest']
      },
      {
        option: '--disableMegaMenu [disableMegaMenu]'
      },
      {
        option: '--hideTitleInHeader [hideTitleInHeader]'
      },
      {
        option: '--disableFooter [disableFooter]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.url);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (typeof args.options.footerEmphasis !== "undefined" && !(args.options.footerEmphasis in Emphasis)) {
          return `${args.options.footerEmphasis} is not a valid option for footerEmphasis. Allowed values Lightest|Light|Dark|Darkest`;
        }

        if (typeof args.options.footerLayout !== "undefined" && !(args.options.footerLayout in FooterLayout)) {
          return `${args.options.footerLayout} is not a valid option for footerLayout. Allowed values Simple|Extended`;
        }

        if (typeof args.options.headerEmphasis !== "undefined" && !(args.options.headerEmphasis in Emphasis)) {
          return `${args.options.headerEmphasis} is not a valid option for headerEmphasis. Allowed values Lightest|Light|Dark|Darkest`;
        }

        if (typeof args.options.headerLayout !== "undefined" && !(args.options.headerLayout in HeaderLayout)) {
          return `${args.options.headerLayout} is not a valid option for headerLayout. Allowed values Standard|Compact|Minimal|Extended`;
        }

        if (typeof args.options.logoAlignment !== "undefined" && !(args.options.logoAlignment in Alignment)) {
          return `${args.options.logoAlignment} is not a valid option for logoAlignment. Allowed values Left|Center|Right`;
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const headerLayout = args.options.headerLayout ? HeaderLayout[args.options.headerLayout] : null;
    const headerEmphasis = args.options.headerEmphasis ? Emphasis[args.options.headerEmphasis] : null;
    const logoAlignment = args.options.logoAlignment ? Alignment[args.options.logoAlignment] : null;
    const footerLayout = args.options.footerLayout ? FooterLayout[args.options.footerLayout] : null;
    const footerEmphasis = args.options.footerEmphasis ? Emphasis[args.options.footerEmphasis] : null;
    const hideTitleInHeader = typeof args.options.hideTitleInHeader !== "undefined" ? args.options.hideTitleInHeader.toLowerCase() === "true" : null;
    const disableMegaMenu = typeof args.options.disableMegaMenu !== 'undefined' ? args.options.disableMegaMenu.toLowerCase() === "true" : null;
    const disableFooter = typeof args.options.disableFooter !== 'undefined' ? args.options.disableFooter.toLowerCase() === "true" : null;

    const body: any = {};
    if (headerLayout !== null) {
      body["headerLayout"] = headerLayout;
    }
    if (headerEmphasis !== null) {
      body["headerEmphasis"] = headerEmphasis;
    }
    if (logoAlignment !== null) {
      body["logoAlignment"] = logoAlignment;
    }
    if (footerLayout !== null) {
      body["footerLayout"] = footerLayout;
    }
    if (footerEmphasis !== null) {
      body["footerEmphasis"] = 3 - parseInt(footerEmphasis); // Footer is inverted
    }
    if (hideTitleInHeader !== null) {
      body["hideTitleInHeader"] = hideTitleInHeader;
    }
    if (disableMegaMenu !== null) {
      body["megaMenuEnabled"] = !disableMegaMenu;
    }
    if (disableFooter !== null) {
      body["footerEnabled"] = !disableFooter;
    }

    const requestOptions: any = {
      url: `${args.options.url}/_api/web/SetChromeOptions`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      data: body,
      responseType: 'json'
    };

    request
      .post(requestOptions)
      .then(_ => cb(),
        (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new SpoSiteChromeSetCommand();
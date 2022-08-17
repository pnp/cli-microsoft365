import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  description?: string;
  headerEmphasis?: number;
  headerLayout?: string;
  megaMenuEnabled?: string;
  quickLaunchEnabled?: string;
  siteLogoUrl?: string;
  title?: string;
  webUrl: string;
  footerEnabled?: string;
  searchScope?: string;
}

class SpoWebSetCommand extends SpoCommand {
  private static searchScopeOptions: string[] = ['defaultscope', 'tenant', 'hub', 'site'];

  public get name(): string {
    return commands.WEB_SET;
  }

  public get description(): string {
    return 'Updates subsite properties';
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
        description: typeof args.options.description !== 'undefined',
        headerEmphasis: typeof args.options.headerEmphasis !== 'undefined',
        headerLayout: typeof args.options.headerLayout !== 'undefined',
        megaMenuEnabled: typeof args.options.megaMenuEnabled !== 'undefined',
        siteLogoUrl: typeof args.options.siteLogoUrl !== 'undefined',
        title: typeof args.options.title !== 'undefined',
        quickLaunchEnabled: typeof args.options.quickLaunchEnabled !== 'undefined',
        footerEnabled: typeof args.options.footerEnabled !== 'undefined',
        searchScope: args.options.searchScope !== 'undefined'
      });
      this.trackUnknownOptions(this.telemetryProperties, args.options);
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-t, --title [title]'
      },
      {
        option: '-d, --description [description]'
      },
      {
        option: '--siteLogoUrl [siteLogoUrl]'
      },
      {
        option: '--quickLaunchEnabled [quickLaunchEnabled]'
      },
      {
        option: '--headerLayout [headerLayout]',
        autocomplete: ['standard', 'compact']
      },
      {
        option: '--headerEmphasis [headerEmphasis]',
        autocomplete: ['0', '1', '2', '3']
      },
      {
        option: '--megaMenuEnabled [megaMenuEnabled]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--footerEnabled [footerEnabled]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--searchScope [searchScope]',
        autocomplete: SpoWebSetCommand.searchScopeOptions
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (typeof args.options.quickLaunchEnabled !== 'undefined') {
          if (args.options.quickLaunchEnabled !== 'true' &&
            args.options.quickLaunchEnabled !== 'false') {
            return `${args.options.quickLaunchEnabled} is not a valid boolean value`;
          }
        }

        if (typeof args.options.headerEmphasis !== 'undefined') {
          if (isNaN(args.options.headerEmphasis)) {
            return `${args.options.headerEmphasis} is not a number`;
          }

          if ([0, 1, 2, 3].indexOf(args.options.headerEmphasis) < 0) {
            return `${args.options.headerEmphasis} is not a valid value for headerEmphasis. Allowed values are 0|1|2|3`;
          }
        }

        if (typeof args.options.headerLayout !== 'undefined') {
          if (['standard', 'compact'].indexOf(args.options.headerLayout) < 0) {
            return `${args.options.headerLayout} is not a valid value for headerLayout. Allowed values are standard|compact`;
          }
        }

        if (typeof args.options.megaMenuEnabled !== 'undefined') {
          if (['true', 'false'].indexOf(args.options.megaMenuEnabled) < 0) {
            return `${args.options.megaMenuEnabled} is not a valid boolean value`;
          }
        }

        if (typeof args.options.footerEnabled !== 'undefined') {
          if (args.options.footerEnabled !== 'true' &&
            args.options.footerEnabled !== 'false') {
            return `${args.options.footerEnabled} is not a valid boolean value`;
          }
        }

        if (typeof args.options.searchScope !== 'undefined') {
          const searchScope = args.options.searchScope.toString().toLowerCase();
          if (SpoWebSetCommand.searchScopeOptions.indexOf(searchScope) < 0) {
            return `${args.options.searchScope} is not a valid value for searchScope. Allowed values are DefaultScope|Tenant|Hub|Site`;
          }
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    const payload: any = {};

    this.addUnknownOptionsToPayload(payload, args.options);

    if (args.options.title) {
      payload.Title = args.options.title;
    }
    if (args.options.description) {
      payload.Description = args.options.description;
    }
    if (typeof args.options.siteLogoUrl !== 'undefined') {
      payload.SiteLogoUrl = args.options.siteLogoUrl;
    }
    if (typeof args.options.quickLaunchEnabled !== 'undefined') {
      payload.QuickLaunchEnabled = args.options.quickLaunchEnabled === 'true';
    }
    if (typeof args.options.headerEmphasis !== 'undefined') {
      payload.HeaderEmphasis = args.options.headerEmphasis;
    }
    if (typeof args.options.headerLayout !== 'undefined') {
      payload.HeaderLayout = args.options.headerLayout === 'standard' ? 1 : 2;
    }
    if (typeof args.options.megaMenuEnabled !== 'undefined') {
      payload.MegaMenuEnabled = args.options.megaMenuEnabled === 'true';
    }
    if (typeof args.options.footerEnabled !== 'undefined') {
      payload.FooterEnabled = args.options.footerEnabled === 'true';
    }
    if (typeof args.options.searchScope !== 'undefined') {
      const searchScope = args.options.searchScope.toLowerCase();
      payload.SearchScope = SpoWebSetCommand.searchScopeOptions.indexOf(searchScope);
    }

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/web`,
      headers: {
        'content-type': 'application/json;odata=nometadata',
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json',
      data: payload
    };

    if (this.verbose) {
      logger.logToStderr(`Updating properties of subsite ${args.options.webUrl}...`);
    }

    request
      .patch(requestOptions)
      .then(_ => cb(), (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }
}

module.exports = new SpoWebSetCommand();
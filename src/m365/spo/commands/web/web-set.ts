import commands from '../../commands';
import request from '../../../../request';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

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
  private static searchScopeOptions: string[] =
    ['defaultscope', 'tenant', 'hub', 'site'];

  public get name(): string {
    return commands.WEB_SET;
  }

  public get description(): string {
    return 'Updates subsite properties';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.description = typeof args.options.description !== 'undefined';
    telemetryProps.headerEmphasis = typeof args.options.headerEmphasis !== 'undefined';
    telemetryProps.headerLayout = typeof args.options.headerLayout !== 'undefined';
    telemetryProps.megaMenuEnabled = typeof args.options.megaMenuEnabled !== 'undefined';
    telemetryProps.siteLogoUrl = typeof args.options.siteLogoUrl !== 'undefined';
    telemetryProps.title = typeof args.options.title !== 'undefined';
    telemetryProps.quickLaunchEnabled = typeof args.options.quickLaunchEnabled !== 'undefined';
    telemetryProps.footerEnabled = typeof args.options.footerEnabled !== 'undefined';
    telemetryProps.searchScope = args.options.searchScope !== 'undefined';
    this.trackUnknownOptions(telemetryProps, args.options);
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const payload: any = {};

    this.addUnknownOptionsToPayload(payload, args.options);

    if (args.options.title) {
      payload.Title = args.options.title;
    }
    if (args.options.description) {
      payload.Description = args.options.description;
    }
    if (args.options.siteLogoUrl) {
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
      json: true,
      body: payload
    };

    if (this.verbose) {
      cmd.log(`Updating properties of subsite ${args.options.webUrl}...`);
    }

    request
      .patch(requestOptions)
      .then((): void => {
        if (this.debug) {
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the subsite to update'
      },
      {
        option: '-t, --title [title]',
        description: 'New title for the subsite'
      },
      {
        option: '-d, --description [description]',
        description: 'New description for the subsite'
      },
      {
        option: '--siteLogoUrl [siteLogoUrl]',
        description: 'New site logo URL for the subsite'
      },
      {
        option: '--quickLaunchEnabled [quickLaunchEnabled]',
        description: 'Set to true to enable quick launch and to false to disable it'
      },
      {
        option: '--headerLayout [headerLayout]',
        description: 'Configures the site header. Allowed values standard|compact',
        autocomplete: ['standard', 'compact']
      },
      {
        option: '--headerEmphasis [headerEmphasis]',
        description: 'Configures the site header background. Allowed values 0|1|2|3',
        autocomplete: ['0', '1', '2', '3']
      },
      {
        option: '--megaMenuEnabled [megaMenuEnabled]',
        description: 'Set to \'true\' to change the menu style to megamenu. Set to \'false\' to use the cascading menu style',
        autocomplete: ['true', 'false']
      },
      {
        option: '--footerEnabled [footerEnabled]',
        description: 'Set to \'true\' to enable footer and to \'false\' to disable it',
        autocomplete: ['true', 'false']
      },
      {
        option: '--searchScope [searchScope]',
        description: 'Search scope to set in the site. Allowed values DefaultScope|Tenant|Hub|Site',
        autocomplete: SpoWebSetCommand.searchScopeOptions
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
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

      return this.validateUnknownOptions(args.options, 'web', 'set');
    };
  }
}

module.exports = new SpoWebSetCommand();
import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import * as request from 'request-promise-native';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import { Auth } from '../../../../Auth';
import Utils from '../../../../Utils';
const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  description?: string;
  quickLaunchEnabled?: string;
  siteLogoUrl?: string;
  title?: string;
  webUrl: string;
}

class SpoWebSetCommand extends SpoCommand {
  public get name(): string {
    return commands.WEB_SET;
  }

  public get description(): string {
    return 'Updates subsite properties';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.description = typeof args.options.description !== 'undefined';
    telemetryProps.siteLogoUrl = typeof args.options.siteLogoUrl !== 'undefined';
    telemetryProps.title = typeof args.options.title !== 'undefined';
    telemetryProps.quickLaunchEnabled = typeof args.options.quickLaunchEnabled !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Updating subsite properties...`);
        }

        const payload: any = {};
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

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${accessToken}`,
            'content-type': 'application/json;odata=nometadata',
            accept: 'application/json;odata=nometadata'
          }),
          json: true,
          body: payload
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        if (this.verbose) {
          cmd.log(`Updating properties of subsite ${args.options.webUrl}...`);
        }

        return request.patch(requestOptions)
      })
      .then((res: any): void => {
        if (this.debug) {
          cmd.log('Response:')
          cmd.log(res);
          cmd.log('');
        }

        if (this.debug) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
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
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.webUrl) {
        return 'Required option webUrl missing';
      }
      else {
        const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }
      }

      if (typeof args.options.quickLaunchEnabled !== 'undefined') {
        if (args.options.quickLaunchEnabled !== 'true' &&
          args.options.quickLaunchEnabled !== 'false') {
          return `${args.options.quickLaunchEnabled} is not a valid boolean value`;
        }
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to a SharePoint Online site,
    using the ${chalk.blue(commands.LOGIN)} command.

  Remarks:
  
    To update subsite properties, you have to first log in to SharePoint
    using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso.sharepoint.com`)}.
  
  Examples:
  
    Update subsite title
      ${chalk.grey(config.delimiter)} ${commands.WEB_SET} --webUrl https://contoso.sharepoint.com/sites/team-a --title Team-a

    Hide quick launch on the subsite
      ${chalk.grey(config.delimiter)} ${commands.WEB_SET} --webUrl https://contoso.sharepoint.com/sites/team-a --quickLaunchEnabled false
  ` );
  }
}

module.exports = new SpoWebSetCommand();
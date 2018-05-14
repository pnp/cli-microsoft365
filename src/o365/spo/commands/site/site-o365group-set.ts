import auth from '../../SpoAuth';
import config from '../../../../config';
import * as request from 'request-promise-native';
import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';
import GlobalOptions from '../../../../GlobalOptions';
import { Auth } from '../../../../Auth';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteUrl: string;
  alias: string;
  displayName: string;
  description?: string;
  classification?: string;
  isPublic?: boolean;
  keepOldHomepage?: boolean;
}

class SpoSiteOffice365GroupSetCommand extends SpoCommand {
  public get name(): string {
    return `${commands.SITE_O365GROUP_SET}`;
  }

  public get description(): string {
    return 'Connects site collection to an Office 365 Group';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.description = typeof args.options.description !== 'undefined';
    telemetryProps.classification = typeof args.options.classification !== 'undefined';
    telemetryProps.isPublic = args.options.isPublic === true;
    telemetryProps.keepOldHomepage = args.options.keepOldHomepage === true;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.siteUrl);

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Connecting site collection to Office 365 Group...`);
        }

        const optionalParams: any = {}
        const payload: any = {
          displayName: args.options.displayName,
          alias: args.options.alias,
          isPublic: args.options.isPublic === true,
          optionalParams: optionalParams
        };

        if (args.options.description) {
          optionalParams.Description = args.options.description;
        }
        if (args.options.classification) {
          optionalParams.Classification = args.options.classification;
        }
        if (args.options.keepOldHomepage) {
          optionalParams.CreationOptions = ["SharePointKeepOldHomepage"];
        }

        const requestOptions: any = {
          url: `${args.options.siteUrl}/_api/GroupSiteManager/CreateGroupForSite`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${accessToken}`,
            'content-type': 'application/json;odata=nometadata',
            accept: 'application/json;odata=nometadata',
            json: true
          }),
          body: payload,
          json: true
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((res: any): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        cmd.log(res);

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --siteUrl <siteUrl>',
        description: 'URL of the site collection being connected to new Office 365 Group'
      },
      {
        option: '-a, --alias <alias>',
        description: 'The email alias for the new Office 365 Group that will be created'
      },
      {
        option: '-n, --displayName <displayName>',
        description: 'The name of the new Office 365 Group that will be created'
      },
      {
        option: '-d, --description [description]',
        description: 'The group’s description'
      },
      {
        option: '-c, --classification [classification]',
        description: 'The classification value, if classifications are set for the organization. If no value is provided, the default classification will be set, if one is configured'
      },
      {
        option: '--isPublic',
        description: 'Determines the Office 365 Group’s privacy setting. If set, the group will be public, otherwise it will be private'
      },
      {
        option: '--keepOldHomepage',
        description: 'For sites that already have a modern page set as homepage, set this option, to keep it as the homepage'
      },
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.siteUrl) {
        return 'Required option siteUrl missing';
      }

      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.siteUrl);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      if (!args.options.alias) {
        return 'Required option alias missing';
      }

      if (!args.options.displayName) {
        return 'Required option displayName missing';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site
    using the ${chalk.blue(commands.CONNECT)} command.
        
  Remarks:

    ${chalk.yellow('Attention:')} This command is based on a SharePoint API that is currently
    in preview and is subject to change once the API reached general
    availability.

    To connect site collection to an Office 365 Group, you have to first connect
    to a SharePoint site using the ${chalk.blue(commands.CONNECT)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.

    When connecting site collection to an Office 365 Group, SharePoint will
    create a new group using the specified information. If a group with the same
    name already exists, you will get a ${chalk.grey('The group alias already exists.')}
    error.

  Examples:
  
    Connect site collection to an Office 365 Group
      ${chalk.grey(config.delimiter)} ${this.name} --siteUrl https://contoso.sharepoin.com/sites/team-a --alias team-a --displayName 'Team A'

    Connect site collection to an Office 365 Group and make the group public
      ${chalk.grey(config.delimiter)} ${this.name} --siteUrl https://contoso.sharepoin.com/sites/team-a --alias team-a --displayName 'Team A' --isPublic

    Connect site collection to an Office 365 Group and set the group classification
      ${chalk.grey(config.delimiter)} ${this.name} --siteUrl https://contoso.sharepoin.com/sites/team-a --alias team-a --displayName 'Team A' --classification HBI

    Connect site collection to an Office 365 Group and keep the old home page
      ${chalk.grey(config.delimiter)} ${this.name} --siteUrl https://contoso.sharepoin.com/sites/team-a --alias team-a --displayName 'Team A' --keepOldHomepage

  More information:

    Overview of the "Connect to new Office 365 group" feature
      https://docs.microsoft.com/en-us/sharepoint/dev/features/groupify/groupify-overview
`);
  }
}

module.exports = new SpoSiteOffice365GroupSetCommand();
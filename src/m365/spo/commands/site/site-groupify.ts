import request from '../../../../request';
import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import GlobalOptions from '../../../../GlobalOptions';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

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

class SpoSiteGroupifyCommand extends SpoCommand {
  public get name(): string {
    return `${commands.SITE_GROUPIFY}`;
  }

  public get description(): string {
    return 'Connects site collection to an Microsoft 365 Group';
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
      headers: {
        'content-type': 'application/json;odata=nometadata',
        accept: 'application/json;odata=nometadata',
        json: true
      },
      body: payload,
      json: true
    };

    request
      .post(requestOptions)
      .then((res: any): void => {
        cmd.log(res);

        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --siteUrl <siteUrl>',
        description: 'URL of the site collection being connected to new Microsoft 365 Group'
      },
      {
        option: '-a, --alias <alias>',
        description: 'The email alias for the new Microsoft 365 Group that will be created'
      },
      {
        option: '-n, --displayName <displayName>',
        description: 'The name of the new Microsoft 365 Group that will be created'
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
        description: 'Determines the Microsoft 365 Group’s privacy setting. If set, the group will be public, otherwise it will be private'
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
      return SpoCommand.isValidSharePointUrl(args.options.siteUrl);
    };
  }
}

module.exports = new SpoSiteGroupifyCommand();
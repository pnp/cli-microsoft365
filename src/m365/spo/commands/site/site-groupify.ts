import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
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
        responseType: 'json'
      },
      data: payload,
      responseType: 'json'
    };

    request
      .post(requestOptions)
      .then((res: any): void => {
        logger.log(res);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --siteUrl <siteUrl>'
      },
      {
        option: '-a, --alias <alias>'
      },
      {
        option: '-n, --displayName <displayName>'
      },
      {
        option: '-d, --description [description]'
      },
      {
        option: '-c, --classification [classification]'
      },
      {
        option: '--isPublic'
      },
      {
        option: '--keepOldHomepage'
      },
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    return SpoCommand.isValidSharePointUrl(args.options.siteUrl);
  }
}

module.exports = new SpoSiteGroupifyCommand();
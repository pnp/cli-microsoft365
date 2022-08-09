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
    return commands.SITE_GROUPIFY;
  }

  public get description(): string {
    return 'Connects site collection to an Microsoft 365 Group';
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
        classification: typeof args.options.classification !== 'undefined',
        isPublic: args.options.isPublic === true,
        keepOldHomepage: args.options.keepOldHomepage === true
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
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
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.siteUrl)
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const optionalParams: any = {};
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
}

module.exports = new SpoSiteGroupifyCommand();
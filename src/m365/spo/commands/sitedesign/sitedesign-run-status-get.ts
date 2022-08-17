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
  runId: string;
  webUrl: string;
}

class SpoSiteDesignRunStatusGetCommand extends SpoCommand {
  public get name(): string {
    return commands.SITEDESIGN_RUN_STATUS_GET;
  }

  public get description(): string {
    return 'Gets information about the site scripts executed for the specified site design';
  }

  public defaultProperties(): string[] | undefined {
    return ['ActionTitle', 'SiteScriptTitle', 'OutcomeText'];
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-i, --runId <runId>'
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

        if (!validation.isValidGuid(args.options.runId)) {
          return `${args.options.runId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const data: any = {
      runId: args.options.runId
    };

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRunStatus`,
      headers: {
        accept: 'application/json;odata=nometadata',
        'content-type': 'application/json;odata=nometadata'
      },
      data: data,
      responseType: 'json'
    };

    request.post<{ value: any[] }>(requestOptions)
      .then((res: { value: any[] }): void => {
        logger.log(res.value);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new SpoSiteDesignRunStatusGetCommand();
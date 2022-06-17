import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { SiteDesignRun } from './SiteDesignRun';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteDesignId?: string;
  webUrl: string;
}

class SpoSiteDesignRunListCommand extends SpoCommand {
  public get name(): string {
    return commands.SITEDESIGN_RUN_LIST;
  }

  public get description(): string {
    return 'Lists information about site designs applied to the specified site';
  }

  public defaultProperties(): string[] | undefined {
    return ['ID', 'SiteDesignID', 'SiteDesignTitle', 'StartTime'];
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
        siteDesignId: typeof args.options.siteDesignId !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-i, --siteDesignId [siteDesignId]'
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

        if (args.options.siteDesignId) {
          if (!validation.isValidGuid(args.options.siteDesignId)) {
            return `${args.options.siteDesignId} is not a valid GUID`;
          }
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const data: any = {};
    if (args.options.siteDesignId) {
      data.siteDesignId = args.options.siteDesignId;
    }

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRun`,
      headers: {
        accept: 'application/json;odata=nometadata',
        'content-type': 'application/json;odata=nometadata'
      },
      data: data,
      responseType: 'json'
    };

    request.post<{ value: SiteDesignRun[] }>(requestOptions)
      .then((res: { value: SiteDesignRun[] }): void => {
        if (args.options.output !== 'json') {
          res.value.forEach(d => {
            d.StartTime = new Date(parseInt(d.StartTime)).toLocaleString();
          });
        }

        logger.log(res.value);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new SpoSiteDesignRunListCommand();
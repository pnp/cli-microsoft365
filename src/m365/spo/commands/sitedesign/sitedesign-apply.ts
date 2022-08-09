import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { spo, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  asTask: boolean;
  id: string;
  webUrl: string;
}

class SpoSiteDesignApplyCommand extends SpoCommand {
  public get name(): string {
    return commands.SITEDESIGN_APPLY;
  }

  public get description(): string {
    return 'Applies a site design to an existing site collection';
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
        asTask: args.options.asTask || false
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id <id>'
      },
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--asTask'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    spo
      .getSpoUrl(logger, this.debug)
      .then((spoUrl: string): Promise<any> => {
        const requestBody: any = {
          siteDesignId: args.options.id,
          webUrl: args.options.webUrl
        };

        const requestOptions: any = {
          url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.${args.options.asTask ? 'AddSiteDesignTask' : 'ApplySiteDesign'}`,
          headers: {
            'content-type': 'application/json;charset=utf-8',
            accept: 'application/json;odata=nometadata'
          },
          data: requestBody,
          responseType: 'json'
        };

        return request.post(requestOptions);
      })
      .then((res: any): void => {
        if (res.value) {
          logger.log(res.value);
        }
        else {
          logger.log(res);
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new SpoSiteDesignApplyCommand();
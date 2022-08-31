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

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoUrl: string = await spo.getSpoUrl(logger, this.debug);
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

      const res: any = await request.post(requestOptions);

      if (res.value) {
        logger.log(res.value);
      }
      else {
        logger.log(res);
      }
    } 
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoSiteDesignApplyCommand();
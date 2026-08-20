import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { getBuiltInSiteDesignId, getBuiltInSiteDesignTemplateNames } from './BuiltInSiteDesigns.js';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  asTask: boolean;
  id?: string;
  template?: string;
  builtIn?: boolean;
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
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        asTask: args.options.asTask || false,
        builtIn: args.options.builtIn || false,
        template: typeof args.options.template !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id [id]'
      },
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--asTask'
      },
      {
        option: '-b, --builtIn'
      },
      {
        option: '--template [template]',
        autocomplete: getBuiltInSiteDesignTemplateNames()
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        if (args.options.template && !getBuiltInSiteDesignTemplateNames().includes(args.options.template)) {
          return `${args.options.template} is not a valid built-in site design template. Allowed values are: ${getBuiltInSiteDesignTemplateNames().join(', ')}`;
        }

        if (args.options.asTask && (args.options.builtIn || args.options.template)) {
          return `The asTask option is not supported when applying a built-in site design`;
        }

        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['id', 'template'] }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoUrl: string = await spo.getSpoUrl(logger, this.debug);
      const isBuiltIn: boolean = !!args.options.builtIn || !!args.options.template;
      const siteDesignId: string = args.options.template
        ? getBuiltInSiteDesignId(args.options.template) as string
        : args.options.id as string;

      const requestBody: any = {
        siteDesignId: siteDesignId,
        webUrl: args.options.webUrl
      };

      if (isBuiltIn) {
        requestBody.store = 1;
      }

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
        await logger.log(res.value);
      }
      else {
        await logger.log(res);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSiteDesignApplyCommand();
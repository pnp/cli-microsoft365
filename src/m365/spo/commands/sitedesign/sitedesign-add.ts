import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { ContextInfo, spo } from '../../../../utils/spo';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  title: string;
  webTemplate: string;
  siteScripts: string;
  description?: string;
  previewImageUrl?: string;
  previewImageAltText?: string;
  isDefault?: boolean;
}

class SpoSiteDesignAddCommand extends SpoCommand {
  public get name(): string {
    return commands.SITEDESIGN_ADD;
  }

  public get description(): string {
    return 'Adds site design for creating modern sites';
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
        webTemplate: args.options.webTemplate,
        numSiteScripts: args.options.siteScripts.split(',').length,
        description: (!(!args.options.description)).toString(),
        previewImageUrl: (!(!args.options.previewImageUrl)).toString(),
        previewImageAltText: (!(!args.options.previewImageAltText)).toString(),
        isDefault: args.options.isDefault || false
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-t, --title <title>'
      },
      {
        option: '-w, --webTemplate <webTemplate>',
        autocomplete: ['TeamSite', 'CommunicationSite']
      },
      {
        option: '-s, --siteScripts <siteScripts>'
      },
      {
        option: '-d, --description [description]'
      },
      {
        option: '-m, --previewImageUrl [previewImageUrl]'
      },
      {
        option: '-a, --previewImageAltText [previewImageAltText]'
      },
      {
        option: '--isDefault'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.webTemplate !== 'TeamSite' &&
          args.options.webTemplate !== 'CommunicationSite') {
          return `${args.options.webTemplate} is not a valid web template type. Allowed values TeamSite|CommunicationSite`;
        }

        const siteScripts = args.options.siteScripts.split(',');
        for (let i: number = 0; i < siteScripts.length; i++) {
          const trimmedId: string = siteScripts[i].trim();
          if (!validation.isValidGuid(trimmedId)) {
            return `${trimmedId} is not a valid GUID`;
          }
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoUrl: string = await spo.getSpoUrl(logger, this.debug);
      const requestDigest: ContextInfo = await spo.getRequestDigest(spoUrl);
      const info: any = {
        Title: args.options.title,
        WebTemplate: args.options.webTemplate === 'TeamSite' ? '64' : '68',
        SiteScriptIds: args.options.siteScripts.split(',').map(i => i.trim())
      };

      if (args.options.description) {
        info.Description = args.options.description;
      }
      if (args.options.previewImageUrl) {
        info.PreviewImageUrl = args.options.previewImageUrl;
      }
      if (args.options.previewImageAltText) {
        info.PreviewImageAltText = args.options.previewImageAltText;
      }
      if (args.options.isDefault) {
        info.IsDefault = true;
      }

      const requestOptions: any = {
        url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign`,
        headers: {
          'X-RequestDigest': requestDigest.FormDigestValue,
          'content-type': 'application/json;charset=utf-8',
          accept: 'application/json;odata=nometadata'
        },
        data: { info: info },
        responseType: 'json'
      };

      const res = await request.post(requestOptions);
      logger.log(res);
    } 
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoSiteDesignAddCommand();

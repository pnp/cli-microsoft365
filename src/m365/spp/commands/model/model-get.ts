import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { spp, SppModel } from '../../../../utils/spp.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteUrl: string;
  id?: string;
  title?: string;
  withPublications?: boolean;
}

class SppModelGetCommand extends SpoCommand {
  public get name(): string {
    return commands.MODEL_GET;
  }

  public get description(): string {
    return 'Retrieves information about a document understanding model';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: typeof args.options.id !== 'undefined',
        title: typeof args.options.title !== 'undefined',
        withPublications: !!args.options.withPublications
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --siteUrl <siteUrl>'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '-t, --title [title]'
      },
      {
        option: '--withPublications'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID for option 'id'.`;
        }

        return validation.isValidSharePointUrl(args.options.siteUrl);
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['id', 'title'] });
  }

  #initTypes(): void {
    this.types.string.push('siteUrl', 'id', 'title');
    this.types.boolean.push('withPublications');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.log(`Retrieving model information from ${args.options.siteUrl}...`);
      }

      const siteUrl = urlUtil.removeTrailingSlashes(args.options.siteUrl);
      await spp.assertSiteIsContentCenter(siteUrl);

      let requestUrl = `${siteUrl}/_api/machinelearning/models/`;

      if (args.options.title) {
        let requestTitle = args.options.title.toLowerCase();

        if (!requestTitle.endsWith('.classifier')) {
          requestTitle += '.classifier';
        }

        requestUrl += `getbytitle('${formatting.encodeQueryParameter(requestTitle)}')`;
      }
      else {
        requestUrl += `getbyuniqueid('${args.options.id}')`;
      }

      const requestOptions: CliRequestOptions = {
        url: requestUrl,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const result = await request.get<SppModel>(requestOptions);

      if ((result as any)['odata.null'] === true) {
        throw 'Model not found.';
      }

      if (args.options.withPublications) {
        result.Publications = await odata.getAllItems<any>(`${siteUrl}/_api/machinelearning/publications/getbymodeluniqueid('${result.UniqueId}')`);
      }

      await logger.log({
        ...result,
        ConfidenceScore: result.ConfidenceScore ? JSON.parse(result.ConfidenceScore) : null,
        Explanations: result.Explanations ? JSON.parse(result.Explanations) : null,
        Schemas: result.Schemas ? JSON.parse(result.Schemas) : null,
        ModelSettings: result.ModelSettings ? JSON.parse(result.ModelSettings) : null
      });
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SppModelGetCommand();
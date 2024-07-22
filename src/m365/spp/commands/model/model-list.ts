import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
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
}

class SppModelListCommand extends SpoCommand {
  public get name(): string {
    return commands.MODEL_LIST;
  }

  public get description(): string {
    return 'Retrieve the list of SharePoint Premium unstructured document processing models on the content center site';
  }

  public defaultProperties(): string[] | undefined {
    return ['AIBuilderHybridModelType', 'ContentTypeName', 'LastTrained', 'UniqueId', 'PublicationType'];
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initTypes();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --siteUrl <siteUrl>'
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('siteUrl');
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        return validation.isValidSharePointUrl(args.options.siteUrl);
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const siteUrl = urlUtil.removeTrailingSlashes(args.options.siteUrl);
      const isContentCenter = await spp.isContentCenter(siteUrl);
      if (!isContentCenter) {
        throw `${siteUrl} is not a content site`;
      }

      const requestOptions: CliRequestOptions = {
        url: `${siteUrl}/_api/machinelearning/models`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const res = await request.get<{ value: SppModel[] }>(requestOptions);
      await logger.log(res.value);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SppModelListCommand();
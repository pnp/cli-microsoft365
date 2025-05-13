import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { odata } from '../../../../utils/odata.js';
import { spp } from '../../../../utils/spp.js';
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
    return 'Retrieves the list of unstructured document processing models';
  }

  public defaultProperties(): string[] | undefined {
    return ['AIBuilderHybridModelType', 'ContentTypeName', 'LastTrained', 'UniqueId'];
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
      if (this.verbose) {
        await logger.log(`Retrieving models from ${args.options.siteUrl}...`);
      }

      const siteUrl = urlUtil.removeTrailingSlashes(args.options.siteUrl);
      await spp.assertSiteIsContentCenter(siteUrl, logger, this.verbose);

      const result = await odata.getAllItems<any>(`${siteUrl}/_api/machinelearning/models`);
      await logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SppModelListCommand();
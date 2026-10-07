import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import { odata } from '../../../../utils/odata.js';
import { spp } from '../../../../utils/spp.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  siteUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
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

  public get schema(): z.ZodType | undefined {
    return options;
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
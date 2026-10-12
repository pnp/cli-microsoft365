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
    .alias('u'),
  id: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: `The value specified for option 'id' is not a valid GUID.`
    })
    .optional()
    .alias('i'),
  title: z.string().optional().alias('t'),
  withPublications: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SppModelGetCommand extends SpoCommand {
  public get name(): string {
    return commands.MODEL_GET;
  }

  public get description(): string {
    return 'Retrieves information about a document understanding model';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.id, opts.title].filter(x => x !== undefined).length === 1, {
        message: `Specify either 'id' or 'title', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['id', 'title']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const siteUrl = urlUtil.removeTrailingSlashes(args.options.siteUrl);
      await spp.assertSiteIsContentCenter(siteUrl, logger, this.verbose);

      let result = null;
      if (args.options.title) {
        result = await spp.getModelByTitle(siteUrl, args.options.title, logger, this.verbose);
      }
      else {
        result = await spp.getModelById(siteUrl, args.options.id!, logger, this.verbose);
      }

      if (args.options.withPublications) {
        if (this.verbose) {
          await logger.log(`Retrieving publications for model...`);
        }
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
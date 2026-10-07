import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
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
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SppModelRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.MODEL_REMOVE;
  }

  public get description(): string {
    return 'Deletes a document understanding model';
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
      if (!args.options.force) {
        const confirmationResult = await cli.promptForConfirmation({ message: `Are you sure you want to remove model '${args.options.title ?? args.options.id}'?` });

        if (!confirmationResult) {
          return;
        }
      }

      if (this.verbose) {
        await logger.log(`Removing model from ${args.options.siteUrl}...`);
      }

      const siteUrl = urlUtil.removeTrailingSlashes(args.options.siteUrl);
      await spp.assertSiteIsContentCenter(siteUrl, logger, this.verbose);
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
          accept: 'application/json;odata=nometadata',
          'if-match': '*'
        },
        responseType: 'json'
      };

      const result = await request.delete<any>(requestOptions);
      if (result?.['odata.null'] === true) {
        throw "Model not found.";
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SppModelRemoveCommand();
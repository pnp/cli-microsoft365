import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  url: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    }).alias('u'),
  audienceIds: z.string()
    .refine(audiences => validation.isValidGuidArray(audiences) === true, {
      error: e => `The following GUIDs are invalid: ${e.input}.`
    }).optional(),
  audienceNames: z.string().optional(),
  vivaConnectionsDefaultStart: z.boolean().optional(),
  isInDraftMode: z.boolean().optional(),
  order: z.number()
    .refine(order => validation.isValidPositiveInteger(order) === true, {
      error: e => `'${e.input}' is not a positive integer.`
    }).optional()
});

declare type Options = z.infer<typeof options>;
interface CommandArgs {
  options: Options;
}

class SpoHomeSiteAddCommand extends SpoCommand {
  public get name(): string {
    return commands.HOMESITE_ADD;
  }

  public get description(): string {
    return 'Adds a home site';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(
        (options: Options) => [options.audienceIds, options.audienceNames].filter(o => o !== undefined).length <= 1,
        {
          message: 'You must specify either audienceIds or audienceNames but not both.'
        }
      );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let audiences: string[] = [];
    if (args.options.audienceIds) {
      audiences = args.options.audienceIds.split(',').map(id => id.trim());
    }
    else if (args.options.audienceNames) {
      audiences = await this.transformAudienceNamesToIds(args.options.audienceNames);
    }

    try {
      const spoAdminUrl: string = await spo.getSpoAdminUrl(logger, this.verbose);
      const requestOptions: CliRequestOptions = {
        url: `${spoAdminUrl}/_api/SPHSite/AddHomeSite`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json',
        data: {
          siteUrl: args.options.url,
          audiences: audiences,
          vivaConnectionsDefaultStart: args.options.vivaConnectionsDefaultStart ?? true,
          isInDraftMode: args.options.isInDraftMode ?? false,
          order: args.options.order
        }
      };

      if (this.verbose) {
        await logger.logToStderr(`Adding home site with URL: ${args.options.url}...`);
      }

      const res = await request.post(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async transformAudienceNamesToIds(audienceNames: string): Promise<string[]> {
    const names = audienceNames.split(',');
    const ids: string[] = [];

    for (const name of names) {
      const id = await entraGroup.getGroupIdByDisplayName(name.trim());
      ids.push(id);
    }

    return ids;
  }
}

export default new SpoHomeSiteAddCommand();
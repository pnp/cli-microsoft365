import { z } from 'zod';
import { zod } from '../../../../utils/zod.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import request, { CliRequestOptions } from '../../../../request.js';

const options = globalOptionsZod
  .extend({
    url: zod.alias('u', z.string()
      .refine((url: string) => validation.isValidSharePointUrl(url) === true, url => ({
        message: `'${url}' is not a valid SharePoint Online site URL.`
      }))
    ),
    audiences: zod.alias('audiences', z.string().optional()
      .refine(audiences => {
        if (audiences === undefined) {
          return true;
        }
        const audienceArray = audiences.split(',');
        return audienceArray.every(audience => validation.isValidGuid(audience));
      }, audiences => ({
        message: `'${audiences}' contains one or more invalid GUIDs.`
      })),
    ),
    vivaConnectionsDefaultStart: z.boolean().optional(),
    isInDraftMode: z.boolean().optional(),
    order: z.number()
      .refine(order => validation.isValidPositiveInteger(order) === true, order => ({
        message: `'${order}' is not a positive integer`
      })).optional()
  })
  .strict();

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

  public get schema(): z.ZodTypeAny {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
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
          audiences: args.options.audiences?.split(','),
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
}

export default new SpoHomeSiteAddCommand();
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  id: z.uuid(),
  fileId: z.uuid().optional(),
  fileUrl: z.string().optional(),
  expirationDateTime: z.string()
    .refine(date => validation.isValidISODateTime(date) === true, {
      error: e => `'${e.input}' is not a valid ISO date string.`
    })
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFileSharingLinkSetCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_SHARINGLINK_SET;
  }

  public get description(): string {
    return 'Updates a sharing link of a file';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.fileId, opts.fileUrl].filter(o => o !== undefined).length === 1, {
        error: `Specify 'fileId' or 'fileUrl', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['fileId', 'fileUrl']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Updating sharing link of file ${args.options.fileId || args.options.fileUrl}...`);
    }

    try {
      const fileDetails = await spo.getVroomFileDetails(args.options.webUrl, args.options.fileId, args.options.fileUrl);

      const requestOptions: CliRequestOptions = {
        url: `https://graph.microsoft.com/v1.0/sites/${fileDetails.SiteId}/drives/${fileDetails.VroomDriveID}/items/${fileDetails.VroomItemID}/permissions/${args.options.id}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: {
          expirationDateTime: args.options.expirationDateTime
        }
      };

      const sharingLink = await request.patch<any>(requestOptions);

      await logger.log(sharingLink);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoFileSharingLinkSetCommand();
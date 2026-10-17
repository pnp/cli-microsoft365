import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import { zod } from '../../../../utils/zod.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

enum SharingLinkType {
  view = 'view',
  edit = 'edit'
}

enum SharingLinkScope {
  anonymous = 'anonymous',
  organization = 'organization'
}

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  fileId: z.uuid().optional(),
  fileUrl: z.string().optional(),
  type: zod.coercedEnum(SharingLinkType),
  expirationDateTime: z.string()
    .refine(date => validation.isValidISODateTime(date) === true, {
      error: e => `'${e.input}' is not a valid ISO date string.`
    })
    .optional(),
  scope: zod.coercedEnum(SharingLinkScope).optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFileSharingLinkAddCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_SHARINGLINK_ADD;
  }

  public get description(): string {
    return 'Creates a new sharing link for a file';
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
      })
      .refine(opts => !opts.scope || opts.scope === 'anonymous' || !opts.expirationDateTime, {
        error: `Option expirationDateTime can only be used for links with scope 'anonymous'.`,
        path: ['expirationDateTime']
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Creating a sharing link for file ${args.options.fileId || args.options.fileUrl}...`);
    }

    try {
      const fileDetails = await spo.getVroomFileDetails(args.options.webUrl, args.options.fileId, args.options.fileUrl);

      const requestOptions: CliRequestOptions = {
        url: `https://graph.microsoft.com/v1.0/sites/${fileDetails.SiteId}/drives/${fileDetails.VroomDriveID}/items/${fileDetails.VroomItemID}/createLink`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: {
          type: args.options.type,
          expirationDateTime: args.options.expirationDateTime,
          scope: args.options.scope
        }
      };

      const sharingLink = await request.post<any>(requestOptions);

      await logger.log(sharingLink);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoFileSharingLinkAddCommand();

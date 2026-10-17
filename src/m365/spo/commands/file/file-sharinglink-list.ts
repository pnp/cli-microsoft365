import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import { zod } from '../../../../utils/zod.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

enum SharingLinkScope {
  anonymous = 'anonymous',
  users = 'users',
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
  scope: zod.coercedEnum(SharingLinkScope).optional().alias('s')
});

export declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFileSharingLinkListCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_SHARINGLINK_LIST;
  }

  public get description(): string {
    return 'Lists all the sharing links of a specific file';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'scope', 'roles', 'link'];
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
      await logger.logToStderr(`Retrieving sharing links for file ${args.options.fileId || args.options.fileUrl}...`);
    }

    try {
      const fileDetails = await spo.getVroomFileDetails(args.options.webUrl, args.options.fileId, args.options.fileUrl);
      let url = `https://graph.microsoft.com/v1.0/sites/${fileDetails.SiteId}/drives/${fileDetails.VroomDriveID}/items/${fileDetails.VroomItemID}/permissions?$filter=Link ne null`;
      if (args.options.scope) {
        url += ` and Link/Scope eq '${args.options.scope}'`;
      }

      const sharingLinks = await odata.getAllItems<any>(url);

      if (!args.options.output || !cli.shouldTrimOutput(args.options.output)) {
        await logger.log(sharingLinks);
      }
      else {
        //converted to text friendly output
        await logger.log(sharingLinks.map(i => {
          return {
            id: i.id,
            roles: i.roles.join(','),
            link: i.link.webUrl,
            scope: i.link.scope
          };
        }));
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoFileSharingLinkListCommand();
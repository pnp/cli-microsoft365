import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
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
  fileUrl: z.string().optional(),
  fileId: z.uuid().optional(),
  id: z.string().alias('i'),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFileSharingLinkRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_SHARINGLINK_REMOVE;
  }

  public get description(): string {
    return 'Removes a specific sharing link of a file';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.fileUrl, opts.fileId].filter(o => o !== undefined).length === 1, {
        error: `Specify 'fileUrl' or 'fileId', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['fileUrl', 'fileId']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeSharingLink = async (): Promise<void> => {
      try {
        if (this.verbose) {
          await logger.logToStderr(`Removing sharing link of file ${args.options.fileUrl || args.options.fileId} with id ${args.options.id}...`);
        }

        const fileDetails = await spo.getVroomFileDetails(args.options.webUrl, args.options.fileId, args.options.fileUrl);
        const requestOptions: CliRequestOptions = {
          url: `https://graph.microsoft.com/v1.0/sites/${fileDetails.SiteId}/drives/${fileDetails.VroomDriveID}/items/${fileDetails.VroomItemID}/permissions/${args.options.id}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeSharingLink();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove sharing link ${args.options.id} of file ${args.options.fileUrl || args.options.fileId}?` });

      if (result) {
        await removeSharingLink();
      }
    }
  }
}

export default new SpoFileSharingLinkRemoveCommand();
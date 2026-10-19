import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import Command, { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import { zod } from '../../../../utils/zod.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import spoFileSharingLinkListCommand, { Options as SpoFileSharingLinkListOptions } from './file-sharinglink-list.js';

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
  fileUrl: z.string().optional(),
  fileId: z.uuid().optional(),
  scope: zod.coercedEnum(SharingLinkScope).optional().alias('s'),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFileSharingLinkClearCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_SHARINGLINK_CLEAR;
  }

  public get description(): string {
    return 'Removes sharing links of a file';
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
    const clearSharingLinks: () => Promise<void> = async (): Promise<void> => {
      try {
        if (this.verbose) {
          await logger.logToStderr(`Clearing sharing links for file ${args.options.fileUrl || args.options.fileId}${args.options.scope ? ` with scope ${args.options.scope}` : ''}`);
        }

        const fileDetails = await spo.getVroomFileDetails(args.options.webUrl, args.options.fileId, args.options.fileUrl);
        const sharingLinks = await this.getFileSharingLinks(args.options.webUrl, args.options.fileId, args.options.fileUrl, args.options.scope);

        const requestOptions: CliRequestOptions = {
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        for (const sharingLink of sharingLinks) {
          requestOptions.url = `https://graph.microsoft.com/v1.0/sites/${fileDetails.SiteId}/drives/${fileDetails.VroomDriveID}/items/${fileDetails.VroomItemID}/permissions/${sharingLink.id}`;
          await request.delete(requestOptions);
        }
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await clearSharingLinks();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to clear the sharing links of file ${args.options.fileUrl || args.options.fileId}${args.options.scope ? ` with scope ${args.options.scope}` : ''}?` });

      if (result) {
        await clearSharingLinks();
      }
    }
  }

  private async getFileSharingLinks(webUrl: string, fileId?: string, fileUrl?: string, scope?: SpoFileSharingLinkListOptions['scope']): Promise<any[]> {
    const sharingLinkListOptions: SpoFileSharingLinkListOptions = {
      webUrl: webUrl,
      fileId: fileId,
      fileUrl: fileUrl,
      scope: scope,
      debug: this.debug,
      verbose: this.verbose
    };

    const commandOutput = await cli.executeCommandWithOutput(spoFileSharingLinkListCommand as Command, { options: { ...sharingLinkListOptions, _: [] } });
    const outputParsed = JSON.parse(commandOutput.stdout);
    return outputParsed;
  }
}

export default new SpoFileSharingLinkClearCommand();

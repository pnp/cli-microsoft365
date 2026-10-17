import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { FileProperties } from './FileProperties.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  fileUrl: z.string().optional(),
  fileId: z.uuid().optional().alias('i'),
  clearExistingPermissions: z.boolean().optional().alias('c'),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFileRoleInheritanceBreakCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_ROLEINHERITANCE_BREAK;
  }

  public get description(): string {
    return 'Breaks inheritance of a file. Keeping existing permissions is the default behavior.';
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
    const breakFileRoleInheritance = async (): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr(`Breaking role inheritance for file ${args.options.fileId || args.options.fileUrl}`);
      }
      try {
        const fileURL: string = await this.getFileURL(args, logger);

        const keepExistingPermissions: boolean = !args.options.clearExistingPermissions;

        const requestOptions: CliRequestOptions = {
          url: `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(fileURL)}')/ListItemAllFields/breakroleinheritance(${keepExistingPermissions})`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        await request.post(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await breakFileRoleInheritance();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to break the role inheritance of file ${args.options.fileUrl || args.options.fileId} located in site ${args.options.webUrl}?` });

      if (result) {
        await breakFileRoleInheritance();
      }
    }
  }

  private async getFileURL(args: CommandArgs, logger: Logger): Promise<string> {
    if (args.options.fileUrl) {
      return urlUtil.getServerRelativePath(args.options.webUrl, args.options.fileUrl);
    }

    const file: FileProperties = await spo.getFileById(args.options.webUrl, args.options.fileId!, logger, this.verbose);
    return file.ServerRelativeUrl;
  }
}

export default new SpoFileRoleInheritanceBreakCommand();

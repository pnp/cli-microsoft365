import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
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
  folderUrl: z.string(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFolderRoleInheritanceResetCommand extends SpoCommand {
  public get name(): string {
    return commands.FOLDER_ROLEINHERITANCE_RESET;
  }

  public get description(): string {
    return 'Restores the role inheritance of a folder';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const serverRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.folderUrl);
    const roleFolderUrl: string = urlUtil.getWebRelativePath(args.options.webUrl, args.options.folderUrl);
    let requestUrl: string = `${args.options.webUrl}/_api/web/`;

    const resetFolderRoleInheritance = async (): Promise<void> => {
      try {
        if (roleFolderUrl.split('/').length === 2) {
          requestUrl += `GetList('${formatting.encodeQueryParameter(serverRelativeUrl)}')`;
        }
        else {
          requestUrl += `GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativeUrl)}')/ListItemAllFields`;
        }
        const requestOptions: CliRequestOptions = {
          url: `${requestUrl}/resetroleinheritance`,
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
      await resetFolderRoleInheritance();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to reset the role inheritance of folder ${args.options.folderUrl} located in site ${args.options.webUrl}?` });

      if (result) {
        await resetFolderRoleInheritance();
      }
    }
  }
}

export default new SpoFolderRoleInheritanceResetCommand();
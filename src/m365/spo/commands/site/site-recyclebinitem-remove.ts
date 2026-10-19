import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
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
  ids: z.string()
    .refine(ids => validation.isValidGuidArray(ids) === true, {
      error: e => `The following GUIDs are invalid for the option 'ids': ${validation.isValidGuidArray(e.input as string)}.`
    })
    .alias('i'),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoSiteRecycleBinItemRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_RECYCLEBINITEM_REMOVE;
  }

  public get description(): string {
    return 'Permanently deletes specific items from the site recycle bin';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removeRecycleBinItem(args, logger);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to permanently delete ${args.options.ids.split(',').length} item(s) from the site recycle bin?` });

      if (result) {
        await this.removeRecycleBinItem(args, logger);
      }
    }
  }

  private async removeRecycleBinItem(args: CommandArgs, logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Permanently deleting items from the site recycle bin at site ${args.options.siteUrl}...`);
    }

    try {
      const requestOptions: CliRequestOptions = {
        url: `${args.options.siteUrl}/_api/site/RecycleBin/DeleteByIds`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json',
        data: {
          ids: args.options.ids.split(',')
        }
      };

      await request.post<any>(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSiteRecycleBinItemRemoveCommand();
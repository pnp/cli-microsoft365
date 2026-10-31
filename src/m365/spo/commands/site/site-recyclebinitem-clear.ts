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
  secondary: z.boolean().optional(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoSiteRecycleBinItemClearCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_RECYCLEBINITEM_CLEAR;
  }

  public get description(): string {
    return 'Permanently removes all items in a site recycle bin';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.clearRecycleBin(args, logger);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to clear the recycle bin of site ${args.options.siteUrl}?` });

      if (result) {
        await this.clearRecycleBin(args, logger);
      }
    }
  }

  private async clearRecycleBin(args: CommandArgs, logger: Logger): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Permanently removing all items in recycle bin of site ${args.options.siteUrl}...`);
      }

      const requestOptions: CliRequestOptions = {
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      if (args.options.secondary) {
        if (this.verbose) {
          await logger.logToStderr('Removing all items from the second-stage recycle bin');
        }
        requestOptions.url = `${args.options.siteUrl}/_api/site/RecycleBin/DeleteAllSecondStageItems`;
      }
      else {
        if (this.verbose) {
          await logger.logToStderr('Removing all items from the first-stage recycle bin');
        }
        requestOptions.url = `${args.options.siteUrl}/_api/web/RecycleBin/DeleteAll`;
      }

      const result = await request.post<any>(requestOptions);
      if (result['odata.null'] !== true) {
        throw result;
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSiteRecycleBinItemClearCommand();
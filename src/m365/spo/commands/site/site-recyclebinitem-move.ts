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
    .optional()
    .alias('i'),
  all: z.boolean().optional(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoSiteRecycleBinItemMoveCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_RECYCLEBINITEM_MOVE;
  }

  public get description(): string {
    return 'Moves items from the first-stage recycle bin to the second-stage recycle bin';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(o => [o.ids, o.all].filter(x => x !== undefined).length === 1, {
        error: `Specify either 'ids' or 'all', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['ids', 'all']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.moveRecycleBinItem(args, logger);
    }
    else {
      const result = await cli.promptForConfirmation({ message: 'Are you sure you want to move these items to the second-stage recycle bin?' });

      if (result) {
        await this.moveRecycleBinItem(args, logger);
      }
    }
  }

  private async moveRecycleBinItem(args: CommandArgs, logger: Logger): Promise<void> {
    try {
      const requestOptions: CliRequestOptions = {
        url: `${args.options.siteUrl}/_api/web/recycleBin/MoveAllToSecondStage`,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      if (args.options.ids !== undefined) {
        const ids = args.options.ids.split(',');
        if (this.verbose) {
          await logger.logToStderr(`Moving ${ids.length} items to the second-stage recycle bin.`);
        }

        requestOptions.data = {
          ids: ids
        };
      }
      else if (this.verbose) {
        await logger.logToStderr('Moving all items to the second-stage recycle bin.');
      }

      await request.post<any>(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSiteRecycleBinItemMoveCommand();
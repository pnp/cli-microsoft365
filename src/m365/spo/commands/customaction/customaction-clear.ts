import chalk from 'chalk';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { z } from 'zod';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string().refine(val => validation.isValidSharePointUrl(val) === true, {
    error: e => `${e.input} is not a valid SharePoint Online site URL.`
  }).alias('u'),
  scope: z.enum(['Site', 'Web', 'All']).optional().alias('s'),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoCustomActionClearCommand extends SpoCommand {
  public get name(): string {
    return commands.CUSTOMACTION_CLEAR;
  }

  public get description(): string {
    return 'Deletes all custom actions in the collection';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const clearCustomActions = async (): Promise<void> => {
      try {
        if (args.options.scope && args.options.scope.toLowerCase() !== "all") {
          await this.clearScopedCustomActions(args.options);
        }
        else {
          await this.clearAllScopes(args.options);
        }
      }
      catch (err: any) {
        this.handleRejectedPromise(err);
      }
    };

    if (args.options.force) {
      await clearCustomActions();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to clear all the user custom actions with scope ${chalk.yellow(args.options.scope || 'All')}?` });

      if (result) {
        await clearCustomActions();
      }
    }
  }

  private clearScopedCustomActions(options: Options): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${options.webUrl}/_api/${options.scope}/UserCustomActions/clear`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request.post(requestOptions);
  }

  /**
   * Clear request with `web` scope is send first. 
   * Another clear request is send with `site` scope after.
   */
  private async clearAllScopes(options: Options): Promise<void> {
    options.scope = "Web";

    await this.clearScopedCustomActions(options);

    options.scope = "Site";

    await this.clearScopedCustomActions(options);
  }
}

export default new SpoCustomActionClearCommand();
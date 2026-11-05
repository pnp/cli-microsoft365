import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u')
});
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoRoleDefinitionListCommand extends SpoCommand {
  public get name(): string {
    return commands.ROLEDEFINITION_LIST;
  }

  public get description(): string {
    return 'Gets list of role definitions for the specified site';
  }

  public defaultProperties(): string[] | undefined {
    return ['Id', 'Name'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Getting role definitions list from ${args.options.webUrl}...`);
    }

    try {
      const res = await odata.getAllItems<any>(`${args.options.webUrl}/_api/web/roledefinitions`);
      const response = formatting.setFriendlyPermissions(res);
      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoRoleDefinitionListCommand();
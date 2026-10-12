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
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  id: z.coerce.number()
    .alias('i'),
  force: z.boolean().optional()
});
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoRoleDefinitionRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.ROLEDEFINITION_REMOVE;
  }

  public get description(): string {
    return 'Removes the role definition from the specified site';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removeRoleDefinition(logger, args);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the role definition with id ${args.options.id} from site ${args.options.webUrl}?` });

      if (result) {
        await this.removeRoleDefinition(logger, args);
      }
    }
  }

  private async removeRoleDefinition(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Removing role definition from site ${args.options.webUrl}...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${args.options.webUrl}/_api/web/roledefinitions(${args.options.id})`,
      method: 'delete',
      headers: {
        'X-HTTP-Method': 'DELETE',
        'If-Match': '*',
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      await request.delete(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoRoleDefinitionRemoveCommand();
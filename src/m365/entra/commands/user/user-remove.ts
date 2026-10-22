import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import commands from '../../commands.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import { cli } from '../../../../cli/cli.js';
import GraphCommand from '../../../base/GraphCommand.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.uuid().optional(),
  userName: z.string().refine(name => validation.isValidUserPrincipalName(name), {
    error: e => `'${e.input}' is not a valid user principal name (UPN).`
  }).optional(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraUserRemoveCommand extends GraphCommand {

  public get name(): string {
    return commands.USER_REMOVE;
  }

  public get description(): string {
    return 'Removes a specific user';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.id, options.userName].filter(o => o !== undefined).length === 1, {
        error: `Specify either 'id' or 'userName'.`,
        params: {
          customCode: 'optionSet',
          options: ['id', 'userName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Removing user '${args.options.id || args.options.userName}'...`);
    }

    if (args.options.force) {
      await this.deleteUser(args);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove user '${args.options.id || args.options.userName}'?` });

      if (result) {
        await this.deleteUser(args);
      }
    }
  }

  private async deleteUser(args: CommandArgs): Promise<void> {
    try {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/users/${args.options.id || args.options.userName}`,
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
  }
}

export default new EntraUserRemoveCommand();
import { Logger } from '../../../cli/Logger.js';
import auth from '../../../Auth.js';
import commands from '../commands.js';
import Command, { CommandError, globalOptionsZod } from '../../../Command.js';
import z from 'zod';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  name: z.string().alias('n')
    .refine(async (name) => (await auth.getAllConnections()).some(c => c.name === name), {
      error: e => `Connection with name '${e.input}' does not exist.`
    }),
  newName: z.string()
    .refine(async (newName) => !(await auth.getAllConnections()).some(c => c.name === newName), {
      error: e => `Connection with name '${e.input}' already exists.`
    })
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class ConnectionSetCommand extends Command {
  public get name(): string {
    return commands.SET;
  }

  public get description(): string {
    return 'Rename the specified connection';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const connection = await auth.getConnection(args.options.name);

    if (this.verbose) {
      await logger.logToStderr(`Updating connection '${connection.identityName}', appId: ${connection.appId}, tenantId: ${connection.identityTenantId}...`);
    }

    await auth.updateConnection(connection, args.options.newName);
  }

  public async action(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      await auth.restoreAuth();
    }
    catch (error: any) {
      throw new CommandError(error);
    }

    await this.initAction(args, logger);
    await this.commandAction(logger, args);
  }
}

export default new ConnectionSetCommand();
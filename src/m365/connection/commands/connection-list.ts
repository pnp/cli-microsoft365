import assert from 'assert';
import { z } from 'zod';
import auth from '../../../Auth.js';
import { Logger } from '../../../cli/Logger.js';
import Command, { CommandArgs, CommandError, globalOptionsZod } from '../../../Command.js';
import commands from '../commands.js';

const options = globalOptionsZod.strict();

class ConnectionListCommand extends Command {
  public get name(): string {
    return commands.LIST;
  }

  public get description(): string {
    return 'Show the list of available connections';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'connectedAs', 'authType', 'active'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    const availableConnections = await auth.getAllConnections();

    const output = availableConnections.map(connection => {
      const isCurrentConnection = connection.name === auth.connection?.name;

      return {
        name: connection.name,
        connectedAs: connection.identityName,
        authType: connection.authType,
        active: isCurrentConnection
      };
    }).sort((a, b) => {

      // Asserting name because it is optional, but required at this point.
      assert(a.name !== undefined);
      assert(b.name !== undefined);

      const aName = a.name;
      const bName = b.name;

      return aName.localeCompare(bName);
    });

    await logger.log(output);
  }

  public async action(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      await auth.restoreAuth();
    }
    catch (error: any) {
      throw new CommandError(error);
    }

    await this.initAction(args, logger);
    await this.commandAction(logger);
  }
}

export default new ConnectionListCommand();
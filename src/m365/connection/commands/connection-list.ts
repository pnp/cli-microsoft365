import { Logger } from '../../../cli/Logger.js';
import auth, { AuthType } from '../../../Auth.js';
import commands from '../commands.js';
import Command, { CommandArgs, CommandError } from '../../../Command.js';

class ConnectionListCommand extends Command {
  public get name(): string {
    return commands.LIST;
  }

  public get description(): string {
    return 'Show the list of available connections';
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
        authType: AuthType[connection.authType],
        active: isCurrentConnection
      };
    }).sort((a, b) => a.name!.localeCompare(b.name!));

    await logger.log(output);
  }

  public async action(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      await auth.restoreAuth();
    }
    catch (error: any) {
      throw new CommandError(error);
    }

    this.initAction(args, logger);
    await this.commandAction(logger);
  }
}

export default new ConnectionListCommand();
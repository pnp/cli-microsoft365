import { Logger } from '../../../cli/Logger.js';
import auth, { AuthType } from '../../../Auth.js';
import commands from '../commands.js';
import Command, { CommandArgs, CommandError } from '../../../Command.js';

class ConnectionListCommand extends Command {
  public get name(): string {
    return commands.LIST;
  }

  public get description(): string {
    return "Show the list of available connections";
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'connectedAs', 'authType'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    const availableConnections = await auth.getAllConnections();

    await logger.log(availableConnections.map(i => {
      const isCurrentConnection = auth.connection?.name === i.name;

      return {
        name: i.name,
        connectedAs: i.identityName,
        authType: AuthType[i.authType],
        active: isCurrentConnection
      };
    }));
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
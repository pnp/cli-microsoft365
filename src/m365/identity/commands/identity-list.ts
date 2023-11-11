import { Logger } from '../../../cli/Logger.js';
import auth from '../../../Auth.js';
import commands from '../commands.js';
import Command, { CommandArgs, CommandError } from '../../../Command.js';

class IdentityListCommand extends Command {
  public get name(): string {
    return commands.LIST;
  }

  public get description(): string {
    return "Shows a list of currently signed in identities";
  }

  public defaultProperties(): string[] | undefined {
    return ['identityName', 'authType'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    const availableConnections = await auth.getAllConnections();

    await logger.log(availableConnections.map(i => auth.getIdentityDetails(i, this.debug)));
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

export default new IdentityListCommand();
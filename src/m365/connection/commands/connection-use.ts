import { Logger } from '../../../cli/Logger.js';
import auth from '../../../Auth.js';
import commands from '../commands.js';
import Command, { CommandError } from '../../../Command.js';
import GlobalOptions from '../../../GlobalOptions.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
}

class ConnectionUseCommand extends Command {
  public get name(): string {
    return commands.USE;
  }

  public get description(): string {
    return "When signed in with multiple identities, switch to another connection";
  }

  constructor() {
    super();

    this.#initOptions();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name <name>'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const connection = await auth.getConnection(args.options.name);

    if (this.verbose) {
      await logger.logToStderr(`Switching to connection '${connection.identityName}', appId: ${connection.appId}, tenantId: ${connection.identityTenantId}...`);
    }

    await auth.switchToConnection(connection);

    await logger.log(auth.getIdentityDetails(auth.connection, this.debug));
  }

  public async action(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      await auth.restoreAuth();
    }
    catch (error: any) {
      throw new CommandError(error);
    }

    this.initAction(args, logger);
    await this.commandAction(logger, args);
  }
}

export default new ConnectionUseCommand();
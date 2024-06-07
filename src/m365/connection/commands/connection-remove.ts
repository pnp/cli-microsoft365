import { Logger } from '../../../cli/Logger.js';
import auth from '../../../Auth.js';
import commands from '../commands.js';
import Command, { CommandError } from '../../../Command.js';
import GlobalOptions from '../../../GlobalOptions.js';
import { cli } from '../../../cli/cli.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  force?: boolean;
}

class ConnectionRemoveCommand extends Command {
  public get name(): string {
    return commands.REMOVE;
  }

  public get description(): string {
    return 'Remove the specified connection';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }


  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        force: (!(!args.options.force)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name <name>'
      },
      {
        option: '-f, --force'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const deleteConnection = async (): Promise<void> => {
      const connection = await auth.getConnection(args.options.name);

      if (this.verbose) {
        await logger.logToStderr(`Removing connection '${connection.identityName}', appId: ${connection.appId}, tenantId: ${connection.identityTenantId}...`);
      }

      await auth.removeConnectionInfo(connection, logger, this.debug);
    };

    if (args.options.force) {
      await deleteConnection();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the connection?` });

      if (result) {
        await deleteConnection();
      }
    }
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

export default new ConnectionRemoveCommand();
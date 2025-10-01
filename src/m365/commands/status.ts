import { z } from 'zod';
import auth from '../../Auth.js';
import { Logger } from '../../cli/Logger.js';
import Command, { CommandArgs, CommandError, globalOptionsZod } from '../../Command.js';
import commands from './commands.js';

const options = globalOptionsZod.strict();


class StatusCommand extends Command {
  public get name(): string {
    return commands.STATUS;
  }

  public get description(): string {
    return 'Shows Microsoft 365 login status';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public async commandAction(logger: Logger): Promise<void> {
    if (auth.connection.active) {
      try {
        await auth.ensureAccessToken(auth.defaultResource, logger, this.debug);
      }
      catch (err: any) {
        if (this.debug) {
          await logger.logToStderr(err);
        }

        auth.connection.deactivate();
        throw new CommandError(`Your login has expired. Sign in again to continue. ${err.message}`);
      }

      const details = auth.getConnectionDetails(auth.connection);

      if (this.debug) {
        (details as any).accessTokens = JSON.stringify(auth.connection.accessTokens, null, 2);
      }

      await logger.log(details);
    }
    else {
      const connections = await auth.getAllConnections();
      if (this.verbose) {
        const message = connections.length > 0
          ? `Logged out, signed in connections available`
          : 'Logged out';
        await logger.logToStderr(message);
      }
      else {
        const message = connections.length > 0
          ? `Logged out, signed in connections available`
          : 'Logged out';
        await logger.log(message);
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
    await this.commandAction(logger);
  }
}

export default new StatusCommand();
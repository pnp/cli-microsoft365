import auth from '../../Auth.js';
import { Logger } from '../../cli/Logger.js';
import Command, { CommandArgs, CommandError } from '../../Command.js';
import commands from './commands.js';

class StatusCommand extends Command {
  public get name(): string {
    return commands.STATUS;
  }

  public get description(): string {
    return 'Shows Microsoft 365 login status';
  }

  public async commandAction(logger: Logger): Promise<void> {
    if (auth.service.active) {
      try {
        await auth.ensureAccessToken(auth.defaultResource, logger, this.debug);
      }
      catch (err: any) {
        if (this.debug) {
          await logger.logToStderr(err);
        }

        auth.service.deactivateConnection();
        throw new CommandError(`Your login has expired. Sign in again to continue. ${err.message}`);
      }

      const response = {
        ...auth.getIdentityDetails(auth.service, this.debug),
        connectedAs: auth.service.identityName // (Deprecated), added for backwards compatibility
      };

      if (this.debug) {
        await logger.logToStderr(response);
      }
      else {
        await logger.log(response);
      }
    }
    else {
      const connections = await auth.getAllConnections();
      if (this.verbose) {
        const message = connections.length > 0
          ? `Logged out from Microsoft 365, signed in identities available`
          : 'Logged out from Microsoft 365';
        await logger.logToStderr(message);
      }
      else {
        const message = connections.length > 0
          ? `Logged out, signed in identities available`
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

    this.initAction(args, logger);
    await this.commandAction(logger);
  }
}

export default new StatusCommand();
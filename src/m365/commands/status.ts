import auth, { AuthType } from '../../Auth';
import { Logger } from '../../cli/Logger';
import Command, { CommandArgs, CommandError } from '../../Command';
import { accessToken } from '../../utils/accessToken';
import commands from './commands';

class StatusCommand extends Command {
  public get name(): string {
    return commands.STATUS;
  }

  public get description(): string {
    return 'Shows Microsoft 365 login status';
  }

  public async commandAction(logger: Logger): Promise<void> {
    if (auth.service.connected) {
      try {
        await auth.ensureAccessToken(auth.defaultResource, logger, this.debug);
      }
      catch (err: any) {
        if (this.debug) {
          logger.logToStderr(err);
        }

        auth.service.logout();
        throw new CommandError(`Your login has expired. Sign in again to continue. ${err.message}`);
      }
      
      if (this.debug) {
        logger.logToStderr({
          connectedAs: accessToken.getUserNameFromAccessToken(auth.service.accessTokens[auth.defaultResource].accessToken),
          authType: AuthType[auth.service.authType],
          accessTokens: JSON.stringify(auth.service.accessTokens, null, 2)
        });
      }
      else {
        logger.log({
          connectedAs: accessToken.getUserNameFromAccessToken(auth.service.accessTokens[auth.defaultResource].accessToken)
        });
      }
    }
    else {
      if (this.verbose) {
        logger.logToStderr('Logged out from Microsoft 365');
      }
      else {
        logger.log('Logged out');
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

module.exports = new StatusCommand();
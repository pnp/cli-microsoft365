import auth, { AuthType } from '../../Auth';
import { Logger } from '../../cli';
import Command, { CommandArgs, CommandError } from '../../Command';
import { accessToken } from '../../utils';
import commands from './commands';

class StatusCommand extends Command {
  public get name(): string {
    return commands.STATUS;
  }

  public get description(): string {
    return 'Shows Microsoft 365 login status';
  }

  public commandAction(logger: Logger, args: any, cb: (err?: any) => void): void {

    if (auth.service.connected) {
      auth
        .ensureAccessToken(auth.defaultResource, logger, this.debug)
        .then((): void => {
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
          cb();
        }, (rej: Error): void => {
          if (this.debug) {
            logger.logToStderr(rej);
          }

          logger.log('Your login has expired. Sign in again to continue.');
          auth.service.logout();
          cb(new CommandError(rej.message));
        });
    }
    else {
      if (this.verbose) {
        logger.logToStderr('Logged out from Microsoft 365');
      }
      else {
        logger.log('Logged out');
      }
      cb();
    }
  }

  public action(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    auth
      .restoreAuth()
      .then((): void => {
        this.initAction(args, logger);
        this.commandAction(logger, args, cb);
      }, (error: any): void => {
        cb(new CommandError(error));
      });
  }
}

module.exports = new StatusCommand();
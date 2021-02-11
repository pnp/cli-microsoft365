import auth, { AuthType } from '../../Auth';
import { Logger } from '../../cli';
import Command, { CommandArgs, CommandError } from '../../Command';
import Utils from '../../Utils';
import commands from './commands';

class StatusCommand extends Command {
  public get name(): string {
    return commands.STATUS;
  }

  public get description(): string {
    return 'Shows Microsoft 365 login status';
  }

  public commandAction(logger: Logger, args: {}, cb: (err?: any) => void): void {
    if (auth.service.connected) {
      if (this.debug) {
        logger.logToStderr({
          connectedAs: Utils.getUserNameFromAccessToken(auth.service.accessTokens[auth.defaultResource].accessToken),
          authType: AuthType[auth.service.authType],
          accessTokens: JSON.stringify(auth.service.accessTokens, null, 2)
        });
      }
      else {
        logger.log({
          connectedAs: Utils.getUserNameFromAccessToken(auth.service.accessTokens[auth.defaultResource].accessToken)
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
    cb();
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
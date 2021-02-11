import auth from '../../Auth';
import { Logger } from '../../cli';
import Command, { CommandArgs, CommandError } from '../../Command';
import commands from './commands';

class LogoutCommand extends Command {
  public get name(): string {
    return commands.LOGOUT;
  }

  public get description(): string {
    return 'Log out from Microsoft 365';
  }

  public commandAction(logger: Logger, args: {}, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr('Logging out from Microsoft 365...');
    }

    const logout: () => void = (): void => {
      auth.service.logout();
      cb();
    }

    auth
      .clearConnectionInfo(logger, this.debug)
      .then((): void => {
        logout();
      }, (error: any): void => {
        if (this.debug) {
          logger.logToStderr(new CommandError(error));
        }

        logout();
      });
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

module.exports = new LogoutCommand();
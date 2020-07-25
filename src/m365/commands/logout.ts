import auth from '../../Auth';
import commands from './commands';
import Command, {
  CommandError, CommandAction, CommandArgs,
} from '../../Command';
import * as chalk from 'chalk';
import { CommandInstance } from '../../cli';

class LogoutCommand extends Command {
  public get name(): string {
    return commands.LOGOUT;
  }

  public get description(): string {
    return 'Log out from Microsoft 365';
  }

  public commandAction(cmd: CommandInstance, args: {}, cb: () => void): void {
    if (this.verbose) {
      cmd.log('Logging out from Microsoft 365...');
    }

    const logout: () => void = (): void => {
      auth.service.logout();
      if (this.verbose) {
        cmd.log(chalk.green('DONE'));
      }
      cb();
    }

    auth
      .clearConnectionInfo()
      .then((): void => {
        logout();
      }, (error: any): void => {
        if (this.debug) {
          cmd.log(new CommandError(error));
        }

        logout();
      });
  }

  public action(): CommandAction {
    const cmd: Command = this;
    return function (this: CommandInstance, args: CommandArgs, cb: (err?: any) => void) {
      auth
        .restoreAuth()
        .then((): void => {
          (cmd as any).initAction(args, this);
          cmd.commandAction(this, args, cb);
        }, (error: any): void => {
          cb(new CommandError(error));
        });
    }
  }
}

module.exports = new LogoutCommand();
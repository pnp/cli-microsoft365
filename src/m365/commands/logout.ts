import auth from '../../Auth';
import commands from './commands';
import Command, {
  CommandError, CommandAction, CommandArgs,
} from '../../Command';

const vorpal: Vorpal = require('../../vorpal-init');

class LogoutCommand extends Command {
  public get name(): string {
    return commands.LOGOUT;
  }

  public get description(): string {
    return 'Log out from Microsoft 365';
  }

  public commandAction(cmd: CommandInstance, args: {}, cb: () => void): void {
    const chalk = vorpal.chalk;

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
          args = (cmd as any).processArgs(args);
          (cmd as any).initAction(args, this);

          cmd.commandAction(this, args, cb);
        }, (error: any): void => {
          cb(new CommandError(error));
        });
    }
  }

  public commandHelp(args: any, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.LOGOUT).helpInformation());
    log(
      `  Remarks:

    The ${chalk.blue(commands.LOGOUT)} command logs out from Microsoft 365
    and removes any access and refresh tokens from memory.

  Examples:
  
    Log out from Microsoft 365
      ${commands.LOGOUT}

    Log out from Microsoft 365 in debug mode including detailed debug
    information in the console output
      ${commands.LOGOUT} --debug
`);
  }
}

module.exports = new LogoutCommand();
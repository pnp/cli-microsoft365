import auth from '../../Auth';
import commands from './commands';
import Command, {
  CommandError, CommandAction, CommandArgs
} from '../../Command';
import Utils from '../../Utils';
import { AuthType } from '../../Auth';

const vorpal: Vorpal = require('../../vorpal-init');

class StatusCommand extends Command {
  public get name(): string {
    return commands.STATUS;
  }

  public get description(): string {
    return 'Shows Microsoft 365 login status';
  }

  public commandAction(cmd: CommandInstance, args: {}, cb: (err?: any) => void): void {
    if (auth.service.connected) {
      if (this.debug) {
        cmd.log({
          connectedAs: Utils.getUserNameFromAccessToken(auth.service.accessTokens[auth.defaultResource].value),
          authType: AuthType[auth.service.authType],
          accessTokens: JSON.stringify(auth.service.accessTokens, null, 2),
          refreshToken: auth.service.refreshToken
        });
      }
      else {
        cmd.log({
          connectedAs: Utils.getUserNameFromAccessToken(auth.service.accessTokens[auth.defaultResource].value)
        });
      }
    }
    else {
      if (this.verbose) {
        cmd.log('Logged out from Microsoft 365');
      }
      else {
        cmd.log('Logged out');
      }
    }
    cb();
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
    log(vorpal.find(commands.STATUS).helpInformation());
    log(
      `  Remarks:

    If you are logged in to Microsoft 365, the ${chalk.blue(commands.STATUS)} command
    will show you information about the user or application name used to sign in
    and the details about the stored refresh and access tokens and their
    expiration date and time when run in debug mode.

  Examples:
  
    Show the information about the current login to Microsoft 365
      ${commands.STATUS}
`);
  }
}

module.exports = new StatusCommand();
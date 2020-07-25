import auth from '../../Auth';
import commands from './commands';
import Command, {
  CommandError, CommandAction, CommandArgs
} from '../../Command';
import Utils from '../../Utils';
import { AuthType } from '../../Auth';
import { CommandInstance } from '../../cli';

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
          (cmd as any).initAction(args, this);
          cmd.commandAction(this, args, cb);
        }, (error: any): void => {
          cb(new CommandError(error));
        });
    }
  }
}

module.exports = new StatusCommand();
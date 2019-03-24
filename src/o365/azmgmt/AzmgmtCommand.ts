import Command, { CommandAction, CommandError } from '../../Command';
import auth from './AzmgmtAuth';

export default abstract class AzmgmtCommand extends Command {
  public action(): CommandAction {
    const cmd: AzmgmtCommand = this;

    return function (this: CommandInstance, args: any, cb: (err?: any) => void) {
      auth
        .restoreAuth()
        .then((): void => {
          cmd.initAction(args, this);

          if (!auth.service.connected) {
            cb(new CommandError('Log in to the Azure Management Service first'));
            return;
          }

          cmd.commandAction(this, args, cb);
        }, (error: any): void => {
          cb(new CommandError(error));
        });
    }
  }
}
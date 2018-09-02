import Command, { CommandAction, CommandError } from '../../Command';
import auth from './AadAuth';

export default abstract class AadCommand extends Command {
  public action(): CommandAction {
    const cmd: AadCommand = this;

    return function (this: CommandInstance, args: any, cb: (err?: any) => void) {
      auth
        .restoreAuth()
        .then((): void => {
          cmd.initAction(args);

          if (!auth.service.connected) {
            cb(new CommandError('Connect to Azure Active Directory Graph first'));
            return;
          }

          cmd.commandAction(this, args, cb);
        }, (error: any): void => {
          cb(new CommandError(error));
        });
    }
  }
}
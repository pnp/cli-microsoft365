import Command, { CommandAction, CommandError } from '../../Command';
import auth from './GraphAuth';

export default abstract class GraphCommand extends Command {
  public action(): CommandAction {
    const cmd: GraphCommand = this;

    return function (this: CommandInstance, args: any, cb: (err?: any) => void) {
      auth
        .restoreAuth()
        .then((): void => {
          cmd.initAction(args);

          if (!auth.service.connected) {
            cb(new CommandError('Connect to the Microsoft Graph first'));
            return;
          }

          cmd.commandAction(this, args, cb);
        }, (error: any): void => {
          cb(new CommandError(error));
        });
    }
  }
}
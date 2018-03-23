import Command, { CommandAction, CommandError } from '../../Command';
import appInsights from '../../appInsights';
import auth from './AzmgmtAuth';

export default abstract class AzmgmtCommand extends Command {
  public action(): CommandAction {
    const cmd: AzmgmtCommand = this;

    return function (this: CommandInstance, args: any, cb: () => void) {
      auth
        .restoreAuth()
        .then((): void => {
          cmd._debug = args.options.debug || false;
          cmd._verbose = cmd._debug || args.options.verbose || false;

          appInsights.trackEvent({
            name: cmd.getCommandName(),
            properties: cmd.getTelemetryProperties(args)
          });
          appInsights.flush();

          if (!auth.service.connected) {
            this.log(new CommandError('Connect to the Azure Management Service first'));
            cb();
            return;
          }

          cmd.commandAction(this, args, cb);
        }, (error: any): void => {
          this.log(new CommandError(error));
          cb();
        });
    }
  }
}
import Command, { CommandAction, CommandError } from '../../Command';
import appInsights from '../../appInsights';
import auth from './AadAuth';

export default abstract class AadCommand extends Command {
  public action(): CommandAction {
    const cmd: AadCommand = this;

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
            this.log(new CommandError('Connect to Azure Active Directory Graph site first'));
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
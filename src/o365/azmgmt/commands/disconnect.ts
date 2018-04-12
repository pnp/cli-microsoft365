import auth from '../AzmgmtAuth';
import commands from '../commands';
import config from '../../../config';
import Command, {
  CommandError,
} from '../../../Command';
import appInsights from '../../../appInsights';

const vorpal: Vorpal = require('../../../vorpal-init');

class AzmgmtDisconnectCommand extends Command {
  public get name(): string {
    return commands.DISCONNECT;
  }

  public get description(): string {
    return 'Disconnects from the Azure Management Service';
  }

  public commandAction(cmd: CommandInstance, args: {}, cb: () => void): void {
    const chalk = vorpal.chalk;
    appInsights.trackEvent({
      name: commands.DISCONNECT
    });
    if (this.verbose) {
      cmd.log('Disconnecting from Azure Management Service...');
    }

    const disconnect: () => void = (): void => {
      auth.service.disconnect();
      if (this.verbose) {
        cmd.log(chalk.green('DONE'));
      }
      cb();
    }

    auth
      .clearConnectionInfo()
      .then((): void => {
        disconnect();
      }, (error: any): void => {
        if (this.debug) {
          cmd.log(new CommandError(error));
        }

        disconnect();
      });
  }

  public commandHelp(args: any, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.DISCONNECT).helpInformation());
    log(
      `  Remarks:

    ${chalk.yellow('Attention:')} This command is based on an API that is currently
    in preview and is subject to change once the API reached general
    availability.

    The ${chalk.blue(commands.DISCONNECT)} command disconnects from the Azure
    Management Service and removes any access and refresh tokens from memory.

  Examples:
  
    Disconnect from Azure Management Service
      ${chalk.grey(config.delimiter)} ${commands.DISCONNECT}

    Disconnect from Azure Management Service in debug mode including detailed
    debug information in the console output
      ${chalk.grey(config.delimiter)} ${commands.DISCONNECT} --debug
`);
  }
}

module.exports = new AzmgmtDisconnectCommand();
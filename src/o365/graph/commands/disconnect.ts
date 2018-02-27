import auth from '../GraphAuth';
import commands from '../commands';
import config from '../../../config';
import Command, {
  CommandError,
} from '../../../Command';
import appInsights from '../../../appInsights';

const vorpal: Vorpal = require('../../../vorpal-init');

class GraphDisconnectCommand extends Command {
  public get name(): string {
    return commands.DISCONNECT;
  }

  public get description(): string {
    return 'Disconnects from the Microsoft Graph';
  }

  public commandAction(cmd: CommandInstance, args: {}, cb: () => void): void {
    const chalk = vorpal.chalk;
    appInsights.trackEvent({
      name: commands.DISCONNECT
    });
    if (this.verbose) {
      cmd.log('Disconnecting from Microsoft Graph...');
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

    The ${chalk.blue(commands.DISCONNECT)} command disconnects from the Microsoft Graph
    and removes any access and refresh tokens from memory.

  Examples:
  
    Disconnect from Microsoft Graph
      ${chalk.grey(config.delimiter)} ${commands.DISCONNECT}

    Disconnect from Microsoft Graph in debug mode including detailed debug
    information in the console output
      ${chalk.grey(config.delimiter)} ${commands.DISCONNECT} --debug
`);
  }
}

module.exports = new GraphDisconnectCommand();
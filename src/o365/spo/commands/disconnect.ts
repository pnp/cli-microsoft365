import auth from '../SpoAuth';
import commands from '../commands';
import config from '../../../config';
import Command, {
  CommandHelp, CommandError,
} from '../../../Command';
import appInsights from '../../../appInsights';

const vorpal: Vorpal = require('../../../vorpal-init');

class SpoDisconnectCommand extends Command {
  public get name(): string {
    return commands.DISCONNECT;
  }

  public get description(): string {
    return 'Disconnects from a previously connected SharePoint Online site';
  }

  public commandAction(cmd: CommandInstance, args: {}, cb: () => void): void {
    const chalk = vorpal.chalk;
    appInsights.trackEvent({
      name: commands.DISCONNECT
    });
    if (this.verbose) {
      cmd.log('Disconnecting from SharePoint Online...');
    }

    const disconnect: () => void = (): void => {
      auth.site.disconnect();
      if (this.verbose) {
        cmd.log(chalk.green('DONE'));
      }
      cb();
    }

    auth
      .clearSiteConnectionInfo()
      .then((): void => {
        disconnect();
      }, (error: any): void => {
        if (this.debug) {
          cmd.log(new CommandError(error));
        }

        disconnect();
      });
  }

  public help(): CommandHelp {
    return function (args: any, log: (help: string) => void): void {
      const chalk = vorpal.chalk;
      log(vorpal.find(commands.DISCONNECT).helpInformation());
      log(
        `  Remarks:

    The ${chalk.blue(commands.DISCONNECT)} command disconnects from the previously connected
    SharePoint Online site and removes any access and refresh tokens from memory.

  Examples:
  
    Disconnect from a previously connected SharePoint Online site
      ${chalk.grey(config.delimiter)} ${commands.DISCONNECT}

    Disconnect from a previously connected SharePoint Online site in
    debug mode including detailed debug information in the console output
      ${chalk.grey(config.delimiter)} ${commands.DISCONNECT} --debug
`);
    };
  }
}

module.exports = new SpoDisconnectCommand();
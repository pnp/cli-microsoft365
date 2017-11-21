import auth from '../SpoAuth';
import commands from '../commands';
import config from '../../../config';
import Command, {
  CommandHelp,
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
    cmd.log('Disconnecting from SharePoint Online...');
    auth.site.disconnect();
    cmd.log(chalk.green('DONE'));
    cb();
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
  
    ${chalk.grey(config.delimiter)} ${commands.DISCONNECT}
      disconnects from a previously connected SharePoint Online site

    ${chalk.grey(config.delimiter)} ${commands.DISCONNECT} --verbose
      disconnects from a previously connected SharePoint Online site in
      verbose mode including detailed debug information in the console output
`);
    };
  }
}

module.exports = new SpoDisconnectCommand();
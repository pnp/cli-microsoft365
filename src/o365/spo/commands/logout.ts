import auth from '../SpoAuth';
import commands from '../commands';
import config from '../../../config';
import Command, {
  CommandError,
} from '../../../Command';
import appInsights from '../../../appInsights';

const vorpal: Vorpal = require('../../../vorpal-init');

class SpoLogoutCommand extends Command {
  public get name(): string {
    return commands.LOGOUT;
  }

  public get description(): string {
    return 'Log out from SharePoint Online';
  }

  public alias(): string[] | undefined {
    return [commands.DISCONNECT];
  }

  public commandAction(cmd: CommandInstance, args: {}, cb: () => void): void {
    const chalk = vorpal.chalk;
    
    this.showDeprecationWarning(cmd, commands.DISCONNECT, commands.LOGOUT);

    appInsights.trackEvent({
      name: this.getUsedCommandName(cmd)
    });

    if (this.verbose) {
      cmd.log('Logging out from SharePoint Online...');
    }

    const logOut: () => void = (): void => {
      auth.site.logout();
      if (this.verbose) {
        cmd.log(chalk.green('DONE'));
      }
      cb();
    }

    auth
      .clearSiteConnectionInfo()
      .then((): void => {
        logOut();
      }, (error: any): void => {
        if (this.debug) {
          cmd.log(new CommandError(error));
        }

        logOut();
      });
  }

  public commandHelp(args: any, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.LOGOUT).helpInformation());
    log(
      `  Remarks:

    The ${chalk.blue(commands.LOGOUT)} command logs out from SharePoint Online 
    and removes any access and refresh tokens from memory.

  Examples:
  
    Log out from SharePoint Online
      ${chalk.grey(config.delimiter)} ${commands.LOGOUT}

    Log out from SharePoint Online in debug mode including detailed debug
    information in the console output
      ${chalk.grey(config.delimiter)} ${commands.LOGOUT} --debug
`);
  }
}

module.exports = new SpoLogoutCommand();
import auth from '../SpoAuth';
import config from '../../../config';
import commands from '../commands';
import Command, {
  CommandHelp
} from '../../../Command';

const vorpal: Vorpal = require('../../../vorpal-init');

class SpoStatusCommand extends Command {
  public get name(): string {
    return commands.STATUS;
  }

  public get description(): string {
    return 'Shows SharePoint Online site connection status';
  }
  
  public commandAction(cmd: CommandInstance, args: {}, cb: () => void): void {
    const chalk: any = vorpal.chalk;

    if (auth.site.connected) {
      const expiresAtDate: Date = new Date(0);
      expiresAtDate.setUTCSeconds(auth.service.expiresAt);

      cmd.log(`
Connected to ${auth.site.url}

${chalk.grey('Is tenant admin:')}  ${auth.site.isTenantAdminSite()}
${chalk.grey('AAD resource:')}     ${auth.service.resource}
${chalk.grey('Access token:')}     ${auth.service.accessToken}
${chalk.grey('Refresh token:')}    ${auth.service.refreshToken}
${chalk.grey('Expires at:')}       ${expiresAtDate}
`);
    }
    else {
      cmd.log(`
Not connected to SharePoint Online
`);
    }
    cb();
  }

  public help(): CommandHelp {
    return function (args: any, log: (help: string) => void): void {
      const chalk = vorpal.chalk;
      log(vorpal.find(commands.STATUS).helpInformation());
      log(
        `  Remarks:

    If you are connected to a SharePoint Online, the ${chalk.blue(commands.STATUS)} command
    will show you information about the site to which you are connected, the currently stored
    refresh and access token and the expiration date and time of the access token.

  Examples:
  
    ${chalk.grey(config.delimiter)} ${commands.STATUS}
      shows the information about the current connection to SharePoint Online
`);
    };
  }
}

module.exports = new SpoStatusCommand();
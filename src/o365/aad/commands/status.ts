import auth from '../AadAuth';
import config from '../../../config';
import commands from '../commands';
import Command, {
  CommandError
} from '../../../Command';

const vorpal: Vorpal = require('../../../vorpal-init');

class AadStatusCommand extends Command {
  public get name(): string {
    return commands.STATUS;
  }

  public get description(): string {
    return 'Shows Azure Active Directory Graph connection status';
  }

  public commandAction(cmd: CommandInstance, args: {}, cb: () => void): void {
    const chalk: any = vorpal.chalk;

    auth
      .restoreAuth()
      .then((): void => {
        if (auth.service.connected) {
          const expiresAtDate: Date = new Date(0);
          expiresAtDate.setUTCSeconds(auth.service.expiresAt);

          if (this.verbose) {
            cmd.log(`Connected to ${auth.service.resource}`);
          }
          else {
            cmd.log(auth.service.resource);
          }

          if (this.debug) {
            cmd.log(`
  ${chalk.grey('AAD resource:')}     ${auth.service.resource}
  ${chalk.grey('Access token:')}     ${auth.service.accessToken}
  ${chalk.grey('Refresh token:')}    ${auth.service.refreshToken}
  ${chalk.grey('Expires at:')}       ${expiresAtDate}
  `);
          }
        }
        else {
          if (this.verbose) {
            cmd.log('Not connected to AAD Graph');
          }
          else {
            cmd.log('Not connected');
          }
        }
        cb();
      }, (error: any): void => {
        cmd.log(new CommandError(error));
        cb();
      });
  }

  public commandHelp(args: any, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.STATUS).helpInformation());
    log(
      `  Remarks:

    If you are connected to Azure Active Directory Graph, the ${chalk.blue(commands.STATUS)} command
    will show you information about the currently stored refresh and access token and the
    expiration date and time of the access token when run in debug mode.

  Examples:
  
    Show the information about the current connection to Azure Active Directory Graph
      ${chalk.grey(config.delimiter)} ${commands.STATUS}
`);
  }
}

module.exports = new AadStatusCommand();
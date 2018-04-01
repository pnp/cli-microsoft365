import auth from '../AadAuth';
import config from '../../../config';
import commands from '../commands';
import Command, {
  CommandError
} from '../../../Command';
import Utils from '../../../Utils';
import { AuthType } from '../../../Auth';

const vorpal: Vorpal = require('../../../vorpal-init');

class AadStatusCommand extends Command {
  public get name(): string {
    return commands.STATUS;
  }

  public get description(): string {
    return 'Shows Azure Active Directory Graph connection status';
  }

  public commandAction(cmd: CommandInstance, args: {}, cb: () => void): void {
    auth
      .restoreAuth()
      .then((): void => {
        if (auth.service.connected) {
          if (this.debug) {
            cmd.log({
              connectedAs: Utils.getUserNameFromAccessToken(auth.service.accessToken),
              authType: AuthType[auth.service.authType],
              aadResource: auth.service.resource,
              accessToken: auth.service.accessToken,
              refreshToken: auth.service.refreshToken,
              expiresAt: auth.service.expiresOn
            });
          }
          else {
            cmd.log({
              connectedTo: auth.service.resource,
              connectedAs: Utils.getUserNameFromAccessToken(auth.service.accessToken)
            });
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
    will show you information about the currently stored refresh and access
    token and the expiration date and time of the access token when run in debug
    mode. If you are connected using a user name and password, it will also show
    you the name of the user used to authenticate.

  Examples:
  
    Show the information about the current connection to Azure Active Directory Graph
      ${chalk.grey(config.delimiter)} ${commands.STATUS}
`);
  }
}

module.exports = new AadStatusCommand();
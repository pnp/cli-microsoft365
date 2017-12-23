import auth from '../AadAuth';
import config from '../../../config';
import commands from '../commands';
import GlobalOptions from '../../../GlobalOptions';
import Command, {
  CommandCancel,
  CommandError
} from '../../../Command';
import appInsights from '../../../appInsights';

const vorpal: Vorpal = require('../../../vorpal-init');

interface CommandArgs {
  options: GlobalOptions;
}

class AadConnectCommand extends Command {
  public get name(): string {
    return `${commands.CONNECT}`;
  }

  public get description(): string {
    return 'Connects to the Azure Active Directory Graph';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const chalk: any = vorpal.chalk;

    appInsights.trackEvent({
      name: this.name
    });

    // disconnect before re-connecting
    if (this.debug) {
      cmd.log(`Disconnecting from AAD Graph...`);
    }

    const disconnect: () => void = (): void => {
      auth.service.disconnect();
      if (this.verbose) {
        cmd.log(chalk.green('DONE'));
      }
    }

    const connect: () => void = (): void => {
      if (this.verbose) {
        cmd.log(`Authenticating with AAD Graph...`);
      }

      const resource = 'https://graph.windows.net';

      auth
        .ensureAccessToken(resource, cmd, args.options.debug)
        .then((accessToken: string): Promise<void> => {
          auth.service.resource = resource;
          if (this.verbose) {
            cmd.log(chalk.green('DONE'));
          }

          auth.service.connected = true;
          return auth.storeConnectionInfo();
        })
        .then((): void => {
          cb();
        }, (rej: Error): void => {
          if (this.debug) {
            cmd.log('Error:');
            cmd.log(rej);
            cmd.log('');
          }

          cmd.log(new CommandError(rej.message));
          cb();
        });
    }

    auth
      .clearConnectionInfo()
      .then((): void => {
        disconnect();
        connect();
      }, (error: any): void => {
        if (this.debug) {
          cmd.log(new CommandError(error));
        }

        disconnect();
        connect();
      });
  }

  public cancel(): CommandCancel {
    return (): void => {
      if (auth.interval) {
        clearInterval(auth.interval);
      }
    }
  }

  public commandHelp(args: CommandArgs, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.CONNECT).helpInformation());
    log(
      `  Remarks:
    
    Using the ${chalk.blue(commands.CONNECT)} command you can connect to the Azure Active
    Directory Graph to manage your AAD objects.

    The ${chalk.blue(commands.CONNECT)} command uses device code OAuth flow with the standard
    Microsoft Azure Xplat-CLI Azure AD application to connect to the AAD Graph.
    
    When connecting to the AAD Graph, the ${chalk.blue(commands.CONNECT)} command stores in memory
    the access token and the refresh token. Both tokens are cleared from memory
    after exiting the CLI or by calling the ${chalk.blue(commands.DISCONNECT)} command.

  Examples:
  
    Connect to the AAD Graph
      ${chalk.grey(config.delimiter)} ${commands.CONNECT}

    Connect to the AAD Graph in debug mode including detailed debug information in
    the console output
      ${chalk.grey(config.delimiter)} ${commands.CONNECT} --debug
`);
  }
}

module.exports = new AadConnectCommand();
import auth from '../AzmgmtAuth';
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

class AzmgmtConnectCommand extends Command {
  public get name(): string {
    return `${commands.CONNECT}`;
  }

  public get description(): string {
    return 'Connects to the Azure Management Service';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const chalk: any = vorpal.chalk;

    appInsights.trackEvent({
      name: this.name
    });

    // disconnect before re-connecting
    if (this.debug) {
      cmd.log(`Disconnecting from Azure Management Service...`);
    }

    const disconnect: () => void = (): void => {
      auth.service.disconnect();
      auth.service.resource = 'https://management.azure.com/';
      if (this.verbose) {
        cmd.log(chalk.green('DONE'));
      }
    }

    const connect: () => void = (): void => {
      if (this.verbose) {
        cmd.log(`Authenticating with Azure Management Service...`);
      }

      auth
        .ensureAccessToken('', cmd, args.options.debug)
        .then((accessToken: string): Promise<void> => {
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

    ${chalk.yellow('Attention:')} This command is based on an API that is currently
    in preview and is subject to change once the API reached general
    availability.
    
    Using the ${chalk.blue(commands.CONNECT)} command you can connect to
    the Azure Management Service to manage your Azure objects.

    The ${chalk.blue(commands.CONNECT)} command uses device code OAuth flow
    to connect to the Azure Management Service.
    
    When connecting to the Azure Management Service,
    the ${chalk.blue(commands.CONNECT)} command stores in memory the access token
    and the refresh token. Both tokens are cleared from memory after exiting
    the CLI or by calling the ${chalk.blue(commands.DISCONNECT)} command.

  Examples:
  
    Connect to the Azure Management Service
      ${chalk.grey(config.delimiter)} ${commands.CONNECT}

    Connect to the Azure Management Service in debug mode including detailed
    debug information in the console output
      ${chalk.grey(config.delimiter)} ${commands.CONNECT} --debug
`);
  }
}

module.exports = new AzmgmtConnectCommand();
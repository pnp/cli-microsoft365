import auth from '../GraphAuth';
import config from '../../../config';
import commands from '../commands';
import GlobalOptions from '../../../GlobalOptions';
import Command, {
  CommandCancel,
  CommandOption,
  CommandValidate,
  CommandError
} from '../../../Command';
import appInsights from '../../../appInsights';
import { AuthType } from '../../../Auth';
import * as fs from 'fs';

const vorpal: Vorpal = require('../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  authType?: string;
  userName?: string;
  password?: string;
  applicationId?: string;
  certificateFile?: string;
  thumbprint?: string;
}

class GraphConnectCommand extends Command {
  public get name(): string {
    return `${commands.CONNECT}`;
  }

  public get description(): string {
    return 'Connects to the Microsoft Graph';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const chalk: any = vorpal.chalk;

    appInsights.trackEvent({
      name: this.name
    });

    // disconnect before re-connecting
    if (this.debug) {
      cmd.log(`Disconnecting from Microsoft Graph...`);
    }

    const disconnect: () => void = (): void => {
      auth.service.disconnect();
      auth.service.resource = 'https://graph.microsoft.com';
      if (this.verbose) {
        cmd.log(chalk.green('DONE'));
      }
    }

    const connect: () => void = (): void => {
      if (this.verbose) {
        cmd.log(`Authenticating with Microsoft Graph...`);
      }

      if (args.options.authType === 'password') {
        auth.service.authType = AuthType.Password;
        auth.service.userName = args.options.userName;
        auth.service.password = args.options.password;
      } else if (args.options.authType === 'certificate') {
        auth.service.authType = AuthType.Certificate;
        auth.service.applicationId = args.options.applicationId;
        if (!fs.existsSync(args.options.certificateFile as string)) {
          cb(new CommandError(`File does not exist: ${args.options.certificateFile}`));
          return;
        }
        auth.service.certificate = fs.readFileSync(args.options.certificateFile as string, 'utf8').toString();
        auth.service.thumbprint = args.options.thumbprint;
      }

      auth
        .ensureAccessToken(auth.service.resource, cmd, this.debug)
        .then((accessToken: string): Promise<void> => {
          if (this.verbose) {
            cmd.log(chalk.green('DONE'));
          }

          auth.service.connected = true;
          return auth.storeConnectionInfo();
        })
        .then((): void => {
          cb();
        }, (rej: string): void => {
          if (this.debug) {
            cmd.log('Error:');
            cmd.log(rej);
            cmd.log('');
          }

          if (rej !== 'Polling_Request_Cancelled') {
            cb(new CommandError(rej));
            return;
          }
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
      auth.cancel();
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-t, --authType [authType]',
        description: 'The type of authentication to use. Allowed values deviceCode|password. Default deviceCode',
        autocomplete: ['deviceCode', 'password', 'certificate']
      },
      {
        option: '-u, --userName [userName]',
        description: 'Name of the user to authenticate. Required when authType is set to password'
      },
      {
        option: '-p, --password [password]',
        description: 'Password for the user. Required when authType is set to password'
      },
      {
        option: '-a, --applicationId [AAD application id]',
        description: 'Azure AD application id. Required when authType is set to certificate'
      },
      {
        option: '-c, --certificateFile [certificate file]',
        description: 'File with certificate private key. Required when authType is set to certificate'
      },
      {
        option: '--thumbprint [thumbprint]',
        description: 'Certificate thumbprint. Required when authType is set to certificate'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (args.options.authType === 'password') {
        if (!args.options.userName) {
          return 'Required option userName missing';
        }

        if (!args.options.password) {
          return 'Required option password missing';
        }
      }
      if (args.options.authType === 'certificate') {
        if (!args.options.certificateFile) {
          return 'Required option certificateFile missing';
        }

        if (!args.options.thumbprint) {
          return 'Required option thumbprint missing';
        }

        if (!args.options.applicationId) {
          return 'Required option applicationId missing';
        }
      }

      return true;
    };
  }

  public commandHelp(args: CommandArgs, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.CONNECT).helpInformation());
    log(
      `  Remarks:
    
    Using the ${chalk.blue(commands.CONNECT)} command you can connect to the Microsoft Graph.

    By default, the ${chalk.blue(commands.CONNECT)} command uses device code OAuth flow
    to connect to the Microsoft Graph. Alternatively, you can
    authenticate using a user name and password or a certificate, which is convenient for CI/CD
    scenarios, but which comes with its own limitations. See the Office 365 CLI
    manual for more information.
    
    When connecting to the Microsoft Graph, the ${chalk.blue(commands.CONNECT)} command stores
    in memory the access token and the refresh token. Both tokens are cleared
    from memory after exiting the CLI or by calling the ${chalk.blue(commands.DISCONNECT)}
    command.

    When connecting to the Microsoft Graph using the user name and
    password, next to the access and refresh token, the Office 365 CLI will
    store the user credentials so that it can automatically reauthenticate if
    necessary. Similarly to the tokens, the credentials are removed by
    reconnecting using the device code or by calling the ${chalk.blue(commands.DISCONNECT)}
    command.

    When connecting to the Microsoft Graph using a certificate you are required to create your 
    own Azure AD application and grant permissions accordingly. You are alsoo required to 
    specify the OFFICE365CLI_TENANT environment variable which should have the value of 
    your tenant name; for instance contoso.onmicrosoft.com. 
    Not all commands will work with a certificate as not all features in the 
    Microsoft Graph supports App-only policies.

  Examples:
  
    Connect to the Microsoft Graph using the device code
      ${chalk.grey(config.delimiter)} ${commands.CONNECT}

    Connect to the Microsoft Graph using the device code in debug mode including
    detailed debug information in the console output
      ${chalk.grey(config.delimiter)} ${commands.CONNECT} --debug

    Connect to the Microsoft Graph using a user name and password
      ${chalk.grey(config.delimiter)} ${commands.CONNECT} --authType password --userName user@contoso.com --password pass@word1

    Connect to the Microsoft Graph using a certificate
      ${chalk.grey(config.delimiter)} ${commands.CONNECT} --authType certificate --certificateFile cert.pem 
        --thumbprint d712ebab09e3a9788e9d1a234ea4ac98d173c6c3 --applicationId b269214b-7ed2-4d60-9fb2-064c7b79a4a3
`);
  }
}

module.exports = new GraphConnectCommand();
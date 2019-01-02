import auth from '../AzmgmtAuth';
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
  certificateFile?: string;
  thumbprint?: string;
}

class AzmgmtLoginCommand extends Command {
  public get name(): string {
    return `${commands.LOGIN}`;
  }

  public get description(): string {
    return 'Log in to the Azure Management Service';
  }

  public alias(): string[] | undefined {
    return [commands.CONNECT];
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const chalk: any = vorpal.chalk;

    this.showDeprecationWarning(cmd, commands.CONNECT, commands.LOGIN);

    appInsights.trackEvent({
      name: this.getUsedCommandName(cmd)
    });

    // disconnect before re-connecting
    if (this.debug) {
      cmd.log(`Logging out from Azure Management Service...`);
    }

    const logout: () => void = (): void => {
      auth.service.logout();
      auth.service.resource = 'https://management.azure.com/';
      if (this.verbose) {
        cmd.log(chalk.green('DONE'));
      }
    }

    const login: () => void = (): void => {
      if (this.verbose) {
        cmd.log(`Authenticating with Azure Management Service...`);
      }

      switch (args.options.authType) {
        case 'password':
          auth.service.authType = AuthType.Password;
          auth.service.userName = args.options.userName;
          auth.service.password = args.options.password;
          break;
        case 'certificate':
          auth.service.authType = AuthType.Certificate;
          auth.service.certificate = fs.readFileSync(args.options.certificateFile as string, 'utf8');
          auth.service.thumbprint = args.options.thumbprint;
          break;
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
        logout();
        login();
      }, (error: any): void => {
        if (this.debug) {
          cmd.log(new CommandError(error));
        }

        logout();
        login();
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
        description: 'The type of authentication to use. Allowed values certificate|deviceCode|password. Default deviceCode',
        autocomplete: ['certificate', 'deviceCode', 'password']
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
        option: '-c, --certificateFile [certificateFile]',
        description: 'Path to the file with certificate private key. Required when authType is set to certificate'
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

        if (!fs.existsSync(args.options.certificateFile)) {
          return `File '${args.options.certificateFile}' does not exist`;
        }

        if (!args.options.thumbprint) {
          return 'Required option thumbprint missing';
        }
      }

      return true;
    };
  }

  public commandHelp(args: CommandArgs, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.LOGIN).helpInformation());
    log(
      `  Remarks:

    ${chalk.yellow('Attention:')} This command is based on an API that is currently
    in preview and is subject to change once the API reached general
    availability.
    
    Using the ${chalk.blue(commands.LOGIN)} command you can log in to
    the Azure Management Service to manage your Azure objects.

    By default, the ${chalk.blue(commands.LOGIN)} command uses device code OAuth flow
    to log in to the Azure Management Service. Alternatively, you can
    authenticate using a user name and password or certificate, which are
    convenient for CI/CD scenarios, but which come with their own limitations.
    See the Office 365 CLI manual for more information.
    
    When logging in to the Azure Management Service, the ${chalk.blue(commands.LOGIN)}
    command stores in memory the access token and the refresh token. Both tokens
    are cleared from memory after exiting the CLI or by calling the
    ${chalk.blue(commands.LOGOUT)} command.

    When logging in to the Azure Management Service using the user name and
    password, next to the access and refresh token, the Office 365 CLI will
    store the user credentials so that it can automatically re-authenticate if
    necessary. Similarly to the tokens, the credentials are removed by
    re-authenticating using the device code or by calling the ${chalk.blue(commands.LOGOUT)}
    command.

    When logging in to the Azure Management Service using a certificate,
    the Office 365 CLI will store the contents of the certificate so that it can
    automatically re-authenticate if necessary. The contents of the certificate
    are removed by re-authenticating using the device code or by calling
    the ${chalk.blue(commands.LOGOUT)} command.

    To log in to the Azure Management Service using a certificate,
    you will typically create a custom Azure AD application. To use this
    application with the Office 365 CLI, you will set the ${chalk.grey('OFFICE365CLI_AADAADAPPID')}
    environment variable to the application's ID and the ${chalk.grey('OFFICE365CLI_TENANT')}
    environment variable to the ID of the Azure AD tenant, where you created
    the Azure AD application.

  Examples:
  
    Log in to the Azure Management Service using the device code
      ${chalk.grey(config.delimiter)} ${commands.LOGIN}

    Log in to the Azure Management Service using the device code in debug mode
    including detailed debug information in the console output
      ${chalk.grey(config.delimiter)} ${commands.LOGIN} --debug

    Log in to the Azure Management Service using a user name and password
      ${chalk.grey(config.delimiter)} ${commands.LOGIN} --authType password --userName user@contoso.com --password pass@word1

    Log in to the Azure Management Service using a certificate
      ${chalk.grey(config.delimiter)} ${commands.LOGIN} --authType certificate --certificateFile /Users/user/dev/localhost.pfx --thumbprint 47C4885736C624E90491F32B98855AA8A7562AF1
`);
  }
}

module.exports = new AzmgmtLoginCommand();
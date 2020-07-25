import auth from '../../Auth';
import commands from './commands';
import GlobalOptions from '../../GlobalOptions';
import Command, {
  CommandOption,
  CommandValidate,
  CommandError,
  CommandAction
} from '../../Command';
import { AuthType } from '../../Auth';
import * as fs from 'fs';
import { CommandInstance } from '../../cli';
import * as chalk from 'chalk';

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

class LoginCommand extends Command {
  public get name(): string {
    return `${commands.LOGIN}`;
  }

  public get description(): string {
    return 'Log in to Microsoft 365';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.authType = args.options.authType || 'deviceCode';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    // disconnect before re-connecting
    if (this.debug) {
      cmd.log(`Logging out from Microsoft 365...`);
    }

    const logout: () => void = (): void => {
      auth.service.logout();
      if (this.verbose) {
        cmd.log(chalk.green('DONE'));
      }
    }

    const login: () => void = (): void => {
      if (this.verbose) {
        cmd.log(`Signing in to Microsoft 365...`);
      }

      switch (args.options.authType) {
        case 'password':
          auth.service.authType = AuthType.Password;
          auth.service.userName = args.options.userName;
          auth.service.password = args.options.password;
          break;
        case 'certificate':
          auth.service.authType = AuthType.Certificate;
          auth.service.certificate = fs.readFileSync(args.options.certificateFile as string, 'base64');
          auth.service.thumbprint = args.options.thumbprint;
          auth.service.password = args.options.password;
          break;
        case 'identity':
          auth.service.authType = AuthType.Identity;
          auth.service.userName = args.options.userName;
          break;
      }

      auth
        .ensureAccessToken(auth.defaultResource, cmd, this.debug)
        .then((): void => {
          if (this.verbose) {
            cmd.log(chalk.green('DONE'));
          }

          auth.service.connected = true;
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

  public action(): CommandAction {
    const cmd: Command = this;
    return function (this: CommandInstance, args: CommandArgs, cb: (err?: any) => void) {
      auth
        .restoreAuth()
        .then((): void => {
          (cmd as any).initAction(args, this);
          cmd.commandAction(this, args, cb);
        }, (error: any): void => {
          cb(new CommandError(error));
        });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-t, --authType [authType]',
        description: 'The type of authentication to use. Allowed values certificate|deviceCode|password|identity. Default deviceCode',
        autocomplete: ['certificate', 'deviceCode', 'password', 'identity']
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
}

module.exports = new LoginCommand();
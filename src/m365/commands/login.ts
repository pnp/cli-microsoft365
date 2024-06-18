import fs from 'fs';
import auth, { AuthType, CloudType } from '../../Auth.js';
import { Logger } from '../../cli/Logger.js';
import Command, {
  CommandError
} from '../../Command.js';
import config from '../../config.js';
import GlobalOptions from '../../GlobalOptions.js';
import { misc } from '../../utils/misc.js';
import commands from './commands.js';
import { settingsNames } from '../../settingsNames.js';
import { cli } from '../../cli/cli.js';

interface CommandArgs {
  options: Options;
}

interface AccessToken {
  expiresOn: Date | string | null;
  accessToken: string;
}

interface Options extends GlobalOptions {
  authType?: string;
  cloud?: string;
  userName?: string;
  password?: string;
  certificateFile?: string;
  certificateBase64Encoded?: string;
  thumbprint?: string;
  appId?: string;
  tenant?: string;
  secret?: string;
  connectionName?: string;
  ensure?: boolean;
}

class LoginCommand extends Command {
  private static allowedAuthTypes: string[] = ['certificate', 'deviceCode', 'password', 'identity', 'browser', 'secret'];

  public get name(): string {
    return commands.LOGIN;
  }

  public get description(): string {
    return 'Log in to Microsoft 365';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        authType: args.options.authType || cli.getSettingWithDefaultValue<string>(settingsNames.authType, 'deviceCode'),
        cloud: args.options.cloud ?? CloudType.Public
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-t, --authType [authType]',
        autocomplete: LoginCommand.allowedAuthTypes
      },
      {
        option: '-u, --userName [userName]'
      },
      {
        option: '-p, --password [password]'
      },
      {
        option: '-c, --certificateFile [certificateFile]'
      },
      {
        option: '--certificateBase64Encoded [certificateBase64Encoded]'
      },
      {
        option: '--thumbprint [thumbprint]'
      },
      {
        option: '--appId [appId]'
      },
      {
        option: '--tenant [tenant]'
      },
      {
        option: '-s, --secret [secret]'
      },
      {
        option: '--cloud [cloud]',
        autocomplete: misc.getEnums(CloudType)
      },
      {
        option: '--connectionName [connectionName]'
      },
      {
        option: '--ensure'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const authType = args.options.authType || cli.getSettingWithDefaultValue<string>(settingsNames.authType, 'deviceCode');

        if (authType === 'password') {
          if (!args.options.userName) {
            return 'Required option userName missing';
          }

          if (!args.options.password) {
            return 'Required option password missing';
          }
        }

        if (authType === 'certificate') {
          if (args.options.certificateFile && args.options.certificateBase64Encoded) {
            return 'Specify either certificateFile or certificateBase64Encoded, but not both.';
          }

          if (!args.options.certificateFile && !args.options.certificateBase64Encoded) {
            return 'Specify either certificateFile or certificateBase64Encoded';
          }

          if (args.options.certificateFile) {
            if (!fs.existsSync(args.options.certificateFile)) {
              return `File '${args.options.certificateFile}' does not exist`;
            }
          }
        }

        if (authType &&
          LoginCommand.allowedAuthTypes.indexOf(authType) < 0) {
          return `'${authType}' is not a valid authentication type. Allowed authentication types are ${LoginCommand.allowedAuthTypes.join(', ')}`;
        }

        if (authType === 'secret') {
          if (!args.options.secret) {
            return 'Required option secret missing';
          }
        }

        if (args.options.cloud &&
          typeof CloudType[args.options.cloud as keyof typeof CloudType] === 'undefined') {
          return `${args.options.cloud} is not a valid value for cloud. Valid options are ${misc.getEnums(CloudType).join(', ')}`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    // disconnect before re-connecting
    if (this.debug) {
      await logger.logToStderr(`Logging out from Microsoft 365...`);
    }

    const deactivate: () => void = (): void => auth.connection.deactivate();

    const login: () => Promise<void> = async (): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr(`Signing in to Microsoft 365...`);
      }

      const authType = args.options.authType || cli.getSettingWithDefaultValue<string>(settingsNames.authType, 'deviceCode');
      auth.connection.appId = args.options.appId || config.cliEntraAppId;
      auth.connection.tenant = args.options.tenant || config.tenant;
      auth.connection.name = args.options.connectionName;

      switch (authType) {
        case 'password':
          auth.connection.authType = AuthType.Password;
          auth.connection.userName = args.options.userName;
          auth.connection.password = args.options.password;
          break;
        case 'certificate':
          auth.connection.authType = AuthType.Certificate;
          auth.connection.certificate = args.options.certificateBase64Encoded ? args.options.certificateBase64Encoded : fs.readFileSync(args.options.certificateFile as string, 'base64');
          auth.connection.thumbprint = args.options.thumbprint;
          auth.connection.password = args.options.password;
          break;
        case 'identity':
          auth.connection.authType = AuthType.Identity;
          auth.connection.userName = args.options.userName;
          break;
        case 'browser':
          auth.connection.authType = AuthType.Browser;
          break;
        case 'secret':
          auth.connection.authType = AuthType.Secret;
          auth.connection.secret = args.options.secret;
          break;
      }

      if (args.options.cloud) {
        auth.connection.cloudType = CloudType[args.options.cloud as keyof typeof CloudType];
      }
      else {
        auth.connection.cloudType = CloudType.Public;
      }

      await this.ensureAccessToken(logger);

      const details = auth.getConnectionDetails(auth.connection);

      if (this.debug) {
        (details as any).accessToken = JSON.stringify(auth.connection.accessTokens, null, 2);
      }

      await logger.log(details);
    };

    const shouldReconnect = this.shouldReconnect(args.options);

    if (shouldReconnect) {
      deactivate();
      await login();
    }
    else {
      await this.ensureAccessToken(logger);
    }
  }

  public async action(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      await auth.restoreAuth();
    }
    catch (error: any) {
      throw new CommandError(error);
    }

    this.initAction(args, logger);
    await this.commandAction(logger, args);
  }

  private shouldReconnect(options: Options): boolean {
    const authType = options.authType || cli.getSettingWithDefaultValue<string>(settingsNames.authType, 'deviceCode');

    if (!auth.connection.active) {
      return true;
    }

    if (!options.ensure) {
      return true;
    }

    if (authType.toLocaleLowerCase() !== AuthType[auth.connection.authType].toLocaleLowerCase()) {
      return true;
    }

    if (options.cloud && options.cloud.toLocaleLowerCase() !== CloudType[auth.connection.cloudType].toLocaleLowerCase()) {
      return true;
    }

    if (options.appId && options.appId !== auth.connection.appId) {
      return true;
    }

    if (options.tenant && options.tenant !== auth.connection.tenant) {
      return true;
    }

    if (authType === 'password' && (options.password && options.userName !== auth.connection.userName)) {
      return true;
    }

    if (authType === 'certificate' && (options.certificateFile && (auth.connection.certificate !== fs.readFileSync(options.certificateFile as string, 'base64')))) {
      return true;
    }

    if (authType === 'identity' && (options.userName && options.userName !== auth.connection.userName)) {
      return true;
    }

    if (authType === 'secret' && (options.secret && options.secret !== auth.connection.secret)) {
      return true;
    }

    const now: Date = new Date();
    const accessToken: AccessToken | undefined = auth.connection.accessTokens[auth.defaultResource];

    const expiresOn: Date = accessToken && accessToken.expiresOn ?
      // if expiresOn is serialized from the service file, it's set as a string
      // if it's coming from MSAL, it's a Date
      typeof accessToken.expiresOn === 'string' ? new Date(accessToken.expiresOn) : accessToken.expiresOn
      : new Date(0);

    if (expiresOn < now) {
      return true;
    }

    return false;
  }

  private async ensureAccessToken(logger: Logger): Promise<void> {
    try {
      await auth.ensureAccessToken(auth.defaultResource, logger, this.debug);
      auth.connection.active = true;
    }
    catch (error: any) {
      if (this.debug) {
        await logger.logToStderr('Error:');
        await logger.logToStderr(error);
        await logger.logToStderr('');
      }

      await auth.clearConnectionInfo();
      throw new CommandError(error.message);
    }
  }
}

export default new LoginCommand();
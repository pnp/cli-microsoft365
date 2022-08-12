import * as fs from 'fs';
import auth, { AuthType } from '../../Auth';
import { Logger } from '../../cli';
import Command, {
  CommandError
} from '../../Command';
import config from '../../config';
import GlobalOptions from '../../GlobalOptions';
import commands from './commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  authType?: string;
  userName?: string;
  password?: string;
  certificateFile?: string;
  certificateBase64Encoded?: string;
  thumbprint?: string;
  appId?: string;
  tenant?: string;
  secret?: string;
}

class LoginCommand extends Command {
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
        authType: args.options.authType || 'deviceCode'
      });
    });
  }
  
  #initOptions(): void {
    this.options.unshift(
      {
        option: '-t, --authType [authType]',
        autocomplete: ['certificate', 'deviceCode', 'password', 'identity', 'browser']
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
        option: '--secret [secret]'
      }
    );
  }
  
  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.authType === 'password') {
          if (!args.options.userName) {
            return 'Required option userName missing';
          }
    
          if (!args.options.password) {
            return 'Required option password missing';
          }
        }
    
        if (args.options.authType === 'certificate') {
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
    
        if (args.options.authType === 'secret') {
          if (!args.options.secret) {
            return 'Required option secret missing';
          }
        }
    
        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    // disconnect before re-connecting
    if (this.debug) {
      logger.logToStderr(`Logging out from Microsoft 365...`);
    }

    const logout: () => void = (): void => auth.service.logout();

    const login: () => Promise<void> = async (): Promise<void> => {
      if (this.verbose) {
        logger.logToStderr(`Signing in to Microsoft 365...`);
      }

      auth.service.appId = args.options.appId || config.cliAadAppId;
      auth.service.tenant = args.options.tenant || config.tenant;

      switch (args.options.authType) {
        case 'password':
          auth.service.authType = AuthType.Password;
          auth.service.userName = args.options.userName;
          auth.service.password = args.options.password;
          break;
        case 'certificate':
          auth.service.authType = AuthType.Certificate;
          auth.service.certificate = args.options.certificateBase64Encoded ? args.options.certificateBase64Encoded : fs.readFileSync(args.options.certificateFile as string, 'base64');
          auth.service.thumbprint = args.options.thumbprint;
          auth.service.password = args.options.password;
          break;
        case 'identity':
          auth.service.authType = AuthType.Identity;
          auth.service.userName = args.options.userName;
          break;        
        case 'browser':
          auth.service.authType = AuthType.Browser;
          break;
        case 'secret':
          auth.service.authType = AuthType.Secret;
          auth.service.secret = args.options.secret;
          break;
      }

      try {
        await auth.ensureAccessToken(auth.defaultResource, logger, this.debug);
        auth.service.connected = true;

      }
      catch(error: any) {
        if (this.debug) {
          logger.logToStderr('Error:');
          logger.logToStderr(error);
          logger.logToStderr('');
        }

        throw new CommandError(error.message);
      }
    };

    try {
      await auth.clearConnectionInfo();
    }
    catch (error: any) {
      if (this.debug) {
        logger.logToStderr(new CommandError(error));
      }
    }
    finally {
      logout();
      await login();
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
}

module.exports = new LoginCommand();
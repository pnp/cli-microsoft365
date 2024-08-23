import fs from 'fs';
import { z } from 'zod';
import auth, { AuthType, CloudType } from '../../Auth.js';
import Command, {
  CommandError, globalOptionsZod
} from '../../Command.js';
import { Logger } from '../../cli/Logger.js';
import { cli } from '../../cli/cli.js';
import config from '../../config.js';
import { settingsNames } from '../../settingsNames.js';
import { zod } from '../../utils/zod.js';
import commands from './commands.js';
import { accessToken as accessTokenUtil } from '../../utils/accessToken.js';

const options = globalOptionsZod
  .extend({
    authType: zod.alias('t', z.enum(['certificate', 'deviceCode', 'password', 'identity', 'browser', 'secret', 'accessToken']).optional()),
    cloud: z.nativeEnum(CloudType).optional().default(CloudType.Public),
    userName: zod.alias('u', z.string().optional()),
    password: zod.alias('p', z.string().optional()),
    certificateFile: zod.alias('c', z.string().optional()
      .refine(filePath => !filePath || fs.existsSync(filePath), filePath => ({
        message: `Certificate file ${filePath} does not exist`
      }))),
    certificateBase64Encoded: z.string().optional(),
    thumbprint: z.string().optional(),
    appId: z.string().optional(),
    tenant: z.string().optional(),
    secret: zod.alias('s', z.string().optional()),
    accessToken: zod.alias('a', z.string().or(z.array(z.string())).optional()),
    connectionName: z.string().optional()
  })
  .strict();
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class LoginCommand extends Command {
  public get name(): string {
    return commands.LOGIN;
  }

  public get description(): string {
    return 'Log in to Microsoft 365';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => options.authType !== 'password' || options.userName, {
        message: 'Username is required when using password authentication'
      })
      .refine(options => options.authType !== 'password' || options.password, {
        message: 'Password is required when using password authentication'
      })
      .refine(options => options.authType !== 'certificate' || !(options.certificateFile && options.certificateBase64Encoded), {
        message: 'Specify either certificateFile or certificateBase64Encoded, but not both.'
      })
      .refine(options => options.authType !== 'certificate' || options.certificateFile || options.certificateBase64Encoded, {
        message: 'Specify either certificateFile or certificateBase64Encoded'
      })
      .refine(options => options.authType !== 'secret' || options.secret, {
        message: 'Secret is required when using secret authentication'
      })
      .refine(options => options.authType !== 'accessToken' || options.accessToken, {
        message: 'accessToken is required when using accessToken authentication'
      })
      .refine(options => !(options.authType === 'accessToken' && options.accessToken && !this.validatesAccessTokensAreForSingleTenant(options)), {
        message: 'The provided accessToken is not for the specified tenant or the access tokens are not for the same tenant'
      })
      .refine(options => !(options.authType === 'accessToken' && options.accessToken && !this.validatesAccessTokensAreForSingleApp(options)), {
        message: 'The provided access token is not for the specified app or the access tokens are not for the same app'
      })
      .refine(options => !(options.authType === 'accessToken' && options.accessToken && !this.validatesAccessTokensAreForSingleResource(options)), {
        message: 'Specify access tokens that are not for the same resource'
      })
      .refine(options => !(options.authType === 'accessToken' && options.accessToken && !this.validatesAccessTokensNotExpired(options)), {
        message: 'The provided access token has expired'
      });
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
        case 'accessToken':
          const accessTokens = typeof args.options.accessToken === 'string' ? [args.options.accessToken] : args.options.accessToken as string[];
          auth.connection.authType = AuthType.AccessToken;
          auth.connection.appId = accessTokenUtil.getTenantIdFromAccessToken(accessTokens[0]);
          auth.connection.tenant = accessTokenUtil.getAppIdFromAccessToken(accessTokens[0]);

          for (const token of accessTokens) {
            const resource = accessTokenUtil.getAudienceFromAccessToken(token);
            const expiresOn = accessTokenUtil.getExpirationFromAccessToken(token);

            auth.connection.accessTokens[resource] = {
              expiresOn: expiresOn as Date || null,
              accessToken: token
            };
          };

          break;
      }

      auth.connection.cloudType = args.options.cloud;

      try {
        if (auth.connection.authType !== AuthType.AccessToken) {
          await auth.ensureAccessToken(auth.defaultResource, logger, this.debug);
        }
        else {
          for (const resource of Object.keys(auth.connection.accessTokens)) {
            await auth.ensureAccessToken(resource, logger, this.debug);
          }
        }
        auth.connection.active = true;
      }
      catch (error: any) {
        if (this.debug) {
          await logger.logToStderr('Error:');
          await logger.logToStderr(error);
          await logger.logToStderr('');
        }

        if (error instanceof Error) {
          throw new CommandError(error.message);
        }
        else {
          throw new CommandError(error);
        }
      }

      const details = auth.getConnectionDetails(auth.connection);

      if (this.debug) {
        (details as any).accessToken = JSON.stringify(auth.connection.accessTokens, null, 2);
      }


      await logger.log(details);
    };

    deactivate();
    await login();
  }

  public async action(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      await auth.restoreAuth();
    }
    catch (error: any) {
      throw new CommandError(error);
    }

    await this.initAction(args, logger);
    await this.commandAction(logger, args);
  }

  private validatesAccessTokensAreForSingleTenant(options: Options): boolean {
    const accessTokens = typeof options.accessToken === 'string' ? [options.accessToken] : options.accessToken as string[];
    let tenant = options.tenant || config.tenant;

    for (const token of accessTokens) {
      const tenantIdInAccessToken = accessTokenUtil.getTenantIdFromAccessToken(token);

      if (tenant !== 'common' && tenant !== tenantIdInAccessToken) {
        return false;
      }

      tenant = tenantIdInAccessToken;
    };

    return true;
  }

  private validatesAccessTokensAreForSingleApp(options: Options): boolean {
    const accessTokens = typeof options.accessToken === 'string' ? [options.accessToken] : options.accessToken as string[];
    let appId = options.appId || config.cliEnvEntraAppId || '';

    for (const token of accessTokens) {
      const appIdInAccessToken = accessTokenUtil.getAppIdFromAccessToken(token);

      if (appId !== '' && appId !== appIdInAccessToken) {
        return false;
      }

      appId = appIdInAccessToken;
    };

    return true;
  }

  private validatesAccessTokensAreForSingleResource(options: Options): boolean {
    const accessTokens = typeof options.accessToken === 'string' ? [options.accessToken] : options.accessToken as string[];
    const resources: string[] = [];

    if (accessTokens.length === 1) {
      return true;
    }

    for (const token of accessTokens) {
      const resource = accessTokenUtil.getAudienceFromAccessToken(token);

      if (resources.indexOf(resource) > -1) {
        return false;
      }

      resources.push(resource);
    };

    return true;
  }

  private validatesAccessTokensNotExpired(options: Options): boolean {
    const accessTokens = typeof options.accessToken === 'string' ? [options.accessToken] : options.accessToken as string[];

    for (const token of accessTokens) {
      const expiresOn = accessTokenUtil.getExpirationFromAccessToken(token);

      const accessToken = {
        expiresOn: expiresOn as Date || null,
        accessToken: token
      };

      if (auth.accessTokenExpired(accessToken)) {
        return false;
      }
    };

    return true;
  }
}

export default new LoginCommand();
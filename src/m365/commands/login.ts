import fs from 'fs';
import { z } from 'zod';
import auth, { AccessToken, AuthType, CloudType } from '../../Auth.js';
import Command, {
  CommandError, globalOptionsZod
} from '../../Command.js';
import { Logger } from '../../cli/Logger.js';
import { cli } from '../../cli/cli.js';
import { settingsNames } from '../../settingsNames.js';
import { zod } from '../../utils/zod.js';
import commands from './commands.js';

const options = globalOptionsZod
  .extend({
    authType: zod.alias('t', zod.coercedEnum(AuthType).optional()),
    cloud: zod.coercedEnum(CloudType).optional().default(CloudType.Public),
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
    connectionName: z.string().optional(),
    ensure: z.boolean().optional()
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
      .refine(options => typeof options.appId !== 'undefined' || cli.getConfig().get(settingsNames.clientId), {
        message: `appId is required. TIP: use the "m365 setup" command to configure the default appId`
      })
      .refine(options => options.authType !== 'password' || options.userName, {
        message: 'Username is required when using password authentication',
        path: ['userName']
      })
      .refine(options => options.authType !== 'password' || options.password, {
        message: 'Password is required when using password authentication',
        path: ['password']
      })
      .refine(options => options.authType !== 'certificate' || !(options.certificateFile && options.certificateBase64Encoded), {
        message: 'Specify either certificateFile or certificateBase64Encoded, but not both.',
        path: ['certificateBase64Encoded']
      })
      .refine(options => options.authType !== 'certificate' ||
        options.certificateFile ||
        options.certificateBase64Encoded ||
        cli.getConfig().get(settingsNames.clientCertificateFile) ||
        cli.getConfig().get(settingsNames.clientCertificateBase64Encoded), {
        message: 'Specify either certificateFile or certificateBase64Encoded',
        path: ['certificateFile']
      })
      .refine(options => options.authType !== 'secret' ||
        options.secret ||
        cli.getConfig().get(settingsNames.clientSecret), {
        message: 'Secret is required when using secret authentication',
        path: ['secret']
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.debug) {
      await logger.logToStderr(`Logging out from Microsoft 365...`);
    }

    if (this.shouldLogin(args.options)) {
      auth.connection.deactivate();
      await this.login(logger, args);
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

    await this.initAction(args, logger);
    await this.commandAction(logger, args);
  }

  private shouldLogin(options: Options): boolean {
    if (!auth.connection.active) {
      return true;
    }

    if (!options.ensure) {
      return true;
    }

    const authType = options.authType || cli.getSettingWithDefaultValue<string>(settingsNames.authType, 'deviceCode') as AuthType;
    if (authType !== auth.connection.authType) {
      return true;
    }

    if (options.cloud !== auth.connection.cloudType) {
      return true;
    }

    if (options.appId && options.appId !== auth.connection.appId) {
      return true;
    }

    if (options.tenant && options.tenant !== auth.connection.tenant) {
      return true;
    }

    if (authType === AuthType.Password && (options.password && options.userName !== auth.connection.userName)) {
      return true;
    }

    if (authType === AuthType.Certificate && (options.certificateFile && (auth.connection.certificate !== fs.readFileSync(options.certificateFile as string, 'base64')))) {
      return true;
    }

    if (authType === AuthType.Identity && (options.userName && options.userName !== auth.connection.userName)) {
      return true;
    }

    if (authType === AuthType.Secret && (options.secret && options.secret !== auth.connection.secret)) {
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

      throw new CommandError(error.message);
    }
  }

  private getCertificate(options: Options): string | undefined {
    // command args take precedence over settings
    if (options.certificateFile) {
      return fs.readFileSync(options.certificateFile).toString('base64');
    }
    if (options.certificateBase64Encoded) {
      return options.certificateBase64Encoded;
    }
    return cli.getConfig().get(settingsNames.clientCertificateFile) ||
      cli.getConfig().get(settingsNames.clientCertificateBase64Encoded);
  };

  private async login(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Signing in to Microsoft 365...`);
    }

    const authType = args.options.authType || cli.getSettingWithDefaultValue<string>(settingsNames.authType, 'deviceCode');
    auth.connection.appId = args.options.appId || cli.getClientId();
    auth.connection.tenant = args.options.tenant || cli.getTenant();
    auth.connection.name = args.options.connectionName;

    switch (authType) {
      case 'password':
        auth.connection.authType = AuthType.Password;
        auth.connection.userName = args.options.userName;
        auth.connection.password = args.options.password;
        break;
      case 'certificate':
        auth.connection.authType = AuthType.Certificate;
        auth.connection.certificate = this.getCertificate(args.options);
        auth.connection.thumbprint = args.options.thumbprint;
        auth.connection.password = args.options.password ?? cli.getConfig().get(settingsNames.clientCertificatePassword);
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
        auth.connection.secret = args.options.secret || cli.getConfig().get(settingsNames.clientSecret);
        break;
    }

    auth.connection.cloudType = args.options.cloud;
    await this.ensureAccessToken(logger);

    const details = auth.getConnectionDetails(auth.connection);

    if (this.debug) {
      (details as any).accessToken = JSON.stringify(auth.connection.accessTokens, null, 2);
    }

    await logger.log(details);
  };
}

export default new LoginCommand();
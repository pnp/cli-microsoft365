import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { zod } from '../../../../utils/zod.js';
import config from '../../../../config.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { Logger } from '../../../../cli/Logger.js';
import { cli } from '../../../../cli/cli.js';
import { AppCreationOptions, AppInfo, entraApp } from '../../../../utils/entraApp.js';
import { accessToken } from '../../../../utils/accessToken.js';
import auth from '../../../../Auth.js';

const options = globalOptionsZod
  .extend({
    name: zod.alias('n', z.string().optional().default('CLI for M365')),
    scopes: zod.alias('s', z.string().optional().default('minimal')),
    saveToConfig: z.boolean().optional()
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class CliAppAddCommand extends GraphCommand {
  public get name(): string {
    return commands.APP_ADD;
  }

  public get description(): string {
    return 'Creates a Microsoft Entra application registration for CLI for Microsoft 365';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => {
        const scopes = options.scopes;

        if (!scopes.includes(',')) {
          return scopes === 'minimal' || scopes === 'all';
        }

        const scopeList = scopes.split(',').map(s => s.trim());

        return scopeList.every(scope => scope.startsWith('https'));
      }, {
        message: "Scopes must be 'minimal', 'all', or comma-separated list of URLs starting with 'https'. 'minimal' and 'all' cannot be combined with other scopes.",
        path: ['scopes']
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const options: AppCreationOptions = {
        allowPublicClientFlows: true,
        apisDelegated: this.getScopes(args.options),
        implicitFlow: false,
        multitenant: false,
        name: args.options.name,
        platform: 'publicClient',
        redirectUris: 'http://localhost,https://localhost,https://login.microsoftonline.com/common/oauth2/nativeclient'
      };
      const apis = await entraApp.resolveApis({
        options,
        logger,
        verbose: this.verbose,
        debug: this.debug
      });
      const appInfo: AppInfo = await entraApp.createAppRegistration({
        options,
        unknownOptions: {},
        apis,
        logger,
        verbose: this.verbose,
        debug: this.debug
      });
      appInfo.tenantId = accessToken.getTenantIdFromAccessToken(auth.connection.accessTokens[auth.defaultResource].accessToken);
      await entraApp.grantAdminConsent({
        appInfo,
        appPermissions: entraApp.appPermissions,
        adminConsent: true,
        logger,
        debug: this.debug
      });

      if (args.options.saveToConfig) {
        cli.getConfig().set('clientId', appInfo.appId);
        cli.getConfig().set('tenantId', appInfo.tenantId);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getScopes(options: Options): string {
    if (options.scopes === 'all') {
      return config.allScopes.join(',');
    }
    else if (options.scopes === 'minimal') {
      return config.minimalScopes.join(',');
    }
    return options.scopes;
  }
}

export default new CliAppAddCommand();
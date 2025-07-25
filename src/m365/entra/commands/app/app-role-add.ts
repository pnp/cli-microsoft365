import { v4 } from 'uuid';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { cli } from '../../../../cli/cli.js';
import { entraApp } from '../../../../utils/entraApp.js';
import { formatting } from '../../../../utils/formatting.js';
import { zod } from '../../../../utils/zod.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

const allowedMembers = ['usersGroups', 'applications', 'both'] as const;

const options = globalOptionsZod
  .extend({
    allowedMembers: zod.alias('m', z.enum(allowedMembers)),
    appId: z.string().uuid().optional(),
    appObjectId: z.string().uuid().optional(),
    appName: z.string().optional(),
    claim: zod.alias('c', z.string()),
    name: zod.alias('n', z.string()),
    description: zod.alias('d', z.string())
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

interface AppInfo {
  appRoles: {
    allowedMemberTypes: ('User' | 'Application')[];
    description: string;
    displayName: string;
    id: string;
    value: string;
  }[];
  id: string;
}

class EntraAppRoleAddCommand extends GraphCommand {
  public get name(): string {
    return commands.APP_ROLE_ADD;
  }

  public get description(): string {
    return 'Adds role to the specified Entra app registration';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => [options.appId, options.appObjectId, options.appName].filter(Boolean).length === 1, {
        message: 'Specify either appId, appObjectId, or appName but not multiple'
      })
      .refine(options => options.claim.length <= 120, {
        message: 'Claim must not be longer than 120 characters'
      })
      .refine(options => !options.claim.startsWith('.'), {
        message: 'Claim must not begin with .'
      })
      .refine(options => /^[\w:!#$%&'()*+,-.\/:;<=>?@\[\]^+_`{|}~]+$/.test(options.claim), {
        message: 'Claim can contain only the following characters a-z, A-Z, 0-9, :!#$%&\'()*+,-./:;<=>?@[]^+_`{|}~]+'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const appId = await this.getAppObjectId(args, logger);
      const appInfo = await this.getAppInfo(appId, logger);

      if (this.verbose) {
        await logger.logToStderr(`Adding role ${args.options.name} to Microsoft Entra app ${appInfo.id}...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/myorganization/applications/${appInfo.id}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: {
          appRoles: appInfo.appRoles.concat({
            displayName: args.options.name,
            description: args.options.description,
            id: v4(),
            value: args.options.claim,
            allowedMemberTypes: this.getAllowedMemberTypes(args)
          })
        }
      };

      await request.patch(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getAppInfo(appId: string, logger: Logger): Promise<AppInfo> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving information about roles for Microsoft Entra app ${appId}...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/myorganization/applications/${appId}?$select=id,appRoles`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request.get(requestOptions);
  }

  private getAllowedMemberTypes(args: CommandArgs): ('User' | 'Application')[] {
    switch (args.options.allowedMembers) {
      case 'usersGroups':
        return ['User'];
      case 'applications':
        return ['Application'];
      case 'both':
      default:
        return ['User', 'Application'];
    }
  }

  private async getAppObjectId(args: CommandArgs, logger: Logger): Promise<string> {
    if (args.options.appObjectId) {
      return args.options.appObjectId;
    }

    const { appId, appName } = args.options;

    if (this.verbose) {
      await logger.logToStderr(`Retrieving information about Microsoft Entra app ${appId ? appId : appName}...`);
    }

    if (appId) {
      const app = await entraApp.getAppRegistrationByAppId(appId, ['id']);
      return app.id!;
    }
    else {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/myorganization/applications?$filter=displayName eq '${formatting.encodeQueryParameter(appName as string)}'&$select=id`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const res = await request.get<{ value: { id: string }[] }>(requestOptions);

      if (res.value.length === 1) {
        return res.value[0].id;
      }

      if (res.value.length === 0) {
        throw `No Microsoft Entra application registration with name ${appName} found`;
      }

      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', res.value);
      const result = await cli.handleMultipleResultsFound<{ id: string }>(`Multiple Microsoft Entra application registrations with name '${appName}' found.`, resultAsKeyValuePair);
      return result.id;
    }
  }
}

export default new EntraAppRoleAddCommand();
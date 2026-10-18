import { Application } from '@microsoft/microsoft-graph-types';
import { v4 } from 'uuid';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { entraApp } from '../../../../utils/entraApp.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  appId: z.uuid().optional(),
  appObjectId: z.uuid().optional(),
  appName: z.string().optional(),
  name: z.string().alias('n'),
  description: z.string().alias('d'),
  allowedMembers: z.enum(['usersGroups', 'applications', 'both']).alias('m'),
  claim: z.string()
    .max(120, `Claim must not be longer than 120 characters`)
    .refine(claim => !claim.startsWith('.'), {
      error: 'Claim must not begin with .'
    })
    .refine(claim => /^[\w:!#$%&'()*+,-./:;<=>?@[\]^+_`{|}~]+$/.test(claim), {
      error: `Claim can contain only the following characters a-z, A-Z, 0-9, :!#$%&'()*+,-./:;<=>?@[]^+_\`{|}~]+`
    })
    .alias('c')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraAppRoleAddCommand extends GraphCommand {
  public get name(): string {
    return commands.APP_ROLE_ADD;
  }

  public get description(): string {
    return 'Adds role to the specified Entra app registration';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.appId, options.appObjectId, options.appName].filter(o => o !== undefined).length === 1, {
        error: 'Specify either appId, appObjectId, or appName',
        params: {
          customCode: 'optionSet',
          options: ['appId', 'appObjectId', 'appName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const appInfo = await this.getAppInfo(args, logger);

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
          appRoles: appInfo.appRoles!.concat({
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

  private getAllowedMemberTypes(args: CommandArgs): ('User' | 'Application')[] {
    switch (args.options.allowedMembers) {
      case 'usersGroups':
        return ['User'];
      case 'applications':
        return ['Application'];
      case 'both':
        return ['User', 'Application'];
      default:
        return [];
    }
  }

  private async getAppInfo(args: CommandArgs, logger: Logger): Promise<Application> {
    const { appObjectId, appId, appName } = args.options;

    if (this.verbose) {
      await logger.logToStderr(`Retrieving information about Microsoft Entra app ${appObjectId ? appObjectId : (appId ? appId : appName)}...`);
    }

    if (appObjectId) {
      return await entraApp.getAppRegistrationByObjectId(appObjectId, ['id', 'appRoles']);
    }
    else if (appId) {
      return await entraApp.getAppRegistrationByAppId(appId, ['id', 'appRoles']);
    }
    else {
      return await entraApp.getAppRegistrationByAppName(appName!, ['id', 'appRoles']);
    }
  }
}

export default new EntraAppRoleAddCommand();
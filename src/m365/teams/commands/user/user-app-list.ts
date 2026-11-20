import { TeamsAppDefinition, TeamsAppInstallation } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  userId: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value must be a valid GUID.'
    })
    .optional(),
  userName: z.string()
    .refine(val => validation.isValidUserPrincipalName(val), {
      message: 'The value must be a valid user principal name.'
    })
    .optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsUserAppListCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_APP_LIST;
  }

  public get description(): string {
    return 'Lists the apps installed in the personal scope of the specified user';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema.refine(opts => [opts.userId, opts.userName].filter(o => o !== undefined).length === 1, {
      message: `Specify either 'userId' or 'userName', but not both.`,
      params: {
        customCode: 'optionSet',
        options: ['userId', 'userName']
      }
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const userId: string = (await this.getUserId(args)).value;
      const endpoint: string = `${this.resource}/v1.0/users/${formatting.encodeQueryParameter(userId)}/teamwork/installedApps?$expand=teamsAppDefinition,teamsApp`;

      const items = await odata.getAllItems<TeamsAppInstallation>(endpoint);
      items.forEach(i => {
        const userAppId: string = Buffer.from(i.id as string, 'base64').toString('ascii');
        const appId: string = userAppId.substring(userAppId.indexOf("##") + 2, userAppId.length);
        (i as any).appId = appId;
      });

      if (!cli.shouldTrimOutput(args.options.output)) {
        await logger.log(items);
      }
      else {
        await logger.log(items.map(i => {
          return {
            id: i.id,
            appId: (i as any).appId,
            displayName: (i.teamsAppDefinition as TeamsAppDefinition).displayName,
            version: (i.teamsAppDefinition as TeamsAppDefinition).version
          };
        }));
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getUserId(args: CommandArgs): Promise<{ value: string }> {
    if (args.options.userId) {
      return { value: args.options.userId };
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/users/${formatting.encodeQueryParameter(args.options.userName!)}/id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request.get<{ value: string; }>(requestOptions);
  }
}

export default new TeamsUserAppListCommand();
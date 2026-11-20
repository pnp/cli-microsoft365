import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().optional(),
  name: z.string().optional(),
  userId: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value must be a valid GUID.'
    })
    .optional(),
  userName: z.string()
    .refine(val => validation.isValidUserPrincipalName(val), {
      message: 'The value must be a valid user principal name.'
    })
    .optional(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsUserAppRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_APP_REMOVE;
  }

  public get description(): string {
    return 'Uninstalls an app from the personal scope of the specified user.';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.id, opts.name].filter(o => o !== undefined).length === 1, {
        message: `Specify either 'id' or 'name', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['id', 'name']
        }
      })
      .refine(opts => [opts.userId, opts.userName].filter(o => o !== undefined).length === 1, {
        message: `Specify either 'userId' or 'userName', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['userId', 'userName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeApp = async (): Promise<void> => {
      const appId: string = await this.getAppId(args);
      // validation ensures that here we have either userId or userName
      const userId: string = (args.options.userId ?? args.options.userName) as string;
      const endpoint: string = `${this.resource}/v1.0`;

      if (this.verbose) {
        await logger.logToStderr(`Removing app with ID ${args.options.id} for user ${args.options.userId}`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${endpoint}/users/${formatting.encodeQueryParameter(userId)}/teamwork/installedApps/${appId}`,
        headers: {
          'accept': 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      try {
        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeApp();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the app with id ${args.options.id} for user ${args.options.userId ?? args.options.userName}?` });

      if (result) {
        await removeApp();
      }
    }
  }

  private async getAppId(args: CommandArgs): Promise<string> {
    if (args.options.id) {
      return args.options.id;
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/users/${args.options.userId}/teamwork/installedApps?$expand=teamsAppDefinition&$filter=teamsAppDefinition/displayName eq '${formatting.encodeQueryParameter(args.options.name as string)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: { id: string; }[] }>(requestOptions);

    if (response.value.length === 1) {
      return response.value[0].id;
    }

    if (response.value.length === 0) {
      throw `The specified Teams app does not exist`;
    }

    const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', response.value);
    const result: { id: string } = (await cli.handleMultipleResultsFound(`Multiple Teams apps with name '${args.options.name}' found.`, resultAsKeyValuePair)) as { id: string };
    return result.id;
  }
}

export default new TeamsUserAppRemoveCommand();
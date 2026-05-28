import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import PowerAutomateCommand from '../../../base/PowerAutomateCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  flowName: z.uuid(),
  environmentName: z.string().alias('e'),
  roleName: z.enum(['CanView', 'CanEdit']),
  userId: z.uuid().optional(),
  userName: z.string().refine(name => validation.isValidUserPrincipalName(name), {
    error: e => `'${e.input}' is not a valid userName.`
  }).optional(),
  groupId: z.uuid().optional(),
  groupName: z.string().optional(),
  asAdmin: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class FlowOwnerEnsureCommand extends PowerAutomateCommand {
  public get name(): string {
    return commands.OWNER_ENSURE;
  }

  public get description(): string {
    return 'Assigns/updates permissions to a Power Automate flow';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema.refine(
      options => [options.userId, options.userName, options.groupId, options.groupName].filter(x => x !== undefined).length === 1,
      {
        error: 'Specify either userId, userName, groupId, or groupName, but not multiple.',
        params: {
          customCode: 'optionSet',
          options: ['userId', 'userName', 'groupId', 'groupName']
        }
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Assigning permissions for ${args.options.userId || args.options.userName || args.options.groupId || args.options.groupName} with permissions ${args.options.roleName} to Power Automate flow ${args.options.flowName}`);
      }

      let id = '';
      if (args.options.userId) {
        id = args.options.userId;
      }
      else if (args.options.userName) {
        id = await entraUser.getUserIdByUpn(args.options.userName);
      }
      else if (args.options.groupId) {
        id = args.options.groupId;
      }
      else {
        id = await entraGroup.getGroupIdByDisplayName(args.options.groupName!);
      }

      let type: string;
      if (args.options.userId || args.options.userName) {
        type = 'User';
      }
      else {
        type = 'Group';
      }

      const requestOptions: CliRequestOptions = {
        url: `${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple/${args.options.asAdmin ? 'scopes/admin/' : ''}environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.flowName)}/modifyPermissions?api-version=2016-11-01`,
        headers: {
          accept: 'application/json'
        },
        data: {
          put: [
            {
              properties: {
                principal: {
                  id: id,
                  type: type
                },
                roleName: args.options.roleName
              }
            }
          ]
        },
        responseType: 'json'
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new FlowOwnerEnsureCommand();
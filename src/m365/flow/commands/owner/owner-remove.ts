import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
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
  userId: z.uuid().optional(),
  userName: z.string().refine(name => validation.isValidUserPrincipalName(name), {
    error: e => `'${e.input}' is not a valid userName.`
  }).optional(),
  groupId: z.uuid().optional(),
  groupName: z.string().optional(),
  asAdmin: z.boolean().optional(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class FlowOwnerRemoveCommand extends PowerAutomateCommand {
  public get name(): string {
    return commands.OWNER_REMOVE;
  }

  public get description(): string {
    return 'Removes owner permissions to a Power Automate flow';
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
        await logger.logToStderr(`Removing owner ${args.options.userId || args.options.userName || args.options.groupId || args.options.groupName} from flow ${args.options.flowName} in environment ${args.options.environmentName}`);
      }

      const removeFlowOwner = async (): Promise<void> => {
        let idToRemove: string;
        if (args.options.userId) {
          idToRemove = args.options.userId;
        }
        else if (args.options.userName) {
          idToRemove = await entraUser.getUserIdByUpn(args.options.userName);
        }
        else if (args.options.groupId) {
          idToRemove = args.options.groupId;
        }
        else {
          idToRemove = await entraGroup.getGroupIdByDisplayName(args.options.groupName!);
        }

        const requestOptions: CliRequestOptions = {
          url: `${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple/${args.options.asAdmin ? 'scopes/admin/' : ''}environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.flowName)}/modifyPermissions?api-version=2016-11-01`,
          headers: {
            accept: 'application/json'
          },
          data: {
            delete: [
              {
                id: idToRemove
              }
            ]
          },
          responseType: 'json'
        };
        await request.post(requestOptions);
      };

      if (args.options.force) {
        await removeFlowOwner();
      }
      else {
        const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove owner '${args.options.groupId || args.options.groupName || args.options.userId || args.options.userName}' from the specified flow?` });

        if (result) {
          await removeFlowOwner();
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new FlowOwnerRemoveCommand();
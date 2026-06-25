import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  rosterId: z.string(),
  userId: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value is not a valid GUID.'
    })
    .optional(),
  userName: z.string()
    .refine(val => validation.isValidUserPrincipalName(val), {
      message: 'The value is not a valid user principal name (UPN).'
    })
    .optional(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PlannerRosterMemberRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.ROSTER_MEMBER_REMOVE;
  }

  public get description(): string {
    return 'Removes a member from a Microsoft Planner Roster';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodType | undefined {
    return schema
      .refine(opts => [opts.userId, opts.userName].filter(x => x !== undefined).length === 1, {
        message: `Specify exactly one of the following options: 'userId' or 'userName'.`,
        params: {
          customCode: 'optionSet',
          options: ['userId', 'userName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Removing member ${args.options.userName || args.options.userId} from the Microsoft Planner Roster`);
    }

    if (args.options.force) {
      await this.removeRosterMember(args);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove member '${args.options.userId || args.options.userName}'?` });

      if (result) {
        await this.removeRosterMember(args);
      }
    }
  }

  private async getUserId(args: CommandArgs): Promise<string> {
    if (args.options.userId) {
      return args.options.userId;
    }

    return entraUser.getUserIdByUpn(args.options.userName!);
  }

  private async removeRosterMember(args: CommandArgs): Promise<void> {
    try {
      const rosterMembersContinue = await this.removeLastMemberConfirmation(args);
      if (rosterMembersContinue) {
        const userId = await this.getUserId(args);

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/beta/planner/rosters/${args.options.rosterId}/members/${userId}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        await request.delete(requestOptions);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async removeLastMemberConfirmation(args: CommandArgs): Promise<boolean> {
    if (!args.options.force) {
      const rosterMembers = await odata.getAllItems(`${this.resource}/beta/planner/rosters/${args.options.rosterId}/members?$select=Id`);
      if (rosterMembers.length === 1) {
        const result = await cli.promptForConfirmation({ message: 'You are about to remove the last member of this Roster. When this happens, the Roster and all its contents will be deleted within 30 days. Are you sure you want to proceed?' });

        return result;
      }
    }

    return true;
  }
}

export default new PlannerRosterMemberRemoveCommand();

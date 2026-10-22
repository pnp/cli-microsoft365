import auth from '../../../../Auth.js';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { odata } from '../../../../utils/odata.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  userId: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value is not a valid GUID.'
    })
    .optional(),
  userName: z.string()
    .refine(val => validation.isValidUserPrincipalName(val), {
      message: 'The value is not a valid user principal name (UPN).'
    })
    .optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PlannerRosterPlanListCommand extends GraphCommand {
  public get name(): string {
    return commands.ROSTER_PLAN_LIST;
  }

  public get description(): string {
    return 'Lists all Microsoft Planner Roster plans for a specific user';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodType | undefined {
    return schema
      .refine(opts => !opts.userId || !opts.userName, {
        message: `Specify either 'userId' or 'userName', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['userId', 'userName']
        }
      });
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'title', 'createdDateTime', 'owner'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[this.resource].accessToken);
    if (isAppOnlyAccessToken && !args.options.userId && !args.options.userName) {
      this.handleError(`Specify at least 'userId' or 'userName' when using application permissions.`);
    }
    else if (!isAppOnlyAccessToken && (args.options.userId || args.options.userName)) {
      this.handleError(`The options 'userId' or 'userName' cannot be used when obtaining Microsoft Planner Roster plans using delegated permissions.`);
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving Microsoft Planner Roster plans for user: ${args.options.userId || args.options.userName || 'current user'}.`);
    }

    let requestUrl: string = `${this.resource}/beta/`;
    if (args.options.userId || args.options.userName) {
      requestUrl += `users/${args.options.userId || args.options.userName}`;
    }
    else {
      requestUrl += 'me';
    }
    requestUrl += '/planner/rosterPlans';

    try {
      const items = await odata.getAllItems(requestUrl);
      await logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PlannerRosterPlanListCommand();

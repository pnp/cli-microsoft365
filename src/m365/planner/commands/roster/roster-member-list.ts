import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  rosterId: z.string()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PlannerRosterMemberListCommand extends GraphCommand {
  public get name(): string {
    return commands.ROSTER_MEMBER_LIST;
  }

  public get description(): string {
    return 'Lists members of the specified Microsoft Planner Roster';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Retrieving members of the specified Microsoft Planner Roster');
    }

    try {
      const response = await odata.getAllItems(`${this.resource}/beta/planner/rosters/${args.options.rosterId}/members`);
      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PlannerRosterMemberListCommand();

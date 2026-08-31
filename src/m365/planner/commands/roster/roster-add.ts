import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = globalOptionsZod.strict();

class PlannerRosterAddCommand extends GraphCommand {
  public get name(): string {
    return commands.ROSTER_ADD;
  }

  public get description(): string {
    return 'Creates a new Microsoft Planner Roster';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Creating a new Microsoft Planner Roster');
    }

    try {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/beta/planner/rosters`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        data: {},
        responseType: 'json'
      };

      const response = await request.post(requestOptions);
      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PlannerRosterAddCommand();

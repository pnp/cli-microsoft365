import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PlannerRosterRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.ROSTER_REMOVE;
  }

  public get description(): string {
    return 'Removes a Microsoft Planner Roster';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removeRoster(args, logger);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove roster ${args.options.id}?` });

      if (result) {
        await this.removeRoster(args, logger);
      }
    }
  }

  private async removeRoster(args: CommandArgs, logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Removing roster ${args.options.id}`);
    }
    try {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/beta/planner/rosters/${args.options.id}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      await request.delete(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PlannerRosterRemoveCommand();

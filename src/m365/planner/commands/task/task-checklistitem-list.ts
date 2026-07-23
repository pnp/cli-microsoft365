import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  taskId: z.string().alias('i')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PlannerTaskChecklistItemListCommand extends GraphCommand {
  public get name(): string {
    return commands.TASK_CHECKLISTITEM_LIST;
  }

  public get description(): string {
    return 'Lists the checklist items of a Planner task.';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'title', 'isChecked'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/planner/tasks/${formatting.encodeQueryParameter(args.options.taskId)}/details?$select=checklist`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      const res = await request.get<any>(requestOptions);
      if (!args.options.output || !cli.shouldTrimOutput(args.options.output)) {
        await logger.log(res.checklist);
      }
      else {
        //converted to text friendly output
        const output = Object.getOwnPropertyNames(res.checklist).map(prop => ({ id: prop, ...(res.checklist as any)[prop] }));
        await logger.log(output);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PlannerTaskChecklistItemListCommand();

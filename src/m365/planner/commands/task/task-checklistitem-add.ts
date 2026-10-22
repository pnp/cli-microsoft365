import { PlannerTaskDetails } from '@microsoft/microsoft-graph-types';
import { v4 } from 'uuid';
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
  taskId: z.string().alias('i'),
  title: z.string().alias('t'),
  isChecked: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PlannerTaskChecklistItemAddCommand extends GraphCommand {
  public get name(): string {
    return commands.TASK_CHECKLISTITEM_ADD;
  }

  public get description(): string {
    return 'Adds a new checklist item to a Planner task.';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const etag = await this.getTaskDetailsEtag(args.options.taskId);
      const body: PlannerTaskDetails = {
        checklist: {
          // Generate new GUID for new task checklist item
          [v4()]: {
            '@odata.type': 'microsoft.graph.plannerChecklistItem',
            title: args.options.title,
            isChecked: args.options.isChecked || false
          }
        }
      };

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/planner/tasks/${formatting.encodeQueryParameter(args.options.taskId)}/details`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          prefer: 'return=representation',
          'if-match': etag
        },
        responseType: 'json',
        data: body
      };

      const result = await request.patch<PlannerTaskDetails>(requestOptions);
      if (!cli.shouldTrimOutput(args.options.output)) {
        await logger.log(result.checklist);
      }
      else {
        // Transform checklist item object to text friendly format
        const output = Object.getOwnPropertyNames(result.checklist).map(prop => ({ id: prop, ...(result.checklist as any)[prop] }));
        await logger.log(output);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getTaskDetailsEtag(taskId: string): Promise<string> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/planner/tasks/${formatting.encodeQueryParameter(taskId)}/details`,
      headers: {
        accept: 'application/json;odata.metadata=minimal'
      },
      responseType: 'json'
    };

    const task = await request.get<any>(requestOptions);
    return task['@odata.etag'];
  }
}

export default new PlannerTaskChecklistItemAddCommand();

import { PlannerTaskDetails } from '@microsoft/microsoft-graph-types';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  taskId: string;
  force?: boolean;
}

class PlannerTaskChecklistItemRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.TASK_CHECKLISTITEM_REMOVE;
  }

  public get description(): string {
    return 'Removes the checklist item from the planner task';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        force: (!(!args.options.force)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '-i, --id <id>' },
      { option: '--taskId <taskId>' },
      { option: '-f, --force' }
    );
  }

  #initTypes(): void {
    this.types.string.push('id', 'taskId');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removeChecklistitem(args);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the checklist item with id ${args.options.id} from the planner task?` });

      if (result) {
        await this.removeChecklistitem(args);
      }
    }
  }

  private async removeChecklistitem(args: CommandArgs): Promise<void> {
    try {
      const task = await this.getTaskDetails(args.options.taskId);
      if (!task.checklist || !(task.checklist as any)[args.options.id]) {
        throw `The specified checklist item with id ${args.options.id} does not exist`;
      }

      const requestOptionsTaskDetails: CliRequestOptions = {
        url: `${this.resource}/v1.0/planner/tasks/${args.options.taskId}/details`,
        headers: {
          'accept': 'application/json;odata.metadata=none',
          'If-Match': (task as any)['@odata.etag'],
          'Prefer': 'return=representation'
        },
        responseType: 'json',
        data: {
          checklist: {
            [args.options.id]: null
          }
        }
      };

      await request.patch(requestOptionsTaskDetails);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getTaskDetails(taskId: string): Promise<PlannerTaskDetails> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/planner/tasks/${formatting.encodeQueryParameter(taskId)}/details?$select=checklist`,
      headers: {
        accept: 'application/json;odata.metadata=minimal'
      },
      responseType: 'json'
    };

    return await request.get(requestOptions);
  }
}

export default new PlannerTaskChecklistItemRemoveCommand();
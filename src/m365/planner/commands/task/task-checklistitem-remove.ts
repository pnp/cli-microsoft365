import { PlannerTaskDetails } from '@microsoft/microsoft-graph-types';
import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  taskId: string;
  confirm?: boolean;
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
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        confirm: (!(!args.options.confirm)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '-i, --id <id>' },
      { option: '--taskId <taskId>' },
      { option: '--confirm' }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.confirm) {
      await this.removeChecklistitem(args);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the checklist item with id ${args.options.id} from the planner task?`
      });

      if (result.continue) {
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

module.exports = new PlannerTaskChecklistItemRemoveCommand();
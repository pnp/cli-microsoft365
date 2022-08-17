import { PlannerTaskDetails } from '@microsoft/microsoft-graph-types';
import { Cli, Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (args.options.confirm) {
      this.removeChecklistitem(logger, args, cb);
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the checklist item with id ${args.options.id} from the planner task?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          this.removeChecklistitem(logger, args, cb);
        }
      });
    }
  }

  private removeChecklistitem(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    this
      .getTaskDetails(args.options.taskId)
      .then(task => {
        if (!task.checklist || !(task.checklist as any)[args.options.id]) {
          return Promise.reject(`The specified checklist item with id ${args.options.id} does not exist`);
        }

        const requestOptionsTaskDetails: any = {
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

        return request.patch(requestOptionsTaskDetails);
      })
      .then(_ => cb(), (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getTaskDetails(taskId: string): Promise<PlannerTaskDetails> {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/planner/tasks/${encodeURIComponent(taskId)}/details?$select=checklist`,
      headers: {
        accept: 'application/json;odata.metadata=minimal'
      },
      responseType: 'json'
    };

    return request.get(requestOptions);
  }
}

module.exports = new PlannerTaskChecklistItemRemoveCommand();
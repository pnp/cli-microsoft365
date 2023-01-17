import { PlannerTaskDetails } from '@microsoft/microsoft-graph-types';
import { v4 } from 'uuid';
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
  taskId: string;
  title: string;
  isChecked?: boolean;
}

class PlannerTaskChecklistItemAddCommand extends GraphCommand {
  public get name(): string {
    return commands.TASK_CHECKLISTITEM_ADD;
  }

  public get description(): string {
    return 'Adds a new checklist item to a Planner task.';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'title', 'isChecked'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        isChecked: args.options.isChecked || false
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '-i, --taskId <taskId>' },
      { option: '-t, --title <title>' },
      { option: '--isChecked' }
    );
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
      if (args.options.output === 'json') {
        logger.log(result.checklist);
      }
      else {
        // Transform checklist item object to text friendly format
        const output = Object.getOwnPropertyNames(result.checklist).map(prop => ({ id: prop, ...(result.checklist as any)[prop] }));
        logger.log(output);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getTaskDetailsEtag(taskId: string): Promise<string> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/planner/tasks/${formatting.encodeQueryParameter(taskId)}/details`,
      headers: {
        accept: 'application/json;odata.metadata=minimal'
      },
      responseType: 'json'
    };

    return request
      .get(requestOptions)
      .then((task: any) => task['@odata.etag'],
        () => Promise.reject('Planner task was not found.'));
  }
}

module.exports = new PlannerTaskChecklistItemAddCommand();
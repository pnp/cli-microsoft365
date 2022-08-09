import { PlannerTaskDetails } from '@microsoft/microsoft-graph-types';
import { AxiosRequestConfig } from 'axios';
import { v4 } from 'uuid';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { accessToken } from '../../../../utils/accessToken';
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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (accessToken.isAppOnlyAccessToken(auth.service.accessTokens[this.resource].accessToken)) {
      this.handleError('This command does not support application permissions.', logger, cb);
      return;
    }

    this
      .getTaskDetailsEtag(args.options.taskId)
      .then(etag => {
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

        const requestOptions: AxiosRequestConfig = {
          url: `${this.resource}/v1.0/planner/tasks/${encodeURIComponent(args.options.taskId)}/details`,
          headers: {
            accept: 'application/json;odata.metadata=none',
            prefer: 'return=representation',
            'if-match': etag
          },
          responseType: 'json',
          data: body
        };

        return request.patch<PlannerTaskDetails>(requestOptions);
      })
      .then((result): void => {
        if (args.options.output === 'json') {
          logger.log(result.checklist);
        }
        else {
          // Transform checklist item object to text friendly format
          const output = Object.getOwnPropertyNames(result.checklist).map(prop => ({ id: prop, ...(result.checklist as any)[prop] }));
          logger.log(output);
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getTaskDetailsEtag(taskId: string): Promise<string> {
    const requestOptions: AxiosRequestConfig = {
      url: `${this.resource}/v1.0/planner/tasks/${encodeURIComponent(taskId)}/details`,
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
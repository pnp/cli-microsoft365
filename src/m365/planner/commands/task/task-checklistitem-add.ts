import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import { accessToken } from '../../../../utils/accessToken';
import { PlannerTaskDetails } from '@microsoft/microsoft-graph-types';
import { AxiosRequestConfig } from 'axios';
import { v4 } from 'uuid';
import GlobalOptions from '../../../../GlobalOptions';
import Auth from '../../../../Auth';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import request from '../../../../request';

interface ExtendedPlannerTaskDetails extends PlannerTaskDetails {
  '@odata.etag': string;
}

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  taskId: string;
  title: string;
  isChecked?: boolean;
}

class PlannerChecklistItemAddCommand extends GraphCommand {
  public get name(): string {
    return commands.TASK_CHECKLISTITEM_ADD;
  }

  public get description(): string {
    return 'Adds a new checklist item to a Planner task.';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.isChecked = args.options.isChecked || false;
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (accessToken.isAppOnlyAccessToken(Auth.service.accessTokens[this.resource].accessToken)) {
      this.handleError('This command does not support application permissions.', logger, cb);
      return;
    }

    this
      .getTaskDetails(args.options.taskId)
      .then(details => {
        const etag: string = details['@odata.etag'];

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
      .then((res): void => {
        logger.log(res.checklist);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getTaskDetails(taskId: string): Promise<ExtendedPlannerTaskDetails> {
    const requestOptions: AxiosRequestConfig = {
      url: `${this.resource}/v1.0/planner/tasks/${encodeURIComponent(taskId)}/details`,
      headers: {
        accept: 'application/json;odata.metadata=minimal'
      },
      responseType: 'json'
    };

    return request.get(requestOptions);
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      { option: '-i, --taskId <taskId>' },
      { option: '-t, --title <title>' },
      { option: '--isChecked' }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new PlannerChecklistItemAddCommand();
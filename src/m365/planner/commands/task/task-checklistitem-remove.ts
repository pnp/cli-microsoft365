import { Cli, Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
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

class PlannerTaskChecklistitemRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.TASK_CHECKLISTITEM_REMOVE;
  }

  public get description(): string {
    return 'Removes the Checklistitem from the Planner task';
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
        message: `Are you sure you want to remove the Checklist item from the Planner task?`
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
      .getTaskDetailsEtag(args.options.taskId)
      .then(etag => {
        const requestOptionsTaskDetails: any = {
          url: `${this.resource}/v1.0/planner/tasks/${args.options.taskId}/details`,
          headers: {
            'accept': 'application/json;odata.metadata=none',
            'If-Match': etag,
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
      .then((res: any): void => {
        logger.log(res.checklist);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getTaskDetailsEtag(taskId: string): Promise<string> {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/planner/tasks/${encodeURIComponent(taskId)}/details`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    return request
      .get(requestOptions)
      .then((response: any) => {
        const etag: string | undefined = response ? response['@odata.etag'] : undefined;
        if (!etag) {
          return Promise.reject(`Error fetching task details`);
        }
        return Promise.resolve(etag);
      });
  }
  public options(): CommandOption[] {
    const options: CommandOption[] = [
      { option: '-i, --id <id>' },
      { option: '--taskId <taskId>' },
      { option: '--confirm' }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

}

module.exports = new PlannerTaskChecklistitemRemoveCommand();
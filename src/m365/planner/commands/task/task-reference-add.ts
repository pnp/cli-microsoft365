import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  taskId: string;
  url: string;
  alias?: string;
  type?: string;
}

class PlannerTaskReferenceAddCommand extends GraphCommand {  
  public get name(): string {
    return commands.TASK_REFERENCE_ADD;
  }

  public get description(): string {
    return 'Adds a new reference to a Planner task';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getTaskDetailsEtag(args.options.taskId)
      .then(etag => {
        const requestOptionsTaskDetails: any = {
          url: `${this.resource}/v1.0/planner/tasks/${encodeURIComponent(args.options.taskId)}/details`,
          headers: {
            'accept': 'application/json;odata.metadata=none',
            'If-Match': etag,
            'Prefer': 'return=representation'
          },
          responseType: 'json',
          data: {
            references: {
              [formatting.openTypesEncoder(args.options.url)]: {
                '@odata.type': 'microsoft.graph.plannerExternalReference',
                previewPriority: ' !',
                ...(args.options.alias && {alias: args.options.alias}),
                ...(args.options.type && {type: args.options.type})
              }
            }
          }
        };

        return request.patch(requestOptionsTaskDetails);
      })
      .then((res: any): void => {
        logger.log(res.references);
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
      { option: '-i, --taskId <taskId>' },
      { option: '-u, --url <url>' },
      { option: '--alias [alias]' },
      { option: '--type [type]' }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
  
  public validate(args: CommandArgs): boolean | string {
    if (args.options.type && ['powerpoint', 'word', 'excel', 'other'].indexOf(args.options.type.toLocaleLowerCase()) === -1) {
      return `${args.options.type} is not a valid type value. Allowed values PowerPoint|Word|Excel|Other`;
    } 

    return true;
  }
}

module.exports = new PlannerTaskReferenceAddCommand();

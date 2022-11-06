import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
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

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        alias: typeof args.options.alias !== 'undefined',
        type: args.options.type
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '-i, --taskId <taskId>' },
      { option: '-u, --url <url>' },
      { option: '--alias [alias]' },
      {
        option: '--type [type]',
        autocomplete: ['PowerPoint', 'Word', 'Excel', 'Other']
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.type && ['powerpoint', 'word', 'excel', 'other'].indexOf(args.options.type.toLocaleLowerCase()) === -1) {
          return `${args.options.type} is not a valid type value. Allowed values PowerPoint|Word|Excel|Other`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const etag = await this.getTaskDetailsEtag(args.options.taskId);
      const requestOptionsTaskDetails: any = {
        url: `${this.resource}/v1.0/planner/tasks/${formatting.encodeQueryParameter(args.options.taskId)}/details`,
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
              ...(args.options.alias && { alias: args.options.alias }),
              ...(args.options.type && { type: args.options.type })
            }
          }
        }
      };
      const res = await request.patch<any>(requestOptionsTaskDetails);
      logger.log(res.references);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getTaskDetailsEtag(taskId: string): Promise<string> {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/planner/tasks/${formatting.encodeQueryParameter(taskId)}/details`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    return request
      .get(requestOptions)
      .then((response: any) => response['@odata.etag']);
  }
}

module.exports = new PlannerTaskReferenceAddCommand();

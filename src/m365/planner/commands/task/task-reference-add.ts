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
    this.#initTypes();
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

  #initTypes(): void {
    this.types.string.push('taskId', 'url', 'alias', 'type');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const etag = await this.getTaskDetailsEtag(args.options.taskId);
      const requestOptionsTaskDetails: CliRequestOptions = {
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
      await logger.log(res.references);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getTaskDetailsEtag(taskId: string): Promise<string> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/planner/tasks/${formatting.encodeQueryParameter(taskId)}/details`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    const response = await request.get<any>(requestOptions);
    return response['@odata.etag'];
  }
}

export default new PlannerTaskReferenceAddCommand();

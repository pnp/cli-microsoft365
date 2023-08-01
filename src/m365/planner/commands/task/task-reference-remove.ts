import { PlannerTaskDetails } from '@microsoft/microsoft-graph-types';
import { Cli } from '../../../../cli/Cli.js';
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
  url?: string;
  alias?: string;
  taskId: string;
  force?: boolean;
}

class PlannerTaskReferenceRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.TASK_REFERENCE_REMOVE;
  }

  public get description(): string {
    return 'Removes the reference from the Planner task';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        url: typeof args.options.url !== 'undefined',
        alias: typeof args.options.alias !== 'undefined',
        force: (!(!args.options.force)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '-u, --url [url]' },
      { option: '--alias [alias]' },
      { option: '-i, --taskId <taskId>' },
      { option: '-f, --force' }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.url && args.options.url.indexOf('https://') !== 0 && args.options.url.indexOf('http://') !== 0) {
          return 'The url option should contain a valid URL. A valid URL starts with http(s)://';
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['url', 'alias'] }
    );
  }

  #initTypes(): void {
    this.types.string.push('taskId', 'alias');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removeReference(logger, args);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the reference from the Planner task?`
      });

      if (result.continue) {
        await this.removeReference(logger, args);
      }
    }
  }

  private async removeReference(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const { etag, url } = await this.getTaskDetailsEtagAndUrl(args.options);
      const requestOptionsTaskDetails: CliRequestOptions = {
        url: `${this.resource}/v1.0/planner/tasks/${args.options.taskId}/details`,
        headers: {
          'accept': 'application/json;odata.metadata=none',
          'If-Match': etag,
          'Prefer': 'return=representation'
        },
        responseType: 'json',
        data: {
          references: {
            [formatting.openTypesEncoder(url)]: null
          }
        }
      };

      await request.patch(requestOptionsTaskDetails);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getTaskDetailsEtagAndUrl(options: Options): Promise<{ etag: string, url: string }> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/planner/tasks/${formatting.encodeQueryParameter(options.taskId)}/details`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    let url: string = options.url!;

    const taskDetails = await request.get<PlannerTaskDetails>(requestOptions);
    if (options.alias) {
      const urls: string[] = [];

      if (taskDetails.references) {
        Object.entries(taskDetails.references!).forEach((ref: any) => {
          if (ref[1].alias?.toLocaleLowerCase() === options.alias!.toLocaleLowerCase()) {
            urls.push(decodeURIComponent(ref[0]));
          }
        });
      }

      if (urls.length === 0) {
        throw `The specified reference with alias ${options.alias} does not exist`;
      }

      if (urls.length > 1) {
        throw `Multiple references with alias ${options.alias} found. Pass one of the following urls within the "--url" option : ${urls}`;
      }

      url = urls[0];
    }

    return { etag: (taskDetails as any)['@odata.etag'], url };
  }
}

export default new PlannerTaskReferenceRemoveCommand();
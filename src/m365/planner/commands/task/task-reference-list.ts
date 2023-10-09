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
}

class PlannerTaskReferenceListCommand extends GraphCommand {
  public get name(): string {
    return commands.TASK_REFERENCE_LIST;
  }

  public get description(): string {
    return 'Retrieve the references of the specified planner task';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initTypes();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --taskId <taskId>'
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('taskId');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/planner/tasks/${formatting.encodeQueryParameter(args.options.taskId)}/details?$select=references`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      const res = await request.get<any>(requestOptions);
      await logger.log(res.references);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PlannerTaskReferenceListCommand();

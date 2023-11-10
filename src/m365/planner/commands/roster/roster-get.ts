import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
}

class PlannerRosterGetCommand extends GraphCommand {
  public get name(): string {
    return commands.ROSTER_GET;
  }

  public get description(): string {
    return 'Retrieve information about a specific Microsoft Planner Roster';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initTypes();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--id <id>'
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('id');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving information about Microsoft Planner Roster with id ${args.options.id}`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/beta/planner/rosters/${args.options.id}`,
      headers: {
        'accept': 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      const response = await request.get(requestOptions);
      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PlannerRosterGetCommand();
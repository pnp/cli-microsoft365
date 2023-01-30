import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

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
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--id <id>'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving information about Microsoft Planner Roster with id ${args.options.id}`);
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
      logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new PlannerRosterGetCommand();
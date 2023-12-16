import { cli } from '../../../../cli/cli.js';
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
  force?: boolean;
}

class PlannerRosterRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.ROSTER_REMOVE;
  }

  public get description(): string {
    return 'Removes a Microsoft Planner Roster';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--id <id>'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('id');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removeRoster(args, logger);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove roster ${args.options.id}?` });

      if (result) {
        await this.removeRoster(args, logger);
      }
    }
  }

  private async removeRoster(args: CommandArgs, logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Removing roster ${args.options.id}`);
    }
    try {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/beta/planner/rosters/${args.options.id}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      await request.delete(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PlannerRosterRemoveCommand();
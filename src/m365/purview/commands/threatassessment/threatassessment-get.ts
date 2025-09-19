import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  withResults?: boolean;
}

class PurviewThreatAssessmentGetCommand extends GraphCommand {
  public get name(): string {
    return commands.THREATASSESSMENT_GET;
  }

  public get description(): string {
    return 'Get a threat assessment';
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
        includeResults: !!args.options.includeResults,
        withResults: !!args.options.withResults
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id <id>'
      },
      {
        option: '--withResults'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID.`;
        }
        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving threat assessment with id ${args.options.id}`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/informationProtection/threatAssessmentRequests/${args.options.id}${args.options.withResults ? '?$expand=results' : ''}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const res: any = await request.get(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PurviewThreatAssessmentGetCommand();
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { validation } from '../../../../utils/validation';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  includeResults?: boolean;
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
        includeResults: !!args.options.includeResults
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id <id>'
      },
      {
        option: '--includeResults'
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
        logger.logToStderr(`Retrieving threat assessment with id ${args.options.id}`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/informationProtection/threatAssessmentRequests/${args.options.id}${args.options.includeResults ? '?$expand=results' : ''}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const res: any = await request.get(requestOptions);
      logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new PurviewThreatAssessmentGetCommand();
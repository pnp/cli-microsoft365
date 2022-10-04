import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils/validation';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: DateTimeOptions;
}

interface DateTimeOptions extends GlobalOptions {
  fromDateTime: string;
  toDateTime?: string;
}

class TeamsReportDirectroutingcallsCommand extends GraphCommand {
  public get name(): string {
    return commands.REPORT_DIRECTROUTINGCALLS;
  }

  public get description(): string {
    return 'Get details about direct routing calls made within a given time period';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'calleeNumber', 'callerNumber', 'startDateTime'];
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
        toDateTime: typeof args.options.toDateTime !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--fromDateTime <fromDateTime>'
      },
      {
        option: '--toDateTime [toDateTime]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidISODateDashOnly(args.options.fromDateTime)) {
          return 'The fromDateTime is not a valid ISO date string (YYYY-MM-DD)';
        }

        if (args.options.toDateTime &&
          !validation.isValidISODateDashOnly(args.options.toDateTime)) {
          return 'The toDateTime is not a valid ISO date string (YYYY-MM-DD)';
        }

        if (Math.ceil((new Date(args.options.toDateTime || new Date().toISOString()).getTime() - new Date(args.options.fromDateTime).getTime()) / (1000 * 3600 * 24)) > 90) {
          return 'The maximum number of days between fromDateTime and toDateTime cannot exceed 90';
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const toDateTimeParameter: string = encodeURIComponent(args.options.toDateTime ? args.options.toDateTime : new Date().toISOString());

    const requestOptions: any = {
      url: `${this.resource}/v1.0/communications/callRecords/getDirectRoutingCalls(fromDateTime=${encodeURIComponent(args.options.fromDateTime)},toDateTime=${toDateTimeParameter})`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      const res: { value: any[] } = await request.get<{ value: any[] }>(requestOptions);
      logger.log(res);
    } 
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new TeamsReportDirectroutingcallsCommand();
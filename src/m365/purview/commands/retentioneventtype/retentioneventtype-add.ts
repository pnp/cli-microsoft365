import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  displayName: string;
  description?: string;
}

class PurviewRetentionEventTypeAddCommand extends GraphCommand {
  public get name(): string {
    return commands.RETENTIONEVENTTYPE_ADD;
  }

  public get description(): string {
    return 'Create a retention event type';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        description: typeof args.options.description !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --displayName <displayName>'
      },
      {
        option: '-d, --description [description]'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const requestBody = {
      displayName: args.options.displayName,
      description: args.options.description
    };

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/security/triggerTypes/retentionEventTypes`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      data: requestBody,
      responseType: 'json'
    };

    try {
      const response = await request.post(requestOptions);
      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataPromise(err);
    }
  }
}

export default new PurviewRetentionEventTypeAddCommand();
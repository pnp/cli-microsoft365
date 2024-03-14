import GlobalOptions from '../../../../GlobalOptions.js';
import { Logger } from '../../../../cli/Logger.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import request, { CliRequestOptions } from '../../../../request.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
}

class VivaEngageCommunityGetCommand extends GraphCommand {
  public get name(): string {
    return commands.ENGAGE_COMMUNITY_GET;
  }

  public get description(): string {
    return 'Gets information of a Viva Engage community';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: args.options.id !== undefined
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '-i, --id <id>' }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Getting the information of Viva Engage community with id '${args.options.id}'...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/beta/employeeExperience/communities/${args.options.id}`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const res: any = await request.get(requestOptions);

      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new VivaEngageCommunityGetCommand();
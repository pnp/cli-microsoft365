import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import PowerAutomateCommand from '../../../base/PowerAutomateCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environmentName: string;
  flowName: string;
  name: string;
  includeTriggerInformation?: boolean
}

class FlowRunGetCommand extends PowerAutomateCommand {
  public get name(): string {
    return commands.RUN_GET;
  }

  public get description(): string {
    return 'Gets information about a specific run of the specified Microsoft Flow';
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'startTime', 'endTime', 'status', 'triggerName'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name <name>'
      },
      {
        option: '--flowName <flowName>'
      },
      {
        option: '-e, --environmentName <environmentName>'
      },
      {
        option: '--includeTriggerInformation'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.flowName)) {
          return `${args.options.flowName} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        includeTriggerInformation: !!args.options.includeTriggerInformation
      });
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving information about run ${args.options.name} of Microsoft Flow ${args.options.flowName}...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/providers/Microsoft.ProcessSimple/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.flowName)}/runs/${formatting.encodeQueryParameter(args.options.name)}?api-version=2016-11-01`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    try {
      const res = await request.get<any>(requestOptions);

      res.startTime = res.properties.startTime;
      res.endTime = res.properties.endTime || '';
      res.status = res.properties.status;
      res.triggerName = res.properties.trigger.name;

      if (args.options.includeTriggerInformation && res.properties.trigger.outputsLink) {
        const triggerInformationOptions: CliRequestOptions = {
          url: res.properties.trigger.outputsLink.uri,
          headers: {
            accept: 'application/json',
            'x-anonymous': true
          },
          responseType: 'json'
        };
        const triggerInformationResponse = await request.get<any>(triggerInformationOptions);
        res.triggerInformation = triggerInformationResponse.body;
      }

      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new FlowRunGetCommand();
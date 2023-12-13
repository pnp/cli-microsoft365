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
  withTrigger?: boolean;
  withActions?: string | boolean;
}

interface FlowLink {
  uri: string;
}

interface Trigger {
  endTime: string
  name: string;
  originHistoryName: string;
  isAborted: boolean;
  inputsLink: FlowLink;
  outputsLink: FlowLink;
  sourceHistoryName: string
  startTime: string
  status: string
}

interface Action {
  [actionKey: string]: {
    code: string;
    endTime: string;
    startTime: string;
    status: string;
    inputsLink?: FlowLink;
    outputsLink?: FlowLink;
    input?: any;
    output?: any;
  }
}

interface Run {
  id: string;
  name: string;
  properties: {
    startTime: string,
    endTime: string,
    status: string,
    code: string,
    trigger: Trigger,
    actions?: Action
  },
  type: string;
}

interface RunResult extends Run {
  endTime?: string
  startTime?: string;
  status?: string;
  triggerName?: string;
  triggerInformation?: any;
  actions?: Action;
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
      },
      {
        option: '--withTrigger'
      },
      {
        option: '--withActions [withActions]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.flowName)) {
          return `${args.options.flowName} is not a valid GUID`;
        }

        if (args.options.withActions && (typeof args.options.withActions !== 'string' && typeof args.options.withActions !== 'boolean')) {
          return 'the withActions parameter must be a string or boolean';
        }

        return true;
      }
    );
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        includeTriggerInformation: !!args.options.includeTriggerInformation,
        withTrigger: !!args.options.withTrigger,
        withActions: typeof args.options.withActions !== 'undefined'
      });
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving information about run ${args.options.name} of Microsoft Flow ${args.options.flowName}...`);
    }

    if (args.options.includeTriggerInformation) {
      await this.warn(logger, `Parameter 'includeTriggerInformation' is deprecated. Please use 'withTrigger instead`);
    }

    const actionsParameter = args.options.withActions ? '$expand=properties%2Factions&' : '';
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/providers/Microsoft.ProcessSimple/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.flowName)}/runs/${formatting.encodeQueryParameter(args.options.name)}?${actionsParameter}api-version=2016-11-01`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    try {
      const res: RunResult = await request.get<Run>(requestOptions);
      res.startTime = res.properties.startTime;
      res.endTime = res.properties.endTime || '';
      res.status = res.properties.status;
      res.triggerName = res.properties.trigger.name;

      if ((args.options.includeTriggerInformation || args.options.withTrigger) && res.properties.trigger.outputsLink) {
        res.triggerInformation = await this.getTriggerInformation(res);
      }

      if (!!args.options.withActions) {
        res.actions = await this.getActionsInformation(res, args.options.withActions);
      }

      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getTriggerInformation(res: RunResult): Promise<any> {
    return await this.requestAdditionalInformation(res.properties.trigger.outputsLink.uri);
  }

  private async getActionsInformation(res: RunResult, withActions: boolean | string): Promise<any> {
    const chosenActions = typeof withActions === 'string' ? withActions.split(',') : null;
    const actionsResult: any = {};

    for (const action in res.properties.actions) {
      if (!res.properties.actions[action] || (chosenActions && chosenActions.indexOf(action) === -1)) { continue; }

      actionsResult[action] = res.properties.actions[action];
      if (!!res.properties.actions[action].inputsLink?.uri) {
        actionsResult[action].input = await this.requestAdditionalInformation(res.properties.actions[action].inputsLink?.uri);
      }

      if (!!res.properties.actions[action].outputsLink?.uri) {
        actionsResult[action].output = await this.requestAdditionalInformation(res.properties.actions[action].outputsLink?.uri);
      }
    }
    return actionsResult;
  }

  private async requestAdditionalInformation(requestUri: string | undefined): Promise<any> {
    const additionalInformationOptions: CliRequestOptions = {
      url: requestUri,
      headers: {
        accept: 'application/json',
        'x-anonymous': true
      },
      responseType: 'json'
    };
    const additionalInformationResponse = await request.get<any | string>(additionalInformationOptions);
    return additionalInformationResponse.body ? additionalInformationResponse.body : additionalInformationResponse;
  }
}

export default new FlowRunGetCommand();
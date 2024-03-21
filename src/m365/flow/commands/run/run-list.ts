import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import PowerAutomateCommand from '../../../base/PowerAutomateCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environmentName: string;
  flowName: string;
  status?: string;
  triggerStartTime?: string;
  triggerEndTime?: string;
  withTrigger?: boolean
  asAdmin?: boolean;
}

interface PowerAutomateFlowRun {
  name: string;
  startTime: string;
  status: string;
  properties: {
    startTime: string;
    status: string;
  }
}

class FlowRunListCommand extends PowerAutomateCommand {
  public readonly allowedStatusses: string[] = ['Succeeded', 'Running', 'Failed', 'Cancelled'];

  public get name(): string {
    return commands.RUN_LIST;
  }

  public get description(): string {
    return 'Lists runs of the specified Microsoft Flow';
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'startTime', 'status'];
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
        status: typeof args.options.status !== 'undefined',
        triggerStartTime: typeof args.options.triggerStartTime !== 'undefined',
        triggerEndTime: typeof args.options.triggerEndTime !== 'undefined',
        withTrigger: !!args.options.withTrigger,
        asAdmin: !!args.options.asAdmin
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--flowName <flowName>'
      },
      {
        option: '-e, --environmentName <environmentName>'
      },
      {
        option: '--status [status]',
        autocomplete: this.allowedStatusses
      },
      {
        option: '--triggerStartTime [triggerStartTime]'
      },
      {
        option: '--triggerEndTime [triggerEndTime]'
      },
      {
        option: '--withTrigger'
      },
      {
        option: '--asAdmin'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.flowName)) {
          return `${args.options.flowName} is not a valid GUID`;
        }

        if (args.options.status && this.allowedStatusses.indexOf(args.options.status) === -1) {
          return `'${args.options.status}' is not a valid status. Allowed values are: ${this.allowedStatusses.join(',')}`;
        }

        if (args.options.triggerStartTime && !validation.isValidISODateTime(args.options.triggerStartTime)) {
          return `'${args.options.triggerStartTime}' is not a valid datetime.`;
        }

        if (args.options.triggerEndTime && !validation.isValidISODateTime(args.options.triggerEndTime)) {
          return `'${args.options.triggerEndTime}' is not a valid datetime.`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving list of runs for Microsoft Flow ${args.options.flowName}...`);
    }

    let url: string = `${this.resource}/providers/Microsoft.ProcessSimple/${args.options.asAdmin ? 'scopes/admin/' : ''}environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.flowName)}/runs?api-version=2016-11-01`;
    const filters = this.getFilters(args.options);
    if (filters.length > 0) {
      url += `&$filter=${filters.join(' and ')}`;
    }
    try {
      const items = await odata.getAllItems<PowerAutomateFlowRun>(url);

      if (args.options.output === 'json' && args.options.withTrigger) {
        await this.retrieveTriggerInformation(items);
      }

      if (args.options.output !== 'json' && items.length > 0) {
        items.forEach(i => {
          i.startTime = i.properties.startTime;
          i.status = i.properties.status;
        });
      }

      await logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getFilters(options: Options): string[] {
    const filters = [];
    if (options.status) {
      filters.push(`status eq '${options.status}'`);
    }
    if (options.triggerStartTime) {
      filters.push(`startTime ge ${options.triggerStartTime}`);
    }
    if (options.triggerEndTime) {
      filters.push(`startTime lt ${options.triggerEndTime}`);
    }
    return filters;
  }

  private async retrieveTriggerInformation(items: PowerAutomateFlowRun[]): Promise<void> {
    const tasks = items.map(async (item: any) => {
      const requestOptions: CliRequestOptions = {
        url: item.properties.trigger.outputsLink.uri,
        headers: {
          accept: 'application/json',
          'x-anonymous': true
        },
        responseType: 'json'
      };
      const response = await request.get<{ body: any }>(requestOptions);
      item.triggerInformation = response.body;
    });

    await Promise.all(tasks);
  }
}

export default new FlowRunListCommand();
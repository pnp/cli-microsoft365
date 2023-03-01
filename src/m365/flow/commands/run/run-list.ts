import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { formatting } from '../../../../utils/formatting';
import { validation } from '../../../../utils/validation';
import { AzmgmtItemsListCommand } from '../../../base/AzmgmtItemsListCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environmentName: string;
  flowName: string;
  status?: string;
  triggerStartTime?: string;
  triggerEndTime?: string;
  asAdmin?: boolean;
}

class FlowRunListCommand extends AzmgmtItemsListCommand<{ name: string, startTime: string, status: string, properties: { startTime: string, status: string } }> {
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
        asAdmin: !!args.options.asAdmin
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-f, --flowName <flowName>'
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
      logger.logToStderr(`Retrieving list of runs for Microsoft Flow ${args.options.flowName}...`);
    }

    let url: string = `${this.resource}providers/Microsoft.ProcessSimple/${args.options.asAdmin ? 'scopes/admin/' : ''}environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.flowName)}/runs?api-version=2016-11-01`;
    const filters = this.getFilters(args.options);
    if (filters.length > 0) {
      url += `&$filter=${filters.join(' and ')}`;
    }
    try {
      await this.getAllItems(url, logger, true);

      if (this.items.length > 0) {
        this.items.forEach(i => {
          i.startTime = i.properties.startTime;
          i.status = i.properties.status;
        });

        logger.log(this.items);
      }
      else {
        if (this.verbose) {
          logger.logToStderr('No runs found');
        }
      }
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
}

module.exports = new FlowRunListCommand();
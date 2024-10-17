import { Logger } from '../../../cli/Logger.js';
import GlobalOptions from '../../../GlobalOptions.js';
import { formatting } from '../../../utils/formatting.js';
import { odata } from '../../../utils/odata.js';
import PowerAutomateCommand from '../../base/PowerAutomateCommand.js';
import commands from '../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environmentName: string;
  sharingStatus?: string;
  includeSolutions?: boolean;
  asAdmin?: boolean;
}

interface PowerAutomateFlow {
  name: string;
  id: string;
  displayName: string;
  properties: {
    displayName: string;
  }
}

class FlowListCommand extends PowerAutomateCommand {
  private allowedSharingStatuses = ['all', 'personal', 'ownedByMe', 'sharedWithMe'];

  public get name(): string {
    return commands.LIST;
  }

  public get description(): string {
    return 'Lists Power Automate flows in the given environment';
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'displayName'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        sharingStatus: typeof args.options.sharingStatus !== 'undefined',
        includeSolutions: !!args.options.includeSolutions,
        asAdmin: !!args.options.asAdmin
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-e, --environmentName <environmentName>'
      },
      {
        option: '--sharingStatus [sharingStatus]',
        autocomplete: this.allowedSharingStatuses
      },
      {
        option: '--includeSolutions'
      },
      {
        option: '--asAdmin'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.asAdmin && args.options.sharingStatus) {
          return `The options asAdmin and sharingStatus cannot be specified together.`;
        }

        if (args.options.sharingStatus && !this.allowedSharingStatuses.some(status => status === args.options.sharingStatus)) {
          return `${args.options.sharingStatus} is not a valid sharing status. Allowed values are: ${this.allowedSharingStatuses.join(', ')}`;
        }

        return true;
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('environmentName', 'sharingStatus');
    this.types.boolean.push('includeSolutions', 'asAdmin');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Getting Power Automate flows${args.options.asAdmin && ' as admin'} in environment '${args.options.environmentName}'...`);
    }

    try {
      const {
        environmentName,
        asAdmin,
        sharingStatus,
        includeSolutions
      } = args.options;

      let items: PowerAutomateFlow[] = [];

      if (sharingStatus === 'personal') {
        const url = this.getApiUrl(environmentName, asAdmin, includeSolutions, 'personal');
        items = await odata.getAllItems<PowerAutomateFlow>(url);
      }
      else if (sharingStatus === 'sharedWithMe') {
        const url = this.getApiUrl(environmentName, asAdmin, includeSolutions, 'team');
        items = await odata.getAllItems<PowerAutomateFlow>(url);
      }
      else if (sharingStatus === 'all') {
        let url = this.getApiUrl(environmentName, asAdmin, includeSolutions, 'personal');
        items = await odata.getAllItems<PowerAutomateFlow>(url);

        url = this.getApiUrl(environmentName, asAdmin, includeSolutions, 'team');
        const teamFlows = await odata.getAllItems<PowerAutomateFlow>(url);
        items = items.concat(teamFlows);
      }
      else {
        const url = this.getApiUrl(environmentName, asAdmin, includeSolutions);
        items = await odata.getAllItems<PowerAutomateFlow>(url);
      }

      // Remove duplicates
      items = items.filter((flow, index, self) =>
        index === self.findIndex(f => f.id === flow.id)
      );

      if (args.options.output && args.options.output !== 'json') {
        items.forEach(flow => {
          flow.displayName = flow.properties.displayName;
        });
      }
      await logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getApiUrl(environmentName: string, asAdmin?: boolean, includeSolutionFlows?: boolean, filter?: 'personal' | 'team',): string {
    const baseEndpoint = `${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple`;
    const environmentSegment = `/environments/${formatting.encodeQueryParameter(environmentName)}`;
    const adminSegment = `/scopes/admin${environmentSegment}/v2`;
    const flowsEndpoint = '/flows?api-version=2016-11-01';
    const filterQuery = filter === 'personal' || filter === 'team' ? `&$filter=search('${filter}')` : '';
    const includeQuery = includeSolutionFlows ? '&include=includeSolutionCloudFlows' : '';

    return `${baseEndpoint}${asAdmin ? adminSegment : environmentSegment}${flowsEndpoint}${filterQuery}${includeQuery}`;
  }
}

export default new FlowListCommand();
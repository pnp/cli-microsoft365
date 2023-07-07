import { Logger } from '../../../cli/Logger';
import GlobalOptions from '../../../GlobalOptions';
import { formatting } from '../../../utils/formatting';
import { AzmgmtItemsListCommand } from '../../base/AzmgmtItemsListCommand';
import commands from '../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environmentName: string;
  sharingStatus?: string;
  includeSolutions?: boolean;
  asAdmin?: boolean;
}

class FlowListCommand extends AzmgmtItemsListCommand<{ name: string, id: string, displayName: string, properties: { displayName: string } }> {
  private allowedSharingStatuses = ['all', 'personal', 'ownedByMe', 'sharedWithMe'];

  public get name(): string {
    return commands.LIST;
  }

  public get description(): string {
    return 'Lists Microsoft Flows in the given environment';
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'displayName'];
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
          return `${args.options.sharingStatus} is not a valid sharing status. Allowed values are: ${this.allowedSharingStatuses.join(',')}`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const {
        environmentName,
        asAdmin,
        sharingStatus,
        includeSolutions
      } = args.options;

      if (sharingStatus === 'personal') {
        const url = this.getApiUrl(environmentName, asAdmin, includeSolutions, 'personal');
        await this.getAllItems(url, logger, true);
      }
      else if (sharingStatus === 'sharedWithMe') {
        const url = this.getApiUrl(environmentName, asAdmin, includeSolutions, 'team');
        await this.getAllItems(url, logger, true);
      }
      else if (sharingStatus === 'all') {
        let url = this.getApiUrl(environmentName, asAdmin, includeSolutions, 'personal');
        await this.getAllItems(url, logger, true);

        url = this.getApiUrl(environmentName, asAdmin, includeSolutions, 'team');
        await this.getAllItems(url, logger, false);
      }
      else {
        const url = this.getApiUrl(environmentName, asAdmin, includeSolutions);
        await this.getAllItems(url, logger, true);
      }

      // Remove duplicates
      this.items = this.items.filter((flow, index, self) =>
        index === self.findIndex(f => f.id === flow.id)
      );

      if (this.items.length > 0) {
        this.items.forEach(i => {
          i.displayName = i.properties.displayName;
        });

        logger.log(this.items);
      }
      else {
        if (this.verbose) {
          logger.logToStderr('No Flows found');
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getApiUrl(environmentName: string, asAdmin?: boolean, includeSolutionFlows?: boolean, filter?: 'personal' | 'team',): string {
    let url = `${this.resource}providers/Microsoft.ProcessSimple${asAdmin ? '/scopes/admin' : ''}/environments/${formatting.encodeQueryParameter(environmentName)}/flows?api-version=2016-11-01`;

    if (filter === 'personal') {
      url += `&$filter=search('personal')`;
    }
    else if (filter === 'team') {
      url += `&$filter=search('team')`;
    }

    if (includeSolutionFlows) {
      url += '&include=includeSolutionCloudFlows';
    }

    return url;
  }
}

module.exports = new FlowListCommand();
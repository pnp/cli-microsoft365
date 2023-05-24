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
  asAdmin: boolean;
}

class FlowListCommand extends AzmgmtItemsListCommand<{ name: string, displayName: string, properties: { displayName: string } }> {
  private allowedSharingStatusses = ['all', 'personal', 'ownedByMe', 'sharedWithMe'];

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
        autocomplete: this.allowedSharingStatusses
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

        if (args.options.sharingStatus && !this.allowedSharingStatusses.some(status => status === args.options.sharingStatus)) {
          return `${args.options.sharingStatus} is not a valid sharing status. Allowed values are: ${this.allowedSharingStatusses.join(',')}`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const url: string = `${this.resource}providers/Microsoft.ProcessSimple${args.options.asAdmin ? '/scopes/admin' : ''}/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows?api-version=2016-11-01`;

    try {
      if (args.options.asAdmin || !args.options.sharingStatus || args.options.sharingStatus === 'ownedByMe') {
        await this.getAllItems(url, logger, true);
      }
      else if (args.options.sharingStatus === 'personal') {
        await this.getFilteredItems(url, logger, 'personal', true);
      }
      else if (args.options.sharingStatus === 'sharedWithMe') {
        await this.getFilteredItems(url, logger, 'team', true);
      }
      else {
        await this.getFilteredItems(url, logger, 'personal', true);
        await this.getFilteredItems(url, logger, 'team', false);
      }

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

  private async getFilteredItems(url: string, logger: Logger, filter: string, firstRun: boolean): Promise<void> {
    await this.getAllItems(`${url}&$filter=search('${filter}')`, logger, firstRun);
  }
}

module.exports = new FlowListCommand();
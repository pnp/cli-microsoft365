import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils/odata';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ListInstance } from './ListInstance';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  properties?: string;
  filter?: string;
}

interface FieldProperties {
  selectProperties: string[];
  expandProperties: string[];
}

class SpoListListCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_LIST;
  }

  public get description(): string {
    return 'Lists all available list in the specified site';
  }

  public defaultProperties(): string[] | undefined {
    return ['Title', 'Url', 'Id'];
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
    this.#initTelemetry();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-p, --properties [properties]'
      },
      {
        option: '-f, --filter [filter]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.webUrl)
    );
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        properties: typeof args.options.properties !== 'undefined',
        filter: typeof args.options.filter !== 'undefined'
      });
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving all lists in site at ${args.options.webUrl}...`);
    }

    try {
      const fieldProperties = this.formatSelectProperties(args.options.properties);
      const queryParams = [`$expand=${fieldProperties.expandProperties.join(',')}`, `$select=${fieldProperties.selectProperties.join(',')}`];

      if (args.options.filter) {
        queryParams.push(`$filter=${args.options.filter}`);
      }

      const listInstances = await odata.getAllItems<ListInstance>(`${args.options.webUrl}/_api/web/lists?${queryParams.join('&')}`);
      listInstances.forEach(l => {
        l.Url = l.RootFolder.ServerRelativeUrl;
      });

      logger.log(listInstances);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private formatSelectProperties(fields: string | undefined): FieldProperties {
    const selectProperties: any[] = ['RootFolder/ServerRelativeUrl'];
    const expandProperties: any[] = ['RootFolder'];

    if (!fields) {
      selectProperties.push('*');
    }

    if (fields) {
      fields.split(',').forEach((field) => {
        const subparts = field.trim().split('/');
        if (subparts.length > 1) {
          expandProperties.push(subparts[0]);
        }
        selectProperties.push(field.trim());
      });
    }

    return {
      selectProperties: [...new Set(selectProperties)],
      expandProperties: [...new Set(expandProperties)]
    };
  }
}

module.exports = new SpoListListCommand();
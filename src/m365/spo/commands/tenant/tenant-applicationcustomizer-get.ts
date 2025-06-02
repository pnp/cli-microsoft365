import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { cli } from '../../../../cli/cli.js';
import { ListItemInstance } from '../listitem/ListItemInstance.js';
import { ListItemListOptions, spoListItem } from '../../../../utils/spoListItem.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  title?: string;
  id?: string;
  clientSideComponentId?: string;
  tenantWideExtensionComponentProperties?: boolean;
}

class SpoTenantApplicationCustomizerGetCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_APPLICATIONCUSTOMIZER_GET;
  }

  public get description(): string {
    return 'Get an application customizer that is installed tenant wide';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        title: typeof args.options.title !== 'undefined',
        id: typeof args.options.id !== 'undefined',
        clientSideComponentId: typeof args.options.clientSideComponentId !== 'undefined',
        tenantWideExtensionComponentProperties: !!args.options.tenantWideExtensionComponentProperties
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-t, --title [title]'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '-c, --clientSideComponentId [clientSideComponentId]'
      },
      {
        option: '-p, --tenantWideExtensionComponentProperties'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && isNaN(parseInt(args.options.id))) {
          return `${args.options.id} is not a number`;
        }

        if (args.options.clientSideComponentId && !validation.isValidGuid(args.options.clientSideComponentId)) {
          return `${args.options.clientSideComponentId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['title', 'id', 'clientSideComponentId'] });
  }

  #initTypes(): void {
    this.types.string.push('title', 'id', 'clientSideComponentId');
    this.types.boolean.push('tenantWideExtensionComponentProperties');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const appCatalogUrl = await spo.getTenantAppCatalogUrl(logger, this.debug);

      if (!appCatalogUrl) {
        throw 'No app catalog URL found';
      }

      let filter: string;
      if (args.options.title) {
        filter = `Title eq '${args.options.title}'`;
      }
      else if (args.options.id) {
        filter = `Id eq ${args.options.id}`;
      }
      else {
        filter = `TenantWideExtensionComponentId eq '${args.options.clientSideComponentId}'`;
      }

      const options: ListItemListOptions = {
        webUrl: appCatalogUrl,
        listUrl: '/Lists/TenantWideExtensions',
        filter: `TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer' and ${filter}`
      };

      const listItems = await spoListItem.getListItems(options, logger, this.verbose);

      if (listItems) {
        if (listItems.length === 0) {
          throw 'The specified application customizer was not found';
        }

        let listItemInstance: ListItemInstance;
        if (listItems.length > 1) {
          const resultAsKeyValuePair = formatting.convertArrayToHashTable('Id', listItems);
          listItemInstance = await cli.handleMultipleResultsFound<ListItemInstance>(`Multiple application customizers with ${args.options.title || args.options.clientSideComponentId} were found.`, resultAsKeyValuePair);
        }
        else {
          listItemInstance = listItems[0];
        }

        if (!args.options.tenantWideExtensionComponentProperties) {
          await logger.log(listItemInstance);
        }
        else {
          const properties = formatting.tryParseJson((listItemInstance as any).TenantWideExtensionComponentProperties);
          await logger.log(properties);
        }
      }
      else {
        throw 'The specified application customizer was not found';
      }
    }
    catch (err: any) {
      return this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoTenantApplicationCustomizerGetCommand();
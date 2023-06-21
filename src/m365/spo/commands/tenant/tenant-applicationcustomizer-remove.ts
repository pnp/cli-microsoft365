import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { ListItemInstance } from '../listitem/ListItemInstance.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  title?: string;
  id?: string;
  clientSideComponentId?: string;
  force?: boolean;
}

class SpoTenantApplicationCustomizerRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_APPLICATIONCUSTOMIZER_REMOVE;
  }

  public get description(): string {
    return 'Removes an application customizer that is installed tenant wide.';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        title: typeof args.options.title !== 'undefined',
        id: typeof args.options.id !== 'undefined',
        clientSideComponentId: typeof args.options.clientSideComponentId !== 'undefined',
        force: !!args.options.force
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
        option: '-c, --clientSideComponentId  [clientSideComponentId]'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id) {
          const id: number = parseInt(args.options.id);
          if (isNaN(id)) {
            return `${args.options.id} is not a valid list item ID`;
          }
        }

        if (args.options.clientSideComponentId &&
          !validation.isValidGuid(args.options.clientSideComponentId as string)) {
          return `${args.options.clientSideComponentId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['title', 'id', 'clientSideComponentId'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (args.options.force) {
        return await this.removeTenantApplicationCustomizer(logger, args);
      }

      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the tenant applicationcustomizer ${args.options.id || args.options.title || args.options.clientSideComponentId}?`
      });

      if (result.continue) {
        await this.removeTenantApplicationCustomizer(logger, args);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  public async getTenantApplicationCustomizerId(logger: Logger, args: CommandArgs, requestUrl: string): Promise<number> {
    if (this.verbose) {
      await logger.logToStderr(`Getting the tenant application customizer ${args.options.id || args.options.title || args.options.clientSideComponentId}`);
    }

    const filter: string[] = [`TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer'`];
    if (args.options.title) {
      filter.push(`Title eq '${args.options.title}'`);
    }
    else if (args.options.id) {
      filter.push(`Id eq '${args.options.id}'`);
    }
    else if (args.options.clientSideComponentId) {
      filter.push(`TenantWideExtensionComponentId eq '${args.options.clientSideComponentId}'`);
    }

    const listItemInstances: ListItemInstance[] = await odata.getAllItems(`${requestUrl}/items?$filter=${filter.join(' and ')}&$select=Id`);

    if (listItemInstances.length === 0) {
      throw 'The specified application customizer was not found';
    }

    if (listItemInstances.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('Id', listItemInstances);
      listItemInstances[0] = await Cli.handleMultipleResultsFound<ListItemInstance>(`Multiple application customizers with ${args.options.title || args.options.clientSideComponentId} were found. Choose the correct ID:`, `Multiple application customizers with ${args.options.title || args.options.clientSideComponentId} were found. Please disambiguate (IDs): ${listItemInstances.map(item => item.Id).join(', ')}`, resultAsKeyValuePair);
    }

    return listItemInstances[0].Id;
  }

  private async removeTenantApplicationCustomizer(logger: Logger, args: CommandArgs): Promise<void> {
    const appCatalogUrl = await spo.getTenantAppCatalogUrl(logger, this.debug);

    if (!appCatalogUrl) {
      throw 'No app catalog URL found';
    }

    const listServerRelativeUrl: string = urlUtil.getServerRelativePath(appCatalogUrl, '/lists/TenantWideExtensions');
    const requestUrl = `${appCatalogUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
    const id = await this.getTenantApplicationCustomizerId(logger, args, requestUrl);

    if (this.verbose) {
      await logger.logToStderr(`Removing tenant application customizer ${args.options.id || args.options.title || args.options.clientSideComponentId}`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${requestUrl}/items(${id})`,
      method: 'POST',
      headers: {
        'X-HTTP-Method': 'DELETE',
        'If-Match': '*',
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    await request.post(requestOptions);
  }
}

export default new SpoTenantApplicationCustomizerRemoveCommand();
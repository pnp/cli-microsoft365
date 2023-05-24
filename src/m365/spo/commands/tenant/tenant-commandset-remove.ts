import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { odata } from '../../../../utils/odata';
import { spo } from '../../../../utils/spo';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ListItemInstance } from '../listitem/ListItemInstance';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  title?: string;
  id?: string;
  clientSideComponentId?: string;
  confirm?: boolean;
}

class SpoTenantCommandSetRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_COMMANDSET_REMOVE;
  }

  public get description(): string {
    return 'Removes a ListView Command Set that is installed tenant wide.';
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
        confirm: !!args.options.confirm
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
        option: '--confirm'
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
      if (args.options.confirm) {
        return await this.removeTenantCommandSet(logger, args);
      }

      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the tenant commandset ${args.options.id || args.options.title || args.options.clientSideComponentId}?`
      });

      if (result.continue) {
        await this.removeTenantCommandSet(logger, args);
      }
    }
    catch (err: any) {
      logger.log(err);
      this.handleRejectedODataJsonPromise(err);
    }
  }

  public async getTenantCommandSet(logger: Logger, args: CommandArgs, requestUrl: string): Promise<number> {
    if (this.verbose) {
      logger.logToStderr(`Getting the tenant command set ${args.options.id || args.options.title || args.options.clientSideComponentId}`);
    }

    const filter: string[] = [`startswith(TenantWideExtensionLocation,'ClientSideExtension.ListViewCommandSet')`];
    if (args.options.title) {
      filter.push(`Title eq '${args.options.title}'`);
    }
    else if (args.options.id) {
      filter.push(`Id eq ${args.options.id}`);
    }
    else if (args.options.clientSideComponentId) {
      filter.push(`TenantWideExtensionComponentId eq '${args.options.clientSideComponentId}'`);
    }

    const listItemInstances: ListItemInstance[] = await odata.getAllItems<ListItemInstance>(`${requestUrl}/items?$filter=${filter.join(' and ')}`);

    if (listItemInstances.length === 0) {
      throw 'The specified command set was not found';
    }

    if (listItemInstances.length > 1) {
      throw `Multiple command sets with ${args.options.title || args.options.clientSideComponentId} were found. Please disambiguate (IDs): ${listItemInstances.map(item => item.Id).join(', ')}`;
    }

    return listItemInstances[0].Id;
  }

  private async removeTenantCommandSet(logger: Logger, args: CommandArgs): Promise<void> {
    const appCatalogUrl = await spo.getTenantAppCatalogUrl(logger, this.debug);
    if (!appCatalogUrl) {
      throw 'No app catalog URL found';
    }

    const listServerRelativeUrl: string = urlUtil.getServerRelativePath(appCatalogUrl, '/lists/TenantWideExtensions');

    const id = await this.getTenantCommandSet(logger, args, `${appCatalogUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`);

    if (this.verbose) {
      logger.logToStderr(`Removing tenant command set ${args.options.id || args.options.title || args.options.clientSideComponentId}`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${appCatalogUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/items(${id})`,
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

module.exports = new SpoTenantCommandSetRemoveCommand();
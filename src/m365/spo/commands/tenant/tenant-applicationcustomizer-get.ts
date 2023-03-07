import { Logger } from '../../../../cli/Logger';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { spo } from '../../../../utils/spo';
import { CommandError } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ListItemInstanceCollection } from '../listitem/ListItemInstanceCollection';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  title?: string;
  id?: string;
  clientSideComponentId?: string;
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
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        title: typeof args.options.title !== 'undefined',
        id: typeof args.options.id !== 'undefined',
        clientSideComponentId: typeof args.options.clientSideComponentId !== 'undefined'
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
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
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

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const appCatalogUrl = await spo.getTenantAppCatalogUrl(logger, this.debug);
    if (!appCatalogUrl) {
      throw new CommandError('No app catalog URL found');
    }

    let filter: string = '';
    if (args.options.title) {
      filter = `Title eq '${args.options.title}'`;
    }
    else if (args.options.id) {
      filter = `GUID eq '${args.options.id}'`;
    }
    else if (args.options.clientSideComponentId) {
      filter = `TenantWideExtensionComponentId eq '${args.options.clientSideComponentId}'`;
    }

    const listServerRelativeUrl: string = urlUtil.getServerRelativePath(appCatalogUrl, '/lists/TenantWideExtensions');
    const reqOptions: CliRequestOptions = {
      url: `${appCatalogUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/items?$filter=${filter}`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    let listItemInstances: ListItemInstanceCollection | undefined;
    try {
      listItemInstances = await request.get<ListItemInstanceCollection>(reqOptions);
    }
    catch (err: any) {
      return this.handleRejectedODataJsonPromise(err);
    }

    if (listItemInstances) {
      if (listItemInstances.value.length === 0) {
        throw new CommandError('The specified application customizer was not found');
      }

      if (listItemInstances.value.length > 1) {
        throw new CommandError(`Multiple application customizers with ${args.options.title || args.options.clientSideComponentId} were found. Please disambiguate (IDs): ${listItemInstances.value.map(item => item.GUID).join(', ')}`);
      }

      listItemInstances.value.forEach(v => delete v['ID']);

      logger.log(listItemInstances.value[0]);
    }
    else {
      throw new CommandError('The specified application customizer was not found');
    }
  }
}

module.exports = new SpoTenantApplicationCustomizerGetCommand();
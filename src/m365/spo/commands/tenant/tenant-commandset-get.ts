import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { ListItemInstanceCollection } from '../listitem/ListItemInstanceCollection.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  title?: string;
  id?: string;
  clientSideComponentId?: string;
}

class SpoTenantCommandSetGetCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_COMMANDSET_GET;
  }

  public get description(): string {
    return 'Get a ListView Command Set that is installed tenant wide';
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

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const appCatalogUrl = await spo.getTenantAppCatalogUrl(logger, this.debug);
    if (!appCatalogUrl) {
      throw new CommandError('No app catalog URL found');
    }

    let filter: string = `startswith(TenantWideExtensionLocation,'ClientSideExtension.ListViewCommandSet')`;

    if (args.options.title) {
      filter += ` and Title eq '${args.options.title}'`;
    }
    else if (args.options.id) {
      filter += ` and Id eq ${args.options.id}`;
    }
    else if (args.options.clientSideComponentId) {
      filter += ` and TenantWideExtensionComponentId eq '${args.options.clientSideComponentId}'`;
    }

    const listServerRelativeUrl: string = urlUtil.getServerRelativePath(appCatalogUrl, '/lists/TenantWideExtensions');
    const reqOptions: CliRequestOptions = {
      url: `${appCatalogUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/items?$filter=${filter}`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const listItemInstances = await request.get<ListItemInstanceCollection>(reqOptions);

      if (listItemInstances?.value.length > 0) {
        if (listItemInstances.value.length > 1) {
          throw `Multiple ListView Command Sets with ${args.options.title || args.options.clientSideComponentId} were found. Please disambiguate (IDs): ${listItemInstances.value.map(item => item.Id).join(', ')}`;
        }

        listItemInstances.value.forEach(v => delete v['ID']);

        await logger.log(listItemInstances.value[0]);
      }
      else {
        throw 'The specified ListView Command Set was not found';
      }
    }
    catch (err: any) {
      return this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoTenantCommandSetGetCommand();
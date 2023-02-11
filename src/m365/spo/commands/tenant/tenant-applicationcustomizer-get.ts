import { Logger } from '../../../../cli/Logger';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { spo } from '../../../../utils/spo';
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
    try {
      const spoUrl: string = await spo.getSpoUrl(logger, this.debug);
      console.log(`${spoUrl}/_api/SP_TenantSettings_Current`);
      const requestOptions: any = {
        url: `${spoUrl}/_api/SP_TenantSettings_Current`,
        headers: {
          accept: 'application/json;odata=nometadata'
        }
      };

      const res: string = await request.get(requestOptions);
      const json = JSON.parse(res);
      console.log(json);
      const appCatalogUrl: string | undefined = json.CorporateCatalogUrl;

      if (!appCatalogUrl) {
        throw 'No app catalog URL found';
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
      console.log(`${appCatalogUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/items?$filter=${filter}`);
      const reqOptions: CliRequestOptions = {
        url: `${appCatalogUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/items?$filter=${filter}`,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const listItemInstances = await request.get<ListItemInstanceCollection>(reqOptions);
      console.log(listItemInstances);
      listItemInstances.value.forEach(v => delete v['ID']);
      console.log(listItemInstances);

      if (listItemInstances.value.length === 0) {
        throw 'The specified application customizer was not found';
      }

      if (listItemInstances.value.length > 1) {
        throw `Multiple application customizers with ${args.options.title || args.options.clientSideComponentId} was found. Please disambiguate (IDs): ${listItemInstances.value.map(item => item.GUID).join(', ')}`;
      }

      logger.log(listItemInstances.value[0]);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoTenantApplicationCustomizerGetCommand();
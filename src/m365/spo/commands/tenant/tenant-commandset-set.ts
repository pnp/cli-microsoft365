import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { urlUtil } from '../../../../utils/urlUtil';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { spo } from '../../../../utils/spo';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  newTitle?: string;
  listType?: string;
  clientSideComponentId?: string;
  clientSideComponentProperties?: string;
  webTemplate?: string;
  location?: string;
}

class SpoTenantCommandSetSetCommand extends SpoCommand {
  private static readonly listTypes: string[] = ['List', 'Library', 'SitePages'];
  private static readonly locations: string[] = ['ContextMenu', 'CommandBar', 'Both'];

  public get name(): string {
    return commands.TENANT_COMMANDSET_SET;
  }

  public get description(): string {
    return 'Updates a ListView Command Set that is installed tenant wide.';
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
        newTitle: typeof args.options.newTitle !== 'undefined',
        listType: args.options.listType,
        clientSideComponentId: typeof args.options.clientSideComponentId !== 'undefined',
        clientSideComponentProperties: typeof args.options.clientSideComponentProperties !== 'undefined',
        webTemplate: typeof args.options.webTemplate !== 'undefined',
        location: args.options.location
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id <id>'
      },
      {
        option: '-t, --newTitle [newTitle]'
      },
      {
        option: '-l, --listType [listType]',
        autocomplete: SpoTenantCommandSetSetCommand.listTypes
      },
      {
        option: '-c, --clientSideComponentId [clientSideComponentId]'
      },
      {
        option: '-p, --clientSideComponentProperties [clientSideComponentProperties]'
      },
      {
        option: '-w, --webTemplate [webTemplate]'
      },
      {
        option: '--location [location]',
        autocomplete: SpoTenantCommandSetSetCommand.locations
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!args.options.newTitle &&
          !args.options.listType &&
          !args.options.clientSideComponentId &&
          !args.options.clientSideComponentProperties &&
          !args.options.webTemplate &&
          !args.options.location) {
          return 'Specify at least one property to update';
        }

        if (args.options.clientSideComponentId && !validation.isValidGuid(args.options.clientSideComponentId)) {
          return `${args.options.clientSideComponentId} is not a valid GUID`;
        }

        if (args.options.listType && SpoTenantCommandSetSetCommand.listTypes.indexOf(args.options.listType) < 0) {
          return `${args.options.listType} is not a valid list type. Allowed values are ${SpoTenantCommandSetSetCommand.listTypes.join(', ')}`;
        }

        if (args.options.location && SpoTenantCommandSetSetCommand.locations.indexOf(args.options.location) < 0) {
          return `${args.options.location} is not a valid location. Allowed values are ${SpoTenantCommandSetSetCommand.locations.join(', ')}`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const appCatalogUrl = await spo.getTenantAppCatalogUrl(logger, this.debug);

      if (!appCatalogUrl) {
        throw 'No app catalog URL found';
      }

      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(appCatalogUrl, '/lists/TenantWideExtensions');
      const listItem = await this.getListItemById(logger, appCatalogUrl, listServerRelativeUrl, args.options.id);

      if (listItem.TenantWideExtensionLocation.indexOf("ClientSideExtension.ListViewCommandSet") === -1) {
        throw 'The item is not a ListViewCommandSet';
      }

      await this.updateTenantWideExtension(appCatalogUrl, args.options, listServerRelativeUrl, logger);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getListItemById(logger: Logger, webUrl: string, listServerRelativeUrl: string, id: string): Promise<any> {
    if (this.verbose) {
      logger.logToStderr(`Getting the list item by id ${id}`);
    }
    const reqOptions: CliRequestOptions = {
      url: `${webUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/Items(${id})`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return await request.get<any>(reqOptions);
  }

  private async updateTenantWideExtension(appCatalogUrl: string, options: Options, listServerRelativeUrl: string, logger: Logger): Promise<void> {
    if (this.verbose) {
      logger.logToStderr('Updating tenant wide extension to the TenantWideExtensions list');
    }

    const formValues: any = [];
    if (options.newTitle !== undefined) {
      formValues.push({
        FieldName: 'Title',
        FieldValue: options.newTitle
      });
    }

    if (options.clientSideComponentId !== undefined) {
      formValues.push({
        FieldName: 'TenantWideExtensionComponentId',
        FieldValue: options.clientSideComponentId
      });
    }

    if (options.location !== undefined) {
      formValues.push({
        FieldName: 'TenantWideExtensionLocation',
        FieldValue: this.getLocation(options.location)
      });
    }

    if (options.listType !== undefined) {
      formValues.push({
        FieldName: 'TenantWideExtensionListTemplate',
        FieldValue: this.getListTemplate(options.listType)
      });
    }

    if (options.clientSideComponentProperties !== undefined) {
      formValues.push({
        FieldName: 'TenantWideExtensionComponentProperties',
        FieldValue: options.clientSideComponentProperties
      });
    }

    if (options.webTemplate !== undefined) {
      formValues.push({
        FieldName: 'TenantWideExtensionWebTemplate',
        FieldValue: options.webTemplate
      });
    }

    const requestOptions: CliRequestOptions =
    {
      url: `${appCatalogUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/Items(${options.id})/ValidateUpdateListItem()`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      data: {
        formValues: formValues
      },
      responseType: 'json'
    };

    await request.post(requestOptions);
  }

  private getLocation(location: string | undefined): string {
    switch (location) {
      case 'Both':
        return 'ClientSideExtension.ListViewCommandSet';
      case 'ContextMenu':
        return 'ClientSideExtension.ListViewCommandSet.ContextMenu';
      default:
        return 'ClientSideExtension.ListViewCommandSet.CommandBar';
    }
  }

  private getListTemplate(listTemplate: string | undefined): string {
    switch (listTemplate) {
      case 'SitePages':
        return '119';
      case 'Library':
        return '101';
      default:
        return '100';
    }
  }
}

module.exports = new SpoTenantCommandSetSetCommand();
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import Command from '../../../../Command';
import { Cli } from '../../../../cli/Cli';
import { Options as spoListItemSetCommandOptions } from '../listitem/listitem-set';
import * as spoListItemSetCommand from '../listitem/listitem-set';
import { urlUtil } from '../../../../utils/urlUtil';
import request, { CliRequestOptions } from '../../../../request';
import { ListItemInstanceCollection } from '../listitem/ListItemInstanceCollection';
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
    return 'Update a ListView Command Set that is installed tenant wide.';
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
        option: '-i, --clientSideComponentId [clientSideComponentId]'
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

      await this.updateTenantWideExtension(appCatalogUrl, args.options, logger);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  public async getTenantCommandSet(logger: Logger, options: Options, requestUrl: string): Promise<number> {
    if (this.verbose) {
      logger.logToStderr(`Getting the tenant command set ${options.id}`);
    }

    const filter: string = `startswith(TenantWideExtensionLocation,'ClientSideExtension.ListViewCommandSet') and Id eq '${options.id}'`;

    const reqOptions: CliRequestOptions = {
      url: `${requestUrl}/items?$filter=${filter}`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const listItemInstances: ListItemInstanceCollection = await request.get<ListItemInstanceCollection>(reqOptions);

    if (listItemInstances.value.length === 0) {
      throw 'The specified command set was not found';
    }

    return listItemInstances.value[0].Id;
  }

  private async updateTenantWideExtension(appCatalogUrl: string, options: Options, logger: Logger): Promise<void> {
    if (this.verbose) {
      logger.logToStderr('Updating tenant wide extension to the TenantWideExtensions list');
    }

    const listServerRelativeUrl: string = urlUtil.getServerRelativePath(appCatalogUrl, '/lists/TenantWideExtensions');
    const requestUrl = `${appCatalogUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
    const id = await this.getTenantCommandSet(logger, options, requestUrl);

    const commandOptions: spoListItemSetCommandOptions = {
      id: id.toString(),
      webUrl: appCatalogUrl,
      listUrl: `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/Lists/TenantWideExtensions`
    };

    if (options.newTitle) {
      commandOptions.Title = options.newTitle;
    }

    if (options.clientSideComponentId) {
      commandOptions.TenantWideExtensionComponentId = options.clientSideComponentId;
    }

    if (options.location) {
      commandOptions.TenantWideExtensionLocation = this.getLocation(options.location);
    }

    if (options.listType) {
      commandOptions.TenantWideExtensionListTemplate = this.getListTemplate(options.listType);
    }

    if (options.clientSideComponentProperties) {
      commandOptions.TenantWideExtensionComponentProperties = options.clientSideComponentProperties;
    }

    if (options.webTemplate) {
      commandOptions.TenantWideExtensionWebTemplate = options.webTemplate;
    }

    await Cli.executeCommand(spoListItemSetCommand as Command, { options: { ...commandOptions, _: [] } });
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
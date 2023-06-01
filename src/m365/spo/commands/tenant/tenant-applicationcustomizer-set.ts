import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { formatting } from '../../../../utils/formatting';
import { odata } from '../../../../utils/odata';
import { spo } from '../../../../utils/spo';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import * as os from 'os';
import { Options as spoListItemSetOptions } from '../listitem/listitem-set';
import * as spoListItemSet from '../listitem/listitem-set';
import Command from '../../../../Command';
import { ListItemInstance } from '../listitem/ListItemInstance';
import { Cli } from '../../../../cli/Cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  title?: string;
  clientSideComponentId?: string;
  newTitle?: string;
  clientSideComponentProperties?: string;
  webTemplate?: string;
}

class SpoTenantApplicationCustomizerSetCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_APPLICATIONCUSTOMIZER_SET;
  }

  public get description(): string {
    return 'Update an Application Customizer that is deployed as a tenant-wide extension';
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
        newTitle: typeof args.options.newTitle !== 'undefined',
        clientSideComponentProperties: typeof args.options.clientSideComponentProperties !== 'undefined',
        webTemplate: typeof args.options.webTemplate !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id [id]'
      },
      {
        option: '-t, --title [title]'
      },
      {
        option: '-c, --clientSideComponentId  [clientSideComponentId]'
      },
      {
        option: '--newTitle [newTitle]'
      },
      {
        option: '-p, --clientSideComponentProperties [clientSideComponentProperties]'
      },
      {
        option: '-w, --webTemplate [webTemplate]'
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

        if (!args.options.newTitle && !args.options.clientSideComponentProperties && !args.options.webTemplate) {
          return `Please specify an option to be updated`;
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
      const appCatalogUrl = await spo.getTenantAppCatalogUrl(logger, this.debug);
      if (!appCatalogUrl) {
        throw 'No app catalog URL found';
      }

      const listItem = await this.getListItem(logger, args.options, appCatalogUrl);
      await this.updateTenantWideExtension(appCatalogUrl, args.options, logger, listItem);
    }
    catch (err: any) {
      return this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getListItem(logger: Logger, options: Options, appCatalogUrl: string): Promise<ListItemInstance> {
    const { title, id, clientSideComponentId } = options;
    const filter = title ? `Title eq '${title}'` : id ? `Id eq '${id}'` : `TenantWideExtensionComponentId eq '${clientSideComponentId}'`;

    if (this.verbose) {
      logger.logToStderr(`Getting tenant-wide application customizer: "${title || id || clientSideComponentId}"...`);
    }

    const listServerRelativeUrl: string = urlUtil.getServerRelativePath(appCatalogUrl, '/lists/TenantWideExtensions');
    const listItemInstances = await odata.getAllItems<ListItemInstance>(`${appCatalogUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/items?$filter=TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer' and ${filter}`);
    if (listItemInstances) {
      if (listItemInstances.length === 0) {
        throw 'The specified application customizer was not found';
      }

      if (listItemInstances.length > 1) {
        throw `Multiple application customizers with ${title ? `title '${title}'` : `ClientSideComponentId '${clientSideComponentId}'`} found. Please disambiguate using IDs: ${os.EOL}${listItemInstances.map(item => `- ${(item as any).Id}`).join(os.EOL)}`;
      }
    }
    else {
      throw 'The specified application customizer was not found';
    }
    return listItemInstances[0];
  }

  private async updateTenantWideExtension(appCatalogUrl: string, options: Options, logger: Logger, listItem: ListItemInstance): Promise<void> {
    const { title, id, clientSideComponentId, newTitle, clientSideComponentProperties, webTemplate } = options;

    if (this.verbose) {
      logger.logToStderr(`Updating tenant-wide application customizer: "${title || id || clientSideComponentId}"...`);
    }

    const itemId = listItem.Id;
    const commandOptions: spoListItemSetOptions = {
      webUrl: appCatalogUrl,
      listUrl: `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/Lists/TenantWideExtensions`,
      id: itemId.toString(),
      ...(newTitle && { Title: newTitle }),
      ...(clientSideComponentProperties && { TenantWideExtensionComponentProperties: clientSideComponentProperties }),
      ...(webTemplate && { TenantWideExtensionWebTemplate: webTemplate })
    };

    await Cli.executeCommandWithOutput(spoListItemSet as Command, { options: { ...commandOptions, _: [] } });
  }
}

module.exports = new SpoTenantApplicationCustomizerSetCommand();

import GlobalOptions from '../../../../GlobalOptions.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  title?: string;
  clientSideComponentId?: string;
  newTitle?: string;
  listType?: string;
  newClientSideComponentId?: string;
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
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        title: typeof args.options.title !== 'undefined',
        id: typeof args.options.id !== 'undefined',
        clientSideComponentId: typeof args.options.clientSideComponentId !== 'undefined',
        newTitle: typeof args.options.newTitle !== 'undefined',
        listType: args.options.listType,
        newClientSideComponentId: typeof args.options.newClientSideComponentId !== 'undefined',
        clientSideComponentProperties: typeof args.options.clientSideComponentProperties !== 'undefined',
        webTemplate: typeof args.options.webTemplate !== 'undefined',
        location: args.options.location
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
        option: '-c, --clientSideComponentId [clientSideComponentId]'
      },
      {
        option: '--newTitle [newTitle]'
      },
      {
        option: '-l, --listType [listType]',
        autocomplete: SpoTenantCommandSetSetCommand.listTypes
      },
      {
        option: '--newClientSideComponentId [newClientSideComponentId]'
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
          !args.options.newClientSideComponentId &&
          !args.options.clientSideComponentProperties &&
          !args.options.webTemplate &&
          !args.options.location) {
          return 'Specify at least one property to update';
        }

        if (args.options.id && isNaN(parseInt(args.options.id))) {
          return `${args.options.id} is not a number`;
        }

        if (args.options.clientSideComponentId && !validation.isValidGuid(args.options.clientSideComponentId)) {
          return `${args.options.clientSideComponentId} is not a valid GUID`;
        }

        if (args.options.newClientSideComponentId && !validation.isValidGuid(args.options.newClientSideComponentId)) {
          return `${args.options.newClientSideComponentId} is not a valid GUID`;
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

  #initOptionSets(): void {
    this.optionSets.push({ options: ['title', 'id', 'clientSideComponentId'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const appCatalogUrl = await spo.getTenantAppCatalogUrl(logger, this.debug);

      if (!appCatalogUrl) {
        throw 'No app catalog URL found';
      }

      if (args.options.newClientSideComponentId !== undefined) {
        const componentManifest = await this.getComponentManifest(appCatalogUrl, args.options.newClientSideComponentId, logger);
        const clientComponentManifest = JSON.parse(componentManifest.ClientComponentManifest);

        if (clientComponentManifest.extensionType !== "ListViewCommandSet") {
          throw `The extension type of this component is not of type 'ListViewCommandSet' but of type '${clientComponentManifest.extensionType}'`;
        }

        const solution = await this.getSolutionFromAppCatalog(appCatalogUrl, componentManifest.SolutionId, logger);

        if (!solution.ContainsTenantWideExtension) {
          throw `The solution does not contain an extension that can be deployed to all sites. Make sure that you've entered the correct component Id.`;
        }
        else if (!solution.SkipFeatureDeployment) {
          throw 'The solution has not been deployed to all sites. Make sure to deploy this solution to all sites.';
        }
      }

      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(appCatalogUrl, '/lists/TenantWideExtensions');
      const listItemId: number = await this.getListItemId(appCatalogUrl, args.options, listServerRelativeUrl, logger);

      await this.updateTenantWideExtension(logger, appCatalogUrl, args.options, listServerRelativeUrl, listItemId);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getComponentManifest(appCatalogUrl: string, clientSideComponentId: string, logger: Logger): Promise<any> {
    if (this.verbose) {
      logger.logToStderr('Retrieving component manifest item from the ComponentManifests list on the app catalog site so that we get the solution id');
    }

    const camlQuery = `<View><ViewFields><FieldRef Name='ClientComponentId'></FieldRef><FieldRef Name='SolutionId'></FieldRef><FieldRef Name='ClientComponentManifest'></FieldRef></ViewFields><Query><Where><Eq><FieldRef Name='ClientComponentId' /><Value Type='Guid'>${clientSideComponentId}</Value></Eq></Where></Query></View>`;
    const commandOptions: spoListItemListCommandOptions = {
      webUrl: appCatalogUrl,
      listUrl: `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/Lists/ComponentManifests`,
      camlQuery: camlQuery,
      verbose: this.verbose,
      debug: this.debug,
      output: 'json'
    };

    const output = await Cli.executeCommandWithOutput(spoListItemListCommand as Command, { options: { ...commandOptions, _: [] } });

    if (this.verbose) {
      logger.logToStderr(output.stderr);
    }

    const outputParsed = JSON.parse(output.stdout);

    if (outputParsed.length === 0) {
      throw 'No component found with the specified clientSideComponentId found in the component manifest list. Make sure that the application is added to the application catalog';
    }

    return outputParsed[0];
  }

  private async getSolutionFromAppCatalog(appCatalogUrl: string, solutionId: string, logger: Logger): Promise<Solution> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving solution with id ${solutionId} from the application catalog`);
    }

    const camlQuery = `<View><ViewFields><FieldRef Name='SkipFeatureDeployment'></FieldRef><FieldRef Name='ContainsTenantWideExtension'></FieldRef></ViewFields><Query><Where><Eq><FieldRef Name='AppProductID' /><Value Type='Guid'>${solutionId}</Value></Eq></Where></Query></View>`;
    const commandOptions: spoListItemListCommandOptions = {
      webUrl: appCatalogUrl,
      listUrl: `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/AppCatalog`,
      camlQuery: camlQuery,
      verbose: this.verbose,
      debug: this.debug,
      output: 'json'
    };

    const output = await Cli.executeCommandWithOutput(spoListItemListCommand as Command, { options: { ...commandOptions, _: [] } });

    if (this.verbose) {
      logger.logToStderr(output.stderr);
    }

    const outputParsed = JSON.parse(output.stdout);

    if (outputParsed.length === 0) {
      throw `No component found with the solution id ${solutionId}. Make sure that the solution is available in the app catalog`;
    }

    return outputParsed[0];
  }

  private async getListItemId(appCatalogUrl: string, options: Options, listServerRelativeUrl: string, logger: Logger): Promise<number> {
    const { title, id, clientSideComponentId } = options;
    const filter = title ? `Title eq '${title}'` : id ? `Id eq '${id}'` : `TenantWideExtensionComponentId eq '${clientSideComponentId}'`;

    if (this.verbose) {
      logger.logToStderr(`Getting tenant-wide listview commandset: "${title || id || clientSideComponentId}"...`);
    }

    const listItemInstances = await odata.getAllItems<ListItemInstance>(`${appCatalogUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/Items?$filter=startswith(TenantWideExtensionLocation, 'ClientSideExtension.ListViewCommandSet') and ${filter}`);

    if (!listItemInstances || listItemInstances.length === 0) {
      throw 'The specified listview commandset was not found';
    }

    if (listItemInstances.length > 1) {
      throw `Multiple listview commandsets with ${title ? `title '${title}'` : `ClientSideComponentId '${clientSideComponentId}'`} found. Please disambiguate using IDs: ${os.EOL}${listItemInstances.map(item => `- ${(item as any).Id}`).join(os.EOL)}`;
    }

    return listItemInstances[0].Id;
  }

  private async updateTenantWideExtension(logger: Logger, appCatalogUrl: string, options: Options, listServerRelativeUrl: string, listItemId: number): Promise<void> {
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

    if (options.newClientSideComponentId !== undefined) {
      formValues.push({
        FieldName: 'TenantWideExtensionComponentId',
        FieldValue: options.newClientSideComponentId
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
      url: `${appCatalogUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/Items(${listItemId})/ValidateUpdateListItem()`,
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

export default new SpoTenantCommandSetSetCommand();